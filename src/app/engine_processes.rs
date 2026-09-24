//! The processes this app starts, and how they are made to go with it.
//!
//! What this app starts would otherwise outlive it: the Office application that
//! draws a document's page, the browser that draws an SVG document, the player a
//! video preview runs in, and the probes that size one. Each of them is ended where
//! it is ended deliberately — an engine let go for idle, a player stopped — and what
//! this module is for is the ways an app does not get to choose: ended from Task
//! Manager, taken down by a crash, closed by a logoff.
//!
//! Two mechanisms, covering different halves of that.
//!
//! A job object holds every process adopted here, with `KILL_ON_JOB_CLOSE` set. The
//! handle is never closed by this app, so what closes it is this process ending —
//! however it ends — and the system ends everything in the job with it. That is the
//! only mechanism that asks nothing of a process that is dying, which is exactly
//! what a crash is: nothing of ours runs, and the OS does the rest.
//!
//! A record on disk covers what the job cannot: a process the job would not take,
//! and a process left by a run of this app that is already gone. Each run writes
//! what it started to `%LOCALAPPDATA%\rust-hover-preview\engines\<pid>.state`, and
//! the next run reads every file there — ending what the runs that are gone left
//! behind, and leaving alone what a run that is still alive is holding, because
//! another session of the same user is another run and its engines are its own.
//!
//! What is never done is acting on a process id by itself. An id recorded hours ago
//! may belong to something else by now, so a record is acted on only when the
//! process still runs the image the record names *and* started when the record says
//! it did. The browser is the one thing looked for by name, and it is found by its
//! parent — one of our own runs — with the further condition that that run is dead:
//! a live parent means the browser belongs to an app that is running, which is not
//! ours to end.

use std::ffi::c_void;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::{Duration, Instant};

use directories::BaseDirs;
use once_cell::sync::Lazy;
use windows::core::PCWSTR;
use windows::Win32::Foundation::{CloseHandle, FILETIME, HANDLE};
use windows::Win32::System::Diagnostics::ToolHelp::{
    CreateToolhelp32Snapshot, Process32FirstW, Process32NextW, PROCESSENTRY32W, TH32CS_SNAPPROCESS,
};
use windows::Win32::System::JobObjects::{
    AssignProcessToJobObject, CreateJobObjectW, JobObjectExtendedLimitInformation,
    SetInformationJobObject, JOBOBJECT_EXTENDED_LIMIT_INFORMATION,
    JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE,
};
use windows::Win32::System::Threading::{
    GetCurrentProcess, GetExitCodeProcess, GetProcessTimes, OpenProcess,
    QueryFullProcessImageNameW, TerminateProcess, PROCESS_NAME_WIN32,
    PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_SET_QUOTA, PROCESS_TERMINATE,
};

/// The browser the animated-document engine runs as, by image name.
pub const BROWSER_IMAGE: &str = "msedgewebview2.exe";

/// How a console program of this app's is started: without a console window of its own.
///
/// Every engine here that is asked for a file rather than for a document — `magick.exe` above
/// all, and FFmpeg's `ffprobe` and `ffplay` — is a console program, and this app is not: a
/// console program started by a program that has no console is given one of its own by
/// Windows, which is a window that appears for as long as the launch lasts. What says no to it
/// is this flag, on every spawn of one; a program that draws its own window, like the Office
/// applications and LibreOffice's launcher, is a program of the other kind and is started as
/// it is.
pub const CREATE_NO_WINDOW: u32 = 0x0800_0000;

/// The longest image path `QueryFullProcessImageNameW` is given room for.
const MAX_IMAGE_PATH: usize = 260;

/// What `GetExitCodeProcess` reports for a process that is still running.
const STILL_ACTIVE: u32 = 259;

/// How long a process that has been asked to end is given to go. A terminate is a
/// request, and what waits on it here is the start of its replacement: two engines
/// of one family at once is the thing being avoided, so the wait is short and the
/// answer to a process that outlasts it is to start nothing.
const PROCESS_GONE_WAIT: Duration = Duration::from_secs(2);

/// How often the wait above looks.
const PROCESS_GONE_POLL: Duration = Duration::from_millis(25);

// ------------------------------------------------------------------- the job

/// The job every process adopted here is put in, or nothing when one could not be
/// made — which costs the immediate end a crash would otherwise get, and nothing
/// else, since the record below is the other half.
///
/// The handle is kept as a number so that it can live in a `static`, and it is
/// deliberately never closed: closing it is what ends every process in the job, and
/// what closes it is this process ending, which is when that is wanted.
static JOB: Lazy<Option<usize>> = Lazy::new(create_job);

fn job() -> Option<HANDLE> {
    JOB.map(|handle| HANDLE(handle as *mut c_void))
}

fn create_job() -> Option<usize> {
    unsafe {
        let job = CreateJobObjectW(None, PCWSTR::null()).ok()?;

        let mut limits = JOBOBJECT_EXTENDED_LIMIT_INFORMATION::default();
        limits.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;

        let set = SetInformationJobObject(
            job,
            JobObjectExtendedLimitInformation,
            std::ptr::addr_of!(limits).cast(),
            std::mem::size_of::<JOBOBJECT_EXTENDED_LIMIT_INFORMATION>() as u32,
        );

        if set.is_err() {
            let _ = CloseHandle(job);
            return None;
        }

        Some(job.0 as usize)
    }
}

/// Put a process in the job, so that whatever ends this app ends it too.
///
/// Best effort, and only ever a process this app started: a refusal — a process
/// already inside a job that will not nest is the way that happens — costs the
/// immediate end and leaves the record to do it on the next start, so nothing here
/// is reported. What is adopted without being recorded is a probe — the `ffprobe`
/// and the cropdetect `ffmpeg` a video's geometry is read by — which lives a few
/// dozen milliseconds and has no window to leave behind: writing one down would be a
/// record made and struck off again inside the hover that started it.
pub fn adopt(pid: u32) -> bool {
    let Some(job) = job() else {
        return false;
    };

    unsafe {
        let Ok(process) = OpenProcess(PROCESS_SET_QUOTA | PROCESS_TERMINATE, false, pid) else {
            return false;
        };

        let adopted = AssignProcessToJobObject(job, process).is_ok();
        let _ = CloseHandle(process);

        adopted
    }
}

// --------------------------------------------------------------- the records

/// What a recorded process is, which is what decides how far an ending reaches.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Kind {
    /// An engine: a process that draws something for this app, ended with its tier
    /// when that tier is let go of — including every one of them at once.
    Engine,
    /// The player a video preview runs in. It is the picture the user is watching,
    /// so what ends it is its hover ending and not a tier going away.
    Player,
}

/// A process this app started, and what it takes to know it again later.
struct Owned {
    /// The image it runs, by file name, which is also what says what it is: Word,
    /// Excel, PowerPoint, the browser, the player.
    image: String,
    pid: u32,
    /// When the process started, in the ticks the state file carries. Zero when that
    /// could not be read, which is a record that can only be matched by name.
    created: u64,
    kind: Kind,
}

static OWNED: Lazy<Mutex<Vec<Owned>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// Take charge of an engine this app started: put it in the job, hold it, and write
/// it down.
pub fn record(image: &str, pid: u32) {
    record_as(Kind::Engine, image, pid);
}

/// Take charge of the player a video preview runs in.
///
/// Written down the way an engine is, so that a run which never got to end it —
/// killed, crashed, or closed with the player still up — is answered for by the next
/// one. It is the record that makes this worth doing where the job alone was not:
/// the job takes the player with the app, but a process the job would not take is a
/// looping player left on screen, and the only thing that reaches one of those is
/// the next run.
///
/// The kind is the difference that matters: the player is held in the same list as
/// the engines and is kept out of `terminate_all_owned`, because a preview type
/// switched off or a worker given up on is nothing to do with a video that is
/// playing.
pub fn record_player(image: &str, pid: u32) {
    record_as(Kind::Player, image, pid);
}

fn record_as(kind: Kind, image: &str, pid: u32) {
    if pid == 0 {
        return;
    }

    adopt(pid);

    let created = creation_time(pid).unwrap_or(0);
    if let Ok(mut owned) = OWNED.lock() {
        owned.retain(|owned| owned.pid != pid);
        owned.push(Owned {
            image: image.to_string(),
            pid,
            created,
            kind,
        });
    }

    persist();
}

/// Stop holding a process, which is what a caller does once it has seen it go.
pub fn forget(pid: u32) {
    if let Ok(mut owned) = OWNED.lock() {
        owned.retain(|owned| owned.pid != pid);
    }

    persist();
}

/// End a process this app started, if it is still there and still the one recorded.
/// `true` when it is gone — ended here, or already gone when it was asked about.
///
/// A record is kept while the process it names may still be alive, because a
/// terminate is a request: what has to know about a process that has not gone yet is
/// whatever is about to start its replacement.
pub fn terminate_owned(pid: u32) -> bool {
    let Some((image, created)) = record_of(pid) else {
        return true;
    };

    let requested = terminate_verified(pid, &image, (created != 0).then_some(created));
    let gone = !is_running(pid);

    if gone {
        forget(pid);
    }

    requested || gone
}

/// End every engine this app started that is still recorded.
///
/// Engines and not everything held: the player a video preview runs in is recorded
/// too, and it is the one process in this list the user is looking at. What this is
/// called for is a tier being let go of — a preview type switched off in the tray, a
/// worker given up on — and a video that is playing has nothing to do with either.
/// What ends a player is its hover ending, and what answers for one this app never
/// got to end is the record.
pub fn terminate_all_owned() {
    for pid in recorded_pids(None, Some(Kind::Engine)) {
        terminate_owned(pid);
    }
}

/// End whatever this app still holds that runs this image, and say whether anything
/// of it is left.
///
/// This is what keeps one engine per family: a process this app started for a family
/// is ended here, and its replacement is started only once it is gone. What the
/// answer is used for is refusing to start a second engine beside one that would not
/// end — a preview that does not render this time, which is nothing next to two
/// Office processes on the machine.
pub fn end_recorded(image: &str) -> bool {
    let pids = recorded_pids(Some(image), None);
    if pids.is_empty() {
        return true;
    }

    for pid in &pids {
        terminate_owned(*pid);
    }

    let deadline = Instant::now() + PROCESS_GONE_WAIT;
    for pid in &pids {
        while is_running(*pid) && Instant::now() < deadline {
            std::thread::sleep(PROCESS_GONE_POLL);
        }
    }

    pids.into_iter().all(|pid| {
        let gone = !is_running(pid);
        if gone {
            forget(pid);
        }

        gone
    })
}

fn record_of(pid: u32) -> Option<(String, u64)> {
    let owned = OWNED.lock().ok()?;

    owned
        .iter()
        .find(|owned| owned.pid == pid)
        .map(|owned| (owned.image.clone(), owned.created))
}

/// The ids this app is holding, for one image or for all of them, and for one kind of
/// process or for both.
fn recorded_pids(image: Option<&str>, kind: Option<Kind>) -> Vec<u32> {
    OWNED
        .lock()
        .map(|owned| {
            owned
                .iter()
                .filter(|owned| match image {
                    Some(image) => owned.image.eq_ignore_ascii_case(image),
                    None => true,
                })
                .filter(|owned| kind.is_none_or(|kind| owned.kind == kind))
                .map(|owned| owned.pid)
                .collect()
        })
        .unwrap_or_default()
}

// --------------------------------------------------------------- the record on disk

/// The folder the run files live in.
///
/// `RHP_ENGINE_STATE` names a folder for one run, the way `RHP_WEBVIEW_PROFILE` names
/// a profile, and it is there for the same reason: a test must be able to reap what
/// it wrote without being confused with the records of the app somebody is actually
/// running.
fn state_folder() -> Option<PathBuf> {
    if let Some(folder) = std::env::var_os("RHP_ENGINE_STATE") {
        return Some(PathBuf::from(folder));
    }

    Some(
        BaseDirs::new()?
            .data_local_dir()
            .join("rust-hover-preview")
            .join("engines"),
    )
}

/// This run's file, named by its own process id.
fn run_file() -> Option<PathBuf> {
    Some(state_folder()?.join(format!("{}.state", std::process::id())))
}

/// Write what this run is holding, so that a run which never gets to end them can
/// still be answered for.
///
/// Best effort, and rewritten whole rather than appended to: a file written while a
/// process is being started is the file as it was a moment ago, which is the half
/// that matters — what it leaves off is a process nobody recorded, which is what a
/// crash between the two would have left anyway.
fn persist() {
    let Some(path) = run_file() else {
        return;
    };
    let Some(folder) = path.parent() else {
        return;
    };

    if std::fs::create_dir_all(folder).is_err() {
        return;
    }

    let mut text = format!("run {} {}\n", std::process::id(), creation_time_self());

    if let Ok(owned) = OWNED.lock() {
        for owned in owned.iter() {
            text.push_str(&format!(
                "engine {} {} {}\n",
                owned.image, owned.pid, owned.created
            ));
        }
    }

    let _ = std::fs::write(&path, text);
}

/// End what the runs that are gone left behind, and take their files with them.
///
/// Called once at startup, before anything of this run can start a process: what a
/// previous run could not end — because it was killed, or crashed, or lost power —
/// is ended here instead, and it is ended before this run could add another beside
/// it.
pub fn reap_leftovers() {
    let Some(folder) = state_folder() else {
        return;
    };
    let Ok(entries) = std::fs::read_dir(&folder) else {
        return;
    };

    for entry in entries.flatten() {
        let path = entry.path();
        if path.extension().and_then(|extension| extension.to_str()) != Some("state") {
            continue;
        }

        let Ok(text) = std::fs::read_to_string(&path) else {
            continue;
        };

        // A file whose run is still alive is that run's. Another session of the same
        // user is another run, and what it is holding is not this one's to end.
        if run_is_alive(&text) {
            continue;
        }

        reap(&text);
        let _ = std::fs::remove_file(&path);
    }
}

/// Whether the run that wrote this file is still that process, rather than one that
/// has since been given its id.
fn run_is_alive(text: &str) -> bool {
    let Some((pid, created)) = run_line(text) else {
        return false;
    };

    // A run whose own start could not be read is one that cannot be told from a
    // process that took its id, and a file that cannot be vouched for is one this
    // leaves alone: what that costs is a file left lying, which is nothing next to
    // ending a process that is not ours.
    created != 0 && process_matches(pid, created)
}

/// End the engines a run recorded, and the browser it was playing documents in.
fn reap(text: &str) {
    for line in text.lines() {
        let Some((image, pid, created)) = engine_line(line) else {
            continue;
        };

        terminate_verified(pid, &image, (created != 0).then_some(created));
    }

    if let Some((run_pid, _)) = run_line(text) {
        end_browsers_started_by(run_pid);
    }
}

/// End the browsers started by a run that is over.
///
/// The browser is the one process looked for by name, and what makes that safe is
/// the two conditions together: its parent is one of our own runs, and that run is
/// gone. A live parent is an app that is running — another session's, or this run's
/// own before it has closed its engine — and a browser it is holding is not ours to
/// end. The browser's own children are its business: ending it ends them.
pub fn end_browsers_started_by(run_pid: u32) {
    if run_pid == 0 || is_running(run_pid) {
        return;
    }

    for pid in processes_named_by_parent(BROWSER_IMAGE, run_pid) {
        terminate_verified(pid, BROWSER_IMAGE, None);
    }
}

/// End every browser this app started, whether or not it was recorded.
///
/// The browser is the one process looked for by name, and the parent is what makes
/// that safe: a browser the WebView2 loader starts runs as a child of the process
/// that asked for it, and a browser started by anything else is a child of that
/// something else. What this catches is the engine whose browser could not be told
/// apart from an earlier one's — which is a browser that was never recorded, and so
/// would otherwise be left to the next run's reaper.
pub fn end_our_browsers() {
    for pid in processes_named_by_parent(BROWSER_IMAGE, std::process::id()) {
        terminate_verified(pid, BROWSER_IMAGE, None);
    }
}

/// The process ids of the folders a previous run left its browser's state in.
///
/// Every folder under the browser's profile root was made by one of our runs and
/// named for the run that made it, so a folder that is not this run's names a run
/// whose browser may still be holding it. This is what reaches a browser left by a
/// version of this app that did not write a record of it.
pub fn stale_run_pids(root: &Path) -> Vec<u32> {
    let ours = std::process::id();
    let Ok(entries) = std::fs::read_dir(root) else {
        return Vec::new();
    };

    entries
        .flatten()
        .filter(|entry| entry.path().is_dir())
        .filter_map(|entry| entry.file_name().to_string_lossy().parse::<u32>().ok())
        .filter(|pid| *pid != ours)
        .collect()
}

/// The run line of a state file: the process that wrote it, and when it started.
fn run_line(text: &str) -> Option<(u32, u64)> {
    let line = text.lines().find(|line| line.starts_with("run "))?;
    let mut parts = line.split_whitespace();

    if parts.next()? != "run" {
        return None;
    }

    let pid = parts.next()?.parse().ok()?;
    let created = parts.next()?.parse().ok()?;

    Some((pid, created))
}

/// An engine line of a state file: the image it runs, its id, and when it started.
fn engine_line(line: &str) -> Option<(String, u32, u64)> {
    let mut parts = line.split_whitespace();

    if parts.next()? != "engine" {
        return None;
    }

    let image = parts.next()?.to_string();
    let pid = parts.next()?.parse().ok()?;
    let created = parts.next()?.parse().ok()?;

    Some((image, pid, created))
}

// ------------------------------------------------------------- the processes

/// End a process, having first confirmed it is the one the record names.
///
/// The creation time is the half that makes this safe to do to an id read from a
/// file: ids are recycled, and the name alone is not enough when the process being
/// asked about was recorded hours ago. A record without one — the start could not be
/// read when it was written — is matched by name, which is what the callers that
/// hold a process in hand can afford and a file cannot.
pub fn terminate_verified(pid: u32, image: &str, created: Option<u64>) -> bool {
    unsafe {
        let Ok(process) = OpenProcess(
            PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION,
            false,
            pid,
        ) else {
            // One that cannot be opened is one that is gone, or one that is not ours
            // to end; either way there is nothing here to do.
            return false;
        };

        let named = image_path(process)
            .map(|path| file_name(&path).eq_ignore_ascii_case(image))
            .unwrap_or(false);
        let same = created.is_none_or(|created| creation_time_of(process) == Some(created));

        let ended = named && same && TerminateProcess(process, 1).is_ok();
        let _ = CloseHandle(process);

        ended
    }
}

/// Whether a process is still running, by id. One that cannot be opened is one that
/// is gone.
pub fn is_running(pid: u32) -> bool {
    unsafe {
        let Ok(process) = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid) else {
            return false;
        };

        let mut code = 0u32;
        let running = GetExitCodeProcess(process, &mut code).is_ok() && code == STILL_ACTIVE;
        let _ = CloseHandle(process);

        running
    }
}

/// Whether the process with this id is that same process: the id and the moment it
/// started, together.
fn process_matches(pid: u32, created: u64) -> bool {
    unsafe {
        let Ok(process) = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid) else {
            return false;
        };

        let matches = creation_time_of(process) == Some(created);
        let _ = CloseHandle(process);

        matches
    }
}

/// The processes running an image of this name, by executable name.
pub fn processes_named(image_name: &str) -> Vec<u32> {
    processes_with_parent(image_name)
        .into_iter()
        .map(|(pid, _)| pid)
        .collect()
}

/// The processes running an image of this name whose parent is `parent`.
pub fn processes_named_by_parent(image_name: &str, parent: u32) -> Vec<u32> {
    processes_with_parent(image_name)
        .into_iter()
        .filter(|(_, found_parent)| *found_parent == parent)
        .map(|(pid, _)| pid)
        .collect()
}

/// The processes running an image of this name, each with the process that started
/// it.
fn processes_with_parent(image_name: &str) -> Vec<(u32, u32)> {
    let mut found = Vec::new();

    unsafe {
        let Ok(snapshot) = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) else {
            return found;
        };

        let mut entry = PROCESSENTRY32W {
            dwSize: std::mem::size_of::<PROCESSENTRY32W>() as u32,
            ..Default::default()
        };

        if Process32FirstW(snapshot, &mut entry).is_ok() {
            loop {
                let name = String::from_utf16_lossy(&entry.szExeFile);
                if name.trim_end_matches('\0').eq_ignore_ascii_case(image_name) {
                    found.push((entry.th32ProcessID, entry.th32ParentProcessID));
                }
                if Process32NextW(snapshot, &mut entry).is_err() {
                    break;
                }
            }
        }

        let _ = CloseHandle(snapshot);
    }

    found
}

/// When a process started, in the ticks a `FILETIME` counts. Nothing when that
/// cannot be read.
fn creation_time(pid: u32) -> Option<u64> {
    unsafe {
        let process = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid).ok()?;
        let created = creation_time_of(process);
        let _ = CloseHandle(process);

        created
    }
}

/// When this process started.
fn creation_time_self() -> u64 {
    unsafe { creation_time_of(GetCurrentProcess()).unwrap_or(0) }
}

fn creation_time_of(process: HANDLE) -> Option<u64> {
    let mut created = FILETIME::default();
    let mut exited = FILETIME::default();
    let mut kernel = FILETIME::default();
    let mut user = FILETIME::default();

    unsafe {
        GetProcessTimes(process, &mut created, &mut exited, &mut kernel, &mut user).ok()?;
    }

    Some(((created.dwHighDateTime as u64) << 32) | created.dwLowDateTime as u64)
}

fn image_path(process: HANDLE) -> Option<String> {
    let mut name = [0u16; MAX_IMAGE_PATH];
    let mut length = name.len() as u32;

    unsafe {
        QueryFullProcessImageNameW(
            process,
            PROCESS_NAME_WIN32,
            windows::core::PWSTR(name.as_mut_ptr()),
            &mut length,
        )
        .ok()?;
    }

    Some(String::from_utf16_lossy(&name[..length as usize]))
}

/// A path's last element, which is what a record names an image by: what is stored
/// is the name and not the folder, so an Office installed somewhere else is still
/// the same application.
fn file_name(path: &str) -> &str {
    path.rsplit(['\\', '/']).next().unwrap_or(path)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reads_the_lines_a_state_file_is_made_of() {
        let text = "run 4242 133700\nengine EXCEL.EXE 5150 133800\nengine msedgewebview2.exe 5151 133900\n";

        assert_eq!(run_line(text), Some((4242, 133700)));
        assert_eq!(
            engine_line("engine EXCEL.EXE 5150 133800"),
            Some(("EXCEL.EXE".to_string(), 5150, 133800))
        );
        assert_eq!(engine_line("run 4242 133700"), None);
        assert_eq!(run_line("engine EXCEL.EXE 5150 133800"), None);
    }

    /// A file whose run cannot be vouched for is left alone rather than acted on:
    /// what it costs is a file left lying, which is nothing next to ending a process
    /// that is not ours.
    #[test]
    fn a_run_with_no_start_time_is_not_taken_for_a_live_one() {
        assert!(!run_is_alive("run 4242 0\n"));
        assert!(!run_is_alive("run 4242\n"));
        assert!(!run_is_alive(""));
    }

    #[test]
    fn reads_an_image_by_its_last_element() {
        assert_eq!(file_name(r"C:\Program Files\Office\EXCEL.EXE"), "EXCEL.EXE");
        assert_eq!(file_name("msedgewebview2.exe"), "msedgewebview2.exe");
        assert_eq!(file_name(""), "");
    }

    /// This process is one the module can answer for: the checks the reaper is built
    /// on, against something whose answers are known.
    #[test]
    fn knows_this_process_is_running_and_when_it_started() {
        let pid = std::process::id();
        let created = creation_time(pid).expect("a start time for this process");

        assert!(is_running(pid));
        assert!(process_matches(pid, created));
        assert!(!process_matches(pid, created + 1), "a different process");
        assert_eq!(creation_time_self(), created);
    }

    /// The one thing that makes an id from a file safe to act on: a process that is
    /// not the one recorded is never ended.
    #[test]
    fn will_not_end_a_process_that_is_not_the_one_recorded() {
        let pid = std::process::id();
        let created = creation_time(pid).expect("a start time for this process");
        let image = unsafe { image_path(GetCurrentProcess()) }
            .map(|path| file_name(&path).to_string())
            .expect("this process's own image");

        assert!(
            !terminate_verified(pid, "not-our-image.exe", Some(created)),
            "an image that does not match"
        );
        assert!(
            !terminate_verified(pid, &image, Some(created + 1)),
            "a start time that does not match, image and all"
        );
        assert!(is_running(pid), "and it is still running");
    }

    /// A process that stays up long enough to be asked about, standing in for an
    /// engine: what the reaper does to one it was told about is the same whatever the
    /// process is, and a test is not going to start an Office application to find out.
    fn stand_in() -> std::process::Child {
        std::process::Command::new("ping")
            .args(["-n", "30", "127.0.0.1"])
            .stdout(std::process::Stdio::null())
            .spawn()
            .expect("a stand-in process")
    }

    /// The image name a record for this process would carry.
    fn image_of(pid: u32) -> String {
        let process = unsafe { OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid) }
            .expect("the stand-in to be openable");
        let image = image_path(process)
            .map(|path| file_name(&path).to_string())
            .expect("its image");

        unsafe {
            let _ = CloseHandle(process);
        }

        image
    }

    /// The whole point of the reaper, and the line it must not cross: what a run that
    /// is gone left behind is ended, and what a run that is still alive is holding is
    /// not touched — another session of the same user is another run.
    ///
    /// One test rather than two, because the folder it reaps from is named in the
    /// environment, and two tests setting that at once are two tests reading each
    /// other's files.
    #[test]
    fn reaps_a_run_that_is_gone_and_leaves_a_live_one_alone() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-engine-tests")
            .join(std::process::id().to_string());
        let _ = std::fs::remove_dir_all(&folder);
        std::fs::create_dir_all(&folder).expect("a test folder");
        std::env::set_var("RHP_ENGINE_STATE", &folder);

        let mut abandoned = stand_in();
        let mut held = stand_in();

        // One run that is gone — an id no process holds — holding the first process,
        // and one that is alive — this one — holding the second.
        std::fs::write(
            folder.join("gone.state"),
            format!(
                "run 999999999 1\nengine {} {} {}\n",
                image_of(abandoned.id()),
                abandoned.id(),
                creation_time(abandoned.id()).expect("a start time")
            ),
        )
        .expect("a written file");
        std::fs::write(
            folder.join("live.state"),
            format!(
                "run {} {}\nengine {} {} {}\n",
                std::process::id(),
                creation_time_self(),
                image_of(held.id()),
                held.id(),
                creation_time(held.id()).expect("a start time")
            ),
        )
        .expect("a written file");

        reap_leftovers();

        assert!(
            !is_running(abandoned.id()),
            "the abandoned process is ended"
        );
        assert!(is_running(held.id()), "the one a live run holds is not");
        assert!(
            !folder.join("gone.state").exists(),
            "and its file goes with it"
        );
        assert!(
            folder.join("live.state").exists(),
            "a run that is alive keeps its file"
        );

        let _ = held.kill();
        let _ = held.wait();
        let _ = abandoned.wait();
        let _ = std::fs::remove_dir_all(&folder);
        std::env::remove_var("RHP_ENGINE_STATE");
    }
}

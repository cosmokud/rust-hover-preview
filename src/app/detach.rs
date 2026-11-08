//! Detach Mode: a pinned window handed to a process of this app's own.
//!
//! A pin is one window and one whole set of process-wide facts — its state, its bubble, its film
//! host, its chrome strips, the player it plays through, the thumbnail it is drawn by — and those
//! facts are singletons by construction (see `ui::preview_window::pin_window::PIN_STATE`, and the
//! statics of every `pin_*.rs` beside it). A second pin *inside* this process would therefore be
//! the first pin's window with a second name on it. So a detached window is not a second pin here:
//! it is a second process of the same executable, begun in a mode that takes one pin up on the file
//! it is handed and never opens a hover at all.
//!
//! What that buys is the whole of what a pin is, without a line of it being made per-window: the
//! caption, the bubble, the taskbar stand-in, the walk, the menus, the player. What it costs is one
//! process per detached window, and what it needs is the three things this module is:
//!
//! - **What is handed over.** The file, the media box the pin was showing it in, and whether the
//!   window is maximized. The box is the *media's* rather than the window's, because a take-up in
//!   the new process builds the window around a media box (see `pin_install::take_up_pinned_window`):
//!   handed the media box, the child builds the window the parent is standing in.
//! - **When the hand is finished.** The child sets an event once its window is up and drawn, and
//!   only then does the tray app's own pin come down — so the file is never in neither window, and
//!   a child that dies before its window appears leaves the pin it came from standing.
//! - **What ends it.** Nothing here outlives the tray app: the child is put in the same job every
//!   engine of this app's is in (`engine_processes::adopt`), which ends it the moment the app does,
//!   and it watches the process that began it as well, for the machine where a job cannot be
//!   joined. A detached window has one way out of its own — its Close, its Escape, its menu — and
//!   that ending is the process's own: the loop it draws on sets `RUNNING` false and `main` follows
//!   the same shutdown the tray app's does.
//!
//! A detached window is a separate *process*, and that is what makes it inert with respect to
//! Explorer: no hook of this app's runs in it, so nothing of the listing — a click, a hover, a key,
//! `Tray → Pin Mode → Update Preview` — can reach the file in it. It reads the configuration file
//! like any other run, so the tray's own rows reach it the way they reach a pin: the ones written
//! down take effect live, and the ones a window only reads at its take-up (the engine a film is
//! played by, the scale a box was measured at) reach it at its next start, which is the answer a
//! pin has always given.

use std::ffi::OsString;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicU32, Ordering};
use std::sync::Mutex;
use std::time::Duration;

use once_cell::sync::Lazy;
use windows::core::PCWSTR;
use windows::Win32::Foundation::{CloseHandle, BOOL, HANDLE};
use windows::Win32::System::Threading::{
    CreateEventW, OpenEventW, SetEvent, WaitForSingleObject, EVENT_MODIFY_STATE,
};

/// The switch that says this process is a detached window rather than the tray app.
const DETACHED: &str = "--detach";
/// The media box the window is asked to stand around, in physical pixels, as `x,y,x,y`.
const BOX: &str = "--box";
/// The event this process is to set once its window is up, and the one the tray app waits on.
const READY: &str = "--ready";
/// Whether the pin that was handed over was maximized.
const MAXIMIZED: &str = "--maximized";
/// Where the sound the pin was playing had got to, in seconds into its file.
const AT: &str = "--at";
/// The process that began this one, waited on so that a window outlives nobody.
const PARENT: &str = "--parent";

/// How often the wait for a detached window's first paint looks.
///
/// The wait is a thread's own and nothing is behind it: the pin it was handed over from is still
/// on the screen until this answers, which is the whole reason the wait exists.
const READY_POLL: Duration = Duration::from_millis(50);

/// How long a hand-off is given before the wait is given up on, leaving the pin it came from
/// standing. A document handed to an Office engine can spend fifteen seconds being rendered and
/// still arrive, so this is generous rather than tight; a child that never answers at all is a
/// child whose window the user never sees, and a pin that stays is the better of the two.
const READY_WAIT: Duration = Duration::from_secs(300);

/// A media box on the screen: left, top, right, bottom, in physical pixels.
pub(crate) type MediaBox = (i32, i32, i32, i32);

/// What a detached window was begun with.
#[derive(Clone, PartialEq, Debug)]
pub(crate) struct OpenRequest {
    /// The file the window is to stand on.
    pub(crate) path: PathBuf,
    /// The media box the pin it was handed over from was showing it in.
    pub(crate) region: MediaBox,
    /// Whether that pin's window was maximized, which is put back on after the take-up rather than
    /// asked for by the take-up: a maximized window is a box chosen for the whole window, and only
    /// the window that now exists can be given one.
    pub(crate) maximized: bool,
    /// Where the sound of the file had got to in the pin this was handed over from, for the file
    /// that pin was showing, and nothing for every other kind: a picture, a film's first frame and
    /// a page are read the same wherever they are opened, and a second into a sound is not in the
    /// file at all — it is where the player on the other side of this hand-off had got to, and the
    /// window that takes the file up cannot ask a process that no longer has it.
    pub(crate) at: Option<f64>,
    /// The event to set once the window is up, or nothing where nobody is waiting — a window begun
    /// by hand is one whose first paint nobody owes an answer to.
    pub(crate) ready: Option<String>,
    /// The process that began this one, if it was begun by one.
    pub(crate) parent: Option<u32>,
}

/// How this process was started.
#[derive(PartialEq, Debug)]
pub(crate) enum Startup {
    /// The tray app, which is what a launch with no switch of ours on it is.
    Tray,
    /// A detached window.
    Detached(OpenRequest),
    /// `--detach` with no file or no box behind it. Nothing is started at all: a request this app
    /// cannot answer must not become a second tray icon, which is what reading it as `Tray` would
    /// be.
    Refused,
}

/// What this process is drawing, once `main` has said which of the three it is.
static OPEN: Lazy<Mutex<Option<OpenRequest>>> = Lazy::new(|| Mutex::new(None));

/// Whether this process is a detached window, published so that the loop's own tick may ask
/// without taking the lock above on every turn.
static DETACHED_HERE: AtomicBool = AtomicBool::new(false);

/// How many windows this process has begun, which is what keeps two hand-offs' events apart when
/// both are made from the same run.
static HANDED_OVER: AtomicU32 = AtomicU32::new(1);

/// How this process was started, read from its own command line.
///
/// The command line is this app's own on both sides of a hand-off — the switch is one `main` wrote
/// and one `main` reads — so nothing here parses a line a shell made: an `OsString` per argument,
/// which is what `std::env::args_os` hands back and what a path with a space or a character outside
/// this machine's code page survives inside.
pub(crate) fn startup_mode() -> Startup {
    parse(std::env::args_os().skip(1))
}

/// The same question asked of arguments handed in rather than of this process's own.
fn parse(args: impl IntoIterator<Item = OsString>) -> Startup {
    let mut args = args.into_iter();
    let mut asked = false;
    let mut path: Option<PathBuf> = None;
    let mut region: Option<MediaBox> = None;
    let mut maximized = false;
    let mut at: Option<f64> = None;
    let mut ready: Option<String> = None;
    let mut parent: Option<u32> = None;

    while let Some(arg) = args.next() {
        match arg.to_str() {
            Some(DETACHED) => {
                asked = true;
                path = args.next().map(PathBuf::from);
            }
            Some(BOX) => region = args.next().and_then(|value| reader_box(&value)),
            Some(READY) => ready = args.next().and_then(|value| value.into_string().ok()),
            // A pid is a number, and a child that cannot read its parent's is a child with one
            // guard fewer rather than one that does not start.
            Some(PARENT) => {
                parent = args
                    .next()
                    .and_then(|value| value.into_string().ok())
                    .and_then(|value| value.parse().ok())
            }
            Some(MAXIMIZED) => maximized = true,
            Some(AT) => at = args.next().and_then(|value| reader_at(&value)),
            // Anything else is not this switch's business and not a reason to read the launch as a
            // detached one: what is looked for is `--detach`, and nothing else here can add it.
            _ => {}
        }
    }

    if !asked {
        return Startup::Tray;
    }

    match (path, region) {
        (Some(path), Some(region)) if !path.as_os_str().is_empty() => {
            Startup::Detached(OpenRequest {
                path,
                region,
                maximized,
                at,
                ready,
                parent,
            })
        }
        // A file is a file and a box is a box: a request missing either is a request this app
        // cannot stand a window around, and a window it cannot stand is one it does not begin.
        _ => Startup::Refused,
    }
}

/// Read a box written as `x,y,x,y`, or nothing where the writing is not one.
fn reader_box(value: &std::ffi::OsStr) -> Option<MediaBox> {
    let value = value.to_str()?;
    let mut parts = value.split(',');
    let mut number = || {
        parts
            .next()
            .and_then(|part| part.trim().parse::<i32>().ok())
    };

    let region = (number()?, number()?, number()?, number()?);

    // A box with no room in it is not a box a window can be built around: what a take-up does with
    // one is divide by it.
    if parts.next().is_some() || region.2 <= region.0 || region.3 <= region.1 {
        return None;
    }

    Some(region)
}

/// Read where a sound was left off, written as a number of seconds, or nothing where the writing is
/// not one: a second this app cannot read is a sound that is taken up by its own rule, which is a
/// window that plays rather than a hand-off that does not happen.
fn reader_at(value: &std::ffi::OsStr) -> Option<f64> {
    let seconds = value.to_str()?.trim().parse::<f64>().ok()?;

    (seconds.is_finite() && seconds >= 0.0).then_some(seconds)
}

/// Where a sound this window was handed over is to begin: the second the pin it came from had got
/// to, for the file that pin was showing, and nothing for any other.
///
/// It is the one fact about a hand-off a file cannot be asked for on the other side of it, and it is
/// what the window that takes the file up reads before it falls back to the way a hover of that
/// file drops a sound in (see the loop's own start of a sound and `pin_detach`).
pub(crate) fn the_sound_starts_at(path: &Path) -> Option<f64> {
    open_request()
        .filter(|request| request.path == path)
        .and_then(|request| request.at)
}

/// Publish what this process was started for. Called once, by `main`, before any thread that asks
/// is begun.
pub(crate) fn begin(request: OpenRequest) {
    DETACHED_HERE.store(true, Ordering::Release);
    if let Ok(mut open) = OPEN.lock() {
        *open = Some(request);
    }
}

/// Whether this process is a detached window.
pub(crate) fn is_detached() -> bool {
    DETACHED_HERE.load(Ordering::Acquire)
}

/// What this window was begun with, or nothing where this process is the tray app.
pub(crate) fn open_request() -> Option<OpenRequest> {
    OPEN.lock().ok().and_then(|open| open.clone())
}

/// Tell whoever began this window that it is up, which is the moment the pin it was handed over
/// from comes down.
///
/// A failure is nothing to act on: a window nobody is waiting on is a window begun by hand, and an
/// event that has gone is a tray app that ended while this one was loading — in both cases the
/// window itself is the thing that matters, and it is already standing.
pub(crate) fn the_window_is_up() {
    let Some(name) = open_request().and_then(|request| request.ready) else {
        return;
    };
    let name: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();

    unsafe {
        let Ok(event) = OpenEventW(EVENT_MODIFY_STATE, BOOL(0), PCWSTR(name.as_ptr())) else {
            return;
        };
        let _ = SetEvent(event);
        let _ = CloseHandle(event);
    }
}

/// Let this process go, which is what a detached window has instead of a tray's `Exit`: its one
/// window has ended, so there is nothing left for it to be.
pub(crate) fn the_process_may_go() {
    crate::RUNNING.store(false, Ordering::SeqCst);
}

/// End this process when the one that began it goes.
///
/// The job every hand-off joins ends a child with the app that began it, and this is the same
/// ending asked the other way: a job that could not be joined (a run already inside one that will
/// not nest) is the case the job cannot cover, and a window of this app's left standing over the
/// desktop with nothing to close it is the whole of what that would cost. The wait is a thread of
/// its own and asks nothing of the loop; the loop reads `RUNNING` on its next turn and ends the way
/// it ends when its own window is closed.
pub(crate) fn watch_the_parent() {
    use windows::Win32::System::Threading::{
        OpenProcess, WaitForSingleObject, PROCESS_SYNCHRONIZE,
    };

    let Some(parent) = open_request().and_then(|request| request.parent) else {
        return;
    };

    std::thread::spawn(move || {
        let process = unsafe { OpenProcess(PROCESS_SYNCHRONIZE, BOOL(0), parent) };
        // A process that cannot be opened is one that has gone: a pid that is no longer there is a
        // window whose owner is not, and this one has nothing left to stand beside.
        let Ok(process) = process else {
            the_process_may_go();
            return;
        };
        unsafe {
            // The call waits until the process it is on ends, and a process that has already ended
            // answers at once: both are the same answer here.
            let _ = WaitForSingleObject(process, u32::MAX);
            let _ = CloseHandle(process);
        }
        the_process_may_go();
    });
}

/// Hand a pinned window to a process of its own: the file, the media box it is in, and whether the
/// window is maximized.
///
/// Answers nothing about the window's own ending. What comes down when the new window is up is the
/// pin this was asked of, and that is the ready event's business — the tray app's pin stays standing
/// until the child says its window is there, so what the user sees is a window being *replaced* by
/// the same window rather than a file that blinks out and back.
pub(crate) fn hand_off(
    path: &Path,
    region: MediaBox,
    maximized: bool,
    at: Option<f64>,
) -> std::io::Result<()> {
    let exe = std::env::current_exe()?;
    let serial = HANDED_OVER.fetch_add(1, Ordering::Relaxed);
    let ready = format!(
        "Local\\rust-hover-preview-detach-{}-{serial}",
        std::process::id()
    );
    let event = create_the_ready_event(&ready)?;

    let mut command = crate::app::engine_processes::hidden_command(exe);
    command
        .arg(DETACHED)
        .arg(path)
        .arg(BOX)
        .arg(format!(
            "{},{},{},{}",
            region.0, region.1, region.2, region.3
        ))
        .arg(READY)
        .arg(&ready)
        .arg(PARENT)
        .arg(care_of().to_string());
    if maximized {
        command.arg(MAXIMIZED);
    }
    // Where the sound had got to, written down to the thousandth of a second it is a position in:
    // the number is the other side's answer read back as one, which is why it crosses as text rather
    // than as a shape of its own.
    if let Some(at) = at.filter(|seconds| seconds.is_finite() && *seconds >= 0.0) {
        command.arg(AT).arg(format!("{at:.3}"));
    }

    let child = command.spawn()?;

    // **The job is the tray app's alone.** A hand-off made from a detached window cannot put its
    // child in the job that outlives it — a job is a process's own handle — and the job a detached
    // window *does* own is the one that ends with it, which is right for its engines and wrong here:
    // the new window would be ended by the hand-off that began it. So the chain is bound to the tray
    // app by the process id it was handed instead, and by the job only where this process is the one
    // that owns it (see `care_of` and `watch_the_parent`).
    if !is_detached() {
        crate::app::engine_processes::adopt(child.id());
    }

    let path = path.to_path_buf();
    // The handle crosses as a number for the reason every handle in this tree does — a handle is a
    // raw pointer, and the wait is a thread of its own (see `PinWindow::hwnd`).
    let event = event.0 as isize;
    std::thread::spawn(move || wait_for_the_hand_over(child, event, &path));
    Ok(())
}

/// The process a hand-off is bound to: the one that owns the tray icon.
///
/// It is this process where this process is the tray app, and the process that began this window
/// where this one is a detached window being handed a pin again — the tray app's own id, carried
/// down the chain rather than the immediate parent's, so that a window handed over twice is still
/// ended by the app the user closes rather than by the window that handed it on.
fn care_of() -> u32 {
    open_request()
        .and_then(|request| request.parent)
        .unwrap_or_else(std::process::id)
}

/// The event a hand-off is waited on, made before the child is begun so that no child can set an
/// event nobody has made yet.
fn create_the_ready_event(name: &str) -> std::io::Result<HANDLE> {
    let name: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();

    // Manual reset and initially clear: the event is set once and read once, and a waiter's own
    // read is not allowed to clear it for a waiter that has not looked yet.
    unsafe { CreateEventW(None, BOOL(1), BOOL(0), PCWSTR(name.as_ptr())) }
        .map_err(|error| std::io::Error::other(error.to_string()))
}

/// Wait for a hand-off to be answered: the event set, or the child gone, or the wait given up on.
///
/// The pin this was handed over from is taken down only where the window it became is standing on
/// the file it was handed — a pin the user closed, or walked to another file, while the child was
/// loading is not this hand-off's to end. That is the one guard the whole road needs: everything
/// else about it is a wait.
fn wait_for_the_hand_over(mut child: std::process::Child, event: isize, path: &Path) {
    let waited_on = HANDLE(event as *mut std::ffi::c_void);
    let started = std::time::Instant::now();

    while started.elapsed() < READY_WAIT {
        // A wait on the event is the whole of this loop's cadence: the child sets it the moment its
        // window is drawn, so the pin stands for as short a time as the load takes and no shorter.
        let waited = unsafe { WaitForSingleObject(waited_on, READY_POLL.as_millis() as u32) };

        if waited == windows::Win32::Foundation::WAIT_OBJECT_0 {
            if crate::ui::preview_window::pinned_path().as_deref() == Some(path) {
                // The end this brings is a hand-off rather than a close: the hand asked for a window
                // of its own, and the file under the pointer is not a file it has just moved onto
                // (see `note_the_pin_was_handed_off`).
                crate::ui::preview_window::note_the_pin_was_handed_off();
                crate::ui::preview_window::request_pin_end();
            }
            break;
        }

        // A child that has gone without answering is a window there is nothing to hand over to:
        // the pin it came from stays, which is the answer the user can still act on.
        if !matches!(child.try_wait(), Ok(None)) {
            break;
        }
    }

    unsafe {
        let _ = CloseHandle(waited_on);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn arguments(line: &[&str]) -> Vec<OsString> {
        line.iter().map(OsString::from).collect()
    }

    fn a_request(line: &[&str]) -> OpenRequest {
        match parse(arguments(line)) {
            Startup::Detached(request) => request,
            other => panic!("not a detached request: {other:?}"),
        }
    }

    #[test]
    fn a_launch_with_no_switch_of_ours_on_it_is_the_tray_app() {
        assert_eq!(parse(arguments(&[])), Startup::Tray);
        assert_eq!(parse(arguments(&["--verbose"])), Startup::Tray);
    }

    #[test]
    fn a_hand_off_names_the_file_the_box_and_whether_the_window_was_maximized() {
        let request = a_request(&[
            "--detach",
            r"C:\pictures\My Pictures\a file.png",
            "--box",
            "100,200,900,700",
            "--maximized",
            "--ready",
            "Local\\rust-hover-preview-detach-11-1",
            "--parent",
            "4242",
        ]);

        assert_eq!(
            request.path,
            PathBuf::from(r"C:\pictures\My Pictures\a file.png")
        );
        assert_eq!(request.region, (100, 200, 900, 700));
        assert!(request.maximized);
        assert_eq!(
            request.ready.as_deref(),
            Some("Local\\rust-hover-preview-detach-11-1")
        );
        assert_eq!(request.parent, Some(4242));
    }

    #[test]
    fn a_name_outside_this_machines_code_page_survives_the_hand_off() {
        let request = a_request(&["--detach", "C:\\图片\\写真.png", "--box", "0,0,10,10"]);
        assert_eq!(request.path, PathBuf::from("C:\\图片\\写真.png"));
    }

    #[test]
    fn a_request_with_no_file_or_no_box_begins_nothing() {
        // Neither half.
        assert_eq!(parse(arguments(&["--detach"])), Startup::Refused);
        // A file with no box.
        assert_eq!(
            parse(arguments(&["--detach", "C:\\file.png"])),
            Startup::Refused
        );
        // A box with no file.
        assert_eq!(
            parse(arguments(&["--detach", "--box", "0,0,10,10"])),
            Startup::Refused
        );
        // An empty name where a file should be.
        assert_eq!(
            parse(arguments(&["--detach", "", "--box", "0,0,10,10"])),
            Startup::Refused
        );
    }

    #[test]
    fn a_box_that_is_not_a_box_is_not_a_window() {
        for written in [
            "0,0,0,10",
            "0,0,10,0",
            "10,10,0,0",
            "0,0,10",
            "0,0,10,10,10",
            "0,0,ten,10",
            "",
        ] {
            assert_eq!(
                parse(arguments(&["--detach", "C:\\file.png", "--box", written])),
                Startup::Refused,
                "{written} was read as a box"
            );
        }

        // Whitespace around the numbers is not the writing this app does, but it is not a mistake
        // worth standing on either.
        assert_eq!(
            a_request(&["--detach", "C:\\file.png", "--box", " 0, 0, 10, 10 "]).region,
            (0, 0, 10, 10)
        );
    }

    #[test]
    fn a_switch_this_build_does_not_know_is_not_an_instruction() {
        let request = a_request(&[
            "--detach",
            "C:\\file.png",
            "--box",
            "0,0,10,10",
            "--something-new",
        ]);
        assert_eq!(request.region, (0, 0, 10, 10));
        assert_eq!(request.parent, None);
        assert_eq!(request.ready, None);
        assert!(!request.maximized);
    }
}

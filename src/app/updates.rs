//! The one thing in this app that asks the network a question: whether a newer
//! release than the one running has been published, and — where one has — the
//! installer for it, fetched when the click that puts it on asks for it.
//!
//! The click that puts it on is a question with three answers, and only one of
//! them installs anything: `Auto` fetches the installer and runs it, `Manual`
//! opens the release page in the user's own browser instead, and `Cancel` is
//! nothing at all.
//!
//! What asks is the app's own start, so that the answer is waiting by the time
//! anyone looks for it, and the tray menu, because an opening is the one other
//! moment a user is looking for one: `show_context_menu` asks for a check as the
//! menu is built, the check runs on a thread of its own, and the row above `Run
//! at Startup` reports what the last one found. Nothing here is on a hover's
//! path, and a check that never happens costs a preview nothing.
//!
//! What is trusted is this repository's own releases. Both addresses are this
//! repository's own — the newest release, for the version, and the release that
//! version names, for its installer — and both are over HTTPS under the
//! machine's own certificate store and proxy settings, WinHTTP, so nothing is
//! bundled for either; and an installer that arrives is checked against the
//! length the response promised, and for the two bytes every Windows executable
//! opens with, before it is run. It is run with the installer's own silent
//! switch, which replaces this app, and with the switch that starts it again
//! afterwards. The app ends itself as it hands over, so the copy the installer
//! has to terminate is one that is already leaving.

use crate::app::dialogs;
use crate::RUNNING;
use once_cell::sync::Lazy;
use std::ffi::c_void;
use std::fs::{self, File};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Mutex;
use std::time::{Duration, Instant};
use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::HWND;
use windows::Win32::Networking::WinHttp::{
    WinHttpCloseHandle, WinHttpConnect, WinHttpOpen, WinHttpOpenRequest, WinHttpQueryDataAvailable,
    WinHttpQueryHeaders, WinHttpReadData, WinHttpReceiveResponse, WinHttpSendRequest,
    WinHttpSetTimeouts, WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY, WINHTTP_FLAG_SECURE,
    WINHTTP_QUERY_CONTENT_LENGTH, WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_QUERY_FLAG_NUMBER64,
    WINHTTP_QUERY_STATUS_CODE,
};
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::Shell::ShellExecuteW;
use windows::Win32::UI::WindowsAndMessaging::{
    MessageBoxW, SetWindowsHookExW, UnhookWindowsHookEx, IDNO, IDYES, MB_ICONINFORMATION,
    MB_ICONWARNING, MB_OK, MB_SETFOREGROUND, MB_YESNOCANCEL, SW_SHOWNORMAL, WH_CBT,
};

/// Where this app's releases are published, and the two files a check asks for.
/// They are the releases' *own* assets rather than a query against the releases
/// API: the version is read at `releases/latest/download`, GitHub's stable
/// address for an asset of the newest release, and the installer is asked for in
/// the release that version names — by the tag it is published under, and by the
/// name the packager gave the file there — so the check is two addresses built
/// here, the answers are files the releases already carry, and nothing has to
/// parse a response format or count against an API rate limit.
const RELEASE_HOST: &str = "github.com";
const LATEST_PATH: &str = "/cosmokud/rust-hover-preview/releases/latest/download";
const RELEASE_PATH: &str = "/cosmokud/rust-hover-preview/releases/download";
const VERSION_ASSET: &str = "version.txt";

/// The page a `Manual` answer opens: the release the version is read from, which
/// is the update the row is offering — `latest` without the asset path the
/// version is read at. It is the one address here that is handed to something
/// other than this app's own request, and the only one that is opened rather than
/// fetched: the browser the user already has, and nothing of this app runs on the
/// way.
const LATEST_PAGE: &str = "/cosmokud/rust-hover-preview/releases/latest";

/// What the installer is called where it is fetched to, which is this app's own
/// name for it: the file it was fetched from is named for the version it carries,
/// and nothing outside this app reads the name it is written under.
const INSTALLER_FILE: &str = "rust-hover-preview-setup.exe";

/// What this app calls itself on the wire. GitHub answers a request without one
/// with a refusal, and a version in it is what tells a release page's own
/// statistics apart from a browser's.
const USER_AGENT: &str = concat!("RustHoverPreview/", env!("CARGO_PKG_VERSION"));

/// How often the check may be answered, however many times it is asked for. An
/// hour is often enough for a release, and an opening of the menu is not a
/// reason to ask GitHub anything: the check is skipped while the last one is
/// within this, and the hour is counted however that one ended, so a machine
/// that is offline does not ask again on the next menu. Nothing about it is
/// written down — a run has no record of the run before it — so the check a run
/// makes as it starts is always its first.
const CHECK_INTERVAL: Duration = Duration::from_secs(60 * 60);

/// Nothing smaller than this is this app's installer — the smallest one ever
/// published is a few megabytes — so a truncated download is refused rather
/// than run.
const MIN_INSTALLER_BYTES: u64 = 256 * 1024;

const READ_CHUNK_BYTES: u32 = 64 * 1024;

/// The version of the newest release, where it is newer than the one running:
/// what the menu row is built from, and the whole of what a check keeps. Nothing
/// is fetched for it — the installer is the click's business — so this is a
/// version on offer rather than an update waiting to be put on.
static OFFER: Lazy<Mutex<Option<String>>> = Lazy::new(|| Mutex::new(None));

/// When this run last made a check, which is the whole of what the interval is
/// counted from. It goes with the run that made it: a run that has just started
/// has asked nothing yet, and what some earlier run found is no reason for this
/// one to wait a moment before asking.
static LAST_CHECK: Lazy<Mutex<Option<Instant>>> = Lazy::new(|| Mutex::new(None));

/// Whether a check is running, so that the menu being opened twice in a row
/// does not start two of them.
static CHECKING: AtomicBool = AtomicBool::new(false);

/// Ask for a check, without waiting for one. The app's own start is one caller:
/// a run that has just started has asked nothing, and the answer is worth having
/// before anyone opens the menu for it. The tray menu is the other: an opening
/// is the one moment a user is looking for an update, and the one moment the
/// answer is worth having again.
pub(crate) fn request_check() {
    if CHECKING.swap(true, Ordering::SeqCst) {
        return;
    }

    std::thread::spawn(|| {
        check();
        CHECKING.store(false, Ordering::SeqCst);
    });
}

/// The version of the update on offer, where one is: what the menu row is built
/// from, and `None` where there is nothing to say.
pub(crate) fn available() -> Option<String> {
    OFFER.lock().ok().and_then(|offer| offer.clone())
}

/// Put the update on, once the user has said so — which is the `Auto` answer and
/// nothing else. The installer is fetched here rather than ahead of the check that
/// found it: a click is what a download of a few megabytes is worth, and a row
/// nobody clicked costs the network nothing. The installer then runs silently —
/// `/S` is its own switch for that, and `/R`, which only a silent installer reads,
/// is what starts the app again once the new version is in place. `false` is
/// answered where there is nothing on offer and where nothing could be fetched;
/// the two are one answer to the caller, and only the last is one the user is told
/// about.
pub(crate) fn install() -> bool {
    let Some(version) = available() else {
        return false;
    };

    let Some(installer) = fetch_installer(&version) else {
        download_failed(&version);
        return false;
    };

    std::process::Command::new(installer)
        .arg("/S")
        .arg("/R")
        .spawn()
        .is_ok()
}

/// The other way to take the update, which is the `Manual` answer: the release
/// page, opened in the browser the user already has. Nothing is fetched and
/// nothing is run here — the page is where the installer is, and taking it from
/// there is the user's own business — so the app is left exactly where it was.
pub(crate) fn open_release_page() {
    let url = wide(&format!("https://{RELEASE_HOST}{LATEST_PAGE}"));

    unsafe {
        let _ = ShellExecuteW(
            HWND::default(),
            w!("open"),
            PCWSTR(url.as_ptr()),
            PCWSTR::null(),
            PCWSTR::null(),
            SW_SHOWNORMAL,
        );
    }
}

/// What the click on the update row is answered with: the buttons of the dialog,
/// left to right, and the three things that can be done with the update.
pub(crate) enum Answer {
    /// Put the update on, which is what the row did before it had three answers.
    Auto,
    /// Take the update by hand: the release page, in the browser.
    Manual,
    /// Nothing, and a menu that stays as it was.
    Cancel,
}

/// Ask, and answer with what the user said. It stands where it does because the
/// click it follows may end the app: what a user is agreeing to with `Auto` is an
/// update that puts itself on, and a window that comes back as the new version,
/// while `Manual` is the release page in their browser and `Cancel` is a click
/// that does nothing. The app's other dialog, the one about a download that
/// failed, is only ever reached past `Auto`.
///
/// The dialog is given no owner, and is set to the foreground, for the same
/// reason: the one window this app owns is the tray's, which is never shown and
/// has no place on screen for a dialog to be centred over.
///
/// The question is asked with a message box — the platform's own dialog for one —
/// which carries the three buttons this question needs but names two of them for
/// a question this app is not asking; see `dialogs::hook` for what is done about
/// that. A click with nothing on offer is answered `Cancel` without a dialog at
/// all, which is the same click that does nothing.
pub(crate) fn ask() -> Answer {
    let Some(version) = available() else {
        return Answer::Cancel;
    };

    let caption = wide("Rust Hover Preview");
    let text = wide(&format!(
        "Version {version} is available.\n\n\
         Auto: Download and install now, then restart the app on the new version.\n\
         Manual: Open the release page in your browser."
    ));

    dialogs::set_button_names("Auto", "Manual");
    dialogs::begin();

    let hook =
        unsafe { SetWindowsHookExW(WH_CBT, Some(dialogs::hook), None, GetCurrentThreadId()) };

    let answer = unsafe {
        MessageBoxW(
            HWND::default(),
            PCWSTR(text.as_ptr()),
            PCWSTR(caption.as_ptr()),
            MB_YESNOCANCEL | MB_ICONINFORMATION | MB_SETFOREGROUND,
        )
    };

    if let Ok(hook) = hook {
        unsafe {
            let _ = UnhookWindowsHookEx(hook);
        }
    }

    dialogs::end();

    if answer == IDYES {
        Answer::Auto
    } else if answer == IDNO {
        Answer::Manual
    } else {
        Answer::Cancel
    }
}

/// What a click that asked for the update and could not have it is told. It is
/// the app's only dialog about something going wrong: without it, a download that
/// failed would be a click that did nothing, and the row it was clicked on is
/// still there to be clicked again.
fn download_failed(version: &str) {
    let caption = wide("Rust Hover Preview");
    let text = wide(&format!(
        "Version {version} could not be downloaded.\n\nNothing has changed. Check your \
         connection and try again."
    ));

    unsafe {
        MessageBoxW(
            HWND::default(),
            PCWSTR(text.as_ptr()),
            PCWSTR(caption.as_ptr()),
            MB_OK | MB_ICONWARNING | MB_SETFOREGROUND,
        );
    }
}

/// One check, whole: the version the release is published under, where that is
/// newer than the one running, and the time the check is noted as having
/// happened.
fn check() {
    if !due_for_check() {
        return;
    }

    // What a check comes back with is a version and nothing else: the installer
    // for it is the click's business, so a row nobody has clicked costs the
    // network nothing at all.
    if let Some(version) = newer_release() {
        if let Ok(mut offer) = OFFER.lock() {
            *offer = Some(version);
        }
    }

    // Noted however this one ended: a check that found nothing, and one that
    // could not reach GitHub, are both checks that were made, and asking again on
    // the next opening of the menu is the spam the interval is there to prevent.
    if let Ok(mut last) = LAST_CHECK.lock() {
        *last = Some(Instant::now());
    }
}

/// The version the newest release is published under, where it is newer than
/// this build. A release that carries no version — every release published
/// before this app asked for one — answers with nothing, which is the same
/// answer as a machine that is offline, and the same as an up-to-date one: the
/// menu has nothing to add.
fn newer_release() -> Option<String> {
    let body =
        request(RELEASE_HOST, &format!("{LATEST_PATH}/{VERSION_ASSET}"))?.read_all(4 * 1024)?;
    let text = String::from_utf8(body).ok()?;
    let version = text.trim();

    is_newer(version, env!("CARGO_PKG_VERSION")).then(|| version.to_owned())
}

/// Whether a release is one to offer: a version this app reads, newer than the
/// one running. What is running is read as the release it names and whether it
/// is a pre-release of it — a build made from a pre-release tag, `0.3.4-rc.1`,
/// is the release `0.3.4` before it — so the stable release that ends a
/// pre-release cycle is newer than it and brings it back onto the stable line,
/// while the release before that one is not offered as a step backwards. A
/// running version nothing can be read from is read as older than every release
/// there is, so it is offered whatever is published rather than being left where
/// it is.
fn is_newer(published: &str, running: &str) -> bool {
    let Some(published) = parse_version(published) else {
        return false;
    };

    let (release, pre_release) = match running.split_once('-') {
        Some((release, _)) => (parse_version(release).unwrap_or((0, 0, 0)), true),
        None => (parse_version(running).unwrap_or((0, 0, 0)), false),
    };

    published > release || (published == release && pre_release)
}

/// The address of one release's installer, which is the release the version
/// names rather than the newest one — the version is read before the download
/// for exactly this reason. What is asked for is the file the packager built,
/// under the name it gave it there (`<binary>_<version>_<arch>-setup.exe`), and
/// a release that carries no such file answers `404`, which is a download that
/// failed: the click is told so, and the row it was clicked on stays where it was.
fn installer_path(version: &str) -> String {
    format!("{RELEASE_PATH}/v{version}/rust-hover-preview_{version}_x64-setup.exe")
}

/// Fetch the installer for a release, answering with the file it is in, ready to
/// be run. It is written where the app's other transient files go — the folder a
/// page render in flight is written to, which a startup clears out — because
/// nothing waits on it once it has been run.
fn fetch_installer(version: &str) -> Option<PathBuf> {
    let directory = installer_dir();
    fs::create_dir_all(&directory).ok()?;

    let installer = directory.join(INSTALLER_FILE);

    if fetch_installer_body(&installer, version).is_none() {
        // A body that failed is a file that would never be run, and it has no
        // reason to be left where the next click would find it.
        let _ = fs::remove_file(&installer);
        return None;
    }

    Some(installer)
}

/// The body of the installer, written to `path`, answering only where what
/// arrived is a whole file: the length the response promised, a size this app's
/// installer has, and the two bytes a Windows executable opens with — which
/// together are what tells an installer apart from a truncated download and
/// from the page a release that carries no such asset answers with.
fn fetch_installer_body(path: &Path, version: &str) -> Option<u64> {
    let mut file = File::create(path).ok()?;
    let written = request(RELEASE_HOST, &installer_path(version))?.read_into(&mut file)?;
    file.flush().ok()?;

    (written >= MIN_INSTALLER_BYTES && starts_with_mz(path)).then_some(written)
}

/// Where the installer is written: this app's own folder under the temp folder,
/// the same one a page render in flight is written to, and one a startup clears
/// out — so a download that was ended mid-flight is cleared away with it.
fn installer_dir() -> PathBuf {
    std::env::temp_dir().join("rust-hover-preview")
}

/// Whether the last check is far enough behind to ask again. A run with no
/// record of one — which is every run, until it makes its first — is due.
fn due_for_check() -> bool {
    LAST_CHECK
        .lock()
        .ok()
        .and_then(|last| *last)
        .is_none_or(|last| last.elapsed() >= CHECK_INTERVAL)
}

/// Remove what earlier versions left for the check: the time they wrote down of
/// when they last asked, and the installer they fetched before any click. Neither
/// is kept here any more — the hour is the run's own memory, and the installer is
/// fetched for the click that asks for it — so both go at startup, with the other
/// files earlier versions left behind.
pub(crate) fn discard_old_files() {
    let Some(dirs) = directories::BaseDirs::new() else {
        return;
    };

    let directory = dirs
        .data_local_dir()
        .join("rust-hover-preview")
        .join("update");

    let _ = fs::remove_file(directory.join("last-check"));
    let _ = fs::remove_file(directory.join(INSTALLER_FILE));
    let _ = fs::remove_file(directory.join(format!("{INSTALLER_FILE}.part")));

    // Only where those three were the whole of it: the folder is nothing this
    // version writes, and an empty one is a folder a user need not have.
    let _ = fs::remove_dir(&directory);
}

/// A version as the three numbers it is made of, which is what the tags this
/// repository publishes are (`v0.3.0`), and what a comparison of two of them
/// needs. Anything else — a tag with a suffix, a line that is not a version, a
/// body that came back from something other than the release — answers with
/// nothing, which reads as "no update" rather than as an update to a version
/// nothing can be compared against.
fn parse_version(text: &str) -> Option<(u32, u32, u32)> {
    let text = text.trim();
    let text = text.strip_prefix('v').unwrap_or(text);

    let mut parts = text.split('.');
    let major = parts.next()?.parse().ok()?;
    let minor = parts.next()?.parse().ok()?;
    let patch = parts.next()?.parse().ok()?;

    parts.next().is_none().then_some((major, minor, patch))
}

/// Whether a file opens with the two bytes every Windows executable does. It is
/// not a signature and does not pretend to be one — what makes the download
/// trustworthy is the HTTPS connection it came over, from this repository's own
/// release — but it does catch a download that is a page rather than a program.
fn starts_with_mz(path: &Path) -> bool {
    let mut signature = [0u8; 2];
    File::open(path)
        .and_then(|mut file| file.read_exact(&mut signature))
        .is_ok()
        && &signature == b"MZ"
}

fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}

/// One WinHTTP handle, closed when it goes out of scope. A request is three of
/// them — the session, the connection to the host, and the request itself — and
/// they are opened in that order inside one function, so a step that fails
/// closes the ones before it and nothing is left open.
struct Handle(*mut c_void);

impl Handle {
    fn new(raw: *mut c_void) -> Option<Self> {
        (!raw.is_null()).then_some(Handle(raw))
    }
}

impl Drop for Handle {
    fn drop(&mut self) {
        unsafe {
            let _ = WinHttpCloseHandle(self.0);
        }
    }
}

/// A response that is ready to be read: the request to read it from, the
/// handles it hangs off, and the length the server said the body would be,
/// where it said.
///
/// The whole chain is held rather than just the request, and in the order it is
/// closed in: a connection, or a session, closed while the request under it is
/// still being read is a body that fails halfway with
/// `ERROR_WINHTTP_CONNECTION_ERROR` — which is exactly what a first read does
/// when the handles that opened it have gone out of scope.
struct Reply {
    request: Handle,
    _connect: Handle,
    _session: Handle,
    length: Option<u64>,
}

/// Ask one host for one path, and answer with the response only where the
/// server said the file was there. A redirect — which is how
/// `releases/latest/download` reaches the release it names — is followed by
/// WinHTTP itself, and the headers read here are the ones the response ended up
/// with; a `404`, which is what a release published before this app asked for
/// these names looks like, is answered with nothing and is not an error.
fn request(host: &str, path: &str) -> Option<Reply> {
    let agent = wide(USER_AGENT);
    let host = wide(host);
    let target = wide(path);

    unsafe {
        // The automatic proxy is the one a user configures in Windows, PAC
        // script and all, rather than the machine-wide WinHTTP setting an
        // explicit proxy would name: a check is a request on the user's behalf
        // and belongs on their connection.
        let session = Handle::new(WinHttpOpen(
            PCWSTR(agent.as_ptr()),
            WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY,
            PCWSTR::null(),
            PCWSTR::null(),
            0,
        ))?;

        // Nothing waits longer than this for any part of it: a check that is
        // slow is a check nobody asked to watch, and the next opening of the
        // menu asks again anyway.
        let _ = WinHttpSetTimeouts(session.0, 5_000, 5_000, 5_000, 10_000);

        let connect = Handle::new(WinHttpConnect(session.0, PCWSTR(host.as_ptr()), 443, 0))?;

        let request = Handle::new(WinHttpOpenRequest(
            connect.0,
            w!("GET"),
            PCWSTR(target.as_ptr()),
            PCWSTR::null(),
            PCWSTR::null(),
            std::ptr::null(),
            WINHTTP_FLAG_SECURE,
        ))?;

        WinHttpSendRequest(request.0, None, None, 0, 0, 0).ok()?;
        WinHttpReceiveResponse(request.0, std::ptr::null_mut()).ok()?;

        let mut status = 0u32;
        let mut size = std::mem::size_of::<u32>() as u32;
        let mut index = 0u32;
        WinHttpQueryHeaders(
            request.0,
            WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
            PCWSTR::null(),
            Some(&mut status as *mut u32 as *mut c_void),
            &mut size,
            &mut index,
        )
        .ok()?;

        if status != 200 {
            return None;
        }

        let mut length = 0u64;
        let mut size = std::mem::size_of::<u64>() as u32;
        let mut index = 0u32;
        let length = WinHttpQueryHeaders(
            request.0,
            WINHTTP_QUERY_CONTENT_LENGTH | WINHTTP_QUERY_FLAG_NUMBER64,
            PCWSTR::null(),
            Some(&mut length as *mut u64 as *mut c_void),
            &mut size,
            &mut index,
        )
        .is_ok()
        .then_some(length);

        Some(Reply {
            request,
            _connect: connect,
            _session: session,
            length,
        })
    }
}

impl Reply {
    /// The whole body, where it is no longer than the caller allows. Both files
    /// this app asks for are known sizes — a version is a line, an installer is
    /// megabytes — so the cap is what keeps an answer that is neither from being
    /// read into memory at all.
    fn read_all(&self, cap: usize) -> Option<Vec<u8>> {
        let mut body = Vec::new();
        let mut chunk = vec![0u8; 4 * 1024];

        loop {
            let read = self.read_chunk(&mut chunk)?;

            if read == 0 {
                return Some(body);
            }

            if body.len() + read > cap {
                return None;
            }

            body.extend_from_slice(&chunk[..read]);
        }
    }

    /// The whole body, written where the caller says, and the count of what was
    /// written. A body that stopped short of the length the response promised
    /// is refused rather than handed on, which is the one way a download this
    /// side can see failing without a status code saying so.
    fn read_into(&self, file: &mut File) -> Option<u64> {
        let mut written = 0u64;
        let mut chunk = vec![0u8; READ_CHUNK_BYTES as usize];

        loop {
            // A download is the one thing here that can outlast the app: an
            // exit is not a reason to keep reading megabytes for a menu that
            // is already gone.
            if !RUNNING.load(Ordering::Acquire) {
                return None;
            }

            let read = self.read_chunk(&mut chunk)?;

            if read == 0 {
                break;
            }

            file.write_all(&chunk[..read]).ok()?;
            written += read as u64;
        }

        match self.length {
            Some(length) if length != written => None,
            _ => Some(written),
        }
    }

    /// One read of whatever the connection has, waiting for it when it has
    /// nothing yet. A finished body is a read of nothing, which is what ends
    /// the two loops above.
    fn read_chunk(&self, chunk: &mut [u8]) -> Option<usize> {
        let mut available = 0u32;

        unsafe {
            WinHttpQueryDataAvailable(self.request.0, &mut available).ok()?;

            if available == 0 {
                return Some(0);
            }

            let mut read = 0u32;
            let wanted = available.min(chunk.len() as u32);

            WinHttpReadData(
                self.request.0,
                chunk.as_mut_ptr() as *mut c_void,
                wanted,
                &mut read,
            )
            .ok()?;

            Some(read as usize)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_version_is_the_three_numbers_it_is_made_of() {
        assert_eq!(parse_version("0.2.14"), Some((0, 2, 14)));
        assert_eq!(parse_version("v0.3.0"), Some((0, 3, 0)));
        assert_eq!(parse_version(" 1.10.2\n"), Some((1, 10, 2)));
        assert!(parse_version(env!("CARGO_PKG_VERSION")).is_some());
    }

    #[test]
    fn anything_else_is_not_a_version() {
        assert_eq!(parse_version("0.3"), None);
        assert_eq!(parse_version("0.3.0-beta"), None);
        assert_eq!(parse_version("main"), None);
        assert_eq!(parse_version(""), None);
        assert_eq!(parse_version("v"), None);
    }

    #[test]
    fn a_release_is_newer_only_when_it_is() {
        let newer = |published: &str| parse_version(published) > parse_version("0.2.14");

        assert!(newer("0.2.15"));
        assert!(newer("0.3.0"));
        assert!(newer("1.0.0"));
        assert!(!newer("0.2.14"));
        assert!(!newer("0.2.13"));
    }

    /// A build made from a pre-release tag is older than the release its own
    /// version names, which is how it comes back to the stable line: the `0.3.4`
    /// that ends the `0.3.4-rc.1` cycle is offered to it, and so is anything
    /// later, while the `0.3.3` before it is not — that would be a step
    /// backwards rather than a way back.
    #[test]
    fn a_pre_release_build_is_brought_back_onto_the_release_it_names() {
        assert!(is_newer("0.3.4", "0.3.4-rc.1"));
        assert!(is_newer("0.3.5", "0.3.4-rc.1"));
        assert!(!is_newer("0.3.3", "0.3.4-rc.1"));
        assert!(!is_newer("0.3.4-rc.2", "0.3.4-rc.1"));

        assert!(!is_newer("0.3.4", "0.3.4"));
        assert!(is_newer("0.3.4", "main"));
    }

    /// The installer is asked for in the release its version names, under the
    /// name the packager gave that file there. It is the one address the deploy
    /// workflow and this app have to agree on, since a release carries no second
    /// copy of the installer under a name of its own.
    #[test]
    fn the_installer_is_addressed_by_the_release_and_its_own_name() {
        assert_eq!(
            installer_path("0.3.0"),
            "/cosmokud/rust-hover-preview/releases/download/v0.3.0/\
             rust-hover-preview_0.3.0_x64-setup.exe"
        );
    }

    /// The host answers, and a redirect is followed to the release it names —
    /// which is the request half of every check, and the half a machine with no
    /// connection fails on. Run it by hand: `cargo test -- --ignored`.
    #[test]
    #[ignore = "asks the network"]
    fn a_release_page_answers() {
        assert!(request(RELEASE_HOST, "/cosmokud/rust-hover-preview/releases/latest").is_some());
    }

    /// The version asset, where a release carries one. It passes with nothing
    /// to check against a release published before the workflow wrote
    /// `version.txt` — which is the answer the check itself reads as "no
    /// update", and the same answer an offline machine gets.
    #[test]
    #[ignore = "asks the network"]
    fn the_published_version_is_a_version() {
        let body = request(RELEASE_HOST, &format!("{LATEST_PATH}/{VERSION_ASSET}"))
            .and_then(|reply| reply.read_all(4 * 1024));

        if let Some(body) = body {
            let text = String::from_utf8_lossy(&body);
            assert!(
                parse_version(&text).is_some(),
                "published version: {text:?}"
            );
        }
    }

    /// The other half — the download — against a body that takes more than one
    /// read, from a host that is not this repository's. `Cargo.toml` stands in
    /// for an installer: what is being checked is that a body arrives whole
    /// through `read_into`, which is the call the installer is fetched with,
    /// and that the length it was promised is the length it got.
    #[test]
    #[ignore = "asks the network"]
    fn a_body_arrives_whole() {
        let reply = request(
            "raw.githubusercontent.com",
            "/cosmokud/rust-hover-preview/main/Cargo.toml",
        )
        .expect("the manifest should be there");

        let path = std::env::temp_dir().join("rust-hover-preview-update-probe.toml");
        let mut file = File::create(&path).unwrap();
        let written = reply
            .read_into(&mut file)
            .expect("the body should arrive whole");
        drop(file);

        let body = fs::read_to_string(&path).unwrap();
        let _ = fs::remove_file(&path);

        assert!(written > 1024, "written: {written}");
        assert!(body.contains("rust-hover-preview"), "body: {body:.80}");
    }
}

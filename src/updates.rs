//! The one thing in this app that asks the network a question: whether a newer
//! release than the one running has been published, and — where one has — the
//! installer for it, fetched and kept ready for the click that puts it on.
//!
//! What asks is the tray menu, because that is the one moment a user is looking
//! for an answer: `show_context_menu` asks for a check as the menu is built, the
//! check runs on a thread of its own, and the row above `Run at Startup` reports
//! what the last one found. Nothing here runs at startup, nothing here is on a
//! hover's path, and a check that never happens costs a preview nothing.
//!
//! What is trusted is this repository's own releases. Both URLs are compiled in
//! and both are GitHub's, over HTTPS under the machine's own certificate store
//! and proxy settings — WinHTTP, so nothing is bundled for either — and an
//! installer that arrives is checked against the length the response promised,
//! and for the two bytes every Windows executable opens with, before it is kept.
//! It is then run with the installer's own silent switch, which replaces this
//! app, and with the switch that starts it again afterwards. The app ends itself
//! as it hands over, so the copy the installer has to terminate is one that is
//! already leaving.

use crate::RUNNING;
use once_cell::sync::Lazy;
use std::ffi::c_void;
use std::fs::{self, File};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Mutex;
use std::time::{Duration, SystemTime, UNIX_EPOCH};
use windows::core::{w, PCWSTR};
use windows::Win32::Networking::WinHttp::{
    WinHttpCloseHandle, WinHttpConnect, WinHttpOpen, WinHttpOpenRequest, WinHttpQueryDataAvailable,
    WinHttpQueryHeaders, WinHttpReadData, WinHttpReceiveResponse, WinHttpSendRequest,
    WinHttpSetTimeouts, WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY, WINHTTP_FLAG_SECURE,
    WINHTTP_QUERY_CONTENT_LENGTH, WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_QUERY_FLAG_NUMBER64,
    WINHTTP_QUERY_STATUS_CODE,
};

/// Where a release is published, and the two names this app asks for. They are
/// the release's *own* assets rather than a query against the releases API:
/// `releases/latest/download/<name>` is GitHub's stable address for an asset of
/// the newest release, so the check is a fixed pair of URLs, the answer is a
/// file the release already carries, and nothing here has to parse a response
/// format or count against an API rate limit. It is the deploy workflow that
/// puts both names in every release — a plain version, and the installer copied
/// out under a name that does not change from version to version.
const RELEASE_HOST: &str = "github.com";
const RELEASE_PATH: &str = "/cosmokud/rust-hover-preview/releases/latest/download";
const VERSION_ASSET: &str = "version.txt";
const INSTALLER_ASSET: &str = "rust-hover-preview-setup.exe";

/// What this app calls itself on the wire. GitHub answers a request without one
/// with a refusal, and a version in it is what tells a release page's own
/// statistics apart from a browser's.
const USER_AGENT: &str = concat!("RustHoverPreview/", env!("CARGO_PKG_VERSION"));

/// How often the check may be answered, however many times it is asked for. An
/// hour is often enough for a release, and an opening of the menu is not a
/// reason to ask GitHub anything: the check is skipped while the last one is
/// within this, and the timestamp is written however that one ended, so a
/// machine that is offline does not ask again on the next menu.
const CHECK_INTERVAL: Duration = Duration::from_secs(60 * 60);

/// Nothing smaller than this is this app's installer — the smallest one ever
/// published is a few megabytes — so a truncated download is refused rather
/// than offered.
const MIN_INSTALLER_BYTES: u64 = 256 * 1024;

const READ_CHUNK_BYTES: u32 = 64 * 1024;

/// The installer fetched for a release newer than the one running, waiting for
/// the click that puts it on.
struct Ready {
    version: String,
    installer: PathBuf,
}

static READY: Lazy<Mutex<Option<Ready>>> = Lazy::new(|| Mutex::new(None));

/// Whether a check is running, so that the menu being opened twice in a row
/// does not start two of them.
static CHECKING: AtomicBool = AtomicBool::new(false);

/// Ask for a check, without waiting for one. The tray menu is the only caller:
/// it is the one moment a user is looking for an update, and the one moment the
/// answer is worth having.
pub(crate) fn request_check() {
    if CHECKING.swap(true, Ordering::SeqCst) {
        return;
    }

    std::thread::spawn(|| {
        check();
        CHECKING.store(false, Ordering::SeqCst);
    });
}

/// The version of the update waiting to be installed, where one is: what the
/// menu row is built from, and `None` where there is nothing to say.
pub(crate) fn available() -> Option<String> {
    READY
        .lock()
        .ok()
        .and_then(|ready| ready.as_ref().map(|ready| ready.version.clone()))
}

/// Put the fetched update on. The installer runs silently — `/S` is its own
/// switch for that, and `/R`, which only a silent installer reads, is what
/// starts the app again once the new version is in place — and `false` is
/// answered when there is nothing fetched to run.
pub(crate) fn install() -> bool {
    let installer = READY
        .lock()
        .ok()
        .and_then(|ready| ready.as_ref().map(|ready| ready.installer.clone()));

    let Some(installer) = installer else {
        return false;
    };

    std::process::Command::new(installer)
        .arg("/S")
        .arg("/R")
        .spawn()
        .is_ok()
}

/// One check, whole: the version the release is published under, the installer
/// for it where that version is newer than this one, and the time the check is
/// written down as having happened.
fn check() {
    if !due_for_check() {
        return;
    }

    if let Some(version) = newer_release() {
        // An installer of the same release, or of a newer one, is already what
        // the menu is offering: fetching it again would be a download nothing
        // is waiting for.
        let already = available().map_or(false, |staged| {
            parse_version(&staged) >= parse_version(&version)
        });

        if !already {
            fetch_installer(&version);
        }
    }

    write_last_check();
}

/// The version the newest release is published under, where it is newer than
/// this build. A release that carries no version — every release published
/// before this app asked for one — answers with nothing, which is the same
/// answer as a machine that is offline, and the same as an up-to-date one: the
/// menu has nothing to add.
fn newer_release() -> Option<String> {
    let body =
        request(RELEASE_HOST, &format!("{RELEASE_PATH}/{VERSION_ASSET}"))?.read_all(4 * 1024)?;
    let text = String::from_utf8(body).ok()?;
    let version = text.trim();

    (parse_version(version) > parse_version(env!("CARGO_PKG_VERSION"))).then(|| version.to_owned())
}

/// Fetch the installer for a release and keep it where the click will find it.
/// It is written under a name of its own first and renamed into place once it
/// has been read whole, so what the menu offers is never half a file.
fn fetch_installer(version: &str) -> Option<()> {
    let directory = update_dir()?;
    fs::create_dir_all(&directory).ok()?;

    let installer = directory.join(INSTALLER_ASSET);
    let partial = directory.join(format!("{INSTALLER_ASSET}.part"));

    if fetch_installer_body(&partial).is_none() {
        // Nothing is left behind for the next check to find: a body that failed
        // is a file that would never be offered and has no reason to be kept.
        let _ = fs::remove_file(&partial);
        return None;
    }

    // A download of an older release already waiting here is replaced: what the
    // menu is offering is the newest one there is.
    fs::rename(&partial, &installer).ok()?;

    if let Ok(mut ready) = READY.lock() {
        *ready = Some(Ready {
            version: version.to_owned(),
            installer,
        });
    }

    Some(())
}

/// The body of the installer, written to `path`, answering only where what
/// arrived is a whole file: the length the response promised, a size this app's
/// installer has, and the two bytes a Windows executable opens with — which
/// together are what tells an installer apart from a truncated download and
/// from the page a release that carries no such asset answers with.
fn fetch_installer_body(path: &Path) -> Option<u64> {
    let mut file = File::create(path).ok()?;
    let written = request(RELEASE_HOST, &format!("{RELEASE_PATH}/{INSTALLER_ASSET}"))?
        .read_into(&mut file)?;
    file.flush().ok()?;

    (written >= MIN_INSTALLER_BYTES && starts_with_mz(path)).then_some(written)
}

/// What the check keeps, which is its own folder rather than the user's: the
/// time of the last check, and the installer a click is waiting on. It sits
/// with the app's other per-user folders, beside the browser profiles and the
/// engine records, and not among the settings a user edits.
fn update_dir() -> Option<PathBuf> {
    directories::BaseDirs::new()
        .map(|dirs| {
            dirs.data_local_dir()
                .join("rust-hover-preview")
                .join("update")
        })
        .or_else(|| {
            Some(
                std::env::temp_dir()
                    .join("rust-hover-preview")
                    .join("update"),
            )
        })
}

fn last_check_path() -> Option<PathBuf> {
    Some(update_dir()?.join("last-check"))
}

/// Whether the last check is far enough behind to ask again. A machine with no
/// record of one — the first menu after this feature arrives — is due.
fn due_for_check() -> bool {
    let Some(path) = last_check_path() else {
        return true;
    };

    fs::read_to_string(path)
        .ok()
        .and_then(|text| text.trim().parse::<u64>().ok())
        .map_or(true, |last| {
            now_secs().saturating_sub(last) >= CHECK_INTERVAL.as_secs()
        })
}

/// Write down that a check happened, whichever way it went. A check that found
/// nothing, and one that could not reach GitHub, are both checks that were
/// made: asking again on the next opening of the menu is the spam the interval
/// is there to prevent.
fn write_last_check() {
    let Some(directory) = update_dir() else {
        return;
    };

    let _ = fs::create_dir_all(&directory);
    let _ = fs::write(directory.join("last-check"), now_secs().to_string());
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

fn now_secs() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|since| since.as_secs())
        .unwrap_or(0)
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
        let body = request(RELEASE_HOST, &format!("{RELEASE_PATH}/{VERSION_ASSET}"))
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

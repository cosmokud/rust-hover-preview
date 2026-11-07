//! Installing a missing engine with `winget`, for the `Codecs → Install` rows.
//!
//! Four decisions are the whole of it, and each is one a reader would otherwise reverse:
//!
//! **It runs unelevated.** `winget` installs into the user's own profile — its shims in
//! `%LOCALAPPDATA%\Microsoft\WinGet\Links`, its packages under `…\WinGet\Packages` — so an elevated
//! run would install for the administrator that token belongs to and leave this user's rows exactly
//! as missing as they were. A UAC prompt raised by a package's own installer is expected and is the
//! installer's, not this app's.
//!
//! **The menu's thread never waits.** An install is minutes — LibreOffice's and calibre's installers
//! are the two slow ones — so it runs on a thread of its own, and what comes back is one `WM_APP`
//! message on the tray window rather than a blocked message pump.
//!
//! **One retry, and only over the sources.** A failure `winget` blames on its sources is answered by
//! `winget source update` and the same install once more: that is the per-user failure a stale index
//! causes, and it is what the README walks a user through. Nothing is removed and nothing is
//! retried with interactivity — a second prompt-driven install is a window nobody asked for.
//!
//! **The `PATH` is re-read from the registry before the re-probe.** `winget` broadcasts an
//! environment change, and a process that was already running does not see it: without this the
//! engine it has just installed is still looked for in a `PATH` that does not name it, and reads as
//! missing — the one answer installing it was for.

use crate::app::dialogs;
use crate::app::engine_processes::hidden_command;
use crate::formats::codecs::{self, Row};
use once_cell::sync::Lazy;
use std::os::windows::io::AsRawHandle;
use std::path::Path;
use std::process::{Child, Output, Stdio};
use std::sync::Mutex;
use std::time::Duration;
use windows::core::PCWSTR;
use windows::Win32::Foundation::{HANDLE, WAIT_TIMEOUT};
use windows::Win32::System::Registry::{
    RegCloseKey, RegEnumValueW, RegOpenKeyExW, RegQueryValueExW, HKEY, HKEY_CURRENT_USER,
    HKEY_LOCAL_MACHINE, KEY_READ, REG_EXPAND_SZ,
};
use windows::Win32::System::Threading::WaitForSingleObject;

/// How long one `winget install` is given. The two slow engines are minutes rather than seconds —
/// LibreOffice's installer alone is a download of a few hundred megabytes — and the wait is a
/// background thread's rather than a hover's, so it is generous. What outlasts it is killed and
/// reported rather than waited on.
const INSTALL_TIMEOUT: Duration = Duration::from_secs(20 * 60);

/// How long the `winget source update` that answers a source-flavoured failure is given: it is the
/// index being downloaded rather than a package, so it is minutes of the same kind but fewer.
const SOURCE_UPDATE_TIMEOUT: Duration = Duration::from_secs(5 * 60);

/// How much of winget's output a failed install shows: its errors are the last lines of it, and the
/// whole of a build log is not a box anybody reads.
const TAIL_LINES: usize = 12;

/// The two keys a process's environment is composed from, machine first and user over it.
const MACHINE_ENVIRONMENT: &str = r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment";
const USER_ENVIRONMENT: &str = "Environment";

/// The install in flight, by package id, so the menu can say which row is being installed and no
/// second one is started while it is (see `install`).
static IN_FLIGHT: Lazy<Mutex<Option<&'static str>>> = Lazy::new(|| Mutex::new(None));

/// Which engine is being installed right now, if one is: the `Codecs → Install` rows read it to
/// label the row and to keep every row unpickable while it is in flight.
pub(crate) fn in_flight() -> Option<&'static str> {
    IN_FLIGHT.lock().ok().and_then(|slot| *slot)
}

/// Note that the worker has finished with the install slot, which the tray window does when the
/// message the worker posted reaches it (`WM_TRAY_INSTALL_FINISHED`).
pub(crate) fn settle() {
    if let Ok(mut slot) = IN_FLIGHT.lock() {
        *slot = None;
    }
}

/// Install one engine of the `Codecs` rows, on a thread of its own.
///
/// Nothing is started where `winget` is not there — the rows are only offered where it is (see
/// `codecs::winget`) — and nothing is started while another install is in flight: one install at a
/// time is the whole of the guard, and there is never a second worker.
pub(crate) fn install(row: &Row) {
    let Some(id) = row.winget else {
        return;
    };

    let Some(winget) = codecs::winget() else {
        return;
    };

    {
        let Ok(mut slot) = IN_FLIGHT.lock() else {
            return;
        };

        if slot.is_some() {
            return;
        }

        *slot = Some(id);
    }

    let name = row.name;

    std::thread::spawn(move || {
        started(&winget, name, id);
        crate::shell::tray::note_install_finished();
    });
}

/// The worker: install, retry once over the sources, re-probe, and report what went wrong.
fn started(winget: &Path, name: &'static str, id: &'static str) {
    let (installed, said) = match attempt(winget, id) {
        // A failure that names the sources is answered once, by rebuilding them: that is the
        // per-user failure a stale index causes, and it is what the README walks a user through.
        // Removing a source is not part of it — that is the user's call.
        (false, said) if source_flavoured(&said) => {
            let _ = run(winget, &["source", "update"], SOURCE_UPDATE_TIMEOUT);
            attempt(winget, id)
        }
        answer => answer,
    };

    // Before the re-probe, and not after it: the shim `winget` just wrote is on a `PATH` this
    // process was not told about (see the module's own note).
    refresh_path();
    codecs::refresh();

    if installed {
        codecs::note_installed(id);
    } else {
        dialogs::report_install_failure(name, &tail(&said));
    }
}

/// One `winget install`, answered with whether it succeeded and with what it said.
fn attempt(winget: &Path, id: &str) -> (bool, String) {
    run(
        winget,
        &[
            "install",
            "-e",
            "--id",
            id,
            "--silent",
            "--accept-package-agreements",
            "--accept-source-agreements",
            "--disable-interactivity",
            "--source",
            "winget",
        ],
        INSTALL_TIMEOUT,
    )
}

/// One `winget` run, waited on for at most `timeout`: whether it succeeded, and stdout and stderr
/// together — which of the two `winget` writes an error to is not something to lean on.
fn run(winget: &Path, args: &[&str], timeout: Duration) -> (bool, String) {
    let child = hidden_command(winget)
        .args(args)
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn();

    let Ok(child) = child else {
        return (
            false,
            format!("winget could not be started: {}", winget.display()),
        );
    };

    let Some(output) = wait_bounded(child, timeout) else {
        return (
            false,
            "winget did not finish within the time it was given and was ended.".to_string(),
        );
    };

    let mut said = String::from_utf8_lossy(&output.stdout).into_owned();
    said.push_str(&String::from_utf8_lossy(&output.stderr));

    (output.status.success(), said)
}

/// Wait for a child, giving up after `timeout` — the same wait `video_probe` makes for a probe, for
/// a program whose deadline is measured in minutes: a child still running at the deadline is killed
/// before anything is read of it, and what it left behind is not read at all.
fn wait_bounded(mut child: Child, timeout: Duration) -> Option<Output> {
    let handle = HANDLE(child.as_raw_handle());
    let waited = unsafe { WaitForSingleObject(handle, timeout.as_millis() as u32) };

    if waited == WAIT_TIMEOUT {
        let _ = child.kill();
        let _ = child.wait();
        return None;
    }

    child.wait_with_output().ok()
}

/// Whether winget's own words say the sources are what failed — a code from the source half of its
/// `0x8a1500xx` range, or a line naming a source. Loose on purpose: what a false positive costs is
/// one `source update` and one more attempt, where a false negative costs the retry that would have
/// repaired a stale index.
fn source_flavoured(said: &str) -> bool {
    let said = said.to_ascii_lowercase();

    said.contains("source") || said.contains("0x8a150")
}

/// The end of what `winget` said, which is where its errors are.
fn tail(said: &str) -> String {
    let lines: Vec<&str> = said
        .lines()
        .map(str::trim_end)
        .filter(|line| !line.trim().is_empty())
        .collect();

    lines[lines.len().saturating_sub(TAIL_LINES)..].join("\n")
}

/// Point this process's `PATH` at the one the machine has now, machine values first and this user's
/// over them — the order Windows composes a process's `PATH` in.
///
/// `winget` writes its shims to `%LOCALAPPDATA%\Microsoft\WinGet\Links`, adds the folder to the
/// user's `Path` and broadcasts the change; a process that is already running is not told, and its
/// `PATH` keeps the folders it started with. Reading the two keys back is what makes the folder a
/// part of this process's search, and what makes the engine an install just put there a preview can
/// lean on without a restart.
///
/// Expanded from the values themselves, since a `REG_EXPAND_SZ` `Path` is written with
/// `%USERPROFILE%`-style names in it rather than with the names expanded.
fn refresh_path() {
    let machine = environment(HKEY_LOCAL_MACHINE, MACHINE_ENVIRONMENT);
    let user = environment(HKEY_CURRENT_USER, USER_ENVIRONMENT);

    let mut named: Vec<(String, String)> = machine.clone();
    named.extend(user.iter().cloned());

    let path_of = |values: &[(String, String)]| {
        values
            .iter()
            .find(|(name, _)| name.eq_ignore_ascii_case("Path"))
            .map(|(_, value)| expand(value, &named))
    };

    let composed: Vec<String> = [path_of(&machine), path_of(&user)]
        .into_iter()
        .flatten()
        .filter(|part| !part.trim().is_empty())
        .collect();

    if !composed.is_empty() {
        std::env::set_var("PATH", composed.join(";"));
    }
}

/// Every value of one environment key that holds text, as the name it is written under and what it
/// holds. Anything else — a number, a binary value — is not an environment variable and is skipped.
fn environment(root: HKEY, key: &str) -> Vec<(String, String)> {
    let mut values = Vec::new();
    let key: Vec<u16> = key.encode_utf16().chain(std::iter::once(0)).collect();

    unsafe {
        let mut handle = HKEY::default();

        if RegOpenKeyExW(root, PCWSTR(key.as_ptr()), 0, KEY_READ, &mut handle).is_err() {
            return values;
        }

        let mut index = 0;

        loop {
            let mut name = [0u16; 512];
            let mut length = name.len() as u32;

            if RegEnumValueW(
                handle,
                index,
                windows::core::PWSTR(name.as_mut_ptr()),
                &mut length,
                None,
                None,
                None,
                None,
            )
            .is_err()
            {
                break;
            }

            index += 1;

            let name = String::from_utf16_lossy(&name[..length as usize]);

            if let Some(value) = value_of(handle, &name) {
                values.push((name, value));
            }
        }

        let _ = RegCloseKey(handle);
    }

    values
}

/// One value of an environment key, read as text whatever size it is: a `REG_EXPAND_SZ` keeps its
/// names unexpanded, which is what the caller wants — the names are expanded against the whole key
/// rather than one value at a time.
fn value_of(handle: HKEY, name: &str) -> Option<String> {
    let name: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();

    unsafe {
        let mut kind = REG_EXPAND_SZ;
        let mut size: u32 = 0;

        let sized = RegQueryValueExW(
            handle,
            PCWSTR(name.as_ptr()),
            None,
            Some(&mut kind),
            None,
            Some(&mut size),
        )
        .is_ok()
            && size >= 2;

        if !sized {
            return None;
        }

        let mut buffer = vec![0u8; size as usize];
        let read = RegQueryValueExW(
            handle,
            PCWSTR(name.as_ptr()),
            None,
            Some(&mut kind),
            Some(buffer.as_mut_ptr()),
            Some(&mut size),
        )
        .is_ok();

        if !read || (kind != REG_EXPAND_SZ && kind != windows::Win32::System::Registry::REG_SZ) {
            return None;
        }

        Some(decode(&buffer[..size.min(buffer.len() as u32) as usize]))
    }
}

/// A counted registry string as text: UTF-16, ending at its terminator.
fn decode(bytes: &[u8]) -> String {
    let units: Vec<u16> = bytes
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .take_while(|unit| *unit != 0)
        .collect();

    String::from_utf16_lossy(&units)
}

/// `%name%` written out, from the environment keys being read and then from this process's own
/// environment. A name neither holds is left as it was written: an unexpanded name in a `PATH` is a
/// folder that is not there, which is what it was before this ran.
fn expand(value: &str, named: &[(String, String)]) -> String {
    let mut expanded = String::with_capacity(value.len());
    let mut rest = value;

    while let Some(open) = rest.find('%') {
        let Some(close) = rest[open + 1..].find('%').map(|at| at + open + 1) else {
            break;
        };

        let name = &rest[open + 1..close];

        expanded.push_str(&rest[..open]);

        let found = named
            .iter()
            .find(|(other, _)| other.eq_ignore_ascii_case(name))
            .map(|(_, value)| value.clone())
            .or_else(|| std::env::var(name).ok());

        match found {
            Some(value) => expanded.push_str(&value),
            None => expanded.push_str(&rest[open..=close]),
        }

        rest = &rest[close + 1..];
    }

    expanded.push_str(rest);
    expanded
}

#[cfg(test)]
mod tests {
    use super::*;

    /// `%name%` is what an environment value is written with, and what is written with it is what
    /// the `PATH` a probe reads is composed of: a name the keys hold is written out, a name they do
    /// not is left as it stands rather than dropped.
    #[test]
    fn a_path_value_is_read_with_its_names_written_out() {
        let named = vec![
            ("SystemRoot".to_string(), "S:\\Windows".to_string()),
            ("LOCALAPPDATA".to_string(), "S:\\appdata\\local".to_string()),
        ];

        assert_eq!(
            expand("%SystemRoot%\\system32;S:\\bin", &named),
            "S:\\Windows\\system32;S:\\bin"
        );
        assert_eq!(
            expand("%LOCALAPPDATA%\\Microsoft\\WinGet\\Links", &named),
            "S:\\appdata\\local\\Microsoft\\WinGet\\Links"
        );
        assert_eq!(
            expand("%NOT_A_NAME%\\bin", &named),
            "%NOT_A_NAME%\\bin",
            "a name no key holds is left as it was written"
        );
        assert_eq!(expand("S:\\bin", &named), "S:\\bin");
        assert_eq!(expand("%unterminated", &named), "%unterminated");
    }

    /// What is shown for an install that failed is the end of what winget said, which is where its
    /// errors are.
    #[test]
    fn the_output_tail_keeps_the_last_lines_that_say_anything() {
        let said: String = (1..=20)
            .map(|line| format!("line {line}\r\n"))
            .collect::<String>()
            + "0x8a15000f : no source\n";

        let shown = tail(&said);
        let lines: Vec<&str> = shown.lines().collect();

        assert_eq!(lines.len(), TAIL_LINES);
        assert_eq!(lines.last(), Some(&"0x8a15000f : no source"));
        assert_eq!(tail(""), "", "an install that said nothing shows nothing");
    }

    /// A failure over the sources is the one that is retried, and it is recognised by winget's own
    /// words: that retry is what repairs a stale index without elevation or a source being removed.
    #[test]
    fn a_failure_winget_blames_on_its_sources_is_the_one_that_is_retried() {
        assert!(source_flavoured(
            "Failed when searching source; results will not be included"
        ));
        assert!(source_flavoured("0x8a15000F : the source data is missing"));
        assert!(!source_flavoured(
            "Installer failed with exit code: 1603 (the package is already installed)"
        ));
    }
}

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
//! **The install is visible, and the retry lives in the batch script.** What runs is a batch script
//! written to `%TEMP%`, started in a console window of its own: the user watches `winget` download
//! and install, and the second attempt — `winget source update` and the same install once more,
//! the answer to the stale index the README walks a user through — is that same window's. Nothing
//! is removed and nothing is retried interactively. Capturing the output instead was the thing
//! this replaced: a silent, minutes-long download is a menu row saying `installing…` and a window
//! nobody can read, and the one time an install matters is the time it fails. The script is the
//! installer's engines page's script too (`packaging/nsis/installer.nsi`), so the two run the same
//! flags and the same retry — keeping them identical is the point of the text being one function.
//!
//! **The `PATH` is re-read from the registry before the re-probe.** `winget` broadcasts an
//! environment change, and a process that was already running does not see it: without this the
//! engine it has just installed is still looked for in a `PATH` that does not name it, and reads as
//! missing — the one answer installing it was for.

use crate::app::dialogs;
use crate::formats::codecs::{self, Row};
use once_cell::sync::Lazy;
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::Mutex;
use std::time::Duration;
use windows::core::PCWSTR;
use windows::Win32::System::Registry::{
    RegCloseKey, RegEnumValueW, RegOpenKeyExW, RegQueryValueExW, HKEY, HKEY_CURRENT_USER,
    HKEY_LOCAL_MACHINE, KEY_READ, REG_EXPAND_SZ,
};

/// `CREATE_NEW_CONSOLE`: the flag that makes the script's console a window of its own rather than
/// one borrowed from whatever started it — which, for a process with no console at all, there is
/// none of. `engine_processes::hidden_command` is the opposite flag and other callers' business.
const CREATE_NEW_CONSOLE: u32 = 0x0000_0010;

/// How often the worker looks for the exit code the script leaves beside itself. The window, and
/// with it the user, is the progress display: this only has to notice a file appear.
const POLL: Duration = Duration::from_millis(250);

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

/// The worker: run the script, re-probe, and report what the exit code says went wrong.
///
/// Nothing is reported where the window was closed before the script left its exit code: a window
/// closed that early takes the install with it, and there is no failure to hand the user that the
/// window did not already say.
fn started(winget: &Path, name: &'static str, id: &'static str) {
    let code = run(winget, id);

    // Before the re-probe, and not after it: the shim `winget` just wrote is on a `PATH` this
    // process was not told about (see the module's own note).
    refresh_path();
    codecs::refresh();

    match code {
        Some(0) => codecs::note_installed(id),
        Some(code) => dialogs::report_install_failure(name, code),
        None => {}
    }
}

/// The two files an install leaves in `%TEMP%`, under the package id: the script run in the
/// console window, and the `.rc` that carries the exit code back here — the same two paths the
/// installer's engines page writes, so the two can never be running different installs.
fn paths(id: &str) -> (PathBuf, PathBuf) {
    let stem = format!("rust-hover-preview-winget-{id}");
    let temp = std::env::temp_dir();

    (
        temp.join(format!("{stem}.cmd")),
        temp.join(format!("{stem}.rc")),
    )
}

/// The batch text, as one string from one function: the app runs it and the installer's engines
/// page writes it, and the parity between them is what the text being here buys.
///
/// The retry is inside it rather than in either caller, and so is the reason it is a console
/// window's business: `winget` prints why it failed, and only a window that stays up long enough
/// to be read says it. `%ERRORLEVEL%` is read after the `if` block rather than inside it, where
/// batch would have expanded it before the block ran. The redirect comes before the `echo`
/// because `echo %ERRORLEVEL%> file` is a handle-`0` redirect with no output at all the moment
/// the code is a single digit.
fn script(winget: &Path, id: &str, rc: &Path) -> String {
    let install = format!(
        " install -e --id {id} --source winget --silent --accept-package-agreements \
         --accept-source-agreements --disable-interactivity"
    );

    [
        "@echo off".to_string(),
        format!("title Rust Hover Preview - installing {id}"),
        format!("\"{}\"{install}", winget.display()),
        "if errorlevel 1 (".to_string(),
        "    echo.".to_string(),
        "    echo First attempt failed. Updating the winget sources and retrying once.".to_string(),
        "    echo.".to_string(),
        format!("    \"{}\" source update", winget.display()),
        format!("    \"{}\"{install}", winget.display()),
        ")".to_string(),
        format!("> \"{}\" echo %ERRORLEVEL%", rc.display()),
        "echo.".to_string(),
        "echo Finished. This window can be closed.".to_string(),
        "pause >nul".to_string(),
    ]
    .join("\r\n")
        + "\r\n"
}

/// Start the script in a console of its own, and answer with the exit code it left behind — or
/// nothing where the window went before it wrote one.
///
/// There is no deadline of this side's: the user is watching the install, closing the window is
/// the way to end it, and a timeout would end an install nobody wanted ended.
fn run(winget: &Path, id: &str) -> Option<i32> {
    let (path, rc) = paths(id);

    // A stale `.rc` would be read as this install's answer the moment it appeared.
    let _ = std::fs::remove_file(&rc);
    std::fs::write(&path, script(winget, id, &rc)).ok()?;

    let comspec = std::env::var_os("COMSPEC").unwrap_or_else(|| "cmd.exe".into());

    let mut child = Command::new(comspec)
        .arg("/c")
        .arg(&path)
        .creation_flags(CREATE_NEW_CONSOLE)
        .spawn()
        .ok()?;

    loop {
        // A code that does not read back yet is a file still being written, and one more look is
        // cheaper than a failure read off half of it.
        if let Ok(text) = std::fs::read_to_string(&rc) {
            if let Ok(code) = text.trim().parse() {
                return Some(code);
            }
        }

        if let Ok(Some(_)) = child.try_wait() {
            return None;
        }

        std::thread::sleep(POLL);
    }
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

    /// The script is the whole contract this app has with the installer's engines page, which
    /// writes the same text: the package, both attempts, the source, and where the exit code is
    /// left. Admin rights, elevation and interactivity are the package's own business.
    #[test]
    fn the_script_carries_the_package_both_installs_and_where_the_code_is_left() {
        let winget = Path::new("S:\\winget\\winget.exe");
        let rc = Path::new("S:\\temp\\rust-hover-preview-winget-Gyan.FFmpeg.rc");

        let script = script(winget, "Gyan.FFmpeg", rc);
        let written = script.to_ascii_lowercase();

        assert_eq!(written.matches("--id gyan.ffmpeg").count(), 2);
        assert_eq!(written.matches("--source winget").count(), 2);
        assert_eq!(
            written
                .matches("\"s:\\winget\\winget.exe\" install")
                .count(),
            2
        );
        assert!(
            written.contains("if errorlevel 1"),
            "the retry is the script's"
        );
        assert!(
            written.contains("source update"),
            "and so is the source update"
        );
        assert!(written.contains("s:\\temp\\rust-hover-preview-winget-gyan.ffmpeg.rc"));
        assert!(
            script.is_ascii(),
            "cmd reads a batch file in its own codepage"
        );
    }
}

//! What this app hands over to a program of the user's own — and, on an elevated run, the token it
//! is handed with.
//!
//! **A child inherits its parent's token, and an elevated token is an administrator's.** Every
//! program this app starts while it is elevated is therefore a program running with administrator
//! rights that nobody was prompted for: a browser, an editor, a media player, an archive tool — and
//! `winget`, which installs into the profile of whichever administrator the token belongs to rather
//! than into this user's own. That is not what any of those rows means: `Open in Default App` is
//! "look at this file in the program I chose", not "give that program my administrator's account".
//!
//! **The token to hand over is the Shell's own.** `explorer.exe` is the one process in the session
//! that is certainly the desk the user is sitting at, and it runs with that account's ordinary token:
//! duplicated from it, a program this app starts is started as the desktop user rather than as the
//! caller (see [`as_the_desktop_user`]). A path is not a program, so it is handed to `explorer.exe`
//! instead, whose own `ShellExecute` is what turns it into one — and which starts it with the same
//! ordinary token, whether the Explorer that performs it is the copy already running or a second one
//! this process begins (see [`open_as_the_desktop_user`]).
//!
//! **The Shell's automation object is not that road on this platform, measured.** `Shell.Application`
//! is the object the well-known "start this unelevated" trick names, and it is registered here with
//! an in-process server (`shell32.dll`) and **no** `LocalServer32`: `CoCreateInstance` with
//! `CLSCTX_LOCAL_SERVER` answers `REGDB_E_CLASSNOTREG` (`0x80040154`), so a call through it is a
//! hand-off that silently never happens. An in-process object would run with this process's own
//! token anyway. What is written here is the mechanism rather than the name: a token taken from the
//! process that is already at the desk.
//!
//! **It is a setting, and it starts off.** `elevate_hand_offs` — the tray's
//! `System → Elevate Hand-Offs (Warning)` row — is off by default, which means every hand-off here is
//! started as the desktop user; on, the app starts them itself, with the token it has. The warning in
//! that label is what the on state is: some program of a user's own really is one that only works as
//! an administrator, and there is no way for this app to know which.
//!
//! **A preview engine is not a hand-off and keeps this process's token.** `mpv`, `ffplay`,
//! LibreOffice, ImageMagick, calibre, PeaZip and f3d are started to be *looked at*, not used: the
//! window is composited into the hover, a hook DLL is written into the process, a pipe is read and a
//! job object holds it. Every one of those crossings needs the two processes to be of one integrity
//! level — an injected DLL cannot be loaded into a process of another, and a hand-written starter
//! with a duplicated token would have to carry the pipes, the job and the hook along with it — so an
//! engine is started by `app::engine_processes` as it always was. What is de-elevated is exactly what
//! leaves this app's hands for the user's.

use std::os::windows::process::CommandExt;

use windows::core::{PCWSTR, PWSTR};
use windows::Win32::Foundation::{CloseHandle, HANDLE};
use windows::Win32::Security::{
    DuplicateTokenEx, SecurityImpersonation, TokenPrimary, TOKEN_ALL_ACCESS, TOKEN_DUPLICATE,
    TOKEN_QUERY,
};
use windows::Win32::System::Threading::{
    CreateProcessWithTokenW, OpenProcess, OpenProcessToken, LOGON_WITH_PROFILE,
    PROCESS_CREATION_FLAGS, PROCESS_INFORMATION, PROCESS_QUERY_LIMITED_INFORMATION, STARTUPINFOW,
};
use windows::Win32::UI::Shell::{
    ShellExecuteExW, SEE_MASK_ASYNCOK, SEE_MASK_HMONITOR, SHELLEXECUTEINFOW,
};
use windows::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL;

/// Whether a hand-off should be started as the desktop user rather than with this process's token.
///
/// Two facts decide it, and both are read every time a hand-off is made: an app that is not elevated
/// has nothing to de-elevate, and a user who ticked `Elevate Hand-Offs (Warning)` has asked for the
/// opposite. Nothing is cached, because the two are one atomic read and a file read, and a hand-off
/// is a click rather than a frame.
pub(crate) fn through_the_shell() -> bool {
    if !crate::app::elevation::is_elevated() {
        return false;
    }

    !crate::CONFIG
        .lock()
        .map(|config| config.elevate_hand_offs)
        .unwrap_or(false)
}

/// Open `file` the way a hand-off of this app's always has — the Shell's `open` verb on that path or
/// address — with one difference: on an elevated run the desktop user performs it, so the program
/// that comes up is the desk's rather than an administrator's.
///
/// `monitor` is the display a preview was standing on, which is what a pin's own open asks for; the
/// Shell's own open takes no display hint, so a hand-off made as the desktop user lands wherever the
/// program decides. That is the one thing given up for the token, and it is the right way round: a
/// window on the wrong display is moved by hand, an administrator's browser is not.
pub(crate) fn open(file: &str, monitor: Option<isize>) -> bool {
    if through_the_shell() {
        return open_as_the_desktop_user(file);
    }

    let file = wide(file);
    let verb = wide("open");

    let mut info = SHELLEXECUTEINFOW {
        cbSize: std::mem::size_of::<SHELLEXECUTEINFOW>() as u32,
        fMask: SEE_MASK_ASYNCOK
            | match monitor {
                Some(_) => SEE_MASK_HMONITOR,
                None => 0,
            },
        lpVerb: PCWSTR(verb.as_ptr()),
        lpFile: PCWSTR(file.as_ptr()),
        nShow: SW_SHOWNORMAL.0,
        ..Default::default()
    };

    if let Some(monitor) = monitor {
        info.Anonymous.hMonitor = HANDLE(monitor as *mut core::ffi::c_void);
    }

    // Safety: the two strings are this function's own buffers living past the call, and the monitor
    // is a number read off a window rather than a pointer this side keeps.
    unsafe { ShellExecuteExW(&mut info) }.is_ok()
}

/// Start `program` with `arguments`, owned by the desktop user on an elevated run and by this
/// process otherwise — what `winget` and an update's installer are started with.
///
/// `arguments` is handed over as one command line, which is what both of those are: `winget` with
/// its own switches, and the installer with `/S /R`.
pub(crate) fn run(program: &str, arguments: &str) -> bool {
    if through_the_shell() {
        return as_the_desktop_user(program, &format!("\"{program}\" {arguments}"));
    }

    crate::app::engine_processes::hidden_command(program)
        .raw_arg(arguments)
        .spawn()
        .is_ok()
}

/// Start a copy of this app as the desktop user, which is what a tray replacing itself needs: the
/// copy that comes back is the unelevated one a launch by hand would start.
///
/// It is [`open`]'s sibling rather than its caller: the row that asks for it is turning an
/// administrator's token **off**, so the token cannot be the thing the setting decides — a restart
/// that came back elevated would leave the tray standing as the state the row has just stopped
/// asking for.
pub(crate) fn start_a_copy(exe: &str) -> bool {
    if as_the_desktop_user(exe, &format!("\"{exe}\"")) {
        return true;
    }

    // The token could not be had — a Shell this process may not open, or a machine whose Secondary
    // Logon service will not start. The Shell is asked to open the path instead: an Explorer begun
    // here is a second instance that forwards the path to the one already running, which is the
    // desk's own and starts the copy with its token.
    let explorer = explorer_path();
    crate::app::engine_processes::hidden_command(&explorer)
        .arg(exe)
        .spawn()
        .is_ok()
}

/// Open a path or an address with the desktop user's token: `explorer.exe` is handed it as its one
/// argument, and the Shell that performs it turns it into the program the user's own settings say.
///
/// **A path is not a program, so it is not started as one.** What is started is the Shell with the
/// path, which is the same call `Start → Run` makes; what it opens is the desktop user's program,
/// because both the Explorer already running and a second one begun here run with that token.
fn open_as_the_desktop_user(file: &str) -> bool {
    let explorer = explorer_path();

    if as_the_desktop_user(&explorer, &format!("\"{explorer}\" \"{file}\"")) {
        return true;
    }

    // The token could not be had: the Shell is begun with this process's own, and it is still the
    // running Explorer that performs the open, so the child is the desk's either way.
    crate::app::engine_processes::hidden_command(&explorer)
        .arg(file)
        .spawn()
        .is_ok()
}

/// Start `program` with `command_line` as the account at the desk, by duplicating the Shell's own
/// token and starting the process with it.
///
/// **The token is the whole of it.** `CreateProcessWithTokenW` needs `SeImpersonatePrivilege`, which
/// an elevated administrator's token carries and an ordinary run never asks for; a token that cannot
/// be duplicated, or a process the machine refuses to start with it, answers `false` so that the
/// caller can fall back rather than pretend.
fn as_the_desktop_user(program: &str, command_line: &str) -> bool {
    let Some(shell) = crate::app::elevation::shell_process_id() else {
        return false;
    };

    // Safety: every handle this call opens is closed here, the two strings live past the call, and
    // the structures are this function's own.
    unsafe {
        let Ok(process) = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, shell) else {
            return false;
        };

        let mut token = HANDLE::default();
        let opened = OpenProcessToken(process, TOKEN_DUPLICATE | TOKEN_QUERY, &mut token);
        // Safety: the handle `OpenProcess` made and nothing else holds.
        let _ = CloseHandle(process);

        if opened.is_err() {
            return false;
        }

        let mut primary = HANDLE::default();
        let duplicated = DuplicateTokenEx(
            token,
            TOKEN_ALL_ACCESS,
            None,
            SecurityImpersonation,
            TokenPrimary,
            &mut primary,
        );
        // Safety: the handle `OpenProcessToken` made and nothing else holds.
        let _ = CloseHandle(token);

        if duplicated.is_err() {
            return false;
        }

        let program = wide(program);
        let mut command_line = wide(command_line);
        let startup = STARTUPINFOW {
            cb: std::mem::size_of::<STARTUPINFOW>() as u32,
            ..Default::default()
        };
        let mut info = PROCESS_INFORMATION::default();

        let started = CreateProcessWithTokenW(
            primary,
            LOGON_WITH_PROFILE,
            PCWSTR(program.as_ptr()),
            PWSTR(command_line.as_mut_ptr()),
            PROCESS_CREATION_FLAGS(0),
            None,
            PCWSTR::null(),
            &startup,
            &mut info,
        );

        // Safety: the duplicated token this call made, which the process does not keep.
        let _ = CloseHandle(primary);

        if started.is_err() {
            return false;
        }

        // Safety: the two handles `CreateProcessWithTokenW` answered with, which nothing here waits
        // on: the copy is the desktop user's process, not a child whose exit this app has a use for.
        let _ = CloseHandle(info.hThread);
        let _ = CloseHandle(info.hProcess);

        true
    }
}

/// The Shell's own executable, which is the program a path is handed to when what is wanted is for
/// the Shell to open it.
fn explorer_path() -> String {
    std::env::var_os("SystemRoot")
        .map(std::path::PathBuf::from)
        .unwrap_or_else(|| std::path::PathBuf::from(r"C:\Windows"))
        .join("explorer.exe")
        .to_string_lossy()
        .into_owned()
}

/// A string as the platform's calls want it: UTF-16, terminated.
fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The one call that makes this module worth having: it is only asked for where a hand-off has
    /// something to de-elevate, and never by an ordinary run.
    #[test]
    fn an_ordinary_run_hands_nothing_to_the_shell() {
        // A run that is not elevated has no elevated token for a child to inherit, so the token it
        // hands over is the same one either way — and the road that keeps the monitor hint and the
        // job object is the one an ordinary run takes.
        if !crate::app::elevation::is_elevated() {
            assert!(
                !through_the_shell(),
                "an unelevated app never asks the Shell to start its children"
            );
        }
    }

    #[test]
    fn a_path_the_shell_is_handed_is_utf16_terminated() {
        // A folder with a space in it, spelled the way the Shell takes a path.
        let path = String::from("C:")
            + &char::from(92).to_string()
            + "a b"
            + &char::from(92).to_string()
            + "c.exe";
        let wide = wide(&path);

        assert_eq!(wide.last(), Some(&0));
        assert_eq!(String::from_utf16_lossy(&wide[..wide.len() - 1]), path);
    }

    /// The Shell is named by the machine rather than by a path this app decided, and the answer is
    /// the one file that exists on every Windows install.
    #[test]
    fn the_shell_is_the_one_the_machine_names() {
        let explorer = explorer_path();

        assert!(
            explorer.to_ascii_lowercase().ends_with("explorer.exe"),
            "the Shell is `explorer.exe`, not `{explorer}`"
        );
        assert!(
            std::path::Path::new(&explorer).is_file(),
            "the Shell named by the machine is a file: {explorer}"
        );
    }
}

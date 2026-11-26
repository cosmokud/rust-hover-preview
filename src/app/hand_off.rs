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
//! **The platform's own answer is to let the Shell start it.** `explorer.exe` runs with the ordinary
//! token of the account at the desktop, so a program it starts runs with that token — Windows'
//! documented behaviour, and the mechanism behind the well-known "run this unelevated" trick. This
//! module's whole job is that one call: `Shell.Application`'s `ShellExecute`, performed through the
//! Shell's automation object in the local server that is already running, which starts the target as
//! the desktop user rather than as the caller.
//!
//! **It is a setting, and it starts off.** `elevate_hand_offs` — the tray's
//! `System → Elevate Hand-Offs (Warning)` row — is off by default, which means every hand-off here is
//! started through the Shell's token; on, the app starts them itself, with the token it has. The
//! warning in that label is what the on state is: some program of a user's own really is one that
//! only works as an administrator, and there is no way for this app to know which.
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
use windows::core::{BSTR, PCWSTR, VARIANT};
use windows::Win32::Foundation::HANDLE;
use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_LOCAL_SERVER};
use windows::Win32::UI::Shell::{
    IShellDispatch2, Shell, ShellExecuteExW, SEE_MASK_ASYNCOK, SEE_MASK_HMONITOR, SHELLEXECUTEINFOW,
};
use windows::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL;

/// Whether a hand-off should be started with the Shell's own (unelevated) token rather than this
/// process's.
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

/// Ask the Shell to open `file` with `verb` — `open`, `runas`, `print`, or one of the file's own —
/// and answer whether the Shell took the request.
///
/// `arguments` is the command line for a program named as the file, which is what a hand-off with
/// parameters is: `winget install …` is `winget.exe` with its own switches. `None` is the ordinary
/// case of a file the Shell opens by its association.
pub(crate) fn shell_execute(file: &str, arguments: Option<&str>, verb: &str) -> bool {
    // Safety: the Shell's automation object is a local server of the platform's own, and every
    // interface this reaches is asked for and answered rather than assumed.
    unsafe {
        // `Shell.Application` is the platform's own automation object, and `IShellDispatch2` is the
        // interface whose `ShellExecute` takes a verb and a command line — the one the documented
        // "start this unelevated" road is walked with.
        let Ok(shell): windows::core::Result<IShellDispatch2> =
            CoCreateInstance(&Shell, None, CLSCTX_LOCAL_SERVER)
        else {
            return false;
        };

        let arguments = arguments.map_or_else(VARIANT::default, VARIANT::from);
        let directory = VARIANT::default();
        let verb = VARIANT::from(verb);
        let show = VARIANT::from(1i32);

        shell
            .ShellExecute(&BSTR::from(file), &arguments, &directory, &verb, &show)
            .is_ok()
    }
}

/// Open `file` the way a hand-off of this app's always has — the Shell's `open` verb on that path or
/// address — with one difference: on an elevated run the Shell performs it, so the program that
/// comes up is the desktop user's rather than an administrator's.
///
/// `monitor` is the display a preview was standing on, which is what a pin's own open asks for; the
/// Shell's automation object takes no display hint, so a de-elevated hand-off lands wherever the
/// program decides. That is the one thing given up for the token, and it is the right way round: a
/// window on the wrong display is moved by hand, an administrator's browser is not.
pub(crate) fn open(file: &str, monitor: Option<isize>) -> bool {
    if through_the_shell() {
        return shell_execute(file, None, "open");
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
        return shell_execute(program, Some(arguments), "open");
    }

    crate::app::engine_processes::hidden_command(program)
        .raw_arg(arguments)
        .spawn()
        .is_ok()
}

/// A string as the platform's calls want it: UTF-16, terminated.
fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The one call that makes this module worth having: it is the Shell's own object that performs
    /// a hand-off, and it is only asked for where a hand-off has something to de-elevate.
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
}

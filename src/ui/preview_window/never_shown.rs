//! Beginning a viewer whose window is never on a screen: `CreateProcessW` with a `STARTUPINFO`
//! that says so, and the pipe its own loading report arrives on.
//!
//! **Why the launch is this module's and not `Command`'s.** f3d's window is born visible —
//! measured on 4.0.1: at 180 ms it stands at 0,0, 316x339, decorated and black, and about 30 ms
//! later it is moved to `--position`/`--resolution` — so `--position` cannot keep it off the screen,
//! the position being applied after the window is created and shown. What does, measured, is
//! `SW_HIDE` in the `STARTUPINFO` the process is begun with: no window of the process is visible at
//! birth, after the move to the parking origin, or for the whole run, and once this app shows the
//! window itself the client has real content (sampled `0x333333`, `0x323232`, `0x656566`, where a
//! never-shown window's DC reads `0xFFFFFF`). `Command` cannot carry one — `CommandExt::show_window`
//! is unstable in this toolchain (`windows_process_extensions_show_window`, issue #127544) — so the
//! first `ShowWindow` in that window's life is `f3d_player::show_the_window`'s.
//!
//! **And the argv is still the launch's**: `f3d_player::launch` builds its `std::process::Command`
//! exactly as every other launch does, and this reads the program, the arguments and the creation
//! flags off it — so the command line the tests read back is the command line begun.
//!
//! **The report is the one handle of this app's the child is given**, besides the `NUL` device on
//! stdin and stderr: `STARTF_USESTDHANDLES` means every handle the child is to see is named here,
//! and a pipe is the only way the loading marker reaches a reader that is not the console (see
//! `f3d_load`).

use super::*;
use std::ffi::OsStr;
use std::os::windows::io::{AsRawHandle, FromRawHandle, OwnedHandle};
use windows::Win32::Foundation::{SetHandleInformation, BOOL, HANDLE_FLAGS, HANDLE_FLAG_INHERIT};
use windows::Win32::Security::SECURITY_ATTRIBUTES;
use windows::Win32::System::Pipes::CreatePipe;
use windows::Win32::System::Threading::{
    CreateProcessW, GetExitCodeProcess, PROCESS_CREATION_FLAGS, PROCESS_INFORMATION,
    STARTF_USESHOWWINDOW, STARTF_USESTDHANDLES, STARTUPINFOW,
};

/// The quote a command line's argument is wrapped in, and the escape the CRT reads back.
const QUOTE: u16 = b'"' as u16;
/// What an odd count of is a quote being escaped, and an even one a literal backslash.
const BACKSLASH: u16 = b'\\' as u16;
/// The two characters that make an argument one that has to be quoted at all.
const SPACE: u16 = b' ' as u16;
const TAB: u16 = b'\t' as u16;

/// A viewer begun by hand: the process this app holds, the id it is booked in under, and the read
/// end of the pipe its own report arrives on.
///
/// The process is an owned handle rather than a `Child` — there is no `Child` to be had from a
/// `CreateProcessW` — which is what says the handle is closed once, here, on every road that ends
/// with this struct dropped (see `video_playback::PlayerProcess`).
pub(super) struct NeverShown {
    pub(super) process: OwnedHandle,
    pub(super) pid: u32,
    /// Handed to the report's own reader once, and read by it for as long as the viewer writes
    /// (see `f3d_load::drain_it`).
    pub(super) report: Option<std::fs::File>,
}

impl NeverShown {
    /// Whether the process has ended, asked with a wait of nothing and answered by its exit code.
    ///
    /// `ExitStatus::from_raw` is what makes this the same answer `Child::try_wait` gives, and it is
    /// the stable half of a `Child` this road has none of.
    pub(super) fn try_wait(&self) -> std::io::Result<Option<std::process::ExitStatus>> {
        use std::os::windows::process::ExitStatusExt;

        let process = self.handle();
        if unsafe { WaitForSingleObject(process, 0) } != WAIT_OBJECT_0 {
            return Ok(None);
        }
        let mut code = 0u32;
        unsafe { GetExitCodeProcess(process, &mut code) }.map_err(a_win32_error)?;
        Ok(Some(std::process::ExitStatus::from_raw(code)))
    }

    /// End the process, on the terms `Child::kill` does: a terminate request, and nothing waited on.
    pub(super) fn kill(&self) -> std::io::Result<()> {
        unsafe { TerminateProcess(self.handle(), 1) }.map_err(a_win32_error)
    }

    fn handle(&self) -> HANDLE {
        HANDLE(self.process.as_raw_handle())
    }
}

/// A Win32 failure as the `io::Error` the `Child` road would have answered with.
fn a_win32_error(error: windows::core::Error) -> std::io::Error {
    std::io::Error::from_raw_os_error(error.code().0)
}

/// Begin `command` with a `STARTUPINFO` that says its window is never to be shown, and answer the
/// process, its id, and the read end of the pipe it reports on (see the module note).
///
/// The program is given as `lpApplicationName` rather than only as the command line's first token,
/// so a path with spaces in it is a path and not two arguments; the command line carries it anyway,
/// because that is the `argv[0]` the viewer reads back. **It is a path and not a name**: a name is
/// searched for on `PATH` by `Command` and never by `lpApplicationName`, and what is handed in here
/// is what `f3d_render::program` resolved — one of three paths.
///
/// Safety: nothing here is unsound, but the calls are Win32's: the handles this creates are the
/// only ones the child inherits, and every one of them is closed on every path out of here.
pub(super) fn spawn(command: &mut std::process::Command) -> Option<NeverShown> {
    let program = command.get_program();
    let application = wide(program);
    let mut line = a_command_line(program, command.get_args());

    // **`NUL`, made inheritable by hand**: a handle this app opens is opened without
    // `HANDLE_FLAG_INHERIT`, and a child named a handle it cannot inherit reads and writes nowhere.
    let nul = std::fs::OpenOptions::new()
        .read(true)
        .write(true)
        .open("NUL")
        .ok()?;
    let nul_handle = HANDLE(nul.as_raw_handle());
    unsafe { SetHandleInformation(nul_handle, HANDLE_FLAG_INHERIT.0, HANDLE_FLAG_INHERIT) }.ok()?;

    let inheritable = SECURITY_ATTRIBUTES {
        nLength: std::mem::size_of::<SECURITY_ATTRIBUTES>() as u32,
        lpSecurityDescriptor: std::ptr::null_mut(),
        bInheritHandle: BOOL(1),
    };
    let mut read = HANDLE::default();
    let mut write = HANDLE::default();
    unsafe { CreatePipe(&mut read, &mut write, Some(&inheritable), 0) }.ok()?;
    // **The read end is this app's alone**: a child holding it is a pipe a reader of the child's
    // own could never see close — and a pipe that cannot be said to be this app's is no launch.
    if unsafe { SetHandleInformation(read, HANDLE_FLAG_INHERIT.0, HANDLE_FLAGS(0)) }.is_err() {
        let _ = unsafe { CloseHandle(read) };
        let _ = unsafe { CloseHandle(write) };
        return None;
    }

    let startup = a_startup_info(nul_handle, write);
    let mut begun = PROCESS_INFORMATION::default();
    let started = unsafe {
        CreateProcessW(
            PCWSTR(application.as_ptr()),
            PWSTR(line.as_mut_ptr()),
            None,
            None,
            BOOL(1),
            // The creation flags the launch asked for — `CREATE_NO_WINDOW`, the one every engine of
            // this app's is begun with (see `engine_processes::hidden_command`).
            PROCESS_CREATION_FLAGS(engine_processes::CREATE_NO_WINDOW),
            None,
            PCWSTR::null(),
            &startup,
            &mut begun,
        )
    };
    // **The write end is the child's from here, on every path**: a write end this app still held
    // would keep the pipe open past the viewer's end, and its reader would wait for a close that
    // cannot come. The `NUL` handle goes with this function's scope, which is the same every path.
    let _ = unsafe { CloseHandle(write) };
    if started.is_err() {
        let _ = unsafe { CloseHandle(read) };
        return None;
    }
    // Nothing of this app's waits on the thread, and the handle would outlive the one use it has.
    let _ = unsafe { CloseHandle(begun.hThread) };

    Some(NeverShown {
        // SAFETY: both handles are this call's, unclosed until now, and neither is closed again.
        process: unsafe { OwnedHandle::from_raw_handle(begun.hProcess.0) },
        pid: begun.dwProcessId,
        report: Some(unsafe { std::fs::File::from_raw_handle(read.0) }),
    })
}

/// The `STARTUPINFO` a viewer is begun with: **the one pair of flags that keeps its window off every
/// screen** (see the module note), the show state it names, and the three handles it hands over.
fn a_startup_info(nul: HANDLE, report: HANDLE) -> STARTUPINFOW {
    STARTUPINFOW {
        cb: std::mem::size_of::<STARTUPINFOW>() as u32,
        dwFlags: STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES,
        wShowWindow: SW_HIDE.0 as u16,
        hStdInput: nul,
        hStdOutput: report,
        hStdError: nul,
        ..Default::default()
    }
}

/// The command line `CreateProcessW` is handed: the program and then every argument, each quoted as
/// the CRT reads it back (see `quoted`).
fn a_command_line<'a>(program: &'a OsStr, args: impl Iterator<Item = &'a OsStr>) -> Vec<u16> {
    let mut line: Vec<u16> = Vec::new();
    for argument in std::iter::once(program).chain(args) {
        if !line.is_empty() {
            line.push(SPACE);
        }
        line.extend(quoted(argument));
    }
    line.push(0);
    line
}

/// One argument as the CRT's own parser reads it back, quoted the way it expects.
///
/// The rules are the ones every Windows program's `main` is given its arguments by: an argument
/// that is empty, or holds a space, a tab or a quote, is wrapped in quotes; a run of backslashes
/// before a quote is doubled and one more added, so the quote is the escaped one; and a run at the
/// end of an argument is doubled, so the closing quote is not escaped. Anything else is itself.
///
/// What this is for is the one thing a model's launch cannot lose: **a model's path with a space in
/// it is one argument**, and a path that reached f3d as two is a file it cannot open.
fn quoted(argument: &OsStr) -> Vec<u16> {
    let units: Vec<u16> = argument.encode_wide().collect();
    let plain = !units.is_empty()
        && !units
            .iter()
            .any(|unit| matches!(*unit, SPACE | TAB | QUOTE));
    if plain {
        return units;
    }

    let mut quoted: Vec<u16> = Vec::with_capacity(units.len() + 2);
    quoted.push(QUOTE);
    let mut backslashes = 0usize;
    for unit in units {
        match unit {
            QUOTE => {
                quoted.extend(std::iter::repeat_n(BACKSLASH, backslashes * 2 + 1));
                quoted.push(QUOTE);
                backslashes = 0;
            }
            BACKSLASH => backslashes += 1,
            unit => {
                quoted.extend(std::iter::repeat_n(BACKSLASH, backslashes));
                quoted.push(unit);
                backslashes = 0;
            }
        }
    }
    quoted.extend(std::iter::repeat_n(BACKSLASH, backslashes * 2));
    quoted.push(QUOTE);
    quoted
}

/// One string as the UTF-16 a Win32 call takes, terminated.
fn wide(value: &OsStr) -> Vec<u16> {
    value.encode_wide().chain(Some(0)).collect()
}

/// A process this app can begin in a test and be sure ends: the machine's own command processor,
/// told to leave with a code a test can read back.
#[cfg(test)]
pub(super) fn a_stand_in_process() -> std::process::Command {
    let shell = std::env::var_os("COMSPEC").expect("a machine has a command processor");
    let mut command = engine_processes::hidden_command(shell);
    command.args(["/c", "exit 3"]);
    command
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A command line as the string the desktop would see it as.
    fn the_line(line: &[u16]) -> String {
        assert_eq!(line.last(), Some(&0), "a command line is a terminated one");
        String::from_utf16(&line[..line.len() - 1]).expect("the line is what was put in it")
    }

    fn the_line_of(program: &str, args: &[&str]) -> String {
        let args: Vec<&OsStr> = args.iter().map(|arg| OsStr::new(*arg)).collect();
        the_line(&a_command_line(OsStr::new(program), args.into_iter()))
    }

    /// **A model's path with a space in it is one argument**, and so is a path with a quote, a run
    /// of backslashes before one, and a trailing one: every rule the CRT's own parser reads a
    /// command line back by, on the one line a model's launch hands it (see `quoted`).
    #[test]
    fn a_path_with_a_space_survives_the_command_line() {
        assert_eq!(
            the_line_of("f3d.exe", &["C:/models/a-model.stl"]),
            "f3d.exe C:/models/a-model.stl",
            "an argument that needs no quoting is itself"
        );
        assert_eq!(
            the_line_of("f3d.exe", &["C:/my models/a-model.stl"]),
            "f3d.exe \"C:/my models/a-model.stl\"",
            "a space is what quotes an argument"
        );
        assert_eq!(
            the_line_of("C:/Program Files/f3d/f3d.exe", &[]),
            "\"C:/Program Files/f3d/f3d.exe\"",
            "and the program is an argument like any other"
        );
        assert_eq!(
            the_line_of("f3d.exe", &[""]),
            "f3d.exe \"\"",
            "an empty argument is a pair of quotes, not nothing"
        );
        assert_eq!(
            the_line_of("f3d.exe", &["a\tb"]),
            "f3d.exe \"a\tb\"",
            "a tab is a space to the parser"
        );
        assert_eq!(
            the_line_of("f3d.exe", &["a\"b"]),
            "f3d.exe \"a\\\"b\"",
            "a quote is escaped inside the quotes it forced"
        );
        assert_eq!(
            the_line_of("f3d.exe", &["C:/my models/"]),
            "f3d.exe \"C:/my models/\"",
            "a trailing backslash inside quotes is the end of the path, not an escape"
        );
        assert_eq!(
            the_line_of("f3d.exe", &["C:/my models/sub\\"]),
            "f3d.exe \"C:/my models/sub\\\\\"",
            "so a run of backslashes before the closing quote is doubled"
        );
        assert_eq!(
            the_line_of("f3d.exe", &["a\\\"b"]),
            "f3d.exe \"a\\\\\\\"b\"",
            "and a run before a quote is doubled and one more, which is the quote escaped"
        );
        assert_eq!(
            the_line_of("f3d.exe", &["one", "two three", "four"]),
            "f3d.exe one \"two three\" four",
            "every argument is its own, in the order the launch asked for them"
        );
    }

    /// **The viewer is begun with a show state and not a placement**: the window is hidden from the
    /// first instruction it runs, and the three handles it is given are the `NUL` device on both
    /// ends nothing reads and the pipe its report arrives on (see `a_startup_info`).
    #[test]
    fn a_viewer_is_begun_hidden_and_told_which_handles_it_has() {
        let nul = HANDLE(0x1000 as *mut _);
        let report = HANDLE(0x2000 as *mut _);
        let startup = a_startup_info(nul, report);

        assert_eq!(
            startup.cb as usize,
            std::mem::size_of::<STARTUPINFOW>(),
            "the struct says its own size, which is how Windows reads it"
        );
        assert!(
            startup.dwFlags.contains(STARTF_USESHOWWINDOW),
            "the show state below is one the viewer is given"
        );
        assert!(
            startup.dwFlags.contains(STARTF_USESTDHANDLES),
            "and the handles below are the ones it is given"
        );
        assert_eq!(
            startup.wShowWindow, SW_HIDE.0 as u16,
            "hidden: the first time this window is on a screen is this app's own show"
        );
        assert_eq!(
            startup.hStdOutput, report,
            "the report is the pipe's write end"
        );
        assert_eq!(startup.hStdInput, nul, "stdin is the device nothing reads");
        assert_eq!(
            startup.hStdError, nul,
            "and stderr the same device, so a viewer's complaint is not a window"
        );
    }

    /// **A viewer begun by hand is begun for real**: the pipe it is given is one this app holds the
    /// read end of, and it closes when the viewer does — a write end this app kept would keep it open
    /// past the end of the process, and the report's reader would wait on a close that cannot come
    /// (see `spawn`). The rest of what the launch is asked is `video_playback`'s own test, which is
    /// where the `PlayerProcess` it becomes is asked whether it has ended.
    #[test]
    fn a_hand_begun_viewer_is_given_a_pipe_that_closes_with_it() {
        let mut command = a_stand_in_process();
        let mut begun = spawn(&mut command).expect("a stand-in process this app holds");

        let mut report = begun.report.take().expect("the report it was given");
        let mut bytes = Vec::new();
        assert_eq!(
            std::io::Read::read_to_end(&mut report, &mut bytes).expect("a pipe reads"),
            0,
            "the pipe closes with the viewer rather than being held open by this app"
        );
    }
}

use std::env;
use std::fs;
use std::os::windows::ffi::OsStrExt;
use std::path::Path;
use windows::core::{w, PCWSTR};
use windows::Win32::System::Registry::{
    RegCloseKey, RegDeleteValueW, RegOpenKeyExW, RegQueryValueExW, RegSetValueExW, HKEY,
    HKEY_CURRENT_USER, KEY_READ, KEY_SET_VALUE, REG_SZ,
};

const STARTUP_KEY: PCWSTR = w!(r"Software\Microsoft\Windows\CurrentVersion\Run");
const APP_NAME: PCWSTR = w!("RustHoverPreview");

/// Write the path of the copy that is running into the entry, and say whether it took.
///
/// One write, and the path is this process's own rather than one remembered from whenever
/// the entry was last written: an app moved, or installed again over itself, is a path that
/// no longer exists, and an entry naming it starts nothing at the next logon.
fn write_entry(path: &Path) -> bool {
    // The value is a counted string, and its terminator is part of what is written: a path
    // whose bytes stop one short of it is read back carrying whatever was in the registry
    // next to it.
    let wide: Vec<u16> = path
        .as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect();

    unsafe {
        let mut hkey: HKEY = HKEY::default();
        if RegOpenKeyExW(HKEY_CURRENT_USER, STARTUP_KEY, 0, KEY_SET_VALUE, &mut hkey).is_err() {
            return false;
        }

        let written =
            RegSetValueExW(hkey, APP_NAME, 0, REG_SZ, Some(wide.align_to::<u8>().1)).is_ok();
        let _ = RegCloseKey(hkey);
        written
    }
}

pub fn enable_startup() {
    if let Ok(exe_path) = env::current_exe() {
        write_entry(&exe_path);
    }
}

pub fn disable_startup() {
    unsafe {
        let mut hkey: HKEY = HKEY::default();
        if RegOpenKeyExW(HKEY_CURRENT_USER, STARTUP_KEY, 0, KEY_SET_VALUE, &mut hkey).is_ok() {
            let _ = RegDeleteValueW(hkey, APP_NAME);
            let _ = RegCloseKey(hkey);
        }
    }
}

/// Whether this app is registered to start when the user logs in.
///
/// Asked of the registry rather than of the configuration, because the registry is what
/// actually starts it: the value is the entry's presence, so an entry the user turned off in
/// Task Manager, or one another tool removed, is answered as the app that will not start
/// rather than as the value `config.ini` remembers choosing.
pub fn is_startup_enabled() -> bool {
    unsafe {
        let mut hkey: HKEY = HKEY::default();
        if RegOpenKeyExW(HKEY_CURRENT_USER, STARTUP_KEY, 0, KEY_READ, &mut hkey).is_ok() {
            let result = RegQueryValueExW(hkey, APP_NAME, None, None, None, None).is_ok();
            let _ = RegCloseKey(hkey);
            return result;
        }
    }
    false
}

/// What the entry holds, as the text it is read back as.
///
/// An entry this app writes is a path, and the shape of it can be told from the type alone —
/// but the entry is the registry's, and a value is whatever is in it, so what is read is
/// the bytes rather than the type's promise of them: a `REG_SZ` carrying a number is still
/// text to be judged, not a length to be trusted.
fn registered_command() -> Option<String> {
    unsafe {
        let mut hkey: HKEY = HKEY::default();
        if RegOpenKeyExW(HKEY_CURRENT_USER, STARTUP_KEY, 0, KEY_READ, &mut hkey).is_err() {
            return None;
        }

        // The size is asked for before the bytes are, the way the API counts a string: one
        // call for how many bytes there are, and one for the bytes themselves.
        let mut size: u32 = 0;
        let sized = RegQueryValueExW(hkey, APP_NAME, None, None, None, Some(&mut size)).is_ok()
            // A string is at least its own terminator; anything shorter is a value that
            // names no executable rather than one that names it emptily.
            && size >= 2;

        let command = if sized {
            let mut buffer = vec![0u8; size as usize];
            let read = RegQueryValueExW(
                hkey,
                APP_NAME,
                None,
                None,
                Some(buffer.as_mut_ptr()),
                Some(&mut size),
            )
            .is_ok();

            read.then(|| decode_string(&buffer[..size.min(buffer.len() as u32) as usize]))
        } else {
            None
        };

        let _ = RegCloseKey(hkey);
        command
    }
}

/// A counted registry string as text: UTF-16, ending at its terminator.
///
/// The terminator is where the string ends whatever the buffer behind it holds, and an odd
/// trailing byte is a byte this reads nothing out of rather than a panic.
fn decode_string(bytes: &[u8]) -> String {
    let units: Vec<u16> = bytes
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .take_while(|unit| *unit != 0)
        .collect();

    String::from_utf16_lossy(&units)
}

/// The executable a startup entry names, out of the command line the value holds.
///
/// Windows runs the value as a command line, so what is in the registry need not be a bare
/// path: a path holding a space is quoted, or half of it would run, and arguments can
/// follow the executable. Only the executable is the app — an argument is not a path to
/// another copy of it — so the name is taken from in front of them.
fn executable_of(command: &str) -> &str {
    let trimmed = command.trim();

    match trimmed.strip_prefix('"') {
        Some(quoted) => quoted.split('"').next().unwrap_or(quoted),
        None => trimmed.split_whitespace().next().unwrap_or(trimmed),
    }
}

/// Whether two spellings are the path of one file, as the machine reads a path.
///
/// The spelling itself settles it where it can: a Windows path carries no case, and the
/// verbatim prefix names the same file as without it. Where the two read differently the
/// files themselves are asked, which is what settles a short name for a folder or a
/// junction — a spelling this app never writes, and would otherwise replace over one file
/// with itself. A path that does not resolve is asked nothing: an entry naming a copy that
/// has been deleted is not this one, and is the case this exists for.
fn same_path(registered: &str, running: &Path) -> bool {
    let left = crate::paths::plain_path(Path::new(registered.trim()));
    let right = crate::paths::plain_path(running);

    if left.eq_ignore_ascii_case(&right) {
        return true;
    }

    match (fs::canonicalize(&left), fs::canonicalize(&right)) {
        (Ok(left), Ok(right)) => {
            crate::paths::plain_path(&left).eq_ignore_ascii_case(&crate::paths::plain_path(&right))
        }
        _ => false,
    }
}

/// Whether the entry, as it stands, starts the copy of the app that is running.
fn entry_names_this_copy() -> bool {
    let (Some(command), Ok(exe_path)) = (registered_command(), env::current_exe()) else {
        return false;
    };

    same_path(executable_of(&command), &exe_path)
}

/// Put the entry back on the copy of the app that is running, and say whether it was
/// rewritten.
///
/// What the entry says and what this process is are two things that can come apart: a
/// portable copy run once from a folder of its own, an older version still installed where
/// it put itself, an app moved since — and each writes its own path in, which is what the
/// next logon starts. The entry then names a version of this app that the tray the user is
/// looking at does not belong to, or nothing at all.
///
/// Only a path is touched, and only where the entry is already there: whether this app
/// starts at all is the user's choice, and a choice to have it off is not one this
/// overrides. Where the entry names this copy it is left exactly as it is, an argument of
/// the user's own included, and where it names another it is written in this app's own
/// spelling, since that is what starts it.
pub fn repair_startup_entry() -> bool {
    if !is_startup_enabled() || entry_names_this_copy() {
        return false;
    }

    match env::current_exe() {
        Ok(exe_path) => write_entry(&exe_path),
        Err(_) => false,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn an_entry_is_read_as_the_executable_it_names() {
        // The path this app writes.
        assert_eq!(
            executable_of(r"C:\apps\rust-hover-preview\rust-hover-preview.exe"),
            r"C:\apps\rust-hover-preview\rust-hover-preview.exe"
        );
        // A path holding a space is quoted, and the quotes are not part of the name.
        assert_eq!(
            executable_of(r#""C:\Program Files\Rust Hover Preview\app.exe""#),
            r"C:\Program Files\Rust Hover Preview\app.exe"
        );
        // Arguments follow the executable and are not another copy of the app.
        assert_eq!(
            executable_of(r#""C:\apps\app.exe" --silent"#),
            r"C:\apps\app.exe"
        );
        assert_eq!(
            executable_of(r"C:\apps\app.exe --silent"),
            r"C:\apps\app.exe"
        );
        // Surrounding whitespace is not part of the name either.
        assert_eq!(executable_of("  C:\\apps\\app.exe  "), r"C:\apps\app.exe");
        // A value naming nothing at all is read as nothing rather than as a path.
        assert_eq!(executable_of(""), "");
        assert_eq!(executable_of(r#""""#), "");
    }

    #[test]
    fn one_path_in_another_spelling_is_this_path() {
        let running = Path::new(r"C:\apps\rust-hover-preview\rust-hover-preview.exe");

        // The same path, written as this app writes it.
        assert!(same_path(
            r"C:\apps\rust-hover-preview\rust-hover-preview.exe",
            running
        ));
        // A Windows path carries no case.
        assert!(same_path(
            r"C:\Apps\Rust-Hover-Preview\Rust-Hover-Preview.EXE",
            running
        ));
        // The verbatim prefix names the same file as without it, and is what the Shell
        // canonicalizes to, so an entry holding one is not a different copy.
        assert!(same_path(
            r"\\?\C:\apps\rust-hover-preview\rust-hover-preview.exe",
            running
        ));
        // Whitespace around the value is not part of the path.
        assert!(same_path(
            "  C:\\apps\\rust-hover-preview\\rust-hover-preview.exe  ",
            running
        ));
    }

    #[test]
    fn another_copy_of_the_app_is_not_this_copy() {
        let running = Path::new(r"C:\apps\rust-hover-preview\rust-hover-preview.exe");

        // The stray portable copy: another folder, another version, the thing the check
        // exists for.
        assert!(!same_path(
            r"C:\Users\test\Downloads\rust-hover-preview.exe",
            running
        ));
        // A path that no longer resolves at all is not this one either, and is not asked
        // about: it is exactly what has to be replaced.
        assert!(!same_path(
            r"C:\apps\rust-hover-preview-0.3.6\app.exe",
            running
        ));
        // An empty value names nothing.
        assert!(!same_path("", running));
    }

    #[test]
    fn a_counted_string_is_read_up_to_its_terminator() {
        let counted = |text: &str| -> Vec<u8> {
            text.encode_utf16()
                .chain(std::iter::once(0))
                .flat_map(u16::to_le_bytes)
                .collect()
        };

        assert_eq!(
            decode_string(&counted(r"C:\apps\app.exe")),
            r"C:\apps\app.exe"
        );
        // Whatever is in the buffer behind the terminator is not part of the string.
        assert_eq!(decode_string(&counted("app.exe")), "app.exe");
        // A string that is nothing but its terminator names nothing.
        assert_eq!(decode_string(&counted("")), "");
        // An odd trailing byte is a byte this reads nothing out of.
        assert_eq!(decode_string(&[0x41, 0x00, 0x42]), "A");
    }
}

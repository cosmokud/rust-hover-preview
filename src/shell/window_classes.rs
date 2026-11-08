//! Which windows are Explorer's, and which of them is in front: the two answers the preview loop
//! and the Explorer hook both have to give the same way.
//!
//! Both sides ask these, and they are asked here rather than in the hook for that reason: a pin key
//! or a bubble restored only where Explorer is in front, and a pin taken as the window under the
//! hand only where the window beside it is an Explorer frame, are decisions the preview thread makes
//! on its own, out of facts about the desktop it reads for itself. Nothing of the hook's state is
//! here — no cache of views, no item under the pointer — so this is a module both sides own, and it
//! kept the hook's name only because the classes it names are the shell's.
//!
//! **Only the class**, and never the process behind a window: explorer.exe also owns the desktop, the
//! taskbar, the Start menu and the search box, and reading one of those as an Explorer window is what
//! made the foreground test answer yes with the desktop in front (see `is_explorer_window`, which is
//! the cached form of the same question).

use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{GetClassNameW, GetForegroundWindow};

/// Whether the window in front is Explorer's, asked of the foreground window and one class lookup.
///
/// This is the form the low-level hooks use, because a `WH_KEYBOARD_LL` callback may not take a lock
/// and the hook's own cached form of the question does. It is also the form the preview thread asks:
/// the class of a window is not something that changes under the answer, so the cache the hook keeps
/// is a memo of the lookup rather than a reading of the desktop.
pub(crate) fn the_foreground_is_explorer() -> bool {
    unsafe {
        let foreground = GetForegroundWindow();
        if foreground.is_invalid() {
            return false;
        }
        explorer_class_matches(foreground)
    }
}

/// Whether a window is one of Explorer's folder windows, by its class name.
///
/// The class is the whole of the test, and both of Explorer's are matched: the browser frame a folder
/// is drawn in (`CabinetWClass`), and the older frame (`ExploreWClass`) a shell that still has one
/// puts up. The rest of the shell is explorer.exe as well — `Progman` and `WorkerW` are the desktop,
/// the taskbar and the Start menu are windows of its own — which is why the process is not asked
/// instead: it answers yes for every one of them.
pub(crate) fn explorer_class_matches(hwnd: HWND) -> bool {
    unsafe {
        let mut class_name = [0u16; 256];
        let len = GetClassNameW(hwnd, &mut class_name);
        if len <= 0 {
            return false;
        }

        let class_name = &class_name[..len as usize];
        utf16_contains_ascii_ignore_case(class_name, b"cabinetwclass")
            || utf16_contains_ascii_ignore_case(class_name, b"explorerwclass")
    }
}

/// Allocation-free equivalent of lowercasing the class name and searching for an ASCII needle. A
/// `u16` outside ASCII becomes U+FFFD under `to_string_lossy` and can never match, so it is skipped.
pub(crate) fn utf16_contains_ascii_ignore_case(haystack: &[u16], needle: &[u8]) -> bool {
    if needle.is_empty() || haystack.len() < needle.len() {
        return false;
    }

    haystack.windows(needle.len()).any(|window| {
        window
            .iter()
            .zip(needle)
            .all(|(ch, byte)| *ch < 0x80 && (*ch as u8).eq_ignore_ascii_case(byte))
    })
}

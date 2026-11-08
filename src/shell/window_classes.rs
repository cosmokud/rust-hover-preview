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
use windows::Win32::UI::WindowsAndMessaging::{
    GetAncestor, GetClassNameW, GetForegroundWindow, GA_ROOT,
};

/// The frame a window stands in: the top-level window whose band every part of it is drawn in.
///
/// **A press inside a listing is a press on one of the frame's children.** The window the desktop
/// names for a press made in a folder view — the window under the point, the one a placement is
/// asked to put another window behind — is the tree, the view's own `DirectUIHWND`, the list's rows,
/// a toolbar or the status bar, as often as it is the `CabinetWClass` frame. Every question this app
/// asks of such a window (is this Explorer, is this the listing a pin goes behind) is a question
/// about the frame, so the frame is what they are asked about (see `explorer_class_matches`).
///
/// One `GetAncestor`, no lock and no allocation, so a low-level hook may ask it. A window that is
/// already top-level, or one the desktop no longer keeps, answers itself.
pub(crate) fn the_frame_of(hwnd: HWND) -> HWND {
    if hwnd.is_invalid() {
        return hwnd;
    }

    // Safety: asks the desktop for the ancestor of a handle it gave out; dereferences nothing.
    let frame = unsafe { GetAncestor(hwnd, GA_ROOT) };
    if frame.is_invalid() {
        hwnd
    } else {
        frame
    }
}

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

/// Whether a window is part of one of Explorer's folder windows, by the class of the frame it stands
/// in.
///
/// The class is the whole of the test, and both of Explorer's are matched: the browser frame a folder
/// is drawn in (`CabinetWClass`), and the older frame (`ExploreWClass`) a shell that still has one
/// puts up. The rest of the shell is explorer.exe as well — `Progman` and `WorkerW` are the desktop,
/// the taskbar and the Start menu are windows of its own — which is why the process is not asked
/// instead: it answers yes for every one of them.
///
/// **Asked of the frame rather than of the window handed in, because a click inside a listing lands
/// on a child of it** (see `the_frame_of`). Asked of the child alone this answered no for the whole
/// of a window the user was working in: a press in the file view was never read as a press in a
/// listing, so the pin raised over that listing was never put back behind it, and one answered as a
/// press on the child left a placement asked to put a window behind a child — a call the desktop
/// does nothing with.
pub(crate) fn explorer_class_matches(hwnd: HWND) -> bool {
    let frame = the_frame_of(hwnd);

    unsafe {
        let mut class_name = [0u16; 256];
        let len = GetClassNameW(frame, &mut class_name);
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

#[cfg(test)]
mod tests {
    use super::*;
    use windows::core::{w, PCWSTR};
    use windows::Win32::Foundation::{LPARAM, LRESULT, WPARAM};
    use windows::Win32::System::LibraryLoader::GetModuleHandleW;
    use windows::Win32::UI::WindowsAndMessaging::{
        CreateWindowExW, DefWindowProcW, DestroyWindow, RegisterClassW, WINDOW_EX_STYLE, WNDCLASSW,
        WS_CHILD, WS_POPUP,
    };

    unsafe extern "system" fn an_empty_window(
        hwnd: HWND,
        msg: u32,
        wparam: WPARAM,
        lparam: LPARAM,
    ) -> LRESULT {
        // Safety: the default procedure answers a message this window's own class never handles.
        unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
    }

    /// The windows one of these tests makes, put away when the test ends: a panic in an assertion
    /// leaves none of them on the desktop for the suite running beside it to count as an Explorer
    /// window.
    #[derive(Default)]
    struct WindowsOfTheTest {
        made: Vec<HWND>,
    }

    impl WindowsOfTheTest {
        /// A window of this process's own named as the Shell names one of its own: class names are
        /// per process, so a name Explorer holds is free here.
        ///
        /// Safety: the window is made where no display is and is never shown, and the test puts it
        /// away with the rest of the set.
        unsafe fn a_window_named(&mut self, class: PCWSTR, parent: Option<HWND>) -> HWND {
            let instance = GetModuleHandleW(None).expect("this module");
            let wc = WNDCLASSW {
                lpfnWndProc: Some(an_empty_window),
                hInstance: instance.into(),
                lpszClassName: class,
                ..Default::default()
            };
            // Already registered by an earlier window of this test: the same class, once.
            let _ = unsafe { RegisterClassW(&wc) };

            let style = if parent.is_some() { WS_CHILD } else { WS_POPUP };
            // Safety: the class is registered above, and the window is made where no display is.
            let hwnd = unsafe {
                CreateWindowExW(
                    WINDOW_EX_STYLE::default(),
                    class,
                    w!(""),
                    style,
                    -32000,
                    -32000,
                    320,
                    240,
                    parent.unwrap_or_default(),
                    None,
                    None,
                    None,
                )
            }
            .expect("a window of the test's own");

            self.made.push(hwnd);
            hwnd
        }
    }

    impl Drop for WindowsOfTheTest {
        fn drop(&mut self) {
            // Safety: every handle is a window this test made, and each is destroyed once here.
            unsafe {
                for window in self.made.drain(..) {
                    let _ = DestroyWindow(window);
                }
            }
        }
    }

    /// **A press anywhere inside a folder window is a press in Explorer, whatever part of it the
    /// click landed on.** Every child class the Shell draws a listing with is the window the desktop
    /// names for that press — the view, its rows, the tree, the toolbars, the status bar — so the
    /// answer is the frame's and not the window's own (see `explorer_class_matches` and
    /// `the_frame_of`).
    ///
    /// It is one answer for every question the app asks about a window: the pin's demotion, the
    /// engine's keyboard, a hover's admission, and whether the bubble's key has a listing to come
    /// back in.
    #[test]
    fn a_click_inside_a_listing_is_a_click_on_explorer() {
        let mut windows = WindowsOfTheTest::default();

        // Safety: every window is this test's own, made where no display is and shown never, and the
        // set puts them away when the test ends.
        let (frame, children, another_program) = unsafe {
            let frame = windows.a_window_named(w!("CabinetWClass"), None);
            let view = windows.a_window_named(w!("SHELLDLL_DefView"), Some(frame));
            let rows_drawn_by_the_view = windows.a_window_named(w!("DirectUIHWND"), Some(view));
            let rows = windows.a_window_named(w!("SysListView32"), Some(view));
            let inside_the_views_rows =
                windows.a_window_named(w!("SysListView32"), Some(rows_drawn_by_the_view));
            let tree = windows.a_window_named(w!("SysTreeView32"), Some(frame));
            let toolbar = windows.a_window_named(w!("ToolbarWindow32"), Some(frame));
            let status_bar = windows.a_window_named(w!("msctls_statusbar32"), Some(frame));
            let another_program = windows.a_window_named(w!("RustHoverPreviewWindow"), None);

            (
                frame,
                [
                    view,
                    rows_drawn_by_the_view,
                    rows,
                    inside_the_views_rows,
                    tree,
                    toolbar,
                    status_bar,
                ],
                another_program,
            )
        };

        assert!(
            explorer_class_matches(frame),
            "the frame a folder is drawn in"
        );
        for child in children {
            assert!(
                explorer_class_matches(child),
                "a click inside the listing lands on a child of the frame it is drawn in"
            );
            assert_eq!(
                the_frame_of(child),
                frame,
                "every part of the listing answers the frame it stands in"
            );
        }

        assert!(
            !explorer_class_matches(another_program),
            "a window of another program is not Explorer's, whatever it stands over"
        );
        assert!(
            !explorer_class_matches(HWND::default()),
            "no window is no window"
        );
    }
}

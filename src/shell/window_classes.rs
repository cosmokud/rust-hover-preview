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

use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{
    FindWindowExW, GetAncestor, GetClassNameW, GetDlgItem, GetForegroundWindow, GA_ROOT,
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

/// Whether the window in front is a file listing — Explorer's own, or a Common Item Dialog.
///
/// This is the form the low-level hooks use, because a `WH_KEYBOARD_LL` callback may not take a lock
/// and the hook's own cached form of the question does. It is also the form the preview thread asks:
/// the class of a window is not something that changes under the answer, so the cache the hook keeps
/// is a memo of the lookup rather than a reading of the desktop.
///
/// **Both kinds of listing are answers here, because every caller means "the user is in a listing"** —
/// a pin key pressed over a dialog is a press about the dialog, the bubble's key has a listing to come
/// back in, and the state machine's active path is the one that asks where the pointer is. Explorer's
/// own class alone read a dialog as somebody else's program: the state dropped to
/// `VisibleNotFocused`, the pointer was never probed, and a file the user was looking at in Chrome's
/// `Ctrl`+`O` had no preview at all (see `app::seam::a_listing_is_previewed`, which is where the
/// three kinds of listing are told apart).
///
/// **The dialog half of that is the setting's**, and it is asked of the published fact rather than of
/// the configuration for the reason everything here is: this is called from the keyboard hook, which
/// may take no lock. Off, the answer is the one every caller had before a dialog was a listing at all —
/// a dialog is no listing, and the roads that answer here stay quiet in it (see
/// `app::seam::a_listing_is_previewed`).
pub(crate) fn the_foreground_is_explorer() -> bool {
    unsafe {
        let foreground = GetForegroundWindow();
        if foreground.is_invalid() {
            return false;
        }
        crate::app::seam::a_listing_is_previewed(foreground)
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

/// The control id of the name field every Common Item Dialog carries — `File name:` on an Open or
/// a Save As, `Folder:` on a folder pick. It is what tells a file dialog from the other windows drawn
/// in the same class: a message box has a `Static` and buttons and no name field at all, and a
/// property sheet is a `SysTabControl32` over nested `#32770`s with none either (see
/// `file_dialog_class_matches`).
pub(crate) const THE_NAME_EDIT_CONTROL_ID: i32 = 1148;

/// Whether a window is part of a **Common Item Dialog** — every Vista+ file dialog the machine
/// draws for another program: Chrome's or Notepad's `Ctrl`+`O`, a `Save As`, a folder pick.
///
/// They are listings and are nothing like Explorer's frame. Each is an owned `#32770` in the process
/// that asked for it, and its contents are the shell's own: a `DUIViewWndClassName` holding a
/// `DirectUIHWND`, the rows inside that, the navigation pane and an address breadcrumb. So the folder
/// it is showing is read from the dialog (see `explorer_hook::file_dialog`) and its items are named
/// by the shell the same way a folder view's are — which is why it is a listing the hook has to
/// recognize rather than a window of no interest.
///
/// **The three facts are asked of the frame, and each of them alone would be wrong.** The class
/// `#32770` is the dialog class as a whole — a message box, a property sheet and a print dialog are
/// all of it — so the name field's control id is what narrows it to a *file* dialog, and the
/// `DUIViewWndClassName` is what says the shell is drawing rows in it rather than a message or a page
/// of tabs. A `#32770` without either is one this app has no listing in, and answering yes for it
/// would put a hover over anything a program asks the user.
///
/// One class compare and two window lookups, with no lock and no allocation: a `WH_MOUSE_LL` callback
/// reaches this through `app::seam::a_listing_is_previewed`, so nothing here may wait on anything. It is not memoized
/// here for the same reason the Explorer classes are not: the ask is cheap enough that a hook call
/// paying for a lock to save it would cost more than the lookups, and the hook's own cached form of
/// the question is the memo (`is_explorer_window`).
pub(crate) fn file_dialog_class_matches(hwnd: HWND) -> bool {
    let frame = the_frame_of(hwnd);
    if frame.is_invalid() {
        return false;
    }

    unsafe {
        let mut class_name = [0u16; 256];
        let len = GetClassNameW(frame, &mut class_name);
        if len <= 0 || !utf16_contains_ascii_ignore_case(&class_name[..len as usize], b"#32770") {
            return false;
        }

        // The name field: a message box or a property sheet has none, and a dialog that has one is
        // a dialog asking for a file.
        if GetDlgItem(frame, THE_NAME_EDIT_CONTROL_ID).is_err() {
            return false;
        }

        // And the shell's own row host inside it, which a `#32770` that is a window of its own
        // process's (a message box, a settings page) does not have.
        FindWindowExW(frame, None, w!("DUIViewWndClassName"), PCWSTR::null()).is_ok()
    }
}

// Three kinds of frame are a listing now — Explorer's folder window, a Common Item Dialog and
// Everything's results — and they are three predicates rather than one union, because each is a
// setting of its own where a preview is decided (`app::seam::a_listing_is_previewed`). Anything that
// means "is this Explorer's own folder window" asks `explorer_class_matches`, which stays narrow:
// neither of the others is an Explorer window, and counting one as one would move the window counts
// the pace and the covered-window rule are read from.

/// The class of voidtools Everything's frame, and of the frames of a second instance beside it.
///
/// The bare name is 1.5.0.1423b's own; an instance that is not the first is reported to carry the
/// instance in brackets after it (`EVERYTHING_(1.5a)`), which is read here as the same class because
/// a frame is a frame — and the bracket is what keeps the rest of that program's windows out, since
/// every one of them puts something plainer after the same letters: `EVERYTHING_TASKBAR_NOTIFICATION`
/// is its IPC window and `EVERYTHING_RESULT_LIST_FOCUS` a hidden stand-in for the rows' focus, and
/// neither of them is a window a listing is drawn in.
pub(crate) fn everything_class_matches(hwnd: HWND) -> bool {
    let frame = the_frame_of(hwnd);
    if frame.is_invalid() {
        return false;
    }

    unsafe {
        let mut class_name = [0u16; 256];
        let len = GetClassNameW(frame, &mut class_name);
        len > 0 && a_frame_class_of_everything(&class_name[..len as usize])
    }
}

/// Whether a window class is Everything's frame — see `everything_class_matches` for why the two
/// forms are read as one class and why the bracket is what tells them from the program's others.
///
/// Allocation-free and one pass, because a `WH_MOUSE_LL` callback reaches this through
/// `app::seam::a_listing_is_previewed` and may wait on nothing.
pub(crate) fn a_frame_class_of_everything(class_name: &[u16]) -> bool {
    const THE_FRAME: &[u8] = b"everything";
    const AN_INSTANCE: &[u8] = b"everything_(";

    if class_name.len() < THE_FRAME.len() {
        return false;
    }

    class_name.len() == THE_FRAME.len() && starts_with_ascii_ignore_case(class_name, THE_FRAME)
        || starts_with_ascii_ignore_case(class_name, AN_INSTANCE)
}

/// Whether a class name begins with an ASCII needle, compared one code unit at a time and without
/// allocating: a `u16` outside ASCII can never match a byte of it.
pub(crate) fn starts_with_ascii_ignore_case(haystack: &[u16], needle: &[u8]) -> bool {
    haystack.len() >= needle.len()
        && haystack
            .iter()
            .zip(needle)
            .all(|(ch, byte)| *ch < 0x80 && (*ch as u8).eq_ignore_ascii_case(byte))
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
        CreateWindowExW, DefWindowProcW, DestroyWindow, RegisterClassW, HMENU, WINDOW_EX_STYLE,
        WNDCLASSW, WS_CHILD, WS_POPUP,
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
            self.a_child_named(class, parent, None)
        }

        /// The same window with a **control id**, which is what a dialog's own controls are named by
        /// and what `GetDlgItem` reads them back through: a `#32770` is a real dialog to the desktop
        /// only where its controls carry ids, and the name field's is what tells a file dialog from
        /// a message box (see `file_dialog_class_matches`).
        ///
        /// Safety: as above.
        unsafe fn a_child_named(
            &mut self,
            class: PCWSTR,
            parent: Option<HWND>,
            control_id: Option<i32>,
        ) -> HWND {
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
                    control_id
                        .map(|id| HMENU(id as *mut core::ffi::c_void))
                        .unwrap_or_default(),
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

    /// **A `#32770` is a listing only where it is a file dialog**, and the three windows below are
    /// the ones that prove it: a file dialog's frame carries the name field and the shell's row host,
    /// a message box carries neither, and a property sheet is a page of tabs over the same class. Read
    /// by the class alone, every one of them would be a listing — which is what a hover over a
    /// program's own message would have been, and what the control id and the row host are here to
    /// stop.
    ///
    /// What a dialog is *for* is the other half: it is a listing, and never an Explorer window.
    /// Explorer's own class is what the window counts, the pace and the covered-window rule are read
    /// from, and a dialog counted as one would make the count depend on how many programs happen to
    /// have a file dialog open.
    #[test]
    fn a_file_dialog_is_a_listing_and_never_an_explorer_window() {
        let mut windows = WindowsOfTheTest::default();

        // Safety: every window is this test's own, made where no display is and shown never, and
        // the set puts them away when the test ends.
        let (dialog, rows_host, name_edit, message_box, property_sheet) = unsafe {
            let dialog = windows.a_window_named(w!("#32770"), None);
            let rows_host = windows.a_window_named(w!("DUIViewWndClassName"), Some(dialog));
            let name_edit =
                windows.a_child_named(w!("Edit"), Some(dialog), Some(THE_NAME_EDIT_CONTROL_ID));
            let rows = windows.a_window_named(w!("DirectUIHWND"), Some(rows_host));
            let _ = windows.a_window_named(w!("CtrlNotifySink"), Some(rows));

            // A message box: `Static` and a button, and no name field at all.
            let message_box = windows.a_window_named(w!("#32770"), None);
            let _ = windows.a_child_named(w!("Static"), Some(message_box), Some(0xFFFF));
            let _ = windows.a_child_named(w!("Button"), Some(message_box), Some(1));

            // A property sheet: a tab control over nested pages, and the row host none of them has.
            let property_sheet = windows.a_window_named(w!("#32770"), None);
            let _ = windows.a_child_named(w!("SysTabControl32"), Some(property_sheet), None);

            (dialog, rows_host, name_edit, message_box, property_sheet)
        };

        assert!(
            file_dialog_class_matches(dialog),
            "a dialog with a name field and the shell's rows in it is a file listing"
        );
        for child in [rows_host, name_edit] {
            assert!(
                file_dialog_class_matches(child),
                "a click inside a dialog lands on a child of the frame it is drawn in"
            );
            assert_eq!(
                the_frame_of(child),
                dialog,
                "every part of the dialog answers the frame it stands in"
            );
        }

        assert!(
            !file_dialog_class_matches(message_box),
            "a message box is the same class and no file dialog"
        );
        assert!(
            !file_dialog_class_matches(property_sheet),
            "and a property sheet is tabs rather than rows"
        );
        assert!(
            !everything_class_matches(message_box) && !everything_class_matches(property_sheet),
            "and neither is any other kind of listing either: a hover there would be a hover over a program's own window"
        );

        // **A dialog is a listing and never an Explorer window.** The window counts, the pace and
        // the covered-window rule are Explorer's own, and a dialog is neither counted nor hidden —
        // a dialog in front of a folder window is what that folder window is still visible behind.
        assert!(
            !explorer_class_matches(dialog) && !explorer_class_matches(rows_host),
            "Explorer's own class stays Explorer's"
        );
    }

    /// The three classes a window can be a listing in are told apart by the answer the counts are read
    /// from, and the frames of the rest of the desktop are the same answer they always were.
    #[test]
    fn the_desktop_is_no_listing_of_any_kind() {
        let mut windows = WindowsOfTheTest::default();

        // Safety: this test's own windows, made where no display is and shown never.
        let (folder, desktop, taskbar) = unsafe {
            (
                windows.a_window_named(w!("CabinetWClass"), None),
                windows.a_window_named(w!("Progman"), None),
                windows.a_window_named(w!("Shell_TrayWnd"), None),
            )
        };

        assert!(
            explorer_class_matches(folder),
            "Explorer's own frame is a listing, as it always was"
        );
        assert!(
            !file_dialog_class_matches(folder) && !everything_class_matches(folder),
            "and it is only that kind of listing"
        );
        for window in [desktop, taskbar] {
            assert!(
                !explorer_class_matches(window)
                    && !file_dialog_class_matches(window)
                    && !everything_class_matches(window),
                "the rest of explorer.exe is no listing at all: no hover is over a file there"
            );
        }
    }

    /// **Everything's frame is a listing and never an Explorer window**, and the frames around it in
    /// that program's own process are neither: its IPC window and the hidden stand-in its rows take
    /// their focus from carry the same letters, and reading either as a listing would put a hover over
    /// a window no row is ever drawn in.
    #[test]
    fn an_everything_frame_is_a_listing_of_its_own_kind() {
        let mut windows = WindowsOfTheTest::default();

        // Safety: this test's own windows, made where no display is and shown never.
        let (frame, list, ipc, focus_stand_in) = unsafe {
            let frame = windows.a_window_named(w!("EVERYTHING"), None);
            let list = windows.a_window_named(w!("SysListView32"), Some(frame));
            (
                frame,
                list,
                windows.a_window_named(w!("EVERYTHING_TASKBAR_NOTIFICATION"), None),
                windows.a_window_named(w!("EVERYTHING_RESULT_LIST_FOCUS"), Some(frame)),
            )
        };

        assert!(
            everything_class_matches(frame),
            "Everything's frame is the window its rows are drawn in"
        );
        assert!(
            everything_class_matches(list) && the_frame_of(list) == frame,
            "and a click inside its rows lands on a child of that frame"
        );
        assert!(
            everything_class_matches(focus_stand_in),
            "a hidden stand-in drawn in the frame is part of it, as any child of a listing is"
        );
        assert!(
            !everything_class_matches(ipc),
            "and the program's own IPC window carries the same letters and is nobody's frame"
        );
        assert!(
            !explorer_class_matches(frame) && !file_dialog_class_matches(frame),
            "and it is no Explorer window and no dialog: the counts and the pace stay Explorer's own"
        );

        // A second instance's frame, which is the same name with the instance in brackets after it.
        assert!(a_frame_class_of_everything(
            &"EVERYTHING_(1.5a)".encode_utf16().collect::<Vec<_>>()
        ));
        assert!(a_frame_class_of_everything(
            &"everything".encode_utf16().collect::<Vec<_>>()
        ));
        for not_the_frame in [
            "EVERYTHING_TASKBAR_NOTIFICATION",
            "EVERYTHING_RESULT_LIST_FOCUS",
            "EVERYTHIN",
            "CabinetWClass",
            "",
        ] {
            assert!(
                !a_frame_class_of_everything(&not_the_frame.encode_utf16().collect::<Vec<_>>()),
                "{not_the_frame:?} is not the frame a listing is drawn in"
            );
        }
    }

    /// **The class of a dialog is not the setting, and the setting is what decides whether a preview
    /// is made in one.** `listing_class_matches` keeps answering for a `#32770` file dialog whatever the
    /// switch says — it is asked from a `WH_MOUSE_LL` callback, which may take no lock, and Explorer's
    /// own reading of the desktop must not change under a setting — so the gate is `a_listing_is_previewed`,
    /// which reads the published fact beside the class (see `app::seam::BROWSE_DIALOG_PREVIEWS`).
    ///
    /// Off, a dialog is the window it was before the feature: not a listing, and a folder view is
    /// untouched by the switch — it is Explorer's own class, which the fact is never asked about.
    #[test]
    fn the_dialog_gate_follows_the_published_fact_and_the_class_does_not() {
        let mut windows = WindowsOfTheTest::default();
        let was = crate::app::seam::browse_dialog_previews();

        // Safety: this test's own windows, made where no display is and shown never.
        let (dialog, folder) = unsafe {
            let dialog = windows.a_window_named(w!("#32770"), None);
            let _ = windows.a_window_named(w!("DUIViewWndClassName"), Some(dialog));
            let _ = windows.a_child_named(w!("Edit"), Some(dialog), Some(THE_NAME_EDIT_CONTROL_ID));
            (dialog, windows.a_window_named(w!("CabinetWClass"), None))
        };

        crate::app::seam::publish_browse_dialog_previews(true);
        assert!(
            crate::app::seam::a_listing_is_previewed(dialog),
            "on, a dialog is a listing a file is previewed in"
        );

        crate::app::seam::publish_browse_dialog_previews(false);
        assert!(
            !crate::app::seam::a_listing_is_previewed(dialog),
            "off, it is no listing at all — exactly as it was before the dialog was one"
        );
        assert!(
            file_dialog_class_matches(dialog),
            "and the class still answers for it: a hook callback may take no lock, and the class is \
             the desktop's fact rather than the setting's"
        );
        assert!(
            crate::app::seam::a_listing_is_previewed(folder),
            "a folder view is Explorer's own and the switch is not about it"
        );

        crate::app::seam::publish_browse_dialog_previews(was);
    }
}

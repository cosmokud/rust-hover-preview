use super::everything::{a_folder_of_our_own, a_picture};
use super::*;
use windows::Win32::Foundation::{LRESULT, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, RegisterClassW, WINDOW_EX_STYLE, WNDCLASSW,
    WS_CHILD, WS_POPUP, WS_VISIBLE,
};

/// The cells one row of a Q-Dir pane is drawn with: the name, then the date and the kind its default
/// columns put beside it — no folder, because the pane shows one folder.
fn a_row(name: &str) -> Vec<String> {
    vec![
        name.to_string(),
        "7/5/2026 9:25 PM".to_string(),
        "PNG File".to_string(),
    ]
}

/// **The pane's folder is the whole of what a row is joined onto**, and the file's own entry is what
/// vouches for it — the name alone is a word the pane drew.
#[test]
fn a_name_joined_onto_the_panes_folder_is_the_file() {
    let folder = a_folder_of_our_own("q-dir-joined");
    let picture = a_picture(&folder, "holiday.png");

    assert_eq!(
        row_file_in_folder(&a_row("holiday.png"), &folder),
        normalize_media_path(picture),
        "a name and the folder the pane is showing is the file the row stands for"
    );
}

/// A folder drawn in the pane is a row with nothing to preview, and a name no file answers is no file:
/// `item_path_in_folder` refuses both.
#[test]
fn a_folder_row_and_a_name_no_file_has_answer_nothing() {
    let folder = a_folder_of_our_own("q-dir-refused");
    let inside = folder.join("album");
    std::fs::create_dir_all(&inside).expect("a folder of the test's own");
    a_picture(&inside, "holiday.png");

    assert_eq!(
        row_file_in_folder(&a_row("album"), &folder),
        None,
        "a folder is not a file this app previews"
    );
    assert_eq!(
        row_file_in_folder(&a_row("gone.png"), &folder),
        None,
        "a name with no file behind it is no file"
    );
}

/// **An extensionless name is the shell's hidden-extension listing**, answered by the one media file of
/// that stem — and refused where two files share it, because a name that belongs to two files belongs
/// to neither.
#[test]
fn an_extensionless_name_answers_through_the_one_file_of_that_stem() {
    let folder = a_folder_of_our_own("q-dir-stem");
    let picture = a_picture(&folder, "holiday.png");

    assert_eq!(
        row_file_in_folder(&a_row("holiday"), &folder),
        normalize_media_path(picture),
        "a name the pane draws without an extension is the sole file of that stem"
    );

    a_picture(&folder, "holiday.jpg");
    assert_eq!(
        row_file_in_folder(&a_row("holiday"), &folder),
        None,
        "a second file of one stem is a name that belongs to neither"
    );
}

/// A cell 0 that is a whole path is taken as the path it is, the way a row a shell hand-off names
/// is: the pane's folder is not joined onto it.
#[test]
fn a_cell_that_is_a_whole_path_is_the_file_it_names() {
    let folder = a_folder_of_our_own("q-dir-rooted");
    let picture = a_picture(&folder, "a-picture.png");
    let picture = std::fs::canonicalize(&picture).expect("the file the test made");

    // A folder that is nothing of this file's: the answer cannot have come from the join.
    let other = a_folder_of_our_own("q-dir-rooted-elsewhere");
    assert_eq!(
        row_file_in_folder(&a_row(&picture.to_string_lossy()), &other),
        Some(picture.clone()),
        "an absolute name is the path it names, whatever folder the pane shows"
    );
}

/// A name is a name, whatever it is spelled in: Q-Dir draws the shell's display name and the join is
/// the filesystem's.
#[test]
fn a_unicode_name_resolves() {
    let folder = a_folder_of_our_own("q-dir-unicode");
    let picture = a_picture(&folder, "naïve — 日本語 (.png");

    assert_eq!(
        row_file_in_folder(&a_row("naïve — 日本語 (.png"), &folder),
        normalize_media_path(picture),
        "the name is taken as the pane draws it, and nothing is escaped or decoded"
    );
}

/// **A pane's address bar names a folder, or it names nothing.** It is trimmed and it is absolute and it
/// is there; a placeholder, a URL, a bare relative name and an empty field are none of those, and a path
/// on no drive of this machine is not one either.
#[test]
fn only_an_absolute_folder_that_is_there_is_the_panes_folder() {
    let folder = a_folder_of_our_own("q-dir-address");

    assert_eq!(
        pane_folder_from_edit_text(&folder.to_string_lossy()),
        Some(folder.clone()),
        "an absolute folder is the folder the pane is showing"
    );
    assert_eq!(
        pane_folder_from_edit_text(&format!("  {}  ", folder.to_string_lossy())),
        Some(folder.clone()),
        "and its whitespace is the field's, not the path's"
    );
    assert_eq!(
        pane_folder_from_edit_text(&format!("{}\\", folder.to_string_lossy())),
        Some(folder.clone()),
        "a trailing separator is where the shell puts one, and the folder is the same folder"
    );

    for not_a_folder in [
        &*folder.join("nothing-here").to_string_lossy(),
        "&Abbrechen",
        "",
        "   ",
        "https://example.test/a",
        "holiday.png",
        "C:",
    ] {
        assert_eq!(
            pane_folder_from_edit_text(not_a_folder),
            None,
            "{not_a_folder:?} is not a folder a row could be joined onto"
        );
    }

    let _ = std::fs::remove_dir_all(&folder);
}

/// A window of this test's own, in a class it names: the tests below are about a frame holding lists
/// and about which of them answers for the keyboard, and not about a pixel of any of them.
///
/// Safety: the test puts every window it makes away, and none is ever shown.
fn a_window(class: windows::core::PCWSTR, parent: Option<HWND>, visible: bool) -> HWND {
    unsafe extern "system" fn an_empty_window(
        hwnd: HWND,
        msg: u32,
        wparam: WPARAM,
        lparam: LPARAM,
    ) -> LRESULT {
        // Safety: the default procedure answers a message this window's own class never handles.
        unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
    }

    let instance = unsafe { GetModuleHandleW(None) }.expect("this module");
    let wc = WNDCLASSW {
        lpfnWndProc: Some(an_empty_window),
        hInstance: instance.into(),
        lpszClassName: class,
        ..Default::default()
    };
    let _ = unsafe { RegisterClassW(&wc) };

    let style = if parent.is_some() {
        WS_CHILD | WS_VISIBLE
    } else {
        WS_POPUP | WS_VISIBLE
    };
    debug_assert!(visible, "a list with no room drawn in it is no pane");
    unsafe {
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
    .expect("a window of the test's own")
}

/// The chain one pane is drawn in, as 12.83 draws it: the pane's own `WTL_XSplitterWindow`, its
/// `#32770`, the three `ATL:` hosts, and the `SHELLDLL_DefView` the rows are in (see the note at the top
/// of `q_dir`). Handed the frame, it answers the pane's list and every window on the way to it.
///
/// Safety: the caller puts every window it makes away, and none is ever shown.
fn a_pane_of(frame: HWND) -> (HWND, HWND) {
    let splitter = a_window(w!("WTL_XSplitterWindow"), Some(frame), true);
    let dialog = a_window(w!("#32770"), Some(splitter), true);
    let host = a_window(w!("ATL:00000001401C3970"), Some(dialog), true);
    let inner = a_window(w!("ATL:00000001401C7BB0"), Some(host), true);
    let def_view = a_window(w!("SHELLDLL_DefView"), Some(inner), true);
    let list = a_window(w!("SysListView32"), Some(def_view), true);

    (splitter, list)
}

/// **The pane the keyboard is on is the one whose list is the focused window of Q-Dir's thread**, and a
/// frame of four panes is why the question is asked at all: the panes beside it keep the row they were
/// last left selected, and previewing the file of one of those is previewing a file the user did not
/// pick (see `focused_list`).
///
/// The focus is put on the list by the test itself, which is Q-Dir's own arrangement: the pane the user
/// last walked holds the keyboard, and the pane a click moved to is the one holding it since.
#[test]
fn the_pane_holding_the_focus_is_the_one_the_keyboard_is_on() {
    use windows::Win32::UI::Input::KeyboardAndMouse::SetFocus;

    let was = crate::app::seam::q_dir_previews();
    crate::app::seam::publish_q_dir_previews(true);

    // Safety: this test's own windows, made where no display is and shown never.
    let frame = a_window(w!("ATL:00000001401C5A80"), None, true);
    let (left_pane, left_list) = a_pane_of(frame);
    let (right_pane, right_list) = a_pane_of(frame);

    // The keyboard in the left pane: the frame answers that pane's list and nothing of the right
    // pane's, which is the whole of the rule.
    let _ = unsafe { SetFocus(left_list) };
    assert_eq!(
        focused_list(frame),
        Some(left_list),
        "the pane holding the keyboard is the one the frame answers with"
    );
    assert_eq!(
        focused_list(left_pane),
        Some(left_list),
        "and a window of the pane is answered the same way as the frame it stands in"
    );
    assert_eq!(
        focused_list(right_pane),
        None,
        "a pane that holds no keyboard is no listing: its rows are the rows it was last left on"
    );

    // The keyboard moving to the right pane answers the right pane, and never the pane it came from.
    let _ = unsafe { SetFocus(right_list) };
    assert_eq!(
        focused_list(frame),
        Some(right_list),
        "the pane the keyboard has moved to is the one it is in"
    );

    // And a window of the frame that is no list draws no rows at all: the pane's chrome is not a
    // listing, so no file is answered for it.
    let _ = unsafe { SetFocus(right_pane) };
    assert_eq!(
        focused_list(frame),
        None,
        "chrome holding the keyboard names no pane's rows, and nothing is guessed at"
    );

    for window in [left_list, right_list, left_pane, right_pane, frame] {
        // Safety: the windows are this test's own and are put away here.
        let _ = unsafe { DestroyWindow(window) };
    }

    crate::app::seam::publish_q_dir_previews(was);
}

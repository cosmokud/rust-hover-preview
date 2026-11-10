//! **Q-Dir**'s panes: a listing that is nobody's Shell window either, and the second one read out of
//! its own address space (see `list_view`).
//!
//! **A frame holds four panes.** The frame's class is `ATL:00000001401C5A80` — an ATL host's own
//! class, whose name carries the address of its window procedure — and each pane is
//! `WTL_XSplitterWindow`×3 → `#32770` → three `ATL:` windows → `SHELLDLL_DefView` → `SysListView32`.
//! So the list is always found **from the item's or the point's own window** and never as "the first
//! list in the frame": four panes, four lists, and answering with the first would preview the file
//! drawn in another pane at the same place.
//!
//! `Shell.Windows()` holds nothing for those panes and their rows carry no `ItemIndex` and no path, so
//! what replaces the walk is the list's own messages (see `list_view`) — and the folder a row is in is
//! the pane's address bar: the visible `Edit` inside the same `#32770` the list is in, whose text is
//! the absolute path that pane is showing. Q-Dir navigates by rewriting that `Edit`, which is what the
//! memo's TTL covers.
//!
//! A row's cell 0 is the shell's display name, so it carries the extension only while Windows shows
//! one: a name with no extension is answered the way a dialog's is, by the sole media file of that stem
//! in the pane's folder (see `file_dialog::item_path_in_folder`).

use windows::core::PCWSTR;
use windows::Win32::UI::WindowsAndMessaging::{
    EnumChildWindows, FindWindowExW, GetParent, GetWindowTextW,
};

use super::*;

/// The `#32770` a pane's rows are drawn in, out of the list itself.
fn the_pane_container(list: HWND) -> Option<HWND> {
    let mut at = list;
    for _ in 0..THE_PANE_ANCESTOR_LIMIT {
        // Safety: asks the desktop about a handle it gave out, and dereferences nothing.
        let parent = unsafe { GetParent(at) }.ok()?;
        if parent.is_invalid() {
            return None;
        }
        at = parent;

        let mut class_name = [0u16; 256];
        // Safety: as above, into this function's own buffer.
        let len = unsafe { GetClassNameW(at, &mut class_name) };
        if len > 0
            && crate::shell::window_classes::utf16_contains_ascii_ignore_case(
                &class_name[..len as usize],
                b"#32770",
            )
        {
            return Some(at);
        }
    }

    None
}

/// How far up a pane's own chrome the search for its `#32770` walks: the measured chain from the list
/// to it is five windows long, and a frame is nowhere near this high.
const THE_PANE_ANCESTOR_LIMIT: usize = 16;

/// The folder one pane is showing, read out of the address bar of the container its rows are in.
///
/// The `Edit` is the one *visible* one whose text is an absolute path that is there: the frame's other
/// `ComboBox`es are history dropdowns, and a pane's own address bar is
/// `ComboBoxEx32` → `ComboBox` → `Edit` inside its own `#32770`.
fn the_pane_folder(list: HWND) -> Option<PathBuf> {
    struct TheSearch {
        found: Option<PathBuf>,
    }

    unsafe extern "system" fn the_address_edit(hwnd: HWND, lparam: LPARAM) -> BOOL {
        let search = &mut *(lparam.0 as *mut TheSearch);
        if search.found.is_some() {
            return BOOL(0);
        }

        let mut class_name = [0u16; 256];
        // Safety: the class name is this window's and is written into this callback's own buffer.
        let len = GetClassNameW(hwnd, &mut class_name);
        if len <= 0
            || !crate::shell::window_classes::utf16_contains_ascii_ignore_case(
                &class_name[..len as usize],
                b"edit",
            )
        {
            return BOOL(1);
        }

        // Safety: asks the desktop about a window it enumerated, and dereferences nothing.
        if !IsWindowVisible(hwnd).as_bool() {
            return BOOL(1);
        }

        let mut text = [0u16; 1024];
        // Safety: as above, into this callback's own buffer.
        let len = GetWindowTextW(hwnd, &mut text).max(0) as usize;
        let text = String::from_utf16_lossy(&text[..len.min(text.len())]);
        let Some(folder) = pane_folder_from_edit_text(&text) else {
            return BOOL(1);
        };

        search.found = Some(folder);
        BOOL(0)
    }

    let container = the_pane_container(list)?;
    let mut search = TheSearch { found: None };
    // Safety: the callback is a plain function of this module's, and it is handed a pointer to a
    // struct that lives on this frame until the enumeration is over.
    unsafe {
        let _ = EnumChildWindows(
            container,
            Some(the_address_edit),
            LPARAM(&mut search as *mut TheSearch as isize),
        );
    }

    search.found
}

/// The folder an address bar's text names, or nothing where it names none.
///
/// Trimmed and absolute and there, and nothing else: a pane's bar also reads `&Abbrechen` under a
/// German shell, `https://…` under a view of the web, and a bare relative name while the user is
/// still typing — none of those is a folder a row could be joined onto.
pub(super) fn pane_folder_from_edit_text(text: &str) -> Option<PathBuf> {
    let text = text.trim();
    if text.is_empty() {
        return None;
    }

    let path = Path::new(text);
    if !path.is_absolute() || !path.is_dir() {
        return None;
    }

    Some(path.to_path_buf())
}

/// The folder the list's pane is showing, memoized against **the list** for `FOLDER_PROBE_MS`.
///
/// **Against the list rather than the frame**, because a frame holds four panes and four folders: one
/// answer for the frame would join a row of one pane onto another pane's folder. The refresh is the
/// dialog's (`file_dialog::dialog_folder`) and for the same reason: the read is a walk of another
/// process's windows, so it is kept, and Q-Dir navigates by rewriting that `Edit` rather than by moving
/// its windows — which is a change nothing but this TTL and `forget_window_views` can see.
pub(super) fn pane_folder(resolver: &mut ItemResolver, list: HWND) -> Option<PathBuf> {
    let list_key = list.0 as isize;

    let kept = resolver
        .pane_folder
        .as_ref()
        .filter(|memo| {
            memo.frame == list_key
                && memo.read_at.elapsed() < Duration::from_millis(FOLDER_PROBE_MS)
        })
        .map(|memo| memo.answer.clone());
    if let Some(answer) = kept {
        return answer;
    }

    let answer = the_pane_folder(list);
    resolver.pane_folder = Some(DialogFolderMemo {
        frame: list_key,
        answer: answer.clone(),
        read_at: Instant::now(),
    });

    answer
}

/// The file a row's cells name in the folder the pane is showing, or nothing where they name none.
///
/// Cell 0 is the name, and everything after it is a date or a kind rather than a folder — the pane
/// shows one folder, so the folder is the pane's and not the row's. An absolute name is taken as the
/// path it is; anything else is joined the way a dialog's display name is (`item_path_in_folder`),
/// which refuses a folder row, gates the join on the file's own entry, and answers an extensionless
/// name through the sole media file of that stem.
pub(super) fn row_file_in_folder(cells: &[String], folder: &Path) -> Option<PathBuf> {
    let name = cells.first()?.trim();
    if name.is_empty() {
        return None;
    }

    if Path::new(name).is_absolute() {
        return normalize_media_path(PathBuf::from(name));
    }

    item_path_in_folder(folder, name)
}

/// The file the item `item_file_path` was handed stands for, in the Q-Dir pane it was read in: the row
/// the item names, joined onto the folder that pane is showing (see `row_file_in_folder`).
pub(super) fn item_path(
    resolver: &mut ItemResolver,
    window: &ItemWindow,
    item: &HoveredItem,
) -> Option<PathBuf> {
    // The switch is read here as well as where the window was classified, because a tick that
    // classified a Q-Dir window while it was on must answer nothing after it is turned off (see
    // `app::seam::a_q_dir_is_previewed`).
    if !crate::app::seam::q_dir_previews() {
        return None;
    }

    // The pane the item is drawn in, from the item's own window: a frame holds four of them.
    let drawn_in = if window.drawn_in != 0 {
        window.drawn_in
    } else {
        item.native_window
    };
    let list = the_q_dir_list_of(HWND(drawn_in as *mut core::ffi::c_void))?;
    let folder = pane_folder(resolver, list)?;
    let answer = row_of_list(resolver, list, item.index? - 1)
        .and_then(|row| row_file_in_folder(&row.cells, &folder));

    // The row the answer was read from against the file it named, because an unanswered row is the
    // ordinary case here — a folder, a name no file answers, a kind nothing previews.
    note_pin_click!(
        "q-dir item  {}  in {}  -> {}",
        item.name,
        folder.display(),
        answer
            .as_deref()
            .map_or_else(|| "-".to_string(), |path| path.display().to_string()),
    );

    answer
}

/// The list a pane draws its rows in, out of the window an item of it reports as its own.
///
/// **The pane is found from the window handed in, never from the frame.** A frame holds four panes, so
/// "the first list in the frame" is the file drawn in another pane — the one under the pointer is the
/// one whose list is above that point, and each ancestor of it is asked for its own
/// `SHELLDLL_DefView`/`SysListView32` until the frame is reached.
pub(super) fn the_q_dir_list_of(window: HWND) -> Option<HWND> {
    if window.is_invalid() || !crate::shell::window_classes::q_dir_class_matches(window) {
        return None;
    }

    let frame = crate::shell::window_classes::the_frame_of(window);
    let mut at = window;
    loop {
        if let Some(list) = a_list_under(at) {
            return Some(list);
        }
        if at == frame {
            return None;
        }

        // Safety: asks the desktop for the parent of a handle it gave out; dereferences nothing.
        let parent = unsafe { GetParent(at) }.ok()?;
        if parent.is_invalid() {
            return None;
        }
        at = parent;
    }
}

/// The list the point's own window is in, where that window is in a Q-Dir pane.
pub(super) fn the_q_dir_list_under(point: POINT) -> Option<HWND> {
    // Safety: asks the desktop what is drawn at a point of the display; dereferences nothing.
    the_q_dir_list_of(unsafe { WindowFromPoint(point) })
}

/// The visible `SysListView32` below `ancestor`, through the `SHELLDLL_DefView` the pane draws its
/// rows in — or nothing where that ancestor holds no pane.
fn a_list_under(ancestor: HWND) -> Option<HWND> {
    // Safety: two lookups of a window the desktop already holds, by class and nothing else.
    let def_view =
        unsafe { FindWindowExW(ancestor, None, w!("SHELLDLL_DefView"), PCWSTR::null()) }.ok()?;
    if def_view.is_invalid() {
        return None;
    }

    // Safety: as above.
    let list =
        unsafe { FindWindowExW(def_view, None, w!("SysListView32"), PCWSTR::null()) }.ok()?;
    // The visible one with room to draw a row: a pane that is not showing draws no rows, and a list
    // with no room in it is not the one the user is pointing at (see `is_a_list_to_read`).
    is_a_list_to_read(list).then_some(list)
}

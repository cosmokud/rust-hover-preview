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
    EnumChildWindows, FindWindowExW, GetGUIThreadInfo, GetParent, GetWindowThreadProcessId,
    SendMessageTimeoutW, GUITHREADINFO, SMTO_ABORTIFHUNG, SMTO_BLOCK,
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

        // **`WM_GETTEXT`, and not `GetWindowTextW`.** The address bar is another process's control
        // and Q-Dir keeps its text in the control rather than in the window's own title, which is
        // what `GetWindowTextW` answers for a window of somebody else's process: it reads the
        // desktop's copy of the title and not the control's, and it answers nothing here. A message
        // is what crosses processes, and it is bounded so that a pane busy painting answers nothing
        // rather than holding the tick.
        let mut text = [0u16; 1024];
        let mut answer = 0usize;
        // `wParam` is how many characters may be copied, and a message that copies none answers
        // empty whatever the bar says.
        let sent = unsafe {
            SendMessageTimeoutW(
                hwnd,
                0x000D,
                windows::Win32::Foundation::WPARAM(text.len()),
                LPARAM(text.as_mut_ptr() as isize),
                SMTO_BLOCK | SMTO_ABORTIFHUNG,
                THE_ADDRESS_BAR_MS,
                Some(&mut answer),
            )
        };
        if sent.0 == 0 {
            return BOOL(1);
        }

        let len = answer.min(text.len());
        let text = String::from_utf16_lossy(&text[..len]);
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

/// How long the address bar has to answer before its pane is left for the next tick.
///
/// Its own number rather than the row reader's (`list_view::ANSWER_MS`), because the two failures are
/// not the same failure: a list that will not answer is a listing that cannot be read at all, while a
/// bar that will not answer is a pane whose folder is not known yet — the rows are still there, and
/// the next tick asks again.
const THE_ADDRESS_BAR_MS: u32 = 250;

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

/// The list the **keyboard's own** focus is on, out of the frame the keyboard is driving — or nothing
/// where that frame is no Q-Dir frame, or where no pane is holding the focus.
///
/// **A pane holds the keyboard itself.** Q-Dir puts the file's own `SysListView32` down as the focused
/// window of its thread, so `GetGUIThreadInfo` names the list a key pressed at that moment goes to. The
/// panes beside it are drawn the same way and hold no focus, which is what tells the pane the user is
/// walking from the panes that merely have a row selected from before: Q-Dir leaves a pane's selection
/// exactly as it was when **a click in another pane** moved to it, and one of those is drawn in the
/// shell's grey inactive colour rather than the blue of the listing the keyboard is in (measured on
/// 12.83: `active` and `focus` both name the pane's list, and every other pane is a list with a row
/// still selected).
///
/// The focus is read of the **thread** rather than of the foreground window, because the two are not the
/// same question: `GetFocus` answers for the thread that calls it — the app's own — and the window in
/// front is not the thread's to answer for.
///
/// A frame whose list the keyboard is in and whose panes say nothing about a focus answers nothing here,
/// which is left to the caller rather than to a second rule: a pane that is not the one the keyboard is
/// in must never be previewed in, so a guess is worse than no preview at all.
pub(super) fn focused_list(hwnd: HWND) -> Option<HWND> {
    let frame = crate::shell::window_classes::the_frame_of(hwnd);
    if frame.is_invalid() || !crate::app::seam::a_q_dir_is_previewed(frame) {
        return None;
    }

    let mut q_dir_pid = 0u32;
    // Safety: asks the desktop which thread draws the frame it gave the handle of, into this
    // function's own integer and dereferencing nothing.
    let thread = unsafe { GetWindowThreadProcessId(frame, Some(&mut q_dir_pid)) };
    if thread == 0 {
        return None;
    }

    let mut info = GUITHREADINFO {
        cbSize: std::mem::size_of::<GUITHREADINFO>() as u32,
        ..Default::default()
    };
    // Safety: the struct is this frame's own and is filled by the call it is handed to; it names no
    // window of another process's to be read through, it is asked about one.
    if unsafe { GetGUIThreadInfo(thread, &mut info) }.is_err() {
        return None;
    }

    let focus = info.hwndFocus;
    if focus.is_invalid() || crate::shell::window_classes::the_frame_of(focus) != frame {
        return None;
    }

    // The list the focused window **is**, or the one it is drawn in: Q-Dir leaves the pane its whole
    // focus — the list on 12.83 — and a shell that puts the focus on a window holding it is the same
    // pane either way.
    let list = the_q_dir_list_of(focus)?;

    // **And the window handed in is answered only where it is that pane's own.** A frame is asked about
    // whichever pane holds the keyboard, and a window inside one of its panes is asked about that pane:
    // the other panes hold no keyboard, so nothing of theirs is a listing a key was pressed in.
    let relates = hwnd_is_same_or_ancestor(list, hwnd) || hwnd_is_same_or_ancestor(hwnd, list);

    relates.then_some(list)
}

/// The list a pane draws its rows in, out of the window an item of it reports as its own.
///
/// **The pane is found from the window handed in, never from the frame.** A frame holds four panes, so
/// "the first list in the frame" is the file drawn in another pane — the one under the pointer is the
/// one whose list is above that point, and each ancestor of it is asked for its own
/// `SHELLDLL_DefView`/`SysListView32` until the frame is reached.
pub(super) fn the_q_dir_list_of(window: HWND) -> Option<HWND> {
    // Asked in the setting's own form rather than of the class alone, so that a Q-Dir pane is read no
    // further the moment `Enable in Q-Dir` is off (see `app::seam::a_q_dir_is_previewed`).
    if window.is_invalid()
        || !crate::app::seam::a_q_dir_is_previewed(crate::shell::window_classes::the_frame_of(
            window,
        ))
    {
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

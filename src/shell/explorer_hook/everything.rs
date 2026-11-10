//! voidtools **Everything**'s result list: a listing that is nobody's Shell window, read out of
//! Everything's own address space.
//!
//! **Nothing about this listing is a position the rest of the hook can ask for.** Everything's frame
//! is not a Shell window, so `Shell.Windows()` has nothing for it, no Shell view is registered for
//! it, and its rows are a plain `SysListView32` of another process whose items carry no `ItemIndex`
//! and no path. A search also spans folders, so there is no folder a name could be resolved against
//! either: the row's own cells are the only witness there is, which is why they are the whole of what
//! this module reads (see `row_file`).
//!
//! What replaces the walk is the list's own messages, through the shared reader in `list_view` (see
//! there for the page, the messages and the cost).
//!
//! **And they are never read through UI Automation.** Reading that list through UIA is reported to
//! hang Everything and to leave it holding gigabytes, so every road that could reach
//! `ElementFromPoint`/`GetFocusedElement` asks the frame's class first and comes here instead — see
//! `uia_item_from_point` and `uia_item_from_focus`, which is where every walk starts.

use super::*;

/// The Everything frame a window stands in, where the app previews in that listing.
///
/// It is the branch every road takes before it asks UI Automation anything, and it is taken on the
/// class of the frame alone: the ask is a `GetAncestor` and two class compares, which is what keeps
/// the answer out of a walk that must never be made in this window (see the note at the top of this
/// module, and `app::seam::an_everything_is_previewed` for the setting beside the class).
pub(super) fn the_everything_frame_of(window: HWND) -> Option<isize> {
    let frame = crate::shell::window_classes::the_frame_of(window);
    crate::app::seam::an_everything_is_previewed(frame).then_some(frame.0 as isize)
}

/// One row of the Everything list under `frame`, read by its index — zero-based, as the list numbers
/// its rows.
pub(super) fn row_of(resolver: &mut ItemResolver, frame: isize, index: i32) -> Option<ListViewRow> {
    let list = the_list_in(HWND(frame as *mut core::ffi::c_void))?;
    row_of_list(resolver, list, index)
}

/// The row a point on the screen is on, in the Everything list under `frame`.
pub(super) fn row_at(
    resolver: &mut ItemResolver,
    frame: isize,
    point: POINT,
) -> Option<ListViewRow> {
    let list = the_list_in(HWND(frame as *mut core::ffi::c_void))?;
    row_at_in_list(resolver, list, point)
}

/// The row the keyboard is on in the Everything list under `frame`.
pub(super) fn focused_row(resolver: &mut ItemResolver, frame: isize) -> Option<ListViewRow> {
    let list = the_list_in(HWND(frame as *mut core::ffi::c_void))?;
    focused_row_in_list(resolver, list)
}

/// The file a row's cells name, or nothing where they name none.
///
/// **The row's own text is the only witness there is.** Everything holds no path for a row beyond the
/// columns it draws, and which columns those are is the user's arrangement: so the name is cell 0 and
/// the folder is the first other cell that *is* one — an absolute path that is a directory, which a
/// size (`51,964 KB`), a date or a count never is — and the join is gated the way every route here is
/// gated (`normalize_media_path` reads the file's own entry, so a path is only taken where the file is
/// there, is a file, and is a kind this build shows).
///
/// A folder, a file no kind previews, and a row whose columns show no folder all answer nothing.
/// Nothing is looked up by name in any folder: a search's rows come from every folder the user pointed
/// it at, and the same name in two of them is two files.
pub(super) fn row_file(cells: &[String]) -> Option<PathBuf> {
    let name = cells.first()?.trim();
    if name.is_empty() {
        return None;
    }

    let folder = cells.iter().skip(1).find_map(|cell| {
        let cell = cell.trim();
        let path = Path::new(cell);
        (!cell.is_empty() && path.is_absolute() && path.is_dir()).then_some(path)
    })?;

    normalize_media_path(folder.join(name))
}

/// The file the item `item_file_path` was handed stands for, in the Everything listing it was read in:
/// the row the item names, and the file that row's cells name (see `row_file`).
///
/// It is the one place an Everything item becomes a file, so the pointer's row, the keyboard's row and
/// any walk by index all answer through it.
pub(super) fn item_path(
    resolver: &mut ItemResolver,
    window: &ItemWindow,
    item: &HoveredItem,
) -> Option<PathBuf> {
    // The switch is read here as well as where the window was classified, because a tick that
    // classified an Everything window while it was on must answer nothing after it is turned off (see
    // `app::seam::an_everything_is_previewed`).
    if !crate::app::seam::everything_previews() {
        return None;
    }

    let answer =
        row_of(resolver, window.frame, item.index? - 1).and_then(|row| row_file(&row.cells));

    // The row the answer was read from against the file it named, because a row that names no file is
    // the ordinary case here — a folder, a kind nothing previews, a column set without the folder's
    // path — and reads here as `-` rather than as a run that never looked.
    note_pin_click!(
        "everything item  {}  -> {}",
        item.name,
        answer
            .as_deref()
            .map_or_else(|| "-".to_string(), |path| path.display().to_string()),
    );

    answer
}

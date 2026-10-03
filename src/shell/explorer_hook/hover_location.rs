//! Where the pointer is, and what place it is over: one reading of the cursor per tick,
//! and the four facts a view describes itself with.
//!
//! Those facts are compared one field at a time rather than as the single key they used
//! to be written as, because they are not equally reliable: the URL answers every time
//! and the folder is a walk out to a filesystem path. Compared as one string, the same
//! place read twice reads as a change, and a preview of a file that never moved is
//! taken down and put back a moment later.

use super::*;

/// What a look at the view under the pointer said about the place it is showing, one
/// fact per field: the folder the view has open, the search root under it, the URL it
/// was opened with, and the view's own window.
///
/// It stands where a single formatted key did — the first fact that answered, written
/// as a string — and the difference between the two is the whole of it: the facts are
/// not equally reliable. The URL is the browser object's own and answers every time,
/// while the folder is a walk out through the shell's objects to a filesystem path,
/// and a path on a share, on a slow disk or in a library answers nothing to that walk
/// on some looks and a path on others. Read as one key, the same place came out as
/// `folder:–` on one look and `url:–` on the next, and each was read as a *change*:
/// the preview of a file that never moved was taken down, the gate armed, and the same
/// preview put back a moment later, which is the blink. Kept apart, a fact is compared
/// only where both looks answered it — see [`hover_location_changed`].
#[derive(Clone, Default)]
pub(super) struct HoverLocation {
    pub(super) folder: Option<String>,
    pub(super) search_root: Option<String>,
    pub(super) location_url: Option<String>,
    pub(super) view_hwnd: Option<isize>,
}

impl HoverLocation {
    /// The place a look's hints describe, as the facts that look managed to read.
    ///
    /// A fact is read into the form it is compared in — see `location_fact_key` — so that
    /// the Shell spelling one place another way on the next look is not a difference of
    /// place. What is stored is that form rather than what the Shell said, because the
    /// only thing a stored fact is for is being compared with the next look's.
    pub(super) fn of(hints: &HoverResolverHints) -> Self {
        Self {
            folder: hints.current_folder.as_deref().map(location_fact_key),
            search_root: hints.search_root.as_deref().map(location_fact_key),
            location_url: hints.location_url.as_deref().map(location_fact_key),
            view_hwnd: hints.shell_view_hwnd,
        }
    }

    /// Whether the look answered anything about the place at all. A look that read
    /// none of the four is not a place to compare against: it is a shell that said
    /// nothing, and nothing follows from it either way.
    pub(super) fn was_answered(&self) -> bool {
        self.folder.is_some()
            || self.search_root.is_some()
            || self.location_url.is_some()
            || self.view_hwnd.is_some()
    }
}

/// The form a location fact is compared in, so that two answers describing one place are
/// one answer.
///
/// A fact out of the Shell is the Shell's own spelling of it, and the Shell does not spell
/// one place the same way on every look. The folder a view has open is canonicalized where
/// that succeeds — which is where the verbatim `\\?\` form comes from — and is left as the
/// view reported it where it does not, so the same folder answers `\\?\D:\Pictures` on one
/// look and `D:\Pictures` on the next, on the volumes where canonicalizing is the thing
/// that fails from time to time. Read as it comes, that difference is a difference of
/// *place*, and the preview of a file that never moved is taken down on the look that
/// happens to answer the other spelling — and put back on the one after it, which is a
/// preview that blinks at a pointer which has not moved at all. Case is the other half of
/// the same coin: Windows reads two spellings of a path as one path, and so does a fact
/// that is one, and a trailing separator names the place named without it.
///
/// What is *not* normalized is a search's own URL: a query's text is the query, and two
/// searches that differ in case are two searches.
pub(super) fn location_fact_key(fact: &str) -> String {
    let trimmed = fact.trim();
    let unverbatim = match trimmed.strip_prefix(r"\\?\UNC\") {
        Some(share) => format!(r"\\{share}"),
        None => trimmed.strip_prefix(r"\\?\").unwrap_or(trimmed).to_string(),
    };

    let unrooted = unverbatim.trim_end_matches(['\\', '/']);
    let unrooted = if unrooted.is_empty() {
        unverbatim.as_str()
    } else {
        unrooted
    };

    if is_search_ms_url(unrooted) {
        unrooted.to_string()
    } else {
        unrooted.to_ascii_lowercase()
    }
}

/// Whether two looks at the view describe different places.
///
/// A fact counts only where *both* looks answered it. A look that could not walk the
/// folder out of the view leaves that fact unanswered rather than answering another
/// one, and two places are not told apart by one of them failing to say where it is:
/// reading a missing answer as a different place is what blinked the preview. Where
/// both looks did answer, a difference is a move — another folder, another search,
/// another tab of the same window — and is read as the change it is.
pub(super) fn hover_location_changed(previous: &HoverLocation, current: &HoverLocation) -> bool {
    /// Whether two looks disagree about one fact, where both of them read it.
    fn differs<T: PartialEq>(previous: &Option<T>, current: &Option<T>) -> bool {
        match (previous.as_ref(), current.as_ref()) {
            (Some(previous), Some(current)) => previous != current,
            _ => false,
        }
    }

    differs(&previous.folder, &current.folder)
        || differs(&previous.search_root, &current.search_root)
        || differs(&previous.location_url, &current.location_url)
        || differs(&previous.view_hwnd, &current.view_hwnd)
}

/// Make `place` the baseline in `slot`, and answer whether it is a different place from the one
/// that was in hand: a landing rather than a pick, asked of the keyboard's item and of the
/// listing's own selection alike.
pub(super) fn take_place(slot: &mut Option<HoverLocation>, place: HoverLocation) -> bool {
    let moved = slot
        .as_ref()
        .is_some_and(|previous| hover_location_changed(previous, &place));
    *slot = Some(place);

    moved
}

pub(super) fn is_pressed_or_down_state(state: u16) -> bool {
    (state & 0x8000) != 0 || (state & 0x0001) != 0
}

/// Whether the resolved off-trigger virtual key is currently held.
pub(super) fn key_is_down(vk: i32) -> bool {
    unsafe {
        let state = GetAsyncKeyState(vk) as u16;
        (state & 0x8000) != 0
    }
}

/// What one tick knows about the pointer: where it is, the scale of the display it is on,
/// and the window it is over.
///
/// One reading of each per tick rather than one per caller, because every question a tick
/// asks about the pointer is a question about one instant: what the hand has done is one
/// reading of "the mouse has moved", what is under it is one window, and the display it is
/// on is one monitor. A caller that reads the cursor for itself in the middle of a tick is
/// asking about a later instant than the tick it belongs to — and paying a syscall for the
/// privilege. The one read that is deliberately fresh is the one the preview loop makes
/// before it lays a hover out (see `replay_where_the_pointer_is`): that one is about where
/// the preview goes, and it happens on the thread that draws it.
#[derive(Clone, Copy)]
pub(super) struct PointerTick {
    pub(super) point: POINT,
    /// The scale of the display the pointer is on, which is what a move is measured against
    /// at the tolerance it is given (see `KeyboardPointerPause::move_threshold_px`).
    pub(super) dpi: u32,
    /// The window under the pointer, leaf first: what a frame's views are told apart by
    /// (see `ItemWindow`).
    pub(super) window: HWND,
}

impl PointerTick {
    /// The pointer as one reading, from a cursor position already in hand.
    ///
    /// It is the one place a snapshot is built: the scale of the display is asked of the
    /// display the point is on, and the window is asked of the point itself, so a caller
    /// that has read the cursor is a caller that has everything else here (see
    /// `read_pointer` for the callers that have not).
    pub(super) fn of(point: POINT) -> Self {
        Self {
            point,
            dpi: monitor_dpi_from_point(point.x, point.y),
            window: unsafe { WindowFromPoint(point) },
        }
    }
}

/// Where the pointer is and what is under it, as one reading.
pub(super) fn read_pointer() -> Option<PointerTick> {
    let mut point = POINT::default();
    if unsafe { GetCursorPos(&mut point) }.is_err() {
        return None;
    }

    Some(PointerTick::of(point))
}

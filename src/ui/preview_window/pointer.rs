//! What the pointer's own questions are answered from: the box of the item it is over, the
//! regions a press is held for, and the scrollbar a hand can drag.

use super::*;

/// The item the hover on screen is about, as the box the view draws it in: the
/// hook's own answer about the pointer, published with every look at the item under
/// it and withdrawn with the window.
///
/// The count above covers a load that lands after a hide has been *sent*. It cannot
/// cover the moment before one, and that is where the flash the eye catches lives:
/// the hook decides a hover on a tick of its own while the frame lands on a tick of
/// this loop's, and a pointer can cross a whole row of the list in between — a file
/// previewed, and taken down again by the hook's very next look. So the reveal is
/// asked one more question, of the pointer itself: the box says which item the hover
/// was resolved from, a pointer outside it is a hover that has moved on, and the
/// frame that lands for one is dropped rather than revealed — the hook's next look
/// answers for whatever the pointer is on instead. The same box is what tells the
/// hook a pointer that has crossed to another item has moved at all, which is not a
/// distance to measure (see `pointer_item_holds`).
///
/// A box that cannot be told — the view reported no item, the hover is the
/// keyboard's, nothing at all is on screen — is no constraint, and the reveal is
/// governed by the count alone.
pub(super) static HOVER_POINTER_BOX: Mutex<Option<(i32, i32, i32, i32)>> = Mutex::new(None);

/// Note the item the pointer is on, as the box the view draws it in. Published by the
/// Explorer hook wherever it reads the item under the pointer, which is what keeps it
/// the item the pointer is on *now* rather than the one the last preview was for (see
/// `HOVER_POINTER_BOX`).
pub fn publish_pointer_item_box(bounds: (i32, i32, i32, i32)) {
    if let Ok(mut published) = HOVER_POINTER_BOX.lock() {
        *published = Some(bounds);
    }
}

/// Withdraw it: what is on screen is no longer a pointer's hover.
pub(super) fn clear_pointer_item_box() {
    if let Ok(mut published) = HOVER_POINTER_BOX.lock() {
        *published = None;
    }
}

/// Whether a box on screen holds a point. Half-open, so a point on the box's right or
/// bottom edge is outside it — which is how a window hit-test reads a rectangle, and
/// what the dismissal of a mouse preview is decided by.
pub(super) fn box_holds(x: i32, y: i32, region: (i32, i32, i32, i32)) -> bool {
    let (left, top, right, bottom) = region;

    x >= left && x < right && y >= top && y < bottom
}

/// Whether a point is still on the item the hover on screen is about — the one
/// question a preview is revealed against, and the one the hook reads a move off (see
/// `HOVER_POINTER_BOX`). A point outside a published box is a pointer that has left
/// the file, and no box at all is nothing to hold a reveal back with.
pub fn pointer_item_holds(x: i32, y: i32) -> bool {
    let Ok(published) = HOVER_POINTER_BOX.lock() else {
        return true;
    };

    (*published)
        .map(|region| box_holds(x, y, region))
        .unwrap_or(true)
}

/// The box the item under the pointer is drawn in, as the last look at it published, or
/// nothing when no look has answered with one (see `HOVER_POINTER_BOX`).
///
/// The hook asks for it where one look has to be told from another: the box a hover's
/// preview was resolved from is the item that preview is about, so a look that answered
/// nothing but left that same box under the pointer is a read that failed, while a look
/// that found another item publishes another box — an item with no preview of its own
/// being an item all the same (see `read_failure_is_the_same_item`).
pub fn pointer_item_box() -> Option<(i32, i32, i32, i32)> {
    HOVER_POINTER_BOX
        .lock()
        .ok()
        .and_then(|published| *published)
}

/// Whether the pointer is on that item this moment, read from the cursor: the
/// reveal's own question. A pointer that cannot be read is not a pointer that has
/// left, so an answer that could not be had holds nothing back either.
///
/// A pinned preview is not the pointer's and is not held to this question at all: the pin *is*
/// the item the pointer was on when it was taken up, so a pointer that has since gone to the
/// other end of the desktop — onto the pinned window itself, most likely — has left nothing that
/// is on screen (see `PIN_ACTIVE`).
pub(super) fn pointer_on_the_hovered_item() -> bool {
    if pinned() {
        return true;
    }

    cursor_position()
        .map(|cursor| pointer_item_holds(cursor.x, cursor.y))
        .unwrap_or(true)
}

/// Lines one wheel notch moves a text preview. Three is the step a text editor
/// takes, and it keeps a screenful to a few notches.
pub(super) const TEXT_SCROLL_LINES_PER_NOTCH: i64 = 3;
pub(super) const WHEEL_DELTA: i32 = 120;

/// How far behind the point a preview was opened from the region reaches, in
/// logical pixels. The pointer travels forwards from there, so this is only there
/// to keep the pixel under a hand at rest inside the region.
pub(super) const TEXT_SCROLL_ANCHOR_SLACK_PIXELS: f32 = 1.0;
/// How far above and below the row the pointer is on the journey to a preview may
/// wander before it is out of it. The journey is made across a row of the list, so
/// the band is the hand's, not the preview's.
pub(super) const TEXT_SCROLL_CORRIDOR_SLACK_PIXELS: f32 = 12.0;

/// How far either side of the scrollbar's column a press still counts as a press
/// on the bar, in logical pixels. It is deliberately small: the bar is thin, so a
/// hand aiming at it needs some slack, but everything further left is text, and a
/// press in the text is the start of a selection rather than a scroll.
pub(super) const TEXT_SCROLL_BAR_PRESS_SLACK_PIXELS: f32 = 8.0;

/// A distance written in logical pixels — the pixels of a display at 100% — in the
/// pixels of the display it is drawn on.
///
/// Every margin this app's placement is written around is a distance under a hand
/// rather than a count of pixels, so all of them are scaled this way: a standoff that
/// is comfortable at 100% is a sliver at 200%, and a room worth taking at 100% is a
/// room a preview is squeezed into at 200%. The text, the scrollbar and the grace
/// below are scaled for the same reason.
pub(super) fn logical_px(dpi: u32, logical_pixels: f32) -> i32 {
    (logical_pixels * dpi as f32 / 96.0).round() as i32
}

/// The grace `text_scroll_far_edge_grace_pixels` asks for past the far edge of a
/// text preview, at the display the preview is on.
///
/// That edge is the one the pointer arrives at last, and the one it can overshoot:
/// crossing the gap to reach the preview is a movement towards it, so the hand is
/// still moving when it gets there, and what usually waits at the end of the
/// journey is the scrollbar — a thin target sitting at the very edge of the frame.
/// A few pixels past it would otherwise take the preview down with it, which is
/// what this is for. It goes on the far side whichever side that is: the preview to
/// the right of the cursor is the common case, and then it is the right edge.
pub(super) fn far_edge_grace(dpi: u32, configured_pixels: f32) -> i32 {
    logical_px(dpi, configured_pixels)
}

/// The grace as `config.ini` has it, so a hand-edited distance is used as written
/// and a config that names none keeps the default.
pub(super) fn configured_far_edge_grace_pixels() -> f32 {
    CONFIG
        .lock()
        .map(|config| config.text_scroll_far_edge_grace_pixels)
        .unwrap_or(DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS)
}

/// A region on screen: left, top, right, bottom.
pub(super) type ScreenRegion = (i32, i32, i32, i32);

/// The region a preview of an item is kept off, as the `Avoid` setting measured it off
/// the item, and the one thing about it a placement cannot see: whether it is a *column*
/// of the item's view rather than the text the item itself draws.
///
/// An item's own text is a box, and a preview that comes out over one is stepped off it in
/// whichever direction asks the least — past its right edge, past its left, under it or
/// over it. A column is not one item's text: the view draws every row's in it — the `Name`
/// column of a `Details` or `Content` row at `Avoid Filename Column`, the columns a row
/// writes beside its name at `Avoid Details` — so a preview that comes out over a column
/// covers the rows beside the item however it is placed vertically, and the ways out of it
/// are the two to its sides. The ways off the item's own text are taken only where the
/// display leaves no room in either, since a preview over the item it describes is worse
/// than one over its neighbours (see `avoiding_text`).
#[derive(Clone, Copy)]
pub(crate) struct AvoidRegion {
    /// Where the region is on screen: left, top, right, bottom.
    pub(super) region: ScreenRegion,
    /// Whether it is a column of the view rather than the item's own text.
    pub(super) column: bool,
}

impl AvoidRegion {
    /// The text the item itself draws, which a preview is stepped off in either axis.
    pub(super) fn text(region: ScreenRegion) -> Self {
        Self {
            region,
            column: false,
        }
    }

    /// A column of the view — a region every row of it draws its text in — which a
    /// preview is stepped off to one of its two sides.
    pub(super) fn column(region: ScreenRegion) -> Self {
        Self {
            region,
            column: true,
        }
    }
}

/// The regions that keep a preview on screen, in screen coordinates, or `None`
/// when the preview on screen is not one that holds the pointer: the journey to a
/// text preview and the preview itself, or the box a waiting spinner occupies.
pub(super) static POINTER_HOLD_REGIONS: Lazy<Mutex<Option<Vec<ScreenRegion>>>> =
    Lazy::new(|| Mutex::new(None));

/// Where the preview that is on screen was opened from: the cursor that hovered
/// the file, or the middle of the focused item. The hold region is built from
/// this point and the preview's box, so the whole path between them is inside it.
pub(super) static TEXT_SCROLL_ANCHOR: Lazy<Mutex<Option<(i32, i32)>>> =
    Lazy::new(|| Mutex::new(None));

/// Whether the preview on screen is a text preview with more lines than it can
/// show, which is the condition for everything above.
pub(super) static TEXT_PREVIEW_SCROLLABLE: AtomicBool = AtomicBool::new(false);

/// Whether a pointer is using the preview on screen: any text preview in full
/// mode, which the pointer can rest on to select from. Published beside the region
/// so the check for it stays an atomic read.
pub(super) static TEXT_PREVIEW_HOLDING: AtomicBool = AtomicBool::new(false);

/// Whether what is on screen is a wait rather than a preview: the spinner a
/// document's page is being rendered behind.
///
/// The pointer is held by this the way it is held by a text preview, and for a
/// reason of its own: there is nothing under the spinner to hand the pointer back
/// to, and a hover that is dismissed while its page is on the way loses the page
/// it was waiting for — the render finishes, but the hover it was for is gone.
///
/// What it holds the pointer *through* is the item the wait is for and not the box
/// the spinner occupies: the spinner is placed a pixel off the pointer and follows
/// it, so a box of its own is one the pointer can never leave — which would be a
/// preview no fast hand could close on its way to somewhere else (see
/// `preview_pointer_hold`).
pub(super) static WAITING_PREVIEW_HOLDING: AtomicBool = AtomicBool::new(false);

/// Whether a drag that began on a page the engine is drawing is still down, which is a
/// question the rectangle on screen cannot answer for itself: the page is under the
/// pointer only where the drag started, and a drag that orbits the view or pans it is
/// read long after the pointer has walked off the box the page was drawn in.
///
/// The hook keeps it — a press inside the engine's rectangle arms it, and a tick with no
/// button down at all stands it down, which is the only reading that says the drag is
/// over. This side only reads it, and through a predicate of its own rather than through
/// the published regions, because a region published here is a region the wheel hook
/// reads and the wheel belongs to the page (see `note_engine_page_drag`).
pub(super) static ENGINE_PAGE_DRAG: AtomicBool = AtomicBool::new(false);

/// The regions that keep a preview alive: the journey to it, and the preview.
///
/// A preview is placed beside what it belongs to rather than over it, so the
/// pointer has to travel to reach it — across a gap, sometimes against the side
/// the placement chose. Joining the two means that journey never leaves the
/// regions, however the preview ended up placed relative to the cursor. The
/// journey is the first of the two regions and the preview is the second.
///
/// The journey reaches the preview's *nearest* point and stops there. Reaching
/// across to its far corner instead — which is what a region built from the two
/// corners of both rectangles does — covers every file in the list beside the
/// preview, from the row the pointer is on down to the preview's bottom: a text
/// preview is as tall as the display allows, and inside the region the item under
/// the pointer is not resolved at all, so those files stop previewing for as long
/// as the preview is up.
pub(super) fn text_scroll_hold_regions(
    preview: ScreenRegion,
    anchor: (i32, i32),
    far_edge_grace: i32,
    dpi: u32,
) -> [ScreenRegion; 2] {
    let (left, top, right, bottom) = preview;

    // The preview, with the margin a hand that overshot the edge it was travelling
    // towards needs — that edge is the far one, and what usually waits just past it
    // is the scrollbar.
    let preview_region = if anchor.0 < left {
        (left, top, right + far_edge_grace, bottom)
    } else if anchor.0 >= right {
        (left - far_edge_grace, top, right, bottom)
    } else {
        (left, top, right, bottom)
    };

    // The journey: from the point the preview was opened from to the nearest point
    // of the preview, and no further. It is as tall as the journey is and not as
    // tall as the preview is, so a pointer beside a preview crosses a row of the
    // list and nothing else.
    let near_x = anchor.0.clamp(left, right);
    let near_y = anchor.1.clamp(top, bottom);

    let anchor_slack = logical_px(dpi, TEXT_SCROLL_ANCHOR_SLACK_PIXELS);
    let corridor_slack = logical_px(dpi, TEXT_SCROLL_CORRIDOR_SLACK_PIXELS);

    let corridor = (
        anchor.0.min(near_x) - anchor_slack,
        anchor.1.min(near_y) - corridor_slack,
        anchor.0.max(near_x) + anchor_slack,
        anchor.1.max(near_y) + corridor_slack,
    );

    [corridor, preview_region]
}

/// Publish — or withdraw — the regions in which the pointer keeps what is on
/// screen alive, and note which of the two things that hold a pointer is on screen.
///
/// The Explorer hook polls this to decide whether the pointer over the preview
/// means "the user is reading this" or "dismiss it and show what is underneath",
/// and the wheel hook asks the same question before it decides whether the wheel
/// belongs to Explorer or to the preview. One thing holds the pointer through a
/// region: a text preview in full mode, which the pointer can rest on to select from
/// and scroll. The other — the spinner a page is being rendered behind, which has
/// nothing under it to hand the pointer back to and no page yet to be shown in its
/// place — is noted here rather than drawn, because what it holds the pointer through
/// is the item the wait is for (see `preview_pointer_hold`).
pub(super) unsafe fn publish_pointer_hold(hwnd: HWND) {
    // The pointer can only be on a window that is on screen, and while this one is not —
    // a document is drawn by the engine, in a window of its own — nothing of this app's is
    // under the pointer and there is nothing for a hold to keep alive. Asking the window
    // rather than the media is what keeps the hold from outliving the spinner it was
    // published for: a hold left standing over a document is a preview the pointer can
    // never close, because the pointer arriving at it is exactly what the hold refuses.
    if !IsWindowVisible(hwnd).as_bool() {
        clear_pointer_hold();
        return;
    }

    let (text, waiting) = CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| {
            let media = media.as_ref()?;
            Some((
                media
                    .text_state
                    .as_ref()
                    .map(|state| (state.dpi, state.can_scroll())),
                media.media_type.is_loading(),
            ))
        })
        .unwrap_or((None, false));

    // The wheel only belongs to a preview that can move under it, but the pointer
    // is held by any text preview in full mode: selecting and copying needs a
    // pointer that can rest on the preview whether or not it scrolls.
    TEXT_PREVIEW_SCROLLABLE.store(
        text.map(|(_, can_scroll)| can_scroll).unwrap_or(false),
        Ordering::Release,
    );
    TEXT_PREVIEW_HOLDING.store(text.is_some(), Ordering::Release);
    WAITING_PREVIEW_HOLDING.store(waiting, Ordering::Release);

    let mut keep_alive = if let Some((dpi, _)) = text {
        let mut rect = RECT::default();
        GetWindowRect(hwnd, &mut rect)
            .ok()
            .map(|_| (rect.left, rect.top, rect.right, rect.bottom))
            .map(|preview| {
                let anchor = TEXT_SCROLL_ANCHOR
                    .lock()
                    .ok()
                    .and_then(|anchor| *anchor)
                    // Without an anchor — a preview that was already on screen when this
                    // started, say — the preview's own corner stands in for it, which
                    // leaves the region as the preview alone.
                    .unwrap_or((preview.0, preview.1));

                text_scroll_hold_regions(
                    preview,
                    anchor,
                    far_edge_grace(dpi, configured_far_edge_grace_pixels()),
                    dpi,
                )
                .to_vec()
            })
    } else {
        // A wait publishes no region of its own. It holds the pointer through the item
        // it is waiting for, which is a question the hook asks of its own item box rather
        // than a rectangle this side hands it (see `preview_pointer_hold`): the spinner
        // is placed at the hand and follows it, so a box of its own would be one the
        // pointer could never leave.
        None
    };

    // A pinned window holds the pointer through the whole of itself. It is a window the user put
    // there and is reading, so a region built for a hover — the journey to it and the preview —
    // is the wrong answer for a place that was chosen by a drag; and what the region is *for*
    // here is the wheel, which belongs to a scrollable text preview inside a pin exactly as it
    // does to one on a hover. Nothing pinned is asked at all: this runs on every tick of a hover
    // and on every repaint of one (see `PIN_ACTIVE`).
    if pinned() {
        if let Some((window, _, _)) = pinned_window_box() {
            keep_alive = Some(vec![(window.0, window.1, window.2, window.3)]);
        }
    }

    if let Ok(mut published) = POINTER_HOLD_REGIONS.lock() {
        *published = keep_alive;
    }
}

/// Remember the point a preview was opened from, which is what the hold region
/// stretches back to.
pub(super) fn set_text_scroll_anchor(x: i32, y: i32) {
    if let Ok(mut anchor) = TEXT_SCROLL_ANCHOR.lock() {
        *anchor = Some((x, y));
    }
}

pub(super) fn clear_pointer_hold() {
    TEXT_PREVIEW_SCROLLABLE.store(false, Ordering::Release);
    TEXT_PREVIEW_HOLDING.store(false, Ordering::Release);
    WAITING_PREVIEW_HOLDING.store(false, Ordering::Release);
    // A drag is held through the page it began on, and a preview that is being taken down
    // is not one: whatever the hand was doing to it ends with it, and the latch is cleared
    // here rather than by the tick that finds the buttons up so that every caller's own
    // withdrawal is the whole of what a stale latch needs.
    ENGINE_PAGE_DRAG.store(false, Ordering::Release);
    if let Ok(mut published) = POINTER_HOLD_REGIONS.lock() {
        *published = None;
    }
    if let Ok(mut anchor) = TEXT_SCROLL_ANCHOR.lock() {
        *anchor = None;
    }
}

/// Whether the preview on screen is a text preview that scrolls. Cheap enough for
/// the Explorer hook to ask on every poll tick.
pub fn text_preview_scrollable() -> bool {
    TEXT_PREVIEW_SCROLLABLE.load(Ordering::Acquire)
}

/// Whether the pointer is inside a region that keeps what is on screen alive, without
/// waiting on a lock the preview thread may be holding, so the Explorer hook can ask on
/// every poll tick.
///
/// A text preview holds the pointer through the regions it published — the journey to it
/// and the preview itself — because a hand reading or selecting from one is on its way
/// there or already there.
///
/// A wait holds the pointer through the *file* it is waiting for rather than through the
/// box its spinner occupies. The spinner is placed at the hand and follows it, so a
/// region of its own would be a box the pointer could never leave — and a preview the
/// hook could never close, however fast the hand is moving on to something else. What a
/// wait is owed is the question the item box answers: a pointer still inside the item the
/// hover was resolved from is a pointer still waiting for that file, and one outside it
/// has gone wherever it liked, spinner or no spinner (see `HOVER_POINTER_BOX`).
///
/// An item box nobody could be read for is *not* a hold here, which is the reverse of the
/// rule a reveal follows, and for the reason the hold exists: a hold is the hook leaving the
/// mouse alone, so it is only ever taken on an answer. A wait whose item could not be
/// read is a wait the pointer may still dismiss — what it costs is a page read again from
/// the cache, and what the other reading costs is a preview nothing can close.
///
/// A page the engine is drawing holds the pointer through its own rectangle while the
/// document on screen is a page that runs, and that is the third thing beside the regions
/// above: a page the user has clicked into is a page the user is working, so the pointer
/// arriving on it is the arrival that hands the page its drag, not the arrival that takes
/// the preview down. It is held through the rectangle and through a drag that began in it —
/// an orbit carries the pointer outside the page it is orbiting — and only for a document
/// that runs, so an SVG or a font, which is drawn in the engine's window and cannot be
/// touched, is dismissed by the pointer as it always was. The hold is what keeps a hover
/// from being sticky here: without it the preview would be gone the moment the pointer
/// arrived, and there would be nothing left to click.
pub fn preview_pointer_hold(x: i32, y: i32) -> bool {
    if TEXT_PREVIEW_HOLDING.load(Ordering::Acquire) {
        let Ok(published) = POINTER_HOLD_REGIONS.lock() else {
            return false;
        };

        return (*published)
            .as_ref()
            .map(|regions| {
                regions.iter().any(|(left, top, right, bottom)| {
                    x >= *left && x < *right && y >= *top && y < *bottom
                })
            })
            .unwrap_or(false);
    }

    // Gated on whether the engine has a window up at all, and it is gated first because this is
    // asked at the pointer's own rate: 66 times a second, from the hook's tick, for a window that
    // is usually a picture. Behind the gate are a `PathBuf` clone out of what the engine is
    // holding, a name test and two `GetWindowRect`s — all evaluated as arguments to the call
    // below, and all thrown away by it whenever no page is running. The gate is the engine's own
    // answer, which is one atomic load (see `webview_preview::is_showing`).
    if webview_preview::is_showing()
        && engine_page_holds(
            x,
            y,
            preview_screen_rect().unwrap_or((0, 0, 0, 0)),
            webview_preview::showing_path().is_some_and(|path| html_is_engine_drawn(&path)),
            ENGINE_PAGE_DRAG.load(Ordering::Acquire),
        )
    {
        return true;
    }

    WAITING_PREVIEW_HOLDING.load(Ordering::Acquire)
        && pointer_item_box().is_some_and(|item| box_holds(x, y, item))
}

/// Whether the engine's own rectangle holds a point, as the page rule reads it: a document
/// that runs is held through the box the engine drew it in, and through a drag that began
/// in that box, and nothing else.
///
/// The box is the engine's own and not the hold regions published beside the text preview
/// on purpose. A published region is a region the wheel hook reads to decide whether the
/// wheel is this app's, and the wheel of a page that runs is the page's own — a page being
/// scrolled here is the page moving, and reading it as a scroll of the text preview behind
/// it moves something the user is not looking at. So this is a predicate of its own and the
/// page is held by it without ever being published.
pub(super) fn engine_page_holds(
    x: i32,
    y: i32,
    rect: (i32, i32, i32, i32),
    runs: bool,
    dragging: bool,
) -> bool {
    if !runs {
        return false;
    }

    box_holds(x, y, rect) || dragging
}

/// Note whether a drag that began on a page the engine is drawing is still down, which the
/// hook reads once a tick from the buttons it has already read for itself.
///
/// The hook keeps this rather than the preview side working it out from the pointer alone,
/// because a drag cannot be seen from where the pointer has got to: the page is under the
/// hand only where the drag began, and everything after that is the page being orbited or
/// panned. A press inside the engine's rectangle arms it, a tick with no button down at all
/// stands it down, and a tick in between leaves it standing — which is the reading a drag
/// that has wandered off the page is given.
pub fn note_engine_page_drag(down: bool) {
    ENGINE_PAGE_DRAG.store(down, Ordering::Release);
}

/// The published preview region, without blocking. The wheel hook runs inside a
/// system-wide hook procedure, where waiting on a lock held by the preview thread
/// would stall every wheel message on the desktop — so a lock it cannot take
/// immediately means the wheel is not ours to take either.
///
/// This is the preview itself and not the journey to it: the wheel belongs to the
/// preview while the pointer is on (or just past) its edge, not while the pointer
/// is still crossing the row in front of it, where the wheel is Explorer's.
pub fn text_scroll_keep_alive_try() -> Option<(i32, i32, i32, i32)> {
    if !text_preview_scrollable() {
        return None;
    }

    POINTER_HOLD_REGIONS
        .try_lock()
        .ok()
        .and_then(|held| held.as_ref().and_then(|regions| regions.last().copied()))
}

/// Move the text preview on screen to `first_line` and repaint it.
///
/// The frame is re-rendered rather than slid: only the lines coming into view are
/// styled, the window it is drawn in does not move, and a document that is
/// scrolled through costs a window of highlighting per step rather than a
/// re-read.
pub(super) unsafe fn scroll_text_preview(hwnd: HWND, first_line: usize) {
    // Taken out of the media so the lock is not held while the lines are styled
    // and painted.
    let Some((path, options, dpi, width, height, current)) =
        CURRENT_MEDIA.lock().ok().and_then(|media| {
            media.as_ref().and_then(|media| {
                media.text_state.as_ref().map(|scroll| {
                    (
                        scroll.path.clone(),
                        scroll.options,
                        scroll.dpi,
                        scroll.width,
                        scroll.height,
                        scroll.first_line,
                    )
                })
            })
        })
    else {
        return;
    };

    if first_line == current {
        return;
    }

    // A selection is a range of the frame that is on screen, so it is dropped
    // rather than left pointing at lines that have moved.
    let Some(frame) =
        text_preview::render_scrolled(&path, first_line, width, height, dpi, options, None)
    else {
        return;
    };

    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        let Some(media) = media.as_mut() else {
            return;
        };

        // The hover may have moved on while this frame was rendered.
        if media.text_state.as_ref().map(|state| &state.path) != Some(&path) {
            return;
        }

        media.frames[0] = Arc::new(ImageFrame::new(frame.pixels, frame.width, frame.height, 0));

        if let Some(state) = media.text_state.as_mut() {
            state.first_line = frame.first_line;
            state.visible_lines = frame.visible_lines;
            state.scrollbar = frame.scrollbar;
            state.lines = frame.lines;
            state.selection = None;
        }
    }

    render_layered_preview(hwnd);
}

/// Repaint the text preview where it is, with whatever is selected in it now.
///
/// A selection is painted into the frame rather than drawn over it, so changing
/// one costs a re-render of the window that is on screen — a screenful of lines,
/// all of which are cached after the first pass.
pub(super) unsafe fn repaint_text_preview(hwnd: HWND) {
    let Some((path, options, dpi, width, height, first_line, selection)) =
        CURRENT_MEDIA.lock().ok().and_then(|media| {
            media.as_ref().and_then(|media| {
                media.text_state.as_ref().map(|state| {
                    (
                        state.path.clone(),
                        state.options,
                        state.dpi,
                        state.width,
                        state.height,
                        state.first_line,
                        state.selection,
                    )
                })
            })
        })
    else {
        return;
    };

    let Some(frame) =
        text_preview::render_scrolled(&path, first_line, width, height, dpi, options, selection)
    else {
        return;
    };

    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        let Some(media) = media.as_mut() else {
            return;
        };

        if media.text_state.as_ref().map(|state| &state.path) != Some(&path) {
            return;
        }

        media.frames[0] = Arc::new(ImageFrame::new(frame.pixels, frame.width, frame.height, 0));

        if let Some(state) = media.text_state.as_mut() {
            state.lines = frame.lines;
        }
    }

    render_layered_preview(hwnd);
}

/// Whether a press at `(x, y)` in window coordinates lands on the scrollbar, and
/// if so, where the drag starts.
pub(super) fn text_scroll_drag_target(x: i32, y: i32) -> Option<usize> {
    CURRENT_MEDIA.lock().ok().and_then(|media| {
        let media = media.as_ref()?;
        let scroll = media.text_state.as_ref()?;
        let scrollbar = scroll.scrollbar?;

        // The whole column counts, not just the groove: the bar is thin, and a
        // press a few pixels to its left is a press on the bar as far as the user
        // is concerned.
        let slack = (TEXT_SCROLL_BAR_PRESS_SLACK_PIXELS * scroll.dpi as f32 / 96.0).round() as i32;
        let (left, top, right, bottom) = scrollbar.track;
        if x < left - slack || x >= right + slack || y < top || y >= bottom {
            return None;
        }

        Some(text_preview::scroll_line_at_track_y(
            scrollbar.track,
            scrollbar.thumb,
            y,
            scroll.visible_lines,
            scroll.scrollable_lines,
        ))
    })
}

/// The document line a drag to `y` in window coordinates asks for.
pub(super) fn drag_target_for_y(y: i32) -> Option<usize> {
    CURRENT_MEDIA.lock().ok().and_then(|media| {
        let scroll = media.as_ref()?.text_state.as_ref()?;
        let scrollbar = scroll.scrollbar?;
        Some(text_preview::scroll_line_at_track_y(
            scrollbar.track,
            scrollbar.thumb,
            y,
            scroll.visible_lines,
            scroll.scrollable_lines,
        ))
    })
}

pub(super) fn set_text_scroll_dragging(dragging: bool) -> bool {
    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return false;
    };

    media
        .as_mut()
        .and_then(|media| media.text_state.as_mut())
        .map(|scroll| {
            let was = scroll.dragging;
            scroll.dragging = dragging;
            was
        })
        .unwrap_or(false)
}

pub(super) fn is_text_scroll_dragging() -> bool {
    CURRENT_MEDIA
        .lock()
        .map(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| state.dragging)
                .unwrap_or(false)
        })
        .unwrap_or(false)
}

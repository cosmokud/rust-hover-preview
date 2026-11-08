//! What the pointer's own questions are answered from: the box of the item it is over, the
//! regions a press is held for, and the scrollbar a hand can drag.

use super::*;

use crate::app::seam::FONT_SIZE_CHOICES;

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
///
/// It is one packed word in the seam rather than the locked `Option` it was while both
/// sides were one process: the hook publishes it from its own tick, which may take no
/// lock, and a torn box names a place no item is drawn in (see
/// `crate::app::seam::POINTER_ITEM_BOX`).
pub fn publish_pointer_item_box(bounds: (i32, i32, i32, i32)) {
    crate::app::seam::POINTER_ITEM_BOX.publish(Some(bounds));
}

/// Withdraw it: what is on screen is no longer a pointer's hover.
pub(super) fn clear_pointer_item_box() {
    crate::app::seam::POINTER_ITEM_BOX.publish(None);
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
/// `publish_pointer_item_box`). A point outside a published box is a pointer that has left
/// the file, and no box at all is nothing to hold a reveal back with.
pub fn pointer_item_holds(x: i32, y: i32) -> bool {
    crate::app::seam::POINTER_ITEM_BOX
        .get()
        .map(|region| box_holds(x, y, region))
        .unwrap_or(true)
}

/// The box the item under the pointer is drawn in, as the last look at it published, or
/// nothing when no look has answered with one (see `publish_pointer_item_box`).
///
/// The hook asks for it where one look has to be told from another: the box a hover's
/// preview was resolved from is the item that preview is about, so a look that answered
/// nothing but left that same box under the pointer is a read that failed, while a look
/// that found another item publishes another box — an item with no preview of its own
/// being an item all the same (see `read_failure_is_the_same_item`).
pub fn pointer_item_box() -> Option<(i32, i32, i32, i32)> {
    crate::app::seam::POINTER_ITEM_BOX.get()
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
pub(crate) type ScreenRegion = (i32, i32, i32, i32);

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
    pub(crate) region: ScreenRegion,
    /// Whether it is a column of the view rather than the item's own text.
    pub(crate) column: bool,
}

impl AvoidRegion {
    /// The text the item itself draws, which a preview is stepped off in either axis.
    pub(crate) fn text(region: ScreenRegion) -> Self {
        Self {
            region,
            column: false,
        }
    }

    /// A column of the view — a region every row of it draws its text in — which a
    /// preview is stepped off to one of its two sides.
    pub(crate) fn column(region: ScreenRegion) -> Self {
        Self {
            region,
            column: true,
        }
    }
}

/// The regions that keep a preview on screen, in screen coordinates, or none at all
/// when the preview on screen is not one that holds the pointer: the journey to a
/// text preview and the preview itself, or the box a waiting spinner occupies.
///
/// **Two packed words in the seam rather than a `Vec` behind a lock.** The reader is the
/// Explorer hook, inside a tick that may take no lock, and the writer is this loop on the
/// other side of a process boundary, where a `Vec` is a length and a heap the other side
/// does not share. No caller ever publishes more than two boxes, and each is packed into
/// its slot whole, so a box read half from one publish and half from the next is not a
/// thing that can happen (see `crate::app::seam::HOLD_REGION_JOURNEY`).
pub(super) static POINTER_HOLD_REGIONS: HoldRegions = HoldRegions;

/// The handle over the two published boxes a preview holds the pointer through (see
/// `POINTER_HOLD_REGIONS`).
pub(super) struct HoldRegions;

impl HoldRegions {
    /// Publish these boxes in order, and withdraw every slot past them.
    pub(super) fn publish(&self, regions: &[ScreenRegion]) {
        crate::app::seam::HOLD_REGION_JOURNEY.publish(regions.first().copied());
        crate::app::seam::HOLD_REGION_PREVIEW.publish(regions.get(1).copied());
    }

    /// Whether one of them holds a point. Half-open, as a window hit-test reads a box.
    fn holds(&self, x: i32, y: i32) -> bool {
        [
            crate::app::seam::HOLD_REGION_JOURNEY.get(),
            crate::app::seam::HOLD_REGION_PREVIEW.get(),
        ]
        .into_iter()
        .flatten()
        .any(|region| box_holds(x, y, region))
    }

    /// The boxes as last published, in their slots, so a caller that is only standing in
    /// for this loop can put them back.
    #[cfg(test)]
    pub(super) fn get(&self) -> [Option<ScreenRegion>; 2] {
        [
            crate::app::seam::HOLD_REGION_JOURNEY.get(),
            crate::app::seam::HOLD_REGION_PREVIEW.get(),
        ]
    }

    /// Put back the slots `get` handed out, for a test that stood in for this loop.
    #[cfg(test)]
    pub(super) fn restore(&self, regions: [Option<ScreenRegion>; 2]) {
        crate::app::seam::HOLD_REGION_JOURNEY.publish(regions[0]);
        crate::app::seam::HOLD_REGION_PREVIEW.publish(regions[1]);
    }
}

/// Where the preview that is on screen was opened from: the cursor that hovered
/// the file, or the middle of the focused item. The hold region is built from
/// this point and the preview's box, so the whole path between them is inside it.
pub(super) static TEXT_SCROLL_ANCHOR: Lazy<Mutex<Option<(i32, i32)>>> =
    Lazy::new(|| Mutex::new(None));

/// Whether the preview on screen is a text preview with more lines than it can
/// show, which is the condition for everything above.
pub(super) static TEXT_PREVIEW_SCROLLABLE: AtomicBool = AtomicBool::new(false);

// `TEXT_PREVIEW_HOLDING`, `WAITING_PREVIEW_HOLDING` and `ENGINE_PAGE_DRAG` are the seam's
// own facts now, re-exported by this module's parent: the hook leaves the pointer alone on
// them and the loop publishes them, and while both sides were one process they were three
// statics here (see `crate::app::seam`).

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
    //
    // **And the window that is asked is the one the media is drawn in, which a pin makes a
    // different window from the handle here.** A pin that stands in a window of its own has put
    // the preview's window down (see `take_up_pinned_window`), and the handle the loop's
    // heartbeat carries is that preview's: it answers "nothing of ours is on the screen" for the
    // whole life of the pin, so every heartbeat withdrew the hold — the regions, the scroll's own
    // answer and the size's — and a notch over a pinned page belonged to Explorer while the
    // scrollbar, which is a mouse message to the pin's own window, still dragged. What is on
    // screen is the window the pin stands in, which is the same question
    // `hover_facts::cursor_preview_hover` already had to learn to ask (see `pinned_window`).
    let drawn_in = if pinned() { pinned_window() } else { hwnd };
    if !IsWindowVisible(drawn_in).as_bool() {
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

    let box_at_hand = keep_alive
        .as_ref()
        .and_then(|regions| regions.last().copied());
    POINTER_HOLD_REGIONS.publish(keep_alive.as_deref().unwrap_or_default());

    publish_text_wheel_boxes(
        box_at_hand,
        TEXT_PREVIEW_SCROLLABLE.load(Ordering::Acquire),
        TEXT_PREVIEW_HOLDING.load(Ordering::Acquire),
        pinned(),
    );
}

/// Publish — or withdraw — the two boxes a wheel notch over a text preview is claimed by: the scroll's,
/// which is the preview itself and not the journey to it, and the size's, which is the same box asked
/// about a page that fits the one it is drawn in.
///
/// The answer is a rule rather than a reading, and it is handed in as four values for that reason: a box
/// that can only be tested by putting a window on somebody's screen is a box that goes untested (see
/// `text_size_after_notches`). What the loop has in hand when it publishes is what this is given.
///
/// **The two are not one box.** A page that scrolls hands the wheel with its scroll, and a page short
/// enough to fit its box is a page no scroll will ever reach — but it is exactly one a reader has a
/// reason to enlarge, and a `Ctrl` and a wheel over it is the only grip there is on how large a pinned
/// page is drawn. That is a pin's question alone, because a hover that fits its box is a preview the
/// pointer has not settled on (see `TEXT_SCROLL_REGION` and `TEXT_SIZE_REGION`).
pub(super) fn publish_text_wheel_boxes(
    region: Option<ScreenRegion>,
    scrollable: bool,
    holding: bool,
    pinned: bool,
) {
    crate::app::seam::TEXT_SCROLL_REGION.publish(if scrollable { region } else { None });
    crate::app::seam::TEXT_SIZE_REGION.publish(if holding && pinned { region } else { None });
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
    POINTER_HOLD_REGIONS.publish(&[]);
    // And the two boxes the wheel hook reads are withdrawn with it, so a wheel turned over where a
    // preview was is the listing's and not a preview's that is no longer there.
    crate::app::seam::TEXT_SCROLL_REGION.publish(None);
    crate::app::seam::TEXT_SIZE_REGION.publish(None);
    if let Ok(mut anchor) = TEXT_SCROLL_ANCHOR.lock() {
        *anchor = None;
    }
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
/// has gone wherever it liked, spinner or no spinner (see `POINTER_ITEM_BOX`).
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
        return POINTER_HOLD_REGIONS.holds(x, y);
    }

    // The page rule, read off the two facts the hook cannot ask for itself: whether the document the
    // engine is drawing is one that runs, and where the engine's window stands. Both are the preview
    // side's answers, published once a tick (see `crate::app::seam::ENGINE_PAGE_RUNS`). The gate on
    // the engine having a window up at all is the first of the two: a page that runs is a page the
    // engine is drawing, so a run of the rule with nothing showing cannot be reached.
    if ENGINE_PAGE_RUNS.load(Ordering::Acquire)
        && engine_page_holds(
            x,
            y,
            preview_screen_rect().unwrap_or((0, 0, 0, 0)),
            true,
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

/// Move the text preview on screen to `first_line` and repaint it.
///
/// The frame is re-rendered rather than slid: only the lines coming into view are
/// styled, the window it is drawn in does not move, and a document that is
/// scrolled through costs a window of highlighting per step rather than a
/// re-read.
///
/// **And a page moved by a line or two is the frame already on screen, shifted.** A held arrow's
/// press is mostly a page scrolled a line, and laying a screenful out and painting every run of it
/// again cost 4.5 ms of the 6.9 a press took — the frame that is on screen is the same frame, and the
/// rows that came in from an edge are the only ones not already painted (see `PaintedFrame`).
pub(super) unsafe fn scroll_text_preview(hwnd: HWND, first_line: usize) {
    // Taken out of the media so the lock is not held while the lines are styled
    // and painted — the frame on screen as well, as a handle rather than a copy: a screenful is a
    // megabyte and a half.
    let Some((path, options, dpi, width, height, current, painted)) =
        CURRENT_MEDIA.lock().ok().and_then(|media| {
            media.as_ref().and_then(|media| {
                let scroll = media.text_state.as_ref()?;
                let painted = media.current_frame().cloned().map(|frame| {
                    (
                        frame,
                        scroll.first_line,
                        scroll.lines.clone(),
                        scroll.selection,
                        scroll.find.highlights(),
                    )
                });
                Some((
                    scroll.path.clone(),
                    scroll.options,
                    scroll.dpi,
                    scroll.width,
                    scroll.height,
                    scroll.first_line,
                    painted,
                ))
            })
        })
    else {
        return;
    };

    if first_line == current {
        return;
    }

    // What is selected is kept exactly as it is: a selection is a place in the document, so the
    // frame that comes back draws the same text somewhere else rather than the selection having
    // to be moved to where the text went (see `Selection`). A find bar's places are taken out beside
    // it and drawn into the same frame, because a highlight is a place in the document too and one
    // painted over the frame afterwards would cover a selection the reader made (see
    // `FindHighlights`).
    let Some((selection, find)) = CURRENT_MEDIA.lock().ok().and_then(|media| {
        media
            .as_ref()
            .and_then(|media| media.text_state.as_ref())
            .map(|state| (state.selection, state.find.highlights()))
    }) else {
        return;
    };

    // An empty selection is no selection — `paint` reads it that way — so it is said that way here
    // too, and a page whose caret was put down without a drag is a page with nothing painted into it.
    let selection = selection.filter(|selection| !selection.is_empty());

    // What was painted into the frame on screen besides its text, which is only nothing: a page with
    // a selection or a search painted into it is a page whose pixels are not the ones the layout would
    // paint elsewhere, and it is drawn again rather than shifted (see `PaintedFrame`).
    let painted = match painted.as_ref() {
        Some((frame, first_line, lines, selection, find))
            if selection.is_none_or(|selection| selection.is_empty()) && find.is_none() =>
        {
            Some(text_preview::PaintedFrame {
                pixels: &frame.pixels,
                width: frame.width,
                height: frame.height,
                first_line: *first_line,
                lines,
                selection: None,
                find: None,
            })
        }
        _ => None,
    };

    // Held: the page being scrolled is drawn from the document that drew it, so a line moved by a
    // caret or a notch is not a look at the file (see `render_scrolled_held`).
    let Some(frame) = text_preview::render_scrolled_held(
        painted, &path, first_line, width, height, dpi, options, selection, find,
    ) else {
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
            state.selection = selection;
        }
    }

    render_layered_preview(hwnd);
}

/// Lay a pinned page of text out again in the box its window has just been given — one a hand has
/// dragged it to, a maximize, a restore, a size reset, or a display that changed — keeping the line
/// it is scrolled to, what is selected in it and the size the reader put it to.
///
/// It is the whole of what a box asks of a page, and nothing lays the page out again behind it: a read
/// begins at the document's first line, because nothing carries the reader's line into a frame that
/// is not the one being replaced, so the load the same box change starts for every other kind is that
/// page drawn a second time from the top (see `relayout_pinned_media`).
///
/// A box decides what a page of text shows rather than how large it is: the font is a fixed,
/// display-scaled size, and the box decides how many lines it holds and how long each of them may
/// be before it wraps (see `render_scrolled`). So a window under a hand is a page with more lines in
/// it and longer ones — and not the page it was drawn as, stretched — which is what a resize does to
/// every other kind and the one thing it may not do to this one (see `compose_media_into_band`).
///
/// It is not a read and not a parse: the document is already read and the lines on screen already
/// styled (see `DOCUMENTS` and `styled_window`), so what this costs is a screenful laid out and
/// painted, which is the cost a wheel's own scroll pays (see `scroll_text_preview`).
///
/// The band the frame lands in is the box it is drawn for, so the paint that follows composes it row
/// for row rather than sampling it.
pub(super) fn relayout_text_for_the_box(content: ScreenRegion) {
    let width = (content.2 - content.0).max(1) as u32;
    let height = (content.3 - content.1).max(1) as u32;

    // Read out of the media, so the lock is not held while the page is laid out and painted — and
    // answered with nothing at all where the pin is showing a page of another kind, or no page at
    // all, which is what keeps a resize of every other kind the paint it has always been.
    let Some((path, options, dpi, drawn, first_line, selection, find)) =
        CURRENT_MEDIA.lock().ok().and_then(|media| {
            media
                .as_ref()
                .filter(|media| media.media_type == MediaType::Text)
                .and_then(|media| media.text_state.as_ref())
                .map(|state| {
                    (
                        state.path.clone(),
                        state.options,
                        state.dpi,
                        (state.width, state.height),
                        state.first_line,
                        state.selection,
                        state.find.highlights(),
                    )
                })
        })
    else {
        return;
    };

    // A box the page is already drawn for owes nothing: a drag that has met the room's edge brings
    // the same size again, and a screenful laid out for it is the work below paid for nothing.
    if drawn == (width, height) {
        return;
    }

    // A find bar that is up stays up across the box: its matches are places in the document, so the
    // page laid out again draws the same places on the same lines, wherever the new width wraps them
    // (see `FindState`).
    let Some(frame) = text_preview::render_scrolled(
        &path, first_line, width, height, dpi, options, selection, find,
    ) else {
        return;
    };

    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        let Some(media) = media.as_mut() else {
            return;
        };

        // The pin may have been taken up on another file while this page was being laid out.
        if media.text_state.as_ref().map(|state| &state.path) != Some(&path) {
            return;
        }

        media.frames[0] = Arc::new(ImageFrame::new(frame.pixels, frame.width, frame.height, 0));

        if let Some(state) = media.text_state.as_mut() {
            state.width = frame.width;
            state.height = frame.height;
            state.first_line = frame.first_line;
            state.visible_lines = frame.visible_lines;
            state.scrollable_lines = frame.scrollable_lines;
            state.scrollbar = frame.scrollbar;
            state.lines = frame.lines;
        }
    }
}

/// Lines a tick of a selection drag carried past the edge of the box moves the page: a screenful
/// every two seconds at the tick the preview loop runs at, which is a speed a hand can stop.
const TEXT_DRAG_LINES_PER_TICK: i64 = 3;

/// Where a selection drag that has left the box asks the page to go, and the point the selection is
/// grown to afterwards — or nothing while the pointer is inside the box or no drag is running.
///
/// The point comes back past the edge the pointer left rather than where the pointer is: the frame
/// that comes back is the page a few lines on (or back), and the nearest line to a point just
/// outside it is the line at its edge. That is what makes the selection grow with the scroll —
/// growing it to the pointer's own line first would leave the far end standing still while the text
/// moved under it (see `position_in` and `step_text_drag_selection`).
pub(super) fn text_drag_step() -> Option<(usize, (i32, i32))> {
    let Ok(media) = CURRENT_MEDIA.lock() else {
        return None;
    };
    let state = media.as_ref().and_then(|media| media.text_state.as_ref())?;
    if !state.selecting {
        return None;
    }

    let (x, y) = state.drag_pointer?;
    let (step, edge) = if y < 0 {
        (-TEXT_DRAG_LINES_PER_TICK, -1)
    } else if y >= state.height as i32 {
        (TEXT_DRAG_LINES_PER_TICK, state.height as i32 + 1)
    } else {
        return None;
    };

    Some((state.scrolled_by(step), (x, edge)))
}

/// Follow a selection drag that has been carried past the edge of the box: the page is taken a step
/// in the direction the drag left by, and the selection is grown to the line at the edge it came
/// back with, so a range longer than the box shows can be made with the pointer alone.
///
/// Nothing is drawn where the page cannot move any further: the scroll itself is a no-op at either
/// end of the document, and a selection grown to the edge line of the frame that is already on
/// screen has nowhere left to grow to.
pub(super) unsafe fn step_text_drag_selection(hwnd: HWND) {
    let Some((line, point)) = text_drag_step() else {
        return;
    };

    scroll_text_preview(hwnd, line);

    if extend_text_selection(point.0, point.1) == Some(true) {
        repaint_text_preview(hwnd);
    }
}

/// The size a pinned page is drawn at after a wheel has been turned over it with `Ctrl` held: the
/// steps the tray's own `Text Size` offers, walked one per notch — so the sizes a reader can reach in
/// a pin are the sizes the tray lists, and a size hand-edited into `config.ini` between two of them
/// steps to the next one in the direction the wheel was turned.
///
/// Nothing moves past either end of the list: the largest step is the largest page and the smallest
/// is the smallest, and a notch that has nowhere to go leaves the size it found rather than taking
/// the page to the other end of the scale.
pub(super) fn text_size_after_notches(current: u32, notches: i32) -> u32 {
    let steps = || FONT_SIZE_CHOICES.iter().map(|(percent, _)| *percent);
    let mut size = current;

    for _ in 0..notches.abs() {
        size = if notches > 0 {
            steps()
                .filter(|percent| *percent > size)
                .min()
                .unwrap_or(size)
        } else {
            steps()
                .filter(|percent| *percent < size)
                .max()
                .unwrap_or(size)
        };
    }

    size
}

/// Draw the pinned page again at a step of its own text size, keeping its box, its place in the
/// document and what is selected in it.
///
/// The size is the pin's own rather than the file's or the configuration's: it is held in the page's
/// state for as long as the pin lives, and it is the one thing a reader who has zoomed into a page
/// does not expect to have to do again on the next file of a walk (see `keep_text_place`).
///
/// One step is taken per tick that any notch arrived in rather than per whole notch: a high-
/// resolution wheel reports less than a notch per message, and a count truncated to whole notches is
/// most of its travel thrown away.
pub(super) unsafe fn step_pinned_text_size(hwnd: HWND, notches: i32) {
    let Some(size) = reader_text_size() else {
        return;
    };

    let percent = text_size_after_notches(size, notches);
    if percent != size {
        redraw_pinned_page(hwnd, Some(percent));
    }
}

/// Lay the page a pin is showing out again from the options the configuration now selects, keeping
/// the line it is read at, what is selected in it and the size the reader put it to.
///
/// A page's glyphs and its Markdown are painted into its frame rather than composed into the
/// window, so a theme or a `Markdown` switch is a page drawn again and not a repaint of the frame
/// already in hand (see `set_theme`). Every other setting a tick is owed a repaint for is a page
/// this says nothing about, and nothing is laid out where the options it already has are the
/// options asked for.
///
/// The size the reader put the page to is the page's own, and it is kept across all of this. What
/// overrides it is `font_scale`, which is the tray's own `Font Size` row and the one thing that puts
/// a page already up to another size: the reader's own goes with the pick the way it goes when the
/// pin is closed (see `keep_text_place` and `publish_text_font_size`).
pub(super) unsafe fn redraw_pinned_page(hwnd: HWND, font_scale: Option<u32>) {
    // A page is a pin's: a hover of the same file is rebuilt from its own load, at the size the
    // configuration names, rather than drawn again under a frame a hover is composed from.
    if !pinned() {
        return;
    }

    let Some((path, before, dpi, width, height, first_line, selection, find)) =
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
                        state.find.highlights(),
                    )
                })
            })
        })
    else {
        return;
    };

    let configured = current_text_options();
    let options = TextPreviewOptions {
        theme: configured.theme,
        markdown_mode: configured.markdown_mode,
        font_scale_percent: font_scale.unwrap_or(before.font_scale_percent),
        // Whether the page is more than something to look at is the pin's own answer rather than
        // a setting's (see `current_text_options`).
        full_mode: before.full_mode,
    };
    if options == before {
        return;
    }

    // Nothing is laid out while the lock is held: the render is a screenful of text, and the lock is
    // the one a paint takes (see `scroll_text_preview`). The find bar's places are drawn with it: a
    // page redrawn at another size or theme is the same document, and its search is not thrown away
    // by a repaint (see `FindHighlights`).
    let Some(frame) = text_preview::render_scrolled(
        &path, first_line, width, height, dpi, options, selection, find,
    ) else {
        return;
    };

    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        let Some(media) = media.as_mut() else {
            return;
        };

        // The pin may have moved on while this page was drawn.
        if media.text_state.as_ref().map(|state| &state.path) != Some(&path) {
            return;
        }

        media.frames[0] = Arc::new(ImageFrame::new(frame.pixels, frame.width, frame.height, 0));

        if let Some(state) = media.text_state.as_mut() {
            state.options = options;
            state.first_line = frame.first_line;
            state.visible_lines = frame.visible_lines;
            state.scrollbar = frame.scrollbar;
            state.lines = frame.lines;
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
    let Some((path, options, dpi, width, height, first_line, selection, find)) =
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
                        state.find.highlights(),
                    )
                })
            })
        })
    else {
        return;
    };

    // Held, for the reason `scroll_text_preview` is: a selection painted into the page is a page
    // drawn again from the text it already has, not a look at the file (see `render_scrolled_held`).
    let Some(frame) = text_preview::render_scrolled_held(
        None, &path, first_line, width, height, dpi, options, selection, find,
    ) else {
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

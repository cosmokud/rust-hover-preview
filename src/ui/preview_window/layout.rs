//! Where a preview is put: the room it has, the box a pointer's own preview takes beside it,
//! the box a keyboard preview takes next to what is focused, and the clearance that keeps a
//! window off the text it was covering.

use super::*;

/// Computed preview window layout
pub(super) struct PreviewLayout {
    pub(super) pos_x: i32,
    pub(super) pos_y: i32,
    pub(super) max_width: u32,
    pub(super) max_height: u32,
    pub(super) preview_w: u32,
    pub(super) preview_h: u32,
}

/// Compared and printed rather than only copied, because the placements this is asked for are
/// read back as rectangles: a test that asserts on one wants to say what the rectangle was, and
/// an inequality between two displays is how "one display rather than the union" is stated.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) struct ScreenBounds {
    pub(super) left: i32,
    pub(super) top: i32,
    pub(super) right: i32,
    pub(super) bottom: i32,
}

impl ScreenBounds {
    pub(super) fn height(self) -> i32 {
        self.bottom - self.top
    }

    /// The room the display has: the whole work area, which is the largest box anything
    /// shown on this display can be drawn in — every placement mode takes its own room
    /// out of this one, so a box this size bounds all of them.
    ///
    /// It is what an engine that has to draw a preview *before* the file can be measured
    /// is asked for, rather than the room the hover's own layout came out at: a picture
    /// the image converter develops is developed at the size it is then shown at, so a
    /// room smaller than this is a picture that can never be shown any larger however
    /// much room its preview is given afterwards (see `PendingLoad::room`).
    pub(super) fn room(self) -> (u32, u32) {
        (
            (self.right - self.left).max(1) as u32,
            (self.bottom - self.top).max(1) as u32,
        )
    }

    /// The same room as a box on screen, which is what the sizing functions take: they measure
    /// into a room and never read where it stands (see `scale_dimensions`).
    pub(super) fn region(self) -> ScreenRegion {
        (self.left, self.top, self.right, self.bottom)
    }
}

/// Where the pointer is, in screen coordinates.
pub(super) fn cursor_position() -> Option<POINT> {
    let mut point = POINT::default();
    unsafe { GetCursorPos(&mut point) }.ok()?;
    Some(point)
}

/// The top edge that centers a `height`-tall preview on `center`, kept inside the
/// display.
///
/// Centering is the intent and the screen edge is the limit: a preview that fits
/// under a cursor near the top is placed where the cursor is rather than pushed
/// down the display, and one tall enough to reach an edge is moved only as far as
/// that edge allows.
pub(super) fn centered_top(center: i32, height: i32, bounds: ScreenBounds) -> i32 {
    let lowest = (bounds.bottom - height).max(bounds.top);
    (center - height / 2).clamp(bounds.top, lowest)
}

/// The least room a way out of the text may leave before a preview is resized into
/// it. Below this the room is a sliver — the tail past a name that fills its row, the
/// last strip of a display under a row at the bottom — and a preview squeezed into it
/// says less than the one left over the name would have.
pub(super) const MIN_AVOID_ROOM_PIXELS: f32 = 64.0;

/// A layout moved off the region the item it describes draws, as the `Avoid` setting
/// measured it: the name the file is listed under at `Avoid Filename`, and that name
/// with the columns a row writes beside it at `Avoid Details`.
///
/// What a placement is kept clear of, and how far: the region the `Avoid` setting
/// measured off the item — and what kind of region it is, which decides which steps off
/// it are steps at all — the distance the placement keeps from it, and the pointer a
/// mouse hover's placement is held clear of as well (see `avoiding_text`).
pub(super) struct Clearance {
    /// The region the `Avoid` setting measured off the hovered item: the name the file
    /// is listed under, that name with the columns a row writes beside it, or `None`
    /// where the setting is off or the view reported no text for the item — with whether
    /// it is a column of the view rather than the item's own text, which is what the ways
    /// out are read against (see `AvoidRegion`).
    pub(super) text: Option<AvoidRegion>,
    /// The distance the placement keeps from `text`, so the text is stepped off rather
    /// than touched at its edge. Already in the pixels of the display the placement is
    /// for — like the least room a way out is worth taking, which `avoiding_text`
    /// scales from the logical distance it is written as.
    pub(super) gap: i32,
    /// Where the pointer is, for a placement that is a mouse hover's: a way out that
    /// would put the preview over the pointer is not one the hover can use. A mouse
    /// preview is dismissed the moment the cursor touches it, so one placed *under* the
    /// cursor is dismissed the instant it appears — and put back the same way a moment
    /// later, for as long as the pointer sits there, which is a preview that blinks at
    /// the hand rather than one that is read. A pointer clear of the preview is what
    /// makes the dismissal mean what it says: the pointer arriving at the preview is
    /// the user asking for it to go. So a way out that keeps the pointer clear is
    /// preferred over one that does not, whatever size the two offer — and a way out
    /// that covers it is still taken where no other is left, since a preview moved as
    /// far as the display allows is all there is to give. The keyboard's placements
    /// carry no pointer: what they clear is the focused item, and a preview over a
    /// parked pointer is allowed there (see `compute_keyboard_layout`).
    pub(super) cursor: Option<(i32, i32)>,
}

/// A preview is placed beside what it belongs to rather than over it, and the text of
/// the item it came from is part of what it belongs to: that item stays readable while
/// its preview is up, which is what the `Avoid` setting asks for. The placement the
/// position mode chose is therefore moved by the shortest step that clears that region —
/// past its right edge, past its left, under it or over it, whichever asks the least
/// of the preview — and only a step the display has room for is taken, so a preview
/// moved off one edge is never pushed off another. A step that would put the preview
/// over the pointer is not one a mouse hover can use either, whether or not it is the
/// shortest: the rules the ways out are held to are the ones `Clearance` names.
///
/// A *column* is the one region the ways over and under do not clear. The view draws
/// every row's text in it — the `Name` column at `Avoid Filename Column`, the columns
/// beside the name at `Avoid Details` — so a preview stepped off the item's own row is
/// still over the rows next to it, which is the whole of what the setting exists to
/// keep the preview off: in a `Details` or `Content` view, a preview placed above or
/// below the hovered item covers its neighbours' names exactly as one placed over it
/// would. The ways out of a column are therefore the two to its sides, and the four
/// ways are taken only where the display leaves no room in either — a preview over the
/// neighbours' names is then the lesser thing, because the item the preview describes
/// is the one it would cover instead (see `AvoidRegion`).
///
/// A preview too large for every one of those rooms is *resized* into the roomiest of
/// them rather than left where it covers the text. That is the case a preview filling
/// the display lands in: nothing can be moved into place beside a name while the
/// preview is as wide and as tall as the display, so it is the size that gives, and the
/// preview shows as much as the room beside the name can hold — which is the same rule
/// that sized it in the first place, applied to the room that is left. A room too small
/// to be worth having is not taken at all, so a preview is never squeezed into a sliver
/// to get off a name that a usable preview would have covered anyway.
pub(super) fn avoiding_text(
    layout: PreviewLayout,
    orig_dims: (u32, u32),
    preview_scale: PreviewScale,
    clearance: Clearance,
    bounds: ScreenBounds,
    dpi: u32,
) -> PreviewLayout {
    let Clearance { text, gap, cursor } = clearance;
    let Some(AvoidRegion { region, column }) = text else {
        return layout;
    };
    let (text_left, text_top, text_right, text_bottom) = region;

    let (left, top) = (layout.pos_x, layout.pos_y);
    let (width, height) = (layout.preview_w as i32, layout.preview_h as i32);
    let min_room = logical_px(dpi, MIN_AVOID_ROOM_PIXELS);

    // What the placement is kept clear of: the region itself for the text one item
    // draws, and the region's own span *down the display* for a column — every row of
    // the view draws its text in the column, so a preview that overlaps it across is
    // over the column wherever it sits, and the ways out of it are the two to its
    // sides, the ways over and under the row being only the ones left where the
    // display has no room in either (see below).
    let (block_top, block_bottom) = match column {
        true => (bounds.top, bounds.bottom),
        false => (text_top, text_bottom),
    };

    let covers_text = left < text_right
        && left + width > text_left
        && top < block_bottom
        && top + height > block_top;
    if !covers_text {
        return layout;
    }

    // Where a preview clear of the text would sit: past the text's right edge, before
    // its left one, under it and over it, each a gap away from it. The ways over and
    // under are read off the region's own box, which for a column is the row it was
    // measured at: they are the ways out a preview of a column takes only where the
    // display leaves no room beside it (see the note on `clear_of_column` below).
    let past_right = text_right + gap;
    let before_left = text_left - gap;
    let under = text_bottom + gap;
    let over = text_top - gap;

    // The four ways out of the text, each as the room the display has on that side,
    // whether that room is a width — a step to either side — or a height, where the
    // preview sits in it, and whether that anchor is the preview's own far edge. A
    // step right or down grows away from the text from its near edge and a step left
    // or up from its far one.
    let ways_out = [
        (bounds.right - past_right, true, past_right, false),
        (before_left - bounds.left, true, before_left, true),
        (bounds.bottom - under, false, under, false),
        (over - bounds.top, false, over, true),
    ];

    // The ways out are compared by whether they keep the pointer clear, then by whether
    // they keep the preview clear of a column, then by the preview's own size, then by
    // how short the step is — see the notes on `cursor` and on `clear_of_column`.
    let mut best: Option<(bool, bool, i64, i32, PreviewLayout)> = None;
    for (room, along_width, anchor, far_edge) in ways_out {
        let room = room.max(0);

        // The box the way out offers: the room where it constrains the preview, and
        // what the mode allowed it where it does not — a preview moved off the text is
        // never enlarged by the move.
        let (max_width, max_height) = if along_width {
            (room.min(layout.max_width as i32) as u32, layout.max_height)
        } else {
            (layout.max_width, room.min(layout.max_height as i32) as u32)
        };
        if max_width == 0 || max_height == 0 {
            continue;
        }

        let (preview_w, preview_h) = scale_dimensions(
            orig_dims.0,
            orig_dims.1,
            max_width,
            max_height,
            preview_scale,
        );
        if preview_w == 0 || preview_h == 0 {
            continue;
        }

        // A way out that costs the preview its size is only taken where the room left
        // is worth having; one that costs it nothing is taken however little room it
        // leaves, since the preview was already going to be that small.
        let natural_size = if along_width { width } else { height };
        let size_there = if along_width {
            preview_w as i32
        } else {
            preview_h as i32
        };
        if size_there < natural_size && room < min_room {
            continue;
        }

        let placement = PreviewLayout {
            pos_x: if along_width {
                if far_edge {
                    anchor - preview_w as i32
                } else {
                    anchor
                }
            } else {
                left
            },
            pos_y: if along_width {
                top
            } else if far_edge {
                anchor - preview_h as i32
            } else {
                anchor
            },
            max_width,
            max_height,
            preview_w,
            preview_h,
        };

        // Whether this way out leaves the preview clear of the region's own column: a
        // way clear of it leaves every row's text readable, where one over or under the
        // item covers the rows beside it — which is the whole of what a column is kept
        // off for. The two ways to the sides are clear of it by construction; the ways
        // over and under it are only as clear as the place the mode put the preview
        // already was, which is what leaves them as the ways out of last resort. For
        // the text of one item every way out answers `true`, so nothing is chosen by
        // this (see the note on `Clearance`).
        let clear_of_column = !column
            || placement.pos_x >= text_right
            || placement.pos_x + preview_w as i32 <= text_left;

        // A way out that keeps the pointer clear of the preview comes first, then one
        // that is clear of a column, then the largest preview, and the shortest move
        // breaks a tie: every way out that fits the preview as it stands offers it the
        // same size, so those are the ones the move decides between, and only a preview
        // that has to shrink is chosen between by what the room holds. See the note on
        // `cursor` above for why the pointer comes ahead of both.
        let clear_of_cursor = cursor.is_none_or(|(x, y)| {
            !box_holds(
                x,
                y,
                (
                    placement.pos_x,
                    placement.pos_y,
                    placement.pos_x + preview_w as i32,
                    placement.pos_y + preview_h as i32,
                ),
            )
        });
        let area = preview_w as i64 * preview_h as i64;
        let step = (placement.pos_x - left).abs() + (placement.pos_y - top).abs();
        let better = match &best {
            Some((best_clear, best_column, best_area, best_step, _)) => {
                if clear_of_cursor != *best_clear {
                    // One of the two keeps the pointer clear and the other does not, and
                    // that is the whole of the choice between them.
                    clear_of_cursor
                } else if clear_of_column != *best_column {
                    // And one of them is clear of a column the other covers, which for a
                    // column is the whole of what the move is for.
                    clear_of_column
                } else {
                    area > *best_area || (area == *best_area && step < *best_step)
                }
            }
            None => true,
        };
        if better {
            best = Some((clear_of_cursor, clear_of_column, area, step, placement));
        }
    }

    match best {
        Some((_, _, _, _, placement)) => placement,
        None => layout,
    }
}

/// How far off the pointer a preview is placed, in logical pixels: the margin the
/// position modes are written around, and the gap the waiting spinner is kept at as well.
///
/// The spinner is the preview window with an arc in it, so a pointer that lands on it is a
/// pointer that has stopped clicking and probing the file it is waiting on: what a wait is
/// owed is the hand's own corner and not the hand itself, and the same gap every other
/// preview keeps is a gap the arc can be touched across but not spawned in.
pub(super) const POINTER_STANDOFF_PIXELS: f32 = 20.0;

/// Compute preview layout for mouse hover (relative to cursor position)
///
/// `placement` is what the hover asks for — the size the preview was measured at,
/// the text to keep it off, the position mode it follows and the scale it is drawn
/// with — the same reading a pending load keeps to place its preview again as the
/// pointer moves (see `HoverPlacement`). `dpi` is the display the pointer is on,
/// which is what the margins the placement is written around are scaled by.
///
/// A placement that is the waiting spinner (`at_the_pointer_corner`) is placed by its own
/// rule: the corner nearest the pointer is put a pointer gap off it — the same gap every
/// other preview keeps, so the arc is beside the hand rather than under it — in whichever
/// of the four quadrants the display has room for the spinner, and the name it covers is
/// not stepped around: a spinner waiting on the page of the file under the hand says what
/// it is by being at the hand's own corner, and one placed a row away from it says nothing
/// about what is being waited on. Every other preview keeps the margin its position mode
/// leaves and the room the `Avoid` setting asks for.
pub(super) fn compute_mouse_layout(
    cursor_x: i32,
    cursor_y: i32,
    placement: HoverPlacement,
    bounds: ScreenBounds,
    dpi: u32,
) -> Option<PreviewLayout> {
    let HoverPlacement {
        orig_dims,
        avoid,
        follow_cursor,
        preview_scale,
        at_the_pointer_corner,
    } = placement;

    let offset = logical_px(dpi, POINTER_STANDOFF_PIXELS);
    let (orig_w, orig_h) = (orig_dims.0 as i32, orig_dims.1 as i32);

    if follow_cursor || at_the_pointer_corner {
        let quadrants = [
            (
                bounds.right - cursor_x - offset,
                bounds.bottom - cursor_y - offset,
                cursor_x + offset,
                cursor_y + offset,
            ), // BR
            (
                cursor_x - bounds.left - offset,
                bounds.bottom - cursor_y - offset,
                bounds.left,
                cursor_y + offset,
            ), // BL
            (
                bounds.right - cursor_x - offset,
                cursor_y - bounds.top - offset,
                cursor_x + offset,
                bounds.top,
            ), // TR
            (
                cursor_x - bounds.left - offset,
                cursor_y - bounds.top - offset,
                bounds.left,
                bounds.top,
            ), // TL
        ];

        let mut best_quadrant = 0;
        let mut best_scale: f32 = 0.0;

        for (i, &(avail_w, avail_h, _, _)) in quadrants.iter().enumerate() {
            if avail_w <= 0 || avail_h <= 0 {
                continue;
            }
            let scale = scale_in_room(
                avail_w as f32,
                avail_h as f32,
                orig_w as f32,
                orig_h as f32,
                preview_scale,
            );
            if scale > best_scale {
                best_scale = scale;
                best_quadrant = i;
            }
        }

        if best_scale <= 0.0 {
            return None;
        }

        let (avail_w, avail_h, _, _) = quadrants[best_quadrant];
        let max_width = avail_w.max(1) as u32;
        let max_height = avail_h.max(1) as u32;

        let (preview_w, preview_h) = scale_dimensions(
            orig_dims.0,
            orig_dims.1,
            max_width,
            max_height,
            preview_scale,
        );
        let media_width = preview_w as i32;
        let media_height = preview_h as i32;

        if media_width <= 0 || media_height <= 0 {
            return None;
        }

        let (pos_x, pos_y) = match best_quadrant {
            0 => (cursor_x + offset, cursor_y + offset),
            1 => (cursor_x - offset - media_width, cursor_y + offset),
            2 => (cursor_x + offset, cursor_y - offset - media_height),
            3 => (
                cursor_x - offset - media_width,
                cursor_y - offset - media_height,
            ),
            _ => (cursor_x + offset, cursor_y + offset),
        };

        let layout = PreviewLayout {
            pos_x,
            pos_y,
            max_width,
            max_height,
            preview_w,
            preview_h,
        };

        // The spinner is placed and left there: the step off the name is what being at
        // the pointer's own corner is instead of.
        if at_the_pointer_corner {
            return Some(layout);
        }

        Some(avoiding_text(
            layout,
            orig_dims,
            preview_scale,
            Clearance {
                text: avoid,
                gap: offset,
                cursor: Some((cursor_x, cursor_y)),
            },
            bounds,
            dpi,
        ))
    } else {
        let left_width = cursor_x - bounds.left - offset;
        let right_width = bounds.right - cursor_x - offset;
        let full_height = bounds.height();

        let left_scale = scale_in_room(
            left_width as f32,
            full_height as f32,
            orig_w as f32,
            orig_h as f32,
            preview_scale,
        );
        let right_scale = scale_in_room(
            right_width as f32,
            full_height as f32,
            orig_w as f32,
            orig_h as f32,
            preview_scale,
        );

        let (use_left, max_width, max_height) = if left_scale > right_scale && left_width > 0 {
            (true, left_width.max(1) as u32, full_height as u32)
        } else if right_width > 0 {
            (false, right_width.max(1) as u32, full_height as u32)
        } else {
            return None;
        };

        let (preview_w, preview_h) = scale_dimensions(
            orig_dims.0,
            orig_dims.1,
            max_width,
            max_height,
            preview_scale,
        );
        let media_width = preview_w as i32;
        let media_height = preview_h as i32;

        if media_width <= 0 || media_height <= 0 {
            return None;
        }

        let pos_x = if use_left {
            cursor_x - offset - media_width
        } else {
            cursor_x + offset
        };
        // Best position means beside the cursor, not in the middle of the
        // display. Centering on the cursor's own line keeps a small preview where
        // the pointer is instead of floating at the screen's center — which for a
        // cursor near the top put most of the preview below it — and the clamp is
        // what still keeps a tall preview inside the display.
        let pos_y = centered_top(cursor_y, media_height, bounds);

        let layout = PreviewLayout {
            pos_x,
            pos_y,
            max_width,
            max_height,
            preview_w,
            preview_h,
        };

        Some(avoiding_text(
            layout,
            orig_dims,
            preview_scale,
            Clearance {
                text: avoid,
                gap: offset,
                cursor: Some((cursor_x, cursor_y)),
            },
            bounds,
            dpi,
        ))
    }
}

/// How far off the item a keyboard preview is placed, in logical pixels.
pub(super) const KEYBOARD_GAP_PIXELS: f32 = 10.0;

/// The least room beside an item a keyboard preview will squeeze into before it
/// stops treating the item as something to sit beside, in logical pixels. Below this
/// the free space past the item's edge — or past the region a row is kept off — is a
/// sliver, and the preview is placed from the item's middle instead — see
/// `compute_keyboard_layout`.
pub(super) const MIN_BESIDE_ROOM_PIXELS: f32 = 64.0;

/// What placing a keyboard preview needs: the item it is about — its box, the region
/// the `Avoid` setting keeps it off, and whether its text is drawn as a row — and the
/// size and mode the placement is made at.
///
/// It is the keyboard's answer to `HoverPlacement`, which is the same set of facts
/// read at a cursor rather than at the focused item, and it is asked for in the same
/// way: once as the preview opens, and again with the size a text preview measured
/// itself at, which is the one thing that changes between the two.
#[derive(Clone, Copy)]
pub(super) struct KeyboardPlacement {
    /// The box the item occupies on screen, which a box item's preview is placed
    /// beside and a row's is not — see `compute_keyboard_layout`.
    pub(super) item_rect: (i32, i32, i32, i32),
    /// The region the `Avoid` setting keeps a preview off: the item's own text, the
    /// name alone, or the column the name sits in, as the setting has it — or `None`
    /// when nothing is kept off, which a keyboard preview is never asked with: the hook
    /// reads `Avoid Nothing` as `Avoid Filename` and answers with the item's own box
    /// when the view reported no text (see
    /// `explorer_hook::HoveredItem::keyboard_avoid_box`).
    ///
    /// It is read as the item's own text whatever the setting measured it as, rather
    /// than as a column the view draws every row in: a column's ways out are its own
    /// two sides, and a keyboard preview of a row is already placed past the region's
    /// right edge (see `compute_keyboard_layout`), so the region is left as what a
    /// fallback placement is stepped off.
    pub(super) avoid: Option<ScreenRegion>,
    /// Whether the item draws anything beside the piece its name is drawn in, which
    /// with the item's shape is what says it is a row of its view rather than a box —
    /// see `explorer_hook::ItemText`.
    pub(super) columns: bool,
    /// The media's own size, which the preview is scaled into the room by.
    pub(super) orig_dims: (u32, u32),
    /// The position mode: whether the preview grows away from the item or is centred
    /// beside it.
    pub(super) follow_cursor: bool,
    /// The share of the media's own size the preview is drawn at.
    pub(super) preview_scale: PreviewScale,
}

/// Compute preview layout for keyboard hover (relative to the focused item's box)
/// Positions the preview so it doesn't block the selected file item
///
/// `placement` is what the item asks for: its box, the region the `Avoid` setting
/// keeps a preview off — the boxes each piece of the item's own text is drawn in,
/// which the hook reads off the item's children in one batched call (see
/// `explorer_hook::item_text_box`) — whether the item draws its text as a row of its
/// view, and the size and mode the placement is made at. The region is what the
/// placement is kept clear of *and* where a row's placement is measured from, so a row
/// is only cleared as far as the setting asks; with nothing kept off, an item is
/// placed by the position mode alone. See `avoiding_text` and `KeyboardPlacement`.
///
/// A row of the view is one by the columns it draws beside its name rather than by its
/// box: the row's box is as wide as the *view* the row is drawn in, not as wide as the
/// display, so a `Details` row of a window a quarter of the display across is a row all
/// the same — and read by its box it would be placed past the whole of its columns at
/// every way of avoiding.
///
/// `dpi` is the display the item is on, which is what the margins this is written
/// around — the gap it keeps off the item and the least room beside one that is worth
/// sitting in — are scaled by.
pub(super) fn compute_keyboard_layout(
    placement: KeyboardPlacement,
    bounds: ScreenBounds,
    dpi: u32,
) -> Option<PreviewLayout> {
    let KeyboardPlacement {
        item_rect,
        avoid,
        columns,
        orig_dims,
        follow_cursor,
        preview_scale,
    } = placement;

    let (item_left, item_top, item_right, item_bottom) = item_rect;
    let gap = logical_px(dpi, KEYBOARD_GAP_PIXELS);
    let min_beside_room = logical_px(dpi, MIN_BESIDE_ROOM_PIXELS);
    let (orig_w, orig_h) = (orig_dims.0 as i32, orig_dims.1 as i32);

    // An item far wider than it is tall whose text is drawn as a row — the name with
    // the columns of a `Details` or `Content` row beside it, which the hook reads off
    // the item's own text — is a row of the list: a box as wide as the view with its
    // text written into the left end of it, whatever the view is doing on the display.
    let item_width = (item_right - item_left).max(0);
    let item_height = (item_bottom - item_top).max(1);
    let row_shaped = columns && item_width >= item_height * 4;

    // What is *beside* a row is not what its edges leave: the room past the row's
    // right edge is the space the view itself is not using — a sliver at the
    // window's edge, which is where a preview squeezed beside a row used to land.
    // The room a row really offers is the empty tail past the region the `Avoid`
    // setting keeps a preview off, and a keyboard preview is placed in it: just past
    // that region's edge, at the row's own line, sized by the tail and the display's
    // height. That is the placement a box item gets past its right edge, with the
    // region's edge standing in for the box's — so at `Avoid Details` it is the
    // placement a `Details` row has always had, past all of its columns, while at
    // `Avoid Filename` the preview is only taken past the name: the columns drawn
    // after it are within what the setting allows a preview to cover.
    //
    // A row with no tail — a narrow view, a name long enough to fill it — and a row
    // with no region to be placed from at all are both left to the placement below,
    // which anchors them at their middle: there is nowhere beside such a row to put a
    // preview, and the display's own room is all there is. (The keyboard path is never
    // asked with no region — see `KeyboardPlacement`.)
    if row_shaped {
        if let Some(tail_right) = avoid
            .map(|(_, _, right, _)| right)
            .filter(|right| *right > item_left && bounds.right - *right - gap >= min_beside_room)
        {
            let max_width = (bounds.right - tail_right - gap).max(1) as u32;
            let room_below = bounds.bottom - item_bottom - gap;
            let room_above = item_top - bounds.top - gap;
            // Follow Cursor grows the preview away from the row — from below it
            // when the larger room is there, from above it when it is not — while
            // Best Position centres it on the row, the way the mouse path centres
            // one on the cursor's line.
            let max_height = if follow_cursor {
                room_below.max(room_above).max(1) as u32
            } else {
                bounds.height().max(1) as u32
            };

            let (preview_w, preview_h) = scale_dimensions(
                orig_dims.0,
                orig_dims.1,
                max_width,
                max_height,
                preview_scale,
            );
            if preview_w == 0 || preview_h == 0 {
                return None;
            }

            let pos_x = tail_right + gap;
            let pos_y = if !follow_cursor {
                centered_top((item_top + item_bottom) / 2, preview_h as i32, bounds)
            } else if room_below >= room_above {
                item_bottom + gap
            } else {
                item_top - gap - preview_h as i32
            };

            let layout = PreviewLayout {
                pos_x,
                pos_y,
                max_width,
                max_height,
                preview_w,
                preview_h,
            };

            return Some(avoiding_text(
                layout,
                orig_dims,
                preview_scale,
                Clearance {
                    text: avoid.map(AvoidRegion::text),
                    gap,
                    cursor: None,
                },
                bounds,
                dpi,
            ));
        }
    }

    // What the placement is anchored at: the item's own edges for a box, and its
    // *middle* for a row that leaves no tail to be placed in (see above). Anchoring
    // at the middle is what the mouse path does with the cursor, so such a row is
    // read the way a hover over it is, and the preview is allowed to cover the rest
    // of it.
    let (anchor_left, anchor_top, anchor_right, anchor_bottom) = if row_shaped {
        let center_x = (item_left + item_right) / 2;
        let center_y = (item_top + item_bottom) / 2;
        (center_x, center_y, center_x, center_y)
    } else {
        (item_left, item_top, item_right, item_bottom)
    };

    if follow_cursor {
        // Quadrant-based positioning relative to what the item is anchored at
        let quadrants = [
            // Bottom-Right of it
            (
                bounds.right - anchor_right - gap,
                bounds.bottom - anchor_bottom - gap,
                anchor_right + gap,
                anchor_bottom + gap,
            ),
            // Bottom-Left of it
            (
                anchor_left - bounds.left - gap,
                bounds.bottom - anchor_bottom - gap,
                bounds.left,
                anchor_bottom + gap,
            ),
            // Top-Right of it
            (
                bounds.right - anchor_right - gap,
                anchor_top - bounds.top - gap,
                anchor_right + gap,
                bounds.top,
            ),
            // Top-Left of it
            (
                anchor_left - bounds.left - gap,
                anchor_top - bounds.top - gap,
                bounds.left,
                bounds.top,
            ),
        ];

        let mut best_quadrant = 0;
        let mut best_scale: f32 = 0.0;

        for (i, &(avail_w, avail_h, _, _)) in quadrants.iter().enumerate() {
            if avail_w <= 0 || avail_h <= 0 {
                continue;
            }
            let scale = scale_in_room(
                avail_w as f32,
                avail_h as f32,
                orig_w as f32,
                orig_h as f32,
                preview_scale,
            );
            if scale > best_scale {
                best_scale = scale;
                best_quadrant = i;
            }
        }

        if best_scale <= 0.0 {
            return None;
        }

        let (avail_w, avail_h, _, _) = quadrants[best_quadrant];
        let max_width = avail_w.max(1) as u32;
        let max_height = avail_h.max(1) as u32;

        let (preview_w, preview_h) = scale_dimensions(
            orig_dims.0,
            orig_dims.1,
            max_width,
            max_height,
            preview_scale,
        );
        let media_width = preview_w as i32;
        let media_height = preview_h as i32;

        if media_width <= 0 || media_height <= 0 {
            return None;
        }

        let (pos_x, pos_y) = match best_quadrant {
            0 => (anchor_right + gap, anchor_bottom + gap),
            1 => (anchor_left - gap - media_width, anchor_bottom + gap),
            2 => (anchor_right + gap, anchor_top - gap - media_height),
            3 => (
                anchor_left - gap - media_width,
                anchor_top - gap - media_height,
            ),
            _ => (anchor_right + gap, anchor_bottom + gap),
        };

        let layout = PreviewLayout {
            pos_x,
            pos_y,
            max_width,
            max_height,
            preview_w,
            preview_h,
        };

        Some(avoiding_text(
            layout,
            orig_dims,
            preview_scale,
            Clearance {
                text: avoid.map(AvoidRegion::text),
                gap,
                cursor: None,
            },
            bounds,
            dpi,
        ))
    } else {
        // Best spot mode: choose the left or right side of what the item is anchored
        // at — its own edges for a box, its middle for a row that leaves no tail to
        // be placed in (see above).
        //
        // The room a side offers is the room past that anchor, which for a box is
        // what keeps the preview off the file it describes. A box that leaves no room
        // on either side is placed from its middle with the display's own room, since
        // a preview squeezed into what is left past its edge is a sliver while one
        // placed from its middle takes the size the display allows; a row without a
        // tail is already anchored there.
        let edge_left_width = anchor_left - bounds.left - gap;
        let edge_right_width = bounds.right - anchor_right - gap;

        let (left_anchor_x, right_anchor_x, left_width, right_width) =
            if edge_left_width < min_beside_room && edge_right_width < min_beside_room {
                let center = ((anchor_left + anchor_right) / 2).clamp(bounds.left, bounds.right);
                (
                    center,
                    center,
                    center - bounds.left - gap,
                    bounds.right - center - gap,
                )
            } else {
                (anchor_left, anchor_right, edge_left_width, edge_right_width)
            };

        let full_height = bounds.height();

        let left_scale = scale_in_room(
            left_width as f32,
            full_height as f32,
            orig_w as f32,
            orig_h as f32,
            preview_scale,
        );
        let right_scale = scale_in_room(
            right_width as f32,
            full_height as f32,
            orig_w as f32,
            orig_h as f32,
            preview_scale,
        );

        let (use_left, max_width, max_height) = if left_scale > right_scale && left_width > 0 {
            (true, left_width.max(1) as u32, full_height as u32)
        } else if right_width > 0 {
            (false, right_width.max(1) as u32, full_height as u32)
        } else {
            return None;
        };

        let (preview_w, preview_h) = scale_dimensions(
            orig_dims.0,
            orig_dims.1,
            max_width,
            max_height,
            preview_scale,
        );
        let media_width = preview_w as i32;
        let media_height = preview_h as i32;

        if media_width <= 0 || media_height <= 0 {
            return None;
        }

        let pos_x = if use_left {
            left_anchor_x - gap - media_width
        } else {
            right_anchor_x + gap
        };
        // The same rule as the mouse path, centered on the line the preview
        // belongs to: beside the item it describes, not adrift in the display.
        let pos_y = centered_top((anchor_top + anchor_bottom) / 2, media_height, bounds);

        let layout = PreviewLayout {
            pos_x,
            pos_y,
            max_width,
            max_height,
            preview_w,
            preview_h,
        };

        Some(avoiding_text(
            layout,
            orig_dims,
            preview_scale,
            Clearance {
                text: avoid.map(AvoidRegion::text),
                gap,
                cursor: None,
            },
            bounds,
            dpi,
        ))
    }
}

/// A pinned window's box, kept on a display: a window dragged past an edge leaves a caption's
/// worth of itself behind, and a window dragged wholly off one is put back on it. The display
/// it is kept on is the one the caption is nearest — which is the one the hand is on.
///
/// The display arrives as an argument rather than being worked out here, and that is the whole
/// of why this function has tests at all. It is arithmetic on a rectangle, and it was untestable
/// because it reached into the display driver for one of its two arguments: the one clamp in
/// this file that decides where a window the user is looking at is allowed to be stranded, and
/// the only way to ask it anything was to attach a monitor to it (see `displays`).
///
/// Which display it is asked about is the decision the seam makes testable rather than this one:
/// the anchor is the box's own middle, so a window wide enough to have a middle on a display it
/// is mostly not on is kept on the display the hand is on rather than the one its left edge is
/// against, and a recorder reads that anchor back as a point (see `RecordedDisplays`).
pub(super) fn clamp_pinned_box(
    box_: ScreenRegion,
    dpi: u32,
    displays: &dyn Displays,
) -> ScreenRegion {
    let keep = logical_px(dpi, PIN_KEEP_ON_SCREEN_PIXELS).max(8);
    let width = (box_.2 - box_.0).max(1);
    let height = (box_.3 - box_.1).max(1);

    let anchor = displays::display_at(displays, box_.0 + width / 2, box_.1 + height / 2).work_area;
    let horizontal_keep = keep.min(width);
    let vertical_keep = keep.min(height);

    let mut left = box_.0;
    let mut top = box_.1;

    if left + horizontal_keep > anchor.right {
        left = anchor.right - horizontal_keep;
    }
    if left < anchor.left - width + horizontal_keep {
        left = anchor.left - width + horizontal_keep;
    }
    if top + vertical_keep > anchor.bottom {
        top = anchor.bottom - vertical_keep;
    }
    if top < anchor.top {
        top = anchor.top;
    }

    (left, top, left + width, top + height)
}

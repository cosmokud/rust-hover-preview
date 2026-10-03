use super::*;

/// A display to the right of `bounds()`, so the two can be told apart by the answer
/// rather than by their position.
fn the_second_display() -> ScreenBounds {
    ScreenBounds {
        left: 1000,
        top: 0,
        right: 2000,
        bottom: 800,
    }
}

/// How much of a pinned window is left on the display it is kept on, in the pixels of a
/// 100% display: a caption's worth, which is the band the buttons that close the pin are
/// drawn in (see `PIN_KEEP_ON_SCREEN_PIXELS`).
fn a_captions_worth() -> i32 {
    logical_px(TEST_DPI, PIN_KEEP_ON_SCREEN_PIXELS).max(8)
}

/// A locked window, a caption's worth of it left on the display.
///
/// The two failures this states are the ones the clamp exists for: a window dragged past
/// an edge keeps a caption on screen, so its buttons can still be pressed, and one
/// dragged wholly off is put back. Both were unreachable to a test, because the clamp
/// worked the display out for itself — the reader is `displays`, which is what this
/// passes in (see `clamp_pinned_box`).
#[test]
fn a_locked_window_keeps_a_caption_on_the_display_it_is_kept_on() {
    let desk = RecordedDisplays::one_display(bounds(), TEST_DPI);
    let keep = a_captions_worth();

    // Past the right edge: the caption's worth comes back, and the window keeps the size
    // the hand pulled it to rather than being squeezed into the room that is left.
    let dragged_right = clamp_pinned_box((1000, 300, 1400, 700), TEST_DPI, &desk);
    assert_eq!(
        dragged_right.0,
        bounds().right - keep,
        "a window dragged past the right edge is pulled back to leave its caption on screen"
    );
    assert_eq!(
        (
            dragged_right.2 - dragged_right.0,
            dragged_right.3 - dragged_right.1
        ),
        (400, 400),
        "and it is moved, not resized: the size a hand pulled it to is the size it keeps"
    );

    // Wholly off the right, and it is put back rather than kept as a caption on nothing.
    let dragged_away = clamp_pinned_box((1600, 300, 2000, 700), TEST_DPI, &desk);
    assert_eq!(
        dragged_away.0,
        bounds().right - keep,
        "a window dragged wholly off one display is put back on it"
    );
    assert_eq!(
        (
            dragged_away.2 - dragged_away.0,
            dragged_away.3 - dragged_away.1
        ),
        (400, 400),
        "at the size it was dragged to, which is what the pin's own bound is later read from"
    );

    // And the same at each of the other three edges, because a caption is on whichever side
    // the window was carried to and the four are four arms of the same decision.
    let past_bottom = clamp_pinned_box((300, 780, 700, 1180), TEST_DPI, &desk);
    assert_eq!(
        past_bottom.1,
        bounds().bottom - keep,
        "and at the bottom, from the other direction"
    );

    // The top edge is the odd one out and deliberately so: a caption is drawn *above* the
    // media, so a window pulled off the top has its buttons still on screen with no help
    // from a sliver, and the whole of the top edge is pulled back rather than a caption's
    // worth.
    let past_top = clamp_pinned_box((300, -400, 700, 0), TEST_DPI, &desk);
    assert_eq!(
        past_top.1,
        bounds().top,
        "a window off the top is put back on it wholly"
    );

    let past_left = clamp_pinned_box((-400, 300, 0, 700), TEST_DPI, &desk);
    assert_eq!(
        past_left.2,
        bounds().left + keep,
        "and at the left, from the other direction"
    );
}

/// A window wholly off a display is kept on the one its middle is on, which is the one
/// the caption is nearest — and for a box wider than the gap between two displays that is
/// the only thing there is to go on.
///
/// The anchor is the box's own middle and not its left edge, and the two disagree exactly
/// where this defect lives: a window carried from one display to the next has its left edge
/// over the display it came from, so a clamp that asked about the left edge would put the
/// user's window back on the monitor they dragged it off.
#[test]
fn a_locked_window_is_kept_on_the_display_its_middle_is_on() {
    let desk = RecordedDisplays::one_display(the_second_display(), TEST_DPI);

    // The left edge is over the first display and the middle is over the second, which is
    // the only case where the two disagree.
    let straddling = clamp_pinned_box((900, 300, 1300, 700), TEST_DPI, &desk);
    assert_eq!(
        straddling,
        (900, 300, 1300, 700),
        "a window the middle of which is on the second display is not moved at all, whatever \
             its left edge is over"
    );

    // And the point the clamp asked about is the middle, which a test can read back rather
    // than infer from the box that came out.
    let fresh = RecordedDisplays::one_display(bounds(), TEST_DPI);
    clamp_pinned_box((300, 300, 700, 700), TEST_DPI, &fresh);
    assert_eq!(
        fresh.points_asked_about(),
        vec![(500, 500)],
        "the display is asked about the middle of the box rather than its corner"
    );
}

/// The caption's worth is a distance under a hand, so it grows with the display it is
/// measured on — and a window on a 200% display is kept back further than the same window
/// on a 100% one.
///
/// The alternative is a caption left half the width a hand can hit, on the one display
/// where the pixels are twice as big and the window is being read from further away.
#[test]
fn the_room_a_caption_is_given_grows_with_the_display_it_is_measured_on() {
    let desk = RecordedDisplays::one_display(bounds(), 192);
    let clamped = clamp_pinned_box((1000, 300, 1400, 700), 192, &desk);

    assert_eq!(
        clamped.0,
        bounds().right - logical_px(192, PIN_KEEP_ON_SCREEN_PIXELS).max(8),
        "a 200% display keeps a 200% caption, which is the same size under a hand"
    );
}

/// A window kept on a display whose scale could not be read is kept on the primary one
/// rather than on the union of them.
///
/// The machine this stands in for is the one a display-change arrives on, and it is the
/// only caller of this clamp that could not previously be reached at all: the fallback was
/// inside the function that did the Win32 call, so a test could not give it a machine that
/// refuses (see `displays::display_at`).
#[test]
fn a_window_on_a_display_that_cannot_be_named_is_kept_on_the_primary_one() {
    // The same drag, against a machine that can name a display and against one that cannot. The
    // second answers the primary's right edge, which is a rectangle covering one display and
    // not the union of them — a union would leave the window where the hand left it.
    let dragged = (1000, 300, 1400, 700);
    let named = RecordedDisplays::one_display(the_second_display(), TEST_DPI);
    let unnamed = RecordedDisplays::no_display_to_name(bounds());
    assert_eq!(
        clamp_pinned_box(dragged, TEST_DPI, &named),
        dragged,
        "a machine that names the second display leaves the window where the hand put it"
    );
    assert_eq!(
        clamp_pinned_box(dragged, TEST_DPI, &unnamed).0,
        bounds().right - a_captions_worth(),
        "and one that cannot names the primary, whose own right edge is what it is kept \
             against — a rectangle covering one display, not the union of them"
    );
}

#[test]
fn leaves_a_placement_that_is_already_clear_of_the_text() {
    // The name is drawn to the left of the cursor's column, so the preview beside
    // the cursor is already off it.
    let name = (100, 300, 400, 320);
    let placement = placed(layout(420, 300, 300, 300), (300, 300), name, None, bounds());

    assert_eq!((placement.pos_x, placement.pos_y), (420, 300));
}

#[test]
fn moves_a_preview_out_of_the_name_it_covers() {
    // A row of a Details view: the pointer's item draws its name in a band twenty
    // pixels tall, and the preview came out beside the cursor with its top inside
    // that band. Down is the shortest way out, so it ends up just under the name,
    // where its own column already was.
    let name = (100, 300, 400, 320);
    let placement = placed(layout(120, 300, 300, 300), (300, 300), name, None, bounds());

    assert_eq!((placement.pos_x, placement.pos_y), (120, 340));
    assert_eq!((placement.preview_w, placement.preview_h), (300, 300));
}

/// A column of the view — the `Name` column at `Avoid Filename Column`, the columns
/// a row writes beside its name at `Avoid Details` — is not the text of the one item
/// the hover is about: every row draws its text in it, so a preview stepped off the
/// item's own row and left at the column's width covers the rows beside it, which is
/// exactly what the setting keeps a preview off. The ways out of a column are the two
/// to its sides, however much room the ways under and over it offer — see
/// `AvoidRegion` and `avoiding_text`.
#[test]
fn steps_a_preview_off_a_column_to_the_side_and_not_under_the_row() {
    // The `Name` column the name above is drawn in — the same region, as the way of
    // avoiding that keeps the whole column off measures it.
    let column = AvoidRegion::column((100, 300, 400, 320));

    // The placement the text of one item is stepped *under* the row by — see
    // `moves_a_preview_out_of_the_name_it_covers` — is taken past the column's right
    // edge instead: the name the preview came out over is the one of the item's own
    // row, and the step leaves every row's name readable.
    let placement = placed_off(
        layout(120, 300, 300, 300),
        (300, 300),
        column,
        None,
        bounds(),
    );
    assert_eq!((placement.pos_x, placement.pos_y), (420, 300));
    assert_eq!((placement.preview_w, placement.preview_h), (300, 300));

    // A preview that came out *below* the row is off the column's width as well,
    // rather than off the row's own line: what the last row of the display draws in
    // the column is covered by a preview the row's own step would clear.
    let placement = placed_off(
        layout(120, 500, 300, 300),
        (300, 300),
        column,
        None,
        bounds(),
    );
    assert_eq!((placement.pos_x, placement.pos_y), (420, 500));
}

/// A preview beside the column is left where the position mode put it, above or
/// below the item's own row: what a column is kept off is the column, and a preview
/// clear of it across covers no row's text wherever it sits.
#[test]
fn leaves_a_placement_beside_a_column_where_it_is() {
    let column = AvoidRegion::column((100, 300, 400, 320));

    let below = placed_off(
        layout(420, 500, 300, 300),
        (300, 300),
        column,
        None,
        bounds(),
    );
    assert_eq!((below.pos_x, below.pos_y), (420, 500));

    let above = placed_off(layout(420, 0, 300, 300), (300, 300), column, None, bounds());
    assert_eq!((above.pos_x, above.pos_y), (420, 0));
}

/// Where the display has no room beside a column, the ways out of the item's own
/// text are the ones left: a preview over the neighbours' names is then the lesser
/// thing, because the item the preview describes is the one it would cover instead.
/// The ways under and over a column are read off the row the column was measured at
/// for exactly this — see `avoiding_text`.
#[test]
fn steps_a_preview_under_the_row_where_the_display_leaves_no_room_beside_a_column() {
    // A column reaching both edges of the display, so neither side of it is a way
    // out: the room past its right edge and the room before its left one are both
    // nothing.
    let column = AvoidRegion::column((0, 300, 1000, 320));
    let placement = placed_off(
        layout(120, 300, 300, 300),
        (300, 300),
        column,
        None,
        bounds(),
    );

    assert_eq!((placement.pos_x, placement.pos_y), (120, 340));
    assert_eq!((placement.preview_w, placement.preview_h), (300, 300));
}

/// A way out that would put the preview over the pointer is not one a mouse hover
/// can use. A mouse preview is dismissed the moment the cursor touches it, so one
/// placed over the cursor is dismissed the instant it appears — and put back the
/// same way a moment later, for as long as the pointer sits there, which is a
/// preview that blinks at the hand rather than one that is read. The tile figures
/// are the case that reaches it: the pointer is on a thumbnail with the label below
/// it, and the step that clears the label by rising over it is the thumbnail the
/// pointer is standing on — see `avoiding_text`.
#[test]
fn keeps_a_placement_off_the_pointer_it_is_for() {
    let label = (400, 600, 700, 620);
    let thumbnail = (500, 450);

    // With nothing said about the pointer, the shortest step that clears the label
    // is the one over it — which lands on the thumbnail the pointer is on.
    let blind = placed(
        layout(420, 400, 300, 300),
        (300, 300),
        label,
        None,
        bounds(),
    );
    assert_eq!((blind.pos_x, blind.pos_y), (420, 280));
    assert!(
        box_holds(
            thumbnail.0,
            thumbnail.1,
            (
                blind.pos_x,
                blind.pos_y,
                blind.pos_x + blind.preview_w as i32,
                blind.pos_y + blind.preview_h as i32,
            ),
        ),
        "the step over the label is the thumbnail the pointer is on"
    );

    // Held to the pointer, the same hover takes the way out that clears both: the
    // room before the label, which the thumbnail's own column has.
    let placement = placed(
        layout(420, 400, 300, 300),
        (300, 300),
        label,
        Some(thumbnail),
        bounds(),
    );
    assert_eq!((placement.pos_x, placement.pos_y), (80, 400));
    assert!(
        !box_holds(
            thumbnail.0,
            thumbnail.1,
            (
                placement.pos_x,
                placement.pos_y,
                placement.pos_x + placement.preview_w as i32,
                placement.pos_y + placement.preview_h as i32,
            ),
        ),
        "the preview is clear of the pointer"
    );
    assert!(
        placement.pos_x + placement.preview_w as i32 <= label.0,
        "and still clear of the label it was moved off"
    );
}

#[test]
fn takes_the_side_when_the_display_has_no_room_under_the_name() {
    // The same placement on a display that ends below it: the preview cannot drop
    // under the name at its own size, so it goes past the name's right edge, which
    // the display has room for.
    let name = (100, 300, 400, 320);
    let short = ScreenBounds {
        bottom: 500,
        ..bounds()
    };
    let placement = placed(layout(120, 300, 300, 300), (300, 300), name, None, short);

    assert_eq!((placement.pos_x, placement.pos_y), (420, 300));
    assert_eq!((placement.preview_w, placement.preview_h), (300, 300));
}

#[test]
fn resizes_a_preview_too_large_to_get_off_the_name() {
    // A preview as large as the display allows, beside a row whose name spans most
    // of it: no way out holds the preview as it stands, so the size is what gives
    // and the roomiest way out is taken. Under the name, that is the full width the
    // mode allowed and the height the display leaves below the row.
    let name = (100, 300, 700, 320);
    let placement = placed(layout(0, 0, 1000, 800), (1000, 800), name, None, bounds());

    assert_eq!((placement.pos_x, placement.pos_y), (0, 340));
    assert_eq!((placement.preview_w, placement.preview_h), (575, 460));
    // The name it was moved off is clear of it: the top edge is the row's bottom
    // plus the gap the claim above was measured with.
    assert!(placement.pos_y >= 320 + 20);
}

#[test]
fn leaves_a_preview_alone_when_every_way_out_is_a_sliver() {
    // A display with no room worth having on any side of the name: the placement is
    // what the mode chose, since a preview squeezed into a sliver says less than the
    // one left over the name.
    let name = (40, 80, 360, 100);
    let tight = ScreenBounds {
        left: 0,
        top: 0,
        right: 400,
        bottom: 180,
    };
    let placement = placed(layout(150, 80, 200, 90), (200, 90), name, None, tight);

    assert_eq!((placement.pos_x, placement.pos_y), (150, 80));
}

#[test]
fn keeps_the_preview_inside_the_display_it_moves_on() {
    // Clearing the name to the right falls short of what the preview needs here, so
    // the step under the name is taken instead — resized into the room it leaves
    // when even that is not enough for it as it stands.
    let name = (100, 300, 700, 320);
    let narrow = ScreenBounds {
        right: 900,
        ..bounds()
    };
    let placement = placed(layout(690, 100, 200, 300), (200, 300), name, None, narrow);

    // 720 is the name's right edge plus the gap, and 720 + 200 leaves the display;
    // 340 is its bottom plus the gap, and 340 + 300 does not.
    assert_eq!((placement.pos_x, placement.pos_y), (690, 340));
}

/// The room the display has is its whole work area: the largest box anything shown on
/// it can be drawn in, and the box an engine that has to draw a preview before the file
/// can be measured is asked for (see `PendingLoad::room`). A work area with no room in
/// it is a pixel rather than nothing, because a box of nothing is an engine asked to
/// write nothing.
#[test]
fn a_display_room_is_the_whole_of_its_work_area() {
    assert_eq!(bounds().room(), (1000, 800));

    assert_eq!(
        ScreenBounds {
            left: 100,
            top: 40,
            right: 1800,
            bottom: 1040,
        }
        .room(),
        (1700, 1000),
        "the room is the work area whatever corner it is anchored at"
    );

    assert_eq!(
        ScreenBounds {
            left: 0,
            top: 0,
            right: 0,
            bottom: 0,
        }
        .room(),
        (1, 1),
        "a display with no room is asked for a pixel rather than for nothing"
    );
}

/// A planned preview is inside the display it was planned for: the box it takes is
/// within the room it was given, that room is within the room the display itself has,
/// and the whole of it — place and size, both axes — is within that display's work area.
fn assert_inside_the_display(layout: PreviewLayout, bounds: ScreenBounds) {
    assert!(
        layout.preview_w <= layout.max_width && layout.preview_h <= layout.max_height,
        "the preview takes {} by {} of the {} by {} it was given",
        layout.preview_w,
        layout.preview_h,
        layout.max_width,
        layout.max_height
    );

    // The room a placement hands its preview is taken out of the display — the side of
    // the pointer it goes beside, the quadrant it grows into, the room left past a name —
    // so the display's own room bounds every one of them. That is what makes it the box
    // to ask an engine for when the file cannot be measured before it is drawn: a picture
    // developed into a smaller box than this could never be drawn at the size a layout
    // asks for (see `PendingLoad::room`).
    let room = bounds.room();
    assert!(
        layout.max_width <= room.0 && layout.max_height <= room.1,
        "the room a layout was given, {} by {}, leaves the display's own room of {} by {}",
        layout.max_width,
        layout.max_height,
        room.0,
        room.1
    );

    assert!(
        layout.pos_x >= bounds.left
            && layout.pos_y >= bounds.top
            && layout.pos_x + layout.preview_w as i32 <= bounds.right
            && layout.pos_y + layout.preview_h as i32 <= bounds.bottom,
        "the preview at {},{} is {} by {}, which leaves the display of {} by {} at {},{}",
        layout.pos_x,
        layout.pos_y,
        layout.preview_w,
        layout.preview_h,
        bounds.right - bounds.left,
        bounds.bottom - bounds.top,
        bounds.left,
        bounds.top
    );
}

/// A display and the scale it is drawn at, which is one case for a placement test.
type ScaledDisplay = (ScreenBounds, u32);

/// What the placement matrices are made of.
struct PlacementCases {
    displays: Vec<ScaledDisplay>,
    media: [(u32, u32); 7],
    scales: [PreviewScale; 4],
}

/// The displays, the display scales and the media every placement test is run
/// over: a 1080p display, a 4K one, a 4K one that is the second display rather than
/// the first, and a portrait one — each at 100%, 150% and 200%, because the margins
/// a placement is written around are scaled by the display it is on and the ones
/// that came apart at a scale are what the scaling is for — against the shapes media
/// comes in: pages both ways up, a slide, the waiting spinner, a picture, and things
/// far wider and far taller than any display.
fn placement_cases() -> PlacementCases {
    let geometries = [
        ScreenBounds {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1040,
        },
        ScreenBounds {
            left: 0,
            top: 0,
            right: 3840,
            bottom: 2090,
        },
        ScreenBounds {
            left: 1920,
            top: 0,
            right: 5760,
            bottom: 2090,
        },
        ScreenBounds {
            left: 0,
            top: 0,
            right: 1440,
            bottom: 2500,
        },
    ];
    let media = [
        (794, 1123),
        (1123, 794),
        (1920, 1080),
        (36, 36),
        (100, 100),
        (8000, 120),
        (120, 8000),
    ];
    let scales = [
        PreviewScale::FitToScreen,
        PreviewScale::FitToScreenReduced(50),
        PreviewScale::Percent(100),
        PreviewScale::Percent(400),
    ];
    // Every display at every scale, as one list: the margins a placement is
    // written around are scaled by the display it is on, so a display at 150% is
    // not the display at 100%, and the pair is what a case is.
    let displays = geometries
        .iter()
        .flat_map(|bounds| [TEST_DPI, 144, 192].map(|dpi| (*bounds, dpi)))
        .collect();

    PlacementCases {
        displays,
        media,
        scales,
    }
}

/// The one thing a hovered preview may never do: leave the display it was planned
/// for. Every position mode, every scale, every shape of media, anchored at each
/// corner of the display and at its middle, with a name under the pointer and with a
/// row across the display's top to be kept off — each as the text one item draws and
/// as a column the view draws every row in — because a preview that grows past its
/// display takes an edge and a room that disagree to find, and the ones that
/// disagree are not the ones anyone hovers over on purpose.
#[test]
fn places_every_hover_inside_its_display() {
    let PlacementCases {
        displays,
        media,
        scales,
    } = placement_cases();
    let mut placed = 0usize;

    for (bounds, dpi) in displays {
        let points = [
            (bounds.left, bounds.top),
            (bounds.right - 1, bounds.top),
            (bounds.left, bounds.bottom - 1),
            (bounds.right - 1, bounds.bottom - 1),
            (
                (bounds.left + bounds.right) / 2,
                (bounds.top + bounds.bottom) / 2,
            ),
        ];

        for (orig_width, orig_height) in media {
            for preview_scale in scales {
                for follow_cursor in [true, false] {
                    for (cursor_x, cursor_y) in points {
                        let regions = [
                            None,
                            // The name of the item the pointer is on.
                            Some(AvoidRegion::text((
                                cursor_x - 100,
                                cursor_y - 8,
                                cursor_x + 100,
                                cursor_y + 8,
                            ))),
                            // A row's text across the whole display's top, which a
                            // preview above the middle of the display covers.
                            Some(AvoidRegion::text((
                                bounds.left,
                                bounds.top,
                                bounds.right,
                                bounds.top + 40,
                            ))),
                            // The same two regions read as columns of the view — as
                            // `Avoid Filename Column` and `Avoid Details` measure
                            // them — which a placement is kept off to the side (see
                            // `AvoidRegion`).
                            Some(AvoidRegion::column((
                                cursor_x - 100,
                                cursor_y - 8,
                                cursor_x + 100,
                                cursor_y + 8,
                            ))),
                            Some(AvoidRegion::column((
                                bounds.left,
                                bounds.top,
                                bounds.right,
                                bounds.top + 40,
                            ))),
                        ];

                        for avoid in regions {
                            let placement = HoverPlacement {
                                orig_dims: (orig_width, orig_height),
                                avoid,
                                follow_cursor,
                                preview_scale,
                                // Only the waiting spinner is placed at the
                                // pointer's own corner, and only it is ever that size.
                                at_the_pointer_corner: (orig_width, orig_height) == (36, 36),
                            };

                            let layout =
                                compute_mouse_layout(cursor_x, cursor_y, placement, bounds, dpi);
                            let Some(layout) = layout else {
                                continue;
                            };

                            assert_inside_the_display(layout, bounds);
                            placed += 1;
                        }
                    }
                }
            }
        }
    }

    // A matrix that answers nothing checks nothing: nearly every combination of
    // point, media, scale and setting has a layout to check.
    assert!(
        placed >= 3000,
        "{placed} of the matrix's layouts were placed, which is too few to have checked the rest"
    );
}

/// The same promise for the preview a keyboard hover raises, which is placed from
/// the item rather than from the pointer: a row of a list, a box item, and an item
/// the display has scrolled half off its top and half off its bottom — the cases
/// where the room beside an item and the item's own edges disagree.
#[test]
fn places_every_keyboard_hover_inside_its_display() {
    let PlacementCases {
        displays,
        media,
        scales,
    } = placement_cases();
    let mut placed = 0usize;

    for (bounds, dpi) in displays {
        let items = [
            // A row of a list, drawn across the view at its middle.
            (
                bounds.left,
                (bounds.top + bounds.bottom) / 2,
                bounds.right,
                (bounds.top + bounds.bottom) / 2 + 40,
            ),
            // A box item, the way an icon view draws one.
            (
                bounds.left + 100,
                bounds.top + 100,
                bounds.left + 220,
                bounds.top + 220,
            ),
            // The first row of a scrolled list, half off the display's top.
            (bounds.left, bounds.top - 20, bounds.right, bounds.top + 20),
            // And the last one, half off its bottom.
            (
                bounds.left,
                bounds.bottom - 20,
                bounds.right,
                bounds.bottom + 20,
            ),
        ];

        for (item_left, item_top, item_right, item_bottom) in items {
            for (orig_width, orig_height) in media {
                for preview_scale in scales {
                    for follow_cursor in [true, false] {
                        let avoids = [
                            None,
                            // The name the row is listed under, at its left end.
                            Some((
                                item_left + 8,
                                item_top + 8,
                                item_left + 220,
                                item_bottom - 8,
                            )),
                            // And the row's whole width, as `Avoid Details` reads it.
                            Some((item_left, item_top, item_right, item_bottom)),
                        ];

                        // Read as a row of the view and as a box item, since what
                        // the item draws beside its name decides which of the two
                        // placements it gets and both are promises of their own:
                        // the item is placed from the region it is kept off or
                        // from its own edges, whichever its text says it is.
                        for columns in [true, false] {
                            for avoid in avoids {
                                let Some(layout) = compute_keyboard_layout(
                                    KeyboardPlacement {
                                        item_rect: (item_left, item_top, item_right, item_bottom),
                                        avoid,
                                        columns,
                                        orig_dims: (orig_width, orig_height),
                                        follow_cursor,
                                        preview_scale,
                                    },
                                    bounds,
                                    dpi,
                                ) else {
                                    continue;
                                };

                                assert_inside_the_display(layout, bounds);
                                placed += 1;
                            }
                        }
                    }
                }
            }
        }
    }

    assert!(
        placed >= 2000,
        "{placed} of the matrix's layouts were placed, which is too few to have checked the rest"
    );
}

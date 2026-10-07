use super::*;

/// A keyboard placement of one item, at the size and mode the figures are easy to
/// read in: the media's own size at 100%, and `Best Position`, so the place comes
/// out of the room beside the item alone. `columns` is whether the item is read as
/// a row of its view or as a box item — see `KeyboardPlacement`.
fn keyboard_placement(
    item: (i32, i32, i32, i32),
    avoid: Option<ScreenRegion>,
    columns: bool,
) -> KeyboardPlacement {
    KeyboardPlacement {
        item_rect: item,
        avoid,
        columns,
        orig_dims: (400, 300),
        follow_cursor: false,
        preview_scale: PreviewScale::Percent(100),
    }
}

/// A keyboard preview of a row is placed in the room past the region the `Avoid`
/// setting keeps it off, so a row is cleared only as far as the setting asks: past
/// every column at `Avoid Details`, and only past the name at `Avoid Filename`,
/// where the columns drawn after the name are the room the preview takes.
#[test]
fn a_keyboard_rows_tail_begins_at_the_region_it_is_kept_off() {
    let row = (0, 100, 1000, 140);

    let past_the_name = compute_keyboard_layout(
        keyboard_placement(row, Some((20, 104, 120, 136)), true),
        bounds(),
        TEST_DPI,
    )
    .expect("a placement past the name");

    let past_every_column = compute_keyboard_layout(
        keyboard_placement(row, Some((20, 104, 900, 136)), true),
        bounds(),
        TEST_DPI,
    )
    .expect("a placement past the row's columns");

    assert_eq!(past_the_name.pos_x, 130, "just past the name");
    assert_eq!(past_every_column.pos_x, 910, "just past the columns");
    assert!(
        past_the_name.pos_x < past_every_column.pos_x,
        "a narrower region leaves more of the row to be covered"
    );
}

/// A row is a row whatever the window is doing: a `Details` row of a window
/// narrower than half the display still takes its placement from the region the
/// `Avoid` setting keeps it off — past the name, past the `Name` column, past the
/// row's columns — rather than from its own right edge, which is where the same
/// row read as a box would put its preview at every one of those settings. The row
/// below is a sixth of the display across, its text written into the left end of
/// it, as `Details` draws one.
#[test]
fn a_rows_placement_is_measured_from_the_region_whatever_the_window_is() {
    let row = (0, 100, 420, 124);
    let gap = logical_px(TEST_DPI, KEYBOARD_GAP_PIXELS);

    // The region each way of avoiding keeps off, and the edge it leaves the
    // preview past: the name's own width, the `Name` column, and the columns.
    let regions = [
        ((12, 103, 180, 122), 180, "past the name"),
        ((12, 103, 220, 122), 220, "past the `Name` column"),
        ((12, 103, 418, 122), 418, "past the row's columns"),
    ];

    for (region, region_right, what) in regions {
        let layout = compute_keyboard_layout(
            keyboard_placement(row, Some(region), true),
            bounds(),
            TEST_DPI,
        )
        .unwrap_or_else(|| panic!("a placement {what}"));

        assert_eq!(layout.pos_x, region_right + gap, "{what}");
    }

    // Which leaves the name's placement inside the row it belongs to, where the
    // row's own right edge would have put it had the row been read as a box.
    let past_the_name = compute_keyboard_layout(
        keyboard_placement(row, Some((12, 103, 180, 122)), true),
        bounds(),
        TEST_DPI,
    )
    .expect("a placement past the name");

    assert!(
        past_the_name.pos_x < row.2,
        "the name's tail is inside the row: {} against the row's own edge {}",
        past_the_name.pos_x,
        row.2
    );
}

/// An item that draws nothing beside its name — the label under an icon — is a box
/// item: its preview is placed beside the box itself whichever region is kept off,
/// since the label is a piece drawn inside that box and not a row of the view.
#[test]
fn a_box_items_tail_is_its_own_edge_and_not_its_labels() {
    let tile = (0, 100, 120, 220);
    let label = (25, 190, 95, 206);
    let gap = logical_px(TEST_DPI, KEYBOARD_GAP_PIXELS);

    for avoid in [Some(label), None] {
        let layout =
            compute_keyboard_layout(keyboard_placement(tile, avoid, false), bounds(), TEST_DPI)
                .expect("a placement beside the tile");

        assert_eq!(layout.pos_x, tile.2 + gap, "just past the tile's own edge");
    }
}

/// A row the placement was given no region for is placed by the position mode
/// alone: it is anchored at its middle, the way a hover over it is read, and the
/// preview is allowed to cover it. No keyboard preview is asked this way — the hook
/// reads `Avoid Nothing` as `Avoid Filename` (see `KeyboardPlacement`) — and the
/// answer is what a caller with no region to be placed from gets.
#[test]
fn a_keyboard_row_with_nothing_kept_off_is_placed_by_position_alone() {
    let layout = compute_keyboard_layout(
        keyboard_placement((0, 100, 1000, 140), None, true),
        bounds(),
        TEST_DPI,
    )
    .expect("a placement");

    assert_eq!(layout.pos_x, 510, "half the row's width, and the gap");
}

/// A page — a PDF's, or one a document was drawn as — is drawn at the room the display
/// has, because the room is free quality there. Every setting at or above
/// `100%` asks for at least that room, so they are one setting for a page; only
/// a setting below it is a size the user picked, and it is answered by
/// reducing the fitted size rather than by ignoring it. Which page setting is
/// read is the kind's own: a PDF page follows `ebook_scale` and a page drawn for
/// a document follows `document_scale`, and neither moves for the picture scale.
#[test]
fn a_page_takes_the_room_the_display_has() {
    let pdf = PathBuf::from(r"C:\docs\report.pdf");
    let vector_scale = DEFAULT_VECTOR_SCALE;
    let document_scale = PreviewScale::Percent(75);

    for configured in [
        PreviewScale::FitToScreen,
        PreviewScale::Percent(100),
        PreviewScale::Percent(200),
        PreviewScale::Percent(400),
    ] {
        assert_eq!(
            effective_preview_scale(
                &pdf,
                HoverScales {
                    picture: configured,
                    ebook: configured,
                    vector: vector_scale,
                    document: document_scale,
                    ..hover_scales()
                }
            ),
            PreviewScale::FitToScreen
        );
    }

    assert_eq!(
        effective_preview_scale(
            &pdf,
            HoverScales {
                picture: PreviewScale::Percent(400),
                ebook: PreviewScale::Percent(50),
                vector: vector_scale,
                document: document_scale,
                ..hover_scales()
            }
        ),
        PreviewScale::FitToScreenReduced(50)
    );
    assert_eq!(
        effective_preview_scale(
            &pdf,
            HoverScales {
                picture: PreviewScale::Percent(400),
                ebook: PreviewScale::Percent(25),
                vector: vector_scale,
                document: document_scale,
                ..hover_scales()
            }
        ),
        PreviewScale::FitToScreenReduced(25),
        "a page follows its own share of the display, whatever the picture scale says"
    );
    assert_eq!(
        effective_preview_scale(
            &pdf,
            HoverScales {
                picture: PreviewScale::Percent(400),
                ebook: PreviewScale::FitToScreen,
                vector: vector_scale,
                document: document_scale,
                ..hover_scales()
            }
        ),
        PreviewScale::FitToScreen
    );
}

/// A document's scale means the room the display has rather than the size the file
/// asks for, so `Fit to Screen` is the whole of that room and a percentage is a
/// share of it — half the display at `50%`, a tenth at `10%` — whatever scale the
/// pictures beside it are drawn at. A share the configuration can hold but the menu
/// does not offer is honored rather than rounded to a menu entry.
#[test]
fn a_document_is_drawn_at_its_share_of_the_room() {
    let svg = PathBuf::from(r"C:\art\clock.svg");
    let configured = PreviewScale::Percent(100);
    let ebook = PreviewScale::FitToScreen;

    assert_eq!(
        effective_preview_scale(
            &svg,
            HoverScales {
                picture: configured,
                vector: PreviewScale::FitToScreen,
                ebook,
                ..hover_scales()
            }
        ),
        PreviewScale::FitToScreen
    );
    for percent in [75, 50, 25, 10] {
        assert_eq!(
            effective_preview_scale(
                &svg,
                HoverScales {
                    picture: configured,
                    vector: PreviewScale::Percent(percent),
                    ebook,
                    ..hover_scales()
                }
            ),
            PreviewScale::FitToScreenReduced(percent),
            "{percent}% of the room"
        );
    }

    assert_eq!(
        effective_preview_scale(
            &svg,
            HoverScales {
                picture: configured,
                vector: PreviewScale::Percent(60),
                ebook,
                ..hover_scales()
            }
        ),
        PreviewScale::FitToScreenReduced(60),
        "a share the menu does not offer is the share it is"
    );

    for configured in [
        PreviewScale::FitToScreen,
        PreviewScale::Percent(100),
        PreviewScale::Percent(400),
    ] {
        assert_eq!(
            effective_preview_scale(
                &svg,
                HoverScales {
                    picture: configured,
                    vector: PreviewScale::Percent(100),
                    ebook,
                    ..hover_scales()
                }
            ),
            PreviewScale::FitToScreen,
            "asking for the whole room or more is the whole room"
        );
    }
}

/// A sound is laid out over a share of the display, the share the Audio
/// Scaling setting names: the room its card is measured at is the work
/// area of the display the pointer is on times that share, at the
/// display's own DPI. The width room is the share of the width plus
/// twice the card's own margin — the padding the card is drawn inside,
/// read at the font size the card is built at for the share, which is
/// the default text size scaled by the share's fraction of the 10%
/// anchor — and the height room is the plain share of the height, which
/// is only the guard that answers a too-short room with nothing. The
/// answer carries the font size the card is built at alongside the
/// room.
#[test]
fn a_sound_is_laid_out_over_its_share_of_the_display() {
    let display = ScreenBounds {
        left: 0,
        top: 0,
        right: 3440,
        bottom: 1440,
    };

    for (percent, room, font) in [
        (5, (188, 72), 63),
        (10, (374, 144), 125),
        (15, (562, 216), 188),
        (20, (748, 288), 250),
        (25, (936, 360), 313),
    ] {
        let answered =
            dimensions::audio_box_room(display, PreviewScale::Percent(percent), TEST_DPI);
        assert_eq!(
            (answered.width, answered.height),
            room,
            "{percent}% of a 3440x1440 work area"
        );
        assert_eq!(
            answered.font_scale_percent, font,
            "{percent}% of a 3440x1440 work area builds its card at {font}%"
        );
    }
}

/// A font is drawn at a share of the room the way a document is, and it is the fourth
/// setting of its own: the specimen is a page of this app's making, so the share decides
/// how large the type is drawn — and what a PDF, a document and a page are configured at
/// leaves it where it is.
#[test]
fn a_specimen_is_drawn_at_its_share_of_the_room() {
    let font = PathBuf::from(r"C:\fonts\Inter-Regular.woff2");
    let picture = PreviewScale::Percent(400);
    let ebook = PreviewScale::FitToScreen;
    let vector = PreviewScale::Percent(75);

    assert_eq!(
        effective_preview_scale(
            &font,
            HoverScales {
                picture,
                vector,
                ebook,
                ..hover_scales()
            }
        ),
        PreviewScale::FitToScreenReduced(DEFAULT_FONT_SCALE_PERCENT),
        "a specimen starts at half the room, whatever the other kinds are set to"
    );

    for percent in [75, 50, 25, 10] {
        assert_eq!(
            effective_preview_scale(
                &font,
                HoverScales {
                    picture,
                    vector,
                    ebook,
                    font: PreviewScale::Percent(percent),
                    ..hover_scales()
                }
            ),
            PreviewScale::FitToScreenReduced(percent),
            "{percent}% of the room"
        );
    }

    for configured in [
        PreviewScale::FitToScreen,
        PreviewScale::Percent(100),
        PreviewScale::Percent(400),
    ] {
        assert_eq!(
            effective_preview_scale(
                &font,
                HoverScales {
                    picture,
                    vector,
                    ebook,
                    font: configured,
                    ..hover_scales()
                }
            ),
            PreviewScale::FitToScreen,
            "asking for the whole room or more is the whole room"
        );
    }

    // A file that is not a font keeps the picture scale, whatever the specimen's is.
    let png = PathBuf::from(r"C:\art\photo.png");
    assert_eq!(
        effective_preview_scale(
            &png,
            HoverScales {
                picture: PreviewScale::Percent(200),
                vector,
                ebook,
                font: PreviewScale::Percent(10),
                ..hover_scales()
            }
        ),
        PreviewScale::Percent(200)
    );
}

/// The picture scale is about another kind of preview, so a document's size does
/// not move when it does: the two settings are read one each rather than one for
/// both.
#[test]
fn a_documents_size_is_not_the_pictures_setting() {
    let svg = PathBuf::from(r"C:\art\clock.svg");
    let picture = PreviewScale::Percent(100);

    assert_eq!(
        effective_preview_scale(
            &svg,
            HoverScales {
                picture,
                vector: DEFAULT_VECTOR_SCALE,
                ebook: PreviewScale::Percent(25),
                document: PreviewScale::Percent(75),
                font: PreviewScale::Percent(10),
                ..hover_scales()
            }
        ),
        PreviewScale::FitToScreen,
        "a document follows its own share, not a page's"
    );

    // And a picture is still drawn at the picture scale, whatever the document
    // settings say.
    let png = PathBuf::from(r"C:\art\photo.png");
    assert_eq!(
        effective_preview_scale(
            &png,
            HoverScales {
                picture: PreviewScale::Percent(200),
                vector: PreviewScale::FitToScreen,
                ebook: PreviewScale::Percent(25),
                document: PreviewScale::Percent(75),
                font: PreviewScale::Percent(10),
                ..hover_scales()
            }
        ),
        PreviewScale::Percent(200)
    );
}

/// A share of the fitted size is not a share of the media's own size, and the
/// difference is the point: `50%` is half the media — the same size whatever
/// room it has — while a reduced fit is half of what the room would have
/// allowed, so it grows with the room.
#[test]
fn reduces_the_fitted_size_by_the_configured_share() {
    // 800 by 600 in a 400 by 300 room: the fit is half the media's size, so
    // half of the fit is a quarter of it.
    assert_eq!(
        scale_dimensions(800, 600, 400, 300, PreviewScale::FitToScreen),
        (400, 300)
    );
    assert_eq!(
        scale_dimensions(800, 600, 400, 300, PreviewScale::FitToScreenReduced(50)),
        (200, 150)
    );
    assert_eq!(
        scale_dimensions(800, 600, 400, 300, PreviewScale::FitToScreenReduced(25)),
        (100, 75)
    );

    // The same media in a room that allows all of it: 50% of the media is 400
    // by 300, while half of the fitted 800 by 600 is not.
    assert_eq!(
        scale_dimensions(800, 600, 1900, 1000, PreviewScale::Percent(50)),
        (400, 300)
    );
    assert_eq!(
        scale_dimensions(800, 600, 1900, 1000, PreviewScale::FitToScreenReduced(50)),
        (667, 500)
    );
}

/// A document whose page has not been rendered yet is a spinner, and the
/// spinner is placed at its own size: fitted to the display it would be a
/// screen-sized square with a spinner drawn in the middle of it.
#[test]
fn a_document_with_no_page_yet_waits_in_the_spinners_own_box() {
    let waiting = PathBuf::from(r"C:\docs\not-rendered-yet.docx");
    let scale = effective_preview_scale(
        &waiting,
        HoverScales {
            picture: PreviewScale::Percent(400),
            vector: DEFAULT_VECTOR_SCALE,
            ebook: PreviewScale::Percent(25),
            document: PreviewScale::Percent(10),
            font: PreviewScale::Percent(DEFAULT_FONT_SCALE_PERCENT),
            ..hover_scales()
        },
    );

    assert_eq!(scale, PreviewScale::Percent(100));
    assert_eq!(
        scale_dimensions(
            office_preview::WAITING_BOX,
            office_preview::WAITING_BOX,
            1920,
            1080,
            scale,
        ),
        (office_preview::WAITING_BOX, office_preview::WAITING_BOX)
    );
}

/// A page's share of the display is read for a bitmap — a workbook's corner where
/// no page can be exported — as the same share of its own size, and the whole of
/// the display as the bitmap at the size it is: a quarter of the room is a quarter
/// of the picture, and a fit never stretches one to fill the screen. A reduced fit
/// asks for the room reduced to its share, so a bitmap is asked for that share of
/// its own size — the size the plain fit would have left it at, reduced.
#[test]
fn a_bitmap_follows_a_pages_share_without_being_enlarged() {
    for percent in [75, 50, 25, 10] {
        assert_eq!(
            bitmap_at_display_scale(PreviewScale::Percent(percent)),
            PreviewScale::Percent(percent),
            "{percent}% of the picture"
        );
        assert_eq!(
            bitmap_at_display_scale(PreviewScale::FitToScreenReduced(percent)),
            PreviewScale::Percent(percent),
            "{percent}% of the fitted picture"
        );
    }

    assert_eq!(
        bitmap_at_display_scale(PreviewScale::FitToScreen),
        PreviewScale::Percent(100),
        "the whole room is the picture at its own size"
    );
}

/// The spinner is the arc alone: the box it is drawn in is transparent, so a
/// hover that is waiting shows a spinner rather than a square of its own.
#[test]
fn draws_the_spinner_without_a_box_around_it() {
    let frame = render_loading_frame(64, 64, 0.0);
    assert_eq!(frame.len(), 64 * 64 * 4);

    // The corners — and so the box the arc is drawn in — are nothing at all.
    for corner in [0usize, 63, 64 * 63, 64 * 64 - 1] {
        assert_eq!(&frame[corner * 4..corner * 4 + 4], &[0, 0, 0, 0]);
    }

    // The arc is drawn, and the halo under it and the arc's own tail are
    // partly there rather than filled in.
    let alphas: Vec<u8> = frame
        .as_chunks::<4>()
        .0
        .iter()
        .map(|pixel| pixel[3])
        .collect();
    assert!(alphas.iter().any(|&alpha| alpha > 200), "the arc is drawn");
    assert!(
        alphas.iter().any(|&alpha| (1..=200).contains(&alpha)),
        "the halo and the tail fade rather than fill"
    );
}

/// The waiting frame is the size of the spinner in it: over a turn of the
/// spinner, the arc and its halo reach each of the box's edges, so a spinner
/// placed at the pointer's corner is the spinner at the pointer rather than an
/// empty frame around one.
#[test]
fn the_waiting_frame_is_the_size_of_the_spinner_in_it() {
    let side = office_preview::WAITING_BOX as usize;
    let (mut min_x, mut min_y) = (side, side);
    let (mut max_x, mut max_y) = (0, 0);

    // The arc ends in a tail that fades to nothing, so one frame has fewer of
    // the box's edges inked than the next: the frame is measured against the
    // whole turn rather than against the arc's position in it.
    for quarter in 0..4 {
        let angle = quarter as f32 * std::f32::consts::FRAC_PI_2;
        let frame = render_loading_frame(side as u32, side as u32, angle);
        for (index, pixel) in frame.as_chunks::<4>().0.iter().enumerate() {
            if pixel[3] == 0 {
                continue;
            }
            let (x, y) = (index % side, index / side);
            min_x = min_x.min(x);
            min_y = min_y.min(y);
            max_x = max_x.max(x);
            max_y = max_y.max(y);
        }
    }

    // Nothing but the fade the halo ends in is left between the spinner and
    // the frame it is placed in.
    assert!(
        min_x <= 2 && min_y <= 2,
        "the spinner reaches its frame's near edges, {min_x} and {min_y} of {side}"
    );
    assert!(
        max_x >= side - 3 && max_y >= side - 3,
        "and its far ones, {max_x} and {max_y} of {side}"
    );
}

/// A load that may be about to finish is given the delay `spinner_delay_ms` names
/// before the spinner goes up — the same one for every kind of wait — while a delay
/// of nothing is a spinner that goes up with the load.
#[test]
fn puts_the_spinner_up_once_the_load_has_run_for_the_delay() {
    let load = |age: Duration, delay: Duration, upgrade: bool| PendingLoad {
        generation: 1,
        hide_epoch: 0,
        path: PathBuf::new(),
        started: Instant::now() - age,
        pos_x: 0,
        pos_y: 0,
        width: 64,
        height: 64,
        room: (1920, 1040),
        spinner_shown: false,
        spinner_delay: delay,
        spinner_pos: (0, 0),
        spinner_side: office_preview::WAITING_BOX,
        placement: None,
        upgrade,
        awaiting_engine: false,
    };
    let default_delay = Duration::from_millis(DEFAULT_SPINNER_DELAY_MS);

    // A load that may be about to finish: not yet, and due once it has run for
    // the delay.
    assert!(!load(Duration::from_millis(20), default_delay, false).spinner_due());
    assert!(load(default_delay, default_delay, false).spinner_due());

    // A delay of nothing puts the spinner up with the load, and a load given a
    // longer one than it has run for is not due yet.
    assert!(load(Duration::ZERO, Duration::ZERO, false).spinner_due());
    assert!(!load(Duration::from_secs(1), Duration::from_secs(5), false).spinner_due());

    // An upgrade never: what is on screen stays until the page replaces it.
    assert!(!load(Duration::from_secs(10), default_delay, true).spinner_due());

    // And a spinner that is already up is not put up a second time.
    let mut showing = load(Duration::from_secs(10), default_delay, false);
    showing.spinner_shown = true;
    assert!(!showing.spinner_due());
}

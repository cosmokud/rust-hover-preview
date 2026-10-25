use super::*;

/// **A step that grows is a step whose window has to be down before its film is**, and one that
/// shrinks is the other way round (see `apply_pin_drag`). The direction is read off the box the step
/// came from against the one it goes to, on either axis.
#[test]
fn a_resize_step_grows_where_either_side_of_it_is_pulled_out() {
    let at = (100, 100, 500, 400);

    assert!(
        a_resize_step_grows(at, (100, 100, 600, 400)),
        "a box pulled wider grew"
    );
    assert!(
        a_resize_step_grows(at, (100, 100, 500, 500)),
        "and so did one pulled taller"
    );
    assert!(
        a_resize_step_grows(at, (200, 100, 700, 400)),
        "and a box pulled wider and shorter at once, because the side it exposes is the one the \
         order is about — the two boxes cross rather than nest there, so no order keeps the film \
         inside and the side that grew is the one the order can still get right"
    );
    assert!(
        !a_resize_step_grows(at, (100, 100, 400, 300)),
        "a box smaller on both sides shrank"
    );
    assert!(
        !a_resize_step_grows(at, at),
        "and a step that changed neither side is no growth, so it takes the shrinking order: the two \
         commits then write the same box whichever goes first"
    );
}

#[test]
fn a_resize_keeps_the_shape_of_what_is_pinned() {
    // A media box 400 x 300, dragged by its right edge 200 to the right: the media grows by
    // half in both directions rather than being stretched sideways, the edge that was not
    // dragged stays where it was, and the window grows around the line it was on — which is
    // what an image viewer does with a picture it is being made wider.
    let content = dragged(edge(false, false, true, false), 200, 0);
    assert_eq!(content, (400, 225, 1000, 675));
    assert_eq!(
        (content.2 - content.0) as f64 / (content.3 - content.1) as f64,
        400.0 / 300.0,
        "the media keeps its shape"
    );

    // And the same drag from the left edge grows it away to the left, with the right edge it
    // did not drag the one that stays.
    assert_eq!(
        dragged(edge(true, false, false, false), -200, 0),
        (200, 225, 800, 675)
    );

    // A vertical edge keeps the horizontal line it was on for the same reason.
    assert_eq!(
        dragged(edge(false, false, false, true), 0, 150),
        (300, 300, 900, 750)
    );
    assert_eq!(
        dragged(edge(false, true, false, false), 0, -150),
        (300, 150, 900, 600)
    );

    // A corner asks for both sides at once, and what the shape allows is one of the two
    // answers: the nearer of them to what the hand asked for is the one the box takes.
    assert_eq!(
        dragged(edge(false, false, true, true), 200, 30),
        (400, 300, 1000, 750)
    );
    // And the corner opposite the one being dragged is the one that stays.
    assert_eq!(
        dragged(edge(true, true, false, false), 100, 75),
        (500, 375, 800, 600),
        "the bottom-right corner of the box did not move"
    );
}

/// A corner dragged against the room the media is kept inside moves the box with the hand the whole
/// way, a pixel at a time, and stops at the room's edge rather than jumping to it.
///
/// This is the shape of the bug it was written for: the two answers a shape has for a corner — one
/// worked out from the width the hand asked for, one from the height — were compared *after* each had
/// been cut to the room, so the nearer of the two could be the box at the room's edge one step and a
/// box a hundred pixels inside it the next, with the hand moving a pixel. A corner pulled along the
/// top of the screen therefore leapt back and forth between those two boxes for as long as the drag
/// lasted, which is the snapping a corner does at the edge of a display and nowhere else.
#[test]
fn a_corner_dragged_against_the_room_stops_at_its_edge_and_never_jumps_to_it() {
    // The pin of that bug, as the log of it had it: a media box 516 by 730 whose top right corner is
    // pulled up and to the right until the room's own top and right are reached, which for this
    // shape is the same box from both sides.
    let press = (1092, 323, 1608, 1053);
    let aspect = (press.2 - press.0) as f64 / (press.3 - press.1) as f64;
    let space = |content| PinSpace {
        content,
        room: ScreenBounds {
            left: 0,
            top: 0,
            right: 1836,
            bottom: 1400,
        },
        overlay: true,
        transport: false,
        caption: 0,
    };

    let mut previous = press;
    for travel in 1..=400 {
        // One pixel of hand on each edge, every step: the top going up and the right going right.
        let (dx, dy) = (travel, -travel);
        let content = resize_pinned_content(
            space(press),
            edge(false, true, true, false),
            dx,
            dy,
            96,
            PinFrame::Shaped,
        );

        assert!(
            content.0 >= 0 && content.1 >= 0 && content.2 <= 1836 && content.3 <= 1400,
            "the box {content:?} left the room the drag is kept inside"
        );
        assert!(
            ((content.2 - content.0) as f64 / (content.3 - content.1) as f64 - aspect).abs() < 0.02,
            "the box {content:?} is not the shape the media is"
        );

        let (was_width, was_height) = (previous.2 - previous.0, previous.3 - previous.1);
        let (width, height) = (content.2 - content.0, content.3 - content.1);
        assert!(
            (width - was_width).abs() <= 2 && (height - was_height).abs() <= 2,
            "a pixel of hand took the box from {previous:?} to {content:?}"
        );

        previous = content;
    }

    // And it does reach the room, which is the whole of what the drag was for.
    assert_eq!(
        (previous.1, previous.2),
        (0, 1836),
        "the corner stopped short of the room's own edge: {previous:?}"
    );
}

/// A pin shown another file keeps the box it has: what the new file is given is the largest box
/// of its own shape that the room can hold at the scale it is given, in the middle of it. So the
/// window a swap comes out of stands where the window was and is no larger than the room,
/// whatever the two files' shapes are, and what is drawn in it is the shape of the file that
/// replaces the pin's rather than the shape of the file it replaces.
#[test]
fn a_swap_fits_the_new_shape_into_the_box_the_pin_has() {
    // A media box 800 by 500, with a window around it that is nobody's business here.
    let room = (100, 200, 900, 700);

    for shape in [
        (1920u32, 1080u32),
        (1080, 1920),
        (100, 100),
        (4000, 400),
        (1, 1),
        (500, 800),
        (800, 500),
    ] {
        let content = pin_update_box(room, shape, PreviewScale::FitToScreen);
        let width = (content.2 - content.0).max(1);
        let height = (content.3 - content.1).max(1);
        let shape_ratio = shape.0 as f64 / shape.1 as f64;

        assert!(
            content.0 >= room.0
                && content.1 >= room.1
                && content.2 <= room.2
                && content.3 <= room.3,
            "a {shape:?} file was given a box outside the pin's own: {content:?}"
        );
        assert!(
            (width as f64 / height as f64 - shape_ratio).abs() < 0.02 * shape_ratio,
            "the box {content:?} does not have the shape of the file ({shape:?})"
        );
        assert!(
            (content.0 - room.0 - (room.2 - content.2)).abs() <= 1
                && (content.1 - room.1 - (room.3 - content.3)).abs() <= 1,
            "the box {content:?} is not in the middle of the pin's own"
        );
    }
}

/// The room a swap is fitted into is the bound the pin carries rather than the box the file
/// before it left, so a pin that is shown file after file keeps the size it went up with: what
/// fitting one shape into the box another came out of does is take a side off it, and a run of
/// files whose proportions disagree would walk the window down to nothing one file at a time.
///
/// And the bound is a side rather than a box, so a shape the pin was never given still gets the
/// whole of it: the tall file below is given 800 pixels of height — more than the box the pin
/// went up with had — because 800 is the side the user's own hand accepted.
#[test]
fn a_swap_is_fitted_into_the_bound_rather_than_the_box_the_last_file_left() {
    // A pin that went up at 800 by 500, whose longest side is now what every file may use, and
    // the box it stands in now: a swap may take the window down from the bound and never past
    // it.
    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };
    let bound = 800;
    let mut current = (100, 200, 900, 700);
    let centre = ((current.0 + current.2) / 2, (current.1 + current.3) / 2);

    // The shapes a listing holds, and the one setting that fills the bound whatever the file's
    // own pixels are — the scale that shrinks fastest when the room is the last file's box.
    let shapes = [(1920u32, 1080u32), (1080, 1920), (4000, 400), (100, 100)];

    // Twice round the same files: what the pin is given for one is the same answer in both
    // passes, which is the whole of what the old box being the room had taken away.
    let mut passes: Vec<Vec<(i32, i32)>> = Vec::new();
    for _ in 0..2 {
        let mut sizes = Vec::new();

        for shape in shapes {
            let space = PinSwapSpace {
                current,
                bound: Some(bound),
                transport_bar: false,
                overlay: true,
                caption: pinned_caption_height(96, None),
                maximized: false,
                room: bounds,
            };
            current = pin_update_box(
                pin_swap_room(space, bounds, 96),
                shape,
                PreviewScale::FitToScreen,
            );

            let (width, height) = (current.2 - current.0, current.3 - current.1);
            let middle = ((current.0 + current.2) / 2, (current.1 + current.3) / 2);
            assert!(
                width <= bound && height <= bound,
                "a {shape:?} file was given a box larger than the bound: {current:?}"
            );
            assert!(
                (middle.0 - centre.0).abs() <= 1 && (middle.1 - centre.1).abs() <= 1,
                "the box {current:?} has moved from where the window stands"
            );

            sizes.push((width, height));
        }

        passes.push(sizes);
    }

    assert_eq!(
        passes[0], passes[1],
        "a file was given a different box the second time it was shown"
    );
}

/// A maximized pin walks its files against the display's room, and not against the box the
/// file before each one left.
///
/// This is the same rule the bound keeps above, in the one place it was not kept: a
/// maximized window is the one that fills its display, so every file it walks to is fitted
/// to the room afresh. Measured against the last file's box instead, each step is fitted
/// into a box the step before shrank — 16:9 into 16:9, then 4:3 into that, then 1:1 into
/// that — so a window walked far enough shrinks to the smallest shape in the folder and
/// keeps shrinking, while the caption goes on drawing the restore glyph the whole way,
/// because the maximize is a state about the room and was never given up.
#[test]
fn a_maximized_pin_walks_its_files_against_the_room_and_not_the_last_box() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let folder = std::env::temp_dir().join("rust-hover-preview-pin-maximized");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };

    // The room a maximized window's media is laid out in: the display's own, less the
    // caption above it (see `pinned_room`).
    let room = pinned_room(bounds, 96, false, false, pinned_caption_height(96, None));

    // Five files of real shapes, walked in an order that makes the ratchet visible: a wide
    // one, a tall one, a square, a strip, and the wide one again. What a user hammering
    // next through a mixed folder actually does.
    let shapes = [
        (1920u32, 1080u32),
        (1080, 1920),
        (1000, 1000),
        (4000, 400),
        (1920, 1080),
    ];

    // A file per shape, each a real picture of that size, so that the measure is an answer
    // and the swap a laying out rather than a wait (see `pin_swap_awaits`).
    let files: Vec<PathBuf> = shapes
        .iter()
        .map(|(width, height)| {
            let path = folder.join(format!("{width}x{height}.png"));
            image::save_buffer(
                &path,
                &vec![0x40u8; (*width * *height * 3) as usize],
                *width,
                *height,
                image::ExtendedColorType::Rgb8,
            )
            .expect("a written picture");
            // The picture's own box, held where the measure would leave it: what a swap lays
            // out is the answer to a measure read off the preview thread, and a test has no
            // thread to wait for (see `picture_box`).
            hold_box(
                &path,
                &file_version(&path),
                &MeasureScope::File,
                Some((*width, *height)),
            );
            path
        })
        .collect();

    // Walk the folder, carrying the box each file was laid out in as the box the next one
    // is swapped into. That carrying is the whole of the bug, and it is why the walk has to
    // be continuous rather than restarted: a maximized swap that measures the new file
    // against the box the last one left makes every step a smaller step, so the window
    // shrinks a little on every keypress and never comes back. A walk restarted from the
    // room each time would ratchet identically on every lap and look perfectly stable,
    // which is why the box is carried through rather than reset.
    //
    // Three laps is enough to see it, and not so many that the test is slow: each step is
    // one fit of one shape.
    let mut current = room_bounds_as_region(room);
    let mut laps: Vec<Vec<(i32, i32)>> = Vec::new();

    for _ in 0..3 {
        let mut sizes = Vec::new();

        for path in &files {
            let space = PinSwapSpace {
                current,
                bound: None,
                transport_bar: false,
                overlay: false,
                caption: pinned_caption_height(96, None),
                maximized: true,
                // The room the production reader builds, caption and all, and not the bare
                // display: a test that laid its files out against a room nothing lays out
                // against is a test of its own fixture.
                room: pinned_room(bounds, 96, false, false, pinned_caption_height(96, None)),
            };
            let Some(PinBox::Measured(laid_out)) = pin_update_content(space, path, bounds, 96)
            else {
                panic!("a measured picture was not laid out");
            };

            current = laid_out;
            sizes.push((current.2 - current.0, current.3 - current.1));
        }

        laps.push(sizes);
    }

    // Every lap of the same five files is the same as every other. A window measured
    // against the last file's box cannot be: each lap comes back smaller than the one
    // before, which is the window shrinking away under a caption still drawing the restore
    // glyph — because nothing ever gave the maximize up, so nothing said it had stopped
    // being one.
    assert_eq!(
        laps[0], laps[1],
        "a maximized window was given a different box the second time a file was shown"
    );
    assert_eq!(
        laps[1], laps[2],
        "a maximized window kept shrinking over a third lap of the same files"
    );

    // And a file that fills the room still fills it, so "maximized" is a window of the
    // room's size and not a name for no larger a box than before.
    //
    // The room is 1920 by 1050 — the display's own 1080 less the caption. A 16:9 file
    // fitted to it fills the height and is 1867 wide, since 16:9 of 1050 is 1866.67: the
    // picture's own shape at the room's own height, which is the largest this display has
    // to give it. Under the ratchet the same file came back 105 by 59.
    assert_eq!(
        laps[0][0],
        (1867, 1050),
        "a maximized pin's own 16:9 file no longer fills the room"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// A maximized window walks onto a file smaller than the room, and the file
/// is drawn as large as the display can show it.
///
/// The maximize itself fits the file to the room — up as well as down (see
/// `pinned_media_box`) — and the walk a maximized window takes is laid out by
/// the same rule, because the maximize is a state about the room and never
/// stopped being one. A swap that kept the percentage the hover scale names
/// instead would draw a small file at its own size while the caption went on
/// drawing the restore glyph: the button would say maximized and the preview
/// would not be, which is the one thing a maximize is not.
#[test]
fn a_maximized_pin_fits_a_file_smaller_than_the_room_to_the_room() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let folder = std::env::temp_dir().join("rust-hover-preview-pin-maximized-small");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };

    // The room a maximized window's media is laid out in: the display's own,
    // less the caption above it (see `pinned_room`).
    let room = pinned_room(bounds, 96, false, false, pinned_caption_height(96, None));

    // A picture smaller than the room in both directions — the one shape a
    // percentage of its own size and a fit to the room disagree about, and so
    // the one a wrong scale is shown by.
    let path = folder.join("400x300.png");
    image::save_buffer(
        &path,
        &vec![0x40u8; (400 * 300 * 3) as usize],
        400,
        300,
        image::ExtendedColorType::Rgb8,
    )
    .expect("a written picture");

    // And the box the picture is measured at, held where the measure would leave it (see
    // `picture_box`).
    hold_box(
        &path,
        &file_version(&path),
        &MeasureScope::File,
        Some((400, 300)),
    );

    let space = PinSwapSpace {
        current: room_bounds_as_region(room),
        bound: None,
        transport_bar: false,
        overlay: false,
        caption: pinned_caption_height(96, None),
        maximized: true,
        room,
    };

    let Some(PinBox::Measured(laid_out)) = pin_update_content(space, &path, bounds, 96) else {
        panic!("a measured picture was not laid out");
    };

    // The room is 1920 by 1050. A 4:3 picture fitted to it is height-bound —
    // 1050 of height, and 4:3 of that is 1400 of width — which is the box the
    // maximize the window is in gives it, and not the 400 by 300 a percentage
    // of its own size would.
    assert_eq!(
        (laid_out.2 - laid_out.0, laid_out.3 - laid_out.1),
        (1400, 1050),
        "a maximized window's small file is fitted to the room, up as well as down"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// A double-click that asks for the window's size back is the whole of what a fresh take-up of the
/// file would have been given: the file measured as its kind is measured, fitted into the room the
/// pin's own chrome leaves at the share the scaling setting names for that kind, and put in the
/// middle of that room — with the size a hand pulled the window to given up and a maximize with it.
///
/// The size comes from the *setting* rather than from the box the window is standing in, and both
/// ways it can be wrong are asked of: a window pulled down to 200 by 150 comes back 400 by 300 at
/// the picture's own `100%`, and one pulled out to 800 by 600 comes back down to 200 by 150 at
/// `50%`.
#[test]
fn a_window_the_hand_sized_is_put_back_to_the_size_the_setting_asks_for() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let folder = std::env::temp_dir().join("rust-hover-preview-pin-reset-size");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let path = folder.join("400x300.png");
    image::save_buffer(
        &path,
        &vec![0x40u8; (400 * 300 * 3) as usize],
        400,
        300,
        image::ExtendedColorType::Rgb8,
    )
    .expect("a written picture");

    // And the picture's own box, held where the measure read off the preview thread would leave
    // it (see `picture_box`).
    hold_box(
        &path,
        &file_version(&path),
        &MeasureScope::File,
        Some((400, 300)),
    );

    // A pin showing that picture, with its window pulled by hand to a size of its own: a box in the
    // corner of the screen it stands on, the bound such a drag leaves behind written down with it
    // (see `overlay_pin`), and a maximize standing that the reset has to give up.
    let room = work_area_at(0, 0);
    let pulled_small = (room.left, room.top, room.left + 200, room.top + 150);
    let pulled_wide = (room.left, room.top, room.left + 800, room.top + 600);

    // The media behind the window, so that the tick this asks is a tick with a pin still standing:
    // a pin whose media reads as gone is taken down rather than given a box (see
    // `pin_media_is_alive`).
    let mut media = create_loading_media(400, 300);
    media.media_type = MediaType::StaticImage;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }

    // One reset through the loop's own drain, which is where a double-click's own reset is answered:
    // the pair itself is what asks for it, and a press is a procedure this test has no window for
    // (see `press_carries_the_pin`).
    let reset = || {
        let mut request = None;
        let mut wait = None;
        let mut started = None;
        let mut offset = 0.0;
        let mut paused = None;
        ask_pin(PinCommand::ResetSize);
        unsafe {
            pin_command_request(
                HWND(0x1000 as *mut _),
                &mut request,
                &mut wait,
                &mut started,
                &mut offset,
                &mut paused,
                false,
                None,
            );
        }
        request
    };

    for (setting, content, size) in [
        (PreviewScale::Percent(100), pulled_small, (400, 300)),
        (PreviewScale::Percent(50), pulled_wide, (200, 150)),
    ] {
        let _picture_scale = PictureScale::stood_at(setting);

        let mut pin = overlay_pin(content, PinChrome::always());
        pin.path = path.clone();
        pin.restore = Some((0, 0, 800, 600));
        stand_pin(Some(pin));

        let Some(PreviewMessage::PinBox(reset_box)) = reset() else {
            panic!("a reset at {setting:?} asked for no box");
        };

        assert_eq!(
            (reset_box.2 - reset_box.0, reset_box.3 - reset_box.1),
            size,
            "the window is not the size a picture of 400 by 300 is given at {setting:?}"
        );

        // The room the pin leaves its media is the screen's own work area, this pin's chrome being
        // drawn over its media: what the box is put in the middle of is asked as the rule rather
        // than as a number, the display being the machine's own.
        let (left, right) = (reset_box.0 - room.left, room.right - reset_box.2);
        assert!(
            (left - right).abs() <= 1,
            "the window is not in the middle of the room it stands on: {reset_box:?} in {room:?}"
        );

        let (bound, restore, written) = pin_state()
            .and_then(|pinned| {
                let pin = pinned.pin()?;
                Some((pin.bound, pin.restore, pin.content))
            })
            .expect("the pin this test stood");
        assert_eq!(
            written, reset_box,
            "the box asked for is not the box the pin was written with"
        );
        assert_eq!(
            bound, None,
            "the size a hand pulled the window to was left behind as the bound the files after it \
             are fitted into"
        );
        assert_eq!(
            restore, None,
            "a reset left the maximize it found standing, so the caption goes on drawing a restore"
        );
    }

    let _ = std::fs::remove_dir_all(&folder);
    stand_pin(previous_pin);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
}

/// A sound's card is the one kind a pair hands no size back to: a card is framed by nothing — no
/// hand can pull one to another size — so the size it has is the size its setting asked for, and
/// what a double-click owes it is the middle. The measure goes with it, so a card is never measured
/// again at a scale the pin was not taken up at (see `reset_pinned_size`).
#[test]
fn a_card_is_given_the_middle_and_keeps_the_size_it_is_drawn_at() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    // A card at a size of its own, in the corner of the screen it stands on, showing a sound no
    // machine has ever seen: a measure taken for this file would be a wait rather than a size, and
    // the whole of what is asked here is that none is taken at all.
    let room = work_area_at(0, 0);
    let content = (room.left, room.top, room.left + 320, room.top + 96);
    let mut pin = PinnedPreview {
        frame: PinFrame::None,
        overlay: false,
        content,
        caption: 0,
        bound: None,
        ..PinnedPreview::for_test()
    };
    pin.path = std::env::temp_dir().join("a-card-no-machine-has.flac");
    stand_pin(Some(pin));

    let mut media = create_loading_media(320, 96);
    media.media_type = MediaType::Audio;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }

    let mut request = None;
    let mut wait = None;
    let mut started = None;
    let mut offset = 0.0;
    let mut paused = None;
    ask_pin(PinCommand::ResetSize);
    unsafe {
        pin_command_request(
            HWND(0x1000 as *mut _),
            &mut request,
            &mut wait,
            &mut started,
            &mut offset,
            &mut paused,
            false,
            None,
        );
    }

    let Some(PreviewMessage::PinBox(reset_box)) = request else {
        panic!("a reset of a card asked for no box");
    };
    assert_eq!(
        (reset_box.2 - reset_box.0, reset_box.3 - reset_box.1),
        (320, 96),
        "a card was given a size of its own rather than the one it is drawn at"
    );

    let (left, right) = (reset_box.0 - room.left, room.right - reset_box.2);
    assert!(
        (left - right).abs() <= 1,
        "a card was not put in the middle of the room it stands on: {reset_box:?} in {room:?}"
    );

    stand_pin(previous_pin);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
}

fn room_bounds_as_region(room: ScreenBounds) -> ScreenRegion {
    (room.left, room.top, room.right, room.bottom)
}

/// The maximize button is a toggle, and both of its presses change the pin.
///
/// A window with no restore put aside is maximized to the room and remembers the box it had;
/// a window with one is put back to it and gives it up. Which of the two it is has to be
/// read from the restore that came *in* — the two answers are never the same value, so a
/// button that checks its answer against the state it is writing silently does nothing at
/// all, in both directions, forever. That is what a check written against the outgoing
/// value does, and it is invisible from the tests below unless the decision itself is
/// asked directly.
#[test]
fn a_maximize_button_answers_in_both_directions() {
    let room = ScreenBounds {
        left: 0,
        top: 30,
        right: 1920,
        bottom: 1080,
    };
    let standing_in = (700, 400, 1700, 1000);
    let wide = (1920u32, 1080u32);

    // The first press: nothing is maximized, so the file is fitted to the room and the box
    // it stood in is put aside for the way back.
    let (maximized, remembered) = pin_maximize_decided(PinMaximize {
        frame: PinFrame::Shaped,
        content: standing_in,
        restore: None,
        shape: Some(wide),
        room,
    });

    assert_eq!(
        remembered,
        Some(standing_in),
        "a maximize did not put the box it replaced aside"
    );
    let (width, height) = (maximized.2 - maximized.0, maximized.3 - maximized.1);
    assert_eq!(
        (width, height),
        (1867, 1050),
        "a maximize is not the file fitted to the room"
    );

    // The second press, on the state the first left: the box comes back and the maximize is
    // given up, so a third press maximizes again rather than restoring a second time.
    let (restored, given_up) = pin_maximize_decided(PinMaximize {
        frame: PinFrame::Shaped,
        content: maximized,
        restore: remembered,
        shape: Some(wide),
        room,
    });

    assert_eq!(given_up, None, "a restore left a maximize behind");

    // The whole of the button in one property: what a press decides is never the state it
    // was handed. A maximize remembers a box and a restore gives it up, so the two can
    // never agree — which is what makes it safe to guard the write on the state that was
    // read, and unsafe to guard it on the value being written, since that could never
    // match. A button that did the latter does nothing at all, in both directions, and
    // nothing else about the pin would say so.
    for (restore_in, restore_out) in [(None, remembered), (remembered, given_up)] {
        assert_ne!(
            restore_in, restore_out,
            "a press of the maximize button left the pin exactly as it found it"
        );
    }

    // The box that comes back is the box that file wants rather than the one the maximize
    // filled the room with — a window put back into the maximized box would be a 16:9
    // picture in a 1867-wide box it can never grow out of, which is the stretch the button
    // was fixed for. And it is the same box every time, so a restore is not a reshuffle.
    let (restored_again, _) = pin_maximize_decided(PinMaximize {
        frame: PinFrame::Shaped,
        content: restored,
        restore: Some(standing_in),
        shape: Some(wide),
        room,
    });
    assert_eq!(
        restored_again, restored,
        "a restore of the same file gave it two different boxes"
    );
    assert_ne!(
        restored, maximized,
        "a restore is the maximized box, which is the stretch the button was fixed for"
    );
    let (width, height) = (restored.2 - restored.0, restored.3 - restored.1);
    assert!(
        width <= standing_in.2 - standing_in.0 && height <= standing_in.3 - standing_in.1,
        "a restore grew the window past the box the user had: {restored:?}"
    );

    // And the state the first press wrote is what the second one reads: a toggle that took
    // its answer from anywhere else would get this wrong.
    let (again, _) = pin_maximize_decided(PinMaximize {
        frame: PinFrame::Shaped,
        content: restored,
        restore: given_up,
        shape: Some(wide),
        room,
    });
    assert_eq!(again, maximized, "the button is not a toggle");

    // A kind laid out to whatever box it is given has no shape to fit, and the room is the
    // whole of what maximizing means for it.
    let (free, remembered) = pin_maximize_decided(PinMaximize {
        frame: PinFrame::Free,
        content: standing_in,
        restore: None,
        shape: Some(wide),
        room,
    });
    assert_eq!(
        (free.2 - free.0, free.3 - free.1),
        (1920, 1050),
        "a page of text is not maximized to the room"
    );
    assert_eq!(remembered, Some(standing_in));
}

/// A restore down gives the file on screen a box of *its own* shape, at the size the user
/// had the window at.
///
/// The box put aside belongs to whatever was on screen when the maximize was pressed. A
/// window that has since walked along the pin to a file of another shape and then restores
/// down is put into a box sized and shaped for the old file — which is a stretched window,
/// and a walk through shapes is the surest way to arrive at one. The size stays the user's,
/// because a restore undoes a maximize and does not resize; only the shape is the file's.
#[test]
fn a_restore_down_gives_the_file_on_screen_a_box_of_its_own_shape() {
    let room = ScreenBounds {
        left: 0,
        top: 30,
        right: 1920,
        bottom: 1080,
    };

    // The box a 16:9 window was left at before it was maximized, and the file that was in
    // it: its own shape, fitted to that box, which is where it was before the maximize and
    // is where a restore puts it back.
    let chosen = (700, 400, 1700, 1000);
    let before = (1920u32, 1080u32);

    // The file the walk has since landed on. Its shape is nothing like the box's, and the
    // aspect of what comes back has to follow the file rather than the box.
    let after = (1080u32, 1920u32);

    // A restore of the file the box belongs to is that box again, to the pixel: a restore
    // undoes a maximize and does not resize, so the box the user had is the answer.
    let fitted_before = pin_restored_box(before, chosen, room);
    assert_eq!(
        pin_restored_box(before, fitted_before, room),
        fitted_before,
        "a restore of the file a box was fitted to moved it"
    );

    let restored = pin_restored_box(after, chosen, room);
    let (width, height) = (restored.2 - restored.0, restored.3 - restored.1);

    // The aspect is the file's, which is the whole of what a stretched window is not. Had
    // the box been put back as it stood, this would be 1000 by 600 — a 1.67 ratio for a
    // 0.56 file, drawn 3× too wide.
    assert_eq!(
        (width as f32 / height as f32).round(),
        (after.0 as f32 / after.1 as f32).round(),
        "a tall file was restored into a wide box and would be drawn stretched"
    );

    // And the box is never larger than the one the user chose: the file is fitted into
    // their size rather than replacing it, so a restore takes nothing away.
    assert!(
        width <= chosen.2 - chosen.0 && height <= chosen.3 - chosen.1,
        "a restore grew the window past the size the user had: {restored:?}"
    );

    // A file whose shape *is* the box's shape comes back at that size, in the middle of the
    // room, which is what makes the rule above a fit and not a guess at a smaller box.
    let square = (600u32, 600u32);
    assert_eq!(
        pin_restored_box(square, (0, 0, 600, 600), room),
        (660, 255, 1260, 855),
        "a square file in a square box comes back at that size"
    );
}

/// A pin taken up as a portrait leaves the whole of its longest side for the files that follow
/// it: the widescreen file below is given 2000 pixels of width, where the box the pin went up in
/// would have given it 1000 and shown it at a third of the size (see `PinnedPreview::bound`).
#[test]
fn a_tall_pin_leaves_its_side_for_the_widescreen_file_that_follows_it() {
    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 3840,
        bottom: 2160,
    };
    // A pin that went up as a 1000 by 2000 portrait, whose longest side is 2000.
    let space = PinSwapSpace {
        current: (1000, 100, 2000, 2100),
        bound: Some(2000),
        transport_bar: false,
        overlay: true,
        caption: pinned_caption_height(96, None),
        maximized: false,
        room: bounds,
    };
    let room = pin_swap_room(space, bounds, 96);
    assert_eq!(room, (500, 100, 2500, 2100));

    // The 16:9 file that follows uses all of it: 2000 by 1125, in the middle of the room.
    assert_eq!(
        pin_update_box(room, (1920, 1080), PreviewScale::FitToScreen),
        (500, 538, 2500, 1663)
    );

    // And a portrait that follows a landscape gets the same answer from its own side of it.
    assert_eq!(
        pin_update_box(room, (1080, 1920), PreviewScale::FitToScreen),
        (938, 100, 2063, 2100)
    );
}

/// The scale the tray names decides the box inside the bound: a percentage takes the file's own
/// size, reduced only where the bound cannot hold it, and fit-to-screen fills the bound's side
/// — and the bound is the ceiling for both, so a file can be shown smaller than it, and never
/// larger.
#[test]
fn a_swap_takes_the_scale_setting_and_the_bound_is_its_ceiling() {
    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };
    let space = PinSwapSpace {
        current: (100, 200, 900, 700),
        bound: Some(800),
        transport_bar: false,
        overlay: true,
        caption: pinned_caption_height(96, None),
        maximized: false,
        room: bounds,
    };
    let room = pin_swap_room(space, bounds, 96);

    // The room is the bound in both directions, in the middle of the box the pin stands in: a
    // file of any shape may use the whole of the side the user gave the pin.
    assert_eq!(room, (100, 50, 900, 850));

    // A small file at a percentage is its own size, in the middle of that room: what the setting
    // asks for is the file's pixels, not the room's.
    assert_eq!(
        pin_update_box(room, (400, 250), PreviewScale::Percent(100)),
        (300, 325, 700, 575)
    );

    // The same file at fit-to-screen fills the bound's side, its own shape deciding the other.
    assert_eq!(
        pin_update_box(room, (400, 250), PreviewScale::FitToScreen),
        (100, 200, 900, 700)
    );

    // And a file larger than the bound, at the same percentage, is reduced to it rather than
    // drawn at its own pixels.
    assert_eq!(
        pin_update_box(room, (2000, 1000), PreviewScale::Percent(100)),
        (100, 250, 900, 650)
    );
}

/// A bound outlives the display it was measured on: a window carried to a smaller display, or
/// one the desktop was rearranged under, is fitted into the room that is there rather than into
/// the room the bound was given, so a swap cannot put a window back up that the display has no
/// room to show.
#[test]
fn a_bound_larger_than_the_display_is_capped_by_the_room_it_is_on() {
    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 1280,
        bottom: 800,
    };
    let space = PinSwapSpace {
        current: (0, 0, 2000, 3000),
        bound: Some(3000),
        transport_bar: true,
        overlay: false,
        caption: pinned_caption_height(96, None),
        maximized: false,
        room: bounds,
    };

    // The room of a kind with bands is the work area less the caption and the transport bar,
    // and what stands in the middle of the box the window has now is the bound cut to it — a
    // side longer than the display's own room is bounded by the room.
    let room = pin_swap_room(space, bounds, 96);
    assert_eq!((room.2 - room.0, room.3 - room.1), (1280, 800 - 30 - 30));

    // And what is fitted into it stays inside it: a tall file the size of the bound is taken
    // down to the display rather than standing 3000 pixels tall on an 800 pixel screen.
    let content = pin_update_box(room, (2000, 3000), PreviewScale::Percent(100));
    assert!(
        content.2 - content.0 <= 1280 && content.3 - content.1 <= 800 - 30 - 30,
        "a box larger than the room was given back: {content:?}"
    );
}

/// A pin that went up on a file drawn to its own box — a page of text, a listing, a sound's card
/// — has no bound yet, and the file that follows it is laid out against the whole of the room
/// the display has rather than inside a square of that box: a box the file was *drawn* to is no
/// size to fit anything into, and it is the scale its kind names that decides what comes of it
/// (see `pin_swap_room` and `PinnedPreview::bound`).
#[test]
fn a_pin_with_no_bound_is_fitted_into_the_room_the_display_has() {
    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };

    // A kind whose chrome is drawn over its media: the room is the whole display, wherever the
    // small window the pin went up in stands.
    let space = PinSwapSpace {
        current: (760, 440, 1160, 640),
        bound: None,
        transport_bar: false,
        overlay: true,
        caption: pinned_caption_height(96, None),
        maximized: false,
        room: bounds,
    };
    let room = pin_swap_room(space, bounds, 96);
    assert_eq!(room, (0, 0, 1920, 1080));

    // A 16:9 file at fit-to-screen fills it — where the box the card left would have shown it
    // at a fraction of the size.
    assert_eq!(
        pin_update_box(room, (1920, 1080), PreviewScale::FitToScreen),
        (0, 0, 1920, 1080)
    );

    // And a percentage is the file's own pixels, inside that room.
    assert_eq!(
        pin_update_box(room, (400, 250), PreviewScale::Percent(400)),
        (160, 40, 1760, 1040)
    );

    // A kind with bands is given the room those leave: the display's work area less the caption
    // and the transport bar.
    let banded = PinSwapSpace {
        current: (760, 440, 1160, 640),
        bound: None,
        transport_bar: true,
        overlay: false,
        caption: pinned_caption_height(96, None),
        maximized: false,
        room: bounds,
    };
    assert_eq!(pin_swap_room(banded, bounds, 96), (0, 30, 1920, 1050));
}

/// A box change that the pin's own frame already covers owes it no decode, and the kinds a box lays
/// out again rather than scales are owed one whatever the frame's size is. A page is the first of
/// those as much as a picture is: what the box decides about a PDF, a slide or a book's page is how
/// many pixels the page is drawn at, not what is on it.
///
/// This is the rule the window a user drags small and opens again lives or dies by. The band is the
/// pin's own box and the frame in the media is stretched into it, so a frame decoded for a bigger
/// box is every pixel the smaller box would have been given: decoding the smaller size replaces it,
/// and a window shut small and opened again is then soft from a thumbnail until a whole decode of
/// the file lands. What the frame answers is asked of the frame rather than of the cache, so it holds
/// however many hovers have gone through the cache since (see `box_is_covered`).
///
/// A page of text is the other side of it, and the side that keeps the rule honest: text is drawn at
/// a fixed size, so half the box is half the lines rather than the same page smaller, and a frame of
/// it stretched into a box it was not laid out for is a different pageful rather than a soft one.
#[test]
fn a_box_the_frames_pixels_already_cover_owes_the_pin_no_decode() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let held = |kind: MediaType, width: u32, height: u32| {
        if let Ok(mut media) = CURRENT_MEDIA.lock() {
            *media = Some(static_image_media(
                Arc::new(ImageFrame::new(
                    vec![0u8; (width * height * 4) as usize],
                    width,
                    height,
                    0,
                )),
                kind,
            ));
        }
    };
    let path = PathBuf::from("a-pin-frame.png");
    let owed = |path: &Path, box_: ScreenRegion, road| {
        relayout_pinned_media(path, box_, 96, None, road).is_some()
    };

    // A picture whose frame was decoded for a box the size of a display, dragged small.
    held(MediaType::StaticImage, 1600, 1200);
    assert!(
        !owed(&path, (100, 100, 500, 400), PinRelayoutRoad::BoxChange),
        "a 400x300 box is answered by the 1600x1200 frame the pin already has: it is every pixel \
             that box would have been decoded to, and the frame an enlargement back is stretched \
             from"
    );
    assert!(
        owed(&path, (100, 100, 1900, 1525), PinRelayoutRoad::BoxChange),
        "while a box bigger than the frame is not, and is decoded for"
    );
    assert!(
        owed(&path, (100, 100, 500, 400), PinRelayoutRoad::TakeUp),
        "and a take-up is never covered: the media on screen is the file the pin is leaving, and \
             the size of its frame says nothing about the file arriving"
    );

    // Every raster kind this app draws is covered alike, and a page is one of them: the two differ
    // in what rasterized them and in nothing a box change asks.
    for kind in [
        MediaType::Dds,
        MediaType::Design,
        MediaType::Magick,
        MediaType::StaticImage,
        MediaType::Pdf,
        MediaType::Office,
        MediaType::Libre,
        MediaType::Calibre,
        MediaType::Comic,
    ] {
        held(kind, 1600, 1200);
        assert!(
            !owed(&path, (100, 100, 500, 400), PinRelayoutRoad::BoxChange),
            "{kind:?} is the file rasterized at a box, so a smaller box of it is the frame it \
             already has"
        );
    }

    // And the kinds a box lays out again rather than scales are owed one whatever the frame's size
    // is: half the box is half the lines at the same size, and that smaller pageful is what a frame
    // of such a kind would have been drawn as.
    for kind in [MediaType::Text, MediaType::Archive, MediaType::Peazip] {
        held(kind, 1600, 1200);
        assert!(
            owed(&path, (100, 100, 500, 400), PinRelayoutRoad::BoxChange),
            "{kind:?} is laid out at the box rather than scaled into it, so its frame is drawn \
             again at the box it is now shown in"
        );
    }

    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous;
    }
}

/// **The temporary `RHP_RESIZE_TRACE` diagnostic measures a step only where the machine named it**,
/// which is the whole of what makes it safe to leave in the tree: a run with neither variable set
/// opens no step, writes no line, and a step that is opened is one a test can read back (see
/// `pin_resize_probe`).
#[test]
fn a_resize_step_is_measured_only_where_the_diagnostic_asked_for_it() {
    pin_resize_probe::stand_the_knobs_in(Some((false, false)));
    pin_resize_probe::begin_step((100, 200, 400, 600), true);
    assert!(
        pin_resize_probe::a_step_started().is_none(),
        "nothing is measured with neither variable set"
    );
    pin_resize_probe::end_step();
    assert!(
        !pin_resize_probe::the_film_is_hidden(),
        "and no film is off the screen for a diagnostic nobody asked for"
    );

    pin_resize_probe::stand_the_knobs_in(Some((false, true)));
    pin_resize_probe::begin_step((100, 200, 400, 600), true);
    assert!(
        pin_resize_probe::a_step_started().is_some(),
        "and a step is opened where the trace did ask for one"
    );
    pin_resize_probe::forget_the_step();
    pin_resize_probe::stand_the_knobs_in(None);
}

/// **A step's line carries the box the pointer asked for, the direction the step went in, and each
/// cost the step paid** — the numbers a future agent reads to see where a step's time went. A cost
/// written as a dash is one nothing measured, and a fit that was asked for asynchronously says so,
/// because that is the difference between a step that waited on mpv's thread and one that did not.
#[test]
fn a_resize_step_writes_its_box_its_direction_and_the_costs_it_paid() {
    let took = |ms: u64| Some(Duration::from_micros(ms * 1000));
    let mut step = pin_resize_probe::a_step_at((100, 200, 400, 600), true);
    step.paint = took(3);
    step.host = took(1);
    step.fit = took(2);
    step.fit_is_async = true;
    step.asked_for = Some((
        HWND(std::ptr::null_mut()),
        (100, 200, 300, 400),
        Instant::now(),
    ));

    let line = pin_resize_probe::a_step_line(&step);
    for field in [
        "box 100,200,300x400",
        "grew",
        "paint 3.000ms",
        "host 1.000ms",
        "fit 2.000ms async",
        "want 100,200 300x400",
    ] {
        assert!(line.contains(field), "{field} is in the line: {line}");
    }

    // A step that paid none of them says so rather than printing a zero, and one that shrank says
    // that: the two are what a line is read against.
    let mut bare = pin_resize_probe::a_step_at((100, 200, 400, 600), false);
    bare.paint = took(3);
    let line = pin_resize_probe::a_step_line(&bare);
    assert!(line.contains("shrank"), "{line}");
    assert!(line.contains("paint 3.000ms host - fit -"), "{line}");
    assert!(line.contains("child none"), "{line}");
}

/// **The box a step asked the film's window for is read as the distance it was found at**, on each of
/// the four edges a box is written in: what a later tick's `GetWindowRect` is compared against, and
/// the only number that says whether the film's window arrived at all (see `pin_resize_probe`).
#[test]
fn an_observed_film_box_is_read_as_the_distance_it_was_off_by() {
    let want = (100, 200, 300, 400);

    assert_eq!(
        pin_resize_probe::the_delta(want, want),
        (0, 0, 0, 0),
        "a window standing where it was asked for is no distance at all"
    );
    assert_eq!(
        pin_resize_probe::the_delta(want, (100, 200, 200, 400)),
        (0, 0, -100, 0),
        "and one still a step behind is short by the width it has not taken yet"
    );
    assert_eq!(
        pin_resize_probe::the_delta(want, (90, 210, 310, 390)),
        (-10, 10, 10, -10),
        "a box asked for somewhere else is off on every edge it moved"
    );
}

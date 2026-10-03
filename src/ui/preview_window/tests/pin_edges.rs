use super::*;

/// The rule a sound picked into a pin lives or dies by: the card is built from what the machine
/// has for the file, and that verdict is a *probe* — a source reader, or an `ffprobe` run, both
/// of them felt — so what a pin is owed for one is the probe rather than a card that cannot be
/// built yet. The probe's own two answers are the branches beside it: a file the machine will
/// not play is no preview at all, and one it plays is the card at its own size, laid out at the
/// box the card's own measure came out at rather than at the box the pin is standing in (see
/// `pin_update_content` and `audio_box`).
#[test]
fn a_sound_picked_into_a_pin_is_the_wait_for_its_probe_until_that_has_answered() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let folder = std::env::temp_dir().join("rust-hover-preview-pin-sound");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };
    let space = PinSwapSpace {
        current: (760, 440, 1160, 640),
        bound: None,
        transport_bar: false,
        overlay: false,
        caption: pinned_caption_height(96, None),
        maximized: false,
        room: bounds,
    };

    // What the measure a card is asked for is keyed by, so that an answer can be held for a
    // file the way the measure thread would have held it.
    let scope = {
        let options = current_audio_options();
        MeasureScope::Room {
            cap_width: (bounds.right - bounds.left).max(1) as u32,
            cap_height: bounds.height().max(1) as u32,
            dpi: 96,
            theme: options.theme,
            font_scale_percent: options.font_scale_percent,
        }
    };

    // A playable sound whose card has been measured, which is the state a hover leaves the file
    // in: the card is drawn at its own size, in the middle of the box the pin has.
    let playable = folder.join("song.mp3");
    std::fs::write(
        &playable,
        b"ID3\x04\x00\x00\x00\x00\x00\x00\x10\x00\x00\x00",
    )
    .expect("a written file");
    audio_track::remember(
        &playable,
        audio_track::Probed::Track(audio_track::Track {
            player: audio_track::Player::Ffmpeg,
            codec: Some("MP3".to_string()),
            rate: Some(44_100),
            channels: Some(2),
            bitrate: Some(192_000),
            duration: Some(180.0),
        }),
    );
    hold_box(
        &playable,
        &file_version(&playable),
        &scope,
        Some((240, 200)),
    );

    let card_box = centred_at((240, 200), (960, 540));

    assert_eq!(
        pin_update_content(space, &playable, bounds, 96),
        Some(PinBox::Measured(card_box)),
        "a measured card is drawn at its own size rather than at the box the pin has"
    );

    // And the box another file left — a picture's, with the bound its take-up wrote — is no
    // size for a card either, nor is the card put where the display would put it: what comes
    // back is the card box in the middle of the pin's own, wherever the pin is standing.
    let after_a_picture = PinSwapSpace {
        current: (100, 100, 500, 400),
        bound: Some(400),
        transport_bar: false,
        overlay: false,
        caption: pinned_caption_height(96, None),
        maximized: false,
        room: bounds,
    };

    assert_eq!(
        pin_update_content(after_a_picture, &playable, bounds, 96),
        Some(PinBox::Measured(centred_at((240, 200), (300, 250)))),
        "the box and the bound another file left are no size for a card"
    );

    // A maximize is no part of it either: a card offers no maximize to stay in, so a swap to
    // one takes the card's own box even where the window is maximized — and the take-up gives
    // the maximize up with it (see `pin_restore_after`).
    let maximized = PinSwapSpace {
        current: (0, 30, 1920, 1050),
        bound: Some(1400),
        transport_bar: false,
        overlay: false,
        caption: pinned_caption_height(96, None),
        maximized: true,
        room: bounds,
    };

    assert_eq!(
        pin_update_content(maximized, &playable, bounds, 96),
        Some(PinBox::Measured(card_box)),
        "a maximized window is given a card's own box like any other"
    );

    // A file the machine will not play is the other answer a probe leaves behind, and it is no
    // preview at all: there is no card to swap in, so the pin keeps the file it is showing.
    let silent = folder.join("silence.mp3");
    std::fs::write(&silent, b"ID3\x04\x00\x00\x00\x00\x00\x00\x20\x00\x00\x00")
        .expect("a written file");
    audio_track::remember(&silent, audio_track::Probed::Nothing);
    hold_box(&silent, &file_version(&silent), &scope, None);

    assert_eq!(
        pin_update_content(space, &silent, bounds, 96),
        None,
        "a file no engine here plays has no card"
    );

    // And a sound nothing has looked at yet is the wait for the probe that would say which of
    // the two it is — started by the measure this call takes, and answered on the thread it
    // runs on.
    let fresh = folder.join("unheard.mp3");
    std::fs::write(&fresh, b"ID3\x04\x00\x00\x00\x00\x00\x00\x30\x00\x00\x00")
        .expect("a written file");

    assert_eq!(
        pin_update_content(space, &fresh, bounds, 96),
        Some(PinBox::Waiting),
        "a sound no probe has answered for is a wait"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// A file drawn to whatever box it is given is laid out in the room rather than fitted into the
/// bound, which is what a sound's card is already given: the bound is a size some *other* file
/// came out at, and neither a page of text nor a card is drawn to that or fitted into it. The
/// window keeps the place it stands on the display and only its size changes about the middle of
/// it (see `pin_keeps_its_box`).
#[test]
fn a_text_file_picked_into_a_pin_is_laid_out_in_the_room_and_not_fitted_into_the_bound() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let folder = std::env::temp_dir().join("rust-hover-preview-pin-text");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("notes.txt");
    std::fs::write(&path, b"one line\nand another\n").expect("a written file");

    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };
    let space = |bound| PinSwapSpace {
        current: (760, 440, 1160, 640),
        bound,
        transport_bar: false,
        overlay: false,
        caption: pinned_caption_height(96, None),
        maximized: false,
        room: bounds,
    };

    let laid_out = |bound| {
        let Some(PinBox::Measured(laid_out)) = pin_update_content(space(bound), &path, bounds, 96)
        else {
            panic!("a page of text was not laid out");
        };
        laid_out
    };

    // The bound of the file before it is no ceiling for this one, and what says so is that the
    // box comes out the same whatever that bound is: a page of text is measured against the
    // room it is drawn in, so the number a picture left behind has nothing to say about it.
    // Under the bound it would have been fitted into, the two boxes below differ by all of it.
    let small = laid_out(Some(400));
    let large = laid_out(Some(1600));

    assert_eq!(
        small, large,
        "the box the file before it left was used as the ceiling for a page of text anyway"
    );
    assert_ne!(
        small,
        space(Some(400)).current,
        "the page kept the box the pin already had rather than being laid out for itself"
    );

    // The middle is the pin's own: a window the hand has moved keeps its place on the display
    // while it changes size about it.
    assert_eq!(
        (small.0 + small.2) / 2,
        (space(None).current.0 + space(None).current.2) / 2,
        "the window moved off the place it stood rather than changing size about its middle"
    );
    assert_eq!(
        (small.1 + small.3) / 2,
        (space(None).current.1 + space(None).current.3) / 2
    );

    // And it neither takes a bound nor gives one: a box laid out in the room is a size the file
    // was drawn to and no ceiling for the files after it (see `pin_bound_after`).
    assert_eq!(pin_bound_after(None, true, small), None);

    let _ = std::fs::remove_dir_all(&folder);
}

/// The bound a pin has once the file just taken up is on screen, which is the rule a pin that
/// follows a listing lives or dies by: one the window already has is carried over untouched, a
/// pin that has none keeps none while the file on screen is drawn to its own box, and the first
/// file with a shape of its own locks the longest side of the box it came out at (see
/// `pin_bound_after`).
#[test]
fn a_pin_taken_up_on_a_file_drawn_to_its_box_takes_a_bound_from_the_first_shaped_file() {
    // The boxes two of those kinds leave: a short page of text and a sound's card, both small
    // and neither of them a size anything was asked to fit inside.
    let page = (0, 0, 300, 900);
    let card = (100, 200, 500, 400);

    // A first pin on one of them: no bound.
    assert_eq!(pin_bound_after(None, true, card), None);

    // Another of them following it leaves the pin unbound still, however much of a side that
    // one happens to have.
    assert_eq!(pin_bound_after(None, true, page), None);

    // The first file with a shape of its own takes the longest side of the box it came out at —
    // here the 900 the tall page left, which is the size this window has shown it can hold.
    assert_eq!(pin_bound_after(None, false, page), Some(900));

    // And a bound the window already has is never rewritten: not by another shaped file, and
    // not by one drawn to its own box, which keeps the box it is given.
    assert_eq!(pin_bound_after(Some(800), false, page), Some(800));
    assert_eq!(pin_bound_after(Some(800), true, card), Some(800));
}

/// The maximize a swap leaves behind, which is what the file after a card is laid out from: a
/// window shown a picture stays maximized, and one shown a sound's card gives the maximize up —
/// a card is its own size, offers no maximize to stay in, and the box it stands in is no box
/// for a maximum that is still standing (see `pin_restore_after`).
#[test]
fn a_swap_to_a_card_gives_up_the_maximize_the_window_was_in() {
    let restore = (100, 100, 500, 400);

    assert_eq!(
        pin_restore_after(Some(restore), false),
        Some(restore),
        "a kind a maximize can be shown for keeps the state the window was in"
    );
    assert_eq!(
        pin_restore_after(Some(restore), true),
        None,
        "a card offers no maximize to stay in"
    );
    assert_eq!(pin_restore_after(None, true), None);
}

/// The case the rule above is for: a pin taken up on a sound's card — a 400 by 200 box with no
/// size of a file in it — shown a picture. The picture is laid out against the room the display
/// has at the scale its kind names rather than inside the card's box, and the side that comes
/// out of it is the bound the files after it are fitted into (see `pin_bound_after`).
#[test]
fn a_picture_shown_to_a_sound_pin_fills_the_room_and_takes_the_bound() {
    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };
    // The card the pin went up on, in the middle of the display.
    let space = PinSwapSpace {
        current: (760, 440, 1160, 640),
        bound: None,
        transport_bar: false,
        overlay: false,
        caption: pinned_caption_height(96, None),
        maximized: false,
        room: bounds,
    };

    // The room is the display's own, which the card's box does not shrink.
    let room = pin_swap_room(space, bounds, 96);
    assert_eq!(room, (0, 15, 1920, 1065));

    // A 4:3 photograph at fit-to-screen fills the room's height, and the card's 400 pixels are
    // no ceiling for it.
    let content = pin_update_box(room, (4000, 3000), PreviewScale::FitToScreen);
    assert_eq!(content, (260, 15, 1660, 1065));
    assert_eq!(pin_bound_after(None, false, content), Some(1400));

    // The file after it is fitted into a square of that side, cut to the room it is on.
    let next = PinSwapSpace {
        current: content,
        bound: Some(1400),
        transport_bar: false,
        overlay: false,
        caption: pinned_caption_height(96, None),
        maximized: false,
        room: bounds,
    };
    let next_room = pin_swap_room(next, bounds, 96);
    assert_eq!(
        (next_room.2 - next_room.0, next_room.3 - next_room.1),
        (1400, 1050)
    );
}

/// The box a measure answers with while it runs is a wait rather than a shape, and what makes
/// it a wait is the read running behind it: the placeholder is a square, so laying a swap out
/// at it is a square window whatever the file is (see `box_is_the_wait` and `pin_swap_awaits`).
#[test]
fn a_placeholder_box_is_a_wait_only_while_a_read_is_running_for_it() {
    let path = std::env::temp_dir()
        .join("rust-hover-preview-wait-tests")
        .join("waited-page.pdf");
    let wait = (office_preview::WAITING_BOX, office_preview::WAITING_BOX);

    assert!(box_is_the_wait(wait));
    assert!(
        !box_is_the_wait((1920, 1080)),
        "a box of a file's own shape is not the wait for one"
    );

    // Nothing is reading this file, so what it was measured by is its own box, whatever that
    // is — and there is nothing for a swap to wait for.
    assert!(!pin_swap_awaits(&path, wait));

    // A read is running for it, and the box in hand is what that read answered with.
    begin_measure(&path);
    assert!(pin_swap_awaits(&path, wait));
    assert!(
        !pin_swap_awaits(&path, (1920, 1080)),
        "a box that is the file's own is not the wait, however a read is going"
    );

    end_measure(&path);
    assert!(
        !pin_swap_awaits(&path, wait),
        "a placeholder nothing is reading behind is a size that never changes, and a wait on \
             it would never end"
    );
}

/// The same drag on a pin whose chrome is drawn *over* its media, which is every kind this app
/// draws for itself: the window there is the media's own box, so a drag has no caption and no
/// bar to give room to — and the box it comes out with is the media's either way.
#[test]
fn a_resize_of_a_pin_whose_chrome_is_over_its_media_has_no_room_to_leave() {
    let edge = edge(false, false, true, false);

    // What the drag produces is the media box, and the same one whether or not the kind has
    // bands of its own: a window a caption taller than the picture in it is a picture the paint
    // stretches into the band, and the band being where the next drag measures the box from is
    // what made a hand that kept at it watch the picture grow a caption per gesture.
    let media = (400, 300, 800, 600);
    let plain = dragged_from(media, PinFrame::Shaped, edge, 200, 0);
    let over = dragged_overlay(media, PinFrame::Shaped, edge, 200, 0, true);
    assert_eq!(plain, (400, 225, 1000, 675));
    assert_eq!(
        over, plain,
        "the window is the media and nothing is added to it"
    );

    // A gesture that asks for the box it already has is the sharpest way to say it: nothing
    // about the box may move, however many times it is done.
    let mut box_ = over;
    for _ in 0..8 {
        box_ = dragged_overlay(box_, PinFrame::Shaped, edge, 0, 0, true);
    }
    assert_eq!(box_, over, "a drag of nothing changes nothing");

    // And one that asks for a bigger box leaves the media's own shape in the box it leaves.
    let grown = dragged_overlay(over, PinFrame::Shaped, edge, 200, 0, true);
    assert_eq!(grown, (400, 150, 1200, 750));
    assert_eq!(
        (grown.2 - grown.0) as f64 / (grown.3 - grown.1) as f64,
        4.0 / 3.0,
        "the media keeps its shape"
    );
}

/// A window is pulled smaller by the same drags that pull it larger, from every edge and every
/// corner: an inward pull is a box under the hand, and a box that only ever grew was a resize
/// with half of its range missing.
#[test]
fn a_pinned_window_is_pulled_smaller_from_every_edge_and_corner() {
    assert_eq!(
        dragged(edge(false, false, true, false), -100, 0),
        (400, 337, 700, 562),
        "the right edge pulled in"
    );
    assert_eq!(
        dragged(edge(true, false, false, false), 100, 0),
        (500, 337, 800, 562),
        "the left edge pulled in"
    );
    assert_eq!(
        dragged(edge(false, false, false, true), 0, -75),
        (450, 300, 750, 525),
        "the bottom edge pulled in"
    );
    assert_eq!(
        dragged(edge(false, true, false, false), 0, 75),
        (450, 375, 750, 600),
        "the top edge pulled in"
    );
    assert_eq!(
        dragged(edge(false, false, true, true), -100, -75),
        (400, 300, 700, 525),
        "the bottom-right corner pulled in"
    );
    assert_eq!(
        dragged(edge(true, true, false, false), 100, 75),
        (500, 375, 800, 600),
        "the top-left corner pulled in"
    );
}

/// A resize begun from a box a drag left partly off the room keeps the dragged place where
/// the hand did not change the size: a maximized window fills the room in one dimension, so
/// asking past the room is refused and the size comes back unchanged — and centering an
/// unchanged room-sized box would snap it back onto the room's edge, forgetting the drag.
/// That is Maximize → drag → resize jumping back to the top; a drag → resize while not
/// maximized never fills the room, so it never meets the clamp at all.
#[test]
fn a_resize_from_a_dragged_maximized_box_keeps_the_dragged_place() {
    // A maximized 4:3 picture fills the 1200x900 room exactly, then is carried to
    // (100, 90): partly off the room's right and bottom edges.
    let moved = (100, 90, 1300, 990);

    // Pulling the left edge further out cannot grow past the room, so the box is
    // already what the hand asked for: the top stays at the dragged line 90
    // rather than snapping back to the room's top 0.
    assert_eq!(
        dragged_overlay(
            moved,
            PinFrame::Shaped,
            edge(true, false, false, false),
            -25,
            0,
            true
        ),
        moved,
        "the left edge pulled out past the room"
    );
    // The user's own gesture: pulling the top border up cannot grow past the
    // room either, so the left stays at the dragged line 100.
    assert_eq!(
        dragged_overlay(
            moved,
            PinFrame::Shaped,
            edge(false, true, false, false),
            0,
            -25,
            true
        ),
        moved,
        "the top edge pulled up past the room"
    );
}

/// The same promise for a kind laid out to its box: a maximized page fills the room in
/// both dimensions, so a refused grow on one axis must not re-center the other back
/// onto the room either.
#[test]
fn a_resize_from_a_dragged_maximized_page_keeps_the_dragged_place() {
    let moved = (100, 90, 1300, 990);

    assert_eq!(
        dragged_from(
            moved,
            PinFrame::Free,
            edge(false, true, false, false),
            0,
            -25
        ),
        moved,
        "the top edge pulled up past the room"
    );
    assert_eq!(
        dragged_from(
            moved,
            PinFrame::Free,
            edge(true, false, false, false),
            -25,
            0
        ),
        moved,
        "the left edge pulled out past the room"
    );
}

/// Maximize → drag → resize keeps the place the drag left the window at, whichever edge
/// is pulled. This is the whole of the snapping: a maximized window fills the room, so on the
/// axis the hand is not on there is no room left for the box to be centred in, and the clamp
/// that keeps it inside the room collapses onto the room's edge instead — a top-edge drag
/// taking a 16:9 picture from left 100 to 33, and a left-edge drag taking it from top 90 to
/// 19, which is the window jumping to the top of the screen the moment an edge is touched.
///
/// Note that the picture's own shape is kept on both axes, so the size on the *untouched* axis
/// changes too: a guard that only kept the place when the size had not moved never sees this
/// case, which is why the snapping survived it.
#[test]
fn a_resize_of_a_maximized_box_keeps_the_dragged_place_on_both_axes() {
    // A maximized picture filling the 1200x900 room, then carried to (100, 90).
    let moved = (100, 90, 1300, 990);

    // Every edge and corner, for a picture that keeps its shape as it grows.
    let cases: [(&str, PinResize, i32, i32); 5] = [
        ("top", edge(false, true, false, false), 0, 25),
        ("bottom", edge(false, false, false, true), 0, 25),
        ("left", edge(true, false, false, false), 25, 0),
        ("right", edge(false, false, true, false), 25, 0),
        ("bottom-right", edge(false, false, true, true), 25, 25),
    ];
    for (name, e, dx, dy) in cases {
        let got = dragged_overlay(moved, PinFrame::Shaped, e, dx, dy, true);
        // Only the line under the hand may move: an axis the hand is not on has to stay where
        // the drag left it, which is the whole of what the snapping broke.
        if !(e.left || e.right) {
            assert_eq!(got.0, moved.0, "the left line moves on a {name} drag");
        }
        if !(e.top || e.bottom) {
            assert_eq!(got.1, moved.1, "the top line moves on a {name} drag");
        }
    }

    // And for a page laid out to its box, which resizes one dimension at a time — the
    // same promise, asked of the axis the page's resize does not touch.
    for (name, e, dx, dy) in [
        ("top", edge(false, true, false, false), 0i32, 25i32),
        ("left", edge(true, false, false, false), 25i32, 0i32),
    ] {
        let got = dragged_overlay(moved, PinFrame::Free, e, dx, dy, true);
        if !(e.left || e.right) {
            assert_eq!(got.0, moved.0, "a page's left line moves on a {name} drag");
        }
        if !(e.top || e.bottom) {
            assert_eq!(got.1, moved.1, "a page's top line moves on a {name} drag");
        }
    }
}

/// A drag cannot take the media off the display it is on, and the shape it keeps is kept
/// against that limit rather than cut by the placement afterwards: a box pulled wider than the
/// room has, whose other side was then trimmed to bring it back, is a picture of another shape
/// inside the box drawn around it.
#[test]
fn a_drag_is_stopped_by_the_room_rather_than_cut_by_it() {
    // A box against the top of the room has nowhere to grow upwards, so the line it is centered
    // on is a line it is moved along rather than held to: a right-edge pull still grows it.
    let at_the_top = (400, 0, 800, 300);
    let content = dragged_from(
        at_the_top,
        PinFrame::Shaped,
        edge(false, false, true, false),
        600,
        0,
    );
    assert_eq!(content, (400, 0, 1200, 600));
    assert_eq!(
        (content.2 - content.0) as f64 / (content.3 - content.1) as f64,
        400.0 / 300.0,
        "and what it grew to is still the shape of the media"
    );

    // The same for a corner: the box may not be pulled past the room's own edge.
    let content = dragged_from(
        (400, 300, 800, 600),
        PinFrame::Shaped,
        edge(false, false, true, true),
        10_000,
        10_000,
    );
    assert_eq!(content, (400, 300, 1200, 900));
}

/// The four places a window can be resized from are found where they are drawn, and the middle
/// of it is not one of them: a band is a band, and a pointer that is not on one is a pointer
/// with nothing of the frame under it.
#[test]
fn a_pinned_window_says_which_of_its_edges_a_point_is_on() {
    let pin = PinnedPreview {
        path: PathBuf::from("picture.png"),
        bound: Some(600),
        content: (100, 100, 700, 500),
        restore: None,
        dpi: 96,
        transport_bar: false,
        transport_live: false,
        frame: PinFrame::Shaped,
        overlay: false,
        hides_chrome: false,
        caption: pinned_caption_height(96, None),
        chrome: PinChrome::always(),
        collapsed: false,
        bubble_pause: None,
        hovered: None,
        pressed: None,
        tooltip: PinTooltip::default(),
        dragging: None,
        parked: false,
        transport: PinTransport::default(),
        volume: PinVolume::default(),
        audio_hovered: None,
        audio_pressed: None,
    };
    let (width, height) = pin.window_size();

    // A side, at the middle of it, on all four sides.
    assert_eq!(
        pin.resize_edge(2, height / 2),
        Some(edge(true, false, false, false))
    );
    assert_eq!(
        pin.resize_edge(width - 2, height / 2),
        Some(edge(false, false, true, false))
    );
    assert_eq!(
        pin.resize_edge(width / 2, 2),
        Some(edge(false, true, false, false))
    );
    assert_eq!(
        pin.resize_edge(width / 2, height - 2),
        Some(edge(false, false, false, true))
    );

    // A corner, on all four of them — the top two included, which are the two a caption asked
    // about first would have swallowed.
    assert_eq!(pin.resize_edge(2, 2), Some(edge(true, true, false, false)));
    assert_eq!(
        pin.resize_edge(width - 2, 2),
        Some(edge(false, true, true, false))
    );
    assert_eq!(
        pin.resize_edge(2, height - 2),
        Some(edge(true, false, false, true))
    );
    assert_eq!(
        pin.resize_edge(width - 2, height - 2),
        Some(edge(false, false, true, true))
    );

    // A corner's band is wider than a side's, so a point just inside one is the corner.
    assert_eq!(
        pin.resize_edge(10, 10),
        Some(edge(true, true, false, false)),
        "inside the corner's band"
    );

    // And the media, the caption off to one side of it, and every point outside the bands are
    // none of the eight.
    assert_eq!(pin.resize_edge(width / 2, height / 2), None);
    assert_eq!(pin.resize_edge(40, 40), None);
    assert_eq!(pin.resize_edge(40, height / 2), None);
    assert_eq!(pin.resize_edge(width / 2, 40), None);
}

/// A press that lands on the window the engine draws a document in begins what a press on the
/// same place of any other kind's band begins: a resize on an edge of the pin's own window —
/// the edges that window covers, which is every one of them but the caption's — and a move
/// anywhere else. The point is read from the pointer rather than from a message, so it arrives
/// in screen coordinates and the window's own origin is taken out here.
#[test]
fn a_press_on_the_engines_window_begins_the_drag_its_place_means() {
    // A pin whose window is (100, 100, 700, 500): the caption is the strip along the top, and
    // the engine's window is everything below it and the full width.
    let window = (100, 100, 700, 500);
    let action_at =
        |x: i32, y: i32| pinned_engine_press_action(window, 96, PinFrame::Shaped, (x, y));
    let edge_of = |action: PinDragAction| match action {
        PinDragAction::Resize(edge) => Some(edge),
        PinDragAction::Move => None,
    };

    // The middle of the band is the media, which is the handle a window's body is.
    assert!(matches!(action_at(400, 300), PinDragAction::Move));

    // The left, right and bottom edges of the pin's window are the engine's window's own —
    // the caption is above them — so a press there is a resize, in the same bands the window
    // procedure finds them in.
    assert_eq!(
        edge_of(action_at(100, 300)),
        Some(edge(true, false, false, false)),
        "the left edge"
    );
    assert_eq!(
        edge_of(action_at(700, 300)),
        Some(edge(false, false, true, false)),
        "the right edge"
    );
    assert_eq!(
        edge_of(action_at(400, 500)),
        Some(edge(false, false, false, true)),
        "the bottom edge"
    );
    assert_eq!(
        edge_of(action_at(100, 500)),
        Some(edge(true, false, false, true)),
        "and a corner of two of them"
    );

    // A hand just inside the edge is the media, which is the same border a press on the
    // opaque parts of the pin is measured by — eight pixels at this scale.
    assert!(matches!(action_at(109, 300), PinDragAction::Move));
    assert!(matches!(action_at(400, 490), PinDragAction::Move));

    // A window that offers no resize at all — a sound's card — moves from its edges as it
    // does from its middle: the edges are not the pin's to offer there.
    let card = |x: i32, y: i32| pinned_engine_press_action(window, 96, PinFrame::None, (x, y));
    assert!(matches!(card(100, 300), PinDragAction::Move));
    assert!(matches!(card(400, 300), PinDragAction::Move));
}

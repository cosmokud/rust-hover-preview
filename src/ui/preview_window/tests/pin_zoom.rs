use super::*;

/// The room every one of these zooms is measured in: a display of 1000 by 800 at the origin, which is
/// what `bounds` is in `tests` — and a picture whose shape the boxes below keep, since a pinned
/// picture is only ever given boxes of its own shape (`PinFrame::Shaped`).
const DESK: u32 = TEST_DPI;

fn display() -> RecordedDisplays {
    RecordedDisplays::one_display(bounds(), DESK)
}

/// What a notch comes to, for a pin standing at `content` with the view it is handed.
fn zoomed(
    content: ScreenRegion,
    zoom: PinZoom,
    notches: i32,
    pointer: Option<(i32, i32)>,
) -> PinZoomPlan {
    pin_zoom_after(
        &PinZoomInputs {
            content,
            zoom,
            dpi: DESK,
            room: bounds(),
            pointer,
        },
        notches,
        &display(),
    )
    .expect("a notch that asks for something")
}

/// A picture 400 by 300 standing in the middle of the desk: four thirds against a room of five
/// quarters, so the fit of it is limited by the room's *width* at 1000 by 750 and the room has a
/// hundred pixels of it to spare at the top and the bottom.
fn a_picture() -> ScreenRegion {
    (300, 250, 700, 550)
}

/// The same picture taken all the way to the cap, one notch at a time with the pointer held still.
fn at_the_cap(pointer: (i32, i32)) -> PinZoomPlan {
    let mut plan = zoomed(a_picture(), PinZoom::filling(a_picture()), 1, None);
    let mut content = plan.content;
    let mut zoom = plan.zoom;

    for _ in 0..40 {
        plan = zoomed(content, zoom, 1, Some(pointer));
        content = plan.content;
        zoom = plan.zoom;
    }

    plan
}

#[test]
fn one_notch_up_grows_the_window_about_the_pictures_own_middle() {
    // A 400 x 300 picture grown by a tenth is 440 x 330, and a box grown about its middle moves its
    // opposite sides by half the difference each way: 20 to the left and 20 to the right, 15 up and
    // 15 down. That is the whole of the anchor before the window reaches its cap, and it is what
    // keeps a picture in the place the hand was looking at rather than walking it off the screen in
    // one direction.
    let plan = zoomed(a_picture(), PinZoom::filling(a_picture()), 1, None);
    assert_eq!(plan.content, (280, 235, 720, 565));
    assert_eq!(plan.zoom.drawn, (440.0, 330.0));
    assert_eq!(plan.zoom.pan, (0.0, 0.0));
    assert_eq!(
        ((plan.content.2 - plan.content.0) as f64) / ((plan.content.3 - plan.content.1) as f64),
        4.0 / 3.0,
        "the window keeps the picture's shape"
    );

    // Three notches in one gesture are the same step taken further rather than three boxes: a wheel
    // turned fast lands on the same size a wheel turned slowly would have arrived at.
    let three = zoomed(a_picture(), PinZoom::filling(a_picture()), 3, None);
    let step = PIN_ZOOM_PER_NOTCH.powi(3);
    assert_eq!(
        three.content.2 - three.content.0,
        (400.0 * step).round() as i32
    );
    assert_eq!(
        three.content.3 - three.content.1,
        (300.0 * step).round() as i32
    );
}

#[test]
fn one_notch_down_shrinks_it_by_exactly_what_a_notch_up_grew_it_by() {
    // Out is in divided rather than a step of its own, which is the whole of what lets a hand come
    // back from either end without counting: the same notch back lands on the box it started from.
    let grown = zoomed(a_picture(), PinZoom::filling(a_picture()), 1, None);
    let back = zoomed(grown.content, grown.zoom, -1, None);

    assert_eq!(back.content, a_picture());
    assert_eq!(back.zoom.drawn, (400.0, 300.0));
}

#[test]
fn the_window_stops_growing_where_fit_to_screen_would_have_stopped_it() {
    // Fit to Screen for a 4:3 picture in a 1000 x 800 room is limited by the room's width: 1000 by
    // 750, which is the largest the window may ever be. So the picture goes on growing inside a
    // window that has stopped rather than the window going on growing to any size at all.
    let plan = at_the_cap((500, 400));

    assert_eq!(
        plan.content,
        (0, 25, 1000, 775),
        "the cap is the fit-to-screen box"
    );
    assert!(
        plan.zoom.drawn.0 > 4000.0,
        "the picture has gone on growing inside a window that stopped: {}",
        plan.zoom.drawn.0
    );
    assert_eq!(plan.zoom.share_of(plan.content), plan.zoom.drawn.0 / 1000.0);
    assert!(
        plan.zoom.picture_box((1000, 750)).0 < 0,
        "and the window is a box the picture overhangs"
    );
}

#[test]
fn a_picture_beyond_a_fit_is_given_a_cap_of_its_own_size_rather_than_being_shrunk_to_one() {
    // A photograph taken up at `100%` on a screen smaller than it is comes up larger than any fit —
    // and the cap is there to stop a window growing, so a notch that asked to *grow* one may not be
    // answered by shrinking it. What it is instead is the first notch of a picture already inside its
    // window: the box holds still and the picture grows within it.
    let huge = (0, 0, 2000, 1500);
    let plan = zoomed(huge, PinZoom::filling(huge), 1, Some((500, 400)));

    assert_eq!(
        (
            plan.content.2 - plan.content.0,
            plan.content.3 - plan.content.1
        ),
        (2000, 1500),
        "the window holds the size it already had"
    );
    assert_eq!(plan.zoom.drawn.0, 2200.0, "and the picture grows inside it");

    // And a notch the other way takes the same window down rather than leaving it at the size it came
    // up at, so a photograph that arrived too large can be brought back to a fit.
    let out = zoomed(huge, PinZoom::filling(huge), -1, None);
    assert_eq!(
        (out.content.2 - out.content.0, out.content.3 - out.content.1),
        (1818, 1364)
    );
    assert_eq!(out.zoom.pan, (0.0, 0.0));
}

#[test]
fn zooming_out_stops_when_the_picture_is_a_hundredth_of_the_screen() {
    // The floor is the display's own room at a hundredth of itself, in *either* direction, so a box
    // of the picture's shape stops at the first of the two to be reached: a hundredth of 1000 by 800
    // is 10 by 8, and a 4:3 picture reaching 8 tall is 10.67 wide — so it is the *height* that sets
    // the floor here and the picture stops at 10.67 by 8 rather than shrinking past the limit in the
    // other direction.
    let mut content = a_picture();
    let mut zoom = PinZoom::filling(a_picture());

    for _ in 0..200 {
        let plan = zoomed(content, zoom, -1, None);
        content = plan.content;
        zoom = plan.zoom;
    }

    let (width, height) = region_size(content);
    assert!(
        width >= 10 && height >= 8,
        "a picture never shrinks below a hundredth of the room: {width} by {height}"
    );
    assert!(
        width <= 12 && height <= 9,
        "and it is taken all the way down to it: {width} by {height}"
    );
    // The shape is the picture's own to within the pixel a box on a screen is whole in, which at this
    // size is all the accuracy there is to ask for: 10.67 by 8 is four thirds and 11 by 8 is not.
    assert!(
        (width as f64 - height as f64 * 4.0 / 3.0).abs() <= 1.0,
        "the shape is kept to the very end: {width} by {height}"
    );
}

#[test]
fn past_the_cap_the_anchor_is_the_pointer_rather_than_the_pictures_middle() {
    // A window at the cap with the picture filling it, and the hand a quarter of the way in from the
    // left. What a zoom in owes that hand is that the *point of the picture under the pointer* stays
    // under it — which a picture grown about its own middle would not: the point under the pointer
    // would slide away from it, and a hand zooming into a face would watch the face walk off.
    let window = (0, 0, 1000, 800);
    let zoom = PinZoom::filling(window);
    let pointer = (250, 400);
    let plan = zoomed(window, zoom, 1, Some(pointer));

    assert_eq!(plan.content, window, "the window is at its cap and holds");
    assert_eq!(plan.zoom.drawn.0, 1100.0);

    // Where the point of the picture that was under the pointer has landed. The picture is drawn a
    // tenth larger than the window and carried by the pan, so the point that was 250 to the left of
    // the middle has to still be 250 to the left of it.
    let (drawn_width, _) = plan.zoom.drawn;
    let left = (1000.0 - drawn_width) / 2.0 + plan.zoom.pan.0;
    let centre = left + drawn_width / 2.0;
    let landed = centre + (pointer.0 as f32 - 500.0) / 1000.0 * drawn_width;
    assert!(
        (landed - pointer.0 as f32).abs() < 0.01,
        "the point under the pointer stayed under it: {landed} against {}",
        pointer.0
    );

    // And a hand at the very middle of the window is the case the two anchors agree on, which is why
    // a zoom into the middle of a picture looks the same either side of the cap.
    let middle = zoomed(window, zoom, 1, Some((500, 400)));
    assert_eq!(middle.zoom.pan, (0.0, 0.0));
}

#[test]
fn a_notch_that_first_takes_the_picture_past_the_window_never_pulls_it_off_one() {
    // The notch that reaches the cap in the same breath as it takes the picture past it is the one
    // worth pinning down: the window grows on both sides and the picture grows more, so a picture
    // anchored on the pointer by the whole of its overhang would be pulled off the far side of a
    // window it was covering a moment ago. What it is owed is the overhang the *window* left room for
    // and no more, which is a picture that still covers the box it is drawn in.
    for pointer in [(0, 0), (500, 400), (999, 799), (0, 400), (999, 400)] {
        let plan = zoomed(
            a_picture(),
            PinZoom::filling(a_picture()),
            12,
            Some(pointer),
        );
        let box_ = plan.zoom.picture_box(region_size(plan.content));

        assert!(
            box_.0 <= 0 && box_.1 <= 0 && box_.2 >= 1 && box_.3 >= 1,
            "a picture covers its window at the cap, pointed at {pointer:?}: {box_:?}"
        );
    }
}

#[test]
fn zooming_back_out_of_the_window_consumes_the_overhang_before_it_shrinks_it() {
    // A picture half as large again as its window, carried a little. A notch out takes a tenth off
    // the *picture*, so the window holds still while the overhang is spent — and the moment the
    // picture is the window's own size again the window starts coming down with it and the pan is
    // given up, because a picture that size has nowhere to sit.
    let window = (0, 0, 1000, 800);
    let inside = PinZoom::still((1500.0, 1200.0), (-80.0, 40.0));

    let first = zoomed(window, inside, -1, Some((300, 300)));
    assert_eq!(first.content, window, "the overhang is spent first");
    assert!(first.zoom.drawn.0 > 1000.0 && first.zoom.drawn.0 < 1500.0);
    assert_ne!(
        first.zoom.pan, inside.pan,
        "and the pointer is still the anchor"
    );

    let mut content = window;
    let mut zoom = inside;
    let mut shrunk = false;

    for _ in 0..40 {
        let plan = zoomed(content, zoom, -1, Some((300, 300)));
        shrunk |= plan.content != window;
        content = plan.content;
        zoom = plan.zoom;
    }

    assert!(shrunk, "the window comes down once the picture fits it");
    assert_eq!(zoom.pan, (0.0, 0.0), "a picture filling it sits nowhere");
    assert_eq!(
        zoom.picture_box(region_size(content)),
        (0, 0, region_size(content).0, region_size(content).1),
        "a picture that fits is the whole of the band it is drawn in"
    );
}

#[test]
fn a_notch_of_zero_is_no_plan_at_all() {
    // The hook counts notches in wheel units and the loop divides by a notch, so a count that came to
    // nothing has to answer nothing rather than a box it happened to work out.
    assert!(pin_zoom_after(
        &PinZoomInputs {
            content: a_picture(),
            zoom: PinZoom::filling(a_picture()),
            dpi: DESK,
            room: bounds(),
            pointer: None,
        },
        0,
        &display(),
    )
    .is_none());
}

#[test]
fn a_picture_is_drawn_on_the_whole_of_a_window_it_fills_and_inside_one_it_does_not() {
    // Scale one is the ordinary pinned picture: the box the media lands in is the whole band, which is
    // the case every repaint of everything but a zoomed picture goes through.
    let filling = PinZoom::filling((0, 0, 1000, 800));
    assert_eq!(filling.picture_box((1000, 800)), (0, 0, 1000, 800));

    // And twice the window's size is a picture centred until a hand carries it: half of the overhang
    // hangs off each side.
    let inside = PinZoom::still((2000.0, 1600.0), (0.0, 0.0));
    assert_eq!(inside.picture_box((1000, 800)), (-500, -400, 1500, 1200));

    // A pan moves it the way the hand went: a pointer carried to the right shows the left of the
    // picture, so the box steps left.
    let carried = PinZoom::still((2000.0, 1600.0), (120.0, -60.0));
    assert_eq!(carried.picture_box((1000, 800)), (-380, -460, 1620, 1140));

    // A picture with no room to be moved has nowhere to be but the whole window, whatever the pan
    // remembers of a bigger one.
    let back = PinZoom::still((1000.0, 800.0), (400.0, 400.0));
    assert_eq!(back.picture_box((1000, 800)), (0, 0, 1000, 800));
}

#[test]
fn a_box_laid_out_again_puts_a_zoomed_picture_back_to_filling_it() {
    // Every road that chooses a box for the whole of a window — a walk, a maximize, a double-click, an
    // edge dragged — owes the picture back to the size of that box, and owes it to nothing else: a
    // kind the wheel does not zoom has no view to reset and must not be given one.
    let mut pin = PinnedPreview {
        content: (100, 100, 500, 400),
        zoom: Some(PinZoom::still((1600.0, 1200.0), (200.0, -100.0))),
        ..PinnedPreview::for_test()
    };

    reset_pin_zoom(&mut pin);
    assert_eq!(pin.zoom, Some(PinZoom::filling((100, 100, 500, 400))));

    let mut video = PinnedPreview {
        zoom: None,
        ..PinnedPreview::for_test()
    };
    reset_pin_zoom(&mut video);
    assert_eq!(video.zoom, None);
}

#[test]
fn the_wheel_is_handed_the_box_of_a_pinned_picture_and_of_nothing_else() {
    // The lock every test that reaches the pin's own state takes: what is shared is process-wide, and
    // a pin stood up here is a pin taken out from under whatever test beside this one is standing its
    // own (see `PIN_TESTS_ONE_AT_A_TIME`).
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    // What the wheel hook reads has to be answered without waiting for the pin's lock, and it has to
    // be answered about the picture on the screen and nothing else: a bubble has no window to point
    // at, a film's picture belongs to a player of somebody else's, and a hand already carrying the
    // window is not a hand asking for a zoom.
    let stood = |pin: PinnedPreview| {
        stand_pin(Some(pin));
        pin_zoom_keep_alive_try()
    };

    let picture = PinnedPreview {
        content: (100, 100, 500, 400),
        ..PinnedPreview::for_test()
    };
    assert_eq!(stood(picture), Some((100, 100, 500, 400)));

    assert_eq!(
        stood(PinnedPreview {
            collapsed: true,
            ..PinnedPreview::for_test()
        }),
        None,
        "a bubble has no window of its own"
    );

    assert_eq!(
        stood(PinnedPreview {
            zoom: None,
            ..PinnedPreview::for_test()
        }),
        None,
        "a kind the wheel does not zoom is not a box to point at"
    );

    stand_pin(None);
    assert_eq!(pin_zoom_keep_alive_try(), None, "and no pin is no box");
}

#[test]
fn a_wheel_that_asks_for_more_notches_than_any_hand_can_turn_is_answered_with_a_size_that_holds() {
    // The hook accumulates whatever a device reports, and a single wheel message can carry a delta of
    // tens of thousands of notches: a step of that many would be a drawn size no `f32` holds, and the
    // box the picture is then placed in is a rectangle of whole pixels. So the one thing "no hard
    // limit" still owes is a number the placement can add to, and it is asked here at both ends —
    // the ceiling and the floor are the only two counts this function is not free on.
    let window = (0, 0, 1000, 800);
    let huge = zoomed(window, PinZoom::filling(window), i32::MAX, Some((500, 400)));
    assert_eq!(huge.zoom.drawn.0, PIN_ZOOM_CEILING);
    assert!(
        huge.zoom.picture_box((1000, 800)).2 < i32::MAX,
        "and the box it lands in is a rectangle with room either side of it"
    );

    let tiny = zoomed(window, PinZoom::filling(window), i32::MIN, None);
    let (width, height) = region_size(tiny.content);
    assert!(
        width >= 10 && height >= 8,
        "a notch out that no hand can turn stops at the floor rather than going past it: {width} by {height}"
    );
}

/// A zoom is a *travel* and not a jump, and the whole of what makes it one is that it closes on the size
/// the wheel asked for: every frame brings the picture nearer to it, no frame takes it past it, and the
/// frames stop once it is there.
///
/// The frames are asked for at the loop's own cadence, sixteen milliseconds apart, because that is the
/// cadence a travel is taken at — and what is asserted is the shape of the whole travel rather than any
/// one frame of it, since that is the difference between this and the two things it is not: a curve of a
/// fixed length, which cannot take a second notch in the middle of its own travel, and a chase that
/// closes a fixed share of the distance each frame, which arrives at a rate that depends on how many
/// frames a second the machine happens to run at.
#[test]
fn a_travel_closes_on_the_size_the_wheel_asked_for_and_never_passes_it() {
    let (from, toward) = (400.0f32, 440.0f32);
    let (mut drawn, mut velocity) = (from, 0.0f32);
    let mut frames = 0;

    loop {
        let (next, next_velocity) = pin_zoom_spring(drawn, toward, velocity, 0.016);
        assert!(
            next >= drawn && next <= toward,
            "frame {frames} took the picture from {drawn} to {next}, and a zoom may only travel the way it was asked to"
        );

        drawn = next;
        velocity = next_velocity;
        frames += 1;

        if drawn == toward {
            break;
        }
        assert!(
            frames < 120,
            "a travel of two seconds is not a travel: {drawn} of {toward}"
        );
    }

    assert!(
        frames > 3,
        "a step is a travel of some frames and not a jump: {frames}"
    );
    assert_eq!(
        velocity, 0.0,
        "and it arrives at rest rather than at a speed"
    );
}

/// A notch taken while the last one is still travelling is counted from the size the wheel *asked for*
/// and not from the size the picture happens to be passing — which is the whole of what makes a fast
/// wheel and a slow one land in the same place, and what makes a notch in and a notch out leave the
/// picture exactly where it was however fast they were turned.
#[test]
fn a_notch_taken_mid_travel_is_counted_from_the_size_the_wheel_asked_for() {
    let window = (0, 0, 1000, 800);
    let settled = PinZoom::filling(window);
    let asked = |zoom| PinZoomInputs {
        content: window,
        zoom,
        dpi: DESK,
        room: bounds(),
        pointer: None,
    };

    let first = pin_zoom_toward(&asked(settled), 1);

    // The picture part of the way to the size it was asked for, which is the state a second notch
    // arrives in whenever a hand turns the wheel faster than the picture can follow.
    let mid = PinZoom {
        drawn: (settled.drawn.0 * 1.05, settled.drawn.1 * 1.05),
        toward: first,
        ..settled
    };

    let second = pin_zoom_toward(&asked(mid), 1);
    assert!(
        (second - first * PIN_ZOOM_PER_NOTCH).abs() < 0.01,
        "the second notch is counted from what the first one asked for: {second} against {first}"
    );
    assert!(
        (second - mid.drawn.0 * PIN_ZOOM_PER_NOTCH).abs() > 1.0,
        "and not from where the picture has got to"
    );

    // And two notches together are the one step they would have been: a fast wheel and a slow one come
    // out at the same size.
    assert!(
        (second - pin_zoom_toward(&asked(settled), 2)).abs() < 0.01,
        "two notches are the same size whether they were taken apart or together"
    );
}

/// A travel of a picture the window has stopped growing for leaves the window exactly where it is and
/// carries the picture inside it — which is the whole of what a zoom past the cap is, taken a frame at a
/// time rather than in one step.
///
/// It is the frame the *loop* takes, and it is where the two stages of a zoom are held apart: the window
/// is the box the cap allows and nothing the wheel asks for may move it again, while the picture goes on
/// growing inside it and stays anchored on the point the wheel read (see `PinZoom::anchor`).
#[test]
fn a_travel_past_the_cap_leaves_the_window_where_it_is() {
    let cap = at_the_cap((500, 400));
    let stopped = cap.content;
    let toward = cap.zoom.drawn.0 * PIN_ZOOM_PER_NOTCH;
    let asking = |zoom| PinZoomInputs {
        content: stopped,
        zoom,
        dpi: DESK,
        room: bounds(),
        pointer: Some((500, 400)),
    };

    let step = pin_zoom_step(
        &asking(PinZoom {
            toward,
            ..cap.zoom
        }),
        0.016,
        &display(),
    )
    .expect("a picture on its way to a size is a frame to take");

    assert_eq!(
        step.content, stopped,
        "the window has stopped and stays stopped"
    );
    assert!(
        step.zoom.drawn.0 > cap.zoom.drawn.0 && step.zoom.drawn.0 < toward,
        "and the picture is part of the way to the size it was asked for: {} of {toward}",
        step.zoom.drawn.0
    );
    assert_eq!(step.zoom.toward, toward, "aimed where the wheel asked");
    assert_ne!(step.zoom.velocity, 0.0, "and still moving");

    assert!(
        pin_zoom_step(&asking(cap.zoom), 0.016, &display()).is_none(),
        "a picture already at the size it was asked for has nothing to carry"
    );
}

/// A wheel notch does not move the picture at all: what it writes down is where the picture is *going*,
/// and the frames that take it there are the loop's.
///
/// That is the whole of why a notch is cheap enough to be counted on a hook every wheel message on the
/// desktop waits for — it is one number written under one lock — and it is what the loop's cadence is
/// then kept for: a picture with somewhere to be is a picture with something to animate (see
/// `pin_zoom_in_flight`).
#[test]
fn a_notch_aims_the_picture_and_leaves_it_where_it_is() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous = take_pin_for_a_test();

    let content = (100, 100, 500, 400);
    let pin = overlay_pin(content, PinChrome::always());
    let drawn = pin.zoom.expect("a picture the wheel zooms").drawn;
    stand_pin(Some(pin));

    assert!(
        !pin_zoom_in_flight(),
        "a picture nothing has been asked of is going nowhere"
    );
    assert_eq!(
        pin_zoom_oversize_box(),
        None,
        "and a picture the size of its window is not a picture a frame is owed to"
    );

    aim_pin_zoom(1);

    let zoom = pin_state()
        .and_then(|pinned| pinned.pin().and_then(|pin| pin.zoom))
        .expect("the pin this test stood up");
    assert_eq!(
        zoom.drawn, drawn,
        "a notch writes down where the picture is going and does not move it"
    );
    assert!(
        (zoom.toward - drawn.0 * PIN_ZOOM_PER_NOTCH).abs() < 0.01,
        "and what it wrote down is one notch further out: {} of {}",
        zoom.toward,
        drawn.0 * PIN_ZOOM_PER_NOTCH
    );
    assert!(
        pin_zoom_in_flight(),
        "so there is a travel for the loop to take"
    );

    stand_pin(None);
    stand_pin(previous);
}

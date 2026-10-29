use super::*;

// What the wheel hook is handed a message as: the delta in the high word of `mouseData`, the pointer in
// its own point, and a code that says the message is one to look at.
use windows::core::w;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, POINT, RECT, WAIT_OBJECT_0, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DestroyWindow, GetWindowRect, IsIconic, SendMessageW, ShowWindow, HC_ACTION,
    MSLLHOOKSTRUCT, SC_MINIMIZE, SW_HIDE, SW_SHOWNOACTIVATE, WINDOW_EX_STYLE, WM_MOUSEWHEEL,
    WM_SYSCOMMAND, WS_POPUP,
};

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
    notches: f32,
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
    let mut plan = zoomed(a_picture(), PinZoom::filling(a_picture()), 1.0, None);
    let mut content = plan.content;
    let mut zoom = plan.zoom;

    for _ in 0..40 {
        plan = zoomed(content, zoom, 1.0, Some(pointer));
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
    let plan = zoomed(a_picture(), PinZoom::filling(a_picture()), 1.0, None);
    assert_eq!(plan.content, (280, 235, 720, 565));
    assert_eq!(plan.zoom.drawn, (440.0, 330.0));
    assert_eq!(plan.zoom.pan, (0.0, 0.0));
    assert_eq!(
        ((plan.content.2 - plan.content.0) as f64) / ((plan.content.3 - plan.content.1) as f64),
        4.0 / 3.0,
        "the window keeps the picture's shape"
    );

    // Three notches in one gesture are the same step taken further rather than three boxes: a wheel
    // turned fast lands on the same size a wheel turned slowly would have arrived at. The ideal box is
    // 532.4 by 399.3, and a window is whole pixels: its four edges are that rect rounded once each, so
    // the box is 532 across — the fraction under a pixel is given up — and 400 down, the pixel the
    // outward rounding costs the height. Both exact: what a fast wheel lands on is the box a slow one
    // arrives at, to the pixel.
    let three = zoomed(a_picture(), PinZoom::filling(a_picture()), 3.0, None);
    assert_eq!(three.content.2 - three.content.0, 532);
    assert_eq!(three.content.3 - three.content.1, 400);
}

#[test]
fn one_notch_down_shrinks_it_by_exactly_what_a_notch_up_grew_it_by() {
    // Out is in divided rather than a step of its own, which is the whole of what lets a hand come
    // back from either end without counting: the same notch back lands on the box it started from.
    let grown = zoomed(a_picture(), PinZoom::filling(a_picture()), 1.0, None);
    let back = zoomed(grown.content, grown.zoom, -1.0, None);

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
    let plan = zoomed(huge, PinZoom::filling(huge), 1.0, Some((500, 400)));

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
    let out = zoomed(huge, PinZoom::filling(huge), -1.0, None);
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
        let plan = zoomed(content, zoom, -1.0, None);
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
    let plan = zoomed(window, zoom, 1.0, Some(pointer));

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
    let middle = zoomed(window, zoom, 1.0, Some((500, 400)));
    assert_eq!(middle.zoom.pan, (0.0, 0.0));
}

#[test]
fn a_notch_inside_the_window_never_pulls_a_picture_it_was_covering_off_one() {
    // A hand is always inside the window it turns the wheel in, because a notch is counted for a
    // pointer inside the pin's own box and for nothing else, and what such a hand is owed is a
    // picture that goes on covering its window. The notches are turned one at a time because that is
    // the shape a hand turns them in: the window grows with the picture a frame at a time, so the
    // notch that takes the picture past the cap is a notch the window has already nearly reached.
    for pointer in [
        (300, 250),
        (699, 549),
        (300, 549),
        (699, 250),
        (500, 400),
        (400, 300),
    ] {
        let mut plan = zoomed(
            a_picture(),
            PinZoom::filling(a_picture()),
            1.0,
            Some(pointer),
        );

        for _ in 0..11 {
            plan = zoomed(plan.content, plan.zoom, 1.0, Some(pointer));
        }

        let box_ = plan.zoom.picture_box(region_size(plan.content));

        assert!(
            box_.0 <= 0 && box_.1 <= 0 && box_.2 >= 1 && box_.3 >= 1,
            "a picture covers its window at the cap, pointed at {pointer:?}: {box_:?}"
        );
    }
}

#[test]
fn a_notch_that_grows_the_window_with_the_picture_keeps_the_pointer_anchored() {
    // The one notch where the two anchors are asked for at once: the picture is taken past the cap by
    // the same notch that takes the window out to it, so the window grows about the picture's middle
    // while the picture grows about the pointer. What the pointer is owed is the point of the picture
    // under it — the share of the picture to its left does not move — and the window's own growth is
    // not a reason to move it: an anchor measured in the window's shares instead would let the point
    // under the hand drift by the window's growth, which is four pixels of a picture this size.
    let window = (5, 4, 995, 796);
    let zoom = PinZoom::filling(window);
    let pointer = (50, 400);
    let under_the_pointer = |content: ScreenRegion, zoom: &PinZoom| {
        let (left, _, _, _) = zoom.picture_box(region_size(content));

        (pointer.0 - (content.0 + left)) as f32 / zoom.drawn.0
    };

    let before = under_the_pointer(window, &zoom);
    let plan = zoomed(window, zoom, 1.0, Some(pointer));

    assert_eq!(
        plan.content,
        (0, 0, 1000, 800),
        "the window is taken out to the cap"
    );
    assert_eq!(plan.zoom.picture_box((1000, 800)), (1, -35, 1090, 836));
    assert!(
        (under_the_pointer(plan.content, &plan.zoom) - before).abs() < 0.002,
        "the point of the picture under the pointer stays under it: {} against {before}",
        under_the_pointer(plan.content, &plan.zoom)
    );
}

#[test]
fn a_notch_over_the_backdrop_keeps_the_point_of_the_backdrop_under_the_pointer() {
    // A picture half again as large as its window and carried most of the way off it: the window is
    // 0..1000 across and the picture is drawn from 450, so the backdrop — the `Image -> Zoom ->
    // Background` colour, black by a fresh configuration — is showing down the window's left side,
    // and the hand is over it.
    let window = (0, 0, 1000, 800);
    let carried = PinZoom::still((1500.0, 1200.0), (700.0, 0.0), centred_on(window));
    let pointer = (200, 400);

    assert_eq!(
        carried.picture_box((1000, 800)),
        (450, -200, 1950, 1000),
        "the picture is drawn with the backdrop beside it"
    );

    let plan = zoomed(window, carried, 1.0, Some(pointer));

    // The point of the backdrop under the hand is not the picture's to keep, so what holds under it is
    // its *place in the picture*: a hand 250 to the left of the picture's left edge is 275 to the left
    // of it once the picture has grown by a tenth. What the wheel may not do is pull the picture back
    // over the window, which is what a placement bounded by the overhang does (see `pin_zoom_to`).
    assert_eq!(plan.content, window, "the window is at its cap and holds");
    assert_eq!(plan.zoom.drawn, (1650.0, 1320.0));
    assert_eq!(
        plan.zoom.picture_box((1000, 800)),
        (475, -260, 2125, 1060),
        "the backdrop under the pointer keeps its width, and the picture a hand carried off stays off"
    );
}

#[test]
fn a_travel_over_the_backdrop_keeps_the_point_of_the_picture_under_its_anchor() {
    // The same as the travel at the cap, for a picture the hand has carried off its window: each frame
    // re-places the picture from where the last one left it, so a frame that bounded the placement by
    // the window's overhang would walk the picture back over the window a little at a time — a picture
    // that creeps across the hand as it is zoomed rather than one that follows it.
    let window = (0, 0, 1000, 800);
    let carried = PinZoom::still((1500.0, 1200.0), (700.0, 0.0), centred_on(window));
    let anchor = (200, 400);
    let toward = carried.drawn.0 * PIN_ZOOM_PER_NOTCH.powi(6);
    let asking = |zoom| PinZoomInputs {
        content: window,
        zoom,
        dpi: DESK,
        room: bounds(),
        pointer: Some(anchor),
    };

    // Where in the picture the point under the anchor is, as a share of the picture's width: negative,
    // since the point is off the picture's own left edge — which is the whole of what a hand over the
    // backdrop has to be anchored on.
    let under_the_anchor = |zoom: &PinZoom| {
        let (left, _, _, _) = zoom.picture_box((1000, 800));

        (anchor.0 - left) as f32 / zoom.drawn.0
    };

    let mut zoom = carried;
    zoom.toward = toward;
    let before = under_the_anchor(&zoom);
    assert!(
        before < -0.1,
        "the point the wheel read is well beside the picture: {before}"
    );

    let mut frames = 0;

    while let Some(plan) = pin_zoom_step(&asking(zoom), 0.016, &display()) {
        zoom = plan.zoom;
        frames += 1;
        assert!(
            (under_the_anchor(&zoom) - before).abs() < 0.005,
            "frame {frames} moved the picture under the anchor: {} against {before}",
            under_the_anchor(&zoom)
        );
        assert!(frames < 120, "a travel of two seconds is not a travel");
    }

    assert!(
        zoom.drawn.0 > carried.drawn.0 * 1.5,
        "and the picture did travel: {} of {toward}",
        zoom.drawn.0
    );
}

#[test]
fn zooming_back_out_of_the_window_consumes_the_overhang_before_it_shrinks_it() {
    // A picture half as large again as its window, carried a little. A notch out takes a tenth off
    // the *picture*, so the window holds still while the overhang is spent — and the moment the
    // picture is the window's own size again the window starts coming down with it and the pan is
    // given up, because a picture that size has nowhere to sit.
    let window = (0, 0, 1000, 800);
    let inside = PinZoom::still((1500.0, 1200.0), (-80.0, 40.0), centred_on(window));

    let first = zoomed(window, inside, -1.0, Some((300, 300)));
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
        let plan = zoomed(content, zoom, -1.0, Some((300, 300)));
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
fn a_window_shrinking_back_out_follows_itself_and_not_the_panned_picture() {
    // A picture half as large again as its window, carried a long way off it: the window is 0..1000
    // across and the picture is drawn from 450, so the backdrop — the `Image -> Zoom -> Background`
    // colour, black by a fresh configuration — is showing down the window's left side. Zooming out
    // shrinks the window back toward the picture's size, and where that smaller window lands is
    // about its own middle: a window re-placed about the picture's carried centre would walk across
    // the screen with the pan, following the image inside the window instead of the window itself.
    let window = (0, 0, 1000, 800);
    let carried = PinZoom::still((1500.0, 1200.0), (700.0, 0.0), centred_on(window));
    assert_eq!(
        carried.picture_box((1000, 800)),
        (450, -200, 1950, 1000),
        "the backdrop is showing down the window's left side"
    );

    // Notches out spend the overhang first — the window holds while the picture is larger than it —
    // so the wheel is turned until the window itself starts coming down, the hand over the backdrop
    // the whole way, as a hand zooming out of a pan is.
    let (mut content, mut zoom) = (window, carried);
    for _ in 0..40 {
        let plan = zoomed(content, zoom, -1.0, Some((200, 400)));
        content = plan.content;
        zoom = plan.zoom;
        if content != window {
            break;
        }
    }
    assert_ne!(
        content, window,
        "the window comes down once the picture fits it"
    );
    let (before, after) = (
        ((window.0 + window.2) / 2, (window.1 + window.3) / 2),
        ((content.0 + content.2) / 2, (content.1 + content.3) / 2),
    );
    assert_eq!(
        after, before,
        "the smaller window sits about the window's own middle and not the carried picture's: \
         {after:?} against {before:?}"
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
        0.0,
        &display(),
    )
    .is_none());
}

#[test]
fn a_picture_is_drawn_on_the_whole_of_a_window_it_fills_and_inside_one_it_does_not() {
    // Scale one is the ordinary pinned picture: the box the media lands in is the whole band, which is
    // the case every repaint of everything but a zoomed picture goes through.
    let window = (0, 0, 1000, 800);
    let filling = PinZoom::filling(window);
    assert_eq!(filling.picture_box((1000, 800)), (0, 0, 1000, 800));

    // And twice the window's size is a picture centred until a hand carries it: half of the overhang
    // hangs off each side.
    let inside = PinZoom::still((2000.0, 1600.0), (0.0, 0.0), centred_on(window));
    assert_eq!(inside.picture_box((1000, 800)), (-500, -400, 1500, 1200));

    // A pan moves it the way the hand went: a pointer carried to the right shows the left of the
    // picture, so the box steps left.
    let carried = PinZoom::still((2000.0, 1600.0), (120.0, -60.0), centred_on(window));
    assert_eq!(carried.picture_box((1000, 800)), (-380, -460, 1620, 1140));

    // A picture with no room to be moved has nowhere to be but the whole window, whatever the pan
    // remembers of a bigger one.
    let back = PinZoom::still((1000.0, 800.0), (400.0, 400.0), centred_on(window));
    assert_eq!(back.picture_box((1000, 800)), (0, 0, 1000, 800));
}

/// A picture the wheel has not taken past its window *is* its window, whatever whole pixel the box's
/// outward rounding lands on.
///
/// The box a travel places is its ideal rect rounded once per edge — `floor` or `ceil` of the ideal
/// width against the `round` the picture's own whole pixels are — so the box can be a pixel narrower
/// than the picture it holds. Read against those two numbers, a picture under the cap was a picture
/// *carried* on every frame whose ideal size fell in the wrong fraction of a pixel, and a travel
/// under the cap crosses it hundreds of times: the road flipped between the band scaled into the
/// window and the picture carried inside it several times per pixel of travel (see
/// `PinZoom::picture_box` and the place beside it).
#[test]
fn a_picture_under_the_cap_is_the_whole_band_of_its_window_on_every_frame() {
    // Two shapes, and both parities of the middle a box is seeded on: a 400 by 300 window's middle is
    // a whole pixel and a 401 by 301 one's is a half.
    for content in [
        (300, 250, 700, 550),
        (300, 250, 701, 551),
        (300, 250, 620, 430),
        (300, 250, 700, 475),
    ] {
        let asking = PinZoomInputs {
            content,
            zoom: PinZoom::filling(content),
            dpi: DESK,
            room: bounds(),
            pointer: None,
        };
        let cap = pin_zoom_cap(&asking);
        let (start, _) = region_size(content);
        let mut drawn = start as f32;
        let mut narrowed = 0;

        // A travel the wheel has not taken past the cap, a twentieth of a pixel a frame: the ideal
        // width sweeps every fraction of a pixel there is, and the box the rounding gives it is the
        // whole of what each frame is asked about.
        while drawn < cap * 0.95 {
            let plan = pin_zoom_to(&asking, drawn, &display());
            let (width, height) = region_size(plan.content);
            let (drawn_width, drawn_height) = plan.zoom.drawn_size();

            assert_eq!(
                plan.zoom.picture_box((width, height)),
                (0, 0, width, height),
                "frame {drawn} of a {cap} cap: a picture the wheel has not taken past its window is \
                 the band of it — {width} by {height} holding {:?} — and this one took the carried road",
                plan.zoom.drawn
            );
            assert_eq!(
                plan.zoom.picture_place((width, height)),
                (0.0, 0.0, width as f32, height as f32),
                "frame {drawn}: the place the sampler is handed is the same band — {:?} in {width} \
                 by {height} — and this one took the carried road",
                plan.zoom.picture_place((width, height))
            );

            // And the walk reaches the frames this is about: the picture's own whole pixels one wider
            // than the box the outward rounding gave it, which is the pixel the two roads disagreed
            // over. A walk that reached none of them would be no test of this at all.
            if drawn_width > width || drawn_height > height {
                narrowed += 1;
            }

            drawn += 0.05;
        }

        assert!(
            narrowed > 0,
            "the walk crosses the pixel the box's outward rounding takes away: {content:?}"
        );
    }
}

#[test]
fn a_frame_the_paint_fills_is_not_a_frame_the_pan_gate_carries() {
    // The picture is at the cap, where the window has stopped and the picture has not: a walk a
    // twentieth of a pixel a frame takes the picture's own whole pixels one wider than the window it
    // stands in on those frames whose ideal size falls in that pixel — and a picture within that
    // pixel of its window *is* its window, so the paint draws the whole band and a pan has nowhere to
    // put it.
    //
    // The pan gate compared those whole pixels against the band without the pixel of tolerance, so on
    // those frames it let the pointer through and computed a pan: a pan carried in the state on a
    // frame the paint ignored, and used by the first frame after it where the picture really does
    // overhang. The gate asks the paint's own question now (see `fills`), so a frame the paint fills
    // is a frame with no pan at all.
    let window = (0, 25, 1000, 775);
    let mut content = window;
    let mut zoom = PinZoom::still((1000.0, 750.0), (0.0, 0.0), centred_on(window));
    let pointer = Some((700, 500));
    let mut overhanging = 0;

    for step in 0..=200 {
        let plan = pin_zoom_to(
            &PinZoomInputs {
                content,
                zoom,
                dpi: DESK,
                room: bounds(),
                pointer,
            },
            1000.0 + step as f32 * 0.05,
            &display(),
        );

        let (width, height) = region_size(plan.content);
        if plan.zoom.picture_box((width, height)) == (0, 0, width, height) {
            assert_eq!(
                plan.zoom.pan,
                (0.0, 0.0),
                "frame {:?} put the picture on the whole band of a {width} by {height} window and \
                 carried it by the pointer as well: {:?}",
                plan.zoom.drawn,
                plan.zoom.pan
            );
        } else {
            overhanging += 1;
        }

        content = plan.content;
        zoom = plan.zoom;
    }

    assert!(
        overhanging > 0,
        "the walk goes on past the pixel of tolerance into the carried road"
    );
}

#[test]
fn a_box_laid_out_again_puts_a_zoomed_picture_back_to_filling_it() {
    // Every road that chooses a box for the whole of a window — a walk, a maximize, a double-click, an
    // edge dragged — owes the picture back to the size of that box, and owes it to nothing else: a
    // kind the wheel does not zoom has no view to reset and must not be given one.
    let mut pin = PinnedPreview {
        content: (100, 100, 500, 400),
        zoom: Some(PinZoom::still(
            (1600.0, 1200.0),
            (200.0, -100.0),
            centred_on((100, 100, 500, 400)),
        )),
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
    let huge = zoomed(
        window,
        PinZoom::filling(window),
        i32::MAX as f32,
        Some((500, 400)),
    );
    assert_eq!(huge.zoom.drawn.0, PIN_ZOOM_CEILING);
    assert!(
        huge.zoom.picture_box((1000, 800)).2 < i32::MAX,
        "and the box it lands in is a rectangle with room either side of it"
    );

    let tiny = zoomed(window, PinZoom::filling(window), i32::MIN as f32, None);
    let (width, height) = region_size(tiny.content);
    assert!(
        width >= 10 && height >= 8,
        "a notch out that no hand can turn stops at the floor rather than going past it: {width} by {height}"
    );
}

#[test]
fn a_wheel_notch_over_a_pin_belongs_to_the_pin_while_the_pointer_is_on_it() {
    // The lock every test that reaches the pin's own state takes, and the pin this one stands for the
    // hook to find: what the hook asks is the pin's window box, which is process-wide.
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous = take_pin_for_a_test();

    let content = (100, 100, 500, 400);
    stand_pin(Some(overlay_pin(content, PinChrome::always())));

    // Which window the pin's keys and wheel are answered for is published by the loop once a tick,
    // and which window the user is in is asked of Windows rather than remembered: a notch belongs to
    // the pin only while the foreground window is the pin's own. The pin is therefore stood up as
    // whatever window is in front here. A desktop with nothing in front has no window for a pin to
    // be, and a test with no pin in front has nothing to ask — the caret's own arithmetic is covered
    // in `key_input`'s tests (see `pin_owns_caret`).
    let foreground = unsafe { GetForegroundWindow() };
    if foreground.0.is_null() {
        stand_pin(None);
        stand_pin(previous);
        return;
    }
    crate::shell::key_input::publish_pin_key_owner(foreground.0 as u64, 0);

    // The count and the signal are the process's own, so both are cleared here: what is read below has
    // to be this notch's and not one left behind by a test beside this one.
    crate::shell::wheel_input::publish_preview_under_the_pointer(false);
    let _ = wheel_input::take_pin_zoom_delta();
    unsafe {
        let _ = ResetEvent(preview_channel_event());
    }

    // A wheel message the way the system hands it over: a low-level hook carries the delta in the high
    // word of `mouseData` and the pointer in its own point, and the point here is inside the pin's box.
    let info = MSLLHOOKSTRUCT {
        pt: POINT {
            x: content.0 + 10,
            y: content.1 + 10,
        },
        mouseData: (WHEEL_DELTA as u32) << 16,
        flags: 0,
        time: 0,
        dwExtraInfo: 0,
    };

    unsafe {
        let _ = wheel_input::wheel_hook_proc(
            HC_ACTION as i32,
            WPARAM(WM_MOUSEWHEEL as usize),
            LPARAM(&info as *const MSLLHOOKSTRUCT as isize),
        );
    }

    assert_eq!(
        wheel_input::take_pin_zoom_delta(),
        WHEEL_DELTA,
        "a notch over a pinned picture is the pin's and not Explorer's"
    );

    // And the loop that reads that count is woken where it is counted: the count is an atomic written
    // on the hook's own thread, so a loop waiting for a pin would otherwise read it at the end of its
    // own wait — 62 ms measured, which is four frames of a 144 Hz display and the whole of what makes
    // a zoom start unevenly (see `wake_preview_loop`).
    unsafe {
        assert_eq!(
            WaitForSingleObject(preview_channel_event(), 0),
            WAIT_OBJECT_0,
            "and the wait that would have read it is ended by the notch itself"
        );
    }

    // And the same notch over the same box belongs to whoever the user is working in the moment the
    // pin is not the window they are in. A pin behind another program is the whole of what
    // `Pin Mode → Always On Top` off asks for, and a wheel turned there is not this app's to take:
    // it is the program in front that is being scrolled, and taking it would leave that program
    // unscrolled and zoom a window the user cannot even see.
    let _ = wheel_input::take_pin_zoom_delta();
    unsafe {
        let _ = ResetEvent(preview_channel_event());
    }
    crate::shell::key_input::publish_pin_key_owner((foreground.0 as u64) ^ 0x1000, 0);

    let handed_on = unsafe {
        wheel_input::wheel_hook_proc(
            HC_ACTION as i32,
            WPARAM(WM_MOUSEWHEEL as usize),
            LPARAM(&info as *const MSLLHOOKSTRUCT as isize),
        )
    };

    assert_eq!(
        wheel_input::take_pin_zoom_delta(),
        0,
        "a pin the user is not in does not take the wheel from the program they are working in"
    );
    assert_ne!(
        handed_on,
        LRESULT(1),
        "and the notch is handed on rather than swallowed, so that program scrolls"
    );
    unsafe {
        assert_ne!(
            WaitForSingleObject(preview_channel_event(), 0),
            WAIT_OBJECT_0,
            "and the loop that would zoom the pin is not woken for a notch that is not its own"
        );
    }

    // A pin that is `Pin Mode → Always On Top` owns the notch from the hover alone, whichever
    // window the keyboard is in: it stands over every other window, so the box the preview
    // published is the box the pointer is over, and the click that gate used to need was a click
    // a topmost pin never needed (see `publish_pin_always_on_top`).
    let _ = wheel_input::take_pin_zoom_delta();
    unsafe {
        let _ = ResetEvent(preview_channel_event());
    }
    crate::shell::key_input::publish_pin_always_on_top(true);

    let swallowed = unsafe {
        wheel_input::wheel_hook_proc(
            HC_ACTION as i32,
            WPARAM(WM_MOUSEWHEEL as usize),
            LPARAM(&info as *const MSLLHOOKSTRUCT as isize),
        )
    };

    assert_eq!(
        wheel_input::take_pin_zoom_delta(),
        WHEEL_DELTA,
        "a pin that is always on top is the window the pointer is in"
    );
    assert_eq!(
        swallowed,
        LRESULT(1),
        "and the notch is the pin's rather than what stands behind it"
    );
    crate::shell::key_input::publish_pin_always_on_top(false);

    // And the same notch is the pin's while the pin is the surface the pointer is standing on,
    // which is the case the keyboard's question gets wrong: a pin raised over the listing the user
    // is working in has the pointer over it and the keyboard in the listing behind it, and a notch
    // turned there is the pin's rather than the list's. It is asked of the desktop once a tick and
    // published, because a low-level hook may not walk the desktop itself (see
    // `PREVIEW_UNDER_POINTER`).
    crate::shell::wheel_input::publish_preview_under_the_pointer(true);
    let _ = wheel_input::take_pin_zoom_delta();
    unsafe {
        let _ = ResetEvent(preview_channel_event());
    }

    let swallowed = unsafe {
        wheel_input::wheel_hook_proc(
            HC_ACTION as i32,
            WPARAM(WM_MOUSEWHEEL as usize),
            LPARAM(&info as *const MSLLHOOKSTRUCT as isize),
        )
    };

    assert_eq!(
        wheel_input::take_pin_zoom_delta(),
        WHEEL_DELTA,
        "a pin the pointer is on is scrolled without the click that used to be needed to reach it"
    );
    assert_eq!(
        swallowed,
        LRESULT(1),
        "and the notch is not left for the listing to scroll behind the pin"
    );
    crate::shell::wheel_input::publish_preview_under_the_pointer(false);
    stand_pin(None);
    stand_pin(previous);
    crate::shell::key_input::publish_pin_key_owner(0, 0);
}

/// A wheel that reports less than a notch per message is counted by the unit rather than by the notch:
/// a high-resolution wheel — and the smoothing driver a plain one can be given — sends a fraction of a
/// notch at a time, and a count rounded to whole notches throws two of every three of those away and
/// turns the third into a whole notch, which is a zoom that moves in lurches of a tenth of a picture
/// rather than in the thirds it was asked for.
///
/// The three thirds are asked of the size the wheel has already asked for rather than of the size the
/// picture has reached, which is the same rule a whole notch is counted by: two hands turning the same
/// amount come out at the same size whatever the frames between them (see `pin_zoom_toward`).
#[test]
fn a_wheel_that_reports_less_than_a_notch_at_a_time_is_counted_by_the_unit() {
    let window = (0, 0, 1000, 800);
    let settled = PinZoom::filling(window);
    let asked = |zoom| PinZoomInputs {
        content: window,
        zoom,
        dpi: DESK,
        room: bounds(),
        pointer: None,
    };

    // A wheel of 40 units against the 120 of a notch, sent as three messages.
    let mut zoom = settled;
    for _ in 0..3 {
        zoom.toward = pin_zoom_toward(&asked(zoom), 40.0 / 120.0);
    }

    assert!(
        (zoom.toward - pin_zoom_toward(&asked(settled), 1.0)).abs() < 0.01,
        "three thirds of a notch are a notch: {} of {}",
        zoom.toward,
        pin_zoom_toward(&asked(settled), 1.0)
    );
    assert!(
        zoom.toward > settled.toward * 1.05,
        "and none of the three was thrown away: {} of {}",
        zoom.toward,
        settled.toward
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

    let first = pin_zoom_toward(&asked(settled), 1.0);

    // The picture part of the way to the size it was asked for, which is the state a second notch
    // arrives in whenever a hand turns the wheel faster than the picture can follow.
    let mut mid = settled;
    mid.drawn = (settled.drawn.0 * 1.05, settled.drawn.1 * 1.05);
    mid.toward = first;

    let second = pin_zoom_toward(&asked(mid), 1.0);
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
        (second - pin_zoom_toward(&asked(settled), 2.0)).abs() < 0.01,
        "two notches are the same size whether they were taken apart or together"
    );
}

/// A wheel's ask is *released* over the frames that follow the notch rather than landing at once: a
/// notch landed whole is a jump, and what a hand turning a wheel is looking at is a movement — and one
/// that goes on after the hand has stopped, which is what is left of the ask when the last notch has
/// been turned.
///
/// What is asserted is the curve in *seconds*: nine tenths of the ask within a tenth of a second, no
/// frame past the ask, and the same release whatever the frames it was taken in. The last of those is
/// the whole of what a share of the remainder buys and a share of the ask would not (see
/// `pin_zoom_release`).
#[test]
fn a_wheels_ask_is_released_over_the_frames_that_follow_it() {
    let (from, toward) = (1000.0f32, 1100.0f32);
    let of_the_ask = |released: f32| (released - from) / (toward - from);

    // Fourteen frames of seven milliseconds is not yet a tenth of a second's worth of release, and the
    // frame that crosses it is nine tenths of the ask.
    let mut released = from;
    for _ in 0..14 {
        released = pin_zoom_release(released, toward, 0.007);
        assert!(
            released >= from && released <= toward,
            "a release may only travel the way it was asked to: {released}"
        );
    }
    assert!(
        of_the_ask(released) < 0.9,
        "a tenth of a second is where nine tenths of the ask lands, and this is not there yet: {}",
        of_the_ask(released)
    );

    released = pin_zoom_release(released, toward, 0.007);
    assert!(
        of_the_ask(released) >= 0.9,
        "and the frame past it is: {}",
        of_the_ask(released)
    );

    // The same curve in whatever frames it is taken: two frames of half the length are one frame of the
    // whole, which is what a machine running at twice the frame rate has to come out at.
    let once = pin_zoom_release(from, toward, 0.016);
    let twice = pin_zoom_release(pin_zoom_release(from, toward, 0.008), toward, 0.008);
    assert!(
        (once - twice).abs() < 0.001,
        "a machine that takes its frames twice as often takes the same release: {once} against {twice}"
    );

    // And the travel is a travel of *some* frames rather than a jump: a frame of the loop's own cadence
    // brings the picture nearer to the ask, and the frames stop once it is there.
    let mut frames = 0;
    let mut released = from;
    while released != toward {
        released = pin_zoom_release(released, toward, 0.016);
        frames += 1;
        assert!(
            frames < 120,
            "a travel of two seconds is not a travel: {released}"
        );
    }
    assert!(
        frames > 3,
        "a step is a travel of some frames and not a jump: {frames}"
    );

    // And a release that has all but arrived *has* arrived, or the loop would keep the frame cadence for
    // a hundredth of a pixel for the rest of the pin's life.
    assert_eq!(pin_zoom_release(toward - 0.1, toward, 0.007), toward);
}

#[test]
fn an_image_is_brought_home_by_the_zoom_out_that_reaches_the_fit() {
    // A picture half as large again as its window, carried most of the way off it: the picture is drawn
    // from x=450, so the backdrop — black by a fresh configuration — is showing down the window's left
    // side, and the hand turns the wheel out from there. What the image inside the window is owed is the
    // place the fit would put it in *before* the size at which it has nowhere else to sit: the last of
    // the picture's overhang is where the offset a hand left is given up, so the moment the picture is
    // the window's own size there is nothing left to drop at once — which, without this, is the whole of
    // the pan in a single frame.
    let window = (0, 0, 1000, 800);
    let carried = PinZoom::still((1500.0, 1200.0), (700.0, 0.0), centred_on(window));
    let asking = |zoom| PinZoomInputs {
        content: window,
        zoom,
        dpi: DESK,
        room: bounds(),
        pointer: Some((200, 400)),
    };

    // The wheel is turned out one notch at a time, each notch walked to its end, which is the shape a
    // hand turns it in and the shape the loop draws it in. What is kept is the last frame at which the
    // picture still *overhangs* its window: the frame the offset a hand left has to be gone by, if the
    // fit is not to be reached by dropping the whole of it at once.
    let mut content = window;
    let mut zoom = carried;
    let mut last_overhang = None;
    let mut notches = 0;

    while region_size(content).0 >= 1000 && notches < 20 {
        zoom.toward = zoom.drawn.0 / PIN_ZOOM_PER_NOTCH;
        let mut frames = 0;

        while zoom.drawn.0 != zoom.toward {
            let plan = pin_zoom_step(&asking(zoom), 0.016, &display())
                .expect("a picture on its way to a size is a frame to take");
            let was = zoom.drawn.0;

            let (drawn_width, drawn_height) = plan.zoom.drawn_size();
            let (window_width, window_height) = region_size(plan.content);

            // The last frame at which the *pan is in play at all*: a picture whole pixels larger than
            // its window is the only picture `picture_box` has anywhere else to put than the whole of
            // it, which is why a picture a fraction of a pixel larger is no use to this test.
            if drawn_width > window_width || drawn_height > window_height {
                last_overhang = Some((plan.content, plan.zoom));
            }

            zoom = plan.zoom;
            content = plan.content;
            frames += 1;
            assert!(frames < 240, "a travel is not forever");
            assert!(
                zoom.drawn.0 < was,
                "the wheel was turned out and the picture comes down: {} of {was}",
                zoom.drawn.0
            );
        }

        notches += 1;
    }

    // The first box the wheel puts the window in once the picture is under the fit: the ideal width is
    // 1500 / 1.1^5 = 931.4, and the box is the whole pixels that rect covers, so it is 932 of them.
    assert_eq!(
        region_size(content).0,
        932,
        "the wheel has taken the picture under the fit"
    );

    let (content, zoom) = last_overhang.expect("the wheel was turned at least once");
    let (width, height) = region_size(content);
    let (left, top, right, bottom) = zoom.picture_box((width, height));

    assert!(
        left <= 0 && top <= 0 && right >= width && bottom >= height,
        "the last frame the picture overhangs its window is already where the fit would put it, with no \
         backdrop left showing: {left}..{right} of {width} against a picture {:?} and an offset {:?}",
        zoom.drawn,
        zoom.pan
    );
    assert!(
        zoom.pan.0.abs() < 1.0 && zoom.pan.1.abs() < 1.0,
        "which is the offset a hand left given up by then, and not dropped at the fit: {:?} of 700",
        zoom.pan
    );
}

static PIN_ZOOM_TEST_MUTEX: Mutex<()> = Mutex::new(());

#[test]
fn the_image_swings_past_its_place_and_comes_back() {
    let _guard = PIN_ZOOM_TEST_MUTEX.lock().unwrap();
    // The bounce: the image is brought home over the last of the overhang, and the swing is laid over
    // that — past the place it is being brought to and back again, so that the image lands rather than
    // stops. What it swings past is the *image's* own offset: the window's box is the wheel's own
    // through every frame of it (see `pin_zoom_to`).
    let window = (0, 0, 1000, 800);
    let asking = |zoom| PinZoomInputs {
        content: window,
        zoom,
        dpi: DESK,
        room: bounds(),
        pointer: Some((200, 400)),
    };
    let mut zoom = PinZoom::still((1100.0, 880.0), (451.0, 0.0), centred_on(window));
    zoom.toward = 1000.0;

    // Where the fit would put the picture at the size it stands at: the middle, with what it overhangs
    // by hanging off both sides alike.
    let home = |zoom: &PinZoom| (1000.0 - zoom.drawn.0) / 2.0;
    if let Ok(mut config) = crate::CONFIG.lock() {
        config.pin_zoom_bounce_speed_ms = 20;
    }

    let mut swang = 0.0f32;
    let mut frames = 0;

    while let Some(plan) = pin_zoom_step(&asking(zoom), 0.016, &display()) {
        zoom = plan.zoom;
        let (left, _, _, _) = zoom.picture_box((1000, 800));
        let past = home(&zoom) - left as f32;
        swang = swang.max(past);
        frames += 1;
        assert!(frames < 240, "a travel and its bounce are not forever");
    }

    assert!(
        swang > 8.0,
        "the image is swung past its place by more than a hair: {swang} pixels"
    );
    assert!(
        swang < 100.0,
        "and never far off its window: {swang} pixels past the middle"
    );
    assert!(
        zoom.bounce == 0.0,
        "and the swing is over with, leaving nothing to undo"
    );
    assert!(
        (home(&zoom) - zoom.picture_box((1000, 800)).0 as f32).abs() < 2.0,
        "so it comes to rest in the place the fit would put it"
    );

    if let Ok(mut config) = crate::CONFIG.lock() {
        config.pin_zoom_bounce_speed_ms = crate::config::config::DEFAULT_PIN_ZOOM_BOUNCE_SPEED_MS;
    }
}

/// An intermediate zoom out that leaves the image oversized neither bounces nor gives up its pan:
/// the bounce and the return home belong to reaching the fit itself, so an image zoomed out within
/// its oversized range does not bounce towards center only to bounce right back to its pan.
#[test]
fn intermediate_zoom_out_does_not_bounce_and_zoom_out_to_fit_clamps_and_bounces_cleanly() {
    let _guard = PIN_ZOOM_TEST_MUTEX.lock().unwrap();
    let window = (0, 0, 1000, 800);
    let asking = |zoom| PinZoomInputs {
        content: window,
        zoom,
        dpi: DESK,
        room: bounds(),
        pointer: Some((200, 400)),
    };

    // Picture starts at 1200px (oversized, window is 1000px), panned significantly.
    let mut zoom = PinZoom::still((1200.0, 960.0), (300.0, 50.0), centred_on(window));

    // First notch out: 1200 / 1.10 = 1090.9 > cap (1000.0).
    let next_toward = pin_zoom_toward(&asking(zoom), -1.0);
    assert!(
        next_toward > 1000.0,
        "the first notch out leaves the image still oversized: {next_toward}"
    );
    zoom.toward = next_toward;

    // Walk the travel to 1090.9:
    while zoom.drawn.0 != zoom.toward {
        let plan = pin_zoom_step(&asking(zoom), 0.016, &display()).expect("frame in flight");
        zoom = plan.zoom;
        assert_eq!(
            zoom.bounce, 0.0,
            "intermediate zoom out to oversized never starts a bounce"
        );
    }
    // Pan is preserved:
    assert!(
        zoom.pan.0 > 200.0,
        "pan is not prematurely reduced by home_share while still oversized: {}",
        zoom.pan.0
    );

    // Second notch out: from 1090.9, 1090.9 / 1.10 = 991.7 < cap (1000.0).
    // It clamps to cap (1000.0)!
    let fit_toward = pin_zoom_toward(&asking(zoom), -1.0);
    assert_eq!(
        fit_toward, 1000.0,
        "a zoom out crossing the fit boundary clamps to fit (1000.0) first"
    );
    zoom.toward = fit_toward;

    if let Ok(mut config) = crate::CONFIG.lock() {
        config.pin_zoom_bounce_speed_ms = 20;
    }

    let mut swang = 0.0f32;
    let mut saw_bounce = false;
    let mut frames = 0;

    while let Some(plan) = pin_zoom_step(&asking(zoom), 0.016, &display()) {
        zoom = plan.zoom;
        if zoom.bounce > 0.0 {
            saw_bounce = true;
        }
        let (left, _, _, _) = zoom.picture_box((1000, 800));
        let past = 0.0 - left as f32;
        swang = swang.max(past);
        frames += 1;
        assert!(frames < 240, "a travel and its bounce are not forever");
    }

    assert!(
        saw_bounce,
        "the bounce triggered when heading home to the fit"
    );
    assert!(swang > 4.0, "the bounce swung past center: {swang} pixels");
    assert_eq!(zoom.bounce, 0.0, "bounce finished");
    assert_eq!(zoom.pan, (0.0, 0.0), "pan came home completely to zero");
    assert_eq!(
        zoom.picture_box((1000, 800)),
        (0, 0, 1000, 800),
        "the image settled cleanly at fit to screen without bouncing back"
    );

    if let Ok(mut config) = crate::CONFIG.lock() {
        config.pin_zoom_bounce_speed_ms = crate::config::config::DEFAULT_PIN_ZOOM_BOUNCE_SPEED_MS;
    }
}

#[test]
fn pin_zoom_bounce_duration_respects_config() {
    let _guard = PIN_ZOOM_TEST_MUTEX.lock().unwrap();
    let window = (0, 0, 1000, 800);
    let asking = |zoom| PinZoomInputs {
        content: window,
        zoom,
        dpi: DESK,
        room: bounds(),
        pointer: Some((200, 400)),
    };

    // When configured to 0 ms: no bounce occurs.
    if let Ok(mut config) = crate::CONFIG.lock() {
        config.pin_zoom_bounce_ms = 0;
    }
    let mut zoom = PinZoom::still((1100.0, 880.0), (451.0, 0.0), centred_on(window));
    zoom.toward = 1000.0;
    while let Some(plan) = pin_zoom_step(&asking(zoom), 0.016, &display()) {
        zoom = plan.zoom;
        assert_eq!(zoom.bounce, 0.0, "0 ms config disables bounce");
    }

    // Reset back to default
    if let Ok(mut config) = crate::CONFIG.lock() {
        config.pin_zoom_bounce_ms = crate::config::config::DEFAULT_PIN_ZOOM_BOUNCE_MS;
    }
}

/// When the bounce gate is switched off, the image still smoothly returns to the fit, but skips
/// the overshoot bounce entirely and lands directly at center.
#[test]
fn pin_zoom_bounce_gate_disabled_skips_overshoot() {
    let _guard = PIN_ZOOM_TEST_MUTEX.lock().unwrap();
    let window = (0, 0, 1000, 800);
    let asking = |zoom| PinZoomInputs {
        content: window,
        zoom,
        dpi: DESK,
        room: bounds(),
        pointer: Some((200, 400)),
    };

    if let Ok(mut config) = crate::CONFIG.lock() {
        config.pin_zoom_bounce_gate = false;
    }

    let mut zoom = PinZoom::still((1100.0, 880.0), (451.0, 0.0), centred_on(window));
    zoom.toward = 1000.0;
    let home = |zoom: &PinZoom| (1000.0 - zoom.drawn.0) / 2.0;
    let mut swang = 0.0f32;

    while let Some(plan) = pin_zoom_step(&asking(zoom), 0.016, &display()) {
        zoom = plan.zoom;
        let (left, _, _, _) = zoom.picture_box((1000, 800));
        let past = home(&zoom) - left as f32;
        swang = swang.max(past);
    }

    // Because gate is false, it glides home and stops without swinging past center
    assert!(
        swang < 1.0,
        "when gate is disabled, image never swings past center: swang={swang}"
    );
    assert_eq!(zoom.pan, (0.0, 0.0));
    assert_eq!(zoom.picture_box((1000, 800)), (0, 0, 1000, 800));

    // Reset back
    if let Ok(mut config) = crate::CONFIG.lock() {
        config.pin_zoom_bounce_gate = crate::config::config::DEFAULT_PIN_ZOOM_BOUNCE_GATE;
    }
}

/// Bounce speed and power can be customized via configuration.
#[test]
fn pin_zoom_bounce_speed_and_power_configurable() {
    let _guard = PIN_ZOOM_TEST_MUTEX.lock().unwrap();
    let window = (0, 0, 1000, 800);
    let asking = |zoom| PinZoomInputs {
        content: window,
        zoom,
        dpi: DESK,
        room: bounds(),
        pointer: Some((200, 400)),
    };

    if let Ok(mut config) = crate::CONFIG.lock() {
        config.pin_zoom_bounce_gate = true;
        config.pin_zoom_bounce_speed_ms = 40;
        config.pin_zoom_bounce_power = 0.08;
    }

    let mut zoom = PinZoom::still((1100.0, 880.0), (451.0, 0.0), centred_on(window));
    zoom.toward = 1000.0;
    let home = |zoom: &PinZoom| (1000.0 - zoom.drawn.0) / 2.0;
    let mut swang = 0.0f32;

    while let Some(plan) = pin_zoom_step(&asking(zoom), 0.016, &display()) {
        zoom = plan.zoom;
        let (left, _, _, _) = zoom.picture_box((1000, 800));
        let past = home(&zoom) - left as f32;
        swang = swang.max(past);
    }

    // With power = 0.08 on pan 451.0, overshoot is ~451 * 0.08 = ~36 pixels
    assert!(
        swang > 20.0,
        "higher power yields proportionally larger overshoot: swang={swang}"
    );
    assert_eq!(zoom.pan, (0.0, 0.0));

    // Reset back
    if let Ok(mut config) = crate::CONFIG.lock() {
        config.pin_zoom_bounce_gate = crate::config::config::DEFAULT_PIN_ZOOM_BOUNCE_GATE;
        config.pin_zoom_bounce_speed_ms = crate::config::config::DEFAULT_PIN_ZOOM_BOUNCE_SPEED_MS;
        config.pin_zoom_bounce_power = crate::config::config::DEFAULT_PIN_ZOOM_BOUNCE_POWER;
    }
}

#[test]
fn a_picture_sitting_in_the_middle_neither_bounces_nor_keeps_the_loop_awake() {
    // The bounce is for an image that has somewhere to be brought home *from*: a picture nothing has
    // carried off its window has no offset to give up. Such a picture is the ordinary one — every pin
    // zoomed past the fit to look at something closely, with nothing ever dragged — so what a zoom out of
    // one owes is the motion it always was: no swing, and no frame cadence kept for one (see
    // `pin_zoom_bounce`).
    let window = (0, 0, 1000, 800);
    // A picture overhanging its window with nothing carried off it: everything the ordinary pin zoomed in
    // to look at something closely is, and the state the stretch is crossed from.
    let mut zoom = PinZoom::still((1200.0, 960.0), (0.0, 0.0), centred_on(window));
    let asking = |zoom| PinZoomInputs {
        content: window,
        zoom,
        dpi: DESK,
        room: bounds(),
        pointer: Some((500, 400)),
    };

    assert!(
        zoom.drawn.0 > 1000.0 && zoom.pan == (0.0, 0.0),
        "the picture overhangs its window and sits in the middle of it"
    );
    let mut frames = 0;

    while zoom.drawn.0 > 1000.0 {
        zoom.toward = zoom.drawn.0 / PIN_ZOOM_PER_NOTCH;

        while let Some(plan) = pin_zoom_step(&asking(zoom), 0.016, &display()) {
            zoom = plan.zoom;
            frames += 1;
            assert_eq!(
                zoom.bounce, 0.0,
                "frame {frames} swings an image that has nothing to bring home"
            );
            assert_eq!(zoom.pan, (0.0, 0.0), "and gives up an offset it never had");
            assert!(frames < 240, "a travel is not forever");
        }
    }

    assert!(frames > 1, "the wheel was turned at least once");
}

#[test]
fn a_zoom_in_gives_up_nothing_of_an_offset() {
    // The other half of the rule: a hand turning the wheel *in* is opening room for the picture's offset
    // rather than taking it away, so a notch in over the last of the overhang carries the offset it has
    // — the wheel's own anchoring, whole (see `home_share`).
    let window = (0, 0, 1000, 800);
    let carried = PinZoom::still((1024.0, 819.2), (383.0, 0.0), centred_on(window));
    let asking = |zoom| PinZoomInputs {
        content: window,
        zoom,
        dpi: DESK,
        room: bounds(),
        pointer: Some((500, 400)),
    };

    let mut zoom = carried;
    zoom.toward = zoom.drawn.0 * PIN_ZOOM_PER_NOTCH;
    let mut frames = 0;

    while let Some(plan) = pin_zoom_step(&asking(zoom), 0.016, &display()) {
        zoom = plan.zoom;
        frames += 1;
        assert!(
            zoom.pan.0 > carried.pan.0 * 0.9,
            "frame {frames} gives up an offset a zoom in has no business giving up: {}",
            zoom.pan.0
        );
        assert!(frames < 240, "a travel is not forever");
    }
}

/// A frame of a travel is the wheel's ask *released* and not the ask itself: a notch's own tenth of a
/// picture is carried out over the frames that follow it, which is what leaves the picture still moving
/// after the hand has stopped (see `pin_zoom_release`).
#[test]
fn a_frame_of_a_travel_is_the_wheels_ask_released_and_not_the_ask_itself() {
    let window = (0, 0, 1000, 800);
    let settled = PinZoom::filling(window);
    let toward = settled.drawn.0 * PIN_ZOOM_PER_NOTCH;
    let asking = |zoom| PinZoomInputs {
        content: window,
        zoom,
        dpi: DESK,
        room: bounds(),
        pointer: Some((500, 400)),
    };

    let mut travelling = settled;
    travelling.toward = toward;
    let step = pin_zoom_step(&asking(travelling), 0.007, &display())
        .expect("a picture on its way to a size is a frame to take");

    // A frame of seven milliseconds releases 1 - 0.1^0.07 of what is left of the ask — 14.89 pixels of
    // this notch's hundred, which is where the picture is drawn.
    assert!(
        (step.zoom.drawn.0 - 1014.89).abs() < 0.05,
        "a frame releases a seventh of what is left of the ask: {} of 1014.89",
        step.zoom.drawn.0
    );
    assert_eq!(
        step.zoom.toward, toward,
        "the ask is where the wheel put it"
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

    let mut travelling = cap.zoom;
    travelling.toward = toward;
    let step = pin_zoom_step(&asking(travelling), 0.016, &display())
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
    assert!(step.zoom.drawn.0 < toward, "and still moving");

    assert!(
        pin_zoom_step(&asking(cap.zoom), 0.016, &display()).is_none(),
        "a picture already at the size it was asked for has nothing to carry"
    );
}

/// A travel past the cap zooms about the point the wheel read and about nothing else: the point of the
/// picture under it at the first frame is the point of the picture under it at the last, however many
/// frames the travel took.
///
/// It is the property the whole of the pan arithmetic exists for, and it is the one a travel is most
/// likely to lose: each frame re-anchors the picture from where the last one left it, so a frame that
/// placed the picture by anything other than the same rule would walk the point under the hand away from
/// it a little at a time — which is a picture that creeps sideways as it is zoomed.
#[test]
fn a_travel_past_the_cap_keeps_the_point_of_the_picture_under_its_anchor() {
    // An anchor well off the picture's own middle, so that the picture's share of it is a number a frame
    // that placed it wrongly would move: an anchor at the middle is one a picture can be zoomed about
    // without placing anything at all.
    let anchor = (900, 700);
    let cap = at_the_cap(anchor);
    let toward = cap.zoom.drawn.0 * PIN_ZOOM_PER_NOTCH.powi(6);
    let asking = |zoom| PinZoomInputs {
        content: cap.content,
        zoom,
        dpi: DESK,
        room: bounds(),
        pointer: Some(anchor),
    };

    // Where in the picture the point under the anchor is, as a share of the picture's width: the whole
    // of what "anchored" means is that this number does not move.
    let under_the_anchor = |zoom: &PinZoom| {
        let centre = (cap.content.0 + cap.content.2) / 2 + zoom.pan.0.round() as i32;

        (anchor.0 - centre) as f32 / zoom.drawn.0
    };

    let mut zoom = cap.zoom;
    zoom.toward = toward;
    let before = under_the_anchor(&zoom);
    let mut frames = 0;

    while let Some(plan) = pin_zoom_step(&asking(zoom), 0.016, &display()) {
        zoom = plan.zoom;
        frames += 1;
        assert!(
            (under_the_anchor(&zoom) - before).abs() < 0.005,
            "frame {frames} moved the picture under the anchor: {} against {before}",
            under_the_anchor(&zoom)
        );
        assert!(frames < 120, "a travel of two seconds is not a travel");
    }

    assert!(
        zoom.drawn.0 > cap.zoom.drawn.0 * 1.5,
        "and the picture did travel: {} of {toward}",
        zoom.drawn.0
    );
}

/// A travel's frames place the picture so that the point the wheel read stays under the pointer — in the
/// floats the picture is really *drawn* at, and not only in the arithmetic that decided them.
///
/// The two are not the same question, and the difference is the whole of what an eye sees at the end of a
/// zoom: `drawn` is a float the release walks by a fraction of a pixel a frame, and a picture drawn
/// through a box of whole pixels is a picture on the pixel grid — every time the place crosses a half
/// pixel the whole picture steps a pixel, which a hand that is holding still reads as a jolt *backwards*,
/// since the jolt is the same size whatever the frame's own travel was. Measured on six notches taken at
/// the cap, the point under the wheel stepped a whole pixel back on a frame that had moved a third of one
/// forwards; the place a sampler is handed holds it to a hundredth (see `PinZoom::picture_place`).
#[test]
fn a_travel_keeps_the_point_of_the_picture_under_the_wheel_where_it_is_drawn() {
    for anchor in [(900, 700), (100, 100), (500, 400), (890, 60)] {
        let cap = at_the_cap(anchor);
        let toward = cap.zoom.drawn.0 * PIN_ZOOM_PER_NOTCH.powi(6);
        let asking = |zoom| PinZoomInputs {
            content: cap.content,
            zoom,
            dpi: DESK,
            room: bounds(),
            pointer: Some(anchor),
        };
        let (width, height) = region_size(cap.content);

        // Which point of the picture the wheel read, as a share of the picture's whole width, read out
        // of the floats the place is worked from rather than out of a box of pixels: a share taken from a
        // box that is half a pixel out would spend the whole of this test's own room for being half a
        // pixel out. The pointer's own screen point is asked in the *window's* coordinates, which is where
        // the pan and the place are both measured (see `PinZoom::picture_place`).
        let anchor = (anchor.0 - cap.content.0, anchor.1 - cap.content.1);
        let (left, _, drawn_width, _) = cap.zoom.picture_place((width, height));
        let share = (anchor.0 as f32 - left) / drawn_width;

        let mut zoom = cap.zoom;
        zoom.toward = toward;
        let mut frames = 0;

        while let Some(plan) = pin_zoom_step(&asking(zoom), 0.016, &display()) {
            zoom = plan.zoom;
            let (left, _, drawn_width, _) = zoom.picture_place((width, height));
            let drawn_at = left + share * drawn_width;

            assert!(
                (drawn_at - anchor.0 as f32).abs() <= 0.01,
                "frame {frames} drew the point of the picture the wheel read {} pixels from where the \
                 hand left it",
                drawn_at - anchor.0 as f32
            );
            frames += 1;
            assert!(frames < 120, "a travel of two seconds is not a travel");
        }

        assert!(
            zoom.drawn.0 > cap.zoom.drawn.0 * 1.5,
            "and the picture did travel: {} of {toward}",
            zoom.drawn.0
        );
    }
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

    aim_pin_zoom(1.0);

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

/// A page is taken up as a picture the wheel zooms, because a page is a frame this window draws like
/// any other: what rasterized it — a renderer, a converter, a comic's own reader — is nothing a notch
/// or a box change asks (see `box_is_covered`), and the view it arrives with is its picture filling
/// the window the take-up just laid it out in.
///
/// What the take-up must keep refusing is here beside it, because the difference is the whole of what
/// the gate is for: a document the engine draws is a window of somebody else's behind the band, a
/// film is a player's frames arriving one at a time rather than a picture this window holds, and a
/// card is a layout painted at the box rather than a picture rasterized into it.
#[test]
fn a_page_is_a_picture_the_wheel_zooms_and_a_card_or_an_engines_window_is_not() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    // Every kind goes through the one road: the media in the slot is the kind's, and the pin that
    // comes back is the whole of what the take-up decided — including whether it is zoomable at all.
    let stood = |kind: MediaType| {
        let mut media = create_loading_media(400, 300);
        media.media_type = kind;
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = Some(media);
        }
        take_up_pinned_window(&PathBuf::from("a-pinned-file"), (100, 80, 800, 400))
    };

    for kind in [
        MediaType::StaticImage,
        MediaType::Pdf,
        MediaType::Comic,
        MediaType::Calibre,
        MediaType::Libre,
        MediaType::Office,
    ] {
        let pin = stood(kind);
        assert_eq!(
            pin.zoom,
            Some(PinZoom::filling(pin.content)),
            "{kind:?} is a frame this window draws, so the wheel zooms it: the pin is taken up with \
             its picture filling the window"
        );
    }

    for kind in [
        MediaType::EngineSvg,
        MediaType::EngineFont,
        MediaType::Video,
        MediaType::NativeVideo,
        MediaType::Text,
        MediaType::Archive,
        MediaType::Peazip,
        MediaType::Audio,
    ] {
        let pin = stood(kind);
        assert_eq!(
            pin.zoom, None,
            "{kind:?} is not a still picture this window holds, so there is no view for the wheel \
             to write"
        );
    }

    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    stand_pin(previous_pin);
}

/// A page of text that fits the box it is drawn in is still a page a reader can resize: `Ctrl` and a
/// wheel over a pin is the page's own size, and a page with no scroll left to give Explorer is the
/// one a reader most has a reason to enlarge (see `text_size_keep_alive_try`).
///
/// A page that *does* scroll is reached by the scroll's own region, which is the branch the page's
/// own tests cover (see `text_scroll_keep_alive_try`); what is asked here is only that the second
/// region answers for a pin whose page has nothing to scroll, and that it answers for a page rather
/// than for whatever a pin happens to show.
#[test]
fn a_short_page_is_still_a_page_a_ctrl_wheel_resizes() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_regions = POINTER_HOLD_REGIONS
        .lock()
        .ok()
        .and_then(|mut held| held.take());
    let previous_holding = TEXT_PREVIEW_HOLDING.swap(true, Ordering::AcqRel);

    let content = (100, 100, 500, 400);
    stand_pin(Some(overlay_pin(content, PinChrome::always())));

    // A page of two lines in a box that holds far more than two: it cannot scroll, which is the
    // whole of what this test is about.
    let folder = std::env::temp_dir().join("rust-hover-preview-pin-zoom");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("a-page-that-fits.txt");
    std::fs::write(
        &path,
        "one line
another
",
    )
    .expect("a written page");
    let Some(page) = load_text_preview(&path, 400, 300, 96, current_text_options()) else {
        panic!("a page of text was not loaded")
    };
    assert!(
        !page
            .text_state
            .as_ref()
            .map(|state| state.can_scroll())
            .unwrap_or(true),
        "the page fits the box it was loaded in, so it has no scroll of its own"
    );
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = Some(page);
    }

    // What the loop publishes for a pin: the window's own box, which is what the hook reads.
    if let Ok(mut held) = POINTER_HOLD_REGIONS.lock() {
        *held = Some(vec![content]);
    }

    assert!(
        text_size_keep_alive_try().is_some(),
        "a pin whose page fits its box publishes a region the size's own wheel reads"
    );
    assert!(
        text_scroll_keep_alive_try().is_none(),
        "and the scroll's own region is not it: a page that cannot scroll has nothing to give"
    );

    // And the region is a page's, not a pin's: a pinned picture is zoomed by a `Ctrl` and a wheel
    // rather than resized, and the window it publishes is the same box (see
    // `pin_zoom_keep_alive_try`).
    TEXT_PREVIEW_HOLDING.store(false, Ordering::Release);
    assert!(
        text_size_keep_alive_try().is_none(),
        "a pin that is not showing a page is not a page to resize"
    );

    TEXT_PREVIEW_HOLDING.store(previous_holding, Ordering::Release);
    if let Ok(mut held) = POINTER_HOLD_REGIONS.lock() {
        *held = previous_regions;
    }
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    stand_pin(previous_pin);
}

/// Where this file's windows stand: no display is there, so a run of the suite is no window of the
/// user's and no place the hand can land in.
const NOWHERE: i32 = -32000;

/// The screen box a window is standing at, as the desktop has it rather than as this app meant it.
///
/// Safety: the handle is a window of this test's or of the app's own, and the question is a
/// measurement of it.
unsafe fn window_rect(hwnd: HWND) -> ScreenRegion {
    let mut rect = RECT::default();
    let _ = GetWindowRect(hwnd, &mut rect);
    (rect.left, rect.top, rect.right, rect.bottom)
}

/// A picture pin standing in a window of its own at `box_`, with the hover's window down beside it:
/// what a picture pin is on the screen, and the arrangement the two roads below are asked in.
///
/// The hover's window is a window of the test's own, hidden, because that is the state a picture pin
/// leaves it in — and because it is the window a road that asked the wrong one would answer from.
///
/// The caller holds the pin tests' lock and puts back every global written here: the handle the app
/// reads its preview window by, the pin's own window, the flag saying which window the pin stands in,
/// and the pin itself.
///
/// Safety: both windows are the caller's to destroy.
unsafe fn a_picture_pin_in_a_window_of_its_own(box_: ScreenRegion, zoom: PinZoom) -> (HWND, HWND) {
    let mut pin = overlay_pin(box_, PinChrome::always());
    pin.zoom = Some(zoom);
    stand_pin(Some(pin));
    pin_stands_in_its_own_window(true);

    pin_real_window::stand_real_pin_window_for_a_test(None);
    let pin_hwnd = pin_real_window::real_pin_window_at(box_).expect("the pin's own window");

    let preview = CreateWindowExW(
        WINDOW_EX_STYLE::default(),
        w!("STATIC"),
        w!(""),
        WS_POPUP,
        NOWHERE,
        NOWHERE,
        320,
        240,
        None,
        None,
        None,
        None,
    )
    .expect("a window of the test's own");
    PREVIEW_HWND.store(preview.0 as isize, Ordering::SeqCst);

    (pin_hwnd, preview)
}

/// A wheel's frame is one present of the window the pin stands in: the picture, the size and the place
/// reach the compositor in the one call that sets them (see `step_pin_zoom`).
///
/// **The fault this is a regression test for.** A picture pin stands in a window of its own, and a
/// wheel's frame was painted into the preview's — the hover's window, which a pin has put down. The
/// window on the screen was left a frame behind, and what caught it up was the tick's own restore: a
/// `MoveWindow` made before the paint that followed it, so the frame the window already had was
/// stretched into the box the frame after it asked for, on every frame of a travel. That is the shake
/// a hand sees zooming in or out below the fit — the wheel's own window changing size, which is the
/// one part of a zoom this app has never been allowed to get wrong.
#[test]
fn a_wheel_frame_sizes_the_window_the_pin_stands_in() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_own = pin_is_its_own_window();
    let previous_pin = take_pin_for_a_test();
    let previous_window = pin_real_window::real_pin_window();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);

    let box_ = (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240);
    // A notch aimed at and not yet arrived at: 352 by 264 from a picture 320 wide, which is a frame of
    // a travel rather than the size the wheel has asked for (see `pin_zoom_release`).
    let mut travel = PinZoom::filling(box_);
    travel.toward = 352.0;
    let (pin_hwnd, preview) = unsafe { a_picture_pin_in_a_window_of_its_own(box_, travel) };

    unsafe { step_pin_zoom(0.05) };

    let content = pin_state()
        .and_then(|state| state.pin().map(|pin| pin.content))
        .expect("the pin that is up");
    assert_ne!(
        content, box_,
        "a frame of a travel is a box of its own, so something was sized by it"
    );
    assert_eq!(
        unsafe { window_rect(pin_hwnd) },
        content,
        "the pin's own window is the window the frame is drawn into, and the one call that draws it \
         is the one that sizes it"
    );
    assert_eq!(
        unsafe { window_rect(preview) },
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        "and the hover's window is left where it stands: the pin's picture is not the hover's"
    );

    // Safety: both windows are this test's own, and every global written here is put back.
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    unsafe {
        let _ = DestroyWindow(pin_hwnd);
        let _ = DestroyWindow(preview);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    pin_stands_in_its_own_window(previous_own);
    stand_pin(previous_pin);
}

/// The tick puts a pin's window back up only where something took it off the screen, and a window
/// Windows minimized is not one of those: a pin that stands in a window of its own is a window like any
/// other, and its own taskbar button is the way down and the way back.
///
/// **The fault this is a regression test for.** The tick asked `the_pin_window_is_down`, which counts a
/// minimized window as a down one, and so every minimize of a real pin was undone a tick later. The
/// pin's own taskbar button could only ever bring it back: clicking it while the pin was the window in
/// front minimized the window — Windows' own gesture, and the one the button exists for — and this app
/// put it up again, so the button restored and never minimized. A `Win`+`D` and a hand's own minimize
/// were one and the same ask, though only the first is the desktop taking the window with everything
/// else, and the second is a window the user put away. What is left to the tick is a window that was
/// *hidden*: the state no button of Windows has a way back from (see `the_pin_owes_a_show`).
#[test]
fn the_tick_puts_up_a_pin_only_where_something_hid_its_window() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_own = pin_is_its_own_window();
    let previous_pin = take_pin_for_a_test();
    let previous_window = pin_real_window::real_pin_window();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);

    let box_ = (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240);
    let (pin_hwnd, preview) =
        unsafe { a_picture_pin_in_a_window_of_its_own(box_, PinZoom::filling(box_)) };
    // On the screen, which is the state the road is asked about — and where no display is, so it is
    // no window of the user's.
    unsafe {
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
    }

    // The window the user is working in, which after a minimize from the pin's own button is any window
    // but the desktop: the Shell hands the front on to whatever is behind the pin.
    let foreground = preview;

    assert!(
        !the_pin_owes_a_show(pin_hwnd, foreground),
        "a pin whose own window is on the screen is not one to put up again, though the hover's \
         window is down"
    );

    // The Shell's own minimize of that window: what its taskbar button sends while the pin is the window
    // in front, and what `Show Desktop` sends to every window of the desktop.
    unsafe {
        let _ = SendMessageW(
            pin_hwnd,
            WM_SYSCOMMAND,
            WPARAM(SC_MINIMIZE as usize),
            LPARAM(0),
        );
    }
    assert!(
        unsafe { IsIconic(pin_hwnd) }.as_bool(),
        "the Shell's ask puts the pin's own window down, which is the button the pin wears"
    );
    assert!(
        !the_pin_owes_a_show(pin_hwnd, foreground),
        "and a pin the Shell minimized is the Shell's to bring back: a tick that put it up again is a \
         taskbar button that can restore and never minimize"
    );

    // What is left to the tick: a window something took off the screen rather than minimized — no button
    // of Windows has a road back from this one, which is why the restore exists at all.
    unsafe {
        let _ = ShowWindow(pin_hwnd, SW_HIDE);
    }
    assert!(
        the_pin_owes_a_show(pin_hwnd, foreground),
        "a pin whose window was hidden is one the tick puts back up"
    );

    // And the show it makes on that answer, which is the whole of what the question is for: the pin comes
    // back on the screen and at the box it stands in.
    unsafe { show_pinned_window(pinned_window()) };
    assert!(!the_pin_window_is_down(pin_hwnd), "the pin comes back up");
    assert_eq!(
        unsafe { window_rect(pin_hwnd) },
        box_,
        "and it comes back where it stood rather than wherever the Shell left it minimized"
    );

    // Safety: both windows are this test's own, and every global written here is put back.
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    unsafe {
        let _ = DestroyWindow(pin_hwnd);
        let _ = DestroyWindow(preview);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    pin_stands_in_its_own_window(previous_own);
    stand_pin(previous_pin);
}

/// A road that carries the hover's handle paints the window the pin stands in: the loop's own handle is
/// the window it pumps messages for, and a picture pin has put that one down (see
/// `render_layered_preview`).
///
/// **The fault this is a regression test for.** Every paint the loop made for a pin — a spinner put
/// away, a playhead turned, a frame laid out again, a menu dismissed — went into the hover's window,
/// which nobody can see. What did reach the screen was the tick's restore, painting the pin's window at
/// the box the pin's state held, on every tick, whatever frame the road had meant to present. Asked of
/// the state instead, the window is one the painter answers for itself.
#[test]
fn a_paint_asked_for_the_hover_reaches_the_window_the_pin_stands_in() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_own = pin_is_its_own_window();
    let previous_pin = take_pin_for_a_test();
    let previous_window = pin_real_window::real_pin_window();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);

    let box_ = (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240);
    let (pin_hwnd, preview) =
        unsafe { a_picture_pin_in_a_window_of_its_own(box_, PinZoom::filling(box_)) };

    // A pin whose box has moved on and whose frame has not been presented yet: what a road that paints
    // owes the window is a size the window is not standing in.
    let grown = (box_.0, box_.1, box_.0 + 352, box_.1 + 264);
    pin_state()
        .and_then(|mut pinned| pinned.pin_mut().map(|pin| pin.content = grown))
        .expect("the pin that is up");

    unsafe { render_layered_preview(preview) };

    assert_eq!(
        unsafe { window_rect(pin_hwnd) },
        grown,
        "the pin's own window is the window a pin's paint is drawn into, whatever handle the road had"
    );
    assert_eq!(
        unsafe { window_rect(preview) },
        box_,
        "and the hover's window is left where it stands: it is down, and it holds no picture of a pin"
    );

    // Safety: both windows are this test's own, and every global written here is put back.
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    unsafe {
        let _ = DestroyWindow(pin_hwnd);
        let _ = DestroyWindow(preview);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    pin_stands_in_its_own_window(previous_own);
    stand_pin(previous_pin);
}

#[test]
fn a_travel_is_placed_about_one_centre_and_never_half_a_pixel_either_side_of_it() {
    // A box whose middle is a half pixel — an odd width and an odd height, which a window a fit or a
    // hand chose is as often as not — is the case the placement used to lose half of. The middle was
    // read back off the whole-pixel box every frame, the half was dropped by the division, and the box
    // the next frame placed was half a pixel to one side of the true middle; a frame later it was on
    // the other side. What an eye saw was the whole window shimmering sideways as it grew.
    //
    // Every box below stays well inside the 1000 by 800 desk, so `clamp_pinned_box` is a no-op and
    // what is asserted is the placement alone.
    let seed = (100, 100, 501, 401);
    let ideal = centred_on(seed);
    let mut content = seed;
    let mut zoom = PinZoom::still((401.0, 301.0), (0.0, 0.0), ideal);
    let mut boxes = Vec::new();

    for drawn in 401..=520 {
        let plan = pin_zoom_to(
            &PinZoomInputs {
                content,
                zoom,
                dpi: DESK,
                room: bounds(),
                pointer: None,
            },
            drawn as f32,
            &display(),
        );

        content = plan.content;
        zoom = plan.zoom;
        boxes.push(content);
    }

    assert_eq!(
        boxes[0], seed,
        "the box the notch found is the box the travel starts from"
    );

    for (at, window) in boxes.iter().enumerate() {
        let centre = centred_on(*window);
        assert!(
            (centre.0 - ideal.0).abs() <= 0.5 && (centre.1 - ideal.1).abs() <= 0.5,
            "the box at frame {at} is centred at {centre:?}, not within half a pixel of {ideal:?}"
        );
    }

    // And the half pixel never takes turns: a centre that comes back to where it was two frames ago
    // is a window that stepped one way and then the other.
    let centres: Vec<(f32, f32)> = boxes.iter().map(|window| centred_on(*window)).collect();
    for (at, three) in centres.windows(3).enumerate() {
        assert!(
            !(three[2] == three[0] && three[1] != three[0]),
            "the window's centre came back to {:?} at frame {at} after {:?}: it is placed by the \
             whole-pixel box rather than by the one centre a travel is owed",
            three[2],
            three[1]
        );
    }
}

#[test]
fn a_notch_on_a_box_a_hand_has_just_moved_zooms_about_where_the_hand_left_it() {
    // The wheel is turned on the box the hand left, and the box grows about its middle: the middle the
    // hand put there, not the whole-pixel middle read back off it. The box is 430 by 330 — an even
    // width whose grown size is odd, so its middle is a whole pixel and the box a notch in has to be
    // half of one either side of it.
    //
    // The box stands well inside the 1000 by 800 desk, so `clamp_pinned_box` is a no-op.
    let left_behind = (100, 100, 530, 430);
    let left_where = centred_on(left_behind);
    let mut zoom = PinZoom::still((430.0, 330.0), (0.0, 0.0), left_where);
    let mut content = left_behind;

    zoom.toward = 430.0 * PIN_ZOOM_PER_NOTCH;

    let mut centres = Vec::new();
    let mut grown = false;

    for frame in 0..30 {
        let plan = pin_zoom_step(
            &PinZoomInputs {
                content,
                zoom,
                dpi: DESK,
                room: bounds(),
                pointer: None,
            },
            0.016,
            &display(),
        );

        let Some(plan) = plan else { break };

        content = plan.content;
        zoom = plan.zoom;
        grown |= zoom.drawn.0 > 430.0;

        let centre = centred_on(content);
        assert!(
            (centre.0 - left_where.0).abs() <= 0.5 && (centre.1 - left_where.1).abs() <= 0.5,
            "frame {frame} put the box's middle at {centre:?}, not within half a pixel of where the \
             hand left it, {left_where:?}"
        );
        assert!(
            content.0 <= left_behind.0 && content.2 >= left_behind.2,
            "frame {frame} grew the box off to one side: {content:?} of {left_behind:?}"
        );
        centres.push(centre);
    }

    assert!(grown, "the box the hand left is the box the notch grew");

    for (at, three) in centres.windows(3).enumerate() {
        assert!(
            !(three[2] == three[0] && three[1] != three[0]),
            "the box's middle came back to {:?} at frame {at} after {:?}: the notch is placed by the \
             box the hand left rather than about it",
            three[2],
            three[1]
        );
    }
}

#[test]
fn a_hand_that_carries_a_window_carries_the_centre_by_the_distance_it_moved() {
    // A move changes no size, so the one centre a travel is placed about moves with the box: a window a
    // hand carried forty pixels left and twenty-five down is grown about the middle it has now, and not
    // the one it was carried from. It is the only road that moves the centre by a distance rather than
    // seating it on a box (see `PinZoom::carried`), and it is a window drag away: a sign the wrong way
    // round here is a notch that grows the box about a point the hand never put it at, which no other
    // test would see.
    //
    // The box a notch is placed in is the box it already has, so what the placement can say is exactly
    // where the hand left it: forty pixels left and twenty-five down.
    let was = (100, 100, 530, 430);
    let mut zoom = PinZoom::still((430.0, 330.0), (0.0, 0.0), centred_on(was));
    zoom.carried((-40, 25));

    let plan = pin_zoom_to(
        &PinZoomInputs {
            content: was,
            zoom,
            dpi: DESK,
            room: bounds(),
            pointer: None,
        },
        430.0,
        &display(),
    );

    assert_eq!(plan.content, (60, 125, 490, 455));
}

#[test]
fn a_travel_moves_every_edge_one_way_and_never_back() {
    // A travel is one direction of the window's box: a window growing about one centre opens its left
    // and top edges up and left and its right and bottom edges down and right, and a window coming
    // down closes them back the same way. An edge that turned round on itself is the window shaking,
    // whatever arithmetic shook it.
    //
    // Four thirds and widescreen, each out from its own box past the cap and back: the cap is where the
    // window stops while the picture goes on growing, which is the one place in a travel the box has a
    // reason to stand still, and a box that stands still is not an edge that went back.
    //
    // Both pictures stand in the middle of the 1000 by 800 desk — 400 by 300 and 480 by 270, whose
    // middles are whole pixels — and the largest box either asks for is the whole width of the desk
    // between the work area's edges, so `clamp_pinned_box` is a no-op all the way out and what is
    // asserted is the placement alone.
    for seed in [(300, 250, 700, 550), (260, 265, 740, 535)] {
        let (width, height) = region_size(seed);
        let drawn = width as f32;
        let mut content = seed;
        let mut zoom = PinZoom::still((drawn, height as f32), (0.0, 0.0), centred_on(seed));
        let mut boxes = Vec::new();

        let frame = |drawn: f32, content: &mut ScreenRegion, zoom: &mut PinZoom| {
            let plan = pin_zoom_to(
                &PinZoomInputs {
                    content: *content,
                    zoom: *zoom,
                    dpi: DESK,
                    room: bounds(),
                    pointer: None,
                },
                drawn,
                &display(),
            );
            *content = plan.content;
            *zoom = plan.zoom;
            plan.content
        };

        // A thousand pixels of travel in four hundred frames, which is a step of two and a half and
        // takes a picture of either shape well past its cap and back.
        for step in 0..=400 {
            boxes.push(frame(drawn + step as f32 * 2.5, &mut content, &mut zoom));
        }
        for step in (0..=400).rev() {
            boxes.push(frame(drawn + step as f32 * 2.5, &mut content, &mut zoom));
        }

        let out = &boxes[..=400];
        let back = &boxes[400..];

        for (at, pair) in out.windows(2).enumerate() {
            let (was, is) = (pair[0], pair[1]);
            assert!(
                is.0 <= was.0 && is.1 <= was.1 && is.2 >= was.2 && is.3 >= was.3,
                "the box at frame {at} on the way out moved an edge back: {is:?} of {was:?}"
            );
        }
        for (at, pair) in back.windows(2).enumerate() {
            let (was, is) = (pair[0], pair[1]);
            assert!(
                is.0 >= was.0 && is.1 >= was.1 && is.2 <= was.2 && is.3 <= was.3,
                "the box at frame {at} on the way back moved an edge out: {is:?} of {was:?}"
            );
        }

        assert_eq!(
            out.iter().map(|box_| region_size(*box_).0).max(),
            Some(1000),
            "the travel reached the cap, where the window stops and the picture does not"
        );
        assert_eq!(
            *boxes.last().unwrap(),
            seed,
            "and the travel back in lands the box it started from"
        );
    }
}

#[test]
fn a_box_keeps_its_place_for_exactly_as_long_as_it_keeps_its_size() {
    // What an eye reads as a window *moving* rather than being resized is the box sliding with its size
    // held — which is exactly what taking the middle off a whole-pixel box every frame produces, half a
    // pixel at a time. So a box either keeps both its size and its place, or it changes both; and what
    // it changes in one frame is a pixel per edge and no more, which is the whole of the difference
    // between a resize and the freeze-then-jump a size rounded away from the ideal one leaves behind.
    //
    // A slow wheel: a fifth of a pixel a frame, the cadence a hand turning a notch over a second asks
    // for. The box stands well inside the desk, so `clamp_pinned_box` is a no-op.
    let seed = (300, 250, 700, 550);
    let mut content = seed;
    let mut zoom = PinZoom::still((400.0, 300.0), (0.0, 0.0), centred_on(seed));
    let mut last: Option<ScreenRegion> = None;

    for step in 0..=500 {
        let plan = pin_zoom_to(
            &PinZoomInputs {
                content,
                zoom,
                dpi: DESK,
                room: bounds(),
                pointer: None,
            },
            400.0 + step as f32 * 0.2,
            &display(),
        );

        content = plan.content;
        zoom = plan.zoom;

        if let Some(was) = last {
            let (size, was_size) = (region_size(content), region_size(was));

            if size == was_size {
                assert_eq!(
                    content, was,
                    "the box moved with its size held at {size:?}: {content:?} of {was:?}"
                );
            } else {
                assert!(
                    (size.0 - was_size.0).abs() <= 2 && (size.1 - was_size.1).abs() <= 2,
                    "the box jumped by {:?} in one frame: {content:?} of {was:?}",
                    (size.0 - was_size.0, size.1 - was_size.1)
                );
                assert!(
                    (content.0 - was.0).abs() <= 1
                        && (content.1 - was.1).abs() <= 1
                        && (content.2 - was.2).abs() <= 1
                        && (content.3 - was.3).abs() <= 1,
                    "an edge jumped more than a pixel in one frame: {content:?} of {was:?}"
                );
            }
        }

        last = Some(content);
    }
}

#[test]
fn a_box_held_at_the_desks_edge_never_takes_turns_between_two_answers() {
    // A travel is monotonic and so is the desk's clamp, and the two together are one direction of the
    // box: every edge opens or holds, and an edge the work area has pinned holds where it pinned it.
    // That is what the placement rests on — the clamp is asked of this frame's whole-pixel rect alone,
    // never of the box the frame before left — and until now it was only written down in a commit
    // message: every other travel in this suite keeps clear of the desk's edge, so no box here was ever
    // held at one. This is that travel: out past the work area's top edge a pixel a frame, from a seed
    // standing forty pixels clear of it.
    let seed = (300, 40, 700, 340);
    let work_area = bounds();
    let mut content = seed;
    let mut zoom = PinZoom::still((400.0, 300.0), (0.0, 0.0), centred_on(seed));
    let mut boxes = Vec::new();

    for step in 0..=600 {
        let plan = pin_zoom_to(
            &PinZoomInputs {
                content,
                zoom,
                dpi: DESK,
                room: work_area,
                pointer: None,
            },
            400.0 + step as f32,
            &display(),
        );

        content = plan.content;
        zoom = plan.zoom;
        boxes.push(content);
    }

    // One way only: a box taking turns between two answers is an edge that moved back.
    for (at, pair) in boxes.windows(2).enumerate() {
        let (was, is) = (pair[0], pair[1]);
        assert!(
            is.0 <= was.0 && is.1 <= was.1 && is.2 >= was.2 && is.3 >= was.3,
            "frame {at} on the way out moved an edge back: {is:?} of {was:?}"
        );
    }

    // And the edge the desk pinned stays pinned: the work area's top is not a line the box crosses and
    // comes back over, it is the line the box's top is held at for the rest of the travel.
    let pinned = boxes
        .iter()
        .position(|box_| box_.1 == work_area.top)
        .expect("the travel runs out past the work area's top edge");
    assert!(
        pinned > 0,
        "the seed stands clear of the edge, so the pin is the travel's and not its start"
    );
    for (at, box_) in boxes.iter().enumerate().skip(pinned) {
        assert_eq!(
            box_.1, work_area.top,
            "frame {at} came back off the top edge the desk pinned at frame {pinned}"
        );
    }
}

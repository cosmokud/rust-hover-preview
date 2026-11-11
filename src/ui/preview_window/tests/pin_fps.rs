//! A wheel turned over a pinned picture's pace control: the notch is the pace's and never the
//! picture's zoom, it is a glide rather than a step, and a hand that takes the pace back re-bases it.

use super::*;

// What the wheel hook is handed a message as (see `wheel_hook_proc`).
use windows::Win32::Foundation::{LPARAM, LRESULT, POINT, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::{HC_ACTION, MSLLHOOKSTRUCT, WM_MOUSEWHEEL};

/// A picture that plays, pinned at a box, with its pace panel up: the state every notch below is
/// turned over.
fn a_pinned_picture_with_its_pace_panel_up(open: bool) -> PinnedPreview {
    let content = (100, 100, 500, 400);
    let mut pin = with_pin_zoom(overlay_pin(content, PinChrome::always()), 1.0);
    pin.transport_bar = true;
    pin.transport_animated = true;
    pin.fps = PinFps {
        detected: 24.0,
        open,
        ..PinFps::default()
    };
    pin
}

/// Turn one notch at a point of the pin, answering what the hook said about it.
fn turn_a_notch_at(window: ScreenRegion, x: i32, y: i32) -> LRESULT {
    let info = MSLLHOOKSTRUCT {
        pt: POINT {
            x: window.0 + x,
            y: window.1 + y,
        },
        mouseData: (WHEEL_DELTA as u32) << 16,
        flags: 0,
        time: 0,
        dwExtraInfo: 0,
    };
    unsafe {
        crate::shell::wheel_input::wheel_hook_proc(
            HC_ACTION as i32,
            WPARAM(WM_MOUSEWHEEL as usize),
            LPARAM(&info as *const MSLLHOOKSTRUCT as isize),
        )
    }
}

/// Stand the pin and the boxes the hook reads, and drain both counts so what is read below is this
/// notch's (see `stand_pin` and `publish_pin_always_on_top`).
fn stand_the_pace_pin(pin: PinnedPreview) -> (ScreenRegion, i32, i32) {
    crate::app::seam::publish_pin_always_on_top(true);
    crate::app::seam::publish_preview_under_the_pointer(false);
    let window = pin.window_box();
    let (width, _) = pin.window_size();
    stand_pin(Some(pin));
    // The box the picture's own zoom claims, which is the whole pin: the pace's box has to win over
    // it for a notch over the panel to be anything but a zoom (see `PIN_ZOOM_REGION`).
    crate::app::seam::PIN_ZOOM_REGION.publish(pin_model::pin_zoom_keep_alive_try());
    let _ = crate::app::seam::take_pin_fps_delta();
    let _ = crate::app::seam::take_pin_zoom_delta();
    (window, width, 0)
}

fn put_the_room_back(previous_pin: Option<PinnedPreview>) {
    crate::app::seam::publish_pin_always_on_top(false);
    crate::app::seam::PIN_ZOOM_REGION.publish(None);
    stand_pin(None);
    stand_pin(previous_pin);
}

/// The pace a wheel notch asks for, added to the pace the pin is at: the tick's own arithmetic, so
/// what the glide below is driven from is the answer the loop would have asked for.
fn ask_for_a_notch(pace: u32, carry: Option<i64>) -> i64 {
    let base = fps_wheel_base_units(carry, pace);
    fps_after_wheel_delta(base, WHEEL_DELTA, current_animated_fps_max())
}

/// Run the glide to its end, answering how many ticks it took and the pace it landed on.
fn glide_to_the_end() -> (u32, u32) {
    let held = || {
        pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| pin.fps))
            .expect("the pin the test stands")
    };
    for ticks in 1..=60 {
        ramp_pin_fps();
        if held().target.is_none() {
            return (ticks, held().pace());
        }
    }
    panic!("the glide never landed");
}

/// The tick's half of a notch, driven the way the loop drives it: the count the hook left is turned
/// into the pace the wheel asked for, and the glide below moves the slider toward it (see
/// `set_pin_fps_target` and `ramp_pin_fps`).
fn take_the_notch_the_hook_left() -> u32 {
    let pace = pin_fps_pace().expect("a pin's pace");
    let units = ask_for_a_notch(pace, None);
    set_pin_fps_target(units);
    pace
}

/// A notch over the pace's own button is the pace's before it is the picture's zoom, and it is
/// swallowed so the listing behind the pin does not move as well. One notch is one frame a second,
/// and it is a glide that lands on it exactly (see `PinFps` and `ramp_pin_fps`).
#[test]
fn a_wheel_notch_over_the_pace_button_is_the_paces_and_moves_one_frame_a_notch() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();

    let pin = a_pinned_picture_with_its_pace_panel_up(false);
    let button = pinned_fps_button(&pin).expect("a picture that plays carries a pace button");
    let (window, _, _) = stand_the_pace_pin(pin);

    let swallowed = turn_a_notch_at(
        window,
        (button.left + button.right) / 2,
        (button.top + button.bottom) / 2,
    );

    assert_eq!(
        crate::app::seam::take_pin_fps_delta(),
        WHEEL_DELTA,
        "the notch over the button is the pace being turned"
    );
    assert_eq!(
        crate::app::seam::take_pin_zoom_delta(),
        0,
        "and it is not also the picture being zoomed"
    );
    assert_eq!(
        swallowed,
        LRESULT(1),
        "and it is not left for the listing behind the pin"
    );

    // And the tick the count is read on asks for one frame more, which the glide lands on rather
    // than stepping to: the notch is in flight for the first tick and on its pace after a handful.
    let pace = take_the_notch_the_hook_left();
    assert!(
        pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| pin.fps.target.is_some()))
            .unwrap_or(false),
        "the notch is a glide rather than a step"
    );
    let (ticks, landed) = glide_to_the_end();
    assert_eq!(landed, pace + 1, "one notch is one frame a second");
    assert!(
        ticks <= 20,
        "and it lands within a few ticks of the pin's own wait: {ticks}"
    );

    put_the_room_back(previous_pin);
}

/// The same notch over the open panel — which floats over the picture the zoom road claims whole.
#[test]
fn a_wheel_notch_over_the_open_pace_panel_is_the_paces() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();

    let pin = a_pinned_picture_with_its_pace_panel_up(true);
    let panel = pinned_fps_panel(&pin)
        .expect("a picture that plays carries a pace panel")
        .panel;
    let (window, _, _) = stand_the_pace_pin(pin);

    let swallowed = turn_a_notch_at(
        window,
        (panel.left + panel.right) / 2,
        (panel.top + panel.bottom) / 2,
    );

    assert_eq!(
        crate::app::seam::take_pin_fps_delta(),
        WHEEL_DELTA,
        "the notch over the panel is the pace's"
    );
    assert_eq!(
        crate::app::seam::take_pin_zoom_delta(),
        0,
        "and the picture is not zoomed under it"
    );
    assert_eq!(swallowed, LRESULT(1));

    put_the_room_back(previous_pin);
}

/// A notch anywhere else on the picture is the zoom's, and the pace is left alone.
#[test]
fn a_wheel_notch_over_the_picture_itself_is_still_the_zoom() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();

    // The panel is open, so the boundary is the panel's edge rather than the button's: a notch in
    // the middle of the picture is neither.
    let pin = a_pinned_picture_with_its_pace_panel_up(true);
    let content = pin.content;
    let (window, _, _) = stand_the_pace_pin(pin);

    let swallowed = turn_a_notch_at(window, (content.2 - content.0) / 2, 20);

    assert_eq!(
        crate::app::seam::take_pin_zoom_delta(),
        WHEEL_DELTA,
        "a notch off both controls is the picture being zoomed"
    );
    assert_eq!(
        crate::app::seam::take_pin_fps_delta(),
        0,
        "and the pace is left where it was"
    );
    assert_eq!(swallowed, LRESULT(1), "and it is still the pin's to take");

    put_the_room_back(previous_pin);
}

/// The pace a notch asks for glides: it climbs toward the wheel's fine target a tick at a time,
/// never past it and never backwards, and lands on exactly one frame on — at either end of the
/// slider as well (see `set_pin_fps_target` and `ramp_pin_fps`).
#[test]
fn a_wheel_notchs_pace_glides_to_its_target_and_holds_at_the_ends() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();

    let pin = a_pinned_picture_with_its_pace_panel_up(true);
    stand_the_pace_pin(pin);

    let held = || {
        pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| pin.fps))
            .expect("the pin the test stands")
    };

    let pace = pin_fps_pace().expect("a pin's pace");
    assert_eq!(pace, 24, "the file's own pace, in whole frames");

    let target = ask_for_a_notch(pace, None);
    set_pin_fps_target(target);
    assert_eq!(held().target, Some(target), "the notch is owed a glide");
    assert_eq!(
        pin_fps_pace(),
        Some(pace),
        "and the slider is still drawn where the pace is"
    );

    let mut applied = 0;
    let mut shown = pace;
    for ticks in 1..=60 {
        ramp_pin_fps();
        let fps = held();
        assert!(fps.applied >= applied, "the glide only ever climbs");
        assert!(fps.applied <= target, "and it never goes past the target");
        assert!(fps.pace() >= shown, "and the drawn pace with it");
        applied = fps.applied;
        shown = fps.pace();
        if fps.target.is_none() {
            break;
        }
        assert!(ticks < 60, "a one-frame notch is a handful of ticks");
    }
    assert_eq!(
        i64::from(held().wanted.expect("the landed pace")),
        i64::from(pace) + 1,
        "one notch is one frame a second, exactly"
    );
    assert_eq!(
        held().applied,
        i64::from(pace + 1) * WHEEL_FPS_UNITS_PER_FRAME
    );

    // And the two ends hold: the bottom of the slider is one frame a second and the top is the
    // ceiling this app's own setting names.
    let max = current_animated_fps_max();
    assert_eq!(
        fps_from_wheel_units(fps_after_wheel_delta(
            WHEEL_FPS_UNITS_PER_FRAME,
            -WHEEL_DELTA,
            max
        )),
        1,
        "a notch down at the bottom stays at one"
    );
    assert_eq!(
        fps_from_wheel_units(fps_after_wheel_delta(
            i64::from(max) * WHEEL_FPS_UNITS_PER_FRAME,
            WHEEL_DELTA,
            max
        )),
        max,
        "and one up at the top stays at the ceiling"
    );

    // A picture at the ceiling lands on it rather than past it.
    set_pin_fps(max);
    let target = ask_for_a_notch(max, None);
    set_pin_fps_target(target);
    let (_, landed) = glide_to_the_end();
    assert_eq!(landed, max, "the glide lands on the ceiling, not past it");

    put_the_room_back(previous_pin);
}

/// A hand that takes the pace after a notch moved it drops the glide, and the next notch is added to
/// the pace the hand left rather than to the one the wheel had asked for (see `set_pin_fps` and
/// `fps_wheel_base_units`).
#[test]
fn a_dragged_knob_after_a_notch_re_bases_the_wheel() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();

    let pin = a_pinned_picture_with_its_pace_panel_up(true);
    stand_the_pace_pin(pin);

    // A notch the loop has taken: its carry is where the wheel asked to go.
    let carry = ask_for_a_notch(24, None);
    set_pin_fps_target(carry);
    assert!(
        pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| pin.fps.target.is_some()))
            .unwrap_or(false),
        "the glide is in flight"
    );

    // And then a hand on the knob, which is a pace in hand rather than a request: nothing is left
    // to glide toward, and the carry is re-based on it.
    set_pin_fps(40);
    let held = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.fps))
        .expect("the pin the test stands");
    assert_eq!(held.wanted, Some(40), "the drag is played at once");
    assert_eq!(held.target, None, "with nothing left to glide toward");
    assert_eq!(
        held.applied,
        40 * WHEEL_FPS_UNITS_PER_FRAME,
        "the glide starts from the pace in hand"
    );

    let base = fps_wheel_base_units(Some(carry), 40);
    let units = fps_after_wheel_delta(base, WHEEL_DELTA, current_animated_fps_max());
    assert_eq!(
        fps_from_wheel_units(units),
        41,
        "the pace the drag set plus one notch is that pace plus one"
    );

    // And the glide a notch begins from there is the pace the drag left, with no jump.
    set_pin_fps_target(units);
    let (_, landed) = glide_to_the_end();
    assert_eq!(landed, 41);

    put_the_room_back(previous_pin);
}

/// A double press on the panel is the pace put back to the file's own: the pace a hand set is a pace
/// a second press takes back, which is what the zero of a slider whose zero is the file means (see
/// `reset_pin_fps`).
#[test]
fn a_double_press_on_the_pace_panel_puts_it_back_to_the_files_own() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let mut pin = a_pinned_picture_with_its_pace_panel_up(true);
    pin.fps.wanted = Some(30);
    let panel = pinned_fps_panel(&pin)
        .expect("a picture that plays carries a pace panel")
        .panel;
    let content = pin.content;
    stand_the_pace_pin(pin);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        let mut media = create_loading_media(400, 300);
        media.media_type = MediaType::AnimatedGif;
        media.animation = AnimationPlayback {
            detected: 24.0,
            wanted: Some(30),
        };
        *current = Some(media);
    }

    let hwnd = HWND(0x1000 as *mut _);
    let window = a_window_at(content);
    let (x, y) = ((panel.left + panel.right) / 2, panel.top + 40);

    // Two presses in the same place inside the panel, the second of which is the pair. The release
    // between them is stood in for here, since a pair is two presses of a slider whose knob was let go
    // of in between (see `pinned_fps_release`).
    assert!(unsafe { pinned_press_on(hwnd, x, y, &window) });
    with_pin(|pin| {
        pin.fps.dragging = false;
        pin.fps.held = None;
    });
    assert!(unsafe { pinned_press_on(hwnd, x, y, &window) });

    let held = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.fps))
        .expect("the pin the test stands");
    assert_eq!(held.wanted, None, "the file plays at its own pace again");
    assert_eq!(held.target, None, "with nothing left to glide toward");
    assert_eq!(
        held.applied,
        fps_from_wheel_units(held.applied) as i64 * WHEEL_FPS_UNITS_PER_FRAME
    );
    assert_eq!(held.pace(), 24, "which is the pace the pin opened at");
    assert!(!held.dragging, "no knob is held by the pair");
    assert!(held.open, "and the panel it was pressed on stays up");
    assert_nothing_is_carried();

    if let Ok(current) = CURRENT_MEDIA.lock() {
        let media = current.as_ref().expect("the media the test stood");
        assert_eq!(
            media.animation.wanted, None,
            "and the picture on screen is played at the file's own pace"
        );
    }

    with_pin(|pin| pin.last_press = None);
    stand_pin(previous_pin);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
}

fn assert_nothing_is_carried() {
    assert!(
        pin_state()
            .and_then(|pinned| pinned.pin().and_then(|pin| pin.dragging))
            .is_none(),
        "the pair carries nothing"
    );
}

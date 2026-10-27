use super::*;

/// A pin that was given the keyboard and no longer has it asks for it back — and a pin that
/// never had it, or that still has it, does not.
///
/// Both facts are unarguable, which is why they are the whole question: the claim is this app's
/// own record of a press, and `GetFocus` is Windows' answer about the caret. The fault behind
/// this is that an arrow press on a pin can be taken by the player instead. The player is begun
/// again on every navigation, every seek and every resize; its window is created *without*
/// `WS_EX_NOACTIVATE` — measured, the style is absent the whole time between the window
/// appearing and this app's monitor thread styling it — and all four of FFmpeg's arrows are
/// seeks, so an arrow that reaches the player is a seek however the symptom is described.
#[test]
fn a_pin_that_claims_the_keyboard_and_does_not_have_it_asks_for_it_back() {
    assert!(
        pin_keyboard_wanted_back(true, false),
        "a claim standing while Windows says the pin does not have the caret is a claim that \
             has stopped being true, and the pin's arrows are then the keys of whatever is in front \
             of it"
    );
    assert!(
        !pin_keyboard_wanted_back(false, false),
        "a pin that was never given the keyboard must not ask for it: that is asking for a \
             keyboard nobody handed over, and the claim would be written on the asking"
    );
    assert!(
        !pin_keyboard_wanted_back(true, true),
        "a pin that still has the caret holds the keyboard it claimed, and asking again is the \
             ask that takes the focus from whatever the user is now working in"
    );
}

/// The arrows on a pinned window are the pin's own, and none of them is ever a seek.
///
/// The mapping itself is asserted above this one; what is asserted here is the property the
/// whole fault turns on. There is no key-driven seek anywhere in this app — a seek is what the
/// transport bar's mouse asks for, and the only keys this app has ever posted to a player are
/// the pause and the bar's own hold, which is the same key — so an arrow cannot be a seek
/// whatever state the pin is in. That is what makes "navigates on the first press, seeks on the
/// second" a statement about *which window* got the key rather than about what the key means,
/// and it is why a loop is given by beginning the file again rather than by posting an arrow at
/// it (see `video_launch::rewind_launch_seconds`).
///
/// The two arrows that are a page of text's own are commands of this app's like the walks are: a
/// walk is what they come to in a window with no page behind them, and that is asked where the
/// file in the window is known rather than here (see `scroll_pinned_text`).
#[test]
fn an_arrow_on_a_pinned_window_is_never_a_seek() {
    for (vk, expected, name) in [
        (VK_LEFT.0 as i32, PinCommand::Previous, "left"),
        (VK_RIGHT.0 as i32, PinCommand::Next, "right"),
    ] {
        assert_eq!(
            pinned_key_command(vk),
            Some(expected),
            "{name} walks the pin's own folder, in both the pin's directions and not only in \
                 the reading one"
        );
    }

    for (vk, expected, name) in [
        (VK_UP.0 as i32, PinCommand::PageBack, "up"),
        (VK_DOWN.0 as i32, PinCommand::PageOn, "down"),
    ] {
        assert_eq!(
            pinned_key_command(vk),
            Some(expected),
            "{name} is the page's own command, which is a walk of the folder to a pin that \
                 holds no page"
        );
    }
}

/// The hook and the pin's own window answer the same keys, and mean the same thing by them.
///
/// These are two copies of one table, because a `WH_KEYBOARD_LL` callback may not take the
/// pin's lock to ask for the other one (see `key_input::PIN_KEY_COMMANDS`). A number that drifts
/// between them is a key this app acts on *wrongly* rather than one that stops working, and it
/// is invisible from either side alone: each side is right about itself.
///
/// The set of keys being the same set is the whole of **one owner per keypress**, which is what
/// the swallow rests on. A key the pin's window answers and the hook does not name reaches
/// FFmpeg whenever the caret is in the player's window — and every key bound to a seek in that
/// player is a fixed increment, so an arrow that arrives there is a step and not a walk. A key
/// the hook names and the pin's window does not answer is acted on for a caret the window was
/// never in. Neither is a missing key: both are a key answered by the wrong party, or by both.
///
/// **And the same is asked of the class of the message, not only of the key.** A chord is not a
/// walk and not a hold however plain the key under the modifier is: the pin's own window eats
/// `Alt+F4` and `Alt+Tab` without answering them, so a hook that counted `WM_SYSKEYDOWN` was
/// walking this window's folder for a `Ctrl`+Left and holding its film for a `Ctrl`+Space —
/// each one a key acted on for a caret the hook owns and answered by nothing for a caret the
/// window was in.
#[test]
fn the_hook_and_the_pin_window_answer_the_same_keys_the_same_way() {
    use crate::shell::key_input::{pin_key_command, pin_key_message_acts, PIN_KEY_COMMANDS};

    for (vk, command) in PIN_KEY_COMMANDS {
        assert_eq!(
            hook_pin_key_command(command),
            pinned_key_command(vk),
            "the hook counts `{vk:#x}` as command {command}, and that number has to mean what \
                 the pin's own window would have done with the key, or a walk arrives as a close"
        );
    }

    // Both directions of the set, over every virtual key a `wParam` can hold: a key one side
    // answers and the other does not is the double action above, and it is exactly what makes
    // "one owner per keypress" a property of the code rather than a hope about it.
    let drifted: Vec<i32> = (0..=0xFF)
        .filter(|vk| pinned_key_command(*vk).is_some() != pin_key_command(*vk).is_some())
        .collect();

    assert!(
        drifted.is_empty(),
        "these keys are answered by one of the two mappings and not the other, which is a key \
             acted on once for a caret this app is not in and passed on once for a caret it is: \
             {drifted:#x?}"
    );

    // The class, over the four messages a key can arrive as. `pinned_key_message_command` is
    // what this window's own procedure asks, so this is the procedure's rule read back rather
    // than a second copy of it — which is the point: the two sides ask one function, and this
    // is what pins that function to the answer both of them need.
    for (message, acted_on, name) in [
        (WM_KEYDOWN, true, "a key-down"),
        (WM_SYSKEYDOWN, false, "a system key-down"),
        (WM_KEYUP, false, "a key release"),
        (
            windows::Win32::UI::WindowsAndMessaging::WM_SYSKEYUP,
            false,
            "a system release",
        ),
    ] {
        assert_eq!(
            pin_key_message_acts(message),
            acted_on,
            "{name} of a key a standing pin answers is {acted_on} answered by the hook, and a \
                 chord the pin's own window throws away must not be counted for a caret this app is \
                 in merely because the caret is in FFmpeg's window instead"
        );
        assert_eq!(
            pinned_key_message_command(message, VK_LEFT.0 as i32, 0).is_some(),
            acted_on,
            "and {name} is answered by the pin's own window on the same question, or the same \
                 key is a walk for one of the two windows this app has and nothing for the other"
        );
    }
}

/// A park is not taken back by the flag it is written in, and this is the whole of why: the band
/// is transparent while a film is playing because FFmpeg's own window stands in it, so a park that
/// ends while that window is not there is a hole in the desktop in every band a player of this
/// app's is drawn in. A resize's release used to be exactly that — the relaunch that answers it
/// wrote the flag down the moment it began a replacement, which is a player that is running and
/// has no window yet — and what the user saw for the length of the wait was the file behind, and
/// the windows beside it, through the window they were holding the edge of.
///
/// So the settle asks the window rather than the flag, and a park that is answered while the
/// replacement is still starting stays standing, holding the last frame it took scaled to the box
/// the drag settled on. It is asked with the window rather than with a pid because the two overlap
/// during a relaunch and a pid cannot say which of them is on screen (see `video_window_for`).
///
/// The frame is given up with the park and not one tick before it: it is the band's picture for as
/// long as the band has no window in it, and a pin with no cover standing pays nothing to hold a
/// picture the size of a display.
#[test]
fn a_park_is_taken_back_only_once_the_players_own_window_is_there() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    stand_pin(Some(PinnedPreview {
        path: PathBuf::from("parked-until-the-replacement-arrives.mkv"),
        ..PinnedPreview::for_test()
    }));

    let mut media = create_loading_media(320, 240);
    media.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }

    assert!(
        park_pinned_player_for_seek(HWND(0x1000 as *mut _), (0, 0)),
        "a drag puts the player's window away, which is the half of the band that is then this \
             app's to fill"
    );
    // The machine this runs on has no player of this app's, so the frame a real park would have
    // taken off the desktop is stood in for it (see `hold_video_window_frame`).
    stand_video_frame_for_a_test([0u8, 0, 255, 255].repeat(16), 4, 4);
    assert!(
        held_video_frame().is_some(),
        "and the frame that was on screen a pointer message ago is held for the drag to scale"
    );

    assert!(
        !settle_pinned_park_onto(false),
        "a replacement that is running and has no window yet is not a picture to see through the \
             band, so the park is not taken back"
    );
    assert!(
        pin_player_is_parked(),
        "and the band keeps the frame it has rather than becoming a hole in the desktop — which is \
             what this very settle used to do, for as long as a player takes to open a file"
    );
    assert!(
        !settle_pinned_park_onto(false),
        "a settle asked again while the replacement is still starting is still the same answer"
    );

    assert!(
        settle_pinned_park_onto(true),
        "and it is taken back the moment the player's own window is standing in the band"
    );
    assert!(
        !pin_player_is_parked(),
        "which is the whole of what the park was holding off: the band is transparent again \
             because there is a window of somebody else's in it"
    );
    assert!(
        !settle_pinned_park_onto(true),
        "and a park that has been taken back cannot be taken back again, so a relaunch that finds \
             no drag in flight does not write a flag nobody will ever ask about"
    );
    assert!(
        held_video_frame().is_none(),
        "the frame goes with it rather than being held for the rest of the run: a pin that is not \
             being dragged pays nothing for a picture the size of a display"
    );

    // And with no pin up there is no flag, so a settle is answered by nothing at all.
    stand_pin(None);
    assert!(
        !settle_pinned_park_onto(true),
        "a park over no pin at all cannot be standing, and so cannot be taken back"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A window that hands the pointer back the way the machine does, by delivering the capture
/// notice back into the window procedure before the call returns.
///
/// `ReleaseCapture` sends `WM_CAPTURECHANGED` to the window it took the capture from, on the
/// same thread and *during* the call, so a window procedure answering one is answering it
/// while the caller is still inside its own release. Standing that up here is the whole of
/// what makes the two ends of a capture tellable apart on a machine with no desktop: a
/// recorder that merely recorded the release would let a park be undone twice with nothing in
/// the test ever noticing.
///
/// Shared rather than kept to this file, because it is the stand-in for the machine and not
/// for one road: whether a release that reached an arm is an answer or a theft of what the
/// press armed cannot be told without it (see `ws_k_chrome_nav`).
pub(super) struct CapturingWindow {
    inner: RecordedPinWindow,
}

impl CapturingWindow {
    /// This window, keeping the rest of a drag's list on the recorder underneath.
    pub(super) fn around(inner: RecordedPinWindow) -> Self {
        Self { inner }
    }

}

impl PinWindow for CapturingWindow {
    fn hwnd(&self) -> isize {
        self.inner.hwnd()
    }

    fn pointer(&self) -> Option<(i32, i32)> {
        self.inner.pointer()
    }

    fn window_box(&self, hwnd: isize) -> Option<ScreenRegion> {
        self.inner.window_box(hwnd)
    }

    fn capture(&self, hwnd: isize) {
        self.inner.capture(hwnd)
    }

    fn release_capture(&self, hwnd: isize) {
        self.inner.release_capture(hwnd);
        // On the machine this is `ReleaseCapture` sending `WM_CAPTURECHANGED` back into this
        // thread's own window procedure, which is where `pin_capture_lost` is called from and
        // still inside the release that caused it.
        pin_capture_lost(self);
    }

    fn set_focusable(&self, hwnd: isize, focusable: bool) {
        self.inner.set_focusable(hwnd, focusable)
    }

    fn set_focus(&self, hwnd: isize) {
        self.inner.set_focus(hwnd)
    }

    fn set_foreground(&self, hwnd: isize) {
        self.inner.set_foreground(hwnd)
    }

    fn hide_pin_windows(&self) {
        self.inner.hide_pin_windows()
    }

    fn hide_pin_bubble(&self) {
        self.inner.hide_pin_bubble()
    }

    fn unpark_player_window(&self, band: Option<ScreenRegion>) {
        self.inner.unpark_player_window(band)
    }

    fn repaint(&self) {
        self.inner.repaint()
    }

    fn post(&self, hwnd: isize, message: u32) {
        self.inner.post(hwnd, message)
    }
}

/// The relayout a finished resize asked for, taken off the request slot.
///
/// Taken rather than read because the slot is a machine value the loop drains on its next turn and
/// other tests write it too, so a test that left one behind would be laying out whatever pin ran
/// next. Shared with the tests about what a drag's end asks for (see `ws_k_chrome_nav`).
pub(super) fn take_relayout_request() -> Option<ScreenRegion> {
    PIN_BOX_REQUEST
        .lock()
        .ok()
        .and_then(|mut request| request.take())
}

/// Two presses are one double-click by the machine's own two measures, and neither of them alone
/// is: the second has to arrive inside the time Windows reads a double-click in and inside the
/// rectangle it lets the pointer wander in between the two, and a press outside either is the first
/// press of a pair of its own (see `within_double_click`).
#[test]
fn two_presses_are_one_double_click_only_inside_the_machines_own_measures() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();
    stand_pin(Some(overlay_pin((100, 100, 500, 400), PinChrome::always())));

    let (x, y) = (200, 200);

    // A pin that has not been pressed has no second press behind the first one.
    assert!(
        !pin_double_click_press(x, y),
        "the first press of a window is nobody's second"
    );
    assert!(
        pin_double_click_press(x, y),
        "and the press after it, at the same place, is the second of the two"
    );

    // The rectangle is the machine's own and the boundary is what decides: a press the pointer
    // could not have wandered to is a press of its own, and the press after *it* pairs with it in
    // its turn rather than with the one three presses back.
    let (reach_x, reach_y) = unsafe {
        (
            GetSystemMetrics(SM_CXDOUBLECLK) / 2,
            GetSystemMetrics(SM_CYDOUBLECLK) / 2,
        )
    };
    let far = (x + reach_x + 1, y);
    assert!(
        !pin_double_click_press(far.0, far.1),
        "a press outside the box the pointer may wander in is not a double-click of the one before \
         it"
    );
    assert!(
        pin_double_click_press(far.0, far.1 + reach_y),
        "and a press inside that box of the press before it is one"
    );

    // The time is the machine's own too, and it is asked of the rule itself: a test cannot wait a
    // double-click out, and nothing about two presses a moment apart would say the measure was
    // read at all.
    let now = Instant::now();
    let read = u64::from(unsafe { GetDoubleClickTime() });
    assert!(
        within_double_click(now, (x, y), now + Duration::from_millis(read), (x, y)),
        "a press inside the time Windows reads a double-click in is part of one"
    );
    assert!(
        !within_double_click(now, (x, y), now + Duration::from_millis(read + 1), (x, y)),
        "and one after that time is a press of its own"
    );

    stand_pin(previous_pin);
}

/// A press that is the second of a pair puts the pinned window in the middle of the screen it
/// stands on rather than taking hold of it: the place is written down and no drag is begun — a
/// window moved out from under a press that is still holding the pointer is a window the next
/// pointer move carries back to where the first press began (see `press_carries_the_pin`).
///
/// The size reset is stood off here, so that what this asks is the move alone and the size a hand
/// gave the window is left exactly as it was: the pair that asks for a size back as well is the
/// test below (see `DoubleClickSettings`).
#[test]
fn a_paired_press_on_a_pinned_window_puts_it_in_the_middle_of_its_screen() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let _settings = DoubleClickSettings::stood_at(true, false);

    // A picture in the corner of the screen it is on, and a media slot a press on a picture is the
    // window's own in: a page of text under the point would have the press instead (see
    // `pinned_content_is_the_pins`).
    let room = work_area_at(0, 0);
    let content = (room.left, room.top, room.left + 400, room.top + 300);
    stand_pin(Some(overlay_pin(content, PinChrome::always())));
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        let mut media = create_loading_media(400, 300);
        media.media_type = MediaType::StaticImage;
        *current = Some(media);
    }

    // A point well inside the picture, clear of the caption's buttons above it and of the resize
    // bands around it.
    let (x, y) = (100, pinned_caption_height(96, None) + 100);
    let hwnd = HWND(0x1000 as *mut _);
    let content_now = || pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.content));

    assert!(
        unsafe { pinned_press(hwnd, x, y) },
        "a press on the picture is the window's to act on"
    );
    assert_eq!(
        content_now(),
        Some(content),
        "the first press of a pair moves the window nowhere: what it begins is a drag, and a drag \
         that has not been carried is not a move"
    );

    assert!(
        unsafe { pinned_press(hwnd, x, y) },
        "and the second press of the pair is the window's to act on as well"
    );
    let moved = content_now().expect("the pin this test put up");
    assert_ne!(
        moved, content,
        "the pair moved nothing: a double-click asks for the window to be put in the middle of the \
         screen it stands on"
    );
    assert_eq!(
        (moved.2 - moved.0, moved.3 - moved.1),
        (content.2 - content.0, content.3 - content.1),
        "and it is a move and nothing else — the size is the one the window had"
    );

    stand_pin(previous_pin);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
}

/// And where the configuration has the size reset on as well, the pair asks the loop for the size
/// rather than writing one itself: a size is what a media is laid out at rather than a place it is
/// carried to, so what a press leaves behind is the request and the box it stands in — the box the
/// loop will lay the file out at is the loop's to write (see `reset_pinned_size`).
#[test]
fn a_paired_press_asks_for_the_size_back_rather_than_writing_one() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let _settings = DoubleClickSettings::stood_at(true, true);

    // The same picture in the corner of the screen it is on, and a point well inside it.
    let room = work_area_at(0, 0);
    let content = (room.left, room.top, room.left + 400, room.top + 300);
    stand_pin(Some(overlay_pin(content, PinChrome::always())));
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        let mut media = create_loading_media(400, 300);
        media.media_type = MediaType::StaticImage;
        *current = Some(media);
    }

    let (x, y) = (100, pinned_caption_height(96, None) + 100);
    let hwnd = HWND(0x1000 as *mut _);
    let content_now = || pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.content));

    assert!(
        unsafe { pinned_press(hwnd, x, y) },
        "a press on the picture is the window's to act on"
    );
    assert!(
        take_pin_command().is_none(),
        "the first press of a pair asks the loop for nothing"
    );

    assert!(
        unsafe { pinned_press(hwnd, x, y) },
        "and the second press of the pair is the window's to act on as well"
    );
    assert_eq!(
        take_pin_command(),
        Some(PinCommand::ResetSize),
        "a pair with the size reset on did not ask the loop for the window's size back"
    );
    assert_eq!(
        content_now(),
        Some(content),
        "and the press wrote a box of its own: what a pair asks for is a size, and a size is a \
         thing the loop lays the file out at"
    );

    stand_pin(previous_pin);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
}

/// A press on the media of a picture the wheel has taken inside its window carries the picture; a
/// press on a picture that fills its window carries the window, exactly as it always has.
///
/// The two halves differ by nothing but what the wheel has done to the picture between them, which is
/// the whole of the claim: the gesture is the same press, and which of the two things under it is the
/// handle is the picture's own size against the window it is drawn in (see `PinZoom::picture_box`).
#[test]
fn a_press_on_a_picture_bigger_than_its_window_carries_the_picture_rather_than_the_window() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();

    let content = (100, 100, 500, 400);
    let hwnd = HWND(0x1000 as *mut _);
    let window = a_window_at(content);
    let dragging = || pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin.dragging));

    // A picture the size of its window has nowhere to be carried, so the press stays the window's own
    // handle and nothing is taken hold of.
    stand_pin(Some(overlay_pin(content, PinChrome::always())));
    assert!(
        !begin_pin_pan(hwnd, &window),
        "a picture the size of its window is not a picture with room to be carried"
    );
    assert!(dragging().is_none(), "and nothing is taken hold of for one");

    // And the same press on the same picture once the wheel has taken it inside that window.
    stand_pin(Some(with_pin_zoom(
        overlay_pin(content, PinChrome::always()),
        3.0,
    )));
    assert!(
        begin_pin_pan(hwnd, &window),
        "a picture drawn larger than its window is the hand's own handle"
    );
    assert!(
        matches!(
            dragging().map(|drag| drag.action),
            Some(PinDragAction::Pan { from }) if from == (0.0, 0.0)
        ),
        "a press begins a pan of it, measured from the pan the pin already had"
    );
    assert!(
        window.calls().contains(&PinWindowCall::Capture),
        "and the pointer is taken for the duration of the drag"
    );

    // What a move carries the pan out to: the picture's own centre travels with the pointer, and the
    // window it is drawn in does not move at all — which is the whole of the difference between a pan
    // and the window drag it replaced.
    super::probes::step_the_pin_drag_by(40, 25);
    unsafe { apply_pin_drag(hwnd) };

    let zoom = pin_state()
        .and_then(|pinned| pinned.pin().and_then(|pin| pin.zoom))
        .expect("the pin this test stood up");
    assert!(
        (zoom.pan.0 - 40.0).abs() <= 2.0 && (zoom.pan.1 - 25.0).abs() <= 2.0,
        "the picture went the way the hand did and by as much: {:?}",
        zoom.pan
    );
    assert_eq!(
        zoom.drawn,
        {
            let (width, height) = region_size(content);
            (width as f32 * 3.0, height as f32 * 3.0)
        },
        "and a pan is a place rather than a size: the picture is drawn exactly as large as it was"
    );
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.content)),
        Some(content),
        "and the window the press was on is not the drag's to move"
    );

    stand_pin(None);
    stand_pin(previous_pin);
}

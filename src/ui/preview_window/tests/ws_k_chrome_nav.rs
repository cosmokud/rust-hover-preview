use super::*;

// WS-K Phase 1: the capture discipline (a press that arms nothing must claim
// nothing), the swap's reconcile (a take-up carries no gesture), the two
// covers (a file step and a box change), and the player's own mouse.
//
// RED (pre-fix): the seams below — `begin_covered_box_change`,
// `cover_step_swap_for_video`, `reconcile_swap_take_up` — do not exist yet, so
// this file does not compile; the capture tests are behavioural red through
// the arms that already exist.

/// A playing pin, as a covered road's press finds it.
fn playing_pin(content: ScreenRegion, from: f64) -> PinnedPreview {
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = PathBuf::from("ws-k-chrome-nav.mkv");
    pin.transport.begun(from, true, false);
    pin.transport.duration = Some(200.0);
    pin.transport.subtitle = Some(2);
    pin.volume.level = 40;
    pin.volume.playing_at = 40;
    pin
}

/// A banded video pin (caption above, transport below — what a video pin is,
/// since its chrome is not drawn over its media), for the bar's own questions.
fn banded_video_pin(content: ScreenRegion, from: f64) -> PinnedPreview {
    PinnedPreview {
        path: PathBuf::from("ws-k-chrome-nav.mkv"),
        content,
        bound: Some(400),
        restore: None,
        dpi: 96,
        transport_bar: true,
        transport_live: true,
        frame: PinFrame::Shaped,
        overlay: false,
        hides_chrome: false,
        caption: pinned_caption_height(96, Some(MediaType::Video)),
        chrome: PinChrome::always(),
        collapsed: false,
        bubble_pause: None,
        hovered: None,
        pressed: None,
        tooltip: PinTooltip::default(),
        dragging: None,
        parked: false,
        transport: {
            let mut transport = PinTransport::default();
            transport.begun(from, true, false);
            transport.duration = Some(200.0);
            transport
        },
        volume: PinVolume::default(),
        audio_hovered: None,
        audio_pressed: None,
    }
}

/// A video behind the pin, as a covered press finds it.
fn stand_video_media() -> Option<MediaData> {
    let previous = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }
    previous
}

fn restore_media(previous: Option<MediaData>) {
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous;
    }
}

fn restore(
    previous_pin: Option<PinnedPreview>,
    previous_pid: u32,
    previous_media: Option<MediaData>,
) {
    take_gesture_snapshot();
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    clear_pending_pinned_relaunch();
    clear_park_swap_arm();
    clear_restart_count();
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    restore_media(previous_media);
}

/// The state effects of an end-of-road relaunch, without starting a real
/// player: the pid the record names, the generation tag, the transport write.
fn fake_end_relaunch(pid: u32, from: f64, holding: bool) {
    VIDEO_PID.store(pid, Ordering::SeqCst);
    note_pinned_relaunch(pid);
    update_pin_transport(|transport| transport.begun(from, true, holding));
}

fn spend_the_swap_bound() {
    if let Ok(mut held) = PIN_PARK_SWAP.lock() {
        if let Some(swap) = held.as_mut() {
            swap.since = Some(Instant::now() - PIN_PARK_SWAP_TIMEOUT * 4);
        }
    }
}

/// A point on the seek track of a banded pin's bar, and the proof that it is
/// on the seek track and not on the play button or the volume button beside
/// it: the bar is laid out by arithmetic a test would otherwise have to
/// reproduce to find a point on it at all.
fn seek_track_point() -> (i32, i32) {
    let bar = pinned_transport_geometry().expect("a banded pin has a bar");
    let (x, y) = (bar.width / 2, bar.top + bar.height / 2);
    assert_eq!(
        pin_chrome::transport_part_at(x, y - bar.top, bar.width, bar.height, bar.dpi, bar.live),
        Some(pin_chrome::TransportPart::Seek),
        "the middle of the bar is the seek track"
    );
    (x, y)
}

/// What a press left armed: a button, a scrub's aim, a drag, a knob.
fn armed() -> (bool, Option<f64>, bool, bool) {
    pin_state()
        .and_then(|pinned| {
            pinned.pin().map(|pin| {
                (
                    pin.pressed.is_some() || pin.transport.pressed.is_some(),
                    pin.transport.seeking,
                    pin.dragging.is_some(),
                    pin.volume.dragging,
                )
            })
        })
        .unwrap_or((false, None, false, false))
}

/// K1 PRESS: a seek press with no second to seek to arms nothing — no aim, no
/// button, no cover, and no capture, because every release arm answers out of
/// what the press armed and a press that armed nothing left the window holding
/// the pointer for the rest of the process (the state a pin is in between a
/// file step and the probe that answers for the file in it).
#[test]
fn a_seek_press_with_no_second_to_seek_to_claims_nothing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let mut pin = banded_video_pin((100, 80, 420, 320), 30.0);
    pin.transport.duration = None;
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();

    let (x, y) = seek_track_point();
    assert!(
        !unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) },
        "a press with no second to seek to is not the bar's to act on"
    );
    assert_eq!(
        armed(),
        (false, None, false, false),
        "nothing is claimed: no button, no aim, no drag, no knob"
    );
    assert!(
        !pin_player_is_parked(),
        "and no cover stands over a film nothing is going to end"
    );
    assert!(
        !gesture_snapshot_active(),
        "and no dead interval is armed behind a press that armed nothing"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K1 keeps the press that *can* seek: a bar a second can be aimed on is the
/// bar's own to take hold of, capture and all.
#[test]
fn a_seek_press_with_a_second_still_arms_the_scrub() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(banded_video_pin((100, 80, 420, 320), 30.0)));
    forget_pin_park_swap();
    forget_video_frame();

    let (x, y) = seek_track_point();
    assert!(
        unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) },
        "a bar with a length to seek on is taken hold of"
    );
    let (_, aim, _, _) = armed();
    assert!(
        aim.is_some(),
        "and the scrub has its aim, which is what the release takes the file to"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K2 RELEASE: an arm that finds nothing to answer for still lets go of the
/// pointer. The transport arm is the one the leak was actually found in: a
/// press with no second to aim armed nothing, and this arm's own early return
/// stood between that and the release of the capture.
#[test]
fn a_transport_release_with_nothing_armed_gives_the_pointer_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(banded_video_pin((100, 80, 420, 320), 30.0)));
    forget_pin_park_swap();
    take_gesture_snapshot();

    let (x, y) = seek_track_point();
    let window = RecordedPinWindow::new(0x1000);
    assert!(
        !unsafe { pinned_transport_release(HWND(0x1000 as *mut _), x, y, &window) },
        "nothing armed is nothing to answer"
    );
    assert_eq!(
        window.calls(),
        vec![PinWindowCall::ReleaseCapture],
        "and the pointer is let go of through the seam before this returns"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K2 RELEASE for the card's own buttons, which a rebuilt pin can empty out
/// from under a press the same way.
#[test]
fn a_card_release_with_no_button_held_gives_the_pointer_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(banded_video_pin((100, 80, 420, 320), 30.0)));

    let (x, y) = seek_track_point();
    let window = RecordedPinWindow::new(0x1000);
    assert!(
        !unsafe { pinned_audio_control_release(HWND(0x1000 as *mut _), x, y, &window) },
        "no button of the card was held"
    );
    assert_eq!(
        window.calls(),
        vec![PinWindowCall::ReleaseCapture],
        "so the pointer is let go of"
    );

    stand_pin(previous_pin);
    restore_media(previous_media);
}

/// K2 RELEASE for the road as a whole: a release whose every arm finds nothing
/// armed is still the end of a press somewhere, so the umbrella hands the
/// pointer back rather than answering nothing at all.
#[test]
fn a_release_with_nothing_armed_anywhere_gives_the_pointer_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(banded_video_pin((100, 80, 420, 320), 30.0)));
    forget_pin_park_swap();
    take_gesture_snapshot();

    let (x, y) = seek_track_point();
    let window = RecordedPinWindow::new(0x1000);
    assert!(
        !unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) },
        "nothing anywhere is armed"
    );
    // One release per arm, which is the rule each arm keeps on its own: the first of them is the
    // capture going back and the rest are the machine saying it has none, and what must not appear
    // among them is any other window work — a refused release ends here and does nothing else.
    let calls = window.calls();
    assert!(
        !calls.is_empty() && calls.iter().all(|call| *call == PinWindowCall::ReleaseCapture),
        "the road lets the pointer go and does nothing else (got {calls:?})"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K2 RELEASE for the knob's own arm: a press the swap disarmed under the
/// hand leaves `dragging` standing nowhere, and this arm's early return is
/// then the last thing between that and a window that holds the pointer.
#[test]
fn a_volume_release_with_no_knob_held_gives_the_pointer_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(banded_video_pin((100, 80, 420, 320), 30.0)));

    let window = RecordedPinWindow::new(0x1000);
    assert!(
        !unsafe { pinned_volume_release(HWND(0x1000 as *mut _), &window) },
        "no knob was held, so there is no knob to let go of"
    );
    assert_eq!(
        window.calls(),
        vec![PinWindowCall::ReleaseCapture],
        "and the pointer goes back rather than staying for the rest of the run"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K2 RELEASE for the drag's arm, which is the arm a title bar and an edge
/// both end in: the pin rebuilt under the hand takes its drag with it, and
/// this is the arm that is left holding the pointer.
#[test]
fn a_drag_end_with_no_drag_gives_the_pointer_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(banded_video_pin((100, 80, 420, 320), 30.0)));
    take_gesture_snapshot();

    let window = RecordedPinWindow::new(0x1000);
    assert!(
        !finish_pin_drag(HWND(0x1000 as *mut _), &window),
        "no drag is no drag to end"
    );
    assert_eq!(
        window.calls(),
        vec![PinWindowCall::ReleaseCapture],
        "so the pointer is let go of here too: one rule, every arm"
    );

    restore(previous_pin, previous_pid, previous_media);
}
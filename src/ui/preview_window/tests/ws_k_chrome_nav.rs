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
        !calls.is_empty()
            && calls
                .iter()
                .all(|call| *call == PinWindowCall::ReleaseCapture),
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
/// K3 RECONCILE: a swap under a hand reconciles what a file step inherits. A
/// snapshot nothing will ever take answers `video_drag_hold_apply` false for the
/// rest of the run, and the take-up's own `release_pin_capture` re-enters
/// `pin_capture_lost`, whose snapshot arm then relaunches the film the pin just
/// left - at its second, at its box - over the player begun for the file that
/// replaced it.
#[test]
fn a_swap_reconciles_the_gesture_it_is_taking_up_with() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));
    forget_pin_park_swap();
    assert!(
        gesture_press_freeze(),
        "a gesture killed a player of its own"
    );
    let generation = current_pinned_generation();
    fake_end_relaunch(102, 30.0, false);
    with_pin(|pin| pin.transport.drag_held = true);
    video_drag_hold_set(true);

    reconcile_swap_take_up(true);

    assert!(
        !gesture_snapshot_active(),
        "the snapshot goes with the step: no end will relaunch a file that is no longer on screen"
    );
    assert!(
        pending_pinned_relaunch().is_none(),
        "and the relaunch behind it is reaped rather than left for the next run's sweep"
    );
    assert!(
        !video_drag_holding(),
        "the claim on the film a gesture was holding goes with it: the next film to be dragged is \
         not answered with a claim about this one"
    );
    assert!(
        current_pinned_generation() == generation,
        "and no generation is bumped for a reconcile, which is not a gesture"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "and that relaunch is the outgoing film's own, so its kill is confirmed: nothing of the \
         film the pin left is left standing"
    );
    assert!(
        orphaned_video_kills().is_empty(),
        "reaped rather than parked for the next run's sweep"
    );

    video_drag_hold_set(false);
    restore(previous_pin, previous_pid, previous_media);
}

/// K3 RECONCILE: a cover standing for the file coming in is carried onto the pin
/// that file is taken up in, because the settle is what takes it down.
#[test]
fn a_swap_carries_a_cover_forward_to_the_incoming_file() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    assert!(cover_step_swap_for_video(true), "the step covers");

    reconcile_swap_take_up(true);

    assert!(
        pin_park_carried_forward(),
        "the cover survives the reconcile, so the take-up's own pin is written parked"
    );
    assert!(
        pin_player_is_parked(),
        "and the standing cover is still standing"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K3 RECONCILE: a cover with nothing behind it is given up rather than carried,
/// so a pin taken up over a picture is never left holding a frozen film.
#[test]
fn a_swap_gives_up_a_cover_nothing_is_coming_to_be_handed() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    assert!(cover_step_swap_for_video(true), "a cover stands");

    reconcile_swap_take_up(false);

    assert!(
        !pin_player_is_parked(),
        "the flag is down: nothing is coming to be handed this band"
    );
    assert!(
        !pin_park_carried_forward(),
        "and no record is left for a settle to answer"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K4 COVER: a step onto a video parks the outgoing frame - the band holds a
/// picture, never the desktop, until the new player has a window of its own.
#[test]
fn a_step_onto_a_video_parks_the_outgoing_frame() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    assert!(cover_step_swap_for_video(true), "a step onto a video parks");
    assert!(pin_player_is_parked(), "the cover stands across the swap");
    assert!(
        PIN_PARK_SWAP
            .lock()
            .ok()
            .and_then(|held| held.as_ref().map(|swap| swap.replacing))
            .unwrap_or(false),
        "for the incoming player behind it"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "and the outgoing player is killed: the cover stands over nothing"
    );
    assert!(
        orphaned_video_kills().is_empty(),
        "a confirmed kill is reaped, never parked"
    );

    // Rapid steps coalesce: a second step extends the standing cover rather than
    // stacking a second park on the first, and the newest file wins.
    fake_end_relaunch(103, 0.0, false);
    assert!(
        cover_step_swap_for_video(true),
        "a rapid second step covers again"
    );
    assert!(pin_player_is_parked(), "still one cover");
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "and the film the first step began is killed by the second"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K4 COVER: a file that paints itself is not covered, and a cover standing over
/// one is given up: there is no player coming to be handed this band, so a
/// frozen film left standing over a picture is worse than the hole it hides.
#[test]
fn a_step_onto_a_picture_raises_no_cover() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));
    forget_pin_park_swap();
    forget_video_frame();

    assert!(
        !cover_step_swap_for_video(false),
        "a picture paints itself: nothing to cover and nothing to hand the cover to"
    );
    assert!(!pin_player_is_parked(), "so no cover stands");
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        101,
        "and the outgoing player is left to the take-down that always killed it"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K4 COVER with nothing behind it is nothing: no pin up, or no player, and there
/// is no band to cover and no frame to read.
#[test]
fn a_step_with_nothing_playing_raises_no_cover() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));
    forget_pin_park_swap();
    assert!(
        !cover_step_swap_for_video(true),
        "with no player there is no window to put away and no frame to read"
    );

    VIDEO_PID.store(101, Ordering::SeqCst);
    stand_pin(None);
    forget_pin_park_swap();
    assert!(
        !cover_step_swap_for_video(true),
        "and with no pin up there is no band to cover"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K4 COVER hands the band over: the settle waits for the incoming player's
/// window rather than swapping onto whatever merely exists, and takes the cover
/// down in the same tick it puts that window up.
#[test]
fn a_step_hands_the_cover_to_the_window_that_replaces_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    assert!(cover_step_swap_for_video(true), "the step covers");

    let window = RecordedPinWindow::new(0x1000);
    fake_end_relaunch(103, 0.0, false);

    // A replacement's window is on screen within milliseconds of being begun and
    // has decoded nothing: the cover holds until the bound, not until the window.
    assert!(
        !settle_pinned_park_where(&window, true),
        "an empty window in the band is not a band to hand back"
    );
    assert!(pin_player_is_parked(), "so the cover stands");

    spend_the_swap_bound();
    assert!(
        settle_pinned_park_where(&window, true),
        "and the bound spent with one standing, the band goes back to it"
    );
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some(content)),
            PinWindowCall::Repaint,
        ],
        "placed before shown, repainted in the same tick"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K5 PRESS: a box change snapshots, parks the cover and kills - the WS-J press
/// half, at the box the window stands at rather than at a gesture's.
#[test]
fn a_box_change_freezes_parks_and_kills() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    let generation = current_pinned_generation();

    assert!(
        unsafe { begin_covered_box_change() },
        "a maximize over a live player takes the covered road"
    );
    assert!(
        gesture_snapshot_active(),
        "the press snapshots: the end relaunch owns the resume"
    );
    assert!(pin_player_is_parked(), "and the cover stands");
    assert!(
        PIN_PARK_SWAP
            .lock()
            .ok()
            .and_then(|held| held.as_ref().map(|swap| swap.replacing))
            .unwrap_or(false),
        "for a relaunch behind it"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "the press kills: VIDEO_PID == 0 mid-road asserts silence"
    );
    assert!(
        current_pinned_generation() > generation,
        "a newer road supersedes whatever relaunch is still in flight"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K5 PRESS with nothing to kill is the legacy road: no snapshot, no cover.
#[test]
fn a_box_change_with_no_player_stays_on_the_legacy_road() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));
    forget_pin_park_swap();

    assert!(
        !unsafe { begin_covered_box_change() },
        "with VIDEO_PID == 0 there is nothing to kill: no snapshot, no cover"
    );
    assert!(!gesture_snapshot_active(), "and no dead interval is armed");
    assert!(!pin_player_is_parked(), "and no cover stands");

    restore(previous_pin, previous_pid, previous_media);
}

/// K5 PRESS is not run twice for one relayout: a resize's drag already ran it
/// and left its cover standing, and running it again would kill the replacement
/// the first one is waiting for.
#[test]
fn a_box_change_over_a_standing_cover_does_not_press_again() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    assert!(
        park_pinned_player(HWND(0x1000 as *mut _), (content.0, content.1), true),
        "a resize's drag parks its own cover"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        101,
        "and its press has not killed yet - the release owns that"
    );

    assert!(
        !unsafe { begin_covered_box_change() },
        "so a box change over it is the same relayout and presses nothing again"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        101,
        "leaving the player the resize's own press stood behind alone"
    );
    assert!(
        pin_player_is_parked(),
        "and the standing cover still standing"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K5 END takes the snapshot once: a second end relaunches nothing, so a
/// maximize is exactly one relaunch at the final box.
#[test]
fn a_box_change_takes_the_snapshot_exactly_once() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    assert!(
        unsafe { begin_covered_box_change() },
        "the press arms one relaunch"
    );

    let first = take_gesture_snapshot();
    assert!(first.is_some(), "the end owns the one relaunch");
    assert!(
        !gesture_snapshot_active(),
        "the take disarms the dead interval"
    );
    assert_eq!(
        take_gesture_snapshot(),
        None,
        "a second end relaunches nothing: one press, one player"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K5 WHOLE ROAD: press, kill, one covered swap at the final box - a playing film
/// is playing after, placed before shown, repainted in the same tick, and never
/// at the box the press found.
#[test]
fn a_box_change_swaps_once_at_the_final_box() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let final_box = (0, 0, 800, 600);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    assert!(
        unsafe { begin_covered_box_change() },
        "the press snapshots, parks and kills"
    );
    let snapshot = take_gesture_snapshot().expect("the end owns the relaunch");
    fake_end_relaunch(
        102,
        snapshot.seconds,
        gesture_end_holding(snapshot.was_playing, snapshot.was_held),
    );
    spend_the_swap_bound();

    let window = RecordedPinWindow::new(0x1000);
    // The maximize has already written the new box into the pin, so the cover was
    // painted at that box rather than the one the press found.
    with_pin(|pin| pin.content = final_box);
    assert!(
        settle_pinned_park_where(&window, true),
        "the current replacement swaps"
    );
    assert!(!pin_player_is_parked(), "and the cover comes down with it");
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some(final_box)),
            PinWindowCall::Repaint,
        ],
        "one swap at the final box: placed before shown, repainted in the same tick"
    );
    assert_eq!(VIDEO_PID.load(Ordering::SeqCst), 102, "exactly one player");
    assert!(
        !settle_pinned_park_where(&window, true),
        "a spent swap never swaps twice"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K5 BOX MATH: a maximize fits the shape to the room and remembers the box;
/// restore gives the remembered box back and forgets it. The cover is scaled to
/// whichever box the pin stands at, so this is the box it is scaled to.
#[test]
fn a_maximize_remembers_its_box_and_a_restore_gives_it_back() {
    let room = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };
    let content = (100, 100, 500, 400);

    let (maximized, restore) = pin_maximize_decided(PinMaximize {
        frame: PinFrame::Shaped,
        content,
        restore: None,
        shape: Some((400, 300)),
        room,
    });
    assert_eq!(
        restore,
        Some(content),
        "a maximize remembers the box it had"
    );
    assert!(
        maximized.2 - maximized.0 <= room.right - room.left
            && maximized.3 - maximized.1 <= room.bottom - room.top,
        "fitted inside the room, got {maximized:?}"
    );

    let (back, restore) = pin_maximize_decided(PinMaximize {
        frame: PinFrame::Shaped,
        content: maximized,
        restore,
        shape: Some((400, 300)),
        room,
    });
    assert_eq!(restore, None, "a restore gives the maximize up");
    assert_eq!(
        (back.2 - back.0, back.3 - back.1),
        (content.2 - content.0, content.3 - content.1),
        "and puts back the size the user had"
    );
}

/// K6 MOUSE: the player's window is made mouse-transparent where the monitor
/// thread already asserts its styles, so a hand on the band cannot reach FFmpeg's
/// own bindings. `left double-click toggle full screen` is compiled into the
/// player and no option unbinds it (verified against ffplay 9.0.2), so this has to
/// be a window property and not a flag.
#[test]
fn the_player_window_is_mouse_transparent() {
    let style = player_ex_style(0);

    assert_eq!(
        style & WS_EX_TRANSPARENT.0 as isize,
        WS_EX_TRANSPARENT.0 as isize,
        "clicks on the band fall through the player's own window rather than reaching its SDL \
         handler"
    );
    assert_eq!(
        style & WS_EX_NOACTIVATE.0 as isize,
        WS_EX_NOACTIVATE.0 as isize,
        "and it still cannot take the keyboard"
    );
    assert_eq!(
        style & WS_EX_TOPMOST.0 as isize,
        WS_EX_TOPMOST.0 as isize,
        "and it is still kept above Explorer"
    );
    assert_eq!(
        style & WS_EX_TOOLWINDOW.0 as isize,
        WS_EX_TOOLWINDOW.0 as isize,
        "and it is still out of the taskbar"
    );
    assert_eq!(
        style & WS_EX_LAYERED.0 as isize,
        0,
        "and it is not made layered: a layered window draws from UpdateLayeredWindow and this \
         one is SDL's own surface, so this is the bit that could stop a film rendering"
    );
    // Idempotent, and everything the window was created with is kept: the monitor
    // thread asserts this list about five times a millisecond.
    assert_eq!(
        player_ex_style(style),
        style,
        "re-asserting changes nothing"
    );
    assert_eq!(
        player_ex_style(0x0004_0000isize) & 0x0004_0000isize,
        0x0004_0000isize,
        "and a style the player was created with is not dropped"
    );
}

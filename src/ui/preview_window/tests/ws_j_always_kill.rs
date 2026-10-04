use super::*;

// WS-J Phase 1: every gesture press kills; the clock freezes over the dead
// interval; every gesture end relaunches exactly once.
//
// RED (pre-fix): the `gesture_*` / `kill_pinned_player_async` seams below do
// not exist, so this file does not compile — the same red WS-G/WS-H started
// from. Behavioural red through the old fns is impossible where noted: the
// old road holds with a blind `P` toggle, so a frozen clock, a dead pid and
// a snapshot are unrepresentable without the fix.

/// A playing pin, as a gesture's press finds it: begun at `from`, volume 40
/// at 40, subtitle track 2, a known length.
fn playing_pin(content: ScreenRegion, from: f64) -> PinnedPreview {
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = PathBuf::from("ws-j-always-kill.mkv");
    pin.transport.begun(from, true, false);
    pin.transport.duration = Some(200.0);
    pin.transport.subtitle = Some(2);
    pin.volume.level = 40;
    pin.volume.playing_at = 40;
    pin
}

/// A paused pin, as a gesture's press finds it: held at `at`, track 1.
fn paused_pin(content: ScreenRegion, at: f64) -> PinnedPreview {
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = PathBuf::from("ws-j-always-kill-paused.mkv");
    pin.transport.begun(at, true, false);
    pin.transport.duration = Some(200.0);
    pin.transport.subtitle = Some(1);
    pin.transport.held(at);
    pin.volume.level = 40;
    pin.volume.playing_at = 40;
    pin
}

/// A video behind the pin, as the press finds it: the kind the kill road
/// is about. Taken away and handed back by each test that stands one.
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
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    restore_media(previous_media);
}

/// The state effects of an end-of-gesture relaunch, without starting a real
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

/// PRESS snapshots everything the end relaunch needs and freezes the clock:
/// the band reads the snapshot second for the whole dead interval.
#[test]
fn press_snapshots_and_freezes_the_clock() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));

    assert!(
        gesture_press_freeze(),
        "a press over a live player takes the kill road"
    );

    let snapshot = take_gesture_snapshot().expect("the press snapshots");
    assert!(
        (snapshot.seconds - 30.0).abs() < 1.0,
        "the playhead second, got {}",
        snapshot.seconds
    );
    assert_eq!(snapshot.level, 40, "the volume level");
    assert_eq!(snapshot.subtitle, Some(2), "the subtitle track");
    assert_eq!(snapshot.content, content, "the content box");
    assert!(snapshot.was_playing, "the film was playing");
    assert!(!snapshot.was_held, "and nothing was holding it");

    let frozen = pin_state().and_then(|pinned| {
        pinned.pin().map(|pin| {
            (
                pin.transport.started.is_none(),
                pin.transport.paused_at,
                pin.transport.pending_hold,
                pin.transport.drag_held,
                pin_playhead(&pin.transport),
            )
        })
    });
    assert_eq!(
        frozen,
        Some((
            true,
            Some(snapshot.seconds),
            false,
            true,
            Some(snapshot.seconds)
        )),
        "the clock is frozen at the snapshot: no start, held there, nothing \
         owed, the gesture's claim on it, and the playhead reads the snapshot"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// PRESS over a paused film freezes too, with no gesture claim: was-held
/// rides the end relaunch, was-playing is false.
#[test]
fn press_over_a_paused_film_freezes_without_a_claim() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(paused_pin(content, 30.0)));

    assert!(
        gesture_press_freeze(),
        "a press over a live-but-held player still takes the kill road"
    );

    let snapshot = take_gesture_snapshot().expect("the press snapshots");
    assert!(!snapshot.was_playing, "the film was not playing");
    assert!(snapshot.was_held, "it was held");

    let claim = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.transport.drag_held))
        .unwrap_or(true);
    assert!(
        !claim,
        "no gesture claim over a film the hand never started: the end \
         relaunch carries was-held, and the swap leaves it held"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// PRESS with no player lives is the legacy road: no snapshot, no freeze.
#[test]
fn press_with_no_player_snapshots_nothing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));

    assert!(
        !gesture_press_freeze(),
        "with VIDEO_PID == 0 there is nothing to kill: no snapshot, no \
         freeze, the legacy road"
    );
    assert!(!gesture_snapshot_active(), "and no dead interval is armed");

    restore(previous_pin, previous_pid, previous_media);
}

/// KILL is image-verified terminate with async confirm, never a UI-thread
/// wait: a fake pid confirms on the spot, VIDEO_PID == 0 mid-gesture.
#[test]
fn kill_confirms_async_and_leaves_no_pid_mid_gesture() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));
    assert!(gesture_press_freeze(), "the press snapshots");

    assert!(
        kill_pinned_player_async(),
        "a dead fake pid confirms on the spot: terminate is a request, the \
         confirm is the read after it, and neither ever waits"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "VIDEO_PID == 0 mid-gesture asserts silence: SW_HIDE never stops \
         audio, only death does"
    );
    assert!(
        gesture_snapshot_active(),
        "the dead interval is still armed: the end relaunch owns the resume"
    );
    assert!(
        orphaned_video_kills().is_empty(),
        "a confirmed kill is reaped, never parked"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// Killing nothing is already reaped.
#[test]
fn kill_with_no_player_is_already_reaped() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);

    assert!(kill_pinned_player_async(), "pid zero is already nothing");

    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
}

/// DURING: scrub and volume steps rewrite aim/level only — no spawn, no
/// key, no further kill.
#[test]
fn steps_rewrite_aim_and_level_only() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let pin = playing_pin((100, 80, 420, 320), 30.0);
    stand_pin(Some(pin));
    assert!(gesture_press_freeze(), "the press snapshots");
    assert!(kill_pinned_player_async(), "the press kills");
    let frozen_at = pin_state()
        .and_then(|pinned| pinned.pin().and_then(|pin| pin.transport.paused_at))
        .expect("the press freezes the clock");

    // Scrub steps: the aim moves, the frozen hold does not.
    for aimed in [31.0, 45.5, 90.0] {
        update_pin_transport(|transport| transport.seeking = Some(aimed));
    }
    // Volume steps: the owed level moves, no player is told.
    set_pin_volume(50);
    set_pin_volume(60);

    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "no step spawns: the dead interval has no player"
    );
    assert_eq!(
        pending_pinned_relaunch(),
        None,
        "and none is in flight behind the cover"
    );
    let kept = pin_state().and_then(|pinned| {
        pinned.pin().map(|pin| {
            (
                pin.transport.seeking,
                pin.transport.paused_at,
                pin.transport.pending_hold,
                pin.transport.drag_held,
                pin.transport.started.is_none(),
                pin.volume.level,
            )
        })
    });
    assert_eq!(
        kept,
        Some((Some(90.0), Some(frozen_at), false, true, true, 60)),
        "steps move the aim and the owed level and leave the frozen hold, \
         the owed flag, the claim and the stopped clock where the press put \
         them"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// DURING posts no key: the hold applicator is a no-op over a dead interval
/// and leaves the gesture's claim standing.
#[test]
fn during_posts_no_key_and_keeps_the_claim() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));
    assert!(gesture_press_freeze(), "the press snapshots");
    assert!(kill_pinned_player_async(), "the press kills");

    assert!(
        !video_drag_hold_apply(true),
        "no key mid-gesture: the blind P toggle stays off every gesture path"
    );
    assert!(
        !video_drag_hold_apply(false),
        "neither does the release arm fire early"
    );
    let kept = pin_state().and_then(|pinned| {
        pinned
            .pin()
            .map(|pin| (pin.transport.paused_at.is_some(), pin.transport.drag_held))
    });
    assert_eq!(
        kept,
        Some((true, true)),
        "the frozen hold and the gesture's claim survive the tick's settle: \
         the end relaunch carries both"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// END takes the snapshot once: a second end is a no-op, so rapid steps
/// coalesce to exactly one relaunch.
#[test]
fn end_takes_the_snapshot_exactly_once() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));
    assert!(gesture_press_freeze(), "the press snapshots");
    assert!(kill_pinned_player_async(), "the press kills");

    let first = take_gesture_snapshot();
    assert!(first.is_some(), "the end owns the one relaunch");
    assert!(
        !gesture_snapshot_active(),
        "the take disarms the dead interval"
    );
    assert_eq!(
        take_gesture_snapshot(),
        None,
        "a second end relaunches nothing: rapid steps coalesce to one"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// END holding = was-held || gesture-held: a playing film resumes, a paused
/// one stays held.
#[test]
fn end_holding_restores_prior_transport() {
    assert!(
        gesture_end_holding(true, false),
        "a playing film carries the gesture hold onto the relaunch"
    );
    assert!(
        gesture_end_holding(false, true),
        "a paused film carries was-held onto the relaunch"
    );
    assert!(
        gesture_end_holding(true, true),
        "either is held: the swap tells them apart, not the relaunch"
    );
    assert!(
        !gesture_end_holding(false, false),
        "nothing held relaunches playing"
    );
}

/// The full dead interval down the state road: press, kill, end relaunch at
/// the snapshot, one covered swap — a playing film is playing after, at the
/// snapshot second.
#[test]
fn whole_gesture_resumes_a_playing_film_at_the_snapshot() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    // PRESS: snapshot, cover, kill.
    assert!(gesture_press_freeze(), "the press snapshots");
    assert!(
        park_pinned_player(HWND(0x1000 as *mut _), (100, 80), false),
        "the cover stands"
    );
    assert!(kill_pinned_player_async(), "the press kills");
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "mid-gesture: no player, no audio"
    );

    // END: the one relaunch at the snapshot second, carrying the hold.
    let snapshot = take_gesture_snapshot().expect("the end owns the relaunch");
    assert!(
        (snapshot.seconds - 30.0).abs() < 1.0,
        "the end relaunches at the snapshot, never snapshot + duration"
    );
    fake_end_relaunch(
        102,
        snapshot.seconds,
        gesture_end_holding(snapshot.was_playing, snapshot.was_held),
    );
    spend_the_swap_bound();

    let window = RecordedPinWindow::new(0x1000);
    assert!(
        settle_pinned_park_where(&window, true),
        "the current replacement swaps"
    );
    assert!(!pin_player_is_parked(), "and the cover comes down with it");
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some(content)),
            PinWindowCall::Repaint,
        ],
        "a single swap: placed before shown, repainted in the same tick"
    );
    assert_eq!(VIDEO_PID.load(Ordering::SeqCst), 102, "exactly one player");
    assert_eq!(
        pending_pinned_relaunch(),
        None,
        "the swap consumes the in-flight record"
    );

    let transport = pin_state().and_then(|pinned| {
        pinned.pin().map(|pin| {
            (
                pin.transport.paused_at,
                pin.transport.pending_hold,
                pin.transport.drag_held,
                pin_playhead(&pin.transport),
            )
        })
    });
    let (paused, owed, claimed, at) = transport.expect("the transport stands");
    assert_eq!(
        paused, None,
        "a playing film is playing after: the swap ends the gesture hold with no key"
    );
    assert!(!owed, "nothing owed");
    assert!(!claimed, "no claim left");
    assert!(
        at.is_some_and(|at| (at - snapshot.seconds).abs() < 0.12),
        "resume within 120ms of the snapshot, got {at:?} vs {}",
        snapshot.seconds
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// The same road for a paused film: the end relaunch carries was-held, the
/// swap leaves it held, and the hold is still owed to the new window.
#[test]
fn whole_gesture_keeps_a_paused_film_paused() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(paused_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    assert!(gesture_press_freeze(), "the press snapshots");
    assert!(
        park_pinned_player(HWND(0x1000 as *mut _), (100, 80), false),
        "the cover stands"
    );
    assert!(kill_pinned_player_async(), "the press kills");

    let snapshot = take_gesture_snapshot().expect("the end owns the relaunch");
    assert!(!snapshot.was_playing && snapshot.was_held, "paused, held");
    fake_end_relaunch(
        102,
        snapshot.seconds,
        gesture_end_holding(snapshot.was_playing, snapshot.was_held),
    );
    spend_the_swap_bound();

    let window = RecordedPinWindow::new(0x1000);
    assert!(
        settle_pinned_park_where(&window, true),
        "the swap is the same road whether the film was playing or paused"
    );

    let transport = pin_state().and_then(|pinned| {
        pinned.pin().map(|pin| {
            (
                pin.transport.paused_at,
                pin.transport.pending_hold,
                pin.transport.drag_held,
            )
        })
    });
    assert_eq!(
        transport,
        Some((Some(snapshot.seconds), true, false)),
        "a paused film is still held at the snapshot second with the hold \
         still owed: the loop delivers it through the new window"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// The park order is capture, paint, hide — and a failed capture keeps the
/// last good frame: never transparent.
#[test]
fn park_captures_before_it_hides_and_never_goes_transparent() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_trace();
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    assert!(
        park_pinned_player(HWND(0x1000 as *mut _), (100, 80), false),
        "the press parks"
    );
    assert_eq!(
        park_trace(),
        vec!["capture", "paint", "hide"],
        "capture before hide: the band holds a frame, never the desktop"
    );

    assert!(
        !hold_video_window_frame(),
        "with no player window there is no picture to take"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// Double-press is idempotent: the second press snapshots nothing new and
/// kills nothing further.
#[test]
fn double_press_is_idempotent() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));
    assert!(gesture_press_freeze(), "the first press snapshots");
    assert!(kill_pinned_player_async(), "the first press kills");

    assert!(
        gesture_press_freeze(),
        "a second press over the dead interval stays on the kill road"
    );
    assert!(
        kill_pinned_player_async(),
        "killing nothing is already reaped"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "no pid returns mid-gesture"
    );
    let second = take_gesture_snapshot().expect("the end owns the one relaunch");
    assert!(
        (second.seconds - 30.0).abs() < 1.0,
        "the snapshot is the same frozen second, not a second kill"
    );
    assert!(
        second.was_playing && !second.was_held,
        "and the first press's claim stands: the second press wrote nothing"
    );
    assert!(
        orphaned_video_kills().is_empty(),
        "no orphan from a press that found nothing to kill"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// Teardown mid-dead-band leaves no orphan: a confirmed kill is reaped, an
/// unconfirmed one rides the reaper, and the snapshot goes with the pin.
#[test]
fn teardown_during_the_dead_band_reaps() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    assert!(gesture_press_freeze(), "the press snapshots");
    assert!(
        park_pinned_player(HWND(0x1000 as *mut _), (100, 80), false),
        "the cover stands"
    );
    assert!(kill_pinned_player_async(), "the press kills");

    end_pin_beside_the_state();

    assert_eq!(
        pending_pinned_relaunch(),
        None,
        "no in-flight relaunch: the press kills, it never launches"
    );
    assert!(
        orphaned_video_kills().is_empty(),
        "a confirmed kill is reaped on the spot, never parked"
    );
    assert!(
        !gesture_snapshot_active(),
        "the snapshot goes with the pin: no stale resume for the next gesture"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "no orphan ffplay, no zombie audio"
    );

    // The unconfirmed half, stood by hand: a kill still dying is owned by the
    // list, and the loop's settle drains it once the death confirms.
    orphaned_video_kills_push_for_test(101);
    settle_video_retirement();
    assert!(
        orphaned_video_kills().is_empty(),
        "the reaper drops only confirmed-gone kills, and only by reaping them"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// Swap regressions kept: a cover with no window holds past its bound, and a
/// spent swap never swaps twice.
#[test]
fn swap_regressions_hold_without_a_window_and_never_twice() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    assert!(
        park_pinned_player(HWND(0x1000 as *mut _), (100, 80), true),
        "the press parks"
    );
    if let Ok(mut held) = PIN_PARK_SWAP.lock() {
        if let Some(swap) = held.as_mut() {
            swap.since = Some(Instant::now() - PIN_PARK_SWAP_TIMEOUT * 4);
        }
    }

    let window = RecordedPinWindow::new(0x1000);
    assert!(
        !settle_pinned_park_where(&window, false),
        "four times the bound with no window behind the band is not a band \
         to hand back"
    );
    assert!(pin_player_is_parked(), "so the cover stands");
    assert_eq!(
        park_swap_last_arm(),
        None,
        "no arm for a park still standing"
    );

    // A window arrives: one swap, placed and repainted in the same tick.
    fake_end_relaunch(102, 30.0, false);
    spend_the_swap_bound();
    assert!(
        settle_pinned_park_where(&window, true),
        "the current replacement swaps"
    );
    assert!(
        !settle_pinned_park_where(&window, true),
        "a spent swap never swaps twice"
    );
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some(content)),
            PinWindowCall::Repaint,
        ],
        "one place-before-show, one repaint"
    );

    restore(previous_pin, previous_pid, previous_media);
}

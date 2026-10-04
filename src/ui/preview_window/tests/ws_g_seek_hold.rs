use super::*;

// WS-G Phase 1: a seeking player is a held player — silent, covered, swapped
// once, restored to its prior transport state.
//
// H1 (S1): seek takes no hold, so audio plays through the scrub. Red below is
// the swap leaving a seek's hold standing (T3) and the missing re-arm/gate
// helpers failing to compile (T5, T8).
// H2 (S2): release bypasses the hold (unheld relaunch) — T2 pins the carry;
// the cover itself stands, T9 pins it.
// H3 (S2 alt): the swap bound is spent during the scrub, so the first tick
// after release swaps onto a window that merely exists — T5 re-arms it.

/// A seek press is a drag-hold decision: a playing film holds, a paused one is
/// left alone. No key is posted for a film with nothing playing to stop, and
/// none for a hold this app already owns (re-press mid-aim).
#[test]
fn seek_press_holds_a_playing_film_and_leaves_a_paused_one() {
    assert_eq!(
        video_drag_hold_decision(true, true, false),
        VideoDragHold::Hold,
        "a press over a playing film holds it, silencing audio with the picture"
    );
    assert_eq!(
        video_drag_hold_decision(true, false, false),
        VideoDragHold::Leave,
        "a press over a paused film holds nothing: was-paused rides the release \
         relaunch as holding and stays paused"
    );
    assert_eq!(
        video_drag_hold_decision(true, true, true),
        VideoDragHold::Leave,
        "a second press while the aim is still held posts no second key: a \
         toggle per press is an unpause, the WS-E lesson"
    );
}

/// The hold a seek press takes survives the release relaunch as owed, and the
/// owed hold defers while the aim is still the gesture's — so scrub steps and
/// the swap wait post no key at all.
#[test]
fn a_seek_hold_survives_its_release_relaunch_and_owes_while_aiming() {
    let mut transport = PinTransport::default();
    transport.begun(0.0, true, false);
    assert!(
        transport_playing(&transport, true),
        "the premise: the film is playing when the hand lands on the bar"
    );

    // The press: the key landed, so the hold is written with the gesture's
    // claim on it.
    transport.held(30.0);
    transport.drag_held = true;

    // The release: one relaunch at the second the hand stopped at, carrying
    // the hold.
    transport.begun(90.0, true, true);
    assert_eq!(
        transport.paused_at,
        Some(90.0),
        "the relaunch is begun at the aimed second and held there, or the new \
         player plays audibly over a bar drawn paused"
    );
    assert!(
        transport.pending_hold,
        "and the hold is owed: the player just begun has no window for the \
         key to be posted to"
    );
    assert!(
        transport.drag_held,
        "with the gesture's claim carried onto the new player, or the owed \
         hold lands the same tick the gesture lets go of it — two toggles, \
         one playing film over a pause glyph"
    );
    assert!(
        !pending_hold_delivers(transport.pending_hold, transport.drag_held),
        "so the owed hold defers while the aim is still the gesture's: no \
         per-step toggle, no same-tick double"
    );
}

/// Scrub steps move the aim and nothing else: no key, no hold written, no
/// claim made or dropped.
#[test]
fn scrub_aiming_moves_nothing_but_the_aim() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.transport.drag_held = true;
    pin.transport.paused_at = Some(30.0);
    stand_pin(Some(pin));

    for aimed in [31.0, 45.5, 90.0] {
        update_pin_transport(|transport| transport.seeking = Some(aimed));
    }

    let kept = pin_state().and_then(|pinned| {
        pinned.pin().map(|pin| {
            (
                pin.transport.seeking,
                pin.transport.paused_at,
                pin.transport.pending_hold,
                pin.transport.drag_held,
            )
        })
    });
    assert_eq!(
        kept,
        Some((Some(90.0), Some(30.0), false, true)),
        "a scrub updates the aimed second and leaves the hold, the owed flag \
         and the claim exactly where the press put them"
    );

    stand_pin(previous_pin);
}

/// The swap ends a seek's hold without a key: the replacement has been playing
/// since the release relaunched it, so posting one would pause the film the
/// swap just uncovered. The owed hold is taken back, the claim dropped, and a
/// playing film is playing after.
#[test]
fn the_swap_ends_the_seek_hold_without_a_key() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(100, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = PathBuf::from("seek-swap-ends-the-hold.mkv");
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    // The press: cover over the live player, hold with the gesture's claim.
    assert!(
        park_pinned_player_for_seek(HWND(0x1000 as *mut _), (100, 80)),
        "the press parks: the cover stands before anything relaunches"
    );
    update_pin_transport(|transport| {
        transport.begun(0.0, true, false);
        transport.held(30.0);
        transport.drag_held = true;
    });

    // The release: one relaunch at the aimed second, carrying the hold. The
    // pid swap is the relaunch the cover is waiting for.
    update_pin_transport(|transport| transport.begun(90.0, true, true));
    VIDEO_PID.store(200, Ordering::SeqCst);

    // The bound spent from the release, which is what the re-arm buys: the
    // swap waits on the new player rather than on a window that merely exists.
    if let Ok(mut held) = PIN_PARK_SWAP.lock() {
        if let Some(swap) = held.as_mut() {
            swap.since = Some(Instant::now() - PIN_PARK_SWAP_TIMEOUT);
        }
    }

    let window = RecordedPinWindow::new(0x1000);
    assert!(
        settle_pinned_park_where(&window, true),
        "a window standing in the band with the bound spent is the swap"
    );
    assert!(!pin_player_is_parked(), "and the cover comes down with it");
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some((100, 80, 420, 320))),
            PinWindowCall::Repaint,
        ],
        "placed before unflagged, repainted in the same tick: no desktop \
         between the placeholder and the player"
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
        Some((None, false, false)),
        "the swap takes back the owed hold and drops the gesture's claim with \
         no key posted: the replacement was playing all along, so a playing \
         film is playing after"
    );
    assert!(
        matches!(park_swap_last_arm(), Some((ParkSwap::TimedOut, _))),
        "on the timeout arm, against the new player"
    );

    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
}

/// A paused film stays paused across the seek swap: no claim, so no take-back,
/// and the owed hold is left for the loop to deliver through the new window.
#[test]
fn a_paused_film_stays_paused_across_the_seek_swap() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(100, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = PathBuf::from("seek-swap-keeps-the-pause.mkv");
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    assert!(
        park_pinned_player_for_seek(HWND(0x1000 as *mut _), (100, 80)),
        "the press parks, held film or not: the cover is about the relaunch \
         behind it either way"
    );
    // Was-paused: the press holds nothing, the release carries the pause.
    update_pin_transport(|transport| {
        transport.begun(0.0, true, false);
        transport.held(30.0);
    });
    update_pin_transport(|transport| transport.begun(90.0, true, true));
    VIDEO_PID.store(200, Ordering::SeqCst);
    if let Ok(mut held) = PIN_PARK_SWAP.lock() {
        if let Some(swap) = held.as_mut() {
            swap.since = Some(Instant::now() - PIN_PARK_SWAP_TIMEOUT);
        }
    }

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
        Some((Some(90.0), true, false)),
        "a paused film is still held at the aimed second with the hold still \
         owed: the loop delivers it through the new window, and no take-back \
         unpauses what the hand never started"
    );

    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
}

/// The release relaunch re-arms the swap bound: the stamp the scrub's ticks
/// kept is spent, so the 600 ms runs from the new player rather than from the
/// press. Without this the first tick after release swaps onto whatever
/// window merely exists — published in milliseconds, empty for hundreds.
#[test]
fn the_release_relaunch_re_arms_the_swap_bound() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    stand_pin(Some(pin));
    forget_pin_park_swap();

    // A long scrub: the stamp below is where a real one would have put it,
    // and the whole bound is spent before anything relaunches.
    if let Ok(mut held) = PIN_PARK_SWAP.lock() {
        *held = Some(PinParkSwap {
            player: 100,
            replacing: true,
            awaiting_relaunch: true,
            owner: ParkArm::Seek,
            since: Some(Instant::now() - PIN_PARK_SWAP_TIMEOUT * 4),
        });
    }

    rearm_seek_cover_for_relaunch();

    let rearmed = PIN_PARK_SWAP.lock().ok().and_then(|held| *held);
    assert!(
        rearmed.is_some_and(|swap| swap.since.is_none()
            && swap.awaiting_relaunch
            && swap.player == 100
            && swap.replacing),
        "the re-arm spends the stamp and keeps the cover: the wait runs from \
         the relaunch, over the player it replaces"
    );

    forget_pin_park_swap();
    stand_pin(previous_pin);
}

/// A seek cover with no window behind it holds past its bound: the band is
/// opaque and there is nothing to hand it to, so the placeholder stands and
/// the wait extends rather than restarts.
#[test]
fn a_seek_cover_with_no_window_holds_past_its_bound() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = PathBuf::from("seek-cover-holds-without-a-window.mkv");
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    assert!(
        park_pinned_player_for_seek(HWND(0x1000 as *mut _), (100, 80)),
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
    assert!(
        pin_player_is_parked(),
        "so the cover stands, holding the frame the press took"
    );
    assert_eq!(
        window.calls(),
        Vec::<PinWindowCall>::new(),
        "asking nothing of either window: the player's is not put up and the \
         band needs no repaint with no upgrade in place"
    );
    assert_eq!(
        park_swap_last_arm(),
        None,
        "no arm was taken for a park that is still standing"
    );

    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
}

/// A press with no second to seek to arms nothing: no cover over a player
/// that is going nowhere, and no hold on a film the hand did not move.
#[test]
fn a_press_with_no_second_to_seek_to_arms_nothing() {
    assert!(
        !seek_press_arms(None),
        "an unknown length is no aim: the press parks nothing and holds \
         nothing, or the cover stands over a playing film with no relaunch \
         coming to end it"
    );
    assert!(
        seek_press_arms(Some(90.0)),
        "and an aimed second arms the whole gesture: cover, hold and the one \
         relaunch the release makes"
    );
}

/// Press→scrub→release down the real road: the aim the press stores survives
/// the hold its key posts, so a scrub moves it and the release has both a
/// second to take the file to and a hold to carry onto the relaunch.
///
/// The earlier tests set `seeking` and `drag_held` by hand, which is why the
/// wipe went uncaught: `held` clears `seeking`, so the hold the press took
/// refused every scrub step at the drag's own guard and left the release
/// with nothing to relaunch — the press's cover stranded over a held film.
#[test]
fn press_scrub_release_keeps_the_aim_and_the_hold() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();

    // A playing film with a known length, as the press finds it.
    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.transport.begun(0.0, true, false);
    pin.transport.duration = Some(200.0);
    stand_pin(Some(pin));

    // The press: an aimed second stored the way `pinned_transport_press`
    // stores it, then the hold its key posts recorded the way
    // `seek_press_hold` records it — one key for the whole gesture.
    let aimed = pin_state()
        .and_then(|pinned| {
            pinned
                .pin()
                .and_then(|pin| pin_seconds_at(&pin.transport, 0.15))
        })
        .expect("a known length aims the press");
    update_pin_transport(|transport| transport.seeking = Some(aimed));
    seek_press_hold_apply(30.0);

    let after_press = pin_state().and_then(|pinned| {
        pinned.pin().map(|pin| {
            (
                pin.transport.seeking,
                pin.transport.paused_at,
                pin.transport.drag_held,
            )
        })
    });
    assert_eq!(
        after_press,
        Some((Some(aimed), Some(30.0), true)),
        "the hold silences the film with the gesture's claim on it and leaves \
         the aim standing, or the scrub has nothing to move and the release \
         nothing to relaunch"
    );

    // The scrub: steps move the aim the way `pinned_transport_drag` moves
    // it, and nothing else.
    for share in [0.25, 0.5] {
        update_pin_transport(|transport| {
            transport.seeking = pin_seconds_at(transport, share);
        });
    }

    let after_scrub = pin_state().and_then(|pinned| {
        pinned.pin().map(|pin| {
            (
                pin.transport.seeking,
                pin.transport.paused_at,
                pin.transport.pending_hold,
                pin.transport.drag_held,
            )
        })
    });
    assert_eq!(
        after_scrub,
        Some((Some(100.0), Some(30.0), false, true)),
        "a scrub moves the aimed second and leaves the hold, the owed flag \
         and the claim where the press put them"
    );

    // The release: reads the aim the way `pinned_transport_release` reads
    // it, and the relaunch carries the hold the way `seek_pinned_playback`
    // carries it.
    let aim = pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin.transport.seeking));
    assert_eq!(
        aim,
        Some(100.0),
        "the release has a second to take the file to, or there is no \
         relaunch and the press's cover stands over a held film"
    );
    assert!(
        pinned_is_held(),
        "and a hold to carry onto it, or the new player plays audibly over \
         a bar drawn paused"
    );
    let holding = pinned_is_held();
    update_pin_transport(|transport| transport.begun(100.0, true, holding));

    let after_release = pin_state().and_then(|pinned| {
        pinned.pin().map(|pin| {
            (
                pin.transport.seeking,
                pin.transport.paused_at,
                pin.transport.pending_hold,
                pin.transport.drag_held,
            )
        })
    });
    assert_eq!(
        after_release,
        Some((None, Some(100.0), true, true)),
        "the release relaunches once at the aimed second carrying the hold: \
         the aim spent, the hold written down as owing with the gesture's \
         claim on it"
    );

    stand_pin(previous_pin);
}

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

    let pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
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

/// A seek cover with no window behind it, and nothing in flight that can put
/// one there, holds while it is still inside the cover's own bound and is
/// given up once it is not.
///
/// **The second half is WS-L and it overturns what this test used to say.** Such a cover used to
/// stand for ever, because `awaiting_relaunch` was an expectation nobody ever checked: the release
/// had already come, and if its relaunch got no player there was nothing left to satisfy it. An
/// opaque band is not a picture that is merely stale — its pixels are claimed by this app's own
/// window at the compositor's hit test, so every click in the video area was answered here rather
/// than by the player. A cover with a scrub still held is a different thing entirely and is held
/// for as long as the hand stays down (see `cover_relaunch_is_still_possible`).
#[test]
fn a_seek_cover_with_nothing_in_flight_is_given_up_by_the_cover_bound() {
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
            swap.since = Some(Instant::now());
        }
    }

    let window = RecordedPinWindow::new(0x1000);
    assert!(
        !settle_pinned_park_where(&window, false),
        "a bound is a bound on a wait, not a delay: nothing has been waited for \
         yet, and the cover holds"
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

    // Past the cover's own bound there is nothing behind the band, nothing in flight that can put
    // a window there — no scrub, no snapshot, no relaunch begun — and so nothing left to wait for.
    // The cover is given up rather than held, which is the whole of what WS-L changed here: an
    // opaque band this app's own window hit-tests is not a picture that is merely stale.
    if let Ok(mut held) = PIN_PARK_SWAP.lock() {
        if let Some(swap) = held.as_mut() {
            swap.since = Some(Instant::now() - PIN_PARK_COVER_TIMEOUT * 2);
        }
    }
    assert!(
        settle_pinned_park_where(&window, false),
        "four times the cover's own bound with nothing behind the band and nothing \
         in flight gives the cover up"
    );
    assert!(
        !pin_player_is_parked(),
        "so no cover stands for ever over a player nothing is going to put there"
    );
    assert!(
        matches!(park_swap_last_arm(), Some((ParkSwap::Abandoned, _))),
        "and the arm is the one the cover's own bound ends on"
    );

    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
}

/// A video behind the pin, as the press road reads the kind off the slot.
fn stand_video_media() -> Option<MediaData> {
    let previous = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }
    previous
}

/// A pin on a film with a live bar: the shape the transport bar's own press is asked
/// about, with the strip standing in bands rather than over the picture so that the bar
/// has rows of its own to be pressed in.
fn a_banded_video_pin(film: &Path, content: ScreenRegion) -> PinnedPreview {
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = film.to_path_buf();
    pin.transport_bar = true;
    pin.transport_live = true;
    pin.overlay = false;
    pin.hides_chrome = false;
    pin.caption = pinned_caption_height(96, Some(MediaType::Video));
    pin.transport.begun(0.0, true, false);
    pin.transport.duration = Some(200.0);
    pin.volume.level = 40;
    pin.volume.playing_at = 40;
    pin
}

/// The second a scrub is aiming at, read out of the pin.
fn the_seek_aim() -> Option<f64> {
    pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin.transport.seeking))
}

/// The line mpv is taken to a second with, as this app writes it (see `mpv_ipc::seek_absolute`).
fn the_seek_line(seconds: f64) -> String {
    format!(r#"{{"command":["seek",{seconds:.3},"absolute"],"async":true}}"#) + "\n"
}

/// A seek on a pin whose player is mpv is a live seek and nothing else: the press tells the
/// player nothing, and the film is left doing what it was doing.
///
/// **The key is counted rather than read off the film, because where it lands is nowhere a test
/// can see** (see `ffplay_window::pause_posts`): there is no player's window here to post one
/// through, so a key posted and a key never posted leave the same trace — and the trace the
/// transport keeps is written only where that post is believed to have landed (see
/// `seek_press_hold`). mpv reads the posted `p` as `cycle pause`, so the user's report is the key:
/// a playing film is stopped by the press and a paused one is *started* by it, and both are what a
/// seek must not be.
#[test]
fn a_seek_press_and_release_on_an_mpv_pin_posts_no_key_and_leaves_the_play_state_it_found() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("mpv-seek-keeps-the-play-state.mkv");
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    stand_ffplay_in(Some(false));
    mpv_ipc::remember_pipe(101, "test-pipe", false);

    // Both play states, because the key a seek press posts is a toggle: a playing film is stopped
    // by it and a paused one is started by it, and neither is a seek.
    for (was_paused, what) in [
        (false, "a playing film"),
        (true, "a film the user had paused"),
    ] {
        let mut pin = a_banded_video_pin(&film, content);
        if was_paused {
            pin.transport.held(30.0);
        }
        let found = (pin.transport.paused_at, pin.transport.drag_held);
        stand_pin(Some(pin));
        ffplay_window::clear_pause_posts();
        mpv_ipc::clear_noted();
        clear_restart_count();

        // The press the hand makes, at the track the bar itself says is one: nothing here
        // reproduces the bar's arithmetic.
        let bar = pinned_transport_geometry().expect("a banded video pin carries a bar");
        let (x, y) = (bar.width / 2, bar.top + bar.height / 2);
        let window = RecordedPinWindow::new(0x1000);
        assert!(
            unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) },
            "the track is the bar's to act on ({what})"
        );
        let aim = the_seek_aim().expect("an aimed second for the scrub to move");
        assert_eq!(
            ffplay_window::pause_posts(),
            0,
            "mpv is told nothing by a seek press: the pause key it would be posted is a toggle \
             that stops {what} the moment the hand takes the bar, and the release posts no second \
             key to start it again"
        );

        assert!(
            unsafe { pinned_transport_release(HWND(0x1000 as *mut _), x, y, &window) },
            "and the release is the bar's own ({what})"
        );
        assert_eq!(
            mpv_ipc::noted_lines(),
            vec![the_seek_line(aim)],
            "the live seek is the whole of the seek: mpv is taken to the aimed second while it \
             runs ({what})"
        );
        assert_eq!(
            restart_count(),
            0,
            "and no player is ended and begun again for it ({what})"
        );

        let after = pin_state().and_then(|pinned| {
            pinned.pin().map(|pin| {
                (
                    pin.transport.paused_at,
                    pin.transport.drag_held,
                    pin.transport.seeking,
                )
            })
        });
        let (held_at, claim, aim_left) = after.expect("the pin is still up");
        assert_eq!(
            claim, found.1,
            "the press leaves no claim on the film behind it: a hold this app believes it took is \
             a hold the next film to be dragged is answered with ({what})"
        );
        assert_eq!(
            aim_left, None,
            "and the aim is spent by the release rather than left standing ({what})"
        );
        assert_eq!(
            held_at.is_some(),
            was_paused,
            "the film is still doing what it was doing: {what}"
        );
        if was_paused {
            assert_eq!(
                held_at,
                Some(aim),
                "a film the user had paused is held at the second the seek landed on, or the bar \
                 springs back to the second the pause began at"
            );
        } else {
            assert_eq!(
                held_at, None,
                "and a film that was playing is playing, with nothing written down as held"
            );
        }
    }

    // FFmpeg's player on the same press: it can be told nothing, so a seek is a player ended and
    // begun again at the aimed second, and it is left exactly as it was (see
    // `seek_pinned_playback`).
    mpv_ipc::clear_noted();
    clear_restart_count();
    stand_pin(Some(a_banded_video_pin(&film, content)));
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    let bar = pinned_transport_geometry().expect("a banded video pin carries a bar");
    let (x, y) = (bar.width / 2, bar.top + bar.height / 2);
    let window = RecordedPinWindow::new(0x1000);
    assert!(
        unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) },
        "FFmpeg's press is the bar's to act on"
    );
    assert!(
        unsafe { pinned_transport_release(HWND(0x1000 as *mut _), x, y, &window) },
        "and so is its release"
    );
    assert_eq!(
        restart_count(),
        1,
        "FFmpeg's player is still ended and begun again at the aimed second"
    );
    assert!(
        mpv_ipc::noted_lines().is_empty(),
        "and nothing of that road reaches an mpv pipe: FFmpeg's player listens on none"
    );

    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    clear_restart_count();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
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

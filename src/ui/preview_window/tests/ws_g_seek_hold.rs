use super::*;

// The stand-in for the machine's own `ReleaseCapture` re-entrancy, shared with the tests about the
// roads a release ends (see `pin_input`).
use super::pin_input::CapturingWindow;

// WS-G: what a hand on the seekbar does to the film it is aimed at.
//
// mpv's film is told nothing by the press and taken to a second at the release, and draws all the
// way through (see the mpv test below). FFmpeg's player cannot be told anything, so its seek is a
// player ended and begun again — at the release, where the second is known, behind a cover raised
// there, over a film that was playing until then. A hold a gesture took is ended by the swap.

/// The hold decision itself, which is the drag's now that a hand on the bar takes none: a playing
/// film holds, a paused one is left alone, and a hold this app already owns is not taken twice.
#[test]
fn a_hold_is_taken_for_a_playing_film_and_left_alone_for_a_paused_one() {
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

/// A hold a gesture takes survives its release relaunch as owed, and the owed hold defers while the
/// aim is still the gesture's — so scrub steps and the swap wait post no key at all.
#[test]
fn a_gesture_hold_survives_its_release_relaunch_and_owes_while_aiming() {
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

/// The swap ends a gesture's hold without a key: the replacement has been playing since the release
/// relaunched it, so posting one would pause the film the swap just uncovered. The owed hold is
/// taken back, the claim dropped, and a playing film is playing after.
#[test]
fn the_swap_ends_a_gestures_hold_without_a_key() {
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

    // The park: cover over the live player, hold with the gesture's claim.
    assert!(
        park_pinned_player_for_seek(HWND(0x1000 as *mut _), (100, 80)),
        "the release parks: the cover stands before anything relaunches"
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

    // The bound runs from the relaunch: the park leaves the stamp unspent, the first tick after it
    // takes one, and the swap is the tick that finds it spent. Here it is spent outright.
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
        "the release parks, held film or not: the cover is about the relaunch \
         behind it either way"
    );
    // Was-paused: nothing held it, and the relaunch carries the pause.
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

/// A seek cover with no window behind it, and nothing in flight that can put
/// one there, holds while it is still inside the cover's own bound and is
/// given up once it is not.
///
/// **The second half is WS-L and it overturns what this test used to say.** Such a cover used to
/// stand for ever, because `awaiting_relaunch` was an expectation nobody ever checked: the release
/// had already come, and if its relaunch got no player there was nothing left to satisfy it. An
/// opaque band is not a picture that is merely stale — its pixels are claimed by this app's own
/// window at the compositor's hit test, so every click in the video area was answered here rather
/// than by the player. A cover whose wait can still be satisfied is a different thing entirely
/// and is held while it can (see `cover_relaunch_is_still_possible` and
/// `a_cover_whose_relaunch_has_not_begun_waits_rather_than_spending_its_bound`).
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
        "the release parks"
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
/// player nothing, and the film is left doing what it was doing — playing, in the window it was
/// already drawing into, for the whole of the scrub and past the release.
///
/// **The key is counted rather than read off the film, because where it lands is nowhere a test
/// can see** (see `ffplay_window::pause_posts`): there is no player's window here to post one
/// through, so a key posted and a key never posted leave the same trace — and the trace the
/// transport keeps is written only where a key is believed to have landed (see
/// `video_drag_hold_apply`). mpv reads the posted `p` as `cycle pause`, so the user's report is the key:
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

        // Nothing of an mpv seek ends its player, so there is no hole for a cover to stand in and
        // no replacement for one to wait for: the cover a park would raise hides the film's window
        // and paints the still over the band for as long as it stands, over {what} that is playing
        // behind it. Neither end of the gesture is a park on this route (see
        // `pinned_transport_press` and `pinned_transport_release`).
        assert!(
            !pin_player_is_parked(),
            "an mpv seek press leaves the film where it is drawing ({what})"
        );

        // And the scrub the hand makes tells the player no more than the press did: the step moves
        // the aim under the hand, and the film keeps playing in the window it is already in until
        // the release takes it to that second (see `pinned_transport_drag`).
        let press_aim = aim;
        assert!(
            unsafe { pinned_transport_drag(HWND(0x1000 as *mut _), bar.width / 4) },
            "a step along the bar is the track's while an aim stands ({what})"
        );
        let aim = the_seek_aim().expect("the second the step moved the aim to");
        assert_ne!(
            aim, press_aim,
            "the step moves the aim under the hand ({what})"
        );
        assert!(
            ffplay_window::pause_posts() == 0 && mpv_ipc::noted_lines().is_empty(),
            "and the player is told nothing while the hand is down: no key for the film it would \
             stop, and no second written down the pipe before the hand lets go ({what})"
        );
        assert!(
            !pin_player_is_parked(),
            "with the film still drawing where it was for the length of the scrub ({what})"
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
        assert!(
            !pin_player_is_parked() && park_record_owner().is_none(),
            "and the release has no cover of its own to give up, so the bound a \
             parked seek waits out is never begun ({what})"
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

    // FFmpeg's player on the same hand: it can be told nothing, so a seek is a player ended and
    // begun again at the aimed second — and the end waits for the release, where that second is
    // known, rather than landing on the press and pausing the film for the whole of the scrub (see
    // `pinned_transport_release` and `seek_pinned_playback`).
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
    let press_aim = the_seek_aim().expect("an aimed second for the scrub to move");
    assert_eq!(
        restart_count(),
        0,
        "and the film plays on under the hand: the player is ended at the release, not here"
    );
    assert!(
        !pin_player_is_parked() && park_record_owner().is_none(),
        "with no cover over it, since nothing has ended for one to hide"
    );

    assert!(
        unsafe { pinned_transport_drag(HWND(0x1000 as *mut _), bar.width / 4) },
        "a step along the bar is the track's while an aim stands"
    );
    let aim = the_seek_aim().expect("the second the step moved the aim to");
    assert_ne!(aim, press_aim, "the step moves the aim under the hand");
    assert_eq!(
        restart_count(),
        0,
        "and the scrub ends no player at all: it moves an aim and nothing else"
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
    assert_eq!(
        the_seek_aim(),
        None,
        "the aim the press stored is the one the release spent, or there is no relaunch at all"
    );
    assert!(
        pin_player_is_parked() && park_record_owner() == Some(ParkArm::Seek),
        "and the release raised a seek's cover over the frame the hand let go on, which stands \
         until the player replacing the ended one has decoded a frame"
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

/// An mpv seek this app cannot hand to the player falls back to the relaunch every other road takes,
/// and that relaunch is covered like every other one: a player ended with no cover over the band is a
/// hole in the desktop shaped like a film, for as long as the file takes to open and to decode.
#[test]
fn an_mpv_seek_whose_pipe_is_gone_raises_the_cover_before_it_ends_the_player() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("mpv-seek-with-no-pipe.mkv");
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    stand_ffplay_in(Some(false));
    // No pipe is remembered for this player, so the live seek has nowhere to go.
    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    clear_restart_count();

    for (was_paused, what) in [
        (false, "a playing film"),
        (true, "a film the user had paused"),
    ] {
        let mut pin = a_banded_video_pin(&film, content);
        if was_paused {
            pin.transport.held(30.0);
        }
        stand_pin(Some(pin));
        forget_pin_park_swap();
        clear_restart_count();

        let bar = pinned_transport_geometry().expect("a banded video pin carries a bar");
        let (x, y) = (bar.width / 2, bar.top + bar.height / 2);
        let window = RecordedPinWindow::new(0x1000);
        assert!(
            unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) },
            "the track is the bar's to act on ({what})"
        );
        assert!(
            unsafe { pinned_transport_release(HWND(0x1000 as *mut _), x, y, &window) },
            "and the release is the bar's own ({what})"
        );

        assert_eq!(
            restart_count(),
            1,
            "a seek the pipe cannot take is a player ended and begun again ({what})"
        );
        assert!(
            pin_player_is_parked() && park_record_owner() == Some(ParkArm::Seek),
            "and the cover stands over the hole that relaunch leaves, raised where the end was \
             decided rather than over a film that was still playing ({what})"
        );
        assert!(
            mpv_ipc::noted_lines().is_empty(),
            "nothing is written down a pipe that is not there ({what})"
        );
    }

    forget_pin_park_swap();
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

/// **The strip a seek's cover is painted with stands where the hand let go, and the bar is the only
/// place that mistake can show.** The release gives the pointer back before it raises the cover,
/// and the notice that comes back from doing that (`pin_capture_lost`) takes the aim with it —
/// while the cover is painted *from* the transport. So the strip under the still was drawn at the
/// second the *old* player's own clock had reached, which is the position the thumb stood at
/// before the drag, and that strip is what the compositor is handed until the relaunch rebases the
/// clock — the flash between the park's paint and the release's own, for as long as the kill and the
/// spawn take. The film plays throughout, so nothing else about the gesture shows it.
#[test]
fn the_cover_a_seek_release_raises_draws_the_bar_where_the_hand_let_go() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("seek-cover-draws-the-aim.mkv");
    stand_video_media();
    // FFmpeg's route rather than the machine's, because the cover is the road only that player
    // takes: a seek on mpv is a line down its pipe and a film nobody ends (see `mpv_ipc`).
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, false)));

    take_gesture_snapshot();
    clear_pending_pinned_relaunch();
    stand_pin(Some(a_banded_video_pin(&film, content)));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    clear_restart_count();

    // The hand: a press on the track, a step along it, and a release where it stopped — through
    // the machine's own re-entrancy, where the capture the press took is handed back and
    // `pin_capture_lost` runs before the release's own arm has raised anything.
    let hwnd = HWND(0x1000 as *mut _);
    let bar = pinned_transport_geometry().expect("a banded video pin carries a bar");
    let y = bar.top + bar.height / 2;
    let let_go_at = (bar.width * 5) / 8;
    let window = CapturingWindow::around(a_window_at(content));
    assert!(
        unsafe { pinned_transport_press(hwnd, bar.width / 2, y) },
        "the track is the bar's to act on"
    );
    assert!(
        unsafe { pinned_transport_drag(hwnd, let_go_at) },
        "a step along the bar moves the aim under the hand"
    );
    let aim = the_seek_aim().expect("a second for the release to take the file to");
    assert!(
        unsafe { pinned_transport_release(hwnd, let_go_at, y, &window) },
        "and the release is the bar's own"
    );
    assert_eq!(
        restart_count(),
        1,
        "the release is one player ended and begun again at the aimed second, behind its cover"
    );

    // The surface a paint draws into is keyed by the window, so the very surface the park painted
    // is the one read back here rather than a second paint of it.
    let (width, height) = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_size()))
        .expect("a pin to have a size");
    let strip = |surface: *mut u8| -> Vec<u8> {
        let all =
            unsafe { std::slice::from_raw_parts(surface, width as usize * height as usize * 4) };
        let rows = pinned_transport_height(96, true) as usize;
        all[(height as usize - rows) * width as usize * 4..].to_vec()
    };
    let covered = strip(
        ensure_layered_surface(hwnd.0 as isize, width as u32, height as u32)
            .expect("the surface the park painted"),
    );
    assert!(
        covered.iter().any(|channel| *channel != 0),
        "the cover painted a strip at all: an unpainted one would make the comparison below say \
         nothing"
    );

    // The same strip with the aim standing, which is what the cover is meant to be.
    update_pin_transport(|transport| transport.seeking = Some(aim));
    unsafe { render_pinned_preview_at(hwnd, content.0, content.1) };
    let at_the_aim = strip(
        ensure_layered_surface(hwnd.0 as isize, width as u32, height as u32)
            .expect("the surface the paint drew into"),
    );

    assert!(
        covered == at_the_aim,
        "the strip the cover is painted with stands at {aim:.3} seconds, where the hand let go, \
         rather than at the second the player it ended had reached"
    );

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    take_gesture_snapshot();
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    clear_pending_pinned_relaunch();
    clear_park_swap_arm();
    clear_restart_count();
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

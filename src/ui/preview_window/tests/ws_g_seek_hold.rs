use super::*;

// WS-G: what a hand on the seekbar does to the film it is aimed at.
//
// The press tells the player nothing on either route, and the film plays on under the hand. mpv is
// taken to the second over its pipe at the release and keeps the player it has; FFmpeg's player —
// which cannot be told anything — is begun again there, and **nothing is taken off the screen for
// it**: the registration the outgoing player's window is drawn through is the picture until the
// replacement's source has taken it over, and only then is the outgoing player ended (see
// `pin_playback::relaunch_pinned_player` and `pin_film_thumbnail::hand_the_picture_over`).

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

/// **The press on the bar tells the player nothing on either route**, and the film plays on under
/// the hand: no key, no restart, and the aim moves with the pointer. The second is only known at
/// the release, and nothing about the film may change before then.
#[test]
fn a_seek_press_and_scrub_tell_the_player_nothing_on_either_route() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("seek-press-tells-nothing.mkv");

    for (engine, mpv, ffplay) in [
        (VideoEngine::Mpv, true, false),
        (VideoEngine::Ffmpeg, false, true),
    ] {
        stand_mpv_in(Some(mpv));
        stand_ffplay_in(Some(ffplay));
        stand_video_engine_in(Some((engine, true)));
        stand_pin(Some(a_banded_video_pin(&film, content)));
        ffplay_window::clear_pause_posts();
        mpv_ipc::clear_noted();
        clear_restart_count();
        let hides = pin_film_host::host_hides_asked();

        let bar = pinned_transport_geometry().expect("a banded video pin carries a bar");
        let (x, y) = (bar.width / 2, bar.top + bar.height / 2);
        let window = RecordedPinWindow::new(0x1000);
        let what = format!("{engine:?}");
        assert!(
            unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) },
            "the track is the bar's to act on ({what})"
        );
        let press_aim = the_seek_aim().expect("an aimed second for the scrub to move");
        assert!(
            unsafe { pinned_transport_drag(HWND(0x1000 as *mut _), bar.width / 4) },
            "a step along the bar is the track's while an aim stands ({what})"
        );
        assert_ne!(
            the_seek_aim(),
            Some(press_aim),
            "the step moves the aim under the hand ({what})"
        );

        assert_eq!(
            ffplay_window::pause_posts(),
            0,
            "no pause key: a toggle would stop the film the moment the hand took the bar ({what})"
        );
        assert_eq!(
            restart_count(),
            0,
            "and no player is ended while the second is still being chosen ({what})"
        );
        assert_eq!(
            pin_film_host::host_hides_asked(),
            hides,
            "and the picture is never taken off the screen for a press ({what})"
        );

        // The release is the release's own; this test is about the press and the scrub.
        let _ = unsafe { pinned_transport_release(HWND(0x1000 as *mut _), x, y, &window) };
        mpv_ipc::forget_pipe(101);
        mpv_ipc::clear_noted();
        clear_restart_count();
    }

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// **A release on FFmpeg's route is one player begun again at the aimed second, and it says why**:
/// `player restarted for seek`. mpv is asked down its pipe and begins nothing. On both routes the
/// film is on screen throughout: no host is hidden, nothing is painted flat over the band, and the
/// player that was drawing it is still the one the bar reads from until its replacement stands.
#[test]
fn a_seek_release_restarts_once_and_the_film_stays_on_screen() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("seek-release-restarts-once.mkv");

    // FFmpeg's player: the player cannot be told a second, so a seek is a relaunch.
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    stand_pin(Some(a_banded_video_pin(&film, content)));
    clear_restart_count();
    pin_trace::clear_noted();
    let hides = pin_film_host::host_hides_asked();

    let bar = pinned_transport_geometry().expect("a banded video pin carries a bar");
    let (x, y) = (bar.width / 2, bar.top + bar.height / 2);
    let window = RecordedPinWindow::new(0x1000);
    assert!(unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) });
    assert!(unsafe { pinned_transport_drag(HWND(0x1000 as *mut _), bar.width / 4) });
    assert!(unsafe { pinned_transport_release(HWND(0x1000 as *mut _), x, y, &window) });

    assert_eq!(restart_count(), 1, "one seek, one player begun");
    assert_eq!(
        last_restart_reason(),
        Some(RestartReason::Seek),
        "and the line says what it was for"
    );
    assert_eq!(
        the_seek_aim(),
        None,
        "and the aim is spent by the release rather than left standing"
    );
    assert_eq!(
        pin_film_host::host_hides_asked(),
        hides,
        "the host the picture is drawn in is never hidden for a relaunch"
    );
    assert!(
        !window
            .calls()
            .iter()
            .any(|call| matches!(call, PinWindowCall::Repaint)),
        "and the pin's window is not repainted for a relaunch: there is no cover to take down and          nothing of the picture to put back"
    );
    let lines = pin_trace::noted_lines();
    assert!(
        !lines
            .iter()
            .any(|line| line.starts_with("film host hidden while a pin lives")),
        "and a run says nothing was taken off the screen: {lines:?}"
    );

    // mpv: the same hand, and the player is taken to the second over its pipe. Nothing is begun.
    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    clear_restart_count();
    stand_mpv_in(Some(true));
    stand_ffplay_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    mpv_ipc::remember_pipe(101, "test-pipe", false);
    VIDEO_PID.store(101, Ordering::SeqCst);
    stand_pin(Some(a_banded_video_pin(&film, content)));
    let hides = pin_film_host::host_hides_asked();

    let bar = pinned_transport_geometry().expect("a banded video pin carries a bar");
    let (x, y) = (bar.width / 2, bar.top + bar.height / 2);
    assert!(unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) });
    assert!(unsafe { pinned_transport_drag(HWND(0x1000 as *mut _), bar.width / 4) });
    let aim = the_seek_aim().expect("the second the hand stopped at");
    assert!(unsafe { pinned_transport_release(HWND(0x1000 as *mut _), x, y, &window) });
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![the_seek_line(aim)],
        "mpv is taken to the aimed second while it runs"
    );
    assert_eq!(restart_count(), 0, "and no player is begun again for it");
    assert_eq!(
        pin_film_host::host_hides_asked(),
        hides,
        "and nothing of the picture is hidden on that route either"
    );

    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    clear_restart_count();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// **A seek's release ends nothing on the spot.** FFmpeg's player is the one picture the pin has
/// until the player that replaces it has drawn its own, so the player under the hand is retired by
/// the loop once that picture stands rather than killed at the release — the kill is the band the
/// host's black shows through for the length of a decode (see `retire_replaced_player`).
///
/// The line read here is the relaunch's own `from`, which is the pid of the player that was on
/// screen: a release that ended it first leaves `VIDEO_PID` cleared before the relaunch reads it,
/// and the line says `from 0` with no player for the retirement to hold.
#[test]
fn a_seek_release_ends_no_player_before_its_replacement_can_draw() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let film = PathBuf::from("seek-release-ends-nothing.mkv");

    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    stand_pin(Some(a_banded_video_pin(&film, content)));
    clear_restart_count();
    pin_trace::clear_noted();

    let bar = pinned_transport_geometry().expect("a banded video pin carries a bar");
    let (x, y) = (bar.width / 2, bar.top + bar.height / 2);
    let window = RecordedPinWindow::new(0x1000);
    assert!(unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) });
    assert!(unsafe { pinned_transport_drag(HWND(0x1000 as *mut _), bar.width / 4) });
    assert!(unsafe { pinned_transport_release(HWND(0x1000 as *mut _), x, y, &window) });

    assert_eq!(
        restart_count(),
        1,
        "the seek is one relaunch on this engine"
    );
    let lines = pin_trace::noted_lines();
    assert!(
        lines
            .iter()
            .any(|line| line.starts_with("player restarted for seek from 101 to ")),
        "the player on screen is still the one the relaunch replaces, so nothing ended it: {lines:?}"
    );
    assert!(
        !lines
            .iter()
            .any(|line| line.starts_with("player ended while standing")),
        "no end is written against a picture that is still on screen: {lines:?}"
    );

    clear_restart_count();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

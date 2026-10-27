use super::*;

// A player that a cover ended on purpose is not a file that failed.
//
// `pin_media_failed_before_a_frame` is the other side of `pin_media_is_alive`, and the two have to
// read the same in-flight fact: a park whose replacement is on its way. Without it the tick after a
// step's own player reads it as a dead file — inside the three-second give-up — and the pin falls to
// the failure mark before the replacement can arrive. That is the "window breaks permanently" after
// a Next: the film's player is killed, and the pin is left on the cross.

/// A video behind the pin, as the take-up reads the kind off the slot, with a player that is
/// not running: the state `pin_media_failed_before_a_frame` is about.
fn stand_a_video(width: u32, height: u32) -> Option<MediaData> {
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    stand_mpv_in(Some(false));
    let previous = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let mut media = create_loading_media(width, height);
    media.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }
    previous
}

#[test]
fn a_covers_own_player_is_not_a_failed_file() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_media = stand_a_video(640, 360);
    let previous_pin = take_pin_for_a_test();
    forget_pin_park_swap();
    VIDEO_PID.store(0, Ordering::SeqCst);

    let path = PathBuf::from("pin-player-failure.mkv");
    install(take_up_pinned_window(&path, (100, 80, 800, 320)));
    let started = Some(Instant::now());

    // A player that is gone with nothing on its way is a file that failed, and the two facts
    // below are what say so — no cover stands yet.
    assert_eq!(
        pin_media_failed_before_a_frame(started).as_deref(),
        Some(path.as_path()),
        "a player that is gone with nothing on its way is read as a failed file",
    );

    // A file step's cover: the player is killed on purpose and a replacement is on its way. The
    // same state must not read as a failed file, or the pin is stuck on the cross before the
    // replacement lands.
    VIDEO_PID.store(1, Ordering::SeqCst);
    assert!(
        park_pinned_player_for_step(pinned_window(), (100, 80)),
        "a step parks the player it is about to replace",
    );
    assert_eq!(
        pin_media_failed_before_a_frame(started),
        None,
        "a player a cover ended is not a file that failed",
    );

    // The cover is taken back once the new window is there; then the same gone player is refused
    // again, on purpose: what is left is a file the player really did not draw.
    forget_pin_park_swap();
    with_pin(|pin| pin.parked = false);
    assert_eq!(
        pin_media_failed_before_a_frame(started).as_deref(),
        Some(path.as_path()),
        "with no cover standing, a gone player is a failed file again",
    );

    VIDEO_PID.store(0, Ordering::SeqCst);
    forget_pin_park_swap();
    forget_video_frame();
    stand_pin(previous_pin);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_video_engine_in(None);
}

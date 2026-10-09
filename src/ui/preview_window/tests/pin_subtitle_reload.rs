use super::*;
use super::super::subtitle_files::VideoSubtitlesSetting;

/// A video behind the pin, as the reload's own guard reads it: the kind this app plays with
/// FFmpeg, taken away and handed back by the test that stands one — the same shape `ws_j` uses
/// for the same question.
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

/// A pin standing on `film` at a box, playing from the thirtieth second — the state a hover
/// leaves behind when it is pinned. No player is begun here; every question below is about the
/// relaunch being asked for, not about the process it would leave behind.
fn a_pin_playing(film: &Path, content: ScreenRegion) -> PinnedPreview {
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = film.to_path_buf();
    pin.transport.begun(30.0, true, false);
    pin.transport.duration = Some(200.0);
    pin.transport.subtitle = Some(0);
    pin.volume.level = 40;
    pin.volume.playing_at = 40;
    pin
}

/// A pin showing the film whose subtitles have just been copied is begun again, once, so the
/// frame it draws next is drawn with them — while a hover keeps the answer it was given.
///
/// This is the pin-only reload the user asked for: a pinned window has its film on screen
/// already, so a copy landing in the background is worth a stop and a start (see
/// `reload_pinned_subtitles`), which is exactly what a hover must not do under the pointer.
#[test]
fn a_pin_showing_the_film_restarts_its_player_when_the_copied_subtitles_are_ready() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let film = PathBuf::from("pin-subtitle-reload-film.mkv");
    stand_pin(Some(a_pin_playing(&film, (100, 80, 420, 320))));
    clear_restart_count();

    assert!(
        reload_pinned_subtitles(&film),
        "a pin showing the film is begun again for the copy that just landed"
    );
    assert_eq!(
        restart_count(),
        1,
        "one relaunch, at the second the bar is showing — the same player a track change \
         begins, with the copy now drawn (see `restart_pinned_player`)"
    );

    // The same answer about another film is not this pin's to act on: a copy landing for a file
    // the pin is not showing is a file the pin has nothing to reload for.
    assert!(
        !reload_pinned_subtitles(&PathBuf::from("pin-subtitle-reload-other.mkv")),
        "an answer about another film is dropped"
    );
    assert_eq!(restart_count(), 1, "and nothing is begun for it");

    // A video the engine draws is not this app's to reload: no extraction runs for one, and the
    // engine shows no subtitles at all for a copy to reach.
    let mut engine_video = create_loading_media(320, 240);
    engine_video.media_type = MediaType::NativeVideo;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(engine_video);
    }
    assert!(
        !reload_pinned_subtitles(&film),
        "a pin an engine is drawing is refused"
    );
    assert_eq!(restart_count(), 1, "and it begins nothing either");

    // And a pin with no file up is no window to reload.
    stand_pin(None);
    assert!(
        !reload_pinned_subtitles(&film),
        "no pin is nothing to begin again"
    );
    assert_eq!(restart_count(), 1, "and nothing is begun with none up");

    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    restore_media(previous_media);
    clear_restart_count();
}

/// A geometry answer for `film` carrying a ready copy, as the probe would have left it once the
/// extraction finished — what the take-up's question is read from (see `video_copy_ready`). No
/// files are written: the question is whether the answer holds a copy, not whether one is there
/// to open.
fn stand_a_ready_copy(film: &Path) {
    video_geometry_cache().insert(
        VideoGeometryKey {
            path: film.to_path_buf(),
            version: file_version(film),
        },
        ProbedGeometry::Measured(VideoGeometry {
            width: 1920,
            height: 1080,
            frame_width: 1920,
            frame_height: 1080,
            crop: None,
            duration: Some(200.0),
            subtitles: SubtitleStreams { count: 1, first: 0 },
            sidecar: None,
            derived: Some(DerivedSubtitles {
                tracks: vec![Some(PathBuf::from("pin-subtitle-reload-sub0.ass"))],
                fonts: None,
            }),
            subtitle_codecs: vec!["ass".to_string()],
            attachment_codecs: Vec::new(),
            subtitle_extraction_failed: false,
        }),
    );
}

/// A pin taken up on a player that was begun before the copy existed is begun again here, where
/// the copy is ready by the time the pin is made — the corner the ready-message cannot reach
/// because it has already been and gone (see `reload_adopted_subtitles`).
#[test]
fn an_adopted_pin_is_restarted_when_the_copy_it_predates_is_already_ready() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    // The switch on: this is the take-up's answer with the setting a film's
    // subtitles are wanted for, which is the answer it has always given (see
    // `VideoSubtitlesSetting`).
    let _setting = VideoSubtitlesSetting::stood_at(true);
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let film = PathBuf::from("pin-subtitle-reload-adopted.mkv");
    stand_a_ready_copy(&film);
    stand_pin(Some(a_pin_playing(&film, (100, 80, 420, 320))));

    // The hover this pin was taken up on was begun while the copy was still coming — the state
    // every launch writes down (see `note_player_subtitles`) — so the adopted player is drawing
    // no subtitles and there is a ready copy to begin it again for.
    note_player_subtitles(&film, false);
    clear_restart_count();

    assert!(
        reload_adopted_subtitles(&film),
        "a player begun before the copy existed is begun again now that it does"
    );
    assert_eq!(
        restart_count(),
        1,
        "one relaunch, at the second the bar is showing, the way a seek begins one"
    );

    // A player begun with its subtitles — the copies themselves, a sidecar, or the film's own
    // track — is left alone: there is nothing on screen to correct.
    note_player_subtitles(&film, true);
    assert!(
        !reload_adopted_subtitles(&film),
        "a player already drawing subtitles is not begun again"
    );
    assert_eq!(restart_count(), 1, "and nothing is begun for one");

    // And a film with no copy ready is not this either, whatever its player drew: what is
    // missing is the thing there would be something to reload for.
    let other = PathBuf::from("pin-subtitle-reload-adopted-none.mkv");
    stand_pin(Some(a_pin_playing(&other, (100, 80, 420, 320))));
    note_player_subtitles(&other, false);
    assert!(
        !reload_adopted_subtitles(&other),
        "no copy ready is nothing to reload for"
    );
    assert_eq!(restart_count(), 1, "and nothing is begun");

    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    restore_media(previous_media);
    clear_restart_count();
}

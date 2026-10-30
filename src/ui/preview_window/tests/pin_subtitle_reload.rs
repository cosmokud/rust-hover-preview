use super::super::subtitle_files::VideoSubtitlesSetting;
use super::*;

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

/// A pin showing the film whose subtitles have just been copied: **mpv is simply handed the file** and
/// draws it without the player, its window or its picture moving, and **FFmpeg's player is begun for
/// nothing at all** — the copy is preloaded for the next relaunch, which draws it out of the film's own
/// geometry (see `reload_pinned_subtitles` and `preload_the_subtitles_for_the_next_relaunch`).
///
/// This is the pin-only answer the user asked for: a pinned window has its film on screen already, and
/// a copy landing in the background must not swap the picture for it; a hover is deliberately never
/// begun again for one.
#[test]
fn a_pin_hands_its_late_copied_subtitles_to_mpv_and_preloads_them_for_ffplays_next_relaunch() {
    let _room = the_room_to_itself();
    let _setting = VideoSubtitlesSetting::stood_at(true);

    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let film = PathBuf::from("pin-subtitle-reload-film.mkv");
    let folder = a_copied_track(&film, &PathBuf::from("sub0.ass"));
    stand_pin(Some(a_pin_playing(&film, (100, 80, 420, 320))));
    clear_restart_count();

    // **mpv draws it from the file it is handed**: the player this app holds is the one it was, and
    // nothing is begun again for the copy.
    stand_mpv_in(Some(true));
    stand_ffplay_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    mpv_ipc::clear_noted();
    pin_trace::clear_noted();
    mpv_ipc::remember_pipe(101, "a-subtitle-pipe", false);

    assert!(
        reload_pinned_subtitles(&film),
        "a pin showing the film is handed the copy that just landed"
    );
    assert_eq!(
        restart_count(),
        0,
        "and no player is begun again: the picture on screen is the one the film plays on"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        101,
        "the player is the one it was, so there is no handover and no audio to hand over"
    );
    let commands = mpv_ipc::noted_lines();
    assert_eq!(
        commands.len(),
        1,
        "one command, and it is the file: {commands:?}"
    );
    assert!(
        commands[0].contains("\"sub-add\"")
            && commands[0].contains("sub0.ass")
            && commands[0].contains("\"select\""),
        "mpv is told to add the copy and draw it: {commands:?}"
    );
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("subtitle added without a restart")),
        "and the run says the copy cost no restart"
    );

    // **FFmpeg's player can be told nothing, so the copy waits for the next relaunch** — and the
    // picture on screen is not changed hands, nor the film decoded a second time, for it.
    clear_restart_count();
    mpv_ipc::forget_pipe(101);
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    VIDEO_PID.store(101, Ordering::SeqCst);
    pin_trace::clear_noted();
    assert!(
        reload_pinned_subtitles(&film),
        "the copy is answered on the engine that cannot be told as well"
    );
    assert_eq!(
        restart_count(),
        0,
        "and no player is begun for it: the copy is preloaded, not a reason for a swap"
    );
    assert_eq!(
        last_restart_reason(),
        None,
        "so nothing says a relaunch happened"
    );
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("subtitle preloaded for the next relaunch")),
        "and the run says the copy is waiting for one: {:?}",
        pin_trace::noted_lines()
    );

    // The relaunch the user asks for next is where it is activated: the launch's own `-vf` chain names
    // the copy the film's geometry holds, and one line says it is carrying it (see
    // `preload_the_subtitles_for_the_next_relaunch`).
    pin_trace::clear_noted();
    restart_pinned_player(&film, (100, 80, 420, 320), 30.0, false, RestartReason::Seek);
    let launched = take_last_spawned_command().expect("a player begun for the seek");
    assert!(
        launched.args.windows(2).any(|pair| pair[0] == "-vf"
            && pair[1].contains("subtitles=")
            && pair[1].contains("sub0.ass")),
        "the relaunch draws the copy: {:?}",
        launched.args
    );
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("subtitle carried into a relaunch")),
        "and the one line that says so is written: {:?}",
        pin_trace::noted_lines()
    );

    // The same answer about another film is not this pin's to act on: a copy landing for a file
    // the pin is not showing is a file the pin has nothing to preload for.
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
    mpv_ipc::clear_noted();
    pin_trace::clear_noted();
    mpv_ipc::forget_pipe(101);
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    let _ = std::fs::remove_dir_all(&folder);
}

/// A geometry answer for `film` carrying a ready copy, as the probe would have left it once the
/// extraction finished — and a real file for it, because a copy is only handed to a player that has one
/// to open. Answers the folder the test removes (see `video_copy_ready` and `video_subtitle_copy`).
fn a_copied_track(film: &Path, name: &Path) -> PathBuf {
    let folder = std::env::temp_dir().join(format!(
        "preview-subtitle-copy-{}",
        film.file_stem()
            .map(|stem| stem.to_string_lossy().into_owned())
            .unwrap_or_else(|| "film".to_string()),
    ));
    std::fs::create_dir_all(&folder).expect("a folder for the copied track is creatable");
    let copy = folder.join(name);
    std::fs::write(&copy, b"[Script Info]\n").expect("a stand-in copy is writable");
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
                tracks: vec![Some(copy)],
                fonts: None,
            }),
            subtitle_codecs: vec!["ass".to_string()],
            attachment_codecs: Vec::new(),
            subtitle_extraction_failed: false,
        }),
    );
    folder
}

/// A geometry answer for `film` carrying `count` subtitle streams and nothing else, as the probe
/// would have left it: what `video_subtitles` reads the pin's next track out of.
fn stand_subtitle_streams(film: &Path, count: usize) {
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
            subtitles: SubtitleStreams { count, first: 0 },
            sidecar: None,
            derived: None,
            subtitle_codecs: Vec::new(),
            attachment_codecs: Vec::new(),
            subtitle_extraction_failed: false,
        }),
    );
}

/// The `T` key steps a film's tracks both ways there are: FFmpeg's player is told a track by being
/// begun again with one, and mpv is told `cycle sub` down the pipe it listens on (see
/// `step_pinned_subtitle` and `mpv_ipc`).
///
/// Both routes are stood in rather than read off the machine: with mpv installed the default
/// engine would pick it, and each half of this is the answer one route gives.
#[test]
fn a_pinned_film_steps_its_subtitles_by_a_relaunch_on_ffplay_and_by_a_command_on_mpv() {
    let _room = the_room_to_itself();
    let _setting = VideoSubtitlesSetting::stood_at(true);

    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let film = PathBuf::from("pin-subtitle-step-film.mkv");
    stand_subtitle_streams(&film, 2);
    stand_pin(Some(a_pin_playing(&film, (100, 80, 420, 320))));
    clear_restart_count();

    // FFmpeg's player is told a track by being begun again with one, so the key is a relaunch.
    // mpv is stood out first, because with mpv installed the default engine would pick it and the
    // key would answer the mpv route before this test asked about the FFmpeg one (see `stand_mpv_in`).
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    step_pinned_subtitle();
    assert_eq!(
        restart_count(),
        1,
        "a film FFmpeg plays steps to the next track by a relaunch"
    );
    assert_eq!(
        pinned_subtitle(),
        Some(1),
        "and the choice is written down, as it is for every relaunch this one survives"
    );

    // mpv draws its own tracks and reports nothing back, so there is no number here to write
    // down: the step is `cycle sub` down the pipe the player listens on, and nothing is begun.
    clear_restart_count();
    mpv_ipc::clear_noted();
    stand_ffplay_in(Some(false));
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    // The FFmpeg half above began a player, which wrote its own pid down: this half is about the
    // player this app was holding before that.
    VIDEO_PID.store(101, Ordering::SeqCst);
    mpv_ipc::remember_pipe(101, "test-pipe", false);
    step_pinned_subtitle();
    assert_eq!(
        restart_count(),
        0,
        "mpv draws its own tracks, so the key begins nothing"
    );
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![format!(
            "{}\n",
            r#"{"command":["cycle","sub"],"async":true}"#
        )],
        "the step is mpv's own word for walking its tracks, written down the pipe it listens on"
    );
    assert_eq!(
        pinned_subtitle(),
        Some(1),
        "and it does not claim a track it never chose: mpv is the one that chose"
    );

    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    restore_media(previous_media);
    clear_restart_count();
}

/// A pin taken up on a player that was begun before the copy existed is corrected here, where
/// the copy is ready by the time the pin is made — the corner the ready-message cannot reach
/// because it has already been and gone (see `reload_adopted_subtitles`).
///
/// **mpv is handed the copy and FFmpeg's player is given the copy to carry**, and the difference is the
/// whole of what a player that can be spoken to is worth here: the same corner costs a film mpv plays
/// nothing, and the one it draws is not begun again for a subtitle either (see
/// `preload_the_subtitles_for_the_next_relaunch`).
#[test]
fn an_adopted_pin_is_handed_the_copy_it_predates_without_a_relaunch() {
    let _room = the_room_to_itself();

    // The switch on: this is the take-up's answer with the setting a film's
    // subtitles are wanted for, which is the answer it has always given (see
    // `VideoSubtitlesSetting`).
    let _setting = VideoSubtitlesSetting::stood_at(true);
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let film = PathBuf::from("pin-subtitle-reload-adopted.mkv");
    let folder = a_copied_track(&film, &PathBuf::from("sub0.ass"));
    stand_pin(Some(a_pin_playing(&film, (100, 80, 420, 320))));

    // The hover this pin was taken up on was begun while the copy was still coming — the state
    // every launch writes down (see `note_player_subtitles`) — so the adopted player is drawing
    // no subtitles and there is a ready copy to hand it.
    note_player_subtitles(&film, false);
    clear_restart_count();

    // **mpv is handed the file**: the player adopted by the take-up is left running, which is the
    // second decode the take-up exists to avoid.
    stand_mpv_in(Some(true));
    stand_ffplay_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    mpv_ipc::clear_noted();
    mpv_ipc::remember_pipe(101, "an-adopted-subtitle-pipe", false);
    assert!(
        reload_adopted_subtitles(&film),
        "a player begun before the copy existed is handed it now that it does"
    );
    assert_eq!(restart_count(), 0, "and no player is begun again for it");
    assert_eq!(
        mpv_ipc::noted_lines().len(),
        1,
        "the copy is the one command the adopted player is sent"
    );

    // **FFmpeg's player is begun for nothing at all**: the copy is preloaded, and the relaunch the user
    // asks for next is what draws it (see `preload_the_subtitles_for_the_next_relaunch`).
    clear_restart_count();
    mpv_ipc::forget_pipe(101);
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    assert!(
        reload_adopted_subtitles(&film),
        "a player that can be told nothing has the copy to carry"
    );
    assert_eq!(
        restart_count(),
        0,
        "and no player is begun for it: the copy waits for a relaunch of the user's"
    );
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("subtitle preloaded for the next relaunch")),
        "and the run says so: {:?}",
        pin_trace::noted_lines()
    );

    // A player begun with its subtitles — the copies themselves, a sidecar, or the film's own
    // track — is left alone: there is nothing on screen to correct.
    note_player_subtitles(&film, true);
    assert!(
        !reload_adopted_subtitles(&film),
        "a player already drawing subtitles has nothing to preload"
    );
    assert_eq!(restart_count(), 0, "and nothing is begun for one");

    // And a film with no copy ready is not this either, whatever its player drew: what is
    // missing is the thing there would be something to reload for.
    let other = PathBuf::from("pin-subtitle-reload-adopted-none.mkv");
    stand_pin(Some(a_pin_playing(&other, (100, 80, 420, 320))));
    note_player_subtitles(&other, false);
    assert!(
        !reload_adopted_subtitles(&other),
        "no copy ready is nothing to reload for"
    );
    assert_eq!(restart_count(), 0, "and nothing is begun");

    // **A hosted hover's moved player is not begun again either**: it was moved into the pin rather
    // than launched there, and the copy rides the next relaunch the user asks for.
    stand_pin(Some(a_pin_playing(&film, (100, 80, 420, 320))));
    let host = unsafe { hover_film_host::take_up_the_hover_film_host((100, 80, 420, 320)) }
        .expect("a host standing in for the moved hover");
    // The take-up places the host on the desktop, and a host an earlier test left on screen is a black
    // rectangle of the app's over it: it is composited as nothing for the whole of this test (see
    // `keep_a_window_of_the_suite_unseen`).
    unsafe { keep_a_window_of_the_suite_unseen(host) };
    hover_film_host::note_hover_film_host_player(101, &film);
    note_player_subtitles(&film, false);
    assert!(
        reload_adopted_subtitles(&film),
        "the copy is preloaded for a player that cannot be told"
    );
    assert_eq!(restart_count(), 0, "and nothing is begun");
    unsafe { hover_film_host::take_down_the_hover_film_host() };
    assert!(!unsafe { IsWindow(host) }.as_bool(), "the stand goes");

    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    restore_media(previous_media);
    clear_restart_count();
    clear_pending_subtitles();
    mpv_ipc::clear_noted();
    mpv_ipc::forget_pipe(101);
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    let _ = std::fs::remove_dir_all(&folder);
}

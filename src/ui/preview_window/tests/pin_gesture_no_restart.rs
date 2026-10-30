use super::*;

// **A gesture does not touch the film.** The pinned film is a thumbnail of a source window nothing
// a hand does moves, so a press, a move, a drag, a resize and their ends are this app's own window
// work: no key is posted, no player is killed or begun again, nothing is hidden, and no reason line
// is written.
// The only relaunches left are the ones a user action needs on a player that cannot be told
// anything — a seek, a level, a subtitle, the end of a file, and the engine's own take-up (see
// `RestartReason`).

/// A video behind the pin, with the given FFmpeg/mpv answer: the media slot the press road reads.
fn stand_video_media(ffplay: bool, mpv: bool) -> Option<MediaData> {
    stand_ffplay_in(Some(ffplay));
    stand_mpv_in(Some(mpv));
    let previous = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }
    previous
}

/// A playing pin on a film, as a gesture finds it.
fn a_playing_film_pin(film: &str, content: ScreenRegion) -> PinnedPreview {
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = PathBuf::from(film);
    pin.transport.begun(30.0, true, false);
    pin.transport.duration = Some(200.0);
    pin.volume.level = 40;
    pin.volume.playing_at = 40;
    pin
}

/// Nothing of the film was touched by the gesture: the player is the one that was there, nothing
/// was begun, and nothing of the picture was hidden.
fn nothing_of_the_film_was_touched(pid: u32, what: &str) {
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        pid,
        "the player is the one the gesture found, with no kill behind it ({what})"
    );
    assert_eq!(
        restart_count(),
        0,
        "and no player was begun again for it ({what})"
    );
    assert_eq!(
        last_restart_reason(),
        None,
        "so no reason line was written, and there is no reason a gesture could write ({what})"
    );
}

/// **A press that never moves is a press and a release, and it does nothing to the film** — on
/// either route: same pid, no relaunch, no reason line, no cover, and the pin is still the pin that
/// was pressed.
#[test]
fn a_press_with_no_movement_does_nothing_to_the_film_on_either_route() {
    let _room = the_room_to_itself();

    for (ffplay, mpv, what) in [
        (true, false, "FFmpeg's route"),
        (false, true, "mpv's route"),
    ] {
        let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
        let previous_pin = take_pin_for_a_test();
        let previous_media = stand_video_media(ffplay, mpv);
        clear_restart_count();

        let content = (100, 80, 420, 320);
        let film = format!("a-click-does-nothing-{ffplay}.mkv");
        let pin = a_playing_film_pin(&film, content);
        let route = video_route(&pin.path);
        assert!(
            matches!(route, VideoRoute::Ffplay) == ffplay
                && matches!(route, VideoRoute::Mpv) == mpv,
            "the test stands on {what}, and the route says {route:?}"
        );
        stand_pin(Some(pin));

        let window = a_window_at(content);
        begin_pin_drag(HWND(0x1000 as *mut _), &window, PinDragAction::Move, true);
        assert!(
            pin_is_dragging(),
            "the press is the pin's own and the window it stands in is selected ({what})"
        );
        assert!(pinned(), "the pin is still up after the press ({what})");
        nothing_of_the_film_was_touched(101, what);

        // No step of a hand: the release is the whole of the rest of it.
        assert!(
            finish_pin_drag(HWND(0x1000 as *mut _), &window),
            "the release is the drag's own ({what})"
        );
        nothing_of_the_film_was_touched(101, what);
        assert!(
            pinned() && !pin_is_dragging(),
            "and the pin is still up and no longer carried, exactly as a click left it ({what})"
        );

        clear_restart_count();
        stand_mpv_in(None);
        stand_ffplay_in(None);
        stand_pin(previous_pin);
        VIDEO_PID.store(previous_pid, Ordering::SeqCst);
        if let Ok(mut media) = CURRENT_MEDIA.lock() {
            *media = previous_media;
        }
    }
}

/// **A full drag and a resize of an FFmpeg film leave the player alone from the press to the
/// settle**: the same pid throughout, no reason line, no cover, and the thumbnail road is never
/// asked for a source placement — so the film the pin is drawing has no gap in it.
#[test]
fn a_drag_and_a_resize_keep_the_same_player_throughout() {
    let _room = the_room_to_itself();

    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media(true, false);
    clear_restart_count();
    pin_film_thumbnail::clear_the_source_placements();

    let content = (100, 80, 420, 320);
    stand_pin(Some(a_playing_film_pin(
        "a-drag-touches-nothing.mkv",
        content,
    )));
    assert_eq!(
        video_route(&PathBuf::from("a-drag-touches-nothing.mkv")),
        VideoRoute::Ffplay,
        "the test is about FFmpeg's player, which is the route the kill road was for"
    );

    for (action, travelling, what) in [
        (PinDragAction::Move, (60, -40), "a drag"),
        (
            PinDragAction::Resize(edge(false, false, true, false)),
            (80, 0),
            "a resize",
        ),
    ] {
        let before = pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| pin.window_box()))
            .expect("a pin up");
        let window = a_window_at(before);
        begin_pin_drag(HWND(0x1000 as *mut _), &window, action, true);
        assert!(pin_is_dragging(), "{what} begins a drag");
        nothing_of_the_film_was_touched(101, what);

        for step in 1..=4 {
            super::probes::step_the_pin_drag_by(travelling.0 / 4 * step, travelling.1 / 4 * step);
            unsafe { apply_pin_drag(HWND(0x1000 as *mut _)) };
            nothing_of_the_film_was_touched(101, &format!("{what}, step {step}"));
        }

        let after = pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| pin.window_box()))
            .expect("a pin up");
        assert_ne!(after, before, "{what} really moved the window");

        assert!(
            finish_pin_drag(HWND(0x1000 as *mut _), &window),
            "{what} is let go of through its release"
        );
        nothing_of_the_film_was_touched(101, &format!("{what}, let go"));
        if matches!(action, PinDragAction::Resize(_)) {
            assert!(
                PIN_BOX_REQUEST
                    .lock()
                    .ok()
                    .and_then(|request| *request)
                    .is_some(),
                "a resize asks for its media to be laid out again at the box it settled on"
            );
            drain_pin_box_request();
        }
    }

    assert_eq!(
        pin_film_thumbnail::the_source_placements(),
        0,
        "and the thumbnail's source is never placed for a gesture: it was placed once, at the \
         film's own size, and a gesture moves only the box (see `pin_film_thumbnail`)"
    );

    clear_restart_count();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// **A level let go of on an FFmpeg film is still one player begun again**, and it says why:
/// `player restarted for volume`. The press itself told the player nothing.
#[test]
fn a_volume_change_restarts_once_and_says_why() {
    let _room = the_room_to_itself();

    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media(true, false);
    clear_restart_count();

    let mut pin = a_playing_film_pin("a-volume-change-restarts.mkv", (100, 80, 420, 320));
    pin.volume.level = 70;
    pin.volume.playing_at = 40;
    stand_pin(Some(pin));

    settle_pin_volume();

    assert_eq!(
        restart_count(),
        1,
        "FFmpeg's player takes a level only by being begun at one, so the change is one relaunch"
    );
    assert_eq!(
        last_restart_reason(),
        Some(RestartReason::Volume),
        "and the line says what it was for"
    );
    clear_restart_count();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

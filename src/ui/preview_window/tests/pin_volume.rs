use super::*;

// What the wheel hook is handed a message as (see `wheel_hook_proc`).
use windows::Win32::Foundation::{LPARAM, LRESULT, POINT, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::{HC_ACTION, MSLLHOOKSTRUCT, WM_MOUSEWHEEL};

/// The volume popup belongs to the button that opened it and to nothing else: it is kept while
/// the pointer is anywhere in that control — the panel itself, or the strip the button sits in
/// — held while the knob is being dragged, and put away once the pointer has left the whole of
/// it, which is a repaint the window owes.
#[test]
fn a_volume_popup_is_kept_on_its_control_and_put_away_once_the_pointer_has_left_it() {
    let _room = the_room_to_itself();
    let now = Instant::now();
    let content = (100, 100, 500, 400);
    let far = Some((-1000, -1000));

    let mut pin = overlay_pin(content, PinChrome::always());
    pin.transport_bar = true;
    pin.volume = PinVolume {
        level: 40,
        playing_at: 40,
        open: true,
        dragging: false,
        ..Default::default()
    };

    let window = pin.window_box();
    let (width, height) = pin.window_size();
    let strip = pinned_transport_height(pin.dpi, true);
    let popup = pin_chrome::volume_popup_layout(width, (height - strip).max(0), strip, pin.dpi);
    let on_the_popup = Some((
        window.0 + (popup.panel.left + popup.panel.right) / 2,
        window.1 + (popup.panel.top + popup.panel.bottom) / 2,
    ));

    // A hand on the panel keeps it, and the strip it came out of stays showing with it — the
    // popup is drawn in this window's own rows above that strip, and a bar that went away
    // underneath it would take the button that opened it with it.
    pin.chrome = PinChrome {
        caption: false,
        bar: false,
        until: None,
    };
    assert!(refresh_pin_chrome(&mut pin, now, on_the_popup));
    assert!(pin.volume.open);
    assert!(pin.chrome.bar);

    // The button's own strip keeps it too: the pointer is on the control, one row below the
    // panel.
    let at_the_button = Some((window.0 + width - 20, window.1 + height - 15));
    assert!(!refresh_pin_chrome(&mut pin, now, at_the_button));
    assert!(pin.volume.open);

    // A knob being held keeps it wherever the pointer has got to: a drag that has left the
    // panel is a level being taken to an end, not a popup being dismissed.
    pin.volume.dragging = true;
    assert!(!refresh_pin_chrome(&mut pin, now, far));
    assert!(pin.volume.open);

    // And a pointer away from the whole of it puts it away, which is the repaint.
    pin.volume.dragging = false;
    assert!(refresh_pin_chrome(&mut pin, now, far));
    assert!(!pin.volume.open);

    // A window whose box has changed has no popup left to put away, and asking again is not a
    // change: a popup that has gone costs nothing to keep gone.
    assert!(!refresh_pin_chrome(&mut pin, now, far));
    assert!(!pin.volume.open);
}

/// The level a pin is given is the tray's at the moment it is taken up, and what a hand does to
/// it afterwards is that window's own: nothing of it is written back to `Volume → Video`.
#[test]
fn a_level_moved_on_a_pin_is_the_pins_and_not_the_setting() {
    let _room = the_room_to_itself();

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    // The `Remember` rows are the test's subject by negation: with them off, a level moved
    // on the pin is the pin's own and the settings are left where they were, whatever the
    // machine the test runs on remembers (see `PinLevels::remembering`).
    let mut _levels = PinLevels::set(100, 100);
    _levels.remembering(false, false);

    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::NativeVideo;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(video);
    }
    stand_pin(Some(overlay_pin((100, 100, 420, 340), PinChrome::always())));

    let setting = current_video_volume();
    set_pin_volume(12);

    assert_eq!(pinned_volume_level(), 12, "the pin is playing at 12");
    assert_eq!(
        current_video_volume(),
        setting,
        "`Volume → Video` is left exactly where the user put it"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// The tray's two levels, for a test that is about which of them a pin is holding. Put back when
/// the guard is dropped, so a test beside this one is answered with the machine's own.
struct PinLevels {
    video: u32,
    audio: u32,
    /// The two `Remember` rows beside them, put where `remembering` says and put back with the
    /// levels (see `pin_level_is_remembered`).
    remember: Option<(bool, bool)>,
}

impl PinLevels {
    fn set(video: u32, audio: u32) -> Self {
        let mut config = CONFIG.lock().expect("the configuration");
        let was = PinLevels {
            video: config.video_volume,
            audio: config.audio_volume,
            remember: None,
        };
        config.video_volume = video;
        config.audio_volume = audio;

        was
    }

    /// The two `Remember` rows under the two levels, for a test whose subject is what a level
    /// moved on a pin's own knob outlives (see `PinVolume::kept`).
    fn remembering(&mut self, audio: bool, video: bool) {
        if let Ok(mut config) = CONFIG.lock() {
            config.remember_audio_volume = audio;
            config.remember_video_volume = video;
        }
        self.remember = Some((audio, video));
    }

    /// The two levels moved again under a guard already taken, which is what reading the setting
    /// back out of the file looks like to the pin holding it.
    fn set_levels(&mut self, video: u32, audio: u32) {
        if let Ok(mut config) = CONFIG.lock() {
            config.video_volume = video;
            config.audio_volume = audio;
        }
    }
}

impl Drop for PinLevels {
    fn drop(&mut self) {
        if let Ok(mut config) = CONFIG.lock() {
            config.video_volume = self.video;
            config.audio_volume = self.audio;
            if let Some((audio, video)) = self.remember {
                config.remember_audio_volume = audio;
                config.remember_video_volume = video;
            }
        }
    }
}

/// A pin's level is the level of the kind it is playing: a sound's card at `Volume → Audio` is
/// not what a film stepped onto from it is played at, and the film the tray has muted is
/// played at `Volume → Video` — silent, whatever the sound it replaces was at.
#[test]
fn a_film_stepped_onto_from_a_sound_is_played_at_the_films_own_level() {
    let _room = the_room_to_itself();

    let _levels = PinLevels::set(0, 100);
    let previous_pin = take_pin_for_a_test();

    // The pin is up on a sound, which the tray has at 100%.
    stand_pin(Some(PinnedPreview {
        volume: PinVolume {
            audio: true,
            level: 100,
            playing_at: 100,
            ..Default::default()
        },
        ..sound_pin()
    }));

    assert_eq!(
        pinned_audio_level(),
        100,
        "the sound is at `Volume → Audio`"
    );
    assert_eq!(
        pinned_volume_level(),
        0,
        "the film that replaces it is at `Volume → Video`, so it is silent"
    );

    // A knob turned on a film's own bar is still that film's, and is kept for the next film —
    // which is the half of this that is not about the two settings being one.
    stand_pin(Some(PinnedPreview {
        volume: PinVolume {
            audio: false,
            level: 12,
            playing_at: 12,
            ..Default::default()
        },
        ..overlay_pin((100, 100, 420, 340), PinChrome::always())
    }));

    assert_eq!(
        pinned_volume_level(),
        12,
        "a level moved on a film's bar is the next film's too"
    );

    stand_pin(previous_pin);
}

/// The other half of the same thing, and the half a knob has to do for it: a pin that walked off
/// a sound is still holding the sound's level, and a knob turned on the film's bar makes that
/// level the film's own. Without this the level moved is read as the sound's and dropped, so the
/// film a hand set the knob on is the one film it is kept for (see `set_pin_volume`).
#[test]
fn a_level_moved_on_a_films_bar_is_the_films_and_not_the_sounds() {
    let _room = the_room_to_itself();

    let _levels = PinLevels::set(0, 100);
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    // A film the media engine is drawing, with the pin behind it still holding the sound's level.
    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::NativeVideo;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(video);
    }
    stand_pin(Some(PinnedPreview {
        volume: PinVolume {
            audio: true,
            level: 100,
            playing_at: 100,
            ..Default::default()
        },
        ..overlay_pin((100, 100, 420, 340), PinChrome::always())
    }));

    // The knob is turned on the film's bar, so the level in hand is the film's.
    set_pin_volume(30);

    assert_eq!(
        pinned_volume_level(),
        30,
        "a level moved on a film's bar is the pin's, and is kept for the next film"
    );
    assert_eq!(
        pinned_audio_level(),
        100,
        "and the sound it walked off is still at `Volume → Audio`, which the knob was not on"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// The round trip a sound's card makes through a film, both ways round: with
/// `Volume → Audio`'s `Remember` on a level a hand left on that card is the level the next sound
/// is played at and a film stepped onto in between does not throw it away, and with the row off
/// the knob goes with the sound it was turned on and every switch is the setting's own.
///
/// The step in the middle of the remembered half is the whole of it.
/// `remember_pin_volume` puts the level into the configuration as the knob moves, but
/// `save_remembered_pin_volume` writes the file only where the two disagree — and they never do,
/// because the first of those has just made them agree — so the file still says what the tray
/// said before the drag. Anything that reads the setting back hands the pin a tray at 10% again:
/// the watcher after `config.ini` is edited, or the next run. One slot cannot answer for both
/// kinds at once, so a pin holding only the film's 0% by then plays the returning sound at 10%,
/// out of a card its own knob was left at 100%.
#[test]
fn a_sound_stepped_back_onto_from_a_film_is_at_the_level_its_own_card_was_left_at() {
    // Serialized with the pin tests: this stands pins and takes the media, which unlocked is a
    // window another test's pin is standing in.
    let _room = the_room_to_itself();

    let mut levels = PinLevels::set(0, 10);
    levels.remembering(true, false);
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let mut sound = create_loading_media(320, 240);
    sound.media_type = MediaType::Audio;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(sound);
    }

    // The pin comes up on a sound at the tray's 10%, and a hand turns its knob to 100%.
    stand_pin(Some(PinnedPreview {
        volume: PinVolume {
            audio: true,
            level: 10,
            playing_at: 10,
            ..Default::default()
        },
        ..sound_pin()
    }));
    set_pin_volume(100);

    // A film is stepped onto, at `Volume → Video`, which the tray has at 0%.
    stand_pin(Some(PinnedPreview {
        volume: pin_volume_taken_up(pin_holding(), false),
        ..overlay_pin((100, 100, 420, 340), PinChrome::always())
    }));
    assert_eq!(
        pinned_volume_level(),
        0,
        "the film is at `Volume → Video`, so it is silent"
    );

    // The setting is read back out of the file the release never wrote, and says 10% again.
    levels.set_levels(0, 10);

    // And the sound is stepped onto from it, which is the whole of the round trip.
    stand_pin(Some(PinnedPreview {
        volume: pin_volume_taken_up(pin_holding(), true),
        ..sound_pin()
    }));
    assert_eq!(
        pinned_audio_level(),
        100,
        "the sound is at the level its own card was left at, and not at what the setting says"
    );

    // The same round trip with the row off, where none of the above is written down at all: a
    // knob belongs to the window it was turned on and to no other file whatever, so the setting
    // is what every switch is answered by (see `pin_level_is_remembered`).
    levels.remembering(false, false);
    stand_pin(Some(PinnedPreview {
        volume: PinVolume {
            audio: true,
            level: 10,
            playing_at: 10,
            ..Default::default()
        },
        ..sound_pin()
    }));
    set_pin_volume(100);
    stand_pin(Some(PinnedPreview {
        volume: pin_volume_taken_up(pin_holding(), false),
        ..overlay_pin((100, 100, 420, 340), PinChrome::always())
    }));
    assert_eq!(
        current_audio_volume(),
        10,
        "and nothing was written down, so there is nothing to read back either"
    );

    stand_pin(Some(PinnedPreview {
        volume: pin_volume_taken_up(pin_holding(), true),
        ..sound_pin()
    }));
    assert_eq!(
        pinned_audio_level(),
        10,
        "the knob's 100% went with the sound it was turned on, and the setting is what is left"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// What the pin that is up is holding, as the take-up reads it before the file it is coming up
/// on is known.
fn pin_holding() -> Option<PinVolume> {
    pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.volume))
}

/// `Volume → Audio`'s `Remember` is a claim about `config.ini`, and this is the claim: the level
/// in hand is put where the file will be asked for it again, and the file is written.
///
/// The level is already in the configuration before the question is asked, and it always is —
/// `remember_pin_volume` put it there on every step of the drag, and again on the click that
/// moved nothing before the release — so a question shaped as "is the file out of step with the
/// level?" is answered no every time, and the file was never written by either of the two rows
/// (`a_level_moved_on_a_pin_is_the_pins_and_not_the_setting` covers the half that keeps the
/// level in hand, which is a different thing entirely from writing it down).
#[test]
fn a_remembered_level_is_written_even_where_the_configuration_already_holds_it() {
    let mut config = AppConfig {
        remember_audio_volume: true,
        remember_video_volume: true,
        audio_volume: 100,
        video_volume: 0,
        ..Default::default()
    };

    assert!(
        put_remembered_pin_volume(&mut config, Some(MediaType::Audio), 100),
        "the level is in the configuration already, which is what a drag is, and the file is \
             written all the same"
    );

    assert!(
        put_remembered_pin_volume(&mut config, Some(MediaType::Video), 0),
        "and a film's own row is asked the same question of its own level"
    );

    // A level that is not remembered is the window's own and there is nothing to write down.
    config.remember_audio_volume = false;
    config.remember_video_volume = false;
    assert!(
        !put_remembered_pin_volume(&mut config, Some(MediaType::Audio), 100),
        "with the row off the knob moves the window alone"
    );

    // And a kind with no row of its own is neither half's answer.
    assert!(
        !put_remembered_pin_volume(&mut config, Some(MediaType::Pdf), 100),
        "a page has no level to remember"
    );
}

/// A level moved on a pinned film: mpv is told it while it runs, and FFmpeg's player is told it at
/// its mixer session — the gate this ticket turns on, asserted on both routes.
///
/// Neither route's settle begins a player: mpv takes the level as it plays, so the knob itself is
/// the write, and FFmpeg's level is its session's (see `set_pin_volume` and `settle_pin_volume`).
/// The `Remember` rows are stood off so that a test never writes the configuration file a level
/// would be remembered into.
#[test]
fn a_level_moved_on_a_pinned_player_is_told_to_it_and_begins_nothing() {
    let _room = the_room_to_itself();

    let mut levels = PinLevels::set(100, 50);
    levels.remembering(false, false);
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(video);
    }

    let content = (100, 100, 420, 340);
    let film = PathBuf::from("pin-live-volume-film.mkv");
    let stand = |level: u32, playing_at: u32| {
        let mut pin = overlay_pin(content, PinChrome::always());
        pin.path = film.clone();
        pin.transport.begun(30.0, true, false);
        pin.transport.duration = Some(200.0);
        pin.volume.level = level;
        pin.volume.playing_at = playing_at;
        stand_pin(Some(pin));
    };

    clear_restart_count();
    mpv_ipc::clear_noted();

    // FFmpeg's player takes a level at its mixer session, so nothing is begun again.
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    stand(70, 40);
    set_pin_volume(70);
    assert!(
        mpv_ipc::noted_lines().is_empty(),
        "FFmpeg's player has no pipe, so nothing is written to one"
    );
    settle_pin_volume();
    assert_eq!(
        restart_count(),
        0,
        "FFmpeg's player takes the level at its mixer session, so none is begun again"
    );
    assert_eq!(
        pinned_playback_state().map(|(_, _, _, volume)| volume.playing_at),
        Some(70),
        "and the claim moves, because the player was told"
    );

    // mpv is told the level while it runs, and neither the knob nor the settle begins anything.
    clear_restart_count();
    mpv_ipc::clear_noted();
    stand_ffplay_in(Some(false));
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    // The ffplay half above began a player, which wrote its own pid down: this half is about the
    // player this app was holding before that.
    VIDEO_PID.store(101, Ordering::SeqCst);
    mpv_ipc::remember_pipe(101, "test-pipe", false);
    stand(70, 40);
    set_pin_volume(70);
    flush_pin_volume_level();
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![format!(
            "{}\n",
            r#"{"command":["set_property","volume",70],"async":true}"#
        )],
        "the level is set on mpv's own property as the knob moves, which is the whole point of \
         the route"
    );
    assert_eq!(restart_count(), 0, "and the knob begins nothing");

    // The settle the knob's release makes is a no-op on mpv even where the claim is stale.
    stand(70, 40);
    settle_pin_volume();
    assert_eq!(
        restart_count(),
        0,
        "mpv takes a level while it runs, so its settle begins nothing again"
    );

    // The machine's own engine: the level is its own while it runs, so the knob is the write and
    // nothing is begun again — the arm this pins alongside FFmpeg's player's relaunch.
    stand_native_video_media();
    clear_restart_count();
    mpv_ipc::clear_noted();
    stand(70, 40);
    set_pin_volume(70);
    assert_eq!(
        restart_count(),
        0,
        "the engine takes the level while it runs, so the knob begins nothing"
    );
    assert!(
        mpv_ipc::noted_lines().is_empty(),
        "and the engine listens on no pipe for a level to be written to"
    );

    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
}

/// **A level dragged up from nothing on a pinned mpv film turns the player's audio on**, which is
/// what makes the drag do anything at all: a film begun at `Volume → Video` 0% is given no audio
/// rather than a level of nothing (see `mpv_args`), so a level set on its player is one nothing
/// plays at and the knob that moved it went on being silent. The track is asked back for by the
/// player's own road — `aid` set to `auto`, after the level so that the audio opens at the level
/// in hand (see `mpv_ipc::set_volume`) — and the asking is the player's own fact rather than the
/// drag's: the rest of the drag, and the settle, are the level alone.
///
/// The `Remember` rows are stood off so that a test never writes the configuration file a level
/// would be remembered into.
#[test]
fn a_level_moved_up_from_nothing_on_a_pinned_mpv_asks_the_track_back() {
    let _room = the_room_to_itself();

    let mut levels = PinLevels::set(0, 50);
    levels.remembering(false, false);
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(video);
    }

    let content = (100, 100, 420, 340);
    let film = PathBuf::from("pin-muted-launch-film.mkv");
    let stand = |level: u32, playing_at: u32| {
        let mut pin = overlay_pin(content, PinChrome::always());
        pin.path = film.clone();
        pin.transport.begun(30.0, true, false);
        pin.transport.duration = Some(200.0);
        pin.volume.level = level;
        pin.volume.playing_at = playing_at;
        stand_pin(Some(pin));
    };

    // The player behind the pin is one a level of nothing began — a pipe to be told on, and no
    // track at all, which is the launch `mpv_args` makes of that level.
    stand_ffplay_in(Some(false));
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    VIDEO_PID.store(101, Ordering::SeqCst);
    mpv_ipc::remember_pipe(101, "a-muted-launch-pipe", true);
    clear_restart_count();
    mpv_ipc::clear_noted();
    stand(0, 0);

    // A hand takes the knob up, and the level is heard: the property first, and the track the
    // level opens at asked for after it.
    set_pin_volume(60);
    flush_pin_volume_level();
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![
            format!(
                "{}\n",
                r#"{"command":["set_property","volume",60],"async":true}"#
            ),
            format!(
                "{}\n",
                r#"{"command":["set_property","aid","auto"],"async":true}"#
            ),
        ],
        "a level above nothing on a player begun at nothing is the level and the track together"
    );
    assert_eq!(
        pinned_playback_state().map(|(_, _, _, volume)| volume.playing_at),
        Some(60),
        "and the claim of what the player is playing at moves only with a level it was told"
    );

    // The rest of the drag is the level alone — the track is a fact of the player now — and the
    // release begins nothing: mpv took the level as it played.
    set_pin_volume(80);
    flush_pin_volume_level();
    assert_eq!(
        mpv_ipc::noted_lines()[2..],
        [format!(
            "{}\n",
            r#"{"command":["set_property","volume",80],"async":true}"#
        )],
        "the track is asked for once per player, not once per step of a drag"
    );
    clear_restart_count();
    settle_pin_volume();
    assert_eq!(
        restart_count(),
        0,
        "mpv took the level while it ran, so the settle begins nothing"
    );

    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
}

/// A level moved on one of the tray's two `Volume` rows is given to the pin that is already
/// up, which is the one road a pin's own level does not open on its own.
///
/// A pin plays at the level it was taken up at and at whatever a hand has done on its knob since,
/// and a row moved in the tray used to be the setting's alone: the film on screen went on being
/// heard at what it was until the pin was closed and taken up again, which reads as a tray that
/// has stopped working (see `tray_volume_moved`). Every route answers with what it always has,
/// because it is the knob's own road and not a second one. The `Remember` rows are stood off so
/// that a test never writes the configuration file a level would be remembered into.
#[test]
fn a_level_moved_in_the_tray_is_given_to_the_pin_that_is_up() {
    let _room = the_room_to_itself();

    let mut levels = PinLevels::set(10, 50);
    levels.remembering(false, false);
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(video);
    }

    let content = (100, 100, 420, 340);
    let film = PathBuf::from("tray-live-volume-film.mkv");
    let stand = |level: u32, playing_at: u32| {
        let mut pin = overlay_pin(content, PinChrome::always());
        pin.path = film.clone();
        pin.transport.begun(30.0, true, false);
        pin.transport.duration = Some(200.0);
        pin.volume.level = level;
        pin.volume.playing_at = playing_at;
        stand_pin(Some(pin));
    };

    // FFmpeg's player, whose level is its mixer session's: the film the tray has turned down to
    // nothing is told at the mixer, and nothing is begun again.
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    clear_restart_count();
    stand(100, 100);
    levels.set_levels(0, 50);
    tray_volume_moved(false);
    assert_eq!(pinned_volume_level(), 0, "the pin is playing at 0%");
    assert_eq!(
        restart_count(),
        0,
        "FFmpeg's player takes the level at its mixer session, so the row begins nothing"
    );
    assert_eq!(
        pinned_playback_state().map(|(_, _, _, volume)| volume.playing_at),
        Some(0),
        "and the claim moves with the level the player was told"
    );

    // mpv's, which can be told one while it plays: the level is written down its own pipe and
    // nothing is begun again.
    clear_restart_count();
    stand_ffplay_in(Some(false));
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    // The player above was begun again, which wrote its own pid down: this half is about the one
    // this app was holding before that.
    VIDEO_PID.store(101, Ordering::SeqCst);
    mpv_ipc::remember_pipe(101, "tray-live-volume-pipe", false);
    mpv_ipc::clear_noted();
    stand(0, 0);
    levels.set_levels(100, 50);
    tray_volume_moved(false);
    assert_eq!(pinned_volume_level(), 100, "the pin is playing at 100%");
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![format!(
            "{}\n",
            r#"{"command":["set_property","volume",100],"async":true}"#
        )],
        "mpv is told the level while it runs, which is what makes the row free for it"
    );
    assert_eq!(restart_count(), 0, "and nothing is begun again");

    // And the other half's row, which is about a sound: the film under it is left exactly where
    // it is, the same answer a film stepped onto from a sound is given (see `pinned_level`).
    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    clear_restart_count();
    stand(100, 100);
    levels.set_levels(100, 0);
    tray_volume_moved(true);
    assert_eq!(
        pinned_volume_level(),
        100,
        "a sound's row is not a hand on the film on screen"
    );
    assert_eq!(restart_count(), 0);
    assert!(
        mpv_ipc::noted_lines().is_empty(),
        "and nothing is written to a pipe for it"
    );

    // And a sound's own row, which is the same road for a card: the level is the pin's own at
    // once, and the player a card begins is owed the new one — the clock behind a card and its
    // player are the loop's, so this is the loop's flag rather than a player begun here (see
    // `settle_pinned_audio_volume`).
    let mut sound = create_loading_media(320, 240);
    sound.media_type = MediaType::Audio;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(sound);
    }
    PIN_AUDIO_VOLUME.store(false, Ordering::SeqCst);
    stand_pin(Some(PinnedPreview {
        volume: PinVolume {
            audio: true,
            level: 100,
            playing_at: 100,
            ..Default::default()
        },
        ..sound_pin()
    }));
    levels.set_levels(50, 0);
    tray_volume_moved(true);
    assert_eq!(
        pinned_audio_level(),
        0,
        "the sound's own row is a hand on the card under it, and it is at 0%"
    );
    assert!(
        PIN_AUDIO_VOLUME.load(Ordering::Acquire),
        "and the loop is told to begin the player a card has at the new level"
    );
    PIN_AUDIO_VOLUME.store(false, Ordering::SeqCst);

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
}

/// A hand holding the knob owns the level: no publisher moves the thumb under the pointer, the
/// tray's own `Volume` row included (see `tray_volume_moved`).
#[test]
fn a_tray_row_moved_while_the_knob_is_held_leaves_the_level_in_hand() {
    let _room = the_room_to_itself();

    let mut levels = PinLevels::set(10, 50);
    levels.remembering(false, false);
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(video);
    }

    let mut pin = overlay_pin((100, 100, 420, 340), PinChrome::always());
    pin.path = PathBuf::from("tray-held-volume-film.mkv");
    pin.transport.begun(30.0, true, false);
    pin.transport.duration = Some(200.0);
    pin.volume.level = 37;
    pin.volume.playing_at = 37;
    pin.volume.open = true;
    pin.volume.dragging = true;
    stand_pin(Some(pin));

    levels.set_levels(80, 50);
    tray_volume_moved(false);

    assert_eq!(
        pinned_volume_level(),
        37,
        "a knob being dragged is the level in hand, not the row's"
    );
    assert_eq!(
        pinned_playback_state().map(|(_, _, _, volume)| volume.playing_at),
        Some(37),
        "and the level the player was told is the one the drag took it to"
    );
    assert!(
        pin_volume_open(),
        "the popup in hand is left open where the drag put it"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// The media behind the pin, as the machine's own engine rather than a player of its own: the
/// native arms of `set_pin_volume` and `settle_pin_volume` are read off the media's kind, so this
/// is what a test about the other routes pins them with.
fn stand_native_video_media() {
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        if let Some(media) = media.as_mut() {
            media.media_type = MediaType::NativeVideo;
        }
    }
}

/// A sound's filter is the file's own measured gain and **never the app's level**, which is the
/// session's (see `audio_session::ask_the_mixer`): a file nothing has measured is played as the file
/// holds it, and a measured one is normalised at every level including nothing.
#[test]
fn a_sounds_filter_is_the_files_gain_and_never_the_apps_level() {
    assert_eq!(audio_gain_filter(None), None);
    assert_eq!(
        audio_gain_filter(Some(1.5)).as_deref(),
        Some("volume=1.5000")
    );
    assert_eq!(
        audio_gain_filter(Some(1.0)).as_deref(),
        Some("volume=1.0000"),
        "a gain of one is a gain, not the absence of one"
    );
}

/// The box on screen a wheel notch over the volume control belongs to: the transport strip's
/// button while the popup is closed, that button and the panel above it once it is open, and a
/// sound's card button where the pin has no strip at all (see `crate::app::seam::PIN_VOLUME_REGION`).
#[test]
fn a_wheel_notch_over_the_volume_control_is_answered_by_the_box_that_control_is_drawn_in() {
    let _room = the_room_to_itself();
    let previous = take_pin_for_a_test();

    let content = (100, 100, 500, 400);
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.transport_bar = true;
    let window = pin.window_box();
    let button = pinned_volume_button(&pin).expect("a strip carries a volume button");
    stand_pin(Some(pin));

    assert_eq!(
        crate::app::seam::PIN_VOLUME_REGION.get(),
        Some((
            window.0 + button.left,
            window.1 + button.top,
            window.0 + button.right,
            window.1 + button.bottom,
        )),
        "a closed popup leaves the notch the button's own box"
    );

    // And opening the panel widens the answer to take it in: the level is turned on the panel
    // above the button without the hand having to find the button again.
    let mut pin = take_pin_for_a_test().expect("the pin the test stands");
    pin.volume.open = true;
    let popup = pinned_volume_panel(&pin).expect("a panel to hang off the button");
    stand_pin(Some(pin));

    assert_eq!(
        crate::app::seam::PIN_VOLUME_REGION.get(),
        Some((
            window.0 + button.left.min(popup.panel.left),
            window.1 + button.top.min(popup.panel.top),
            window.0 + button.right.max(popup.panel.right),
            window.1 + button.bottom.max(popup.panel.bottom),
        )),
        "an open popup is the panel and the button it hangs from"
    );

    // A sound has no strip, and its card's own button is the control the same panel hangs off.
    let sound = PinnedPreview {
        content,
        zoom: Some(PinZoom::filling(content)),
        frame: PinFrame::None,
        overlay: false,
        hides_chrome: false,
        caption: 0,
        ..PinnedPreview::for_test()
    };
    let sound_window = sound.window_box();
    let sound_button = pinned_volume_button(&sound).expect("a card carries a volume button");
    stand_pin(Some(sound));

    assert_eq!(
        crate::app::seam::PIN_VOLUME_REGION.get(),
        Some((
            sound_window.0 + sound_button.left,
            sound_window.1 + sound_button.top,
            sound_window.0 + sound_button.right,
            sound_window.1 + sound_button.bottom,
        )),
        "a sound's card button is the control with no transport strip to carry one"
    );

    stand_pin(None);
    assert_eq!(
        crate::app::seam::PIN_VOLUME_REGION.get(),
        None,
        "and no pin is no box"
    );
    stand_pin(previous);
}

/// A notch over the volume control is the level's before it is the pin's zoom: the button sits
/// inside the box the zoom road claims whole, and the panel can float over a page that scrolls,
/// so the hook has to ask the volume first and leave neither count behind (see
/// `wheel_hook_proc`).
#[test]
fn a_wheel_notch_over_the_volume_control_is_the_levels_and_not_the_pins_zoom() {
    let _room = the_room_to_itself();
    let previous = take_pin_for_a_test();

    let content = (100, 100, 500, 400);
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.transport_bar = true;
    pin.volume.open = true;
    let button = pinned_volume_button(&pin).expect("a strip carries a volume button");
    let window = pin.window_box();
    stand_pin(Some(pin));

    crate::app::seam::publish_pin_always_on_top(true);
    crate::app::seam::publish_preview_under_the_pointer(false);
    let _ = crate::app::seam::take_pin_volume_delta();
    let _ = crate::app::seam::take_pin_zoom_delta();

    let info = MSLLHOOKSTRUCT {
        pt: POINT {
            x: window.0 + (button.left + button.right) / 2,
            y: window.1 + (button.top + button.bottom) / 2,
        },
        mouseData: (WHEEL_DELTA as u32) << 16,
        flags: 0,
        time: 0,
        dwExtraInfo: 0,
    };
    let swallowed = unsafe {
        crate::shell::wheel_input::wheel_hook_proc(
            HC_ACTION as i32,
            WPARAM(WM_MOUSEWHEEL as usize),
            LPARAM(&info as *const MSLLHOOKSTRUCT as isize),
        )
    };

    assert_eq!(
        crate::app::seam::take_pin_volume_delta(),
        WHEEL_DELTA,
        "the notch over the button is the level being turned"
    );
    assert_eq!(
        crate::app::seam::take_pin_zoom_delta(),
        0,
        "and it is not also the picture being zoomed"
    );
    assert_eq!(
        swallowed,
        LRESULT(1),
        "and the notch is not left for the listing behind the pin"
    );

    crate::app::seam::publish_pin_always_on_top(false);
    stand_pin(None);
    stand_pin(previous);
}

/// A wheel turned over the sound card's own volume button opens the panel, and the panel does not
/// close on the next tick while the pointer is still on the button that opened it (see
/// `refresh_pin_volume`).
#[test]
fn a_volume_popup_opened_over_a_sounds_button_is_kept_while_the_pointer_is_on_it() {
    let content = (100, 100, 500, 400);
    let mut sound = PinnedPreview {
        content,
        zoom: Some(PinZoom::filling(content)),
        frame: PinFrame::None,
        overlay: false,
        hides_chrome: false,
        caption: 0,
        ..PinnedPreview::for_test()
    };
    sound.volume.open = true;
    let button = pinned_volume_button(&sound).expect("a card carries a volume button");
    let window = sound.window_box();
    let cursor = Some((
        window.0 + (button.left + button.right) / 2,
        window.1 + (button.top + button.bottom) / 2,
    ));

    assert!(
        !refresh_pin_volume(&mut sound, cursor),
        "the pointer on the button that opened the panel does not put it away"
    );
    assert!(sound.volume.open, "and the panel the notch opened stays up");
}

/// A notch moves the level by five percent of the range, a smooth wheel's smaller messages add up
/// to the same, and both stop at the ends `set_pin_volume` draws (see `volume_after_wheel_delta`).
#[test]
fn a_wheel_notch_steps_the_level_by_five_percent_and_stops_at_the_ends() {
    let notch = |level: u32, notches: i64| {
        let units = (0..notches.abs()).fold(
            i64::from(level) * WHEEL_LEVEL_UNITS_PER_PERCENT,
            |units, _| volume_after_wheel_delta(units, WHEEL_DELTA * notches.signum() as i32),
        );
        volume_from_wheel_units(units)
    };
    assert_eq!(notch(50, 1), 55, "up is louder");
    assert_eq!(notch(50, -1), 45, "and down is quieter");
    assert_eq!(notch(98, 1), 100, "the top holds");
    assert_eq!(notch(2, -1), 0, "and so does nothing");
    assert_eq!(notch(50, -100), 0, "however far the wheel turns");

    // Twelve 10-unit messages are one notch, exactly: a smooth wheel's travel is not lost to a
    // whole-notch carry, and does not overshoot the +5% a notched wheel lands on.
    let mut units = 50 * WHEEL_LEVEL_UNITS_PER_PERCENT;
    for _ in 0..(WHEEL_DELTA / 10) {
        units = volume_after_wheel_delta(units, 10);
    }
    assert_eq!(
        units,
        50 * WHEEL_LEVEL_UNITS_PER_PERCENT + WHEEL_DELTA as i64,
        "twelve sub-notch messages carry exactly one notch, with nothing dropped"
    );
    assert_eq!(
        volume_from_wheel_units(units),
        55,
        "and exactly one notch on"
    );
    assert_eq!(
        volume_from_wheel_units(volume_after_wheel_delta(
            100 * WHEEL_LEVEL_UNITS_PER_PERCENT - 1,
            10
        )),
        100,
        "the top clamps the fine carry too"
    );
    assert_eq!(
        volume_from_wheel_units(volume_after_wheel_delta(1, -10)),
        0,
        "and so does nothing"
    );
}

/// A notch arrives on a carry built from a level a drag, a tray row or a key has since moved: the
/// notch is added to the level in hand rather than to that older position, which is what used to
/// step the thumb back to where the hand had been. A carry the wheel is still gliding toward is
/// kept whole instead, so sub-notch travel survives a ramp (see `wheel_base_units`).
#[test]
fn a_wheel_notch_starts_from_the_level_in_hand_unless_a_glide_is_still_running() {
    let per = WHEEL_LEVEL_UNITS_PER_PERCENT;
    let fresh = 60 * per;

    // A drag raised the level under a carry left three notches below it.
    assert_eq!(wheel_base_units(Some(30 * per), 60, false), fresh);
    // And lowered it under a carry above it.
    assert_eq!(wheel_base_units(Some(90 * per), 60, false), fresh);
    // A carry built from the level in hand is that carry, sub-notch travel and all.
    let travelling = fresh + 5;
    assert_eq!(volume_from_wheel_units(travelling), 60);
    assert_eq!(wheel_base_units(Some(travelling), 60, false), travelling);
    // A glide in flight is the wheel's own level still on its way, so its carry outlives a level
    // that would otherwise re-base it.
    assert_eq!(wheel_base_units(Some(30 * per), 60, true), 30 * per);
    // Nothing carried is the level itself.
    assert_eq!(wheel_base_units(None, 60, false), fresh);
    assert_eq!(wheel_base_units(None, 60, true), fresh);
}

/// A drag or a wheel burst asks for a level on every message, but the engine hears it once per
/// preview tick, and what it hears is the level the hand stopped on — never a queue of stale ones
/// (see `flush_pin_volume_level`).
#[test]
fn a_level_burst_reaches_mpv_once_per_tick_with_the_latest_level() {
    let _room = the_room_to_itself();

    let mut levels = PinLevels::set(100, 50);
    levels.remembering(false, false);
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(video);
    }

    let content = (100, 100, 420, 340);
    let film = PathBuf::from("pin-coalesced-volume-film.mkv");
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = film;
    pin.transport.begun(30.0, true, false);
    pin.transport.duration = Some(200.0);
    pin.volume.level = 40;
    pin.volume.playing_at = 40;
    stand_pin(Some(pin));

    stand_ffplay_in(Some(false));
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    mpv_ipc::remember_pipe(101, "coalesced-volume-pipe", false);
    clear_restart_count();
    mpv_ipc::clear_noted();

    // Three steps of one drag, one tick.
    set_pin_volume(41);
    set_pin_volume(55);
    set_pin_volume(70);
    assert!(
        mpv_ipc::noted_lines().is_empty(),
        "nothing is written while the hand is still moving"
    );

    flush_pin_volume_level();
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![format!(
            "{}\n",
            r#"{"command":["set_property","volume",70],"async":true}"#
        )],
        "one command, and it is the latest level"
    );
    assert_eq!(
        pinned_playback_state().map(|(_, _, _, volume)| volume.playing_at),
        Some(70),
        "the claim follows the one level that was actually told"
    );

    // A second flush in the same tick sends nothing: the burst is spent.
    flush_pin_volume_level();
    assert_eq!(mpv_ipc::noted_lines().len(), 1, "no second command");

    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
}

/// **A wheel burst and a release begin no player at all**: FFmpeg's level is its mixer session's, so
/// the quiet window is the claims' settling and nothing is replaced at either end of it (see
/// `settle_pin_volume_after_quiet` and `settle_pin_volume`).
#[test]
fn a_wheel_burst_and_a_release_begin_no_player() {
    let _room = the_room_to_itself();

    let mut levels = PinLevels::set(100, 50);
    levels.remembering(false, false);
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(video);
    }

    let content = (100, 100, 420, 340);
    let film = PathBuf::from("pin-quiet-volume-film.mkv");
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = film;
    pin.transport.begun(30.0, true, false);
    pin.transport.duration = Some(200.0);
    pin.volume.level = 40;
    pin.volume.playing_at = 40;
    stand_pin(Some(pin));

    stand_ffplay_in(Some(true));
    stand_mpv_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    clear_restart_count();

    // A normal wheel's notches arrive with gaps between ticks; every one resets the countdown.
    for level in [45, 50, 55, 60, 65] {
        set_pin_volume(level);
        settle_pin_volume_after_quiet();
        assert_eq!(
            restart_count(),
            0,
            "the player is not replaced while the wheel is still turning"
        );
    }
    assert_eq!(restart_count(), 0, "a burst of notches begins nothing");

    // Then the hand stops, the quiet window passes, and the level is told at the mixer.
    for _ in 0..PIN_VOLUME_SETTLE_TICKS {
        settle_pin_volume_after_quiet();
    }
    assert_eq!(
        restart_count(),
        0,
        "the level is the session's, so nothing is begun for it"
    );

    // A release settles at once, however recently the knob moved, and begins nothing either.
    set_pin_volume(80);
    settle_pin_volume();
    assert_eq!(
        restart_count(),
        0,
        "a release is the end of the drag and begins no player"
    );

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
}

/// The box a wheel notch over the volume control is answered by is published, not resolved: the
/// wheel hook asks it while `PIN_STATE` is held by the preview thread, which is half of the notches
/// of a turned wheel falling through to Explorer (see `pin_volume_keep_alive_try` and
/// `PIN_VOLUME_REGION`).
#[test]
fn a_wheel_notch_is_answered_by_the_volume_box_while_the_pin_state_is_held() {
    let _room = the_room_to_itself();
    let previous = take_pin_for_a_test();

    let content = (100, 100, 500, 400);
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.transport_bar = true;
    let button = pinned_volume_button(&pin).expect("a strip carries a volume button");
    let window = pin.window_box();
    stand_pin(Some(pin));

    crate::app::seam::publish_pin_always_on_top(true);
    crate::app::seam::publish_preview_under_the_pointer(false);
    let _ = crate::app::seam::take_pin_volume_delta();

    // The lock the hook used to wait on, held for the whole of the question.
    let held = pin_state().expect("the pin state the test published");

    assert_eq!(
        crate::app::seam::PIN_VOLUME_REGION.get(),
        Some((
            window.0 + button.left,
            window.1 + button.top,
            window.0 + button.right,
            window.1 + button.bottom,
        )),
        "the box is read without the lock the preview thread is holding"
    );

    let info = MSLLHOOKSTRUCT {
        pt: POINT {
            x: window.0 + (button.left + button.right) / 2,
            y: window.1 + (button.top + button.bottom) / 2,
        },
        mouseData: (WHEEL_DELTA as u32) << 16,
        flags: 0,
        time: 0,
        dwExtraInfo: 0,
    };
    let swallowed = unsafe {
        crate::shell::wheel_input::wheel_hook_proc(
            HC_ACTION as i32,
            WPARAM(WM_MOUSEWHEEL as usize),
            LPARAM(&info as *const MSLLHOOKSTRUCT as isize),
        )
    };
    assert_eq!(
        crate::app::seam::take_pin_volume_delta(),
        WHEEL_DELTA,
        "the notch is the level's even with the pin held by another thread"
    );
    assert_eq!(swallowed, LRESULT(1), "and it is not left for Explorer");
    drop(held);

    crate::app::seam::publish_pin_always_on_top(false);
    stand_pin(None);
    stand_pin(previous);
}

/// The level a wheel asks for is a glide: it climbs toward the wheel's fine target a step a tick,
/// never past it and never backwards, and lands on the exact level the notch names (see
/// `ramp_pin_volume`).
#[test]
fn a_wheel_notchs_level_glides_to_its_target_and_lands_on_it_exactly() {
    let _room = the_room_to_itself();
    let previous = take_pin_for_a_test();

    let content = (100, 100, 500, 400);
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.transport_bar = true;
    pin.volume = PinVolume {
        level: 50,
        playing_at: 50,
        ..Default::default()
    };
    stand_pin(Some(pin));

    let per = WHEEL_LEVEL_UNITS_PER_PERCENT;
    let target = volume_after_wheel_delta(50 * per, WHEEL_DELTA);
    set_pin_volume_target(target);

    let held = || {
        pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| pin.volume))
            .expect("the pin the test stands")
    };

    assert_eq!(
        held().level,
        50,
        "the popup is still drawn where the level is"
    );
    assert_eq!(
        held().target,
        Some(target as u32),
        "and the notch is owed a glide"
    );

    let mut applied = 0;
    let mut level = 0;
    let mut ticks = 0;
    loop {
        ramp_pin_volume();
        ticks += 1;
        let volume = held();
        assert!(
            volume.applied >= applied,
            "the glide only ever climbs toward the target"
        );
        assert!(
            i64::from(volume.applied) <= target,
            "and it never goes past it"
        );
        assert!(volume.level >= level, "and the drawn level with it");
        applied = volume.applied;
        level = volume.level;

        if volume.target.is_none() {
            assert_eq!(
                i64::from(volume.applied),
                target,
                "the glide stops exactly on the target"
            );
            break;
        }
        assert!(
            ticks < 60,
            "a five percent notch is a handful of ticks, not a second of them"
        );
    }

    // One 120-unit notch is five percent, exactly: 50 -> 55, not 54 or 56.
    assert_eq!(
        i64::from(held().level) * per,
        target,
        "the notch settles on its own fine level"
    );
    assert_eq!(held().level, 55, "which is five percent on");

    stand_pin(previous);
}

/// A knob dragged on the popup, a tray row and a key are levels in hand rather than requests: they
/// are applied at once, and a glide a wheel left behind is dropped rather than carried on with (see
/// `set_pin_volume` and `ramp_pin_volume`).
#[test]
fn a_dragged_knob_moves_the_level_at_once_and_drops_a_wheels_glide() {
    let _room = the_room_to_itself();
    let previous = take_pin_for_a_test();

    let content = (100, 100, 500, 400);
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.transport_bar = true;
    pin.volume = PinVolume {
        level: 50,
        playing_at: 50,
        ..Default::default()
    };
    stand_pin(Some(pin));

    let held = || {
        pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| pin.volume))
            .expect("the pin the test stands")
    };

    set_pin_volume(64);
    assert_eq!(held().level, 64, "the drag is heard at once");
    assert_eq!(held().target, None, "with nothing left to glide toward");

    // A wheel request mid-drag is replaced by the hand that came after it.
    let per = WHEEL_LEVEL_UNITS_PER_PERCENT;
    set_pin_volume_target(volume_after_wheel_delta(64 * per, WHEEL_DELTA));
    assert_eq!(held().level, 64, "the request alone moves no level");
    set_pin_volume(20);
    assert_eq!(held().level, 20);
    assert_eq!(held().target, None);
    assert_eq!(
        i64::from(held().applied),
        20 * per,
        "the glide starts from the level in hand"
    );
    ramp_pin_volume();
    assert_eq!(held().level, 20, "and there is nothing to glide toward");

    stand_pin(previous);
}

/// Every notch of a burst moves the wheel's target five percent on, and the applied level stays
/// behind it until the ticks catch up (see `set_pin_volume_target` and `ramp_pin_volume`).
#[test]
fn a_burst_of_wheel_notches_moves_the_target_five_percent_a_notch() {
    let _room = the_room_to_itself();
    let previous = take_pin_for_a_test();

    let content = (100, 100, 500, 400);
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.transport_bar = true;
    pin.volume = PinVolume {
        level: 50,
        playing_at: 50,
        ..Default::default()
    };
    stand_pin(Some(pin));

    let per = WHEEL_LEVEL_UNITS_PER_PERCENT;
    let mut units = 50 * per;
    for notch in 1..=3 {
        units = volume_after_wheel_delta(units, WHEEL_DELTA);
        set_pin_volume_target(units);
        let volume = pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| pin.volume))
            .expect("the pin the test stands");
        assert_eq!(
            volume.target,
            Some(units as u32),
            "notch {notch} is another five percent of the range"
        );
        assert_eq!(volume.level, 50, "and the applied level has not jumped");
    }
    assert_eq!(volume_from_wheel_units(units), 65, "three notches up");

    stand_pin(previous);
}

/// A notch is chrome on every kind, and a film's thumb is drawn in the strip a repaint reaches:
/// the flag a notch sets is drained once a tick, so the thumb steps at the cadence the wheel holds
/// the loop at rather than on whatever incidental repaint comes next — and a sound's card, whose
/// paint is the card rather than chrome, is left to its own arm (see `pin_chrome_owes_a_repaint`).
#[test]
fn a_wheel_notch_on_a_film_owes_one_chrome_repaint_a_tick() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(video);
    }

    let content = (100, 100, 500, 400);
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.transport_bar = true;
    pin.volume = PinVolume {
        level: 50,
        playing_at: 50,
        ..Default::default()
    };
    stand_pin(Some(pin));
    AUDIO_CARD_DIRTY.store(false, Ordering::Release);

    // Two notches inside the one tick: the level moves twice, and the tick still owes one paint.
    let per = WHEEL_LEVEL_UNITS_PER_PERCENT;
    set_pin_volume_target(50 * per);
    set_pin_volume_target(55 * per);
    assert!(
        pin_chrome_owes_a_repaint(),
        "a notch on a film asks for the chrome the thumb is drawn in"
    );
    assert!(
        !pin_chrome_owes_a_repaint(),
        "and the notch after it is the same repaint, not another"
    );

    // A sound's card is refreshed by its own arm before this one is asked: the drain here leaves
    // its flag standing rather than repainting a card the window's paint does not rebuild.
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        if let Some(media) = media.as_mut() {
            media.media_type = MediaType::Audio;
        }
    }
    AUDIO_CARD_DIRTY.store(true, Ordering::Release);
    assert!(
        !pin_chrome_owes_a_repaint(),
        "a sound's card is not owed the film's chrome repaint"
    );
    assert!(
        AUDIO_CARD_DIRTY.swap(false, Ordering::AcqRel),
        "and the flag is still the card's own to spend"
    );

    AUDIO_CARD_DIRTY.store(false, Ordering::Release);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// **A dragged knob on a pin FFmpeg plays is heard on every tick of the drag**, a film's and a
/// sound's alike: the level is owed the mixer session while the hand is still moving rather than
/// only from the release, so `PinVolume::pending` is set by each step, `flush_pin_volume_level`
/// spends it, and the claim follows it (see `pin_level_is_immediate` and `ffplay_holds_the_level`).
///
/// The player is a noted one that never ends, which is what `is_video_process_running` needs of it,
/// and nothing is begun at either end of the drag: the level is a session's.
#[test]
fn a_dragged_level_on_a_pin_ffmpeg_plays_reaches_it_tick_by_tick() {
    let _room = the_room_to_itself();

    let mut levels = PinLevels::set(100, 50);
    levels.remembering(false, false);
    let previous_image = video_process_image();
    let previous_pid = VIDEO_PID.load(Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    video.video_process = Some(video_playback::PlayerProcess::Noted(101));
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(video);
    }
    note_player_process(101, FFPLAY_IMAGE_NAME);
    stand_ffplay_in(Some(true));
    stand_mpv_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    clear_restart_count();

    let held = || pin_holding().expect("the pin the test stands");
    let mut pin = overlay_pin((100, 100, 420, 340), PinChrome::always());
    pin.path = PathBuf::from("pin-ffplay-dragged-level-film.mkv");
    pin.transport.begun(30.0, true, false);
    pin.transport.duration = Some(200.0);
    pin.volume.level = 40;
    pin.volume.playing_at = 40;
    stand_pin(Some(pin));

    assert!(
        ffplay_holds_the_level(),
        "the player on screen is FFmpeg's and still holds a process"
    );

    // Three steps of one drag, each with its own tick between them.
    for level in [45, 55, 70] {
        set_pin_volume(level);
        assert_eq!(held().level, level, "the step {level} is heard at once");
        assert!(
            held().pending,
            "and the tick is owed it, rather than the release being"
        );
        flush_pin_volume_level();
        assert!(!held().pending, "the tick spends what it was owed");
        assert_eq!(
            held().playing_at,
            level,
            "the claim follows the level the mixer was asked for"
        );
    }

    // A sound's card is the other FFmpeg route, and the level is its session's just the same.
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        if let Some(media) = media.as_mut() {
            media.media_type = MediaType::Audio;
        }
    }
    let mut sound = sound_pin();
    sound.path = PathBuf::from("pin-ffplay-dragged-level-sound.mp3");
    sound.volume = PinVolume {
        audio: true,
        level: 40,
        playing_at: 40,
        ..Default::default()
    };
    stand_pin(Some(sound));
    for level in [50, 60] {
        set_pin_volume(level);
        assert!(held().pending, "a sound's card is owed its level live too");
        flush_pin_volume_level();
        assert_eq!(held().playing_at, level, "and it is told on the same tick");
    }

    assert_eq!(
        restart_count(),
        0,
        "the level is a session's, so the drag begins, ends and restarts nothing"
    );
    settle_pin_volume();
    assert_eq!(
        restart_count(),
        0,
        "and a release the claim already agrees with begins nothing either"
    );

    note_player_process(previous_pid, previous_image);
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    clear_restart_count();
}

/// A drag measures the hand against the panel it began on, whatever the pin's chrome or card does
/// under it: the box is frozen at the press, so the groove and the thumb drawn in it stay where the
/// pointer found them, and the freeze goes with the release (see `PinVolume::held`).
#[test]
fn a_drag_keeps_the_popup_it_began_on_while_the_pins_chrome_relays_out() {
    let _room = the_room_to_itself();
    let previous = take_pin_for_a_test();

    let mut pin = overlay_pin((100, 100, 500, 400), PinChrome::always());
    pin.transport_bar = true;
    pin.volume = PinVolume {
        level: 50,
        playing_at: 50,
        open: true,
        ..Default::default()
    };
    let frozen = pinned_volume_panel(&pin).expect("a strip carries a volume panel");
    stand_pin(Some(pin));

    // The press takes hold of this box and the strip it hung from goes away under the hand.
    with_pin(|pin| {
        pin.volume.dragging = true;
        pin.volume.held = Some(frozen);
        pin.chrome = PinChrome {
            caption: false,
            bar: false,
            until: None,
        };
        pin.transport_bar = false;
    });
    let dragging_panel = pin_state()
        .and_then(|pinned| pinned.pin().map(pinned_volume_panel))
        .flatten()
        .expect("a held drag always answers a panel");
    assert_eq!(
        dragging_panel, frozen,
        "the panel and the groove the hand is measured against are the press's own"
    );

    // What the layout would answer unfrozen is another box altogether, which is the jump this
    // freeze is for.
    with_pin(|pin| pin.volume.dragging = false);
    let unfrozen = pin_state()
        .and_then(|pinned| pinned.pin().map(pinned_volume_panel))
        .flatten();
    assert_ne!(
        unfrozen,
        Some(frozen),
        "the chrome's own layout has moved the panel out from under the hand"
    );

    // Letting go of the knob drops the freeze with it, so the next press takes the box it finds.
    assert!(close_pin_volume(), "the popup was open");
    let after = pin_holding().expect("the pin the test stands");
    assert!(!after.dragging, "and the knob is no longer held");
    assert_eq!(after.held, None, "the frozen box goes with the drag");

    stand_pin(previous);
}

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

/// A level moved on a pinned film: mpv is told it while it runs, and FFmpeg's player is begun
/// again at it — the gate this ticket turns on, asserted on both routes.
///
/// FFmpeg's player takes a level only by being started at one, so the settle is a player replaced;
/// mpv takes one as it plays, so the knob itself is the write and the settle begins nothing (see
/// `set_pin_volume` and `settle_pin_volume`). The `Remember` rows are stood off so that a test
/// never writes the configuration file a level would be remembered into.
#[test]
fn a_level_moved_on_a_pinned_mpv_is_told_to_the_player_while_ffplays_is_begun_again() {
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

    // FFmpeg's player takes a level only by being begun again at one.
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
        1,
        "a level moved on FFmpeg's player is a player begun again at the level, as it always was"
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

    // FFmpeg's player, which is told nothing while it runs: the film the tray has turned down to
    // nothing is begun again at nothing, from the second it was watching.
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
        1,
        "FFmpeg's player takes a level by being begun at one, so the row is a player begun again"
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

/// A sound's level reaches the player's filter at every level including the top: the player's own
/// `-volume` knob is silent at 100, so a sound at `Volume → Audio` 100% goes through `volume=1.00`
/// like a film's — and a level of nothing is still a player playing silence, not no player at all
/// (see `audio_volume_filter`).
#[test]
fn a_sounds_level_reaches_the_filter_at_every_level_including_the_top() {
    assert_eq!(audio_volume_filter(100, None), "volume=1.00");
    assert_eq!(audio_volume_filter(99, None), "volume=0.99");
    assert_eq!(
        audio_volume_filter(0, None),
        "volume=0.00",
        "silence that still plays, so the card's clock moves and the file stays seekable"
    );
    assert_eq!(audio_volume_filter(100, Some(1.5)), "volume=1.5000");
}

/// The box on screen a wheel notch over the volume control belongs to: the transport strip's
/// button while the popup is closed, that button and the panel above it once it is open, and a
/// sound's card button where the pin has no strip at all (see `pin_volume_keep_alive_try`).
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
        pin_volume_keep_alive_try(),
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
        pin_volume_keep_alive_try(),
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
        pin_volume_keep_alive_try(),
        Some((
            sound_window.0 + sound_button.left,
            sound_window.1 + sound_button.top,
            sound_window.0 + sound_button.right,
            sound_window.1 + sound_button.bottom,
        )),
        "a sound's card button is the control with no transport strip to carry one"
    );

    stand_pin(None);
    assert_eq!(pin_volume_keep_alive_try(), None, "and no pin is no box");
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

    crate::shell::key_input::publish_pin_always_on_top(true);
    crate::shell::wheel_input::publish_preview_under_the_pointer(false);
    let _ = wheel_input::take_pin_volume_delta();
    let _ = wheel_input::take_pin_zoom_delta();

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
        wheel_input::wheel_hook_proc(
            HC_ACTION as i32,
            WPARAM(WM_MOUSEWHEEL as usize),
            LPARAM(&info as *const MSLLHOOKSTRUCT as isize),
        )
    };

    assert_eq!(
        wheel_input::take_pin_volume_delta(),
        WHEEL_DELTA,
        "the notch over the button is the level being turned"
    );
    assert_eq!(
        wheel_input::take_pin_zoom_delta(),
        0,
        "and it is not also the picture being zoomed"
    );
    assert_eq!(
        swallowed,
        LRESULT(1),
        "and the notch is not left for the listing behind the pin"
    );

    crate::shell::key_input::publish_pin_always_on_top(false);
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

/// FFmpeg's player takes a level only by being begun again, so a wheel burst must not begin one per
/// notch: the relaunch waits for `PIN_VOLUME_SETTLE_TICKS` quiet ticks and then happens once, while
/// a drag released still settles at once (see `settle_pin_volume_after_quiet` and
/// `settle_pin_volume`).
#[test]
fn a_wheel_burst_relaunches_ffplay_only_after_the_quiet_window() {
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

    // Then the hand stops, the quiet window passes, and one player takes the level.
    for _ in 0..PIN_VOLUME_SETTLE_TICKS {
        settle_pin_volume_after_quiet();
    }
    assert_eq!(
        restart_count(),
        1,
        "one relaunch once the level has been quiet"
    );

    // And the ticks after it begin nothing more.
    for _ in 0..PIN_VOLUME_SETTLE_TICKS {
        settle_pin_volume_after_quiet();
    }
    assert_eq!(
        restart_count(),
        1,
        "the quiet window is spent, not repeated"
    );

    // A release still settles at once, however recently the knob moved.
    set_pin_volume(80);
    settle_pin_volume();
    assert_eq!(
        restart_count(),
        2,
        "a release is the end of the drag and settles immediately"
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

use super::*;

/// The volume popup belongs to the button that opened it and to nothing else: it is kept while
/// the pointer is anywhere in that control — the panel itself, or the strip the button sits in
/// — held while the knob is being dragged, and put away once the pointer has left the whole of
/// it, which is a repaint the window owes.
#[test]
fn a_volume_popup_is_kept_on_its_control_and_put_away_once_the_pointer_has_left_it() {
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
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

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
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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

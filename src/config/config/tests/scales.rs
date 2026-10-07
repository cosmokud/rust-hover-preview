use super::*;

#[test]
fn avoid_mode_reads_back_what_it_writes() {
    for mode in [
        AvoidMode::Off,
        AvoidMode::Filename,
        AvoidMode::FilenameColumn,
        AvoidMode::Details,
    ] {
        let written = mode.as_str();
        assert_eq!(
            AvoidMode::from_str(written),
            Some(mode),
            "`{written}` read back"
        );
    }

    assert_eq!(AvoidMode::from_str("  DETAILS "), Some(AvoidMode::Details));
    assert_eq!(
        AvoidMode::from_str("every column"),
        None,
        "a value that names no way of avoiding is not one"
    );
}

/// The name-only way of avoiding is what the app does unless it is told otherwise,
/// so a configuration that says nothing about it keeps a preview off the name.
#[test]
fn a_fresh_configuration_avoids_the_name_alone() {
    assert_eq!(AppConfig::default().avoid_mode, AvoidMode::Filename);
}

#[test]
fn office_engine_idle_reads_back_what_it_writes() {
    for idle in [
        EngineIdle::Seconds(0),
        EngineIdle::Seconds(600),
        EngineIdle::Seconds(MAX_OFFICE_ENGINE_IDLE_SECS),
        EngineIdle::Indefinite,
    ] {
        let written = idle.as_str();
        assert_eq!(
            EngineIdle::from_str(&written),
            Some(idle),
            "`{written}` read back"
        );
    }
}

#[test]
fn office_engine_idle_takes_the_words_a_person_would_write() {
    for written in ["indefinitely", "Indefinite", " forever ", "ALWAYS"] {
        assert_eq!(
            EngineIdle::from_str(written),
            Some(EngineIdle::Indefinite),
            "`{written}`"
        );
    }

    assert_eq!(
        EngineIdle::from_str(" 900 "),
        Some(EngineIdle::Seconds(900))
    );
    assert_eq!(
        EngineIdle::from_str("999999"),
        Some(EngineIdle::Seconds(MAX_OFFICE_ENGINE_IDLE_SECS)),
        "a number past the ceiling is reduced to it"
    );
    assert_eq!(
        EngineIdle::from_str("soon"),
        None,
        "a value that is neither a time nor a word is not one"
    );
}

#[test]
fn office_engine_idle_expires_on_its_own_clock() {
    let minute = Duration::from_secs(60);

    assert!(
        EngineIdle::Seconds(0).has_expired(Duration::ZERO),
        "an engine let go as soon as it has drawn a page is idle at once"
    );
    assert!(!EngineIdle::Seconds(60).has_expired(minute - Duration::from_secs(1)));
    assert!(EngineIdle::Seconds(60).has_expired(minute));
    assert!(EngineIdle::Seconds(60).has_expired(Duration::from_secs(9_999)));
    assert!(
        !EngineIdle::Indefinite.has_expired(Duration::from_secs(365 * 24 * 60 * 60)),
        "an engine kept for the life of the app never goes idle"
    );
}

/// A drawing is drawn at the whole room the display has unless the file says
/// otherwise — which is what the setting starts at, and what a fresh install
/// writes.
#[test]
fn a_drawings_scale_starts_at_the_whole_room() {
    let config = AppConfig::default();

    assert_eq!(config.vector_scale, PreviewScale::FitToScreen);
    assert_eq!(config.vector_scale.as_str(), "fit");
}

/// A page's scale is a setting of its own as well: a PDF, a document a page was drawn
/// for and a hand-edited picture scale are three answers to three questions, and one
/// of them changing leaves the others where they were.
#[test]
fn a_pages_scale_is_read_from_its_own_key() {
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "preview_scale", Some("400".to_string()));
    ini.set(CONFIG_SECTION, "vector_scale", Some("75".to_string()));
    ini.set(CONFIG_SECTION, "ebook_scale", Some("25".to_string()));
    ini.set(CONFIG_SECTION, "document_scale", Some("10".to_string()));

    let config = read_file(&mut ini);

    assert_eq!(config.ebook_scale, PreviewScale::Percent(25));
    assert_eq!(config.document_scale, PreviewScale::Percent(10));
    assert_eq!(config.vector_scale, PreviewScale::Percent(75));
    assert_eq!(config.preview_scale, PreviewScale::Percent(400));

    // The words a person would write are read for a page key the same way they are
    // for a document's, since it is the same value read against the same whole.
    let mut ini = Ini::new();
    ini.set(
        CONFIG_SECTION,
        "ebook_scale",
        Some(" Fit to Screen ".to_string()),
    );
    ini.set(CONFIG_SECTION, "document_scale", Some("50%".to_string()));

    let config = read_file(&mut ini);

    assert_eq!(config.ebook_scale, PreviewScale::FitToScreen);
    assert_eq!(config.document_scale, PreviewScale::Percent(50));
}

/// The same for a video: a video's scale is a setting of its own like the four
/// document scales beside it, so one key changing leaves the others where they were.
/// A picture and a video are drawn at the same share by default, which the two keys
/// keep apart rather than sharing.
#[test]
fn a_videos_scale_is_read_from_its_own_key() {
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "preview_scale", Some("400".to_string()));
    ini.set(CONFIG_SECTION, "video_scale", Some("50".to_string()));

    let config = read_file(&mut ini);

    assert_eq!(config.video_scale, PreviewScale::Percent(50));
    assert_eq!(config.preview_scale, PreviewScale::Percent(400));

    let mut ini = Ini::new();
    ini.set(
        CONFIG_SECTION,
        "video_scale",
        Some(" Fit to Screen ".to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(config.video_scale, PreviewScale::FitToScreen);
}

/// A video's scale is a setting of its own: a file that says nothing about one leaves it
/// where a fresh installation starts, whatever the file says its pictures are drawn at —
/// one share is not an answer to the other's question.
#[test]
fn a_file_without_a_video_scale_leaves_it_where_it_starts() {
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "preview_scale", Some("25".to_string()));

    let config = read_file(&mut ini);

    assert_eq!(config.preview_scale, PreviewScale::Percent(25));
    assert_eq!(config.video_scale, DEFAULT_VIDEO_SCALE);

    let config = AppConfig::default();

    assert_eq!(config.preview_scale, DEFAULT_PREVIEW_SCALE);
    assert_eq!(config.video_scale, DEFAULT_VIDEO_SCALE);
}

/// And the same for an animation: a setting of its own beside the video and picture
/// scales, so one key changing leaves the others where they were — and a file that has
/// no key for it leaves it where a fresh installation starts, like the video's.
#[test]
fn an_animations_scale_is_read_from_its_own_key() {
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "preview_scale", Some("400".to_string()));
    ini.set(CONFIG_SECTION, "video_scale", Some("75".to_string()));
    ini.set(CONFIG_SECTION, "animated_scale", Some("50".to_string()));

    let config = read_file(&mut ini);

    assert_eq!(config.animated_scale, PreviewScale::Percent(50));
    assert_eq!(config.video_scale, PreviewScale::Percent(75));
    assert_eq!(config.preview_scale, PreviewScale::Percent(400));

    let mut ini = Ini::new();
    ini.set(
        CONFIG_SECTION,
        "animated_scale",
        Some(" Fit to Screen ".to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(config.animated_scale, PreviewScale::FitToScreen);

    // A file with no key for it leaves the setting where a fresh installation starts,
    // whatever share the file draws its pictures at.
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "preview_scale", Some("25".to_string()));

    let config = read_file(&mut ini);

    assert_eq!(config.animated_scale, DEFAULT_ANIMATED_SCALE);

    let config = AppConfig::default();

    assert_eq!(config.animated_scale, DEFAULT_ANIMATED_SCALE);
}

/// A sound's card is a tenth of the display unless the file says otherwise:
/// the size the card was measured at across the files it was tried on, which
/// is what a fresh install writes and what the menu marks as the default.
#[test]
fn a_sounds_card_starts_at_a_tenth_of_the_display() {
    let config = AppConfig::default();

    assert_eq!(config.audio_scale, PreviewScale::Percent(10));
    assert_eq!(config.audio_scale, DEFAULT_AUDIO_SCALE);
    assert_eq!(config.audio_scale.as_str(), "10");
}

/// The card's scale is written the way a document scale is, so the words a
/// person would write by hand are the words it reads — and every share a card
/// can be given is one the setting keeps, so a choice made in the menu is
/// still the choice after a restart.
#[test]
fn a_sounds_scale_reads_back_what_it_writes() {
    for percent in [5, 10, 15, 20, 25, 100] {
        assert_eq!(
            PreviewScale::from_audio_str(&PreviewScale::Percent(percent).as_str()),
            Some(PreviewScale::Percent(percent)),
            "`{percent}%` read back"
        );
    }
}

/// A hand-edited `config.ini` is read through the sound's own bounds: the
/// setting is a share of the display, so a `0` — a share of nothing — is the
/// share a fresh installation starts at, a percentage past the whole display is
/// the whole display rather than a card wider than one, and a value that names
/// no share at all is `None`, which leaves the setting where it is.
#[test]
fn a_hand_edited_sounds_scale_is_brought_back_to_the_display() {
    assert_eq!(
        PreviewScale::from_audio_str("150"),
        Some(PreviewScale::Percent(100)),
        "a share past the whole display is the whole display"
    );
    assert_eq!(
        PreviewScale::from_audio_str("0"),
        Some(PreviewScale::Percent(10)),
        "a share of nothing is the share a fresh installation starts at"
    );
    assert_eq!(
        PreviewScale::from_audio_str("banana"),
        None,
        "a value that names no share is not one"
    );
    assert_eq!(
        PreviewScale::from_audio_str(" 25% "),
        Some(PreviewScale::Percent(25)),
        "the words a person would write are read for them"
    );
}

/// A sound's scale is a setting of its own, like the picture and video scales
/// beside it: one key changing leaves the others where they were, and a file
/// that has no key for it — one written before the kind had a scale of its
/// own — leaves it where a fresh installation starts.
#[test]
fn a_sounds_scale_is_read_from_its_own_key() {
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "preview_scale", Some("400".to_string()));
    ini.set(CONFIG_SECTION, "video_scale", Some("75".to_string()));
    ini.set(CONFIG_SECTION, "animated_scale", Some("50".to_string()));
    ini.set(CONFIG_SECTION, "audio_scale", Some("20".to_string()));

    let config = read_file(&mut ini);

    assert_eq!(config.audio_scale, PreviewScale::Percent(20));
    assert_eq!(config.video_scale, PreviewScale::Percent(75));
    assert_eq!(config.animated_scale, PreviewScale::Percent(50));
    assert_eq!(config.preview_scale, PreviewScale::Percent(400));

    // The words a person would write are read for a sound's key the same way
    // they are for a document's, since it is the same value read against the
    // same whole.
    let mut ini = Ini::new();
    ini.set(
        CONFIG_SECTION,
        "audio_scale",
        Some(" Fit to Screen ".to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(config.audio_scale, PreviewScale::FitToScreen);

    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "audio_scale", Some("150".to_string()));

    let config = read_file(&mut ini);

    assert_eq!(
        config.audio_scale,
        PreviewScale::Percent(100),
        "a hand-edited share past the whole display is the whole display"
    );

    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "audio_scale", Some("0".to_string()));

    let config = read_file(&mut ini);

    assert_eq!(
        config.audio_scale,
        PreviewScale::Percent(10),
        "a hand-edited share of nothing is the share a fresh installation starts at"
    );

    // A file with no key for it leaves the setting where a fresh installation
    // starts, whatever share the file draws its pictures at.
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "preview_scale", Some("25".to_string()));

    let config = read_file(&mut ini);

    assert_eq!(config.preview_scale, PreviewScale::Percent(25));
    assert_eq!(config.audio_scale, DEFAULT_AUDIO_SCALE);

    let config = AppConfig::default();

    assert_eq!(config.audio_scale, DEFAULT_AUDIO_SCALE);
}

/// A page is drawn at the whole room the display has unless the file says
/// otherwise — which is what the two page scales start at, and what a fresh
/// install writes.
#[test]
fn a_pages_scale_starts_at_the_whole_room() {
    let config = AppConfig::default();

    assert_eq!(config.ebook_scale, PreviewScale::FitToScreen);
    assert_eq!(config.ebook_scale.as_str(), "fit");
    assert_eq!(config.document_scale, PreviewScale::FitToScreen);
    assert_eq!(config.document_scale.as_str(), "fit");
}

/// The scale is written the way the picture scale is, so the words a person
/// would write by hand are the words it reads: `fit`, a number, either with a
/// percent sign or without — and a reduced fit of the display, which is the
/// share the fit is reduced to, written around the word for the room it is a
/// share of.
#[test]
fn a_drawing_scale_takes_the_words_a_person_would_write() {
    for (written, expected) in [
        ("fit", PreviewScale::FitToScreen),
        ("Fit to Screen", PreviewScale::FitToScreen),
        (" 75 ", PreviewScale::Percent(75)),
        ("25%", PreviewScale::Percent(25)),
        ("10", PreviewScale::Percent(10)),
        ("screen 75", PreviewScale::FitToScreenReduced(75)),
        ("Screen 50", PreviewScale::FitToScreenReduced(50)),
        ("75 of screen", PreviewScale::FitToScreenReduced(75)),
        ("screen-25", PreviewScale::FitToScreenReduced(25)),
        ("screen10", PreviewScale::FitToScreenReduced(10)),
    ] {
        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "vector_scale", Some(written.to_string()));

        let config = read_file(&mut ini);

        assert_eq!(config.vector_scale, expected, "`{written}` read back");
    }
}

/// A bitmap's reduced fit is a scale of its own — `screen 75`, the display's
/// fitted size reduced to a share of it — and the three bitmap settings read
/// it through the one cascade every scale is read by, so which of them a
/// `config.ini` key belongs to changes nothing. A plain number keeps meaning
/// a share of the file's own size, which is what it has always meant, and a
/// value that names no scale is not one.
#[test]
fn a_bitmaps_reduced_fit_takes_the_words_a_person_would_write() {
    for (written, expected) in [
        ("screen 75", PreviewScale::FitToScreenReduced(75)),
        (" 75 of screen ", PreviewScale::FitToScreenReduced(75)),
        ("Screen-50", PreviewScale::FitToScreenReduced(50)),
        ("screen10", PreviewScale::FitToScreenReduced(10)),
        ("50", PreviewScale::Percent(50)),
        ("fit", PreviewScale::FitToScreen),
    ] {
        for key in ["preview_scale", "video_scale", "animated_scale"] {
            let mut ini = Ini::new();
            ini.set(CONFIG_SECTION, key, Some(written.to_string()));

            let config = read_file(&mut ini);

            let scale = match key {
                "preview_scale" => config.preview_scale,
                "video_scale" => config.video_scale,
                _ => config.animated_scale,
            };

            assert_eq!(scale, expected, "`{key}` at `{written}` read back");
        }
    }

    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "video_scale", Some("screen".to_string()));

    let config = read_file(&mut ini);

    assert_eq!(
        config.video_scale, DEFAULT_VIDEO_SCALE,
        "a share with no number in it is not one, so the setting is left where it was"
    );
}

/// A reduced fit is written back as the share it asks for, so a share picked
/// from the `By Screen` group of a bitmap's submenu is still the share after
/// a restart: what the file holds is the scale the setting is.
#[test]
fn a_bitmaps_reduced_fit_reads_back_what_it_writes() {
    for scale in [
        PreviewScale::FitToScreen,
        PreviewScale::FitToScreenReduced(75),
        PreviewScale::FitToScreenReduced(10),
        PreviewScale::Percent(50),
    ] {
        let written = scale.as_str();

        assert_eq!(
            PreviewScale::from_str(&written),
            Some(scale),
            "`{written}` read back"
        );
    }

    assert_eq!(
        PreviewScale::FitToScreenReduced(75).as_str(),
        "screen 75",
        "the share a reduced fit asks for is what the file is written with"
    );
}

/// A specimen's scale is the fifth of them and a setting of its own like the four:
/// one key changing leaves the others where they were.
#[test]
fn a_specimens_scale_is_read_from_its_own_key() {
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "preview_scale", Some("400".to_string()));
    ini.set(CONFIG_SECTION, "vector_scale", Some("75".to_string()));
    ini.set(CONFIG_SECTION, "ebook_scale", Some("25".to_string()));
    ini.set(CONFIG_SECTION, "document_scale", Some("10".to_string()));
    ini.set(
        CONFIG_SECTION,
        "font_scale",
        Some(" Fit to Screen ".to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(config.font_scale, PreviewScale::FitToScreen);
    assert_eq!(config.vector_scale, PreviewScale::Percent(75));
    assert_eq!(config.ebook_scale, PreviewScale::Percent(25));
    assert_eq!(config.document_scale, PreviewScale::Percent(10));
    assert_eq!(config.preview_scale, PreviewScale::Percent(400));
}

/// A design document's scale is the fifth of them and a setting of its own like the
/// four: it starts at the whole of the room, and one key changing leaves the others
/// where they were.
#[test]
fn a_designs_scale_is_read_from_its_own_key() {
    assert_eq!(DEFAULT_DESIGN_SCALE, PreviewScale::FitToScreen);
    assert_eq!(AppConfig::default().design_scale, DEFAULT_DESIGN_SCALE);

    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "vector_scale", Some("75".to_string()));
    ini.set(CONFIG_SECTION, "ebook_scale", Some("25".to_string()));
    ini.set(CONFIG_SECTION, "document_scale", Some("fit".to_string()));
    ini.set(CONFIG_SECTION, "font_scale", Some("50".to_string()));
    ini.set(CONFIG_SECTION, "design_scale", Some(" 10% ".to_string()));

    let config = read_file(&mut ini);

    assert_eq!(config.design_scale, PreviewScale::Percent(10));
    assert_eq!(config.vector_scale, PreviewScale::Percent(75));
    assert_eq!(config.ebook_scale, PreviewScale::Percent(25));
    assert_eq!(config.document_scale, PreviewScale::FitToScreen);
    assert_eq!(config.font_scale, PreviewScale::Percent(50));
}

/// Which face of a collection a specimen is of is a key of its own: a file that has never
/// named one previews the first face, one that names a face is read at it, and a number
/// outside the range the menu offers is brought back into it rather than kept.
#[test]
fn the_face_of_a_collection_is_read_from_its_own_key() {
    let config = AppConfig::default();
    assert_eq!(config.ttc_face, DEFAULT_TTC_FACE);

    for (written, expected) in [
        ("3", 3),
        ("10", MAX_TTC_FACE),
        ("0", DEFAULT_TTC_FACE),
        ("99", MAX_TTC_FACE),
    ] {
        let mut ini = Ini::new();
        ini.set(CONFIG_SECTION, "ttc_face", Some(written.to_string()));

        let mut config = AppConfig::default();
        config.apply_ini(&ini);

        assert_eq!(config.ttc_face, expected, "`{written}` read back");
    }

    // A face and a specimen's scale are two keys: one does not move the other.
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "font_scale", Some("25".to_string()));
    ini.set(CONFIG_SECTION, "ttc_face", Some("2".to_string()));

    let config = read_file(&mut ini);

    assert_eq!(config.ttc_face, 2);
    assert_eq!(config.font_scale, PreviewScale::Percent(25));
}

/// A specimen is drawn at half the room the display has unless the file says otherwise:
/// the share an SVG document starts at, which is what a fresh install writes and what
/// the menu marks as the default.
#[test]
fn a_specimens_scale_starts_at_half_the_room() {
    let config = AppConfig::default();

    assert_eq!(config.font_scale, PreviewScale::Percent(50));
    assert_eq!(config.font_scale.as_str(), "50");
    assert_eq!(config.font_scale, DEFAULT_FONT_SCALE);
}

/// A specimen's backdrop is a key of its own: fonts are a kind of their own, so the setting
/// is read from its own name and from nothing else.
#[test]
fn a_specimens_backdrop_is_read_from_its_own_key() {
    let mut ini = Ini::new();
    ini.set(
        CONFIG_SECTION,
        "font_background",
        Some("checkerboard".to_string()),
    );

    let config = read_file(&mut ini);
    assert_eq!(config.font_background, TransparentBackground::Checkerboard);

    // A file that says nothing about one leaves the setting where it starts, which for a
    // specimen is the page it is written on.
    let ini = Ini::new();

    let mut config = AppConfig::default();
    config.apply_ini(&ini);
    assert_eq!(config.font_background, DEFAULT_FONT_BACKGROUND);
}

/// A design document's backdrop is a key of its own for the reason a specimen's is:
/// the kind is one of its own, so the key a picture was written under is read for a
/// picture and leaves this one where it starts.
#[test]
fn a_designs_backdrop_is_read_from_its_own_key() {
    let mut ini = Ini::new();
    ini.set(
        CONFIG_SECTION,
        "image_background",
        Some("white".to_string()),
    );
    ini.set(
        CONFIG_SECTION,
        "design_background",
        Some("checkerboard".to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(
        config.design_background,
        TransparentBackground::Checkerboard
    );
    assert_eq!(config.image_background, TransparentBackground::White);
}

/// The two volumes are two settings, each read from its own key, and each reduced to what a
/// player may be handed: a file written before the kind existed has no key for a sound's
/// volume and plays at the level a fresh installation does.
#[test]
fn reads_each_volume_from_its_own_key() {
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "video_volume", Some("35".to_string()));
    ini.set(CONFIG_SECTION, "audio_volume", Some("80".to_string()));

    let config = read_file(&mut ini);
    assert_eq!(config.video_volume, 35);
    assert_eq!(config.audio_volume, 80);

    let mut older = Ini::new();
    older.set(CONFIG_SECTION, "video_volume", Some("10".to_string()));
    assert_eq!(
        read_file(&mut older).audio_volume,
        DEFAULT_AUDIO_VOLUME,
        "a file written before the kind existed plays a sound at the level a fresh one does"
    );

    let mut loud = Ini::new();
    loud.set(CONFIG_SECTION, "audio_volume", Some("400".to_string()));
    assert_eq!(
        read_file(&mut loud).audio_volume,
        MAX_VOLUME,
        "a level past the loudest is the loudest, not what the file asked a player for"
    );
}

/// Where a sound starts is read from its own key, the way a volume is: a file that has
/// never named it leaves the setting where this build starts, a file that names one of the
/// four is read as it, and a file that names something else is left at the default rather
/// than starting sounds somewhere no one asked for.
#[test]
fn reads_where_a_sound_starts_from_its_own_key() {
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "audio_seek", Some("middle".to_string()));
    assert_eq!(read_file(&mut ini).audio_seek, AudioSeek::Middle);

    let mut older = Ini::new();
    older.set(CONFIG_SECTION, "audio_volume", Some("35".to_string()));
    assert_eq!(
        read_file(&mut older).audio_seek,
        DEFAULT_AUDIO_SEEK,
        "a file written before the setting existed starts a sound where a fresh one does"
    );

    let mut unknown = Ini::new();
    unknown.set(CONFIG_SECTION, "audio_seek", Some("sideways".to_string()));
    assert_eq!(
        read_file(&mut unknown).audio_seek,
        DEFAULT_AUDIO_SEEK,
        "a value the app cannot read is answered with the way it starts rather than guessed at"
    );
}

/// Where a *pinned* sound starts is a setting of its own, read from its own
/// key the way the hover's is — and it is a setting that starts at the
/// beginning rather than at where a sound was left: a pin is a sound the
/// user asked to hear, which is the one answer that does not have to be a
/// memory for.
#[test]
fn reads_where_a_pinned_sound_starts_from_its_own_key() {
    assert_eq!(
        AppConfig::default().pin_mode_audio_seek,
        DEFAULT_PIN_MODE_AUDIO_SEEK,
        "a pinned sound starts at the beginning unless the file says otherwise"
    );
    assert_eq!(
        DEFAULT_PIN_MODE_AUDIO_SEEK,
        AudioSeek::Start,
        "the pin's setting starts at the beginning, not where a sound was left"
    );

    let mut ini = Ini::new();
    ini.set(
        CONFIG_SECTION,
        "pin_mode_audio_seek",
        Some("middle".to_string()),
    );
    assert_eq!(read_file(&mut ini).pin_mode_audio_seek, AudioSeek::Middle);

    // What the menu writes is what the file reads back: every way of
    // starting a sound is one the setting keeps, so a choice made here is
    // still the choice after a restart.
    for seek in [
        AudioSeek::Remember,
        AudioSeek::Start,
        AudioSeek::Middle,
        AudioSeek::Random,
    ] {
        let mut ini = Ini::new();
        ini.set(
            CONFIG_SECTION,
            "pin_mode_audio_seek",
            Some(seek.as_str().to_string()),
        );
        assert_eq!(
            read_file(&mut ini).pin_mode_audio_seek,
            seek,
            "`{}` read back",
            seek.as_str()
        );
    }

    let mut older = Ini::new();
    older.set(CONFIG_SECTION, "audio_seek", Some("random".to_string()));
    assert_eq!(
        read_file(&mut older).pin_mode_audio_seek,
        DEFAULT_PIN_MODE_AUDIO_SEEK,
        "a file written before the setting existed starts a pinned sound where a fresh one does, whatever the hover's own setting is"
    );

    let mut unknown = Ini::new();
    unknown.set(
        CONFIG_SECTION,
        "pin_mode_audio_seek",
        Some("sideways".to_string()),
    );
    assert_eq!(
        read_file(&mut unknown).pin_mode_audio_seek,
        DEFAULT_PIN_MODE_AUDIO_SEEK,
        "a value the app cannot read is answered with the way it starts rather than guessed at"
    );

    // The two settings are two answers to one question, so a file that names
    // one names it alone: the hover's is left where the file put it, and the
    // pin's where the file put that.
    let mut both = Ini::new();
    both.set(CONFIG_SECTION, "audio_seek", Some("random".to_string()));
    both.set(
        CONFIG_SECTION,
        "pin_mode_audio_seek",
        Some("middle".to_string()),
    );
    let config = read_file(&mut both);
    assert_eq!(config.audio_seek, AudioSeek::Random);
    assert_eq!(config.pin_mode_audio_seek, AudioSeek::Middle);
}

/// Which files a pin's own previous/next buttons step through is one of its own keys, and
/// it is written under the name the tray's `Nav File Types` submenu shows: `category` is
/// the folder narrowed to the pinned file's own kind of thing, and every file this build
/// could preview is what a file that names neither — or names nothing at all — is read as.
#[test]
fn reads_which_files_the_pin_steps_through_from_its_own_key() {
    let mut ini = Ini::new();
    ini.set(
        CONFIG_SECTION,
        "pin_nav_file_types",
        Some("category".to_string()),
    );
    assert_eq!(
        read_file(&mut ini).pin_nav_file_types,
        PinNavFileTypes::Category
    );

    let mut unknown = Ini::new();
    unknown.set(
        CONFIG_SECTION,
        "pin_nav_file_types",
        Some("sideways".to_string()),
    );
    assert_eq!(
        read_file(&mut unknown).pin_nav_file_types,
        DEFAULT_PIN_NAV_FILE_TYPES,
        "a value the app cannot read is answered with what a fresh one walks rather than \
         guessed at"
    );

    let mut older = Ini::new();
    older.set(CONFIG_SECTION, "pin_key", Some("space".to_string()));
    assert_eq!(
        read_file(&mut older).pin_nav_file_types,
        DEFAULT_PIN_NAV_FILE_TYPES,
        "a file written before the setting existed walks every file a fresh one walks"
    );
}

/// The key is written under the name it is read under, so the file a user edits by hand
/// is the file the app wrote: what comes back out of `save` is what went in, and the
/// heading table names it (see `every_setting_the_app_writes_is_one_the_table_names`).
#[test]
fn the_pin_navigation_setting_survives_a_write_and_a_read() {
    let config = AppConfig {
        pin_nav_file_types: PinNavFileTypes::Category,
        ..AppConfig::default()
    };

    let mut written = Ini::new();
    let _ = written.read(ordered_text(&config.to_ini()));
    assert_eq!(
        written.get(CONFIG_SECTION, "pin_nav_file_types"),
        Some("category".to_string())
    );

    let mut read_back = Ini::new();
    read_back
        .read(ordered_text(&config.to_ini()))
        .expect("a file this app wrote is one it can read");
    assert_eq!(
        read_file(&mut read_back).pin_nav_file_types,
        PinNavFileTypes::Category
    );
}

/// What a pin collapsed into its bubble does with what it is playing is two settings, each
/// read from its own key: a video behind one and a sound behind the other, so a file that
/// turns one off leaves the other where it was. Both hold what is playing unless the file
/// says otherwise, which is what the tray offers and what a file written before the two
/// switches existed is read as.
#[test]
fn a_collapsed_pin_holds_what_it_plays_by_each_kinds_own_switch() {
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "pin_pause_video", Some("false".to_string()));
    ini.set(CONFIG_SECTION, "pin_pause_audio", Some("true".to_string()));

    let config = read_file(&mut ini);
    assert!(!config.pin_pause_video);
    assert!(config.pin_pause_audio);

    let mut older = Ini::new();
    older.set(CONFIG_SECTION, "pin_key", Some("space".to_string()));

    let config = read_file(&mut older);
    assert_eq!(
        config.pin_pause_video, DEFAULT_PIN_PAUSE_VIDEO,
        "a file written before the switch existed holds a film the way a fresh one does"
    );
    assert_eq!(
        config.pin_pause_audio, DEFAULT_PIN_PAUSE_AUDIO,
        "and a sound the same way"
    );
}

/// A texture is offered two backdrops of the four, which is what `dds_image` is written
/// against: the two that show what stands behind a preview are read as the white page the
/// setting starts at, so a file that still names one is not left holding a value the menu
/// beside it has no item for.
#[test]
fn a_textures_backdrop_is_one_of_the_two_it_is_offered() {
    for kept in [TransparentBackground::Black, TransparentBackground::White] {
        assert_eq!(sanitize_dds_background(kept), kept, "`{kept:?}` is offered");
    }

    for dropped in [
        TransparentBackground::Transparent,
        TransparentBackground::Checkerboard,
    ] {
        assert_eq!(
            sanitize_dds_background(dropped),
            DEFAULT_DDS_BACKGROUND,
            "`{dropped:?}` is not a backdrop a texture is offered"
        );
    }

    let mut ini = Ini::new();
    ini.set(
        CONFIG_SECTION,
        "dds_background",
        Some("checkerboard".to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(config.dds_background, DEFAULT_DDS_BACKGROUND);
}

/// A page is offered three backdrops of the four, which is what `html_background` is
/// written against: everything but transparency, which is not something a page is drawn
/// over, so a file that still names it is read as the white page the setting starts at
/// rather than left holding a value the menu beside it has no item for.
#[test]
fn a_pages_backdrop_is_one_of_the_three_it_is_offered() {
    for kept in [
        TransparentBackground::White,
        TransparentBackground::Black,
        TransparentBackground::Checkerboard,
    ] {
        assert_eq!(
            sanitize_html_background(kept),
            kept,
            "`{kept:?}` is offered"
        );
    }

    assert_eq!(
        sanitize_html_background(TransparentBackground::Transparent),
        DEFAULT_HTML_BACKGROUND,
        "transparency is not a backdrop a page is offered"
    );

    let mut ini = Ini::new();
    ini.set(
        CONFIG_SECTION,
        "html_background",
        Some("transparent".to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(config.html_background, DEFAULT_HTML_BACKGROUND);
}

/// A page's backdrop is a key of its own for the reason a specimen's is: the kind is one
/// of its own, so the key a picture was written under is read for a picture and leaves
/// this one where it starts.
#[test]
fn a_pages_backdrop_is_read_from_its_own_key() {
    let mut ini = Ini::new();
    ini.set(
        CONFIG_SECTION,
        "image_background",
        Some("checkerboard".to_string()),
    );
    ini.set(CONFIG_SECTION, "html_background", Some("black".to_string()));

    let config = read_file(&mut ini);

    assert_eq!(config.html_background, TransparentBackground::Black);

    // A file that says nothing about one leaves the setting where it starts, which for a
    // page is the page it is written on.
    let ini = Ini::new();

    let mut config = AppConfig::default();
    config.apply_ini(&ini);
    assert_eq!(config.html_background, DEFAULT_HTML_BACKGROUND);
}

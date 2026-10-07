use super::*;

/// The one budget where an older build had two — each engine had a cache of its own. A file
/// that still names either key is read once for the larger of the two, since one budget
/// replaces both, and neither line is written again.
#[test]
fn one_document_cache_replaces_the_two_budgets_an_older_file_names() {
    let older = |keys: &[(&str, &str)]| {
        let written = written_file();
        let mut ini = Ini::new();
        for (section, held) in written.get_map_ref() {
            for (key, value) in held {
                if key != "document_cache_mb" {
                    ini.set(section, key, value.clone());
                }
            }
        }
        for (key, value) in keys {
            ini.set(CONFIG_SECTION, key, Some((*value).to_string()));
        }

        ini
    };

    let mut config = AppConfig::default();
    config.apply_ini(&older(&[
        ("office_cache_mb", "1024"),
        ("libre_cache_mb", "32"),
    ]));
    assert_eq!(
        config.document_cache_mb, 1024,
        "the larger of the two budgets is the one a single cache is given"
    );

    let mut config = AppConfig::default();
    config.apply_ini(&older(&[("libre_cache_mb", "64")]));
    assert_eq!(
        config.document_cache_mb, 64,
        "and a file naming one of them alone is read as it stands"
    );
    assert!(
        config.differs(&older(&[("libre_cache_mb", "64")])),
        "a file naming a budget this build does not write is one to write again"
    );
}

/// A setting the file does not have is one the app has to write: the line was deleted, a
/// whole section was, or the setting is one this build has and the file was written before
/// it existed. A file that holds everything this build writes is one there is nothing to do.
#[test]
fn a_setting_the_file_does_not_have_is_a_reason_to_write_it() {
    let config = AppConfig::default();

    let mut partial = Ini::new();
    partial.set(CONFIG_SECTION, "preview_enabled", Some("true".to_string()));

    assert!(
        config.differs(&partial),
        "a file holding one setting of fifty does not say what the app is using"
    );
    assert!(
        !config.differs(&written_file()),
        "a file holding every one of them is one to leave alone"
    );
}

/// The file is what this app writes and nothing else: a key of the user's own, or a section
/// of their own, is a difference like any other, and the write it asks for is what drops it.
/// The one thing this does not touch is a comment, which is not a key and is not compared.
#[test]
fn a_key_the_app_does_not_write_is_a_reason_to_write_it() {
    let mut ini = written_file();
    assert!(
        !AppConfig::default().differs(&ini),
        "a file holding what the app writes is one to leave alone"
    );

    ini.set(CONFIG_SECTION, "cat", Some("yes".to_string()));
    assert!(
        AppConfig::default().differs(&ini),
        "a key of the user's own is not a key the app writes"
    );

    let mut own_section = written_file();
    own_section.set("mine", "key", Some("yes".to_string()));
    assert!(
        AppConfig::default().differs(&own_section),
        "and neither is a section of their own"
    );
}

/// A value the app cannot read is one it does not keep, and a file holding one is a file to
/// write: what was written by hand — a tone map that is not one of them, a delay past the
/// ceiling, a face of a collection past the last one there is — is replaced by the value the
/// app is actually using, so the file says what the app does rather than what it could not
/// read.
#[test]
fn a_value_the_app_cannot_read_is_written_back_as_the_one_it_uses() {
    let mut ini = written_file();
    ini.set(
        CONFIG_SECTION,
        "hdr_tone_map",
        Some("reinharzzz".to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(
        config.hdr_tone_map, DEFAULT_HDR_TONE_MAP,
        "a curve that is not one of them leaves the setting where it starts"
    );
    assert!(
        config.differs(&ini),
        "and the file, which names no curve the app reads, is one to write"
    );

    // And the same for a value a setting reduces: a delay past its ceiling is read as the
    // ceiling, so a file asking for more than that is not one the app would write.
    let mut ini = written_file();
    ini.set(
        CONFIG_SECTION,
        "spinner_delay_ms",
        Some((MAX_SPINNER_DELAY_MS * 10).to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(config.spinner_delay_ms, MAX_SPINNER_DELAY_MS);
    assert!(config.differs(&ini));

    // A tick is reduced too, at both ends: below the system's own clock the number
    // asks for a wait nothing can honour, and past a second it says nothing that
    // turning previews off does not say better.
    let mut ini = written_file();
    ini.set(CONFIG_SECTION, "tick_ms", Some("0".to_string()));

    let config = read_file(&mut ini);

    assert_eq!(
        config.tick_ms, MIN_TICK_MS,
        "a tick of nothing is the floor"
    );
    assert!(config.differs(&ini));
}

/// A file the app writes is a file it reads back as itself: reading the text `save` puts on
/// disk gives the same settings, and writing those out again gives the same text. This is
/// what the whole check rests on — a file that read back as something else would be written
/// on every read, the watcher would see that write as a change, and the app would spend the
/// rest of the run rewriting a file it had just written.
#[test]
fn a_file_the_app_wrote_is_one_it_reads_back_as_itself() {
    let config = AppConfig {
        theme: TextTheme::Dark,
        avoid_mode: AvoidMode::Details,
        hdr_tone_map: Curve::Aces,
        hdr_exposure: -2.5,
        decode_budget_gb: 0.5,
        office_engine_idle: EngineIdle::Indefinite,
        webview_idle: EngineIdle::Seconds(60),
        audio_scale: PreviewScale::Percent(20),
        ebook_scale: PreviewScale::Percent(25),
        image_background: TransparentBackground::Transparent,
        text_scroll_far_edge_grace_pixels: 12.5,
        document_cache_mb: 1024,
        tick_ms: 47,
        trigger_key_affect_pin_mode: true,
        ..Default::default()
    };

    for config in [AppConfig::default(), config] {
        let written = ordered_text(&config.to_ini());

        let mut ini = Ini::new();
        ini.read(written.clone()).expect("a file this app wrote");

        let mut read_back = AppConfig::default();
        read_back.apply_ini(&ini);

        assert_eq!(
            ordered_text(&read_back.to_ini()),
            written,
            "the file is read back as the file it is"
        );
        assert!(
            !read_back.differs(&ini),
            "and it is a file there is nothing to write"
        );
    }
}

/// The repair runs on every read, so a file it has already been through has to come out of
/// it untouched: a file written every time it is read is a file whose mtime moves under the
/// watcher, and the app would keep reading its own write back for as long as it ran.
#[test]
fn a_file_the_repair_has_been_through_is_left_alone() {
    let mut ini = written_file();
    ini.set(
        lists::IMAGE.section,
        "extensions",
        Some(lists::IMAGE_EXTENSIONS_BEFORE_SVG.to_string()),
    );

    assert!(
        lists::repair_older_lists(&mut ini),
        "the file had a list of the app's own from before it grew"
    );

    let once = ordered_text(&ini);

    assert!(
        !lists::repair_older_lists(&mut ini),
        "the file it made is one there is nothing left to do to"
    );
    assert_eq!(
        ordered_text(&ini),
        once,
        "and it is the file it was left as"
    );
}

/// The cost of telling this app's own older lists from a user's by their entries, said out
/// loud: an entry taken out of a list comes back where what is left is exactly a list this
/// app shipped. `dds` is the entry to take out — the list without it is the one this app
/// shipped before it was added — while a list with anything else changed is the user's and
/// is kept, which the test above this one covers.
#[test]
fn an_entry_taken_out_of_a_list_can_come_back() {
    let mut ini = Ini::new();
    ini.set(
        lists::IMAGE.section,
        "extensions",
        Some(lists::sanitize_extension_list(lists::IMAGE_EXTENSIONS_BEFORE_DDS).join(",")),
    );

    let config = read_file(&mut ini);

    assert!(
        config.image_extensions.contains(&"dds".to_string()),
        "the list is read as this app's own older one, and `dds` is put back into it"
    );
}

/// A marker is answered once and gone with the answer: what it asks for is applied by the
/// load that finds it, so one left behind would be applied again on a later start, after
/// the user had set something of their own.
#[test]
fn a_reset_marker_is_taken_once_and_removed() {
    let folder = std::env::temp_dir().join("rust-hover-preview-marker-test");

    let _ = fs::remove_dir_all(&folder);
    fs::create_dir_all(&folder).unwrap();

    assert_eq!(
        take_reset_markers(&folder),
        (false, false),
        "an installation that left nothing has nothing to apply"
    );

    fs::write(folder.join("reset-settings.marker"), b"").unwrap();

    assert_eq!(take_reset_markers(&folder), (true, false));
    assert!(
        !folder.join("reset-settings.marker").exists(),
        "and the marker went with the answer"
    );
    assert_eq!(
        take_reset_markers(&folder),
        (false, false),
        "so the next start applies nothing"
    );

    fs::write(folder.join("reset-extensions.marker"), b"").unwrap();

    assert_eq!(take_reset_markers(&folder), (false, true));

    let _ = fs::remove_dir_all(&folder);
}

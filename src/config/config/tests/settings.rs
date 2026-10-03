use super::*;

/// The delay a hover's load is given before the spinner goes up is a setting of its
/// own, read from its own key in milliseconds: `0` is a delay like any other, and a
/// number past the ceiling is reduced to it.
#[test]
fn the_spinner_delay_is_read_from_its_own_key() {
    assert_eq!(
        AppConfig::default().spinner_delay_ms,
        DEFAULT_SPINNER_DELAY_MS,
        "a load is given a quarter of a second unless the file says otherwise"
    );

    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "spinner_delay_ms", Some("900".to_string()));

    let config = read_file(&mut ini);
    assert_eq!(config.spinner_delay_ms, 900);

    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "spinner_delay_ms", Some("0".to_string()));

    let config = read_file(&mut ini);
    assert_eq!(config.spinner_delay_ms, 0, "`0` is a delay like any other");

    let mut ini = Ini::new();
    ini.set(
        CONFIG_SECTION,
        "spinner_delay_ms",
        Some((MAX_SPINNER_DELAY_MS + 1).to_string()),
    );

    let config = read_file(&mut ini);
    assert_eq!(config.spinner_delay_ms, MAX_SPINNER_DELAY_MS);
}

/// The settings section is written under the headings the tray lists its menus under, in
/// the order the tray lists them, with the keys of a heading in alphabetical order and the
/// file lists left to the sections below — which is the shape a person reads, rather than
/// one alphabetical run of every key the app knows.
#[test]
fn the_settings_are_written_under_the_headings_the_tray_lists_them_under() {
    let mut ini = Ini::new();
    // Set out of order, and out of the order the headings run in, so what is tested is the
    // writer's order rather than the order they were handed over in.
    ini.set(CONFIG_SECTION, "video_volume", Some("50".to_string()));
    ini.set(CONFIG_SECTION, "avoid_mode", Some("details".to_string()));
    ini.set(CONFIG_SECTION, "preview_enabled", Some("true".to_string()));
    ini.set(
        CONFIG_SECTION,
        "image_background",
        Some("white".to_string()),
    );
    ini.set(
        lists::IMAGE.section,
        "extensions",
        Some("png,jpg".to_string()),
    );

    assert_eq!(
        ordered_text(&ini),
        "\
[settings]
; General
preview_enabled=true

; Placement
avoid_mode=details

; Background
image_background=white

; Volume
video_volume=50

[image]
extensions=png,jpg
"
    );
}

/// The engine settings are written under a heading of their own, below the one the caches
/// and the budget are under, which is where the tray lists them.
#[test]
fn the_engine_settings_are_written_under_a_heading_of_their_own() {
    let written = ordered_text(&AppConfig::default().to_ini());

    let performance = written
        .find("; Performance\n")
        .expect("the settings are grouped under a heading for what the app costs");
    let engine = written
        .find("; Engine\n")
        .expect("and under one for the engines themselves");
    let advanced = written
        .find("; Advanced\n")
        .expect("and under one for the settings the tray has no item for");

    assert!(
        performance < engine,
        "the caches are listed above the engines"
    );
    assert!(engine < advanced, "and the engines above the rest");

    let under_engine = &written[engine..advanced];
    for key in [
        "libreoffice_idle",
        "office_engine",
        "office_engine_idle",
        "webview_idle",
    ] {
        assert!(under_engine.contains(key), "`{key}` is under `Engine`");
    }
    for key in ["decode_budget_gb", "image_cache_mb"] {
        assert!(!under_engine.contains(key), "`{key}` is not under `Engine`");
    }

    // A file this build wrote is a file there is nothing to write again, which is what keeps
    // the watcher from writing the file it has just read back.
    assert!(
        !headings_are_old(&written),
        "a file this build wrote needs nothing done to it"
    );
}

/// A file grouped the way an older build grouped it is written again under the headings of
/// this one. A heading is a comment, so what such a file holds is read exactly as any other
/// file is: the grouping is the one thing about it that is not what this build writes, and
/// the one thing that a write puts right.
#[test]
fn a_file_written_before_the_settings_were_regrouped_is_written_again() {
    let written = ordered_text(&AppConfig::default().to_ini());

    // The same file as an older build wrote it: the engines had no heading of their own, so
    // the settings they are named by sat among the caches and the budget.
    let older = written.replace("; Engine\n", "");
    assert!(
        headings_are_old(&older),
        "a file with no heading for the engines is one to write again"
    );

    // A file with every heading, listed in the order the build before this one listed them —
    // the engines above the caches — is one to write again as well: the menus were
    // rearranged, and the headings say so.
    let (head, rest) = written
        .split_once("; Performance\n")
        .expect("a heading for the caches");
    let (performance, rest) = rest
        .split_once("; Engine\n")
        .expect("a heading for the engines");
    let (engine, tail) = rest
        .split_once("; Advanced\n")
        .expect("a heading for the rest");
    let rearranged =
        format!("{head}; Engine\n{engine}; Performance\n{performance}; Advanced\n{tail}");
    assert!(
        headings_are_old(&rearranged),
        "a file listing the headings in another order is one to write again"
    );

    // An editor that saves the file with the other line ending leaves one there is nothing
    // to write either: the heading is still the line it was, and what an editor adds to the
    // end of it is not the app's business.
    assert!(
        !headings_are_old(&written.replace('\n', "\r\n")),
        "a heading is not read by the line ending it happens to have"
    );

    // And the keys are read the same either way, which is why the grouping can be put right
    // without anything being migrated: what moved is the comment, not the setting.
    let mut read_back = Ini::new();
    assert!(read_back.read(older).is_ok());
    assert_eq!(
        read_back.get(CONFIG_SECTION, "office_engine"),
        Some("microsoft_office".to_string())
    );
}

/// Every setting the app writes is a setting the heading table names. The table is what the
/// headings of a file are written from, and it is what the repair reads to tell whether a
/// file is missing a setting, so a key `save` writes that the table does not name would be
/// written under `; Ungrouped` — and its absence from a file would never be noticed.
#[test]
fn every_setting_the_app_writes_is_one_the_table_names() {
    let written = AppConfig::default().to_ini();
    let keys = written
        .get_map_ref()
        .get(CONFIG_SECTION)
        .expect("a settings section");

    for key in keys.keys() {
        assert!(
            SETTING_GROUPS
                .iter()
                .any(|(_, group)| group.contains(&key.as_str())),
            "`{key}` is written by `save` and named by no heading"
        );
    }

    assert_eq!(
        keys.len(),
        SETTING_GROUPS
            .iter()
            .map(|(_, group)| group.len())
            .sum::<usize>(),
        "the table names as many settings as `save` writes"
    );
}

/// A heading is a comment, so the file the app writes is one it can read back: the reader
/// steps over the headings and the blank lines they are kept apart by, and every value
/// comes back as it was written.
#[test]
fn a_file_written_under_the_headings_reads_back_as_it_was() {
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "avoid_mode", Some("details".to_string()));
    ini.set(CONFIG_SECTION, "hover_delay_ms", Some("200".to_string()));
    ini.set(
        lists::VECTOR.section,
        "extensions",
        Some("svg,svgz".to_string()),
    );

    let mut read_back = Ini::new();
    read_back
        .read(ordered_text(&ini))
        .expect("a file this app wrote is one it can read");

    assert_eq!(
        read_back.get(CONFIG_SECTION, "avoid_mode"),
        Some("details".to_string())
    );
    assert_eq!(
        read_back.get(CONFIG_SECTION, "hover_delay_ms"),
        Some("200".to_string())
    );
    assert_eq!(
        read_back.get(lists::VECTOR.section, "extensions"),
        Some("svg,svgz".to_string())
    );
}

/// A setting the table has not been told about is written last, under a heading of its
/// own: a key added to `save` and forgotten there lands at the bottom of the file where it
/// can be seen, rather than inside a heading it has nothing to do with or nowhere at all.
#[test]
fn a_setting_the_table_does_not_know_is_written_under_its_own_heading() {
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "something_new", Some("1".to_string()));
    ini.set(CONFIG_SECTION, "preview_enabled", Some("true".to_string()));

    assert_eq!(
        ordered_text(&ini),
        "\
[settings]
; General
preview_enabled=true

; Ungrouped
something_new=1
"
    );
}

/// One setting belongs to one heading. A key listed under two of them is written under
/// both, and since what is read back is the key rather than the heading, the second one is
/// the one the setting takes.
#[test]
fn every_setting_belongs_to_one_heading() {
    let mut seen: Vec<&str> = Vec::new();

    for (heading, keys) in SETTING_GROUPS {
        assert!(
            !keys.is_empty(),
            "the heading `{heading}` lists no settings"
        );

        for key in *keys {
            assert!(
                !seen.contains(key),
                "`{key}` is listed under more than one heading"
            );
            seen.push(key);
        }
    }
}

/// The defaults the tray marks are the ones the app starts at: a picture, a drawing and a
/// design document are drawn over the squares, a specimen and a texture over a white page,
/// and the placement, the two delays and the volume start where the menu says.
#[test]
fn the_defaults_the_tray_marks_are_the_ones_the_app_starts_at() {
    let config = AppConfig::default();

    assert_eq!(config.image_background, DEFAULT_IMAGE_BACKGROUND);
    assert_eq!(config.vector_background, DEFAULT_VECTOR_BACKGROUND);
    assert_eq!(config.design_background, DEFAULT_DESIGN_BACKGROUND);
    assert_eq!(config.image_background, TransparentBackground::Checkerboard);

    assert_eq!(config.font_background, DEFAULT_FONT_BACKGROUND);
    assert_eq!(config.font_background, TransparentBackground::White);
    assert_eq!(config.dds_background, DEFAULT_DDS_BACKGROUND);
    assert_eq!(config.dds_background, TransparentBackground::White);
    assert_eq!(config.html_background, DEFAULT_HTML_BACKGROUND);
    assert_eq!(config.html_background, TransparentBackground::White);

    assert_eq!(config.avoid_mode, DEFAULT_AVOID_MODE);
    assert_eq!(config.avoid_mode, AvoidMode::Filename);
    assert_eq!(config.follow_cursor, DEFAULT_FOLLOW_CURSOR);
    assert!(
        !config.follow_cursor,
        "a preview is placed at its best position"
    );

    assert_eq!(config.hover_delay_ms, DEFAULT_HOVER_DELAY_MS);
    assert_eq!(config.hover_delay_ms, 0);
    assert_eq!(
        config.same_file_rehover_delay_ms,
        DEFAULT_SAME_FILE_REHOVER_DELAY_MS
    );
    assert_eq!(config.same_file_rehover_delay_ms, 200);
    assert_eq!(config.tick_ms, DEFAULT_TICK_MS);
    assert_eq!(config.tick_ms, 15, "the loop looks once a system tick");
    assert_eq!(config.video_volume, DEFAULT_VIDEO_VOLUME);
    assert_eq!(config.video_volume, 0, "a hover never makes a sound");
}

/// The settings reset puts every setting back where this build starts it, and the lists
/// are not settings: what a user has made of one is theirs, so the lists are moved out of
/// the way and put back exactly as they were.
#[test]
fn the_settings_reset_leaves_the_extension_lists_alone() {
    let mut config = AppConfig {
        image_cache_mb: 16,
        document_cache_mb: 2048,
        preview_enabled: false,
        image_extensions: lists::sanitize_extension_list("png,dng"),
        ..Default::default()
    };

    config.text_names.push("notes".to_string());

    let image = config.image_extensions.clone();
    let names = config.text_names.clone();

    config.reset_to_recommended();

    assert_eq!(
        config.image_cache_mb, DEFAULT_IMAGE_CACHE_MB,
        "the cache this release moved is the release's to move"
    );
    assert_eq!(config.document_cache_mb, DEFAULT_DOCUMENT_CACHE_MB);
    assert!(config.preview_enabled);
    assert_eq!(
        config.image_extensions, image,
        "and the lists stay the user's"
    );
    assert_eq!(config.text_names, names);
}

/// The lists reset is the other half and the whole of what it does: an edited list is the
/// built-in one again, and no setting moves.
#[test]
fn the_lists_reset_restores_the_built_in_lists_alone() {
    let mut config = AppConfig {
        image_cache_mb: 16,
        image_extensions: lists::sanitize_extension_list("png,dng"),
        ebook_extensions: Vec::new(),
        ..Default::default()
    };

    config.reset_extension_lists();

    assert_eq!(
        config.image_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_IMAGE_EXTENSIONS)
    );
    assert_eq!(
        config.ebook_extensions,
        crate::formats::text_formats::sanitize_extension_list(lists::DEFAULT_EBOOK_EXTENSIONS),
        "a list emptied by hand is a list this puts back"
    );
    assert_eq!(config.image_cache_mb, 16, "and nothing else was touched");
}

/// What the two reset rows are offered on, and what the question each of them asks is
/// written from: the difference between the configuration and the one this build
/// recommends, named by the keys the file writes it under.
#[test]
fn a_configuration_at_the_recommended_values_has_nothing_apart() {
    let config = AppConfig::default();

    assert!(config.settings_apart_from_recommended().is_empty());
    assert!(config.lists_apart_from_built_in().is_empty());
}

#[test]
fn the_differences_are_named_by_the_keys_the_file_writes() {
    let config = AppConfig {
        image_cache_mb: 16,
        image_extensions: lists::sanitize_extension_list("png,dng"),
        ..Default::default()
    };

    assert_eq!(
        config.settings_apart_from_recommended(),
        vec![(
            "image_cache_mb".to_string(),
            "16".to_string(),
            DEFAULT_IMAGE_CACHE_MB.to_string(),
        )]
    );
    assert_eq!(
        config.lists_apart_from_built_in(),
        vec!["image".to_string()],
        "the section the edit sits under, and only that one"
    );
}

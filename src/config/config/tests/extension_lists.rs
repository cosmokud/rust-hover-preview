use super::*;

/// The font list is the fifth of them and behaves like the rest: a file that has never
/// named it is written out with the built-in entries, and one that has been edited keeps
/// what the user wrote.
#[test]
fn the_font_list_is_written_out_and_read_back() {
    let config = AppConfig::default();
    assert_eq!(
        config.font_extensions,
        crate::formats::text_formats::sanitize_extension_list(lists::DEFAULT_FONT_EXTENSIONS)
    );

    let mut ini = Ini::new();
    ini.set(
        lists::FONT.section,
        "extensions",
        Some(".OTF,ttf".to_string()),
    );

    let config = read_file(&mut ini);
    assert_eq!(config.font_extensions, vec!["otf", "ttf"]);

    // A file with no section at all is answered with the built-in list, and the file is one
    // to write out again with it, since a key that is gone is not what the app is using.
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "run_at_startup", Some("true".to_string()));

    let config = read_file(&mut ini);
    assert_eq!(
        config.font_extensions,
        crate::formats::text_formats::sanitize_extension_list(lists::DEFAULT_FONT_EXTENSIONS)
    );
    assert!(
        config.differs(&ini),
        "and the file, which has no such section, is one to write"
    );
}

/// Fonts are their own kind under `Preview Types`: the gate is a switch of its own, and
/// switching it leaves the list and every other kind where they were.
#[test]
fn the_fonts_gate_is_a_switch_of_its_own() {
    let mut config = AppConfig::default();
    assert!(PreviewType::Fonts.enabled_in(&config));

    PreviewType::Fonts.set_enabled_in(&mut config, false);
    assert!(!PreviewType::Fonts.enabled_in(&config));
    assert!(PreviewType::Vector.enabled_in(&config));
    assert!(PreviewType::Images.enabled_in(&config));

    PreviewType::Fonts.set_enabled_in(&mut config, true);
    assert!(PreviewType::Fonts.enabled_in(&config));
}

/// The pictures the ImageMagick engine is asked about are a list of their own in a section
/// of their own — written from the built-in list on first run, normalized on the way in
/// and out, and read back from the file — and the kind they belong to has a gate of its
/// own under `Preview Types`.
///
/// What the engine's own work costs is not a setting of its own, and there is nothing to
/// assert about one: a picture it develops is held in the image cache, under the budget
/// pictures have always been held under, and a file whose frame has been given up is
/// developed again rather than kept in a file of the app's own (see `imagemagick_render`).
#[test]
fn the_pictures_the_magick_engine_is_asked_about_are_a_list_and_a_gate_of_their_own() {
    let config = AppConfig::default();
    assert_eq!(
        config.magick_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_MAGICK_EXTENSIONS)
    );

    let mut ini = Ini::new();
    ini.set(
        lists::MAGICK.section,
        "extensions",
        Some(".NEF, cr3".to_string()),
    );

    let config = read_file(&mut ini);
    assert_eq!(config.magick_extensions, vec!["nef", "cr3"]);

    // A section that is gone is answered with the built-in list, and the file is one to
    // write out again with it, since a key that is gone is not what the app is using.
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "run_at_startup", Some("true".to_string()));

    let config = read_file(&mut ini);
    assert_eq!(
        config.magick_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_MAGICK_EXTENSIONS)
    );
    assert!(
        config.differs(&ini),
        "and the file, which has no such section, is one to write"
    );

    // And a file holding the list this app shipped before the names nobody had asked for were
    // added to it is this app's own rather than a user's edit, so it is brought up to the list
    // of now — which is how an installation that already exists is given them.
    let mut ini = Ini::new();
    ini.set(
        lists::MAGICK.section,
        "extensions",
        Some(lists::MAGICK_EXTENSIONS_BEFORE_THE_REST.to_string()),
    );

    let config = read_file(&mut ini);
    assert_eq!(
        config.magick_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_MAGICK_EXTENSIONS)
    );
    for name in ["sun", "pict", "rgb", "ase", "fax"] {
        assert!(
            config.magick_extensions.iter().any(|entry| entry == name),
            "`{name}` is one of the names added since"
        );
    }

    // The gate is not a switch of its own: a picture an engine develops is a picture, so
    // it is switched by the Images gate and by nothing else, and the list is left where it
    // was either way.
    let mut config = AppConfig::default();
    assert!(PreviewType::Magick.enabled_in(&config));

    PreviewType::Magick.set_enabled_in(&mut config, false);
    assert!(!PreviewType::Magick.enabled_in(&config));
    assert!(!PreviewType::Images.enabled_in(&config));
    assert!(PreviewType::Libre.enabled_in(&config));
    assert_eq!(
        config.magick_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_MAGICK_EXTENSIONS)
    );

    PreviewType::Images.set_enabled_in(&mut config, true);
    assert!(PreviewType::Magick.enabled_in(&config));
}

/// And the books the ebook engine is asked about are the same shape once more: a section of
/// their own in `config.ini`, read back the way it is written, behind a gate that is the one
/// books have always had.
///
/// The engine's own work costs nothing to bound and has no setting of its own, which is why
/// there is none to assert about: a book it converted is kept as a page, under the budget pages
/// have always been kept under, and there is no idle time because there is no process — see
/// `calibre_render`, which is why the `Engine` submenu has no `Calibre TTL` row.
#[test]
fn the_books_the_calibre_engine_is_asked_about_are_a_list_and_a_gate_of_their_own() {
    let config = AppConfig::default();
    assert_eq!(
        config.calibre_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_CALIBRE_EXTENSIONS)
    );

    let mut ini = Ini::new();
    ini.set(
        lists::CALIBRE.section,
        "extensions",
        Some(".MOBI, epub".to_string()),
    );

    let config = read_file(&mut ini);
    assert_eq!(config.calibre_extensions, vec!["mobi", "epub"]);

    // A section that is gone is answered with the built-in list, and the file is one to write
    // out again with it, since a key that is gone is not what the app is using.
    let mut ini = Ini::new();
    ini.set(CONFIG_SECTION, "run_at_startup", Some("true".to_string()));

    let config = read_file(&mut ini);
    assert_eq!(
        config.calibre_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_CALIBRE_EXTENSIONS)
    );
    assert!(
        config.differs(&ini),
        "and the file, which has no such section, is one to write"
    );

    // And the gate is over the list rather than through it: a user who switches books off has
    // not edited which books the engine is asked about, so switching them back on restores
    // exactly what was configured.
    let mut config = AppConfig::default();
    PreviewType::Calibre.set_enabled_in(&mut config, false);
    assert!(!PreviewType::Ebook.enabled_in(&config));
    assert_eq!(
        config.calibre_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_CALIBRE_EXTENSIONS)
    );

    PreviewType::Ebook.set_enabled_in(&mut config, true);
    assert!(PreviewType::Calibre.enabled_in(&config));
}

/// The three other engine-drawn kinds share their gate the same way: a document an engine
/// drew is switched by the `Document` gate, an archive a listing engine read by the
/// `Archives` gate and a book an ebook engine converted by the `Ebook` gate, whichever of
/// the pair the switch is thrown from.
#[test]
fn an_engine_drawn_kind_is_switched_by_the_kind_it_belongs_to() {
    let mut config = AppConfig::default();

    for (kind, gate) in [
        (PreviewType::Libre, PreviewType::Document),
        (PreviewType::Peazip, PreviewType::Archives),
        (PreviewType::Magick, PreviewType::Images),
        (PreviewType::Calibre, PreviewType::Ebook),
    ] {
        kind.set_enabled_in(&mut config, false);

        assert!(!kind.enabled_in(&config));
        assert!(
            !gate.enabled_in(&config),
            "the gate the kind belongs to came down with it"
        );

        gate.set_enabled_in(&mut config, true);

        assert!(kind.enabled_in(&config), "and back up with it");
        assert!(gate.enabled_in(&config));
    }
}

/// The switch over the two names a page of HTML goes by is a key of its own, and it
/// starts off: a `.htm` is a page of text until something says otherwise, which is what
/// every file written before the setting existed says by not having it.
#[test]
fn a_page_of_html_is_markup_until_the_tray_says_otherwise() {
    assert_eq!(AppConfig::default().render_html, DEFAULT_RENDER_HTML);

    // A file that names no such key reads as the markup and is written back with the key
    // the next time it is saved, so the file the app keeps is one that can be read again.
    let mut ini = Ini::new();
    let config = read_file(&mut ini);
    assert!(!config.render_html);

    let written = ordered_text(&config.to_ini());
    assert!(
        written.contains("render_html=false"),
        "the key is written under the heading the table names it in:\n{written}"
    );

    // And the switch round-trips as the switch itself rather than as a value read one
    // way and written the other.
    for setting in [true, false] {
        let config = AppConfig {
            render_html: setting,
            ..Default::default()
        };

        let written = ordered_text(&config.to_ini());
        let mut read_back = Ini::new();
        read_back
            .read(written)
            .expect("a file this app wrote is one it can read");

        let mut config = AppConfig::default();
        config.apply_ini(&read_back);
        assert_eq!(
            config.render_html, setting,
            "`render_html={setting}` read back"
        );
    }
}

/// A file holding the built-in list of an earlier version has never been edited,
/// so the entries added to the list since then are put into it. Without this, a
/// format added to a built-in list would preview on a fresh installation only.
#[test]
fn a_list_holding_the_apps_own_older_image_entries_takes_the_ones_added_to_it() {
    let mut ini = Ini::new();
    ini.set(
        lists::IMAGE.section,
        "extensions",
        Some(lists::IMAGE_EXTENSIONS_BEFORE_SVG.to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(
        config.image_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_IMAGE_EXTENSIONS),
        "the list the app shipped before is read as the list it ships now"
    );
}

/// The `[video]` list is two lists now — the names Windows' own codecs read and the names only
/// FFmpeg's player does — and a file written before the split holds the two of them as one
/// list. A list holding exactly the entries this app shipped then is this app's own rather
/// than an edit somebody made, so it is split as the file is read: the media engine is asked
/// about the half it can read, and the rest of the old list is written down beside it.
#[test]
fn a_video_list_from_before_the_split_is_read_as_the_two_lists_of_now() {
    let mut ini = written_file_before_this_build(&[lists::FFMPEG.section]);
    ini.set(
        lists::VIDEO.section,
        "extensions",
        Some(lists::VIDEO_EXTENSIONS_BEFORE_THE_SPLIT.to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(
        config.video_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_VIDEO_EXTENSIONS),
        "the engine is asked about the names it can read"
    );
    assert_eq!(
        config.ffmpeg_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_FFMPEG_EXTENSIONS),
        "and the rest of the old list is the player's"
    );
    assert!(
        config.differs(&ini),
        "a file written before the split is one to write again"
    );
}

/// A list somebody edited is their own, and the split leaves it alone: a `[video]` list with a
/// name added to it is not the list this app shipped, so it is kept exactly as it is — the
/// repair is for lists this app wrote — and what the section beside it is given is the list of
/// now, since a section the file does not have is one the app writes.
#[test]
fn a_video_list_of_the_users_own_is_left_as_it_is() {
    let mut ini = written_file_before_this_build(&[lists::FFMPEG.section]);
    let edited = format!("{},film-of-mine", lists::VIDEO_EXTENSIONS_BEFORE_THE_SPLIT);
    ini.set(lists::VIDEO.section, "extensions", Some(edited.clone()));

    let config = read_file(&mut ini);

    assert_eq!(
        config.video_extensions,
        lists::sanitize_extension_list(&edited),
        "an edit is read as the edit it is"
    );
    assert_eq!(
        config.ffmpeg_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_FFMPEG_EXTENSIONS),
        "and the player's list is the one this build writes"
    );
}

/// And the same for the `[peazip]` list, which is new to this table with the backends beside
/// the console archiver: a file holding the entries this app shipped before those tools were
/// driven has never been edited, so it is brought up to the list of now — which is how an
/// installation that already exists is given the names the archiver's own table never
/// declared, and is what keeps a `.arc` or a `.br` from previewing on a fresh installation
/// only.
#[test]
fn a_list_holding_the_apps_own_peazip_entries_takes_the_backends_added_to_them() {
    let mut ini = Ini::new();
    ini.set(
        lists::PEAZIP.section,
        "extensions",
        Some(lists::PEAZIP_EXTENSIONS_BEFORE_THE_BACKENDS.to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(
        config.peazip_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_PEAZIP_EXTENSIONS),
        "the list the app shipped before is read as the list it ships now"
    );
    for name in ["arc", "zpaq", "br", "bcm", "lpaq8"] {
        assert!(
            config.peazip_extensions.iter().any(|entry| entry == name),
            "`{name}` is one of the names added with the backends"
        );
    }
}

/// And the same for the lists a name moved between when the book kind grew one of its own: `cbz`
/// was the `[archive]` list's until a comic became a book, `lit` was the `[peazip]` list's until
/// the ebook engine was the one asked about it, and `chm` has been both — the ebook engine drew a
/// page for one for a build and the archiver has it again, because a help file is not worth two
/// seconds of waiting. A file holding any of those as this app shipped it has never been edited —
/// nobody types these — so all of them are brought up to the lists of now.
///
/// It is the case worth having a test for rather than the repairs separately: a file written
/// between two changes holds one list of each pair, and `chm` can be in either of two of them.
#[test]
fn a_list_holding_the_apps_own_entries_loses_the_names_that_became_books() {
    let mut ini = Ini::new();
    ini.set(
        lists::ARCHIVE.section,
        "extensions",
        Some(lists::ARCHIVE_EXTENSIONS_BEFORE_THE_COMICS.to_string()),
    );
    ini.set(
        lists::PEAZIP.section,
        "extensions",
        Some(lists::PEAZIP_EXTENSIONS_BEFORE_THE_HELP_FILE.to_string()),
    );
    ini.set(
        lists::CALIBRE.section,
        "extensions",
        Some(lists::CALIBRE_EXTENSIONS_WITH_THE_HELP_FILE.to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(
        config.archive_extensions,
        crate::formats::text_formats::sanitize_archive_extension_list(
            lists::DEFAULT_ARCHIVE_EXTENSIONS
        ),
        "a comic is not an archive any more"
    );
    assert_eq!(
        config.peazip_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_PEAZIP_EXTENSIONS),
        "and the help file the engine had taken is the archiver's again"
    );
    assert_eq!(
        config.calibre_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_CALIBRE_EXTENSIONS),
        "and the engine's list is without it, with the Microsoft Reader book it kept"
    );

    for name in ["cbz", "lit"] {
        assert!(
            !config.archive_extensions.iter().any(|entry| entry == name)
                && !config.peazip_extensions.iter().any(|entry| entry == name),
            "`{name}` is a book now, so no listing list carries it"
        );
    }

    assert!(
        config.peazip_extensions.iter().any(|entry| entry == "chm"),
        "and the help file the engine had taken for a build is the archiver's again"
    );
    assert!(
        !config.calibre_extensions.iter().any(|entry| entry == "chm"),
        "so a hover on one is a listing rather than a two-second wait"
    );
    assert!(
        config.calibre_extensions.iter().any(|entry| entry == "lit"),
        "and the Microsoft Reader book is the engine's, which is where it stays"
    );

    // And the list the comics went to is a section the file has never held, so the built-in
    // entries come back with the key rather than the names being lost on the way over.
    assert_eq!(
        config.ebook_extensions,
        crate::formats::text_formats::sanitize_extension_list(lists::DEFAULT_EBOOK_EXTENSIONS),
        "a file with no `[ebook]` section is given the built-in list"
    );
    assert!(
        config.differs(&ini),
        "and the file, which holds none of that, is one to write"
    );
}

/// Names taken out of a built-in list leave the files already written with them: the
/// list the app shipped with those names is this app's own — nobody typed it — so it is
/// read as the list of now, and a name the engine cannot read stops being asked about
/// on an installation that has been running since before they were taken out.
#[test]
fn a_list_holding_the_names_the_engine_cannot_read_loses_them() {
    let mut ini = Ini::new();
    ini.set(
        lists::LIBRE.section,
        "extensions",
        Some(lists::LIBRE_EXTENSIONS_WITH_THE_NAMES_THE_ENGINE_CANNOT_READ.to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(
        config.libre_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_LIBRE_EXTENSIONS),
        "the list the app shipped before is read as the list it ships now"
    );
    for name in ["swf", "epub", "qxp", "pm3", "vssm", "uof"] {
        assert!(
            !config.libre_extensions.iter().any(|entry| entry == name),
            "`{name}` is one of the names that left it"
        );
    }
}

/// A name this app no longer writes is a key it does not write, and that is all it is: the
/// setting the line once named is not read from it — this app has no way to know what the
/// name meant — and the line goes with every other line that is not the app's, which is what
/// keeps a file from carrying the history of the names it has been written under.
#[test]
fn a_name_the_app_no_longer_writes_is_dropped_with_the_line_it_is_on() {
    let mut ini = written_file();
    ini.set(CONFIG_SECTION, "svg_scale", Some("75".to_string()));

    assert!(
        AppConfig::default().differs(&ini),
        "a line the app does not write is one the file is written again for"
    );

    let mut config = AppConfig::default();
    config.apply_ini(&ini);

    assert_eq!(
        config.vector_scale, DEFAULT_VECTOR_SCALE,
        "and the setting it once named is where a fresh installation starts"
    );
}

/// An old file, read as the app reads one: the lists it holds are brought up to the ones this
/// build ships, which is what gives an installation that already exists the formats added
/// since — while the settings it holds under names this app no longer writes are not read at
/// all, so those go back to their defaults and the lines are dropped with the write.
#[test]
fn an_old_file_keeps_its_lists_and_loses_the_names_this_app_no_longer_writes() {
    let mut ini = written_file();
    ini.set(CONFIG_SECTION, "off_trigger_key", Some("ctrl".to_string()));
    ini.set(CONFIG_SECTION, "avoid_filename", Some("true".to_string()));
    ini.set(CONFIG_SECTION, "svg_scale", Some("75".to_string()));
    ini.set(
        lists::IMAGE.section,
        "extensions",
        Some(lists::IMAGE_EXTENSIONS_BEFORE_DDS.to_string()),
    );
    ini.set(
        lists::DESIGN.section,
        "extensions",
        Some(lists::DESIGN_EXTENSIONS_BEFORE_AI.to_string()),
    );

    assert!(
        lists::repair_older_lists(&mut ini),
        "the lists are what the repair has to do with a file like this"
    );

    let mut config = AppConfig::default();
    config.apply_ini(&ini);

    for extension in ["dds", "avif", "heic", "heif", "jxl"] {
        assert!(
            config.image_extensions.contains(&extension.to_string()),
            "`{extension}` is in the list of now"
        );
    }
    assert!(config.design_extensions.contains(&"ai".to_string()));

    assert_eq!(
        config.trigger_key, "alt",
        "a name the app does not write says nothing, however clear it looks"
    );
    assert_eq!(config.avoid_mode, DEFAULT_AVOID_MODE);
    assert_eq!(config.vector_scale, DEFAULT_VECTOR_SCALE);

    assert!(
        config.differs(&ini),
        "and the file, holding lines the app does not write, is one to write again"
    );
}

/// The design list's own version of the same: a file holding a list the app shipped
/// before two more names were added to it — `ai`, and then `cdr` and `procreate` — is
/// the app's own older list, so those entries reach an installation that already exists
/// rather than a fresh one only — and a list anyone has edited is kept exactly as it is.
#[test]
fn a_list_holding_the_apps_own_design_entries_takes_the_ones_added_to_them() {
    for shipped in [
        lists::DESIGN_EXTENSIONS_BEFORE_AI,
        lists::DESIGN_EXTENSIONS_BEFORE_CDR_AND_PROCREATE,
    ] {
        let mut ini = Ini::new();
        ini.set(
            lists::DESIGN.section,
            "extensions",
            Some(shipped.to_string()),
        );

        let config = read_file(&mut ini);

        assert_eq!(
            config.design_extensions,
            lists::sanitize_extension_list(lists::DEFAULT_DESIGN_EXTENSIONS),
            "the list the app shipped before (`{shipped}`) is read as the list it ships now"
        );
    }

    let edited = format!("dng,{}", lists::DESIGN_EXTENSIONS_BEFORE_AI);
    let mut ini = Ini::new();
    ini.set(lists::DESIGN.section, "extensions", Some(edited.clone()));

    let config = read_file(&mut ini);

    assert_eq!(
        config.design_extensions,
        lists::sanitize_extension_list(&edited),
        "a list with an entry of its own is the user's and is kept as written"
    );
}

/// The same, one list's worth of entries later: a file holding the list the app
/// shipped before the formats Windows has a codec for were added to it is the
/// app's own older list, so those four reach an installation that already exists
/// rather than a fresh one only.
#[test]
fn a_list_holding_the_apps_own_image_entries_takes_the_codec_formats_added_to_them() {
    let mut ini = Ini::new();
    ini.set(
        lists::IMAGE.section,
        "extensions",
        Some(lists::IMAGE_EXTENSIONS_BEFORE_CODEC_FORMATS.to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(
        config.image_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_IMAGE_EXTENSIONS),
        "the list the app shipped before is read as the list it ships now"
    );
    for extension in ["avif", "heic", "heif", "jxl"] {
        assert!(
            config.image_extensions.contains(&extension.to_string()),
            "`{extension}` was added to the built-in list"
        );
    }
}

/// The same list one move later, the other way round: a file holding the list the app
/// shipped while `svg` and `svgz` were entries of it is the app's own older list, so
/// those two entries are given up — the kind they belong to names them now — rather
/// than left in a list of pictures they were never pictures of.
#[test]
fn a_list_holding_the_apps_own_image_entries_gives_the_documents_up() {
    let mut ini = Ini::new();
    ini.set(
        lists::IMAGE.section,
        "extensions",
        Some(lists::IMAGE_EXTENSIONS_WITH_SVG.to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(
        config.image_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_IMAGE_EXTENSIONS),
        "the list the app shipped before is read as the list it ships now"
    );
    assert!(!config.image_extensions.contains(&"svg".to_string()));
    assert!(!config.image_extensions.contains(&"svgz".to_string()));
}

/// And the vector list's own version of it: a file holding the list the app shipped
/// before the documents were added to it takes them, so an installation that already
/// exists keeps previewing an `svg` when the image list gives it up.
#[test]
fn a_list_holding_the_apps_own_vector_entries_takes_the_documents_added_to_them() {
    let mut ini = Ini::new();
    ini.set(
        lists::VECTOR.section,
        "extensions",
        Some(lists::VECTOR_EXTENSIONS_BEFORE_SVG.to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(
        config.vector_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_VECTOR_EXTENSIONS),
        "the list the app shipped before is read as the list it ships now"
    );
    for extension in ["svg", "svgz"] {
        assert!(
            config.vector_extensions.contains(&extension.to_string()),
            "`{extension}` is a drawing and belongs to the vector list"
        );
    }
}

/// And the spellings of an encapsulated PostScript file, which were added to that list
/// after the documents were: a file holding the list as it stood before them takes them
/// too, so a `.epsf` or an `.ept` is a drawing an installation already exists previews.
#[test]
fn a_list_holding_the_apps_own_vector_entries_takes_the_eps_spellings_added_to_them() {
    let mut ini = Ini::new();
    ini.set(
        lists::VECTOR.section,
        "extensions",
        Some(lists::VECTOR_EXTENSIONS_BEFORE_THE_EPS_SPELLINGS.to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(
        config.vector_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_VECTOR_EXTENSIONS)
    );
    for extension in ["epsf", "epi", "ept", "ept2", "ept3"] {
        assert!(
            config.vector_extensions.contains(&extension.to_string()),
            "`{extension}` is another spelling of the same drawing"
        );
    }
}

/// And the image list's own version: a file holding the list the app shipped before the
/// codec's AVC still was added to it takes the name, so a `.avci` is a picture an
/// installation that already exists previews.
#[test]
fn a_list_holding_the_apps_own_image_entries_takes_the_avc_still_added_to_them() {
    let mut ini = Ini::new();
    ini.set(
        lists::IMAGE.section,
        "extensions",
        Some(lists::IMAGE_EXTENSIONS_BEFORE_AVCI.to_string()),
    );

    let config = read_file(&mut ini);

    assert_eq!(
        config.image_extensions,
        lists::sanitize_extension_list(lists::DEFAULT_IMAGE_EXTENSIONS)
    );
    assert!(config.image_extensions.contains(&"avci".to_string()));
}

/// The same list with one entry of the user's own in it is the user's list, not
/// the app's: the entries added since are left out of it.
#[test]
fn an_image_list_anyone_has_edited_is_kept_as_written() {
    let written = format!("dng,{}", lists::IMAGE_EXTENSIONS_BEFORE_SVG);

    let mut ini = Ini::new();
    ini.set(lists::IMAGE.section, "extensions", Some(written.clone()));

    let config = read_file(&mut ini);

    assert_eq!(
        config.image_extensions,
        lists::sanitize_extension_list(&written),
        "what the file says is what the list is"
    );
    assert!(!config.image_extensions.contains(&"svg".to_string()));
}

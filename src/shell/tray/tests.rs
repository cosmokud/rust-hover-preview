use super::*;

/// The order the submenu is built in is the order it is read in, and it is
/// built top down: the engine that is never let go is the topmost item and the
/// one let go at once is the bottom one, which is what the choices are listed
/// in.
#[test]
fn engine_idle_is_offered_longest_first() {
    assert_eq!(
        ENGINE_IDLE_CHOICES.map(|idle| engine_idle_label(idle, DEFAULT_OFFICE_ENGINE_IDLE_SECS)),
        [
            "Indefinitely".to_string(),
            "1 hour".to_string(),
            "30 minutes".to_string(),
            "10 minutes (Default)".to_string(),
            "5 minutes".to_string(),
            "1 minute".to_string(),
            "0 seconds".to_string(),
        ]
    );
}

/// The halves of the `Background` submenu list the backdrops their settings hold, in the
/// order the ids are handed out in, and each id resolves back to the backdrop its item was
/// listed for — which is what makes a click select what it named. Every half marks the
/// backdrop its own setting starts at, which is not the same one for every half, and not
/// every half lists all four: the texture's offers two of them and a page's three.
#[test]
fn every_offered_background_is_one_the_setting_keeps() {
    assert_eq!(
        BACKGROUND_CHOICES.map(|choice| background_label(choice, DEFAULT_IMAGE_BACKGROUND)),
        [
            "Transparent".to_string(),
            "Black".to_string(),
            "White".to_string(),
            "Checkerboard (Default)".to_string(),
        ]
    );

    // The mark is the setting's own answer: one backdrop per half carries it, and which
    // one it is is read from the constant that half's setting starts at.
    for (choices, default) in [
        (&BACKGROUND_CHOICES[..], DEFAULT_IMAGE_BACKGROUND),
        (&BACKGROUND_CHOICES[..], DEFAULT_VECTOR_BACKGROUND),
        (&BACKGROUND_CHOICES[..], DEFAULT_FONT_BACKGROUND),
        (&BACKGROUND_CHOICES[..], DEFAULT_DESIGN_BACKGROUND),
        (&DDS_BACKGROUND_CHOICES[..], DEFAULT_DDS_BACKGROUND),
        (&HTML_BACKGROUND_CHOICES[..], DEFAULT_HTML_BACKGROUND),
    ] {
        let marked: Vec<TransparentBackground> = choices
            .iter()
            .copied()
            .filter(|choice| background_label(*choice, default).ends_with(" (Default)"))
            .collect();

        assert_eq!(
            marked,
            [default],
            "the backdrop a half starts at is the one it marks"
        );
    }

    for (index, background) in BACKGROUND_CHOICES.iter().enumerate() {
        assert_eq!(background_at(index as u16), Some(*background));
    }

    assert_eq!(
        background_at(BACKGROUND_CHOICES.len() as u16),
        None,
        "an id past the last item is not one the menu offered"
    );

    // The texture's half is a range and a table of its own, two backdrops wide.
    for (index, background) in DDS_BACKGROUND_CHOICES.iter().enumerate() {
        assert_eq!(dds_background_at(index as u16), Some(*background));
    }

    assert_eq!(
        dds_background_at(DDS_BACKGROUND_CHOICES.len() as u16),
        None,
        "a backdrop the texture's half does not offer is not one of its items"
    );

    // And a page's half is a range and a table of its own too, three backdrops wide.
    for (index, background) in HTML_BACKGROUND_CHOICES.iter().enumerate() {
        assert_eq!(html_background_at(index as u16), Some(*background));
    }

    assert_eq!(
        html_background_at(HTML_BACKGROUND_CHOICES.len() as u16),
        None,
        "a backdrop a page's half does not offer is not one of its items"
    );
}

/// An item of one half of the submenu is never an item of another, whatever it was
/// listed at — and never an id something else in the menu hands out either, which is
/// the failure this range was moved for: a backdrop of a specimen and the text
/// preview's `Full Mode` item were one id, and the backdrop was read first.
#[test]
fn the_six_halves_of_the_background_submenu_carry_different_ids() {
    // Each half is as wide as the choices it offers, which is two for the texture's half,
    // three for a page's and four for the rest: a range that were wider than the items in
    // it would take an id from the half beside it.
    let halves = [
        (
            ID_TRAY_IMAGE_BACKGROUND_BASE,
            BACKGROUND_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_FONT_BACKGROUND_BASE,
            BACKGROUND_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_DDS_BACKGROUND_BASE,
            DDS_BACKGROUND_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_DESIGN_BACKGROUND_BASE,
            BACKGROUND_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_VECTOR_BACKGROUND_BASE,
            BACKGROUND_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_HTML_BACKGROUND_BASE,
            HTML_BACKGROUND_CHOICES.len() as u16,
        ),
    ];

    for (index, half) in halves.iter().enumerate() {
        for other in &halves[index + 1..] {
            let ours = half.0..half.0 + half.1;
            let theirs = other.0..other.0 + other.1;

            assert!(
                !ours.contains(&other.0) && !theirs.contains(&half.0),
                "the ranges {ours:?} and {theirs:?} overlap"
            );
        }
    }

    // The ids around them, of the menus that grew up beside the `Background` one. A
    // page's half is in the stretch the `Volume` submenu's two halves gave up, so the
    // ids those used to be are among the ones that must stay away from it.
    for elsewhere in [
        ID_TRAY_THEME_LIGHT,
        ID_TRAY_MARKDOWN_RENDERED,
        ID_TRAY_OPEN_CONFIG,
        ID_TRAY_SCALE_BASE,
        ID_TRAY_VIDEO_SCALE_BASE,
        ID_TRAY_TYPE_IMAGES,
        ID_TRAY_TRIGGER_ENABLED,
        ID_TRAY_RENDER_HTML,
    ] {
        for half in halves {
            let ours = half.0..half.0 + half.1;

            assert!(
                !ours.contains(&elsewhere),
                "the range {ours:?} contains {elsewhere}, which is another item's id"
            );
        }
    }
}

/// The rows of the `Pin Mode` submenu are ids of their own, and neither is the row of the
/// pin itself: the three hang from one submenu and are clicked in the same place, so a
/// collision here is a click that switches previews off where it meant to leave a pin
/// following the listing.
///
/// What each switch starts at is the answer the tray draws its ticks from, so it is checked
/// with them: a row whose setting starts one way and whose id is read another is a switch
/// that appears to do nothing for one click.
#[test]
fn the_pin_mode_rows_are_ids_of_their_own() {
    for (row, other) in [
        (ID_TRAY_PIN, ID_TRAY_PIN_UPDATE),
        (ID_TRAY_PIN, ID_TRAY_PIN_UPDATE_HOVER),
        (ID_TRAY_PIN_UPDATE, ID_TRAY_PIN_UPDATE_HOVER),
        (ID_TRAY_PIN, ID_TRAY_ENABLE),
        (ID_TRAY_PIN, ID_TRAY_TRIGGER_ENABLED),
        (ID_TRAY_PIN_UPDATE, ID_TRAY_TRIGGER_ENABLED),
        (ID_TRAY_PIN_UPDATE_HOVER, ID_TRAY_TRIGGER_ENABLED),
        (ID_TRAY_PIN_NAV_ALL, ID_TRAY_PIN_NAV_CATEGORY),
        (ID_TRAY_PIN_NAV_ALL, ID_TRAY_PIN),
        (ID_TRAY_PIN_NAV_ALL, ID_TRAY_PIN_UPDATE),
        (ID_TRAY_PIN_NAV_ALL, ID_TRAY_PIN_UPDATE_HOVER),
        (ID_TRAY_PIN_NAV_CATEGORY, ID_TRAY_PIN),
        (ID_TRAY_PIN_NAV_CATEGORY, ID_TRAY_PIN_UPDATE),
        (ID_TRAY_PIN_NAV_CATEGORY, ID_TRAY_PIN_UPDATE_HOVER),
        (ID_TRAY_PIN_NAV_ALL, ID_TRAY_PIN_PAUSE_AUDIO),
        (ID_TRAY_PIN_NAV_ALL, ID_TRAY_PIN_PAUSE_VIDEO),
    ] {
        assert_ne!(row, other, "two rows of the menu share the id {row}");
    }

    // And outside every range a click is read against before this row is, so a walk's
    // two answers are never answered as an item of a submenu of their own.
    for (base, len) in [
        (
            ID_TRAY_LIBREOFFICE_IDLE_BASE,
            ENGINE_IDLE_CHOICES.len() as u16,
        ),
        (ID_TRAY_SCALE_BASE, BITMAP_SCALE_CHOICES.len() as u16),
    ] {
        for row in [ID_TRAY_PIN_NAV_ALL, ID_TRAY_PIN_NAV_CATEGORY] {
            assert!(
                !(base..base + len).contains(&row),
                "the {row} row is inside the {base} range, so a click on it is read as that \
                 submenu's"
            );
        }
    }

    let defaults = crate::config::config::AppConfig::default();

    assert!(
        defaults.pin_update_enabled,
        "a pin follows what the user picks unless it is switched off"
    );
    assert!(
        !defaults.pin_update_on_hover,
        "the pointer's own hover is not one of the ways until it is asked for"
    );
    assert_eq!(
        defaults.pin_nav_file_types, DEFAULT_PIN_NAV_FILE_TYPES,
        "a pin walks the whole folder until it is told otherwise"
    );
}

/// The trigger key's second row is a switch of its own beside the one that watches the key,
/// and it starts off: a pin is a window the user put there, and the key that stops hovers is
/// not what takes it down until it is asked for.
#[test]
fn the_trigger_keys_pin_row_starts_off_and_is_an_id_of_its_own() {
    let defaults = crate::config::config::AppConfig::default();

    assert_eq!(
        defaults.trigger_key_affect_pin_mode, DEFAULT_TRIGGER_KEY_AFFECT_PIN_MODE,
        "a fresh pin is not the trigger key's to take down"
    );

    for row in [
        ID_TRAY_TRIGGER_ENABLED,
        ID_TRAY_TRIGGER_DISABLE,
        ID_TRAY_TRIGGER_ENABLE,
        ID_TRAY_PIN,
    ] {
        assert_ne!(
            ID_TRAY_TRIGGER_AFFECT_PIN, row,
            "the trigger key's own rows share the id {row}"
        );
    }
}

/// The two halves of the `Volume` submenu carry a range apiece, and the levels they offer
/// are the table's own: silence and the decades above it, each louder than the one before,
/// which the menu lists the other way round — the whole of the scale at the top.
///
/// It is a test about ids for the reason the timing one below is — the two halves share a
/// table, so a range that ran into the other would hand a click to the wrong setting — and
/// about the table for the reason the font sizes' own test is: a level no item offers is a
/// setting a user can reach only by editing the file.
#[test]
fn the_two_volume_submenus_carry_a_range_apiece() {
    let bases = [ID_TRAY_VIDEO_VOLUME_BASE, ID_TRAY_AUDIO_VOLUME_BASE];
    let width = VOLUME_CHOICES.len() as u16;

    assert_eq!(VOLUME_CHOICES[0], 0, "the topmost item is silence");
    assert_eq!(
        *VOLUME_CHOICES.last().expect("a last level"),
        100,
        "the bottom one is the whole of it"
    );
    assert!(
        VOLUME_CHOICES.windows(2).all(|pair| pair[0] < pair[1]),
        "every level is louder than the one above it: {VOLUME_CHOICES:?}"
    );
    for default in [DEFAULT_VIDEO_VOLUME, DEFAULT_AUDIO_VOLUME] {
        assert!(
            VOLUME_CHOICES.contains(&default),
            "{default}% is marked as a default and is no item of the menu"
        );
    }

    for (index, base) in bases.iter().enumerate() {
        for above in &bases[..index] {
            assert!(
                above + width <= *base,
                "the range at {base} overlaps the one at {above}"
            );
        }
    }

    // The gate for sounds is a command of its own, in the slack the font sizes leave
    // rather than in either range: a level is not a switch, and neither is a switch a
    // level — a collision here is a click that turns previews off where it meant to
    // change a volume.
    for base in bases {
        assert!(
            !(base..base + width).contains(&ID_TRAY_TYPE_AUDIO),
            "the sound gate's id is inside the range at {base}"
        );
    }

    // And the third item of the same submenu — where a sound starts — is a range of its
    // own as well: it shares the menu with both halves rather than a table, and a click on
    // one of its ways is never a click on a level of either (see the test below).
    let seek = ID_TRAY_AUDIO_SEEK_BASE..ID_TRAY_AUDIO_SEEK_BASE + AUDIO_SEEK_CHOICES.len() as u16;
    for base in bases {
        let range = base..base + width;
        assert!(
            !range.contains(&seek.start) && !seek.contains(&range.start),
            "the ranges {range:?} and {seek:?} overlap"
        );
    }
}

/// The `Volume → Audio Seek` submenu lists a way of starting a sound for every way the
/// setting has, in the order the ids are handed out in, with the one the setting starts at
/// marked as the default — and each id resolves back to the way its item was listed for,
/// which is what makes a click start a sound where it named.
#[test]
fn every_offered_way_of_starting_a_sound_is_one_the_setting_keeps() {
    assert_eq!(
        AUDIO_SEEK_CHOICES.map(audio_seek_label),
        [
            "Remember (Default)".to_string(),
            "From the Start".to_string(),
            "From the Middle".to_string(),
            "Random".to_string()
        ]
    );

    let marked: Vec<AudioSeek> = AUDIO_SEEK_CHOICES
        .iter()
        .copied()
        .filter(|seek| audio_seek_label(*seek).ends_with(" (Default)"))
        .collect();

    assert_eq!(
        marked,
        [DEFAULT_AUDIO_SEEK],
        "the way the setting starts at is the way the menu marks"
    );

    for (index, seek) in AUDIO_SEEK_CHOICES.iter().enumerate() {
        assert_eq!(audio_seek_at(index as u16), Some(*seek));
    }

    assert_eq!(
        audio_seek_at(AUDIO_SEEK_CHOICES.len() as u16),
        None,
        "an id past the last item is not one the menu offered"
    );

    // The ways the submenu offers are the ways the setting holds, and each of them is
    // named: a way the menu has no words for is one a user cannot pick, and a way the
    // setting has that the menu does not list is one they cannot reach but by editing
    // `config.ini` — which is the arrangement every other value menu here keeps to.
    assert_eq!(
        AUDIO_SEEK_CHOICES.len(),
        4,
        "every way of starting a sound is offered: {AUDIO_SEEK_CHOICES:?}"
    );
    assert!(
        AUDIO_SEEK_CHOICES.contains(&DEFAULT_AUDIO_SEEK),
        "the way the setting starts at is one of the items"
    );
}

/// The three `Timing` submenus list the same delays, one range apiece, and a range
/// that were wider than the items in it would take an id from the submenu below it —
/// which is a click selecting a delay that was never listed, for a setting nobody
/// asked for.
#[test]
fn the_three_timing_submenus_carry_a_range_apiece() {
    let bases = [
        ID_TRAY_DELAY_BASE,
        ID_TRAY_REHOVER_DELAY_BASE,
        ID_TRAY_SETTLING_DELAY_BASE,
    ];
    let width = TIMING_DELAY_CHOICES_MS.len() as u16;

    // The table is the steps themselves: no wait at all at the top, a whole second at
    // the bottom, and every delay larger than the one above it.
    assert_eq!(TIMING_DELAY_CHOICES_MS[0], 0, "the topmost item is no wait");
    assert_eq!(
        *TIMING_DELAY_CHOICES_MS.last().expect("a last delay"),
        1000,
        "the bottom one is a whole second"
    );
    assert!(
        TIMING_DELAY_CHOICES_MS
            .windows(2)
            .all(|pair| pair[0] < pair[1]),
        "every delay is larger than the one above it: {TIMING_DELAY_CHOICES_MS:?}"
    );

    // Every value a default mark is read from is one of the items: a default the menu
    // does not offer would leave the setting starting at nothing marked.
    for default_ms in [
        DEFAULT_HOVER_DELAY_MS,
        DEFAULT_SAME_FILE_REHOVER_DELAY_MS,
        DEFAULT_SETTLING_DELAY_MS,
    ] {
        assert!(
            TIMING_DELAY_CHOICES_MS.contains(&default_ms),
            "{default_ms} ms is marked as a default and is no item of the menu"
        );
    }

    for (index, base) in bases.iter().enumerate() {
        for above in &bases[..index] {
            assert!(
                above + width <= *base,
                "the range at {base} overlaps the one at {above}"
            );
        }
    }

    for (position, delay_ms) in TIMING_DELAY_CHOICES_MS.iter().enumerate() {
        assert_eq!(
            timing_delay_at(position as u16),
            Some(*delay_ms),
            "an item resolves back to the delay it was listed for"
        );
    }

    assert_eq!(
        timing_delay_at(width),
        None,
        "an id past the last item is not one the menu offered"
    );
}

/// The `Avoid` submenu lists every way the setting can be in, in the order the ids
/// are handed out in, and each id resolves back to the way its item was listed for
/// — which is what makes a click select what it named.
#[test]
fn every_offered_avoid_mode_is_one_the_setting_keeps() {
    assert_eq!(
        AVOID_CHOICES.map(avoid_label),
        [
            "Avoid Nothing".to_string(),
            "Avoid Filename (Default)".to_string(),
            "Avoid Filename Column".to_string(),
            "Avoid Details".to_string()
        ]
    );

    let marked: Vec<AvoidMode> = AVOID_CHOICES
        .iter()
        .copied()
        .filter(|mode| avoid_label(*mode).ends_with(" (Default)"))
        .collect();

    assert_eq!(
        marked,
        [DEFAULT_AVOID_MODE],
        "the way the setting starts at is the way the menu marks"
    );

    for (index, mode) in AVOID_CHOICES.iter().enumerate() {
        assert_eq!(avoid_mode_at(index as u16), Some(*mode));
    }

    assert_eq!(
        avoid_mode_at(AVOID_CHOICES.len() as u16),
        None,
        "an id past the last item is not one the menu offered"
    );
}

/// The `… Scaling` submenus are one range each, and the `Avoid` items sit in
/// the slack between them: a click on a share of the display is never read as a way
/// of avoiding the item a preview is about, and the other way round.
#[test]
fn the_avoid_submenu_carries_ids_of_its_own() {
    let avoid = ID_TRAY_AVOID_BASE..ID_TRAY_AVOID_BASE + AVOID_CHOICES.len() as u16;
    let document_scales = [
        ID_TRAY_VECTOR_SCALE_BASE,
        ID_TRAY_TEXT_SCALE_BASE,
        ID_TRAY_EBOOK_SCALE_BASE,
        ID_TRAY_DOCUMENT_SCALE_BASE,
        ID_TRAY_FONT_SCALE_BASE,
        ID_TRAY_DESIGN_SCALE_BASE,
    ]
    .map(|base| base..base + DOCUMENT_SCALE_CHOICES.len() as u16);

    for range in &document_scales {
        assert!(
            !range.contains(&avoid.start) && !avoid.contains(&range.start),
            "the ranges {range:?} and {avoid:?} overlap"
        );
    }
    assert!(
        avoid.start > ID_TRAY_POSITION_BEST,
        "the avoid items are listed after the position choices"
    );

    // One range per submenu as well: a click on one scale is never read as a click on
    // the submenu beside it.
    for (index, range) in document_scales.iter().enumerate() {
        for other in document_scales.iter().skip(index + 1) {
            assert!(
                !range.contains(&other.start) && !other.contains(&range.start),
                "the ranges {range:?} and {other:?} overlap"
            );
        }
    }
}

/// The `Render HTML` row is an id of its own, and one of the slack rather than of a
/// range: it sits where the text preview's `Full Mode` item did, four ids above the
/// Markdown rows this submenu also holds, and a click on either of them is a click on
/// the setting it names.
///
/// It is also read the way the menu draws it — from the setting on the configuration
/// rather than from a value remembered between the click and the row — so the answer the
/// row shows and the answer the click gives are the same one.
#[test]
fn the_render_html_row_is_an_id_of_its_own() {
    for other in [
        ID_TRAY_THEME_LIGHT,
        ID_TRAY_THEME_DARK,
        ID_TRAY_MARKDOWN_RENDERED,
        ID_TRAY_MARKDOWN_SOURCE,
        ID_TRAY_FONT_100,
        ID_TRAY_TYPE_IMAGES,
        ID_TRAY_TYPE_TEXT,
        ID_TRAY_VECTOR_BACKGROUND_BASE,
        ID_TRAY_IMAGE_BACKGROUND_BASE,
    ] {
        assert_ne!(
            ID_TRAY_RENDER_HTML, other,
            "the text preview's rows share the id {other}"
        );
    }

    const _: () = assert!(
        ID_TRAY_RENDER_HTML > ID_TRAY_MARKDOWN_SOURCE,
        "the row is listed after the Markdown popup it hangs beside"
    );

    // And outside every range the click reads before the row is: a row that fell in one
    // of them would be answered as an item of that submenu and never reach the switch
    // (see the `cmd if` arms above `ID_TRAY_RENDER_HTML` in `tray_window_proc`).
    for (base, len) in [
        (
            ID_TRAY_IMAGE_BACKGROUND_BASE,
            BACKGROUND_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_VECTOR_BACKGROUND_BASE,
            BACKGROUND_CHOICES.len() as u16,
        ),
        (ID_TRAY_SCALE_BASE, BITMAP_SCALE_CHOICES.len() as u16),
        (ID_TRAY_ENGINE_IDLE_BASE, ENGINE_IDLE_CHOICES.len() as u16),
        (
            ID_TRAY_HTML_BACKGROUND_BASE,
            HTML_BACKGROUND_CHOICES.len() as u16,
        ),
    ] {
        assert!(
            !(base..base + len).contains(&ID_TRAY_RENDER_HTML),
            "the row is inside the {base} range, so a click on it is read as that row's"
        );
    }

    let defaults = crate::config::config::AppConfig::default();
    assert_eq!(
        defaults.render_html, DEFAULT_RENDER_HTML,
        "a fresh install shows a page as its markup"
    );
}

/// Every `… Scaling` submenu offers the whole room a document can be given and then
/// the shares of it, in that order, and each id resolves back to the share its item
/// was listed for — which is what makes a click select what it named. What each
/// submenu marks as the default is the share its own setting starts at.
#[test]
fn every_offered_document_scale_is_one_the_setting_keeps() {
    let drawing_default = DEFAULT_VECTOR_SCALE;

    assert_eq!(
        DOCUMENT_SCALE_CHOICES.map(|scale| document_scale_label(scale, drawing_default)),
        [
            "Fit to Screen (Default)".to_string(),
            "75%".to_string(),
            "50%".to_string(),
            "25%".to_string(),
            "10%".to_string(),
        ]
    );

    for default in [
        drawing_default,
        DEFAULT_EBOOK_SCALE,
        DEFAULT_DOCUMENT_SCALE,
        DEFAULT_FONT_SCALE,
    ] {
        assert_eq!(
            DOCUMENT_SCALE_CHOICES.map(|scale| document_scale_label(scale, default)),
            [
                if default == PreviewScale::FitToScreen {
                    "Fit to Screen (Default)"
                } else {
                    "Fit to Screen"
                },
                "75%",
                if default == PreviewScale::Percent(50) {
                    "50% (Default)"
                } else {
                    "50%"
                },
                "25%",
                "10%",
            ]
            .map(String::from),
            "one share is marked the default at {default:?}"
        );
    }

    for (index, scale) in DOCUMENT_SCALE_CHOICES.iter().enumerate() {
        assert_eq!(document_scale_at(index as u16), Some(*scale));
    }

    assert_eq!(
        document_scale_at(DOCUMENT_SCALE_CHOICES.len() as u16),
        None,
        "an id past the last item is not one the menu offered"
    );
}

/// What the menu writes is what the file reads back: every share it offers is one
/// the setting holds, so a choice made here is still the choice after a restart.
#[test]
fn every_offered_document_scale_round_trips_through_the_file() {
    for scale in DOCUMENT_SCALE_CHOICES {
        let written = scale.as_str();

        assert_eq!(
            PreviewScale::from_str(&written),
            Some(scale),
            "`{written}` read back"
        );
    }
}

/// A share of the display is never read as a cache size, and no cache size is read as
/// another's: the three caches are a range each, and the shares added beside them are ranges
/// of their own.
#[test]
fn the_document_scale_ranges_are_not_another_submenus_range() {
    let caches = [
        ID_TRAY_IMAGE_CACHE_BASE..ID_TRAY_IMAGE_CACHE_BASE + CACHE_SIZE_CHOICES_MB.len() as u16,
        ID_TRAY_DOCUMENT_CACHE_BASE
            ..ID_TRAY_DOCUMENT_CACHE_BASE + CACHE_SIZE_CHOICES_MB.len() as u16,
        ID_TRAY_IMAGE_DISK_CACHE_BASE
            ..ID_TRAY_IMAGE_DISK_CACHE_BASE + CACHE_SIZE_CHOICES_MB.len() as u16,
        ID_TRAY_GENERAL_DISK_CACHE_BASE
            ..ID_TRAY_GENERAL_DISK_CACHE_BASE + CACHE_SIZE_CHOICES_MB.len() as u16,
        ID_TRAY_THEME_CUSTOM_BASE..ID_TRAY_IMAGE_CACHE_BASE,
    ];

    for (index, cache) in caches.iter().enumerate() {
        for other in &caches[index + 1..] {
            assert!(
                !cache.contains(&other.start) && !other.contains(&cache.start),
                "the ranges {cache:?} and {other:?} overlap"
            );
        }
    }

    for base in [
        ID_TRAY_VECTOR_SCALE_BASE,
        ID_TRAY_TEXT_SCALE_BASE,
        ID_TRAY_EBOOK_SCALE_BASE,
        ID_TRAY_DOCUMENT_SCALE_BASE,
        ID_TRAY_FONT_SCALE_BASE,
        ID_TRAY_DESIGN_SCALE_BASE,
    ] {
        let scales = base..base + DOCUMENT_SCALE_CHOICES.len() as u16;

        for other in &caches {
            assert!(
                !other.contains(&base) && !scales.contains(&other.start),
                "the ranges {scales:?} and {other:?} overlap"
            );
        }
    }
}

/// The `Image Scaling`, `Video Scaling` and `Animated Scaling` submenus are one
/// range each, and none of them reaches into another or into the display shares the
/// document scales beside them hand out: a click on a share of a bitmap is never read
/// as a click on another setting's share. Each lists every share the setting can be
/// asked for, in order, and every id resolves back to the share its item was listed
/// for — which is what makes a click select what it named.
#[test]
fn the_bitmap_scaling_submenus_carry_ids_of_their_own() {
    let bitmap_bases = [
        ID_TRAY_SCALE_BASE,
        ID_TRAY_VIDEO_SCALE_BASE,
        ID_TRAY_ANIMATED_SCALE_BASE,
    ];
    let ranges: Vec<(u16, u16)> = bitmap_bases
        .iter()
        .map(|base| (*base, base + BITMAP_SCALE_CHOICES.len() as u16))
        .collect();
    let font_scales = (
        ID_TRAY_FONT_SCALE_BASE,
        ID_TRAY_FONT_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
    );
    let design_scales = (
        ID_TRAY_DESIGN_SCALE_BASE,
        ID_TRAY_DESIGN_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
    );
    let ebook_scales = (
        ID_TRAY_EBOOK_SCALE_BASE,
        ID_TRAY_EBOOK_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
    );
    let document_scales = (
        ID_TRAY_DOCUMENT_SCALE_BASE,
        ID_TRAY_DOCUMENT_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
    );
    let vector_scales = (
        ID_TRAY_VECTOR_SCALE_BASE,
        ID_TRAY_VECTOR_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
    );
    let text_scales = (
        ID_TRAY_TEXT_SCALE_BASE,
        ID_TRAY_TEXT_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
    );

    let overlaps = |ours: (u16, u16), theirs: (u16, u16)| {
        (ours.0 < theirs.1 && theirs.0 < ours.1).then_some((ours, theirs))
    };

    for (index, range) in ranges.iter().enumerate() {
        for other in ranges.iter().skip(index + 1) {
            assert_eq!(
                overlaps(*range, *other),
                None,
                "the ranges {range:?} and {other:?} overlap"
            );
        }

        for document in [
            font_scales,
            design_scales,
            ebook_scales,
            document_scales,
            vector_scales,
            text_scales,
        ] {
            assert_eq!(
                overlaps(*range, document),
                None,
                "the range {range:?} and the display shares {document:?} overlap"
            );
        }
    }

    for base in bitmap_bases {
        for (index, scale) in BITMAP_SCALE_CHOICES.iter().enumerate() {
            assert_eq!(bitmap_scale_at(index as u16), Some(*scale));
        }

        assert_eq!(
            bitmap_scale_at(BITMAP_SCALE_CHOICES.len() as u16),
            None,
            "an id past the last item of the range at {base} is not one it offered"
        );
    }
}

/// The `Image Scaling` and `Video Scaling` submenus offer the shares a bitmap can be
/// drawn at — the share of its own size, rather than the share of the display the
/// document scales beside them are — in one order and with one set of labels: what a
/// share is called does not depend on which of the two is asking, and exactly one
/// label — the share each setting starts at — reads as the default. Every share an
/// item can pick is one the setting keeps, so a choice made here is still the choice
/// after a restart.
#[test]
fn every_offered_bitmap_scale_is_one_the_setting_keeps() {
    assert_eq!(
        BITMAP_SCALE_CHOICES.map(|scale| bitmap_scale_label(scale, DEFAULT_PREVIEW_SCALE)),
        [
            "Fit to Screen".to_string(),
            "400%".to_string(),
            "300%".to_string(),
            "200%".to_string(),
            "150%".to_string(),
            "100% (Default)".to_string(),
            "50%".to_string(),
            "25%".to_string(),
        ]
    );

    for default in [
        DEFAULT_PREVIEW_SCALE,
        DEFAULT_VIDEO_SCALE,
        DEFAULT_ANIMATED_SCALE,
    ] {
        let marked: Vec<String> = BITMAP_SCALE_CHOICES
            .iter()
            .map(|scale| bitmap_scale_label(*scale, default))
            .filter(|label| label.ends_with(" (Default)"))
            .collect();

        assert_eq!(
            marked,
            [bitmap_scale_label(default, default)],
            "one share is the default at {default:?}"
        );
    }

    for scale in BITMAP_SCALE_CHOICES {
        let written = scale.as_str();

        assert_eq!(
            PreviewScale::from_str(&written),
            Some(scale),
            "`{written}` read back"
        );
    }
}

/// The font sizes are listed largest first, `110%` between the `125%` and `100%` it
/// sits between, and every size carries an id of its own — so a click selects the
/// size its item named.
#[test]
fn the_font_sizes_are_listed_largest_first() {
    assert!(
        FONT_SIZE_CHOICES
            .windows(2)
            .all(|pair| pair[0].0 > pair[1].0),
        "every size is smaller than the one above it: {FONT_SIZE_CHOICES:?}"
    );
    assert!(
        FONT_SIZE_CHOICES.contains(&(110, ID_TRAY_FONT_110)),
        "110% is offered"
    );

    for (index, (_, id)) in FONT_SIZE_CHOICES.iter().enumerate() {
        let above = &FONT_SIZE_CHOICES[..index];
        assert!(
            !above.iter().any(|(_, earlier)| earlier == id),
            "an id names one size and one only: {id}"
        );
    }
}

/// Every item a person can pick is one the setting can hold, so what the menu
/// writes is what the menu reads back and marks on the next open.
#[test]
fn every_offered_idle_time_is_one_the_setting_keeps() {
    for idle in ENGINE_IDLE_CHOICES {
        assert_eq!(EngineIdle::from_str(&idle.as_str()), Some(idle), "{idle:?}");
    }

    assert_eq!(
        engine_idle_at(ENGINE_IDLE_CHOICES.len() as u16),
        None,
        "an id past the last item is not one the menu offered"
    );
    assert_eq!(engine_idle_at(0), Some(EngineIdle::Indefinite));
}

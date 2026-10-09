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
        (ID_TRAY_PIN_MINIMIZE_BUBBLE, ID_TRAY_PIN_MINIMIZE_TASKBAR),
        (ID_TRAY_PIN_MINIMIZE_BUBBLE, ID_TRAY_PIN),
        (ID_TRAY_PIN_MINIMIZE_BUBBLE, ID_TRAY_PIN_UPDATE),
        (ID_TRAY_PIN_MINIMIZE_BUBBLE, ID_TRAY_PIN_UPDATE_HOVER),
        (ID_TRAY_PIN_MINIMIZE_BUBBLE, ID_TRAY_PIN_NAV_ALL),
        (ID_TRAY_PIN_MINIMIZE_BUBBLE, ID_TRAY_PIN_NAV_CATEGORY),
        (ID_TRAY_PIN_MINIMIZE_BUBBLE, ID_TRAY_PIN_PAUSE_AUDIO),
        (ID_TRAY_PIN_MINIMIZE_BUBBLE, ID_TRAY_PIN_PAUSE_VIDEO),
        (ID_TRAY_PIN_MINIMIZE_TASKBAR, ID_TRAY_PIN),
        (ID_TRAY_PIN_MINIMIZE_TASKBAR, ID_TRAY_PIN_UPDATE),
        (ID_TRAY_PIN_MINIMIZE_TASKBAR, ID_TRAY_PIN_UPDATE_HOVER),
        (ID_TRAY_PIN_MINIMIZE_TASKBAR, ID_TRAY_PIN_NAV_ALL),
        (ID_TRAY_PIN_MINIMIZE_TASKBAR, ID_TRAY_PIN_NAV_CATEGORY),
        (ID_TRAY_PIN_MINIMIZE_TASKBAR, ID_TRAY_PIN_PAUSE_AUDIO),
        (ID_TRAY_PIN_MINIMIZE_TASKBAR, ID_TRAY_PIN_PAUSE_VIDEO),
    ] {
        assert_ne!(row, other, "two rows of the menu share the id {row}");
    }

    // And outside every range a click is read against before this row is, so a walk's
    // two answers are never answered as an item of a submenu of their own — the
    // backdrop halves beside the pin block among them, since a row inside one would
    // be read as that half's own backdrop.
    for (base, len) in [
        (
            ID_TRAY_LIBREOFFICE_IDLE_BASE,
            ENGINE_IDLE_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_IMAGE_BACKGROUND_BASE,
            BACKGROUND_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_HTML_BACKGROUND_BASE,
            HTML_BACKGROUND_CHOICES.len() as u16,
        ),
        (ID_TRAY_SCALE_BASE, BITMAP_SCALE_CHOICES.len() as u16),
    ] {
        for row in [
            ID_TRAY_PIN_NAV_ALL,
            ID_TRAY_PIN_NAV_CATEGORY,
            ID_TRAY_PIN_MINIMIZE_BUBBLE,
            ID_TRAY_PIN_MINIMIZE_TASKBAR,
        ] {
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
    assert_eq!(
        defaults.pin_minimize_to, DEFAULT_PIN_MINIMIZE_TO,
        "a pin put away goes to the bubble on the desktop until it is told otherwise"
    );
}

/// The two places a minimize can put a pin are named for where the pin goes, and the one the
/// setting starts at is marked as the default, so a user reading the submenu is told both which
/// place is chosen and which one a setting nobody has changed would have chosen.
#[test]
fn the_minimize_rows_say_where_a_pin_goes() {
    assert_eq!(
        pin_minimize_label(PinMinimizeTo::Bubble),
        "To Bubble (Default)"
    );
    assert_eq!(pin_minimize_label(PinMinimizeTo::Taskbar), "To Taskbar");
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

/// The `Volume → Pin Mode Audio Seek` submenu lists the same four
/// ways of starting a sound the `Audio Seek` one does, in the order
/// the ids are handed out in, but marks the way *this* setting
/// starts at — the beginning — rather than the way the hover's does,
/// and each id of its own range resolves back to the way its item
/// was listed for, which is what makes a click start a pinned sound
/// where it named.
#[test]
fn every_offered_way_of_starting_a_pinned_sound_is_one_the_setting_keeps() {
    assert_eq!(
        AUDIO_SEEK_CHOICES.map(pin_mode_audio_seek_label),
        [
            "Remember".to_string(),
            "From the Start (Default)".to_string(),
            "From the Middle".to_string(),
            "Random".to_string()
        ]
    );

    let marked: Vec<AudioSeek> = AUDIO_SEEK_CHOICES
        .iter()
        .copied()
        .filter(|seek| pin_mode_audio_seek_label(*seek).ends_with(" (Default)"))
        .collect();

    assert_eq!(
        marked,
        [DEFAULT_PIN_MODE_AUDIO_SEEK],
        "the way the pin's setting starts at is the way the pin's menu marks"
    );

    assert_eq!(
        DEFAULT_PIN_MODE_AUDIO_SEEK,
        AudioSeek::Start,
        "the pin's menu marks the beginning, not where a sound was left"
    );

    for (index, seek) in AUDIO_SEEK_CHOICES.iter().enumerate() {
        assert_eq!(audio_seek_at(index as u16), Some(*seek));
    }

    assert_eq!(
        audio_seek_at(AUDIO_SEEK_CHOICES.len() as u16),
        None,
        "an id past the last item is not one the menu offered"
    );

    // The two `Audio Seek` submenus are two ranges of their own, so
    // a click on one is never a click on the other — and neither is
    // a click on a level of either volume half, which is the bargain
    // `the_two_volume_submenus_carry_a_range_apiece` holds the
    // hover's range to.
    let pinned = ID_TRAY_PIN_MODE_AUDIO_SEEK_BASE
        ..ID_TRAY_PIN_MODE_AUDIO_SEEK_BASE + AUDIO_SEEK_CHOICES.len() as u16;
    let hover = ID_TRAY_AUDIO_SEEK_BASE..ID_TRAY_AUDIO_SEEK_BASE + AUDIO_SEEK_CHOICES.len() as u16;
    assert!(
        !pinned.contains(&hover.start) && !hover.contains(&pinned.start),
        "the ranges {pinned:?} and {hover:?} overlap"
    );

    for base in [ID_TRAY_VIDEO_VOLUME_BASE, ID_TRAY_AUDIO_VOLUME_BASE] {
        let range = base..base + VOLUME_CHOICES.len() as u16;
        assert!(
            !range.contains(&pinned.start) && !pinned.contains(&range.start),
            "the ranges {range:?} and {pinned:?} overlap"
        );
    }
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
            "Nothing".to_string(),
            "Filename (Default)".to_string(),
            "Filename Column".to_string(),
            "Details".to_string()
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

/// The `Image`, `Video` and `Animated` submenus are one
/// range each, and none of them reaches into another or into the display shares the
/// document scales beside them hand out: a click on a share of a bitmap is never read
/// as a click on another setting's share. The `Audio` submenu beside
/// them is one range of the display shares, and it overlaps none of those or
/// of the bitmap ranges either. Each lists every share the setting can be
/// asked for, in order, and every id resolves back to the share its item was
/// listed for — which is what makes a click select what it named.
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
    let audio_scales = (
        ID_TRAY_AUDIO_SCALE_BASE,
        ID_TRAY_AUDIO_SCALE_BASE + AUDIO_SCALE_CHOICES.len() as u16,
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
            audio_scales,
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

    // The sound's shares answer the question the document scales answer —
    // of the display rather than of a file's own size — so the range sits
    // in their block, and it overlaps none of those either.
    for document in [
        font_scales,
        design_scales,
        ebook_scales,
        document_scales,
        vector_scales,
        text_scales,
    ] {
        assert_eq!(
            overlaps(audio_scales, document),
            None,
            "the range {audio_scales:?} and the display shares {document:?} overlap"
        );
    }

    for (index, scale) in AUDIO_SCALE_CHOICES.iter().enumerate() {
        assert_eq!(audio_scale_at(index as u16), Some(*scale));
    }

    assert_eq!(
        audio_scale_at(AUDIO_SCALE_CHOICES.len() as u16),
        None,
        "an id past the last item of the range at {ID_TRAY_AUDIO_SCALE_BASE} is not one it offered"
    );
}

/// The rows the redesign renamed say what they do: the pin's own
/// follow, the two walks a pin takes, the level a pin keeps, and the
/// update a check found.
#[test]
fn the_renamed_rows_say_what_they_do() {
    assert_eq!(pin_update_label(), "Follow Selection");
    assert_eq!(
        pin_nav_label(crate::config::config::PinNavFileTypes::All),
        "All Files"
    );
    assert_eq!(
        pin_nav_label(crate::config::config::PinNavFileTypes::Category),
        "Same Category"
    );
    assert_eq!(remember_volume_label(), "Remember Level");
    assert_eq!(update_available_label("0.3.4"), "Update Available (v0.3.4)");
    assert_eq!(system_menu_label("0.4.0"), "System (v0.4.0)");
}

/// The `Audio` range sits in the stretch between the volume
/// toggles and the `Codecs` commands, and every id of it answers as a
/// share of the display and nothing else: the range overlaps no other
/// id the app hands out — not another scale's range, not the avoid
/// ways, the reset rows, the volume switches, the engine-idle times,
/// the cache sizes, the font sizes, the theme folder's own ids, the
/// codec rows, the delays and the tick, the away timer, the persistent
/// toggles, the hardware-acceleration row, nor any single row of the
/// menu — and no two of those overlap each other either. Holding the
/// range against the scale ranges alone is how a range landing on the
/// reset rows went unnoticed, so every id the app hands out is held
/// against it here.
#[test]
fn the_audio_scaling_range_sits_apart_from_every_other_id() {
    let audio_scales = (
        ID_TRAY_AUDIO_SCALE_BASE,
        ID_TRAY_AUDIO_SCALE_BASE + AUDIO_SCALE_CHOICES.len() as u16,
    );

    // Every other id the app hands out, as the half-open range the
    // window proc reads it as: a submenu's range is its base plus the
    // choices it is as wide as, and a row of its own is one id wide.
    let mut others: Vec<(u16, u16)> = vec![
        // The scale ranges, each as wide as the shares its submenu lists.
        (ID_TRAY_SCALE_BASE, ID_TRAY_SCALE_BASE + BITMAP_SCALE_CHOICES.len() as u16),
        (
            ID_TRAY_VIDEO_SCALE_BASE,
            ID_TRAY_VIDEO_SCALE_BASE + BITMAP_SCALE_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_ANIMATED_SCALE_BASE,
            ID_TRAY_ANIMATED_SCALE_BASE + BITMAP_SCALE_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_VECTOR_SCALE_BASE,
            ID_TRAY_VECTOR_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_EBOOK_SCALE_BASE,
            ID_TRAY_EBOOK_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_DOCUMENT_SCALE_BASE,
            ID_TRAY_DOCUMENT_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_FONT_SCALE_BASE,
            ID_TRAY_FONT_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_DESIGN_SCALE_BASE,
            ID_TRAY_DESIGN_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_TEXT_SCALE_BASE,
            ID_TRAY_TEXT_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
        ),
        // The ways of keeping a preview off the item it is about.
        (ID_TRAY_AVOID_BASE, ID_TRAY_AVOID_BASE + AVOID_CHOICES.len() as u16),
        // The engine-idle times, one range apiece.
        (
            ID_TRAY_ENGINE_IDLE_BASE,
            ID_TRAY_ENGINE_IDLE_BASE + ENGINE_IDLE_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_WEBVIEW_IDLE_BASE,
            ID_TRAY_WEBVIEW_IDLE_BASE + ENGINE_IDLE_CHOICES.len() as u16,
        ),
        (
            ID_TRAY_LIBREOFFICE_IDLE_BASE,
            ID_TRAY_LIBREOFFICE_IDLE_BASE + ENGINE_IDLE_CHOICES.len() as u16,
        ),
        // The cache sizes, one range per cache.
        (
            ID_TRAY_IMAGE_CACHE_BASE,
            ID_TRAY_IMAGE_CACHE_BASE + CACHE_SIZE_CHOICES_MB.len() as u16,
        ),
        (
            ID_TRAY_DOCUMENT_CACHE_BASE,
            ID_TRAY_DOCUMENT_CACHE_BASE + CACHE_SIZE_CHOICES_MB.len() as u16,
        ),
        (
            ID_TRAY_IMAGE_DISK_CACHE_BASE,
            ID_TRAY_IMAGE_DISK_CACHE_BASE + CACHE_SIZE_CHOICES_MB.len() as u16,
        ),
        (
            ID_TRAY_GENERAL_DISK_CACHE_BASE,
            ID_TRAY_GENERAL_DISK_CACHE_BASE + CACHE_SIZE_CHOICES_MB.len() as u16,
        ),
        // The delays the timing submenus offer, and the loop's own tick.
        (
            ID_TRAY_DELAY_BASE,
            ID_TRAY_DELAY_BASE + TIMING_DELAY_CHOICES_MS.len() as u16,
        ),
        (
            ID_TRAY_REHOVER_DELAY_BASE,
            ID_TRAY_REHOVER_DELAY_BASE + TIMING_DELAY_CHOICES_MS.len() as u16,
        ),
        (
            ID_TRAY_SETTLING_DELAY_BASE,
            ID_TRAY_SETTLING_DELAY_BASE + TIMING_DELAY_CHOICES_MS.len() as u16,
        ),
        (ID_TRAY_TICK_BASE, ID_TRAY_TICK_BASE + TICK_CHOICES_MS.len() as u16),
        // The away times, the persistent toggles above them, and the
        // theme folder's own items up to where the cache sizes begin.
        (
            ID_TRAY_AFK_TIMER_BASE,
            ID_TRAY_AFK_TIMER_BASE + AFK_TIMER_CHOICES_SECS.len() as u16,
        ),
        (ID_TRAY_ENGINE_PERSISTENT_BASE, ID_TRAY_ENGINE_PERSISTENT_BASE + 3),
        (ID_TRAY_THEME_CUSTOM_BASE, ID_TRAY_IMAGE_CACHE_BASE),
        // The codec rows, wider than the list is long.
        (ID_TRAY_CODEC_BASE, ID_TRAY_CODEC_BASE + CODEC_COMMANDS),
    ];
    // The font sizes, each the one id its size carries.
    others.extend(FONT_SIZE_CHOICES.map(|(_, id)| (id, id + 1)));
    // And the single rows: the reset pair, the update check, the volume
    // switches, the hardware-acceleration row, and every other lone row
    // the menu holds.
    others.extend(
        [
            ID_TRAY_RESET_SETTINGS,
            ID_TRAY_RESET_LISTS,
            ID_TRAY_CHECK_UPDATES,
            ID_TRAY_NORMALIZE_VOLUME,
            ID_TRAY_NORMALIZE_VIDEO_VOLUME,
            ID_TRAY_REMEMBER_VOLUME,
            ID_TRAY_REMEMBER_VIDEO_VOLUME,
            ID_TRAY_VIDEO_HW_ACCEL,
            ID_TRAY_PRIORITIZE_KEYBOARD,
            ID_TRAY_EXIT,
            ID_TRAY_STARTUP,
            ID_TRAY_UPDATE,
            ID_TRAY_ENABLE,
            ID_TRAY_PIN,
            ID_TRAY_PIN_UPDATE,
            ID_TRAY_PIN_UPDATE_HOVER,
            ID_TRAY_PIN_PAUSE_AUDIO,
            ID_TRAY_PIN_PAUSE_VIDEO,
            ID_TRAY_PIN_NAV_ALL,
            ID_TRAY_PIN_NAV_CATEGORY,
            ID_TRAY_TRIGGER_DISABLE,
            ID_TRAY_TRIGGER_ENABLE,
            ID_TRAY_TRIGGER_ENABLED,
            ID_TRAY_TRIGGER_AFFECT_PIN,
            ID_TRAY_ENGINE_OFFICE_MS,
            ID_TRAY_ENGINE_OFFICE_LIBRE,
            ID_TRAY_VIDEO_ENGINE_FALLBACK,
            ID_TRAY_POSITION_FOLLOW,
            ID_TRAY_POSITION_BEST,
            ID_TRAY_OPEN_CONFIG,
            ID_TRAY_THEME_LIGHT,
            ID_TRAY_THEME_DARK,
            ID_TRAY_MARKDOWN_RENDERED,
            ID_TRAY_MARKDOWN_SOURCE,
            ID_TRAY_RENDER_HTML,
            ID_TRAY_TYPE_IMAGES,
            ID_TRAY_TYPE_VIDEOS,
            ID_TRAY_TYPE_AUDIO,
            ID_TRAY_TYPE_TEXT,
            ID_TRAY_TYPE_EBOOK,
            ID_TRAY_TYPE_ARCHIVES,
            ID_TRAY_TYPE_DOCUMENT,
            ID_TRAY_TYPE_FONTS,
            ID_TRAY_TYPE_DESIGN,
            ID_TRAY_TYPE_VECTOR,
        ]
        .map(|id| (id, id + 1)),
    );

    let overlaps = |ours: (u16, u16), theirs: (u16, u16)| {
        (ours.0 < theirs.1 && theirs.0 < ours.1).then_some((ours, theirs))
    };

    for other in &others {
        assert_eq!(
            overlaps(audio_scales, *other),
            None,
            "the audio range {audio_scales:?} and the id {other:?} overlap"
        );
    }

    for (index, range) in others.iter().enumerate() {
        for other in others.iter().skip(index + 1) {
            assert_eq!(
                overlaps(*range, *other),
                None,
                "the ranges {range:?} and {other:?} overlap"
            );
        }
    }
}

/// The `Image`, `Video` and `Animated` submenus offer the shares a
/// bitmap is drawn at — the shares of the display's fitted size and of a
/// bitmap's own size, rather than the share of the display the document
/// scales beside them are — in one order and with one set of labels: what a
/// share is called does not depend on which of the three is asking, and
/// exactly one label — the share each setting starts at — reads as the
/// default. Every share an item can pick is one the setting keeps, so a
/// choice made here is still the choice after a restart.
#[test]
fn every_offered_bitmap_scale_is_one_the_setting_keeps() {
    assert_eq!(
        BITMAP_SCALE_CHOICES.map(|scale| bitmap_scale_label(scale, DEFAULT_PREVIEW_SCALE)),
        [
            "Fit to Screen".to_string(),
            "75%".to_string(),
            "50%".to_string(),
            "25%".to_string(),
            "10%".to_string(),
            "400%".to_string(),
            "300%".to_string(),
            "200%".to_string(),
            "150%".to_string(),
            "100% (Default)".to_string(),
            "50%".to_string(),
            "25%".to_string(),
        ]
    );

    // The two groups the `By Screen` and `By Own Size` rows stand
    // between: every share of the display's fitted size first, then
    // every share of a bitmap's own size — the boundary the
    // submenu's separators and its two rows are placed at (see
    // `append_bitmap_scale_menu`).
    let own_size_begin = BITMAP_SCALE_CHOICES
        .iter()
        .position(|choice| matches!(choice, PreviewScale::Percent(_)))
        .expect("a share of a bitmap's own size is offered");
    assert!(
        BITMAP_SCALE_CHOICES[..own_size_begin]
            .iter()
            .all(|choice| !matches!(choice, PreviewScale::Percent(_))),
        "the screen group holds the fitted-size shares alone"
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

/// The `Audio` submenu offers the shares of the display a sound's card
/// is laid out over — the share of the room rather than of a file's own size —
/// in one order, and the labels are the document scale's, which name the same
/// question: what a share is called does not depend on which of the two is
/// asking, and exactly one label — the share the setting starts at — reads as
/// the default. Every share an item can pick is one the setting keeps, so a
/// choice made here is still the choice after a restart.
#[test]
fn every_offered_audio_scale_is_one_the_setting_keeps() {
    assert_eq!(
        AUDIO_SCALE_CHOICES.map(|scale| document_scale_label(scale, DEFAULT_AUDIO_SCALE)),
        [
            "25%".to_string(),
            "20%".to_string(),
            "15%".to_string(),
            "10% (Default)".to_string(),
            "7%".to_string(),
            "5%".to_string(),
        ]
    );

    let marked: Vec<PreviewScale> = AUDIO_SCALE_CHOICES
        .iter()
        .copied()
        .filter(|scale| document_scale_label(*scale, DEFAULT_AUDIO_SCALE).ends_with(" (Default)"))
        .collect();

    assert_eq!(
        marked,
        [DEFAULT_AUDIO_SCALE],
        "one share is the default at {DEFAULT_AUDIO_SCALE:?}"
    );

    for scale in AUDIO_SCALE_CHOICES {
        let written = scale.as_str();

        assert_eq!(
            PreviewScale::from_audio_str(&written),
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

/// The `Engine → Video Engine` rows are ids of their own, and none of them falls inside
/// another submenu's range: a click read as an idle time or as a backdrop would set that setting
/// instead of the engine the user named.
///
/// The engines the submenu lists are the whole of what the setting holds and every one of them is
/// named, with the one the app starts at marked — the mark is read from the constant the setting
/// starts at rather than written into the label, the way every other value menu here reads it.
#[test]
fn the_video_engine_rows_are_ids_of_their_own() {
    let engines =
        ID_TRAY_VIDEO_ENGINE_BASE..ID_TRAY_VIDEO_ENGINE_BASE + VIDEO_ENGINE_CHOICES.len() as u16;

    assert!(
        !engines.contains(&ID_TRAY_VIDEO_ENGINE_FALLBACK),
        "the switch at the top of the submenu is not one of the engines below it"
    );

    for (base, len, what) in [
        (
            ID_TRAY_LIBREOFFICE_IDLE_BASE,
            ENGINE_IDLE_CHOICES.len() as u16,
            "an idle time",
        ),
        (
            ID_TRAY_IMAGE_BACKGROUND_BASE,
            BACKGROUND_CHOICES.len() as u16,
            "a backdrop",
        ),
    ] {
        let range = base..base + len;

        assert!(
            !range.contains(&ID_TRAY_VIDEO_ENGINE_FALLBACK)
                && !range.contains(&engines.start)
                && !engines.contains(&range.start),
            "the video rows and the range at {base} overlap, so a click on {what} is answered \
             as the other"
        );
    }

    assert_eq!(
        VIDEO_ENGINE_CHOICES.map(video_engine_label),
        [
            "Best (Default)".to_string(),
            "Native".to_string(),
            "FFmpeg".to_string(),
            "Native (FFmpeg above 3.2MP)".to_string(),
        ]
    );

    let marked: Vec<VideoEngine> = VIDEO_ENGINE_CHOICES
        .iter()
        .copied()
        .filter(|engine| video_engine_label(*engine).ends_with(" (Default)"))
        .collect();

    assert_eq!(
        marked,
        [DEFAULT_VIDEO_ENGINE],
        "the engine the setting starts at is the one the menu marks"
    );

    // Every engine an item can pick is one the setting holds, so a choice made here is still the
    // choice after a restart.
    for engine in VIDEO_ENGINE_CHOICES {
        let written = engine.as_str();

        assert_eq!(
            VideoEngine::from_str(written),
            Some(engine),
            "`{written}` read back"
        );
    }

    let defaults = crate::config::config::AppConfig::default();

    assert_eq!(
        defaults.video_engine, DEFAULT_VIDEO_ENGINE,
        "the best engine the machine has is the choice until the user makes one"
    );
    assert_eq!(
        defaults.video_engine_fallback, DEFAULT_VIDEO_ENGINE_FALLBACK,
        "a file the chosen engine cannot play is played by another one until that is turned off"
    );
    assert!(
        DEFAULT_VIDEO_ENGINE_FALLBACK,
        "the app starts with the fallback on"
    );

    // The rows that are greyed are the ones naming a player this machine has not got: the two the
    // app and Windows supply are always here, and both FFmpeg-based choices need `ffplay`.
    assert!(
        VideoEngine::Best.installed() && VideoEngine::Native.installed(),
        "the app's own answer and the engine Windows ships are never greyed"
    );
    assert_eq!(
        VideoEngine::Ffmpeg.installed(),
        crate::formats::codecs::ffplay_available(),
        "FFmpeg's row is offered exactly where `ffplay` is installed"
    );
    assert_eq!(
        VideoEngine::Hybrid.installed(),
        crate::formats::codecs::ffplay_available(),
        "and the hybrid needs `ffplay` for the half of it that hands a film over"
    );
}

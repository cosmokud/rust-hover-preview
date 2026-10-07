use super::*;

/// The three settings the card's menu holds, put where a test
/// asks and put back when the guard is dropped, so a test beside
/// this one is answered with the machine's own (the house
/// pattern: `SeekSettings` in `pin_audio_seek`).
struct MenuSettings {
    seek: AudioSeek,
    shuffle: bool,
    loop_: bool,
}

impl MenuSettings {
    fn put(seek: AudioSeek, shuffle: bool, loop_: bool) -> Self {
        let mut config = CONFIG.lock().expect("the configuration");
        let was = MenuSettings {
            seek: config.pin_mode_audio_seek,
            shuffle: config.pin_mode_audio_shuffle,
            loop_: config.pin_mode_audio_loop,
        };
        config.pin_mode_audio_seek = seek;
        config.pin_mode_audio_shuffle = shuffle;
        config.pin_mode_audio_loop = loop_;

        was
    }
}

impl Drop for MenuSettings {
    fn drop(&mut self) {
        if let Ok(mut config) = CONFIG.lock() {
            config.pin_mode_audio_seek = self.seek;
            config.pin_mode_audio_shuffle = self.shuffle;
            config.pin_mode_audio_loop = self.loop_;
        }
    }
}

/// The rows the menu holds are the settings it is the face of:
/// the first page is the two mode toggles — each marked as the
/// setting it stands for is — and the row that opens the seek
/// choices, and the seek page is the four choices, in the order
/// the tray lists them in, with the one the pin plays by marked.
#[test]
fn the_menu_holds_the_settings_it_is_the_face_of() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Middle, true, false);

    let top = menu_rows(false);
    let labels: Vec<&str> = top.iter().map(|row| row.label.as_str()).collect();
    assert_eq!(labels, ["Shuffle Mode", "Loop", "Seek"]);
    assert!(top[0].marked, "shuffle is on, so its row is marked");
    assert!(!top[1].marked, "loop is off, so its row is not");
    assert!(
        !top[2].marked,
        "the seek row opens the choices, so it is not one of them"
    );

    let choices = menu_rows(true);
    let labels: Vec<&str> = choices.iter().map(|row| row.label.as_str()).collect();
    assert_eq!(
        labels,
        ["Remember", "From the Start", "From the Middle", "Random"]
    );
    let marked = choices.iter().position(|row| row.marked);
    assert_eq!(
        marked,
        Some(2),
        "the pin plays from the middle, so that choice is the one marked"
    );

    // A configuration that names the defaults is a menu that
    // marks the defaults: shuffle off, loop on, and the
    // beginning the way a pinned sound starts.
    let _settings = MenuSettings::put(
        DEFAULT_PIN_MODE_AUDIO_SEEK,
        DEFAULT_PIN_MODE_AUDIO_SHUFFLE,
        DEFAULT_PIN_MODE_AUDIO_LOOP,
    );
    let top = menu_rows(false);
    assert!(!top[0].marked, "shuffle starts off, so its row is not marked");
    assert!(top[1].marked, "loop starts on, so its row is");
    let choices = menu_rows(true);
    let marked = choices.iter().position(|row| row.marked);
    assert_eq!(
        marked,
        Some(1),
        "a pinned sound starts at the beginning, so that choice is the one marked"
    );
}

/// The panel a menu opens hangs from the cell the card's mark is
/// drawn in: its left edge on the cell's own left edge, below
/// the cell's own row rather than over it, and inside the window
/// the pin stands in — and a menu that is not up is nothing at
/// all, which is what a paint with no panel to draw is answered
/// with.
#[test]
fn the_menu_hangs_from_the_cell_the_mark_is_drawn_in() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, false, true);

    // A pin stood up for this test and whatever was there
    // before it put back when the test is done — even when it
    // is done by panicking — because a test that leaves a pin
    // up is a test every test after it is standing inside (see
    // `stand_pin`).
    struct PutBack(Option<PinnedPreview>);
    impl Drop for PutBack {
        fn drop(&mut self) {
            stand_pin(self.0.take());
        }
    }
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(sound_pin()));

    assert!(
        pinned_menu_geometry().is_none(),
        "a menu that is not up is not painted"
    );

    with_pin(|pin| pin.menu.open = true);

    let Some(paint) = pinned_menu_geometry() else {
        panic!("a menu that is up is painted");
    };

    // The cell the menu hangs from, in the window's own
    // coordinates: the card's own box for the cell, moved down
    // by the band the card is drawn in (see
    // `pinned_audio_volume_popup`).
    let (width, height, cell) = pin_state()
        .and_then(|pinned| {
            let pin = pinned.pin()?;
            let (width, height) = pin.window_size();
            let (top, _) = pinned_band_rows(
                height,
                pin.caption,
                pinned_transport_height(pin.dpi, pin.transport_bar),
                pin.overlay,
            );
            let cell = audio_preview::control_box(
                CardControl::Menu,
                (pin.content.2 - pin.content.0).max(1) as u32,
                pin.dpi,
                pinned_audio_options(pin),
                true,
            )?;
            Some((
                width,
                height,
                RECT {
                    top: cell.top + top,
                    ..cell
                },
            ))
        })
        .expect("the pin that was installed");

    assert_eq!(
        paint.popup.panel.left, cell.left,
        "the panel's left edge is the cell's own"
    );
    assert!(
        paint.popup.panel.top >= cell.bottom,
        "the panel hangs below the cell's row, not over it"
    );
    assert!(paint.popup.panel.right <= width, "the panel is inside the window");
    assert!(
        paint.popup.panel.bottom <= height,
        "and so is its bottom edge"
    );

    // The rows are the first page's, because the menu was put up
    // on it, and each says what the setting it stands for says.
    let labels: Vec<&str> = paint.rows.iter().map(|row| row.label.as_str()).collect();
    assert_eq!(labels, ["Shuffle Mode", "Loop", "Seek"]);
    assert!(!paint.rows[0].marked, "shuffle is off");
    assert!(paint.rows[1].marked, "loop is on");
    assert!(!paint.rows[2].marked, "the seek row is not a choice");
}

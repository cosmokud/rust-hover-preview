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

/// What the configuration file was when a test began: the bytes
/// it held, no file at all, or a file there that cannot be read —
/// which is a file a test cannot write back either, so it is left
/// exactly where it is.
enum ConfigFileWas {
    Bytes(Vec<u8>),
    NoFile,
    Unreadable,
}

/// The configuration file as a test found it, put back when the
/// guard is dropped. A press on a menu's check row writes the
/// configuration the way every setting this app's own menus write
/// is, so a test that presses one leaves the file the press wrote
/// unless it is put back.
struct ConfigFilePutBack {
    path: Option<std::path::PathBuf>,
    was: ConfigFileWas,
}

impl ConfigFilePutBack {
    fn take() -> Self {
        let path = crate::config::config::AppConfig::config_path();
        let was = match path.as_ref() {
            Some(path) if path.exists() => match std::fs::read(path) {
                Ok(bytes) => ConfigFileWas::Bytes(bytes),
                Err(_) => ConfigFileWas::Unreadable,
            },
            _ => ConfigFileWas::NoFile,
        };
        Self { path, was }
    }
}

impl Drop for ConfigFilePutBack {
    fn drop(&mut self) {
        let Some(path) = self.path.as_ref() else {
            return;
        };

        match &self.was {
            ConfigFileWas::Bytes(bytes) => {
                if let Some(parent) = path.parent() {
                    let _ = std::fs::create_dir_all(parent);
                }
                let _ = std::fs::write(path, bytes);
            }
            // There was no file before the test, so the file the
            // test's own press wrote is taken away again.
            ConfigFileWas::NoFile => {
                let _ = std::fs::remove_file(path);
            }
            // A file that cannot be read is one that cannot be
            // written back, so it is left where it is.
            ConfigFileWas::Unreadable => {}
        }
    }
}

/// A pin stood up for a test and whatever was there before it put
/// back when the test is done — even when it is done by panicking —
/// because a test that leaves a pin up is a test every test after it
/// is standing inside (see `stand_pin`).
struct PutBack(Option<PinnedPreview>);
impl Drop for PutBack {
    fn drop(&mut self) {
        stand_pin(self.0.take());
    }
}

/// The rows the menu holds are the settings it is the face of:
/// the first page is the two mode toggles — each carrying the
/// checkbox of the setting it stands for, checked where the
/// setting is on and empty where it is off — and the row that
/// opens the seek choices, and the seek page is the four
/// choices, in the order the tray lists them in, with the one
/// the pin plays by carrying its disc.
#[test]
fn the_menu_holds_the_settings_it_is_the_face_of() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Middle, true, false);

    let top = menu_rows();
    let labels: Vec<&str> = top.iter().map(|row| row.label.as_str()).collect();
    assert_eq!(labels, ["Shuffle Mode", "Loop", "Seek"]);
    assert!(
        matches!(top[0].mark, pin_chrome::MenuMark::Check(true)),
        "shuffle is on, so its row is checked"
    );
    assert!(
        matches!(top[1].mark, pin_chrome::MenuMark::Check(false)),
        "loop is off, so its row is an empty box"
    );
    assert!(
        matches!(top[2].mark, pin_chrome::MenuMark::None),
        "the seek row opens the choices, so it carries no mark"
    );

    let choices = seek_rows();
    let labels: Vec<&str> = choices.iter().map(|row| row.label.as_str()).collect();
    assert_eq!(
        labels,
        ["Remember", "From the Start", "From the Middle", "Random"]
    );
    let marked = choices
        .iter()
        .position(|row| matches!(row.mark, pin_chrome::MenuMark::Bullet));
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
    let top = menu_rows();
    assert!(
        matches!(top[0].mark, pin_chrome::MenuMark::Check(false)),
        "shuffle starts off, so its row is an empty box"
    );
    assert!(
        matches!(top[1].mark, pin_chrome::MenuMark::Check(true)),
        "loop starts on, so its row is checked"
    );
    let choices = seek_rows();
    let marked = choices
        .iter()
        .position(|row| matches!(row.mark, pin_chrome::MenuMark::Bullet));
    assert_eq!(
        marked,
        Some(1),
        "a pinned sound starts at the beginning, so that choice is the one marked"
    );
}

/// The panel a menu opens hangs from the gear the card's
/// menu opens from: it opens downward from the band the gear
/// stands in, with its right edge tucked to the gear's own
/// left edge a small gap off it — the placement that leaves a
/// flyout room to the menu's right — and inside the window
/// the pin stands in — and a menu that is not up is nothing
/// at all, which is what a paint with no panel to draw is
/// answered with.
#[test]
fn the_menu_hangs_from_the_gear_the_card_s_menu_opens_from() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, false, true);
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

    // The gear the menu hangs from, in the window's own
    // coordinates: the card's own box for it, moved down
    // by the band the card is drawn in (see
    // `pinned_audio_volume_popup`).
    let (width, height, gear) = pin_state()
        .and_then(|pinned| {
            let pin = pinned.pin()?;
            let (width, height) = pin.window_size();
            let (top, _) = pinned_band_rows(
                height,
                pin.caption,
                pinned_transport_height(pin.dpi, pin.transport_bar),
                pin.overlay,
            );
            let gear = audio_preview::control_box(
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
                    top: gear.top + top,
                    ..gear
                },
            ))
        })
        .expect("the pin that was installed");

    // The panel's right edge is tucked to the gear's own
    // left edge, one small gap off it (see
    // `pin_chrome::menu_popup_from_button`). At the pin's
    // own 96-DPI scale the gap is the four logical pixels
    // the panel is held off the button it came from.
    assert_eq!(
        paint.popup.panel.right,
        gear.left - 4,
        "the panel's right edge is one gap left of the gear's own left edge"
    );
    assert!(
        paint.popup.panel.top >= gear.bottom,
        "the panel hangs below the band the gear stands in, not over it"
    );
    assert!(paint.popup.panel.right <= width, "the panel is inside the window");
    assert!(
        paint.popup.panel.bottom <= height,
        "and so is its bottom edge"
    );

    // The rows are the first page's, because the menu was put up
    // on it, and each carries the mark of the setting it stands
    // for.
    let labels: Vec<&str> = paint.rows.iter().map(|row| row.label.as_str()).collect();
    assert_eq!(labels, ["Shuffle Mode", "Loop", "Seek"]);
    assert!(
        matches!(paint.rows[0].mark, pin_chrome::MenuMark::Check(false)),
        "shuffle is off, so its row is an empty box"
    );
    assert!(
        matches!(paint.rows[1].mark, pin_chrome::MenuMark::Check(true)),
        "loop is on, so its row is checked"
    );
    assert!(
        matches!(paint.rows[2].mark, pin_chrome::MenuMark::None),
        "the seek row is not a choice"
    );

    // A menu put up on its main page holds no flyout:
    // the Seek row is the door the flyout comes
    // through, not the flyout itself.
    assert!(
        paint.flyout.is_none(),
        "the main page is the menu's three rows"
    );
}

/// A press on one of the two mode rows turns the setting it
/// stands for over, writes it down, and puts the panel away:
/// the row is a check toggle, so a press on it is an answer
/// rather than a door. The setting is written to `config.ini`
/// the way every setting this app's own menus write is, which
/// is why the file the press writes is put back what it was
/// when the test is done.
#[test]
fn a_press_on_a_mode_row_turns_its_setting_over_and_writes_it_down() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _file = ConfigFilePutBack::take();
    let _settings = MenuSettings::put(AudioSeek::Start, false, true);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(sound_pin()));

    with_pin(|pin| pin.menu.open = true);
    let Some(paint) = pinned_menu_geometry() else {
        panic!("a menu that is up is painted");
    };

    // The middle of the first row, which is the row the
    // Shuffle Mode setting is the face of.
    let x = (paint.popup.panel.left + paint.popup.panel.right) / 2;
    let y = paint.popup.rows_top + paint.popup.row_height / 2;
    assert!(
        unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on a row of the menu is the menu's own"
    );

    // The setting is the opposite of what it was, in the
    // configuration the press holds...
    assert!(
        CONFIG
            .lock()
            .expect("the configuration")
            .pin_mode_audio_shuffle,
        "shuffle was off, so the press turned it on"
    );
    // ... and in the file the press wrote it down to, which
    // is what a read of the configuration from where it
    // lives answers, where there is a file to write to at
    // all.
    if crate::config::config::AppConfig::config_path().is_some() {
        assert!(
            crate::config::config::AppConfig::load().pin_mode_audio_shuffle,
            "the file the press wrote holds the setting turned over"
        );
    }

    // The rows show the setting as it now is: the box the
    // first row carries is checked.
    let rows = menu_rows();
    assert!(
        matches!(rows[0].mark, pin_chrome::MenuMark::Check(true)),
        "the Shuffle Mode row is checked now"
    );

    // And the panel is away, because a toggle is an answer
    // rather than a door.
    assert!(
        pinned_menu_geometry().is_none(),
        "the panel is put away by the press that answered it"
    );
}

/// The second mode row turns its own setting over the same
/// way: the Loop row is the checkbox of the loop setting,
/// and a press on it is the press on the Shuffle Mode row's
/// own, asked of the other row.
#[test]
fn a_press_on_the_loop_row_turns_the_loop_setting_over() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _file = ConfigFilePutBack::take();
    let _settings = MenuSettings::put(AudioSeek::Start, false, true);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(sound_pin()));

    with_pin(|pin| pin.menu.open = true);
    let Some(paint) = pinned_menu_geometry() else {
        panic!("a menu that is up is painted");
    };

    // The middle of the second row, which is the row the
    // Loop setting is the face of.
    let x = (paint.popup.panel.left + paint.popup.panel.right) / 2;
    let y = paint.popup.rows_top + paint.popup.row_height + paint.popup.row_height / 2;
    assert!(
        unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on a row of the menu is the menu's own"
    );

    assert!(
        !CONFIG
            .lock()
            .expect("the configuration")
            .pin_mode_audio_loop,
        "loop was on, so the press turned it off"
    );
    if crate::config::config::AppConfig::config_path().is_some() {
        assert!(
            !crate::config::config::AppConfig::load().pin_mode_audio_loop,
            "the file the press wrote holds the setting turned over"
        );
    }

    let rows = menu_rows();
    assert!(
        matches!(rows[1].mark, pin_chrome::MenuMark::Check(false)),
        "the Loop row is an empty box now"
    );
    assert!(
        pinned_menu_geometry().is_none(),
        "and the panel is put away"
    );
}

/// A press anywhere outside the panel puts it away: the panel
/// is over the card, and a hand that has come for what is
/// under it is a hand that has left the menu. The point is
/// one the panel does not hold and not the gear the menu
/// came from, whose own press is the press that opens and
/// closes it.
#[test]
fn a_press_outside_the_panel_puts_it_away() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, false, true);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(sound_pin()));

    with_pin(|pin| pin.menu.open = true);
    let Some(paint) = pinned_menu_geometry() else {
        panic!("a menu that is up is painted");
    };

    // A point left of the panel, in the row of its first
    // row: not the panel, and not the gear the menu came
    // out of, which stands to the panel's right.
    let x = paint.popup.panel.left - 10;
    let y = paint.popup.rows_top + paint.popup.row_height / 2;
    assert!(
        unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
        "a hand off the panel is still the menu's to answer"
    );
    assert!(
        pinned_menu_geometry().is_none(),
        "so the panel is put away"
    );
}

/// The Seek row opens its flyout beside the menu: the
/// flyout's panel is right of the main menu's own
/// right edge, a small gap off it, and top-aligned
/// with the Seek row, so the menu the flyout came
/// from stays wholly visible beside it — and the
/// flyout holds the four seek choices, in the order
/// the tray lists them in, with the one the pin
/// plays by carrying its disc.
#[test]
fn the_seek_row_opens_its_flyout_beside_the_menu() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Middle, true, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(sound_pin()));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.seek = true;
    });

    let Some(paint) = pinned_menu_geometry() else {
        panic!("a menu that is up is painted");
    };
    let Some(flyout) = paint.flyout.as_ref() else {
        panic!("the Seek row puts the flyout up")
    };

    // The flyout is beside the main menu, not over it:
    // its panel starts a small gap right of the main
    // menu's own right edge — the same gap the menu is
    // held off the gear it came from — and its top is
    // the Seek row's own top, the row the flyout is
    // hung from.
    assert_eq!(
        flyout.popup.panel.left,
        paint.popup.panel.right + 4,
        "the flyout sits right of the menu's own right edge"
    );
    assert_eq!(
        flyout.popup.panel.top,
        paint.popup.rows_top + 2 * paint.popup.row_height,
        "the flyout is top-aligned with the Seek row"
    );

    // The main menu is still inside the window and
    // uncovered, and the flyout is inside the window,
    // which is what keeps it reachable.
    let (width, height) = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_size()))
        .expect("the pin that was installed");
    assert!(paint.popup.panel.left >= 0, "the menu is inside the window");
    assert!(paint.popup.panel.right <= width, "and uncovered to its right");
    assert!(flyout.popup.panel.right <= width, "the flyout is inside the window");
    assert!(flyout.popup.panel.bottom <= height, "and so is its bottom");

    // The flyout holds the four seek choices, in the
    // order the tray lists them in, with the one the
    // pin plays by carrying its disc.
    let labels: Vec<&str> = flyout.rows.iter().map(|row| row.label.as_str()).collect();
    assert_eq!(
        labels,
        ["Remember", "From the Start", "From the Middle", "Random"]
    );
    let marked = flyout
        .rows
        .iter()
        .position(|row| matches!(row.mark, pin_chrome::MenuMark::Bullet));
    assert_eq!(
        marked,
        Some(2),
        "the pin plays from the middle, so that choice is the one marked"
    );
}

/// The flyout's rows answer a press each as its own
/// choice: a press on any row of the flyout is that
/// row's own — the choice pressed is the one the pin
/// plays by from then on, and the menu goes away with
/// the answer — rather than a press outside the menu.
#[test]
fn the_flyout_s_rows_answer_presses() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Remember, false, true);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(sound_pin()));

    // The four choices, in the order the flyout lists
    // them in.
    let choices = [
        AudioSeek::Remember,
        AudioSeek::Start,
        AudioSeek::Middle,
        AudioSeek::Random,
    ];

    for (index, choice) in choices.into_iter().enumerate() {
        with_pin(|pin| {
            pin.menu.open = true;
            pin.menu.seek = true;
        });
        let Some(paint) = pinned_menu_geometry() else {
            panic!("a menu that is up is painted");
        };
        let Some(flyout) = paint.flyout.as_ref() else {
            panic!("the flyout is up")
        };

        // The middle of the flyout's own row, which is
        // the row the choice is the face of.
        let x = (flyout.popup.panel.left + flyout.popup.panel.right) / 2;
        let y = flyout.popup.rows_top
            + index as i32 * flyout.popup.row_height
            + flyout.popup.row_height / 2;
        assert!(
            unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
            "a press on the flyout's row {index} is the menu's own"
        );

        assert_eq!(
            CONFIG
                .lock()
                .expect("the configuration")
                .pin_mode_audio_seek,
            choice,
            "the choice pressed is the one the pin plays by"
        );
        assert!(
            pinned_menu_geometry().is_none(),
            "and the menu goes away with the answer"
        );
    }
}

/// A choice made in the flyout is written down, and
/// the whole menu goes away with it — both panels,
/// because the flyout is the menu's own. The choice
/// is written to `config.ini` the way every setting
/// this app's own menus write is, which is why the
/// file the press writes is put back what it was when
/// the test is done.
#[test]
fn a_choice_made_in_the_flyout_is_written_down_and_puts_the_whole_menu_away() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _file = ConfigFilePutBack::take();
    let _settings = MenuSettings::put(AudioSeek::Start, false, true);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(sound_pin()));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.seek = true;
    });
    let Some(paint) = pinned_menu_geometry() else {
        panic!("a menu that is up is painted");
    };
    let Some(flyout) = paint.flyout.as_ref() else {
        panic!("the flyout is up")
    };

    // The middle of the flyout's last row, which is
    // the Random choice.
    let x = (flyout.popup.panel.left + flyout.popup.panel.right) / 2;
    let y = flyout.popup.rows_top
        + 3 * flyout.popup.row_height
        + flyout.popup.row_height / 2;
    assert!(
        unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on a row of the flyout is the menu's own"
    );

    // The choice is the one the pin plays by from then
    // on, in the configuration the press holds...
    assert_eq!(
        CONFIG
            .lock()
            .expect("the configuration")
            .pin_mode_audio_seek,
        AudioSeek::Random,
        "the pin played from the start, so the press chose Random"
    );
    // ... and in the file the press wrote it down to,
    // which is what a read of the configuration from
    // where it lives answers, where there is a file to
    // write to at all.
    if crate::config::config::AppConfig::config_path().is_some() {
        assert_eq!(
            crate::config::config::AppConfig::load().pin_mode_audio_seek,
            AudioSeek::Random,
            "the file the press wrote holds the choice made"
        );
    }

    // And the whole menu is away — both panels, because
    // the flyout is the menu's own.
    assert!(
        pinned_menu_geometry().is_none(),
        "the whole menu goes away with the choice"
    );
}

/// A press on the Seek row while its flyout is up is
/// the one that puts the flyout back: the row is the
/// door the flyout came through, so a hand pressing it
/// again is a hand taking the flyout back — the menu
/// itself stays up, its three rows still showing.
#[test]
fn a_press_on_the_seek_row_while_its_flyout_is_up_puts_it_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, false, true);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(sound_pin()));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.seek = true;
    });
    let Some(paint) = pinned_menu_geometry() else {
        panic!("a menu that is up is painted");
    };

    // The middle of the Seek row, the main menu's
    // third row.
    let x = (paint.popup.panel.left + paint.popup.panel.right) / 2;
    let y = paint.popup.rows_top
        + 2 * paint.popup.row_height
        + paint.popup.row_height / 2;
    assert!(
        unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on the Seek row is the menu's own"
    );

    // The menu is still up, on its three rows, with the
    // flyout put back.
    let Some(paint) = pinned_menu_geometry() else {
        panic!("the menu is still up")
    };
    assert!(paint.flyout.is_none(), "the flyout is put back");
    let labels: Vec<&str> = paint.rows.iter().map(|row| row.label.as_str()).collect();
    assert_eq!(labels, ["Shuffle Mode", "Loop", "Seek"]);
}

/// A press outside both panels puts the whole menu
/// away: the panels are over the card, and a hand that
/// has come for what is under them is a hand that has
/// left the menu — whether the press lands left of the
/// main menu, in the gap between the two panels, or
/// below the flyout.
#[test]
fn a_press_outside_both_panels_puts_the_whole_menu_away() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, false, true);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(sound_pin()));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.seek = true;
    });
    let Some(paint) = pinned_menu_geometry() else {
        panic!("a menu that is up is painted");
    };
    let Some(flyout) = paint.flyout.as_ref() else {
        panic!("the flyout is up")
    };

    // Three points neither panel holds, in the row of
    // the menu's first row or below the flyout: not the
    // panels, and not the gear the menu came out of,
    // which stands in the band above them.
    let first_row = paint.popup.rows_top + paint.popup.row_height / 2;
    let presses = [
        (paint.popup.panel.left - 10, first_row),
        (paint.popup.panel.right + 2, first_row),
        (
            (flyout.popup.panel.left + flyout.popup.panel.right) / 2,
            flyout.popup.panel.bottom + 10,
        ),
    ];
    for (x, y) in presses {
        with_pin(|pin| {
            pin.menu.open = true;
            pin.menu.seek = true;
        });
        assert!(
            unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
            "a hand off both panels is still the menu's to answer"
        );
        assert!(
            pinned_menu_geometry().is_none(),
            "so the whole menu is put away"
        );
    }
}

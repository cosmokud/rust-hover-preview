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

/// A right-click on an audio pin opens the card's menu at the point it
/// landed — with no panel before the press and the panel there the moment
/// it has landed, which is no animation — and a second right-click while
/// the panel is up moves it to the new point rather than putting it away.
/// A point near the window's own corner is held inside the window.
#[test]
fn a_right_click_opens_the_menu_at_the_point_on_an_audio_pin() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, false, true);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(sound_pin()));

    assert!(
        pinned_menu_geometry().is_none(),
        "a menu that has not been asked for is not painted"
    );

    let (width, height) = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_size()))
        .expect("the pin that was installed");
    let hwnd = HWND(0x1000 as *mut _);

    // A point that leaves the panel room: the panel's own top-left is the
    // point itself, the moment the press has landed.
    let (x, y) = (20, 20);
    unsafe { open_pin_menu(hwnd, x, y) };
    let Some(paint) = pinned_menu_geometry() else {
        panic!("a menu that is up is painted");
    };
    assert_eq!(
        (paint.popup.panel.left, paint.popup.panel.top),
        (x, y),
        "the panel's top-left is the point the right-click landed"
    );

    // It opens on its main page: the two toggles and the Seek row, each
    // carrying the mark of the setting it stands for.
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
        paint.flyout.is_none(),
        "the main page is the menu's three rows"
    );

    // A second right-click moves the panel to the new point, and the menu
    // stays up.
    let (x, y) = (30, 40);
    unsafe { open_pin_menu(hwnd, x, y) };
    let Some(paint) = pinned_menu_geometry() else {
        panic!("the menu is still up");
    };
    assert_eq!(
        (paint.popup.panel.left, paint.popup.panel.top),
        (x, y),
        "a second right-click moves the panel to the new point"
    );

    // A right-click near the far corner would put the panel off the
    // window, so the placement holds it inside.
    unsafe { open_pin_menu(hwnd, width - 2, height - 2) };
    let Some(paint) = pinned_menu_geometry() else {
        panic!("the menu is still up");
    };
    assert!(
        paint.popup.panel.right <= width && paint.popup.panel.bottom <= height,
        "a panel asked for near the corner is held inside the window"
    );
    assert!(paint.popup.panel.left >= 0 && paint.popup.panel.top >= 0);
}

/// A right-click on a pin that is not showing a sound's card opens
/// nothing at all: the menu is the card's own, and only a sound's card
/// carries one (see `pin_shows_an_audio_card`).
#[test]
fn a_right_click_on_a_pin_that_is_not_a_sound_opens_nothing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _put_back = PutBack(take_pin_for_a_test());
    // A framed card is not a sound's own card: the three facts the menu
    // is gated on are settled against the kind.
    let mut pin = sound_pin();
    pin.frame = PinFrame::Shaped;
    stand_pin(Some(pin));

    unsafe { open_pin_menu(HWND(0x1000 as *mut _), 20, 20) };
    assert!(
        pinned_menu_geometry().is_none(),
        "a pin that is not a sound opens no menu"
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
/// one the panel does not hold.
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

    // A point left of the panel, in the row of its first row: not the
    // panel, and not a row of any of the menu's.
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
    // menu's own right edge — the gap a panel is held
    // off the panel beside it — and its top is
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

/// A press on the Seek row is the door the flyout comes through: the
/// flyout opens beside the menu, and the menu itself stays up, its
/// three rows still showing. A hand that has come to press rather than
/// wait for the hover does not have to.
#[test]
fn a_press_on_the_seek_row_opens_its_flyout_and_keeps_the_menu_up() {
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
    assert!(paint.flyout.is_none(), "the menu opens on its main page");

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

    // The menu is still up, on its three rows, with the flyout up
    // beside it.
    let Some(paint) = pinned_menu_geometry() else {
        panic!("the menu is still up")
    };
    assert!(paint.flyout.is_some(), "the press opened the flyout");
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
    // the menu's first row, in the gap between the two
    // panels, or below the flyout: a hand off both
    // panels has left the menu.
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

/// The screen point a window-relative point of these tests' sound pin
/// is: its window's box begins at (100, 100), so a point on the menu is
/// taken to the screen the way the chrome tick takes the pointer.
fn on_screen(point: (i32, i32)) -> (i32, i32) {
    (100 + point.0, 100 + point.1)
}

/// The window-relative middle of a row of the main menu.
fn menu_row_point(paint: &PinMenuPaint, index: usize) -> (i32, i32) {
    (
        (paint.popup.panel.left + paint.popup.panel.right) / 2,
        paint.popup.rows_top
            + index as i32 * paint.popup.row_height
            + paint.popup.row_height / 2,
    )
}

/// Ask the chrome tick about a screen point, for the pin that is up,
/// answering whether anything changed.
fn tick(cursor: Option<(i32, i32)>, now: Instant) -> bool {
    pin_state()
        .and_then(|mut pinned| {
            let pin = pinned.pin_mut()?;
            Some(refresh_pin_chrome(pin, now, cursor))
        })
        .unwrap_or(false)
}

/// Whether the pin that is up has its Seek flyout up.
fn flyout_is_up() -> bool {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.menu.seek))
        .unwrap_or(false)
}

/// Hovering the Seek row opens its flyout after a short delay, and not
/// before: the first tick starts the timer, the flyout waits while the
/// pointer has not rested long enough, and it opens on the tick that
/// finds the delay elapsed.
#[test]
fn hovering_the_seek_row_opens_the_flyout_after_a_short_delay() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, false, true);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(sound_pin()));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.point = (20, 20);
    });
    let paint = pinned_menu_geometry().expect("the menu is up");
    let seek = on_screen(menu_row_point(&paint, 2));

    let t0 = Instant::now();

    // The first tick the pointer is on the Seek row starts the timer;
    // the flyout is not up yet, however many ticks pass before the
    // delay is reached.
    tick(Some(seek), t0);
    assert!(!flyout_is_up(), "the flyout waits for the hover delay");
    tick(Some(seek), t0 + Duration::from_millis(100));
    assert!(
        !flyout_is_up(),
        "the flyout is not up before the delay has elapsed"
    );

    // Once the pointer has rested there for the delay, it opens.
    assert!(
        tick(Some(seek), t0 + Duration::from_millis(200)),
        "the flyout opening is a change worth a repaint"
    );
    assert!(flyout_is_up(), "the flyout opens after the delay");
}

/// Leaving the Seek row before the delay has elapsed cancels the timer:
/// a hand crossing the row on its way somewhere else opens nothing, and
/// waiting there afterwards does not open it either.
#[test]
fn leaving_the_seek_row_before_the_delay_cancels_its_flyout() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, false, true);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(sound_pin()));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.point = (20, 20);
    });
    let paint = pinned_menu_geometry().expect("the menu is up");
    let seek = on_screen(menu_row_point(&paint, 2));
    let shuffle = on_screen(menu_row_point(&paint, 0));

    let t0 = Instant::now();
    tick(Some(seek), t0);
    // The pointer leaves the row before the delay: the timer is cleared.
    tick(Some(shuffle), t0 + Duration::from_millis(50));
    assert!(!flyout_is_up(), "it did not open before the delay");

    // And it never opens on this rest, however long it lasts.
    tick(Some(shuffle), t0 + Duration::from_millis(400));
    assert!(
        !flyout_is_up(),
        "leaving the Seek row before the delay cancels the flyout"
    );
}

/// A flyout that is up stays up while the pointer is over the Seek row,
/// the gap between the panels, or the flyout itself, and goes when the
/// pointer hovers another main-menu row or leaves the menu entirely.
#[test]
fn the_flyout_stays_up_across_the_gap_and_hides_on_another_row() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, false, true);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(sound_pin()));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.point = (20, 20);
    });

    // Open the flyout by resting on the Seek row past the delay.
    let paint = pinned_menu_geometry().expect("the menu is up");
    let seek = on_screen(menu_row_point(&paint, 2));
    let t0 = Instant::now();
    tick(Some(seek), t0);
    tick(Some(seek), t0 + Duration::from_millis(200));
    assert!(flyout_is_up(), "the flyout opened");

    // Over the flyout itself: it stays.
    let paint = pinned_menu_geometry().expect("the menu is up");
    let flyout = paint.flyout.as_ref().expect("the flyout is up");
    let on_flyout = on_screen((
        (flyout.popup.panel.left + flyout.popup.panel.right) / 2,
        flyout.popup.rows_top + flyout.popup.row_height / 2,
    ));
    tick(Some(on_flyout), t0 + Duration::from_millis(300));
    assert!(flyout_is_up(), "the flyout stays up under the pointer");

    // In the gap between the two panels, at the Seek row's own height:
    // it stays up across it.
    let paint = pinned_menu_geometry().expect("the menu is up");
    let flyout = paint.flyout.as_ref().expect("the flyout is up");
    let gap = on_screen((
        (paint.popup.panel.right + flyout.popup.panel.left) / 2,
        paint.popup.rows_top + 2 * paint.popup.row_height + paint.popup.row_height / 2,
    ));
    tick(Some(gap), t0 + Duration::from_millis(400));
    assert!(
        flyout_is_up(),
        "the flyout stays up while the pointer crosses the gap"
    );

    // On the Shuffle Mode row: it hides.
    let shuffle = on_screen(menu_row_point(&paint, 0));
    assert!(
        tick(Some(shuffle), t0 + Duration::from_millis(500)),
        "the flyout going is a change worth a repaint"
    );
    assert!(
        !flyout_is_up(),
        "the flyout hides when the pointer hovers another row"
    );

    // Opened again, then the pointer leaves the menu entirely: it hides.
    let seek = on_screen(menu_row_point(&paint, 2));
    tick(Some(seek), t0 + Duration::from_millis(600));
    tick(Some(seek), t0 + Duration::from_millis(800));
    assert!(flyout_is_up(), "the flyout opened again");
    tick(Some((-1000, -1000)), t0 + Duration::from_millis(900));
    assert!(
        !flyout_is_up(),
        "the flyout hides when the pointer leaves the menu entirely"
    );
}

/// The row under the pointer is the row the paint washes: the main
/// menu's own row where the pointer is on the main menu, and the
/// flyout's row where it is on the flyout.
#[test]
fn the_row_under_the_pointer_is_the_row_that_is_washed() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, false, true);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(sound_pin()));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.point = (20, 20);
    });
    let paint = pinned_menu_geometry().expect("the menu is up");
    let loop_row = on_screen(menu_row_point(&paint, 1));
    let t0 = Instant::now();

    assert!(
        tick(Some(loop_row), t0),
        "the row under the pointer is a change worth a repaint"
    );
    let paint = pinned_menu_geometry().expect("the menu is up");
    assert_eq!(
        paint.hover,
        Some(1),
        "the Loop row is the one the pointer is on"
    );
    assert!(
        paint.flyout.is_none(),
        "the flyout is down on the main page"
    );

    // And on the flyout: its own row is the one washed.
    with_pin(|pin| pin.menu.seek = true);
    let paint = pinned_menu_geometry().expect("the menu is up");
    let flyout = paint.flyout.as_ref().expect("the flyout is up");
    let choice = on_screen((
        (flyout.popup.panel.left + flyout.popup.panel.right) / 2,
        flyout.popup.rows_top + 2 * flyout.popup.row_height + flyout.popup.row_height / 2,
    ));
    tick(Some(choice), t0 + Duration::from_millis(50));
    let paint = pinned_menu_geometry().expect("the menu is up");
    assert_eq!(
        paint.flyout.as_ref().expect("the flyout is up").hover,
        Some(2),
        "the flyout's third row is the one the pointer is on"
    );
}

/// A press that landed on another window is closed from the hook side: the ask the
/// Explorer hook publishes is taken by the loop's own tick, which puts the whole menu
/// away — both panels — and leaves the pin standing. A tick with no such ask leaves a
/// menu that is up exactly where it is.
#[test]
fn the_loop_takes_a_foreign_press_ask_and_puts_the_menu_away() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, false, true);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(sound_pin()));

    // Any ask another test left behind is taken before this one asks its own.
    take_pin_menu_dismiss();

    // A menu that is up with its flyout beside it: both panels are what the ask puts away.
    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.seek = true;
    });
    assert!(pinned_menu_geometry().is_some(), "the menu is up");

    // A tick that no press reached leaves it where it is.
    assert!(
        !take_pin_menu_dismiss(),
        "with no ask there is nothing for the loop to close"
    );
    assert!(pinned_menu_geometry().is_some(), "so the menu stays up");

    // The hook's ask for a press on another window, taken on the loop's tick.
    request_pin_menu_dismiss();
    assert!(
        take_pin_menu_dismiss(),
        "the ask is taken and the menu is the thing it closes"
    );
    assert!(pinned_menu_geometry().is_none(), "both panels are away");
    assert!(pinned(), "and the pin that was standing is left standing");
}

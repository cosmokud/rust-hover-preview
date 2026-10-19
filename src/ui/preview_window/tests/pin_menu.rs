use super::*;

/// The three settings the card's menu holds, put where a test
/// asks and put back when the guard is dropped, so a test beside
/// this one is answered with the machine's own (the house
/// pattern: `SeekSettings` in `pin_audio_seek`).
struct MenuSettings {
    seek: AudioSeek,
    end_of_file: PinEndOfFile,
    shuffle: bool,
}

impl MenuSettings {
    fn put(seek: AudioSeek, end_of_file: PinEndOfFile, shuffle: bool) -> Self {
        let mut config = CONFIG.lock().expect("the configuration");
        let was = MenuSettings {
            seek: config.pin_mode_audio_seek,
            end_of_file: config.pin_mode_audio_end_of_file,
            shuffle: config.pin_mode_audio_shuffle,
        };
        config.pin_mode_audio_seek = seek;
        config.pin_mode_audio_end_of_file = end_of_file;
        config.pin_mode_audio_shuffle = shuffle;

        was
    }
}

impl Drop for MenuSettings {
    fn drop(&mut self) {
        if let Ok(mut config) = CONFIG.lock() {
            config.pin_mode_audio_seek = self.seek;
            config.pin_mode_audio_end_of_file = self.end_of_file;
            config.pin_mode_audio_shuffle = self.shuffle;
        }
    }
}

/// The app's own configuration file while a test has it: a folder of the test's own holding
/// it, taken away when the guard is dropped. A press on a menu's check row writes the
/// configuration the way every setting this app's own menus write is, and the machine's own
/// file is the one a run of the app watches every second and takes what it finds in — so a
/// test that pressed a row over the machine's own would change what that run had
/// (`ConfigFolderOfItsOwn`). The file is one file, so the guard holds the turn to have it
/// until the test is done.
struct ConfigFilePutBack {
    _folder: crate::config::config::ConfigFolderOfItsOwn,
}

impl ConfigFilePutBack {
    fn take() -> Self {
        Self {
            _folder: crate::config::config::ConfigFolderOfItsOwn::take(),
        }
    }
}

/// The media slot stood at a kind for the length of a test and put back when the guard is
/// dropped. A pinned window's menu holds the rows its kind has — the hold and the switch are
/// a film's or a sound's — and the kind is read off the slot rather than off the pin (see
/// `bubble_menu_rows`), so a test of them has to say which kind is on screen.
struct MediaKindPutBack {
    was: Option<MediaData>,
}

impl MediaKindPutBack {
    fn put(kind: MediaType) -> Self {
        let was = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
        let media = static_image_media(Arc::new(ImageFrame::new(vec![0u8; 4], 1, 1, 0)), kind);
        if let Ok(mut slot) = CURRENT_MEDIA.lock() {
            *slot = Some(media);
        }

        Self { was }
    }
}

impl Drop for MediaKindPutBack {
    fn drop(&mut self) {
        if let Ok(mut slot) = CURRENT_MEDIA.lock() {
            *slot = self.was.take();
        }
    }
}

/// Whether a text preview is on screen, stood where a test asks and put back when the guard is
/// dropped: the one fact that tells a pinned page's own `Select All` / `Copy` menu from the
/// window's, and a fact the preview loop publishes rather than a test's to leave standing (see
/// `TEXT_PREVIEW_HOLDING`).
struct TextHoldingPutBack(bool);

impl TextHoldingPutBack {
    fn put(holding: bool) -> Self {
        Self(TEXT_PREVIEW_HOLDING.swap(holding, Ordering::AcqRel))
    }
}

impl Drop for TextHoldingPutBack {
    fn drop(&mut self) {
        TEXT_PREVIEW_HOLDING.store(self.0, Ordering::Release);
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

/// The bubble a collapsed pin leaves, put up where a test asks and taken away again when the
/// guard is dropped: its handle and the menu it holds are slots for the whole process, and a
/// test that leaves either filled is a test every test after it stands inside (see
/// `PIN_BUBBLE_HWND` and `BUBBLE_MENU`).
struct BubblePutBack;

impl BubblePutBack {
    fn put(hwnd: HWND) -> Self {
        PIN_BUBBLE_HWND.store(hwnd.0 as isize, Ordering::SeqCst);
        Self
    }
}

impl Drop for BubblePutBack {
    fn drop(&mut self) {
        PIN_BUBBLE_HWND.store(0, Ordering::SeqCst);
        if let Ok(mut menu) = BUBBLE_MENU.lock() {
            *menu = None;
        }
    }
}

/// A pin showing a picture this app draws itself, big enough that a right-click on it leaves the
/// whole of its menu's room on the screen beside it — the kind that shares its menu with a
/// bubble, which a sound's card does not (see `pin_shows_an_audio_card`).
fn a_picture_whose_menu_fits_the_screen() -> PinnedPreview {
    let mut pin = PinnedPreview::for_test();
    pin.content = (100, 100, 500, 420);
    pin
}

/// A pin stood up for a test showing a sound's card big enough that a right-click on it
/// leaves the whole of the menu's room on the screen beside it: the panel is placed against
/// the display's own work area rather than against the card's box, so the card is no longer
/// asked to be tall enough to hold its own menu — and these tests keep a card that would
/// have been, which is the shape a card on a real machine has (see
/// `pin_chrome::menu_popup_from_point`).
fn a_sound_card_whose_menu_fits_the_screen() -> PinnedPreview {
    PinnedPreview {
        content: (100, 100, 500, 420),
        ..sound_pin()
    }
}

/// The rows the menu holds are the settings it is the face of, and what
/// the pin does with them: the two ways out of the pin,
/// the rule, the playback of the file in it, the rule again,
/// and then the settings — the two doors, neither carrying
/// a mark and both carrying the arrow a door carries — with the shuffle
/// between them carrying the check it is the face of, the end-of-file
/// flyout is the three ends with the one the pin plays by carrying its disc,
/// and the seek flyout is the four choices, in the order the tray
/// lists them in, with the one the pin plays by carrying its disc.
#[test]
fn the_menu_holds_the_settings_it_is_the_face_of() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Middle, PinEndOfFile::Continue, true);

    let top = menu_rows();
    let labels: Vec<&str> = top.iter().map(|row| row.label.as_str()).collect();
    assert_eq!(
        labels,
        [
            "To Bubble",
            "To Taskbar",
            "Close",
            "",
            "Play",
            "Next File",
            "Previous File",
            "",
            "Shuffle",
            "End of File",
            "Seek",
        ]
    );
    assert!(
        matches!(top[3].mark, pin_chrome::MenuMark::Separator),
        "the rule between the ways out of the pin and the playback is a rule, not a row"
    );
    assert!(
        matches!(
            top[row_at(&top, "End of File")].mark,
            pin_chrome::MenuMark::Door
        ),
        "the end-of-file row opens the choices, so it is a door rather than an answer"
    );
    assert!(
        matches!(
            top[row_at(&top, "Shuffle")].mark,
            pin_chrome::MenuMark::Check(true)
        ),
        "and the row between the doors is the check the shuffle is the face of"
    );
    assert!(
        matches!(top[row_at(&top, "Seek")].mark, pin_chrome::MenuMark::Door),
        "and so is the seek row"
    );

    let ends = end_of_file_rows();
    let labels: Vec<&str> = ends.iter().map(|row| row.label.as_str()).collect();
    assert_eq!(labels, ["Pause", "Loop", "Jump"]);
    let marked = ends
        .iter()
        .position(|row| matches!(row.mark, pin_chrome::MenuMark::Bullet));
    assert_eq!(
        marked,
        Some(2),
        "the pin steps on at the end of a file, so that choice is the one marked"
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
    // marks the defaults: a loop at the end of a file, the walk in
    // the folder's own order, and the
    // beginning the way a pinned sound starts.
    let _settings = MenuSettings::put(
        DEFAULT_PIN_MODE_AUDIO_SEEK,
        DEFAULT_PIN_MODE_AUDIO_END_OF_FILE,
        DEFAULT_PIN_MODE_AUDIO_SHUFFLE,
    );
    let top = menu_rows();
    assert!(
        matches!(
            top[row_at(&top, "Shuffle")].mark,
            pin_chrome::MenuMark::Check(false)
        ),
        "the shuffle starts off, so its row is unchecked"
    );
    let ends = end_of_file_rows();
    let marked = ends
        .iter()
        .position(|row| matches!(row.mark, pin_chrome::MenuMark::Bullet));
    assert_eq!(
        marked,
        Some(1),
        "a pinned sound loops, so that choice is the one marked"
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

/// The hold row is named for what it will do rather than for what the row already does,
/// the way the bubble's own hold is: `Pause` while the sound behind the card is playing
/// and `Play` while it is held — the same answer the card's own play button gives, and
/// the row holds its place in the menu whichever of the two it says (see `menu_rows` and
/// `a_sound_is_playing`).
///
/// The media slot is the one place the answer is read from, so it is taken for the length
/// of the test and put back what it was: a player of this app's is the process in it, and
/// a test notes one rather than beginning one (see `PlayerProcess::Noted`).
#[test]
fn the_hold_row_is_named_for_what_it_will_do() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    // Nothing behind the card: the row offers to set the sound going.
    let rows = menu_rows();
    assert_eq!(
        rows[row_at(&rows, "Play")].label,
        "Play",
        "with nothing playing the hold row is a play"
    );
    assert!(
        rows.iter().all(|row| row.label != "Pause"),
        "and no row of it is a pause"
    );

    // A player of this app's behind the sound: the row offers to hold it, in the same
    // place it sat before — only what it says moves.
    let mut media = static_image_media(
        Arc::new(ImageFrame::new(vec![0u8; 4], 1, 1, 0)),
        MediaType::Audio,
    );
    media.video_process = Some(PlayerProcess::Noted(4));
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = Some(media);
    }
    let rows = menu_rows();
    assert_eq!(
        rows[row_at(&rows, "Pause")].label,
        "Pause",
        "with a player running the hold row is a pause"
    );

    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = previous_media;
    }
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

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    assert!(
        pinned_menu_geometry().is_none(),
        "a menu that has not been asked for is not painted"
    );

    let hwnd = HWND(0x1000 as *mut _);

    // A point that leaves the panel room: the panel's own top-left is the
    // point itself, the moment the press has landed.
    let (x, y) = on_screen((20, 20));
    unsafe { open_pin_menu(hwnd, x, y, PinMenuPage::Chrome) };
    let Some(paint) = pinned_menu_geometry() else {
        panic!("a menu that is up is painted");
    };
    assert_eq!(
        (paint.popup.panel.left, paint.popup.panel.top),
        (x, y),
        "the panel's top-left is the point the right-click landed"
    );

    // It opens on its main page: the ways out of the pin, the rule, the playback, the rule
    // again, and the Shuffle check between the two doors.
    let labels: Vec<&str> = paint.rows.iter().map(|row| row.label.as_str()).collect();
    assert_eq!(
        labels,
        [
            "To Bubble",
            "To Taskbar",
            "Close",
            "",
            "Play",
            "Next File",
            "Previous File",
            "",
            "Shuffle",
            "End of File",
            "Seek",
        ]
    );
    assert!(
        matches!(
            paint.rows[row_at(&paint.rows, "End of File")].mark,
            pin_chrome::MenuMark::Door
        ),
        "the end-of-file row opens the choices, so it is a door rather than an answer"
    );
    assert!(
        matches!(
            paint.rows[row_at(&paint.rows, "Shuffle")].mark,
            pin_chrome::MenuMark::Check(_)
        ),
        "and the row between the doors is the check the shuffle is the face of"
    );
    assert!(
        matches!(
            paint.rows[row_at(&paint.rows, "Seek")].mark,
            pin_chrome::MenuMark::Door
        ),
        "and so is the seek row"
    );
    assert!(
        paint.flyout.is_none(),
        "the main page is the menu's own rows"
    );

    // A second right-click moves the panel to the new point, and the menu
    // stays up.
    let (x, y) = on_screen((30, 40));
    unsafe { open_pin_menu(hwnd, x, y, PinMenuPage::Chrome) };
    let Some(paint) = pinned_menu_geometry() else {
        panic!("the menu is still up");
    };
    assert_eq!(
        (paint.popup.panel.left, paint.popup.panel.top),
        (x, y),
        "a second right-click moves the panel to the new point"
    );

    // A right-click near the far corner of the screen would put the panel off
    // it, so the placement holds it inside: the room a panel is held in is the
    // display's own work area and not the card's box, which is what lets a menu
    // taller than the card keep every one of its rows (see `pin_menu_window`).
    let screen = work_area_at(x, y);
    unsafe {
        open_pin_menu(
            hwnd,
            screen.right - 2,
            screen.bottom - 2,
            PinMenuPage::Chrome,
        )
    };
    let Some(paint) = pinned_menu_geometry() else {
        panic!("the menu is still up");
    };
    assert!(
        paint.popup.panel.right <= screen.right && paint.popup.panel.bottom <= screen.bottom,
        "a panel asked for near the corner of the screen is held inside it"
    );
    assert!(
        paint.popup.panel.left >= screen.left && paint.popup.panel.top >= screen.top,
        "and it is held inside the work area's own top-left"
    );
}

/// A right-click on a pin that is not showing a sound's card opens the window's own menu: the
/// same rows a bubble's menu holds, and no panel beside it, because nothing on it is a setting a
/// flyout hangs off (see `bubble_menu_rows` and `pinned_menu_paint`).
#[test]
fn a_right_click_on_a_pin_that_is_not_a_sound_opens_the_window_menu() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _put_back = PutBack(take_pin_for_a_test());
    let _kind = MediaKindPutBack::put(MediaType::StaticImage);
    // A framed card is not a sound's own card: the two menus are told apart by the
    // pin's own three facts rather than by the media's kind (see `pin_shows_an_audio_card`).
    stand_pin(Some(a_picture_whose_menu_fits_the_screen()));

    unsafe { open_pin_menu(HWND(0x1000 as *mut _), 20, 20, PinMenuPage::Chrome) };
    let Some(paint) = pinned_menu_geometry() else {
        panic!("a pin that is not a sound opens the window's own menu")
    };
    let labels: Vec<&str> = paint.rows.iter().map(|row| row.label.as_str()).collect();
    assert_eq!(
        labels,
        [
            "Next File",
            "Previous File",
            "",
            "Minimize to Bubble",
            "Minimize to Taskbar",
            "",
            "To Taskbar",
            "Close",
        ],
        "the window's menu is the menu a bubble holds"
    );
    assert!(
        paint.flyout.is_none(),
        "no row of it opens a panel beside it"
    );
}

/// Which page of its own menu a right-click asks a pinned window for: the window's rows on a
/// sound's card, on the body and the title bar of a picture, and on the title bar of a pinned
/// page; the page's own two commands over the text of a pinned page; and nothing at all on the
/// transport bar of a film, which is the player's own strip along the bottom (see
/// `pin_right_click_page`).
#[test]
fn a_right_click_asks_for_the_page_its_row_of_the_window_holds() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _put_back = PutBack(take_pin_for_a_test());
    let _no_page = TextHoldingPutBack::put(false);

    // A sound's card is all of it: no caption above the card and no bar under it, so
    // every row of the window is the menu's.
    stand_pin(Some(sound_pin()));
    assert_eq!(
        pin_right_click_page(0),
        Some(PinMenuPage::Chrome),
        "the first row of a sound's card is the menu's"
    );
    assert_eq!(
        pin_right_click_page(200),
        Some(PinMenuPage::Chrome),
        "and so is a row well down it"
    );

    // A picture: the title bar above its media is the window's, and so is the picture under it,
    // which is the window's own to draw.
    let picture = PinnedPreview::for_test();
    let caption = picture.caption;
    stand_pin(Some(picture));
    assert_eq!(
        pin_right_click_page(0),
        Some(PinMenuPage::Chrome),
        "the title bar of a picture is the menu's"
    );
    assert_eq!(
        pin_right_click_page(caption + 1),
        Some(PinMenuPage::Chrome),
        "and the picture under it too"
    );

    // A film's transport bar, which is the strip along the bottom of its window: the bar's own,
    // with the picture above it the window's.
    let mut film = PinnedPreview::for_test();
    film.transport_bar = true;
    film.transport_live = true;
    stand_pin(Some(film));
    let bar = pinned_transport_geometry().expect("a film's window carries a transport bar");
    assert_eq!(
        pin_right_click_page(bar.top),
        None,
        "the bar's own first row is the player's"
    );
    assert_eq!(
        pin_right_click_page(bar.top + bar.height - 1),
        None,
        "and so is its last"
    );
    assert_eq!(
        pin_right_click_page(bar.top - 1),
        Some(PinMenuPage::Chrome),
        "while the row above the bar is the window's"
    );

    // A page of text, with a text preview on screen: the text asks for the page's own page —
    // its two commands above the window's rows — and the title bar above it for the window's.
    let _page = TextHoldingPutBack::put(true);
    let mut page = PinnedPreview::for_test();
    page.caption = pinned_caption_height(96, Some(MediaType::Text));
    let caption = page.caption;
    stand_pin(Some(page));
    assert_eq!(
        pin_right_click_page(caption - 1),
        Some(PinMenuPage::Chrome),
        "the title bar of a pinned page is the menu's"
    );
    assert_eq!(
        pin_right_click_page(caption),
        Some(PinMenuPage::Text),
        "and the text under it asks for the page's own rows above them"
    );
    assert_eq!(
        pin_right_click_page(caption + 40),
        Some(PinMenuPage::Text),
        "wherever in the text the hand is"
    );
}

/// A right-click in the text of a pinned page opens the window's own menu with the page's own two
/// commands above the rows the pin shares with a bubble, in the same panel and the same design
/// (see `PinMenuPage` and `bubble_menu_rows`).
#[test]
fn a_pinned_page_holds_its_own_commands_above_the_windows_rows() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _put_back = PutBack(take_pin_for_a_test());
    let _kind = MediaKindPutBack::put(MediaType::Text);
    let _page = TextHoldingPutBack::put(true);
    stand_pin(Some(a_picture_whose_menu_fits_the_screen()));

    unsafe { open_pin_menu(HWND(0x1000 as *mut _), 20, 20, PinMenuPage::Text) };
    let Some(paint) = pinned_menu_geometry() else {
        panic!("a right-click in the text of a pinned page opens the window's own menu")
    };
    let labels: Vec<&str> = paint.rows.iter().map(|row| row.label.as_str()).collect();
    assert_eq!(
        labels,
        [
            "Select All",
            "Copy",
            "",
            "Next File",
            "Previous File",
            "",
            "Minimize to Bubble",
            "Minimize to Taskbar",
            "",
            "To Taskbar",
            "Close",
        ],
        "the page's own commands stand above the rows the pin shares with a bubble"
    );
    assert!(
        matches!(paint.rows[2].mark, pin_chrome::MenuMark::Separator),
        "and the rule between the two groups is drawn where the page's own rows end"
    );
    assert!(
        paint.flyout.is_none(),
        "no row of it opens a panel beside it"
    );
}

/// A press on a row a pinned window's menu shares with the bubble's leaves the ask that row
/// leaves there: the two walks and the close the same, and the one row the two walk differently
/// — `To Taskbar` — putting this pin away into the app's own place in the taskbar from the
/// window it is standing in (see `shared_menu_row_press` and `PinCommand::MinimizeTo`).
#[test]
fn a_press_on_a_row_a_pin_shares_with_a_bubble_leaves_its_ask() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _put_back = PutBack(take_pin_for_a_test());
    let _kind = MediaKindPutBack::put(MediaType::StaticImage);
    stand_pin(Some(a_picture_whose_menu_fits_the_screen()));

    // What each row is for, as the ask it leaves behind with the loop.
    let asks = [
        ("Next File", PinCommand::Next),
        ("Previous File", PinCommand::Previous),
        ("Close", PinCommand::Close),
        ("To Taskbar", PinCommand::MinimizeTo(PinMinimizeTo::Taskbar)),
    ];

    for (label, command) in asks {
        with_pin(|pin| {
            pin.menu.open = true;
            pin.menu.point = on_screen((20, 20));
        });
        // The queue is emptied first, so what is read is the ask this press left.
        take_pin_command();
        let paint = pinned_menu_geometry().expect("the menu is up");
        let (x, y) = menu_row_point(&paint, row_at(&paint.rows, label));
        assert!(
            unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
            "a hand on the {label} row is the menu's own"
        );
        assert_eq!(
            take_pin_command(),
            Some(command),
            "the {label} row asks for what it says"
        );
        assert!(
            pinned_menu_geometry().is_none(),
            "and the menu goes away with the answer"
        );
    }
}

/// A press on either of the two places a minimize can put a pin writes the setting the tray's
/// own `Pin Mode → Minimize` writes, for the pin that is up as much as for a bubble: the row is
/// the setting's face rather than an ask, so nothing is left with the loop and the menu goes
/// away with the pick (see `shared_menu_row_press`).
#[test]
fn a_press_on_a_minimize_pick_writes_the_setting_the_tray_writes() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _config_file = ConfigFilePutBack::take();
    let _put_back = PutBack(take_pin_for_a_test());
    let _kind = MediaKindPutBack::put(MediaType::StaticImage);
    let _target = PinMinimizeToPutBack::put(PinMinimizeTo::Bubble);
    stand_pin(Some(a_picture_whose_menu_fits_the_screen()));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.point = on_screen((20, 20));
    });
    take_pin_command();
    let paint = pinned_menu_geometry().expect("the menu is up");
    let (x, y) = menu_row_point(&paint, row_at(&paint.rows, "Minimize to Taskbar"));
    assert!(
        unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on a pick is the menu's own"
    );
    assert_eq!(
        CONFIG.lock().map(|config| config.pin_minimize_to).ok(),
        Some(PinMinimizeTo::Taskbar),
        "the pick wrote the place the tray's own row writes"
    );
    assert_eq!(
        take_pin_command(),
        None,
        "and nothing is left with the loop: a pick is a setting rather than an ask"
    );
    assert!(
        pinned_menu_geometry().is_none(),
        "and the menu goes away with the pick"
    );
}

/// The switch a pinned window's menu shares with the bubble's is turned over where it is pressed
/// and the panel stays up with the mark it now carries — there is nothing to pick from and the
/// mark is the answer (see `shared_menu_row_press`). A film's menu holds the switch because a
/// film is a kind with playback for it to hold back.
#[test]
fn the_switch_a_pin_shares_with_a_bubble_stays_up_with_its_mark() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _switches = PausePreviewSettings::stood_at(true, false);
    let _put_back = PutBack(take_pin_for_a_test());
    let _kind = MediaKindPutBack::put(MediaType::Video);
    let mut film = a_picture_whose_menu_fits_the_screen();
    // A film held where it stands, so that the hold row says the same thing on every machine
    // (see `bubble_is_playing`).
    film.transport.paused_at = Some(0.0);
    stand_pin(Some(film));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.point = on_screen((20, 20));
    });
    take_pin_command();
    let paint = pinned_menu_geometry().expect("the menu is up");
    assert_eq!(
        paint.rows[0].label, "Play",
        "a held film's menu offers to set it going"
    );
    let (x, y) = menu_row_point(&paint, row_at(&paint.rows, "Pause on Minimize"));
    assert!(
        unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on the switch is the menu's own"
    );
    assert_eq!(
        take_pin_command(),
        None,
        "a switch asks the loop for nothing at all"
    );
    let paint = pinned_menu_geometry().expect("the panel stays up over the switch");
    let row = row_at(&paint.rows, "Pause on Minimize");
    assert!(
        matches!(paint.rows[row].mark, pin_chrome::MenuMark::Check(false)),
        "the video's switch was on, so the press turned it off and the mark says so"
    );
}

/// A row of the panel that reaches across the window's own title bar is the row's: the panel is
/// over the chrome, so the press is answered the way the same row is answered lower down rather
/// than beginning the drag a title bar begins (see `pinned_press`).
#[test]
fn a_press_on_a_row_over_the_title_bar_is_the_rows_own() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _put_back = PutBack(take_pin_for_a_test());
    let _kind = MediaKindPutBack::put(MediaType::StaticImage);
    let mut pin = PinnedPreview::for_test();
    pin.content = (0, 0, 500, 420);
    let caption = pin.caption;
    stand_pin(Some(pin));

    with_pin(|pin| {
        pin.menu.open = true;
        // A panel at the window's own top-left corner, which is where a right-click in the
        // title bar there would have put it: its first rows lie in the title bar itself.
        pin.menu.point = (0, 0);
    });
    take_pin_command();
    let paint = pinned_menu_geometry().expect("the menu is up");
    let (x, y) = menu_row_point(&paint, row_at(&paint.rows, "Next File"));
    assert!(
        y < caption,
        "the row this test presses lies in the window's own title bar"
    );

    assert!(
        unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
        "a press on a row of the panel is the menu's"
    );
    assert_eq!(
        take_pin_command(),
        Some(PinCommand::Next),
        "the row the hand was on is the row that answered"
    );
    assert!(
        pinned_menu_geometry().is_none(),
        "and the menu went away with its answer"
    );
}

/// The painted menu carries the Seek row's arrow with the flyout
/// down — the arrow is a property of the main menu's panel, the
/// one placement answering it whether the flyout is up or not —
/// and the flyout's own panel and rows are carried only while the
/// flyout is up: the paint carries no flyout at all until it is,
/// which is what keeps the row's door sign up without its room.
#[test]
fn the_menu_carries_its_arrow_with_the_flyout_down() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    with_pin(|pin| pin.menu.open = true);
    let Some(paint) = pinned_menu_geometry() else {
        panic!("a menu that is up is painted");
    };

    // The arrow, carried while the flyout is down: not nothing,
    // the side the flyout opens on.
    assert_ne!(
        paint.popup.arrow,
        pin_chrome::MenuArrow::None,
        "the Seek row carries its arrow with the flyout down"
    );

    // The flyout's panel and rows are absent until the flyout
    // is up.
    assert!(
        paint.flyout.is_none(),
        "the flyout's panel is carried only while the flyout is up"
    );

    // And when the flyout is up, its rows are carried and the
    // arrow is the side its panel is on — the same arrow the
    // menu carried without it.
    let arrow = paint.popup.arrow;
    with_pin(|pin| pin.menu.seek = true);
    let Some(paint) = pinned_menu_geometry() else {
        panic!("the menu is still up");
    };
    let Some(flyout) = paint.flyout.as_ref() else {
        panic!("the flyout is up")
    };
    assert_eq!(
        flyout.rows.len(),
        4,
        "the flyout's rows are carried while the flyout is up"
    );
    assert_eq!(
        paint.popup.arrow, arrow,
        "the arrow does not move when the flyout comes up"
    );
    match paint.popup.arrow {
        pin_chrome::MenuArrow::Right => assert!(
            flyout.popup.panel.left >= paint.popup.panel.right,
            "the flyout opens on the arrow's side, right"
        ),
        pin_chrome::MenuArrow::Left => assert!(
            flyout.popup.panel.right <= paint.popup.panel.left,
            "the flyout opens on the arrow's side, left"
        ),
        pin_chrome::MenuArrow::None => {
            panic!("the Seek row carries an arrow while the menu is up")
        }
    }
}

/// A press on the Jump row is the door its flyout comes through: the
/// flyout opens beside the menu, and the menu itself stays up, its
/// two rows still showing. A hand that has come to press rather than
/// wait for the hover does not have to.
#[test]
fn a_press_on_the_jump_row_opens_its_flyout_and_keeps_the_menu_up() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    with_pin(|pin| pin.menu.open = true);
    let Some(paint) = pinned_menu_geometry() else {
        panic!("a menu that is up is painted");
    };
    assert!(paint.flyout.is_none(), "the menu opens on its main page");

    // The middle of the End of File row, named rather than numbered: the dialogue's door.
    let (x, y) = menu_row_point(&paint, row_at(&paint.rows, "End of File"));
    assert!(
        unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on the End of File row is the menu's own"
    );

    // The menu is still up, on its rows, with the flyout up beside it and hung on
    // the row the press landed on.
    let Some(paint) = pinned_menu_geometry() else {
        panic!("the menu is still up")
    };
    let Some(flyout) = paint.flyout.as_ref() else {
        panic!("the press opened the flyout")
    };
    assert_eq!(
        flyout.popup.panel.top,
        paint.popup.bands[row_at(&paint.rows, "End of File")].top,
        "the End of File flyout is top-aligned with the row it was opened from"
    );
    assert_eq!(
        flyout.popup.panel.left,
        paint.popup.panel.right + 4,
        "and it sits beside the menu it came from"
    );
    assert_ne!(
        paint.popup.arrow,
        pin_chrome::MenuArrow::None,
        "the End of File row carries the arrow a door carries, whether the flyout is up or not"
    );
    let labels: Vec<&str> = paint.rows.iter().map(|row| row.label.as_str()).collect();
    assert_eq!(
        labels,
        [
            "To Bubble",
            "To Taskbar",
            "Close",
            "",
            "Play",
            "Next File",
            "Previous File",
            "",
            "Shuffle",
            "End of File",
            "Seek",
        ]
    );
}

/// A press on the `Shuffle` row turns the check it carries over, writes it down, and
/// leaves the panel up showing it: the row is a check toggle with no flyout of its own,
/// so it is the one row of this menu a press answers without putting the menu away — and
/// what the check now says is the row's own answer. The setting is written to `config.ini`
/// the way every setting this app's own menus write is, which is why the file the press
/// writes is put back what it was when the test is done.
#[test]
fn a_press_on_the_shuffle_row_turns_it_over_and_writes_it_down() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _file = ConfigFilePutBack::take();
    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    with_pin(|pin| pin.menu.open = true);
    let Some(paint) = pinned_menu_geometry() else {
        panic!("a menu that is up is painted");
    };

    // The middle of the Shuffle row, the check between the two doors.
    let (x, y) = menu_row_point(&paint, row_at(&paint.rows, "Shuffle"));
    assert!(
        unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on the Shuffle row is the menu's own"
    );

    // The setting is the opposite of what it was, in the
    // configuration the press holds...
    assert!(
        CONFIG
            .lock()
            .expect("the configuration")
            .pin_mode_audio_shuffle,
        "the walk was in the folder's own order, so the press turned the shuffle on"
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

    // The rows show the setting as it now is: the check row is ticked.
    let rows = menu_rows();
    assert!(
        matches!(
            rows[row_at(&rows, "Shuffle")].mark,
            pin_chrome::MenuMark::Check(true)
        ),
        "the Shuffle row is ticked now"
    );

    // And the panel is still up, because the row answers where it stands rather
    // than opening anything: the mark it carries is the answer.
    let Some(paint) = pinned_menu_geometry() else {
        panic!("the panel is left up by the press that answered it")
    };
    assert!(
        paint.flyout.is_none(),
        "and no flyout comes with it: the check row opens nothing"
    );
}

/// A menu standing with the whole screen to its right opens its flyout to the right of
/// it, and one asked for at the screen's own right edge opens it to the left: the room a
/// flyout is placed in is the display's own work area and not the card's box, so what
/// says which side it opens on is where the menu stands on the screen — which is also why
/// a card no wider than the menu still has a flyout to open (see `pinned_menu_paint`).
#[test]
fn the_flyout_opens_right_of_the_menu_unless_the_screens_right_edge_says_otherwise() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    let hwnd = HWND(0x1000 as *mut _);
    let screen = work_area_at(0, 0);

    // With the screen open to the right of it, the flyout is placed right of the menu it
    // came from.
    unsafe { open_pin_menu(hwnd, 120, 120, PinMenuPage::Chrome) };
    let paint = pinned_menu_geometry().expect("the menu is up");
    let (x, y) = menu_row_point(&paint, row_at(&paint.rows, "Seek"));
    assert!(
        unsafe { pinned_menu_press(hwnd, x, y) },
        "a press on the Seek row is the menu's own"
    );
    let paint = pinned_menu_geometry().expect("the menu is up");
    let flyout = paint.flyout.as_ref().expect("the flyout is up");
    assert_eq!(
        paint.popup.arrow,
        pin_chrome::MenuArrow::Right,
        "the Seek row's door points at the flyout on its right"
    );
    assert!(
        flyout.popup.panel.left >= paint.popup.panel.right,
        "and the flyout stands right of the menu it came from"
    );

    // Asked for at the screen's own right edge, it flips to the menu's left — the standard
    // Windows behaviour — with the menu itself left where the click put it.
    unsafe { open_pin_menu(hwnd, screen.right - 2, screen.top + 20, PinMenuPage::Chrome) };
    let paint = pinned_menu_geometry().expect("the menu is up");
    let (x, y) = menu_row_point(&paint, row_at(&paint.rows, "Seek"));
    assert!(
        unsafe { pinned_menu_press(hwnd, x, y) },
        "a press on the Seek row is the menu's own"
    );
    let paint = pinned_menu_geometry().expect("the menu is up");
    let flyout = paint.flyout.as_ref().expect("the flyout is up");
    assert_eq!(
        paint.popup.arrow,
        pin_chrome::MenuArrow::Left,
        "the Seek row's door points at the flyout on its left near the screen's own edge"
    );
    assert!(
        flyout.popup.panel.right <= paint.popup.panel.left,
        "and the flyout stands left of the menu it came from"
    );
    assert!(
        paint.popup.panel.right <= screen.right,
        "and the menu itself is still held inside the screen"
    );
}

/// A press on one of the end-of-file flyout's choice rows makes that end of file the one
/// the pin plays by from then on, and the menu goes away with the answer — the same bargain
/// the seek flyout's rows make (see `the_flyout_s_rows_answer_presses`).
#[test]
fn a_press_on_an_end_of_file_choice_sets_it_and_puts_the_menu_away() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _file = ConfigFilePutBack::take();
    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.end_of_file = true;
    });
    let Some(paint) = pinned_menu_geometry() else {
        panic!("a menu that is up is painted");
    };
    let Some(flyout) = paint.flyout.as_ref() else {
        panic!("the end-of-file flyout is up")
    };

    // The middle of the flyout's `Jump` row, which is the end the pin is not
    // playing by: Loop is on, so the press chooses Jump. The row is named by what it
    // says, and the band it is aimed with is the one it was laid out in (see
    // `flyout_row_point`).
    let jump_row = flyout
        .rows
        .iter()
        .position(|row| row.label == "Jump")
        .expect("the end-of-file flyout lists Jump");
    let (x, y) = flyout_row_point(flyout, jump_row);
    assert!(
        unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on a row of the flyout is the menu's own"
    );

    assert_eq!(
        CONFIG
            .lock()
            .expect("the configuration")
            .pin_mode_audio_end_of_file,
        PinEndOfFile::Continue,
        "the press chose Jump, so the pin steps on at the end of a file now"
    );
    if crate::config::config::AppConfig::config_path().is_some() {
        assert_eq!(
            crate::config::config::AppConfig::load().pin_mode_audio_end_of_file,
            PinEndOfFile::Continue,
            "the file the press wrote holds the choice made"
        );
    }

    let rows = end_of_file_rows();
    let marked = rows
        .iter()
        .position(|row| matches!(row.mark, pin_chrome::MenuMark::Bullet));
    assert_eq!(
        marked
            .and_then(|index| rows.get(index))
            .map(|row| row.label.as_str()),
        Some("Jump"),
        "Jump is the end of file in force now"
    );
    assert!(
        pinned_menu_geometry().is_none(),
        "and the whole menu goes away with the choice"
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

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

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
    assert!(pinned_menu_geometry().is_none(), "so the panel is put away");
}

/// The rows that do something rather than answer a question leave their ask with the
/// preview loop, and put the menu away with it: the two ways out say where the pin goes, the
/// close takes it down, and the three playback rows ask for the same the card's own buttons
/// and the arrow keys do (see `PinCommand`).
#[test]
fn the_rows_that_do_something_leave_their_ask_with_the_loop() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    // What each row is for, as the ask it leaves behind with the loop.
    let asks = [
        ("To Bubble", PinCommand::MinimizeTo(PinMinimizeTo::Bubble)),
        ("To Taskbar", PinCommand::MinimizeTo(PinMinimizeTo::Taskbar)),
        ("Close", PinCommand::Close),
        ("Play", PinCommand::TogglePlayback),
        ("Next File", PinCommand::Next),
        ("Previous File", PinCommand::Previous),
    ];

    for (label, command) in asks {
        with_pin(|pin| {
            pin.menu.open = true;
            pin.menu.point = on_screen((20, 20));
        });
        // The queue is emptied first, so what is read is the ask this press left.
        take_pin_command();
        let paint = pinned_menu_geometry().expect("the menu is up");
        let (x, y) = menu_row_point(&paint, row_at(&paint.rows, label));
        assert!(
            unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
            "a hand on the {label} row is the menu's own"
        );
        assert_eq!(
            take_pin_command(),
            Some(command),
            "the {label} row asks for what it says"
        );
        assert!(
            pinned_menu_geometry().is_none(),
            "and the menu goes away with the card it was opened over"
        );
    }
}

/// A rule is not a row a hand answers, so a press on one is nothing at all: no ask is left
/// behind and the menu stays exactly where it is.
#[test]
fn a_press_on_a_rule_is_nothing_at_all() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.point = on_screen((20, 20));
    });
    take_pin_command();
    let paint = pinned_menu_geometry().expect("the menu is up");
    // The first rule of the menu, between the ways out of the pin and the playback.
    let rule = paint
        .rows
        .iter()
        .position(|row| row.label.is_empty())
        .expect("the menu holds a rule between two of its groups");
    let (x, y) = menu_row_point(&paint, rule);
    assert!(
        unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on a rule is the menu's own"
    );
    assert!(
        pinned_menu_geometry().is_some(),
        "and the menu stays up over a rule"
    );
    assert_eq!(
        take_pin_command(),
        None,
        "and a rule asks for nothing at all"
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

    let _settings = MenuSettings::put(AudioSeek::Middle, PinEndOfFile::Continue, true);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

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
        paint.popup.bands[row_at(&paint.rows, "Seek")].top,
        "the flyout is top-aligned with the Seek row"
    );

    // The main menu is still inside the screen and
    // uncovered, and the flyout is inside the screen,
    // which is what keeps it reachable: the room a
    // panel is held in is the display's own work area
    // and not the card's box, which is what lets a
    // menu taller than the card keep every row (see
    // `pinned_menu_paint`).
    let screen = work_area_at(0, 0);
    assert!(
        paint.popup.panel.left >= screen.left,
        "the menu is inside the screen"
    );
    assert!(
        paint.popup.panel.right <= screen.right,
        "and uncovered to its right"
    );
    assert!(
        flyout.popup.panel.right <= screen.right,
        "the flyout is inside the screen"
    );
    assert!(
        flyout.popup.panel.bottom <= screen.bottom,
        "and so is its bottom"
    );

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

    let _settings = MenuSettings::put(AudioSeek::Remember, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

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
    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

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
    let y = flyout.popup.rows_top + 3 * flyout.popup.row_height + flyout.popup.row_height / 2;
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
/// two rows still showing. A hand that has come to press rather than
/// wait for the hover does not have to.
#[test]
fn a_press_on_the_seek_row_opens_its_flyout_and_keeps_the_menu_up() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    with_pin(|pin| pin.menu.open = true);
    let Some(paint) = pinned_menu_geometry() else {
        panic!("a menu that is up is painted");
    };
    assert!(paint.flyout.is_none(), "the menu opens on its main page");

    // The middle of the Seek row, the last of the menu's own rows.
    let (x, y) = menu_row_point(&paint, row_at(&paint.rows, "Seek"));
    assert!(
        unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on the Seek row is the menu's own"
    );

    // The menu is still up, on its own rows, with the flyout up
    // beside it.
    let Some(paint) = pinned_menu_geometry() else {
        panic!("the menu is still up")
    };
    assert!(paint.flyout.is_some(), "the press opened the flyout");
    let labels: Vec<&str> = paint.rows.iter().map(|row| row.label.as_str()).collect();
    assert_eq!(
        labels,
        [
            "To Bubble",
            "To Taskbar",
            "Close",
            "",
            "Play",
            "Next File",
            "Previous File",
            "",
            "Shuffle",
            "End of File",
            "Seek",
        ]
    );
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

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

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

/// The screen point a point of these tests' sound card is: the card's own box begins at
/// (100, 100), so a right-click on it at a point of its own is a point on the screen the
/// way the window's procedure takes one (see `client_to_screen`). What a panel is read
/// back as is already on the screen, because that is where it is drawn — in a window of
/// its own over the screen rather than inside the card (`pin_menu_window`).
fn on_screen(point: (i32, i32)) -> (i32, i32) {
    (100 + point.0, 100 + point.1)
}

/// The middle of a row of the main menu, taken from the band the panel laid
/// the row out in rather than from arithmetic on the row height: a rule between two groups of
/// rows takes half of one, so the row under one is not a row's height further down (see
/// `pin_chrome::MenuRowBand`). It is a point on the screen, because that is where the
/// panel is drawn (see `pin_menu_window`).
fn menu_row_point(paint: &PinMenuPaint, index: usize) -> (i32, i32) {
    let band = paint
        .popup
        .bands
        .iter()
        .find(|band| band.answers == index)
        .unwrap_or_else(|| panic!("the menu lays the row a hand counts {index} out in a band"));

    (
        (paint.popup.panel.left + paint.popup.panel.right) / 2,
        band.top + band.height / 2,
    )
}

/// Where a row of a menu sits in it, counted by what it says: the rows a hand counts are the
/// rows themselves, and the number one of them sits at moves the moment another is put above
/// it (see `pin_menu::PIN_MENU_ROWS`).
fn row_at(rows: &[pin_chrome::MenuRow], label: &str) -> usize {
    rows.iter()
        .position(|row| row.label == label)
        .unwrap_or_else(|| panic!("the menu holds the {label} row"))
}

/// The middle of a row of a flyout: what
/// a press or a hover on that row is aimed with. It is taken from the band the panel laid
/// the row out in rather than from arithmetic on the row height (see
/// `pin_chrome::MenuRowBand`), and it is a point on the screen with it.
fn flyout_row_point(flyout: &PinMenuFlyout, index: usize) -> (i32, i32) {
    let band = flyout
        .popup
        .bands
        .iter()
        .find(|band| band.answers == index)
        .unwrap_or_else(|| panic!("the flyout lays the row a hand counts {index} out in a band"));

    (
        (flyout.popup.panel.left + flyout.popup.panel.right) / 2,
        band.top + band.height / 2,
    )
}

/// The middle of a row of the bubble's own menu, taken from the band the panel laid the row out
/// in — the same answer `menu_row_point` gives for a pinned window's panel (see
/// `pin_chrome::MenuRowBand`).
fn bubble_row_point(index: usize) -> (i32, i32) {
    let menu = BUBBLE_MENU.lock().expect("the bubble's menu");
    let menu = menu.as_ref().expect("the bubble's menu is up");
    let band = menu
        .popup
        .bands
        .iter()
        .find(|band| band.answers == index)
        .unwrap_or_else(|| panic!("the panel lays the row a hand counts {index} out in a band"));

    (
        (menu.popup.panel.left + menu.popup.panel.right) / 2,
        band.top + band.height / 2,
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

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.point = on_screen((20, 20));
    });
    let paint = pinned_menu_geometry().expect("the menu is up");
    let seek = menu_row_point(&paint, row_at(&paint.rows, "Seek"));

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

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.point = on_screen((20, 20));
    });
    let paint = pinned_menu_geometry().expect("the menu is up");
    let seek = menu_row_point(&paint, row_at(&paint.rows, "Seek"));
    let end_of_file = menu_row_point(&paint, row_at(&paint.rows, "End of File"));

    let t0 = Instant::now();
    tick(Some(seek), t0);
    // The pointer leaves the row before the delay: the timer is cleared.
    tick(Some(end_of_file), t0 + Duration::from_millis(50));
    assert!(!flyout_is_up(), "it did not open before the delay");

    // And it never opens on this rest, however long it lasts.
    tick(Some(end_of_file), t0 + Duration::from_millis(400));
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

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.point = on_screen((20, 20));
    });

    // Open the flyout by resting on the Seek row past the delay.
    let paint = pinned_menu_geometry().expect("the menu is up");
    let seek = menu_row_point(&paint, row_at(&paint.rows, "Seek"));
    let t0 = Instant::now();
    tick(Some(seek), t0);
    tick(Some(seek), t0 + Duration::from_millis(200));
    assert!(flyout_is_up(), "the flyout opened");

    // Over the flyout itself: it stays.
    let paint = pinned_menu_geometry().expect("the menu is up");
    let flyout = paint.flyout.as_ref().expect("the flyout is up");
    let on_flyout = (
        (flyout.popup.panel.left + flyout.popup.panel.right) / 2,
        flyout.popup.rows_top + flyout.popup.row_height / 2,
    );
    tick(Some(on_flyout), t0 + Duration::from_millis(300));
    assert!(flyout_is_up(), "the flyout stays up under the pointer");

    // In the gap between the two panels, at the Seek row's own height:
    // it stays up across it.
    let paint = pinned_menu_geometry().expect("the menu is up");
    let flyout = paint.flyout.as_ref().expect("the flyout is up");
    let gap = (
        (paint.popup.panel.right + flyout.popup.panel.left) / 2,
        menu_row_point(&paint, row_at(&paint.rows, "Seek")).1,
    );
    tick(Some(gap), t0 + Duration::from_millis(400));
    assert!(
        flyout_is_up(),
        "the flyout stays up while the pointer crosses the gap"
    );

    // On the End of File row: it hides.
    let end_of_file = menu_row_point(&paint, row_at(&paint.rows, "End of File"));
    assert!(
        tick(Some(end_of_file), t0 + Duration::from_millis(500)),
        "the flyout going is a change worth a repaint"
    );
    assert!(
        !flyout_is_up(),
        "the flyout hides when the pointer hovers another row"
    );

    // Opened again, then the pointer leaves the menu entirely: it hides.
    let seek = menu_row_point(&paint, row_at(&paint.rows, "Seek"));
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

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.point = on_screen((20, 20));
    });
    let paint = pinned_menu_geometry().expect("the menu is up");
    let seek_row = menu_row_point(&paint, row_at(&paint.rows, "Seek"));
    let t0 = Instant::now();

    assert!(
        tick(Some(seek_row), t0),
        "the row under the pointer is a change worth a repaint"
    );
    let paint = pinned_menu_geometry().expect("the menu is up");
    assert_eq!(
        paint.hover,
        Some(row_at(&paint.rows, "Seek")),
        "the Seek row is the one the pointer is on"
    );
    assert!(
        paint.flyout.is_none(),
        "the flyout is down on the main page"
    );

    // And on the flyout: its own row is the one washed.
    with_pin(|pin| pin.menu.seek = true);
    let paint = pinned_menu_geometry().expect("the menu is up");
    let flyout = paint.flyout.as_ref().expect("the flyout is up");
    let choice = (
        (flyout.popup.panel.left + flyout.popup.panel.right) / 2,
        flyout.popup.rows_top + 2 * flyout.popup.row_height + flyout.popup.row_height / 2,
    );
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

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

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

/// The same ask puts the bubble's own menu away: with the pin's window hidden the bubble is
/// the pin's own window, so a press the hook read on the browser beneath it — or on the
/// desktop — landed outside the bubble as surely as it landed outside the card's, and the
/// bubble's window never saw it to answer it (see `bubble_window_is_at` and
/// `close_pin_bubble_menu`).
#[test]
fn the_same_ask_puts_the_bubbles_menu_away() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    // The bubble's own menu up in the circle a collapsed pin left, with the bubble's handle
    // in the slot the loop's tick reads it from — a window nobody has to put on a screen
    // for a menu to be stood up in a slot (see `open_bubble_menu`).
    let _bubble = BubblePutBack::put(HWND(0x2000 as *mut _));
    let rows = bubble_menu_chrome_rows(Some(MediaType::Audio), true);
    let popup = bubble_menu_popup((0, 44), 400, 400, 96, &rows);
    if let Ok(mut menu) = BUBBLE_MENU.lock() {
        *menu = Some(BubbleMenu {
            rows,
            popup,
            hover: None,
            dpi: 96,
        });
    }
    assert!(bubble_menu_is_open(), "the bubble's menu is up");

    // The hook's ask for a press that landed on another window, taken on the loop's tick.
    request_pin_menu_dismiss();
    assert!(
        take_pin_menu_dismiss(),
        "the ask is taken and a menu is the thing it closes"
    );
    assert!(!bubble_menu_is_open(), "the bubble's menu is away with it");
    assert!(!pin_menu_is_open(), "and the card's own was never up");
    assert!(pinned(), "while the pin that was standing is left standing");
}

/// A press on the bubble's `Pause on Minimize` row turns the kind's own switch over and leaves
/// the panel up with the check it now carries: the one row of a bubble's menu that answers where
/// it stands, the way the card's own `Shuffle` row answers (see `bubble_menu_press`).
#[test]
fn a_press_on_the_bubbles_pause_switch_turns_it_over() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    // Both switches on, so the press below has one of them to turn off, and a pin whose card
    // the standing menu is read against.
    let _switches = PausePreviewSettings::stood_at(true, true);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    // A sound behind the bubble: the switch the menu's row carries is the sound's own half of
    // the pair, and a kind with no playback has no row to press at all (see `pin_pause_audio`).
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        let mut media = create_loading_media(8, 4);
        media.media_type = MediaType::Audio;
        *slot = Some(media);
    }

    // The bubble's own menu up in the circle a collapsed pin left, stood in the slot the loop's
    // tick reads it from, the way the dismiss test above stands it (see `open_bubble_menu`).
    let hwnd = HWND(0x2000 as *mut _);
    let _bubble = BubblePutBack::put(hwnd);
    let rows = bubble_menu_chrome_rows(Some(MediaType::Audio), true);
    let popup = bubble_menu_popup((0, 44), 400, 400, 96, &rows);
    if let Ok(mut menu) = BUBBLE_MENU.lock() {
        *menu = Some(BubbleMenu {
            rows,
            popup,
            hover: None,
            dpi: 96,
        });
    }

    // The middle of the switch's own band, named by what the row says rather than by the number
    // it happens to sit at.
    let index = row_at(
        &bubble_menu_chrome_rows(Some(MediaType::Audio), true),
        "Pause on Minimize",
    );
    let (x, y) = bubble_row_point(index);
    assert!(
        unsafe { bubble_menu_press(hwnd, x, y) },
        "a hand on the switch is the menu's own"
    );

    assert!(
        !CONFIG.lock().expect("the configuration").pin_pause_audio,
        "the sound's switch was on, so the press turned it off"
    );
    let check = {
        let menu = BUBBLE_MENU.lock().expect("the bubble's menu");
        let menu = menu.as_ref().expect("the panel stays up over the switch");
        matches!(menu.rows[index].mark, pin_chrome::MenuMark::Check(false))
    };
    assert!(
        check,
        "and the check the panel carries is the switch it now stands for"
    );

    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A press on one of the bubble's two minimize picks writes the setting the tray's own
/// `Pin Mode → Minimize` writes and puts the panel away with it: the two rows are one group
/// rather than two switches, so a press on one is a choice made and not a mark turned over (see
/// `shared_menu_row_press`).
#[test]
fn a_press_on_the_bubbles_minimize_pick_writes_the_setting() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _config_file = ConfigFilePutBack::take();
    let _put_back = PutBack(take_pin_for_a_test());
    let _target = PinMinimizeToPutBack::put(PinMinimizeTo::Bubble);
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    // A sound behind the bubble, which is the kind the bubble's rows are read from.
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        let mut media = create_loading_media(8, 4);
        media.media_type = MediaType::Audio;
        *slot = Some(media);
    }

    // The bubble's own menu up in the circle a collapsed pin left, the way the switch's own test
    // stands it (see `open_bubble_menu`).
    let hwnd = HWND(0x2000 as *mut _);
    let _bubble = BubblePutBack::put(hwnd);
    let rows = bubble_menu_chrome_rows(Some(MediaType::Audio), true);
    let index = row_at(&rows, "Minimize to Taskbar");
    let popup = bubble_menu_popup((0, 44), 400, 400, 96, &rows);
    if let Ok(mut menu) = BUBBLE_MENU.lock() {
        *menu = Some(BubbleMenu {
            rows,
            popup,
            hover: None,
            dpi: 96,
        });
    }

    let (x, y) = bubble_row_point(index);
    assert!(
        unsafe { bubble_menu_press(hwnd, x, y) },
        "a hand on a minimize pick is the menu's own"
    );
    assert_eq!(
        CONFIG.lock().expect("the configuration").pin_minimize_to,
        PinMinimizeTo::Taskbar,
        "the pick wrote the place the tray's own row writes"
    );
    assert!(
        !bubble_menu_is_open(),
        "and the panel goes away with the pick, the way a menu a choice was made in does"
    );

    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// Closing the menu from the window side puts both panels away and leaves
/// the pin standing: the answer a captured outside press gets, rather than
/// the hook's ask a tick later.
#[test]
fn closing_the_menu_from_the_window_side_puts_both_panels_away() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    let hwnd = HWND(0x1000 as *mut _);
    unsafe { open_pin_menu(hwnd, 20, 20, PinMenuPage::Chrome) };
    assert!(pin_menu_is_open(), "the menu is up");
    assert!(pinned_menu_geometry().is_some(), "both panels are painted");

    assert!(
        unsafe { close_pin_menu(hwnd) },
        "closing an open menu says it came down"
    );
    assert!(!pin_menu_is_open(), "and the menu is away");
    assert!(pinned_menu_geometry().is_none(), "both panels with it");
    assert!(pinned(), "while the pin that was standing is left standing");

    assert!(
        !unsafe { close_pin_menu(hwnd) },
        "closing a menu that is down says there was nothing to put away"
    );
}

/// A point on neither of the card's menu's panels is outside it, and a point on
/// a row is not: what a captured press aimed at another app or the taskbar is
/// told apart by. The panels are drawn in a window of their own over the screen,
/// so the point is a place on the screen and what says it is on the menu is a
/// panel's own rows (see `pin_menu_point_is_outside`).
#[test]
fn a_point_off_both_panels_is_outside_and_one_on_a_row_is_not() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let _settings = MenuSettings::put(AudioSeek::Start, PinEndOfFile::Loop, false);
    let _put_back = PutBack(take_pin_for_a_test());
    stand_pin(Some(a_sound_card_whose_menu_fits_the_screen()));

    unsafe { open_pin_menu(HWND(0x1000 as *mut _), 120, 120, PinMenuPage::Chrome) };
    let paint = pinned_menu_geometry().expect("the menu is up");

    // The middle of the Shuffle row, which is a row the menu holds.
    let (x, y) = menu_row_point(&paint, row_at(&paint.rows, "Shuffle"));
    assert!(
        !pin_menu_point_is_outside(x, y),
        "a press on a row of the menu is on the menu"
    );

    // A point beside the panel, in the row of its first row: not the panel, and not
    // a row of any of the menu's.
    let x = paint.popup.panel.left - 10;
    let y = paint.popup.rows_top + paint.popup.row_height / 2;
    assert!(
        pin_menu_point_is_outside(x, y),
        "a press off the panel is outside"
    );

    // And a press that landed on another window or on the desktop.
    assert!(
        pin_menu_point_is_outside(-1_000, -1_000),
        "a press from another window is outside"
    );
}

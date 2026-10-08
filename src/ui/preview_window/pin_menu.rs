//! The menu a pinned sound's card opens from a right-click on its
//! window: the state of it, the rows it holds, and what a press
//! on one of them does.
//!
//! The menu is the card's own rather than the tray's, so its rows are
//! read out of the same settings the tray's own `Pin Mode` submenus
//! write, and its panel is drawn by `pin_chrome` rather than by the
//! menu the shell draws: a menu that floats over the media is a panel
//! a hand reads and then does something about, and it is drawn the way
//! this window draws its own panels (see `pin_chrome::menu`).

use super::*;

/// The state of the card's menu: whether its panel is up, and
/// whether the Seek row's flyout is up beside it.
///
/// Both are the window's own rather than the configuration's, because
/// they are what the pointer is looking at right now: a panel up is up
/// until a press puts it away, and a flyout up is up until the press
/// that put it up puts it back or a choice in it is made (see
/// `pinned_menu_press`).
#[derive(Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct PinMenu {
    /// Whether the panel is up.
    pub(super) open: bool,
    /// Whether the Seek row's flyout is up beside the menu,
    /// rather than the menu holding its three rows alone.
    pub(super) seek: bool,
    /// Where the right-click that opened the panel landed, in the
    /// window's own coordinates: the panel's top-left, held inside
    /// the window by the placement (see `pin_chrome::menu_popup_from_point`).
    pub(super) point: (i32, i32),
}

/// What a repaint draws the card's menu from: the
/// panel it floats in, the rows in it, and the
/// Seek row's flyout beside it when the flyout is
/// up.
///
/// A value rather than a reference, for the reason every other thing a
/// repaint draws from is: the pin's lock is let go of before anything
/// is drawn (see `PinnedPaint`).
pub(super) struct PinMenuPaint {
    pub(super) popup: pin_chrome::MenuPopup,
    pub(super) rows: Vec<pin_chrome::MenuRow>,
    /// The Seek row's flyout, when it is up: its panel
    /// right of the main menu's, top-aligned with the
    /// Seek row, and the rows in it. The main menu's
    /// panel is `popup` as the flyout's own placement
    /// pulled it left to make the flyout room (see
    /// `pin_chrome::menu_flyout_from_menu`).
    pub(super) flyout: Option<PinMenuFlyout>,
}

/// What a repaint draws the Seek row's flyout from:
/// the panel it floats in beside the main menu's,
/// and the rows in it.
///
/// A value rather than a reference, for the reason
/// `PinMenuPaint` is.
pub(super) struct PinMenuFlyout {
    pub(super) popup: pin_chrome::MenuPopup,
    pub(super) rows: Vec<pin_chrome::MenuRow>,
}

/// The ways a pinned sound can start, in the order the menu lists them:
/// the same four, in the same order, the tray's own
/// `Volume → Pin Mode Audio Seek` submenu offers (see
/// `shell::tray::ids::AUDIO_SEEK_CHOICES`), held here because that
/// list is the tray's own and this is the card's.
const SEEK_CHOICES: [AudioSeek; 4] = [
    AudioSeek::Remember,
    AudioSeek::Start,
    AudioSeek::Middle,
    AudioSeek::Random,
];

/// What each of the seek choices is listed as: the words the tray lists
/// the same four under (see `shell::tray::submenus::
/// pin_mode_audio_seek_label`), without the mark that menu puts on the
/// default — the card's menu marks the one in force instead (see
/// `menu_rows`).
const SEEK_LABELS: [&str; 4] = [
    "Remember",
    "From the Start",
    "From the Middle",
    "Random",
];

/// The rows the card's menu holds: the two mode
/// toggles and the row that opens the seek
/// choices.
///
/// The main menu is always these three rows — the
/// seek choices are the flyout's own rows, beside
/// the menu's rather than in place of them (see
/// `seek_rows`).
///
/// The rows are read out of the configuration rather than held in the
/// pin's state, because what a row *says* is what the setting says: a
/// marker is the setting's own answer to "which one is this?", and a
/// toggle is the setting's own opposite. Reading them at the paint
/// rather than remembering them is also what makes a change to a
/// setting show the moment it is made, from wherever it was made.
pub(super) fn menu_rows() -> Vec<pin_chrome::MenuRow> {
    let (shuffle, loop_) = pin_mode_audio_toggles();
    vec![
        pin_chrome::MenuRow {
            label: "Shuffle Mode".to_string(),
            mark: pin_chrome::MenuMark::Check(shuffle),
        },
        pin_chrome::MenuRow {
            label: "Loop".to_string(),
            mark: pin_chrome::MenuMark::Check(loop_),
        },
        pin_chrome::MenuRow {
            label: "Seek".to_string(),
            mark: pin_chrome::MenuMark::None,
        },
    ]
}

/// The rows the Seek row's flyout holds: the four
/// seek choices, in the order the tray lists them
/// in, with the one the pin plays by carrying its
/// disc.
///
/// Read out of the configuration for the same
/// reason the main menu's rows are (see
/// `menu_rows`).
pub(super) fn seek_rows() -> Vec<pin_chrome::MenuRow> {
    let current = pin_mode_audio_seek();
    SEEK_CHOICES
        .iter()
        .zip(SEEK_LABELS)
        .map(|(choice, label)| pin_chrome::MenuRow {
            label: label.to_string(),
            mark: if *choice == current {
                pin_chrome::MenuMark::Bullet
            } else {
                pin_chrome::MenuMark::None
            },
        })
        .collect()
}

/// Where a pinned sound starts, which is the pin's own setting's answer.
fn pin_mode_audio_seek() -> AudioSeek {
    CONFIG
        .lock()
        .ok()
        .map(|config| config.pin_mode_audio_seek)
        .unwrap_or(DEFAULT_PIN_MODE_AUDIO_SEEK)
}

/// The two mode toggles, read the same way. A lock that cannot be taken
/// is answered with the defaults rather than with nothing, because a row
/// that says nothing is a row a hand cannot read.
fn pin_mode_audio_toggles() -> (bool, bool) {
    let Ok(config) = CONFIG.lock() else {
        return (
            DEFAULT_PIN_MODE_AUDIO_SHUFFLE,
            DEFAULT_PIN_MODE_AUDIO_LOOP,
        );
    };

    (config.pin_mode_audio_shuffle, config.pin_mode_audio_loop)
}

/// The two mode toggles and where a pinned sound starts, written the way
/// every setting this app's own menus write is: under the
/// configuration's lock and saved (see
/// `shell::tray::commands::setters`).
fn toggle_pin_mode_audio_shuffle() {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_mode_audio_shuffle = !config.pin_mode_audio_shuffle;
        config.save();
    }
}

fn toggle_pin_mode_audio_loop() {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_mode_audio_loop = !config.pin_mode_audio_loop;
        config.save();
    }
}

fn set_pin_mode_audio_seek(seek: AudioSeek) {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_mode_audio_seek = seek;
        config.save();
    }
}

/// Where the card's menu floats, or nothing while it is closed: the
/// panel anchored at the point the right-click that opened it landed,
/// holding the rows of the page it is showing.
///
/// It is asked of the pin rather than recomputed at the paint, so that
/// the panel drawn is the panel a press is answered against (see
/// `pinned_volume_geometry`).
pub(super) fn pinned_menu_geometry() -> Option<PinMenuPaint> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;

    pin.menu.open.then(|| pinned_menu_paint(pin)).flatten()
}

/// The menu as it is painted: the rows the main
/// page holds, the panel anchored at the point the
/// right-click that opened it landed, and the
/// Seek row's flyout beside it when the flyout is
/// up.
///
/// The panel is anchored at the pin's own remembered click point, in
/// the window's own coordinates, and held inside the window by the
/// placement (see `pin_chrome::menu_popup_from_point`), so a panel
/// opened on a hand near an edge is a panel a hand can still reach.
fn pinned_menu_paint(pin: &PinnedPreview) -> Option<PinMenuPaint> {
    if !pin_shows_an_audio_card(pin) {
        return None;
    }

    let rows = menu_rows();
    let choices = pin.menu.seek.then(seek_rows);
    let (width, height) = pin.window_size();

    // The flyout's own placement is what says where the main
    // menu itself is, because the two panels are placed together
    // to keep both of them inside the window (see
    // `pin_chrome::menu_flyout_from_menu`).
    let (popup, flyout) = match choices {
        Some(choices) => {
            let placed = pin_chrome::menu_flyout_from_menu(
                &pin_chrome::menu_popup_from_point(pin.menu.point, width, height, pin.dpi, &rows),
                width,
                height,
                pin.dpi,
                &choices,
            );
            (
                placed.menu,
                Some(PinMenuFlyout {
                    popup: placed.popup,
                    rows: choices,
                }),
            )
        }
        None => (
            pin_chrome::menu_popup_from_point(pin.menu.point, width, height, pin.dpi, &rows),
            None,
        ),
    };

    Some(PinMenuPaint { popup, rows, flyout })
}

/// A press on a pinned window with the card's menu up, answering whether
/// it was the menu's: a press on a row is that row's own — a mode
/// toggles and the panel goes away with the answer, the Seek row opens
/// its flyout beside the menu and the panel stays up to hold it, a row
/// of the flyout is the choice the pin plays by from then on and the
/// whole menu goes away with it — while a press anywhere else is a press
/// that has left the menu, which puts both panels away.
///
/// Nothing is captured and nothing is held, because a menu is a thing a
/// hand reads rather than a thing it carries: the press is answered
/// where it landed, and the panel is one paint away from being gone (see
/// `open_pin_menu`).
pub(super) unsafe fn pinned_menu_press(hwnd: HWND, x: i32, y: i32) -> bool {
    let Some(paint) = pinned_menu_geometry() else {
        return false;
    };

    // The row the press landed on, if it landed on one: each panel's own
    // rows are what a press is answered against, and the pad around them
    // and everywhere outside both panels is not a row (see
    // `pin_chrome::menu_row_at`).
    let stays_up = match pin_chrome::menu_row_at(&paint.popup, x, y) {
        Some(index) => match index {
            // The two mode toggles: the setting is turned over and the
            // panel goes away, because a toggle is an answer rather than
            // a door.
            0 => {
                toggle_pin_mode_audio_shuffle();
                false
            }
            1 => {
                toggle_pin_mode_audio_loop();
                false
            }
            // The seek row: the door the flyout comes through. While the
            // flyout is down, the press is the one that opens it beside
            // the menu; while it is up, the press is the one that puts
            // it back — the menu itself stays up, its rows still
            // showing.
            _ => {
                with_pin(|pin| pin.menu.seek = !pin.menu.seek);
                true
            }
        },
        None => match paint
            .flyout
            .as_ref()
            .and_then(|flyout| pin_chrome::menu_row_at(&flyout.popup, x, y))
        {
            // A row of the flyout: the one pressed is the one the pin
            // plays by from then on, and the whole menu goes away with
            // the answer — both panels, because the flyout is the
            // menu's own — the way a menu a choice was made in does.
            Some(index) => {
                if let Some(choice) = SEEK_CHOICES.get(index) {
                    set_pin_mode_audio_seek(*choice);
                }
                false
            }
            // A press anywhere else is a press outside both panels,
            // which puts the whole menu away: the panels are over the
            // card, and a hand that has come for what is under them is
            // a hand that has left the menu.
            None => false,
        },
    };

    if !stays_up {
        with_pin(|pin| pin.menu.open = false);
    }
    render_layered_preview(hwnd);

    true
}

/// The card's menu put up by a right-click on the pin's window, at the
/// point the right-click landed: a second right-click while the panel is
/// up moves it to the new point rather than putting it away.
///
/// It is the pin's own menu and only a sound's card carries one, so a
/// right-click on a pin showing anything else opens nothing at all (see
/// `pin_shows_an_audio_card`).
///
/// A menu put up is put up on its main page, the
/// flyout put away: a hand that opens the menu has
/// come for the toggles, and the choices are one
/// row away.
pub(super) unsafe fn open_pin_menu(hwnd: HWND, x: i32, y: i32) {
    let audio = pin_state()
        .and_then(|pinned| pinned.pin().map(pin_shows_an_audio_card))
        .unwrap_or(false);
    if !audio {
        return;
    }

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.seek = false;
        pin.menu.point = (x, y);
    });
    render_layered_preview(hwnd);
}

//! The menu a pinned sound's card opens from the gear in its
//! top corner: the state of it, the rows it holds, and what a press
//! on one of them does.
//!
//! The menu is the card's own rather than the tray's, so its rows are
//! read out of the same settings the tray's own `Pin Mode` submenus
//! write, and its panel is drawn by `pin_chrome` rather than by the
//! menu the shell draws: a menu that floats over the media is a panel
//! a hand reads and then does something about, and it is drawn the way
//! this window draws its own panels (see `pin_chrome::menu`).

use super::*;

/// The state of the card's menu: whether its panel is up, and which
/// of its two pages it is showing.
///
/// Both are the window's own rather than the configuration's, because
/// they are what the pointer is looking at right now: a panel up is up
/// until a press puts it away, and the page it shows is the page the
/// last row pressed opened (see `pinned_menu_press`).
#[derive(Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct PinMenu {
    /// Whether the panel is up.
    pub(super) open: bool,
    /// Whether the panel is showing the seek choices rather than the
    /// two mode toggles and the row that opens them.
    pub(super) seek: bool,
}

/// What a repaint draws the card's menu from: the panel it floats in,
/// and the rows in it.
///
/// A value rather than a reference, for the reason every other thing a
/// repaint draws from is: the pin's lock is let go of before anything
/// is drawn (see `PinnedPaint`).
pub(super) struct PinMenuPaint {
    pub(super) popup: pin_chrome::MenuPopup,
    pub(super) rows: Vec<pin_chrome::MenuRow>,
    /// Which of the menu's two pages the rows are, which is
    /// what a press on one of them is answered against (see
    /// `pinned_menu_press`).
    pub(super) seek: bool,
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

/// The rows the card's menu holds, for the page it is showing: the two
/// mode toggles and the row that opens the seek choices, or the seek
/// choices themselves.
///
/// The rows are read out of the configuration rather than held in the
/// pin's state, because what a row *says* is what the setting says: a
/// marker is the setting's own answer to "which one is this?", and a
/// toggle is the setting's own opposite. Reading them at the paint
/// rather than remembering them is also what makes a change to a
/// setting show the moment it is made, from wherever it was made.
pub(super) fn menu_rows(seek_page: bool) -> Vec<pin_chrome::MenuRow> {
    if !seek_page {
        let (shuffle, loop_) = pin_mode_audio_toggles();
        return vec![
            pin_chrome::MenuRow {
                label: "Shuffle Mode".to_string(),
                marked: shuffle,
            },
            pin_chrome::MenuRow {
                label: "Loop".to_string(),
                marked: loop_,
            },
            pin_chrome::MenuRow {
                label: "Seek".to_string(),
                marked: false,
            },
        ];
    }

    let current = pin_mode_audio_seek();
    SEEK_CHOICES
        .iter()
        .zip(SEEK_LABELS)
        .map(|(choice, label)| pin_chrome::MenuRow {
            label: label.to_string(),
            marked: *choice == current,
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
/// panel hung from the gear the card's menu opens from, holding the
/// rows of the page it is showing.
///
/// It is asked of the pin rather than recomputed at the paint, so that
/// the panel drawn is the panel a press is answered against (see
/// `pinned_volume_geometry`).
pub(super) fn pinned_menu_geometry() -> Option<PinMenuPaint> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;

    pin.menu.open.then(|| pinned_menu_paint(pin)).flatten()
}

/// The menu as it is painted: the rows the page it is showing holds,
/// and the panel hung from the gear's own box for those rows to be
/// drawn in.
///
/// The panel is hung from the gear's own box, from the card's own
/// arithmetic rather than from anything kept beside it, so a panel
/// opened after a change to Audio Scaling hangs off the gear as the
/// card drew it (see `pinned_audio_volume_popup`, which is the same
/// road the volume button's panel takes).
fn pinned_menu_paint(pin: &PinnedPreview) -> Option<PinMenuPaint> {
    if !pin_shows_an_audio_card(pin) {
        return None;
    }

    let rows = menu_rows(pin.menu.seek);
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

    Some(PinMenuPaint {
        popup: pin_chrome::menu_popup_from_bullet(
            RECT {
                top: gear.top + top,
                ..gear
            },
            width,
            height,
            pin.dpi,
            &rows,
        ),
        rows,
        seek: pin.menu.seek,
    })
}

/// A press on a pinned window with the card's menu up, answering whether
/// it was the menu's: a press on a row is that row's own — a mode
/// toggles and the panel goes away with the answer, the seek row opens
/// the seek choices and the panel stays up to hold them, and a seek
/// choice is the one the pin plays by from then on — while a press
/// anywhere else is a press that has left the menu, which puts the panel
/// away.
///
/// The gear that opens the menu is left out on purpose: a
/// press on it is the press every card control follows, held
/// and acted on at the release, and the release is what
/// toggles the panel (see `pinned_audio_control_release`) —
/// putting the panel away here would have the release open it
/// straight back up, which is the same bargain the volume
/// button's press makes (see the top of `pinned_press`).
///
/// Nothing is captured and nothing is held, because a menu is a thing a
/// hand reads rather than a thing it carries: the press is answered
/// where it landed, and the panel is one paint away from being gone (see
/// `toggle_pin_menu`).
pub(super) unsafe fn pinned_menu_press(hwnd: HWND, x: i32, y: i32) -> bool {
    let Some(paint) = pinned_menu_geometry() else {
        return false;
    };

    // A press on the gear the menu came out of is the press that opens
    // and closes it, left to the road every card control follows (see
    // `pinned_audio_control_press`).
    if pressed_on_the_gear(x, y) {
        return false;
    }

    // The row the press landed on, if it landed on one: the panel's own
    // rows are what a press is answered against, and the pad around them
    // and everything outside the panel is not a row (see
    // `pin_chrome::menu_row_at`).
    let stays_up = match pin_chrome::menu_row_at(&paint.popup, x, y) {
        Some(index) if paint.seek => {
            // The seek choices: the one pressed is the one the pin plays
            // by from then on, and the panel goes away with the answer,
            // the way a menu a choice was made in does.
            if let Some(choice) = SEEK_CHOICES.get(index) {
                set_pin_mode_audio_seek(*choice);
            }
            false
        }
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
            // The seek row: the panel stays up, holding the choices it
            // opens in place of the row that opened them.
            _ => {
                with_pin(|pin| pin.menu.seek = true);
                true
            }
        },
        // A press anywhere else is a press outside the menu, which puts
        // the panel away: the panel is over the card, and a hand that
        // has come for what is under it is a hand that has left the
        // menu.
        None => false,
    };

    if !stays_up {
        with_pin(|pin| pin.menu.open = false);
    }
    render_layered_preview(hwnd);

    true
}

/// Whether a press is on the gear the card's menu opens from: asked
/// of the card's own layout rather than of the panel, because the
/// gear is a fact of the card whether the menu is up or not (see
/// `pin_audio_control_at`).
fn pressed_on_the_gear(x: i32, y: i32) -> bool {
    pin_state()
        .and_then(|pinned| {
            let pin = pinned.pin()?;
            Some(pin_audio_control_at(pin, x, y) == Some(CardControl::Menu))
        })
        .unwrap_or(false)
}

/// The card's menu put up or put away by the press on the gear that
/// opens it: the gear's press and release, which is the road every card
/// control follows (see `pinned_audio_control_release`).
///
/// A menu put up is put up as the page it opens on, whatever page it
/// was showing when it was last put away: a hand that opens the menu has
/// come for the toggles, and the choices are one row away.
pub(super) unsafe fn toggle_pin_menu(hwnd: HWND) {
    with_pin(|pin| {
        pin.menu.open = !pin.menu.open;
        pin.menu.seek = false;
    });
    render_layered_preview(hwnd);
}

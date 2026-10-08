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
    /// The row of the main menu the pointer is over, for the wash the
    /// row under the hand is painted with, or nothing where the
    /// pointer is not on a row (see `refresh_pin_menu`).
    pub(super) hover_main: Option<usize>,
    /// The row of the flyout the pointer is over, for the same wash.
    pub(super) hover_flyout: Option<usize>,
    /// When the pointer came onto the Seek row, while the hover timer
    /// that opens the flyout is running: cleared the moment the
    /// pointer leaves the row, so that a hand only passing over it
    /// opens nothing (see `refresh_pin_menu`).
    pub(super) seek_since: Option<Instant>,
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
    /// The row of the main menu the pointer is over, for the
    /// wash the row under the hand is painted with.
    pub(super) hover: Option<usize>,
    /// The Seek row's flyout, when it is up: its panel
    /// beside the main menu's, top-aligned with the
    /// Seek row, and the rows in it. The main menu's
    /// panel is `popup` as the flyout's own placement
    /// leaves it — on one side of the point the right-click
    /// landed or the other (see `pin_chrome::menu_flyout_from_menu`).
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
    /// The row of the flyout the pointer is over, for its wash.
    pub(super) hover: Option<usize>,
}

/// How long the pointer rests on the Seek row before its flyout opens:
/// the Windows-like hover delay, long enough that a hand on its way
/// across the row does not open a panel it did not mean to, and short
/// enough that a hand meaning to use the row does not wait (see
/// `refresh_pin_menu`).
const PIN_MENU_HOVER_DELAY: Duration = Duration::from_millis(150);

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
            label: "Shuffle".to_string(),
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
    // `pin_chrome::menu_flyout_from_menu`). With the flyout down,
    // the main menu is placed alone — and carries the same arrow
    // the flyout's placement answers, both placements asking the
    // one tier the flyout opens in, so the Seek row points the way
    // its flyout opens whether the flyout is up or not.
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
                    hover: pin.menu.hover_flyout,
                }),
            )
        }
        None => (
            pin_chrome::menu_popup_from_point(pin.menu.point, width, height, pin.dpi, &rows),
            None,
        ),
    };

    Some(PinMenuPaint {
        popup,
        rows,
        hover: pin.menu.hover_main,
        flyout,
    })
}

/// Ask the card's menu about the pointer, on the tick that already drives
/// the pin's chrome, answering whether anything about it changed — which is
/// whether the window owes a repaint for it.
///
/// Three things are settled here and nowhere else, because they are all
/// questions about where the pointer is right now rather than about anything
/// the window was sent: the row the pointer is on, which is washed; whether
/// the pointer has rested on the Seek row long enough for its flyout to open;
/// and whether a flyout that is up is still wanted. A pointer that merely
/// crosses the Seek row opens nothing — the timer is cleared the moment it
/// leaves — and a flyout that is up stays up while the pointer is on the Seek
/// row, in the gap between the panels, or on the flyout itself, and goes when
/// the pointer is on another row or off the menu entirely (see
/// `pin_chrome::menu_flyout_gap_holds`).
///
/// The pointer is read as a place on the screen and taken against the
/// window's own box, the conversion `refresh_pin_volume` makes, for the same
/// reason it makes it: a pointer that has left the pin sends it nothing more,
/// so a wash or a flyout left up under one is a mark on a hand that is gone.
pub(super) fn refresh_pin_menu(
    pin: &mut PinnedPreview,
    now: Instant,
    cursor: Option<(i32, i32)>,
) -> bool {
    if !pin.menu.open {
        return false;
    }

    let Some(paint) = pinned_menu_paint(pin) else {
        return false;
    };

    let window = pin.window_box();
    let at = cursor.map(|(x, y)| (x - window.0, y - window.1));

    let hover_main = at.and_then(|(x, y)| pin_chrome::menu_row_at(&paint.popup, x, y));
    let hover_flyout = paint
        .flyout
        .as_ref()
        .and_then(|flyout| at.and_then(|(x, y)| pin_chrome::menu_row_at(&flyout.popup, x, y)));

    // The Seek row is the main menu's last: it is the row the flyout hangs
    // from, whatever the panel holds (see `menu_rows`).
    let seek_row = (paint.rows.len() as i32 - 1).max(0) as usize;
    let on_seek = hover_main == Some(seek_row);

    // The timer: it starts when the pointer comes onto the Seek row and is
    // cleared the moment it leaves, so a crossing opens nothing.
    let seek_since = match on_seek {
        true => pin.menu.seek_since.or(Some(now)),
        false => None,
    };
    let rested = seek_since.is_some_and(|since| now.duration_since(since) >= PIN_MENU_HOVER_DELAY);

    // Whether a flyout that is up is still wanted: the pointer is on the Seek
    // row, the gap, or the flyout itself.
    let kept = on_seek
        || hover_flyout.is_some()
        || paint.flyout.as_ref().is_some_and(|flyout| {
            at.is_some_and(|(x, y)| {
                pin_chrome::menu_flyout_gap_holds(&paint.popup, &flyout.popup, x, y)
            })
        });

    let seek = (pin.menu.seek || rested) && kept;

    let changed = pin.menu.hover_main != hover_main
        || pin.menu.hover_flyout != hover_flyout
        || pin.menu.seek != seek
        || pin.menu.seek_since != seek_since;
    pin.menu.hover_main = hover_main;
    pin.menu.hover_flyout = hover_flyout;
    pin.menu.seek = seek;
    pin.menu.seek_since = seek_since;

    changed
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
            // The seek row: the door the flyout comes through. A press on
            // it opens the flyout beside the menu and the menu stays up,
            // its rows still showing — the same door the hover opens, for
            // a hand that has come to press rather than wait.
            _ => {
                with_pin(|pin| pin.menu.seek = true);
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
        pin.menu.hover_main = None;
        pin.menu.hover_flyout = None;
        pin.menu.seek_since = None;
    });
    render_layered_preview(hwnd);
}

/// Take the ask the Explorer hook left for a press that landed on another window and put the
/// card's menu away where it is up, answering whether it came down — a repaint is owed where
/// it did. It is the preview loop's own read of the ask, taken on its tick (see
/// `PIN_MENU_DISMISS_REQUESTED`), and only the menu goes: the pin itself is left standing,
/// which is what a press on another window means here (see
/// `requests::request_pin_menu_dismiss`).
pub(super) fn take_pin_menu_dismiss() -> bool {
    let asked = PIN_MENU_DISMISS_REQUESTED.swap(false, Ordering::AcqRel);
    if !asked {
        return false;
    }

    let mut was_up = false;
    with_pin(|pin| {
        if pin.menu.open {
            was_up = true;
            pin.menu.open = false;
        }
    });

    // The ask being taken, as this side of the seam sees it: the
    // flag's value before the swap, and whether the menu was up to
    // put away. The hook's side of the same ask is the trace its
    // pinned tick writes (see `note_pin_click!`).
    crate::shell::explorer_hook::note_pin_click!(
        "DISMISS taken  asked {}  menu_up {}",
        asked as u8,
        was_up as u8,
    );

    was_up
}

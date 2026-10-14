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

/// The state of the card's menu: whether its panel is up, and which row's flyout is
/// up beside it.
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
    /// Whether the End of File row's flyout is up beside the menu, rather than the menu
    /// holding its rows alone. Only one flyout is ever up: opening one puts
    /// the other away (see `refresh_pin_menu`).
    pub(super) end_of_file: bool,
    /// Whether the Seek row's flyout is up beside the menu, for the same reason
    /// the End of File row's is.
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
    /// When the pointer came onto the End of File row, while the hover timer
    /// that opens the flyout is running: cleared the moment the
    /// pointer leaves the row, so that a hand only passing over it
    /// opens nothing (see `refresh_pin_menu`).
    pub(super) end_of_file_since: Option<Instant>,
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

/// How long the pointer rests on a row that opens a flyout before it opens:
/// the Windows-like hover delay, long enough that a hand on its way
/// across the row does not open a panel it did not mean to, and short
/// enough that a hand meaning to use the row does not wait (see
/// `refresh_pin_menu`).
const PIN_MENU_HOVER_DELAY: Duration = Duration::from_millis(150);

/// The rows of the card's menu: `End of File` and `Seek` are the doors its flyouts hang from,
/// the first and the last of them, and `Shuffle` between the two answers on the spot — a
/// check a press turns over, with no flyout of its own (see `menu_rows` and
/// `pinned_menu_press`). The same two doors are what the hover timers watch and what a press
/// on either opens (see `refresh_pin_menu`).
const END_OF_FILE_ROW: usize = 0;
/// The row of the card's menu that answers where it stands rather than opening a panel: the
/// check `Shuffle` carries (see `menu_rows`).
const SHUFFLE_ROW: usize = 1;

/// The row of the card's menu that opens the seek choices: the last of its rows, whatever
/// the panel holds (see `menu_rows`).
fn seek_row_of(rows: &[pin_chrome::MenuRow]) -> usize {
    rows.len().saturating_sub(1)
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

/// The rows the card's menu holds: the door that opens the end-of-file choices, the check
/// that says whether a forward step of the pin's walk shuffles, and the door that opens the
/// seek choices.
///
/// The main menu is always these three rows — the choices are the flyouts' own rows, beside
/// the menu's rather than in place of them (see `end_of_file_rows` and `seek_rows`).
///
/// The two doors carry no mark of their own: what each choice is, is the flyout's to say, and
/// what says a row opens one is the arrow at its own right (see `pin_chrome::MenuMark::Door`).
/// The row between them is the one that answers here, and it carries the check the setting has
/// now (see `pin_audio_shuffle`).
///
/// The rows a hand counts here are the rows themselves, with no rule between them: the doors
/// the flyouts hang from are the first and the last of these (see `pinned_menu_paint`).
pub(super) fn menu_rows() -> Vec<pin_chrome::MenuRow> {
    vec![
        pin_chrome::MenuRow {
            label: "End of File".to_string(),
            mark: pin_chrome::MenuMark::Door,
        },
        pin_chrome::MenuRow {
            label: "Shuffle".to_string(),
            mark: pin_chrome::MenuMark::Check(pin_audio_shuffle()),
        },
        pin_chrome::MenuRow {
            label: "Seek".to_string(),
            mark: pin_chrome::MenuMark::Door,
        },
    ]
}

/// The ends of file a pinned sound can have, in the order the tray's own `End of File`
/// submenu offers them (see `shell::tray::ids::PIN_END_OF_FILE_CHOICES`), held here
/// because that list is the tray's own and this is the card's.
const END_OF_FILE_CHOICES: [PinEndOfFile; 3] = [
    PinEndOfFile::Pause,
    PinEndOfFile::Loop,
    PinEndOfFile::Continue,
];

/// What each of the end-of-file choices is listed as: the words the tray lists
/// the same three under, without the mark that menu puts on the default — the
/// card's menu marks the one in force instead (see `end_of_file_rows`).
const END_OF_FILE_LABELS: [&str; 3] = ["Pause", "Loop", "Jump"];

/// The rows the End of File row's flyout holds: the three ends the pin can have, in the order
/// the tray lists them in, with the one the pin plays by carrying its disc.
///
/// All three are one radio group — the sound goes round, steps on, or is left where it ended —
/// so there is no switch above them and no rule between anything, and the first row of the
/// flyout is the first end a hand reads (see `pinned_menu_press`).
///
/// Read out of the configuration at the paint rather than remembered, so that a
/// change to the setting shows the moment it is made, from wherever it was made —
/// the same setting the tray's own `End of File` submenu writes (see `menu_rows`).
pub(super) fn end_of_file_rows() -> Vec<pin_chrome::MenuRow> {
    let current = pin_mode_audio_end_of_file();
    END_OF_FILE_CHOICES
        .iter()
        .zip(END_OF_FILE_LABELS)
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

/// What a pinned sound does at the end of its file — the same setting the tray's own
/// `End of File` submenu writes. A lock that cannot be taken is answered with the default
/// rather than with nothing, because a row that says nothing is a row a hand cannot read.
fn pin_mode_audio_end_of_file() -> PinEndOfFile {
    CONFIG
        .lock()
        .ok()
        .map(|config| config.pin_mode_audio_end_of_file)
        .unwrap_or(DEFAULT_PIN_MODE_AUDIO_END_OF_FILE)
}

/// Whether a forward step of a pinned sound's walk is a random sound of the folder rather than
/// the one beside it — the same setting the tray's own `Shuffle` row writes, asked the way the
/// end of file above it is.
fn pin_audio_shuffle() -> bool {
    CONFIG
        .lock()
        .ok()
        .map(|config| config.pin_mode_audio_shuffle)
        .unwrap_or(DEFAULT_PIN_MODE_AUDIO_SHUFFLE)
}

/// The end of file and the shuffle, and where a pinned sound starts, written
/// the way every setting this app's own menus write is: under the
/// configuration's lock and saved (see
/// `shell::tray::commands::setters`).
fn set_pin_audio_end_of_file(end_of_file: PinEndOfFile) {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_mode_audio_end_of_file = end_of_file;
        config.save();
    }
}

/// The check the `Shuffle` row carries, turned over where it is pressed: the row answers on
/// the spot rather than opening a flyout of its own (see `pinned_menu_press`).
fn toggle_pin_audio_shuffle() {
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_mode_audio_shuffle = !config.pin_mode_audio_shuffle;
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
/// flyout the row under the hand opens beside it.
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
    // The End of File row's flyout where it is up, else the Seek row's: only one flyout is
    // ever up, so the one that is up is the panel's (see `refresh_pin_menu`).
    let choices = if pin.menu.end_of_file {
        Some(end_of_file_rows())
    } else if pin.menu.seek {
        Some(seek_rows())
    } else {
        None
    };
    // Which row of the main menu the flyout hangs from: the row the hand opened it from,
    // which is the main menu's first for the end-of-file list and its last for the seeks (see
    // `menu_rows`).
    let hang_row = match pin.menu.end_of_file {
        true => END_OF_FILE_ROW,
        false => seek_row_of(&rows),
    };
    let (width, height) = pin.window_size();

    // The flyout's own placement is what says where the main
    // menu itself is, because the two panels are placed together
    // to keep both of them inside the window (see
    // `pin_chrome::menu_flyout_from_menu`). With the flyout down,
    // the main menu is placed alone — and carries the same arrow
    // the flyout's placement answers, both placements asking the
    // one tier the flyout opens in, so a row that opens one points the way
    // its flyout opens whether the flyout is up or not.
    let (popup, flyout) = match choices {
        Some(choices) => {
            let placed = pin_chrome::menu_flyout_from_menu(
                &pin_chrome::menu_popup_from_point(pin.menu.point, width, height, pin.dpi, &rows),
                width,
                height,
                pin.dpi,
                hang_row,
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
/// the pointer has rested on a row that opens a flyout long enough for it to
/// open; and whether a flyout that is up is still wanted. A pointer that merely
/// crosses such a row opens nothing — the timer is cleared the moment it
/// leaves — and a flyout that is up stays up while the pointer is on its row,
/// in the gap between the panels, or on the flyout itself, and goes when
/// the pointer is on another row or off the menu entirely (see
/// `pin_chrome::menu_flyout_gap_holds`). Opening one flyout puts the other
/// away, so only one is ever up.
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

    // The End of File row is the main menu's first and the Seek row its last: the rows the
    // flyouts hang from, whatever the panel holds (see `menu_rows`). The Shuffle row between
    // them opens nothing, so no timer watches it — a check is a press rather than a rest (see
    // `pinned_menu_press`). The rows counted here are the rows a hand counts, which for this
    // panel is every row of it.
    let end_of_file_row = END_OF_FILE_ROW;
    let seek_row = seek_row_of(&paint.rows);
    let on_end_of_file = hover_main == Some(end_of_file_row);
    let on_seek = hover_main == Some(seek_row);
    // And the row the flyout that is up hangs from, which is the row the crossing between
    // the two panels is measured at (see `pinned_menu_paint`).
    let hang_row = match pin.menu.end_of_file {
        true => end_of_file_row,
        false => seek_row,
    };

    // The timers: each starts when the pointer comes onto its row and is
    // cleared the moment it leaves, so a crossing opens nothing.
    let end_of_file_since = match on_end_of_file {
        true => pin.menu.end_of_file_since.or(Some(now)),
        false => None,
    };
    let seek_since = match on_seek {
        true => pin.menu.seek_since.or(Some(now)),
        false => None,
    };
    let rested_end_of_file = end_of_file_since
        .is_some_and(|since| now.duration_since(since) >= PIN_MENU_HOVER_DELAY);
    let rested_seek =
        seek_since.is_some_and(|since| now.duration_since(since) >= PIN_MENU_HOVER_DELAY);

    // Whether a flyout that is up is still wanted: the pointer is on its row,
    // the gap, or the flyout itself.
    let in_flyout = hover_flyout.is_some()
        || paint.flyout.as_ref().is_some_and(|flyout| {
            at.is_some_and(|(x, y)| {
                pin_chrome::menu_flyout_gap_holds(&paint.popup, &flyout.popup, hang_row, x, y)
            })
        });
    let kept_end_of_file = on_end_of_file || (pin.menu.end_of_file && in_flyout);
    let kept_seek = on_seek || (pin.menu.seek && in_flyout);

    let mut end_of_file = (pin.menu.end_of_file || rested_end_of_file) && kept_end_of_file;
    let mut seek = (pin.menu.seek || rested_seek) && kept_seek;
    // Only one flyout is ever up: a hand that has rested on one row has left the
    // other, so the rested one wins and the kept one goes with the row it hung from.
    if rested_end_of_file {
        seek = false;
    } else if rested_seek {
        end_of_file = false;
    }

    let changed = pin.menu.hover_main != hover_main
        || pin.menu.hover_flyout != hover_flyout
        || pin.menu.end_of_file != end_of_file
        || pin.menu.seek != seek
        || pin.menu.end_of_file_since != end_of_file_since
        || pin.menu.seek_since != seek_since;
    pin.menu.hover_main = hover_main;
    pin.menu.hover_flyout = hover_flyout;
    pin.menu.end_of_file = end_of_file;
    pin.menu.seek = seek;
    pin.menu.end_of_file_since = end_of_file_since;
    pin.menu.seek_since = seek_since;

    changed
}

/// A press on a pinned window with the card's menu up, answering whether
/// it was the menu's: a press on a door row opens its flyout beside the menu
/// and the panel stays up to hold it, a press on the `Shuffle` row turns the
/// check it carries over and leaves the panel up to show it, a row
/// of a flyout is the choice the pin plays by from then on and the
/// whole menu goes away with it — while a press anywhere else is a press
/// that has left the menu, which puts both panels away.
///
/// The rows a hand counts are what a press is answered against, and the door rows of the
/// main menu are the panel's first and its last (see `menu_rows`); there is no rule inside
/// either flyout, so its first row is its first choice (see `pin_chrome::menu_row_at`).
///
/// Nothing is held, because a menu is a thing a
/// hand reads rather than a thing it carries: the press is answered
/// where it landed, and the panel is one paint away from being gone (see
/// `open_pin_menu`). The pointer itself is held while the menu is up —
/// taken in `open_pin_menu` and let go below with the menu — so a press
/// anywhere outside the window still arrives here to be answered.
pub(super) unsafe fn pinned_menu_press(hwnd: HWND, x: i32, y: i32) -> bool {
    let Some(paint) = pinned_menu_geometry() else {
        return false;
    };

    // The row the press landed on, if it landed on one: each panel's own
    // rows are what a press is answered against, and the pad around them
    // and everywhere outside both panels is not a row (see
    // `pin_chrome::menu_row_at`).
    let stays_up = match pin_chrome::menu_row_at(&paint.popup, x, y) {
        // The door the end-of-file choices come through: a press on it opens its flyout
        // beside the menu and the menu stays up, its rows still showing — the same doors the
        // hover opens, for a hand that has come to press rather than wait. Opening one puts
        // the other away, so only one is ever up.
        Some(END_OF_FILE_ROW) => {
            with_pin(|pin| {
                pin.menu.end_of_file = true;
                pin.menu.seek = false;
            });
            true
        }
        // The row that answers on the spot: a press turns the check over, and the panel
        // stays up with the mark it now carries — the one row of this menu a press does not
        // put away, because there is nothing to choose from and the mark is the answer.
        Some(SHUFFLE_ROW) => {
            toggle_pin_audio_shuffle();
            true
        }
        Some(_) => {
            with_pin(|pin| {
                pin.menu.seek = true;
                pin.menu.end_of_file = false;
            });
            true
        }
        None => match paint
            .flyout
            .as_ref()
            .and_then(|flyout| pin_chrome::menu_row_at(&flyout.popup, x, y))
        {
            // A row of the flyout: the one pressed is the one the pin
            // plays by from then on, and the whole menu goes away with
            // the answer — both panels, because the flyout is the
            // menu's own — the way a menu a choice was made in does. The
            // End of File flyout's rows are the ends of file; the Seek flyout's are the seeks.
            Some(index) => {
                let end_of_file_open = pin_state()
                    .and_then(|pinned| pinned.pin().map(|pin| pin.menu.end_of_file))
                    .unwrap_or(false);
                if end_of_file_open {
                    if let Some(end_of_file) = END_OF_FILE_CHOICES.get(index) {
                        set_pin_audio_end_of_file(*end_of_file);
                    }
                } else if let Some(choice) = SEEK_CHOICES.get(index) {
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
        // The menu held the pointer while it was up, so it lets it go
        // with it: a window still holding the pointer after its menu is
        // gone is a window every click on the desktop is delivered to
        // instead of whatever the click was aimed at.
        release_pin_capture(hwnd);
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
/// A menu put up is put up on its main page, both
/// flyouts put away: a hand that opens the menu has
/// come for the choices, and they are one row away.
pub(super) unsafe fn open_pin_menu(hwnd: HWND, x: i32, y: i32) {
    let audio = pin_state()
        .and_then(|pinned| pinned.pin().map(pin_shows_an_audio_card))
        .unwrap_or(false);
    if !audio {
        return;
    }

    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.end_of_file = false;
        pin.menu.seek = false;
        pin.menu.point = (x, y);
        pin.menu.hover_main = None;
        pin.menu.hover_flyout = None;
        pin.menu.end_of_file_since = None;
        pin.menu.seek_since = None;
    });
    // The menu is this window's own, so the window takes the pointer while
    // it is up: a press anywhere else — another app, the taskbar — arrives
    // here first, and the first one puts the menu away rather than reaching
    // through to whatever was aimed at. Without it no outside press ever
    // reaches this window at all, and only the hook's poll sees one a tick
    // later, which is the press the menu stayed up through.
    let _ = SetCapture(hwnd);
    render_layered_preview(hwnd);
}

/// Whether the card's menu is up: what an outside press is answered against
/// before anything the press would otherwise do (see `window_proc`).
pub(super) fn pin_menu_is_open() -> bool {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.menu.open))
        .unwrap_or(false)
}

/// Whether a point in the window's own coordinates is outside the window:
/// what a press the menu's capture delivered from elsewhere is told apart
/// by. A captured press aimed at another app or the taskbar arrives here
/// with coordinates outside the window's own box, and answering it as a
/// press on the card would drag the window for a hand on another window.
pub(super) fn pin_menu_point_is_outside(x: i32, y: i32) -> bool {
    let Some(size) = pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.window_size()))
    else {
        return false;
    };
    x < 0 || y < 0 || x >= size.0 || y >= size.1
}

/// Put the card's menu away where it is up, answering whether it came down:
/// the whole menu goes — both panels — and the pointer with it, while the
/// pin itself is left standing. A press that arrived from elsewhere is what
/// this answers; a repaint is owed where it did.
pub(super) unsafe fn close_pin_menu(hwnd: HWND) -> bool {
    let mut was_up = false;
    with_pin(|pin| {
        if pin.menu.open {
            was_up = true;
            pin.menu.open = false;
        }
    });
    if was_up {
        release_pin_capture(hwnd);
        render_layered_preview(hwnd);
    }
    was_up
}

/// Take the ask the Explorer hook left for a press that landed on another window and put the
/// card's menu away where it is up, answering whether it came down — a repaint is owed where
/// it did. It is the preview loop's own read of the ask, taken on its tick (see
/// `PIN_MENU_DISMISS_REQUESTED`), and only the menu goes: the pin itself is left standing,
/// which is what a press on another window means here (see
/// `requests::request_pin_menu_dismiss`).
///
/// The bubble's own menu goes the same way and from the same ask, because with the pin's
/// window hidden the bubble is the pin's own window: a press the hook read was outside the
/// bubble's as surely as it was outside the card's (see `bubble_window_is_at` and
/// `close_pin_bubble_menu`).
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
    let bubble_was_up = unsafe { close_pin_bubble_menu() };

    // The ask being taken, as this side of the seam sees it: the
    // flag's value before the swap, and whether a menu was up to
    // put away. The hook's side of the same ask is the trace its
    // pinned tick writes (see `note_pin_click!`).
    crate::shell::explorer_hook::note_pin_click!(
        "DISMISS taken  asked {}  menu_up {}",
        asked as u8,
        (was_up || bubble_was_up) as u8,
    );

    was_up || bubble_was_up
}

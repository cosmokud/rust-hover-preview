//! The menu a pinned window opens from a right-click on it: the state of
//! it, the rows it holds, and what a press on one of them does.
//!
//! A sound's card holds a menu of its own — the two ways out of the pin,
//! the playback of the file in it, and the settings it plays by, two of them
//! behind flyouts — while every other kind holds the same rows a bubble's
//! menu holds, down to the two places a minimize can put the pin (see
//! `bubble_menu_rows`). Both are the window's own rather than the tray's, so
//! their rows are read out of the same settings the tray's own `Pin Mode`
//! submenus write, and their panels are drawn by `pin_chrome` rather than by
//! the menu the shell draws: a menu that floats over the media is a panel
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
    /// Which page of the menu the panel holds: the chrome's own rows, or the two commands a
    /// pinned page of text stands above them (see `PinMenuPage`).
    pub(super) page: PinMenuPage,
    /// Whether the End of File row's flyout is up beside the menu, rather than the menu
    /// holding its rows alone. Only one flyout is ever up: opening one puts
    /// the other away (see `refresh_pin_menu`).
    pub(super) end_of_file: bool,
    /// Whether the Seek row's flyout is up beside the menu, for the same reason
    /// the End of File row's is.
    pub(super) seek: bool,
    /// Where the right-click that opened the panel landed, on the screen: the
    /// panel's top-left, held inside the display's own work area by the placement (see
    /// `pin_chrome::menu_popup_from_point`).
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

/// Which page a right-click asks a pinned window's own menu for: the rows its chrome holds, or —
/// over the text of a pinned page — those rows under the page's own two commands, so the text can
/// be selected and copied without the window's menu being a page away.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(super) enum PinMenuPage {
    /// The window's own rows, which is what a right-click on the title bar asks for.
    #[default]
    Chrome,
    /// A pinned page's text, whose own `Select All` and `Copy` stand above the window's rows.
    Text,
}

impl PinMenuPage {
    /// How many of the panel's rows are the page's own, above the rows the pin shares with the
    /// bubble.
    fn own_rows(self) -> usize {
        match self {
            PinMenuPage::Chrome => 0,
            PinMenuPage::Text => TEXT_MENU_ROWS.len(),
        }
    }
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
    /// The scale the panels were laid out at, which is the scale they are painted at: the
    /// card's own, so that a panel's pads, rows and marks are the size its placement
    /// measured them (see `pin_chrome::menu_panel_sizes`).
    pub(super) dpi: u32,
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

/// What a pinned page's own page of the menu offers above the window's rows: the whole frame
/// selected, what is selected put on the clipboard, and the next place it appears — the same three
/// commands a hover's own text menu holds the first two of (see `show_text_preview_menu`).
#[derive(Clone, Copy, PartialEq, Eq)]
enum TextMenuRow {
    SelectAll,
    Copy,
    /// The page taken to the next place what is selected appears in it, which is the one find a
    /// window nobody may type in can offer: the selection is the query (see `find_next`).
    FindNext,
    /// The rule between the page's own commands and the window's rows under them.
    Separator,
}

/// The rows a pinned page's own page of the menu holds, from the top of the panel down, above the
/// rows the pin shares with the bubble.
const TEXT_MENU_ROWS: [TextMenuRow; 4] = [
    TextMenuRow::SelectAll,
    TextMenuRow::Copy,
    TextMenuRow::FindNext,
    TextMenuRow::Separator,
];

impl TextMenuRow {
    /// What the row is written as: a command's own name, and nothing at all on a rule.
    fn label(self) -> &'static str {
        match self {
            TextMenuRow::SelectAll => "Select All",
            TextMenuRow::Copy => "Copy",
            TextMenuRow::FindNext => "Find Next",
            TextMenuRow::Separator => "",
        }
    }

    /// The mark the row carries: a command is a row that answers on the spot, and a rule carries
    /// the hairline between the two groups (see `pin_chrome::MenuMark`).
    fn mark(self) -> pin_chrome::MenuMark {
        match self {
            TextMenuRow::Separator => pin_chrome::MenuMark::Separator,
            TextMenuRow::SelectAll | TextMenuRow::Copy | TextMenuRow::FindNext => {
                pin_chrome::MenuMark::None
            }
        }
    }
}

/// The rows a pinned page's own page of the menu holds, above the rows the pin shares with the
/// bubble.
fn text_menu_rows() -> Vec<pin_chrome::MenuRow> {
    TEXT_MENU_ROWS
        .iter()
        .map(|row| pin_chrome::MenuRow {
            label: row.label().to_string(),
            mark: row.mark(),
        })
        .collect()
}

/// What a press on one of a pinned page's own rows asks, answering whether the panel stays up:
/// every command answers and goes away the way a menu a command was chosen in does, while a rule is
/// not a row a hand answers at all.
unsafe fn text_menu_row_press(hwnd: HWND, row: TextMenuRow) -> bool {
    match row {
        TextMenuRow::SelectAll => select_all_text_preview(hwnd),
        TextMenuRow::Copy => {
            copy_text_preview(hwnd);
        }
        TextMenuRow::FindNext => find_next_in_text_preview(hwnd),
        TextMenuRow::Separator => return true,
    }

    false
}

/// How long the pointer rests on a row that opens a flyout before it opens:
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
const SEEK_LABELS: [&str; 4] = ["Remember", "From the Start", "From the Middle", "Random"];

/// One row of the card's menu, named rather than numbered: the rows a hand counts are the
/// rows themselves, and the number one of them sits at moves the moment another is put in
/// above it. A press, a hover and a flyout are all answered against the name, which is the
/// thing that says what the row does (see `PIN_MENU_ROWS`).
#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum PinMenuRow {
    /// Collapse the pin into the round bubble it leaves on the desktop.
    ToBubble,
    /// Collapse the pin into the app's own stand-in in the taskbar.
    ToTaskbar,
    /// Take the pin down.
    Close,
    /// The rule between two groups of rows: one between the ways out of the pin and the
    /// playback it holds, and one between that and the settings the pin plays by.
    Separator,
    /// Hold what is playing where it stands, or set it going again: the same pause the card's
    /// clock and a Space in the window answer. It is named for what it will do, the way the
    /// bubble's own hold is (see `menu_rows`).
    Pause,
    /// The file after the one pinned, the same walk the card's own arrow takes.
    NextFile,
    /// The file before the one pinned.
    PreviousFile,
    /// Whether a forward step of the pin's walk shuffles: the one row that answers where it
    /// stands rather than opening a panel.
    Shuffle,
    /// The door the ends of file are chosen through.
    EndOfFile,
    /// The door the seeks are chosen through.
    Seek,
}

/// The rows the card's menu holds, from the top of the panel to the bottom: the two ways out
/// of the pin, the rule, the playback of the file in it, the rule again, and the three
/// settings the pin plays by — the shuffle that answers on the spot, and the two doors the
/// flyouts hang from.
///
/// The same two doors are what the hover timers watch and what a press on either opens (see
/// `refresh_pin_menu`).
const PIN_MENU_ROWS: [PinMenuRow; 11] = [
    PinMenuRow::ToBubble,
    PinMenuRow::ToTaskbar,
    PinMenuRow::Close,
    PinMenuRow::Separator,
    PinMenuRow::Pause,
    PinMenuRow::NextFile,
    PinMenuRow::PreviousFile,
    PinMenuRow::Separator,
    PinMenuRow::Shuffle,
    PinMenuRow::EndOfFile,
    PinMenuRow::Seek,
];

impl PinMenuRow {
    /// What the row is written as: a rule carries nothing at all, because there is nothing on
    /// one to read. The hold is named for what it will do rather than for what it already does,
    /// the way the bubble's own hold is (see `bubble_menu_row_label`): `Pause` while the sound is
    /// playing and `Play` while it is held — the same answer the card's own play button gives
    /// (see `a_sound_is_playing`).
    fn label(self, playing: bool) -> &'static str {
        match self {
            PinMenuRow::ToBubble => "To Bubble",
            PinMenuRow::ToTaskbar => "To Taskbar",
            PinMenuRow::Close => "Close",
            PinMenuRow::Separator => "",
            PinMenuRow::Pause if playing => "Pause",
            PinMenuRow::Pause => "Play",
            PinMenuRow::NextFile => "Next File",
            PinMenuRow::PreviousFile => "Previous File",
            PinMenuRow::Shuffle => "Shuffle",
            PinMenuRow::EndOfFile => "End of File",
            PinMenuRow::Seek => "Seek",
        }
    }

    /// The mark the row carries: a rule, a door to a flyout, the check the shuffle is the face
    /// of, and nothing at all for a row that answers on the spot.
    fn mark(self) -> pin_chrome::MenuMark {
        match self {
            PinMenuRow::Separator => pin_chrome::MenuMark::Separator,
            PinMenuRow::EndOfFile | PinMenuRow::Seek => pin_chrome::MenuMark::Door,
            PinMenuRow::Shuffle => pin_chrome::MenuMark::Check(pin_audio_shuffle()),
            PinMenuRow::ToBubble
            | PinMenuRow::ToTaskbar
            | PinMenuRow::Close
            | PinMenuRow::Pause
            | PinMenuRow::NextFile
            | PinMenuRow::PreviousFile => pin_chrome::MenuMark::None,
        }
    }
}

/// Where a row of the card's menu sits in the panel, by name rather than by the number it
/// happens to be at: what a press and a hover are answered against (see `PIN_MENU_ROWS`).
fn row_of(row: PinMenuRow) -> Option<usize> {
    PIN_MENU_ROWS.iter().position(|held| *held == row)
}

/// The rows the card's menu holds: the two ways out of the pin, the playback of the file in
/// it, and the settings the pin plays by — the door that opens the end-of-file choices, the
/// check that says whether a forward step of the pin's walk shuffles, and the door that opens
/// the seek choices.
///
/// The two doors carry no mark of their own: what each choice is, is the flyout's to say, and
/// what says a row opens one is the arrow at its own right (see `pin_chrome::MenuMark::Door`).
/// The row between them is the one that answers here, and it carries the check the setting has
/// now (see `pin_audio_shuffle`).
///
/// Read out of the configuration at the paint rather than remembered, so that a
/// change to the setting shows the moment it is made, from wherever it was made —
/// the same setting the tray's own `End of File` submenu writes (see `menu_rows`).
///
/// The hold is named at the paint as well, from what the player behind the card is doing:
/// `Pause` while the sound is playing and `Play` while it is held (see
/// `a_sound_is_playing`).
///
/// The rows a hand counts here are the rows themselves, with no rule between them: the doors
/// the flyouts hang from are named rather than numbered (see `row_of` and
/// `pinned_menu_paint`).
pub(super) fn menu_rows() -> Vec<pin_chrome::MenuRow> {
    let playing = a_sound_is_playing();
    PIN_MENU_ROWS
        .iter()
        .map(|row| pin_chrome::MenuRow {
            label: row.label(playing).to_string(),
            mark: row.mark(),
        })
        .collect()
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

/// Whether the pin that is up has the card's menu up, or nothing for a pin that is not there.
///
/// It is what the player's own raises are held off by while the panel is up (see `player_order`):
/// the panel floats in a window of this app's own over the film, so a raise puts the film over it,
/// and the panel is back only on the tick's own `render_pin_menu` — a menu that blinks wherever it
/// overlaps the picture. The same question as `pin_volume_open`, asked of the other panel a pin
/// floats (see `pin_menu_window`).
pub(super) fn pin_menu_open() -> bool {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.menu.open && !pin.collapsed))
        .unwrap_or(false)
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

    pin.menu.open.then(|| pinned_menu_paint(pin))
}

/// The menu as it is painted: the rows the main
/// page holds, the panel anchored at the point the
/// right-click that opened it landed, and the
/// flyout the row under the hand opens beside it.
///
/// The panel is anchored at the pin's own remembered click point, on the
/// screen, and held inside the display's own work area by the placement (see
/// `pin_chrome::menu_popup_from_point`), so a panel opened on a hand near an edge is
/// a panel a hand can still reach.
fn pinned_menu_paint(pin: &PinnedPreview) -> PinMenuPaint {
    // The rows the pin's menu is the face of: a sound's card holds its own — the two ways out of
    // the pin, the playback of the file in it, and the settings it plays by — while every other
    // kind holds the rows a bubble's menu holds, which is the same menu under the same press
    // (see `bubble_menu_chrome_rows`).
    let card = pin_shows_an_audio_card(pin);
    let rows = match card {
        true => menu_rows(),
        // A right-click in the text of a pinned page holds the page's own two commands above the
        // rows the pin shares with the bubble (see `PinMenuPage::Text`).
        false => {
            let mut rows = match pin.menu.page {
                PinMenuPage::Chrome => Vec::new(),
                PinMenuPage::Text => text_menu_rows(),
            };
            rows.extend(bubble_menu_chrome_rows(
                current_media_type(),
                pin_media_is_playing(pin),
            ));
            rows
        }
    };
    // The End of File row's flyout where it is up, else the Seek row's: only one flyout is
    // ever up, so the one that is up is the panel's, and both hang off rows only a sound's
    // card carries (see `refresh_pin_menu`).
    let choices = if !card {
        None
    } else if pin.menu.end_of_file {
        Some(end_of_file_rows())
    } else if pin.menu.seek {
        Some(seek_rows())
    } else {
        None
    };
    // Which row of the main menu the flyout hangs from: the row the hand opened it from,
    // which is the named door that flyout hangs from, wherever the panel now holds it (see
    // `row_of`).
    let hang_row = row_of(match pin.menu.end_of_file {
        true => PinMenuRow::EndOfFile,
        false => PinMenuRow::Seek,
    })
    .unwrap_or(0);
    // The panel is placed on the screen rather than inside the card: the window it is drawn
    // in is one of its own over the display's work area, so the room a panel is held
    // inside is the display's own and not the card's — which is what lets a menu taller
    // than the card keep every one of its rows, and what keeps the flyout on the right of
    // the menu until the right edge of the screen says otherwise (see `pin_menu_window`).
    // The placement is asked in the work area's own coordinates and the panels are moved
    // onto the screen afterwards, because a press on one is answered in the screen's own.
    let screen = work_area_at(pin.menu.point.0, pin.menu.point.1);
    let point = (
        pin.menu.point.0 - screen.left,
        pin.menu.point.1 - screen.top,
    );
    let width = screen.right - screen.left;
    let height = screen.bottom - screen.top;

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
                &pin_chrome::menu_popup_from_point(point, width, height, pin.dpi, &rows),
                width,
                height,
                pin.dpi,
                hang_row,
                &choices,
            );
            (
                pin_chrome::menu_popup_shifted(&placed.menu, screen.left, screen.top),
                Some(PinMenuFlyout {
                    popup: pin_chrome::menu_popup_shifted(&placed.popup, screen.left, screen.top),
                    rows: choices,
                    hover: pin.menu.hover_flyout,
                }),
            )
        }
        None => (
            pin_chrome::menu_popup_shifted(
                &pin_chrome::menu_popup_from_point(point, width, height, pin.dpi, &rows),
                screen.left,
                screen.top,
            ),
            None,
        ),
    };

    PinMenuPaint {
        popup,
        rows,
        hover: pin.menu.hover_main,
        flyout,
        dpi: pin.dpi,
    }
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
/// The pointer is read as a place on the screen, which is the coordinates the
/// panels are held in: they are drawn in a window of their own over the screen rather
/// than inside the card, so what the pointer has left is the menu itself rather than a
/// window it was standing over — and a wash or a flyout left up under a hand that is
/// gone is a mark on nothing.
pub(super) fn refresh_pin_menu(
    pin: &mut PinnedPreview,
    now: Instant,
    cursor: Option<(i32, i32)>,
) -> bool {
    if !pin.menu.open {
        return false;
    }

    let paint = pinned_menu_paint(pin);

    let at = cursor;

    let hover_main = at.and_then(|(x, y)| pin_chrome::menu_row_at(&paint.popup, x, y));
    let hover_flyout = paint
        .flyout
        .as_ref()
        .and_then(|flyout| at.and_then(|(x, y)| pin_chrome::menu_row_at(&flyout.popup, x, y)));

    // The End of File row and the Seek row are the doors the flyouts hang from, named rather
    // than numbered, whatever the panel holds (see `row_of`). The rows between them answer on
    // the spot or rule the two groups apart, so no timer watches them — a check is a press
    // rather than a rest, and a rule is nothing at all (see `pinned_menu_press`). The rows
    // counted here are the rows a hand counts, which for this
    // panel is every row of it.
    let on_end_of_file = row_of(PinMenuRow::EndOfFile).is_some_and(|row| hover_main == Some(row));
    let on_seek = row_of(PinMenuRow::Seek).is_some_and(|row| hover_main == Some(row));
    // And the row the flyout that is up hangs from, which is the row the crossing between
    // the two panels is measured at (see `pinned_menu_paint`).
    let hang_row = match pin.menu.end_of_file {
        true => row_of(PinMenuRow::EndOfFile),
        false => row_of(PinMenuRow::Seek),
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
    let rested_end_of_file =
        end_of_file_since.is_some_and(|since| now.duration_since(since) >= PIN_MENU_HOVER_DELAY);
    let rested_seek =
        seek_since.is_some_and(|since| now.duration_since(since) >= PIN_MENU_HOVER_DELAY);

    // Whether a flyout that is up is still wanted: the pointer is on its row,
    // the gap, or the flyout itself.
    let in_flyout = hover_flyout.is_some()
        || paint.flyout.as_ref().is_some_and(|flyout| {
            hang_row.is_some_and(|hang_row| {
                at.is_some_and(|(x, y)| {
                    pin_chrome::menu_flyout_gap_holds(&paint.popup, &flyout.popup, hang_row, x, y)
                })
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

/// A press on a pinned window with its menu up, answering whether it was the
/// menu's: a press on a row is that row's own and the menu answers it, while a
/// press anywhere else is a press that has left the panel, which puts the menu away.
///
/// Which rows a press is answered against is the pin's own answer as much as the
/// panel's is: a sound's card holds the card's menu, with its two doors and their
/// flyouts (see `card_menu_press`), while every other kind holds the rows it shares
/// with a bubble (see `shared_menu_press`).
///
/// Nothing is held, because a menu is a thing a
/// hand reads rather than a thing it carries: the press is answered
/// where it landed, and the panel is one paint away from being gone (see
/// `open_pin_menu`). The pointer itself is held while the menu is up —
/// taken in `open_pin_menu` and let go below with the menu — so a press
/// anywhere on the screen still arrives here to be answered, over whatever
/// window it was aimed at.
///
/// `x` and `y` are a place on the screen rather than in the card's own
/// window, because the panels are: the menu is drawn in a window of its own
/// over the screen (see `pin_menu_window`), and the card's own window is
/// the one holding the pointer that carried the press here.
pub(super) unsafe fn pinned_menu_press(hwnd: HWND, x: i32, y: i32) -> bool {
    let Some(paint) = pinned_menu_geometry() else {
        return false;
    };

    // Which menu the pin's own is, asked of the pin's own three facts rather than of the media's
    // kind, because those are what the rows were built from at the paint (see
    // `pin_shows_an_audio_card`).
    let card = pin_state()
        .and_then(|pinned| pinned.pin().map(pin_shows_an_audio_card))
        .unwrap_or(false);
    let stays_up = match card {
        true => card_menu_press(&paint, x, y),
        false => shared_menu_press(hwnd, &paint, x, y),
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
    // And the window the menu was drawn in goes away with the panels that were
    // answered, or is painted again where a row left it standing — a flyout
    // opened, a check turned over (see `render_pin_menu`).
    render_pin_menu();

    true
}

/// What a press on a row of the card's own menu means: a press on a door row opens its flyout
/// beside the menu and the panel stays up to hold it, a press on the `Shuffle` row turns the
/// check it carries over and leaves the panel up to show it, a row of a flyout is the choice
/// the pin plays by from then on and the whole menu goes away with it — while a press anywhere
/// else is a press that has left the menu, which puts both panels away.
///
/// The rows a hand counts are what a press is answered against, by name rather than by
/// number (see `row_of`); there is no rule inside
/// either flyout, so its first row is its first choice (see `pin_chrome::menu_row_at`).
///
/// Answers whether the panel stays up.
unsafe fn card_menu_press(paint: &PinMenuPaint, x: i32, y: i32) -> bool {
    // The row the press landed on, if it landed on one: each panel's own
    // rows are what a press is answered against, and the pad around them
    // and everywhere outside both panels is not a row (see
    // `pin_chrome::menu_row_at`).
    let pressed = pin_chrome::menu_row_at(&paint.popup, x, y)
        .and_then(|index| PIN_MENU_ROWS.get(index).copied());

    match pressed {
        // The door the end-of-file choices come through: a press on it opens its flyout
        // beside the menu and the menu stays up, its rows still showing — the same doors the
        // hover opens, for a hand that has come to press rather than wait. Opening one puts
        // the other away, so only one is ever up.
        Some(PinMenuRow::EndOfFile) => {
            with_pin(|pin| {
                pin.menu.end_of_file = true;
                pin.menu.seek = false;
            });
            true
        }
        // The other door, the same bargain the other way.
        Some(PinMenuRow::Seek) => {
            with_pin(|pin| {
                pin.menu.seek = true;
                pin.menu.end_of_file = false;
            });
            true
        }
        // The row that answers on the spot: a press turns the check over, and the panel
        // stays up with the mark it now carries — the one row of this menu a press does not
        // put away, because there is nothing to choose from and the mark is the answer.
        Some(PinMenuRow::Shuffle) => {
            toggle_pin_audio_shuffle();
            true
        }
        // The two ways out of the pin are collapses that take the card off the screen with
        // them, so the menu goes away with the window it was opened over.
        Some(PinMenuRow::ToBubble) => {
            ask_pin(PinCommand::MinimizeTo(PinMinimizeTo::Bubble));
            false
        }
        Some(PinMenuRow::ToTaskbar) => {
            ask_pin(PinCommand::MinimizeTo(PinMinimizeTo::Taskbar));
            false
        }
        // Take the pin down: the same close the tray and the card's own close button ask for.
        Some(PinMenuRow::Close) => {
            ask_pin(PinCommand::Close);
            false
        }
        // The playback of the file in the card: every one of them goes through the loop the
        // way the key that asks for the same does, because the answer is a player begun or
        // ended rather than a state this window can set.
        Some(PinMenuRow::Pause) => {
            ask_pin(PinCommand::TogglePlayback);
            false
        }
        Some(PinMenuRow::NextFile) => {
            ask_pin(PinCommand::Next);
            false
        }
        Some(PinMenuRow::PreviousFile) => {
            ask_pin(PinCommand::Previous);
            false
        }
        // A rule is not a row a hand answers, so a press on one is nothing at all: the menu
        // stays exactly where it is.
        Some(PinMenuRow::Separator) => true,
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
    }
}

/// What a press on the menu a pinned window shares with the bubble means: the rows a bubble's
/// menu holds, answered the same way on both menus wherever the two meet (see
/// `shared_menu_row_press`) — with the one row the two walk differently, `To Taskbar`, putting
/// this pin away into the app's own place in the taskbar from the window it is standing in,
/// which is where the pin's own minimize button can put it (see `PinCommand::MinimizeTo`).
///
/// No row of this menu opens a flyout, so the panel is the whole of what a press is answered
/// against and a press outside it has left the menu (see `pinned_menu_paint`).
///
/// Answers whether the panel stays up.
unsafe fn shared_menu_press(hwnd: HWND, paint: &PinMenuPaint, x: i32, y: i32) -> bool {
    let page = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.menu.page))
        .unwrap_or_default();

    // The page's own rows stand above the ones the pin shares with the bubble, so a press on one
    // of them is the page's own and is answered before the shared rows are counted from the top
    // of the panel (see `PinMenuPage::own_rows`).
    let pressed_row = pin_chrome::menu_row_at(&paint.popup, x, y);
    if let Some(row) = pressed_row.and_then(|index| TEXT_MENU_ROWS.get(index)) {
        if page == PinMenuPage::Text {
            return text_menu_row_press(hwnd, *row);
        }
    }

    // The rows the panel was painted from are the rows a press is answered against, read the
    // way the paint reads them so that an index into the panel is an index into the pin's own
    // rows (see `bubble_menu_rows`).
    let pressed = pressed_row
        .and_then(|index| index.checked_sub(page.own_rows()))
        .and_then(|index| bubble_menu_rows(current_media_type()).get(index).copied());

    match pressed {
        Some(row) => shared_menu_row_press(row, || {
            ask_pin(PinCommand::MinimizeTo(PinMinimizeTo::Taskbar))
        }),
        // A press anywhere else is a press outside the panel, which puts the menu away: the
        // panel is over the window, and a hand that has come for what is under it is a hand
        // that has left the menu.
        None => false,
    }
}

/// The card's menu put up by a right-click on the pin's window, at the
/// point the right-click landed: a second right-click while the panel is
/// up moves it to the new point rather than putting it away.
///
/// Whether the point a right-click landed at is one the menu opens for, and which of its pages it
/// asks for, is not asked here but of the caller, which is the one that knows the window it landed
/// in: a point on a film's transport bar belongs to the player, and one on the text of a pinned
/// page asks for the page's own rows above the window's (see `pin_right_click_page`).
///
/// A menu put up is put up on the page the caller names and on its main rows alone,
/// both flyouts put away: a hand that opens the menu has
/// come for the choices, and they are one row away.
///
/// `x` and `y` are a place on the screen rather than in the card's own
/// window, which is the window the panel is placed against and the window
/// the press on it is answered in (see `pin_menu_window`).
pub(super) unsafe fn open_pin_menu(hwnd: HWND, x: i32, y: i32, page: PinMenuPage) {
    with_pin(|pin| {
        pin.menu.open = true;
        pin.menu.page = page;
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
    render_pin_menu();
}

/// Which page of the window's own menu, if any, a right-click in row `y` of a pinned window asks
/// for, in the window's own coordinates.
///
/// The question is a row's rather than a point's because the two strips that are not the menu's
/// are as wide as the window: a kind that plays has the player's own bar along the bottom, where
/// a right-click is the bar's (see `pinned_transport_geometry`).
///
/// A pinned page of text is the one kind whose media is the window's to answer for: a right-click
/// in it opens the menu's text page — the page's own `Select All` and `Copy` above the window's
/// rows — rather than the menu the text holds on a hover, which is the page's own only while
/// there is no window around it (see `PinMenuPage`).
///
/// A sound's card is the one kind with neither a bar nor a caption — the card is all of it, its
/// own title and controls drawn inside it — so every row of it is the menu's, which is where the
/// menu came from (see `pin_shows_an_audio_card`).
///
/// A kind whose media is a window of somebody else's — FFmpeg's player, the browser an engine page
/// is drawn by — never sends the band's right-clicks here at all: the window standing in the band
/// takes them, and the caption above it is the way to this menu.
pub(super) fn pin_right_click_page(y: i32) -> Option<PinMenuPage> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;

    if pin_shows_an_audio_card(pin) {
        return Some(PinMenuPage::Chrome);
    }

    // The bar a film plays by, which is the strip along the bottom of its window: the same band a
    // press is asked about, read off the pin already in hand rather than through
    // `pinned_transport_geometry`, which would take this very state's lock again (see
    // `pinned_transport_press`).
    if pinned_transport_bar_of(pin).is_some_and(|bar| bar.wanted && y >= bar.top) {
        return None;
    }

    // The text under the chrome: a right-click there is the page's own, with the window's rows
    // under its two commands rather than a menu of its own.
    if y >= pin.caption && TEXT_PREVIEW_HOLDING.load(Ordering::Acquire) {
        return Some(PinMenuPage::Text);
    }

    Some(PinMenuPage::Chrome)
}

/// Whether the card's menu is up: what an outside press is answered against
/// before anything the press would otherwise do (see `window_proc`).
pub(super) fn pin_menu_is_open() -> bool {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.menu.open))
        .unwrap_or(false)
}

/// Whether a point on the screen is on neither of the card's menu's panels:
/// what a press the menu's capture delivered from elsewhere is told apart by.
/// A captured press aimed at another app or the taskbar arrives here from a
/// place the menu holds nothing at, and answering it as a press on the card
/// would drag the window for a hand on another window.
///
/// It is asked of the panels rather than of the card's own box, because the
/// panels are no longer the card's: they are drawn in a window of their own
/// over the screen, and a hand on the card beside the menu has left the menu
/// exactly as a hand on another app has.
pub(super) fn pin_menu_point_is_outside(x: i32, y: i32) -> bool {
    let Some(paint) = pinned_menu_geometry() else {
        return false;
    };
    let on_a_row = pin_chrome::menu_row_at(&paint.popup, x, y).is_some()
        || paint
            .flyout
            .as_ref()
            .is_some_and(|flyout| pin_chrome::menu_row_at(&flyout.popup, x, y).is_some());
    !on_a_row
}

/// Put the card's menu away where it is up, answering whether it came down:
/// the whole menu goes — both panels — and the pointer with it, and the window
/// they were drawn in with them, while the
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
        render_pin_menu();
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
///
/// The window the card's menu was drawn in goes with it, because a panel nothing will ever
/// paint again is a window the desktop keeps for nothing (see `render_pin_menu`).
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
    unsafe { render_pin_menu() };

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

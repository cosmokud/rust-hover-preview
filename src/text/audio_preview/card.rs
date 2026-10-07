//! What a sound is previewed as: the options a caller hands in, the facts a track holds, the
//! name a card is headed with, the scroll a name the card has no room for is drawn with, and
//! the two calls that put the card on screen (`measure` and `render`).
//!
//! The questions about *where* on a card something is are answered here too — where the bar is,
//! which of a pinned card's own controls a point is on, where one of them is drawn — out of the
//! same arithmetic `page` lays the row out with. They are here rather than in the page because
//! they are what a caller outside asks: a button hit-tested by one arithmetic and drawn by
//! another answers a press in the middle of the facts line, so the boxes are rebuilt from the
//! page's own row rather than kept beside the card.

use super::page::{bar_band, bar_row, build_page, control_boxes, holds, paint, text_width};
use crate::config::config::TextTheme;
use crate::readers::audio_track::Track;
use crate::text::text_paint::{DibSurface, TextMetrics};
use crate::text::text_theme;
use std::path::Path;
use std::time::{Duration, Instant};
use windows::Win32::Foundation::RECT;
use windows::Win32::Graphics::Gdi::{CreateCompatibleDC, DeleteDC};

/// The mark the card opens with: a round bullet rather than an icon glyph, so the card asks
/// for no face beyond the fixed-pitch one the page is painted in. It is drawn in the accent
/// color, which is what makes it read as a mark rather than as punctuation.
pub(super) const BULLET: &str = "●";

/// The cell the bullet is drawn in, in header advances: the mark and the room after it.
pub(super) const BULLET_CELL_ADVANCES: i32 = 2;

/// Room between the last fact and the clock at the right edge.
pub(super) const TIME_GAP_ADVANCES: i32 = 3;

/// The room the clock is given, in body advances: enough for `1:02:03 / 1:02:03`, which is the
/// longest readout there is. What the clock takes is right-aligned inside it, so the facts are
/// laid out against a fixed edge and the two do not move as the seconds do.
pub(super) const CLOCK_ADVANCES: i32 = 17;

/// The narrowest a card is worth drawing, in body advances — the width floor the
/// card is laid out at, so a room of a few pixels answers with a card that looks
/// like one rather than a sliver. The floor is the smallest card that still holds
/// a pin's controls, and it holds more than those controls need: the four buttons
/// are carved out of the bar's own row, standing in the gap the track gives up, so
/// what the floor is really holding is the bar — a bar a handful of pixels long is
/// not the control that takes a sound to a second of it (see `control_boxes`).
pub(super) const MIN_CONTENT_ADVANCES: i32 = 34;

/// The size level the name is set in, which is the level an archive's header is set in. The
/// bullet is drawn at the same level as the text beside it, so the two share a baseline.
pub(super) const HEADER_LEVEL: u8 = 3;

/// The hairline under the name, and the room kept around it.
pub(super) const RULE_PIXELS: i32 = 1;
pub(super) const RULE_GAP_PIXELS: i32 = 5;

/// The bar at the foot of the card: a hairline track with the played part drawn over it, a
/// little thicker, so what is left and what has been heard are the same line at two weights.
pub(super) const BAR_PIXELS: i32 = 3;
pub(super) const TRACK_PIXELS: i32 = 1;

/// Room between the facts and the bar.
pub(super) const BAR_GAP_PIXELS: i32 = 10;

/// How far below the bar a press is still answered as a press on the bar: the top of the row is
/// already generous and this is the other end of the same bargain, because a hand aiming at a
/// three-pixel line from above misses the line but not the row, and one aiming from below misses
/// it as often. It stops at the card's own bottom edge (see `bar_share_at`).
pub(super) const BAR_REACH_PIXELS: i32 = 8;

/// The side of one of the card's own buttons at a display's scale, and the room between two of
/// them: a button is a square a hand can be asked to hit, and the gap is what keeps three of them
/// from reading as one wide target. Both are compact on purpose — a button is wider than the bar
/// is tall and stands centred on it, reaching into the gap above and the margin below, so the card
/// a pin shows is the size the card a hover shows is and the buttons are carved out of it rather
/// than added to it (see `bar_row`).
pub(super) const CONTROL_SIDE_PIXELS: i32 = 16;
pub(super) const CONTROL_GAP_PIXELS: i32 = 2;

/// The room around the two window buttons a pinned card carries in its top
/// margin: between each of them and the card's own border, and between the
/// two of them. A logical pixel, like every other pixel of the card — at
/// 200% it is two physical pixels, and the buttons stay the same proportion
/// of the card they are part of.
pub(super) const WINDOW_BUTTON_GAP_PIXELS: i32 = 2;

/// How far below a drawn window button a press is still answered as a press
/// on it: the gap below the drawn box, plus a pixel of the name line's own
/// invisible leading — room the line has and the name's glyphs do not, which
/// is what keeps the cushion from ever reaching the glyphs (see
/// `window_button_boxes`).
pub(super) const WINDOW_BUTTON_CUSHION_PIXELS: i32 = 3;

/// How long the block takes to cross a bar of unknown length, in seconds, and the share of the
/// bar it takes.
pub(super) const CHASE_SECONDS: f64 = 2.0;
pub(super) const CHASE_DIVISOR: i32 = 5;

/// How long a name rests at either end of its travel before it turns around.
pub(super) const NAME_HOLD: Duration = Duration::from_millis(1000);

/// The time one advance of a name's own font is worth: the speed a scroll is counted in, so
/// that a card drawn at any scale is scrolled across at the same speed to the eye.
const NAME_ADVANCE_MS: u128 = 100;

/// The options a card is built with, passed in rather than read from the configuration here so
/// a caller's intent cannot drift from what is drawn.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct AudioPreviewOptions {
    pub(crate) theme: TextTheme,
    pub(crate) font_scale_percent: u32,
}

/// How one fact is drawn: the codec's name is the one fact that is not a quantity or a plain
/// word, so it is the one drawn in the page's own foreground.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum FactKind {
    /// What the file is — `FLAC`, `MP3`, or the extension it carries where no engine named one.
    Name,
    /// A quantity: a sample rate, a bitrate.
    Number,
    /// A word: how many channels there are, as a person reads it.
    Word,
}

/// One fact, ready to be drawn: what it says and which color it is set in.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Fact {
    pub text: String,
    pub kind: FactKind,
}

/// What the card says, and how far along the sound is.
///
/// The clock's two numbers are the preview loop's business rather than the file's: a sound
/// played by Windows' engine reports its own position, and one played by FFmpeg's is timed by
/// this app's clock, so what arrives here is whatever the player had to say when the card was
/// drawn. `elapsed` is nothing where no player is running — a file nothing will play
/// is drawn without one — and `duration` is nothing where the file does not say.
#[derive(Debug, Clone, PartialEq)]
pub(crate) struct Card {
    pub name: String,
    pub facts: Vec<Fact>,
    pub duration: Option<f64>,
    pub elapsed: Option<f64>,
    /// How far the name is scrolled, in pixels, where the card has no room for the whole of
    /// it: what a repaint of a card whose name moves hands over, and nothing at all for the
    /// card a hover is measured with or the first frame of one (see [`NameScroll`]).
    pub name_offset: i32,
    /// What the controls a pinned card carries are saying: whether a player is running behind
    /// it, the level this window is playing at, and which button the pointer is over and holding.
    ///
    /// Nothing at all for the card a hover shows, which carries no controls — a hover's own
    /// window is a window nobody is in, and a click on one of those lands on the file behind it
    /// (see `control_at`). It is also what keeps the bar spanning margin to margin there: the
    /// row the four buttons and the bar share is only laid out when there are buttons in it.
    pub controls: Option<CardChrome>,
}

/// The parts of a pinned sound's card a pointer can be on: the three buttons at the left of the
/// bar, the bar itself, the volume button at its right, and the two window buttons standing in
/// the card's top margin.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum CardControl {
    /// The file before the one pinned, the walk the caption's own arrow starts.
    Previous,
    /// The button that holds a sound where it stands and sets it going again.
    Play,
    /// The file after the one pinned, the same walk the other way.
    Next,
    /// The bar itself, which is pressed to take the file to a second of it.
    Seek,
    /// The button that opens the volume popup, which is the pin's own control rather than the
    /// player's.
    Volume,
    /// The window button that shrinks the pin into the bubble it leaves, which stands where
    /// the button did (see `pinned_minimize_box`). It is a window button rather than one of
    /// the card's own controls because what it acts on is the pin rather than the sound: it
    /// is asked for as the pin's own minimize command (see `pinned_audio_control_release`).
    Minimize,
    /// The window button that ends the pin, the player and the window together, which is the
    /// pin's own close command asked for (see `pinned_audio_control_release`).
    Close,
    /// The menu the card's mark opens, in the cell the mark is drawn in: the
    /// shuffle, loop and seek switches of a pinned sound. The cell is a
    /// control only while the window buttons are up, because the stripes the
    /// cell becomes while they are are what that band is for (see `paint`),
    /// and it is hit-tested before the window buttons so that the cell wins
    /// in its own box (see `control_at`).
    Menu,
}

/// What a card's own controls are saying, and what a pointer is doing on them: whether a player
/// is running behind the card, the level this window is playing at, and which of the two
/// questions about the pointer — is it over one of them, is it holding one — is yes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub(crate) struct CardChrome {
    /// Whether a player is running, which is what the play button's glyph says.
    pub playing: bool,
    /// The level this pin is playing at, which is what the volume button's speaker says and what
    /// the panel that opens off it is drawn at.
    pub volume: u32,
    pub hovered: Option<CardControl>,
    pub pressed: Option<CardControl>,
    /// Whether the two window buttons are showing, which is what the card is painted
    /// from: a hand near the window's top border or near the buttons themselves.
    pub window_buttons: bool,
}

/// A name scrolled across a card that has no room for it: how far it has moved, which way it
/// is going, when the hold at either end runs out, and the two numbers the motion comes out
/// of — how far the name travels before the end of it is in sight, and the advance of the font
/// it is set in, which the speed is counted in.
///
/// The page never moves a name by itself: it draws the name at whatever offset it is handed,
/// whole, clipped to the card (see `Card::name_offset`). This is the state a tick advances,
/// and it lives beside the card rather than inside it because both the page and the loop need
/// to ask it things: the page is handed the offset, and the loop asks whether the name moves
/// at all, which is what its cadence is read from (see `AUDIO_NAME_REPAINT`).
#[derive(Debug, Clone, Copy)]
pub(crate) struct NameScroll {
    /// Pixels the name has been moved left from its resting place.
    offset: i32,
    /// Pixels it may move before the end of it is in sight, which is nothing at all for a
    /// name the card has room for — a name that never moves.
    pub(super) travel: i32,
    /// Pixels one advance of the name's font is: the unit the speed is counted in, so that a
    /// card at 200% scrolls the same characters a second as one at 100%.
    advance: i32,
    /// Which way it is going: one for left, minus one for back right.
    direction: i32,
    /// When the name may move again: the hold at the start, and the hold at either end.
    hold_until: Instant,
}

impl NameScroll {
    /// The scroll a card of `width` pixels draws `name` with: how far the name runs past the
    /// room the card leaves for it, at the font the card is drawn in, and nothing at all where
    /// the name fits — a card whose name is drawn once and held.
    pub(crate) fn of(name: &str, width: u32, dpi: u32, options: AudioPreviewOptions) -> Self {
        let Some((advance, room)) = name_room(width, dpi, options) else {
            return Self {
                offset: 0,
                travel: 0,
                advance: 1,
                direction: 1,
                hold_until: Instant::now(),
            };
        };

        Self {
            offset: 0,
            travel: (text_width(name, advance) - room).max(0),
            advance,
            direction: 1,
            hold_until: Instant::now() + NAME_HOLD,
        }
    }

    /// Whether the name moves at all, which is what the cadence it is repainted at is read
    /// from: a card that has to scroll cannot be redrawn at the rate the clock's own seconds
    /// are worth watching at.
    pub(crate) fn moves(&self) -> bool {
        self.travel > 0
    }

    /// How far the name is scrolled right now, in pixels.
    pub(crate) fn offset(&self) -> i32 {
        self.offset
    }

    /// Move the name on by one repaint, `cadence` after the last one: it rests at either end
    /// for `NAME_HOLD`, and between the holds it travels at about an advance a `NAME_ADVANCE_MS`
    /// — ping-ponging for as long as the card is up, so the end of a long name is shown and the
    /// beginning comes back rather than the name being read once and left where it stopped.
    ///
    /// The run's own start is held for the same time as its ends: a name that has just appeared
    /// is a name a person is reading, and a card that started moving the moment it was put up
    /// would begin at the one instant the reading starts.
    pub(crate) fn advance(&mut self, now: Instant, cadence: Duration) {
        if !self.moves() || now < self.hold_until {
            return;
        }

        // The advance is the font's own, so the speed is the same at every scale: at ten
        // advances a second, one repaint of `cadence` is that share of an advance — never less
        // than a pixel, because a step of nothing is a name repainted forever without moving.
        let step = ((self.advance as u128 * cadence.as_millis()) / NAME_ADVANCE_MS).max(1) as i32;
        let next = self.offset + step * self.direction;

        if next >= self.travel {
            self.offset = self.travel;
            self.direction = -1;
            self.hold_until = now + NAME_HOLD;
        } else if next <= 0 {
            self.offset = 0;
            self.direction = 1;
            self.hold_until = now + NAME_HOLD;
        } else {
            self.offset = next;
        }
    }
}

/// The room a card of `width` pixels leaves for its name, and the advance of the font the name
/// is set in: the two numbers a scroll is made of (see [`NameScroll`]).
///
/// It is the card's own layout, counted the way `build_page` counts it — the content box less
/// the mark's cell, both in advances of the header's font — and it is worked out here for the
/// one caller that has none of that in hand: the preview loop, which owns the scroll and needs
/// to know both where it turns around and how fast to move it.
fn name_room(width: u32, dpi: u32, options: AudioPreviewOptions) -> Option<(i32, i32)> {
    let dc = unsafe { CreateCompatibleDC(None) };
    if dc.0.is_null() {
        return None;
    }

    let result = TextMetrics::new(dc, dpi, options.font_scale_percent).map(|metrics| {
        let advance = metrics.advance[HEADER_LEVEL as usize].max(1);
        let room = width as i32 - metrics.padding * 2 - advance * BULLET_CELL_ADVANCES;

        (advance, room)
    });
    unsafe {
        let _ = DeleteDC(dc);
    }

    result
}

/// The box the card wants, bounded by what the display can give it.
pub(crate) fn measure(
    card: &Card,
    max_width: u32,
    max_height: u32,
    dpi: u32,
    options: AudioPreviewOptions,
) -> Option<(u32, u32)> {
    let theme = text_theme::loaded(options.theme)?;

    let dc = unsafe { CreateCompatibleDC(None) };
    if dc.0.is_null() {
        return None;
    }
    let result = TextMetrics::new(dc, dpi, options.font_scale_percent)
        .map(|metrics| build_page(card, theme, &metrics, max_width, max_height))
        .map(|page| (page.width, page.height));
    unsafe {
        let _ = DeleteDC(dc);
    }

    result
}

/// The card painted into the box the layout settled on.
pub(crate) fn render(
    card: &Card,
    width: u32,
    height: u32,
    dpi: u32,
    options: AudioPreviewOptions,
) -> Option<(Vec<u8>, u32, u32)> {
    let theme = text_theme::loaded(options.theme)?;

    let surface = DibSurface::create(width, height)?;
    let metrics = TextMetrics::new(surface.dc, dpi, options.font_scale_percent)?;
    let page = build_page(card, theme, &metrics, width, height);
    paint(&surface, &page, theme, metrics.scale);

    Some((surface.pixels(), surface.width, surface.height))
}

/// Where a press on a card's bar is, as the share of the file it stands for — or nothing at all
/// where the point is not on the bar.
///
/// A bar is three pixels of track with a hairline under it, which is not something a hand can be
/// asked to hit. The band a press is answered against is therefore a band rather than a line:
/// the bar's own pixels, the gap above it — the same gap every other line of the card is set out
/// with — and `BAR_REACH_PIXELS` below it, which is the other end of the same bargain. Above is
/// where the bar's own room is and it is left generous, because the hand is aiming downwards at a
/// line it can see; below there is the card's own margin, and a hand aiming upwards from the
/// bottom edge of the card lands as often beside the line as on it. The reach is clamped at the
/// card's own bottom edge, so it can never swallow a press on the margin further down — which is
/// a hand carrying the window, not a hand on the bar (see `pinned_press`).
///
/// What the share is a share of is the caller's question and not this function's: a file that
/// does not say how long it is is drawn with a block crossing the track rather than a played part
/// of it, and there is no second of such a file for a press to mean (see [`Card::duration`]).
///
/// The bar's row is asked of the same arithmetic the card is laid out with, and the width it is
/// measured across is the width the card was drawn at. `controls` says whether the card carries
/// its own buttons: the bar then fills what is left of the content box between them, and the
/// share is measured across that bar alone — a press on the button at the end of the row is a
/// press on that button, not on the first second of the file.
pub(crate) fn bar_share_at(
    x: i32,
    y: i32,
    width: u32,
    dpi: u32,
    options: AudioPreviewOptions,
    controls: bool,
) -> Option<f64> {
    let dc = unsafe { CreateCompatibleDC(None) };
    if dc.0.is_null() {
        return None;
    }

    let share = TextMetrics::new(dc, dpi, options.font_scale_percent).and_then(|metrics| {
        let row = bar_row(&metrics);
        let boxes = control_boxes(&metrics, width, controls);
        let left = boxes.map_or(metrics.padding, |boxes| boxes.bar.left);
        let right = boxes.map_or(width as i32 - metrics.padding, |boxes| boxes.bar.right);
        let across = right - left;

        // A card too narrow for the bar to be a line of anything is a card with no bar on it,
        // which is what an empty page draws.
        if across <= 0 {
            return None;
        }

        if !bar_band(row, left, right, &metrics).contains(x, y) {
            return None;
        }

        Some(((x - left) as f64 / across as f64).clamp(0.0, 1.0))
    });

    unsafe {
        let _ = DeleteDC(dc);
    }

    share
}

/// Which of a pinned card's own controls a point is on, or nothing at all where it is on none of
/// them — including the whole of a card that carries none, which is the card a hover shows.
///
/// The boxes are rebuilt from the same arithmetic `build_page` lays the row out with, which is the
/// whole of why this is a function and not a set of coordinates a caller keeps: a button hit-tested
/// by one arithmetic and drawn by another answers a press in the middle of the facts line (see
/// `BarRow`).
///
/// The two window buttons in the top margin are answered before anything else, and the order is
/// not arbitrary: a press on one is a press on a button of the window, which is a thing a hand
/// is on before it is a hand on the card, and the cushion above a drawn button reaches into the
/// name line's own leading — the one band of the card the row's controls do not share rows with,
/// which is what keeps the two questions apart even where they overlap (see `window_button_boxes`).
///
/// The volume button is answered next, and the order after that is not arbitrary either: it is
/// this app's own control rather than the player's, exactly as the transport bar's own volume
/// button is answered before that bar's play button (see `pin_chrome::transport_part_at`), and
/// the two are the same control drawn in two places. The seek is answered last because its band
/// is a band rather than a box, and a band is only worth asking about once every box in the row
/// has been.
pub(crate) fn control_at(
    x: i32,
    y: i32,
    width: u32,
    dpi: u32,
    options: AudioPreviewOptions,
    controls: bool,
    window_buttons: bool,
) -> Option<CardControl> {
    let dc = unsafe { CreateCompatibleDC(None) };
    if dc.0.is_null() {
        return None;
    }

    let found = TextMetrics::new(dc, dpi, options.font_scale_percent).and_then(|metrics| {
        let boxes = control_boxes(&metrics, width, controls)?;
        let row = bar_row(&metrics);

        // The menu the cell the card's mark is drawn in opens, asked of
        // before anything else the card carries: the cell is its own box,
        // and a hand in it is a hand on the menu. It is a control only
        // while the window buttons are up, because the stripes the cell
        // becomes while they are are what that band is for (see `paint`),
        // and it stands first in the chain so that the cell wins in its
        // own box (see the order above).
        let menu = window_buttons.then(|| {
            boxes
                .rect(CardControl::Menu)
                .filter(|rect| holds(*rect, x, y))
                .map(|_| CardControl::Menu)
        });

        // The two window buttons, which a hand on the top margin is on before
        // it is on anything the card carries lower down (see the order above).
        menu.flatten()
            .or_else(|| boxes.window_button_at(x, y))
            .or_else(|| {
                boxes
                    .rect(CardControl::Volume)
                    .filter(|rect| holds(*rect, x, y))
                    .map(|_| CardControl::Volume)
            })
            .or_else(|| {
                boxes
                    .rect(CardControl::Previous)
                    .filter(|rect| holds(*rect, x, y))
                    .map(|_| CardControl::Previous)
            })
            .or_else(|| {
                boxes
                    .rect(CardControl::Play)
                    .filter(|rect| holds(*rect, x, y))
                    .map(|_| CardControl::Play)
            })
            .or_else(|| {
                boxes
                    .rect(CardControl::Next)
                    .filter(|rect| holds(*rect, x, y))
                    .map(|_| CardControl::Next)
            })
            .or_else(|| {
                bar_band(row, boxes.bar.left, boxes.bar.right, &metrics)
                    .contains(x, y)
                    .then_some(CardControl::Seek)
            })
    });

    unsafe {
        let _ = DeleteDC(dc);
    }

    found
}

/// The box one of a pinned card's own controls is drawn in, in the card's own coordinates — the
/// box a press on it is answered against, read back out of the same arithmetic (see
/// [`control_at`]).
///
/// For the two window buttons in the top margin this is the box a press is answered against
/// rather than the box the glyph is drawn in: the drawn box plus the cushion above it (see
/// [`WindowButtonBox`]), which is the box the bubble a minimize leaves stands on (see
/// `pinned_minimize_box`).
///
/// A card that carries no controls has no box for any of them, which is the answer a hover's card
/// gives for every one: a hover's window is a window nobody is in.
pub(crate) fn control_box(
    control: CardControl,
    width: u32,
    dpi: u32,
    options: AudioPreviewOptions,
    controls: bool,
) -> Option<RECT> {
    let dc = unsafe { CreateCompatibleDC(None) };
    if dc.0.is_null() {
        return None;
    }

    let box_of = TextMetrics::new(dc, dpi, options.font_scale_percent)
        .and_then(|metrics| control_boxes(&metrics, width, controls))
        .and_then(|boxes| boxes.rect(control));

    unsafe {
        let _ = DeleteDC(dc);
    }

    box_of
}

/// The facts line of a file: what it holds, in the order a person reads it — what the format
/// is, how fast it was sampled, how many channels it has, how much room a second of it takes.
///
/// Every part is left out rather than guessed at, and a file no engine named a codec for is
/// named by the extension it carries, which is what a card for it can honestly say.
pub(crate) fn facts_of(track: &Track, path: &Path) -> Vec<Fact> {
    let mut facts = vec![Fact {
        text: track
            .codec
            .clone()
            .unwrap_or_else(|| extension_of(path).to_uppercase()),
        kind: FactKind::Name,
    }];

    if let Some(rate) = track.rate {
        facts.push(Fact {
            text: rate_label(rate),
            kind: FactKind::Number,
        });
    }
    if let Some(channels) = track.channels {
        facts.push(Fact {
            text: channel_label(channels),
            kind: FactKind::Word,
        });
    }
    if let Some(bitrate) = track.bitrate {
        facts.push(Fact {
            text: bitrate_label(bitrate),
            kind: FactKind::Number,
        });
    }

    facts
}

/// The name a card is headed with: the file's own name, which is what a person looking at the
/// card is identifying — the folder it is in is already said by the hover that put the card up,
/// and a card is not wide enough for a whole path.
///
/// It is the card's business rather than a caller's because two of them need it and they have
/// to agree: the card is drawn with it, and the scroll a card too narrow for it needs is
/// measured from it (see [`NameScroll::of`]).
pub(crate) fn name_of(path: &Path) -> String {
    path.file_name()
        .map(|name| name.to_string_lossy().to_string())
        .unwrap_or_default()
}

fn extension_of(path: &Path) -> String {
    path.extension()
        .map(|extension| extension.to_string_lossy().to_string())
        .unwrap_or_else(|| "audio".to_string())
}

/// A sample rate as a person reads it: kHz to the decimal the fraction asks for, and the whole
/// number where it asks for none — `44.1 kHz`, `48 kHz`, `22.05 kHz`.
pub(super) fn rate_label(rate: u32) -> String {
    let khz = rate as f64 / 1000.0;
    let rounded = (khz * 100.0).round() / 100.0;

    if rounded.fract() == 0.0 {
        format!("{rounded:.0} kHz")
    } else if (rounded * 10.0).fract() == 0.0 {
        format!("{rounded:.1} kHz")
    } else {
        format!("{rounded:.2} kHz")
    }
}

/// How many channels, as a word where a word is what people say: `mono`, `stereo`, and the
/// count for everything else.
pub(super) fn channel_label(channels: u16) -> String {
    match channels {
        0 => String::new(),
        1 => "mono".to_string(),
        2 => "stereo".to_string(),
        6 => "5.1".to_string(),
        8 => "7.1".to_string(),
        channels => format!("{channels} ch"),
    }
}

/// A bitrate as a person reads it, in the unit the number belongs in: kilobits for anything
/// that came off a codec, bits for a stream that is only a few hundred a second.
pub(super) fn bitrate_label(bitrate: u32) -> String {
    if bitrate >= 1000 {
        format!("{} kbps", (bitrate as f64 / 1000.0).round() as u64)
    } else {
        format!("{bitrate} bps")
    }
}

/// A position as a clock: `9:22`, and `1:02:03` where the file runs past an hour.
pub(super) fn clock(seconds: f64) -> String {
    let total = seconds.max(0.0) as u64;
    let (hours, minutes, seconds) = (total / 3600, (total % 3600) / 60, total % 60);

    if hours > 0 {
        format!("{hours}:{minutes:02}:{seconds:02}")
    } else {
        format!("{minutes}:{seconds:02}")
    }
}

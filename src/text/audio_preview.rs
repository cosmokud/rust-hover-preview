//! The card a sound file is previewed as.
//!
//! A sound has no picture, so what a hover shows is what a person wants to know about one: a
//! mark, the file's name, what the file holds, and how far into it the sound is while it
//! plays. It is painted with the same GDI layer a text preview uses and colored by the same
//! theme, which is what keeps the three painted previews looking like one app — and it is
//! made of text and rules and nothing else, because that is what the layer draws.
//!
//! Two calls share the work, as they do for text and for an archive: `measure` answers how big
//! the card wants to be before the window is placed, and `render` fills the box the layout
//! settled on. Both build the same page and neither keeps it — what is cached is the track the
//! card is drawn from (`audio_track`), so a repaint of the clock is a layout rather than a
//! probe, and a second hover of the file is a layout rather than a probe as well.
//!
//! Two things change while the card is on screen, and both are the preview loop's: the clock,
//! with the bar under it, which is drawn from a player that is running, and a name the card
//! has no room for, which is scrolled across the card sideways rather than being cut short
//! (see [`NameScroll`] and `Card::name_offset`). A painted preview of this app's is otherwise
//! drawn once and held, so the preview loop is what asks for the card again while a sound is
//! playing (see `repaint_audio_card`).
//!
//! The card carries its own controls when it is the card a pinned window is showing: the three
//! buttons at the left of its bar row, the bar itself, and the volume button at the bar's right.
//! Every question about where one of them is is answered against the card's own layout (see
//! [`control_at`], [`control_box`] and [`bar_share_at`]) rather than against anything kept beside
//! it, because a button laid out by one arithmetic and hit-tested by another answers a press in
//! the middle of the facts line. Only a pinned window asks: a hover's own window is a window
//! nobody is in, and a click on one of those lands on the file behind it, so the card a hover
//! shows is the card it has always been — no buttons, and a bar from margin to margin (see
//! [`Card::controls`]). The row the buttons stand in is the bar's own line either way, so the
//! card a pin shows is the size the card a hover shows is: what the buttons cost is the width of
//! the bar, not the height of the card (see `bar_row`).

use crate::config::config::TextTheme;
use crate::readers::audio_track::Track;
use crate::text::pin_chrome::{self, ControlGlyph};
use crate::text::text_paint::{
    blend, fill_rect, plain_style, readable, rgb, scaled, DibSurface, RunPainter, TextMetrics,
    TextStyle, BODY_LEVEL,
};
use crate::text::text_theme::{self, LoadedTheme};
use std::path::Path;
use std::time::{Duration, Instant};
use windows::Win32::Foundation::RECT;
use windows::Win32::Graphics::Gdi::{CreateCompatibleDC, DeleteDC};

/// The mark the card opens with: a round bullet rather than an icon glyph, so the card asks
/// for no face beyond the fixed-pitch one the page is painted in. It is drawn in the accent
/// color, which is what makes it read as a mark rather than as punctuation.
const BULLET: &str = "●";

/// The cell the bullet is drawn in, in header advances: the mark and the room after it.
const BULLET_CELL_ADVANCES: i32 = 2;

/// Room between the last fact and the clock at the right edge.
const TIME_GAP_ADVANCES: i32 = 3;

/// The room the clock is given, in body advances: enough for `1:02:03 / 1:02:03`, which is the
/// longest readout there is. What the clock takes is right-aligned inside it, so the facts are
/// laid out against a fixed edge and the two do not move as the seconds do.
const CLOCK_ADVANCES: i32 = 17;

/// The narrowest a card is worth drawing, in body advances: the bar is what asks for it, so a
/// file with a short name still gets a card that looks like one.
const MIN_CONTENT_ADVANCES: i32 = 34;

/// The size level the name is set in, which is the level an archive's header is set in. The
/// bullet is drawn at the same level as the text beside it, so the two share a baseline.
const HEADER_LEVEL: u8 = 3;

/// The hairline under the name, and the room kept around it.
const RULE_PIXELS: i32 = 1;
const RULE_GAP_PIXELS: i32 = 5;

/// The bar at the foot of the card: a hairline track with the played part drawn over it, a
/// little thicker, so what is left and what has been heard are the same line at two weights.
const BAR_PIXELS: i32 = 3;
const TRACK_PIXELS: i32 = 1;

/// Room between the facts and the bar.
const BAR_GAP_PIXELS: i32 = 10;

/// How far below the bar a press is still answered as a press on the bar: the top of the row is
/// already generous and this is the other end of the same bargain, because a hand aiming at a
/// three-pixel line from above misses the line but not the row, and one aiming from below misses
/// it as often. It stops at the card's own bottom edge (see `bar_share_at`).
const BAR_REACH_PIXELS: i32 = 8;

/// The side of one of the card's own buttons at a display's scale, and the room between two of
/// them: a button is a square a hand can be asked to hit, and the gap is what keeps three of them
/// from reading as one wide target. Both are compact on purpose — a button is wider than the bar
/// is tall and stands centred on it, reaching into the gap above and the margin below, so the card
/// a pin shows is the size the card a hover shows is and the buttons are carved out of it rather
/// than added to it (see `bar_row`).
const CONTROL_SIDE_PIXELS: i32 = 16;
const CONTROL_GAP_PIXELS: i32 = 2;

/// How long the block takes to cross a bar of unknown length, in seconds, and the share of the
/// bar it takes.
const CHASE_SECONDS: f64 = 2.0;
const CHASE_DIVISOR: i32 = 5;

/// How long a name rests at either end of its travel before it turns around.
const NAME_HOLD: Duration = Duration::from_millis(1000);

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
/// bar, the bar itself, and the volume button at its right.
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
    travel: i32,
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
/// [`BarRow`]).
///
/// The volume button is answered first, and the order is not arbitrary: it is this app's own
/// control rather than the player's, exactly as the transport bar's own volume button is answered
/// before that bar's play button (see `pin_chrome::transport_part_at`), and the two are the same
/// control drawn in two places. The seek is answered last because its band is a band rather than a
/// box, and a band is only worth asking about once every box in the row has been.
pub(crate) fn control_at(
    x: i32,
    y: i32,
    width: u32,
    dpi: u32,
    options: AudioPreviewOptions,
    controls: bool,
) -> Option<CardControl> {
    let dc = unsafe { CreateCompatibleDC(None) };
    if dc.0.is_null() {
        return None;
    }

    let found = TextMetrics::new(dc, dpi, options.font_scale_percent).and_then(|metrics| {
        let boxes = control_boxes(&metrics, width, controls)?;
        let row = bar_row(&metrics);

        boxes
            .rect(CardControl::Volume)
            .filter(|rect| holds(*rect, x, y))
            .map(|_| CardControl::Volume)
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
fn rate_label(rate: u32) -> String {
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
fn channel_label(channels: u16) -> String {
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
fn bitrate_label(bitrate: u32) -> String {
    if bitrate >= 1000 {
        format!("{} kbps", (bitrate as f64 / 1000.0).round() as u64)
    } else {
        format!("{bitrate} bps")
    }
}

/// A position as a clock: `9:22`, and `1:02:03` where the file runs past an hour.
fn clock(seconds: f64) -> String {
    let total = seconds.max(0.0) as u64;
    let (hours, minutes, seconds) = (total / 3600, (total % 3600) / 60, total % 60);

    if hours > 0 {
        format!("{hours}:{minutes:02}:{seconds:02}")
    } else {
        format!("{minutes}:{seconds:02}")
    }
}

// ----------------------------------------------------------------------- page

/// One run of the card: a piece of text drawn from `origin` into the box `x`–`x + width`, in
/// one style.
struct PageRun {
    text: String,
    /// The box the run is laid out in and drawn in: how the runs beside it are placed, and
    /// what the run is clipped to.
    x: i32,
    width: i32,
    /// Where the text itself starts. It is the box's own left edge for every run of the card
    /// but one — a name the card has no room for, which is drawn whole and scrolled under the
    /// box, so what moves is the origin and what stays is the box that clips it (see
    /// `Card::name_offset`).
    origin: i32,
    style: TextStyle,
}

/// The bar and what is drawn over it.
#[derive(Debug)]
struct Bar {
    left: i32,
    width: i32,
    top: i32,
    height: i32,
    track_height: i32,
    /// Where the played part starts and how wide it is. A bar with no length to measure
    /// against — a file that does not say how long it is — is a block crossing the track
    /// instead, which is the difference between a sound and a stream.
    fill_start: i32,
    fill_width: i32,
}

/// The row the bar is drawn in: where it begins beneath the facts, how tall it is, and how thick
/// the track under the played part is.
///
/// These are the numbers the bar is drawn from *and* the numbers a press on it is answered
/// against, which is the whole of why they are one set rather than two: a bar laid out by one
/// arithmetic and hit-tested by another is a bar whose middle few pixels answer for a line
/// somewhere else on the card (see [`bar_share_at`]).
struct BarRow {
    top: i32,
    height: i32,
    track_height: i32,
}

/// The bar's row: the bar's own height, whether or not buttons are standing in it. A button is
/// taller than that line, and stands centred on it rather than growing the row — which is what
/// makes a card a pin shows the same size as the card a hover shows, with the four buttons carved
/// out of the bar's line instead of the card growing to hold them (see `control_boxes`).
fn bar_row(metrics: &TextMetrics) -> BarRow {
    BarRow {
        // The same walk down the page `build_page` makes, from the top margin to the bar: the
        // name, the rule under it, the facts, and the gap the bar is set out after.
        top: metrics.padding
            + metrics.line_height[HEADER_LEVEL as usize]
            + scaled(RULE_GAP_PIXELS, metrics.scale)
            + scaled(RULE_PIXELS, metrics.scale)
            + scaled(RULE_GAP_PIXELS, metrics.scale)
            + metrics.line_height[BODY_LEVEL as usize]
            + scaled(BAR_GAP_PIXELS, metrics.scale),
        height: scaled(BAR_PIXELS, metrics.scale),
        track_height: scaled(TRACK_PIXELS, metrics.scale),
    }
}

/// The band around the bar a press is answered against, in the card's own coordinates: the
/// bar's row, the gap above it and `BAR_REACH_PIXELS` below it — the last clamped at the card's
/// own bottom edge, so the band cannot reach into the margin that carries the window.
///
/// This is one band for every question that asks where the bar is (see [`bar_share_at`] and
/// [`control_at`]), because a band that is generous for one question and not for the other is a
/// bar whose own middle answers for a press and whose ends do not.
struct BarBand {
    left: i32,
    right: i32,
    top: i32,
    bottom: i32,
}

impl BarBand {
    fn contains(&self, x: i32, y: i32) -> bool {
        x >= self.left && x < self.right && y >= self.top && y < self.bottom
    }
}

fn bar_band(row: BarRow, left: i32, right: i32, metrics: &TextMetrics) -> BarBand {
    // The card's own bottom edge is the row's bottom plus the page's own margin, which is where
    // the card stops rather than a number of its own: the clamp is what keeps the reach below the
    // bar inside the card.
    let card_bottom = row.top + row.height + metrics.padding;

    BarBand {
        left,
        right,
        top: row.top - scaled(BAR_GAP_PIXELS, metrics.scale),
        bottom: (row.top + row.height + scaled(BAR_REACH_PIXELS, metrics.scale)).min(card_bottom),
    }
}

/// The boxes of the controls a pinned card carries, laid out across the card's own content box:
/// the three buttons at its left, the bar filling what is left of the row, and the volume button
/// against its right edge.
///
/// They are one walk rather than five numbers each because the card is drawn from these boxes and
/// pressed against them: a button laid out here and hit-tested elsewhere answers a press in the
/// middle of the facts line (see [`BarRow`]). Copied rather than borrowed so that one walk can
/// answer for both the bar's own edges and the whole row (see `bar_share_at`).
#[derive(Clone, Copy)]
struct CardBoxes {
    previous: RECT,
    play: RECT,
    next: RECT,
    volume: RECT,
    bar: RECT,
}

impl CardBoxes {
    /// The box one of the row's controls is drawn in, which is the box a press on it is answered
    /// against — except the bar's, whose is its own box inside the row and whose band is wider.
    fn rect(&self, control: CardControl) -> Option<RECT> {
        Some(match control {
            CardControl::Previous => self.previous,
            CardControl::Play => self.play,
            CardControl::Next => self.next,
            CardControl::Seek => self.bar,
            CardControl::Volume => self.volume,
        })
    }
}

/// The row's boxes, laid out for a card of `width` pixels at a display's scale — or nothing at all
/// where the card carries no controls, which is the card a hover shows and the card every caller
/// that is only asking *where* something is uses.
fn control_boxes(metrics: &TextMetrics, width: u32, controls: bool) -> Option<CardBoxes> {
    if !controls {
        return None;
    }

    let side = scaled(CONTROL_SIDE_PIXELS, metrics.scale);
    let gap = scaled(CONTROL_GAP_PIXELS, metrics.scale);
    let row = bar_row(metrics);

    // The buttons stand centred on the bar's own line, so they reach above it into the gap and
    // below it into the margin: which is what keeps the card the size a hover's card is rather
    // than a taller one. A button is asked for at this centre and nothing else, so a press is
    // answered by the box it is drawn in (see `control_at`).
    let centre = row.top + row.height / 2;
    let top = centre - side / 2;
    let button = |left: i32| RECT {
        left,
        top,
        right: left + side,
        bottom: top + side,
    };

    let left = metrics.padding;
    let right = width as i32 - metrics.padding;

    let previous = button(left);
    let play = button(previous.right + gap);
    let next = button(play.right + gap);
    let volume = button(right - side);
    let bar_left = next.right + gap;

    Some(CardBoxes {
        previous,
        play,
        next,
        volume,
        // The bar is what is left of the content box between the two ends of the row, and it is
        // the bar's own line rather than the buttons' height: a seek is answered against the
        // track a person can see, which is where it is drawn (see `bar_band`). A card too narrow
        // to hold both is a card whose bar has no width, and `bar_share_at` answers nothing for
        // it rather than a share of a line it does not have.
        bar: RECT {
            left: bar_left,
            top: row.top,
            right: (volume.left - gap).max(bar_left),
            bottom: row.top + row.height,
        },
    })
}

/// Whether a point is inside a box, in the card's own coordinates. Half-open at the right and the
/// bottom, so that two boxes which touch are two boxes and not one row of pixels answered twice.
fn holds(rect: RECT, x: i32, y: i32) -> bool {
    x >= rect.left && x < rect.right && y >= rect.top && y < rect.bottom
}

/// A painted card: what to draw and how big it came out.
struct Page {
    header: Vec<PageRun>,
    header_top: i32,
    header_height: i32,
    rule_top: i32,
    rule_height: i32,
    facts: Vec<PageRun>,
    facts_top: i32,
    facts_height: i32,
    bar: Bar,
    /// The boxes the row's controls are drawn in and answered against, and nothing at all for a
    /// card that carries none.
    boxes: Option<CardBoxes>,
    /// What those controls are saying, which is the wash under the pointer and the glyph on each
    /// of them: kept beside `boxes` rather than inside it because the hit test rebuilds the boxes
    /// and has no business with any of this (see `control_at`).
    chrome: Option<CardChrome>,
    width: u32,
    height: u32,
    padding: i32,
}

fn build_page(
    card: &Card,
    theme: &LoadedTheme,
    metrics: &TextMetrics,
    box_width: u32,
    box_height: u32,
) -> Page {
    let header_advance = metrics.advance[HEADER_LEVEL as usize].max(1);
    let body_advance = metrics.advance[BODY_LEVEL as usize].max(1);
    let header_height = metrics.line_height[HEADER_LEVEL as usize];
    let body_height = metrics.line_height[BODY_LEVEL as usize];
    let padding = metrics.padding;

    let rule_gap = scaled(RULE_GAP_PIXELS, metrics.scale);
    let rule_height = scaled(RULE_PIXELS, metrics.scale);
    let bar_row = bar_row(metrics);

    let page_color = rgb(theme.background());
    let foreground = rgb(theme.foreground());
    let muted = blend(foreground, page_color, 0.45);
    let number = readable(
        rgb(theme.style_for_scopes(&["constant.numeric"]).foreground),
        page_color,
    );
    let accent = readable(
        rgb(theme.style_for_scopes(&["support.function"]).foreground),
        page_color,
    );

    // The width the card would like: the facts line with the clock beside it, or the narrowest
    // bar there is, whichever asks for more — clamped to the box the way every preview's size
    // is. The name is not in that count, and that is the one rule this line is: a name too long
    // for the card is not a reason to ask the display for a card a hundred advances wide, so
    // what gives way is the name, which is drawn whole and scrolled across the card (see
    // `Card::name_offset`). A name the card *does* have room for is inside these two anyway,
    // since the width is the line beneath it.
    let bullet_cell = header_advance * BULLET_CELL_ADVANCES;
    let facts_line = facts_width(&card.facts, body_advance)
        + body_advance * TIME_GAP_ADVANCES
        + body_advance * CLOCK_ADVANCES;
    // A card whose facts are narrow still needs room for the four buttons the bar is carved out
    // beside: the bar is what takes a sound to a second of it, and a bar a hundred pixels long is
    // not the same control as one that fills the card. The bar gives way for them rather than the
    // card giving way — a pin's card is the width a hover's is, and what the buttons cost is the
    // track's width (see `control_boxes`).
    let content = facts_line.max(body_advance * MIN_CONTENT_ADVANCES);
    let width = (content + padding * 2).clamp(1, box_width.max(1) as i32) as u32;

    // A card with no room for a name and a line under it is a card that cannot be drawn, which
    // is the answer an archive page gives in the same place.
    if (width as i32) < padding * 2 + body_advance
        || (box_height as i32) < padding * 2 + header_height
    {
        return empty_page(width, padding);
    }

    let content_left = padding;
    let content_right = width as i32 - padding;

    // The name is drawn whole and clipped to the box the mark leaves it, and one the box has no
    // room for is scrolled under that box rather than cut short: what moves is where the run
    // starts, and the box that clips it never does. A run's own origin is its box's left edge
    // for every other run of the card, so this is the one place the two are told apart (see
    // `PageRun::origin` and `Card::name_offset`).
    let name_left = content_left + bullet_cell;
    let header = vec![
        PageRun {
            text: BULLET.to_string(),
            x: content_left,
            width: bullet_cell,
            origin: content_left,
            style: {
                let mut style = plain_style(HEADER_LEVEL);
                style.foreground = accent;
                style
            },
        },
        PageRun {
            text: card.name.clone(),
            x: name_left,
            width: content_right - name_left,
            origin: name_left - card.name_offset.max(0),
            style: {
                let mut style = plain_style(HEADER_LEVEL);
                style.foreground = foreground;
                style.bold = true;
                style
            },
        },
    ];

    let mut top = padding;
    let header_top = top;
    top += header_height + rule_gap;
    let rule_top = top;
    top += rule_height + rule_gap;
    let facts_top = top;

    let bar_width = content_right - content_left;
    let mut facts = fact_runs(
        &card.facts,
        content_left,
        bar_width - body_advance * (TIME_GAP_ADVANCES + CLOCK_ADVANCES),
        body_advance,
        foreground,
        muted,
        number,
    );
    facts.extend(clock_runs(card, content_right, body_advance, muted, accent));

    let boxes = control_boxes(metrics, width, card.controls.is_some());
    let bar_left = boxes.map_or(content_left, |boxes| boxes.bar.left);
    let bar_span = boxes.map_or(bar_width, |boxes| boxes.bar.right - boxes.bar.left);
    let (fill_start, fill_width) = fill_span(card, bar_span, metrics.scale);

    Page {
        header,
        header_top,
        header_height,
        rule_top,
        rule_height,
        facts,
        facts_top,
        facts_height: body_height,
        bar: Bar {
            left: bar_left,
            width: bar_span,
            top: bar_row.top,
            height: bar_row.height,
            track_height: bar_row.track_height,
            fill_start,
            fill_width,
        },
        boxes,
        chrome: card.controls,
        width,
        height: box_height
            .min((bar_row.top + bar_row.height + padding) as u32)
            .max(1),
        padding,
    }
}

fn empty_page(width: u32, padding: i32) -> Page {
    Page {
        header: Vec::new(),
        header_top: 0,
        header_height: 0,
        rule_top: 0,
        rule_height: 0,
        facts: Vec::new(),
        facts_top: 0,
        facts_height: 0,
        bar: Bar {
            left: 0,
            width: 0,
            top: 0,
            height: 0,
            track_height: 0,
            fill_start: 0,
            fill_width: 0,
        },
        boxes: None,
        chrome: None,
        width,
        height: 1,
        padding,
    }
}

/// Where the played part of the bar starts and how wide it is: the played share of the whole
/// where the file says how long it is, and a block crossing the track where it does not.
fn fill_span(card: &Card, bar_width: i32, scale: f32) -> (i32, i32) {
    let Some(elapsed) = card.elapsed else {
        return (0, 0);
    };

    match card.duration {
        Some(duration) if duration > 0.0 => {
            let filled =
                ((bar_width as f64 * elapsed / duration).round() as i32).clamp(0, bar_width);
            (0, filled)
        }
        // A sound still playing and no length to measure it against: what is shown instead is
        // that it is playing at all — a block that crosses the track and starts over, so a
        // hover on a stream is never a bar that has stood still. It travels from one edge to
        // the other rather than from past one edge to past the other, so that there is no
        // instant of the cycle at which the bar is empty and the sound looks stopped.
        _ => {
            let block = (bar_width / CHASE_DIVISOR)
                .max(scaled(6, scale))
                .min(bar_width);
            let phase = (elapsed % CHASE_SECONDS) / CHASE_SECONDS;
            let travel = (bar_width - block).max(0);
            let lead = (phase * f64::from(travel)) as i32;
            (lead, block.min(bar_width - lead).max(0))
        }
    }
}

/// The readout at the right edge of the facts line: what has been played in the accent, and
/// the whole it is measured against in the page's own muted gray. A card with no player behind
/// it has only the whole to show, and one whose file does not say how long it is has only the
/// number that moves.
fn clock_runs(
    card: &Card,
    right: i32,
    advance: i32,
    muted: [u8; 3],
    accent: [u8; 3],
) -> Vec<PageRun> {
    let mut parts: Vec<(String, [u8; 3])> = Vec::new();
    match (card.elapsed, card.duration) {
        (Some(elapsed), Some(duration)) if duration > 0.0 => {
            parts.push((clock(elapsed), accent));
            parts.push((" / ".to_string(), muted));
            parts.push((clock(duration), muted));
        }
        (Some(elapsed), _) => parts.push((clock(elapsed), accent)),
        (None, Some(duration)) if duration > 0.0 => parts.push((clock(duration), muted)),
        _ => {}
    }

    let total: i32 = parts
        .iter()
        .map(|(text, _)| text_width(text, advance))
        .sum();
    let mut x = right - total;

    parts
        .into_iter()
        .map(|(text, color)| {
            let width = text_width(&text, advance);
            let run = PageRun {
                text,
                x,
                width,
                origin: x,
                style: {
                    let mut style = plain_style(BODY_LEVEL);
                    style.foreground = color;
                    style
                },
            };
            x += width;
            run
        })
        .collect()
}

/// The facts as runs, each in its own color, with the separators between them — cut to the
/// room the clock leaves, so what gives way on a narrow card is the last fact rather than the
/// clock or the names of the facts before it.
#[allow(clippy::too_many_arguments)]
fn fact_runs(
    facts: &[Fact],
    left: i32,
    room: i32,
    advance: i32,
    foreground: [u8; 3],
    muted: [u8; 3],
    number: [u8; 3],
) -> Vec<PageRun> {
    let mut runs = Vec::new();
    let mut x = left;
    let right = left + room.max(0);

    for (index, fact) in facts.iter().enumerate() {
        if x >= right {
            break;
        }

        if index > 0 {
            let separator = " · ";
            let separator_width = text_width(separator, advance);
            if x + separator_width > right {
                break;
            }
            runs.push(PageRun {
                text: separator.to_string(),
                x,
                width: separator_width,
                origin: x,
                style: {
                    let mut style = plain_style(BODY_LEVEL);
                    style.foreground = muted;
                    style
                },
            });
            x += separator_width;
        }

        let text = cut_to_width(&fact.text, right - x, advance);
        let width = text_width(&text, advance);
        runs.push(PageRun {
            text,
            x,
            width,
            origin: x,
            style: {
                let mut style = plain_style(BODY_LEVEL);
                style.foreground = match fact.kind {
                    FactKind::Name => foreground,
                    FactKind::Number => number,
                    FactKind::Word => muted,
                };
                style
            },
        });
        x += width;
        if width == 0 {
            break;
        }
    }

    runs
}

// -------------------------------------------------------------------- drawing

fn paint(surface: &DibSurface, page: &Page, theme: &LoadedTheme, scale: f32) {
    let page_color = rgb(theme.background());
    let foreground = rgb(theme.foreground());
    let rule_color = blend(page_color, foreground, 0.18);
    let accent = readable(
        rgb(theme.style_for_scopes(&["support.function"]).foreground),
        page_color,
    );

    fill_rect(
        surface,
        RECT {
            left: 0,
            top: 0,
            right: surface.width as i32,
            bottom: surface.height as i32,
        },
        page_color,
    );

    let mut painter = RunPainter::new(surface, scale);
    for (runs, top, height) in [
        (&page.header, page.header_top, page.header_height),
        (&page.facts, page.facts_top, page.facts_height),
    ] {
        for run in runs {
            painter.draw(
                &run.text,
                run.origin,
                RECT {
                    left: run.x,
                    top,
                    right: run.x + run.width,
                    bottom: top + height,
                },
                &run.style,
                run.style.foreground,
                page_color,
            );
        }
    }

    if page.rule_height > 0 {
        fill_rect(
            surface,
            RECT {
                left: page.padding,
                top: page.rule_top,
                right: surface.width as i32 - page.padding,
                bottom: page.rule_top + page.rule_height,
            },
            rule_color,
        );
    }

    if page.bar.width > 0 && page.bar.height > 0 {
        // The track is centred in the row rather than laid out from its top, so that a row a
        // button stands across puts the line where the button is rather than a third of the way
        // down it (see `BarRow`).
        let centre = page.bar.top + page.bar.height / 2;
        let track_top = centre - page.bar.track_height / 2;
        fill_rect(
            surface,
            RECT {
                left: page.bar.left,
                top: track_top,
                right: page.bar.left + page.bar.width,
                bottom: track_top + page.bar.track_height,
            },
            rule_color,
        );

        if page.bar.fill_width > 0 {
            // The played part is `BAR_PIXELS` thick about the same centre, which is what makes it
            // read as the same line at a heavier weight: it is the bar's own height
            // whichever controls stand in the row (see `BarRow`).
            let played = scaled(BAR_PIXELS, scale);
            let played_top = centre - played / 2;
            fill_rect(
                surface,
                RECT {
                    left: page.bar.left + page.bar.fill_start,
                    top: played_top,
                    right: page.bar.left + page.bar.fill_start + page.bar.fill_width,
                    bottom: played_top + played,
                },
                accent,
            );
        }
    }

    if let Some(boxes) = page.boxes.as_ref() {
        let Some(chrome) = page.chrome.as_ref() else {
            return;
        };

        // The wash first, for the whole of the row, and then the glyphs on top of it: a button
        // lit under a pointer is the only place a pointer is visible on a card, and the card
        // itself is what shows it — the pin's own chrome is drawn over the media band, and the
        // card *is* the media band (see `PinnedPreview::audio_hovered`).
        for control in [
            CardControl::Previous,
            CardControl::Play,
            CardControl::Next,
            CardControl::Volume,
        ] {
            let Some(rect) = boxes.rect(control) else {
                continue;
            };

            let wash = if chrome.pressed == Some(control) {
                Some(0.18)
            } else if chrome.hovered == Some(control) {
                Some(0.10)
            } else {
                None
            };
            if let Some(amount) = wash {
                fill_rect(surface, rect, blend(page_color, foreground, amount));
            }
        }

        for (control, glyph) in [
            (CardControl::Previous, ControlGlyph::Previous),
            (
                CardControl::Play,
                if chrome.playing {
                    ControlGlyph::Pause
                } else {
                    ControlGlyph::Play
                },
            ),
            (CardControl::Next, ControlGlyph::Next),
            (CardControl::Volume, ControlGlyph::Volume(chrome.volume)),
        ] {
            if let Some(rect) = boxes.rect(control) {
                pin_chrome::paint_card_control(surface, rect, glyph, foreground, scale);
            }
        }
    }
}

// ----------------------------------------------------------------------- text

/// The width of a run of text, counted the way the rest of the page is: the face is fixed
/// pitch, so a count of characters is the whole measurement.
fn text_width(text: &str, advance: i32) -> i32 {
    text.chars().count() as i32 * advance.max(1)
}

/// How wide the facts are with the separators between them.
fn facts_width(facts: &[Fact], advance: i32) -> i32 {
    let separator = text_width(" · ", advance);
    let parts: i32 = facts
        .iter()
        .map(|fact| text_width(&fact.text, advance))
        .sum();

    parts + separator * facts.len().saturating_sub(1) as i32
}

/// A name cut to the room it has, with an ellipsis where it was cut. A name that does not fit
/// at all is dropped rather than drawn over the card.
fn cut_to_width(text: &str, room: i32, advance: i32) -> String {
    let advance = advance.max(1);
    let room = (room / advance).max(0) as usize;
    let characters = text.chars().count();
    if characters <= room {
        return text.to_string();
    }
    if room == 0 {
        return String::new();
    }

    let mut cut: String = text.chars().take(room - 1).collect();
    cut.push('…');

    cut
}

#[cfg(test)]
mod tests;

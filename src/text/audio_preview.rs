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

use crate::config::config::TextTheme;
use crate::readers::audio_track::Track;
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
/// drawn. `elapsed` is nothing where no player is running — a card at `Volume → Audio` 0% is
/// drawn without one — and `duration` is nothing where the file does not say.
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
    let bar_height = scaled(BAR_PIXELS, metrics.scale);
    let track_height = scaled(TRACK_PIXELS, metrics.scale);
    let bar_gap = scaled(BAR_GAP_PIXELS, metrics.scale);

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
    top += body_height + bar_gap;
    let bar_top = top;

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
    facts.extend(clock_runs(
        card,
        content_right,
        body_advance,
        muted,
        accent,
    ));

    let (fill_start, fill_width) = fill_span(card, bar_width, metrics.scale);

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
            left: content_left,
            width: bar_width,
            top: bar_top,
            height: bar_height,
            track_height,
            fill_start,
            fill_width,
        },
        width,
        height: box_height.min((bar_top + bar_height + padding) as u32).max(1),
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
            let filled = ((bar_width as f64 * elapsed / duration).round() as i32).clamp(0, bar_width);
            (0, filled)
        }
        // A sound still playing and no length to measure it against: what is shown instead is
        // that it is playing at all — a block that crosses the track and starts over, so a
        // hover on a stream is never a bar that has stood still. It travels from one edge to
        // the other rather than from past one edge to past the other, so that there is no
        // instant of the cycle at which the bar is empty and the sound looks stopped.
        _ => {
            let block = (bar_width / CHASE_DIVISOR).max(scaled(6, scale)).min(bar_width);
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
    let rule_color = blend(page_color, rgb(theme.foreground()), 0.18);
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
            fill_rect(
                surface,
                RECT {
                    left: page.bar.left + page.bar.fill_start,
                    top: page.bar.top,
                    right: page.bar.left + page.bar.fill_start + page.bar.fill_width,
                    bottom: page.bar.top + page.bar.height,
                },
                accent,
            );
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
mod tests {
    use super::*;
    use crate::readers::audio_track::Player;

    fn card() -> Card {
        Card {
            name: "Kind of Blue - So What.flac".to_string(),
            facts: vec![
                Fact {
                    text: "FLAC".to_string(),
                    kind: FactKind::Name,
                },
                Fact {
                    text: "44.1 kHz".to_string(),
                    kind: FactKind::Number,
                },
                Fact {
                    text: "stereo".to_string(),
                    kind: FactKind::Word,
                },
                Fact {
                    text: "1006 kbps".to_string(),
                    kind: FactKind::Number,
                },
            ],
            duration: Some(562.0),
            elapsed: Some(67.0),
            name_offset: 0,
        }
    }

    /// The name of a card too narrow for it is scrolled rather than cut short: the page is
    /// built at the box the layout settled on, so the test's card is the one a card of a
    /// measured box is built from, with the offset a tick would have reached.
    fn scrolled(card: &Card, width: u32, offset: i32) -> Page {
        let dc = unsafe { CreateCompatibleDC(None) };
        let metrics = TextMetrics::new(dc, 96, options().font_scale_percent).expect("metrics");
        let theme = text_theme::loaded(options().theme).expect("the bundled theme");
        let mut card = card.clone();
        card.name_offset = offset;

        let page = build_page(&card, theme, &metrics, width, width);
        unsafe {
            let _ = DeleteDC(dc);
        }

        page
    }

    fn options() -> AudioPreviewOptions {
        AudioPreviewOptions {
            theme: TextTheme::Dark,
            font_scale_percent: 125,
        }
    }

    /// The card is a page like any other painted preview's: measured before the window is
    /// placed, drawn into the box the layout settled on, and opaque everywhere.
    #[test]
    fn draws_the_card_a_sound_is_previewed_as() {
        let (width, height) = measure(&card(), 4096, 2160, 96, options()).expect("a measured card");
        assert!(
            width > 128 && height > 32,
            "a card is a page with room for a name and a bar in it: {width}x{height}"
        );

        let (pixels, width, height) =
            render(&card(), width, height, 96, options()).expect("a painted card");
        assert_eq!(pixels.len(), (width * height * 4) as usize);
        assert!(
            pixels.as_chunks::<4>().0.iter().all(|pixel| pixel[3] == 255),
            "a page is a page: what stands behind the card is the card"
        );
    }

    /// A card with nothing to say about its length still draws, and one with no room at all is
    /// answered with the page every painted preview gives in that place.
    #[test]
    fn draws_a_sound_whose_length_is_not_known() {
        let mut unknown = card();
        unknown.duration = None;

        let (width, height) = measure(&unknown, 4096, 2160, 96, options()).expect("a measured card");
        let painted = render(&unknown, width, height, 96, options()).expect("a painted card");
        assert_eq!(painted.0.len(), (painted.1 * painted.2 * 4) as usize);

        let (_, cramped) = measure(&card(), 8, 8, 96, options()).expect("a measured card");
        assert_eq!(cramped, 1, "a card with no room is one pixel and no page");
    }

    /// The facts of a track are the ones it holds: what a file does not say is left out rather
    /// than guessed at, and a file no engine named a codec for is named by its extension.
    #[test]
    fn writes_the_facts_a_file_has_and_no_others() {
        let track = Track {
            player: Player::Native,
            codec: Some("FLAC".to_string()),
            rate: Some(44_100),
            channels: Some(2),
            bitrate: Some(1_006_000),
            duration: Some(562.0),
        };
        let facts = facts_of(&track, Path::new("C:/music/track.flac"));
        let words: Vec<&str> = facts.iter().map(|fact| fact.text.as_str()).collect();
        assert_eq!(words, ["FLAC", "44.1 kHz", "stereo", "1006 kbps"]);

        let bare = Track {
            player: Player::Ffmpeg,
            codec: None,
            rate: None,
            channels: None,
            bitrate: None,
            duration: None,
        };
        let facts = facts_of(&bare, Path::new("C:/music/track.ape"));
        assert_eq!(facts.len(), 1, "nothing is said that the file did not say");
        assert_eq!(facts[0].text, "APE");
        assert_eq!(facts[0].kind, FactKind::Name);
    }

    /// The numbers a card is written with, at the lengths and rates a file can have.
    #[test]
    fn writes_a_position_the_way_a_player_does() {
        assert_eq!(clock(0.0), "0:00");
        assert_eq!(clock(67.4), "1:07");
        assert_eq!(clock(562.0), "9:22");
        assert_eq!(clock(3723.0), "1:02:03");
        assert_eq!(rate_label(44_100), "44.1 kHz");
        assert_eq!(rate_label(48_000), "48 kHz");
        assert_eq!(rate_label(22_050), "22.05 kHz");
        assert_eq!(channel_label(1), "mono");
        assert_eq!(channel_label(2), "stereo");
        assert_eq!(channel_label(6), "5.1");
        assert_eq!(bitrate_label(320_000), "320 kbps");
        assert_eq!(bitrate_label(320), "320 bps");
    }

    /// The bar is the played share of the whole, and a block crossing the track where there is
    /// no whole to measure against: what a card with no player shows is an empty bar.
    #[test]
    fn fills_the_bar_by_where_the_sound_is() {
        let mut known = card();
        known.duration = Some(100.0);
        known.elapsed = Some(25.0);
        assert_eq!(fill_span(&known, 400, 1.0), (0, 100));

        known.elapsed = Some(562.0);
        assert_eq!(
            fill_span(&known, 400, 1.0),
            (0, 400),
            "a position past the end is the whole of the bar and no more"
        );

        known.duration = None;
        let (_, width) = fill_span(&known, 400, 1.0);
        assert!(
            width > 0 && width < 400,
            "a sound whose length is not known is a block crossing the track: {width}"
        );

        known.elapsed = None;
        assert_eq!(
            fill_span(&known, 400, 1.0),
            (0, 0),
            "and a card at no volume at all is a bar with nothing in it"
        );
    }

    /// The readout is the played part and the whole, and it is dropped rather than guessed at
    /// where there is neither.
    #[test]
    fn writes_the_clock_the_card_is_drawn_with() {
        let mut card = card();
        let colors = ([0, 0, 0], [1, 1, 1]);

        let runs = clock_runs(&card, 400, 8, colors.0, colors.1);
        let text: String = runs.iter().map(|run| run.text.as_str()).collect();
        assert_eq!(text, "1:07 / 9:22");
        assert!(
            runs.iter().map(|run| run.width).sum::<i32>() <= 400,
            "the readout is drawn inside the room it was given"
        );

        card.duration = None;
        let text: String = clock_runs(&card, 400, 8, colors.0, colors.1)
            .iter()
            .map(|run| run.text.as_str())
            .collect();
        assert_eq!(text, "1:07", "a file that does not say how long it is has no whole");

        card.elapsed = None;
        assert!(
            clock_runs(&card, 400, 8, colors.0, colors.1).is_empty(),
            "and a card with no player behind it has nothing to say about time"
        );
    }

    /// A name the card has no room for is not a reason to draw a wider card: the width comes
    /// from the facts line and the bar under it, and the name is the thing that gives way.
    #[test]
    fn a_name_that_does_not_fit_does_not_widen_the_card() {
        let mut long = card();
        long.name = "18 - The Longest Track Name On This Album (Remastered, 2026).flac".to_string();
        let mut short = card();
        short.name = "2.flac".to_string();

        let long_box = measure(&long, 4096, 2160, 96, options()).expect("a measured card");
        let short_box = measure(&short, 4096, 2160, 96, options()).expect("a measured card");
        assert_eq!(
            long_box, short_box,
            "the card is the size its facts ask for whatever its name is"
        );

        // And the card really has no room for the name, which is what makes the two above the
        // same box rather than two names that both fit.
        let (width, _) = long_box;
        assert!(
            NameScroll::of(&long.name, width, 96, options()).moves(),
            "the name the test is about is one the card cannot fit"
        );
        assert!(
            !NameScroll::of(&short.name, width, 96, options()).moves(),
            "and one it can is left where it is"
        );
    }

    /// The name is drawn whole and clipped to the box the card leaves it — never cut short —
    /// and the offset a tick has reached is the whole of what moves: the origin of the run.
    #[test]
    fn draws_the_whole_name_where_the_offset_puts_it() {
        let mut long = card();
        long.name = "18 - The Longest Track Name On This Album (Remastered, 2026).flac".to_string();
        let (width, _) = measure(&long, 4096, 2160, 96, options()).expect("a measured card");

        let resting = scrolled(&long, width, 0);
        let name = resting.header.last().expect("the run the name is drawn in");
        assert_eq!(name.text, long.name, "the name is drawn whole, not cut short");
        assert_eq!(
            name.origin, name.x,
            "a name at rest starts at the left edge of the box the mark leaves it"
        );
        assert!(
            name.x + name.width <= width as i32,
            "and that box is the card's own content box: it ends inside the card"
        );

        let moved = scrolled(&long, width, 40);
        let name = moved.header.last().expect("the run the name is drawn in");
        assert_eq!(name.text, long.name, "a scrolled name is still the whole name");
        assert_eq!(name.origin, name.x - 40, "the offset is what moves it");
        assert_eq!(
            (
                resting.header.last().expect("the name").x,
                resting.header.last().expect("the name").width
            ),
            (name.x, name.width),
            "and the box it is clipped to stays where it is"
        );
    }

    /// The scroll starts at the beginning of the name, holds there, runs to the end of it,
    /// holds again and comes back — ping-ponging for as long as the card is up — and each end
    /// is where the page draws it.
    #[test]
    fn scrolls_the_name_from_one_end_of_itself_to_the_other_and_back() {
        let mut long = card();
        long.name = "18 - The Longest Track Name On This Album (Remastered, 2026).flac".to_string();
        let (width, _) = measure(&long, 4096, 2160, 96, options()).expect("a measured card");
        let step = Duration::from_millis(33);

        let mut scroll = NameScroll::of(&long.name, width, 96, options());
        assert!(scroll.moves() && scroll.offset() == 0, "it starts at the start");

        // The hold a card is put up with: repaints inside it move nothing at all.
        scroll.advance(Instant::now(), step);
        assert_eq!(scroll.offset(), 0, "the name rests before it begins to move");

        let first_end = scroll.travel;
        assert!(first_end > 0, "the name has an end past the box to reach");

        // Left until the end of the name comes into sight, where it stops and holds.
        let arrival = Instant::now() + NAME_HOLD + Duration::from_millis(1);
        for _ in 0..first_end + 1 {
            scroll.advance(arrival, step);
        }
        assert_eq!(scroll.offset(), first_end, "the name stops at the end of itself");
        scroll.advance(arrival, step);
        assert_eq!(
            scroll.offset(),
            first_end,
            "and holds there rather than turning at once"
        );

        let far = scrolled(&long, width, scroll.offset());
        let name = far.header.last().expect("the run the name is drawn in");
        assert_eq!(
            name.origin,
            name.x - first_end,
            "the far end of the ping-pong is the whole name shown"
        );

        // And back to the beginning, once the hold at the far end is over.
        let returning = arrival + NAME_HOLD + Duration::from_millis(1);
        for _ in 0..first_end {
            scroll.advance(returning, step);
        }
        assert_eq!(scroll.offset(), 0, "the name comes back to where it started");

        let home = scrolled(&long, width, scroll.offset());
        let name = home.header.last().expect("the run the name is drawn in");
        assert_eq!(name.origin, name.x, "and the near end is its resting place");
    }
}

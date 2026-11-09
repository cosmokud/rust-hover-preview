//! The page a card is laid out and painted from: the runs of text on it, the bar's row and the
//! band a press on it is answered against, the boxes of the controls a pinned card carries, and
//! the paint that fills them.
//!
//! None of it is kept: `measure` and `render` build the same page from the same card and hold
//! neither, so the row, the boxes and the paint are one set of arithmetic read three ways — which
//! is the whole of why a bar a person aims at and a bar drawn under a name are the same line.

use super::card::{
    clock, AudioPreviewOptions, Card, CardChrome, CardControl, Fact, FactKind, BAR_GAP_PIXELS,
    BAR_PIXELS, BAR_REACH_PIXELS, CHASE_DIVISOR, CHASE_SECONDS, CONTROL_GAP_PIXELS,
    CONTROL_SIDE_PIXELS, HEADER_LEVEL, MARK_CELL_ADVANCES, MIN_CONTENT_ADVANCES, RULE_GAP_PIXELS,
    RULE_PIXELS, TIME_GAP_ADVANCES, TRACK_PIXELS, WINDOW_BUTTON_CUSHION_PIXELS,
    WINDOW_BUTTON_GAP_PIXELS, WINDOW_BUTTON_REVEAL_PIXELS,
};
use crate::text::pin_chrome::{
    self, fill_disc, fill_ring, stroke_segment, surface_pixels, ControlGlyph,
};
use crate::text::text_paint::{
    blend, fill_rect, plain_style, readable, rgb, scaled, DibSurface, RunPainter, TextMetrics,
    TextStyle, BODY_LEVEL,
};
use crate::text::text_theme::LoadedTheme;
use windows::Win32::Foundation::RECT;
use windows::Win32::Graphics::Gdi::{CreateCompatibleDC, DeleteDC};

// ----------------------------------------------------------------------- page

/// One run of the card: a piece of text drawn from `origin` into the box `x`–`x + width`, in
/// one style.
pub(super) struct PageRun {
    pub(super) text: String,
    /// The box the run is laid out in and drawn in: how the runs beside it are placed, and
    /// what the run is clipped to.
    pub(super) x: i32,
    pub(super) width: i32,
    /// Where the text itself starts. It is the box's own left edge for every run of the card
    /// but one — a name the card has no room for, which is drawn whole and scrolled under the
    /// box, so what moves is the origin and what stays is the box that clips it (see
    /// `Card::name_offset`).
    pub(super) origin: i32,
    pub(super) style: TextStyle,
}

/// The bar and what is drawn over it.
#[derive(Debug)]
pub(super) struct Bar {
    pub(super) left: i32,
    pub(super) width: i32,
    pub(super) top: i32,
    pub(super) height: i32,
    pub(super) track_height: i32,
    /// Where the played part starts and how wide it is. A bar with no length to measure
    /// against — a file that does not say how long it is — is a block crossing the track
    /// instead, which is the difference between a sound and a stream.
    pub(super) fill_start: i32,
    pub(super) fill_width: i32,
}

/// The ripples a beat travels as: three, which is what the room a header line leaves a mark of
/// this size can hold apart — and three is what the loader the ripples come from draws.
const RIPPLE_COUNT: usize = 3;
/// The parting marks of the facts line: the one between two facts, and the one between the played
/// clock and the whole it is measured against — so a line of parts reads as one sentence.
const FACT_SEPARATOR: &str = " · ";
const CLOCK_SEPARATOR: &str = " / ";

/// How much softer the ripple at the end of its travel is drawn than one still standing just off
/// the dot, which is how that loader's own rings fade as they go: the beat as it arrives is the
/// full accent and the widest ripple a card can hold is two fifths of it (see `mark`).
const RIPPLE_FADE: f32 = 0.6;

/// The dot's radius as a share of what its ripples travel: a fifth, which is the share that
/// loader draws its own dot at — so the mark a card is left with is the small centre of a wide
/// ripple rather than a shape the widest of them only just clears (see `mark`).
const MARK_DOT_SHARE: f32 = 0.2;

/// The mark beside the name: the dot the card is drawn with, the ripples a beat travels as around
/// it, and where the two stand.
///
/// The dot is drawn rather than set as the bullet it replaces, and drawn smaller than that glyph
/// was: it is the centre the ripples are measured from, so it is the one part of the mark whose
/// size a beat does not change and whose centre never moves (see `mark`).
#[derive(Clone, Copy)]
pub(super) struct Mark {
    pub(super) center_x: f32,
    pub(super) center_y: f32,
    /// The dot's own radius, which a beat does not touch: the ripples travel and the dot stands.
    pub(super) radius: f32,
    /// How far out the front ripple stands with the beat at the top of its scale — the room the
    /// mark has, and the length a ripple's own travel is a share of (see `paint`).
    pub(super) reach: f32,
    /// The ripples' own radii, the smallest first, and all of them nothing for a card at rest.
    /// A ripple whose radius has fallen inside the dot is a ripple the beat has taken back in.
    pub(super) ripples: [f32; RIPPLE_COUNT],
    /// How thick each ripple's own line is drawn where the ripple stands a dot's radius off the
    /// mark: within that it is drawn thinner the nearer it is home, so a ripple that has arrived
    /// is a line of nothing (see `paint`).
    pub(super) thickness: f32,
}

/// The row the bar is drawn in: where it begins beneath the facts, how tall it is, and how thick
/// the track under the played part is.
///
/// These are the numbers the bar is drawn from *and* the numbers a press on it is answered
/// against, which is the whole of why they are one set rather than two: a bar laid out by one
/// arithmetic and hit-tested by another is a bar whose middle few pixels answer for a line
/// somewhere else on the card (see `bar_share_at`).
pub(super) struct BarRow {
    pub(super) top: i32,
    pub(super) height: i32,
    pub(super) track_height: i32,
}

/// The row the facts line is drawn in: where it begins under the rule, and how tall it is.
///
/// These are the numbers the line is drawn from *and* the band a pointer engages its scroll on, for
/// the reason [`BarRow`] is one set rather than two: a row laid out by one arithmetic and hit-tested
/// by another answers a hand somewhere else on the card (see `audio_preview::facts_rect`).
#[derive(Clone, Copy)]
pub(super) struct FactsRow {
    pub(super) top: i32,
    pub(super) height: i32,
}

/// The facts line's row: the same walk down the page `build_page` makes, from the rule under the name
/// to the line itself, whose height is the body line's.
pub(super) fn facts_row(metrics: &TextMetrics) -> FactsRow {
    let (_, rule_top) = header_line(metrics);

    FactsRow {
        top: rule_top + scaled(RULE_PIXELS, metrics.scale) + scaled(RULE_GAP_PIXELS, metrics.scale),
        height: metrics.line_height[BODY_LEVEL as usize],
    }
}

/// The bar's row: the bar's own height, whether or not buttons are standing in it. A button is
/// taller than that line, and stands centred on it rather than growing the row — which is what
/// makes a card a pin shows the same size as the card a hover shows, with the four buttons carved
/// out of the bar's line instead of the card growing to hold them (see `control_boxes`).
pub(super) fn bar_row(metrics: &TextMetrics) -> BarRow {
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
/// This is one band for every question that asks where the bar is (see `bar_share_at` and
/// `control_at`), because a band that is generous for one question and not for the other is a
/// bar whose own middle answers for a press and whose ends do not.
pub(super) struct BarBand {
    pub(super) left: i32,
    pub(super) right: i32,
    pub(super) top: i32,
    pub(super) bottom: i32,
}

impl BarBand {
    pub(super) fn contains(&self, x: i32, y: i32) -> bool {
        x >= self.left && x < self.right && y >= self.top && y < self.bottom
    }
}

pub(super) fn bar_band(row: BarRow, left: i32, right: i32, metrics: &TextMetrics) -> BarBand {
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
/// the three buttons at its left, the bar filling what is left of the row, the volume button
/// against its right edge, and the four window buttons standing in the top corner.
///
/// They are one walk rather than five numbers each because the card is drawn from these boxes and
/// pressed against them: a button laid out here and hit-tested elsewhere answers a press in the
/// middle of the facts line (see [`BarRow`]). Copied rather than borrowed so that one walk can
/// answer for both the bar's own edges and the whole row (see `bar_share_at`).
#[derive(Clone, Copy)]
pub(super) struct CardBoxes {
    pub(super) previous: RECT,
    pub(super) play: RECT,
    pub(super) next: RECT,
    pub(super) volume: RECT,
    pub(super) bar: RECT,
    pub(super) always_on_top: WindowButtonBox,
    pub(super) detach: WindowButtonBox,
    pub(super) minimize: WindowButtonBox,
    pub(super) close: WindowButtonBox,
}

/// One of the four window buttons a pinned card carries in its top corner, with the
/// two boxes it has: the one its glyph is drawn in, and the one a press on it is
/// answered against.
///
/// The two are not the same box, and that is the whole of the shape: the button is
/// small — twice the side the top margin alone would hold — so a hand is given the
/// window's own top edge above the drawn box and a sliver of the name line's
/// invisible leading below it (see [`window_button_boxes`]).
#[derive(Clone, Copy)]
pub(super) struct WindowButtonBox {
    /// The box the button's glyph is drawn in.
    pub(super) drawn: RECT,
    /// The box a press is answered against: the drawn box, widened to the window's
    /// own top edge above and by the cushion below.
    pub(super) hit: RECT,
}

impl CardBoxes {
    /// The box one of the row's controls is drawn in, which is the box a press on it is answered
    /// against — except the bar's, whose is its own box inside the row and whose band is wider,
    /// and the four window buttons', whose is the box their glyphs are drawn in plus the cushion
    /// above them (see [`WindowButtonBox`]).
    pub(super) fn rect(&self, control: CardControl) -> Option<RECT> {
        Some(match control {
            CardControl::Previous => self.previous,
            CardControl::Play => self.play,
            CardControl::Next => self.next,
            CardControl::Seek => self.bar,
            CardControl::Volume => self.volume,
            CardControl::AlwaysOnTop => self.always_on_top.hit,
            CardControl::Detach => self.detach.hit,
            CardControl::Minimize => self.minimize.hit,
            CardControl::Close => self.close.hit,
        })
    }

    /// Which of the window buttons in the top margin a point is on, if any.
    ///
    /// Asked of the cushion each is answered against rather than of the drawn boxes,
    /// because a hand aiming at a button this small is given the room above it
    /// (see [`window_button_boxes`]).
    pub(super) fn window_button_at(&self, x: i32, y: i32) -> Option<CardControl> {
        if holds(self.always_on_top.hit, x, y) {
            return Some(CardControl::AlwaysOnTop);
        }
        if holds(self.detach.hit, x, y) {
            return Some(CardControl::Detach);
        }
        if holds(self.minimize.hit, x, y) {
            return Some(CardControl::Minimize);
        }
        if holds(self.close.hit, x, y) {
            return Some(CardControl::Close);
        }
        None
    }
}

/// The row's boxes, laid out for a card of `width` pixels at a display's scale — or nothing at all
/// where the card carries no controls, which is the card a hover shows and the card every caller
/// that is only asking *where* something is uses.
pub(super) fn control_boxes(
    metrics: &TextMetrics,
    width: u32,
    controls: bool,
) -> Option<CardBoxes> {
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
    let (always_on_top, detach, minimize, close) = window_button_boxes(metrics, width as i32);

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
        always_on_top,
        detach,
        minimize,
        close,
    })
}

/// Where the name line's own box and the rule under it stand: the line is
/// centred between the window's own top border and the rule, so the ink of
/// its glyphs — the neon circle and the name — is as far from the one as
/// from the other. The rule keeps the place it has always had: the line's
/// own box moves within the room above it, and the card keeps the height it
/// has always had.
///
/// Worked out here for the caller that has no page in hand — the hit test,
/// which is asked of the card's own layout rather than of a page that was
/// built — as well as for `build_page`, which lays
/// the page out around the same two lines.
fn header_line(metrics: &TextMetrics) -> (i32, i32) {
    let header_height = metrics.line_height[HEADER_LEVEL as usize];
    let rule_gap = scaled(RULE_GAP_PIXELS, metrics.scale);

    let rule_top = metrics.padding + header_height + rule_gap;
    let header_top = (rule_top
        - metrics.ink_top[HEADER_LEVEL as usize]
        - metrics.ink_bottom[HEADER_LEVEL as usize]
        - 1)
        / 2;

    (header_top, rule_top)
}

/// The side one of the four window buttons is drawn at: twice what the card's own top
/// margin would hold on its own — the margin's height less the gap above the button
/// and the gap below it — which is room the margin does not have, so a button stands
/// in the name line's own rows as well (see [`window_button_boxes`]).
fn window_button_side(metrics: &TextMetrics) -> i32 {
    let gap = scaled(WINDOW_BUTTON_GAP_PIXELS, metrics.scale);
    (metrics.padding - gap * 2).max(0) * 2
}

/// The bottom of the box a press on a window button is answered against: the drawn
/// box's own bottom plus the cushion below it, which is the bottom of the band the
/// buttons stand in (see [`window_button_boxes`]).
fn window_button_reach(metrics: &TextMetrics) -> i32 {
    let gap = scaled(WINDOW_BUTTON_GAP_PIXELS, metrics.scale);
    window_button_side(metrics) + gap + scaled(WINDOW_BUTTON_CUSHION_PIXELS, metrics.scale)
}

/// The band a pointer is near the window's own top border in, at the display's scale
/// and font size: the bottom of the box a press on any window button is answered
/// against. A hand anywhere in it is a hand in the band the buttons stand in, which is
/// one of the two axes a hand asks for them on — the other being [`window_button_span`]
/// — the one thing on a sound's card that comes and goes, because a button twice the
/// margin's own side stands in the name's own rows (see [`window_button_reach`]).
pub(crate) fn window_button_band(dpi: u32, options: AudioPreviewOptions) -> i32 {
    window_button_layout(dpi, options, window_button_reach).unwrap_or(0)
}

/// The columns the four window buttons and a hand's margin around them reach across, on a
/// card of `width` pixels: from the leftmost button's own left to the rightmost one's own
/// right, widened by `WINDOW_BUTTON_REVEAL_PIXELS` on each side at the display's scale.
///
/// The horizontal half of what asks for the buttons — the band is the vertical one (see
/// [`window_button_band`]) — so a hand in the top band but across the card from them is a
/// hand near neither.
pub(crate) fn window_button_span(dpi: u32, options: AudioPreviewOptions, width: i32) -> (i32, i32) {
    window_button_layout(dpi, options, |metrics| {
        let (always_on_top, _, _, close) = window_button_boxes(metrics, width);
        let margin = scaled(WINDOW_BUTTON_REVEAL_PIXELS, metrics.scale);
        (
            always_on_top.drawn.left - margin,
            close.drawn.right + margin,
        )
    })
    .unwrap_or((0, 0))
}

/// Window-button arithmetic asked of a throwaway display context made for this one pass and
/// released here: the buttons are laid out from the metrics of the font the card is built at,
/// which no caller keeps (see [`window_button_band`] and [`window_button_span`]).
fn window_button_layout<R>(
    dpi: u32,
    options: AudioPreviewOptions,
    f: impl FnOnce(&TextMetrics) -> R,
) -> Option<R> {
    let dc = unsafe { CreateCompatibleDC(None) };
    if dc.0.is_null() {
        return None;
    }

    let answer = TextMetrics::new(dc, dpi, options.font_scale_percent).map(|metrics| f(&metrics));

    unsafe {
        let _ = DeleteDC(dc);
    }

    answer
}

/// The four window buttons a pinned card carries in its top corner: the always-on-top that keeps
/// the window over other windows, the detach that hands it to a process of its own, the minimize
/// that shrinks the pin into the bubble it leaves, and the close that ends the pin, the player and
/// the window together.
///
/// Each button is a square stood against the card's right border with the same gap on every
/// side that has room — above it to the window's own top edge, between each pair of
/// them, and to the border — and a side twice what the margin alone would hold, which is room the
/// margin does not have: the buttons stand in the name line's own rows, and the name runs
/// underneath them while they are up, its box and its scroll unchanged (see `paint`).
///
/// A button's *hit* box is its drawn box widened upward to the window's own top edge and by the
/// cushion below it: the buttons are small, so a hand is given the room above them and
/// a pixel of the name line's invisible leading, and the name's own rows below that as well —
/// the room the buttons stand in, which a hit box that stopped at the name line's top would
/// cut off, leaving a box smaller than the button it answers for. The cushion is a logical
/// pixel like the gap it is counted from, so it is the same share of the button at every
/// scale (see `WINDOW_BUTTON_CUSHION_PIXELS`).
pub(super) fn window_button_boxes(
    metrics: &TextMetrics,
    width: i32,
) -> (
    WindowButtonBox,
    WindowButtonBox,
    WindowButtonBox,
    WindowButtonBox,
) {
    let gap = scaled(WINDOW_BUTTON_GAP_PIXELS, metrics.scale);
    let side = window_button_side(metrics);

    // The close against the right border, the minimize beside it, the detach to the left of
    // the minimize and always-on-top to the left of the detach, each with the same gap and side.
    let close = RECT {
        left: (width - gap - side).max(0),
        top: gap,
        right: (width - gap).max(0),
        bottom: gap + side,
    };
    let minimize = RECT {
        left: (close.left - gap - side).max(0),
        top: gap,
        right: (close.left - gap).max(0),
        bottom: gap + side,
    };
    let detach = RECT {
        left: (minimize.left - gap - side).max(0),
        top: gap,
        right: (minimize.left - gap).max(0),
        bottom: gap + side,
    };
    let always_on_top = RECT {
        left: (detach.left - gap - side).max(0),
        top: gap,
        right: (detach.left - gap).max(0),
        bottom: gap + side,
    };

    // The cushion below the drawn box: the gap the button stands in, plus a pixel
    // of the name line's invisible leading — a logical pixel like the gap it is
    // counted from, so the cushion is the same share of the button at every scale
    // (see `WINDOW_BUTTON_CUSHION_PIXELS`). The hit box it ends reaches into the
    // name's own rows, which is the room the buttons stand in: they are twice the
    // margin's own side, and a hit box that stopped at the name line's top would
    // answer for less than the button it is asked about.
    let bottom = window_button_reach(metrics);

    (
        WindowButtonBox {
            drawn: always_on_top,
            hit: RECT {
                left: always_on_top.left,
                top: 0,
                right: always_on_top.right,
                bottom,
            },
        },
        WindowButtonBox {
            drawn: detach,
            hit: RECT {
                left: detach.left,
                top: 0,
                right: detach.right,
                bottom,
            },
        },
        WindowButtonBox {
            drawn: minimize,
            hit: RECT {
                left: minimize.left,
                top: 0,
                right: minimize.right,
                bottom,
            },
        },
        WindowButtonBox {
            drawn: close,
            hit: RECT {
                left: close.left,
                top: 0,
                right: close.right,
                bottom,
            },
        },
    )
}

/// Whether a point is inside a box, in the card's own coordinates. Half-open at the right and the
/// bottom, so that two boxes which touch are two boxes and not one row of pixels answered twice.
pub(super) fn holds(rect: RECT, x: i32, y: i32) -> bool {
    x >= rect.left && x < rect.right && y >= rect.top && y < rect.bottom
}

/// A painted card: what to draw and how big it came out.
pub(super) struct Page {
    pub(super) header: Vec<PageRun>,
    pub(super) header_top: i32,
    pub(super) header_height: i32,
    pub(super) rule_top: i32,
    pub(super) rule_height: i32,
    pub(super) facts: Vec<PageRun>,
    pub(super) facts_top: i32,
    pub(super) facts_height: i32,
    pub(super) bar: Bar,
    /// The mark the card is drawn with: the dot it is marked by and the ripples a beat stands
    /// around it, which are none at all for a card at rest — a card with no player running
    /// behind it, hovered or pinned (see `Card::beat` and `Mark`).
    pub(super) mark: Option<Mark>,
    /// The boxes the row's controls are drawn in and answered against, and nothing at all for a
    /// card that carries none.
    pub(super) boxes: Option<CardBoxes>,
    /// What those controls are saying, which is the wash under the pointer and the glyph on each
    /// of them: kept beside `boxes` rather than inside it because the hit test rebuilds the boxes
    /// and has no business with any of this (see `control_at`).
    pub(super) chrome: Option<CardChrome>,
    pub(super) width: u32,
    pub(super) height: u32,
    pub(super) padding: i32,
}

/// The mark a card is drawn with: the dot, and the ripples a beat has stood around it.
///
/// The dot is centred on the mark's own cell — half a header advance in from the card's margin,
/// which is where the ink of the glyph it replaces stood — and, on the line, on the vertical
/// middle of the text line it marks (`ink_top` to `ink_bottom`), shifted by any configured
/// `offset_y`. The ripples are drawn from that same centre, so what
/// a beat does is travel around the mark and never off it. The room they
/// have is the nearest of the three things they could reach: the two edges of the card the mark
/// stands between and the box the name begins in, each less half a ripple's own line and a pixel
/// of air — so the widest ripple is the widest a mark of this size can be drawn without lapping
/// the letter beside it, and the dot is a fifth of it.
///
/// The ripples travel as a train rather than each on its own: the front one leads and the ones
/// behind it keep their spacing, so a beat of half the scale stands the front halfway out and a
/// bass that drops pulls it back in, taking the ripples behind it under the dot one at a time —
/// the same travel the single ring made, drawn as the ripple it was always meant to be (see
/// `Card::beat`).
fn mark(beat: f32, metrics: &TextMetrics, header_top: i32, name_left: i32, offset_y: f32) -> Mark {
    let beat = beat.clamp(0.0, 1.0);

    let advance = metrics.advance[HEADER_LEVEL as usize].max(1) as f32;
    let ink_top = metrics.ink_top[HEADER_LEVEL as usize] as f32;
    let ink_bottom = metrics.ink_bottom[HEADER_LEVEL as usize] as f32;
    let center_x = metrics.padding as f32 + advance / 2.0;
    let center_y = header_top as f32 + (ink_top + ink_bottom + 1.0) / 2.0 + offset_y;
    let thickness = scaled(1, metrics.scale) as f32;
    let reach =
        (center_x.min(center_y).min(name_left as f32 - center_x) - thickness / 2.0 - 1.0).max(0.0);
    let radius = reach * MARK_DOT_SHARE;

    // The train a beat sends out: the front ripple travels the whole room and the ones behind it
    // follow at the same spacing, so a beat that only whispers stands a ripple just off the dot
    // rather than crowding three of them into a blur, and the ripples a falling bass has taken
    // back inside the dot are the ones the card no longer draws (see `paint`).
    let mut ripples = [0.0; RIPPLE_COUNT];
    if beat > 0.0 {
        let travel = (reach - radius).max(0.0);
        let front = travel * beat;
        let gap = travel / RIPPLE_COUNT as f32;
        for (index, ripple) in ripples.iter_mut().enumerate() {
            *ripple = radius + front - gap * (RIPPLE_COUNT - 1 - index) as f32;
        }
    }

    Mark {
        center_x,
        center_y,
        radius,
        reach,
        ripples,
        thickness,
    }
}

pub(super) fn build_page(
    card: &Card,
    theme: &LoadedTheme,
    metrics: &TextMetrics,
    box_width: u32,
    box_height: u32,
    mark_offset_y: f32,
) -> Page {
    let header_advance = metrics.advance[HEADER_LEVEL as usize].max(1);
    let body_advance = metrics.advance[BODY_LEVEL as usize].max(1);
    let header_height = metrics.line_height[HEADER_LEVEL as usize];
    let padding = metrics.padding;

    let rule_height = scaled(RULE_PIXELS, metrics.scale);
    let bar_row = bar_row(metrics);
    let facts_row = facts_row(metrics);

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

    // The width the card takes: the room it is given, floored at the narrowest
    // card worth drawing — the card fills the room and its seek bar stretches with
    // it, so every share of the display is visibly a different size and every
    // file's card is the same width at a given share. The facts line is no longer
    // in that count: the room decides how many facts fit, the last giving way
    // first (see `fact_runs`). The name is not in the count either, and that is
    // the one rule this line is: a name too long for the card is not a reason to
    // ask the display for a card a hundred advances wide, so what gives way is the
    // name, which is drawn whole and scrolled across the card (see
    // `Card::name_offset`).
    let mark_cell = header_advance * MARK_CELL_ADVANCES;
    let content = ((box_width as i32) - padding * 2).max(body_advance * MIN_CONTENT_ADVANCES);
    let width = (content + padding * 2).max(1) as u32;

    // A room with no room for a name and a line under it is a room that cannot be
    // drawn in, which is the answer an archive page gives in the same place. The
    // width is never the reason: the floor above keeps the card at the narrowest
    // card worth drawing, so a room of a few pixels answers with that card's width
    // and a one-pixel height rather than with nothing at all.
    if (box_height as i32) < padding * 2 + header_height {
        return empty_page(width, padding);
    }

    let content_left = padding;
    let content_right = width as i32 - padding;

    // The name is drawn whole and clipped to the box the mark leaves it, and one the box has no
    // room for is scrolled under that box rather than cut short: what moves is where the run
    // starts, and the box that clips it never does. A run's own origin is its box's left edge
    // for every other run of the card, so this is the one place the two are told apart (see
    // `PageRun::origin` and `Card::name_offset`).
    let name_left = content_left + mark_cell;
    let header = vec![PageRun {
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
    }];

    let (header_top, rule_top) = header_line(metrics);
    let facts_top = facts_row.top;

    let bar_width = content_right - content_left;
    // The facts line ends where the clock's readout begins, less the gap kept off it: the room it is
    // drawn in is the room the readout really takes rather than one reserved for the longest readout
    // there is (see `facts_right`), and a line that runs past it ends on the last fact it holds whole
    // (see `facts_edge`), so nothing of it is cut at rest.
    let mut facts = fact_runs(
        &card.facts,
        content_left,
        facts_edge(
            &card.facts,
            content_left,
            facts_right(card, content_right, body_advance),
            body_advance,
        ),
        card.facts_offset,
        body_advance,
        muted,
        foreground,
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
        facts_height: facts_row.height,
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
        mark: Some(mark(
            card.beat,
            metrics,
            header_top,
            name_left,
            mark_offset_y,
        )),
        width,
        // The height is the card's own: the name line, the rule, the facts line
        // and the bar, with the card's own margin under it — content-driven at
        // every share, so the card stays a one-line strip whatever room it is
        // given (the room's own height is only the guard above, see
        // `audio_box_room`).
        height: ((bar_row.top + bar_row.height + padding) as u32).max(1),
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
        mark: None,
        width,
        height: 1,
        padding,
    }
}

/// Where the played part of the bar starts and how wide it is: the played share of the whole
/// where the file says how long it is, and a block crossing the track where it does not.
pub(super) fn fill_span(card: &Card, bar_width: i32, scale: f32) -> (i32, i32) {
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

/// What the readout at the right edge of the facts line says and what has been played in the
/// accent, and the whole it is measured against in the page's own muted gray. A card with no player
/// behind it has only the whole to show, and one whose file does not say how long it is has only the
/// number that moves.
pub(super) fn clock_runs(
    card: &Card,
    right: i32,
    advance: i32,
    muted: [u8; 3],
    accent: [u8; 3],
) -> Vec<PageRun> {
    let parts = clock_parts(card);
    let total: i32 = parts.iter().map(|text| text_width(text, advance)).sum();
    let mut x = right - total;
    // The played number is the one part of the readout that moves, and it is drawn in the accent the
    // bar's played part is drawn in; the whole beside it is the page's muted gray.
    let played = card.elapsed.is_some();

    parts
        .into_iter()
        .enumerate()
        .map(|(index, text)| {
            let width = text_width(&text, advance);
            let run = PageRun {
                text,
                x,
                width,
                origin: x,
                style: {
                    let mut style = plain_style(BODY_LEVEL);
                    style.foreground = if index == 0 && played { accent } else { muted };
                    style
                },
            };
            x += width;
            run
        })
        .collect()
}

/// The readout at the right edge of the facts line, part by part: what has been played, the whole it
/// is measured against, and nothing at all where there is neither. Measured as well as drawn from
/// these, so the room the facts give up for it is the room it takes (see `facts_right`).
fn clock_parts(card: &Card) -> Vec<String> {
    match (card.elapsed, card.duration) {
        (Some(elapsed), Some(duration)) if duration > 0.0 => {
            vec![clock(elapsed), CLOCK_SEPARATOR.to_string(), clock(duration)]
        }
        (Some(elapsed), _) => vec![clock(elapsed)],
        (None, Some(duration)) if duration > 0.0 => vec![clock(duration)],
        _ => Vec::new(),
    }
}

/// How wide the readout is, which is the room the facts line gives up for it.
fn clock_width(card: &Card, advance: i32) -> i32 {
    clock_parts(card)
        .iter()
        .map(|text| text_width(text, advance))
        .sum()
}

/// Where the facts line ends: one `TIME_GAP_ADVANCES` gap left of the readout the clock is drawn in,
/// measured as it will be drawn rather than reserved for the longest readout there is — the room
/// `1:02:03 / 1:02:03` needs is a file an hour long's, and every card under one was giving eighteen
/// advances of its facts line away for a clock that never took them.
///
/// A card with no clock behind it ends the line at the content box's own edge: there is nothing to
/// keep the room off.
pub(super) fn facts_right(card: &Card, content_right: i32, advance: i32) -> i32 {
    let clock = clock_width(card, advance);

    if clock == 0 {
        content_right
    } else {
        content_right - clock - advance * TIME_GAP_ADVANCES
    }
}

/// The facts a line draws, in the order it draws them: a fact with nothing to say is left out whole
/// rather than given a parting mark of its own — `channels` is empty for a sound that states none,
/// and a mark with nothing behind it is a mark between nothing and nothing.
fn laid_facts(facts: &[Fact]) -> impl Iterator<Item = &Fact> {
    facts.iter().filter(|fact| !fact.text.is_empty())
}

/// How wide the whole facts line wants to be, its parting marks included — the marks it really draws,
/// one between two facts rather than one before every fact but the first: the length a card the room
/// has no place for scrolls across itself (see `Card::facts_offset`).
pub(super) fn facts_width(facts: &[Fact], advance: i32) -> i32 {
    let mut width = 0;

    for (index, fact) in laid_facts(facts).enumerate() {
        if index > 0 {
            width += text_width(FACT_SEPARATOR, advance);
        }
        width += text_width(&fact.text, advance);
    }

    width
}

/// The right edge of the box the facts line is drawn in: the room the clock leaves, drawn back to the
/// end of the last whole fact that fits in it — so a box that cannot hold the line ends it on a fact
/// rather than cutting the one that runs past it (see `fact_runs`).
///
/// A line the room holds whole keeps the room's own edge: there is nothing to draw back from. So does
/// a room too narrow for even the first fact, which is that fact cut (see `fact_runs`).
///
/// Read of the line at rest, so the edge is one number per card: what the box clips as the line
/// travels under it is the scroll's business, not the edge's.
pub(super) fn facts_edge(facts: &[Fact], left: i32, right: i32, advance: i32) -> i32 {
    let mut edge = left;

    for (index, fact) in laid_facts(facts).enumerate() {
        let mark = if index == 0 {
            0
        } else {
            text_width(FACT_SEPARATOR, advance)
        };
        let next = edge + mark + text_width(&fact.text, advance);
        if next > right {
            break;
        }
        edge = next;
    }

    if edge == left || edge - left == facts_width(facts, advance) {
        right
    } else {
        edge
    }
}

/// The facts as runs, each in its own ink, laid from `left` into the room that ends at `right` — the
/// room the clock leaves, drawn back to the last fact it holds whole (see [`facts_edge`]) — with a
/// parting mark before every fact after the first.
///
/// A fact is drawn whole where the whole of it fits: the room is the room the readout really leaves,
/// so a fact that fits is never the one dropped. A line the room has no space for is the card's second
/// scroll rather than a cut line: `offset` is how far the whole line has been carried back under the
/// box, and what the box cannot hold is out of sight for now (see `Card::facts_offset`).
///
/// A parting mark is laid only where the fact after it is: the line ends on a character at every
/// offset of its travel, never on the mark of a fact that has not come into the room yet.
// Eight: the line, the room it is drawn in, and the three inks a fact can be set in — the page
// reads all of them in one place and hands them to the one walk that lays them out.
#[allow(clippy::too_many_arguments)]
fn fact_runs(
    facts: &[Fact],
    left: i32,
    right: i32,
    offset: i32,
    advance: i32,
    muted: [u8; 3],
    foreground: [u8; 3],
    number: [u8; 3],
) -> Vec<PageRun> {
    let mut runs = Vec::new();
    let mut x = left - offset.max(0);
    let mut first = true;

    for fact in laid_facts(facts) {
        let mark = if first {
            0
        } else {
            text_width(FACT_SEPARATOR, advance)
        };
        let start = x + mark;
        let width = text_width(&fact.text, advance);
        let laid = (start + width).min(right) > start.max(left);

        if laid {
            if !first {
                push_fact_run(&mut runs, FACT_SEPARATOR, x, left, right, advance, muted);
            }

            let ink = match fact.kind {
                FactKind::Format => number,
                FactKind::Reading => foreground,
                FactKind::Note => muted,
            };
            push_fact_run(&mut runs, &fact.text, start, left, right, advance, ink);
        }

        x = start + width;
        first = false;
    }

    runs
}

/// One piece of the facts line laid at `x`, drawn into the part of it inside the box `left`–`right`
/// and nowhere else — so nothing of the line is ever drawn over the clock — and the `x` the piece
/// after it begins at.
///
/// The piece is clipped to that box rather than cut to it: a line the room cannot hold is a line
/// scrolled across the card, and what is past the edge of the room is out of sight for now (see
/// `Card::facts_offset`).
fn push_fact_run(
    runs: &mut Vec<PageRun>,
    text: &str,
    x: i32,
    left: i32,
    right: i32,
    advance: i32,
    ink: [u8; 3],
) -> i32 {
    let width = text_width(text, advance);
    let box_left = x.max(left);
    let box_right = (x + width).min(right);

    if box_right > box_left {
        runs.push(PageRun {
            text: text.to_string(),
            x: box_left,
            width: box_right - box_left,
            origin: x,
            style: {
                let mut style = plain_style(BODY_LEVEL);
                style.foreground = ink;
                style
            },
        });
    }

    x + width
}

// -------------------------------------------------------------------- drawing

pub(super) fn paint(surface: &DibSurface, page: &Page, theme: &LoadedTheme, scale: f32) {
    let page_color = rgb(theme.background());
    let foreground = rgb(theme.foreground());
    let rule_color = blend(page_color, foreground, 0.18);
    let number = readable(
        rgb(theme.style_for_scopes(&["constant.numeric"]).foreground),
        page_color,
    );
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

    if let Some(mark) = page.mark {
        // The mark and its ripples are drawn over the runs rather than under them, because each
        // run's own box is filled with the page as it is painted (see `paint_run`) — and they are
        // drawn from one centre, so a beat travels around the dot rather than off it.
        let buffer = unsafe {
            std::slice::from_raw_parts_mut(
                surface_pixels(surface),
                surface.width as usize * surface.height as usize * 4,
            )
        };
        fill_disc(
            buffer,
            surface.width as i32,
            mark.center_x,
            mark.center_y,
            mark.radius,
            accent,
            1.0,
        );
        for radius in mark.ripples {
            // A ripple the beat has taken back inside the dot is one the card has put away, and
            // so is every ripple of a card at rest: what is left of a beat is its front.
            if radius <= mark.radius {
                continue;
            }
            // The line it is drawn with narrows as the ripple comes home — its whole width once it
            // stands a dot's radius out, a share of that within, and a line of nothing at the dot
            // — which is what keeps the end of a fall from being a change in the mark's size: the
            // ripples have gone by the time the beat reaches its rest (see `Mark::thickness`).
            let out = radius - mark.radius;
            let home = (out / mark.radius).clamp(0.0, 1.0);
            let share = (out / (mark.reach - mark.radius)).clamp(0.0, 1.0);
            fill_ring(
                buffer,
                surface.width as i32,
                mark.center_x,
                mark.center_y,
                radius,
                mark.thickness * home,
                accent,
                1.0 - RIPPLE_FADE * share,
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

        // The window buttons in the top corner, painted last of all and painted as
        // nothing but their marks: no wash under them and no box around them in any
        // state, the ink being the only thing that changes. At rest it is the card's
        // own foreground, which is the ink the row's glyphs are painted in, and while
        // the pointer is on one or holding it it is the accent, which is the color the
        // played part of the bar is drawn in.
        //
        // The pin beside them says something of its own — whether the window is kept on top — so
        // it says it in the ink the numbers on the facts line are read in, and under the hand a
        // brighter one: the accent is what a button with no state of its own wears there, and a pin
        // kept on top that took it would say nothing but that the hand is near.
        //
        // The detach beside it is the second of the two, and says the same way what became of the
        // window it stands on: a pin is a window this process is showing the file in, and a detached
        // window is one of its own — so the button that made it one wears the ink of a state it is in
        // rather than the ink of a choice it still offers (see `CardChrome::detached`).
        //
        // They are the one thing on the card that comes and goes: a button twice the
        // margin's own side stands in the name's own rows, so a hand near the buttons
        // themselves — in their own band and across their own columns — is what asks for
        // them (see `window_button_band` and `window_button_span`), and while they are up
        // the name runs underneath them —
        // the box each stands in is the page's own colour, which is what keeps a name
        // scrolling under one from running under its mark.
        if chrome.window_buttons {
            for (control, button) in [
                (CardControl::AlwaysOnTop, &boxes.always_on_top),
                (CardControl::Detach, &boxes.detach),
                (CardControl::Minimize, &boxes.minimize),
                (CardControl::Close, &boxes.close),
            ] {
                fill_rect(surface, button.drawn, page_color);

                let lit = chrome.hovered == Some(control) || chrome.pressed == Some(control);
                // The two window buttons that carry a state of the window's own — that it is kept on
                // top, that it is a window of its own — are drawn in the numbers' ink, and light a
                // step further toward white under the hand. The other two are the card's own ink, lit
                // to the accent, which is what a button with no state to show wears.
                let wearing_state = match control {
                    CardControl::AlwaysOnTop => chrome.always_on_top,
                    CardControl::Detach => chrome.detached,
                    _ => false,
                };
                let color = match (wearing_state, lit) {
                    (false, false) => foreground,
                    (false, true) => accent,
                    (true, false) => number,
                    // Lit a step further toward white: the mark keeps the ink its state is read in
                    // and is the brightest thing in that corner while the hand is on it.
                    (true, true) => blend(number, [255, 255, 255], 0.35),
                };
                paint_window_button(surface, button.drawn, control, color, scale);
            }
        }
    }
}

/// The card's own pixels, in the form the marks are written in. It is taken per mark rather
/// than once for the whole paint so that a mark drawn as boxes of ink goes through `fill_rect`
/// with no other borrow of the surface live beside it (see `paint_window_button`).
#[allow(clippy::mut_from_ref)] // The pixels are the surface's own, behind the pointer `surface_pixels` hands back.
fn card_pixels(surface: &DibSurface) -> &mut [u8] {
    unsafe {
        std::slice::from_raw_parts_mut(
            surface_pixels(surface),
            surface.width as usize * surface.height as usize * 4,
        )
    }
}

/// One of the window buttons, drawn as the one mark that says what it
/// is: a pin for the always-on-top, two windows for the detach, a dash for the minimize
/// and two crossing strokes for the close.
///
/// The marks are drawn by hand rather than with the caption's glyph machinery, because that
/// machinery's smallest mark has a floor of six pixels and a button this small is barely
/// wider than that — a caption's mark would overflow the box it stands in (see
/// `pin_chrome::paint_glyph`). Each mark is a stroke one logical pixel thick, centred in
/// the box and clear of its edges, and nothing of the button is drawn but the stroke: the
/// ink is the only thing that changes, in any state.
fn paint_window_button(
    surface: &DibSurface,
    rect: RECT,
    control: CardControl,
    ink: [u8; 3],
    scale: f32,
) {
    let width = surface.width as i32;

    // The centre of the box, on the middle of a pixel rather than on the edge of one,
    // which is where a one-pixel stroke is a whole pixel rather than a half of two.
    let center_x = ((rect.left + rect.right) / 2) as f32 + 0.5;
    let center_y = ((rect.top + rect.bottom) / 2) as f32 + 0.5;
    let stroke = scaled(1, scale) as f32;

    // The mark spans the middle of the box and is clear of its edges: a quarter of the
    // side in from each end, which is half the box for the dash to run and the span of
    // the cross's own diagonal.
    let side = (rect.right - rect.left) as f32;
    let inset = side / 4.0;

    match control {
        CardControl::AlwaysOnTop => {
            // The box the cross's own strokes are painted in, which is the box the pin is drawn
            // in as well: its ends sit on pixel centres a quarter of the side in from each edge,
            // and a stroke is painted in every pixel a half-stroke's reach away from them,
            // snapped out to whole pixels. It is this box rather than a square of the pin's own
            // that makes it stand exactly the height of the cross beside it (see `draw_pin_top`,
            // which takes the square the caption's own marks are drawn in and is a size larger).
            let reach = (stroke / 2.0).max(0.5) + 0.5;
            let mark = RECT {
                left: (rect.left as f32 + inset - reach).floor() as i32 + 1,
                top: (rect.top as f32 + inset - reach).floor() as i32 + 1,
                right: (rect.right as f32 - inset + reach - 1.0).ceil() as i32,
                bottom: (rect.bottom as f32 - inset + reach - 1.0).ceil() as i32,
            };

            let thick = stroke.round().max(1.0) as i32;
            let middle_x = (mark.left + mark.right) / 2;
            let middle_y = (mark.top + mark.bottom) / 2;
            // Where the arrow's head begins, which is a row below the ceiling and never past the
            // mark's last row: a mark too small for a stem is one the head is drawn at the bottom
            // of rather than one whose head is drawn outside the mark (see `mark`).
            let head_top = (mark.top + thick + 1).min(mark.bottom - 1);

            // The ceiling: the mark's own first rows, as wide as the mark.
            fill_rect(
                surface,
                RECT {
                    left: mark.left,
                    top: mark.top,
                    right: mark.right,
                    bottom: (mark.top + thick).min(mark.bottom),
                },
                ink,
            );

            // The stem down to the mark's last row, as wide as the ceiling is thick.
            let stem_left = middle_x - thick / 2;
            fill_rect(
                surface,
                RECT {
                    left: stem_left.max(mark.left),
                    top: head_top,
                    right: (stem_left + thick).min(mark.right),
                    bottom: mark.bottom,
                },
                ink,
            );

            // The head: two arms from the top of the stem out to the mark's middle row.
            let tip = (middle_x as f32 + 0.5, head_top as f32 + 0.5);
            let arm = (((mark.right - mark.left) / 2) - thick - 1).max(1) as f32;
            for direction in [-1.0, 1.0] {
                stroke_segment(
                    card_pixels(surface),
                    width,
                    tip,
                    (
                        middle_x as f32 + 0.5 + direction * arm,
                        middle_y as f32 + 0.5,
                    ),
                    thick as f32,
                    ink,
                    1.0,
                );
            }
        }
        CardControl::Detach => {
            // Two windows, one stepped out behind the other, as the caption's own detach button is
            // drawn (see `draw_detached_window`): the front window's whole outline and the back
            // one's top and right edges. Runs of whole rows and columns rather than strokes for the
            // reason the dash and the cross are — a box this small has no room for an edge that is
            // a ramp on either side of itself — and the mark stands inside the box's own inset,
            // where the marks beside it stand.
            let thick = stroke.round().max(1.0) as i32;
            let quarter = inset as i32;
            let left = rect.left + quarter;
            let top = rect.top + quarter;
            let right = rect.right - quarter;
            let bottom = rect.bottom - quarter;
            // How far the window behind is stepped out from the one in front: a quarter of the
            // mark, and never so far that the window in front has no room left to be one.
            let step = ((right - left) / 4).max(1).min((right - left) / 2);

            // The window in front: the whole outline, stepped down and left of the one behind it.
            let (front_left, front_top) = (left, top + step);
            let (front_right, front_bottom) = (right - step, bottom);
            for edge in [
                RECT {
                    left: front_left,
                    top: front_top,
                    right: front_right,
                    bottom: front_top + thick,
                },
                RECT {
                    left: front_left,
                    top: front_bottom - thick,
                    right: front_right,
                    bottom: front_bottom,
                },
                RECT {
                    left: front_left,
                    top: front_top,
                    right: front_left + thick,
                    bottom: front_bottom,
                },
                RECT {
                    left: front_right - thick,
                    top: front_top,
                    right: front_right,
                    bottom: front_bottom,
                },
            ] {
                fill_rect(surface, edge, ink);
            }

            // The window behind: its top edge across the whole mark, and its right edge down to
            // where the window in front takes the corner — the two edges the front one does not
            // cover, a whole second outline being one thicker square rather than two windows.
            fill_rect(
                surface,
                RECT {
                    left: left + step,
                    top,
                    right,
                    bottom: (top + thick).min(bottom),
                },
                ink,
            );
            fill_rect(
                surface,
                RECT {
                    left: (right - thick).max(left),
                    top: top + thick,
                    right,
                    bottom: bottom - step,
                },
                ink,
            );
        }
        CardControl::Minimize => {
            stroke_segment(
                card_pixels(surface),
                width,
                (center_x - inset, center_y),
                (center_x + inset, center_y),
                stroke,
                ink,
                1.0,
            );
        }
        CardControl::Close => {
            // From the pixel the box's own first row and column hold clear of the
            // corner, to the last one that is clear of the far corner: the ends of a
            // stroke sit on the middle of a pixel, like the centre above.
            let from_x = rect.left as f32 + inset + 0.5;
            let from_y = rect.top as f32 + inset + 0.5;
            let to_x = rect.right as f32 - inset - 1.0 + 0.5;
            let to_y = rect.bottom as f32 - inset - 1.0 + 0.5;
            stroke_segment(
                card_pixels(surface),
                width,
                (from_x, from_y),
                (to_x, to_y),
                stroke,
                ink,
                1.0,
            );
            stroke_segment(
                card_pixels(surface),
                width,
                (from_x, to_y),
                (to_x, from_y),
                stroke,
                ink,
                1.0,
            );
        }
        // The row's controls are drawn by the caption's own glyph machinery, at a
        // size a button this small cannot hold (see `paint_card_control`).
        _ => {}
    }
}

// ----------------------------------------------------------------------- text

/// The width of a run of text, counted the way the rest of the page is: the face is fixed
/// pitch, so a count of characters is the whole measurement.
pub(super) fn text_width(text: &str, advance: i32) -> i32 {
    text.chars().count() as i32 * advance.max(1)
}

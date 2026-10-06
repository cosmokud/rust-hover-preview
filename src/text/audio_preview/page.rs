//! The page a card is laid out and painted from: the runs of text on it, the bar's row and the
//! band a press on it is answered against, the boxes of the controls a pinned card carries, and
//! the paint that fills them.
//!
//! None of it is kept: `measure` and `render` build the same page from the same card and hold
//! neither, so the row, the boxes and the paint are one set of arithmetic read three ways — which
//! is the whole of why a bar a person aims at and a bar drawn under a name are the same line.

use super::card::{
    clock, Card, CardChrome, CardControl, Fact, FactKind, BAR_GAP_PIXELS, BAR_PIXELS,
    BAR_REACH_PIXELS, BULLET, BULLET_CELL_ADVANCES, CHASE_DIVISOR, CHASE_SECONDS, CLOCK_ADVANCES,
    CONTROL_GAP_PIXELS, CONTROL_SIDE_PIXELS, HEADER_LEVEL, MIN_CONTENT_ADVANCES, RULE_GAP_PIXELS,
    RULE_PIXELS, TIME_GAP_ADVANCES, TRACK_PIXELS, WINDOW_BUTTON_CUSHION_PIXELS,
    WINDOW_BUTTON_GAP_PIXELS,
};
use crate::text::pin_chrome::{self, stroke_segment, surface_pixels, ControlGlyph};
use crate::text::text_paint::{
    blend, fill_rect, plain_style, readable, rgb, scaled, DibSurface, RunPainter, TextMetrics,
    TextStyle, BODY_LEVEL,
};
use crate::text::text_theme::LoadedTheme;
use windows::Win32::Foundation::RECT;

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
/// against its right edge, and the two window buttons standing in the top margin.
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
    /// The minimize that shrinks the pin into its bubble, and the close that ends it.
    pub(super) minimize: WindowButtonBox,
    pub(super) close: WindowButtonBox,
}

/// One of the two window buttons a pinned card carries in its top margin, with the
/// two boxes it has: the one its glyph is drawn in, and the one a press on it is
/// answered against.
///
/// The two are not the same box, and that is the whole of the shape: the button is
/// very small — the margin's own height less the room around it — so a hand is
/// given the window's own top edge above the drawn box and a sliver of the name
/// line's invisible leading below it (see [`window_button_boxes`]).
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
    /// and the two window buttons', whose is the box their glyphs are drawn in plus the cushion
    /// above them (see [`WindowButtonBox`]).
    pub(super) fn rect(&self, control: CardControl) -> Option<RECT> {
        Some(match control {
            CardControl::Previous => self.previous,
            CardControl::Play => self.play,
            CardControl::Next => self.next,
            CardControl::Seek => self.bar,
            CardControl::Volume => self.volume,
            CardControl::Minimize => self.minimize.hit,
            CardControl::Close => self.close.hit,
        })
    }

    /// Which of the two window buttons in the top margin a point is on, if either.
    ///
    /// Asked of the cushion each is answered against rather than of the drawn boxes,
    /// because a hand aiming at a button this small is given the room above it
    /// (see [`window_button_boxes`]).
    pub(super) fn window_button_at(&self, x: i32, y: i32) -> Option<CardControl> {
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
    let (minimize, close) = window_button_boxes(metrics, width as i32);

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
        minimize,
        close,
    })
}

/// The two window buttons a pinned card carries in its top margin: the minimize that shrinks the
/// pin into the bubble it leaves, and the close that ends the pin, the player and the window
/// together.
///
/// The card's top margin is the band the name line begins below, so each button is a square the
/// margin's own height less the gap above it and the gap below it, stood against the card's right
/// border with the same gap on that side and between the two of them. They are carved out of
/// margin the card already has — the bargain the bar-row controls make as well — so a card that
/// carries them is the card a hover's is and the name keeps the whole width of the card to
/// scroll across (see `build_page`).
///
/// A button's *hit* box is its drawn box widened upward to the window's own top edge and by the
/// cushion below it: the two buttons are very small, so a hand is given the room above them and
/// a pixel of the name line's invisible leading — the one band of the line the name's glyphs do
/// not sit in. The leading is read off the font the name is set in, and where a display scale
/// yields a font with none of it the cushion stops at the line's own top instead, so a hit box
/// never reaches the glyphs (see [`TextMetrics::internal_leading`]).
pub(super) fn window_button_boxes(
    metrics: &TextMetrics,
    width: i32,
) -> (WindowButtonBox, WindowButtonBox) {
    let gap = scaled(WINDOW_BUTTON_GAP_PIXELS, metrics.scale);
    // The margin's own height less the gap above and the gap below, which is what
    // fixes the side: a button is never taller than the band it stands in.
    let side = (metrics.padding - gap * 2).max(0);

    // The close against the right border, and the minimize beside it with the same
    // gap between them as each has to the border.
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

    // The cushion: up to the window's own top edge, and down by the gap plus a
    // pixel of the name line's invisible leading — clamped to the line's own top
    // where the font at this scale has no leading to borrow. A logical pixel like
    // the gap it is counted from, so the cushion is the same share of the button
    // at every scale (see `WINDOW_BUTTON_CUSHION_PIXELS`).
    let leading = metrics.internal_leading[HEADER_LEVEL as usize];
    let cushion = scaled(WINDOW_BUTTON_CUSHION_PIXELS, metrics.scale);
    let bottom = (minimize.bottom + cushion).min(metrics.padding + leading);

    (
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

pub(super) fn build_page(
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

/// The readout at the right edge of the facts line: what has been played in the accent, and
/// the whole it is measured against in the page's own muted gray. A card with no player behind
/// it has only the whole to show, and one whose file does not say how long it is has only the
/// number that moves.
pub(super) fn clock_runs(
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

pub(super) fn paint(surface: &DibSurface, page: &Page, theme: &LoadedTheme, scale: f32) {
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

        // The two window buttons in the top margin, painted last of all and painted as
        // nothing but their marks: no wash under them and no box around them in any
        // state, the ink being the only thing that changes. At rest it is the card's
        // own foreground, which is the ink the row's glyphs are painted in, and while
        // the pointer is on one or holding it it is the accent, which is the color the
        // played part of the bar is drawn in (see `build_page`).
        for (control, button) in [
            (CardControl::Minimize, &boxes.minimize),
            (CardControl::Close, &boxes.close),
        ] {
            let lit = chrome.hovered == Some(control) || chrome.pressed == Some(control);
            paint_window_button(
                surface,
                button.drawn,
                control,
                if lit { accent } else { foreground },
                scale,
            );
        }
    }
}

/// One of the two window buttons, drawn as the one mark that says what it is: a dash for the
/// minimize, two crossing strokes for the close.
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
    let buffer = unsafe {
        std::slice::from_raw_parts_mut(
            surface_pixels(surface),
            surface.width as usize * surface.height as usize * 4,
        )
    };
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
        CardControl::Minimize => {
            stroke_segment(
                buffer,
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
            stroke_segment(buffer, width, (from_x, from_y), (to_x, to_y), stroke, ink, 1.0);
            stroke_segment(buffer, width, (from_x, to_y), (to_x, from_y), stroke, ink, 1.0);
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

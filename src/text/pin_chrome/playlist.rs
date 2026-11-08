//! The playlist a pinned sound's card grows under it: the rows of the pin's own walk, drawn in
//! the window the card is in and in the same palette, as a band below the card rather than as a
//! window of their own.
//!
//! It is a band and not a panel because the card above it does not move for it: the window is
//! measured as the card plus this band (see `pinned_window_box_of`), and what a row costs is rows
//! of the window the card never reaches. The band is painted opaque — the theme's own page at full
//! alpha — because the pin is a layered window hit-tested by the shape of its pixels, and a band
//! left see-through is a press that falls to the listing behind it (see `pin_real_window`).
//!
//! Everything here is band-local: a box the band was given, the rows that fit in it, and where a
//! point of it lands. What the rows *say* is built once per order by `pin_playlist`, and what the
//! band's own row is on the screen is the caller's business, so this file never asks a pin
//! anything.

use super::caption::fit_title;
use super::primitives::{fill_box, ChromePalette};
use crate::readers::audio_track::{self, Probed};
use crate::text::audio_preview;
use crate::text::text_paint::{self, DibSurface};
use std::path::Path;
use windows::Win32::Foundation::RECT;

/// What one row is drawn as: the two lines at its left and the length at its right.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct PlaylistRow {
    /// The file's own name without its extension: what a person reading the band is
    /// identifying, the folder being the one the pin was taken up in already.
    pub(crate) name: String,
    /// What the file says about itself, read once when the rows are built and drawn the way the
    /// card draws its own facts line: the format in the numbers' ink and everything else the file
    /// says in the muted one, parted by the same ` · ` in that same muted gray (see
    /// `audio_preview::facts_of` and `audio_preview::page`).
    pub(crate) facts: Vec<audio_preview::Fact>,
    /// How long it plays, or nothing where the file does not say — a row's right column is a
    /// clock rather than a dash, so a lengthless file carries nothing there.
    pub(crate) clock: Option<String>,
}

/// How the rows of a band are laid out: the box they scroll in, how many of them fit, and the
/// four edges the two lines and the clock are drawn against.
///
/// It is band-local, so a row's own box is one multiply away and a press is answered against the
/// same numbers the paint drew from rather than against a second arithmetic.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct PlaylistLayout {
    pub(crate) width: i32,
    pub(crate) row_height: i32,
    /// The rows the band can draw, from the first row that fits to the last.
    pub(crate) rows_top: i32,
    pub(crate) rows_bottom: i32,
    pub(crate) visible: i32,
    /// Where a row's name begins: past the pad, the stripe the shown row carries and a gap.
    pub(crate) text_left: i32,
    /// The last column a row's name may reach: the length is right-aligned in its column and drawn
    /// narrower than it, so the name is held off the drawn text's own left edge — at the widest
    /// shape a length is ever drawn as, which is one cut for every row whatever its clock says.
    pub(crate) name_right: i32,
    /// The last column a row's facts line may reach: the clock's whole column taken off the width,
    /// which is the edge the note under the name has always kept.
    pub(crate) facts_right: i32,
    pub(crate) clock_left: i32,
    pub(crate) clock_right: i32,
    /// The scrollbar's own track, left and right, whenever the rows do not all fit.
    pub(crate) track: Option<(i32, i32)>,
}

/// The rows a paint is handed, and the three a band draws one of them differently for.
pub(crate) struct PlaylistBand<'a> {
    pub(crate) layout: PlaylistLayout,
    /// The rows that fit, the first of them being the row number `first`.
    pub(crate) rows: &'a [PlaylistRow],
    pub(crate) first: usize,
    /// How many rows the whole list holds, which is what a scrollbar is measured against.
    pub(crate) total: usize,
    /// The row the pin is showing, which carries the accent stripe at its left edge.
    pub(crate) showing: Option<usize>,
    /// The row the hand is on, and the row the hand has picked: either one marquees a name the
    /// band has no room for (see `pin_playlist`).
    pub(crate) hovered: Option<usize>,
    pub(crate) picked: Option<usize>,
    /// How far the marqueed name has been scrolled, in pixels to the left.
    pub(crate) name_offset: i32,
}

/// How tall a row is, in the display's own pixels: two lines of text and the pad above and below
/// them.
pub(crate) fn playlist_row_height(scale: f32) -> i32 {
    text_paint::scaled(ROW_PIXELS as i32, scale).max(1)
}

/// The box a row's first line is set in, and its second: the two are centred in the row together,
/// so a band at any scale keeps the pair on the middle of its row rather than on its top.
fn line_boxes(layout: &PlaylistLayout, row_top: i32, scale: f32) -> (RECT, RECT) {
    let name_cell = text_paint::scaled(text_paint::LEVEL_FONT_PIXELS[NAME_LEVEL as usize], scale);
    let facts_cell = text_paint::scaled(FACTS_PIXELS as i32, scale);
    let top = row_top + ((layout.row_height - name_cell - facts_cell) / 2).max(0);

    (
        RECT {
            left: layout.text_left,
            top,
            right: layout.name_right.max(layout.text_left),
            bottom: top + name_cell,
        },
        RECT {
            left: layout.text_left,
            top: top + name_cell,
            right: layout.facts_right.max(layout.text_left),
            bottom: top + name_cell + facts_cell,
        },
    )
}

/// The band a box leaves its rows: the pad taken off both ends, the rows whole ones that fit, and
/// the two columns a row is drawn in — with the scrollbar's own room taken off the right of the
/// width whenever there is anything to scroll.
pub(crate) fn playlist_layout(width: i32, height: i32, scale: f32, rows: usize) -> PlaylistLayout {
    let pad = text_paint::scaled(PAD_PIXELS as i32, scale).max(0);
    let row_height = playlist_row_height(scale);
    let stripe = text_paint::scaled(STRIPE_PIXELS as i32, scale).max(0);
    let bar = text_paint::scaled(SCROLLBAR_PIXELS as i32, scale).max(1);
    let gap = text_paint::scaled(SCROLLBAR_GAP_PIXELS as i32, scale).max(0);
    let track_margin = text_paint::scaled(TRACK_MARGIN_PIXELS as i32, scale).max(0);
    let clock = text_paint::scaled(CLOCK_PIXELS as i32, scale).max(0);
    let clock_gap = text_paint::scaled(CLOCK_GAP_PIXELS as i32, scale).max(0);
    let name_gap = text_paint::scaled(NAME_GAP_PIXELS as i32, scale).max(0);

    let rows_top = pad.min(height);
    let rows_bottom = (height - pad).max(rows_top);
    let visible = ((rows_bottom - rows_top) / row_height).max(0);
    let scrolling = rows.saturating_sub(visible as usize) > 0;
    // The scrollbar is a mark on the window's own edge rather than a column of the band's: it is
    // held the margin off the edge and the rows are held the gap off it, so what the rows give up
    // for it is its own width and a breath beside it and no more.
    let track_right = (width - track_margin).max(0);
    let track_left = (track_right - bar).max(0);
    // What the rows give up on their right: the pad the band holds everything off its edge by, or
    // the bar's own column and the breath beside it where there is one — the bar stands on the
    // window's edge rather than inside the pad, so the two are counted from different ends.
    let margin = if scrolling {
        width - track_left + gap
    } else {
        pad
    };

    let text_left = pad + stripe + gap;
    let clock_right = (width - margin).max(text_left);
    let clock_left = (clock_right - clock).max(text_left);
    // The name is cut against the length as it is drawn rather than against the column holding it:
    // the column is wider than its own text, so a cut on the column's left edge leaves a gap no
    // clock asked for. Nothing to measure on leaves it there, which overlaps no text either way.
    let drawn_left = widest_clock(scale).map_or(clock_left, |widest| {
        clock_origin(clock_right, clock_left, widest)
    });

    PlaylistLayout {
        width,
        row_height,
        rows_top,
        rows_bottom,
        visible,
        text_left,
        name_right: (drawn_left - name_gap).max(text_left),
        facts_right: (clock_left - clock_gap).max(text_left),
        clock_left,
        clock_right,
        track: scrolling.then_some((track_left, track_right)),
    }
}

/// The last row the band can scroll to: every row that does not fit, and none where they all do.
pub(crate) fn playlist_scroll_max(rows: usize, visible: usize) -> usize {
    rows.saturating_sub(visible)
}

/// Which edge of the window a band-local row is the height's own handle on.
///
/// The band is the only thing a pinned sound's window is ever resized by — the card above it is
/// drawn at the share it was taken up at and is never pulled — so the band's two edges are the
/// window's two edges, and neither strip is one a press picks a row with (see `pin_playlist::drag`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum PlaylistEdge {
    /// The window's own first rows: dragging them up grows the band and carries the card up with
    /// them, the window's bottom edge staying exactly where it was.
    Top,
    /// The window's own last rows: dragging them down grows the band, and nothing above moves.
    Bottom,
}

/// Which of the band's two edges a window row is on, or nothing where it is the band's own body.
///
/// `top` is the row the band begins at in the window, which is what tells the window's own first
/// rows from the band's: the top edge is the window's, and a press on it is a press above the band.
pub(crate) fn playlist_edge_at(top: i32, height: i32, scale: f32, y: i32) -> Option<PlaylistEdge> {
    let edge = text_paint::scaled(EDGE_PIXELS as i32, scale).max(1);
    let bottom = top + height;

    if (0..edge).contains(&y) {
        return Some(PlaylistEdge::Top);
    }
    if (bottom - edge..bottom).contains(&y) {
        return Some(PlaylistEdge::Bottom);
    }

    None
}

/// The row a band-local `y` is on, or nothing where that row is the pad's or the band has run out
/// of rows before it.
pub(crate) fn playlist_row_at(
    layout: &PlaylistLayout,
    scroll: usize,
    rows: usize,
    y: i32,
) -> Option<usize> {
    if y < layout.rows_top || y >= layout.rows_bottom || layout.row_height <= 0 {
        return None;
    }

    let step = ((y - layout.rows_top) / layout.row_height) as usize;
    let index = scroll + step;

    (step < layout.visible.max(0) as usize && index < rows).then_some(index)
}

/// The row the arrows step to, and the scroll that keeps it in sight.
///
/// The first press with nothing picked starts on the row the pin is showing where its walk has one
/// and on the band's first row in sight where it has none. Either end holds rather than wrapping,
/// and the scroll is a whole row because a band never draws half of one.
pub(crate) fn playlist_step_pick(
    picked: Option<usize>,
    showing: Option<usize>,
    scroll: usize,
    rows: usize,
    visible: usize,
    direction: i32,
) -> Option<(usize, usize)> {
    let last = rows.checked_sub(1)?;
    let from = picked.or(showing).unwrap_or(scroll).min(last);
    let picked = (from as i32 + direction).clamp(0, last as i32) as usize;
    let scroll = scroll
        .min(picked)
        .max((picked + 1).saturating_sub(visible.max(1)));

    Some((picked, scroll))
}

/// The scrollbar's thumb, top and bottom: the share of the rows the band shows, carried down the
/// track by the share it has scrolled.
pub(crate) fn playlist_thumb(layout: &PlaylistLayout, scroll: usize, rows: usize) -> (i32, i32) {
    let room = (layout.rows_bottom - layout.rows_top).max(1);
    let visible = layout.visible.max(1) as usize;
    let total = rows.max(visible);

    let height = ((room as i64 * visible as i64) / total as i64) as i32;
    let height = height.max((layout.row_height / 2).max(2)).min(room);
    let travel = (room - height).max(0);
    let max = playlist_scroll_max(rows, visible).max(1);
    let top = layout.rows_top + (travel * scroll.min(max) as i32) / max as i32;

    (top, top + height)
}

/// The scroll a thumb dragged to `top` means: the inverse of `playlist_thumb`, so that a press on
/// the bar and the bar's own drawing are one arithmetic rather than two that drift.
///
/// Nothing at all for a band whose rows all fit, which is a band with no bar to drag.
pub(crate) fn playlist_scroll_at(layout: &PlaylistLayout, rows: usize, top: i32) -> usize {
    let room = (layout.rows_bottom - layout.rows_top).max(1);
    let visible = layout.visible.max(1) as usize;
    let max = playlist_scroll_max(rows, visible);
    let (thumb_top, thumb_bottom) = playlist_thumb(layout, 0, rows);
    let travel = room - (thumb_bottom - thumb_top);

    if travel <= 0 || max == 0 {
        return 0;
    }

    let carried = (top - layout.rows_top).clamp(0, travel);
    ((carried as i64 * max as i64 + (travel as i64 / 2)) / travel as i64) as usize
}

/// The share of the thumb a press on the track has hold of, so that the bar does not jump under the
/// hand that is carrying it: where on the thumb the press landed, and the middle of it for a press
/// on the track itself — which is the one gesture a bar answers by jumping.
pub(crate) fn playlist_thumb_grab(
    layout: &PlaylistLayout,
    scroll: usize,
    rows: usize,
    y: i32,
) -> i32 {
    let (thumb_top, thumb_bottom) = playlist_thumb(layout, scroll, rows);

    if (thumb_top..thumb_bottom).contains(&y) {
        return y - thumb_top;
    }

    (thumb_bottom - thumb_top) / 2
}

/// Whether a band-local point is the scrollbar's own track rather than a row's.
pub(crate) fn playlist_track_holds(layout: &PlaylistLayout, x: i32, y: i32) -> bool {
    let Some((left, right)) = layout.track else {
        return false;
    };

    x >= left && x < right && y >= layout.rows_top && y < layout.rows_bottom
}

/// How far a name has to travel before its end is in sight, at the room a row's left column
/// leaves, and the advance of the font it is set in: the two numbers a marquee is made of, worked
/// out once when a row becomes the one the hand is on rather than on every frame it moves.
///
/// Nothing for a name the column has room for, which is a name drawn once and held.
pub(crate) fn name_travel(name: &str, room: i32, scale: f32) -> Option<(i32, i32)> {
    if room <= 0 || name.is_empty() {
        return None;
    }

    let surface = DibSurface::create(1, 1)?;
    let total = measure(&surface, &name_style(), name, scale);
    let travel = total - room;
    if travel <= 0 {
        return None;
    }

    let characters = name.chars().count().max(1) as i32;

    Some((travel, (total / characters).max(1)))
}

/// The band's own pixels: the page its rows sit on, the wash and the stripe a row carries, the
/// scrollbar, and the two lines of every row.
///
/// It is a surface of the band's own size and it is painted *whole*, so what the caller does with
/// it is carry its rows across: a band is opaque by design, and there is nothing of the window
/// under it left to composite through.
pub(crate) fn paint_playlist(
    surface: &DibSurface,
    palette: &ChromePalette,
    band: &PlaylistBand,
    scale: f32,
) {
    let width = surface.width as i32;
    let height = surface.height as i32;
    if width <= 0 || height <= 0 {
        return;
    }

    let layout = band.layout;
    let pixels = unsafe {
        std::slice::from_raw_parts_mut(
            surface.bits(),
            surface.width as usize * surface.height as usize * 4,
        )
    };

    let page = palette.background;
    let wash = palette.hover(ROW_WASH);
    let whole = RECT {
        left: 0,
        top: 0,
        right: width,
        bottom: height,
    };
    fill_box(pixels, width, whole, page, 1.0);

    // The hairline the card above ends on, so that a row scrolling up to the seam has something
    // to disappear under rather than a row of the card's own ink to run into.
    fill_box(
        pixels,
        width,
        RECT {
            left: 0,
            top: 0,
            right: width,
            bottom: (scale.round() as i32).max(1),
        },
        palette.hover(HAIRLINE),
        1.0,
    );

    let stripe = text_paint::scaled(STRIPE_PIXELS as i32, scale).max(0);
    let name_style = name_style();
    let facts_style = facts_style();

    // **The facts line is set the way the card sets its own**, which is two inks rather than one:
    // the format is the numbers' ink, so it is the one part of the line a person's eye lands on,
    // and everything the file says beside it — the rate, how many channels there are, the bitrate,
    // the room it takes — is the muted gray. The ` · ` between two of them is that same gray, so
    // the line reads as parts of one sentence rather than as a run of one colour (see
    // `audio_preview::page`).
    let muted = text_paint::blend(palette.background, palette.foreground, MUTED_INK);
    let separator = " · ";

    // The runs are drawn through GDI, which writes colour and leaves the alpha byte at zero (see
    // the module documentation of `pin_chrome`), so each run's own box is given its coverage back
    // once it has been drawn — the same sealing the caption does for its title.
    let mut boxes = Vec::with_capacity(band.rows.len() * 6);
    let mut painter = text_paint::RunPainter::new(surface, scale);
    // The second line is a pixel under the first, so it is measured and drawn at the scale that
    // lands it there rather than at the band's own (see `facts_scale`).
    let small = facts_scale(scale);
    let mut small_painter = text_paint::RunPainter::new(surface, small);

    for (step, row) in band.rows.iter().enumerate() {
        let index = band.first + step;
        let top = layout.rows_top + step as i32 * layout.row_height;
        let bottom = (top + layout.row_height).min(layout.rows_bottom);
        if bottom <= top {
            continue;
        }

        let picked = band.picked == Some(index);
        let under = if picked { wash } else { page };
        let band_rect = RECT {
            left: 0,
            top,
            right: width,
            bottom,
        };
        if picked {
            fill_box(pixels, width, band_rect, under, 1.0);
        }

        if band.showing == Some(index) && stripe > 0 {
            fill_box(
                pixels,
                width,
                RECT {
                    left: 0,
                    top,
                    right: stripe,
                    bottom,
                },
                palette.accent,
                1.0,
            );
        }

        let (name_box, facts_box) = line_boxes(&layout, top, scale);

        // A name the left column has room for is drawn once, cut with an ellipsis where it has
        // none — while the row the hand is on travels instead, which is the whole of what picking
        // one buys a name too long to read.
        let marqueed = band.hovered == Some(index) || picked;
        let shown_name = if marqueed {
            row.name.clone()
        } else {
            fit_title(
                surface,
                &name_style,
                &row.name,
                name_box.right - name_box.left,
                scale,
            )
        };
        let origin = if marqueed {
            name_box.left - band.name_offset
        } else {
            name_box.left
        };

        painter.draw(
            &shown_name,
            origin,
            name_box,
            &name_style,
            palette.foreground,
            under,
        );
        boxes.push(name_box);

        // The parts of the facts line, one after another, each in its own ink and each parted from
        // the one before it by the muted gray — and the whole of it cut where the column ends, which
        // is what a line of parts rather than one string is drawn by.
        let mut x = facts_box.left;
        for (at, fact) in row.facts.iter().enumerate() {
            let ink = match fact.kind {
                audio_preview::FactKind::Format => palette.number,
                audio_preview::FactKind::Reading => palette.foreground,
                audio_preview::FactKind::Note => muted,
            };

            if at > 0 {
                let width = measure(surface, &facts_style, separator, small);
                if x + width > facts_box.right {
                    break;
                }
                small_painter.draw(
                    separator,
                    x,
                    RECT {
                        left: x,
                        right: x + width,
                        ..facts_box
                    },
                    &facts_style,
                    muted,
                    under,
                );
                boxes.push(RECT {
                    left: x,
                    right: x + width,
                    ..facts_box
                });
                x += width;
            }

            let room = facts_box.right - x;
            let text = fit_title(surface, &facts_style, &fact.text, room, small);
            let width = measure(surface, &facts_style, &text, small);
            if width == 0 {
                break;
            }
            let run = RECT {
                left: x,
                right: x + width,
                ..facts_box
            };
            small_painter.draw(&text, x, run, &facts_style, ink, under);
            boxes.push(run);
            x += width;

            if width >= room {
                break;
            }
        }
        boxes.push(facts_box);

        if let Some(clock) = row.clock.as_deref() {
            let measured = measure(surface, &facts_style, clock, scale);
            let clock_box = RECT {
                left: layout.clock_left,
                top: facts_box.top,
                right: layout.clock_right,
                bottom: facts_box.bottom,
            };
            if clock_box.right > clock_box.left {
                // The length is the accent, which is the ink the card's own clock and its bar are
                // drawn in: the two say the same thing about the same second (see
                // `audio_preview::page`).
                painter.draw(
                    clock,
                    clock_origin(clock_box.right, clock_box.left, measured),
                    clock_box,
                    &facts_style,
                    palette.accent,
                    under,
                );
                boxes.push(clock_box);
            }
        }

        // **A row's name, its facts and its length are two painters' work on one DC**, and each
        // remembers only the face it selected itself: neither may trust its cache once the row is
        // done, or the next row's run comes out in the face the other left there.
        painter.forget_the_face();
        small_painter.forget_the_face();
    }

    // The scrollbar goes last, over the rows' right edge and out of their columns' way: the track
    // it runs in is the room the rows were held off the window's margin for.
    if let Some((left, right)) = layout.track {
        let track = RECT {
            left,
            top: layout.rows_top,
            right,
            bottom: layout.rows_bottom,
        };
        fill_box(
            pixels,
            width,
            track,
            text_paint::blend(page, palette.foreground, TRACK_INK),
            1.0,
        );

        // The thumb is the caller's own scroll, which is the row the band was handed first.
        let (thumb_top, thumb_bottom) = playlist_thumb(&layout, band.first, band.total);
        fill_box(
            pixels,
            width,
            RECT {
                left,
                top: thumb_top,
                right,
                bottom: thumb_bottom,
            },
            text_paint::blend(page, palette.foreground, THUMB_INK),
            1.0,
        );
    }

    // GDI left the alpha byte of every run's box at zero, and the band is opaque: what it drew is
    // given its own coverage back before the band is carried across.
    for box_ in &boxes {
        for y in box_.top.max(0)..box_.bottom.min(height) {
            for x in box_.left.max(0)..box_.right.min(width) {
                pixels[(y as usize * width as usize + x as usize) * 4 + 3] = 255;
            }
        }
    }
}

/// One run, measured at the scale it is drawn at, without a box to draw it in: what a name's
/// travel and a clock's right-hand edge are worked out from.
fn measure(surface: &DibSurface, style: &text_paint::TextStyle, text: &str, scale: f32) -> i32 {
    super::primitives::measure_text(surface, style, text, scale)
}

/// Where a length is drawn from: its column is a fixed room and the text is set off the right of it,
/// so the column's own left edge is the furthest a short clock can start and the widest one starts
/// there.
pub(crate) fn clock_origin(right: i32, left: i32, measured: i32) -> i32 {
    (right - measured).max(left)
}

/// The width the widest length is drawn at, at a band's own scale: what the name column is cut
/// against (see `CLOCK_REFERENCE`). Nothing where there is no DC to measure on.
fn widest_clock(scale: f32) -> Option<i32> {
    let surface = DibSurface::create(1, 1)?;

    Some(measure(&surface, &facts_style(), CLOCK_REFERENCE, scale))
}

/// The style a row's name is set in: the card's own header level, on the face a caption is set in
/// — a row is chrome under a card, and the two are one window.
fn name_style() -> text_paint::TextStyle {
    styled(NAME_LEVEL)
}

/// The style a row's facts and its clock are set in: the theme's own body text, at the size one
/// pixel under it — a facts line is the note under the name rather than a second name, and a pixel
/// is the whole of the difference a row needs.
pub(crate) fn facts_style() -> text_paint::TextStyle {
    styled(FACTS_LEVEL)
}

/// The scale the facts line is set at, which is the band's own times the ratio that lands it a pixel
/// under the body level.
///
/// A size is a level of the page's own table and the step under the body is three pixels, so the one
/// pixel is asked for as a scale: a font is made at the size its style names times the scale, so the
/// scale that lands `FACTS_PIXELS` where `LEVEL_FONT_PIXELS[FACTS_LEVEL]` would have stood is the
/// display's own times the ratio of the two.
fn facts_scale(scale: f32) -> f32 {
    let body = text_paint::LEVEL_FONT_PIXELS[FACTS_LEVEL as usize] as f32;

    scale * (FACTS_PIXELS / body)
}

fn styled(level: u8) -> text_paint::TextStyle {
    text_paint::TextStyle {
        foreground: [0, 0, 0],
        background: None,
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        level,
        face: ROW_FACE,
    }
}

/// The row a band draws a file as: its name without the extension, the parts the file says about
/// itself, and how long it plays.
///
/// Every part is left out rather than guessed at, which is the card's own rule for its facts line
/// (see `audio_preview::facts_of`): a file nothing has probed is a file whose only fact is the
/// extension it carries and the room it takes. What it costs is a probe and a stat per row, which
/// is why the rows are built when the walk moves and never at a paint.
pub(crate) fn playlist_row(path: &Path) -> PlaylistRow {
    let track = match audio_track::probed(path) {
        Probed::Track(track) => Some(track),
        _ => None,
    };

    let facts = match track.as_ref() {
        Some(track) => audio_preview::facts_of(track, path),
        // A file nothing probed says only be read to say without one: the extension it carries
        // and the room it takes, which is the one fact no probe answers for.
        None => vec![
            audio_preview::Fact {
                text: audio_preview::extension_of(path).to_uppercase(),
                kind: audio_preview::FactKind::Format,
            },
            audio_preview::Fact {
                text: audio_preview::size_label(path),
                kind: audio_preview::FactKind::Note,
            },
        ],
    };

    PlaylistRow {
        name: path
            .file_stem()
            .map(|stem| stem.to_string_lossy().to_string())
            .unwrap_or_default(),
        facts,
        clock: track
            .as_ref()
            .and_then(|track| track.duration)
            .map(audio_preview::clock),
    }
}

/// How tall a row is in logical pixels: two lines of text and the pad above and below them.
const ROW_PIXELS: f32 = 44.0;

/// The pad the rows are held off every edge of the band by, and the room the drag handle at the
/// band's bottom edge is drawn in — the pad is the wider of the two on purpose, so that no press
/// on a row is ever a press on the handle.
const PAD_PIXELS: f32 = 10.0;

/// The room the accent stripe of the row the pin is showing takes at a row's left edge.
const STRIPE_PIXELS: f32 = 3.0;

/// The scrollbar in the band's right margin: how wide its track is, the gap that holds the rows
/// off it, and how far a row's name may run before the scrollbar's own room is taken off.
const SCROLLBAR_PIXELS: f32 = 5.0;
const SCROLLBAR_GAP_PIXELS: f32 = 6.0;
/// How far the scrollbar's track is held off the window's own right edge: the bar is a mark on that
/// edge, and a margin wider than a couple of pixels reads as a column the band does not have.
const TRACK_MARGIN_PIXELS: f32 = 2.0;

/// The clock's column: the room `1:02:03` is set in at the body level, and the gap that holds the
/// facts line off it.
const CLOCK_PIXELS: f32 = 56.0;
const CLOCK_GAP_PIXELS: f32 = 8.0;

/// The widest length a row's clock is ever drawn as, which is what the name column is cut against:
/// six zeroes are the face's widest figures, so no clock of any length or any hour runs into it.
const CLOCK_REFERENCE: &str = "0:00:00";

/// How far the name column is held off the left of the drawn length: the column is wider than the
/// text it holds, so a name held off the column's edge stops short of the ink.
const NAME_GAP_PIXELS: f32 = 5.0;

/// The face a row is set in: the one Windows writes a window's title in, which is the caption's
/// own, so a row of a band and the strip above the card read as one window.
const ROW_FACE: &str = "Segoe UI";

/// The level the card's own name is set at, and the body level under it — the card's two, so a
/// row is a card's line of the same two sizes.
const NAME_LEVEL: u8 = 3;
const FACTS_LEVEL: u8 = text_paint::BODY_LEVEL;

/// The size a row's facts line is set at: the body level less the one pixel that makes it the note
/// under the name (see `facts_scale`).
const FACTS_PIXELS: f32 = text_paint::LEVEL_FONT_PIXELS[FACTS_LEVEL as usize] as f32 - 1.0;

/// How far the wash a picked row takes lifts off the band's page, which is the lift a caption
/// button takes under the pointer: a row is a thing a hand has chosen, not a control lit up.
const ROW_WASH: f32 = 0.10;

/// How far the muted ink a facts line's words and its separators are drawn in stands off the page,
/// which is the amount the card's own facts line is muted by (see `audio_preview::page`).
const MUTED_INK: f32 = 0.45;

/// How far the seam at the band's top is drawn off the page, and the scrollbar's two inks: the
/// same two amounts a page of text draws its own bar in.
const HAIRLINE: f32 = 0.18;
const TRACK_INK: f32 = 0.10;
const THUMB_INK: f32 = 0.45;

/// How many rows a notch of the wheel steps the band by: whole rows, so that a notch never leaves
/// a row of a line at the top of the band.
pub(crate) const WHEEL_ROWS: i32 = 3;

/// How close to either of the window's edges a press is the band's height being dragged rather
/// than a row being picked — the two strips of a window that a press picks nothing with.
pub(crate) const EDGE_PIXELS: f32 = 6.0;

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
    /// What the file says about itself, joined with ` · ` and read once when the rows are built:
    /// the format, the sample rate, the channels, the bitrate and the room it takes.
    pub(crate) facts: String,
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
    /// The last column the left half of a row may reach: the clock's own room is taken off the
    /// width, so a name never runs under the length.
    pub(crate) name_right: i32,
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
    let facts_cell = text_paint::scaled(text_paint::LEVEL_FONT_PIXELS[FACTS_LEVEL as usize], scale);
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
            right: layout.name_right.max(layout.text_left),
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
    let clock = text_paint::scaled(CLOCK_PIXELS as i32, scale).max(0);
    let clock_gap = text_paint::scaled(CLOCK_GAP_PIXELS as i32, scale).max(0);

    let rows_top = pad.min(height);
    let rows_bottom = (height - pad).max(rows_top);
    let visible = ((rows_bottom - rows_top) / row_height).max(0);
    let scrolling = rows.saturating_sub(visible as usize) > 0;
    let margin = if scrolling { bar + gap } else { 0 };

    let text_left = pad + stripe + gap;
    let clock_right = (width - pad - margin).max(text_left);
    let clock_left = (clock_right - clock).max(text_left);

    PlaylistLayout {
        width,
        row_height,
        rows_top,
        rows_bottom,
        visible,
        text_left,
        name_right: (clock_left - clock_gap).max(text_left),
        clock_left,
        clock_right,
        track: scrolling.then(|| {
            let right = (width - pad).max(0);
            (right - bar, right)
        }),
    }
}

/// The last row the band can scroll to: every row that does not fit, and none where they all do.
pub(crate) fn playlist_scroll_max(rows: usize, visible: usize) -> usize {
    rows.saturating_sub(visible)
}

/// Whether a band-local `y` is the height's own handle rather than a row's: the strip at the
/// band's bottom edge, which is the one part of a band a press does not pick anything with.
pub(crate) fn playlist_handle_holds(height: i32, scale: f32, y: i32) -> bool {
    let handle = text_paint::scaled(HANDLE_PIXELS as i32, scale).max(1);

    y >= height - handle && y < height
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

    // The runs are drawn through GDI, which writes colour and leaves the alpha byte at zero (see
    // the module documentation of `pin_chrome`), so each run's own box is given its coverage back
    // once it has been drawn — the same sealing the caption does for its title.
    let mut boxes = Vec::with_capacity(band.rows.len() * 3);
    let mut painter = text_paint::RunPainter::new(surface, scale);

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

        painter.draw(
            &fit_title(
                surface,
                &facts_style,
                &row.facts,
                facts_box.right - facts_box.left,
                scale,
            ),
            facts_box.left,
            facts_box,
            &facts_style,
            palette.number,
            under,
        );
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
                painter.draw(
                    clock,
                    (clock_box.right - measured).max(clock_box.left),
                    clock_box,
                    &facts_style,
                    palette.number,
                    under,
                );
                boxes.push(clock_box);
            }
        }
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

/// The style a row's name is set in: the card's own header level, on the face a caption is set in
/// — a row is chrome under a card, and the two are one window.
fn name_style() -> text_paint::TextStyle {
    styled(NAME_LEVEL)
}

/// The style a row's facts and its clock are set in: the theme's own body text, in the ink a
/// card reads its numbers in.
fn facts_style() -> text_paint::TextStyle {
    styled(FACTS_LEVEL)
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

    let mut parts: Vec<String> = match track.as_ref() {
        Some(track) => audio_preview::facts_of(track, path)
            .into_iter()
            .map(|fact| fact.text)
            .collect(),
        None => vec![audio_preview::extension_of(path).to_uppercase()],
    };
    // The room the file takes is the one part no probe answers for, and the one a person picking
    // one of ten mixes reads: it is always there.
    parts.push(format!(
        "{:.1} MB",
        size_in_bytes(path) as f64 / BYTES_IN_A_MEGABYTE
    ));

    PlaylistRow {
        name: path
            .file_stem()
            .map(|stem| stem.to_string_lossy().to_string())
            .unwrap_or_default(),
        facts: parts.join(" · "),
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
const SCROLLBAR_PIXELS: f32 = 4.0;
const SCROLLBAR_GAP_PIXELS: f32 = 6.0;

/// The clock's column: the room `1:02:03` is set in at the body level, and the gap that holds the
/// name column off it.
const CLOCK_PIXELS: f32 = 56.0;
const CLOCK_GAP_PIXELS: f32 = 8.0;

/// The face a row is set in: the one Windows writes a window's title in, which is the caption's
/// own, so a row of a band and the strip above the card read as one window.
const ROW_FACE: &str = "Segoe UI";

/// The level the card's own name is set at, and the body level under it — the card's two, so a
/// row is a card's line of the same two sizes.
const NAME_LEVEL: u8 = 3;
const FACTS_LEVEL: u8 = text_paint::BODY_LEVEL;

/// How far the wash a picked row takes lifts off the band's page, which is the lift a caption
/// button takes under the pointer: a row is a thing a hand has chosen, not a control lit up.
const ROW_WASH: f32 = 0.10;

/// How far the seam at the band's top is drawn off the page, and the scrollbar's two inks: the
/// same two amounts a page of text draws its own bar in.
const HAIRLINE: f32 = 0.18;
const TRACK_INK: f32 = 0.10;
const THUMB_INK: f32 = 0.45;

/// How many rows a notch of the wheel steps the band by: whole rows, so that a notch never leaves
/// a row of a line at the top of the band.
pub(crate) const WHEEL_ROWS: i32 = 3;

/// How close to the band's bottom edge a press is the height's handle rather than a row's — the
/// one strip of the band a press does not pick anything with.
pub(crate) const HANDLE_PIXELS: f32 = 6.0;

/// The room a file takes, as the one fact of a row no probe answers for.
fn size_in_bytes(path: &Path) -> u64 {
    std::fs::metadata(path).map(|meta| meta.len()).unwrap_or(0)
}

const BYTES_IN_A_MEGABYTE: f64 = 1048576.0;

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
use super::primitives::{
    draw_cross, fill_box, fill_ring, fill_round_rect, measure_run, stroke_round_rect,
    stroke_segment, ChromePalette, GLYPH_STROKE_PIXELS,
};
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
    /// The search field at the band's bottom: the rows end where it begins, and its own bottom is
    /// held the pad off the window's last row so that it never meets the drag strip there.
    ///
    /// It is a pill: the box is filled with the band's own page blended a little toward its ink and
    /// rounded to half its height, and the two columns inside it are the mark's and the query's.
    pub(crate) search_top: i32,
    pub(crate) search_bottom: i32,
    pub(crate) search_left: i32,
    pub(crate) search_right: i32,
    /// Where the search mark's square begins, inside the pill's own left pad.
    pub(crate) search_mark: i32,
    /// Where the query begins: past the pill's pad, the mark and the gap after it. A press names a
    /// place in the query from this column, so the mark is never a character of it.
    pub(crate) search_text_left: i32,
    /// Where the query ends: the pill's right pad taken off, and the gap the pill keeps before the
    /// clear button held off that again, so that a query never runs under the × that clears it.
    pub(crate) search_text_right: i32,
    /// The clear button's own square at the pill's right, as the columns it runs between: the × a
    /// press puts the query back to nothing with (see `search_close_holds`).
    pub(crate) search_close_left: i32,
    pub(crate) search_close_right: i32,
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
    /// The whole list the band was built from, and the positions into it the query lets through:
    /// the band draws `shown` and never the rows it hides, so a scroll, a scrollbar and a row's
    /// place are all counted in positions while `showing`, `picked` and the marquees stay rows.
    pub(crate) rows: &'a [PlaylistRow],
    pub(crate) shown: &'a [usize],
    /// The position the first row drawn is, which is what the scroll counts in.
    pub(crate) first: usize,
    /// How many positions the list holds, which is what a scrollbar is measured against.
    pub(crate) total: usize,
    /// The row the pin is showing, which carries the accent stripe at its left edge.
    pub(crate) showing: Option<usize>,
    /// The row the hand has picked, which is the row the band washes (see `pin_playlist`).
    pub(crate) picked: Option<usize>,
    /// The rows whose names are travelling, and how far each has been carried, in pixels to the
    /// left: the row the hand is on, the row it has picked and the row the pin is showing, any of
    /// which can be a name the column has no room for. A name on its way back to its resting place
    /// is still one of them (see `pin_playlist::set_the_marquee`).
    pub(crate) scrolled: &'a [(usize, i32)],
    /// The query field at the band's bottom, and the line a band whose query hides every row says
    /// instead of its rows.
    pub(crate) search: Option<PlaylistSearch<'a>>,
    pub(crate) note: Option<&'a str>,
}

/// The band's search field, as a paint reads it: the query, where its caret stands, and whether
/// the field is the one the keys are going to.
///
/// Whether the text is the placeholder is the text's own answer — an empty query is what draws
/// the placeholder — so there is nothing here to keep in step with it.
pub(crate) struct PlaylistSearch<'a> {
    pub(crate) text: &'a str,
    /// A byte index on a character boundary, as the field keeps it (see `text_find_field::Field`).
    pub(crate) caret: usize,
    /// The selected range of the query, in the text's own order, where the field holds one: the wash
    /// drawn between its two ends is what a press carried across the text or a `Ctrl`+`A` puts there,
    /// and the keys that cut text take the range rather than a character (see
    /// `text_find_field::Field`).
    pub(crate) selection: Option<(usize, usize)>,
    /// Whether the caret is drawn, and the bar with it.
    pub(crate) focused: bool,
    /// Whether the caret's blink is in its lit half, which is the half the bar is drawn in: a field that
    /// holds the caret turns its bar over at the pace the system's own caret blinks at (see
    /// `pin_playlist::advance_the_search`).
    pub(crate) caret_lit: bool,
    /// How much of the field is in sight, from `0.0` — parked out of the band's bottom edge — to
    /// `1.0`, whole. It is the one thing the band's own animation draws with: the field is carried
    /// down by the share of its height that is out and clipped to the strip the layout kept for it,
    /// so what a hand sees is a pill sliding under the window's edge rather than one blinking away
    /// (see `pin_playlist::advance_the_search`).
    pub(crate) slide: f32,
}

/// The rows a query lets through: a case-insensitive substring of the file's own name, in the order
/// the walk gave them.
///
/// **Only the name**, never the facts line or the clock: those are what the file says about itself,
/// and a hand looking for a track is naming a track. A query of nothing but spaces hides nothing,
/// which is the one reading of an empty field a reader would not have asked for.
pub(crate) fn playlist_shown(rows: &[PlaylistRow], query: &str) -> Vec<usize> {
    let needle = query.trim().to_lowercase();
    if needle.is_empty() {
        return (0..rows.len()).collect();
    }

    rows.iter()
        .enumerate()
        .filter(|(_, row)| row.name.to_lowercase().contains(&needle))
        .map(|(at, _)| at)
        .collect()
}

/// How tall a row is, in the display's own pixels: two lines of text and the pad above and below
/// them.
pub(crate) fn playlist_row_height(scale: f32) -> i32 {
    text_paint::scaled(ROW_PIXELS as i32, scale).max(1)
}

/// How far outside the band a pointer still counts as reaching for the band's field, in the
/// display's own pixels.
///
/// **The field is at the band's own bottom edge, so a hand arrives at it from under the window or
/// from beside it rather than through the rows on top of it.** A band that answered only inside its
/// own box is a band whose field appears once the pointer has already crossed the last row, and the
/// gesture a hand makes for a field at the bottom of a window is one that lands a few pixels short.
/// This is that shortfall: the reach the field answers to, and nothing else — the row a name travels
/// for is still only ever the pointer's inside the band's own rect (see `pin_playlist::set_the_hover`).
pub(crate) fn playlist_reach(scale: f32) -> i32 {
    text_paint::scaled(REACH_PIXELS as i32, scale).max(0)
}

/// The row the band's search field begins at, in the band's own rows: the field is anchored to the
/// band's bottom, so it floats over the last rows of the list rather than sitting in a strip the rows
/// are kept out of.
///
/// It is asked without a layout of its own because it is the answer a pointer is measured against
/// before there is a band to lay one out for — a hand asks for the field at the field's own rows, and
/// never at a row of the list above it (see `pin_playlist::hover_at_the_band`).
pub(crate) fn search_field_top(height: i32, scale: f32) -> i32 {
    let pad = text_paint::scaled(PAD_PIXELS as i32, scale).max(0);
    let search = text_paint::scaled(SEARCH_PIXELS as i32, scale).max(0);

    (height - pad - search).max(pad.min(height))
}

/// The height a band stands at, snapped to the rows it can draw: the pad above and below them and a
/// whole number of rows between.
///
/// **A band's bottom edge is a row's bottom edge.** The field floats over the last row rather than a
/// strip being kept from it, so a band left at an arbitrary height is a band whose list ends in a row
/// cut in half — a row of nothing at the bottom of the window. The height a band opens at and every
/// height a drag gives it are snapped here, which is the one place the two are counted.
pub(crate) fn snapped_playlist_height(height: i32, scale: f32) -> i32 {
    let pad = text_paint::scaled(PAD_PIXELS as i32, scale).max(0);
    let row = playlist_row_height(scale);
    let room = (height - 2 * pad).max(0);
    // Rounded to the nearest row rather than down: a drag that has come half a row past its last whole
    // row is a drag asking for that row.
    let rows = ((room + row / 2) / row).max(1);

    2 * pad + rows * row
}

/// The top row the search pill has at the share of it that is in sight: the whole strip down to the
/// band's own last row for a field parked under it, and the strip's own top for one in sight whole.
///
/// It is the line a row peeking under the field is cut against, so that the pill takes the row back as
/// it comes up rather than a band of page standing under it (see `paint_playlist` and `paint_search`).
fn painted_field_top(layout: &PlaylistLayout, slide: f32) -> i32 {
    let out = ((1.0 - slide.clamp(0.0, 1.0)) * (layout.search_bottom - layout.search_top) as f32)
        .round() as i32;

    layout.search_top + out
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

    let search = text_paint::scaled(SEARCH_PIXELS as i32, scale).max(0);
    let mark = text_paint::scaled(SEARCH_MARK_PIXELS as i32, scale).max(0);
    let search_pad = text_paint::scaled(SEARCH_PAD_PIXELS as i32, scale).max(0);
    let mark_gap = text_paint::scaled(SEARCH_MARK_GAP_PIXELS as i32, scale).max(0);

    let rows_top = pad.min(height);
    // **The rows are the band's own and the field is drawn over them.** A band is a list, and a band
    // that kept a strip of its own bottom for a field is a band that ended in a row of nothing: the
    // rows reach the band's bottom pad, and the field floats over the last of them while it is in sight
    // (see `paint_search` and `snapped_playlist_height`).
    let rows_bottom = (height - pad).max(rows_top);
    // The field is a constant of the layout still, anchored to the band's bottom rather than to its own
    // strip: its rows are the last rows of the band, and a band taller or shorter than this one gives
    // the rows above it what it took rather than stretching the field.
    let search_top = search_field_top(height, scale);
    let search_bottom = (search_top + search).min(rows_bottom).max(search_top);
    let search_left = pad.min(width);
    let search_right = (width - pad).max(search_left);
    // The mark is the first thing in the pill and the query is set after it, each held the pill's own
    // pad off the edge it begins at: a query drawn over the mark is a mark no one can see, and a
    // press on the mark is a press before the first character rather than one in it.
    let search_mark = search_left + search_pad;
    let search_text_left = (search_mark + mark + mark_gap).min(search_right);
    // **The clear button is the last thing in the pill**, at the same pad off its right edge the mark
    // keeps off its left, and the query is held the mark's own gap off it: a × a query runs under is a
    // × nothing can press, and a field whose text reaches its own edge has nowhere to say "clear".
    let close = text_paint::scaled(SEARCH_CLOSE_PIXELS as i32, scale).max(0);
    let search_close_right = (search_right - search_pad).max(search_text_left);
    let search_close_left = (search_close_right - close).max(search_text_left);
    let search_text_right = (search_close_left - mark_gap).max(search_text_left);
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
        search_top,
        search_bottom,
        search_left,
        search_right,
        search_mark,
        search_text_left,
        search_text_right,
        search_close_left,
        search_close_right,
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

/// The rows the search field's own box runs between, in the band's coordinates.
pub(crate) fn search_box(layout: &PlaylistLayout) -> (i32, i32) {
    (layout.search_top, layout.search_bottom)
}

/// Whether a band-local point is the search field's rather than a row's or the band's.
pub(crate) fn search_field_holds(layout: &PlaylistLayout, x: i32, y: i32) -> bool {
    let (top, bottom) = search_box(layout);

    (layout.search_left..layout.search_right).contains(&x) && (top..bottom).contains(&y)
}

/// Whether a band-local point of the field is the clear button's rather than the query's: the × at
/// the pill's right, and the pad between it and the pill's own edge, which is the one press on a
/// field that asks for nothing in it (see `paint_search`, which draws the × where this names it).
///
/// The button is a mark of the pill's own row rather than a box drawn around it, so its rows are the
/// whole field's: a hand made to find a mark ten pixels square is a hand the field was not drawn for.
pub(crate) fn search_close_holds(layout: &PlaylistLayout, x: i32, y: i32) -> bool {
    let (top, bottom) = search_box(layout);

    (layout.search_close_left..layout.search_right).contains(&x) && (top..bottom).contains(&y)
}

/// The column the query's first character is drawn at: the one the layout kept past the mark.
fn search_text_left(layout: &PlaylistLayout) -> i32 {
    layout.search_text_left
}

/// The place in the query a point of the field names: a press puts the caret between the characters
/// it landed between, measured from the column the query is drawn at rather than from the box's edge.
pub(crate) fn search_index_at(layout: &PlaylistLayout, text: &str, x: i32, scale: f32) -> usize {
    if text.is_empty() {
        return 0;
    }

    // shortcut: a query too wide for the field is cut at its right edge rather than scrolled under
    // its caret; upgrade if a query long enough to lose its own caret shows up.
    let Some(surface) = DibSurface::create(1, 1) else {
        return text.len();
    };
    let columns = measure_run(&surface, &facts_style(), text, facts_scale(scale));
    let character = super::find_bar::index_at_column(&columns, layout.search_text_left, x);

    text.char_indices()
        .nth(character)
        .map_or(text.len(), |(at, _)| at)
}

/// Whether a band-local point is the scrollbar's to answer rather than a row's.
///
/// The band a press grabs it in is wider than the mark it paints: it runs from the last column the
/// length can be inked in — `clock_right` is that column's exclusive edge, so one pixel left of it is
/// the length's own last pixel — to the window's own right edge, which is the two pixels past the
/// track as well as the track. Nothing is drawn in that room beyond the bar: a hand reaching for the
/// bar is not a hand made to find four pixels of it (see `pin_playlist::press`).
pub(crate) fn playlist_track_holds(layout: &PlaylistLayout, x: i32, y: i32) -> bool {
    layout.track.is_some()
        && x >= layout.clock_right - 1
        && x < layout.width
        && y >= layout.rows_top
        && y < layout.rows_bottom
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
    let mut boxes = Vec::new();
    let mut painter = text_paint::RunPainter::new(surface, scale);
    // The second line is a pixel under the first, so it is measured and drawn at the scale that
    // lands it there rather than at the band's own (see `facts_scale`).
    let small = facts_scale(scale);
    let mut small_painter = text_paint::RunPainter::new(surface, small);

    // The rows drawn are the positions the query let through, and everything a row carries is its
    // own index: a stripe and a wash are the row the pin shows and the row the hand picked, which
    // the filter never renumbers.
    let first = band.first.min(band.shown.len());
    let last = (first + layout.visible.max(0) as usize).min(band.shown.len());
    let drawn = band.shown.get(first..last).unwrap_or(&[]);

    // The rows drawn are the positions the query let through, and everything a row carries is its
    // own index: a stripe and a wash are the row the pin shows and the row the hand picked, which
    // the filter never renumbers.
    let mut places: Vec<(usize, i32, i32)> = drawn
        .iter()
        .copied()
        .enumerate()
        .map(|(step, index)| {
            let top = layout.rows_top + step as i32 * layout.row_height;
            let bottom = (top + layout.row_height).min(layout.rows_bottom);
            (index, top, bottom)
        })
        .collect();

    // **The room under the last whole row is a row peeking rather than a band of empty page.** A band
    // is snapped to whole rows, so there is normally nothing left under them; a band caught between two
    // rows — a drag in flight, or a height a clamp held short — leaves a part of the next row, and a page
    // there is a band that reads as a list which ended above its own bottom. The next row is drawn into
    // it and cut at the band's own bottom pad, which is the row the field floats over when it is in sight
    // (see `snapped_playlist_height` and `paint_search`).
    let peek_floor = layout.rows_bottom.min(height);
    if let Some(index) = band.shown.get(last).copied() {
        let top = layout.rows_top + (last - first) as i32 * layout.row_height;
        let bottom = (top + layout.row_height).min(peek_floor);
        if bottom > top {
            places.push((index, top, bottom));
        }
    }

    boxes.reserve(places.len() * 6);

    for (index, top, bottom) in places {
        let Some(row) = band.rows.get(index) else {
            continue;
        };
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

        // **A row is cut where the row is**, which for the rows that fit is their own bottom and for
        // the row peeking under the field is the pill's top: the two lines keep their place inside
        // the row, so a row cut short shows the part of itself that is not under the field yet. A box
        // cut to nothing draws nothing at all, which is a row the pill has come up over.
        let (mut name_box, mut facts_box) = line_boxes(&layout, top, scale);
        name_box.bottom = name_box.bottom.min(bottom).max(name_box.top);
        facts_box.bottom = facts_box.bottom.min(bottom).max(facts_box.top);

        // A name the left column has room for is drawn once, cut with an ellipsis where it has
        // none — while a name the band is travelling is drawn whole and moved under the column
        // instead, which is the whole of what a row the hand is on, one it has picked or the one the
        // pin is showing buys a name too long to read (see `pin_playlist::set_the_marquee`).
        let scrolled = band
            .scrolled
            .iter()
            .find(|(row, _)| *row == index)
            .map(|(_, offset)| *offset);
        let shown_name = match scrolled {
            Some(_) => row.name.clone(),
            None => fit_title(
                surface,
                &name_style,
                &row.name,
                name_box.right - name_box.left,
                scale,
            ),
        };
        let origin = name_box.left - scrolled.unwrap_or(0);

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

    // A band whose query hides every row says so rather than drawing an empty room: the note is
    // where the first row would have been, in the ink and the size of a row's own facts line.
    if let Some(note) = band.note {
        let run = RECT {
            left: layout.text_left,
            top: layout.rows_top
                + ((layout.row_height - text_paint::scaled(FACTS_PIXELS as i32, small)).max(0) / 2),
            right: layout.facts_right,
            bottom: layout.rows_bottom,
        };
        small_painter.draw(note, layout.text_left, run, &facts_style, muted, page);
        small_painter.forget_the_face();
        boxes.push(run);
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

    // The search field over the last of it, in the room the layout kept for it.
    if let Some(search) = band.search.as_ref() {
        boxes.extend(paint_search(
            surface, palette, &layout, search, scale, muted,
        ));
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

/// The search field's own pixels: a pill, the search mark at its left, the query after it, the caret
/// when it is focused, and the muted placeholder where the query is nothing.
///
/// Answers the boxes GDI wrote into, which the band's own seal gives their coverage back — a field
/// drawn but not sealed is a hole in an opaque band, and a hole in an opaque band is a press that
/// falls through to the listing behind the pin.
fn paint_search(
    surface: &DibSurface,
    palette: &ChromePalette,
    layout: &PlaylistLayout,
    search: &PlaylistSearch,
    scale: f32,
    muted: [u8; 3],
) -> Vec<RECT> {
    let width = surface.width as i32;
    let height = surface.height as i32;
    if width <= 0 || height <= 0 {
        return Vec::new();
    }

    let pixels = unsafe {
        std::slice::from_raw_parts_mut(
            surface.bits(),
            surface.width as usize * surface.height as usize * 4,
        )
    };

    let page = palette.background;
    // **The field is carried down by the share of it that is out of sight, and the rows it floats over
    // are not.** The pill is drawn in full at the row its slide puts it at, and everything of it that
    // has gone past the band's own bottom pad is painted back to the band's page after: what a hand sees
    // is a pill sliding under the window's edge and being cut there, which is the parking animation — a
    // pill that faded or blinked would be one that disappeared where it stood (see
    // `pin_playlist::advance_the_search`).
    let field_top = painted_field_top(layout, search.slide);
    let field = RECT {
        left: layout.search_left,
        top: field_top,
        right: layout.search_right,
        bottom: field_top + (layout.search_bottom - layout.search_top),
    };
    if field.right <= field.left || field.bottom <= field.top {
        return Vec::new();
    }
    // The rows the field stands in, which is the whole of it a hand may see: the field floats over the
    // last rows of the list, so what is above it is a row's and what is below it is the band's own pad,
    // which is not the field's to paint over.
    let strip = RECT {
        left: 0,
        top: layout.search_top,
        right: width,
        bottom: (layout.search_top + (layout.search_bottom - layout.search_top)).min(height),
    };

    // **The pill.** The plate is the band's own page lifted a little toward its ink, so that the field
    // reads as a surface in the band rather than as a band of two colours, and it is rounded to half
    // its height — past that a radius stops being a curve and becomes the box it started as.
    let radius = (field.bottom - field.top) as f32 / 2.0;
    let edge = text_paint::scaled(SEARCH_EDGE_PIXELS as i32, scale).max(1);
    let plate = text_paint::blend(page, palette.foreground, SEARCH_INK);
    // **The plate is filled a hairline short of the pill, and the outline is drawn after it.** A fill
    // drawn over an outline covers exactly the rows the outline is, so the border of a field is drawn
    // last of the two: a pill whose bottom edge is repainted by the box its text sits in is a field
    // with a border on three sides, which is what the bottom of this band used to show.
    fill_round_rect(pixels, width, field, radius, plate, plate, 1.0);
    stroke_round_rect(
        pixels,
        width,
        field,
        radius,
        edge as f32,
        palette.hover(SEARCH_EDGE_INK),
        1.0,
    );

    // **The mark**: the ring and the handle of a magnifier, drawn as the two strokes it is rather than
    // as a glyph, because a glyph of a face the theme chose is a mark whose weight the theme decided
    // (see `paint_search_mark`).
    let centre_y = (field.top + field.bottom) as f32 / 2.0;
    paint_search_mark(
        pixels,
        width,
        layout.search_mark,
        centre_y,
        text_paint::scaled(SEARCH_MARK_PIXELS as i32, scale),
        muted,
    );

    // **The clear button**, at the pill's other end: the × that puts the query back to nothing, drawn
    // in the muted ink the mark beside it is drawn in — a control about the text rather than text of
    // its own, and the query is the one bright thing in the field (see `search_close_holds`, which is
    // the press it answers).
    draw_cross(
        pixels,
        width,
        (layout.search_close_left + layout.search_close_right) / 2,
        centre_y as i32,
        layout.search_close_right - layout.search_close_left,
        (SEARCH_CLOSE_STROKE * scale).max(1.0),
        muted,
    );

    let text_left = search_text_left(layout);
    // **The query is set at the size of a row's facts line**, which is the note under a name rather than
    // a second name: the pill is a control inside the band, and a query in the band's own body size
    // would be the loudest thing in a window of names. It is one scale for the face, the box it sits
    // in and the caret that stands between its characters, so the three cannot disagree.
    let small = facts_scale(scale);
    let level = facts_style().level as usize;
    let cell = text_paint::scaled(text_paint::LEVEL_FONT_PIXELS[level], small).max(1);
    // **The run is the pill's own inside, never its edge.** GDI opaques the box it is handed, so a run
    // that reached the pill's last row would paint the border there back to the plate — a field with a
    // hairline on three sides. It is held the outline's own width inside the box, which leaves the
    // ring of the pill whole and the whole of a line's height to set the query in.
    //
    // It is one line of the face, centred in the pill rather than the pill's whole height: the box is
    // what GDI opaques, and a box a taller pill hands it is the pill's border rows painted over.
    let top = centre_y as i32 - cell / 2;
    let run = RECT {
        left: text_left,
        top: top.max(field.top + edge),
        right: layout.search_text_right.max(text_left),
        bottom: (top + cell).min(field.bottom - edge),
    };

    let mut painter = text_paint::RunPainter::new(surface, small);
    let style = facts_style();
    // A run opaques its box, so the text's own box is filled with the pill's plate first: what is
    // written over the field is set on the field rather than on the band's page.
    let under = RECT {
        left: run.left,
        right: run.right,
        ..run
    };
    fill_box(pixels, width, under, plate, 1.0);

    // The boxes the query's own pieces fill, which the band's seal gives their coverage back: a piece
    // drawn by GDI leaves the alpha byte of its box at zero, and a hole in an opaque band is a press
    // that falls through it (see `paint_playlist`).
    let mut boxes = Vec::new();

    if search.text.is_empty() {
        painter.draw(PLACEHOLDER, text_left, run, &style, muted, plate);
        // The placeholder is drawn as one run, so the box it filled is the whole of what the band's
        // seal has to give back.
        boxes.push(run);
    } else {
        let cut = fit_title(surface, &style, search.text, run.right - run.left, small);
        // **A selection is a wash between its two ends**, in the ink the find bar washes its own query
        // with: a field in a list and a bar over a page are one editor's widget, and a selection a hand
        // can see is a selection the keys that cut text act on (see `text_find_field::Field`).
        //
        // It is the field's while the field holds the caret, and only then: a wash left standing over
        // a field the keys have left is a selection nothing on the window acts on.
        let characters = cut.chars().count();
        let selected = search
            .selection
            .filter(|_| search.focused)
            .map(|(from, to)| {
                (
                    super::find_bar::char_index(search.text, from).min(characters),
                    super::find_bar::char_index(search.text, to).min(characters),
                )
            })
            .filter(|(from, to)| from != to);

        // The columns the run's own characters end at, measured once and read by the wash and by every
        // piece: one measurement, so that neither can name a place the other does not.
        let columns = measure_run(surface, &style, &cut, small);
        let wash = text_paint::blend(plate, palette.foreground, SEARCH_SELECTION_WASH);
        if let Some((from, to)) = selected {
            let washed = RECT {
                left: text_left + columns.get(from).copied().unwrap_or(0),
                top: run.top,
                right: text_left + columns.get(to).copied().unwrap_or(0),
                bottom: run.bottom,
            };
            if washed.right > washed.left {
                fill_box(pixels, width, washed, wash, 1.0);
            }
        }

        // **The run is drawn in the pieces the selection cuts it into**, each at its own column and
        // over its own background: a selected piece carries the wash and the rest carry the plate, so
        // the run and its wash cannot disagree about where one ends and the other begins.
        let pieces: Vec<char> = cut.chars().collect();
        let mut edges = vec![0usize, characters];
        if let Some((from, to)) = selected {
            edges.push(from);
            edges.push(to);
        }
        edges.sort_unstable();
        edges.dedup();

        for pair in edges.windows(2) {
            let (from, to) = (pair[0], pair[1]);
            if from >= to {
                continue;
            }

            let text: String = pieces[from..to].iter().collect();
            let left = text_left + columns[from];
            let piece = RECT {
                left,
                right: text_left + columns[to],
                ..run
            };
            if piece.right <= piece.left {
                continue;
            }
            let background = match selected {
                Some((start, end)) if start <= from && to <= end => wash,
                _ => plate,
            };
            painter.draw(&text, left, piece, &style, palette.foreground, background);
            boxes.push(piece);
        }
    }
    painter.forget_the_face();

    // **The caret is drawn solid, and it blinks**: a field that holds the caret says so with a bar that
    // turns over at the pace the system's own text fields do, and a bar that never blinked would be a
    // caret a hand cannot tell from the field's own mark (see `pin_playlist::advance_the_search`, which
    // is the tick the two halves of it are counted on).
    if search.focused && search.caret_lit {
        let caret = search.caret.min(search.text.len());
        let carried = measure(surface, &style, &search.text[..caret], small);
        let bar = text_paint::scaled(CARET_PIXELS as i32, scale).max(1);
        let at = (text_left + carried).clamp(run.left, (run.right - bar).max(run.left));
        // **The caret is the size of the face, not of the line box.** The run is a line of the font with
        // its leading above and below the glyphs, and a bar drawn the height of it stands taller than
        // every character it is between — so it is the font's own pixels, centred in the pill.
        let top = centre_y as i32 - cell / 2;

        // The accent, which is the ink the clock and the card's own caret are drawn in — and the
        // height of a line of the face the query is set in, so that a caret is the size of the
        // characters it stands between rather than of the box the line is laid out in.
        fill_box(
            pixels,
            width,
            RECT {
                left: at,
                top,
                right: at + bar,
                bottom: top + cell,
            },
            palette.accent,
            0.9,
        );
    }

    // **Whatever the pill has taken below its own rows is given back to the band's page**, and the rows'
    // own bottom is where the cut falls: what is under them is the band's pad, which nothing else in
    // this paint writes on, so the page is what belongs there.
    let below = RECT {
        top: strip.bottom,
        bottom: (field.bottom).min(height),
        ..strip
    };
    if below.bottom > below.top {
        fill_box(pixels, width, below, page, 1.0);
    }

    boxes
}

/// The magnifier at the field's left: a ring and the handle off its lower right, which is the mark a
/// person reads a text field by — two strokes of one face, so it matches the query beside it whatever
/// the theme's typeface is.
///
/// The ring's radius and the handle's length are the square's own numbers rather than a font's: a
/// glyph of a theme's face is a mark whose weight the theme decided, and this is chrome the band draws
/// itself as it draws its chevrons and its play mark.
fn paint_search_mark(
    pixels: &mut [u8],
    width: i32,
    left: i32,
    centre_y: f32,
    side: i32,
    ink: [u8; 3],
) {
    if side <= 0 {
        return;
    }

    let stroke = (GLYPH_STROKE_PIXELS * side as f32 / SEARCH_MARK_PIXELS).max(1.0);
    // The ring takes the square's upper two thirds and the handle the rest, so that the mark is the
    // square it was given rather than a smaller one floating in it.
    let radius = side as f32 * 0.30;
    let centre_x = left as f32 + radius + stroke / 2.0;
    let centre_y = centre_y - side as f32 * 0.06;
    fill_ring(pixels, width, centre_x, centre_y, radius, stroke, ink, 1.0);
    stroke_segment(
        pixels,
        width,
        (
            centre_x + radius * std::f32::consts::FRAC_1_SQRT_2,
            centre_y + radius * std::f32::consts::FRAC_1_SQRT_2,
        ),
        (
            centre_x + radius + stroke / 2.0,
            centre_y + radius + stroke / 2.0,
        ),
        stroke,
        ink,
        1.0,
    );
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

/// How tall the search field is, in logical pixels: a constant of the layout, so that a band dragged
/// taller or shorter gives the rows above it what it took rather than stretching the field.
///
/// It is a pill of one line of the query and the pad around it — a field in a list is a control a hand
/// reaches for, not a bar the band is built around, so it is kept to the height of a row's own note
/// rather than of a row.
const SEARCH_PIXELS: f32 = 24.0;

/// What an empty field says, in the muted ink and nothing else: it is a note about what the field is
/// for rather than a character it holds.
const PLACEHOLDER: &str = "Search names";

/// How wide the caret bar is: a hairline of accent rather than a stem, because it stands between
/// characters and is read as the place the next one lands rather than as a mark of its own.
const CARET_PIXELS: f32 = 1.5;

/// The clear button's cross: how wide the × is drawn by, and the stroke it is drawn at — the find
/// bar's own close mark to the pixel, so that the two fields of one app carry one button rather than
/// one button drawn two ways.
const SEARCH_CLOSE_PIXELS: f32 = 10.0;
const SEARCH_CLOSE_STROKE: f32 = 1.6;

/// The search mark's square, and the two gaps the pill keeps around it: its own pad off the left edge
/// of the pill, and the room between the mark and the first character of the query.
const SEARCH_MARK_PIXELS: f32 = 11.0;
const SEARCH_PAD_PIXELS: f32 = 10.0;
const SEARCH_MARK_GAP_PIXELS: f32 = 6.0;

/// How far the pill stands off the band's page: enough that the two read as a field inside the band
/// rather than as a band of two colours, and little enough that it is a tint of the band's own page
/// rather than a second surface in the window.
const SEARCH_INK: f32 = 0.07;

/// How far the hairline around the pill stands off the page, which is a shade under the line a card's
/// own boxes are drawn with: a pill's outline says where the field ends, and it is not a border.
const SEARCH_EDGE_INK: f32 = 0.22;

/// How wide that hairline is, in the display's own pixels: one row of the window, which is all an edge
/// of a field in a list is — and the width the query's own run is held inside it by (see
/// `paint_search`).
const SEARCH_EDGE_PIXELS: f32 = 1.0;

/// The pad the rows are held off every edge of the band by, and the room the drag handle at the
/// band's bottom edge is drawn in — the pad is the wider of the two on purpose, so that no press
/// on a row is ever a press on the handle.
const PAD_PIXELS: f32 = 10.0;

/// How far outside the band a pointer still counts as reaching for the band's field, in logical
/// pixels: the shortfall of a hand that comes at the field from under the window's own bottom edge
/// or from beside it, rather than through the rows above it (see `playlist_reach`).
const REACH_PIXELS: f32 = 16.0;

/// The room the accent stripe of the row the pin is showing takes at a row's left edge.
const STRIPE_PIXELS: f32 = 3.0;

/// The scrollbar in the band's right margin: how wide its track is — a bar a hand is not the reader
/// of, so it is the thinnest mark the display draws a column in — the gap that holds the rows off it,
/// and how far a row's name may run before the scrollbar's own room is taken off.
///
/// The width is the paint's alone: what a press grabs is the room from the length's last pixel to the
/// window's own edge, which is far wider than the mark itself (see `playlist_track_holds`).
const SCROLLBAR_PIXELS: f32 = 4.0;
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

/// How far the wash under a selected range of the query stands off the pill's own plate, which is the
/// amount the find bar washes its own selection by: the two fields are one editor's widget, and a
/// selection in one reads as a selection in the other (see `find_bar`).
const SEARCH_SELECTION_WASH: f32 = 0.25;

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

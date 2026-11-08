//! The card a pinned page's find bar is drawn in: where its two rows sit over the page, what each of
//! its controls answers, and the pixels of all of it.
//!
//! **A card rather than a strip.** A page of text is this app's own pixels the whole of the band down,
//! so there is no hosted window for the card to stand over and it is drawn into the pin's own surface
//! — the road the volume popup takes over a picture, one window simpler (see `pin_chrome_strips`).
//! It floats over the page's top-right corner, under the caption, and it is placed in the window's own
//! coordinates because that is where a press is answered (see `text_find::find_bar`).
//!
//! **Everything here is the display's scale.** A pad, a row, the glyphs and the shadow are all
//! `text_paint::scaled` pixels, so the card at 150% is the card rather than a small one in a large
//! corner.
//!
//! The card says what the reader typed and nothing else: the query as a field — its characters, the
//! washed selection between two of them and the caret at the focus end — the count of the place they
//! are on, the two switches and the three commands. Nothing is a Windows edit control: the query is
//! this app's own text in this app's own pixels, and every place in it is a byte index the field
//! keeps and this card turns into a column (see `text_find_field`).
//!
//! **A query too long for the field scrolls under the caret, and the field is what cuts it.** The
//! scroll is worked out from the caret as the paint measures the run, rather than stored (see
//! `query_run`), so nothing about it can go stale; and the run, the wash under its selection and the
//! caret are all clipped to the field's box — a character half outside it is cut at its edge rather
//! than drawn over the card's padding (see `paint_find_bar`).
//!
//! **The caret blinks, and only the tick turns it.** A page of text is not on the frame cadence, so a
//! caret that asked for a paint of its own would put one there; instead the blink is the field's own
//! clock, asked on the ticks a pin already takes, and this paint is handed the lit half it is in (see
//! `FindBarPaint::caret_lit` and `text_find::settle`).
//!
//! **A control is a button and is drawn as one.** The pointer over one of the five washes its own box,
//! a press washes it further and inks its mark with the accent, and the two chevrons are the buttons a
//! hand holds down rather than presses once (see `FindBarPaint` and `text_find`).

use super::bubble::{TOOLTIP_GAP_PIXELS, TOOLTIP_PADDING_PIXELS, TOOLTIP_RADIUS};
use super::primitives::{
    caption_style, draw_cross, fill_box, fill_round_rect, measure_run, measure_text,
    stroke_round_rect, stroke_segment, surface_pixels, ChromePalette,
};
use crate::text::text_paint::{self, DibSurface};
use std::cell::RefCell;
use windows::Win32::Foundation::RECT;

/// The card's size and the room its rows are given, in the pixels a display's scale multiplies.
const FIND_BAR_WIDTH: i32 = 320;
const FIND_BAR_ROW_HEIGHT: i32 = 22;
const FIND_BAR_PADDING: f32 = 7.0;
const FIND_BAR_MARGIN: f32 = 10.0;
const FIND_BAR_RADIUS: f32 = 9.0;
const FIND_BAR_SHADOW: f32 = 2.0;
/// The room the count keeps at the right end of the query's row, where `current/total` is drawn, at
/// its narrowest: a column narrower than the count it holds is a count with its last digits cut off,
/// so a count that needs more than this takes more (see `count_text`).
pub(super) const FIND_BAR_COUNT_WIDTH: i32 = 44;
/// The gap between the two step buttons and the close, and the wider one between the switches and the
/// commands — two groups a hand reaches for for different reasons.
const FIND_BAR_GAP: i32 = 2;
const FIND_BAR_GROUP_GAP: i32 = 9;

/// The card's marks, in the pixels a display's scale multiplies: a chevron nine across and five tall,
/// which is what makes it read as a chevron rather than as a narrow `v`, and a stroke uniform in every
/// direction rather than the caption's diagonals — those are narrowed by the root of two so that a
/// walk of whole columns comes out at their weight, and a mark drawn with `stroke_segment` has no walk
/// to answer to. The close cross takes a span of its own beside them, so that it weighs what the two
/// chevrons do rather than half again as much.
const FIND_BAR_CHEVRON_WIDTH: i32 = 9;
const FIND_BAR_CHEVRON_HEIGHT: i32 = 5;
const FIND_BAR_CLOSE_SPAN: i32 = 8;
const FIND_BAR_MARK_STROKE: f32 = 1.6;

/// How far a control's box is washed toward the card's own text while the pointer is on it, and how
/// much further while it is held: the two washes a caption's button takes in the same two states (see
/// `caption::paint_button`), and the wash a switch that is on carries under both of them.
const FIND_BAR_HOVER_WASH: f32 = 0.10;
const FIND_BAR_PRESSED_WASH: f32 = 0.18;
const FIND_BAR_ON_WASH: f32 = 0.24;

/// How far the field's selected characters are washed toward the card's text: the same shade the
/// page itself washes its selection with where its theme names no selection colour of its own, so a
/// selection in the query reads as the selection under it (see `layout::paint` and its
/// `blend(page, foreground, 0.25)`).
const FIND_BAR_SELECTION_WASH: f32 = 0.25;

/// The room the caret keeps from the field's edge once a query too wide for the field has scrolled to
/// keep it in view, in the pixels a display's scale multiplies: enough that the caret is never flush
/// against an edge, and no more — every pixel of it is a pixel of the query the field could have
/// shown instead.
pub(super) const FIND_BAR_CARET_MARGIN: i32 = 2;

/// Where a find bar is, in the window's own coordinates: the card, the row a query is typed into, the
/// count at its right end, and the five controls under it.
#[derive(Clone, Copy, Default)]
pub(crate) struct FindBar {
    pub(crate) panel: RECT,
    pub(crate) query: RECT,
    pub(crate) count: RECT,
    /// The case switch, written `Aa`.
    pub(crate) case: RECT,
    /// The whole-word switch, written `W`.
    pub(crate) whole_word: RECT,
    pub(crate) previous: RECT,
    pub(crate) next: RECT,
    pub(crate) close: RECT,
}

/// One of the card's five controls, named rather than numbered: a press is answered against the name,
/// and the number a control sits at moves the moment another is put beside it.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum FindBarPart {
    Case,
    WholeWord,
    Previous,
    Next,
    Close,
}

impl FindBar {
    /// Where a point in the window falls on the card's controls, if it falls on one: the switches are
    /// asked first, then the three commands, so two boxes that a page too narrow has pushed into each
    /// other are still answered in the order a reader reads them.
    pub(crate) fn part_at(&self, x: i32, y: i32) -> Option<FindBarPart> {
        [
            (self.case, FindBarPart::Case),
            (self.whole_word, FindBarPart::WholeWord),
            (self.previous, FindBarPart::Previous),
            (self.next, FindBarPart::Next),
            (self.close, FindBarPart::Close),
        ]
        .into_iter()
        .find(|(rect, _)| inside(*rect, x, y))
        .map(|(_, part)| part)
    }

    /// Whether a point is anywhere on the card, its controls and its own face alike: the card floats
    /// over the page, so nothing under it is what a hand there meant (see `text_find::press`).
    pub(crate) fn holds(&self, x: i32, y: i32) -> bool {
        inside(self.panel, x, y)
    }

    /// Whether a point is in the row the query is drawn in, which is where a press places the caret
    /// rather than asking nothing of the bar (see `text_find::press_the_field`).
    pub(crate) fn holds_query(&self, x: i32, y: i32) -> bool {
        inside(self.query, x, y)
    }
}

/// What the card says: the query, the count of the place the reader is on, and the two switches.
///
/// A value rather than a borrow, like every other answer a paint is handed: the page's state is read
/// under its lock and the lock is let go of before anything is drawn (see `PinnedPaint`).
pub(crate) struct FindBarPaint {
    pub(crate) query: String,
    /// The caret's place and the other end of the selection, in byte indices of `query` on character
    /// boundaries. `anchor` equal to `caret` is a caret with nothing selected; anything else is a
    /// selection drawn washed between the two, with the caret at `caret` whichever end that is (see
    /// `text_find_field::Field`).
    pub(crate) caret: usize,
    pub(crate) anchor: usize,
    /// Whether the caret is drawn at this moment, which is the blink the field's own clock turns
    /// over: solid from every edit, and off for half of each blink after it (see `CaretBlink`).
    pub(crate) caret_lit: bool,
    /// The control the pointer is over, and the one a press is holding: what a control's box is washed
    /// for, and what the two chevrons step from while they are held (see `text_find::hover` and
    /// `text_find::press`).
    pub(crate) hovered: Option<FindBarPart>,
    pub(crate) pressed: Option<FindBarPart>,
    /// The match the reader is on, or nothing while the query is empty or found nowhere — which is
    /// the `0/0` a count reads as.
    pub(crate) current: Option<usize>,
    pub(crate) total: usize,
    pub(crate) case_sensitive: bool,
    pub(crate) whole_word: bool,
}

/// What the card's count says: the place the reader is on out of the places the query was found, or
/// the `0/0` a count reads as when the query is empty or found nowhere.
///
/// It is named here rather than in the paint because the card's layout is written from it: the count's
/// own column is as wide as its text, so the field beside it is laid out against the count really
/// drawn and not against one count's worth of room (see `find_bar_layout`).
pub(crate) fn count_text(current: Option<usize>, total: usize) -> String {
    if total == 0 {
        "0/0".to_string()
    } else {
        format!("{}/{}", current.map_or(0, |current| current + 1), total)
    }
}

/// Where the card is drawn for a page of `width` x `height` whose box begins at `top`: against the
/// page's top-right corner, inside the page's own margin, and as wide as the page leaves room for —
/// a card wider than its page would be a card drawn over the frame's edge.
///
/// `count` is what the card is about to say at its right end, and the column it is given is measured
/// from it rather than fixed: a page searched for a word it holds two thousand of has a wider count
/// than a page searched for one it holds nine of, and the room is taken out of the field rather than
/// out of the count's last digits (see `count_text`).
///
/// Every rectangle is in the window's coordinates rather than the page's: the page is the band between
/// the pin's chrome, and a press arrives in the window's own points (see `text_find::find_bar`).
pub(crate) fn find_bar_layout(width: i32, height: i32, top: i32, dpi: u32, count: &str) -> FindBar {
    let scale = dpi as f32 / 96.0;
    let margin = text_paint::scaled(FIND_BAR_MARGIN as i32, scale);
    let padding = text_paint::scaled(FIND_BAR_PADDING as i32, scale);
    let row = text_paint::scaled(FIND_BAR_ROW_HEIGHT, scale);
    let gap = text_paint::scaled(FIND_BAR_GAP, scale).max(1);
    let group_gap = text_paint::scaled(FIND_BAR_GROUP_GAP, scale).max(gap);
    // The column is the room the count's own text needs, and never less than the room a count has
    // always been given: a count that fits moves nothing as it changes, and a count that does not is
    // drawn whole rather than cut at its last digit.
    let count_width = text_paint::scaled(FIND_BAR_COUNT_WIDTH, scale)
        .max(query_columns(count, dpi).last().copied().unwrap_or(0));
    // A command is a box the row's own height across, so a hand has a square to land in rather than a
    // mark to hit: the mark inside it is the smaller glyph the caption's buttons are drawn with.
    let side = row;

    // The card keeps its margin where the page has room for it and gives it up rather than paint past
    // the page's edge on one that does not.
    let right = (width - margin).max(1);
    let panel_width = text_paint::scaled(FIND_BAR_WIDTH, scale)
        .min((right - margin).max(padding * 2 + side))
        .min(right);
    let wanted_height = padding * 2 + row * 2;
    let panel_top = top + margin;
    // A page too short to hold the card's two rows draws no card at all rather than a card painted
    // over the transport bar below it: there is nowhere for a hand to type on such a page anyway.
    if height - margin < wanted_height {
        return FindBar::default();
    }
    let panel_height = wanted_height;

    let panel = RECT {
        left: right - panel_width,
        top: panel_top,
        right,
        bottom: panel_top + panel_height,
    };

    let inner_right = panel.right - padding;
    let query_row = RECT {
        left: panel.left + padding,
        top: panel.top + padding,
        right: (inner_right - count_width - padding).max(panel.left + padding + 1),
        bottom: panel.top + padding + row,
    };
    let count = RECT {
        left: (inner_right - count_width).max(query_row.right),
        top: query_row.top,
        right: inner_right,
        bottom: query_row.bottom,
    };

    let command_row_top = query_row.bottom + padding;

    // The commands from the right: close, then the two steps, then the two switches, each in a box of
    // its own with the commands tighter together than the switches are (see `FIND_BAR_GROUP_GAP`).
    let close = RECT {
        left: inner_right - side,
        top: command_row_top,
        right: inner_right,
        bottom: command_row_top + row,
    };
    let next = RECT {
        left: close.left - gap - side,
        top: command_row_top,
        right: close.left - gap,
        bottom: command_row_top + row,
    };
    let previous = RECT {
        left: next.left - gap - side,
        top: command_row_top,
        right: next.left - gap,
        bottom: command_row_top + row,
    };
    let whole_word = RECT {
        left: previous.left - group_gap - side,
        top: command_row_top,
        right: previous.left - group_gap,
        bottom: command_row_top + row,
    };
    let case = RECT {
        left: whole_word.left - gap - side,
        top: command_row_top,
        right: whole_word.left - gap,
        bottom: command_row_top + row,
    };

    // A page narrow enough to push a mark off the card's left edge keeps every box on it: an
    // off-card name would answer at a point nothing is drawn at.
    let on_the_card = |rect: RECT| RECT {
        left: rect.left.max(panel.left + padding),
        ..rect
    };

    FindBar {
        panel,
        query: query_row,
        count,
        case: on_the_card(case),
        whole_word: on_the_card(whole_word),
        previous: on_the_card(previous),
        next: on_the_card(next),
        close: on_the_card(close),
    }
}

/// The room a switch's tooltip takes below the card, in the pixels a display's scale multiplies: the
/// gap it hangs by, its own padding about its cell, and the shadow it casts. It is the row the surface
/// a find bar is written on keeps below the card, because the name is written there and carried across
/// below the card (see `find_bar_surface_size` and `paint_tooltip`).
fn tooltip_room(dpi: u32) -> i32 {
    let scale = dpi as f32 / 96.0;
    let (padding, gap, _, shadow) = tooltip_shape(scale);
    gap + padding + tooltip_cell(scale) + shadow
}

/// The card's text height and the numbers a switch's tooltip is drawn from at one scale: the room it
/// keeps about its name, the gap it hangs below the card by, its corner radius and the shadow under
/// it — the four the caption's own tooltip is drawn with (see `paint_tooltip`).
fn tooltip_cell(scale: f32) -> i32 {
    text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[text_paint::BODY_LEVEL as usize],
        scale,
    )
}

fn tooltip_shape(scale: f32) -> (i32, i32, f32, i32) {
    (
        text_paint::scaled(TOOLTIP_PADDING_PIXELS as i32, scale),
        text_paint::scaled(TOOLTIP_GAP_PIXELS as i32, scale),
        (TOOLTIP_RADIUS * scale).max(1.0),
        (2.0 * scale).round().max(1.0) as i32,
    )
}

/// The name a switch says while a pointer rests on it, or nothing for a control that has nothing to
/// say: `Aa` and `W` are not words, where the two chevrons and the cross are the marks every window a
/// hand has ever used has drawn its own steps and its own close with (see `paint_find_bar`).
///
/// A pressed switch says it as well, because a press is a hand on it: a name that went away the moment
/// the button was pressed would leave the reader guessing what the box they just washed meant.
fn switch_tooltip(part: FindBarPart) -> Option<&'static str> {
    match part {
        FindBarPart::Case => Some("Match Case"),
        FindBarPart::WholeWord => Some("Match Whole Word"),
        _ => None,
    }
}

/// The size of the surface a find bar is written on: the card, and the row below it a switch's tooltip
/// hangs in.
///
/// The surface is the caller's, and it is the card's own size plus that row because the name is
/// written on it and carried across the window below the card, the way the card's own text is written
/// on it and carried across the card (see `paint_find_bar` and `carry_text_across`).
pub(crate) fn find_bar_surface_size(card: &FindBar, dpi: u32) -> (u32, u32) {
    (
        (card.panel.right - card.panel.left).max(1) as u32,
        (card.panel.bottom - card.panel.top).max(1) as u32 + tooltip_room(dpi).max(0) as u32,
    )
}

/// Whether a point is inside a rectangle, which is half-open the way every box of the chrome is:
/// two buttons that meet share an edge rather than a column.
fn inside(rect: RECT, x: i32, y: i32) -> bool {
    x >= rect.left && x < rect.right && y >= rect.top && y < rect.bottom
}

thread_local! {
    /// A one-pixel surface the query's own columns are measured on, kept for the life of the thread:
    /// a press, every move of the drag after it and every blink of the caret ask the same question,
    /// and a surface — with the font selected into it — made per message would be the field's cost
    /// for the one thing it draws. It is not the card's own surface, because a press arrives before
    /// any paint of the bar it is a press on.
    static MEASURE_SURFACE: RefCell<Option<DibSurface>> = const { RefCell::new(None) };
}

/// Where the query is drawn from, given how wide it is: the field's left edge where it fits, and its
/// right edge less the run where it does not. A query too long for its row rests with its end — the
/// end a reader is typing at — on the field, and its beginning is what is cut (see `query_run`, which
/// is where the caret scrolls it to from here).
pub(super) fn query_origin(bar: &FindBar, measured: i32) -> i32 {
    (bar.query.right - measured).min(bar.query.left)
}

/// Where a query's characters are drawn in its field: the columns they end at, and the column the run
/// begins at, both in the window's coordinates.
///
/// One answer rather than two, because a press has to name the character a paint put under the
/// pointer: `paint_find_bar` and `query_index_at` both read this, so a field that has scrolled is
/// pressed where it is drawn rather than a field's width away from the hand (see `query_run`).
pub(super) struct QueryRun {
    /// The run's prefix widths, beginning with a zero — so `columns[k]` ends character `k` — and the
    /// run's whole width last.
    pub(super) columns: Vec<i32>,
    /// The column, in the window's coordinates, the first character begins at.
    pub(super) origin: i32,
}

/// The run a query is drawn as in a field with the caret at `caret`, a character index.
///
/// **The scroll is derived from the caret and stored nowhere**, so a resize, a paste that replaces the
/// query and an undo can none of them leave the field looking at a column the text has left. A query
/// that fits never scrolls: there is nothing to bring into the field, and a margin taken from a run
/// the reader can see whole would cut its beginning to show them air.
///
/// Past that, the rule every single-line field follows. The run rests where `query_origin` puts it, and
/// while the caret's own column is inside the field it stays exactly there — a caret moving about
/// inside the part of the query the reader can see does not shuffle the text under it. Once the caret
/// is past either edge the run moves the least it can to put the caret back inside, a margin clear of
/// the edge, so that a caret typing at the end of a long query is a character the reader sees land
/// rather than one drawn on the field's last column.
pub(super) fn query_run(bar: &FindBar, columns: Vec<i32>, caret: usize, dpi: u32) -> QueryRun {
    let measured = columns.last().copied().unwrap_or(0);
    let rest = query_origin(bar, measured);
    if measured <= bar.query.right - bar.query.left {
        return QueryRun {
            columns,
            origin: rest,
        };
    }

    let caret_column = columns.get(caret).copied().unwrap_or(measured);
    let margin = text_paint::scaled(FIND_BAR_CARET_MARGIN, dpi as f32 / 96.0).max(1);
    let settled = rest + caret_column;
    let origin = if settled < bar.query.left + margin {
        bar.query.left + margin - caret_column
    } else if settled > bar.query.right - margin {
        bar.query.right - margin - caret_column
    } else {
        rest
    };

    QueryRun { columns, origin }
}

/// The byte index in `query` a point names: the character boundary nearest the point, measured from
/// the column the query is really drawn at rather than from the field's left edge.
///
/// A press between two letters means the boundary between them, so a point on a character's left
/// half puts the caret before it and one on its right half after it. The index is a byte index on a
/// character boundary, which is what the field's caret is (see `text_find_field::Field`).
///
/// `caret` is where the field's caret stands, which is what scrolls a query too wide for its row: a
/// press is answered against the run that caret left on the field, not against one right-aligned from
/// a field that has scrolled somewhere else (see `query_run`).
pub(crate) fn query_index_at(bar: &FindBar, query: &str, caret: usize, x: i32, dpi: u32) -> usize {
    if query.is_empty() {
        return 0;
    }

    let columns = query_columns(query, dpi);
    let run = query_run(bar, columns, char_index(query, caret), dpi);
    let character = index_at_column(&run.columns, run.origin, x);

    query
        .char_indices()
        .nth(character)
        .map_or(query.len(), |(at, _)| at)
}

/// The columns a query's characters end at, measured on the kept surface at the display's scale.
///
/// The style's colours are the card's own, but only its metrics are read: a font's extents do not
/// depend on the ink it is asked for, and the palette is not read here because a press can arrive
/// while the theme is being reloaded and a place in a query is not worth waiting for one.
fn query_columns(query: &str, dpi: u32) -> Vec<i32> {
    MEASURE_SURFACE.with(|cell| {
        let mut cell = cell.borrow_mut();
        if cell.is_none() {
            *cell = DibSurface::create(1, 1);
        }

        cell.as_ref().map_or_else(
            || vec![0],
            |surface| measure_run(surface, &caption_style([0, 0, 0]), query, dpi as f32 / 96.0),
        )
    })
}

/// The place in a run a column names: the boundary nearest the point, given the column the run is
/// drawn from and the width of every character before it.
///
/// `columns` is the run's own prefix widths, beginning with a zero: `columns[k]` is the width of the
/// first `k` characters, and its last element is the run's whole width. The origin is passed rather
/// than assumed to be the field's left edge, because a query too long for its row is drawn from the
/// right (see `query_origin`).
pub(super) fn index_at_column(columns: &[i32], origin: i32, x: i32) -> usize {
    let at = x - origin;
    if at <= 0 || columns.len() < 2 {
        return 0;
    }

    for index in 1..columns.len() {
        let (left, right) = (columns[index - 1], columns[index]);
        if at < right {
            let before = (at - left) * 2 < right - left;
            return if before { index - 1 } else { index };
        }
    }

    columns.len() - 1
}

/// Paint the card: its panel and hairline, the two switches where they are on, the three marks, the
/// query with its caret, and the count.
///
/// The panel and the marks are drawn straight into the window's own pixels, because that is what they
/// are — a card over a picture, with the page behind it. The text is not, because GDI needs a device
/// context: it is drawn onto the surface the caller keeps, and then carried across the card (see
/// `carry_text_across`), which is the road the caption's tooltip and the menu's labels take.
#[allow(clippy::too_many_arguments)] // The card, what it says, and the surface it is written on are one paint.
pub(crate) fn paint_find_bar(
    buffer: &mut [u8],
    width: i32,
    palette: &ChromePalette,
    bar: &FindBar,
    state: &FindBarPaint,
    surface: &DibSurface,
    dpi: u32,
) {
    if bar.panel.bottom <= bar.panel.top || bar.panel.right <= bar.panel.left {
        return;
    }

    let scale = dpi as f32 / 96.0;
    let radius = (FIND_BAR_RADIUS * scale).max(1.0);
    let shadow = text_paint::scaled(FIND_BAR_SHADOW as i32, scale);
    let stroke = (FIND_BAR_MARK_STROKE * scale).max(1.0);
    let panel_color = palette.hover(0.05);

    // The shadow, the panel, and a hairline around it: a flat panel rather than a shaded one, because
    // a query written across a gradient has a shade running through it — the same reason the
    // caption's tooltip is flat (see `paint_tooltip`).
    fill_round_rect(
        buffer,
        width,
        RECT {
            top: bar.panel.top + shadow,
            bottom: bar.panel.bottom + shadow,
            ..bar.panel
        },
        radius,
        [0, 0, 0],
        [0, 0, 0],
        0.32,
    );
    // The panel is opaque rather than almost so: the query and the count are written on a surface of
    // the card's own and carried over it as whole pixels, and a panel that is even a shade transparent
    // shows the run's box as a band of another colour across itself.
    fill_round_rect(
        buffer,
        width,
        bar.panel,
        radius,
        panel_color,
        panel_color,
        1.0,
    );
    stroke_round_rect(
        buffer,
        width,
        bar.panel,
        radius,
        (1.0 * scale).round().max(1.0),
        palette.hover(0.32),
        1.0,
    );

    // A switch that is on is a button held down: the same wash and the same accent ink a caption's
    // button uses for the state it is in, so that `Aa` on and `Aa` off are told apart at a glance.
    //
    // The pointer's own two states are washed over that wash rather than instead of it, in this order,
    // and `washed` below is the same order applied to a colour — the two have to agree, because the run
    // of a switch's name is written against the colour its box was washed to and carried across as
    // whole pixels (see `carry_text_across`).
    let toward = if palette.dark {
        [255, 255, 255]
    } else {
        [0, 0, 0]
    };
    let washed = |part: FindBarPart, on: bool| {
        let mut face = panel_color;
        if on {
            face = text_paint::blend(face, palette.accent, FIND_BAR_ON_WASH);
        }
        if state.pressed == Some(part) {
            face = text_paint::blend(face, toward, FIND_BAR_PRESSED_WASH);
        } else if state.hovered == Some(part) {
            face = text_paint::blend(face, toward, FIND_BAR_HOVER_WASH);
        }
        face
    };

    let switched = |part: FindBarPart| match part {
        FindBarPart::Case => state.case_sensitive,
        FindBarPart::WholeWord => state.whole_word,
        _ => false,
    };

    for (rect, part) in [
        (bar.case, FindBarPart::Case),
        (bar.whole_word, FindBarPart::WholeWord),
        (bar.previous, FindBarPart::Previous),
        (bar.next, FindBarPart::Next),
        (bar.close, FindBarPart::Close),
    ] {
        let on = switched(part);
        if on {
            fill_round_rect(
                buffer,
                width,
                rect,
                (radius * 0.5).max(1.0),
                palette.accent,
                palette.accent,
                FIND_BAR_ON_WASH,
            );
        }
        let amount = if state.pressed == Some(part) {
            Some(FIND_BAR_PRESSED_WASH)
        } else if state.hovered == Some(part) {
            Some(FIND_BAR_HOVER_WASH)
        } else {
            None
        };
        if let Some(amount) = amount {
            fill_round_rect(
                buffer,
                width,
                rect,
                (radius * 0.5).max(1.0),
                toward,
                toward,
                amount,
            );
        }
    }

    // A control whose press is held, or a switch that is on, carries the accent: the ink is what says
    // which button was pressed rather than merely that a pointer is over the row.
    let ink_of = |part: FindBarPart| {
        if switched(part) || state.pressed == Some(part) {
            palette.accent
        } else {
            palette.foreground
        }
    };

    for (rect, part, pointing_up) in [
        (bar.previous, FindBarPart::Previous, true),
        (bar.next, FindBarPart::Next, false),
    ] {
        let center_x = (rect.left + rect.right) / 2;
        let center_y = (rect.top + rect.bottom) / 2;
        draw_step_mark(
            buffer,
            width,
            center_x,
            center_y,
            scale,
            ink_of(part),
            pointing_up,
        );
    }

    let close_center_x = (bar.close.left + bar.close.right) / 2;
    let close_center_y = (bar.close.top + bar.close.bottom) / 2;
    draw_cross(
        buffer,
        width,
        close_center_x,
        close_center_y,
        text_paint::scaled(FIND_BAR_CLOSE_SPAN, scale).max(4),
        stroke,
        ink_of(FindBarPart::Close),
    );

    // The text: the query, its caret, the count and the two switch names. Drawn on the caller's
    // surface, which is the card's own size, and sealed over the boxes the runs fill (see
    // `paint_tooltip` for why GDI's alpha byte has to be given back).
    let cell = text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[text_paint::BODY_LEVEL as usize],
        scale,
    );
    let centered_top = |rect: RECT| rect.top + ((rect.bottom - rect.top - cell) / 2).max(0);
    let style = caption_style(palette.foreground);
    // Everything the text is drawn against is the card's own surface, so a box of the window's is
    // taken back to the card's corner first (see `carry_text_across`).
    let on_the_surface = |rect: RECT| RECT {
        left: rect.left - bar.panel.left,
        top: rect.top - bar.panel.top,
        right: rect.right - bar.panel.left,
        bottom: rect.bottom - bar.panel.top,
    };
    // **The field is the limit on everything the query draws.** A character the run has scrolled half
    // outside the field is cut at its edge rather than drawn over the card's padding, and the box that
    // is sealed and carried across is this clipped one and not the run's whole width: a box as wide as
    // the run would seal the columns the field never showed and carry them over its own panel as a band
    // of another colour (see `seal_text` and `carry_text_across`).
    let in_the_field = |rect: RECT| RECT {
        left: rect.left.max(bar.query.left),
        top: rect.top.max(bar.query.top),
        right: rect.right.min(bar.query.right),
        bottom: rect.bottom.min(bar.query.bottom),
    };
    let mut painter = text_paint::RunPainter::new(surface, scale);
    // The window-coordinate boxes the runs really fill, so that what is sealed and carried across is
    // the run's own box and not a button's whole width: a sealed pixel the run never reached would be
    // a blank one carried over the panel as black.
    let mut boxes: Vec<RECT> = Vec::new();

    // The query, as the columns its characters end at: measured once, and read by everything that has
    // to place something in the run — the wash under a selection, the pieces the run is drawn in, and
    // the caret. A query wider than the room it has is scrolled by the caret and cut at the field: the
    // two ends of the selection are counted in characters of the run rather than in its bytes, and the
    // column the run begins at comes from the caret alone (see `query_run`).
    let columns = if state.query.is_empty() {
        vec![0]
    } else {
        measure_run(surface, &style, &state.query, scale)
    };

    // The two ends of the selection, in characters of the run rather than in its bytes: a column is
    // one character wide here, and a character outside the basic plane is one column as well as one
    // `char` (see `measure_run`).
    let characters = state.query.chars().count();
    let caret = char_index(&state.query, state.caret).min(characters);
    let anchor = char_index(&state.query, state.anchor).min(characters);
    let selected = (anchor != caret).then(|| (anchor.min(caret), anchor.max(caret)));

    let QueryRun { columns, origin } = query_run(bar, columns, caret, dpi);
    let measured = columns.last().copied().unwrap_or(0);
    let line_top = centered_top(bar.query);

    // A selection is washed in the field the way the page washes its own. It is put down in the
    // window's pixels before the runs are carried across, so that a glyph over it is composited over
    // the wash rather than over the panel; the pieces below then draw their own boxes in the same
    // wash, which keeps what is carried across the run's own colour.
    let wash = text_paint::blend(panel_color, palette.foreground, FIND_BAR_SELECTION_WASH);
    if let Some((from, to)) = selected {
        let box_ = in_the_field(RECT {
            left: origin + columns[from],
            top: line_top,
            right: origin + columns[to],
            bottom: line_top + cell,
        });
        if box_.right > box_.left {
            fill_box(buffer, width, box_, wash, 1.0);
        }
    }

    // The run is drawn in the pieces the selection cuts it into, each at its own column and over its
    // own background: a selected piece's background is the wash, so the run and its selection cannot
    // disagree about where one ends and the other begins.
    if !state.query.is_empty() {
        let pieces: Vec<char> = state.query.chars().collect();
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
            let left = origin + columns[from];
            // The piece is drawn from its own column, which for a scrolled run can be off the field:
            // the text's origin stays there so the characters in view are the ones the columns named,
            // and the clipped box is what cuts the ones the field has no room for.
            let box_ = in_the_field(RECT {
                left,
                top: line_top,
                right: origin + columns[to],
                bottom: line_top + cell,
            });
            if box_.right <= box_.left {
                continue;
            }
            let background = match selected {
                Some((start, end)) if start <= from && to <= end => wash,
                _ => panel_color,
            };
            painter.draw(
                &text,
                left - bar.panel.left,
                on_the_surface(box_),
                &style,
                palette.foreground,
                background,
            );
            boxes.push(box_);
        }
    }

    let count = count_text(state.current, state.total);
    let count_width = measure_text(surface, &style, &count, scale);
    // The count is right-aligned at the row's end and grows leftwards over the room the layout gave
    // it, rather than being cut at the column a shorter count was laid out in (see `find_bar_layout`).
    let count_left = (bar.count.right - count_width).max(bar.query.left);
    let count_top = centered_top(bar.count);
    let count_box = RECT {
        left: count_left,
        top: count_top,
        right: bar.count.right,
        bottom: count_top + cell,
    };
    painter.draw(
        &count,
        count_left - bar.panel.left,
        on_the_surface(count_box),
        &style,
        palette.foreground,
        panel_color,
    );
    boxes.push(count_box);

    for (rect, text, part) in [
        (bar.case, "Aa", FindBarPart::Case),
        (bar.whole_word, "W", FindBarPart::WholeWord),
    ] {
        let measured = measure_text(surface, &style, text, scale);
        let left = rect.left + ((rect.right - rect.left - measured) / 2).max(0);
        let top = centered_top(rect);
        let box_ = RECT {
            left,
            top,
            right: (left + measured).min(rect.right),
            bottom: top + cell,
        };
        painter.draw(
            text,
            left - bar.panel.left,
            on_the_surface(box_),
            &style,
            ink_of(part),
            washed(part, switched(part)),
        );
        boxes.push(box_);
    }

    // The name a switch says while a pointer rests on it: a card of its own, hanging below the card,
    // the way a caption's tooltip hangs below the strip — because the card has no room for a word that
    // is not the query (see `switch_tooltip` and `paint_tooltip`).
    //
    // Its panel is put down straight into the window's pixels, because it is a panel over the page; its
    // name is not, because GDI needs a device context, so it is written on the surface the card's text
    // is written on — in the row `find_bar_surface_size` keeps below the card for it — and carried
    // across with everything else (see `seal_text` and `carry_text_across`).
    let buffer_height = (buffer.len() / (width.max(1) as usize * 4)) as i32;
    if let Some(text) = state.hovered.and_then(switch_tooltip) {
        let anchor = if state.hovered == Some(FindBarPart::WholeWord) {
            bar.whole_word
        } else {
            bar.case
        };
        let (padding, gap, radius, shadow) = tooltip_shape(scale);
        let text_width = measure_text(surface, &style, text, scale);
        let panel_width = (text_width + padding * 2).min(width.max(0));
        let panel_height = cell + padding;
        let top = bar.panel.bottom + gap;
        let left =
            ((anchor.left + anchor.right - panel_width) / 2).clamp(0, (width - panel_width).max(0));
        let panel = RECT {
            left,
            top,
            right: left + panel_width,
            bottom: top + panel_height,
        };

        // A name with nowhere to be drawn is not drawn at all: half a tooltip over the page's last row
        // is worse than the name waiting for a window with room for it.
        if panel_width > 0
            && panel_height > 0
            && panel.bottom + shadow <= buffer_height
            && panel.bottom - bar.panel.top <= surface.height as i32
        {
            fill_round_rect(
                buffer,
                width,
                RECT {
                    top: panel.top + shadow,
                    bottom: panel.bottom + shadow,
                    ..panel
                },
                radius,
                [0, 0, 0],
                [0, 0, 0],
                0.32,
            );
            fill_round_rect(
                buffer,
                width,
                panel,
                radius,
                palette.hover(0.0),
                palette.hover(0.0),
                1.0,
            );
            stroke_round_rect(
                buffer,
                width,
                panel,
                radius,
                (1.0 * scale).round().max(1.0),
                palette.hover(0.30),
                1.0,
            );

            let name_top = panel.top + ((panel_height - cell) / 2).max(0);
            let name_box = RECT {
                left: panel.left + padding,
                top: name_top,
                right: (panel.left + padding + text_width.max(0)).min(panel.right),
                bottom: (name_top + cell).min(panel.bottom),
            };
            painter.draw(
                text,
                name_box.left - bar.panel.left,
                on_the_surface(name_box),
                &style,
                palette.foreground,
                palette.hover(0.0),
            );
            boxes.push(name_box);
        }
    }

    drop(painter);

    // Every run was drawn over the whole of its own box, so the boxes are given their coverage back;
    // what is then carried across is those boxes and no more, and the marks and the panel already in
    // the window's pixels are left as they are.
    for box_ in boxes {
        seal_text(surface, on_the_surface(box_));
    }

    carry_text_across(surface, buffer, width, bar.panel);

    // The caret: one column at the place in the run the caret is at, placed where the query was really
    // drawn rather than off the field's left edge — a long query is scrolled to keep this column in
    // view, and a caret counted from the left of the field is a caret nowhere near the text. With a
    // selection it stands at the focus end of it, which is where the next character typed will land.
    //
    // It is the text's own line and not the row's: `cell` is the run's height and `line_top` its top,
    // so the column spans that box, inset a pixel at either end so that it sits inside the line rather
    // than on its edges. It is drawn solid, at full coverage and in the accent, while its blink is on:
    // the field is the one caret a pinned page has, and it blinks at the rate the system's own caret
    // blinks rather than putting a static page on the frame cadence for ever.
    //
    // It is put down after the runs have been carried across and not before: a run's box is carried as
    // whole pixels, so a caret painted under one is a column of panel colour again. And it is a column
    // of the field: the field's own last column is the furthest right it may stand, so a caret at the
    // end of a run exactly as wide as its row is inside the row rather than over the count beside it.
    if state.caret_lit {
        let caret_width = text_paint::scaled(1, scale).max(1);
        let inset = text_paint::scaled(1, scale).max(1);
        let caret_x = (origin + columns.get(caret).copied().unwrap_or(measured))
            .min(bar.query.right - caret_width);
        if caret_x >= bar.query.left {
            fill_box(
                buffer,
                width,
                RECT {
                    left: caret_x,
                    top: line_top + inset,
                    right: caret_x + caret_width,
                    bottom: (line_top + cell - inset).max(line_top + inset + 1),
                },
                palette.accent,
                1.0,
            );
        }
    }
}

/// The character index of a byte index: a caret is a place in a UTF-8 string, and a column is a place
/// among its characters, so a paint or a press walks the one to name the other. A byte index that has
/// somehow landed inside a character is read as the boundary before it rather than panicking.
fn char_index(text: &str, byte: usize) -> usize {
    text.char_indices().take_while(|(at, _)| *at < byte).count()
}

/// One chevron pointing up or down, the mark the two step buttons carry: two arms closing on the one
/// apex the mark points at, drawn as strokes for the reason the caption's sideways chevrons are (see
/// `draw_chevron`).
fn draw_step_mark(
    buffer: &mut [u8],
    width: i32,
    center_x: i32,
    center_y: i32,
    scale: f32,
    color: [u8; 3],
    pointing_up: bool,
) {
    let (mark_width, mark_height) = step_mark_shape(scale);
    // One width in every direction rather than the caption's diagonals, which are narrowed by the root
    // of two so that a walk of columns comes out at their weight: there is no walk here, and a mark
    // drawn with `stroke_segment` is the width it is asked for.
    let thickness = (FIND_BAR_MARK_STROKE * scale).max(1.0);

    for (far, apex) in step_mark_legs(center_x, center_y, mark_width, mark_height, pointing_up) {
        // The two legs meet at one point: the stroke rounds that shared end into a join and the two
        // free ends into caps, which is the whole of what a chevron needs (see `stroke_segment`).
        stroke_segment(buffer, width, far, apex, thickness, color, 1.0);
    }
}

/// How big a step mark is at `scale`: as wide as it is asked for, and a whole even number of pixels
/// tall.
///
/// Even is what keeps the apex on a pixel centre: an odd height that had its box centred on the button
/// would put the apex on a boundary between two pixels, where the join is drawn as two half-lit ones.
pub(super) fn step_mark_shape(scale: f32) -> (i32, i32) {
    let width = text_paint::scaled(FIND_BAR_CHEVRON_WIDTH, scale).max(6);
    let height = text_paint::scaled(FIND_BAR_CHEVRON_HEIGHT, scale).max(2);
    (width, ((height + 1) / 2) * 2)
}

/// The two legs a step mark is drawn from, as `(free end, apex)` pairs, for a mark `width` across and
/// `height` tall.
///
/// The apex is on a pixel centre, and the free ends are a whole number of pixels above or below it, so
/// the join and both caps land on the same grid rather than each of them a half pixel off it. The
/// mark's box is centred on the button's own middle and not its legs' middle: the chevron's mass sits
/// nearer its apex, so a mark centred by its ends is a mark that reads as hanging off the middle of
/// the button it is in.
pub(super) fn step_mark_legs(
    center_x: i32,
    center_y: i32,
    width: i32,
    height: i32,
    pointing_up: bool,
) -> [((f32, f32), (f32, f32)); 2] {
    let (middle_x, middle_y) = (center_x as f32 + 0.5, center_y as f32 + 0.5);
    let half = width as f32 / 2.0;
    let reach = height as f32 / 2.0;
    let (apex_y, far_y) = if pointing_up {
        (middle_y - reach, middle_y + reach)
    } else {
        (middle_y + reach, middle_y - reach)
    };

    [
        ((middle_x - half, far_y), (middle_x, apex_y)),
        ((middle_x + half, far_y), (middle_x, apex_y)),
    ]
}

/// Give the alpha byte back to a box GDI drew text over, so that what is carried across is the run
/// rather than nothing: GDI leaves it at zero, and a run drawn over the whole of its own box is the
/// box's coverage to seal (see `paint_tooltip`).
fn seal_text(surface: &DibSurface, rect: RECT) {
    let pixels = unsafe {
        std::slice::from_raw_parts_mut(
            surface_pixels(surface),
            surface.width as usize * surface.height as usize * 4,
        )
    };

    for y in rect.top.max(0)..rect.bottom.min(surface.height as i32) {
        for x in rect.left.max(0)..rect.right.min(surface.width as i32) {
            pixels[(y as usize * surface.width as usize + x as usize) * 4 + 3] = 255;
        }
    }
}

/// Put the text a card-sized surface was drawn on into the window's own pixels, at the card's corner
/// and over the panel that was drawn there first.
///
/// A pixel the run did not reach is left exactly as it was, which is what makes this the text rather
/// than a rectangle of card-coloured surface: the run's background is written over the whole of its
/// box, and its coverage is what separates the letters from the box around them. The card's panel and
/// marks are already in the buffer and none of them is touched (see `composite_text_into`, the same
/// road at the panel's own corner).
fn carry_text_across(surface: &DibSurface, out: &mut [u8], width: i32, panel: RECT) {
    let source = unsafe {
        std::slice::from_raw_parts(
            surface_pixels(surface),
            surface.width as usize * surface.height as usize * 4,
        )
    };

    for y in 0..surface.height as i32 {
        for x in 0..surface.width as i32 {
            let from = (y as usize * surface.width as usize + x as usize) * 4;
            let alpha = source[from + 3] as u32;
            if alpha == 0 {
                continue;
            }

            let (out_x, out_y) = (panel.left + x, panel.top + y);
            if out_x < 0 || out_y < 0 || out_x >= width {
                continue;
            }
            let to = (out_y as usize * width as usize + out_x as usize) * 4;
            if to + 4 > out.len() {
                continue;
            }

            let pixel = &mut out[to..to + 4];
            for channel in 0..3 {
                let over = source[from + channel] as u32;
                let under = pixel[channel] as u32;
                pixel[channel] = (over + (under * (255 - alpha)) / 255).min(255) as u8;
            }
            pixel[3] = 255;
        }
    }
}

//! The menu a pinned sound's card opens from a right-click on its window:
//! the panel that floats over the media below the card, and the rows in
//! it.
//!
//! The panel is placed at the point the press landed and drawn the way a
//! tool tip is drawn — the theme's own page colour, a hairline around it, a
//! shadow under it, and labels written through GDI onto a surface of the
//! window's own size and carried across (`paint_menu_popup`) — because a menu
//! is the same kind of thing over the media a tool tip is: a panel a hand
//! reads and then does something about.
//!
//! The rows are whatever the window asks for ([`MenuRow`]): the geometry
//! holds a row count, not a kind of row, so the one panel carries the mode
//! toggles and the seek choices without either knowing about the other. What
//! a press on a row means is the window's question and not this module's;
//! where a point answers a row is ([`menu_row_at`]).
//!
//! The main menu's panel is only as wide as its longest row plus the marks
//! and the pads, measured at the display's own scale
//! ([`menu_main_panel_width`]); the Seek row's flyout keeps a fixed width
//! ([`MENU_PANEL_WIDTH_PIXELS`]), the way it always has.

use super::bubble::composite_text_into;
use super::primitives::{
    caption_style, fill_box, fill_disc, fill_polygon, fill_round_rect, measure_text,
    stroke_round_rect, stroke_segment, surface_pixels, ChromePalette,
};
use crate::text::text_paint::{self, DibSurface};
use windows::Win32::Foundation::RECT;

/// The width of the seek flyout's panel, the height of a row, and the
/// room the panel holds its rows in, in the units a display's scale
/// multiplies. The flyout's width is a fixed one rather than one measured
/// from the labels, for the same reason the volume popup's is: a panel
/// that is one size every time it opens is a panel a hand learns, and the
/// longest label the flyout carries fits the width with room to spare.
/// The main menu's own panel is measured from its rows instead (see
/// `menu_main_panel_width`).
pub(super) const MENU_PANEL_WIDTH_PIXELS: f32 = 150.0;
const MENU_ROW_HEIGHT_PIXELS: f32 = 22.0;
pub(super) const MENU_PAD_PIXELS: f32 = 6.0;

/// The narrowest the main menu's own panel is ever held, whatever its rows
/// measure: the room the marks stand in and the pads on either side of it,
/// with room for several characters of the theme's own face besides, so
/// that a menu of the shortest rows still paints a panel a hand can read
/// rather than a sliver. Every label the menu carries today is wider than
/// it — the longest, "Shuffle Mode", asks for a panel past it at every
/// scale — which is why the floor never holds the real menu open (see
/// `menu_main_panel_width`).
pub(super) const MENU_PANEL_MIN_PIXELS: f32 = 64.0;

/// How far the panel is held off the band the button stands
/// in, and how round its corners are: the tooltip's own
/// radius, because the two panels are the two of the same
/// kind.
const MENU_PANEL_GAP_PIXELS: f32 = 4.0;
const MENU_RADIUS: f32 = 4.0;

/// The column the mark of a row stands in, before the labels, and
/// the size of the mark itself: a check mark a mode is turned over
/// with, or a disc the size of a point of ink, both small enough to
/// sit inside a row without being the row.
pub(super) const MENU_MARKER_ROOM_PIXELS: f32 = 16.0;
const MENU_MARKER_RADIUS_PIXELS: f32 = 2.5;
/// The span of the check mark a check row carries, drawn alone in
/// the marker column when the setting the row stands for is on and
/// not at all when it is off: there is no box around it.
const MENU_CHECK_PIXELS: f32 = 9.0;
/// The side of the triangle the Seek row carries at its right,
/// pointing toward the panel the row opens: small enough to sit in
/// the row's own pad without being the row.
const MENU_ARROW_PIXELS: f32 = 7.0;
/// The wash a row under the pointer carries: the same lift a
/// caption button takes under the hand, so that a menu and the
/// chrome above it read as one app.
const MENU_ROW_HOVER: f32 = 0.10;

/// Which side of a panel the arrow on its last row points to — the
/// side the panel the row opens sits on, so that the mark reads the
/// way the row behaves.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub(crate) enum MenuArrow {
    /// Nothing: the row opens nothing, so it carries no arrow.
    #[default]
    None,
    /// The panel the row opens sits to this one's right.
    Right,
    /// The panel the row opens sits to this one's left.
    Left,
}

/// Where a card's menu is, in the window's own coordinates: the panel it
/// floats in, and the rows inside it.
pub(crate) struct MenuPopup {
    /// The panel the menu is drawn in.
    pub(crate) panel: RECT,
    /// The top of the first row, which is the panel's own top moved down by
    /// the pad the panel holds its rows in.
    pub(crate) rows_top: i32,
    /// The height of one row, which every row of the panel is.
    pub(crate) row_height: i32,
    /// How many rows the panel holds: what the panel's own height was worked
    /// out from, and what a point is answered against (`menu_row_at`).
    pub(crate) rows: i32,
    /// Which way the arrow on the panel's last row points, or nothing for a
    /// panel whose rows open nothing. It is the flyout's tier's answer
    /// rather than the window's, because the tier is what knows which side
    /// the panel a row opens sits on (`menu_flyout_tier`) — which is why
    /// the main menu's own panel carries it whether the flyout is up or
    /// not, the same tier asked either way.
    pub(crate) arrow: MenuArrow,
}

/// The mark a row of a card's menu carries, which is what says
/// what the row is: a check for a setting that is on or off, a
/// disc for the one choice of several that is in force, and
/// nothing at all for a row that is a door to somewhere rather
/// than an answer.
pub(crate) enum MenuMark {
    /// Nothing: the row carries no mark at all.
    None,
    /// The check a row carries where the setting it stands for is
    /// on; a row whose setting is off carries no mark at all, so the
    /// mark is the check alone with no box around it.
    Check(bool),
    /// The disc of the choice in force.
    Bullet,
}

/// One row of a card's menu: the label it is written with, and
/// the mark it carries, which is what says what the row is — a
/// check for a mode that is on or off, a disc for the seek
/// that is the one the pin plays by, and nothing at all for a
/// row that opens something rather than answers.
pub(crate) struct MenuRow {
    pub(crate) label: String,
    pub(crate) mark: MenuMark,
}

/// The sizes a menu's panels are placed by, at a display's
/// own scale: the width of a panel, the height of a row,
/// the pad a panel holds its rows in, and the gap a panel
/// is held off the panel beside it. The width is the seek
/// flyout's own fixed one — the main menu's panel is
/// measured from its rows instead (see
/// `menu_main_panel_width`), so the flyout's placements
/// and the tier they ask are the ones that read this
/// width, and the main menu's own placement asks for the
/// row height, the pad and the gap only.
fn menu_panel_sizes(width: i32, dpi: u32) -> (i32, i32, i32, i32) {
    let scale = dpi as f32 / 96.0;
    (
        text_paint::scaled(MENU_PANEL_WIDTH_PIXELS as i32, scale)
            .max(8)
            .min(width.max(8)),
        text_paint::scaled(MENU_ROW_HEIGHT_PIXELS as i32, scale).max(1),
        text_paint::scaled(MENU_PAD_PIXELS as i32, scale).max(0),
        text_paint::scaled(MENU_PANEL_GAP_PIXELS as i32, scale).max(0),
    )
}

/// The width of the main menu's own panel, at the display's own
/// scale, measured from the rows it is handed: the longest row's
/// label as the theme draws it, plus the room the marks stand in
/// and the pad the panel holds its rows in on either side of it —
/// the room the panel's own paint asks for (`paint_menu_popup`),
/// so the panel is exactly as wide as its longest row needs it
/// and no wider.
///
/// Measured the way the card's own layout measures for the window
/// buttons' band (`window_button_band`): a throwaway display
/// context, made for this one layout pass and released on its
/// drop path, with the theme font the labels are painted in. The
/// scale is the display's own — the one the panel's labels are
/// painted at — so what is measured is the room the paint draws.
/// A context that cannot be made is a panel of the floor below,
/// and a window narrower than the panel is the panel held to the
/// window's own width, which is the holding the placement clamps
/// to (`menu_popup_from_point`).
fn menu_main_panel_width(rows: &[MenuRow], dpi: u32, width: i32) -> i32 {
    let scale = dpi as f32 / 96.0;

    // The throwaway context the measurement is made on: a memory
    // DC with a single pixel of nothing drawn on it, created for
    // this one pass and released when it ends, the way the card's
    // layout measures for the window buttons' band.
    let longest = DibSurface::create(1, 1).map(|surface| {
        let style = caption_style([0, 0, 0]);
        rows.iter()
            .map(|row| measure_text(&surface, &style, &row.label, scale))
            .max()
            .unwrap_or(0)
    });

    let label = longest.unwrap_or(0);
    let marker_room = text_paint::scaled(MENU_MARKER_ROOM_PIXELS as i32, scale).max(0);
    let pads = 2 * text_paint::scaled(MENU_PAD_PIXELS as i32, scale).max(0);

    (label + marker_room + pads)
        .max(text_paint::scaled(MENU_PANEL_MIN_PIXELS as i32, scale))
        .min(width.max(8))
}

/// The menu a card's right-click opens, anchored at the point the
/// click landed in the window's own coordinates: the panel's top-left
/// is that point, held inside the window — moved left only as far as
/// the window's own right edge demands and up only as far as its
/// bottom does — so a click near an edge opens a panel a hand can
/// still reach.
///
/// The panel is placed here rather than recomputed from anything the
/// card drew, because a right-click is a point on the window rather
/// than a button: the old button-anchored menu hung off the gear's
/// own box, and the gear is gone (see `pin_menu::open_pin_menu`).
///
/// `rows` is what the panel is sized to hold, so the panel and the rows the
/// window puts in it cannot disagree about how tall the panel is, nor about
/// how wide: its own width is measured from the longest of them
/// (`menu_main_panel_width`). A window
/// too small to hold the panel at all is a panel held at its own top-left
/// corner, with the pair narrowed to fit where a flyout is up beside it
/// (see `menu_flyout_from_menu`).
///
/// The panel carries the arrow on its last row pointing the way the flyout
/// opens, from the same tier the flyout's own placement asks
/// (`menu_flyout_tier`) — so the arrow is there whether the flyout is up
/// or not, and it is the same arrow either way.
pub(crate) fn menu_popup_from_point(
    point: (i32, i32),
    width: i32,
    height: i32,
    dpi: u32,
    rows: &[MenuRow],
) -> MenuPopup {
    let (_flyout_width, row_height, pad, _gap) = menu_panel_sizes(width, dpi);

    // The main menu's own width is measured from the rows it
    // holds, not the flyout's fixed one: the panel is only as
    // wide as its longest row plus the marks and the pads (see
    // `menu_main_panel_width`).
    let panel_width = menu_main_panel_width(rows, dpi, width);

    let panel_height = (pad * 2 + rows.len() as i32 * row_height).max(1);

    // The panel's top-left is the point the menu was asked for, held
    // between the window's own edges: a left or top past the window is
    // held at it, and one that would run the panel off the right or the
    // bottom is moved back by the room it lacks. A window smaller than
    // the panel leaves the room a negative number, which is the
    // panel held at the window's own top-left corner.
    let left = point.0.min(width.saturating_sub(panel_width)).max(0);
    let top = point.1.min(height.saturating_sub(panel_height)).max(0);

    let placed = MenuPopup {
        panel: RECT {
            left,
            top,
            right: left + panel_width,
            bottom: top + panel_height,
        },
        rows_top: top + pad,
        row_height,
        rows: rows.len() as i32,
        arrow: MenuArrow::None,
    };

    // Which way the Seek row's arrow points: the tier the flyout's own
    // placement asks (`menu_flyout_tier`), of the panel as it stands
    // here. The panel's placement is this function's alone — the tier's
    // narrowing of the main menu stays a flyout-up concern — but the
    // arrow is carried either way, so the Seek row points the way its
    // flyout opens whether the flyout is up or not, and the answer the
    // flyout's placement gives is the one the menu carried without it.
    let (_menu_left, _menu_right, _flyout_left, arrow) =
        menu_flyout_tier(&placed, width, dpi);

    MenuPopup { arrow, ..placed }
}

/// Where the seek flyout is, in the window's own
/// coordinates: the main menu's panel as the flyout's
/// own placement leaves it, and the flyout's panel
/// beside it.
pub(crate) struct MenuFlyout {
    /// The main menu's panel, as `menu_popup_from_point`
    /// placed it — left where the anchor put it, moved only
    /// where the window is too narrow to hold the two panels
    /// at their own width, which is the one case the two are
    /// narrowed to fit. Its last row carries the arrow saying
    /// which side the flyout ended up on.
    pub(crate) menu: MenuPopup,
    /// The flyout's own panel: beside the main menu — right
    /// of its right edge where it fits there, left of its
    /// left edge where it does not — top-aligned with the
    /// seek row.
    pub(crate) popup: MenuPopup,
}

/// Where the flyout's tier puts the two panels' own edges, and which way
/// the Seek row's arrow points because of it: the flyout's first place is
/// a small gap right of the main menu's own right edge; where that does
/// not fit it flips to a small gap left of the menu's own left edge; and
/// where neither fits, the two are narrowed to half the room the window
/// leaves between its own edges — the menu at the window's left edge and
/// the flyout at its right, which points the arrow right.
///
/// The one computation of that tier, asked both by the main menu's own
/// placement (`menu_popup_from_point`) and by the flyout's
/// (`menu_flyout_from_menu`), so that the arrow the main menu's panel
/// carries is always the side its flyout opens on, whether the flyout is
/// up or not: the main menu's own placement takes the arrow from this
/// and leaves the panel where it stands — the narrowing the tier answers
/// for the main menu being a flyout-up concern — and the flyout's own
/// placement applies the whole answer.
fn menu_flyout_tier(
    menu: &MenuPopup,
    width: i32,
    dpi: u32,
) -> (i32, i32, i32, MenuArrow) {
    let (panel_width, _row_height, _pad, gap) = menu_panel_sizes(width, dpi);

    if menu.panel.right + gap + panel_width <= width {
        (
            menu.panel.left,
            menu.panel.right,
            menu.panel.right + gap,
            MenuArrow::Right,
        )
    } else if menu.panel.left - gap - panel_width >= 0 {
        (
            menu.panel.left,
            menu.panel.right,
            menu.panel.left - gap - panel_width,
            MenuArrow::Left,
        )
    } else {
        let half = ((width - gap) / 2).max(8).min(width.max(8));
        (0, half, width - half, MenuArrow::Right)
    }
}

/// The seek flyout the menu's Seek row opens, placed from
/// the main menu the row belongs to: a small gap right of
/// the main menu's own right edge, top-aligned with the
/// Seek row — the main menu's third row — so the menu the
/// flyout came from stays wholly visible beside it.
///
/// Where the flyout does not fit to the right of the menu,
/// it flips to the menu's left, a small gap left of the
/// menu's own left edge with the same top alignment — the
/// standard Windows behavior — and the arrow on the Seek
/// row flips with it. The menu is anchored where it stands
/// in both, because the point the right-click landed is
/// the user's own and only a window too narrow for the
/// panel at all moves it.
///
/// A window too narrow for the flyout on either side is
/// the two panels narrowed to half the room the window
/// leaves between its own edges — the menu at the window's
/// left edge and the flyout right of it. Either way the
/// menu stays wholly visible and uncovered, and the flyout
/// stays inside the window to be pressed, which is the
/// same discipline the menu's own panel is placed by
/// (`menu_popup_from_point`): never off an edge, never
/// past the bottom, narrowed if the window is narrower
/// than the panel.
pub(crate) fn menu_flyout_from_menu(
    menu: &MenuPopup,
    width: i32,
    height: i32,
    dpi: u32,
    flyout_rows: &[MenuRow],
) -> MenuFlyout {
    let (panel_width, row_height, pad, _gap) = menu_panel_sizes(width, dpi);

    // The flyout hangs from the Seek row, which is the
    // main menu's third row: its top is that row's own
    // top, moved up only as far as the window's own
    // bottom makes it and never past the top — the holding
    // the menu's own panel has, because a flyout off the
    // bottom is rows a hand cannot read.
    let flyout_height = (pad * 2 + flyout_rows.len() as i32 * row_height).max(1);
    let top = (menu.rows_top + 2 * menu.row_height)
        .min(height.saturating_sub(flyout_height).max(0))
        .max(0);

    // Where the two panels' own edges go, and which way the
    // Seek row's arrow points because of it: the tier the
    // main menu's own placement asks for the same arrow
    // (`menu_flyout_tier`), applied whole here — the main
    // menu narrowed where the window is too narrow for the
    // two panels at their own width, the flyout at the
    // tier's own place beside it — so the arrow does not
    // move when the flyout comes up.
    let (menu_left, menu_right, flyout_left, arrow) = menu_flyout_tier(menu, width, dpi);
    // The flyout's own width is the room the window leaves
    // it, which is the panel's own width everywhere but the
    // narrowed pair, where it is the half the window leaves.
    let flyout_width = (width - flyout_left).min(panel_width);

    MenuFlyout {
        menu: MenuPopup {
            panel: RECT {
                left: menu_left,
                top: menu.panel.top,
                right: menu_right,
                bottom: menu.panel.bottom,
            },
            rows_top: menu.rows_top,
            row_height: menu.row_height,
            rows: menu.rows,
            arrow,
        },
        popup: MenuPopup {
            panel: RECT {
                left: flyout_left,
                top,
                right: flyout_left + flyout_width,
                bottom: top + flyout_height,
            },
            rows_top: top + pad,
            row_height,
            rows: flyout_rows.len() as i32,
            arrow: MenuArrow::None,
        },
    }
}

/// Whether a point in the window lies in the gap between the two
/// panels of a menu: the room a pointer crosses on its way between the
/// main menu and its flyout, which keeps the flyout up while the
/// pointer is in it rather than making the hand re-find the Seek row.
///
/// The gap is the room between the panels' two facing edges, and it
/// runs from the top of whichever of the Seek row and the flyout
/// begins higher to the bottom of whichever ends lower, so that a
/// pointer crossing at the row the flyout is hung from is in it even
/// where a short window has moved the flyout off that row.
pub(crate) fn menu_flyout_gap_holds(menu: &MenuPopup, flyout: &MenuPopup, x: i32, y: i32) -> bool {
    let (left, right) = if flyout.panel.left >= menu.panel.right {
        (menu.panel.right, flyout.panel.left)
    } else {
        (flyout.panel.right, menu.panel.left)
    };

    let seek_top = menu.rows_top + (menu.rows - 1) * menu.row_height;
    let seek_bottom = seek_top + menu.row_height;
    let top = flyout.panel.top.min(seek_top);
    let bottom = flyout.panel.bottom.max(seek_bottom);

    x >= left && x <= right && y >= top && y <= bottom
}

/// Which row of a menu a point is, or nothing at all where the point is not
/// inside the rows: the pad the panel holds its rows in is no row, and neither
/// is anywhere outside the panel, above it or below it — a press there is a
/// press outside the menu, which is what closes it.
pub(crate) fn menu_row_at(popup: &MenuPopup, x: i32, y: i32) -> Option<usize> {
    if x < popup.panel.left || x >= popup.panel.right {
        return None;
    }

    let rows_bottom = popup.rows_top + popup.rows * popup.row_height;
    if y < popup.rows_top || y >= rows_bottom {
        return None;
    }

    Some(((y - popup.rows_top) / popup.row_height.max(1)) as usize)
}

/// Paint a card's menu into the window's own pixels, over what is behind it:
/// the panel, the wash under the row the pointer is on, and the rows' labels
/// and marks in it.
///
/// The panel and its shadow are put down straight into the buffer, because
/// that is what they are: a thing drawn over the picture, with the picture
/// behind it. The labels are not, because GDI needs a device context — so
/// they are drawn onto a surface of the window's own size, which the caller
/// keeps blanked between paints, and then carried across the panel they were
/// written over, which is the only place they are wanted (`composite_text_into`).
/// It is the same road a tool tip's name takes, because a menu's labels are
/// the same kind of thing a tool tip's name is.
///
/// `hover` is the row the pointer is on, if it is on one, which is painted
/// under a wash of the theme's own: it is the same tick that decides where
/// the pointer is and what the wash is drawn from (see `refresh_pin_menu`).
pub(crate) fn paint_menu_popup(
    buffer: &mut [u8],
    width: i32,
    palette: &ChromePalette,
    popup: &MenuPopup,
    rows: &[MenuRow],
    hover: Option<usize>,
    surface: &DibSurface,
    scale: f32,
) {
    let panel = popup.panel;
    if panel.right <= panel.left || panel.bottom <= panel.top || rows.is_empty() {
        return;
    }

    let radius = (MENU_RADIUS * scale).max(1.0);
    let shadow = (2.0 * scale).round().max(1.0) as i32;

    // The shadow, the panel and its hairline are the tooltip's own three, so a
    // menu reads as the same kind of floating panel a tool tip is: a flat face
    // of the theme's own page colour rather than a shaded one, because the
    // labels are laid over it in a flat colour of their own; a hairline that
    // keeps it off a picture of a colour close to its own; and a shadow under
    // it that makes it float over the media rather than sit in it.
    fill_round_rect(
        buffer,
        width,
        RECT {
            left: panel.left,
            top: panel.top + shadow,
            right: panel.right,
            bottom: panel.bottom + shadow,
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

    let pad = text_paint::scaled(MENU_PAD_PIXELS as i32, scale).max(0);
    let marker_room = text_paint::scaled(MENU_MARKER_ROOM_PIXELS as i32, scale).max(0);
    let label_left = panel.left + pad + marker_room;
    let marker = panel.left as f32 + pad as f32 + marker_room as f32 / 2.0;
    let marker_radius = (MENU_MARKER_RADIUS_PIXELS * scale).max(0.5);

    let style = caption_style(palette.foreground);
    let cell = text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[text_paint::BODY_LEVEL as usize],
        scale,
    );

    // The two colours a row's own band is drawn in: the panel's own page
    // where the pointer is elsewhere, and the theme's wash where it is on
    // the row. The wash is the label's background as well as the row's, so
    // that the run carried across from the surface lands on the wash rather
    // than punching the panel's page through it.
    let base = palette.hover(0.0);
    let wash = palette.hover(MENU_ROW_HOVER);

    // Each row's own box, which its label is drawn in: the label is centred
    // in the row's own height, held to the room the marker leaves it and to
    // the panel's own far pad, so a label longer than the room is cut by the
    // box it is drawn in rather than running past the panel's edge.
    let mut boxes = Vec::with_capacity(rows.len());
    for (index, row) in rows.iter().enumerate() {
        let row_top = popup.rows_top + index as i32 * popup.row_height;
        let top = row_top + ((popup.row_height - cell) / 2).max(0);
        let measured = measure_text(surface, &style, &row.label, scale);
        boxes.push(RECT {
            left: label_left,
            top,
            right: (label_left + measured.max(0)).min(panel.right - pad),
            bottom: (top + cell).min(panel.bottom),
        });

        // The row under the pointer is washed before anything is drawn in
        // it, so that its label, its mark and its arrow are all read over
        // the wash rather than under it.
        if hover == Some(index) {
            fill_box(
                buffer,
                width,
                RECT {
                    left: panel.left,
                    top: row_top,
                    right: panel.right,
                    bottom: row_top + popup.row_height,
                },
                wash,
                1.0,
            );
        }
    }

    let mut painter = text_paint::RunPainter::new(surface, scale);
    for (index, (box_, row)) in boxes.iter().zip(rows).enumerate() {
        // The label's background is the row's own colour, so that where the
        // run is not a letter it is exactly what the row already is.
        painter.draw(
            &row.label,
            box_.left,
            *box_,
            &style,
            palette.foreground,
            if hover == Some(index) { wash } else { base },
        );

        // The mark of the row, in the room the labels
        // are held away from, which is the room a menu
        // keeps for the answer to "what is this row?".
        // It is drawn straight into the panel rather than
        // through GDI, because it is a shape and not a
        // letter, and the panel is already drawn under it.
        let row_top = popup.rows_top + index as i32 * popup.row_height;
        let row_middle = row_top as f32 + (popup.row_height / 2) as f32 + 0.5;
        match row.mark {
            MenuMark::None => {}
            MenuMark::Check(true) => {
                paint_menu_check(buffer, width, marker, row_middle, scale, palette.foreground)
            }
            // A setting that is off carries nothing at all: the check alone
            // says whether it is on, and an empty box is not drawn.
            MenuMark::Check(false) => {}
            MenuMark::Bullet => fill_disc(
                buffer,
                width,
                marker,
                row_middle,
                marker_radius,
                palette.foreground,
                1.0,
            ),
        }
    }

    // The arrow on the panel's last row, at the row's own right, pointing
    // the way the panel the row opens sits: the row that opens the flyout
    // is the panel's last, and the flyout's tier is what knows which side
    // the flyout sits on (`menu_flyout_tier`) — which is why the main
    // menu's panel carries the arrow whether the flyout is up or not.
    if popup.arrow != MenuArrow::None {
        let index = (popup.rows - 1).max(0);
        let row_top = popup.rows_top + index * popup.row_height;
        let center_y = row_top as f32 + (popup.row_height / 2) as f32 + 0.5;
        let half = (MENU_ARROW_PIXELS * scale / 2.0).max(1.5);
        let center_x = panel.right as f32 - pad as f32 - half;
        let points: [(f32, f32); 3] = match popup.arrow {
            MenuArrow::Right => [
                (center_x - half, center_y - half),
                (center_x - half, center_y + half),
                (center_x + half, center_y),
            ],
            MenuArrow::Left => [
                (center_x + half, center_y - half),
                (center_x + half, center_y + half),
                (center_x - half, center_y),
            ],
            MenuArrow::None => [(0.0, 0.0); 3],
        };
        fill_polygon(buffer, width, &points, palette.foreground, 1.0);
    }

    // GDI leaves the alpha byte of everything it draws at zero (see the module
    // documentation), and a run is drawn over the whole of its box and not
    // only over its letters, so each box is given its own coverage back here.
    // It is the same sealing the tooltip does for its one name, done for each
    // row's own box, and done on the surface of its own rather than on the
    // window's because this text is not part of the window — it is a panel
    // over the media underneath it.
    let pixels = unsafe {
        std::slice::from_raw_parts_mut(
            surface_pixels(surface),
            surface.width as usize * surface.height as usize * 4,
        )
    };
    for box_ in &boxes {
        for y in box_.top.max(0)..box_.bottom.min(surface.height as i32) {
            for x in box_.left.max(0)..box_.right.min(surface.width as i32) {
                pixels[(y as usize * surface.width as usize + x as usize) * 4 + 3] = 255;
            }
        }
    }

    // And the runs are carried across onto the panel they were written over in
    // the surface, which is the only place they are wanted: the surface is the
    // size of the whole window and the rest of it is blank memory, which over
    // the media would be a sheet of nothing.
    composite_text_into(surface, buffer, width, panel);
}

/// The check a check row is marked with: a small check mark at the
/// row's own middle, in the room the labels are held away from — no
/// box around it, because the row that is on is read from the check
/// alone and a row that is off carries nothing at all.
///
/// Two strokes one logical pixel thick, drawn the way the window
/// buttons' own marks are: from the mark's left shoulder down to its
/// middle and up to its right shoulder.
fn paint_menu_check(
    buffer: &mut [u8],
    width: i32,
    center_x: f32,
    center_y: f32,
    scale: f32,
    ink: [u8; 3],
) {
    let side = (MENU_CHECK_PIXELS * scale).round().max(5.0);
    let third = side / 3.0;
    let stroke = text_paint::scaled(1, scale) as f32;

    stroke_segment(
        buffer,
        width,
        (center_x - third, center_y),
        (center_x, center_y + third),
        stroke,
        ink,
        1.0,
    );
    stroke_segment(
        buffer,
        width,
        (center_x, center_y + third),
        (center_x + third, center_y - third),
        stroke,
        ink,
        1.0,
    );
}

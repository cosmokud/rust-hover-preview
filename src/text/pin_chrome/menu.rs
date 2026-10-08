//! The menu a pinned sound's card opens from its gear: the panel that
//! floats over the media below the card, and the rows in it.
//!
//! The panel is placed by the button the press came from and drawn the way a
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

use super::bubble::composite_text_into;
use super::primitives::{
    caption_style, fill_disc, fill_round_rect, measure_text, stroke_box, stroke_round_rect,
    stroke_segment, surface_pixels, ChromePalette,
};
use crate::text::text_paint::{self, DibSurface};
use windows::Win32::Foundation::RECT;

/// The width of the panel, the height of a row, and the room the panel holds
/// its rows in, in the units a display's scale multiplies. The width is a
/// fixed one rather than one measured from the labels, for the same reason the
/// volume popup's is: a panel that is one size every time it opens is a panel
/// a hand learns, and the longest label the menu carries fits the width with
/// room to spare.
const MENU_PANEL_WIDTH_PIXELS: f32 = 150.0;
const MENU_ROW_HEIGHT_PIXELS: f32 = 22.0;
const MENU_PAD_PIXELS: f32 = 6.0;

/// How far the panel hangs below the band the button stands
/// in, and how round its corners are: the tooltip's own
/// radius, because the two panels are the two of the same
/// kind.
const MENU_PANEL_GAP_PIXELS: f32 = 4.0;
const MENU_RADIUS: f32 = 4.0;

/// The column the mark of a row stands in, before the labels, and
/// the size of the mark itself: a checkbox a mode is turned over
/// with, or a disc the size of a point of ink, both small enough to
/// sit inside a row without being the row.
const MENU_MARKER_ROOM_PIXELS: f32 = 16.0;
const MENU_MARKER_RADIUS_PIXELS: f32 = 2.5;
/// The side of the checkbox a check row is marked with: a square
/// big enough for the check drawn in it to be read, and small
/// enough to sit in the room the marks stand in.
const MENU_CHECKBOX_PIXELS: f32 = 9.0;

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
}

/// The mark a row of a card's menu carries, which is what says
/// what the row is: a check for a setting that is on or off, a
/// disc for the one choice of several that is in force, and
/// nothing at all for a row that is a door to somewhere rather
/// than an answer.
pub(crate) enum MenuMark {
    /// Nothing: the row carries no mark at all.
    None,
    /// A checkbox, checked where the setting the row stands
    /// for is on, and empty where it is off.
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
/// is held off the button it came from and off the panel
/// beside it.
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

/// The menu a card's gear opens, hung from that button's own
/// box in the window's own coordinates — the box the card's
/// own arithmetic answers for the gear, which reaches the
/// window's own top edge above the drawn glyph and down to
/// the bottom of the band the window buttons stand in. The
/// panel hangs below that band rather than over it, with its
/// right edge tucked to the button's own left edge a small
/// gap off it — the placement that leaves a flyout room to
/// the panel's right — and its rows held in from its ends.
///
/// `rows` is what the panel is sized to hold, so the panel and the rows the
/// window puts in it cannot disagree about how tall the panel is. The panel
/// is kept inside the window it belongs to — a panel off the side of a window
/// is a row a hand cannot reach, and one off the bottom is rows that cannot be
/// read at all — so a window too narrow to hold the panel beside the button is
/// a panel narrowed to the window and held at its own left (because a panel
/// running off the side of a window is rows no hand can read), and one too
/// short to hold the panel below the button is a panel held at the window's
/// own bottom, where the hand that opened it can still reach it.
pub(crate) fn menu_popup_from_button(
    button: RECT,
    width: i32,
    height: i32,
    dpi: u32,
    rows: &[MenuRow],
) -> MenuPopup {
    let (panel_width, row_height, pad, gap) = menu_panel_sizes(width, dpi);

    let panel_height = (pad * 2 + rows.len() as i32 * row_height).max(1);

    // The right edge is tucked to the button's own left edge, a
    // small gap off it, so a flyout has room to the panel's
    // right. It is moved right only as far as the window's own
    // left edge makes it — a window too narrow to hold the panel
    // beside the button is a panel held at the window's own
    // left edge — and never past the window's own right.
    let right = button
        .left
        .saturating_sub(gap)
        .clamp(panel_width, width.max(panel_width));
    let left = right - panel_width;
    // The panel hangs below the band the button stands in rather
    // than over it — the way the volume popup floats clear of the
    // button that opened it — moved up only as far as the
    // window's own bottom makes it, and never past the top.
    let top = (button.bottom + gap)
        .min(height.saturating_sub(panel_height).max(0))
        .max(0);

    MenuPopup {
        panel: RECT {
            left,
            top,
            right,
            bottom: top + panel_height,
        },
        rows_top: top + pad,
        row_height,
        rows: rows.len() as i32,
    }
}

/// Where the seek flyout is, in the window's own
/// coordinates: the main menu's panel as the flyout's
/// own placement leaves it, and the flyout's panel
/// beside it.
pub(crate) struct MenuFlyout {
    /// The main menu's panel, as `menu_popup_from_button`
    /// placed it — pulled left of the button where the
    /// window is too narrow to hold the flyout beside it
    /// there, so that the flyout fits beside it at the
    /// window's right edge instead.
    pub(crate) menu: MenuPopup,
    /// The flyout's own panel: right of the main menu's
    /// right edge, top-aligned with the seek row.
    pub(crate) popup: MenuPopup,
}

/// The seek flyout the menu's Seek row opens, placed from
/// the main menu the row belongs to: right of the main
/// menu's own right edge, top-aligned with the Seek row —
/// the main menu's third row — so the menu the flyout
/// came from stays wholly visible beside it.
///
/// The two panels are placed together, because the button
/// the menu hangs from stands in the window's top corner,
/// where a menu tucked to it leaves a flyout no room but
/// the corner's own: a window too narrow for the flyout
/// beside the menu at the button's side is a menu pulled
/// left, off the button, until the flyout fits beside it
/// at the window's own right edge, and a window too narrow
/// for the two panels at their own width is the two of
/// them narrowed to half the room the window leaves
/// between its own edges — the menu at the window's left
/// edge, the flyout right of it. Either way the menu
/// stays wholly visible and uncovered, and the flyout
/// stays inside the window to be pressed, which is the
/// same discipline the menu's own panel is placed by
/// (`menu_popup_from_button`): never off an edge, never
/// past the bottom, narrowed if the window is narrower
/// than the panel.
pub(crate) fn menu_flyout_from_menu(
    menu: &MenuPopup,
    width: i32,
    height: i32,
    dpi: u32,
    flyout_rows: &[MenuRow],
) -> MenuFlyout {
    let (panel_width, row_height, pad, gap) = menu_panel_sizes(width, dpi);

    // The flyout hangs from the Seek row, which is the main
    // menu's third row: its top is that row's own top, moved
    // up only as far as the window's own bottom makes it and
    // never past the top — the holding the menu's own panel
    // has, because a flyout off the bottom is rows a hand
    // cannot read.
    let flyout_height = (pad * 2 + flyout_rows.len() as i32 * row_height).max(1);
    let top = (menu.rows_top + 2 * menu.row_height)
        .min(height.saturating_sub(flyout_height).max(0))
        .max(0);

    // Where the two panels' own edges go. The flyout's
    // first place is a small gap right of the main menu's
    // right edge; the main menu is pulled left of the
    // button to make that room where the window is too
    // narrow for it there; and the two are narrowed to fit
    // the window where it is too narrow for them at their
    // own width.
    let (menu_left, menu_right, flyout_left) = if menu.panel.right + gap + panel_width <= width {
        (menu.panel.left, menu.panel.right, menu.panel.right + gap)
    } else if width >= 2 * panel_width + gap {
        (
            width - 2 * panel_width - gap,
            width - panel_width - gap,
            width - panel_width,
        )
    } else {
        let half = ((width - gap) / 2).max(8).min(width.max(8));
        (0, half, width - half)
    };
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
        },
    }
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
/// the panel, and the rows' labels and marks in it.
///
/// The panel and its shadow are put down straight into the buffer, because
/// that is what they are: a thing drawn over the picture, with the picture
/// behind it. The labels are not, because GDI needs a device context — so
/// they are drawn onto a surface of the window's own size, which the caller
/// keeps blanked between paints, and then carried across the panel they were
/// written over, which is the only place they are wanted (`composite_text_into`).
/// It is the same road a tool tip's name takes, because a menu's labels are
/// the same kind of thing a tool tip's name is.
pub(crate) fn paint_menu_popup(
    buffer: &mut [u8],
    width: i32,
    palette: &ChromePalette,
    popup: &MenuPopup,
    rows: &[MenuRow],
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
    }

    let mut painter = text_paint::RunPainter::new(surface, scale);
    for (index, (box_, row)) in boxes.iter().zip(rows).enumerate() {
        // The label's background is the panel's own colour, so that where the
        // run is not a letter it is exactly what the panel already is.
        painter.draw(
            &row.label,
            box_.left,
            *box_,
            &style,
            palette.foreground,
            palette.hover(0.0),
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
            MenuMark::Check(on) => paint_menu_check(
                buffer,
                width,
                marker,
                row_middle,
                on,
                scale,
                palette.foreground,
            ),
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

/// The checkbox a check row is marked with: a square at the
/// row's own middle, in the room the labels are held away
/// from, its outline a stroke one logical pixel thick like
/// the window buttons' own marks, with the check drawn in
/// when the setting the row stands for is on and the square
/// left empty when it is off — the empty square says as much
/// as the checked one, which is what a check a hand can read
/// is.
fn paint_menu_check(
    buffer: &mut [u8],
    width: i32,
    center_x: f32,
    center_y: f32,
    on: bool,
    scale: f32,
    ink: [u8; 3],
) {
    let side = (MENU_CHECKBOX_PIXELS * scale).round().max(5.0);
    let half = side / 2.0;
    let square_left = (center_x - half).round() as i32;
    let square_top = (center_y - half).round() as i32;
    let square = RECT {
        left: square_left,
        top: square_top,
        right: square_left + side as i32,
        bottom: square_top + side as i32,
    };

    // The square's outline, drawn inside its own edges so
    // that the box is the size it was asked for rather than
    // a stroke wider on each side.
    let stroke = text_paint::scaled(1, scale) as f32;
    stroke_box(buffer, width, square, stroke, ink, 1.0);

    if !on {
        return;
    }

    // The check: two strokes from the box's own left shoulder
    // down to its middle and up to its right shoulder, each
    // the same one logical pixel the box around them is drawn
    // with.
    let third = side / 3.0;
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

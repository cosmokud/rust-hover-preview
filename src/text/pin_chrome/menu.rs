//! The menu a pinned sound's card opens from its bullet: the panel that
//! floats over the media below the card, and the rows in it.
//!
//! The panel is placed by the cell the press came from and drawn the way a
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
    caption_style, fill_disc, fill_round_rect, measure_text, stroke_round_rect, surface_pixels,
    ChromePalette,
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

/// How far the panel hangs below the bullet's own row, and how round its
/// corners are: the tooltip's own radius, because the two panels are the two
/// of the same kind.
const MENU_PANEL_GAP_PIXELS: f32 = 4.0;
const MENU_RADIUS: f32 = 4.0;

/// The column the mark of the row in force stands in, before the labels, and
/// the size of the mark itself: a disc of the ink, small enough to sit inside
/// a row without being the row.
const MENU_MARKER_ROOM_PIXELS: f32 = 16.0;
const MENU_MARKER_RADIUS_PIXELS: f32 = 2.5;

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

/// One row of a card's menu: the label it is written with, and whether it is
/// the row in force — which is what the mark before the label says, for a mode
/// that is on and for the seek that is the one the pin plays by.
pub(crate) struct MenuRow {
    pub(crate) label: String,
    pub(crate) marked: bool,
}

/// The menu a card's bullet opens, hung from that bullet's own box in the
/// window's own coordinates: a panel below the bullet's row, with its left
/// edge on the bullet's own left edge and its rows held in from its ends.
///
/// `rows` is what the panel is sized to hold, so the panel and the rows the
/// window puts in it cannot disagree about how tall the panel is. The panel
/// is kept inside the window it belongs to — a panel off the side of a window
/// is a row a hand cannot reach, and one off the bottom is rows that cannot be
/// read at all — so a window too narrow to hold the panel from the bullet's
/// edge is a panel held at the window's own left (and a window narrower than
/// the panel itself is a panel narrowed to it, because a panel running off the
/// side of a window is rows no hand can read), and one too short to hold
/// the panel below the bullet is a panel held at the window's own bottom,
/// where the hand that opened it can still reach it.
pub(crate) fn menu_popup_from_bullet(
    bullet: RECT,
    width: i32,
    height: i32,
    dpi: u32,
    rows: &[MenuRow],
) -> MenuPopup {
    let scale = dpi as f32 / 96.0;
    let panel_width = text_paint::scaled(MENU_PANEL_WIDTH_PIXELS as i32, scale)
        .max(8)
        .min(width.max(8));
    let row_height = text_paint::scaled(MENU_ROW_HEIGHT_PIXELS as i32, scale).max(1);
    let pad = text_paint::scaled(MENU_PAD_PIXELS as i32, scale).max(0);
    let gap = text_paint::scaled(MENU_PANEL_GAP_PIXELS as i32, scale).max(0);

    let panel_height = (pad * 2 + rows.len() as i32 * row_height).max(1);

    // The left edge is the bullet's own, moved left only as far as the
    // window's own edge makes it, and never past it.
    let left = bullet
        .left
        .min(width.saturating_sub(panel_width).max(0))
        .max(0);
    // The panel hangs below the bullet's row rather than over it — the way the
    // volume popup floats clear of the button that opened it — moved up only
    // as far as the window's own bottom makes it, and never past the top.
    let top = (bullet.bottom + gap)
        .min(height.saturating_sub(panel_height).max(0))
        .max(0);

    MenuPopup {
        panel: RECT {
            left,
            top,
            right: left + panel_width,
            bottom: top + panel_height,
        },
        rows_top: top + pad,
        row_height,
        rows: rows.len() as i32,
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

        // The mark of the row that is in force: a disc of the ink in the room
        // the labels are held away from, which is the room a menu keeps for
        // the answer to "which one is this?". It is drawn straight into the
        // panel rather than through GDI, because it is a shape and not a
        // letter, and the panel is already drawn under it.
        if row.marked {
            let row_top = popup.rows_top + index as i32 * popup.row_height;
            fill_disc(
                buffer,
                width,
                marker,
                row_top as f32 + (popup.row_height / 2) as f32 + 0.5,
                marker_radius,
                palette.foreground,
                1.0,
            );
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

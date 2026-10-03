//! The small helpers the window and the hook share: the point a mouse message carries, the
//! rows copied out of a surface, and what the Explorer hook asks of a pinned window.

use super::*;

/// Copy a surface painted by `pin_chrome` into the rows of a pinned window's own surface. The
/// two are the same width — a caption is as wide as the window it is on — and what the surface
/// carries is already the form the layered window is in (see `pin_chrome`).
pub(super) fn copy_surface_rows_into(
    surface: &DibSurface,
    out: &mut [u8],
    out_width: u32,
    origin_y: u32,
) {
    let width = surface.width as usize;
    let height = surface.height as usize;
    if width == 0 || height == 0 || (out_width as usize) < width {
        return;
    }

    let row_bytes = width * 4;
    let out_row_bytes = out_width as usize * 4;
    let source = unsafe { std::slice::from_raw_parts(surface.bits(), row_bytes * height) };

    for row in 0..height {
        let start = (origin_y as usize + row) * out_row_bytes;
        let end = start + row_bytes;
        if end > out.len() {
            break;
        }
        out[start..end].copy_from_slice(&source[row * row_bytes..(row + 1) * row_bytes]);
    }
}

/// Where a scroll of `lines` from the preview's current position lands.
pub(super) fn text_scroll_target(lines: i64) -> Option<usize> {
    CURRENT_MEDIA.lock().ok().and_then(|media| {
        media
            .as_ref()
            .and_then(|media| media.text_state.as_ref())
            .map(|scroll| scroll.scrolled_by(lines))
    })
}

/// The point a mouse message was delivered at, in window coordinates.
pub(super) fn message_point(lparam: LPARAM) -> (i32, i32) {
    let x = (lparam.0 & 0xFFFF) as u16 as i16 as i32;
    let y = ((lparam.0 >> 16) & 0xFFFF) as u16 as i16 as i32;
    (x, y)
}

/// Whether a point on a pinned window is the volume button: what a press on the popup's own button
/// has to be told apart from a press anywhere else, since the one keeps the popup and the other
/// puts it away (see `pinned_press`).
///
/// A sound's own button is asked about too, and for the same reason: it is the same control drawn
/// in a different place, and a press on it that closed the popup would have the release open it
/// straight back up.
pub(super) fn pin_point_is_volume_button(x: i32, y: i32) -> bool {
    let on_the_bar = pinned_transport_geometry().is_some_and(|bar| {
        y >= bar.top
            && pin_chrome::transport_part_at(
                x,
                y - bar.top,
                bar.width,
                bar.height,
                bar.dpi,
                bar.live,
            ) == Some(pin_chrome::TransportPart::Volume)
    });

    on_the_bar
        || pin_state()
            .and_then(|pinned| {
                let pin = pinned.pin()?;
                Some(pin_audio_control_at(pin, x, y) == Some(CardControl::Volume))
            })
            .unwrap_or(false)
}

/// Whether a press on the media belongs to the pin rather than to the media.
///
/// It is the pin's for every kind and every part of the frame — a hand on the picture carries the
/// window the way a hand on a title bar does, which is what a window's body is for — with one
/// exception: a text preview has business of its own under the pointer, and the two places that
/// business is in are the scrollbar (a drag of the thumb) and the text itself (a selection). What
/// is left of a page is its margins — above the first line, below the last, and the gutters
/// either side of the column — and the margins are a handle, as the caption is.
pub(super) fn pinned_content_is_the_pins(x: i32, y: i32) -> bool {
    let (media_x, media_y) = media_point(x, y);
    if text_scroll_drag_target(media_x, media_y).is_some() {
        return false;
    }

    let Ok(media) = CURRENT_MEDIA.lock() else {
        return true;
    };
    let Some(media) = media.as_ref() else {
        return true;
    };
    let Some(state) = media.text_state.as_ref() else {
        return true;
    };

    !text_preview::point_is_on_text(&state.lines, media_x, media_y)
}

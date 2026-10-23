//! The small helpers the window and the hook share: the point a mouse message carries, and what
//! the Explorer hook asks of a pinned window.

use super::*;

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
/// exception: a page of text is worked in from edge to edge, so the whole of its frame is the
/// page's. The text is what a drag selects and the margins around it are where a caret at a line's
/// own end goes, while the scrollbar is the page's own drag; none of it is a handle that carries
/// the window (see `pinned_press`).
pub(super) fn pinned_content_is_the_pins() -> bool {
    CURRENT_MEDIA
        .lock()
        .map(|media| {
            media
                .as_ref()
                .map(|media| media.text_state.is_none())
                .unwrap_or(true)
        })
        .unwrap_or(true)
}

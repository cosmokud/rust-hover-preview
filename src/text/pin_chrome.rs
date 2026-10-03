//! The chrome a pinned preview wears: the caption above its media, and the round bubble it
//! collapses into.
//!
//! Both are painted into a surface of their own and copied onto the frame's, which is what
//! keeps the drawing of them out of the file that holds the preview window and its media: a
//! caption is a bar, a title and three glyphs, and a bubble is a circle — none of it knows
//! what a file is.
//!
//! The colors are the text preview's own theme rather than a palette of this app's (see
//! [`ChromePalette`]): a preview is a page of text, a picture, a card or a video, and the
//! chrome around one belongs to the same app whichever it turned out to be.
//!
//! Everything written here is written *premultiplied* — the color a pixel carries is its own
//! color times its coverage — which is the form the layered surface is in and the form
//! `UpdateLayeredWindow` with `AC_SRC_ALPHA` reads (see `compose_preview_row`). The one
//! exception is GDI, which knows nothing of an alpha channel: what it draws leaves the alpha
//! byte at zero, so a caption is closed by handing every pixel its coverage back.
//!
//! The chrome lives in four files below: `transport` for the bar a playing file is driven and
//! drawn from, `caption` for the strip above it, `bubble` for the two panels that float over
//! the media rather than standing in it, and `primitives` for the palette and the shapes, marks
//! and text runs the other three are drawn from. What is left in this file is the way in: every
//! name the rest of the tree reaches as `pin_chrome::`, and the imports the tests beside it read
//! the four through.

mod bubble;
mod caption;
mod primitives;
mod transport;

pub(crate) use bubble::{
    paint_bubble, paint_failure_mark, paint_tooltip, tooltip_layout, BubbleMark, TooltipText,
};
pub(crate) use caption::{
    button_at, button_boxes, measure_caption_text, paint_caption, Caption, CaptionButton,
};
pub(crate) use primitives::ChromePalette;
pub(crate) use transport::{
    paint_card_control, paint_transport, paint_volume_popup, transport_part_at, transport_share_at,
    volume_popup_from_button, volume_popup_layout, volume_share_at, ControlGlyph, TransportPart,
    TransportState, VolumePopup,
};

#[cfg(test)]
use crate::text::text_paint::{self, DibSurface};
#[cfg(test)]
use caption::{fit_title, longest_prefix_that_fits, paint_glyph, CaptionButtonBox, BUTTON_PIXELS};
#[cfg(test)]
use primitives::{caption_style, clock_text, draw_chevron, measure_text, GLYPH_PIXELS};
#[cfg(test)]
use transport::{
    draw_track_step, transport_layout, volume_thumb_row, VOLUME_COLLAR_PIXELS, VOLUME_PANEL_GAP,
    VOLUME_PANEL_WIDTH, VOLUME_THUMB_RADIUS,
};
#[cfg(test)]
use windows::Win32::Foundation::RECT;

#[cfg(test)]
mod tests;

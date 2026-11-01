//! The caret a pinned page of text draws, and the read-only pill its caption wears.
//!
//! **The caret is a mark over the composed frame, not a part of it.** A selection is a colour laid
//! under the text and belongs in the frame the page is painted into; a caret is a bar laid *over* a
//! character, and it goes away again every half second — so a frame drawn again for each phase would
//! restyle and repaint a screenful of text twice a second for a mark one pixel wide. The bar is drawn
//! into the composited band instead, from the place the page says the caret is (see
//! `text_preview::caret_rect` and `text_caret::caret_mark`).
//!
//! **The read-only pill is a caption's own.** The page is a view of a file and nothing a reader does
//! to it reaches the disk (see `text_caret`), and a caption that says so is the difference between a
//! window that ignores a keystroke and a window that is broken. The pill is a hairline ring, an
//! accent wash, a padlock drawn from this module's primitives and the words `Read only`, and it is
//! drawn at the caption's own stroke width and scale so that it belongs to the strip it stands on.

use super::primitives::{
    caption_style, fill_box, fill_round_rect, measure_text, stroke_round_rect, stroke_top_arc,
    surface_pixels, ChromePalette, GLYPH_STROKE_PIXELS,
};
use crate::text::text_paint::{self, DibSurface};
use windows::Win32::Foundation::RECT;

/// How wide the bar a page's caret is drawn is, in the units a display's scale multiplies. One
/// logical pixel: a caret is a hairline over the text rather than a mark beside it, and every editor
/// on the desktop draws one this wide.
const CARET_PIXELS: f32 = 1.0;

/// How much of the caret's own line height the bar takes. Less than the whole of it, so that the bar
/// reads as a mark between two lines rather than as a character that has been painted over.
const CARET_HEIGHT_SHARE: f32 = 0.82;

/// Draw the caret a page is holding, where the page says it stands.
///
/// `rect` is the bar's place among the painted lines of the frame (see `text_preview::caret_rect`),
/// and `origin_y` is the row of the composited band the frame's own first line lands at: the page is
/// drawn below the caption, and a caret measured in the frame's coordinates has to be carried into
/// the band's the way the frame itself is.
///
/// Nothing is drawn where the place is not one of the lines on screen, which is what a caret
/// scrolled past the edge of the box is — the page is what moves to keep a caret visible, and a bar
/// drawn at the frame's edge for one that is not on it would be a mark in the wrong place (see
/// `TextPreviewState::scroll_to_show`).
pub(crate) fn paint_text_caret(
    buffer: &mut [u8],
    width: i32,
    rect: RECT,
    origin_y: i32,
    palette: &ChromePalette,
    dpi: u32,
) {
    let scale = dpi as f32 / 96.0;
    let stroke = text_paint::scaled(CARET_PIXELS as i32, scale).max(1);

    // The bar is the line's height less a little at either end, centered on the line: a caret that
    // filled the whole of its line would touch the descenders above and the ascenders below.
    let inset = ((rect.bottom - rect.top) as f32 * (1.0 - CARET_HEIGHT_SHARE) / 2.0).round() as i32;
    let top = rect.top + origin_y + inset;
    let bottom = rect.bottom + origin_y - inset;

    fill_box(
        buffer,
        width,
        RECT {
            left: rect.left,
            top,
            right: rect.left + stroke,
            bottom,
        },
        palette.foreground,
        1.0,
    );
}

/// How tall the read-only pill is at a display's scale, and how much room it keeps beside the caption
/// buttons: the whole of what `caption::title_room` takes off for it, and the whole of what the pill
/// is drawn in.
pub(super) const READ_ONLY_PILL_PIXELS: f32 = 20.0;

/// How much room the pill keeps clear of the name beside it and of the buttons beyond it.
pub(super) const READ_ONLY_PILL_GAP_PIXELS: f32 = 8.0;

/// The words the pill says, which are the whole of what it is for: a window that ignores a keystroke
/// says so rather than leaving the reader to work out why.
const READ_ONLY_WORDS: &str = "Read only";

/// How wide the pill is at a display's scale: the padlock, the gap after it, the words, and the
/// padding at either end — measured through the caption's own face, because the words are drawn in it.
pub(super) fn read_only_pill_width(surface: &DibSurface, dpi: u32) -> i32 {
    let scale = dpi as f32 / 96.0;
    let padding = text_paint::scaled(7, scale);
    let lock = text_paint::scaled(9, scale);
    let gap = text_paint::scaled(5, scale);
    let words = measure_text(surface, &caption_style([0, 0, 0]), READ_ONLY_WORDS, scale);

    2 * padding + lock + gap + words
}

/// The pill's box inside a caption of `width`, standing against `title_right` — the left edge of the
/// leftmost caption button — so that it is the last thing said before the buttons begin.
///
/// It is `title_right`'s own answer rather than an arithmetic of its own, because the two have to be
/// one answer: a pill placed beside a group of buttons the caption was not drawn with is a pill
/// either under the buttons or a gap away from them.
pub(super) fn read_only_pill_box(
    surface: &DibSurface,
    width: i32,
    height: i32,
    title_right: i32,
    dpi: u32,
) -> RECT {
    let scale = dpi as f32 / 96.0;
    let pill = text_paint::scaled(READ_ONLY_PILL_PIXELS as i32, scale);
    let gap = text_paint::scaled(READ_ONLY_PILL_GAP_PIXELS as i32, scale);
    let right = title_right - gap;
    let left = right - read_only_pill_width(surface, dpi);
    let top = ((height - pill) / 2).max(0);

    RECT {
        left: left.max(0),
        top,
        right: right.min(width),
        bottom: (top + pill).min(height),
    }
}

/// Paint the read-only pill: a hairline ring, an accent wash, a padlock and the words.
///
/// The padlock is drawn by hand from this module's own primitives rather than set from a font or an
/// emoji, for the reason every other mark on the chrome is: an emoji is whatever the machine's own
/// emoji font draws it as, at a size and a color this app does not choose, and an icon font is a file
/// the app would have to ship and load. Two rounded boxes and an arc are the whole of a padlock.
pub(super) fn paint_read_only_pill(
    surface: &DibSurface,
    palette: &ChromePalette,
    box_: RECT,
    dpi: u32,
) {
    if box_.right <= box_.left || box_.bottom <= box_.top {
        return;
    }

    let scale = dpi as f32 / 96.0;
    let stroke = (GLYPH_STROKE_PIXELS * scale).max(1.0);
    let radius = (box_.bottom - box_.top) as f32 / 2.0;
    let buffer = unsafe {
        std::slice::from_raw_parts_mut(
            surface_pixels(surface),
            surface.width as usize * surface.height as usize * 4,
        )
    };
    let width = surface.width as i32;

    // The wash and its ring: the accent moved most of the way back toward the bar, so the pill is a
    // tint of the caption rather than a block of color on it, with a hairline of the accent itself.
    let wash = text_paint::blend(palette.background, palette.accent, 0.14);
    fill_round_rect(buffer, width, box_, radius, wash, wash, 1.0);
    stroke_round_rect(buffer, width, box_, radius, 1.0, palette.accent, 0.55);

    // The padlock stands on the pill's own left padding, its body on the strip's middle line and its
    // shackle above it: a body, a ring that is the top half of one, and the two together read as a
    // lock at the size a caption has.
    let padding = text_paint::scaled(7, scale);
    let lock = text_paint::scaled(9, scale);
    let left = box_.left + padding;
    let middle = (box_.top + box_.bottom) as f32 / 2.0;
    let body_height = (lock as f32 * 0.52).round().max(2.0);
    let body = RECT {
        left,
        top: (middle + 1.0 - body_height / 2.0) as i32,
        right: left + lock,
        bottom: (middle + 1.0 + body_height / 2.0) as i32,
    };
    let body_radius = (body_height * 0.28).max(1.0);
    stroke_round_rect(
        buffer,
        width,
        body,
        body_radius,
        stroke,
        palette.accent,
        1.0,
    );

    // The shackle: an arc standing on the body's top edge, open at the bottom, drawn a little wider
    // than the body so that it reads as a hoop rather than as a lid.
    let shackle_radius = (lock as f32 * 0.28).max(1.5);
    stroke_top_arc(
        buffer,
        width,
        (
            (body.left + body.right) as f32 / 2.0,
            body.top as f32 - shackle_radius * 0.15,
        ),
        shackle_radius,
        stroke,
        palette.accent,
        1.0,
    );

    // The words, in the caption's own face and at the caption's own size, so that the pill reads as
    // something the strip is saying rather than as a label from another program.
    let words_left = left + lock + text_paint::scaled(5, scale);
    let cell = text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[text_paint::BODY_LEVEL as usize],
        scale,
    );
    let top = ((box_.bottom - box_.top - cell) / 2 + box_.top).max(box_.top);
    let style = caption_style(palette.accent);
    let mut painter = text_paint::RunPainter::new(surface, scale);
    // **The run's box stops at the pill's own padding, not at its edge.** A run is drawn with an
    // opaque background over the whole of its box rather than over its glyphs alone (see
    // `paint_tooltip`), so a box reaching the edge paints that background over the ring's right cap
    // and leaves the pill cut off where the words stop — which is exactly what a run drawn to
    // `box_.right` did.
    painter.draw(
        READ_ONLY_WORDS,
        words_left,
        RECT {
            left: words_left,
            top,
            right: (box_.right - padding).max(words_left + 1),
            bottom: (top + cell).min(box_.bottom),
        },
        &style,
        palette.accent,
        wash,
    );
    drop(painter);

    // GDI leaves the alpha byte of what it draws at zero, and the caption seals its own coverage
    // after this is called (see `paint_caption`) — so nothing is sealed here, and the pill's own
    // pixels are the caption's to hand back.
}

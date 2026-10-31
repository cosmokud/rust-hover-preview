//! The card a pinned window shows when the player drawing its film has said the file is damaged.
//!
//! **It is a card and not a badge.** What it says is a guess the player made and this app passed on —
//! that a *file* may be broken — so it stands in the middle of the picture for a few seconds, once,
//! and gets out of the way: a mark that stayed would be a verdict on a film somebody may be watching
//! happily. Its three beats are `engines::corruption_log`'s (`CARD_FADE_IN`, `CARD_HOLD`,
//! `CARD_FADE_OUT`), and what this module draws is one frame of them.
//!
//! **It dims the picture before it names it.** A warning over a bright scene is the one thing a
//! translucent panel cannot promise to be read on, so the band is covered by a scrim of the theme's
//! own black and the panel stands on that — dimming and card fading together, because they are one
//! object rather than two.
//!
//! The surface it is painted on is the band's own box rather than the window's (see
//! `CARD_SURFACE` and `pin_chrome_strips`): a hosted film composites over the pin's layered surface,
//! so a card drawn into that surface would be a card under the film. The strip carries these pixels
//! over the picture instead, which is the same road the caption and the volume popup take.

use super::primitives::{caption_style, fill_round_rect, measure_text, stroke_round_rect};
use super::{surface_pixels, ChromePalette};
use crate::text::text_paint::{self, DibSurface};
use windows::Win32::Foundation::RECT;

/// The two things the card says: what it is, and what was seen.
pub(super) const CARD_TITLE: &str = "This video file may be corrupted";
pub(super) const CARD_DETAIL: &str = "playback errors detected";

/// How much of the picture the scrim covers, at the card's own full coverage.
const SCRIM_COVERAGE: f32 = 0.38;

/// The panel's measurements, in logical pixels — every one of them scaled by the display's DPI
/// before it is drawn, which is the one rule this file has about size (see `text_paint::scaled`,
/// and `ui::preview_window::pointer::logical_px` for the other half of the same arithmetic).
const PADDING_PIXELS: i32 = 20;
const RADIUS_PIXELS: f32 = 10.0;
/// The room between the two lines, which is what makes them one card rather than two labels.
const GAP_PIXELS: i32 = 6;
/// The margin the panel keeps off the band's own edges: a panel wider than the band less two of
/// these has stopped being a card and become a band of its own, and one the band cannot hold at
/// all is drawn not at all.
const MARGIN_PIXELS: i32 = 24;
/// How thick the panel's outline is drawn, in logical pixels.
const OUTLINE_PIXELS: f32 = 1.0;

/// Draw one frame of the card over a band, into the surface its strip will carry.
///
/// `coverage` is how much of the card is on screen this frame, from nothing to whole (see
/// `corruption_log::card_coverage`). A band too small to hold the panel is left alone rather than
/// drawn into, which is the answer the spinner gives a box it does not fit in for the same reason.
pub(crate) fn paint_corruption_card(
    surface: &DibSurface,
    palette: &ChromePalette,
    coverage: f32,
    dpi: u32,
) {
    let width = surface.width as i32;
    let height = surface.height as i32;
    if width <= 0 || height <= 0 {
        return;
    }

    let scale = dpi as f32 / 96.0;
    let padding = text_paint::scaled(PADDING_PIXELS, scale);
    let gap = text_paint::scaled(GAP_PIXELS, scale);
    let margin = text_paint::scaled(MARGIN_PIXELS, scale);
    let radius = (RADIUS_PIXELS * scale).max(1.0);
    let cell = text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[text_paint::BODY_LEVEL as usize],
        scale,
    );
    let style = caption_style(palette.foreground);
    let title_width = measure_text(surface, &style, CARD_TITLE, scale);
    let detail_width = measure_text(surface, &style, CARD_DETAIL, scale);

    // The panel is the wider of the two lines with the padding either side of it, and two cells
    // with the gap and the padding between and around them.
    let panel_width = (title_width.max(detail_width) + padding * 2).min(width - margin * 2);
    let panel_height = cell * 2 + gap + padding * 2;
    if panel_width <= padding * 2 || panel_height + margin * 2 > height {
        return;
    }

    let left = (width - panel_width) / 2;
    let top = (height - panel_height) / 2;
    let panel = RECT {
        left,
        top,
        right: left + panel_width,
        bottom: top + panel_height,
    };

    let pixels = unsafe {
        std::slice::from_raw_parts_mut(
            surface_pixels(surface),
            surface.width as usize * surface.height as usize * 4,
        )
    };

    // The scrim, over the whole band and under everything else. The surface is blank before this,
    // so a premultiplied black at the scrim's own coverage is four bytes a pixel: the colour is
    // already nothing, and only the coverage has to be written.
    let alpha = (SCRIM_COVERAGE * coverage.clamp(0.0, 1.0) * 255.0).round() as u8;
    for pixel in pixels.as_chunks_mut::<4>().0 {
        pixel[0] = 0;
        pixel[1] = 0;
        pixel[2] = 0;
        pixel[3] = alpha;
    }

    // The panel and its outline, put down opaquely: what fades is the panel's box, scaled once at
    // the end (see below), because a fill that blended at the card's coverage would have to be
    // walked again for every step of the fade — and the text on it cannot fade that way at all.
    let face = palette.hover(0.0);
    fill_round_rect(pixels, width, panel, radius, face, face, 1.0);
    stroke_round_rect(
        pixels,
        width,
        panel,
        radius,
        (OUTLINE_PIXELS * scale).round().max(1.0),
        palette.hover(0.30),
        1.0,
    );

    // The two lines, through GDI, on the surface the panel is already on: the run's background is
    // the panel's own colour, so where the run is not a letter it is exactly what the panel already
    // is, and the box each line is given is the one cell of text it draws in.
    let mut painter = text_paint::RunPainter::new(surface, scale);
    let mut draw_line = |text: &str, top: i32| {
        painter.draw(
            text,
            panel.left + padding,
            RECT {
                left: panel.left,
                top,
                right: panel.right,
                bottom: (top + cell).min(panel.bottom),
            },
            &style,
            palette.foreground,
            face,
        );
    };
    draw_line(CARD_TITLE, panel.top + padding);
    draw_line(CARD_DETAIL, panel.top + padding + cell + gap);
    drop(painter);

    // GDI leaves the alpha byte of everything it draws at zero (see the module documentation over
    // `pin_chrome`), and a run covers the whole of its box rather than only its letters — so a pixel
    // of the panel with no coverage on it is a pixel GDI wrote, and it is handed its own back. Then
    // the whole panel is faded by the card's coverage, which for a premultiplied pixel is each of
    // its four channels scaled.
    let fade = coverage.clamp(0.0, 1.0);
    for y in panel.top.max(0)..panel.bottom.min(height) {
        for x in panel.left.max(0)..panel.right.min(width) {
            let index = (y as usize * surface.width as usize + x as usize) * 4;
            let Some(pixel) = pixels.get_mut(index..index + 4) else {
                continue;
            };
            let alpha = if pixel[3] == 0 {
                255.0
            } else {
                f32::from(pixel[3])
            };
            for channel in pixel.iter_mut().take(3) {
                *channel = (f32::from(*channel) * fade).round().clamp(0.0, 255.0) as u8;
            }
            pixel[3] = (alpha * fade).round().clamp(0.0, 255.0) as u8;
        }
    }
}

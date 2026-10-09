//! The two things a pinned window draws over its media rather than standing in it: the panel a
//! caption's button is described in, and the circle a pin collapses into.
//!
//! A name and a bubble float over a picture somebody else painted, which is all they have in
//! common and the whole of why they are here together: each is placed by the window rather than
//! by the strip it belongs to, each is put down straight into the window's own pixels with the
//! picture behind it, and each carries a shadow of its own so that it reads as lying over the
//! desktop rather than as a hole cut in it.
//!
//! A bubble is the theme's own page colour with a ring of its accent around it and the picture
//! that was pinned inside it where there is one (`paint_bubble`). Where there is not, it carries a
//! mark saying what kind of thing it stands for (`BubbleMark`). The failure mark is the other
//! answer to the same question - a file a pinned window was shown and could not draw at all
//! (`paint_failure_mark`).

use super::primitives::{
    caption_style, draw_cross, fill_box, fill_disc, fill_round_rect, put, stroke_round_rect,
    surface_pixels, ChromePalette,
};
use crate::text::text_paint::{self, DibSurface};
use windows::Win32::Foundation::RECT;

/// The room a tooltip takes around its text, and how far it floats below the button it
/// belongs to, in the units a display's scale multiplies.
const TOOLTIP_PADDING_PIXELS: f32 = 8.0;
const TOOLTIP_GAP_PIXELS: f32 = 6.0;
const TOOLTIP_RADIUS: f32 = 4.0;

/// Where a caption's tooltip is, in the window's own coordinates: the panel it is drawn in,
/// which is a panel of its own rather than any part of the strip.
///
/// It is *below* the caption, over the media, and not inside the strip. A name like "Open With
/// Adobe Photoshop" is wider than the two buttons it would have to sit among, and a tooltip
/// written inside a thirty-pixel bar either covers the buttons beside it or is a box of text
/// with a row of glyphs showing through it. A tooltip that covers the buttons it is describing
/// has covered them for as long as it was up, and the hand cannot move to the one next to it
/// without waiting for it to go first.
///
/// So it hangs off the bottom edge of the strip, over the media underneath, which is the same
/// arrangement as the volume popup and for the same reason (see `volume_popup_layout`).
pub(crate) fn tooltip_layout(
    width: i32,
    height: i32,
    caption_height: i32,
    anchor: RECT,
    text_width: i32,
    dpi: u32,
) -> Option<RECT> {
    let scale = dpi as f32 / 96.0;
    let padding = text_paint::scaled(TOOLTIP_PADDING_PIXELS as i32, scale);
    let gap = text_paint::scaled(TOOLTIP_GAP_PIXELS as i32, scale);
    let cell = text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[text_paint::BODY_LEVEL as usize],
        scale,
    );

    let panel_width = (text_width + padding * 2).min(width.max(0));
    let panel_height = cell + padding;
    // The panel is hung *below* the caption, so the room it needs is measured from the bottom of
    // the strip and not from the top of the window: a guard that asked only whether the panel is
    // taller than the window would admit a panel that runs off the bottom of one, and a name cut
    // off across its middle is a name that is not a name.
    let top = caption_height + gap;
    if panel_width <= 0 || panel_height <= 0 || top + panel_height > height {
        return None;
    }

    // Centred under the button it belongs to, then pulled inside the window's own sides, and
    // hung below the caption with the gap it was given — a name touching the bar it describes
    // reads as part of the bar.
    let left =
        ((anchor.left + anchor.right - panel_width) / 2).clamp(0, (width - panel_width).max(0));

    Some(RECT {
        left,
        top,
        right: left + panel_width,
        bottom: top + panel_height,
    })
}

/// A caption's tooltip as the painter is given it: the name, how wide it was measured at, and
/// the surface it is to be written on.
///
/// The three are one thing rather than three arguments because they are one thing: a name is
/// measured on a surface and drawn on the same one, and a pair that could be given apart is a
/// pair that can be given wrong — a name measured in one font and drawn in another, or a name
/// measured on a surface that is not the one it is drawn on.
pub(crate) struct TooltipText<'a> {
    pub(crate) text: &'a str,
    /// The name's width in the face and at the scale it is drawn at, as `measure_caption_text`
    /// measured it.
    pub(crate) width: i32,
    /// A surface with a device context to draw it on, which the window's own surface is not.
    pub(crate) surface: &'a DibSurface,
}

/// Paint the name of a caption button as a panel of its own, floated over the media below the
/// strip (see `tooltip_layout` for why it is not drawn into the strip).
///
/// The panel and its outline are put down straight into the window's own pixels, because that is
/// what they are: a thing drawn over the picture, with the picture behind it. The name on top of
/// them is not, because GDI needs a device context — so it is drawn onto a surface of its own,
/// and then carried across the pixels the panel already covers, which is the only place it is
/// wanted.
pub(crate) fn paint_tooltip(
    buffer: &mut [u8],
    width: i32,
    palette: &ChromePalette,
    panel: RECT,
    name: TooltipText<'_>,
    scale: f32,
) {
    let TooltipText {
        text,
        width: text_width,
        surface,
    } = name;
    if text.is_empty() || panel.right <= panel.left || panel.bottom <= panel.top {
        return;
    }

    let padding = text_paint::scaled(TOOLTIP_PADDING_PIXELS as i32, scale);
    let radius = (TOOLTIP_RADIUS * scale).max(1.0);
    let cell = text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[text_paint::BODY_LEVEL as usize],
        scale,
    );

    // The panel is the theme's own page colour moved off the caption's, so that a name reads
    // against a surface rather than against the bar it belongs to; the hairline is what keeps
    // it off a picture of a colour close to its own. It is flat rather than shaded, because the
    // name is laid over it in a flat colour of its own and a name across a gradient has a shade
    // running through it — which is the one thing a tool tip is not.
    //
    // A shadow under it, the panel's own shape held a couple of pixels lower, is what makes it
    // float over the picture rather than sit in it, the same as the volume popup's.
    let shadow = (2.0 * scale).round().max(1.0) as i32;
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

    let style = caption_style(palette.foreground);

    // The name is measured against this same font, so the run is given a box of the cell's own
    // height and the room `text_width` leaves is the room the panel was made from. Its
    // background is the panel's own colour, so that where the run is not a letter it is exactly
    // what the panel already is.
    let top = panel.top + ((panel.bottom - panel.top - cell) / 2).max(0);
    let rect = RECT {
        left: panel.left,
        top,
        right: (panel.left + padding + text_width.max(0)).min(panel.right),
        bottom: (top + cell).min(panel.bottom),
    };
    let mut painter = text_paint::RunPainter::new(surface, scale);
    painter.draw(
        text,
        panel.left + padding,
        rect,
        &style,
        palette.foreground,
        palette.hover(0.0),
    );

    // GDI leaves the alpha byte of everything it draws at zero (see the module documentation),
    // and a run is drawn over the whole of its box and not only over its letters, so the box is
    // given its own coverage back here. It is the same sealing the caption does for its title,
    // and it is done on the surface of its own rather than on the caption's because this text is
    // not part of the caption — it is a panel over the media underneath it.
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

    // And the run is carried across onto the panel it was written over in the surface, which is
    // the only place it is wanted: the surface is the size of the whole window and the rest of
    // it is blank memory, which over the media would be a sheet of nothing.
    composite_text_into(surface, buffer, width, panel);
}

/// Put the text a surface was drawn on into the window's own pixels, over the panel it belongs
/// to, in the premultiplied form the layered surface is composited from.
///
/// A pixel the run did not reach is left exactly as it was, which is what makes this the name
/// rather than a rectangle of text-coloured panel: the run's background is written over the
/// whole of its box, and its coverage is what separates the letters from the box
/// around them. The menu's labels are carried across by the same road
/// (`pin_chrome::menu`).
pub(super) fn composite_text_into(surface: &DibSurface, out: &mut [u8], width: i32, panel: RECT) {
    let source = unsafe {
        std::slice::from_raw_parts(
            surface.bits(),
            surface.width as usize * surface.height as usize * 4,
        )
    };
    let height = surface.height as i32;

    for y in panel.top.max(0)..panel.bottom.min(height) {
        let left = panel.left.max(0);
        let from = (y as usize * surface.width as usize + left as usize) * 4;
        let to = (y as usize * width as usize + left as usize) * 4;
        let count = ((panel.right - left).max(0) as usize) * 4;
        if from + count > source.len() || to + count > out.len() {
            // A row that does not fit is a row of nothing: the rest of the panel is still
            // worth carrying, and a name that stopped halfway is worse than one that stopped
            // short of the bottom.
            continue;
        }

        for (pixel, drawn) in out[to..to + count]
            .as_chunks_mut::<4>()
            .0
            .iter_mut()
            .zip(source[from..from + count].as_chunks::<4>().0)
        {
            let alpha = drawn[3] as u32;
            if alpha == 0 {
                continue;
            }
            for channel in 0..3 {
                let over = drawn[channel] as u32;
                let under = pixel[channel] as u32;
                pixel[channel] = (over + (under * (255 - alpha)) / 255).min(255) as u8;
            }
            pixel[3] = 255;
        }
    }
}

/// The mark a bubble carries when the pin it stands for has no picture to show inside it.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum BubbleMark {
    /// Something that plays, and is playing: a play triangle.
    Play,
    /// Something that plays, and is paused: two bars.
    Pause,
    /// A page: text, an archive's listing, a document, a book.
    Page,
    /// A picture whose frame is not in hand.
    Picture,
}

impl BubbleMark {
    /// Whether this mark is a playback one — which is the one family of marks drawn *over* a
    /// picture rather than instead of one: a bubble standing for a film or a sound has something
    /// of its own behind the glyph (a frame, or the histogram of the level on screen) and the
    /// glyph says what the file is doing on top of it, where a page and a picture have nothing
    /// inside the circle but themselves (see `paint_bubble`).
    pub(crate) fn is_playback(self) -> bool {
        matches!(self, Self::Play | Self::Pause)
    }
}

/// What a bubble draws inside its ring where the pin it stands for has something to show there:
/// the picture the file is, or the level the sound on screen is being heard at.
///
/// The picture is borrowed rather than owned because the picture *is* — what a bubble shows of a
/// film already decoded is that frame's own bytes, and a bubble that copied them would be a second
/// copy of a picture this app is holding at the very moment the pin collapses (see `bubble_art`).
pub(crate) enum BubbleArt<'a> {
    /// A picture already cropped to `side × side` BGRA (an image, or a video frame).
    Picture {
        pixels: &'a [u8],
        width: u32,
        height: u32,
    },
    /// The live level of the sound on screen (0.0..=1.0) and the animation phase in radians.
    Audio { level: f32, phase: f32 },
    /// No art: the face alone, with the mark on it.
    None,
}

/// Paint the round bubble a collapsed pin becomes: a face of the theme's own color, a ring around
/// it, the picture that was pinned inside it where there is one — or, for a sound, the row of bars
/// the level on screen is drawn as — and a soft shadow under it so that it reads as floating over
/// whatever is behind it.
///
/// It is written straight into the buffer rather than through a surface: a bubble is a circle
/// and a picture, and neither of the two needs a device context. What the buffer is, is the
/// premultiplied BGRA of a layered window (see the module documentation).
pub(crate) fn paint_bubble(
    buffer: &mut [u8],
    width: u32,
    height: u32,
    palette: &ChromePalette,
    art: BubbleArt<'_>,
    mark: BubbleMark,
) {
    let width = width as i32;
    let height = height as i32;
    let scale = (width as f32 / 44.0).clamp(0.5, 4.0);

    let radius = (width.min(height) as f32) / 2.0 - 1.0;
    let center = (width as f32 / 2.0, height as f32 / 2.0);
    let ring_thickness = (1.6 * scale).clamp(1.0, 4.0);

    // The ring is the picture's own colour where there is a picture to take one from, and the
    // theme's accent where there is not — which is the whole of what "the bubble belongs to the
    // file" means at a glance. A hairline of the theme around somebody else's photograph is a
    // hairline of this app around it; a hairline of the photograph's own colour is the picture
    // itself coming to its own edge, which is what a window onto a file should look like. A
    // picture that is transparent everywhere has no colour to give, and the accent stands (see
    // `average_color`).
    let ring = match &art {
        BubbleArt::Picture { pixels, .. } => average_color(pixels).unwrap_or(palette.accent),
        _ => palette.accent,
    };

    buffer.fill(0);

    // The shadow: the ring's own coverage, held a couple of pixels lower and spread a little,
    // at a fraction of an opaque black. It is what makes a flat circle read as a thing lying
    // over the desktop rather than a hole in it.
    let shadow_offset = (1.5 * scale).max(1.0);
    let shadow_spread = (1.5 * scale).max(1.0);
    for y in 0..height {
        for x in 0..width {
            let dx = x as f32 + 0.5 - center.0;
            let dy = y as f32 + 0.5 - (center.1 + shadow_offset);
            let distance = (dx * dx + dy * dy).sqrt();
            let coverage =
                ((radius + shadow_spread - distance) / shadow_spread).clamp(0.0, 1.0) * 0.35;
            if coverage > 0.0 {
                put(buffer, width, x, y, [0, 0, 0], coverage);
            }
        }
    }

    for y in 0..height {
        for x in 0..width {
            let dx = x as f32 + 0.5 - center.0;
            let dy = y as f32 + 0.5 - center.1;
            let distance = (dx * dx + dy * dy).sqrt();
            let coverage = (radius - distance + 0.5).clamp(0.0, 1.0);
            if coverage <= 0.0 {
                continue;
            }

            // The picture inside the bubble, clipped to the circle the ring leaves: what a
            // pinned preview looks like from outside the window it collapsed out of.
            let inner = radius - ring_thickness;
            let sample = match &art {
                BubbleArt::Picture {
                    pixels,
                    width: picture_width,
                    height: picture_height,
                } if distance <= inner && *picture_width > 0 && *picture_height > 0 => {
                    let u = ((x as f32 + 0.5 - center.0) / (inner * 2.0) + 0.5).clamp(0.0, 0.999);
                    let v = ((y as f32 + 0.5 - center.1) / (inner * 2.0) + 0.5).clamp(0.0, 0.999);
                    let sx = (u * *picture_width as f32) as u32;
                    let sy = (v * *picture_height as f32) as u32;
                    let index = (sy as usize * *picture_width as usize + sx as usize) * 4;

                    // A pixel the picture does not cover is not a pixel of the picture: what is
                    // behind the bubble is the desktop, so a hole in the frame has to be the face
                    // rather than a hole in the circle, and the face is what the fall-through
                    // below writes.
                    pixels
                        .get(index..index + 4)
                        .filter(|pixel| pixel[3] != 0)
                        .map(|pixel| [pixel[2], pixel[1], pixel[0], pixel[3]])
                }
                _ => None,
            };

            if let Some(sample) = sample {
                put(
                    buffer,
                    width,
                    x,
                    y,
                    [sample[0], sample[1], sample[2]],
                    coverage * (sample[3] as f32 / 255.0),
                );
                continue;
            }

            // The ring, and the face inside it where no picture stands in for one.
            if distance > inner {
                put(buffer, width, x, y, ring, coverage);
            } else {
                put(buffer, width, x, y, palette.background, coverage);
            }
        }
    }

    // The bars the level is drawn as, laid down before the glyph because the glyph goes over
    // them: what the mark of a playing sound says is what the *sound* is doing, and a row of bars
    // drawn through a play triangle is neither a triangle nor a meter (see `draw_histogram`).
    if let BubbleArt::Audio { level, phase } = &art {
        draw_histogram(
            buffer,
            width,
            center,
            radius - ring_thickness,
            ring,
            *level,
            *phase,
        );
    }

    // The mark, and the two cases it is drawn in. A bubble with nothing inside it carries the mark
    // as its whole answer, and a bubble standing for something that plays carries it *over* what
    // it is showing — which is the play/pause of the file on screen, and is the one thing about a
    // bubble that changes while nothing else does.
    if matches!(&art, BubbleArt::None) || mark.is_playback() {
        if matches!(&art, BubbleArt::Picture { .. }) {
            // A glyph over a photograph is a glyph over anything a photograph can be, and half of
            // what a photograph can be is as bright as the glyph. So a faint dark disc goes down
            // first, which does not hide the frame — it is a third of an opaque black, at half
            // the face — and does make the triangle read on a white sky. The audio face needs
            // none of it: it is the theme's own background, and the theme is what the glyph was
            // chosen against.
            fill_disc(
                buffer,
                width,
                center.0,
                center.1,
                (radius - ring_thickness) * 0.5,
                [0, 0, 0],
                0.35,
            );
        }

        paint_mark(
            buffer,
            width,
            center,
            radius - ring_thickness,
            palette.foreground,
            palette.accent,
            mark,
        );
    }
}

/// The mean colour of a picture, sampled rather than walked whole.
///
/// **It is the cheap average and it is meant to be.** A bubble is drawn in the moment a pin
/// collapses, on the thread that owns the frame, and a pass over every pixel of a 4K frame at that
/// moment is a stall the user reads as the bubble taking a breath before it appears. A few
/// thousand samples answer the same question — a ring is one colour wide and not a survey — so the
/// stride is worked out from the picture's own length, and two pictures of a million pixels and of
/// a hundred cost the same to ask.
///
/// A pixel the picture does not cover is not a pixel of it: an alpha of zero is skipped rather
/// than counted as black, so the average of a frame with a hole in it is the average of what is
/// drawn. A picture that is transparent everywhere has no colour to give, and answers nothing —
/// which the caller reads as *no colour from here* and answers with the theme's accent.
pub(super) fn average_color(pixels: &[u8]) -> Option<[u8; 3]> {
    let count = pixels.len() / 4;
    if count == 0 {
        return None;
    }

    let step = (count / 4096).max(1);

    // Summed in the order a colour is written — the buffer is BGRA and a palette is RGB — so that
    // what comes out is a colour the ring can be given directly.
    let mut sum = [0u32; 3];
    let mut seen = 0u32;

    for index in (0..count).step_by(step) {
        let pixel = &pixels[index * 4..index * 4 + 4];
        if pixel[3] == 0 {
            continue;
        }

        sum[0] += pixel[2] as u32;
        sum[1] += pixel[1] as u32;
        sum[2] += pixel[0] as u32;
        seen += 1;
    }

    (seen > 0).then(|| {
        [
            (sum[0] / seen) as u8,
            (sum[1] / seen) as u8,
            (sum[2] / seen) as u8,
        ]
    })
}

/// A symmetrical row of bars across the bubble, driven by one loudness number and a phase.
///
/// The sound on screen is a sound and not a spectrum: what a bubble over one is asked is *is it
/// making a noise, and how much*, which is one number — the system's own output peak, read as a
/// call rather than as a transform (see `audio_meter`). A frequency analysis would be a window of
/// samples per repaint, on the thread that draws the bubble, for a shape that at forty-four pixels
/// across reads the same either way.
///
/// A row of bars rather than one, and shaped rather than equal, because that is what a meter is:
/// one bar says how loud and stops there, where a row says *something is playing* — which is the
/// whole of what a bubble standing for a sound has to say, the card's own clock being the other
/// half of it and off the screen by then. The shapes are one another delayed by a fixed step of
/// the phase, so the row reads as a wave travelling along it rather than as a block of level.
///
/// **Every pixel written asks the circle**, and the asking is per pixel rather than taken for
/// granted from the bars: the row is as wide as the face is, and the face narrows at its own ends,
/// so a bar at either end growing to its full height would have its corners out in the ring —
/// written over the bubble's edge, which is the one thing this must never do (see `paint_bubble`).
fn draw_histogram(
    buffer: &mut [u8],
    width: i32,
    center: (f32, f32),
    inner: f32,
    color: [u8; 3],
    level: f32,
    phase: f32,
) {
    const BARS: usize = 7;

    let level = level.clamp(0.0, 1.0);
    if level <= 0.0 {
        return;
    }

    // The bar and the gap are the same width, which is what makes the row read as bars rather
    // than as a filled band at a glance: the width is a share of the face, so the row is the same
    // fraction of every bubble it is drawn in.
    let bar_width = (inner * 2.0 / (BARS as f32 * 2.0)).max(2.0);
    let gap = bar_width;
    let total = BARS as f32 * bar_width + (BARS - 1) as f32 * gap;
    let mut left = center.0 - total / 2.0;

    for bar in 0..BARS {
        let shape = 0.35 + 0.65 * (0.5 + 0.5 * (phase + bar as f32 * 0.9).sin());
        let half = inner * 0.86 * level * shape / 2.0;
        let top = (center.1 - half).round() as i32;
        let bottom = (center.1 + half).round() as i32;

        // A bar is at least one row tall whatever the level rounds to: a bar that drew nothing
        // would be a gap in the row, and a gap in the row is a bar that is not there.
        for x in (left.round() as i32)..((left + bar_width).round() as i32) {
            for y in top..bottom.max(top + 1) {
                let dx = x as f32 + 0.5 - center.0;
                let dy = y as f32 + 0.5 - center.1;
                if (dx * dx + dy * dy).sqrt() <= inner {
                    put(buffer, width, x, y, color, 1.0);
                }
            }
        }

        left += bar_width + gap;
    }
}

fn paint_mark(
    buffer: &mut [u8],
    width: i32,
    center: (f32, f32),
    inner: f32,
    foreground: [u8; 3],
    accent: [u8; 3],
    mark: BubbleMark,
) {
    let size = (inner * 0.62).max(4.0);
    match mark {
        BubbleMark::Play => {
            // A triangle pointing right, drawn as a scan of half-widths: the widest at the
            // left edge, narrowing to the point.
            let left = center.0 - size * 0.32;
            let right = center.0 + size * 0.52;
            let half = size * 0.5;
            let steps = (right - left).max(1.0) as i32;
            for step in 0..=steps {
                let x = left + step as f32;
                let share = 1.0 - (step as f32 / steps as f32);
                let bar = half * share;
                for y in 0..=(bar.ceil() as i32) {
                    put(
                        buffer,
                        width,
                        x as i32,
                        (center.1 - y as f32) as i32,
                        foreground,
                        1.0,
                    );
                    put(
                        buffer,
                        width,
                        x as i32,
                        (center.1 + y as f32) as i32,
                        foreground,
                        1.0,
                    );
                }
            }
        }
        BubbleMark::Pause => {
            // Two bars: the same mark a player's own pause button is, and the answer to the same
            // question the triangle answers the other way. They are drawn as two boxes rather
            // than as one box with a slot cut out of it, because there is no cutting out here —
            // every shape in this module is written one pixel at a time and the gap between the
            // two is the space the two boxes leave (see `fill_box`).
            let bar = (size * 0.16).max(1.0);
            let height = size * 0.62;
            let gap = size * 0.12;
            let left = center.0 - bar - gap / 2.0;

            for offset in [0.0, bar + gap] {
                fill_box(
                    buffer,
                    width,
                    RECT {
                        left: (left + offset).round() as i32,
                        top: (center.1 - height / 2.0).round() as i32,
                        right: (left + offset + bar).round() as i32,
                        bottom: (center.1 + height / 2.0).round() as i32,
                    },
                    foreground,
                    1.0,
                );
            }
        }
        BubbleMark::Page => {
            let lines = 3;
            let line_height = (size * 0.12).max(1.0);
            let gap = (size * 0.16).max(1.0);
            let total = lines as f32 * line_height + (lines - 1) as f32 * gap;
            let mut top = center.1 - total / 2.0;
            for line in 0..lines {
                let inset = if line == lines - 1 { size * 0.22 } else { 0.0 };
                fill_box(
                    buffer,
                    width,
                    RECT {
                        left: (center.0 - size * 0.42) as i32,
                        top: top.round() as i32,
                        right: (center.0 + size * 0.42 - inset) as i32,
                        bottom: (top + line_height).round() as i32,
                    },
                    foreground,
                    1.0,
                );
                top += line_height + gap;
            }
        }
        BubbleMark::Picture => {
            // A dot of the theme's accent: a picture whose own frame is not in hand, said in
            // one shape rather than guessed at.
            let dot = (size * 0.26).max(2.0);
            for y in 0..=(dot.ceil() as i32 * 2) {
                for x in 0..=(dot.ceil() as i32 * 2) {
                    let dx = x as f32 - dot;
                    let dy = y as f32 - dot;
                    let distance = (dx * dx + dy * dy).sqrt();
                    let coverage = (dot - distance + 0.5).clamp(0.0, 1.0);
                    put(
                        buffer,
                        width,
                        center.0 as i32 - dot as i32 + x,
                        center.1 as i32 - dot as i32 + y,
                        accent,
                        coverage,
                    );
                }
            }
        }
    }
}

/// The mark that stands in for a file a pinned window was shown and could not draw: a cross of
/// two strokes walked a pixel at a time, on a panel of the theme's own page.
///
/// It is a mark rather than a frame of a file because there is no file behind it: nothing was
/// read, nothing was decoded, and there is nothing to read the next time the file is visited —
/// which is the whole of why this is drawn rather than loaded, and why it is never put in the
/// image cache (see `unplayable_media`). It is a cross rather than the empty band a file this
/// app has no preview for leaves, because the two are different answers: this one says the file
/// was tried and failed, where the other says it was never previewable at all.
///
/// The cross is the caption's own close mark (`draw_cross`) walked across the whole panel, at
/// a stroke of its own rather than the glyph's, and every pixel of it is opaque — which is what
/// lets the same premultiplied buffer be a preview's frame: at full coverage the premultiplied
/// and the straight readings of a pixel are the same pixel (see the module documentation).
///
/// The panel is filled first and in full, so what the cross is drawn over is the theme rather
/// than whatever the buffer held: a mark that were transparent between its strokes would show
/// the tray's backdrop through and read as a broken picture rather than as an answer.
pub(crate) fn paint_failure_mark(
    buffer: &mut [u8],
    width: u32,
    height: u32,
    palette: &ChromePalette,
) {
    let side = width.min(height);
    if side == 0 || buffer.len() < width as usize * height as usize * 4 {
        return;
    }

    let (w, h) = (width as i32, height as i32);
    fill_box(
        buffer,
        w,
        RECT {
            left: 0,
            top: 0,
            right: w,
            bottom: h,
        },
        palette.background,
        1.0,
    );

    // Too small for a cross to be a cross, and the panel is the answer anyway: a mark the size of
    // a few pixels is read as the file being small, which is a different thing entirely.
    if side < 16 {
        return;
    }

    let side = side as f32;
    let span = (side * 0.62).round() as i32;
    let stroke = (side * 0.085).round().max(2.0);
    draw_cross(
        buffer,
        w,
        w / 2,
        h / 2 - (stroke.round() as i32) / 2,
        span,
        stroke,
        palette.foreground,
    );
}

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

use crate::text::text_paint::{self, DibSurface};
use crate::CONFIG;
use windows::Win32::Foundation::{RECT, SIZE};
use windows::Win32::Graphics::Gdi::{GetTextExtentPoint32W, SelectObject};

/// The face a caption's title is set in: the one Windows writes a window's title in, so a
/// pinned preview's caption reads like the caption of anything else on the desktop.
const CAPTION_FACE: &str = "Segoe UI";

/// How wide a caption button is, in the units a display's scale multiplies: the width
/// Windows 11 gives one, so the three land where a hand expects them.
const BUTTON_PIXELS: f32 = 46.0;

/// The room a title is kept clear of the left edge by.
const TITLE_PADDING_PIXELS: f32 = 12.0;

/// How thick a glyph's strokes are drawn.
const GLYPH_STROKE_PIXELS: f32 = 1.4;

/// How wide the glyph inside a caption button is.
const GLYPH_PIXELS: f32 = 10.0;

/// The colors a caption and a bubble are painted in.
///
/// They come from the theme the text previews are painted with — One Dark Pro, Atom One
/// Light, or a `.tmTheme` of the user's own — rather than from a palette of this app's, so
/// that the chrome of a pinned preview belongs to the same app as the page inside it. What
/// the theme is *read* for is three colors: its page, its text, and one it uses for
/// something else, which is drawn as the accent.
pub(crate) struct ChromePalette {
    /// The theme's page color: the caption's bar and the bubble's face.
    pub(crate) background: [u8; 3],
    /// The theme's text color: the title and the glyphs.
    pub(crate) foreground: [u8; 3],
    /// A color the theme spends on something else — a keyword, which every theme colors —
    /// drawn as the accent: the bubble's ring, and the played part of a transport's bar.
    pub(crate) accent: [u8; 3],
    /// Whether the theme is a dark one, which decides which way a hairline is shaded.
    pub(crate) dark: bool,
}

impl ChromePalette {
    /// The palette the configuration currently names, or `None` when no theme can be read
    /// at all — which the bundled themes make impossible in practice, and which is answered
    /// with no chrome rather than with guessed colors (see `text_theme::loaded`).
    pub(crate) fn current() -> Option<Self> {
        let kind = CONFIG.lock().ok()?.theme;
        let theme = crate::text::text_theme::loaded(kind)?;

        let background = text_paint::rgb(theme.background());
        let foreground = text_paint::rgb(theme.foreground());
        let accent = text_paint::rgb(
            theme
                .style_for_scopes(&["keyword.control", "keyword"])
                .foreground,
        );

        Some(Self {
            background,
            foreground,
            accent,
            dark: text_paint::luminance(background) <= 0.5,
        })
    }

    /// A color blended toward the opposite end of the theme, for the surface a pointer is
    /// over: a caption button lights up, and the rest of the caption stays where it is.
    pub(crate) fn hover(&self, amount: f32) -> [u8; 3] {
        let toward = if self.dark {
            [255, 255, 255]
        } else {
            [0, 0, 0]
        };
        text_paint::blend(self.background, toward, amount)
    }
}

/// The parts of a transport bar a pointer can be on.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum TransportPart {
    /// The button that pauses and resumes.
    Play,
    /// The bar itself, which is dragged to seek.
    Seek,
}

/// Where the parts of a transport bar are, in the bar's own coordinates: the button, the bar
/// itself, and the two labels beside them.
pub(crate) struct TransportLayout {
    pub(crate) play: RECT,
    pub(crate) bar: RECT,
    pub(crate) elapsed: RECT,
    pub(crate) total: RECT,
}

/// The room each part of a transport bar is given at a display's scale.
pub(crate) fn transport_layout(width: i32, height: i32, dpi: u32) -> TransportLayout {
    let scale = dpi as f32 / 96.0;
    let padding = text_paint::scaled(10, scale);
    let play_side = (height - text_paint::scaled(8, scale)).clamp(8, height.max(8));
    let label = text_paint::scaled(52, scale);
    let gap = (padding / 2).max(2);

    let play = RECT {
        left: padding,
        top: (height - play_side) / 2,
        right: padding + play_side,
        bottom: (height + play_side) / 2,
    };
    let elapsed = RECT {
        left: play.right + gap,
        top: 0,
        right: play.right + gap + label,
        bottom: height,
    };
    let total = RECT {
        left: (width - padding - label).max(elapsed.right),
        top: 0,
        right: (width - padding).max(elapsed.right),
        bottom: height,
    };
    let bar_height = text_paint::scaled(4, scale).clamp(2, height.max(2));
    let bar = RECT {
        left: elapsed.right + gap,
        top: (height - bar_height) / 2,
        right: (total.left - gap).max(elapsed.right + gap),
        bottom: (height + bar_height) / 2,
    };

    TransportLayout {
        play,
        bar,
        elapsed,
        total,
    }
}

/// Which part of a transport bar a point is on, in the bar's own coordinates.
pub(crate) fn transport_part_at(
    x: i32,
    y: i32,
    width: i32,
    height: i32,
    dpi: u32,
) -> Option<TransportPart> {
    let layout = transport_layout(width, height, dpi);
    let inside = |rect: RECT| x >= rect.left && x < rect.right && y >= rect.top && y < rect.bottom;

    if inside(layout.play) {
        return Some(TransportPart::Play);
    }
    // The bar is a thin line, so what is dragged is a band around it rather than the line: a
    // target four pixels high is one a hand cannot hold on to.
    let reach = text_paint::scaled(10, dpi as f32 / 96.0).max(6);
    if y >= layout.bar.top - reach
        && y < layout.bar.bottom + reach
        && x >= layout.bar.left
        && x < layout.bar.right
    {
        return Some(TransportPart::Seek);
    }

    None
}

/// Where along a transport bar a point is, as a share of the file: what a press on the bar asks
/// the player to be taken to.
pub(crate) fn transport_share_at(x: i32, width: i32, dpi: u32) -> f64 {
    let layout = transport_layout(width, 0, dpi);
    let span = (layout.bar.right - layout.bar.left).max(1) as f64;

    ((x - layout.bar.left) as f64 / span).clamp(0.0, 1.0)
}

/// What a transport bar is drawn from.
pub(crate) struct TransportState {
    /// Whether a player is running, which is what the button's glyph says.
    pub(crate) playing: bool,
    /// Where the playhead is, and how long the file is: nothing for either is a bar with no
    /// length to draw a playhead against.
    pub(crate) position: Option<f64>,
    pub(crate) duration: Option<f64>,
    pub(crate) hovered: Option<TransportPart>,
    pub(crate) pressed: Option<TransportPart>,
}

/// Paint a transport bar into a surface the size of the strip: the play button, the bar with the
/// part of the file that has been played filled in, and the two clocks beside them.
pub(crate) fn paint_transport(
    surface: &DibSurface,
    palette: &ChromePalette,
    state: &TransportState,
    dpi: u32,
) {
    let width = surface.width as i32;
    let height = surface.height as i32;
    let scale = dpi as f32 / 96.0;
    let layout = transport_layout(width, height, dpi);

    text_paint::fill_rect(
        surface,
        RECT {
            left: 0,
            top: 0,
            right: width,
            bottom: height,
        },
        palette.background,
    );

    // A hairline along the top, the way the caption carries one along its bottom: the two strips
    // are the frame the media sits in.
    text_paint::fill_rect(
        surface,
        RECT {
            left: 0,
            top: 0,
            right: width,
            bottom: 1,
        },
        palette.hover(0.18),
    );

    paint_play_button(surface, palette, state, layout.play, scale);

    // The bar: a track, what has been played filled in, and a thumb at the playhead.
    let filled = match (state.position, state.duration) {
        (Some(position), Some(duration)) if duration > 0.0 => {
            position.clamp(0.0, duration) / duration
        }
        _ => 0.0,
    };
    let thumb =
        layout.bar.left + (((layout.bar.right - layout.bar.left) as f64) * filled).round() as i32;

    {
        let buffer = unsafe {
            std::slice::from_raw_parts_mut(
                surface.bits(),
                surface.width as usize * surface.height as usize * 4,
            )
        };

        fill_box(buffer, width, layout.bar, palette.hover(0.28), 1.0);
        fill_box(
            buffer,
            width,
            RECT {
                left: layout.bar.left,
                top: layout.bar.top,
                right: thumb.max(layout.bar.left),
                bottom: layout.bar.bottom,
            },
            palette.accent,
            1.0,
        );

        // The thumb is drawn only where there is a length to place it against: a bar with no
        // duration is a file whose container says nothing, and a thumb on it would be a position
        // that means nothing.
        if state.duration.is_some() {
            let radius = text_paint::scaled(5, scale).clamp(3, 12);
            let center_y = (layout.bar.top + layout.bar.bottom) / 2;
            for y in -radius..=radius {
                for x in -radius..=radius {
                    let distance = ((x * x + y * y) as f32).sqrt();
                    let coverage = (radius as f32 - distance + 0.5).clamp(0.0, 1.0);
                    put(
                        buffer,
                        width,
                        thumb + x,
                        center_y + y,
                        palette.foreground,
                        coverage,
                    );
                }
            }
        }
    }

    paint_time(
        surface,
        palette,
        state.position,
        layout.elapsed,
        scale,
        false,
    );
    paint_time(surface, palette, state.duration, layout.total, scale, true);

    // GDI leaves the alpha byte of everything it draws at zero (see the module documentation).
    let buffer = unsafe {
        std::slice::from_raw_parts_mut(
            surface.bits(),
            surface.width as usize * surface.height as usize * 4,
        )
    };
    for pixel in buffer.as_chunks_mut::<4>().0 {
        pixel[3] = 255;
    }
}

fn paint_play_button(
    surface: &DibSurface,
    palette: &ChromePalette,
    state: &TransportState,
    rect: RECT,
    scale: f32,
) {
    if state.hovered == Some(TransportPart::Play) || state.pressed == Some(TransportPart::Play) {
        let wash = if state.pressed == Some(TransportPart::Play) {
            palette.hover(0.18)
        } else {
            palette.hover(0.10)
        };
        text_paint::fill_rect(surface, rect, wash);
    }

    let ink = palette.foreground;
    let width = surface.width as i32;
    let center_x = (rect.left + rect.right) / 2;
    let center_y = (rect.top + rect.bottom) / 2;
    let size = text_paint::scaled(9, scale).clamp(6, rect.bottom - rect.top);

    let buffer = unsafe {
        std::slice::from_raw_parts_mut(
            surface.bits(),
            surface.width as usize * surface.height as usize * 4,
        )
    };

    if state.playing {
        // Pause: two bars.
        let bar = (size / 3).max(2);
        let gap = (size / 5).max(1);
        fill_box(
            buffer,
            width,
            RECT {
                left: center_x - gap - bar,
                top: center_y - size / 2,
                right: center_x - gap,
                bottom: center_y + size / 2,
            },
            ink,
            1.0,
        );
        fill_box(
            buffer,
            width,
            RECT {
                left: center_x + gap,
                top: center_y - size / 2,
                right: center_x + gap + bar,
                bottom: center_y + size / 2,
            },
            ink,
            1.0,
        );
        return;
    }

    // Play: a triangle, drawn as a scan of half-widths, widest at the left edge.
    let left = center_x - size / 3;
    let half = size / 2;
    for step in 0..=size {
        let x = left + step;
        let share = 1.0 - (step as f32 / size as f32);
        for y in 0..=(half as f32 * share) as i32 {
            put(buffer, width, x, center_y - y, ink, 1.0);
            put(buffer, width, x, center_y + y, ink, 1.0);
        }
    }
}

/// One of the two clocks: where the file is, or how long it is. A file with no length is drawn as
/// a pair of dashes rather than as a zero, which is the truth about it.
fn paint_time(
    surface: &DibSurface,
    palette: &ChromePalette,
    seconds: Option<f64>,
    rect: RECT,
    scale: f32,
    align_right: bool,
) {
    let text = clock_text(seconds);
    let style = text_paint::TextStyle {
        foreground: palette.foreground,
        background: None,
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        level: text_paint::BODY_LEVEL,
        face: CAPTION_FACE,
    };

    let cell = text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[text_paint::BODY_LEVEL as usize],
        scale,
    );
    let top = ((surface.height as i32 - cell) / 2).max(0);
    let left = if align_right {
        let measured = measure_text(surface, &style, &text, scale);
        (rect.right - measured - 1).max(rect.left)
    } else {
        rect.left + 1
    };

    let mut painter = text_paint::RunPainter::new(surface, scale);
    painter.draw(
        &text,
        left,
        RECT {
            left: rect.left,
            top,
            right: rect.right,
            bottom: (top + cell).min(surface.height as i32),
        },
        &style,
        palette.foreground,
        palette.background,
    );
}

/// How a number of seconds is written on a transport bar: minutes and seconds, hours where there
/// are any, and dashes where there is nothing to write.
pub(crate) fn clock_text(seconds: Option<f64>) -> String {
    let Some(seconds) = seconds.filter(|value| value.is_finite() && *value >= 0.0) else {
        return "--:--".to_string();
    };

    let total = seconds.round() as u64;
    let (hours, minutes, seconds) = (total / 3600, (total % 3600) / 60, total % 60);

    if hours > 0 {
        format!("{hours}:{minutes:02}:{seconds:02}")
    } else {
        format!("{minutes}:{seconds:02}")
    }
}

/// The width of a run of text at a style, which the two clocks need to be placed against the
/// edges they sit on.
fn measure_text(
    surface: &DibSurface,
    style: &text_paint::TextStyle,
    text: &str,
    scale: f32,
) -> i32 {
    unsafe {
        let font = text_paint::create_font(
            text_paint::scaled(text_paint::LEVEL_FONT_PIXELS[style.level as usize], scale),
            style,
        );
        if font.0.is_null() {
            return 0;
        }

        let previous = SelectObject(surface.dc, font);
        let wide: Vec<u16> = text.encode_utf16().collect();
        let mut extent = SIZE::default();
        let measured = GetTextExtentPoint32W(surface.dc, &wide, &mut extent).as_bool();
        let _ = SelectObject(surface.dc, previous);
        let _ = windows::Win32::Graphics::Gdi::DeleteObject(font);

        if measured {
            extent.cx
        } else {
            0
        }
    }
}

/// The mark a media band carries when the player that would be filling it has been stopped: what
/// a paused picture looks like where the picture would be.
///
/// The mark is placed in the middle of the band the media occupies, which is a row range of the
/// window's own surface rather than a surface of its own (see `render_pinned_preview_at`).
pub(crate) fn paint_paused_mark(
    buffer: &mut [u8],
    surface_width: u32,
    band_top: u32,
    band_height: u32,
    palette: &ChromePalette,
) {
    let width = surface_width as i32;
    let size = (band_height as f32 * 0.14).clamp(18.0, 128.0);
    let center_x = width / 2;
    let center_y = band_top as i32 + band_height as i32 / 2;

    let left = center_x - (size / 3.0) as i32;
    for step in 0..=(size as i32) {
        let x = left + step;
        let share = 1.0 - (step as f32 / size);
        for y in 0..=((size / 2.0) * share) as i32 {
            put(buffer, width, x, center_y - y, palette.foreground, 0.75);
            put(buffer, width, x, center_y + y, palette.foreground, 0.75);
        }
    }
}

/// The three buttons a caption carries, in the order Windows has them: minimize, maximize or
/// restore, close.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum CaptionButton {
    Minimize,
    Maximize,
    Close,
}

/// What a caption is drawn from: the name of what is pinned, whether it is maximized, whether it
/// offers a maximize at all, and which button the pointer is on, if any.
pub(crate) struct Caption<'a> {
    pub(crate) title: &'a str,
    pub(crate) maximized: bool,
    /// Whether this pin has a maximum to go to. A sound's card does not: the card is its own
    /// size, so there is no larger box to give it and the button would be one that does nothing
    /// (see `pin_frame`).
    pub(crate) maximizable: bool,
    pub(crate) hovered: Option<CaptionButton>,
    pub(crate) pressed: Option<CaptionButton>,
}

/// Where a caption's buttons are, in the caption's own coordinates: each one's box, left to
/// right. They are the caption's own height and sit against its right edge, which is where a
/// hand goes for a window it wants to close — and the same place Windows puts them.
///
/// A caption without a maximize carries two buttons rather than three, and the two it carries
/// keep their own places at that edge: closing a window is a gesture made by position, and a
/// close button that moved because the window it is on has nothing to maximize would be a
/// window whose close button is where a maximize is on every other one.
pub(crate) fn button_boxes(
    width: i32,
    height: i32,
    dpi: u32,
    maximizable: bool,
) -> Vec<CaptionButtonBox> {
    let kinds: &[CaptionButton] = if maximizable {
        &[
            CaptionButton::Minimize,
            CaptionButton::Maximize,
            CaptionButton::Close,
        ]
    } else {
        &[CaptionButton::Minimize, CaptionButton::Close]
    };

    let button = text_paint::scaled(BUTTON_PIXELS as i32, dpi as f32 / 96.0)
        .max(1)
        .min((width / kinds.len().max(1) as i32).max(1));

    kinds
        .iter()
        .enumerate()
        .map(|(index, kind)| {
            let from_the_right = kinds.len() as i32 - 1 - index as i32;
            let right = width - from_the_right * button;
            CaptionButtonBox {
                kind: *kind,
                rect: RECT {
                    left: right - button,
                    top: 0,
                    right,
                    bottom: height,
                },
            }
        })
        .collect()
}

/// One caption button and the box it occupies.
#[derive(Clone, Copy)]
pub(crate) struct CaptionButtonBox {
    pub(crate) kind: CaptionButton,
    pub(crate) rect: RECT,
}

/// Which button a point in the caption's own coordinates is on, if any.
pub(crate) fn button_at(
    x: i32,
    y: i32,
    width: i32,
    height: i32,
    dpi: u32,
    maximizable: bool,
) -> Option<CaptionButton> {
    button_boxes(width, height, dpi, maximizable)
        .into_iter()
        .find(|button| {
            x >= button.rect.left
                && x < button.rect.right
                && y >= button.rect.top
                && y < button.rect.bottom
        })
        .map(|button| button.kind)
}

/// Paint a caption into a surface the size of the strip: the bar, the three buttons with
/// whatever state the pointer has put them in, and the name of what is pinned.
pub(crate) fn paint_caption(
    surface: &DibSurface,
    palette: &ChromePalette,
    caption: &Caption,
    dpi: u32,
) {
    let width = surface.width as i32;
    let height = surface.height as i32;
    let scale = dpi as f32 / 96.0;

    text_paint::fill_rect(
        surface,
        RECT {
            left: 0,
            top: 0,
            right: width,
            bottom: height,
        },
        palette.background,
    );

    // A hairline along the bottom, so the caption reads as a strip over a picture of any
    // color rather than dissolving into one that happens to match it.
    let hairline = palette.hover(0.18);
    text_paint::fill_rect(
        surface,
        RECT {
            left: 0,
            top: height - 1,
            right: width,
            bottom: height,
        },
        hairline,
    );

    let buttons = button_boxes(width, height, dpi, caption.maximizable);
    for button in &buttons {
        paint_button(surface, palette, caption, *button, scale);
    }

    let title_right = buttons
        .iter()
        .map(|button| button.rect.left)
        .min()
        .unwrap_or(width);
    paint_title(surface, palette, caption.title, title_right, scale);
}

fn paint_button(
    surface: &DibSurface,
    palette: &ChromePalette,
    caption: &Caption,
    button: CaptionButtonBox,
    scale: f32,
) {
    let pressed = caption.pressed == Some(button.kind);
    let hovered = caption.hovered == Some(button.kind);

    if pressed || hovered {
        // The close button is washed with the red Windows washes it with, whichever theme is
        // in use: it is the one button whose meaning does not depend on the palette. The
        // other two are washed with the theme's own background moved a step toward its text.
        let wash = if button.kind == CaptionButton::Close {
            if pressed {
                [178, 32, 32]
            } else {
                [196, 43, 28]
            }
        } else if pressed {
            palette.hover(0.18)
        } else {
            palette.hover(0.10)
        };

        text_paint::fill_rect(surface, button.rect, wash);
    }

    let ink = if button.kind == CaptionButton::Close && (pressed || hovered) {
        [255, 255, 255]
    } else {
        palette.foreground
    };

    paint_glyph(surface, button, caption.maximized, ink, scale);
}

fn paint_glyph(
    surface: &DibSurface,
    button: CaptionButtonBox,
    maximized: bool,
    ink: [u8; 3],
    scale: f32,
) {
    let buffer = unsafe {
        std::slice::from_raw_parts_mut(
            surface.bits(),
            surface.width as usize * surface.height as usize * 4,
        )
    };
    let width = surface.width as i32;

    let center_x = (button.rect.left + button.rect.right) / 2;
    let center_y = (button.rect.top + button.rect.bottom) / 2;
    let glyph = text_paint::scaled(GLYPH_PIXELS as i32, scale).max(6);
    let stroke = (GLYPH_STROKE_PIXELS * scale).max(1.0);
    let half = glyph / 2;

    match button.kind {
        CaptionButton::Minimize => {
            let top = center_y;
            fill_box(
                buffer,
                width,
                RECT {
                    left: center_x - half,
                    top,
                    right: center_x + half + 1,
                    bottom: top + stroke.round().max(1.0) as i32,
                },
                ink,
                1.0,
            );
        }
        // The restore glyph: the window behind, then the window in front of it, the way
        // Windows draws the pair. A button that means "restore" is the one on a window that
        // is maximized.
        CaptionButton::Maximize if maximized => {
            let inset = text_paint::scaled(3, scale).max(1);
            stroke_box(
                buffer,
                width,
                RECT {
                    left: center_x - half + inset,
                    top: center_y - half,
                    right: center_x + half + 1,
                    bottom: center_y + half + 1 - inset,
                },
                stroke,
                ink,
                1.0,
            );
            stroke_box(
                buffer,
                width,
                RECT {
                    left: center_x - half,
                    top: center_y - half + inset,
                    right: center_x + half + 1 - inset,
                    bottom: center_y + half + 1,
                },
                stroke,
                ink,
                1.0,
            );
        }
        CaptionButton::Maximize => {
            stroke_box(
                buffer,
                width,
                RECT {
                    left: center_x - half,
                    top: center_y - half,
                    right: center_x + half + 1,
                    bottom: center_y + half + 1,
                },
                stroke,
                ink,
                1.0,
            );
        }
        CaptionButton::Close => draw_cross(buffer, width, center_x, center_y, glyph, stroke, ink),
    }
}

fn paint_title(
    surface: &DibSurface,
    palette: &ChromePalette,
    title: &str,
    available_right: i32,
    scale: f32,
) {
    let padding = text_paint::scaled(TITLE_PADDING_PIXELS as i32, scale);
    let left = padding;
    let right = available_right - padding;
    if right <= left || title.is_empty() {
        return;
    }

    let style = text_paint::TextStyle {
        foreground: palette.foreground,
        background: None,
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        level: text_paint::BODY_LEVEL,
        face: CAPTION_FACE,
    };

    let fitted = fit_title(surface, &style, title, right - left, scale);
    if fitted.is_empty() {
        return;
    }

    // ExtTextOut draws from the top of a text cell, so the title is centered by giving the
    // run a box the height of the cell rather than the whole caption.
    let cell = text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[text_paint::BODY_LEVEL as usize],
        scale,
    );
    let top = ((surface.height as i32 - cell) / 2).max(0);
    let rect = RECT {
        left,
        top,
        right,
        bottom: (top + cell).min(surface.height as i32),
    };

    let mut painter = text_paint::RunPainter::new(surface, scale);
    painter.draw(
        &fitted,
        left + 1,
        rect,
        &style,
        palette.foreground,
        palette.background,
    );

    // GDI leaves the alpha byte of everything it draws at zero, and a caption is opaque
    // wherever it is painted: the strip's own coverage is handed back here (see the module
    // documentation).
    let buffer = unsafe {
        std::slice::from_raw_parts_mut(
            surface.bits(),
            surface.width as usize * surface.height as usize * 4,
        )
    };
    for pixel in buffer.as_chunks_mut::<4>().0 {
        pixel[3] = 255;
    }
}

/// The longest beginning of `title` that fits `available` pixels, with an ellipsis after it
/// where any of the name had to be left out. Measured rather than guessed: a caption is the
/// file's own name, and a name cut by a character count is either a name with room to spare
/// or one whose end was cut off twice.
fn fit_title(
    surface: &DibSurface,
    style: &text_paint::TextStyle,
    title: &str,
    available: i32,
    scale: f32,
) -> String {
    if available <= 0 {
        return String::new();
    }

    unsafe {
        let font = text_paint::create_font(
            text_paint::scaled(text_paint::LEVEL_FONT_PIXELS[style.level as usize], scale),
            style,
        );
        if font.0.is_null() {
            return title.to_string();
        }

        let previous = SelectObject(surface.dc, font);
        let measured = |text: &str| -> i32 {
            let wide: Vec<u16> = text.encode_utf16().collect();
            if wide.is_empty() {
                return 0;
            }
            let mut extent = SIZE::default();
            if GetTextExtentPoint32W(surface.dc, &wide, &mut extent).as_bool() {
                extent.cx
            } else {
                0
            }
        };

        let full = measured(title);
        let fitted = if full <= available {
            title.to_string()
        } else {
            let ellipsis = measured("…");
            let budget = (available - ellipsis).max(0);
            let mut best = 0usize;
            for (index, _) in title.char_indices().skip(1) {
                if measured(&title[..index]) > budget {
                    break;
                }
                best = index;
            }

            if best == 0 {
                String::new()
            } else {
                format!("{}…", &title[..best])
            }
        };

        let _ = SelectObject(surface.dc, previous);
        let _ = windows::Win32::Graphics::Gdi::DeleteObject(font);

        fitted
    }
}

/// The mark a bubble carries when the pin it stands for has no picture to show inside it.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum BubbleMark {
    /// Something that plays: a video or a sound.
    Play,
    /// A page: text, an archive's listing, a document, a book.
    Page,
    /// A picture whose frame is not in hand.
    Picture,
}

/// Paint the round bubble a collapsed pin becomes: a face of the theme's own color, a ring of
/// its accent around it, the picture that was pinned inside it where there is one, and a soft
/// shadow under it so that it reads as floating over whatever is behind it.
///
/// It is written straight into the buffer rather than through a surface: a bubble is a circle
/// and a picture, and neither of the two needs a device context. What the buffer is, is the
/// premultiplied BGRA of a layered window (see the module documentation).
pub(crate) fn paint_bubble(
    buffer: &mut [u8],
    width: u32,
    height: u32,
    palette: &ChromePalette,
    thumbnail: Option<(&[u8], u32, u32)>,
    mark: BubbleMark,
) {
    let width = width as i32;
    let height = height as i32;
    let scale = (width as f32 / 44.0).clamp(0.5, 4.0);

    let radius = (width.min(height) as f32) / 2.0 - 1.0;
    let center = (width as f32 / 2.0, height as f32 / 2.0);
    let ring = (1.6 * scale).clamp(1.0, 4.0);

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
            let inner = radius - ring;
            match thumbnail {
                Some((pixels, thumb_width, thumb_height))
                    if distance <= inner && thumb_width > 0 && thumb_height > 0 =>
                {
                    let u = ((x as f32 + 0.5 - center.0) / (inner * 2.0) + 0.5).clamp(0.0, 0.999);
                    let v = ((y as f32 + 0.5 - center.1) / (inner * 2.0) + 0.5).clamp(0.0, 0.999);
                    let sx = (u * thumb_width as f32) as u32;
                    let sy = (v * thumb_height as f32) as u32;
                    let index = (sy as usize * thumb_width as usize + sx as usize) * 4;
                    if index + 3 < pixels.len() {
                        let sample = &pixels[index..index + 4];
                        let color = [sample[2], sample[1], sample[0]];
                        put(
                            buffer,
                            width,
                            x,
                            y,
                            color,
                            coverage * (sample[3] as f32 / 255.0),
                        );
                        continue;
                    }
                }
                _ => {}
            }

            // The ring, and the face inside it where no picture stands in for one.
            if distance > inner {
                put(buffer, width, x, y, palette.accent, coverage);
            } else {
                put(buffer, width, x, y, palette.background, coverage);
            }
        }
    }

    if thumbnail.is_none() {
        paint_mark(
            buffer,
            width,
            center,
            radius - ring,
            palette.foreground,
            palette.accent,
            mark,
        );
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

/// One pixel written premultiplied (see the module documentation).
fn put(buffer: &mut [u8], width: i32, x: i32, y: i32, color: [u8; 3], coverage: f32) {
    if x < 0 || y < 0 || x >= width {
        return;
    }

    let index = (y as usize * width as usize + x as usize) * 4;
    if index + 3 >= buffer.len() {
        return;
    }

    let alpha = (coverage.clamp(0.0, 1.0) * 255.0).round() as u32;
    if alpha == 0 {
        return;
    }

    let existing = [
        buffer[index],
        buffer[index + 1],
        buffer[index + 2],
        buffer[index + 3],
    ];
    let existing_alpha = existing[3] as u32;
    let inverse = 255 - alpha;

    // What is already there is premultiplied too, so it composites the way any two
    // premultiplied colors do.
    buffer[index] = (((color[2] as u32 * alpha) + existing[0] as u32 * inverse + 127) / 255) as u8;
    buffer[index + 1] =
        (((color[1] as u32 * alpha) + existing[1] as u32 * inverse + 127) / 255) as u8;
    buffer[index + 2] =
        (((color[0] as u32 * alpha) + existing[2] as u32 * inverse + 127) / 255) as u8;
    buffer[index + 3] = (alpha + (existing_alpha * inverse + 127) / 255).min(255) as u8;
}

fn fill_box(buffer: &mut [u8], width: i32, rect: RECT, color: [u8; 3], coverage: f32) {
    for y in rect.top..rect.bottom {
        for x in rect.left..rect.right {
            put(buffer, width, x, y, color, coverage);
        }
    }
}

/// The outline of a box, drawn inside its own edges so that a glyph is the size it was asked
/// for rather than a stroke wider on each side.
fn stroke_box(
    buffer: &mut [u8],
    width: i32,
    rect: RECT,
    thickness: f32,
    color: [u8; 3],
    coverage: f32,
) {
    let thickness = thickness.round().max(1.0) as i32;
    let (left, top, right, bottom) = (rect.left, rect.top, rect.right, rect.bottom);

    fill_box(
        buffer,
        width,
        RECT {
            left,
            top,
            right,
            bottom: (top + thickness).min(bottom),
        },
        color,
        coverage,
    );
    fill_box(
        buffer,
        width,
        RECT {
            left,
            top: (bottom - thickness).max(top),
            right,
            bottom,
        },
        color,
        coverage,
    );
    fill_box(
        buffer,
        width,
        RECT {
            left,
            top,
            right: (left + thickness).min(right),
            bottom,
        },
        color,
        coverage,
    );
    fill_box(
        buffer,
        width,
        RECT {
            left: (right - thickness).max(left),
            top,
            right,
            bottom,
        },
        color,
        coverage,
    );
}

/// The two diagonals of a cross, walked a column at a time so that a stroke is one width
/// whoever draws it and whatever the display's scale is.
fn draw_cross(
    buffer: &mut [u8],
    width: i32,
    center_x: i32,
    center_y: i32,
    span: i32,
    thickness: f32,
    color: [u8; 3],
) {
    let thickness = thickness.round().max(1.0) as i32;
    let half = span / 2;

    for step in 0..=span.max(1) {
        let x = center_x - half + step;
        let offset = step - half;
        for depth in 0..thickness {
            put(buffer, width, x, center_y + offset + depth, color, 1.0);
            put(buffer, width, x, center_y - offset + depth, color, 1.0);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_clock_is_written_the_way_a_player_writes_one() {
        assert_eq!(clock_text(Some(7.4)), "0:07");
        assert_eq!(clock_text(Some(62.0)), "1:02");
        assert_eq!(clock_text(Some(3723.0)), "1:02:03");

        // A file whose container says nothing, and one whose length is nonsense, are the same
        // answer: a bar with no length to draw a playhead against.
        assert_eq!(clock_text(None), "--:--");
        assert_eq!(clock_text(Some(f64::NAN)), "--:--");
        assert_eq!(clock_text(Some(-1.0)), "--:--");
    }

    #[test]
    fn the_buttons_sit_against_the_right_edge_in_the_order_windows_has_them() {
        let buttons = button_boxes(600, 30, 96, true);

        assert_eq!(buttons[0].kind, CaptionButton::Minimize);
        assert_eq!(buttons[1].kind, CaptionButton::Maximize);
        assert_eq!(buttons[2].kind, CaptionButton::Close);
        assert!(buttons[0].rect.left < buttons[1].rect.left);
        assert!(buttons[1].rect.left < buttons[2].rect.left);
        assert_eq!(buttons[2].rect.right, 600);

        // And a point is on the button it looks like it is on, or on none of them.
        let close = buttons[2].rect;
        assert_eq!(
            button_at(close.left + 1, 5, 600, 30, 96, true),
            Some(CaptionButton::Close)
        );
        assert_eq!(button_at(1, 5, 600, 30, 96, true), None);
        assert_eq!(button_at(close.left + 1, 40, 600, 30, 96, true), None);
    }

    /// A pin with nothing to maximize — a sound's card — carries the two buttons that mean
    /// something on it, packed against the right edge the way Windows packs a window's: the one
    /// that closes it keeps the place it has on every other caption, and the one beside it is
    /// the minimize that was there before.
    #[test]
    fn a_caption_without_a_maximize_keeps_the_close_button_where_it_was() {
        let three = button_boxes(600, 30, 96, true);
        let two = button_boxes(600, 30, 96, false);

        assert_eq!(two.len(), 2);
        assert_eq!(two[0].kind, CaptionButton::Minimize);
        assert_eq!(two[1].kind, CaptionButton::Close);

        let close = three
            .iter()
            .find(|button| button.kind == CaptionButton::Close)
            .expect("a close button")
            .rect;
        let minimize = three
            .iter()
            .find(|button| button.kind == CaptionButton::Minimize)
            .expect("a minimize button")
            .rect;

        assert_eq!(two[1].rect.left, close.left);
        assert_eq!(two[1].rect.right, close.right);
        assert_eq!(two[0].rect.right, two[1].rect.left);
        assert_eq!(two[0].rect.right - two[0].rect.left, minimize.right - minimize.left);

        // And the space the button used to take is a button's, not a hole: it is the minimize
        // that has moved along into it.
        let maximize = three
            .iter()
            .find(|button| button.kind == CaptionButton::Maximize)
            .expect("a maximize button")
            .rect;
        assert_eq!(
            button_at(maximize.left + 1, 5, 600, 30, 96, false),
            Some(CaptionButton::Minimize)
        );
        assert_eq!(
            button_at(maximize.left + 1, 5, 600, 30, 96, true),
            Some(CaptionButton::Maximize)
        );
    }

    #[test]
    fn a_transport_bar_is_taken_hold_of_where_it_was_pressed() {
        let layout = transport_layout(800, 30, 96);
        let middle = (layout.bar.left + layout.bar.right) / 2;
        let share = transport_share_at(middle, 800, 96);
        assert!(
            (share - 0.5).abs() < 0.05,
            "the middle of the bar is half of it"
        );

        // A press past either end is the end it is past, which is what keeps a drag from asking
        // for a second of a file that is not there.
        assert_eq!(transport_share_at(0, 800, 96), 0.0);
        assert_eq!(transport_share_at(800, 800, 96), 1.0);
    }

    #[test]
    fn the_parts_of_a_transport_bar_are_hit_where_they_are_drawn() {
        let layout = transport_layout(800, 30, 96);

        let play = (layout.play.left + layout.play.right) / 2;
        assert_eq!(
            transport_part_at(play, 15, 800, 30, 96),
            Some(TransportPart::Play)
        );

        let bar = (layout.bar.left + layout.bar.right) / 2;
        assert_eq!(
            transport_part_at(bar, 15, 800, 30, 96),
            Some(TransportPart::Seek)
        );

        // Nothing is hit where nothing is drawn.
        assert_eq!(transport_part_at(2, 2, 800, 30, 96), None);
    }
}

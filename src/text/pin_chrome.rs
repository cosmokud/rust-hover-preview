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
use std::cell::RefCell;
use windows::Win32::Foundation::{RECT, SIZE};
use windows::Win32::Graphics::Gdi::{GetTextExtentPoint32W, SelectObject};

/// The face a caption's title is set in: the one Windows writes a window's title in, so a
/// pinned preview's caption reads like the caption of anything else on the desktop.
const CAPTION_FACE: &str = "Segoe UI";

/// How wide a caption button is, in the units a display's scale multiplies: the width
/// Windows 11 gives one, so the three land where a hand expects them.
const BUTTON_PIXELS: f32 = 46.0;

/// The room a tooltip takes around its text, and how far it floats below the button it
/// belongs to, in the units a display's scale multiplies.
const TITLE_PADDING_PIXELS: f32 = 12.0;
const GLYPH_STROKE_PIXELS: f32 = 1.4;
const GLYPH_PIXELS: f32 = 10.0;
const TOOLTIP_PADDING_PIXELS: f32 = 8.0;
const TOOLTIP_GAP_PIXELS: f32 = 6.0;
const TOOLTIP_RADIUS: f32 = 4.0;

/// The room a volume popup takes at a display's scale: the panel that floats over the media
/// above the bar, and the groove the level is drawn in inside it.
///
/// The panel is a strip of glass over somebody else's picture, so it is kept no wider than it has
/// to be: the thumb and its collar, and a couple of pixels of panel either side of them.
const VOLUME_PANEL_WIDTH: i32 = 24;
const VOLUME_PANEL_HEIGHT: i32 = 98;
/// How far the groove is held off each end of the panel, which is the room the thumb takes at
/// either end of it: a thumb that was clipped by the panel at 100% would be a level drawn short,
/// and anything beyond that room is panel being drawn over a picture it is covering.
const VOLUME_PANEL_INSET: i32 = 14;
/// How far the panel floats above the bar it belongs to.
const VOLUME_PANEL_GAP: i32 = 8;
const VOLUME_PANEL_RADIUS: f32 = 9.0;
const VOLUME_TRACK_WIDTH: i32 = 6;
/// The radius of the button's own wash: the button is a chip rather than the square the strip's
/// room for it would otherwise draw, since it is a control a hand comes back to.
const VOLUME_BUTTON_RADIUS: i32 = 7;
/// The radius of the knob on the groove, which is the bar's own thumb made rounder: it is held
/// rather than aimed at, so it is drawn as something a finger fits.
const VOLUME_THUMB_RADIUS: f32 = 8.0;
/// How much wider the collar around the knob is than the knob itself: the ring of panel colour
/// that keeps a knob drawn over the filled part of the groove reading as a knob.
const VOLUME_COLLAR_PIXELS: f32 = 1.5;

/// The colors a caption and a bubble are painted in: the theme's page, its text, and one it
/// spends on something else, which is drawn as the accent.
///
/// They come from the theme the text previews are painted with — One Dark Pro, Atom One Light,
/// or a `.tmTheme` of the user's own — so that the chrome of a pinned preview belongs to the
/// same app as the page inside it.
pub(crate) struct ChromePalette {
    /// The caption's bar and the bubble's face.
    pub(crate) background: [u8; 3],
    /// The title and the glyphs.
    pub(crate) foreground: [u8; 3],
    /// The bubble's ring, and the played part of a transport's bar.
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
    /// The button that opens the volume popup, which is the pin's own control rather than the
    /// player's.
    Volume,
}

/// Where the parts of a transport bar are, in the bar's own coordinates: the button, the bar
/// itself, the two labels beside them, and the volume button at the far end.
pub(crate) struct TransportLayout {
    pub(crate) play: RECT,
    pub(crate) bar: RECT,
    pub(crate) elapsed: RECT,
    pub(crate) total: RECT,
    pub(crate) volume: RECT,
}

/// The room each part of a transport bar is given at a display's scale.
///
/// A bar whose player cannot be told anything is laid out without the button: the two clocks and
/// the track begin at the strip's own padding rather than after a control that is not there (see
/// `TransportState::interactive`).
pub(crate) fn transport_layout(
    width: i32,
    height: i32,
    dpi: u32,
    interactive: bool,
) -> TransportLayout {
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
    // The volume button is the last thing on the bar, against the right edge the way it is on
    // every player that has one — and it keeps its box whether or not the bar it sits on has
    // controls, since it is this app's own control rather than the player's (see
    // `TransportPart::Volume`).
    let volume = RECT {
        left: (width - padding - play_side).max(padding),
        top: (height - play_side) / 2,
        right: (width - padding).max(padding),
        bottom: (height + play_side) / 2,
    };
    let content_left = if interactive {
        play.right + gap
    } else {
        padding
    };
    let elapsed = RECT {
        left: content_left,
        top: 0,
        right: content_left + label,
        bottom: height,
    };
    let total = RECT {
        left: (volume.left - gap - label).max(elapsed.right),
        top: 0,
        right: (volume.left - gap).max(elapsed.right),
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
        volume,
    }
}

/// Which part of a transport bar a point is on, in the bar's own coordinates.
///
/// The volume button is answered on a bar that is not `interactive` as well, and it is answered
/// first: it is not the player's control but this app's, and what it does — take the level it is
/// given, and start the player again at it where the player cannot be told anything — is a thing
/// both engines can be asked. A bar that is not `interactive` has no button and no track for a
/// press to have found, which is what the early answer after it says.
pub(crate) fn transport_part_at(
    x: i32,
    y: i32,
    width: i32,
    height: i32,
    dpi: u32,
    interactive: bool,
) -> Option<TransportPart> {
    let layout = transport_layout(width, height, dpi, interactive);
    let inside = |rect: RECT| x >= rect.left && x < rect.right && y >= rect.top && y < rect.bottom;

    if inside(layout.volume) {
        return Some(TransportPart::Volume);
    }

    if !interactive {
        return None;
    }

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

/// Where a volume popup is, in the window's own coordinates: the panel it floats in, and the
/// groove inside it the level is drawn against.
pub(crate) struct VolumePopup {
    pub(crate) panel: RECT,
    pub(crate) track: RECT,
}

/// The popup a volume button opens, as it is placed against the strip the button sits in: a panel
/// floating over the media above the bar, hung from the button's own right edge and kept inside
/// the window it belongs to — a popup drawn off the side of the window is a level that cannot be
/// dragged to.
///
/// `strip_top` is the row the transport bar begins at in the window and `strip_height` how tall
/// that strip is, which is what the button's own box is computed from, so the popup and the
/// button cannot drift apart: what opens it is drawn where it was pressed.
pub(crate) fn volume_popup_layout(
    width: i32,
    strip_top: i32,
    strip_height: i32,
    dpi: u32,
) -> VolumePopup {
    // The button is the transport bar's own, and the bar begins `strip_top` rows down the
    // window, so the box it opens from is the strip's box moved down the window — which is what
    // makes the two answers one answer (see `volume_popup_from_button`).
    let in_strip = transport_layout(width, strip_height, dpi, true).volume;
    let button = RECT {
        top: in_strip.top + strip_top,
        bottom: in_strip.bottom + strip_top,
        ..in_strip
    };

    volume_popup_from_button(button, width, dpi)
}

/// The popup a volume button opens, hung from that button's own box in the window's own
/// coordinates: the same panel the transport bar's button opens, from wherever the button that
/// opened it is drawn.
///
/// Split from `volume_popup_layout` so that a button drawn on a sound's card can open the very
/// same panel — the card's own row is a window away from the transport bar (see
/// `audio_preview::CardControl::Volume`), and a second panel for it would be a second set of
/// numbers for one control this app has once.
pub(crate) fn volume_popup_from_button(button: RECT, width: i32, dpi: u32) -> VolumePopup {
    let scale = dpi as f32 / 96.0;

    let panel_width = text_paint::scaled(VOLUME_PANEL_WIDTH, scale).max(8);
    let panel_height = text_paint::scaled(VOLUME_PANEL_HEIGHT, scale).max(8);
    let gap = text_paint::scaled(VOLUME_PANEL_GAP, scale).max(1);

    let right = button.right.clamp(0, width.max(0));
    let left = (right - panel_width).max(0);
    let right = left + panel_width;
    // The panel floats above the button rather than over it: a level is aimed at from the side
    // it is heard on, and a panel covering the button that opened it is a panel the hand has to
    // reach through to close again.
    let bottom = (button.top - gap).max(panel_height);
    let panel = RECT {
        left,
        top: bottom - panel_height,
        right,
        bottom,
    };

    // The groove, which is what the level is measured on: it is held off both ends of the panel
    // by the room the thumb takes, so a level of nothing and a level of everything are both drawn
    // whole, and it can never be shorter than a level that can be aimed at.
    let track_width = text_paint::scaled(VOLUME_TRACK_WIDTH, scale).max(2);
    let inset = text_paint::scaled(VOLUME_PANEL_INSET, scale)
        .min(((panel.bottom - panel.top) - 8) / 2)
        .max(1);
    let center = (panel.left + panel.right) / 2;
    let track = RECT {
        left: center - track_width / 2,
        top: panel.top + inset,
        right: center - track_width / 2 + track_width,
        bottom: panel.bottom - inset,
    };

    VolumePopup { panel, track }
}

/// Where along a volume popup's groove a point is, as a share of the level: the bottom of the
/// groove is nothing and the top of it is everything, which is the way a level is read.
pub(crate) fn volume_share_at(y: i32, track: RECT) -> f64 {
    let span = (track.bottom - track.top).max(1) as f64;
    ((track.bottom - y) as f64 / span).clamp(0.0, 1.0)
}

/// The row a level is drawn at on a popup's groove.
pub(crate) fn volume_thumb_row(track: RECT, volume: u32) -> i32 {
    let share = volume.min(100) as f64 / 100.0;
    let span = (track.bottom - track.top).max(1) as f64;
    track.bottom - (span * share).round() as i32
}

/// Where along a transport bar a point is, as a share of the file: what a press on the bar asks
/// the player to be taken to.
pub(crate) fn transport_share_at(x: i32, width: i32, dpi: u32, interactive: bool) -> f64 {
    let layout = transport_layout(width, 0, dpi, interactive);
    let span = (layout.bar.right - layout.bar.left).max(1) as f64;

    ((x - layout.bar.left) as f64 / span).clamp(0.0, 1.0)
}

/// What a transport bar is drawn from.
pub(crate) struct TransportState {
    /// Whether the player behind the bar can be told anything at all.
    ///
    /// FFmpeg's player cannot: it reports no position, takes no pause, and can only be taken to
    /// another second by being ended and begun again — so a pinned video it plays carries a bar
    /// with no button and nothing to drag, and what is left is a read-out drawn from this app's
    /// own clock over the player's start (see `pin_playhead`). The media engine's bar is the one
    /// with controls, because every one of them is a question it answers.
    pub(crate) interactive: bool,
    /// Whether a player is running, which is what the button's glyph says.
    pub(crate) playing: bool,
    /// Where the playhead is, and how long the file is: nothing for either is a bar with no
    /// length to draw a playhead against.
    pub(crate) position: Option<f64>,
    pub(crate) duration: Option<f64>,
    pub(crate) hovered: Option<TransportPart>,
    pub(crate) pressed: Option<TransportPart>,
    /// The level this pin's own volume control is at, and whether its popup is open. The button
    /// carries both: a level of nothing is a speaker with no sound coming out of it, and an open
    /// popup is a button held down rather than one a hand is merely over (`paint_volume_button`).
    pub(crate) volume: u32,
    pub(crate) volume_open: bool,
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
    let layout = transport_layout(width, height, dpi, state.interactive);

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

    // The button is drawn only where there is a player to press it: a bar with no controls is a
    // read-out, and a button that does nothing is a promise the app cannot keep.
    if state.interactive {
        paint_play_button(surface, palette, state, layout.play, scale);
    }

    // The volume button, which is on the bar whether or not the player behind it answers
    // anything: the level is this app's to keep, and both engines are given it (see
    // `TransportPart::Volume`).
    paint_volume_button(surface, palette, state, layout.volume, scale);

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
                surface_pixels(surface),
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
            surface_pixels(surface),
            surface.width as usize * surface.height as usize * 4,
        )
    };
    for pixel in buffer.as_chunks_mut::<4>().0 {
        pixel[3] = 255;
    }
}

/// The two bars a play/pause button shows while the thing behind it is playing, which is the
/// whole of what "pause" is drawn as: two uprights side by side, held apart far enough that they
/// do not read as one mark.
fn draw_pause_glyph(
    buffer: &mut [u8],
    width: i32,
    center_x: i32,
    center_y: i32,
    size: i32,
    ink: [u8; 3],
) {
    let bar = (size / 3).max(2);
    let gap = (size / 5).max(1);
    for left in [center_x - gap - bar, center_x + gap] {
        fill_box(
            buffer,
            width,
            RECT {
                left,
                top: center_y - size / 2,
                right: left + bar,
                bottom: center_y + size / 2,
            },
            ink,
            1.0,
        );
    }
}

/// The triangle a play/pause button shows while the thing behind it is stopped, drawn as a scan
/// of half-widths, widest at its left edge.
fn draw_play_glyph(
    buffer: &mut [u8],
    width: i32,
    center_x: i32,
    center_y: i32,
    size: i32,
    ink: [u8; 3],
) {
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

/// One step of a walk through files, as a transport draws it: a bar against the side the step
/// comes from, and a triangle pointing the way it goes — ⏮ and ⏭, the marks every player has
/// drawn for this since there were players.
///
/// Not the caption's chevron beside it, and the difference is a claim about two different
/// things rather than a matter of taste. A chevron on a caption is the walk with no end, drawn as
/// one open mark because a caption's row is five buttons and a name, and a pair of uprights
/// there would be read as a resize handle; on a transport it is a control with room round it,
/// and the mark a hand has been reaching for beside a play button for thirty years is the bar
/// and the triangle. Drawing the caption's mark beside a play button would put two marks for one
/// thing on one row, a hand apart, which is what a row of four controls is not for.
fn draw_track_step(
    buffer: &mut [u8],
    width: i32,
    center_x: i32,
    center_y: i32,
    span: i32,
    color: [u8; 3],
    forward: bool,
) {
    let bar_width = (span / 4).max(1);
    let gap = (span / 6).max(1);
    let head = (span - bar_width - gap).max(1);
    let half = span / 2;

    // The bar: at the far edge of the mark, which is the leading edge for a step forward and the
    // trailing one for a step back, so that the bar is always the wall the triangle is pushed off.
    let bar_left = if forward {
        center_x + half - bar_width
    } else {
        center_x - half
    };
    fill_box(
        buffer,
        width,
        RECT {
            left: bar_left,
            top: center_y - half,
            right: bar_left + bar_width,
            bottom: center_y + half,
        },
        color,
        1.0,
    );

    // The triangle, walked as a scan of half-widths in the direction the step goes: widest at the
    // far edge and closing to nothing at the bar, which is what a triangle is.
    //
    // The far edge is the last *column* of the mark rather than its first, on both sides: the mark
    // is a whole number of columns wide, so the two ends it has are a pixel either side of the
    // centreline rather than at it. Measuring the far edge from the centre instead is what makes
    // the back one a column wider than the on one and its triangle sit across its own bar.
    let far = if forward {
        center_x - half
    } else {
        center_x + half - 1
    };
    for step in 0..=head {
        let x = if forward { far + step } else { far - step };
        let share = 1.0 - (step as f32 / head as f32);
        for y in 0..=(half as f32 * share).round() as i32 {
            put(buffer, width, x, center_y - y, color, 1.0);
            put(buffer, width, x, center_y + y, color, 1.0);
        }
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
            surface_pixels(surface),
            surface.width as usize * surface.height as usize * 4,
        )
    };

    // One glyph or the other, whichever the player behind the button says the truth is: the two
    // are the transport's own, and a card's row draws them at the size its row gives them
    // (see `paint_card_control`).
    if state.playing {
        draw_pause_glyph(buffer, width, center_x, center_y, size, ink);
    } else {
        draw_play_glyph(buffer, width, center_x, center_y, size, ink);
    }
}

/// Paint the volume button: a speaker with the sound leaving it, or the same speaker crossed out
/// where the level is nothing.
///
/// What the button says is the level this pin is playing at rather than what the tray's setting
/// is, and that is the whole reason it is drawn here: the level belongs to the window it is on
/// (see `PinnedPreview::volume`), so the bar is the only place it can be read off.
fn paint_volume_button(
    surface: &DibSurface,
    palette: &ChromePalette,
    state: &TransportState,
    rect: RECT,
    scale: f32,
) {
    let pressed = state.pressed == Some(TransportPart::Volume);
    let buffer = unsafe {
        std::slice::from_raw_parts_mut(
            surface_pixels(surface),
            surface.width as usize * surface.height as usize * 4,
        )
    };
    let width = surface.width as i32;

    // A button a hand is over lights up, and a button whose popup is open stays lit while it is:
    // the pointer is on the level by then, and what the popup came out of should still say where
    // it belongs. The wash is the button's own shape rather than the strip it sits in — a square
    // of a colour around a speaker is a box, and this is one of the two controls on the bar a hand
    // is meant to find.
    if state.hovered == Some(TransportPart::Volume) || pressed || state.volume_open {
        let radius = text_paint::scaled(VOLUME_BUTTON_RADIUS, scale) as f32;
        let wash = if pressed || state.volume_open {
            palette.hover(0.18)
        } else {
            palette.hover(0.10)
        };
        fill_round_rect(buffer, width, rect, radius, wash, wash, 1.0);
    }

    let center_x = (rect.left + rect.right) as f32 / 2.0;
    let center_y = (rect.top + rect.bottom) as f32 / 2.0;
    let side = (rect.bottom - rect.top).max(8);
    let size = (text_paint::scaled(17, scale) as f32).clamp(9.0, side as f32);
    let stroke = (GLYPH_STROKE_PIXELS * scale * 1.1).max(1.0);

    paint_speaker(
        buffer,
        width,
        (center_x, center_y),
        size,
        palette.foreground,
        stroke,
        state.volume,
    );
}

/// The speaker a volume button is drawn as: a body with a cone on it, and the sound leaving the
/// cone as one or two arcs — whose number is how loud it is — or a cross where it is not.
///
/// How many arcs a level gets is the icon every player uses: at 1% and at 34% the same mark is
/// drawn and neither is wrong, and what a button is read for at a glance is whether there is any
/// sound at all.
fn paint_speaker(
    buffer: &mut [u8],
    width: i32,
    center: (f32, f32),
    size: f32,
    ink: [u8; 3],
    stroke: f32,
    volume: u32,
) {
    let (center_x, center_y) = center;
    let half = size / 2.0;
    let point = |x: f32, y: f32| (center_x + x * half, center_y + y * half);

    // The body: the back of the speaker and the cone that leaves it, as one shape.
    fill_polygon(
        buffer,
        width,
        &[
            point(-0.78, -0.26),
            point(-0.30, -0.26),
            point(0.10, -0.62),
            point(0.10, 0.62),
            point(-0.30, 0.26),
            point(-0.78, 0.26),
        ],
        ink,
        1.0,
    );

    if volume == 0 {
        // Crossed out, and crossed out where the sound would be: a speaker drawn without its
        // arcs reads as a small icon rather than as one that is making no noise.
        stroke_segment(
            buffer,
            width,
            point(0.30, -0.34),
            point(0.80, 0.34),
            stroke,
            ink,
            1.0,
        );
        stroke_segment(
            buffer,
            width,
            point(0.30, 0.34),
            point(0.80, -0.34),
            stroke,
            ink,
            1.0,
        );
        return;
    }

    let center = point(0.14, 0.0);
    stroke_arc(buffer, width, center, half * 0.44, stroke, ink, 1.0);
    if volume >= 34 {
        stroke_arc(buffer, width, center, half * 0.86, stroke, ink, 1.0);
    }
}

/// One of the marks this app's own transport is drawn in, named so that a button which has to
/// say one of them does not have to know how: the two a play/pause button is either, the pair a
/// step through files is, and the speaker a level is.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum ControlGlyph {
    /// A triangle: the thing behind it is stopped.
    Play,
    /// Two bars: the thing behind it is going.
    Pause,
    /// A bar and a triangle pointing back: the step before this one.
    Previous,
    /// A bar and a triangle pointing on: the step after this one.
    Next,
    /// A speaker with the sound leaving it, crossed out at a level of nothing.
    Volume(u32),
}

/// Paint one of a sound's card's own control buttons into a page rather than into a strip of the
/// window: the marks are the transport's, drawn at the size a card's row gives them and in the
/// page's own foreground, because a button on a card is a control of that page and not a strip of
/// chrome beside it (see `audio_preview::CardControl`).
///
/// It draws the mark and nothing else. The wash under a pointer is the card's own paint, drawn
/// from the theme the page is painted with — this module's palette is the caption's, which a
/// card does not have — and a wash drawn here as well would be the same button lit twice at two
/// colours.
///
/// Everything here is written premultiplied (see the module documentation), which a page cannot
/// see: `DibSurface::pixels` forces the alpha byte of every pixel it hands back to 255, so a
/// mark written at full coverage reads the same either way. Every glyph below is therefore drawn
/// at coverage 1.0 only — a mark at partial coverage would be composited twice over the page's
/// own opaque pixel and come out darker than the same mark on a strip.
pub(crate) fn paint_card_control(
    surface: &DibSurface,
    rect: RECT,
    glyph: ControlGlyph,
    ink: [u8; 3],
    scale: f32,
) {
    let width = surface.width as i32;
    let center_x = (rect.left + rect.right) / 2;
    let center_y = (rect.top + rect.bottom) / 2;
    let side = (rect.bottom - rect.top).max(8);

    let buffer = unsafe {
        std::slice::from_raw_parts_mut(
            surface_pixels(surface),
            surface.width as usize * surface.height as usize * 4,
        )
    };

    match glyph {
        ControlGlyph::Play => draw_play_glyph(
            buffer,
            width,
            center_x,
            center_y,
            size_in(scale, side, 9),
            ink,
        ),
        ControlGlyph::Pause => draw_pause_glyph(
            buffer,
            width,
            center_x,
            center_y,
            size_in(scale, side, 9),
            ink,
        ),
        ControlGlyph::Previous => draw_track_step(
            buffer,
            width,
            center_x,
            center_y,
            size_in(scale, side, 12),
            ink,
            false,
        ),
        ControlGlyph::Next => draw_track_step(
            buffer,
            width,
            center_x,
            center_y,
            size_in(scale, side, 12),
            ink,
            true,
        ),
        // The speaker is drawn at the transport's own size rather than the play button's: it is
        // the widest mark of the three, and at the smaller size it loses the cone it is read by.
        ControlGlyph::Volume(volume) => paint_speaker(
            buffer,
            width,
            (center_x as f32, center_y as f32),
            size_in(scale, side, 17) as f32,
            ink,
            (GLYPH_STROKE_PIXELS * scale * 1.1).max(1.0),
            volume,
        ),
    }
}

/// The size a mark of `nominal` pixels is drawn at inside a box of `side` pixels, never larger
/// than the box and never smaller than the floor every mark in this file shares: a mark smaller
/// than that is a speck, and a mark larger than its button is a mark on the card beside it.
fn size_in(scale: f32, side: i32, nominal: i32) -> i32 {
    text_paint::scaled(nominal, scale).clamp(6, side)
}

/// Paint the popup a volume button opens: a panel floating over the media above the bar, the
/// groove in it, the level the groove has been taken to, and the knob that is dragged.
///
/// It is written straight into the window's own surface rather than into a strip of its own,
/// because that is what it is: a thing drawn over the picture, with the picture behind it. What
/// is behind it is the media wherever the media is this app's pixels and the player's own window
/// everywhere else, which is what brings the pin's window above the player's while it is open
/// (see `pin_volume_open`).
pub(crate) fn paint_volume_popup(
    buffer: &mut [u8],
    width: i32,
    palette: &ChromePalette,
    popup: &VolumePopup,
    volume: u32,
    held: bool,
) {
    let panel_height = (popup.panel.bottom - popup.panel.top).max(1) as f32;
    let scale = panel_height / VOLUME_PANEL_HEIGHT as f32;
    let radius = (VOLUME_PANEL_RADIUS * scale).max(1.0);
    let thumb = (VOLUME_THUMB_RADIUS * scale).max(2.0);

    // The shadow: the panel's own shape held a couple of pixels lower, so that a panel over a
    // picture reads as floating over it rather than as a hole cut in it.
    let shadow = (2.0 * scale).round().max(1.0) as i32;
    fill_round_rect(
        buffer,
        width,
        RECT {
            left: popup.panel.left,
            top: popup.panel.top + shadow,
            right: popup.panel.right,
            bottom: popup.panel.bottom + shadow,
        },
        radius,
        [0, 0, 0],
        [0, 0, 0],
        0.32,
    );

    // The panel: the theme's page colour, a shade lighter at the top than at the bottom so that
    // it reads as a surface a knob sits on. A hairline around it keeps it off a video of a colour
    // close to its own.
    fill_round_rect(
        buffer,
        width,
        popup.panel,
        radius,
        palette.hover(0.08),
        palette.hover(0.0),
        0.97,
    );
    stroke_round_rect(
        buffer,
        width,
        popup.panel,
        radius,
        (1.0 * scale).round().max(1.0),
        palette.hover(0.32),
        1.0,
    );

    // The groove the level is read against, and the part of it that has been reached: the fill
    // runs from the knob down, which is the way a level is filled in.
    let track_radius = (popup.track.right - popup.track.left) as f32 / 2.0;
    let groove = palette.hover(0.22);
    fill_round_rect(
        buffer,
        width,
        popup.track,
        track_radius,
        groove,
        groove,
        0.9,
    );

    let row = volume_thumb_row(popup.track, volume);
    if volume > 0 {
        fill_round_rect(
            buffer,
            width,
            RECT {
                left: popup.track.left,
                top: row,
                right: popup.track.right,
                bottom: popup.track.bottom,
            },
            track_radius,
            palette.accent,
            palette.accent,
            1.0,
        );
    }

    // The knob, with the panel's own colour for a collar: it is drawn over the filled part of the
    // groove as often as over the empty one, and the collar is what keeps it a knob either way.
    let center_x = (popup.track.left + popup.track.right) as f32 / 2.0;
    let center_y = row as f32;
    fill_disc(
        buffer,
        width,
        center_x,
        center_y,
        thumb + (VOLUME_COLLAR_PIXELS * scale).max(1.0),
        palette.hover(0.0),
        1.0,
    );
    let ink = if held {
        palette.accent
    } else {
        palette.foreground
    };
    fill_disc(buffer, width, center_x, center_y, thumb, ink, 1.0);
}

/// How wide a run of text is in a caption's face at a display's scale, measured through the
/// same font the run is drawn in.
///
/// It is a question rather than an estimate because the text is a program's name: "Open With
/// Adobe Photoshop" and "Open With Photos" are nowhere near the same width, and a panel sized
/// for one of them clips the other.
pub(crate) fn measure_caption_text(surface: &DibSurface, text: &str, dpi: u32) -> i32 {
    if text.is_empty() {
        return 0;
    }

    measure_text(surface, &caption_style([0, 0, 0]), text, dpi as f32 / 96.0)
}

/// The style every piece of caption text is drawn in, whatever colour it is asked for.
fn caption_style(foreground: [u8; 3]) -> text_paint::TextStyle {
    text_paint::TextStyle {
        foreground,
        background: None,
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        level: text_paint::BODY_LEVEL,
        face: CAPTION_FACE,
    }
}

/// Fill a rounded rectangle, shaded from one colour at its top edge to another at its bottom: a
/// popup panel is the one thing here that is not a flat surface.
fn fill_round_rect(
    buffer: &mut [u8],
    width: i32,
    rect: RECT,
    radius: f32,
    top: [u8; 3],
    bottom: [u8; 3],
    coverage: f32,
) {
    if rect.right <= rect.left || rect.bottom <= rect.top {
        return;
    }

    let height = (rect.bottom - rect.top) as f32;
    for y in rect.top..rect.bottom {
        let color = text_paint::blend(top, bottom, (y - rect.top) as f32 / height);
        for x in rect.left..rect.right {
            let coverage = coverage * round_rect_coverage(x, y, rect, radius);
            if coverage > 0.0 {
                put(buffer, width, x, y, color, coverage);
            }
        }
    }
}

/// The outline of a rounded rectangle, drawn inside its own edge so that the shape under it is
/// the size it was asked for.
fn stroke_round_rect(
    buffer: &mut [u8],
    width: i32,
    rect: RECT,
    radius: f32,
    thickness: f32,
    color: [u8; 3],
    coverage: f32,
) {
    let thickness = thickness.round().max(1.0) as i32;
    let inner = RECT {
        left: rect.left + thickness,
        top: rect.top + thickness,
        right: rect.right - thickness,
        bottom: rect.bottom - thickness,
    };
    if inner.right <= inner.left || inner.bottom <= inner.top {
        return;
    }

    let inner_radius = (radius - thickness as f32).max(0.0);
    for y in rect.top..rect.bottom {
        for x in rect.left..rect.right {
            let outline = (round_rect_coverage(x, y, rect, radius)
                - round_rect_coverage(x, y, inner, inner_radius))
            .max(0.0);
            if outline > 0.0 {
                put(buffer, width, x, y, color, coverage * outline);
            }
        }
    }
}

/// How much of a pixel a rounded rectangle covers, worked out from the distance to its edge: one
/// shape, so that a panel, its outline and the groove in it cannot disagree about where they end.
fn round_rect_coverage(x: i32, y: i32, rect: RECT, radius: f32) -> f32 {
    let half_width = (rect.right - rect.left) as f32 / 2.0;
    let half_height = (rect.bottom - rect.top) as f32 / 2.0;
    if half_width <= 0.0 || half_height <= 0.0 {
        return 0.0;
    }

    let radius = radius.clamp(0.0, half_width.min(half_height));
    let center_x = rect.left as f32 + half_width;
    let center_y = rect.top as f32 + half_height;
    let dx = ((x as f32 + 0.5 - center_x).abs() - (half_width - radius)).max(0.0);
    let dy = ((y as f32 + 0.5 - center_y).abs() - (half_height - radius)).max(0.0);
    let distance = (dx * dx + dy * dy).sqrt() - radius;

    (0.5 - distance).clamp(0.0, 1.0)
}

/// Fill a circle: the knob a level is dragged by, and the collar around it.
fn fill_disc(
    buffer: &mut [u8],
    width: i32,
    center_x: f32,
    center_y: f32,
    radius: f32,
    color: [u8; 3],
    coverage: f32,
) {
    let radius = radius.max(0.5);
    let left = (center_x - radius - 1.0).floor() as i32;
    let right = (center_x + radius + 1.0).ceil() as i32;
    let top = (center_y - radius - 1.0).floor() as i32;
    let bottom = (center_y + radius + 1.0).ceil() as i32;

    for y in top..=bottom {
        for x in left..=right {
            let dx = x as f32 + 0.5 - center_x;
            let dy = y as f32 + 0.5 - center_y;
            let edge = (radius - (dx * dx + dy * dy).sqrt() + 0.5).clamp(0.0, 1.0);
            if edge > 0.0 {
                put(buffer, width, x, y, color, coverage * edge);
            }
        }
    }
}

/// Fill a polygon, sampled rather than scanned: the shapes drawn here are a dozen pixels across,
/// so a handful of samples a pixel costs less than a scanline that has to be right about every
/// edge of a shape as small as a glyph.
fn fill_polygon(
    buffer: &mut [u8],
    width: i32,
    points: &[(f32, f32)],
    color: [u8; 3],
    coverage: f32,
) {
    const SAMPLES: i32 = 4;

    if points.len() < 3 {
        return;
    }

    let left = points
        .iter()
        .map(|p| p.0)
        .fold(f32::INFINITY, f32::min)
        .floor() as i32;
    let right = points
        .iter()
        .map(|p| p.0)
        .fold(f32::NEG_INFINITY, f32::max)
        .ceil() as i32;
    let top = points
        .iter()
        .map(|p| p.1)
        .fold(f32::INFINITY, f32::min)
        .floor() as i32;
    let bottom = points
        .iter()
        .map(|p| p.1)
        .fold(f32::NEG_INFINITY, f32::max)
        .ceil() as i32;

    for y in top..=bottom {
        for x in left..=right {
            let mut inside = 0;
            for sample_y in 0..SAMPLES {
                for sample_x in 0..SAMPLES {
                    let point = (
                        x as f32 + (sample_x as f32 + 0.5) / SAMPLES as f32,
                        y as f32 + (sample_y as f32 + 0.5) / SAMPLES as f32,
                    );
                    if point_in_polygon(point, points) {
                        inside += 1;
                    }
                }
            }

            if inside > 0 {
                let edge = inside as f32 / (SAMPLES * SAMPLES) as f32;
                put(buffer, width, x, y, color, coverage * edge);
            }
        }
    }
}

fn point_in_polygon(point: (f32, f32), points: &[(f32, f32)]) -> bool {
    let mut inside = false;
    let mut previous = points.len() - 1;

    for current in 0..points.len() {
        let (x, y) = points[current];
        let (previous_x, previous_y) = points[previous];

        if (y > point.1) != (previous_y > point.1)
            && point.0 < (previous_x - x) * (point.1 - y) / (previous_y - y) + x
        {
            inside = !inside;
        }

        previous = current;
    }

    inside
}

/// One band of a circle, drawn as the pixels a given distance from a center and within the spread
/// of angles a speaker's sound is drawn in: the two arcs a volume button carries.
fn stroke_arc(
    buffer: &mut [u8],
    width: i32,
    center: (f32, f32),
    radius: f32,
    thickness: f32,
    color: [u8; 3],
    coverage: f32,
) {
    let half = (thickness / 2.0).max(0.5);
    let reach = radius + half + 1.0;
    let left = (center.0 - reach).floor() as i32;
    let right = (center.0 + reach).ceil() as i32;
    let top = (center.1 - reach).floor() as i32;
    let bottom = (center.1 + reach).ceil() as i32;

    // The ends of an arc are square to the axis it opens from, with a sliver of fade so that an
    // end is a soft edge rather than a stair step.
    const SPREAD: f32 = 0.85;
    let feather = 0.14;

    for y in top..=bottom {
        for x in left..=right {
            let dx = x as f32 + 0.5 - center.0;
            let dy = y as f32 + 0.5 - center.1;
            let distance = (dx * dx + dy * dy).sqrt();
            let radial = (half + 0.5 - (distance - radius).abs()).clamp(0.0, 1.0);
            if radial <= 0.0 {
                continue;
            }

            let angle = dy.abs().atan2(dx);
            let spread = ((SPREAD - angle) / feather).clamp(0.0, 1.0);
            if spread > 0.0 {
                put(buffer, width, x, y, color, coverage * radial * spread);
            }
        }
    }
}

/// A line between two points, drawn as the pixels within half its thickness of it.
fn stroke_segment(
    buffer: &mut [u8],
    width: i32,
    from: (f32, f32),
    to: (f32, f32),
    thickness: f32,
    color: [u8; 3],
    coverage: f32,
) {
    let half = (thickness / 2.0).max(0.5);
    let left = (from.0.min(to.0) - half - 1.0).floor() as i32;
    let right = (from.0.max(to.0) + half + 1.0).ceil() as i32;
    let top = (from.1.min(to.1) - half - 1.0).floor() as i32;
    let bottom = (from.1.max(to.1) + half + 1.0).ceil() as i32;

    for y in top..=bottom {
        for x in left..=right {
            let distance = distance_to_segment((x as f32 + 0.5, y as f32 + 0.5), from, to);
            let edge = (half + 0.5 - distance).clamp(0.0, 1.0);
            if edge > 0.0 {
                put(buffer, width, x, y, color, coverage * edge);
            }
        }
    }
}

fn distance_to_segment(point: (f32, f32), from: (f32, f32), to: (f32, f32)) -> f32 {
    let (dx, dy) = (to.0 - from.0, to.1 - from.1);
    let length_squared = dx * dx + dy * dy;
    let along = if length_squared <= f32::EPSILON {
        0.0
    } else {
        (((point.0 - from.0) * dx + (point.1 - from.1) * dy) / length_squared).clamp(0.0, 1.0)
    };

    let nearest = (from.0 + along * dx, from.1 + along * dy);
    ((point.0 - nearest.0).powi(2) + (point.1 - nearest.1).powi(2)).sqrt()
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
    let style = caption_style(palette.foreground);

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
/// The buttons a caption carries, in the order they sit in: the two or three that are a
/// window's own — minimize, maximize or restore, close — and the three that are a pin's,
/// packed to their left.
///
/// The window's are the ones a hand has been reaching for on every window on the desktop,
/// and they are where Windows puts them. The pin's are new: the file before this one, the
/// file after it, and the file handed to whatever the machine has filed it under (see
/// `shell::pin_navigation`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum CaptionButton {
    /// The file before the one pinned, in the order the folder it was taken up in is
    /// showing them.
    Previous,
    /// The file after the one pinned, the same walk the other way.
    Next,
    /// Hand the pinned file to whatever the Shell has filed it under. A pin previews a file
    /// rather than opening it, so this is the only way out of it into the program that owns
    /// the format.
    OpenWith,
    /// Hand the pinned file to a program the user names out loud, rather than to the one the
    /// machine has settled on. Beside `OpenWith` because the two are the same gesture two
    /// directions: this one opens the Shell's own list of what could open it, which is the only
    /// way into a second program when the default is the wrong one — the button beside it can
    /// only ever reach the one the Shell has already chosen.
    OpenWithList,
    Minimize,
    Maximize,
    Close,
}

/// The pin's own four, in the order they are packed. They are separate from the window's
/// group because a window's group is measured by itself and cannot give any of its room:
/// a close button that moved because a pin grew a walk would be a close button in a new
/// place on every caption the moment this app was updated.
const NAV_BUTTONS: [CaptionButton; 4] = [
    CaptionButton::Previous,
    CaptionButton::Next,
    CaptionButton::OpenWith,
    CaptionButton::OpenWithList,
];

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
///
/// The pin's own four — the file before, the file after, the file opened by whatever
/// the machine has filed it under, and the file opened by a program named out loud — are packed
/// beside that group and not inside it, and are dropped whole where a caption is too narrow to
/// carry them. Which of the two is given up is not a question: the walk is a thing a window
/// only has once it is a pin, and closing a pin is not.
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

    // Measured off the window's group alone, so the walk beside it can be any width at all and
    // the three land where they landed before it existed.
    let button = text_paint::scaled(BUTTON_PIXELS as i32, dpi as f32 / 96.0)
        .max(1)
        .min((width / kinds.len().max(1) as i32).max(1));

    let mut boxes: Vec<CaptionButtonBox> = kinds
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
        .collect();

    // The walk is packed against the group it hangs off, as wide as the buttons beside it: a
    // caption's buttons are one size, and a narrower one in the middle of them is a row of
    // targets a hand has to find rather than count.
    //
    // A caption too narrow for all four keeps the window's group and loses the walk whole.
    // The buttons a hand has been reaching for on every window it has ever had are the ones
    // that are not given up, and half a walk beside them is a set of targets with no known
    // order to them.
    if width - (kinds.len() as i32 + NAV_BUTTONS.len() as i32) * button < 0 {
        return boxes;
    }

    // Packed from the group's own left edge outward, so the walk reads left-to-right in the
    // order it is written: the file before this one, the file after it, the hand-off to
    // another program, and then the hand-off to one named out loud. Walking out from the edge
    // and reversing would put them the other way round, which puts `Next` where a hand reaches
    // for `Previous` — and puts the list beside the default rather than beyond it, which is
    // where a hand that has just tried the default and found it wanting goes next.
    let group_left = width - kinds.len() as i32 * button;
    let mut nav: Vec<CaptionButtonBox> = NAV_BUTTONS
        .iter()
        .enumerate()
        .map(|(index, kind)| {
            let left = group_left - (NAV_BUTTONS.len() as i32 - index as i32) * button;
            CaptionButtonBox {
                kind: *kind,
                rect: RECT {
                    left,
                    top: 0,
                    right: left + button,
                    bottom: height,
                },
            }
        })
        .collect();
    nav.append(&mut boxes);

    nav
}

/// One caption button and the box it occupies.
#[derive(Clone, Copy, PartialEq, Eq)]
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
    let length = surface.width as usize * surface.height as usize * 4;

    // Copied rather than drawn again where the same strip has been asked for before, which is
    // what a repaint is: a pinned window repaints sixty times a second to move a playhead that
    // is drawn on the bar and not on this, and a caption that redrew its own name every one of
    // those frames paid a font create and delete, a handful of `GetTextExtentPoint32W` and a
    // per-pixel walk of every glyph to arrive at the same bytes.
    //
    // The key is every input the drawing below reads, not the ones that seemed to matter: the
    // size and the scale, the four theme colors, the name, and which button the pointer is on
    // and holding down. A key missing the hovered or pressed button would be a caption whose
    // button kept the wash it had when the pointer first crossed it and never gave it back.
    let reused = CAPTION_STRIP.with(|cell| {
        let cached = cell.borrow();
        let Some(strip) = cached.as_ref() else {
            return false;
        };
        if strip.pixels.len() != length || !strip.painted_from(surface, palette, caption, dpi) {
            return false;
        }

        unsafe {
            std::slice::from_raw_parts_mut(surface_pixels(surface), length)
                .copy_from_slice(&strip.pixels);
        }
        true
    });
    if reused {
        return;
    }

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

    // The tooltip is not drawn here: it is a panel of its own hanging below the strip, over
    // the media, and is the window's own business to place (see `tooltip_layout`). All this
    // hands over is which button the name belongs to and what it says.
    //
    // GDI leaves the alpha byte of everything it draws at zero, and a caption is opaque wherever
    // it is painted, so the strip's own coverage is handed back here — after the last of the
    // text rather than after each run, because a caption whose title was sealed and whose
    // tooltip was sealed separately would leave the tooltip's run as a hole in the bar (see the
    // module documentation).
    let pixels = unsafe { std::slice::from_raw_parts_mut(surface_pixels(surface), length) };
    for pixel in pixels.as_chunks_mut::<4>().0 {
        pixel[3] = 255;
    }

    // Kept as it stands rather than as the drawing that made it, so what the next repaint of
    // this same strip copies is the bytes the eye has already been shown — including the
    // coverage the seal above has just handed back, which a copy of the strip before it would
    // not have carried.
    let painted = pixels.to_vec();
    CAPTION_STRIP.with(|cell| {
        *cell.borrow_mut() = Some(CaptionStrip {
            width: surface.width,
            height: surface.height,
            dpi,
            title: caption.title.to_string(),
            maximized: caption.maximized,
            maximizable: caption.maximizable,
            hovered: caption.hovered,
            pressed: caption.pressed,
            background: palette.background,
            foreground: palette.foreground,
            accent: palette.accent,
            dark: palette.dark,
            pixels: painted,
        });
    });
}

// One caption strip kept whole, with everything it was drawn from.
//
// Kept per thread because a strip is painted into a surface the preview thread owns and read
// back off that same one: one cache shared between two windows would be handing each of them
// the other's pixels.
thread_local! {
    static CAPTION_STRIP: RefCell<Option<CaptionStrip>> = const { RefCell::new(None) };
}

/// A caption strip, and the whole of what it was drawn from.
struct CaptionStrip {
    width: u32,
    height: u32,
    dpi: u32,
    title: String,
    maximized: bool,
    maximizable: bool,
    hovered: Option<CaptionButton>,
    pressed: Option<CaptionButton>,
    background: [u8; 3],
    foreground: [u8; 3],
    accent: [u8; 3],
    dark: bool,
    pixels: Vec<u8>,
}

impl CaptionStrip {
    /// Whether this strip is the one asked for: every input the drawing reads, so a theme
    /// change, a resized window, another display's scale, another file, or a pointer that has
    /// moved onto — or off — a button all of them redraw rather than copy.
    ///
    /// All four theme colors are compared and not only the two the strip's own pixels are mostly
    /// made of, because the close button is washed with a red that does not come from the theme
    /// at all and every other wash is the background moved toward the text — which branches on
    /// `dark`. A key left off the background would hand a dark theme's caption to a light one,
    /// and one left off the hovered button would leave the wash under a pointer that has since
    /// walked away.
    fn painted_from(
        &self,
        surface: &DibSurface,
        palette: &ChromePalette,
        caption: &Caption,
        dpi: u32,
    ) -> bool {
        self.width == surface.width
            && self.height == surface.height
            && self.dpi == dpi
            && self.title == caption.title
            && self.maximized == caption.maximized
            && self.maximizable == caption.maximizable
            && self.hovered == caption.hovered
            && self.pressed == caption.pressed
            && self.background == palette.background
            && self.foreground == palette.foreground
            && self.accent == palette.accent
            && self.dark == palette.dark
    }
}

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
/// whole of its box, and its coverage is what separates the letters from the box around them.
fn composite_text_into(surface: &DibSurface, out: &mut [u8], width: i32, panel: RECT) {
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
            surface_pixels(surface),
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
        CaptionButton::Previous => {
            draw_chevron(buffer, width, center_x, center_y, glyph, stroke, ink, false)
        }
        CaptionButton::Next => {
            draw_chevron(buffer, width, center_x, center_y, glyph, stroke, ink, true)
        }
        CaptionButton::OpenWith => {
            draw_open_with(buffer, width, center_x, center_y, glyph, stroke, ink)
        }
        CaptionButton::OpenWithList => {
            draw_open_with_list(buffer, width, center_x, center_y, glyph, stroke, ink)
        }
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

    let style = caption_style(palette.foreground);

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
}

/// The longest beginning of `title` that fits `budget` pixels, with an ellipsis after it.
///
/// Searched rather than walked. A prefix only ever gets wider, so the longest beginning that
/// fits is found in about seven measurements whatever the length of the name — where walking it
/// was one `GetTextExtentPoint32W` per character, over a deep folder path, on every repaint
/// that only moved a playhead.
///
/// It asks a `width` rather than a font, which is what let it be tested at all: the walk it
/// replaced and the search that replaced it had each been written out a second time in the
/// tests, so the tests agreed with the search and said nothing whatever about the one the
/// caption is actually drawn from. One definition, two callers, is the only shape in which
/// breaking this one is visible.
///
/// A name whose first character is already too wide for the ellipsis is not a name with room
/// cut off it, so nothing comes back rather than an ellipsis alone.
fn longest_prefix_that_fits(
    title: &str,
    budget: i32,
    width: &mut dyn FnMut(&str) -> i32,
) -> String {
    let boundaries: Vec<usize> = title
        .char_indices()
        .skip(1)
        .map(|(index, _)| index)
        .collect();

    // How many of the boundaries fit, which is the one the walk used to arrive at by stopping at
    // the first that did not.
    let mut low = 0usize;
    let mut high = boundaries.len();
    while low < high {
        let middle = low + (high - low) / 2;
        if width(&title[..boundaries[middle]]) <= budget {
            low = middle + 1;
        } else {
            high = middle;
        }
    }

    // The last boundary is the one before the final character: a name cut at the very last
    // character is a name with a character missing, and it is cut where the width ran out
    // rather than as late as it could be.
    match low.checked_sub(1).map(|index| boundaries[index]) {
        Some(best) => format!("{}…", &title[..best]),
        None => String::new(),
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

        // One buffer for every measurement below, filled in place. A caption is repainted on
        // every pointer move across it, and the alternative was an allocation per measurement.
        let mut wide: Vec<u16> = Vec::new();
        let measured = |text: &str, wide: &mut Vec<u16>| -> i32 {
            wide.clear();
            wide.extend(text.encode_utf16());
            if wide.is_empty() {
                return 0;
            }
            let mut extent = SIZE::default();
            if GetTextExtentPoint32W(surface.dc, wide, &mut extent).as_bool() {
                extent.cx
            } else {
                0
            }
        };

        let full = measured(title, &mut wide);
        let fitted = if full <= available {
            title.to_string()
        } else {
            let ellipsis = measured("…", &mut wide);
            let budget = (available - ellipsis).max(0);
            longest_prefix_that_fits(title, budget, &mut |text| measured(text, &mut wide))
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

/// The surface's pixels, which every box, disc and glyph here is drawn into.
///
/// # Safety
///
/// The surface must be one GDI is not drawing into, and must not be handed back to GDI
/// until the slice built from this has ended. A raw pointer rather than a `&mut`, because
/// the surface is reached by shared reference and the pixels behind it are the only thing
/// written: the caller gets a slice for the length of its own block, and nothing else.
unsafe fn surface_pixels(surface: &DibSurface) -> *mut u8 {
    surface.bits()
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

/// One chevron pointing left or right, walked a column at a time the way the cross is.
///
/// A chevron rather than an arrowhead: a walk has no end, so what the two buttons mean is
/// which way along it to go and not where it stops.
///
/// The vertex is at the end the chevron points to and the arms open away from it. Measured
/// from the middle of the span instead, the mark is a cross rather than a chevron — which is
/// what a `<` and a `>` would both be drawn as.
///
/// The arms part *half* a span either side of the middle row, so that the mark comes out the
/// size the close cross and the maximize box beside it are drawn at.
///
/// Walked in whole pixels rather than in `stroke_segment` because a vertex is the whole of
/// what a chevron says: an anti-aliased stroke a pixel wide rounds its own point off.
#[allow(clippy::too_many_arguments)] // Which way it points is a fact about the mark, not a drawing flag.
fn draw_chevron(
    buffer: &mut [u8],
    width: i32,
    center_x: i32,
    center_y: i32,
    span: i32,
    thickness: f32,
    color: [u8; 3],
    pointing_right: bool,
) {
    let thickness = thickness.round().max(1.0) as i32;
    let half = span / 2;
    for step in 0..=span.max(1) {
        // The columns are walked the same way whichever way the chevron points — what says
        // which way it points is whether the arms are widest at the first column or the last.
        // Measuring the arms from the middle instead is what draws an X; taking the columns
        // one way and the arms the other is what draws both buttons as the same arrow.
        let x = center_x - half + step;
        let reach = if pointing_right { span - step } else { step };
        // The reach is rounded to the nearer row, so that the column the arms part from is
        // the one row on its own and not the two the rounding would otherwise leave on it.
        // The reach is *half* the span and not the whole of it: arms that reach as far as the
        // mark is walked are drawn a whole span either side of the middle row, which is twice
        // as tall as every other glyph in the bar.
        let offset = (reach * half * 2 + span.max(1)) / (span.max(1) * 2);

        // As many rows either side of the middle as the stroke is wide, taken in opposite
        // directions, so that the two arms are each other turned about the mark's own middle
        // row at any thickness: a stroke that ran the same way on both sides would close the
        // arms into a blob on the row where they meet.
        for depth in 0..thickness {
            put(buffer, width, x, center_y - offset + depth, color, 1.0);
            put(buffer, width, x, center_y + offset - depth, color, 1.0);
        }
    }
}

/// An arrow leaving a box: the file handed to whatever the machine has filed it under.
///
/// Drawn as the box and the arrow together rather than as a box with a mark in it, because the
/// button's whole meaning is the leaving. The box is drawn in every side but the one the arrow
/// goes out of, and the reading it has to carry is *leaving this thing* — a box drawn too small
/// to read as a box with a big arrow on top of it is the upload mark, and a button that means
/// "hand this to another program" must not be drawn as a button that means "send this away".
///
/// The tail starts at the middle of the box and crosses its open side to run out past the
/// corner, so that the arrow is seen leaving rather than hovering over.
fn draw_open_with(
    buffer: &mut [u8],
    width: i32,
    center_x: i32,
    center_y: i32,
    span: i32,
    thickness: f32,
    color: [u8; 3],
) {
    let edge = thickness.round().max(1.0);
    let half = ((span as f32 / 2.0) - edge / 2.0).max(0.0);
    let (middle_x, middle_y) = (center_x as f32 + 0.5, center_y as f32 + 0.5);
    let point = |x: f32, y: f32| (middle_x + x * half, middle_y + y * half);
    let open_at = |y: f32| (point(-1.0, y), point(1.0, y));

    // The file: a box with its top right corner left off, so that the corner the arrow goes
    // out of is the one that is not drawn.
    let (top_left, top_right) = open_at(-1.0);
    let (bottom_left, bottom_right) = open_at(1.0);
    stroke_segment(buffer, width, top_left, top_right, edge, color, 1.0);
    stroke_segment(buffer, width, bottom_left, bottom_right, edge, color, 1.0);
    stroke_segment(buffer, width, top_left, bottom_left, edge, color, 1.0);
    stroke_segment(
        buffer,
        width,
        top_right,
        point(1.0, -1.0 / 3.0),
        edge,
        color,
        1.0,
    );

    // The hand-off: a tail out of the middle of the file, through the open side and past the
    // corner it goes out of, and the head it ends in — one stroke along the top of it and one
    // along the side of it, which is the same pair a chevron is drawn in.
    let tip = point(1.0, -1.0);
    stroke_segment(buffer, width, point(-0.2, 0.2), tip, edge, color, 1.0);
    stroke_segment(buffer, width, tip, point(1.0 / 3.0, -1.0), edge, color, 1.0);
    stroke_segment(buffer, width, tip, point(1.0, -1.0 / 3.0), edge, color, 1.0);
}

/// A list of rows above a chevron pointing down: the file handed to a program named out loud.
///
/// It is not the hand-off mark beside it with a mark added, and the difference is the whole of
/// what the two buttons are: the one before opens the file with the program the machine has
/// already chosen, which is a *going* and is drawn as an arrow leaving a box, while this one
/// opens a list of what else could — which is a *choosing*, and is drawn as the list itself. Two
/// boxes and two arrows on one strip would be a row of targets a hand has to read one at a time,
/// which is what a walk beside the window's own three is not for.
///
/// The rows are three lines of the same length rather than a bulleted list, because a bullet is
/// a dot this size does not survive and a line is. The chevron below them is the same pair of
/// strokes a `Previous` is drawn in, which is what makes the two read as one family.
fn draw_open_with_list(
    buffer: &mut [u8],
    width: i32,
    center_x: i32,
    center_y: i32,
    span: i32,
    thickness: f32,
    color: [u8; 3],
) {
    let edge = thickness.round().max(1.0);

    // The list: three rows across the top of the mark, each as wide as the mark itself, so
    // that it reads as a list rather than as three specks. A row shorter than the chevron
    // below it is a bullet, and a bullet at ten pixels is a dot. It is walked in whole pixels
    // for the reason the chevron is: a row of a list is a row, and a row drawn in halves is
    // a row of gaps at the scale a display runs at.
    let half = span / 2;
    for offset in [0, half / 2, half] {
        for x in (center_x - half)..=(center_x + half) {
            put(buffer, width, x, center_y - half + offset, color, 1.0);
        }
    }

    // The chevron below them, pointing down and drawn as the two arms a `Previous` is drawn
    // in, reaching to the same edge the list's top row does: the list is what is being
    // offered, and this is the sign that there is more of it than the button can hold. A
    // chevron that stopped short of the top row would leave the mark sitting high on its own
    // button, which is the one thing a row of marks beside each other cannot afford.
    let shoulder = half as f32 * 0.8;
    let top = center_y as f32 + half as f32 * 0.35;
    let (left, right) = (
        (center_x as f32 - shoulder, top),
        (center_x as f32 + shoulder, top),
    );
    let tip = (center_x as f32 + 0.5, center_y as f32 + half as f32 + 0.5);
    stroke_segment(buffer, width, left, tip, edge, color, 1.0);
    stroke_segment(buffer, width, right, tip, edge, color, 1.0);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_cut_name_is_the_whole_name_where_it_fits_and_the_longest_beginning_where_it_does_not() {
        // Six pixels a character, so a budget of 24 is four characters and no more. The width is
        // a closure rather than a font so the arithmetic of a name can be pinned down exactly;
        // `a_real_font_cuts_the_name_where_the_measured_width_runs_out` is what pins the font.
        let mut width = |text: &str| text.chars().count() as i32 * 6;

        assert_eq!(longest_prefix_that_fits("abcdef", 24, &mut width), "abcd…");
        assert_eq!(longest_prefix_that_fits("abcdef", 30, &mut width), "abcde…");
        // Nothing fits beside the ellipsis, so nothing is drawn — a lone ellipsis is not a name.
        assert_eq!(longest_prefix_that_fits("abcdef", 0, &mut width), "");
        assert_eq!(longest_prefix_that_fits("abcdef", 5, &mut width), "");
        // One boundary that fits, and the name is kept to its first character.
        assert_eq!(longest_prefix_that_fits("abcdef", 6, &mut width), "a…");
        // A single character has no boundary to cut at, so it is not cut either.
        assert_eq!(longest_prefix_that_fits("a", 0, &mut width), "");
    }

    #[test]
    fn a_cut_name_is_cut_on_a_character_rather_than_inside_one() {
        // A name of multi-byte characters: every boundary the search tries is the start of a
        // character, so a cut cannot land in the middle of one and leave half a glyph.
        let title = "äöüäöüäöü";
        let mut width = |text: &str| text.chars().count() as i32 * 10;

        let cut = longest_prefix_that_fits(title, 25, &mut width);

        assert_eq!(cut, "äö…");
        assert!(title.starts_with(cut.trim_end_matches('…')));
        assert!(cut.is_char_boundary(cut.len()));
    }

    /// The search and the walk it replaced must answer the same thing for every budget, because
    /// the search is the walk with fewer measurements and the whole claim of the change is that
    /// the caption is not shorter for it.
    ///
    /// The property is asked of `longest_prefix_that_fits`, which is the function `fit_title`
    /// draws its name through — a second copy of the search would agree with this loop whatever
    /// the caption did, which is how the three tests above it came to pass with `fit_title`
    /// deleted outright.
    #[test]
    fn the_search_agrees_with_walking_it() {
        let title = "D:\\some\\folder\\a rather long file name.txt";
        let width = |text: &str| text.chars().count() as i32 * 7;

        for budget in 0..(title.chars().count() as i32 * 7) {
            let boundaries: Vec<usize> = title
                .char_indices()
                .skip(1)
                .map(|(index, _)| index)
                .collect();

            let mut walked = 0usize;
            for &index in &boundaries {
                if width(&title[..index]) > budget {
                    break;
                }
                walked = index;
            }

            let expected = if walked == 0 {
                String::new()
            } else {
                format!("{}…", &title[..walked])
            };

            assert_eq!(
                longest_prefix_that_fits(title, budget, &mut |text| width(text)),
                expected,
                "the search and the walk differ at a budget of {budget}"
            );
        }
    }

    /// A caption is cut by what its own font measures, and not by anything a test can arrange:
    /// every width above is six or ten pixels a character because a real one is not, and a name
    /// cut at the wrong place in a real font is a name with its last character gone.
    ///
    /// So the search is asked through `fit_title` on a surface with a font in it — the path
    /// that has no test at all when the search is copied rather than shared, which is what left
    /// this function free to be wrong for as long as it was here.
    #[test]
    fn a_real_font_cuts_the_name_where_the_measured_width_runs_out() {
        let palette = ChromePalette {
            background: [250, 250, 250],
            foreground: [30, 30, 30],
            accent: [10, 90, 200],
            dark: false,
        };
        let surface = DibSurface::create(600, 30).expect("a surface");
        let style = caption_style(palette.foreground);
        let title = "D:\\some\\folder\\a rather long file name.txt";

        // The whole name where it fits, exactly, and something cut a character or two off where
        // it does not — measured rather than counted, so the last few characters go in the order
        // their own widths say they can and not in the order a name reads.
        let whole = measure_text(&surface, &style, title, 1.0);
        assert_eq!(fit_title(&surface, &style, title, whole, 1.0), title);

        let cut = fit_title(&surface, &style, title, whole - 1, 1.0);
        assert!(
            title.starts_with(cut.strip_suffix('…').unwrap_or_default()),
            "`{cut}` is not a beginning of `{title}`"
        );
        assert!(
            title.len() - cut.trim_end_matches('…').len() > 1,
            "a name one pixel too wide loses the ellipsis and at least one character to it"
        );

        // And nothing at all where even the first character will not fit beside the ellipsis,
        // which a search over a made-up width could not have shown.
        assert_eq!(fit_title(&surface, &style, title, 1, 1.0), "");
        assert_eq!(fit_title(&surface, &style, title, 0, 1.0), "");

        // Every room in between: what comes back is a beginning of the name with an ellipsis
        // after it, it is not the whole name, and it is as wide as the room allows — never
        // wider, which is the whole claim of measuring rather than counting.
        for available in 1..whole {
            let fitted = fit_title(&surface, &style, title, available, 1.0);
            assert!(
                title.starts_with(fitted.strip_suffix('…').unwrap_or(&fitted)),
                "`{fitted}` is not a beginning of `{title}`"
            );
            assert_ne!(fitted, title, "a name that does not fit was not cut");
            assert!(
                measure_text(&surface, &style, &fitted, 1.0) <= available,
                "`{fitted}` is wider than the {available} pixels it was given"
            );
        }
    }

    /// The popup the transport bar's own button opens, and the same panel hung from that button
    /// given directly, are one panel and not two: a sound's card opens the very same one from a
    /// button drawn on its own row, so a level that was placed one way there and another way on
    /// the bar would be a level that moves as the file it is playing changes.
    #[test]
    fn the_popup_hung_from_a_button_is_the_one_the_transport_bar_opens() {
        for (width, height, top) in [(600, 30, 0), (600, 40, 200), (320, 30, 55)] {
            for dpi in [96, 144] {
                let strip = transport_layout(width, height, dpi, true).volume;
                let button = RECT {
                    top: strip.top + top,
                    bottom: strip.bottom + top,
                    ..strip
                };

                let panel = volume_popup_layout(width, top, height, dpi);
                let from_button = volume_popup_from_button(button, width, dpi);
                let (a, b) = (
                    RECT {
                        left: panel.panel.left,
                        top: panel.panel.top,
                        right: panel.panel.right,
                        bottom: panel.panel.bottom,
                    },
                    RECT {
                        left: from_button.panel.left,
                        top: from_button.panel.top,
                        right: from_button.panel.right,
                        bottom: from_button.panel.bottom,
                    },
                );
                assert_eq!(a, b, "{width}x{height} at dpi {dpi}");

                let groove = RECT {
                    left: panel.track.left,
                    top: panel.track.top,
                    right: panel.track.right,
                    bottom: panel.track.bottom,
                };
                let hung = RECT {
                    left: from_button.track.left,
                    top: from_button.track.top,
                    right: from_button.track.right,
                    bottom: from_button.track.bottom,
                };
                assert_eq!(groove, hung, "{width}x{height} at dpi {dpi}");

                // And the panel is still where it belongs: hung from the button's own right edge,
                // above it, and inside the window it belongs to.
                assert!(panel.panel.right <= width && panel.panel.left >= 0);
                // And it floats above the button rather than over it — unless the window is too
                // short for that, which is what the panel's own floor is for: a level that cannot
                // be aimed at is worse than one drawn a little high.
                let panel_height = panel.panel.bottom - panel.panel.top;
                assert!(
                    panel.panel.bottom <= button.top || panel.panel.top == 0,
                    "the panel is not hung over the button it came out of"
                );
                assert!(
                    panel.panel.bottom - panel.panel.top == panel_height,
                    "{width}x{height} at dpi {dpi}"
                );
            }
        }
    }

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

        // The walk comes first in the run and the window's three after it, because the
        // window's are the ones against the right edge and the walk hangs off them.
        assert_eq!(buttons.len(), 7);
        assert_eq!(buttons[0].kind, CaptionButton::Previous);
        assert_eq!(buttons[1].kind, CaptionButton::Next);
        assert_eq!(buttons[2].kind, CaptionButton::OpenWith);
        assert_eq!(buttons[3].kind, CaptionButton::OpenWithList);
        assert_eq!(buttons[4].kind, CaptionButton::Minimize);
        assert_eq!(buttons[5].kind, CaptionButton::Maximize);
        assert_eq!(buttons[6].kind, CaptionButton::Close);
        for pair in buttons.windows(2) {
            assert!(pair[0].rect.left < pair[1].rect.left);
        }
        assert_eq!(buttons[6].rect.right, 600);

        // The window's three keep the widths and the places they have always had: 46 pixels
        // each at 100%, packed to 600, which is what a hand has been reaching for on every
        // window it has ever had.
        for (index, left) in [462, 508, 554].into_iter().enumerate() {
            assert_eq!(
                buttons[index + 4].rect,
                RECT {
                    left,
                    top: 0,
                    right: left + 46,
                    bottom: 30,
                },
                "the window's button at {left} has moved"
            );
        }

        // And the walk is the same width as the buttons beside it, laid end to end and
        // touching the group rather than leaving a gap in it.
        assert_eq!(buttons[3].rect.right, buttons[4].rect.left);
        for button in &buttons {
            assert_eq!(
                button.rect.right - button.rect.left,
                46,
                "a caption's buttons are one size"
            );
        }

        // And a point is on the button it looks like it is on, or on none of them.
        let close = buttons[6].rect;
        assert_eq!(
            button_at(close.left + 1, 5, 600, 30, 96, true),
            Some(CaptionButton::Close)
        );
        assert_eq!(button_at(1, 5, 600, 30, 96, true), None);
        assert_eq!(button_at(close.left + 1, 40, 600, 30, 96, true), None);
    }

    /// Every button a caption carries is found by pointing at it. Painting and hit-testing are
    /// both read off the one list of boxes, so a button that is drawn and is not on, or is on
    /// and is not drawn, is a caption the hand and the eye disagree about.
    #[test]
    fn every_button_of_a_caption_is_the_one_under_the_pointer() {
        let buttons = button_boxes(600, 30, 96, true);
        assert_eq!(buttons.len(), 7);

        for button in &buttons {
            let middle_x = (button.rect.left + button.rect.right) / 2;
            assert_eq!(
                button_at(middle_x, 5, 600, 30, 96, true),
                Some(button.kind),
                "the middle of {:?} is not on it",
                button.kind
            );
            // Each end of the box is its own, and the pixel past the last one is the title's.
            assert_eq!(
                button_at(button.rect.left, 0, 600, 30, 96, true),
                Some(button.kind)
            );
            assert_eq!(
                button_at(button.rect.right - 1, 29, 600, 30, 96, true),
                Some(button.kind)
            );
        }

        assert_eq!(
            button_at(buttons[0].rect.left - 1, 5, 600, 30, 96, true),
            None,
            "the strip to the left of the walk is the title's"
        );
    }

    /// A caption too narrow for the walk carries the window's buttons and nothing else. A
    /// pin is worth less than a window it can be closed from, and half a walk is a set of
    /// targets with no known order to them — so the whole group goes, and the buttons that
    /// stay are exactly where they would have been on a caption wide enough to carry it.
    #[test]
    fn a_caption_narrower_than_its_own_walk_keeps_the_window_s_buttons() {
        let narrow = button_boxes(200, 30, 96, true);

        // The buttons are 46 wide either way here — 200 / 3 is 66, and the width is the
        // smaller of the two — so the walk's own four would need 7 * 46 = 322 of a strip
        // 200 wide, and the strip is the window's alone.
        assert_eq!(narrow.len(), 3);
        assert_eq!(narrow[0].kind, CaptionButton::Minimize);
        assert_eq!(narrow[1].kind, CaptionButton::Maximize);
        assert_eq!(narrow[2].kind, CaptionButton::Close);
        assert_eq!(narrow[2].rect.right, 200);

        // And the pointer agrees with the painter about what is there.
        for button in &narrow {
            assert_eq!(
                button_at(button.rect.left + 1, 5, 200, 30, 96, true),
                Some(button.kind)
            );
        }
        assert_eq!(button_at(0, 5, 200, 30, 96, true), None);
        for kind in [
            CaptionButton::Previous,
            CaptionButton::Next,
            CaptionButton::OpenWith,
            CaptionButton::OpenWithList,
        ] {
            assert!(
                narrow.iter().all(|button| button.kind != kind),
                "a caption too narrow for the walk does not carry {kind:?}"
            );
        }
    }

    /// A pin with nothing to maximize — a sound's card — carries the two buttons that mean
    /// something on it, packed against the right edge the way Windows packs a window's: the one
    /// that closes it keeps the place it has on every other caption, and the one beside it is
    /// the minimize that was there before. Neither of them moves for the walk, which sits
    /// against the group and is measured off the same strip.
    #[test]
    fn a_caption_without_a_maximize_keeps_the_close_button_where_it_was() {
        let three = button_boxes(600, 30, 96, true);
        let two = button_boxes(600, 30, 96, false);

        assert_eq!(two.len(), 6);
        assert_eq!(two[0].kind, CaptionButton::Previous);
        assert_eq!(two[1].kind, CaptionButton::Next);
        assert_eq!(two[2].kind, CaptionButton::OpenWith);
        assert_eq!(two[3].kind, CaptionButton::OpenWithList);
        assert_eq!(two[4].kind, CaptionButton::Minimize);
        assert_eq!(two[5].kind, CaptionButton::Close);

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

        assert_eq!(two[5].rect.left, close.left);
        assert_eq!(two[5].rect.right, close.right);
        assert_eq!(two[4].rect.right, two[5].rect.left);
        assert_eq!(
            two[4].rect.right - two[4].rect.left,
            minimize.right - minimize.left
        );

        // And the space the button used to take is a button's, not a hole: it is the minimize
        // that has moved along into it, with the walk along beside it.
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
        let layout = transport_layout(800, 30, 96, true);
        let middle = (layout.bar.left + layout.bar.right) / 2;
        let share = transport_share_at(middle, 800, 96, true);
        assert!(
            (share - 0.5).abs() < 0.05,
            "the middle of the bar is half of it"
        );

        // A press past either end is the end it is past, which is what keeps a drag from asking
        // for a second of a file that is not there.
        assert_eq!(transport_share_at(0, 800, 96, true), 0.0);
        assert_eq!(transport_share_at(800, 800, 96, true), 1.0);
    }

    /// The volume button is the last thing on the bar and it answers a press whatever the player
    /// behind the bar can be told: a level is this app's own, and FFmpeg's player takes one as well
    /// as the media engine does — by being started again at it.
    #[test]
    fn the_volume_button_answers_on_a_bar_with_no_other_controls() {
        let live = transport_layout(800, 30, 96, true);
        let readout = transport_layout(800, 30, 96, false);

        // It keeps its own box at the right edge of the strip, and it is the same box whether or
        // not the bar carries a play button: a control that moved because the engine changed would
        // be one a hand has to look for.
        assert_eq!(live.volume, readout.volume);
        assert_eq!(
            live.volume.right,
            800 - 10,
            "against the strip's own padding"
        );
        assert!(
            live.volume.left > live.total.right,
            "the clocks end before it"
        );

        let x = (live.volume.left + live.volume.right) / 2;
        let y = (live.volume.top + live.volume.bottom) / 2;
        assert_eq!(
            transport_part_at(x, y, 800, 30, 96, true),
            Some(TransportPart::Volume)
        );
        assert_eq!(
            transport_part_at(x, y, 800, 30, 96, false),
            Some(TransportPart::Volume),
            "and on a bar whose player cannot be told anything"
        );

        // What the button did not take is not the track's: the room it occupies comes off the end
        // the clocks were drawn at.
        assert!(readout.total.right <= readout.volume.left);
        assert!(readout.bar.right <= readout.total.left);
    }

    /// The popup hangs over the button that opened it, inside the window it belongs to, and its
    /// groove is what the level is measured against: the bottom of it is nothing and the top of it
    /// is everything.
    #[test]
    fn a_volume_popup_is_placed_over_its_button_and_read_bottom_up() {
        let (width, strip_top, strip_height) = (800, 400, 30);
        let popup = volume_popup_layout(width, strip_top, strip_height, 96);
        let button = transport_layout(width, strip_height, 96, true).volume;

        // Above the bar, hung from the button's own right edge, and clear of the button itself.
        assert_eq!(popup.panel.right, button.right);
        assert_eq!(
            popup.panel.bottom,
            strip_top + button.top - VOLUME_PANEL_GAP
        );
        assert!(popup.panel.top >= 0);
        assert_eq!(
            popup.panel.right - popup.panel.left,
            VOLUME_PANEL_WIDTH,
            "the panel is the width a level is drawn in"
        );

        // The groove is inside the panel and centered in it, with the room the knob takes at either
        // end: a level of everything is a knob that is still whole. The room is the knob's own
        // radius, and it is asked for with the collar the knob is drawn with rather than tightly.
        assert!(popup.track.left > popup.panel.left && popup.track.right < popup.panel.right);
        assert_eq!(
            (popup.track.left + popup.track.right) / 2,
            (popup.panel.left + popup.panel.right) / 2
        );

        let collar = (VOLUME_THUMB_RADIUS + VOLUME_COLLAR_PIXELS).ceil() as i32;
        let ends = popup.track.top - popup.panel.top;
        assert!(
            ends >= collar && popup.panel.bottom - popup.track.bottom >= collar,
            "the groove is held off each end of the panel by the knob and its collar"
        );

        // And the panel is kept close around the knob: it is a strip of glass over somebody else's
        // picture, so what it is wider than the knob and its collar by is a couple of pixels and no
        // more, at either side and at both ends.
        let beside = (popup.panel.right - popup.panel.left - collar * 2) / 2;
        assert!(
            (1..=4).contains(&beside) && ends <= collar + 8,
            "the panel hugs the knob: {beside} beside it, {ends} past its ends"
        );

        // A point on the groove is the share of the level it is at, and a point past either end is
        // the end it is past.
        assert_eq!(volume_share_at(popup.track.bottom, popup.track), 0.0);
        assert_eq!(volume_share_at(popup.track.top, popup.track), 1.0);
        let middle = (popup.track.top + popup.track.bottom) / 2;
        assert!((volume_share_at(middle, popup.track) - 0.5).abs() < 0.02);
        assert_eq!(volume_share_at(0, popup.track), 1.0);
        assert_eq!(volume_share_at(strip_top + strip_height, popup.track), 0.0);

        // And what is drawn at a level is drawn where it was read from: the knob of 100% is at the
        // top of the groove, the knob of nothing is at the bottom, and the two are inside it.
        assert_eq!(volume_thumb_row(popup.track, 100), popup.track.top);
        assert_eq!(volume_thumb_row(popup.track, 0), popup.track.bottom);
        assert!(volume_thumb_row(popup.track, 50) < volume_thumb_row(popup.track, 25));
    }

    /// A window too narrow for the panel is not a popup drawn off the side of it: it is a panel
    /// that starts at the window's own edge, where the hand that opened it can still reach it.
    #[test]
    fn a_volume_popup_is_kept_inside_a_narrow_window() {
        let popup = volume_popup_layout(30, 100, 30, 96);

        assert!(popup.panel.left >= 0);
        assert!(popup.panel.right > popup.panel.left);
        assert!(popup.track.left >= popup.panel.left);
        assert!(popup.track.bottom > popup.track.top);
    }

    fn scratch(label: &str) -> std::path::PathBuf {
        let root = std::env::var_os("COMMANDCODE_SCRATCHPAD")
            .map(std::path::PathBuf::from)
            .unwrap_or_else(std::env::temp_dir)
            .join("pin-chrome")
            .join(label);
        std::fs::create_dir_all(&root).expect("a fixture directory");
        root
    }

    fn write_png(path: std::path::PathBuf, bgra: &[u8], width: u32, height: u32) {
        let mut rgba = bgra.to_vec();
        for pixel in rgba.as_chunks_mut::<4>().0 {
            pixel.swap(0, 2);
        }
        image::save_buffer(path, &rgba, width, height, image::ExtendedColorType::Rgba8)
            .expect("a written picture");
    }

    fn image_buffer(width: u32, height: u32) -> Vec<u8> {
        vec![0u8; width as usize * height as usize * 4]
    }

    /// One painted button's own width pasted into a strip of several, so that a row of glyphs is
    /// laid out side by side the way the bar lays them out. The source is a button wide and the
    /// destination is the whole strip, so the two are not the same stride.
    fn blit_cell(
        out: &mut [u8],
        width: u32,
        height: u32,
        source: &[u8],
        source_width: u32,
        x: u32,
    ) {
        for row in 0..height as usize {
            for column in 0..source_width as usize {
                let from = (row * source_width as usize + column) * 4;
                let to = (row * width as usize + x as usize + column) * 4;
                if to + 3 < out.len() && from + 3 < source.len() {
                    out[to..to + 4].copy_from_slice(&source[from..from + 4]);
                }
            }
        }
    }

    /// A caption is opaque everywhere, and a tooltip drawn on one of its own surfaces carries
    /// its text across with the coverage that text has.
    ///
    /// The name is written through GDI, and GDI knows nothing of an alpha channel: it leaves the
    /// alpha byte of everything it draws at zero. A caption is copied into the layered surface
    /// a row at a time and composited with `AC_SRC_ALPHA`, so a pixel left at zero is a hole in
    /// the title bar — and a run of text is drawn with an opaque background over the whole of
    /// its box, not just over its glyphs, so the hole is a box rather than a letter.
    ///
    /// This is the one thing about a caption's text a test on its layout cannot see, and the
    /// reason the caption hands its own coverage back after the last of its runs.
    #[test]
    fn a_caption_is_opaque_everywhere_it_is_painted() {
        let palette = ChromePalette {
            background: [250, 250, 250],
            foreground: [30, 30, 30],
            accent: [10, 90, 200],
            dark: false,
        };
        let width = 600u32;
        // A caption is as tall as a pinned window's caption is given at a display's scale, which
        // is the height the surface is really created at (see `pinned_caption_height`).
        let height = 30u32;
        let surface = DibSurface::create(width, height).expect("a surface");

        // With a name on it and without: the coverage is the caption's own, and a caption that
        // only happened to be opaque on a strip with nothing written on it is not opaque.
        for title in ["picture.png", ""] {
            let caption = Caption {
                title,
                maximized: false,
                maximizable: true,
                hovered: Some(CaptionButton::OpenWith),
                pressed: None,
            };
            paint_caption(&surface, &palette, &caption, 96);

            let pixels = unsafe {
                std::slice::from_raw_parts(surface.bits(), width as usize * height as usize * 4)
            };
            let transparent = pixels
                .as_chunks::<4>()
                .0
                .iter()
                .filter(|pixel| pixel[3] != 255)
                .count();
            assert_eq!(
                transparent,
                0,
                "a caption is opaque wherever it is painted, and {transparent} of its pixels were not"
            );
        }
    }

    /// The strip a repaint copies is the strip that was painted, and every change that would
    /// have moved a pixel on it is a change that misses the cache.
    ///
    /// A cached caption that does not repaint is the worst failure this module has available: a
    /// window whose title bar shows the last file's name, or a button that stays washed under a
    /// pointer that has walked off it, with nothing in the window to say so. So each input the
    /// drawing reads is changed on its own and the strip is required to change with it — and
    /// then the *same* strip is asked for twice and required to come back byte for byte, which
    /// is what makes the cache worth having at all.
    #[test]
    fn a_caption_is_repainted_when_anything_it_is_drawn_from_changes() {
        let palette = ChromePalette {
            background: [250, 250, 250],
            foreground: [30, 30, 30],
            accent: [10, 90, 200],
            dark: false,
        };
        let dark = ChromePalette {
            background: [30, 30, 30],
            foreground: [220, 220, 220],
            accent: [10, 140, 240],
            dark: true,
        };
        // A name short enough to be drawn whole on the strip: a caption whose name is cut away
        // entirely is the same picture whichever file it is, so a test on one of those could not
        // tell a stale strip from a changed one.
        let caption = Caption {
            title: "Readme.md",
            maximized: false,
            maximizable: true,
            hovered: Some(CaptionButton::OpenWith),
            pressed: None,
        };
        let surface = DibSurface::create(600, 30).expect("a surface");

        let strip = |palette: &ChromePalette, caption: &Caption, dpi: u32| {
            paint_caption(&surface, palette, caption, dpi);
            surface.pixels()
        };

        let first = strip(&palette, &caption, 96);
        // The same strip again, and byte for byte the one before: a cached caption is the bytes
        // the eye has already been shown, or it is a caption that flickers for nothing.
        assert_eq!(
            strip(&palette, &caption, 96),
            first,
            "the same caption asked for twice is not the strip it was"
        );

        // And every one of these is a different strip. Each is asked for straight after the
        // first one rather than after the one before it, so that only what it changes differs —
        // a cache missing one field is only caught by a change that is *only* that field.
        let changed = [
            (
                "another file",
                &palette,
                &Caption {
                    title: "Cargo.toml",
                    ..caption
                },
                96,
            ),
            (
                "the hovered button left",
                &palette,
                &Caption {
                    hovered: None,
                    ..caption
                },
                96,
            ),
            (
                "the pressed button arrived",
                &palette,
                &Caption {
                    hovered: None,
                    pressed: Some(CaptionButton::Close),
                    ..caption
                },
                96,
            ),
            (
                "a close button is washed red and another one is not",
                &palette,
                &Caption {
                    hovered: Some(CaptionButton::Close),
                    ..caption
                },
                96,
            ),
            (
                "a window with a maximum to go to and one without",
                &palette,
                &Caption {
                    maximizable: false,
                    ..caption
                },
                96,
            ),
            (
                "a maximized window draws a restore pair and an ordinary one does not",
                &palette,
                &Caption {
                    maximized: true,
                    ..caption
                },
                96,
            ),
            ("another display's scale", &palette, &caption, 144),
            ("another theme", &dark, &caption, 96),
        ];
        for (what, other_palette, other_caption, other_dpi) in changed {
            // Straight back to the strip painted above, so the cache holds it and the only thing
            // that differs about the next paint is what this case changes.
            assert_eq!(
                strip(&palette, &caption, 96),
                first,
                "{what}: the base moved"
            );
            assert_ne!(
                strip(other_palette, other_caption, other_dpi),
                first,
                "`{what}` painted the caption that was already there"
            );
        }

        // And a wider window is not a narrow one's strip with room at the end of it: the buttons
        // are against the right edge and the name is cut against the left one, so both of them
        // move with the window.
        let wide = DibSurface::create(900, 30).expect("a surface");
        paint_caption(&wide, &palette, &caption, 96);
        assert_ne!(
            wide.pixels(),
            first,
            "a resized window painted the strip of the width it had"
        );
    }

    /// The pictures this test writes are the design under review: the six marks the caption's
    /// own buttons carry, at both scales, side by side in one strip. A row of marks is judged
    /// as a row — one of them twice the size of the rest, or a third the size, or sitting a
    /// pixel or two off the middle of its own button, is what a strip shows and a test on one
    /// mark at a time cannot.
    #[test]
    fn draws_the_caption_glyphs() {
        let dir = scratch("caption-glyphs");
        // The marks on the bar's own background: the glyphs are ink, so a picture that is not
        // the colour behind them shows nothing, which is what a black buffer and black ink do.
        let bar = [250u8, 250, 250];
        let marks = [
            ("previous", CaptionButton::Previous, false),
            ("next", CaptionButton::Next, false),
            ("open-with", CaptionButton::OpenWith, false),
            ("open-with-list", CaptionButton::OpenWithList, false),
            ("minimize", CaptionButton::Minimize, false),
            ("maximize", CaptionButton::Maximize, false),
            ("restore", CaptionButton::Maximize, true),
            ("close", CaptionButton::Close, false),
        ];

        for (dpi, tag) in [(96u32, "96dpi"), (192u32, "192dpi")] {
            let scale = dpi as f32 / 96.0;
            let side = text_paint::scaled(BUTTON_PIXELS as i32, scale).max(1);
            let height = side as u32;
            let width = (side * marks.len() as i32) as u32;
            let mut out = image_buffer(width, height);
            let mut sizes: Vec<(&str, i32, i32)> = Vec::new();

            for (index, (name, kind, maximized)) in marks.iter().enumerate() {
                let surface = DibSurface::create(side as u32, height).expect("a surface");
                // The bar behind the mark, opaque, so that a column the glyph did not reach
                // reads as the background and not as ink. Filled through the same raw bits
                // every glyph here is drawn into.
                let filled = unsafe {
                    std::slice::from_raw_parts_mut(
                        surface.bits(),
                        4 * side as usize * height as usize,
                    )
                };
                for pixel in filled.as_chunks_mut::<4>().0 {
                    pixel[0] = bar[2];
                    pixel[1] = bar[1];
                    pixel[2] = bar[0];
                    pixel[3] = 255;
                }

                paint_glyph(
                    &surface,
                    CaptionButtonBox {
                        kind: *kind,
                        rect: RECT {
                            left: 0,
                            top: 0,
                            right: side,
                            bottom: side,
                        },
                    },
                    *maximized,
                    [56u8, 58, 66],
                    scale,
                );

                let painted = surface.pixels();
                let ink = |x: usize, y: usize| painted[(y * side as usize + x) * 4] < 128;

                // The mark's own bounds in the surface it was painted into, as the min and
                // max of every row and column that carries any ink at all.
                let mut low = (i32::MAX, i32::MAX);
                let mut high = (i32::MIN, i32::MIN);
                for y in 0..height as usize {
                    for x in 0..side as usize {
                        if !ink(x, y) {
                            continue;
                        }
                        low = (low.0.min(x as i32), low.1.min(y as i32));
                        high = (high.0.max(x as i32), high.1.max(y as i32));
                    }
                }
                assert!(
                    low.0 > i32::MIN && high.0 > i32::MIN,
                    "{name} drew nothing at all at {tag}"
                );

                // The mark sits in the middle of its own button, which is the half of "beauty"
                // a row of glyphs is judged on that no single-glyph test can see: a mark one
                // pixel off its centre is a mark the hand does not aim where the eye is.
                let middle = side / 2;
                let (across, down) = ((low.0 + high.0) / 2 - middle, (low.1 + high.1) / 2 - middle);
                assert!(
                    across.abs() <= 1 && down.abs() <= 1,
                    "{tag} {name} is {across}px across and {down}px down from the middle of its button"
                );

                // And it is the size of the rest of the row, which is the other half. A glyph
                // drawn twice the height of its neighbours is the one the report is about, and
                // a picture is the only place that shows: the chevron once reached a whole span
                // either side of the middle row, twice every other mark's height. The
                // minimize bar is a bar, so it is not held to the height of the rest.
                let (tall, wide) = (high.1 - low.1 + 1, high.0 - low.0 + 1);
                sizes.push((*name, tall, wide));

                println!(
                    "{tag:>6} {name:<9} x {}..={} y {}..={} ({}x{})",
                    low.0,
                    high.0,
                    low.1,
                    high.1,
                    high.0 - low.0 + 1,
                    high.1 - low.1 + 1
                );

                let at = (index as u32) * side as u32;
                blit_cell(&mut out, width, height, &painted, side as u32, at);
            }

            // The walk buttons are the size of the marks they sit beside, which is the whole
            // of the report they were redrawn for: they used to reach a whole span either side
            // of the middle row and come out twice the height of every other mark in the bar.
            //
            // The size they have to match is the glyph square the strip is drawn on — a mark
            // is asked for a span, and a span across and a span down is what it should be. The
            // close cross is not the yardstick: a diagonal stroke laid down a pixel at a time
            // overshoots its own ends by the stroke's width, so the cross is a row or two taller
            // than it is wide at 200% and always has been. Comparing to it would be holding the
            // chevron to the cross's overhang rather than to the size a glyph is asked for.
            let glyph = text_paint::scaled(GLYPH_PIXELS as i32, scale).max(6);
            let square = glyph + 1;
            for (name, tall, wide) in &sizes {
                if *name != "previous" && *name != "next" {
                    continue;
                }
                assert_eq!(
                    (*tall, *wide),
                    (square, square),
                    "{tag} {name} is {tall}x{wide} beside the {square}x{square} a glyph is asked for"
                );
            }

            // And the row is one size the rest of the way: nothing stands much taller than the
            // rest, which is the property a picture shows and a per-glyph test cannot. The
            // minimize bar is not counted: it is a bar, and is one row tall by being a bar.
            let heights = sizes
                .iter()
                .filter(|(name, ..)| *name != "minimize")
                .map(|(_, tall, _)| *tall);
            let tallest = heights.clone().max().unwrap_or_default();
            let smallest = heights.min().unwrap_or_default();
            assert!(
                tallest - smallest <= 2,
                "{tag} the row runs from {smallest} to {tallest} rows tall"
            );

            let path = dir.join(format!("caption-glyphs-{tag}.png"));
            write_png(path.clone(), &out, width, height);
            println!("{}", path.display());
        }
    }

    /// The two walk buttons point opposite ways, each points the way it walks, and each is
    /// drawn the size of every other glyph in the bar.
    ///
    /// A chevron whose arms open on both sides of its middle is a cross, and one whose arms
    /// open away from the end the columns stop at points the wrong way — either of which is a
    /// button that reads as something other than what it does, and neither of which a test on
    /// the boxes alone could see. Arms that reach further than the mark is wide are the third
    /// of the same kind: a row of marks that are not one size is a row of things that are not
    /// one kind of thing, and the walk buttons are the two biggest of the six.
    #[test]
    fn the_two_walk_buttons_are_chevrons_pointing_opposite_ways() {
        // A box with room for the whole glyph and a row to spare: the arms part half a span
        // either side of the middle row, so a box the glyph's own width would cut the open
        // end off and leave a test that passes for a chevron that is really a stub.
        const W: i32 = 24;
        const H: i32 = 24;
        let span = 10;
        let ink = [0u8, 0, 0];

        // The rows drawn in one column. The buffer is filled with a colour first, so that a
        // blank column reads as blank rather than as ink.
        fn drawn(buffer: &[u8], x: i32) -> Vec<i32> {
            (0..H)
                .filter(|y| buffer[((*y * W + x) * 4) as usize] == 0)
                .collect()
        }

        let mut left = vec![255u8; (W * H * 4) as usize];
        draw_chevron(&mut left, W, 12, 12, span, 1.0, ink, false);
        let mut right = vec![255u8; (W * H * 4) as usize];
        draw_chevron(&mut right, W, 12, 12, span, 1.0, ink, true);

        // A chevron has a vertex: the end it points at is a single row, and the arms open
        // away from it to two that part as they go. A cross has two rows at both ends and
        // four through its middle, so the end that is one row is what says a chevron — and
        // the point is at the end the button walks off, which is a different end for each.
        let point_end = 12 - span / 2;
        let open_end = 12 + span / 2;
        for (buffer, point, open, name) in [
            (&left, point_end, open_end, "the back one"),
            (&right, open_end, point_end, "the on one"),
        ] {
            assert_eq!(
                drawn(buffer, point).len(),
                1,
                "{name} points at one row of its end, and not two: {:?}",
                drawn(buffer, point)
            );
            assert_eq!(
                drawn(buffer, open).len(),
                2,
                "{name} opens to two rows at the other end: {:?}",
                drawn(buffer, open)
            );
        }

        // The two are the same shape facing the other way, so what the back one draws in a
        // column is what the on one draws in the column reflected about the middle of the
        // glyph. Two buttons drawn the same way round would be equal column for column
        // instead, and this is what catches that. Only the columns the glyph itself covers
        // are compared: the rest of the box is empty on both sides and reflects out of it.
        for x in 12 - span / 2..=12 + span / 2 {
            assert_eq!(
                drawn(&left, x),
                drawn(&right, 2 * 12 - x),
                "the two chevrons are each other turned about at x={x}"
            );
        }

        // The arms part half a span either side of the middle row, which is what the close
        // cross and the maximize box are drawn at: a chevron taller than the glyph beside it
        // is an arrow in a row of marks, not a mark in a row.
        for buffer in [&left, &right] {
            let ink = |x: i32, y: i32| buffer[((y * W + x) * 4) as usize] == 0;
            let (top, bottom) = (0..H)
                .find(|y| (point_end..=open_end).any(|x| ink(x, *y)))
                .zip(
                    (0..H)
                        .rev()
                        .find(|y| (point_end..=open_end).any(|x| ink(x, *y))),
                )
                .expect("a chevron that was drawn");
            assert_eq!(
                bottom - top + 1,
                span + 1,
                "a chevron is a glyph tall, whatever else it is"
            );
        }
    }

    /// The two steps of a walk are one another's mirror: a bar at the far edge of the mark and a
    /// triangle beside it pointing the way the step goes — ⏮ beside a step back and ⏭ beside a
    /// step on, which is what a hand has been reaching for beside a play button for thirty years.
    ///
    /// The bar is the wall the triangle is pushed off, and it is at a different end for each of
    /// them; so is the triangle's wide end. A mark whose triangle is walked from the wrong end is
    /// a mark whose bar and triangle are the wrong distance apart and, once the walk runs past
    /// the far edge of the button, a mark cut in half at one end and drawn twice at the other —
    /// none of which a test on the boxes alone can see.
    #[test]
    fn the_two_steps_of_a_walk_are_mirrors_of_one_another() {
        const W: i32 = 32;
        const H: i32 = 24;
        let centre = 16;
        let span = 14;
        let ink = [0u8, 0, 0];

        // The rows drawn in one column. The buffer is filled with a colour first, so that a
        // blank column reads as blank rather than as ink.
        fn drawn(buffer: &[u8], x: i32) -> Vec<i32> {
            (0..H)
                .filter(|y| buffer[((*y * W + x) * 4) as usize] == 0)
                .collect()
        }

        let mut on = vec![255u8; (W * H * 4) as usize];
        draw_track_step(&mut on, W, centre, 12, span, ink, true);
        let mut back = vec![255u8; (W * H * 4) as usize];
        draw_track_step(&mut back, W, centre, 12, span, ink, false);

        // The two are the same mark facing the other way, so what the back one draws in a column
        // is what the on one draws in the column reflected about the centreline — which is the line
        // between the two middle columns, a whole mark being an even number of columns wide. Two
        // marks drawn the same way round would be equal column for column instead, and a mark drawn
        // a column too wide would fail only on one of the two ends, and this is what catches that.
        for x in 0..W {
            let mirror = 2 * centre - 1 - x;
            if !(0..W).contains(&mirror)
                || (drawn(&back, x).is_empty() && drawn(&on, mirror).is_empty())
            {
                continue;
            }
            assert_eq!(
                drawn(&back, x),
                drawn(&on, mirror),
                "the two steps are each other turned about at x={x}"
            );
        }

        // And both are inside the button they are drawn in, which is what a walk past the far edge
        // would otherwise have lost: neither mark is a column wider than the other, or wider than
        // the button.
        for (buffer, name) in [(&on, "the on one"), (&back, "the back one")] {
            let mut columns = (0..W).filter(|x| !drawn(buffer, *x).is_empty());
            assert_eq!(
                columns.next(),
                Some(centre - span / 2),
                "{name} begins at its own near edge"
            );
            assert_eq!(
                columns.next_back(),
                Some(centre + span / 2 - 1),
                "{name} ends at its own far one"
            );
        }
    }

    #[test]
    fn the_parts_of_a_transport_bar_are_hit_where_they_are_drawn() {
        let layout = transport_layout(800, 30, 96, true);

        let play = (layout.play.left + layout.play.right) / 2;
        assert_eq!(
            transport_part_at(play, 15, 800, 30, 96, true),
            Some(TransportPart::Play)
        );

        let bar = (layout.bar.left + layout.bar.right) / 2;
        assert_eq!(
            transport_part_at(bar, 15, 800, 30, 96, true),
            Some(TransportPart::Seek)
        );

        // Nothing is hit where nothing is drawn.
        assert_eq!(transport_part_at(2, 2, 800, 30, 96, true), None);
    }

    /// A bar whose player cannot be told anything is a read-out: FFmpeg's player reports no
    /// position, takes no pause, and can only be "seeked" by being ended and begun again at a
    /// second — so its bar carries no button and nothing to drag, and the two clocks begin where
    /// the strip does rather than after a control that is not there.
    #[test]
    fn a_bar_with_no_controls_has_none_to_press() {
        let dead = transport_layout(800, 30, 96, false);
        let live = transport_layout(800, 30, 96, true);

        assert_eq!(dead.elapsed.left, 10, "the clock begins at the padding");
        assert_eq!(
            live.elapsed.left,
            dead.elapsed.left + (live.play.right - live.play.left) + 5,
            "a bar with a button begins after it"
        );

        // The track is longer by the room the button does not take, so a playhead drawn against
        // it is drawn against the same file rather than against a bar with a hole in it.
        assert!(dead.bar.left < live.bar.left);
        assert!(
            dead.bar.right - dead.bar.left > live.bar.right - live.bar.left,
            "the space a button does not take is the track's"
        );

        // And neither the button's place nor the track answers a press.
        let play = (live.play.left + live.play.right) / 2;
        let bar = (dead.bar.left + dead.bar.right) / 2;
        assert_eq!(transport_part_at(play, 15, 800, 30, 96, false), None);
        assert_eq!(transport_part_at(bar, 15, 800, 30, 96, false), None);
        assert_eq!(
            transport_part_at(play, 15, 800, 30, 96, true),
            Some(TransportPart::Play)
        );
    }
}

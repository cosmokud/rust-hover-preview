//! The bar a pinned preview plays from: where its parts are, what a press on one of them means,
//! and the glyphs this app's own transport is drawn in.
//!
//! It is a strip of the window rather than part of the media (see the module documentation): the
//! button that pauses and resumes, the bar with the part of the file that has been played filled
//! in, the two clocks beside them, and the volume button at the far end - which is this app's own
//! control rather than the player's, so it is answered on a bar whose player cannot be told
//! anything at all, and so the panel it opens is placed and drawn here too.
//!
//! The marks are named rather than left to each caller to draw (`ControlGlyph`), because a sound's
//! card draws the same three in a row of its own (see `paint_card_control`), and two sets of
//! numbers for one control this app has once is a button that moves between two windows.

use super::primitives::{
    fill_box, fill_disc, fill_polygon, fill_round_rect, paint_time, put, stroke_arc,
    stroke_round_rect, stroke_segment, surface_pixels, ChromePalette, GLYPH_STROKE_PIXELS,
};
use crate::text::text_paint::{self, DibSurface};
use windows::Win32::Foundation::RECT;

/// The room a volume popup takes at a display's scale: the panel that floats over the media
/// above the bar, and the groove the level is drawn in inside it.
///
/// The panel is a strip of glass over somebody else's picture, so it is kept no wider than it has
/// to be: the thumb and its collar, and a couple of pixels of panel either side of them.
pub(super) const VOLUME_PANEL_WIDTH: i32 = 24;
const VOLUME_PANEL_HEIGHT: i32 = 98;
/// How far the groove is held off each end of the panel, which is the room the thumb takes at
/// either end of it: a thumb that was clipped by the panel at 100% would be a level drawn short,
/// and anything beyond that room is panel being drawn over a picture it is covering.
const VOLUME_PANEL_INSET: i32 = 14;
/// How far the panel floats above the bar it belongs to.
pub(super) const VOLUME_PANEL_GAP: i32 = 8;
const VOLUME_PANEL_RADIUS: f32 = 9.0;
const VOLUME_TRACK_WIDTH: i32 = 6;
/// The radius of the button's own wash: the button is a chip rather than the square the strip's
/// room for it would otherwise draw, since it is a control a hand comes back to.
const VOLUME_BUTTON_RADIUS: i32 = 7;
/// The radius of the knob on the groove, which is the bar's own thumb made rounder: it is held
/// rather than aimed at, so it is drawn as something a finger fits.
pub(super) const VOLUME_THUMB_RADIUS: f32 = 8.0;
/// How much wider the collar around the knob is than the knob itself: the ring of panel colour
/// that keeps a knob drawn over the filled part of the groove reading as a knob.
pub(super) const VOLUME_COLLAR_PIXELS: f32 = 1.5;

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
/// `TransportState::interactive`, which is now `false` for no kind this app has — it is kept
/// because the layout is the only thing that reads it).
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
/// floating over the media above the bar, centered on the button and kept inside
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

    // The panel sits centered on the button's own middle — a hand aims at the
    // button, and the level answers from the middle of it — rather than hung off
    // the button's right edge, and kept inside the window: a window too narrow to
    // hold the panel centered is a panel held at the window's own edge instead,
    // where the hand that opened it can still reach it.
    let center = (button.left + button.right) / 2;
    let left = (center - panel_width / 2).clamp(0, width.saturating_sub(panel_width).max(0));
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
///
/// The layout is built at the strip's own height, which is the whole of what makes this the same
/// bar the thumb is drawn on. The play button is sized from the height it is drawn at, and the
/// track is laid out between the button and the clocks that button leaves behind, so a layout
/// built at any other height is a bar of a different length somewhere else on the strip: the
/// share it reads off a press is then measured across a track that is not the one drawn, and the
/// thumb lands short of the hand at one end of the bar and past it at the other, while the middle
/// stays where it is by the accident that the two tracks share a centre (see
/// `a_press_puts_the_thumb_under_the_hand_wherever_it_landed`).
pub(crate) fn transport_share_at(
    x: i32,
    width: i32,
    height: i32,
    dpi: u32,
    interactive: bool,
) -> f64 {
    let layout = transport_layout(width, height, dpi, interactive);
    let span = (layout.bar.right - layout.bar.left).max(1) as f64;

    ((x - layout.bar.left) as f64 / span).clamp(0.0, 1.0)
}

/// What a transport bar is drawn from.
pub(crate) struct TransportState {
    /// Whether the player behind the bar can be told anything at all.
    ///
    /// Every control on the bar is a question the player answers, and the two players here answer
    /// them for different reasons rather than to different depths. The media engine is *asked* —
    /// a pause is a pause and a position is reported — while FFmpeg's player is *posted to*: its
    /// window takes a key pressed on it, which is how a hold and a track change both reach it, and
    /// a drag of the bar still ends and begins a player because that is the only way to say
    /// "this second" rather than one of the ten-second steps its own keys offer (see
    /// `preview_window::seek_pinned_playback`).
    ///
    /// So a bar is drawn for a kind or not at all, and this says whether its parts are drawn —
    /// because the one kind with a bar this app cannot drive would be a bar with a button on it
    /// that does nothing, and a button that does nothing is a promise the app cannot keep (see
    /// `pin_transport_live`).
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
pub(super) fn draw_track_step(
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

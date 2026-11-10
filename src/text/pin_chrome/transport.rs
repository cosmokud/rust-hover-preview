//! The bar a pinned preview plays from: where its parts are, what a press on one of them means,
//! and the glyphs this app's own transport is drawn in.
//!
//! It is a strip of the window rather than part of the media (see the module documentation): the
//! button that pauses and resumes, the bar with the part of the file that has been played filled
//! in, the two clocks beside them, and one control at the far end - which is this app's own rather
//! than the player's, so it is answered on a bar whose player cannot be told anything at all, and
//! so the panel it opens is placed and drawn here too.
//!
//! **The far end is one control and not two.** A film's is its level, which this app keeps because
//! FFmpeg's player cannot be told one; a picture that plays has no level at all - a file of frames
//! has no sound to hand a player and no player to hand it to - so its own control there is the pace
//! those frames move at, and a count of the frames a hand on its track is standing at is written
//! beside the clock.
//!
//! The marks are named rather than left to each caller to draw (`ControlGlyph`), because a sound's
//! card draws the same three in a row of its own (see `paint_card_control`), and two sets of
//! numbers for one control this app has once is a button that moves between two windows.

use super::primitives::{
    caption_style, fill_box, fill_disc, fill_polygon, fill_round_rect, fill_triangle, measure_text,
    paint_time, put, stroke_arc, stroke_round_rect, stroke_segment, surface_pixels, ChromePalette,
    GLYPH_STROKE_PIXELS,
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
/// The room a pace's panel takes at a display's scale: the volume panel widened by the figures it
/// has to hold, because what a pace's panel says is a number written across the top of it and a
/// number in a level's own narrow panel is a figure read out of a slot (see `paint_fps_popup_text`).
const FPS_PANEL_WIDTH: i32 = 46;
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
    /// player's - and one a bar of a picture that plays does not carry, since a file of frames has
    /// no sound for a level to be a level of (see `transport_layout`).
    Volume,
    /// The button that opens the FPS popup, which is the pin's own control the same way the
    /// volume button is — and one only a picture that plays carries, since a film's pace is its
    /// player's and a file's own is what moves here (see `PinFps`).
    Fps,
}

/// Where the parts of a transport bar are, in the bar's own coordinates: the button, the bar
/// itself, the two labels beside them, and the volume button at the far end.
pub(crate) struct TransportLayout {
    pub(crate) play: RECT,
    pub(crate) bar: RECT,
    /// The clock before the bar, held against it: what is written beside the track is that far
    /// from it and no further, so a strip does not open a gap between a clock and the bar it
    /// belongs to.
    pub(crate) elapsed: RECT,
    pub(crate) total: RECT,
    /// The level's button at the far end of the bar, and a box of no width on a bar that has none
    /// (see `TransportPart::Volume`).
    pub(crate) volume: RECT,
    /// The FPS button at the far end of a bar a picture that plays carries, and a box of no width
    /// on every other bar: the same square as the level's, because it is the same button a hand
    /// reaches for, saying the one thing its own slider changes.
    pub(crate) fps: RECT,
    /// The room left of the clock for the count of frames a hand on the track is standing at, and a
    /// box of no width on a bar that writes no count.
    pub(crate) frames: RECT,
}

/// The room held between the parts of a transport bar at a display's scale: one part is never
/// written through the next, and a strip whose parts touched would be one run of ink.
pub(super) fn strip_gap(scale: f32) -> i32 {
    (text_paint::scaled(10, scale) / 2).max(2)
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
    animated: bool,
) -> TransportLayout {
    let scale = dpi as f32 / 96.0;
    let padding = text_paint::scaled(10, scale);
    let play_side = (height - text_paint::scaled(8, scale)).clamp(8, height.max(8));
    let label = text_paint::scaled(52, scale);
    let gap = strip_gap(scale);
    // The room a picture that plays keeps left of its clock for the count of frames under the hand.
    // It is held open whether or not a hand is on the track, because the count is written beside a
    // bar that must not move while it is dragged - and given up where the strip is too narrow to
    // hold both it and a track worth dragging, since a strip a tenth of a window wide is a bar
    // nobody can aim at.
    let frames_room = if animated && width >= text_paint::scaled(360, scale) {
        text_paint::scaled(74, scale)
    } else {
        0
    };

    let play = RECT {
        left: padding,
        top: (height - play_side) / 2,
        right: padding + play_side,
        bottom: (height + play_side) / 2,
    };
    // The far end of the bar, against the right edge the way it is on every player that has one,
    // and keeping its box whether or not the bar it sits on has controls, since it is this app's
    // own control rather than the player's (see `TransportPart::Volume`).
    let last = RECT {
        left: (width - padding - play_side).max(padding),
        top: (height - play_side) / 2,
        right: (width - padding).max(padding),
        bottom: (height + play_side) / 2,
    };
    let nothing = RECT {
        left: last.left,
        top: last.top,
        right: last.left,
        bottom: last.bottom,
    };
    // One control at the far end: the level on a bar with a player behind it, and the pace on a
    // picture that plays, which has no level to show (see `TransportPart::Fps`).
    let (volume, fps) = if animated {
        (nothing, last)
    } else {
        (last, nothing)
    };
    let content_left = if interactive {
        play.right + gap
    } else {
        padding
    };
    let frames = RECT {
        left: content_left,
        top: 0,
        right: content_left + frames_room,
        bottom: height,
    };
    let elapsed = RECT {
        left: frames.right,
        top: 0,
        right: frames.right + label,
        bottom: height,
    };
    let controls_left = if animated { fps.left } else { volume.left };
    let total = RECT {
        left: (controls_left - gap - label).max(elapsed.right),
        top: 0,
        right: (controls_left - gap).max(elapsed.right),
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
        fps,
        frames,
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
    animated: bool,
) -> Option<TransportPart> {
    let layout = transport_layout(width, height, dpi, interactive, animated);
    let inside = |rect: RECT| x >= rect.left && x < rect.right && y >= rect.top && y < rect.bottom;

    if !animated && inside(layout.volume) {
        return Some(TransportPart::Volume);
    }
    // The FPS button is this app's own control too, and it is answered on a bar whose
    // player cannot be told anything at all — a picture that plays has no player (see
    // `TransportPart::Fps`).
    if animated && inside(layout.fps) {
        return Some(TransportPart::Fps);
    }

    if !interactive && !animated {
        return None;
    }

    if interactive && inside(layout.play) {
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
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
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
#[cfg(test)]
pub(crate) fn volume_popup_layout(
    width: i32,
    strip_top: i32,
    strip_height: i32,
    dpi: u32,
) -> VolumePopup {
    // The button is the transport bar's own, and the bar begins `strip_top` rows down the
    // window, so the box it opens from is the strip's box moved down the window — which is what
    // makes the two answers one answer (see `volume_popup_from_button`).
    let in_strip = transport_layout(width, strip_height, dpi, true, false).volume;
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
    popup_from_button(button, width, dpi, VOLUME_PANEL_WIDTH)
}

/// The panel a bar's pace button opens, hung from that button's own box: the level's panel in
/// every way but the room across it, since what a pace's panel holds is a number written over its
/// groove and a number in a level's own narrow panel would be a figure read out of a slot.
///
/// It is placed from the button the same way and for the same reason - the panel a press is
/// answered against is the panel that is drawn (see `paint_fps_popup`).
pub(crate) fn fps_popup_from_button(button: RECT, width: i32, dpi: u32) -> VolumePopup {
    popup_from_button(button, width, dpi, FPS_PANEL_WIDTH)
}

fn popup_from_button(button: RECT, width: i32, dpi: u32, panel_width_in: i32) -> VolumePopup {
    let scale = dpi as f32 / 96.0;

    let panel_width = text_paint::scaled(panel_width_in, scale).max(8);
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

/// The footprint a volume popup paints in the window: its panel, the shadow held under it, and a
/// pixel of margin around the edge the panel's antialiasing can land in — clamped to the window,
/// because a popup at its edge has nothing to paint outside it.
///
/// It is what a strip of the popup's own is placed at, so it is the same shadow the painter draws
/// (see `paint_volume_popup`) rather than a second set of numbers for one panel.
pub(crate) fn volume_popup_bounds(popup: &VolumePopup, width: i32, height: i32) -> RECT {
    let panel_height = (popup.panel.bottom - popup.panel.top).max(1) as f32;
    let scale = panel_height / VOLUME_PANEL_HEIGHT as f32;
    let shadow = (2.0 * scale).round().max(1.0) as i32;
    let margin = (1.0 * scale).round().max(1.0) as i32;

    RECT {
        left: (popup.panel.left - margin).clamp(0, width),
        top: (popup.panel.top - margin).clamp(0, height),
        right: (popup.panel.right + margin).clamp(0, width),
        bottom: (popup.panel.bottom + shadow + margin).clamp(0, height),
    }
}

/// Where along a volume popup's groove a point is, as a share of the level: the bottom of the
/// groove is nothing and the top of it is everything, which is the way a level is read.
pub(crate) fn volume_share_at(y: i32, track: RECT) -> f64 {
    let span = (track.bottom - track.top).max(1) as f64;
    ((track.bottom - y) as f64 / span).clamp(0.0, 1.0)
}

/// The pace a point on a popup's groove names, between `1` at the bottom and `max` at the top.
///
/// A pace has no zero: the bottom of the groove is one frame a second, which is a picture a
/// hand can read a frame at a time, and what the slider says at the top is the file's own
/// ceiling rather than a second one of its own (`animated_fps_max`).
pub(crate) fn fps_at(y: i32, track: RECT, max: u32) -> u32 {
    let span = (track.bottom - track.top).max(1) as f64;
    let share = ((track.bottom - y) as f64 / span).clamp(0.0, 1.0);
    let top = max.max(1) as f64;
    ((1.0 + share * (top - 1.0)).round() as u32).clamp(1, max.max(1))
}

/// The row a pace is drawn at on a popup's groove, which is the level's own row for the share it
/// stands at: one groove, read bottom upwards, in whole frames a second.
pub(crate) fn fps_thumb_row(track: RECT, fps: u32, max: u32) -> i32 {
    let top = max.max(1) as f64;
    let share = ((fps.clamp(1, max.max(1)) as f64 - 1.0) / (top - 1.0)).clamp(0.0, 1.0);
    let span = (track.bottom - track.top).max(1) as f64;
    track.bottom - (span * share).round() as i32
}

/// The row a level is drawn at on a popup's groove.
pub(crate) fn volume_thumb_row(track: RECT, volume: u32) -> i32 {
    let share = volume.min(100) as f64 / 100.0;
    let span = (track.bottom - track.top).max(1) as f64;
    track.bottom - (span * share).round() as i32
}

/// The room a track has to name a place in the file: its own columns less the one the last of them
/// already is.
///
/// A bar of `n` columns spans `n - 1` steps between its first and its last, and a file read as `n`
/// is a press on the last column that lands a frame short of the end and a thumb drawn one column
/// past the hand that put it there — so the track's own length is measured from its first column to
/// its last rather than across the room *between* the two ends (see
/// `a_press_puts_the_thumb_under_the_hand_wherever_it_landed`).
fn bar_span(bar: RECT) -> i32 {
    (bar.right - bar.left - 1).max(1)
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
    animated: bool,
) -> f64 {
    let layout = transport_layout(width, height, dpi, interactive, animated);
    // The track's own columns, of which the last is the end of the file: a share read across the
    // room *between* the two ends leaves the last frame of a file under a strip's rightmost pixel
    // out of reach, and a press there is the one place a hand means "the end" (see `bar_span`).
    let span = f64::from(bar_span(layout.bar));

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
    /// Whether the bar belongs to a picture that plays rather than a film: its seek is the
    /// file's frames, and the FPS button at its end is the pace those frames play at.
    pub(crate) animated: bool,
    /// Whether the FPS popup is open — a button whose own slider is up is a button held down, the
    /// way an open volume popup is.
    pub(crate) fps_open: bool,
    /// The frame of the file a hand on the track is standing at, and how many the file plays: a
    /// drag of a picture that plays is a walk through its frames rather than through seconds of a
    /// player, and the count of where the hand is written beside the clock says which frame the
    /// picture is about to be left on (see `pin_frame_share`).
    pub(crate) frames: Option<(usize, usize)>,
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
    let gap = strip_gap(scale);
    let layout = transport_layout(width, height, dpi, state.interactive, state.animated);

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
    // `TransportPart::Volume`). A picture that plays has no level to keep and no button to keep it
    // on, so its far end carries the pace's own instead.
    if !state.animated {
        paint_volume_button(surface, palette, state, layout.volume, scale);
    }

    // And the FPS button, where the bar belongs to a picture that plays: the pace is this
    // app's own control for the same reason the level is (see `TransportPart::Fps`).
    if state.animated {
        paint_fps_button(surface, palette, state, layout.fps, scale);
    }

    // The bar: a track, what has been played filled in, and a thumb at the playhead.
    let filled = match (state.position, state.duration) {
        (Some(position), Some(duration)) if duration > 0.0 => {
            position.clamp(0.0, duration) / duration
        }
        _ => 0.0,
    };
    let thumb = layout.bar.left + (f64::from(bar_span(layout.bar)) * filled).round() as i32;

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

    // The clock before the bar, written against it rather than laid out from the strip's own
    // padding: a clock that starts where the strip starts is a clock a bar's own length away from
    // the strip's, and the gap it opens is a gap in the middle of one read-out. And where a hand
    // is on the track of a picture that plays, the frame it is standing at is written in the room
    // held open left of the clock, so that the two are one read-out beside the track - the count
    // has a room of its own whether or not it is written, which is why the clock does not move
    // under a hand that is dragging the bar it belongs to (see `paint_frame_count`).
    if let Some((frame, total)) = state.frames {
        paint_frame_count(
            surface,
            palette,
            (frame, total),
            RECT {
                left: layout.frames.left,
                top: 0,
                right: (layout.frames.right - gap / 2).max(layout.frames.left),
                bottom: height,
            },
            scale,
        );
    }
    paint_time(
        surface,
        palette,
        state.position,
        layout.elapsed,
        scale,
        true,
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
pub(super) fn draw_pause_glyph(
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

/// The triangle a play/pause button shows while the thing behind it is stopped.
pub(super) fn draw_play_glyph(
    buffer: &mut [u8],
    width: i32,
    center_x: i32,
    center_y: i32,
    size: i32,
    ink: [u8; 3],
) {
    // A triangle pointing right: its base down the left edge, over every row the mark is tall, and
    // its point a column past the last one it is drawn in — the mark's own size, drawn as the slope
    // from the base to the point rather than as the flight of steps a walk of whole columns makes
    // of it (see `fill_triangle`).
    let base = (center_x - size / 3) as f32;
    let half = (size / 2) as f32;
    fill_triangle(
        buffer,
        width,
        [
            (base, center_y as f32 - half),
            (base, center_y as f32 + half + 1.0),
            (base + size as f32 + 1.0, center_y as f32 + 0.5),
        ],
        ink,
        1.0,
    );
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

    // The triangle, drawn by the three lines that make it rather than walked as a scan of
    // half-widths: what a scan cannot draw is a slanted edge, and this mark is nothing else (see
    // `fill_triangle`). The far edge is the last *column* of the mark on both sides, for the reason
    // the mark's own width is counted in columns: the two ends it has are a pixel either side of the
    // centreline rather than at it, and a triangle measured from the centre instead comes out a
    // column too wide on one of the two. The base is the edge the far end of the mark is, and the
    // point is a column past it.
    let far = if forward {
        center_x - half
    } else {
        center_x + half - 1
    };
    let (base, point) = if forward {
        (far as f32, (far + head + 1) as f32)
    } else {
        ((far + 1) as f32, (far - head) as f32)
    };
    fill_triangle(
        buffer,
        width,
        [
            (base, center_y as f32 - half as f32),
            (base, center_y as f32 + half as f32 + 1.0),
            (point, center_y as f32 + 0.5),
        ],
        color,
        1.0,
    );
}

/// Write the frame a hand on the track is standing at against the file's own count of them —
/// `[212/1021]` — in the room the clock leaves, answering the width it took so the clock is held
/// off it rather than written through it.
///
/// It is the count a drag is read by: a picture that plays is moved a frame at a time, and a track
/// of a few hundred columns over a file of a few thousand frames is a place in the file that a
/// second of the run does not name. Nothing is written where the room it is given cannot hold it,
/// which is a strip too narrow to have kept any room at all (see `transport_layout`).
fn paint_frame_count(
    surface: &DibSurface,
    palette: &ChromePalette,
    (frame, total): (usize, usize),
    rect: RECT,
    scale: f32,
) {
    if rect.right <= rect.left {
        return;
    }

    let text = format!("[{frame}/{total}]");
    let style = caption_style(palette.number);
    let measured = measure_text(surface, &style, &text, scale);
    if measured <= 0 || measured >= rect.right - rect.left {
        return;
    }

    let cell = text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[text_paint::BODY_LEVEL as usize],
        scale,
    );
    let top = ((surface.height as i32 - cell) / 2).max(0);
    let box_ = RECT {
        left: rect.right - measured,
        top,
        right: rect.right,
        bottom: (top + cell).min(surface.height as i32),
    };
    let mut painter = text_paint::RunPainter::new(surface, scale);
    painter.draw(
        &text,
        box_.left,
        box_,
        &style,
        palette.number,
        palette.background,
    );
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

/// Paint the FPS button: the three letters, drawn as the word they are rather than as a mark a
/// hand has to learn first.
///
/// A picture that plays has no player behind it to report a pace and no speaker to point at,
/// and what the button opens is a slider the file's own pace is read onto — so the button says
/// the one thing the slider changes, the way the volume button says the one thing its own
/// slider changes.
fn paint_fps_button(
    surface: &DibSurface,
    palette: &ChromePalette,
    state: &TransportState,
    rect: RECT,
    scale: f32,
) {
    let pressed = state.pressed == Some(TransportPart::Fps);
    let buffer = unsafe {
        std::slice::from_raw_parts_mut(
            surface_pixels(surface),
            surface.width as usize * surface.height as usize * 4,
        )
    };
    let width = surface.width as i32;

    // The chip the volume button lights as, from the same radius: a bar with two controls on
    // it whose washes were two shapes would be two kinds of button on one strip.
    if state.hovered == Some(TransportPart::Fps) || pressed || state.fps_open {
        let radius = text_paint::scaled(VOLUME_BUTTON_RADIUS, scale) as f32;
        let wash = if pressed || state.fps_open {
            palette.hover(0.18)
        } else {
            palette.hover(0.10)
        };
        fill_round_rect(buffer, width, rect, radius, wash, wash, 1.0);
    }

    // The word itself, centred in the button's own box and clipped to it: the box is the
    // square the volume button takes, and three letters at the strip's own text size are what
    // fits inside one of those.
    let style = caption_style(palette.foreground);
    let cell = text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[text_paint::BODY_LEVEL as usize],
        scale,
    )
    .min((rect.bottom - rect.top).max(1));
    let measured = measure_text(surface, &style, "FPS", scale);
    let left = (rect.left + ((rect.right - rect.left - measured) / 2)).max(rect.left);
    let top = (rect.top + ((rect.bottom - rect.top - cell) / 2)).max(rect.top);
    let mut painter = text_paint::RunPainter::new(surface, scale);
    painter.draw(
        "FPS",
        left,
        RECT {
            left: rect.left,
            top,
            right: rect.right,
            bottom: (top + cell).min(rect.bottom),
        },
        &style,
        palette.foreground,
        palette.background,
    );
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
/// Everything here is written premultiplied (see the module documentation), against a page that has
/// been filled with the card's own colour first and is read back opaque: `DibSurface::pixels`
/// forces the alpha byte of every pixel it hands back to 255 before the card is shown, which is the
/// byte a pixel that is a blend of two opaque colours already carries. So a glyph whose edge runs
/// across the pixel grid is drawn in the shares of it the shape covers — the way the two window
/// buttons beside it are (`paint_window_button`) — rather than walked in whole columns.
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
pub(super) fn size_in(scale: f32, side: i32, nominal: i32) -> i32 {
    text_paint::scaled(nominal, scale).clamp(6, side)
}

/// Paint the popup a volume button opens: a panel floating over the media above the bar, the
/// groove in it, the level the groove has been taken to, and the knob that is dragged.
///
/// It is drawn into the window's own surface and into a strip of its own: the panel floats over
/// this app's pixels, and over a hosted film — a child of the pin's window, which composites
/// over its parent's layered surface — only a strip of the popup's own stands above the picture
/// (see `pin_chrome_strips` and `volume_popup_bounds`).
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

/// Paint the popup the FPS button opens: the panel and groove the volume panel is drawn with, a
/// knob at the frame a second the file is being held at, and that number written over the groove
/// where a hand set it.
///
/// The number is drawn through GDI, which needs a device context of its own — so what is painted
/// here is the panel, its groove and its knob, and the figure is written over it by the caller
/// that has a surface to write on (see `paint_fps_popup_text`).
pub(crate) fn paint_fps_popup(
    buffer: &mut [u8],
    width: i32,
    palette: &ChromePalette,
    popup: &VolumePopup,
    fps: u32,
    max: u32,
    held: bool,
) {
    let panel_height = (popup.panel.bottom - popup.panel.top).max(1) as f32;
    let scale = panel_height / VOLUME_PANEL_HEIGHT as f32;
    let radius = (VOLUME_PANEL_RADIUS * scale).max(1.0);
    let thumb = (VOLUME_THUMB_RADIUS * scale).max(2.0);

    // The shadow and the panel, exactly as the volume panel's own: one panel, drawn twice for
    // one bar's two controls - and flat rather than shaded from the top, because the figures the
    // pace is read by are written across the top of it in a colour of their own, with no panel of
    // their own behind them (see `paint_fps_popup_text`).
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
    let face = fps_panel_colour(palette);
    fill_round_rect(buffer, width, popup.panel, radius, face, face, 0.97);
    stroke_round_rect(
        buffer,
        width,
        popup.panel,
        radius,
        (1.0 * scale).round().max(1.0),
        palette.hover(0.32),
        1.0,
    );

    // The groove, and the stretch of it the pace covers: filled from the knob down, the way a
    // level is, because what is below the knob is the pace from one frame a second up to it.
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

    let row = fps_thumb_row(popup.track, fps, max);
    if fps > 1 {
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

/// The colour a pace's panel is: the theme's page moved off the caption's the way the tool tip is,
/// one colour over the whole of the panel rather than a shade from its top to its bottom.
///
/// It is one colour because the figures are written across the top of it with no panel of their
/// own under them: a run whose background is the panel's own colour is a run that is only ink, and
/// a run across a gradient is a patch of some shade of the panel with a number in it (see
/// `paint_fps_popup_text`).
fn fps_panel_colour(palette: &ChromePalette) -> [u8; 3] {
    palette.hover(0.06)
}

/// Write the pace a popup is holding over its own groove's top, in the strip's text: the number a
/// hand set, which is the whole of what the panel is for.
///
/// **It is written in the theme's own blue and with nothing behind it.** The number is the one
/// figure on this strip that is a setting rather than a state of the file, and it is read against
/// the panel the way a function is read against a page: One Dark Pro's neon blue, Atom One Light's
/// strong one, and whatever blue a theme of the user's own names (see `ChromePalette::blue`). Its
/// background is the panel's own flat colour, so that where the run is not a figure it is exactly
/// what the panel already is, and the alpha byte GDI leaves behind is handed back the panel's own:
/// a run sealed whole over a panel held a shade off whole would be a patch of the panel cut out of
/// it.
///
/// The panel and its knob are painted as pixels, and the figure is written through GDI —
/// `RunPainter` needs a device context, which the surface a popup is painted on has and the raw
/// bytes its panel is filled in do not (see `paint_fps_popup`).
pub(crate) fn paint_fps_popup_text(
    surface: &DibSurface,
    palette: &ChromePalette,
    popup: &VolumePopup,
    fps: u32,
    dpi: u32,
) {
    let rect = fps_popup_text_rect(popup, dpi);
    if rect.right <= rect.left || rect.bottom <= rect.top {
        return;
    }

    let scale = dpi as f32 / 96.0;
    let style = caption_style(palette.blue);
    let text = fps.to_string();
    let measured = measure_text(surface, &style, &text, scale);
    let left = (rect.left + ((rect.right - rect.left - measured) / 2)).max(rect.left);
    let box_ = RECT {
        right: (left + measured).min(rect.right),
        ..rect
    };

    // The panel's own alpha byte under the figures, put back after GDI has drawn through it.
    let stride = surface.width as usize;
    let alpha_under: Vec<u8> = {
        let pixels = unsafe {
            std::slice::from_raw_parts(
                surface_pixels(surface),
                stride * surface.height as usize * 4,
            )
        };
        let mut alphas = Vec::with_capacity(
            ((box_.right - box_.left).max(0) * (box_.bottom - box_.top).max(0)) as usize,
        );
        for y in box_.top.max(0)..box_.bottom.min(surface.height as i32) {
            for x in box_.left.max(0)..box_.right.min(surface.width as i32) {
                alphas.push(pixels[(y as usize * stride + x as usize) * 4 + 3]);
            }
        }
        alphas
    };

    let mut painter = text_paint::RunPainter::new(surface, scale);
    painter.draw(
        &text,
        left,
        box_,
        &style,
        palette.blue,
        fps_panel_colour(palette),
    );
    drop(painter);

    let pixels = unsafe {
        std::slice::from_raw_parts_mut(
            surface_pixels(surface),
            stride * surface.height as usize * 4,
        )
    };
    let mut at = 0;
    for y in box_.top.max(0)..box_.bottom.min(surface.height as i32) {
        for x in box_.left.max(0)..box_.right.min(surface.width as i32) {
            if let Some(alpha) = alpha_under.get(at) {
                pixels[(y as usize * stride + x as usize) * 4 + 3] = *alpha;
            }
            at += 1;
        }
    }
}

/// Where the pace a popup is holding is written, given the popup and the room one line of the
/// strip's own text takes at that scale: the row above the groove's top, which is the one room in
/// the panel the thumb never reaches.
pub(crate) fn fps_popup_text_rect(popup: &VolumePopup, dpi: u32) -> RECT {
    let scale = dpi as f32 / 96.0;
    let cell = text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[text_paint::BODY_LEVEL as usize],
        scale,
    )
    .max(1);
    let top = (popup.panel.top + 1).max(0);

    RECT {
        left: popup.panel.left + 1,
        top,
        right: popup.panel.right - 1,
        bottom: (top + cell).min(popup.track.top.max(top + cell)),
    }
}

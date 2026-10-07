//! What is drawn while a file is still being read: the spinner, the arc it turns, and the
//! frames that arc is painted into.

use super::*;

/// Turn of a pin spinner's arc, kept at the old animation floor: a wait
/// indicator needs no more than 30 Hz, and doubling its repaints would buy
/// nothing (see `PinLoad::due` and `PinWait::due`).
pub(super) const SPINNER_TURN_MS: u32 = 33;
/// Render a single frame of the loading spinner animation (BGRA pixels).
///
/// The frame is the arc and nothing else: the box it is drawn in is transparent,
/// so what a hover that is waiting shows is a spinner rather than a square of its
/// own. The arc is white, with a soft dark halo drawn under it, because a preview
/// goes over whatever Explorer happens to be drawing — the halo is what keeps the
/// spinner visible over a light background, and the white arc is what keeps it
/// visible over a dark one.
pub(super) fn render_loading_frame(width: u32, height: u32, angle: f32) -> Vec<u8> {
    let total_pixels = (width as usize) * (height as usize);
    let mut pixels = vec![0u8; total_pixels * 4];

    let cx = width as f32 / 2.0;
    let cy = height as f32 / 2.0;

    // Spinner proportional to window size, clamped for aesthetics
    let radius = (width.min(height) as f32 * 0.08).clamp(10.0, 32.0);
    let thickness = (radius * 0.32).clamp(2.5, 7.0);
    // How far the halo reaches past the arc, and how much of it there is where it
    // meets the arc's own edge. Both are drawn from the arc's size, so a large
    // spinner gets a halo in proportion.
    let halo_width = thickness;
    let halo_opacity = 0.85;
    // What the halo is made of: dark enough to read over a light file list, and
    // light enough to read as a shadow rather than a second arc.
    let halo_shade = 12.0;

    let two_pi = std::f32::consts::PI * 2.0;
    let arc_length = std::f32::consts::PI * 1.5; // 270-degree arc

    // Only iterate over the bounding box of the spinner's own ring and halo
    let reach = radius + thickness + halo_width + 1.0;
    let min_x = ((cx - reach).max(0.0)) as u32;
    let max_x = ((cx + reach).min(width as f32 - 1.0)) as u32;
    let min_y = ((cy - reach).max(0.0)) as u32;
    let max_y = ((cy + reach).min(height as f32 - 1.0)) as u32;

    for y in min_y..=max_y {
        for x in min_x..=max_x {
            let dx = x as f32 + 0.5 - cx;
            let dy = y as f32 + 0.5 - cy;
            let dist = (dx * dx + dy * dy).sqrt();

            let ring_dist = (dist - radius).abs();
            if ring_dist > thickness + halo_width {
                continue;
            }

            let pixel_angle = dy.atan2(dx);
            let relative = (pixel_angle - angle).rem_euclid(two_pi);
            if relative > arc_length {
                continue;
            }

            // Smooth gradient: ease-in from tail (transparent) to head (bright)
            let t = relative / arc_length;
            let t_smooth = t * t; // quadratic ease-in

            // Anti-aliased smooth edge, and the halo under it: full where the arc
            // covers it and fading out from the arc's edge, which is what leaves a
            // dark rim past an arc that is white.
            let arc = (1.0 - (ring_dist - thickness + 1.0).max(0.0)).clamp(0.0, 1.0) * t_smooth;
            let halo = (1.0 - (ring_dist - thickness).max(0.0) / halo_width).clamp(0.0, 1.0)
                * halo_opacity
                * t_smooth;

            // The arc over the halo, over nothing at all: the alpha is how much of
            // the two there is, and the colour is what is left of them once that is
            // known — white where the arc covers, the halo's shade where only the
            // halo does.
            let alpha = arc + halo * (1.0 - arc);
            if alpha <= 0.0 {
                continue;
            }
            let shade = ((255.0 * arc + halo_shade * halo * (1.0 - arc)) / alpha).clamp(0.0, 255.0);

            let idx = ((y * width + x) * 4) as usize;
            let shade = shade as u8;
            pixels[idx] = shade; // B
            pixels[idx + 1] = shade; // G
            pixels[idx + 2] = shade; // R
            pixels[idx + 3] = (alpha.clamp(0.0, 1.0) * 255.0) as u8;
        }
    }

    pixels
}

/// Create a loading animation MediaData for the given dimensions
pub(super) fn create_loading_media(width: u32, height: u32) -> MediaData {
    let pixels = render_loading_frame(width, height, 0.0);
    let frame = ImageFrame::new(pixels, width, height, 33);
    MediaData {
        frames: vec![Arc::new(frame)],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::Loading,
        stream_cancel: None,
        video_process: None,
        loading_start: Some(Instant::now()),
        text_state: None,
        audio_options: None,
    }
}

/// The pin's own arc, in the window's own pixels: how large its ring is, how thick it is, how
/// far its halo reaches past it, and how far past that the box it is walked in has to go.
///
/// It is the arc a hover waits with — the same 270° sweep, the same white over a dark halo, the
/// same ease from tail to head (see `render_loading_frame`) — at the size a *window* wants. A
/// hover's arc is sized by the box it is placed in, which is the arc's own box and is small,
/// because it stands at a pointer's corner; this one is drawn into the middle of a window the
/// user is reading, where a ring the size of the corner spinner is a mark on a picture rather
/// than an answer about the window. It is drawn straight onto the band rather than rendered into
/// a frame of its own and copied over, which is what keeps a repaint every thirty-odd
/// milliseconds from allocating a square of pixels to throw away.
pub(super) const PIN_ARC_RADIUS: f32 = 20.0;
pub(super) const PIN_ARC_THICKNESS: f32 = 5.0;
/// How far the halo reaches past the arc, and how much of it there is where it meets the arc's
/// own edge — both drawn from the arc's own size, so a larger arc gets a halo in proportion.
/// The shade is what the halo is made of: dark enough to read over a light file list, and
/// light enough to read as a shadow rather than a second arc (see `render_loading_frame`).
pub(super) const PIN_ARC_HALO: f32 = 5.0;
pub(super) const PIN_ARC_HALO_OPACITY: f32 = 0.85;
pub(super) const PIN_ARC_HALO_SHADE: f32 = 12.0;
/// How far from the centre the arc's own ring reaches: the halo, and one pixel more for the
/// edge it fades out over. The box the drawing walks is this either side of the centre.
pub(super) const PIN_ARC_REACH: i32 =
    (PIN_ARC_RADIUS + PIN_ARC_THICKNESS + PIN_ARC_HALO + 1.0) as i32;

/// Publish a pinned window's arc — how long its wait has been running — or take it down, which
/// is what a window that has answered, or a pin that has gone, is left with.
///
/// A wait publishes nothing until it has run for the delay `spinner_delay_ms` names, because a
/// load that answers inside the delay is a load the user never sees a spinner for, and that is
/// the whole of what the delay is for (see `PinLoad::due`).
pub(super) fn pin_arc_set(since: Option<Duration>) {
    // A wait that has not run a whole millisecond is still a wait, so a `0` — which is how
    // this says there is no arc — is never what a wait publishes.
    PIN_ARC.store(
        since.map_or(0, |since| (since.as_millis() as u64).max(1)),
        Ordering::Release,
    );
}

/// Draw the arc a pinned window is waiting for a file with, in the middle of its media band.
///
/// It goes *over* the band rather than in place of it: the file the pin is showing is the pin
/// until the new one has answered, and a band cleared for a spinner would be a window with
/// nothing in it for the length of a decode — the wait this exists for made visible as the
/// freeze it was meant to answer. The halo is what makes that legible over either a light
/// picture or a dark one, which is the same reason a hover's own arc carries one.
///
/// A band with no room for the arc is left as it is rather than drawn into: a window smaller
/// than the arc is a window whose wait is the caption changing, and half a ring is a mark
/// rather than a wait.
pub(super) fn paint_pin_spinner(out: &mut [u8], out_width: u32, band_top: i32, band_height: i32) {
    let since = PIN_ARC.load(Ordering::Acquire);
    if since == 0 {
        return;
    }

    let band_top = band_top.max(0) as u32;
    let band_height = band_height.max(1) as u32;
    if band_height <= PIN_ARC_REACH as u32 * 2 || out_width <= PIN_ARC_REACH as u32 * 2 {
        return;
    }

    // The middle of the band: the middle of the window for a kind whose chrome is drawn over its
    // media, and the middle of what is left between the two strips for a kind whose chrome has
    // bands of its own (see `pinned_band_rows`).
    let centre_x = out_width as f32 / 2.0;
    let centre_y = band_top as f32 + band_height as f32 / 2.0;

    let left = (centre_x as i32 - PIN_ARC_REACH).max(0) as u32;
    let right = ((centre_x as i32 + PIN_ARC_REACH) as u32).min(out_width - 1);
    let top = (centre_y as i32 - PIN_ARC_REACH).max(0) as u32;
    let bottom = ((centre_y as i32 + PIN_ARC_REACH) as u32).min(band_top + band_height - 1);

    let elapsed = PIN_ARC_BASE
        .elapsed()
        .saturating_sub(Duration::from_millis(since));
    let angle = elapsed.as_secs_f32() * std::f32::consts::PI * 2.4;

    let two_pi = std::f32::consts::PI * 2.0;
    let arc_length = std::f32::consts::PI * 1.5; // 270-degree arc
    let out_width = out_width as usize;

    for y in top..=bottom {
        for x in left..=right {
            let dx = x as f32 + 0.5 - centre_x;
            let dy = y as f32 + 0.5 - centre_y;
            let dist = (dx * dx + dy * dy).sqrt();

            // Only the arc's own ring and the halo the arc's edge fades out over.
            let ring_dist = (dist - PIN_ARC_RADIUS).abs();
            if ring_dist > PIN_ARC_THICKNESS + PIN_ARC_HALO {
                continue;
            }

            let relative = (dy.atan2(dx) - angle).rem_euclid(two_pi);
            if relative > arc_length {
                continue;
            }

            // Smooth gradient: ease-in from tail (transparent) to head (bright), with the
            // anti-aliased edge of the arc over the anti-aliased edge of its halo.
            let t_smooth = (relative / arc_length).powi(2);
            let arc =
                (1.0 - (ring_dist - PIN_ARC_THICKNESS + 1.0).max(0.0)).clamp(0.0, 1.0) * t_smooth;
            let halo = (1.0 - (ring_dist - PIN_ARC_THICKNESS).max(0.0) / PIN_ARC_HALO)
                .clamp(0.0, 1.0)
                * PIN_ARC_HALO_OPACITY
                * t_smooth;

            let index = (y as usize * out_width + x as usize) * 4;
            let Some(pixel) = out.get_mut(index..index + 4) else {
                continue;
            };

            blend_pin_arc_pixel(arc, halo, pixel);
        }
    }
}

/// Put one pixel of a pinned window's arc over one pixel of its band.
///
/// The arc over the halo, over the band: the coverage is how much of the two there is, and the
/// shade is what is left of them once that is known — white where the arc covers, the halo's
/// shade where only the halo does. What a layered window's surface holds is premultiplied
/// coverage (see `compose_preview_row`), so the shade is scaled by it on the way in and the
/// band's own pixel by what is left of it.
pub(super) fn blend_pin_arc_pixel(arc: f32, halo: f32, destination: &mut [u8]) {
    let alpha = (arc + halo * (1.0 - arc)).clamp(0.0, 1.0);
    if alpha <= 0.0 {
        return;
    }

    let shade = ((255.0 * arc + PIN_ARC_HALO_SHADE * halo * (1.0 - arc)) / alpha).clamp(0.0, 255.0);
    let keep = 1.0 - alpha;

    // The three colour channels, and the coverage after them: the arc is a colour over the
    // band and not a colour in place of it, which is what makes the halo read as a shadow.
    for channel in destination.iter_mut().take(3) {
        *channel = (shade * alpha + f32::from(*channel) * keep).clamp(0.0, 255.0) as u8;
    }
    destination[3] = (alpha * 255.0 + f32::from(destination[3]) * keep).clamp(0.0, 255.0) as u8;
}

/// The spinner's own geometry, in the frame's coordinates: how large the ring is, how thick
/// it is, and how far its corner sits from the frame's own.
pub(super) const SPINNER_OVERLAY_RADIUS: f32 = 8.0;
pub(super) const SPINNER_OVERLAY_THICKNESS: f32 = 2.5;
pub(super) const SPINNER_OVERLAY_PADDING: f32 = 12.0;

/// The box the corner spinner is drawn in: where it sits in the frame, and how large a box
/// holds it whole.
///
/// One answer for the copying and for the drawing, which have to agree: what the spinner is
/// drawn into is exactly what was copied out of the frame (see `overlay_loading_spinner`).
/// `None` for a frame too small to hold one at all, which is a frame the spinner is not drawn
/// on.
pub(super) fn spinner_overlay_box(width: u32, height: u32) -> Option<FrameBox> {
    if width < 24 || height < 24 {
        return None;
    }

    // How far the halo reaches past the centre, and one pixel more for the edge it fades
    // out over: the same reach the drawing below walks, so the box is never smaller than
    // what is drawn in it.
    let reach = SPINNER_OVERLAY_RADIUS + SPINNER_OVERLAY_THICKNESS + 4.0 + 1.0;
    let cx =
        width as f32 - SPINNER_OVERLAY_PADDING - SPINNER_OVERLAY_RADIUS - SPINNER_OVERLAY_THICKNESS;
    let cy = height as f32
        - SPINNER_OVERLAY_PADDING
        - SPINNER_OVERLAY_RADIUS
        - SPINNER_OVERLAY_THICKNESS;

    let left = ((cx - reach).max(0.0)) as u32;
    let top = ((cy - reach).max(0.0)) as u32;
    let right = ((cx + reach).min(width as f32 - 1.0)) as u32;
    let bottom = ((cy + reach).min(height as f32 - 1.0)) as u32;

    Some(FrameBox {
        left,
        top,
        width: right - left + 1,
        height: bottom - top + 1,
    })
}

/// Render a small loading spinner overlay onto a box of BGRA pixels copied out of a frame, in
/// place: a spinning arc in the bottom-right corner with a semi-transparent dark backdrop
/// circle.
///
/// The pixels are the box's and the geometry is the frame's — the corner the arc sits in is a
/// padding off the frame's own bottom-right, not off the box's — which is what keeps a spinner
/// drawn into a corner of a frame the same spinner it was when the whole frame was copied to
/// draw it (see `spinner_overlay_box`).
pub(super) fn overlay_loading_spinner(
    pixels: &mut [u8],
    area: FrameBox,
    frame_width: u32,
    frame_height: u32,
    angle: f32,
) {
    if pixels.len() < area.bytes() {
        return;
    }

    let (left, top) = (area.left, area.top);
    let (width, height) = (area.width, area.height);

    let radius = SPINNER_OVERLAY_RADIUS;
    let thickness = SPINNER_OVERLAY_THICKNESS;
    let backdrop_r = radius + thickness + 4.0;

    // Center of the spinner in the bottom-right corner of the frame
    let cx = frame_width as f32 - SPINNER_OVERLAY_PADDING - radius - thickness;
    let cy = frame_height as f32 - SPINNER_OVERLAY_PADDING - radius - thickness;

    let min_x = (((cx - backdrop_r - 1.0).max(0.0)) as u32).max(left);
    let max_x =
        (((cx + backdrop_r + 1.0).min(frame_width as f32 - 1.0)) as u32).min(left + width - 1);
    let min_y = (((cy - backdrop_r - 1.0).max(0.0)) as u32).max(top);
    let max_y =
        (((cy + backdrop_r + 1.0).min(frame_height as f32 - 1.0)) as u32).min(top + height - 1);

    let two_pi = std::f32::consts::PI * 2.0;
    let arc_length = std::f32::consts::PI * 1.5;

    for y in min_y..=max_y {
        for x in min_x..=max_x {
            let dx = x as f32 + 0.5 - cx;
            let dy = y as f32 + 0.5 - cy;
            let dist = (dx * dx + dy * dy).sqrt();
            let idx = (((y - top) * width + (x - left)) * 4) as usize;
            if idx + 3 >= pixels.len() {
                continue;
            }

            // Semi-transparent dark backdrop circle
            if dist <= backdrop_r {
                let edge = (1.0 - (dist - backdrop_r + 1.0).max(0.0)).clamp(0.0, 1.0);
                let bg_alpha = 0.45 * edge;
                if bg_alpha > 0.0 {
                    pixels[idx] = ((pixels[idx] as f32) * (1.0 - bg_alpha)) as u8;
                    pixels[idx + 1] = ((pixels[idx + 1] as f32) * (1.0 - bg_alpha)) as u8;
                    pixels[idx + 2] = ((pixels[idx + 2] as f32) * (1.0 - bg_alpha)) as u8;
                }
            }

            // Spinner ring
            let ring_dist = (dist - radius).abs();
            if ring_dist > thickness + 1.0 {
                continue;
            }
            let edge_alpha = (1.0 - (ring_dist - thickness + 1.0).max(0.0)).clamp(0.0, 1.0);
            if edge_alpha <= 0.0 {
                continue;
            }
            let pixel_angle = dy.atan2(dx);
            let relative = (pixel_angle - angle).rem_euclid(two_pi);
            if relative <= arc_length {
                let t = relative / arc_length;
                let t_smooth = t * t;
                let alpha = edge_alpha * t_smooth;
                let blend = |bg_c: u8, fg: u8, a: f32| -> u8 {
                    ((bg_c as f32) * (1.0 - a) + (fg as f32) * a).clamp(0.0, 255.0) as u8
                };
                pixels[idx] = blend(pixels[idx], 255, alpha);
                pixels[idx + 1] = blend(pixels[idx + 1], 255, alpha);
                pixels[idx + 2] = blend(pixels[idx + 2], 255, alpha);
            }
        }
    }
}

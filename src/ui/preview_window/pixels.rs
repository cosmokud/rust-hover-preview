//! Turning a decoded picture into the pixels a window is painted with: the header-sniffing
//! decodes, the tone-mapped read of a high-dynamic-range file, and the compose, scale and blend a
//! frame goes through on its way into a band.

use super::*;

/// Decode an image by sniffing its magic bytes rather than by the name it is written under.
///
/// A picture is decoded by its header always: a `.dat` holding a PNG is a picture, and a
/// `.png` holding a container is a container — which is the kind's question, and it has been
/// asked by the time anything here is reached.
pub(super) fn decode_image_with_header_check(path: &PathBuf) -> Option<image::DynamicImage> {
    let mut reader = image::ImageReader::open(path)
        .ok()?
        .with_guessed_format()
        .ok()?;
    reader.limits(image_decode_limits());

    reader.decode().ok()
}

/// Read image dimensions by sniffing magic bytes instead of trusting the extension.
///
/// A format the `image` crate has no reader for at all is a file it cannot answer for
/// rather than a file that is not a picture, so the codec Windows has is asked before
/// the answer is no: what a hover onto a `.heic`, or onto a still `.webp`, is measured
/// from is its own frame; see `wic_image`.
pub(super) fn image_dimensions_with_header_check(path: &PathBuf) -> Option<(u32, u32)> {
    image::ImageReader::open(path)
        .ok()?
        .with_guessed_format()
        .ok()?
        .into_dimensions()
        .ok()
        .or_else(|| codec_dimensions(path))
}

/// The size a picture of a format this app's own decoder does not read is measured
/// from: the frame the codec Windows has for it reports, and — where that codec is the
/// WebP one and the machine has none, which is what a Windows 10 machine usually is —
/// libwebp's, which is in the binary; see `wic_image` and `webp_image`. A `.dds` of a
/// format the codec does not read is measured by the decoder this app carries for the
/// rest of that format; see `dds_image`.
pub(super) fn codec_dimensions(path: &Path) -> Option<(u32, u32)> {
    wic_image::dimensions(path)
        .or_else(|| dds_image::dimensions(path))
        .or_else(|| webp_image::dimensions(path))
}

/// The eight-bit picture a picture whose samples are light is shown as.
///
/// `None` is every picture that holds levels already — a PNG, a JPEG, a BMP, and every
/// other format this app's reader decodes — because a level put through a transfer
/// function a second time is a washed-out picture rather than a corrected one.
///
/// What arrives as one of these is an `.exr` and a Radiance `.hdr`, which are the two
/// float kinds the `image` crate has: three channels for a `.hdr` — and for the `.exr`
/// that was written without an alpha — and four for the one that carries it. A
/// single-channel file is the decoder's to widen, and it arrives as one of the two as
/// well. What the curve does with them is `tone_map`'s; what is done here is the shape,
/// and an alpha is not light and is not put through a curve.
pub(super) fn tone_mapped_image(img: &image::DynamicImage) -> Option<image::RgbaImage> {
    let (samples, channels, width, height) = match img {
        image::DynamicImage::ImageRgb32F(buffer) => (
            buffer.as_raw().as_slice(),
            3,
            buffer.width(),
            buffer.height(),
        ),
        image::DynamicImage::ImageRgba32F(buffer) => (
            buffer.as_raw().as_slice(),
            4,
            buffer.width(),
            buffer.height(),
        ),
        _ => return None,
    };

    let tone = tone_map::ToneMap::current();
    let count = width as usize * height as usize;
    let mut pixels = vec![0u8; count * 4];

    for index in 0..count {
        let texel = &samples[index * channels..][..channels];

        let alpha = if channels == 4 { texel[3] } else { 1.0 };

        let at = index * 4;
        pixels[at] = tone.encode(texel[0]);
        pixels[at + 1] = tone.encode(texel[1]);
        pixels[at + 2] = tone.encode(texel[2]);
        pixels[at + 3] = tone_mapped_alpha(alpha);
    }

    image::RgbaImage::from_raw(width, height, pixels)
}

/// An alpha as a level: brought into the range and scaled, with no curve applied — what
/// the alpha of a float texture gets as well, and for the same reason: coverage is not
/// light, and a picture composited over a backdrop with a curve put on its alpha is a
/// picture that fades differently from every other one beside it.
pub(super) fn tone_mapped_alpha(alpha: f32) -> u8 {
    if !alpha.is_finite() {
        return 0;
    }

    (alpha.clamp(0.0, 1.0) * 255.0 + 0.5) as u8
}

/// Convert RGBA pixels to BGRA for Windows GDI
///
/// Shared with the readers that hand back their own pixels rather than going through the
/// `image` crate — a texture this app decodes itself, a codec's own frame — because a
/// frame is composed in one order whatever produced it (see `dds_image`).
pub(crate) fn rgba_to_bgra(rgba: &[u8]) -> Vec<u8> {
    let mut bgra = Vec::with_capacity(rgba.len());
    for chunk in rgba.chunks(4) {
        if chunk.len() == 4 {
            bgra.push(chunk[2]); // B
            bgra.push(chunk[1]); // G
            bgra.push(chunk[0]); // R
            bgra.push(chunk[3]); // A
        }
    }
    bgra
}

pub(super) fn checkerboard_color(x: u32, y: u32) -> (u8, u8, u8) {
    if ((x / 16) + (y / 16)).is_multiple_of(2) {
        (224, 224, 224)
    } else {
        (144, 144, 144)
    }
}

/// Compose `bgra` into `out` (BGRA, top-down) with `background` applied to the
/// alpha channel.
///
/// Writes into the caller's buffer so a repaint can target the layered window's
/// DIB directly, and walks it row by row so the per-pixel background position is
/// a row/column counter instead of a division.
///
/// `opaque` is the frame's own answer to whether every one of its pixels has an alpha of 255
/// (see `ImageFrame`). Where it does, the frame *is* the composed surface and the whole of it
/// is one copy: a pixel the backdrop cannot be seen through is the pixel the blend below would
/// have written, in every backdrop kind, because both the premultiply a transparent backdrop
/// asks for and the blend over an opaque one come to the pixel's own bytes where the alpha is
/// 255. That is the difference between a copy and a division per channel per pixel, on every
/// frame of a video or an animation drawn at the size of the display.
pub(super) fn compose_preview_pixels_into(
    bgra: &[u8],
    width: u32,
    height: u32,
    background: TransparentBackground,
    opaque: bool,
    out: &mut [u8],
) {
    let width = width as usize;
    let expected = width * height as usize * 4;
    if width == 0 || out.len() < expected {
        return;
    }

    // A short source used to end up zero padded; keep that.
    let usable = bgra.len() / 4 * 4;
    if usable < expected {
        out[usable..expected].fill(0);
    }

    if opaque && usable >= expected {
        out[..expected].copy_from_slice(&bgra[..expected]);
        return;
    }

    let row_bytes = width * 4;
    for (y, (src_row, dst_row)) in bgra
        .chunks_exact(row_bytes)
        .zip(out[..expected].chunks_exact_mut(row_bytes))
        .enumerate()
    {
        compose_preview_row(src_row, dst_row, background, 0, y as u32);
    }
}

/// The same for a box of a frame rather than the whole of it: the box is where it sits in the
/// frame, which is both what the checkerboard's squares are placed by and what the destination's
/// rows are offset by.
///
/// It is what the corner spinner is composed back through: what is drawn over a frame is drawn
/// into a copy of the corner it sits in rather than into a copy of the frame, and the box is
/// what that copy is (see `render_layered_preview_at`). A box carries what was drawn over the
/// frame, so it is blended however opaque the frame under it was — it is a few thousand pixels
/// either way.
pub(super) fn compose_preview_block_into(
    bgra: &[u8],
    area: FrameBox,
    background: TransparentBackground,
    out: &mut [u8],
    out_width: u32,
) {
    let row_bytes = area.width as usize * 4;
    let out_row_bytes = out_width as usize * 4;
    let expected = area.bytes();

    if area.width == 0 || area.height == 0 || bgra.len() < expected {
        return;
    }

    for (y, src_row) in bgra[..expected].chunks_exact(row_bytes).enumerate() {
        let row = area.top as usize + y;
        let start = row * out_row_bytes + area.left as usize * 4;
        let Some(dst_row) = out.get_mut(start..start + row_bytes) else {
            return;
        };

        compose_preview_row(src_row, dst_row, background, area.left, row as u32);
    }
}

/// A box of a frame's pixels, copied out whole: the corner a spinner is drawn into.
pub(super) fn copy_frame_box_into(
    frame: &[u8],
    frame_width: u32,
    area: FrameBox,
    out: &mut Vec<u8>,
) {
    let row_bytes = area.width as usize * 4;
    let frame_row_bytes = frame_width as usize * 4;
    out.clear();
    out.resize(area.bytes(), 0);

    for y in 0..area.height as usize {
        let from = (area.top as usize + y) * frame_row_bytes + area.left as usize * 4;
        let Some(source) = frame.get(from..from + row_bytes) else {
            return;
        };

        out[y * row_bytes..(y + 1) * row_bytes].copy_from_slice(source);
    }
}

pub(super) fn compose_preview_row(
    src_row: &[u8],
    dst_row: &mut [u8],
    background: TransparentBackground,
    x_offset: u32,
    y: u32,
) {
    match background {
        TransparentBackground::Transparent => {
            for (px, dst) in src_row
                .as_chunks::<4>()
                .0
                .iter()
                .zip(dst_row.as_chunks_mut::<4>().0.iter_mut())
            {
                let b = px[0] as u32;
                let g = px[1] as u32;
                let r = px[2] as u32;
                let a = px[3] as u32;

                dst[0] = ((b * a + 127) / 255) as u8;
                dst[1] = ((g * a + 127) / 255) as u8;
                dst[2] = ((r * a + 127) / 255) as u8;
                dst[3] = a as u8;
            }
        }
        // The background is settled once per row rather than once per pixel: it
        // cannot change inside the loop, and as a per-pixel match it cost a
        // branch on every pixel of every animation frame.
        TransparentBackground::Black => {
            for (px, dst) in src_row
                .as_chunks::<4>()
                .0
                .iter()
                .zip(dst_row.as_chunks_mut::<4>().0.iter_mut())
            {
                blend_pixel_over(px, dst, 0, 0, 0);
            }
        }
        TransparentBackground::White => {
            for (px, dst) in src_row
                .as_chunks::<4>()
                .0
                .iter()
                .zip(dst_row.as_chunks_mut::<4>().0.iter_mut())
            {
                blend_pixel_over(px, dst, 255, 255, 255);
            }
        }
        TransparentBackground::Checkerboard => {
            for (x, (px, dst)) in src_row
                .as_chunks::<4>()
                .0
                .iter()
                .zip(dst_row.as_chunks_mut::<4>().0.iter_mut())
                .enumerate()
            {
                // Squares are placed by where a pixel is in the *frame*, which for a box is
                // not where it is in the box.
                let (r, g, b) = checkerboard_color(x as u32 + x_offset, y);
                blend_pixel_over(px, dst, b as u32, g as u32, r as u32);
            }
        }
    }
}

/// One pixel blended over an opaque background, written opaquely.
#[inline]
pub(super) fn blend_pixel_over(px: &[u8], dst: &mut [u8], bg_b: u32, bg_g: u32, bg_r: u32) {
    let a = px[3] as u32;
    let inv_a = 255 - a;

    dst[0] = ((px[0] as u32 * a + bg_b * inv_a + 127) / 255) as u8;
    dst[1] = ((px[1] as u32 * a + bg_g * inv_a + 127) / 255) as u8;
    dst[2] = ((px[2] as u32 * a + bg_r * inv_a + 127) / 255) as u8;
    dst[3] = 255;
}

/// The scale a preview of `orig` size takes inside a room, for the scale the
/// configuration asked for.
///
/// One function answers it for every use — the size a preview is given, and the
/// room the position modes choose between — so the scale the layout plans and the
/// scale the renderer draws at cannot come apart.
pub(super) fn scale_in_room(
    room_width: f32,
    room_height: f32,
    orig_width: f32,
    orig_height: f32,
    preview_scale: PreviewScale,
) -> f32 {
    let fit_scale = (room_width / orig_width).min(room_height / orig_height);

    // A requested percentage is honored when it fits; anything larger than the
    // available area falls back to the fit scale so nothing is ever clipped. A
    // scale that is a reduction of the fit is a share *of* it, and is applied
    // after it.
    let scale = match preview_scale.target_scale() {
        Some(target_scale) => target_scale.min(fit_scale),
        None => fit_scale,
    };

    scale * preview_scale.fit_share()
}

/// Scale media dimensions to the requested preview scale while never exceeding
/// `max_width`/`max_height`, so the preview always stays fully inside the screen.
pub(super) fn scale_dimensions(
    orig_width: u32,
    orig_height: u32,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> (u32, u32) {
    let scale = scale_in_room(
        max_width as f32,
        max_height as f32,
        orig_width as f32,
        orig_height as f32,
        preview_scale,
    );

    // Rounded so a fit scale lands on the exact available size, then clamped so
    // float error can never push the preview past the screen edge.
    let new_width = (orig_width as f32 * scale)
        .round()
        .clamp(1.0, max_width.max(1) as f32) as u32;
    let new_height = (orig_height as f32 * scale)
        .round()
        .clamp(1.0, max_height.max(1) as f32) as u32;

    (new_width, new_height)
}

/// Animation frames stream continuously, so shrinking keeps the cheap nearest
/// filter while enlarging uses a smoother filter to avoid blocky previews.
pub(super) fn frame_resize_filter(
    orig_width: u32,
    orig_height: u32,
    target_width: u32,
    target_height: u32,
) -> image::imageops::FilterType {
    if target_width > orig_width || target_height > orig_height {
        image::imageops::FilterType::Triangle
    } else {
        image::imageops::FilterType::Nearest
    }
}

/// Decode a single GIF frame from canvas to an ImageFrame
pub(super) fn decode_gif_frame_to_image(
    canvas: &[u8],
    gif_width: u32,
    gif_height: u32,
    target_width: u32,
    target_height: u32,
    delay_ms: u32,
) -> Option<ImageFrame> {
    let scaled = if target_width != gif_width || target_height != gif_height {
        let img = image::RgbaImage::from_raw(gif_width, gif_height, canvas.to_vec())?;
        let resized = image::imageops::resize(
            &img,
            target_width,
            target_height,
            frame_resize_filter(gif_width, gif_height, target_width, target_height),
        );
        resized.into_raw()
    } else {
        canvas.to_vec()
    };

    let bgra = rgba_to_bgra(&scaled);

    Some(ImageFrame::new(bgra, target_width, target_height, delay_ms))
}

/// Composite a GIF frame onto the canvas
pub(super) fn composite_gif_frame(
    canvas: &mut [u8],
    frame: &gif::Frame,
    gif_width: u32,
    gif_height: u32,
) {
    let frame_x = frame.left as usize;
    let frame_y = frame.top as usize;
    let frame_w = frame.width as usize;
    let frame_h = frame.height as usize;

    for y in 0..frame_h {
        for x in 0..frame_w {
            let src_idx = (y * frame_w + x) * 4;
            let dst_x = frame_x + x;
            let dst_y = frame_y + y;
            if dst_x < gif_width as usize && dst_y < gif_height as usize {
                let dst_idx = (dst_y * gif_width as usize + dst_x) * 4;
                if src_idx + 3 < frame.buffer.len() {
                    let alpha = frame.buffer[src_idx + 3];
                    if alpha > 0 {
                        canvas[dst_idx] = frame.buffer[src_idx];
                        canvas[dst_idx + 1] = frame.buffer[src_idx + 1];
                        canvas[dst_idx + 2] = frame.buffer[src_idx + 2];
                        canvas[dst_idx + 3] = alpha;
                    }
                }
            }
        }
    }
}

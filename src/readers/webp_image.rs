//! A still WebP, decoded by libwebp: the reader for a machine whose Windows has no
//! codec for one.
//!
//! The codec Windows has for WebP is not one Windows ships — it is the **WebP Image
//! Extension**, a Store package Windows 11 comes with and Windows 10 usually has to
//! be given — so a machine without it is a machine the codec path answers with nothing
//! for. A picture is not something this app loses to a missing package: libwebp is in
//! the binary already, because it is what plays the WebP that *moves* (see the
//! animated reader in `preview_window`), so a still is decoded by it where the codec
//! has nothing to say. The codec is still asked first — it is the reader the machine
//! prefers, and it is the one that also answers for the formats this app has no decoder
//! for at all (see `wic_image`) — so what reaches this module is the WebP the codec has
//! no answer for.
//!
//! What the decoder is asked for is the box the layout planned rather than the file's
//! own size, which is the shape the codec path takes and the reason it takes it: a
//! forty-megapixel photograph is a fraction of a preview, and libwebp's scaler takes
//! the size as it decodes. What comes back is the BGRA a frame is composed in, so the
//! resample and the two conversions the app's own reader pays for are not paid for
//! here at all.
//!
//! Nothing here is allowed to take the app down with it, and none of it is done before
//! the file has earned it: the twelve bytes that say whether a file is a WebP at all
//! are read first, so a `.heic` the codec declined costs a header rather than the whole
//! file; what is read whole is bounded by the budget every other reader is read under;
//! and the file's own frame is checked against that budget before anything is
//! allocated. A file that is not a WebP, a file that will not decode and a picture past
//! the budget are one answer: no preview.

use crate::config::config::{frame_bytes_within_budget, read_within_budget};
use libwebp_sys::{
    MODE_BGRA, VP8_STATUS_OK, WebPDecode, WebPDecoderConfig, WebPGetInfo,
    WebPInitDecoderConfig, WebPRGBABuffer,
};
use std::io::Read;
use std::path::Path;

/// A still WebP's own size, which is the size the layout places it at.
///
/// `None` is a file that is not a WebP and one that will not decode, and both are the
/// same answer: no preview.
pub fn dimensions(path: &Path) -> Option<(u32, u32)> {
    dimensions_of(&read(path)?)
}

/// A still WebP decoded to exactly `width` by `height`, in the order the preview's own
/// frame is composed in: BGRA, top-down, four bytes to the pixel.
pub fn decode(path: &Path, width: u32, height: u32) -> Option<Vec<u8>> {
    if width == 0 || height == 0 {
        return None;
    }

    let bytes = read(path)?;
    let (source_width, source_height) = dimensions_of(&bytes)?;

    // The budget every other reader is handed before it allocates: it is asked of the
    // file's own shape rather than of the preview's, because the preview's is this
    // side's to choose and the file's is not.
    frame_bytes_within_budget(source_width, source_height, 4)?;

    let stride = (width as usize).checked_mul(4)?;
    let mut pixels = vec![0u8; stride.checked_mul(height as usize)?];

    // The frame is written into the buffer this side owns rather than into one of
    // libwebp's that would then be copied out of, and the size it is written at is the
    // preview's own.
    let mut config: WebPDecoderConfig = unsafe { std::mem::zeroed() };
    if unsafe { WebPInitDecoderConfig(&mut config) } == 0 {
        return None;
    }

    config.output.colorspace = MODE_BGRA;
    config.output.is_external_memory = 1;
    config.options.use_scaling = 1;
    config.options.scaled_width = width as i32;
    config.options.scaled_height = height as i32;
    config.output.u.RGBA = WebPRGBABuffer {
        rgba: pixels.as_mut_ptr(),
        stride: stride as i32,
        size: pixels.len(),
    };

    let status = unsafe { WebPDecode(bytes.as_ptr(), bytes.len(), &mut config) };

    (status == VP8_STATUS_OK).then_some(pixels)
}

/// The file's bytes, when it is a WebP at all.
///
/// The container's own signature is settled first — `RIFF`, a little-endian size, and
/// `WEBP` — so a file that is not one costs twelve bytes and no more. What is read
/// whole beyond that is read under the budget a hover may ask for, the same way the
/// animated reader reads its own file.
fn read(path: &Path) -> Option<Vec<u8>> {
    let mut head = [0u8; 12];
    std::fs::File::open(path).ok()?.read_exact(&mut head).ok()?;

    if &head[0..4] != b"RIFF" || &head[8..12] != b"WEBP" {
        return None;
    }

    read_within_budget(path)
}

/// A WebP's own size, read from its header: what the layout is placed from, and what
/// the budget above is asked of.
///
/// A header that will not report one is a file libwebp has refused, which is the same
/// answer a file that will not decode gets.
fn dimensions_of(bytes: &[u8]) -> Option<(u32, u32)> {
    let (mut width, mut height) = (0i32, 0i32);
    let read = unsafe { WebPGetInfo(bytes.as_ptr(), bytes.len(), &mut width, &mut height) };

    (read != 0 && width > 0 && height > 0).then_some((width as u32, height as u32))
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A file that is not a WebP is answered with no picture rather than with a guess,
    /// and the name it carries has nothing to do with it: the container's own signature
    /// is what this reader asks.
    #[test]
    fn a_file_that_is_not_a_webp_is_no_picture() {
        let folder = std::env::temp_dir().join("rust-hover-preview-webp-tests");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let png_named_webp = folder.join("actually-a-png.webp");
        std::fs::write(&png_named_webp, b"\x89PNG\r\n\x1a\n this is not a WebP at all")
            .expect("a written file");

        let riff_but_not_webp = folder.join("riff-of-something-else.webp");
        let mut bytes = b"RIFF".to_vec();
        bytes.extend_from_slice(&[0, 0, 0, 0]);
        bytes.extend_from_slice(b"WAVEfmt ");

        std::fs::write(&riff_but_not_webp, &bytes).expect("a written file");

        for path in [
            &png_named_webp,
            &riff_but_not_webp,
            &folder.join("missing.webp"),
        ] {
            assert_eq!(dimensions(path), None, "{path:?}");
            assert_eq!(decode(path, 16, 16), None, "{path:?}");
        }
    }
}


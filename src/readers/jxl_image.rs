//! JPEG XL, decoded by `jxl-oxide`: the reader for a machine whose Windows has no codec
//! for one.
//!
//! The codec Windows has for JPEG XL is not one Windows ships — it is the **JPEG XL
//! Image Extension**, a Store package Windows 11 comes with and a Windows 10 machine
//! usually has to be given — so the codec is still asked first (see `wic_image`) and what
//! reaches this module is the `.jxl` it has no answer for. `jxl-oxide` answers that, and is
//! a better fit than the codec would have been anyway: a JPEG XL file is not a picture, it
//! is a sequence of frames in time, and a codec that reads one of them as a still has thrown
//! away half of what the file is.
//!
//! Whether a file is one picture or a sequence is not written anywhere a byte pattern can be
//! read — the flag is in the codestream's own image header, behind whatever container boxes
//! the file carries — so a `.jxl` whose name says nothing has to be opened before it can be
//! said what it is. The size is in the same header, so both are answered from the first
//! sixty-four kilobytes rather than from all of it; only the pixels need the whole codestream,
//! and only the pixels are read under a hover's budget.
//!
//! The one thing in here that is not obvious from the signatures, and the one a future
//! reader will get wrong, is the timing. **A frame's duration is in ticks, not in
//! milliseconds** (see `ticks_to_ms`).

use crate::config::config::{decode_budget_bytes, frame_bytes_within_budget, read_within_budget};
use jxl_oxide::{AllocTracker, InitializeResult, JxlImage, NullCms, PixelFormat, Render};
use std::io::Read;
use std::path::Path;

/// How much of a file is read to answer the two questions its header answers.
///
/// A container puts a signature, a box header and however much metadata the encoder
/// chose to embed in front of the codestream, so this is generous rather than tight; it
/// is still a fraction of a file, and a file whose header lies further back than this is
/// answered with no preview rather than by reading all of it to find out.
const HEADER_BYTES: u64 = 64 * 1024;

/// The two bytes a bare codestream begins with, which are also the first two bits of the
/// image header's own signature — the header is what says a file is a JPEG XL at all.
const CODESTREAM_SIGNATURE: [u8; 2] = [0xff, 0x0a];

/// The twelve bytes a container begins with: a box header of size twelve, the box type
/// `JXL `, and the signature that follows it.
const CONTAINER_SIGNATURE: [u8; 12] = [
    0x00, 0x00, 0x00, 0x0c, b'J', b'X', b'L', 0x20, 0x0d, 0x0a, 0x87, 0x0a,
];

/// Whether the file's own header says it is a sequence in time rather than one picture.
///
/// This is what the head probe asks before anything else about a `.jxl`, and it is the
/// cheapest real decoder there is: the flag is inside the codestream, so it is read by
/// feeding the opening of the file to a decoder that has not been told anything about it.
/// A file that is not a JPEG XL, one that is not there, and one whose header is past what
/// the opening is read are all the same answer: it is not one of these.
pub(crate) fn is_animated(path: &Path) -> bool {
    open_header(path).is_some_and(|image| image.image_header().metadata.animation.is_some())
}

/// One decoded frame, full canvas, in the order a frame of this app is composed in: BGRA,
/// top-down, four bytes to the pixel (see `rgba_to_bgra` in `preview_window`).
///
/// `keyframe_index` counts keyframes, not frames — a JPEG XL animation may hold frames
/// between them that are only ever blended into the one after — and a picture past the box
/// the layout planned is answered with nothing rather than drawn smaller, since there is no
/// scaler to ask for a box the way libwebp has one (see `webp_image`).
pub(crate) fn decode_frame(
    path: &Path,
    keyframe_index: usize,
    max_width: u32,
    max_height: u32,
) -> Option<Vec<u8>> {
    if max_width == 0 || max_height == 0 {
        return None;
    }

    // The file is read whole, and under the budget a hover may ask for, the same way the
    // animated readers read their own.
    let bytes = read_within_budget(path)?;
    let image = open(&bytes)?;

    let (width, height) = size(&image)?;
    if width > max_width || height > max_height {
        return None;
    }

    // The budget every other reader is handed before it allocates, asked of the frame the
    // file would be decoded into: this app composes in four bytes to the pixel.
    frame_bytes_within_budget(width, height, 4)?;

    compose(&image, &image.render_frame(keyframe_index).ok()?)
}

/// The number of keyframes in the file, each keyframe's delay in **milliseconds**, and the
/// size every frame of it is decoded at.
///
/// `None` is a file that is not one, one whose header carries no animation, and one holding a
/// single keyframe — one answer, and the cue to fall back to the still path.
pub(crate) fn animation(path: &Path) -> Option<(Vec<u32>, (u32, u32))> {
    let bytes = read_within_budget(path)?;
    let image = open(&bytes)?;

    // A file whose frames are not all loaded is one whose timing is not all known: the
    // delays below are the whole of what the engine is told, so a sequence that stops
    // halfway is a sequence this app does not play at all rather than one played at the
    // wrong speed after its last frame.
    if !image.is_loading_done() {
        return None;
    }

    let timing = image.image_header().metadata.animation.as_ref()?;
    let (width, height) = size(&image)?;

    let keyframes = image.num_loaded_keyframes();
    if keyframes < 2 {
        return None;
    }

    let mut delays = Vec::with_capacity(keyframes);
    for index in 0..keyframes {
        // A frame with no duration of its own is one the format says is shown with the
        // frame after it rather than on its own; the delay is passed on as it is read,
        // because zero is what it means.
        let ticks = image.frame_header(index)?.duration;
        delays.push(ticks_to_ms(
            ticks,
            timing.tps_numerator,
            timing.tps_denominator,
        ));
    }

    Some((delays, (width, height)))
}

/// A frame's duration, in ticks, as the milliseconds the animation engine is given.
///
/// **The conversion, which is the one thing in this module a future reader will get
/// wrong: a duration is in ticks, not in milliseconds.** The animation header gives a
/// ticks-per-second as a fraction, and a frame's duration counts ticks of it, so a
/// millisecond is `ticks * denominator * 1000 / numerator` — a file at the usual
/// thousand ticks a second converts a hundred ticks to a hundred milliseconds, and a
/// file at one tick a second converts one tick to a second.
///
/// The arithmetic is done in 128 bits rather than 64, and it has to be: `ticks` and
/// `denominator` are both `u32` and both are whatever the file says they are, and
/// `u32::MAX * u32::MAX * 1000` is about three orders of magnitude past what 64 bits
/// carry. Widening to 64 is not enough — a file claiming the largest tick and the
/// coarsest timebase at once would wrap into a delay of a fraction of a millisecond in a
/// release build, and panic outright in a debug one. The result is held to what the
/// engine's own `u32` can carry. A still has no animation header at all and so no
/// timebase, which is why the numerator is asked about before it divides by it.
fn ticks_to_ms(ticks: u32, numerator: u32, denominator: u32) -> u32 {
    if numerator == 0 {
        return 0;
    }

    let milliseconds = u128::from(ticks) * u128::from(denominator) * 1000 / u128::from(numerator);

    milliseconds.min(u128::from(u32::MAX)) as u32
}

/// A file's own size, which is the size the layout is placed at and the size the budget
/// is asked of. A shape with no pixels in it is a file this app does not preview.
fn size(image: &JxlImage) -> Option<(u32, u32)> {
    let (width, height) = (image.width(), image.height());

    (width > 0 && height > 0).then_some((width, height))
}

/// The opening of a file, when its first bytes say it is a JPEG XL at all.
///
/// The two ways a file begins are read before any of it is, so a `.png` the codec
/// declined costs a few kilobytes rather than a whole decode. What is read beyond the
/// signature is bounded rather than whole: a file's whole bytes are read under the
/// budget only by the two calls that need the frames.
fn read_head(path: &Path) -> Option<Vec<u8>> {
    let mut head = Vec::new();
    std::fs::File::open(path)
        .ok()?
        .take(HEADER_BYTES)
        .read_to_end(&mut head)
        .ok()?;

    begins_like_jxl(&head).then_some(head)
}

/// Whether the opening bytes given are those of a JPEG XL, in either of the two forms it
/// arrives in — a naked codestream, or a container wrapping one.
///
/// This is the whole of the cheap question, and it is asked *before* the file is opened
/// for anything else. It has to be, because the head probe asks whether a file is a
/// sequence for a great many files that are not JPEG XL at all: a PDF, a ZIP, a JPEG, an
/// executable. Reading sixty-four kilobytes off each of those to discover they do not
/// begin `FF 0A` is a cost this app's hover path cannot pay on every file it is handed.
///
/// The twelve-byte container signature is the longest of the two, so that is what this
/// asks for; a two-byte codestream signature is decided by the same read.
pub(crate) fn begins_like_jxl(front: &[u8]) -> bool {
    front.starts_with(&CODESTREAM_SIGNATURE) || front.starts_with(&CONTAINER_SIGNATURE)
}

/// The file's header, read from the opening of the file and no further.
///
/// This is what a file's size and a file's being a sequence are both answered from, and
/// it is deliberately not the whole file: the whole codestream is read by the two calls
/// that need frames. The decoder is fed what there is until it either has the header or
/// has run out of file to read — a header that needs more than this is a file answered
/// with no preview, not a file read in full to find out.
fn open_header(path: &Path) -> Option<JxlImage> {
    let bytes = read_head(path)?;
    let mut uninitialized = JxlImage::builder().build_uninit();
    let mut fed = 0;

    loop {
        let consumed = uninitialized.feed_bytes(&bytes[fed..]).ok()?;
        fed += consumed;

        match uninitialized.try_init().ok()? {
            InitializeResult::Initialized(image) => return Some(image),
            InitializeResult::NeedMoreData(next) => {
                uninitialized = next;

                // A decoder that has taken every byte there is and still wants more wants
                // a file this reader will not read further for — and one that is taking
                // none of the bytes it is given is not going to take them next time
                // either, so both are a file with no header in it.
                if consumed == 0 || fed >= bytes.len() {
                    return None;
                }
            }
        }
    }
}

/// A whole file's bytes decoded into a decoder.
///
/// The caller has read the bytes under the budget a hover may ask for; the decoder is
/// handed that same budget as an allocation limit of its own, so a header that claims an
/// enormous shape is refused inside the decode rather than by an allocation this side
/// made on the header's word.
fn open(bytes: &[u8]) -> Option<JxlImage> {
    let budget = usize::try_from(decode_budget_bytes()).ok()?;

    let mut image = JxlImage::builder()
        .alloc_tracker(AllocTracker::with_limit(budget))
        .read(bytes)
        .ok()?;

    // No colour management system is in the binary, and `NullCms` is how that is said to
    // the decoder: a file whose colour is only reachable through an ICC transform is
    // answered with no preview rather than with a picture in the wrong colour. The colour
    // work the decoder does itself — XYB, the transfer functions, the tone map — goes
    // through no CMS and is unaffected.
    image.set_cms(NullCms);

    Some(image)
}

/// One rendered frame, composed the way a frame of this app is composed.
///
/// The samples come out of the stream rather than out of a frame buffer because the
/// stream is the one that has the file's orientation applied and the one that has already
/// chosen which of the file's channels are colour, black and alpha; a frame buffer holds
/// every extra channel a JPEG XL file can carry, spot colours and depth among them, which
/// a preview has no use for.
fn compose(image: &JxlImage, render: &Render) -> Option<Vec<u8>> {
    let format = image.pixel_format();

    // A CMYK picture's colour is only reachable through a transform the decoder has to
    // be handed a colour management system for, and this binary has none (see `open`), so
    // one is answered with no preview rather than with a picture whose channels have been
    // read as though they were red, green and blue.
    if format.has_black() {
        return None;
    }

    let mut stream = render.stream();

    let (width, height, channels) = (stream.width(), stream.height(), stream.channels());
    if width == 0 || height == 0 || channels != format.channels() as u32 {
        return None;
    }

    let count = (width as usize)
        .checked_mul(height as usize)?
        .checked_mul(channels as usize)?;
    let mut samples = vec![0u8; count];

    if stream.write_to_buffer(&mut samples) != count {
        return None;
    }

    let pixel_count = count / channels as usize;
    let mut pixels = Vec::with_capacity(pixel_count.checked_mul(4)?);

    for sample in samples.chunks_exact(channels as usize) {
        let texel = match format {
            PixelFormat::Gray => [sample[0], sample[0], sample[0], 255],
            PixelFormat::Graya => [sample[0], sample[0], sample[0], sample[1]],
            // The colour channels are in red, green, blue order, which is the reverse of
            // the order a frame of this app is composed in; the alpha is where it is.
            PixelFormat::Rgb | PixelFormat::Rgba => {
                let alpha = if format.has_alpha() { sample[3] } else { 255 };
                [sample[2], sample[1], sample[0], alpha]
            }
            // Refused above, and named here so that a file of this kind reaching this line
            // is a mistake rather than an unhandled case.
            PixelFormat::Cmyk | PixelFormat::Cmyka => return None,
        };

        pixels.extend_from_slice(&texel);
    }

    Some(pixels)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A folder of files that are not JPEG XLs, each named as though it were one.
    ///
    /// A valid codestream cannot be hand-authored — a few bits of image header, then a
    /// frame header, then entropy-coded data whose meaning depends on both — so the
    /// end-to-end cases a real fixture would cover come from a real encoder (`cjxl`), and
    /// what is tested here is everything that has to hold for a file that is *not* one:
    /// that this reader answers with nothing rather than with a guess, and never with a
    /// panic.
    fn folder() -> std::path::PathBuf {
        let folder = std::env::temp_dir().join("rust-hover-preview-jxl-tests");
        std::fs::create_dir_all(&folder).expect("a test folder");

        folder
    }

    #[test]
    fn a_file_that_is_not_a_jxl_is_nothing() {
        let folder = folder();

        let png_named_jxl = folder.join("actually-a-png.jxl");
        std::fs::write(
            &png_named_jxl,
            b"\x89PNG\r\n\x1a\n this is not a JPEG XL at all",
        )
        .expect("a written file");

        // Every byte value in turn, which is what a file of no particular kind looks like
        // and what neither signature is a prefix of.
        let noise = folder.join("noise.jxl");
        std::fs::write(&noise, (0u8..=255).cycle().take(4096).collect::<Vec<u8>>())
            .expect("a written file");

        for path in [&png_named_jxl, &noise, &folder.join("missing.jxl")] {
            assert!(!is_animated(path), "{path:?}");
            assert_eq!(decode_frame(path, 0, 64, 64), None, "{path:?}");
            assert_eq!(animation(path), None, "{path:?}");
        }
    }

    #[test]
    fn a_container_that_ends_before_its_codestream_is_nothing() {
        let folder = folder();

        // A container whose `jxlc` box declares sixty-four bytes of codestream and
        // carries none of them: the signature is there, the header is not, and the header
        // is the whole of what this reader reads a `.jxl` for.
        let mut bytes = CONTAINER_SIGNATURE.to_vec();
        bytes.extend_from_slice(&[0x00, 0x00, 0x00, 0x40, b'j', b'x', b'l', b'c']);
        let truncated = folder.join("truncated-container.jxl");
        std::fs::write(&truncated, bytes).expect("a written file");

        assert!(!is_animated(&truncated));
        assert_eq!(decode_frame(&truncated, 0, 64, 64), None);
        assert_eq!(animation(&truncated), None);
    }

    #[test]
    fn a_codestream_that_stops_after_its_header_plays_nothing() {
        let folder = folder();

        // A bare codestream, cut off just after where its image header is. What is
        // asserted is what a file like this must never be: a sequence in time, or a frame
        // out of it. Its header may or may not be one a decoder accepts — the header is a
        // few bits long — so the size is not asked about.
        let mut bytes = CODESTREAM_SIGNATURE.to_vec();
        bytes.extend_from_slice(&[0u8; 24]);
        let truncated = folder.join("truncated-codestream.jxl");
        std::fs::write(&truncated, bytes).expect("a written file");

        assert!(!is_animated(&truncated));
        assert_eq!(decode_frame(&truncated, 0, 64, 64), None);
        assert_eq!(animation(&truncated), None);
    }

    /// The conversion a future reader will get wrong if this module does not say so
    /// clearly, checked against the two timebases a file is written with in practice.
    #[test]
    fn ticks_become_the_milliseconds_they_are() {
        // A thousand ticks a second, the usual timebase: a hundred ticks is a tenth of a
        // second, and it is the number a millisecond answer has to be, not the ticks.
        assert_eq!(ticks_to_ms(100, 1000, 1), 100);
        assert_eq!(ticks_to_ms(1, 1000, 1), 1);
        assert_eq!(ticks_to_ms(1500, 1000, 1), 1500);

        // A frame's duration is a *timecode* rather than a length, which is the other way
        // a file may be written: a frame starting at 250 ticks past the second is shown
        // for half a second, not for 250 of them.
        assert_eq!(ticks_to_ms(500, 100, 1), 5000);

        // One tick a second, and the whole of a second for one tick.
        assert_eq!(ticks_to_ms(1, 1, 1), 1000);

        // A still has no timebase at all, and a file that claims one of zero is not
        // divided by it.
        assert_eq!(ticks_to_ms(100, 0, 1), 0);

        // A file with a very short tick and a long frame is held to what the engine can
        // carry rather than wrapping into a delay of no time at all.
        assert_eq!(ticks_to_ms(u32::MAX, 1, 1), u32::MAX);
    }
}

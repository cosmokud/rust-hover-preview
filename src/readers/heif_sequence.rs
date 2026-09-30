//! The pictures that move inside a HEIF or AVIF container, decoded through the media
//! engine's *decoder* rather than its player.
//!
//! An AVIF sequence (`avis` brand) and a HEIC/HEIF sequence (`mif1` sample group) are
//! ISOBMFF containers, and this asks Windows for their frames one at a time as an
//! `IMFSourceReader`, each sample pushed into the same `Vec<ImageFrame>` queue an animated
//! GIF feeds — the drawing code must not know the difference, which is why the app's own
//! video path is deliberately not used here. **Microsoft has never documented that Media
//! Foundation can open either brand as a source, so `None` is the designed answer rather
//! than a bug to be chased**: the caller falls back to the still path, which reads the same
//! file through WIC. And the startup the media engine needs is not done here —
//! `codecs::mf_started()` starts it once for the process, and asking it also puts the
//! calling thread in an apartment, which is what every COM object below needs.

use crate::formats::codecs;
use crate::readers::video_player;
use std::path::Path;
use std::sync::atomic::{AtomicBool, Ordering};
use windows::Win32::Media::MediaFoundation::{
    IMFAttributes, IMFMediaBuffer, IMFMediaType, IMFSample, IMFSourceReader, MFCreateAttributes,
    MFCreateMediaType, MFCreateSourceReaderFromByteStream, MFMediaType_Video, MFVideoFormat_ARGB32,
    MFVideoFormat_RGB32, MF_MT_FRAME_RATE, MF_MT_FRAME_SIZE, MF_MT_MAJOR_TYPE,
    MF_MT_PIXEL_ASPECT_RATIO, MF_MT_SAMPLE_SIZE, MF_MT_SUBTYPE, MF_SOURCE_READERF_ENDOFSTREAM,
    MF_SOURCE_READER_DISCONNECT_MEDIASOURCE_ON_SHUTDOWN, MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING,
    MF_SOURCE_READER_FIRST_VIDEO_STREAM,
};

/// What every frame of a sequence costs, all of them, before this gives up.
///
/// A hover is not a video player: it holds its frames for as long as the pointer rests and
/// then keeps them until the next preview takes their place, so an hours-long sequence
/// decoded whole is a few gigabytes of this app's heap over one hover. The cap is what a
/// preview is worth — tens of megabytes of frames is far past any moving picture a hover
/// shows — and a file past it is answered with `None`, which is the still path, which is
/// also the right preview for a file too long to animate anyway.
const RETAINED_BYTES: usize = 64 * 1024 * 1024;

/// How many frames a sequence may have, whatever their size.
///
/// The byte ceiling above is what bounds ordinary files; this bounds the other kind, a
/// sequence of frames one pixel across that would otherwise be read until the reader ran
/// out of itself. Two thousand frames is more than any hover shows before it has finished
/// opening.
const MAX_FRAMES: usize = 2048;

// The two ends of how long a frame is held, in milliseconds. The floor is the GIF and the
// animated WebP paths' and for the same reason: a container that claims a zero-length or
// sub-millisecond frame would otherwise spin the render loop at a rate a preview is never
// meant to draw at. The ceiling is a hover's — a file whose timing says one frame lasts a
// minute is not animating, it is a broken timestamp, and a real frame that is merely slow (a
// slideshow's two seconds) is under it.
const MIN_FRAME_DELAY_MS: u32 = 33;
const MAX_FRAME_DELAY_MS: u32 = 1000;

/// A sequence decoded whole: its frames, and how long each is held.
///
/// The delay is beside the frame rather than in a list beside the list, because a delay
/// index that does not line up with a frame index is a delay shown against the wrong
/// picture, and the pairing is what the caller actually wants — one `ImageFrame` per entry.
pub(crate) struct Sequence {
    /// The size every frame is, which is the size the file's own track was asked for.
    pub(crate) width: u32,
    pub(crate) height: u32,
    /// The frames in presentation order: BGRA pixels, and the hold before the next frame.
    pub(crate) frames: Vec<(Vec<u8>, u32)>,
}

/// Decode the sequence into frames, or `None`.
///
/// The picture's size is read from the track's own media type as part of the decode, so
/// there is no separate measuring call here: a HEIF sequence this engine will not open has
/// no size either, and a still HEIC — which has a size — is measured and read by
/// `wic_image` rather than here.
///
/// `max_width`/`max_height` are the box the preview is laid out in: the decoder is asked
/// for the picture at the size that fits inside it, so the engine does the scaling and a
/// forty-megapixel sequence is not decoded whole to be shrunk afterwards. `cancel` is
/// checked between samples, so a hover that moves on stops here rather than finishing a
/// decode nobody is waiting for.
///
/// `None` is every failure — a file that will not open, a track the engine cannot decode,
/// a box of nothing, a picture past the byte ceiling — and it is also a file with fewer
/// than two frames: a single-frame file is a *still*, and a still is drawn by the
/// still path, which reads more formats and costs a fraction of this. Returning it as a
/// one-frame sequence here would be a preview that cannot move, decoded by the slower of
/// the two readers that can produce it.
pub(crate) fn decode_sequence(
    path: &Path,
    max_width: u32,
    max_height: u32,
    cancel: &AtomicBool,
) -> Option<Sequence> {
    if max_width == 0 || max_height == 0 || !codecs::mf_started() {
        return None;
    }
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    let byte_stream = video_player::open_stream(path)?;
    let attributes = reader_attributes()?;
    let reader =
        unsafe { MFCreateSourceReaderFromByteStream(&byte_stream, Some(&attributes)) }.ok()?;

    // The track's own media type: the size it says it has, and — for the timing below —
    // the rate it says it runs at.
    let native = unsafe { reader.GetNativeMediaType(first_video_stream(), 0) }.ok()?;
    let native_size = frame_size(&native)?;
    let fallback_delay_ms = frame_rate_ms(&native);

    // The box the preview was planned at, and then whatever the engine actually agreed to
    // produce — which is not always what was asked for.
    let wanted = fit_within(native_size, (max_width, max_height));
    let (width, height, force_opaque) = negotiated_size(&reader, native_size, wanted)?;

    let bytes_per_frame = crate::config::config::frame_bytes_within_budget(width, height, 4)?;
    let mut pixels: Vec<Vec<u8>> = Vec::new();
    let mut times: Vec<i64> = Vec::new();
    let mut durations: Vec<i64> = Vec::new();
    let mut retained = 0usize;

    loop {
        if cancel.load(Ordering::Acquire) {
            return None;
        }

        let mut stream_flags = 0u32;
        let mut timestamp = 0i64;
        let mut sample: Option<IMFSample> = None;

        let read = unsafe {
            reader.ReadSample(
                first_video_stream(),
                0,
                None,
                Some(&mut stream_flags),
                Some(&mut timestamp),
                Some(&mut sample),
            )
        };

        // The end of a stream is not a failure: the reader answers the last read with a
        // success, no sample, and this flag, which is where the loop below is meant to
        // stop rather than a `None` that is read as a decode that failed.
        if read.is_err() || stream_flags & MF_SOURCE_READERF_ENDOFSTREAM.0 as u32 != 0 {
            break;
        }

        let Some(sample) = sample else {
            break;
        };

        // A sample's own presentation time, in 100-nanosecond units, and the length of the
        // sample itself. The sample's own time is preferred over the one the reader hands
        // back beside it, and the reader's is the fallback for a sample that carries none.
        let time = match unsafe { sample.GetSampleTime() } {
            Ok(sample_time) if sample_time != 0 => sample_time,
            _ => timestamp,
        };
        let duration = unsafe { sample.GetSampleDuration() }.unwrap_or(0);

        let mut frame = sample_pixels(&sample, width, bytes_per_frame)?;

        // What came back is only usable if its alpha byte means something. `ARGB32` has a
        // defined one; `RGB32` does not, and a frame whose alpha is undefined is a frame
        // that may compose to nothing at all, so it is made opaque here rather than left
        // to the drawing code to guess at.
        if force_opaque {
            video_player::force_opaque(&mut frame);
        }

        retained = retained.saturating_add(frame.len());
        if retained > RETAINED_BYTES {
            return None;
        }

        pixels.push(frame);
        times.push(time);
        durations.push(duration);

        if pixels.len() >= MAX_FRAMES {
            break;
        }
    }

    // A still is not a sequence, and neither is a file that decoded nothing at all.
    if pixels.len() < 2 {
        return None;
    }

    let delays = frame_delays_ms(&times, &durations, fallback_delay_ms);
    let frames = pixels.into_iter().zip(delays).collect();

    Some(Sequence {
        width,
        height,
        frames,
    })
}

/// How long each frame is held, in milliseconds.
///
/// Every time in this API is in **100-nanosecond units**, so a span is divided by 10,000
/// to be milliseconds — dividing by 1,000 (as for microseconds) makes every animation
/// run ten times too fast, and not dividing at all leaves a number that is meaningless to
/// the caller. The presentation times are **absolute**, so a frame's hold is the
/// *difference* to the next one, not the next one's own time: taking the times as
/// durations makes the first frame instant and every frame after it longer than the file.
///
/// Three cases, in order of preference:
///
///  * a frame with a next one is held until that one's time arrives;
///  * the last frame has no next one, so it is held for its own `MF_PT_DURATION`;
///  * a span that comes out zero, negative, or absent — two samples with the same
///    timestamp, a container that did not write the timing — falls back to the track's
///    frame rate, and to the shared floor where the track did not say that either.
///
/// What comes out is clamped to the two bounds above, which are the render loop's and a
/// hover's.
fn frame_delays_ms(times: &[i64], durations: &[i64], fallback_ms: Option<u32>) -> Vec<u32> {
    let fallback = fallback_ms
        .filter(|ms| *ms > 0)
        .unwrap_or(MIN_FRAME_DELAY_MS)
        .clamp(MIN_FRAME_DELAY_MS, MAX_FRAME_DELAY_MS);

    times
        .iter()
        .enumerate()
        .map(|(index, time)| {
            let span = match times.get(index + 1) {
                Some(next) => next.saturating_sub(*time),
                None => durations.get(index).copied().unwrap_or(0),
            };

            // A stream whose samples do not come out in order gives a negative span, and
            // one that repeats a timestamp gives a zero. Neither is a length, so both take
            // the fallback rather than dividing towards zero into another zero.
            if span <= 0 {
                return fallback;
            }

            // A span that overflows a `u32` of milliseconds is a file claiming a frame
            // lasts over forty-nine days; the ceiling is what it is taken as.
            let milliseconds = u32::try_from(span / 10_000).unwrap_or(MAX_FRAME_DELAY_MS);

            milliseconds.clamp(MIN_FRAME_DELAY_MS, MAX_FRAME_DELAY_MS)
        })
        .collect()
}

/// The track's frame rate as the hold of one frame, in milliseconds.
///
/// This is the fallback for a file whose samples carry no timing of their own — a
/// constant-rate sequence written by a camera, where the rate is the only timing there is.
fn frame_rate_ms(media_type: &IMFMediaType) -> Option<u32> {
    let packed = unsafe { media_type.GetUINT64(&MF_MT_FRAME_RATE) }.ok()?;

    let (numerator, denominator) = video_player::unpack_pair(packed);

    milliseconds_per_frame(numerator, denominator)
}

/// One frame's hold, in milliseconds, from a rate written as a fraction of frames a
/// second.
///
/// A rate of `numerator / denominator` frames a second means one frame lasts
/// `denominator / numerator` seconds, so the hold in milliseconds is
/// `denominator * 1000 / numerator` — note which half is the numerator and which the
/// denominator here, which is the opposite of the order the arguments are given in. A
/// 30/1 rate is thirty-three milliseconds, and a 30000/1001 rate is the same to within a
/// millisecond. Both halves are widened before the division, since a numerator of
/// `u32::MAX` denominated into it is a rate no animation has and must not be the one that
/// overflows.
fn milliseconds_per_frame(numerator: u32, denominator: u32) -> Option<u32> {
    if numerator == 0 || denominator == 0 {
        return None;
    }

    let milliseconds = (u64::from(denominator) * 1_000) / u64::from(numerator);

    u32::try_from(milliseconds).ok().filter(|ms| *ms > 0)
}

/// A media type's frame size, out of the packed `UINT64` `MF_MT_FRAME_SIZE` is: the high
/// half is the width and the low one the height, which is the reverse of the way a
/// `(u32, u32)` reads (see `video_player::unpack_pair`).
fn frame_size(media_type: &IMFMediaType) -> Option<(u32, u32)> {
    let packed = unsafe { media_type.GetUINT64(&MF_MT_FRAME_SIZE) }.ok()?;
    let (width, height) = video_player::unpack_pair(packed);

    (width > 0 && height > 0).then_some((width, height))
}

/// The largest shape of the picture that fits inside the box, at the picture's own ratio.
///
/// A sequence is never scaled *up*: a box larger than the picture is the pinned-window
/// case, which the drawing side already handles for every other format, and asking a
/// decoder for more pixels than the file has costs a frame that is blurrier than the one
/// it replaced. Both halves are taken in `u64` so a shape that overflows a `u32` cannot
/// wrap into a size that looks small enough to keep.
fn fit_within(size: (u32, u32), bounds: (u32, u32)) -> (u32, u32) {
    let (source_width, source_height) = (u64::from(size.0), u64::from(size.1));
    let (max_width, max_height) = (u64::from(bounds.0), u64::from(bounds.1));

    let width = source_width.min(max_width);
    let height_at_that_width = (source_height * width / source_width).max(1);

    let (width, height) = if height_at_that_width <= max_height {
        (width, height_at_that_width)
    } else {
        (
            (source_width * max_height / source_height).max(1),
            max_height,
        )
    };

    (
        width.min(u64::from(u32::MAX)) as u32,
        height.min(u64::from(u32::MAX)) as u32,
    )
}

/// The size the reader ended up producing, and whether its frames' alpha has to be
/// forced.
///
/// The wanted size is asked for first and the track's own size is the second try: a
/// decoder that will not scale generally says so, and taking what it will give is better
/// than answering a hover with no preview at all when the caller had asked for a smaller
/// picture than the file has. What came back is read from the reader rather than assumed,
/// because a video processor is entitled to hand back something else — a stride it chose,
/// an interlaced picture, a size rounded to a macroblock — and a frame copied at the
/// wrong shape is a frame of garbage rather than a wrong preview.
///
/// The third value is what the copy above then depends on, and it is read here rather
/// than remembered from the ask: the ask is a preference the reader may decline, so which
/// of the two formats the frames are actually in is only knowable from what it settled on.
fn negotiated_size(
    reader: &IMFSourceReader,
    native: (u32, u32),
    wanted: (u32, u32),
) -> Option<(u32, u32, bool)> {
    if !set_output_type(reader, wanted) && !set_output_type(reader, native) {
        return None;
    }

    let current = unsafe { reader.GetCurrentMediaType(first_video_stream()) }.ok()?;
    let size = frame_size(&current)?;

    let subtype = unsafe { current.GetGUID(&MF_MT_SUBTYPE) }.ok()?;

    Some((size.0, size.1, subtype == MFVideoFormat_RGB32))
}

/// Ask the reader for decoded frames at `size`, in the format the preview is composed in.
///
/// The video processing attribute is what makes this answer at all: without it the reader
/// is asked whether the *decoder* delivers RGB32, which no HEVC or AV1 decoder does, and
/// every sequence on the machine would be answered with no.
fn set_output_type(reader: &IMFSourceReader, size: (u32, u32)) -> bool {
    // `ARGB32` first, for the alpha: it is the one of the two whose alpha byte is defined,
    // and a frame composed over a backdrop with an undefined alpha composes to nothing.
    // What it holds is premultiplied, which is identical to straight for every opaque
    // frame — which is nearly every frame of a sequence — and closer to right for the
    // translucent ones than an undefined byte is. `RGB32` is kept as the second try for
    // the decoder that will only give that one.
    [MFVideoFormat_ARGB32, MFVideoFormat_RGB32]
        .into_iter()
        .any(|subtype| {
            let Ok(frames) = (unsafe { MFCreateMediaType() }) else {
                return false;
            };

            unsafe { frames.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Video) }.is_err()
                || unsafe { frames.SetGUID(&MF_MT_SUBTYPE, &subtype) }.is_err()
                || unsafe {
                    frames.SetUINT64(&MF_MT_FRAME_SIZE, (u64::from(size.0) << 32) | u64::from(size.1))
                }
                .is_err()
                || unsafe { frames.SetUINT32(&MF_MT_SAMPLE_SIZE, size.0.saturating_mul(4)) }
                    .is_err()
                // A square pixel is asked for explicitly: without it the processor may
                // carry the container's own aspect ratio into a picture that is then
                // stretched, since the frames are copied at the size above and nowhere
                // else is the ratio applied.
                || unsafe {
                    frames.SetUINT64(&MF_MT_PIXEL_ASPECT_RATIO, (1u64 << 32) | 1)
                }
                .is_err()
                || unsafe {
                    reader.SetCurrentMediaType(first_video_stream(), None, &frames)
                }
                .is_err()
        })
}

/// The attributes a reader for a hover is made with.
///
/// The two are both about not holding on to the file: video processing is what turns what
/// the decoder hands back into something composable, and disconnecting the media source
/// on shutdown is what lets the reader let go of the byte stream — and the file handle
/// under it — the moment this module drops the reader, rather than when the process ends.
/// A hover that walks a folder of sequences would otherwise hold every one of them open.
fn reader_attributes() -> Option<IMFAttributes> {
    let mut attributes: Option<IMFAttributes> = None;
    if unsafe { MFCreateAttributes(&mut attributes, 2) }.is_err() {
        return None;
    }

    let attributes = attributes?;
    if unsafe { attributes.SetUINT32(&MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING, 1) }.is_err() {
        return None;
    }
    if unsafe { attributes.SetUINT32(&MF_SOURCE_READER_DISCONNECT_MEDIASOURCE_ON_SHUTDOWN, 1) }
        .is_err()
    {
        return None;
    }

    Some(attributes)
}

/// One sample's pixels, copied out row by row and the sample let go of.
///
/// The sample is turned into a single contiguous buffer first, because a sample may hold
/// several and only the first `bytes_per_frame` of the first would be the picture. It is
/// then copied *a row at a time* rather than as one block, because a media buffer's row
/// is not required to be exactly `width * 4` bytes: a decoder that pads each row out to a
/// multiple of 16 hands back a buffer that is longer than the picture, and copying it as
/// one block shifts every row after the first by the padding and comes out as a picture
/// sheared diagonally.
///
/// The buffer is unlocked before this returns whatever happens, including on the paths
/// that return `None` — a locked media buffer belongs to the reader's pool and does not
/// come back, so a leak here is a decoder that stops handing out frames after a few
/// hundred.
fn sample_pixels(sample: &IMFSample, width: u32, bytes_per_frame: usize) -> Option<Vec<u8>> {
    let row_bytes = (width as usize).checked_mul(4)?;
    let buffer: IMFMediaBuffer = unsafe { sample.ConvertToContiguousBuffer() }.ok()?;

    let mut data: *mut u8 = std::ptr::null_mut();
    let mut length = 0u32;
    if unsafe { buffer.Lock(&mut data, None, Some(&mut length)) }.is_err() {
        return None;
    }

    let mut pixels = Vec::new();

    // A buffer that is not the whole picture is a decode that came out short, and a short
    // frame handed to the queue is a frame drawn from whatever followed it. Nothing is
    // worth guessing at here, so a buffer that cannot fill the frame is no frame at all.
    if !data.is_null() && (length as usize) >= bytes_per_frame {
        // SAFETY: `data` is the pointer the buffer just locked, it is not null, and it is
        // known to describe at least `length` bytes of a live frame. The lock is released by
        // `UnlockBuffer` on the path out, below.
        let available = unsafe { std::slice::from_raw_parts(data, length as usize) };

        pixels.reserve(bytes_per_frame);
        for row in 0..(bytes_per_frame / row_bytes) {
            let start = row * row_bytes;
            // Unreachable while the buffer is known to be the whole frame, but the slice is
            // indexed rather than sliced so that a reader cannot have it both ways.
            let Some(row_pixels) = available.get(start..start + row_bytes) else {
                break;
            };
            pixels.extend_from_slice(row_pixels);
        }
    }

    unsafe {
        let _ = buffer.Unlock();
    }

    (pixels.len() == bytes_per_frame).then_some(pixels)
}

/// The stream a video is asked for, which is the first one the container has.
fn first_video_stream() -> u32 {
    MF_SOURCE_READER_FIRST_VIDEO_STREAM.0 as u32
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The folder every fixture here is written into, named after the module so a folder
    /// of failed tests is recognisable next to the others.
    fn folder() -> std::path::PathBuf {
        let folder = std::env::temp_dir().join("rust-hover-preview-heif-sequence-tests");
        std::fs::create_dir_all(&folder).expect("a test folder");
        folder
    }

    /// A file of `bytes` written under `name`, returned by path.
    fn write(name: &str, bytes: &[u8]) -> std::path::PathBuf {
        let path = folder().join(name);
        std::fs::write(&path, bytes).expect("a written file");
        path
    }

    /// What nothing at all decodes to: a file that is not a picture, and a file that is
    /// not there. There is no real sequence fixture here, and that is deliberate: **a fixture
    /// that decodes on the machine that made it is a fixture that says nothing about any other
    /// machine**, since whether Windows opens a `mif1` or `avis` sequence at all is a property
    /// of that machine's Media Foundation rather than of this code.
    #[test]
    fn a_file_that_is_not_a_sequence_is_answered_with_nothing() {
        let not_a_container = write("not-a-container.avif", &[0x8f; 512]);
        let missing = folder().join("missing.avif");

        let cancel = AtomicBool::new(false);
        // `is_none` rather than an equality: what a caller checks is whether there is a
        // sequence, and there is nothing about a `Sequence` worth comparing.
        assert!(
            decode_sequence(&missing, 256, 256, &cancel).is_none(),
            "a file that is not there is not a sequence"
        );
        assert!(
            decode_sequence(&not_a_container, 256, 256, &cancel).is_none(),
            "random bytes are not a container"
        );
    }

    /// A container that starts correctly and then says nothing sensible — a real `ftyp` box,
    /// even the `avis` brand a sequence declares, and no track behind it, which is the shape of
    /// input most likely to reach this module in the wild. Every one of these is `None`.
    #[test]
    fn a_container_that_stops_after_its_header_is_answered_with_nothing() {
        let mut truncated_ftyp = Vec::new();
        truncated_ftyp.extend_from_slice(&[0x00, 0x00, 0x00, 0x18]); // box size
        truncated_ftyp.extend_from_slice(b"ftyp");
        truncated_ftyp.extend_from_slice(b"avif"); // major brand, truncated
        truncated_ftyp.extend_from_slice(&[0x00, 0x00, 0x00, 0x00]); // minor version
        truncated_ftyp.extend_from_slice(&[0xff, 0xff, 0xff, 0xff]); // claimed box size

        let truncated = write("truncated.avif", &truncated_ftyp);

        // Structurally shaped, semantically empty: a declared `mif1` item whose payload is
        // shorter than its own header.
        let mut empty_item = Vec::new();
        empty_item.extend_from_slice(&[0x00, 0x00, 0x00, 0x10]); // box size
        empty_item.extend_from_slice(b"ftyp");
        empty_item.extend_from_slice(b"mif1");
        empty_item.extend_from_slice(&[0x00, 0x00, 0x00, 0x00]);
        empty_item.extend_from_slice(&[0x00, 0x00, 0x00, 0x08]); // an item box
        empty_item.extend_from_slice(b"mdat");

        let itemless = write("itemless.heif", &empty_item);

        for path in [truncated, itemless] {
            let cancel = AtomicBool::new(false);
            assert!(decode_sequence(&path, 256, 256, &cancel).is_none());
        }
    }

    /// A box of nothing at all is refused before anything is opened.
    #[test]
    fn a_box_of_nothing_is_refused() {
        let path = write("refused.avif", &[0u8; 32]);
        let cancel = AtomicBool::new(false);

        assert!(decode_sequence(&path, 0, 256, &cancel).is_none());
        assert!(decode_sequence(&path, 256, 0, &cancel).is_none());
    }

    /// A cancel that is already set stops the decode before the file is opened at all,
    /// which is what a hover that has already moved on needs.
    #[test]
    fn a_cancelled_hover_decodes_nothing() {
        let path = write("cancelled.avif", &[0u8; 64]);

        assert!(decode_sequence(&path, 256, 256, &AtomicBool::new(true)).is_none());
    }

    /// A frame's hold is the gap to the *next* frame, not the next frame's own time, and
    /// the last frame is held for its own duration.
    ///
    /// The times are the absolute presentation times in hundred-nanosecond units that
    /// Media Foundation hands over, so a file shown at ten frames a second carries times
    /// 0, 1,000,000, 2,000,000 … and every one of them is a moment, not a length.
    #[test]
    fn delays_are_the_gaps_between_presentation_times() {
        // Ten frames a second, expressed as absolute timestamps, with the last frame
        // declaring its own length rather than having a next one to measure to.
        let times = [0i64, 1_000_000, 2_000_000];
        let durations = [0i64, 0, 500_000];

        assert_eq!(
            frame_delays_ms(&times, &durations, None),
            vec![100, 100, 50],
            "a hundred-nanosecond span of a million is a hundred milliseconds"
        );
    }

    /// Timing that says nothing falls back to the frame rate, and to the floor where even
    /// that is nothing — a sequence of identical timestamps must not spin the render loop.
    #[test]
    fn timing_that_says_nothing_falls_back_to_the_frame_rate() {
        // Every sample at the same instant, which is what a container with no per-sample
        // timing produces.
        let times = [0i64, 0, 0];
        let durations = [0i64, 0, 0];

        assert_eq!(
            frame_delays_ms(&times, &durations, Some(40)),
            vec![40, 40, 40],
            "the track's own rate is the only timing there is"
        );
        assert_eq!(
            frame_delays_ms(&times, &durations, None),
            vec![MIN_FRAME_DELAY_MS; 3],
            "and the floor is what is left"
        );
    }

    /// A frame rate packed the way `MF_MT_FRAME_RATE` is packed: 30/1 and 30000/1001.
    ///
    /// The unpacking is tested through `milliseconds_per_frame`, the half of
    /// `frame_rate_ms` that does the arithmetic — asking a real media type for this would
    /// make the test a test of whether the machine has a codec, which is the one thing
    /// this module cannot promise.
    #[test]
    fn a_frame_rate_is_read_out_of_the_packed_rational() {
        // A whole number of frames a second, and the fraction one that a film is written
        // at. The second is the case that catches a reader dividing by the wrong half.
        assert_eq!(milliseconds_per_frame(30, 1), Some(33));
        assert_eq!(
            milliseconds_per_frame(30_000, 1_001),
            Some(33),
            "1001/30 frames a second is still about thirty-three milliseconds"
        );
        assert_eq!(milliseconds_per_frame(0, 1), None, "no rate at all");
        assert_eq!(milliseconds_per_frame(30, 0), None, "a rate over nothing");
    }
}

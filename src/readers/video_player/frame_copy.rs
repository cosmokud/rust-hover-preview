//! The file as the media stack reads it and the picture as the compositor composes it: the
//! stream, what a probe finds in it, and the surface a frame is copied or resampled out of.
//! See the module above for the reasoning.

use crate::formats::codecs;
use crate::readers::audio_track;
use std::os::windows::ffi::OsStrExt;
use std::path::Path;
use windows::core::{Interface, GUID, PCWSTR};
use windows::Win32::Graphics::Imaging::{
    GUID_WICPixelFormat32bppBGRA, IWICBitmap, IWICBitmapLock, WICBitmapCacheOnLoad,
    WICBitmapLockWrite,
};
use windows::Win32::Media::MediaFoundation::{
    IMFAttributes, IMFByteStream, IMFSourceReader, MFAudioFormat_AAC, MFAudioFormat_ADTS,
    MFAudioFormat_ALAC, MFAudioFormat_AMR_NB, MFAudioFormat_AMR_WB, MFAudioFormat_DTS,
    MFAudioFormat_Dolby_AC3, MFAudioFormat_Dolby_DDPlus, MFAudioFormat_FLAC, MFAudioFormat_Float,
    MFAudioFormat_MP3, MFAudioFormat_Opus, MFAudioFormat_PCM, MFAudioFormat_Vorbis,
    MFAudioFormat_WMAudioV8, MFAudioFormat_WMAudioV9, MFAudioFormat_WMAudio_Lossless,
    MFCreateMFByteStreamOnStream, MFCreateMediaType, MFCreateSourceReaderFromByteStream,
    MFMediaType_Audio, MF_BYTESTREAM_ORIGIN_NAME, MF_MT_AUDIO_NUM_CHANNELS,
    MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_AVG_BITRATE, MF_MT_MAJOR_TYPE, MF_MT_SUBTYPE,
    MF_PD_DURATION, MF_SOURCE_READER_FIRST_AUDIO_STREAM, MF_SOURCE_READER_MEDIASOURCE,
};
use windows::Win32::System::Com::{IStream, STGM_READ, STGM_SHARE_DENY_NONE};
use windows::Win32::UI::Shell::SHCreateStreamOnFileEx;

/// A surface for the engine to deliver into: a bitmap in the format a frame is composed
/// in, made at the size the preview came out at.
pub(super) fn surface(width: u32, height: u32) -> Option<IWICBitmap> {
    let factory = crate::readers::wic_image::factory()?;

    unsafe {
        factory
            .CreateBitmap(
                width,
                height,
                &GUID_WICPixelFormat32bppBGRA,
                WICBitmapCacheOnLoad,
            )
            .ok()
    }
}

/// The file as the media stack's own stream, which is the form every reader here is handed and
/// is shared with `heif_sequence` (see there).
///
/// The path goes in as it is — verbatim prefix and all — because `SHCreateStreamOnFileEx`
/// is handed a path rather than being asked to resolve a URL, which is what
/// `pdf_preview::open_document` settled for the same reason. The name is set on the stream
/// as well: the handler that opens a byte stream is chosen by the name it carries, and a
/// stream made from a file has none.
pub(crate) fn open_stream(path: &Path) -> Option<IMFByteStream> {
    let wide: Vec<u16> = path
        .as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect();

    let file: IStream = unsafe {
        SHCreateStreamOnFileEx(
            PCWSTR(wide.as_ptr()),
            STGM_READ.0 | STGM_SHARE_DENY_NONE.0,
            0,
            false,
            None::<&IStream>,
        )
    }
    .ok()?;

    let stream = unsafe { MFCreateMFByteStreamOnStream(&file) }.ok()?;
    let attributes: IMFAttributes = stream.cast().ok()?;
    unsafe { attributes.SetString(&MF_BYTESTREAM_ORIGIN_NAME, PCWSTR(wide.as_ptr())) }.ok()?;

    Some(stream)
}

/// The alpha byte a frame of this app's is composed with, which is what every frame is
/// written with whether the codec behind it had an opinion or not.
///
/// It is a named constant rather than the literal at each of the four places that write one,
/// because the four are the same decision: the samplers below write it because their
/// arithmetic has no alpha channel to carry, the copy writes it because the converter that
/// produced the frame has no reason to have written one either (see [`force_opaque`]), and
/// `heif_sequence` writes it because the format it read the frame in does not have the byte
/// at all.
const OPAQUE: u8 = 255;

/// A frame's alpha bytes forced opaque, which is what every frame of this app's is composed with
/// and is shared with `heif_sequence`.
///
/// The fourth byte of a pixel is set rather than taken from whatever produced the frame, because
/// a picture has no transparency here: the window a video used to be played in was opaque, so a
/// frame that carried an alpha of nothing composites to a preview that fades out or opens blank.
/// What the engine hands over in `RGB32` has no alpha written into it at all (see
/// `heif_sequence::set_output_type`).
///
/// That last point was settled by measurement for `ARGB32` as well rather than assumed from the
/// format's name, since the natural question is whether the colour converter writes `0xFF` into
/// a destination that has an alpha channel and so makes all of this unnecessary. It does not: a
/// 1920 x 800 frame read straight out of the engine's own surface after a transfer into
/// `MFVideoFormat_ARGB32` carried a histogram of the fourth byte over the whole surface of
/// `253`, `254` and `255` — which is not opaque, is not stable, and is not a value anybody
/// could have drawn with — and the very first frame of a session, before the converter had
/// written anything at all into the surface it had just been handed, was `0` throughout. So the
/// byte is written here and this function is not merely a convenience for the one caller who
/// needs a whole frame of it forced: it is the reason a video preview is opaque at all.
///
/// It is no longer on the video path itself, which sets the byte as it copies rather than
/// walking the frame afterwards, because a pass of its own over thirty megabytes is a pass of
/// its own over thirty megabytes (see [`copy_locked`]). What is left here is `heif_sequence`,
/// which has no copy of its own to fold it into.
pub(crate) fn force_opaque(pixels: &mut [u8]) {
    // `as_chunks_mut` rather than `chunks_exact_mut`: a frame is a few million bytes and this
    // is a per-pixel pass over every one of them, so the bounds check the other form carries is
    // worth taking out. A frame that is not a whole number of pixels is left alone rather than
    // shortened, since it is not a frame.
    for pixel in pixels.as_chunks_mut::<4>().0 {
        pixel[3] = OPAQUE;
    }
}

/// A pair packed the way Media Foundation packs one into a `UINT64` — `MF_MT_FRAME_SIZE`'s width
/// over height, a pixel aspect ratio's numerator over denominator, `MF_MT_FRAME_RATE`'s numerator
/// over denominator — the first of the two in the high half and the second in the low.
///
/// It is written as the unpacking it is rather than as arithmetic because it is the reverse of
/// the way a `(u32, u32)` reads, and the mistake is a file that opens and then decodes into a
/// transposed frame.
pub(crate) fn unpack_pair(packed: u64) -> (u32, u32) {
    ((packed >> 32) as u32, packed as u32)
}

/// The surface locked for writing, with the bytes and the row stride the lock reports: the
/// prologue [`copy_locked`] and [`resample_locked`] both open with.
///
/// The lock is handed back rather than let go of here, because what the bytes are read through
/// borrows it and a picture is only drawn while it is held.
fn locked_surface(bitmap: &IWICBitmap) -> Option<(IWICBitmapLock, &[u8], usize)> {
    let lock = unsafe { bitmap.Lock(std::ptr::null(), WICBitmapLockWrite.0 as u32) }.ok()?;
    let stride = (unsafe { lock.GetStride() }).ok()? as usize;

    let mut size: u32 = 0;
    let mut data: *mut u8 = std::ptr::null_mut();
    if unsafe { lock.GetDataPointer(&mut size, &mut data) }.is_err() || data.is_null() {
        return None;
    }

    // SAFETY: `data` is the pointer the lock just reported, it is not null, and it describes the
    // `size` bytes the same call reports. The lock is handed back with it and is what lets the
    // surface go when the picture has been drawn.
    Some((
        lock,
        unsafe { std::slice::from_raw_parts(data, size as usize) },
        stride,
    ))
}

/// Take the picture out of a locked bitmap and into the box, scaled by this side: the road every
/// frame of a preview shown above the picture's own size takes (see `scales_here`).
///
/// It is [`copy_locked`]'s other half and is written in the same terms — BGRA both ways, the alpha
/// forced opaque, the caller's buffer rewritten rather than replaced — with the one difference
/// that the picture is read out of the engine's own surface a row at a time, as the rows are
/// needed, rather than copied out whole first: a row of the picture that two rows of the box are
/// mixed from is read twice from the lock and touched once in memory.
pub(super) fn resample_locked(
    bitmap: &IWICBitmap,
    pixels: &mut Vec<u8>,
    picture: (u32, u32),
    box_size: (u32, u32),
    rows: &mut [Vec<u8>; 2],
) -> bool {
    let Some((_lock, source, stride)) = locked_surface(bitmap) else {
        return false;
    };

    // The buffer is the frame's own and is kept between frames, and every byte of it is written by
    // the rows below: what it needs is the room rather than a zeroing, which at the size of a
    // display was a pass of its own over thirty megabytes sixty times a second. The room it needs
    // is the box's, which a buffer still holding a frame of an older box — a pinned window dragged
    // to another size — does not have yet.
    let row_bytes = box_size.0 as usize * 4;
    let Some(needed) = row_bytes.checked_mul(box_size.1 as usize) else {
        return false;
    };
    if pixels.len() != needed {
        pixels.resize(needed, 0);
    }

    scale_rows(source, stride, picture, box_size, rows, pixels)
}

// How many rows of the picture the scaler has read on this thread since the last time it was
// asked, which is the only question a cache of two rows can be judged by.
//
// It is counted rather than reasoned about because the whole of the cache is invisible from
// the picture it draws: a resampler that reads every row twice and one that reads each once
// draw exactly the same frame. It is thread-local because the tests that ask are run beside
// one another and a counter they all wrote to would be counting every test's frame, and it is
// compiled in only under `cfg(test)` because a number only a test reads is not worth an
// increment on the preview thread.
#[cfg(test)]
thread_local! {
    static ROWS_READ: std::cell::Cell<u64> = const { std::cell::Cell::new(0) };
}

// How many source rows a resample read per destination row of the box: the number the cache
// exists to keep at one per source row of the picture, and which no test can get at any other
// way.
#[cfg(test)]
pub(super) fn rows_read_per_destination_row(
    source: &[u8],
    picture: (u32, u32),
    box_size: (u32, u32),
) -> f64 {
    let mut out = vec![0u8; box_size.0 as usize * box_size.1 as usize * 4];
    let mut rows = [Vec::new(), Vec::new()];
    ROWS_READ.with(|read| read.set(0));

    assert!(scale_rows(
        source,
        picture.0 as usize * 4,
        picture,
        box_size,
        &mut rows,
        &mut out
    ));

    ROWS_READ.with(|read| read.replace(0)) as f64 / f64::from(box_size.1)
}

/// One source pixel for one destination pixel, in the 16.16 the two axes are mapped in.
///
/// It is named because a step of exactly this is the one that says the axis needs nothing done
/// to it, which is a different kind of fact from the mapping itself and is not obvious from the
/// arithmetic: every column lands on the column it started on, so the whole of the horizontal
/// half of the scaling is a copy.
pub(super) const A_WHOLE_PIXEL: u64 = 1 << 16;

/// The scaling itself: the picture at its own size read into a buffer of the box's.
///
/// Bilinear and separable — a row of the box is two rows of the picture read across to the box's
/// width and then mixed, and a pixel of that row is two of those columns mixed — which is the
/// reading `preview_window::resample_into_band` gives a picture being dragged to another size, in
/// the same 16.16 arithmetic and the same 256ths. It is the cheapest filter a picture can be
/// shown above its own size at without the edges of the file arriving a whole destination pixel
/// wide, which is what was being looked at.
///
/// The row of the picture a row of the box takes its lower half from is the row the next one takes
/// its upper half from wherever the box is the larger of the two, and reading it across twice is
/// half the scaling's work done twice over, so the two rows a row of the box is mixed from are
/// kept in `rows` between the rows that need them. Both of them are looked for, and neither is
/// read across twice while the buffer that has it is still holding it: at four times the
/// picture's size a pair of rows is mixed into sixteen rows of the box, and the cache as it
/// stood was reading the lower one on all sixteen of them, so it kept the row for the ticks
/// that needed it and paid for the ones that did not. Below the picture's own size the rows no
/// longer come in pairs and every row is read for itself, which is `resample_into_band`'s
/// arrangement for the same reason: a scale down is not what this is here for.
pub(super) fn scale_rows(
    source: &[u8],
    stride: usize,
    picture: (u32, u32),
    box_size: (u32, u32),
    rows: &mut [Vec<u8>; 2],
    out: &mut [u8],
) -> bool {
    let (picture_width, picture_height) = picture;
    let (box_width, box_height) = box_size;

    if picture_width == 0 || picture_height == 0 || box_width == 0 || box_height == 0 {
        return false;
    }

    let row_bytes = box_width as usize * 4;
    let Some(needed) = row_bytes.checked_mul(box_height as usize) else {
        return false;
    };

    // Every row of the picture this reads has to be there, and the last of them is the one that
    // says so: a lock is as long as the surface's own rows, and a picture sized by the file and
    // delivered by the engine is one this can be short of only where the engine wrote less than it
    // said it would — half a frame being no frame.
    let Some(picture_bytes) = stride
        .checked_mul(picture_height as usize - 1)
        .and_then(|rows| rows.checked_add(picture_width as usize * 4))
    else {
        return false;
    };
    if source.len() < picture_bytes || out.len() < needed {
        return false;
    }

    // Where a destination row and column map back to in the picture, in 16.16, a pixel's own half
    // taken off so that a destination pixel stands on the centre of the source pixel it lands in:
    // the mapping `resample_into_band` makes, and the whole of what the loop repeats.
    let step_x = ((picture_width as u64) << 16) / box_width as u64;
    let step_y = ((picture_height as u64) << 16) / box_height as u64;

    for row in rows.iter_mut() {
        row.resize(row_bytes, 0);
    }

    // Which row of the picture each of the two holds, which is what says whether it has to be read
    // across for the row of the box in hand at all.
    let mut held: [Option<u32>; 2] = [None, None];

    for y in 0..box_height as usize {
        let sy = (y as u64 * step_y + step_y / 2).saturating_sub(0x8000);
        let upper_row = (((sy >> 16) as usize).min(picture_height as usize - 1)) as u32;
        let lower_row = (upper_row + 1).min(picture_height - 1);
        let weight = ((sy & 0xFFFF) >> 8) as u32;

        // The row of the box above the one in hand ends where this one starts wherever the picture
        // is being enlarged, so the buffer the last row was mixed from is the one to use here
        // wherever it still holds the row this one needs: only a row neither buffer holds is read
        // across again.
        let upper = match held.iter().position(|row| *row == Some(upper_row)) {
            Some(index) => index,
            None => {
                interpolate_row(
                    source,
                    stride,
                    picture_width,
                    upper_row as usize,
                    step_x,
                    &mut rows[0],
                );
                held[0] = Some(upper_row);

                0
            }
        };

        // The lower half of the pair is asked for the same way and for the same reason, and this
        // is where the cache was not doing its job: a pair of rows mixed into four of the box is
        // read across five times, and one mixed into sixteen is read across seventeen, all of it
        // into a buffer that has held the row since the first of them. Reading it again costs a
        // pass over the whole picture for nothing, which at a large enlargement is most of what
        // the scaling is.
        let lower = if lower_row == upper_row {
            // The last row of the picture is its own lower half — the mapping is clamped to the
            // picture's own rows, and the row below the last one does not exist. It is mixed
            // with itself, which is the row, so it is read once and read for both.
            upper
        } else if let Some(index) = held.iter().position(|row| *row == Some(lower_row)) {
            index
        } else {
            let index = 1 - upper;

            interpolate_row(
                source,
                stride,
                picture_width,
                lower_row as usize,
                step_x,
                &mut rows[index],
            );
            held[index] = Some(lower_row);

            index
        };

        let destination = &mut out[y * row_bytes..(y + 1) * row_bytes];
        blend_rows(&rows[upper], &rows[lower], weight, destination);
    }

    true
}

/// One row of the picture read across to the box's width: the whole of the horizontal half of the
/// scaling, and the reason a row of the picture is read where a row of the box asks for one.
///
/// The alpha is not read at all — a picture is opaque here, so the fourth byte of every pixel
/// written is 255 whatever the codec put beside it (see [`force_opaque`]).
///
/// A box exactly as wide as the picture is the one case where this is not a resample at all, and
/// it is worth the branch because it is the case a box that enlarged one axis only arrives at: a
/// file whose pixels are not square is stretched along one axis at every size, so the other
/// comes out equal, and every row of every frame of it would be read by the loop below, two
/// pixels at a time, mixed by nothing, to arrive at the row it was given. A step of
/// [`A_WHOLE_PIXEL`] says exactly that — one pixel per pixel, with the half-pixel offset the
/// mapping takes leaving every column standing on its own — and the whole of the row comes then
/// from [`copy_row_opaque`], which is what copying a row here already is.
pub(super) fn interpolate_row(
    source: &[u8],
    stride: usize,
    picture_width: u32,
    picture_row: usize,
    step: u64,
    out: &mut [u8],
) {
    #[cfg(test)]
    ROWS_READ.with(|read| read.set(read.get() + 1));

    let start = picture_row * stride;
    let picture_row = &source[start..start + picture_width as usize * 4];

    if step == A_WHOLE_PIXEL {
        copy_row_opaque(picture_row, out);
        return;
    }

    let last = picture_width as usize - 1;

    for (x, pixel) in out.as_chunks_mut::<4>().0.iter_mut().enumerate() {
        let sx = (x as u64 * step + step / 2).saturating_sub(0x8000);
        let left = ((sx >> 16) as usize).min(last);
        let right = (left + 1).min(last);
        let weight = ((sx & 0xFFFF) >> 8) as u32;

        let (near, far) = (left * 4, right * 4);
        for channel in 0..3 {
            let mixed = picture_row[near + channel] as u32 * (256 - weight)
                + picture_row[far + channel] as u32 * weight;

            pixel[channel] = ((mixed + 128) >> 8) as u8;
        }

        pixel[3] = OPAQUE;
    }
}

/// Mix the two rows of the picture a row of the box sits between, by how far into the pair that
/// row falls: the vertical half of the scaling, weighed in the same 256ths [`interpolate_row`]
/// weighs a column by.
fn blend_rows(upper: &[u8], lower: &[u8], weight: u32, out: &mut [u8]) {
    let near_weight = 256 - weight;

    for ((pixel, upper), lower) in out
        .as_chunks_mut::<4>()
        .0
        .iter_mut()
        .zip(upper.as_chunks::<4>().0)
        .zip(lower.as_chunks::<4>().0)
    {
        for channel in 0..3 {
            let mixed = upper[channel] as u32 * near_weight + lower[channel] as u32 * weight;

            pixel[channel] = ((mixed + 128) >> 8) as u8;
        }

        pixel[3] = OPAQUE;
    }
}

/// Copy a locked bitmap out as the preview's frame, which is the one place a video's pixels
/// are touched where the engine is the one that scaled them.
///
/// The copy and the alpha are one pass, which is the whole of what this function is for: a
/// frame at the size of a display is thirty megabytes, so a second pass over it is a second
/// thirty megabytes read back out of main memory and thirty more written over the compositor's
/// copy — a hundred megabytes of traffic for a byte in every fourth position, on the preview
/// thread, once for each frame of the file. Written as two passes it was two traversals of a
/// buffer that does not fit in a cache, and the second one paid for every byte of the first.
/// Written as one there is nothing to pay for at all: the fourth byte is set as the pixel goes
/// past rather than by walking over what was just written.
pub(super) fn copy_locked(
    bitmap: &IWICBitmap,
    pixels: &mut Vec<u8>,
    width: u32,
    height: u32,
) -> bool {
    let Some(stride) = (width as usize).checked_mul(4) else {
        return false;
    };
    let Some(needed) = stride.checked_mul(height as usize) else {
        return false;
    };

    let Some((_lock, source, source_stride)) = locked_surface(bitmap) else {
        return false;
    };
    if source.len() < needed {
        return false;
    }

    // The buffer is the frame's own and is kept between frames, and every byte of it is written
    // by the rows below: what it needs is the room rather than a zeroing, which at the size of a
    // display was a pass of its own over thirty megabytes sixty times a second.
    if pixels.len() != needed {
        pixels.resize(needed, 0);
    }

    for row in 0..height as usize {
        let from = row * source_stride;
        let to = row * stride;

        copy_row_opaque(&source[from..from + stride], &mut pixels[to..to + stride]);
    }

    true
}

/// One row of the engine's surface into one row of the box's, with every pixel of it opaque as
/// it goes past rather than by a walk over the row afterwards.
///
/// The two are separate functions for the reason every other row-level step in this file is
/// one: this is the whole of [`copy_locked`] that can be looked at without a bitmap to lock,
/// and the alpha is the part of it that is easy to get wrong. A pixel past the end of the
/// pair is left as it was, which is the same rule [`force_opaque`] keeps and for the same
/// reason — the two slices are the same length here, so a byte over is a byte over rather than
/// a frame.
pub(super) fn copy_row_opaque(from: &[u8], to: &mut [u8]) {
    // `as_chunks` rather than `chunks_exact`: pairing the two by index costs a bounds check
    // per pixel on the way round a whole frame, and the pairing is the guarantee that each
    // pixel is copied with its own alpha written beside it.
    for (pixel, from) in to
        .as_chunks_mut::<4>()
        .0
        .iter_mut()
        .zip(from.as_chunks::<4>().0)
    {
        pixel.copy_from_slice(&[from[0], from[1], from[2], OPAQUE]);
    }
}

/// Whether this machine can play `path` as a sound, and what the file says it holds.
///
/// It is the one question the sound list cannot answer, asked the way every other question in
/// this app is asked — of the machine rather than of a table. The file is opened as a source
/// reader (the header parse a hover pays for and nothing more), its own media type is read for
/// the facts the card is drawn with, and then the reader is asked for *decoded* PCM on that
/// stream: a source reader can only give what a registered decoder can produce, so a yes there
/// is the whole of "this will play".
///
/// The last fact is the file's own length, read off the reader's presentation — which is the
/// container's answer rather than a decoder's, and is what the probe is asked for beyond the
/// card: where a sound starts is a position or a share of its length, and a share of a length
/// nothing has read yet is not a place at all (see `audio_seek`). It is read from the media
/// source rather than from the media engine because the engine is not started by a probe — and
/// a duration in hundred-nanosecond units, which is the unit the property is written in, is
/// what the seconds the rest of the app speaks in are made of. A container that does not say
/// leaves it out, and the engine's own answer is what the card is drawn from then.
pub fn audio_probe(path: &Path) -> Option<audio_track::Track> {
    if !codecs::mf_started() {
        return None;
    }

    let byte_stream = open_stream(path)?;
    let reader =
        unsafe { MFCreateSourceReaderFromByteStream(&byte_stream, None::<&IMFAttributes>) }.ok()?;

    // A file with no audio stream at all — a film, or a container of something else — is not a
    // sound, and there is nothing to be asked about it beyond that.
    let media_type =
        unsafe { reader.GetNativeMediaType(MF_SOURCE_READER_FIRST_AUDIO_STREAM.0 as u32, 0) }
            .ok()?;

    let subtype = unsafe { media_type.GetGUID(&MF_MT_SUBTYPE) }.ok()?;
    let rate = unsafe { media_type.GetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND) }.ok();
    let channels = unsafe { media_type.GetUINT32(&MF_MT_AUDIO_NUM_CHANNELS) }
        .ok()
        .and_then(|channels| u16::try_from(channels).ok());
    let bitrate = unsafe { media_type.GetUINT32(&MF_MT_AVG_BITRATE) }.ok();
    let duration = presentation_duration(&reader);

    // The decoder, asked for the only way that answers it: a stream the reader will hand back
    // as PCM is a stream this machine has a decoder for.
    let pcm = unsafe { MFCreateMediaType() }.ok()?;
    unsafe { pcm.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio) }.ok()?;
    unsafe { pcm.SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_PCM) }.ok()?;
    unsafe { reader.SetCurrentMediaType(MF_SOURCE_READER_FIRST_AUDIO_STREAM.0 as u32, None, &pcm) }
        .ok()?;

    Some(audio_track::Track {
        player: audio_track::Player::Native,
        codec: codec_name(&subtype),
        rate: rate.filter(|rate| *rate > 0),
        channels: channels.filter(|channels| *channels > 0),
        bitrate: bitrate.filter(|bitrate| *bitrate > 0),
        duration,
    })
}

/// How long the file a reader was opened over plays, in seconds, where the container says.
///
/// The property is the media source's own rather than a stream's, and it is written in
/// hundred-nanosecond units — the unit every duration in this API is kept in — so what is
/// handed back is the number of seconds the rest of the app works in. A file that does not say,
/// and a reader that will not answer, are one answer here: nothing, which is a card drawn
/// without a length and a hover that starts a sound at its beginning rather than at a share of
/// a length nothing knows.
fn presentation_duration(reader: &IMFSourceReader) -> Option<f64> {
    let hundred_nanoseconds = unsafe {
        reader.GetPresentationAttribute(MF_SOURCE_READER_MEDIASOURCE.0 as u32, &MF_PD_DURATION)
    }
    .ok()
    .and_then(|value| u64::try_from(&value).ok())?;

    let seconds = hundred_nanoseconds as f64 / 10_000_000.0;

    (seconds.is_finite() && seconds > 0.0).then_some(seconds)
}

/// What a stream's own codec is called, where this app has a name for it.
///
/// The engine names its formats by the GUID of the stream rather than by a word, and the words
/// beside them are the ones a person reads on a label: what is left unnamed is a codec this
/// app has no name for, which the card answers for with the extension the file carries (see
/// `audio_preview::facts_of`).
fn codec_name(subtype: &GUID) -> Option<String> {
    const CODECS: &[(GUID, &str)] = &[
        (MFAudioFormat_MP3, "MP3"),
        (MFAudioFormat_AAC, "AAC"),
        (MFAudioFormat_ADTS, "AAC"),
        (MFAudioFormat_FLAC, "FLAC"),
        (MFAudioFormat_ALAC, "ALAC"),
        (MFAudioFormat_WMAudioV8, "WMA"),
        (MFAudioFormat_WMAudioV9, "WMA"),
        (MFAudioFormat_WMAudio_Lossless, "WMA Lossless"),
        (MFAudioFormat_Dolby_AC3, "Dolby Digital"),
        (MFAudioFormat_Dolby_DDPlus, "Dolby Digital Plus"),
        (MFAudioFormat_DTS, "DTS"),
        (MFAudioFormat_AMR_NB, "AMR"),
        (MFAudioFormat_AMR_WB, "AMR-WB"),
        (MFAudioFormat_Opus, "Opus"),
        (MFAudioFormat_Vorbis, "Vorbis"),
        (MFAudioFormat_PCM, "PCM"),
        (MFAudioFormat_Float, "PCM"),
    ];

    CODECS
        .iter()
        .find(|(codec, _)| codec == subtype)
        .map(|(_, name)| (*name).to_string())
}

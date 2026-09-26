//! The media engine Windows has, for the previews it plays.
//!
//! A video is played by `ffplay` when FFmpeg is installed, and by this when it is not; a sound
//! is played by this where its own decoders reach the format, and by that player where they do
//! not — the other way round, and deliberately: what the engine gives a sound is a player
//! inside this app's own process, with no window, no process to supervise and a position the
//! card can be drawn from (see `audio_track` and `audio_preview`).
//!
//! The two are the same preview to the rest of the app — the same window, the same
//! placement, the same box — because what comes out of here is frames, drawn where every
//! other frame is drawn, rather than a player's window standing in for the preview. That
//! is the shape the media engine is asked for: it is created in *frame-server* mode,
//! which is what it is by default when no playback window is named, and then
//!
//!   * it decodes, paces itself, plays the audio and loops — none of which this app does;
//!   * this side asks, once a tick, whether a frame is due (`OnVideoStreamTick`) and takes
//!     it (`TransferVideoFrame`) into a bitmap of its own.
//!
//! What that buys over the ffplay path is the whole of the window machinery a player's
//! own window needs — the style monitor, the topmost re-assertion, the PID record, the
//! job object — none of which exists here, because there is no second window and no
//! second process. What it costs is that the frames are copied once per frame.
//!
//! The engine is what decides which files it can play: whatever Windows 11 decodes out of
//! the box, plus whatever a codec extension from the Microsoft Store has added — which is
//! the question the tray's `Codecs` submenu answers for the machine it is running on.
//!
//! The three questions asked of the engine from outside are all answered here. [`dimensions`]
//! is a video's shape, asked of a source reader before there is anything to play. [`audio_track`]
//! is whether this machine has a decoder for a sound and what the file says about it, asked the
//! same way — the reader is asked for *decoded* PCM, which it can only give where the decoder is
//! registered, and the media type the file declares is read for the rest. And [`position`] and
//! [`duration`] are what the running session says about itself, for the card's clock.

use crate::formats::codecs;
use crate::readers::audio_track;
use std::cell::RefCell;
use std::os::windows::ffi::OsStrExt;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Arc;
use windows::core::{implement, GUID, IUnknown, Interface, BSTR, PCWSTR};
use windows::Win32::Foundation::RECT;
use windows::Win32::Graphics::Imaging::{
    GUID_WICPixelFormat32bppBGRA, IWICBitmap, WICBitmapCacheOnLoad, WICBitmapLockWrite,
};
use windows::Win32::Media::MediaFoundation::{
    CLSID_MFMediaEngineClassFactory, IMFAttributes, IMFByteStream, IMFMediaEngine,
    IMFMediaEngineClassFactory, IMFMediaEngineEx, IMFMediaEngineNotify, IMFMediaEngineNotify_Impl,
    MFAudioFormat_AAC, MFAudioFormat_ADTS, MFAudioFormat_ALAC, MFAudioFormat_AMR_NB,
    MFAudioFormat_AMR_WB, MFAudioFormat_DTS, MFAudioFormat_Dolby_AC3, MFAudioFormat_Dolby_DDPlus,
    MFAudioFormat_FLAC, MFAudioFormat_Float, MFAudioFormat_MP3, MFAudioFormat_Opus,
    MFAudioFormat_PCM, MFAudioFormat_Vorbis, MFAudioFormat_WMAudioV8, MFAudioFormat_WMAudioV9,
    MFAudioFormat_WMAudio_Lossless, MFCreateAttributes, MFCreateMFByteStreamOnStream,
    MFCreateMediaType, MFCreateSourceReaderFromByteStream, MFMediaType_Audio, MFVideoFormat_ARGB32,
    MFARGB, MF_BYTESTREAM_ORIGIN_NAME, MF_MEDIA_ENGINE_CALLBACK, MF_MEDIA_ENGINE_EVENT_ERROR,
    MF_MEDIA_ENGINE_READY_HAVE_CURRENT_DATA, MF_MEDIA_ENGINE_VIDEO_OUTPUT_FORMAT,
    MF_MT_AUDIO_NUM_CHANNELS, MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_AVG_BITRATE, MF_MT_FRAME_SIZE,
    MF_MT_MAJOR_TYPE, MF_MT_PIXEL_ASPECT_RATIO, MF_MT_SUBTYPE, MF_SOURCE_READER_FIRST_AUDIO_STREAM,
    MF_SOURCE_READER_FIRST_VIDEO_STREAM,
};
use windows::Win32::System::Com::{
    CoCreateInstance, IStream, CLSCTX_INPROC_SERVER, STGM_READ, STGM_SHARE_DENY_NONE,
};
use windows::Win32::UI::Shell::SHCreateStreamOnFileEx;

/// What the letterboxing is filled with. A video is an opaque rectangle — the window it
/// used to be played in was opaque too — so the bars inside a frame are black rather than
/// the backdrop the rest of a frame is composited over.
const BORDER: MFARGB = MFARGB {
    rgbBlue: 0,
    rgbGreen: 0,
    rgbRed: 0,
    rgbAlpha: 255,
};

/// The engine's event sink, which is what the engine needs before it will run at all —
/// `MF_MEDIA_ENGINE_CALLBACK` is required in every mode.
///
/// One event is acted on and the rest are counted: an engine that reports an error has
/// nothing left to hand over, and what this side does about that is stop asking it for
/// frames. The callback arrives on a thread of the engine's own, so what it touches is an
/// atomic and nothing else.
#[implement(IMFMediaEngineNotify)]
struct Notify {
    failed: Arc<AtomicBool>,
}

impl IMFMediaEngineNotify_Impl for Notify_Impl {
    fn EventNotify(&self, event: u32, param1: usize, param2: u32) -> windows::core::Result<()> {
        if event == MF_MEDIA_ENGINE_EVENT_ERROR.0 as u32 {
            let _ = (param1, param2);
            self.failed.store(true, Ordering::Release);
        }

        Ok(())
    }
}

// The playback this thread is running, if it is running one.
//
// It is thread-local because it is not shared: the preview thread is the only thread that
// starts a video, asks for its frames or stops it, and a media engine belongs to the
// apartment that made it. The two threads that have to agree about *which* engine plays a
// video — the load worker measuring a file and the preview thread placing it — only read
// `codecs::plays_video_natively`, which is not this state.
thread_local! {
    static SESSION: RefCell<Option<Session>> = const { RefCell::new(None) };
}

/// One playing video: the engine, the surface its frames are delivered into, and what the
/// surface is currently the size of.
struct Session {
    engine: IMFMediaEngine,
    /// The source the engine was handed. The engine holds it too; this is the handle it
    /// was handed, kept until the video is over so that nothing it is still reading from
    /// can go out of scope under it.
    byte_stream: IMFByteStream,
    /// Where a video's frames are delivered — and nothing at all for a sound, which has no
    /// frames to deliver and no box to draw one in. A session without a surface is the whole
    /// of what playing a sound costs this side.
    bitmap: Option<IWICBitmap>,
    width: u32,
    height: u32,
    path: PathBuf,
    failed: Arc<AtomicBool>,
}

/// The size a video asks to be shown at — its frame, corrected for the pixel shape the
/// file says it has — or `None` when this machine cannot open the file at all.
///
/// This is the measuring half of the path, and it runs on whatever thread the layout is
/// on: what a video's size is has to be known before a preview is placed, and the answer
/// is also the one that decides whether there is a preview to place. A file no reader
/// claims — an `.flv`, a `.rmvb`, an MPEG program stream — is answered with no size, which
/// is how the layout drops it rather than opening a box nothing would be drawn into.
///
/// It is the source reader that is asked rather than the engine, because this question is
/// asked before there is anything to play: the reader opens the file, reports the first
/// video stream's own media type and is dropped, so a hover that is turned down costs a
/// header parse rather than a player.
pub fn dimensions(path: &Path) -> Option<(u32, u32)> {
    if !codecs::mf_started() {
        return None;
    }

    let byte_stream = open_stream(path)?;
    let reader =
        unsafe { MFCreateSourceReaderFromByteStream(&byte_stream, None::<&IMFAttributes>) }.ok()?;
    let media_type =
        unsafe { reader.GetNativeMediaType(MF_SOURCE_READER_FIRST_VIDEO_STREAM.0 as u32, 0) }
            .ok()?;

    // The frame's size and the shape of a pixel, both packed into one `UINT64` each: the
    // width is the high half and the height the low one, and a pixel aspect ratio is a
    // numerator over a denominator the same way.
    let frame = unsafe { media_type.GetUINT64(&MF_MT_FRAME_SIZE) }.ok()?;
    let (frame_width, frame_height) = ((frame >> 32) as u32, frame as u32);
    if frame_width == 0 || frame_height == 0 {
        return None;
    }

    let aspect = unsafe { media_type.GetUINT64(&MF_MT_PIXEL_ASPECT_RATIO) }.unwrap_or(1 << 32);
    let (aspect_width, aspect_height) = ((aspect >> 32) as u32, aspect as u32);

    // A pixel that is not square makes the picture a different shape from its frame, and
    // which axis grows is whichever the ratio takes past one — a DVD's 720 x 480 is shown
    // as 4:3 by a pixel that is wider than it is tall.
    let (width, height) = match (aspect_width, aspect_height) {
        (0, _) | (_, 0) => (frame_width, frame_height),
        (w, h) if w > h => (frame_width.saturating_mul(w) / h, frame_height),
        (w, h) if h > w => (frame_width, frame_height.saturating_mul(h) / w),
        _ => (frame_width, frame_height),
    };

    Some((width, height))
}

/// Start playing `path` into a surface of `width` by `height`, at `volume` per cent.
///
/// Anything already playing is stopped first, so a video is never two videos. A call that
/// could not start one leaves nothing behind rather than a session that will never produce
/// a frame: [`is_playing`] answers for that, and the hover it was for is answered with no
/// preview.
pub fn play(path: &Path, width: u32, height: u32, volume: u32) {
    stop();

    if width == 0 || height == 0 || !codecs::mf_started() {
        return;
    }

    if let Some(session) = Session::begin(path, Some((width, height)), volume) {
        SESSION.with(|slot| *slot.borrow_mut() = Some(session));
    }
}

/// Start playing `path` as a sound at `volume` per cent: the same engine, the same file and the
/// same loop, with no surface and nothing to draw.
///
/// Anything already playing is stopped first, so a sound is never two sounds. A call that could
/// not start one leaves nothing behind rather than a session that will never make a noise:
/// [`is_playing`] answers for that, and the card the hover shows stands alone.
pub fn play_audio(path: &Path, volume: u32) {
    stop();

    if !codecs::mf_started() {
        return;
    }

    if let Some(session) = Session::begin(path, None, volume) {
        SESSION.with(|slot| *slot.borrow_mut() = Some(session));
    }
}

/// Give the surface a new size, which is what a preview that is placed again at another
/// size asks for.
///
/// The engine is told nothing: what it delivers is scaled into whatever rectangle the
/// frame transfer names, so the only thing a new size costs is a new bitmap to deliver
/// into.
pub fn resize(width: u32, height: u32) {
    SESSION.with(|slot| {
        let mut slot = slot.borrow_mut();
        let Some(session) = slot.as_mut() else {
            return;
        };

        // A sound has no surface to resize, and asking one for a new one would be a video's
        // question put to a file that has no picture.
        if session.bitmap.is_none() {
            return;
        }

        if session.width == width && session.height == height || width == 0 || height == 0 {
            return;
        }

        let Some(bitmap) = surface(width, height) else {
            return;
        };

        session.bitmap = Some(bitmap);
        session.width = width;
        session.height = height;
    });
}

/// The file being played, which is what a hover that lands on the same file again compares
/// against rather than restarting it.
pub fn playing_path() -> Option<PathBuf> {
    SESSION.with(|slot| slot.borrow().as_ref().map(|session| session.path.clone()))
}

/// Whether a video is playing, which is also whether one that was just asked for started.
pub fn is_playing() -> bool {
    SESSION.with(|slot| {
        slot.borrow()
            .as_ref()
            .is_some_and(|session| !session.failed.load(Ordering::Acquire))
    })
}

/// How far into the file the running session has played, in seconds.
///
/// It is the engine's own clock rather than this app's, which is what makes it the right
/// answer for a sound: an engine that decodes, paces itself and loops is the one thing that
/// knows where the sound is, and a card drawn from a wall clock beside it would drift from
/// what is being heard. `None` is no session, and a position the engine will not report yet.
pub fn position() -> Option<f64> {
    SESSION.with(|slot| {
        let slot = slot.borrow();
        let session = slot.as_ref()?;

        let seconds = unsafe { session.engine.GetCurrentTime() };

        (seconds.is_finite() && seconds >= 0.0).then_some(seconds)
    })
}

/// How long the file plays, in seconds, where the running session has said.
///
/// A duration is not known the moment a session exists — the engine reads the container's own
/// header first — so a card drawn before it lands is a card with no whole to measure against,
/// and the one drawn a moment later has it (see `audio_preview::Card`).
pub fn duration() -> Option<f64> {
    SESSION.with(|slot| {
        let slot = slot.borrow();
        let session = slot.as_ref()?;

        let seconds = unsafe { session.engine.GetDuration() };

        (seconds.is_finite() && seconds > 0.0).then_some(seconds)
    })
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
/// What comes back is a track with no duration in it: how long a file plays is the engine's
/// answer once one is running, and the engine is not started by a probe.
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

    // The decoder, asked for the only way that answers it: a stream the reader will hand back
    // as PCM is a stream this machine has a decoder for.
    let pcm = unsafe { MFCreateMediaType() }.ok()?;
    unsafe { pcm.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio) }.ok()?;
    unsafe { pcm.SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_PCM) }.ok()?;
    unsafe {
        reader.SetCurrentMediaType(MF_SOURCE_READER_FIRST_AUDIO_STREAM.0 as u32, None, &pcm)
    }
    .ok()?;

    Some(audio_track::Track {
        player: audio_track::Player::Native,
        codec: codec_name(&subtype),
        rate: rate.filter(|rate| *rate > 0),
        channels: channels.filter(|channels| *channels > 0),
        bitrate: bitrate.filter(|bitrate| *bitrate > 0),
        duration: None,
    })
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

/// Take the frame the engine has ready into `pixels`, answering with the size it was
/// written at — or `None` when there is nothing new to take.
///
/// The frame that comes out is the preview's own composition: BGRA, top-down, four bytes
/// to the pixel, at the size the surface is. It is handed to the caller's buffer rather
/// than returned in one of its own, because this runs for every frame of a video that is
/// playing and a fresh megabyte a frame is a megabyte a frame.
pub fn copy_frame_into(pixels: &mut Vec<u8>) -> Option<(u32, u32)> {
    SESSION.with(|slot| {
        let mut slot = slot.borrow_mut();
        let session = slot.as_mut()?;

        session.copy_into(pixels)
    })
}

/// Stop playing and let the engine go. Nothing outside the process is involved, so there is
/// nothing here to wait for: the frames stop being asked for and the player is gone.
pub fn stop() {
    SESSION.with(|slot| {
        if let Some(session) = slot.borrow_mut().take() {
            unsafe { session.engine.Shutdown() }.ok();

            // The player is let go before the source it was reading from is. What the
            // stream is held for is the calls a running engine still makes on it, and
            // there are none of those left.
            drop(session.byte_stream);
        }
    });
}

impl Session {
    fn begin(path: &Path, surface_size: Option<(u32, u32)>, volume: u32) -> Option<Self> {
        let byte_stream = open_stream(path)?;

        // A video is played into a surface of its own; a sound is played into nothing, and the
        // engine is left with no output for a picture at all.
        let (bitmap, width, height) = match surface_size {
            Some((width, height)) => (Some(surface(width, height)?), width, height),
            None => (None, 0, 0),
        };
        let failed = Arc::new(AtomicBool::new(false));

        let mut attributes: Option<IMFAttributes> = None;
        unsafe { MFCreateAttributes(&mut attributes, 3) }.ok()?;
        let attributes = attributes?;

        let notify: IUnknown = Notify {
            failed: Arc::clone(&failed),
        }
        .into();
        unsafe { attributes.SetUnknown(&MF_MEDIA_ENGINE_CALLBACK, &notify) }.ok()?;

        // Frame-server mode delivers frames in one format or another, and this is the one
        // this app composes in: `MFVideoFormat_ARGB32` is a D3D `A8R8G8B8`, which is the
        // same bytes in the same order as the WIC bitmap below. It is asked for where there
        // is a surface to deliver into, and not for a sound.
        if bitmap.is_some() {
            unsafe {
                attributes.SetGUID(&MF_MEDIA_ENGINE_VIDEO_OUTPUT_FORMAT, &MFVideoFormat_ARGB32)
            }
            .ok()?;
        }

        let factory: IMFMediaEngineClassFactory = unsafe {
            CoCreateInstance(&CLSID_MFMediaEngineClassFactory, None, CLSCTX_INPROC_SERVER)
        }
        .ok()?;

        // No playback window and no playback visual: the engine is left in frame-server
        // mode, which is the mode that delivers frames to this side and still renders the
        // audio itself.
        let engine: IMFMediaEngine = unsafe { factory.CreateInstance(0, &attributes) }.ok()?;
        let engine_ex: IMFMediaEngineEx = engine.cast().ok()?;

        // The file is handed over as a stream rather than as a URL, for the reason the PDF
        // engine is: the Shell gives a hovered path in its verbatim form — `\\?\C:\…` —
        // and that form is not a URL anything will open. The path comes with the stream as
        // the name to read it by, since a stream has no name of its own and the handler
        // that opens one is chosen by what the file is called.
        let url = BSTR::from(path.to_string_lossy().as_ref());
        unsafe { engine_ex.SetSourceFromByteStream(&byte_stream, &url) }.ok()?;

        unsafe { engine.SetLoop(true) }.ok()?;

        // A preview is muted by default, and a muted one is asked for as mute rather than
        // as a volume of nothing: the engine is then free to leave the audio path out
        // altogether, which is what the volume setting means at zero on the FFmpeg path
        // as well.
        if volume == 0 {
            let _ = unsafe { engine.SetMuted(true) };
        } else {
            let _ = unsafe { engine.SetVolume(f64::from(volume) / 100.0) };
        }

        unsafe { engine.Load() }.ok()?;
        unsafe { engine.Play() }.ok()?;

        Some(Self {
            engine,
            byte_stream,
            bitmap,
            width,
            height,
            path: path.to_path_buf(),
            failed,
        })
    }

    fn copy_into(&mut self, pixels: &mut Vec<u8>) -> Option<(u32, u32)> {
        // A sound has no frames to take, and the session that plays one is never asked for
        // any: what the card beside it is drawn from is the engine's clock and not its output.
        let bitmap = self.bitmap.as_ref()?;

        if self.failed.load(Ordering::Acquire) {
            return None;
        }

        // A frame is only there to be taken once the engine has current data.
        // `OnVideoStreamTick` reports a frame as due some while before one is queued, and
        // a transfer asked for in that window is *failed* rather than answered late —
        // which is a video that never starts, an `.mp4` being the format that shows it.
        if unsafe { self.engine.GetReadyState() } < MF_MEDIA_ENGINE_READY_HAVE_CURRENT_DATA.0 as u16
        {
            return None;
        }

        unsafe { self.engine.OnVideoStreamTick() }.ok()?;

        let rect = RECT {
            left: 0,
            top: 0,
            right: self.width as i32,
            bottom: self.height as i32,
        };
        let destination: IUnknown = bitmap.cast().ok()?;

        // The whole frame, drawn into the whole surface: what the engine is asked for is
        // the box the layout planned, and it scales the picture into that box and fills
        // what is left of it with the border colour.
        unsafe {
            self.engine
                .TransferVideoFrame(&destination, None, &rect, Some(&BORDER))
        }
        .ok()?;

        copy_locked(bitmap, pixels, self.width, self.height).then_some((self.width, self.height))
    }
}

/// A surface for the engine to deliver into: a bitmap in the format a frame is composed
/// in, made at the size the preview came out at.
fn surface(width: u32, height: u32) -> Option<IWICBitmap> {
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

/// The file as the media stack's own stream, which is the form every reader here is handed.
///
/// The path goes in as it is — verbatim prefix and all — because `SHCreateStreamOnFileEx`
/// is handed a path rather than being asked to resolve a URL, which is what
/// `pdf_preview::open_document` settled for the same reason. The name is set on the stream
/// as well: the handler that opens a byte stream is chosen by the name it carries, and a
/// stream made from a file has none.
fn open_stream(path: &Path) -> Option<IMFByteStream> {
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

/// Copy a locked bitmap out as the preview's frame, which is the one place a video's pixels
/// are touched.
///
/// The alpha is forced opaque rather than taken from the codec. What the engine delivers is
/// a picture, and a picture has no transparency of its own here: the window a video used to
/// be played in was opaque, so a frame that carried an alpha of nothing would be a preview
/// that faded out rather than one that is drawn.
fn copy_locked(bitmap: &IWICBitmap, pixels: &mut Vec<u8>, width: u32, height: u32) -> bool {
    let Some(stride) = (width as usize).checked_mul(4) else {
        return false;
    };
    let Some(needed) = stride.checked_mul(height as usize) else {
        return false;
    };

    let Ok(lock) = (unsafe { bitmap.Lock(std::ptr::null(), WICBitmapLockWrite.0 as u32) }) else {
        return false;
    };
    let Ok(source_stride) = (unsafe { lock.GetStride() }) else {
        return false;
    };

    let mut size: u32 = 0;
    let mut data: *mut u8 = std::ptr::null_mut();
    if unsafe { lock.GetDataPointer(&mut size, &mut data) }.is_err() || data.is_null() {
        return false;
    }
    if (size as usize) < needed {
        return false;
    }

    pixels.clear();
    pixels.resize(needed, 0);

    let source = unsafe { std::slice::from_raw_parts(data, size as usize) };
    for row in 0..height as usize {
        let from = row * source_stride as usize;
        let to = row * stride;
        pixels[to..to + stride].copy_from_slice(&source[from..from + stride]);

        for pixel in pixels[to..to + stride].as_chunks_mut::<4>().0 {
            pixel[3] = 0xFF;
        }
    }

    true
}

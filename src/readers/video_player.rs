//! The media engine Windows has, for a video preview on a machine without FFmpeg.
//!
//! A video is played by `ffplay` when FFmpeg is installed, and by this when it is not.
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

use crate::formats::codecs;
use std::cell::RefCell;
use std::os::windows::ffi::OsStrExt;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Arc;
use windows::core::{implement, BSTR, IUnknown, Interface, PCWSTR};
use windows::Win32::Foundation::RECT;
use windows::Win32::Graphics::Imaging::{
    GUID_WICPixelFormat32bppBGRA, IWICBitmap, WICBitmapCacheOnLoad, WICBitmapLockWrite,
};
use windows::Win32::Media::MediaFoundation::{
    IMFAttributes, IMFByteStream, IMFMediaEngine, IMFMediaEngineClassFactory, IMFMediaEngineEx,
    IMFMediaEngineNotify, IMFMediaEngineNotify_Impl, MFCreateAttributes,
    MFCreateMFByteStreamOnStream, MFCreateSourceReaderFromByteStream, MFARGB,
    MF_BYTESTREAM_ORIGIN_NAME, MF_MEDIA_ENGINE_CALLBACK, MF_MEDIA_ENGINE_EVENT_ERROR,
    MF_MEDIA_ENGINE_READY_HAVE_CURRENT_DATA, MF_MEDIA_ENGINE_VIDEO_OUTPUT_FORMAT, MF_MT_FRAME_SIZE,
    MF_MT_PIXEL_ASPECT_RATIO, MF_SOURCE_READER_FIRST_VIDEO_STREAM, MFVideoFormat_ARGB32,
    CLSID_MFMediaEngineClassFactory,
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
    bitmap: IWICBitmap,
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
    let reader = unsafe { MFCreateSourceReaderFromByteStream(&byte_stream, None::<&IMFAttributes>) }
        .ok()?;
    let media_type = unsafe {
        reader.GetNativeMediaType(MF_SOURCE_READER_FIRST_VIDEO_STREAM.0 as u32, 0)
    }
    .ok()?;

    // The frame's size and the shape of a pixel, both packed into one `UINT64` each: the
    // width is the high half and the height the low one, and a pixel aspect ratio is a
    // numerator over a denominator the same way.
    let frame = unsafe { media_type.GetUINT64(&MF_MT_FRAME_SIZE) }.ok()?;
    let (frame_width, frame_height) = ((frame >> 32) as u32, frame as u32);
    if frame_width == 0 || frame_height == 0 {
        return None;
    }

    let aspect = unsafe { media_type.GetUINT64(&MF_MT_PIXEL_ASPECT_RATIO) }
        .unwrap_or(1 << 32);
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

    if let Some(session) = Session::begin(path, width, height, volume) {
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

        if session.width == width && session.height == height || width == 0 || height == 0 {
            return;
        }

        let Some(bitmap) = surface(width, height) else {
            return;
        };

        session.bitmap = bitmap;
        session.width = width;
        session.height = height;
    });
}

/// The file being played, which is what a hover that lands on the same file again compares
/// against rather than restarting it.
pub fn playing_path() -> Option<PathBuf> {
    SESSION.with(|slot| {
        slot.borrow()
            .as_ref()
            .map(|session| session.path.clone())
    })
}

/// Whether a video is playing, which is also whether one that was just asked for started.
pub fn is_playing() -> bool {
    SESSION.with(|slot| {
        slot.borrow()
            .as_ref()
            .is_some_and(|session| !session.failed.load(Ordering::Acquire))
    })
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
    fn begin(path: &Path, width: u32, height: u32, volume: u32) -> Option<Self> {
        let byte_stream = open_stream(path)?;
        let bitmap = surface(width, height)?;
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
        // same bytes in the same order as the WIC bitmap below.
        unsafe { attributes.SetGUID(&MF_MEDIA_ENGINE_VIDEO_OUTPUT_FORMAT, &MFVideoFormat_ARGB32) }
            .ok()?;

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
        let destination: IUnknown = self.bitmap.cast().ok()?;

        // The whole frame, drawn into the whole surface: what the engine is asked for is
        // the box the layout planned, and it scales the picture into that box and fills
        // what is left of it with the border colour.
        unsafe {
            self.engine
                .TransferVideoFrame(&destination, None, &rect, Some(&BORDER))
        }
        .ok()?;

        copy_locked(&self.bitmap, pixels, self.width, self.height)
            .then_some((self.width, self.height))
    }
}

/// A surface for the engine to deliver into: a bitmap in the format a frame is composed
/// in, made at the size the preview came out at.
fn surface(width: u32, height: u32) -> Option<IWICBitmap> {
    let factory = crate::readers::wic_image::factory()?;

    unsafe {
        factory
            .CreateBitmap(width, height, &GUID_WICPixelFormat32bppBGRA, WICBitmapCacheOnLoad)
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

        for pixel in pixels[to..to + stride].chunks_exact_mut(4) {
            pixel[3] = 0xFF;
        }
    }

    true
}

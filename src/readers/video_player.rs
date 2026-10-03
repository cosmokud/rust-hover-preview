//! The media engine Windows has, for the previews it plays.
//!
//! A video is played by this where its decoders reach the file and by `ffplay` where they do
//! not; a sound is played by this where its own decoders reach the format, and by that player
//! where they do not — the same order for both, and the same question asked of both, which is
//! whether this machine has a decoder for the file (see [`plays`] and [`audio_probe`]).
//!
//! The order is the engine's first because of what only it can be asked: a pause, a seek and a
//! position that are real rather than a player restarted at a second, and frames this app draws
//! itself — which is what lets a pinned window of one be resized, and dragged by its picture.
//! FFmpeg's player is what is left for the files the engine has no decoder for.
//!
//! The two are the same preview to the rest of the app — the same window, the same
//! placement, the same box — because what comes out of here is frames, drawn where every
//! other frame is drawn, rather than a player's window standing in for the preview. That
//! is the shape the media engine is asked for: it is created in *frame-server* mode,
//! which is what it is by default when no playback window is named, and then
//!
//!   * it decodes, paces itself, plays the audio and loops — none of which this app does;
//!   * this side asks, once a tick, whether a frame is due (`OnVideoStreamTick`) and takes it
//!     (`TransferVideoFrame`) into a bitmap of its own, from the crop the probe settled on
//!     where it settled on one (see [`Crop`]).
//!
//! # Where the decoding happens, and why it is not where it should be
//!
//! Every video this app plays is decoded in software, and the first thing to be clear about why
//! is that attaching a DXGI device manager cannot fix it on this machine. A media engine handed
//! one through `MF_MEDIA_ENGINE_DXGI_MANAGER` is documented to use the display's own hardware for
//! the decoding instead, and it does not: the media stack here has no hardware HEVC decoder
//! registered at all, so there is nothing for a manager to select.
//!
//! That makes the measurement this starts from a trap, and it is worth stating plainly because it
//! has been read the wrong way twice. Attaching a manager on a 144 fps 1440p HEVC file shown at
//! a preview box of a display's own size takes the engine from 44.02 s of a core over eight
//! seconds to 0.11 s — which reads as a four-hundred-fold win and is not one. The 0.11 s is an
//! idle engine: it is decoding nothing, because every transfer fails and no frame is ever
//! produced. The control settles it. Remove the device manager and the cost per frame is
//! identical; only the buffer the samples arrive in changes. So the manager was never selecting a
//! hardware decoder here, and the only thing it was selecting was a way for the engine to stop.
//!
//! That is why one is not attached, and not because it would be dearer to build: on this machine
//! the media engine will not then hand a frame over at all. `TransferVideoFrame` accepts a DXGI
//! surface *or* a WIC bitmap, a hardware-decoded frame lives in a texture on the GPU, and the
//! obvious reading is that the destination has to become a texture too — a buffer made by
//! `MFCreateDXGISurfaceBuffer` over a D3D11 texture, read back through a staging copy. That was
//! built, and it is refused. Measured, on this machine, over a 1440p HEVC file and a 1920x800
//! H.264 one:
//!
//!   * a WIC bitmap destination — the one this app already uses — answers the *first* transfer
//!     and every transfer after it with `E_NOINTERFACE` (`0x80004002`). The first frame is
//!     decoded before the DXGI pipeline is up; every frame after it is decoded on the GPU, and
//!     the engine will not render a GPU frame into a bitmap in system memory. That was true with
//!     the output format set to `ARGB32`, to `RGB32`, to `NV12` and unset entirely.
//!   * a `MFCreateDXGISurfaceBuffer` destination is refused on the first transfer as well as the
//!     rest, over a plain `D3D11_USAGE_DEFAULT` texture, over the same texture bound as a render
//!     target, over a surface from `IDXGIDevice::CreateSurface`, at the box's size and at the
//!     picture's own size, and over all of those again with the same four output formats.
//!
//! So the engine will render into a bitmap only while it is decoding in software, and into
//! nothing at all while it is decoding on the GPU. Attaching a manager therefore is not a slower
//! preview: it is one frame, then three seconds of refusals, then the file handed to FFmpeg's
//! player (see [`TRANSFER_FAILURES_GIVE_UP`]) — which is a working preview on a machine that has
//! FFmpeg's player and no preview at all on a machine that does not.
//!
//! What is left of it, and the reason this paragraph is here rather than a branch somewhere, is
//! that the failure above is precisely the one that used to be invisible. It was measured twice
//! before: once as a frame pipeline that froze on the first frame and reported nothing, and once
//! as a hundred and forty-four frames a second decoded in software for a box that could show
//! sixty of them. Anyone reaching for this again should read this first, and should read
//! [`failing_path`] second, because that is the thing standing between the attempt and a preview
//! that hangs.
//!
//! # The one other way a media engine is asked for frames, and why it is not this one
//!
//! There is a second way, and it was built and measured and thrown away, and what it found is
//! worth more than the code was. `IMFMediaEngineEx::EnableWindowlessSwapchainMode` has the engine
//! create a swap chain and present into it, and `GetVideoSwapchainHandle` hands that chain over
//! as an opaque handle; the handle is wrapped with
//! `IDXGIFactoryMedia::CreateSwapChainForCompositionSurfaceHandle`, `GetBuffer(0)` is the engine's
//! own back buffer, and it is read through a staging copy. No `TransferVideoFrame` anywhere in
//! it, so the refusal above is never asked — and the engine is still an `IMFMediaEngine`, which
//! is what makes it look like the answer to everything in this file.
//!
//! It is not, and the reason is narrower and duller than "Media Foundation does not work":
//!
//!   * The mode itself is accepted. So is the handle, the chain, `GetBuffer(0)` and the
//!     `CopyResource` out of it. Every step answers `S_OK`.
//!   * What answers is empty. `Map` on the staging copy succeeds and hands back a row pitch of
//!     *zero* and no address at all, over and over, for as long as the file plays. Nothing is
//!     ever written into the engine's back buffer: zero frames, zero non-zero bytes, and the
//!     clock standing still at zero while the engine believes it is playing.
//!   * The engine will only render into the chain if something is *compositing* the chain. This
//!     is the part that decides it, and it is not a bug in the probe: with no visual and no
//!     window attached, the swap chain has no consumer, so the engine has nothing to present
//!     into and writes nothing. Presenting the chain by hand — recycling a buffer and turning
//!     the rotation on, which puts no pixels anywhere and is not "showing" anything — is what a
//!     flip-model chain needs to advance at all, and it deadlocks the session on the second
//!     read rather than producing a frame.
//!
//! So the path needs a DirectComposition visual to live inside, which means the app's own
//! window and the compositor that draws it, which is a different and much larger project than
//! the one this file is. Nothing about the decoding is unreachable: what is unreachable is
//! *reading* a frame back out of a chain that nobody is displaying.
//!
//! Two of the things learned along the way are worth writing down because they cost the most and
//! are the sort of thing that would be learned again from scratch:
//!
//!   * A `DXGI_SWAP_CHAIN_DESC1` written with `..Default::default()` carries a
//!     `DXGI_SAMPLE_DESC` whose `Count` is **zero**, and every composition swap chain made over
//!     one is refused `DXGI_ERROR_INVALID_CALL` on this machine — every swap effect, every buffer
//!     count, every alpha mode, both formats, and with no media engine in the process at all.
//!     `Count` of one is what makes it, and the same is true of a `D3D11_TEXTURE2D_DESC` for the
//!     staging copy, which is refused `E_INVALIDARG` without it. C++ samples leave both at zero
//!     and get away with it, which is exactly why this reads as a Media Foundation problem and
//!     is not one.
//!   * `IMFDXGIDeviceManager` has no `SetDevice`. The only call that binds a device to one is
//!     `ResetDevice`, which is what Chromium's media foundation renderer does with the manager it
//!     locks; and the device has to be the engine's device rather than merely this side's, since
//!     a staging copy of a buffer on one device cannot be made with another's context.
//!
//! And the third, which is the one that cost the most time and is the reason this section
//! exists at all: `EnableWindowlessSwapchainMode` and `GetVideoSwapchainHandle` have to be asked
//! for *after* `Play`, not before `Load`, and the handle has to be polled rather than asked for
//! once. Before the engine is running the first is accepted and never acted on and the second
//! answers `S_OK` and a null handle. And `CreateSwapChainForCompositionSurfaceHandle` is a race:
//! asked for promptly it answers `S_OK`, and asked for a few hundred milliseconds later — the
//! engine having made its own chain over the same handle in the meantime — it answers
//! `DXGI_ERROR_ALREADY_EXISTS`. Which is the same finding as the one above, wearing a different
//! hat: there is always already a chain there, and it is not one anybody is showing.
//!
//! # The third way, which works and is still not the answer
//!
//! There is a third, and unlike the two above it does hand pixels back: an `IMFSourceReader`
//! handed an `IMFDXGIDeviceManager` through `MF_SOURCE_READER_D3D_MANAGER`, with advanced video
//! processing on, decoding the file through the same software HEVC decoder and answering
//! `ReadSample` with an `IMFDXGIBuffer` over a texture that a staging copy maps and reads. No
//! compositor is involved at any point and the pixels come back as decoded picture in system
//! memory, which is the thing the engine above would not do for anything. So the refusal is
//! narrower than it reads: it belongs to that engine rather than to Media Foundation, which is
//! worth having written down before anyone concludes the latter.
//!
//! Measured on this machine over eight seconds with the process's own clock, five runs of each
//! because the number moves a few milliseconds between them: 20.3–24.3 ms of CPU a frame to
//! decode alone, 25.3–31.2 ms a frame to decode and copy the staging texture back, against the
//! 128.99 ms a frame this side spends handing one over (345 frames in 8.01 s for 44.50 s of CPU,
//! the same instrument as `tests::video_take_cost`). Four to six times cheaper, which is a real
//! result and not a rounding error. It is also nowhere near the two milliseconds or so a hardware
//! decode of a 1440p picture would cost, and it could not be, because the decoder underneath is
//! the software one and no arrangement of attributes changes which decoder is registered. Only
//! `NV12` would negotiate as an output type at all — `BGRA` was refused `MF_E_INVALIDMEDIATYPE` at
//! every size and every frame rate tried — and the NV12 staging layout that had the most reason
//! to be trouble was not trouble at all.
//!
//! **So the frame pipeline is not being rewritten on a source reader**, and the measurements are
//! why: a source reader would hand this app a presentation clock, a WASAPI audio client and an
//! error channel that the engine gives away for free, and would buy four to six times on a
//! machine whose decoder is software whatever asks for it. The spike that measured all of this is
//! gone and the numbers are what is left of it, which is the arrangement the rest of this file
//! takes anyway. What it is worth is the last thing it found, and that is the way to read every
//! other measurement here: **a claim that hardware decode engaged is gated on CPU cost and never
//! on having been handed a GPU buffer.** A device manager wraps a software decoder's output in an
//! `IMFDXGIBuffer` as readily as a hardware one's — here the DXGI buffer came back over the same
//! twenty milliseconds as the plain system-memory one — so the buffer says which path produced the
//! frame and says nothing whatever about what decoded it.
//!
//! A tick is not a frame. The clock this side ticks on is a vertical blank, sixty times a
//! second whatever the file runs at, so most ticks of a film find the engine offering the
//! picture it offered the last four of them, and a tick that finds one is answered without
//! taking it at all (see [`is_a_new_frame`]). That is most of the difference between a video
//! previewed on the engine and a video previewed by FFmpeg's player: the same file is
//! decoded, copied and handed to the compositor as many times as it has frames rather than
//! as many times as the display refreshes, and a 4K one is four times the frame of a
//! 1080p one throughout. It is also why the two ends of it have to be read together: a 144 fps
//! file is decoded at 144 fps and drawn at sixty, so most of what the engine decodes is never
//! shown — nearly free on a GPU, and on a preview box of a display's own size it is the largest
//! single cost there is (see `tests::video_take_cost`).
//!
//! Who *scales* the picture is settled by the two sizes alone: a box meaningfully larger than
//! the picture is one this side scales into it ([`scale_rows`]), because which filter the
//! engine reads a picture at is not a question this app can ask it, let alone choose. What each
//! side of that choice costs is written where it is made (see [`play`]).
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
//! The questions asked of the engine from outside are all answered here. [`dimensions`]
//! is a video's shape, asked of a source reader before there is anything to play. [`plays`]
//! is whether a decoder here can hand a video's own frames back at all — what the router
//! asks before it hands a file to this engine or to FFmpeg's player. [`audio_track`] is the
//! same question about a sound, asked the same way — the reader is asked for *decoded* PCM,
//! which it can only give where the decoder is registered, and the media type the file
//! declares is read for the rest. And [`position`] and [`duration`] are what the running
//! session says about itself, for the card's clock and a pinned window's transport bar.

use crate::formats::codecs;
use crate::formats::head;
use crate::paths::plain_path;
use crate::readers::audio_track;
use once_cell::sync::Lazy;
use std::cell::RefCell;
use std::collections::HashMap;
use std::os::windows::ffi::OsStrExt;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};
use windows::core::{implement, IUnknown, Interface, BSTR, GUID, PCWSTR};
use windows::Win32::Foundation::RECT;
use windows::Win32::Graphics::Imaging::{
    GUID_WICPixelFormat32bppBGRA, IWICBitmap, IWICBitmapLock, WICBitmapCacheOnLoad,
    WICBitmapLockWrite,
};
use windows::Win32::Media::MediaFoundation::{
    CLSID_MFMediaEngineClassFactory, IMFAttributes, IMFByteStream, IMFMediaEngine,
    IMFMediaEngineClassFactory, IMFMediaEngineEx, IMFMediaEngineNotify, IMFMediaEngineNotify_Impl,
    IMFSourceReader, MFAudioFormat_AAC, MFAudioFormat_ADTS, MFAudioFormat_ALAC,
    MFAudioFormat_AMR_NB, MFAudioFormat_AMR_WB, MFAudioFormat_DTS, MFAudioFormat_Dolby_AC3,
    MFAudioFormat_Dolby_DDPlus, MFAudioFormat_FLAC, MFAudioFormat_Float, MFAudioFormat_MP3,
    MFAudioFormat_Opus, MFAudioFormat_PCM, MFAudioFormat_Vorbis, MFAudioFormat_WMAudioV8,
    MFAudioFormat_WMAudioV9, MFAudioFormat_WMAudio_Lossless, MFCreateAttributes,
    MFCreateMFByteStreamOnStream, MFCreateMediaType, MFCreateSourceReaderFromByteStream,
    MFMediaType_Audio, MFMediaType_Video, MFVideoFormat_ARGB32, MFVideoFormat_RGB32,
    MFVideoNormalizedRect, MFARGB, MF_BYTESTREAM_ORIGIN_NAME, MF_MEDIA_ENGINE_CALLBACK,
    MF_MEDIA_ENGINE_EVENT_ERROR, MF_MEDIA_ENGINE_READY_HAVE_CURRENT_DATA,
    MF_MEDIA_ENGINE_READY_HAVE_METADATA, MF_MEDIA_ENGINE_VIDEO_OUTPUT_FORMAT,
    MF_MT_AUDIO_NUM_CHANNELS, MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_AVG_BITRATE, MF_MT_FRAME_SIZE,
    MF_MT_MAJOR_TYPE, MF_MT_PIXEL_ASPECT_RATIO, MF_MT_SUBTYPE, MF_PD_DURATION,
    MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING, MF_SOURCE_READER_FIRST_AUDIO_STREAM,
    MF_SOURCE_READER_FIRST_VIDEO_STREAM, MF_SOURCE_READER_MEDIASOURCE,
};
use windows::Win32::System::Com::{
    CoCreateInstance, IStream, CLSCTX_INPROC_SERVER, STGM_READ, STGM_SHARE_DENY_NONE,
};
use windows::Win32::UI::Shell::SHCreateStreamOnFileEx;

/// What the letterboxing is filled with, where a box has any letterboxing in it at all.
///
/// The box a preview is placed at is the shape of the picture that goes into it — the probe's
/// crop where it settled on one, the frame itself where it did not — so an ordinary preview is
/// filled to its own edges and this is never reached. What is left for it is the box that has
/// been given another shape since: a pinned window dragged by its edges, into which the engine
/// scales the picture and pads what is left. A video is an opaque rectangle — the window it
/// used to be played in was opaque too — so the padding is black rather than the backdrop the
/// rest of a frame is composited over.
const BORDER: MFARGB = MFARGB {
    rgbBlue: 0,
    rgbGreen: 0,
    rgbRed: 0,
    rgbAlpha: 255,
};

/// How long a seek a session was asked for may wait for the engine to be ready for one.
///
/// Reading a file's own header is a fraction of a second's work for a file on this machine —
/// it is the header, not the file — so this is a give-up rather than a wait anything is
/// expected to reach. What it is here for is the file the engine never gets to the header of:
/// what a sound is answered with then is its beginning, rather than a seek asked for again on
/// every read of its clock for as long as it is hovered.
const SEEK_GIVE_UP: Duration = Duration::from_secs(3);

/// How long a session is given to hand over its first frame before the engine is taken to be
/// failing at the file rather than slow to start it.
///
/// Reading a file's first frame is a fraction of a second's work for a file on this machine, so
/// this is a give-up rather than a wait anything is expected to reach. It is generous on purpose:
/// what the answer does is hand the file to FFmpeg's player, and a file that was merely slow to
/// start would be a preview taken away from the engine that could have drawn it.
const FIRST_FRAME_GIVE_UP: Duration = Duration::from_secs(3);

/// How many transfers in a row have to fail before a session is a session that has failed.
///
/// A transfer that fails is not a video that has stopped: it is a video the engine is offering
/// a frame of and will not give over. One of those is ordinary — a seek asks the engine for a
/// frame it has not decoded yet, a loop lands on the time it started from, a gap in a file is a
/// gap — and a run of them is what is not ordinary. What this bounds is therefore a *run*, and
/// it is counted in ticks rather than seconds for the reason [`FIRST_FRAME_GIVE_UP`] is counted
/// in seconds: the two are the same length of time measured two ways, the preview loop turning
/// sixty times a second, and a bound in seconds would have to assume that sixty itself.
///
/// One hundred and eighty-eight is three seconds of those ticks — the nearest whole tick at or
/// above [`FIRST_FRAME_GIVE_UP`], which is the resolution a count of ticks has — and deliberately
/// so: the answer to "this engine cannot draw this file" should arrive in about as long as the
/// answer to "this engine has not drawn this file", whatever order the two mistakes arrive in. A
/// bound a great deal shorter would hand a perfectly good film over to FFmpeg's player every time
/// a network drive hiccuped for a quarter of a second.
///
/// What this is for is the fault that has no symptom at all. A session that has drawn a frame
/// and cannot draw another keeps the frame it had and reports nothing: the picture on screen is
/// the picture from a second ago and every question the app asks it is answered well. That is
/// exactly what the DXGI device manager does to a preview on this machine — one frame off the
/// software path, then nothing but refusals off the hardware one (see this module's own note on
/// where the decoding happens).
const TRANSFER_FAILURES_GIVE_UP: u32 = 188;

/// Whether a run of transfers that all failed is a session that has failed, which is the whole
/// of what the count above is for.
///
/// A function of one number rather than of a session and a clock because the question has to be
/// answerable without one: the run *is* the evidence, and what it has to be judged against is a
/// count rather than a moment, since a session that took a seek to get here has no useful moment
/// to measure a gap against.
fn a_run_of_refusals_is_a_failure(run: u32) -> bool {
    run >= TRANSFER_FAILURES_GIVE_UP
}

/// How much bigger than the picture a box has to be before this side is the one that scales into
/// it, as a share of the picture's own size: one part in fifty.
///
/// The choice is between the two engines, and they cost very different amounts. The engine
/// scales on the GPU as part of the frame it is already writing, so its share of the work is
/// nothing this app can measure; this side's share is a scalar bilinear over every pixel of
/// every frame, on the preview thread (see [`scale_rows`]). So the answer to "who is bigger" has
/// to be a threshold and not a comparison, and what it is worth is a trade between a picture
/// very slightly too small for its box — where the engine scales and the difference is a
/// hundredth of a pixel per pixel, and nobody would ever have known — against the alternative,
/// where a box one pixel wider than the picture sends a whole film down the scalar path and
/// pays for it on every frame.
///
/// What makes this the number rather than anything near it is that the two sizes do not move
/// together. A box is the picture at the scale setting times the share of the work area it was
/// given, and the layout computes both in integers: a 1920-wide picture at 50% is 960 and
/// nothing is round, so the smallest enlargement a box of a whole size can arrive at is a
/// fraction of a percent, and a threshold of a hair would fire on nearly every file at every
/// setting but 100%. Two per cent is above every one of those and far below the enlargements
/// worth resampling for, and the largest share that is not worth it — a 4K file shown on a
/// 4K display by `fit`, where the two sizes are within a pixel of one another — is exactly the
/// case this is here for.
const ENLARGEMENT_WORTH_RESAMPLING: f64 = 1.02;

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
    /// The stream the engine was handed, kept until the video is over so that nothing it is
    /// still reading from can go out of scope under it.
    byte_stream: IMFByteStream,
    /// Where a video's frames are delivered, and nothing at all for a sound — which has no
    /// frames to deliver and no box to draw one in.
    bitmap: Option<IWICBitmap>,
    /// The box the preview draws at: the size of the frame this side hands over, and the size
    /// the engine is asked to deliver at only while it is the one scaling (see `scaled`).
    width: u32,
    height: u32,
    /// The size the picture has in the file, which is what the engine is asked for where this
    /// side scales it.
    picture: (u32, u32),
    /// Whether this side scales the picture into the box rather than the engine: the two sizes
    /// above are what settles it, and what each of the two costs is written on [`play`].
    scaled: bool,
    /// The two rows of the picture already read across to the box's width, kept between the
    /// rows of the box mixed from them, and empty while the engine is the one scaling.
    rows: [Vec<u8>; 2],
    /// Where in the frame the picture is taken from, where the probe settled on a crop: the
    /// rectangle the engine's frame transfer is asked for. The whole frame is `None`, which is
    /// a file the probe found no crop in and every sound.
    source: Option<MFVideoNormalizedRect>,
    path: PathBuf,
    failed: Arc<AtomicBool>,
    /// Whether a frame of the file has been handed over at all: a session is not a promise that
    /// a picture will come of it, since the engine accepts a file whose decoder and converter
    /// are both there and whose *pipeline* is not (see [`failing_path`]).
    ///
    /// It says whether a frame has *ever* been handed over and is not cleared by a transfer that
    /// fails after one, because that is a different fault: a file that played and then met a bad
    /// sector is not a file the engine cannot draw, and handing it to FFmpeg's player over one
    /// costs the engine its decoder for the rest of a film it was getting on with (see
    /// [`transfer_failures`] and [`failing_before_a_frame`]).
    drew: bool,
    /// How many transfers in a row the engine has refused, reset by the one that succeeds.
    ///
    /// This is what distinguishes a video that is not moving from a video this side cannot get
    /// frames out of, and the distinction is invisible from outside: both answer the loop's tick
    /// with no frame and neither reports anything wrong. The run is bounded and reaching
    /// [`TRANSFER_FAILURES_GIVE_UP`] of it is the session failing, which is what makes this
    /// something the app can be told about rather than a picture that freezes.
    transfer_failures: u32,
    /// The time of the frame the caller is holding, in the hundred-nanosecond units the engine
    /// ticks in, and `None` where the next frame is owed whatever the engine says about it — a
    /// session that has just been begun, resized or sought somewhere has a picture the caller
    /// has not seen, and a seek in particular can land on the very time it has just drawn
    /// (see [`is_a_new_frame`]).
    drawn: Option<i64>,
    /// Where the sound was asked to start, while the engine has not taken it there yet: `Load`
    /// answers before the header is there, so the position is kept and made on the first tick
    /// that finds the engine loaded (see `apply_seek`).
    pending_seek: Option<f64>,
    /// When the session was started, which bounds the wait for a pending seek: an engine that
    /// has not read the header of its file in this long is not going to, and a seek left
    /// standing is a seek made on every read of the clock after it.
    began: Instant,
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

    let frame = unsafe { media_type.GetUINT64(&MF_MT_FRAME_SIZE) }.ok()?;
    let (frame_width, frame_height) = unpack_pair(frame);
    if frame_width == 0 || frame_height == 0 {
        return None;
    }

    let aspect = unsafe { media_type.GetUINT64(&MF_MT_PIXEL_ASPECT_RATIO) }.unwrap_or(1 << 32);
    let (aspect_width, aspect_height) = unpack_pair(aspect);

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

/// The part of a video's own frame a preview is drawn from: the rectangle the geometry probe's
/// cropdetect pass settled on, in the frame's own pixels, together with the frame it is a part
/// of.
///
/// It is the one thing about a crop that both players have to be told, and each is told it in
/// its own terms — FFmpeg's player as a filter on its command line, and the engine as the source
/// rectangle of its frame transfer, which is normalized over the frame rather than in pixels
/// (see [`source_rect`]). What it is for is the difference between a box with black bars in it
/// and a picture that fills the box: the box is placed at the crop's shape, so a *whole* frame
/// drawn into one is scaled down to fit and padded with the border colour on the two sides the
/// file's own bars leave over.
#[derive(Clone, Copy)]
pub struct Crop {
    pub x: u32,
    pub y: u32,
    pub width: u32,
    pub height: u32,
    /// The frame the rectangle is a part of: the probe reads both together, and what the engine
    /// is told is the rectangle as a share of this.
    pub frame_width: u32,
    pub frame_height: u32,
}

/// The picture a preview of a video is drawn from: the size the file is shown at, and — where the
/// probe settled on one — the part of the frame that picture is cut from.
///
/// The two halves are what the choice of who scales the picture is made of, so they are handed
/// over together rather than one being worked out from the other: a box larger than the picture
/// is a picture this side scales into it, and a box that is not is one the engine scales into it
/// (see [`play`]).
#[derive(Clone, Copy, Default)]
pub struct Picture {
    /// The size the picture is shown at — the box at 100%, and what every other share of the
    /// setting is a share of. A sound has no picture and is played with none of this.
    pub width: u32,
    pub height: u32,
    /// Where in the frame the picture is taken from, where the file carries its own bars: the
    /// rectangle both players are told about, each in its own terms (see [`Crop`]).
    pub crop: Option<Crop>,
}

/// A crop as the rectangle the engine's frame transfer is asked for — the same region, in the
/// normalized coordinates that call takes — or nothing at all for a rectangle with no area in
/// it or a frame with no size.
fn source_rect(crop: Crop) -> Option<MFVideoNormalizedRect> {
    if crop.width == 0 || crop.height == 0 || crop.frame_width == 0 || crop.frame_height == 0 {
        return None;
    }

    let frame_width = crop.frame_width as f32;
    let frame_height = crop.frame_height as f32;

    Some(MFVideoNormalizedRect {
        left: crop.x as f32 / frame_width,
        top: crop.y as f32 / frame_height,
        right: (crop.x + crop.width) as f32 / frame_width,
        bottom: (crop.y + crop.height) as f32 / frame_height,
    })
}

/// Start playing `path` into a surface of `width` by `height`, at `volume` per cent, from the
/// picture `picture` names.
///
/// The two sizes are what settles who scales the picture into that box: a box meaningfully
/// larger than the picture is a file being shown above its own size, and one the engine is
/// asked for at the picture's own size and this side scales (see [`scale_rows`]), while a box
/// that is not larger is the engine's to fill as it always was (see `scales_here`). What the
/// first costs is a resample of every frame and what the second costs is nothing beyond the
/// copy every frame paid before it — a preview drawn at or below the picture's own size is the
/// copy it always was, and a preview drawn a per cent above it is deliberately left there too,
/// since the difference between the two engines' scaling and this one's is not worth a whole
/// film at a hundredth of a pixel (see [`ENLARGEMENT_WORTH_RESAMPLING`]).
///
/// Anything already playing is stopped first, so a video is never two videos. A call that
/// could not start one leaves nothing behind rather than a session that will never produce
/// a frame: [`is_playing`] answers for that, and the hover it was for is answered with no
/// preview.
pub fn play(path: &Path, width: u32, height: u32, volume: u32, picture: Picture) {
    stop();

    if width == 0 || height == 0 || !codecs::mf_started() {
        return;
    }

    if let Some(session) = Session::begin(path, Some((width, height)), picture, volume, 0.0) {
        SESSION.with(|slot| *slot.borrow_mut() = Some(session));
    }
}

/// Start playing `path` as a sound at `volume` per cent and `start` seconds in: the same engine,
/// the same file and the same loop, with no surface and nothing to draw.
///
/// Anything already playing is stopped first, so a sound is never two sounds. A call that could
/// not start one leaves nothing behind rather than a session that will never make a noise:
/// [`is_playing`] answers for that, and the card the hover shows stands alone.
///
/// Where the sound starts is the caller's answer rather than this one's — what a file's own
/// length is, and what the tray has been asked for, are questions this side is not asked (see
/// `audio_seek`). What it is handed is a number of seconds, and a sound that starts at the
/// beginning is handed zero.
pub fn play_audio(path: &Path, volume: u32, start: f64) {
    stop();

    if !codecs::mf_started() {
        return;
    }

    // A sound has no picture, so there is no frame for a crop to be a part of and none is
    // handed over: the argument is a video's question and is answered as one by the caller.
    if let Some(session) = Session::begin(path, None, Picture::default(), volume, start) {
        SESSION.with(|slot| *slot.borrow_mut() = Some(session));
    }
}

/// Hold the sound or picture where it is, or let it go on.
///
/// It is the engine's own pause, which is the one thing this app cannot do for itself: what is
/// playing is decoded ahead of the clock, and a pause that only stopped drawing would go on
/// playing. What it is asked by is a pinned preview's transport bar — the one window of this
/// app's that has buttons of its own — and it does nothing where nothing is playing, which is a
/// pin whose file has already been let go.
pub fn set_paused(paused: bool) {
    SESSION.with(|slot| {
        let mut slot = slot.borrow_mut();
        let Some(session) = slot.as_mut() else {
            return;
        };

        let engine = &session.engine;
        let _ = unsafe {
            if paused {
                engine.Pause()
            } else {
                engine.Play()
            }
        };

        // The picture the caller is holding is not known to be the picture the engine is
        // holding: a file asked to pause with a frame queued up hands that frame over on
        // the way, and whether the next tick reports it as a new one is the engine's
        // business rather than a question this app can answer from here. What a pause or
        // an unpause costs is one frame drawn that may have been drawn already, which is
        // nothing beside the frame that is not drawn because a hold was mistaken for one.
        session.drawn = None;
    });
}

/// Take the sound of what is playing to `volume` per cent, which is what a pinned preview's own
/// volume control asks for the moment its knob is moved — the one thing about a running session
/// that can be changed while it runs, since a seek is deferred and a pause is a state.
///
/// A level of nothing is asked for as mute rather than as a volume of zero, which is the form a
/// session is started in as well: the engine is then free to leave the audio path out
/// altogether. Nothing is done where nothing is playing, which is a pin whose file has already
/// been let go of; the level a player is started at is the caller's answer, as it has always
/// been (see [`play`]).
pub fn set_volume(volume: u32) {
    SESSION.with(|slot| {
        let slot = slot.borrow();
        let Some(session) = slot.as_ref() else {
            return;
        };

        let engine = &session.engine;
        unsafe {
            if volume == 0 {
                let _ = engine.SetMuted(true);
            } else {
                let _ = engine.SetMuted(false);
                let _ = engine.SetVolume(f64::from(volume.min(100)) / 100.0);
            }
        }
    });
}

/// Take the sound that is playing to `seconds` into its file, which is what a start position
/// the probe could not work out asks for once the engine has said how long the file is.
///
/// It is asked of the session rather than of the engine because the seek may have to wait for
/// one (see `Session::pending_seek`), and it does nothing where nothing is playing — a sound
/// whose file has been left by the time the length lands is a sound that is over.
pub fn seek(seconds: f64) {
    SESSION.with(|slot| {
        if let Some(session) = slot.borrow_mut().as_mut() {
            session.seek(seconds);
        }
    });
}

/// Give the surface a new size, which is what a preview that is placed again at another
/// size asks for.
///
/// The engine is told nothing: what it delivers is scaled into whatever rectangle the
/// frame transfer names, so the only thing a new size costs is a new bitmap to deliver
/// into.
///
/// Which rectangle that is depends on the box the new size makes, and the two sides of the
/// picture's own size cost different things. A box meaningfully larger than the picture is
/// this side's to scale, and the surface it reads is the picture's own — the same surface
/// whatever the box is, so a window dragged about up there is dragged about without a bitmap
/// being made. A box at or below the picture's size is the engine's to scale into, and the
/// surface has to *be* that box: the engine writes where it is told to, so a box that changed
/// is a bitmap that is made again (see `scales_here`).
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

        let scaled = scales_here(session.picture, width, height);

        // The surface has to be made again wherever the engine is the one that scales — every box
        // it is handed is the box it writes — and wherever the scaling has just come back to this
        // side, whose surface is the picture's rather than the box's.
        if !scaled || scaled != session.scaled {
            let size = if scaled {
                session.picture
            } else {
                (width, height)
            };

            let Some(bitmap) = surface(size.0, size.1) else {
                return;
            };

            session.bitmap = Some(bitmap);
        }

        session.scaled = scaled;
        session.width = width;
        session.height = height;

        // A box is not a picture, and the frame the caller is holding was drawn at the size
        // the last box was: the next frame is owed to it whatever the engine makes of the
        // time, because the time has not changed at all. This is where the frame is first
        // asked for again, so it is also where the tick has to be believed for the first
        // time after it.
        session.drawn = None;
    });
}

/// Whether this side scales the picture into the box rather than asking the engine to: a box
/// meaningfully larger than the picture on either axis is a file shown above its own size,
/// which is the one case the engine's own scaling is not asked for.
///
/// "Meaningfully" is [`ENLARGEMENT_WORTH_RESAMPLING`] and the difference matters because of
/// what the two sides cost, which is written on [`play`]: this side's share is a scalar
/// bilinear over every pixel of every frame and the engine's is a share of a frame it is
/// writing anyway, so a box one pixel wider than the picture is a whole film resampled on the
/// preview thread to correct for a hundredth of a pixel. A preview at 100% is the picture
/// itself and is not scaled by anybody, and that is where a hover leaves a video most of the
/// time — the resampler is for a file deliberately shown larger than it is, not for the
/// rounding of a box that was meant to be its own size.
///
/// Either axis rather than both, because the two are not always scaled alike — a file whose
/// pixels are not square is stretched along one of them at every size — and a picture already
/// being asked for at its own size may as well be scaled by the one sampler for both
/// directions. A picture with no size at all is a sound, which has nothing to scale and
/// nothing to ask for.
fn scales_here(picture: (u32, u32), width: u32, height: u32) -> bool {
    picture.0 > 0
        && picture.1 > 0
        && (f64::from(width) > f64::from(picture.0) * ENLARGEMENT_WORTH_RESAMPLING
            || f64::from(height) > f64::from(picture.1) * ENLARGEMENT_WORTH_RESAMPLING)
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

/// The file the engine is failing at, if it is failing at one: a session with a surface that has
/// been up a while and has not handed over a single frame of it, or one whose transfers have been
/// refused for long enough to be a fault rather than a hiccup.
/// It is how the app finds out that a question a probe answered yes to was still the wrong
/// question. A probe asks the engine's *parts* — a decoder for the stream and a converter to the
/// format this app composes in — and the engine plays through a pipeline of its own, which can
/// have no mode for a file the parts both accept: an H.264 film in a chroma the hardware decoder
/// does not do is decoded and converted happily by a source reader, and never drawn by the
/// engine. What is left of such a file without this is a preview of the placeholder pixels — the
/// pixels a video preview is opened with and that nothing replaces — so the app asks this on the
/// tick and answers it with [`mark_unplayable`], which is what hands the file to FFmpeg's player.
///
/// Only a session with a surface is asked about: a sound has no frames to hand over and is never
/// waiting for one.
pub fn failing_path() -> Option<PathBuf> {
    SESSION.with(|slot| {
        let slot = slot.borrow();
        let session = slot.as_ref()?;

        // A session that was asked to pause hands nothing over because nothing was asked of it,
        // which is not a session failing at its file: a pin whose pause was pressed before the
        // first frame arrived is a file this side stopped.
        if unsafe { session.engine.IsPaused() }.as_bool() {
            return None;
        }

        (engine_is_failing(
            session.bitmap.is_some(),
            session.drew,
            session.transfer_failures,
            session.began.elapsed(),
        ))
        .then(|| session.path.clone())
    })
}

/// Whether the file behind a running session is one the engine is failing at, in the sense
/// [`failing_path`] uses the words: a session with a surface to draw into that is either not
/// getting frames or cannot take the ones it is offered.
///
/// It is two faults in one condition because they are the same fault arrived at by two roads,
/// and what they have in common is the length of time: three seconds of an engine that will not
/// draw this file and three seconds of an engine that will not give this one a frame are both a
/// file FFmpeg's player has to be handed. What they are *not* is one thing about `drew`, and
/// that is the half of the condition with the reasoning on it: `drew` says a frame has been
/// handed over at all, so it is asked about a session that has not, and a session that has drawn
/// is asked about the run of refusals alone — the frame it saw being exactly what makes a run of
/// refusals worth reporting rather than merely embarrassing.
///
/// The four arguments rather than a `&Session` because the third is a number this does not own
/// — it belongs to a tick — and a question that has to be asked with a borrow of a session on
/// the preview thread is a question that cannot be asked in a test at all.
fn engine_is_failing(surface: bool, drew: bool, refusals: u32, age: Duration) -> bool {
    surface
        && (a_run_of_refusals_is_a_failure(refusals) || (!drew && age >= FIRST_FRAME_GIVE_UP))
}

/// The file the engine is failing at before it has drawn a frame of it, if it is failing at one
/// there: the union of the two ways a pinned window's session comes to have no picture in it,
/// where [`failing_path`] waits out only the second of them.
///
/// The first is the engine saying so: `MF_MEDIA_ENGINE_EVENT_ERROR` raised into [`Notify`], the
/// engine admitting outright that it cannot play the file it was handed rather than taking its
/// time over it, and worth acting on the tick it arrives rather than a second later. The second
/// is [`FIRST_FRAME_GIVE_UP`] running out with still nothing drawn, which is the engine with
/// nothing to report because it has no pipeline to report on (see [`failing_path`]) — the same
/// failure arrived at by silence rather than by an error, and so only ever to be found by
/// waiting.
///
/// [`failing_path`] is not this and still is what it is, because a hover has shown nothing yet:
/// it can leave a file that is merely slow to start on the engine another tick and be no worse
/// off, while a preview taken away from the engine that could have drawn it is a real loss. A
/// pinned window is in no such position, having been put up for this file and standing there
/// until something else is put up instead — so the tick the engine gives up on the file is the
/// tick the window is a box of placeholder pixels that nothing further is coming to replace, and
/// that tick is the one this asks about.
///
/// A session that has drawn a frame is left out, which is the same distinction drawn the other
/// way round: a file that played and then met a bad sector is not a file the engine cannot
/// draw, and handing it to FFmpeg's player over a fault that has nothing to do with what plays
/// it costs the engine its decoder for the rest of a file it was getting on with. What
/// [`mark_unplayable`] writes down is that the probe's answer was wrong about the file, and a
/// file that has drawn a frame is one the probe got right.
///
/// A paused session is left out for the reason [`failing_path`] gives, and it is the whole of
/// it: a pin whose pause was pressed before the first frame arrived is a file this side
/// stopped, and a window its own transport has stopped is not a broken preview to be taken away
/// from the reader of it.
///
/// Only a session with a surface is asked about, as in [`failing_path`]: a sound has no frames
/// to hand over and is never waiting for one, so an engine failing at a sound is not this
/// either.
pub fn failing_before_a_frame() -> Option<PathBuf> {
    SESSION.with(|slot| {
        let slot = slot.borrow();
        let session = slot.as_ref()?;

        // A session that was asked to pause hands nothing over because nothing was asked of it,
        // which is not a session failing at its file: a pin whose pause was pressed before the
        // first frame arrived is a file this side stopped.
        if unsafe { session.engine.IsPaused() }.as_bool() {
            return None;
        }

        (session.bitmap.is_some()
            && !session.drew
            && (session.failed.load(Ordering::Acquire)
                || session.began.elapsed() >= FIRST_FRAME_GIVE_UP))
            .then(|| session.path.clone())
    })
}

/// Write a file down as one the media engine cannot draw, and let go of the session that was
/// failing at it, so that what plays the file from here on is FFmpeg's player.
///
/// The answer is held where the probe's own answer is held and the same way — per file and the
/// version of it, in the same map — because it is the same question. What a probe answers yes to
/// and the engine then fails at is that answer being wrong about this file, and a correction
/// belongs where the answer was read from: the routing asks [`plays`] and gets this, the next
/// hover of the file takes the other road, and the load that follows agrees with it. A file
/// written again is a file asked about again, since the version is part of the key.
///
/// The session is stopped rather than left running: an engine that has not drawn a frame of the
/// file in all this time is not going to, and what it is holding is the file.
pub fn mark_unplayable(path: &Path) {
    let key = head::key(path);

    if let Ok(mut held) = PLAYABLE.lock() {
        if held.len() >= PLAYABLE_MAX_ENTRIES {
            held.clear();
        }
        held.insert(key, false);
    }

    if playing_path().as_deref() == Some(path) {
        stop();
    }
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

/// Take the running session to the position it was asked to start at, where the engine has not
/// taken it there yet.
///
/// A seek cannot be made the instant a session is begun — the engine has not read the file's own
/// header — so it is kept and made here instead, and what this is for is the *moment* it is
/// made: it is called on the tick a sound is on screen for, which is a sixtieth of a second,
/// while the clock the card is drawn from is read four times a second. A seek waiting on the
/// slower of those is a quarter of a second of the file's beginning heard before the sound is
/// where it was asked to start, and this is what makes what is heard of the beginning the time
/// it takes to read a header instead (see `Session::take_pending`).
pub fn apply_seek() {
    SESSION.with(|slot| {
        if let Some(session) = slot.borrow_mut().as_mut() {
            session.take_pending();
        }
    });
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

/// Whether this machine can play `path` as a video: the question the router asks before it hands
/// a file to this engine rather than to FFmpeg's player.
///
/// It is the question `audio_probe` asks about a sound, asked about a picture and the same way:
/// the file is opened as a source reader, it has to hold a video stream, and the reader is then
/// asked to produce *decoded* RGB32 on it — a decoder *and* a converter, which is what the
/// engine needs before a frame can be handed over. A no here is a file FFmpeg's player takes,
/// and the two engines between them are why a video is previewed at all on a machine that has
/// only one of them.
///
/// What is deliberately not asked is whether the file would play *well*: a file this engine opens
/// and then fails on is answered where every other failure about a preview is, and a probe that
/// decoded a frame to find out would make every video hover pay for a decoder's worth of work.
pub fn can_play(path: &Path) -> bool {
    if !codecs::mf_started() {
        return false;
    }

    let Some(byte_stream) = open_stream(path) else {
        return false;
    };

    // The reader is asked for decoded frames in the format this app composes in, and that is a
    // question with two halves: a decoder for the stream, and something to turn what the decoder
    // hands back into RGB32. Asked without `MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING` it asks
    // whether the *decoder itself* delivers RGB32, which no H.264 decoder does: every video of
    // every kind would be answered no, and the file would go to FFmpeg's player for a reason
    // that has nothing to do with the file.
    let mut attributes: Option<IMFAttributes> = None;
    if unsafe { MFCreateAttributes(&mut attributes, 1) }.is_err() {
        return false;
    }
    let Some(attributes) = attributes else {
        return false;
    };
    if unsafe { attributes.SetUINT32(&MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING, 1) }.is_err() {
        return false;
    }

    let Ok(reader) =
        (unsafe { MFCreateSourceReaderFromByteStream(&byte_stream, Some(&attributes)) })
    else {
        return false;
    };

    // A file with no video stream at all — a sound, a container of something else — is not a
    // video, and there is nothing to be asked about it beyond that.
    if unsafe { reader.GetNativeMediaType(MF_SOURCE_READER_FIRST_VIDEO_STREAM.0 as u32, 0) }
        .is_err()
    {
        return false;
    }

    let Ok(frames) = (unsafe { MFCreateMediaType() }) else {
        return false;
    };
    if unsafe { frames.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Video) }.is_err() {
        return false;
    }
    if unsafe { frames.SetGUID(&MF_MT_SUBTYPE, &MFVideoFormat_RGB32) }.is_err() {
        return false;
    }

    // The decoder, asked for the only way that answers it: a stream the reader will hand back as
    // frames is a stream this machine has a decoder for.
    unsafe {
        reader.SetCurrentMediaType(MF_SOURCE_READER_FIRST_VIDEO_STREAM.0 as u32, None, &frames)
    }
    .is_ok()
}

/// What the engine said about a file, held by the path and the version of the file it was said
/// about: what the router asks before every video it previews, answered once per version.
///
/// It is `audio_track`'s arrangement for the same reason: a probe opens a file and builds a
/// decoder chain for it, and a hover that asks twice — a measure and a render, a hover that comes
/// back to a file — must not pay for that twice. A file that is written again is a file whose
/// answer is asked again, because the version is part of the key.
pub fn plays(path: &Path) -> bool {
    let key = head::key(path);

    if let Some(answer) = PLAYABLE
        .lock()
        .ok()
        .and_then(|held| held.get(&key).copied())
    {
        return answer;
    }

    let answer = can_play(path);

    if let Ok(mut held) = PLAYABLE.lock() {
        if held.len() >= PLAYABLE_MAX_ENTRIES {
            held.clear();
        }
        held.insert(key, answer);
    }

    answer
}

/// What the engine has been asked about, one answer per file and version. It is bounded the way
/// the sound tracks are: a run that hovers a library of films does not grow a map forever, and
/// what a clear costs is one probe per file rather than an unbounded map.
static PLAYABLE: Lazy<Mutex<HashMap<head::Key, bool>>> = Lazy::new(|| Mutex::new(HashMap::new()));
const PLAYABLE_MAX_ENTRIES: usize = 512;

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

/// Take the frame the engine has ready into `pixels`, answering with the size it was
/// written at — or `None` when there is nothing new to take.
///
/// The frame that comes out is the preview's own composition: BGRA, top-down, four bytes
/// to the pixel, at the size the surface is. It is handed to the caller's buffer rather
/// than returned in one of its own, because this runs for every frame of a video that is
/// playing and a fresh megabyte a frame is a megabyte a frame.
///
/// `None` is the whole of "there is nothing to repaint for", and a caller is to read it that
/// way and leave both its buffer and the compositor alone. It is a deliberately wider answer
/// than it was: a tick that finds the engine still holding the picture already on the screen
/// is a tick with no work in it, and the buffer is left as it was rather than filled again
/// with the picture the caller already has (see [`is_a_new_frame`]). It is also wider than the
/// failure it stands for, which is deliberate and is where the narrower question is asked
/// instead: of the four answers [`Take`] has, only one is a frame, and the three that are not
/// are all a tick with nothing to paint — including the one that is a fault, which the session
/// keeps to itself and which [`failing_path`] is the question to ask about (see
/// [`Session::refused`]).
pub fn copy_frame_into(pixels: &mut Vec<u8>) -> Option<(u32, u32)> {
    SESSION.with(|slot| {
        let mut slot = slot.borrow_mut();
        let session = slot.as_mut()?;

        match session.copy_into(pixels) {
            Take::Drawn(size) => Some(size),
            Take::Nothing | Take::Refused | Take::Failed => None,
        }
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
    fn begin(
        path: &Path,
        surface_size: Option<(u32, u32)>,
        picture: Picture,
        volume: u32,
        start: f64,
    ) -> Option<Self> {
        let byte_stream = open_stream(path)?;

        // A video is played into a surface of its own; a sound is played into nothing, and the
        // engine is left with no output for a picture at all.
        let (bitmap, width, height, scaled) = match surface_size {
            Some((width, height)) => {
                // Whose scaling the picture is handed to is settled here, once, because it is what
                // the surface is made of as well as what every frame of the session costs: the
                // picture's own size where this side scales it, and the box itself where the engine
                // does (see `scales_here`).
                let scaled = scales_here((picture.width, picture.height), width, height);
                let (surface_width, surface_height) = if scaled {
                    (picture.width, picture.height)
                } else {
                    (width, height)
                };

                (
                    Some(surface(surface_width, surface_height)?),
                    width,
                    height,
                    scaled,
                )
            }
            None => (None, 0, 0, false),
        };

        // What the picture is taken from, which is settled here rather than at every frame:
        // the crop where the probe settled on one, and the whole frame where it did not.
        let source = picture.crop.and_then(source_rect);

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
        // and that form is not a URL anything will open. The name goes on the stream as the
        // *plain* form, which is the one the engine resolves: handed the verbatim path the
        // resolution fails, and the failure arrives *after* `Play` has already answered — as
        // an error on the notify callback a moment later — so a sound was a card whose clock
        // never moved, with nothing on this side able to say why. The stream itself is still
        // opened on the verbatim path — long paths and odd names open by it, and it is not a
        // name anything parses — so what changes is only the name handed over beside it.
        let url = BSTR::from(plain_path(path).as_str());
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

        let mut session = Self {
            engine,
            byte_stream,
            bitmap,
            width,
            height,
            picture: (picture.width, picture.height),
            scaled,
            rows: [Vec::new(), Vec::new()],
            source,
            path: path.to_path_buf(),
            failed,
            drew: false,
            transfer_failures: 0,
            // The first frame of a session is the first picture the caller has of the file,
            // so it is owed one whatever the engine makes of a tick that finds its own
            // first frame already on its clock.
            drawn: None,
            pending_seek: None,
            began: Instant::now(),
        };

        // A sound that was asked to start somewhere other than the beginning is taken there
        // before anything is heard of the beginning of it: the engine has the file's header by
        // the time `Load` has answered for a local file often enough that this call lands, and
        // a call it was too early for is kept and made again on the first tick after it (see
        // `take_pending`).
        session.seek(start);

        Some(session)
    }

    /// Ask to be taken to `seconds` into the file. Nothing is a position rather than the
    /// beginning: the beginning is somewhere a transport bar can be dragged to, and a seek
    /// dropped because the second asked for is the zeroth is a bar that does nothing at its
    /// own left-hand end.
    fn seek(&mut self, seconds: f64) {
        if !seconds.is_finite() {
            return;
        }

        self.pending_seek = Some(seconds.max(0.0));
        self.take_pending();
    }

    /// The seek this session is still owed, made where the engine is ready for one.
    ///
    /// What "ready" means is the engine's own ready state rather than a wait of this side's: a
    /// seek asked for before the media is loaded is one the engine refuses, and one asked for
    /// the moment it lands is what this app wants — so what is watched is the state that says
    /// the file's own header has been read.
    ///
    /// A seek is made once and not checked afterwards, because a source that will not seek and
    /// a position past the end of a file that says nothing about its length are both answered
    /// by the engine playing on from where it is. The one length that is read is where the
    /// file's own says the position asked for is past it, which is a remembered position past
    /// a file edited since it was asked for (see `audio_seek::planned`).
    fn take_pending(&mut self) {
        let Some(target) = self.pending_seek else {
            return;
        };

        // The wait is bounded by the session's own age and not by the age of the seek: a seek
        // asked of a file that has been playing for a minute is not one that arrived late, and
        // an engine that has not read a header by now is not going to (see `SEEK_GIVE_UP`).
        if unsafe { self.engine.GetReadyState() } < MF_MEDIA_ENGINE_READY_HAVE_METADATA.0 as u16 {
            if self.began.elapsed() >= SEEK_GIVE_UP {
                self.pending_seek = None;
            }
            return;
        }

        let duration = unsafe { self.engine.GetDuration() };
        if duration.is_finite() && duration > 0.0 && target >= duration {
            self.pending_seek = None;
            return;
        }

        let _ = unsafe { self.engine.SetCurrentTime(target) };
        self.pending_seek = None;

        // A seek is not told apart from playing by the picture: the engine moves its own
        // clock and the next tick can hand over a frame whose time is the one the caller
        // is already holding, and on a video cut at a whole second — or on a bar dragged
        // back to where it was — that is not a corner case but the ordinary one. So a
        // position change is a new picture owed, whatever its time turns out to say.
        self.drawn = None;
    }

    /// The one place a frame is asked for, and the only place a session is found to be failing
    /// at its file.
    ///
    /// It is a tick of the preview loop: the engine's own ready state, its own offer of a frame,
    /// its own identity for that frame, the transfer, and then the copy out of the surface the
    /// transfer wrote into. Each of those can end the tick with nothing to paint, and each says
    /// something different about why — which is what [`Take`] is for, since only one of them is
    /// a fault and a fault flattened into the other two is a fault nothing is done about.
    fn copy_into(&mut self, pixels: &mut Vec<u8>) -> Take {
        // A sound has no frames to take, and the session that plays one is never asked for
        // any: what the card beside it is drawn from is the engine's clock and not its
        // output. It is the one of the three answers below that is not about a tick at all.
        let Some(bitmap) = self.bitmap.as_ref() else {
            return Take::Nothing;
        };

        if self.failed.load(Ordering::Acquire) {
            return Take::Failed;
        }

        // A frame is only there to be taken once the engine has current data.
        // `OnVideoStreamTick` reports a frame as due some while before one is queued, and
        // a transfer asked for in that window is *failed* rather than answered late —
        // which is a video that never starts, an `.mp4` being the format that shows it.
        if unsafe { self.engine.GetReadyState() } < MF_MEDIA_ENGINE_READY_HAVE_CURRENT_DATA.0 as u16
        {
            return Take::Nothing;
        }

        // The tick is where the engine says whether there is a *new* frame, and answers
        // with the presentation time of the one there is. `S_FALSE` — which arrives as
        // anything but `S_OK` here — is the engine saying there is not, and is the
        // ordinary answer on two ticks in three of a file that runs at half the tick: a
        // tick that finds no new frame is no frame, and the caller is told so rather than
        // shown the picture it already has.
        let pts = match unsafe { self.engine.OnVideoStreamTick() } {
            Ok(pts) => pts,
            Err(_) => return Take::Nothing,
        };

        // And the engine saying there is one is not the same question as whether it is a
        // different one: while a film plays, the same picture is offered on every tick
        // until the next one is ready, and copying one out of the engine and into the
        // compositor again is thirty megabytes of memory traffic at the size of a 4K
        // display for a picture nobody is waiting for. The time is the identity, and a
        // frame at a time already drawn is left in the engine's own surface where it is.
        if !is_a_new_frame(self.drawn, pts) {
            return Take::Nothing;
        }

        // The rectangle the engine is asked to write: the box, where it is the one scaling, and the
        // picture's own size where this side is — where what comes back is the picture as the file
        // holds it, for the sampler below to read into the box (see `scales_here`).
        let (delivered_width, delivered_height) = if self.scaled {
            self.picture
        } else {
            (self.width, self.height)
        };

        let rect = RECT {
            left: 0,
            top: 0,
            right: delivered_width as i32,
            bottom: delivered_height as i32,
        };
        // A surface that cannot be asked to be a destination is a surface frames cannot be taken out
        // of, which is the fault rather than the ordinary answer: a tick that said otherwise
        // would be a tick that froze the preview without ever reporting anything wrong, which is
        // the whole thing this side is here to stop. It cannot actually happen — the cast is a
        // COM identity query and the bitmap is a COM object — and it is counted rather than
        // swallowed anyway, since the cost of being wrong here is a frozen preview and the cost
        // of counting it is three seconds of a session that was never going to work.
        let Ok(destination) = bitmap.cast::<IUnknown>() else {
            return self.refused();
        };

        // The picture drawn into the whole surface: the engine is asked for the box the layout
        // planned and for the part of the frame that goes into it — the crop the probe settled
        // on, where there is one, which is the shape that box was placed at. What is left of
        // the box is filled with the border colour, which an ordinary preview has nothing of;
        // asked for the whole frame while the box is the shape of a crop inside it, the engine
        // is left with the difference to pad, and the preview grows black bars down the sides
        // the file's own bars do not cover.
        if unsafe {
            self.engine.TransferVideoFrame(
                &destination,
                self.source.as_ref().map(std::ptr::from_ref),
                &rect,
                Some(&BORDER),
            )
        }
        .is_err()
        {
            return self.refused();
        }

        // One that got through is a run of nothing, whatever the run was: a file that hiccuped
        // and came back is a file that is playing, and a counter that survived it would be
        // counting two sessions' faults as one.
        self.transfer_failures = 0;

        // A frame of the file has been drawn: whatever becomes of this session later, it is not
        // a session that never had one (see `failing_path`). What is drawn is what the
        // caller is holding from here on, and it is held only once the engine has written
        // it: a transfer that failed leaves the surface as it was, so recording the frame
        // as taken would be a picture the caller never saw remembered as one it has.
        self.drew = true;
        self.drawn = Some(pts);

        let copied = if self.scaled {
            resample_locked(
                bitmap,
                pixels,
                self.picture,
                (self.width, self.height),
                &mut self.rows,
            )
        } else {
            copy_locked(bitmap, pixels, self.width, self.height)
        };

        // A surface that would not open is not the engine failing at the file and is not a run
        // of anything: the frame was written, the caller was not shown it, and the next tick
        // finds the same time and leaves it where the engine has it. What is reported to the
        // caller is the one answer it has always been given for that, which is no frame.
        if copied {
            Take::Drawn((self.width, self.height))
        } else {
            Take::Nothing
        }
    }

    /// A frame was offered and could not be taken out of the engine: counted into the run, and
    /// the session failed once the run is long enough to be a fault rather than a hiccup.
    ///
    /// It sets the *engine's own* failure flag rather than one of its own, and that is the whole
    /// of what makes the fault visible outside: the flag is what [`is_playing`] and both of the
    /// failing paths read, so a session whose transfers keep being refused is answered exactly
    /// as one whose engine gave an error event — which is the truth, and the reason the three
    /// seconds after which the app hands the file to FFmpeg's player arrives on a schedule
    /// rather than never.
    fn refused(&mut self) -> Take {
        self.transfer_failures = self.transfer_failures.saturating_add(1);

        if a_run_of_refusals_is_a_failure(self.transfer_failures) {
            self.failed.store(true, Ordering::Release);

            return Take::Failed;
        }

        Take::Refused
    }
}

/// What a tick of the preview loop was answered with, which is a question the caller cannot ask
/// and the session has to keep the answer to.
///
/// Three of these are no frame at all and the caller is given one answer for all of them, which
/// is right — a tick with nothing in it is a tick with nothing to paint and nothing to say about
/// it. What they are not is the same *fact*, and keeping them apart here is the only place they
/// can be kept apart: one of the three is a fault, and a fault that cannot be told from the
/// ordinary answer is a fault nothing is ever done about.
enum Take {
    /// A frame, at the size it was written for the caller to compose.
    Drawn((u32, u32)),
    /// The engine was not offering a frame this tick: it has none new, or not one yet, or the
    /// session has been asked for nothing at all. This is the ordinary answer on two ticks in
    /// three of any film, and the one [`is_a_new_frame`]'s dedup rests on being able to tell
    /// apart from the next.
    Nothing,
    /// The engine had a frame this tick and would not give it over, which is this side's fault
    /// and not the file's, and which is counted into a run so that enough of them is not one.
    Refused,
    /// The session has failed, by an error event out of the engine or by a run of refusals this
    /// side reached the bound of, and is handing over nothing at all from here on.
    Failed,
}

/// Whether the frame the engine says is ready at `pts` is one the caller has not been given.
///
/// The tick carries the presentation time of the picture it is offering, and a picture is not
/// new because a tick happened: the clock this side ticks on is a vertical blank and runs at
/// whatever refresh the display is set to, so between one frame of a film and the next the
/// engine offers that same frame three or four times over. Every one of those offers costs a
/// transfer out of the engine, a copy into the caller and a hand of the whole frame to the
/// compositor, and none of them is a picture anybody is waiting for — which is why this is the
/// difference between a preview on the engine and a preview of a film at four times its own
/// size, since the frame is the largest thing in the loop and it is proportional to the square
/// of the file.
///
/// `drawn` is `None` wherever the next frame is owed rather than merely new, and each of the
/// places that leaves it `None` is a place where the caller's picture and the engine's have no
/// reason to agree: a session has drawn nothing yet, a box has changed size under a frame
/// drawn at the old one, and a seek has moved a clock a transport bar is drawn from. A seek is
/// the one that earns the distinction outright — a bar dragged back to where it was asks for
/// the very time just drawn, and a picture that genuinely changed hands back a time the caller
/// is already holding.
fn is_a_new_frame(drawn: Option<i64>, pts: i64) -> bool {
    drawn != Some(pts)
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
fn resample_locked(
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
fn rows_read_per_destination_row(source: &[u8], picture: (u32, u32), box_size: (u32, u32)) -> f64 {
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
const A_WHOLE_PIXEL: u64 = 1 << 16;

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
fn scale_rows(
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
fn interpolate_row(
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
fn copy_locked(bitmap: &IWICBitmap, pixels: &mut Vec<u8>, width: u32, height: u32) -> bool {
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
fn copy_row_opaque(from: &[u8], to: &mut [u8]) {
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

#[cfg(test)]
mod tests {
    use super::*;

    /// The name a media source is read by is the Shell's path with the verbatim prefix
    /// taken off, which is the whole of what the engine was failing on: a verbatim path
    /// is not a URL, and the engine resolves this name as one.
    #[test]
    fn reads_the_plain_form_of_a_verbatim_path() {
        assert_eq!(
            plain_path(Path::new(r"\\?\C:\Music\track.mp3")),
            r"C:\Music\track.mp3",
            "the verbatim form of a local path is the path it is written around"
        );
        assert_eq!(
            plain_path(Path::new(r"\\?\UNC\server\share\track.mp3")),
            r"\\server\share\track.mp3",
            "and the verbatim form of a share keeps its server"
        );
        assert_eq!(
            plain_path(Path::new(r"C:\Music\track.mp3")),
            r"C:\Music\track.mp3",
            "a path that was never verbatim is left exactly as it is"
        );
        assert_eq!(
            plain_path(Path::new(r"\\server\share\track.mp3")),
            r"\\server\share\track.mp3",
            "and so is a share written the ordinary way"
        );
    }

    /// A file the engine has been watched failing at is a file FFmpeg's player plays from then
    /// on: the mark is written where the probe's own answer is kept and read by the same
    /// question, so what the routing asks about the file is the answer this side learned by
    /// watching it rather than the one it guessed at.
    #[test]
    fn a_file_the_engine_failed_at_is_ffmpegs_from_here_on() {
        // A file of this module's own holding nothing a decoder could read, so that the answer
        // the mark corrects is one about a file no engine can play rather than about whatever
        // this machine happens to have decoders for.
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-video-tests")
            .join("unplayable");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let path = folder.join("not-a-film.mp4");
        std::fs::write(&path, b"not a film at all").expect("the file");

        mark_unplayable(&path);

        assert!(
            !plays(&path),
            "the mark is what the router is answered with, not the probe"
        );
    }

    /// The pixel at a row and column of a frame of the box's own width.
    fn pixel_at(bgra: &[u8], width: usize, row: usize, column: usize) -> [u8; 4] {
        let at = (row * width + column) * 4;

        bgra[at..at + 4].try_into().expect("four bytes to a pixel")
    }

    /// A picture read above its own size is read *between* its pixels, which is the whole of
    /// what this side scales a preview above 100% for: the edge of a picture that was two pixels
    /// wide arrives as the four steps between them rather than as the two it was made of.
    #[test]
    fn an_enlarged_edge_is_read_between_its_pixels() {
        // A 2x2 picture of one edge: black down the left of it and blue down the right.
        let source = [
            0u8, 0, 0, 255, 255, 0, 0, 255, //
            0, 0, 0, 255, 255, 0, 0, 255,
        ];
        let mut out = vec![0u8; 4 * 4 * 4];
        let mut rows = [Vec::new(), Vec::new()];

        assert!(scale_rows(&source, 8, (2, 2), (4, 4), &mut rows, &mut out));

        // The outermost pixels stand on the two the picture has, and each of the two between them
        // is a quarter of the way from one to the other — where a picture read a source pixel at a
        // time would have arrived as `0, 0, 255, 255`.
        let expected = [0u8, 64, 191, 255];

        for row in 0..4 {
            for (column, level) in expected.iter().enumerate() {
                assert_eq!(
                    pixel_at(&out, 4, row, column),
                    [*level, 0, 0, 255],
                    "row {row}, column {column} of the edge, and every pixel of it opaque"
                );
            }
        }
    }

    /// The other direction is read the same way — a picture read below its own size lands between
    /// four of its pixels and is answered with their average — which is what a file whose pixels
    /// are not square asks of the one axis it is stretched along at every size.
    #[test]
    fn a_shrunken_picture_is_read_between_its_pixels_too() {
        // Black above blue, read into the one pixel between them.
        let source = [
            0u8, 0, 0, 255, 0, 0, 0, 255, //
            255, 0, 0, 255, 255, 0, 0, 255,
        ];
        let mut out = vec![0u8; 4];
        let mut rows = [Vec::new(), Vec::new()];

        assert!(scale_rows(&source, 8, (2, 2), (1, 1), &mut rows, &mut out));

        assert_eq!(out, [128, 0, 0, 255], "the pixel standing between them");
    }

    /// Who scales the picture is one question and these are its two sizes: a box meaningfully larger
    /// than the picture on either axis is a file shown above its own size and is this side's to
    /// scale, and everything at or below the picture's own size is the engine's, as it always
    /// was.
    #[test]
    fn a_box_larger_than_the_picture_is_the_one_this_side_scales() {
        assert!(
            scales_here((640, 480), 1280, 960),
            "shown at twice its size"
        );

        assert!(
            !scales_here((640, 480), 640, 480),
            "a preview at 100% is the picture itself, and nothing is scaled by anybody"
        );
        assert!(!scales_here((640, 480), 320, 240), "a preview below 100%");

        assert!(
            scales_here((720, 480), 872, 480),
            "a pixel wider than it is tall is a stretch along one axis, which is still this side's"
        );
        assert!(
            !scales_here((0, 0), 320, 240),
            "a sound has no picture to scale, and a box of nothing is not one either"
        );
    }

    /// A box a hair bigger than the picture is not a file shown above its own size, and the
    /// difference is the whole of this side's argument: the two sizes do not come out of the
    /// layout equal even at 100%, and a comparison rather than a share resamples a whole film
    /// for a hundredth of a pixel. So the three cases are the three answers, and the middle one
    /// is the one that used to be the first.
    #[test]
    fn only_an_enlargement_worth_the_resampler_comes_to_this_side() {
        assert!(
            !scales_here((1920, 1080), 1920, 1080),
            "an exact match is the picture itself, and is what a preview at 100% is"
        );

        // One per cent: a 4K file on a 4K display, which is what `fit` places, and the largest
        // enlargement that is not worth resampling for by any distance.
        assert!(
            !scales_here((3840, 2160), 3878, 2182),
            "a box a per cent bigger than a 4K picture is left to the engine"
        );
        assert!(
            !scales_here((1920, 1080), 1939, 1094),
            "and the same a per cent on either axis, however the rounding fell"
        );

        assert!(
            !scales_here((640, 480), 641, 480),
            "nor is one pixel, which is what the rounding of a box that was meant to be the picture gives"
        );
        assert!(!scales_here((640, 480), 640, 481), "on either axis alone");

        assert!(
            scales_here((640, 480), 960, 720),
            "while one and a half times is a file deliberately shown larger than it is"
        );
        assert!(
            scales_here((1920, 1080), 2560, 1440),
            "and a 1080p file enlarged onto a 4K display, which is the case the resampler exists for"
        );
        assert!(
            scales_here((1920, 1080), 1920, 2200),
            "along one axis only, which is what a picture of non-square pixels asks of at every size"
        );
    }

    /// The two rows a row of the box is mixed from are kept between the rows that need them,
    /// and the only way to see whether they are is to count what is read: a resampler that
    /// reads each row twice and one that reads each once draw exactly the same frame. At twice
    /// the picture's size each source row is wanted by two rows of the box and at four times by
    /// four, so the two numbers are one half and one quarter — and a cache that keeps the upper
    /// row but reads the lower one again on every tick comes to one and a bit either way.
    #[test]
    fn every_row_of_the_picture_is_read_once_however_far_it_is_stretched() {
        // A picture big enough for the mapping's arithmetic to be the whole of the cost and
        // small enough to be written out by hand.
        let picture = (100u32, 100u32);
        let source = vec![0u8; picture.0 as usize * picture.1 as usize * 4];

        let doubled = rows_read_per_destination_row(&source, picture, (200, 200));
        assert!(
            doubled <= 0.5,
            "a row of the picture is wanted by two of the box, so it is read once each: {doubled:.3}"
        );

        let quadrupled = rows_read_per_destination_row(&source, picture, (400, 400));
        assert!(
            quadrupled <= 0.25,
            "and by four at four times the size, for the same reason: {quadrupled:.3}"
        );

        // A box the same size as the picture is the other end of it: one row wanted by one row,
        // and one read for it. It is also the last row of the picture under every other box,
        // where the pair is a row and the same row again and so is read once rather than twice.
        let matched = rows_read_per_destination_row(&source, picture, (100, 100));
        assert!(
            matched <= 1.0,
            "the last row of a picture is its own lower half, so a whole frame reads one row fewer than it has: {matched:.3}"
        );

        // One axis equal and the other stretched is what a file of non-square pixels asks for at
        // every size, and it is asked along the height here: the row cache does not care which
        // axis it is, and the horizontal half of this is a copy rather than a resample.
        let stretched = rows_read_per_destination_row(&source, picture, (100, 200));
        assert!(
            stretched <= 0.5,
            "and a box stretched on one axis only reads each row of the picture once for it as well: {stretched:.3}"
        );

        let stretched_across = rows_read_per_destination_row(&source, picture, (200, 100));
        assert!(
            stretched_across <= 1.0,
            "and one row for one row is one read for one row, whatever the columns are doing: {stretched_across:.3}"
        );
    }

    /// A box as wide as the picture needs no resampling across, and the row it is given is the row
    /// it was handed rather than a mix of it with itself: which is the difference between reading
    /// a row and copying it, and is what a picture of non-square pixels gets on the axis that is
    /// not being stretched.
    #[test]
    fn a_box_as_wide_as_the_picture_is_given_the_row_unchanged() {
        // One row of a picture, with the alpha bytes a converter chose rather than this app.
        let source = [
            10u8, 20, 30, 0, //
            40, 50, 60, 253, 70, 80, 90, 254, 100, 110, 120, 255,
        ];
        let mut row = [0u8; 16];

        interpolate_row(&source, 16, 4, 0, A_WHOLE_PIXEL, &mut row);

        let expected = [
            [10u8, 20, 30, 255],
            [40, 50, 60, 255],
            [70, 80, 90, 255],
            [100, 110, 120, 255],
        ];

        for (column, pixel) in expected.iter().enumerate() {
            assert_eq!(
                pixel_at(&row, 4, 0, column),
                *pixel,
                "column {column} arrived at from the source pixel and no other, and is opaque"
            );
        }
    }

    /// The alpha byte the colour converter writes into `ARGB32` is not one anybody could draw
    /// with, so the copy writes it rather than reading it: which is what a row of the engine's
    /// surface has to be for the preview to be opaque at all.
    #[test]
    fn a_copied_row_is_opaque_whatever_the_converter_wrote_in_it() {
        // Three pixels carrying the three alpha values a transfer into `MFVideoFormat_ARGB32`
        // has been seen to leave behind, plus the all-zero alpha of a frame nothing has been
        // written into yet.
        let source = [
            10u8, 20, 30, 0, //
            40, 50, 60, 253, 70, 80, 90, 254, 100, 110, 120, 255,
        ];
        let mut row = [0u8; 16];

        copy_row_opaque(&source, &mut row);

        let expected = [
            [10u8, 20, 30, 255],
            [40, 50, 60, 255],
            [70, 80, 90, 255],
            [100, 110, 120, 255],
        ];

        for (column, pixel) in expected.iter().enumerate() {
            assert_eq!(
                pixel_at(&row, 4, 0, column),
                *pixel,
                "pixel {column} keeps the colour it arrived with and is handed an alpha that can be composited"
            );
        }
    }

    /// A picture the engine wrote less of than it said it would is answered with no frame rather
    /// than with one drawn from the rows it managed: half a picture is not a picture, and the
    /// frame the preview is holding is a better answer than a half-filled one.
    #[test]
    fn a_picture_short_of_its_own_size_is_no_frame_at_all() {
        // One row of a two-row picture, read into a box with room for it.
        let source = [0u8; 8];
        let mut out = vec![9u8; 4 * 4 * 4];
        let mut rows = [Vec::new(), Vec::new()];

        assert!(!scale_rows(&source, 8, (2, 2), (4, 4), &mut rows, &mut out));
        assert!(
            out.iter().all(|byte| *byte == 9),
            "and what was there is left as it was"
        );
    }

    /// The engine offering the frame it offered last is not a frame, and a session that has
    /// drawn nothing is owed the first picture whatever the engine has to say about it. This is
    /// the whole of the question, and it is asked in hundred-nanosecond units because that is
    /// what the tick is answered in.
    #[test]
    fn a_frame_the_engine_is_still_holding_is_not_drawn_again() {
        assert!(
            is_a_new_frame(None, 0),
            "the first frame of a session is new however the engine times it"
        );

        assert!(
            is_a_new_frame(Some(0), 333_333),
            "and so is the frame that follows the one drawn, a hundredth of a second later"
        );

        assert!(
            !is_a_new_frame(Some(333_333), 333_333),
            "while the same time is the same picture however many ticks it is offered over"
        );

        // What a file that loops hands back, and what a seek hands back that a bar was
        // dragged to: the beginning of the file arriving again, and the time of a frame
        // that did change arriving as the time of one that did not.
        assert!(
            is_a_new_frame(Some(12_000_000), 0),
            "the loop of a file starts its times over without starting its frames"
        );
        assert!(
            is_a_new_frame(None, 12_000_000),
            "and a seek leaves the caller owed the time it lands on, exactly as it was"
        );
    }

    /// A run of transfers that all failed is a session that has failed, and a run that does not
    /// is not: which is the whole of what the bound is for, and the only thing standing between
    /// a video that is not moving and a video this side has stopped being able to ask for.
    ///
    /// The two ends are what a file on a network drive is and what this module's own frame path
    /// was: a gap, a seek and a loop are a tick or two of this, while a pipeline that cannot give
    /// a frame up at all is three seconds of it and never anything else. A bound anywhere near
    /// the first number would hand a film over to another engine every time the network hiccuped,
    /// and a bound nowhere near the second would leave a preview frozen on a frame from a second
    /// ago reporting nothing whatever was wrong.
    #[test]
    fn a_transfer_that_keeps_failing_is_a_session_that_has_failed() {
        assert!(
            !a_run_of_refusals_is_a_failure(0),
            "a session that has transferred every frame is not failing"
        );

        assert!(
            !a_run_of_refusals_is_a_failure(1),
            "one refused transfer is a seek, a loop landing on its own first time, or a gap"
        );
        assert!(
            !a_run_of_refusals_is_a_failure(TRANSFER_FAILURES_GIVE_UP - 1),
            "and so is a run a tick short of the bound, however long the film has been playing"
        );

        assert!(
            a_run_of_refusals_is_a_failure(TRANSFER_FAILURES_GIVE_UP),
            "while a run of the whole bound is a pipeline that will not give a frame up"
        );
        assert!(
            a_run_of_refusals_is_a_failure(TRANSFER_FAILURES_GIVE_UP * 2),
            "and a longer one is the same fault rather than a worse one"
        );

        // The bound is three seconds of the preview loop's own ticks, which is the same length of
        // time as the wait a session gets for its first frame. A file this side cannot take
        // frames out of is a file another engine has to take over, and the two answers ought to
        // arrive in about the same time whichever order the two mistakes happen in — within a
        // tick, which is the resolution a count of ticks has.
        let bound = Duration::from_millis(TRANSFER_FAILURES_GIVE_UP as u64 * 16);
        assert!(
            bound >= FIRST_FRAME_GIVE_UP && bound < FIRST_FRAME_GIVE_UP + Duration::from_millis(16),
            "the bound is FIRST_FRAME_GIVE_UP counted in the loop's ticks: {bound:?}"
        );
    }

    /// The four answers a tick can have are four different facts and the caller is given one
    /// answer for three of them, which is right — a tick with nothing to paint is a tick with
    /// nothing to say — but only because the session keeps them apart for itself. The two
    /// faults this module now has to tell apart are a session that never drew and a session
    /// that cannot take a frame out of the one it has, and the difference between reporting the
    /// second and not the first is this condition.
    #[test]
    fn a_file_whose_frames_cannot_be_taken_out_is_a_failure_after_the_first_frame_too() {
        let give_up = FIRST_FRAME_GIVE_UP;
        let fresh = FIRST_FRAME_GIVE_UP * 2;
        let short = TRANSFER_FAILURES_GIVE_UP - 1;
        let long = TRANSFER_FAILURES_GIVE_UP;

        assert!(
            engine_is_failing(true, false, 0, give_up),
            "an engine that has been up long enough with nothing on the screen is failing at the \
             file, which is what it was always asked about"
        );
        assert!(
            !engine_is_failing(true, false, 0, Duration::ZERO),
            "while one that has only just started is a file that is merely slow, and taking it \
             away would be a real loss"
        );

        // The case this is for. A session that has drawn a frame and cannot draw another keeps
        // the frame it had, and every question the app asks about it is answered well: the
        // picture on screen is the picture from three seconds ago and nothing is wrong with the
        // preview as far as anything outside can see.
        assert!(
            engine_is_failing(true, true, long, Duration::ZERO),
            "a session whose transfers keep being refused is failing whatever it has drawn"
        );

        // And what must not move with it. `drew` means a frame has been handed over, and the
        // reason is written on `failing_before_a_frame`: a film that played and then met a bad
        // sector is a film the probe got right about, and handing it to another engine over a
        // fault that has nothing to do with what plays it costs the engine its decoder for the
        // rest of a file it was getting on with.
        assert!(
            !engine_is_failing(true, true, short, fresh),
            "a run of refusals short of the bound is a hiccup, and a session that has drawn a \
             frame is not a session that cannot draw"
        );
        assert!(
            !engine_is_failing(true, true, 0, fresh),
            "and a session that is playing normally is not one no matter how long it has been up"
        );

        assert!(
            !engine_is_failing(false, false, long, fresh),
            "a sound has no frames to hand over and is never waiting for one, however it is asked"
        );
    }

    /// What this side costs to take a frame of a video: frames actually drawn, wall time, and
    /// the CPU the process burned while it did — against a real file, played at the box the
    /// layout would have placed it at.
    ///
    /// It is the measurement this module's own numbers are argued from, and it is here rather
    /// than in a harness of its own because the thing being measured is a thread-local session
    /// and a static this module owns; nothing outside can reach them. Three things are
    /// measured because a change to this path moves them independently: how many frames the
    /// caller was actually handed, how long the window took, and what the process's own clock
    /// says it spent. The first says whether the picture moved at all, the second whether the
    /// loop kept up with the file, and the third is the whole of the question this module
    /// exists to answer — a preview that costs a core of the CPU is a preview that makes every
    /// other hover on the desktop stutter.
    ///
    /// The box matters as much as the file. A 320 x 240 box measures the *engine* and nothing
    /// this side does: the copies, the resample and the hand to the compositor are all
    /// proportional to the number of pixels, so a small box hides the entire cost. The default
    /// is a 2560 x 1440 file at `fit` on this machine's display, which is a little over
    /// 2493 x 1400 — a box this side does *not* scale into (see `ENLARGEMENT_WORTH_RESAMPLING`),
    /// so what it measures is the copy and not the resampler. `RHP_VIDEO_PERF_BOX` overwrites it
    /// as `WIDTHxHEIGHT` for a machine whose display is another size.
    ///
    /// The window is eight seconds with the first one spent settling, which is where a first
    /// frame's decoder comes up and where a hardware pipeline's device is made; a run that
    /// started its clock at `Play` would be measuring setup rather than a preview. The tick is
    /// sixteen milliseconds, which is `preview_window::FRAME_WAIT_MS` — that constant belongs to
    /// the loop and is not this file's to read, and a measurement taken at any other cadence is
    /// not the loop being measured.
    ///
    /// The CPU is the process's own, read from `GetProcessTimes` on this process: kernel and
    /// user together, in 100-nanosecond units, over the same window. Shelling out to an external
    /// timer would be a second program and its own precision to measure a first.
    ///
    /// Run it by hand:
    ///
    /// ```text
    /// cargo test video_take_cost -- --ignored --nocapture
    /// ```
    ///
    /// The file it is most worth running on is one that decodes as a rate the display cannot
    /// show — a 144 fps file of near-duplicate frames — because that is where the difference
    /// between decoding on the GPU and decoding in software is the largest thing in the
    /// measurement, and where the waste of decoding frames nobody is shown is at its height.
    #[test]
    #[ignore = "plays the file named in RHP_VIDEO_PERF at the size of the display"]
    fn video_take_cost() {
        let Ok(path) = std::env::var("RHP_VIDEO_PERF") else {
            println!("set RHP_VIDEO_PERF to the path of a video to play");
            return;
        };

        let path = PathBuf::from(path);
        let (width, height) = box_from_env((2493, 1400));
        let picture = dimensions(&path).expect("the file's own frame size");

        println!("\n--- {} ---", path.display());
        println!(
            "the picture is {}x{}, shown in a box of {width}x{height}",
            picture.0, picture.1
        );

        // The crop the geometry probe would have settled on is not asked for: this is a
        // measurement of the frame pipeline and the file's own bars are its own business. A
        // file with a crop is measured a few thousand pixels narrower than this box, which is
        // what it is played at in the app too.
        play(
            &path,
            width,
            height,
            0,
            Picture {
                width: picture.0,
                height: picture.1,
                crop: None,
            },
        );

        assert!(is_playing(), "the engine took the file at all");
        println!(
            "this side is the one that scales into the box: {}",
            scales_here(picture, width, height)
        );

        // The first frame is waited for rather than assumed, because a file the engine cannot
        // draw has none to arrive and a window measured over a session that never drew is a
        // measurement of nothing (see `failing_path`).
        let mut pixels = Vec::new();
        let ticks = Duration::from_millis(TICK_MS);
        let waiting = Instant::now();
        let mut drew = false;

        while waiting.elapsed() < FIRST_FRAME_GIVE_UP {
            if copy_frame_into(&mut pixels).is_some() {
                drew = true;
                break;
            }

            std::thread::sleep(ticks);
        }

        if !drew {
            println!(
                "no frame in {} ms — nothing to measure",
                waiting.elapsed().as_millis()
            );
            stop();
            return;
        }

        // A second of the file is played before the clock is read, so that what is measured is
        // a preview and not a pipeline coming up: a first frame's decoder, and a device that
        // has not been made yet if there is a GPU path to make one.
        let settling = Instant::now();
        let mut settling_frames = 0;
        while settling.elapsed() < Duration::from_secs(1) {
            if copy_frame_into(&mut pixels).is_some() {
                settling_frames += 1;
            }

            std::thread::sleep(ticks);
        }

        // And what the process's own clock says before and after: the kernel's and the user's
        // together, since a decode that is handed to a driver runs partly in one and partly in
        // the other and neither on its own is the cost. A process whose own CPU time cannot be
        // read has no measurement to make, and says so rather than reporting a zero.
        let Some(before) = cpu_hundred_nanoseconds() else {
            println!("this process's own CPU time cannot be read — nothing to measure");
            stop();
            return;
        };

        let started = Instant::now();
        let mut frames = 0;
        let mut refused = 0;

        while started.elapsed() < WINDOW {
            if copy_frame_into(&mut pixels).is_some() {
                frames += 1;
            } else if failing_path().is_some() || !is_playing() {
                refused += 1;
            }

            std::thread::sleep(ticks);
        }

        let wall = started.elapsed();
        let cpu = cpu_hundred_nanoseconds().map_or(0, |now| now.saturating_sub(before));
        let seconds = wall.as_secs_f64();

        println!(
            "\n{frames} frames drawn in {:.2} s ({:.1} a second, over {settling_frames} more \
             while it settled)",
            seconds,
            frames as f64 / seconds,
        );
        println!(
            "process CPU {:.2} s over {:.2} s — {:.1}% of one core",
            cpu as f64 / 10_000_000.0,
            seconds,
            cpu as f64 / 10_000_000.0 / seconds * 100.0,
        );
        println!(
            "each frame cost {:.2} ms of CPU to hand over",
            if frames > 0 {
                cpu as f64 / 10_000_000.0 / frames as f64 * 1000.0
            } else {
                f64::NAN
            }
        );
        println!("{refused} ticks came back with nothing to draw");
        println!(
            "the session is still playing: {}",
            is_playing(),
        );

        stop();
    }

    /// How long the preview loop waits between its ticks, which is the cadence the measurement
    /// above is taken at: `preview_window::FRAME_WAIT_MS`, written out here because that
    /// constant belongs to the loop and this module does not read it.
    const TICK_MS: u64 = 16;

    /// How long `video_take_cost` measures a preview for, once it has settled: long enough that
    /// the answer is a number rather than a rounding, and short enough to run before a film is
    /// over — which matters, because a 144 fps file of eight seconds is over a thousand frames
    /// of decode and the file the measurement is most worth running on is twelve minutes long.
    const WINDOW: Duration = Duration::from_secs(8);

    /// The box a preview is played at, from `RHP_VIDEO_PERF_BOX` as `WIDTHxHEIGHT` or the
    /// default this module's own measurements have been taken at: a 1440p file at `fit` on a
    /// 2560-wide display, which is the box the layout hands `play` and therefore the size at
    /// which every copy and every hand to the compositor is paid for.
    fn box_from_env(default: (u32, u32)) -> (u32, u32) {
        let Ok(setting) = std::env::var("RHP_VIDEO_PERF_BOX") else {
            return default;
        };

        let Some((width, height)) = setting.split_once('x') else {
            println!("RHP_VIDEO_PERF_BOX is not WIDTHxHEIGHT — using {default:?}");
            return default;
        };

        match (width.trim().parse(), height.trim().parse()) {
            (Ok(width), Ok(height)) if width > 0 && height > 0 => (width, height),
            _ => {
                println!("RHP_VIDEO_PERF_BOX is not two whole numbers — using {default:?}");
                default
            }
        }
    }

    /// What this process has spent on the CPU, in 100-nanosecond units: kernel and user
    /// together, since a preview that hands its decoding to a driver spends the time in
    /// whichever of the two the driver spends it in.
    fn cpu_hundred_nanoseconds() -> Option<u64> {
        use windows::Win32::Foundation::FILETIME;
        use windows::Win32::System::Threading::{GetCurrentProcess, GetProcessTimes};

        let mut created = FILETIME::default();
        let mut exited = FILETIME::default();
        let mut kernel = FILETIME::default();
        let mut user = FILETIME::default();

        // SAFETY: all four out-parameters are the caller's own, initialised, and live for the
        // duration of the call; `GetCurrentProcess` is a pseudo-handle that is always valid and
        // is what the process's own times are read from. A refusal reads as no measurement
        // rather than as a zero, which would flatter everything this is used for.
        unsafe {
            GetProcessTimes(
                GetCurrentProcess(),
                &mut created,
                &mut exited,
                &mut kernel,
                &mut user,
            )
        }
        .ok()?;

        let as_units =
            |time: FILETIME| ((time.dwHighDateTime as u64) << 32) | time.dwLowDateTime as u64;

        Some(as_units(kernel) + as_units(user))
    }
}

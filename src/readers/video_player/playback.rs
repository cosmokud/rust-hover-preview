//! What the app asks of a running engine — start, stop, pause, seek, resize — and the file-level
//! probes that decide whether one is started at all. See the module above for the reasoning.

use super::frame_copy::{open_stream, surface, unpack_pair};
use super::session::{Session, Take};
use crate::formats::{codecs, head};
use once_cell::sync::Lazy;
use std::cell::RefCell;
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::Duration;
use windows::core::implement;
use windows::Win32::Media::MediaFoundation::{
    IMFAttributes, IMFMediaEngineNotify, IMFMediaEngineNotify_Impl, MFCreateAttributes,
    MFCreateMediaType, MFCreateSourceReaderFromByteStream, MFMediaType_Video, MFVideoFormat_RGB32,
    MFVideoNormalizedRect, MFARGB, MF_MEDIA_ENGINE_EVENT_ERROR,
    MF_MEDIA_ENGINE_EVENT_FIRSTFRAMEREADY, MF_MT_FRAME_SIZE, MF_MT_MAJOR_TYPE,
    MF_MT_PIXEL_ASPECT_RATIO, MF_MT_SUBTYPE, MF_SOURCE_READER_ENABLE_VIDEO_PROCESSING,
    MF_SOURCE_READER_FIRST_VIDEO_STREAM,
};

/// What the letterboxing is filled with, where a box has any letterboxing in it at all.
///
/// The box a preview is placed at is the shape of the picture that goes into it — the probe's
/// crop where it settled on one, the frame itself where it did not — so an ordinary preview is
/// filled to its own edges and this is never reached. What is left for it is the box that has
/// been given another shape since: a pinned window dragged by its edges, into which the engine
/// scales the picture and pads what is left. A video is an opaque rectangle — the window it
/// used to be played in was opaque too — so the padding is black rather than the backdrop the
/// rest of a frame is composited over.
pub(super) const BORDER: MFARGB = MFARGB {
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
pub(super) const SEEK_GIVE_UP: Duration = Duration::from_secs(3);

/// How long a session is given to hand over its first frame before the engine is taken to be
/// failing at the file rather than slow to start it.
///
/// Reading a file's first frame is a fraction of a second's work for a file on this machine, so
/// this is a give-up rather than a wait anything is expected to reach. It is generous on purpose:
/// what the answer does is hand the file to FFmpeg's player, and a file that was merely slow to
/// start would be a preview taken away from the engine that could have drawn it.
pub(super) const FIRST_FRAME_GIVE_UP: Duration = Duration::from_secs(3);

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
pub(super) const TRANSFER_FAILURES_GIVE_UP: u32 = 188;

/// Whether a run of transfers that all failed is a session that has failed, which is the whole
/// of what the count above is for.
///
/// A function of one number rather than of a session and a clock because the question has to be
/// answerable without one: the run *is* the evidence, and what it has to be judged against is a
/// count rather than a moment, since a session that took a seek to get here has no useful moment
/// to measure a gap against.
pub(super) fn a_run_of_refusals_is_a_failure(run: u32) -> bool {
    run >= TRANSFER_FAILURES_GIVE_UP
}

/// How much bigger than the picture a box has to be before this side is the one that scales into
/// it, as a share of the picture's own size: one part in fifty.
///
/// The choice is between the two engines, and they cost very different amounts. The engine
/// scales on the GPU as part of the frame it is already writing, so its share of the work is
/// nothing this app can measure; this side's share is a scalar bilinear over every pixel of
/// every frame, on the preview thread (see [`scale_rows`](super::frame_copy::scale_rows)). So the answer to "who is bigger" has
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
/// Two events are acted on and the rest are thrown away: an engine that reports an error has
/// nothing left to hand over, and what this side does about that is stop asking it for
/// frames; an engine that reports its first frame has a picture to give, which is the only
/// thing that makes a frame transfer worth asking for. The callback arrives on a thread of the
/// engine's own, so what it touches is an atomic and nothing else.
#[implement(IMFMediaEngineNotify)]
pub(super) struct Notify {
    pub(super) failed: Arc<AtomicBool>,
    /// That the engine has decoded a frame of the file and handed it over, which is its own word
    /// about the file and the one thing here that means a picture exists to be taken.
    pub(super) first_frame: Arc<AtomicBool>,
}

impl IMFMediaEngineNotify_Impl for Notify_Impl {
    fn EventNotify(&self, event: u32, param1: usize, param2: u32) -> windows::core::Result<()> {
        // The failure the app acts on first: an error is the flag being set, and there is
        // nothing for this engine to hand over after that.
        if event == MF_MEDIA_ENGINE_EVENT_ERROR.0 as u32 {
            let _ = (param1, param2);
            self.failed.store(true, Ordering::Release);
        }

        // The engine saying it has decoded a frame of the file and handed it over, which is the
        // only answer that means there is a picture here to take. Everything above this line is
        // about a file the engine cannot play; this is about one it can.
        if event == MF_MEDIA_ENGINE_EVENT_FIRSTFRAMEREADY.0 as u32 {
            self.first_frame.store(true, Ordering::Release);
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
pub(super) fn source_rect(crop: Crop) -> Option<MFVideoNormalizedRect> {
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
/// asked for at the picture's own size and this side scales (see [`scale_rows`](super::frame_copy::scale_rows)), while a box
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
pub(super) fn scales_here(picture: (u32, u32), width: u32, height: u32) -> bool {
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
///
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
pub(super) fn engine_is_failing(surface: bool, drew: bool, refusals: u32, age: Duration) -> bool {
    surface && (a_run_of_refusals_is_a_failure(refusals) || (!drew && age >= FIRST_FRAME_GIVE_UP))
}

/// The file the engine is failing at before it has drawn a frame of it, if it is failing at one
/// there: the engine saying so, on a window that has been put up for the file and is standing
/// there until something else is put up instead.
///
/// What it is read is the one thing the engine can be asked that means failure — an
/// `MF_MEDIA_ENGINE_EVENT_ERROR` raised into [`Notify`], the engine admitting outright that it
/// cannot play the file it was handed rather than taking its time over it, and worth acting on
/// the tick it arrives rather than a second later.
///
/// There is deliberately no time floor beside it. There was [`FIRST_FRAME_GIVE_UP`], and it was
/// there for a silence: a session with a surface that had drawn nothing for three seconds was
/// read as the engine with nothing to report because it has no pipeline to report on (see
/// [`failing_path`]) — the same failure arrived at by not-arrival rather than by an error. But
/// a file being slow to decode is not that, and a cold read of a large one can exceed three
/// seconds of opening before there is a frame in it at all: on that file the floor fired, the
/// engine was taken away over the disk's speed, and the pin was left with a preview of the
/// placeholder — the very flash the frame gate exists to prevent. A wait this long is better
/// ended by the engine's own word or by the user closing the pin or stepping to another file
/// than by a clock that cannot tell a slow file from a broken one.
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

        // There is no time floor here, and the reason is the file rather than the engine: a
        // pinned window is put up for this file and stands there until something else is put up
        // instead, so a wait that ends by a clock is a file called unplayable because this machine
        // read it slowly. A cold read of a large file is exactly that — seconds of opening it
        // before there is a frame in it — and giving up on the clock hands it to FFmpeg's player
        // over a fault that is the disk's and not the codec's. What ends the wait is the engine
        // saying so, or the user closing the pin or stepping to another file.
        (session.bitmap.is_some() && !session.drew && session.failed.load(Ordering::Acquire))
            .then(|| session.path.clone())
    })
}

/// Write a file down as one the media engine cannot draw, and let go of the session that was
/// failing at it, so that what plays the file from here on is FFmpeg's player.
///
/// On a machine that has FFmpeg's player this has almost nothing left to correct, because that
/// player was already what the file was routed to: the engine is only playing a video there
/// because it is not installed at all. What this still settles on such a machine is the audio,
/// where the engine is asked first for a format it can decode and this is the correction when it
/// then failed anyway — and it is the correction for a *video* on a machine with no FFmpeg, where
/// the file stops having any preview rather than gaining one (see
/// [`plays_video_natively`](crate::formats::codecs::plays_video_natively)).
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
///
/// It is asked for every film the route would draw with the media engine — which, since the engine
/// is a setting, is a small film under the hybrid, every film under an explicit `Native`, and every
/// film on a machine with no `ffplay` at all; a film FFmpeg's player takes never opens the file
/// (see `video_hw::resolve_video_engine` and
/// [`plays_video_natively`](crate::formats::codecs::plays_video_natively)).
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
/// with the picture the caller already has (see [`is_a_new_frame`](super::session::is_a_new_frame)). It is also wider than the
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

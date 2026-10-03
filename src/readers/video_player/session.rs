//! One playing file: the session the engine plays it in, the seek it is still owed, and the one
//! place a frame is ever asked of the engine. See the module above for the reasoning.

use super::frame_copy::{copy_locked, open_stream, resample_locked, surface};
use super::playback::{
    a_run_of_refusals_is_a_failure, scales_here, source_rect, Notify, Picture, BORDER, SEEK_GIVE_UP,
};
use crate::paths::plain_path;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Arc;
use std::time::Instant;
use windows::core::{IUnknown, Interface, BSTR};
use windows::Win32::Foundation::RECT;
use windows::Win32::Graphics::Imaging::IWICBitmap;
use windows::Win32::Media::MediaFoundation::{
    CLSID_MFMediaEngineClassFactory, IMFAttributes, IMFByteStream, IMFMediaEngine,
    IMFMediaEngineClassFactory, IMFMediaEngineEx, MFCreateAttributes, MFVideoFormat_ARGB32,
    MFVideoNormalizedRect, MF_MEDIA_ENGINE_CALLBACK, MF_MEDIA_ENGINE_READY_HAVE_CURRENT_DATA,
    MF_MEDIA_ENGINE_READY_HAVE_METADATA, MF_MEDIA_ENGINE_VIDEO_OUTPUT_FORMAT,
};
use windows::Win32::System::Com::{CoCreateInstance, CLSCTX_INPROC_SERVER};

/// One playing video: the engine, the surface its frames are delivered into, and what the
/// surface is currently the size of.
pub(super) struct Session {
    pub(super) engine: IMFMediaEngine,
    /// The stream the engine was handed, kept until the video is over so that nothing it is
    /// still reading from can go out of scope under it.
    pub(super) byte_stream: IMFByteStream,
    /// Where a video's frames are delivered, and nothing at all for a sound — which has no
    /// frames to deliver and no box to draw one in.
    pub(super) bitmap: Option<IWICBitmap>,
    /// The box the preview draws at: the size of the frame this side hands over, and the size
    /// the engine is asked to deliver at only while it is the one scaling (see `scaled`).
    pub(super) width: u32,
    pub(super) height: u32,
    /// The size the picture has in the file, which is what the engine is asked for where this
    /// side scales it.
    pub(super) picture: (u32, u32),
    /// Whether this side scales the picture into the box rather than the engine: the two sizes
    /// above are what settles it, and what each of the two costs is written on [`play`](super::playback::play).
    pub(super) scaled: bool,
    /// The two rows of the picture already read across to the box's width, kept between the
    /// rows of the box mixed from them, and empty while the engine is the one scaling.
    rows: [Vec<u8>; 2],
    /// Where in the frame the picture is taken from, where the probe settled on a crop: the
    /// rectangle the engine's frame transfer is asked for. The whole frame is `None`, which is
    /// a file the probe found no crop in and every sound.
    source: Option<MFVideoNormalizedRect>,
    pub(super) path: PathBuf,
    pub(super) failed: Arc<AtomicBool>,
    /// The engine's own word that it decoded a frame of the file and handed it over. It is a
    /// different and a stronger statement than [`Self::drew`], which records only that a
    /// transfer was *asked for*: a surface the engine has not drawn into is a bitmap of zeros,
    /// and a frame taken from one is a picture of nothing (see `copy_into`).
    first_frame: Arc<AtomicBool>,
    /// Whether a frame of the file has been handed over at all: a session is not a promise that
    /// a picture will come of it, since the engine accepts a file whose decoder and converter
    /// are both there and whose *pipeline* is not (see [`failing_path`](super::playback::failing_path)).
    ///
    /// It says whether a frame has *ever* been handed over and is not cleared by a transfer that
    /// fails after one, because that is a different fault: a file that played and then met a bad
    /// sector is not a file the engine cannot draw, and handing it to FFmpeg's player over one
    /// costs the engine its decoder for the rest of a film it was getting on with (see
    /// [`transfer_failures`] and [`failing_before_a_frame`](super::playback::failing_before_a_frame)).
    pub(super) drew: bool,
    /// How many transfers in a row the engine has refused, reset by the one that succeeds.
    ///
    /// This is what distinguishes a video that is not moving from a video this side cannot get
    /// frames out of, and the distinction is invisible from outside: both answer the loop's tick
    /// with no frame and neither reports anything wrong. The run is bounded and reaching
    /// [`TRANSFER_FAILURES_GIVE_UP`](super::playback::TRANSFER_FAILURES_GIVE_UP) of it is the session failing, which is what makes this
    /// something the app can be told about rather than a picture that freezes.
    pub(super) transfer_failures: u32,
    /// The time of the frame the caller is holding, in the hundred-nanosecond units the engine
    /// ticks in, and `None` where the next frame is owed whatever the engine says about it — a
    /// session that has just been begun, resized or sought somewhere has a picture the caller
    /// has not seen, and a seek in particular can land on the very time it has just drawn
    /// (see [`is_a_new_frame`]).
    pub(super) drawn: Option<i64>,
    /// Where the sound was asked to start, while the engine has not taken it there yet: `Load`
    /// answers before the header is there, so the position is kept and made on the first tick
    /// that finds the engine loaded (see `apply_seek`).
    pending_seek: Option<f64>,
    /// When the session was started, which bounds the wait for a pending seek: an engine that
    /// has not read the header of its file in this long is not going to, and a seek left
    /// standing is a seek made on every read of the clock after it.
    pub(super) began: Instant,
}

impl Session {
    pub(super) fn begin(
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
        let first_frame = Arc::new(AtomicBool::new(false));

        let mut attributes: Option<IMFAttributes> = None;
        unsafe { MFCreateAttributes(&mut attributes, 3) }.ok()?;
        let attributes = attributes?;

        let notify: IUnknown = Notify {
            failed: Arc::clone(&failed),
            first_frame: Arc::clone(&first_frame),
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
            first_frame,
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
    pub(super) fn seek(&mut self, seconds: f64) {
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
    pub(super) fn take_pending(&mut self) {
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
    pub(super) fn copy_into(&mut self, pixels: &mut Vec<u8>) -> Take {
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

        // And ready is not the same as decoded. `HAVE_CURRENT_DATA` says a frame is queued for
        // the *stream*, which on a cold read is true before the engine has drawn a single pixel
        // of the file — and the surface it is delivered into is a bitmap of zeros until it does,
        // so a transfer asked for in that window succeeds and hands over a rectangle of black.
        // The engine's own report of a first frame is the only thing that says a picture exists,
        // and until it is made there is nothing here to take: the engine is not ticked, no
        // transfer is asked for, nothing is written to the caller's pixels and the session is not
        // recorded as having drawn.
        if !self.first_frame.load(Ordering::Acquire) {
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
    /// of what makes the fault visible outside: the flag is what
    /// [`is_playing`](super::playback::is_playing) and both of the
    /// failing paths read, so a session whose transfers keep being refused is answered exactly
    /// as one whose engine gave an error event — which is the truth, and the reason the three
    /// seconds after which the app hands the file to FFmpeg's player arrives on a schedule
    /// rather than never.
    pub(super) fn refused(&mut self) -> Take {
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
pub(super) enum Take {
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
pub(super) fn is_a_new_frame(drawn: Option<i64>, pts: i64) -> bool {
    drawn != Some(pts)
}

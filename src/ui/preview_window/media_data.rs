//! One media: what is on screen right now, the frames it is made of, and the geometry cache
//! the films were probed into.

use super::*;

pub(super) const VIDEO_GEOMETRY_CACHE_MAX_ENTRIES: usize = 512;
pub(super) static CURRENT_MEDIA: Lazy<Mutex<Option<MediaData>>> = Lazy::new(|| Mutex::new(None));
/// What a probed geometry is only valid for: the file and the version of it that
/// was probed, so a video replaced in place is probed again rather than cropped and
/// sized by the answer about the file it used to be — the same rule every other
/// held picture in this module follows.
#[derive(Clone, PartialEq, Eq, Hash)]
pub(super) struct VideoGeometryKey {
    pub(super) path: PathBuf,
    pub(super) version: FileVersion,
}

/// Not `Copy`, for the same reason the geometry it holds is not: the sidecar in there is a
/// path, so every read out of the cache is a clone of the answer rather than a copy of it
/// (see `cached_video_geometry`).
#[derive(Clone)]
pub(super) enum ProbedGeometry {
    /// The shape the probe read, and the crop the detector settled on.
    Measured(VideoGeometry),
    /// The answer that there is nothing to measure: a file neither FFmpeg nor the media
    /// engine will open. It is an answer like any other and is held like one, so a file
    /// that cannot be measured is not probed again on every hover — and so the hover
    /// that is waiting for a probe can be told that the probe is done (see `video_box`).
    Unmeasurable,
}

pub(super) static VIDEO_GEOMETRY_CACHE: Lazy<Mutex<HashMap<VideoGeometryKey, ProbedGeometry>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// The geometry cache's own guard, a poisoned lock included.
///
/// What is behind it is a map of facts — a shape, a crop, the answer that there is neither —
/// and there is no invariant a panic could have left half applied, so a lock a panicked probe
/// poisoned is one to read anyway rather than one that reads as an empty cache for the rest
/// of the run: an empty cache is `video_probe_due` true on every hover, which is a probe
/// started, and waited on, for every file that is hovered (see `hidden_epoch` for the same
/// reading of a lock that holds a fact).
pub(super) fn video_geometry_cache(
) -> MutexGuard<'static, HashMap<VideoGeometryKey, ProbedGeometry>> {
    match VIDEO_GEOMETRY_CACHE.lock() {
        Ok(cache) => cache,
        Err(poisoned) => poisoned.into_inner(),
    }
}

/// Media data that can be either static or animated
pub(super) struct MediaData {
    /// Frames behind an `Arc` so a frame handed over by the image cache is the same
    /// allocation rather than a copy of it — a still is decoded once and then shared by
    /// every hover that shows it again (see `ImageCacheEntry`).
    ///
    /// The two places that write into a frame in place ask for it mutably, which copies
    /// only if something else is still holding it: a loading frame and a native video
    /// frame, neither of which is a frame the cache has.
    pub(super) frames: Vec<Arc<ImageFrame>>,
    /// Shared frame queue for streaming decode (animated formats append here)
    pub(super) shared_frames: Option<Arc<Mutex<StreamedFrames>>>,
    /// Signal from the background thread that all frames have been decoded
    pub(super) all_frames_loaded: Option<Arc<AtomicBool>>,
    pub(super) current_frame: usize,
    pub(super) last_frame_time: Instant,
    pub(super) media_type: MediaType,
    /// Cancellation token for background decode work.
    pub(super) stream_cancel: Option<Arc<AtomicBool>>,
    // For video playback using ffplay
    pub(super) video_process: Option<Child>,
    pub(super) loading_start: Option<Instant>,
    /// Where a text preview is scrolled to, when it scrolls at all.
    pub(super) text_state: Option<TextPreviewState>,
}

/// Paint a band of a layered window one flat opaque colour, for a window whose media is not there.
///
/// A layered window's hit testing and its compositing are both answered by the shape of its
/// pixels, which is why a video pin leaves its middle band transparent for another process's
/// window to show through. That arrangement has one failure mode and this is it: hide the player
/// and the band stops being a hole for something and becomes a hole, with the desktop visible
/// through a window the hand is dragging.
///
/// The alpha is forced opaque rather than blended, because a translucent fill over nothing is the
/// same hole at lower contrast, and because the point of the band while a drag lasts is to stop
/// being a picture at all.
pub(super) fn fill_band_opaque(out: &mut [u8], width: u32, origin_y: u32, height: u32) {
    let stride = width as usize * 4;
    if stride == 0 {
        return;
    }

    let rows = out.len() / stride;
    let first = (origin_y as usize).min(rows);
    let last = (origin_y as usize + height as usize).min(rows);
    for row in first..last {
        let Some(pixels) = out.get_mut(row * stride..row * stride + stride) else {
            return;
        };
        for pixel in pixels.chunks_exact_mut(4) {
            pixel.copy_from_slice(&[0, 0, 0, 255]);
        }
    }
}

impl MediaData {
    /// The frame on screen, which a document the engine draws does not have.
    pub(super) fn current_frame(&self) -> Option<&ImageFrame> {
        self.frames.get(self.current_frame).map(AsRef::as_ref)
    }

    pub(super) fn current_pixels(&self) -> &[u8] {
        self.current_frame()
            .map(|frame| frame.pixels.as_slice())
            .unwrap_or(&[])
    }

    /// Repaint a sound's card with the clock as it stands, answering whether the frame on
    /// screen changed.
    ///
    /// A painted preview is drawn once and held, so the two things about a card that move are
    /// drawn by asking for the page again — the same arrangement a text preview's scrolling has
    /// (see `repaint_text_preview`), and what keeps the card's own layout in one place: the box
    /// it was painted in is the box it is painted in again, and only the clock, the bar under
    /// it, the scroll of a name the card has no room for and what the card's own controls are
    /// saying differ.
    ///
    /// `chrome` carries the whole of that last one, the loop's `playing` included — it is read in
    /// one look by `pinned_audio_chrome` rather than handed over beside itself, because both
    /// callers of this are reached with the media's own lock held and a pin's lock asked from
    /// under it is the re-entrancy `toggle_pinned_playback` is written down for.
    pub(super) fn refresh_audio_card(
        &mut self,
        path: &Path,
        elapsed: Option<f64>,
        duration: Option<f64>,
        dpi: u32,
        name_offset: i32,
        chrome: Option<CardChrome>,
    ) -> bool {
        let Some(frame) = self.frames.first() else {
            return false;
        };
        let (width, height) = (frame.width, frame.height);

        let Some(card) = audio_card(path, elapsed, duration, name_offset, chrome) else {
            return false;
        };
        let Some((pixels, width, height)) =
            audio_preview::render(&card, width, height, dpi, current_audio_options())
        else {
            return false;
        };

        self.frames[0] = Arc::new(ImageFrame::new(pixels, width, height, 0));
        true
    }

    /// Paint the card again at a box of a different size: what a pinned preview of a sound
    /// costs when its window is maximized or resized. It is the question `refresh_audio_card`
    /// answers asked of the box the card is being given rather than of the one it already has —
    /// and it is handed the loop's own clock whole (`AudioCardClock`), which is the same thing a
    /// repaint reads out of it (see `relayout_pinned_media`).
    pub(super) fn relayout_audio_card(
        &mut self,
        path: &Path,
        clock: AudioCardClock,
        elapsed: Option<f64>,
        duration: Option<f64>,
        size: (u32, u32),
    ) -> bool {
        let Some(card) = audio_card(path, elapsed, duration, clock.name_offset, clock.chrome)
        else {
            return false;
        };
        let Some((pixels, width, height)) = audio_preview::render(
            &card,
            size.0.max(1),
            size.1.max(1),
            clock.dpi,
            current_audio_options(),
        ) else {
            return false;
        };

        self.frames[0] = Arc::new(ImageFrame::new(pixels, width, height, 0));
        true
    }

    pub(super) fn current_width(&self) -> u32 {
        self.current_frame().map(|frame| frame.width).unwrap_or(0)
    }

    pub(super) fn current_height(&self) -> u32 {
        self.current_frame().map(|frame| frame.height).unwrap_or(0)
    }

    /// Whether the frame on screen is opaque everywhere, which is what lets a repaint copy
    /// it rather than blend it (see `ImageFrame::opaque`).
    pub(super) fn current_frame_is_opaque(&self) -> bool {
        self.current_frame().is_some_and(|frame| frame.opaque)
    }

    /// Check if all frames have finished streaming
    pub(super) fn is_fully_loaded(&self) -> bool {
        match &self.all_frames_loaded {
            Some(flag) => flag.load(Ordering::Acquire),
            None => true, // No streaming = already complete
        }
    }

    /// Pull newly decoded frames from the shared buffer, then give back the
    /// frames that have already been played.
    ///
    /// Frames are only taken while the retained window has room: the decoder
    /// waits once its queue is full, so this is what keeps a long animation from
    /// decoding itself into memory faster than it is shown. The window stops
    /// meaning anything once the decoder is done with the file — there is nothing
    /// left to hold back, and the frames it finished with are frames of the
    /// animation whatever is in hand.
    pub(super) fn sync_shared_frames(&mut self) {
        let Some(shared) = self.shared_frames.clone() else {
            return;
        };

        let retained_bytes: usize = self.frames.iter().map(|frame| frame.pixels.len()).sum();
        if retained_bytes < ANIMATION_RETAINED_BYTES || self.is_fully_loaded() {
            let result = shared.lock();
            if let Ok(mut streamed) = result {
                if !streamed.queue.is_empty() {
                    // A streamed frame is made by the decoder thread and handed over once,
                    // so wrapping it here is the only allocation it ever needs.
                    self.frames.extend(streamed.queue.drain(..).map(Arc::new));
                }
            }
        }

        self.release_played_frames(retained_bytes);
    }

    /// Whether the player has given back frames it already showed.
    pub(super) fn frames_were_released(&self) -> bool {
        self.shared_frames
            .as_ref()
            .and_then(|shared| shared.lock().ok().map(|streamed| streamed.released))
            .unwrap_or(false)
    }

    /// Give back the frames behind the playhead once enough of them have piled up.
    /// Without this a long animation would either stop part-way at a fixed size
    /// cap or keep its whole decoded length in memory; with it, playback stays
    /// inside a fixed window while the decoder replays the file to loop.
    ///
    /// Two things have to hold before a frame is given back, and both of them are
    /// about a promise the player makes to itself: the frames that are dropped are
    /// frames that have to be decoded again. The window has to be full, which is
    /// the only thing a release is for — an animation that fits in
    /// `ANIMATION_RETAINED_BYTES` is held whole, plays from beginning to end and
    /// wraps back into the frame it started on, and the file is read once for all
    /// of it. And the decoder has to still be working, which is settled under the
    /// same lock the decoder takes as it ends a pass: a decoder that has finished
    /// with a file it decoded whole has nothing to decode again.
    pub(super) fn release_played_frames(&mut self, retained_bytes: usize) {
        let Some(shared) = self.shared_frames.clone() else {
            return;
        };

        let keep_from = self.current_frame.saturating_sub(1);
        if keep_from == 0 {
            return;
        }

        if retained_bytes < ANIMATION_RETAINED_BYTES {
            return;
        }

        let played_bytes: usize = self.frames[..keep_from]
            .iter()
            .map(|frame| frame.pixels.len())
            .sum();
        if played_bytes < ANIMATION_RELEASE_BYTES {
            return;
        }

        let Ok(mut streamed) = shared.lock() else {
            return;
        };
        if streamed.decoded {
            return;
        }

        self.frames.drain(..keep_from);
        self.current_frame -= keep_from;

        streamed.released = true;
    }

    pub(super) fn advance_frame(&mut self) -> bool {
        // Pull in any new frames from streaming decode
        self.sync_shared_frames();

        let frame_count = self.frames.len();
        if frame_count <= 1 {
            return false;
        }

        let fully_loaded = self.is_fully_loaded();
        let mut advanced = false;

        // Allow skipping multiple frames per call to keep up with real time.
        for _ in 0..frame_count {
            let delay = Duration::from_millis(effective_frame_delay_ms(
                &self.media_type,
                self.frames[self.current_frame].delay_ms,
            ) as u64);
            if self.last_frame_time.elapsed() >= delay {
                let next = self.current_frame + 1;
                if next < frame_count {
                    // More decoded frames ahead — advance normally
                    self.current_frame = next;
                    self.last_frame_time += delay;
                    advanced = true;
                } else if fully_loaded && !self.frames_were_released() {
                    // All frames decoded and still in memory — safe to loop back
                    // to start
                    self.current_frame = 0;
                    self.last_frame_time += delay;
                    advanced = true;
                } else {
                    // Still streaming, or the start of the animation has been
                    // released to stay inside the memory window: hold this frame
                    // until the next one arrives. Keep the next streamed frame
                    // immediately eligible instead of adding another full-frame
                    // delay.
                    self.last_frame_time = Instant::now()
                        .checked_sub(delay)
                        .unwrap_or_else(Instant::now);
                    break;
                }
            } else {
                break;
            }
        }

        // The playhead is snapped forward when it is late by more than a second
        // past the delay the frame it is on asks for — a machine that slept, a
        // decode that took far longer than the frame it was for — so that catching
        // up cannot run on across several loops.
        //
        // The frame's own delay is part of what late means, and leaving it out is
        // what a frame held for a long time used to cost: a GIF frame that waits
        // two seconds is a frame waiting for two seconds and not a playhead that
        // has fallen behind, so a snap on a flat second reset the clock every tick
        // and a delay of a whole second or more could never be reached — an
        // animation whose first frame is held for 1.2 seconds sat on that frame
        // for good.
        let waiting_for = Duration::from_millis(effective_frame_delay_ms(
            &self.media_type,
            self.frames[self.current_frame].delay_ms,
        ) as u64)
            + Duration::from_secs(1);
        if self.last_frame_time.elapsed() > waiting_for {
            self.last_frame_time = Instant::now();
        }

        advanced
    }

    /// Returns true if this media is an animation still being decoded
    pub(super) fn is_streaming(&self) -> bool {
        matches!(
            self.media_type,
            MediaType::AnimatedGif
                | MediaType::AnimatedApng
                | MediaType::AnimatedWebP
                | MediaType::AnimatedHeif
                | MediaType::AnimatedJxl
        ) && !self.is_fully_loaded()
    }

    pub(super) fn should_draw_streaming_overlay(&self) -> bool {
        if !self.is_streaming() || self.frames.len() > 1 {
            return false;
        }

        self.loading_start
            .map(|s| s.elapsed() <= Duration::from_millis(STREAMING_SPINNER_MAX_MS))
            .unwrap_or(false)
    }

    pub(super) fn update_loading_frame(&mut self) -> bool {
        if !matches!(self.media_type, MediaType::Loading) {
            return false;
        }
        if self.last_frame_time.elapsed() >= Duration::from_millis(33) {
            if !self.frames.is_empty() {
                let width = self.frames[0].width;
                let height = self.frames[0].height;
                if let Some(start) = self.loading_start {
                    let elapsed_secs = start.elapsed().as_secs_f32();
                    let angle = elapsed_secs * 2.0 * std::f32::consts::PI * 1.2;
                    Arc::make_mut(&mut self.frames[0])
                        .set_pixels(render_loading_frame(width, height, angle));
                }
            }
            self.last_frame_time = Instant::now();
            return true;
        }
        false
    }

    pub(super) fn cancel_background_work(&mut self) {
        if let Some(flag) = self.stream_cancel.take() {
            flag.store(true, Ordering::Release);
        }
    }

    /// Take the frame the media engine has ready into the one frame a preview of this kind
    /// is composed of, answering whether there was one to take.
    ///
    /// A video played by the media engine is not a file that is decoded once and drawn
    /// again — it is a new picture every frame — so the frame the placeholder was made of
    /// is the frame every one after it lands in. That is also what keeps a playing video
    /// from allocating: the buffer is written into rather than replaced, and the frames
    /// that arrive between repaints are simply the ones that are never seen.
    pub(super) fn take_native_video_frame(&mut self) -> bool {
        // Asked for mutably rather than reached into through the `Arc`: a native video frame
        // is written over and over, so this is the one place the frame is genuinely owned
        // alone and the borrow is not a copy.
        let Some(frame) = self.frames.first_mut().map(Arc::make_mut) else {
            return false;
        };

        let Some((width, height)) = video_player::copy_frame_into(&mut frame.pixels) else {
            return false;
        };

        frame.width = width;
        frame.height = height;
        // Every pixel the copy wrote was forced opaque, so the frame a video lands in is one
        // a repaint can copy rather than blend — which is the whole of what makes a video at
        // the size of the display affordable to draw sixty times a second (see
        // `video_player::copy_locked`). It is also now the only way this flag is ever set:
        // the copy above answers with nothing unless the engine has reported a frame of its
        // own, so a surface it has not drawn into never reaches the answer, and there is no
        // zeroed rectangle here for a painter to read as a picture.
        frame.opaque = true;

        true
    }
}

/// The theme, Markdown rendering and font size the configuration currently selects, read once
/// per hover so a measure and the render that follow agree.
///
/// Whether a text preview is in full mode is not one of the configuration's answers: it is what
/// being pinned *means* for one. A hover is something to read and then to leave; a window the
/// user put on the screen with a key and a caption is somewhere to work, and full mode is what
/// makes one of those out of the other — the scrollbar the whole document is reachable through,
/// the text a selection can be made of, and the keys that put it on the clipboard.
pub(super) fn current_text_options() -> TextPreviewOptions {
    let full_mode = pinned();

    CONFIG
        .lock()
        .map(|cfg| TextPreviewOptions {
            theme: cfg.theme,
            markdown_mode: cfg.markdown_mode,
            font_scale_percent: cfg.text_font_scale_percent,
            full_mode,
        })
        .unwrap_or(TextPreviewOptions {
            theme: TextTheme::Light,
            markdown_mode: MarkdownMode::Rendered,
            font_scale_percent: DEFAULT_TEXT_FONT_SCALE_PERCENT,
            full_mode,
        })
}

/// The options an archive preview is laid out and painted with, read from the
/// configuration the way a text preview's are.
pub(super) fn current_archive_options() -> ArchivePreviewOptions {
    CONFIG
        .lock()
        .map(|cfg| ArchivePreviewOptions {
            theme: cfg.theme,
            font_scale_percent: cfg.text_font_scale_percent,
        })
        .unwrap_or(ArchivePreviewOptions {
            theme: TextTheme::Light,
            font_scale_percent: DEFAULT_TEXT_FONT_SCALE_PERCENT,
        })
}

/// Whether the preview on screen is one whose appearance is baked into its
/// painted frame rather than recomposited from shared pixels.
pub(super) fn current_media_is_painted() -> bool {
    CURRENT_MEDIA
        .lock()
        .map(|media| {
            media
                .as_ref()
                .map(|media| media.media_type.is_painted())
                .unwrap_or(false)
        })
        .unwrap_or(false)
}

/// Which kind of preview is on screen, if one is.
///
/// The media the renderer built already knows what it is, so this is the kind of
/// preview that is up rather than a classification of the file it came from. A
/// spinner is no kind: what it is standing in for has not been decided yet.
pub(super) fn current_media_kind() -> Option<PreviewType> {
    CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| media.as_ref().map(|media| media.media_type.kind()))
        .flatten()
}

/// How long a preview waits for a page before it stops waiting. An engine can be
/// held by a dialog inside Office, and a spinner that never ends is worse than
/// the picture the document saved — so past this the preview comes down and the
/// file is left alone.
pub(super) const OFFICE_RENDER_WAIT_SECS: u64 = 25;

/// How long the pointer has to have been on a file before an engine is asked to come up for it.
///
/// An engine's launch is a second or more, and an engine that is up is what the first hover of a
/// session on a document of its kind is otherwise paying for: the ask is made while the hover
/// waits, so that the start overlaps what it is waiting for rather than following it (see
/// `warm_engines_for`).
///
/// What the wait is for is the pointer rather than the engine. A hand crossing a folder is on a
/// new file every few dozen milliseconds, and what a hover is about is the file it comes to rest
/// on: an engine asked for on every file a sweep touches would be a machine full of
/// applications for a folder nobody looked at. A sixth of a second is long enough for a hand
/// that was going somewhere else to be somewhere else, and short enough that a hand that stopped
/// has the engine starting before it has finished looking at the file.
pub(super) const WARM_SETTLE_MS: Duration = Duration::from_millis(150);

/// An SVG document as `MediaData`: a kind, and no frame at all, because the engine draws
/// it in a window of its own.
///
/// This is what the loader answers with for a document, and the install path reads it as
/// the signal to hand the hover over: nothing of this app's goes on screen for one, so
/// there is no frame to install and no size to place it by.
///
/// A page of HTML the engine draws is the same kind of document as far as this is
/// concerned, and this is what it is answered with too: the one kind carries both, and the
/// engine's window is the preview either way (see `html_is_engine_drawn`, which is what
/// decides between this and a painted text preview).
pub(super) fn engine_svg_media() -> MediaData {
    MediaData {
        frames: Vec::new(),
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::EngineSvg,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
    }
}

/// A font as `MediaData`: a kind and nothing else, the same shape as a document — the
/// specimen is the engine's to draw, so there is no frame of this app's to install.
pub(super) fn engine_font_media() -> MediaData {
    MediaData {
        media_type: MediaType::EngineFont,
        ..engine_svg_media()
    }
}

/// The media a pinned window is given for a file the engine draws, when the hover it was
/// taken up from left none behind: the same kind, carrying the same nothing.
///
/// A document, a specimen and a page of HTML are shown in the engine's own window, so the
/// handover that puts one on screen deliberately leaves the media slot empty (see the loop's
/// own handover) — and a pin taken up over one finds that emptiness rather than a frame. What
/// is needed is the *kind* and nothing more: the pin reads it to know what shape the window is
/// framed by and what chrome it carries, and asks the engine itself whether the thing it is a
/// window onto is still there (see `pin_media_is_alive`).
///
/// `None` for a file this app draws itself, which is the whole of what this answers: such a
/// file has a frame of its own in the slot already, and installing an engine kind over it
/// would take a picture's window for a document's.
pub(super) fn pinned_engine_media(path: &Path) -> Option<MediaData> {
    match engine_kind_of(path)? {
        PreviewType::Vector | PreviewType::Text => Some(engine_svg_media()),
        PreviewType::Fonts => Some(engine_font_media()),
        _ => None,
    }
}

/// A still image as `MediaData`: one frame, nothing streaming.
///
/// The kind arrives with the frame rather than being decided here: a texture is a still
/// picture like any other, and the one thing that makes it a kind of its own is what it is
/// drawn over — which is the loader's to know, since the loader is what held the file.
pub(super) fn static_image_media(frame: Arc<ImageFrame>, kind: MediaType) -> MediaData {
    MediaData {
        frames: vec![frame],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: kind,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
    }
}

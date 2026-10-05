//! The shapes a preview arrives in: what a message asks for, what kind of file it is, the
//! frames a decoded picture is made of, and the geometry a film was probed at.

use super::*;

/// Budget for the frames one animation keeps decoded. An animation that fits is
/// decoded once and loops from memory; a larger one plays through a sliding
/// window and its decoder starts over when the animation wraps.
pub(super) const ANIMATION_RETAINED_BYTES: usize = 128 * 1024 * 1024;
/// Frames the streaming decoder may keep queued ahead of playback before it
/// waits, so decoding can never run away from what is on screen.
pub(super) const ANIMATION_QUEUE_FRAMES: usize = 6;
/// How much already-played footage piles up behind the playhead before it is
/// released. Releasing in blocks keeps the sliding window from moving per frame.
pub(super) const ANIMATION_RELEASE_BYTES: usize = 16 * 1024 * 1024;
/// Floor for an animation frame's own delay: a file saying zero — or a still
/// timebase, which is zero too — is a frame the playhead would advance through
/// as fast as the message pump allows, spinning the render loop on a picture
/// that never appears to change. 16 ms lets authored-60fps files play at speed;
/// anything slower is untouched.
pub(super) const MIN_ANIMATION_FRAME_DELAY_MS: u32 = 16;
/// Frames an animation is given before it is handed over.
///
/// Two, rather than a startup buffer: the frame the preview opens on is the first
/// one, and holding a buffer before showing anything is a delay the user reads as
/// the preview not appearing. The decoder runs on ahead of playback once it has
/// been handed over, so it is the queue depth that keeps playback fed, not this —
/// and every frame still reaches the screen, because the handover only decides
/// when the preview opens, not how much of the animation is played.
pub(super) const ANIMATION_STARTUP_FRAMES: usize = 2;
pub(super) const STREAMING_SPINNER_MAX_MS: u64 = 1500;
#[derive(Clone)]
pub enum PreviewMessage {
    /// A preview of the file the pointer hovers, opened from the cursor it was
    /// hovered at.
    ///
    /// The region that comes with it is what the `Avoid` setting measured off the
    /// hovered item — its name, and the columns beside it at `Avoid Details` — which
    /// the placement is kept off at either way of avoiding, and whether that region is
    /// a column of the item's view, which decides *how* a placement is kept off it: a
    /// column is stepped off to one of its sides, an item's own text in whichever
    /// direction asks the least. See `AvoidRegion` and `avoiding_text`. It is `None`
    /// when the setting is off, when the view reported no text for the item, or when
    /// the walk found no item at the cursor.
    Show(PathBuf, i32, i32, Option<AvoidRegion>),
    /// A preview of the focused item, whose box comes with it, and the region the
    /// `Avoid` setting keeps it off — the item's own text, or the name alone, or the
    /// column the name sits in, depending on the way the setting is on. That region is
    /// where its preview is placed from as well as what it is kept off, the way a
    /// hovered item's is, for an item that draws its text as a row of its view: the
    /// last field says whether it does, read off the item's own text by the hook (see
    /// `explorer_hook::ItemText`). See `compute_keyboard_layout`.
    ShowKeyboard(PathBuf, i32, i32, i32, i32, Option<ScreenRegion>, bool),
    Hide,
    Refresh,
    /// A preview type was switched on or off. Only a preview whose own kind is
    /// now off is rebuilt, and it is rebuilt from the hover it came from, so it
    /// goes away on the spot rather than at the next pointer move.
    RefreshTypes,
    /// The pin key was pressed while a preview was on screen: what is up becomes a
    /// window of its own — captioned, movable, always on top — and the hover machinery
    /// stays quiet behind it until it is closed. The rect is the box the preview is
    /// already in, so nothing about the picture moves when it is pinned.
    Pin {
        path: PathBuf,
        rect: ScreenRegion,
    },
    /// The pinned window was given another box — maximized, restored, or resized by an
    /// edge — and the media in it is laid out again for that box. It is the same
    /// question a hover asks, with the box already answered.
    PinBox(ScreenRegion),
    /// The file the user picked while a preview was pinned, which the pin is asked to show
    /// instead of the one it has: the window keeps its place and its size, and the media in it
    /// becomes the new file's, fitted to the box the pin already has (see `Pin Mode → Update
    /// Preview`). The Explorer hook sends it — for a click, for a key, and — where the setting
    /// asks for it — for the pointer settling on another file (see `update_pinned_preview`).
    PinUpdate(PathBuf),
    /// Whether the pin key is watched was changed in the tray, or previews themselves
    /// were turned off. A pin on screen is ended here rather than left as a window
    /// nothing would ever take down again.
    PinChanged,
    /// The render tier is done with a document: a page is waiting in the cache,
    /// or there is no page. The generation is the hover that asked for it, so a
    /// render landing after the pointer has moved on is ignored — the page is
    /// still cached for the next hover either way.
    OfficeRenderReady {
        path: PathBuf,
        generation: u64,
        ok: bool,
    },
    /// A video's probe is done: the geometry is waiting in the cache, or the answer is
    /// that there is none. The generation is the hover that was waiting on it, so a probe
    /// landing after the pointer has moved on is ignored — the answer is held for the next
    /// hover either way (see `video_probe`).
    VideoProbed {
        path: PathBuf,
        generation: u64,
    },
    /// A measure that reads a file is done: the box it answers with is held, or the answer is
    /// that the reader has none for the file — which is not a wait that can be answered, so it
    /// is the one answer a wait comes down on (see `measured_off_the_tick`).
    ///
    /// What was waiting on it is the hover that is on screen, and that is the whole of what
    /// this answer has to be matched to: a box is measured per version of the file, and a hover
    /// of a file whose version has changed is measured again rather than answered with the box
    /// of the version before it.
    MeasureProbed {
        path: PathBuf,
        size: Option<(u32, u32)>,
    },
    /// The planner is done with a question the pin asked — a walk along the folder, or the
    /// name of the program a file would open with. Neither question can be asked on the
    /// thread that draws the pin, so both are asked elsewhere and answered here (see
    /// `PinPlanner`).
    ///
    /// A walk is answered whatever it found: a folder that could not be read and a folder
    /// with nothing in it to step onto are both answers, not the absence of one, so the pin's
    /// arc comes down when they land rather than waiting out a bound (see `answer_pin_job`).
    PinAnswered(PinPlanned),
    /// The small subtitle files a film's own tracks were copied into are ready, sent the
    /// moment the extraction's one pass finishes (see `subtitle_files`).
    ///
    /// Only a pinned window acts on it: a pin has its film on screen already, so what the
    /// copy landing changes is what the frame after it should be drawn with — the player is
    /// begun again to draw that (see `reload_pinned_subtitles`). A hover is deliberately not
    /// begun again under the pointer, and this message is not for one (see the note on the
    /// first hover in `video_launch::subtitle_filter`).
    VideoSubtitlesReady(PathBuf),
    /// The ImageMagick engine is done with a file: the picture it developed is in hand, or
    /// there is none — a file it cannot read is remembered as one it will not draw. The
    /// generation is the hover that was waiting on it, so a conversion landing after the
    /// pointer has moved on is ignored; the picture itself is held for the hover that asked
    /// either way (see `magick_render_is_due`).
    MagickReady {
        path: PathBuf,
        generation: u64,
        ok: bool,
    },
    /// The PeaZip engine is done with a file: the archive's table of contents is in hand, read
    /// into the listing cache, or there is none — a file it cannot open is remembered as one it
    /// will not list. The generation is the hover that was waiting on it, so a listing landing
    /// after the pointer has moved on is ignored; the listing itself is held under the file's own
    /// key either way, so the next hover of it is a read rather than a launch (see
    /// `peazip_render_is_due`).
    PeazipReady {
        path: PathBuf,
        generation: u64,
        ok: bool,
    },
}

/// Represents different types of media we can display
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum MediaType {
    StaticImage,
    /// A `.dds` texture, which is a still picture decoded by this app and drawn by this
    /// window like any other — a kind of its own for the one thing about it that differs:
    /// what a preview is drawn over. A texture's alpha channel is as often a mask, a
    /// height or a channel a tool never filled in as it is transparency, so its backdrop
    /// is the tray's own setting rather than a picture's (see the `Background` submenu and
    /// `dds_image`).
    Dds,
    /// An SVG document the engine draws in a window of its own. This side holds no
    /// frame for one: the kind is what the preview loop reads to hand the hover over,
    /// and the media it comes in arrives with nothing in it. It is a kind of its own
    /// rather than a static image because a document is composited over a backdrop of
    /// its own (see the tray's `Background` submenu) — a backdrop the engine is given
    /// rather than one this app composes.
    EngineSvg,
    /// A font the engine draws a specimen of in a window of its own, the same shape as a
    /// document: this side holds no frame for one, and what it answers about a font — the
    /// lines the file's own character map covers, and the face a collection is drawn by —
    /// comes from `font_preview`. It is a kind of its own for the reason above it, and
    /// because what it is drawn over is a page of its own rather than a document's.
    EngineFont,
    AnimatedGif,
    AnimatedApng,
    AnimatedWebP,
    /// An AVIF or HEIF image sequence, played frame by frame out of the media engine
    /// Windows has and fed through the same queue every other animation is (see
    /// `heif_sequence`).
    AnimatedHeif,
    /// An animated JPEG XL, read by `jxl-oxide` (see `jxl_image`).
    AnimatedJxl,
    /// A video played by `ffplay`, whose own window is the preview while it is up.
    Video,
    /// A video the media engine Windows has decodes and this window draws — the kind of
    /// video preview a machine without FFmpeg gets (see `video_player`).
    ///
    /// It is a kind of its own rather than a difference inside `Video` because the two are
    /// drawn by different things: the player's window is the preview for one, and the
    /// layered window this app owns is the preview for the other, which is the one
    /// question `render_layered_preview_at` asks of a frame.
    NativeVideo,
    /// A sound: a card of what the file holds, painted into this window's own frame like an
    /// archive's page, with the sound itself played by one of the two engines behind it (see
    /// `audio_preview`). Nothing of the player is drawn — the card is the whole of what is on
    /// screen — which is why a sound is a kind of its own here rather than a video with no
    /// frames in it.
    Audio,
    Pdf,
    Text,
    Archive,
    /// An archive an installed PeaZip listed for this app — a cabinet file, an iso, a disk image,
    /// a Linux package, a single-stream `.gz` or `.zst` — drawn as the same page an archive this
    /// app read itself is drawn as, because that is what it is: a list of what the file holds,
    /// painted by this window into a frame of its own.
    ///
    /// It is a kind of its own for the gate alone, the way the picture an image converter develops
    /// is: the switch over these is not the switch for archives, so a user who wants their isos
    /// left alone is not asking for their zips to be left alone. See `peazip_formats` and
    /// `peazip_render`.
    Peazip,
    Office,
    /// A design document, previewed from the picture its own format keeps of the whole
    /// thing: the merged image at the end of a Photoshop file, or the flattened
    /// document a project container holds beside its layers. It is a kind of its own
    /// for the reason the texture above is — the gate over it is not the gate over
    /// pictures, so a user who wants none of them has a switch that is not the switch
    /// for images — while what draws one is this window, into a frame composed like
    /// any other.
    Design,
    /// A vector drawing: Windows' metafiles, and the preview an encapsulated PostScript
    /// file carries. It is a kind of its own — drawn into this window's frame like a
    /// picture, but drawn by the drawing layer rather than decoded, and composited over a
    /// backdrop of its own — and what makes it worth a kind of its own is the size: the
    /// records are replayed at whatever box the preview is shown at, so a drawing is
    /// sharp at any size the display has.
    Vector,
    /// A document drawn by a render engine rather than read here — CorelDRAW above all:
    /// what comes back is a page, which is sharp at whatever size the preview is shown at,
    /// and what such a file keeps of itself is a thumbnail this app does not show. See
    /// `libre_formats` and `libreoffice_render`.
    Libre,
    /// A picture an installed ImageMagick developed for this app — a camera raw above all:
    /// what comes back is a PNG, which is decoded and drawn like the picture it is, over the
    /// backdrop a picture is drawn over and at the share of its own size a picture is drawn
    /// at. It is the picture kind's second half, and the switch over it is the switch for
    /// pictures; see `magick_formats` and `imagemagick_render`.
    Magick,
    /// A page an installed Calibre converted a book into — a Kindle or Mobipocket file, an EPub,
    /// a FictionBook, a scanned book — drawn exactly as a PDF page is: a frame of this app's own,
    /// made from the first page of the PDF the engine wrote that says anything about the book, at
    /// the share of the display a book is drawn over. It is the book kind's second half, and the
    /// switch over it is the switch for books; see `calibre_formats` and `calibre_render`.
    ///
    /// It is a kind of its own rather than `Pdf` because the two are gated apart: the switch a user
    /// throws over a book the app drew itself is not the switch over one an engine had to convert,
    /// and only one of the two costs a conversion to show.
    Calibre,
    /// The first plate of a comic book, read out of the container it is filed in — a `.cbz`, a
    /// `.cbr` or a `.cbc`: a plate is decoded the way a picture is, and then drawn the way a page
    /// is, at the share of the display a book is drawn over. It is the book kind's *other* half,
    /// and what it has in common with the two above is the whole of why it is here: what a hover on
    /// a book shows is a page, whichever reader drew one. See `ebook_formats` and `comic_preview`,
    /// and note that what draws it is this window rather than an engine — there is nothing to wait
    /// for and nothing to keep.
    Comic,
    /// A file the pin was shown and its player could not draw, standing as the cross
    /// `pin_chrome::paint_failure_mark` draws. It is a kind of its own because it has to be:
    /// the whole of what it is for is that `pin_media_is_alive` reads it as *there* — a frame
    /// this app holds is a frame nothing outside this thread can take away, and that is the
    /// answer a window standing over this mark wants — and because a file that was tried and
    /// failed must stay distinguishable from a file that was never previewable, whose pin keeps
    /// the file it already had and shows nothing of the new one at all.
    ///
    /// Its own kind rather than a still image of the cross, which is the cheaper arrangement and
    /// the dishonest one: everything that keys on `StaticImage` would answer about this file as
    /// though the cross were its picture — the image gates, the animated-image handling, the
    /// cache key it would be held under — and a file that fails must never be cached as the
    /// mark that stood in for it (see `unplayable_media`).
    Unplayable,
    Loading,
}

impl MediaType {
    /// The kind of preview this media is, as the tray's gates name them.
    pub(super) fn kind(&self) -> Option<PreviewType> {
        match self {
            Self::StaticImage
            | Self::AnimatedGif
            | Self::AnimatedApng
            | Self::AnimatedWebP
            | Self::AnimatedHeif
            | Self::AnimatedJxl => Some(PreviewType::Images),
            // A texture is a picture as far as the gates go: the list a `.dds` is in is the
            // image list, and the switch for pictures is the switch for it.
            Self::Dds => Some(PreviewType::Images),
            Self::EngineSvg => Some(PreviewType::Vector),
            Self::EngineFont => Some(PreviewType::Fonts),
            Self::Video | Self::NativeVideo => Some(PreviewType::Videos),
            // A sound is its own kind in the tray as well as here: the switch over it is the
            // switch over the sound list, not the video's.
            Self::Audio => Some(PreviewType::Audio),
            Self::Text => Some(PreviewType::Text),
            Self::Pdf => Some(PreviewType::Ebook),
            Self::Archive => Some(PreviewType::Archives),
            // And an archive an engine listed: the page is an archive's page in every way that
            // matters — a frame of this app's own, painted from a listing — and the switch over
            // it is the archive switch, because what a user turns off is archives.
            Self::Peazip => Some(PreviewType::Peazip),
            Self::Office => Some(PreviewType::Document),
            // A design document is drawn into a frame like any picture, and the switch
            // over it is its own: the picture *is* what the file keeps of the document,
            // but a user who wants none of them is not asking for pictures to be off.
            Self::Design => Some(PreviewType::Design),
            // The same for a document an engine drew, at the gate over the document kind.
            Self::Libre => Some(PreviewType::Libre),
            // And for a picture one developed: it is a picture in every way that matters —
            // a frame of this app's own, drawn like any other — and the switch over it is the
            // switch for pictures, because what a user turns off is pictures.
            Self::Magick => Some(PreviewType::Magick),
            // And a book an engine converted: a page like any other, at the gate over books,
            // because what a user turns off is books.
            Self::Calibre => Some(PreviewType::Calibre),
            // And a comic, which is the same kind of preview as a PDF — a page of a book — drawn by
            // a reader of this app's own rather than by an engine, so the switch over it is the
            // switch for books and nothing else.
            Self::Comic => Some(PreviewType::Ebook),
            Self::Vector => Some(PreviewType::Vector),
            // The mark a failed file stands as is this app's own answer rather than a preview of
            // the file's kind, so it is behind no gate: turning a kind off is a statement about
            // what the pointer is offered, and a window already standing over a file it could not
            // draw does not become a file that can be hidden the moment the user hides films.
            Self::Unplayable => None,
            Self::Loading => None,
        }
    }

    /// Whether this is the spinner standing in for a preview that is not ready.
    pub(super) fn is_loading(&self) -> bool {
        matches!(self, Self::Loading)
    }

    /// Whether this is a document the engine draws rather than a frame this app holds.
    /// The preview loop asks before it reaches for a frame, because there is none.
    pub(super) fn is_engine(&self) -> bool {
        matches!(self, Self::EngineSvg | Self::EngineFont)
    }

    /// Whether this is a video the media engine decodes and this window draws, which is
    /// what the tick asks before it pulls a frame.
    pub(super) fn is_native_video(&self) -> bool {
        matches!(self, Self::NativeVideo)
    }

    /// Whether this kind's picture is a window of its own — another process's window standing
    /// behind the pin's band — rather than a frame this window draws.
    ///
    /// It is the question a park is asked, and the whole of it: what a drag puts away is the
    /// window behind the band, so a band this app fills itself has nothing to put away, and the
    /// flat fill a park paints would be laid over a picture this window is drawing — a frame of
    /// this app's own, blanked for the length of a drag and answered with a stray decode. FFmpeg's
    /// player and the browser that draws an engine document are the two of them; the media engine's
    /// video is drawn into this window's own surface like any other frame (see
    /// `park_pinned_player`).
    pub(super) fn draws_in_a_window_of_its_own(&self) -> bool {
        matches!(self, Self::Video) || self.is_engine()
    }

    /// Whether this is a sound, whose card is the one painted preview that changes while it is
    /// on screen: the clock and the bar under it are drawn from a player that is running.
    pub(super) fn is_audio(&self) -> bool {
        matches!(self, Self::Audio)
    }

    /// Whether this preview's appearance is painted into its own frame rather
    /// than recomposited from shared pixels, which is what decides whether a
    /// theme switch means rebuilding it.
    pub(super) fn is_painted(&self) -> bool {
        matches!(
            self,
            Self::Text | Self::Archive | Self::Peazip | Self::Audio
        )
    }

    /// Whether this kind has a picture for the round bubble a collapsed pin becomes. The
    /// kinds whose frame *is* a picture have one; a page of text, an archive's listing and a
    /// sound's card are layouts of text, which at a bubble's size is a grey smear rather than
    /// a picture of anything, so they are drawn as a mark instead (see `pin_chrome`).
    pub(super) fn has_bubble_picture(&self) -> bool {
        matches!(
            self,
            Self::StaticImage
                | Self::AnimatedGif
                | Self::AnimatedApng
                | Self::AnimatedWebP
                | Self::AnimatedHeif
                | Self::AnimatedJxl
                | Self::Dds
                | Self::Design
                | Self::Vector
                | Self::Magick
                | Self::Pdf
                | Self::Comic
                | Self::Office
                | Self::Libre
                | Self::Calibre
                // The cross is a frame like any other, so a bubble collapsed over a file that
                // failed shows the failure rather than a mark standing in for a picture.
                | Self::Unplayable
        )
    }

    /// The mark the bubble carries when there is no picture to stand in for one.
    pub(super) fn bubble_mark(&self) -> pin_chrome::BubbleMark {
        match self {
            Self::Video | Self::NativeVideo | Self::Audio => pin_chrome::BubbleMark::Play,
            Self::Text | Self::Archive | Self::Peazip => pin_chrome::BubbleMark::Page,
            _ => pin_chrome::BubbleMark::Picture,
        }
    }
}

/// A single frame of image data
#[derive(Clone)]
pub(super) struct ImageFrame {
    pub(super) pixels: Vec<u8>,
    pub(super) width: u32,
    pub(super) height: u32,
    pub(super) delay_ms: u32, // Delay before next frame (for animations)
    /// Whether every pixel of the frame has an alpha of 255, which is what lets a repaint
    /// copy the frame instead of blending it pixel by pixel.
    ///
    /// A frame that is opaque everywhere is the surface it is drawn on already, whatever the
    /// backdrop behind it is, so composing it is a copy of its bytes — and at the size of a
    /// display that is the difference between a few gigabytes a second and a few hundred
    /// megabytes, on every frame of a video or an animation (see
    /// `compose_preview_pixels_into`).
    ///
    /// It is asked where the pixels are made, off the thread that draws them, and never
    /// guessed: a frame whose producer has not asked keeps the blend, which is the same
    /// picture by a longer road. `false` is therefore always safe and `true` never is — the
    /// one producer that says so without asking is the video path, which forces the alpha of
    /// every pixel it writes (see `copy_locked`).
    pub(super) opaque: bool,
}

impl ImageFrame {
    /// A frame of `pixels`, with its opacity asked of the pixels themselves.
    ///
    /// The question is a pass over the frame, which is why it is asked here rather than by
    /// whoever composes it: a pass per repaint would cost what the copy it enables saves.
    pub(super) fn new(pixels: Vec<u8>, width: u32, height: u32, delay_ms: u32) -> Self {
        let opaque = pixels_are_opaque(&pixels);

        Self {
            pixels,
            width,
            height,
            delay_ms,
            opaque,
        }
    }

    /// Replace the pixels of a frame, with the opacity asked of them again.
    pub(super) fn set_pixels(&mut self, pixels: Vec<u8>) {
        self.opaque = pixels_are_opaque(&pixels);
        self.pixels = pixels;
    }
}

/// Whether every pixel of a frame is opaque, which is what a repaint copies rather than
/// blends (see `ImageFrame`).
pub(super) fn pixels_are_opaque(pixels: &[u8]) -> bool {
    pixels
        .as_chunks::<4>()
        .0
        .iter()
        .all(|pixel| pixel[3] == 255)
}

/// Frames an animated preview streams while it plays. The decoder appends to the
/// queue and the player drains it; `released` records that the player gave back
/// frames it already showed, after which the animation can no longer loop from
/// memory and the decoder has to start the file over.
///
/// The two are read and written under this one lock, and that is what makes the
/// arrangement sound: a pass that ends with nothing given back leaves every frame
/// in the player's hands and the file needs no decoder again (`decoded`), while a
/// pass that ends after frames were given back is a pass to run again. Settling
/// both under the same lock is what keeps a release and the end of a decode from
/// passing each other, which is a player left holding the last frame of an
/// animation whose decoder has already gone.
pub(super) struct StreamedFrames {
    pub(super) queue: VecDeque<ImageFrame>,
    pub(super) released: bool,
    /// Whether a pass was decoded whole with nothing given back, so the player
    /// holds the file entire and loops it from memory. Nothing is released after
    /// this: a frame dropped out of a file the decoder is done with could never be
    /// decoded again.
    pub(super) decoded: bool,
}

/// What a text preview on screen keeps so that it can be worked with: where it is
/// scrolled to, what is selected in it, and the numbers a press is tested against.
///
/// A text preview in full mode always has one of these, whether or not the
/// document is longer than the frame — a selection needs somewhere to live even
/// when there is nothing to scroll. Without full mode a text preview keeps none of
/// it: nothing scrolls, nothing is selected, and a pointer over it dismisses it the
/// way any other preview is dismissed.
pub(super) struct TextPreviewState {
    pub(super) path: PathBuf,
    pub(super) options: TextPreviewOptions,
    pub(super) dpi: u32,
    pub(super) width: u32,
    pub(super) height: u32,
    /// Document line the frame starts at, and how many it shows.
    pub(super) first_line: usize,
    pub(super) visible_lines: usize,
    /// Lines the preview can reach, which is also the range its scrollbar is
    /// drawn and dragged in.
    pub(super) scrollable_lines: usize,
    /// The bar drawn in the frame, kept so a drag can be tested against it.
    pub(super) scrollbar: Option<text_preview::ScrollBar>,
    /// Whether the pointer is currently dragging the thumb.
    pub(super) dragging: bool,
    /// The painted lines, which is what a press is turned back into a place in the
    /// text against, and what a selection is copied from.
    pub(super) lines: Vec<text_preview::FrameLine>,
    /// What is selected in the frame on screen, and whether a drag is extending it.
    pub(super) selection: Option<text_preview::Selection>,
    pub(super) selecting: bool,
}

impl TextPreviewState {
    pub(super) fn max_first_line(&self) -> usize {
        self.scrollable_lines.saturating_sub(self.visible_lines)
    }

    /// The line a scroll of `lines` from here lands on, kept inside the document.
    pub(super) fn scrolled_by(&self, lines: i64) -> usize {
        (self.first_line as i64 + lines).clamp(0, self.max_first_line() as i64) as usize
    }

    pub(super) fn can_scroll(&self) -> bool {
        self.max_first_line() > 0
    }
}

#[derive(Clone, Copy)]
pub(super) struct VideoCrop {
    pub(super) width: u32,
    pub(super) height: u32,
    pub(super) x: u32,
    pub(super) y: u32,
}

/// Not `Copy`, and that is the sidecar's doing: it is a path rather than a number, so an answer
/// read out of the cache is cloned rather than taken by value (see `video_probe`).
#[derive(Clone)]
pub(super) struct VideoGeometry {
    /// The shape the preview is placed at: the crop where the probe settled on one, and the
    /// frame itself where it did not.
    pub(super) width: u32,
    pub(super) height: u32,
    /// The frame that shape was read from, which is the whole the crop is a part of: what the
    /// engine's own source rectangle is normalized over (see `video_player::Crop`).
    pub(super) frame_width: u32,
    pub(super) frame_height: u32,
    pub(super) crop: Option<VideoCrop>,
    /// How long the file plays, where the probe read it: what a pinned preview's transport bar
    /// is drawn against. Neither engine hands a length over — FFmpeg's player reports nothing at
    /// all, and the media engine's duration is only known once it is playing — so the probe that
    /// measures the file is asked for this at the same time it is asked for the shape.
    pub(super) duration: Option<f64>,
    /// How many subtitle streams the file carries and which of them the player would reach for
    /// by itself — the whole of what a track choice is made of here (see `video_subtitles`).
    pub(super) subtitles: SubtitleStreams,
    /// The subtitle file lying beside the film, where the folder holds one — the whole of
    /// that file, which is why the filter beside it names no track (see `subtitle_filter`).
    ///
    /// It is here rather than looked for at the launch because looking for it is a walk of the
    /// film's whole folder, and this struct is what the launch reads instead of reading the
    /// directory (see `video_sidecar`).
    pub(super) sidecar: Option<PathBuf>,
    /// The small files this app's own extraction copied the film's subtitle tracks into, with
    /// the container's fonts that came out beside them, where the extraction has answered (see
    /// `subtitle_files`). `None` while nothing has been extracted yet, which is the state the
    /// first hover of a film is answered in.
    ///
    /// A hover that draws one of these opens a few-dozen-kilobyte file rather than the film,
    /// which is the whole of the difference the extraction exists to make — measured, a cold
    /// 1.4 GB film drawn from its extracted `.ass` came up in 492 ms against 14 904 ms with the
    /// film's own embedded track, which streams the whole container before the first frame (see
    /// `video_launch::subtitle_filter`).
    pub(super) derived: Option<DerivedSubtitles>,
    /// The codec name of every subtitle stream, by subtitle-relative index, and of every
    /// attachment in the container's own order: the two lists the extraction's one command is
    /// built from, held beside the answer so that the launch which asks for the pass needs no
    /// second read of the header (see `subtitle_files::request_extraction`).
    pub(super) subtitle_codecs: Vec<String>,
    pub(super) attachment_codecs: Vec<String>,
    /// Whether the one extraction pass has failed, which is the flag that keeps it from being
    /// asked for again: the film is then drawn *without* subtitles and stays fast, because the
    /// film's own embedded track is never named (see `video_launch::subtitle_filter`).
    pub(super) subtitle_extraction_failed: bool,
}

/// The small files this app copied a film's own subtitle tracks into, with the fonts
/// that came out of the container beside them (see `subtitle_files`).
#[derive(Clone, Default)]
pub(super) struct DerivedSubtitles {
    /// By subtitle-relative index: the file for that track, or `None` for one that
    /// could not be copied (a codec with no small form).
    pub(super) tracks: Vec<Option<PathBuf>>,
    /// The folder the container's attached fonts were dumped into, where any were.
    pub(super) fonts: Option<PathBuf>,
}

impl DerivedSubtitles {
    /// The copied file for one subtitle-relative track, where that track was copied at all: a
    /// track whose codec has no small form has no file, and its slot is `None` rather than a
    /// path to something that is not there (see `subtitle_files`).
    pub(super) fn track(&self, index: usize) -> Option<&Path> {
        self.tracks.get(index).and_then(|track| track.as_deref())
    }
}

/// A file's subtitle streams, counted and ordered the way FFmpeg's `-sst s:` specifier orders
/// them: from zero, counting only subtitle streams, whichever numbers the video and audio
/// streams around them happen to carry.
///
/// Both numbers are needed and neither is enough alone. The count is what a next-track press
/// wraps within, so a key pressed on a one-track file is a key that does nothing rather than a
/// track that does not exist; the first is what a relaunch is given before the user has chosen
/// anything, which is the player's own choice written down rather than this app's guess at it —
/// and writing it down is the point, because a relaunch that left it unnamed would be a relaunch
/// that let the player's own default stand in for a choice this app had forgotten to keep.
///
/// The default is *no* streams rather than one, because "the header was never read" and "the
/// header said there are none" are the same answer to every question here — which is not the same
/// thing as reading the tracks out of it.
#[derive(Clone, Copy, PartialEq, Eq, Debug, Default)]
pub(super) struct SubtitleStreams {
    pub(super) count: usize,
    pub(super) first: usize,
}

impl SubtitleStreams {
    /// What a relaunch is told when nothing has been chosen: the player's own first choice,
    /// written down so that it survives the relaunch, and nothing at all for a file with no
    /// subtitles — where naming a track is a track in a file that has none.
    pub(super) fn chosen(&self) -> Option<usize> {
        (self.count > 0).then_some(self.first)
    }
}

/// The file a hover message is about.
pub(super) fn show_path(show: &PreviewMessage) -> Option<&PathBuf> {
    match show {
        PreviewMessage::Show(path, ..) | PreviewMessage::ShowKeyboard(path, ..) => Some(path),
        _ => None,
    }
}

/// A hover about to be replayed, anchored where it belongs now.
///
/// A mouse hover is replayed where the pointer is rather than where the hover opened:
/// what is being put back is the preview of the file under the hand, and the display
/// it is being put back on is the one the pointer is on now — the point it opened from
/// resolves the display it came from, which is the one that is gone. A keyboard hover
/// is the item's own place and is replayed as it came.
///
/// It is asked of every mouse hover the loop takes up, and not only of the ones being
/// replayed. The point a `Show` arrives with is one the Explorer hook sampled at the top
/// of its own tick — ahead of a walk through the shell and of the look for the `Avoid`
/// region the layout is placed by — so the better part of a frame has passed by the time
/// anything is laid out from it, and a fast hand covers dozens of pixels in that time. The
/// one thing the layout does with the point is keep the box clear of it, which is worth
/// something only while the point is still the hand's.
pub(super) fn replay_where_the_pointer_is(show: Option<PreviewMessage>) -> Option<PreviewMessage> {
    match (show, cursor_position()) {
        (Some(PreviewMessage::Show(path, _, _, avoid)), Some(cursor)) => {
            Some(PreviewMessage::Show(path, cursor.x, cursor.y, avoid))
        }
        (show, _) => show,
    }
}

/// Whether a hover message is one the loop may act on, given whether a preview is pinned.
///
/// A pinned preview is the whole of what this app is showing, and the hover machinery is refused
/// at three doors for it: the Explorer hook's `show_preview` and `show_preview_keyboard`, which
/// answer nothing while a pin is up, and this one. It is asked here as well because the hook is
/// not the only thing a hover comes from — what is on screen is recorded as a hover message
/// whoever put it there, so an answer that lands for the file the *pin* is showing (a box
/// measured, a video probed, a page drawn, a setting changed in the tray, a hover that was in
/// flight when the pin came up) is replayed out of that record as a hover, and one acted on would
/// stop the pin's media and put a hover where the window was (see the swap's own record and the
/// gate in `run_preview_window`).
///
/// What a pin answers for itself is not a hover: a take-up, a box and a take-down go on to the
/// match as they always did.
pub(super) fn hover_is_shown(message: &PreviewMessage, pinned: bool) -> bool {
    !pinned
        || !matches!(
            message,
            PreviewMessage::Show(..) | PreviewMessage::ShowKeyboard(..)
        )
}

/// Whether the preview of `path` is a text preview — the form a caller with no hover of its own
/// asks, which is one entry read rather than the three this used to make.
pub(super) fn is_text_preview(path: &Path) -> bool {
    HoverFacts::read(path).is_text()
}

/// Whether a preview of `path` is painted into the box it is given rather than scaled within it.
pub(super) fn page_is_painted(path: &Path) -> bool {
    HoverFacts::read(path).is_painted_page()
}

pub(super) fn effective_frame_delay_ms(media_type: &MediaType, source_delay_ms: u32) -> u32 {
    match media_type {
        MediaType::AnimatedWebP => {
            let fps = current_webp_playback_fps();
            let min_delay_ms = (1000 / fps).max(1);
            source_delay_ms.max(min_delay_ms)
        }
        _ => source_delay_ms,
    }
}

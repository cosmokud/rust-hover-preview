//! What the pointer is over: the file it names, the kind that file is read as, the box it was
//! measured into, and the measurements a hover has had taken off the tick.

use super::*;

/// What a box measured off this thread was measured against.
///
/// It is part of what a held box is keyed by because a box is only the answer for what it was
/// measured against: a page's own size is the file's, whatever it is drawn on, while a
/// listing's page is wrapped to the room it is shown in at the text settings it is wrapped
/// for — and a box held for another room is a page laid out to the wrong one.
#[derive(Clone, PartialEq, Eq, Hash)]
pub(super) enum MeasureScope {
    /// The file's own size: a page's, a plate's, a drawing's declared extent, a specimen's
    /// box.
    File,
    /// A page wrapped to the room it is drawn in, at the text settings it is wrapped for.
    Room {
        cap_width: u32,
        cap_height: u32,
        dpi: u32,
        theme: TextTheme,
        font_scale_percent: u32,
    },
}

/// A box measured off the preview thread, and what it was measured against.
pub(super) struct MeasuredBox {
    pub(super) path: PathBuf,
    pub(super) version: FileVersion,
    pub(super) scope: MeasureScope,
    /// The box, or `None` for a file its reader has no answer for — which is an answer too,
    /// and one worth holding: a document that will not open is not one to read again on every
    /// hover.
    pub(super) size: Option<(u32, u32)>,
}

/// The boxes measured off the preview thread, newest first.
///
/// It is a table of its own rather than the readers' own memos, and that is the point: what a
/// reader remembers is dropped wholesale when it fills, and a box that fell out of one of
/// those memos would be measured again the moment it was asked for — on the thread that draws
/// the hover, which is what these measures are kept off. What is held here is answered from
/// here, and the reader's own memo is left to the side that draws the file (see
/// `measured_off_the_tick`).
pub(super) static MEASURED_BOXES: Lazy<Mutex<Vec<MeasuredBox>>> =
    Lazy::new(|| Mutex::new(Vec::new()));

/// Entries the table of measured boxes holds before it is emptied.
pub(super) const MEASURED_BOXES_MAX_ENTRIES: usize = 256;

/// The measures running right now, by the file being measured.
///
/// One thread per file rather than one per hover: a hover that lands on a file whose measure is
/// already running waits for that one instead of starting a second read of the same bytes, and
/// what the layout asks to place that wait is a question about this list (see
/// `measure_waiting`). A file whose version changes while it is being read is measured again by
/// the next hover: this list says a read is running, and what that read answered is held only
/// for the version it was a read of (see `hold_box`).
pub(super) static MEASURING: Lazy<Mutex<Vec<PathBuf>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// Entries the list of running measures holds before it is emptied. It is a list of reads in
/// flight, so it is never more than a handful long; the ceiling is what a stuck thread would
/// cost.
pub(super) const MEASURING_MAX_ENTRIES: usize = 64;

/// The box held for this version of this file, measured against `scope` — or `None` when
/// nothing is held for it, which is not the same answer as a held `None`.
pub(super) fn held_box(
    path: &Path,
    version: &FileVersion,
    scope: &MeasureScope,
) -> Option<Option<(u32, u32)>> {
    let boxes = MEASURED_BOXES.lock().ok()?;

    boxes
        .iter()
        .find(|held| held.path == path && held.version == *version && held.scope == *scope)
        .map(|held| held.size)
}

/// Hold the box a measure answered with.
pub(super) fn hold_box(
    path: &Path,
    version: &FileVersion,
    scope: &MeasureScope,
    size: Option<(u32, u32)>,
) {
    let Ok(mut boxes) = MEASURED_BOXES.lock() else {
        return;
    };

    boxes.retain(|held| !(held.path == path && held.version == *version && held.scope == *scope));
    if boxes.len() >= MEASURED_BOXES_MAX_ENTRIES {
        boxes.clear();
    }

    boxes.insert(
        0,
        MeasuredBox {
            path: path.to_path_buf(),
            version: version.clone(),
            scope: scope.clone(),
            size,
        },
    );
}

/// Say that this file is being measured, answering whether one already was.
pub(super) fn begin_measure(path: &Path) -> bool {
    let Ok(mut measuring) = MEASURING.lock() else {
        return false;
    };

    if measuring.iter().any(|running| running == path) {
        return false;
    }

    if measuring.len() >= MEASURING_MAX_ENTRIES {
        measuring.clear();
    }

    measuring.push(path.to_path_buf());

    true
}

/// Say that the measure of this file is done with.
pub(super) fn end_measure(path: &Path) {
    let Ok(mut measuring) = MEASURING.lock() else {
        return;
    };

    measuring.retain(|running| running != path);
}

/// Whether this file is being measured right now: the question the layout places the wait by.
///
/// It is asked of the file the hover is on, straight after the layout measured it, and what it
/// says is whether that measure handed back the wait for a read rather than a box. The two
/// cannot disagree — the wait is placed exactly when the measure the layout has just taken
/// started this file's read (see `measured_off_the_tick`).
pub(super) fn measure_waiting(path: &Path) -> bool {
    MEASURING
        .lock()
        .map(|measuring| measuring.iter().any(|running| running == path))
        .unwrap_or(false)
}

/// Measure a file on a thread of its own, and tell the preview loop.
///
/// The measure is a read that can be felt — a PDF opened, an archive's table of contents
/// walked, a document parsed, a specimen read — and the hover waits for it, so it is taken
/// here rather than on the preview thread: what is on screen while it runs is the spinner, and
/// the hover it belongs to is replayed when the answer lands, laid out at the box that answer
/// is held under. It is the shape a video's probe has, for the same reason (see
/// `spawn_video_probe`).
///
/// A measure whose hover has moved on is not wasted: what it answered is held for the next
/// hover of the file, so nothing here is cancelled or waited for.
///
/// What a hover waits for is the answer rather than the read, and the answer is sent whatever
/// became of the read — a measure that panicked included. A thread that unwound past the two
/// calls below would leave the file on the list of reads running with nothing to take it off
/// it, and what a hover on that file would be from then on is a spinner nothing ends: the
/// layout goes on placing the wait, since a read for the file is running, and the cap a wait
/// is given is the engines' and does not stand behind a read (see `measured_off_the_tick`
/// and `awaiting_engine`). A read that came apart is answered with the reader's own "nothing
/// for this file" and held like it, so a file whose read comes apart is not read again on
/// every hover either (see `spawn_video_probe`, whose probe answers through the same guard).
pub(super) fn spawn_measure_probe(
    path: PathBuf,
    version: FileVersion,
    scope: MeasureScope,
    measure: impl FnOnce() -> Option<(u32, u32)> + Send + 'static,
) {
    std::thread::spawn(move || {
        // A read that panicked is no box, and no box is an answer; what it must not be is an
        // unwind past the mark that says the read is done with (see above).
        let size = std::panic::catch_unwind(std::panic::AssertUnwindSafe(measure)).unwrap_or(None);

        // The box is held before the read is marked done, so a hover that arrives while this
        // thread is between the two finds the answer rather than starting a read of its own.
        hold_box(&path, &version, &scope, size);
        end_measure(&path);

        notify_measured(&path, size);
    });
}

/// A measure has been taken off the preview thread, and the box it answered with is held. Sent
/// from the thread the measure ran on, through the same channel every other answer arrives on,
/// so the hover that was waiting for it is replayed the moment there is a box to place it with
/// — or, where the reader has no box for the file at all, told that the wait is over (see
/// `MeasureProbed`).
pub(super) fn notify_measured(path: &Path, size: Option<(u32, u32)>) {
    send_preview(PreviewMessage::MeasureProbed {
        path: path.to_path_buf(),
        size,
    });
}

/// The box a measure that reads a file answers with: the box this side already holds, or the
/// wait for one that is being measured now.
///
/// `measure` is the reader's own measure — the same call this side would otherwise make on its
/// own thread — and it runs on a thread of this function's own making, once per file version
/// and scope. What comes back to the hovering call meanwhile is `waiting`, which
/// is what the layout places a hover at until the answer lands (see `measure_waiting` and
/// `MeasureProbed`): the spinner's own box for a read that is about the file, and the room
/// a card will take for one that is about the box it is laid out in, so that the wait is
/// never laid out at a size the answer will not stand at.
pub(super) fn measured_off_the_tick(
    path: &Path,
    scope: MeasureScope,
    waiting: (u32, u32),
    measure: impl FnOnce() -> Option<(u32, u32)> + Send + 'static,
) -> Option<(u32, u32)> {
    let version = file_version(path);

    if let Some(held) = held_box(path, &version, &scope) {
        return held;
    }

    if begin_measure(path) {
        spawn_measure_probe(path.to_path_buf(), version, scope, measure);
    }

    Some(waiting)
}

/// Which preview surface the pointer is currently on.
#[derive(Clone, Copy)]
pub struct PreviewCursorHover {
    pub image: bool,
    pub video: bool,
    /// A document the engine draws, in a window of its own: a preview of this app's in every
    /// way but the window it is drawn in. A document and a specimen are pictures, so the
    /// pointer takes them the way it takes any other preview — arriving at the engine's
    /// window is arriving at the preview, and a document is closed by that arrival. A page
    /// that runs is not a picture: its own rectangle holds the pointer instead, which is a
    /// question asked of the hold and not of this surface (see `preview_pointer_hold`).
    pub engine: bool,
}

impl PreviewCursorHover {
    pub const NONE: Self = Self {
        image: false,
        video: false,
        engine: false,
    };

    pub fn any(self) -> bool {
        self.image || self.video || self.engine
    }
}

/// Single shared pointer probe for both preview kinds. Callers gate it on
/// "a preview can be under the pointer"; the fast path keeps the cost at a few
/// atomic reads whenever nothing is on screen.
pub fn cursor_preview_hover() -> PreviewCursorHover {
    let preview_hwnd = PREVIEW_HWND.load(Ordering::SeqCst);
    let video_hwnd = VIDEO_HWND.load(Ordering::SeqCst);
    let video_pid = VIDEO_PID.load(Ordering::SeqCst);
    // A document is drawn by the engine, in a window of its own rather than this app's, so
    // its surface is asked for by its own handle: it is the same preview to the pointer.
    let engine_hwnd = webview_preview::showing_hwnd();

    // PREVIEW_HWND is created once at startup and never cleared, so visibility
    // is what tells us whether the layered window is actually on screen.
    let preview_visible =
        preview_hwnd != 0 && unsafe { IsWindowVisible(HWND(preview_hwnd as *mut _)).as_bool() };
    if !preview_visible && video_hwnd == 0 && video_pid == 0 && engine_hwnd == 0 {
        return PreviewCursorHover::NONE;
    }

    unsafe {
        use windows::Win32::UI::WindowsAndMessaging::{GetCursorPos, IsChild, WindowFromPoint};

        let mut cursor_pos = POINT::default();
        if GetCursorPos(&mut cursor_pos).is_err() {
            return PreviewCursorHover::NONE;
        }

        let hwnd_under_cursor = WindowFromPoint(cursor_pos);
        if hwnd_under_cursor.is_invalid() {
            return PreviewCursorHover::NONE;
        }

        let hwnd_ptr = hwnd_under_cursor.0 as isize;
        let image = preview_hwnd != 0 && hwnd_ptr == preview_hwnd;
        // A document is drawn by a browser inside the engine's window, so what the pointer
        // is over is a window of the browser's — a child of the engine's, one or two levels
        // down — and not the engine's own window at all. The engine's window is what the
        // preview is, so the question is whether the window under the pointer is that
        // window or one inside it: comparing the two handles alone never matched, and the
        // touch that closes a document was the one thing that never came from here.
        let engine = engine_hwnd != 0
            && (hwnd_ptr == engine_hwnd
                || IsChild(HWND(engine_hwnd as *mut _), hwnd_under_cursor).as_bool());

        // A hit on the stored HWND is enough; the process-ID fallback covers the
        // race window where ffplay's window exists but VIDEO_HWND isn't stored yet.
        let mut video = video_hwnd != 0 && hwnd_ptr == video_hwnd;
        if !video && video_pid != 0 {
            use windows::Win32::UI::WindowsAndMessaging::GetWindowThreadProcessId;

            let mut window_pid: u32 = 0;
            GetWindowThreadProcessId(hwnd_under_cursor, Some(&mut window_pid));
            video = window_pid == video_pid;
        }

        PreviewCursorHover {
            image,
            video,
            engine,
        }
    }
}

/// Every scale a hover is laid out by, read together so that a measure and the render
/// that follows it agree on all of them — one read of the configuration rather than a
/// handful of them, and one answer per kind of preview.
pub(super) fn current_hover_scales() -> HoverScales {
    CONFIG
        .lock()
        .map(|cfg| HoverScales::of(&cfg))
        .unwrap_or(HoverScales {
            picture: PreviewScale::Percent(DEFAULT_PREVIEW_SCALE_PERCENT),
            video: PreviewScale::Percent(DEFAULT_VIDEO_SCALE_PERCENT),
            animated: PreviewScale::Percent(DEFAULT_ANIMATED_SCALE_PERCENT),
            ebook: DEFAULT_EBOOK_SCALE,
            document: DEFAULT_DOCUMENT_SCALE,
            font: DEFAULT_FONT_SCALE,
            design: DEFAULT_DESIGN_SCALE,
            vector: DEFAULT_VECTOR_SCALE,
        })
}

/// Everything one hover needs to know about the file under the hand, asked once.
///
/// This is the thing the layout already wants to carry. Installing a hover asks six questions
/// about one file — what its content is, what kind its name makes it, whether it is drawn as a
/// video, a sound, a page or an engine window, what size it is measured at, and whether what is
/// on screen for it is a wait — and every one of them used to open the same file's directory
/// entry for itself and take the process-wide configuration lock across the read. One hover of
/// one file therefore paid about twenty `fs::metadata` calls for a question with one answer, on
/// the thread that pumps this window's own messages, which is the thread a pin's caption is
/// dispatched on.
///
/// It is read once, at the top of the arm that installs the hover, and the questions below are
/// handed it rather than a path: the entry, the front of the file and what the bytes named come
/// from [`crate::formats::content_type::Probe`], the kind and the list answers come from the one
/// configuration read beside them, and every question that does not need one of those is a
/// comparison against a field. What is left asking the disk is the question that genuinely has
/// to — whether a `.ts` carries transport packets, what an `.ai` keeps at its front, whether a
/// font parses — and each of those is answered once per file and held.
///
/// A caller with no hover of its own to hand builds one for the file it is asking about, which
/// is what every site outside the `Show` arm does: it costs one entry read where the question
/// used to make several, and it costs nothing at all after the first hover of the same version
/// of the file, because the head and the content answer are both caches keyed by that version.
pub(super) struct HoverFacts {
    pub(super) probe: crate::formats::content_type::Probe,
    /// What this file is, in one answer: its own bytes' kind, the kind its name's lists claim,
    /// which half of a book or a drawing it is, and who draws it. This is
    /// [`crate::formats::routing::resolve`]'s answer, and the predicates below are its fields —
    /// which is what turns twelve scattered answers to one question into lookups.
    pub(super) route: crate::formats::routing::Route,
    /// The four list answers the eleven predicates used to each go and take the configuration's
    /// lock for. They are asked here rather than where they are used because they are cheap — a
    /// name compared against a list of extensions — and because a question asked twice for one
    /// hover is two locks where one was.
    pub(super) video_named: bool,
    pub(super) audio_named: bool,
    pub(super) archive_named: bool,
    pub(super) peazip_named: bool,
    /// The three of the tray's switches that gate the three kinds whose gate changes what is
    /// drawn rather than whether it is, and so are asked inside the predicates that gate.
    pub(super) video_enabled: bool,
    pub(super) audio_enabled: bool,
    pub(super) text_enabled: bool,
    /// What of this app's own reads this file, asked with the kind already settled — which is
    /// the contract `native_formats::job_for` is written for and the reason it takes a kind
    /// rather than working one out. A book is the case that needed it: which half of the kind a
    /// `.cbz` is costs a read of the file, and the loader used to ask it a second time for
    /// itself under a lock of its own.
    pub(super) native_job: Option<native_formats::NativeJob>,
    pub(super) scales: HoverScales,
    pub(super) follow_cursor: bool,
}

/// What a file's content says it is, with the configuration's lock not held across the file.
///
/// One question asked from the hook, the layout, the loader and the engines' own request sides,
/// every one of them for the same file in the same hover, and every one of them taking the
/// process-wide configuration lock across a `content_type::of` — which reads the file's first
/// four kilobytes on a miss. A slow volume therefore froze every thread of the app, including
/// the one pumping the preview window's own messages, which is the thread a pin's caption is
/// dispatched on.
///
/// The fix is the order rather than a copy of the configuration: the file's own entry is read
/// first, without the lock, and the lists are taken under it for the lookup that consults
/// them. A file that cannot be read has no entry and takes the other form, which is the same
/// answer for a file nothing can say anything about.
///
/// It is now the content field of [`HoverFacts`], which is the whole of what this was for: a
/// caller that has a hover's answer in hand reads a field.
pub(super) fn content_of(path: &Path) -> crate::formats::content_type::Content {
    HoverFacts::read(path).route.content
}

/// Whether the file's own bytes name one of this app's kinds *other* than `kind`.
pub(super) fn content_names_another_kind(path: &Path, kind: PreviewType) -> bool {
    HoverFacts::read(path).names_another_kind(kind)
}

impl HoverFacts {
    /// Everything one hover of one file needs to know about that file.
    ///
    /// The two halves are read in this order and not the other way round: the file's own bytes
    /// first, with nothing held, and the configuration's lists after — because the lists are
    /// what the bytes' answer is turned into a kind by, and a guard held across the reading is a
    /// guard every thread of the app waits on for as long as the volume takes to answer it (see
    /// `is_text`, whose own account of the bug is the one this whole struct retires).
    pub(super) fn read(path: &Path) -> Self {
        let probe = crate::formats::content_type::Probe::read(path);

        let Ok(config) = CONFIG.lock() else {
            // Nothing is claimed, so nothing is drawn, and a file a hover cannot read its
            // configuration for is the answer every kind's gate gives when the tray is
            // shut (see `PreviewType::enabled`).
            let route = crate::formats::routing::resolve(path, &AppConfig::default(), &probe);

            return Self {
                probe,
                route,
                video_named: false,
                audio_named: false,
                archive_named: false,
                peazip_named: false,
                video_enabled: false,
                audio_enabled: false,
                text_enabled: false,
                native_job: None,
                scales: current_hover_scales(),
                follow_cursor: true,
            };
        };

        // One answer to every question this hover is going to ask, asked of the file's own bytes
        // and the configuration's lists together — which is what `routing::resolve` is for, and
        // the reason it exists rather than a thirteenth predicate here.
        //
        // The order inside it is the one this function used to have to state: the file is read
        // first, with nothing held, and the lists are consulted after.
        let route = crate::formats::routing::resolve(path, &config, &probe);

        // And the four answers this side has that routing does not: which of the lists' names the
        // file carries, which gates the tray has thrown, and the scales a hover is laid out by.
        let audio_named = crate::formats::lists::AUDIO.claims(path, &config);
        let archive_named = crate::formats::lists::ARCHIVE.claims(path, &config);
        let peazip_named = crate::formats::lists::PEAZIP.claims(path, &config);
        let video_enabled = PreviewType::Videos.enabled_in(&config);
        let audio_enabled = PreviewType::Audio.enabled_in(&config);
        let text_enabled = PreviewType::Text.enabled_in(&config);
        let scales = HoverScales::of(&config);
        let follow_cursor = config.follow_cursor;

        // The video lists are the one exception to "consulted after, and only in memory": the two
        // names they share with the text lists are settled by whether the file holds MPEG-TS
        // packets, which is a `File::open` and a read. They were copied out under the guard for
        // exactly that reason, and this line is the second half of a fix whose first half is
        // `media_engine_plays` doing the same: a guard held across a read is a guard every thread
        // of the app waits on for as long as the volume takes, and this one is taken on the thread
        // that pumps this window's own messages.
        let video_lists = (
            config.video_extensions.clone(),
            config.ffmpeg_extensions.clone(),
        );
        drop(config);

        let video_named =
            video_formats::claims_any_video_name_in(path, &video_lists.0, &video_lists.1);

        // And last, with the guard gone: which of this app's own readers does the work. That is
        // the one question here that still opens the file — a book's is settled by what an `.ai`
        // keeps at an offset, a picture's by its own head — so it is the one that must not be
        // asked with a guard in hand. It takes the entry read above, which makes it a lookup
        // rather than a second `fs::metadata`, and it takes the lock for the one list it consults.
        //
        // A configuration that will not open is answered as the page a book most often is, which
        // is what this arm has always done (see `native_formats::page_job`).
        let native_job = route.named.and_then(|kind| {
            CONFIG
                .lock()
                .ok()
                .and_then(|config| native_formats::job_for(path, kind, &config, probe.facts()))
        });

        Self {
            probe,
            route,
            video_named,
            audio_named,
            archive_named,
            peazip_named,
            video_enabled,
            audio_enabled,
            text_enabled,
            native_job,
            scales,
            follow_cursor,
        }
    }

    /// The kind this file is drawn as: its own bytes' answer where they named one, and the name's
    /// after them.
    ///
    /// It is the loader's kind and the layout's, and it is one answer because it is asked once:
    /// the content tier's verdict outranks the lists for the loader — a `.docx` whose bytes are
    /// an MP4 is loaded as the video it is — and the same answer is what a file whose bytes named
    /// nothing is measured at.
    pub(super) fn routed_kind(&self) -> Option<PreviewType> {
        match self.route.content {
            crate::formats::content_type::Content::Kind(kind) => Some(kind),
            crate::formats::content_type::Content::Foreign => None,
            crate::formats::content_type::Content::Unknown => self.route.named,
        }
    }

    /// What of this app's own reads a book: the first plate out of the container, or the page the
    /// PDF engine draws.
    ///
    /// A configuration that will not open is answered as the page a book most often is, which
    /// is what this arm has always done (see `native_formats::page_job`).
    pub(super) fn book_job(&self) -> native_formats::NativeJob {
        self.native_job.unwrap_or(native_formats::NativeJob::Pdf)
    }

    /// Whether an engine's listing engine would list this file, which is the file's own bytes
    /// first and the listing list after them (see `peazip_formats::is_engine_archive`).
    pub(super) fn engine_archive(&self) -> bool {
        match self.route.content {
            crate::formats::content_type::Content::Kind(PreviewType::Peazip) => true,
            crate::formats::content_type::Content::Kind(_)
            | crate::formats::content_type::Content::Foreign => false,
            crate::formats::content_type::Content::Unknown => self.peazip_named,
        }
    }

    /// Whether this file's own bytes name one of this app's kinds *other* than `kind`.
    ///
    /// It is the question an engine tier asks before it starts anything, and what it says no to
    /// is a file that is called what it is: a name and a content that agree are answered with no
    /// opinion at all, and only a disagreement — a picture under a document's name — is a file
    /// whose engine must not be started.
    pub(super) fn names_another_kind(&self, kind: PreviewType) -> bool {
        matches!(
            self.route.content,
            crate::formats::content_type::Content::Kind(named) if named != kind
        )
    }

    /// Whether this file is an SVG document rather than a drawing the drawing layer replays.
    ///
    /// The name is the whole of it, and deliberately so: an `svg` that is not a document is a
    /// drawing the browser refuses rather than a metafile, and the renderer has the last word on
    /// whether a document is a document at all (see `svg_preview::is_svg_file`).
    pub(super) fn svg_document(&self) -> bool {
        self.route.drawing == crate::formats::routing::Drawing::Svg
    }

    /// Whether a page of HTML is one the browser engine draws.
    ///
    /// It is the one term of the route that is about the machine rather than about the file, so
    /// it is asked where a caller already has the route rather than picked out of three names
    /// a fourth time (see `web_page_of` for the form a caller with no route asks).
    pub(super) fn html_drawn_by_the_engine(&self) -> bool {
        self.route.drawn_by
            == crate::formats::routing::DrawnBy::WebView(crate::formats::routing::WebPage::Html)
            && webview_preview::draws(self.probe.path())
    }
}

impl HoverFacts {
    /// Whether the preview of this file is a text preview.
    ///
    /// It is one question rather than a chain of exclusions: what kind a file has is the
    /// router's answer, and a file is drawn as text exactly when that answer is text. It used
    /// to be written out here as "the text lists claim it and no kind asked earlier does",
    /// with the kinds listed one by one — and the list had been left short, so a name written
    /// into the text list beside a listing engine's or a picture converter's was measured as
    /// text and drawn as the other thing (see `formats::routing`).
    ///
    /// What the file's own bytes say comes first, as it does for the loader that draws it and
    /// for the box it is painted into: a file whose content is another kind is not drawn as
    /// text whatever it is called, and one whose content is text is drawn as text even where
    /// the name is a kind the lists would have claimed first.
    ///
    /// Both of the questions below consult the lists, and both of them used to open the file —
    /// the content one reads four kilobytes on a miss, and the router's video claim reads a
    /// `.ts` to tell a film from a TypeScript file — so the guard used to be held across two
    /// `File::open`s on the thread that pumps this window's messages, and every other thread of
    /// the app waited on that guard, including the one that would end an engine or answer the
    /// tray. Neither is asked again here: this is the answer (see `HoverFacts`).
    pub(super) fn is_text(&self) -> bool {
        match self.route.content {
            crate::formats::content_type::Content::Kind(PreviewType::Text) => true,
            // Another kind, or a format no kind here previews at all: neither is drawn as
            // text, and the second is drawn as nothing.
            crate::formats::content_type::Content::Kind(_)
            | crate::formats::content_type::Content::Foreign => false,
            // The kind the hook called it, asked of the same table the hook asked: a name the
            // text lists hold and an earlier list also claims is that earlier kind, and a
            // preview measured as text would be placed as one and drawn as the other. The
            // switch is part of the question, as it is wherever the text lists are asked — a
            // kind turned off in the tray is not drawn at all.
            crate::formats::content_type::Content::Unknown => {
                self.text_enabled && self.route.named == Some(PreviewType::Text)
            }
        }
    }

    /// Whether a preview of this file is painted into the box it is given rather than scaled
    /// within it: a text file, an archive this app read itself, and an archive an engine
    /// listed are pages of one kind — painted at a fixed font size, so the box the layout
    /// planned for one is the box it draws into, and the frame that comes back is that box
    /// rather than a size to be fitted to a space.
    ///
    /// One question, asked in the two places that have to agree about a kind: the share it is
    /// drawn at (`effective_preview_scale`) and the box the loader is handed, which is what
    /// the window ends up sized to. Asking it in one place is the point — a kind left out of
    /// one of them is a preview that is drawn at the planned size and loaded against the free
    /// room of the display, which is a page stretched to the screen, and that is exactly what
    /// an archive an engine listed was.
    ///
    /// The engine's own question is the third term, asked the way the engine asks it — the
    /// file's bytes first and the name after them — so an archive it lists under a name no
    /// list holds (a `.cab` renamed to `.dat`) is a page here too.
    pub(super) fn is_painted_page(&self) -> bool {
        // A page the engine draws is not painted into the box at all, so it is not this rule.
        if self.html_drawn_by_the_engine() {
            return false;
        }

        self.is_text() || self.archive_named || self.engine_archive() || self.is_audio()
    }

    /// Whether this file is drawn as a video: the name the video list carries, or the bytes
    /// of a video under a name that list does not have.
    ///
    /// Every question about a video goes through this one answer — whether its shape has to
    /// be probed, whether the wait for it is shown, and whether the player takes over the
    /// window rather than this app drawing its frames — because the loader plays the file its
    /// bytes name, and a hover whose picture is played but whose frames are awaited would sit
    /// on a first frame that nothing ever replaces.
    pub(super) fn is_video(&self) -> bool {
        // It was asked from four places in one hover — the probe's due question, both
        // dimension questions and the `Show` arm — and each of those used to read the file and
        // take the configuration lock for it. The content is a cache hit after the first, so
        // the entry read is all that is repeated; the lock is no longer held across the read
        // at all (see `HoverFacts`).
        let named = match self.route.content {
            crate::formats::content_type::Content::Kind(PreviewType::Videos) => true,
            _ => self.video_named,
        };

        named && self.video_enabled
    }

    /// Whether the preview of this file is a sound: what the file's own bytes say it is — the
    /// verdict a probe left behind included — and, for a name no table names, the sound list.
    ///
    /// It is asked the way the video's is asked and for the same reason: a sound is drawn as a
    /// card by this app rather than by a player, so the layout has to know one when it sees
    /// one — which for a renamed file, or for a container whose streams hold only a song, is a
    /// question about the content rather than about the name.
    pub(super) fn is_audio(&self) -> bool {
        if !self.audio_enabled {
            return false;
        }

        match self.route.content {
            crate::formats::content_type::Content::Kind(PreviewType::Audio) => true,
            _ => self.audio_named,
        }
    }
}

/// A box of pixels inside a frame: where it sits in the frame, and how large it is.
///
/// It is what the corner spinner is drawn through — the box is copied out of the frame, drawn
/// into, and composed back at the place it came from — so the four numbers travel together
/// rather than as an argument to everything that touches one (see `render_layered_preview_at`).
#[derive(Clone, Copy)]
pub(super) struct FrameBox {
    pub(super) left: u32,
    pub(super) top: u32,
    pub(super) width: u32,
    pub(super) height: u32,
}

impl FrameBox {
    /// The area of the box in bytes, at four bytes to the pixel.
    pub(super) fn bytes(self) -> usize {
        self.width as usize * self.height as usize * 4
    }
}

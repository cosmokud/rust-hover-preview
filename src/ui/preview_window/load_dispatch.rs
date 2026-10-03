//! How a load is asked for and dispatched: the request a worker thread is handed, the answer
//! it comes back with, and the dispatch that decides which loader a kind is given.

use super::*;

/// Load media (image, animated image, text, or video) with appropriate loader
///
/// `max_width` x `max_height` is the box the loader may draw in, and every loader
/// draws its frame inside it — resized, clamped or rasterized at that size. That is
/// the rule rather than a convenience: the preview window is sized to the frame that
/// comes back from here rather than to the box the layout planned (see
/// `render_layered_preview_at`), and the box is what was fitted to the display, so a
/// frame larger than the box is a preview hanging off the edge of the display. A
/// source that draws at whatever size it is asked for — the Windows PDF engine, whose
/// destination is in DIPs and comes back scaled by the display — has to be drawn back
/// into the box it was given; see `pdf_preview::fit_drawn_page`.
///
/// The one source with no frame to draw is an SVG document, which is the engine's: what
/// comes back for one is the kind alone, and the install path hands the hover over
/// rather than putting anything of this app's up (see `MediaType::EngineSvg`).
pub(super) fn load_media(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    dpi: u32,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    // Every loader below reads the file's bytes, so a file whose content is still
    // in the cloud is answered here rather than after a download the user never
    // asked for. The hook refuses these too; this is the boundary that reads, so
    // it decides for itself rather than trusting that nothing reaches it.
    //
    // The probe below answers that from the file's own entry, so it is not asked twice: the
    // loader is the fourth thing to ask this file what it is in one hover, and each of them
    // was reading the entry for itself (see `HoverFacts`).
    let hover = HoverFacts::read(path);

    if hover.probe.needs_download() {
        return None;
    }

    // What the file's content says it is comes ahead of what its name does, where the two
    // disagree: a `.docx` whose bytes are an MP4 is loaded as the video it is, and a format
    // no kind of this app previews is loaded as nothing at all — see `content_type` for
    // what settles that, and `load_media_of_kind` for where the kind is handed on.
    match hover.route.content {
        crate::formats::content_type::Content::Kind(_) => {}
        crate::formats::content_type::Content::Foreign => return None,
        crate::formats::content_type::Content::Unknown => {}
    }

    // What the file is, is the router's answer: one order, asked once, and the same one the hook
    // that admitted this hover asked (see `formats::routing`) — and asked of the entry already
    // in hand rather than of a second reading of it.
    if hover.routed_kind().is_none() {
        // A name no list claims has always been the picture path's, and a drawing among those is
        // still the drawing layer's: what it is, is its own header's answer rather than its
        // name's, and an `svg` a hand-edited list no longer names is a document this app can
        // draw. The hook refuses such a file before a hover reaches this far (see
        // `explorer_hook::is_media_file`), so this is the answer for the hover that came the
        // other way — through the content, which named no kind either.
        if hover.svg_document() {
            return webview_preview::draws(path).then(engine_svg_media);
        }

        return load_picture(path, max_width, max_height, preview_scale, &cancel);
    }

    load_media_of_kind(
        &hover,
        path,
        max_width,
        max_height,
        preview_scale,
        dpi,
        cancel,
    )
}

/// The loader for a file whose content named a kind of its own — see `content_type`.
///
/// It is the arm the chain above would have taken had the file been named what its
/// content says it is, reached by the kind rather than by the name: the same loaders, and
/// the same one for each kind. What every one of them reads is the file itself and never
/// the name it is under, which is what makes this a routing rather than a rename.
///
/// The gates are not asked here, exactly as they are not asked by the chain: the hook
/// asked them before a hover could reach this path at all.
pub(super) fn load_media_of_kind(
    hover: &HoverFacts,
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    dpi: u32,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    match hover.routed_kind() {
        Some(PreviewType::Videos) => {
            load_video_thumbnail(path, max_width, max_height, preview_scale)
        }
        // A book is one kind with two readers, and which of the two a file is, is asked of the
        // table that names them rather than assumed here: a page is the PDF engine's, a comic is
        // the first plate read out of the container it is published in, and a comic reached by
        // its *content* would otherwise be read as a page of a PDF that does not exist (see
        // `native_formats`). What cannot be read for — a configuration that will not open — is
        // answered as the page a book most often is.
        Some(PreviewType::Ebook) => match hover.book_job() {
            native_formats::NativeJob::Comic => {
                load_comic_page(path, max_width, max_height, preview_scale)
            }
            _ => load_pdf_first_page(path, max_width, max_height, preview_scale),
        },
        Some(PreviewType::Archives) => load_archive_preview(
            path,
            max_width,
            max_height,
            dpi,
            current_archive_options(),
            MediaType::Archive,
            &cancel,
        ),
        Some(PreviewType::Document) => {
            load_office_preview(path, max_width, max_height, preview_scale, &cancel)
                .or_else(|| load_engine_page_for_office(path, max_width, max_height, preview_scale))
        }
        Some(PreviewType::Libre) => libreoffice_render::rendered_page(path).and_then(|page| {
            load_engine_page(
                &page,
                MediaType::Libre,
                max_width,
                max_height,
                preview_scale,
            )
        }),
        Some(PreviewType::Magick) => {
            load_magick_picture(path, max_width, max_height, preview_scale, &cancel)
        }
        // An archive an engine listed is loaded as an archive: the listing it produced is in the
        // same cache under the same key, so the page is measured and painted from it without this
        // arm knowing where it came from — and a file the engine has not answered for yet is a
        // listing the cache does not hold, which is the wait the hover is already in.
        Some(PreviewType::Peazip) => load_archive_preview(
            path,
            max_width,
            max_height,
            dpi,
            current_archive_options(),
            MediaType::Peazip,
            &cancel,
        ),
        Some(PreviewType::Calibre) => calibre_render::rendered_page(path)
            .and_then(|page| load_book_page(&page, max_width, max_height, preview_scale)),
        Some(PreviewType::Design) => {
            load_design_preview(path, max_width, max_height, preview_scale)
        }
        // Which half of the drawing kind this is, is the name's to say here rather than the
        // content's: a document is drawn by the browser engine and a metafile by the drawing
        // layer, and the content has already answered that the file is a drawing at all.
        Some(PreviewType::Vector) => {
            if hover.svg_document() {
                webview_preview::draws(path).then(engine_svg_media)
            } else {
                load_vector_preview(path, max_width, max_height, preview_scale)
            }
        }
        // A page of HTML the engine draws is handed over the way a document is: the media carries
        // the kind and no frame, and the engine's window is the preview (see `engine_svg_media`).
        Some(PreviewType::Text) => {
            if hover.html_drawn_by_the_engine() {
                Some(engine_svg_media())
            } else {
                load_text_preview(path, max_width, max_height, dpi, current_text_options())
            }
        }
        // A sound: a card of what the file holds, painted like an archive's page. The facts are
        // the probe's and are already in hand — the measure that read them is what laid this
        // hover out (see `audio_box`) — and the player the card is drawn against is started by
        // the loop, where every other preview is put up.
        Some(PreviewType::Audio) => load_audio_card(path, max_width, max_height, dpi),
        Some(PreviewType::Fonts) => (font_preview::probe(path).is_some()
            && webview_preview::draws(path))
        .then(engine_font_media),
        Some(PreviewType::Images) => {
            load_picture(path, max_width, max_height, preview_scale, &cancel)
        }

        // A file whose content named no kind and whose name no list claims is the picture
        // path's, and the caller above has already answered that one — a name no list has ever
        // claimed is the picture it ends up decoded as.
        None => None,
    }
}

/// The picture path: the animated reader the file's own bytes call for, and everything else as
/// the still picture it is.
///
/// It is where the chain above ends for every name that reaches it, and the arm a picture the
/// content named is loaded by. Which reader is asked about it is the job the router gives a
/// picture (`native_formats::picture_job`), and that job is the file's own bytes first: a `.gif`
/// with one frame in it is a still here, and the animated reader is never asked about one. That
/// job is the file's answer alone, so nothing here takes the configuration or holds its lock
/// across the read of the head.
pub(super) fn load_picture(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    cancel: &Arc<AtomicBool>,
) -> Option<MediaData> {
    let job = native_formats::picture_job(path);

    match job {
        native_formats::NativeJob::AnimatedGif => {
            if let Some(media) = load_animated_gif(
                path,
                max_width,
                max_height,
                preview_scale,
                Arc::clone(cancel),
            ) {
                return Some(media);
            }
        }
        native_formats::NativeJob::AnimatedWebp => {
            if let Some(media) = load_animated_webp(
                path,
                max_width,
                max_height,
                preview_scale,
                Arc::clone(cancel),
            ) {
                return Some(media);
            }
        }
        native_formats::NativeJob::AnimatedApng => {
            if let Some(media) = load_animated_apng(
                path,
                max_width,
                max_height,
                preview_scale,
                Arc::clone(cancel),
            ) {
                return Some(media);
            }
        }
        native_formats::NativeJob::AnimatedJxl => {
            if let Some(media) = load_animated_jxl(
                path,
                max_width,
                max_height,
                preview_scale,
                Arc::clone(cancel),
            ) {
                return Some(media);
            }
        }
        native_formats::NativeJob::AnimatedHeif => {
            if let Some(media) = load_animated_heif(
                path,
                max_width,
                max_height,
                preview_scale,
                Arc::clone(cancel),
            ) {
                return Some(media);
            }
        }
        native_formats::NativeJob::Picture
        | native_formats::NativeJob::PictureCodec
        | native_formats::NativeJob::Text
        | native_formats::NativeJob::SvgDocument
        | native_formats::NativeJob::Metafile
        | native_formats::NativeJob::Eps
        | native_formats::NativeJob::FontSpecimen
        | native_formats::NativeJob::Pdf
        | native_formats::NativeJob::Comic
        | native_formats::NativeJob::Psd
        | native_formats::NativeJob::Project
        | native_formats::NativeJob::ArchiveZip
        | native_formats::NativeJob::ArchiveSevenZ
        | native_formats::NativeJob::ArchiveRar
        | native_formats::NativeJob::ArchiveTar
        | native_formats::NativeJob::ArchiveTarGz
        | native_formats::NativeJob::VideoMediaFoundation
        | native_formats::NativeJob::AudioMediaFoundation => {}
    }

    // What is left is a still: a picture that never moved, or one whose animated reader
    // answered nothing for it. Nothing is decoded twice to find that out.
    if cancel.load(Ordering::Acquire) {
        return None;
    }
    load_static_image(path, max_width, max_height, preview_scale)
}

/// Result from background image loading thread
pub(super) struct LoadResult {
    pub(super) generation: u64,
    pub(super) path: PathBuf,
    pub(super) media: Option<MediaData>,
    /// Nothing to draw yet, and a page on the way: the preview stays pending
    /// rather than being dropped, so the page has something to replace.
    pub(super) awaiting_render: bool,
}

/// A decode request consumed by the dedicated loader worker.
pub(super) struct LoadRequest {
    pub(super) generation: u64,
    pub(super) path: PathBuf,
    pub(super) max_width: u32,
    pub(super) max_height: u32,
    pub(super) preview_scale: PreviewScale,
    pub(super) dpi: u32,
    pub(super) cancel: Arc<AtomicBool>,
}

pub(super) type LoadRequestSlot = Arc<(Mutex<Option<LoadRequest>>, Condvar)>;

pub(super) fn queue_load_request(slot: &LoadRequestSlot, request: LoadRequest) {
    let (lock, cvar) = &**slot;
    if let Ok(mut pending) = lock.lock() {
        *pending = Some(request);
        cvar.notify_one();
    }
}

pub(super) fn clear_load_request(slot: &LoadRequestSlot) {
    let (lock, _) = &**slot;
    if let Ok(mut pending) = lock.lock() {
        *pending = None;
    }
}

pub(super) fn spawn_load_worker(
    request_slot: LoadRequestSlot,
    result_tx: Sender<LoadResult>,
) -> std::thread::JoinHandle<()> {
    std::thread::spawn(move || {
        // A PDF page is rendered through Windows.Data.Pdf and a picture of a format
        // this app has no decoder for is decoded by the codec Windows has, so this
        // thread needs an apartment before the first load asks for either.
        pdf_preview::initialize_apartment();
        wic_image::initialize_apartment();

        while RUNNING.load(Ordering::Acquire) {
            let mut request = {
                let (lock, cvar) = &*request_slot;
                let mut pending = match lock.lock() {
                    Ok(guard) => guard,
                    Err(_) => break,
                };

                while pending.is_none() && RUNNING.load(Ordering::Acquire) {
                    pending = match cvar.wait_timeout(pending, Duration::from_millis(200)) {
                        Ok((guard, _)) => guard,
                        Err(_) => return,
                    };
                }

                if !RUNNING.load(Ordering::Acquire) {
                    break;
                }

                match pending.take() {
                    Some(req) => req,
                    None => continue,
                }
            };

            // Coalesce any queued requests so we decode only the newest target.
            {
                let (lock, _) = &*request_slot;
                if let Ok(mut pending) = lock.lock() {
                    if let Some(newer) = pending.take() {
                        request.cancel.store(true, Ordering::Release);
                        request = newer;
                    }
                } else {
                    break;
                }
            }

            if request.cancel.load(Ordering::Acquire) {
                continue;
            }

            let media = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                load_media(
                    &request.path,
                    request.max_width,
                    request.max_height,
                    request.preview_scale,
                    request.dpi,
                    Arc::clone(&request.cancel),
                )
            }))
            .unwrap_or(None);

            // Every engine a hover can be waiting on is asked here, and a kind left out is
            // not merely a wait that is not shown: the loop reads this as "there is nothing
            // to draw and nothing coming", which is the branch that hides the window and
            // drops the hover — so the engine is never asked for its answer and the file
            // has no preview at all. The listing the PeaZip engine owes is such a wait, and so
            // is the page the ebook engine converts a book into (see `peazip_render_is_due` and
            // `calibre_render_is_due`).
            let awaiting_render = media.is_none()
                && (office_render_is_due(&request.path, request.max_width)
                    || libre_render_is_due(&request.path)
                    || magick_render_is_due(&request.path)
                    || peazip_render_is_due(&request.path)
                    || calibre_render_is_due(&request.path));

            let _ = result_tx.send(LoadResult {
                generation: request.generation,
                path: request.path.clone(),
                media,
                awaiting_render,
            });
        }
    })
}

/// How long a hover's load may run before the spinner is put up for it, as
/// `spinner_delay_ms` in `config.ini` names it.
///
/// Read when a load starts rather than captured once, so an edit applies to the next
/// hover without a restart. The window is hidden while a load runs, so a load that
/// finishes inside the delay has gone straight from nothing to the preview — what the
/// delay is for — and `0` is a spinner that goes up with the load; see `spinner_due`.
pub(super) fn load_spinner_delay() -> Duration {
    let millis = CONFIG
        .lock()
        .map(|config| sanitize_spinner_delay_ms(config.spinner_delay_ms))
        .unwrap_or(DEFAULT_SPINNER_DELAY_MS);

    Duration::from_millis(millis)
}

/// What placing a hover's preview again needs, kept on a pending load.
///
/// The size is not measured again when the preview follows the pointer: a cursor
/// moving along the item it belongs to finds the same media, and measuring it per
/// tick would re-read a header, a listing or a document sixty times a second to
/// learn what the hover already knew. What is recomputed is the place — that is
/// what the cursor decides.
///
/// A text preview is the one size measured again before the wait starts — its height
/// is the rows its lines wrap into at the width the box came out with — and what is
/// kept here is the size that measure answered with, so a re-placement steps around
/// the name for the rows the frame really has (see `text_preview_layout`).
#[derive(Clone, Copy)]
pub(super) struct HoverPlacement {
    pub(super) orig_dims: (u32, u32),
    /// The region the Explorer hook measured off the hovered item, and what kind of
    /// region it is — the item's own text, or a column the view draws every row's text
    /// in — which is what the placement is kept off (see `AvoidRegion`).
    pub(super) avoid: Option<AvoidRegion>,
    pub(super) follow_cursor: bool,
    pub(super) preview_scale: PreviewScale,
    /// Whether this placement is the waiting spinner's own box rather than a
    /// preview's: it is then placed at the pointer's own corner — a pointer gap off
    /// the cursor, in the quadrant the display has room for it, with the name it
    /// covers left alone — and kept there while the wait runs. See
    /// `compute_mouse_layout` and `waiting_placement`.
    pub(super) at_the_pointer_corner: bool,
}

/// The placement a hover's waiting spinner is given: the arc's own box at the pointer's
/// own corner — a pointer gap off the hand — whatever the preview it is waiting for is.
///
/// It is the placement a document waiting on a page has always been placed by, and
/// every other kind of load is shown the same way: what a hover shows while it waits
/// is the arc at the hand that hovered the file — which is what says the wait is for
/// the file under it — rather than a preview-sized frame with a spinner in the middle
/// of it, placed where the preview will land and saying nothing about the hand.
///
/// The gap is what keeps the arc off the cursor. The window it is drawn in is the one
/// the pointer's own messages land on, and a hand kept off the arc is a hand still
/// clicking and probing the file the wait is for.
pub(super) fn waiting_placement(placement: HoverPlacement) -> HoverPlacement {
    HoverPlacement {
        orig_dims: (office_preview::WAITING_BOX, office_preview::WAITING_BOX),
        preview_scale: PreviewScale::Percent(100),
        at_the_pointer_corner: true,
        ..placement
    }
}

/// What placing a pending load again moved: the spinner's own box, which is what is
/// on screen while the load runs, and the preview's, which is what the media lands at
/// and what an engine drawing it is told to draw in.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) struct Followed {
    pub(super) spinner: bool,
    pub(super) preview: bool,
}

/// Tracks a pending background load so we can show the spinner while it runs.
pub(super) struct PendingLoad {
    pub(super) generation: u64,
    /// The hide count this load was started under. A hide after it is the pointer
    /// having left the file while the load ran, and is what the load's reveal is
    /// checked against so the frame it lands with is dropped rather than shown
    /// (see `HIDDEN_EPOCH`).
    pub(super) hide_epoch: u64,
    /// The file this load is for, which is what decides whether the engine plays it
    /// and what the engine is pointed at.
    pub(super) path: PathBuf,
    pub(super) started: Instant,
    pub(super) pos_x: i32,
    pub(super) pos_y: i32,
    pub(super) width: u32,
    pub(super) height: u32,
    /// The room the display the hover is on has (`ScreenBounds::room`), which is the box
    /// a page on its way is asked for in and the box a picture the image converter
    /// develops is developed at. A slide is exported at the width the render is asked
    /// for, so the room the display has is the sharpest page that display can show; and
    /// a picture is developed *at* the size it is then drawn at, so a room smaller than
    /// that is a picture that can never be shown any larger.
    ///
    /// It is deliberately not the room this load's own layout came out at. A hover
    /// waiting on an engine is laid out as the spinner's own box at the pointer (see
    /// `waiting_placement`), and the room that layout comes out at is the corner of the
    /// display the spinner was put in — a box that says nothing about how large the
    /// preview will be drawn, and for a picture a ceiling that would leave it thumbnail
    /// sized for good.
    pub(super) room: (u32, u32),
    pub(super) spinner_shown: bool,
    /// How long this load may run before the spinner is put up for it, read from
    /// `spinner_delay_ms` when the load started. Every kind of wait is given the same
    /// one — a decode, a page Office is rendering, a browser that has to start.
    pub(super) spinner_delay: Duration,
    /// Where the spinner is drawn and placed while this load runs — the arc's own box
    /// at the pointer's corner (see `waiting_placement`) — and not the box the preview
    /// will arrive in (`pos_x`, `pos_y`, `width`, `height`).
    pub(super) spinner_pos: (i32, i32),
    pub(super) spinner_side: u32,
    /// The mouse hover this load came from, if it was one. A preview that is still
    /// on its way follows the pointer, so it is placed again for a cursor that has
    /// moved along the item since; a keyboard hover's placement belongs to the
    /// item and carries none.
    pub(super) placement: Option<HoverPlacement>,
    /// Whether this load is replacing what is already on screen — the page that
    /// arrived for the hover that is up — rather than opening a new preview. An
    /// upgrade never shows the spinner: what is there stays where it is, at its own
    /// size, until the page is ready.
    pub(super) upgrade: bool,
    /// Whether this load is waiting on an engine to produce what no reader here could
    /// draw: the page Office is rendering, the page the render engine beside it is
    /// converting, the picture the image converter is developing, the listing Peazip is
    /// printing. Set where a load comes back with nothing to draw and one of those
    /// engines is owed the file (see `awaiting_render` in the loader).
    ///
    /// It is what the cap on waiting is read against, and it is on the wait rather than
    /// on the request for a reason: an engine answers a file once, so a hover whose
    /// request was folded into one already in flight — the file the engine is working on
    /// is the file this hover asked about — has nothing left that would ever answer it,
    /// and a wait that is not bounded is a spinner that runs for good (see
    /// `OFFICE_RENDER_WAIT_SECS`).
    ///
    /// A video's probe is the other wait marked this way, and for the reason above rather
    /// than a reason of its own: what it waits on is outside this side, and a probe that
    /// answered nothing would leave the hover with nothing here that ends it. The probe is
    /// given a bound of its own (`VIDEO_PROBE_TIMEOUT_SECS`), so the cap behind this is
    /// the second line rather than the first.
    pub(super) awaiting_engine: bool,
}

impl PendingLoad {
    /// Whether the spinner is due for this load: once it has run for the delay
    /// `spinner_delay_ms` names, which is the same moment for every kind of wait — a
    /// decode, a page Office is rendering, a browser that has to start.
    ///
    /// An upgrade is never due — what is on screen stays where it is, at its own
    /// size, until the page replaces it — and a load that already has its spinner
    /// up is not due again.
    pub(super) fn spinner_due(&self) -> bool {
        !self.spinner_shown && !self.upgrade && self.started.elapsed() >= self.spinner_delay
    }

    /// Place this load's preview — and the spinner standing in for it — again for
    /// `cursor`, when it is one that follows the pointer: the size the hover measured,
    /// its `Avoid` region and its scale, and the display the pointer is on now — the
    /// room it has (`dpi` is that display's scale, and `cursor` is where it is).
    /// Answers what moved, so the window is only moved when the spinner's own box has
    /// and an engine is only told when the preview's has.
    ///
    /// The two are not the same place. A wait is the arc's own box at the pointer's
    /// corner whatever the preview is (see `waiting_placement`), so a wait that
    /// follows a moving hand is the wait at the hand, while the preview's place is
    /// read for the media that lands at it and for an engine told where to draw.
    ///
    /// A load that answered nothing, or a keyboard hover, is left where it is.
    pub(super) fn follow_pointer(&mut self, cursor: POINT, dpi: u32) -> Followed {
        let mut followed = Followed {
            spinner: false,
            preview: false,
        };
        let Some(placement) = self.placement else {
            return followed;
        };

        let bounds = work_area_at(cursor.x, cursor.y);

        // The room an engine is asked for follows the hand the preview does, because the
        // display the hand is on is the one the hover is now waiting on: a pointer that has
        // crossed to another display is a picture developed for that display's room.
        self.room = bounds.room();

        if let Some(layout) = compute_mouse_layout(cursor.x, cursor.y, placement, bounds, dpi) {
            followed.preview = (layout.pos_x, layout.pos_y) != (self.pos_x, self.pos_y)
                || (layout.preview_w, layout.preview_h) != (self.width, self.height);
            if followed.preview {
                self.pos_x = layout.pos_x;
                self.pos_y = layout.pos_y;
                self.width = layout.preview_w;
                self.height = layout.preview_h;
            }
        }

        let waiting = waiting_placement(placement);
        if let Some(layout) = compute_mouse_layout(cursor.x, cursor.y, waiting, bounds, dpi) {
            followed.spinner = (layout.pos_x, layout.pos_y) != self.spinner_pos
                || layout.preview_w != self.spinner_side;
            if followed.spinner {
                self.spinner_pos = (layout.pos_x, layout.pos_y);
                self.spinner_side = layout.preview_w;
            }
        }

        followed
    }
}

/// Whether an engine's answer belongs to the wait under the spinner even though it names
/// another hover: the answer is about the file that wait is on, it is an answer that
/// produced what that wait is for, and the wait is one an engine owes the file.
///
/// An engine answers a file once — a request made while it was already working on the
/// same file is folded into that work, and what comes back names the request before it.
/// Read as an answer about a hover that has gone, it would leave the wait with no request
/// that could answer it and no cap that could end it, and the page, picture or listing it
/// names — which is in hand — would not be shown until the file was hovered afresh. It is
/// what `hovered` is not: an answer that is the wait's own without naming its generation
/// (see `awaiting_engine`).
pub(super) fn answer_belongs_to_the_wait(
    ready_path: &Path,
    ready_ok: bool,
    shown: Option<&Path>,
    pending: Option<&PendingLoad>,
) -> bool {
    ready_ok
        && shown == Some(ready_path)
        && pending.is_some_and(|pl| pl.awaiting_engine && pl.path == ready_path)
}

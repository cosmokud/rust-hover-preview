//! One loader per kind of file, each answering a room with a `MediaData`: pictures, engine
//! drawings, comics, vectors, documents, text, archives and a film's first frame.

use super::*;

/// Load a static image (JPG, PNG, BMP, static WebP, etc.)
///
/// Decoding one costs a full decode, a resample and two whole-buffer conversions,
/// and a file list is a place a pointer is swept back and forth over, so a frame
/// that has already been built is handed back rather than built again.
pub(super) fn load_static_image(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    // The header carries the image's own size, which is what the layout and the
    // resample are computed from, so it is read first: it names the box the frame
    // would be decoded into, and so the key that frame is held under.
    let dimensions = image_dimensions_with_header_check(path);

    // A texture is a picture to everything below this line, and a kind of its own to what
    // draws it: what a `.dds` preview is composited over is the tray's texture backdrop
    // rather than a picture's (see the `Background` submenu and `dds_image`).
    let kind = if dds_image::is_dds_file(path) {
        MediaType::Dds
    } else {
        MediaType::StaticImage
    };

    let cache_key = dimensions.map(|(width, height)| {
        let (target_width, target_height) =
            scale_dimensions(width, height, max_width, max_height, preview_scale);

        ImageCacheKey {
            path: path.to_path_buf(),
            version: file_version(path),
            width: target_width,
            height: target_height,
        }
    });

    if let Some(key) = cache_key.as_ref() {
        if let Some(frame) = image_cache_get(key) {
            return Some(static_image_media(frame, kind));
        }
    }

    // A header that would not report its dimensions is not a reason to refuse the
    // file: the decoder has the last word on whether it is an image at all, and a
    // frame measured this way is simply not held.
    //
    // A picture of a format this app's own decoder does not read is asked of the codec
    // Windows has for it instead, and that one is handed the box the layout planned
    // rather than the file's own size: what it decodes is the preview, and what it
    // hands back is already in the pixel order the frame is composed in, so the
    // resample and the two conversions are not paid for either; see `wic_image`.
    let (pixels, target_width, target_height) = if wic_image::is_codec_file(path) {
        let (width, height) = match cache_key.as_ref() {
            Some(key) => (key.width, key.height),
            // A picture whose size would not be read has no size to take a share of,
            // so what it is decoded into is the box the layout has.
            None => (max_width, max_height),
        };

        // The WebP codec is the one of them that is a Store package rather than
        // something Windows has, so a machine without it — a Windows 10 machine,
        // usually — is answered by libwebp instead, which is in the binary for the
        // picture that moves; see `webp_image`. A `.dds` is the one of them the codec
        // reads a smaller set of than the format holds, so what it has no answer for —
        // the uncompressed formats, BC4 and BC5 — is asked of a decoder of this app's
        // own; see `dds_image`. Both are guarded by the file's own header, so being
        // asked about a picture that is neither costs a header rather than a file.
        let pixels = wic_image::decode(path, width, height)
            .or_else(|| dds_image::decode(path, width, height))
            .or_else(|| webp_image::decode(path, width, height))?;

        (pixels, width, height)
    } else {
        let img = decode_image_with_header_check(path)?;

        // A picture whose samples are light rather than levels — an EXR, a Radiance HDR —
        // is brought into eight bits before anything else is done with it. It is the one
        // thing that has to see the whole of the file's range, and the order is also the
        // cheaper one: what follows is a resample, and resampling one byte to the channel
        // is a quarter of the memory and a fraction of the time of resampling four bytes
        // of float (see `tone_map`).
        let img = match tone_mapped_image(&img) {
            Some(toned) => image::DynamicImage::ImageRgba8(toned),
            None => img,
        };

        let (orig_width, orig_height) = img.dimensions();
        let (width, height) = match cache_key.as_ref() {
            Some(key) => (key.width, key.height),
            None => scale_dimensions(
                orig_width,
                orig_height,
                max_width,
                max_height,
                preview_scale,
            ),
        };

        let resized = if width != orig_width || height != orig_height {
            img.resize_exact(width, height, image::imageops::FilterType::Triangle)
        } else {
            img
        };

        let rgba = resized.to_rgba8();

        (rgba_to_bgra(rgba.as_raw()), width, height)
    };

    let frame = Arc::new(ImageFrame::new(pixels, target_width, target_height, 0));

    if let Some(key) = cache_key {
        image_cache_put(key, Arc::clone(&frame));
    }

    Some(static_image_media(frame, kind))
}

/// The picture a design document is previewed from.
///
/// Three answers, in one order. Where an engine is installed the document itself is drawn
/// by it — `libreoffice_render` — and what comes back is a page, sharp at whatever size the
/// preview is shown at. Behind it are the pictures these formats keep of their own work,
/// which is what a machine without the engine is answered with, and what a document the
/// engine cannot read is answered with as well: the planar merged picture Photoshop writes
/// at the end of a file, decoded by `psd_image`, and the picture a project container, a
/// CorelDRAW document or a PostScript one holds, read by `project_image`, `cdr_image` and
/// `eps_image`. Which of those is asked is settled by the file's own header rather than by
/// its name, and each hands back the picture decoded into the box the layout planned rather
/// than at the size it is — which for a layered document can be enormous, and is why the
/// box is what is asked for rather than the size.
///
/// Everywhere else a design preview is a frame of this app's own: it is composed like a
/// picture, held in the image cache like one under the size it was made for, drawn over the
/// backdrop the tray keeps for this kind — `design_background`, a setting of its own — and
/// laid out at the share of the display `design_scale` names, which is what
/// `effective_preview_scale` has already read by the time this is called.
pub(super) fn load_design_preview(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    let (source_width, source_height) = design_dimensions(path)?;
    let (target_width, target_height) = scale_dimensions(
        source_width,
        source_height,
        max_width,
        max_height,
        preview_scale,
    );

    let key = ImageCacheKey {
        path: path.to_path_buf(),
        version: file_version(path),
        width: target_width,
        height: target_height,
    };

    if let Some(frame) = image_cache_get(&key) {
        return Some(static_image_media(frame, MediaType::Design));
    }

    // The page the engine drew comes first, where there is one; the readers below are the
    // fallback for a machine without the engine and for a document it has not drawn — a
    // name this list and the `[libre]` list both hold is drawn by the engine, and this is
    // only reached for one of those where these previews are what is being asked for. What
    // the engine drew comes back at the size the page fitted into the box rather than at
    // the box, so the frame is built from what was drawn.
    let (pixels, width, height) = if let Some(page) = libreoffice_render::rendered_page(path) {
        pdf_preview::render_first_page(&page, target_width, target_height)?
    } else {
        let pixels = if psd_image::is_psd_file(path) {
            psd_image::decode(path, target_width, target_height)
        } else {
            project_image::decode(path, target_width, target_height)
                .or_else(|| eps_image::decode(path, target_width, target_height))
        }?;

        (pixels, target_width, target_height)
    };

    let frame = Arc::new(ImageFrame::new(pixels, width, height, 0));

    image_cache_put(key, Arc::clone(&frame));

    Some(static_image_media(frame, MediaType::Design))
}

/// The page a render engine drew for a document, as a frame of `kind`.
///
/// What the engine hands back is a PDF, and this is the whole of what is asked of it: the
/// page's own size, the box the layout would place that size in, and page 1 rendered into
/// it — drawn at the size it is shown at rather than scaled up from a picture, which is the
/// reason a document an engine drew is worth asking for at all. The pages are cached by the
/// PDF path itself, under the size they were drawn at.
pub(super) fn load_engine_page(
    page: &Path,
    kind: MediaType,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    let (page_width, page_height) = pdf_preview::page_dimensions(page)?;
    let (target_width, target_height) = scale_dimensions(
        page_width,
        page_height,
        max_width,
        max_height,
        preview_scale,
    );
    let (pixels, width, height) =
        pdf_preview::render_first_page(page, target_width, target_height)?;

    Some(static_image_media(
        Arc::new(ImageFrame::new(pixels, width, height, 0)),
        kind,
    ))
}

/// The first plate of a comic, drawn into the box the layout measured it for.
///
/// It is `load_pdf_first_page` for a book that is a container rather than a document: what is drawn
/// is a picture that is already in the file, decoded at the size the box asks for and composited
/// over the backdrop the book kind is drawn over. Nothing is waited on and nothing is converted —
/// the plate is read out of the archive, decoded and scaled in one go — so this is the one book of
/// the kind whose preview is made on the side that shows it.
pub(super) fn load_comic_page(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    let (page_width, page_height) = comic_preview::dimensions(path)?;

    let (target_width, target_height) = scale_dimensions(
        page_width,
        page_height,
        max_width,
        max_height,
        preview_scale,
    );
    let pixels = comic_preview::decode(path, target_width, target_height)?;

    Some(static_image_media(
        Arc::new(ImageFrame::new(pixels, target_width, target_height, 0)),
        MediaType::Comic,
    ))
}

/// The page a converted book is drawn from, drawn into the box the layout measured it for.
///
/// It is `load_engine_page` for the one engine whose page is a whole book rather than a page of a
/// document: what is drawn is the first page of that book which says anything about it, and the box
/// is that page's own size, so a book whose first page is a cover of one colour is previewed from
/// the page behind it rather than as the colour (see `pdf_preview::book_page`).
pub(super) fn load_book_page(
    page: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    let book = pdf_preview::book_page(page)?;
    let (page_width, page_height) = book.size;

    let (target_width, target_height) = scale_dimensions(
        page_width,
        page_height,
        max_width,
        max_height,
        preview_scale,
    );
    let (pixels, width, height) = pdf_preview::render_book_page(page, target_width, target_height)?;

    Some(static_image_media(
        Arc::new(ImageFrame::new(pixels, width, height, 0)),
        MediaType::Calibre,
    ))
}

/// The page the render engine has drawn for an Office document — the fallback for a document
/// whose own application is not installed, and the page itself where the tray has asked the
/// engine for every Office document. Shown as the Office document it is either way.
///
/// Nothing is converted here, and nothing is waited on: a document the engine has not drawn
/// yet is answered with nothing, which is the wait the hover is already in — the loop has
/// asked the engine for the page, and the hover is replayed when it lands (see
/// `libre_render_is_due`). The page that *is* there is read at the share the layout measured
/// it for, and that share is the Office kind's: the file is what it is whichever engine drew
/// it (see `effective_preview_scale`).
///
/// A page the engine drew while it was the one being asked is not read once the choice is the
/// application's: a page is what the engine that is drawn by is asked for, and a document the
/// application is drawing is not shown the other engine's page until it is ready (see
/// `office_formats::page_engine`).
pub(super) fn load_engine_page_for_office(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    if office_formats::page_engine(path) != Some(OfficeEngine::LibreOffice) {
        return None;
    }

    let page = libreoffice_render::rendered_page(path)?;
    load_engine_page(
        &page,
        MediaType::Office,
        max_width,
        max_height,
        preview_scale,
    )
}

/// The picture the ImageMagick engine developed for a file, drawn as the picture it is.
///
/// Nothing is converted here, and nothing is waited on. Three things answer a hover in the
/// order they are worth asking:
///
/// * the frame this app has already built for this file at this box, which is the image cache
///   every other picture is kept in — a hit costs a lookup and no engine at all, which is what
///   a pointer swept back and forth over a folder of raws is answered with;
/// * the picture the engine developed, which is what the frame is built from: the one held in
///   the engine's hand for the hover that asked, or the page it wrote down for the hovers after
///   it — decoded from the bytes the engine wrote the picture as rather than from the source
///   file, resampled into the box the layout planned, and held in that same cache under the
///   file, its version and that box;
/// * and nothing at all, which is the wait the hover is already in — the loop has asked the
///   engine for the picture, and the hover is replayed when the answer lands (see
///   `magick_render_is_due`).
///
/// What the frame is composited over is `image_background` and what share of its own size it
/// is drawn at is the picture's, because that is what it is: a picture of this app's, in the
/// format a frame is composed in, held like any other.
pub(super) fn load_magick_picture(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    cancel: &Arc<AtomicBool>,
) -> Option<MediaData> {
    // The size the engine developed this file at — which is the picture's own size rather
    // than the file's, since what the engine writes is a picture fitted into the room — and
    // what the frame is keyed by, together with the box it is drawn in.
    let cache_key = match imagemagick_render::dimensions(path) {
        Some((width, height)) => {
            let (target_width, target_height) =
                scale_dimensions(width, height, max_width, max_height, preview_scale);

            Some(ImageCacheKey {
                path: path.to_path_buf(),
                version: file_version(path),
                width: target_width,
                height: target_height,
            })
        }
        // A file the engine has developed nothing for is a file with no size to key a frame
        // by: what is coming is the engine's answer, and the frame it is drawn as is built
        // from that answer rather than placed in the cache under a size nothing knows.
        None => None,
    };

    if let Some(key) = cache_key.as_ref() {
        if let Some(frame) = image_cache_get(key) {
            return Some(static_image_media(frame, MediaType::Magick));
        }
    }

    // The picture the engine developed: the one in its hand, which is what a hover that asked
    // for it is waiting for — a hover that has moved on leaves it where it is, since what it is
    // waiting for is its own replay — or the page it wrote for the hovers after that one, which
    // is every hover since a restart.
    let developed = match imagemagick_render::take_developed(path) {
        Some(developed) => developed,
        None => imagemagick_render::read_page(path)?,
    };
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    let (target_width, target_height) = scale_dimensions(
        developed.width,
        developed.height,
        max_width,
        max_height,
        preview_scale,
    );
    // A page that is there but does not decode is not a picture: it is given up, so that the
    // file behind it is developed again rather than read into the same nothing on every hover.
    let Some(image) = decode_png(&developed.png) else {
        imagemagick_render::forget_page(path);
        return None;
    };
    let (orig_width, orig_height) = image.dimensions();
    let resized = if target_width != orig_width || target_height != orig_height {
        image.resize_exact(
            target_width,
            target_height,
            image::imageops::FilterType::Triangle,
        )
    } else {
        image
    };

    let rgba = resized.to_rgba8();
    let frame = Arc::new(ImageFrame::new(
        rgba_to_bgra(rgba.as_raw()),
        target_width,
        target_height,
        0,
    ));

    if let Some(key) = cache_key {
        image_cache_put(key, Arc::clone(&frame));
    }

    Some(static_image_media(frame, MediaType::Magick))
}

/// The picture the engine wrote, decoded from the bytes it wrote it as — under the budget
/// every other decode of this app is answered under, and with the format asked for rather
/// than taken on trust.
pub(super) fn decode_png(png: &[u8]) -> Option<image::DynamicImage> {
    let mut reader = image::ImageReader::new(std::io::Cursor::new(png))
        .with_guessed_format()
        .ok()?;
    reader.limits(image_decode_limits());

    reader.decode().ok()
}

/// The drawing a vector file is previewed from.
///
/// Two readers answer for these files and both hand back a frame of the same kind: the
/// records an `.eps` carries, and the records a `.wmf` or an `.emf` is. Both ask the file
/// itself what it is rather than trusting its name, so which is asked first is only a
/// question of which is cheaper to turn down — the metafile reader is asked first for the
/// two names it is written as, and the encapsulated PostScript reader first for everything
/// else, since either of them refuses a file that is not its own after a header.
///
/// What comes back is the drawing replayed at the box the layout planned rather than a
/// picture resampled into it, which is what makes a preview of one sharp at any size the
/// display has. A file carrying a picture instead — a TIFF preview, which is what some
/// writers leave in an `.eps` — is resampled the way every other picture is.
pub(super) fn load_vector_preview(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    let (source_width, source_height) = vector_dimensions(path)?;
    let (target_width, target_height) = scale_dimensions(
        source_width,
        source_height,
        max_width,
        max_height,
        preview_scale,
    );

    let key = ImageCacheKey {
        path: path.to_path_buf(),
        version: file_version(path),
        width: target_width,
        height: target_height,
    };

    if let Some(frame) = image_cache_get(&key) {
        return Some(static_image_media(frame, MediaType::Vector));
    }

    let pixels = if metafile_image::is_metafile_name(path) {
        metafile_image::decode(path, target_width, target_height)
            .or_else(|| eps_image::decode(path, target_width, target_height))
    } else {
        eps_image::decode(path, target_width, target_height)
            .or_else(|| metafile_image::decode(path, target_width, target_height))
    }?;

    let frame = Arc::new(ImageFrame::new(pixels, target_width, target_height, 0));

    image_cache_put(key, Arc::clone(&frame));

    Some(static_image_media(frame, MediaType::Vector))
}

/// Render the first page of a PDF through the PDF engine built into Windows.
pub(super) fn load_pdf_first_page(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    let (page_width, page_height) = pdf_preview::page_dimensions(path).unwrap_or((
        pdf_preview::DEFAULT_PAGE_WIDTH,
        pdf_preview::DEFAULT_PAGE_HEIGHT,
    ));
    let (target_width, target_height) = scale_dimensions(
        page_width,
        page_height,
        max_width,
        max_height,
        preview_scale,
    );

    let (pixels, width, height) =
        pdf_preview::render_first_page(path, target_width, target_height)?;

    let frame = ImageFrame::new(pixels, width, height, 0);

    Some(MediaData {
        frames: vec![Arc::new(frame)],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::Pdf,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
    })
}

/// Render a page of an Office document into the box the layout planned.
///
/// One source, and it is one this side reads rather than produces: the page the render tier
/// is holding in memory for the document — drawn by Office itself, or by the render engine
/// beside it where the document's own application is not installed, and shown under this
/// kind either way. A document with no page yet is answered with nothing, which is the wait
/// the hover is in, and the page is asked for by the loop the moment the hover is up (see
/// `request_office_render` and `libre_render_is_due`). The source's own aspect ratio is
/// preserved inside the box, so a page that is not the shape the layout assumed is
/// letterboxed instead of stretched.
pub(super) fn load_office_preview(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    cancel: &Arc<AtomicBool>,
) -> Option<MediaData> {
    let (source_width, source_height) =
        office_preview::measure(path).unwrap_or_else(|| office_formats::default_page_size(path));
    let (target_width, target_height) = scale_dimensions(
        source_width,
        source_height,
        max_width,
        max_height,
        preview_scale,
    );

    let (pixels, width, height) =
        office_preview::render(path, target_width, target_height, Some(cancel))?;

    let frame = ImageFrame::new(pixels, width, height, 0);

    Some(MediaData {
        frames: vec![Arc::new(frame)],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::Office,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
    })
}

/// Render a text file into the box the layout planned for it.
///
/// Unlike an image, text is not scaled to the box: the font is a fixed,
/// display-scaled size, and the box decides how many lines and columns are shown.
/// That is why the caller hands over the planned preview size rather than the
/// free space around the cursor — the two are the same thing to this renderer,
/// and using the planned size keeps the painted frame and the window in step.
pub(super) fn load_text_preview(
    path: &Path,
    width: u32,
    height: u32,
    dpi: u32,
    options: TextPreviewOptions,
) -> Option<MediaData> {
    let frame = text_preview::render_scrolled(path, 0, width, height, dpi, options, None)?;

    // Any text preview in full mode keeps its state, scrollable or not: the
    // pointer rests on it, its text can be selected, and a document that happens to
    // fit is simply one that cannot be scrolled.
    let state = options.full_mode.then(|| TextPreviewState {
        path: path.to_path_buf(),
        options,
        dpi,
        width: frame.width,
        height: frame.height,
        first_line: frame.first_line,
        visible_lines: frame.visible_lines,
        scrollable_lines: frame.scrollable_lines,
        scrollbar: frame.scrollbar,
        dragging: false,
        lines: frame.lines,
        selection: None,
        selecting: false,
    });

    let frame = ImageFrame::new(frame.pixels, frame.width, frame.height, 0);

    Some(MediaData {
        frames: vec![Arc::new(frame)],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::Text,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: state,
    })
}

/// Load an archive's contents as a page of its own, the way a text preview is
/// loaded: measured first, then painted into exactly the box the layout planned.
///
/// Which kind the page comes out as is the caller's to say, because the same page is drawn for
/// two of them: an archive this app read itself is shown under `Archives`, and one an engine
/// listed under `Peazip` — the same reader of the same listing and the same painted frame, with
/// only the gate over the preview on screen differing between them.
pub(super) fn load_archive_preview(
    path: &Path,
    width: u32,
    height: u32,
    dpi: u32,
    options: ArchivePreviewOptions,
    media_type: MediaType,
    cancel: &AtomicBool,
) -> Option<MediaData> {
    let (pixels, width, height) = archive_preview::render(path, width, height, dpi, options)?;
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    Some(MediaData {
        frames: vec![Arc::new(ImageFrame::new(pixels, width, height, 0))],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
    })
}

/// Extract video thumbnail using ffmpeg and create frames for preview
pub(super) fn load_video_thumbnail(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    // Which engine plays this file is settled here, once, and everything below follows from it:
    // the media engine Windows has draws it into this app's own window, and FFmpeg's player draws
    // it in a window of its own (see `media_engine_plays`). The third answer is what this arm is
    // — a file nothing on this machine will play has no preview at all, which is the whole of
    // what a name only `[ffmpeg]` carries means on a machine with no FFmpeg, and what a name of
    // `[video]` means there once the engine has turned the file down. A box nothing would ever be
    // drawn into is not a preview, and the load that returned nothing is the branch that takes
    // the window down rather than one that fills it with a placeholder.
    let native = match video_route(path) {
        VideoRoute::MediaEngine => true,
        VideoRoute::Ffplay => false,
        VideoRoute::NoPreview => return None,
    };

    let geometry = match probe_video_geometry(path) {
        ProbedGeometry::Measured(geometry) => geometry,
        // No geometry, and the engine that would play the file cannot open it either: a
        // file this machine has no reader for is answered with no preview rather than with
        // a 16:9 box nothing would be drawn into.
        ProbedGeometry::Unmeasurable if native => return None,
        // FFmpeg is the one that would play it, so the box the layout uses is the one it
        // has always used for a file ffprobe could not measure.
        ProbedGeometry::Unmeasurable => VideoGeometry {
            width: 1920,
            height: 1080,
            frame_width: 1920,
            frame_height: 1080,
            crop: None,
            duration: None,
            // A file nothing could read is a file whose subtitle streams are unknown, which is
            // answered the same way as a file known to have none: the player is told nothing,
            // because a `-sst` guessed at is a specifier it may refuse outright (see
            // `video_subtitles`).
            subtitles: SubtitleStreams::default(),
            // And no sidecar either, which is the same answer for the same reason: this box is
            // built here rather than read out of the cache, and the only thing that answers the
            // question — the folder beside the film — is a read this thread has no reason to make
            // for a file nothing could read in the first place (see `video_sidecar`).
            sidecar: None,
            // And nothing derived either, for the same reason once more: no extraction runs for
            // a film nothing could read, so there are no small files and no pass to have failed
            // (see `subtitle_files`).
            derived: None,
            subtitle_extraction_failed: false,
        },
    };

    let (target_width, target_height) = scale_dimensions(
        geometry.width,
        geometry.height,
        max_width,
        max_height,
        preview_scale,
    );

    // Create a placeholder frame (dark gray) while video plays
    let placeholder_pixels = vec![40u8; (target_width * target_height * 4) as usize];

    let frame = ImageFrame::new(placeholder_pixels, target_width, target_height, 0);

    Some(MediaData {
        frames: vec![Arc::new(frame)],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: if native {
            MediaType::NativeVideo
        } else {
            MediaType::Video
        },
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
    })
}

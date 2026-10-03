use super::*;

/// The whole path a hover takes for a document that is already on disk: measure
/// it, place it, load it. Ignored, and driven by `RHP_OFFICE_PROBE` —
/// `$env:RHP_OFFICE_PROBE = "C:\docs\one.xlsx"; cargo test -- --ignored --nocapture office_hover_probe`
/// — for a document whose preview does not appear.
#[test]
#[ignore = "reads the files named in RHP_OFFICE_PROBE"]
fn office_hover_probe() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    // The drawing half of the app runs in a multithreaded apartment.
    pdf_preview::initialize_apartment();

    let Ok(list) = std::env::var("RHP_OFFICE_PROBE") else {
        println!("set RHP_OFFICE_PROBE to one or more paths, separated by ';'");
        return;
    };

    // A display with a cursor on it, which is what a hover arrives with.
    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };
    let (cursor_x, cursor_y) = (900, 500);
    let dpi = 96;

    for path in list
        .split(';')
        .map(str::trim)
        .filter(|path| !path.is_empty())
    {
        let path = PathBuf::from(path);
        println!("\n--- {} ---", path.display());

        // Which engine draws a page for this document, which is the question a preview
        // that arrives smaller than it should turn on: a page the render engine beside
        // Office drew is a page of the document's own size, and a hover laid out as the
        // wait for one places it at the spinner's box (see `libre_formats`).
        println!(
            "engines: office tier = {}, chosen engine = {:?}, render engine = {:?}, \
                 application installed = {}",
            office_render_is_due(&path, 800),
            office_formats::page_engine(&path),
            libre_formats::engine_page_kind(&path),
            office_formats::app_installed(&path),
        );
        println!(
            "render engine page: {}, refused = {}",
            libreoffice_render::rendered_page(&path)
                .map(|page| page.display().to_string())
                .unwrap_or_else(|| "none".to_string()),
            libreoffice_render::refused(&path)
        );

        let configured = current_hover_scales().picture;
        let scale = effective_preview_scale(&path, current_hover_scales());
        println!("scale: configured {configured:?}, effective {scale:?}");

        let Some(dimensions) = media_dimensions(&path, bounds, dpi) else {
            println!("measure: nothing — the hover shows no preview");
            continue;
        };
        println!("measure: {dimensions:?}");

        let Some(layout) = compute_mouse_layout(
            cursor_x,
            cursor_y,
            HoverPlacement {
                orig_dims: dimensions,
                avoid: None,
                follow_cursor: false,
                preview_scale: scale,
                at_the_pointer_corner: false,
            },
            bounds,
            dpi,
        ) else {
            println!("layout: none — the hover shows no preview");
            continue;
        };
        println!(
            "layout: {}x{} at ({}, {}), free room {}x{}",
            layout.preview_w,
            layout.preview_h,
            layout.pos_x,
            layout.pos_y,
            layout.max_width,
            layout.max_height
        );

        let cancel = Arc::new(AtomicBool::new(false));
        let started = Instant::now();
        match load_media(
            &path,
            layout.max_width,
            layout.max_height,
            scale,
            dpi,
            Arc::clone(&cancel),
        ) {
            Some(media) => println!(
                "loaded: {}x{}, {} frame(s), in {:?}",
                media.current_width(),
                media.current_height(),
                media.frames.len(),
                started.elapsed()
            ),
            None => println!(
                "loaded: nothing — the hover blinks, in {:?}",
                started.elapsed()
            ),
        }
    }
}

/// The whole path a hover takes for a document an installed render engine draws, which
/// is the one path with a wait in the middle of it: measured as the wait for a page, the
/// page asked for, the wait watched, and then — when the engine has drawn it — measured
/// and loaded from the page. It also answers the order of the lists, which is what
/// decides whether the engine is asked about a file at all. Ignored, and driven by
/// `RHP_LIBRE_PROBE` —
/// `$env:RHP_LIBRE_PROBE = "C:\art\logo.cdr"; cargo test -- --ignored --nocapture libre_hover_probe`
/// — for a document whose preview does not appear.
#[test]
#[ignore = "reads the files named in RHP_LIBRE_PROBE and starts the installed LibreOffice"]
fn libre_hover_probe() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    pdf_preview::initialize_apartment();

    let Ok(list) = std::env::var("RHP_LIBRE_PROBE") else {
        println!("set RHP_LIBRE_PROBE to one or more paths, separated by ';'");
        return;
    };

    // A name the app would not hand the engine can be forced into this run's list,
    // which is how the give-up itself is watched: the files that show it are the ones
    // the engine cannot draw at all, and no name it answers for does that. This is the
    // list as it is held in memory, so the file on disk is not touched.
    if let Ok(forced) = std::env::var("RHP_LIBRE_PROBE_FORCE") {
        if let Ok(mut config) = crate::CONFIG.lock() {
            for name in forced
                .split(',')
                .map(str::trim)
                .filter(|name| !name.is_empty())
            {
                let name = name.trim_start_matches('.').to_lowercase();
                if !config.libre_extensions.contains(&name) {
                    config.libre_extensions.push(name.clone());
                }
                println!("forced `{name}` into [libre] for this run");
            }
        }
    }

    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };
    let dpi = 96;
    let (cursor_x, cursor_y) = (900, 500);

    for path in list
        .split(';')
        .map(str::trim)
        .filter(|path| !path.is_empty())
    {
        let path = PathBuf::from(path);
        println!("\n--- {} ---", path.display());
        println!("engine available: {}", libreoffice_render::available());
        let text_lists = {
            let config = crate::CONFIG.lock().expect("the configuration");
            crate::formats::routing::named_as(&path, &config, PreviewType::Text)
        };
        println!(
            "kinds: video = {}, pdf = {}, office = {}, libre = {}, design = {}, vector = {}, text = {}",
            named_as(&path, PreviewType::Videos),
            pdf_preview::is_pdf_file(&path),
            named_as(&path, PreviewType::Document),
            named_as(&path, PreviewType::Libre),
            named_as(&path, PreviewType::Design),
            named_as(&path, PreviewType::Vector),
            text_lists,
        );

        let scale = effective_preview_scale(&path, current_hover_scales());
        println!("scale: {scale:?}");
        println!("render due: {}", libre_render_is_due(&path));
        println!(
            "measure before the page: {:?} — the spinner's own box is {}",
            media_dimensions(&path, bounds, dpi),
            office_preview::WAITING_BOX
        );

        // The hover the page is asked for, which is what the loop does the moment a
        // document like this is missed. A file this kind does not hold — one another
        // list reads, or one the engine has already turned down — has nothing to ask
        // for and nothing to wait on.
        let Some(_waiting) = request_libre_render(&path, 0) else {
            println!("request: nothing — the engine is not asked about this file");
            continue;
        };

        let started = Instant::now();
        let mut page = None;
        while page.is_none()
            && !libreoffice_render::refused(&path)
            && started.elapsed() < Duration::from_secs(90)
        {
            std::thread::sleep(Duration::from_millis(250));
            page = libreoffice_render::rendered_page(&path);
        }
        println!(
            "engine: page = {}, refused = {}, after {:?}",
            page.as_ref()
                .map(|page| page.display().to_string())
                .unwrap_or_else(|| "none".to_string()),
            libreoffice_render::refused(&path),
            started.elapsed()
        );

        if page.is_none() {
            continue;
        }

        // The replay, which is what the loop does when the page lands: the hover is
        // measured again — this time from the page — placed, and loaded.
        let Some(dimensions) = media_dimensions(&path, bounds, dpi) else {
            println!("measure after the page: nothing — the hover shows no preview");
            continue;
        };
        println!("measure after the page: {dimensions:?}");

        let Some(layout) = compute_mouse_layout(
            cursor_x,
            cursor_y,
            HoverPlacement {
                orig_dims: dimensions,
                avoid: None,
                follow_cursor: false,
                preview_scale: scale,
                at_the_pointer_corner: false,
            },
            bounds,
            dpi,
        ) else {
            println!("layout: none — the hover shows no preview");
            continue;
        };

        let cancel = Arc::new(AtomicBool::new(false));
        let started = Instant::now();
        match load_media(
            &path,
            layout.max_width,
            layout.max_height,
            scale,
            dpi,
            Arc::clone(&cancel),
        ) {
            Some(media) => println!(
                "loaded: {}x{}, {} frame(s), in {:?}",
                media.current_width(),
                media.current_height(),
                media.frames.len(),
                started.elapsed()
            ),
            None => println!(
                "loaded: nothing — the hover blinks, in {:?}",
                started.elapsed()
            ),
        }
    }
}

/// The whole path a hover takes for an archive an installed PeaZip lists, which is the
/// third path with a wait in the middle of it: measured as the wait for a listing, the
/// listing asked for, the wait watched, and then — when the engine has answered —
/// measured and loaded from the listing itself. Ignored, and driven by `RHP_PEAZIP_PROBE` —
/// `$env:RHP_PEAZIP_PROBE = "C:\downloads\backup.cab"; cargo test -- --ignored --nocapture peazip_hover_probe`
/// — for an archive whose preview does not appear.
///
/// It also answers the routing, which is what such a preview is usually about: the name the
/// configured list carries, what the file's bytes and its name together make of it, whether
/// the engine is installed to list it, and whether this hover is the wait for one at all.
/// That last question is asked here of the loader's own predicate rather than of the list,
/// because the list is only half of it: a file the engine will be asked about is still a
/// file with no preview if the load that missed it is not read as a wait (see
/// `awaiting_render` in the loader worker).
#[test]
#[ignore = "reads the files named in RHP_PEAZIP_PROBE and starts the installed PeaZip"]
fn peazip_hover_probe() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let Ok(list) = std::env::var("RHP_PEAZIP_PROBE") else {
        println!("set RHP_PEAZIP_PROBE to one or more paths, separated by ';'");
        return;
    };

    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };
    let dpi = 96;
    let (cursor_x, cursor_y) = (900, 500);

    for path in list
        .split(';')
        .map(str::trim)
        .filter(|path| !path.is_empty())
    {
        let path = PathBuf::from(path);
        println!("\n--- {} ---", path.display());

        let claimed = crate::CONFIG
            .lock()
            .map(|config| crate::formats::lists::PEAZIP.claims(&path, &config))
            .unwrap_or(false);

        println!(
            "kinds: peazip list = {claimed}, engine archive = {}, engine available = {}",
            peazip_formats::is_engine_archive(&path),
            peazip_render::available()
        );
        println!(
            "state: refused = {}, listed = {}, render due = {}",
            peazip_render::refused(&path),
            peazip_render::listed(&path),
            peazip_render_is_due(&path)
        );

        let scale = effective_preview_scale(&path, current_hover_scales());
        println!("scale: {scale:?}");
        println!(
            "measure before the listing: {:?} — the spinner's own box is {}",
            media_dimensions(&path, bounds, dpi),
            office_preview::WAITING_BOX
        );

        // The hover the listing is asked for, which is what the loop does the moment an
        // archive like this is missed. A file this kind does not hold — one another list
        // reads, one whose bytes are another kind, or one the engine has already turned
        // down — has nothing to ask for and nothing to wait on.
        let Some(_waiting) = request_peazip_render(&path, 0) else {
            println!("request: nothing — the engine is not asked about this file");
            continue;
        };

        let started = Instant::now();
        let mut listing = None;
        while listing.is_none()
            && !peazip_render::refused(&path)
            && started.elapsed() < Duration::from_secs(90)
        {
            std::thread::sleep(Duration::from_millis(250));
            listing = crate::readers::archive_listing::listing_for(&path, None);
        }

        println!(
            "engine: listed = {}, refused = {}, after {:?}",
            listing.is_some(),
            peazip_render::refused(&path),
            started.elapsed()
        );

        if let Some(listing) = &listing {
            println!(
                "listing: {} entries, {} bytes over them, packed {:?}, file {}",
                listing.entries.len(),
                listing.total_size,
                listing.packed_total,
                listing.file_size
            );
            for entry in listing.entries.iter().take(4) {
                println!(
                    "  {} ({} bytes, packed {:?}, dir {})",
                    entry.name, entry.size, entry.packed, entry.is_dir
                );
            }
        }

        // The replay, which is what the loop does when the listing lands: the hover is
        // measured again — this time from the listing — placed, and loaded.
        let Some(dimensions) = media_dimensions(&path, bounds, dpi) else {
            println!("measure after the listing: nothing — the hover shows no preview");
            continue;
        };
        println!("measure after the listing: {dimensions:?}");

        let Some(layout) = compute_mouse_layout(
            cursor_x,
            cursor_y,
            HoverPlacement {
                orig_dims: dimensions,
                avoid: None,
                follow_cursor: false,
                preview_scale: scale,
                at_the_pointer_corner: false,
            },
            bounds,
            dpi,
        ) else {
            println!("layout: none — the hover shows no preview");
            continue;
        };
        println!(
            "layout: {}x{} at ({}, {}), free room {}x{}",
            layout.preview_w,
            layout.preview_h,
            layout.pos_x,
            layout.pos_y,
            layout.max_width,
            layout.max_height
        );

        let cancel = Arc::new(AtomicBool::new(false));
        let started = Instant::now();
        match load_media(
            &path,
            layout.max_width,
            layout.max_height,
            scale,
            dpi,
            Arc::clone(&cancel),
        ) {
            Some(media) => println!(
                "loaded: {}x{}, {} frame(s), in {:?}",
                media.current_width(),
                media.current_height(),
                media.frames.len(),
                started.elapsed()
            ),
            None => println!(
                "loaded: nothing — the hover blinks, in {:?}",
                started.elapsed()
            ),
        }
    }
}

/// The whole path a hover takes for a comic, which is the book kind's fastest path: what the
/// file's name and its bytes together make of it, which plate the reader chooses out of the
/// container, the size it is placed at, and the frame it is drawn into. Ignored, and driven by
/// `RHP_COMIC_PROBE` —
/// `$env:RHP_COMIC_PROBE = "F:\manga\vol1.cbz;F:\manga\vol2.cbr"; cargo test -- --ignored --nocapture comic_hover_probe`
/// — for a comic whose preview does not appear, and for working through a folder of them one at
/// a time.
///
/// Nothing here waits on anything: a comic is read rather than converted, so the whole of what a
/// hover does is a listing and one plate — and the timings printed are what says so.
#[test]
#[ignore = "reads the files named in RHP_COMIC_PROBE"]
fn comic_hover_probe() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let Ok(list) = std::env::var("RHP_COMIC_PROBE") else {
        println!("set RHP_COMIC_PROBE to one or more paths, separated by ';'");
        return;
    };

    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };
    let dpi = 96;
    let (cursor_x, cursor_y) = (900, 500);

    for path in list
        .split(';')
        .map(str::trim)
        .filter(|path| !path.is_empty())
        .map(PathBuf::from)
    {
        println!("\n--- {} ---", path.display());

        let (claimed, page_name) = crate::CONFIG
            .lock()
            .map(|config| {
                (
                    crate::formats::lists::EBOOK.claims(&path, &config),
                    crate::formats::lists::EBOOK.claims(&path, &config),
                )
            })
            .unwrap_or((false, false));

        println!(
            "kinds: ebook list = {claimed}, page name = {page_name}, comic name = {}, preview = {}",
            claimed && !page_name,
            previewed_as(&path, PreviewType::Ebook)
        );

        let started = Instant::now();
        let Some(plate) = comic_preview::first_page_name(&path) else {
            println!(
                "plate: none — the container holds no page this app can draw, after {:?}",
                started.elapsed()
            );
            continue;
        };
        println!("plate: `{plate}` chosen in {:?}", started.elapsed());

        let started = Instant::now();
        let dimensions = comic_preview::dimensions(&path);
        println!("size: {dimensions:?} read in {:?}", started.elapsed());

        let scale = effective_preview_scale(&path, current_hover_scales());
        println!("scale: {scale:?}");

        let Some(measured) = media_dimensions(&path, bounds, dpi) else {
            println!("measure: nothing — the hover shows no preview");
            continue;
        };
        println!("measure: {measured:?}");

        let Some(layout) = compute_mouse_layout(
            cursor_x,
            cursor_y,
            HoverPlacement {
                orig_dims: measured,
                avoid: None,
                follow_cursor: false,
                preview_scale: scale,
                at_the_pointer_corner: false,
            },
            bounds,
            dpi,
        ) else {
            println!("layout: none — the hover shows no preview");
            continue;
        };
        println!(
            "layout: {}x{} at ({}, {}), free room {}x{}",
            layout.preview_w,
            layout.preview_h,
            layout.pos_x,
            layout.pos_y,
            layout.max_width,
            layout.max_height
        );

        let cancel = Arc::new(AtomicBool::new(false));
        let started = Instant::now();
        match load_media(
            &path,
            layout.max_width,
            layout.max_height,
            scale,
            dpi,
            Arc::clone(&cancel),
        ) {
            Some(media) => println!(
                "loaded: {}x{}, {} frame(s), in {:?}",
                media.current_width(),
                media.current_height(),
                media.frames.len(),
                started.elapsed()
            ),
            None => println!(
                "loaded: nothing — the hover blinks, in {:?}",
                started.elapsed()
            ),
        }
    }
}

/// The whole path a hover takes for a book an installed Calibre converts, which is the
/// render engine's path with a slower engine behind it: measured as the wait for a page, the
/// conversion asked for, the page watched for, and then — when the engine has answered —
/// measured and loaded from the page itself. Ignored, and driven by `RHP_CALIBRE_PROBE` —
/// `$env:RHP_CALIBRE_PROBE = "C:\books\book.mobi;C:\books\book.epub"; cargo test -- --ignored --nocapture calibre_hover_probe`
/// — for a book whose preview does not appear, and for working through a list of the formats
/// this engine is asked about one at a time.
///
/// It also answers the routing, which is what such a preview is usually about: the name the
/// configured list carries, what the file's bytes and its name together make of it, whether
/// the engine is installed to convert it, and whether this hover is the wait for one at all.
/// That last question is asked here of the loader's own predicate rather than of the list,
/// because the list is only half of it: a file the engine will be asked about is still a file
/// with no preview if the load that missed it is not read as a wait (see `awaiting_render` in
/// the loader worker).
#[test]
#[ignore = "reads the files named in RHP_CALIBRE_PROBE and starts the installed Calibre"]
fn calibre_hover_probe() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let Ok(list) = std::env::var("RHP_CALIBRE_PROBE") else {
        println!("set RHP_CALIBRE_PROBE to one or more paths, separated by ';'");
        return;
    };

    let bounds = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };
    let dpi = 96;
    let (cursor_x, cursor_y) = (900, 500);

    // The threads a preview really runs on are in an apartment before they ask for a page
    // (see `spawn_load_worker` and `run_preview_window`), and a probe asks for one from a test
    // thread that is not: a page's size is read through `Windows.Data.Pdf`, so the probe has to
    // put itself in one first, exactly as the two threads do.
    pdf_preview::initialize_apartment();
    wic_image::initialize_apartment();

    for path in list
        .split(';')
        .map(str::trim)
        .filter(|path| !path.is_empty())
    {
        let path = PathBuf::from(path);
        println!("\n--- {} ---", path.display());

        let claimed = crate::CONFIG
            .lock()
            .map(|config| crate::formats::lists::CALIBRE.claims(&path, &config))
            .unwrap_or(false);

        println!(
            "kinds: calibre list = {claimed}, engine ebook = {}, engine available = {}",
            calibre_formats::is_engine_ebook(&path),
            calibre_render::available()
        );
        println!(
            "state: refused = {}, page = {:?}, render due = {}",
            calibre_render::refused(&path),
            calibre_render::rendered_page(&path),
            calibre_render_is_due(&path)
        );

        let scale = effective_preview_scale(&path, current_hover_scales());
        println!("scale: {scale:?}");
        println!(
            "measure before the page: {:?} — the spinner's own box is {}",
            media_dimensions(&path, bounds, dpi),
            office_preview::WAITING_BOX
        );

        // The hover the page is asked for, which is what the loop does the moment a book like
        // this is missed. A file this kind does not hold — one another list reads, one whose
        // bytes are another kind, or one the engine has already turned down — has nothing to
        // ask for and nothing to wait on.
        let Some(_waiting) = request_calibre_render(&path, 0) else {
            println!("request: nothing — the engine is not asked about this file");
            continue;
        };

        let started = Instant::now();
        let mut page = None;
        while page.is_none()
            && !calibre_render::refused(&path)
            && started.elapsed() < Duration::from_secs(300)
        {
            std::thread::sleep(Duration::from_millis(250));
            page = calibre_render::rendered_page(&path);
        }

        println!(
            "engine: page = {:?}, refused = {}, after {:?}",
            page,
            calibre_render::refused(&path),
            started.elapsed()
        );

        if let Some(page) = &page {
            println!(
                "page: {} bytes, {:?} — previewed from {:?}",
                std::fs::metadata(page).map(|meta| meta.len()).unwrap_or(0),
                pdf_preview::page_dimensions(page),
                pdf_preview::book_page(page)
            );
        }

        // The replay, which is what the loop does when the page lands: the hover is measured
        // again — this time from the page — placed, and loaded.
        let Some(dimensions) = media_dimensions(&path, bounds, dpi) else {
            println!("measure after the page: nothing — the hover shows no preview");
            continue;
        };
        println!("measure after the page: {dimensions:?}");

        let Some(layout) = compute_mouse_layout(
            cursor_x,
            cursor_y,
            HoverPlacement {
                orig_dims: dimensions,
                avoid: None,
                follow_cursor: false,
                preview_scale: scale,
                at_the_pointer_corner: false,
            },
            bounds,
            dpi,
        ) else {
            println!("layout: none — the hover shows no preview");
            continue;
        };
        println!(
            "layout: {}x{} at ({}, {}), free room {}x{}",
            layout.preview_w,
            layout.preview_h,
            layout.pos_x,
            layout.pos_y,
            layout.max_width,
            layout.max_height
        );

        let cancel = Arc::new(AtomicBool::new(false));
        let started = Instant::now();
        match load_media(
            &path,
            layout.max_width,
            layout.max_height,
            scale,
            dpi,
            Arc::clone(&cancel),
        ) {
            Some(media) => println!(
                "loaded: {}x{}, {} frame(s), in {:?}",
                media.current_width(),
                media.current_height(),
                media.frames.len(),
                started.elapsed()
            ),
            None => println!(
                "loaded: nothing — the hover blinks, in {:?}",
                started.elapsed()
            ),
        }
    }
}

#[test]
#[ignore = "builds an animation larger than the retained window of its own"]
fn sliding_window_probe() {
    use image::codecs::gif::{GifEncoder, Repeat};
    use image::{Delay, Frame, Rgba, RgbaImage};
    use std::io::BufWriter;

    let path = std::env::temp_dir().join("rhp-sliding-window-probe.gif");
    let (side, frame_count, delay_ms) = (1024u32, 40u32, 40u32);

    {
        let file = File::create(&path).expect("a file to write the probe animation to");
        let mut encoder = GifEncoder::new_with_speed(BufWriter::new(file), 30);
        encoder.set_repeat(Repeat::Infinite).expect("looping");
        for index in 0..frame_count {
            let shade = (index * 251 / frame_count) as u8;
            let image = RgbaImage::from_pixel(side, side, Rgba([shade, 40, 255 - shade, 255]));
            encoder
                .encode_frame(Frame::from_parts(
                    image,
                    0,
                    0,
                    Delay::from_numer_denom_ms(delay_ms, 1),
                ))
                .expect("a frame");
        }
    }

    let decoded_mb = u64::from(side) * u64::from(side) * 4 * u64::from(frame_count) / 1048576;
    println!(
        "\n--- {} ({} frames, {} MB decoded) ---",
        path.display(),
        frame_count,
        decoded_mb
    );

    let cancel = Arc::new(AtomicBool::new(false));
    let Some(mut media) = load_animated_gif(
        &path,
        1920,
        1080,
        PreviewScale::Percent(100),
        Arc::clone(&cancel),
    ) else {
        println!("load: None");
        return;
    };

    let start = Instant::now();
    let mut last_frame = media.current_frame;
    let mut advances = 0usize;
    let mut wraps = 0usize;
    let mut next_log = Instant::now();

    while start.elapsed() < Duration::from_secs(12) {
        if media.advance_frame() {
            advances += 1;
            if media.current_frame <= last_frame {
                wraps += 1;
            }
            last_frame = media.current_frame;
        }

        if Instant::now() >= next_log {
            next_log = Instant::now() + Duration::from_millis(1000);
            println!(
                "  t={:>5.2}s frame={:<4} held={:<4} loaded={:<5} released={:<5} advances={} wraps={}",
                start.elapsed().as_secs_f32(),
                media.current_frame,
                media.frames.len(),
                media.is_fully_loaded(),
                media.frames_were_released(),
                advances,
                wraps
            );
        }

        std::thread::sleep(Duration::from_millis(4));
    }

    println!(
        "after 12s: advances={} wraps={} frame={} held={}",
        advances,
        wraps,
        media.current_frame,
        media.frames.len()
    );
    media.cancel_background_work();
    let _ = std::fs::remove_file(&path);
}

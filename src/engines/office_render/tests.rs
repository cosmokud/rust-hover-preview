use super::*;

/// A request rendered where no worker is running, which is where these tests are.
///
/// The generation is the one the counter holds, so the check that a worker which
/// has been given up on starts nothing is satisfied: what these tests are asking
/// about is a render, not the worker that would normally make it.
fn render_here(engines: &mut Engines, request: &RenderRequest) -> RenderOutcome {
    render_request(engines, request, WORKER_GENERATION.load(Ordering::Acquire))
}

/// A document of this module's own, in a folder the tests share: a page is kept by the
/// document it was drawn for, so the names are what keep one test's document out of
/// another's.
fn document(name: &str) -> PathBuf {
    let folder = std::env::temp_dir()
        .join("rust-hover-preview-office-tests")
        .join("documents");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let path = folder.join(name);
    std::fs::write(&path, b"a document").expect("a written document");

    path
}

/// A slide is exported at the width the render was asked for, so a deck that was
/// first previewed on a smaller display holds a page a larger one is right to ask
/// for again — and one exported at the cap is not, which is what keeps asking for
/// it from being a loop. A workbook's picture is the other raster Office draws, and
/// it is nobody's export: it is the used range at the range's own size, so a larger
/// display is not a reason to ask for one.
#[test]
fn asks_for_a_slide_again_when_a_wider_display_wants_one() {
    let source = document("wider.pptx");
    let deck = document_cache::store(
        &source,
        OfficeEngine::MicrosoftOffice.as_str(),
        PageKind::Png,
        &picture_bytes(1280, 720),
    )
    .expect("a kept slide");

    assert!(
        page_is_narrower_than(&source, &deck, 1920),
        "a wider box replaces the page"
    );
    assert!(
        page_is_narrower_than(&source, &deck, 3840),
        "and a display larger than a slide is ever exported for asks for the widest one"
    );
    assert!(
        !page_is_narrower_than(&source, &deck, 1200),
        "a box the page already fills does not"
    );

    let at_the_cap = document_cache::store(
        &source,
        OfficeEngine::MicrosoftOffice.as_str(),
        PageKind::Png,
        &picture_bytes(MAX_SLIDE_EXPORT_WIDTH, 1080),
    )
    .expect("a kept slide");

    assert!(
        !page_is_narrower_than(&source, &at_the_cap, 3840),
        "a page already at the cap is not asked for again"
    );

    // A workbook's picture is the other PNG a family draws, and it has no export width to
    // ask again for: what it is is the used range at the size the range came out at,
    // whatever box the render was asked for — so a second render would write the same
    // picture, and asking for one would be a render paid on every hover.
    let workbook = document("wider.xlsx");
    let picture = document_cache::store(
        &workbook,
        OfficeEngine::MicrosoftOffice.as_str(),
        PageKind::Png,
        &picture_bytes(800, 600),
    )
    .expect("a kept picture");

    assert!(
        !page_is_narrower_than(&workbook, &picture, 3840),
        "a workbook's picture is not asked for again for a larger display"
    );

    let _ = std::fs::remove_file(&workbook);
    let _ = std::fs::remove_file(&source);
}

/// A PNG of one colour, written the way a slide's export and a workbook's picture are.
fn picture_bytes(width: u32, height: u32) -> Vec<u8> {
    let mut written = std::io::Cursor::new(Vec::new());
    image::RgbaImage::from_pixel(width, height, image::Rgba([10, 20, 30, 255]))
        .write_to(&mut written, image::ImageFormat::Png)
        .expect("a written slide");

    written.into_inner()
}

/// The engine's own lifecycle: one is started for each family, let go, and
/// nothing is left behind.
#[test]
#[ignore = "starts the installed Office"]
fn engine_lifecycle_probe() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

    for app_kind in [OfficeApp::Word, OfficeApp::Excel, OfficeApp::PowerPoint] {
        println!("\n--- {app_kind:?} ---");
        println!(
            "before: {:?}",
            engine_processes::processes_named(app_kind.image_name())
        );

        let Some(engine) = Engine::create(app_kind) else {
            println!("created: no");
            continue;
        };
        println!(
            "created: attached={} owned_pid={}",
            engine.attached, engine.owned_pid
        );

        drop(engine);
        std::thread::sleep(Duration::from_secs(3));
        println!(
            "after: {:?}",
            engine_processes::processes_named(app_kind.image_name())
        );
    }
}

/// A page that cannot be read is not a page: the side that draws a preview drops
/// it, so the document is rendered again rather than answered with a preview that
/// blinks away every time it is hovered.
#[test]
fn forgets_a_page_it_cannot_read() {
    use crate::readers::office_preview::{measure, source_kind, SourceKind};

    // The test is about a page this tier produced, and what a hover asks of it: what the
    // app's own `config.ini` holds is not what it is about, so the engine choice is the
    // one that asks the tier (see `office_formats::page_engine`).
    if let Ok(mut config) = crate::CONFIG.lock() {
        config.office_engine = crate::config::config::OfficeEngine::MicrosoftOffice;
    }

    let source = document("held.xlsx");

    // A picture that can be read is a page, and its own size is what the layout
    // places the preview by — and it is a picture rather than a page in the layout's
    // eyes, since what a workbook is answered with is only as good as its pixels.
    document_cache::store(
        &source,
        OfficeEngine::MicrosoftOffice.as_str(),
        PageKind::Png,
        &picture_bytes(2, 2),
    )
    .expect("a kept page");
    assert_eq!(measure(&source), Some((2, 2)), "a page that can be read");
    assert_eq!(source_kind(&source), SourceKind::Raster);

    // One that cannot is dropped, and the answer is that nothing is rendered yet.
    document_cache::store(
        &source,
        OfficeEngine::MicrosoftOffice.as_str(),
        PageKind::Png,
        b"not a picture at all",
    )
    .expect("a kept page");
    assert_eq!(source_kind(&source), SourceKind::None, "nothing to draw");
    assert!(held_page(&source).is_none(), "the broken page is gone");

    let _ = std::fs::remove_file(&source);
}

/// A file that was downloaded carries a zone identifier, which is what puts
/// Word and Excel into Protected View — where a page cannot be exported — so it
/// is rendered from a copy made without it. A file that was never downloaded
/// must not be copied: a large document is expensive to copy for nothing.
#[test]
fn sees_a_zone_identifier() {
    // A folder of this module's own: the tests run beside each other, and one
    // of them clearing its fixtures must not take another's with it.
    let folder = std::env::temp_dir()
        .join("rust-hover-preview-office-tests")
        .join("zones");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let path = folder.join("marked.docx");
    std::fs::write(&path, b"a document").expect("a written document");
    let plain = path.to_string_lossy().to_string();
    assert!(!has_zone_identifier(&plain), "nothing marks it yet");

    // The stream Windows writes beside a downloaded file, written the way any
    // process could write it.
    std::fs::write(
        format!("{plain}:Zone.Identifier"),
        b"[ZoneTransfer]\r\nZoneId=3\r\n",
    )
    .expect("a written stream");
    assert!(has_zone_identifier(&plain), "the stream is seen");

    let _ = std::fs::remove_file(&path);
}

/// Renders documents that are already on disk, named in `RHP_OFFICE_PROBE`
/// (separated by `;`), through the real code path and reports what happened.
///
/// Ignored like the smoke test, and the way to look at a document that will not
/// preview:
/// `$env:RHP_OFFICE_PROBE = "C:\docs\one.xlsx;C:\docs\two.xls"`
/// `cargo test -- --ignored --nocapture office_render_probe`
#[test]
#[ignore = "starts the installed Office"]
fn office_render_probe() {
    // What the app's own `config.ini` holds is not what this probe is about: it measures
    // this tier, so the setting that asks the tier is the one it runs under.
    if let Ok(mut config) = crate::CONFIG.lock() {
        config.office_engine = crate::config::config::OfficeEngine::MicrosoftOffice;
    }

    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

    let Ok(list) = std::env::var("RHP_OFFICE_PROBE") else {
        println!("set RHP_OFFICE_PROBE to one or more paths, separated by ';'");
        return;
    };

    for path in list
        .split(';')
        .map(str::trim)
        .filter(|path| !path.is_empty())
    {
        let path = PathBuf::from(path);
        println!("\n--- {} ---", path.display());
        println!(
            "exists: {} container: {:?} engine: {:?}",
            path.exists(),
            container_kind(&path),
            app_for(&path)
        );

        let mut engines = Engines::new();
        let request = RenderRequest {
            source: path.clone(),
            width: 1280,
            height: 800,
            generation: 1,
            requested: Instant::now(),
        };

        let started = Instant::now();
        let outcome = render_here(&mut engines, &request);
        println!(
            "rendered: {} in {:?}",
            outcome == RenderOutcome::Rendered,
            started.elapsed()
        );
        match app_for(&path).and_then(|app_kind| engines.get(app_kind)) {
            Some(engine) => println!(
                "engine: attached={} owned_pid={}",
                engine.attached, engine.owned_pid
            ),
            None => println!("engine: none"),
        }
        println!(
            "last failure: {}",
            last_failure().unwrap_or_else(|| "none recorded".to_string())
        );
        match held_page(&path) {
            Some(page) => println!(
                "page: {:?} ({} bytes)",
                page.kind,
                std::fs::metadata(&page.path)
                    .map(|metadata| metadata.len())
                    .unwrap_or(0)
            ),
            None => println!("page: none"),
        }

        // The half a hover does after the render: measure it, then draw it —
        // on a thread of its own, in a multithreaded apartment, which is where
        // the app does both.
        let drawing = path.clone();
        let (measured, drawn) = std::thread::spawn(move || {
            crate::readers::pdf_preview::initialize_apartment();
            let measured = crate::readers::office_preview::measure(&drawing);
            let drawn = crate::readers::office_preview::render(&drawing, 1200, 900, None);
            (measured, drawn)
        })
        .join()
        .expect("the drawing thread");

        println!("measured: {measured:?}");
        match drawn {
            Some((pixels, width, height)) => {
                println!("drawn: {width}x{height} ({} pixels)", pixels.len() / 4)
            }
            None => println!("drawn: nothing"),
        }

        engines.drop_all();
    }
}

/// A document of the given family, written by the application itself, with a
/// little content in it so that a page has something to show.
fn write_sample(app_kind: OfficeApp, path: &Path) -> Option<()> {
    let app = Object::create(app_kind.prog_id())?;
    // The hygiene the engine applies, which a fixture that leaves it out hangs
    // on: an alert is a dialog no one is present to answer.
    let _ = app.set("DisplayAlerts", alerts_off(app_kind));

    let collection = app.member(match app_kind {
        OfficeApp::Word => "Documents",
        OfficeApp::Excel => "Workbooks",
        OfficeApp::PowerPoint => "Presentations",
    })?;
    let document = Object::from_variant(collection.call("Add", &[])?)?;

    match app_kind {
        OfficeApp::Word => {}
        OfficeApp::Excel => {
            if let Some(sheet) = document
                .member("Worksheets")
                .and_then(|sheets| sheets.item(1))
            {
                for (cell, value) in [
                    ("A1", "Item"),
                    ("B1", "Qty"),
                    ("A2", "Widget"),
                    ("B2", "3"),
                    ("A3", "Gadget"),
                    ("B3", "7"),
                ] {
                    if let Some(range) = sheet
                        .call_args("Range", &[VARIANT::from(cell)])
                        .and_then(Object::from_variant)
                    {
                        let _ = range.set("Value2", VARIANT::from(value));
                    }
                }
            }
        }
        OfficeApp::PowerPoint => {
            if let Some(slides) = document.member("Slides") {
                let _ = slides.call_args(
                    "Add",
                    &[VARIANT::from(1i32), VARIANT::from(1i32)], // index, layout
                );
            }
        }
    }

    let saved = match app_kind {
        OfficeApp::Word => document.call(
            "SaveAs2",
            &[
                ("FileName", path_variant(path)),
                ("FileFormat", VARIANT::from(12i32)), // wdFormatXMLDocument
            ],
        ),
        OfficeApp::Excel => document.call(
            "SaveAs",
            &[
                ("Filename", path_variant(path)),
                ("FileFormat", VARIANT::from(51i32)), // xlOpenXMLWorkbook
            ],
        ),
        OfficeApp::PowerPoint => document.call(
            "SaveAs",
            &[
                ("FileName", path_variant(path)),
                ("Format", VARIANT::from(24i32)), // ppSaveAsOpenXMLPresentation
            ],
        ),
    };

    let _ = document.call("Close", &[("SaveChanges", VARIANT::from(false))]);
    let _ = app.call("Quit", &[]);

    saved.map(|_| ())
}

/// One document of each family, written and then rendered through the real
/// code path, reporting what every step did.
///
/// Ignored because it starts the installed Office and writes sample documents
/// into the scratchpad. Run it when a document produces no page:
/// `cargo test -- --ignored --nocapture office_render_smoke_test`.
#[test]
#[ignore = "starts the installed Office and writes sample documents"]
fn office_render_smoke_test() {
    // What the app's own `config.ini` holds is not what this probe is about: it measures
    // this tier, so the setting that asks the tier is the one it runs under.
    if let Ok(mut config) = crate::CONFIG.lock() {
        config.office_engine = crate::config::config::OfficeEngine::MicrosoftOffice;
    }

    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

    let folder = std::env::var_os("COMMANDCODE_SCRATCHPAD")
        .map(PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("office-smoke-test");
    std::fs::create_dir_all(&folder).expect("a scratch folder");

    for (app_kind, name) in [
        (OfficeApp::Word, "sample.docx"),
        (OfficeApp::Excel, "sample.xlsx"),
        (OfficeApp::PowerPoint, "sample.pptx"),
    ] {
        println!("\n--- {name} ---");
        let path = folder.join(name);
        let _ = std::fs::remove_file(&path);

        let started = Instant::now();
        match write_sample(app_kind, &path) {
            Some(()) => println!("written: {} in {:?}", path.display(), started.elapsed()),
            None => {
                println!(
                    "written: no ({})",
                    last_failure().unwrap_or_else(|| "no failure recorded".to_string())
                );
                continue;
            }
        }

        // What the hover would show before anything is rendered: nothing, so the
        // spinner's own box is what the layout places, and the page is asked for
        // in the box that family's pages have.
        println!(
            "before a render: source {:?}, spinner box {}, render box {:?}",
            crate::readers::office_preview::source_kind(&path),
            crate::readers::office_preview::WAITING_BOX,
            crate::formats::office_formats::default_page_size(&path)
        );

        let mut engines = Engines::new();
        let request = RenderRequest {
            source: path.clone(),
            width: 1280,
            height: 800,
            generation: 1,
            requested: Instant::now(),
        };
        let started = Instant::now();
        let outcome = render_here(&mut engines, &request);
        println!(
            "rendered: {} in {:?}",
            outcome == RenderOutcome::Rendered,
            started.elapsed()
        );

        // What the engine made of the instance: an attached one is the user's
        // and is never hidden, quit or ended; one this app started is.
        match engines.get(app_kind) {
            Some(engine) => println!(
                "engine: attached={} owned_pid={}",
                engine.attached, engine.owned_pid
            ),
            None => println!("engine: none"),
        }

        if outcome != RenderOutcome::Rendered {
            println!(
                "last failure: {}",
                last_failure().unwrap_or_else(|| "none recorded".to_string())
            );
        }
        match held_page(&path) {
            Some(page) => println!(
                "held page: {:?} ({} bytes)",
                page.kind,
                std::fs::metadata(&page.path)
                    .map(|metadata| metadata.len())
                    .unwrap_or(0)
            ),
            None => println!("held page: none"),
        }

        // Drawing happens on a thread of its own, in a multithreaded
        // apartment, which is where the app draws too: this thread is
        // apartment-threaded for the automation above, and a WinRT call waited
        // on from a single-threaded apartment deadlocks without a pump.
        let drawing = path.clone();
        let started = Instant::now();
        let drawn = std::thread::spawn(move || {
            crate::readers::pdf_preview::initialize_apartment();
            crate::readers::office_preview::render(&drawing, 800, 600, None)
        })
        .join()
        .ok()
        .flatten();

        match drawn {
            Some((pixels, width, height)) => println!(
                "drawn: {width}x{height}, {} pixels, in {:?}",
                pixels.len() / 4,
                started.elapsed()
            ),
            None => println!("drawn: nothing, in {:?}", started.elapsed()),
        }

        engines.drop_all();
    }
}

/// The engines the tier holds, one per family, side by side.
///
/// A folder can hold a document, a workbook and a deck, and one engine for the
/// whole tier meant the pointer crossing between them quit one application and
/// started another every time it did. Each family keeps its own now, so what
/// the second pass below times is an engine that is already there rather than a
/// cold start, and what it counts is three engines held at once.
///
/// Ignored because it starts the installed Office and writes sample documents.
/// `cargo test -- --ignored --nocapture office_engines_side_by_side`
#[test]
#[ignore = "starts the installed Office and writes sample documents"]
fn office_engines_side_by_side() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

    let folder = std::env::var_os("COMMANDCODE_SCRATCHPAD")
        .map(PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("office-engines");
    std::fs::create_dir_all(&folder).expect("a scratch folder");

    // A sample per family and a copy of it: the copy is another path, so the
    // page cache cannot answer for it and the second pass really renders.
    let mut samples: Vec<(&str, OfficeApp, PathBuf, PathBuf)> = Vec::new();
    for (app_kind, name) in [
        (OfficeApp::Word, "sample.docx"),
        (OfficeApp::Excel, "sample.xlsx"),
        (OfficeApp::PowerPoint, "sample.pptx"),
    ] {
        let path = folder.join(name);
        let again = folder.join(format!("again-{name}"));
        let _ = std::fs::remove_file(&path);
        let _ = std::fs::remove_file(&again);

        if write_sample(app_kind, &path).is_none() {
            println!("{name}: not written, skipped");
            continue;
        }
        let _ = std::fs::copy(&path, &again);
        samples.push((name, app_kind, path, again));
    }

    // One holder across all three families, which is what the tier keeps.
    let mut engines = Engines::new();

    let render_one = |engines: &mut Engines, path: &Path| {
        let request = RenderRequest {
            source: path.to_path_buf(),
            width: 1280,
            height: 800,
            generation: 1,
            requested: Instant::now(),
        };
        let started = Instant::now();
        let outcome = render_here(engines, &request);
        (outcome, started.elapsed())
    };

    println!("\n--- cold, one family after another ---");
    for (name, _, path, _) in &samples {
        let (outcome, took) = render_one(&mut engines, path);
        println!(
            "{name}: rendered {} in {took:?}, engines held {}",
            outcome == RenderOutcome::Rendered,
            engines.iter().count()
        );
    }

    println!("\n--- warm, a second document of each family ---");
    for (name, _, _, again) in &samples {
        let (outcome, took) = render_one(&mut engines, again);
        println!(
            "{name}: rendered {} in {took:?}, engines held {}",
            outcome == RenderOutcome::Rendered,
            engines.iter().count()
        );
    }

    // What the tier is for: the family rendered first is still held after the
    // other two have been asked for.
    for (name, app_kind, _, _) in &samples {
        assert!(
            engines.get(*app_kind).is_some(),
            "{name} left no engine behind"
        );
    }

    engines.drop_all();
}

/// Renders several documents one after another on a single warm engine, which
/// is what the tier does across hovers, and reports what each render cost.
///
/// The list is `RHP_WARM_PROBE`, separated by `;`. Every path must be distinct
/// — a page that is already held answers from the cache and never reaches the
/// engine — so the same workbook is listed once under two names.
///
/// Ignored because it starts the installed Office.
/// `cargo test -- --ignored --nocapture excel_warm_render_probe`
#[test]
#[ignore = "starts the installed Office"]
fn excel_warm_render_probe() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

    let Ok(list) = std::env::var("RHP_WARM_PROBE") else {
        println!("set RHP_WARM_PROBE to one or more paths, separated by ';'");
        return;
    };
    let paths: Vec<PathBuf> = list
        .split(';')
        .map(str::trim)
        .filter(|path| !path.is_empty())
        .map(PathBuf::from)
        .collect();

    let mut engines = Engines::new();
    for path in &paths {
        let request = RenderRequest {
            source: path.clone(),
            width: 1280,
            height: 800,
            generation: 1,
            requested: Instant::now(),
        };

        let engine_before = app_for(path)
            .and_then(|app_kind| engines.get(app_kind))
            .map(|engine| engine.owned_pid);
        let started = Instant::now();
        let outcome = render_here(&mut engines, &request);
        let took = started.elapsed();
        let kind = held_page(path).map(|page| page.kind);
        let engine_after = app_for(path)
            .and_then(|app_kind| engines.get(app_kind))
            .map(|engine| engine.owned_pid);

        println!(
            "{}: rendered {} in {took:?}, page {:?}, engine {:?} -> {:?}, last failure: {}",
            path.file_name().unwrap_or_default().to_string_lossy(),
            outcome == RenderOutcome::Rendered,
            kind,
            engine_before,
            engine_after,
            last_failure().unwrap_or_else(|| "none".to_string())
        );
    }

    engines.drop_all();
}

/// `xlCalculationManual`: the setting that says a workbook is not to be worked
/// out again, whose reach the probe below measures.
const XL_CALCULATION_MANUAL: i32 = -4135;

/// Whether a workbook really is opened without being recalculated.
///
/// The workbook this measures is saved *stale*: one cell is changed while
/// calculation is manual, so the cell beside it keeps the value it had rather
/// than being worked out again. What is on disk is then a workbook whose stored
/// value and whose computed value differ — which is the only thing that makes a
/// recalculation on open visible. A formula of the time would say nothing: a
/// volatile one is recalculated whenever a workbook is opened, whatever the
/// calculation mode is.
///
/// Ignored because it starts the installed Excel and writes a workbook.
/// `cargo test -- --ignored --nocapture excel_calculation_probe`
#[test]
#[ignore = "starts the installed Excel and writes a workbook"]
fn excel_calculation_probe() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

    let folder = std::env::var_os("COMMANDCODE_SCRATCHPAD")
        .map(PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("office-calculation");
    std::fs::create_dir_all(&folder).expect("a scratch folder");
    let path = folder.join("stale.xlsx");
    let _ = std::fs::remove_file(&path);

    {
        let Some(app) = Object::create(OfficeApp::Excel.prog_id()) else {
            println!("no Excel");
            return;
        };
        let _ = app.set("DisplayAlerts", alerts_off(OfficeApp::Excel));
        with_a_workbook(&app, || {
            let _ = app.set("Calculation", VARIANT::from(XL_CALCULATION_MANUAL));
            // Saving recalculates a workbook unless it is told not to, which would
            // leave the file holding the worked-out value rather than the stale one
            // this needs it to hold.
            let _ = app.set("CalculateBeforeSave", VARIANT::from(false));
        });

        let Some(book) = app
            .member("Workbooks")
            .and_then(|books| books.call("Add", &[]))
            .and_then(Object::from_variant)
        else {
            println!("no workbook");
            return;
        };
        let Some(sheet) = book.member("Worksheets").and_then(|sheets| sheets.item(1)) else {
            println!("no worksheet");
            return;
        };

        set_cell(&sheet, "A1", VARIANT::from(1i32));
        set_cell(&sheet, "B1", VARIANT::from("=A1+1"));
        println!("B1 once entered:        {:?}", read_cell(&sheet, "B1"));

        // The change manual calculation is supposed to hide.
        set_cell(&sheet, "A1", VARIANT::from(5i32));
        println!("B1 after A1 became 5:   {:?}", read_cell(&sheet, "B1"));

        let _ = book.call("SaveAs", &[("Filename", path_variant(&path))]);
        println!("B1 as it was saved:     {:?}", read_cell(&sheet, "B1"));

        let _ = book.call("Close", &[("SaveChanges", VARIANT::from(false))]);
        let _ = app.call("Quit", &[]);
    }

    println!(
        "opened with calculation manual:  B1 = {:?}",
        read_saved(&path, "B1", true)
    );
    println!(
        "opened with Excel's own setting: B1 = {:?}",
        read_saved(&path, "B1", false)
    );
    println!("(a recalculation is the 6; the saved value is the 2)");
}

/// One cell of a worksheet, by address.
fn read_cell(sheet: &Object, address: &str) -> Option<String> {
    sheet
        .call_args("Range", &[VARIANT::from(address)])
        .and_then(Object::from_variant)
        .and_then(|cell| cell.value("Value2"))
        .map(|value| value.to_string())
}

fn set_cell(sheet: &Object, address: &str, value: VARIANT) {
    if let Some(cell) = sheet
        .call_args("Range", &[VARIANT::from(address)])
        .and_then(Object::from_variant)
    {
        let _ = cell.set("Value2", value);
    }
}

/// One cell of a workbook on disk, opened with — or without — the setting the
/// engine applies first.
fn read_saved(path: &Path, address: &str, manual: bool) -> Option<String> {
    let app = Object::create(OfficeApp::Excel.prog_id())?;
    let _ = app.set("DisplayAlerts", alerts_off(OfficeApp::Excel));

    if manual {
        // The engine's own step, through the same helper it uses: Excel answers
        // for its calculation mode only while a workbook is open.
        with_a_workbook(&app, || {
            let _ = app.set("Calculation", VARIANT::from(XL_CALCULATION_MANUAL));
        });
    }

    let book = app
        .member("Workbooks")
        .and_then(|books| {
            books.call(
                "Open",
                &[
                    ("FileName", path_variant(path)),
                    ("ReadOnly", VARIANT::from(true)),
                ],
            )
        })
        .and_then(Object::from_variant)?;

    let value = book
        .member("Worksheets")
        .and_then(|sheets| sheets.item(1))
        .and_then(|sheet| read_cell(&sheet, address));

    let _ = book.call("Close", &[("SaveChanges", VARIANT::from(false))]);
    let _ = app.call("Quit", &[]);
    value
}

use super::*;

/// The page the engine is pointed at is what makes a document the size of the window
/// it was given: the document goes in an image, and an image is scaled to its box
/// whatever size it asks for. What a document may not do — run code, leave the page —
/// is what an image may not do either, which is why it is drawn this way rather than
/// being put in the page as markup.
///
/// A document whose name needs escaping is the case the URL builder and the attribute
/// have to agree on: an `&` written as itself ends the attribute early and the image
/// is a document the browser never found.
#[test]
fn the_page_draws_the_document_as_an_image_that_fills_it() {
    let folder = std::env::temp_dir().join("rust-hover-preview-frame-tests");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let document = folder.join("a document & one.svg");
    std::fs::write(
        &document,
        br#"<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"></svg>"#,
    )
    .expect("a written file");

    let (page, url) =
        frame_page(&document, 42, TransparentBackground::Black).expect("a page for a document");
    let html = std::fs::read_to_string(&page).expect("a written page");

    assert!(html.contains("width:100%;height:100%;object-fit:contain"));
    assert!(html.contains("file:///"));
    assert!(html.contains("a%20document%20&amp;%20one.svg"));
    assert!(html.contains("?v=42"));
    assert!(
        !html.contains("conic-gradient"),
        "a backdrop the controller can be given is the controller's"
    );
    assert!(url.ends_with("?v=42"));
    assert!(url.starts_with("file:///"));

    // A checkerboard is the one backdrop it cannot be given, because it is drawn by
    // whatever composites the frame and the page is what composites this one.
    let (page, _) = frame_page(&document, 42, TransparentBackground::Checkerboard)
        .expect("a page for a document");
    let html = std::fs::read_to_string(&page).expect("a written page");

    assert!(html.contains("conic-gradient"));
    assert!(html.contains("#e0e0e0"));
    assert!(html.contains("background-size:32px 32px"));
}

/// A specimen's page is the font itself in the page, the lines the file's own character
/// map covers, and a heading the font's `name` table supplies — the three things the page
/// carries rather than asks the engine for. Two of the four backdrops are the page's too:
/// the checkerboard's squares, and the ink a specimen is drawn in, which is dark on the
/// light backdrops, light on black, and light with a shadow where there is no backdrop at
/// all to be read against.
#[test]
fn the_page_draws_the_font_with_its_own_lines_and_a_name() {
    let folder = std::env::temp_dir().join("rust-hover-preview-font-page-tests");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let font = folder.join("specimen.ttf");
    std::fs::write(&font, specimen_font()).expect("a written font");

    let (page, url) =
        font_page(&font, 7, TransparentBackground::Black, 0).expect("a page for a font");
    let html = std::fs::read_to_string(&page).expect("a written page");

    assert!(html.contains("@font-face"));
    assert!(html.contains("font-family:\"RHPPreviewFont\""));
    assert!(
        html.contains("specimen.ttf?v=7"),
        "the font is in the page as the file it is, at the version that was read"
    );
    assert!(html.contains("Test Family Regular"));
    assert!(html.contains("The quick brown fox jumps over the lazy dog."));
    assert!(
        html.contains("<p class=\"line pangram\" dir=\"auto\">"),
        "a line is drawn in its own script's direction, Arabic and Hebrew being written right to left"
    );
    assert!(html.contains("#f2f2f2"), "light ink on a black backdrop");
    assert!(
        !html.contains("conic-gradient"),
        "a backdrop the controller can be given is the controller's"
    );
    assert!(
        !html.contains("text-shadow"),
        "and so is the colour the text is drawn in"
    );
    assert!(url.ends_with("?v=7"));
    assert!(
        url.contains("font-black-0.html"),
        "the page is named for the backdrop and the face it was written for"
    );

    // The two backdrops the page owns: the squares, and the ink that reads on them.
    let (page, _) =
        font_page(&font, 7, TransparentBackground::Checkerboard, 0).expect("a page");
    let html = std::fs::read_to_string(&page).expect("a written page");

    assert!(html.contains("conic-gradient"));
    assert!(html.contains("#1a1a1a"), "dark ink on the light squares");

    // And the one with no backdrop to be read against: the ink carries its own shadow.
    let (page, _) = font_page(&font, 7, TransparentBackground::Transparent, 0).expect("a page");
    let html = std::fs::read_to_string(&page).expect("a written page");

    assert!(html.contains("text-shadow"));
    assert!(!html.contains("conic-gradient"));

    let _ = std::fs::remove_dir_all(&folder);
}

/// A page of HTML is the one thing previewed in a page of this app's own, so the wrapper
/// is the one construct that widens what a previewed file may do — and what it is given
/// is two allowances, both of them the page's own: its origin, without which a page's
/// stylesheets and pictures would not load, and script, without which a page that draws
/// itself is a blank rectangle. What is kept back is every way *out* of the frame, and
/// those two together are not the loosening they would be for same-origin content, since
/// the wrapper and the document are two files and so two origins. The page is named for
/// the file it draws, so a hover on one page and then on another is two pages the browser
/// has not seen, and the URL carries the version for the reason the other pages' do.
#[test]
fn the_page_a_page_of_html_is_drawn_in_is_a_sandboxed_frame_of_the_file_itself() {
    let folder = std::env::temp_dir().join("rust-hover-preview-html-page-tests");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let one = folder.join("one page.html");
    let two = folder.join("another page.html");
    std::fs::write(&one, "<!doctype html><title>one</title>").expect("a written file");
    std::fs::write(&two, "<!doctype html><title>two</title>").expect("a written file");

    let (page, url) =
        html_page(&one, 42, TransparentBackground::Black).expect("a page for a page of html");
    let html = std::fs::read_to_string(&page).expect("a written page");

    let (other_page, _) =
        html_page(&two, 42, TransparentBackground::Black).expect("a page for a page of html");
    let other_name = other_page
        .file_name()
        .expect("a named page")
        .to_string_lossy()
        .into_owned();

    assert_ne!(
        page.file_name().expect("a named page"),
        other_page.file_name().expect("a named page"),
        "two pages of html are two pages the browser has not seen, and the target is \
         hashed into the wrapper's name so one is not rewritten under the other"
    );
    assert!(
        other_name.starts_with("html-black-"),
        "the wrapper is named for the backdrop it was written for, as {other_name} is"
    );
    assert!(url.ends_with("?v=42"), "the URL carries the version");
    assert!(url.starts_with("file:///"));

    assert!(
        html.contains("sandbox=\"allow-same-origin allow-scripts\""),
        "the frame keeps its own origin, which is what keeps its stylesheets, and is \
         given script, which is the whole of what a page that draws itself can be shown by"
    );
    let target = escape_attribute(&file_url(&one).expect("a url"));
    assert!(
        html.contains(&format!("{target}?v=42")),
        "the target is the file itself, at the version that was read"
    );

    for refused in ["allow-forms", "allow-popups", "allow-top-navigation"] {
        assert!(
            !html.contains(refused),
            "{refused} is a way out of the frame, so a page previewed cannot do it"
        );
    }

    let _ = std::fs::remove_dir_all(&folder);
}

/// A document is looked at, not handled — and the page is where that has to be said,
/// because the page is the only thing standing between the pointer and a browser that
/// answers it.
///
/// An image is draggable in a browser by default, and a drag of one is not a message this
/// app ever sees: Chromium begins its own drag session on the engine's thread, complete
/// with the ghost of the drawing following the pointer, and that session is a modal loop
/// that takes the pointer for itself. The pin takes the same pointer for itself at almost
/// the same moment (`begin_pin_drag` in `preview_window`), and two owners of one pointer
/// is a press whose release arrives to neither: the pin holds it with no release coming,
/// and the browser's drag never ends. That is the fault — a pinned document that hangs
/// the preview — and it is the *drag* of it, not the drawing, so it is the drag this page
/// refuses.
///
/// Both halves are needed, and neither is enough alone: `draggable="false"` is what tells
/// the browser the image is not a thing to be carried off the page, and `pointer-events:
/// none` is what stops the pointer from reaching it at all — no `mousedown`, so no drag to
/// begin, and nothing else either, which is the whole of what "no interaction" means here.
/// A page of HTML is the one document a hand is on, so it keeps both: the same assertion
/// against `html_page` is the other half of this test, and it is what keeps the fix from
/// being the blunt one that takes interaction away from the only document that has any.
#[test]
fn a_document_is_drawn_in_a_page_that_cannot_be_dragged_or_pointed_at() {
    let document = "file:///C:/art/a%20drawing.svg";

    // Every backdrop, because the drawing is the same page under each of them and a
    // refusal that only one of them carried would be a drag waiting to be found.
    for background in [
        TransparentBackground::Transparent,
        TransparentBackground::Black,
        TransparentBackground::White,
        TransparentBackground::Checkerboard,
    ] {
        let html = frame_html(document, 42, background);

        let start = html
            .find("<img")
            .expect("the document is an image on the page");
        let image = &html[start..start + html[start..].find('>').expect("a closed tag") + 1];
        assert!(
            image.contains("draggable=\"false\""),
            "the drawing is not a thing the browser can begin a drag of ({background:?}): \
             {image}"
        );
        assert!(
            html.contains("pointer-events:none"),
            "and the pointer does not reach it at all, so there is no press to begin one \
             with ({background:?})"
        );
    }

    // The one document that is interacted with keeps both, so the refusal above is a
    // refusal of pictures and not of the browser.
    let folder = std::env::temp_dir().join("rust-hover-preview-drag-tests");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let page_file = folder.join("a page.html");
    std::fs::write(&page_file, "<!doctype html><title>a page</title>").expect("a written page");

    let (page, _) = html_page(&page_file, 42, TransparentBackground::Black)
        .expect("a page for a page of html");
    let html = std::fs::read_to_string(&page).expect("a written page");

    assert!(
        !html.contains("draggable=\"false\""),
        "a page of HTML is a program the user clicks into, and it keeps its own drags"
    );
    assert!(
        !html.contains("pointer-events:none"),
        "and the pointer still reaches it, which is what a page being worked in means"
    );

    // A specimen is looked at on the same terms as a drawing, and is the one of the two
    // where a drag can begin in text rather than in a picture: type a pointer can sweep
    // across is type a pointer can begin a drag out of.
    let html = specimen_html(
        "file:///C:/art/specimen.ttf",
        7,
        TransparentBackground::Black,
        "Test Family Regular",
        &["The quick brown fox jumps over the lazy dog.".to_string()],
    );

    assert!(
        html.contains("pointer-events:none"),
        "a specimen is looked at, so the pointer does not reach it either"
    );
    assert!(
        html.contains("user-select:none"),
        "and none of its type can be swept across, which is what a drag out of text is"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// The one kind of document the engine hands to a browser as a page rather than as an
/// image is the one kind that runs, and it is the two names a page goes by under the one
/// setting that says a page is drawn at all. The setting is an argument here and not a
/// lock, because which names count is a question the text lists can be asked, and a test
/// that took the application's one configuration lock to ask it would be waiting on the
/// lock it is already holding. So each name is asked twice, under each setting, and the
/// setting can only turn the answer off: a document and a specimen are pictures to a
/// browser, and a picture is never anything but a picture however it was named, so `.svg`
/// and `.ttf` are not pages under either setting, and `.xhtml` is XML the text preview
/// reads, and is not one of the two names either.
#[test]
fn a_page_runs_only_under_one_of_the_two_names_a_page_goes_by() {
    for name in ["page.html", "page.htm", "some/deeper/page.HTML"] {
        assert!(
            page_runs_under(true, Path::new(name)),
            "`{name}` is a page of HTML, and a page of HTML is drawn as a page"
        );
        assert!(
            !page_runs_under(false, Path::new(name)),
            "`{name}` is markup until the tray says otherwise, and so is drawn as text"
        );
    }

    for name in ["clock.svg", "specimen.ttf", "page.xhtml"] {
        assert!(
            !page_runs_under(true, Path::new(name)),
            "`{name}` is not one of the two names a page goes by, so it is drawn and not run"
        );
        assert!(
            !page_runs_under(false, Path::new(name)),
            "`{name}` is not a page under either setting, and the switch cannot make it one"
        );
    }
}

/// A box that changes under a document is answered by the half of the ask that owns it, and
/// what owns it is the *window* rather than the want (see `place` and `box_change`).
#[test]
fn a_box_that_changes_moves_the_window_of_the_document_the_engine_holds() {
    // The document the engine is holding, at a box that has moved: the window's.
    assert_eq!(box_change(true, true, false), BoxChange::Window);

    // A document still on its way, at a box that has moved: the want's, and nothing at all is
    // asked of the engine — the page is drawn in the box the newest want asks for when it
    // lands.
    assert_eq!(box_change(true, false, false), BoxChange::Want);

    // A box that has not moved: nothing, whichever half owns it — a drag is many of these.
    assert_eq!(box_change(true, true, true), BoxChange::Nothing);
    assert_eq!(box_change(true, false, true), BoxChange::Nothing);

    // And a file the engine neither holds nor is owed has nothing to move, whatever the box
    // says: a box belongs to the preview it was measured for, and one for another file is a
    // `show` rather than this. A *held* file that is no longer wanted is the same answer —
    // a preview that has moved on has no box of this file's to move, and asking for it would
    // take the want from the file that now owns it.
    assert_eq!(box_change(false, false, false), BoxChange::Nothing);
    assert_eq!(box_change(false, true, false), BoxChange::Nothing);
}

/// A drag of a pinned document is a flood of boxes, and the engine's thread is one that
/// takes a single command per pass — so what a box is asked for has to cost the engine one
/// move however many boxes arrive, and never a backlog (see `PLACED`, `PLACE_ASKED`).
#[test]
fn a_drag_publishes_one_box_rather_than_a_queue_of_them() {
    /// The box a placement is published under, read without taking the cell's contents out
    /// of it — a test asks what is there, and the engine is what takes one.
    fn published_area() -> Option<Area> {
        PLACED
            .lock()
            .ok()
            .and_then(|placed| placed.as_ref().map(|p| p.area))
    }

    let path = Path::new("D:/Pictures/dragged.svg");
    let area = |x: i32| Area {
        x,
        y: 40,
        width: 640,
        height: 480,
    };

    // A stand-in for the engine's thread: it owes nothing, so every ask below finds no
    // engine to send to, which is the only part of `ask_place` a test can reach. What it
    // does reach is the cell, and the cell is where the coalescing is decided. The engine
    // is read in a block of its own because `ask_place` takes the same lock, and a lock
    // this thread already holds is not a lock it can wait for.
    PLACE_ASKED.store(false, Ordering::Release);
    drop_placement();
    {
        let Ok(engine) = ENGINE.lock() else {
            return;
        };
        assert!(
            engine.is_none(),
            "a placement needs no engine to be published into its cell"
        );
    }

    // A first box is published, and it is the one left behind.
    ask_place(path, area(10), TransparentBackground::Black);
    assert_eq!(
        published_area(),
        Some(area(10)),
        "a box that changed under a held document is published for the engine to move to"
    );

    // The next two are asked while the first is still owed, which is what a drag is: the
    // pointer outruns the engine's thread. Each replaces the cell, and none of them is
    // asked of the engine — which is the bound, and it is what a drag of a thousand moves
    // now costs.
    PLACE_ASKED.store(true, Ordering::Release);
    ask_place(path, area(20), TransparentBackground::Black);
    ask_place(path, area(30), TransparentBackground::Black);

    let published = published_area();
    assert_eq!(
        published,
        Some(area(30)),
        "a drag leaves the box the hand is at now, not a backlog of where it was"
    );

    // Taking the placement up is the engine's, and it gives back the flag as it does so:
    // a flag left owed is a window that stops following the hand for the rest of the run.
    let taken = take_placement().expect("the placement the drag left behind");
    assert_eq!(
        taken.area,
        area(30),
        "and it is the newest one, not the first"
    );
    assert_eq!(taken.path, path, "for the document the window is showing");
    assert!(
        !PLACE_ASKED.load(Ordering::Acquire),
        "taking a placement up gives back the flag that says one is owed"
    );

    // A hide throws the box away with the flag, so a box asked for before the window went
    // down is not carried out against whatever is put up next.
    drop_placement();
    ask_place(path, area(40), TransparentBackground::Black);
    drop_placement();
    assert!(
        published_area().is_none(),
        "a box published before a hide belongs to a window that is off screen"
    );
    assert!(
        !PLACE_ASKED.load(Ordering::Acquire),
        "and the ask goes with it, rather than being owed to a window that is gone"
    );
}

/// The keyboard follows the document, and nothing else does. A page that runs is a program
/// the user clicks into, and a click that refused to activate it would leave a page on
/// screen that no key could reach; every other document is a picture, and a picture is
/// not what the pointer being over it means the user has left the window they are working
/// in. So the two answers are one each way round, and the two styles with them: a window
/// carrying `WS_EX_NOACTIVATE` cannot be activated by anything, so it has to be taken off
/// before a click into a running page can be answered with an activation at all — and put
/// back on for the next document, which is never one.
#[test]
fn the_keyboard_goes_to_a_page_that_runs_and_to_nothing_else() {
    assert_eq!(
        mouse_activate_answers(true),
        LRESULT(1),
        "MA_ACTIVATE, so a click into a page that runs reaches the page"
    );
    assert_eq!(
        mouse_activate_answers(false),
        LRESULT(3),
        "MA_NOACTIVATE, so a click on a document leaves the caret where it is"
    );

    let bare = WS_EX_TOOLWINDOW.0 as isize | WS_EX_TOPMOST.0 as isize;
    let noactivate = WS_EX_NOACTIVATE.0 as isize;
    let refusing = bare | noactivate;

    assert_eq!(
        ex_style_for(refusing, true),
        bare,
        "a page that runs is a window that can be activated, and is otherwise untouched"
    );
    assert_eq!(
        ex_style_for(bare, true),
        bare,
        "a style with nothing to clear is left as it is"
    );
    assert_eq!(
        ex_style_for(bare, false),
        refusing,
        "a document is refused activation, and the tool window and topmost are kept"
    );
    assert_eq!(
        ex_style_for(refusing, false),
        refusing,
        "and a window already refusing is not asked twice"
    );
}

/// A font of this test's own making: an sfnt with a `cmap` covering the pangram and a
/// `name` table naming it, which is everything a specimen's page is built from.
fn specimen_font() -> Vec<u8> {
    let mut codes: Vec<u32> = "The quick brown fox jumps over the lazy dog."
        .chars()
        .map(u32::from)
        .collect();
    codes.sort_unstable();
    codes.dedup();

    let mut cmap = Vec::new();
    cmap.extend_from_slice(&12u16.to_be_bytes()); // format
    cmap.extend_from_slice(&0u16.to_be_bytes()); // reserved
    cmap.extend_from_slice(&(16u32 + codes.len() as u32 * 12).to_be_bytes()); // length
    cmap.extend_from_slice(&0u32.to_be_bytes()); // language
    cmap.extend_from_slice(&(codes.len() as u32).to_be_bytes()); // numGroups
    for (index, code) in codes.iter().enumerate() {
        cmap.extend_from_slice(&code.to_be_bytes());
        cmap.extend_from_slice(&code.to_be_bytes());
        cmap.extend_from_slice(&(index as u32 + 1).to_be_bytes());
    }

    let mut cmap_table = Vec::new();
    cmap_table.extend_from_slice(&0u16.to_be_bytes()); // version
    cmap_table.extend_from_slice(&1u16.to_be_bytes()); // numTables
    cmap_table.extend_from_slice(&3u16.to_be_bytes()); // Windows
    cmap_table.extend_from_slice(&10u16.to_be_bytes()); // UCS-4
    cmap_table.extend_from_slice(&12u32.to_be_bytes()); // offset
    cmap_table.extend_from_slice(&cmap);

    let mut strings = Vec::new();
    let mut records = Vec::new();
    for (name_id, text) in [(1u16, "Test Family"), (2u16, "Regular")] {
        let offset = strings.len();
        for unit in text.encode_utf16() {
            strings.extend_from_slice(&unit.to_be_bytes());
        }

        records.extend_from_slice(&3u16.to_be_bytes()); // Windows
        records.extend_from_slice(&1u16.to_be_bytes()); // Unicode BMP
        records.extend_from_slice(&0x0409u16.to_be_bytes()); // English (United States)
        records.extend_from_slice(&name_id.to_be_bytes());
        records.extend_from_slice(&((strings.len() - offset) as u16).to_be_bytes());
        records.extend_from_slice(&(offset as u16).to_be_bytes());
    }

    let mut name_table = Vec::new();
    name_table.extend_from_slice(&0u16.to_be_bytes()); // format
    name_table.extend_from_slice(&2u16.to_be_bytes()); // count
    name_table.extend_from_slice(&((6 + 2 * 12) as u16).to_be_bytes()); // stringOffset
    name_table.extend_from_slice(&records);
    name_table.extend_from_slice(&strings);

    let tables: [(&[u8; 4], &[u8]); 2] = [
        (b"cmap", cmap_table.as_slice()),
        (b"name", name_table.as_slice()),
    ];
    let count = tables.len();
    let mut directory = Vec::new();
    let mut data = Vec::new();

    for (tag, table) in tables {
        directory.extend_from_slice(tag);
        directory.extend_from_slice(&0u32.to_be_bytes()); // checksum, which nothing reads
        directory.extend_from_slice(&((12 + count * 16 + data.len()) as u32).to_be_bytes());
        directory.extend_from_slice(&(table.len() as u32).to_be_bytes());
        data.extend_from_slice(table);
        while !data.len().is_multiple_of(4) {
            data.push(0);
        }
    }

    let mut font = Vec::new();
    font.extend_from_slice(b"\x00\x01\x00\x00");
    font.extend_from_slice(&(count as u16).to_be_bytes());
    font.extend_from_slice(&[0u8; 6]); // the binary-search fields
    font.extend_from_slice(&directory);
    font.extend_from_slice(&data);
    font
}

/// What a hover's path becomes when it is handed to the engine: the Shell's
/// verbatim form is not a URL, and a document reached through it was a document the
/// engine never opened.
#[test]
fn turns_a_verbatim_path_into_a_url_a_browser_opens() {
    for (path, expected) in [
        (r"C:\art\clock.svg", "file:///C:/art/clock.svg"),
        (r"\\?\C:\art\clock.svg", "file:///C:/art/clock.svg"),
        (r"\\?\C:\a b\c#d.svg", "file:///C:/a%20b/c%23d.svg"),
        (r"\\?\UNC\server\share\a.svg", "file://server/share/a.svg"),
        (r"\\server\share\a.svg", "file://server/share/a.svg"),
    ] {
        assert_eq!(
            file_url(Path::new(path)).as_deref(),
            Some(expected),
            "{path}"
        );
    }

    assert_eq!(file_url(Path::new(r"relative\a.svg")), None);
}

/// What the engine costs on this machine: beginning it, pointing it at a document,
/// and pointing it at another one once it is warm. Ignored, and driven by
/// `RHP_WEBVIEW_PROBE` — `$env:RHP_WEBVIEW_PROBE = "C:\art\one.svg; C:\art\two.svg";
/// cargo test --release -- --ignored --nocapture webview_probe` — because it puts a
/// window on the screen and starts a browser.
#[test]
#[ignore = "starts the WebView2 runtime and shows a window"]
fn webview_probe() {
    let paths: Vec<PathBuf> = std::env::var("RHP_WEBVIEW_PROBE")
        .unwrap_or_default()
        .split(';')
        .map(str::trim)
        .filter(|path| !path.is_empty())
        .map(PathBuf::from)
        .collect();

    if paths.is_empty() {
        println!("set RHP_WEBVIEW_PROBE to one or more paths, separated by ';'");
        return;
    }

    println!("runtime: {:?}", runtime_version());
    println!("available: {}", is_available());

    let area = std::env::var("RHP_WEBVIEW_PROBE_AREA")
        .ok()
        .and_then(|size| {
            let (width, height) = size.split_once('x')?;
            Some(Area {
                x: 60,
                y: 60,
                width: width.trim().parse().ok()?,
                height: height.trim().parse().ok()?,
            })
        })
        .unwrap_or(Area {
            x: 60,
            y: 60,
            width: 800,
            height: 800,
        });

    for path in &paths {
        // Each document is measured from nothing on screen, so what the wait below
        // measures is this document rather than the window the last one left up.
        hide();
        let mut cleared = Duration::ZERO;
        while is_showing() && cleared < Duration::from_secs(2) {
            std::thread::sleep(Duration::from_millis(10));
            cleared += Duration::from_millis(10);
        }

        let drawn = draws(path);
        println!("{}: drawn={drawn}", path.display());

        let started = Instant::now();
        // Black by default, because a probe that measures what a document is drawn
        // at reads the screen and a transparent window shows the desktop through it,
        // which reads as a document that fills its window whatever it actually
        // draws. White is for a document drawn in dark strokes.
        let background = match std::env::var("RHP_WEBVIEW_PROBE_BACKGROUND").as_deref() {
            Ok("white") => TransparentBackground::White,
            _ => TransparentBackground::Black,
        };
        show(path, area, background);

        // The engine answers on its own thread; this is the wait for it to have
        // arrived rather than a measurement of the navigation itself.
        let mut waited = Duration::ZERO;
        while drawn && !is_showing() && waited < Duration::from_secs(5) {
            std::thread::sleep(Duration::from_millis(20));
            waited = started.elapsed();
        }

        let timings = last_timings();
        println!(
            "  showing={} after {} ms (environment {} ms, controller {} ms, navigation {} ms)",
            is_showing(),
            started.elapsed().as_millis(),
            timings.environment_ms,
            timings.controller_ms,
            timings.navigate_ms
        );

        // Kept up long enough for the screen to be looked at, which is what a probe
        // measuring what a document is drawn at needs and what a probe measuring a
        // navigation does not.
        let hold = std::env::var("RHP_WEBVIEW_PROBE_HOLD_MS")
            .ok()
            .and_then(|ms| ms.trim().parse().ok())
            .unwrap_or(1500);
        std::thread::sleep(Duration::from_millis(hold));
    }

    hide();
    std::thread::sleep(Duration::from_millis(200));
    println!("after hide: showing={}", is_showing());

    shutdown();
    println!("after shutdown: showing={}", is_showing());
}

use super::*;

/// A wait holds the pointer through the item it is waiting for rather than through the
/// box its spinner occupies: the spinner is placed at the hand and follows it, so a box
/// of its own is one the pointer can never leave — and the wait would be a preview no
/// fast hand could close on its way to somewhere else (see `preview_pointer_hold`).
#[test]
fn holds_the_pointer_through_the_item_a_wait_is_for() {
    let _stand_in = POINTER_STAND_IN.lock().expect("the pointer's own state");

    // A spinner box at the pointer, as one on screen is, and the item the hover was
    // resolved from a row away from it.
    let spinner = (100, 100, 136, 136);
    *POINTER_HOLD_REGIONS.lock().expect("the published regions") = Some(vec![spinner]);
    publish_pointer_item_box((100, 200, 300, 220));

    WAITING_PREVIEW_HOLDING.store(true, Ordering::Release);
    assert!(
        !preview_pointer_hold(118, 118),
        "the spinner's own box is not a hold"
    );
    assert!(
        preview_pointer_hold(150, 210),
        "the item the wait is for holds the pointer"
    );
    assert!(
        !preview_pointer_hold(150, 230),
        "a pointer below that item has left the file it was waiting for"
    );

    // A wait whose item nobody could be read for is not a hold: a hold is the hook
    // leaving the mouse alone, so it is only ever taken on an answer, and the reading
    // that cannot happen is a preview nothing can close.
    clear_pointer_item_box();
    assert!(
        !preview_pointer_hold(150, 210),
        "a wait with no item box of its own holds nothing"
    );

    // A hold that is no longer noted holds nothing, whatever region was left behind.
    WAITING_PREVIEW_HOLDING.store(false, Ordering::Release);
    assert!(
        !preview_pointer_hold(118, 118),
        "nothing is holding the pointer"
    );

    clear_pointer_hold();
}

/// A page the engine is drawing holds the pointer through the engine's own rectangle, and
/// through a drag that began in it once the drag has carried the pointer off the page.
///
/// The rectangle is read here as an argument and the drag as a flag rather than taken from
/// the state of the process, because both are the outside world's answer: where the engine
/// has put the window and where the hand has got to are not this side's to know. What is
/// decided here — and what the test is for — is the rule those two answers are read by, and
/// it is the whole of what makes a pointer on a page a pointer that is not a dismissal.
#[test]
fn a_page_that_runs_holds_the_pointer_where_the_engine_drew_it() {
    // The engine's window, placed where a page is drawn.
    let page = (200, 150, 600, 450);

    assert!(
        engine_page_holds(400, 300, page, true, false),
        "a pointer standing on a page that runs is the user working on the page"
    );
    assert!(
        !engine_page_holds(600, 450, page, true, false),
        "a point on the far corner is outside a half-open box"
    );
    assert!(
        !engine_page_holds(700, 300, page, true, false),
        "a pointer that has left the page has left the preview it was holding"
    );

    // A document the engine draws that is not a page that runs — a drawing, a specimen, a
    // page drawn with `render_html` switched off — is not held: the pointer on it is the
    // pointer closing the preview, which is what it always did.
    assert!(
        !engine_page_holds(400, 300, page, false, false),
        "a document that does not run holds nothing, the pointer dismisses it as before"
    );
    assert!(
        !engine_page_holds(400, 300, page, false, true),
        "and a drag standing on one is not a hold either — the page was never there to drag"
    );

    // A drag that began inside the page is the page's own, and an orbit carries the pointer
    // well outside the box the page was drawn in; that is the whole of what the latch is for.
    assert!(
        engine_page_holds(900, 700, page, true, true),
        "a drag begun on the page holds the pointer wherever the page's own view has taken it"
    );
}

/// A pinned preview is the whole of what is on screen, and no hover is shown beside it or in
/// its place — not even the hovers the loop makes for itself out of the record of what is
/// pinned, which is what a folder, a tab or a window changed under a pin used to leave behind.
/// What the pin answers for itself is not a hover and goes on (see `hover_is_shown`).
#[test]
fn a_hover_is_not_shown_while_a_preview_is_pinned() {
    let hover = PreviewMessage::Show(PathBuf::from("D:\\Pictures\\cat.png"), 100, 200, None);
    let keyboard = PreviewMessage::ShowKeyboard(
        PathBuf::from("D:\\Pictures\\cat.png"),
        100,
        200,
        300,
        220,
        None,
        false,
    );

    assert!(
        hover_is_shown(&hover, false),
        "a hover is shown while no preview is pinned"
    );
    assert!(
        !hover_is_shown(&hover, true),
        "and none is shown while one is, however the hover was asked for"
    );
    assert!(
        !hover_is_shown(&keyboard, true),
        "a keyboard hover is a hover"
    );
    assert!(
        hover_is_shown(&PreviewMessage::Hide, true),
        "while the pin's own take-down is not one"
    );
    assert!(
        hover_is_shown(
            &PreviewMessage::Pin {
                path: PathBuf::from("D:\\Pictures\\cat.png"),
                rect: (100, 200, 300, 220),
            },
            true
        ),
        "nor is the take-up that puts a window up"
    );
}

/// A reveal is held to the item the hover was resolved from: the pointer inside
/// that item's box is a pointer still on the file, one outside it is a hover that
/// has moved on, and a hover whose item was never read holds nothing back — which
/// is what keeps a frame from going up for a file the hand has already left, and
/// what keeps a keyboard preview or an unknown item from being held to a box that
/// is not theirs (see `HOVER_POINTER_BOX`).
#[test]
fn holds_a_reveal_to_the_item_the_hover_was_resolved_from() {
    let _stand_in = POINTER_STAND_IN.lock().expect("the pointer's own state");

    publish_pointer_item_box((100, 200, 300, 220));

    assert!(pointer_item_holds(100, 200), "the item's own corner holds");
    assert!(pointer_item_holds(299, 219), "and so does its far one");
    assert!(
        !pointer_item_holds(99, 210),
        "a point past its left edge does not"
    );
    assert!(!pointer_item_holds(150, 220), "nor one a row below it");

    clear_pointer_item_box();
    assert!(
        pointer_item_holds(0, 0),
        "an item nothing was read for holds anything"
    );
}

/// One hover of one file, asked exactly the six questions the `Show` arm asks and in the
/// order it asks them — which is the shape whose entry reads the test below is about.
fn ask_as_the_show_arm_does(path: &PathBuf, bounds: ScreenBounds, dpi: u32) {
    let hover = HoverFacts::read(path);

    let _ = effective_preview_scale_of(&hover, hover.scales);
    let _ = page_is_on_the_way(path);
    let _ = video_probe_due(&hover);
    let _ = media_dimensions_of(&hover, path, bounds, dpi);
    let _ = measure_waiting(path);
    let _ = hover.is_video();
}

/// Six questions about one file cost one reading of that file's directory entry, which is
/// what the entry is read for.
///
/// It was one per question. The layout asked what the content was, the scale asked it
/// again, both dimension arms asked it again, the video question asked it again and the
/// text layout asked it a sixth time — and each of those six went to the volume for the
/// same answer, on the thread that pumps this window's own messages, which is the thread a
/// pin's caption is dispatched on. A hover of a file on a slow volume paid the read six
/// times over for a question with one answer.
///
/// The figures are counted rather than argued about: `content_type::Probe::read` is the
/// one place a hover's questions reach the disk for, so its count is the count of
/// `fs::metadata` calls those questions made, whatever the caches above it did.
#[test]
fn a_hover_reads_the_files_entry_once() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let folder = std::env::temp_dir().join("rust-hover-preview-one-probe");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let picture = folder.join("tomcat.png");
    write_test_png(&picture, false);

    // A file the caches have never seen, so nothing below is answered from what a
    // previous hover of the same version left behind.
    let _ = std::fs::remove_file(&picture);
    write_test_png(&picture, false);

    crate::formats::content_type::count_entry_reads_from_now();
    ask_as_the_show_arm_does(&picture, bounds(), TEST_DPI);
    let reads = crate::formats::content_type::entry_reads();

    assert_eq!(
        reads, 1,
        "six questions about one file, and the file's directory entry read {reads} times"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// The same file read once for the whole hover rather than once for each question, and the
/// answer is the same either way: threading the reading down changed where the answer came
/// from, not what it is.
///
/// It is asked for a file each of the six questions has an opinion about — a picture under
/// a document's name, where the content is a picture and the name a document, so the two
/// answers cannot be the same question read twice.
#[test]
fn one_reading_answers_the_same_questions_six_readings_did() {
    let folder = std::env::temp_dir().join("rust-hover-preview-one-probe-answers");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let renamed = folder.join("report.docx");
    write_test_png(&renamed, false);

    let hover = HoverFacts::read(&renamed);

    // The bytes are a picture's, so the layout measures it as the picture it is whatever
    // it is called, and the loader is handed the picture's kind for the same reason.
    assert_eq!(hover.routed_kind(), Some(PreviewType::Images));
    assert!(!hover.is_video());
    assert!(!hover.is_audio());
    assert!(!hover.is_text());
    assert!(!hover.is_painted_page());

    // The name is still a document's, so the engine tier is still the one that asks
    // whether this is a document a page is owed for — and says no, which is the whole of
    // what `content_type` exists to prevent.
    assert!(previewed_as(&renamed, PreviewType::Document));
    assert!(hover.names_another_kind(PreviewType::Document));
    assert!(
        !office_render_is_due(&renamed, 800),
        "and no engine is started for a file whose bytes are a picture's"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// What a file's bytes say it is decides the share it is drawn at, the way they decide
/// the loader that draws it and the box it is placed in: a picture left under a video's
/// name is drawn at the picture's share, and one left under a document's name or a text
/// name's at the picture's share too. The shares are given values of their own so that
/// the answer says which of them was read, and the files are real ones because the
/// question is asked of the file's own header.
#[test]
fn a_share_follows_the_content_rather_than_the_name() {
    let folder = std::env::temp_dir().join("rust-hover-preview-content-share");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let scales = HoverScales {
        picture: PreviewScale::Percent(100),
        video: PreviewScale::Percent(50),
        animated: PreviewScale::Percent(100),
        document: PreviewScale::FitToScreen,
        ebook: PreviewScale::FitToScreen,
        design: PreviewScale::Percent(25),
        vector: PreviewScale::Percent(25),
        font: PreviewScale::Percent(25),
    };

    for name in ["tomcat.mp4", "tomcat.docx", "tomcat.txt"] {
        let path = folder.join(name);
        write_test_png(&path, false);

        assert_eq!(
            effective_preview_scale(&path, scales),
            PreviewScale::Percent(100),
            "`{name}` holds a picture, so the picture's share is what it is drawn at"
        );
    }

    // And a name with nothing behind it keeps the share its name asks for, which is
    // every file this module's other tests are about.
    assert_eq!(
        effective_preview_scale(Path::new(r"C:\docs\report.pdf"), scales),
        PreviewScale::FitToScreen,
        "a page is still laid out at the page's share of the room"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// And the box it is placed in is the picture's too: a picture under a text name is
/// measured as the picture it is rather than read as a page of text — which for bytes
/// that are not text is no measurement at all, and a hover that never appears for a file
/// that would otherwise be drawn.
#[test]
fn a_box_follows_the_content_rather_than_the_name() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let folder = std::env::temp_dir().join("rust-hover-preview-content-box");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let path = folder.join("tomcat.txt");
    write_test_png(&path, false);

    assert!(
        picture_dimensions(&path).is_some(),
        "the fixture is a picture the header reader can measure"
    );
    assert_eq!(
        media_dimensions(&path, bounds(), TEST_DPI),
        picture_dimensions(&path),
        "so the box is the picture's rather than the text measure's nothing"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// A page is painted into the box the layout planned for it, and an archive an engine listed is
/// one of those pages rather than a size to fit the room it was given. It is the one question
/// the share a kind is drawn at and the box the loader is handed both ask, and a kind left out
/// of either is a listing stretched to the display: what an engine's own archive was, for as
/// long as the two places listed their kinds by hand.
#[test]
fn a_page_is_painted_whether_this_app_read_the_archive_or_an_engine_listed_it() {
    let folder = std::env::temp_dir().join("rust-hover-preview-painted");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = |name: &str| folder.join(name);

    // A text file, an archive this app reads itself, and archives an engine lists — a cabinet,
    // a FreeArc archive and a stream its tool weighs: every one of them is a page.
    for name in [
        "notes.txt",
        "photos.zip",
        "backup.cab",
        "backup.arc",
        "readme.bz2",
    ] {
        let file = path(name);
        std::fs::write(&file, b"a file, of a sort").expect("a written file");

        assert!(
            page_is_painted(&file),
            "`{name}` is a page painted into the box it is given"
        );
    }

    // And a picture is not: it is a size of its own, drawn into whatever room it is given.
    let picture = path("tomcat.png");
    write_test_png(&picture, false);
    assert!(
        !page_is_painted(&picture),
        "a picture is a size of its own rather than a page"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// Every question about a file is asked of what its bytes say, which for these three is
/// the difference between an engine being started and one being left alone: no page is
/// asked of Office for a picture under a document's name, no wait for one is placed at
/// the pointer, the player takes over a video under a picture's name, and text under a
/// document's name is drawn as the text it is.
#[test]
fn a_foreign_engine_is_never_started_for_a_file_that_is_not_its_own() {
    if let Ok(mut config) = CONFIG.lock() {
        config.document_preview_enabled = true;
        config.video_preview_enabled = true;
    }

    let folder = std::env::temp_dir().join("rust-hover-preview-content-engines");
    std::fs::create_dir_all(&folder).expect("a test folder");

    // A picture under a document's name: Word is never asked for a page, and the hover
    // is not placed as the wait for one.
    let renamed = folder.join("report.docx");
    write_test_png(&renamed, false);

    assert!(
        previewed_as(&renamed, PreviewType::Document),
        "the name is the document list's, which is what answered before this"
    );
    assert!(
        !office_render_is_due(&renamed, 800),
        "and the bytes are a picture's, so no page is asked of Office"
    );
    assert!(
        !page_is_on_the_way(&renamed),
        "nor is the hover placed at the pointer as the wait for one"
    );

    // A video under a picture's name: the probe has an answer to fetch, the wait for it
    // is shown, and the player takes the window over.
    let renamed = folder.join("clip.png");
    std::fs::write(&renamed, b"\x00\x00\x00\x20ftypisom").expect("a written video");

    assert!(
        !named_as(&renamed, PreviewType::Videos),
        "the name is the picture list's, which is what answered before this"
    );
    assert!(
        HoverFacts::read(&renamed).is_video(),
        "and the bytes are a video's, so a video is what is drawn"
    );
    assert!(
        video_probe_due(&HoverFacts::read(&renamed)),
        "whose shape is probed like any other video's"
    );

    // And text under a document's name, which is the box it is painted into and the
    // frame its lines wrap in.
    let renamed = folder.join("letter.docx");
    std::fs::write(&renamed, b"{\\rtf1\\ansi\\deff0 hello}").expect("a written document");

    let listed_as_text = {
        let config = CONFIG.lock().expect("the configuration");
        crate::formats::routing::named_as(&renamed, &config, PreviewType::Text)
    };

    assert!(!listed_as_text, "the name is not one the text lists carry");
    assert!(
        is_text_preview(&renamed),
        "and the bytes are text, so text is what draws it"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// An Office document is asked of one engine and not the other: the application that owns
/// the format draws the page where it is installed, and the render engine beside it draws
/// one where it is not.
///
/// Both drawing it is what a **smaller version** of a preview in front of the right one
/// was: the engine's page is a page of the document's own size, and the hover it arrived
/// in had been laid out as the wait for a page — the spinner's box at the pointer — so it
/// was drawn at the spinner's share of the display and shown there for the second or two
/// the application took to draw the real page over it. Neither drawing it is a hover
/// waiting for a page nothing was asked to draw.
///
/// Which application is installed is the machine's answer rather than this app's, so the
/// expectation is written from it; the rule is this app's, and it is what is asserted.
#[test]
fn an_office_document_is_asked_of_one_engine_and_not_the_other() {
    if let Ok(mut config) = CONFIG.lock() {
        config.document_preview_enabled = true;
        // What the app's own `config.ini` holds is not what this test is about: it asks
        // the machine, and the setting is pinned to the one that asks the machine.
        config.office_engine = OfficeEngine::MicrosoftOffice;
        config.office_extensions = crate::formats::text_formats::sanitize_extension_list(
            crate::formats::lists::DEFAULT_OFFICE_EXTENSIONS,
        );
    }

    let folder = std::env::temp_dir().join("rust-hover-preview-office-engines");
    std::fs::create_dir_all(&folder).expect("a test folder");

    // A document its own application draws, on a machine that has one: the tier is asked
    // for the page, and the render engine is not — it would be a second rendering of the
    // same document, at a size and a place the layout had never measured a page for.
    let named = folder.join("report.docx");
    std::fs::write(&named, b"PK\x03\x04\x00\x00\x00\x00").expect("a written document");
    let installed = office_formats::app_installed(&named);

    assert!(
        previewed_as(&named, PreviewType::Document),
        "the name is the document list's, which is what a document is known by: every \
             format of Office's is a container, and a container says nothing about itself"
    );
    assert_eq!(
        office_render_is_due(&named, 800),
        installed,
        "the application that owns the format is asked for a page exactly where it is here"
    );
    assert_eq!(
        libre_formats::engine_page_kind(&named).is_some(),
        !installed,
        "and the render engine beside it is what draws one exactly where it is not, so a \
             document is never drawn twice and never left undrawn"
    );

    // And a name no family claims is nobody's: an application that could not be resolved
    // from it is not asked for a page, and the render engine draws a document of its own
    // kinds rather than one whose name it was never shown (see `engine_page_kind`).
    let unclaimed = folder.join("report.bin");
    std::fs::write(&unclaimed, b"PK\x03\x04\x00\x00\x00\x00").expect("a written document");

    assert!(
        !office_formats::app_installed(&unclaimed),
        "no family answers for a name like this one"
    );
    assert_eq!(
        libre_formats::engine_page_kind(&unclaimed),
        None,
        "and the render engine is not asked about it either"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// An engine answers a file once: the answer to a request that was folded into work
/// already in flight names the hover before the one that is waiting, and it is still
/// that wait's own answer. Read as an answer about a hover that has gone, it would
/// leave the spinner standing over work the engine has already finished.
#[test]
fn an_engine_answer_names_the_hover_before_the_one_waiting_on_it() {
    let page = PathBuf::from(r"C:\docs\notes.docx");
    let other = PathBuf::from(r"C:\docs\other.docx");

    let load = |path: PathBuf, awaiting_engine: bool| PendingLoad {
        generation: 2,
        hide_epoch: 0,
        path,
        started: Instant::now(),
        pos_x: 0,
        pos_y: 0,
        width: 64,
        height: 64,
        room: (1920, 1040),
        spinner_shown: true,
        spinner_delay: Duration::from_millis(DEFAULT_SPINNER_DELAY_MS),
        spinner_pos: (0, 0),
        spinner_side: office_preview::WAITING_BOX,
        placement: None,
        upgrade: false,
        awaiting_engine,
    };

    let waiting = load(page.clone(), true);

    // The answer to the hover before this one, for the same file, is this wait's own.
    assert!(answer_belongs_to_the_wait(
        &page,
        true,
        Some(page.as_path()),
        Some(&waiting)
    ));

    // An engine that drew nothing has nothing to replay — the load that reads the
    // file again is what takes the spinner down, so this is not an answer to hand on.
    assert!(!answer_belongs_to_the_wait(
        &page,
        false,
        Some(page.as_path()),
        Some(&waiting)
    ));

    // Another file's page is not this wait's, and a wait that is not on an engine at
    // all — a decode, a probe — is not one an engine's answer belongs to.
    assert!(!answer_belongs_to_the_wait(
        &other,
        true,
        Some(page.as_path()),
        Some(&waiting)
    ));
    assert!(!answer_belongs_to_the_wait(
        &page,
        true,
        Some(page.as_path()),
        Some(&load(page.clone(), false))
    ));

    // And nothing waiting is nothing to answer: a page that lands after the pointer
    // has left belongs to the hover that has gone.
    assert!(!answer_belongs_to_the_wait(
        &page,
        true,
        Some(page.as_path()),
        None
    ));
    assert!(!answer_belongs_to_the_wait(
        &page,
        true,
        None,
        Some(&waiting)
    ));
}

/// A preview that is still on its way follows the pointer: it is placed again
/// for a cursor that has moved along the item, kept where it is when the
/// cursor has not moved, and left alone when the hover it came from was the
/// keyboard's rather than the pointer's.
///
/// Two places follow the cursor, and they are not the same one: the preview's,
/// which is where the media lands, and the spinner's own, which is the arc at the
/// pointer's corner whatever box the preview will arrive in.
#[test]
fn a_pending_preview_follows_the_pointer() {
    let pending = |placement: Option<HoverPlacement>| PendingLoad {
        generation: 1,
        hide_epoch: 0,
        path: PathBuf::new(),
        started: Instant::now(),
        pos_x: 0,
        pos_y: 0,
        width: 0,
        height: 0,
        room: (1920, 1040),
        spinner_shown: true,
        spinner_delay: Duration::from_millis(DEFAULT_SPINNER_DELAY_MS),
        spinner_pos: (0, 0),
        spinner_side: 0,
        placement,
        upgrade: false,
        awaiting_engine: false,
    };
    let placement = HoverPlacement {
        orig_dims: (800, 600),
        avoid: None,
        follow_cursor: false,
        preview_scale: PreviewScale::FitToScreen,
        at_the_pointer_corner: false,
    };

    let mut pl = pending(Some(placement));
    let followed = pl.follow_pointer(POINT { x: 300, y: 300 }, TEST_DPI);
    assert!(followed.preview, "placed for the cursor");
    let placed = (pl.pos_x, pl.pos_y, pl.width, pl.height);

    // The cursor has not moved, so neither place has changed: nothing to move.
    assert_eq!(
        pl.follow_pointer(POINT { x: 300, y: 300 }, TEST_DPI),
        Followed {
            spinner: false,
            preview: false
        }
    );
    assert_eq!((pl.pos_x, pl.pos_y, pl.width, pl.height), placed);

    // The cursor has moved: the preview is placed again, somewhere else.
    let followed = pl.follow_pointer(POINT { x: 200, y: 300 }, TEST_DPI);
    assert!(followed.preview);
    assert_ne!((pl.pos_x, pl.pos_y, pl.width, pl.height), placed);

    // A wait is the spinner's own box rather than the preview's, whatever the
    // preview is: an 800 by 600 picture is waited for in the arc's own box at the
    // pointer's own corner, a pointer gap off it — the box a document waiting on a
    // page is placed in — while the preview keeps the place and the size its own
    // layout gave it.
    let mut waiting = pending(Some(placement));
    let followed = waiting.follow_pointer(POINT { x: 300, y: 300 }, TEST_DPI);
    assert!(followed.spinner);
    let gap = logical_px(TEST_DPI, POINTER_STANDOFF_PIXELS);
    assert_eq!(
        waiting.spinner_pos,
        (300 + gap, 300 + gap),
        "the spinner's own corner is a gap off the cursor, not under it"
    );
    assert_eq!(waiting.spinner_side, office_preview::WAITING_BOX);
    let preview = compute_mouse_layout(300, 300, placement, work_area_at(300, 300), TEST_DPI)
        .expect("a placed preview");
    assert_eq!(
        (waiting.pos_x, waiting.pos_y),
        (preview.pos_x, preview.pos_y),
        "the preview keeps the place it will arrive at"
    );
    assert_eq!(
        (waiting.width, waiting.height),
        (preview.preview_w, preview.preview_h),
        "and the box it will arrive in"
    );

    // A cursor moving along the item takes the spinner with it — still the arc's
    // own box, still a gap off the hand — and the preview's place with it.
    let followed = waiting.follow_pointer(POINT { x: 500, y: 300 }, TEST_DPI);
    assert!(followed.spinner);
    assert_eq!(waiting.spinner_pos, (500 + gap, 300 + gap));
    assert_eq!(waiting.spinner_side, office_preview::WAITING_BOX);

    // A keyboard hover's placement is the item's own: it follows nothing.
    let mut keyboard = pending(None);
    assert_eq!(
        keyboard.follow_pointer(POINT { x: 640, y: 400 }, TEST_DPI),
        Followed {
            spinner: false,
            preview: false
        }
    );
    assert_eq!((keyboard.pos_x, keyboard.pos_y), (0, 0));
}

/// The layout applies the reduction the same way the size it plans does, so a
/// preview placed for a page is the reduction of the one fit-to-screen would
/// have placed — not the room's own size with a percentage applied to it
/// somewhere else.
#[test]
fn a_layout_reduces_a_page_by_the_configured_share() {
    let page = |preview_scale: PreviewScale| HoverPlacement {
        orig_dims: (800, 600),
        avoid: None,
        follow_cursor: false,
        preview_scale,
        at_the_pointer_corner: false,
    };

    let full = compute_mouse_layout(
        300,
        300,
        page(PreviewScale::FitToScreen),
        bounds(),
        TEST_DPI,
    )
    .expect("a placed page");
    let half = compute_mouse_layout(
        300,
        300,
        page(PreviewScale::FitToScreenReduced(50)),
        bounds(),
        TEST_DPI,
    )
    .expect("a placed page");

    assert_eq!(
        (full.preview_w, full.preview_h),
        (half.preview_w * 2, half.preview_h * 2)
    );
}

/// The spinner a hover is waiting on is placed at the pointer's own corner — the one
/// of the four the display has room for, a pointer gap off the cursor so the window is
/// not under it — with no step off the name it covers, so the wait stays at the hand
/// that is waiting on it while leaving that hand free to click and probe the file
/// underneath (see `waiting_placement`).
#[test]
fn places_the_waiting_spinner_at_the_pointers_own_corner() {
    let side = office_preview::WAITING_BOX;
    let name = (100, 300, 400, 320);
    let gap = logical_px(TEST_DPI, POINTER_STANDOFF_PIXELS);
    let spinner = |cursor_x: i32, cursor_y: i32| {
        compute_mouse_layout(
            cursor_x,
            cursor_y,
            HoverPlacement {
                orig_dims: (side, side),
                avoid: Some(AvoidRegion::text(name)),
                follow_cursor: false,
                preview_scale: PreviewScale::Percent(100),
                at_the_pointer_corner: true,
            },
            bounds(),
            TEST_DPI,
        )
        .expect("a placed spinner")
    };

    // Room in every quadrant: the spinner sits in the pointer's own corner, over the
    // name it is waiting on, one pointer gap off the cursor rather than in it.
    let placement = spinner(300, 300);
    assert_eq!((placement.pos_x, placement.pos_y), (300 + gap, 300 + gap));
    assert_eq!((placement.preview_w, placement.preview_h), (side, side));

    // And the room that layout comes out at is that corner of the display rather than
    // the display itself. It is the spinner's own room and says nothing about how large
    // the preview it is waiting for will be drawn, which is why it is not the room the
    // engine that draws that preview is asked for (see `PendingLoad::room`).
    assert_eq!(
        (placement.max_width, placement.max_height),
        (
            (bounds().right - 300 - gap) as u32,
            (bounds().bottom - 300 - gap) as u32,
        ),
        "the room is the corner the spinner was put in, not the display"
    );

    // With the display ending just past the pointer there is no room in the
    // quadrant it grows into, so the spinner takes the corner that is visible:
    // its box placed to the pointer's left, the same gap off it.
    let placement = spinner(980, 300);
    assert_eq!(
        (placement.pos_x, placement.pos_y),
        (980 - gap - side as i32, 300 + gap)
    );
}

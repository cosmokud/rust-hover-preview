//! What a key press on a pinned page of text costs, and what it must not: a document read off the
//! disk, a screenful laid out and painted again for a caret that only moved, or a queue of presses
//! answered one turn at a time so that the page keeps moving after the hand has let go.

use super::*;

/// A pinned page of a real file, drawn once so the document is in hand, with the media slot put back
/// when the guard is dropped.
pub(super) struct PageOfAFile {
    /// What the media slot held before this page, put back when the guard is dropped.
    was: Option<MediaData>,
}

impl PageOfAFile {
    /// A page of `lines` lines in the scratchpad, drawn at 200 by 100 so a screenful is a handful of
    /// lines and a caret reaches the edge of it quickly.
    pub(super) fn put(label: &str, lines: usize) -> Self {
        let root = std::env::temp_dir().join("rhp-text-page-presses");
        std::fs::create_dir_all(&root).expect("a fixture directory");
        let path = root.join(format!("{label}.txt"));
        let text: String = (0..lines)
            .map(|line| format!("// line {line} of a page a press is measured on\n"))
            .collect();
        std::fs::write(&path, text).expect("a written page");

        let was = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

        let options = TextPreviewOptions {
            theme: TextTheme::Light,
            markdown_mode: MarkdownMode::Source,
            font_scale_percent: 100,
            full_mode: true,
        };
        let frame = text_preview::render_scrolled(&path, 0, 200, 100, 96, options, None, None)
            .expect("a drawn page");

        let mut media = create_loading_media(200, 100);
        media.media_type = MediaType::Text;
        media.frames[0] = Arc::new(ImageFrame::new(frame.pixels, frame.width, frame.height, 0));
        media.text_state = Some(TextPreviewState {
            path: path.clone(),
            options,
            dpi: 96,
            width: 200,
            height: 100,
            first_line: frame.first_line,
            visible_lines: frame.visible_lines,
            scrollable_lines: frame.scrollable_lines,
            scrollbar: frame.scrollbar,
            dragging: false,
            lines: frame.lines,
            selection: None,
            selecting: false,
            focused: false,
            caret: Some((0, 0)),
            caret_goal: None,
            caret_blink: text_find_field::CaretBlink::default(),
            pending_caret: Vec::new(),
            word_drag: None,
            click_run: None,
            drag_pointer: None,
            find: FindState::default(),
        });

        if let Ok(mut slot) = CURRENT_MEDIA.lock() {
            *slot = Some(media);
        }

        // The page is in the slot; its path is what the state holds, so nothing else is kept.
        let _ = &path;
        Self { was }
    }

    /// The line the page's caret stands on.
    fn caret_line(&self) -> Option<usize> {
        text_caret::caret_now().map(|place| place.0)
    }

    /// The moves the page has been told since the last tick.
    fn queued(&self) -> usize {
        CURRENT_MEDIA
            .lock()
            .ok()
            .and_then(|media| {
                media
                    .as_ref()
                    .and_then(|media| media.text_state.as_ref())
                    .map(|state| state.pending_caret.len())
            })
            .unwrap_or(0)
    }

    /// The selection the page holds now, if any.
    fn selection(&self) -> Option<text_preview::Selection> {
        CURRENT_MEDIA.lock().ok().and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .and_then(|state| state.selection)
        })
    }

    /// A press inside the page's body, which is the only way the caret's keys become the page's: the
    /// caret lands where the press did and the page takes the reader's focus (see
    /// `begin_text_selection` and `text_caret::page_has_focus`).
    fn press_page(&self, x: i32, y: i32) {
        assert!(
            begin_text_selection(x, y),
            "a press on the page's own text is the page's to answer"
        );
    }

    /// The point at the left edge of a place's own character, in the frame's coordinates: a press
    /// there lands on that place and no other (see `text_preview::caret_rect` and `position_in`).
    fn point_at(&self, place: (usize, usize)) -> (i32, i32) {
        let rect = CURRENT_MEDIA
            .lock()
            .ok()
            .and_then(|media| {
                media.as_ref().and_then(|media| {
                    media
                        .text_state
                        .as_ref()
                        .and_then(|state| text_preview::caret_rect(&state.lines, place))
                })
            })
            .expect("the page draws a character at that place");

        (rect.left + 1, (rect.top + rect.bottom) / 2)
    }

    /// The step the loop was asked for, taken off the planner's slot so the next ask is its own.
    fn walked_step(&self) -> Option<i32> {
        PIN_JOBS
            .0
            .lock()
            .ok()
            .and_then(|mut pending| pending.take())
            .and_then(|job| match job {
                PinJob::Walk { step, .. } => Some(step),
                _ => None,
            })
    }

    /// The line the frame starts at.
    fn first_line(&self) -> Option<usize> {
        CURRENT_MEDIA.lock().ok().and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| state.first_line)
        })
    }

    /// A press of `down` left for the loop, the way the pin's own window leaves one.
    fn queue_down(&self) {
        ask_pin(PinCommand::CaretDown);
    }

    /// Answer every press left since the last answer, the way the tick does, and answer what the
    /// loop was asked for: the wait a walk leaves behind, or nothing where the pin answered alone.
    fn answer_left_presses(&self) -> Option<PinWait> {
        let mut request = None;
        let mut wait = None;
        let mut started = None;
        let mut offset = 0.0;
        let mut paused = None;
        unsafe {
            pin_command_request(
                HWND(0x1000 as *mut _),
                &mut request,
                &mut wait,
                &mut started,
                &mut offset,
                &mut paused,
                true,
                None,
            );
        }
        wait
    }

    /// A press of `down` left for the loop, answered at once the way a tick answers it.
    fn press_down(&self) {
        self.queue_down();
        self.answer_left_presses();
    }

    /// Run the caret's moves the way the tick does.
    fn settle(&self) {
        unsafe {
            text_caret::settle(HWND(0x1000 as *mut _));
        }
    }
}

impl Drop for PageOfAFile {
    fn drop(&mut self) {
        if let Ok(mut slot) = CURRENT_MEDIA.lock() {
            *slot = self.was.take();
        }
    }
}

/// A page's own object, standing in for a pin: the caret's keys are the pin's only while the pin is
/// up, and a press is the pin's chrome's to leave for the loop.
struct APinForThePage(Option<PinnedPreview>);

impl APinForThePage {
    fn stand() -> Self {
        let previous = take_pin_for_a_test();
        stand_pin(Some(PinnedPreview::for_test()));
        Self(previous)
    }
}

impl Drop for APinForThePage {
    fn drop(&mut self) {
        stand_pin(self.0.take());
    }
}

/// A caret move on a page already on screen reads the page's own document rather than the disk.
///
/// The stat is what a document read starts with, and a caret moved a key at a time asked it on every
/// press: a file's metadata is a trip out of the process, and a held arrow is that per repeat. The
/// page on screen was drawn from a document already read, and the move is a question about that
/// document as it is drawn (see `drawn_lines_held`).
#[test]
fn a_warm_caret_move_reads_nothing_from_the_disk() {
    let _room = the_room_to_itself();
    let page = PageOfAFile::put("stat", 40);
    let _pin = APinForThePage::stand();
    TEXT_PREVIEW_HOLDING.store(true, Ordering::Release);

    // The press a user makes in the page before the caret's keys are the page's, and nothing of the
    // disk is read for it (see `text_caret::page_has_focus`).
    page.press_page(5, 5);

    // The render that put the page up is the one that pays the stat; everything after it is the
    // reader's.
    let stats = text_preview::document_stat_count();

    for _ in 0..3 {
        page.press_down();
        page.settle();
    }

    assert_eq!(
        page.caret_line(),
        Some(3),
        "the presses moved the caret, so the road below is the one that was measured"
    );
    assert_eq!(
        text_preview::document_stat_count(),
        stats,
        "no press asked the disk whether the file had changed: the page already has its document, \
         and a move is a question about the page"
    );

    TEXT_PREVIEW_HOLDING.store(false, Ordering::Release);
}

/// Every press a hand left behind is run in one move of the caret and drawn once.
///
/// A key held down repeats faster than a page is drawn, and the presses are counted on the hook's
/// thread and answered on the loop's: one answered per tick is a caret that keeps moving after the
/// hand has let go. They are drained together, run in the order they were pressed, and drawn as the
/// one frame they end on (see `pin_command_request` and `text_caret::settle`).
#[test]
fn a_backlog_of_presses_is_one_move_of_the_caret_and_one_frame() {
    let _room = the_room_to_itself();
    let page = PageOfAFile::put("backlog", 40);
    let _pin = APinForThePage::stand();
    TEXT_PREVIEW_HOLDING.store(true, Ordering::Release);

    page.press_page(5, 5);

    let frames = text_preview::frame_render_count();
    for _ in 0..5 {
        page.queue_down();
    }
    page.answer_left_presses();
    page.settle();

    assert_eq!(
        page.caret_line(),
        Some(5),
        "five presses are five lines, in the order they were pressed"
    );
    assert!(
        text_preview::frame_render_count() - frames <= 1,
        "and they are drawn as one frame, however many presses made them"
    );

    TEXT_PREVIEW_HOLDING.store(false, Ordering::Release);
}

/// A press that scrolls the page costs less than a frame of a 60 Hz display, and a press that does
/// not costs a fraction of that.
///
/// The budget is coarse on purpose — the machine a test runs on is not the machine a reader uses —
/// but a press that reads a document, restyles it or paints a screenful twice is far past it, and
/// that is the regression this is here to catch. Measured on the machine this was written on (a
/// debug build, which is what the suite is): 1.0–1.7 ms for a press inside the page, 4.1–4.3 ms for
/// one that scrolls it.
#[test]
fn a_caret_press_costs_less_than_a_frame() {
    let _room = the_room_to_itself();
    let page = PageOfAFile::put("cost", 40);
    let _pin = APinForThePage::stand();
    TEXT_PREVIEW_HOLDING.store(true, Ordering::Release);

    page.press_page(5, 5);

    let mut took = Vec::new();
    for _ in 0..8 {
        let began = Instant::now();
        page.press_down();
        page.settle();
        took.push(began.elapsed());
    }

    let average = took.iter().sum::<Duration>() / took.len() as u32;
    assert!(
        average < Duration::from_millis(10),
        "a press took {average:?} on average, and a page that answers a key a frame late is a \
         caret the reader is waiting for (the presses: {took:?})"
    );

    assert!(
        page.first_line().is_some_and(|first| first > 0),
        "and one of them scrolled the page, which is the press this measures"
    );

    TEXT_PREVIEW_HOLDING.store(false, Ordering::Release);
}

/// A chevron on the caption above a page of text and the menu's `Next File` row over it are the
/// folder's walk whatever the pin holds: the page answers the arrows, and nothing a hand pressed
/// (see `PinCommand::NextFile` and `pin_command_request`).
#[test]
fn a_caption_chevron_and_a_menu_row_step_the_folder_over_a_page_of_text() {
    let _room = the_room_to_itself();
    let page = PageOfAFile::put("walk", 40);
    let _pin = APinForThePage::stand();
    TEXT_PREVIEW_HOLDING.store(true, Ordering::Release);

    // Wide enough for a caption carrying the whole of its walk, which a narrow one drops whole.
    let mut pin = PinnedPreview::for_test();
    pin.content = (100, 100, 800, 300);
    stand_pin(Some(pin));

    let queued = || {
        CURRENT_MEDIA
            .lock()
            .ok()
            .and_then(|media| {
                media
                    .as_ref()
                    .and_then(|media| media.text_state.as_ref())
                    .map(|state| state.pending_caret.len())
            })
            .unwrap_or(0)
    };
    // One press through the loop's own drain, as the tick makes it: the wait a walk owes, and nothing
    // a caret move would have left behind.
    let answer_left_presses = || {
        let mut request = None;
        let mut wait = None;
        let mut started = None;
        let mut offset = 0.0;
        let mut paused = None;
        unsafe {
            pin_command_request(
                HWND(0x1000 as *mut _),
                &mut request,
                &mut wait,
                &mut started,
                &mut offset,
                &mut paused,
                true,
                None,
            );
        }
        wait
    };

    let stood = page.caret_line();

    let (x, y) = super::ws_k_chrome_nav::a_caption_button_point(pin_chrome::CaptionButton::Next);
    assert!(
        unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on the caption's walk is the caption's to answer"
    );
    let window = super::pin_input::CapturingWindow::around(RecordedPinWindow::new(0x1000));
    assert!(
        unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) },
        "and its release asks the step"
    );
    let asked = take_pin_command();
    assert_eq!(
        asked,
        Some(PinCommand::NextFile),
        "a chevron is the folder's, whatever page the pin holds"
    );
    // Left for the loop again, the way the window leaves a command it asked for.
    ask_pin(asked.expect("the press asked for a step"));
    assert!(
        answer_left_presses().is_some(),
        "so the loop asked the folder for the next file"
    );
    assert_eq!(queued(), 0, "and the press queued no move of the page");

    if let Ok(mut pending) = PIN_JOBS.0.lock() {
        *pending = None;
    }

    // The menu's `Next File` row, found by what it says rather than by arithmetic on the rows.
    unsafe { open_pin_menu(HWND(0x1000 as *mut _), 20, 20, PinMenuPage::Body) };
    let paint = pinned_menu_geometry().expect("the menu is up");
    let row = paint
        .rows
        .iter()
        .position(|row| row.label == "Next File")
        .expect("the panel holds the walk's row");
    let band = paint
        .popup
        .bands
        .iter()
        .find(|band| band.answers == row)
        .expect("the row is laid out in a band");
    let (x, y) = (
        (paint.popup.panel.left + paint.popup.panel.right) / 2,
        band.top + band.height / 2,
    );
    assert!(
        unsafe { pinned_menu_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on the row is the menu's"
    );
    let asked = take_pin_command();
    assert_eq!(
        asked,
        Some(PinCommand::NextFile),
        "and the row is the same walk the chevron asks for"
    );
    ask_pin(asked.expect("the row asked for a step"));
    assert!(
        answer_left_presses().is_some(),
        "which the loop asks the folder for"
    );
    assert_eq!(queued(), 0, "and neither press queued a caret move");
    assert_eq!(page.caret_line(), stood, "so the caret is where it stood");

    if let Ok(mut pending) = PIN_JOBS.0.lock() {
        *pending = None;
    }
    TEXT_PREVIEW_HOLDING.store(false, Ordering::Release);
}

/// A point on the caption that is no button and no resize edge: the empty drag area of the title
/// bar, which is a press that carries the window.
fn a_caption_drag_area_point() -> (i32, i32) {
    let caption = pinned_caption_geometry().expect("a pin with a caption above it");
    let framed = caption.frame != PinFrame::None;

    (0..caption.height)
        .flat_map(|y| (0..caption.width).map(move |x| (x, y)))
        .find(|(x, y)| {
            pin_chrome::button_at(*x, *y, caption.width, caption.height, caption.dpi, framed)
                .is_none()
                && !pin_state().is_some_and(|pinned| {
                    pinned
                        .pin()
                        .is_some_and(|pin| pin.resize_edge(*x, *y).is_some())
                })
        })
        .expect("this caption draws an empty drag area")
}

/// A point on the page's own scrollbar, in the frame's coordinates.
fn a_scrollbar_point() -> (i32, i32) {
    let scrollbar = CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .and_then(|state| state.scrollbar)
        })
        .expect("a page longer than its box draws a scrollbar");
    let (left, top, right, bottom) = scrollbar.track;

    ((left + right) / 2, (top + bottom) / 2)
}

/// A page nobody has pressed into walks the folder with every key the caret would take: the four
/// arrows, the two page keys and the two ends are the folder's until the page has the reader's
/// focus, and none of them queues a move of a caret the page is not drawing (see
/// `text_caret::page_has_focus` and `pin_command_request`).
#[test]
fn an_unfocused_page_walks_the_folder_with_every_walk_key() {
    let _room = the_room_to_itself();
    let page = PageOfAFile::put("walk-unfocused", 40);
    let _pin = APinForThePage::stand();
    TEXT_PREVIEW_HOLDING.store(true, Ordering::Release);

    assert!(
        !text_caret::page_has_focus(),
        "a page that has just been pinned is nobody's focus"
    );

    for (command, step) in [
        (PinCommand::Next, 1),
        (PinCommand::CaretDown, 1),
        (PinCommand::PageOn, 1),
        (PinCommand::PageBottom, 1),
        (PinCommand::Previous, -1),
        (PinCommand::CaretUp, -1),
        (PinCommand::PageBack, -1),
        (PinCommand::PageTop, -1),
    ] {
        ask_pin(command);
        assert!(
            page.answer_left_presses().is_some(),
            "{command:?} over a page nobody has pressed into is the folder's"
        );
        assert_eq!(
            page.walked_step(),
            Some(step),
            "{command:?} asked the folder for that file"
        );
        assert_eq!(
            page.queued(),
            0,
            "{command:?} queued no move of a page that has no focus"
        );
    }

    assert_eq!(
        page.caret_line(),
        Some(0),
        "and the caret the page was drawn with has not moved"
    );

    TEXT_PREVIEW_HOLDING.store(false, Ordering::Release);
}

/// The focus is the page's to take and the caption's to take away: a press inside the body makes the
/// arrow the caret's, and a press on the caption — its empty drag area and its walk alike — hands the
/// arrow back to the folder (see `text_caret::page_has_focus` and `unfocus_page`).
#[test]
fn a_page_press_takes_the_focus_and_a_caption_press_gives_it_back() {
    let _room = the_room_to_itself();
    let page = PageOfAFile::put("focus", 40);
    let _pin = APinForThePage::stand();
    TEXT_PREVIEW_HOLDING.store(true, Ordering::Release);

    // Wide enough for a caption carrying the whole of its walk.
    let mut pin = PinnedPreview::for_test();
    pin.content = (100, 100, 800, 300);
    stand_pin(Some(pin));

    let hwnd = HWND(0x1000 as *mut _);

    // The page is drawn unfocused, so the arrow is the folder's.
    ask_pin(PinCommand::Next);
    assert!(
        page.answer_left_presses().is_some(),
        "an unfocused page walks"
    );
    assert_eq!(page.queued(), 0, "and queues nothing of its own");

    // A press in the body: the caret is the page's, and the arrow with it.
    page.press_page(5, 5);
    assert!(
        text_caret::page_has_focus(),
        "a press in the page is what gives it the focus"
    );
    ask_pin(PinCommand::Next);
    assert!(
        page.answer_left_presses().is_none(),
        "a focused page answers the arrow rather than walking the folder"
    );
    assert_eq!(
        page.queued(),
        1,
        "and what it did was queue a move of the caret"
    );
    page.settle();

    // The empty drag area of the caption: the focus goes and the arrow walks again.
    let (x, y) = a_caption_drag_area_point();
    assert!(
        unsafe { pinned_press(hwnd, x, y) },
        "a hand on the caption is the caption's to answer"
    );
    assert!(
        !text_caret::page_has_focus(),
        "a press on the caption gives the page's focus up"
    );
    ask_pin(PinCommand::Next);
    assert!(
        page.answer_left_presses().is_some(),
        "so the arrow is the folder's again"
    );
    assert_eq!(page.queued(), 0, "and the page queued nothing");
    page.walked_step();

    // And a button on the caption is the same press as far as the page is concerned.
    page.press_page(5, 5);
    assert!(text_caret::page_has_focus());
    let (x, y) = super::ws_k_chrome_nav::a_caption_button_point(pin_chrome::CaptionButton::Next);
    assert!(unsafe { pinned_press(hwnd, x, y) });
    assert!(
        !text_caret::page_has_focus(),
        "a caption button gives the focus up too"
    );
    ask_pin(PinCommand::Next);
    assert!(
        page.answer_left_presses().is_some(),
        "and the arrow walks the folder from there"
    );

    if let Ok(mut pending) = PIN_JOBS.0.lock() {
        *pending = None;
    }
    TEXT_PREVIEW_HOLDING.store(false, Ordering::Release);
}

/// A press that is not on the caption leaves the page focused: the scrollbar and the find bar are
/// the page's own tools rather than the window's chrome, so the caret's keys stay the page's across
/// them (see `text_caret::unfocus_page`).
#[test]
fn a_press_on_the_scrollbar_or_the_find_bar_leaves_the_page_focused() {
    let _room = the_room_to_itself();
    let page = PageOfAFile::put("scroll-and-bar", 40);
    let _pin = APinForThePage::stand();
    TEXT_PREVIEW_HOLDING.store(true, Ordering::Release);

    let mut pin = PinnedPreview::for_test();
    pin.content = (100, 100, 800, 300);
    stand_pin(Some(pin));

    let hwnd = HWND(0x1000 as *mut _);
    page.press_page(5, 5);
    assert!(text_caret::page_has_focus());

    // The scrollbar's column, in the window's coordinates rather than the page's.
    let (x, y) = a_scrollbar_point();
    let caption = pinned_caption_geometry()
        .expect("a pin with a caption above it")
        .height;
    assert!(
        !unsafe { pinned_press(hwnd, x, y + caption) },
        "a press on the scrollbar is neither the caption's nor the window's to carry"
    );
    assert!(
        text_caret::page_has_focus(),
        "so the page keeps the focus a press in it gave"
    );
    ask_pin(PinCommand::Next);
    assert!(
        page.answer_left_presses().is_none(),
        "and the arrow is still the caret's"
    );
    assert_eq!(page.queued(), 1);
    page.settle();

    // The find bar's query, floating over the page's top-right corner.
    text_find::open_text_find();
    unsafe { text_find::settle(hwnd) };
    let card = text_find::find_bar()
        .expect("the card the bar is drawn in")
        .card;
    let (x, y) = (
        (card.query.left + card.query.right) / 2,
        (card.query.top + card.query.bottom) / 2,
    );
    assert!(text_find::press(x, y), "a press in the query is the bar's");
    assert!(
        text_caret::page_has_focus(),
        "and a press on the bar leaves the page focused"
    );

    // With the bar away again the arrow is the page's still.
    text_find::note(FindEdit::Close);
    unsafe { text_find::settle(hwnd) };
    ask_pin(PinCommand::Next);
    assert!(
        page.answer_left_presses().is_none(),
        "the page answers the arrow it was pressed into for"
    );
    assert_eq!(page.queued(), 1, "which is a move of its caret");

    TEXT_PREVIEW_HOLDING.store(false, Ordering::Release);
}

/// The caret is drawn only while the page has the focus, and taking the focus away moves it nowhere:
/// a page pressed into again draws the bar exactly where it was left, at the place a press at its own
/// coordinates lands on (see `text_caret::caret_mark` and `page_has_focus`).
#[test]
fn the_caret_is_drawn_only_while_the_page_has_the_focus() {
    let _room = the_room_to_itself();
    let page = PageOfAFile::put("caret-drawn", 40);
    let _pin = APinForThePage::stand();
    TEXT_PREVIEW_HOLDING.store(true, Ordering::Release);

    let mut pin = PinnedPreview::for_test();
    pin.content = (100, 100, 800, 300);
    stand_pin(Some(pin));

    assert!(
        text_caret::caret_mark().is_none(),
        "a page nobody has pressed into draws no caret, though it holds one"
    );

    page.press_page(5, 5);
    let drawn = text_caret::caret_mark().expect("a focused page draws the caret it holds");
    assert!(drawn.lit, "and it is solid from the press that put it down");

    // A few characters along, so the bar taken away is one the press below has to find again.
    for _ in 0..8 {
        ask_pin(PinCommand::Next);
        page.answer_left_presses();
    }
    page.settle();
    let stood = text_caret::caret_mark()
        .expect("the caret is drawn where the moves left it")
        .at;

    let (x, y) = a_caption_drag_area_point();
    assert!(unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) });
    assert!(
        text_caret::caret_mark().is_none(),
        "a press on the caption draws no caret"
    );
    assert_eq!(
        text_caret::caret_now(),
        Some(stood),
        "and the caret is still the page's, where it was left"
    );

    let (x, y) = page.point_at(stood);
    page.press_page(x, y);
    assert_eq!(
        text_caret::caret_mark()
            .expect("the page draws the caret again")
            .at,
        stood,
        "a press at the caret's own coordinates draws it where it stood"
    );

    TEXT_PREVIEW_HOLDING.store(false, Ordering::Release);
}

/// A page that has not been pressed into still reads: `Ctrl`+`A`, `Ctrl`+`C`, `F3` and `Ctrl`+`F` are
/// the page's whether or not it has the focus, because they read the document rather than move a
/// caret in it (see `text_caret::page_has_focus`).
#[test]
fn a_read_key_is_still_answered_while_the_page_is_unfocused() {
    let _room = the_room_to_itself();
    let page = PageOfAFile::put("read-keys", 40);
    let _pin = APinForThePage::stand();
    TEXT_PREVIEW_HOLDING.store(true, Ordering::Release);

    // Wide enough for a caption carrying the whole of its walk, which a narrow one drops whole.
    let mut pin = PinnedPreview::for_test();
    pin.content = (100, 100, 800, 300);
    stand_pin(Some(pin));

    let hwnd = HWND(0x1000 as *mut _);

    // A page that was worked in and then clicked away from: a press on a word, the second press that
    // takes it, and the caption pressed to give the focus up. What the reader selected is the page's
    // and stays, which is the selection the find below looks for.
    let (x, y) = page.point_at((0, 3));
    page.press_page(x, y);
    page.press_page(x, y);
    let taken = page
        .selection()
        .expect("a second press takes the word under it");
    assert!(
        !taken.is_empty(),
        "which is a selection rather than a place"
    );
    let (x, y) = a_caption_drag_area_point();
    assert!(unsafe { pinned_press(hwnd, x, y) });
    assert!(
        !text_caret::page_has_focus(),
        "the caption gave the page's focus up"
    );
    assert_eq!(
        page.selection(),
        Some(taken),
        "and the selection it was left with is untouched"
    );

    // `F3`, the page's own find next, through the loop the window leaves it with.
    ask_pin(PinCommand::FindNext);
    page.answer_left_presses();
    let found = page.selection().expect("the page still holds a selection");
    assert_ne!(
        found, taken,
        "`F3` took the selection to the next place it is found"
    );
    assert!(
        (found.anchor.0, found.anchor.1) > (taken.anchor.0, taken.anchor.1),
        "which is ahead of it in the document, and found with no focus on the page"
    );

    // `Ctrl`+`A`, which the tick answers with the whole page selected.
    unsafe { select_all_text_preview(hwnd) };
    assert_eq!(
        page.selection().map(|all| all.anchor),
        Some((0, 0)),
        "the whole page is selected from its beginning"
    );

    // And `Ctrl`+`F`, which opens the bar over it.
    text_find::open_text_find();
    assert!(
        text_find::is_open(),
        "the bar comes up over a page nobody has pressed into"
    );

    TEXT_PREVIEW_HOLDING.store(false, Ordering::Release);
}

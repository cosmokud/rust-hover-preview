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

    /// Answer every press left since the last answer, the way the tick does.
    fn press_down(&self) {
        self.queue_down();
        self.answer_left_presses();
    }

    fn answer_left_presses(&self) {
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

//! Working in a page of text: where a press anchors a selection, what a scroll does with it, and
//! what the page's own commands in the pin's menu do (see `text_selection` and `pointer`).

use super::*;

/// A page of text put where a press and a menu answer against it, with the media slot put back
/// when the guard is dropped: what a selection is made in is the lines the frame holds, and a
/// test has no document to read them out of (see `TextPreviewState`).
struct TextPagePutBack(Option<MediaData>);

impl TextPagePutBack {
    fn put() -> Self {
        let was = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
        let mut media = create_loading_media(200, 100);
        media.media_type = MediaType::Text;
        media.text_state = Some(TextPreviewState {
            path: PathBuf::from("a-page-no-test-can-read.txt"),
            options: current_text_options(),
            dpi: 96,
            width: 200,
            height: 100,
            first_line: 0,
            visible_lines: 2,
            scrollable_lines: 2,
            scrollbar: None,
            dragging: false,
            // A line is all a press on `Select All` asks the frame to hold: what the command
            // does is measured against the lines themselves and not against what they draw.
            lines: vec![text_preview::FrameLine {
                top: 0,
                height: 20,
                runs: Vec::new(),
                source: (0, 0),
            }],
            selection: None,
            selecting: false,
            word_drag: None,
            click_run: None,
            drag_pointer: None,
        });
        if let Ok(mut slot) = CURRENT_MEDIA.lock() {
            *slot = Some(media);
        }

        Self(was)
    }
}

impl Drop for TextPagePutBack {
    fn drop(&mut self) {
        if let Ok(mut slot) = CURRENT_MEDIA.lock() {
            *slot = self.0.take();
        }
    }
}

/// A pin stood up for a test and whatever was there before put back when the guard is dropped,
/// even when the test is done by panicking (see `stand_pin`).
struct PinPutBack(Option<PinnedPreview>);

impl PinPutBack {
    fn take() -> Self {
        Self(take_pin_for_a_test())
    }
}

impl Drop for PinPutBack {
    fn drop(&mut self) {
        stand_pin(self.0.take());
    }
}

/// What the page's text state holds selected now, if anything.
fn selection_now() -> Option<text_preview::Selection> {
    CURRENT_MEDIA.lock().ok().and_then(|media| {
        media
            .as_ref()
            .and_then(|media| media.text_state.as_ref())
            .and_then(|state| state.selection)
    })
}

/// The word a double click took, if the page is still holding one for the drag that followed it.
fn word_drag_now() -> Option<text_preview::Selection> {
    CURRENT_MEDIA.lock().ok().and_then(|media| {
        media
            .as_ref()
            .and_then(|media| media.text_state.as_ref())
            .and_then(|state| state.word_drag)
    })
}

/// Whether the page's text state is holding a drag open now.
fn selecting_now() -> bool {
    CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| state.selecting)
        })
        .unwrap_or(false)
}

/// A pin a menu can be placed against: a picture with room on the screen beside it for the whole
/// panel (see `pin_chrome::menu_popup_from_point`).
fn a_pin_whose_menu_fits_the_screen() -> PinnedPreview {
    let mut pin = PinnedPreview::for_test();
    pin.content = (100, 100, 500, 420);
    pin
}

/// A page of a test's own with one painted line on it and nothing done to it yet: every character
/// is ten pixels wide, so a point can be put on a character's cell by counting them, and no press,
/// no word and no selection is left over from an earlier one (see `TextPagePutBack`).
fn put_painted_line(text: &str) {
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        if let Some(state) = media.as_mut().and_then(|media| media.text_state.as_mut()) {
            state.lines = vec![text_preview::FrameLine {
                top: 0,
                height: 20,
                runs: vec![text_preview::FrameRun {
                    x: 0,
                    width: text.chars().count() as i32 * 10,
                    text: text.to_string(),
                }],
                source: (0, 0),
            }];
            state.selection = None;
            state.word_drag = None;
            state.selecting = false;
            state.drag_pointer = None;
            state.click_run = None;
        }
    }
}

/// A second press on the page, inside the time and box the machine reads a double click in, takes
/// the word under the hand rather than leaving the empty selection a single press makes — and the
/// word stays whole, because the drag a second press arms is grown by whole runs of characters
/// (see `begin_text_selection` and `extend_word_selection`).
#[test]
fn a_second_press_on_a_word_takes_the_whole_of_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _media = TextPagePutBack::put();

    put_painted_line("hello world");

    // The middle of the `l` at index 2, which is the third ten-pixel cell of the run.
    assert!(
        begin_text_selection(25, 10),
        "a press on the page is the page's"
    );
    assert_eq!(
        selection_now(),
        Some(text_preview::Selection {
            anchor: (0, 3),
            caret: (0, 3),
        }),
        "a first press is an empty selection where it landed"
    );

    // The same point again, well inside the double click time any machine is set to.
    assert!(
        begin_text_selection(25, 10),
        "a second press is the page's too"
    );
    assert_eq!(
        selection_now(),
        Some(text_preview::Selection {
            anchor: (0, 0),
            caret: (0, 5),
        }),
        "and the second of the pair takes the word under the hand"
    );
    assert!(
        selecting_now(),
        "the word is the selection the hand asked for, and the drag after it is armed to grow it \
         by whole words"
    );
}

/// Pressing the same place again climbs the ladder a reader knows: the second press takes the word
/// under the hand, and the third takes the whole line that word is on — the jump made when a word
/// is not enough and the page is more than is wanted. A press after the third stays on the line
/// rather than walking off the end of the ladder, and no drag is armed behind it: a line is whole
/// already, and a hand that moved a little would cut the highlight back to a word (see
/// `begin_text_selection` and `line_in`).
#[test]
fn a_third_press_on_a_word_takes_the_whole_of_the_line() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _media = TextPagePutBack::put();

    put_painted_line("hello world");
    let line = text_preview::Selection {
        anchor: (0, 0),
        caret: (1, 0),
    };

    assert!(
        begin_text_selection(25, 10) && begin_text_selection(25, 10),
        "the pair of presses is the page's"
    );
    assert_eq!(
        selection_now(),
        Some(text_preview::Selection {
            anchor: (0, 0),
            caret: (0, 5),
        }),
        "the second of the pair takes the word under the hand"
    );

    assert!(
        begin_text_selection(25, 10),
        "the press after that is the page's too"
    );
    assert_eq!(
        selection_now(),
        Some(line),
        "and the third of the run takes the whole line the word is on"
    );
    assert!(
        !selecting_now(),
        "nothing is armed behind a line, which is whole already"
    );

    assert_eq!(
        extend_text_selection(85, 10),
        Some(false),
        "so a hand that moves along the line is not a drag of the page's"
    );
    assert_eq!(
        selection_now(),
        Some(line),
        "and the line stays what the press took"
    );

    assert!(
        begin_text_selection(25, 10),
        "a press after the third is still the page's"
    );
    assert_eq!(
        selection_now(),
        Some(line),
        "and stays on the line, the last rung of the ladder"
    );
}

/// What the word a double click took is kept for is the drag that follows it, and that drag is
/// grown by whole runs of characters: the hand that has reached another word takes all of it, one
/// that has come back over the word the press took leaves it whole, and the release is what gives
/// the word up — a press of its own begins a drag grown by the character. That is the gesture every
/// editor has, and what makes a word taken on its last letter survive the first small move of the
/// hand (see `extend_word_selection`).
#[test]
fn a_drag_after_a_double_click_grows_the_highlight_by_whole_words() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _media = TextPagePutBack::put();
    put_painted_line("hello world");

    // A double click on the `l` at index 2 takes `hello` and arms the drag that follows it.
    assert!(
        begin_text_selection(25, 10) && begin_text_selection(25, 10),
        "the pair of presses is the page's"
    );

    // The middle of the `r` of `world`, which a drag grown by the character would have stopped
    // inside.
    assert_eq!(
        extend_text_selection(85, 10),
        Some(true),
        "a drag that has left the word the press took is the page's"
    );
    assert_eq!(
        selection_now(),
        Some(text_preview::Selection {
            anchor: (0, 0),
            caret: (0, 11),
        }),
        "the word the hand has reached is taken whole, the gap before it with it"
    );

    assert_eq!(
        extend_text_selection(15, 10),
        Some(true),
        "and a hand coming back over the word is the page's too"
    );
    assert_eq!(
        selection_now(),
        Some(text_preview::Selection {
            anchor: (0, 0),
            caret: (0, 5),
        }),
        "a hand back on the word the press took leaves it whole rather than cutting it down"
    );

    assert!(end_text_selection(), "the release ends the drag");
    assert_eq!(
        word_drag_now(),
        None,
        "and the word a double click took goes with the drag it was a drag of"
    );

    // The next press is a press of its own: nothing of the pair is left in the page, so the drag
    // after it is grown the way every other drag is, by the character under the hand.
    put_painted_line("hello world");
    assert!(
        begin_text_selection(85, 10),
        "a press on the page is the page's"
    );
    assert_eq!(
        extend_text_selection(95, 10),
        Some(true),
        "and the drag from it takes the character it has reached"
    );
    assert_eq!(
        selection_now(),
        Some(text_preview::Selection {
            anchor: (0, 9),
            caret: (0, 10),
        }),
        "which is one character wide rather than the whole word it is in"
    );
}

/// A press with Shift held keeps the anchor the selection already has and moves its far end,
/// while a plain press anchors where it landed: that is the editor's own way of growing a range
/// a click at a time (see `anchor_for_press`).
#[test]
fn a_press_with_shift_extends_the_selection_it_already_has() {
    let held = text_preview::Selection {
        anchor: (2, 3),
        caret: (4, 0),
    };

    assert_eq!(
        anchor_for_press(Some(held), true, (7, 5)),
        (2, 3),
        "a press with Shift held keeps the anchor and moves only the far end"
    );
    assert_eq!(
        anchor_for_press(Some(held), false, (7, 5)),
        (7, 5),
        "a plain press begins a selection where it landed"
    );
    assert_eq!(
        anchor_for_press(None, true, (7, 5)),
        (7, 5),
        "and a press with nothing selected has no anchor to keep"
    );
}

/// A page of a real file, loaded the way a pin loads one: a copy is taken of the document rather
/// than of the frame, so a test asking for one has to have a document to read it out of (see
/// `text_of`). The pin a full-mode page is read from is stood up by the caller.
fn a_page_of(text: &str, name: &str, width: u32, height: u32) -> Option<MediaData> {
    let folder = std::env::temp_dir().join("rust-hover-preview-text-selection");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join(name);
    std::fs::write(&path, text).expect("a written page");

    load_text_preview(&path, width, height, 96, current_text_options())
}

/// A copy with nothing selected is the whole page rather than the frame the box shows: the pin's
/// own `Copy` row asks for the page, and a page longer than its box has more in it than the frame
/// holds (see `text_preview_clipboard_text`).
#[test]
fn a_copy_with_nothing_selected_is_the_whole_page() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _pin = PinPutBack::take();
    let _media = TextPagePutBack::put();
    stand_pin(Some(PinnedPreview::for_test()));

    let Some(page) = a_page_of(
        "one\ntwo\nthree\nfour\n",
        "a-page-of-four-lines.txt",
        200,
        40,
    ) else {
        panic!("a page of text was not loaded")
    };
    let shown = page
        .text_state
        .as_ref()
        .expect("a pinned page keeps state")
        .visible_lines;
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = Some(page);
    }

    assert!(
        shown < 4,
        "the box shows less of the page than the page has, so a copy of the frame would be short"
    );
    assert_eq!(
        text_preview_clipboard_text().as_deref(),
        Some("one\r\ntwo\r\nthree\r\nfour"),
        "the copy is the whole document rather than the lines the box holds"
    );
}

/// A page of text is worked in from edge to edge: the text is what a drag selects and the margins
/// around it are where a caret at a line's own end goes, so a press anywhere on it is the page's
/// rather than a handle that carries the window (see `pinned_content_is_the_pins`).
#[test]
fn a_press_anywhere_on_a_page_of_text_is_the_pages_own() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _pin = PinPutBack::take();
    let _media = TextPagePutBack::put();

    assert!(
        !pinned_content_is_the_pins(),
        "a page of text is the page's from edge to edge, margins and all"
    );

    // And the press the window procedure makes on it is not the window's to carry: a hand that
    // lands anywhere on the page begins a selection instead (see `pinned_press`).
    let pin = overlay_pin((100, 100, 500, 420), PinChrome::always());
    let caption = pin.caption;
    stand_pin(Some(pin));
    assert!(
        !unsafe { pinned_press(HWND(0x1000 as *mut _), 10, caption + 10) },
        "a press inside a page of text begins a selection rather than carrying the window"
    );

    // And the same slot with no text in it is a picture's: a kind with nothing under the pointer
    // is a handle for moving the window (see `pinned_press`).
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        if let Some(media) = slot.as_mut() {
            media.text_state = None;
        }
    }
    assert!(
        pinned_content_is_the_pins(),
        "a kind with no text under the pointer is a handle for moving the window"
    );
}

/// A press on the page's own `Select All` selects the whole of what the frame shows, and the menu
/// goes away with the command the way a menu a command was chosen in does (see
/// `text_menu_row_press` and `select_all_text_preview`).
#[test]
fn a_press_on_select_all_selects_the_page_and_puts_the_menu_away() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _pin = PinPutBack::take();
    let _media = TextPagePutBack::put();
    stand_pin(Some(a_pin_whose_menu_fits_the_screen()));

    let hwnd = HWND(0x1000 as *mut _);
    unsafe { open_pin_menu(hwnd, 20, 20, PinMenuPage::Text) };
    let Some(paint) = pinned_menu_geometry() else {
        panic!("the page's own menu is up")
    };

    // The middle of the first row, which is the page's own `Select All`: it stands above the rows
    // the pin shares with a bubble (see `pinned_menu_paint`).
    let (x, y) = (
        (paint.popup.panel.left + paint.popup.panel.right) / 2,
        paint.popup.rows_top + paint.popup.row_height / 2,
    );
    assert_eq!(
        pin_chrome::menu_row_at(&paint.popup, x, y),
        Some(0),
        "the point this test presses is on the menu's first row"
    );

    assert!(
        unsafe { pinned_menu_press(hwnd, x, y) },
        "a hand on the page's own command is the menu's"
    );
    assert_eq!(
        selection_now(),
        Some(text_preview::Selection {
            anchor: (0, 0),
            caret: (usize::MAX, usize::MAX),
        }),
        "the command selected the whole of what the frame shows"
    );
    assert!(
        pinned_menu_geometry().is_none(),
        "and the menu goes away with the command"
    );
}

/// A selection drag carried past the edge of the box is followed by the page: the page is taken a
/// step the way the drag left by, and the selection is grown to the line at the edge the page came
/// back with, so a range longer than the box shows can be made with the pointer alone (see
/// `text_drag_step`).
#[test]
fn a_drag_carried_past_the_edge_of_a_page_keeps_selecting() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _pin = PinPutBack::take();
    let _media = TextPagePutBack::put();
    stand_pin(Some(PinnedPreview::for_test()));

    let lines: String = (0..40).map(|n| format!("line {n}\n")).collect();
    let Some(page) = a_page_of(&lines, "a-page-of-forty-lines.txt", 240, 80) else {
        panic!("a page of text was not loaded")
    };
    let (height, reachable) = page
        .text_state
        .as_ref()
        .map(|state| (state.height as i32, state.max_first_line()))
        .expect("a pinned page keeps state");
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = Some(page);
    }
    assert!(
        reachable > 0,
        "the page is longer than the box, so there is somewhere for a drag to follow it to"
    );

    // The drag is put where a hand's own moves would have left it, through the state those moves
    // write (see `extend_text_selection`).
    let drag = |selecting: bool, pointer: Option<(i32, i32)>, first_line: usize| {
        if let Ok(mut slot) = CURRENT_MEDIA.lock() {
            if let Some(state) = slot.as_mut().and_then(|media| media.text_state.as_mut()) {
                state.selecting = selecting;
                state.drag_pointer = pointer;
                state.first_line = first_line;
            }
        }
    };

    drag(false, Some((10, height + 5)), 0);
    assert_eq!(
        text_drag_step(),
        None,
        "nothing is following a pointer that is not dragging"
    );

    drag(true, Some((10, height / 2)), 0);
    assert_eq!(
        text_drag_step(),
        None,
        "and a pointer inside the box is a place the page is already showing"
    );

    drag(true, Some((10, height + 5)), 0);
    assert_eq!(
        text_drag_step(),
        Some((3, (10, height + 1))),
        "a drag below the box takes the page three lines on and grows the selection to the line \
         that comes back at the bottom"
    );

    drag(true, Some((10, -5)), 6);
    assert_eq!(
        text_drag_step(),
        Some((3, (10, -1))),
        "one above it takes the page three lines back from where it stands, to the line at the top"
    );

    drag(true, Some((10, -5)), 1);
    assert_eq!(
        text_drag_step(),
        Some((0, (10, -1))),
        "and a drag past the top of the document asks for the line it is already on, so only the \
         selection grows"
    );
}

/// A find takes the page to the next place the selection appears and selects it, wherever in the
/// document that is and whether or not the box was already showing it; and a find with nowhere to
/// go leaves the page exactly as it was, because the reader asked for the one after this rather
/// than for the top of the document (see `find_next_in_text_preview`).
#[test]
fn a_find_takes_the_page_to_the_next_place_its_selection_appears() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _pin = PinPutBack::take();
    let _media = TextPagePutBack::put();
    stand_pin(Some(PinnedPreview::for_test()));

    // Forty lines with the word on two of them, so the occurrence that is found is one the box is
    // not already showing and the page has to move to it.
    let text: String = (0..40)
        .map(|n| {
            format!(
                "line {n}{}\n",
                if n == 5 || n == 12 {
                    " needle here"
                } else {
                    ""
                }
            )
        })
        .collect();
    let Some(page) = a_page_of(&text, "a-page-with-needles.txt", 240, 60) else {
        panic!("a page of text was not loaded")
    };
    let reachable = page
        .text_state
        .as_ref()
        .map(|state| state.max_first_line())
        .expect("a pinned page keeps state");
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = Some(page);
    }
    assert!(
        reachable > 0,
        "the page is longer than the box, so a find has a page to take somewhere"
    );

    let hwnd = HWND(0x1000 as *mut _);
    let needle = |line: usize, from: usize| text_preview::Selection {
        anchor: (line, from),
        caret: (line, from + 6),
    };
    let select = |selection: text_preview::Selection| {
        if let Ok(mut slot) = CURRENT_MEDIA.lock() {
            if let Some(state) = slot.as_mut().and_then(|media| media.text_state.as_mut()) {
                state.selection = Some(selection);
            }
        }
    };
    let page_now = || {
        CURRENT_MEDIA.lock().ok().and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| (state.selection, state.first_line, state.visible_lines))
        })
    };

    // The first needle, selected where a reader's own double press would have left it.
    select(needle(5, 7));

    unsafe { find_next_in_text_preview(hwnd) };
    let (selection, first_line, visible) = page_now().expect("a pinned page keeps state");
    assert_eq!(
        selection,
        Some(needle(12, 8)),
        "the next needle is the one twelve lines down"
    );
    assert!(
        (first_line..first_line + visible).contains(&12),
        "and the page was taken to a box that shows it, which is re-laid out for the width it was \
         found at"
    );

    unsafe { find_next_in_text_preview(hwnd) };
    assert_eq!(
        page_now().and_then(|page| page.0),
        Some(needle(5, 7)),
        "and from the last of them the search wraps round to the first"
    );

    // A selection the document holds once has no next one, and asking for it leaves the page where
    // it is rather than taking it back to the top.
    select(text_preview::Selection {
        anchor: (0, 0),
        caret: (0, 6),
    });
    let before = page_now();
    unsafe { find_next_in_text_preview(hwnd) };
    assert_eq!(
        page_now(),
        before,
        "a fruitless find leaves the page exactly as it was"
    );
}

/// The page's own menu holds the same find the key does: `Find Next` is a command the page answers
/// on the spot, and the panel goes away with it the way a menu a command was chosen in does (see
/// `text_menu_row_press`).
#[test]
fn a_press_on_the_pages_find_next_row_takes_the_page_to_the_next_place() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _pin = PinPutBack::take();
    let _media = TextPagePutBack::put();
    stand_pin(Some(a_pin_whose_menu_fits_the_screen()));

    let text = "needle here\nnothing\na needle again\n";
    let Some(page) = a_page_of(text, "a-page-with-two-needles.txt", 240, 80) else {
        panic!("a page of text was not loaded")
    };
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = Some(page);
    }
    let needle = |line: usize, from: usize| text_preview::Selection {
        anchor: (line, from),
        caret: (line, from + 6),
    };
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        if let Some(state) = slot.as_mut().and_then(|media| media.text_state.as_mut()) {
            state.selection = Some(needle(0, 0));
        }
    }

    let hwnd = HWND(0x1000 as *mut _);
    unsafe { open_pin_menu(hwnd, 20, 20, PinMenuPage::Text) };
    let Some(paint) = pinned_menu_geometry() else {
        panic!("the page's own menu is up")
    };

    // The third of the page's own rows: `Select All` and `Copy` are above it (see `TEXT_MENU_ROWS`).
    let (x, y) = (
        (paint.popup.panel.left + paint.popup.panel.right) / 2,
        paint.popup.rows_top + paint.popup.row_height * 2 + paint.popup.row_height / 2,
    );
    assert_eq!(
        pin_chrome::menu_row_at(&paint.popup, x, y),
        Some(2),
        "the point this test presses is on the third of the page's own rows"
    );

    assert!(
        unsafe { pinned_menu_press(hwnd, x, y) },
        "a hand on the page's own command is the menu's"
    );
    assert_eq!(
        selection_now(),
        Some(needle(2, 2)),
        "the command found the next place the selection appears"
    );
}

/// A wheel turned with `Ctrl` over a pinned page steps through the sizes the tray's own `Text Size`
/// offers, one per notch, and a size the configuration left between two of them steps to the next one
/// in the direction the wheel was turned: the same sizes at both doors, and nothing past either end
/// of the list (see `text_size_after_notches`).
#[test]
fn a_ctrl_wheel_steps_through_the_sizes_the_tray_offers() {
    assert_eq!(
        text_size_after_notches(125, 1),
        150,
        "one notch up is the next size the tray lists"
    );
    assert_eq!(
        text_size_after_notches(150, -1),
        125,
        "and one notch down is the one below it"
    );
    assert_eq!(
        text_size_after_notches(100, 2),
        125,
        "a turn of two notches is two steps"
    );
    assert_eq!(
        text_size_after_notches(133, 1),
        150,
        "a size between two steps goes to the next one up"
    );
    assert_eq!(
        text_size_after_notches(133, -1),
        125,
        "and to the next one down the other way"
    );
    assert_eq!(
        text_size_after_notches(400, 3),
        400,
        "the largest size has nowhere larger to go"
    );
    assert_eq!(
        text_size_after_notches(70, -3),
        70,
        "and neither has the smallest"
    );
}

/// A `Ctrl` and a wheel draws a pinned page again at the next size, in the box it already has, and
/// the size is the pin's own rather than the file's or the configuration's: the page keeps its place
/// in the document and what is selected in it, and nothing is drawn at all where the wheel has
/// nowhere left to go (see `step_pinned_text_size`).
#[test]
fn a_control_wheel_draws_a_pinned_page_again_at_another_size() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _pin = PinPutBack::take();
    let _media = TextPagePutBack::put();
    stand_pin(Some(PinnedPreview::for_test()));

    let folder = std::env::temp_dir().join("rust-hover-preview-text-selection");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("a-page-to-zoom.txt");
    std::fs::write(
        &path,
        (0..40).map(|n| format!("line {n}\n")).collect::<String>(),
    )
    .expect("a written page");

    let options = TextPreviewOptions {
        font_scale_percent: 125,
        ..current_text_options()
    };
    let Some(page) = load_text_preview(&path, 240, 80, 96, options) else {
        panic!("a page of text was not loaded")
    };
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = Some(page);
    }

    // Where the reader was: a place in the document with a word of it selected.
    let selected = text_preview::Selection {
        anchor: (3, 0),
        caret: (5, 4),
    };
    let (before_height, before_width) = {
        let mut slot = CURRENT_MEDIA.lock().expect("the media slot");
        let media = slot.as_mut().expect("the page the test put there");
        let state = media
            .text_state
            .as_mut()
            .expect("a pinned page keeps state");
        state.first_line = 8;
        state.selection = Some(selected);

        (
            state.lines.first().expect("a line of the page").height,
            state.width,
        )
    };

    let hwnd = HWND(0x1000 as *mut _);
    unsafe { step_pinned_text_size(hwnd, 1) };

    let (percent, height, first_line, selection, width) = {
        let slot = CURRENT_MEDIA.lock().expect("the media slot");
        let media = slot.as_ref().expect("the page");
        let state = media
            .text_state
            .as_ref()
            .expect("a pinned page keeps state");

        (
            state.options.font_scale_percent,
            state.lines.first().expect("a line of the page").height,
            state.first_line,
            state.selection,
            state.width,
        )
    };

    assert_eq!(
        percent, 150,
        "the page is drawn at the next size the tray lists"
    );
    assert!(
        height > before_height,
        "and it is drawn larger: the glyphs are the size that changed rather than the box"
    );
    assert_eq!(
        width, before_width,
        "the page keeps the box it was drawn in, because a pin's box is the user's"
    );
    assert_eq!(
        (first_line, selection),
        (8, Some(selected)),
        "and where the reader was and what they had selected are kept, reflowed into it"
    );

    // The other way is the size it came from, and the size is held for as long as the pin lives.
    let size_now = || {
        CURRENT_MEDIA.lock().ok().and_then(|slot| {
            slot.as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| state.options.font_scale_percent)
        })
    };

    unsafe { step_pinned_text_size(hwnd, -1) };
    assert_eq!(
        size_now(),
        Some(125),
        "a notch the other way is the size it was drawn at before"
    );

    // And nothing is drawn where the wheel has nowhere to go.
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        if let Some(state) = slot.as_mut().and_then(|media| media.text_state.as_mut()) {
            state.options.font_scale_percent = 400;
        }
    }
    unsafe { step_pinned_text_size(hwnd, 1) };
    assert_eq!(
        size_now(),
        Some(400),
        "the largest size is where the scale stops"
    );
}

/// The size a reader put a page to is the reader's own, so it survives the page being laid out
/// again: a window given another box re-renders the page, and what that render reads is the
/// configuration rather than the pin — which is exactly the value the carried size replaces (see
/// `keep_text_place`).
///
/// Without this the size is a setting that lasts until the first resize, which is the one thing a
/// reader who has just chosen it is most likely to do.
#[test]
fn a_resize_keeps_the_size_the_reader_put_the_page_to() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _pin = PinPutBack::take();
    let _media = TextPagePutBack::put();
    // A page keeps its state because a pin is what draws it in full mode, and the state is the
    // whole of what a size and a place are carried on (see `load_text_preview`).
    stand_pin(Some(PinnedPreview::for_test()));

    let folder = std::env::temp_dir().join("rust-hover-preview-text-selection");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("a-page-put-to-a-size.txt");
    std::fs::write(
        &path,
        (0..40).map(|n| format!("line {n}\n")).collect::<String>(),
    )
    .expect("a written page");

    // The page as the reader left it: drawn at a size of their own, at a place in the document.
    let sized = TextPreviewOptions {
        font_scale_percent: 200,
        ..current_text_options()
    };
    let Some(page) = load_text_preview(&path, 240, 120, 96, sized) else {
        panic!("a page of text was not loaded")
    };
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = Some(page);
    }
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        if let Some(state) = slot.as_mut().and_then(|media| media.text_state.as_mut()) {
            state.first_line = 4;
        }
    }

    // A relayout answers with a page read at the configuration's size, which is the road a box
    // change takes (see `relayout_pinned_media` and `pin_media_load`).
    let Some(relaid) = load_text_preview(&path, 320, 160, 96, current_text_options()) else {
        panic!("the page was not laid out again")
    };
    assert_ne!(
        relaid
            .text_state
            .as_ref()
            .map(|state| state.options.font_scale_percent),
        Some(200),
        "the relayout read the configuration rather than the pin, which is the value to carry over"
    );

    let kept = keep_text_place(relaid);
    let state = kept.text_state.as_ref().expect("a pinned page keeps state");
    assert_eq!(
        state.options.font_scale_percent, 200,
        "the size the reader chose is carried onto the page laid out again"
    );
    assert_eq!(
        state.first_line, 4,
        "and so is the place in the document, which is what the same carry is for"
    );
}

/// The box a pin comes out at when the hand at the pointer drags one of its edges from where it
/// stands: the press that drag came out of is placed under the pointer, and the drag is applied the
/// way a pointer move applies it (see `apply_pin_drag`).
fn a_hand_pulls(edge_: PinResize, dx: i32, dy: i32) -> ScreenRegion {
    let at = cursor_screen_point().expect("the pointer is somewhere on the desktop");
    let window = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_box()))
        .expect("a pin up");

    with_pin(|pin| {
        pin.dragging = Some(PinDrag {
            from: (at.0 - dx, at.1 - dy),
            window,
            action: PinDragAction::Resize(edge_),
            delivered: true,
            carried: (i32::MIN, i32::MIN),
        });
    });
    unsafe { apply_pin_drag(HWND(0x1000 as *mut _)) };

    let after = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.content))
        .expect("a pin up");
    with_pin(|pin| pin.dragging = None);
    after
}

/// What the pin's own media holds now: the box the page on screen was drawn for, the frame that came
/// out of drawing it, and the four facts about it a resize is asked to leave alone — the size of a
/// line, the size the reader put the page to, the place in the document they were at, and how many
/// lines of it the box holds (see `TextPreviewState`).
fn the_page_now() -> ((u32, u32), (u32, u32), i32, u32, usize, usize) {
    let slot = CURRENT_MEDIA.lock().expect("the media slot");
    let media = slot.as_ref().expect("the page the test put there");
    let state = media
        .text_state
        .as_ref()
        .expect("a pinned page keeps state");

    (
        (state.width, state.height),
        (media.frames[0].width, media.frames[0].height),
        state.lines.first().expect("a line of the page").height,
        state.options.font_scale_percent,
        state.first_line,
        state.visible_lines,
    )
}

/// The last line of the file the page on screen reaches: a box that is wider fits more of a page
/// whose lines wrap, because where each of them ends is the box's answer rather than the file's.
fn the_page_reaches_now() -> usize {
    let slot = CURRENT_MEDIA.lock().expect("the media slot");

    slot.as_ref()
        .and_then(|media| media.text_state.as_ref())
        .and_then(|state| state.lines.last().map(|line| line.source.0))
        .expect("a line of the page")
}

/// A page of text under a hand is *laid out* for the box the hand is dragging it to: the same letters
/// at the same size, reflowed into a wider or a taller box — more lines of them where the box is
/// taller, and more of the document where it is wider and the lines wrap — where every other kind of
/// preview is the frame it was decoded with, sampled into that box (see `relayout_text_for_the_box`).
///
/// A resize is the one road a box changes on with a frame of its own to draw per pointer move, and a
/// page is the one kind a box *holds* rather than shapes: the font is a size the reader has put the
/// page to, and what the box decides is how many lines of it are on screen and how long each of them
/// may be before it wraps. A page left as the frame it was drawn as is that page pulled out of shape,
/// which is what a picture in a box of another shape *is* — and that is the second half below: the road
/// a page is given must not be one a picture is drawn on, or a resize would be drawing pixels of its
/// own over the ones the file answered with.
#[test]
fn a_page_is_drawn_again_for_the_box_the_hand_took_it_to_and_a_picture_is_not() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _pin = PinPutBack::take();
    let _media = TextPagePutBack::put();

    let folder = std::env::temp_dir().join("rust-hover-preview-text-selection");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("a-page-a-hand-resizes.txt");
    std::fs::write(
        &path,
        // Lines long enough to wrap in the boxes below, so that how much of the document is in the
        // window is a question about the width of the box as well as its height.
        (0..60)
            .map(|n| format!("line {n} {}\n", "of a page a hand resizes ".repeat(6)))
            .collect::<String>(),
    )
    .expect("a written page");

    // A pin standing in a box of the page's own size, in the room the app reads a box against, with
    // the chrome a page is given: bands above and below its media rather than drawn over it, and no
    // shape for a box to keep, which is how a page is framed rather than fitted (see `pin_frame`).
    let room = work_area_at(100, 100);
    let box_ = (
        room.left + 200,
        room.top + 200,
        room.left + 440,
        room.top + 320,
    );
    let mut pin = PinnedPreview::for_test();
    pin.content = box_;
    pin.frame = PinFrame::Free;
    pin.overlay = false;
    pin.zoom = None;
    stand_pin(Some(pin));

    // The page as the reader has it: at a size of their own and at a place in the document, both of
    // which a resize is asked to leave alone.
    let options = TextPreviewOptions {
        font_scale_percent: 125,
        ..current_text_options()
    };
    let Some(page) = load_text_preview(&path, 240, 120, 96, options) else {
        panic!("a page of text was not loaded")
    };
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = Some(page);
    }
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        if let Some(state) = slot.as_mut().and_then(|media| media.text_state.as_mut()) {
            state.first_line = 6;
        }
    }

    let (_, _, line_height, _, _, lines_on_screen) = the_page_now();

    // The hand pulls the right edge of the box out: the edge under the hand is the one that moves, and
    // the page is drawn again for the box it comes out at — the frame is that box's own size, so what
    // the pin composes into it is the page row for row rather than a page sampled to fit.
    assert_eq!(
        a_hand_pulls(edge(false, false, true, false), 80, 0),
        (box_.0, box_.1, box_.2 + 80, box_.3),
        "the edge under the hand is the edge that moved"
    );
    assert_eq!(
        the_page_now(),
        ((320, 120), (320, 120), line_height, 125, 6, lines_on_screen),
        "a wider box is the same page at the same size and the same place in the document, drawn \
         again for the box the hand has taken it to"
    );

    // And pulling the bottom edge down holds more of those lines rather than larger ones: what a
    // taller box buys is more of the page, which is the whole of what a page in a window is.
    assert_eq!(
        a_hand_pulls(edge(false, false, false, true), 0, 180),
        (box_.0, box_.1, box_.2 + 80, box_.3 + 180)
    );
    let (_, _, taller_line_height, size, first_line, taller_lines) = the_page_now();
    assert!(
        taller_lines > lines_on_screen,
        "a taller box holds more lines of the page: {taller_lines} against {lines_on_screen}"
    );
    assert_eq!(
        (taller_line_height, size, first_line),
        (line_height, 125, 6),
        "and they are the size the reader put the page to, at the place in the document they were at"
    );

    // And a wider box wraps each of those long lines into fewer lines of its own, so more of the
    // document is in the window at the same size of letters: what the width of the box changes is
    // where a line ends, which is the half of a page's layout the height cannot show.
    let reached = the_page_reaches_now();
    let wide = a_hand_pulls(edge(false, false, true, false), 400, 0);
    let (w, h) = (
        (wide.2 - wide.0).max(1) as u32,
        (wide.3 - wide.1).max(1) as u32,
    );
    assert!(
        w > 320,
        "the box is wider than the one the bottom edge was pulled out of: {w} pixels"
    );
    let ((drawn_w, drawn_h), (frame_w, frame_h), wider_line_height, size, first_line, _) =
        the_page_now();
    let wider_reached = the_page_reaches_now();
    assert_eq!(
        ((drawn_w, drawn_h), (frame_w, frame_h)),
        ((w, h), (w, h)),
        "a wider box is the page drawn again for it, row for row"
    );
    assert!(
        wider_reached > reached,
        "and it holds more of the document than the same box unwrapped would: {} against \
         {reached} lines of the file",
        wider_reached
    );
    assert_eq!(
        (wider_line_height, size, first_line),
        (line_height, 125, 6),
        "with the letters and the place in the document untouched by the width"
    );

    // The guard: the same drag on a picture leaves the frame the decoder answered with alone — the
    // paint samples it into the band the box has become, which is what every other kind of preview is
    // and what this road was made a page's own to keep (see `compose_media_into_band`).
    let mut pin = PinnedPreview::for_test();
    pin.content = box_;
    stand_pin(Some(pin));

    let mut picture = create_loading_media(240, 120);
    picture.media_type = MediaType::StaticImage;
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = Some(picture);
    }

    assert_ne!(
        a_hand_pulls(edge(false, false, true, false), 80, 0),
        box_,
        "the box of a picture is the one the hand dragged too"
    );
    let slot = CURRENT_MEDIA.lock().expect("the media slot");
    let picture = slot.as_ref().expect("the picture the test put there");
    assert_eq!(
        (picture.frames[0].width, picture.frames[0].height),
        (240, 120),
        "a picture under a hand is the frame it was decoded as, and only a page is drawn again"
    );
}

/// A page the reader has put to a size of their own is *loaded* at that size, so the page a box
/// change ends on is drawn at it rather than at the configuration's: the size carried onto the
/// page laid out again is the size it was asked for as well, and a page that came back at the
/// configuration's would be a reader's own size that the next resize undoes on screen while the
/// state still claims it (see `pin_media_load`, `reader_text_size` and `keep_text_place`).
///
/// A pin's own load is the one that carries a reader's size, and it is the only one: the size is
/// what the page *on screen* is drawn at, so a hover's load has none and a pin showing anything
/// but a page has none either.
#[test]
fn a_page_laid_out_again_is_loaded_at_the_size_the_reader_put_it_to() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _pin = PinPutBack::take();
    let _media = TextPagePutBack::put();
    stand_pin(Some(PinnedPreview::for_test()));

    let folder = std::env::temp_dir().join("rust-hover-preview-text-selection");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("a-page-laid-out-again.txt");
    std::fs::write(
        &path,
        (0..120).map(|n| format!("line {n}\n")).collect::<String>(),
    )
    .expect("a written page");

    // The page as the reader has it, at twice the size the configuration draws a page at: what a
    // load for the pin is asked at, and what a load for a hover is not.
    let reader_size = 200;
    let options = TextPreviewOptions {
        font_scale_percent: reader_size,
        ..current_text_options()
    };
    assert_ne!(
        current_text_options().font_scale_percent,
        reader_size,
        "the configuration's own size has to be another one for this to say anything"
    );
    let Some(page) = load_text_preview(&path, 240, 120, 96, options) else {
        panic!("a page of text was not loaded")
    };
    let sized_line = page
        .text_state
        .as_ref()
        .expect("a pinned page keeps state")
        .lines[0]
        .height;
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = Some(page);
    }

    // The box change a maximize asks for, on the road the pin's own load runs on.
    let answer = pin_media_load(&path, 400, 300, 96, PreviewScale::FitToScreen);
    let Some(relaid) = answer
        .recv_timeout(Duration::from_secs(5))
        .expect("the load answered")
    else {
        panic!("the page was not laid out again")
    };

    let state = relaid
        .text_state
        .as_ref()
        .expect("a pinned page keeps state");
    assert_eq!(
        state.options.font_scale_percent, reader_size,
        "the page is loaded at the size the reader put it to rather than at the configuration's"
    );
    assert_eq!(
        (state.width, state.height),
        (400, 300),
        "and at the box it was asked for, which is the one the window has"
    );
    assert_eq!(
        state.lines[0].height, sized_line,
        "so what it is drawn with is the reader's own letters: a page loaded at the configuration's \
         size would have come back at half this line height"
    );

    // And a load with no page in the pin to read a size off is the configuration's, which is what
    // every hover is and what a pin showing a picture walking on to a file is too.
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = None;
    }
    let answer = pin_media_load(&path, 400, 300, 96, PreviewScale::FitToScreen);
    let Some(plain) = answer
        .recv_timeout(Duration::from_secs(5))
        .expect("the load answered")
    else {
        panic!("the page was not laid out again")
    };
    assert_eq!(
        plain
            .text_state
            .as_ref()
            .expect("a pinned page keeps state")
            .options
            .font_scale_percent,
        current_text_options().font_scale_percent,
        "a load with no page of the reader's own to carry the size off is the configuration's"
    );
}

/// The theme the configuration names, stood where a test wants it and put back when the guard is
/// dropped — the house pattern for a setting a test has to move (see `CardFontSettings`).
struct ThemeSetting(TextTheme);

impl ThemeSetting {
    fn stood_at(theme: TextTheme) -> Self {
        let mut config = crate::CONFIG.lock().expect("the configuration");
        let was = Self(config.theme);
        config.theme = theme;
        was
    }
}

impl Drop for ThemeSetting {
    fn drop(&mut self) {
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.theme = self.0;
        }
    }
}

/// The settings a page's frame is painted from are in the frame — its theme and the Markdown it is
/// laid out as — so a pinned page drawn again is laid out from what the configuration now names,
/// keeping the line it is read at and the size the reader put it to, which is the page's own
/// rather than one of the settings (see `redraw_pinned_page`).
///
/// The one thing that puts a page already up to another size is the tray's own `Font Size` row, and
/// the reader's size goes with the pick the way it goes when the pin is closed.
#[test]
fn a_setting_draws_the_pinned_page_again_at_the_size_the_reader_put_it_to() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _pin = PinPutBack::take();
    let _media = TextPagePutBack::put();
    stand_pin(Some(PinnedPreview::for_test()));

    let folder = std::env::temp_dir().join("rust-hover-preview-text-selection");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("a-page-under-a-setting.txt");
    std::fs::write(
        &path,
        (0..40).map(|n| format!("line {n}\n")).collect::<String>(),
    )
    .expect("a written page");

    // The page as the reader has it, at twice the size the configuration draws a page at.
    let reader_size = 200;
    let options = TextPreviewOptions {
        font_scale_percent: reader_size,
        ..current_text_options()
    };
    assert_ne!(
        current_text_options().font_scale_percent,
        reader_size,
        "the configuration's own size has to be another one for this to say anything"
    );
    let Some(page) = load_text_preview(&path, 240, 120, 96, options) else {
        panic!("a page of text was not loaded")
    };
    let sized_line = page
        .text_state
        .as_ref()
        .expect("a pinned page keeps state")
        .lines[0]
        .height;
    let drawn_with = options.theme;
    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = Some(page);
    }

    let page_of = || {
        CURRENT_MEDIA.lock().ok().and_then(|slot| {
            slot.as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| (state.options, state.lines[0].height))
        })
    };
    assert_eq!(
        page_of().map(|(options, _)| options.theme),
        Some(drawn_with),
        "the page on screen is drawn in the theme it was loaded with"
    );

    // A theme picked in the tray, and the page laid out again for it.
    let wanted = match drawn_with {
        TextTheme::Light => TextTheme::Dark,
        _ => TextTheme::Light,
    };
    let _theme = ThemeSetting::stood_at(wanted);
    let hwnd = HWND(0x1000 as *mut _);
    unsafe { redraw_pinned_page(hwnd, None) };
    assert_eq!(
        page_of().map(|(options, _)| options.theme),
        Some(wanted),
        "a theme the configuration now names is a page painted again in it"
    );
    assert_eq!(
        page_of().map(|(options, _)| options.font_scale_percent),
        Some(reader_size),
        "and the size the reader put it to is the page's own, so it is kept"
    );
    assert_eq!(
        page_of().map(|(_, line)| line),
        Some(sized_line),
        "which is the reader's own letters, not the configuration's"
    );

    // The tray's own `Font Size` row, which is the one thing that puts a page already up to
    // another size: the reader's goes with the pick the way it goes when the pin is closed.
    unsafe { redraw_pinned_page(hwnd, Some(150)) };
    assert_eq!(
        page_of().map(|(options, _)| options.font_scale_percent),
        Some(150),
        "the size the row names is the page's new size"
    );
    assert!(
        page_of().is_some_and(|(_, line)| line < sized_line),
        "and the page is drawn at it rather than only saying so"
    );

    // A page a hover is drawing is not this: a hover is rebuilt from its own load at the
    // configuration's size, so nothing is laid out under it (see `redraw_pinned_page`).
    stand_pin(None);
    unsafe { redraw_pinned_page(hwnd, Some(300)) };
    assert_eq!(
        page_of().map(|(options, _)| options.font_scale_percent),
        Some(150),
        "a page with no pin over it is left as the hover drew it"
    );
}

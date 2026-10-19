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
            }],
            selection: None,
            selecting: false,
            last_press: None,
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

/// A second press on the page, inside the time and box the machine reads a double click in, takes
/// the word under the hand rather than leaving the empty selection a single press makes — and the
/// word stays whole, because what a double click took is not the far end of a drag (see
/// `begin_text_selection`).
#[test]
fn a_second_press_on_a_word_takes_the_whole_of_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _media = TextPagePutBack::put();

    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        if let Some(state) = media.as_mut().and_then(|media| media.text_state.as_mut()) {
            state.lines = vec![text_preview::FrameLine {
                top: 0,
                height: 20,
                runs: vec![text_preview::FrameRun {
                    x: 0,
                    width: 110,
                    text: "hello world".to_string(),
                }],
            }];
        }
    }

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
        !selecting_now(),
        "the word is the selection the hand asked for rather than the start of a drag"
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

/// A scroll keeps what is selected with the text it covers: the frame is scrolled by whole lines,
/// so both ends move by the same number of them rather than the highlight being dropped (see
/// `selection_scrolled`).
#[test]
fn a_scroll_moves_a_selection_with_the_text_it_covers() {
    let held = text_preview::Selection {
        anchor: (4, 1),
        caret: (7, 2),
    };

    assert_eq!(
        selection_scrolled(held, 3),
        text_preview::Selection {
            anchor: (1, 1),
            caret: (4, 2),
        },
        "a frame scrolled three lines on leaves the same text three lines higher"
    );
    assert_eq!(
        selection_scrolled(held, -2),
        text_preview::Selection {
            anchor: (6, 1),
            caret: (9, 2),
        },
        "and one scrolled back up two leaves it two lines lower"
    );
    assert_eq!(
        selection_scrolled(
            text_preview::Selection {
                anchor: (1, 0),
                caret: (2, 1),
            },
            4,
        ),
        text_preview::Selection {
            anchor: (0, 0),
            caret: (0, 1),
        },
        "a range the scroll takes past the top is held at the first line rather than lost"
    );
    assert_eq!(
        selection_scrolled(
            text_preview::Selection {
                anchor: (0, 0),
                caret: (usize::MAX, usize::MAX),
            },
            5,
        ),
        text_preview::Selection {
            anchor: (0, 0),
            caret: (usize::MAX - 5, usize::MAX),
        },
        "a selection of everything keeps reaching past what the frame shows"
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

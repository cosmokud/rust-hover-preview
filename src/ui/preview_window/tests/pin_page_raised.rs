//! The window a pin of a page is owed is raised where the page is drawn, and never before: a pin put
//! up on a source that has presented nothing is a band the Shell pictures with the source's window
//! icon, which is the giant default application pictogram a cold pin of an HTML page or a specimen
//! showed (see `show_pinned_window` and `the_pins_page_is_on_the_screen`).

use super::*;

/// Where a window of a test's is opened: off every display, so that nothing else in the suite can see
/// it (see `a_window_of_the_suite_outlived_its_test`).
const NOWHERE: i32 = -32000;

/// **A pin of an engine kind waits for its page, and every other pin waits for nothing.**
///
/// The gate is the engine's own answer — the file its window is showing — and not the window's existence:
/// a browser that has been created and has navigated nowhere has a window, and what the Shell draws of it
/// is its icon rather than a page. A picture, a card, a film or a drawing this app draws itself has its
/// band in hand the instant the window does, so it is never held back.
#[test]
fn a_pin_of_a_page_is_held_back_until_the_page_is_on_the_screen() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_own = pin_is_its_own_window();

    let page = Path::new("a pinned page.html");
    let drawing = Path::new("a pinned drawing.svg");

    for (path, kind) in [
        (page, MediaType::EngineSvg),
        (page, MediaType::EngineFont),
        (drawing, MediaType::EngineSvg),
    ] {
        let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
        pin.path = path.to_path_buf();
        stand_pin(Some(pin));
        pin_stands_in_its_own_window(true);
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            let mut media = create_loading_media(320, 240);
            media.media_type = kind;
            *current = Some(media);
        }

        // A browser with a window and nothing in it: the engine has published no file at all.
        webview_preview::stand_the_shown_document_in(None);
        assert!(
            !the_pins_page_is_on_the_screen(),
            "{kind:?} of {path:?} is held back until the engine names the file its window is showing"
        );

        // And a page arriving for another file is the same answer: what is on screen is not this pin's.
        webview_preview::stand_the_shown_document_in(Some(Path::new("another page.html")));
        assert!(
            !the_pins_page_is_on_the_screen(),
            "a page the engine is showing for a file the pin is not on is no page of this pin's"
        );

        // The page in hand: the pin is the window that frames it.
        webview_preview::stand_the_shown_document_in(Some(path));
        assert!(
            the_pins_page_is_on_the_screen(),
            "{kind:?} of {path:?} is raised on the page the engine has published"
        );
    }

    // A picture this app draws itself has its band the moment the window has: nothing is waited for, and
    // a page another pin was showing does not hold it back either.
    let picture = Path::new("a pinned picture.png");
    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = picture.to_path_buf();
    stand_pin(Some(pin));
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        let mut media = create_loading_media(320, 240);
        media.media_type = MediaType::StaticImage;
        *current = Some(media);
    }
    webview_preview::stand_the_shown_document_in(None);
    assert!(
        the_pins_page_is_on_the_screen(),
        "a picture pin is its own window's to draw, so it is never held back for a page"
    );

    // Nothing pinned at all is the same answer, which is the hover's road: it asks nothing of this.
    stand_pin(None);
    webview_preview::stand_the_shown_document_in(None);
    assert!(
        the_pins_page_is_on_the_screen(),
        "no pin is no page to wait for"
    );

    webview_preview::stand_the_shown_document_in(None);
    pin_stands_in_its_own_window(previous_own);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
    stand_pin(previous_pin);
}

/// **And the window it is held back from is not shown until then, while the Shell is not asked to picture
/// a band nothing has presented.**
///
/// The two halves are one decision: `show_pinned_window` refuses the show, and `the_page_was_placed` refuses
/// a *new* registration against a page that is not on the screen. A registration already standing is kept —
/// a walk between two pages is one page replacing another, never a band emptied — which is why the rule is
/// asked of a fresh registration alone (see `registration_is_owed`).
#[test]
fn the_window_and_the_registration_are_owed_the_page_rather_than_the_wait() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_own = pin_is_its_own_window();
    let previous_window = pin_real_window::real_pin_window();

    let page = Path::new("a heldback page.html");
    let mut pin = overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    );
    pin.path = page.to_path_buf();
    stand_pin(Some(pin));
    pin_stands_in_its_own_window(true);

    // The pin's own window, made off the screen and at its box, which is what a show would put up. A
    // window of a test's that is on a display is one the rest of the suite can see — a film's source
    // picker prefers the largest window on the desktop, so a test's own would be picked by the next one
    // (see `a_window_of_the_suite_outlived_its_test`).
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        let mut media = create_loading_media(320, 240);
        media.media_type = MediaType::EngineSvg;
        *current = Some(media);
    }
    webview_preview::stand_the_shown_document_in(None);
    let box_ = (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240);
    let pin_hwnd = unsafe {
        pin_real_window::stand_real_pin_window_for_a_test(None);
        pin_real_window::real_pin_window_at(box_).expect("the pin's own window")
    };

    pin_trace::clear_noted();
    // Safety: the window is one this test made, and the show only places and shows it.
    unsafe { show_pinned_window(pin_hwnd) };
    assert!(
        !unsafe { IsWindowVisible(pin_hwnd) }.as_bool(),
        "a pin whose page is not on the screen is not a window on the screen: what its band would hold \
         is the Shell's picture of a page that has presented nothing, which is the source's icon"
    );
    let trace = pin_trace::noted_lines();
    assert!(
        trace
            .iter()
            .any(|line| line.starts_with("pin window: waiting for the page of")),
        "and the wait is written down as the wait it is: {trace:?}"
    );
    assert!(
        !trace.iter().any(|line| line.starts_with("show hwnd")),
        "no show of it was made at all: {trace:?}"
    );

    // The page in hand, and the same ask puts the window up with the page in its band.
    webview_preview::stand_the_shown_document_in(Some(page));
    // Safety: as above.
    unsafe { show_pinned_window(pin_hwnd) };
    assert!(
        unsafe { IsWindowVisible(pin_hwnd) }.as_bool(),
        "a pin whose page is on the screen is the window that frames it"
    );

    // Safety: the windows are this test's own, and the state it wrote is put back below.
    unsafe {
        pin_chrome_strips::take_the_strips_down();
        let _ = DestroyWindow(pin_hwnd);
    }
    pin_real_window::stand_real_pin_window_for_a_test(previous_window);
    webview_preview::stand_the_shown_document_in(None);
    pin_stands_in_its_own_window(previous_own);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
    stand_pin(previous_pin);
}

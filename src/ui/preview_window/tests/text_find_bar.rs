//! The find bar over a pinned page: what it holds, what a key asks of it, and where its query is
//! found (see `text_find`).
//!
//! Nothing here draws the bar — the card over the page is the chrome's, and its pixels are asked
//! about with the chrome's own tests. What is asked here is the state machine a key and a tick share:
//! the query, the count, the match the reader is on, and the page that does or does not move.

use super::text_selection::{a_page_of, PinPutBack};
use super::*;
// The two keys a bar's field could not hold while the keyboard hook swallowed them: the letter `t`,
// which a film answers as a subtitle step, and the `Tab` a field swallows rather than tabs with.
use windows::Win32::UI::Input::KeyboardAndMouse::{VK_T, VK_TAB};

/// A page of a test's own, put where the bar answers against it. The media slot is put back by the
/// guards the tests take beside it (see `TextPagePutBack` and `PinPutBack`).
fn put_page(text: &str, name: &str, width: u32, height: u32) {
    let Some(page) = a_page_of(text, name, width, height) else {
        panic!("a page of text was not loaded")
    };

    if let Ok(mut slot) = CURRENT_MEDIA.lock() {
        *slot = Some(page);
    }
}

/// A handle no road here draws into: the bar's own searches and the page's state are what a test
/// asks about, and a window of the suite's would be a window made for nothing.
fn a_pin_window() -> HWND {
    HWND(0x1000 as *mut _)
}

/// Whether the bar is up, its query, how many places it found, which of them is current, and the two
/// switches — the whole of what a count and a step are read off.
fn bar_now() -> (bool, String, usize, Option<usize>, bool, bool) {
    CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| {
                    (
                        state.find.open,
                        state.find.field.text().to_string(),
                        state.find.matches.len(),
                        state.find.current,
                        state.find.case_sensitive,
                        state.find.whole_word,
                    )
                })
        })
        .expect("a pinned page keeps a find bar")
}

/// The document line the page is scrolled to.
fn first_line_now() -> usize {
    CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| state.first_line)
        })
        .expect("a pinned page keeps its place")
}

/// Where the caret and the other end of its selection stand in the query, which is what a move and an
/// extension are read off (see `text_find_field::Field::caret` and `anchor`).
fn caret_now() -> (usize, usize) {
    CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| (state.find.field.caret(), state.find.field.anchor()))
        })
        .expect("a pinned page keeps a find bar")
}

/// Clear whatever the bar was last asked, so that what a helper reads next is this key's alone.
fn clear_the_queue() {
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        if let Some(find) = media
            .as_mut()
            .and_then(|media| media.text_state.as_mut())
            .map(|state| &mut state.find)
        {
            find.pending.clear();
        }
    }
}

/// Type a run of characters into the bar, searching after each one the way the tick does, and settle
/// whatever the last one asked for.
fn type_into_the_bar(text: &str) {
    for character in text.chars() {
        text_find::note(FindEdit::Typed(character));
        unsafe { text_find::settle(a_pin_window()) };
    }
}

/// A page of forty lines with a needle on two of them, fourteen lines apart, so a match can be
/// outside the box the reader is looking at.
fn a_page_with_two_needles(name: &str) {
    let text: String = (0..40)
        .map(|n| {
            format!(
                "line {n}{}\n",
                if n == 5 || n == 19 {
                    " needle here"
                } else {
                    ""
                }
            )
        })
        .collect();

    put_page(&text, name, 240, 60);
}

/// An empty query finds nothing and finds it for every character typed: the query, the matches and
/// the current one are settled on every keystroke, and a backspace is a keystroke like any other
/// (see `text_find::settle`).
#[test]
fn a_query_is_typed_into_the_bar_and_every_keystroke_finds_it() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    put_page(
        "needle here\nnothing\nneedle again\n",
        "a-bar-typed-into.txt",
        240,
        80,
    );

    text_find::open_text_find();
    unsafe { text_find::settle(a_pin_window()) };

    let (open, query, count, current, ..) = bar_now();
    assert!(open, "the bar is up");
    assert_eq!(query, "", "and it opens with nothing to look for");
    assert_eq!(
        (count, current),
        (0, None),
        "an empty query holds no match and no current one"
    );

    type_into_the_bar("need");
    let (_, query, count, current, ..) = bar_now();
    assert_eq!(query, "need");
    assert_eq!(count, 2, "both needles are found, one at a time");
    assert_eq!(
        current,
        Some(0),
        "and the reader is put on the first match at or after their line"
    );

    type_into_the_bar("lx");
    let (_, query, count, current, ..) = bar_now();
    assert_eq!(query, "needlx");
    assert_eq!(
        (count, current),
        (0, None),
        "a query found nowhere has no current match"
    );

    text_find::note(FindEdit::Backspace);
    text_find::note(FindEdit::Backspace);
    unsafe { text_find::settle(a_pin_window()) };
    let (_, query, count, ..) = bar_now();
    assert_eq!(query, "need");
    assert_eq!(
        count, 2,
        "a backspace puts back the query the document holds"
    );
}

/// A new query starts at the first match at or after the line the reader was on, and a step walks
/// on from there and wraps at either end — the count is the match the reader is on, not the first
/// one found (see `text_find::settle` and `stepped`).
#[test]
fn a_query_starts_at_the_line_the_reader_is_on_and_a_step_wraps() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_page_with_two_needles("a-bar-started-at.txt");

    text_find::open_text_find();
    type_into_the_bar("needle");

    let (_, _, count, current, ..) = bar_now();
    assert_eq!(
        (count, current),
        (2, Some(0)),
        "the top of the page is the first needle"
    );

    // Step on once: the second match, and from there round to the first.
    text_find::note(FindEdit::Step(1));
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(bar_now().3, Some(1), "a step on takes the next match");

    text_find::note(FindEdit::Step(1));
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(
        bar_now().3,
        Some(0),
        "and from the last one it wraps to the first"
    );

    text_find::note(FindEdit::Step(-1));
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(bar_now().3, Some(1), "a step back wraps the other way");

    // A reader who has scrolled the page starts at the match below the line they are on rather than
    // at the top of the document: line 19's needle, not line 5's.
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        if let Some(state) = media.as_mut().and_then(|media| media.text_state.as_mut()) {
            state.first_line = 10;
        }
    }
    text_find::note(FindEdit::SearchAgain);
    unsafe { text_find::settle(a_pin_window()) };

    let (_, _, count, current, ..) = bar_now();
    assert_eq!(count, 2, "the query still finds both needles");
    assert_eq!(
        current,
        Some(1),
        "and the current one is the first at or after the line the reader is on"
    );
}

/// Closing the bar puts its highlights away and keeps its query: the page stays exactly where the
/// find left it, and the key that opens the bar again searches that query again rather than starting
/// from nothing (see `FindState` and `text_find::open_text_find`).
#[test]
fn a_close_puts_the_highlights_away_and_leaves_the_page_where_it_is() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_page_with_two_needles("a-bar-closed.txt");

    text_find::open_text_find();
    type_into_the_bar("needle");
    text_find::note(FindEdit::Step(1));
    unsafe { text_find::settle(a_pin_window()) };

    let scrolled_to = first_line_now();
    assert!(
        scrolled_to > 0,
        "the second needle is outside the box, so the page was taken to it"
    );

    text_find::note(FindEdit::Close);
    unsafe { text_find::settle(a_pin_window()) };

    let (open, query, count, current, ..) = bar_now();
    assert!(!open, "the bar is away");
    assert_eq!(query, "needle", "and it keeps the query it was closed with");
    assert_eq!(
        (count, current),
        (0, None),
        "while the places it found are not drawn by anything any more"
    );
    assert_eq!(
        first_line_now(),
        scrolled_to,
        "a close is the bar going, not the page moving back"
    );

    text_find::open_text_find();
    unsafe { text_find::settle(a_pin_window()) };
    let (open, query, count, current, ..) = bar_now();
    assert!(open, "and opening it again is the bar back at that query");
    assert_eq!(query, "needle");
    assert_eq!(
        (count, current),
        (2, Some(1)),
        "with the query found again and the reader on the match at their line"
    );
}

/// The two switches are searched again under their own answer: case-blind by default, and a
/// whole-word search takes only what stands as a word (see `text_find::settle`).
#[test]
fn the_two_switches_search_again_under_their_own_answer() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    put_page("Cat cat CAT\nscat cat\n", "a-bar-two-switches.txt", 240, 80);

    text_find::open_text_find();
    type_into_the_bar("cat");

    let (_, _, count, _, case_sensitive, whole_word) = bar_now();
    assert_eq!(
        (case_sensitive, whole_word),
        (false, false),
        "both switches are off"
    );
    assert_eq!(
        count, 5,
        "every case of the letters is found, the one inside `scat` too"
    );

    text_find::note(FindEdit::ToggleCase);
    unsafe { text_find::settle(a_pin_window()) };
    let (_, _, count, _, case_sensitive, _) = bar_now();
    assert!(case_sensitive, "the first switch is on");
    assert_eq!(
        count, 3,
        "and only the lower-case occurrences are left, the one inside `scat` among them"
    );

    text_find::note(FindEdit::ToggleCase);
    text_find::note(FindEdit::ToggleWholeWord);
    unsafe { text_find::settle(a_pin_window()) };
    let (_, _, count, _, case_sensitive, whole_word) = bar_now();
    assert_eq!(
        (case_sensitive, whole_word),
        (false, true),
        "case is back off and whole word on"
    );
    assert_eq!(
        count, 4,
        "every `cat` that stands as a word, and not the one inside `scat`"
    );
}

/// **The bar is the one owner of the pin's keys, and the pin's own walk never sees them.** With a bar
/// up the keyboard hook stands down and every key reaches the pin's window as a message, which is
/// what hands it to the field: `←`/`→` move the caret rather than walking the folder, `PgUp`/`PgDn`
/// step the matches rather than turning the page, and a `Space`, a `T` and a `Tab` ask nothing at all
/// (see `text_find::key_edit_with`). A bar that is away leaves every one of them to the pin.
#[test]
fn the_bars_keys_are_not_the_pins_while_it_is_up() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_page_with_two_needles("a-bar-inert-keys.txt");

    text_find::open_text_find();
    type_into_the_bar("needle");

    // One key through the window's own procedure, answering what the pin was asked for. Nothing is
    // queued for the pin while the bar is up, so `None` is the arm the bar returned on.
    let press = |vk: u16| {
        take_pin_command();
        unsafe { window_proc(a_pin_window(), WM_KEYDOWN, WPARAM(vk as usize), LPARAM(0)) };
        unsafe { text_find::settle(a_pin_window()) };
        take_pin_command()
    };

    assert_eq!(caret_now(), (6, 6), "the caret is after the query typed");
    let (_, query, count, current, ..) = bar_now();
    let standing = (query, count, current, first_line_now());

    assert_eq!(
        press(VK_LEFT.0),
        None,
        "the left arrow walks no folder under a query"
    );
    assert_eq!(caret_now(), (5, 5), "it moves the caret inside the query");
    assert_eq!(press(VK_RIGHT.0), None);
    assert_eq!(caret_now(), (6, 6), "and the right arrow moves it back");

    // The page keys step the matches rather than the page, and the caret stays where it was.
    assert_eq!(press(VK_PRIOR.0), None);
    assert_eq!(
        bar_now().3,
        Some(1),
        "`PgUp` steps the matches back, wrapping to the last"
    );
    assert_eq!(press(VK_NEXT.0), None);
    assert_eq!(bar_now().3, Some(0), "and `PgDn` steps them on again");

    // A Space, a subtitle step and a `Tab` are inert: they are characters a query does not type here,
    // and none of them is the pin's while the bar is up.
    for vk in [VK_SPACE.0, VK_T.0, VK_TAB.0] {
        assert_eq!(press(vk), None, "{vk:#06x} is nobody's key under a query");
    }
    let (_, query, count, current, ..) = bar_now();
    assert_eq!(
        (query, count, current, first_line_now()),
        standing,
        "and nothing the page answers moved under the query: not its place, not its matches"
    );

    // A bar that is away leaves the pin's keys to the pin, which is the gate on the other side.
    text_find::note(FindEdit::Close);
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(
        press(VK_RIGHT.0),
        Some(PinCommand::Next),
        "with the bar away the right arrow is the pin's walk again"
    );
}

/// The caret's keys on the one road a focused pin gives the bar: a `WM_KEYDOWN` read by the window's
/// own `key_edit_with`, with no second copy of them anywhere. `Shift` extends from the caret, `Ctrl`
/// moves by a word, and a page key steps the matches without touching the caret (see
/// `text_find::key_edit_with` and `window_proc`).
#[test]
fn the_caret_keys_move_the_caret_under_a_query() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_page_with_two_needles("a-bar-caret-keys.txt");

    text_find::open_text_find();
    type_into_the_bar("needle");

    let (_, query, count, ..) = bar_now();
    assert_eq!(caret_now(), (6, 6), "the caret is after the query typed");

    let plain = text_find::Modifiers::default();
    let shift = text_find::Modifiers {
        shift: true,
        control: false,
    };
    let control = text_find::Modifiers {
        shift: false,
        control: true,
    };
    let press = |vk: u16, held| {
        let edit =
            text_find::key_edit_with(vk as i32, held).expect("a key the field's one table answers");
        text_find::note(edit);
        unsafe { text_find::settle(a_pin_window()) };
    };

    // One character either way.
    press(VK_LEFT.0, plain);
    assert_eq!(
        caret_now(),
        (5, 5),
        "`←` moves the caret one character back"
    );
    press(VK_RIGHT.0, plain);
    assert_eq!(caret_now(), (6, 6), "`→` moves it forward again");

    // The two ends.
    press(VK_HOME.0, plain);
    assert_eq!(caret_now(), (0, 0), "`Home` puts the caret at the start");
    press(VK_END.0, plain);
    assert_eq!(caret_now(), (6, 6), "`End` puts it at the end");

    // `Shift` extends rather than collapses, and moves no text.
    press(VK_LEFT.0, shift);
    assert_eq!(
        caret_now(),
        (5, 6),
        "`Shift`+`←` grows the selection back from the caret"
    );
    assert_eq!(bar_now().1, query, "and the query's text is left alone");

    // `Ctrl` moves by a word.
    press(VK_LEFT.0, control);
    assert_eq!(
        caret_now(),
        (0, 0),
        "`Ctrl`+`←` goes to the start of the word"
    );
    press(VK_RIGHT.0, control);
    assert_eq!(caret_now(), (6, 6), "`Ctrl`+`→` goes to the end of it");

    // The page keys step the matches and leave the caret where they found it.
    let (caret, anchor) = caret_now();
    press(VK_NEXT.0, plain);
    assert_eq!(bar_now().3, Some(1), "the next match is the reader's");
    assert_eq!((caret, anchor), caret_now(), "and the caret did not move");
    assert_eq!(bar_now().2, count, "over the places the same query found");
    press(VK_PRIOR.0, plain);
    assert_eq!(
        bar_now().3,
        Some(0),
        "and the page key back steps them back"
    );
    assert_eq!(
        (caret, anchor),
        caret_now(),
        "with the caret still where it was"
    );
}

/// A key that moved nothing draws nothing: `←` at the start of a query is an edit the road queues
/// like any other, and a repaint of this band copies the whole page frame under the card (see
/// `text_find::settle` and `text_find::BAR_PAINTS`).
#[test]
fn a_key_that_moved_nothing_asks_for_no_repaint() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_page_with_two_needles("a-bar-no-op-key.txt");

    text_find::open_text_find();
    type_into_the_bar("needle");

    let paints = || text_find::BAR_PAINTS.load(Ordering::Relaxed);
    let plain = text_find::Modifiers::default();
    let press = |vk: u16, held| {
        let edit =
            text_find::key_edit_with(vk as i32, held).expect("a key the field's one table answers");
        text_find::note(edit);
        unsafe { text_find::settle(a_pin_window()) };
    };

    // `Home` from the end of the query is a move, and a move asks for the band once.
    let before = paints();
    press(VK_HOME.0, plain);
    assert_eq!(
        caret_now(),
        (0, 0),
        "the caret is at the start of the query"
    );
    assert_eq!(
        paints(),
        before + 1,
        "a move that moved asked for the band again"
    );

    // `Home` again, `←` where there is nothing to the left, and a `Shift`+`Home` whose selection is
    // already at the start: three edits the road queues and none of them a change the card can draw.
    let shift = text_find::Modifiers {
        shift: true,
        control: false,
    };
    for (vk, held, name) in [
        (VK_HOME.0, plain, "Home"),
        (VK_LEFT.0, plain, "←"),
        (VK_HOME.0, shift, "Shift+Home"),
    ] {
        press(vk, held);
        assert_eq!(caret_now(), (0, 0), "{name} moved nothing");
        assert_eq!(
            paints(),
            before + 1,
            "{name} at the start of the query repainted nothing"
        );
    }

    // And a move after them is drawn, so the gate is the change and not the road.
    press(VK_RIGHT.0, plain);
    assert_eq!(caret_now(), (1, 1));
    assert_eq!(paints(), before + 2, "a move that moved is drawn");
}

/// The repaint gate follows the field's own state and not the scroll a paint derives from it: on a
/// query too wide for its field a caret move that scrolls the run is drawn, and a key that moved
/// nothing is not — even though two different carets can rest at the same scroll (see
/// `text_find::Drawn` and `pin_chrome::query_run`).
#[test]
fn a_caret_move_that_scrolls_a_long_query_is_drawn_and_one_that_moved_nothing_is_not() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_page_with_two_needles("a-bar-scrolled-query.txt");

    text_find::open_text_find();
    type_into_the_bar(
        "a query too long for the card's own field, so that its own beginning is cut away",
    );

    let paints = || text_find::BAR_PAINTS.load(Ordering::Relaxed);
    let plain = text_find::Modifiers::default();
    let press = |vk: u16| {
        let edit = text_find::key_edit_with(vk as i32, plain)
            .expect("a key the field's one table answers");
        text_find::note(edit);
        unsafe { text_find::settle(a_pin_window()) };
    };

    // `Home` off the end of the query is a move that scrolls the run to show its beginning, and what
    // the card draws changed: the run is not where it was.
    let before = paints();
    press(VK_HOME.0);
    assert_eq!(
        caret_now(),
        (0, 0),
        "the caret is at the start of the query"
    );
    assert_eq!(
        paints(),
        before + 1,
        "a move that scrolled the run to its beginning is drawn"
    );

    // `Home` again moved nothing the card draws, so it draws nothing.
    press(VK_HOME.0);
    assert_eq!(
        paints(),
        before + 1,
        "Home at the start of a long query repainted nothing"
    );

    // And `End` scrolls the run back to the caret's own end, which is drawn again.
    press(VK_END.0);
    assert_eq!(
        caret_now().0,
        bar_now().1.len(),
        "the caret is back at the end"
    );
    assert_eq!(
        paints(),
        before + 2,
        "a move that scrolled the run back is drawn"
    );
}

/// What one key lands as on the one road a focused pin gives the bar: the window procedure's
/// `WM_KEYDOWN` arm reads the key with `key_edit_with`, and where the key types a character the
/// `WM_CHAR` that follows it — which exists only because nothing swallowed the key-down — types that
/// character (see `window_proc` and `text_find::key_edit_with`). Answers the one edit the bar was
/// queued, or nothing for a key it only swallows.
fn the_window_is_handed(
    vk: u16,
    held: text_find::Modifiers,
    typed: Option<char>,
) -> Option<FindEdit> {
    clear_the_queue();

    let key = text_find::key_edit_with(vk as i32, held);
    if let Some(edit) = key {
        text_find::note(edit);
    }

    // A key that types a character answers nothing as a key. The character is the `WM_CHAR` road,
    // and a key-down that also queued an edit would type it twice.
    assert!(
        typed.is_none() || key.is_none(),
        "{vk:#06x} types a character and must not also be a key the field answers"
    );
    if let Some(character) = typed {
        text_find::note(FindEdit::Typed(character));
    }

    let queued: Vec<FindEdit> = CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| state.find.pending.clone())
        })
        .unwrap_or_default();
    assert!(
        queued.len() <= 1,
        "{vk:#06x} under {held:?} asked the bar for {queued:?}"
    );
    queued.first().copied()
}

/// **Every key the field needs reaches it by the one road a focused pin gives it**, and the table is
/// the point: each row is a key, the modifiers it is pressed under, the character `TranslateMessage`
/// makes of it where it types one, and the edit it must land as. A key that disappears between the
/// keyboard and the field — as `Space` and `T` did while the hook still swallowed them — fails its
/// row, and a key added to the hook's pin-key table without the bar standing down fails at the hook
/// (see `key_input::a_find_bar_that_is_up_passes_on_every_key_the_pin_answers`).
#[test]
fn every_key_the_field_needs_reaches_it_by_the_road_a_pin_gives_it() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_page_with_two_needles("a-bar-every-key.txt");

    text_find::open_text_find();
    type_into_the_bar("needle");

    let plain = text_find::Modifiers::default();
    let shift = text_find::Modifiers {
        shift: true,
        control: false,
    };
    let control = text_find::Modifiers {
        shift: false,
        control: true,
    };
    let both = text_find::Modifiers {
        shift: true,
        control: true,
    };
    let left = text_find_field::Motion::CharacterLeft;
    let right = text_find_field::Motion::CharacterRight;
    let word_left = text_find_field::Motion::WordLeft;
    let word_right = text_find_field::Motion::WordRight;
    let home = text_find_field::Motion::Home;
    let end = text_find_field::Motion::End;

    // The key, the modifiers, the character it types where it types one, and the edit it lands as.
    let rows: &[(u16, text_find::Modifiers, Option<char>, Option<FindEdit>)] = &[
        // The caret's keys.
        (VK_LEFT.0, plain, None, Some(FindEdit::Move(left, false))),
        (VK_LEFT.0, shift, None, Some(FindEdit::Move(left, true))),
        (
            VK_LEFT.0,
            control,
            None,
            Some(FindEdit::Move(word_left, false)),
        ),
        (VK_LEFT.0, both, None, Some(FindEdit::Move(word_left, true))),
        (VK_RIGHT.0, plain, None, Some(FindEdit::Move(right, false))),
        (VK_RIGHT.0, shift, None, Some(FindEdit::Move(right, true))),
        (
            VK_RIGHT.0,
            control,
            None,
            Some(FindEdit::Move(word_right, false)),
        ),
        (VK_HOME.0, plain, None, Some(FindEdit::Move(home, false))),
        (VK_HOME.0, shift, None, Some(FindEdit::Move(home, true))),
        (VK_HOME.0, control, None, Some(FindEdit::Move(home, false))),
        (VK_END.0, plain, None, Some(FindEdit::Move(end, false))),
        (VK_END.0, shift, None, Some(FindEdit::Move(end, true))),
        // The matches: the arrows, the page keys and `F3` step them.
        (VK_UP.0, plain, None, Some(FindEdit::Step(-1))),
        (VK_PRIOR.0, plain, None, Some(FindEdit::Step(-1))),
        (VK_DOWN.0, plain, None, Some(FindEdit::Step(1))),
        (VK_DOWN.0, control, None, Some(FindEdit::Step(1))),
        (VK_NEXT.0, plain, None, Some(FindEdit::Step(1))),
        (VK_F3.0, plain, None, Some(FindEdit::Step(1))),
        (VK_RETURN.0, plain, None, Some(FindEdit::Step(1))),
        (VK_RETURN.0, shift, None, Some(FindEdit::Step(-1))),
        (VK_ESCAPE.0, plain, None, Some(FindEdit::ClearOrClose)),
        // The edits, which are the field's own keys.
        (VK_BACK.0, plain, None, Some(FindEdit::Backspace)),
        (VK_BACK.0, control, None, Some(FindEdit::DeleteWordLeft)),
        (VK_DELETE.0, plain, None, Some(FindEdit::DeleteForward)),
        (VK_DELETE.0, control, None, Some(FindEdit::DeleteWordRight)),
        (VK_A.0, control, None, Some(FindEdit::SelectAll)),
        (VK_C.0, control, None, Some(FindEdit::Copy)),
        (VK_X.0, control, None, Some(FindEdit::Cut)),
        (VK_V.0, control, None, Some(FindEdit::Paste)),
        (VK_Z.0, control, None, Some(FindEdit::Undo)),
        (VK_Z.0, both, None, Some(FindEdit::Redo)),
        (VK_Y.0, control, None, Some(FindEdit::Redo)),
        // And the characters: a letter, a space, and both cases of the `t` a query was unable to
        // hold while the hook ate its key-down. Each arrives as `WM_CHAR`, never as its key.
        (b'a' as u16, plain, Some('a'), Some(FindEdit::Typed('a'))),
        (VK_SPACE.0, plain, Some(' '), Some(FindEdit::Typed(' '))),
        (VK_T.0, plain, Some('t'), Some(FindEdit::Typed('t'))),
        (VK_T.0, shift, Some('T'), Some(FindEdit::Typed('T'))),
        // And the one key the field only swallows.
        (VK_TAB.0, plain, None, None),
    ];

    for (vk, held, typed, lands) in rows {
        assert_eq!(
            the_window_is_handed(*vk, *held, *typed),
            *lands,
            "{vk:#06x} under {held:?} must reach the field as {lands:?}"
        );
    }

    // And the query really holds them: a space and both cases of `t` are handed to the window's own
    // procedure as the key-down and the `WM_CHAR` that follows it, which is the pair
    // `TranslateMessage` makes of a key nothing swallowed.
    clear_the_queue();
    unsafe { text_find::settle(a_pin_window()) };
    let query_was = bar_now().1;
    assert_eq!(query_was, "needle");
    for (vk, character) in [(VK_SPACE.0, ' '), (VK_T.0, 't'), (VK_T.0, 'T')] {
        let answered =
            unsafe { window_proc(a_pin_window(), WM_KEYDOWN, WPARAM(vk as usize), LPARAM(0)) };
        assert_eq!(
            answered.0, 0,
            "{vk:#06x} is the bar's, and the bar swallows it after the field has read it"
        );
        unsafe {
            window_proc(
                a_pin_window(),
                WM_CHAR,
                WPARAM(character as usize),
                LPARAM(0),
            )
        };
    }
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(
        bar_now().1,
        format!("{query_was} tT"),
        "a space and both cases of `t` are typed into the query"
    );
}

/// A character is the bar's when it is a printable one and nothing else: the two half-characters of a
/// surrogate pair are not a character this bar holds, and the control characters a key has already
/// answered as a key are not typed a second time (see `text_find::typed_character`).
#[test]
fn a_character_is_typed_only_when_it_is_a_printable_one_of_the_basic_plane() {
    assert_eq!(text_find::typed_character('a' as usize), Some('a'));
    assert_eq!(text_find::typed_character('é' as usize), Some('é'));
    assert_eq!(
        text_find::typed_character(8),
        None,
        "a backspace is a key, not a character"
    );
    assert_eq!(text_find::typed_character(13), None, "and so is a return");
    assert_eq!(
        text_find::typed_character(0xD83D),
        None,
        "a high surrogate is half of a character the bar does not hold"
    );
    assert_eq!(
        text_find::typed_character(0xDE00),
        None,
        "and so is a low one"
    );
}

/// `Find Next` with the bar up steps the bar's query rather than the page's selection: the selection
/// is the reader's own and is left exactly where it was, and what moved is the match the bar is on
/// (see `find_next_in_text_preview`).
#[test]
fn a_find_next_with_the_bar_up_steps_its_query_rather_than_the_selection() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_page_with_two_needles("a-bar-find-next.txt");

    text_find::open_text_find();
    type_into_the_bar("needle");

    let selection = text_preview::Selection {
        anchor: (0, 0),
        caret: (0, 4),
    };
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        if let Some(state) = media.as_mut().and_then(|media| media.text_state.as_mut()) {
            state.selection = Some(selection);
        }
    }

    unsafe { find_next_in_text_preview(a_pin_window()) };
    unsafe { text_find::settle(a_pin_window()) };

    assert_eq!(bar_now().3, Some(1), "the bar moved to its next match");
    let still_selected = CURRENT_MEDIA.lock().ok().and_then(|media| {
        media
            .as_ref()
            .and_then(|media| media.text_state.as_ref())
            .and_then(|state| state.selection)
    });
    assert_eq!(
        still_selected,
        Some(selection),
        "and the page's own selection did not move with it"
    );
}

/// The bar belongs to the file it was opened on: a walk onto another document brings a page with a
/// bar that is closed and empty, because a match is a place in a document the reader is no longer
/// reading (see `keep_text_place` and `FindState`).
#[test]
fn another_file_does_not_inherit_the_bar_a_page_was_searched_with() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_page_with_two_needles("a-bar-left-behind.txt");

    text_find::open_text_find();
    type_into_the_bar("needle");
    assert!(
        bar_now().0,
        "the page being walked away from has the bar up"
    );

    let Some(next) = a_page_of("another page\naltogether\n", "a-bar-moved-to.txt", 240, 60) else {
        panic!("a page of text was not loaded")
    };
    let carried = keep_text_place(next);

    let state = carried
        .text_state
        .as_ref()
        .expect("a page of text keeps its state");
    assert!(
        !state.find.open,
        "the next file's page has no bar standing over it"
    );
    assert_eq!(
        state.find.field.text(),
        "",
        "and no query it never searched"
    );
    assert!(
        state.find.matches.is_empty(),
        "and no places in a document it is not showing"
    );
}

/// The card a bar is drawn in is the card its controls answer in, and every one of them asks the thing
/// it is drawn as: a switch turns and searches again under it, the two arrows step the matches, the
/// cross closes the bar, and the card's own face swallows a press without asking anything — nothing
/// under the floating card was what the hand meant (see `text_find::press` and
/// `pin_chrome::FindBar::part_at`).
#[test]
fn a_press_on_the_card_answers_its_own_controls() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    put_page("Cat cat CAT\nscat cat\n", "a-bar-pressed-on.txt", 240, 80);

    text_find::open_text_find();
    type_into_the_bar("cat");
    assert_eq!(
        bar_now().2,
        5,
        "every case of the letters is found to begin with"
    );

    let card = text_find::find_bar()
        .expect("the card the bar is drawn in")
        .card;
    let middle = |rect: RECT| ((rect.left + rect.right) / 2, (rect.top + rect.bottom) / 2);
    let press = |rect: RECT| {
        let (x, y) = middle(rect);
        text_find::press(x, y)
    };

    assert!(press(card.case), "a press on a switch is the card's");
    unsafe { text_find::settle(a_pin_window()) };
    assert!(bar_now().4, "the case switch is on");
    assert_eq!(bar_now().2, 3, "and the search ran again under it");

    assert!(press(card.case));
    unsafe { text_find::settle(a_pin_window()) };
    assert!(!bar_now().4, "the same press turns it off again");
    assert_eq!(bar_now().2, 5, "and every case is found again");

    assert!(press(card.whole_word));
    unsafe { text_find::settle(a_pin_window()) };
    assert!(bar_now().5, "the whole-word switch is on");
    assert_eq!(
        bar_now().2,
        4,
        "and the one inside `scat` is no longer a word"
    );

    assert!(press(card.next));
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(bar_now().3, Some(1), "the arrow on the right steps on");
    assert!(press(card.previous));
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(bar_now().3, Some(0), "and the one on the left steps back");

    let (x, y) = (card.query.left + 1, card.query.top + 1);
    let before = bar_now();
    assert!(
        text_find::press(x, y),
        "the card's own face is the card's too"
    );
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(
        bar_now(),
        before,
        "and a press there asks nothing of the bar"
    );

    assert!(
        !text_find::press(card.panel.left - 2, card.panel.top + 2),
        "a press beside the card is the page's"
    );

    assert!(press(card.close));
    unsafe { text_find::settle(a_pin_window()) };
    assert!(!bar_now().0, "the cross puts the bar away");
    assert!(
        text_find::find_bar().is_none(),
        "and there is no card left to press"
    );
}

/// Whether the bar is holding a repeat: the press's own step is written down before the hold is armed,
/// so this is only ever the one after it (see `Hold`).
fn repeating() -> bool {
    CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| state.find.hold.is_some())
        })
        .unwrap_or(false)
}

/// Forty lines with a needle on two of them, in a box tall enough to hold the card over it: the card
/// is laid out only on a page whose corner has room for its two rows (see `find_bar_layout`).
fn a_tall_page_with_two_needles(name: &str) {
    let text: String = (0..40)
        .map(|n| {
            format!(
                "line {n}{}
",
                if n == 5 || n == 19 {
                    " needle here"
                } else {
                    ""
                }
            )
        })
        .collect();

    put_page(&text, name, 240, 140);
}

/// The middle of one of the card's boxes, in the window's own coordinates.
fn middle_of(rect: RECT) -> (i32, i32) {
    ((rect.left + rect.right) / 2, (rect.top + rect.bottom) / 2)
}

/// The two chevrons are buttons a hand holds down rather than presses once: the press is a step of its
/// own, and the hold steps again on the tick's clock — first after a delay and then once an interval —
/// until the button is let go of (see `text_find::press` and `text_find::step_the_hold`).
#[test]
fn a_held_chevron_steps_once_at_the_press_and_then_on_the_ticks_clock() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_tall_page_with_two_needles("a-bar-held.txt");

    text_find::open_text_find();
    type_into_the_bar("needle");
    assert_eq!(
        bar_now().3,
        Some(0),
        "the reader starts on the first of the two needles"
    );

    let card = text_find::find_bar()
        .expect("the card the bar is drawn in")
        .card;
    let pressed = Instant::now();
    let (x, y) = middle_of(card.next);
    assert!(text_find::press(x, y), "a press on the arrow is the card's");
    assert!(repeating(), "and a chevron pressed arms the hold behind it");

    // The press is a step of its own: the tick that settles it is the only thing that has to run.
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(bar_now().3, Some(1), "the press itself steps once");

    // Nothing more before the delay is up: a hand that lets go inside it has stepped once.
    text_find::step_the_hold(pressed + Duration::from_millis(text_find::HOLD_DELAY_MS - 1));
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(
        bar_now().3,
        Some(1),
        "and the hold steps nothing before its delay"
    );

    // Then one step per interval.
    text_find::step_the_hold(pressed + Duration::from_millis(text_find::HOLD_DELAY_MS + 20));
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(
        bar_now().3,
        Some(0),
        "the first repeat comes a delay after the press"
    );

    text_find::step_the_hold(pressed + Duration::from_millis(text_find::HOLD_DELAY_MS + 40));
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(
        bar_now().3,
        Some(0),
        "and the next one an interval after that, not on every tick"
    );

    text_find::step_the_hold(
        pressed + Duration::from_millis(text_find::HOLD_DELAY_MS + text_find::HOLD_REPEAT_MS + 20),
    );
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(
        bar_now().3,
        Some(1),
        "one interval further on, one step more"
    );
}

/// A release stops the hold where it got to, and so does the key that puts the bar away: a chevron is
/// stepping only while it is held (see `text_find::release` and `text_find::step_the_hold`).
#[test]
fn a_release_or_an_escape_stops_a_held_chevron() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_tall_page_with_two_needles("a-bar-let-go-of.txt");

    text_find::open_text_find();
    type_into_the_bar("needle");

    let card = text_find::find_bar()
        .expect("the card the bar is drawn in")
        .card;
    let pressed = Instant::now();
    let (x, y) = middle_of(card.next);
    assert!(text_find::press(x, y));
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(bar_now().3, Some(1), "the press stepped once");

    assert!(text_find::release(), "the card was holding that press");
    assert!(!repeating(), "and the release let the hold go with it");
    text_find::step_the_hold(pressed + Duration::from_millis(text_find::HOLD_DELAY_MS + 400));
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(
        bar_now().3,
        Some(1),
        "a chevron that is no longer held steps no further"
    );

    // The two arrows are a held button's; the key that puts the bar away is not.
    assert!(text_find::press(x, y));
    assert!(repeating(), "a press arms the hold again");
    assert_eq!(
        text_find::key_edit(VK_ESCAPE.0 as i32),
        Some(FindEdit::ClearOrClose)
    );
    // The first `Esc` empties the query the bar is holding, and the second puts the bar away (see
    // `FindEdit::ClearOrClose`).
    text_find::note(FindEdit::ClearOrClose);
    unsafe { text_find::settle(a_pin_window()) };
    assert!(bar_now().0, "Escape on a query with text leaves the bar up");
    assert_eq!(bar_now().1, "", "and empties the query");
    assert!(
        repeating(),
        "while the button the hand is still on goes on stepping"
    );

    text_find::note(FindEdit::ClearOrClose);
    unsafe { text_find::settle(a_pin_window()) };
    assert!(!bar_now().0, "and an Escape on an empty query puts it away");
    assert!(!repeating(), "the hold goes with the bar");
}

/// The switches and the cross are one-shot: a press on `Aa` turns it and a press on the cross puts the
/// bar away, and neither keeps stepping the matches while it is held (see `text_find::press`).
#[test]
fn the_switches_and_the_cross_are_one_shot_and_arm_no_hold() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_tall_page_with_two_needles("a-bar-one-shot.txt");

    text_find::open_text_find();
    type_into_the_bar("needle");

    let card = text_find::find_bar()
        .expect("the card the bar is drawn in")
        .card;
    let held = Instant::now() + Duration::from_millis(text_find::HOLD_DELAY_MS + 1000);
    let current = bar_now().3;

    for rect in [card.case, card.whole_word] {
        let (x, y) = middle_of(rect);
        assert!(text_find::press(x, y));
        assert!(!repeating(), "a switch pressed arms no hold");
        text_find::step_the_hold(held);
        unsafe { text_find::settle(a_pin_window()) };
        text_find::release();
    }
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(
        bar_now().3,
        current,
        "neither switch walks the matches however long it was held"
    );

    let (x, y) = middle_of(card.close);
    assert!(text_find::press(x, y));
    assert!(!repeating(), "and the cross arms none either");
    text_find::step_the_hold(held);
    unsafe { text_find::settle(a_pin_window()) };
    assert!(!bar_now().0, "the cross puts the bar away");
}

/// A box change keeps the bar: the page is the same document laid out again, so its query is still
/// searched for, its places are still the document's places, and the card is still over it — only the
/// page under the card is wrapped at another width (see `relayout_text_for_the_box` and `FindState`).
#[test]
fn a_box_change_keeps_the_bar_and_the_places_it_found() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    put_page(
        "needle here\nnothing\nneedle again\n",
        "a-bar-box-changed.txt",
        240,
        80,
    );

    text_find::open_text_find();
    type_into_the_bar("needle");
    let before = bar_now();
    assert_eq!(before.2, 2, "both needles are found in the first box");

    relayout_text_for_the_box((0, 0, 180, 120));

    let after = bar_now();
    assert!(after.0, "the bar is still up over the page");
    assert_eq!(after.1, before.1, "at the query it was typed with");
    assert_eq!(
        (after.2, after.4, after.5),
        (before.2, before.4, before.5),
        "and its places and switches are places in the document, which did not move"
    );
    assert!(
        CURRENT_MEDIA
            .lock()
            .ok()
            .and_then(|media| {
                media
                    .as_ref()
                    .and_then(|media| media.text_state.as_ref())
                    .map(|state| (state.width, state.height, state.find.highlights().is_some()))
            })
            .is_some_and(|(width, height, drawn)| width == 180 && height == 120 && drawn),
        "the page is laid out for the new box and the frame is drawn with the places on it"
    );
}

/// `Esc` is the field's clearing key first and the bar's closing key second: a query with text is
/// emptied and the bar left standing over a page with no highlights on it, and the same key on an
/// empty query is the bar going away (see `FindEdit::ClearOrClose` and `text_find::settle`).
#[test]
fn escape_empties_the_query_before_it_closes_the_bar() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_page_with_two_needles("a-bar-escape.txt");

    text_find::open_text_find();
    type_into_the_bar("needle");
    assert_eq!(bar_now().2, 2, "the query is found twice");

    let esc = text_find::key_edit(VK_ESCAPE.0 as i32);
    assert_eq!(
        esc,
        Some(FindEdit::ClearOrClose),
        "Escape is the bar's own key"
    );
    text_find::note(FindEdit::ClearOrClose);
    unsafe { text_find::settle(a_pin_window()) };

    let (open, query, count, current, ..) = bar_now();
    assert!(open, "the bar is still up over the page");
    assert_eq!(query, "", "and the query it was holding is gone");
    assert_eq!(
        (count, current),
        (0, None),
        "with nothing left to highlight"
    );

    text_find::note(FindEdit::ClearOrClose);
    unsafe { text_find::settle(a_pin_window()) };
    assert!(
        !bar_now().0,
        "and Escape on an empty query puts the bar away"
    );
}

/// `Ctrl`+`F` over a bar that is already up selects the whole query instead of appending to it: the
/// next character typed replaces what is there, which is the find field a reader has used before
/// (see `text_find::open_text_find` and `Field::select_all`).
#[test]
fn a_ctrl_f_over_an_open_bar_selects_the_query_rather_than_appending_to_it() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_page_with_two_needles("a-bar-found-again.txt");

    text_find::open_text_find();
    type_into_the_bar("needle");

    text_find::open_text_find();
    unsafe { text_find::settle(a_pin_window()) };
    let (selected, text) = CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| {
                    (
                        state.find.field.selection(),
                        state.find.field.text().to_string(),
                    )
                })
        })
        .expect("a pinned page keeps a find bar");
    assert_eq!(text, "needle", "the query is the one it already held");
    assert_eq!(
        selected,
        Some((0, 6)),
        "and the whole of it is selected for what is typed next"
    );

    text_find::note(FindEdit::Typed('x'));
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(bar_now().1, "x", "so the next character replaces it");
}

/// The keys that step the matches keep stepping them with the field under the caret: `↓`, `Enter`
/// and `F3` go on and `↑` comes back, and none of them moves the caret or the query (see
/// `text_find::key_edit`).
#[test]
fn the_result_keys_still_step_with_the_field_in_hand() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_page_with_two_needles("a-bar-field-steps.txt");

    text_find::open_text_find();
    type_into_the_bar("needle");
    assert_eq!(
        bar_now().3,
        Some(0),
        "the reader starts on the first needle"
    );

    for (vk, step, expected) in [
        (VK_DOWN, 1, Some(1)),
        (VK_F3, 1, Some(0)),
        (VK_RETURN, 1, Some(1)),
        (VK_UP, -1, Some(0)),
    ] {
        assert_eq!(
            text_find::key_edit(vk.0 as i32),
            Some(FindEdit::Step(step)),
            "{vk:?} is a step through the matches"
        );
        text_find::note(FindEdit::Step(step));
        unsafe { text_find::settle(a_pin_window()) };
        assert_eq!(bar_now().3, expected, "{vk:?} stepped to {expected:?}");
    }

    let (open, query, ..) = bar_now();
    assert!(open, "the bar is still up through all of it");
    assert_eq!(query, "needle", "and the query is untouched");
}

/// A press in the query places the caret, a drag from it selects, and a second press of the same
/// place takes the word under the pointer — the page's own press menu, over the field (see
/// `text_find::press_the_field` and `text_find::drag`).
#[test]
fn a_press_in_the_field_places_the_caret_and_a_double_press_takes_its_word() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    // A page with room for the card: a bar laid out over a corner too short for it draws nothing,
    // and a card with nothing drawn has no field to press.
    a_tall_page_with_two_needles("a-bar-field-pressed.txt");

    text_find::open_text_find();
    type_into_the_bar("one two");

    let field_now = || {
        CURRENT_MEDIA
            .lock()
            .ok()
            .and_then(|media| {
                media
                    .as_ref()
                    .and_then(|media| media.text_state.as_ref())
                    .map(|state| (state.find.field.caret(), state.find.field.selection()))
            })
            .expect("a pinned page keeps a find bar")
    };

    let card = text_find::find_bar()
        .expect("the card the bar is drawn in")
        .card;
    let (left, y) = (card.query.left, middle_of(card.query).1);

    assert!(text_find::press(left, y), "the query's row is the card's");
    assert_eq!(
        field_now().0,
        0,
        "a press at the field's left edge is its beginning"
    );
    assert!(
        text_find::a_control_is_held(),
        "and it takes the pointer for the drag that may follow"
    );
    assert!(
        !text_find::drag(card.query.left - 20),
        "a drag that stays on the same character moves nothing"
    );
    assert_eq!(field_now(), (0, None));
    assert!(text_find::release(), "the release ends the drag");

    // A drag from the beginning past the end selects the whole query, and the pointer is free again.
    assert!(text_find::press(left, y));
    assert!(
        text_find::drag(card.query.right + 20),
        "a drag to another place grows the selection"
    );
    assert_eq!(field_now().1, Some((0, 7)), "the whole query is selected");
    assert!(text_find::release());
    assert!(!text_find::a_control_is_held());

    // Two presses of the same place are one double press by the time and box Windows counts them
    // in, and the second takes the word the caret landed in.
    assert!(text_find::press(left, y));
    assert!(text_find::press(left, y));
    assert_eq!(
        field_now().1,
        Some((0, 3)),
        "the word under the pointer, whole"
    );
    assert!(text_find::release());
}

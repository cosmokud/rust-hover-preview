//! The find bar over a pinned page: what it holds, what a key asks of it, and where its query is
//! found (see `text_find`).
//!
//! Nothing here draws the bar — the card over the page is the chrome's, and its pixels are asked
//! about with the chrome's own tests. What is asked here is the state machine a key and a tick share:
//! the query, the count, the match the reader is on, and the page that does or does not move.

use super::text_selection::{a_page_of, PinPutBack};
use super::*;

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
                        state.find.query.clone(),
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

    text_find::note(FindEdit::Deleted);
    text_find::note(FindEdit::Deleted);
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

/// The keys a bar names are the bar's, and the ones it does not are swallowed rather than handed to
/// the page: the page keys, the page's two ends, a Space, a subtitle step and the pin's own walk are
/// all inert under a query, while the two arrows step the matches (see `text_find::take_hooked_key`
/// and `key_edit`).
#[test]
fn the_page_keys_and_the_pins_walk_are_inert_while_the_bar_is_up() {
    let _room = the_room_to_itself();
    let _pin = PinPutBack::take();
    stand_pin(Some(PinnedPreview::for_test()));
    a_page_with_two_needles("a-bar-inert-keys.txt");

    text_find::open_text_find();
    type_into_the_bar("needle");

    let (_, query, count, current, ..) = bar_now();
    let standing = (query, count, current, first_line_now());

    for command in [
        PinCommand::Previous,
        PinCommand::Next,
        PinCommand::PageTop,
        PinCommand::PageBottom,
        PinCommand::TogglePlayback,
        PinCommand::NextSubtitle,
    ] {
        assert!(
            text_find::take_hooked_key(command),
            "{command:?} is a key the bar swallowed whole"
        );
    }
    unsafe { text_find::settle(a_pin_window()) };

    let (_, query, count, current, ..) = bar_now();
    assert_eq!(
        (query, count, current, first_line_now()),
        standing,
        "nothing the page answers moved under the query: not its place, not its matches"
    );

    // The two arrows are the bar's, as the page keys of their direction.
    assert!(text_find::take_hooked_key(PinCommand::PageOn));
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(
        bar_now().3,
        Some(1),
        "the down arrow steps to the next match"
    );

    assert!(text_find::take_hooked_key(PinCommand::PageBack));
    unsafe { text_find::settle(a_pin_window()) };
    assert_eq!(
        bar_now().3,
        Some(0),
        "and the up arrow back to the one before it"
    );

    // And the keys a message would bring rather than the hook: `Esc`, the return, the backspace,
    // the two arrows and `F3` are named; the page keys and the page's ends are only swallowed.
    assert_eq!(
        text_find::key_edit(VK_ESCAPE.0 as i32),
        Some(FindEdit::Close)
    );
    assert_eq!(
        text_find::key_edit(VK_DOWN.0 as i32),
        Some(FindEdit::Step(1))
    );
    assert_eq!(
        text_find::key_edit(VK_UP.0 as i32),
        Some(FindEdit::Step(-1))
    );
    assert_eq!(
        text_find::key_edit(VK_BACK.0 as i32),
        Some(FindEdit::Deleted)
    );
    assert_eq!(text_find::key_edit(VK_F3.0 as i32), Some(FindEdit::Step(1)));
    assert!(
        matches!(
            text_find::key_edit(VK_RETURN.0 as i32),
            Some(FindEdit::Step(_))
        ),
        "a return steps, one way or the other under Shift"
    );
    for vk in [
        VK_PRIOR, VK_NEXT, VK_HOME, VK_END, VK_SPACE, VK_LEFT, VK_RIGHT,
    ] {
        assert_eq!(
            text_find::key_edit(vk.0 as i32),
            None,
            "{vk:?} is a page key the bar only swallows"
        );
    }

    // A bar that is away leaves the pin's keys to the pin.
    text_find::note(FindEdit::Close);
    unsafe { text_find::settle(a_pin_window()) };
    assert!(!text_find::take_hooked_key(PinCommand::PageOn));
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
    assert_eq!(state.find.query, "", "and no query it never searched");
    assert!(
        state.find.matches.is_empty(),
        "and no places in a document it is not showing"
    );
}

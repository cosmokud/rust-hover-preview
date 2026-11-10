//! The find bar's query field: typing at a caret, the selection, the words a `Ctrl` chord takes, the
//! clipboard's one line and the two stacks undo and redo are kept on (see `text_find_field`).
//!
//! Everything here is asked of a field with no window behind it. What a press means and where the
//! caret is drawn are the card's questions, and they are asked where the card's columns are measured
//! (see `text::pin_chrome`).

use super::*;
use text_find_field::{sanitize_pasted, CaretBlink, Field, Motion, QUERY_LIMIT};

/// A field holding `text`, with the caret after its last character.
fn holding(text: &str) -> Field {
    Field::holding(text)
}

/// A `Ctrl`+arrow's unit: a word is the page's own run — letters, digits and the underscore — and the
/// marks between two words are crossed with it rather than counted as words of their own.
#[test]
fn a_word_move_crosses_the_marks_between_one_word_and_the_next() {
    let mut field = holding("one_two one-two");

    field.move_caret(Motion::Home, false);
    field.move_caret(Motion::WordRight, false);
    assert_eq!(
        field.caret(),
        8,
        "the first word takes its underscore and stops before the gap"
    );

    field.move_caret(Motion::WordRight, false);
    assert_eq!(
        field.caret(),
        12,
        "and the gap is crossed with the word after it, not stopped in"
    );

    field.move_caret(Motion::WordRight, false);
    assert_eq!(
        field.caret(),
        15,
        "the whole of `two` is one word, to the end"
    );

    field.move_caret(Motion::WordLeft, false);
    assert_eq!(
        field.caret(),
        12,
        "and back across it, stepping into the word"
    );
    field.move_caret(Motion::WordLeft, false);
    assert_eq!(
        field.caret(),
        8,
        "then across the mark to the word before it"
    );
    field.move_caret(Motion::WordLeft, false);
    assert_eq!(
        field.caret(),
        0,
        "and the first word with its underscore whole"
    );
}

/// **A double press takes the word it landed on, and a second double press on that same word takes the
/// whole query** — the line an editor widens a word's selection to, and a query is one line. A press on
/// another word after that is that word again rather than the line: the widening belongs to the word
/// the hand has already taken (see `Field::select_word`).
#[test]
fn a_double_press_takes_a_word_and_the_same_word_again_takes_the_line() {
    let mut field = holding("cute cat");

    field.select_word(6);
    assert_eq!(
        field.selection(),
        Some((5, 8)),
        "the word the place stands in, and nothing of the one beside it"
    );
    assert_eq!(field.selected_text(), "cat");

    field.select_word(6);
    assert_eq!(
        field.selection(),
        Some((0, 8)),
        "the same word again is the whole of the query"
    );
    assert_eq!(field.selected_text(), "cute cat");

    field.select_word(1);
    assert_eq!(
        field.selected_text(),
        "cute",
        "another word is another word"
    );
    field.select_word(1);
    assert_eq!(
        field.selected_text(),
        "cute cat",
        "and its own second press widens that one"
    );
}

/// A word is the page's own run — letters, digits and the underscore — so a gap between two words is a
/// run of its own, and every place a word is picked by is a character boundary rather than a byte
/// (see `text_preview::run_at` and `Field::word_at`).
#[test]
fn a_word_is_the_pages_run_and_never_half_a_character() {
    let mut field = holding("cute cat");

    field.select_word(4);
    assert_eq!(
        field.selected_text(),
        " ",
        "the gap between two words is a run a press can land in"
    );

    let mut accented = holding("héllo wörld");
    accented.select_word(2);
    let (start, end) = accented.selection().expect("the word the place stands in");
    assert_eq!(
        accented.selected_text(),
        "héllo",
        "a character of more than one byte is one character of the word"
    );
    assert!(
        accented.text().is_char_boundary(start) && accented.text().is_char_boundary(end),
        "and neither end of the range cuts a character in half"
    );
}

/// A typed character lands at the caret, and a selection replaces rather than what it is
/// appended to — including a selection that reaches over the caret's own end (see `Field::insert`).
#[test]
fn a_typed_character_lands_at_the_caret_and_replaces_a_selection() {
    let mut field = holding("cd");
    field.place_caret(1, false);
    field.insert('b');
    assert_eq!(field.text(), "cbd");
    assert_eq!(
        (field.caret(), field.anchor()),
        (2, 2),
        "the caret moves past it"
    );

    field.select_all();
    field.insert('x');
    assert_eq!(field.text(), "x", "the selected query is replaced whole");
    assert_eq!((field.caret(), field.anchor()), (1, 1));

    // The query is capped in characters rather than in bytes, so a character outside the basic plane
    // counts once and is never cut in half.
    let mut full = Field::holding(&"é".repeat(QUERY_LIMIT - 1));
    full.insert('ß');
    full.insert('z');
    assert_eq!(full.text().chars().count(), QUERY_LIMIT);
    assert!(
        full.text().ends_with('ß'),
        "the character that fits is whole"
    );
}

/// A backspace and a delete take the selection, or one character, and do nothing at the edge they are
/// pointed at: no wrap from the beginning of the query to its end, and no panic at either.
#[test]
fn a_backspace_and_a_delete_take_one_character_or_the_selection() {
    let mut field = holding("ab");
    field.backspace();
    assert_eq!(
        field.text(),
        "a",
        "a backspace takes the character before the caret"
    );

    field.place_caret(0, false);
    field.backspace();
    assert_eq!(
        field.text(),
        "a",
        "and nothing at the beginning of the query"
    );

    field.place_caret(1, false);
    field.delete();
    assert_eq!(field.text(), "a", "a delete at the end takes nothing");

    field.place_caret(0, false);
    field.delete();
    assert_eq!(
        field.text(),
        "",
        "and a delete before the last character takes it"
    );

    // A selection is what either key takes first, whichever side of the anchor the caret is on.
    let mut field = holding("one two");
    field.place_caret(7, false);
    field.place_caret(4, true);
    assert_eq!(field.selection(), Some((4, 7)));
    field.backspace();
    assert_eq!(field.text(), "one ");
    assert_eq!(field.caret(), 4);

    let mut field = holding("one two");
    field.select_all();
    field.delete();
    assert_eq!(
        field.text(),
        "",
        "the whole query goes with the caret at its start"
    );
    assert_eq!(field.caret(), 0);
}

/// `Ctrl`+`Backspace` and `Ctrl`+`Delete` take a word, by the same run rule a double click takes it
/// by, and they take the selection first where there is one.
#[test]
fn the_word_deletes_take_the_run_beside_the_caret_and_nothing_at_the_edges() {
    let mut field = holding("one two three");

    field.delete_word_left();
    assert_eq!(field.text(), "one two ");
    assert_eq!(field.caret(), 8, "the caret stands where the word began");

    field.delete_word_left();
    assert_eq!(field.text(), "one ", "the gap before it goes with the word");

    field.place_caret(0, false);
    field.delete_word_left();
    assert_eq!(
        field.text(),
        "one ",
        "nothing before the beginning of the query"
    );

    field.delete_word_right();
    assert_eq!(field.text(), "", "and the word after a caret at its start");

    let mut field = holding("one two");
    field.place_caret(4, false);
    field.delete_word_right();
    assert_eq!(
        field.text(),
        "one ",
        "a forward delete takes the word it stands on"
    );

    let mut field = holding("one two");
    field.place_caret(7, false);
    field.delete_word_right();
    assert_eq!(
        field.text(),
        "one two",
        "and nothing at the end of the query"
    );

    let mut field = holding("one two");
    field.place_caret(0, false);
    field.place_caret(3, true);
    field.delete_word_right();
    assert_eq!(
        field.text(),
        " two",
        "the selection is what the key takes first"
    );
}

/// The caret moves by a character and by a word, and `Shift` is what makes any of it a selection
/// rather than a move: a move without it collapses the range the caret was one end of.
#[test]
fn a_caret_move_extends_with_shift_and_collapses_without_it() {
    let mut field = holding("one two");

    field.move_caret(Motion::Home, false);
    field.move_caret(Motion::CharacterRight, true);
    field.move_caret(Motion::CharacterRight, true);
    assert_eq!(field.selection(), Some((0, 2)));
    assert_eq!(field.caret(), 2, "the caret is the focus end");

    field.move_caret(Motion::WordRight, true);
    assert_eq!(
        field.selection(),
        Some((0, 4)),
        "and word-wise extends the same range"
    );

    field.move_caret(Motion::CharacterLeft, false);
    assert_eq!(
        field.selection(),
        None,
        "a move without Shift collapses the selection"
    );
    assert_eq!((field.caret(), field.anchor()), (3, 3));

    field.move_caret(Motion::End, false);
    assert_eq!(field.caret(), 7, "End is the query's end");
    field.move_caret(Motion::Home, true);
    assert_eq!(
        field.selection(),
        Some((0, 7)),
        "and Shift+Home takes all of it"
    );
}

/// Select All then a character is a query of that one character: the whole of the old query is what
/// the typed run replaces, and one undo puts it back with the selection it had (see `Field::undo`).
#[test]
fn select_all_then_typing_replaces_the_query_and_undo_puts_it_back() {
    let mut field = holding("needle");
    field.select_all();
    assert_eq!(field.selection(), Some((0, 6)));

    field.insert('x');
    assert_eq!(field.text(), "x");

    assert!(field.undo());
    assert_eq!(field.text(), "needle");
    assert_eq!(
        field.selection(),
        Some((0, 6)),
        "undo puts the selection back as well as the query"
    );
}

/// A copy takes the selection and changes nothing; a cut takes it and leaves the query it was taken
/// out of, with the caret where the selection began.
#[test]
fn a_copy_takes_the_selection_and_a_cut_takes_it_out() {
    let mut field = holding("one two");
    field.place_caret(0, false);
    field.place_caret(3, true);

    assert_eq!(field.selected_text(), "one");
    assert_eq!(field.cut(), "one");
    assert_eq!(field.text(), " two");
    assert_eq!(field.caret(), 0);

    field.place_caret(0, false);
    field.place_caret(4, true);
    assert_eq!(field.selected_text(), " two");
    assert_eq!(field.text(), " two", "a copy leaves the query as it was");
}

/// A paste is one line of text, and what a paste holds that a query cannot is what the sanitizer
/// takes out: the breaks become spaces, the controls and the byte-order mark go, a run of spaces is
/// one, and the result is capped in characters (see `sanitize_pasted`).
#[test]
fn a_paste_is_made_one_line_before_it_reaches_the_query() {
    assert_eq!(
        sanitize_pasted("one\r\ntwo\tthree"),
        "one two three",
        "the three breaks a selection can carry are spaces"
    );
    assert_eq!(
        sanitize_pasted("\u{feff}needle"),
        "needle",
        "a byte-order mark is not part of the query it introduces"
    );
    assert_eq!(
        sanitize_pasted("a\u{0007}b\u{001b}c"),
        "abc",
        "a bell and an escape are dropped rather than turned into marks"
    );
    assert_eq!(
        sanitize_pasted("  one      two  "),
        "one two",
        "a run of whitespace is one space, and one at either end is none"
    );
    assert_eq!(sanitize_pasted("\n\t \r\n"), "");

    let long: String = "x".repeat(QUERY_LIMIT + 40);
    assert_eq!(
        sanitize_pasted(&long).chars().count(),
        QUERY_LIMIT,
        "a pasted document is capped at a query"
    );

    // And a paste into a field that already holds text takes only the room it has left, whole
    // characters at a time.
    let mut field = Field::holding(&"é".repeat(QUERY_LIMIT - 2));
    field.paste("abcd");
    assert_eq!(field.text().chars().count(), QUERY_LIMIT);
    assert!(field.text().ends_with("ab"));
}

/// Undo puts back the whole field and redo puts back what the undo took away, and consecutive typed
/// characters are one entry: a caret jump, a delete, a paste or a selection begins a new one (see
/// `Field::snapshot`).
#[test]
fn consecutive_typing_is_one_undo_and_a_caret_jump_begins_another() {
    let mut field = Field::default();
    for character in "abc".chars() {
        field.insert(character);
    }
    assert_eq!(field.text(), "abc");

    assert!(field.undo());
    assert_eq!(field.text(), "", "the whole run of typing goes at once");
    assert!(field.redo());
    assert_eq!(field.text(), "abc", "and redo puts the run back");

    // A caret jumped away from is a place the next character is not joined to: the run before the
    // jump and the run after it are two entries rather than one.
    let mut field = holding("ab");
    field.place_caret(1, false);
    field.insert('x');
    field.place_caret(3, false);
    field.insert('y');
    assert_eq!(field.text(), "axby");
    assert!(field.undo());
    assert_eq!(field.text(), "axb", "only the run since the jump goes");
    assert!(field.undo());
    assert_eq!(field.text(), "ab", "and the character before it next");

    // There is nothing left, and nothing to redo once a new edit has been made.
    assert!(!field.undo());
    let mut field = holding("ab");
    field.insert('c');
    assert!(field.undo());
    field.insert('d');
    assert!(
        !field.redo(),
        "a new edit gives up the redo it cannot reach"
    );
}

/// A double click takes the word under the pointer, by the page's own run rule: letters, digits and
/// the underscore are one run, a gap is its own, and a mark is its own so `->` comes out whole (see
/// `text_preview::run_at`).
#[test]
fn a_double_click_takes_the_word_under_it_by_the_pages_own_run_rule() {
    let field = holding("one->two 3_x");

    assert_eq!(field.word_at(1), (0, 3), "a letter is its word");
    assert_eq!(
        field.word_at(3),
        (3, 5),
        "the whole of `->` and not one of its marks"
    );
    assert_eq!(field.word_at(6), (5, 8), "the word after it");
    assert_eq!(field.word_at(8), (8, 9), "the gap is a run of its own");
    assert_eq!(
        field.word_at(9),
        (9, 12),
        "and the underscore keeps `3_x` one word"
    );
    assert_eq!(
        field.word_at(12),
        (12, 12),
        "past the end of the query there is no word to take"
    );
}

/// The caret is solid from the moment of an edit and turns over once an interval, so a reader typing
/// always sees it and a bar left standing blinks at the system's pace (see `CaretBlink`).
#[test]
fn the_caret_is_solid_from_an_edit_and_blinks_once_an_interval() {
    let now = Instant::now();
    let interval = Duration::from_millis(530);
    let mut blink = CaretBlink::default();

    blink.restart(now, interval);
    assert!(blink.lit(), "an edit leaves the caret solid");
    assert!(
        !blink.advance(now + Duration::from_millis(500), interval),
        "nothing turns over before the interval"
    );
    assert!(blink.lit());

    assert!(blink.advance(now + Duration::from_millis(531), interval));
    assert!(!blink.lit(), "and the first blink is the caret going off");
    assert!(blink.advance(now + Duration::from_millis(1062), interval));
    assert!(blink.lit(), "and coming back");

    blink.restart(now + Duration::from_millis(1100), interval);
    assert!(blink.lit(), "another edit makes it solid again");
}

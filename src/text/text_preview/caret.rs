//! Where a caret in a page of text goes: the moves the page's own keys ask for, measured in the
//! document rather than in the frame.
//!
//! **A caret is a place in the document, and every move is answered here.** A page is drawn a
//! screenful at a time, and the screenful moves under a caret that walks off its edge — so a move
//! that was worked out against the painted lines would have to be asked again after the scroll,
//! and a move over a line too long for the box would count rows the document does not have. The
//! document is therefore what every question below is asked of, once per key press, and the frame
//! is what the answer is then drawn in (see `render_scrolled`).
//!
//! **The goal column is the whole of what makes a vertical move feel like an editor's.** `↓` over a
//! short line and `↓` again has to come back to the column the reader was in, not to the end of the
//! line they passed over — so the column is remembered for as long as the moves stay vertical and
//! dropped the moment one does not (see `Motion::is_vertical`). It is a column of the *document*,
//! counted in characters from the start of the line, which is what a page's own wrapping does not
//! disturb.
//!
//! Nothing here paints, and nothing here knows what a key is: a move is named (`Motion`), and the
//! page's key table turns a key into one (see `pin_caret`). That is what lets every move be asked
//! of a document in a test with no window behind it.

use super::document::Place;
use super::frame::is_word_character;

/// One move a caret can be asked for, named rather than written as a place: what a key asks for is
/// a direction and a unit, and where that lands is the document's answer (see `moved`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum Motion {
    /// One character back, crossing a line end to the end of the line above.
    Left,
    /// One character on, crossing a line end to the start of the line below.
    Right,
    /// The start of the word before the caret, by the page's own run rule.
    WordLeft,
    /// The start of the word after the caret, by the same rule.
    WordRight,
    /// One line up, at the goal column where the line is long enough for it.
    Up,
    /// One line down, the same.
    Down,
    /// The first character of the caret's own line.
    LineStart,
    /// The end of the caret's own line, which is the character before its break.
    LineEnd,
    /// The first character of the document.
    DocumentStart,
    /// The last character of the document.
    DocumentEnd,
    /// A screenful up, at the goal column: a page turned without the caret losing its place in the
    /// line it lands on.
    ScreenUp(usize),
    /// The same, down.
    ScreenDown(usize),
}

impl Motion {
    /// Whether this move keeps the goal column. Only the moves that go straight up and down do:
    /// a character move, a word move and a move to an end all leave the caret somewhere the
    /// reader chose, and the column they are in is the column the next vertical move should use.
    pub fn is_vertical(self) -> bool {
        matches!(
            self,
            Motion::Up | Motion::Down | Motion::ScreenUp(_) | Motion::ScreenDown(_)
        )
    }
}

/// Where `caret` goes for `motion` in a document drawn as `lines`, and the goal column it leaves
/// behind: the column a vertical move is aimed at, or `None` for a move that has left the caret
/// somewhere the reader chose (see `Motion::is_vertical`).
///
/// `goal` is what a vertical move aims at — the column of the last vertical move, or of the move
/// that began the run. A vertical move over a line too short for it lands at the line's end and
/// keeps the goal, which is the whole of why `↓` over a short line and `↑` back comes out where it
/// started.
pub fn moved(
    lines: &[String],
    caret: Place,
    motion: Motion,
    goal: Option<usize>,
) -> (Place, Option<usize>) {
    if lines.is_empty() {
        return (caret, goal);
    }

    let last = lines.len() - 1;
    let (line, character) = (caret.0.min(last), caret.1);
    let here = character.min(lines[line].chars().count());

    // The column a vertical move aims at: the goal where there is one, and the caret's own column
    // where the move begins a run of them.
    let aimed = goal.unwrap_or(here);

    let place = match motion {
        Motion::Left => left(lines, line, here),
        Motion::Right => right(lines, line, here, last),
        Motion::WordLeft => word_left(lines, line, here),
        Motion::WordRight => word_right(lines, line, here, last),
        Motion::Up => vertical(lines, line, aimed, -1),
        Motion::Down => vertical(lines, line, aimed, 1),
        Motion::ScreenUp(step) => vertical(lines, line, aimed, -(step as i64)),
        Motion::ScreenDown(step) => vertical(lines, line, aimed, step as i64),
        Motion::LineStart => (line, 0),
        Motion::LineEnd => (line, line_characters(lines, line)),
        Motion::DocumentStart => (0, 0),
        Motion::DocumentEnd => (last, line_characters(lines, last)),
    };

    let kept = if motion.is_vertical() {
        Some(aimed)
    } else {
        None
    };
    (place, kept)
}

/// How many characters one drawn line holds, which is what the end of a line is.
fn line_characters(lines: &[String], line: usize) -> usize {
    lines.get(line).map_or(0, |text| text.chars().count())
}

/// One character back: the end of the line above at the start of a line, and the start of the
/// document at its beginning.
fn left(lines: &[String], line: usize, at: usize) -> Place {
    if at > 0 {
        (line, at - 1)
    } else if line > 0 {
        (line - 1, line_characters(lines, line - 1))
    } else {
        (0, 0)
    }
}

/// One character on: the start of the line below at the end of a line, and the end of the document
/// at its end — which is where a line's break would be, so a caret at the end of the last line has
/// nowhere further to go.
fn right(lines: &[String], line: usize, at: usize, last: usize) -> Place {
    let end = line_characters(lines, line);
    if at < end {
        (line, at + 1)
    } else if line < last {
        (line + 1, 0)
    } else {
        (line, end)
    }
}

/// The characters of one line, for the two word moves: the run rule is the page's own, and it is a
/// rule about characters rather than about columns, so a line is asked of it as a string.
fn characters_of(lines: &[String], line: usize) -> Vec<char> {
    lines
        .get(line)
        .map(|text| text.chars().collect())
        .unwrap_or_default()
}

/// The start of the word before the caret, or of the run before it where that run is not a word:
/// the run rule a double press takes a word by (see `is_word_character`), so a `Ctrl`+`←` out of a
/// word lands at the start of the word, and one out of a gap lands at the start of the word before
/// the gap.
fn word_left(lines: &[String], line: usize, at: usize) -> Place {
    let mut line = line;
    let mut at = at;

    loop {
        let characters = characters_of(lines, line);
        while at > 0 && !is_word_character(characters[at - 1]) {
            at -= 1;
        }
        while at > 0 && is_word_character(characters[at - 1]) {
            at -= 1;
        }
        // A line that begins with the word the caret was in is where the word starts: the line
        // above is a line the reader is not in, and walking onto it would take a word move over a
        // line end rather than over a word.
        if at > 0 || line == 0 {
            return (line, at);
        }
        // Nothing of this line is before the caret, so the word it is in, if any, is the one that
        // ends the line above.
        line -= 1;
        at = line_characters(lines, line);
    }
}

/// The start of the word after the caret, by the same rule: on over the run the caret is in, over
/// the gap after it, and the caret stands at the first character of the next word — or at the end
/// of the document where there is none.
fn word_right(lines: &[String], line: usize, at: usize, last: usize) -> Place {
    let mut line = line;
    let mut at = at;

    loop {
        let characters = characters_of(lines, line);
        while at < characters.len() && is_word_character(characters[at]) {
            at += 1;
        }
        while at < characters.len() && !is_word_character(characters[at]) {
            at += 1;
        }
        if at < characters.len() || line == last {
            return (line, at);
        }
        // The caret has walked off the end of this line, so the word after it begins the line
        // below — unless that line begins with a gap, which the loop above walks over.
        line += 1;
        at = 0;
    }
}

/// A move `step` lines up or down, at `column` where the line it lands on is long enough for it and
/// at the line's own end where it is not. Nothing past either end of the document: a move that has
/// nowhere to go leaves the caret where it is, which is what makes a key held at the last line
/// leave the page standing.
fn vertical(lines: &[String], line: usize, column: usize, step: i64) -> Place {
    let last = lines.len() - 1;
    let target = (line as i64 + step).clamp(0, last as i64) as usize;
    (target, column.min(line_characters(lines, target)))
}

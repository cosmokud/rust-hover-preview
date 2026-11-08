//! The find bar's query, as a field rather than a box a character is appended to: the text, the
//! caret in it, the selection between the two, and every edit that can change any of them.
//!
//! **Every place in here is a byte index on a character boundary.** The query is a `String` and the
//! text arrives from a paste as well as from the keyboard, so a caret counted in characters would
//! need a walk of the text to place it and a caret counted in bytes that was allowed off a boundary
//! would split a character the moment a backspace took the wrong half of it. Every index the field
//! hands out is one `char_indices` produced, and every range it cuts is one of those.
//!
//! **The field knows nothing about the screen.** Where the caret is drawn, how wide a character is,
//! and which place a press names are the card's questions, and the card answers them with
//! `pin_chrome::find_bar` — a press is turned into a byte index there, from the columns the card
//! measured its run at, and handed here as a place. That split is what lets every edit, every
//! selection and every undo in this file be asked of a field with no window behind it.
//!
//! **Undo is a stack of the whole field.** A snapshot is the query, the caret and the selection
//! together, because undo that put the text back without the caret would leave the reader typing
//! somewhere other than where they were. A run of typed characters is one snapshot rather than one
//! per character: a reader who undoes a word expects the word to go, which is what every editor's
//! own coalescing gives them.

use super::*;

/// The most characters a query may hold. A query is one line of a field one row tall, and a paste
/// of a whole document into it would be a query the search walks with nothing on the card to say
/// which part of it is being read (see `paste`).
pub(super) const QUERY_LIMIT: usize = 512;

/// How many edits a field remembers before it forgets the oldest. A query is at most
/// `QUERY_LIMIT` characters, so a deeper stack costs more memory than the query it is undoing.
const UNDO_DEPTH: usize = 100;

/// A caret move, named rather than written as a place: what a key asks for is a direction and a
/// unit, and where that lands is the field's answer for the text it holds (see `Field::move_caret`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum Motion {
    CharacterLeft,
    CharacterRight,
    WordLeft,
    WordRight,
    Home,
    End,
}

/// The query, the caret in it, the selection between the caret and its anchor, and the two stacks
/// undo and redo are kept on.
///
/// A caret with no selection has `anchor == caret`; when they differ the selection is the range
/// between them, and neither end is the same end in every gesture — a `Shift`+`←` puts the caret
/// before the anchor, and a cut takes the range in document order rather than in the order the
/// two were placed (see `selection`).
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub(super) struct Field {
    text: String,
    caret: usize,
    anchor: usize,
    undo: Vec<Snapshot>,
    redo: Vec<Snapshot>,
    /// Whether the last edit was one character typed, which the next typed character joins as one
    /// undo entry. A caret move, a delete, a paste or a selection begins a new one (see `snapshot`).
    typing: bool,
}

/// The whole of what undo puts back: the query and the two ends the reader was between.
#[derive(Clone, Debug, PartialEq, Eq)]
struct Snapshot {
    text: String,
    caret: usize,
    anchor: usize,
}

impl Field {
    /// The query as the search reads it.
    pub(super) fn text(&self) -> &str {
        &self.text
    }

    /// Where the caret is, and the other end of the selection.
    pub(super) fn caret(&self) -> usize {
        self.caret
    }

    pub(super) fn anchor(&self) -> usize {
        self.anchor
    }

    /// The selected range in the text's own order, or nothing when the two ends are one place.
    pub(super) fn selection(&self) -> Option<(usize, usize)> {
        (self.anchor != self.caret).then_some(if self.anchor < self.caret {
            (self.anchor, self.caret)
        } else {
            (self.caret, self.anchor)
        })
    }

    /// What a copy of the field would take.
    pub(super) fn selected_text(&self) -> String {
        match self.selection() {
            Some((start, end)) => self.text[start..end].to_string(),
            None => String::new(),
        }
    }

    /// Select the whole of the query, which is what `Ctrl`+`A` asks and what `Ctrl`+`F` over a bar
    /// that is already up asks as well (see `text_find::open_text_find`).
    pub(super) fn select_all(&mut self) {
        self.anchor = 0;
        self.caret = self.text.len();
        self.typing = false;
    }

    /// Put an empty query back, with the caret at its beginning: `Esc` on a bar that holds text is
    /// this, and the search that follows finds nothing and puts the highlights away.
    pub(super) fn clear(&mut self) {
        if self.text.is_empty() {
            return;
        }

        self.snapshot();
        self.text.clear();
        self.caret = 0;
        self.anchor = 0;
        self.typing = false;
    }

    /// Put a run of text in the query whole, in place of what it held: what `Ctrl`+`F` over a page
    /// with something selected seeds the bar with, and what a step through the run's search history
    /// puts there (see `text_find::open_text_find` and `text_find::step_the_history`).
    ///
    /// **A seed is not a paste.** `paste` tidies what it is given, and tidying is right for text
    /// from another program and wrong for this one: a reader who selected two spaces and pressed the
    /// key asked for a query of two spaces, so the run is taken as written, capped at `QUERY_LIMIT`
    /// and cut at its first line break — the query is one line, and what follows a break in a
    /// selection is a second line of it rather than more of the same query. The caret is left at the
    /// end of it, where the next character typed lands.
    ///
    /// **The whole of it is selected**, which is the one thing an editor's find box does with a seed
    /// that is not a caret: the reader asked for the text that is now in the box, and the next
    /// character typed is the new query rather than an append to the old one (see `select_all`).
    ///
    /// Answers whether there was anything to seed, so that a page with an empty selection leaves the
    /// query the bar was last left holding alone.
    pub(super) fn seed(&mut self, seeded: &str) -> bool {
        let line = seeded.split(['\r', '\n']).next().unwrap_or_default();
        let taken: String = line
            .chars()
            .filter(|character| !character.is_control())
            .take(QUERY_LIMIT)
            .collect();
        if taken.is_empty() {
            return false;
        }

        self.snapshot();
        self.text = taken;
        self.anchor = 0;
        self.caret = self.text.len();
        self.typing = false;
        true
    }

    /// Put the caret at `at`, or with `extend` keep the anchor where it is and grow the selection
    /// from it. `at` is clamped to a character boundary, so a place the card measured off a glyph
    /// can never split one.
    pub(super) fn place_caret(&mut self, at: usize, extend: bool) {
        let at = boundary_at_or_before(&self.text, at);
        if !extend {
            self.anchor = at;
        }
        self.caret = at;
        // A caret jump is the end of a typed run: the next character typed is an edit of its own,
        // which is what a reader undoing it expects to get back to the place they jumped from (see
        // `snapshot`).
        self.typing = false;
    }

    /// `place_caret` with the answer a drag needs: whether anything drawn changed. A move inside
    /// the character the caret already stands on moves nothing, which is most of the messages a
    /// drag sends.
    pub(super) fn place_caret_changed(&mut self, at: usize, extend: bool) -> bool {
        let before = (self.caret, self.anchor);
        self.place_caret(at, extend);
        (self.caret, self.anchor) != before
    }

    /// Move the caret by `motion`, extending the selection rather than collapsing it where
    /// `extend` says so. Answers whether anything changed, which is what tells a caller a repaint
    /// is owed.
    pub(super) fn move_caret(&mut self, motion: Motion, extend: bool) -> bool {
        let to = match motion {
            Motion::CharacterLeft => previous_boundary(&self.text, self.caret),
            Motion::CharacterRight => next_boundary(&self.text, self.caret),
            Motion::WordLeft => word_left(&self.text, self.caret),
            Motion::WordRight => word_right(&self.text, self.caret),
            Motion::Home => 0,
            Motion::End => self.text.len(),
        };
        // The two ends after the move against the two ends before it, rather than the place against the
        // flag: a `Shift`+`←` with the selection already at the start of the query names `extend` and
        // lands where it stood, which is a move that moved nothing.
        let before = (self.caret, self.anchor);

        if !extend {
            self.anchor = to;
        }
        self.caret = to;
        self.typing = false;
        (self.caret, self.anchor) != before
    }

    /// Type a character at the caret, replacing what is selected.
    ///
    /// A character that would take the query past `QUERY_LIMIT` is dropped whole rather than cut:
    /// half a character is not a query, and neither is a query with its last character's first byte
    /// (see `QUERY_LIMIT`).
    pub(super) fn insert(&mut self, character: char) {
        if character.is_control() {
            return;
        }

        let (start, end) = self.selection().unwrap_or((self.caret, self.caret));
        if self.room(start, end) == 0 {
            return;
        }

        let joined = self.typing;
        if !joined {
            self.snapshot();
        }
        self.text.replace_range(start..end, &character.to_string());
        self.caret = start + character.len_utf8();
        self.anchor = self.caret;
        self.typing = true;
    }

    /// Take the selection out, or the character before the caret where nothing is selected: nothing
    /// at the beginning of the query, and no wrap to its end.
    pub(super) fn backspace(&mut self) {
        if let Some((start, end)) = self.selection() {
            self.delete_range(start, end);
            return;
        }

        if self.caret == 0 {
            return;
        }
        let start = previous_boundary(&self.text, self.caret);
        self.delete_range(start, self.caret);
    }

    /// Take the selection out, or the character after the caret where nothing is selected: nothing
    /// at the end of the query, and no wrap to its beginning.
    pub(super) fn delete(&mut self) {
        if let Some((start, end)) = self.selection() {
            self.delete_range(start, end);
            return;
        }

        if self.caret == self.text.len() {
            return;
        }
        let end = next_boundary(&self.text, self.caret);
        self.delete_range(self.caret, end);
    }

    /// `Ctrl`+`Backspace`: the selection, or the word before the caret.
    pub(super) fn delete_word_left(&mut self) {
        let (start, end) = self
            .selection()
            .unwrap_or((word_left(&self.text, self.caret), self.caret));
        self.delete_range(start, end);
    }

    /// `Ctrl`+`Delete`: the selection, or the word after the caret.
    pub(super) fn delete_word_right(&mut self) {
        let (start, end) = self
            .selection()
            .unwrap_or((self.caret, word_right(&self.text, self.caret)));
        self.delete_range(start, end);
    }

    /// Take the selection out and hand it back for the clipboard. Answers nothing where there is
    /// nothing selected, which is a cut that writes no clipboard at all.
    pub(super) fn cut(&mut self) -> String {
        let cut = self.selected_text();
        if cut.is_empty() {
            return cut;
        }

        if let Some((start, end)) = self.selection() {
            self.delete_range(start, end);
        }
        cut
    }

    /// Paste a run of text at the caret, replacing what is selected.
    ///
    /// The run is one line before it ever reaches the query (see `sanitize_pasted`), and it is
    /// capped twice: by the sanitizer's own limit, and by the room the field has left for it. A
    /// paste longer than the room inserts the room's worth of characters rather than refusing the
    /// paste whole, which is what a reader pasting a sentence into a full field means.
    pub(super) fn paste(&mut self, pasted: &str) {
        let clean = sanitize_pasted(pasted);
        let (start, end) = self.selection().unwrap_or((self.caret, self.caret));
        let room = self.room(start, end);
        if clean.is_empty() || room == 0 {
            return;
        }

        let taken: String = clean.chars().take(room).collect();
        self.snapshot();
        self.text.replace_range(start..end, &taken);
        self.caret = start + taken.len();
        self.anchor = self.caret;
        self.typing = false;
    }

    /// Put back the field as it was before the last edit, answering whether there was one.
    pub(super) fn undo(&mut self) -> bool {
        let Some(previous) = self.undo.pop() else {
            return false;
        };

        self.redo.push(self.snapshot_now());
        cap(&mut self.redo);
        self.restore(previous);
        true
    }

    /// Put back the field an undo took away, answering whether there was one.
    pub(super) fn redo(&mut self) -> bool {
        let Some(next) = self.redo.pop() else {
            return false;
        };

        self.undo.push(self.snapshot_now());
        cap(&mut self.undo);
        self.restore(next);
        true
    }

    /// The word a double click at `at` takes: the page's own run rule, over the query — letters,
    /// digits and the underscore are one run, a gap is a run of its own, and every mark is its own
    /// kind (see `text_preview::run_at`).
    pub(super) fn word_at(&self, at: usize) -> (usize, usize) {
        text_preview::run_at(&self.text, boundary_at_or_before(&self.text, at))
    }

    /// How many characters the query would have left if `start..end` were replaced: the room an
    /// insertion has before the limit, counted in characters rather than bytes so that a paste of
    /// text the query cannot hold is refused in the unit the limit is written in.
    fn room(&self, start: usize, end: usize) -> usize {
        let held = self.text.chars().count() - self.text[start..end].chars().count();
        QUERY_LIMIT.saturating_sub(held)
    }

    fn snapshot_now(&self) -> Snapshot {
        Snapshot {
            text: self.text.clone(),
            caret: self.caret,
            anchor: self.anchor,
        }
    }

    /// Write down the field before an edit, and give up the redo it can no longer reach. A run of
    /// typed characters takes this once, at the character that began the run.
    fn snapshot(&mut self) {
        self.undo.push(self.snapshot_now());
        cap(&mut self.undo);
        self.redo.clear();
    }

    fn restore(&mut self, snapshot: Snapshot) {
        self.text = snapshot.text;
        self.caret = snapshot.caret;
        self.anchor = snapshot.anchor;
        self.typing = false;
    }

    /// Take `start..end` out as one edit, leaving the caret where the cut began. An empty range is
    /// no edit at all: nothing is written down and nothing is redone away.
    fn delete_range(&mut self, start: usize, end: usize) -> bool {
        if start >= end {
            return false;
        }

        self.snapshot();
        self.text.replace_range(start..end, "");
        self.caret = start;
        self.anchor = start;
        self.typing = false;
        true
    }
}

#[cfg(test)]
impl Field {
    /// A field holding `text` with the caret after its last character and nothing behind it to undo:
    /// the state a query arrives in when the test is about the edits rather than about the typing.
    pub(super) fn holding(text: &str) -> Self {
        Self {
            text: text.to_string(),
            caret: text.len(),
            anchor: text.len(),
            ..Self::default()
        }
    }
}

/// Whether the caret is drawn at this moment, and when it is next to be drawn the other way.
///
/// **Solid from the moment of any edit**, so the reader typing sees the caret rather than a blink
/// that happened to be off when they looked. Every edit and every press restarts it (see
/// `restart`), and the tick that owns the cursor asks `advance` rather than reading a clock itself
/// (see `text_find::settle`).
#[derive(Clone, Copy, Debug)]
pub(super) struct CaretBlink {
    lit: bool,
    due: Instant,
}

impl Default for CaretBlink {
    fn default() -> Self {
        Self {
            lit: true,
            due: Instant::now(),
        }
    }
}

impl CaretBlink {
    pub(super) fn lit(&self) -> bool {
        self.lit
    }

    /// Draw the caret solid from `now`, and blink it again one interval later.
    pub(super) fn restart(&mut self, now: Instant, interval: Duration) {
        self.lit = true;
        self.due = now + interval;
    }

    /// Turn the phase over where it has come due, answering whether what is drawn changed. A tick
    /// that arrives after several phases have passed turns it once rather than chasing every one it
    /// missed: a blink is a cadence, and a caret flapping through four phases in one repaint is not
    /// four frames of anything.
    pub(super) fn advance(&mut self, now: Instant, interval: Duration) -> bool {
        if now < self.due {
            return false;
        }

        self.lit = !self.lit;
        self.due = now + interval;
        true
    }
}

/// The last `UNDO_DEPTH` snapshots, and no more: the oldest is given up rather than the newest.
fn cap(stack: &mut Vec<Snapshot>) {
    if stack.len() > UNDO_DEPTH {
        stack.remove(0);
    }
}

/// A pasted run of text as a query can hold it: one line, with its breaks as spaces and its runs of
/// space as one space each.
///
/// A paste is the one road a query gets text the keyboard could not have typed — a whole paragraph,
/// a line copied out of a terminal, a file with a byte-order mark at its front — and the search
/// reads a query as one line of one row (see `paint_find_bar`). So `\r`, `\n` and `\t` become a
/// space; every other control character and the byte-order mark are dropped rather than turned
/// into marks the field cannot draw; and a run of whitespace is one space, because the indentation
/// of a pasted block would otherwise be a query beginning with a dozen spaces. A run at either end
/// is dropped with it rather than kept as a space nothing is searched beside. The result is capped
/// at `QUERY_LIMIT` in characters, never splitting the last one it takes.
pub(super) fn sanitize_pasted(pasted: &str) -> String {
    let mut clean = String::new();
    let mut space_pending = false;

    for character in pasted.chars() {
        if character == '\u{feff}' {
            continue;
        }
        if character.is_whitespace() {
            space_pending = true;
            continue;
        }
        if character.is_control() {
            continue;
        }

        if space_pending && !clean.is_empty() {
            clean.push(' ');
        }
        space_pending = false;
        clean.push(character);
        if clean.chars().count() >= QUERY_LIMIT {
            break;
        }
    }

    clean
}

/// The last character boundary at or before `at`, so that a place the card measured from a column
/// is always a place the field can cut at.
pub(super) fn boundary_at_or_before(text: &str, at: usize) -> usize {
    let mut at = at.min(text.len());
    while at > 0 && !text.is_char_boundary(at) {
        at -= 1;
    }
    at
}

/// The boundary before `at`, or the beginning of the text.
fn previous_boundary(text: &str, at: usize) -> usize {
    text[..at]
        .chars()
        .next_back()
        .map_or(0, |character| at - character.len_utf8())
}

/// The boundary after `at`, or the end of the text.
fn next_boundary(text: &str, at: usize) -> usize {
    text[at..]
        .chars()
        .next()
        .map_or(text.len(), |character| at + character.len_utf8())
}

/// The beginning of the word before `at`: back over the marks between one word and the next, and
/// then back over the word itself. The word rule is the page's own (see `text_preview::run_at`).
fn word_left(text: &str, at: usize) -> usize {
    let mut at = at;

    while at > 0 {
        let before = previous_boundary(text, at);
        if text[before..at]
            .chars()
            .all(text_preview::is_word_character)
        {
            break;
        }
        at = before;
    }
    while at > 0 {
        let before = previous_boundary(text, at);
        if !text[before..at]
            .chars()
            .all(text_preview::is_word_character)
        {
            break;
        }
        at = before;
    }

    at
}

/// The end of the word after `at`: on over the word, and then over the marks after it.
fn word_right(text: &str, at: usize) -> usize {
    let mut at = at;

    while at < text.len() {
        let after = next_boundary(text, at);
        if !text[at..after].chars().all(text_preview::is_word_character) {
            break;
        }
        at = after;
    }
    while at < text.len() {
        let after = next_boundary(text, at);
        if text[at..after].chars().all(text_preview::is_word_character) {
            break;
        }
        at = after;
    }

    at
}

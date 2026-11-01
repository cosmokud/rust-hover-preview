//! The parsed document: the file as it was read, the caches it is kept in, and
//! the styled windows a preview draws a line at a time.
//!
//! Nothing here paints or places a line. What it answers is what a document is,
//! how many lines it has, and which of them carry a color of their own — so a
//! preview can say how tall a file is before any of it is styled, and style only
//! the screenful that is on screen (see [`styled_window`]).
//!
//! Which window a line is styled by is the file's own name and mode deciding, in
//! [`build_document`]: a syntax definition for source ([`window_highlighted`]),
//! the Markdown renderer in [`super::markdown`], or plain and ANSI text read by
//! [`super::source`].

use super::frame::{is_word_character, Selection, TextPreviewOptions};
use super::markdown::{rendered_line_count, window_markdown};
use super::source::{push_ansi_spans, read_source, strip_rtf, AnsiPalette};
use crate::config::config::{MarkdownMode, TextTheme};
use crate::text::text_paint::{text_style, TextStyle, BODY_LEVEL};
use crate::text::text_theme::{self, LoadedTheme};
use once_cell::sync::Lazy;
use std::cell::RefCell;
use std::collections::HashMap;
use std::path::{Path, PathBuf};
#[cfg(test)]
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex};
use std::time::SystemTime;
use syntect::easy::HighlightLines;
use syntect::highlighting::HighlightState;
use syntect::parsing::{ParseState, SyntaxReference, SyntaxSet};

/// Lines a preview can scroll through. A source file is styled a window at a
/// time, so this is not a memory bound but a bound on how far one hover can walk
/// into a file — a few dozen screens of text.
pub(super) const MAX_DOC_LINES: usize = 2000;
/// Lines styled between parse checkpoints. Highlighting is sequential, so the
/// only way to show a window near the end of a file is to have parsed from
/// somewhere before it; this is how far back that somewhere can be, and so what
/// a jump that cannot continue from where it stopped costs.
const CHECKPOINT_INTERVAL_LINES: usize = 32;
const CHECKPOINT_MAX_ENTRIES: usize = 192;
/// Styled windows kept per document. Scrolling a line at a time lands inside the
/// window that is already built, and a repaint is free.
const WINDOW_CACHE_MAX_ENTRIES: usize = 6;
/// Documents one thread keeps parse progress for. Small: it is there so a scroll
/// can continue rather than to remember every file ever hovered.
const SYNTAX_PROGRESS_MAX_ENTRIES: usize = 8;
/// Parsed documents kept in memory, keyed by file and rendering options.
const DOC_CACHE_MAX_ENTRIES: usize = 24;
/// A text file as it is read: the decoded text and where its lines start.
///
/// Styling is deliberately not part of this. A line table costs eight bytes per
/// line and no parsing, which is what lets a preview know how tall a document is
/// — and how far it can scroll — before anything has been colored. The styled
/// lines for the part of it that is on screen are built on demand, in
/// [`styled_window`].
pub(super) struct TextDoc {
    pub(super) text: String,
    /// Byte offset of the start of every line, so `line_starts.len()` is the
    /// file's own line count.
    line_starts: Vec<usize>,
    /// Lines a frame of this document is drawn in, when that is not the file's
    /// own line count. A rendered Markdown document is the one producer whose
    /// lines are its own — a paragraph is one line and the blank line between two
    /// blocks is one the source does not have — so its scroll position, its
    /// scrollbar and the lines it says are left out are all measured in those.
    rendered_lines: Option<usize>,
    producer: Producer,
    /// Whether the read stopped at the byte cap rather than at the end of the
    /// file, in which case a file continues past what a preview can show.
    pub(super) read_truncated: bool,
}

impl TextDoc {
    /// Lines there are to show, in the space a frame is drawn in.
    pub(super) fn total_lines(&self) -> usize {
        self.rendered_lines.unwrap_or(self.line_starts.len())
    }

    /// Lines this preview can reach. A source file is styled a window at a time,
    /// but the range a preview can scroll through is capped so a hover can never
    /// walk an unbounded file.
    pub(super) fn scrollable_lines(&self) -> usize {
        self.total_lines().min(MAX_DOC_LINES)
    }

    fn line(&self, index: usize) -> &str {
        let start = self.line_starts[index];
        let end = self
            .line_starts
            .get(index + 1)
            .copied()
            .unwrap_or(self.text.len());

        self.text[start..end].trim_end_matches('\n')
    }
}

/// Which renderer turns a document's lines into styled spans.
#[derive(Clone)]
enum Producer {
    /// Source and markup: a syntax definition colors the lines, and the parser's
    /// state carries from one line to the next, which is what makes it possible
    /// to start part-way into a file.
    Highlighted { syntax: &'static SyntaxReference },
    /// A Markdown document, laid out as the document it describes.
    Markdown,
    /// Text that needs no coloring of its own, such as an RTF's stripped text.
    Plain,
    /// Text carrying ANSI color escapes.
    Ansi,
}

/// Styled lines for a range of one document. `first` is the document line the
/// first entry belongs to, so callers index by absolute line number.
pub(super) struct TextWindow {
    pub(super) first: usize,
    pub(super) lines: Vec<DocLine>,
}

impl TextWindow {
    pub(super) fn line(&self, index: usize) -> Option<&DocLine> {
        self.lines.get(index.checked_sub(self.first)?)
    }

    fn covers(&self, first: usize, count: usize) -> bool {
        self.first <= first && self.first + self.lines.len() >= first + count
    }
}

/// How far the highlighter has walked through one document, and the states it
/// would need to start again from a few points inside it.
///
/// Highlighting is sequential — a grammar's state at line N depends on every line
/// before it — so the only way to show a window near the end of a file without
/// parsing the whole file is to have parsed it once, in order, and to remember
/// where it was. The checkpoint spacing bounds what a jump costs, and the states
/// are small: a scope stack and a style stack, not the lines themselves.
struct SyntaxProgress {
    /// Line the parse has reached, and the state that resumes exactly there.
    frontier: usize,
    frontier_state: (HighlightState, ParseState),
    checkpoints: Vec<(usize, HighlightState, ParseState)>,
}

impl SyntaxProgress {
    fn continuation(&self, first: usize) -> (usize, HighlightState, ParseState) {
        if self.frontier <= first {
            let (highlight, parse) = self.frontier_state.clone();
            return (self.frontier, highlight, parse);
        }

        self.checkpoints
            .iter()
            .rev()
            .find(|(line, _, _)| *line <= first)
            .map(|(line, highlight, parse)| (*line, highlight.clone(), parse.clone()))
            .unwrap_or_else(|| {
                let (highlight, parse) = self.frontier_state.clone();
                (0, highlight, parse)
            })
    }
}

/// Styled windows built for one document. Plain data, so it is shared between the
/// threads that measure and paint a preview.
#[derive(Default)]
struct WindowCache {
    windows: Vec<Arc<TextWindow>>,
}

/// A parsed document with its window cache. Shared through the global cache, so
/// the windows outlive a single preview and a second hover in the same area costs
/// nothing at all.
pub(super) struct CachedDocument {
    /// What this document was built from, which is also how the per-thread parse
    /// progress finds it again.
    pub(super) key: DocKey,
    pub(super) doc: TextDoc,
    windows: Mutex<WindowCache>,
    /// Every line of the document as it is drawn, built once and kept: the text a caret moves
    /// against and a search reads.
    drawn: Mutex<Option<Arc<Vec<String>>>>,
}

impl CachedDocument {
    /// Every line of the document as it is drawn, built once and kept.
    ///
    /// **Nothing past what a page can reach is walked.** A preview can only scroll `scrollable_lines`
    /// of a file (`MAX_DOC_LINES`), so the rest is a document no frame can show — walking it would be
    /// styling a million lines of a log to move a caret over the first two thousand of them, which is
    /// what made an arrow key cost a whole file.
    ///
    /// **A document nothing renders is its own lines**, and they are sliced out of the text the parse
    /// already holds: styling decides the colour of a span and never its text, so a plain file's lines
    /// are one slice each rather than a document through the highlighter. Only a producer that draws
    /// lines of its own — rendered Markdown, whose paragraphs are its lines — is walked, and that one
    /// is walked once.
    fn drawn_text(&self, options: TextPreviewOptions) -> Arc<Vec<String>> {
        if let Ok(held) = self.drawn.lock() {
            if let Some(drawn) = held.as_ref() {
                return Arc::clone(drawn);
            }
        }

        let total = self.doc.scrollable_lines();
        let mut lines: Vec<String> = Vec::with_capacity(total);
        if self.doc.rendered_lines.is_none() {
            for index in 0..total {
                lines.push(self.doc.line(index).to_string());
            }
        } else if total > 0 {
            if let Some(theme) = text_theme::loaded(options.theme) {
                walk_drawn_lines(self, theme, 0, total - 1, |_, text| {
                    lines.push(text);
                    true
                });
            }
        }

        let drawn = Arc::new(lines);
        if let Ok(mut held) = self.drawn.lock() {
            *held = Some(Arc::clone(&drawn));
        }
        drawn
    }
}

thread_local! {
    /// Where the highlighter has reached in each document, on this thread.
    ///
    /// syntect's parse state is not `Send` — it holds pointers into Oniguruma's
    /// match regions — so it cannot live in the shared document cache no matter
    /// how useful it would be there. Styled windows are shared, because they are
    /// plain data; the state that continues a parse from part-way into a file is
    /// rebuilt per thread, which costs a bounded re-parse at worst.
    static SYNTAX_PROGRESS: RefCell<HashMap<DocKey, SyntaxProgress>> =
        RefCell::new(HashMap::new());
}

/// What a parsed document is cached by. The font scale is not part of it: a
/// document is the same text at any size, and only the layout changes.
#[derive(Clone, PartialEq, Eq, Hash)]
pub(super) struct DocKey {
    path: PathBuf,
    modified: Option<SystemTime>,
    len: u64,
    theme: TextTheme,
    /// Which reading of a user's theme file the document was styled against. The
    /// bundled themes cannot change while the process runs, so only a file theme
    /// carries one, and the same file read again is a different document.
    theme_generation: u64,
    markdown_mode: MarkdownMode,
}

/// Parsed documents, the most recently built last, keyed by the file and the
/// options they were built from. Bounded: the whole cache is dropped when it
/// fills, the way the PDF page and video geometry caches are.
type DocumentCache = Vec<(DocKey, Arc<CachedDocument>)>;

static DOCUMENTS: Lazy<Mutex<DocumentCache>> = Lazy::new(|| Mutex::new(Vec::new()));

/// Syntax definitions from syntect plus the ones it does not bundle, built once
/// on the first text hover. Individual syntaxes are still parsed lazily.
pub(super) static SYNTAXES: Lazy<SyntaxSet> = Lazy::new(two_face::syntax::extra_no_newlines);

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub(super) enum LineBlock {
    #[default]
    Plain,
    /// A line of a fenced or indented code block.
    Code,
    /// A line inside a block quote.
    Quote,
}

#[derive(Default, Clone)]
pub(super) struct DocLine {
    pub(super) spans: Vec<Span>,
    pub(super) block: LineBlock,
    /// Extra left indent, in characters, for nested lists and quotes.
    pub(super) indent: u8,
}

#[derive(Clone)]
pub(super) struct Span {
    pub(super) text: String,
    pub(super) style: TextStyle,
}

/// The parsed document for `path`, from the cache when the file, its mode and
/// its theme are unchanged.
pub(super) fn document(path: &Path, options: TextPreviewOptions) -> Option<Arc<CachedDocument>> {
    let metadata = std::fs::metadata(path).ok()?;
    #[cfg(test)]
    DOCUMENT_STATS.fetch_add(1, Ordering::Relaxed);
    let key = DocKey {
        path: path.to_path_buf(),
        modified: metadata.modified().ok(),
        len: metadata.len(),
        theme: options.theme,
        theme_generation: match options.theme {
            TextTheme::Custom(_) => text_theme::generation(),
            _ => 0,
        },
        markdown_mode: options.markdown_mode,
    };

    if let Ok(cache) = DOCUMENTS.lock() {
        if let Some((_, document)) = cache.iter().find(|(cached, _)| *cached == key) {
            return Some(Arc::clone(document));
        }
    }

    let document = Arc::new(CachedDocument {
        key: key.clone(),
        doc: build_document(path, options)?,
        windows: Mutex::new(WindowCache::default()),
        drawn: Mutex::new(None),
    });

    if let Ok(mut cache) = DOCUMENTS.lock() {
        if cache.len() >= DOC_CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.push((key, Arc::clone(&document)));
    }

    Some(document)
}

/// Lines a copy styles at once. A selection can span a whole document, and styling that in one
/// window is every line of it held at once: the walk is taken a window at a time, and the windows
/// it leaves are the ones a scroll already keeps.
const COPY_WINDOW_LINES: usize = 512;

/// Walk the drawn lines of `[first, last]` in windows, handing each line's own text to `visit`
/// until it answers false.
///
/// What a line is drawn as is not always what the file holds — a rendered Markdown document has
/// lines the file does not — so both the copy and the search are read from here, the way the frame
/// itself is (see `styled_window`).
fn walk_drawn_lines(
    document: &CachedDocument,
    theme: &LoadedTheme,
    first: usize,
    last: usize,
    mut visit: impl FnMut(usize, String) -> bool,
) {
    let mut index = first;
    while index <= last {
        let count = (last - index + 1).min(COPY_WINDOW_LINES);
        let Some(window) = styled_window(&document.key, document, theme, index, count) else {
            return;
        };

        let mut line_index = index;
        while line_index < index + count {
            let Some(line) = window.line(line_index) else {
                return;
            };

            let text: String = line.spans.iter().map(|span| span.text.as_str()).collect();
            if !visit(line_index, text) {
                return;
            }
            line_index += 1;
        }

        index += count;
    }
}

/// A place in a document a caret can be: a document line and a character of that line, which is
/// the shape `Selection`'s two ends already have. Named here because a caret move is a question
/// about the document rather than about the frame, and the page that asks it has a selection to
/// keep in step with the answer.
pub type Place = (usize, usize);

/// Every drawn line of `path`, as it is drawn: what a caret move is measured against, because a
/// line too long for the box is one document line however many rows it takes, and a move that
/// counted rows would be a move through text the document does not have — and what a search reads,
/// because a query is found where the frame draws it (see `find_all`).
///
/// It is built once per document and kept (`CachedDocument::drawn_text`), because the two things that
/// ask for it ask constantly: a caret is moved a key at a time and a query is searched a character at
/// a time, and styling a document for either would be a page of text through the highlighter per
/// keystroke. A handle rather than the lines themselves, so the caller that only reads them pays
/// nothing to ask.
#[cfg(test)]
pub fn drawn_lines(path: &Path, options: TextPreviewOptions) -> Arc<Vec<String>> {
    match document(path, options) {
        Some(document) => document.drawn_text(options),
        None => Arc::new(Vec::new()),
    }
}

/// How many times a document has been looked for on the disk, which is the one cost a key press on
/// a page already on screen must not pay (see `document_held`).
///
/// Counted rather than measured because the count is what a test can assert: a move that reads the
/// lines of the page's own document has no business asking the file system whether the file is
/// still the one it read, and the difference between a page that asks and one that does not is a
/// number that does not depend on how fast the disk is.
#[cfg(test)]
static DOCUMENT_STATS: AtomicU64 = AtomicU64::new(0);

/// The number of times the disk has been asked about a document (see `DOCUMENT_STATS`).
#[cfg(test)]
pub(crate) fn document_stat_count() -> u64 {
    DOCUMENT_STATS.load(Ordering::Relaxed)
}

/// The document `path` was last read as, without asking the disk whether it has changed since.
///
/// **A page on screen already has its document in hand**: it was read to draw the page, and every
/// question a reader asks of that page — a caret moved a line, a search typed a character, a
/// selection copied — is a question about the document as it is drawn. Asking the file system first
/// was a stat per key press, and a held arrow is a press every repeat: a trip out of the process for
/// an answer the cache already holds. A file that changes under a page is noticed the next time a
/// page is *put up*, which is where `document`'s stat is paid and where a stale document would
/// first be seen (see `render_scrolled` and `drawn_lines_held`).
///
/// The key is the path, the mode and the theme — everything that makes it a different document —
/// and not the file's size or time, which are what the stat was for.
pub(super) fn document_held(
    path: &Path,
    options: TextPreviewOptions,
) -> Option<Arc<CachedDocument>> {
    let theme_generation = match options.theme {
        TextTheme::Custom(_) => text_theme::generation(),
        _ => 0,
    };

    let cache = DOCUMENTS.lock().ok()?;
    cache
        .iter()
        .rev()
        .find(|(key, _)| {
            key.path == path
                && key.theme == options.theme
                && key.theme_generation == theme_generation
                && key.markdown_mode == options.markdown_mode
        })
        .map(|(_, document)| Arc::clone(document))
}

/// Lines the page is drawn from, read of the document the page was drawn from rather than of the
/// disk (see `document_held`). The stat is paid only where the cache holds nothing, which is a page
/// that has not been drawn yet rather than one a reader is moving through.
pub fn drawn_lines_held(path: &Path, options: TextPreviewOptions) -> Arc<Vec<String>> {
    match document_held(path, options).or_else(|| document(path, options)) {
        Some(document) => document.drawn_text(options),
        None => Arc::new(Vec::new()),
    }
}

/// What a selection takes out of the document of `path`: the lines it covers as they are drawn,
/// the two ends cut to the characters it names, joined the way Windows joins lines.
///
/// It is read of the document rather than of the frame that is showing it, which is what makes a
/// copy the text itself: a line too long for the box is copied whole and unbroken, a selection
/// reaching past the frame reaches past it here too, and a rendered Markdown document gives the
/// text it was rendered to rather than the source it was rendered from (see `Selection`).
///
/// `shortcut:` a copy slices the lines the page is drawn from, so it is the text a document already
/// holds rather than a second read of it; upgrade only if a copy is ever asked to span a document
/// further than a page can scroll (see `drawn_lines` and `MAX_DOC_LINES`).
pub fn text_of(path: &Path, options: TextPreviewOptions, selection: Selection) -> Option<String> {
    let lines = drawn_lines_held(path, options);
    if lines.is_empty() {
        return None;
    }

    Some(text_of_lines(&lines, selection))
}

/// The lines `selection` covers, taken from the lines the page is drawn from rather than styled
/// again: a copy is a reader asking for text they can already see (see `drawn_lines`).
fn text_of_lines(lines: &[String], selection: Selection) -> String {
    let ((start_line, start_char), (end_line, end_char)) = selection.ordered();
    let last = lines.len().saturating_sub(1);
    let start_line = start_line.min(last);
    // A selection that ends past the last line takes to the end of it. What ends a line selection
    // is the break at the end of the line, and a file whose last line has no break after it has
    // nowhere for a selection of that line to end but the end of the text it does have (see
    // `line_in`).
    let (end_line, end_char) = if end_line > last {
        (last, usize::MAX)
    } else {
        (end_line, end_char)
    };

    let mut taken: Vec<String> = Vec::new();
    for index in start_line..=end_line {
        let Some(text) = lines.get(index) else {
            continue;
        };
        let from = if index == start_line { start_char } else { 0 };
        let to = if index == end_line {
            end_char
        } else {
            usize::MAX
        };
        taken.push(
            text.chars()
                .skip(from)
                .take(to.saturating_sub(from))
                .collect(),
        );
    }

    taken.join("\r\n")
}

/// Where the next occurrence of what `selection` names is, as a selection of its own: the line it
/// was found on and the characters it covers, or nothing where the document holds no other one.
///
/// The document is searched from just after the selection to its own end and then from its
/// beginning back to where the search started, so a word with one occurrence in it is found once
/// rather than for ever, and nothing behind the selection is stepped over.
///
/// `shortcut:` what is searched for is the first line of the selection, so a range spanning several
/// lines is searched for by the line it starts on — which is what a double press takes, and what a
/// reader has when they ask. The search is exact and case-sensitive; upgrade when a case-blind or a
/// wrapped search is asked for.
pub fn find_next(
    path: &Path,
    options: TextPreviewOptions,
    selection: Selection,
) -> Option<Selection> {
    let lines = drawn_lines_held(path, options);
    let total = lines.len();
    if total == 0 {
        return None;
    }

    let ((start_line, start_char), (end_line, end_char)) = selection.ordered();
    let start_line = start_line.min(total - 1);
    let to = if start_line == end_line {
        end_char
    } else {
        usize::MAX
    };
    let needle: Vec<char> = lines
        .get(start_line)?
        .chars()
        .skip(start_char)
        .take(to.saturating_sub(start_char))
        .collect();
    if needle.is_empty() {
        return None;
    }

    // The character the search picks the document up at: just past the occurrence the selection
    // already names, so a wrap around the start line does not find it a second time.
    let resumed_at = start_char + needle.len();

    // The document is read in two passes over lines already in hand — from just past the selection
    // to its end, and then from its beginning back to where the search started — because styling a
    // document again to read its text is the cost the lines are kept to avoid (see `drawn_lines`).
    let search = |from_line: usize, to_line: usize, resumed: usize, stopped: usize| {
        for index in from_line..=to_line {
            let Some(text) = lines.get(index) else {
                continue;
            };
            let from = if index == resumed { resumed_at } else { 0 };
            let to = if index == stopped {
                start_char
            } else {
                usize::MAX
            };
            if let Some(at) = find_in(text, from, to, &needle) {
                return Some((index, at));
            }
        }
        None
    };

    let found = search(start_line, total - 1, start_line, usize::MAX)
        .or_else(|| search(0, start_line, usize::MAX, start_line));

    let (line, at) = found?;
    Some(Selection {
        anchor: (line, at),
        caret: (line, at + needle.len()),
    })
}

/// Every place `query` is found in the document, in the order the document is read: the line each
/// occurrence starts on and the characters it covers, left to right and without overlapping.
///
/// It sits beside [`find_next`] rather than in place of it, because the two are asked different
/// questions: that one takes what is selected as its needle and answers where the *next* occurrence
/// is, walking on past the one the selection names, while a find bar is asked to colour every place
/// at once and to say how many of them there are. Both read the drawn lines the frame is drawn
/// from, so a rendered Markdown document is searched as it is rendered rather than as it was
/// written (see `walk_drawn_lines`).
///
/// `case_sensitive` and `whole_word` are the bar's two switches: a case-blind search compares the
/// characters' own lower-case forms, and a whole-word one takes a match whose first character is a
/// word character only where no word character stands before it, and one whose last is a word
/// character only where no word character stands after it. A query that begins or ends with
/// punctuation therefore has no boundary to meet at that end, which is what lets a search for
/// `->` find it inside `a->b` — the punctuation is what the query is.
///
/// **A needle that spans a wrapped line is not found**, which is the same limit [`find_next`] has
/// and for the same reason: the lines are searched one at a time, and a wrapped line is still one
/// line of the document however many of the frame's rows it takes.
pub fn find_all(
    path: &Path,
    options: TextPreviewOptions,
    query: &str,
    case_sensitive: bool,
    whole_word: bool,
) -> Vec<Selection> {
    let needle: Vec<char> = query.chars().collect();
    if needle.is_empty() {
        return Vec::new();
    }

    let lines = drawn_lines_held(path, options);

    let mut found = Vec::new();
    for (index, text) in lines.iter().enumerate() {
        let characters: Vec<char> = text.chars().collect();
        let mut from = 0usize;

        while let Some(at) = find_in_line(&characters, from, &needle, case_sensitive, whole_word) {
            found.push(Selection {
                anchor: (index, at),
                caret: (index, at + needle.len()),
            });
            // Just past the occurrence, so two of them do not overlap: `aaaa` holds two `aa`, not
            // three.
            from = at + needle.len();
        }
    }

    found
}

/// Where `needle` next sits in one drawn line at or after `from`, as the character it starts at,
/// under the two switches a find bar carries (see [`find_all`]).
fn find_in_line(
    characters: &[char],
    from: usize,
    needle: &[char],
    case_sensitive: bool,
    whole_word: bool,
) -> Option<usize> {
    let length = needle.len();
    if length == 0 || characters.len() < length || from > characters.len() - length {
        return None;
    }

    (from..=characters.len() - length).find(|at| {
        let same = characters[*at..*at + length]
            .iter()
            .zip(needle)
            .all(|(found, wanted)| same_character(*found, *wanted, case_sensitive));

        same && (!whole_word || word_boundaries_hold(characters, *at, needle))
    })
}

/// Whether a match at `at` meets the word boundaries a whole-word search asks for: a query that
/// begins with a word character has to be preceded by one that is not a word character, and one
/// that ends with a word character has to be followed by one that is not. Punctuation at either
/// end of the query asks for nothing at that end (see [`find_all`]).
fn word_boundaries_hold(characters: &[char], at: usize, needle: &[char]) -> bool {
    let last = needle.len() - 1;
    let starts_a_word = is_word_character(needle[0]);
    let ends_a_word = is_word_character(needle[last]);

    let before_is_a_word = at
        .checked_sub(1)
        .is_some_and(|before| is_word_character(characters[before]));
    let after_is_a_word = characters
        .get(at + needle.len())
        .is_some_and(|after| is_word_character(*after));

    (!starts_a_word || !before_is_a_word) && (!ends_a_word || !after_is_a_word)
}

/// Whether two characters are the same one to a search: the character itself, or the same character
/// in another case where the search is not to mind case. A character whose lower-case form is more
/// than one character is compared as the whole of that form, so a search never matches half of one
/// character's case folding.
fn same_character(found: char, wanted: char, case_sensitive: bool) -> bool {
    found == wanted || (!case_sensitive && found.to_lowercase().eq(wanted.to_lowercase()))
}

/// Where `needle` sits in one drawn line, between the two character bounds, if it does.
fn find_in(text: &str, from: usize, to: usize, needle: &[char]) -> Option<usize> {
    let characters: Vec<char> = text.chars().collect();
    let from = from.min(characters.len());
    let to = to.min(characters.len()).max(from);

    characters[from..to]
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|at| from + at)
}

/// Styled lines covering `[first, first + count)` of the document, built on
/// demand and kept for the next preview in the same place.
///
/// The returned window may be wider than what was asked for; callers index it by
/// absolute line number through [`TextWindow::line`].
pub(super) fn styled_window(
    key: &DocKey,
    document: &CachedDocument,
    theme: &LoadedTheme,
    first: usize,
    count: usize,
) -> Option<Arc<TextWindow>> {
    // A window is only ever as long as the document is: what the parser can produce past its
    // last line is nothing, and a request kept at its own length is a window the table can
    // never answer for (see `covers`).
    let count = count.min(document.doc.total_lines().saturating_sub(first));

    if count == 0 {
        return Some(Arc::new(TextWindow {
            first,
            lines: Vec::new(),
        }));
    }

    let mut cache = document.windows.lock().ok()?;
    let last = first + count;

    if let Some(window) = cache
        .windows
        .iter()
        .rev()
        .find(|window| window.covers(first, count))
    {
        return Some(Arc::clone(window));
    }

    let window = match &document.doc.producer {
        Producer::Highlighted { syntax } => {
            window_highlighted(&document.doc, key, syntax, theme, first, last)
        }
        Producer::Markdown => window_markdown(&document.doc, theme, first, last),
        // A line of plain text or ANSI art is styled on its own, so only the
        // window's own lines are touched.
        Producer::Plain => window_plain(&document.doc, theme, first, last),
        Producer::Ansi => window_ansi(&document.doc, theme, first, last),
    };

    let window = Arc::new(window);
    if cache.windows.len() >= WINDOW_CACHE_MAX_ENTRIES {
        cache.windows.remove(0);
    }
    cache.windows.push(Arc::clone(&window));

    Some(window)
}

fn build_document(path: &Path, options: TextPreviewOptions) -> Option<TextDoc> {
    let source = read_source(path)?;
    if source.text.trim().is_empty() {
        return None;
    }

    let extension = extension_of(path);

    // An RTF is markup, not text: it is reduced to the text it carries once, and
    // from there it is an ordinary plain document.
    if extension == "rtf" {
        let stripped = strip_rtf(&source.text).join("\n");
        return Some(TextDoc {
            line_starts: line_starts(&stripped),
            rendered_lines: None,
            text: stripped,
            producer: Producer::Plain,
            read_truncated: source.truncated,
        });
    }

    if is_markdown_extension(&extension) && options.markdown_mode == MarkdownMode::Rendered {
        let mut document = TextDoc {
            line_starts: line_starts(&source.text),
            rendered_lines: None,
            producer: Producer::Markdown,
            text: source.text,
            read_truncated: source.truncated,
        };

        // The walk that renders a document is also the only thing that can say
        // how many lines it draws, so it is run once here, before anything asks
        // to scroll: a scroll position past that count would be a frame with
        // nothing in it.
        let theme = text_theme::loaded(options.theme)?;
        document.rendered_lines = Some(rendered_line_count(&document, theme));
        return Some(document);
    }

    // NFO files carry ANSI color codes around plain text, so they are built from
    // their own escapes rather than from a syntax definition.
    if wants_ansi(&extension) {
        return Some(TextDoc {
            line_starts: line_starts(&source.text),
            rendered_lines: None,
            producer: Producer::Ansi,
            text: source.text,
            read_truncated: source.truncated,
        });
    }

    let syntax_set = &*SYNTAXES;
    let syntax = syntax_set
        .find_syntax_for_file(path)
        .ok()
        .flatten()
        .unwrap_or_else(|| syntax_set.find_syntax_plain_text());

    Some(TextDoc {
        line_starts: line_starts(&source.text),
        rendered_lines: None,
        producer: Producer::Highlighted { syntax },
        text: source.text,
        read_truncated: source.truncated,
    })
}

/// The byte offset every line starts at. One pass over the text, no parsing: this
/// is what a preview uses to know how much there is without coloring any of it.
fn line_starts(text: &str) -> Vec<usize> {
    let mut starts = vec![0usize];
    for (offset, byte) in text.bytes().enumerate() {
        if byte == b'\n' && offset + 1 < text.len() {
            starts.push(offset + 1);
        }
    }
    starts
}

/// A window of a source file: the parser runs from wherever it left off — or from
/// the nearest checkpoint — and only the requested lines are kept.
///
/// This is what makes a preview of a large file cheap: nothing past the window is
/// ever styled, and the state that continues the parse is kept instead of the
/// lines it produced.
fn window_highlighted(
    doc: &TextDoc,
    key: &DocKey,
    syntax: &'static SyntaxReference,
    theme: &LoadedTheme,
    first: usize,
    last: usize,
) -> TextWindow {
    let syntax_set = &*SYNTAXES;
    let (start, highlight_state, parse_state) = SYNTAX_PROGRESS.with(|cell| {
        let mut map = cell.borrow_mut();
        if map.len() > SYNTAX_PROGRESS_MAX_ENTRIES {
            map.clear();
        }

        let progress = map.entry(key.clone()).or_insert_with(|| {
            let (highlight_state, parse_state) = HighlightLines::new(syntax, theme.theme()).state();
            SyntaxProgress {
                frontier: 0,
                frontier_state: (highlight_state, parse_state),
                checkpoints: Vec::new(),
            }
        });

        progress.continuation(first)
    });

    let mut highlighter = HighlightLines::from_state(theme.theme(), highlight_state, parse_state);

    let mut lines = Vec::with_capacity(last.saturating_sub(first));
    let mut checkpoints: Vec<(usize, HighlightState, ParseState)> = Vec::new();
    let mut line_index = start;
    while line_index < last && line_index < doc.total_lines() {
        let spans = highlight_line(&mut highlighter, doc.line(line_index), syntax_set);
        if line_index >= first {
            lines.push(DocLine {
                spans,
                block: LineBlock::Plain,
                indent: 0,
            });
        }

        line_index += 1;
        if line_index % CHECKPOINT_INTERVAL_LINES == 0 {
            let (highlight_state, parse_state) = highlighter.state();
            checkpoints.push((line_index, highlight_state.clone(), parse_state.clone()));
            // Reading the state out consumes the highlighter, so it is rebuilt
            // from the same state to carry on.
            highlighter = HighlightLines::from_state(theme.theme(), highlight_state, parse_state);
        }
    }

    let (highlight_state, parse_state) = highlighter.state();
    SYNTAX_PROGRESS.with(|cell| {
        if let Ok(mut map) = cell.try_borrow_mut() {
            if let Some(progress) = map.get_mut(key) {
                for (line, highlight_state, parse_state) in checkpoints {
                    remember_checkpoint(progress, line, highlight_state, parse_state);
                }
                if line_index > progress.frontier {
                    progress.frontier = line_index;
                    progress.frontier_state = (highlight_state, parse_state);
                }
            }
        }
    });

    TextWindow { first, lines }
}

fn remember_checkpoint(
    progress: &mut SyntaxProgress,
    line: usize,
    highlight_state: HighlightState,
    parse_state: ParseState,
) {
    if progress.checkpoints.len() >= CHECKPOINT_MAX_ENTRIES {
        progress.checkpoints.remove(0);
    }
    progress
        .checkpoints
        .push((line, highlight_state, parse_state));
}

/// Plain lines: one style, and only the window's own lines are built.
fn window_plain(doc: &TextDoc, theme: &LoadedTheme, first: usize, last: usize) -> TextWindow {
    let style = text_style(&theme.style_for_scopes(&["source"]), BODY_LEVEL);
    let mut lines = Vec::with_capacity(last.saturating_sub(first));

    for index in first..last.min(doc.total_lines()) {
        lines.push(DocLine {
            spans: vec![Span {
                text: doc.line(index).to_string(),
                style,
            }],
            block: LineBlock::Plain,
            indent: 0,
        });
    }

    TextWindow { first, lines }
}

/// ANSI art: the escapes are consumed per line, so the window is self-contained.
fn window_ansi(doc: &TextDoc, theme: &LoadedTheme, first: usize, last: usize) -> TextWindow {
    let palette = AnsiPalette::new(theme);
    let base = text_style(&theme.style_for_scopes(&["source"]), BODY_LEVEL);
    let mut lines = Vec::with_capacity(last.saturating_sub(first));

    for index in first..last.min(doc.total_lines()) {
        let mut spans = Vec::new();
        push_ansi_spans(&mut spans, doc.line(index), base, &palette);
        lines.push(DocLine {
            spans,
            block: LineBlock::Plain,
            indent: 0,
        });
    }

    TextWindow { first, lines }
}

/// One line through syntect. A grammar that cannot handle the line still shows
/// it, in the theme's plain color, rather than dropping it.
pub(super) fn highlight_line(
    highlighter: &mut HighlightLines<'_>,
    line: &str,
    syntax_set: &SyntaxSet,
) -> Vec<Span> {
    let Ok(regions) = highlighter.highlight_line(line, syntax_set) else {
        return Vec::new();
    };

    regions
        .into_iter()
        .filter(|(_, text)| !text.is_empty())
        .map(|(style, text)| Span {
            text: text.to_string(),
            style: text_style(&style, BODY_LEVEL),
        })
        .collect()
}

fn is_markdown_extension(extension: &str) -> bool {
    matches!(extension, "md" | "markdown" | "mkd" | "mdown" | "mdx")
}

pub(super) fn wants_ansi(extension: &str) -> bool {
    matches!(extension, "nfo" | "ans" | "asc" | "diz")
}

/// The extension `path` is previewed by, lowercased so `.MD` and `.md` are one
/// extension to every rule that reads it.
pub(super) fn extension_of(path: &Path) -> String {
    path.extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_lowercase()
}

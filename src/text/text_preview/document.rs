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

use super::frame::TextPreviewOptions;
use super::markdown::{rendered_line_count, window_markdown};
use super::source::{push_ansi_spans, read_source, strip_rtf, AnsiPalette};
use crate::config::config::{MarkdownMode, TextTheme};
use crate::text::text_paint::{text_style, TextStyle, BODY_LEVEL};
use crate::text::text_theme::{self, LoadedTheme};
use once_cell::sync::Lazy;
use std::cell::RefCell;
use std::collections::HashMap;
use std::path::{Path, PathBuf};
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
    });

    if let Ok(mut cache) = DOCUMENTS.lock() {
        if cache.len() >= DOC_CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.push((key, Arc::clone(&document)));
    }

    Some(document)
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

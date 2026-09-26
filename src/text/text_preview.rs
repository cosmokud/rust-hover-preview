//! Text previews: one screenful of a text file, colored by the syntax
//! definition its extension names.
//!
//! The file is read once, decoded, turned into styled lines, laid out for the
//! box the caller asks for, and painted with GDI into a BGRA frame — the same
//! frame shape a decoded image arrives in, so everything downstream (the layered
//! surface, the spinner, the hover generation check) is unchanged.
//!
//! Two calls share that work. `measure` answers "how big is this preview?"
//! before anything is painted, the way a PDF page's size is resolved before the
//! window is placed; `render` then lays the same document out for the box the
//! layout settled on and fills it. The parsed document is cached per file, mode
//! and theme, so a second hover, a theme switch or a repaint costs a layout
//! instead of a parse.

use crate::config::config::{MarkdownMode, TextTheme};
use crate::text::text_paint::{
    blend, fill_rect, plain_style, readable, rgb, scaled, text_style, DibSurface, RunPainter,
    TextMetrics, TextStyle, BODY_LEVEL, SIZE_LEVELS,
};
use crate::text::text_theme::{self, LoadedTheme};
use once_cell::sync::Lazy;
use pulldown_cmark::{CodeBlockKind, Event, Options, Parser, Tag, TagEnd};
use std::cell::RefCell;
use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};
use std::time::SystemTime;
use syntect::easy::HighlightLines;
use syntect::highlighting::HighlightState;
use syntect::parsing::{ParseState, SyntaxReference, SyntaxSet};
use windows::Win32::Foundation::RECT;
use windows::Win32::Graphics::Gdi::{CreateCompatibleDC, DeleteDC};

/// Bytes read from the file. A preview shows one screenful, and this bounds the
/// work a hover can trigger on a file that happens to be enormous.
const READ_LIMIT_BYTES: u64 = 2 * 1024 * 1024;

/// Lines a preview can scroll through. A source file is styled a window at a
/// time, so this is not a memory bound but a bound on how far one hover can walk
/// into a file — a few dozen screens of text.
const MAX_DOC_LINES: usize = 2000;

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

const TAB_WIDTH: usize = 4;

/// Parsed documents kept in memory, keyed by file and rendering options.
const DOC_CACHE_MAX_ENTRIES: usize = 24;

/// Shortest the thumb gets, so a document of a thousand screens still has
/// something to grab.
const SCROLLBAR_MIN_THUMB_PIXELS: i32 = 24;

/// Narrow files still get a window wide enough to look like one.
const MIN_CONTENT_CHARS: i32 = 16;

/// The options a preview is built with. They are passed in rather than read from
/// the configuration inside this module so the cache key and the caller's intent
/// cannot drift apart.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextPreviewOptions {
    pub theme: TextTheme,
    pub markdown_mode: MarkdownMode,
    /// Font scale as a percentage of the default text size. It sizes the glyphs,
    /// the line spacing and the page margin together, and it is deliberately not
    /// part of what a parsed document is cached by: the same styled lines are
    /// drawn at any size.
    pub font_scale_percent: u32,
    /// Whether the preview is more than something to look at: it scrolls when the
    /// document is longer than the frame, its text can be selected and copied, and
    /// it keeps the room the scrollbar needs. With it off a frame is a frame: the
    /// text a box cannot hold is said in a line, and nothing responds to a pointer.
    pub full_mode: bool,
}

/// A vertical scrollbar inside a frame, in frame coordinates: the groove it runs
/// in and the thumb that shows where the preview is.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ScrollBar {
    pub track: (i32, i32, i32, i32),
    pub thumb: (i32, i32, i32, i32),
}

/// One run of text as it was painted, with where it landed. The frame keeps these
/// so that a point on the preview can be turned back into a place in the text, and
/// so that a selection has something to be measured against and copied from.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FrameRun {
    /// Left edge, in frame coordinates.
    pub x: i32,
    pub width: i32,
    /// The characters the run shows, which is what a selection counts in.
    pub text: String,
}

/// One painted line, in frame coordinates.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FrameLine {
    pub top: i32,
    pub height: i32,
    pub runs: Vec<FrameRun>,
}

/// A range of a text preview a reader has selected: where the drag started and
/// where it is now, both as (frame line, character).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Selection {
    pub anchor: (usize, usize),
    pub caret: (usize, usize),
}

impl Selection {
    /// The selection in reading order, so a drag that runs backwards selects the
    /// same text as one that runs forwards.
    pub fn ordered(self) -> ((usize, usize), (usize, usize)) {
        if self.anchor <= self.caret {
            (self.anchor, self.caret)
        } else {
            (self.caret, self.anchor)
        }
    }

    pub fn is_empty(self) -> bool {
        self.anchor == self.caret
    }
}

/// One painted text preview: its pixels, what it took to lay it out, and where the
/// text landed.
#[derive(Clone)]
pub struct TextFrame {
    pub pixels: Vec<u8>,
    pub width: u32,
    pub height: u32,
    /// Document line the frame starts at.
    pub first_line: usize,
    /// Lines the frame shows.
    pub visible_lines: usize,
    /// Lines the preview can reach, which is why it scrolls at all, and the range
    /// its scrollbar is drawn in.
    pub scrollable_lines: usize,
    /// Present when the document is longer than the frame, which is also the only
    /// case in which the preview scrolls.
    pub scrollbar: Option<ScrollBar>,
    /// The painted lines, in frame coordinates.
    pub lines: Vec<FrameLine>,
}

/// Where a point falls in painted lines: the line under it and the character
/// within that line.
///
/// A point between two lines takes the nearer one, and a point past the end of a
/// line takes the nearest end of it, so a drag that runs off the text still selects
/// the line it was on rather than losing it.
pub fn position_in(lines: &[FrameLine], x: i32, y: i32) -> (usize, usize) {
    let Some((line_index, line)) = lines.iter().enumerate().min_by_key(|(_, line)| {
        let middle = line.top + line.height / 2;
        (middle - y).abs()
    }) else {
        return (0, 0);
    };

    let mut characters = 0usize;
    let mut closest = 0usize;
    let mut closest_distance = i32::MAX;

    for run in &line.runs {
        let count = run.text.chars().count() as i32;
        if count == 0 {
            continue;
        }
        let advance = (run.width as f32 / count as f32).max(1.0);
        let within = (((x - run.x) as f32 / advance).round() as i32).clamp(0, count);
        let edge = run.x + (within as f32 * advance).round() as i32;
        let distance = (edge - x).abs();
        if distance < closest_distance {
            closest_distance = distance;
            closest = characters + within as usize;
        }
        characters += count as usize;
    }

    (line_index, closest)
}

/// The text of painted lines, either everything they show or the part a selection
/// covers, as it will be pasted: the lines it touches, each cut to the part that
/// was selected, joined with Windows line endings.
///
/// What painted lines hold is what the frame shows, so a copy is the part of the
/// file that was on screen — a line that wrapped is copied as the lines it took.
pub fn text_in(lines: &[FrameLine], selection: Selection) -> String {
    let ((start_line, start_char), (end_line, end_char)) = selection.ordered();
    let mut selected: Vec<String> = Vec::new();

    for (index, line) in lines.iter().enumerate() {
        if index < start_line || index > end_line {
            continue;
        }

        let from = if index == start_line { start_char } else { 0 };
        let to = if index == end_line {
            end_char
        } else {
            usize::MAX
        };
        let mut characters = String::new();
        let mut position = 0usize;

        for run in &line.runs {
            for character in run.text.chars() {
                if position >= from && position < to {
                    characters.push(character);
                }
                position += 1;
            }
        }

        selected.push(characters);
    }

    selected.join("\r\n")
}

/// Everything the painted lines show, which is what a Copy with nothing selected
/// takes.
pub fn frame_text(lines: &[FrameLine]) -> String {
    text_in(
        lines,
        Selection {
            anchor: (0, 0),
            caret: (usize::MAX, usize::MAX),
        },
    )
}

/// The size the preview of `path` wants inside `max_width` x `max_height`.
///
/// The box hugs the content — a two-line file gets a two-line window — and is
/// capped by the space the caller offers, so the preview is never clipped. A
/// file with nothing to show (empty, binary, unreadable) reports no size, which
/// drops the preview instead of opening an empty page.
pub fn measure(
    path: &Path,
    max_width: u32,
    max_height: u32,
    dpi: u32,
    options: TextPreviewOptions,
) -> Option<(u32, u32)> {
    let document = document(path, options)?;
    let theme = text_theme::loaded(options.theme)?;

    let dc = unsafe { CreateCompatibleDC(None) };
    if dc.0.is_null() {
        return None;
    }

    let result = TextMetrics::new(dc, dpi, options.font_scale_percent).map(|metrics| {
        let laid_out = layout(
            &document,
            theme,
            &metrics,
            0,
            max_width,
            max_height,
            0,
            options.full_mode,
        );
        (laid_out.width, laid_out.height)
    });

    unsafe {
        let _ = DeleteDC(dc);
    }

    result
}

/// Paint the preview at a scroll position, together with the scrollbar the frame
/// needs and the numbers that describe where the preview is.
///
/// `first_line` is a request rather than a command: a preview of a document that
/// is nearly over is pulled back so the frame is still full, and a preview that
/// fits is always at line 0. What comes back is what was actually painted, so the
/// caller can scroll from there without guessing.
///
/// Nothing of what was painted is held. Laying a screenful out and drawing every run of it is
/// the whole of what a text hover costs, and that cost is small next to what stands behind it:
/// the document is read once and kept, and the lines that are on screen are styled once and
/// kept per document and window (`DOCUMENTS` and `WindowCache`), which is what makes a second
/// hover a layout instead of a parse. A painted frame is a screenful of pixels — the dearest
/// thing here to hold and the cheapest of the three to make again.
pub fn render_scrolled(
    path: &Path,
    first_line: usize,
    width: u32,
    height: u32,
    dpi: u32,
    options: TextPreviewOptions,
    selection: Option<Selection>,
) -> Option<TextFrame> {
    if width == 0 || height == 0 {
        return None;
    }

    let document = document(path, options)?;
    let theme = text_theme::loaded(options.theme)?;

    let surface = DibSurface::create(width, height)?;
    let metrics = TextMetrics::new(surface.dc, dpi, options.font_scale_percent)?;
    let scrollable_lines = document.doc.scrollable_lines();

    // Whether there is a scrollbar depends on how many lines fit, and the bar
    // takes room the text would otherwise use — so the layout is asked once to
    // find that out and again with the room kept aside.
    let mut laid_out = layout(
        &document,
        theme,
        &metrics,
        first_line,
        width,
        height,
        0,
        options.full_mode,
    );
    let scrollable = options.full_mode && scrollable_lines > laid_out.visible_lines;
    if scrollable {
        laid_out = layout(
            &document,
            theme,
            &metrics,
            first_line,
            width,
            height,
            metrics.scrollbar_space(),
            options.full_mode,
        );
    }

    let scrollbar = if scrollable {
        scrollbar_geometry(
            width as i32,
            height as i32,
            &metrics,
            laid_out.first_line,
            laid_out.visible_lines,
            scrollable_lines,
        )
    } else {
        None
    };

    unsafe {
        paint(
            &surface,
            &laid_out,
            theme,
            &metrics,
            scrollbar.as_ref(),
            selection,
        );
    }

    let lines = laid_out
        .lines
        .iter()
        .map(|line| FrameLine {
            top: line.top,
            height: line.height,
            runs: line
                .runs
                .iter()
                .map(|run| FrameRun {
                    x: run.x,
                    width: run.width,
                    text: run.text.clone(),
                })
                .collect(),
        })
        .collect();

    let frame = TextFrame {
        pixels: surface.pixels(),
        width,
        height,
        first_line: laid_out.first_line,
        visible_lines: laid_out.visible_lines,
        scrollable_lines,
        scrollbar,
        lines,
    };

    Some(frame)
}

// ---------------------------------------------------------------- documents

/// A text file as it is read: the decoded text and where its lines start.
///
/// Styling is deliberately not part of this. A line table costs eight bytes per
/// line and no parsing, which is what lets a preview know how tall a document is
/// — and how far it can scroll — before anything has been colored. The styled
/// lines for the part of it that is on screen are built on demand, in
/// [`styled_window`].
struct TextDoc {
    text: String,
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
    read_truncated: bool,
}

impl TextDoc {
    /// Lines there are to show, in the space a frame is drawn in.
    fn total_lines(&self) -> usize {
        self.rendered_lines.unwrap_or(self.line_starts.len())
    }

    /// Lines this preview can reach. A source file is styled a window at a time,
    /// but the range a preview can scroll through is capped so a hover can never
    /// walk an unbounded file.
    fn scrollable_lines(&self) -> usize {
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
struct TextWindow {
    first: usize,
    lines: Vec<DocLine>,
}

impl TextWindow {
    fn line(&self, index: usize) -> Option<&DocLine> {
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
struct CachedDocument {
    /// What this document was built from, which is also how the per-thread parse
    /// progress finds it again.
    key: DocKey,
    doc: TextDoc,
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
struct DocKey {
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
static SYNTAXES: Lazy<SyntaxSet> = Lazy::new(two_face::syntax::extra_no_newlines);

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
enum LineBlock {
    #[default]
    Plain,
    /// A line of a fenced or indented code block.
    Code,
    /// A line inside a block quote.
    Quote,
}

#[derive(Default, Clone)]
struct DocLine {
    spans: Vec<Span>,
    block: LineBlock,
    /// Extra left indent, in characters, for nested lists and quotes.
    indent: u8,
}

#[derive(Clone)]
struct Span {
    text: String,
    style: TextStyle,
}

/// The parsed document for `path`, from the cache when the file, its mode and
/// its theme are unchanged.
fn document(path: &Path, options: TextPreviewOptions) -> Option<Arc<CachedDocument>> {
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
fn styled_window(
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
fn highlight_line(
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

fn wants_ansi(extension: &str) -> bool {
    matches!(extension, "nfo" | "ans" | "asc" | "diz")
}

/// The extension `path` is previewed by, lowercased so `.MD` and `.md` are one
/// extension to every rule that reads it.
fn extension_of(path: &Path) -> String {
    path.extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_lowercase()
}

// ------------------------------------------------------------------ markdown

/// Walk the CommonMark events once, turning them into styled lines.
///
/// Block structure becomes lines and indent, inline structure becomes runs, and
/// fenced code is handed to the same highlighter the rest of the preview uses,
/// so a fenced block is colored like the language it names.
///
/// A rendered document is the one producer that has to read all of its input: a
/// line's place in the document is only known once the block before it has been
/// walked, so there is no state to resume from part-way through. It is a single
/// linear pass with no grammar work, and only the lines inside the requested
/// window are kept — the rest are counted and dropped as they go past.
struct MarkdownBuilder<'a> {
    theme: &'a LoadedTheme,
    lines: Vec<DocLine>,
    /// Document line the builder has reached, and the window it keeps.
    emitted: usize,
    keep_from: usize,
    keep_to: usize,
    /// Whether the line before the current one was blank, which is what keeps
    /// block spacing down to one line without needing the lines themselves.
    last_blank: bool,
    current: DocLine,
    inline: Vec<Inline>,
    list_stack: Vec<Option<u64>>,
    quote_depth: usize,
    heading: Option<u8>,
    code_block: Option<Option<String>>,
    code_text: String,
    hit_line_cap: bool,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Inline {
    Emphasis,
    Strong,
    Strike,
    Link,
    Image,
}

impl<'a> MarkdownBuilder<'a> {
    fn new(theme: &'a LoadedTheme, keep_from: usize, keep_to: usize) -> Self {
        Self {
            theme,
            lines: Vec::new(),
            emitted: 0,
            keep_from,
            keep_to,
            last_blank: false,
            current: DocLine::default(),
            inline: Vec::new(),
            list_stack: Vec::new(),
            quote_depth: 0,
            heading: None,
            code_block: None,
            code_text: String::new(),
            hit_line_cap: false,
        }
    }

    fn band(&self) -> [u8; 3] {
        blend(
            rgb(self.theme.background()),
            rgb(self.theme.foreground()),
            0.08,
        )
    }

    fn scope_style(&self, scopes: &[&str], level: u8) -> TextStyle {
        let mut style = text_style(&self.theme.style_for_scopes(scopes), level);
        style.level = level;
        style
    }

    /// The style of the text being written right now: the block's own style plus
    /// whatever inline markup is open around it.
    fn current_style(&self) -> TextStyle {
        let mut style = if self.code_block.is_some() {
            self.scope_style(&["markup.raw", "markup.raw.block"], BODY_LEVEL)
        } else if let Some(level) = self.heading {
            let level_scopes = format!("markup.heading.{level}.markdown");
            let level = heading_size_level(level);
            let mut style = self.scope_style(&["markup.heading", level_scopes.as_str()], level);
            style.bold = true;
            style
        } else if self.quote_depth > 0 {
            self.scope_style(&["markup.quote"], BODY_LEVEL)
        } else {
            self.scope_style(&["source"], BODY_LEVEL)
        };

        if self.inline.contains(&Inline::Strong) {
            style.bold = true;
        }
        if self.inline.contains(&Inline::Emphasis) {
            style.italic = true;
        }
        if self.inline.contains(&Inline::Strike) {
            style.strike = true;
        }
        if self.inline.contains(&Inline::Link) {
            style.foreground = self
                .scope_style(&["markup.underline.link"], style.level)
                .foreground;
            style.underline = true;
        }
        if self.inline.contains(&Inline::Image) {
            style.italic = true;
        }

        style
    }

    fn line_indent(&self) -> u8 {
        (self.quote_depth as u8).saturating_add((self.list_stack.len() as u8).saturating_mul(2))
    }

    fn block_kind(&self) -> LineBlock {
        if self.code_block.is_some() {
            LineBlock::Code
        } else if self.quote_depth > 0 {
            LineBlock::Quote
        } else {
            LineBlock::Plain
        }
    }

    fn write(&mut self, text: &str, style: TextStyle) {
        if text.is_empty() {
            return;
        }
        if self.current.spans.is_empty() {
            self.current.block = self.block_kind();
            self.current.indent = self.line_indent();
        }
        self.current.spans.push(Span {
            text: text.to_string(),
            style,
        });
    }

    /// Push the line in progress. A line with no runs is dropped, which is what
    /// keeps tight list items from collecting empty ones.
    fn flush(&mut self) {
        if self.current.spans.is_empty() {
            self.current = DocLine::default();
            return;
        }

        let line = std::mem::take(&mut self.current);
        self.push_line(line);
        self.last_blank = false;
    }

    /// Count one document line, keeping it only when it falls inside the window.
    fn push_line(&mut self, line: DocLine) {
        if self.emitted >= MAX_DOC_LINES {
            self.hit_line_cap = true;
            return;
        }

        if self.emitted >= self.keep_from && self.emitted < self.keep_to {
            self.lines.push(line);
        }
        self.emitted += 1;
    }

    /// One empty line between blocks, never two.
    fn blank_line(&mut self) {
        self.flush();
        if self.emitted == 0 || self.last_blank {
            return;
        }

        self.push_line(DocLine::default());
        self.last_blank = true;
    }

    fn rule_line(&mut self) {
        self.blank_line();
        self.push_line(DocLine::default());
        self.last_blank = true;
    }

    fn item_marker(&mut self) {
        let depth = self.list_stack.len().saturating_sub(1);
        let marker = match self.list_stack.last().copied().flatten() {
            Some(number) => format!("{number}. "),
            None => "• ".to_string(),
        };
        let style = self.scope_style(&["markup.list"], BODY_LEVEL);

        self.current.indent = (self.quote_depth as u8).saturating_add((depth as u8) * 2);
        self.current.block = self.block_kind();
        self.write(&marker, style);
    }

    fn advance_list(&mut self) {
        if let Some(Some(number)) = self.list_stack.last_mut() {
            *number += 1;
        }
    }

    /// A fenced or indented block arrives as one text event; it is split into
    /// lines and highlighted with the syntax its info string names.
    fn end_code_block(&mut self) {
        let language = self.code_block.take().flatten();
        let code = std::mem::take(&mut self.code_text);
        let mut lines: Vec<String> = code.split('\n').map(|line| line.to_string()).collect();
        while lines.last().map(|line| line.is_empty()).unwrap_or(false) {
            lines.pop();
        }

        let syntax_set = &*SYNTAXES;
        let syntax = language
            .as_deref()
            .and_then(|token| syntax_set.find_syntax_by_token(token))
            .unwrap_or_else(|| syntax_set.find_syntax_plain_text());
        let mut highlighter = HighlightLines::new(syntax, self.theme.theme());
        let indent = self.line_indent().saturating_add(1);

        for line in &lines {
            let styled = DocLine {
                spans: highlight_line(&mut highlighter, line, syntax_set),
                block: LineBlock::Code,
                indent,
            };
            self.push_line(styled);
            if self.hit_line_cap {
                return;
            }
        }

        self.last_blank = lines.is_empty();
        self.current = DocLine::default();
    }

    /// The window the walk produced, with the line number it starts at.
    fn finish(mut self) -> TextWindow {
        self.flush();

        // A blank line at the very end of the document is padding rather than
        // content, but one in the middle of a window is a paragraph break.
        if !self.hit_line_cap {
            while self
                .lines
                .last()
                .map(|line| line.spans.is_empty())
                .unwrap_or(false)
            {
                self.lines.pop();
            }
        }

        TextWindow {
            first: self.keep_from,
            lines: self.lines,
        }
    }
}

fn heading_size_level(level: u8) -> u8 {
    match level {
        1 => 1,
        2 => 2,
        3 => 3,
        _ => 4,
    }
}

/// A window of a rendered Markdown document. The walk is linear and cannot be
/// resumed, so it runs over the whole text and keeps only the requested lines.
fn window_markdown(doc: &TextDoc, theme: &LoadedTheme, first: usize, last: usize) -> TextWindow {
    let mut builder = MarkdownBuilder::new(theme, first, last);
    walk_markdown(doc, &mut builder);
    builder.finish()
}

/// How many lines the rendered document has.
///
/// The same walk a window runs, keeping nothing: it is the only way to know, and
/// it is what a preview's scroll range is measured in — a paragraph is one line
/// and the blank line between two blocks is one the source does not have, so the
/// file's own line count is not the one a frame is drawn in.
fn rendered_line_count(doc: &TextDoc, theme: &LoadedTheme) -> usize {
    // Every line is kept, and the walk finishes the way a window does: a blank
    // line at the end of a document is padding rather than content, and counting
    // it would leave the preview one line to scroll into with nothing in it.
    let mut builder = MarkdownBuilder::new(theme, 0, usize::MAX);
    walk_markdown(doc, &mut builder);
    builder.finish().lines.len()
}

/// Walk the document's events into a builder, which keeps the lines it was asked
/// for and counts all of them.
fn walk_markdown(doc: &TextDoc, builder: &mut MarkdownBuilder<'_>) {
    let options =
        Options::ENABLE_TABLES | Options::ENABLE_STRIKETHROUGH | Options::ENABLE_TASKLISTS;

    for event in Parser::new_ext(&doc.text, options) {
        match event {
            Event::Start(tag) => match tag {
                Tag::Paragraph => {}
                Tag::Heading { level, .. } => builder.heading = Some(level as u8),
                Tag::BlockQuote => {
                    if builder.quote_depth == 0 {
                        builder.blank_line();
                    }
                    builder.quote_depth += 1;
                }
                Tag::CodeBlock(kind) => {
                    builder.blank_line();
                    let language = match kind {
                        CodeBlockKind::Fenced(info) => {
                            let token = info.split_whitespace().next().unwrap_or("").to_string();
                            (!token.is_empty()).then_some(token)
                        }
                        CodeBlockKind::Indented => None,
                    };
                    builder.code_block = Some(language);
                }
                Tag::List(start) => {
                    if builder.list_stack.is_empty() {
                        builder.blank_line();
                    }
                    builder.list_stack.push(start);
                }
                Tag::Item => {
                    builder.flush();
                    builder.item_marker();
                }
                Tag::TableCell => {
                    if !builder.current.spans.is_empty() {
                        let style = builder.current_style();
                        builder.write(" | ", style);
                    }
                }
                Tag::Emphasis => builder.inline.push(Inline::Emphasis),
                Tag::Strong => builder.inline.push(Inline::Strong),
                Tag::Strikethrough => builder.inline.push(Inline::Strike),
                Tag::Link { .. } => builder.inline.push(Inline::Link),
                Tag::Image { .. } => {
                    builder.inline.push(Inline::Image);
                    let style = builder.current_style();
                    builder.write("[", style);
                }
                _ => {}
            },
            Event::End(tag) => match tag {
                TagEnd::Paragraph => builder.flush(),
                TagEnd::Heading(_) => {
                    builder.heading = None;
                    builder.blank_line();
                }
                TagEnd::BlockQuote => {
                    builder.quote_depth = builder.quote_depth.saturating_sub(1);
                    builder.blank_line();
                }
                TagEnd::CodeBlock => builder.end_code_block(),
                TagEnd::List(_) => {
                    builder.list_stack.pop();
                    if builder.list_stack.is_empty() {
                        builder.blank_line();
                    }
                }
                TagEnd::Item => {
                    builder.flush();
                    builder.advance_list();
                }
                TagEnd::Table => builder.blank_line(),
                TagEnd::TableRow => builder.flush(),
                TagEnd::Emphasis => builder.inline.retain(|item| *item != Inline::Emphasis),
                TagEnd::Strong => builder.inline.retain(|item| *item != Inline::Strong),
                TagEnd::Strikethrough => builder.inline.retain(|item| *item != Inline::Strike),
                TagEnd::Link => builder.inline.retain(|item| *item != Inline::Link),
                TagEnd::Image => {
                    builder.inline.retain(|item| *item != Inline::Image);
                    let style = builder.current_style();
                    builder.write("]", style);
                }
                _ => {}
            },
            Event::Text(text) => {
                if builder.code_block.is_some() {
                    builder.code_text.push_str(&text);
                } else {
                    let style = builder.current_style();
                    builder.write(&text, style);
                }
            }
            Event::Code(text) => {
                let mut style =
                    builder.scope_style(&["markup.inline.raw", "markup.raw.inline"], BODY_LEVEL);
                style.background = Some(builder.band());
                builder.write(&text, style);
            }
            Event::SoftBreak => {
                let style = builder.current_style();
                builder.write(" ", style);
            }
            Event::HardBreak => builder.flush(),
            Event::Rule => builder.rule_line(),
            Event::TaskListMarker(checked) => {
                let marker = if checked { "[x] " } else { "[ ] " };
                let style = builder.scope_style(&["markup.list"], BODY_LEVEL);
                builder.write(marker, style);
            }
            _ => {}
        }
    }
}

// -------------------------------------------------------------------- source

/// A decoded file with its tabs already expanded. Lines are located by offset in
/// `text` rather than copied out of it.
struct SourceText {
    text: String,
    truncated: bool,
}

fn read_source(path: &Path) -> Option<SourceText> {
    let file = File::open(path).ok()?;
    let mut bytes = Vec::new();
    let mut limited = file.take(READ_LIMIT_BYTES + 1);
    limited.read_to_end(&mut bytes).ok()?;

    let truncated = bytes.len() as u64 > READ_LIMIT_BYTES;
    bytes.truncate(READ_LIMIT_BYTES as usize);

    let text = decode_text(&bytes, legacy_encoding(&extension_of(path)), truncated)?;
    let text = expand_tabs(&text);

    Some(SourceText { text, truncated })
}

/// The single-byte encoding a file without a byte order mark is assumed to use.
/// Scene NFO art predates code pages being a settled thing, so those files get
/// the table they were drawn in.
#[derive(Clone, Copy)]
enum LegacyEncoding {
    Windows1252,
    Cp437,
}

fn legacy_encoding(extension: &str) -> LegacyEncoding {
    if wants_ansi(extension) {
        LegacyEncoding::Cp437
    } else {
        LegacyEncoding::Windows1252
    }
}

/// Decode a text file: byte order marks first, then UTF-8, then the legacy code
/// page. `None` means the bytes are not text at all, which is how a renamed
/// archive or executable stays out of the renderer.
fn decode_text(bytes: &[u8], legacy: LegacyEncoding, truncated: bool) -> Option<String> {
    if bytes.is_empty() {
        return Some(String::new());
    }

    if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
        return Some(String::from_utf8_lossy(&bytes[3..]).into_owned());
    }

    if bytes.starts_with(&[0xFF, 0xFE]) || bytes.starts_with(&[0xFE, 0xFF]) {
        return Some(decode_utf16(&bytes[2..], bytes[0] == 0xFE));
    }

    // NUL bytes are checked before the UTF-8 decode, because they are valid
    // UTF-8: a UTF-16 file without a mark and a renamed executable both read as
    // UTF-8 text holding NULs, and neither is a page to show. A UTF-16 file is
    // half NUL bytes on one side of every pair; anything else is binary.
    if bytes[..bytes.len().min(8192)].contains(&0) {
        return decode_utf16_without_mark(bytes);
    }

    if let Ok(text) = std::str::from_utf8(bytes) {
        // Valid UTF-8 is not the same as a page: bytes that decode into a mostly
        // control characters are a binary file that happened to pass. ANSI-art
        // extensions are exempt, since their own escapes are control characters.
        if !matches!(legacy, LegacyEncoding::Cp437) && is_mostly_control(text) {
            return None;
        }
        return Some(text.to_string());
    }

    // The byte cap can land inside a multi-byte character. Dropping the few
    // bytes a split character occupies keeps a long UTF-8 file from being
    // mistaken for a legacy-encoded one.
    if truncated {
        let head = &bytes[..bytes.len().saturating_sub(4)];
        if let Ok(text) = std::str::from_utf8(head) {
            return Some(text.to_string());
        }
    }

    Some(decode_legacy(bytes, legacy))
}

/// Whether a decoded page is mostly characters no page would show. Text that is
/// mostly control characters did not come from a text file, whatever its bytes
/// happened to decode as.
fn is_mostly_control(text: &str) -> bool {
    let probe: Vec<char> = text.chars().take(8192).collect();
    if probe.is_empty() {
        return false;
    }

    let controls = probe
        .iter()
        .filter(|character| character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
        .count();

    controls * 10 > probe.len()
}

fn decode_utf16(bytes: &[u8], big_endian: bool) -> String {
    let units: Vec<u16> = bytes
        .as_chunks::<2>()
        .0
        .iter()
        .map(|pair| {
            if big_endian {
                u16::from_be_bytes([pair[0], pair[1]])
            } else {
                u16::from_le_bytes([pair[0], pair[1]])
            }
        })
        .collect();

    String::from_utf16_lossy(&units)
}

/// UTF-16 with no mark, recognized from the NUL bytes an ASCII-heavy UTF-16 file
/// carries on one side of every pair. The decoded result is checked before it is
/// trusted, so bytes that merely happen to look like UTF-16 are still reported as
/// the binary they are.
fn decode_utf16_without_mark(bytes: &[u8]) -> Option<String> {
    let probe = &bytes[..bytes.len().min(4096) / 2 * 2];
    if probe.len() < 4 {
        return None;
    }

    let mut zeros_even = 0usize;
    let mut zeros_odd = 0usize;
    for (index, byte) in probe.iter().enumerate() {
        if *byte == 0 {
            if index % 2 == 0 {
                zeros_even += 1;
            } else {
                zeros_odd += 1;
            }
        }
    }

    let pairs = probe.len() / 2;
    let big_endian = zeros_even * 10 > pairs * 6;
    let little_endian = zeros_odd * 10 > pairs * 6;
    if !big_endian && !little_endian {
        return None;
    }

    let text = decode_utf16(bytes, big_endian);

    // A decoded page is made of characters a page can show. Control characters
    // mean the NUL pattern was a coincidence and the file is binary.
    let characters = text.chars().count();
    let controls = text
        .chars()
        .filter(|character| character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
        .count();

    (characters > 0 && controls * 10 <= characters).then_some(text)
}

/// Windows-1252 agrees with Latin-1 outside `0x80..=0x9F`, where it borrows
/// punctuation from its own table. CP437 is a different table over `0x80..=0xFF`.
fn decode_legacy(bytes: &[u8], encoding: LegacyEncoding) -> String {
    bytes
        .iter()
        .map(|byte| match encoding {
            LegacyEncoding::Windows1252 => cp1252_char(*byte),
            LegacyEncoding::Cp437 => cp437_char(*byte),
        })
        .collect()
}

fn cp1252_char(byte: u8) -> char {
    const HIGH: [char; 32] = [
        '\u{20AC}', '\u{FFFD}', '\u{201A}', '\u{0192}', '\u{201E}', '\u{2026}', '\u{2020}',
        '\u{2021}', '\u{02C6}', '\u{2030}', '\u{0160}', '\u{2039}', '\u{0152}', '\u{FFFD}',
        '\u{017D}', '\u{FFFD}', '\u{FFFD}', '\u{2018}', '\u{2019}', '\u{201C}', '\u{201D}',
        '\u{2022}', '\u{2013}', '\u{2014}', '\u{02DC}', '\u{2122}', '\u{0161}', '\u{203A}',
        '\u{0153}', '\u{FFFD}', '\u{017E}', '\u{0178}',
    ];

    match byte {
        0x00..=0x7F => byte as char,
        0x80..=0x9F => HIGH[(byte - 0x80) as usize],
        _ => byte as char,
    }
}

/// CP437's upper half, in code point order from `0x80`.
fn cp437_char(byte: u8) -> char {
    const HIGH: [char; 128] = [
        'Ç', 'ü', 'é', 'â', 'ä', 'à', 'å', 'ç', 'ê', 'ë', 'è', 'ï', 'î', 'ì', 'Ä', 'Å', 'É', 'æ',
        'Æ', 'ô', 'ö', 'ò', 'û', 'ù', 'ÿ', 'Ö', 'Ü', '¢', '£', '¥', '₧', 'ƒ', 'á', 'í', 'ó', 'ú',
        'ñ', 'Ñ', 'ª', 'º', '¿', '⌐', '¬', '½', '¼', '¡', '«', '»', '░', '▒', '▓', '│', '┤', '╡',
        '╢', '╖', '╕', '╣', '║', '╗', '╝', '╜', '╛', '┐', '└', '┴', '┬', '├', '─', '┼', '╞', '╟',
        '╚', '╔', '╩', '╦', '╠', '═', '╬', '╧', '╨', '╤', '╥', '╙', '╘', '╒', '╓', '╫', '╪', '┘',
        '┌', '█', '▄', '▌', '▐', '▀', 'α', 'ß', 'Γ', 'π', 'Σ', 'σ', 'µ', 'τ', 'Φ', 'Θ', 'Ω', 'δ',
        '∞', 'φ', 'ε', '∩', '≡', '±', '≥', '≤', '⌠', '⌡', '÷', '≈', '°', '∙', '·', '√', 'ⁿ', '²',
        '■', '\u{00A0}',
    ];

    match byte {
        0x00..=0x7F => byte as char,
        _ => HIGH[(byte - 0x80) as usize],
    }
}

/// Tabs are replaced with spaces before anything else looks at the text, so the
/// highlighter, the Markdown parser and the layout all see the same columns, and
/// no tab character ever reaches GDI's own tab stops. Carriage returns go with
/// them, which is what makes Windows and Unix line endings read the same.
fn expand_tabs(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    let mut column = 0usize;

    for character in text.chars() {
        match character {
            '\t' => {
                let spaces = TAB_WIDTH - (column % TAB_WIDTH);
                for _ in 0..spaces {
                    out.push(' ');
                }
                column += spaces;
            }
            '\n' => {
                out.push('\n');
                column = 0;
            }
            '\r' => {}
            _ => {
                out.push(character);
                column += 1;
            }
        }
    }

    out
}

// ----------------------------------------------------------------------- RTF

/// Destinations whose contents are not document text: font, color and style
/// tables, embedded pictures, field instructions and headers.
const RTF_IGNORED_DESTINATIONS: &[&str] = &[
    "fonttbl",
    "filetbl",
    "colortbl",
    "stylesheet",
    "listtable",
    "listoverridetable",
    "revtbl",
    "rsidtbl",
    "generator",
    "info",
    "pict",
    "object",
    "themedata",
    "datastore",
    "fldinst",
    "header",
    "headerl",
    "headerr",
    "headerf",
    "footer",
    "footerl",
    "footerr",
    "footerf",
    "footnote",
    "xmlnstbl",
];

/// Reduce RTF to the text it carries.
///
/// Word writes its tables ahead of the body, and those groups are skipped whole;
/// what is left is the document with its paragraphs, tabs, escaped bytes and
/// `\uN` sequences resolved. This is the text-stripping half of RTF and
/// deliberately not a layout engine: the preview shows the file's text, without
/// its formatting.
fn strip_rtf(source: &str) -> Vec<String> {
    let characters: Vec<char> = source.chars().collect();
    let mut lines: Vec<String> = Vec::new();
    let mut line = String::new();
    let mut index = 0usize;
    let mut depth = 0usize;
    let mut ignoring: Option<usize> = None;
    let mut unicode_skip = 1usize;

    while index < characters.len() {
        match characters[index] {
            '{' => {
                depth += 1;
                index += 1;
            }
            '}' => {
                if ignoring.map(|start| depth <= start).unwrap_or(false) {
                    ignoring = None;
                }
                depth = depth.saturating_sub(1);
                index += 1;
            }
            '\\' => {
                index += 1;
                if index >= characters.len() {
                    break;
                }

                match characters[index] {
                    '\\' | '{' | '}' => {
                        if ignoring.is_none() {
                            line.push(characters[index]);
                        }
                        index += 1;
                    }
                    '\'' => {
                        let hex: String = characters.iter().skip(index + 1).take(2).collect();
                        if let Ok(value) = u8::from_str_radix(&hex, 16) {
                            if ignoring.is_none() {
                                line.push(cp1252_char(value));
                            }
                        }
                        index += 3;
                    }
                    '*' => {
                        ignoring = Some(depth);
                        index += 1;
                    }
                    character if character.is_ascii_alphabetic() => {
                        let mut word = String::new();
                        while index < characters.len() && characters[index].is_ascii_alphabetic() {
                            word.push(characters[index]);
                            index += 1;
                        }

                        let mut negative = false;
                        let mut parameter: Option<i64> = None;
                        if index < characters.len()
                            && (characters[index] == '-' || characters[index].is_ascii_digit())
                        {
                            if characters[index] == '-' {
                                negative = true;
                                index += 1;
                            }
                            let mut digits = String::new();
                            while index < characters.len() && characters[index].is_ascii_digit() {
                                digits.push(characters[index]);
                                index += 1;
                            }
                            let value: i64 = digits.parse().unwrap_or(0);
                            parameter = Some(if negative { -value } else { value });
                        }

                        // A trailing space belongs to the control word.
                        if index < characters.len() && characters[index] == ' ' {
                            index += 1;
                        }

                        if word == "uc" {
                            unicode_skip = parameter.unwrap_or(1).max(0) as usize;
                            continue;
                        }
                        if word == "u" {
                            if ignoring.is_none() {
                                if let Some(value) = parameter {
                                    let code = if value < 0 { value + 65536 } else { value };
                                    if let Some(character) = char::from_u32(code as u32) {
                                        line.push(character);
                                    }
                                }
                            }
                            // The escape is followed by `unicode_skip` fallback
                            // characters, which the document does not need.
                            for _ in 0..unicode_skip {
                                skip_rtf_character(&characters, &mut index);
                            }
                            continue;
                        }

                        if ignoring.is_some() {
                            continue;
                        }

                        if RTF_IGNORED_DESTINATIONS.contains(&word.as_str()) {
                            ignoring = Some(depth);
                            continue;
                        }

                        match word.as_str() {
                            "par" | "line" | "sect" => lines.push(std::mem::take(&mut line)),
                            "tab" => line.push('\t'),
                            "cell" => line.push('\t'),
                            "row" => lines.push(std::mem::take(&mut line)),
                            "bullet" => line.push('•'),
                            "emdash" => line.push('—'),
                            "endash" => line.push('–'),
                            "lquote" => line.push('‘'),
                            "rquote" => line.push('’'),
                            "ldblquote" => line.push('“'),
                            "rdblquote" => line.push('”'),
                            _ => {}
                        }
                    }
                    _ => index += 1,
                }
            }
            '\r' | '\n' => index += 1,
            character => {
                if ignoring.is_none() {
                    line.push(character);
                }
                index += 1;
            }
        }
    }

    if !line.trim().is_empty() {
        lines.push(line);
    }

    lines
        .iter()
        .map(|line| line.trim_end().to_string())
        .collect()
}

/// Consume one character of RTF source, whether it is plain text or an escape,
/// which is how `\uN`'s fallback characters are counted.
fn skip_rtf_character(characters: &[char], index: &mut usize) {
    if *index >= characters.len() {
        return;
    }

    if characters[*index] != '\\' {
        *index += 1;
        return;
    }

    *index += 1;
    if *index < characters.len() && characters[*index] == '\'' {
        *index = (*index + 3).min(characters.len());
        return;
    }

    while *index < characters.len() && !characters[*index].is_whitespace() {
        *index += 1;
    }
}

// ---------------------------------------------------------------------- ANSI

struct AnsiPalette {
    foreground: [[u8; 3]; 16],
    background: [[u8; 3]; 16],
}

/// The console palette, pulled toward the page's own contrast: an NFO's colors
/// assume a black console, so on a light page its dark colors would disappear.
impl AnsiPalette {
    const BASE: [[u8; 3]; 16] = [
        [12, 12, 12],
        [197, 15, 31],
        [19, 161, 14],
        [193, 156, 0],
        [0, 55, 218],
        [136, 23, 152],
        [58, 150, 221],
        [204, 204, 204],
        [118, 118, 118],
        [231, 72, 86],
        [22, 198, 12],
        [249, 241, 165],
        [59, 120, 255],
        [180, 0, 158],
        [97, 214, 214],
        [242, 242, 242],
    ];

    fn new(theme: &LoadedTheme) -> Self {
        let page = rgb(theme.background());
        let foreground = rgb(theme.foreground());

        let mut palette = Self {
            foreground: [[0, 0, 0]; 16],
            background: [[0, 0, 0]; 16],
        };

        for index in 0..16 {
            palette.foreground[index] = readable(Self::BASE[index], page);
            palette.background[index] = blend(page, Self::BASE[index], 0.25);
        }

        // Console black is the one color that has to become the page's own text
        // color to stay visible, whichever page it is on.
        palette.foreground[0] = foreground;

        palette
    }
}

fn push_ansi_spans(spans: &mut Vec<Span>, line: &str, base: TextStyle, palette: &AnsiPalette) {
    let mut style = base;
    let mut pending = String::new();
    let mut characters = line.chars().peekable();

    while let Some(character) = characters.next() {
        if character != '\u{1B}' {
            pending.push(character);
            continue;
        }

        if characters.peek() != Some(&'[') {
            continue;
        }
        characters.next();

        let mut parameters = String::new();
        let mut terminator = None;
        for next in characters.by_ref() {
            if next.is_ascii_digit() || next == ';' {
                parameters.push(next);
            } else {
                terminator = Some(next);
                break;
            }
        }

        let flushes = terminator == Some('m');
        if flushes {
            if !pending.is_empty() {
                spans.push(Span {
                    text: std::mem::take(&mut pending),
                    style,
                });
            }
            apply_sgr(&mut style, &parameters, base, palette);
        } else if terminator.is_none() {
            // A truncated sequence: the rest of the line is text.
            break;
        }
    }

    if !pending.is_empty() {
        spans.push(Span {
            text: pending,
            style,
        });
    }
}

fn apply_sgr(style: &mut TextStyle, parameters: &str, base: TextStyle, palette: &AnsiPalette) {
    let values: Vec<u32> = if parameters.is_empty() {
        vec![0]
    } else {
        parameters
            .split(';')
            .map(|value| value.parse().unwrap_or(0))
            .collect()
    };

    for value in values {
        match value {
            0 => {
                style.foreground = base.foreground;
                style.background = None;
                style.bold = false;
                style.underline = false;
            }
            1 => style.bold = true,
            22 => style.bold = false,
            4 => style.underline = true,
            24 => style.underline = false,
            7 => {
                let foreground = style.foreground;
                style.foreground = style.background.unwrap_or(base.foreground);
                style.background = Some(foreground);
            }
            30..=37 => style.foreground = palette.foreground[(value - 30) as usize],
            39 => style.foreground = base.foreground,
            40..=47 => style.background = Some(palette.background[(value - 40) as usize]),
            49 => style.background = None,
            90..=97 => style.foreground = palette.foreground[(value - 90 + 8) as usize],
            100..=107 => style.background = Some(palette.background[(value - 100 + 8) as usize]),
            _ => {}
        }
    }
}

// -------------------------------------------------------------------- layout

struct LaidRun {
    text: String,
    style: TextStyle,
    x: i32,
    width: i32,
}

struct LaidLine {
    runs: Vec<LaidRun>,
    top: i32,
    height: i32,
    block: LineBlock,
}

struct LaidOut {
    lines: Vec<LaidLine>,
    width: u32,
    height: u32,
    /// Document line the frame starts at, after the position was pulled back so
    /// the frame is full.
    first_line: usize,
    /// Document lines the frame shows.
    visible_lines: usize,
}

/// The lines that fit above the box's bottom edge, before anything is said about
/// what did not.
struct BodyLayout {
    lines: Vec<LaidLine>,
    y: i32,
    content_right: i32,
    emitted: usize,
    /// Whether the frame stopped because the box was full rather than because the
    /// document ran out. A frame that is full has nothing to pull back into; one
    /// that is not is the last screenful of the document, waiting to be filled by
    /// the lines above it.
    full: bool,
}

/// Lay the document out inside `box_width` x `box_height`, starting at
/// `first_line`, and report the box the content needs.
///
/// `first_line` is in the lines the frame is drawn in — the document's own, except
/// for a rendered Markdown document, whose lines are the ones its renderer makes.
/// It is a request rather than a command: a frame that would end before the bottom
/// of the box is pulled back until the lines left in the document fill it, so the
/// last screenful of a document is a full one and its last line is on screen.
///
/// A line wider than the box wraps onto the next one — nothing in a preview scrolls
/// sideways, so whatever went past the edge would be out of reach for good — and
/// only the lines the frame shows are ever styled, which is what keeps a preview of
/// a large file cheap. `scrollbar_space` is room kept clear at the right edge for a
/// scrollbar the caller is about to draw, and `full_mode` says whether there is a
/// scrollbar to reach the rest of the document at all: without one, what the box
/// cannot hold is said in a line.
#[allow(clippy::too_many_arguments)] // Each one is a distinct number about the request.
fn layout(
    document: &CachedDocument,
    theme: &LoadedTheme,
    metrics: &TextMetrics,
    first_line: usize,
    box_width: u32,
    box_height: u32,
    scrollbar_space: i32,
    full_mode: bool,
) -> LaidOut {
    let doc = &document.doc;
    let padding = metrics.padding;
    let scrollable_lines = doc.scrollable_lines();
    let min_line_height = metrics
        .line_height
        .iter()
        .copied()
        .min()
        .unwrap_or(1)
        .max(1);

    // How many lines a frame of this height holds, which is what tells a preview
    // that cannot scroll how much of the document it is leaving out.
    let capacity = (((box_height.max(1) as i32 - padding * 2).max(0) / min_line_height) as usize)
        .max(1)
        .min(scrollable_lines.max(1));

    // A line is kept for a note whenever the frame cannot show everything and no
    // scrollbar will say so: a file that was cut at the read cap, whose remainder
    // cannot be reached by scrolling either, and a document in a preview that does
    // not scroll at all.
    let note_needed = doc.read_truncated || (!full_mode && scrollable_lines > capacity);
    let note_height = if note_needed {
        metrics.line_height[BODY_LEVEL as usize]
    } else {
        0
    };

    // The frame starts where it was asked to, kept inside the lines there are: a
    // request past the end of the document is the last line of it, never a box
    // with nothing in it.
    let mut first_line = first_line.min(scrollable_lines.saturating_sub(1));
    let mut body = lay_out_body(
        document,
        theme,
        metrics,
        first_line,
        (box_width, box_height),
        scrollbar_space,
        note_height,
    );

    // A frame is pulled back until it can be filled from where it starts, which is
    // what keeps the last screenful of a document full instead of ending in empty
    // page. What fills it are the lines above it, so those are walked backwards
    // from its top: the pull-back is measured in the document's own lines rather
    // than in the lines that fit, because one line can take two of the box's and a
    // pull-back by the number of lines that fit would leave the end of the
    // document out of reach.
    if first_line > 0 && !body.full {
        let bottom = body_bottom(box_height, padding, note_height);
        let pulled_back = pulled_back_first_line(
            document,
            theme,
            metrics,
            first_line,
            box_width,
            scrollbar_space,
            bottom - body.y,
        );

        if pulled_back != first_line {
            first_line = pulled_back;
            body = lay_out_body(
                document,
                theme,
                metrics,
                first_line,
                (box_width, box_height),
                scrollbar_space,
                note_height,
            );
        }
    }

    if note_needed {
        let bottom = box_height.max(1) as i32 - padding;
        if body.y + note_height <= bottom {
            let remaining = doc.total_lines().saturating_sub(first_line + body.emitted);
            let text = if remaining > 0 {
                if doc.read_truncated {
                    format!("… {remaining} more lines (file truncated)")
                } else {
                    format!("… {remaining} more lines")
                }
            } else {
                "… the rest of the file is not shown".to_string()
            };

            let mut style = plain_style(BODY_LEVEL);
            style.foreground = blend(rgb(theme.foreground()), rgb(theme.background()), 0.45);
            let advance = metrics.advance[BODY_LEVEL as usize].max(1);
            let width = text.chars().count() as i32 * advance;

            body.lines.push(LaidLine {
                runs: vec![LaidRun {
                    text,
                    style,
                    x: padding,
                    width,
                }],
                top: body.y,
                height: note_height,
                block: LineBlock::Plain,
            });
            body.y += note_height;
            body.content_right = body.content_right.max(padding + width);
        }
    }

    LaidOut {
        lines: body.lines,
        width: (body.content_right + padding).clamp(1, box_width.max(1) as i32) as u32,
        height: (body.y + padding).clamp(1, box_height.max(1) as i32) as u32,
        first_line,
        visible_lines: body.emitted,
    }
}

/// The bottom edge of the text area inside a box, with `reserve` kept clear at the
/// end of it for a line the caller is about to add.
fn body_bottom(box_height: u32, padding: i32, reserve: i32) -> i32 {
    (box_height.max(1) as i32 - padding).max(padding + 1) - reserve
}

/// Where a frame whose lines do not fill the box has to start for the box to be
/// full, or the line it was asked for when it already is.
///
/// The frame ends at the last line of the document, so what can still fill it are
/// the lines above that top, walked backwards while the box has room for another.
/// What comes back is the earliest line whose tail still fits — which is where a
/// preview stops when it is scrolled to the end, so the last line is on screen and
/// nothing is left below it.
///
/// The walk is bounded by the room itself: every line takes at least the smallest
/// line height of the box, so no more than that many can be added.
fn pulled_back_first_line(
    document: &CachedDocument,
    theme: &LoadedTheme,
    metrics: &TextMetrics,
    first_line: usize,
    box_width: u32,
    scrollbar_space: i32,
    room: i32,
) -> usize {
    if first_line == 0 || room <= 0 {
        return first_line;
    }

    let padding = metrics.padding;
    let min_line_height = metrics
        .line_height
        .iter()
        .copied()
        .min()
        .unwrap_or(1)
        .max(1);
    let take = ((room + min_line_height - 1) / min_line_height) as usize;
    let from = first_line.saturating_sub(take);
    let Some(window) = styled_window(&document.key, document, theme, from, first_line - from)
    else {
        return first_line;
    };

    let left = padding;
    let right = (box_width.max(1) as i32 - padding - scrollbar_space).max(left + 1);
    let mut start = first_line;
    let mut room = room;

    for index in (from..first_line).rev() {
        let Some(line) = window.line(index) else {
            break;
        };

        // The height the line takes in the frame, which is the height the
        // placement gives it.
        let height = line_height(line, metrics);
        // One line more than the room holds, which is enough of a line that does
        // not fit for the answer to be over the room either way.
        let max_lines = (room / height) as usize + 1;
        let needed = visual_lines_of(line, metrics, left, right, max_lines).len() as i32 * height;
        if needed > room {
            break;
        }

        room -= needed;
        start = index;
    }

    start
}

/// The lines of the document that fit inside the box, starting at `first_line`.
///
/// Only the window that will be drawn is styled: the request to [`styled_window`]
/// is sized from the box, so a file that is a hundred pages long has a screenful
/// highlighted and nothing else.
fn lay_out_body(
    document: &CachedDocument,
    theme: &LoadedTheme,
    metrics: &TextMetrics,
    first_line: usize,
    box_size: (u32, u32),
    scrollbar_space: i32,
    reserve: i32,
) -> BodyLayout {
    let doc = &document.doc;
    let (box_width, box_height) = box_size;
    let padding = metrics.padding;
    let box_width = box_width.max(1) as i32;
    let left = padding;
    let right = (box_width - padding - scrollbar_space).max(left + 1);
    let bottom = body_bottom(box_height, padding, reserve);
    let body_advance = metrics.advance[BODY_LEVEL as usize].max(1);
    let min_line_height = metrics
        .line_height
        .iter()
        .copied()
        .min()
        .unwrap_or(1)
        .max(1);

    // One more line than can fit, so the layout can tell "the frame is full" from
    // "the document ends here".
    let wanted = (((bottom - padding).max(0) / min_line_height) as usize).saturating_add(2);
    let window = styled_window(&document.key, document, theme, first_line, wanted);

    let mut lines: Vec<LaidLine> = Vec::new();
    let mut y = padding;
    let mut content_right = left + body_advance * MIN_CONTENT_CHARS;
    let mut emitted = 0usize;
    let mut full = false;

    if let Some(window) = window {
        let mut index = first_line;
        while index < doc.scrollable_lines() {
            let Some(line) = window.line(index) else {
                break;
            };

            let height = line_height(line, metrics);
            if y + height > bottom {
                full = true;
                break;
            }

            // The visual lines the box still has room for are all of this line a
            // frame can ever show, so that is as far as it is wrapped.
            let max_lines = ((bottom - y) / height) as usize;
            let visual_lines = visual_lines_of(line, metrics, left, right, max_lines);

            let mut placed = false;
            for runs in visual_lines {
                if y + height > bottom {
                    break;
                }

                // An empty line has no runs to measure, so it adds nothing to the
                // width the frame needs.
                let right_edge = runs
                    .iter()
                    .map(|run| run.x + run.width)
                    .max()
                    .unwrap_or(left);
                content_right = content_right.max(right_edge);
                lines.push(LaidLine {
                    runs,
                    top: y,
                    height,
                    block: line.block,
                });
                y += height;
                placed = true;
            }

            if !placed {
                full = true;
                break;
            }
            emitted += 1;
            index += 1;

            // A document that continues past the window has to be styled further
            // before the next line can be laid out; rebuilding the window is how
            // that happens, and it is bounded by the same margin.
            if index >= window.first + window.lines.len() {
                break;
            }
        }
    }

    BodyLayout {
        lines,
        y,
        content_right,
        emitted,
        full,
    }
}

/// One document line as the frame draws it inside its text column: the visual
/// lines it takes, left to right, wrapped at the page edge.
///
/// A line wider than the column continues on the next one. Nothing in a text
/// preview scrolls sideways, so a line that ran past the edge would put whatever it
/// held beyond it out of reach for good, whatever the file is — code and NFO art as
/// much as prose.
///
/// Only `max_lines` of them are wrapped, which is as many as the frame could ever
/// draw: what lies past that is a tail no scroll position can reach, because a
/// position is a document line and every frame starts at the first of a line's own
/// visual lines. A minified file is one line of megabytes, and that is what keeps a
/// hover bounded by the box rather than by the file.
///
/// Both the placement and the pull-back measure a line with this, so the height a
/// line is given and the height it is walked back by cannot drift apart.
fn visual_lines_of(
    line: &DocLine,
    metrics: &TextMetrics,
    left: i32,
    right: i32,
    max_lines: usize,
) -> Vec<Vec<LaidRun>> {
    let indent = left + metrics.indent(line.indent);
    let text_right = match line.block {
        LineBlock::Quote => right - metrics.quote_bar - 4,
        _ => right,
    };

    wrap_line(line, indent, text_right, max_lines, metrics)
}

/// A word or the spaces between words, carrying the style it was colored with.
struct Token {
    text: String,
    style: TextStyle,
    space: bool,
}

/// Walk a line's words and the spaces between them, so a wrapped line breaks
/// between words instead of in the middle of one and keeps its colors across the
/// break.
///
/// The walk stops the moment `visit` answers false. That is what keeps a line of a
/// minified file — megabytes on one line — from being split any further than a
/// frame can draw, without deciding up front how much of it that is: only as much
/// of the line as is drawn is ever copied.
fn for_each_token(line: &DocLine, mut visit: impl FnMut(Token) -> bool) {
    for span in &line.spans {
        let mut text = String::new();
        let mut space = false;

        for character in span.text.chars() {
            let is_space = character == ' ';
            if !text.is_empty() && is_space != space {
                let token = Token {
                    text: std::mem::take(&mut text),
                    style: span.style,
                    space,
                };
                if !visit(token) {
                    return;
                }
            }
            space = is_space;
            text.push(character);
        }

        if !text.is_empty() {
            let token = Token {
                text,
                style: span.style,
                space,
            };
            if !visit(token) {
                return;
            }
        }
    }
}

/// One document line's visual lines, wrapped at the page edge and at most
/// `max_lines` of them. A word wider than the page — a URL, a hash — is cut rather
/// than allowed to run off it, and a space that lands at a wrap point is dropped
/// instead of starting a line.
fn wrap_line(
    line: &DocLine,
    left: i32,
    limit: i32,
    max_lines: usize,
    metrics: &TextMetrics,
) -> Vec<Vec<LaidRun>> {
    let mut lines: Vec<Vec<LaidRun>> = Vec::new();
    let mut runs: Vec<LaidRun> = Vec::new();
    let mut x = left;

    for_each_token(line, |token| {
        let advance = metrics.advance[(token.style.level as usize).min(SIZE_LEVELS - 1)].max(1);
        let width = token.text.chars().count() as i32 * advance;

        if token.space {
            if runs.is_empty() || x + width > limit {
                return true;
            }
            runs.push(LaidRun {
                text: token.text,
                style: token.style,
                x,
                width,
            });
            x += width;
            return true;
        }

        if x > left && x + width > limit {
            // This word belongs to another visual line, and there is no room left
            // in the frame for one.
            if lines.len() + 1 == max_lines {
                return false;
            }
            lines.push(std::mem::take(&mut runs));
            x = left;
        }

        // A word that fits where the line stands is taken whole, so only one too
        // wide for the room left of it is ever split.
        if width <= limit - x {
            runs.push(LaidRun {
                text: token.text,
                style: token.style,
                x,
                width,
            });
            x += width;
            return true;
        }

        // A word too wide for the room left of it is split across the lines it
        // takes, and nothing else reaches here.
        let characters: Vec<char> = token.text.chars().collect();
        let mut offset = 0usize;
        while offset < characters.len() {
            let available = (((limit - x).max(advance)) / advance) as usize;
            let take = (characters.len() - offset).min(available.max(1));
            let text: String = characters[offset..offset + take].iter().collect();
            let chunk = take as i32 * advance;

            runs.push(LaidRun {
                text,
                style: token.style,
                x,
                width: chunk,
            });
            x += chunk;
            offset += take;

            if offset < characters.len() {
                if lines.len() + 1 == max_lines {
                    return false;
                }
                lines.push(std::mem::take(&mut runs));
                x = left;
            }
        }

        true
    });

    if !runs.is_empty() {
        lines.push(runs);
    }
    if lines.is_empty() {
        // An empty line still takes a line's worth of space.
        lines.push(Vec::new());
    }

    lines
}

fn line_height(line: &DocLine, metrics: &TextMetrics) -> i32 {
    line.spans
        .iter()
        .map(|span| metrics.line_height[(span.style.level as usize).min(SIZE_LEVELS - 1)])
        .max()
        .unwrap_or(metrics.line_height[BODY_LEVEL as usize])
}

/// Where the scrollbar sits in a frame showing `visible_lines` of `total_lines`
/// from `first_line`.
///
/// The thumb's length is the share of the document the frame shows, and its
/// position is the share that has been scrolled past it — the two things a
/// reader uses to tell how much more there is.
fn scrollbar_geometry(
    width: i32,
    height: i32,
    metrics: &TextMetrics,
    first_line: usize,
    visible_lines: usize,
    total_lines: usize,
) -> Option<ScrollBar> {
    if total_lines <= visible_lines || visible_lines == 0 || width <= 0 || height <= 0 {
        return None;
    }

    let padding = metrics.padding;
    let bar_width = metrics.scrollbar_width();
    let right = width - padding;
    let left = (right - bar_width).max(0);
    let top = padding;
    let bottom = (height - padding).max(top + 1);
    let track_height = bottom - top;

    let thumb_height = ((track_height as f32) * (visible_lines as f32 / total_lines as f32))
        .round()
        .clamp(
            scaled(SCROLLBAR_MIN_THUMB_PIXELS, metrics.scale).min(track_height) as f32,
            track_height as f32,
        ) as i32;

    let max_first = (total_lines - visible_lines) as f32;
    let travelled = if max_first > 0.0 {
        (first_line as f32 / max_first).clamp(0.0, 1.0)
    } else {
        0.0
    };
    let thumb_top = top + ((track_height - thumb_height) as f32 * travelled).round() as i32;

    Some(ScrollBar {
        track: (left, top, right, bottom),
        thumb: (left, thumb_top, right, thumb_top + thumb_height),
    })
}

/// The document line a drag to `y` inside the track asks for.
///
/// The inverse of the thumb's placement, shared by both places that need it: the
/// thumb is dragged by its middle, and the ends of the track are the ends of the
/// document however long it is.
pub fn scroll_line_at_track_y(
    track: (i32, i32, i32, i32),
    thumb: (i32, i32, i32, i32),
    y: i32,
    visible_lines: usize,
    total_lines: usize,
) -> usize {
    let track_height = (track.3 - track.1).max(1);
    let thumb_height = (thumb.3 - thumb.1).max(1);
    let travel = (track_height - thumb_height).max(1);
    let ratio = ((y - track.1 - thumb_height / 2) as f32 / travel as f32).clamp(0.0, 1.0);

    let max_first = total_lines.saturating_sub(visible_lines);
    (ratio * max_first as f32).round() as usize
}

// ------------------------------------------------------------------ painting

/// The part of a line's selection that falls inside one of its runs, as character
/// offsets within that run, or `None` when the selection does not touch it.
///
/// A line is painted run by run — one per token, which is where the colors change —
/// so a selection, which is measured in the line's characters, has to be mapped
/// back onto the run it lands in. `run_start` is how many characters of the line
/// come before this run.
fn selected_part(
    line_index: usize,
    run_start: usize,
    run_characters: usize,
    selection: Selection,
) -> Option<(usize, usize)> {
    let ((start_line, start_char), (end_line, end_char)) = selection.ordered();
    if line_index < start_line || line_index > end_line || run_characters == 0 {
        return None;
    }

    let line_from = if line_index == start_line {
        start_char
    } else {
        0
    };
    let line_to = if line_index == end_line {
        end_char
    } else {
        usize::MAX
    };

    let from = line_from.saturating_sub(run_start).min(run_characters);
    let to = line_to.saturating_sub(run_start).min(run_characters);

    (from < to).then_some((from, to))
}

/// Paint the background, the block decorations, the text and a selection, and the
/// scrollbar when the document is longer than the frame.
///
/// Every run is drawn with an opaque background rectangle, so the page color
/// fills the box and the pieces GDI leaves alone (the space a line that stops
/// before the edge did not use, the gaps between blocks) are already the right
/// color.
unsafe fn paint(
    surface: &DibSurface,
    laid_out: &LaidOut,
    theme: &LoadedTheme,
    metrics: &TextMetrics,
    scrollbar: Option<&ScrollBar>,
    selection: Option<Selection>,
) {
    let width = surface.width as i32;
    let page = rgb(theme.background());
    let band = blend(page, rgb(theme.foreground()), 0.08);
    // A selection is drawn with the theme's own highlight where it has one, and
    // with a shade of the page where it does not.
    let highlight = theme
        .theme()
        .settings
        .selection
        .map(rgb)
        .unwrap_or_else(|| blend(page, rgb(theme.foreground()), 0.25));
    let highlight_foreground = theme.theme().settings.selection_foreground.map(rgb);
    let selection = selection.filter(|selection| !selection.is_empty());

    fill_rect(
        surface,
        RECT {
            left: 0,
            top: 0,
            right: width,
            bottom: surface.height as i32,
        },
        page,
    );

    for line in &laid_out.lines {
        match line.block {
            LineBlock::Code => fill_rect(
                surface,
                RECT {
                    left: 0,
                    top: line.top,
                    right: width,
                    bottom: line.top + line.height,
                },
                band,
            ),
            LineBlock::Quote => {
                let bar = theme.style_for_scopes(&["markup.quote"]).foreground;
                fill_rect(
                    surface,
                    RECT {
                        left: 0,
                        top: line.top,
                        right: metrics.quote_bar,
                        bottom: line.top + line.height,
                    },
                    rgb(bar),
                );
            }
            LineBlock::Plain => {}
        }
    }

    let mut painter = RunPainter::new(surface, metrics.scale);

    for (line_index, line) in laid_out.lines.iter().enumerate() {
        let mut run_start = 0usize;

        for run in &line.runs {
            let run_characters = run.text.chars().count();
            let selected_part = selection.and_then(|selection| {
                selected_part(line_index, run_start, run_characters, selection)
            });
            run_start += run_characters;

            if run.text.is_empty() {
                continue;
            }

            let run_background = match (run.style.background, line.block) {
                (Some(color), _) => color,
                (None, LineBlock::Code) => band,
                (None, _) => page,
            };

            let rect = RECT {
                left: run.x,
                top: line.top,
                right: run.x + run.width,
                bottom: line.top + line.height,
            };

            match selected_part {
                Some((from, to)) => {
                    // The selected part is painted over its own highlight, so the
                    // run is drawn in up to three pieces: the text before it as
                    // usual, the selection without touching its background, and the
                    // text after it as usual.
                    let characters: Vec<char> = run.text.chars().collect();
                    let advance = run.width as f32 / characters.len() as f32;
                    let piece = |start: usize, end: usize| {
                        characters[start..end].iter().collect::<String>()
                    };

                    painter.draw(
                        &piece(0, from),
                        rect.left,
                        rect,
                        &run.style,
                        run.style.foreground,
                        run_background,
                    );

                    let highlight_rect = RECT {
                        left: run.x + (from as f32 * advance).round() as i32,
                        right: run.x + (to as f32 * advance).round() as i32,
                        ..rect
                    };
                    painter.draw(
                        &piece(from, to),
                        highlight_rect.left,
                        highlight_rect,
                        &run.style,
                        highlight_foreground.unwrap_or(run.style.foreground),
                        highlight,
                    );

                    let tail_x = run.x + (to as f32 * advance).round() as i32;
                    painter.draw(
                        &piece(to, characters.len()),
                        tail_x,
                        RECT {
                            left: tail_x,
                            ..rect
                        },
                        &run.style,
                        run.style.foreground,
                        run_background,
                    );
                }
                None => painter.draw(
                    &run.text,
                    rect.left,
                    rect,
                    &run.style,
                    run.style.foreground,
                    run_background,
                ),
            }
        }
    }

    // The fonts go with the painter, and it puts the surface's own object back
    // before they are deleted.
    drop(painter);

    // The scrollbar goes last so it sits above everything, including a line that
    // ran to the right edge.
    if let Some(scrollbar) = scrollbar {
        let (left, top, right, bottom) = scrollbar.track;
        fill_rect(
            surface,
            RECT {
                left,
                top,
                right,
                bottom,
            },
            blend(page, rgb(theme.foreground()), 0.10),
        );

        let (left, top, right, bottom) = scrollbar.thumb;
        fill_rect(
            surface,
            RECT {
                left,
                top,
                right,
                bottom,
            },
            blend(page, rgb(theme.foreground()), 0.45),
        );
    }
}

// ---------------------------------------------------------------------- tests

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;

    /// Where the fixtures and the pictures of them are written. The scratchpad
    /// the session hands out, so a render can be looked at rather than only
    /// asserted.
    fn scratch(label: &str) -> PathBuf {
        let root = std::env::var_os("COMMANDCODE_SCRATCHPAD")
            .map(PathBuf::from)
            .unwrap_or_else(std::env::temp_dir)
            .join("text-preview")
            .join(label);
        fs::create_dir_all(&root).expect("a fixture directory");
        root
    }

    fn write_png(dir: &Path, name: &str, frame: &TextFrame) {
        let mut rgba = frame.pixels.clone();
        for pixel in rgba.as_chunks_mut::<4>().0 {
            pixel.swap(0, 2);
        }
        image::save_buffer(
            dir.join(format!("{name}.png")),
            &rgba,
            frame.width,
            frame.height,
            image::ExtendedColorType::Rgba8,
        )
        .expect("a written picture");
    }

    /// The painting a text preview does is shared with archive listings, so the
    /// path is worth a smoke test of its own: a page is measured, painted, and
    /// opaque, and what it paints is the frame the compositor gets.
    #[test]
    fn draws_pages_of_text() {
        let dir = scratch("pages");
        let source = dir.join("sample.rs");
        fs::write(
            &source,
            "fn main() {\n    // a comment\n    let answer = 42;\n    println!(\"{answer}\");\n}\n",
        )
        .expect("a source file");

        let markdown = dir.join("sample.md");
        fs::write(
            &markdown,
            "# Heading\n\nSome *emphasis*, a [link](https://example.com), and `code`.\n\n```rust\nlet x = 1;\n```\n",
        )
        .expect("a document");

        for (path, name, full_mode, theme) in [
            (&source, "source-light", false, TextTheme::Light),
            (&markdown, "markdown-dark", true, TextTheme::Dark),
        ] {
            let options = TextPreviewOptions {
                theme,
                markdown_mode: MarkdownMode::Rendered,
                font_scale_percent: 125,
                full_mode,
            };

            let (width, height) =
                measure(path, 1_920, 1_200, 96, options).expect("a measured page");
            let frame = render_scrolled(path, 0, width, height, 96, options, None).expect("a page");

            assert!(frame.width > 0 && frame.height > 0, "{name}");
            assert!(
                frame
                    .pixels
                    .as_chunks::<4>()
                    .0
                    .iter()
                    .all(|pixel| pixel[3] == 255),
                "{name} is a page, so every pixel of it is opaque"
            );

            write_png(&dir, name, &frame);
        }
    }
}

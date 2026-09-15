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

use crate::config::{MarkdownMode, TextTheme};
use crate::text_theme::{self, LoadedTheme};
use once_cell::sync::Lazy;
use pulldown_cmark::{CodeBlockKind, Event, Options, Parser, Tag, TagEnd};
use std::fs::File;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};
use std::time::SystemTime;
use syntect::easy::HighlightLines;
use syntect::highlighting::{Color, FontStyle, Style};
use syntect::parsing::SyntaxSet;
use windows::core::PCWSTR;
use windows::Win32::Foundation::{COLORREF, RECT, SIZE};
use windows::Win32::Graphics::Gdi::{
    CreateCompatibleDC, CreateDIBSection, CreateFontW, DeleteDC, DeleteObject, ExtTextOutW,
    GetTextExtentPoint32W, GetTextMetricsW, SelectObject, SetBkColor, SetBkMode, SetTextColor,
    BITMAPINFO, BITMAPINFOHEADER, BI_RGB, CLEARTYPE_QUALITY, CLIP_DEFAULT_PRECIS, DEFAULT_CHARSET,
    DIB_RGB_COLORS, ETO_CLIPPED, ETO_OPAQUE, FF_MODERN, FIXED_PITCH, HBITMAP, HDC, HFONT, HGDIOBJ,
    OPAQUE, OUT_TT_PRECIS, TEXTMETRICW,
};

/// Bytes read from the file. A preview shows one screenful, and this bounds the
/// work a hover can trigger on a file that happens to be enormous.
const READ_LIMIT_BYTES: u64 = 2 * 1024 * 1024;

/// Lines carried through highlighting and layout. No display fits this many, so
/// whatever is cut here is reported as remaining lines instead of being dropped
/// silently.
const MAX_DOC_LINES: usize = 400;

/// Longest line kept, in characters. Minified sources and one-line data files
/// would otherwise turn a single hover into a megabyte-wide layout.
const MAX_LINE_CHARS: usize = 2000;

const TAB_WIDTH: usize = 4;

/// Parsed documents kept in memory, keyed by file and rendering options.
const DOC_CACHE_MAX_ENTRIES: usize = 24;

/// Size steps the document model uses: body text, four heading levels.
const SIZE_LEVELS: usize = 5;
const LEVEL_FONT_PIXELS: [i32; SIZE_LEVELS] = [13, 20, 17, 15, 14];
const LEVEL_EXTRA_LEADING: [i32; SIZE_LEVELS] = [0, 8, 6, 3, 0];
const BODY_LEVEL: u8 = 0;

const PADDING_PIXELS: f32 = 12.0;
const QUOTE_BAR_PIXELS: i32 = 4;

/// Narrow files still get a window wide enough to look like one.
const MIN_CONTENT_CHARS: i32 = 16;

/// Consolas ships with Windows, is fixed pitch, and carries the box-drawing
/// characters NFO art is made of; bold and italic keep the same advance, which
/// is what lets the layout measure a line by counting characters.
const FONT_FACE: &str = "Consolas";

const MIN_DPI: u32 = 48;
const MAX_DPI: u32 = 480;

/// The options a preview is built with. They are passed in rather than read from
/// the configuration inside this module so the cache key and the caller's intent
/// cannot drift apart.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextPreviewOptions {
    pub theme: TextTheme,
    pub markdown_mode: MarkdownMode,
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

    let result = TextMetrics::new(dc, dpi).map(|metrics| {
        let laid_out = layout(&document, max_width, max_height, &metrics, theme);
        (laid_out.width, laid_out.height)
    });

    unsafe {
        let _ = DeleteDC(dc);
    }

    result
}

/// Paint the preview into a `width` x `height` BGRA frame.
///
/// The frame is always exactly the requested size, with the theme's background
/// showing wherever the text does not reach, which is what keeps the painted
/// frame and the window the layout planned in step.
pub fn render(
    path: &Path,
    width: u32,
    height: u32,
    dpi: u32,
    options: TextPreviewOptions,
) -> Option<(Vec<u8>, u32, u32)> {
    if width == 0 || height == 0 {
        return None;
    }

    let document = document(path, options)?;
    let theme = text_theme::loaded(options.theme)?;

    let surface = DibSurface::create(width, height)?;
    let metrics = TextMetrics::new(surface.dc, dpi)?;
    let laid_out = layout(&document, width, height, &metrics, theme);

    unsafe {
        paint(&surface, &laid_out, theme, &metrics);
    }

    Some((surface.pixels(), width, height))
}

// ---------------------------------------------------------------- documents

/// A parsed document, ready to lay out: styled lines plus what was left out.
struct TextDoc {
    lines: Vec<DocLine>,
    /// Lines the file holds beyond the ones above.
    remaining_lines: usize,
    /// Whether the read stopped at the byte cap rather than at the end of the
    /// file, in which case the remaining count is a lower bound.
    read_truncated: bool,
    /// Whether a line wider than the page continues on the next one. Prose wraps —
    /// a paragraph is one long line and clipping it would hide most of what the
    /// file says — while code, markup and art keep their columns.
    wrap: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
enum LineBlock {
    #[default]
    Plain,
    /// A line of a fenced or indented code block.
    Code,
    /// A line inside a block quote.
    Quote,
}

#[derive(Default)]
struct DocLine {
    spans: Vec<Span>,
    block: LineBlock,
    /// Extra left indent, in characters, for nested lists and quotes.
    indent: u8,
}

struct Span {
    text: String,
    style: TextStyle,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct TextStyle {
    foreground: [u8; 3],
    background: Option<[u8; 3]>,
    bold: bool,
    italic: bool,
    underline: bool,
    strike: bool,
    level: u8,
}

#[derive(Clone, PartialEq, Eq)]
struct DocKey {
    path: PathBuf,
    modified: Option<SystemTime>,
    len: u64,
    options: TextPreviewOptions,
}

/// Parsed documents, the most recently built last, keyed by the file and the
/// options they were built from. Bounded: the whole cache is dropped when it
/// fills, the way the PDF page and video geometry caches are.
type DocumentCache = Vec<(DocKey, Arc<TextDoc>)>;

static DOCUMENTS: Lazy<Mutex<DocumentCache>> = Lazy::new(|| Mutex::new(Vec::new()));

/// Syntax definitions from syntect plus the ones it does not bundle, built once
/// on the first text hover. Individual syntaxes are still parsed lazily.
static SYNTAXES: Lazy<SyntaxSet> = Lazy::new(two_face::syntax::extra_no_newlines);

/// The parsed document for `path`, from the cache when the file, its mode and
/// its theme are unchanged.
fn document(path: &Path, options: TextPreviewOptions) -> Option<Arc<TextDoc>> {
    let metadata = std::fs::metadata(path).ok()?;
    let key = DocKey {
        path: path.to_path_buf(),
        modified: metadata.modified().ok(),
        len: metadata.len(),
        options,
    };

    if let Ok(cache) = DOCUMENTS.lock() {
        if let Some((_, document)) = cache.iter().find(|(cached, _)| *cached == key) {
            return Some(Arc::clone(document));
        }
    }

    let theme = text_theme::loaded(options.theme)?;
    let document = Arc::new(build_document(path, options, theme)?);

    if let Ok(mut cache) = DOCUMENTS.lock() {
        if cache.len() >= DOC_CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.push((key, Arc::clone(&document)));
    }

    Some(document)
}

fn build_document(
    path: &Path,
    options: TextPreviewOptions,
    theme: &LoadedTheme,
) -> Option<TextDoc> {
    let source = read_source(path)?;
    if source.text.trim().is_empty() {
        return None;
    }

    let extension = extension_of(path);

    if extension == "rtf" {
        return Some(plain_document(
            &strip_rtf(&source.text),
            theme,
            source.truncated,
        ));
    }

    if is_markdown_extension(extension) && options.markdown_mode == MarkdownMode::Rendered {
        return Some(markdown_document(
            &source.text,
            theme,
            source.lines.len(),
            source.truncated,
        ));
    }

    // NFO files carry ANSI color codes around plain text, so they are built from
    // their own escapes rather than from a syntax definition.
    if wants_ansi(extension) {
        return Some(ansi_document(&source.lines, theme, source.truncated));
    }

    Some(highlighted_document(path, &source, theme))
}

/// A document whose lines share one style, for text that carries no markup. It is
/// prose, so it wraps.
fn plain_document(lines: &[String], theme: &LoadedTheme, truncated: bool) -> TextDoc {
    let style = text_style(&theme.style_for_scopes(&["source"]), BODY_LEVEL);
    let mut document = TextDoc {
        lines: Vec::new(),
        remaining_lines: 0,
        read_truncated: truncated,
        wrap: true,
    };

    for line in lines.iter().take(MAX_DOC_LINES) {
        document.lines.push(DocLine {
            spans: vec![Span {
                text: line.clone(),
                style,
            }],
            block: LineBlock::Plain,
            indent: 0,
        });
    }

    document.remaining_lines = lines.len().saturating_sub(document.lines.len());
    document
}

fn highlighted_document(path: &Path, source: &SourceText, theme: &LoadedTheme) -> TextDoc {
    let syntax_set = &*SYNTAXES;
    let syntax = syntax_set
        .find_syntax_for_file(path)
        .ok()
        .flatten()
        .unwrap_or_else(|| syntax_set.find_syntax_plain_text());
    let mut highlighter = HighlightLines::new(syntax, theme.theme());

    let mut document = TextDoc {
        lines: Vec::new(),
        remaining_lines: 0,
        read_truncated: source.truncated,
        // A file no grammar claims is prose rather than code — a readme, a log, a
        // note — and prose that is cut off at the page edge is unreadable, so
        // those wrap. Anything with a syntax definition keeps its columns.
        wrap: syntax.name == "Plain Text",
    };

    for line in source.lines.iter().take(MAX_DOC_LINES) {
        let spans = highlight_line(&mut highlighter, line, syntax_set);
        document.lines.push(DocLine {
            spans,
            block: LineBlock::Plain,
            indent: 0,
        });
    }

    document.remaining_lines = source.lines.len().saturating_sub(document.lines.len());
    document
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

fn extension_of(path: &Path) -> &str {
    path.extension().and_then(|ext| ext.to_str()).unwrap_or("")
}

fn text_style(style: &Style, level: u8) -> TextStyle {
    TextStyle {
        foreground: rgb(style.foreground),
        background: None,
        bold: style.font_style.contains(FontStyle::BOLD),
        italic: style.font_style.contains(FontStyle::ITALIC),
        underline: style.font_style.contains(FontStyle::UNDERLINE),
        strike: false,
        level,
    }
}

fn rgb(color: Color) -> [u8; 3] {
    [color.r, color.g, color.b]
}

// ------------------------------------------------------------------ markdown

/// Walk the CommonMark events once, turning them into styled lines.
///
/// Block structure becomes lines and indent, inline structure becomes runs, and
/// fenced code is handed to the same highlighter the rest of the preview uses,
/// so a fenced block is colored like the language it names.
struct MarkdownBuilder<'a> {
    theme: &'a LoadedTheme,
    lines: Vec<DocLine>,
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
    fn new(theme: &'a LoadedTheme) -> Self {
        Self {
            theme,
            lines: Vec::new(),
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
        if self.lines.len() >= MAX_DOC_LINES {
            self.hit_line_cap = true;
            self.current = DocLine::default();
            return;
        }

        let line = std::mem::take(&mut self.current);
        self.lines.push(line);
    }

    /// One empty line between blocks, never two.
    fn blank_line(&mut self) {
        self.flush();
        if self.lines.is_empty()
            || self.lines.len() >= MAX_DOC_LINES
            || self
                .lines
                .last()
                .map(|line| line.spans.is_empty())
                .unwrap_or(true)
        {
            return;
        }

        self.lines.push(DocLine::default());
    }

    fn rule_line(&mut self) {
        self.blank_line();
        if self.lines.len() >= MAX_DOC_LINES {
            return;
        }
        self.lines.push(DocLine::default());
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
            if self.lines.len() >= MAX_DOC_LINES {
                self.hit_line_cap = true;
                break;
            }
            self.lines.push(DocLine {
                spans: highlight_line(&mut highlighter, line, syntax_set),
                block: LineBlock::Code,
                indent,
            });
        }

        self.current = DocLine::default();
    }

    fn finish(mut self, source_lines: usize) -> TextDoc {
        self.flush();
        while self
            .lines
            .last()
            .map(|line| line.spans.is_empty())
            .unwrap_or(false)
        {
            self.lines.pop();
        }

        let remaining_lines = if self.hit_line_cap {
            source_lines.saturating_sub(self.lines.len())
        } else {
            0
        };

        TextDoc {
            lines: self.lines,
            remaining_lines,
            read_truncated: false,
            // A rendered document is prose: its paragraphs wrap.
            wrap: true,
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

fn markdown_document(
    text: &str,
    theme: &LoadedTheme,
    source_lines: usize,
    truncated: bool,
) -> TextDoc {
    let options =
        Options::ENABLE_TABLES | Options::ENABLE_STRIKETHROUGH | Options::ENABLE_TASKLISTS;
    let mut builder = MarkdownBuilder::new(theme);

    for event in Parser::new_ext(text, options) {
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

    let mut document = builder.finish(source_lines);
    document.read_truncated = truncated;
    document
}

// -------------------------------------------------------------------- source

/// A decoded file, split into lines with tabs already expanded.
struct SourceText {
    text: String,
    lines: Vec<String>,
    truncated: bool,
}

fn read_source(path: &Path) -> Option<SourceText> {
    let file = File::open(path).ok()?;
    let mut bytes = Vec::new();
    let mut limited = file.take(READ_LIMIT_BYTES + 1);
    limited.read_to_end(&mut bytes).ok()?;

    let truncated = bytes.len() as u64 > READ_LIMIT_BYTES;
    bytes.truncate(READ_LIMIT_BYTES as usize);

    let text = decode_text(&bytes, legacy_encoding(extension_of(path)), truncated)?;
    let text = expand_tabs(&text);

    Some(SourceText {
        lines: text.lines().map(|line| line.to_string()).collect(),
        text,
        truncated,
    })
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
        .chunks_exact(2)
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

/// An SGR-colored document: the escapes are consumed, and the text they wrapped
/// is emitted with the color they selected.
fn ansi_document(lines: &[String], theme: &LoadedTheme, truncated: bool) -> TextDoc {
    let palette = AnsiPalette::new(theme);
    let base = text_style(&theme.style_for_scopes(&["source"]), BODY_LEVEL);
    let mut document = TextDoc {
        lines: Vec::new(),
        remaining_lines: 0,
        read_truncated: truncated,
        // NFO art is drawn in columns; wrapping it would destroy the picture.
        wrap: false,
    };

    for line in lines.iter().take(MAX_DOC_LINES) {
        let mut spans = Vec::new();
        push_ansi_spans(&mut spans, line, base, &palette);
        document.lines.push(DocLine {
            spans,
            block: LineBlock::Plain,
            indent: 0,
        });
    }

    document.remaining_lines = lines.len().saturating_sub(document.lines.len());
    document
}

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

fn luminance(color: [u8; 3]) -> f32 {
    (0.2126 * color[0] as f32 + 0.7152 * color[1] as f32 + 0.0722 * color[2] as f32) / 255.0
}

fn blend(base: [u8; 3], other: [u8; 3], amount: f32) -> [u8; 3] {
    let mix = |a: u8, b: u8| (a as f32 + (b as f32 - a as f32) * amount).round() as u8;
    [
        mix(base[0], other[0]),
        mix(base[1], other[1]),
        mix(base[2], other[2]),
    ]
}

/// Pull a color away from the page until it can be read on it.
fn readable(color: [u8; 3], background: [u8; 3]) -> [u8; 3] {
    let on_light_page = luminance(background) > 0.5;
    let limit = 0.45;
    let target = if on_light_page {
        [0, 0, 0]
    } else {
        [255, 255, 255]
    };

    let mut result = color;
    for step in 0..=10 {
        let value = luminance(result);
        let too_close = if on_light_page {
            value > limit
        } else {
            value < limit
        };
        if !too_close {
            break;
        }
        result = blend(color, target, step as f32 / 10.0);
    }

    result
}

// -------------------------------------------------------------------- layout

/// Font metrics for one display scale, taken from GDI once per preview.
struct TextMetrics {
    scale: f32,
    padding: i32,
    quote_bar: i32,
    advance: [i32; SIZE_LEVELS],
    line_height: [i32; SIZE_LEVELS],
}

impl TextMetrics {
    fn new(dc: HDC, dpi: u32) -> Option<Self> {
        let dpi = dpi.clamp(MIN_DPI, MAX_DPI);
        let scale = dpi as f32 / 96.0;

        let mut advance = [0i32; SIZE_LEVELS];
        let mut line_height = [0i32; SIZE_LEVELS];

        for level in 0..SIZE_LEVELS {
            let pixels = scaled(LEVEL_FONT_PIXELS[level], scale);
            let font = create_font(pixels, &plain_style(level as u8));
            if font.0.is_null() {
                return None;
            }

            let mut keep = false;
            unsafe {
                let previous = SelectObject(dc, font);

                let mut metrics = TEXTMETRICW::default();
                let measured = GetTextMetricsW(dc, &mut metrics).as_bool();

                // The advance of one character is the whole measurement a fixed
                // pitch face needs.
                let sample = [b'0' as u16];
                let mut extent = SIZE::default();
                let sampled = GetTextExtentPoint32W(dc, &sample, &mut extent).as_bool();

                let _ = SelectObject(dc, previous);
                let _ = DeleteObject(font);

                if measured && sampled && extent.cx > 0 {
                    advance[level] = extent.cx;
                    line_height[level] = metrics.tmHeight
                        + metrics.tmExternalLeading
                        + scaled(LEVEL_EXTRA_LEADING[level], scale);
                    keep = true;
                }
            }

            if !keep {
                return None;
            }
        }

        Some(Self {
            scale,
            padding: scaled(PADDING_PIXELS as i32, scale),
            quote_bar: scaled(QUOTE_BAR_PIXELS, scale),
            advance,
            line_height,
        })
    }

    fn indent(&self, characters: u8) -> i32 {
        characters as i32 * self.advance[BODY_LEVEL as usize]
    }
}

fn plain_style(level: u8) -> TextStyle {
    TextStyle {
        foreground: [0, 0, 0],
        background: None,
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        level,
    }
}

fn scaled(pixels: i32, scale: f32) -> i32 {
    ((pixels as f32) * scale).round().max(1.0) as i32
}

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
}

/// The lines that fit above the box's bottom edge, before anything is said about
/// what did not.
struct BodyLayout {
    lines: Vec<LaidLine>,
    y: i32,
    content_right: i32,
    emitted: usize,
}

/// Lay the document out inside `box_width` x `box_height` and report the box the
/// content needs.
///
/// Lines are clipped at the right edge rather than wrapped — column-aligned code
/// reads better unwrapped — and the document is cut at the bottom edge with the
/// count of what was left out, so a long file says so instead of just stopping.
fn layout(
    document: &TextDoc,
    box_width: u32,
    box_height: u32,
    metrics: &TextMetrics,
    theme: &LoadedTheme,
) -> LaidOut {
    let padding = metrics.padding;
    let tail_height = metrics.line_height[BODY_LEVEL as usize];
    let total = document.lines.len() + document.remaining_lines;

    let mut body = lay_out_body(document, box_width, box_height, metrics, 0);

    if body.emitted < total {
        // The file continues past the page, so the last line is kept for the
        // count: a full page of text that simply stops reads as the whole file.
        body = lay_out_body(document, box_width, box_height, metrics, tail_height);

        let bottom = box_height.max(1) as i32 - padding;
        if body.y + tail_height <= bottom {
            let remaining = total - body.emitted;
            let text = if document.read_truncated {
                format!("… {remaining} more lines (file truncated)")
            } else {
                format!("… {remaining} more lines")
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
                height: tail_height,
                block: LineBlock::Plain,
            });
            body.y += tail_height;
            body.content_right = body.content_right.max(padding + width);
        }
    }

    LaidOut {
        lines: body.lines,
        width: (body.content_right + padding).clamp(1, box_width.max(1) as i32) as u32,
        height: (body.y + padding).clamp(1, box_height.max(1) as i32) as u32,
    }
}

/// The lines of the document that fit inside the box, with `reserve` pixels of the
/// bottom edge held back for whatever the caller wants to add below them.
fn lay_out_body(
    document: &TextDoc,
    box_width: u32,
    box_height: u32,
    metrics: &TextMetrics,
    reserve: i32,
) -> BodyLayout {
    let padding = metrics.padding;
    let box_width = box_width.max(1) as i32;
    let left = padding;
    let right = (box_width - padding).max(left + 1);
    let bottom = (box_height.max(1) as i32 - padding).max(padding + 1) - reserve;
    let body_advance = metrics.advance[BODY_LEVEL as usize].max(1);

    let mut lines: Vec<LaidLine> = Vec::new();
    let mut y = padding;
    let mut content_right = left + body_advance * MIN_CONTENT_CHARS;
    let mut emitted = 0usize;

    for line in &document.lines {
        let height = line_height(line, metrics);
        if y + height > bottom {
            break;
        }

        let indent = left + metrics.indent(line.indent);
        let text_right = match line.block {
            LineBlock::Quote => right - metrics.quote_bar - 4,
            _ => right,
        };

        let visual_lines = if document.wrap {
            wrap_line(line, indent, text_right, metrics)
        } else {
            vec![clip_line(line, indent, text_right, metrics)]
        };

        let mut placed = false;
        for runs in visual_lines {
            if y + height > bottom {
                break;
            }

            let right_edge = runs
                .iter()
                .map(|run| run.x + run.width)
                .max()
                .unwrap_or(indent);
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
            break;
        }
        emitted += 1;
    }

    BodyLayout {
        lines,
        y,
        content_right,
        emitted,
    }
}

/// The runs of one line as it fits between `left` and `limit`, clipped at the
/// edge. This is what code, markup and art get: their columns line up, and a
/// wrapped line would break that.
fn clip_line(line: &DocLine, left: i32, limit: i32, metrics: &TextMetrics) -> Vec<LaidRun> {
    let mut x = left;
    let mut runs = Vec::new();

    for span in &line.spans {
        if x >= limit {
            break;
        }

        let advance = metrics.advance[(span.style.level as usize).min(SIZE_LEVELS - 1)].max(1);
        let available = ((limit - x) / advance) as usize;
        let characters = span
            .text
            .chars()
            .take(MAX_LINE_CHARS.min(available))
            .count();
        if characters == 0 {
            break;
        }

        let text: String = span.text.chars().take(characters).collect();
        let width = characters as i32 * advance;
        runs.push(LaidRun {
            text,
            style: span.style,
            x,
            width,
        });
        x += width;
    }

    runs
}

/// A word or the spaces between words, carrying the style it was colored with.
struct Token {
    text: String,
    style: TextStyle,
    space: bool,
}

/// Split a line into words and spaces, so a wrapped line breaks between words
/// instead of in the middle of one and keeps its colors across the break.
fn tokens_of(line: &DocLine) -> Vec<Token> {
    let mut tokens = Vec::new();

    for span in &line.spans {
        let mut text = String::new();
        let mut space = false;

        for character in span.text.chars() {
            let is_space = character == ' ';
            if !text.is_empty() && is_space != space {
                tokens.push(Token {
                    text: std::mem::take(&mut text),
                    style: span.style,
                    space,
                });
            }
            space = is_space;
            text.push(character);
        }

        if !text.is_empty() {
            tokens.push(Token {
                text,
                style: span.style,
                space,
            });
        }
    }

    tokens
}

/// One document line's visual lines, wrapped at the page edge. A word wider than
/// the page — a URL, a hash — is cut rather than allowed to run off it, and a
/// space that lands at a wrap point is dropped instead of starting a line.
fn wrap_line(line: &DocLine, left: i32, limit: i32, metrics: &TextMetrics) -> Vec<Vec<LaidRun>> {
    let mut lines: Vec<Vec<LaidRun>> = Vec::new();
    let mut runs: Vec<LaidRun> = Vec::new();
    let mut x = left;

    for token in tokens_of(line) {
        let advance = metrics.advance[(token.style.level as usize).min(SIZE_LEVELS - 1)].max(1);
        let characters: Vec<char> = token.text.chars().collect();
        let width = characters.len() as i32 * advance;

        if token.space {
            if runs.is_empty() || x + width > limit {
                continue;
            }
            runs.push(LaidRun {
                text: token.text,
                style: token.style,
                x,
                width,
            });
            x += width;
            continue;
        }

        if x > left && x + width > limit {
            lines.push(std::mem::take(&mut runs));
            x = left;
        }

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
                lines.push(std::mem::take(&mut runs));
                x = left;
            }
        }
    }

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

// ------------------------------------------------------------------ painting

/// A memory DC with a top-down 32-bit DIB section selected into it: the surface
/// GDI draws the preview's text on before it is read back as a frame.
struct DibSurface {
    dc: HDC,
    bitmap: HBITMAP,
    previous_bitmap: HGDIOBJ,
    bits: *mut u8,
    width: u32,
    height: u32,
}

impl DibSurface {
    fn create(width: u32, height: u32) -> Option<Self> {
        let dc = unsafe { CreateCompatibleDC(None) };
        if dc.0.is_null() {
            return None;
        }

        let info = BITMAPINFO {
            bmiHeader: BITMAPINFOHEADER {
                biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                biWidth: width as i32,
                biHeight: -(height as i32),
                biPlanes: 1,
                biBitCount: 32,
                biCompression: BI_RGB.0,
                ..Default::default()
            },
            ..Default::default()
        };

        let mut bits: *mut core::ffi::c_void = std::ptr::null_mut();
        let bitmap = unsafe { CreateDIBSection(dc, &info, DIB_RGB_COLORS, &mut bits, None, 0) };
        let Ok(bitmap) = bitmap else {
            unsafe {
                let _ = DeleteDC(dc);
            }
            return None;
        };

        if bits.is_null() {
            unsafe {
                let _ = DeleteObject(bitmap);
                let _ = DeleteDC(dc);
            }
            return None;
        }

        let previous_bitmap = unsafe { SelectObject(dc, bitmap) };

        Some(Self {
            dc,
            bitmap,
            previous_bitmap,
            bits: bits as *mut u8,
            width,
            height,
        })
    }

    /// The painted surface as BGRA. GDI writes color but not alpha, so the alpha
    /// byte is forced opaque here: a text preview is a page, and the layers above
    /// it composite it as one.
    fn pixels(&self) -> Vec<u8> {
        let length = self.width as usize * self.height as usize * 4;
        let mut pixels = unsafe { std::slice::from_raw_parts(self.bits, length) }.to_vec();
        for pixel in pixels.chunks_exact_mut(4) {
            pixel[3] = 255;
        }
        pixels
    }
}

impl Drop for DibSurface {
    fn drop(&mut self) {
        unsafe {
            let _ = SelectObject(self.dc, self.previous_bitmap);
            let _ = DeleteObject(self.bitmap);
            let _ = DeleteDC(self.dc);
        }
    }
}

/// Fonts are created per style and size, once per paint, and deleted with this
/// holder. The DC is handed back its original object before they go.
struct FontCache {
    fonts: Vec<(TextStyleKey, HFONT)>,
}

#[derive(PartialEq, Eq)]
struct TextStyleKey {
    level: u8,
    bold: bool,
    italic: bool,
    underline: bool,
    strike: bool,
}

impl TextStyleKey {
    fn of(style: &TextStyle) -> Self {
        Self {
            level: (style.level as usize).min(SIZE_LEVELS - 1) as u8,
            bold: style.bold,
            italic: style.italic,
            underline: style.underline,
            strike: style.strike,
        }
    }
}

impl FontCache {
    fn new() -> Self {
        Self { fonts: Vec::new() }
    }

    unsafe fn get(&mut self, style: &TextStyle, scale: f32) -> HFONT {
        let key = TextStyleKey::of(style);
        if let Some((_, font)) = self.fonts.iter().find(|(cached, _)| *cached == key) {
            return *font;
        }

        let pixels = scaled(LEVEL_FONT_PIXELS[key.level as usize], scale);
        let font = create_font(pixels, style);
        self.fonts.push((key, font));
        font
    }
}

impl Drop for FontCache {
    fn drop(&mut self) {
        for (_, font) in &self.fonts {
            unsafe {
                let _ = DeleteObject(*font);
            }
        }
    }
}

fn create_font(pixels: i32, style: &TextStyle) -> HFONT {
    let face: Vec<u16> = FONT_FACE.encode_utf16().chain(std::iter::once(0)).collect();

    unsafe {
        CreateFontW(
            pixels,
            0,
            0,
            0,
            if style.bold { 700 } else { 400 },
            style.italic as u32,
            style.underline as u32,
            style.strike as u32,
            DEFAULT_CHARSET.0 as u32,
            OUT_TT_PRECIS.0 as u32,
            CLIP_DEFAULT_PRECIS.0 as u32,
            CLEARTYPE_QUALITY.0 as u32,
            FIXED_PITCH.0 as u32 | FF_MODERN.0 as u32,
            PCWSTR(face.as_ptr()),
        )
    }
}

fn colorref(color: [u8; 3]) -> COLORREF {
    COLORREF(color[0] as u32 | ((color[1] as u32) << 8) | ((color[2] as u32) << 16))
}

/// Paint the background, the block decorations and then the text.
///
/// Every run is drawn with an opaque background rectangle, so the page color
/// fills the box and the pieces GDI leaves alone (the space a clipped line did
/// not use, the gaps between blocks) are already the right color.
unsafe fn paint(
    surface: &DibSurface,
    laid_out: &LaidOut,
    theme: &LoadedTheme,
    metrics: &TextMetrics,
) {
    let width = surface.width as i32;
    let page = rgb(theme.background());
    let band = blend(page, rgb(theme.foreground()), 0.08);

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

    let mut fonts = FontCache::new();
    let mut previous_font: Option<HGDIOBJ> = None;
    let mut selected: Option<TextStyleKey> = None;

    for line in &laid_out.lines {
        for run in &line.runs {
            let key = TextStyleKey::of(&run.style);
            if selected.as_ref() != Some(&key) {
                let font = fonts.get(&run.style, metrics.scale);
                let previous = SelectObject(surface.dc, font);
                if previous_font.is_none() {
                    previous_font = Some(previous);
                }
                selected = Some(key);
            }

            let wide: Vec<u16> = run.text.encode_utf16().collect();
            if wide.is_empty() {
                continue;
            }

            let run_background = match (run.style.background, line.block) {
                (Some(color), _) => color,
                (None, LineBlock::Code) => band,
                (None, _) => page,
            };

            SetTextColor(surface.dc, colorref(run.style.foreground));
            SetBkColor(surface.dc, colorref(run_background));
            SetBkMode(surface.dc, OPAQUE);

            let rect = RECT {
                left: run.x,
                top: line.top,
                right: run.x + run.width,
                bottom: line.top + line.height,
            };
            let _ = ExtTextOutW(
                surface.dc,
                run.x,
                line.top,
                ETO_OPAQUE | ETO_CLIPPED,
                Some(&rect),
                PCWSTR(wide.as_ptr()),
                wide.len() as u32,
                None,
            );
        }
    }

    if let Some(previous) = previous_font {
        let _ = SelectObject(surface.dc, previous);
    }
    drop(fonts);
}

fn fill_rect(surface: &DibSurface, rect: RECT, color: [u8; 3]) {
    let width = surface.width as i32;
    let height = surface.height as i32;
    let left = rect.left.clamp(0, width);
    let right = rect.right.clamp(0, width);
    let top = rect.top.clamp(0, height);
    let bottom = rect.bottom.clamp(0, height);
    if left >= right || top >= bottom {
        return;
    }

    let pixels = unsafe {
        std::slice::from_raw_parts_mut(surface.bits, width as usize * height as usize * 4)
    };

    for y in top..bottom {
        let row = y as usize * width as usize * 4;
        for x in left..right {
            let index = row + x as usize * 4;
            pixels[index] = color[2];
            pixels[index + 1] = color[1];
            pixels[index + 2] = color[0];
            pixels[index + 3] = 255;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{
        cp437_char, decode_text, expand_tabs, legacy_encoding, markdown_document, measure, render,
        strip_rtf, wants_ansi, LegacyEncoding, LineBlock, TextPreviewOptions, MAX_LINE_CHARS,
    };
    use crate::config::{MarkdownMode, TextTheme};
    use crate::text_theme;
    use std::path::{Path, PathBuf};

    fn options(theme: TextTheme, mode: MarkdownMode) -> TextPreviewOptions {
        TextPreviewOptions {
            theme,
            markdown_mode: mode,
        }
    }

    fn light_theme() -> &'static text_theme::LoadedTheme {
        text_theme::loaded(TextTheme::Light).unwrap()
    }

    /// Write `contents` to a file of its own and hand back the path. The name
    /// keeps whatever extension it is given, since that is what selects the
    /// renderer under test.
    fn fixture(name: &str, contents: &[u8]) -> PathBuf {
        let mut path = std::env::temp_dir();
        path.push(format!("rust-hover-preview-{}-{name}", std::process::id()));
        std::fs::write(&path, contents).unwrap();
        path
    }

    #[test]
    fn byte_order_marks_decide_the_encoding() {
        let utf8_bom = [0xEF, 0xBB, 0xBF, b'h', b'i'];
        assert_eq!(
            decode_text(&utf8_bom, LegacyEncoding::Windows1252, false).unwrap(),
            "hi"
        );

        let utf16_le = [0xFF, 0xFE, b'h', 0x00, b'i', 0x00];
        assert_eq!(
            decode_text(&utf16_le, LegacyEncoding::Windows1252, false).unwrap(),
            "hi"
        );

        let utf16_be = [0xFE, 0xFF, 0x00, b'h', 0x00, b'i'];
        assert_eq!(
            decode_text(&utf16_be, LegacyEncoding::Windows1252, false).unwrap(),
            "hi"
        );

        // UTF-16 without a mark is still recognized from the NUL pattern.
        assert_eq!(
            decode_text(
                &[b'h', 0x00, b'i', 0x00],
                LegacyEncoding::Windows1252,
                false
            )
            .unwrap(),
            "hi"
        );
    }

    #[test]
    fn legacy_bytes_land_on_the_right_characters() {
        // 0xE9 is e-acute in both code pages, 0xDB is a block in CP437 only.
        assert_eq!(
            decode_text(&[0xE9], LegacyEncoding::Windows1252, false).unwrap(),
            "é"
        );
        assert_eq!(cp437_char(0xDB), '█');
        assert_eq!(
            decode_text(&[0xDB], LegacyEncoding::Cp437, false).unwrap(),
            "█"
        );

        // Windows-1252's own range: 0x80 is the euro sign, not a NUL and not a
        // control character.
        assert_eq!(
            decode_text(&[0x80], LegacyEncoding::Windows1252, false).unwrap(),
            "€"
        );
    }

    #[test]
    fn binary_files_are_not_previews() {
        // No NUL byte, but the bytes that do not decode as UTF-8 are not text.
        assert!(decode_text(&[0x00, 0x01, 0x02], LegacyEncoding::Windows1252, false).is_none());
        // A ZIP header is valid UTF-8 and still not a page.
        assert!(decode_text(
            &[0x50, 0x4B, 0x03, 0x04],
            LegacyEncoding::Windows1252,
            false
        )
        .is_none());
        assert!(decode_text(
            b"plain text, control free",
            LegacyEncoding::Windows1252,
            false
        )
        .is_some());
        assert_eq!(
            decode_text(&[], LegacyEncoding::Windows1252, false).unwrap(),
            ""
        );
    }

    #[test]
    fn tabs_expand_to_stops_and_carriage_returns_go() {
        assert_eq!(expand_tabs("a\tb"), "a   b");
        assert_eq!(expand_tabs("abcd\te"), "abcd    e");
        assert_eq!(expand_tabs("a\r\nb"), "a\nb");
        assert_eq!(expand_tabs("\t"), "    ");
    }

    #[test]
    fn rtf_reduces_to_its_text() {
        // `\u8364` is the euro sign by RTF's decimal numbering, and the `?` that
        // follows it is that character's fallback, which the page does not need.
        let source = "{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\f0\\fs20 Hello\\par \
                      World \\'e9 \\u8364? done{\\*\\generator Word}}";
        let lines = strip_rtf(source);
        let text = lines.join("\n");

        assert!(text.contains("Hello"), "{text}");
        assert!(text.contains("World é € done"), "{text}");
        assert!(!text.contains("Arial"), "{text}");
        assert!(!text.contains("generator"), "{text}");
        assert_eq!(lines.len(), 2);
    }

    #[test]
    fn ansi_escapes_become_colors() {
        let document = super::ansi_document(
            &["\u{1B}[31mred\u{1B}[0m plain".to_string()],
            light_theme(),
            false,
        );

        let spans = &document.lines[0].spans;
        assert_eq!(spans.len(), 2);
        assert_eq!(spans[0].text, "red");
        assert_eq!(spans[1].text, " plain");
        assert_ne!(spans[0].style.foreground, spans[1].style.foreground);

        // An NFO is read as CP437 and keeps its escapes out of the text.
        assert!(wants_ansi("nfo"));
        assert!(matches!(legacy_encoding("nfo"), LegacyEncoding::Cp437));
        assert!(matches!(
            legacy_encoding("txt"),
            LegacyEncoding::Windows1252
        ));
    }

    #[test]
    fn markdown_becomes_blocks_and_inline_runs() {
        let source = "# Title\n\nSome *emphasis* and `code`.\n\n- one\n- two\n\n```rust\nfn main() {}\n```\n";
        let document = markdown_document(source, light_theme(), 12, false);

        let heading = &document.lines[0];
        assert!(heading.spans.iter().any(|span| span.text == "Title"));
        assert!(heading
            .spans
            .iter()
            .any(|span| span.style.bold && span.style.level > 0));

        let body = document
            .lines
            .iter()
            .find(|line| line.spans.iter().any(|span| span.text.contains("emphasis")))
            .expect("the paragraph should be laid out");
        assert!(body
            .spans
            .iter()
            .any(|span| span.text.contains("emphasis") && span.style.italic));
        assert!(body
            .spans
            .iter()
            .any(|span| span.text == "code" && span.style.background.is_some()));

        assert!(document.lines.iter().any(|line| line
            .spans
            .first()
            .map(|span| span.text.as_str())
            == Some("• ")));
        assert!(document
            .lines
            .iter()
            .any(|line| line.block == LineBlock::Code && !line.spans.is_empty()));
    }

    #[test]
    fn markdown_source_mode_keeps_the_markup() {
        let source = "# Title\n\nSome *emphasis* here.\n";
        let source_mode = options(TextTheme::Light, MarkdownMode::Source);
        let rendered_mode = options(TextTheme::Light, MarkdownMode::Rendered);
        let path = fixture("markdown-modes.md", source.as_bytes());

        let rendered = super::document(&path, rendered_mode).unwrap();
        let highlighted = super::document(&path, source_mode).unwrap();

        let text_of = |document: &super::TextDoc| {
            document
                .lines
                .iter()
                .flat_map(|line| line.spans.iter())
                .map(|span| span.text.clone())
                .collect::<String>()
        };

        // The source mode shows the markup itself; the rendered mode shows the
        // document it describes.
        assert!(text_of(&highlighted).contains("# Title"));
        assert!(!text_of(&rendered).contains("# Title"));
        assert!(text_of(&rendered).contains("Title"));
        assert!(text_of(&rendered).contains("emphasis"));
        assert!(!text_of(&rendered).contains('*'));

        // And the two modes measure to different boxes.
        let rendered_box = measure(&path, 1200, 900, 96, rendered_mode).unwrap();
        let source_box = measure(&path, 1200, 900, 96, source_mode).unwrap();
        assert_ne!(rendered_box, source_box);

        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn a_text_preview_measures_and_paints_to_the_same_box() {
        let path = fixture("measure.rs", b"fn main() {\n    println!(\"hello\");\n}\n");
        let options = options(TextTheme::Dark, MarkdownMode::Rendered);

        let (width, height) = measure(&path, 1400, 900, 96, options).expect("measured");

        // The box hugs the content: three short lines, not the whole display.
        assert!(width < 1400, "{width}");
        assert!(height < 200, "{height}");

        let (pixels, rendered_width, rendered_height) =
            render(&path, width, height, 96, options).expect("rendered");

        assert_eq!((rendered_width, rendered_height), (width, height));
        assert_eq!(pixels.len(), (width * height * 4) as usize);

        // Every pixel is opaque, and the theme's background is what the page is
        // painted with.
        assert!(pixels.chunks_exact(4).all(|pixel| pixel[3] == 255));
        let background = text_theme::loaded(TextTheme::Dark).unwrap().background();
        assert_eq!(
            &pixels[0..3],
            &[background.b, background.g, background.r],
            "the top-left corner should be the page background"
        );

        // Something was drawn: a completely uniform frame would mean the text
        // never reached the surface.
        let uniform = pixels
            .chunks_exact(4)
            .all(|pixel| pixel[0] == pixels[0] && pixel[1] == pixels[1] && pixel[2] == pixels[2]);
        assert!(!uniform, "the preview painted no text");

        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn the_light_and_dark_themes_paint_different_pages() {
        let path = fixture("themes.txt", b"hello world\n");
        let (width, height) = measure(
            &path,
            800,
            600,
            96,
            options(TextTheme::Light, MarkdownMode::Rendered),
        )
        .unwrap();

        let (light, _, _) = render(
            &path,
            width,
            height,
            96,
            options(TextTheme::Light, MarkdownMode::Rendered),
        )
        .unwrap();
        let (dark, _, _) = render(
            &path,
            width,
            height,
            96,
            options(TextTheme::Dark, MarkdownMode::Rendered),
        )
        .unwrap();

        assert!(light[0] > 200, "the light theme should paint a light page");
        assert!(dark[0] < 100, "the dark theme should paint a dark page");

        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn long_lines_are_clipped_instead_of_widening_the_box() {
        let long = "x".repeat(MAX_LINE_CHARS + 500);
        // A Rust file has a syntax definition, so its columns are kept rather
        // than wrapped.
        let path = fixture("long-line.rs", long.as_bytes());
        let options = options(TextTheme::Light, MarkdownMode::Rendered);

        let (width, height) = measure(&path, 500, 400, 96, options).unwrap();
        assert!(width <= 500, "{width}");
        assert!(height < 100, "one clipped line, not {height}");

        let (pixels, rendered_width, rendered_height) =
            render(&path, width, 400, 96, options).unwrap();
        assert_eq!(
            pixels.len(),
            (rendered_width * rendered_height * 4) as usize
        );

        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn prose_wraps_where_code_is_clipped() {
        let line = "word ".repeat(120);
        let prose = fixture("wrapping.txt", line.as_bytes());
        let code = fixture("wrapping.rs", line.as_bytes());
        let options = options(TextTheme::Light, MarkdownMode::Rendered);

        let (prose_width, prose_height) = measure(&prose, 400, 2000, 96, options).unwrap();
        let (code_width, code_height) = measure(&code, 400, 2000, 96, options).unwrap();

        assert!(prose_width <= 400 && code_width <= 400);
        assert!(
            prose_height > 100,
            "prose should continue on the next line: {prose_height}"
        );
        assert!(
            code_height < 100,
            "code should stay on one line: {code_height}"
        );

        let _ = std::fs::remove_file(&prose);
        let _ = std::fs::remove_file(&code);
    }

    #[test]
    fn an_empty_file_has_no_preview() {
        let path = fixture("empty.txt", b"\n\n");
        assert!(measure(
            &path,
            800,
            600,
            96,
            options(TextTheme::Light, MarkdownMode::Rendered)
        )
        .is_none());
        let _ = std::fs::remove_file(&path);
        assert!(measure(
            Path::new("does-not-exist.txt"),
            800,
            600,
            96,
            options(TextTheme::Light, MarkdownMode::Rendered)
        )
        .is_none());
    }

    #[test]
    fn a_long_file_says_how_much_is_left() {
        let mut source = String::new();
        for line in 0..500 {
            source.push_str(&format!("line {line}\n"));
        }
        let path = fixture("long.txt", source.as_bytes());
        let options = options(TextTheme::Light, MarkdownMode::Rendered);

        // Room for a few lines, nowhere near five hundred.
        let (width, height) = measure(&path, 600, 120, 96, options).unwrap();
        let document = super::document(&path, options).unwrap();

        let dc = unsafe { windows::Win32::Graphics::Gdi::CreateCompatibleDC(None) };
        assert!(!dc.0.is_null());
        let metrics = super::TextMetrics::new(dc, 96).unwrap();
        let laid_out = super::layout(&document, width, height, &metrics, light_theme());
        unsafe {
            let _ = windows::Win32::Graphics::Gdi::DeleteDC(dc);
        }

        let last = laid_out.lines.last().expect("a full page");
        let note = &last.runs[0].text;
        assert!(note.starts_with('…'), "{note}");
        assert!(note.contains("more lines"), "{note}");

        // The note is inside the box it was laid out for.
        assert!(last.top + last.height <= height as i32, "{note}");

        let _ = std::fs::remove_file(&path);
    }
}

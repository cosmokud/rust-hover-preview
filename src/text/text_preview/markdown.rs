//! Markdown as the document it describes: one walk over the CommonMark events,
//! turning block structure into lines and indent and inline structure into runs.
//!
//! A rendered document is the one producer that has to read all of its input: a
//! line's place in the document is only known once the block before it has been
//! walked, so there is no state to resume from part-way through. It is a single
//! linear pass with no grammar work, and only the lines inside the requested
//! window are kept — the rest are counted and dropped as they go past.

use super::document::{
    highlight_line, DocLine, LineBlock, Span, TextDoc, TextWindow, MAX_DOC_LINES, SYNTAXES,
};
use crate::text::text_paint::{blend, rgb, text_style, TextStyle, BODY_LEVEL};
use crate::text::text_theme::LoadedTheme;
use pulldown_cmark::{CodeBlockKind, Event, Options, Parser, Tag, TagEnd};
use syntect::easy::HighlightLines;

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
pub(super) fn window_markdown(
    doc: &TextDoc,
    theme: &LoadedTheme,
    first: usize,
    last: usize,
) -> TextWindow {
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
pub(super) fn rendered_line_count(doc: &TextDoc, theme: &LoadedTheme) -> usize {
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

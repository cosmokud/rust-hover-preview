//! The frame API: the options a preview is built with, the scrollbar it draws,
//! and the painted lines a pointer or a selection is measured against.
//!
//! `measure` and `render_scrolled` are the two calls the rest of the app makes:
//! the first answers "how big is this preview?" before anything is painted, the
//! second lays the same document out for the box the layout settled on and fills
//! it. What is underneath them — the document ([`super::document`]) and the box
//! it is laid out in ([`super::layout`]) — is asked for here, not owned.

use super::document::document;
use super::layout::{layout, paint, scrollbar_geometry};
use crate::config::config::{MarkdownMode, TextTheme};
use crate::text::text_paint::{DibSurface, TextMetrics};
use crate::text::text_theme;
use std::path::Path;
use windows::Win32::Graphics::Gdi::{CreateCompatibleDC, DeleteDC};

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
    /// It is not one of the configuration's answers — a preview is built with it
    /// when it is pinned, and without it otherwise (see `current_text_options`).
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
    /// The document line this painted line belongs to, and the character of that
    /// line its first run starts at: what a point on the frame is turned back into,
    /// and what a selection is measured against (see `Selection`).
    pub source: (usize, usize),
}

/// A range of a text preview a reader has selected: where the drag started and
/// where it is now, both as (document line, character).
///
/// The place is the document's own rather than the frame's, which is what makes a
/// selection outlive the box it was made in: a page drawn again — scrolled, resized,
/// wrapped at another width — draws the same text somewhere else, and a selection
/// that named a line of the frame would name whatever moved into that line instead
/// (see `keep_text_place`).
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

/// The line of painted text nearest a point, by its middle.
fn nearest_line(lines: &[FrameLine], y: i32) -> Option<(usize, &FrameLine)> {
    lines.iter().enumerate().min_by_key(|(_, line)| {
        let middle = line.top + line.height / 2;
        (middle - y).abs()
    })
}

/// Where a point falls inside one painted line: the character boundary nearest it, which is the
/// caret a press places, and the character whose cell the point is actually in, which is the one a
/// word is taken around. The two differ in the right half of a cell, where the caret sits after a
/// character the point is still on.
fn point_in_line(line: &FrameLine, x: i32) -> (usize, usize) {
    let mut characters = 0usize;
    let mut caret = 0usize;
    let mut character = 0usize;
    let mut closest_distance = i32::MAX;

    for run in &line.runs {
        let count = run.text.chars().count() as i32;
        if count == 0 {
            continue;
        }
        let advance = (run.width as f32 / count as f32).max(1.0);
        let within = (x - run.x) as f32 / advance;
        let boundary = (within.round() as i32).clamp(0, count);
        let edge = run.x + (boundary as f32 * advance).round() as i32;
        let distance = (edge - x).abs();
        if distance < closest_distance {
            closest_distance = distance;
            caret = characters + boundary as usize;
            character = characters + (within.floor() as i32).clamp(0, count - 1) as usize;
        }
        characters += count as usize;
    }

    (caret, character)
}

/// Where a point falls in a text preview: the document line under it and the
/// character within that line.
///
/// A point between two lines takes the nearer one, and a point past the end of a
/// line takes the nearest end of it, so a drag that runs off the text still selects
/// the line it was on rather than losing it. What comes back is a place in the
/// document, so a point on a wrapped tail names the character it is drawn under
/// rather than the line of the frame it happens to be (see `FrameLine::source`).
pub fn position_in(lines: &[FrameLine], x: i32, y: i32) -> (usize, usize) {
    let Some((_, line)) = nearest_line(lines, y) else {
        return (0, 0);
    };

    (line.source.0, line.source.1 + point_in_line(line, x).0)
}

/// The kind of character a double click takes a run of: a word is letters, digits and the
/// underscore, a gap is space, and everything else — a bracket, a punctuation mark — is a run of
/// its own, so a double click takes the whole `->` or `==` and not just one of its marks.
#[derive(Clone, Copy, PartialEq, Eq)]
enum RunKind {
    Word,
    Space,
    Other,
}

fn run_kind(character: char) -> RunKind {
    if character.is_alphanumeric() || character == '_' {
        RunKind::Word
    } else if character.is_whitespace() {
        RunKind::Space
    } else {
        RunKind::Other
    }
}

/// The word a point on painted lines is in, as the range a double click selects: the run of the
/// point's own kind of character, taken from the character the point is on rather than from the
/// caret beside it, so a press on the right half of a letter still takes the word that letter is
/// in.
///
/// A line with nothing on it has no word to take, and the range comes back empty there.
pub fn word_in(lines: &[FrameLine], x: i32, y: i32) -> Selection {
    let Some((_, line)) = nearest_line(lines, y) else {
        return Selection {
            anchor: (0, 0),
            caret: (0, 0),
        };
    };

    let (_, at) = point_in_line(line, x);
    let characters: Vec<char> = line.runs.iter().flat_map(|run| run.text.chars()).collect();
    let Some(&character) = characters.get(at) else {
        return Selection {
            anchor: line.source,
            caret: line.source,
        };
    };

    let kind = run_kind(character);
    let before = characters[..at]
        .iter()
        .rev()
        .take_while(|character| run_kind(**character) == kind)
        .count();
    let after = characters[at + 1..]
        .iter()
        .take_while(|character| run_kind(**character) == kind)
        .count();

    Selection {
        anchor: (line.source.0, line.source.1 + at - before),
        caret: (line.source.0, line.source.1 + at + 1 + after),
    }
}

/// The selection a drag that followed a double press makes: the word that press took, grown to the
/// whole of the run of characters the hand has reached.
///
/// A drag grown this way takes whole runs rather than single characters, which is what keeps the
/// word a double press took whole: a hand that has not left it yet leaves it as it is, where a drag
/// by the character would have cut it down to the character it reached. A hand that has left it
/// takes the run it is over — a word is taken whole, and a gap is a run of its own, so a hand in
/// the gap between two words takes the gap and stops where the word after it begins.
///
/// Which side of the word the hand has left decides which end of it stays put: past the word's own
/// end the selection grows from its start, and back before it the word's end is what the hand is
/// coming back from. A line with nothing on it has no word to take, and a drag from there begins
/// where the hand has reached (see `word_in` and `position_in`).
pub fn extend_word_selection(lines: &[FrameLine], word: Selection, x: i32, y: i32) -> Selection {
    let (start, end) = word.ordered();
    let (before, after) = word_in(lines, x, y).ordered();
    let at = position_in(lines, x, y);

    if at >= end {
        Selection {
            anchor: start,
            caret: after.max(end),
        }
    } else if at <= start {
        Selection {
            anchor: end,
            caret: before.min(start),
        }
    } else {
        Selection {
            anchor: start,
            caret: end,
        }
    }
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
            source: line.source,
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

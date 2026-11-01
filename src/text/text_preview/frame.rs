//! The frame API: the options a preview is built with, the scrollbar it draws,
//! and the painted lines a pointer or a selection is measured against.
//!
//! `measure` and `render_scrolled` are the two calls the rest of the app makes:
//! the first answers "how big is this preview?" before anything is painted, the
//! second lays the same document out for the box the layout settled on and fills
//! it. What is underneath them — the document ([`super::document`]) and the box
//! it is laid out in ([`super::layout`]) — is asked for here, not owned.

use super::document::{document, document_held};
use super::layout::{layout, paint, scrollbar_geometry};
use crate::config::config::{MarkdownMode, TextTheme};
use crate::text::text_paint::{DibSurface, TextMetrics};
use crate::text::text_theme;
use std::path::Path;
use windows::Win32::Foundation::RECT;
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
    /// Left edge of the column this line's text is drawn in: where the line's own text begins, past
    /// the gutter the page keeps for its numbers and past the indent the line carries. It is where a
    /// place on a line with no run to measure stands, which is what a caret at the start of an empty
    /// line is (see `x_within_line`).
    pub left: i32,
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

/// Where the caret of a page of text is drawn, when it is drawn at all: the character a vertical
/// bar stands before, and which of its two phases the blink is in.
///
/// **It is a mark over the composed frame rather than a part of it.** A selection is painted into
/// the frame because it is a colour laid under the text, but a caret is a mark laid *over* a
/// character — one that has to go away again every half second, and a frame that was drawn again
/// for each phase would restyle and repaint a screenful of text twice a second. So the caret's
/// place travels beside the frame and the compositor draws the bar (see `paint_text_caret`), and
/// only a phase change asks for a repaint.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextCaret {
    /// The document place the bar stands before, as `Selection`'s two ends are measured: the bar is
    /// drawn at the left edge of that character, and at the right edge of the line where the place
    /// is the line's end.
    pub at: (usize, usize),
    /// Whether the bar is drawn this half of the blink.
    pub lit: bool,
}

/// Every place a find bar's query was found in the document, and which of them the reader is on:
/// what a frame highlights over the page, under the selection the reader made for themselves.
///
/// It is owned rather than borrowed because it is taken out of the page's state before the frame is
/// drawn — the lock that holds the state is let go of while the page is laid out and painted — and
/// it is a list rather than a range because a search is read as a whole: what the frame draws is the
/// part of it that is on screen.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FindHighlights {
    pub matches: Vec<Selection>,
    /// Index into `matches` of the one the reader is on, drawn in a colour of its own so that the
    /// count beside the query can be read off the page as well as off the card.
    pub current: Option<usize>,
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

/// Where the bar for `caret` stands among painted lines, in frame coordinates, or nothing where the
/// place is not one of the lines the frame drew.
///
/// The place is a document one, and a frame draws a screenful of a document — so a caret scrolled
/// off the top or the bottom is a caret with nothing to draw, which is what a bar drawn at the
/// frame's edge would get wrong. A wrapped line is asked for the painted part its own characters
/// fall in, so a caret past the first row of a long line is drawn on the row it is on and not at the
/// start of the line (see `FrameLine::source`).
///
/// The bar's height is the painted line's, and its width is the advance of the character it stands
/// before — which is the run's own width divided by its characters, the same estimate `point_in_line`
/// measures a press with.
pub fn caret_rect(lines: &[FrameLine], caret: (usize, usize)) -> Option<RECT> {
    let (line, character) = caret;
    let painted = lines
        .iter()
        .filter(|painted| painted.source.0 == line)
        .find(|painted| {
            let start = painted.source.1;
            let end = start + painted_text_characters(painted);
            character >= start && character < end
        })
        .or_else(|| {
            // The end of a line is the place past its last character, which is a place no run
            // holds: it is drawn at the right edge of the last run the line has.
            lines
                .iter()
                .filter(|painted| painted.source.0 == line)
                .max_by_key(|painted| painted.source.1)
                .filter(|painted| painted.source.1 + painted_text_characters(painted) == character)
        })?;

    let (x, advance) = x_within_line(painted, character.saturating_sub(painted.source.1));

    Some(RECT {
        left: x,
        top: painted.top,
        right: x + advance.max(1),
        bottom: painted.top + painted.height,
    })
}

/// How many characters one painted line shows, which is what tells a wrapped line's first row from
/// its second (see `FrameLine::source`).
fn painted_text_characters(line: &FrameLine) -> usize {
    line.runs.iter().map(|run| run.text.chars().count()).sum()
}

/// The left edge of character `offset` of a painted line, and the width of the character there —
/// which is the next character's width at the end of the line, so a bar at the end of a run is as
/// wide as the text it stands after rather than as wide as whatever is drawn beside it.
fn x_within_line(line: &FrameLine, offset: usize) -> (i32, i32) {
    let mut before = 0usize;

    for run in &line.runs {
        let count = run.text.chars().count();
        if count == 0 {
            continue;
        }
        let advance = (run.width as f32 / count as f32).max(1.0);

        if offset < before + count {
            let within = offset - before;
            return (
                run.x + (within as f32 * advance).round() as i32,
                advance.round().max(1.0) as i32,
            );
        }
        before += count;
    }

    // Past the last run: the end of the line, at the right edge of the run before it. A line with no
    // run at all is a line with nothing to stand after, so the place is the left edge of the column
    // its text is drawn in rather than the frame's — an empty line is inside the text, not beside the
    // numbers (see `FrameLine::left`).
    match line.runs.last() {
        Some(run) => (
            run.x + run.width,
            (run.width as f32 / run.text.chars().count().max(1) as f32)
                .round()
                .max(1.0) as i32,
        ),
        None => (line.left, 1),
    }
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
    if is_word_character(character) {
        RunKind::Word
    } else if character.is_whitespace() {
        RunKind::Space
    } else {
        RunKind::Other
    }
}

/// Whether a character is part of a word: letters, digits and the underscore, which is exactly the
/// run `run_kind` calls a word — so a double click, a whole-word search and `Ctrl`+arrow over a
/// find bar's query all take the same characters to be one (see `crate::text::text_preview::document`
/// and `run_at`).
pub fn is_word_character(character: char) -> bool {
    character.is_alphanumeric() || character == '_'
}

/// The run of characters a double click at `at` takes, in byte indices: the rule `word_in` reads
/// over painted lines, asked of one string that is not laid out as lines at all — the find bar's
/// query. A gap is a run of its own and every mark is a run of its own, so a double click on `->`
/// takes the whole of it and not just one of its marks.
///
/// `at` is a byte index on a character boundary; one past the end of the text takes nothing.
pub fn run_at(text: &str, at: usize) -> (usize, usize) {
    let Some(character) = text.get(at..).and_then(|tail| tail.chars().next()) else {
        return (text.len(), text.len());
    };

    let kind = run_kind(character);
    let start = text[..at]
        .char_indices()
        .rev()
        .take_while(|(_, before)| run_kind(*before) == kind)
        .last()
        .map_or(at, |(before, _)| before);
    let end = text[at..]
        .char_indices()
        .take_while(|(_, after)| run_kind(*after) == kind)
        .map(|(after, character)| at + after + character.len_utf8())
        .last()
        .unwrap_or(at);

    (start, end)
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

/// The line a point on painted lines is in, as the range a third press selects: the whole of the
/// document line the painted line under the point belongs to, from its first character to the break
/// at its end.
///
/// It is the document's line rather than the painted one, which is what a reader means by "the
/// line": a line too long for the box is drawn as several painted lines, and they are one line all
/// the way down (see `FrameLine::source`). The break at the end is taken with it — the end of a
/// line is what makes it one, and a copy of the line pastes back as a line of its own — and it
/// paints nothing, so the highlight still covers the text the line has and no more.
pub fn line_in(lines: &[FrameLine], y: i32) -> Selection {
    let Some((_, line)) = nearest_line(lines, y) else {
        return Selection {
            anchor: (0, 0),
            caret: (0, 0),
        };
    };

    Selection {
        anchor: (line.source.0, 0),
        caret: (line.source.0 + 1, 0),
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

/// How many screenfuls have been laid out and painted, which is the work a key press on a page
/// must not cause once the page is on screen and the caret is already on it (see `render_scrolled`).
#[cfg(test)]
static FRAME_RENDERS: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

/// The number of screenfuls painted (see `FRAME_RENDERS`).
#[cfg(test)]
pub(crate) fn frame_render_count() -> u64 {
    FRAME_RENDERS.load(std::sync::atomic::Ordering::Relaxed)
}

/// How many screenfuls were the frame already on screen moved rather than painted afresh (see
/// `PaintedFrame`), which is the half of a scroll's cost a test can see without a clock.
#[cfg(test)]
static SHIFTED_FRAMES: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

/// The number of frames shifted rather than painted (see `SHIFTED_FRAMES`).
#[cfg(test)]
pub(crate) fn shifted_frame_count() -> u64 {
    SHIFTED_FRAMES.load(std::sync::atomic::Ordering::Relaxed)
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
#[allow(clippy::too_many_arguments)] // The box, the options and the two things drawn over the page are one frame.
pub fn render_scrolled(
    path: &Path,
    first_line: usize,
    width: u32,
    height: u32,
    dpi: u32,
    options: TextPreviewOptions,
    selection: Option<Selection>,
    find: Option<FindHighlights>,
) -> Option<TextFrame> {
    if width == 0 || height == 0 {
        return None;
    }

    render_scrolled_from(
        document(path, options)?,
        None,
        first_line,
        width,
        height,
        dpi,
        options,
        selection,
        find,
    )
}

/// The same frame for a page that is already on screen, drawn of the document that drew it rather
/// than of a fresh look at the file (see `document_held`).
///
/// A page scrolled a line, a caret moved and a selection dragged all lay the page out again in the
/// same box, and every one of them used to ask the disk first: an arrow at the bottom of a viewport
/// is a scroll, and a held arrow is one every repeat. The document is the page's own either way —
/// what the stat could say about a file that changed is said by the next page put up (see
/// `render_scrolled`).
#[allow(clippy::too_many_arguments)] // The box, the options and the two things drawn over the page are one frame.
pub fn render_scrolled_held(
    previous: Option<PaintedFrame<'_>>,
    path: &Path,
    first_line: usize,
    width: u32,
    height: u32,
    dpi: u32,
    options: TextPreviewOptions,
    selection: Option<Selection>,
    find: Option<FindHighlights>,
) -> Option<TextFrame> {
    if width == 0 || height == 0 {
        return None;
    }

    render_scrolled_from(
        document_held(path, options).or_else(|| document(path, options))?,
        previous,
        first_line,
        width,
        height,
        dpi,
        options,
        selection,
        find,
    )
}

/// A frame already on screen: the pixels and the painted lines it was drawn with, which is what a
/// page scrolled a line or two is moved from rather than laid out and painted again.
///
/// It is a borrow of what is already on screen rather than another copy of it — a screenful is a
/// megabyte and a half — and it is handed to `render_scrolled_held` by the one caller that has the
/// page in hand (see `scroll_text_preview`).
pub struct PaintedFrame<'a> {
    /// The frame as it was drawn, as `TextFrame::pixels` was: BGRA, `width * height * 4` bytes.
    pub pixels: &'a [u8],
    pub width: u32,
    pub height: u32,
    /// The document line the frame starts at, and the lines it painted: a frame asked for at another
    /// place is the same frame moved, and the lines are how far.
    pub first_line: usize,
    pub lines: &'a [FrameLine],
    /// The two things painted *into* a frame that its lines do not carry. A page whose selection or
    /// search has changed since the frame was drawn is a page painted with something else, and its
    /// pixels are not the rows a move would take with it.
    pub selection: Option<Selection>,
    pub find: Option<FindHighlights>,
}

/// The layout and the paint of a frame whose document is already in hand, for the two callers above
/// to differ only in which document they hand it.
#[allow(clippy::too_many_arguments)] // The box, the options and the two things drawn over the page are one frame.
fn render_scrolled_from(
    document: std::sync::Arc<super::document::CachedDocument>,
    previous: Option<PaintedFrame<'_>>,
    first_line: usize,
    width: u32,
    height: u32,
    dpi: u32,
    options: TextPreviewOptions,
    selection: Option<Selection>,
    find: Option<FindHighlights>,
) -> Option<TextFrame> {
    #[cfg(test)]
    FRAME_RENDERS.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
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

    // A page moved under the reader by a line or two is the frame already on screen with its rows
    // shifted, and only the rows that came in from an edge painted — a selection moved by a caret
    // held down is the commonest thing a pinned page does, and laying the whole screenful out and
    // painting every run of it again is 4.5 ms of the 6.9 a press cost (see `PaintedFrame` and
    // `MOST_ROWS_A_FRAME_IS_SHIFTED`).
    let painted_on_screen = previous.as_ref().map(|previous| PaintedFrame {
        pixels: previous.pixels,
        width: previous.width,
        height: previous.height,
        first_line: previous.first_line,
        lines: previous.lines,
        selection: previous.selection,
        find: previous.find.clone(),
    });
    let shifted = painted_on_screen.and_then(|previous| {
        the_shift_of(previous, &laid_out, selection, find.as_ref(), width, height)
    });

    let damage = match shifted {
        Some((moved, content_top, content_bottom, damage)) => {
            #[cfg(test)]
            SHIFTED_FRAMES.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
            copy_the_rows_on_screen(
                &surface,
                previous.as_ref()?.pixels,
                moved,
                content_top,
                content_bottom,
            );
            Some(damage)
        }
        None => None,
    };

    unsafe {
        paint(
            &surface,
            &laid_out,
            theme,
            &metrics,
            scrollbar.as_ref(),
            selection,
            find.as_ref(),
            damage,
        );
    }

    let lines = laid_out
        .lines
        .iter()
        .map(|line| FrameLine {
            top: line.top,
            height: line.height,
            left: line.left,
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

/// How far a frame may be moved under the reader and still be the frame already on screen, in rows.
///
/// A caret that has reached the edge of the page scrolls it by a line or three, and a wheel notch a
/// few more: past this the whole screenful is painted, which is also the answer when the page is not
/// the same page at all. The bound is what keeps a shift a band rather than the frame — a hundred
/// rows of a tall page is still a paint of the rows that came in and nothing more.
const MOST_ROWS_A_FRAME_IS_SHIFTED: i32 = 120;

/// How far the frame already on screen has to move to be the frame just laid out, and the band of
/// rows that is left to paint once it has: `(rows moved, (top, bottom))`.
///
/// `None` is the answer for every page that is not the one on screen moved: another box, another
/// selection or search, a line whose runs are not the ones that were painted, or a move too far for
/// the two frames to share anything worth keeping. Every line the two frames share has to agree to
/// the row — the same runs, the same height and the same distance from where it was — or the pixels
/// are not the pixels this layout would paint and the whole screenful is painted instead.
fn the_shift_of(
    previous: PaintedFrame<'_>,
    laid_out: &super::layout::LaidOut,
    selection: Option<Selection>,
    find: Option<&FindHighlights>,
    width: u32,
    height: u32,
) -> Option<(i32, i32, i32, (i32, i32))> {
    if previous.width != width || previous.height != height {
        return None;
    }
    if previous.pixels.len() != width as usize * height as usize * 4 {
        return None;
    }
    if previous.selection != selection || previous.find.as_ref() != find {
        return None;
    }

    // A source that names two painted rows is not a source a row can be found by, and a rendered
    // document is the one producer that can make one: the pairs are the document's own, and the
    // frame's rows are the renderer's.
    if has_a_repeated_source(previous.lines) || has_a_repeated_source_in(&laid_out.lines) {
        return None;
    }

    let mut moved = None;
    let mut shared = 0usize;
    for line in &laid_out.lines {
        let Some(old) = previous.lines.iter().find(|old| old.source == line.source) else {
            continue;
        };

        if old.height != line.height || !the_same_runs(&old.runs, &line.runs) {
            return None;
        }

        let step = line.top - old.top;
        if moved.is_some_and(|moved| moved != step) {
            return None;
        }
        moved = Some(step);
        shared += 1;
    }

    let moved = moved?;
    // **The rows above and below the text stand still**, and the text is scrolled between them: the
    // page is read through its margin rather than dragged with it, and a frame short enough to leave
    // room at the bottom leaves the same room after a scroll. So the two frames have to agree about
    // both, and a page whose margin or whose room at the bottom has changed is painted again.
    let content_top = laid_out.lines.first().map(|line| line.top)?;
    let content_bottom = laid_out.lines.last().map(|line| line.top + line.height)?;
    let previous_top = previous.lines.first().map(|line| line.top)?;
    let previous_bottom = previous.lines.last().map(|line| line.top + line.height)?;
    if content_top != previous_top || content_bottom != previous_bottom {
        return None;
    }

    // A frame that has not moved is a frame painted again for something other than a scroll, and that
    // something may be anything: it is painted rather than kept.
    if shared == 0 || moved == 0 || moved.abs() > MOST_ROWS_A_FRAME_IS_SHIFTED {
        return None;
    }
    if moved >= content_bottom - content_top || moved <= content_top - content_bottom {
        return None;
    }

    // Which rows are the ones that came in: the band an edge of the frame gave up, which is also
    // where the page's own margin is on that side.
    let damage = if moved > 0 {
        (0, content_top + moved)
    } else {
        (content_bottom + moved, height as i32)
    };

    Some((moved, content_top, content_bottom, damage))
}

fn has_a_repeated_source(lines: &[FrameLine]) -> bool {
    lines.iter().enumerate().any(|(index, line)| {
        lines[..index]
            .iter()
            .any(|earlier| earlier.source == line.source)
    })
}

fn has_a_repeated_source_in(lines: &[super::layout::LaidLine]) -> bool {
    lines.iter().enumerate().any(|(index, line)| {
        lines[..index]
            .iter()
            .any(|earlier| earlier.source == line.source)
    })
}

/// Whether two painted rows carry the same text in the same places: what makes a row of the frame on
/// screen the row this layout would paint.
fn the_same_runs(old: &[FrameRun], new: &[super::layout::LaidRun]) -> bool {
    old.len() == new.len()
        && old
            .iter()
            .zip(new)
            .all(|(old, new)| old.x == new.x && old.width == new.width && old.text == new.text)
}

/// Move the rows of the frame already on screen into the surface just made, by `moved` rows, leaving
/// the band an edge gave up as whatever the surface was made with.
///
/// The margin above the first line is the one part of the frame a scroll does not move: the page is
/// read through it rather than dragged with it, so those rows are copied as they stand and the text
/// under them is what moves (see `the_shift_of`).
fn copy_the_rows_on_screen(
    surface: &DibSurface,
    previous: &[u8],
    moved: i32,
    content_top: i32,
    content_bottom: i32,
) {
    let width = surface.width as usize;
    let height = surface.height as i32;
    let stride = width * 4;
    let rows = content_bottom - content_top - moved.abs();
    if rows <= 0 {
        return;
    }

    let destination =
        unsafe { std::slice::from_raw_parts_mut(surface.bits(), stride * height as usize) };
    let mut copy = |from: i32, to: i32, rows: i32| {
        for row in 0..rows {
            let source = (from + row) as usize * stride;
            let destination_row = (to + row) as usize * stride;
            destination[destination_row..destination_row + stride]
                .copy_from_slice(&previous[source..source + stride]);
        }
    };

    // Where the rows that survive come from and go to: the frame's text moves towards the edge it is
    // being scrolled towards, so a page scrolled on has its old rows move up under the reader.
    let (source_first, destination_first) = if moved >= 0 {
        (content_top, content_top + moved)
    } else {
        (content_top - moved, content_top)
    };

    // The margin above the text and the room below it, which are the same rows on both frames.
    copy(0, 0, content_top);
    copy(content_bottom, content_bottom, height - content_bottom);
    copy(source_first, destination_first, rows);
}

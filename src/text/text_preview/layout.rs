//! The box a document is drawn in: the lines it fits, where each run lands, the
//! scrollbar that reaches the rest of it, and the painting of all of that into a
//! BGRA frame.
//!
//! A line wider than the box wraps onto the next one — nothing in a preview
//! scrolls sideways — and only the lines the frame shows are ever styled, which is
//! what keeps a preview of a large file cheap.
//!
//! `layout` is asked for the box and `paint` is what turns what it answered into
//! pixels. What comes out is what [`super::frame`] hands back to the caller: a
//! screenful of pixels, the numbers that describe where the preview is, and the
//! painted lines a pointer or a selection is measured against.

use super::document::{styled_window, CachedDocument, DocLine, LineBlock};
use super::frame::{FindHighlights, ScrollBar, Selection};
use crate::text::text_paint::{
    blend, fill_rect, luminance, plain_style, rgb, scaled, DibSurface, RunPainter, TextMetrics,
    TextStyle, BODY_LEVEL, SIZE_LEVELS,
};
use crate::text::text_theme::LoadedTheme;
use windows::Win32::Foundation::RECT;

/// Shortest the thumb gets, so a document of a thousand screens still has
/// something to grab.
const SCROLLBAR_MIN_THUMB_PIXELS: i32 = 24;
/// Narrow files still get a window wide enough to look like one.
const MIN_CONTENT_CHARS: i32 = 16;
/// Room between the line numbers and the text they are numbers for, in logical
/// pixels: the numbers are right-aligned against it, so it is also what keeps a
/// five-digit line from touching the code it numbers.
const GUTTER_GAP_PIXELS: i32 = 8;

/// The column a document's own line numbers are drawn in.
///
/// It is as wide as the widest number the document can show, so that the text
/// column does not move as a page is scrolled from line 9 to line 10.
pub(super) fn gutter_width(lines: usize, metrics: &TextMetrics) -> i32 {
    let digits = lines.max(1).to_string().len() as i32;

    digits * metrics.advance[BODY_LEVEL as usize].max(1) + gutter_gap(metrics)
}

/// The room `GUTTER_GAP_PIXELS` names, at the display the page is drawn on.
fn gutter_gap(metrics: &TextMetrics) -> i32 {
    scaled(GUTTER_GAP_PIXELS, metrics.scale).max(1)
}
pub(super) struct LaidRun {
    pub(super) text: String,
    style: TextStyle,
    pub(super) x: i32,
    pub(super) width: i32,
}

pub(super) struct LaidLine {
    pub(super) runs: Vec<LaidRun>,
    pub(super) top: i32,
    pub(super) height: i32,
    /// Where this line's own text begins: the left edge of the column it is drawn in, past the
    /// numbers and past the indent the line carries. A line with nothing on it has no run to be
    /// measured against, and a caret standing at the start of one is a mark on this edge rather
    /// than on the frame's (see `FrameLine::left`).
    pub(super) left: i32,
    block: LineBlock,
    /// The document line this one drawn line belongs to, and the character of that line its
    /// first run starts at: what a point on the frame is turned back into, and what a selection
    /// is measured in. A document line too long for the box is drawn as several of these, and
    /// which part of the line each of them holds is not something its position says (see
    /// `Selection`).
    pub(super) source: (usize, usize),
    /// Whether this is one of the document's own lines rather than the note a frame adds at its
    /// bottom: only the document's own carry a number, and only their first painted part does (see
    /// `paint`).
    pub(super) own_line: bool,
}

pub(super) struct LaidOut {
    pub(super) lines: Vec<LaidLine>,
    pub(super) width: u32,
    pub(super) height: u32,
    /// Document line the frame starts at, after the position was pulled back so
    /// the frame is full.
    pub(super) first_line: usize,
    /// Document lines the frame shows.
    pub(super) visible_lines: usize,
    /// The column kept clear at the left for the document's own line numbers, or zero for a page
    /// that is not read by them (see `gutter_width`).
    pub(super) gutter: i32,
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
pub(super) fn layout(
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

    // **A pinned page keeps a column for its own line numbers, and a hover does not.** The pin is
    // the page a reader works in — a caret, a selection, a find — and a line is referred to by its
    // number there, whatever the file is: source, prose, a listing, a rendered document. A hover is
    // something looked at on the way past, and a column of numbers down its left is width the text
    // does not have. The column is as wide as the widest number the page can reach, so it does not
    // move as the page is scrolled from line 9 to line 10 (see `gutter_width`).
    let gutter = if full_mode {
        gutter_width(scrollable_lines, metrics)
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
        gutter,
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
            gutter,
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
                gutter,
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
                    x: padding + gutter,
                    width,
                }],
                top: body.y,
                height: note_height,
                left: padding + gutter,
                block: LineBlock::Plain,
                // The note stands where the lines it hides begin, so a press on it is a press on
                // the line under it rather than on nothing at all — and it is not that line, so it
                // carries no number of its own (see `own_line`).
                source: (first_line + body.emitted, 0),
                own_line: false,
            });
            body.y += note_height;
            body.content_right = body.content_right.max(padding + gutter + width);
        }
    }

    LaidOut {
        lines: body.lines,
        width: (body.content_right + padding).clamp(1, box_width.max(1) as i32) as u32,
        height: (body.y + padding).clamp(1, box_height.max(1) as i32) as u32,
        first_line,
        visible_lines: body.emitted,
        gutter,
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
#[allow(clippy::too_many_arguments)] // Each one is a number about the box the page is laid out in.
fn pulled_back_first_line(
    document: &CachedDocument,
    theme: &LoadedTheme,
    metrics: &TextMetrics,
    first_line: usize,
    box_width: u32,
    scrollbar_space: i32,
    gutter: i32,
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

    let left = padding + gutter;
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
#[allow(clippy::too_many_arguments)] // Each one is a number about the box the page is laid out in.
fn lay_out_body(
    document: &CachedDocument,
    theme: &LoadedTheme,
    metrics: &TextMetrics,
    first_line: usize,
    box_size: (u32, u32),
    scrollbar_space: i32,
    gutter: i32,
    reserve: i32,
) -> BodyLayout {
    let doc = &document.doc;
    let (box_width, box_height) = box_size;
    let padding = metrics.padding;
    let box_width = box_width.max(1) as i32;
    let left = padding + gutter;
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
            for (source_char, runs) in visual_lines {
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
                    left: text_left(line, metrics, left),
                    block: line.block,
                    source: (index, source_char),
                    own_line: true,
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
/// lines it takes, left to right, wrapped at the page edge, each with the
/// character of the document line it starts at.
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
) -> Vec<(usize, Vec<LaidRun>)> {
    let text_right = match line.block {
        LineBlock::Quote => right - metrics.quote_bar - 4,
        _ => right,
    };

    wrap_line(
        line,
        text_left(line, metrics, left),
        text_right,
        max_lines,
        metrics,
    )
}

/// Where a document line's own text begins inside the column kept for it: the indent the line
/// carries shifts the whole of it, so an indented line starts further right than the column does.
///
/// Both the runs of a line and the edge a caret on an empty line stands on are worked out here, so
/// that a line with nothing to measure still begins where the same line with something on it would
/// (see `LaidLine::left`).
fn text_left(line: &DocLine, metrics: &TextMetrics, left: i32) -> i32 {
    left + metrics.indent(line.indent)
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
/// `max_lines` of them, each with the character of the document line it starts at.
/// A word wider than the page — a URL, a hash — is cut rather than allowed to run
/// off it, and a space that lands at a wrap point is dropped instead of starting a
/// line.
///
/// What the offsets are for is a selection: a copy is taken from the document
/// rather than from what was drawn, so a wrap point has to say where the next
/// visual line picks the document line up — a dropped space is still a character of
/// it — and not how wide the frame is (see `Selection`).
fn wrap_line(
    line: &DocLine,
    left: i32,
    limit: i32,
    max_lines: usize,
    metrics: &TextMetrics,
) -> Vec<(usize, Vec<LaidRun>)> {
    let mut lines: Vec<(usize, Vec<LaidRun>)> = Vec::new();
    let mut runs: Vec<LaidRun> = Vec::new();
    let mut x = left;
    // Characters of the document line laid out so far, and where the visual line
    // being built starts among them.
    let mut emitted = 0usize;
    let mut line_from = 0usize;

    for_each_token(line, |token| {
        let count = token.text.chars().count();
        let advance = metrics.advance[(token.style.level as usize).min(SIZE_LEVELS - 1)].max(1);
        let width = count as i32 * advance;

        if token.space {
            // A space at a wrap point is dropped rather than starting a line, and
            // the character it is stays part of the line it was dropped from.
            emitted += count;
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
            lines.push((line_from, std::mem::take(&mut runs)));
            line_from = emitted;
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
            emitted += count;
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
            emitted += take;

            if offset < characters.len() {
                if lines.len() + 1 == max_lines {
                    return false;
                }
                lines.push((line_from, std::mem::take(&mut runs)));
                line_from = emitted;
                x = left;
            }
        }

        true
    });

    if !runs.is_empty() {
        lines.push((line_from, runs));
    }
    if lines.is_empty() {
        // An empty line still takes a line's worth of space.
        lines.push((0, Vec::new()));
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
pub(super) fn scrollbar_geometry(
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

/// Which of the marks over a run a piece of it belongs to, in the order they are drawn over one
/// another where two of them meet: a place a search found, the one the reader is on, and the
/// selection they made for themselves.
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum RunMark {
    Find,
    Current,
    Selection,
}

/// The ink a found run's glyphs are drawn in over `background`: the theme's own find foreground
/// wins; otherwise whichever of the run's own ink and the page stands further from the match
/// background, so a strong band brightens its text rather than swallowing it.
fn find_ink(ink: [u8; 3], page: [u8; 3], background: [u8; 3], chosen: Option<[u8; 3]>) -> [u8; 3] {
    if let Some(chosen) = chosen {
        return chosen;
    }
    let ink_from = (luminance(ink) - luminance(background)).abs();
    let page_from = (luminance(page) - luminance(background)).abs();
    if page_from > ink_from {
        page
    } else {
        ink
    }
}

/// The part of a run a selection covers, as character offsets within the run, or
/// `None` when it does not touch it.
///
/// A line is painted run by run — one per token, which is where the colors change —
/// so a selection, which is measured in the document's characters, has to be mapped
/// back onto the run it lands in. `source` is the document place the painted line
/// begins at and `run_start` is how many characters of that painted line come
/// before this run, so the run's own characters are the document's
/// `[source.1 + run_start, ...)` — which is also what tells a wrapped line's second
/// half from its first, and what keeps a selection measured on one of them from
/// covering the other (see `Selection`).
fn selected_part(
    source: (usize, usize),
    run_start: usize,
    run_characters: usize,
    selection: Selection,
) -> Option<(usize, usize)> {
    let ((start_line, start_char), (end_line, end_char)) = selection.ordered();
    let (line, base) = source;
    if line < start_line || line > end_line || run_characters == 0 {
        return None;
    }

    let run_from = base + run_start;
    let run_to = run_from + run_characters;
    let from = if line == start_line {
        start_char.max(run_from)
    } else {
        run_from
    };
    let to = if line == end_line {
        end_char.min(run_to)
    } else {
        run_to
    };

    (from < to).then(|| (from - run_from, to - run_from))
}

/// Paint the background, the block decorations, the text and a selection, and the
/// scrollbar when the document is longer than the frame.
///
/// Every run is drawn with an opaque background rectangle, so the page color
/// fills the box and the pieces GDI leaves alone (the space a line that stops
/// before the edge did not use, the gaps between blocks) are already the right
/// color.
#[allow(clippy::too_many_arguments)] // The box, the options and the rows a moved page leaves to paint are one frame.
pub(super) unsafe fn paint(
    surface: &DibSurface,
    laid_out: &LaidOut,
    theme: &LoadedTheme,
    metrics: &TextMetrics,
    scrollbar: Option<&ScrollBar>,
    selection: Option<Selection>,
    find: Option<&FindHighlights>,
    damage: Option<(i32, i32)>,
) {
    let width = surface.width as i32;
    let page = rgb(theme.background());
    let band_color = blend(page, rgb(theme.foreground()), 0.08);
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

    // A find's places are drawn in the theme's find highlight where it has one, and in a shade of
    // the page where it does not. The derived shade is mixed well away from the page so every mark
    // is stronger than the selection's wash; a colour the theme names is used as it stands. Every
    // other place the query was found is a wash of that colour; the one the reader is on is the
    // colour itself, which is what tells the count's current match from the rest of the page (see
    // `FindHighlights`).
    let find_color = theme
        .theme()
        .settings
        .find_highlight
        .map(rgb)
        .unwrap_or_else(|| blend(page, rgb(theme.foreground()), 0.65));
    let find_background = blend(page, find_color, 0.75);
    let find_foreground = theme.theme().settings.find_highlight_foreground.map(rgb);

    // A paint of a band of the frame is a paint of a page moved under the reader: the rows outside it
    // are the ones already on screen, shifted (see `PaintedFrame`). The band clamps every fill and
    // clips the text, so nothing this paints can reach them.
    let band = damage.map(|(top, bottom)| RECT {
        left: 0,
        top,
        right: width,
        bottom,
    });
    let fill = |rect: RECT, color: [u8; 3]| {
        let rect = match band {
            Some(band) => RECT {
                top: rect.top.max(band.top),
                bottom: rect.bottom.min(band.bottom),
                ..rect
            },
            None => rect,
        };
        fill_rect(surface, rect, color);
    };

    fill(
        RECT {
            left: 0,
            top: 0,
            right: width,
            bottom: surface.height as i32,
        },
        page,
    );

    // The rows that are not in the band are already on screen, so their runs are not asked of GDI at
    // all (see `PaintedFrame`).
    //
    // **They are skipped rather than clipped**, and that is not the same thing: GDI draws text
    // inside a clip region with its antialiasing turned down, so a band clipped out of a page is a
    // band whose glyphs are a little lighter than the same glyphs painted with no clip — a page that
    // shimmers where the band meets the rows the shift kept. A run's own rectangle is the room it was
    // given, which is the row it is in, so nothing of a run in the band reaches outside it.
    let in_the_band = |line: &LaidLine| match band {
        Some(band) => line.top + line.height > band.top && line.top < band.bottom,
        None => true,
    };
    for line in laid_out.lines.iter().filter(|line| in_the_band(line)) {
        match line.block {
            LineBlock::Code => fill(
                RECT {
                    left: 0,
                    top: line.top,
                    right: width,
                    bottom: line.top + line.height,
                },
                band_color,
            ),
            LineBlock::Quote => {
                let bar = theme.style_for_scopes(&["markup.quote"]).foreground;
                fill(
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

    // The numbers of the document's own lines, drawn in the column kept clear for them: a page that
    // keeps no column draws none, a painted line that is the wrapped tail of its line draws none,
    // and the note a frame adds at its bottom is not one of the document's lines and draws none
    // either (see `gutter_width`, `own_line` and `source`).
    if laid_out.gutter > 0 {
        let advance = metrics.advance[BODY_LEVEL as usize].max(1);
        let right = metrics.padding + laid_out.gutter - gutter_gap(metrics);
        let mut style = plain_style(BODY_LEVEL);
        style.foreground = blend(rgb(theme.foreground()), rgb(theme.background()), 0.45);

        for line in laid_out.lines.iter().filter(|line| in_the_band(line)) {
            if !line.own_line || line.source.1 != 0 {
                continue;
            }

            let number = (line.source.0 + 1).to_string();
            let width = number.chars().count() as i32 * advance;
            let x = (right - width).max(metrics.padding);
            painter.draw(
                &number,
                x,
                RECT {
                    left: x,
                    top: line.top,
                    right: x + width,
                    bottom: line.top + line.height,
                },
                &style,
                style.foreground,
                page,
            );
        }
    }

    for line in laid_out.lines.iter().filter(|line| in_the_band(line)) {
        let mut run_start = 0usize;

        for run in &line.runs {
            let run_characters = run.text.chars().count();

            // The marks over this run, in the characters the run itself counts: the reader's own
            // selection, and every place a find found, with the current one told from the rest. The
            // search's places do not overlap each other — `find_all` walks past each one it takes —
            // so the only overlap a run can carry is a find under the selection, and the selection
            // wins it (see `RunMark`).
            let mut marks: Vec<(usize, usize, RunMark)> = Vec::new();
            if let Some(part) = selection.and_then(|selection| {
                selected_part(line.source, run_start, run_characters, selection)
            }) {
                marks.push((part.0, part.1, RunMark::Selection));
            }
            if let Some(find) = find {
                for (index, at) in find.matches.iter().enumerate() {
                    if let Some((from, to)) =
                        selected_part(line.source, run_start, run_characters, *at)
                    {
                        let mark = if find.current == Some(index) {
                            RunMark::Current
                        } else {
                            RunMark::Find
                        };
                        marks.push((from, to, mark));
                    }
                }
            }

            run_start += run_characters;

            if run.text.is_empty() {
                continue;
            }

            let run_background = match (run.style.background, line.block) {
                (Some(color), _) => color,
                (None, LineBlock::Code) => band_color,
                (None, _) => page,
            };

            let rect = RECT {
                left: run.x,
                top: line.top,
                right: run.x + run.width,
                bottom: line.top + line.height,
            };

            if marks.is_empty() {
                painter.draw(
                    &run.text,
                    rect.left,
                    rect,
                    &run.style,
                    run.style.foreground,
                    run_background,
                );
                continue;
            }

            // The run is drawn in the pieces the marks cut it into, each with the colour of the
            // highest mark that covers it: the text before a mark as usual, the mark without
            // touching its background, and the text after it as usual.
            let characters: Vec<char> = run.text.chars().collect();
            let advance = run.width as f32 / characters.len().max(1) as f32;
            let mut edges = vec![0usize, run_characters];
            for (from, to, _) in &marks {
                edges.push(*from);
                edges.push(*to);
            }
            edges.sort_unstable();
            edges.dedup();

            for pair in edges.windows(2) {
                let (start, end) = (pair[0], pair[1]);
                if start >= end {
                    continue;
                }

                let mark = marks
                    .iter()
                    .filter(|(from, to, _)| *from <= start && *to >= end)
                    .map(|(_, _, mark)| *mark)
                    .max();
                let (foreground, background) = match mark {
                    Some(RunMark::Selection) => (
                        highlight_foreground.unwrap_or(run.style.foreground),
                        highlight,
                    ),
                    Some(RunMark::Current) => (
                        find_ink(run.style.foreground, page, find_color, find_foreground),
                        find_color,
                    ),
                    Some(RunMark::Find) => (
                        find_ink(run.style.foreground, page, find_background, find_foreground),
                        find_background,
                    ),
                    None => (run.style.foreground, run_background),
                };

                let piece: String = characters[start..end].iter().collect();
                let left = run.x + (start as f32 * advance).round() as i32;
                let right = run.x + (end as f32 * advance).round() as i32;
                painter.draw(
                    &piece,
                    left,
                    RECT {
                        left,
                        right,
                        ..rect
                    },
                    &run.style,
                    foreground,
                    background,
                );
            }
        }
    }

    // The fonts go with the painter, and it puts the surface's own object back
    // before they are deleted.
    drop(painter);

    // The scrollbar goes last so it sits above everything, including a line that
    // ran to the right edge. It is filled whole and not through `fill`: its thumb is the one thing
    // about a page that moves *without* the text moving, so it is not in the band a shifted frame
    // leaves to paint (see `the_shift_of`).
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

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
use super::frame::{ScrollBar, Selection};
use crate::text::text_paint::{
    blend, fill_rect, plain_style, rgb, scaled, DibSurface, RunPainter, TextMetrics, TextStyle,
    BODY_LEVEL, SIZE_LEVELS,
};
use crate::text::text_theme::LoadedTheme;
use windows::Win32::Foundation::RECT;

/// Shortest the thumb gets, so a document of a thousand screens still has
/// something to grab.
const SCROLLBAR_MIN_THUMB_PIXELS: i32 = 24;
/// Narrow files still get a window wide enough to look like one.
const MIN_CONTENT_CHARS: i32 = 16;
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
    block: LineBlock,
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
pub(super) unsafe fn paint(
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

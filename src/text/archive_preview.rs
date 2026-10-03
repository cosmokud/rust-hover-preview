//! Archive previews: what an archive holds, drawn as a tree.
//!
//! The page is a summary header over a capped tree of the archive's own contents
//! — folders first, each row indented by its depth, every file's size
//! right-aligned and every name cut rather than wrapped. It is painted with the
//! same GDI layer a text preview uses and colored by the same theme, which is
//! what keeps the two previews looking like one app; what it never does is ask
//! the highlighter anything, because there is nothing to highlight: the rows come
//! from `archive_listing`, which reads the archive's table of contents and no
//! member of it.
//!
//! Two calls share the work, as they do for text: `measure` answers how big the
//! page wants to be before the window is placed, and `render` fills the box the
//! layout settled on. Both build the same page and neither keeps it — the listing
//! they are built from is what is cached.

use crate::config::config::TextTheme;
use crate::readers::archive_listing::{self, ArchiveEntry, Listing};
use crate::text::text_paint::{
    blend, fill_rect, plain_style, rgb, scaled, DibSurface, RunPainter, TextMetrics, TextStyle,
    BODY_LEVEL,
};
use crate::text::text_theme::{self, LoadedTheme};
use once_cell::sync::Lazy;
use std::path::Path;
use windows::Win32::Foundation::{LPARAM, RECT};
use windows::Win32::Graphics::Gdi::{
    CreateCompatibleDC, DeleteDC, EnumFontFamiliesExW, HDC, LOGFONTW, TEXTMETRICW,
};

/// Rows the tree draws before it stops. Past this the page says how many are
/// left, because a preview is something to glance at, not a file manager.
const MAX_ROWS: usize = 100;

/// Spaces one level of depth is indented by. Three reads as a tree without
/// pushing a deep path off the right edge.
const INDENT_SPACES: i32 = 3;

/// The cell an entry's icon is drawn in, in advances. Fixed, so every name starts
/// at the same column whatever the glyph's own advance is.
const ICON_CELL_ADVANCES: i32 = 2;

/// Room kept after that cell, so a glyph that fills it is not drawn against the
/// name beside it.
const ICON_TEXT_GAP_ADVANCES: i32 = 1;

/// Room between the widest name and the size column.
const SIZE_GAP_ADVANCES: i32 = 3;

/// The size level the summary line is set in. Icons are drawn at the level of
/// the text they stand beside, so the two share a baseline and line up without
/// any nudging.
const HEADER_LEVEL: u8 = 3;

/// The hairline under the header, and the room kept around it.
const RULE_PIXELS: i32 = 1;
const RULE_GAP_PIXELS: i32 = 5;

/// The narrowest a page is worth drawing, in body advances, so an archive of two
/// short names still gets a window that looks like one.
const MIN_CONTENT_ADVANCES: i32 = 16;

/// The icon faces Windows ships, newest first: Windows 11 has Segoe Fluent Icons
/// and Windows 10 the MDL2 set it replaces, and every glyph this preview asks for
/// exists in both.
const ICON_FACES: [&str; 2] = ["Segoe Fluent Icons", "Segoe MDL2 Assets"];

/// The glyphs, by the name the icon list gives them. Neither face is guaranteed
/// to carry a glyph for a given code point, so each is checked by eye once.
const ICON_ARCHIVE: char = '\u{E7B8}';
const ICON_FOLDER: char = '\u{E8B7}';
const ICON_FILE: char = '\u{E8A5}';
const ICON_IMAGE: char = '\u{E91B}';
const ICON_VIDEO: char = '\u{E714}';
const ICON_AUDIO: char = '\u{E8D6}';
const ICON_CODE: char = '\u{E943}';
const ICON_PDF: char = '\u{E9F9}';

/// The options a preview is built with, passed in rather than read from the
/// configuration here so a caller's intent cannot drift from what is drawn.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct ArchivePreviewOptions {
    pub(crate) theme: TextTheme,
    pub(crate) font_scale_percent: u32,
}

/// The box the page wants, bounded by what the display can give it.
pub(crate) fn measure(
    path: &Path,
    max_width: u32,
    max_height: u32,
    dpi: u32,
    options: ArchivePreviewOptions,
) -> Option<(u32, u32)> {
    let listing = archive_listing::listing_for(path, None)?;
    let theme = text_theme::loaded(options.theme)?;
    let name = file_name(path);

    let dc = unsafe { CreateCompatibleDC(None) };
    if dc.0.is_null() {
        return None;
    }
    let result = TextMetrics::new(dc, dpi, options.font_scale_percent)
        .map(|metrics| build_page(&listing, &name, theme, &metrics, max_width, max_height))
        .map(|page| (page.width, page.height));
    unsafe {
        let _ = DeleteDC(dc);
    }

    result
}

/// The page painted into the box the layout settled on.
pub(crate) fn render(
    path: &Path,
    width: u32,
    height: u32,
    dpi: u32,
    options: ArchivePreviewOptions,
) -> Option<(Vec<u8>, u32, u32)> {
    let listing = archive_listing::listing_for(path, None)?;
    let theme = text_theme::loaded(options.theme)?;
    let name = file_name(path);

    let surface = DibSurface::create(width, height)?;
    let metrics = TextMetrics::new(surface.dc, dpi, options.font_scale_percent)?;
    let page = build_page(&listing, &name, theme, &metrics, width, height);
    paint(&surface, &page, theme, metrics.scale);

    Some((surface.pixels(), surface.width, surface.height))
}

fn file_name(path: &Path) -> String {
    path.file_name()
        .map(|name| name.to_string_lossy().to_string())
        .unwrap_or_default()
}

// ----------------------------------------------------------------------- tree

/// One entry, or one directory the archive only implies.
struct Node {
    name: String,
    is_dir: bool,
    size: u64,
    encrypted: bool,
    children: Vec<usize>,
}

/// The entries arranged as a tree, in the arena the walk descends through.
struct Tree {
    nodes: Vec<Node>,
}

impl Tree {
    /// Build the tree from the entries' own paths.
    ///
    /// An entry is a path, not a placement: `docs/report.pdf` states a file and
    /// its folder, and a zip written without `docs/` as an entry of its own still
    /// holds one. A directory is therefore made where it is missing, which is
    /// also why the header counts folders the way a file manager would rather
    /// than the way the archive happens to be stored.
    fn build(entries: &[ArchiveEntry]) -> Self {
        use std::collections::HashMap;

        let mut tree = Self {
            nodes: vec![Node {
                name: String::new(),
                is_dir: true,
                size: 0,
                encrypted: false,
                children: Vec::new(),
            }],
        };
        let mut index: HashMap<(usize, String), usize> = HashMap::new();

        for entry in entries {
            let segments: Vec<&str> = entry
                .name
                .split('/')
                .filter(|segment| !segment.is_empty())
                .collect();
            if segments.is_empty() {
                continue;
            }

            let mut parent = 0usize;
            for (position, segment) in segments.iter().enumerate() {
                let last = position + 1 == segments.len();
                let node = match index.get(&(parent, (*segment).to_string())).copied() {
                    Some(node) => node,
                    None => {
                        tree.nodes.push(Node {
                            name: (*segment).to_string(),
                            is_dir: true,
                            size: 0,
                            encrypted: false,
                            children: Vec::new(),
                        });
                        let node = tree.nodes.len() - 1;
                        tree.nodes[parent].children.push(node);
                        index.insert((parent, (*segment).to_string()), node);
                        node
                    }
                };

                if last && !entry.is_dir {
                    // A name that is a file here and a folder elsewhere is the
                    // archive listing the same place twice, and the folder is
                    // what holds the rest of it.
                    if tree.nodes[node].is_dir && !tree.nodes[node].children.is_empty() {
                        break;
                    }
                    let node = &mut tree.nodes[node];
                    node.is_dir = false;
                    node.size = entry.size;
                    node.encrypted = entry.encrypted;
                } else if !last {
                    // A segment with something under it is a folder, whatever a
                    // member of the same name said first.
                    tree.nodes[node].is_dir = true;
                }

                parent = node;
            }
        }

        tree.sort_children(0);
        tree
    }

    /// Folders first, then files, each by name — case-insensitively, and with
    /// digit runs compared as numbers, so `report2` sits before `report10`.
    fn sort_children(&mut self, index: usize) {
        let children = self.nodes[index].children.clone();
        for child in &children {
            self.sort_children(*child);
        }

        let nodes = &self.nodes;
        let mut children = children;
        children.sort_by(|left, right| {
            let left = &nodes[*left];
            let right = &nodes[*right];
            right
                .is_dir
                .cmp(&left.is_dir)
                .then_with(|| natural_order(&left.name, &right.name))
        });
        self.nodes[index].children = children;
    }

    /// How many files and folders the tree holds, the root not counted.
    fn counts(&self) -> (usize, usize) {
        let mut files = 0usize;
        let mut folders = 0usize;
        for node in self.nodes.iter().skip(1) {
            if node.is_dir {
                folders += 1;
            } else {
                files += 1;
            }
        }

        (files, folders)
    }

    /// The rows to draw, in the order they are drawn: the walk descends into a
    /// folder before moving on, so a folder is followed by what it holds.
    fn rows(&self) -> Vec<TreeRow> {
        let mut rows = Vec::new();
        self.walk(0, 0, &mut rows);

        rows
    }

    fn walk(&self, index: usize, depth: usize, rows: &mut Vec<TreeRow>) {
        for child in self.nodes[index].children.clone() {
            if rows.len() >= MAX_ROWS {
                return;
            }

            if self.nodes[child].is_dir {
                let end = self.push_chain(child, depth, rows);
                self.walk(end, depth + 1, rows);
            } else {
                let node = &self.nodes[child];
                rows.push(TreeRow {
                    depth,
                    name: node.name.clone(),
                    is_dir: false,
                    size: node.size,
                    encrypted: node.encrypted,
                    covers: 1,
                });
            }
        }
    }

    /// Draw one folder, and the folders under it that hold nothing but the next
    /// one, as a single row: `a/b/c/` saves the two rows that would have held a
    /// single name each. Returns the deepest folder the row stands for, whose
    /// children are walked next.
    fn push_chain(&self, start: usize, depth: usize, rows: &mut Vec<TreeRow>) -> usize {
        let mut index = start;
        let mut name = self.nodes[index].name.clone();
        let mut covers = 1usize;

        loop {
            let node = &self.nodes[index];
            let [only] = node.children[..] else {
                break;
            };
            if !self.nodes[only].is_dir {
                break;
            }

            name.push('/');
            name.push_str(&self.nodes[only].name);
            index = only;
            covers += 1;
        }

        rows.push(TreeRow {
            depth,
            name,
            is_dir: true,
            size: 0,
            encrypted: false,
            covers,
        });

        index
    }
}

/// One drawn row of the tree.
struct TreeRow {
    depth: usize,
    name: String,
    is_dir: bool,
    size: u64,
    /// Whether the entry behind this row has to be unlocked before it can be read. A
    /// folder row is never one: a container's own headers are encrypted or they are not,
    /// which is the listing's answer rather than a row's.
    encrypted: bool,
    /// Tree nodes this row stands for, which is more than one where a chain of
    /// single-child folders was collapsed.
    covers: usize,
}

/// What a row says about an entry that needs a password to be read. It is written beside the
/// name rather than in a column of its own: it is rare, and the name is what it applies to.
const ENCRYPTED_MARK: &str = " (encrypted)";

/// Compare two names the way a person reads a list: without case, and with the
/// numbers in them counting as numbers.
fn natural_order(left: &str, right: &str) -> std::cmp::Ordering {
    use std::cmp::Ordering;

    let left: Vec<char> = left.to_lowercase().chars().collect();
    let right: Vec<char> = right.to_lowercase().chars().collect();
    let (mut l, mut r) = (0usize, 0usize);

    while l < left.len() && r < right.len() {
        if left[l].is_ascii_digit() && right[r].is_ascii_digit() {
            let mut l_end = l;
            while l_end < left.len() && left[l_end].is_ascii_digit() {
                l_end += 1;
            }
            let mut r_end = r;
            while r_end < right.len() && right[r_end].is_ascii_digit() {
                r_end += 1;
            }

            let l_number: String = left[l..l_end].iter().collect();
            let r_number: String = right[r..r_end].iter().collect();
            let l_value = l_number.trim_start_matches('0');
            let r_value = r_number.trim_start_matches('0');
            let ordering = l_value
                .len()
                .cmp(&r_value.len())
                .then_with(|| l_value.cmp(r_value))
                .then_with(|| l_number.len().cmp(&r_number.len()));
            if ordering != Ordering::Equal {
                return ordering;
            }

            l = l_end;
            r = r_end;
            continue;
        }

        let ordering = left[l].cmp(&right[r]);
        if ordering != Ordering::Equal {
            return ordering;
        }
        l += 1;
        r += 1;
    }

    (left.len() - l).cmp(&(right.len() - r))
}

// ----------------------------------------------------------------------- page

/// One run of the page: a piece of text drawn at `x`, in one style.
struct PageRun {
    text: String,
    x: i32,
    width: i32,
    style: TextStyle,
}

/// One line of the page, or the hairline under the header.
struct PageLine {
    runs: Vec<PageRun>,
    top: i32,
    height: i32,
    rule: bool,
}

/// A painted page: what to draw and how big it came out.
struct Page {
    lines: Vec<PageLine>,
    width: u32,
    height: u32,
    padding: i32,
}

fn build_page(
    listing: &Listing,
    name: &str,
    theme: &LoadedTheme,
    metrics: &TextMetrics,
    box_width: u32,
    box_height: u32,
) -> Page {
    let body_advance = metrics.advance[BODY_LEVEL as usize].max(1);
    let body_height = metrics.line_height[BODY_LEVEL as usize];
    let header_height = metrics.line_height[HEADER_LEVEL as usize];
    let padding = metrics.padding;
    let gap = body_advance * SIZE_GAP_ADVANCES;
    let icon_cell = body_advance * (ICON_CELL_ADVANCES + ICON_TEXT_GAP_ADVANCES);
    let indent = body_advance * INDENT_SPACES;

    let page_color = rgb(theme.background());
    let foreground = rgb(theme.foreground());
    let muted = blend(foreground, page_color, 0.45);

    let tree = Tree::build(&listing.entries);
    let rows = tree.rows();
    let (files, folders) = tree.counts();

    // The size column is reserved from the widest size in the listing, so what
    // gets cut when the box runs out of room is a name and never a size.
    let size_column = rows
        .iter()
        .filter(|row| !row.is_dir)
        .map(|row| text_width(&format_size(row.size), body_advance))
        .max()
        .unwrap_or(body_advance * 3)
        .max(text_width("size", body_advance));

    // The width the page would like: the widest row, or the header, whichever
    // asks for more. It is clamped to the box the way every preview's size is.
    let header = header_runs(listing, name, files, folders, metrics, foreground, muted);
    let header_width: i32 = header.iter().map(|run| run.width).sum();
    let widest_row = rows
        .iter()
        .map(|row| {
            indent * row.depth as i32
                + icon_cell
                + text_width(&row.name, body_advance)
                + gap
                + if row.is_dir { 0 } else { size_column }
        })
        .max()
        .unwrap_or(0);

    let content_width = header_width
        .max(widest_row)
        .max(body_advance * MIN_CONTENT_ADVANCES);
    let width = (content_width + padding * 2).clamp(1, box_width.max(1) as i32) as u32;

    // How many rows fit under the header and the hairline, with a line kept back
    // for the count of what is not shown when something is not shown.
    let rule_height =
        scaled(RULE_PIXELS, metrics.scale) + scaled(RULE_GAP_PIXELS, metrics.scale) * 2;
    let body_top = padding + header_height + rule_height;
    let room = box_height as i32 - body_top - padding;
    if room < body_height {
        return Page {
            lines: Vec::new(),
            width,
            height: 1,
            padding,
        };
    }

    let fits = (room / body_height) as usize;
    let listing_capped = listing.scan_capped || listing.read_truncated || listing.encrypted_headers;
    let note_needed = rows.len() > fits || listing_capped || listing.entries.is_empty();
    let capacity = if note_needed {
        ((room - body_height) / body_height).max(0) as usize
    } else {
        fits
    };

    let shown = rows.len().min(capacity);
    let total: usize = rows.iter().map(|row| row.covers).sum();
    let covered: usize = rows.iter().take(shown).map(|row| row.covers).sum();
    let hidden = total.saturating_sub(covered) + rows.len().saturating_sub(shown);

    let content_right = width as i32 - padding;
    let name_room = content_right - size_column - gap;

    let mut lines = Vec::new();
    let mut top = padding;

    lines.push(PageLine {
        runs: header,
        top,
        height: header_height,
        rule: false,
    });
    top += header_height + scaled(RULE_GAP_PIXELS, metrics.scale);
    lines.push(PageLine {
        runs: Vec::new(),
        top,
        height: scaled(RULE_PIXELS, metrics.scale),
        rule: true,
    });
    top += scaled(RULE_PIXELS, metrics.scale) + scaled(RULE_GAP_PIXELS, metrics.scale);

    let face = icon_face();

    for row in rows.iter().take(shown) {
        let left = padding + indent * row.depth as i32;
        let mut runs = Vec::new();

        if let Some(face) = face {
            let mut icon_style = plain_style(BODY_LEVEL);
            icon_style.face = face;
            icon_style.foreground = muted;
            runs.push(PageRun {
                text: icon_glyph(row.is_dir, &row.name).to_string(),
                x: left,
                width: icon_cell,
                style: icon_style,
            });
        }

        let name_left = left + if face.is_some() { icon_cell } else { 0 };
        let mut name_style = plain_style(BODY_LEVEL);
        name_style.foreground = foreground;
        name_style.bold = row.is_dir;

        // An entry that cannot be read without a password says so beside its name, and the
        // name gives way to the mark rather than the mark being cut off with it: what the
        // mark says is not in the name, so a long name must not be what hides it.
        let mark = if row.encrypted { ENCRYPTED_MARK } else { "" };
        let room = name_room - name_left - text_width(mark, body_advance);
        let name = cut_to_width(&row.name, room, body_advance);
        let name_width = text_width(&name, body_advance);
        runs.push(PageRun {
            text: name,
            x: name_left,
            width: name_width,
            style: name_style,
        });

        if !mark.is_empty() {
            let mut mark_style = plain_style(BODY_LEVEL);
            mark_style.foreground = muted;
            runs.push(PageRun {
                text: mark.to_string(),
                x: name_left + name_width,
                width: text_width(mark, body_advance),
                style: mark_style,
            });
        }

        if !row.is_dir {
            let mut size_style = plain_style(BODY_LEVEL);
            size_style.foreground = muted;
            let size = format_size(row.size);
            runs.push(PageRun {
                text: size.clone(),
                x: content_right - size_column,
                width: size_column.max(text_width(&size, body_advance)),
                style: size_style,
            });
        }

        lines.push(PageLine {
            runs,
            top,
            height: body_height,
            rule: false,
        });
        top += body_height;
    }

    if note_needed {
        let mut style = plain_style(BODY_LEVEL);
        style.foreground = muted;
        let note = note_text(listing, hidden);
        lines.push(PageLine {
            runs: vec![PageRun {
                text: note.clone(),
                x: padding,
                width: text_width(&note, body_advance),
                style,
            }],
            top,
            height: body_height,
            rule: false,
        });
        top += body_height;
    }

    Page {
        lines,
        width,
        height: box_height.min((top + padding) as u32).max(1),
        padding,
    }
}

/// The summary line: what the file is called, what is in it, how much that takes,
/// and anything about the archive the reader could tell.
fn header_runs(
    listing: &Listing,
    name: &str,
    files: usize,
    folders: usize,
    metrics: &TextMetrics,
    foreground: [u8; 3],
    muted: [u8; 3],
) -> Vec<PageRun> {
    let advance = metrics.advance[HEADER_LEVEL as usize].max(1);
    let body_advance = metrics.advance[BODY_LEVEL as usize].max(1);
    let cell = body_advance * (ICON_CELL_ADVANCES + ICON_TEXT_GAP_ADVANCES);

    let mut runs = Vec::new();
    let mut x = metrics.padding;

    let mut icon_style = plain_style(HEADER_LEVEL);
    icon_style.foreground = muted;
    if let Some(face) = icon_face() {
        icon_style.face = face;
    }
    runs.push(PageRun {
        text: if icon_style.face.is_empty() {
            String::new()
        } else {
            ICON_ARCHIVE.to_string()
        },
        x,
        width: cell,
        style: icon_style,
    });
    x += cell;

    let mut push = |runs: &mut Vec<PageRun>, text: &str, bold: bool, color: [u8; 3]| {
        let mut style = plain_style(HEADER_LEVEL);
        style.bold = bold;
        style.foreground = color;
        runs.push(PageRun {
            text: text.to_string(),
            x,
            width: text_width(text, advance),
            style,
        });
        x += text_width(text, advance);
    };

    push(&mut runs, name, true, foreground);
    push(&mut runs, " · ", false, muted);
    push(
        &mut runs,
        &count_label(files, "file", "files"),
        false,
        foreground,
    );
    push(&mut runs, " · ", false, muted);
    push(
        &mut runs,
        &count_label(folders, "folder", "folders"),
        false,
        foreground,
    );
    push(&mut runs, " · ", false, muted);
    push(
        &mut runs,
        &format_size(listing.total_size),
        false,
        foreground,
    );

    if let Some(packed) = listing.packed_total {
        let total = listing.total_size;
        if total > 0 && packed < total {
            let saved = 100 - (packed.saturating_mul(100) / total) as u32;
            push(&mut runs, " · ", false, muted);
            push(
                &mut runs,
                &format!("{} packed ({saved}% smaller)", format_size(packed)),
                false,
                muted,
            );
        }
    } else if listing.file_size > 0 && listing.file_size < listing.total_size {
        // A format that cannot say what one member takes can still say what the
        // whole archive takes, which is the part of the comparison it knows.
        push(&mut runs, " · ", false, muted);
        push(
            &mut runs,
            &format!("{} on disk", format_size(listing.file_size)),
            false,
            muted,
        );
    }

    if listing.encrypted_headers {
        push(&mut runs, " · ", false, muted);
        push(&mut runs, "encrypted", false, muted);
    } else if listing.entries.iter().any(|entry| entry.encrypted) {
        push(&mut runs, " · ", false, muted);
        push(&mut runs, "some entries encrypted", false, muted);
    }

    if listing.multi_volume {
        push(&mut runs, " · ", false, muted);
        push(&mut runs, "first volume of a set", false, muted);
    }

    if listing.scan_capped || listing.read_truncated {
        push(&mut runs, " · ", false, muted);
        push(&mut runs, "listing stopped early", false, muted);
    }

    runs
}

/// A count with the right word after it, because a header reads `1 folder` and
/// `2 folders` and a preview is a page of prose as much as a list.
fn count_label(count: usize, singular: &str, plural: &str) -> String {
    if count == 1 {
        format!("1 {singular}")
    } else {
        format!("{count} {plural}")
    }
}

/// The note under a tree that does not show everything.
fn note_text(listing: &Listing, hidden: usize) -> String {
    if listing.encrypted_headers {
        return "… the contents are encrypted (a password is needed to list them)".to_string();
    }
    if listing.entries.is_empty() {
        return "… this archive is empty".to_string();
    }
    if listing.scan_capped || listing.read_truncated {
        // What was read holds what is left; what was never read holds more, so
        // the count is a floor and says so.
        return match hidden {
            0 => "… and more items (the listing stopped early)".to_string(),
            1 => "… and at least 1 more item (the listing stopped early)".to_string(),
            hidden => format!("… and at least {hidden} more items (the listing stopped early)"),
        };
    }

    match hidden {
        0 => String::new(),
        1 => "… and 1 more item".to_string(),
        hidden => format!("… and {hidden} more items"),
    }
}

// -------------------------------------------------------------------- drawing

fn paint(surface: &DibSurface, page: &Page, theme: &LoadedTheme, scale: f32) {
    let page_color = rgb(theme.background());
    let rule_color = blend(page_color, rgb(theme.foreground()), 0.18);

    fill_rect(
        surface,
        RECT {
            left: 0,
            top: 0,
            right: surface.width as i32,
            bottom: surface.height as i32,
        },
        page_color,
    );

    let mut painter = RunPainter::new(surface, scale);
    for line in &page.lines {
        if line.rule {
            fill_rect(
                surface,
                RECT {
                    left: page.padding,
                    top: line.top,
                    right: surface.width as i32 - page.padding,
                    bottom: line.top + line.height.max(1),
                },
                rule_color,
            );
            continue;
        }

        for run in &line.runs {
            painter.draw(
                &run.text,
                run.x,
                RECT {
                    left: run.x,
                    top: line.top,
                    right: run.x + run.width,
                    bottom: line.top + line.height,
                },
                &run.style,
                run.style.foreground,
                page_color,
            );
        }
    }
}

// ----------------------------------------------------------------------- text

/// The width of a run of text, counted the way the rest of the page is: the face
/// is fixed pitch, so a count of characters is the whole measurement.
fn text_width(text: &str, advance: i32) -> i32 {
    text.chars().count() as i32 * advance.max(1)
}

/// A name cut to the room it has, with an ellipsis where it was cut. A name that
/// does not fit at all is dropped rather than drawn over the size column.
fn cut_to_width(text: &str, room: i32, advance: i32) -> String {
    let advance = advance.max(1);
    let room = (room / advance).max(0) as usize;
    let characters = text.chars().count();
    if characters <= room {
        return text.to_string();
    }
    if room == 0 {
        return String::new();
    }

    let mut cut: String = text.chars().take(room - 1).collect();
    cut.push('…');

    cut
}

/// A byte count as a person reads it: whole bytes below a kilobyte, one decimal
/// where the fraction says something and none where it does not, so a folder
/// view reads `4 KB` and `1.1 MB` rather than `4.0 KB`.
fn format_size(bytes: u64) -> String {
    const UNITS: [&str; 5] = ["B", "KB", "MB", "GB", "TB"];

    let mut value = bytes as f64;
    let mut unit = 0usize;
    while value >= 1024.0 && unit + 1 < UNITS.len() {
        value /= 1024.0;
        unit += 1;
    }

    if unit == 0 {
        return format!("{bytes} B");
    }

    let rounded = (value * 10.0).round() / 10.0;
    if rounded < 10.0 && rounded.fract() != 0.0 {
        format!("{rounded:.1} {}", UNITS[unit])
    } else {
        format!("{rounded:.0} {}", UNITS[unit])
    }
}

/// The glyph an entry is drawn with, by what it is and what it is called.
fn icon_glyph(is_dir: bool, name: &str) -> char {
    if is_dir {
        return ICON_FOLDER;
    }

    let extension = name
        .rsplit_once('.')
        .map(|(_, extension)| extension.to_ascii_lowercase())
        .unwrap_or_default();

    match extension.as_str() {
        "png" | "jpg" | "jpeg" | "jfif" | "gif" | "bmp" | "webp" | "tif" | "tiff" | "ico"
        | "svg" | "avif" | "heic" | "psd" | "tga" | "qoi" | "pbm" | "ff" => ICON_IMAGE,
        "mp4" | "mkv" | "avi" | "mov" | "wmv" | "webm" | "flv" | "m4v" | "mpg" | "mpeg" | "ts"
        | "mts" | "m2ts" => ICON_VIDEO,
        "mp3" | "wav" | "flac" | "ogg" | "m4a" | "aac" | "wma" | "opus" | "aiff" => ICON_AUDIO,
        "pdf" => ICON_PDF,
        "zip" | "zipx" | "rar" | "7z" | "tar" | "gz" | "tgz" | "jar" | "apk" | "xpi" | "cbz"
        | "bz2" | "xz" | "zst" | "lz" | "lzma" | "cab" | "iso" => ICON_ARCHIVE,
        "rs" | "c" | "h" | "cpp" | "cxx" | "hpp" | "cs" | "js" | "mjs" | "tsx" | "jsx" | "py"
        | "rb" | "go" | "java" | "kt" | "swift" | "php" | "pl" | "lua" | "sh" | "bash" | "ps1"
        | "bat" | "cmd" | "json" | "xml" | "yml" | "yaml" | "toml" | "ini" | "cfg" | "conf"
        | "html" | "htm" | "css" | "scss" | "sql" | "md" | "markdown" | "diff" | "patch" => {
            ICON_CODE
        }
        _ => ICON_FILE,
    }
}

/// The icon face this machine has, decided once. A preview without an icon is
/// still a preview; a preview drawn in a face that does not exist would be a row
/// of boxes.
fn icon_face() -> Option<&'static str> {
    *ICON_FACE
}

static ICON_FACE: Lazy<Option<&'static str>> = Lazy::new(detect_icon_face);

fn detect_icon_face() -> Option<&'static str> {
    let dc = unsafe { CreateCompatibleDC(None) };
    if dc.0.is_null() {
        return None;
    }

    let chosen = ICON_FACES.into_iter().find(|face| face_installed(dc, face));

    unsafe {
        let _ = DeleteDC(dc);
    }

    chosen
}

struct FaceSearch<'a> {
    target: &'a str,
    found: bool,
}

unsafe extern "system" fn face_visitor(
    logfont: *const LOGFONTW,
    _metric: *const TEXTMETRICW,
    _font_type: u32,
    lparam: LPARAM,
) -> i32 {
    let search = &mut *(lparam.0 as *mut FaceSearch);
    let face = &(*logfont).lfFaceName;
    let length = face.iter().position(|c| *c == 0).unwrap_or(face.len());
    let name = String::from_utf16_lossy(&face[..length]);

    if name.eq_ignore_ascii_case(search.target) {
        search.found = true;
        return 0;
    }

    1
}

fn face_installed(dc: HDC, face: &str) -> bool {
    let mut search = FaceSearch {
        target: face,
        found: false,
    };
    let logfont = LOGFONTW {
        lfCharSet: windows::Win32::Graphics::Gdi::DEFAULT_CHARSET,
        ..Default::default()
    };

    unsafe {
        EnumFontFamiliesExW(
            dc,
            &logfont,
            Some(face_visitor),
            LPARAM(&mut search as *mut FaceSearch as isize),
            0,
        );
    }

    search.found
}

// ---------------------------------------------------------------------- tests

#[cfg(test)]
mod tests;

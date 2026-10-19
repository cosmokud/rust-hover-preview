//! What every mark in this module is drawn from: the palette they are painted in, a pixel
//! written premultiplied, the shapes a panel and a glyph are made of, and the walks a caption
//! carries and a file that could not be drawn stands in for.
//!
//! Nothing here knows what it is drawing for, which is the whole of what it is for: a caption's
//! close cross and a failed file's are one function (`draw_cross`), and a panel, its outline and
//! the groove inside it cannot disagree about where they end because they are one shape worked
//! out once (`round_rect_coverage`). The text runs are the exception to "drawn by hand": they are
//! measured through GDI, which is the only thing here that knows what a font is.

use crate::text::text_paint::{self, DibSurface};
use crate::CONFIG;
use windows::Win32::Foundation::{RECT, SIZE};
use windows::Win32::Graphics::Gdi::{GetTextExtentPoint32W, SelectObject};

/// The colors a caption and a bubble are painted in: the theme's page, its text, and one it
/// spends on something else, which is drawn as the accent.
///
/// They come from the theme the text previews are painted with — One Dark Pro, Atom One Light,
/// or a `.tmTheme` of the user's own — so that the chrome of a pinned preview belongs to the
/// same app as the page inside it.
pub(crate) struct ChromePalette {
    /// The caption's bar and the bubble's face.
    pub(crate) background: [u8; 3],
    /// The title and the glyphs.
    pub(crate) foreground: [u8; 3],
    /// The bubble's ring, and the played part of a transport's bar.
    pub(crate) accent: [u8; 3],
    /// Whether the theme is a dark one, which decides which way a hairline is shaded.
    pub(crate) dark: bool,
}

impl ChromePalette {
    /// The palette the configuration currently names, or `None` when no theme can be read
    /// at all — which the bundled themes make impossible in practice, and which is answered
    /// with no chrome rather than with guessed colors (see `text_theme::loaded`).
    pub(crate) fn current() -> Option<Self> {
        let kind = CONFIG.lock().ok()?.theme;
        let theme = crate::text::text_theme::loaded(kind)?;

        let background = text_paint::rgb(theme.background());
        let foreground = text_paint::rgb(theme.foreground());
        // The same accent a sound's card draws its bar and its clock in, so that the knob on this
        // chrome and the track on that card are one color rather than two near neighbours (see
        // `audio_preview::build_page`).
        let accent = text_paint::rgb(theme.style_for_scopes(&["support.function"]).foreground);

        Some(Self {
            background,
            foreground,
            accent,
            dark: text_paint::luminance(background) <= 0.5,
        })
    }

    /// A color blended toward the opposite end of the theme, for the surface a pointer is
    /// over: a caption button lights up, and the rest of the caption stays where it is.
    pub(crate) fn hover(&self, amount: f32) -> [u8; 3] {
        let toward = if self.dark {
            [255, 255, 255]
        } else {
            [0, 0, 0]
        };
        text_paint::blend(self.background, toward, amount)
    }
}

/// The face a caption's title is set in: the one Windows writes a window's title in, so a
/// pinned preview's caption reads like the caption of anything else on the desktop.
const CAPTION_FACE: &str = "Segoe UI";

pub(super) const GLYPH_STROKE_PIXELS: f32 = 1.4;
pub(super) const GLYPH_PIXELS: f32 = 10.0;

/// The surface's pixels, which every box, disc and glyph here is drawn into.
///
/// # Safety
///
/// The surface must be one GDI is not drawing into, and must not be handed back to GDI
/// until the slice built from this has ended. A raw pointer rather than a `&mut`, because
/// the surface is reached by shared reference and the pixels behind it are the only thing
/// written: the caller gets a slice for the length of its own block, and nothing else.
pub(crate) unsafe fn surface_pixels(surface: &DibSurface) -> *mut u8 {
    surface.bits()
}

/// One pixel written premultiplied (see the module documentation).
pub(super) fn put(buffer: &mut [u8], width: i32, x: i32, y: i32, color: [u8; 3], coverage: f32) {
    if x < 0 || y < 0 || x >= width {
        return;
    }

    let index = (y as usize * width as usize + x as usize) * 4;
    if index + 3 >= buffer.len() {
        return;
    }

    let alpha = (coverage.clamp(0.0, 1.0) * 255.0).round() as u32;
    if alpha == 0 {
        return;
    }

    let existing = [
        buffer[index],
        buffer[index + 1],
        buffer[index + 2],
        buffer[index + 3],
    ];
    let existing_alpha = existing[3] as u32;
    let inverse = 255 - alpha;

    // What is already there is premultiplied too, so it composites the way any two
    // premultiplied colors do.
    buffer[index] = (((color[2] as u32 * alpha) + existing[0] as u32 * inverse + 127) / 255) as u8;
    buffer[index + 1] =
        (((color[1] as u32 * alpha) + existing[1] as u32 * inverse + 127) / 255) as u8;
    buffer[index + 2] =
        (((color[0] as u32 * alpha) + existing[2] as u32 * inverse + 127) / 255) as u8;
    buffer[index + 3] = (alpha + (existing_alpha * inverse + 127) / 255).min(255) as u8;
}

pub(super) fn fill_box(buffer: &mut [u8], width: i32, rect: RECT, color: [u8; 3], coverage: f32) {
    for y in rect.top..rect.bottom {
        for x in rect.left..rect.right {
            put(buffer, width, x, y, color, coverage);
        }
    }
}

/// The outline of a box, drawn inside its own edges so that a glyph is the size it was asked
/// for rather than a stroke wider on each side.
pub(super) fn stroke_box(
    buffer: &mut [u8],
    width: i32,
    rect: RECT,
    thickness: f32,
    color: [u8; 3],
    coverage: f32,
) {
    let thickness = thickness.round().max(1.0) as i32;
    let (left, top, right, bottom) = (rect.left, rect.top, rect.right, rect.bottom);

    fill_box(
        buffer,
        width,
        RECT {
            left,
            top,
            right,
            bottom: (top + thickness).min(bottom),
        },
        color,
        coverage,
    );
    fill_box(
        buffer,
        width,
        RECT {
            left,
            top: (bottom - thickness).max(top),
            right,
            bottom,
        },
        color,
        coverage,
    );
    fill_box(
        buffer,
        width,
        RECT {
            left,
            top,
            right: (left + thickness).min(right),
            bottom,
        },
        color,
        coverage,
    );
    fill_box(
        buffer,
        width,
        RECT {
            left: (right - thickness).max(left),
            top,
            right,
            bottom,
        },
        color,
        coverage,
    );
}

/// Fill a rounded rectangle, shaded from one colour at its top edge to another at its bottom: a
/// popup panel is the one thing here that is not a flat surface.
pub(super) fn fill_round_rect(
    buffer: &mut [u8],
    width: i32,
    rect: RECT,
    radius: f32,
    top: [u8; 3],
    bottom: [u8; 3],
    coverage: f32,
) {
    if rect.right <= rect.left || rect.bottom <= rect.top {
        return;
    }

    let height = (rect.bottom - rect.top) as f32;
    for y in rect.top..rect.bottom {
        let color = text_paint::blend(top, bottom, (y - rect.top) as f32 / height);
        for x in rect.left..rect.right {
            let coverage = coverage * round_rect_coverage(x, y, rect, radius);
            if coverage > 0.0 {
                put(buffer, width, x, y, color, coverage);
            }
        }
    }
}

/// The outline of a rounded rectangle, drawn inside its own edge so that the shape under it is
/// the size it was asked for.
pub(super) fn stroke_round_rect(
    buffer: &mut [u8],
    width: i32,
    rect: RECT,
    radius: f32,
    thickness: f32,
    color: [u8; 3],
    coverage: f32,
) {
    let thickness = thickness.round().max(1.0) as i32;
    let inner = RECT {
        left: rect.left + thickness,
        top: rect.top + thickness,
        right: rect.right - thickness,
        bottom: rect.bottom - thickness,
    };
    if inner.right <= inner.left || inner.bottom <= inner.top {
        return;
    }

    let inner_radius = (radius - thickness as f32).max(0.0);
    for y in rect.top..rect.bottom {
        for x in rect.left..rect.right {
            let outline = (round_rect_coverage(x, y, rect, radius)
                - round_rect_coverage(x, y, inner, inner_radius))
            .max(0.0);
            if outline > 0.0 {
                put(buffer, width, x, y, color, coverage * outline);
            }
        }
    }
}

/// How much of a pixel a rounded rectangle covers, worked out from the distance to its edge: one
/// shape, so that a panel, its outline and the groove in it cannot disagree about where they end.
fn round_rect_coverage(x: i32, y: i32, rect: RECT, radius: f32) -> f32 {
    let half_width = (rect.right - rect.left) as f32 / 2.0;
    let half_height = (rect.bottom - rect.top) as f32 / 2.0;
    if half_width <= 0.0 || half_height <= 0.0 {
        return 0.0;
    }

    let radius = radius.clamp(0.0, half_width.min(half_height));
    let center_x = rect.left as f32 + half_width;
    let center_y = rect.top as f32 + half_height;
    let dx = ((x as f32 + 0.5 - center_x).abs() - (half_width - radius)).max(0.0);
    let dy = ((y as f32 + 0.5 - center_y).abs() - (half_height - radius)).max(0.0);
    let distance = (dx * dx + dy * dy).sqrt() - radius;

    (0.5 - distance).clamp(0.0, 1.0)
}

/// Fill a circle: the knob a level is dragged by, the collar around it, and the dot a sound's
/// card is marked with (see `audio_preview`).
pub(crate) fn fill_disc(
    buffer: &mut [u8],
    width: i32,
    center_x: f32,
    center_y: f32,
    radius: f32,
    color: [u8; 3],
    coverage: f32,
) {
    let radius = radius.max(0.5);
    let left = (center_x - radius - 1.0).floor() as i32;
    let right = (center_x + radius + 1.0).ceil() as i32;
    let top = (center_y - radius - 1.0).floor() as i32;
    let bottom = (center_y + radius + 1.0).ceil() as i32;

    for y in top..=bottom {
        for x in left..=right {
            let dx = x as f32 + 0.5 - center_x;
            let dy = y as f32 + 0.5 - center_y;
            let edge = (radius - (dx * dx + dy * dy).sqrt() + 0.5).clamp(0.0, 1.0);
            if edge > 0.0 {
                put(buffer, width, x, y, color, coverage * edge);
            }
        }
    }
}

/// Fill a ring: the outline of a circle at `radius`, `thickness` wide — the shape a sound's beat
/// travels as, around the mark beside its name (see `audio_preview`).
#[allow(clippy::too_many_arguments)]
pub(crate) fn fill_ring(
    buffer: &mut [u8],
    width: i32,
    center_x: f32,
    center_y: f32,
    radius: f32,
    thickness: f32,
    color: [u8; 3],
    coverage: f32,
) {
    let outer = radius + thickness / 2.0;
    let inner = (radius - thickness / 2.0).max(0.0);
    let left = (center_x - outer - 1.0).floor() as i32;
    let right = (center_x + outer + 1.0).ceil() as i32;
    let top = (center_y - outer - 1.0).floor() as i32;
    let bottom = (center_y + outer + 1.0).ceil() as i32;

    for y in top..=bottom {
        for x in left..=right {
            let dx = x as f32 + 0.5 - center_x;
            let dy = y as f32 + 0.5 - center_y;
            let distance = (dx * dx + dy * dy).sqrt();
            // The share of the pixel the band leaves: the ramp the outer edge is drawn with — the
            // ramp a disc's own edge is drawn with (see `fill_disc`) — less the ramp the hole cuts.
            let edge =
                (outer - distance + 0.5).clamp(0.0, 1.0) - (inner - distance + 0.5).clamp(0.0, 1.0);
            if edge > 0.0 {
                put(buffer, width, x, y, color, coverage * edge);
            }
        }
    }
}

/// Fill a triangle, by the distance to each of its three edges: what a pixel carries is the share
/// of it the nearest edge leaves, which is the ramp a disc's edge is drawn with (see `fill_disc`)
/// asked of three straight lines instead of one circle.
///
/// A point is inside a triangle where all three of its edges are on the far side of it, and
/// outside by as much as the furthest of them — so the three are asked per pixel and the shape is
/// whatever they say it is, with no scanline and no rounding to whole columns. That is what a
/// slanted edge needs and a column-by-column scan cannot give it: every column of a triangle is a
/// whole number of pixels tall, and the edge of one is the steps between them. It is what the three
/// marks of this chrome that are nothing but a slope are drawn with: the play mark of a bar, the
/// triangle beside each of its step bars, and the play mark of a collapsed pin's bubble
/// (`paint_mark`).
pub(super) fn fill_triangle(
    buffer: &mut [u8],
    width: i32,
    corners: [(f32, f32); 3],
    color: [u8; 3],
    coverage: f32,
) {
    // Each edge as a line: a unit normal pointing out of the triangle, and where the line it is
    // on crosses the origin. The normal points away from the corner the edge is not on, which is
    // the one answer that does not depend on which way round the three were given.
    let mut edges = [(0.0f32, 0.0f32, 0.0f32); 3];
    for index in 0..3 {
        let from = corners[index];
        let to = corners[(index + 1) % 3];
        let (dx, dy) = (to.0 - from.0, to.1 - from.1);
        let length = (dx * dx + dy * dy).sqrt().max(f32::EPSILON);
        let (nx, ny) = (dy / length, -dx / length);
        let opposite = corners[(index + 2) % 3];
        let outward = if nx * (opposite.0 - from.0) + ny * (opposite.1 - from.1) > 0.0 {
            -1.0
        } else {
            1.0
        };
        let (nx, ny) = (nx * outward, ny * outward);
        edges[index] = (nx, ny, -(nx * from.0 + ny * from.1));
    }

    let mut left = f32::INFINITY;
    let mut right = f32::NEG_INFINITY;
    let mut top = f32::INFINITY;
    let mut bottom = f32::NEG_INFINITY;
    for corner in corners {
        left = left.min(corner.0);
        right = right.max(corner.0);
        top = top.min(corner.1);
        bottom = bottom.max(corner.1);
    }
    // The box, a pixel wider than the corners on every side, so that a pixel the edge only clips
    // is asked about rather than missed.
    let (left, right) = ((left - 1.0).floor() as i32, (right + 1.0).ceil() as i32);
    let (top, bottom) = ((top - 1.0).floor() as i32, (bottom + 1.0).ceil() as i32);

    for y in top..=bottom {
        for x in left..=right {
            let (point_x, point_y) = (x as f32 + 0.5, y as f32 + 0.5);
            let outside = edges
                .iter()
                .map(|&(nx, ny, at)| nx * point_x + ny * point_y + at)
                .fold(f32::NEG_INFINITY, f32::max);
            let edge = (0.5 - outside).clamp(0.0, 1.0);
            if edge > 0.0 {
                put(buffer, width, x, y, color, coverage * edge);
            }
        }
    }
}

/// Fill a polygon, sampled rather than scanned: the shapes drawn here are a dozen pixels across,
/// so a handful of samples a pixel costs less than a scanline that has to be right about every
/// edge of a shape as small as a glyph.
pub(super) fn fill_polygon(
    buffer: &mut [u8],
    width: i32,
    points: &[(f32, f32)],
    color: [u8; 3],
    coverage: f32,
) {
    const SAMPLES: i32 = 4;

    if points.len() < 3 {
        return;
    }

    let left = points
        .iter()
        .map(|p| p.0)
        .fold(f32::INFINITY, f32::min)
        .floor() as i32;
    let right = points
        .iter()
        .map(|p| p.0)
        .fold(f32::NEG_INFINITY, f32::max)
        .ceil() as i32;
    let top = points
        .iter()
        .map(|p| p.1)
        .fold(f32::INFINITY, f32::min)
        .floor() as i32;
    let bottom = points
        .iter()
        .map(|p| p.1)
        .fold(f32::NEG_INFINITY, f32::max)
        .ceil() as i32;

    for y in top..=bottom {
        for x in left..=right {
            let mut inside = 0;
            for sample_y in 0..SAMPLES {
                for sample_x in 0..SAMPLES {
                    let point = (
                        x as f32 + (sample_x as f32 + 0.5) / SAMPLES as f32,
                        y as f32 + (sample_y as f32 + 0.5) / SAMPLES as f32,
                    );
                    if point_in_polygon(point, points) {
                        inside += 1;
                    }
                }
            }

            if inside > 0 {
                let edge = inside as f32 / (SAMPLES * SAMPLES) as f32;
                put(buffer, width, x, y, color, coverage * edge);
            }
        }
    }
}

fn point_in_polygon(point: (f32, f32), points: &[(f32, f32)]) -> bool {
    let mut inside = false;
    let mut previous = points.len() - 1;

    for current in 0..points.len() {
        let (x, y) = points[current];
        let (previous_x, previous_y) = points[previous];

        if (y > point.1) != (previous_y > point.1)
            && point.0 < (previous_x - x) * (point.1 - y) / (previous_y - y) + x
        {
            inside = !inside;
        }

        previous = current;
    }

    inside
}

/// One band of a circle, drawn as the pixels a given distance from a center and within the spread
/// of angles a speaker's sound is drawn in: the two arcs a volume button carries.
pub(super) fn stroke_arc(
    buffer: &mut [u8],
    width: i32,
    center: (f32, f32),
    radius: f32,
    thickness: f32,
    color: [u8; 3],
    coverage: f32,
) {
    let half = (thickness / 2.0).max(0.5);
    let reach = radius + half + 1.0;
    let left = (center.0 - reach).floor() as i32;
    let right = (center.0 + reach).ceil() as i32;
    let top = (center.1 - reach).floor() as i32;
    let bottom = (center.1 + reach).ceil() as i32;

    // The ends of an arc are square to the axis it opens from, with a sliver of fade so that an
    // end is a soft edge rather than a stair step.
    const SPREAD: f32 = 0.85;
    let feather = 0.14;

    for y in top..=bottom {
        for x in left..=right {
            let dx = x as f32 + 0.5 - center.0;
            let dy = y as f32 + 0.5 - center.1;
            let distance = (dx * dx + dy * dy).sqrt();
            let radial = (half + 0.5 - (distance - radius).abs()).clamp(0.0, 1.0);
            if radial <= 0.0 {
                continue;
            }

            let angle = dy.abs().atan2(dx);
            let spread = ((SPREAD - angle) / feather).clamp(0.0, 1.0);
            if spread > 0.0 {
                put(buffer, width, x, y, color, coverage * radial * spread);
            }
        }
    }
}

/// A line between two points, drawn as the pixels within half its thickness of it.
///
/// Every mark here is drawn through it, and so is a pinned sound's card's own two
/// window buttons, whose marks are a dash and a cross this size cannot be told
/// apart from anything else (see `audio_preview::paint_window_button`).
pub(crate) fn stroke_segment(
    buffer: &mut [u8],
    width: i32,
    from: (f32, f32),
    to: (f32, f32),
    thickness: f32,
    color: [u8; 3],
    coverage: f32,
) {
    let half = (thickness / 2.0).max(0.5);
    let left = (from.0.min(to.0) - half - 1.0).floor() as i32;
    let right = (from.0.max(to.0) + half + 1.0).ceil() as i32;
    let top = (from.1.min(to.1) - half - 1.0).floor() as i32;
    let bottom = (from.1.max(to.1) + half + 1.0).ceil() as i32;

    for y in top..=bottom {
        for x in left..=right {
            let distance = distance_to_segment((x as f32 + 0.5, y as f32 + 0.5), from, to);
            let edge = (half + 0.5 - distance).clamp(0.0, 1.0);
            if edge > 0.0 {
                put(buffer, width, x, y, color, coverage * edge);
            }
        }
    }
}

fn distance_to_segment(point: (f32, f32), from: (f32, f32), to: (f32, f32)) -> f32 {
    let (dx, dy) = (to.0 - from.0, to.1 - from.1);
    let length_squared = dx * dx + dy * dy;
    let along = if length_squared <= f32::EPSILON {
        0.0
    } else {
        (((point.0 - from.0) * dx + (point.1 - from.1) * dy) / length_squared).clamp(0.0, 1.0)
    };

    let nearest = (from.0 + along * dx, from.1 + along * dy);
    ((point.0 - nearest.0).powi(2) + (point.1 - nearest.1).powi(2)).sqrt()
}

/// The two diagonals of a cross, drawn as the two strokes they are rather than walked a column at
/// a time so that a stroke is one width whoever draws it and whatever the display's scale is.
///
/// A diagonal is the one mark a walk of whole columns cannot draw: every column of it is a whole
/// number of rows tall, so what comes out is the steps between them, and a cross ten pixels across
/// is short enough that those steps are the mark. Each diagonal is a `stroke_segment`, which is the
/// same ramp the box and the arrow of the hand-off mark beside it are drawn with.
pub(super) fn draw_cross(
    buffer: &mut [u8],
    width: i32,
    center_x: i32,
    center_y: i32,
    span: i32,
    thickness: f32,
    color: [u8; 3],
) {
    // What the mark asks for is the rows a *column* of the diagonal carries, and across the stroke
    // that is the width times the root of two: the strokes are drawn at the width they are, so that
    // the mark weighs what it weighed when it was walked and only its edge is different.
    let edge = (thickness * std::f32::consts::FRAC_1_SQRT_2).max(1.0);
    let half = span as f32 / 2.0;

    // From the middle of one corner pixel to the middle of the one opposite it, which is the two
    // ends the walk drew and no more: a stroke that stopped short of them would leave the corners
    // of the mark lit by less than the rest of it.
    let (middle_x, middle_y) = (center_x as f32 + 0.5, center_y as f32 + 0.5);
    for down in [1.0, -1.0] {
        stroke_segment(
            buffer,
            width,
            (middle_x - half, middle_y - half * down),
            (middle_x + half, middle_y + half * down),
            edge,
            color,
            1.0,
        );
    }
}

/// One chevron pointing left or right: two arms, each of them the stroke it is, closing on the
/// column the mark points at.
///
/// A chevron rather than an arrowhead: a walk has no end, so what the two buttons mean is
/// which way along it to go and not where it stops.
///
/// The vertex is at the end the chevron points to and the arms open away from it. Measured
/// from the middle of the span instead, the mark is a cross rather than a chevron — which is
/// what a `<` and a `>` would both be drawn as.
///
/// The arms part *half* a span either side of the middle row, so that the mark comes out the
/// size the close cross and the maximize box beside it are drawn at.
///
/// Two strokes rather than a walk of whole columns, because an arm is a slanted edge and a walk
/// has nothing but the steps between its columns to draw one with. Each stops half a stroke-width
/// short of the vertex, which is where a stroke's rounded end reaches with the ramp it is drawn by,
/// so the two of them light the column the mark points at and stop inside it: a chevron whose arms
/// met at the vertex carries a pixel more, each way, than the glyph square it is asked for once the
/// display is at 200%.
#[allow(clippy::too_many_arguments)] // Which way it points is a fact about the mark, not a drawing flag.
pub(super) fn draw_chevron(
    buffer: &mut [u8],
    width: i32,
    center_x: i32,
    center_y: i32,
    span: i32,
    thickness: f32,
    color: [u8; 3],
    pointing_right: bool,
) {
    // What the mark asks for is the rows a *column* of an arm carries, and the stroke carrying
    // those rows is narrower across than they are: by the root of two, which is the conversion the
    // close cross's diagonal is drawn with. This arm's shallower slope would ask for a wider stroke,
    // and one wide enough for those rows carries its round ends a row outside the square the walk
    // covered once the display is at 200% — so the arms are drawn the narrower of the two.
    let edge = (thickness * std::f32::consts::FRAC_1_SQRT_2).max(1.0);
    let half = span as f32 / 2.0;
    let (middle_x, middle_y) = (center_x as f32 + 0.5, center_y as f32 + 0.5);
    let point = if pointing_right { half } else { -half };

    // Stopped half a stroke-width short of the vertex, which is where the end it rounds off with
    // sits in the column the mark points at rather than in the one beside it.
    let stop = edge / 2.0;
    for down in [-1.0, 1.0] {
        // From the middle of the end pixel the walk opened at to the middle of the column it
        // pointed into, which is the same pair of ends the walk covered and no more: an arm
        // stopped short of them would leave the mark's own corners lit by less than the rest.
        let far = (middle_x - point, middle_y + down * half);
        let vertex = (middle_x + point, middle_y);
        let (dx, dy) = (vertex.0 - far.0, vertex.1 - far.1);
        let length = (dx * dx + dy * dy).sqrt().max(f32::EPSILON);

        stroke_segment(
            buffer,
            width,
            far,
            (vertex.0 - dx / length * stop, vertex.1 - dy / length * stop),
            edge,
            color,
            1.0,
        );
    }
}

/// An arrow leaving a box: the file handed to whatever the machine has filed it under.
///
/// Drawn as the box and the arrow together rather than as a box with a mark in it, because the
/// button's whole meaning is the leaving. The box is drawn in every side but the one the arrow
/// goes out of, and the reading it has to carry is *leaving this thing* — a box drawn too small
/// to read as a box with a big arrow on top of it is the upload mark, and a button that means
/// "hand this to another program" must not be drawn as a button that means "send this away".
///
/// The tail starts at the middle of the box and crosses its open side to run out past the
/// corner, so that the arrow is seen leaving rather than hovering over.
pub(super) fn draw_open_with(
    buffer: &mut [u8],
    width: i32,
    center_x: i32,
    center_y: i32,
    span: i32,
    thickness: f32,
    color: [u8; 3],
) {
    let edge = thickness.round().max(1.0);
    let half = ((span as f32 / 2.0) - edge / 2.0).max(0.0);
    let (middle_x, middle_y) = (center_x as f32 + 0.5, center_y as f32 + 0.5);
    let point = |x: f32, y: f32| (middle_x + x * half, middle_y + y * half);
    let open_at = |y: f32| (point(-1.0, y), point(1.0, y));

    // The file: a box with its top right corner left off, so that the corner the arrow goes
    // out of is the one that is not drawn.
    let (top_left, top_right) = open_at(-1.0);
    let (bottom_left, bottom_right) = open_at(1.0);
    stroke_segment(buffer, width, top_left, top_right, edge, color, 1.0);
    stroke_segment(buffer, width, bottom_left, bottom_right, edge, color, 1.0);
    stroke_segment(buffer, width, top_left, bottom_left, edge, color, 1.0);
    stroke_segment(
        buffer,
        width,
        top_right,
        point(1.0, -1.0 / 3.0),
        edge,
        color,
        1.0,
    );

    // The hand-off: a tail out of the middle of the file, through the open side and past the
    // corner it goes out of, and the head it ends in — one stroke along the top of it and one
    // along the side of it, which is the same pair a chevron is drawn in.
    let tip = point(1.0, -1.0);
    stroke_segment(buffer, width, point(-0.2, 0.2), tip, edge, color, 1.0);
    stroke_segment(buffer, width, tip, point(1.0 / 3.0, -1.0), edge, color, 1.0);
    stroke_segment(buffer, width, tip, point(1.0, -1.0 / 3.0), edge, color, 1.0);
}

/// A ceiling bar and an upward-pointing arrow: always on top mark.
///
/// The mark is drawn to the glyph square the window's own marks beside it are drawn in: its
/// ceiling is that square's top row and its stem ends on the square's bottom row, so the pin
/// stands as tall as the maximize box and the close cross at whatever scale the display is at.
/// The bar and the stem are laid down in whole pixels rather than as strokes, because the rows
/// they end on are the whole of what "level with the buttons beside it" means, and a stroke
/// centred on the square's boundary would spend half its coverage on the row beyond it.
pub(crate) fn draw_pin_top(
    buffer: &mut [u8],
    width: i32,
    center_x: i32,
    center_y: i32,
    span: i32,
    thickness: f32,
    color: [u8; 3],
) {
    let edge = thickness.round().max(1.0);
    let thick = edge as i32;
    let half = (span / 2).max(1);
    let square = RECT {
        left: center_x - half,
        top: center_y - half,
        right: center_x + half + 1,
        bottom: center_y + half + 1,
    };

    // The ceiling: the square's own top border, as wide as the square.
    fill_box(
        buffer,
        width,
        RECT {
            left: square.left,
            top: square.top,
            right: square.right,
            bottom: (square.top + thick).min(square.bottom),
        },
        color,
        1.0,
    );

    // The stem down to the square's bottom row, as wide as the bar is thick.
    let stem_left = center_x - thick / 2;
    fill_box(
        buffer,
        width,
        RECT {
            left: stem_left.max(square.left),
            top: (square.top + thick + 1).min(square.bottom),
            right: (stem_left + thick).min(square.right),
            bottom: square.bottom,
        },
        color,
        1.0,
    );

    // The head: two arms from the top of the stem out to the arrow's own width, which is the
    // square less the margin the arms have to keep from its sides to stay clear of them.
    let (middle_x, middle_y) = (center_x as f32 + 0.5, center_y as f32 + 0.5);
    let tip = (middle_x, (square.top + thick + 1) as f32 + 0.5);
    let arm = (half - thick - 1).max(1) as f32;
    for side in [-1.0, 1.0] {
        stroke_segment(
            buffer,
            width,
            tip,
            (middle_x + side * arm, middle_y),
            edge,
            color,
            1.0,
        );
    }
}

/// A list of rows above a chevron pointing down: the file handed to a program named out loud.
///
/// It is not the hand-off mark beside it with a mark added, and the difference is the whole of
/// what the two buttons are: the one before opens the file with the program the machine has
/// already chosen, which is a *going* and is drawn as an arrow leaving a box, while this one
/// opens a list of what else could — which is a *choosing*, and is drawn as the list itself. Two
/// boxes and two arrows on one strip would be a row of targets a hand has to read one at a time,
/// which is what a walk beside the window's own three is not for.
///
/// The rows are three lines of the same length rather than a bulleted list, because a bullet is
/// a dot this size does not survive and a line is. The chevron below them is the same pair of
/// strokes a `Previous` is drawn in, which is what makes the two read as one family.
pub(super) fn draw_open_with_list(
    buffer: &mut [u8],
    width: i32,
    center_x: i32,
    center_y: i32,
    span: i32,
    thickness: f32,
    color: [u8; 3],
) {
    let edge = thickness.round().max(1.0);

    // The list: three rows across the top of the mark, each as wide as the mark itself, so
    // that it reads as a list rather than as three specks. A row shorter than the chevron
    // below it is a bullet, and a bullet at ten pixels is a dot. It is walked in whole pixels
    // because a row of a list is a row: the rows are as wide as the mark is, so their ends land
    // on the pixel grid, and a row drawn in halves would be a row of gaps.
    let half = span / 2;
    for offset in [0, half / 2, half] {
        for x in (center_x - half)..=(center_x + half) {
            put(buffer, width, x, center_y - half + offset, color, 1.0);
        }
    }

    // The chevron below them, pointing down and drawn as the two arms a `Previous` is drawn
    // in, reaching to the same edge the list's top row does: the list is what is being
    // offered, and this is the sign that there is more of it than the button can hold. A
    // chevron that stopped short of the top row would leave the mark sitting high on its own
    // button, which is the one thing a row of marks beside each other cannot afford.
    let shoulder = half as f32 * 0.8;
    let top = center_y as f32 + half as f32 * 0.35;
    let (left, right) = (
        (center_x as f32 - shoulder, top),
        (center_x as f32 + shoulder, top),
    );
    let tip = (center_x as f32 + 0.5, center_y as f32 + half as f32 + 0.5);
    stroke_segment(buffer, width, left, tip, edge, color, 1.0);
    stroke_segment(buffer, width, right, tip, edge, color, 1.0);
}

/// The style every piece of caption text is drawn in, whatever colour it is asked for.
pub(super) fn caption_style(foreground: [u8; 3]) -> text_paint::TextStyle {
    text_paint::TextStyle {
        foreground,
        background: None,
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        level: text_paint::BODY_LEVEL,
        face: CAPTION_FACE,
    }
}

/// One of the two clocks: where the file is, or how long it is. A file with no length is drawn as
/// a pair of dashes rather than as a zero, which is the truth about it.
pub(super) fn paint_time(
    surface: &DibSurface,
    palette: &ChromePalette,
    seconds: Option<f64>,
    rect: RECT,
    scale: f32,
    align_right: bool,
) {
    let text = clock_text(seconds);
    let style = caption_style(palette.foreground);

    let cell = text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[text_paint::BODY_LEVEL as usize],
        scale,
    );
    let top = ((surface.height as i32 - cell) / 2).max(0);
    let left = if align_right {
        let measured = measure_text(surface, &style, &text, scale);
        (rect.right - measured - 1).max(rect.left)
    } else {
        rect.left + 1
    };

    let mut painter = text_paint::RunPainter::new(surface, scale);
    painter.draw(
        &text,
        left,
        RECT {
            left: rect.left,
            top,
            right: rect.right,
            bottom: (top + cell).min(surface.height as i32),
        },
        &style,
        palette.foreground,
        palette.background,
    );
}

/// How a number of seconds is written on a transport bar: minutes and seconds, hours where there
/// are any, and dashes where there is nothing to write.
pub(crate) fn clock_text(seconds: Option<f64>) -> String {
    let Some(seconds) = seconds.filter(|value| value.is_finite() && *value >= 0.0) else {
        return "--:--".to_string();
    };

    let total = seconds.round() as u64;
    let (hours, minutes, seconds) = (total / 3600, (total % 3600) / 60, total % 60);

    if hours > 0 {
        format!("{hours}:{minutes:02}:{seconds:02}")
    } else {
        format!("{minutes}:{seconds:02}")
    }
}

/// The width of a run of text at a style, which the two clocks need to be placed against the
/// edges they sit on.
pub(super) fn measure_text(
    surface: &DibSurface,
    style: &text_paint::TextStyle,
    text: &str,
    scale: f32,
) -> i32 {
    unsafe {
        let font = text_paint::create_font(
            text_paint::scaled(text_paint::LEVEL_FONT_PIXELS[style.level as usize], scale),
            style,
        );
        if font.0.is_null() {
            return 0;
        }

        let previous = SelectObject(surface.dc, font);
        let wide: Vec<u16> = text.encode_utf16().collect();
        let mut extent = SIZE::default();
        let measured = GetTextExtentPoint32W(surface.dc, &wide, &mut extent).as_bool();
        let _ = SelectObject(surface.dc, previous);
        let _ = windows::Win32::Graphics::Gdi::DeleteObject(font);

        if measured {
            extent.cx
        } else {
            0
        }
    }
}

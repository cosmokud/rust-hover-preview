//! The strip above a pinned preview's media: the bar, the buttons standing against its right
//! edge, and the name of what is pinned.
//!
//! The window's own three - minimize, maximize or restore, close - are where Windows puts them,
//! and the pin's own four - the file before this one, the file after it, the file handed to
//! whatever the machine has filed it under, and the file handed to a program named out loud -
//! are packed to their left. Both groups are measured off the window's three alone (`button_boxes`),
//! so a close button that moved because a pin grew a walk is not a close button in a new place on
//! every caption the moment this app was updated.
//!
//! A strip is kept whole and copied rather than drawn again when it is asked for twice unchanged,
//! because a pinned window repaints sixty times a second to move a playhead that is drawn on the
//! bar rather than on this (`CaptionStrip`). What a button says is not drawn here at all: a tooltip
//! is a panel over the media below the strip, and the window owns where it goes (see `bubble`).

use super::primitives::{
    caption_style, draw_chevron, draw_cross, draw_open_with, draw_open_with_list, fill_box,
    measure_text, stroke_box, surface_pixels, ChromePalette, GLYPH_PIXELS, GLYPH_STROKE_PIXELS,
};
use crate::text::text_paint::{self, DibSurface};
use std::cell::RefCell;
use windows::Win32::Foundation::{RECT, SIZE};
use windows::Win32::Graphics::Gdi::{GetTextExtentPoint32W, SelectObject};

/// How wide a caption button is, in the units a display's scale multiplies: the width
/// Windows 11 gives one, so the three land where a hand expects them.
pub(super) const BUTTON_PIXELS: f32 = 46.0;

const TITLE_PADDING_PIXELS: f32 = 12.0;

/// How wide a run of text is in a caption's face at a display's scale, measured through the
/// same font the run is drawn in.
///
/// It is a question rather than an estimate because the text is a program's name: "Open With
/// Adobe Photoshop" and "Open With Photos" are nowhere near the same width, and a panel sized
/// for one of them clips the other.
pub(crate) fn measure_caption_text(surface: &DibSurface, text: &str, dpi: u32) -> i32 {
    if text.is_empty() {
        return 0;
    }

    measure_text(surface, &caption_style([0, 0, 0]), text, dpi as f32 / 96.0)
}

/// The buttons a caption carries, in the order they sit in: the two or three that are a
/// window's own — minimize, maximize or restore, close — and the three that are a pin's,
/// packed to their left.
///
/// The window's are the ones a hand has been reaching for on every window on the desktop,
/// and they are where Windows puts them. The pin's are new: the file before this one, the
/// file after it, and the file handed to whatever the machine has filed it under (see
/// `shell::pin_navigation`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum CaptionButton {
    /// The file before the one pinned, in the order the folder it was taken up in is
    /// showing them.
    Previous,
    /// The file after the one pinned, the same walk the other way.
    Next,
    /// Hand the pinned file to whatever the Shell has filed it under. A pin previews a file
    /// rather than opening it, so this is the only way out of it into the program that owns
    /// the format.
    OpenWith,
    /// Hand the pinned file to a program the user names out loud, rather than to the one the
    /// machine has settled on. Beside `OpenWith` because the two are the same gesture two
    /// directions: this one opens the Shell's own list of what could open it, which is the only
    /// way into a second program when the default is the wrong one — the button beside it can
    /// only ever reach the one the Shell has already chosen.
    OpenWithList,
    Minimize,
    Maximize,
    Close,
}

/// The pin's own four, in the order they are packed. They are separate from the window's
/// group because a window's group is measured by itself and cannot give any of its room:
/// a close button that moved because a pin grew a walk would be a close button in a new
/// place on every caption the moment this app was updated.
const NAV_BUTTONS: [CaptionButton; 4] = [
    CaptionButton::Previous,
    CaptionButton::Next,
    CaptionButton::OpenWith,
    CaptionButton::OpenWithList,
];

/// What a caption is drawn from: the name of what is pinned, whether it is maximized, whether it
/// offers a maximize at all, and which button the pointer is on, if any.
pub(crate) struct Caption<'a> {
    pub(crate) title: &'a str,
    pub(crate) maximized: bool,
    /// Whether this pin has a maximum to go to. A sound's card does not: the card is its own
    /// size, so there is no larger box to give it and the button would be one that does nothing
    /// (see `pin_frame`).
    pub(crate) maximizable: bool,
    pub(crate) hovered: Option<CaptionButton>,
    pub(crate) pressed: Option<CaptionButton>,
}

/// Where a caption's buttons are, in the caption's own coordinates: each one's box, left to
/// right. They are the caption's own height and sit against its right edge, which is where a
/// hand goes for a window it wants to close — and the same place Windows puts them.
///
/// A caption without a maximize carries two buttons rather than three, and the two it carries
/// keep their own places at that edge: closing a window is a gesture made by position, and a
/// close button that moved because the window it is on has nothing to maximize would be a
/// window whose close button is where a maximize is on every other one.
///
/// The pin's own four — the file before, the file after, the file opened by whatever
/// the machine has filed it under, and the file opened by a program named out loud — are packed
/// beside that group and not inside it, and are dropped whole where a caption is too narrow to
/// carry them. Which of the two is given up is not a question: the walk is a thing a window
/// only has once it is a pin, and closing a pin is not.
pub(crate) fn button_boxes(
    width: i32,
    height: i32,
    dpi: u32,
    maximizable: bool,
) -> Vec<CaptionButtonBox> {
    let kinds: &[CaptionButton] = if maximizable {
        &[
            CaptionButton::Minimize,
            CaptionButton::Maximize,
            CaptionButton::Close,
        ]
    } else {
        &[CaptionButton::Minimize, CaptionButton::Close]
    };

    // Measured off the window's group alone, so the walk beside it can be any width at all and
    // the three land where they landed before it existed.
    let button = text_paint::scaled(BUTTON_PIXELS as i32, dpi as f32 / 96.0)
        .max(1)
        .min((width / kinds.len().max(1) as i32).max(1));

    let mut boxes: Vec<CaptionButtonBox> = kinds
        .iter()
        .enumerate()
        .map(|(index, kind)| {
            let from_the_right = kinds.len() as i32 - 1 - index as i32;
            let right = width - from_the_right * button;
            CaptionButtonBox {
                kind: *kind,
                rect: RECT {
                    left: right - button,
                    top: 0,
                    right,
                    bottom: height,
                },
            }
        })
        .collect();

    // The walk is packed against the group it hangs off, as wide as the buttons beside it: a
    // caption's buttons are one size, and a narrower one in the middle of them is a row of
    // targets a hand has to find rather than count.
    //
    // A caption too narrow for all four keeps the window's group and loses the walk whole.
    // The buttons a hand has been reaching for on every window it has ever had are the ones
    // that are not given up, and half a walk beside them is a set of targets with no known
    // order to them.
    if width - (kinds.len() as i32 + NAV_BUTTONS.len() as i32) * button < 0 {
        return boxes;
    }

    // Packed from the group's own left edge outward, so the walk reads left-to-right in the
    // order it is written: the file before this one, the file after it, the hand-off to
    // another program, and then the hand-off to one named out loud. Walking out from the edge
    // and reversing would put them the other way round, which puts `Next` where a hand reaches
    // for `Previous` — and puts the list beside the default rather than beyond it, which is
    // where a hand that has just tried the default and found it wanting goes next.
    let group_left = width - kinds.len() as i32 * button;
    let mut nav: Vec<CaptionButtonBox> = NAV_BUTTONS
        .iter()
        .enumerate()
        .map(|(index, kind)| {
            let left = group_left - (NAV_BUTTONS.len() as i32 - index as i32) * button;
            CaptionButtonBox {
                kind: *kind,
                rect: RECT {
                    left,
                    top: 0,
                    right: left + button,
                    bottom: height,
                },
            }
        })
        .collect();
    nav.append(&mut boxes);

    nav
}

/// One caption button and the box it occupies.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) struct CaptionButtonBox {
    pub(crate) kind: CaptionButton,
    pub(crate) rect: RECT,
}

/// Which button a point in the caption's own coordinates is on, if any.
pub(crate) fn button_at(
    x: i32,
    y: i32,
    width: i32,
    height: i32,
    dpi: u32,
    maximizable: bool,
) -> Option<CaptionButton> {
    button_boxes(width, height, dpi, maximizable)
        .into_iter()
        .find(|button| {
            x >= button.rect.left
                && x < button.rect.right
                && y >= button.rect.top
                && y < button.rect.bottom
        })
        .map(|button| button.kind)
}

/// Paint a caption into a surface the size of the strip: the bar, the three buttons with
/// whatever state the pointer has put them in, and the name of what is pinned.
pub(crate) fn paint_caption(
    surface: &DibSurface,
    palette: &ChromePalette,
    caption: &Caption,
    dpi: u32,
) {
    let length = surface.width as usize * surface.height as usize * 4;

    // Copied rather than drawn again where the same strip has been asked for before, which is
    // what a repaint is: a pinned window repaints sixty times a second to move a playhead that
    // is drawn on the bar and not on this, and a caption that redrew its own name every one of
    // those frames paid a font create and delete, a handful of `GetTextExtentPoint32W` and a
    // per-pixel walk of every glyph to arrive at the same bytes.
    //
    // The key is every input the drawing below reads, not the ones that seemed to matter: the
    // size and the scale, the four theme colors, the name, and which button the pointer is on
    // and holding down. A key missing the hovered or pressed button would be a caption whose
    // button kept the wash it had when the pointer first crossed it and never gave it back.
    let reused = CAPTION_STRIP.with(|cell| {
        let cached = cell.borrow();
        let Some(strip) = cached.as_ref() else {
            return false;
        };
        if strip.pixels.len() != length || !strip.painted_from(surface, palette, caption, dpi) {
            return false;
        }

        unsafe {
            std::slice::from_raw_parts_mut(surface_pixels(surface), length)
                .copy_from_slice(&strip.pixels);
        }
        true
    });
    if reused {
        return;
    }

    let width = surface.width as i32;
    let height = surface.height as i32;
    let scale = dpi as f32 / 96.0;

    text_paint::fill_rect(
        surface,
        RECT {
            left: 0,
            top: 0,
            right: width,
            bottom: height,
        },
        palette.background,
    );

    // A hairline along the bottom, so the caption reads as a strip over a picture of any
    // color rather than dissolving into one that happens to match it.
    let hairline = palette.hover(0.18);
    text_paint::fill_rect(
        surface,
        RECT {
            left: 0,
            top: height - 1,
            right: width,
            bottom: height,
        },
        hairline,
    );

    let buttons = button_boxes(width, height, dpi, caption.maximizable);
    for button in &buttons {
        paint_button(surface, palette, caption, *button, scale);
    }

    let title_right = buttons
        .iter()
        .map(|button| button.rect.left)
        .min()
        .unwrap_or(width);
    paint_title(surface, palette, caption.title, title_right, scale);

    // The tooltip is not drawn here: it is a panel of its own hanging below the strip, over
    // the media, and is the window's own business to place (see `tooltip_layout`). All this
    // hands over is which button the name belongs to and what it says.
    //
    // GDI leaves the alpha byte of everything it draws at zero, and a caption is opaque wherever
    // it is painted, so the strip's own coverage is handed back here — after the last of the
    // text rather than after each run, because a caption whose title was sealed and whose
    // tooltip was sealed separately would leave the tooltip's run as a hole in the bar (see the
    // module documentation).
    let pixels = unsafe { std::slice::from_raw_parts_mut(surface_pixels(surface), length) };
    for pixel in pixels.as_chunks_mut::<4>().0 {
        pixel[3] = 255;
    }

    // Kept as it stands rather than as the drawing that made it, so what the next repaint of
    // this same strip copies is the bytes the eye has already been shown — including the
    // coverage the seal above has just handed back, which a copy of the strip before it would
    // not have carried.
    let painted = pixels.to_vec();
    CAPTION_STRIP.with(|cell| {
        *cell.borrow_mut() = Some(CaptionStrip {
            width: surface.width,
            height: surface.height,
            dpi,
            title: caption.title.to_string(),
            maximized: caption.maximized,
            maximizable: caption.maximizable,
            hovered: caption.hovered,
            pressed: caption.pressed,
            background: palette.background,
            foreground: palette.foreground,
            accent: palette.accent,
            dark: palette.dark,
            pixels: painted,
        });
    });
}

// One caption strip kept whole, with everything it was drawn from.
//
// Kept per thread because a strip is painted into a surface the preview thread owns and read
// back off that same one: one cache shared between two windows would be handing each of them
// the other's pixels.
thread_local! {
    static CAPTION_STRIP: RefCell<Option<CaptionStrip>> = const { RefCell::new(None) };
}

/// A caption strip, and the whole of what it was drawn from.
struct CaptionStrip {
    width: u32,
    height: u32,
    dpi: u32,
    title: String,
    maximized: bool,
    maximizable: bool,
    hovered: Option<CaptionButton>,
    pressed: Option<CaptionButton>,
    background: [u8; 3],
    foreground: [u8; 3],
    accent: [u8; 3],
    dark: bool,
    pixels: Vec<u8>,
}

impl CaptionStrip {
    /// Whether this strip is the one asked for: every input the drawing reads, so a theme
    /// change, a resized window, another display's scale, another file, or a pointer that has
    /// moved onto — or off — a button all of them redraw rather than copy.
    ///
    /// All four theme colors are compared and not only the two the strip's own pixels are mostly
    /// made of, because the close button is washed with a red that does not come from the theme
    /// at all and every other wash is the background moved toward the text — which branches on
    /// `dark`. A key left off the background would hand a dark theme's caption to a light one,
    /// and one left off the hovered button would leave the wash under a pointer that has since
    /// walked away.
    fn painted_from(
        &self,
        surface: &DibSurface,
        palette: &ChromePalette,
        caption: &Caption,
        dpi: u32,
    ) -> bool {
        self.width == surface.width
            && self.height == surface.height
            && self.dpi == dpi
            && self.title == caption.title
            && self.maximized == caption.maximized
            && self.maximizable == caption.maximizable
            && self.hovered == caption.hovered
            && self.pressed == caption.pressed
            && self.background == palette.background
            && self.foreground == palette.foreground
            && self.accent == palette.accent
            && self.dark == palette.dark
    }
}

fn paint_button(
    surface: &DibSurface,
    palette: &ChromePalette,
    caption: &Caption,
    button: CaptionButtonBox,
    scale: f32,
) {
    let pressed = caption.pressed == Some(button.kind);
    let hovered = caption.hovered == Some(button.kind);

    if pressed || hovered {
        // The close button is washed with the red Windows washes it with, whichever theme is
        // in use: it is the one button whose meaning does not depend on the palette. The
        // other two are washed with the theme's own background moved a step toward its text.
        let wash = if button.kind == CaptionButton::Close {
            if pressed {
                [178, 32, 32]
            } else {
                [196, 43, 28]
            }
        } else if pressed {
            palette.hover(0.18)
        } else {
            palette.hover(0.10)
        };

        text_paint::fill_rect(surface, button.rect, wash);
    }

    let ink = if button.kind == CaptionButton::Close && (pressed || hovered) {
        [255, 255, 255]
    } else {
        palette.foreground
    };

    paint_glyph(surface, button, caption.maximized, ink, scale);
}

pub(super) fn paint_glyph(
    surface: &DibSurface,
    button: CaptionButtonBox,
    maximized: bool,
    ink: [u8; 3],
    scale: f32,
) {
    let buffer = unsafe {
        std::slice::from_raw_parts_mut(
            surface_pixels(surface),
            surface.width as usize * surface.height as usize * 4,
        )
    };
    let width = surface.width as i32;

    let center_x = (button.rect.left + button.rect.right) / 2;
    let center_y = (button.rect.top + button.rect.bottom) / 2;
    let glyph = text_paint::scaled(GLYPH_PIXELS as i32, scale).max(6);
    let stroke = (GLYPH_STROKE_PIXELS * scale).max(1.0);
    let half = glyph / 2;

    match button.kind {
        CaptionButton::Previous => {
            draw_chevron(buffer, width, center_x, center_y, glyph, stroke, ink, false)
        }
        CaptionButton::Next => {
            draw_chevron(buffer, width, center_x, center_y, glyph, stroke, ink, true)
        }
        CaptionButton::OpenWith => {
            draw_open_with(buffer, width, center_x, center_y, glyph, stroke, ink)
        }
        CaptionButton::OpenWithList => {
            draw_open_with_list(buffer, width, center_x, center_y, glyph, stroke, ink)
        }
        CaptionButton::Minimize => {
            let top = center_y;
            fill_box(
                buffer,
                width,
                RECT {
                    left: center_x - half,
                    top,
                    right: center_x + half + 1,
                    bottom: top + stroke.round().max(1.0) as i32,
                },
                ink,
                1.0,
            );
        }
        // The restore glyph: the window behind, then the window in front of it, the way
        // Windows draws the pair. A button that means "restore" is the one on a window that
        // is maximized.
        CaptionButton::Maximize if maximized => {
            let inset = text_paint::scaled(3, scale).max(1);
            stroke_box(
                buffer,
                width,
                RECT {
                    left: center_x - half + inset,
                    top: center_y - half,
                    right: center_x + half + 1,
                    bottom: center_y + half + 1 - inset,
                },
                stroke,
                ink,
                1.0,
            );
            stroke_box(
                buffer,
                width,
                RECT {
                    left: center_x - half,
                    top: center_y - half + inset,
                    right: center_x + half + 1 - inset,
                    bottom: center_y + half + 1,
                },
                stroke,
                ink,
                1.0,
            );
        }
        CaptionButton::Maximize => {
            stroke_box(
                buffer,
                width,
                RECT {
                    left: center_x - half,
                    top: center_y - half,
                    right: center_x + half + 1,
                    bottom: center_y + half + 1,
                },
                stroke,
                ink,
                1.0,
            );
        }
        CaptionButton::Close => draw_cross(buffer, width, center_x, center_y, glyph, stroke, ink),
    }
}

fn paint_title(
    surface: &DibSurface,
    palette: &ChromePalette,
    title: &str,
    available_right: i32,
    scale: f32,
) {
    let padding = text_paint::scaled(TITLE_PADDING_PIXELS as i32, scale);
    let left = padding;
    let right = available_right - padding;
    if right <= left || title.is_empty() {
        return;
    }

    let style = caption_style(palette.foreground);

    let fitted = fit_title(surface, &style, title, right - left, scale);
    if fitted.is_empty() {
        return;
    }

    // ExtTextOut draws from the top of a text cell, so the title is centered by giving the
    // run a box the height of the cell rather than the whole caption.
    let cell = text_paint::scaled(
        text_paint::LEVEL_FONT_PIXELS[text_paint::BODY_LEVEL as usize],
        scale,
    );
    let top = ((surface.height as i32 - cell) / 2).max(0);
    let rect = RECT {
        left,
        top,
        right,
        bottom: (top + cell).min(surface.height as i32),
    };

    let mut painter = text_paint::RunPainter::new(surface, scale);
    painter.draw(
        &fitted,
        left + 1,
        rect,
        &style,
        palette.foreground,
        palette.background,
    );
}

/// The longest beginning of `title` that fits `budget` pixels, with an ellipsis after it.
///
/// Searched rather than walked. A prefix only ever gets wider, so the longest beginning that
/// fits is found in about seven measurements whatever the length of the name — where walking it
/// was one `GetTextExtentPoint32W` per character, over a deep folder path, on every repaint
/// that only moved a playhead.
///
/// It asks a `width` rather than a font, which is what let it be tested at all: the walk it
/// replaced and the search that replaced it had each been written out a second time in the
/// tests, so the tests agreed with the search and said nothing whatever about the one the
/// caption is actually drawn from. One definition, two callers, is the only shape in which
/// breaking this one is visible.
///
/// A name whose first character is already too wide for the ellipsis is not a name with room
/// cut off it, so nothing comes back rather than an ellipsis alone.
pub(super) fn longest_prefix_that_fits(
    title: &str,
    budget: i32,
    width: &mut dyn FnMut(&str) -> i32,
) -> String {
    let boundaries: Vec<usize> = title
        .char_indices()
        .skip(1)
        .map(|(index, _)| index)
        .collect();

    // How many of the boundaries fit, which is the one the walk used to arrive at by stopping at
    // the first that did not.
    let mut low = 0usize;
    let mut high = boundaries.len();
    while low < high {
        let middle = low + (high - low) / 2;
        if width(&title[..boundaries[middle]]) <= budget {
            low = middle + 1;
        } else {
            high = middle;
        }
    }

    // The last boundary is the one before the final character: a name cut at the very last
    // character is a name with a character missing, and it is cut where the width ran out
    // rather than as late as it could be.
    match low.checked_sub(1).map(|index| boundaries[index]) {
        Some(best) => format!("{}…", &title[..best]),
        None => String::new(),
    }
}

/// The longest beginning of `title` that fits `available` pixels, with an ellipsis after it
/// where any of the name had to be left out. Measured rather than guessed: a caption is the
/// file's own name, and a name cut by a character count is either a name with room to spare
/// or one whose end was cut off twice.
pub(super) fn fit_title(
    surface: &DibSurface,
    style: &text_paint::TextStyle,
    title: &str,
    available: i32,
    scale: f32,
) -> String {
    if available <= 0 {
        return String::new();
    }

    unsafe {
        let font = text_paint::create_font(
            text_paint::scaled(text_paint::LEVEL_FONT_PIXELS[style.level as usize], scale),
            style,
        );
        if font.0.is_null() {
            return title.to_string();
        }

        let previous = SelectObject(surface.dc, font);

        // One buffer for every measurement below, filled in place. A caption is repainted on
        // every pointer move across it, and the alternative was an allocation per measurement.
        let mut wide: Vec<u16> = Vec::new();
        let measured = |text: &str, wide: &mut Vec<u16>| -> i32 {
            wide.clear();
            wide.extend(text.encode_utf16());
            if wide.is_empty() {
                return 0;
            }
            let mut extent = SIZE::default();
            if GetTextExtentPoint32W(surface.dc, wide, &mut extent).as_bool() {
                extent.cx
            } else {
                0
            }
        };

        let full = measured(title, &mut wide);
        let fitted = if full <= available {
            title.to_string()
        } else {
            let ellipsis = measured("…", &mut wide);
            let budget = (available - ellipsis).max(0);
            longest_prefix_that_fits(title, budget, &mut |text| measured(text, &mut wide))
        };

        let _ = SelectObject(surface.dc, previous);
        let _ = windows::Win32::Graphics::Gdi::DeleteObject(font);

        fitted
    }
}

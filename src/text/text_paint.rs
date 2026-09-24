//! The GDI painting layer the previews that draw their own page share.
//!
//! A text preview and an archive listing are the same kind of thing to the
//! renderer: a page made of runs of text drawn on the theme's own background,
//! painted into a top-down 32-bit DIB section and handed back as a BGRA frame —
//! the frame shape a decoded image arrives in, so everything downstream (the
//! layered surface, the spinner, the hover generation check) is unchanged.
//!
//! What lives here is exactly that: the surface, the font metrics and the fonts
//! grown from them, the color helpers the themes are read through, and the two
//! functions that put text and rectangles on the surface. What a page *is* —
//! documents, windows and selections for text, a tree for an archive — stays
//! with the module that owns it.

use crate::config::config::sanitize_text_font_scale_percent;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::sync::Mutex;
use syntect::highlighting::{Color, FontStyle, Style};
use windows::core::PCWSTR;
use windows::Win32::Foundation::{COLORREF, RECT, SIZE};
use windows::Win32::Graphics::Gdi::{
    CreateCompatibleDC, CreateDIBSection, CreateFontW, DeleteDC, DeleteObject, ExtTextOutW,
    GetTextExtentPoint32W, GetTextMetricsW, SelectObject, SetBkColor, SetBkMode, SetTextColor,
    BITMAPINFO, BITMAPINFOHEADER, BI_RGB, CLEARTYPE_QUALITY, CLIP_DEFAULT_PRECIS, DEFAULT_CHARSET,
    DIB_RGB_COLORS, ETO_CLIPPED, ETO_OPAQUE, FF_MODERN, FIXED_PITCH, HBITMAP, HDC, HFONT, HGDIOBJ,
    OPAQUE, OUT_TT_PRECIS, TEXTMETRICW,
};

/// Size steps the document model uses: body text, four heading levels.
pub(crate) const SIZE_LEVELS: usize = 5;
pub(crate) const LEVEL_FONT_PIXELS: [i32; SIZE_LEVELS] = [13, 20, 17, 15, 14];
pub(crate) const LEVEL_EXTRA_LEADING: [i32; SIZE_LEVELS] = [0, 8, 6, 3, 0];
pub(crate) const BODY_LEVEL: u8 = 0;

pub(crate) const PADDING_PIXELS: f32 = 12.0;
pub(crate) const QUOTE_BAR_PIXELS: i32 = 4;

/// The scrollbar drawn when a document is longer than the frame: a thin groove
/// near the right edge, and the room kept clear for it.
pub(crate) const SCROLLBAR_WIDTH_PIXELS: i32 = 6;
pub(crate) const SCROLLBAR_MARGIN_PIXELS: i32 = 4;

/// Consolas ships with Windows, is fixed pitch, and carries the box-drawing
/// characters NFO art is made of; bold and italic keep the same advance, which
/// is what lets a layout measure a line by counting characters.
pub(crate) const FONT_FACE: &str = "Consolas";

pub(crate) const MIN_DPI: u32 = 48;
pub(crate) const MAX_DPI: u32 = 480;

/// Bounds on the combined display and font scale, so a hand-edited font size
/// cannot ask for glyphs larger than a screen or smaller than a pixel.
pub(crate) const MIN_SCALE: f32 = 0.25;
pub(crate) const MAX_SCALE: f32 = 16.0;

pub(crate) fn rgb(color: Color) -> [u8; 3] {
    [color.r, color.g, color.b]
}

pub(crate) fn luminance(color: [u8; 3]) -> f32 {
    (0.2126 * color[0] as f32 + 0.7152 * color[1] as f32 + 0.0722 * color[2] as f32) / 255.0
}

pub(crate) fn blend(base: [u8; 3], other: [u8; 3], amount: f32) -> [u8; 3] {
    let mix = |a: u8, b: u8| (a as f32 + (b as f32 - a as f32) * amount).round() as u8;
    [
        mix(base[0], other[0]),
        mix(base[1], other[1]),
        mix(base[2], other[2]),
    ]
}

/// Pull a color away from the page until it can be read on it.
pub(crate) fn readable(color: [u8; 3], background: [u8; 3]) -> [u8; 3] {
    let on_light_page = luminance(background) > 0.5;
    let limit = 0.45;
    let target = if on_light_page {
        [0, 0, 0]
    } else {
        [255, 255, 255]
    };

    let mut result = color;
    for _ in 0..8 {
        let difference = (luminance(result) - luminance(background)).abs();
        if difference >= limit {
            break;
        }
        result = blend(result, target, 0.25);
    }

    result
}

/// How one run of text is drawn: its color, its decorations, its size level and
/// the face it is set in.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct TextStyle {
    pub(crate) foreground: [u8; 3],
    pub(crate) background: Option<[u8; 3]>,
    pub(crate) bold: bool,
    pub(crate) italic: bool,
    pub(crate) underline: bool,
    pub(crate) strike: bool,
    pub(crate) level: u8,
    pub(crate) face: &'static str,
}

pub(crate) fn plain_style(level: u8) -> TextStyle {
    TextStyle {
        foreground: [0, 0, 0],
        background: None,
        bold: false,
        italic: false,
        underline: false,
        strike: false,
        level,
        face: FONT_FACE,
    }
}

pub(crate) fn text_style(style: &Style, level: u8) -> TextStyle {
    TextStyle {
        foreground: rgb(style.foreground),
        background: None,
        bold: style.font_style.contains(FontStyle::BOLD),
        italic: style.font_style.contains(FontStyle::ITALIC),
        underline: style.font_style.contains(FontStyle::UNDERLINE),
        strike: false,
        level,
        face: FONT_FACE,
    }
}

pub(crate) fn scaled(pixels: i32, scale: f32) -> i32 {
    ((pixels as f32) * scale).round().max(1.0) as i32
}

// -------------------------------------------------------------------- metrics

/// Advance and line height for each size level at one scale. Answering costs a
/// font created and deleted per level, and the answer depends on nothing but the
/// scale, so a display scale and font size are measured once and kept.
const LEVEL_METRICS_MAX_ENTRIES: usize = 16;

#[derive(Clone, Copy)]
struct LevelMetrics {
    advance: [i32; SIZE_LEVELS],
    line_height: [i32; SIZE_LEVELS],
}

static LEVEL_METRICS: Lazy<Mutex<HashMap<(u32, u32), LevelMetrics>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// Font metrics for one display scale and font scale, taken from GDI once per
/// scale and font size.
pub(crate) struct TextMetrics {
    pub(crate) scale: f32,
    pub(crate) padding: i32,
    pub(crate) quote_bar: i32,
    pub(crate) advance: [i32; SIZE_LEVELS],
    pub(crate) line_height: [i32; SIZE_LEVELS],
}

impl TextMetrics {
    /// `dpi` is the display's own scale and `font_scale_percent` the configured
    /// text size on top of it. They multiply into one scale, so the glyphs, the
    /// line spacing and the page margin all grow together and a preview at 200%
    /// is the same page twice the size rather than the same page in a bigger box.
    pub(crate) fn new(dc: HDC, dpi: u32, font_scale_percent: u32) -> Option<Self> {
        let dpi = dpi.clamp(MIN_DPI, MAX_DPI);
        let font_scale_percent = sanitize_text_font_scale_percent(font_scale_percent);
        let font_scale = font_scale_percent as f32 / 100.0;
        // A preview is still a preview: past this the config is asking for a
        // handful of characters per screen, which GDI font sizes stop being
        // useful for.
        let scale = (dpi as f32 / 96.0 * font_scale).clamp(MIN_SCALE, MAX_SCALE);

        let levels = level_metrics(dc, (dpi, font_scale_percent), scale)?;

        Some(Self {
            scale,
            padding: scaled(PADDING_PIXELS as i32, scale),
            quote_bar: scaled(QUOTE_BAR_PIXELS, scale),
            advance: levels.advance,
            line_height: levels.line_height,
        })
    }

    pub(crate) fn indent(&self, characters: u8) -> i32 {
        characters as i32 * self.advance[BODY_LEVEL as usize]
    }

    /// Room the scrollbar and its margin take from the right edge of the text.
    pub(crate) fn scrollbar_space(&self) -> i32 {
        scaled(SCROLLBAR_WIDTH_PIXELS + SCROLLBAR_MARGIN_PIXELS, self.scale)
    }

    pub(crate) fn scrollbar_width(&self) -> i32 {
        scaled(SCROLLBAR_WIDTH_PIXELS, self.scale).max(1)
    }
}

/// The five levels at `scale`, from the cache when this display scale and font
/// size have been measured before.
fn level_metrics(dc: HDC, key: (u32, u32), scale: f32) -> Option<LevelMetrics> {
    if let Ok(cache) = LEVEL_METRICS.lock() {
        if let Some(cached) = cache.get(&key) {
            return Some(*cached);
        }
    }

    let measured = measure_levels(dc, scale)?;

    if let Ok(mut cache) = LEVEL_METRICS.lock() {
        if !cache.contains_key(&key) && cache.len() >= LEVEL_METRICS_MAX_ENTRIES {
            cache.clear();
        }
        cache.insert(key, measured);
    }

    Some(measured)
}

/// Measure the five size levels against `dc`, leaving no font behind: each level
/// is created, asked for its advance and line height, and deleted again.
fn measure_levels(dc: HDC, scale: f32) -> Option<LevelMetrics> {
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

    Some(LevelMetrics {
        advance,
        line_height,
    })
}

// ------------------------------------------------------------------ painting

/// A memory DC with a top-down 32-bit DIB section selected into it: the surface
/// GDI draws a preview page on before it is read back as a frame.
pub(crate) struct DibSurface {
    pub(crate) dc: HDC,
    bitmap: HBITMAP,
    previous_bitmap: HGDIOBJ,
    bits: *mut u8,
    pub(crate) width: u32,
    pub(crate) height: u32,
}

impl DibSurface {
    pub(crate) fn create(width: u32, height: u32) -> Option<Self> {
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
    /// byte is forced opaque here: a page is a page, and the layers above it
    /// composite it as one.
    pub(crate) fn pixels(&self) -> Vec<u8> {
        let length = self.width as usize * self.height as usize * 4;
        let mut pixels = unsafe { std::slice::from_raw_parts(self.bits, length) }.to_vec();
        for pixel in pixels.as_chunks_mut::<4>().0 {
            pixel[3] = 255;
        }
        pixels
    }

    pub(crate) fn bits(&self) -> *mut u8 {
        self.bits
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
pub(crate) struct FontCache {
    fonts: Vec<(TextStyleKey, HFONT)>,
}

#[derive(PartialEq, Eq)]
pub(crate) struct TextStyleKey {
    face: &'static str,
    level: u8,
    bold: bool,
    italic: bool,
    underline: bool,
    strike: bool,
}

impl TextStyleKey {
    pub(crate) fn of(style: &TextStyle) -> Self {
        Self {
            face: style.face,
            level: (style.level as usize).min(SIZE_LEVELS - 1) as u8,
            bold: style.bold,
            italic: style.italic,
            underline: style.underline,
            strike: style.strike,
        }
    }
}

impl FontCache {
    pub(crate) fn new() -> Self {
        Self { fonts: Vec::new() }
    }

    pub(crate) unsafe fn get(&mut self, style: &TextStyle, scale: f32) -> HFONT {
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

pub(crate) fn create_font(pixels: i32, style: &TextStyle) -> HFONT {
    let face: Vec<u16> = style
        .face
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();

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

/// Draw one piece of a run into `rect`, filling that rectangle with its
/// background as it goes, so a piece of a selected run carries the highlight
/// instead of the page. The text is drawn at the rectangle's own corner and
/// clipped to it: a run's box is the room it was given, and nothing of it has any
/// business outside.
pub(crate) unsafe fn paint_run(
    surface: &DibSurface,
    text: &str,
    rect: RECT,
    foreground: [u8; 3],
    background: [u8; 3],
) {
    if text.is_empty() {
        return;
    }

    let wide: Vec<u16> = text.encode_utf16().collect();
    SetTextColor(surface.dc, colorref(foreground));
    SetBkColor(surface.dc, colorref(background));
    SetBkMode(surface.dc, OPAQUE);

    let _ = ExtTextOutW(
        surface.dc,
        rect.left,
        rect.top,
        ETO_OPAQUE | ETO_CLIPPED,
        Some(&rect),
        PCWSTR(wide.as_ptr()),
        wide.len() as u32,
        None,
    );
}

/// Draws runs onto one surface, selecting each run's font as the style changes
/// and putting the surface's own object back when it is dropped — so the fonts
/// can be deleted without a DC still using one.
pub(crate) struct RunPainter<'a> {
    surface: &'a DibSurface,
    fonts: FontCache,
    previous_font: Option<HGDIOBJ>,
    selected: Option<TextStyleKey>,
    scale: f32,
}

impl<'a> RunPainter<'a> {
    pub(crate) fn new(surface: &'a DibSurface, scale: f32) -> Self {
        Self {
            surface,
            fonts: FontCache::new(),
            previous_font: None,
            selected: None,
            scale,
        }
    }

    /// Draw a run of text into `rect`, with its own background, which is also
    /// what fills the part of the rectangle the glyphs do not reach.
    pub(crate) fn draw(
        &mut self,
        text: &str,
        rect: RECT,
        style: &TextStyle,
        foreground: [u8; 3],
        background: [u8; 3],
    ) {
        if text.is_empty() {
            return;
        }

        let key = TextStyleKey::of(style);
        if self.selected.as_ref() != Some(&key) {
            unsafe {
                let font = self.fonts.get(style, self.scale);
                let previous = SelectObject(self.surface.dc, font);
                if self.previous_font.is_none() {
                    self.previous_font = Some(previous);
                }
            }
            self.selected = Some(key);
        }

        unsafe {
            paint_run(self.surface, text, rect, foreground, background);
        }
    }
}

impl Drop for RunPainter<'_> {
    fn drop(&mut self) {
        if let Some(previous) = self.previous_font.take() {
            unsafe {
                let _ = SelectObject(self.surface.dc, previous);
            }
        }
    }
}

pub(crate) fn fill_rect(surface: &DibSurface, rect: RECT, color: [u8; 3]) {
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
        std::slice::from_raw_parts_mut(surface.bits(), width as usize * height as usize * 4)
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

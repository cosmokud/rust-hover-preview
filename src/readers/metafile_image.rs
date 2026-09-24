//! Windows metafiles, played by the drawing layer the system already has.
//!
//! A `.wmf` and an `.emf` are not pictures: they are lists of drawing records, and what
//! turns one into pixels is a device context that plays them back. Windows ships that
//! player, so nothing is rasterized on this side and no decoder is carried for the
//! format: `PlayEnhMetaFile` draws the records into any device context, and a memory
//! context holding one of this app's own DIB sections is one. The drawing is therefore
//! replayed at the size the preview is shown at, which is what makes a metafile sharp at
//! any size the display has rather than only at its own.
//!
//! This is also how the metafile an encapsulated PostScript file carries is drawn: a
//! Windows `.eps` keeps one in place of a picture, and what it keeps is the same records
//! (see `eps_image`), handed here as bytes rather than as a file.
//!
//! Two plays are made for every frame, over a black page and over a white one, because a
//! metafile says nothing about what it did not paint: what the two plays agree on is what
//! the drawing put there, and what they differ by is how much of the page it left alone.
//! That is what recovers the coverage a device context never writes — the drawing layer
//! has no alpha to give — and it is what lets a preview of a drawing stand on whatever
//! the tray says is behind it, exactly as a picture's transparency does.
//!
//! The records of a metafile are data this app hands to the drawing layer rather than
//! code it runs, and the box they are played into is the one the layout planned, which is
//! what keeps a file from asking for a canvas of its own choosing: a drawing is drawn at
//! the size of the preview and never at the size it names.

use crate::config::config::{frame_bytes_within_budget, read_within_budget};
use std::path::Path;
use std::ptr;
use windows::Win32::Foundation::RECT;
use windows::Win32::Graphics::Gdi::{
    CreateCompatibleDC, CreateDIBSection, DeleteDC, DeleteEnhMetaFile, DeleteObject,
    GetEnhMetaFileHeader, PlayEnhMetaFile, SelectObject, SetBkMode, SetEnhMetaFileBits, BITMAPINFO,
    BITMAPINFOHEADER, BI_RGB, DIB_RGB_COLORS, ENHMETAHEADER, HENHMETAFILE, HGDIOBJ, MM_ANISOTROPIC,
    TRANSPARENT,
};
use windows::Win32::System::DataExchange::{SetWinMetaFileBits, METAFILEPICT};

/// What a placeable metafile opens with: the key, the drawing's bounds in twips, and how
/// many twips its writer counted to the inch. It is the only place an old-style `.wmf`
/// says how large its drawing is — the metafile records themselves say nothing about the
/// page they are on — and a file that was written without it is measured by the drawing
/// layer instead, which works the bounds out from the records.
const PLACEABLE_KEY: u32 = 0x9AC6_CDD7;
const PLACEABLE_BYTES: usize = 22;
const DEFAULT_TWIPS: u32 = 1440;
/// The inches a DIP is, which is what a preview is measured in.
const DIP_DPI: u32 = 96;
/// What an enhanced metafile opens with, which is the same window on both ends of it.
const ENHANCED_HEADER: usize = 88;

/// Whether the file is named as one of Windows' two metafiles.
pub fn is_metafile_name(path: &Path) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.eq_ignore_ascii_case("wmf") || ext.eq_ignore_ascii_case("emf"))
        .unwrap_or(false)
}

/// The size of a drawing that is already in memory, which is where an `.eps` keeps one.
pub fn records_dimensions(records: &[u8]) -> Option<(u32, u32)> {
    measure(records)
}

/// Whether those bytes open as one of the two metafiles.
///
/// A preview inside a container or a comment block carries no label saying which of the
/// two kinds it is, so the picture itself is asked: an enhanced metafile says so twice
/// over, and the older form opens with the header every one of them begins with.
pub fn is_metafile_records(records: &[u8]) -> bool {
    if is_enhanced(records) {
        return true;
    }
    if Placeable::read(records).is_some() {
        return true;
    }

    // A type of zero or one and a header of nine words, which is what a picture will not
    // begin with whatever it is a picture of.
    matches!(records.get(..4), Some([0 | 1, 0x00, 0x09, 0x00]))
}

/// The size the drawing asks to be shown at.
pub fn dimensions(path: &Path) -> Option<(u32, u32)> {
    let bytes = read_within_budget(path)?;

    measure(&bytes)
}

/// The drawing played into a frame of this box, as BGRA with the coverage the plays
/// recover.
pub fn decode(path: &Path, width: u32, height: u32) -> Option<Vec<u8>> {
    let bytes = read_within_budget(path)?;

    draw_records(&bytes, width, height)
}

/// The same for a drawing that is already in memory, which is where an `.eps` keeps one.
pub fn draw_records(records: &[u8], width: u32, height: u32) -> Option<Vec<u8>> {
    if width == 0 || height == 0 {
        return None;
    }
    frame_bytes_within_budget(width, height, 4)?;

    let drawing = Drawing::open(records)?;

    drawing.play(width, height)
}

/// The size of the drawing those bytes hold: what its writer declared, where it declared
/// it, and otherwise what the drawing layer worked out from the records.
fn measure(records: &[u8]) -> Option<(u32, u32)> {
    if let Some(placeable) = Placeable::read(records) {
        return Some(placeable.size());
    }

    let drawing = Drawing::open(records)?;
    drawing.bounds()
}

/// What a placeable metafile declares before its records begin: the drawing's bounds in
/// the writer's own units, and how many of them it counted to the inch.
#[derive(Clone, Copy)]
struct Placeable {
    width: i32,
    height: i32,
    twips: u32,
}

impl Placeable {
    fn read(records: &[u8]) -> Option<Self> {
        let header = records.get(..PLACEABLE_BYTES)?;
        if u32::from_le_bytes([header[0], header[1], header[2], header[3]]) != PLACEABLE_KEY {
            return None;
        }

        let corner = |at: usize| i16::from_le_bytes([header[at], header[at + 1]]) as i32;
        let twips = u32::from_le_bytes([header[14], header[15], header[16], header[17]]);

        Some(Self {
            width: (corner(10) - corner(6)).abs().max(1),
            height: (corner(12) - corner(8)).abs().max(1),
            twips: if twips == 0 { DEFAULT_TWIPS } else { twips },
        })
    }

    /// The same bounds in the DIPs a preview is measured in.
    fn size(&self) -> (u32, u32) {
        let to_pixels =
            |units: i32| (units as i64 * DIP_DPI as i64 / self.twips.max(1) as i64).max(1) as u32;

        (to_pixels(self.width), to_pixels(self.height))
    }
}

/// One drawing, opened from the bytes a file or a container handed over.
///
/// A `.wmf` is not what the drawing layer plays: the enhanced form is, and Windows
/// converts one into the other, which is also what works the bounds out for an old file
/// that never declared them. A placeable header is stripped first, because only the
/// records behind it are a metafile, and what it declared is passed along as the mapping
/// the drawing is to be read with — the header's own answer to how much paper it wants.
struct Drawing {
    handle: HENHMETAFILE,
}

impl Drawing {
    fn open(records: &[u8]) -> Option<Self> {
        let placeable = Placeable::read(records);

        let handle = if is_enhanced(records) {
            unsafe { SetEnhMetaFileBits(records) }
        } else {
            // What the header declared is passed along as the mapping the drawing is to
            // be read with — the writer's own answer to how much paper it wants — and
            // only the records behind it are handed over, because a metafile is what
            // follows the header rather than what includes it.
            let mapping = placeable.map(|placeable| METAFILEPICT {
                mm: MM_ANISOTROPIC.0,
                xExt: placeable.width,
                yExt: placeable.height,
                ..Default::default()
            });
            let records = match placeable {
                Some(_) => records.get(PLACEABLE_BYTES..)?,
                None => records,
            };

            unsafe { SetWinMetaFileBits(records, None, mapping.as_ref().map(|m| m as *const _)) }
        };

        if handle.0.is_null() {
            return None;
        }

        Some(Self { handle })
    }

    /// The drawing's own bounds in the units its reference device used, which is what the
    /// layer reports for a file that declared nothing.
    fn bounds(&self) -> Option<(u32, u32)> {
        let mut header = ENHMETAHEADER::default();
        let size = unsafe {
            GetEnhMetaFileHeader(
                self.handle,
                std::mem::size_of::<ENHMETAHEADER>() as u32,
                Some(&mut header),
            )
        };

        if size == 0 {
            return None;
        }

        let width = header.rclBounds.right - header.rclBounds.left + 1;
        let height = header.rclBounds.bottom - header.rclBounds.top + 1;
        if width <= 0 || height <= 0 {
            return None;
        }

        Some((width as u32, height as u32))
    }

    /// The drawing played twice, once over black and once over white, and the frames
    /// compared for what it painted.
    fn play(&self, width: u32, height: u32) -> Option<Vec<u8>> {
        let dark = Surface::create(width, height, 0x00)?;
        let light = Surface::create(width, height, 0xff)?;

        dark.play(self.handle);
        light.play(self.handle);

        Some(recover(dark.pixels(), light.pixels()))
    }
}

impl Drop for Drawing {
    fn drop(&mut self) {
        unsafe {
            let _ = DeleteEnhMetaFile(self.handle);
        }
    }
}

/// Whether those bytes are an enhanced metafile rather than an old-style one.
///
/// The two places one says so are the record it opens with — a header and nothing else —
/// and the signature it spells out inside that header, which is the bytes of the word
/// `EMF` behind a space.
fn is_enhanced(records: &[u8]) -> bool {
    if records.len() < ENHANCED_HEADER {
        return false;
    }

    let word = |at: usize| {
        u32::from_le_bytes([
            records[at],
            records[at + 1],
            records[at + 2],
            records[at + 3],
        ])
    };

    word(0) == 1 && word(40) == 0x464D_4520
}

/// The colours the two plays agree on and the coverage they differ by.
///
/// A pixel the drawing painted is the same colour over either page, so the two plays agree
/// there and the pixel is opaque. A pixel it left alone is the page itself, so the two
/// plays are a whole page apart — and what lies between them is how much of the page the
/// drawing covered, which is the alpha a device context never had to give. The colour is
/// then the one the white page was painted with, taken back out of the page it was mixed
/// with.
fn recover(dark: &[u8], light: &[u8]) -> Vec<u8> {
    let mut frame = Vec::with_capacity(dark.len());

    for (dark, light) in dark
        .as_chunks::<4>()
        .0
        .iter()
        .zip(light.as_chunks::<4>().0.iter())
    {
        let spread: i32 = (0..3)
            .map(|channel| 255 - (light[channel] as i32 - dark[channel] as i32))
            .sum::<i32>();
        let alpha = (spread / 3).clamp(0, 255) as u8;

        let colour = |channel: usize| -> u8 {
            if alpha == 0 {
                return 0;
            }

            let mixed = light[channel] as i32 - (255 - alpha as i32);
            (mixed * 255 / alpha as i32).clamp(0, 255) as u8
        };

        frame.extend_from_slice(&[colour(0), colour(1), colour(2), alpha]);
    }

    frame
}

/// A page of the drawing layer's making: a memory context with a top-down 32-bit DIB
/// section selected, whose bits are the frame it is drawn into.
struct Surface {
    dc: windows::Win32::Graphics::Gdi::HDC,
    bitmap: windows::Win32::Graphics::Gdi::HBITMAP,
    previous: HGDIOBJ,
    bits: *mut u8,
    width: u32,
    height: u32,
}

impl Surface {
    /// A page of this size, washed in one value before anything is drawn on it: black and
    /// white are the two the coverage is read from.
    fn create(width: u32, height: u32, wash: u8) -> Option<Self> {
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

        let mut bits: *mut core::ffi::c_void = ptr::null_mut();
        let Ok(bitmap) =
            (unsafe { CreateDIBSection(dc, &info, DIB_RGB_COLORS, &mut bits, None, 0) })
        else {
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

        let previous = unsafe { SelectObject(dc, bitmap) };
        let surface = Self {
            dc,
            bitmap,
            previous,
            bits: bits as *mut u8,
            width,
            height,
        };

        surface.wash(wash);

        Some(surface)
    }

    /// The page before anything is drawn on it, with the alpha a frame is read as.
    fn wash(&self, value: u8) {
        let count = self.width as usize * self.height as usize;
        let pixels = unsafe { std::slice::from_raw_parts_mut(self.bits, count * 4) };

        for pixel in pixels.as_chunks_mut::<4>().0 {
            pixel.copy_from_slice(&[value, value, value, 0xff]);
        }
    }

    /// The records played onto the page, at the size of the page: what is drawn is the
    /// whole of the drawing inside this box rather than a corner of it.
    fn play(&self, drawing: HENHMETAFILE) {
        let target = RECT {
            left: 0,
            top: 0,
            right: self.width as i32,
            bottom: self.height as i32,
        };

        unsafe {
            let _ = SetBkMode(self.dc, TRANSPARENT);
            let _ = PlayEnhMetaFile(self.dc, drawing, &target as *const RECT);
        }
    }

    fn pixels(&self) -> &[u8] {
        let count = self.width as usize * self.height as usize;

        unsafe { std::slice::from_raw_parts(self.bits, count * 4) }
    }
}

impl Drop for Surface {
    fn drop(&mut self) {
        unsafe {
            let _ = SelectObject(self.dc, self.previous);
            let _ = DeleteObject(self.bitmap);
            let _ = DeleteDC(self.dc);
        }
    }
}

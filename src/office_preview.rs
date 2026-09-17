//! Measuring and drawing a page of an Office document.
//!
//! Two sources can be drawn from, and the better one wins: the page Excel, Word
//! or PowerPoint rendered in the background (see `office_render`), and the
//! picture the document saved inside itself (see `office_thumbnail`). Which of
//! them a preview will be drawn from decides how large it may be — a rendered
//! page and a saved metafile are vector and are drawn at whatever size they are
//! asked for, while a raster thumbnail is only as good as the pixels it saved.
//!
//! The metafiles are drawn with GDI, the same engine the preview window paints
//! with: `PlayEnhMetaFile` maps a metafile's own frame onto the rectangle it is
//! given, so one call is the whole of the scaling.

use crate::office_render::{self, CachedRender, RenderedKind};
use crate::office_thumbnail::{self, Thumbnail, ThumbnailKind};
use crate::pdf_preview;
use crate::text_paint::DibSurface;
use std::ffi::c_void;
use std::path::Path;
use std::sync::atomic::AtomicBool;

use windows::Win32::Foundation::{COLORREF, RECT};
use windows::Win32::Graphics::Gdi::{
    CreateSolidBrush, DeleteEnhMetaFile, DeleteObject, FillRect, PlayEnhMetaFile,
    SetEnhMetaFileBits, SetStretchBltMode, StretchDIBits, BITMAPINFO, COLORONCOLOR, DIB_RGB_COLORS,
    HGDIOBJ, HMETAFILE, MM_ANISOTROPIC, SRCCOPY,
};
use windows::Win32::System::DataExchange::{SetWinMetaFileBits, METAFILEPICT};

/// Which of a document's sources a preview would be drawn from.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum SourceKind {
    /// A page Office rendered, or a page that is on its way.
    Page,
    /// A metafile the document saved: vector, like a page.
    Metafile,
    /// A raster picture the document saved: only as good as its pixels.
    Raster,
    /// Nothing to draw from yet.
    None,
}

impl SourceKind {
    /// Whether a preview drawn from this source may be enlarged to fill the
    /// display. A raster thumbnail may not: enlarging it would only stretch the
    /// pixels it saved, which is why it follows the configured scale instead —
    /// the rule an image follows.
    pub(crate) fn may_be_enlarged(self) -> bool {
        !matches!(self, Self::Raster)
    }
}

/// Which of the document's sources a preview would be drawn from, asked by the
/// layout before it decides how large the preview may be.
pub(crate) fn source_kind(path: &Path) -> SourceKind {
    if office_render::cached_render(path).is_some() {
        return SourceKind::Page;
    }

    match office_thumbnail::thumbnail_for(path, None) {
        Some(thumbnail) if thumbnail.is_metafile() => SourceKind::Metafile,
        Some(_) => SourceKind::Raster,
        None => SourceKind::None,
    }
}

/// The size the layout places a preview of this document from.
///
/// `None` means there is nothing to show and nothing on the way, which drops the
/// preview — a document with no saved picture whose kind has no engine, or the
/// render tier switched off, is answered with no preview rather than with a box
/// that would spin forever.
pub(crate) fn measure(path: &Path) -> Option<(u32, u32)> {
    if let Some(cached) = office_render::cached_render(path) {
        if let Some(dimensions) = rendered_dimensions(&cached) {
            return Some(dimensions);
        }
    }

    if let Some(thumbnail) = office_thumbnail::thumbnail_for(path, None) {
        return Some((thumbnail.width, thumbnail.height));
    }

    // Nothing saved and nothing rendered yet: a page is what is coming, so the
    // preview is placed for a page's box and the spinner fills it.
    if office_render::enabled() {
        return Some((
            pdf_preview::DEFAULT_PAGE_WIDTH,
            pdf_preview::DEFAULT_PAGE_HEIGHT,
        ));
    }

    None
}

/// The rendered page, or the saved picture, drawn into the box the layout
/// planned — preserving the source's own aspect ratio inside it.
pub(crate) fn render(
    path: &Path,
    target_width: u32,
    target_height: u32,
    cancel: Option<&AtomicBool>,
) -> Option<(Vec<u8>, u32, u32)> {
    if let Some(cached) = office_render::cached_render(path) {
        if let Some(frame) = render_cached(&cached, target_width, target_height) {
            return Some(frame);
        }
    }

    let thumbnail = office_thumbnail::thumbnail_for(path, cancel)?;
    render_thumbnail(&thumbnail, target_width, target_height)
}

fn rendered_dimensions(cached: &CachedRender) -> Option<(u32, u32)> {
    match cached.kind {
        RenderedKind::Pdf => pdf_preview::page_dimensions(&cached.path),
        RenderedKind::Png => image::image_dimensions(&cached.path).ok(),
    }
}

fn render_cached(
    cached: &CachedRender,
    target_width: u32,
    target_height: u32,
) -> Option<(Vec<u8>, u32, u32)> {
    match cached.kind {
        // The page is rendered at the size it is shown at rather than enlarged
        // afterwards, which is what keeps a scaled-up preview sharp.
        RenderedKind::Pdf => {
            pdf_preview::render_first_page(&cached.path, target_width, target_height)
        }
        RenderedKind::Png => {
            let image = image::open(&cached.path).ok()?;
            let resized = image.resize_exact(
                target_width,
                target_height,
                image::imageops::FilterType::Triangle,
            );

            Some((
                pdf_preview::opaque_bgra(resized.to_rgba8().as_raw()),
                target_width,
                target_height,
            ))
        }
    }
}

fn render_thumbnail(
    thumbnail: &Thumbnail,
    target_width: u32,
    target_height: u32,
) -> Option<(Vec<u8>, u32, u32)> {
    let (width, height) = fit_within(
        thumbnail.width,
        thumbnail.height,
        target_width,
        target_height,
    )?;

    match thumbnail.kind {
        ThumbnailKind::Emf { .. } | ThumbnailKind::Wmf { .. } => {
            play_metafile(&thumbnail.bytes, thumbnail.kind, width, height)
        }
        ThumbnailKind::Dib => stretch_dib(&thumbnail.bytes, width, height),
        ThumbnailKind::Raster => {
            let image = image::load_from_memory(&thumbnail.bytes).ok()?;
            let resized = image.resize_exact(width, height, image::imageops::FilterType::Triangle);

            Some((
                pdf_preview::opaque_bgra(resized.to_rgba8().as_raw()),
                width,
                height,
            ))
        }
    }
}

/// The source's own aspect ratio inside the caller's box, so a picture that is
/// not the shape the box is is letterboxed rather than stretched.
fn fit_within(width: u32, height: u32, max_width: u32, max_height: u32) -> Option<(u32, u32)> {
    if width == 0 || height == 0 || max_width == 0 || max_height == 0 {
        return None;
    }

    let scale = (max_width as f32 / width as f32).min(max_height as f32 / height as f32);

    Some((
        (width as f32 * scale).round().clamp(1.0, max_width as f32) as u32,
        (height as f32 * scale)
            .round()
            .clamp(1.0, max_height as f32) as u32,
    ))
}

/// Draw a metafile — an EMF, or the WMF an Excel workbook and a legacy document
/// save — into a surface of the given size.
///
/// The page is painted white under it and the alpha byte forced opaque on the way
/// out, exactly as a PDF page is: a page is a page whatever it was rendered from,
/// and the preview composites it over the configured background.
fn play_metafile(
    bytes: &[u8],
    kind: ThumbnailKind,
    width: u32,
    height: u32,
) -> Option<(Vec<u8>, u32, u32)> {
    let surface = DibSurface::create(width, height)?;
    let rect = RECT {
        left: 0,
        top: 0,
        right: width as i32,
        bottom: height as i32,
    };

    unsafe {
        let brush = CreateSolidBrush(COLORREF(0x00FF_FFFF));
        FillRect(surface.dc, &rect, brush);
        let _ = DeleteObject(HGDIOBJ(brush.0));

        match kind {
            // An enhanced metafile carries its own frame, and the play maps that
            // frame onto the rectangle it is handed.
            ThumbnailKind::Emf { .. } => {
                let metafile = SetEnhMetaFileBits(bytes);
                if metafile.0.is_null() {
                    return None;
                }
                let _ = PlayEnhMetaFile(surface.dc, metafile, &rect);
                let _ = DeleteEnhMetaFile(metafile);
            }
            // A Windows metafile has no frame of its own to map: the extent it is
            // drawn at is the METAFILEPICT's, which is what the anisotropic
            // mapping and the target size say. Converting it that way is what
            // makes a placeable metafile placeable — the handle that comes back
            // is an enhanced one, and it is played like any other.
            ThumbnailKind::Wmf { .. } => {
                let mapping = METAFILEPICT {
                    mm: MM_ANISOTROPIC.0,
                    xExt: width as i32,
                    yExt: height as i32,
                    // The handle field is not part of what the mapping says: the
                    // bits are already in hand, and what GDI is told is only the
                    // size to map them onto.
                    hMF: HMETAFILE(std::ptr::null_mut()),
                };
                let metafile = SetWinMetaFileBits(bytes, surface.dc, Some(&mapping));
                if metafile.0.is_null() {
                    return None;
                }
                let _ = PlayEnhMetaFile(surface.dc, metafile, &rect);
                let _ = DeleteEnhMetaFile(metafile);
            }
            _ => return None,
        }
    }

    Some((surface.pixels(), width, height))
}

/// Draw a device-independent bitmap — a `BITMAPINFO` and its bits, with no file
/// header — into a surface of the given size.
fn stretch_dib(bytes: &[u8], width: u32, height: u32) -> Option<(Vec<u8>, u32, u32)> {
    let header_size = read_u32(bytes, 0)? as usize;
    if header_size < 40 || header_size > bytes.len() {
        return None;
    }

    let source_width = read_i32(bytes, 4)?;
    let source_height = read_i32(bytes, 8)?;
    if source_width <= 0 || source_height == 0 {
        return None;
    }

    // The colour table sits between the header and the bits: as many entries as
    // the header says, or as many as the bit depth implies.
    let bit_count = read_u16(bytes, 14)? as u32;
    let used_colors = read_u32(bytes, 32)?;
    let palette_entries = if bit_count <= 8 {
        if used_colors != 0 {
            used_colors
        } else {
            1u32 << bit_count
        }
    } else {
        0
    };
    let bits_offset = header_size + palette_entries as usize * 4;
    if bits_offset >= bytes.len() {
        return None;
    }

    let surface = DibSurface::create(width, height)?;

    unsafe {
        SetStretchBltMode(surface.dc, COLORONCOLOR);
        StretchDIBits(
            surface.dc,
            0,
            0,
            width as i32,
            height as i32,
            0,
            0,
            source_width,
            source_height,
            Some(bytes.as_ptr().add(bits_offset) as *const c_void),
            bytes.as_ptr() as *const BITMAPINFO,
            DIB_RGB_COLORS,
            SRCCOPY,
        );
    }

    Some((surface.pixels(), width, height))
}

fn read_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_le_bytes(
        bytes.get(offset..offset + 2)?.try_into().ok()?,
    ))
}

fn read_u32(bytes: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(
        bytes.get(offset..offset + 4)?.try_into().ok()?,
    ))
}

fn read_i32(bytes: &[u8], offset: usize) -> Option<i32> {
    Some(i32::from_le_bytes(
        bytes.get(offset..offset + 4)?.try_into().ok()?,
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs::File;
    use std::io::Write;
    use windows::Win32::Foundation::RECT as GdiRect;
    use windows::Win32::Graphics::Gdi::{
        CloseMetaFile, CreateMetaFileW, CreateSolidBrush as TestBrush, DeleteMetaFile,
        DeleteObject as TestDelete, GetMetaFileBitsEx, Rectangle,
    };

    /// A Windows metafile recorded by GDI itself: a filled black square over the
    /// whole of its frame. Recording one is how Office's own thumbnails come to
    /// exist, so this is the shape of the bytes the reader meets.
    fn recorded_wmf() -> Vec<u8> {
        unsafe {
            let dc = CreateMetaFileW(None);
            assert!(!dc.0.is_null(), "a recording device context");

            let brush = TestBrush(COLORREF(0));
            let previous = windows::Win32::Graphics::Gdi::SelectObject(dc, HGDIOBJ(brush.0));
            let rect = GdiRect {
                left: 0,
                top: 0,
                right: 100,
                bottom: 50,
            };
            let _ = Rectangle(dc, rect.left, rect.top, rect.right, rect.bottom);
            let _ = windows::Win32::Graphics::Gdi::SelectObject(dc, previous);
            let _ = TestDelete(HGDIOBJ(brush.0));

            let metafile = CloseMetaFile(dc);
            let size = GetMetaFileBitsEx(metafile, 0, None);
            assert!(size > 0, "a recorded metafile has bytes");

            let mut bytes = vec![0u8; size as usize];
            let written =
                GetMetaFileBitsEx(metafile, size, Some(bytes.as_mut_ptr() as *mut c_void));
            assert_eq!(written, size);
            let _ = DeleteMetaFile(metafile);

            bytes
        }
    }

    #[test]
    fn draws_a_recorded_metafile_at_the_size_it_is_asked_for() {
        let bytes = recorded_wmf();

        let (pixels, width, height) = play_metafile(
            &bytes,
            ThumbnailKind::Wmf {
                extent: crate::office_thumbnail::WmfExtent::Twips(1500, 750),
            },
            40,
            20,
        )
        .expect("a drawn metafile");

        assert_eq!((width, height), (40, 20));
        assert_eq!(pixels.len(), 40 * 20 * 4);
        assert!(
            pixels.chunks_exact(4).all(|pixel| pixel[3] == 255),
            "a drawn page is opaque"
        );
        assert!(
            pixels.chunks_exact(4).any(|pixel| pixel[0] < 128),
            "the metafile's own drawing reaches the surface"
        );
    }

    #[test]
    fn fits_a_thumbnail_into_the_box_it_is_given() {
        assert_eq!(fit_within(100, 50, 400, 400), Some((400, 200)));
        assert_eq!(fit_within(100, 50, 400, 100), Some((200, 100)));
        assert_eq!(fit_within(100, 50, 50, 50), Some((50, 25)));
        assert_eq!(fit_within(0, 50, 50, 50), None);
    }

    /// The recorded metafile with the Aldus header that makes it placeable: a box
    /// of 1440 by 720 twips, which is 96 by 48 pixels at 96 DPI.
    fn placeable_wmf(bits: &[u8]) -> Vec<u8> {
        let mut header = Vec::new();
        header.extend_from_slice(&0x9AC6_CDD7u32.to_le_bytes()); // key
        header.extend_from_slice(&0u16.to_le_bytes()); // handle, which a file cannot carry
        header.extend_from_slice(&0i16.to_le_bytes()); // left
        header.extend_from_slice(&0i16.to_le_bytes()); // top
        header.extend_from_slice(&1440i16.to_le_bytes()); // right
        header.extend_from_slice(&720i16.to_le_bytes()); // bottom
        header.extend_from_slice(&1440u16.to_le_bytes()); // twips to the inch
        header.extend_from_slice(&0u32.to_le_bytes()); // reserved
        header.extend_from_slice(&0u16.to_le_bytes()); // checksum, which readers ignore

        header.extend_from_slice(bits);
        header
    }

    /// A document is previewed the way a hover previews one: the file's own header
    /// says it is a package, the package holds a picture of the page, and that
    /// picture is drawn into the box the layout planned. Everything is real Office
    /// behaviour except that the picture was recorded by GDI rather than by Word.
    #[test]
    fn draws_the_picture_a_package_saved() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-office-tests")
            .join("preview");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let path = folder.join("saved-thumbnail.docx");

        let file = File::create(&path).expect("a test package");
        let mut writer = zip::ZipWriter::new(file);
        let stored = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        writer
            .start_file("docProps/thumbnail.wmf", stored)
            .expect("a written entry");
        writer
            .write_all(&placeable_wmf(&recorded_wmf()))
            .expect("a written picture");
        writer.finish().expect("a finished package");

        // The picture states its own size, which is what the layout places it by.
        assert_eq!(measure(&path), Some((96, 48)));

        let (pixels, width, height) = render(&path, 480, 240, None).expect("a drawn picture");
        assert_eq!((width, height), (480, 240));
        assert_eq!(pixels.len(), 480 * 240 * 4);
        assert!(
            pixels.chunks_exact(4).all(|pixel| pixel[3] == 255),
            "the drawn page is opaque"
        );
        assert!(
            pixels.chunks_exact(4).any(|pixel| pixel[0] < 128),
            "the picture reaches the frame"
        );

        let _ = std::fs::remove_file(&path);
    }
}

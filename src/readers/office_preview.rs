//! Measuring and drawing the page rendered for an Office document.
//!
//! A preview is drawn from one source and one only: the page an engine drew and the document
//! cache keeps (see `office_render` and `document_cache`). The picture a document saves inside
//! itself is deliberately not read — it is a thumbnail-sized metafile or bitmap, a couple of
//! hundred pixels across, and a preview drawn from one is either tiny or an enlargement of
//! something that small — so a document whose page is not there yet is answered with a spinner
//! in a box of its own, and the page itself the moment it arrives.
//!
//! One document is drawn by the render engine beside Office rather than by an application of
//! its own: one whose own application is not installed, where there is no Word, Excel or
//! PowerPoint to ask for a page — and one the tray has asked the engine for outright, under
//! `Engine → Select Engine → Office`. What that engine writes is a page like any other here —
//! measured from its own first page, laid out beside the cursor, and drawn at whatever size the
//! layout asks for — which is what `engine_page` reads and what `measure` and `source_kind` ask
//! about before they answer with the wait (see `libre_formats::engine_page_kind` for which
//! documents those are).
//!
//! Which of the two engines is the document's own is one question, asked of the configuration
//! and the machine together (`office_formats::page_engine`), and the answer is what both
//! sources are read through here: a page the tier drew is not the page a hover in the engine's
//! mode is shown, and neither is a page the engine drew when it was the one being asked.
//!
//! Nothing is kept *here*: a page belongs to the document cache, which holds what the user's
//! budget allows and nothing else, and what this module does with one is cheap either way — a
//! page costs a header parse to measure and one raster to draw. The size it reads is
//! remembered by the document it was drawn for, so it cannot outlive the page it was read from
//! (see `document_cache::size`).

use crate::config::config::{image_decode_limits, read_within_budget, OfficeEngine};
use crate::engines::document_cache::{self, Page, PageKind};
use crate::engines::libreoffice_render;
use crate::engines::office_render;
use crate::formats::office_formats;
use crate::readers::pdf_preview;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};

/// The box a preview is placed in while its page is being rendered: the spinner's
/// own frame, since there is nothing else to show yet. The page's own size is what
/// the layout uses the moment it exists, and the window is moved to it.
///
/// The box is the size of the arc and its halo, so a spinner placed flush at the
/// pointer is the spinner at the pointer rather than an empty frame around it;
/// and it is placed at the size it is rather than fitted to the display — there is
/// nothing in it to enlarge — so a hover that is waiting on Office costs a corner
/// of the screen and not the whole of it.
pub(crate) const WAITING_BOX: u32 = 36;

/// What a preview of this document would be drawn from.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum SourceKind {
    /// A page Office exported — vector, drawn at whatever size it is asked for.
    Page,
    /// The bitmap a workbook is answered with where no page can be exported: only
    /// as good as the pixels it holds.
    Raster,
    /// Nothing to draw from yet, and a page on the way.
    None,
}

impl SourceKind {
    /// Whether a preview drawn from this source may be enlarged to fill the
    /// display. A bitmap may not: enlarging it would only stretch the pixels it
    /// holds, which is why it follows the configured scale instead — the rule an
    /// image follows.
    pub(crate) fn may_be_enlarged(self) -> bool {
        !matches!(self, Self::Raster)
    }
}

/// Which of the document's sources a preview would be drawn from, asked by the
/// layout before it decides how large the preview may be.
pub(crate) fn source_kind(path: &Path) -> SourceKind {
    if let Some((page, _)) = usable_page(path) {
        // A page Office exported is drawn at whatever size it is asked for; the picture
        // a workbook is answered with on a machine that cannot export a page is a screen
        // bitmap, and enlarging that would only stretch it. Which of the two a PNG is is not
        // something the kind says any more — a slide's export and a workbook's picture are both
        // PNGs — so the family is asked, exactly as the render tier asks it before drawing
        // anything (see `office_render::page_is_workbook_picture`).
        return match page.kind {
            PageKind::Pdf => SourceKind::Page,
            PageKind::Png if !office_render::page_is_workbook_picture(path, &page) => {
                SourceKind::Page
            }
            PageKind::Png | PageKind::Bmp => SourceKind::Raster,
        };
    }

    // The page the render engine beside Office drew is a page like any other, and for the
    // same reason: it is a PDF, and it is drawn at whatever size it is asked for — which is
    // what the layout has to know before it places one, or the page would be laid out as the
    // wait for it and shown at the spinner's own size. The document it was drawn for is one
    // whose own application is not installed (see `libreoffice_render`), so there is no other
    // source coming that the layout could be waiting for instead.
    if engine_page(path).is_some() {
        return SourceKind::Page;
    }

    SourceKind::None
}

/// The page the render engine beside Office has drawn for this document, where it has drawn
/// one: a PDF under the app's own folder, named for the document and the version of it that
/// was converted. Nothing is started and nothing is waited on — whether there is a page is a
/// read of that folder (see `libreoffice_render`).
///
/// It is read only where the engine is the one that draws this document: a PDF the engine
/// wrote while it was being asked is not the page a hover in the application's mode is shown
/// or measured from (see `office_formats::page_engine`).
fn engine_page(path: &Path) -> Option<PathBuf> {
    matches!(
        office_formats::page_engine(path),
        Some(OfficeEngine::LibreOffice)
    )
    .then(|| libreoffice_render::rendered_page(path))
    .flatten()
}

/// The page the render tier holds for this document, and its own size, where there is one that
/// can be read.
///
/// A page that cannot be read is not a page: one a render cut short, or one something else
/// corrupted, would otherwise be handed to a preview that blinks away the moment it tries to
/// draw it — and it would be trusted for good, because "a page is already kept for this file"
/// is what stops another render. So a page whose size will not come out of it is given up here,
/// which is what gets the document rendered again. This is the side that may ask the PDF engine
/// — its threads are multithreaded apartments — which is why the check lives here rather than
/// where the page is kept.
///
/// It answers with nothing at all where the page kept is not one this document is previewed
/// from any more: a page the render tier produced while it was the engine being asked is not
/// the page a hover in the engine's mode is shown, and what this side drew would be the other
/// engine's work under the current choice (see `office_formats::page_engine`). The page is left
/// where it is rather than dropped, so a choice that comes back to the tier finds it.
fn usable_page(path: &Path) -> Option<(Page, (u32, u32))> {
    if office_formats::page_engine(path) != Some(OfficeEngine::MicrosoftOffice) {
        return None;
    }

    let page = office_render::held_page(path)?;
    if let Some(size) = document_cache::size(path, OfficeEngine::MicrosoftOffice.as_str()) {
        return Some((page, size));
    }

    document_cache::forget(path, OfficeEngine::MicrosoftOffice.as_str());
    None
}

/// The size the layout places a preview of this document from.
///
/// `None` means there is nothing to show and nothing on the way, which drops the
/// preview — a document whose family has no engine, or a render tier that is
/// switched off, is answered with no preview rather than with a box that would
/// spin forever.
pub(crate) fn measure(path: &Path) -> Option<(u32, u32)> {
    if let Some((_, size)) = usable_page(path) {
        return Some(size);
    }

    // Nothing of Office's has been drawn, but the render engine beside it may have drawn a
    // page for a document whose own application is not installed: that page is a page, so the
    // hover is measured from it and placed beside the cursor at the size it will be shown at,
    // rather than laid out as the wait for itself (see `source_kind`).
    if let Some(page) = engine_page(path) {
        if let Some(dimensions) = pdf_preview::page_dimensions(&page) {
            return Some(dimensions);
        }
    }

    // Nothing rendered yet: a page is what is coming, so the preview is a spinner
    // in a box of its own until it arrives.
    if office_render::enabled() {
        return Some((WAITING_BOX, WAITING_BOX));
    }

    None
}

/// The rendered page, drawn into the box the layout planned — preserving the
/// page's own aspect ratio inside it.
pub(crate) fn render(
    path: &Path,
    target_width: u32,
    target_height: u32,
    cancel: Option<&AtomicBool>,
) -> Option<(Vec<u8>, u32, u32)> {
    if cancel
        .map(|cancel| cancel.load(Ordering::Acquire))
        .unwrap_or(false)
    {
        return None;
    }

    let (page, _) = usable_page(path)?;
    render_page(&page, target_width, target_height)
}

/// Draw the page the engine drew into the box the layout planned.
fn render_page(page: &Page, target_width: u32, target_height: u32) -> Option<(Vec<u8>, u32, u32)> {
    match page.kind {
        // The page is rendered at the size it is shown at rather than enlarged
        // afterwards, which is what keeps a scaled-up preview sharp.
        PageKind::Pdf => pdf_preview::render_first_page(&page.path, target_width, target_height),
        PageKind::Png | PageKind::Bmp => {
            // Read under the same limits as every other decode: the file is this app's own
            // render's, but the decoder asking is the same decoder, and this is a decode like
            // any other.
            let bytes = read_within_budget(&page.path)?;
            let mut reader = image::ImageReader::new(Cursor::new(&bytes[..]))
                .with_guessed_format()
                .ok()?;
            reader.limits(image_decode_limits());
            let image = reader.decode().ok()?;
            // A picture can be far larger than the box it is shown in — a
            // worksheet's corner at screen resolution is millions of pixels — so
            // shrinking uses the box filter, which is the fast one, and enlarging
            // the smooth one.
            let resized = if target_width <= image.width() && target_height <= image.height() {
                image.thumbnail_exact(target_width, target_height)
            } else {
                image.resize_exact(
                    target_width,
                    target_height,
                    image::imageops::FilterType::Triangle,
                )
            };

            Some((
                pdf_preview::opaque_bgra(resized.to_rgba8().as_raw()),
                target_width,
                target_height,
            ))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A BMP holding one colour, built the way the workbook picture path builds
    /// one: a `BITMAPINFO` and its pixels with a file header in front of them. The
    /// alpha byte is zero, which is what Excel writes there.
    fn bmp_bytes(width: u32, height: u32, color: [u8; 4]) -> Vec<u8> {
        let pixel_bytes = width * height * 4;
        let mut dib = Vec::new();
        dib.extend_from_slice(&40u32.to_le_bytes()); // header size
        dib.extend_from_slice(&(width as i32).to_le_bytes());
        dib.extend_from_slice(&(height as i32).to_le_bytes());
        dib.extend_from_slice(&1u16.to_le_bytes()); // planes
        dib.extend_from_slice(&32u16.to_le_bytes()); // bits per pixel
        dib.extend_from_slice(&0u32.to_le_bytes()); // BI_RGB
        dib.extend_from_slice(&pixel_bytes.to_le_bytes());
        for _ in 0..4 {
            dib.extend_from_slice(&0i32.to_le_bytes()); // resolution and colours
        }
        for _ in 0..(width * height) {
            dib.extend_from_slice(&color);
        }

        let mut file = Vec::new();
        file.extend_from_slice(b"BM");
        file.extend_from_slice(&((dib.len() + 14) as u32).to_le_bytes());
        file.extend_from_slice(&0u16.to_le_bytes());
        file.extend_from_slice(&0u16.to_le_bytes());
        file.extend_from_slice(&54u32.to_le_bytes()); // where the pixels start
        file.extend_from_slice(&dib);
        file
    }

    /// A page as the document cache keeps one: the bytes written where a page is kept, under
    /// the name the kind of page gives them.
    fn kept_page(kind: PageKind, bytes: &[u8]) -> Page {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-office-tests")
            .join("preview-pages");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let path = folder.join(format!("page.{}", kind.extension()));
        std::fs::write(&path, bytes).expect("a written page");

        Page { path, kind }
    }

    /// The picture a workbook is answered with is drawn like any other frame, and
    /// it is opaque whatever its own alpha bytes say.
    #[test]
    fn draws_a_workbook_picture_opaquely() {
        let page = kept_page(PageKind::Bmp, &bmp_bytes(2, 2, [40, 90, 200, 0]));
        let (pixels, width, height) = render_page(&page, 4, 4).expect("a drawn picture");

        assert_eq!((width, height), (4, 4));
        assert_eq!(pixels.len(), 4 * 4 * 4);
        assert!(
            pixels
                .as_chunks::<4>()
                .0
                .iter()
                .all(|pixel| pixel[3] == 255),
            "a drawn picture is opaque"
        );
        // Blue, green, red — the order the frame is in — and the colour survived.
        assert_eq!(&pixels[..3], &[40, 90, 200]);
    }

    /// A slide is drawn at the size the layout asked for.
    #[test]
    fn draws_a_slide_at_the_size_it_is_asked_for() {
        let slide = image::RgbaImage::from_pixel(8, 4, image::Rgba([10, 20, 30, 255]));
        let mut written = Cursor::new(Vec::new());
        slide
            .write_to(&mut written, image::ImageFormat::Png)
            .expect("a written slide");

        let page = kept_page(PageKind::Png, &written.into_inner());
        let (pixels, width, height) = render_page(&page, 32, 16).expect("a drawn slide");

        assert_eq!((width, height), (32, 16));
        assert_eq!(pixels.len(), 32 * 16 * 4);
        assert_eq!(&pixels[..3], &[30, 20, 10]);
    }
}

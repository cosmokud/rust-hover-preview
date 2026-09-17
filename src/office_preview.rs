//! Measuring and drawing the page rendered for an Office document.
//!
//! A preview is drawn from one source and one only: the page the render tier
//! produced and cached (see `office_render`). The picture a document saves inside
//! itself is deliberately not read — it is a thumbnail-sized metafile or bitmap, a
//! couple of hundred pixels across, and a preview drawn from one is either tiny or
//! an enlargement of something that small — so a document whose page is not there
//! yet is answered with a spinner in a box of its own, and the page itself the
//! moment it arrives.
//!
//! Nothing here is kept in memory between hovers. What a preview is drawn from is
//! a file on disk, and its size is read from that file, so a hover costs what
//! opening the file costs — milliseconds — and nothing can go stale.

use crate::office_render::{self, CachedRender, RenderedKind};
use crate::pdf_preview;
use std::path::Path;
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
    let Some(cached) = usable_render(path) else {
        return SourceKind::None;
    };

    // A page Office exported is drawn at whatever size it is asked for; the picture
    // a workbook is answered with on a machine that cannot export a page is a screen
    // bitmap, and enlarging that would only stretch it.
    match cached.kind {
        RenderedKind::Pdf | RenderedKind::Png => SourceKind::Page,
        RenderedKind::Bmp => SourceKind::Raster,
    }
}

/// The page waiting for this document, if there is one and it can be read.
///
/// A cache file that cannot be read is not a page: one left half-written by a
/// process that was ended mid-render, or one something else has corrupted, would
/// otherwise be handed to a preview that blinks away the moment it tries to draw
/// it — and it would be trusted forever, because "a page is already rendered for
/// this file" is what stops another render. So a file that will not give up its page
/// is dropped here, which is what gets the document rendered again. This is the side
/// that may ask the PDF engine — its threads are multithreaded apartments — which is
/// why the check lives here rather than where the cache is listed.
fn usable_render(path: &Path) -> Option<CachedRender> {
    let cached = office_render::cached_render(path)?;
    if rendered_dimensions(&cached).is_some() {
        return Some(cached);
    }

    let _ = std::fs::remove_file(&cached.path);
    None
}

/// The size the layout places a preview of this document from.
///
/// `None` means there is nothing to show and nothing on the way, which drops the
/// preview — a document whose family has no engine, or a render tier that is
/// switched off, is answered with no preview rather than with a box that would
/// spin forever.
pub(crate) fn measure(path: &Path) -> Option<(u32, u32)> {
    if let Some(cached) = usable_render(path) {
        if let Some(dimensions) = rendered_dimensions(&cached) {
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

    let cached = usable_render(path)?;
    render_cached(&cached, target_width, target_height)
}

fn rendered_dimensions(cached: &CachedRender) -> Option<(u32, u32)> {
    match cached.kind {
        RenderedKind::Pdf => pdf_preview::page_dimensions(&cached.path),
        // A slide's PNG and a workbook's bitmap are both read the way any image is.
        RenderedKind::Png | RenderedKind::Bmp => image::image_dimensions(&cached.path).ok(),
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
        RenderedKind::Png | RenderedKind::Bmp => {
            let image = image::open(&cached.path).ok()?;
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
        let pixel_bytes = (width * height * 4) as u32;
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

    /// The picture a workbook is answered with is drawn like any other frame, and
    /// it is opaque whatever its own alpha bytes say.
    #[test]
    fn draws_a_workbook_picture_opaquely() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-office-tests")
            .join("preview");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let path = folder.join("used-range.bmp");
        std::fs::write(&path, bmp_bytes(2, 2, [40, 90, 200, 0])).expect("a written picture");

        let cached = CachedRender {
            path: path.clone(),
            kind: RenderedKind::Bmp,
        };
        let (pixels, width, height) = render_cached(&cached, 4, 4).expect("a drawn picture");

        assert_eq!((width, height), (4, 4));
        assert_eq!(pixels.len(), 4 * 4 * 4);
        assert!(
            pixels.chunks_exact(4).all(|pixel| pixel[3] == 255),
            "a drawn picture is opaque"
        );
        // Blue, green, red — the order the frame is in — and the colour survived.
        assert_eq!(&pixels[..3], &[40, 90, 200]);

        let _ = std::fs::remove_file(&path);
    }

    /// A slide is drawn at the size the layout asked for.
    #[test]
    fn draws_a_slide_at_the_size_it_is_asked_for() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-office-tests")
            .join("preview");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let path = folder.join("slide.png");

        let slide = image::RgbaImage::from_pixel(8, 4, image::Rgba([10, 20, 30, 255]));
        slide.save(&path).expect("a written slide");

        let cached = CachedRender {
            path: path.clone(),
            kind: RenderedKind::Png,
        };
        let (pixels, width, height) = render_cached(&cached, 32, 16).expect("a drawn slide");

        assert_eq!((width, height), (32, 16));
        assert_eq!(pixels.len(), 32 * 16 * 4);
        assert_eq!(&pixels[..3], &[30, 20, 10]);

        let _ = std::fs::remove_file(&path);
    }
}

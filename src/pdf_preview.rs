use crate::config::PreviewType;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::os::windows::ffi::OsStrExt;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use windows::core::PCWSTR;
use windows::Data::Pdf::{PdfDocument, PdfPage, PdfPageRenderOptions};
use windows::Graphics::Imaging::BitmapEncoder;
use windows::Storage::Streams::{DataReader, IRandomAccessStream, InMemoryRandomAccessStream};
use windows::UI::Color;
use windows::Win32::System::Com::{
    CoInitializeEx, IStream, COINIT_MULTITHREADED, STGM_READ, STGM_SHARE_DENY_NONE,
};
use windows::Win32::System::WinRT::{CreateRandomAccessStreamOverStream, BSOS_DEFAULT};
use windows::Win32::UI::Shell::SHCreateStreamOnFileEx;

/// A PDF header may sit behind leading bytes, so the whole first kilobyte is
/// searched for the signature rather than only its start.
const PDF_HEADER_PROBE_BYTES: usize = 1024;
const PDF_HEADER: &[u8] = b"%PDF-";
const PAGE_DIMENSION_CACHE_MAX_ENTRIES: usize = 512;

/// A4 portrait at 96 DPI, used to place a preview whose page size is not known
/// yet. `PdfPage::Size` reports DIPs, which are also 96ths of an inch.
pub const DEFAULT_PAGE_WIDTH: u32 = 794;
pub const DEFAULT_PAGE_HEIGHT: u32 = 1123;

/// Page size in DIPs, keyed by path. `None` records a file that could not be
/// opened as a PDF, so a broken file is not re-parsed on every hover.
type PageDimensionCache = HashMap<PathBuf, Option<(u32, u32)>>;

static PAGE_DIMENSIONS: Lazy<Mutex<PageDimensionCache>> = Lazy::new(|| Mutex::new(HashMap::new()));

/// Initialize the apartment this module's WinRT calls need. Every thread that
/// probes or renders a page has to call this once before its first call.
pub fn initialize_apartment() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_MULTITHREADED);
    }
}

pub fn is_pdf_file(path: &Path) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_lowercase() == "pdf")
        .unwrap_or(false)
}

/// Whether a PDF preview may be shown for `path`: the file a page would be read
/// from, and the `PDF` gate in the tray's `Toggle Preview Types` submenu.
pub fn is_pdf_preview(path: &Path) -> bool {
    is_pdf_file(path) && PreviewType::Pdf.enabled()
}

/// Page 1's size in DIPs, from the cache when the file has been seen before.
///
/// `None` means the file could not be read as a PDF, which is treated as "not
/// previewable" rather than as "use a guessed size".
pub fn page_dimensions(path: &Path) -> Option<(u32, u32)> {
    if let Ok(cache) = PAGE_DIMENSIONS.lock() {
        if let Some(cached) = cache.get(path) {
            return *cached;
        }
    }

    let dimensions = probe_page_dimensions(path);
    remember_page_dimensions(path, dimensions);

    dimensions
}

/// Record a page size that has been read, so it never has to be read again.
///
/// Both the measure and the render open the document — which, because
/// `Windows.Data.Pdf` is handed a stream, means reading the whole file — so what
/// one of them read is handed to the other's cache rather than left for it to
/// find a second time.
fn remember_page_dimensions(path: &Path, dimensions: Option<(u32, u32)>) {
    if let Ok(mut cache) = PAGE_DIMENSIONS.lock() {
        if !cache.contains_key(path) && cache.len() >= PAGE_DIMENSION_CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.insert(path.to_path_buf(), dimensions);
    }
}

/// Render page 1 into the largest box that fits `max_width` x `max_height`
/// without changing the page's aspect ratio, and return BGRA pixels with the
/// size they were rendered at.
pub fn render_first_page(
    path: &Path,
    max_width: u32,
    max_height: u32,
) -> Option<(Vec<u8>, u32, u32)> {
    if !has_pdf_header(path) {
        return None;
    }

    let document = open_document(path)?;
    let page = document.GetPage(0).ok()?;
    let size = page.Size().ok()?;
    let dimensions = page_size_in_dips(size.Width, size.Height);
    let (render_width, render_height) = fit_page(size.Width, size.Height, max_width, max_height)?;
    remember_page_dimensions(path, dimensions);

    let bytes = render_page(&page, render_width, render_height)?;
    let image = image::load_from_memory(&bytes).ok()?.to_rgba8();
    let (width, height) = (image.width(), image.height());
    if width == 0 || height == 0 {
        return None;
    }

    Some((opaque_bgra(image.as_raw()), width, height))
}

fn probe_page_dimensions(path: &Path) -> Option<(u32, u32)> {
    if !has_pdf_header(path) {
        return None;
    }

    let document = open_document(path)?;
    let size = document.GetPage(0).ok()?.Size().ok()?;
    page_size_in_dips(size.Width, size.Height)
}

fn page_size_in_dips(width: f32, height: f32) -> Option<(u32, u32)> {
    if !width.is_finite() || !height.is_finite() || width <= 0.0 || height <= 0.0 {
        return None;
    }

    Some((
        (width.round() as u32).max(1),
        (height.round() as u32).max(1),
    ))
}

/// The page's own aspect ratio inside the caller's box, so a page that is not
/// the shape the caller assumed is letterboxed instead of stretched.
fn fit_page(width: f32, height: f32, max_width: u32, max_height: u32) -> Option<(u32, u32)> {
    let (page_width, page_height) = page_size_in_dips(width, height)?;
    if max_width == 0 || max_height == 0 {
        return None;
    }

    let scale =
        (max_width as f32 / page_width as f32).min(max_height as f32 / page_height as f32);
    let target_width = (page_width as f32 * scale).round().clamp(1.0, max_width as f32) as u32;
    let target_height = (page_height as f32 * scale).round().clamp(1.0, max_height as f32) as u32;

    Some((target_width, target_height))
}

/// Open the document from a stream over the file itself.
///
/// A PDF is read by seeking: the object graph is found from a table at the end of
/// the file, and a page's content and the resources it names can be anywhere in
/// it. A stream over the file lets the engine touch only what page 1 actually
/// needs, instead of the whole file being read into memory and then copied again
/// into a stream — so what a hover onto a large PDF costs is proportional to the
/// page rather than to the file.
///
/// The file is opened with the path forms the rest of the app produces, the
/// verbatim `\\?\` form included, because `SHCreateStreamOnFileEx` is handed the
/// path rather than the WinRT broker being asked to resolve it — that refusal by
/// `StorageFile.GetFileFromPathAsync` is why this used to be read into memory by
/// hand.
fn open_document(path: &Path) -> Option<PdfDocument> {
    let mut wide: Vec<u16> = path.as_os_str().encode_wide().collect();
    wide.push(0);

    let stream: IRandomAccessStream = unsafe {
        let file = SHCreateStreamOnFileEx(
            PCWSTR(wide.as_ptr()),
            STGM_READ.0 | STGM_SHARE_DENY_NONE.0,
            0,
            false,
            None::<&IStream>,
        )
        .ok()?;

        CreateRandomAccessStreamOverStream(&file, BSOS_DEFAULT).ok()?
    };

    PdfDocument::LoadFromStreamAsync(&stream).ok()?.get().ok()
}

fn render_page(page: &PdfPage, width: u32, height: u32) -> Option<Vec<u8>> {
    let options = PdfPageRenderOptions::new().ok()?;
    options.SetDestinationWidth(width).ok()?;
    options.SetDestinationHeight(height).ok()?;
    // A PDF page carries no background of its own and the preview composites the
    // frame over the configured background, so the page is painted white here.
    options
        .SetBackgroundColor(Color {
            A: 255,
            R: 255,
            G: 255,
            B: 255,
        })
        .ok()?;
    // Uncompressed pixels: the renderer's own encode is cheap and the decode on
    // this side becomes a header parse instead of a PNG inflate.
    options
        .SetBitmapEncoderId(BitmapEncoder::BmpEncoderId().ok()?)
        .ok()?;

    let stream = InMemoryRandomAccessStream::new().ok()?;
    page.RenderWithOptionsToStreamAsync(&stream, &options)
        .ok()?
        .get()
        .ok()?;

    read_stream_bytes(&stream)
}

fn read_stream_bytes(stream: &InMemoryRandomAccessStream) -> Option<Vec<u8>> {
    let length = stream.Size().ok()? as usize;
    if length == 0 {
        return None;
    }

    let input = stream.GetInputStreamAt(0).ok()?;
    let reader = DataReader::CreateDataReader(&input).ok()?;
    reader.LoadAsync(length as u32).ok()?.get().ok()?;

    let mut bytes = vec![0u8; length];
    reader.ReadBytes(&mut bytes).ok()?;
    Some(bytes)
}

/// BGRA for GDI, with the alpha channel written as opaque: the page is painted
/// on an opaque background above, so whatever the encoder leaves in the fourth
/// byte cannot make the preview invisible.
fn opaque_bgra(rgba: &[u8]) -> Vec<u8> {
    let mut bgra = Vec::with_capacity(rgba.len());
    for chunk in rgba.chunks(4) {
        if chunk.len() == 4 {
            bgra.push(chunk[2]); // B
            bgra.push(chunk[1]); // G
            bgra.push(chunk[0]); // R
            bgra.push(255); // A
        }
    }
    bgra
}

/// Confirm the file really is a PDF before it is handed to the renderer, so a
/// mislabeled file does not reach the OS parser at all.
fn has_pdf_header(path: &Path) -> bool {
    let Ok(mut file) = File::open(path) else {
        return false;
    };

    let mut probe = [0u8; PDF_HEADER_PROBE_BYTES];
    let read = match file.read(&mut probe) {
        Ok(read) => read,
        Err(_) => return false,
    };

    probe[..read]
        .windows(PDF_HEADER.len())
        .any(|window| window == PDF_HEADER)
}

use crate::config::config::{
    image_decode_limits, sanitize_ebook_cache_mb, PreviewType, DEFAULT_EBOOK_CACHE_MB,
};
use crate::shell::cloud_files;
use crate::CONFIG;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::os::windows::ffi::OsStrExt;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::SystemTime;
use windows::core::PCWSTR;
use windows::Data::Pdf::{PdfDocument, PdfPage, PdfPageRenderOptions};
use windows::Graphics::Imaging::BitmapEncoder;
use windows::Storage::Streams::{
    DataReader, DataWriter, IRandomAccessStream, InMemoryRandomAccessStream,
};
use windows::Win32::System::Com::{
    CoInitializeEx, IStream, COINIT_MULTITHREADED, STGM_READ, STGM_SHARE_DENY_NONE,
};
use windows::Win32::System::WinRT::{CreateRandomAccessStreamOverStream, BSOS_DEFAULT};
use windows::Win32::UI::Shell::SHCreateStreamOnFileEx;
use windows::UI::Color;

/// A PDF header may sit behind leading bytes, so the whole first kilobyte is
/// searched for the signature rather than only its start.
const PDF_HEADER_PROBE_BYTES: usize = 1024;
const PDF_HEADER: &[u8] = b"%PDF-";
const PAGE_DIMENSION_CACHE_MAX_ENTRIES: usize = 512;

/// A4 portrait at 96 DPI, used to place a preview whose page size is not known
/// yet. `PdfPage::Size` reports DIPs, which are also 96ths of an inch.
pub const DEFAULT_PAGE_WIDTH: u32 = 794;
pub const DEFAULT_PAGE_HEIGHT: u32 = 1123;

/// What a page's size is only valid for: the file, and the version of it the size was
/// read from. A file saved again is another document — one exported at another page
/// size is the case that matters — and its first page can be another shape.
#[derive(Clone, PartialEq, Eq, Hash, Debug)]
struct PageDimensionKey {
    path: PathBuf,
    modified: Option<SystemTime>,
    len: u64,
}

/// Page size in DIPs, keyed by the file and its version. `None` records a file that
/// could not be opened as a PDF, so a broken file is not re-parsed on every hover.
type PageDimensionCache = HashMap<PageDimensionKey, Option<(u32, u32)>>;

static PAGE_DIMENSIONS: Lazy<Mutex<PageDimensionCache>> = Lazy::new(|| Mutex::new(HashMap::new()));

/// Initialize the apartment this module's WinRT calls need. Every thread that
/// probes or renders a page has to call this once before its first call.
pub fn initialize_apartment() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_MULTITHREADED);
    }
}

/// Whether `path` is named with `extension`, whatever case it is written in.
fn named(path: &Path, extension: &str) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.eq_ignore_ascii_case(extension))
        .unwrap_or(false)
}

/// Whether a page may be read from `path`.
///
/// A `.pdf` is one by its name, and with it the two other spellings of the same document the
/// format's own world writes: `pdfa`, the archival profile of a PDF, and `epdf`, the
/// encapsulated one, are both pages the Windows engine opens exactly as it opens a `.pdf`. A
/// `.ai` is one when Illustrator saved it the way it saves one by default — with `Create PDF
/// Compatible File` on, which has been the default since Illustrator 9 — because the document
/// *is* page 1 of a PDF then, and the private data the application writes beside the artwork is
/// what the OS engine reads past. Saved without that compatibility the file is PostScript,
/// which is not a page any engine here can draw, so the bytes are asked rather than believed and
/// a file that answers no is a file with no preview.
///
/// Only the name that needs the question pays for it: the three page spellings are answered
/// without opening anything, and a cloud placeholder is not opened to answer either, which is
/// the rule every gate in this app follows.
pub fn is_pdf_file(path: &Path) -> bool {
    if named(path, "pdf") || named(path, "pdfa") || named(path, "epdf") {
        return true;
    }

    named(path, "ai") && !cloud_files::needs_download(path) && has_pdf_header(path)
}

/// Whether a PDF preview may be shown for `path`: the file a page would be read
/// from, and the `Ebook` gate in the tray's `Preview Types` submenu.
pub fn is_pdf_preview(path: &Path) -> bool {
    is_pdf_file(path) && PreviewType::Ebook.enabled()
}

/// Page 1's size in DIPs, from the cache when the file has been seen before.
///
/// `None` means the file could not be read as a PDF, which is treated as "not
/// previewable" rather than as "use a guessed size".
pub fn page_dimensions(path: &Path) -> Option<(u32, u32)> {
    let key = page_dimension_key(path);
    if let Ok(cache) = PAGE_DIMENSIONS.lock() {
        if let Some(cached) = cache.get(&key) {
            return *cached;
        }
    }

    let dimensions = probe_page_dimensions(path);
    remember_page_dimensions(key, dimensions);

    dimensions
}

/// The file and the version of it a page size is read from, read the way every other
/// held value in this app is: what a file is, is its name as it is now, what it
/// weighed, and when it was last written.
fn page_dimension_key(path: &Path) -> PageDimensionKey {
    let metadata = std::fs::metadata(path).ok();

    PageDimensionKey {
        path: path.to_path_buf(),
        modified: metadata
            .as_ref()
            .and_then(|metadata| metadata.modified().ok()),
        len: metadata.map(|metadata| metadata.len()).unwrap_or(0),
    }
}

/// Record a page size that has been read, so it never has to be read again.
///
/// Both the measure and the render open the document — which, because
/// `Windows.Data.Pdf` is handed a stream, means reading the whole file — so what
/// one of them read is handed to the other's cache rather than left for it to
/// find a second time.
fn remember_page_dimensions(key: PageDimensionKey, dimensions: Option<(u32, u32)>) {
    if let Ok(mut cache) = PAGE_DIMENSIONS.lock() {
        if !cache.contains_key(&key) && cache.len() >= PAGE_DIMENSION_CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.insert(key, dimensions);
    }
}

/// What a held page is only valid for: the file, the version of it that was
/// rendered, and the box it was rendered into.
///
/// The box is part of it because a page is stored as the pixels it was drawn as —
/// the same page at fit-to-screen and at `25%` really is different pixels — so only
/// the size that was asked for can be handed back for it; the version is there
/// because a file edited in place has to be drawn again.
#[derive(Clone, PartialEq, Eq, Hash, Debug)]
struct PageCacheKey {
    path: PathBuf,
    modified: Option<SystemTime>,
    len: u64,
    width: u32,
    height: u32,
}

/// A page being held: the pixels, and when they were last asked for.
struct PageCacheEntry {
    pixels: Vec<u8>,
    width: u32,
    height: u32,
    size: usize,
    last_used: u64,
}

/// The pages held in memory, and how much of the budget they take.
#[derive(Default)]
struct PageCache {
    entries: HashMap<PageCacheKey, PageCacheEntry>,
    bytes: usize,
    /// A counter rather than a clock, so the order pages are dropped in cannot be
    /// changed by the system clock moving.
    tick: u64,
}

static RENDERED_PAGES: Lazy<Mutex<PageCache>> = Lazy::new(|| Mutex::new(PageCache::default()));

/// The memory the cache may hold, read from the configuration each time rather
/// than captured, so an edit to `ebook_cache_mb` applies without a restart.
fn page_cache_limit_bytes() -> usize {
    let megabytes = CONFIG
        .lock()
        .map(|config| sanitize_ebook_cache_mb(config.ebook_cache_mb))
        .unwrap_or(DEFAULT_EBOOK_CACHE_MB);

    megabytes as usize * 1024 * 1024
}

fn page_cache_key(path: &Path, width: u32, height: u32) -> PageCacheKey {
    let metadata = std::fs::metadata(path).ok();

    PageCacheKey {
        path: path.to_path_buf(),
        modified: metadata
            .as_ref()
            .and_then(|metadata| metadata.modified().ok()),
        len: metadata.map(|metadata| metadata.len()).unwrap_or(0),
        width,
        height,
    }
}

/// The page held for this file at this size, if the cache still has it.
fn page_cache_get(key: &PageCacheKey) -> Option<(Vec<u8>, u32, u32)> {
    let limit = page_cache_limit_bytes();
    let mut cache = RENDERED_PAGES.lock().ok()?;

    page_cache_trim(&mut cache, limit);

    cache.tick += 1;
    let tick = cache.tick;

    let entry = cache.entries.get_mut(key)?;
    entry.last_used = tick;

    Some((entry.pixels.clone(), entry.width, entry.height))
}

/// Hold the page just rendered, dropping whatever no longer fits beside it.
fn page_cache_put(key: PageCacheKey, pixels: &[u8], width: u32, height: u32) {
    let limit = page_cache_limit_bytes();
    let Ok(mut cache) = RENDERED_PAGES.lock() else {
        return;
    };

    page_cache_trim(&mut cache, limit);

    let size = pixels.len();
    // A page larger than the whole budget would evict everything else and still
    // not fit, so it is simply not held — which is every page at a budget of
    // nothing, and is what makes that size mean "hold nothing".
    if size > limit {
        return;
    }

    cache.tick += 1;
    let tick = cache.tick;

    if let Some(previous) = cache.entries.insert(
        key,
        PageCacheEntry {
            pixels: pixels.to_vec(),
            width,
            height,
            size,
            last_used: tick,
        },
    ) {
        cache.bytes -= previous.size;
    }
    cache.bytes += size;

    page_cache_trim(&mut cache, limit);
}

/// Drop pages, least recently used first, until the cache fits inside `limit`.
fn page_cache_trim(cache: &mut PageCache, limit: usize) {
    while cache.bytes > limit {
        // Bound to its own statement so the borrow of `entries` has ended before
        // the entry is removed.
        let oldest = cache
            .entries
            .iter()
            .min_by_key(|(_, entry)| entry.last_used)
            .map(|(key, _)| key.clone());

        let Some(oldest) = oldest else {
            break;
        };

        if let Some(dropped) = cache.entries.remove(&oldest) {
            cache.bytes -= dropped.size;
        }
    }
}

/// Trim the cache to the configured size now, which is what the tray asks for when
/// a smaller size is chosen: what is over the new budget is freed at the moment it
/// is set rather than at the next render that happens to pass through here.
pub(crate) fn trim_now() {
    let limit = page_cache_limit_bytes();
    if let Ok(mut cache) = RENDERED_PAGES.lock() {
        page_cache_trim(&mut cache, limit);
    }
}

/// Render page 1 into the largest box that fits `max_width` x `max_height`
/// without changing the page's aspect ratio, and return BGRA pixels with the
/// size they were rendered at.
///
/// What was drawn is held in memory for the next hover of the same file at the same
/// size — the render opens the document and rasters the page, which is the whole of
/// what a PDF hover costs — up to `pdf_cache_mb`, and the cache is consulted before
/// anything is opened, so a hit never pays for the engine at all.
pub fn render_first_page(
    path: &Path,
    max_width: u32,
    max_height: u32,
) -> Option<(Vec<u8>, u32, u32)> {
    let key = page_cache_key(path, max_width, max_height);
    if let Some(page) = page_cache_get(&key) {
        return Some(page);
    }

    if !has_pdf_header(path) {
        return None;
    }

    let document = open_document(path)?;
    remember_opened_dimensions(path, &document);
    let (pixels, width, height) = render_opened_first_page(&document, max_width, max_height)?;
    page_cache_put(key, &pixels, width, height);

    Some((pixels, width, height))
}

/// The same, for a PDF that is already in memory.
///
/// This is the page a document renderer exported: it was written to a file and read
/// back out of it, so it never has to be opened from a path of its own — and it is
/// drawn exactly as a PDF on disk is, at whatever size the preview is shown at.
pub fn render_first_page_of_bytes(
    bytes: &[u8],
    max_width: u32,
    max_height: u32,
) -> Option<(Vec<u8>, u32, u32)> {
    if !has_pdf_header_in(bytes) {
        return None;
    }

    let document = open_document_from_bytes(bytes)?;
    render_opened_first_page(&document, max_width, max_height)
}

/// Page 1's size in DIPs, for a PDF that is already in memory.
///
/// It is read here rather than remembered: what is in memory is one document's
/// page, and the entry holding it carries its own size, so nothing about it can
/// outlive the page it was read from.
pub fn page_dimensions_of_bytes(bytes: &[u8]) -> Option<(u32, u32)> {
    if !has_pdf_header_in(bytes) {
        return None;
    }

    let document = open_document_from_bytes(bytes)?;
    let size = document.GetPage(0).ok()?.Size().ok()?;
    page_size_in_dips(size.Width, size.Height)
}

/// Page 1 of a document that has already been opened, fitted and drawn.
fn render_opened_first_page(
    document: &PdfDocument,
    max_width: u32,
    max_height: u32,
) -> Option<(Vec<u8>, u32, u32)> {
    let page = document.GetPage(0).ok()?;
    let size = page.Size().ok()?;
    let (render_width, render_height) = fit_page(size.Width, size.Height, max_width, max_height)?;

    let bytes = render_page(&page, render_width, render_height)?;
    // The page the engine encoded is read under the same limits as every other decode,
    // so what a hover can ask an allocator for is one question with one answer whatever
    // the file was.
    let mut reader = image::ImageReader::new(std::io::Cursor::new(&bytes[..]))
        .with_guessed_format()
        .ok()?;
    reader.limits(image_decode_limits());

    let image = fit_drawn_page(reader.decode().ok()?, render_width, render_height).to_rgba8();
    let (width, height) = (image.width(), image.height());
    if width == 0 || height == 0 {
        return None;
    }

    Some((opaque_bgra(image.as_raw()), width, height))
}

/// Hand the size of a document that has just been opened to the cache the measure
/// reads from, so a file opened for its render is not opened again to be measured.
///
/// A document whose first page gives no size is recorded as such, which is what
/// keeps a broken file from being parsed on every hover.
fn remember_opened_dimensions(path: &Path, document: &PdfDocument) {
    let dimensions = document
        .GetPage(0)
        .ok()
        .and_then(|page| page.Size().ok())
        .and_then(|size| page_size_in_dips(size.Width, size.Height));

    remember_page_dimensions(page_dimension_key(path), dimensions);
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

    let scale = (max_width as f32 / page_width as f32).min(max_height as f32 / page_height as f32);
    let target_width = (page_width as f32 * scale)
        .round()
        .clamp(1.0, max_width as f32) as u32;
    let target_height = (page_height as f32 * scale)
        .round()
        .clamp(1.0, max_height as f32) as u32;

    Some((target_width, target_height))
}

/// A page the engine drew, drawn into the box it was asked for.
///
/// The destination a page is rendered to is in device-independent pixels, so the
/// engine converts it with the display's own scale instead of taking it as pixels:
/// where the display a hover is on is the one the system is scaled by — a 4K monitor
/// at 150% with no other attached, say — the page comes back half again as large as
/// the box it was given, and twice it at 200%. What comes back is the frame this
/// module hands on, and the preview window is sized to the frame it is given (see
/// `render_layered_preview_at`), so a page left at that size is drawn past the edge of
/// the display the layout fitted it into.
///
/// Fitting it into the box it was asked for is a no-op where the two sizes are already
/// the same thing — a display at 100%, and a page the engine drew at the size it was
/// given — and it keeps the page inside the display at every scale. The page's own
/// shape is kept rather than being stretched into the box, and a page that is already
/// inside it is handed on untouched rather than resampled.
fn fit_drawn_page(
    image: image::DynamicImage,
    max_width: u32,
    max_height: u32,
) -> image::DynamicImage {
    let (max_width, max_height) = (max_width.max(1), max_height.max(1));

    if image.width() <= max_width && image.height() <= max_height {
        return image;
    }

    image.thumbnail(max_width, max_height)
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

/// Open a document that is already in memory.
///
/// A page a document renderer exported is the one PDF here that has no path worth
/// reading: it was written to a scratch file, read back out of it and deleted, so
/// it is handed to the engine as the bytes it is. The stream is this process's own,
/// which makes it the cheaper form rather than a dearer one — no file is opened,
/// and nothing is read twice.
fn open_document_from_bytes(bytes: &[u8]) -> Option<PdfDocument> {
    let stream = InMemoryRandomAccessStream::new().ok()?;

    let output = stream.GetOutputStreamAt(0).ok()?;
    let writer = DataWriter::CreateDataWriter(&output).ok()?;
    writer.WriteBytes(bytes).ok()?;
    writer.StoreAsync().ok()?.get().ok()?;
    // A document is read from wherever it is told to begin, and the writer leaves
    // the stream at its end.
    stream.Seek(0).ok()?;

    // The document is read out of the stream the bytes were just written into,
    // which is the same call the file form makes with a stream of its own.
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
/// byte cannot make the preview invisible. Shared with the Office renderer, which
/// draws a page of its own into the same kind of frame.
pub(crate) fn opaque_bgra(rgba: &[u8]) -> Vec<u8> {
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

    has_pdf_header_in(&probe[..read])
}

/// The same question, asked of bytes that are already in memory.
fn has_pdf_header_in(bytes: &[u8]) -> bool {
    let probe = &bytes[..bytes.len().min(PDF_HEADER_PROBE_BYTES)];

    probe
        .windows(PDF_HEADER.len())
        .any(|window| window == PDF_HEADER)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A key of this module's own, for a page that was never rendered.
    fn key(name: &str) -> PageCacheKey {
        PageCacheKey {
            path: PathBuf::from(name),
            modified: None,
            len: 0,
            width: 100,
            height: 200,
        }
    }

    /// A held page of `size` bytes.
    fn entry(size: usize, last_used: u64) -> PageCacheEntry {
        PageCacheEntry {
            pixels: vec![0u8; size],
            width: 100,
            height: 200,
            size,
            last_used,
        }
    }

    /// The page that was asked for least recently is the one that goes, and the
    /// bytes the cache reports are the bytes that are left in it.
    #[test]
    fn trims_the_page_cache_least_recently_used_first() {
        let mut cache = PageCache::default();
        cache.entries.insert(key("old"), entry(64, 1));
        cache.entries.insert(key("new"), entry(64, 2));
        cache.bytes = 128;

        page_cache_trim(&mut cache, 64);

        assert!(
            cache.entries.contains_key(&key("new")),
            "the newer page stays"
        );
        assert!(
            !cache.entries.contains_key(&key("old")),
            "the older page goes"
        );
        assert_eq!(cache.bytes, 64, "the budget is what is held");

        // A budget of nothing empties it, which is what `0 MB` means.
        page_cache_trim(&mut cache, 0);
        assert!(cache.entries.is_empty());
        assert_eq!(cache.bytes, 0);
    }

    /// The file and its version are what a page's size is read for: a file saved
    /// again — exported at another page size, say — is measured again rather than
    /// placed by the shape of the document it used to be.
    #[test]
    fn measures_a_page_again_when_the_file_is_saved_again() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-pdf-tests")
            .join("dimensions");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let path = folder.join("measured.pdf");
        std::fs::write(&path, b"one page").expect("a written file");

        let first = page_dimension_key(&path);
        assert_eq!(first, page_dimension_key(&path));

        std::fs::write(&path, b"another page, and a longer one").expect("a rewritten file");
        assert_ne!(
            first,
            page_dimension_key(&path),
            "a rewritten file is measured again"
        );

        let _ = std::fs::remove_file(&path);
    }

    /// The file, its version and the box it was drawn into are what a page is held
    /// for: any of them changing is another page.
    #[test]
    fn keys_a_page_by_the_file_its_version_and_the_box() {
        // A folder of this module's own: the tests run beside each other, and one
        // of them clearing its fixtures must not take another's with it.
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-pdf-tests")
            .join("keys");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let path = folder.join("keyed.pdf");
        std::fs::write(&path, b"one").expect("a written file");

        let first = page_cache_key(&path, 800, 600);
        assert_eq!(first, page_cache_key(&path, 800, 600));
        assert_ne!(
            first,
            page_cache_key(&path, 600, 800),
            "another box is another page"
        );

        std::fs::write(&path, b"a rewritten file").expect("a rewritten file");
        assert_ne!(
            first,
            page_cache_key(&path, 800, 600),
            "a rewritten file is another page"
        );

        let _ = std::fs::remove_file(&path);
    }

    /// A page that came back from the engine larger than it was asked for — which is
    /// what a display that is the one the system is scaled by produces, the destination
    /// being in DIPs — is drawn into the box it was asked for. The window is sized to
    /// the page, so a page left at the size the engine drew it at is drawn past the edge
    /// of the display it was fitted into.
    #[test]
    fn draws_a_page_the_engine_overscaled_into_the_box_it_was_given() {
        // A page fitted into a box of 1000 by 1400, drawn by an engine that took the
        // destination as DIPs at 150% of it, and at 200%.
        for (drawn_at, given) in [((1500, 2100), (1000, 1400)), ((2000, 2800), (1000, 1400))] {
            let page = image::DynamicImage::ImageRgba8(image::RgbaImage::from_pixel(
                drawn_at.0,
                drawn_at.1,
                image::Rgba([10, 20, 30, 255]),
            ));

            let drawn = fit_drawn_page(page, given.0, given.1);
            let (width, height) = (drawn.width(), drawn.height());

            assert!(
                width <= given.0 && height <= given.1,
                "a page drawn at {drawn_at:?} is drawn inside the box it was given, not {width} by {height}"
            );
            assert!(
                width.abs_diff(given.0) <= 1 && height.abs_diff(given.1) <= 1,
                "and it takes the size of that box rather than less of it: {width} by {height} for {given:?}"
            );
        }
    }

    /// The page's own shape is what survives the fit: a page that is over on one axis
    /// is drawn down by that axis rather than stretched into the box's shape.
    #[test]
    fn keeps_the_pages_own_shape_when_fitting_it_into_the_box() {
        let page = image::DynamicImage::ImageRgba8(image::RgbaImage::from_pixel(
            2000,
            500,
            image::Rgba([10, 20, 30, 255]),
        ));

        let drawn = fit_drawn_page(page, 1000, 1000);

        assert_eq!((drawn.width(), drawn.height()), (1000, 250));
    }

    /// A page the engine drew at the size it was asked for — every render on a display
    /// at 100% — is handed on untouched: what a page costs is the pixels the engine
    /// drew, and nothing is resampled that does not have to be.
    #[test]
    fn hands_a_page_that_is_already_inside_the_box_on_untouched() {
        let page = image::RgbaImage::from_fn(4, 3, |x, y| image::Rgba([x as u8, y as u8, 0, 255]));

        let drawn = fit_drawn_page(image::DynamicImage::ImageRgba8(page.clone()), 800, 600);

        assert_eq!((drawn.width(), drawn.height()), (4, 3));
        assert_eq!(
            drawn.to_rgba8().as_raw(),
            page.as_raw(),
            "the pixels the engine drew are the pixels that are shown"
        );
    }

    /// Page 1 of the PDFs named in `RHP_PDF_PROBE` (separated by `;`), drawn into a 1000
    /// by 1400 box through the real engine, reporting both sizes.
    ///
    /// Ignored because it needs files, and because the size the engine draws a page at
    /// is the one thing about it that depends on the display the machine is on: a display
    /// that is the one the system is scaled by is handed a page larger than the box it
    /// was given, which the fit above is what answers. This is the way to ask a machine
    /// that shows a document too large what its pages really come back at:
    /// `$env:RHP_PDF_PROBE = "C:\docs\report.pdf"`
    /// `cargo test -- --ignored --nocapture pdf_engine_probe`
    #[test]
    #[ignore = "reads the files named in RHP_PDF_PROBE"]
    fn pdf_engine_probe() {
        let Ok(list) = std::env::var("RHP_PDF_PROBE") else {
            println!("set RHP_PDF_PROBE to one or more paths, separated by ';'");
            return;
        };

        initialize_apartment();

        let (max_width, max_height) = (1000, 1400);
        for path in list
            .split(';')
            .map(str::trim)
            .filter(|path| !path.is_empty())
            .map(PathBuf::from)
        {
            println!(
                "\n--- {} ---\npage: {:?} dips\ndrawn into: {max_width} by {max_height}",
                path.display(),
                page_dimensions(&path)
            );

            let drawn = render_first_page(&path, max_width, max_height);
            let Some((pixels, width, height)) = drawn else {
                println!("nothing was drawn");
                continue;
            };

            println!("drawn: {width} by {height}");
            assert_eq!(
                pixels.len(),
                width as usize * height as usize * 4,
                "the pixels are the size the page reports"
            );
            assert!(
                width <= max_width && height <= max_height,
                "the page is drawn inside the box it was given, not {width} by {height}"
            );
        }
    }
}

use crate::config::config::{image_decode_limits, PreviewType};
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
use windows::Storage::Streams::{DataReader, IRandomAccessStream, InMemoryRandomAccessStream};
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

/// Which page of a converted book a preview is drawn from, keyed by the file and its version the
/// way a size is: a book converted again is a book to look at again.
///
/// `None` records a file that could not be opened at all, so a document that is not a PDF is not
/// parsed on every hover of it. What is *not* remembered is a book whose pages are all one colour:
/// that is an answer about the pages, and it is written down like any other.
type BookPageCache = HashMap<PageDimensionKey, Option<BookPage>>;

static BOOK_PAGES: Lazy<Mutex<BookPageCache>> = Lazy::new(|| Mutex::new(HashMap::new()));

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
/// format's own world writes: `pdfa`, the archival profile of a PDF, and `epdf`, the encapsulated
/// one, are both pages the Windows engine opens exactly as it opens a `.pdf`. Those three names are
/// the `[ebook]` list's rather than a rule of this module's — they are the half of the book kind
/// this reader draws, written down beside the comics it does not (see `ebook_formats`) — so a name
/// taken out of that list is a name this app stops drawing.
///
/// A `.ai` is one when Illustrator saved it the way it saves one by default — with `Create PDF
/// Compatible File` on, which has been the default since Illustrator 9 — because the document
/// *is* page 1 of a PDF then, and the private data the application writes beside the artwork is
/// what the OS engine reads past. Saved without that compatibility the file is PostScript,
/// which is not a page any engine here can draw, so the bytes are asked rather than believed and
/// a file that answers no is a file with no preview. That name is a drawing's rather than a book's,
/// so the answer comes from the file rather than from a list.
///
/// Only the name that needs the question pays for it: the three page spellings are answered
/// without opening anything, and a cloud placeholder is not opened to answer either, which is the
/// rule every gate in this app follows.
///
/// The list is read here because this form is asked by callers that have nothing in hand. A caller
/// that already holds the configuration asks [`is_pdf_file_in`] instead, which is the same
/// question with the list handed to it.
pub fn is_pdf_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| is_pdf_file_in(path, &config.ebook_extensions))
        .unwrap_or(false)
}

/// The same question asked of a list the caller already holds, which is the form the hook asks it
/// in: it resolves a hover with the configuration in hand, and every list it consults it consults
/// through that copy.
///
/// A question that went and read the configuration again would wait on a lock the same thread is
/// already holding — and a lock taken twice on one thread never comes back, so the thread that
/// took it never returns to anything else it was doing. The hook's thread is the one that watches
/// Explorer, and the whole app's other threads queue up behind it on the same lock, which is what
/// the first PDF a pointer came to rest on used to do to it (see `explorer_hook::is_media_file`).
pub fn is_pdf_file_in(path: &Path, extensions: &[String]) -> bool {
    if crate::formats::ebook_formats::matches_page_name(path, extensions) {
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

/// The same, for the page a book is previewed from.
fn remember_book_page(key: PageDimensionKey, page: Option<BookPage>) {
    if let Ok(mut cache) = BOOK_PAGES.lock() {
        if !cache.contains_key(&key) && cache.len() >= PAGE_DIMENSION_CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.insert(key, page);
    }
}

/// Render page 1 into the largest box that fits `max_width` x `max_height`
/// without changing the page's aspect ratio, and return BGRA pixels with the
/// size they were rendered at.
///
/// Nothing is held on this side: the render opens the document and rasters the page, which is
/// the whole of what a PDF hover costs, and it is paid again on the next hover of the same
/// file. What was drawn is a screenful of pixels rather than the document, and holding those
/// costs more than drawing them again (see `Performance → Cache` in the tray).
pub fn render_first_page(
    path: &Path,
    max_width: u32,
    max_height: u32,
) -> Option<(Vec<u8>, u32, u32)> {
    if !has_pdf_header(path) {
        return None;
    }

    let document = open_document(path)?;
    remember_opened_dimensions(path, &document);
    render_opened_first_page(&document, max_width, max_height)
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

    decode_rendered_page(&bytes, render_width, render_height)
}

/// How many of a converted book's opening pages are looked at before its own first page is kept
/// whatever it holds.
const BOOK_PAGES_MAX: u32 = 8;

/// How large a page is drawn to answer whether it says anything. Small, because the question is
/// only whether the page holds more than one colour, and what it costs is paid before every
/// preview of the page it is asked about.
const BOOK_PAGE_PROBE: u32 = 48;

/// The page a converted book is previewed from: which page of the document it is, and that page's
/// own size in DIPs.
///
/// The size is that page's rather than the document's first, because the two are not always the
/// same and it is this page that is drawn (see [`book_page`]).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub struct BookPage {
    pub index: u32,
    pub size: (u32, u32),
}

/// The page of a converted book a preview is drawn from: the first of its opening pages that is
/// not one flat colour, with that page's own size.
///
/// A book's first page is very often not a page of the book. What an EPub, a Kindle file and a
/// Mobipocket one all put first is the cover, and a cover is as likely to be one colour as it is to
/// be artwork: the quick start guide Calibre ships inside its own installation carries a cover that
/// is a single pixel, stretched over a whole page, so what a hover on one of those books shows is a
/// rectangle — which is honestly the book's first page, and says nothing whatever about the book.
/// What is asked of each page in turn is therefore whether it holds more than one colour, and the
/// first one that does is the page the preview is made from.
///
/// What is *not* skipped is a page that has anything at all on it: a title page, a page of text, a
/// photograph, a page with one line drawn on it are all answers to "what is this book", and only a
/// page that is a single colour is not. The walk stops after [`BOOK_PAGES_MAX`] pages, so a book
/// whose opening pages are all one colour — a scan of blank leaves, a cover that is a colour
/// swatch — is still previewed from its own first page rather than from nothing, and a document
/// that cannot be walked at all is answered with page one exactly as it always was.
///
/// What this costs is a small render per page looked at, which is why the answer is held between
/// asks: a hover asks for this twice — the layout to place the preview, the loader to draw it —
/// and a page that has been looked at once is not looked at again (see `BOOK_PAGES`).
pub fn book_page(path: &Path) -> Option<BookPage> {
    let key = page_dimension_key(path);
    if let Ok(cache) = BOOK_PAGES.lock() {
        if let Some(cached) = cache.get(&key) {
            return *cached;
        }
    }

    let chosen = probe_book_page(path);
    remember_book_page(key, chosen);

    chosen
}

fn probe_book_page(path: &Path) -> Option<BookPage> {
    if !has_pdf_header(path) {
        return None;
    }

    let document = open_document(path)?;

    // The first page's size is what a PDF is measured by, and this is the same opening: what is
    // read here is handed to that cache rather than left for it to find a second time.
    remember_opened_dimensions(path, &document);

    let mut first: Option<BookPage> = None;

    for index in 0..BOOK_PAGES_MAX {
        let Ok(page) = document.GetPage(index) else {
            break;
        };
        let Some(size) = page
            .Size()
            .ok()
            .and_then(|size| page_size_in_dips(size.Width, size.Height))
        else {
            break;
        };

        let candidate = BookPage { index, size };

        if first.is_none() {
            first = Some(candidate);
        }

        if !is_flat_page(&page) {
            return Some(candidate);
        }
    }

    first
}

/// Whether a page holds one colour and nothing else: nothing but the colour its own background is,
/// which is a page with nothing on it.
///
/// The page is drawn small rather than read for its content: a PDF's content is a program, and
/// what it draws is the answer — a page whose every pixel comes out the same is a page that says
/// nothing, whatever the program that drew it looks like. A page that cannot be drawn at all is
/// answered with `false`, which keeps it: a page this side cannot look at is not a page it may
/// decide the book is without.
fn is_flat_page(page: &PdfPage) -> bool {
    let Some(bytes) = render_page(page, BOOK_PAGE_PROBE, BOOK_PAGE_PROBE) else {
        return false;
    };
    let Ok(image) = image::load_from_memory(&bytes) else {
        return false;
    };

    is_one_colour(&image.to_rgba8())
}

/// Whether a drawn page holds one colour and nothing else.
fn is_one_colour(image: &image::RgbaImage) -> bool {
    let mut pixels = image.pixels();
    let Some(first) = pixels.next() else {
        return false;
    };

    pixels.all(|pixel| pixel == first)
}

/// The page a converted book is drawn from, fitted into the box the caller asks for.
///
/// What is drawn is the page [`book_page`] chose rather than the document's first, and it is drawn
/// into the box the caller names rather than at the page's own size — so the frame that comes back
/// is the shape the layout measured, made of the page the book is previewed from.
pub fn render_book_page(
    path: &Path,
    max_width: u32,
    max_height: u32,
) -> Option<(Vec<u8>, u32, u32)> {
    let book = book_page(path)?;

    let document = open_document(path)?;
    let page = document.GetPage(book.index).ok()?;

    let (render_width, render_height) = fit_page(
        book.size.0 as f32,
        book.size.1 as f32,
        max_width,
        max_height,
    )?;

    let bytes = render_page(&page, render_width, render_height)?;

    decode_rendered_page(&bytes, render_width, render_height)
}

/// The page the engine encoded, read into the frame this module hands on.
///
/// The page is read under the same limits as every other decode of this app, so what a hover can
/// ask an allocator for is one question with one answer whatever the file was, and the box it was
/// drawn into is the one this side asked for — a page left at the size the display's own scale
/// would make of it is a page drawn past the edge of the display the layout fitted it into (see
/// `fit_drawn_page`).
fn decode_rendered_page(
    bytes: &[u8],
    render_width: u32,
    render_height: u32,
) -> Option<(Vec<u8>, u32, u32)> {
    let mut reader = image::ImageReader::new(std::io::Cursor::new(bytes))
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

    /// The page names are read out of the list a caller hands in — the form the hook asks the
    /// question in, holding the configuration the list lives in — and a name that has left the
    /// list is a name this app stops drawing, the list being the reader's rather than this
    /// module's (see `is_pdf_file_in`).
    #[test]
    fn reads_a_page_name_out_of_the_list_it_is_given() {
        let list = crate::formats::ebook_formats::sanitize_ebook_extensions(
            crate::formats::ebook_formats::DEFAULT_EBOOK_EXTENSIONS,
        );

        for name in ["report.pdf", "archived.pdfa", "encapsulated.epdf"] {
            assert!(
                is_pdf_file_in(Path::new(name), &list),
                "`{name}` is a page the engine opens"
            );
        }

        assert!(
            !is_pdf_file_in(Path::new("notes.txt"), &list),
            "a name the list does not hold is not a page"
        );

        let without_the_page_names = crate::formats::ebook_formats::sanitize_ebook_extensions(
            "cbc,cbr,cbz",
        );
        assert!(
            !is_pdf_file_in(Path::new("report.pdf"), &without_the_page_names),
            "and a name taken out of the list is answered by what it is rather than by a list"
        );
    }

    /// An Illustrator document is a page when it is one and the list has nothing to say about it:
    /// the name is a drawing's, so the file's own header is the whole of the answer.
    #[test]
    fn reads_an_illustrator_document_by_its_own_header() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-pdf-tests")
            .join("illustrator");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let list = crate::formats::ebook_formats::sanitize_ebook_extensions(
            crate::formats::ebook_formats::DEFAULT_EBOOK_EXTENSIONS,
        );

        let compatible = folder.join("artwork.ai");
        std::fs::write(&compatible, b"%PDF-1.7\none page").expect("a written file");
        assert!(
            is_pdf_file_in(&compatible, &list),
            "a document saved with PDF compatibility is page one of a PDF"
        );

        let postscript = folder.join("drawn.ai");
        std::fs::write(&postscript, b"%!PS-Adobe-3.0\n").expect("a written file");
        assert!(
            !is_pdf_file_in(&postscript, &list),
            "and one saved without it is PostScript, which is not a page any engine here draws"
        );

        let _ = std::fs::remove_file(&compatible);
        let _ = std::fs::remove_file(&postscript);
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

    /// A page with one colour on it is a page that says nothing, and a page with anything else on
    /// it — one line, one dot, one shade beside another — is a page of the book.
    ///
    /// It is the whole of what a converted book's opening pages are judged by, so what it is worth
    /// saying is where the line is: one colour is not a page, and every other page is.
    #[test]
    fn reads_a_page_of_one_colour_as_a_page_with_nothing_on_it() {
        let flat = image::RgbaImage::from_pixel(64, 64, image::Rgba([130, 41, 45, 255]));
        assert!(
            is_one_colour(&flat),
            "the cover the quick start guide ships is one pixel of one colour, stretched over a page"
        );

        let blank = image::RgbaImage::from_pixel(64, 64, image::Rgba([255, 255, 255, 255]));
        assert!(is_one_colour(&blank), "and a blank page is one colour too");

        let mut lined = flat.clone();
        lined.put_pixel(32, 32, image::Rgba([129, 41, 45, 255]));
        assert!(
            !is_one_colour(&lined),
            "a page with one pixel of another shade is a page with something on it"
        );

        assert!(
            !is_one_colour(&image::RgbaImage::from_fn(64, 64, |x, _| image::Rgba([
                x as u8, 0, 0, 255
            ]))),
            "and so is a page drawn in a gradient"
        );

        assert!(
            !is_one_colour(&image::RgbaImage::new(0, 0)),
            "a page with no pixels at all is not a page to judge"
        );
    }

    /// Page 1 of the PDFs named in `RHP_PDF_PROBE` (separated by `;`), drawn into a 1000
    /// by 1400 box through the real engine, reporting both sizes — and the page a converted book
    /// would be previewed from, which is the first of its opening pages that holds more than one
    /// colour (see `book_page`).
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
                "\n--- {} ---\npage: {:?} dips\ndrawn into: {max_width} by {max_height}\nbook page: {:?}",
                path.display(),
                page_dimensions(&path),
                book_page(&path)
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

            let Some((pixels, width, height)) = render_book_page(&path, max_width, max_height)
            else {
                println!("no page of the book was drawn");
                continue;
            };

            println!("book drawn: {width} by {height}");
            assert_eq!(
                pixels.len(),
                width as usize * height as usize * 4,
                "the pixels of the book's page are the size it reports"
            );
        }
    }
}

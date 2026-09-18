//! SVG previews: a vector document, drawn at the size the preview is shown at.
//!
//! An SVG is not a picture to a decoder — it is a document of shapes, gradients,
//! texts and images — so this is the PDF path's shape rather than the image path's.
//! The document's own size is read from it for the layout to place the preview from,
//! and the frame is then drawn at the box that comes out, at the display's scale,
//! rather than decoded at some native size and resampled into it: a vector has no
//! native size, and drawing one at the size it is shown is also what keeps an enlarged
//! preview sharp (see PDF Previews, whose measure-then-render this follows).
//!
//! A parsed document is held in memory between hovers, and a drawn frame is not.
//!
//! A hover meets the same document several times — the layout measures it for the box,
//! the loader measures it again to scale that box, and the renderer parses it once more
//! to draw it — and the parse is the cheap half of a vector to keep: a tree of paths and
//! styles is small next to the frame it draws, and it is what a second hover of the same
//! file would otherwise pay for all over again. So the parse is held, keyed by the file
//! and the version of it that was read, and the frame is not: a frame is as large as the
//! display's room, it can be drawn again from a tree that is already in hand, and an SVG
//! frame has no business in the budget `image_cache_mb` counts for photos. Nothing is
//! written to disk either — what is held lives in this process and goes when it does.

use once_cell::sync::Lazy;
use resvg::tiny_skia;
use resvg::usvg;
use std::collections::HashMap;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::SystemTime;

/// The bytes a gzipped document starts with, whatever it is called.
const GZIP_MAGIC: [u8; 2] = [0x1f, 0x8b];

/// The viewport a document with no size of its own is assumed to have, which is what
/// every browser assumes for a replaced element: 300 by 150 CSS pixels, the SVG
/// specification's default. usvg's own default is a square, and a document that is
/// a banner should not be previewed as a box.
const DEFAULT_VIEWPORT_WIDTH: f32 = 300.0;
const DEFAULT_VIEWPORT_HEIGHT: f32 = 150.0;

/// The options every parse starts from, with the machine's fonts read once.
///
/// usvg draws text with the database it is handed and has none of its own, so a
/// document with text in it is drawn without any until the fonts have been read. That
/// read is a scan of the installed font files — not cheap, and not something every
/// hover should pay for again — so it happens on the first SVG preview and the
/// database is kept, the way the text preview's parsed themes are kept.
static BASE_OPTIONS: Lazy<usvg::Options<'static>> = Lazy::new(|| {
    let mut options = usvg::Options::default();
    options.fontdb_mut().load_system_fonts();

    if let Some(size) = usvg::Size::from_wh(DEFAULT_VIEWPORT_WIDTH, DEFAULT_VIEWPORT_HEIGHT) {
        options.default_size = size;
    }

    options
});

/// Whether a file's name is an SVG's — including the gzipped form, which is the same
/// document under a name that says it is compressed.
///
/// The question is what a file is called and nothing else, exactly as the other
/// format gates ask it: the parser has the last word on whether a file is a document
/// at all, the way the decoder has it for a picture, and a file that turns out not to
/// be one is answered with no preview rather than with a guessed size.
pub fn is_svg_file(path: &Path) -> bool {
    matches!(
        path.extension()
            .and_then(|extension| extension.to_str())
            .map(str::to_lowercase)
            .as_deref(),
        Some("svg") | Some("svgz")
    )
}

/// The size the document asks to be drawn at, for the layout to place a preview from.
///
/// The size is the root element's own: its `width` and `height`, with a `viewBox`
/// answering where they are relative or missing — a document that says only
/// `width="100%"` is as wide as its viewBox — which is the size a viewer draws it at
/// rather than the extent of what it happens to contain. A document that will not
/// parse, or whose size is not a size at all, reports nothing, and the hover shows no
/// preview for it the way one onto a picture that will not decode does.
pub fn measure(path: &Path) -> Option<(u32, u32)> {
    tree_size(document(path)?.as_ref())
}

/// Draw the document into the largest box that fits `max_width` by `max_height`
/// without changing its shape, and return BGRA pixels with the size they were drawn
/// at.
///
/// The box is the one the layout computed from `measure`, so the fit here is that
/// same rule's rounding rather than a second opinion about it: what comes back is the
/// size the preview was planned at, letterboxed where rounding left the box a pixel
/// off the document's shape.
pub fn render(
    path: &Path,
    max_width: u32,
    max_height: u32,
    cancel: Option<&AtomicBool>,
) -> Option<(Vec<u8>, u32, u32)> {
    if is_cancelled(cancel) {
        return None;
    }

    let tree = document(path)?;

    if is_cancelled(cancel) {
        return None;
    }

    draw(tree.as_ref(), max_width, max_height)
}

/// Draw a document that is written out as text, which is what a frame of an animated
/// one is: the animation is applied by writing the document again with its values at
/// that moment, so what is drawn is a document rather than a tree that was edited.
pub fn render_text(text: &str, max_width: u32, max_height: u32) -> Option<(Vec<u8>, u32, u32)> {
    let tree = usvg::Tree::from_data_nested(text.as_bytes(), &parse_options()).ok()?;

    draw(&tree, max_width, max_height)
}

/// A parsed document drawn into the largest box that fits the room, drawn at the size
/// it is shown at rather than resampled into it.
fn draw(tree: &usvg::Tree, max_width: u32, max_height: u32) -> Option<(Vec<u8>, u32, u32)> {
    let (width, height) = tree_size(tree)?;
    let (target_width, target_height, scale) = fit(width, height, max_width, max_height)?;

    let mut pixmap = tiny_skia::Pixmap::new(target_width, target_height)?;
    resvg::render(
        tree,
        tiny_skia::Transform::from_scale(scale, scale),
        &mut pixmap.as_mut(),
    );

    Some((
        bgra(pixmap.take_demultiplied()),
        target_width,
        target_height,
    ))
}

/// A document's bytes, read for the parser.
///
/// The nested parse that is handed these bytes ignores an `<image>` element linking to
/// an external file; see `document`.
fn read_document(path: &Path) -> Option<Vec<u8>> {
    let bytes = std::fs::read(path).ok()?;

    if bytes.starts_with(&GZIP_MAGIC) {
        let mut document = Vec::new();
        flate2::read::GzDecoder::new(bytes.as_slice())
            .read_to_end(&mut document)
            .ok()?;

        return Some(document);
    }

    Some(bytes)
}

/// How many parsed documents are held at once.
///
/// What a pointer meets again is the folder it is in, so the handful of documents under
/// it are the ones worth keeping; a document that falls out is parsed again the next
/// time it is hovered, which is the cost this cache exists to save rather than one it
/// turns into a failure.
const MAX_HELD_DOCUMENTS: usize = 16;

/// The file's modification time and length: what says a file is not the one that was
/// parsed last time.
#[derive(Clone, PartialEq, Eq, Hash)]
struct FileVersion {
    modified: Option<SystemTime>,
    len: u64,
}

/// What a held parse is only valid for: the file and the version of it that was read.
#[derive(Clone, PartialEq, Eq, Hash)]
struct DocumentKey {
    path: PathBuf,
    version: FileVersion,
}

/// A parsed document: the tree the renderer draws, the text it was parsed from — which
/// is what the animation pass writes out again with a moment's values in it — and
/// whether it says it moves at all, which is what decides between the engine and this
/// app's own reader.
struct Parsed {
    tree: Arc<usvg::Tree>,
    source: Arc<str>,
    moves: bool,
}

/// A held parse: the document, or the fact that this version of the file is not one.
///
/// A file that is not a document is held as that, rather than as nothing held at all,
/// so a hover onto it costs nothing after the first instead of a read and a failed
/// parse per pass — the same way the PDF page-size cache remembers a file it could not
/// open. What says the answer is still good is the version in the key: a file rewritten
/// is a different key and is read again.
enum Held {
    Parsed(Parsed),
    NotADocument,
}

/// A held parse and when it was last asked for. The stamp is a counter rather than a
/// clock, so the order documents are dropped in cannot be changed by the system clock
/// moving.
struct HeldDocument {
    doc: Held,
    last_used: u64,
}

#[derive(Default)]
struct DocumentCache {
    entries: HashMap<DocumentKey, HeldDocument>,
    tick: u64,
}

/// The parsed documents held between hovers.
///
/// The lock is not held across a parse: the layout measures on the preview thread while
/// a load worker may be drawing the document it was handed, and neither has anything to
/// gain from waiting for the other to finish reading a file. Two threads that parse the
/// same document at the same moment both get a tree and the second one stored is the one
/// kept, which costs a duplicate parse and nothing else.
static DOCUMENTS: Lazy<Mutex<DocumentCache>> = Lazy::new(|| Mutex::new(DocumentCache::default()));

/// The document for `path`, parsed from the file once and held from then on.
///
/// What is parsed is the *nested* form, which ignores an `<image>` element linking to
/// an external file: a document cannot use a hover to have a file beside it — or anywhere
/// else on the machine — opened and drawn. An embedded `data:` image is part of the
/// document and is drawn; a link out of it is not followed. Nothing else the parser can
/// reach is a file either: text is drawn from the font database read once above rather
/// than from a font a document names, and a document's own script is not run at all,
/// because this is a renderer and not a viewer.
fn document(path: &Path) -> Option<Arc<usvg::Tree>> {
    parsed(path).map(|parsed| parsed.tree)
}

/// The document's bytes as the text they are, for the animation pass that writes them
/// out again with a moment's values in them.
pub fn source(path: &Path) -> Option<Arc<str>> {
    parsed(path).map(|parsed| parsed.source)
}

/// Whether the document says it moves at all, which is asked of every hover: is this
/// a document for the engine, or one to draw here? It is answered once per version of
/// the file rather than per hover, because it costs a parse of the document's own XML
/// and the answer cannot change while the file does not.
pub fn moves(path: &Path) -> bool {
    parsed(path).map(|parsed| parsed.moves).unwrap_or(false)
}

/// The held parse of `path`, reading and parsing the file when it is not held.
fn parsed(path: &Path) -> Option<Parsed> {
    let key = DocumentKey {
        path: path.to_path_buf(),
        version: file_version(path),
    };

    if let Some(held) = document_cache_get(&key) {
        return match held {
            Held::Parsed(parsed) => Some(parsed),
            Held::NotADocument => None,
        };
    }

    let parsed = read_document(path).and_then(|bytes| {
        let source: Arc<str> = Arc::from(String::from_utf8(bytes).ok()?);
        let tree = usvg::Tree::from_data_nested(source.as_bytes(), &parse_options())
            .ok()
            .map(Arc::new)?;
        let moves = roxmltree::Document::parse(source.as_ref())
            .map(|document| crate::svg_animation::declares_animation(&document))
            .unwrap_or(false);

        Some(Parsed {
            tree,
            source,
            moves,
        })
    });

    document_cache_put(&key, &parsed);

    parsed
}

/// What is held for a version of a file, when anything is.
fn document_cache_get(key: &DocumentKey) -> Option<Held> {
    let mut cache = DOCUMENTS.lock().ok()?;
    cache.tick += 1;
    let tick = cache.tick;

    let held = cache.entries.get_mut(key)?;
    held.last_used = tick;

    Some(match &held.doc {
        Held::Parsed(parsed) => Held::Parsed(Parsed {
            tree: Arc::clone(&parsed.tree),
            source: Arc::clone(&parsed.source),
            moves: parsed.moves,
        }),
        Held::NotADocument => Held::NotADocument,
    })
}

/// Hold the parse of a file — or the fact that there is none — dropping the least
/// recently used document once the cap is passed.
fn document_cache_put(key: &DocumentKey, parsed: &Option<Parsed>) {
    let Ok(mut cache) = DOCUMENTS.lock() else {
        return;
    };

    cache.tick += 1;
    let tick = cache.tick;

    cache.entries.insert(
        key.clone(),
        HeldDocument {
            doc: match parsed {
                Some(parsed) => Held::Parsed(Parsed {
                    tree: Arc::clone(&parsed.tree),
                    source: Arc::clone(&parsed.source),
                    moves: parsed.moves,
                }),
                None => Held::NotADocument,
            },
            last_used: tick,
        },
    );

    while cache.entries.len() > MAX_HELD_DOCUMENTS {
        let Some(oldest) = cache
            .entries
            .iter()
            .min_by_key(|(_, held)| held.last_used)
            .map(|(key, _)| key.clone())
        else {
            break;
        };

        cache.entries.remove(&oldest);
    }
}

fn file_version(path: &Path) -> FileVersion {
    match std::fs::metadata(path) {
        Ok(metadata) => FileVersion {
            modified: metadata.modified().ok(),
            len: metadata.len(),
        },
        Err(_) => FileVersion {
            modified: None,
            len: 0,
        },
    }
}

/// The options one parse runs with: everything default except the fonts, which are
/// the shared database rather than a database of its own.
fn parse_options() -> usvg::Options<'static> {
    usvg::Options {
        fontdb: Arc::clone(&BASE_OPTIONS.fontdb),
        ..Default::default()
    }
}

/// A tree's own size in whole pixels.
///
/// A document is measured in user units, which are neither whole numbers nor
/// necessarily sane, so what comes back is rounded up to at least one pixel rather
/// than truncated to nothing; a size that is not a size at all — zero, negative, or
/// not a number — is refused, because that is a document nothing could draw and there
/// is nothing for a layout to place.
fn tree_size(tree: &usvg::Tree) -> Option<(u32, u32)> {
    let size = tree.size();
    let (width, height) = (size.width(), size.height());

    if !width.is_finite() || !height.is_finite() || width <= 0.0 || height <= 0.0 {
        return None;
    }

    Some((
        width.round().max(1.0) as u32,
        height.round().max(1.0) as u32,
    ))
}

/// The box the document is drawn into, and the scale that fills it: the largest size
/// that fits the room without changing the document's own shape, so a document that
/// is not the shape the layout assumed is letterboxed rather than stretched.
///
/// The rounding and the clamp are the layout's own, which is what keeps the frame the
/// size the window was given.
fn fit(width: u32, height: u32, max_width: u32, max_height: u32) -> Option<(u32, u32, f32)> {
    let scale = (max_width as f32 / width as f32).min(max_height as f32 / height as f32);

    if !scale.is_finite() || scale <= 0.0 {
        return None;
    }

    let target_width = (width as f32 * scale)
        .round()
        .clamp(1.0, max_width.max(1) as f32) as u32;
    let target_height = (height as f32 * scale)
        .round()
        .clamp(1.0, max_height.max(1) as f32) as u32;

    Some((target_width, target_height, scale))
}

/// Straight-alpha BGRA, which is the frame every preview arrives in: the renderer
/// produces premultiplied RGBA, its own format, and the composition that puts a frame
/// on screen premultiplies for itself.
fn bgra(rgba: Vec<u8>) -> Vec<u8> {
    let mut pixels = rgba;

    for pixel in pixels.chunks_exact_mut(4) {
        pixel.swap(0, 2);
    }

    pixels
}

fn is_cancelled(cancel: Option<&AtomicBool>) -> bool {
    cancel
        .map(|flag| flag.load(Ordering::Acquire))
        .unwrap_or(false)
}

#[cfg(test)]
mod tests {
    use super::*;
    use flate2::write::GzEncoder;
    use flate2::Compression;
    use std::io::Write;
    use std::path::PathBuf;

    /// A file of this module's own folder: the tests run beside each other, and one of
    /// them clearing its fixtures must not take another's with it.
    fn fixture(name: &str, contents: &[u8]) -> PathBuf {
        let folder = std::env::temp_dir().join("rust-hover-preview-svg-tests");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let path = folder.join(name);
        std::fs::write(&path, contents).expect("a written file");
        path
    }

    #[test]
    fn claims_the_names_an_svg_goes_by() {
        for name in ["drawing.svg", "drawing.SVG", "drawing.svgz"] {
            assert!(is_svg_file(Path::new(name)), "{name}");
        }

        for name in ["drawing.png", "drawing.svg.txt", "svg"] {
            assert!(!is_svg_file(Path::new(name)), "{name}");
        }
    }

    #[test]
    fn measures_the_size_the_document_asks_for() {
        let path = fixture(
            "sized.svg",
            br#"<svg xmlns="http://www.w3.org/2000/svg" width="40" height="24"/>"#,
        );

        assert_eq!(measure(&path), Some((40, 24)));
        let _ = std::fs::remove_file(&path);
    }

    /// A document rewritten in place is parsed again: the version of the file is part
    /// of what a held parse is valid for.
    #[test]
    fn a_revised_document_is_parsed_again() {
        let path = fixture(
            "revised.svg",
            br#"<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"/>"#,
        );
        assert_eq!(measure(&path), Some((10, 10)));

        // Longer as well as different, so the version changes whatever resolution the
        // filesystem's clock turns out to keep.
        std::fs::write(
            &path,
            br#"<svg xmlns="http://www.w3.org/2000/svg" width="120" height="80"/>"#,
        )
        .expect("a rewritten document");

        assert_eq!(measure(&path), Some((120, 80)));
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn measures_a_relative_size_by_its_view_box() {
        let path = fixture(
            "viewbox.svg",
            br#"<svg xmlns="http://www.w3.org/2000/svg" width="100%" height="100%" viewBox="0 0 32 8"/>"#,
        );

        assert_eq!(measure(&path), Some((32, 8)));
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn draws_at_the_size_the_layout_asked_for() {
        let path = fixture(
            "square.svg",
            br##"<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10" fill="#ff0000"/></svg>"##,
        );

        let (pixels, width, height) = render(&path, 64, 64, None).expect("a drawn document");
        assert_eq!(
            (width, height),
            (64, 64),
            "drawn to fill the box it was given"
        );
        // BGRA, straight alpha: an opaque red pixel is blue 0, green 0, red 255.
        assert_eq!(&pixels[..4], &[0, 0, 255, 255]);
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn keeps_the_documents_shape_inside_the_box() {
        let path = fixture(
            "wide.svg",
            br#"<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50"/>"#,
        );

        let (_, width, height) = render(&path, 200, 400, None).expect("a drawn document");
        assert_eq!(
            (width, height),
            (200, 100),
            "the box is taller than the shape"
        );
        let _ = std::fs::remove_file(&path);
    }

    /// Text is drawn with the machine's fonts, which have to have been read for it to
    /// be drawn at all: a database with nothing in it leaves the text out, and a
    /// document of nothing but text comes back blank.
    #[test]
    fn draws_text_with_the_machines_fonts() {
        let path = fixture(
            "text.svg",
            br#"<svg xmlns="http://www.w3.org/2000/svg" width="200" height="60"><text x="10" y="40" font-family="Arial" font-size="30">Aa</text></svg>"#,
        );

        let (pixels, _, _) = render(&path, 200, 60, None).expect("a drawn document");
        assert!(
            pixels.chunks_exact(4).any(|pixel| pixel[3] != 0),
            "something was drawn where the text is"
        );
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn a_gzipped_document_is_the_same_document() {
        let mut encoder = GzEncoder::new(Vec::new(), Compression::default());
        encoder
            .write_all(br#"<svg xmlns="http://www.w3.org/2000/svg" width="20" height="10"/>"#)
            .expect("a written document");
        let path = fixture(
            "compressed.svgz",
            &encoder.finish().expect("a finished stream"),
        );

        assert_eq!(measure(&path), Some((20, 10)));
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn a_file_that_is_not_a_document_reports_nothing() {
        let path = fixture("not-a-document.svg", b"just some bytes");

        assert_eq!(measure(&path), None);
        assert_eq!(render(&path, 64, 64, None), None);
        let _ = std::fs::remove_file(&path);
    }
}

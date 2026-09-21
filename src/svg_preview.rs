//! SVG previews: whether a file is a document, and the size it asks to be drawn at.
//!
//! An SVG is not a picture to a decoder — it is a document of shapes, gradients, texts
//! and images — and this app does not draw one. What draws a document is the browser
//! engine that is already on the machine, in a window of its own, whether the document
//! is still or moving: see `webview_preview`. What is left here is the two questions the
//! rest of the app asks before it hands one over.
//!
//! The first is the name, asked the way every other format gate asks it. The renderer
//! has the last word on whether a file is a document at all — that renderer is the
//! engine now — and a file it turns out not to be is answered with no preview rather
//! than with a guessed size.
//!
//! The second is the size the document asks to be drawn at, which is the root element's
//! own: its `width` and `height`, with a `viewBox` answering where they are relative or
//! missing — a document that says only `width="100%"` is as wide as its viewBox — which
//! is the size a viewer draws it at rather than the extent of what it happens to
//! contain. The layout places the preview from it, the way it places a PDF page from the
//! size the PDF engine reports, and the box that comes out is the box the engine's
//! window is given: its page draws the document as an image that fills that box, so a
//! document is drawn at the size the layout planned rather than at the size it asked for
//! (see `webview_preview::frame_page`).
//!
//! Nothing is drawn and nothing is held decoded: a measurement is a read of the file and
//! a parse of it, and what is kept between hovers is the answer. The layout measures a
//! document every time it is hovered, and a pointer swept back and forth over a folder
//! meets the same files again, so the answers are held, keyed by the file and the
//! version of it that was read. Nothing is written to disk either — what is held lives
//! in this process and goes when it does.

use crate::config::{decode_budget_bytes, read_within_budget};
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::SystemTime;

/// The bytes a gzipped document starts with, whatever it is called.
const GZIP_MAGIC: [u8; 2] = [0x1f, 0x8b];

/// The viewport a document with no size of its own is assumed to have, which is what
/// every browser assumes for a replaced element: 300 by 150 CSS pixels, the SVG
/// specification's default. A document that is a banner should not be previewed as a
/// box.
const DEFAULT_VIEWPORT_WIDTH: f32 = 300.0;
const DEFAULT_VIEWPORT_HEIGHT: f32 = 150.0;

/// Whether a file's name is an SVG's — including the gzipped form, which is the same
/// document under a name that says it is compressed.
///
/// The question is what a file is called and nothing else, exactly as the other format
/// gates ask it: the renderer has the last word on whether a file is a document at all,
/// the way the decoder has it for a picture, and a file that turns out not to be one is
/// answered with no preview rather than with a guessed size.
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
/// The size is the root element's own — see this module's documentation for what that
/// means, and for what a document that answers neither its own `width` and `height` nor
/// a `viewBox` is measured at. A file that will not parse as XML, or whose root is not
/// an `svg`, reports nothing, and the hover shows no preview for it the way one onto a
/// picture that will not decode does.
pub fn measure(path: &Path) -> Option<(u32, u32)> {
    let key = DocumentKey {
        path: path.to_path_buf(),
        version: file_version(path),
    };

    match held(&key) {
        Some(Held::Measured(width, height)) => return Some((width, height)),
        Some(Held::NotADocument) => return None,
        None => {}
    }

    let size = read_document(path).and_then(|bytes| measure_bytes(&bytes));
    hold(&key, size);

    size
}

/// A document's bytes, read for the parse.
///
/// A document is read and parsed whole rather than capped at a size of its own — it is a
/// document, and what it costs is a parse proportional to it — so the budget every other
/// file is read under is what keeps "whole" from meaning any size at all. The inflated
/// document counts against it too, which is what a gzipped document needs: a `.svgz` is
/// a few kilobytes that can inflate to a great many.
fn read_document(path: &Path) -> Option<Vec<u8>> {
    let bytes = read_within_budget(path)?;

    if bytes.starts_with(&GZIP_MAGIC) {
        let budget = decode_budget_bytes();
        let mut document = Vec::new();
        flate2::read::GzDecoder::new(bytes.as_slice())
            .take(budget + 1)
            .read_to_end(&mut document)
            .ok()?;

        if document.len() as u64 > budget {
            return None;
        }

        return Some(document);
    }

    Some(bytes)
}

/// The size a document's own text asks for.
///
/// The root element's `width` and `height` are the answer where both of them resolve to
/// a length; a document that names neither of them, or names one in a unit that cannot
/// be resolved without knowing the viewport it is being drawn in, is answered by its
/// `viewBox` — which is the area it draws in rather than a size it asks to be drawn at —
/// and by the default viewport where it has neither.
fn measure_bytes(bytes: &[u8]) -> Option<(u32, u32)> {
    let text = std::str::from_utf8(bytes).ok()?;
    let document = roxmltree::Document::parse(text).ok()?;
    let root = document.root().first_element_child()?;

    if !root.tag_name().name().eq_ignore_ascii_case("svg") {
        return None;
    }

    let named = (
        root.attribute("width").and_then(length),
        root.attribute("height").and_then(length),
    );

    if let (Some(width), Some(height)) = named {
        return Some((whole_pixels(width), whole_pixels(height)));
    }

    if let Some((width, height)) = root.attribute("viewBox").and_then(view_box_size) {
        return Some((whole_pixels(width), whole_pixels(height)));
    }

    Some((
        whole_pixels(DEFAULT_VIEWPORT_WIDTH),
        whole_pixels(DEFAULT_VIEWPORT_HEIGHT),
    ))
}

/// A length as CSS pixels: the number a document wrote, in the unit it wrote it in.
///
/// A length that is not a positive number is not a length — a size of zero or less is a
/// document nothing could place — and a document that writes one is answered the way one
/// that writes nothing at all is.
fn length(value: &str) -> Option<f32> {
    let value = value.trim();
    let bytes = value.as_bytes();

    let mut end = 0;
    while end < bytes.len()
        && (bytes[end].is_ascii_digit()
            || bytes[end] == b'.'
            || (end == 0 && (bytes[end] == b'+' || bytes[end] == b'-')))
    {
        end += 1;
    }

    // An exponent, and only where a digit follows it: `1e3px` is one number, while `10em`
    // is ten of a unit this cannot resolve.
    if end < bytes.len() && (bytes[end] == b'e' || bytes[end] == b'E') {
        let mut exponent = end + 1;
        if exponent < bytes.len() && (bytes[exponent] == b'+' || bytes[exponent] == b'-') {
            exponent += 1;
        }

        if exponent < bytes.len() && bytes[exponent].is_ascii_digit() {
            while exponent < bytes.len() && bytes[exponent].is_ascii_digit() {
                exponent += 1;
            }
            end = exponent;
        }
    }

    let number: f32 = value[..end].trim().parse().ok()?;
    let pixels = number * unit_scale(value[end..].trim())?;

    (pixels.is_finite() && pixels > 0.0).then_some(pixels)
}

/// How many CSS pixels one of the units a document may write its own size in is.
///
/// The absolute units are the ones a measurement can resolve, because they mean the same
/// thing wherever the document is drawn. A percentage is relative to the viewport and
/// `em`/`ex` to a font size, and neither is known here: a document that sizes itself that
/// way is answered by its `viewBox` instead, which is what a viewer falls back on for one
/// too.
fn unit_scale(unit: &str) -> Option<f32> {
    match unit.to_ascii_lowercase().as_str() {
        "" | "px" => Some(1.0),
        "pt" => Some(4.0 / 3.0),
        "pc" => Some(16.0),
        "in" => Some(96.0),
        "mm" => Some(96.0 / 25.4),
        "cm" => Some(96.0 / 2.54),
        "q" => Some(96.0 / 25.4 / 4.0),
        _ => None,
    }
}

/// The size a `viewBox` names, which is the last two of its four numbers: the first two
/// are where the box sits rather than how large it is.
fn view_box_size(value: &str) -> Option<(f32, f32)> {
    let numbers: Vec<f32> = value
        .split(|character: char| character.is_whitespace() || character == ',')
        .filter(|number| !number.is_empty())
        .filter_map(|number| number.parse::<f32>().ok())
        .collect();

    let width = *numbers.get(2)?;
    let height = *numbers.get(3)?;

    (width.is_finite() && width > 0.0 && height.is_finite() && height > 0.0)
        .then_some((width, height))
}

/// A size in whole pixels: rounded to the nearest pixel and at least one, since the
/// layout places a preview from it and a box of no pixels is no preview.
fn whole_pixels(value: f32) -> u32 {
    value.round().max(1.0) as u32
}

/// How many measurements are held at once.
///
/// What a pointer meets again is the folder it is in, so the handful of documents under
/// it are the ones worth keeping; a document that falls out is read and parsed again the
/// next time it is hovered, which is the cost this cache exists to save rather than one
/// it turns into a failure.
const MAX_HELD_MEASUREMENTS: usize = 32;

/// The file's modification time and length: what says a file is not the one that was
/// measured last time.
#[derive(Clone, PartialEq, Eq, Hash)]
struct FileVersion {
    modified: Option<SystemTime>,
    len: u64,
}

/// What a held measurement is only valid for: the file and the version of it that was
/// read.
#[derive(Clone, PartialEq, Eq, Hash)]
struct DocumentKey {
    path: PathBuf,
    version: FileVersion,
}

/// A held measurement: the size the document asks for, or the fact that this version of
/// the file is not a document at all.
///
/// A file that is not one is held as that, rather than as nothing held at all, so a
/// hover onto it costs nothing after the first instead of a read and a failed parse per
/// pass — the same way the PDF page-size cache remembers a file it could not open. What
/// says the answer is still good is the version in the key: a file rewritten is a
/// different key and is read again.
#[derive(Clone, Copy)]
enum Held {
    Measured(u32, u32),
    NotADocument,
}

/// A held measurement and when it was last asked for. The stamp is a counter rather than
/// a clock, so the order documents are dropped in cannot be changed by the system clock
/// moving.
struct HeldMeasurement {
    held: Held,
    last_used: u64,
}

#[derive(Default)]
struct MeasurementCache {
    entries: HashMap<DocumentKey, HeldMeasurement>,
    tick: u64,
}

/// The measurements held between hovers.
///
/// Only measurements: what a document costs this side is two numbers, and what it is
/// drawn as is the engine's business — it is handed the file, not anything this app
/// parsed out of it.
static MEASUREMENTS: Lazy<Mutex<MeasurementCache>> =
    Lazy::new(|| Mutex::new(MeasurementCache::default()));

/// What is held for a version of a file, when anything is.
fn held(key: &DocumentKey) -> Option<Held> {
    let mut cache = MEASUREMENTS.lock().ok()?;
    cache.tick += 1;
    let tick = cache.tick;

    let held = cache.entries.get_mut(key)?;
    held.last_used = tick;

    Some(held.held)
}

/// Hold the size of a file — or the fact that there is none — dropping the least
/// recently used measurement once the cap is passed.
fn hold(key: &DocumentKey, size: Option<(u32, u32)>) {
    let Ok(mut cache) = MEASUREMENTS.lock() else {
        return;
    };

    cache.tick += 1;
    let tick = cache.tick;

    cache.entries.insert(
        key.clone(),
        HeldMeasurement {
            held: match size {
                Some((width, height)) => Held::Measured(width, height),
                None => Held::NotADocument,
            },
            last_used: tick,
        },
    );

    while cache.entries.len() > MAX_HELD_MEASUREMENTS {
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

    /// A document rewritten in place is measured again: the version of the file is part
    /// of what a held measurement is valid for.
    #[test]
    fn a_revised_document_is_measured_again() {
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

    /// A viewBox is four numbers however they are separated, and the size is the last
    /// two of them rather than the first.
    #[test]
    fn reads_a_view_box_written_with_commas() {
        let path = fixture(
            "commas.svg",
            br#"<svg xmlns="http://www.w3.org/2000/svg" viewBox="10,20,64,48"/>"#,
        );

        assert_eq!(measure(&path), Some((64, 48)));
        let _ = std::fs::remove_file(&path);
    }

    /// A size is a length in the units a document may name it in, and those are the
    /// absolute ones — the ones that mean the same thing wherever it is drawn.
    #[test]
    fn reads_a_size_written_in_a_unit() {
        let path = fixture(
            "units.svg",
            br#"<svg xmlns="http://www.w3.org/2000/svg" width="10mm" height="1in"/>"#,
        );

        // Ten millimetres at 96 pixels to the inch is 37.8, and an inch is the 96
        // pixels it is.
        assert_eq!(measure(&path), Some((38, 96)));
        let _ = std::fs::remove_file(&path);
    }

    /// A document that names no size, or one that cannot be resolved without knowing the
    /// viewport it is drawn in, is measured at the viewport every browser assumes.
    #[test]
    fn measures_a_document_with_no_resolvable_size_at_the_default_viewport() {
        for (name, contents) in [
            (
                "unsized.svg",
                br#"<svg xmlns="http://www.w3.org/2000/svg"><rect/></svg>"#.as_slice(),
            ),
            (
                "relative.svg",
                br#"<svg xmlns="http://www.w3.org/2000/svg" width="100%" height="100%"/>"#
                    .as_slice(),
            ),
            (
                "font-relative.svg",
                br#"<svg xmlns="http://www.w3.org/2000/svg" width="4em" height="2ex"/>"#
                    .as_slice(),
            ),
        ] {
            let path = fixture(name, contents);
            assert_eq!(measure(&path), Some((300, 150)), "{name}");
            let _ = std::fs::remove_file(&path);
        }
    }

    /// A size that is not a size — nothing of a width, or less than nothing — is no
    /// answer, so the viewBox answers instead, and the default viewport where there is
    /// none.
    #[test]
    fn refuses_a_size_that_is_not_a_size() {
        let path = fixture(
            "not-a-size.svg",
            br#"<svg xmlns="http://www.w3.org/2000/svg" width="0" height="-4" viewBox="0 0 12 6"/>"#,
        );

        assert_eq!(measure(&path), Some((12, 6)));
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
        for (name, contents) in [
            ("not-a-document.svg", b"just some bytes".as_slice()),
            (
                "not-an-svg-root.svg",
                br#"<html><body>a page</body></html>"#.as_slice(),
            ),
        ] {
            let path = fixture(name, contents);

            assert_eq!(measure(&path), None, "{name}");
            let _ = std::fs::remove_file(&path);
        }
    }
}

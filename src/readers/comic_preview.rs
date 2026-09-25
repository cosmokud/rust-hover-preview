//! The first page of a comic book, read out of the container it is filed in.
//!
//! A comic this app is asked about is not a drawing format at all: a `.cbz` is a zip of plates, a
//! `.cbr` is a rar of them, and a `.cbc` — Calibre's own container — is a zip of several comics'
//! pages under a folder each, with a `comics.txt` naming the comics. What all three have in common
//! is the answer to a hover: what the file holds first, as a page.
//!
//! The page is read here rather than converted by an engine, and that is the whole of why this
//! module exists. The engine that reads comics unpacks the archive, decodes *every* plate, rewrites
//! it — trims borders, adjusts levels, resizes to a device profile and converts it to greyscale
//! unless it is told not to — and builds a document out of the lot, which for the comics this was
//! built against is a hundred megabytes and minutes of work between the hover and the page. What a
//! hover needs is one picture, and one picture is what an archive can hand over directly: the
//! headers are walked for the *names* (which decompresses nothing at all) and then the one entry
//! that is the page is inflated, decoded and drawn. A hundred-megabyte comic costs a page.
//!
//! Which entry is the page is a question about names, and the rules are the ones the comic readers
//! and the engine's own comic input agree on:
//!
//! * **A plate is a picture**, by the name it is written under: `jpg`, `jpeg`, `jpe`, `jfif`, `png`,
//!   `gif`, `bmp`, `tif` and `tiff` — the formats this app decodes a comic's plate with. Anything
//!   else in the container is not a page: a `ComicInfo.xml`, a `comics.txt`, a readme beside the
//!   plates.
//! * **What a Macintosh left behind is not a page.** A zip written on macOS carries a `__MACOSX`
//!   folder of resource forks whose names shadow the real ones, and a plate chosen from there is a
//!   file that is not a picture at all.
//! * **And the first plate is the first by name, counted rather than spelled.** Plates are numbered
//!   and the numbering is padded to different widths by different writers — a `2.jpg` that sorts
//!   after a `10.jpg` is not the second page of anything — so the comparison is by runs of digits:
//!   a run of digits is a number, and everything else is compared as it is written.
//!
//! What is *not* done to it is as deliberate as what is. A plate that is two pages of the comic
//! side by side — a spread, which every scan of a printed book has — is drawn as the one page it is
//! rather than cut in half: a preview is a picture of a page, and half a spread is not a picture of
//! anything. Nothing is written to disk: the plate is inflated into memory under the ceiling one
//! hover may decode for and decoded from there, which is the rule every container reader here
//! follows (see `project_image`). And what is held between hovers is a page's *size*, not the page:
//! a comic is a decode with nothing cached in front of it, so a second hover costs the same decode
//! as the first and nothing more (see `dimensions`).

use crate::config::config::image_decode_limits;
use crate::readers::archive_listing;
use image::GenericImageView;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::SystemTime;

/// The names a comic's plates are written under, which are the picture formats this app decodes a
/// plate with.
///
/// The set is a judgement about comics rather than the whole of what a picture can be: a comic is
/// written in the formats a scanner and a page of artwork are written in, and a `.psd` or an `.exr`
/// is not a plate anybody publishes. A plate of a format this app reads only through the codec
/// Windows has — a `.webp`, an `.avif` — is deliberately absent for a reason that is this
/// module's own: what is decoded here is a *member* rather than a file, and every decoder behind
/// this set reads bytes. Those two would need a stream over memory to be handed to the codec, which
/// is a piece of work of its own and is written down in TODO.md.
const PAGE_NAMES: &[&str] = &[
    "bmp", "gif", "jfif", "jpe", "jpeg", "jpg", "png", "tif", "tiff",
];

/// The folder a Macintosh keeps the resource forks of a zip's members in, whose names shadow the
/// members they belong to.
const MACOSX_FOLDER: &str = "__MACOSX";

/// Entries the size cache holds before it is emptied.
const DIMENSION_CACHE_MAX_ENTRIES: usize = 512;

/// What a size is only valid for: the file, and the version of it the size was read from. A comic
/// saved again is another archive — a volume added to it is the case that matters — and its first
/// page can be another plate.
#[derive(Clone, PartialEq, Eq, Hash, Debug)]
struct DimensionKey {
    path: PathBuf,
    modified: Option<SystemTime>,
    len: u64,
}

/// A plate's size, keyed by the comic and its version. `None` records a container with no plate in
/// it, so a file this reader has no answer for is not opened on every hover.
type DimensionCache = HashMap<DimensionKey, Option<(u32, u32)>>;

static DIMENSIONS: Lazy<Mutex<DimensionCache>> = Lazy::new(|| Mutex::new(HashMap::new()));

/// The size of the first plate of the comic `path` holds, in pixels.
pub fn dimensions(path: &Path) -> Option<(u32, u32)> {
    let key = dimension_key(path);
    if let Ok(cache) = DIMENSIONS.lock() {
        if let Some(cached) = cache.get(&key) {
            return *cached;
        }
    }

    let dimensions = probe_dimensions(path);
    remember_dimensions(key, dimensions);

    dimensions
}

/// The first plate of the comic `path` holds, decoded into `width` by `height` and handed back as
/// BGRA, top down, the way every frame in this app is composed.
pub fn decode(path: &Path, width: u32, height: u32) -> Option<Vec<u8>> {
    if width == 0 || height == 0 {
        return None;
    }

    let name = first_page_name(path)?;
    let bytes = archive_listing::entry_bytes(path, &name)?;

    let mut reader = image::ImageReader::new(std::io::Cursor::new(&bytes))
        .with_guessed_format()
        .ok()?;
    reader.limits(image_decode_limits());
    let plate = reader.decode().ok()?;

    let (source_width, source_height) = plate.dimensions();
    let plate = if source_width != width || source_height != height {
        plate.resize_exact(width, height, image::imageops::FilterType::Triangle)
    } else {
        plate
    };

    Some(crate::ui::preview_window::rgba_to_bgra(
        plate.to_rgba8().as_raw(),
    ))
}

/// The name of the plate a comic is previewed from, out of the archive's own table of contents.
///
/// It is the question the whole reader is built on, and it is public because it is the one a
/// diagnostic asks: a comic whose preview does not appear is a comic whose first plate is either not
/// in the container or not where this reader looked, and which of the two it is is what this answers
/// without decoding anything.
pub fn first_page_name(path: &Path) -> Option<String> {
    let listing = archive_listing::listing_for(path, None)?;

    listing
        .entries
        .iter()
        .filter(|entry| !entry.is_dir && !entry.encrypted && is_page_name(&entry.name))
        .map(|entry| entry.name.as_str())
        .min_by(|left, right| natural_order(left, right))
        .map(str::to_string)
}

/// Whether an entry's name is a plate's: one of the picture formats a plate is written in, and not
/// something a Macintosh left beside the real members.
fn is_page_name(name: &str) -> bool {
    if name.contains(MACOSX_FOLDER) {
        return false;
    }

    let Some((_, extension)) = name.rsplit_once('.') else {
        return false;
    };

    PAGE_NAMES.contains(&extension.to_lowercase().as_str())
}

/// The order two plates are named in: by runs of digits, with everything else compared as it is
/// written.
///
/// It is what makes `2.jpg` come before `10.jpg`, which the plain order of a name does not: a plate
/// is numbered, and a writer pads the numbers to whatever width its own count needs, so the two
/// spellings of the second page are `2` and `02` and neither of them sorts where it belongs. A
/// run of digits is compared as the number it is — and where two runs are the same number, the
/// shorter one comes first, so `2` sits before `02` rather than the other way round.
///
/// Everything else is compared without case, because a plate's own name is written by whoever
/// scanned it rather than by a format: `Chapter` and `chapter` are the same folder of a comic.
fn natural_order(left: &str, right: &str) -> std::cmp::Ordering {
    use std::cmp::Ordering;

    let mut left = left.chars().peekable();
    let mut right = right.chars().peekable();

    loop {
        match (left.next(), right.next()) {
            (None, None) => return Ordering::Equal,
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (Some(left_char), Some(right_char)) => {
                let ordered = if left_char.is_ascii_digit() && right_char.is_ascii_digit() {
                    compare_number(&mut left, &mut right, left_char, right_char)
                } else {
                    let left_char = left_char.to_ascii_lowercase();
                    let right_char = right_char.to_ascii_lowercase();

                    if left_char == right_char {
                        continue;
                    }

                    left_char.cmp(&right_char)
                };

                if ordered != Ordering::Equal {
                    return ordered;
                }
            }
        }
    }
}

/// The order of two runs of digits, each begun by the character already taken from it. Both runs
/// are consumed, so the caller goes on with what follows the number.
fn compare_number(
    left: &mut std::iter::Peekable<std::str::Chars<'_>>,
    right: &mut std::iter::Peekable<std::str::Chars<'_>>,
    left_first: char,
    right_first: char,
) -> std::cmp::Ordering {
    use std::cmp::Ordering;

    let mut left_digits = left_first.to_string();
    let mut right_digits = right_first.to_string();

    while left.peek().is_some_and(char::is_ascii_digit) {
        left_digits.push(left.next().unwrap_or_default());
    }
    while right.peek().is_some_and(char::is_ascii_digit) {
        right_digits.push(right.next().unwrap_or_default());
    }

    // A leading zero is padding rather than a digit of the number: `2` and `02` are the same
    // numeric value, and the shorter spelling comes first.
    let left_number = left_digits.trim_start_matches('0');
    let right_number = right_digits.trim_start_matches('0');

    match left_number.len().cmp(&right_number.len()) {
        Ordering::Equal => match left_number.cmp(right_number) {
            Ordering::Equal => left_digits.len().cmp(&right_digits.len()),
            ordered => ordered,
        },
        ordered => ordered,
    }
}

/// The size the first plate is stored at, read from the plate's own header.
fn probe_dimensions(path: &Path) -> Option<(u32, u32)> {
    let name = first_page_name(path)?;
    let bytes = archive_listing::entry_bytes(path, &name)?;

    picture_dimensions(&bytes)
}

/// A picture's own size, read from its header.
fn picture_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
    image::ImageReader::new(std::io::Cursor::new(bytes))
        .with_guessed_format()
        .ok()?
        .into_dimensions()
        .ok()
}

/// The comic and the version of it a size is read from, read the way every other held value in this
/// app is: what a file is, is its name as it is now, what it weighed, and when it was last written.
fn dimension_key(path: &Path) -> DimensionKey {
    let metadata = std::fs::metadata(path).ok();

    DimensionKey {
        path: path.to_path_buf(),
        modified: metadata
            .as_ref()
            .and_then(|metadata| metadata.modified().ok()),
        len: metadata.map(|metadata| metadata.len()).unwrap_or(0),
    }
}

/// Record a size that has been read, so an archive's table of contents is not walked again for
/// every hover of it.
fn remember_dimensions(key: DimensionKey, dimensions: Option<(u32, u32)>) {
    if let Ok(mut cache) = DIMENSIONS.lock() {
        if !cache.contains_key(&key) && cache.len() >= DIMENSION_CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.insert(key, dimensions);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use image::ImageEncoder;
    use std::io::Write;

    /// A comic holding the plates named, in the order they are named, stored rather than deflated:
    /// what these fixtures are asked is which plate is chosen.
    fn write_comic(path: &Path, members: &[(&str, Vec<u8>)]) {
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        let mut writer = zip::ZipWriter::new(std::fs::File::create(path).expect("a zip to write"));

        for (name, bytes) in members {
            writer.start_file(*name, options).expect("a member");
            writer.write_all(bytes).expect("bytes");
        }

        writer.finish().expect("a finished zip");
    }

    fn plate(width: u32, height: u32) -> Vec<u8> {
        let mut bytes = Vec::new();
        image::codecs::bmp::BmpEncoder::new(&mut bytes)
            .write_image(
                &vec![0u8; (width * height * 3) as usize],
                width,
                height,
                image::ExtendedColorType::Rgb8,
            )
            .expect("a bitmap");

        bytes
    }

    /// The plate a comic is previewed from is the first one by *name* rather than by where the
    /// archive happens to keep it, and the names a comic is filed under are not the names a
    /// container is written in: a `ComicInfo.xml` beside the plates is not a page, and a plate of
    /// the resource fork a Macintosh left in the zip is not one either.
    #[test]
    fn reads_the_first_plate_by_name_rather_than_by_position() {
        let path = std::env::temp_dir().join("rust-hover-preview-comic-order.zip");

        write_comic(
            &path,
            &[
                ("ComicInfo.xml", b"<ComicInfo/>".to_vec()),
                ("__MACOSX/000.jpg", plate(9, 9)),
                ("10.bmp", plate(10, 10)),
                ("2.bmp", plate(2, 2)),
                ("1.bmp", plate(1, 1)),
            ],
        );

        assert_eq!(
            first_page_name(&path).as_deref(),
            Some("1.bmp"),
            "the first plate is the first number, not the first member and not the resource fork"
        );
        assert_eq!(
            dimensions(&path),
            Some((1, 1)),
            "and the size read is that plate's own"
        );

        std::fs::remove_file(&path).ok();
    }

    /// A container with nothing to show a page of is answered with nothing: no preview is the
    /// answer for a box of text, and it is the answer that starts no decoder.
    #[test]
    fn answers_nothing_for_a_box_of_something_else() {
        let path = std::env::temp_dir().join("rust-hover-preview-comic-none.zip");

        write_comic(
            &path,
            &[
                ("comics.txt", b"chapter one\nchapter two\n".to_vec()),
                ("readme.txt", b"a folder of notes\n".to_vec()),
                ("info/", b"".to_vec()),
            ],
        );

        assert_eq!(first_page_name(&path), None);
        assert_eq!(dimensions(&path), None);

        std::fs::remove_file(&path).ok();
    }

    /// A comic is a book rather than an album of unrelated pictures, so the plates are ordered the
    /// way a reader opens them: by the number in the name, whenever the writer's padding would put
    /// them in another order entirely.
    #[test]
    fn counts_the_numbers_in_a_name_rather_than_spelling_them() {
        let mut names = vec!["10.jpg", "2.jpg", "1.jpg", "02.jpg", "100.jpg", "000.jpg"];
        names.sort_by(|left, right| natural_order(left, right));

        assert_eq!(
            names,
            vec!["000.jpg", "1.jpg", "2.jpg", "02.jpg", "10.jpg", "100.jpg"],
            "a plate is the number in its name, and padding is not a bigger number"
        );

        // A comic written in chapters is ordered by the chapter first, and by the number inside it
        // second, which is what comparing the whole name does.
        let mut chapters = vec!["ch10/001.jpg", "ch2/010.jpg", "ch2/2.jpg", "ch1/001.jpg"];
        chapters.sort_by(|left, right| natural_order(left, right));

        assert_eq!(
            chapters,
            vec!["ch1/001.jpg", "ch2/2.jpg", "ch2/010.jpg", "ch10/001.jpg"],
            "a chapter is a number too, and the plates inside it come after their own chapter"
        );

        assert_eq!(
            natural_order("Chapter/1.jpg", "chapter/01.jpg"),
            std::cmp::Ordering::Less,
            "a name is written by whoever scanned it, so it is compared without case"
        );
    }
}

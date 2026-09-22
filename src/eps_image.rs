//! Encapsulated PostScript, read for the picture the file carries of itself.
//!
//! An `.eps` is a PostScript program, and this app has no PostScript interpreter: nothing
//! here parses, evaluates or runs any of it. What is read is the preview a writer leaves
//! inside the file on purpose — for on-screen display, which is the whole reason the
//! format carries one — and there are two shapes it comes in.
//!
//! A Windows file is a container: it opens with a signature no PostScript file does, and
//! behind it a header naming where the program starts and how long it is, and where two
//! previews of it may be — a Windows metafile and a TIFF, either of which may be missing.
//! A Mac file keeps its preview in the comment block the format already has, between
//! `%BeginPreview` and `%EndPreview`, written out as hexadecimal.
//!
//! What either of them holds is a picture of one of two kinds: a metafile, which the
//! drawing layer replays and which is therefore sharp at any size the preview is shown at
//! (`metafile_image`), or a TIFF, decoded by the decoder every other picture in this app
//! is decoded by. Both are preferred in that order where a file carries both, because a
//! metafile is the drawing rather than a picture of it.
//!
//! A file with neither is a program and nothing else, which is what a plotting tool, a
//! typesetting tool, a converter or a newer Illustrator writes: those are answered with
//! no preview, like any other file this app has no reader for.

use crate::config::{decode_budget_bytes, image_decode_limits};
use crate::metafile_image;
use image::GenericImageView;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::fs::{self, File};
use std::io::{Read, Seek, SeekFrom};
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::SystemTime;

/// What a Windows container opens with, which nothing else in the format begins with.
const WINDOWS_SIGNATURE: [u8; 4] = [0xC5, 0xD0, 0xD3, 0xC6];
/// Its header: the signature, three pairs of offsets and lengths — the program, a
/// metafile preview, a TIFF preview — and a checksum.
const WINDOWS_HEADER: usize = 30;
/// How far into a file its comment block is looked through. The block sits at the top of
/// the program, so what is scanned is the head of the file rather than the file.
const COMMENT_PROBE_BYTES: usize = 1 << 20;
const PREVIEW_BEGIN: &[u8] = b"%BeginPreview";
const PREVIEW_END: &[u8] = b"%EndPreview";

/// Entries the size cache holds before it is emptied.
const DIMENSION_CACHE_MAX_ENTRIES: usize = 512;

/// What a size is only valid for: the file, and the version of it the size was read from.
#[derive(Clone, PartialEq, Eq, Hash, Debug)]
struct DimensionKey {
    path: PathBuf,
    modified: Option<SystemTime>,
    len: u64,
}

/// A preview's size, keyed by the file and its version. `None` records a file with no
/// preview in it, so a program that carries none is not looked through again on every
/// hover.
type DimensionCache = HashMap<DimensionKey, Option<(u32, u32)>>;

static DIMENSIONS: Lazy<Mutex<DimensionCache>> = Lazy::new(|| Mutex::new(HashMap::new()));

/// The size of the picture `path` carries a preview of.
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

/// The preview `path` carries, drawn into `width` by `height` and handed back as BGRA,
/// top down, the way every frame in this app is composed.
pub fn decode(path: &Path, width: u32, height: u32) -> Option<Vec<u8>> {
    if width == 0 || height == 0 {
        return None;
    }

    let preview = preview(path)?;

    draw(&preview, width, height)
}

/// The size the preview is of.
fn probe_dimensions(path: &Path) -> Option<(u32, u32)> {
    let preview = preview(path)?;

    match preview {
        Preview::Metafile(records) => metafile_image::records_dimensions(&records),
        Preview::Picture(picture) => picture_dimensions(&picture),
    }
}

/// The preview a file carries, whichever of the two shapes it is in and whichever of the
/// two kinds it holds.
///
/// The file is read for what it declares rather than whole: a Windows container says where
/// its previews are and how long each is, so what is read is the header and then the one
/// preview, and a file that is not a container at all is given up on after its first four
/// bytes — which is what keeps a hover onto a `.wmf` from reading a file this reader has
/// nothing to say about.
fn preview(path: &Path) -> Option<Preview> {
    container_preview(path).or_else(|| comment_preview(path))
}

/// What the Windows container behind the signature names, if it names a preview at all.
///
/// The offsets and lengths are the file's own claim about itself, so each one is measured
/// against the file's length and against the budget before it is used: a header that
/// points past the end of the file, or that claims more bytes than a hover may hold, is
/// answered with no preview rather than with a read of whatever happens to lie there.
fn container_preview(path: &Path) -> Option<Preview> {
    let mut file = File::open(path).ok()?;
    let mut header = [0u8; WINDOWS_HEADER];
    file.read_exact(&mut header).ok()?;

    if header[..4] != WINDOWS_SIGNATURE {
        return None;
    }

    let length = fs::metadata(path).ok()?.len();
    let field = |at: usize| {
        u32::from_le_bytes([header[at], header[at + 1], header[at + 2], header[at + 3]]) as u64
    };
    let mut section = |offset: usize, length_at: usize| -> Option<Vec<u8>> {
        let at = field(offset);
        let size = field(length_at);
        if size == 0 || size > decode_budget_bytes() || at.checked_add(size)? > length {
            return None;
        }

        file.seek(SeekFrom::Start(at)).ok()?;
        let mut bytes = vec![0u8; size as usize];
        file.read_exact(&mut bytes).ok()?;

        Some(bytes)
    };

    if let Some(records) = section(12, 16) {
        return Some(Preview::Metafile(records));
    }

    section(20, 24).map(Preview::Picture)
}

/// The preview the comment block carries, which is the hexadecimal one a Mac writer
/// leaves there.
///
/// The lines after `%BeginPreview` are hex digits behind a comment marker — except the
/// last of them, which is written without one often enough that a reader has to allow for
/// it: what ends the block is the marker that closes it, or a line that is not hex at
/// all, whichever comes first.
fn comment_preview(path: &Path) -> Option<Preview> {
    let mut head = Vec::new();
    File::open(path)
        .ok()?
        .take(COMMENT_PROBE_BYTES as u64)
        .read_to_end(&mut head)
        .ok()?;

    let head = &head[..];
    let start = find(head, PREVIEW_BEGIN)? + PREVIEW_BEGIN.len();

    // The digits behind the marker are the picture's shape — a width, a height, the bits a
    // pixel takes, and the lines of hex that follow — but what is read here is the picture
    // itself rather than the shape written down for it, so the line is stepped over and
    // the hex behind it taken.
    let mut lines = head[start..].split(|byte| *byte == b'\n');
    lines.next()?;

    let mut digits: Vec<u8> = Vec::new();
    for line in lines {
        let line = line.strip_prefix(b"%").unwrap_or(line);
        let line = line.strip_suffix(b"\r").unwrap_or(line);

        if line.starts_with(PREVIEW_END) {
            break;
        }
        if line.is_empty() {
            continue;
        }
        if !line.iter().all(|byte| byte.is_ascii_hexdigit()) {
            break;
        }

        digits.extend_from_slice(line);
    }

    let mut picture = Vec::with_capacity(digits.len() / 2);
    for pair in digits.chunks_exact(2) {
        let (high, low) = (hex(pair[0])?, hex(pair[1])?);
        picture.push(high << 4 | low);
    }

    preview_kind(picture)
}

/// Which of the two kinds those bytes are, by what they open with: a preview is not
/// labelled by the format, so the picture itself is asked.
fn preview_kind(picture: Vec<u8>) -> Option<Preview> {
    if metafile_image::is_metafile_records(&picture) {
        return Some(Preview::Metafile(picture));
    }

    let header = picture.get(..4)?;
    let is_tiff = matches!(
        [header[0], header[1], header[2], header[3]],
        [0x49, 0x49, 0x2A, 0x00] | [0x4D, 0x4D, 0x00, 0x2A]
    );

    is_tiff.then_some(Preview::Picture(picture))
}

/// The preview drawn into the box, whichever kind it is.
fn draw(preview: &Preview, width: u32, height: u32) -> Option<Vec<u8>> {
    match preview {
        Preview::Metafile(records) => metafile_image::draw_records(records, width, height),
        Preview::Picture(picture) => picture_frame(picture, width, height),
    }
}

/// A picture decoded from memory and resampled into the box, which is what every preview
/// made of a picture is: the same decoder, the same budget and the same frame the
/// container previews are made in.
fn picture_frame(bytes: &[u8], width: u32, height: u32) -> Option<Vec<u8>> {
    let mut reader =
        image::ImageReader::new(std::io::Cursor::new(bytes)).with_guessed_format().ok()?;
    reader.limits(image_decode_limits());
    let picture = reader.decode().ok()?;

    let (source_width, source_height) = picture.dimensions();
    let picture = if source_width != width || source_height != height {
        picture.resize_exact(width, height, image::imageops::FilterType::Triangle)
    } else {
        picture
    };

    Some(crate::preview_window::rgba_to_bgra(picture.to_rgba8().as_raw()))
}

/// The size a picture's own header declares.
fn picture_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
    let reader = image::ImageReader::new(std::io::Cursor::new(bytes))
        .with_guessed_format()
        .ok()?;

    reader.into_dimensions().ok()
}

fn hex(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

/// Where a marker appears in those bytes.
fn find(bytes: &[u8], marker: &[u8]) -> Option<usize> {
    bytes
        .windows(marker.len())
        .position(|window| window == marker)
}

/// The two kinds a preview comes in.
enum Preview {
    /// Drawing records, replayed by the drawing layer.
    Metafile(Vec<u8>),
    /// A picture, decoded by the decoder.
    Picture(Vec<u8>),
}

/// The file and the version of it a size is read from, read the way every other held
/// value in this app is: what a file is, is its name as it is now, what it weighed, and
/// when it was last written.
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

/// Record a size that has been read, so a program whose comment block has been looked
/// through once is not looked through again for every hover.
fn remember_dimensions(key: DimensionKey, dimensions: Option<(u32, u32)>) {
    if let Ok(mut cache) = DIMENSIONS.lock() {
        if !cache.contains_key(&key) && cache.len() >= DIMENSION_CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.insert(key, dimensions);
    }
}

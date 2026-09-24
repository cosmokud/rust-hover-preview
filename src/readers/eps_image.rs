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
//! (`metafile_image`), or a picture, decoded by the decoder every other picture in this app
//! is decoded by. Both are preferred in that order where a file carries both, because a
//! metafile is the drawing rather than a picture of it.
//!
//! One shape of picture is read here rather than by that decoder: a colour table with an
//! alpha sample beside its index is what a document written from Photoshop carries, and it
//! is a shape the decoder turns down whole — so the table is looked up and the alpha kept
//! on this side, and every other shape of picture is the decoder's business as it was.
//!
//! A file with neither is a program and nothing else, which is what a plotting tool, a
//! typesetting tool, a converter or a newer Illustrator writes: those are answered with
//! no preview, like any other file this app has no reader for.

use crate::config::config::{decode_budget_bytes, frame_bytes_within_budget, image_decode_limits};
use crate::readers::metafile_image;
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
    for pair in digits.as_chunks::<2>().0 {
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
        // The decoder has the first word on a picture, and the palette reader the second:
        // a colour table with an alpha sample beside its index is a shape the decoder
        // turns down whole, and it is one an EPS carries often enough to be read here.
        Preview::Picture(picture) => picture_frame(picture, width, height)
            .or_else(|| palette_frame(picture, width, height)),
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

    frame_from(picture.to_rgba8(), width, height)
}

/// The picture a palette preview holds, read here because the decoder every other picture
/// goes through will not read this shape at all — an eight-bit index into a colour table,
/// with a sample of alpha beside it, is what a document written from Photoshop carries.
fn palette_frame(bytes: &[u8], width: u32, height: u32) -> Option<Vec<u8>> {
    let palette = Palette::read(bytes)?;

    frame_from(palette.picture(bytes)?, width, height)
}

/// A picture in memory, resampled into the box and handed back as the frame this app
/// composes in — the one ending every picture a preview carries goes through, whether the
/// decoder read it or the palette reader did.
fn frame_from(picture: image::RgbaImage, width: u32, height: u32) -> Option<Vec<u8>> {
    let (source_width, source_height) = picture.dimensions();
    let picture = if source_width != width || source_height != height {
        image::imageops::resize(
            &picture,
            width,
            height,
            image::imageops::FilterType::Triangle,
        )
    } else {
        picture
    };

    Some(crate::ui::preview_window::rgba_to_bgra(picture.as_raw()))
}

/// The size a picture's own header declares.
fn picture_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
    let reader = image::ImageReader::new(std::io::Cursor::new(bytes))
        .with_guessed_format()
        .ok()?;

    // A palette picture with an alpha sample is one the decoder cannot measure either, so
    // what its own directory says is what is answered for it.
    reader
        .into_dimensions()
        .ok()
        .or_else(|| Palette::read(bytes).map(|palette| (palette.width, palette.height)))
}

/// The colour table a palette preview holds, and what the directory beside it says about
/// the picture that indexes it.
///
/// A preview is not written by the reader that will show it: it is a picture a document's
/// writer made to be shown small, and the shapes one comes in are few. What is read here
/// is the one shape the decoder will not take — eight bits an index into a colour table,
/// with eight bits of alpha beside it where the writer kept one, in strips — and everything
/// else about it is refused, so a preview this reader does not understand is a file with
/// no preview rather than a guess at one.
struct Palette {
    width: u32,
    height: u32,
    /// The bytes a pixel takes: the index, and the alpha beside it where there is one.
    samples: usize,
    /// Whether the colour is premultiplied by that alpha, which is what a writer marking
    /// the sample as associated means.
    associated: bool,
    compression: u32,
    strips: Vec<(usize, usize)>,
    /// The table, as the three hundred entries of eight bits a colour is looked up in.
    table: Vec<u8>,
}

/// The two ways a directory writes the numbers in it.
#[derive(Clone, Copy)]
enum Order {
    Little,
    Big,
}

impl Order {
    fn u16(self, bytes: &[u8], at: usize) -> Option<u16> {
        let pair = [*bytes.get(at)?, *bytes.get(at + 1)?];

        Some(match self {
            Self::Little => u16::from_le_bytes(pair),
            Self::Big => u16::from_be_bytes(pair),
        })
    }

    fn u32(self, bytes: &[u8], at: usize) -> Option<u32> {
        let mut quad = [0u8; 4];
        quad.copy_from_slice(bytes.get(at..at + 4)?);

        Some(match self {
            Self::Little => u32::from_le_bytes(quad),
            Self::Big => u32::from_be_bytes(quad),
        })
    }
}

/// One directory entry's values, as the numbers they are: what fits inside the entry sits
/// in it, and what does not is read from where the entry says it is.
fn entry_values(order: Order, bytes: &[u8], entry: usize, count: usize, size: usize) -> Option<Vec<u64>> {
    let length = count.checked_mul(size)?;
    let raw = if length > 4 {
        let at = order.u32(bytes, entry + 8)? as usize;
        bytes.get(at..at.checked_add(length)?)?
    } else {
        bytes.get(entry + 8..entry + 12)?
    };

    match size {
        2 => (0..count)
            .map(|index| order.u16(raw, index * 2).map(u64::from))
            .collect(),
        4 => (0..count)
            .map(|index| order.u32(raw, index * 4).map(u64::from))
            .collect(),
        _ => None,
    }
}

impl Palette {
    /// Read the directory, or refuse: every field this reader depends on has to be there
    /// and has to say what a preview of this shape says.
    fn read(bytes: &[u8]) -> Option<Self> {
        let order = match bytes.get(..2)? {
            b"II" => Order::Little,
            b"MM" => Order::Big,
            _ => return None,
        };
        if order.u16(bytes, 2)? != 42 {
            return None;
        }

        let directory = order.u32(bytes, 4)? as usize;
        let count = order.u16(bytes, directory)? as usize;

        let mut width = None;
        let mut height = None;
        let mut bits: Vec<u64> = Vec::new();
        let mut compression = None;
        let mut photometric = None;
        let mut offsets: Vec<u64> = Vec::new();
        let mut lengths: Vec<u64> = Vec::new();
        let mut samples = None;
        let mut planar = None;
        let mut predictor = None;
        let mut extra: Vec<u64> = Vec::new();
        let mut table: Vec<u64> = Vec::new();

        for index in 0..count {
            let entry = directory.checked_add(2 + index * 12)?;
            let tag = order.u16(bytes, entry)?;
            let kind = order.u16(bytes, entry + 2)?;
            let number = order.u32(bytes, entry + 4)? as usize;
            let size = match kind {
                1 => 1,
                3 => 2,
                4 => 4,
                _ => continue,
            };
            let values = || entry_values(order, bytes, entry, number, size);

            match tag {
                256 => width = values()?.first().copied().map(|value| value as u32),
                257 => height = values()?.first().copied().map(|value| value as u32),
                258 => bits = values()?,
                259 => compression = values()?.first().copied().map(|value| value as u32),
                262 => photometric = values()?.first().copied().map(|value| value as u32),
                273 => offsets = values()?,
                277 => samples = values()?.first().copied().map(|value| value as usize),
                279 => lengths = values()?,
                284 => planar = values()?.first().copied().map(|value| value as u32),
                317 => predictor = values()?.first().copied().map(|value| value as u32),
                320 => table = values()?,
                338 => extra = values()?,
                _ => {}
            }
        }

        let (width, height) = (width?, height?);
        let samples = samples?;
        if width == 0 || height == 0 || !(1..=2).contains(&samples) {
            return None;
        }
        // A colour table is what makes this shape this shape, and an eight-bit index is
        // what the table is looked up by.
        if photometric != Some(3) {
            return None;
        }
        // A directory writes one depth for every sample where they are all the same depth,
        // which is what a preview of this shape does, and a list has to describe every
        // sample it names.
        if bits.is_empty()
            || (bits.len() > 1 && bits.len() < samples)
            || bits.iter().any(|bits| *bits != 8)
        {
            return None;
        }
        if planar.unwrap_or(1) != 1 || predictor.unwrap_or(1) != 1 {
            return None;
        }
        let compression = compression.unwrap_or(1);
        if compression != 1 && compression != 32773 {
            return None;
        }
        // The second sample is the alpha beside the index, and a sample the directory does
        // not name as one is not read: what it holds has to be said rather than assumed.
        if samples == 2 && extra.is_empty() {
            return None;
        }
        if table.len() != 768 {
            return None;
        }

        let strips: Vec<(usize, usize)> = offsets
            .iter()
            .zip(&lengths)
            .map(|(at, length)| Some((*at as usize, *length as usize)))
            .collect::<Option<Vec<_>>>()?;
        if strips.is_empty() {
            return None;
        }
        for (at, length) in &strips {
            if at.checked_add(*length)? > bytes.len() {
                return None;
            }
        }

        Some(Self {
            width,
            height,
            samples,
            associated: matches!(extra.first(), Some(1)),
            compression,
            strips,
            // A table is written in sixteen bits a colour, and read here as the eight a
            // frame is composed in.
            table: table.iter().map(|value| (value >> 8) as u8).collect(),
        })
    }

    /// The picture the table indexes, assembled from the strips the directory names.
    fn picture(&self, bytes: &[u8]) -> Option<image::RgbaImage> {
        // The canvas this builds is asked of the budget the rest of the app reads under
        // before a byte of it is allocated, the way every other reader asks it.
        frame_bytes_within_budget(self.width, self.height, 4)?;

        let pixels = (self.width as usize).checked_mul(self.height as usize)?;
        let expected = pixels.checked_mul(self.samples)?;

        let mut data = Vec::with_capacity(expected);
        for (at, length) in &self.strips {
            let strip = bytes.get(*at..*at + *length)?;
            match self.compression {
                1 => data.extend_from_slice(strip),
                32773 => unpack_packbits(strip, &mut data)?,
                _ => return None,
            }
        }
        if data.len() < expected {
            return None;
        }

        let mut picture = image::RgbaImage::new(self.width, self.height);
        let frame = picture.as_mut();

        for index in 0..pixels {
            let at = index * self.samples;
            // A table is written a colour at a time rather than a pixel at a time: every
            // red in the table, then every green, then every blue — so the three that make
            // one colour are a table's length apart from each other.
            let entry = data[at] as usize;
            let (mut red, mut green, mut blue) = (
                self.table[entry],
                self.table[256 + entry],
                self.table[512 + entry],
            );
            let alpha = if self.samples > 1 { data[at + 1] } else { 255 };

            // An associated alpha is a colour already multiplied by it, and what a frame is
            // composited from is the colour itself — while a pixel that is not there at all
            // has no colour of its own to give back.
            if self.associated && alpha != 255 {
                if alpha == 0 {
                    red = 0;
                    green = 0;
                    blue = 0;
                } else {
                    red = back_out(red, alpha);
                    green = back_out(green, alpha);
                    blue = back_out(blue, alpha);
                }
            }

            frame[index * 4] = red;
            frame[index * 4 + 1] = green;
            frame[index * 4 + 2] = blue;
            frame[index * 4 + 3] = alpha;
        }

        Some(picture)
    }
}

/// A colour taken back out of the alpha it was multiplied by.
fn back_out(colour: u8, alpha: u8) -> u8 {
    (colour as u32 * 255 / alpha as u32).min(255) as u8
}

/// A strip of PackBits, which is the one compression a preview of this shape is written
/// with besides none at all: a count byte saying what the bytes behind it are, as it is in
/// every PackBits row, over the whole strip rather than one row at a time.
fn unpack_packbits(packed: &[u8], out: &mut Vec<u8>) -> Option<()> {
    let mut at = 0;
    while at < packed.len() {
        let count = packed[at] as i8;
        at += 1;

        if count >= 0 {
            let run = count as usize + 1;
            out.extend_from_slice(packed.get(at..at.checked_add(run)?)?);
            at += run;
        } else if count != -128 {
            let value = *packed.get(at)?;
            at += 1;
            out.resize(out.len() + (1 - count as i32) as usize, value);
        }
    }

    Some(())
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

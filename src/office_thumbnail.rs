//! The picture an Office document saved inside itself.
//!
//! Every Word, Excel and PowerPoint document can carry a preview of its own first
//! page, and Office writes one into the file when the document is saved with
//! thumbnails on. That picture is what Explorer draws as the file's icon, and it
//! is what a hover shows: reading it costs a few milliseconds and involves no
//! Office process at all.
//!
//! Where it lives depends on the container. An OOXML package keeps it in a part
//! (`docProps/thumbnail.emf`, `thumbnail.wmf` or `thumbnail.jpeg`); a legacy
//! document keeps it in its summary information stream, as a clipboard-format
//! picture. Both are read here, and both are read by content rather than by name,
//! because the part's extension and the clip format are the producer's claim
//! about bytes that carry their own signature.

use crate::office_formats::{container_kind, OfficeContainer};
use once_cell::sync::Lazy;
use std::fs::File;
use std::io::{Cursor, Read};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::SystemTime;

/// The most a thumbnail may be. A real one is a few kilobytes — a metafile of the
/// page — up to a few hundred for a saved JPEG; the cap is what keeps a
/// mislabeled archive from being read into memory as if it were one.
const MAX_THUMBNAIL_BYTES: u64 = 8 * 1024 * 1024;
/// Entries a package is scanned over. A document has a handful to a few hundred;
/// a zip that is not a document stops being interesting well before this.
const MAX_PACKAGE_ENTRIES: usize = 2048;
/// Properties a summary information section may hold. The real one has around
/// twenty.
const MAX_PROPERTY_SET_PROPERTIES: u32 = 256;
const CACHE_MAX_ENTRIES: usize = 256;
/// The longest side a picture's extent is believed to describe, in pixels. Past
/// this the extent is not a page size but a corrupt field.
const MAX_THUMBNAIL_SIDE: f64 = 20000.0;

/// The stream a legacy document keeps its summary information in — the name
/// begins with the byte `0x05`, not with a letter.
const SUMMARY_INFORMATION_STREAM: &str = "\u{5}SummaryInformation";
/// The property the picture is held under, and the type it has.
///
/// The `PIDSI_*` constants are 1-based, so `PIDSI_THUMBNAIL` is 17 while the ID
/// actually written in the stream is 16 — the two are not the same number and
/// only one of them is in the file.
const PIDSI_THUMBNAIL_ID: u32 = 0x10;
const VT_CF: u32 = 71;

/// Clipboard formats a property-set picture can be in.
const CF_METAFILEPICT: u32 = 3;
const CF_DIB: u32 = 8;
const CF_ENHMETAFILE: u32 = 14;
const CF_DIBV5: u32 = 17;

/// Map modes a `METAFILEPICT` extent counts its units in.
const MM_TEXT: u32 = 1;
const MM_LOMETRIC: u32 = 2;
const MM_HIMETRIC: u32 = 3;
const MM_LOENGLISH: u32 = 4;
const MM_HIENGLISH: u32 = 5;
const MM_TWIPS: u32 = 6;
const MM_ISOTROPIC: u32 = 7;
const MM_ANISOTROPIC: u32 = 8;

/// What the picture's bytes are, which is also how they are drawn.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum ThumbnailKind {
    /// An enhanced metafile, with its frame in hundredths of a millimetre.
    Emf {
        frame_width_01mm: i64,
        frame_height_01mm: i64,
    },
    /// A Windows metafile, with the extent its header or its `METAFILEPICT` gives.
    Wmf { extent: WmfExtent },
    /// A device-independent bitmap: a `BITMAPINFO` followed by its bits.
    Dib,
    /// An image the `image` crate decodes.
    Raster,
}

/// The units a Windows metafile's extent is counted in.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum WmfExtent {
    /// Hundredths of a millimetre.
    Himetric(i64, i64),
    /// Twips, 1440 to the inch.
    Twips(i64, i64),
    /// Device pixels.
    Pixels(i64, i64),
}

impl WmfExtent {
    /// The extent in pixels at 96 DPI.
    fn pixels(self) -> Option<(u32, u32)> {
        let (width, height, units_per_inch) = match self {
            Self::Himetric(width, height) => (width as f64, height as f64, 2540.0),
            Self::Twips(width, height) => (width as f64, height as f64, 1440.0),
            Self::Pixels(width, height) => return pixels_within_cap(width as f64, height as f64),
        };

        pixels_within_cap(
            width * 96.0 / units_per_inch,
            height * 96.0 / units_per_inch,
        )
    }
}

pub(crate) struct Thumbnail {
    pub(crate) bytes: Vec<u8>,
    pub(crate) kind: ThumbnailKind,
    /// The picture's own size in pixels at 96 DPI.
    pub(crate) width: u32,
    pub(crate) height: u32,
}

impl Thumbnail {
    /// Whether the picture is a metafile, which is vector: the preview draws it
    /// at whatever size it needs rather than enlarging a raster.
    pub(crate) fn is_metafile(&self) -> bool {
        matches!(
            self.kind,
            ThumbnailKind::Emf { .. } | ThumbnailKind::Wmf { .. }
        )
    }
}

/// What a held thumbnail is valid for: the file, and the version of it that was
/// read.
type ThumbnailKey = (PathBuf, Option<SystemTime>, u64);

struct CacheEntry {
    key: ThumbnailKey,
    value: Option<Arc<Thumbnail>>,
}

/// Thumbnails that have been read, and the answer "this document has none" along
/// with them — a file list is a place a pointer is swept back and forth over, and
/// a document without a saved picture must not be unzipped again on every pass.
static THUMBNAILS: Lazy<Mutex<Vec<CacheEntry>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// The picture `path` carries, from the cache when it has been read before.
///
/// `None` is both "not read yet and there is nothing there" and "read, and this
/// document has no saved thumbnail"; either way the caller shows the document
/// without one. A failure is remembered for the file's version, the way a PDF's
/// page size is, so a document without a picture costs one look rather than one
/// per hover.
pub(crate) fn thumbnail_for(path: &Path, cancel: Option<&AtomicBool>) -> Option<Arc<Thumbnail>> {
    let key = thumbnail_key(path);

    if let Ok(cache) = THUMBNAILS.lock() {
        if let Some(entry) = cache.iter().find(|entry| entry.key == key) {
            return entry.value.clone();
        }
    }

    let value = read_thumbnail(path, cancel).map(Arc::new);

    if let Ok(mut cache) = THUMBNAILS.lock() {
        if cache.len() >= CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.push(CacheEntry {
            key,
            value: value.clone(),
        });
    }

    value
}

fn thumbnail_key(path: &Path) -> ThumbnailKey {
    let metadata = std::fs::metadata(path).ok();

    (
        path.to_path_buf(),
        metadata
            .as_ref()
            .and_then(|metadata| metadata.modified().ok()),
        metadata.map(|metadata| metadata.len()).unwrap_or(0),
    )
}

fn read_thumbnail(path: &Path, cancel: Option<&AtomicBool>) -> Option<Thumbnail> {
    match container_kind(path)? {
        OfficeContainer::Ooxml => read_package_thumbnail(path, cancel),
        OfficeContainer::Ole => read_ole_thumbnail(path, cancel),
    }
}

/// The picture an OOXML package saved.
///
/// The part is `docProps/thumbnail.*`, with the OpenDocument-style producers
/// writing `Thumbnails/thumbnail.*`; both are matched case-insensitively, because
/// a package part name is not guaranteed to keep the spelling its producer used.
fn read_package_thumbnail(path: &Path, cancel: Option<&AtomicBool>) -> Option<Thumbnail> {
    let file = File::open(path).ok()?;
    let mut archive = zip::ZipArchive::new(file).ok()?;

    let mut best: Option<(u8, usize)> = None;
    for index in 0..archive.len().min(MAX_PACKAGE_ENTRIES) {
        if cancelled(cancel) {
            return None;
        }

        let Ok(entry) = archive.by_index(index) else {
            continue;
        };
        let Some(priority) = thumbnail_part_priority(entry.name()) else {
            continue;
        };

        if best
            .map(|(best_priority, _)| priority < best_priority)
            .unwrap_or(true)
        {
            best = Some((priority, index));
        }
    }

    let (_, index) = best?;
    let mut entry = archive.by_index(index).ok()?;
    if entry.size() > MAX_THUMBNAIL_BYTES {
        return None;
    }

    let mut bytes = Vec::with_capacity(entry.size() as usize);
    let mut limited = entry.by_ref().take(MAX_THUMBNAIL_BYTES + 1);
    limited.read_to_end(&mut bytes).ok()?;
    if bytes.len() as u64 > MAX_THUMBNAIL_BYTES {
        return None;
    }

    classify(bytes)
}

/// How good a package part is as the thumbnail, lower being better; `None` is
/// not a thumbnail part at all.
///
/// The metafiles come first because they are what Office itself writes and they
/// are vector, so a preview drawn from one is sharp at any size; a raster part is
/// what older Office and other producers write, and it is only as good as the
/// pixels it saved.
fn thumbnail_part_priority(name: &str) -> Option<u8> {
    let name = name.to_ascii_lowercase();
    let is_thumbnail =
        name.starts_with("docprops/thumbnail.") || name.starts_with("thumbnails/thumbnail.");
    if !is_thumbnail {
        return None;
    }

    Some(match name.rsplit('.').next()? {
        "emf" => 0,
        "wmf" => 1,
        "png" => 2,
        "jpg" | "jpeg" => 3,
        "bmp" | "gif" | "tif" | "tiff" => 4,
        _ => 5,
    })
}

/// The picture a legacy document saved, out of the summary information stream of
/// its compound file.
///
/// Only that one stream is opened, and the compound file itself is walked no
/// further than finding it — the rest of a 97-2003 document is a format this app
/// has no other business in.
fn read_ole_thumbnail(path: &Path, cancel: Option<&AtomicBool>) -> Option<Thumbnail> {
    if cancelled(cancel) {
        return None;
    }

    let file = File::open(path).ok()?;
    let mut compound = cfb::CompoundFile::open(file).ok()?;
    let stream = compound.open_stream(SUMMARY_INFORMATION_STREAM).ok()?;

    let mut bytes = Vec::new();
    let mut limited = stream.take(MAX_THUMBNAIL_BYTES + 1);
    limited.read_to_end(&mut bytes).ok()?;
    if bytes.len() as u64 > MAX_THUMBNAIL_BYTES {
        return None;
    }

    let (clip_format, picture) = property_set_thumbnail(&bytes)?;
    picture_from_property(clip_format, picture)
}

/// The picture out of a summary information property set: the clipboard format it
/// is stored as, and the bytes that carry it.
///
/// The layout is the one every OLE property set has — a header naming the sets it
/// holds, then each set's own table of properties — read no further than the one
/// property wanted.
fn property_set_thumbnail(bytes: &[u8]) -> Option<(u32, &[u8])> {
    // Byte order, version, system identifier, CLSID, then one entry per set: its
    // format ID and the offset of its section.
    if read_u16(bytes, 0)? != 0xFFFE {
        return None;
    }
    let sets = read_u32(bytes, 24)? as usize;
    if sets == 0 || sets > 8 {
        return None;
    }

    for index in 0..sets {
        let entry = 28 + index * 20;
        let section = read_u32(bytes, entry + 16)? as usize;
        if let Some(picture) = section_thumbnail(bytes, section) {
            return Some(picture);
        }
    }

    None
}

fn section_thumbnail(bytes: &[u8], section: usize) -> Option<(u32, &[u8])> {
    let count = read_u32(bytes, section + 4)?;
    if count > MAX_PROPERTY_SET_PROPERTIES {
        return None;
    }

    for index in 0..count as usize {
        let entry = section + 8 + index * 8;
        if read_u32(bytes, entry)? != PIDSI_THUMBNAIL_ID {
            continue;
        }

        // A property's value: its type, then the clipboard picture — the size of
        // the data that follows it, the format, and the data itself.
        let value = section.checked_add(read_u32(bytes, entry + 4)? as usize)?;
        if read_u32(bytes, value)? != VT_CF {
            return None;
        }

        let size = read_u32(bytes, value + 4)? as usize;
        let clip_format = read_u32(bytes, value + 8)?;
        let data = bytes.get(value + 12..value + 12 + size.checked_sub(4)?)?;

        return Some((clip_format, data));
    }

    None
}

/// What the clip format says the bytes are.
fn picture_from_property(clip_format: u32, data: &[u8]) -> Option<Thumbnail> {
    match clip_format {
        CF_METAFILEPICT => {
            // A METAFILEPICT describing the extent, then the metafile's own bits:
            // the extent it states is what the picture is measured and drawn at,
            // and where the bits start is read off the metafile itself.
            let map_mode = read_u32(data, 0)?;
            let width = read_u32(data, 4)? as i64;
            let height = read_u32(data, 8)? as i64;
            let extent = wmf_extent(map_mode, width, height)?;
            let (width, height) = extent.pixels()?;
            let (start, _) = wmf_bits(data)?;

            Some(Thumbnail {
                bytes: data.get(start..)?.to_vec(),
                kind: ThumbnailKind::Wmf { extent },
                width,
                height,
            })
        }
        CF_ENHMETAFILE => classify_emf(data.to_vec()),
        CF_DIB | CF_DIBV5 => classify_dib(data.to_vec()),
        // CF_BITMAP is a GDI handle, which a file cannot carry, and anything else
        // is a format this app does not draw.
        _ => None,
    }
}

fn wmf_extent(map_mode: u32, width: i64, height: i64) -> Option<WmfExtent> {
    match map_mode {
        MM_TEXT => Some(WmfExtent::Pixels(width, height)),
        MM_HIMETRIC => Some(WmfExtent::Himetric(width, height)),
        MM_TWIPS => Some(WmfExtent::Twips(width, height)),
        // The English and metric map modes count in their own units, which are
        // converted to hundredths of a millimetre on the way in so that one
        // converter serves them all.
        MM_LOMETRIC => Some(WmfExtent::Himetric(width * 10, height * 10)),
        MM_LOENGLISH => Some(WmfExtent::Himetric(width * 254 / 10, height * 254 / 10)),
        MM_HIENGLISH => Some(WmfExtent::Himetric(width * 254 / 100, height * 254 / 100)),
        // The unbounded map modes are what Office writes for a thumbnail, and
        // what it means by them is the metric it records the page in.
        MM_ISOTROPIC | MM_ANISOTROPIC => Some(WmfExtent::Himetric(width, height)),
        _ => None,
    }
}

/// What the bytes are, by their own signature: a metafile says so in its header,
/// and anything else is handed to the image decoder.
fn classify(bytes: Vec<u8>) -> Option<Thumbnail> {
    if emf_frame(&bytes).is_some() {
        return classify_emf(bytes);
    }

    if let Some((start, Some(bounds))) = wmf_bits(&bytes) {
        let (width, height) = WmfExtent::Twips(bounds.0, bounds.1).pixels()?;
        return Some(Thumbnail {
            bytes: bytes.get(start..)?.to_vec(),
            kind: ThumbnailKind::Wmf {
                extent: WmfExtent::Twips(bounds.0, bounds.1),
            },
            width,
            height,
        });
    }

    classify_raster(bytes)
}

fn classify_emf(bytes: Vec<u8>) -> Option<Thumbnail> {
    let (frame_width, frame_height) = emf_frame(&bytes)?;
    let (width, height) = WmfExtent::Himetric(frame_width, frame_height).pixels()?;

    Some(Thumbnail {
        bytes,
        kind: ThumbnailKind::Emf {
            frame_width_01mm: frame_width,
            frame_height_01mm: frame_height,
        },
        width,
        height,
    })
}

fn classify_dib(bytes: Vec<u8>) -> Option<Thumbnail> {
    let (width, height) = dib_dimensions(&bytes)?;

    Some(Thumbnail {
        bytes,
        kind: ThumbnailKind::Dib,
        width,
        height,
    })
}

fn classify_raster(bytes: Vec<u8>) -> Option<Thumbnail> {
    let (width, height) = image::ImageReader::new(Cursor::new(&bytes))
        .with_guessed_format()
        .ok()?
        .into_dimensions()
        .ok()?;
    if width == 0 || height == 0 {
        return None;
    }

    Some(Thumbnail {
        bytes,
        kind: ThumbnailKind::Raster,
        width,
        height,
    })
}

/// An enhanced metafile's frame, in hundredths of a millimetre, from the
/// `ENHMETAHEADER` that begins it.
fn emf_frame(bytes: &[u8]) -> Option<(i64, i64)> {
    // Type 1 is a metafile, and the signature sits 40 bytes in, past the bounds
    // and the frame that came before it.
    if read_u32(bytes, 0)? != 1 || read_u32(bytes, 40)? != 0x464D_4520 {
        return None;
    }

    let width = read_i32(bytes, 32)? as i64 - read_i32(bytes, 24)? as i64;
    let height = read_i32(bytes, 36)? as i64 - read_i32(bytes, 28)? as i64;
    if width <= 0 || height <= 0 {
        return None;
    }

    Some((width, height))
}

/// Where a Windows metafile's own bits begin, and the bounding box a placeable
/// header stated for them.
///
/// A placeable metafile carries an Aldus header that is not part of what GDI
/// draws — the box it states is kept and the header is held back — and a picture
/// out of a property set may carry a `METAFILEPICT` in front of the bits as well.
/// Which of them is there, and how long it is, is not written the same way by
/// every producer, so the header is looked for rather than assumed: a metafile
/// signature within the first few bytes says where the bits start.
fn wmf_bits(data: &[u8]) -> Option<(usize, Option<(i64, i64)>)> {
    for offset in 0..data.len().min(24) {
        if read_u32(data, offset) == Some(0x9AC6_CDD7) {
            let width = read_i16(data, offset + 10)? as i64 - read_i16(data, offset + 6)? as i64;
            let height = read_i16(data, offset + 12)? as i64 - read_i16(data, offset + 8)? as i64;
            let bounds = (width > 0 && height > 0).then_some((width, height));

            // Past the header: eleven words of it.
            return Some((offset + 22, bounds));
        }

        // A metafile without a placeable header: the type, the header length it
        // always has, and the version that says it is one.
        let is_header = read_u16(data, offset)
            .map(|kind| kind == 1 || kind == 2)
            .unwrap_or(false)
            && read_u16(data, offset + 2) == Some(9)
            && read_u16(data, offset + 4) == Some(0x0300);
        if is_header {
            return Some((offset, None));
        }
    }

    None
}

/// A `BITMAPINFOHEADER`'s dimensions, for the device-independent bitmaps a
/// property set carries. The height's sign says which way up the rows are, so
/// what is wanted here is its magnitude.
fn dib_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
    let header_size = read_u32(bytes, 0)?;
    if !(40..=124).contains(&header_size) {
        return None;
    }

    let width = read_i32(bytes, 4)?;
    let height = read_i32(bytes, 8)?.unsigned_abs();
    if width <= 0 || height == 0 {
        return None;
    }

    Some((width as u32, height))
}

fn pixels_within_cap(width: f64, height: f64) -> Option<(u32, u32)> {
    if !width.is_finite() || !height.is_finite() || width <= 0.0 || height <= 0.0 {
        return None;
    }

    let longest = width.max(height);
    let scale = if longest > MAX_THUMBNAIL_SIDE {
        MAX_THUMBNAIL_SIDE / longest
    } else {
        1.0
    };

    Some((
        (width * scale).round().max(1.0) as u32,
        (height * scale).round().max(1.0) as u32,
    ))
}

fn cancelled(cancel: Option<&AtomicBool>) -> bool {
    cancel
        .map(|cancel| cancel.load(Ordering::Acquire))
        .unwrap_or(false)
}

fn read_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_le_bytes(
        bytes.get(offset..offset + 2)?.try_into().ok()?,
    ))
}

fn read_i16(bytes: &[u8], offset: usize) -> Option<i16> {
    Some(i16::from_le_bytes(
        bytes.get(offset..offset + 2)?.try_into().ok()?,
    ))
}

fn read_u32(bytes: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(
        bytes.get(offset..offset + 4)?.try_into().ok()?,
    ))
}

fn read_i32(bytes: &[u8], offset: usize) -> Option<i32> {
    Some(i32::from_le_bytes(
        bytes.get(offset..offset + 4)?.try_into().ok()?,
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn test_folder() -> PathBuf {
        // A folder of this module's own: the tests run beside each other, and one
        // of them clearing its fixtures must not take another's with it.
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-office-tests")
            .join("thumbnail");
        std::fs::create_dir_all(&folder).expect("a test folder");
        folder
    }

    /// A PNG of the given size, encoded the way any producer would.
    fn png_bytes(width: u32, height: u32) -> Vec<u8> {
        let image = image::RgbaImage::from_pixel(width, height, image::Rgba([200, 100, 50, 255]));
        let mut bytes = Vec::new();
        image
            .write_to(&mut Cursor::new(&mut bytes), image::ImageFormat::Png)
            .expect("an encoded PNG");
        bytes
    }

    #[test]
    fn reads_the_thumbnail_part_out_of_a_package() {
        let path = test_folder().join("package-reader.docx");
        let file = File::create(&path).expect("a test package");
        let mut writer = zip::ZipWriter::new(file);
        let stored = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);

        writer
            .start_file("word/document.xml", stored)
            .expect("a written entry");
        writer.write_all(b"<w:document/>").expect("a written part");
        // Written under the spelling Office does not use, because the part name
        // is matched case-insensitively.
        writer
            .start_file("docProps/Thumbnail.PNG", stored)
            .expect("a written entry");
        writer
            .write_all(&png_bytes(4, 3))
            .expect("a written picture");
        writer.finish().expect("a finished package");

        let thumbnail = thumbnail_for(&path, None).expect("a thumbnail");
        assert_eq!(thumbnail.kind, ThumbnailKind::Raster);
        assert_eq!((thumbnail.width, thumbnail.height), (4, 3));
        assert!(!thumbnail.is_metafile());

        // A second look is served from the cache, which is why the answer is the
        // same object rather than a second reading.
        let again = thumbnail_for(&path, None).expect("a cached thumbnail");
        assert!(Arc::ptr_eq(&thumbnail, &again));

        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn remembers_a_package_without_a_thumbnail() {
        let path = test_folder().join("package-empty.xlsx");
        let file = File::create(&path).expect("a test package");
        let mut writer = zip::ZipWriter::new(file);
        let stored = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        writer
            .start_file("xl/workbook.xml", stored)
            .expect("a written entry");
        writer.write_all(b"<workbook/>").expect("a written part");
        writer.finish().expect("a finished package");

        assert!(thumbnail_for(&path, None).is_none());
        assert!(thumbnail_for(&path, None).is_none());

        let _ = std::fs::remove_file(&path);
    }

    /// A summary information stream holding one clipboard-format property.
    fn summary_information(picture: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&0xFFFEu16.to_le_bytes()); // byte order
        bytes.extend_from_slice(&0u16.to_le_bytes()); // version
        bytes.extend_from_slice(&0u32.to_le_bytes()); // system identifier
        bytes.extend_from_slice(&[0u8; 16]); // class ID
        bytes.extend_from_slice(&1u32.to_le_bytes()); // one property set
        bytes.extend_from_slice(&[0u8; 16]); // format ID
        bytes.extend_from_slice(&48u32.to_le_bytes()); // section offset

        // The section: its size, one property, and where that property sits.
        bytes.extend_from_slice(&0u32.to_le_bytes()); // section size
        bytes.extend_from_slice(&1u32.to_le_bytes()); // one property
        bytes.extend_from_slice(&PIDSI_THUMBNAIL_ID.to_le_bytes());
        bytes.extend_from_slice(&16u32.to_le_bytes()); // property offset

        // The property: a clipboard picture — the size of the data after it, the
        // format, and the data.
        bytes.extend_from_slice(&VT_CF.to_le_bytes());
        bytes.extend_from_slice(&((picture.len() as u32 + 4).to_le_bytes()));
        bytes.extend_from_slice(&CF_METAFILEPICT.to_le_bytes());
        bytes.extend_from_slice(picture);

        bytes
    }

    /// The bytes a metafile starts with: the type, the header length it always
    /// has, and the version that says it is one.
    const WMF_HEADER: [u8; 18] = [
        0x01, 0x00, 0x09, 0x00, 0x00, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00,
    ];

    #[test]
    fn reads_the_thumbnail_property_of_a_summary_information_stream() {
        // A METAFILEPICT of 2 by 1 centimetres, then the metafile itself.
        let mut picture = Vec::new();
        picture.extend_from_slice(&MM_ANISOTROPIC.to_le_bytes());
        picture.extend_from_slice(&2000u32.to_le_bytes()); // 2 cm
        picture.extend_from_slice(&1000u32.to_le_bytes()); // 1 cm
        picture.extend_from_slice(&WMF_HEADER);
        picture.extend_from_slice(b"wmf bits");

        let stream = summary_information(&picture);
        let (clip_format, data) = property_set_thumbnail(&stream).expect("the property");
        assert_eq!(clip_format, CF_METAFILEPICT);
        assert_eq!(&data[12..30], &WMF_HEADER);

        let thumbnail = picture_from_property(clip_format, data).expect("a thumbnail");
        assert_eq!(
            thumbnail.kind,
            ThumbnailKind::Wmf {
                extent: WmfExtent::Himetric(2000, 1000)
            }
        );
        // 2 cm and 1 cm at 96 DPI.
        assert_eq!((thumbnail.width, thumbnail.height), (76, 38));
        // The METAFILEPICT is not part of what GDI draws: the picture keeps the
        // metafile's own bits and nothing before them.
        assert_eq!(&thumbnail.bytes[..18], &WMF_HEADER);
    }

    #[test]
    fn the_property_id_is_the_one_in_the_file() {
        // The 1-based `PIDSI_THUMBNAIL` constant is 17; the ID written in the
        // stream is 16, and a stream carrying the wrong one carries no picture.
        let mut picture = Vec::new();
        picture.extend_from_slice(&MM_ANISOTROPIC.to_le_bytes());
        picture.extend_from_slice(&2000u32.to_le_bytes());
        picture.extend_from_slice(&1000u32.to_le_bytes());
        picture.extend_from_slice(&WMF_HEADER);

        let mut stream = summary_information(&picture);
        // The property's ID sits 16 bytes into the section, past its size and
        // its count and the property's own offset.
        stream[48 + 8] = 17;

        assert!(property_set_thumbnail(&stream).is_none());
    }

    #[test]
    fn reads_the_frame_of_an_enhanced_metafile() {
        // An ENHMETAHEADER with a 100 by 50 millimetre frame.
        let mut bytes = vec![0u8; 88];
        bytes[0..4].copy_from_slice(&1u32.to_le_bytes());
        bytes[4..8].copy_from_slice(&88u32.to_le_bytes());
        bytes[24..28].copy_from_slice(&0i32.to_le_bytes()); // frame left
        bytes[28..32].copy_from_slice(&0i32.to_le_bytes()); // frame top
        bytes[32..36].copy_from_slice(&10_000i32.to_le_bytes()); // frame right
        bytes[36..40].copy_from_slice(&5_000i32.to_le_bytes()); // frame bottom
        bytes[40..44].copy_from_slice(&0x464D_4520u32.to_le_bytes()); // " EMF"

        let thumbnail = classify(bytes).expect("a metafile");
        assert!(thumbnail.is_metafile());
        // 100 mm at 96 DPI.
        assert_eq!(thumbnail.width, 378);
        assert_eq!(thumbnail.height, 189);
    }
}

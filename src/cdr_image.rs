//! CorelDRAW documents of the older shape, read for the picture the file keeps of itself.
//!
//! A `.cdr` comes in two shapes and both keep a picture of the drawing beside the drawing
//! itself, which is what this app shows. The newer shape is a zip container, and the
//! picture in it is a member `project_image` reads. The older one — the RIFF container
//! CorelDRAW wrote from version 6 to 12, still the shape of most `.cdr` files met in the
//! wild — is read here.
//!
//! What that container holds is a chunk list, and the picture is the `DISP` chunk: the
//! bitmap CorelDRAW wrote for its own file dialog and for the shell to show, at the size a
//! file manager wants rather than at the size of the drawing. The chunks are walked by the
//! length each one declares, which is what makes the walk seeks rather than a read: what is
//! opened is the one chunk this reader has a question for.
//!
//! Behind the four bytes CorelDRAW writes at the front of that chunk is a
//! device-independent bitmap — the header, the colour table and the rows — and a
//! device-independent bitmap is a `.bmp` with fourteen bytes in front of it. So those bytes
//! are written here rather than a DIB decoded here: the decoder the picture path already
//! has takes a bitmap, and the picture that comes out of it is composed like any other
//! design preview. What that costs is fourteen bytes of copying and nothing compiled in.
//!
//! The picture is small — the ones seen are 96 by 96 and 128 by 128 — so it is a preview
//! that tells a user which file this is, the way the thumbnail a project container holds
//! does; see `project_image` for what that is worth and how it is laid out.

use crate::config::{decode_budget_bytes, image_decode_limits};
use image::GenericImageView;
use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::Path;

/// What every one of these containers opens with, and the form CorelDRAW's own files are
/// written under: `CDRA`, `CDRB` and the rest of that family.
const RIFF: &[u8; 4] = b"RIFF";
const COREL_FORM: &[u8; 3] = b"CDR";
/// The chunk the picture of the document is kept in.
const DISP: &[u8; 4] = b"DISP";
/// How much of a RIFF container is read to find that chunk: a header's worth of each
/// chunk, and the one chunk that is the picture when it is reached.
const CHUNK_HEADER_BYTES: u64 = 8;
/// The four bytes CorelDRAW writes ahead of the bitmap in that chunk. What they mean is
/// CorelDRAW's business; what matters here is where the bitmap starts.
const LEAD_BYTES: usize = 4;
/// The header a bitmap always carries, which is the thing that says the bytes behind it
/// are a bitmap rather than something else CorelDRAW put in a chunk of the same name.
const BITMAP_HEADER_BYTES: u32 = 40;
/// The header a `.bmp` file carries in front of that one.
const BMP_FILE_HEADER_BYTES: usize = 14;

/// The size of the picture the document keeps of itself.
pub fn dimensions(path: &Path) -> Option<(u32, u32)> {
    let (bitmap, _) = bitmap_of(path, false)?;

    Some((bitmap.width, bitmap.height))
}

/// The picture the document keeps of itself, decoded into `width` by `height` and handed
/// back as BGRA, top down, the way every frame in this app is composed.
pub fn decode(path: &Path, width: u32, height: u32) -> Option<Vec<u8>> {
    if width == 0 || height == 0 {
        return None;
    }

    let (_, bitmap) = bitmap_of(path, true)?;

    let mut reader = image::ImageReader::new(std::io::Cursor::new(&bitmap))
        .with_guessed_format()
        .ok()?;
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

/// The bitmap the document's `DISP` chunk opens with: its header always, and the whole of
/// it — the colour table and the rows behind the header — when `whole` asks for the
/// picture rather than for its size.
///
/// The picture is handed back as the file a decoder reads rather than as the DIB the chunk
/// holds: the fourteen bytes a `.bmp` opens with are written here, because what is left of
/// a bitmap once they are, is what a DIB is.
fn bitmap_of(path: &Path, whole: bool) -> Option<(Bitmap, Vec<u8>)> {
    let mut file = File::open(path).ok()?;
    let bytes = file.metadata().ok()?.len();
    let (chunk_at, chunk_bytes) = disp_chunk(&mut file, bytes)?;

    file.seek(SeekFrom::Start(chunk_at + LEAD_BYTES as u64)).ok()?;
    // The header is what says the chunk holds a bitmap, so it is read whether or not the
    // picture is wanted; everything behind it is read only when the picture is.
    let mut header_bytes = vec![0u8; BITMAP_HEADER_BYTES as usize];
    file.read_exact(&mut header_bytes).ok()?;
    let bitmap = Bitmap::read(&header_bytes)?;

    if !whole {
        return Some((bitmap, header_bytes));
    }

    let budget = decode_budget_bytes();
    let dib_bytes = bitmap.dib_bytes()?;
    // What the chunk declares is what it holds: a picture longer than the chunk it was read
    // out of is a file that does not add up, and the rows are not read for it.
    if dib_bytes as u64 > chunk_bytes - LEAD_BYTES as u64 || dib_bytes as u64 > budget {
        return None;
    }

    let mut dib = vec![0u8; dib_bytes];
    dib[..BITMAP_HEADER_BYTES as usize].copy_from_slice(&header_bytes);
    file.read_exact(&mut dib[BITMAP_HEADER_BYTES as usize..]).ok()?;
    let picture = bitmap_file(&bitmap, &dib)?;

    Some((bitmap, picture))
}

/// Where the picture is kept: where the `DISP` chunk's own bytes start and how many of them
/// there are, found by walking the container's chunk list the way any RIFF reader walks one
/// — each chunk declaring its own length, and a chunk that declares a length past the end of
/// the file ending the walk rather than being followed.
fn disp_chunk(file: &mut File, bytes: u64) -> Option<(u64, u64)> {
    let mut opening = [0u8; 12];
    file.seek(SeekFrom::Start(0)).ok()?;
    file.read_exact(&mut opening).ok()?;

    if &opening[0..4] != RIFF || &opening[8..11] != COREL_FORM {
        return None;
    }

    let mut at = 12u64;
    let mut header = [0u8; CHUNK_HEADER_BYTES as usize];
    while at + CHUNK_HEADER_BYTES <= bytes {
        file.seek(SeekFrom::Start(at)).ok()?;
        file.read_exact(&mut header).ok()?;
        let length = u32::from_le_bytes([header[4], header[5], header[6], header[7]]) as u64;

        if &header[0..4] == DISP {
            return (length >= LEAD_BYTES as u64 + BITMAP_HEADER_BYTES as u64)
                .then_some((at + CHUNK_HEADER_BYTES, length));
        }

        // A chunk is written padded to an even length, and the padding is not part of it.
        at += CHUNK_HEADER_BYTES + length + (length % 2);
    }

    None
}

/// The header of a device-independent bitmap, read for the four things this reader has a
/// question about: how large the picture is, how many bits a sample is, how many colour
/// table entries are written before the rows, and how long the whole of it is.
struct Bitmap {
    width: u32,
    height: u32,
    /// The colour table entries written before the rows, each four bytes of blue, green,
    /// red and nothing.
    palette_entries: u32,
    bits: u32,
}

impl Bitmap {
    fn read(bytes: &[u8]) -> Option<Self> {
        if bytes.len() < BITMAP_HEADER_BYTES as usize {
            return None;
        }
        if u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]) != BITMAP_HEADER_BYTES {
            return None;
        }

        let width = i32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]);
        let height = i32::from_le_bytes([bytes[8], bytes[9], bytes[10], bytes[11]]);
        let bits = u16::from_le_bytes([bytes[14], bytes[15]]) as u32;
        if width <= 0 || height == 0 || !matches!(bits, 1 | 4 | 8 | 16 | 24 | 32) {
            return None;
        }

        // A table with nothing declared holds every entry the depth can name.
        let declared = u32::from_le_bytes([bytes[32], bytes[33], bytes[34], bytes[35]]);
        let palette_entries = if declared > 0 {
            declared
        } else if bits <= 8 {
            1 << bits
        } else {
            0
        };

        Some(Self {
            width: width as u32,
            // A negative height is a picture written the way this app holds one, top row
            // first, rather than the bottom-up order a bitmap is usually written in.
            height: height.unsigned_abs(),
            palette_entries,
            bits,
        })
    }

    /// How long the header, the colour table and the rows are together.
    fn dib_bytes(&self) -> Option<usize> {
        let row = (self.width as usize)
            .checked_mul(self.bits as usize)?
            .div_ceil(32)
            .checked_mul(4)?;
        let rows = row.checked_mul(self.height as usize)?;
        let palette = (self.palette_entries as usize).checked_mul(4)?;

        (BITMAP_HEADER_BYTES as usize)
            .checked_add(palette)?
            .checked_add(rows)
    }
}

/// The bitmap as the file a decoder reads: the same header, colour table and rows, with
/// the fourteen bytes a `.bmp` carries in front of them — the two letters it opens with,
/// how long the file is, four bytes that mean nothing, and where the rows start.
fn bitmap_file(bitmap: &Bitmap, dib: &[u8]) -> Option<Vec<u8>> {
    let rows_at = BMP_FILE_HEADER_BYTES + BITMAP_HEADER_BYTES as usize
        + bitmap.palette_entries as usize * 4;
    let total = BMP_FILE_HEADER_BYTES.checked_add(dib.len())?;

    let mut file = Vec::with_capacity(total);
    file.extend_from_slice(b"BM");
    file.extend_from_slice(&u32::try_from(total).ok()?.to_le_bytes());
    file.extend_from_slice(&[0u8; 4]);
    file.extend_from_slice(&u32::try_from(rows_at).ok()?.to_le_bytes());
    file.extend_from_slice(dib);

    Some(file)
}

#[cfg(test)]
mod tests {
    use super::*;
    use image::ImageEncoder;
    use std::io::Write;

    /// The chunk a document keeps its picture in, as CorelDRAW writes one: four bytes
    /// this reader has no use for, and the bitmap behind them.
    fn disp_chunk_of(width: u32, height: u32) -> Vec<u8> {
        let mut file = Vec::new();
        image::codecs::bmp::BmpEncoder::new(&mut file)
            .write_image(
                &vec![0u8; (width * height * 3) as usize],
                width,
                height,
                image::ExtendedColorType::Rgb8,
            )
            .expect("a bitmap");

        let mut chunk = vec![8u8, 0, 0, 0];
        chunk.extend_from_slice(&file[BMP_FILE_HEADER_BYTES..]);

        chunk
    }

    /// A document of that shape: a RIFF container with a version chunk, the picture, and a
    /// chunk behind it that a walk has to step over to be finished.
    fn write_document(path: &Path, width: u32, height: u32) {
        let picture = disp_chunk_of(width, height);

        let mut bytes = Vec::new();
        bytes.extend_from_slice(RIFF);
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(b"CDRB");
        bytes.extend_from_slice(b"vrsn");
        bytes.extend_from_slice(&2u32.to_le_bytes());
        bytes.extend_from_slice(&[0x05, 0x00]);
        bytes.extend_from_slice(b"DISP");
        bytes.extend_from_slice(&(picture.len() as u32).to_le_bytes());
        bytes.extend_from_slice(&picture);
        bytes.extend_from_slice(b"sumi");
        bytes.extend_from_slice(&8u32.to_le_bytes());
        bytes.extend_from_slice(&[0u8; 8]);

        let size = (bytes.len() - 8) as u32;
        bytes[4..8].copy_from_slice(&size.to_le_bytes());

        let mut file = File::create(path).expect("a document to write");
        file.write_all(&bytes).expect("bytes");
    }

    fn temp_path(name: &str) -> std::path::PathBuf {
        std::env::temp_dir().join(name)
    }

    /// The picture a document keeps of itself is read for its size, which is the size the
    /// layout is asked to place — a preview that tells a user which file this is, at the
    /// size the file keeps it.
    #[test]
    fn reads_the_size_of_the_picture_it_keeps() {
        let path = temp_path("rust-hover-preview-cdr-size.cdr");
        write_document(&path, 7, 5);

        assert_eq!(dimensions(&path), Some((7, 5)));

        std::fs::remove_file(&path).ok();
    }

    /// And it is decoded into the box it is asked for, drawn the way every other design
    /// preview is: the bitmap the chunk holds is handed to the decoder as the file it is
    /// one header away from being.
    #[test]
    fn draws_the_picture_it_keeps_into_the_box_it_is_asked_for() {
        let path = temp_path("rust-hover-preview-cdr-draw.cdr");
        write_document(&path, 7, 5);

        let frame = decode(&path, 14, 10).expect("a document with a picture in it");
        assert_eq!(frame.len(), 14 * 10 * 4);
        assert_eq!(frame[3], 255, "and a bitmap has no transparency to carry");

        std::fs::remove_file(&path).ok();
    }

    /// A file that is not one of these is refused by its own header rather than by its
    /// name, whatever it is called: what a file is, is its own answer.
    #[test]
    fn refuses_a_file_that_is_not_one_of_these() {
        let path = temp_path("rust-hover-preview-cdr-refused.cdr");
        std::fs::write(&path, b"PK\x03\x04not a cdr at all").unwrap();

        assert_eq!(dimensions(&path), None);
        assert_eq!(decode(&path, 10, 10), None);

        std::fs::remove_file(&path).ok();
    }
}

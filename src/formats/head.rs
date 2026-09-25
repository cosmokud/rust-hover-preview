//! The head of a file: what its own first bytes say.
//!
//! Read once, for every question about the front of a file. The content tables ask what
//! kind a file is, the readers ask which of them is asked about it, and the layout asks
//! whether what it holds moves — and all three are answered from one read of one window,
//! held under the file's name and version, rather than each opening the file for itself.
//!
//! The window is read to the front of a file first, and only read further when that front
//! settles nothing: the picture families all declare themselves in the first bytes, so a
//! hover onto one costs sixteen bytes of it, and a file whose own form says nothing is the
//! file the whole window was always read for. What the front says is kept as facts —
//! nothing here depends on the configuration — so a list edited between hovers is a list
//! the next hover is decided by.

use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Read, Seek, SeekFrom};
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};
use std::time::SystemTime;

use once_cell::sync::Lazy;

/// How much of a file is read before it is decided whether the rest of it is wanted:
/// enough for every signature written at the very start of a file, which is where the
/// picture families are. RIFF's form type — `WEBP` against `AVI` — is why it is sixteen
/// and not eight.
const SHORT_BYTES: usize = 16;

/// How far a file is read where its front settles nothing, and the window every table is
/// asked about. Nothing further is ever read: every signature this app knows is in the
/// head of a file, and a hover onto a video costs this and not the video.
pub(crate) const PROBE_BYTES: usize = 4096;

/// How many heads are held. What it bounds is a pointer dragged across a large folder.
const HEADS_MAX_ENTRIES: usize = 512;

/// A file being walked by one of the chunk walks below.
type Reader<'a> = BufReader<&'a mut File>;

/// What a file's own head says about the picture in it, where its name cannot say.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) struct PictureNature {
    pub(crate) form: PictureForm,
    /// Whether what the file holds is a sequence in time, which is what the animated scale
    /// is for — as against a picture that merely has more than one page in it.
    pub(crate) moves: bool,
    /// Whether the container is one a camera raw is written in. Which names a raw answers
    /// with, and so which reader is asked about it, is a question the byte tables answer
    /// rather than this one (see `content_type::classify`), so a front of that kind is a
    /// front that settles nothing on its own.
    pub(crate) container_of_a_raw: bool,
}

/// The form a picture's own bytes declare.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum PictureForm {
    /// One frame, or a container with one picture in it.
    Still,
    /// A sequence this app plays: the animated reader for the family is asked first.
    Plays(PictureFamily),
    /// A sequence this app has no reader for: what is drawn is its first frame.
    Unplayable,
    /// More than one page or size in a container that is otherwise still: the first one is
    /// the picture, and the pages after it are not a thing that moves.
    Paged,
}

/// The animated families this app reads itself.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum PictureFamily {
    Gif,
    Webp,
    Apng,
}

/// The head of one file: what was read of it, and what that says.
pub(crate) struct Head {
    bytes: Vec<u8>,
    complete: bool,
    front: Option<&'static [&'static str]>,
    nature: Option<PictureNature>,
}

impl Head {
    /// The window the content tables are asked about.
    pub(crate) fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Whether what is held is the whole window a table is asked about, rather than the front
    /// of a file. A front that settled the question is held as the front it is — the rest of
    /// the window is read only when something asks for it (see `full`).
    pub(crate) fn complete(&self) -> bool {
        self.complete
    }

    /// The names the front answers with, where the form the file is in says what it is
    /// without the signature tables.
    pub(crate) fn front(&self) -> Option<&'static [&'static str]> {
        self.front
    }

    /// What the front says about the picture in the file, where the file is one.
    pub(crate) fn nature(&self) -> Option<PictureNature> {
        self.nature
    }
}

/// The file and the version of it a head was read from. A file saved again is a file whose
/// head is read again rather than held to what it was.
#[derive(Clone, PartialEq, Eq, Hash)]
pub(crate) struct Key {
    path: PathBuf,
    modified: Option<SystemTime>,
    len: u64,
}

/// The head of the file at `path`, read to the front of it and no further where the front
/// settles the question.
pub(crate) fn of(path: &Path) -> Option<Arc<Head>> {
    let key = key(path);

    if let Some(head) = held(&key) {
        return Some(head);
    }

    let head = Arc::new(read(path, false)?);

    Some(keep(key, head))
}

/// The head of the file at `path`, read to the whole window the tables are asked about.
pub(crate) fn full(path: &Path) -> Option<Arc<Head>> {
    let key = key(path);

    if let Some(head) = held(&key) {
        if head.complete() {
            return Some(head);
        }
    }

    let head = Arc::new(read(path, true)?);

    Some(keep(key, head))
}

/// What the head of a picture says about it, where the file's own form is one this module
/// reads. Nothing is decoded, and nothing past the head is read.
pub(crate) fn picture_nature(path: &Path) -> Option<PictureNature> {
    of(path).and_then(|head| head.nature)
}

/// The key a file's head is held under.
pub(crate) fn key(path: &Path) -> Key {
    let metadata = std::fs::metadata(path).ok();

    Key {
        path: path.to_path_buf(),
        modified: metadata
            .as_ref()
            .and_then(|metadata| metadata.modified().ok()),
        len: metadata
            .as_ref()
            .map(|metadata| metadata.len())
            .unwrap_or(0),
    }
}

static HEADS: Lazy<Mutex<HashMap<Key, Arc<Head>>>> = Lazy::new(|| Mutex::new(HashMap::new()));

fn held(key: &Key) -> Option<Arc<Head>> {
    HEADS.lock().ok()?.get(key).map(Arc::clone)
}

fn keep(key: Key, head: Arc<Head>) -> Arc<Head> {
    if let Ok(mut heads) = HEADS.lock() {
        if heads.len() >= HEADS_MAX_ENTRIES {
            heads.clear();
        }
        heads.insert(key, Arc::clone(&head));
    }

    head
}

fn read(path: &Path, whole: bool) -> Option<Head> {
    // A file whose content is not on this machine is not opened: reading its front is what
    // would bring it down. What it is called is the whole of what is known about it.
    if crate::shell::cloud_files::needs_download(path) {
        return None;
    }

    let mut file = File::open(path).ok()?;
    let mut bytes = Vec::new();
    let front_bytes = if whole { PROBE_BYTES } else { SHORT_BYTES };
    (&mut file)
        .take(front_bytes as u64)
        .read_to_end(&mut bytes)
        .ok()?;

    let (front, nature) = facts(&mut file, &bytes);

    // A front whose form this module knows is the whole answer — what it says does not change
    // for the rest of the window being read — so nothing more of it is read. It is still held
    // as the front it is: the tables are asked about the whole window and the rest is read when
    // something asks for them (see `full`). A front that says nothing is a file whose signature
    // may be further in, which is what the rest of the window is for.
    let settled = front.is_some() || nature.is_some();
    let mut complete = whole || bytes.len() < front_bytes;
    if !whole && !settled && !complete {
        let at = bytes.len() as u64;
        let rest = (PROBE_BYTES as u64).saturating_sub(at);
        if file.seek(SeekFrom::Start(at)).is_ok()
            && (&mut file).take(rest).read_to_end(&mut bytes).is_ok()
        {
            complete = true;
        }
    }

    Some(Head {
        bytes,
        complete,
        front,
        nature,
    })
}

/// What the front of a file says: the names its own form answers with, and — where it is a
/// picture — the form the picture is in.
///
/// The walks that some of these need go past the front and are given the whole file for it
/// (a PNG's `acTL` may sit behind an ancillary chunk of any size), but they are walks by
/// length and not reads: no pixels are decoded to answer any of this.
fn facts(
    file: &mut File,
    front: &[u8],
) -> (Option<&'static [&'static str]>, Option<PictureNature>) {
    const PNG: [u8; 8] = [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];

    if front.starts_with(&PNG) {
        let animated = png_is_animated(&mut BufReader::new(&mut *file));
        let form = if animated {
            PictureForm::Plays(PictureFamily::Apng)
        } else {
            PictureForm::Still
        };

        return (Some(&["png"]), Some(nature(form, animated, false)));
    }

    if front.starts_with(b"GIF87a") || front.starts_with(b"GIF89a") {
        let animated = gif_is_animated(&mut BufReader::new(&mut *file));
        let form = if animated {
            PictureForm::Plays(PictureFamily::Gif)
        } else {
            PictureForm::Still
        };

        return (Some(&["gif"]), Some(nature(form, animated, false)));
    }

    if front.len() >= 12 && &front[..4] == b"RIFF" && &front[8..12] == b"WEBP" {
        let animated = webp_is_animated(&mut BufReader::new(&mut *file));
        let form = if animated {
            PictureForm::Plays(PictureFamily::Webp)
        } else {
            PictureForm::Still
        };

        return (Some(&["webp"]), Some(nature(form, animated, false)));
    }

    // The ISO base media family, where the brand at the front of the file is what says
    // whether it is one picture or a sequence of them. A brand this module does not know
    // is a front that settles nothing, which is what leaves a `.mov` to the tables.
    if front.len() >= 12 && &front[4..8] == b"ftyp" {
        return match &front[8..12] {
            // An AVIF sequence, and the sequence form of HEIF: what moves is a sequence in
            // time that this app has no reader for.
            b"avis" | b"msf1" => (
                Some(&["avif"]),
                Some(nature(PictureForm::Unplayable, true, false)),
            ),
            b"avif" | b"av01" => (
                Some(&["avif"]),
                Some(nature(PictureForm::Still, false, false)),
            ),
            b"heic" | b"heix" | b"hevc" | b"hevx" | b"mif1" | b"heim" => (
                Some(&["heic"]),
                Some(nature(PictureForm::Still, false, false)),
            ),
            b"avci" => (
                Some(&["avci"]),
                Some(nature(PictureForm::Still, false, false)),
            ),
            _ => (None, None),
        };
    }

    // A TIFF is the one still container a camera raw is written in — a `.dng`, a `.cr2`, a
    // `.nef` are TIFFs — so its front is a front that settles nothing: whether it is a
    // picture or a raw behind a picture's name is the tables' question.
    if front.starts_with(b"II*\0") || front.starts_with(b"MM\0*") {
        let paged =
            tiff_holds_more_than_one_page(&mut BufReader::new(&mut *file), front[0] == b'I');
        let form = if paged {
            PictureForm::Paged
        } else {
            PictureForm::Still
        };

        return (Some(&["tif"]), Some(nature(form, false, true)));
    }

    // An icon or a cursor: the count of images it holds is in its own header, and more than
    // one of them is a set of sizes rather than a thing that moves.
    if front.len() >= 6
        && (front[..4] == [0x00, 0x00, 0x01, 0x00] || front[..4] == [0x00, 0x00, 0x02, 0x00])
    {
        let cursor = front[..4] == [0x00, 0x00, 0x02, 0x00];
        let count = u16::from_le_bytes([front[4], front[5]]);
        let form = if count > 1 {
            PictureForm::Paged
        } else {
            PictureForm::Still
        };
        let names: &'static [&'static str] = if cursor { &["cur"] } else { &["ico"] };

        return (Some(names), Some(nature(form, false, false)));
    }

    let names: &'static [&'static str] = if front.starts_with(&[0xFF, 0xD8]) {
        // A JPEG, and whether it is one picture or the two of a stereo pair is written in
        // an `APP2` segment that a segment of any size may sit in front of; what is read
        // of it is its first frame, the same as before (see `TODO.md`).
        &["jpg"]
    } else if front.starts_with(&[0x8A, b'M', b'N', b'G']) {
        // A multiple-image PNG: an animation, and one this app has no reader for — it is
        // the engine that draws one, and it draws its first frame.
        return (
            Some(&["mng"]),
            Some(nature(PictureForm::Unplayable, true, false)),
        );
    } else if front.starts_with(b"DDS ") {
        // Whether the surfaces behind the first one are a mip chain, a cube map or a
        // volume is in the header's own flags; what is drawn is the first surface.
        &["dds"]
    } else if front.starts_with(b"SIMPLE  =") {
        // A FITS, whose header records say how many axes it has beyond the picture's two.
        &["fits"]
    } else if front.starts_with(b"gimp xcf ") {
        // An XCF, whose layers are its own structure and not pages.
        &["xcf"]
    } else {
        return (None, None);
    };

    (Some(names), Some(nature(PictureForm::Still, false, false)))
}

const fn nature(form: PictureForm, moves: bool, container_of_a_raw: bool) -> PictureNature {
    PictureNature {
        form,
        moves,
        container_of_a_raw,
    }
}

/// Whether a PNG holds an animation control chunk before its image data.
///
/// An APNG is an ordinary PNG plus an `acTL` chunk ahead of its first `IDAT`, and the
/// animation chunks are defined to precede the image data — so the chunk list is walked,
/// and the walk stops where the pixels start. Nothing is decoded.
fn png_is_animated(reader: &mut Reader<'_>) -> bool {
    if reader.seek(SeekFrom::Start(8)).is_err() {
        return false;
    }

    let mut header = [0u8; 8];
    while reader.read_exact(&mut header).is_ok() {
        let length = u32::from_be_bytes([header[0], header[1], header[2], header[3]]) as i64;

        if &header[4..8] == b"acTL" {
            return true;
        }
        if &header[4..8] == b"IDAT" || &header[4..8] == b"IEND" {
            return false;
        }

        // Step over the chunk body and its CRC.
        if reader.seek(SeekFrom::Current(length + 4)).is_err() {
            return false;
        }
    }

    false
}

/// Whether a GIF holds more than its first frame, which is the whole of what makes one an
/// animation rather than a picture: the frames are the file's own blocks, and whether there
/// is a second one is a thing its structure says.
///
/// Nothing is decoded: the blocks are walked by their own lengths — an extension's
/// sub-blocks are stepped over the same way, since they carry lengths too — so what this
/// costs is a few seeks and no pixels.
fn gif_is_animated(reader: &mut Reader<'_>) -> bool {
    if reader.seek(SeekFrom::Start(0)).is_err() {
        return false;
    }

    // The signature and the logical screen descriptor: `GIF87a` or `GIF89a`, then seven
    // bytes of screen size, colour table and background.
    let mut header = [0u8; 13];
    if reader.read_exact(&mut header).is_err() || &header[..3] != b"GIF" {
        return false;
    }

    // A global colour table follows the descriptor, and the frame blocks after it: `0x2C`
    // opens one and carries its own bounds and local table, `0x21` opens an extension whose
    // sub-blocks carry lengths, and `0x3B` ends the file.
    if header[10] & 0x80 != 0 {
        let table_bytes = 3 * (1u64 << ((header[10] & 0x07) + 1));
        if reader.seek(SeekFrom::Current(table_bytes as i64)).is_err() {
            return false;
        }
    }

    let mut frames = 0usize;
    let mut block = [0u8; 1];

    while reader.read_exact(&mut block).is_ok() {
        match block[0] {
            0x2C => {
                frames += 1;
                if frames > 1 {
                    return true;
                }

                // The frame's bounds and flags: the local colour table, where it has one,
                // sits between the flags and the frame's pixels.
                let mut descriptor = [0u8; 9];
                if reader.read_exact(&mut descriptor).is_err() {
                    return false;
                }
                if descriptor[8] & 0x80 != 0 {
                    let table_bytes = 3 * (1u64 << ((descriptor[8] & 0x07) + 1));
                    if reader.seek(SeekFrom::Current(table_bytes as i64)).is_err() {
                        return false;
                    }
                }

                // The LZW code size byte, then the pixel data as sub-blocks.
                if reader.read_exact(&mut block).is_err() || !skip_gif_sub_blocks(reader) {
                    return false;
                }
            }
            0x21 => {
                // An extension's label byte, then its own sub-blocks.
                if reader.read_exact(&mut block).is_err() || !skip_gif_sub_blocks(reader) {
                    return false;
                }
            }
            0x3B => return false,
            // Anything else is not where the next block can be, so the walk is over.
            _ => return false,
        }
    }

    false
}

/// Step a GIF reader over one run of sub-blocks: a length byte per block, ending at a
/// zero-length one. The pixels and the extensions a specimen of either is skipped by are
/// both held this way, so the one walk serves both.
fn skip_gif_sub_blocks(reader: &mut Reader<'_>) -> bool {
    let mut length = [0u8; 1];
    loop {
        if reader.read_exact(&mut length).is_err() {
            return false;
        }
        if length[0] == 0 {
            return true;
        }
        if reader.seek(SeekFrom::Current(length[0] as i64)).is_err() {
            return false;
        }
    }
}

/// Whether a WebP holds an animation, which its container says: an extended WebP that
/// animates carries an `ANIM` chunk, and its frames are `ANMF` chunks after it.
///
/// The chunks are walked by their own lengths for the same reason the GIF's blocks are:
/// what is asked is what the file holds, and nothing has to be decoded to answer it. A
/// plain `VP8 ` or `VP8L` WebP has no chunk list at all — its picture data is the chunk
/// itself — so the walk stops where the picture starts.
fn webp_is_animated(reader: &mut Reader<'_>) -> bool {
    if reader.seek(SeekFrom::Start(0)).is_err() {
        return false;
    }

    let mut header = [0u8; 12];
    if reader.read_exact(&mut header).is_err() || &header[..4] != b"RIFF" || &header[8..] != b"WEBP"
    {
        return false;
    }

    let mut chunk = [0u8; 8];
    while reader.read_exact(&mut chunk).is_ok() {
        let length = u32::from_le_bytes([chunk[4], chunk[5], chunk[6], chunk[7]]) as i64;

        if &chunk[..4] == b"ANIM" {
            return true;
        }

        // The picture data is the end of anything worth walking: an extended WebP that
        // animates names its animation ahead of its frames, and a plain one is nothing but
        // the picture.
        if &chunk[..4] == b"VP8 " || &chunk[..4] == b"VP8L" {
            return false;
        }

        // Chunks are padded to an even length.
        if reader.seek(SeekFrom::Current(length + length % 2)).is_err() {
            return false;
        }
    }

    false
}

/// Whether a TIFF's first image directory points at another one, which is what a second
/// page of one is. The directory is read for its own entry count and the four bytes after
/// its entries, and nothing of the pictures inside it is read at all.
fn tiff_holds_more_than_one_page(reader: &mut Reader<'_>, little_endian: bool) -> bool {
    let u16_of = |bytes: [u8; 2]| {
        if little_endian {
            u16::from_le_bytes(bytes)
        } else {
            u16::from_be_bytes(bytes)
        }
    };
    let u32_of = |bytes: [u8; 4]| {
        if little_endian {
            u32::from_le_bytes(bytes)
        } else {
            u32::from_be_bytes(bytes)
        }
    };

    let mut first = [0u8; 4];
    if reader.seek(SeekFrom::Start(4)).is_err() || reader.read_exact(&mut first).is_err() {
        return false;
    }
    let first = u32_of(first) as u64;
    if first == 0 {
        return false;
    }

    let mut count = [0u8; 2];
    if reader.seek(SeekFrom::Start(first)).is_err() || reader.read_exact(&mut count).is_err() {
        return false;
    }
    let entries = u16_of(count) as u64;

    let mut next = [0u8; 4];
    let after_entries = first + 2 + entries * 12;
    if reader.seek(SeekFrom::Start(after_entries)).is_err() || reader.read_exact(&mut next).is_err()
    {
        return false;
    }

    u32_of(next) != 0
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;

    /// A file of the bytes given, under a name, in a folder of its own.
    fn sample(folder: &str, name: &str, bytes: &[u8]) -> PathBuf {
        let folder = std::env::temp_dir().join(format!("rust-hover-preview-head-{folder}"));
        std::fs::create_dir_all(&folder).expect("a test folder");
        let path = folder.join(name);
        std::fs::write(&path, bytes).expect("a test file");

        path
    }

    /// The eight bytes every PNG opens with.
    pub(crate) fn png_open() -> Vec<u8> {
        vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A]
    }

    /// A chunk as a PNG writes one: its length, its type, its body and its CRC.
    pub(crate) fn png_chunk(kind: &[u8; 4], body: &[u8]) -> Vec<u8> {
        let mut chunk = (body.len() as u32).to_be_bytes().to_vec();
        chunk.extend_from_slice(kind);
        chunk.extend_from_slice(body);
        chunk.extend_from_slice(&[0, 0, 0, 0]);

        chunk
    }

    /// A GIF's header and logical screen descriptor, with no global colour table.
    pub(crate) fn gif_open() -> Vec<u8> {
        let mut gif = b"GIF89a".to_vec();
        gif.extend_from_slice(&[0x01, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00]);

        gif
    }

    /// A GIF image block: its descriptor, its code size and one run of sub-blocks.
    pub(crate) fn gif_frame() -> Vec<u8> {
        let mut frame = vec![0x2C];
        frame.extend_from_slice(&[0u8; 9]);
        frame.push(0x02);
        frame.extend_from_slice(&[0x03, 0x00, 0x00, 0x00, 0x00]);

        frame
    }

    /// A WebP's container header and one chunk.
    fn webp_chunk(kind: &[u8; 4], body: &[u8]) -> Vec<u8> {
        let mut chunk = kind.to_vec();
        chunk.extend_from_slice(&(body.len() as u32).to_le_bytes());
        chunk.extend_from_slice(body);

        chunk
    }

    /// A TIFF whose first image directory sits at eight and holds no entries, with the four
    /// bytes after them pointing at a second directory or at nothing.
    fn tiff(next_page: bool) -> Vec<u8> {
        let mut tiff = b"II*\0".to_vec();
        tiff.extend_from_slice(&8u32.to_le_bytes());
        tiff.extend_from_slice(&0u16.to_le_bytes());
        tiff.extend_from_slice(&if next_page { 12u32 } else { 0u32 }.to_le_bytes());

        tiff
    }

    #[test]
    fn a_still_picture_and_one_that_moves_are_told_apart_by_their_own_bytes() {
        let still_png = sample(
            "still-png",
            "picture.png",
            &[
                png_open(),
                png_chunk(b"IHDR", &[0u8; 13]),
                png_chunk(b"IDAT", &[0]),
            ]
            .concat(),
        );
        let animated_png = sample(
            "animated-png",
            "picture.png",
            &[
                png_open(),
                png_chunk(b"IHDR", &[0u8; 13]),
                png_chunk(b"acTL", &[0u8; 8]),
                png_chunk(b"IDAT", &[0]),
            ]
            .concat(),
        );

        assert_eq!(
            picture_nature(&still_png).map(|nature| nature.form),
            Some(PictureForm::Still)
        );
        assert_eq!(
            picture_nature(&animated_png),
            Some(nature(PictureForm::Plays(PictureFamily::Apng), true, false)),
            "the same `.png` name, and its own `acTL` chunk is what says which it is"
        );

        let still_gif = sample(
            "still-gif",
            "picture.gif",
            &[gif_open(), gif_frame(), vec![0x3B]].concat(),
        );
        let animated_gif = sample(
            "animated-gif",
            "picture.gif",
            &[gif_open(), gif_frame(), gif_frame(), vec![0x3B]].concat(),
        );

        assert_eq!(
            picture_nature(&still_gif).map(|nature| nature.form),
            Some(PictureForm::Still)
        );
        assert_eq!(
            picture_nature(&animated_gif).map(|nature| nature.form),
            Some(PictureForm::Plays(PictureFamily::Gif))
        );

        let mut still_webp = b"RIFF".to_vec();
        still_webp.extend_from_slice(&0u32.to_le_bytes());
        still_webp.extend_from_slice(b"WEBP");
        still_webp.extend_from_slice(&webp_chunk(b"VP8 ", &[0; 8]));

        let mut animated_webp = b"RIFF".to_vec();
        animated_webp.extend_from_slice(&0u32.to_le_bytes());
        animated_webp.extend_from_slice(b"WEBP");
        animated_webp.extend_from_slice(&webp_chunk(b"VP8X", &[0; 10]));
        animated_webp.extend_from_slice(&webp_chunk(b"ANIM", &[0; 6]));

        assert_eq!(
            picture_nature(&sample("still-webp", "picture.webp", &still_webp))
                .map(|nature| nature.form),
            Some(PictureForm::Still)
        );
        assert_eq!(
            picture_nature(&sample("animated-webp", "picture.webp", &animated_webp)),
            Some(nature(PictureForm::Plays(PictureFamily::Webp), true, false))
        );
    }

    #[test]
    fn a_sequence_this_app_cannot_play_is_still_recognised_as_one() {
        let mut sequence = b"\x00\x00\x00\x20".to_vec();
        sequence.extend_from_slice(b"ftypavis");
        sequence.extend_from_slice(&[0u8; 16]);

        let mut single = b"\x00\x00\x00\x20".to_vec();
        single.extend_from_slice(b"ftypavif");
        single.extend_from_slice(&[0u8; 16]);

        assert_eq!(
            picture_nature(&sample("avis", "picture.avif", &sequence)),
            Some(nature(PictureForm::Unplayable, true, false)),
            "an `avis` brand is a sequence in time, and it is drawn as its first frame"
        );
        assert_eq!(
            picture_nature(&sample("avif", "picture.avif", &single)),
            Some(nature(PictureForm::Still, false, false))
        );

        let mng = sample("mng", "picture.mng", &[0x8A, b'M', b'N', b'G', 0, 0, 0, 0]);
        assert_eq!(
            picture_nature(&mng),
            Some(nature(PictureForm::Unplayable, true, false))
        );
    }

    #[test]
    fn more_than_one_page_is_not_a_thing_that_moves() {
        let single = sample("tiff-one", "page.tif", &tiff(false));
        let paged = sample("tiff-many", "page.tif", &tiff(true));

        assert_eq!(
            picture_nature(&single).map(|nature| nature.form),
            Some(PictureForm::Still)
        );
        assert_eq!(
            picture_nature(&paged).map(|nature| nature.form),
            Some(PictureForm::Paged)
        );

        for path in [&single, &paged] {
            assert!(
                picture_nature(path).is_some_and(|nature| !nature.moves),
                "a TIFF with more pages than one is pages, not an animation"
            );
        }

        let one_icon = sample("ico-one", "icon.ico", &[0x00, 0x00, 0x01, 0x00, 0x01, 0x00]);
        let many_icons = sample(
            "ico-many",
            "icon.ico",
            &[0x00, 0x00, 0x01, 0x00, 0x03, 0x00],
        );

        assert_eq!(
            picture_nature(&one_icon).map(|nature| nature.form),
            Some(PictureForm::Still)
        );
        assert_eq!(
            picture_nature(&many_icons).map(|nature| nature.form),
            Some(PictureForm::Paged)
        );
    }

    #[test]
    fn a_tiff_settles_nothing_on_its_own_because_a_raw_is_written_in_one() {
        let raw = sample("tiff-raw", "shot.cr2", &tiff(false));

        assert!(
            picture_nature(&raw).is_some_and(|nature| nature.container_of_a_raw),
            "the front of a raw is the front of a TIFF, and which reader is asked about \
             one is the byte tables' question"
        );
    }

    #[test]
    fn the_front_is_read_no_further_where_it_settles_the_question() {
        let png = sample(
            "front",
            "picture.png",
            &[
                png_open(),
                png_chunk(b"IHDR", &[0u8; 13]),
                png_chunk(b"IDAT", &[0]),
            ]
            .concat(),
        );

        let front = of(&png).expect("a head");
        assert_eq!(front.bytes().len(), SHORT_BYTES);
        assert!(!front.complete(), "the rest of the window was not read");

        let whole = full(&png).expect("a head");
        assert!(whole.complete());
        assert!(
            whole.bytes().len() > SHORT_BYTES,
            "the whole window is what the tables are asked about"
        );

        // A file whose front is nothing this module knows is read to the whole window,
        // because what it is may be further in than the front.
        let unknown = sample("unknown", "data.bin", &[0x11u8; 64]);
        let head = of(&unknown).expect("a head");
        assert!(head.complete());
        assert_eq!(head.bytes().len(), 64);
    }

    #[test]
    fn both_windows_answer_the_same_question_the_same_way() {
        let png = sample(
            "agree-png",
            "picture.png",
            &[
                png_open(),
                png_chunk(b"IHDR", &[0u8; 13]),
                png_chunk(b"acTL", &[0u8; 8]),
                png_chunk(b"IDAT", &[0u8; 512]),
            ]
            .concat(),
        );
        let gif = sample(
            "agree-gif",
            "picture.gif",
            &[gif_open(), gif_frame(), gif_frame(), vec![0x3B]].concat(),
        );
        let tiff = sample("agree-tiff", "page.tif", &tiff(true));

        for path in [&png, &gif, &tiff] {
            let front = of(path).expect("a head");
            let whole = full(path).expect("a head");

            assert_eq!(
                front.nature(),
                whole.nature(),
                "reading on from the front never changes what the front said"
            );
            assert_eq!(front.front(), whole.front());
        }
    }
}

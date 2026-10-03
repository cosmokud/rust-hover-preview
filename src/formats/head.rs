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
use std::os::windows::fs::MetadataExt;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};
use std::time::SystemTime;

use once_cell::sync::Lazy;

use crate::readers::jxl_image;

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
    /// An AVIF or HEIF image sequence — the ISO base media family, played through the media
    /// engine Windows has rather than by a decoder of this app's own, since HEVC and AV1
    /// have no other one here (see `heif_sequence`).
    Heif,
    /// An animated JPEG XL, which the `jxl-oxide` decoder reads (see `jxl_image`).
    Jxl,
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

impl Key {
    /// The key for a file whose own entry has already been read.
    pub(crate) fn of_metadata(path: &Path, metadata: &std::fs::Metadata) -> Self {
        Self {
            path: path.to_path_buf(),
            modified: metadata.modified().ok(),
            len: metadata.len(),
        }
    }
}

/// What one reading of a file's directory entry settles: that it is there, what version it is
/// at, whether it is a file, and whether its content is on this machine.
///
/// A hover asks a file all of those — is it there, may its content be read, is it one of this
/// app's kinds, is it a file — and every one of them used to be its own question for the
/// volume: five or six reads of the same directory entry, on the thread that draws the hover,
/// for a file whose answers are all in the one entry. What is read here is read once and
/// handed to the questions that follow (see `explorer_hook::normalize_media_path`).
pub(crate) struct Facts {
    key: Key,
    attributes: u32,
    is_file: bool,
}

impl Facts {
    /// Read the file's own entry, or `None` where there is nothing there to read.
    pub(crate) fn read(path: &Path) -> Option<Self> {
        let metadata = std::fs::metadata(path).ok()?;

        Some(Self {
            key: Key::of_metadata(path, &metadata),
            attributes: metadata.file_attributes(),
            is_file: metadata.is_file(),
        })
    }

    /// The version of the file this entry was read at, which is what a head and an answer
    /// about a file's content are held under.
    pub(crate) fn key(&self) -> &Key {
        &self.key
    }

    /// Whether reading the file would have to fetch its content first (see `cloud_files`).
    pub(crate) fn needs_download(&self) -> bool {
        crate::shell::cloud_files::is_remote(self.attributes)
    }

    /// Whether the entry is a file rather than a directory or a device.
    pub(crate) fn is_file(&self) -> bool {
        self.is_file
    }
}

/// The head of the file at `path`, read to the front of it and no further where the front
/// settles the question.
pub(crate) fn of(path: &Path) -> Option<Arc<Head>> {
    held_or_read(path, false, None)
}

/// The same for a caller that has already read the file's own entry, so that the version and
/// the question of whether its content is here are not asked of the volume again (see
/// `Facts`).
pub(crate) fn of_with_facts(path: &Path, facts: &Facts) -> Option<Arc<Head>> {
    held_or_read(path, false, Some(facts))
}

/// The head of the file at `path`, read to the whole window the tables are asked about.
pub(crate) fn full(path: &Path) -> Option<Arc<Head>> {
    held_or_read(path, true, None)
}

/// The same for a caller that has already read the file's own entry (see `Facts`).
pub(crate) fn full_with_facts(path: &Path, facts: &Facts) -> Option<Arc<Head>> {
    held_or_read(path, true, Some(facts))
}

/// The head of a file, from the heads already read where one is held and from the file itself
/// where none is.
fn held_or_read(path: &Path, whole: bool, facts: Option<&Facts>) -> Option<Arc<Head>> {
    let key = facts.map_or_else(|| key(path), |facts| facts.key().clone());

    if let Some(head) = held(&key) {
        if !whole || head.complete() {
            return Some(head);
        }
    }

    let remote = facts.map_or_else(
        || crate::shell::cloud_files::needs_download(path),
        Facts::needs_download,
    );
    let head = Arc::new(read(path, whole, remote)?);

    Some(keep(key, head))
}

/// What the head of a picture says about it, where the file's own form is one this module
/// reads. Nothing is decoded, and nothing past the head is read.
pub(crate) fn picture_nature(path: &Path) -> Option<PictureNature> {
    of(path).and_then(|head| head.nature)
}

/// The key a file's head is held under, read from the file's own entry.
pub(crate) fn key(path: &Path) -> Key {
    std::fs::metadata(path)
        .map(|metadata| Key::of_metadata(path, &metadata))
        .unwrap_or_else(|_| Key {
            path: path.to_path_buf(),
            modified: None,
            len: 0,
        })
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

fn read(path: &Path, whole: bool, remote: bool) -> Option<Head> {
    // A file whose content is not on this machine is not opened: reading its front is what
    // would bring it down. What it is called is the whole of what is known about it.
    if remote {
        return None;
    }

    let mut file = File::open(path).ok()?;
    let mut bytes = Vec::new();
    let front_bytes = if whole { PROBE_BYTES } else { SHORT_BYTES };
    (&mut file)
        .take(front_bytes as u64)
        .read_to_end(&mut bytes)
        .ok()?;

    let (front, nature) = facts(path, &mut file, &bytes);

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
    path: &Path,
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
    //
    // The brand that settles it is the *major* one at the front only most of the time: an
    // encoder is free to name a compatible brand behind a still major brand, and `avis`
    // behind `avif` is the ordinary way an animated AVIF is written. So the whole
    // compatible list is read, and the first brand of either kind that is recognised is
    // the one that decides.
    if front.len() >= 12 && &front[4..8] == b"ftyp" {
        // The compatible brands are past the sixteen bytes the front window holds, so this
        // is a walk of the file and not a read of what is already in hand — the same
        // shape as `webp_is_animated`, and for the same reason: a box of any size may sit
        // in front of the thing being asked about.
        let brands = ftyp_brands(&mut BufReader::new(&mut *file), front);

        // What makes a sequence a sequence is a *sequence* brand, and only that. `mif1` is
        // not one: it is the generic HEIF image brand, and an ordinary single-image HEIC
        // is commonly written with it in front. Treating it as a sequence would give every
        // camera photo the animated scale and a pointless attempt to open it as one.
        //
        // `av01` is likewise a codec brand, not a container brand. It is read only where it
        // is the *major* brand, which is the one position that says what the file is rather
        // than what it was made with: as a compatible brand it appears behind `heic` on
        // files that are HEIC and nothing else, and answering those as AVIF renames them.
        let major = brands.first().copied().unwrap_or([0; 4]);
        let is_sequence = brands
            .iter()
            .any(|brand| matches!(&brand[..], b"avis" | b"msf1"));
        let is_avif = major == *b"av01"
            || brands
                .iter()
                .any(|brand| matches!(&brand[..], b"avis" | b"avif" | b"msf1"));

        if is_sequence {
            // A sequence in time: an AVIF sequence is named `avis`, and the sequence form
            // of HEIF is `msf1`. An AVIF one answers as `avif`; a HEIF one is named after
            // what it codes with, or `heic`.
            let names: &'static [&'static str] = if is_avif { &["avif"] } else { &["heic"] };

            return (
                Some(names),
                Some(nature(PictureForm::Plays(PictureFamily::Heif), true, false)),
            );
        }

        if is_avif {
            return (
                Some(&["avif"]),
                Some(nature(PictureForm::Still, false, false)),
            );
        }

        if brands.iter().any(|brand| {
            matches!(
                &brand[..],
                b"heic" | b"heix" | b"hevc" | b"hevx" | b"mif1" | b"heim"
            )
        }) {
            return (
                Some(&["heic"]),
                Some(nature(PictureForm::Still, false, false)),
            );
        }

        if brands.iter().any(|brand| brand == b"avci") {
            return (
                Some(&["avci"]),
                Some(nature(PictureForm::Still, false, false)),
            );
        }

        return (None, None);
    }

    // A JPEG XL arrives in two forms, and both are common: a naked codestream, which opens
    // with the two signature bytes, and a container, which wraps one in a box. The
    // container is the easy half — a fourcc walk finds the codestream behind `jxlc` — but
    // the question this is asked is not where the pixels are but whether the file is a
    // sequence in time, and that is a flag in the codestream's own image header, behind the
    // box. So the second question is put to a decoder rather than walked.
    //
    // The first is the signature, and it is asked first and of the bytes already in hand,
    // because this arm is reached by every file whose front settled nothing above — which
    // is a great many files that are not JPEG XL at all. Opening one of those to read a
    // header it does not have would put a header parse on this app's hover path for a
    // format it almost never has.
    if jxl_image::begins_like_jxl(front) {
        if jxl_image::is_animated(path) {
            return (
                Some(&["jxl"]),
                Some(nature(PictureForm::Plays(PictureFamily::Jxl), true, false)),
            );
        }

        return (
            Some(&["jxl"]),
            Some(nature(PictureForm::Still, false, false)),
        );
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

/// How many brands an ISO base media file is asked for: the major one and three compatible
/// ones behind it. Past any real file's list, and the bound that keeps a file claiming a
/// very large `ftyp` box from being walked end to end to answer a question its first two
/// entries answer (see `ftyp_brands`).
const FTYP_BRANDS_MAX: usize = 4;

/// The brands an ISO base media file declares: the major one first, then every compatible
/// one behind it, up to a handful.
///
/// The compatible list is the reason this is not simply a read of the bytes at the front.
/// An `ftyp` box is a size, the fourcc `ftyp`, the major brand, a minor version, and then a
/// run of further brands the file says it is also readable as — and an animated AVIF is
/// commonly written with `avif` in front and `avis` in that run. Reading only the first
/// brand is what made an animated AVIF look like a still one.
///
/// The list is past the front window, so it is read out of the file: this is a walk by
/// length, and no pixels are decoded to answer any of it.
///
/// A handful is all that is read, and the bound is what keeps that a cheap question. The
/// size the box declares is the file's own claim, and a file that claims a gigabyte of
/// brands and then has a gigabyte behind it would otherwise be walked four bytes at a
/// time to answer a question its first two entries answer. Four is past any real file's
/// list and well inside what a hover may spend.
///
/// A box that claims no room for anything has no compatible brands in it, which is a still
/// file of the ordinary shape. A file that has run out before the end of its own declared
/// box yields the brands it did declare.
fn ftyp_brands(reader: &mut Reader<'_>, front: &[u8]) -> Vec<[u8; 4]> {
    let mut brands = Vec::new();
    if front.len() < 16 || &front[4..8] != b"ftyp" {
        return brands;
    }

    // The major brand, which is in the front window and is where every file that names one
    // at all names it.
    brands.push([front[8], front[9], front[10], front[11]]);

    // The size the box declares, as the end of the brand list.
    let declared = u32::from_be_bytes([front[0], front[1], front[2], front[3]]) as i64;
    if declared <= 16 {
        return brands;
    }

    if reader.seek(SeekFrom::Start(16)).is_err() {
        return brands;
    }

    let mut at = 16i64;
    while at + 4 <= declared && brands.len() < FTYP_BRANDS_MAX {
        let mut brand = [0u8; 4];
        if reader.read_exact(&mut brand).is_err() {
            break;
        }

        brands.push(brand);
        at += 4;
    }

    brands
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
pub(crate) mod tests;

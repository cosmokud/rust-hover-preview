//! What an archive holds, read from its own metadata and nothing else.
//!
//! A listing is not an extraction. Every format here keeps a table of its own
//! contents — zip at the end of the file, 7z in its header block, RAR in the
//! headers that walk from the front, tar in the 512-byte headers between the
//! data — and reading that table is all this module does. Nothing is
//! decompressed, so a hover onto a 5 GB zip costs what a hover onto a small one
//! does, and the one format that cannot be read that way is bounded instead: a
//! `.tar.gz` has to be inflated to walk its headers, so its stream is capped.
//!
//! Everything a read can do is bounded, in the spirit of the rest of the app: a
//! scan stops at an entry count, a name is cut to a length this side of absurd,
//! and every loop asks the caller's cancel flag, which is the flag the load
//! worker is already handed — so a hover that has moved on stops reading.
//!
//! A format none of those readers has — a cabinet file, an iso, the `.lzh` of an
//! older piece of software — is where the PeaZip engine comes in. It is asked for
//! the archive's table of contents by `peazip_render`, its answer is read here by
//! [`engine_listing`] into the same shape a reader of this app's produces, and it
//! is held under the same key in the same cache — which is what makes the two
//! passes of one hover, and a second hover of the same file, cost one read of
//! the file whatever read it was. Nothing below this line starts anything: what
//! is read here is an answer that has already come back, and a file the engine
//! has not answered for yet is a file with no listing rather than a launch.

use crate::config::config::decode_budget_bytes;
use once_cell::sync::Lazy;
use std::fs::File;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::SystemTime;

mod engine_reports;

pub(crate) use engine_reports::{
    arc_listing, engine_listing, stream_listing, zpaq_listing, zstd_listing,
};

/// Entries one scan will read before it stops counting. A folder of a few
/// archives is a place a pointer is swept over, so the cost of one hover has to
/// be a number this side of what an archive can hold.
const MAX_ENTRIES_SCAN: usize = 20_000;

/// Bytes of a `.tar.gz` stream inflated for one listing. This is the one format
/// whose table is not addressable, so what bounds it is the stream itself.
const MAX_INFLATED_BYTES: u64 = 64 * 1024 * 1024;

/// Characters kept of one entry's name. The name is cut rather than dropped: a
/// name too long to show is still a thing the archive holds.
const MAX_NAME_CHARS: usize = 300;

/// Bytes read to identify a format: enough for a tar header's `ustar` magic and
/// the zip signatures a file can start with.
const PROBE_BYTES: usize = 512 + 8;

/// Listings kept between hovers, keyed by the file and the version of it that was
/// read. A second hover, a repaint, a theme switch or the render pass that
/// follows the measure pass costs a lookup instead of a read.
const LISTING_CACHE_MAX_ENTRIES: usize = 16;

/// The reader a file's own header names. The name it goes by is the fallback:
/// a self-extracting zip starts with `MZ`, and a tar written without the ustar
/// magic is still a tar.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ArchiveKind {
    Zip,
    SevenZ,
    Rar,
    Tar,
    TarGz,
}

/// One thing an archive holds.
#[derive(Debug, Clone)]
pub(crate) struct ArchiveEntry {
    /// The entry's path inside the archive, `/`-separated, without a trailing
    /// slash. Only ever shown, never used as a path.
    pub(crate) name: String,
    pub(crate) size: u64,
    /// The space this entry takes in the archive, where the format says: zip
    /// does per entry, and the formats that compress members against each other
    /// cannot.
    pub(crate) packed: Option<u64>,
    pub(crate) is_dir: bool,
    pub(crate) encrypted: bool,
}

/// What an archive holds, and what could be said about reading it.
#[derive(Debug, Clone)]
pub(crate) struct Listing {
    pub(crate) entries: Vec<ArchiveEntry>,
    /// The archive on disk.
    pub(crate) file_size: u64,
    /// The sum of the entries' own sizes, over the entries that were read.
    pub(crate) total_size: u64,
    /// The same sum over the space the entries take, where the format says.
    pub(crate) packed_total: Option<u64>,
    /// The scan stopped at the entry cap, so the totals are a lower bound.
    pub(crate) scan_capped: bool,
    /// The archive could not be read to the end: what is here is what was
    /// reached.
    pub(crate) read_truncated: bool,
    /// The archive's own file table is encrypted, so there is nothing to list
    /// without a password.
    pub(crate) encrypted_headers: bool,
    /// The archive is one volume of a set.
    pub(crate) multi_volume: bool,
}

impl Listing {
    fn new() -> Self {
        Self {
            entries: Vec::new(),
            file_size: 0,
            total_size: 0,
            packed_total: Some(0),
            scan_capped: false,
            read_truncated: false,
            encrypted_headers: false,
            multi_volume: false,
        }
    }

    /// A listing that could not be made: the file table itself is encrypted.
    fn encrypted_headers(file_size: u64) -> Self {
        Self {
            encrypted_headers: true,
            file_size,
            ..Self::new()
        }
    }

    /// Whether this read has done all the reading it is going to.
    fn at_cap(&self, cancel: Option<&AtomicBool>) -> bool {
        self.entries.len() >= MAX_ENTRIES_SCAN
            || cancel.is_some_and(|cancel| cancel.load(Ordering::Relaxed))
    }

    fn push(
        &mut self,
        name: String,
        size: u64,
        packed: Option<u64>,
        is_dir: bool,
        encrypted: bool,
    ) {
        self.entries.push(ArchiveEntry {
            name,
            size,
            packed,
            is_dir,
            encrypted,
        });
    }

    /// The two totals, over the entries that were read. A format that cannot
    /// state what one entry takes cannot state the sum, and a directory takes
    /// nothing in either column.
    fn settle_totals(&mut self) {
        let mut total = 0u64;
        let mut packed = Some(0u64);

        for entry in self.entries.iter().filter(|entry| !entry.is_dir) {
            total = total.saturating_add(entry.size);
            packed = entry
                .packed
                .and_then(|entry_packed| packed.map(|sum| sum.saturating_add(entry_packed)));
        }

        self.total_size = total;
        self.packed_total = packed;
    }
}

// -------------------------------------------------------------------- reading

/// The listing for `path`, from the cache when the file is unchanged.
pub(crate) fn listing_for(path: &Path, cancel: Option<&AtomicBool>) -> Option<Arc<Listing>> {
    let metadata = std::fs::metadata(path).ok()?;
    let key = ListingKey::of(path, &metadata);

    if let Ok(mut cache) = LISTINGS.lock() {
        if let Some(index) = cache.iter().position(|(cached, _)| *cached == key) {
            let (_, listing) = cache.remove(index);
            cache.insert(0, (key, Arc::clone(&listing)));
            return Some(listing);
        }
    }

    let mut listing = read_listing(path, metadata.len(), cancel)?;
    listing.settle_totals();
    let listing = Arc::new(listing);

    if let Ok(mut cache) = LISTINGS.lock() {
        if cache.len() >= LISTING_CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.insert(0, (key, Arc::clone(&listing)));
    }

    Some(listing)
}

/// Whether a listing for this file is already in hand, without reading anything to find out.
///
/// It is the question an engine is asked before it is asked for a listing — an archive whose
/// listing is held is an archive there is nothing to ask about — and it is a probe of the cache
/// rather than a read, so a caller that asks it on every hover of every file pays a lock and a
/// comparison for each of them.
pub(crate) fn is_listed(path: &Path) -> bool {
    let Ok(metadata) = std::fs::metadata(path) else {
        return false;
    };
    let key = ListingKey::of(path, &metadata);

    LISTINGS
        .lock()
        .map(|cache| cache.iter().any(|(cached, _)| *cached == key))
        .unwrap_or(false)
}

/// One entry's bytes, found by the name a listing of the same file gave it.
///
/// It is what a reader of this app's own asks for when what it wants is not the archive but one
/// thing inside it: the first page of a comic is a picture in a zip or a rar, and a picture is what
/// the comic reader needs rather than the box it is filed in. What that costs is the entry and not
/// the archive, which is the whole of why a hover on a comic is a decode rather than an unpacking:
/// the headers are walked for the name alone — inflating nothing — and only the one entry that was
/// named is read.
///
/// The name is compared the way a listing writes it rather than as the format spells it, so a
/// caller never has to hand back a name the archive would not have recognized: the separators of a
/// Windows-written zip and the control characters of a name that has them are a listing's own doing
/// (see `normalize_name`), and a caller holding one of its names holds one this can find.
///
/// What is read is bounded by the ceiling one hover may decode for: an entry larger than that is
/// read as far as the ceiling and no further, and a picture that is not all there is a picture that
/// does not decode — which is the answer "no preview", arrived at without the allocation that would
/// have been refused anyway.
///
/// Only the two containers a comic is filed in are read this way. A `.cbz` is a zip and a `.cbr` a
/// rar; a 7z, a tar and a gzip stream are not containers a comic is published in, and a reader that
/// wants an entry out of one is a reader this cannot answer.
pub(crate) fn entry_bytes(path: &Path, name: &str) -> Option<Vec<u8>> {
    let limit = decode_budget_bytes();

    match kind_of(path)? {
        ArchiveKind::Zip => zip_entry_bytes(path, name, limit),
        ArchiveKind::Rar => rar_entry_bytes(path, name, limit),
        ArchiveKind::SevenZ | ArchiveKind::Tar | ArchiveKind::TarGz => None,
    }
}

/// One entry read out of a zip, by the name a listing gave it.
///
/// The entries are asked by index rather than by name because the name is a listing's: what is
/// compared is what `normalize_name` makes of each entry's own name, which is the same thing a
/// caller was handed.
fn zip_entry_bytes(path: &Path, name: &str, limit: u64) -> Option<Vec<u8>> {
    let mut archive = zip::ZipArchive::new(File::open(path).ok()?).ok()?;

    for index in 0..archive.len() {
        // The raw view, for the reason `read_zip` takes it: an entry packed by something this build
        // has no decompressor for still has a name, and one whose name matches is the entry to try
        // to read.
        let found = archive
            .by_index_raw(index)
            .map(|entry| normalize_name(entry.name()).as_deref() == Some(name))
            .unwrap_or(false);

        if !found {
            continue;
        }

        // Taken rather than read to the end: what the entry declares is a number a file may be
        // wrong about, so what bounds the read is the read itself.
        let entry = archive.by_index(index).ok()?;
        let mut bytes = Vec::new();
        entry.take(limit).read_to_end(&mut bytes).ok()?;

        return Some(bytes);
    }

    None
}

/// One entry read out of a rar, by the name a listing gave it.
///
/// A rar is walked rather than indexed: the archiver reads its way to the member it was asked for,
/// which is what makes a solid archive — where a member cannot be read without the ones before it —
/// readable at all. The size the member *declares* is what bounds the read here, because the
/// archiver hands back a member rather than a stream of one: a member larger than the ceiling is
/// answered with nothing rather than read, which is the one place this reader refuses a plate the
/// zip reader would have read part of.
fn rar_entry_bytes(path: &Path, name: &str, limit: u64) -> Option<Vec<u8>> {
    let mut archive = unrar::Archive::new(path).open_for_processing().ok()?;

    loop {
        let header = archive.read_header().ok()??;

        let (found, is_dir, declared) = {
            let entry = header.entry();

            (
                normalize_name(&entry.filename.to_string_lossy()).as_deref() == Some(name),
                entry.is_directory(),
                entry.unpacked_size,
            )
        };

        if found && !is_dir {
            if declared > limit {
                return None;
            }

            let (bytes, _) = header.read().ok()?;

            return Some(bytes);
        }

        archive = header.skip().ok()?;
    }
}

/// Hold the listing the engine produced for `path`, under the key a reader of this app's would
/// have been cached under.
///
/// It is what joins the engine's answer to the rest of the path: the measure pass and the render
/// pass that follow ask [`listing_for`] for the same file and find this rather than the readers
/// below, which have nothing to say about a format they do not have — and a second hover of the
/// file costs a lookup rather than a launch, for as long as the file is unchanged and the answer
/// has not been pushed out of the cache by other archives.
pub(crate) fn remember_engine_listing(path: &Path, listing: Listing) {
    let Ok(metadata) = std::fs::metadata(path) else {
        return;
    };
    let key = ListingKey::of(path, &metadata);

    let Ok(mut cache) = LISTINGS.lock() else {
        return;
    };

    if cache.len() >= LISTING_CACHE_MAX_ENTRIES {
        cache.clear();
    }

    cache.retain(|(cached, _)| *cached != key);
    cache.insert(0, (key, Arc::new(listing)));
}

/// The listing as the archive itself states it, read once per file version.
fn read_listing(path: &Path, file_size: u64, cancel: Option<&AtomicBool>) -> Option<Listing> {
    match kind_of(path)? {
        ArchiveKind::Zip => read_zip(path, file_size, cancel),
        ArchiveKind::SevenZ => read_sevenz(path, file_size, cancel),
        ArchiveKind::Rar => read_rar(path, file_size, cancel),
        ArchiveKind::Tar => read_tar(path, file_size, cancel),
        ArchiveKind::TarGz => read_targz(path, file_size, cancel),
    }
}

/// Which reader can open `path`.
///
/// The file's own bytes answer first, because they are the only thing that can:
/// a `.jar` is a zip under a name the zip crate has never heard of, and a `.zip`
/// that turns out to be a 7z under a stale name is still a 7z. The name answers
/// when the header says nothing — the two cases where it has to are the zip that
/// begins with an executable stub and the tar written without the ustar magic.
///
/// It is asked from outside the reader as well as inside it: what a container is, is the
/// question a route has to answer before it can say which reader opens one, and the answer
/// is borrowed from here rather than written down a second time (`native_formats`).
pub(crate) fn kind_of(path: &Path) -> Option<ArchiveKind> {
    if let Some(probe) = read_probe(path) {
        if let Some(kind) = magic_kind(&probe) {
            return Some(kind);
        }
    }

    extension_kind(path)
}

fn magic_kind(probe: &[u8]) -> Option<ArchiveKind> {
    if probe.starts_with(b"PK\x03\x04")
        || probe.starts_with(b"PK\x05\x06")
        || probe.starts_with(b"PK\x07\x08")
    {
        return Some(ArchiveKind::Zip);
    }
    if probe.starts_with(b"7z\xBC\xAF\x27\x1C") {
        return Some(ArchiveKind::SevenZ);
    }
    if probe.starts_with(b"Rar!\x1A\x07\x00") || probe.starts_with(b"Rar!\x1A\x07\x01\x00") {
        return Some(ArchiveKind::Rar);
    }
    if probe.starts_with(b"\x1F\x8B") {
        return Some(ArchiveKind::TarGz);
    }
    // A tar header carries its magic at byte 257; the rest of the block says what
    // the file is, so this is a tar and not something that happens to contain the
    // word.
    if probe.len() >= 263 && &probe[257..262] == b"ustar" {
        return Some(ArchiveKind::Tar);
    }

    None
}

fn extension_kind(path: &Path) -> Option<ArchiveKind> {
    let extension = path
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or("")
        .to_ascii_lowercase();

    match extension.as_str() {
        "zip" | "zipx" | "jar" | "apk" | "xpi" | "cbz" => Some(ArchiveKind::Zip),
        "7z" => Some(ArchiveKind::SevenZ),
        "rar" => Some(ArchiveKind::Rar),
        "tar" => Some(ArchiveKind::Tar),
        "tgz" => Some(ArchiveKind::TarGz),
        _ => {
            let name = path.file_name()?.to_str()?.to_ascii_lowercase();
            name.ends_with(".tar.gz").then_some(ArchiveKind::TarGz)
        }
    }
}

fn read_probe(path: &Path) -> Option<Vec<u8>> {
    let mut file = File::open(path).ok()?;
    let mut probe = vec![0u8; PROBE_BYTES];
    let mut filled = 0usize;

    while filled < probe.len() {
        match file.read(&mut probe[filled..]) {
            Ok(0) => break,
            Ok(read) => filled += read,
            Err(_) => return None,
        }
    }
    probe.truncate(filled);

    Some(probe)
}

fn read_zip(path: &Path, file_size: u64, cancel: Option<&AtomicBool>) -> Option<Listing> {
    let file = File::open(path).ok()?;
    let mut archive = zip::ZipArchive::new(file).ok()?;
    let mut listing = Listing::new();
    listing.file_size = file_size;

    let count = archive.len();
    for index in 0..count {
        if listing.at_cap(cancel) {
            listing.scan_capped = true;
            break;
        }

        // The raw view: a zip whose members were packed by something this build
        // has no decompressor for still lists, because a listing never inflates
        // anything.
        let Ok(entry) = archive.by_index_raw(index) else {
            continue;
        };

        let is_dir = entry.is_dir();
        let Some(name) = normalize_name(entry.name()) else {
            continue;
        };

        listing.push(
            name,
            entry.size(),
            Some(entry.compressed_size()),
            is_dir,
            entry.encrypted(),
        );
    }

    Some(listing)
}

fn read_sevenz(path: &Path, file_size: u64, cancel: Option<&AtomicBool>) -> Option<Listing> {
    let archive = match sevenz_rust2::Archive::open(path) {
        Ok(archive) => archive,
        Err(sevenz_rust2::Error::PasswordRequired) => {
            return Some(Listing::encrypted_headers(file_size))
        }
        Err(_) => return None,
    };

    let mut listing = Listing::new();
    listing.file_size = file_size;
    // 7z encrypts members per block with AES, and states it nowhere per entry.
    let encrypted = archive.blocks.iter().any(|block| {
        block
            .coders
            .iter()
            .any(|coder| coder.encoder_method_id() == SEVENZ_AES_METHOD)
    });

    for entry in &archive.files {
        if listing.at_cap(cancel) {
            listing.scan_capped = true;
            break;
        }
        // An anti-item is a deletion an update archive carries, not a member.
        if entry.is_anti_item {
            continue;
        }
        let Some(name) = normalize_name(entry.name()) else {
            continue;
        };

        let is_dir = entry.is_directory();
        listing.push(
            name,
            if is_dir { 0 } else { entry.size() },
            None,
            is_dir,
            encrypted,
        );
    }

    Some(listing)
}

fn read_rar(path: &Path, file_size: u64, cancel: Option<&AtomicBool>) -> Option<Listing> {
    // Listing walks the archive's headers and skips every member's data, so the
    // cost is the headers whether the archive holds a text file or a film.
    let archive = unrar::Archive::new(path).open_for_listing().ok()?;
    let mut listing = Listing::new();
    listing.file_size = file_size;

    for header in archive {
        if listing.at_cap(cancel) {
            listing.scan_capped = true;
            break;
        }

        let header = match header {
            Ok(header) => header,
            Err(_) => {
                // Nothing read at all is not a listing with a caveat, it is a
                // file this reader cannot answer for.
                if listing.entries.is_empty() {
                    return None;
                }
                listing.read_truncated = true;
                break;
            }
        };

        let Some(name) = normalize_name(&header.filename.to_string_lossy()) else {
            continue;
        };

        listing.multi_volume |= header.is_split();
        let is_dir = header.is_directory();
        listing.push(
            name,
            if is_dir { 0 } else { header.unpacked_size },
            None,
            is_dir,
            header.is_encrypted(),
        );
    }

    Some(listing)
}

fn read_tar(path: &Path, file_size: u64, cancel: Option<&AtomicBool>) -> Option<Listing> {
    let file = File::open(path).ok()?;
    // A plain tar is seekable, so walking past a member costs a seek rather than
    // a read of everything before the next header.
    collect_tar(
        tar::Archive::new(file).entries_with_seek(),
        file_size,
        cancel,
    )
}

fn read_targz(path: &Path, file_size: u64, cancel: Option<&AtomicBool>) -> Option<Listing> {
    let file = File::open(path).ok()?;
    let decoder = flate2::read::GzDecoder::new(file);
    let bounded = decoder.take(MAX_INFLATED_BYTES);
    collect_tar(tar::Archive::new(bounded).entries(), file_size, cancel)
}

fn collect_tar<R: Read>(
    entries: std::io::Result<tar::Entries<'_, R>>,
    file_size: u64,
    cancel: Option<&AtomicBool>,
) -> Option<Listing> {
    let mut listing = Listing::new();
    listing.file_size = file_size;

    let Ok(entries) = entries else {
        return None;
    };

    for entry in entries {
        if listing.at_cap(cancel) {
            listing.scan_capped = true;
            break;
        }

        let entry = match entry {
            Ok(entry) => entry,
            Err(_) => {
                if listing.entries.is_empty() {
                    return None;
                }
                listing.read_truncated = true;
                break;
            }
        };

        let is_dir = entry.header().entry_type().is_dir();
        let Some(name) = entry
            .path()
            .ok()
            .and_then(|path| normalize_name(&path.to_string_lossy()))
        else {
            continue;
        };

        listing.push(
            name,
            entry.header().size().unwrap_or(0),
            None,
            is_dir,
            false,
        );
    }

    Some(listing)
}

/// The AES coder's method id in a 7z block, which is how a 7z states that its
/// members are encrypted.
const SEVENZ_AES_METHOD: [u8; 4] = [0x06, 0xF1, 0x07, 0x01];

/// An entry's name as it is shown: `/`-separated, one line, and not absurdly
/// long.
///
/// A name is the one thing an archive can say that the rest of the app cannot
/// use, so nothing here decides what a name means — it is cleaned up to be drawn,
/// and that is all. Control characters go because a newline in a name would draw
/// a second row, backslashes become separators because a zip written by Windows
/// keeps them, and empty names are not rows.
fn normalize_name(raw: &str) -> Option<String> {
    let mut name = String::with_capacity(raw.len().min(MAX_NAME_CHARS));
    for character in raw.chars().take(MAX_NAME_CHARS) {
        match character {
            '\0' => {}
            '\\' => name.push('/'),
            character if character.is_control() => {}
            character => name.push(character),
        }
    }

    let name = name.trim_matches('/');
    (!name.is_empty()).then(|| name.to_string())
}

// ---------------------------------------------------------------------- cache

#[derive(PartialEq, Eq, Hash)]
struct ListingKey {
    path: PathBuf,
    modified: Option<SystemTime>,
    len: u64,
}

impl ListingKey {
    /// What a listing of this file is held under: the file, and the version of it that was read,
    /// so that an archive saved again is an archive to read again.
    fn of(path: &Path, metadata: &std::fs::Metadata) -> Self {
        Self {
            path: path.to_path_buf(),
            modified: metadata.modified().ok(),
            len: metadata.len(),
        }
    }
}

/// Listings, newest first. Small: it is here so that the measure pass, the render
/// pass and a second hover read the archive once between them.
type ListingCache = Vec<(ListingKey, Arc<Listing>)>;

static LISTINGS: Lazy<Mutex<ListingCache>> = Lazy::new(|| Mutex::new(Vec::new()));

#[cfg(test)]
mod tests;

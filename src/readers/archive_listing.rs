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

use once_cell::sync::Lazy;
use std::fs::File;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::SystemTime;

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

// ------------------------------------------------------------ the engine's answer

/// The line the engine writes between what it says about the archive and the entries themselves.
///
/// Everything above it is the engine talking about its own run — its version, the file it
/// scanned, the archive's own header — and everything below it is the archive's table of
/// contents in the same `Key = Value` shape, one blank line between entries. Splitting there is
/// what keeps the archive's own `Path` out of the listing, since the header of the report states
/// it too.
const ENGINE_SEPARATOR: &str = "----------";

/// The listing the PeaZip engine's own report of an archive gives, read into the shape every
/// reader of this app's produces.
///
/// What the engine is asked for is a technical listing — each entry's own fields, one to a line —
/// and what it answers with is that listing as text, which is the only thing about it this side
/// has to be able to read. The fields are read for what the page needs and nothing else: the
/// entry's path, what it weighs and what it took, whether it is a folder, and whether it is
/// encrypted. Everything the engine says beyond those — the method an entry was packed with, the
/// block it shares, its CRC, its timestamps, its attributes — belongs to an extraction this app
/// never does.
///
/// Two shapes of report are worth naming, because both are ordinary and neither is an error. A
/// report with nothing under the separator is an archive the engine could not open the table of
/// contents of — one whose own headers are encrypted, which it will not list without a password —
/// and it is answered with no listing rather than with an empty one: a page saying an archive
/// holds nothing is a claim, and this is not the place that claim can be made from. And a
/// single-stream format keeps no name inside itself at all — a `bzip2` or a `zstd` file is the
/// compressed bytes and nothing beside them — so the one entry the engine reports for one of
/// those has no path, and what it is called is worked out from the archive's own name (see
/// [`stream_member_name`]).
pub(crate) fn engine_listing(report: &str, archive: &Path, file_size: u64) -> Option<Listing> {
    let (_, entries) = report.split_once(ENGINE_SEPARATOR)?;

    let mut listing = Listing::new();
    listing.file_size = file_size;

    let mut fields: Vec<(&str, &str)> = Vec::new();

    // The report ends with a blank line, and a block is ended by one — so the walk is given one
    // more line than the report has and the last block is closed by it rather than by the end of
    // the text.
    for line in entries.lines().chain(std::iter::once("")) {
        if line.trim().is_empty() {
            if !fields.is_empty() {
                push_engine_entry(&mut listing, &fields, archive);
                fields.clear();
            }
            continue;
        }

        // A field is `Name = Value`, and the value may be empty — which is how the engine says
        // that a field does not apply to this entry. The split is on the first separator, so a
        // path with one of its own is still a path.
        if let Some((name, value)) = line.split_once(" = ") {
            fields.push((name.trim(), value.trim()));
        }
    }

    // An archive with nothing in it and one whose table could not be read are the same report
    // from here, and both are answered the same way (see above).
    if listing.entries.is_empty() {
        return None;
    }

    listing.settle_totals();
    Some(listing)
}

/// One entry out of the fields the engine reported for it.
fn push_engine_entry(listing: &mut Listing, fields: &[(&str, &str)], archive: &Path) {
    let value = |name: &str| {
        fields
            .iter()
            .find(|(field, _)| *field == name)
            .map(|(_, value)| *value)
    };

    let Some(name) = value("Path").map(str::to_string).or_else(|| {
        // An entry the engine reports no path for is the member of a single-stream format, which
        // has no name inside it to report.
        value("Size")?;
        stream_member_name(archive)
    }) else {
        return;
    };

    // A folder is stated twice over by the two shapes of report: the entry's own attributes
    // carry the `D` every extractor reads, and the containers that keep a separate flag state it
    // as one of the entry's fields.
    let is_dir = value("Attributes").is_some_and(|attributes| attributes.contains('D'))
        || value("Folder") == Some("+");

    let size = value("Size")
        .and_then(|size| size.parse::<u64>().ok())
        .unwrap_or(0);
    // A field left empty is a field the engine has no answer for — a member of a solid block
    // shares its packed size with the entries beside it — and a total the page cannot state is
    // what `settle_totals` does with one.
    let packed = value("Packed Size").and_then(|packed| packed.parse::<u64>().ok());
    let encrypted = value("Encrypted") == Some("+");

    if let Some(name) = normalize_name(&name) {
        listing.push(
            name,
            if is_dir { 0 } else { size },
            packed,
            is_dir,
            encrypted,
        );
    }
}

/// What the one member of a single-stream format is called: the archive's own name with the
/// extension it was compressed under taken off, which is the name an extraction writes it under
/// and the name the file had before it was compressed.
///
/// Nothing here can recover the name it *did* have — a `bzip2` stream keeps no name, and a `zstd`
/// one usually keeps none either — so what is shown is the closest true thing rather than a guess
/// at the original.
fn stream_member_name(archive: &Path) -> Option<String> {
    let name = archive.file_name().and_then(|name| name.to_str())?;

    let stem = name
        .rsplit_once('.')
        .map(|(stem, _)| stem)
        .filter(|stem| !stem.is_empty())
        .unwrap_or(name);

    Some(stem.to_string())
}

/// Which reader can open `path`.
///
/// The file's own bytes answer first, because they are the only thing that can:
/// a `.jar` is a zip under a name the zip crate has never heard of, and a `.zip`
/// that turns out to be a 7z under a stale name is still a 7z. The name answers
/// when the header says nothing — the two cases where it has to are the zip that
/// begins with an executable stub and the tar written without the ustar magic.
fn kind_of(path: &Path) -> Option<ArchiveKind> {
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
mod tests {
    use super::*;

    /// A report of a zip, as the engine writes one: every entry states whether it is a folder in
    /// a field of its own as well as in its attributes, and the sizes are the entry's own.
    const ZIP_REPORT: &str = "\n7-Zip 25.01 (x64) : Copyright (c) 1999-2025 Igor Pavlov : 2025-08-03\n\n\
Scanning the drive for archives:\n1 file, 428 bytes (1 KiB)\n\n\
Listing archive: C:\\downloads\\photos.zip\n\n--\nPath = C:\\downloads\\photos.zip\nType = zip\n\
Physical Size = 428\n\n----------\nPath = a.txt\nFolder = -\nSize = 13\nPacked Size = 13\n\
Modified = 2026-09-25 03:25:39.6103164\nAttributes = A\nEncrypted = -\nCRC = 38E6C41A\nMethod = Store\n\n\
Path = inner\nFolder = +\nSize = 0\nPacked Size = 0\nModified = 2026-09-25 03:25:39.6118192\n\
Attributes = D\nEncrypted = -\nCRC = \nMethod = Store\n\n\
Path = inner\\b.txt\nFolder = -\nSize = 13\nPacked Size = 13\nAttributes = A\nEncrypted = -\n\n";

    #[test]
    fn reads_the_entries_out_of_the_engines_own_report() {
        let listing = engine_listing(ZIP_REPORT, Path::new(r"C:\downloads\photos.zip"), 428)
            .expect("a listing");

        assert_eq!(listing.entries.len(), 3);
        assert_eq!(listing.file_size, 428);

        let names: Vec<&str> = listing
            .entries
            .iter()
            .map(|entry| entry.name.as_str())
            .collect();
        // A path inside the archive is `/`-separated wherever the engine wrote a backslash, the
        // way every other reader of this app's names one.
        assert_eq!(names, vec!["a.txt", "inner", "inner/b.txt"]);

        assert!(listing.entries[1].is_dir);
        assert!(!listing.entries[0].is_dir);
        assert_eq!(listing.entries[2].size, 13);
        assert_eq!(listing.entries[2].packed, Some(13));
        assert!(!listing.entries[0].encrypted);

        // The folder takes nothing in either column, which is what the totals are settled over.
        assert_eq!(listing.total_size, 26);
        assert_eq!(listing.packed_total, Some(26));
        assert!(!listing.scan_capped && !listing.read_truncated && !listing.encrypted_headers);
    }

    /// A report of a 7z, which is the other shape: a folder is stated by its attributes alone,
    /// and a member of a solid block has no packed size of its own to state.
    #[test]
    fn reads_a_report_whose_entries_state_less_about_themselves() {
        let report =
            "Path = C:\\downloads\\backup.7z\nType = 7z\nPhysical Size = 203\n\n----------\n\
Path = inner\nSize = 0\nPacked Size = 0\nAttributes = D\nEncrypted = -\n\n\
Path = a.txt\nSize = 13\nPacked Size = 30\nAttributes = A\nEncrypted = -\n\n\
Path = inner\\b.txt\nSize = 13\nPacked Size = \nAttributes = A\nEncrypted = -\n\n";

        let listing =
            engine_listing(report, Path::new(r"C:\downloads\backup.7z"), 203).expect("a listing");

        assert_eq!(listing.entries.len(), 3);
        assert!(listing.entries[0].is_dir);
        assert_eq!(listing.entries[1].packed, Some(30));
        assert_eq!(
            listing.entries[2].packed, None,
            "a member of a solid block has no packed size of its own"
        );

        // One entry the format cannot state a packed size for is a total this page cannot state
        // either, which is what every other reader's `None` means in the same place.
        assert_eq!(listing.total_size, 26);
        assert_eq!(listing.packed_total, None);
    }

    /// A single-stream format keeps no name inside itself, so the one member the engine reports
    /// for one has no path: what it is called is the archive's own name without the extension it
    /// was compressed under, which is where an extraction would put it.
    #[test]
    fn names_the_one_member_of_a_stream_that_keeps_no_name_of_its_own() {
        let report = "Path = C:\\downloads\\notes.txt.zst\nType = zstd\n\n----------\n\
Size = \nPacked Size = \n\n";

        let listing = engine_listing(report, Path::new(r"C:\downloads\notes.txt.zst"), 26)
            .expect("a listing");

        assert_eq!(listing.entries.len(), 1);
        assert_eq!(listing.entries[0].name, "notes.txt");
        assert_eq!(listing.entries[0].size, 0);
        assert_eq!(listing.entries[0].packed, None);
        assert!(!listing.entries[0].is_dir);
    }

    /// What an encrypted archive reports is a report with nothing under the separator, and what
    /// a file that is not an archive at all reports is no separator — both are answered with no
    /// listing rather than with a page claiming the archive is empty.
    #[test]
    fn answers_an_archive_it_could_not_read_with_no_listing() {
        let encrypted =
            "Path = C:\\downloads\\secret.7z\nType = 7z\nPhysical Size = 32\n\n----------\n";
        assert!(engine_listing(encrypted, Path::new("secret.7z"), 32).is_none());

        let not_an_archive = "7-Zip 25.01 (x64)\n\nERROR: notes.txt\nnotes.txt\nOpen ERROR: Can not open the file as archive\n";
        assert!(engine_listing(not_an_archive, Path::new("notes.txt"), 12).is_none());

        assert!(engine_listing("", Path::new("empty.cab"), 0).is_none());
    }

    /// What is held for a file is what the two passes of one hover read, and an archive saved
    /// again is read again rather than answered out of the cache.
    #[test]
    fn holds_the_engines_answer_under_the_files_own_key() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-peazip-listing")
            .join("remembering");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let archive = folder.join("a.cab");
        std::fs::write(&archive, b"MSCF\x00\x00\x00\x00 a cabinet, of a sort")
            .expect("a written file");

        let listing = |name: &str| {
            let mut listing = Listing::new();
            listing.file_size = 32;
            listing.push(name.to_string(), 13, Some(13), false, false);
            listing.settle_totals();
            listing
        };

        remember_engine_listing(&archive, listing("first.txt"));
        let held = listing_for(&archive, None).expect("the answer that is held");
        assert_eq!(held.entries[0].name, "first.txt");
        assert!(
            !held.encrypted_headers,
            "what the engine answered is a listing, not a reader's caveat"
        );

        // Saved again: what was known about the file it was is not what it is now, so the key
        // does not match and the answer is not the one that was held.
        std::fs::write(&archive, b"MSCF\x00\x00\x00\x00 a cabinet, saved again")
            .expect("a written file");
        assert!(
            listing_for(&archive, None).is_none(),
            "no reader here knows the format, and the engine has not answered for this version"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }
}

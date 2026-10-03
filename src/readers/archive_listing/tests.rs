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
    let listing =
        engine_listing(ZIP_REPORT, Path::new(r"C:\downloads\photos.zip"), 428).expect("a listing");

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
    let report = "Path = C:\\downloads\\backup.7z\nType = 7z\nPhysical Size = 203\n\n----------\n\
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

    let listing =
        engine_listing(report, Path::new(r"C:\downloads\notes.txt.zst"), 26).expect("a listing");

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

/// FreeArc's verbose listing, as FreeArc 0.67 — the archiver PeaZip 10.9.0 carries — writes it:
/// a timestamped line per entry, with the attributes a folder is told by, the size, the space
/// the entry takes and its CRC, and the name last, backslashes and all. Captured from the tool.
const ARC_REPORT: &str =
    "FreeArc 0.67 (March 15 2014) listing archive: C:\\downloads\\backup.arc\n\
Date/time              Attr            Size          Packed      CRC Filename\n\
-----------------------------------------------------------------------------\n\
2026-09-25 12:15:27 .D.....               0               0 00000000 inner\n\
2026-09-25 12:15:27 .......              30              77 86e5f392 notes.txt\n\
2026-09-25 12:15:27 .......              12               0 c3dbba13 inner\\deep.txt\n\
-----------------------------------------------------------------------------\n\
3 files, 42 bytes, 77 compressed\n\
All OK\n";

#[test]
fn reads_the_entries_out_of_freearcs_verbose_listing() {
    let listing = arc_listing(ARC_REPORT, 367).expect("a listing");

    assert_eq!(listing.entries.len(), 3);
    assert_eq!(listing.file_size, 367);

    let names: Vec<&str> = listing
        .entries
        .iter()
        .map(|entry| entry.name.as_str())
        .collect();
    assert_eq!(names, vec!["inner", "notes.txt", "inner/deep.txt"]);

    // The attributes are what a folder is told by here: it is the only column that says so,
    // and the rule lines and the totals the tool closes with are not entries at all.
    assert!(listing.entries[0].is_dir);
    assert!(!listing.entries[1].is_dir);
    assert_eq!(listing.entries[1].size, 30);
    assert_eq!(listing.total_size, 42);

    // The packed column is block accounting rather than a size per entry — the second file of
    // a solid block is reported as taking nothing, the block having been paid for by the first
    // — so neither a size nor a total is stated from it.
    assert_eq!(listing.entries[1].packed, None);
    assert_eq!(listing.packed_total, None);
}

/// And a file that is not one of its archives is a line of its own with no entries under it,
/// which is no listing here rather than an empty one.
#[test]
fn answers_an_arc_the_archiver_could_not_read_with_no_listing() {
    let report = "FreeArc 0.67 (March 15 2014) listing archive: C:\\downloads\\notes.txt\n\n\
ERROR: C:\\downloads\\notes.txt isn't archive or this archive is corrupt: archive signature not found at the end of archive.\n";

    assert!(arc_listing(report, 30).is_none());
}

/// What zpaq prints for an archive, as the zpaq PeaZip 10.9.0 carries — zpaqfranz — writes it:
/// the latest version of every name it holds, one line each, with a folder marked by a bracket
/// before its size and by a `#`, and a file by a `+`. Captured from the tool.
const ZPAQ_REPORT: &str =
    "zpaqfranz v62.5h-JIT,SFTP-L,HW BLAKE3,SHA1/2,4,SFX64 v55.1,(2025-07-29)\n\n\
<<C:/downloads/backup.zpaq>>: 1 versions, 3 files, 1.254 bytes (1.22 KB)\n\n\n\
   Date      Time   Size Ratio Name\n\
---------- --------  ---- ----- -----\n\
2026-09-25 12:15:27 [  12   dir # inner/\n\
2026-09-25 12:15:27    12 >999% + inner/deep.txt\n\
2026-09-25 12:15:27    30 >999% + notes.txt\n\n\
                42 (42.00  B) of 42 (42.00  B) in 3 files shown\n\
             1.254 compressed  Ratio 29.857 <<C:/downloads/backup.zpaq>>\n\
0.000s (00:00:00,13.07KB) (all OK)\n";

#[test]
fn reads_the_names_out_of_zpaqs_own_listing() {
    let listing = zpaq_listing(ZPAQ_REPORT, 1254).expect("a listing");

    assert_eq!(listing.entries.len(), 3);

    let names: Vec<&str> = listing
        .entries
        .iter()
        .map(|entry| entry.name.as_str())
        .collect();
    // The folder is stored with the slash that marks it and is named without it, the way every
    // other reader of this app's names one.
    assert_eq!(names, vec!["inner", "inner/deep.txt", "notes.txt"]);

    assert!(listing.entries[0].is_dir);
    assert_eq!(listing.entries[2].size, 30);
    assert_eq!(listing.total_size, 42);

    // zpaq states what an archive weighs, not what one name in it took, so the page gives the
    // file's own weight rather than a saving it cannot state.
    assert_eq!(listing.packed_total, None);
}

/// And what zpaq prints for a file that is not its own is its usage screen with no entry lines
/// in it, which is the same answer: no listing.
#[test]
fn answers_a_zpaq_it_could_not_read_with_no_listing() {
    let usage = "zpaqfranz v62.5h-JIT,SFTP-L,HW BLAKE3,SHA1/2,4,SFX64 v55.1,(2025-07-29)\n\
Usage: zpaqfranz command archive.zpaq files|directories -switches\n\
  a: Append files     | x: Extract            |   t: Test\n";

    assert!(zpaq_listing(usage, 30).is_none());
    assert!(zpaq_listing("", 30).is_none());
}

/// What `zstd -l` prints for a stream, captured from the tool: a header and one line carrying
/// how many frames the file holds, what it weighs and what it weighed before it was compressed.
const ZSTD_REPORT: &str = "Frames  Skips  Compressed  Uncompressed  Ratio  Check  Filename\n\
 1      0      43   B        30   B  0.698  XXH64  C:\\downloads\\notes.txt.zst\n";

#[test]
fn reads_the_size_the_zstd_tool_reports_for_a_stream() {
    let listing =
        zstd_listing(ZSTD_REPORT, Path::new(r"C:\downloads\notes.txt.zst"), 43).expect("a listing");

    assert_eq!(listing.entries.len(), 1);
    assert_eq!(listing.entries[0].name, "notes.txt");
    assert_eq!(listing.entries[0].size, 30);
    assert_eq!(listing.entries[0].packed, Some(43));
    assert_eq!(listing.total_size, 30);
    assert_eq!(listing.packed_total, Some(43));
}

/// The sizes in that table are a number and a unit, and a unit larger than bytes is a size
/// this side reads rather than one it mistakes for a number of bytes.
#[test]
fn reads_the_units_the_zstd_table_writes_its_sizes_in() {
    let report = "Frames  Skips  Compressed  Uncompressed  Ratio  Check  Filename\n\
 1      0     168   B      3.50 KiB  21.333  XXH64  C:\\downloads\\sample.tar.zst\n";

    let listing =
        zstd_listing(report, Path::new(r"C:\downloads\sample.tar.zst"), 168).expect("a listing");

    assert_eq!(listing.entries[0].size, 3584, "3.50 KiB is 3584 bytes");
    assert_eq!(listing.entries[0].packed, Some(168));
}

/// A file the tool will not read is answered with the header and a line saying so, which is no
/// data line and so no listing.
#[test]
fn answers_a_stream_the_zstd_tool_will_not_read_with_no_listing() {
    let report = "Frames  Skips  Compressed  Uncompressed  Ratio  Check  Filename\n\
File \"C:\\downloads\\notes.txt\" not compressed by zstd \n";

    assert!(zstd_listing(report, Path::new(r"C:\downloads\notes.txt.zst"), 30).is_none());
}

/// The one answer here that no tool produced: a single-stream file whose own tool has no
/// listing to give is the member an extraction would write — named after the file, and with no
/// size of its own, since nothing but a decompression of the whole file knows one.
#[test]
fn derives_the_member_of_a_stream_no_tool_can_report_on() {
    let listing = stream_listing(Path::new(r"C:\downloads\notes.txt.br"), 31).expect("a listing");

    assert_eq!(listing.entries.len(), 1);
    assert_eq!(listing.entries[0].name, "notes.txt");
    assert_eq!(listing.entries[0].size, 0);
    assert_eq!(listing.entries[0].packed, None);
    assert!(!listing.entries[0].is_dir);
    assert_eq!(listing.file_size, 31);
    assert_eq!(listing.total_size, 0);
    assert_eq!(listing.packed_total, None);
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
    std::fs::write(&archive, b"MSCF\x00\x00\x00\x00 a cabinet, of a sort").expect("a written file");

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

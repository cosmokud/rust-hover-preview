use super::*;
use crate::readers::archive_listing::listing_for;
use std::fs;
use std::io::Write;

/// A row says which of the entries behind it cannot be read without a password, and a
/// folder row does not: the container's own headers are encrypted or they are not, which
/// is the listing's line rather than a row of the tree.
#[test]
fn a_row_says_when_its_entry_needs_a_password() {
    use crate::readers::archive_listing::ArchiveEntry;

    let entry = |name: &str, encrypted: bool| ArchiveEntry {
        name: name.to_string(),
        size: 8,
        packed: None,
        is_dir: false,
        encrypted,
    };

    let tree = Tree::build(&[
        entry("locked.txt", true),
        entry("open.txt", false),
        entry("folder/locked.txt", true),
    ]);
    let rows = tree.rows();

    let row = |name: &str| {
        rows.iter()
            .find(|row| row.name == name)
            .unwrap_or_else(|| panic!("a row for `{name}`"))
    };

    assert!(row("locked.txt").encrypted, "a locked entry says so");
    assert!(!row("open.txt").encrypted, "an ordinary one does not");
    assert!(
        !row("folder").encrypted,
        "and the folder a locked entry sits in is not locked itself"
    );
}

/// Where one test's fixtures and the pictures of them are written. The
/// scratchpad the session hands out, so a render can be looked at rather
/// than only asserted; the label keeps tests that build the same shapes from
/// writing over each other's files.
fn scratch(label: &str) -> std::path::PathBuf {
    let root = std::env::var_os("COMMANDCODE_SCRATCHPAD")
        .map(std::path::PathBuf::from)
        .unwrap_or_else(std::env::temp_dir)
        .join("archive-preview")
        .join(label);
    fs::create_dir_all(&root).expect("a fixture directory");
    root
}

fn write_zip(path: &Path, entries: &[(&str, usize)]) {
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let file = fs::File::create(path).expect("a zip to write");
    let mut writer = zip::ZipWriter::new(file);

    for (name, size) in entries {
        if name.ends_with('/') {
            writer.add_directory(*name, options).expect("a directory");
        } else {
            writer.start_file(*name, options).expect("a member");
            writer.write_all(&vec![b'x'; *size]).expect("bytes");
        }
    }

    writer.finish().expect("a finished zip");
}

fn write_tar(path: &Path, entries: &[(&str, usize)], gzip: bool) {
    let file = fs::File::create(path).expect("a tar to write");
    if gzip {
        let encoder = flate2::write::GzEncoder::new(file, flate2::Compression::default());
        let mut builder = tar::Builder::new(encoder);
        append_tar(&mut builder, entries);
        builder
            .into_inner()
            .expect("the encoder")
            .finish()
            .expect("gz");
    } else {
        let mut builder = tar::Builder::new(file);
        append_tar(&mut builder, entries);
        builder.into_inner().expect("the file");
    }
}

fn append_tar<W: Write>(builder: &mut tar::Builder<W>, entries: &[(&str, usize)]) {
    for (name, size) in entries {
        let bytes = vec![b'y'; *size];
        let mut header = tar::Header::new_gnu();
        header.set_size(bytes.len() as u64);
        header.set_mode(0o644);
        header.set_cksum();
        builder
            .append_data(&mut header, name, bytes.as_slice())
            .expect("a tar member");
    }
}

/// A zip with the shapes the tree has to survive: implicit folders, an
/// explicit one, a chain of single-child folders, numbers to order, a name
/// from outside ASCII, and a long name to cut.
fn bundle(dir: &Path) -> std::path::PathBuf {
    let path = dir.join("bundle.zip");
    write_zip(
        &path,
        &[
            ("docs/", 0),
            ("docs/report.pdf", 1_200_000),
            ("docs/notes.txt", 4_096),
            ("docs/deep/a/b/c/only.txt", 64),
            ("images/logo.png", 48_000),
            ("images/banner.jpg", 310_000),
            ("src/main.rs", 2_048),
            ("src/lib.rs", 1_024),
            ("report2.txt", 10),
            ("report10.txt", 20),
            ("музыка/трек.mp3", 900),
            (
                "a-very-long-file-name-that-has-to-be-cut-when-the-box-is-narrow.tar.gz",
                1_536,
            ),
            ("README.md", 2_048),
        ],
    );

    path
}

fn render_fixture(dir: &Path, name: &str, archive: &Path, theme: TextTheme, cap: u32) {
    let options = ArchivePreviewOptions {
        theme,
        font_scale_percent: 125,
    };
    let (width, height) = measure(archive, cap, 1_400, 96, options).expect("a measured page");
    let (pixels, width, height) = render(archive, width, height, 96, options).expect("a page");

    let mut rgba = pixels;
    for pixel in rgba.as_chunks_mut::<4>().0 {
        pixel.swap(0, 2);
    }
    image::save_buffer(
        dir.join(format!("{name}.png")),
        &rgba,
        width,
        height,
        image::ExtendedColorType::Rgba8,
    )
    .expect("a written picture");
}

#[test]
fn reads_the_shapes_a_zip_holds() {
    let dir = scratch("shapes");
    let path = bundle(&dir);
    let listing = listing_for(&path, None).expect("a listing");

    let names: Vec<&str> = listing
        .entries
        .iter()
        .map(|entry| entry.name.as_str())
        .collect();
    assert!(names.contains(&"docs/report.pdf"));
    assert!(names.contains(&"docs/deep/a/b/c/only.txt"));
    assert!(names.contains(&"музыка/трек.mp3"));
    assert!(!listing
        .entries
        .iter()
        .any(|entry| entry.name.ends_with('/')));

    let explicit = listing
        .entries
        .iter()
        .find(|entry| entry.name == "docs")
        .expect("the explicit folder");
    assert!(explicit.is_dir);

    assert_eq!(
        listing.total_size,
        1_200_000 + 4_096 + 64 + 48_000 + 310_000 + 2_048 + 1_024 + 10 + 20 + 900 + 1_536 + 2_048
    );
    assert!(listing.packed_total.is_some());
    assert!(!listing.scan_capped);
}

#[test]
fn reads_tar_with_and_without_a_gzip_around_it() {
    let dir = scratch("tar");
    let entries: [(&str, usize); 3] = [("one.txt", 128), ("two/deep.txt", 256), ("three.bin", 512)];

    for (name, gzip) in [("sample.tar", false), ("sample.tar.gz", true)] {
        let path = dir.join(name);
        write_tar(&path, &entries, gzip);
        let listing = listing_for(&path, None).expect("a listing");

        assert_eq!(listing.entries.len(), 3, "{name}");
        assert_eq!(listing.total_size, 128 + 256 + 512, "{name}");
        // A tar states no packed size per member.
        assert_eq!(listing.packed_total, None, "{name}");

        render_fixture(&dir, name, &path, TextTheme::Light, 1_920);
    }
}

#[test]
fn reads_a_sevenz() {
    let dir = scratch("sevenz");
    let source = dir.join("sevenz-source");
    fs::create_dir_all(source.join("nested")).expect("a source tree");
    fs::write(source.join("hello.txt"), b"hello world").expect("a file");
    fs::write(source.join("nested/data.bin"), vec![0u8; 4_096]).expect("a file");

    let path = dir.join("sample.7z");
    sevenz_rust2::compress_to_path(&source, &path).expect("a written 7z");

    let listing = listing_for(&path, None).expect("a listing");
    let names: Vec<&str> = listing
        .entries
        .iter()
        .map(|entry| entry.name.as_str())
        .collect();
    assert!(
        names.iter().any(|name| name.ends_with("hello.txt")),
        "{names:?}"
    );
    assert!(listing.total_size >= 4_096);

    render_fixture(&dir, "sample-7z", &path, TextTheme::Light, 1_920);
}

#[test]
fn refuses_a_file_that_is_not_an_archive() {
    let dir = scratch("garbage");
    let path = dir.join("garbage.rar");
    fs::write(&path, b"this is not a rar file, whatever it is called").expect("a file");

    assert!(listing_for(&path, None).is_none());
}

/// The whole point of the reader: an archive whose members are packed with
/// something this build cannot unpack still lists, because a listing reads
/// the table and not the members. The fixture's two method fields are
/// rewritten to a method no one implements after the zip is written.
#[test]
fn lists_an_archive_it_could_not_unpack() {
    use std::io::{Read, Seek, SeekFrom};

    let dir = scratch("method");
    let path = dir.join("imploded.zip");
    write_zip(&path, &[("packed.bin", 4_096), ("plain.txt", 32)]);

    let mut file = fs::OpenOptions::new()
        .read(true)
        .write(true)
        .open(&path)
        .expect("the fixture");
    let length = file.metadata().expect("its size").len();
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes).expect("its bytes");
    assert_eq!(bytes.len() as u64, length);

    let mut patched = 0usize;
    // The local header states its method at byte 8 of the record and the
    // central directory at byte 10; the end-of-directory record is left
    // alone, because its fields mean something else.
    for (signature, offset) in [(b"PK\x03\x04", 8usize), (b"PK\x01\x02", 10usize)] {
        let mut at = 0usize;
        while at + 4 <= bytes.len() {
            if &bytes[at..at + 4] == signature {
                // implode: a method the zip crate has no reader for
                bytes[at + offset] = 6;
                bytes[at + offset + 1] = 0;
                patched += 1;
            }
            at += 1;
        }
    }
    assert!(patched >= 2, "the fixture states its methods somewhere");

    file.seek(SeekFrom::Start(0)).expect("the start");
    file.write_all(&bytes).expect("the patched archive");
    file.set_len(bytes.len() as u64).expect("its length");

    let listing = listing_for(&path, None).expect("a listing anyway");
    assert_eq!(listing.entries.len(), 2);
    assert!(listing
        .entries
        .iter()
        .any(|entry| entry.name == "packed.bin" && entry.size == 4_096));
}

#[test]
fn stops_reading_when_the_hover_moves_on() {
    use std::sync::atomic::AtomicBool;

    let dir = scratch("cancel");
    let path = dir.join("cancelled.zip");
    write_zip(&path, &[("one.txt", 16), ("two.txt", 16)]);

    let cancelled = AtomicBool::new(true);
    let listing = listing_for(&path, Some(&cancelled)).expect("a listing");
    assert!(listing.entries.is_empty());
    assert!(listing.scan_capped);
}

#[test]
fn reads_an_empty_zip() {
    let dir = scratch("empty");
    let path = dir.join("empty.zip");
    write_zip(&path, &[]);

    let listing = listing_for(&path, None).expect("a listing");
    assert!(listing.entries.is_empty());
    assert_eq!(listing.total_size, 0);
}

#[test]
fn stops_scanning_a_zip_that_holds_enough_entries() {
    let dir = scratch("many");
    let path = dir.join("many.zip");
    let entries: Vec<(String, usize)> = (0..21_000)
        .map(|index| (format!("dir{}/entry{index}.txt", index % 50), 0))
        .collect();
    let borrowed: Vec<(&str, usize)> = entries
        .iter()
        .map(|(name, size)| (name.as_str(), *size))
        .collect();
    write_zip(&path, &borrowed);

    let listing = listing_for(&path, None).expect("a listing");
    assert!(listing.scan_capped);
    assert!(listing.entries.len() <= 20_000);

    render_fixture(&dir, "capped-dark", &path, TextTheme::Dark, 1_920);
}

#[test]
fn orders_names_the_way_a_person_reads_them() {
    use std::cmp::Ordering;

    assert_eq!(natural_order("report2.txt", "report10.txt"), Ordering::Less);
    assert_eq!(natural_order("Report.txt", "report.txt"), Ordering::Equal);
    assert_eq!(natural_order("beta", "Alpha"), Ordering::Greater);
    assert_eq!(natural_order("a", "a b"), Ordering::Less);
}

#[test]
fn writes_sizes_a_person_reads() {
    assert_eq!(format_size(0), "0 B");
    assert_eq!(format_size(512), "512 B");
    assert_eq!(format_size(4_096), "4 KB");
    assert_eq!(format_size(48_000), "47 KB");
    assert_eq!(format_size(1_200_000), "1.1 MB");
    assert_eq!(format_size(12_400_000), "12 MB");
    assert_eq!(format_size(4_294_967_296), "4 GB");
}

#[test]
fn cuts_a_name_to_the_room_it_has() {
    assert_eq!(cut_to_width("report.pdf", 1_000, 7), "report.pdf");
    assert_eq!(cut_to_width("report.pdf", 7 * 6, 7), "repor…");
    assert_eq!(cut_to_width("report.pdf", 0, 7), "");
}

/// The pictures this test writes are the design under review: a page per
/// archive shape, in both themes, and a narrow box to cut names in.
#[test]
fn draws_the_pages() {
    let dir = scratch("pages");
    let path = bundle(&dir);

    render_fixture(&dir, "bundle-light", &path, TextTheme::Light, 1_920);
    render_fixture(&dir, "bundle-dark", &path, TextTheme::Dark, 1_920);
    render_fixture(&dir, "bundle-narrow", &path, TextTheme::Light, 420);

    let empty = dir.join("empty.zip");
    write_zip(&empty, &[]);
    render_fixture(&dir, "empty-light", &empty, TextTheme::Light, 1_920);

    // More rows than a page draws, so the count of what is left is exercised.
    let many = dir.join("many.zip");
    let entries: Vec<(String, usize)> = (0..140)
        .map(|index| {
            (
                format!("dir{:02}/entry{index:03}.txt", index % 7),
                100 + index,
            )
        })
        .collect();
    let borrowed: Vec<(&str, usize)> = entries
        .iter()
        .map(|(name, size)| (name.as_str(), *size))
        .collect();
    write_zip(&many, &borrowed);
    render_fixture(&dir, "many-dark", &many, TextTheme::Dark, 1_920);
}

use super::*;
use flate2::write::ZlibEncoder;
use flate2::Compression;
use std::io::Write;

/// A file of this module's own folder: the tests run beside each other, and one of them
/// clearing its fixtures must not take another's with it.
fn fixture(name: &str, contents: &[u8]) -> PathBuf {
    let folder = std::env::temp_dir().join("rust-hover-preview-font-tests");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join(name);
    std::fs::write(&path, contents).expect("a written file");
    path
}

/// One table of a font of this module's own making.
struct Table {
    tag: [u8; 4],
    bytes: Vec<u8>,
}

fn table(tag: &[u8; 4], bytes: Vec<u8>) -> Table {
    Table { tag: *tag, bytes }
}

/// The code points of a piece of sample text, which is what a font that "covers" it
/// maps — so a fixture is built from the very lines the specimen is checked against.
fn codes(line: &str) -> Vec<u32> {
    line.chars().map(u32::from).collect()
}

/// An sfnt holding exactly these tables: a version, a directory, and the tables
/// themselves.
fn sfnt(tables: &[Table]) -> Vec<u8> {
    let count = tables.len();
    let mut records = Vec::new();
    let mut data = Vec::new();

    for entry in tables {
        records.extend_from_slice(&entry.tag);
        records.extend_from_slice(&0u32.to_be_bytes()); // checksum, which nothing here reads
        records.extend_from_slice(&((12 + count * 16 + data.len()) as u32).to_be_bytes());
        records.extend_from_slice(&(entry.bytes.len() as u32).to_be_bytes());
        data.extend_from_slice(&entry.bytes);
        while !data.len().is_multiple_of(4) {
            data.push(0);
        }
    }

    let mut font = Vec::new();
    font.extend_from_slice(b"\x00\x01\x00\x00");
    font.extend_from_slice(&(count as u16).to_be_bytes());
    font.extend_from_slice(&[0u8; 6]); // the binary-search fields, which nothing here reads
    font.extend_from_slice(&records);
    font.extend_from_slice(&data);
    font
}

/// A `name` table naming the font in Windows English.
fn name_table(family: &str, style: &str) -> Vec<u8> {
    let mut strings = Vec::new();
    let mut records = Vec::new();

    for (name_id, text) in [(1u16, family), (2u16, style)] {
        let offset = strings.len();
        for unit in text.encode_utf16() {
            strings.extend_from_slice(&unit.to_be_bytes());
        }

        records.extend_from_slice(&3u16.to_be_bytes()); // Windows
        records.extend_from_slice(&1u16.to_be_bytes()); // Unicode BMP
        records.extend_from_slice(&0x0409u16.to_be_bytes()); // English (United States)
        records.extend_from_slice(&name_id.to_be_bytes());
        records.extend_from_slice(&((strings.len() - offset) as u16).to_be_bytes());
        records.extend_from_slice(&(offset as u16).to_be_bytes());
    }

    let mut table = Vec::new();
    table.extend_from_slice(&0u16.to_be_bytes()); // format
    table.extend_from_slice(&2u16.to_be_bytes()); // count
    table.extend_from_slice(&((6 + 2 * 12) as u16).to_be_bytes()); // stringOffset
    table.extend_from_slice(&records);
    table.extend_from_slice(&strings);
    table
}

/// A `cmap` holding one format 12 group per code point, each with a glyph of its own.
fn cmap_table(codes: &[u32]) -> Vec<u8> {
    let mut subtable = Vec::new();
    subtable.extend_from_slice(&12u16.to_be_bytes()); // format
    subtable.extend_from_slice(&0u16.to_be_bytes()); // reserved
    subtable.extend_from_slice(&(16u32 + codes.len() as u32 * 12).to_be_bytes());
    subtable.extend_from_slice(&0u32.to_be_bytes()); // language
    subtable.extend_from_slice(&(codes.len() as u32).to_be_bytes()); // numGroups

    for (index, code) in codes.iter().enumerate() {
        subtable.extend_from_slice(&code.to_be_bytes());
        subtable.extend_from_slice(&code.to_be_bytes());
        subtable.extend_from_slice(&(index as u32 + 1).to_be_bytes()); // a glyph, never 0
    }

    let mut table = Vec::new();
    table.extend_from_slice(&0u16.to_be_bytes()); // version
    table.extend_from_slice(&1u16.to_be_bytes()); // numTables
    table.extend_from_slice(&3u16.to_be_bytes()); // Windows
    table.extend_from_slice(&10u16.to_be_bytes()); // UCS-4
    table.extend_from_slice(&12u32.to_be_bytes()); // offset, past the one record
    table.extend_from_slice(&subtable);
    table
}

/// A `cmap` covering exactly the characters of these lines, which is what a font that
/// "has" them is as far as a specimen is concerned.
fn cmap_for(lines: &[&str]) -> Vec<u8> {
    let mut covered: Vec<u32> = lines.iter().flat_map(|line| codes(line)).collect();
    covered.sort_unstable();
    covered.dedup();

    cmap_table(&covered)
}

fn zlib(bytes: &[u8]) -> Vec<u8> {
    let mut encoder = ZlibEncoder::new(Vec::new(), Compression::default());
    encoder.write_all(bytes).expect("a written table");
    encoder.finish().expect("a finished stream")
}

/// The same tables in a WOFF container: each one deflated where that is smaller than
/// the table, and stored as it is where it is not — which is what the specification
/// asks a writer for.
fn woff(tables: &[Table]) -> Vec<u8> {
    let count = tables.len();
    let mut directory = Vec::new();
    let mut data = Vec::new();

    for entry in tables {
        let compressed = zlib(&entry.bytes);
        let (stored, length) = if compressed.len() < entry.bytes.len() {
            (compressed, entry.bytes.len())
        } else {
            (entry.bytes.clone(), entry.bytes.len())
        };

        directory.extend_from_slice(&entry.tag);
        directory.extend_from_slice(&((44 + count * 20 + data.len()) as u32).to_be_bytes());
        directory.extend_from_slice(&(stored.len() as u32).to_be_bytes());
        directory.extend_from_slice(&(length as u32).to_be_bytes());
        directory.extend_from_slice(&0u32.to_be_bytes()); // checksum

        data.extend_from_slice(&stored);
        while !data.len().is_multiple_of(4) {
            data.push(0);
        }
    }

    let total = (44 + count * 20 + data.len()) as u32;
    let mut font = Vec::new();
    font.extend_from_slice(b"wOFF");
    font.extend_from_slice(b"\x00\x01\x00\x00"); // flavor: TrueType
    font.extend_from_slice(&total.to_be_bytes());
    font.extend_from_slice(&(count as u16).to_be_bytes());
    font.extend_from_slice(&0u16.to_be_bytes()); // reserved
    font.extend_from_slice(&total.to_be_bytes()); // totalSfntSize
    font.extend_from_slice(&1u16.to_be_bytes()); // majorVersion
    font.extend_from_slice(&0u16.to_be_bytes()); // minorVersion
    font.extend_from_slice(&[0u8; 20]); // the metadata and private blocks, which there are none of
    font.extend_from_slice(&directory);
    font.extend_from_slice(&data);
    font
}

/// The same tables in a WOFF2 container: a directory of lengths and one Brotli stream
/// holding every table in the order they were listed.
fn woff2(tables: &[Table]) -> Vec<u8> {
    let mut directory = Vec::new();
    let mut stream = Vec::new();

    for entry in tables {
        let index = WOFF2_KNOWN_TAGS
            .iter()
            .position(|known| *known == &entry.tag)
            .expect("a tag the container numbers") as u8;

        // The transform bits are zero, and for a `cmap` or a `name` that means the table
        // is stored as it is: no transformed length follows.
        directory.push(index);
        directory.extend_from_slice(&base128(entry.bytes.len() as u32));
        stream.extend_from_slice(&entry.bytes);
    }

    let compressed = brotli_compress(&stream);
    let total = (48 + directory.len() + compressed.len()) as u32;

    let mut font = Vec::new();
    font.extend_from_slice(b"wOF2");
    font.extend_from_slice(b"\x00\x01\x00\x00"); // flavor: TrueType
    font.extend_from_slice(&total.to_be_bytes());
    font.extend_from_slice(&(tables.len() as u16).to_be_bytes());
    font.extend_from_slice(&0u16.to_be_bytes()); // reserved
    font.extend_from_slice(&(stream.len() as u32).to_be_bytes()); // totalSfntSize
    font.extend_from_slice(&(compressed.len() as u32).to_be_bytes()); // totalCompressedSize
    font.extend_from_slice(&1u16.to_be_bytes()); // majorVersion
    font.extend_from_slice(&0u16.to_be_bytes()); // minorVersion
    font.extend_from_slice(&[0u8; 20]); // the metadata and private blocks, which there are none of
    font.extend_from_slice(&directory);
    font.extend_from_slice(&compressed);
    font
}

/// A length as the WOFF2 directory writes it.
fn base128(value: u32) -> Vec<u8> {
    let mut groups = vec![(value & 0x7f) as u8];
    let mut rest = value >> 7;

    while rest != 0 {
        groups.push((rest & 0x7f) as u8);
        rest >>= 7;
    }

    groups.reverse();
    let last = groups.len() - 1;
    for group in &mut groups[..last] {
        *group |= 0x80;
    }

    groups
}

fn brotli_compress(bytes: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();

    {
        let mut writer = ::brotli::CompressorWriter::new(&mut out, 4096, 5, 22);
        writer.write_all(bytes).expect("a written stream");
    }

    out
}

/// A collection of the given faces: a directory of faces that share one table pool,
/// which is what a `.ttc` is.
fn collection(faces: &[Vec<Table>]) -> Vec<u8> {
    let header = 12 + faces.len() * 4;
    let directories = faces
        .iter()
        .map(|tables| 12 + tables.len() * 16)
        .sum::<usize>();
    let data_start = header + directories;

    let mut built = Vec::new();
    let mut offsets = Vec::new();
    let mut data = Vec::new();

    for tables in faces {
        offsets.push(((header + built.len()) as u32).to_be_bytes());
        built.extend_from_slice(b"\x00\x01\x00\x00");
        built.extend_from_slice(&(tables.len() as u16).to_be_bytes());
        built.extend_from_slice(&[0u8; 6]); // the binary-search fields

        for entry in tables {
            built.extend_from_slice(&entry.tag);
            built.extend_from_slice(&0u32.to_be_bytes()); // checksum
            built.extend_from_slice(&((data_start + data.len()) as u32).to_be_bytes());
            built.extend_from_slice(&(entry.bytes.len() as u32).to_be_bytes());
            data.extend_from_slice(&entry.bytes);
            while !data.len().is_multiple_of(4) {
                data.push(0);
            }
        }
    }

    let mut font = Vec::new();
    font.extend_from_slice(b"ttcf");
    font.extend_from_slice(&1u16.to_be_bytes()); // majorVersion
    font.extend_from_slice(&0u16.to_be_bytes()); // minorVersion
    font.extend_from_slice(&(faces.len() as u32).to_be_bytes());
    for offset in &offsets {
        font.extend_from_slice(offset);
    }
    font.extend_from_slice(&built);
    font.extend_from_slice(&data);
    font
}

/// A font whose `name` table says one thing and whose `cmap` covers the given lines.
fn font(family: &str, lines: &[&str]) -> Vec<u8> {
    sfnt(&[
        table(b"name", name_table(family, "Regular")),
        table(b"cmap", cmap_for(lines)),
    ])
}

#[test]
fn titles_a_font_with_the_name_it_calls_itself() {
    let path = fixture("family.ttf", &font("Test Family", &[SAMPLE_LINES[0]]));
    let specimen = probe(&path).expect("a specimen");

    assert_eq!(specimen.title, "Test Family Regular");
    assert_eq!(specimen.samples, vec![SAMPLE_LINES[0].to_string()]);
    assert_eq!(specimen.faces, 1);

    let _ = std::fs::remove_file(&path);
}

/// The specimen's lines are the font's own coverage: a line is drawn where every
/// character of it is in the font, and nothing is drawn for the scripts it does not have
/// — which is what keeps a page from showing a Latin-only font's Japanese line in a
/// system font and calling it the font.
#[test]
fn draws_the_lines_the_font_covers_and_no_others() {
    let pangram = SAMPLE_LINES[0];
    let japanese = SAMPLE_LINES[1];
    let chinese = SAMPLE_LINES[2];

    let path = fixture("japanese.ttf", &font("Kana Family", &[pangram, japanese]));
    let specimen = probe(&path).expect("a specimen");
    assert_eq!(
        specimen.samples,
        vec![pangram.to_string(), japanese.to_string()],
        "the pangram first, and only the scripts the font holds"
    );

    let path = fixture("chinese.ttf", &font("Han Family", &[chinese]));
    let specimen = probe(&path).expect("a specimen");
    assert_eq!(
        specimen.samples,
        vec![chinese.to_string()],
        "a font with no Latin is shown by the line it does have"
    );

    let _ = std::fs::remove_file(&path);
}

/// A line is drawn whole or not at all: one character the font has not got is one the
/// page would fall back to a system font for, in the middle of the font's own text.
#[test]
fn draws_no_line_a_font_cannot_complete() {
    let mut partial = codes(SAMPLE_LINES[0]);
    partial.retain(|code| *code != u32::from('z'));

    let path = fixture(
        "partial.ttf",
        &sfnt(&[table(b"cmap", cmap_table(&partial))]),
    );
    let specimen = probe(&path).expect("a specimen");

    assert_eq!(specimen.samples.len(), 1);
    assert_ne!(specimen.samples[0], SAMPLE_LINES[0]);

    let _ = std::fs::remove_file(&path);
}

/// Where none of the lines fits — an Arabic font, a symbol font — the specimen is drawn
/// from the font's own characters rather than from nothing.
#[test]
fn falls_back_to_the_fonts_own_characters() {
    let arabic: Vec<u32> = (0x0627..0x063a).collect();
    let path = fixture("arabic.ttf", &sfnt(&[table(b"cmap", cmap_table(&arabic))]));
    let specimen = probe(&path).expect("a specimen");

    assert_eq!(specimen.samples.len(), 1);
    assert!(
        specimen.samples[0]
            .chars()
            .all(|character| arabic.contains(&u32::from(character))),
        "the line is made of the font's own characters: {:?}",
        specimen.samples[0]
    );

    // And the title falls back to the file's own name where the font names itself
    // nothing.
    assert_eq!(specimen.title, "arabic");

    let _ = std::fs::remove_file(&path);
}

/// A font whose glyphs are all private-use ones — an icon font — is drawn from them:
/// they are the font's own characters, and the icons are what a specimen of such a font
/// is for. What is not drawn is the variation selector beside them, which is a mark on
/// the character before it rather than a character of its own.
#[test]
fn draws_an_icon_font_from_its_own_glyphs() {
    let icons: Vec<u32> = (0xe000..0xe008).collect();
    let covered: Vec<u32> = icons.iter().copied().chain([0xfe0f]).collect();

    let path = fixture("icons.ttf", &sfnt(&[table(b"cmap", cmap_table(&covered))]));
    let specimen = probe(&path).expect("a specimen");

    assert_eq!(specimen.samples.len(), 1);
    assert_eq!(specimen.samples[0].chars().count(), icons.len());
    assert!(
        specimen.samples[0]
            .chars()
            .all(|character| icons.contains(&u32::from(character))),
        "the line is the font's own icons: {:?}",
        specimen.samples[0]
    );

    let _ = std::fs::remove_file(&path);
}

#[test]
fn reads_the_same_tables_out_of_either_webfont_container() {
    let pangram = SAMPLE_LINES[0];
    let tables = [
        table(b"name", name_table("Web Font", "Regular")),
        table(b"cmap", cmap_for(&[pangram])),
    ];

    for (name, bytes) in [
        ("webfont.woff", woff(&tables)),
        ("webfont.woff2", woff2(&tables)),
    ] {
        let path = fixture(name, &bytes);
        let specimen = probe(&path).expect("a specimen");

        assert_eq!(specimen.title, "Web Font Regular", "{name}");
        assert_eq!(specimen.samples, vec![pangram.to_string()], "{name}");

        let _ = std::fs::remove_file(&path);
    }
}

/// A stream that goes on past the last table a specimen reads, which is what a real
/// `.woff2` is: the outlines of a text face are tens of kilobytes, and the tables after
/// its `name` — its `post`, its kerning — are read by nobody here. So the two tables
/// that *are* read lie in a stream that has not ended, and they arrive in reads of the
/// decompressor's own size rather than in one: a reader that insists on whole reads
/// stops short of the table it needs. Every one of the 34 `.woff2` files on the machine
/// this was measured on is this shape, and this is the reader that finds them.
#[test]
fn reads_a_webfont_whose_stream_goes_past_what_it_needs() {
    let pangram = SAMPLE_LINES[0];
    let beyond = 16 * 1024;

    // Two shapes of the same thing: the tables a specimen reads at the front, and a big
    // table between them with the rest behind it, which is what a real directory holds.
    for (name, tables) in [
        (
            "trailing.woff2",
            vec![
                table(b"cmap", cmap_for(&[pangram])),
                table(b"name", name_table("Web Font", "Regular")),
                table(b"post", vec![0u8; beyond]),
            ],
        ),
        (
            "split.woff2",
            vec![
                table(b"cmap", cmap_for(&[pangram])),
                table(b"post", vec![0u8; beyond]),
                table(b"name", name_table("Web Font", "Regular")),
                table(b"GPOS", vec![0u8; beyond / 4]),
            ],
        ),
    ] {
        let path = fixture(name, &woff2(&tables));
        let specimen = probe(&path).expect("a specimen");

        assert_eq!(specimen.title, "Web Font Regular", "{name}");
        assert_eq!(specimen.samples, vec![pangram.to_string()], "{name}");

        let _ = std::fs::remove_file(&path);
    }
}

/// A collection is drawn from one face of it, which is the one thing a page cannot be
/// pointed at inside a `.ttc`: the face is written out as a font of its own, and what
/// comes out is that face's tables rather than the face next to it.
#[test]
fn writes_out_the_face_a_specimen_was_read_at() {
    let path = fixture(
        "collection.ttc",
        &collection(&[
            vec![
                table(b"name", name_table("First Family", "Regular")),
                table(b"cmap", cmap_for(&[SAMPLE_LINES[0]])),
            ],
            vec![
                table(b"name", name_table("Second Family", "Bold")),
                table(b"cmap", cmap_for(&[SAMPLE_LINES[1]])),
            ],
        ]),
    );

    let first = probe_face(&path, 0).expect("a specimen");
    assert_eq!(first.faces, 2);
    assert_eq!(first.face, 0);
    assert_eq!(first.title, "First Family Regular (1 of 2)");
    assert_eq!(first.samples, vec![SAMPLE_LINES[0].to_string()]);

    // The face beside it is another font in the same file and a specimen of its own: its
    // own name, its own characters, and a heading saying which face of the file it is.
    let second = probe_face(&path, 1).expect("a specimen");
    assert_eq!(second.face, 1);
    assert_eq!(second.title, "Second Family Bold (2 of 2)");
    assert_eq!(second.samples, vec![SAMPLE_LINES[1].to_string()]);

    // Held per face: the first face asked for again is still the answer it was, and not
    // the answer the second one was read as.
    assert_eq!(probe_face(&path, 0), Some(first.clone()));

    // A face past the last one a file holds is the last one it has, and the heading says
    // which face came out rather than which was asked for.
    let last = probe_face(&path, 7).expect("a specimen");
    assert_eq!(last.face, 1);
    assert_eq!(last.title, "Second Family Bold (2 of 2)");

    // What the engine is handed is the face as a font of its own — another face being
    // another file — and what it holds is that face's tables.
    let folder = std::env::temp_dir().join("rust-hover-preview-font-tests-extracted");
    let first_source = browser_source(&path, &first, &folder).expect("a face");
    let second_source = browser_source(&path, &second, &folder).expect("a face");
    assert_ne!(first_source, path);
    assert_ne!(second_source, first_source, "another face is another file");

    let bytes = std::fs::read(&second_source).expect("a written face");
    let tables = tables_of(&bytes, 0).expect("a font of its own");
    let cmap = CharacterMap::parse(tables.cmap.as_deref().expect("a cmap")).expect("a map");

    // The second face's own map — which holds the kana line — and not the face beside it,
    // whose map is the pangram.
    assert!(cmap.maps('い'), "the second face's own map");
    assert!(!cmap.maps('z'), "and not the face beside it");

    // A single-face font is handed over as it is, with nothing written anywhere.
    let single = fixture("single.ttf", &font("Single Family", &[SAMPLE_LINES[0]]));
    let specimen = probe(&single).expect("a specimen");
    assert_eq!(
        browser_source(&single, &specimen, &folder),
        Some(single.clone())
    );

    let _ = std::fs::remove_file(&path);
    let _ = std::fs::remove_file(&single);
    let _ = std::fs::remove_dir_all(&folder);
}

/// Which face a specimen is read at is the setting's own number less one: faces are
/// numbered from `1` by the menu and from `0` by a collection's own directory, and the
/// number is what the two have to agree about.
#[test]
fn reads_the_face_the_setting_names() {
    let before = crate::CONFIG.lock().expect("a configuration").ttc_face;

    for (named, index) in [(1, 0), (4, 3), (crate::config::config::MAX_TTC_FACE, 9)] {
        crate::CONFIG.lock().expect("a configuration").ttc_face = named;

        assert_eq!(configured_face(), index, "face {named}");
    }

    crate::CONFIG.lock().expect("a configuration").ttc_face = before;
}

/// The scripts that had no line of their own are sampled by their own script's line
/// rather than by the first characters their map happens to hold — which is what an
/// Arabic, Hebrew, Thai or Devanagari font used to be shown by.
#[test]
fn draws_the_line_of_the_script_a_font_is_of() {
    // The four lines, in the order the constants list them.
    for (script, line) in SAMPLE_LINES[6..].iter().enumerate() {
        let path = fixture(
            &format!("script-{script}.ttf"),
            &font("Script Family", &[line]),
        );
        let specimen = probe(&path).expect("a specimen");

        assert_eq!(
            specimen.samples,
            vec![line.to_string()],
            "a font of one script is shown by that script's line and nothing else"
        );

        let _ = std::fs::remove_file(&path);
    }
}

#[test]
fn a_file_that_is_not_a_font_makes_no_specimen() {
    for (name, contents) in [
        ("not-a-font.ttf", b"just some bytes".as_slice()),
        ("empty.ttf", b"".as_slice()),
        // A file that begins like an sfnt but holds copies of nothing.
        ("truncated.ttf", b"\x00\x01\x00\x00\x00\x04".as_slice()),
        // A font with no character map has no characters to show.
        (
            "mapless.ttf",
            sfnt(&[table(b"name", name_table("No Map", "Regular"))]).as_slice(),
        ),
    ] {
        let path = fixture(name, contents);

        assert_eq!(probe(&path), None, "{name}");

        let _ = std::fs::remove_file(&path);
    }
}

/// A font rewritten in place is read again: the version of the file is part of what a
/// held specimen is valid for.
#[test]
fn a_revised_font_is_read_again() {
    let path = fixture("revised.ttf", &font("Before", &[SAMPLE_LINES[0]]));
    assert_eq!(probe(&path).expect("a specimen").title, "Before Regular");

    std::fs::write(&path, font("After", &[SAMPLE_LINES[0], SAMPLE_LINES[1]]))
        .expect("a rewritten font");

    let specimen = probe(&path).expect("a specimen");
    assert_eq!(specimen.title, "After Regular");
    assert_eq!(specimen.samples.len(), 2);

    let _ = std::fs::remove_file(&path);
}

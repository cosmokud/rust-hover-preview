//! Font previews: what a font is, the lines its own character map covers, and the file the
//! browser engine draws it from.
//!
//! A font is not a picture to a decoder either, and this app does not rasterize one — the
//! release binary carries no rasterizer, no shaping stack and no font parser of the kind
//! that draws text. What draws a font is the same browser engine the SVG previews use, in
//! a window of its own, pointed at a page of this app's own with the font in it through
//! `@font-face`; see `webview_preview`. That engine reads all five formats a font goes by,
//! TrueType and CFF outlines, the two webfont containers and a collection of faces, so the
//! drawing costs this app no reader at all.
//!
//! What is left here is the specimen, and the one question the engine cannot answer.
//!
//! The question is what a font *covers*. A browser falls back per glyph and says nothing
//! about it: a page that draws 「いろは」 in a Latin-only font draws it in a system font, at
//! the same size and in the same layout, and a preview that showed it would be claiming
//! something about the font that is not true. So the sample lines *are* the font's own
//! coverage: each one is checked against the font's character map, and only the lines the
//! map answers for are drawn — the pangram always among them where the font has Latin, and
//! a line apiece where it has Japanese, Chinese, Korean, Cyrillic, Greek, Arabic, Hebrew,
//! Thai or Devanagari. A font of a script there is no line for — a Georgian one, a symbol
//! one — is drawn from the characters its own map holds, and so is a font with no script at
//! all: an icon font's glyphs are private-use ones, and a line of them is what a specimen of
//! such a font is for. Both are the same answer reached the other way round.
//!
//! Answering it costs a read of the file under the budget every other read is answered
//! under, and a parse of two of its tables: the character map, and the `name` table the
//! preview is titled with. Nothing is rasterized, nothing is held decoded, and what is kept
//! between hovers is the answer — keyed by the file and the version of it that was read, the
//! way an SVG document's measurement is. The two table readers are a WOFF2 font's Brotli
//! stream and a WOFF one's per-table zlib, which is all that stands between a webfont and
//! the same two tables a `.ttf` holds in the open.
//!
//! One thing is written to disk, and it is the one thing the engine cannot be handed: a page
//! has no syntax for naming a face inside a collection, so a `.ttc` is answered with the face
//! the setting names written out as a font of its own, beside the browser's own profile
//! folder and named for the file, the version of it and the face — where the next run's
//! startup clears it away with everything else that folder held.

use crate::config::config::{decode_budget_bytes, read_within_budget, DEFAULT_TTC_FACE};
use crate::CONFIG;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::SystemTime;

/// The box a font preview is measured at, in the same units a document's own size is:
/// what the layout places the engine's window by.
///
/// A font has no size it asks to be drawn at — what a file holds is outlines, and the text
/// they are drawn as is whatever size a caller asks for — so the box a specimen is placed
/// in is this app's own: a page a shade wider than tall, which is the shape the pangram and
/// a line or two under it want. What the setting beside it names is the share of the
/// display that box takes, the way `vector_scale` names one for a drawing.
pub const SPECIMEN_WIDTH: u32 = 1500;
pub const SPECIMEN_HEIGHT: u32 = 1000;

/// The lines a specimen may be drawn from, in the order they are listed: the pangram every
/// Latin font is judged by, and then one line for each script a font may hold instead of —
/// or as well as — Latin.
///
/// A line is drawn only where the font's own character map covers every character of it;
/// see this module's documentation. Each is the closest thing its script has to a pangram,
/// which is what a specimen is for: the Japanese line is the iroha, the poem the kana were
/// ordered by for a thousand years; the Chinese one is the first line of the Thousand
/// Character Classic, four characters of which carry almost every stroke a Han glyph is
/// built from; the Cyrillic, Greek, Arabic and Hebrew ones are the pangrams those scripts
/// are sampled with; and the Thai and Devanagari ones are the openings of the pangrams
/// those scripts are sampled with, the rest of each being words a font is as likely to be
/// asked for as it is to have — and a line is drawn whole or not at all, so a word a font
/// has not got costs the script its line rather than part of one.
const SAMPLE_LINES: [&str; 10] = [
    "The quick brown fox jumps over the lazy dog.",
    "いろはにほへと ちりぬるを",
    "天地玄黄 宇宙洪荒",
    "다람쥐 헌 쳇바퀴에 타고파",
    "Съешь же ещё этих мягких французских булок да выпей чаю",
    "Ξεσκεπάζω την ψυχοφθόρα βδελυγμία",
    "نص حكيم له سر قاطع وذو شأن عظيم مكتوب على ثوب أخضر ومغلف بجلد أزرق.",
    "דג סקרן שט בים מאוכזב ולפתע מצא חברה",
    "เป็นมนุษย์สุดประเสริฐเลิศคุณค่า กว่าบรรดาฝูงสัตว์เดรัจฉาน",
    "ऋषियों को सताने वाले दुष्ट राक्षसों के राजा रावण का सर्वनाश",
];

/// How many characters the last-resort line carries: what a font that covers none of the
/// lines above is shown by, taken from its own character map. A line of two dozen glyphs is
/// enough to see what a face looks like, and short enough that the line stays one.
const FALLBACK_LINE_CHARACTERS: usize = 32;

/// A specimen: the title the preview is headed with, and the lines drawn under it.
///
/// The first line is the headline — the pangram wherever the font has Latin, and the
/// font's own characters where it has none of the lines this app knows — and the rest are
/// drawn smaller, one script apiece.
#[derive(Clone, PartialEq, Eq, Debug)]
pub struct Specimen {
    pub title: String,
    pub samples: Vec<String>,
    /// How many faces the file holds: more than one is a collection, which the engine can
    /// be pointed at only through a face written out on its own; see `browser_source`.
    pub faces: usize,
    /// Which of those faces this is, as an index into the collection's directory — the face
    /// the setting asked for, reduced to the last one the file holds where it holds fewer
    /// than that, and `0` for a file that is not a collection at all. It is the face the
    /// title counts from one and the one `browser_source` writes out.
    pub face: usize,
}

/// The specimen `path` makes at the face the setting names: see `configured_face`.
///
/// The answer is held between hovers — the file, the version of it that was read, and the
/// face — for the reason a document's measurement is: the layout asks for it on every
/// hover, and a pointer swept back and forth over a folder meets the same files again. A
/// file that turned out not to be a font at all is held as that, so a hover onto it costs
/// nothing after the first.
pub fn probe(path: &Path) -> Option<Specimen> {
    probe_face(path, configured_face())
}

/// The specimen `path` makes at one face of a collection: `face` is an index into the
/// faces a `.ttc` holds, and a file that holds a single font answers it the same way
/// whatever it is.
///
/// The face is part of what a held specimen is valid for, because it is part of what the
/// specimen is: the two faces of one collection are two fonts that share a file, and an
/// answer read at one of them is not an answer for the other.
pub fn probe_face(path: &Path, face: usize) -> Option<Specimen> {
    let key = FontKey {
        path: path.to_path_buf(),
        version: file_version(path),
        face,
    };

    if let Some(held) = held(&key) {
        return held;
    }

    let specimen = read_specimen(path, face);
    hold(&key, specimen.clone());

    specimen
}

/// Which face of a collection a preview is of, as the setting numbers it — from `1`, the
/// first face, down to the last face a file holds.
///
/// It is read from the configuration each time rather than captured, so the tray's `Font
/// Face` setting applies to the next hover rather than to the next run. What comes back is
/// the index the face is read at, which is the setting's own number less one; a file with
/// fewer faces than that is read at the last one it has, so a collection of two is its
/// second face whatever past the second the setting asks for.
pub fn configured_face() -> usize {
    CONFIG
        .lock()
        .map(|config| config.ttc_face)
        .unwrap_or(DEFAULT_TTC_FACE)
        .saturating_sub(1) as usize
}

/// The file the engine is pointed at for `path`: the font itself, or — for a collection,
/// which no page can name a face of — the face the specimen was read at, written out as a
/// font of its own.
///
/// The extracted face lands in `folder`, which is the browser's own profile folder for this
/// run; it is named for the file, the version of it that was read and the face, so the same
/// face of the same collection hovered again is the same file, another face is another, and
/// an edited one is another again. What this costs is one write per face per version, and
/// what it is not is permanent: the folder is cleared at the next start, with the rest of
/// what a run leaves behind.
pub fn browser_source(path: &Path, specimen: &Specimen, folder: &Path) -> Option<PathBuf> {
    if specimen.faces <= 1 {
        return Some(path.to_path_buf());
    }

    let (offset, flavor) = collection_face(path, specimen.face)?;
    let extension = if flavor == *b"OTTO" { "otf" } else { "ttf" };
    let extracted = folder.join("fonts").join(format!(
        "{:016x}-{}-{}.{extension}",
        path_hash(path),
        file_stamp(path),
        specimen.face
    ));

    if extracted.is_file() {
        return Some(extracted);
    }

    let bytes = read_within_budget(path)?;
    let face = extract_face(&bytes, offset)?;

    std::fs::create_dir_all(extracted.parent()?).ok()?;
    std::fs::write(&extracted, face).ok()?;

    Some(extracted)
}

/// The specimen a file makes at one of its faces, read fresh.
fn read_specimen(path: &Path, face: usize) -> Option<Specimen> {
    let bytes = read_within_budget(path)?;
    let tables = tables_of(&bytes, face)?;

    // What a specimen is checked against, and what it is drawn from where none of the
    // lines fit: a font with no character map has no characters, so there is nothing to
    // show and nothing to check.
    let cmap = tables.cmap.as_deref().and_then(CharacterMap::parse)?;
    let samples = sample_lines(&cmap);

    if samples.is_empty() {
        return None;
    }

    Some(Specimen {
        title: specimen_title(path, tables.name.as_deref(), tables.faces, tables.face),
        samples,
        faces: tables.faces,
        face: tables.face,
    })
}

/// The lines a specimen is drawn from: every line of the ten the font's own character map
/// covers, or — where it covers none of them — one line of its own characters.
fn sample_lines(cmap: &CharacterMap) -> Vec<String> {
    let covered: Vec<String> = SAMPLE_LINES
        .iter()
        .filter(|line| covers(cmap, line))
        .map(|line| line.to_string())
        .collect();

    if !covered.is_empty() {
        return covered;
    }

    let characters = cmap.drawable_codes(FALLBACK_LINE_CHARACTERS);

    if characters.is_empty() {
        Vec::new()
    } else {
        vec![characters.into_iter().collect()]
    }
}

/// Whether a font can draw every character of a line.
///
/// All or nothing, and the space between two words is not a character a font has to have:
/// a line drawn with one character falling back to a system font would be a preview of two
/// fonts at once, which is exactly what the coverage check exists to keep off the screen.
fn covers(cmap: &CharacterMap, line: &str) -> bool {
    line.chars()
        .filter(|character| !character.is_whitespace())
        .all(|character| cmap.maps(character))
}

/// What a specimen is headed with: the family and style the font calls itself, or the
/// file's own name where its `name` table is missing or says nothing.
///
/// A collection is noted as such and noted for which of its faces is drawn, because what
/// is drawn is one face of it: the file is a font, and the preview is of the face the
/// setting asked for — or of the last one the file holds, where it holds fewer than that.
fn specimen_title(path: &Path, name: Option<&[u8]>, faces: usize, face: usize) -> String {
    let named = name
        .and_then(name_strings)
        .map(|(family, style)| match style {
            Some(style) if !style.is_empty() => format!("{family} {style}"),
            _ => family,
        });

    let title = named.unwrap_or_else(|| {
        path.file_stem()
            .or_else(|| path.file_name())
            .map(|name| name.to_string_lossy().into_owned())
            .unwrap_or_default()
    });

    if faces > 1 {
        format!("{title} ({} of {faces})", face + 1)
    } else {
        title
    }
}

/// The tables a specimen is read from, out of whichever container the file turned out to
/// be: a font's own table directory, a webfont's compressed one, or one face of a
/// collection.
struct Tables {
    /// The character map, as the font stores it.
    cmap: Option<Vec<u8>>,
    /// The name table, as the font stores it.
    name: Option<Vec<u8>>,
    /// How many faces the file holds: more than one is a collection.
    faces: usize,
    /// Which of those faces these tables came out of, as an index into it — `0` for a file
    /// that holds a single font, whatever face the setting named.
    face: usize,
}

/// Which container the file is, by its own first bytes rather than by what it is called: a
/// collection, one of the two webfont containers, or an sfnt written out in the open.
///
/// `face` is which face of a collection to read, and is ignored by the containers that hold
/// one font: a `.ttf` is its own face whatever number the setting holds. A file that is
/// none of the containers — a `.ttf` that holds something else, a download that never
/// finished — is answered with nothing, which is what keeps a hover from opening a box
/// nothing would be drawn into.
fn tables_of(bytes: &[u8], face: usize) -> Option<Tables> {
    let version = bytes.get(..4)?;

    if version == b"ttcf" {
        return collection_tables(bytes, face);
    }
    if version == b"wOFF" {
        return woff_tables(bytes);
    }
    if version == b"wOF2" {
        return woff2_tables(bytes);
    }
    if is_sfnt_version(version) {
        return sfnt_tables(bytes, 0);
    }

    None
}

/// Whether these four bytes begin an sfnt: the TrueType version, the CFF one, or the older
/// spellings Apple's own fonts still carry.
fn is_sfnt_version(bytes: &[u8]) -> bool {
    bytes == b"\x00\x01\x00\x00" || bytes == b"OTTO" || bytes == b"true" || bytes == b"typ1"
}

/// The two tables, out of the table directory that begins at `at` — a file's own, or the
/// directory of one face inside a collection.
fn sfnt_tables(bytes: &[u8], at: usize) -> Option<Tables> {
    Some(Tables {
        cmap: table(bytes, at, b"cmap").map(<[u8]>::to_vec),
        name: table(bytes, at, b"name").map(<[u8]>::to_vec),
        faces: 1,
        face: 0,
    })
}

/// One table of an sfnt, by its tag: what its own table directory says the table is and
/// where it lies, which is the only place a table's size is written down.
fn table<'a>(bytes: &'a [u8], at: usize, tag: &[u8; 4]) -> Option<&'a [u8]> {
    let count = be_u16(bytes, at + 4)? as usize;

    for index in 0..count {
        let record = at + 12 + index * 16;
        if bytes.get(record..record + 4)? != tag {
            continue;
        }

        let offset = be_u32(bytes, record + 8)? as usize;
        let length = be_u32(bytes, record + 12)? as usize;

        return bytes.get(offset..offset.checked_add(length)?);
    }

    None
}

/// The two tables, out of the face of a collection the setting asks for: a `.ttc` is a
/// directory of faces that share their table data, and what a specimen can be read from is
/// the directory at one of the offsets it lists.
///
/// `face` is that face as an index, and a file that holds fewer faces than it names is read
/// at the last one it has — which is the same reduction `collection_face` makes where the
/// face is written out, so the specimen and the file the engine draws are one face.
fn collection_tables(bytes: &[u8], face: usize) -> Option<Tables> {
    let faces = be_u32(bytes, 8)? as usize;
    if faces == 0 {
        return None;
    }

    let at = face.min(faces - 1);
    let mut tables = sfnt_tables(bytes, be_u32(bytes, 12 + at * 4)? as usize)?;
    tables.faces = faces;
    tables.face = at;

    Some(tables)
}

/// The two tables, out of a WOFF container: the same sfnt tables with each of them
/// compressed on its own and a directory that says where each one lies.
fn woff_tables(bytes: &[u8]) -> Option<Tables> {
    let count = be_u16(bytes, 12)? as usize;
    let mut cmap = None;
    let mut name = None;

    for index in 0..count {
        let record = 44 + index * 20;
        let tag = bytes.get(record..record + 4)?;
        if tag != b"cmap" && tag != b"name" {
            continue;
        }

        let offset = be_u32(bytes, record + 4)? as usize;
        let stored = be_u32(bytes, record + 8)? as usize;
        let original = be_u32(bytes, record + 12)? as usize;
        let data = bytes.get(offset..offset.checked_add(stored)?)?;
        let table = inflate(data, stored, original)?;

        if tag == b"cmap" {
            cmap = Some(table);
        } else {
            name = Some(table);
        }
    }

    Some(Tables {
        cmap,
        name,
        faces: 1,
        face: 0,
    })
}

/// One table of a WOFF container as the font has it: a table the container stored smaller
/// than it is was deflated, and one it stored whole — the spec's own way of saying a table
/// did not compress well — is the table itself.
fn inflate(data: &[u8], stored: usize, original: usize) -> Option<Vec<u8>> {
    if stored >= original {
        return Some(data.to_vec());
    }

    if original as u64 > decode_budget_bytes() {
        return None;
    }

    let mut table = Vec::new();
    flate2::read::ZlibDecoder::new(data)
        .take(original as u64 + 1)
        .read_to_end(&mut table)
        .ok()?;

    (table.len() <= original).then_some(table)
}

/// The tables a WOFF2 font numbers its table directory by, which is how the format keeps
/// the entry small: six bits of the flag byte are an index into this list, and the seventh
/// value of those six — `63` — says the four letters follow the flag.
const WOFF2_KNOWN_TAGS: [&[u8; 4]; 63] = [
    b"cmap", b"head", b"hhea", b"hmtx", b"maxp", b"name", b"OS/2", b"post", b"cvt ", b"fpgm",
    b"glyf", b"loca", b"prep", b"CFF ", b"VORG", b"EBDT", b"EBLC", b"gasp", b"hdmx", b"kern",
    b"LTSH", b"PCLT", b"VDMX", b"vhea", b"vmtx", b"BASE", b"GDEF", b"GPOS", b"GSUB", b"EBSC",
    b"JSTF", b"MATH", b"CBDT", b"CBLC", b"COLR", b"CPAL", b"SVG ", b"sbix", b"acnt", b"avar",
    b"bdat", b"bloc", b"bsln", b"cvar", b"fdsc", b"feat", b"fmtx", b"fvar", b"gvar", b"hsty",
    b"just", b"lcar", b"mort", b"morx", b"opbd", b"prop", b"trak", b"Zapf", b"Silf", b"Glat",
    b"Gloc", b"Feat", b"Sill",
];

/// The two tables, out of a WOFF2 container: a directory that says how large each table is
/// and in what order they were written, and then one Brotli stream holding all of them.
///
/// The stream is decompressed only as far as the last table this needs, and only up to what
/// one hover may decode for, which is what keeps a crafted webfont from being a
/// decompression bomb: what a stream can ask for is bounded before it is inflated, the same
/// way a `.svgz` and an animation's frames are.
///
/// Nothing is reconstructed. A WOFF2 font may store its outlines and its horizontal metrics
/// transformed — that is what its own reader puts back — but the character map and the name
/// are never transformed, so what is read here is the table as the font wrote it.
fn woff2_tables(bytes: &[u8]) -> Option<Tables> {
    let count = be_u16(bytes, 12)? as usize;
    let compressed_size = be_u32(bytes, 20)? as usize;

    let mut offset = 48;
    let mut stream_at = 0usize;
    let mut cmap = None;
    let mut name = None;

    for _ in 0..count {
        let flag = *bytes.get(offset)?;
        offset += 1;

        let index = (flag & 0x3f) as usize;
        let tag: [u8; 4] = if index == 63 {
            let tag = bytes.get(offset..offset + 4)?;
            offset += 4;
            tag.try_into().ok()?
        } else {
            **WOFF2_KNOWN_TAGS.get(index)?
        };

        let original = base128(bytes, &mut offset)? as usize;

        // The two bits above the table's index are the transformation the data is in, and
        // the one table whose transformed form is the *zero* version is the outline pair:
        // a `glyf` is stored transformed at 0 and whole at 3, everything else the other way
        // round. What a transformed table's length in the stream is, is its transformed
        // length, which is the second length the directory carries.
        let version = flag >> 6;
        let transformed = match &tag {
            b"glyf" | b"loca" => version == 0,
            _ => version != 0,
        };
        let length = if transformed {
            base128(bytes, &mut offset)? as usize
        } else {
            original
        };

        if &tag == b"cmap" {
            cmap = Some((stream_at, length));
        } else if &tag == b"name" {
            name = Some((stream_at, length));
        }

        stream_at += length;
    }

    let data = bytes.get(offset..offset.checked_add(compressed_size)?)?;
    let needed = cmap
        .iter()
        .chain(name.iter())
        .map(|(at, length)| at + length)
        .max()
        .unwrap_or(0);

    let stream = decompress(data, needed)?;
    let slice = |held: Option<(usize, usize)>| {
        held.and_then(|(at, length)| stream.get(at..at + length))
            .map(<[u8]>::to_vec)
    };

    Some(Tables {
        cmap: slice(cmap),
        name: slice(name),
        faces: 1,
        face: 0,
    })
}

/// A `UIntBase128`: the way a length is written in a WOFF2 table directory — seven bits to
/// a byte, most significant first, the top bit saying another byte follows.
fn base128(bytes: &[u8], at: &mut usize) -> Option<u32> {
    let mut value: u32 = 0;

    for index in 0..5 {
        let byte = *bytes.get(*at)?;
        *at += 1;

        // A leading zero is not a smaller number, it is a byte the writer had no reason to
        // write; the specification refuses it and so does this.
        if index == 0 && byte == 0x80 {
            return None;
        }

        value = value.checked_mul(128)?.checked_add((byte & 0x7f) as u32)?;

        if byte & 0x80 == 0 {
            return Some(value);
        }
    }

    None
}

/// The decompressed table stream of a WOFF2 font, stopped at `needed` bytes.
///
/// The stream is one Brotli stream, so a table near its start cannot be read without
/// inflating everything before it — and a stream is allowed to say it expands to any size
/// at all. What bounds that is what the decompressor is asked for: the bytes up to
/// `needed` and not one more, which is a stream that stops where it is no longer read
/// rather than one that is inflated whole and thrown away. `needed` is the end of the last
/// table a specimen reads, so it is bounded by what the font says it holds, and by what a
/// hover is allowed to decode for as well.
///
/// A stream that gives less than the directory said it holds is not read further: a
/// truncated one, and one that is not a Brotli stream at all, are both answered the same
/// way — with nothing, which is what keeps a hover from opening a box nothing is drawn in.
fn decompress(data: &[u8], needed: usize) -> Option<Vec<u8>> {
    if needed == 0 || needed as u64 > decode_budget_bytes() {
        return None;
    }

    let mut reader = brotli_decompressor::Decompressor::new(data, 4096);
    let mut stream = Vec::with_capacity(needed.min(1 << 20));
    let mut buffer = vec![0u8; needed.min(1 << 16)];

    while stream.len() < needed {
        let want = (needed - stream.len()).min(buffer.len());

        match reader.read(&mut buffer[..want]) {
            Ok(0) | Err(_) => break,
            Ok(read) => stream.extend_from_slice(&buffer[..read]),
        }
    }

    (stream.len() >= needed).then_some(stream)
}

/// A font's own character map: what it can draw, asked of it one code point at a time.
///
/// Only the Unicode subtables are read — the Windows ones, the Unicode platform's, and the
/// Mac byte-map a symbol font may carry instead — and of those, the best one the `cmap`
/// holds: a full Unicode map over a BMP one, either over a symbol one, which is the map
/// whose code points are the Latin letters shifted into the private use area.
struct CharacterMap {
    bytes: Vec<u8>,
    offset: usize,
    format: u16,
    symbol: bool,
}

impl CharacterMap {
    /// The best subtable of a `cmap`, or nothing for a table with none this reads.
    fn parse(cmap: &[u8]) -> Option<Self> {
        let count = be_u16(cmap, 2)? as usize;
        let mut chosen: Option<(u32, usize, bool)> = None;

        for index in 0..count {
            let record = 4 + index * 8;
            let platform = be_u16(cmap, record)?;
            let encoding = be_u16(cmap, record + 2)?;
            let offset = be_u32(cmap, record + 4)? as usize;
            let format = be_u16(cmap, offset)?;

            let Some(score) = subtable_score(platform, encoding) else {
                continue;
            };
            // The formats this reads: the two a modern font uses, the range one an older
            // one does, and the byte map a symbol font may be written with.
            let weight = match format {
                12 => score * 4 + 3,
                4 => score * 4 + 2,
                6 => score * 4 + 1,
                0 => score * 4,
                _ => continue,
            };

            if chosen.is_none_or(|(best, _, _)| weight > best) {
                chosen = Some((weight, offset, platform == 3 && encoding == 0));
            }
        }

        let (_, offset, symbol) = chosen?;

        Some(Self {
            format: be_u16(cmap, offset)?,
            offset,
            symbol,
            bytes: cmap.to_vec(),
        })
    }

    /// Whether the font draws this character at all.
    fn maps(&self, character: char) -> bool {
        self.glyph(u32::from(character)).is_some()
    }

    /// The glyph the map sends a code point to, or nothing where it sends it nowhere — a
    /// glyph of zero is the font saying it has no such character, however the segment or
    /// the group it fell in was written.
    fn glyph(&self, code: u32) -> Option<u16> {
        // A symbol map is the Latin letters shifted into the private use area, which is
        // what a page that writes `A` is asking such a font for.
        let code = if self.symbol && code < 0x100 {
            code + 0xf000
        } else {
            code
        };

        let glyph = match self.format {
            0 => byte_map_glyph(&self.bytes, self.offset, code)?,
            4 => segment_map_glyph(&self.bytes, self.offset, code)?,
            6 => range_map_glyph(&self.bytes, self.offset, code)?,
            12 => group_map_glyph(&self.bytes, self.offset, code)?,
            _ => return None,
        };

        (glyph != 0).then_some(glyph)
    }

    /// The characters the map holds, in its own order, up to `limit` of them that a
    /// specimen can draw: what a font that covers none of the sample lines is shown by.
    fn drawable_codes(&self, limit: usize) -> Vec<char> {
        let mut characters = Vec::new();
        let mut visit = |code: u32| {
            if let Some(character) = drawable(code) {
                characters.push(character);
            }
            characters.len() < limit
        };

        match self.format {
            0 => byte_map_codes(&self.bytes, self.offset, &mut visit),
            4 => segment_map_codes(&self.bytes, self.offset, &mut visit),
            6 => range_map_codes(&self.bytes, self.offset, &mut visit),
            12 => group_map_codes(&self.bytes, self.offset, &mut visit),
            _ => {}
        }

        characters
    }
}

/// How good a `cmap` subtable is for asking what a font draws: the full Unicode maps first,
/// then the BMP-only ones, and the symbol map — which is a Unicode map, with its code
/// points shifted — behind them. A subtable of a platform this does not read is no score at
/// all.
fn subtable_score(platform: u16, encoding: u16) -> Option<u32> {
    match (platform, encoding) {
        (3, 10) | (0, 4) | (0, 5) | (0, 6) => Some(3),
        (3, 1) | (0, 3) | (0, 2) | (0, 1) | (0, 0) => Some(2),
        (3, 0) => Some(1),
        (1, 0) => Some(1),
        _ => None,
    }
}

/// The glyph a format 4 subtable sends a BMP code point to, by the specification's own
/// arithmetic: the segment the code falls in decides it, and its glyph comes from the
/// segment's delta, from the glyph array the range offset points into, or from neither.
fn segment_map_glyph(bytes: &[u8], offset: usize, code: u32) -> Option<u16> {
    if code > 0xffff {
        return None;
    }

    let code = code as u16;
    let segments = (be_u16(bytes, offset + 6)? / 2) as usize;
    let end_codes = offset + 14;
    let start_codes = end_codes + segments * 2 + 2;
    let deltas = start_codes + segments * 2;
    let range_offsets = deltas + segments * 2;

    for index in 0..segments {
        let end = be_u16(bytes, end_codes + index * 2)?;
        if code > end {
            continue;
        }

        // The segments are in ascending order, so a code below the start of the one that
        // ends past it is in none of them.
        let start = be_u16(bytes, start_codes + index * 2)?;
        if code < start {
            return None;
        }

        let delta = be_u16(bytes, deltas + index * 2)?;
        let range_offset = be_u16(bytes, range_offsets + index * 2)?;

        if range_offset == 0 {
            return Some(code.wrapping_add(delta));
        }

        // The range offset is measured from its own place in the array rather than from the
        // table's start, which is the arithmetic that makes this format what it is.
        let at = range_offsets + index * 2 + range_offset as usize + (code - start) as usize * 2;

        return be_u16(bytes, at);
    }

    None
}

/// The glyph a format 12 subtable sends a code point to: the groups are in ascending order
/// and a group's glyphs are consecutive from its first.
fn group_map_glyph(bytes: &[u8], offset: usize, code: u32) -> Option<u16> {
    let groups = be_u32(bytes, offset + 12)? as usize;

    for index in 0..groups {
        let at = offset + 16 + index * 12;
        let start = be_u32(bytes, at)?;

        if code < start {
            return None;
        }

        if code <= be_u32(bytes, at + 4)? {
            let glyph = be_u32(bytes, at + 8)? as u64 + (code - start) as u64;
            return u16::try_from(glyph).ok();
        }
    }

    None
}

/// The glyph a format 6 subtable sends a code point to: one range, and the glyphs its own
/// array holds.
fn range_map_glyph(bytes: &[u8], offset: usize, code: u32) -> Option<u16> {
    let first = be_u16(bytes, offset + 6)? as u32;
    let count = be_u16(bytes, offset + 8)? as u32;
    let index = code.checked_sub(first)?;

    if index >= count {
        return None;
    }

    be_u16(bytes, offset + 10 + index as usize * 2)
}

/// The glyph a format 0 subtable sends a byte to.
fn byte_map_glyph(bytes: &[u8], offset: usize, code: u32) -> Option<u16> {
    if code > 0xff {
        return None;
    }

    Some(u16::from(*bytes.get(offset + 6 + code as usize)?))
}

/// Whether a code point is one a specimen can draw: not a control or a space, and not a
/// variation selector, which is a mark on the character before it rather than a character.
///
/// A private-use one *is* drawn, which is the whole of what an icon font has: its glyphs
/// are the ones it keeps in a private-use area, with no character of any script behind
/// them, and a line of them is what a specimen of such a font is for. What keeps them off
/// the lines a text font is judged by is that only the last resort asks for them: the lines
/// are the ones this app knows, and a font is drawn from its own characters only where it
/// covers none of them — which no font that draws text ever does.
fn drawable(code: u32) -> Option<char> {
    let character = char::from_u32(code)?;

    if character.is_control() || character.is_whitespace() {
        return None;
    }

    let variation = (0xfe00..=0xfe0f).contains(&code) || (0xe0100..=0xe01ef).contains(&code);

    (!variation).then_some(character)
}

/// Every code point a format 4 subtable maps, in its own order, while `visit` asks for
/// more. Each is walked through the same arithmetic a lookup uses, so the two cannot
/// disagree about what the font draws.
fn segment_map_codes(bytes: &[u8], offset: usize, visit: &mut impl FnMut(u32) -> bool) {
    let Some(segments) = be_u16(bytes, offset + 6).map(|value| (value / 2) as usize) else {
        return;
    };
    let end_codes = offset + 14;
    let start_codes = end_codes + segments * 2 + 2;

    for index in 0..segments {
        let (Some(start), Some(end)) = (
            be_u16(bytes, start_codes + index * 2),
            be_u16(bytes, end_codes + index * 2),
        ) else {
            return;
        };

        for code in start..=end {
            // The segments end with a sentinel that claims the last code point of the BMP;
            // nothing is drawn for it.
            if code == 0xffff {
                continue;
            }

            match segment_map_glyph(bytes, offset, u32::from(code)) {
                Some(glyph) if glyph != 0 => {}
                _ => continue,
            }

            if !visit(u32::from(code)) {
                return;
            }
        }
    }
}

/// Every code point a format 12 subtable maps, the same way.
fn group_map_codes(bytes: &[u8], offset: usize, visit: &mut impl FnMut(u32) -> bool) {
    let Some(groups) = be_u32(bytes, offset + 12) else {
        return;
    };

    for index in 0..groups as usize {
        let at = offset + 16 + index * 12;
        let (Some(start), Some(end), Some(first)) = (
            be_u32(bytes, at),
            be_u32(bytes, at + 4),
            be_u32(bytes, at + 8),
        ) else {
            return;
        };

        // A group is a run of consecutive code points, and a file is free to claim a run
        // this side will never walk to its end; what it is not free to do is make a hover
        // read forever.
        if end.saturating_sub(start) > 0x1_0000 {
            continue;
        }

        for code in start..=end {
            if first + (code - start) == 0 {
                continue;
            }
            if !visit(code) {
                return;
            }
        }
    }
}

/// Every code point a format 6 subtable maps, the same way.
fn range_map_codes(bytes: &[u8], offset: usize, visit: &mut impl FnMut(u32) -> bool) {
    let Some(first) = be_u16(bytes, offset + 6) else {
        return;
    };
    let Some(count) = be_u16(bytes, offset + 8) else {
        return;
    };

    for index in 0..count {
        if be_u16(bytes, offset + 10 + index as usize * 2) == Some(0) {
            continue;
        }
        if !visit(u32::from(first) + u32::from(index)) {
            return;
        }
    }
}

/// Every code point a format 0 subtable maps, the same way.
fn byte_map_codes(bytes: &[u8], offset: usize, visit: &mut impl FnMut(u32) -> bool) {
    for code in 0..256u32 {
        if byte_map_glyph(bytes, offset, code) == Some(0) {
            continue;
        }
        if !visit(code) {
            return;
        }
    }
}

/// The family and style a font calls itself, out of its `name` table: name 1 and name 2,
/// read from the record that names them best — a Windows English one where the font has it,
/// and whatever else it has where it does not.
fn name_strings(name: &[u8]) -> Option<(String, Option<String>)> {
    let count = be_u16(name, 2)? as usize;
    let storage = be_u16(name, 4)? as usize;

    let mut family: Option<(u32, String)> = None;
    let mut style: Option<(u32, String)> = None;

    for index in 0..count {
        let record = 6 + index * 12;
        let (Some(platform), Some(language), Some(name_id)) = (
            be_u16(name, record),
            be_u16(name, record + 4),
            be_u16(name, record + 6),
        ) else {
            return None;
        };

        if name_id != 1 && name_id != 2 {
            continue;
        }

        let length = be_u16(name, record + 8)? as usize;
        let at = storage + be_u16(name, record + 10)? as usize;
        let Some(bytes) = name.get(at..at + length) else {
            continue;
        };
        let Some(text) = decode_name(platform, bytes) else {
            continue;
        };

        let text = text.trim().to_string();
        if text.is_empty() {
            continue;
        }

        let score = name_record_score(platform, language);
        let held = if name_id == 1 {
            &mut family
        } else {
            &mut style
        };

        if held.as_ref().is_none_or(|(best, _)| score > *best) {
            *held = Some((score, text));
        }
    }

    let (_, family) = family?;

    Some((family, style.map(|(_, style)| style)))
}

/// How good a name record is: the English one a family name is expected in first, then the
/// platform records that are Unicode by construction, then everything else.
fn name_record_score(platform: u16, language: u16) -> u32 {
    match (platform, language) {
        (3, 0x0409) => 5,
        (0, _) => 4,
        (3, _) => 3,
        (1, 0) => 2,
        _ => 1,
    }
}

/// One name record as text. The platforms that carry UTF-16 are decoded as such; a Mac
/// record is read a byte to a character, which is right for the letters a family name is
/// spelled in and wrong only for the accents an old Mac font might wear.
fn decode_name(platform: u16, bytes: &[u8]) -> Option<String> {
    match platform {
        0 | 3 => {
            let units: Vec<u16> = bytes
                .as_chunks::<2>()
                .0
                .iter()
                .map(|pair| u16::from_be_bytes([pair[0], pair[1]]))
                .collect();

            Some(String::from_utf16_lossy(&units))
        }
        1 => Some(bytes.iter().map(|byte| char::from(*byte)).collect()),
        _ => None,
    }
}

/// Where the face `face` of a collection begins, and what its outlines are — which is also
/// what decides what the face is written out as.
///
/// The face is clamped to the last one the file holds, the same way `collection_tables`
/// clamps it, so a file edited between the two reads cannot answer with an offset it has
/// no face at.
fn collection_face(path: &Path, face: usize) -> Option<(usize, [u8; 4])> {
    let mut file = File::open(path).ok()?;
    // The tag, the two version numbers, the face count, and then the face offsets — of
    // which the one the setting asks for is the one a preview is of.
    let mut header = [0u8; 16];
    file.read_exact(&mut header).ok()?;

    if &header[..4] != b"ttcf" {
        return None;
    }

    let faces = u32::from_be_bytes([header[8], header[9], header[10], header[11]]) as usize;
    if faces == 0 {
        return None;
    }

    let mut entry = [0u8; 4];
    file.seek(SeekFrom::Start((12 + face.min(faces - 1) * 4) as u64))
        .ok()?;
    file.read_exact(&mut entry).ok()?;

    let offset = u32::from_be_bytes(entry) as u64;
    file.seek(SeekFrom::Start(offset)).ok()?;

    let mut version = [0u8; 4];
    file.read_exact(&mut version).ok()?;

    Some((offset as usize, version))
}

/// One face of a collection as an sfnt of its own: a header and a table directory of this
/// face's tables, with the tables themselves copied where the new offsets say.
///
/// This is what a `.ttc` has to be turned into to be drawn: the faces of a collection share
/// a table pool and are found by offsets into it, and no page has a syntax for naming one of
/// them. A face's directory is an sfnt's directory, so what comes out is the face with its
/// own tables and nobody else's. The checksums are copied as the face wrote them, which is
/// what a reader that draws a font does not check; the header's own adjustment is left as it
/// was for the same reason.
fn extract_face(bytes: &[u8], face: usize) -> Option<Vec<u8>> {
    let version = bytes.get(face..face + 4)?;
    let count = be_u16(bytes, face + 4)? as usize;

    if count == 0 || count > 512 {
        return None;
    }

    let mut records = Vec::with_capacity(count * 16);
    let mut data = Vec::new();

    for index in 0..count {
        let record = face + 12 + index * 16;
        let tag = bytes.get(record..record + 4)?;
        let checksum = be_u32(bytes, record + 4)?;
        let offset = be_u32(bytes, record + 8)? as usize;
        let length = be_u32(bytes, record + 12)? as usize;
        let table = bytes.get(offset..offset.checked_add(length)?)?;

        // Every table's data is copied where this face's own directory says it is, four
        // bytes apart as an sfnt lays them out — and the tag, the checksum and the length
        // are the face's own, so what a reader sees is the face rather than a rewrite of it.
        records.extend_from_slice(tag);
        records.extend_from_slice(&checksum.to_be_bytes());
        records.extend_from_slice(&((12 + count * 16 + data.len()) as u32).to_be_bytes());
        records.extend_from_slice(&(length as u32).to_be_bytes());

        data.extend_from_slice(table);
        while !data.len().is_multiple_of(4) {
            data.push(0);
        }
    }

    let mut font = Vec::with_capacity(12 + count * 16 + data.len());
    font.extend_from_slice(version);

    // The binary-search fields an sfnt header carries: what a reader that walks the
    // directory by halving it reads, and nothing this app's own readers use.
    let mut power = 1usize;
    let mut selector = 0u16;
    while power * 2 <= count {
        power *= 2;
        selector += 1;
    }

    font.extend_from_slice(&(count as u16).to_be_bytes());
    font.extend_from_slice(&((power * 16) as u16).to_be_bytes());
    font.extend_from_slice(&selector.to_be_bytes());
    font.extend_from_slice(&((count * 16 - power * 16) as u16).to_be_bytes());
    font.extend_from_slice(&records);
    font.extend_from_slice(&data);

    Some(font)
}

/// A short, stable name for a path. The extracted face is named by this and the font's
/// version rather than by the font's own name, which can be any character a file name
/// cannot hold.
fn path_hash(path: &Path) -> u64 {
    let mut hash: u64 = 0xcbf2_9ce4_8422_2325;

    for byte in path.to_string_lossy().to_lowercase().as_bytes() {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }

    hash
}

/// A file's version in milliseconds since the epoch, for the name an extracted face is
/// written under: a font edited in place is a different name, so the browser can never
/// answer a hover with the face as it was.
fn file_stamp(path: &Path) -> u64 {
    std::fs::metadata(path)
        .and_then(|meta| meta.modified())
        .ok()
        .and_then(|modified| modified.duration_since(std::time::UNIX_EPOCH).ok())
        .map(|since| since.as_millis() as u64)
        .unwrap_or(0)
}

fn be_u16(bytes: &[u8], at: usize) -> Option<u16> {
    let value = bytes.get(at..at + 2)?;

    Some(u16::from_be_bytes([value[0], value[1]]))
}

fn be_u32(bytes: &[u8], at: usize) -> Option<u32> {
    let value = bytes.get(at..at + 4)?;

    Some(u32::from_be_bytes([value[0], value[1], value[2], value[3]]))
}

/// How many specimens are held at once.
///
/// What a pointer meets again is the folder it is in, so the handful of fonts under it are
/// the ones worth keeping; one that falls out is read again the next time it is hovered,
/// which is the cost this cache exists to save rather than one it turns into a failure.
const MAX_HELD_SPECIMENS: usize = 32;

/// The file's modification time and length: what says a file is not the one that was read
/// last time.
#[derive(Clone, PartialEq, Eq, Hash)]
struct FileVersion {
    modified: Option<SystemTime>,
    len: u64,
}

/// What a held specimen is valid for: the file, the version of it that was read, and the
/// face of a collection it was read at — since another face of the same file is another
/// specimen, and a setting that names one is answered only by the specimen read at it.
#[derive(Clone, PartialEq, Eq, Hash)]
struct FontKey {
    path: PathBuf,
    version: FileVersion,
    face: usize,
}

/// A held specimen and when it was last asked for. The stamp is a counter rather than a
/// clock, so the order specimens are dropped in cannot be changed by the system clock
/// moving.
struct HeldSpecimen {
    specimen: Option<Specimen>,
    last_used: u64,
}

#[derive(Default)]
struct SpecimenCache {
    entries: HashMap<FontKey, HeldSpecimen>,
    tick: u64,
}

/// The specimens held between hovers. A file that turned out not to be a font is held as
/// that, rather than as nothing held at all, so a hover onto it costs nothing after the
/// first — the same way a document that will not parse is remembered as one.
static SPECIMENS: Lazy<Mutex<SpecimenCache>> = Lazy::new(|| Mutex::new(SpecimenCache::default()));

/// What is held for a version of a file, when anything is.
fn held(key: &FontKey) -> Option<Option<Specimen>> {
    let mut cache = SPECIMENS.lock().ok()?;
    cache.tick += 1;
    let tick = cache.tick;

    let held = cache.entries.get_mut(key)?;
    held.last_used = tick;

    Some(held.specimen.clone())
}

/// Hold what a file turned out to be, dropping the least recently used specimen once the
/// cap is passed.
fn hold(key: &FontKey, specimen: Option<Specimen>) {
    let Ok(mut cache) = SPECIMENS.lock() else {
        return;
    };

    cache.tick += 1;
    let tick = cache.tick;

    cache.entries.insert(
        key.clone(),
        HeldSpecimen {
            specimen,
            last_used: tick,
        },
    );

    while cache.entries.len() > MAX_HELD_SPECIMENS {
        let Some(oldest) = cache
            .entries
            .iter()
            .min_by_key(|(_, held)| held.last_used)
            .map(|(key, _)| key.clone())
        else {
            break;
        };

        cache.entries.remove(&oldest);
    }
}

fn file_version(path: &Path) -> FileVersion {
    match std::fs::metadata(path) {
        Ok(metadata) => FileVersion {
            modified: metadata.modified().ok(),
            len: metadata.len(),
        },
        Err(_) => FileVersion {
            modified: None,
            len: 0,
        },
    }
}

#[cfg(test)]
mod tests {
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
}

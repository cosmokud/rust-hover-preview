use crate::config::config::decode_budget_bytes;
use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::Path;

/// The tables a specimen is read from, out of whichever container the file turned out to
/// be: a font's own table directory, a webfont's compressed one, or one face of a
/// collection.
pub(super) struct Tables {
    /// The character map, as the font stores it.
    pub(super) cmap: Option<Vec<u8>>,
    /// The name table, as the font stores it.
    pub(super) name: Option<Vec<u8>>,
    /// How many faces the file holds: more than one is a collection.
    pub(super) faces: usize,
    /// Which of those faces these tables came out of, as an index into it — `0` for a file
    /// that holds a single font, whatever face the setting named.
    pub(super) face: usize,
}

/// Which container the file is, by its own first bytes rather than by what it is called: a
/// collection, one of the two webfont containers, or an sfnt written out in the open.
///
/// `face` is which face of a collection to read, and is ignored by the containers that hold
/// one font: a `.ttf` is its own face whatever number the setting holds. A file that is
/// none of the containers — a `.ttf` that holds something else, a download that never
/// finished — is answered with nothing, which is what keeps a hover from opening a box
/// nothing would be drawn into.
pub(super) fn tables_of(bytes: &[u8], face: usize) -> Option<Tables> {
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
pub(super) const WOFF2_KNOWN_TAGS: [&[u8; 4]; 63] = [
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
pub(super) struct CharacterMap {
    bytes: Vec<u8>,
    offset: usize,
    format: u16,
    symbol: bool,
}

impl CharacterMap {
    /// The best subtable of a `cmap`, or nothing for a table with none this reads.
    pub(super) fn parse(cmap: &[u8]) -> Option<Self> {
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
    pub(super) fn maps(&self, character: char) -> bool {
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
    pub(super) fn drawable_codes(&self, limit: usize) -> Vec<char> {
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
pub(super) fn name_strings(name: &[u8]) -> Option<(String, Option<String>)> {
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
pub(super) fn collection_face(path: &Path, face: usize) -> Option<(usize, [u8; 4])> {
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
pub(super) fn extract_face(bytes: &[u8], face: usize) -> Option<Vec<u8>> {
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
pub(super) fn path_hash(path: &Path) -> u64 {
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
pub(super) fn file_stamp(path: &Path) -> u64 {
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

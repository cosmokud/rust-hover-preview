//! The question asked of the front of a file for each format the table above names, and
//! the byte and bit readers those questions are written with.
//!
//! Nothing here reads the disk or consults the configuration: every one of them is a
//! question about the same four kilobytes, and what it answers with is a name.

use super::probe::Content;
use crate::config::config::PreviewType;

// ------------------------------------------- the heads of the archives above

/// Whether the front of a file is an ARJ archive: the two bytes the format is defined by, and
/// the length of the header that follows them.
///
/// Two bytes alone are a weak thing to hang a format on — a file that happens to begin with them
/// is not an archive — so the word after them is asked as well: an ARJ header length is a small
/// number, and a file that answers with a large one is not this format.
pub(super) fn is_arj_archive(probe: &[u8]) -> bool {
    starts_with(probe, &[0x60, 0xEA])
        && read_le16(probe, 2).is_some_and(|header| (10..=4_096).contains(&header))
}

/// Whether the front of a file is a cpio archive, in any of the three shapes the format is
/// written in: the two ASCII headers that name their own version, and the older binary one,
/// whose sixteen-bit magic a writer may have written either way round.
pub(super) fn is_cpio_archive(probe: &[u8]) -> bool {
    starts_with(probe, b"070701")
        || starts_with(probe, b"070702")
        || starts_with(probe, b"070707")
        || starts_with(probe, &[0xC7, 0x71])
        || starts_with(probe, &[0x71, 0xC7])
}

/// Whether the front of a file is an LHA archive: the method it was compressed with, four bytes
/// in — every one of them begins `-lh`, whatever level and dictionary size follow — and the level
/// digit that has to come after it.
pub(super) fn is_lha_archive(probe: &[u8]) -> bool {
    at(probe, 2, b"-lh") && probe.get(5).is_some_and(|level| level.is_ascii_digit())
}

/// Whether the front of a file is a SquashFS image: the format's magic, in the four byte orders
/// its writers have written it in.
pub(super) fn is_squashfs(probe: &[u8]) -> bool {
    ["hsqs", "sqsh", "shsq", "qshs"]
        .iter()
        .any(|magic| starts_with(probe, magic.as_bytes()))
}

/// Whether the front of a file is a Macintosh volume: the signature that sits at a kilobyte of
/// the volume rather than at the front of the file, and the version word after it — `H+` for the
/// older format and `HX` for the one that replaced it.
pub(super) fn is_hfs_volume(probe: &[u8]) -> bool {
    at(probe, 1024, &[0x42, 0x44]) && (at(probe, 1026, b"H+") || at(probe, 1026, b"HX"))
}

/// Whether the front of a file is one of the packages that declare their own type: a zip
/// whose first entry is a stored `mimetype` entry holding `mime`.
///
/// The local file header is thirty bytes and the entry is stored rather than deflated with
/// no extra field, so the eight characters of the entry's name sit at 30 and its content
/// at 38 — an offset every file of every format that writes its type this way agrees on.
/// What follows the type is the next record of the archive, which is to say the letter
/// `P`: asking for it is what keeps a subtype from answering for the type whose name it
/// begins with, the way `…graphics` begins `…graphics-template`.
pub(super) fn is_declared_package(probe: &[u8], mime: &[u8]) -> bool {
    at(probe, 0, b"PK\x03\x04")
        // Stored rather than deflated: the format requires it of this entry, and it is what
        // makes the offsets below fixed.
        && at(probe, 8, &[0x00, 0x00])
        && at(probe, 26, &[0x08, 0x00, 0x00, 0x00])
        && at(probe, 30, b"mimetype")
        && at(probe, 38, mime)
        && probe.get(38 + mime.len()) == Some(&b'P')
}

/// Whether the front of a file is a DirectDraw Surface: the four-character code and the
/// size of the header that follows it, which is what a texture's magic is.
pub(super) fn is_dds_surface(probe: &[u8]) -> bool {
    starts_with(probe, b"DDS ") && at(probe, 4, &[0x7C, 0x00, 0x00, 0x00])
}

/// Whether the front of a file is an OpenEXR picture: the magic, which is the number the
/// format's first version was written as.
pub(super) fn is_openexr(probe: &[u8]) -> bool {
    starts_with(probe, &[0x76, 0x2F, 0x31, 0x01])
}

/// Whether the front of a file is a Radiance picture, in the header the format has written
/// since it was called RGBE and under the name it has now.
pub(super) fn is_radiance_picture(probe: &[u8]) -> bool {
    starts_with(probe, b"#?RADIANCE") || starts_with(probe, b"#?RGBE")
}

/// Whether the front of a file is a Netpbm picture.
///
/// `P1` to `P6` are two characters and a separator, which is little enough to be an
/// accident, so the first digit of the size that must follow them is asked for as well.
/// `P7` is a PAM file, whose header names its own fields and ends with a word of its own,
/// and which is asked for those rather than for the two characters.
pub(super) fn is_netpbm(probe: &[u8]) -> bool {
    match probe.get(0..3) {
        Some([b'P', b'1'..=b'6', separator]) if separator.is_ascii_whitespace() => {
            probe.get(3).is_some_and(|byte| byte.is_ascii_digit())
        }
        Some([b'P', b'7', b'\n']) => {
            let header = &probe[..probe.len().min(256)];

            contains(header, b"WIDTH") && contains(header, b"HEIGHT") && contains(header, b"ENDHDR")
        }
        _ => false,
    }
}

/// Whether the front of a file is a Kodak Photo CD picture: the marker two kilobytes into
/// the image pac's padded header, or the one an overview pac opens with.
pub(super) fn is_photo_cd(probe: &[u8]) -> bool {
    at(probe, 2048, b"PCD_IPI") || starts_with(probe, b"PCD_OPA")
}

/// Whether the front of a file is a ZSoft PCX picture: the encoding byte the format is
/// defined by, the version and the encoding it carries, and a palette that is filled in.
pub(super) fn is_pcx(probe: &[u8]) -> bool {
    matches!(probe.first(), Some(0x0A))
        && matches!(probe.get(1), Some(0x00 | 0x02 | 0x03 | 0x04 | 0x05))
        && matches!(probe.get(2), Some(0x00 | 0x01))
        && probe.get(3).is_some_and(|bits| *bits > 0)
        && probe
            .get(8..12)
            .is_some_and(|palette| palette != [0x00, 0x00, 0x00, 0x00])
}

/// Whether the front of a file is a QuickDraw PICT picture: the version operator of the
/// drawing, which an on-disk file carries five hundred and twenty-two bytes in.
pub(super) fn is_quickdraw_pict(probe: &[u8]) -> bool {
    at(probe, 522, &[0x00, 0x11, 0x02, 0xFF]) || at(probe, 522, &[0x11, 0x01])
}

/// Whether the front of a file is a JPEG 2000 picture, in either of the two shapes the
/// format is written in.
///
/// The file format is a box structure, and its first box is the signature box — twelve
/// bytes long, so it announces itself at four rather than at the front of the file. What the
/// code stream is instead is the picture with none of the boxes: the marker every codestream
/// opens with, which is the same thing a JPEG 2000 file's `jp2c` box holds.
pub(super) fn is_jpeg2000(probe: &[u8]) -> bool {
    at(probe, 4, &[b'j', b'P', b' ', b' ', 0x0D, 0x0A, 0x87, 0x0A])
        || starts_with(probe, &[0xFF, 0x4F, 0xFF, 0x51])
}

/// Whether the front of a file is a Flexible Image Transport System picture.
///
/// The header is ASCII card images of eighty bytes each, and the first of them is the record
/// that says a picture starts here. The card that has to follow it is asked for as well: a
/// text file that happens to open with `SIMPLE  =` is not a picture of the sky.
pub(super) fn is_fits(probe: &[u8]) -> bool {
    starts_with(probe, b"SIMPLE  =") && contains(probe, b"BITPIX")
}

/// Whether the front of a file is a Silicon Graphics picture: the two-byte magic, the
/// storage the picture is held in, and the bytes to a channel — in either byte order, since
/// the format is written both ways.
pub(super) fn is_silicon_graphics(probe: &[u8]) -> bool {
    let magic = probe.get(0..2);

    if !matches!(magic, Some([0x01, 0xDA]) | Some([0xDA, 0x01])) {
        return false;
    }

    matches!(probe.get(2), Some(0x00 | 0x01)) && matches!(probe.get(3), Some(0x01 | 0x02))
}

/// Whether the front of a file is a camera raw written as a TIFF with a magic of its own.
///
/// The two formats that do this are Olympus's and Panasonic's: both are the TIFF container —
/// the byte order, the number of the format's first version, and the offset of the first
/// directory — with the format's own answer written where that number would be, so what
/// tells them from a picture is the two pairs of characters in place of it. The other
/// direction round, a TIFF is a picture: the number is the number, and the name decides.
pub(super) fn is_raw_tiff_variant(probe: &[u8]) -> bool {
    matches!(
        probe.get(0..4),
        Some([0x49, 0x49, 0x52, 0x4F])
            | Some([0x4D, 0x4D, 0x4F, 0x52])
            | Some([0x49, 0x49, 0x52, 0x53])
            | Some([0x49, 0x49, 0x55, 0x00])
    )
}

/// Whether the names a file's bytes answered with are the container a camera raw is written
/// in rather than a picture of its own.
///
/// Every raw format from Canon, Nikon, Sony, Pentax, Samsung and Adobe is a TIFF — the
/// reading, the recipe beside it and a JPEG preview of what it comes out as, in a set of
/// IFDs — so the common table names the box and not what is inside it, and what the box
/// holds is a `.nef`, a `.cr2`, an `.arw`, a `.dng` or a `.pef` whose own name says which.
/// The test is those two names as a whole: a `.tif` is a picture like any other, and a name
/// the `[magick]` list carries is one no reader here opens at all.
pub(super) fn names_the_container_of_a_raw(names: &[&str]) -> bool {
    names
        .iter()
        .any(|name| name.eq_ignore_ascii_case("tif") || name.eq_ignore_ascii_case("tiff"))
}

/// Whether the front of a file is a RIFF container of one of Corel's formats, which say
/// which one they are in the form type that follows the size word.
pub(super) fn is_riff_form(probe: &[u8], form: &[u8]) -> bool {
    if !starts_with(probe, b"RIFF") && !starts_with(probe, b"RIFX") {
        return false;
    }

    let Some(kind) = probe.get(8..11) else {
        return false;
    };

    kind.eq_ignore_ascii_case(form)
}

/// Whether the front of a file is an enhanced metafile: the record type every one opens
/// with, the signature forty bytes in, and the version that follows it.
pub(super) fn is_enhanced_metafile(probe: &[u8]) -> bool {
    starts_with(probe, &[0x01, 0x00, 0x00, 0x00])
        && at(probe, 40, b" EMF")
        && at(probe, 44, &[0x00, 0x00, 0x01, 0x00])
        && matches!(read_le32(probe, 4), Some(size) if size >= 88)
}

/// Whether the front of a file is a Windows metafile: the placeable header's own key, or a
/// bare metafile, whose header size word is nine for both of the versions it has.
pub(super) fn is_windows_metafile(probe: &[u8]) -> bool {
    if starts_with(probe, &[0xD7, 0xCD, 0xC6, 0x9A]) {
        return true;
    }

    matches!(probe.first(), Some(0x01 | 0x02)) && at(probe, 2, &[0x09, 0x00])
}

/// Whether the front of a file is a Computer Graphics Metafile, in either of the two
/// encodings the format defines: the binary one, whose first element is BEGIN METAFILE
/// with a class and an identifier of zero and one, or the clear-text one, which spells the
/// same element out.
pub(super) fn is_cgm(probe: &[u8]) -> bool {
    if starts_with(probe, b"BEGMF") {
        return true;
    }

    read_be16(probe, 0).is_some_and(|element| element & 0xFFE0 == 0x0020)
}

/// Whether the front of a file is a ClarisWorks document: the version byte of the format,
/// and the four characters every file the application wrote carries beside it.
pub(super) fn is_clarisworks(probe: &[u8]) -> bool {
    probe
        .first()
        .is_some_and(|version| (1..=6).contains(version))
        && (at(probe, 4, b"BOBO") || at(probe, 4, b"CWKJ"))
}

/// Whether the front of a file is a dBASE table, which carries no magic at all: what is
/// asked is the version byte the format registers, a date in the three bytes after it that
/// is a date, the length of the header the records follow, and the terminator that header
/// ends with.
pub(super) fn is_dbase_table(probe: &[u8]) -> bool {
    const VERSIONS: [u8; 17] = [
        0x02, 0x03, 0x04, 0x05, 0x30, 0x31, 0x32, 0x43, 0x62, 0x7B, 0x83, 0x87, 0x8B, 0x8E, 0xCB,
        0xE5, 0xF4,
    ];

    let (Some(version), Some(month), Some(day)) = (probe.first(), probe.get(2), probe.get(3))
    else {
        return false;
    };

    VERSIONS.contains(version)
        && (1..=12).contains(month)
        && (1..=31).contains(day)
        && matches!(read_le16(probe, 8), Some(length) if length >= 0x21)
        && probe.get(27) == Some(&0x00)
}

/// Whether the front of a file is a DXF drawing: the sentinel a binary file of the format
/// opens with, or a text file whose first group is the section code every one starts with.
pub(super) fn is_dxf(probe: &[u8]) -> bool {
    if starts_with(probe, b"AutoCAD Binary DXF") {
        return true;
    }

    starts_with(probe, b"0\r\nSECTION") || starts_with(probe, b"0\nSECTION")
}

/// Whether the front of a file is a PageMaker document of one version: the tag the format
/// writes six bytes in, and the version word a hundred and ten bytes in.
pub(super) fn is_pagemaker(probe: &[u8], version: &[u8]) -> bool {
    (at(probe, 6, &[0xFF, 0x99]) || at(probe, 6, &[0x99, 0xFF])) && at(probe, 110, version)
}

/// Whether the front of a file is a WordPerfect container holding the kind of content the
/// caller asks for: the signature every file of the family carries, the product byte that
/// follows it, and the file-type byte that says what is inside.
pub(super) fn is_wordperfect(probe: &[u8], file_type: u8) -> bool {
    starts_with(probe, &[0xFF, 0x57, 0x50, 0x43])
        && probe.get(8) == Some(&0x01)
        && probe.get(9) == Some(&file_type)
}

/// Whether the front of a file is an H.265 stream: a start code, then a NAL header whose
/// two bytes say this is a parameter set or a picture of the codec rather than a stream of
/// anything else.
///
/// The header is what all of the H.26x codecs share, so what is asked beyond it is the type
/// the codec defines: a video parameter set, a sequence parameter set, a picture parameter
/// set, an access-unit delimiter or a coded picture. A video parameter set is asked for the
/// field the format reserves to a constant, because that is a signature where the type
/// alone is a shape.
pub(super) fn is_hevc_stream(probe: &[u8]) -> bool {
    let Some(body) = nal_body(probe) else {
        return false;
    };

    let (Some(&first), Some(&second)) = (body.first(), body.get(1)) else {
        return false;
    };

    let kind = (first >> 1) & 0x3F;
    let layer = ((first & 0x01) << 5) | (second >> 3);

    if first & 0x80 != 0 || second & 0x07 == 0 || layer > 62 {
        return false;
    }

    match kind {
        // A video parameter set: the sixteen bits the format reserves to ones.
        32 => at(body, 4, &[0xFF, 0xFF]),
        33..=35 | 39 | 40 | 16..=23 => true,
        _ => false,
    }
}

/// Whether the front of a file is an H.266 stream: the same start code, then a NAL header
/// of the newer codec, whose first byte holds a layer identifier the top two bits of which
/// are reserved to zero and whose second byte holds the unit type.
pub(super) fn is_vvc_stream(probe: &[u8]) -> bool {
    let Some(body) = nal_body(probe) else {
        return false;
    };

    let (Some(&first), Some(&second)) = (body.first(), body.get(1)) else {
        return false;
    };

    first & 0xC0 == 0 && second & 0x07 != 0 && (second >> 3) <= 31
}

/// Whether the front of a file is an AV1 stream: a chain of units whose headers and sizes
/// have to land exactly on one another, beginning with the temporal delimiter or the
/// sequence header a stream of the format opens with.
///
/// The units are not start-code framed and their header byte is only a few bits wide, so a
/// single unit is an accident waiting to happen. What is asked is the chain: each unit's
/// declared size is read and the next unit is required to begin where it says, which is a
/// shape arithmetic-coded or length-prefixed data does not produce by accident.
pub(super) fn is_av1_obu_stream(probe: &[u8]) -> bool {
    let mut offset = 0;
    let mut units = 0;

    while units < 3 {
        let Some(&header) = probe.get(offset) else {
            return false;
        };

        if header & 0x81 != 0 {
            return false;
        }

        let kind = (header >> 3) & 0x0F;
        let extension = (header >> 2) & 0x01 != 0;
        let sized = (header >> 1) & 0x01 != 0;

        if !matches!(kind, 1..=8 | 15) {
            return false;
        }

        if !sized {
            // Without a size field a unit runs to the end of the stream, so only the two
            // short ones can be followed — and a stream begins with one of them.
            return units == 0 && matches!(kind, 1 | 2);
        }

        let mut size = 0usize;
        let mut shift = 0;
        let mut index = offset + 1 + usize::from(extension);

        loop {
            let Some(&byte) = probe.get(index) else {
                return false;
            };

            size |= usize::from(byte & 0x7F) << shift;
            index += 1;

            if byte & 0x80 == 0 {
                break;
            }

            shift += 7;
            if shift > 28 {
                return false;
            }
        }

        // A unit whose own bytes are not in the probe has not been followed, so the chain
        // it would have been is not one: what was read has to be a whole number of them.
        let end = index + size;

        if end > probe.len() {
            return false;
        }

        offset = end;
        units += 1;

        if offset == probe.len() {
            break;
        }
    }

    units >= 2
}

/// Whether the front of a file is an AVS sequence header: the start code the family of
/// codecs opens with — which is the one an MPEG-4 stream uses for a different thing — and
/// the profile, level and picture size that follow it.
///
/// The three generations cannot be told apart: they share the start code, they overlap in
/// the profiles they write, and what separates them is the bit offset of the size fields,
/// which is a question for a decoder. What can be told is that the header is not the one
/// MPEG-4 writes, whose profile byte is a value no AVS profile has.
pub(super) fn is_avs_sequence(probe: &[u8]) -> bool {
    if !at(probe, 0, &[0x00, 0x00, 0x01, 0xB0]) {
        return false;
    }

    let (Some(&profile), Some(&level)) = (probe.get(4), probe.get(5)) else {
        return false;
    };

    if !matches!(
        profile,
        0x10 | 0x12 | 0x20 | 0x22 | 0x42 | 0x48 | 0x50 | 0x62 | 0x66 | 0x74 | 0x80 | 0x82
    ) || level == 0
        || level > 0x88
    {
        return false;
    }

    // The size fields sit after the start code, a profile, a level and a progressive
    // flag: the seventeenth bit after the start code is the first of the width.
    const SIZES_AT: usize = 4 * 8 + 17;

    let (Some(width), Some(height)) = (
        read_bits(probe, SIZES_AT, 14),
        read_bits(probe, SIZES_AT + 14, 14),
    ) else {
        return false;
    };

    (16..=8192).contains(&width) && (16..=8192).contains(&height)
}

/// Whether the front of a file is a raw Dirac or VC-2 stream — one format, two names — told
/// by the parse-info prefix, the parse code the format registers, and, where the unit the
/// header points at is inside the probe, an offset the two units agree about.
pub(super) fn is_dirac_stream(probe: &[u8]) -> bool {
    const PARSE_CODES: [u8; 17] = [
        0x00, 0x10, 0x20, 0x30, 0x08, 0x48, 0xC8, 0xE8, 0x0A, 0x0C, 0x0D, 0x0E, 0x4C, 0x09, 0xCC,
        0x88, 0xCB,
    ];

    if !starts_with(probe, b"BBCD") {
        return false;
    }

    let Some(&parse_code) = probe.get(4) else {
        return false;
    };

    if !PARSE_CODES.contains(&parse_code) {
        return false;
    }

    let Some(next) = read_be32(probe, 5) else {
        return false;
    };

    if next < 13 {
        return false;
    }

    // Where the unit it names is in the probe, the two have to agree about where this one
    // ended; where it is not, the parse code has answered already.
    match probe.get(next as usize..) {
        Some(rest) if rest.len() >= 13 => at(rest, 0, b"BBCD") && read_be32(rest, 9) == Some(next),
        _ => true,
    }
}

/// Whether the front of a file is an EVC stream, which is not start-code framed: a run of
/// units, each a four-byte length and a two-byte header — and it is the run that is asked
/// for, because one length and one header are a shape ordinary data has.
pub(super) fn is_evc_stream(probe: &[u8]) -> bool {
    let mut offset = 0;

    for _ in 0..3 {
        let Some(length) = read_be32(probe, offset) else {
            return false;
        };

        if length < 2 {
            return false;
        }

        let start = offset + 4;
        let end = start + length as usize;

        let (Some(&header), true) = (probe.get(start), end <= probe.len()) else {
            return false;
        };

        let kind = (header >> 1) & 0x3F;

        if header & 0x80 != 0 || kind == 0 || kind == 63 {
            return false;
        }

        offset = end;
    }

    true
}

/// Whether the front of a file is a DV stream: the header every DIF block of one opens
/// with, looked for in the first kilobyte, which is a frame's worth of blocks and then
/// some.
pub(super) fn is_dv_stream(probe: &[u8]) -> bool {
    (0..probe.len().saturating_sub(3))
        .take(1200)
        .any(|offset| at(probe, offset, &[0x1F, 0x07, 0x00, 0x3F]))
}

/// Whether the front of a file is a VC-1 test stream: the byte the format writes four bytes
/// in, and the marker that closes the first frame's header.
pub(super) fn is_vc1_test_stream(probe: &[u8]) -> bool {
    if probe.get(3) != Some(&0xC5) {
        return false;
    }

    let Some(size) = read_le32(probe, 4) else {
        return false;
    };

    at(probe, size as usize + 16, &[0x0C, 0x00, 0x00, 0x00])
}

/// Whether the front of a file is a TiVo stream: the number the format writes at the head
/// of every chunk, and the chunk header that follows it.
pub(super) fn is_yop(probe: &[u8]) -> bool {
    let (Some(&first), Some(&second), Some(&third), Some(&fourth)) =
        (probe.first(), probe.get(1), probe.get(2), probe.get(3))
    else {
        return false;
    };

    if first != b'Y' || second != b'O' || third >= 10 || fourth >= 10 {
        return false;
    }

    probe.get(6) != Some(&0)
        && probe.get(7) != Some(&0)
        && probe.get(8).is_some_and(|byte| byte & 1 == 0)
        && probe.get(10).is_some_and(|byte| byte & 1 == 0)
        && read_le16(probe, 18).is_some_and(|size| size >= 920)
}

/// Whether the front of a file is a GameCube RSD stream: the four characters it opens with,
/// and the two sizes it carries that have bounds a file of one is written inside.
pub(super) fn is_rsd(probe: &[u8]) -> bool {
    if !starts_with(probe, b"RSD") {
        return false;
    }

    matches!(probe.get(3), Some(b'2'..=b'6'))
        && read_le32(probe, 8).is_some_and(|count| (1..=256).contains(&count))
        && read_le32(probe, 16).is_some_and(|rate| (1..=384_000).contains(&rate))
}

/// Whether the front of a file is an Interplay C93 movie: the block table the format opens
/// with, whose records name each other by the length of the one before, which is the shape
/// FFmpeg's probe scores as this format.
pub(super) fn is_c93(probe: &[u8]) -> bool {
    let mut index = 1;

    for record in 0..4 {
        let offset = record * 4;

        let (Some(number), Some(length), Some(frames)) = (
            read_le16(probe, offset),
            probe.get(offset + 2),
            probe.get(offset + 3),
        ) else {
            return false;
        };

        if number != index || *length == 0 || *frames == 0 {
            return false;
        }

        index = index.wrapping_add(u16::from(*length));
    }

    true
}

/// Whether the front of a file is a CD Graphics stream: twenty-four byte packets whose
/// command byte is the format's graphics command, or the zero a packet the format does not
/// define carries — over as many packets as the probe holds.
pub(super) fn is_cdg(probe: &[u8]) -> bool {
    const PACKET: usize = 24;
    const EXAMINED: usize = 64;

    let packets = (probe.len() / PACKET).min(EXAMINED);

    if packets < 8 {
        return false;
    }

    let mut commands = 0;

    for packet in 0..packets {
        let Some(byte) = probe.get(packet * PACKET) else {
            return false;
        };

        match byte & 0x3F {
            0x09 => commands += 1,
            0x00 => {}
            _ => return false,
        }
    }

    commands * 4 >= packets * 3
}

/// Whether the front of a file is a Commodore CDXL stream: the header's own fields, which
/// is what FFmpeg's probe reads, because the format has no magic to read.
pub(super) fn is_cdxl(probe: &[u8]) -> bool {
    if probe.len() < 32 {
        return false;
    }

    let (Some(&kind), Some(&planes), Some(&reserved)) =
        (probe.first(), probe.get(19), probe.get(18))
    else {
        return false;
    };

    if kind > 1 || !matches!(planes, 6 | 8 | 24) || reserved != 0 || !at(probe, 29, &[0, 0, 0]) {
        return false;
    }

    let (Some(palette), Some(audio), Some(rate)) = (
        read_be16(probe, 20),
        read_be16(probe, 22),
        read_be16(probe, 24),
    ) else {
        return false;
    };

    if palette == 0
        || (kind == 1 && palette > 512)
        || (kind == 0 && palette > 768)
        || (audio == 0 && rate != 0)
        || (kind == 0 && (probe.get(26) == Some(&0) || rate == 0))
    {
        return false;
    }

    let (Some(width), Some(height)) = (read_be16(probe, 14), read_be16(probe, 16)) else {
        return false;
    };

    if width == 0 || width > 640 || height == 0 || height > 480 {
        return false;
    }

    let (Some(size), Some(&flags)) = (read_be32(probe, 2), probe.get(1)) else {
        return false;
    };

    let channels = 1 + u32::from(flags & 0x10 != 0);

    size > u32::from(palette) + u32::from(audio) * channels + 32
}

/// Whether the front of a file is a Moflex movie: the two characters the format opens with,
/// and the record table that follows, whose records have to chain to a pair of zeros.
pub(super) fn is_moflex(probe: &[u8]) -> bool {
    if read_be16(probe, 0) != Some(0x4C32) {
        return false;
    }

    if read_be16(probe, 12).is_none_or(|field| field == 0) {
        return false;
    }

    let mut offset = 14;

    while let (Some(kind), Some(size)) = (read_be16(probe, offset), read_be16(probe, offset + 2)) {
        if kind == 0 && size == 0 {
            return true;
        }

        if size == 0 {
            return false;
        }

        offset += 4 + usize::from(size);
    }

    false
}

/// Whether the front of a file is an H.261 picture.
///
/// The picture start code of the older codecs in the H.26x family is short — twenty bits,
/// the last of which say nothing — so the code on its own is a shape ordinary data can
/// have. What is asked beside it is the code again later in the probe, which is what a
/// stream of more than one picture carries before each of the rest.
pub(super) fn is_h261_picture(probe: &[u8]) -> bool {
    let start = |probe: &[u8], offset: usize| {
        at(probe, offset, &[0x00, 0x01])
            && probe.get(offset + 2).is_some_and(|byte| byte & 0xF0 == 0)
    };

    if !start(probe, 0) {
        return false;
    }

    (3..probe.len().saturating_sub(2)).any(|offset| start(probe, offset))
}

/// Whether the front of a file is an H.263 picture, whose start code is twenty-two bits
/// and asks for the same second one.
pub(super) fn is_h263_picture(probe: &[u8]) -> bool {
    let start = |probe: &[u8], offset: usize| {
        at(probe, offset, &[0x00, 0x00])
            && probe
                .get(offset + 2)
                .is_some_and(|byte| byte & 0xFC == 0x80)
    };

    if !start(probe, 0) {
        return false;
    }

    (3..probe.len().saturating_sub(2)).any(|offset| start(probe, offset))
}

/// The bytes that follow the first start code a probe opens with, which is where a raw
/// stream's own header begins.
pub(super) fn nal_body(probe: &[u8]) -> Option<&[u8]> {
    let start = probe.iter().position(|byte| *byte != 0)?;

    if start < 2 || probe.get(start) != Some(&1) {
        return None;
    }

    probe.get(start + 1..)
}

/// Whether the front of a file is an H.264 elementary stream: a start code, then the
/// sequence parameter set a stream of one opens with.
pub(super) fn is_h264_stream(probe: &[u8]) -> bool {
    nal_body(probe).is_some_and(|body| body.first().is_some_and(|header| header & 0x1F == 7))
}

/// `count` bits of `probe` beginning `start` bits into it, most significant bit first.
pub(super) fn read_bits(probe: &[u8], start: usize, count: usize) -> Option<u32> {
    if start + count > probe.len() * 8 {
        return None;
    }

    let mut value = 0;

    for bit in start..start + count {
        let byte = probe[bit / 8];
        value = (value << 1) | u32::from((byte >> (7 - bit % 8)) & 1);
    }

    Some(value)
}

/// The two bytes at `offset` as a number, and nothing where the probe is too short.
pub(super) fn read_be16(probe: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_be_bytes(
        probe.get(offset..offset + 2)?.try_into().ok()?,
    ))
}

/// And the little-endian one.
pub(super) fn read_le16(probe: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_le_bytes(
        probe.get(offset..offset + 2)?.try_into().ok()?,
    ))
}

/// The four bytes at `offset` as a number, and nothing where the probe is too short.
pub(super) fn read_be32(probe: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_be_bytes(
        probe.get(offset..offset + 4)?.try_into().ok()?,
    ))
}

/// And the little-endian one.
pub(super) fn read_le32(probe: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(
        probe.get(offset..offset + 4)?.try_into().ok()?,
    ))
}

/// Whether `probe` holds `needle` at `offset`, and nothing where it is too short to tell.
pub(super) fn at(probe: &[u8], offset: usize, needle: &[u8]) -> bool {
    probe
        .get(offset..offset + needle.len())
        .is_some_and(|window| window == needle)
}

/// Whether `probe` opens with `prefix`.
pub(super) fn starts_with(probe: &[u8], prefix: &[u8]) -> bool {
    at(probe, 0, prefix)
}

/// Whether `probe` holds `needle` anywhere in it.
pub(super) fn contains(probe: &[u8], needle: &[u8]) -> bool {
    probe.windows(needle.len()).any(|window| window == needle)
}

/// Whether the front of a file is an Ogg page carrying a sound: the container's own four bytes
/// and the header one of the audio codecs writes inside it.
///
/// What is asked is the codec's own name rather than a guess at the framing, because an Ogg
/// file's pages are multiplexed and what the first page holds is the first stream's header —
/// which is why the video entry above is asked before this one: a Theora film's first page
/// carries Theora's header and not the Vorbis header of its audio track.
pub(super) fn is_ogg_sound(probe: &[u8]) -> bool {
    starts_with(probe, b"OggS")
        && (contains(probe, b"vorbis")
            || contains(probe, b"OpusHead")
            || contains(probe, b"Speex")
            || contains(probe, b"fLaC"))
}

/// What a `.pdb` is, which is the one name in the table that holds two formats.
///
/// The name is written for two things that share nothing. One is the Palm OS database —
/// the AportisDoc and its kin, the ebooks the render engine's own filters read, whose
/// header names the application that wrote them and the records that follow. The other is
/// the Microsoft program database a compiler writes beside its binaries, which is an *MSF*
/// container of debug information and no kind of document at all — and which is the one a
/// developer's folders are full of, and the reason this name is not answered by itself.
///
/// What comes back is the ebook where the file's own header is a Palm OS one, and nothing at
/// all where it is anything else, the program database included. Nothing is left to the
/// name here, deliberately: a `.pdb` that is neither of the two is a file this app has no
/// reader for either way, and the engine is not asked about one.
///
/// A `.pdb` that is a *Mobipocket* book is not one of those, and does not arrive here at all:
/// the two identifiers that tell one — the type `BOOK` and the creator `MOBI` — are a signature
/// of their own above, which is answered before any name is asked about (see `super::signatures::SIGNATURES`).
/// What the signature means is that such a file is the ebook engine's rather than the render
/// engine's, which is right by the file rather than by the name it happens to have been given.
pub(super) fn palm_ebook_or_program_database(probe: &[u8]) -> Content {
    if is_program_database(probe) {
        return Content::Foreign;
    }

    if is_palm_database(probe) {
        Content::Kind(PreviewType::Libre)
    } else {
        Content::Foreign
    }
}

/// Whether the front of a file is a Microsoft program database: the *MSF* container a
/// compiler writes, in either of the two versions the format has been written in — the
/// name of the format opens both, and the bytes after it are the container's own.
pub(super) fn is_program_database(probe: &[u8]) -> bool {
    starts_with(probe, b"Microsoft C/C++ MSF 7.00")
        || starts_with(probe, b"Microsoft C/C++ program database 2.00")
}

/// Whether the front of a file is the database header a Palm OS document opens with.
///
/// There is no signature to ask: the format is a name, the four-character type and creator
/// of the application that wrote it, the dates and identifiers of the database, and the
/// record list that follows — a shape a great many files could be written in. What is
/// asked instead is that the two fields the format is *defined* by hold what a Palm OS
/// application writes there: four printable characters each, opening with a letter, which
/// is what `TEXt` (the AportisDoc and the readers beside it), `BOOK` (a MobiPocket one),
/// `DATA` (a Plucker one) and every identifier a Palm program is registered under have in
/// common — and that the database declares at least one record, which is the file's own
/// account of having something in it.
///
/// Which of those applications the engine can read is the engine's business rather than
/// this one's: what is asked here is only whether the file is a Palm document at all.
pub(super) fn is_palm_database(probe: &[u8]) -> bool {
    // The fixed part of the header: the name and the fields up to the record count.
    const HEADER_BYTES: usize = 78;

    let Some(header) = probe.get(..HEADER_BYTES) else {
        return false;
    };

    let records = u16::from_be_bytes([header[76], header[77]]);

    records > 0 && is_palm_tag(&header[60..64]) && is_palm_tag(&header[64..68])
}

/// Whether four bytes hold one of the two identifiers a Palm database is described by:
/// printable ASCII, and opening with a letter.
pub(super) fn is_palm_tag(tag: &[u8]) -> bool {
    tag.first().is_some_and(|byte| byte.is_ascii_alphabetic())
        && tag.iter().all(|byte| byte.is_ascii_graphic())
}

/// Whether the front of a file is a Mobipocket book — the header every Kindle ebook is, under the
/// five names the ebook engine reads it by.
///
/// It is the same Palm database `is_palm_database` reads, asked about one pair of identifiers
/// rather than about the shape of the two fields: the type `BOOK` and the creator `MOBI`, which is
/// what a Mobipocket file, an `.azw`, a KF8 `.azw3`, an `.azw4` and a `.prc` all carry, and which
/// is what tells one from the Palm ebook the render engine's own filters read — the same header
/// with the identifiers of another application. The record count is asked for beside them, the way
/// it is there: a file whose header declares nothing in it is not a book.
pub(super) fn is_mobipocket(probe: &[u8]) -> bool {
    at(probe, 60, b"BOOK")
        && at(probe, 64, b"MOBI")
        && read_be16(probe, 76).is_some_and(|records| records > 0)
}

/// Whether the front of a file is a FictionBook: the root element every file of the format is
/// written with, behind the declaration every XML file opens with.
///
/// Both halves are asked because either alone is no answer: `<?xml` is every XML file there is, and
/// a root element of this name is what no other document format has. The byte-order mark a file
/// may be written with is skipped, since the declaration is what has to follow it.
pub(super) fn is_fictionbook(probe: &[u8]) -> bool {
    let body = probe.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(probe);

    starts_with(body, b"<?xml") && contains(body, b"<FictionBook")
}

/// Whether the front of a file is a DjVu document: the chunk every file of the format opens with,
/// and the form type that follows it — a document of several pages, a single page, or a page
/// included by another.
pub(super) fn is_djvu(probe: &[u8]) -> bool {
    starts_with(probe, b"AT&TFORM")
        && matches!(
            probe.get(8..12),
            Some([b'D', b'J', b'V', b'M' | b'U' | b'I'])
        )
}

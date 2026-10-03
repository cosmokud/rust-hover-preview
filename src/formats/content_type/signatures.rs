//! What this app can name from the front of a file where the `infer` crate cannot: the
//! names each format is known by to the lists, and how the bytes are asked which one
//! it is.
//!
//! One table and the one enum that says what a test is. Every question the entries
//! point at is written out in [`super::matchers`], which is where a table this size
//! reads from rather than holding.

use super::matchers::*;

/// One format this app can name from the front of a file.
pub(super) struct Signature {
    /// The names the format is known by to the lists, which is what a kind is found from.
    ///
    /// A format is written under more than one name — a Matroska file is an `mkv` and a
    /// `webm`, a transport stream is a `ts` and a `m2ts` — and all of them are listed
    /// because the answer is not "what is this" but "is this what the file is called": a
    /// name the file already carries is the two agreeing, and a name no list claims is a
    /// kind this app does not preview.
    pub(super) names: &'static [&'static str],
    /// How the front of a file is read as this format.
    matches: Matcher,
}

/// How one of the formats below is recognized.
enum Matcher {
    /// A question asked of the bytes, written for the format the way the rest of this app
    /// asks one: a magic, a header, the shape a file's own first kilobyte has.
    ///
    /// Where a test is more than a literal it is named after the format and sits below
    /// the table, because what such a test is worth is a judgement about the format rather
    /// than about the bytes.
    Test(fn(&[u8]) -> bool),
    /// A package that declares its own type in the head the way an OpenDocument, a
    /// StarOffice XML document and a Krita project do. The declaration is the first entry
    /// of the archive and is stored rather than deflated, so it sits at an offset every
    /// such file agrees on — and what it holds is the document's own answer rather than
    /// the box's, which is what lets a package be named without breaking the rule that a
    /// container is not a kind.
    Package(&'static [u8]),
}

impl Signature {
    /// Whether the front of a file is this format.
    pub(super) fn matches(&self, probe: &[u8]) -> bool {
        match self.matches {
            Matcher::Test(test) => test(probe),
            Matcher::Package(mime) => is_declared_package(probe, mime),
        }
    }
}

/// The formats the common table does not carry, which are the ones an engine here reads
/// and no file manager would name.
///
/// Every name the engines' lists carry that `infer` has no signature for is answered
/// here, so that the content of a file decides what it is whichever engine the file
/// belongs to: a picture this app's own decoder reads, a drawing the drawing layer plays,
/// a document the render engine imports, a book the ebook engine reads, and the containers
/// and raw streams FFmpeg plays that no common table names. What a format's test is worth
/// differs, and where it is not a literal the test says what it is: the magic of a format
/// whose header is its own, a start code followed by the header a stream of that codec
/// opens with, or — for the packages that declare a type — the declaration itself.
///
/// Three kinds of name are deliberately *not* here, and each is a judgement about the
/// format rather than a gap:
///
/// * **The containers that say nothing about themselves.** A plain zip, an OLE compound
///   file, and a gzip stream are one answer whatever document they hold — an iWork
///   document, a `.doc`, a `.gnumeric` — so nothing is asked of those bytes, and the name
///   decides. The exception is the package that declares its type beside its own name, at
///   an offset every such file agrees on: that declaration is a fact about the document and
///   is read (see [`Matcher::Package`]).
/// * **The formats whose signature is their own text** — an SVG document, a flat
///   OpenDocument, a Visio `.vdx` — which the text lists answer for, and which a prefix
///   match against ordinary text would claim from them. A file whose *format* is text but
///   whose first line is the format's own — a SYLK identifier record, a DXF section code —
///   is the one case that is asked for, because no text list claims those names.
/// * **The few names no head can tell:** a `.tga`, whose only marker is a footer at the
///   end of the file; a `.flm`, whose magic is thirty-six bytes before its end; the names
///   `mvi` and `mxg` carry, which no FFmpeg probe reads at all; `psp` and `vw`, which no
///   FFmpeg demuxer registers; and `cin`, whose demuxer was removed. Each of those keeps
///   the answer it had before this table existed: the name.
pub(super) const SIGNATURES: &[Signature] = &[
    // ------------------------------------------------------------------------- pictures
    // A DirectDraw Surface: the four-character code and the size of the header that
    // follows it, which is what tells a texture from a file that begins with a word.
    Signature {
        names: &["dds", "dxt1", "dxt5"],
        matches: Matcher::Test(is_dds_surface),
    },
    // OpenEXR, whose magic is a version number written as a number.
    Signature {
        names: &["exr"],
        matches: Matcher::Test(is_openexr),
    },
    // A Radiance picture, and the older spelling the format still writes.
    Signature {
        names: &["hdr"],
        matches: Matcher::Test(is_radiance_picture),
    },
    Signature {
        names: &["ff"],
        matches: Matcher::Test(|probe| starts_with(probe, b"farbfeld")),
    },
    Signature {
        names: &["qoi"],
        matches: Matcher::Test(|probe| starts_with(probe, b"qoif")),
    },
    // The Netpbm family in one entry: its six magics are two characters wide, which is
    // weak on its own, so what is asked is the magic, a separator and the first digit of
    // the size that follows — and for a PAM file the named header the format writes.
    Signature {
        names: &["pam", "pbm", "pgm", "pnm", "ppm"],
        matches: Matcher::Test(is_netpbm),
    },
    // A Kodak Photo CD image pac, whose marker sits two kilobytes into a padded header,
    // and the overview pac, which opens with the other one.
    Signature {
        names: &["pcd", "pcds"],
        matches: Matcher::Test(is_photo_cd),
    },
    Signature {
        names: &["pcx"],
        matches: Matcher::Test(is_pcx),
    },
    Signature {
        names: &["ras", "sun"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0x59, 0xA6, 0x6A, 0x95])),
    },
    // A QuickDraw PICT on disk carries a five-hundred-byte header, so the version operator
    // that opens the picture is at 522 rather than at the front of the file.
    Signature {
        names: &["pct", "pict"],
        matches: Matcher::Test(is_quickdraw_pict),
    },
    // ------------------------------------------------- pictures an engine develops
    // The formats the ImageMagick engine reads that nothing else on a machine opens, and
    // that this app's own tables do not carry. What is asked of each is what the format
    // writes at the front of a file and nothing else, the way every entry above is asked —
    // and a format whose head says nothing, or says something another format says as well,
    // is answered by its name instead, which is what `magick_formats` is for.
    //
    // A camera raw is not in this block: the ones that are TIFFs are answered by the name
    // they carry (see `names_the_container_of_a_raw`), and the ones that are not — the
    // Olympus, Panasonic, Fuji and Sigma containers — are entries of their own below.
    //
    // Nor are the four names whose head is nothing to ask about. A `.xbm` and an `.xpm` are
    // C source — the text lists would have them if a user wanted them read as text — and a
    // `.wbmp` is four bytes of type, header, width and height that any little picture could
    // write. A `.cur` is the one worth saying out loud: its header is the icon format's with
    // a two where the icon writes a one, which is *also* the six bytes a Lotus 1-2-3
    // spreadsheet opens with — so a signature for it would take a `.wk1` away from the render
    // engine, which is a trade a picture nobody hovers is not worth.
    //
    // The JPEG 2000 family, in both of the shapes the format is written in: the file format,
    // whose first box is the signature box four bytes into the file, and the bare code
    // stream, which is the same picture with none of the boxes around it.
    Signature {
        names: &["j2c", "j2k", "jp2", "jpc", "jpm", "jpt"],
        matches: Matcher::Test(is_jpeg2000),
    },
    // The two halves of the PNG family's animated kin: a JPEG Network Graphic, which is a
    // PNG carrying JPEG data, and a Multiple-image Network Graphic, which is several PNGs in
    // one stream. Each opens with the eight bytes every PNG-family file opens with, with its
    // own three letters after the first.
    Signature {
        names: &["jng"],
        matches: Matcher::Test(|probe| {
            starts_with(probe, &[0x8B, b'J', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A])
        }),
    },
    Signature {
        names: &["mng"],
        matches: Matcher::Test(|probe| {
            starts_with(probe, &[0x8A, b'M', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A])
        }),
    },
    // A GIMP document, which opens with the program's own name and the version letter the
    // format is written under.
    Signature {
        names: &["xcf"],
        matches: Matcher::Test(|probe| starts_with(probe, b"gimp xcf ")),
    },
    // ImageMagick's own interchange format, which names itself in the header it opens with —
    // the one format here that is the engine's rather than a camera's or an application's,
    // and the one whose magic is a word rather than a number.
    Signature {
        names: &["miff"],
        matches: Matcher::Test(|probe| starts_with(probe, b"id=ImageMagick")),
    },
    // A JPEG Network Graphic's cousin in the film world: the Kodak/SMPTE frame, in either of
    // the two byte orders the format is written in.
    Signature {
        names: &["dpx"],
        matches: Matcher::Test(|probe| starts_with(probe, b"SDPX") || starts_with(probe, b"XPDS")),
    },
    // The astronomer's picture, whose header is an ASCII card image: the record it opens
    // with, and the card a picture of which no other format writes.
    Signature {
        names: &["fit", "fits", "fts"],
        matches: Matcher::Test(is_fits),
    },
    // The film compositor's picture: a two-byte magic, the storage the samples are held in,
    // and the bytes to a channel — one of each of the pairs the format defines.
    Signature {
        names: &["sgi"],
        matches: Matcher::Test(is_silicon_graphics),
    },
    // The medical scanner's image: a hundred and twenty-eight bytes of preamble that may be
    // anything at all, and the four characters the format puts after it.
    Signature {
        names: &["dcm"],
        matches: Matcher::Test(|probe| at(probe, 128, b"DICM")),
    },
    // A multi-page Paintbrush picture: the 0x3ADE68B1 the format is defined by, written
    // little-endian.
    Signature {
        names: &["dcx"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0xB1, 0x68, 0xDE, 0x3A])),
    },
    // A portable float map, which is a Netpbm header with a scale rather than a maximum
    // after the size: the two characters, the separator, and the digit that has to follow.
    Signature {
        names: &["pfm"],
        matches: Matcher::Test(|probe| {
            matches!(probe.get(0..3), Some([b'P', b'F' | b'f', b'\n']))
                && probe.get(3).is_some_and(|byte| byte.is_ascii_digit())
        }),
    },
    // A planetary image from JPL: the header every one opens with names its own length.
    Signature {
        names: &["vicar"],
        matches: Matcher::Test(|probe| starts_with(probe, b"LBLSIZE=")),
    },
    // The camera raws that are not TIFFs: the Olympus and Panasonic containers, which are a
    // TIFF header with a magic of their own in place of the format's, and the Fuji, Sigma and
    // old Canon containers, which say what they are in their first bytes.
    Signature {
        names: &["orf", "rw2"],
        matches: Matcher::Test(is_raw_tiff_variant),
    },
    Signature {
        names: &["raf"],
        matches: Matcher::Test(|probe| starts_with(probe, b"FUJIFILMCCD-RAW")),
    },
    Signature {
        names: &["x3f"],
        matches: Matcher::Test(|probe| starts_with(probe, b"FOVb")),
    },
    Signature {
        names: &["crw"],
        matches: Matcher::Test(|probe| at(probe, 6, b"HEAPCCDR")),
    },
    // ---------------------------------------------------- archives an engine lists
    // The formats the PeaZip engine reads that nothing else on a machine opens, and that this
    // app has no reader of its own for — see `peazip_formats`. Each is asked what the format
    // writes at the front of a file and nothing else, the way every entry above is asked; a
    // format whose head says nothing, or whose marker sits past what is read, is answered by
    // its name instead, which is what that list is for.
    //
    // Three of them are deliberately *not* here, and each is worth saying out loud:
    //
    // * **A gzip stream.** It opens with `1F 8B` like every other, and a signature naming `gz`
    //   would claim a `.tgz` for this kind: the content of a tarball *is* a gzip stream, so the
    //   names a signature answers with are the only thing consulted, and a `.tgz` would stop
    //   being the archive this app reads itself and become a file waiting on an engine. The
    //   whole of that format is left to the name.
    // * **A disk image's.** An iso and a udf say what they are thirty-two kilobytes into the
    //   file — `CD001` at 32769, a UDF descriptor at 32768 — and what a hover reads of a file is
    //   its first four kilobytes, so neither marker is reachable from here. What is left to the
    //   name is the whole of those two formats (see TODO.md, where the bound is written down).
    // * **And a `.lzma`, and the parts of a split archive.** Neither has a marker at all: an
    //   LZMA stream is the compressed data with nothing in front of it, and `.001` is the first
    //   slice of a file rather than a format. Both are the name's business, as every format
    //   without a head is.
    //
    // A Unix archive — an `ar`, and the `.deb` and `.udeb` packages that are one — which opens
    // with the eight characters every one of them begins with.
    Signature {
        names: &["ar", "deb", "udeb"],
        matches: Matcher::Test(|probe| starts_with(probe, b"!<arch>\n")),
    },
    // An ARJ archive: the two bytes the format is defined by, and the header length that
    // follows them, which is what keeps a file that merely begins with those two bytes from
    // answering as one.
    Signature {
        names: &["arj"],
        matches: Matcher::Test(is_arj_archive),
    },
    // A cabinet file: its four characters and the four zero bytes a cabinet's own header
    // reserves.
    Signature {
        names: &["cab"],
        matches: Matcher::Test(|probe| starts_with(probe, b"MSCF\x00\x00\x00\x00")),
    },
    // A compiled help file: the two words the format opens with, the second of which is the
    // offset its directory sits at — a constant of the format rather than a number that varies.
    Signature {
        names: &["chm"],
        matches: Matcher::Test(|probe| {
            starts_with(probe, b"ITSF\x03\x00\x00\x00") && at(probe, 8, &[0x60, 0x00, 0x00, 0x00])
        }),
    },
    // A cpio archive, in the three shapes the format is written in: the two ASCII headers
    // that name their own version, and the older binary one, which is the same number written
    // either way round.
    Signature {
        names: &["cpio"],
        matches: Matcher::Test(is_cpio_archive),
    },
    // An LHA archive, which says what it is four bytes in: a dashes-and-letters method, the
    // `-lh` every one of them carries, and the level digit after it.
    Signature {
        names: &["lha", "lzh"],
        matches: Matcher::Test(is_lha_archive),
    },
    // A Linux package: the four bytes every RPM opens with.
    Signature {
        names: &["rpm"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0xED, 0xAB, 0xEE, 0xDB])),
    },
    // Apple's archive: a `pkg` installer, a `xip` of a signed one, and the `xar` itself.
    Signature {
        names: &["pkg", "xar", "xip"],
        matches: Matcher::Test(|probe| starts_with(probe, b"xar!")),
    },
    // A Windows image, which the installer's `.esd` is a second shape of.
    Signature {
        names: &["esd", "swm", "wim"],
        matches: Matcher::Test(|probe| starts_with(probe, b"MSWIM\x00\x00\x00")),
    },
    // A SquashFS image, in the four byte orders the format writes its magic in.
    Signature {
        names: &["squashfs"],
        matches: Matcher::Test(is_squashfs),
    },
    // The single-stream compressors that name themselves at the front: a bzip2 stream opens with
    // the three characters every one of them carries and the block size after them, an xz stream
    // with its own six, and a `.z` with two control bytes that no text file and no other format
    // opens with. What is left to the name is what has nothing to ask — a `.lzma`, whose stream is
    // the compressed data and nothing in front of it, and a gzip stream, which is the one name
    // here that another list already reads by a longer spelling of it (see the note above).
    Signature {
        names: &["bzip2", "bz2"],
        matches: Matcher::Test(|probe| {
            starts_with(probe, b"BZh")
                && probe
                    .get(3)
                    .is_some_and(|level| (b'1'..=b'9').contains(level))
        }),
    },
    Signature {
        names: &["xz"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0xFD, b'7', b'z', b'X', b'Z', 0x00])),
    },
    Signature {
        names: &["z"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0x1F, 0x9D])),
    },
    // A zstd stream: the frame magic written little-endian.
    Signature {
        names: &["zst"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0x28, 0xB5, 0x2F, 0xFD])),
    },
    // The four disk images that name themselves at the front: a macOS image, a Virtual PC one,
    // a Hyper-V one, and VMware's — whose descriptor is a text file and whose split extents
    // carry this.
    Signature {
        names: &["dmg"],
        matches: Matcher::Test(|probe| {
            starts_with(probe, b"koly") && at(probe, 4, &[0x00, 0x00, 0x00, 0x04])
        }),
    },
    Signature {
        names: &["vhd"],
        matches: Matcher::Test(|probe| starts_with(probe, b"conectix")),
    },
    Signature {
        names: &["vhdx"],
        matches: Matcher::Test(|probe| starts_with(probe, b"vhdxfile")),
    },
    Signature {
        names: &["vmdk"],
        matches: Matcher::Test(|probe| starts_with(probe, b"KDMV")),
    },
    // A QEMU image: the format's four bytes and the version that follows them.
    Signature {
        names: &["qcow", "qcow2"],
        matches: Matcher::Test(|probe| starts_with(probe, b"QFI\xFB")),
    },
    // A VirtualBox image, whose magic sits sixty-four bytes in rather than at the front.
    Signature {
        names: &["vdi"],
        matches: Matcher::Test(|probe| at(probe, 64, &[0x7F, 0x10, 0xDA, 0xBE])),
    },
    // An Apple file system image, whose container header names itself thirty-two bytes in.
    Signature {
        names: &["apfs"],
        matches: Matcher::Test(|probe| at(probe, 32, b"NXSB")),
    },
    // A Macintosh volume, in either of the two shapes it was written in: the signature at a
    // kilobyte, and the version that follows it.
    Signature {
        names: &["hfs", "hfsx"],
        matches: Matcher::Test(is_hfs_volume),
    },
    // The compressed file system of an embedded device, whose older header says what it is in
    // words rather than in a number.
    Signature {
        names: &["cramfs"],
        matches: Matcher::Test(|probe| at(probe, 16, b"Compressed ROMFS")),
    },
    // ------------------------------------------------------------------------ drawings
    // Windows' two metafiles. Neither is a picture this app decodes: both are replayed by
    // the drawing layer, and what tells them apart is the second word — an enhanced
    // metafile counts from zero there and the older one is nine.
    Signature {
        names: &["emf"],
        matches: Matcher::Test(is_enhanced_metafile),
    },
    Signature {
        names: &["wmf"],
        matches: Matcher::Test(is_windows_metafile),
    },
    // ----------------------------------------------------------------------- documents
    // Text602, which opens with the tool's own four characters.
    Signature {
        names: &["602"],
        matches: Matcher::Test(|probe| starts_with(probe, b"@CT ")),
    },
    // CorelDRAW's RIFF container, which names itself in the form type after the size —
    // cases as CorelDRAW wrote them, and the version letter that follows.
    Signature {
        names: &["cdr"],
        matches: Matcher::Test(|probe| is_riff_form(probe, b"CDR")),
    },
    // Corel's presentation exchange, the other RIFF of the same suite.
    Signature {
        names: &["cmx"],
        matches: Matcher::Test(|probe| is_riff_form(probe, b"CMX")),
    },
    // A Computer Graphics Metafile in its binary encoding — the BEGIN METAFILE element,
    // which is the first thing a file of one holds — or in the clear-text encoding, which
    // spells the same element out.
    Signature {
        names: &["cgm"],
        matches: Matcher::Test(is_cgm),
    },
    // ClarisWorks, whose header is a version byte and the four characters every file it
    // wrote carries beside it.
    Signature {
        names: &["cwk"],
        matches: Matcher::Test(is_clarisworks),
    },
    // A dBASE table, which has no magic: what is asked is the version byte the format is
    // defined by, a date that is a date, and the header length and its terminator.
    Signature {
        names: &["dbf"],
        matches: Matcher::Test(is_dbase_table),
    },
    // AutoCAD's interchange format: the sentinel a binary file opens with, which is
    // unmistakable, or a text file whose first group is the section it starts with.
    Signature {
        names: &["dxf"],
        matches: Matcher::Test(is_dxf),
    },
    // A Hangul document of the version that writes its own name at the front. The version
    // after it is an OLE compound file like any other, and is left to the name.
    Signature {
        names: &["hwp"],
        matches: Matcher::Test(|probe| starts_with(probe, b"HWP Document File")),
    },
    Signature {
        names: &["lwp"],
        matches: Matcher::Test(|probe| starts_with(probe, b"WordPro")),
    },
    // An OS/2 metafile: the begin-document structured field, two bytes into a header whose
    // first word is its own length.
    Signature {
        names: &["met"],
        matches: Matcher::Test(|probe| at(probe, 2, &[0xD3, 0xA8, 0xA8])),
    },
    // PageMaker's two: the same header, told apart by the version word a hundred and ten
    // bytes in.
    Signature {
        names: &["pm6"],
        matches: Matcher::Test(|probe| is_pagemaker(probe, &[0x00, 0x06])),
    },
    Signature {
        names: &["pmd"],
        matches: Matcher::Test(|probe| is_pagemaker(probe, &[0x32, 0x06])),
    },
    // A StarView metafile, which is what the drawing layer's own format is called and what
    // the render engine imports.
    Signature {
        names: &["svm"],
        matches: Matcher::Test(|probe| starts_with(probe, b"VCLMTF")),
    },
    // SYLK: the identifier record a spreadsheet of the format opens with.
    Signature {
        names: &["slk"],
        matches: Matcher::Test(|probe| starts_with(probe, b"ID;P")),
    },
    // WordPerfect's container, which carries a document and a drawing under the same
    // magic and says which one it holds in its file-type byte.
    Signature {
        names: &["wpd"],
        matches: Matcher::Test(|probe| is_wordperfect(probe, 10)),
    },
    Signature {
        names: &["wpg"],
        matches: Matcher::Test(|probe| is_wordperfect(probe, 22)),
    },
    // A Microsoft Works document of the versions that wrote their own header rather than
    // an OLE compound file: the two bytes the format opens with, and the size word twenty
    // bytes in that one of its variants carries.
    //
    // The name is a hazardous one — Kingsoft's office suite writes a `.wps` of its own —
    // which is why what is asked is the header rather than the name.
    Signature {
        names: &["wps"],
        matches: Matcher::Test(|probe| {
            starts_with(probe, &[0x01, 0xFE])
                && (at(probe, 20, &[0xD0, 0x02]) || at(probe, 20, &[0xC4, 0x02]))
        }),
    },
    // Windows Write, which is its own format rather than an OLE compound file.
    Signature {
        names: &["wri"],
        matches: Matcher::Test(|probe| at(probe, 0, &[0x31, 0xBE, 0x00, 0x00])),
    },
    // Word for the Macintosh, whose versions open with the same byte and a version byte of
    // their own, and leave the four bytes after them at zero.
    Signature {
        names: &["mcw"],
        matches: Matcher::Test(|probe| {
            matches!(probe.first(), Some(0xFE))
                && matches!(probe.get(1), Some(0x32 | 0x34 | 0x37))
                && at(probe, 4, &[0x00, 0x00, 0x00, 0x00])
        }),
    },
    // A flat BIFF workbook: the beginning-of-file record an Excel 5.0/95 file and a
    // workspace file both open with. The Office tier reads neither — it is handed OLE
    // compound files and OOXML packages only — so this is asked of the bytes before the
    // name sends the file to an engine that would turn it down.
    Signature {
        names: &["xlw"],
        matches: Matcher::Test(|probe| at(probe, 0, &[0x09, 0x08, 0x08, 0x00, 0x00, 0x05])),
    },
    // The spreadsheets of the years before the current ones: Lotus 1-2-3's four save
    // formats and Quattro Pro's three. Every one of them opens with the same four bytes
    // and is told from the rest by the revision word that follows.
    Signature {
        names: &["123"],
        matches: Matcher::Test(|probe| {
            at(probe, 0, &[0x00, 0x00, 0x1A, 0x00])
                && matches!(probe.get(4..6), Some([0x03, 0x10] | [0x05, 0x10]))
        }),
    },
    Signature {
        names: &["wk3"],
        matches: Matcher::Test(|probe| at(probe, 0, &[0x00, 0x00, 0x1A, 0x00, 0x00, 0x10])),
    },
    Signature {
        names: &["wk4"],
        matches: Matcher::Test(|probe| at(probe, 0, &[0x00, 0x00, 0x1A, 0x00, 0x02, 0x10])),
    },
    Signature {
        names: &["wk1"],
        matches: Matcher::Test(|probe| at(probe, 0, &[0x00, 0x00, 0x02, 0x00, 0x06, 0x04])),
    },
    // `.wks` is two formats — the first Lotus 1-2-3 and a Microsoft Works spreadsheet —
    // and both are asked here, because the name is the only thing the two agree on.
    Signature {
        names: &["wks"],
        matches: Matcher::Test(|probe| {
            at(probe, 0, &[0x00, 0x00, 0x02, 0x00, 0x04, 0x04]) || at(probe, 0, &[0xFF, 0x00, 0x02])
        }),
    },
    Signature {
        names: &["wb2"],
        matches: Matcher::Test(|probe| at(probe, 0, &[0x00, 0x00, 0x02, 0x00, 0x02, 0x10])),
    },
    Signature {
        names: &["wq1"],
        matches: Matcher::Test(|probe| at(probe, 0, &[0x00, 0x00, 0x02, 0x00, 0x20, 0x51])),
    },
    Signature {
        names: &["wq2"],
        matches: Matcher::Test(|probe| at(probe, 0, &[0x00, 0x00, 0x02, 0x00, 0x21, 0x51])),
    },
    // The packages that declare their type in the head: an OpenDocument, the StarOffice
    // documents of the XML generation, and a Krita project. What each holds is the type
    // beside its own name, which is what makes it answerable at all — a zip on its own is
    // not a kind (see [`Matcher::Package`]).
    Signature {
        names: &["odb"],
        matches: Matcher::Package(b"application/vnd.oasis.opendocument.database"),
    },
    Signature {
        names: &["odb"],
        matches: Matcher::Package(b"application/vnd.oasis.opendocument.base"),
    },
    Signature {
        names: &["odc"],
        matches: Matcher::Package(b"application/vnd.oasis.opendocument.chart"),
    },
    Signature {
        names: &["odf"],
        matches: Matcher::Package(b"application/vnd.oasis.opendocument.formula"),
    },
    Signature {
        names: &["odg"],
        matches: Matcher::Package(b"application/vnd.oasis.opendocument.graphics"),
    },
    Signature {
        names: &["odm"],
        matches: Matcher::Package(b"application/vnd.oasis.opendocument.text-master"),
    },
    Signature {
        names: &["otg"],
        matches: Matcher::Package(b"application/vnd.oasis.opendocument.graphics-template"),
    },
    Signature {
        names: &["oth"],
        matches: Matcher::Package(b"application/vnd.oasis.opendocument.text-web"),
    },
    Signature {
        names: &["otm"],
        matches: Matcher::Package(b"application/vnd.oasis.opendocument.text-master-template"),
    },
    Signature {
        names: &["otp"],
        matches: Matcher::Package(b"application/vnd.oasis.opendocument.presentation-template"),
    },
    Signature {
        names: &["ots"],
        matches: Matcher::Package(b"application/vnd.oasis.opendocument.spreadsheet-template"),
    },
    Signature {
        names: &["ott"],
        matches: Matcher::Package(b"application/vnd.oasis.opendocument.text-template"),
    },
    Signature {
        names: &["sxd"],
        matches: Matcher::Package(b"application/vnd.sun.xml.draw"),
    },
    Signature {
        names: &["sxg"],
        matches: Matcher::Package(b"application/vnd.sun.xml.writer.global"),
    },
    Signature {
        names: &["sxi"],
        matches: Matcher::Package(b"application/vnd.sun.xml.impress"),
    },
    Signature {
        names: &["sxm"],
        matches: Matcher::Package(b"application/vnd.sun.xml.math"),
    },
    Signature {
        names: &["sxw"],
        matches: Matcher::Package(b"application/vnd.sun.xml.writer"),
    },
    Signature {
        names: &["stc"],
        matches: Matcher::Package(b"application/vnd.sun.xml.calc.template"),
    },
    Signature {
        names: &["std"],
        matches: Matcher::Package(b"application/vnd.sun.xml.draw.template"),
    },
    Signature {
        names: &["sti"],
        matches: Matcher::Package(b"application/vnd.sun.xml.impress.template"),
    },
    Signature {
        names: &["stw"],
        matches: Matcher::Package(b"application/vnd.sun.xml.writer.template"),
    },
    // A Krita project is the one design container that declares itself the same way.
    Signature {
        names: &["kra"],
        matches: Matcher::Package(b"application/x-krita"),
    },
    // --------------------------------------------------------------------------- ebooks
    // The books the ebook engine reads that say what they are in their own bytes. What is asked of
    // each is what is asked of every entry above — what the format writes at the front of a file,
    // and nothing else — and a book whose head says nothing is answered by its name instead, which
    // is what `calibre_formats` is for.
    //
    // Three kinds of name are deliberately *not* here, and each is a judgement about the format
    // rather than a gap:
    //
    // * **An EPUB is a box like any other**, and the one thing that tells it from a zip is a
    //   declaration — which is the same shape the OpenDocuments above have, and it is read the
    //   same way. A book of the format is the one book here that answers a kind of its own, since
    //   the name `epub` is claimed by no other list: a `.zip` whose declaration says it is an EPub
    //   is one, which is what the declaration is for.
    // * **A PalmDoc is not asked about**, although a `.prc` and a `.pdb` of that shape are books
    //   the engine reads. A Palmobook and an AportisDoc are the same header and the same
    //   identifiers, and the second is the render engine's by its own name — so the name is what
    //   settles which of the two engines a PalmDatabase is for, and a signature asked of the bytes
    //   alone would answer the one thing the two cannot be told apart by (see `GUARDS`, where the
    //   `.pdb` name is asked about already).
    // * **`snb`, `tcr`, `pml` and `htmlz` have no head to ask about at all.** A `.snb` is a SQLite
    //   database — the container every application on the machine keeps its own state in, which is
    //   the last thing a signature should claim — a `.tcr` and a `.pml` are text with a marker too
    //   weak to trust, and a `.htmlz` is a zip. All four are answered by their names.
    //
    // What is asked of a Mobipocket book is the two identifiers the format is defined by, written
    // where every Palm database keeps its own: the type `BOOK` and the creator `MOBI`. It is the
    // whole of the Kindle family under one signature — an `.azw`, an `.azw3`, an `.azw4` and a
    // `.prc` are the same database with the same pair inside — and it is what tells one from the
    // Palm ebook the render engine reads, which is the same header with other identifiers.
    Signature {
        names: &["azw", "azw3", "azw4", "mobi", "prc"],
        matches: Matcher::Test(is_mobipocket),
    },
    // The open one, which declares its own type beside its own name like every package above.
    Signature {
        names: &["epub"],
        matches: Matcher::Package(b"application/epub+zip"),
    },
    // FictionBook: the root element every file of the format is written with, behind the
    // declaration every XML file opens with. No text list claims the name, which is why the
    // signature is worth asking for at all: `<?xml` alone is every XML file there is.
    Signature {
        names: &["fb2"],
        matches: Matcher::Test(is_fictionbook),
    },
    // A scanned book: the chunk every DjVu file opens with, and the form type that follows it —
    // a document of pages, a single page, or a page included by another.
    Signature {
        names: &["djvu"],
        matches: Matcher::Test(is_djvu),
    },
    // And the container of the Sony readers of the 2000s, which writes its own three letters with
    // a zero byte between each of them.
    Signature {
        names: &["lrf"],
        matches: Matcher::Test(|probe| {
            starts_with(probe, &[b'L', 0x00, b'R', 0x00, b'F', 0x00, 0x00, 0x00])
        }),
    },
    // -------------------------------------------------------------------------- sounds
    // The audiobook brand of the container below: `ftypM4B` is an `.m4b`, which the common
    // table's M4A test deliberately does not catch — it asks for the brand `M4A` and nothing
    // else — so without this entry the file would be answered as the MP4 its box is.
    Signature {
        names: &["m4b", "m4a"],
        matches: Matcher::Test(|probe| at(probe, 4, b"ftypM4B")),
    },
    // -------------------------------------------------------------------------- videos
    // The ISO base media family, for the names the common table leaves out of it: `3gp`,
    // `3g2`, `3gpp`, `mj2` and `ismv` are the same `ftyp` box an MP4 and a MOV are.
    Signature {
        names: &["3gp", "3g2", "3gpp", "ismv", "m4v", "mj2", "mov", "mp4"],
        matches: Matcher::Test(|probe| at(probe, 4, b"ftyp")),
    },
    // The transport streams, asked with the same probe the text lists ask for the `.ts`
    // they share with TypeScript: the sync byte, run out over the packet size. A `.m2ts`
    // is the same stream with a timestamp in front of every packet, a `.tod` a JVC one
    // with its packets four bytes longer, and the rest are names for the same transport.
    Signature {
        names: &["ts", "m2ts", "mts", "m2t", "tp", "tr", "tod"],
        matches: Matcher::Test(crate::formats::video_formats::has_mpegts_packets),
    },
    // An MPEG program stream, which is what a `.mpe`, a `.m2p` and a DVD's `.vro` are: the
    // pack header every one of them opens with.
    Signature {
        names: &["mpe", "m2p", "vro"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0x00, 0x00, 0x01, 0xBA])),
    },
    // An MPEG elementary stream: a sequence header rather than a program stream, which is
    // what a `.m2v` is and what a bare `.mpg` sometimes turns out to be.
    Signature {
        names: &["m2v", "m1v", "mpg"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0x00, 0x00, 0x01, 0xB3])),
    },
    // A raw H.264 stream: the start code, then the sequence parameter set one opens with.
    // Which profile it is, is the decoder's business.
    Signature {
        names: &["h264", "264", "h26l", "avc"],
        matches: Matcher::Test(is_h264_stream),
    },
    // And the codecs beside it, each asked for the header a stream of that codec opens
    // with rather than for the start code all of them share. The names of one codec are
    // listed together — `hevc`, `h265` and `265` are one stream under three names — and
    // the three AVS generations are one entry because their headers differ by a field
    // offset no file states, which is a question a decoder answers and a probe cannot.
    Signature {
        names: &["265", "hevc", "h265"],
        matches: Matcher::Test(is_hevc_stream),
    },
    Signature {
        names: &["266", "vvc", "h266"],
        matches: Matcher::Test(is_vvc_stream),
    },
    // AV1, which is not start-code framed at all: a chain of length-prefixed units, the
    // first of which is the temporal delimiter or the sequence header of the stream.
    Signature {
        names: &["av1", "obu"],
        matches: Matcher::Test(is_av1_obu_stream),
    },
    Signature {
        names: &["avs", "avs2", "avs3", "cavs"],
        matches: Matcher::Test(is_avs_sequence),
    },
    Signature {
        names: &["vc1"],
        matches: Matcher::Test(|probe| {
            at(probe, 0, &[0x00, 0x00, 0x01, 0x0F])
                && probe.get(4).is_some_and(|byte| byte & 0xC0 == 0xC0)
        }),
    },
    // Raw Dirac, which is VC-2 under another name: the parse-info prefix, a parse code
    // that is one of the format's, and, where the unit it points at is in the probe, two
    // units that agree about where the first one ended.
    Signature {
        names: &["drc", "vc2"],
        matches: Matcher::Test(is_dirac_stream),
    },
    Signature {
        names: &["apv"],
        matches: Matcher::Test(|probe| starts_with(probe, b"aPv1")),
    },
    Signature {
        names: &["evc"],
        matches: Matcher::Test(is_evc_stream),
    },
    Signature {
        names: &["dv", "dif"],
        matches: Matcher::Test(is_dv_stream),
    },
    Signature {
        names: &["ivf"],
        matches: Matcher::Test(|probe| starts_with(probe, b"DKIF")),
    },
    // MXF, and the Sony camera format whose files are MXF under another name.
    Signature {
        names: &["mxf", "imx"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0x06, 0x0E, 0x2B, 0x34])),
    },
    Signature {
        names: &["nut"],
        matches: Matcher::Test(|probe| starts_with(probe, b"NUT/MULTI")),
    },
    Signature {
        names: &["y4m"],
        matches: Matcher::Test(|probe| starts_with(probe, b"YUV4MPEG2")),
    },
    // Bink: the four characters the format opens with, which its second version shares.
    Signature {
        names: &["bik", "bk2"],
        matches: Matcher::Test(|probe| starts_with(probe, b"BIK")),
    },
    // The games' and tools' own containers, each with a header of its own: an Id RoQ
    // file, a Smacker movie, a GameCube THP, an XMV, a TiVo stream, a RealMedia one, a
    // FILM/CPK, a GXF, an IFV, a KUX, an Interplay MVE, a PMP, a Scaleform USM, a Dahua
    // camera's DAV, a Vivo stream, a VC-1 test stream and a PlayStation STR.
    Signature {
        names: &["roq"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0x84, 0x10, 0xFF, 0xFF, 0xFF, 0xFF])),
    },
    Signature {
        names: &["smk"],
        matches: Matcher::Test(|probe| starts_with(probe, b"SMK2") || starts_with(probe, b"SMK4")),
    },
    Signature {
        names: &["thp"],
        matches: Matcher::Test(|probe| starts_with(probe, b"THP\0")),
    },
    Signature {
        names: &["xmv"],
        matches: Matcher::Test(|probe| {
            at(probe, 12, b"xobX") && matches!(read_le32(probe, 16), Some(1..=4))
        }),
    },
    Signature {
        names: &["yop"],
        matches: Matcher::Test(is_yop),
    },
    Signature {
        names: &["rsd"],
        matches: Matcher::Test(is_rsd),
    },
    Signature {
        names: &["rm", "rmvb"],
        matches: Matcher::Test(|probe| {
            starts_with(probe, b".RMF\x00\x00")
                || starts_with(probe, b".RMP\x00\x00")
                || starts_with(probe, &[0x2E, 0x72, 0x61, 0xFD])
        }),
    },
    // A recorded RealMedia stream, which Real's own recorder writes rather than the server:
    // the seven bytes of its header, or the three a recording of another kind opens with.
    Signature {
        names: &["ivr"],
        matches: Matcher::Test(|probe| {
            starts_with(probe, b".R1M\x00\x01\x01") || starts_with(probe, b".REC")
        }),
    },
    Signature {
        names: &["cpk"],
        matches: Matcher::Test(|probe| starts_with(probe, b"FILM") && at(probe, 16, b"FDSC")),
    },
    Signature {
        names: &["gxf"],
        matches: Matcher::Test(|probe| {
            starts_with(probe, &[0x00, 0x00, 0x00, 0x00, 0x01, 0xBC])
                && at(probe, 10, &[0x00, 0x00, 0x00, 0x00, 0xE1, 0xE2])
        }),
    },
    Signature {
        names: &["ifv"],
        matches: Matcher::Test(|probe| {
            starts_with(
                probe,
                &[
                    0x11, 0xD2, 0xD3, 0xAB, 0xBA, 0xA9, 0xCF, 0x11, 0x8E, 0xE6, 0x00, 0xC0, 0x0C,
                    0x20, 0x53, 0x65, 0x44,
                ],
            )
        }),
    },
    Signature {
        names: &["kux"],
        matches: Matcher::Test(|probe| starts_with(probe, b"KDK\x00\x00")),
    },
    Signature {
        names: &["mve"],
        matches: Matcher::Test(|probe| contains(probe, b"Interplay MVE File\x1A\x00\x1A\x00")),
    },
    Signature {
        names: &["pmp"],
        matches: Matcher::Test(|probe| {
            starts_with(probe, b"pmpm") && at(probe, 4, &[0x01, 0x00, 0x00, 0x00])
        }),
    },
    // Scaleform's video, which FFmpeg has no demuxer for and `file`'s own table names: the
    // four characters the container opens with and the marker thirty-two bytes in.
    Signature {
        names: &["usm"],
        matches: Matcher::Test(|probe| starts_with(probe, b"CRID") && at(probe, 32, b"@UTF")),
    },
    Signature {
        names: &["dav"],
        matches: Matcher::Test(|probe| {
            starts_with(probe, b"DAHUA")
                || (starts_with(probe, b"DHAV")
                    && matches!(probe.get(4), Some(0xF0 | 0xF1 | 0xFC | 0xFD)))
        }),
    },
    Signature {
        names: &["viv"],
        matches: Matcher::Test(|probe| {
            probe.first() == Some(&0)
                && (at(probe, 4, b"Version:Vivo/") || at(probe, 5, b"Version:Vivo/"))
        }),
    },
    Signature {
        names: &["rcv"],
        matches: Matcher::Test(is_vc1_test_stream),
    },
    Signature {
        names: &["str"],
        matches: Matcher::Test(|probe| {
            starts_with(
                probe,
                &[
                    0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00,
                ],
            )
        }),
    },
    Signature {
        names: &["wtv"],
        matches: Matcher::Test(|probe| {
            starts_with(
                probe,
                &[
                    0xB7, 0xD8, 0x00, 0x20, 0x37, 0x49, 0xDA, 0x11, 0xA6, 0x4E, 0x00, 0x07, 0xE9,
                    0x5E, 0xAD, 0x8D,
                ],
            )
        }),
    },
    Signature {
        names: &["nsv"],
        matches: Matcher::Test(|probe| starts_with(probe, b"NSVf") || starts_with(probe, b"NSVs")),
    },
    // The last of the video entries are the ones whose format has no magic to ask for and
    // whose own probe scores a shape instead. Each is the shape FFmpeg's demuxer scores —
    // a block table, a periodic command byte, a header of dimensions and offsets, a table
    // that has to chain, a start code twice over — and each is here rather than in a table
    // of its own for the reason the rest of this one exists: a `.c93` whose bytes are a
    // document is not a document.
    Signature {
        names: &["c93"],
        matches: Matcher::Test(is_c93),
    },
    Signature {
        names: &["cdg"],
        matches: Matcher::Test(is_cdg),
    },
    Signature {
        names: &["cdxl", "xl"],
        matches: Matcher::Test(is_cdxl),
    },
    Signature {
        names: &["moflex"],
        matches: Matcher::Test(is_moflex),
    },
    Signature {
        names: &["h261"],
        matches: Matcher::Test(is_h261_picture),
    },
    Signature {
        names: &["h263"],
        matches: Matcher::Test(is_h263_picture),
    },
    // ----------------------------------------------------------------------------- text
    // An Ogg stream is a video where what it carries is Theora: the same container holds
    // Vorbis and Opus, and a preview of a sound is not a preview — so the codec is asked
    // for here rather than the container being taken for a picture.
    Signature {
        names: &["ogv", "ogm"],
        matches: Matcher::Test(|probe| starts_with(probe, b"OggS") && contains(probe, b"theora")),
    },
    // And the sounds of the same container, which the entry above deliberately leaves to this
    // one: an Ogg page carrying Vorbis, Opus, Speex or a FLAC stream and no Theora is a song.
    // The order of the two is the whole of the difference — a Theora film whose audio track is
    // Vorbis declares both, and it is the video entry that is asked first.
    Signature {
        names: &["ogg", "oga", "opus", "spx"],
        matches: Matcher::Test(is_ogg_sound),
    },
    // The sound formats the common table has no signature for, each named by the magic it
    // writes at the front of a file: the containers FFmpeg's player reads that no file manager
    // would name. A file of one of these names that says nothing at the front of itself is
    // answered by the name it carries, which is what the `[audio]` list is for.
    //
    // Musepack, in both of the shapes the format is written in: the `MPCK` an SV8 file opens
    // with, and the `MP+` its predecessor used.
    Signature {
        names: &["mpc"],
        matches: Matcher::Test(|probe| starts_with(probe, b"MPCK") || starts_with(probe, b"MP+")),
    },
    // WavPack, whose four characters open every file of the format.
    Signature {
        names: &["wv"],
        matches: Matcher::Test(|probe| starts_with(probe, b"wvpk")),
    },
    // True Audio, which names itself in the first four bytes.
    Signature {
        names: &["tta"],
        matches: Matcher::Test(|probe| starts_with(probe, b"TTA1")),
    },
    // TAK, whose magic is the format's own name backwards and a case apart.
    Signature {
        names: &["tak"],
        matches: Matcher::Test(|probe| starts_with(probe, b"tBaK")),
    },
    // OptimFROG, in the two shapes a stream is written in — the one that is only a stream, and
    // the one that keeps its samples in a separate file.
    Signature {
        names: &["ofr", "ofs"],
        matches: Matcher::Test(|probe| starts_with(probe, b"OFR ") || starts_with(probe, b"OFS ")),
    },
    // Shorten, whose four letters open every file of it.
    Signature {
        names: &["shn"],
        matches: Matcher::Test(|probe| starts_with(probe, b"ajkg")),
    },
    // DSD in the Philips container, which is an IFF form of the format's own name.
    Signature {
        names: &["dff"],
        matches: Matcher::Test(|probe| starts_with(probe, b"FRM8")),
    },
    // The two older Unix containers: the `.snd` header every AU file opens with, and the CAF
    // header Apple's own writer uses.
    Signature {
        names: &["au", "snd"],
        matches: Matcher::Test(|probe| starts_with(probe, b".snd")),
    },
    Signature {
        names: &["caf"],
        matches: Matcher::Test(|probe| starts_with(probe, b"caff")),
    },
    // Creative's sample format, which announces itself in as many words.
    Signature {
        names: &["voc"],
        matches: Matcher::Test(|probe| starts_with(probe, b"Creative Voice File")),
    },
    // The two cinema codecs, named by the sync word a frame of each opens with: Dolby's is the
    // two bytes every AC-3 and E-AC-3 frame begins with, and DTS's is the sixteen-bit sync of
    // its own framing.
    Signature {
        names: &["ac3", "eac3"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0x0B, 0x77])),
    },
    Signature {
        names: &["dts", "dtshd"],
        matches: Matcher::Test(|probe| at(probe, 0, &[0x7F, 0xFE, 0x80, 0x01])),
    },
    // RealAudio, whose marker every file of the format is written with — the modern one, which
    // names the codec in its header, and the older `.ra` the extension is for.
    Signature {
        names: &["ra"],
        matches: Matcher::Test(|probe| starts_with(probe, b".ra\xfd")),
    },
    // PostScript, which is what an `.eps` is and what an Illustrator document saved
    // without its compatibility layer is. The PDF half of that family needs no entry here:
    // a page is a page, and the common table names it.
    Signature {
        names: &["eps", "epsi", "ai", "ps"],
        matches: Matcher::Test(|probe| starts_with(probe, b"%!PS-Adobe")),
    },
    // The one text format a document is likely to be mistaken for: an RTF file under a
    // document's name is the text the text lists would have shown anyway.
    Signature {
        names: &["rtf"],
        matches: Matcher::Test(|probe| starts_with(probe, b"{\\rtf")),
    },
    // A collection of fonts is the one font container the common table does not name,
    // because what it holds is faces rather than a face.
    Signature {
        names: &["ttc"],
        matches: Matcher::Test(|probe| starts_with(probe, b"ttcf")),
    },
];

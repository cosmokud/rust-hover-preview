//! What a file's own bytes say it is, where that is not what its name says.
//!
//! One engine — the application Word, Excel and PowerPoint are started through, the
//! LibreOffice a document is converted by, the player FFmpeg draws a video in — is only
//! ever handed a file whose content belongs to it. The Office tier has asked its own
//! header what it is before a path was given to it for as long as it has existed (see
//! `office_formats::container_kind`); what this module adds is the same question asked
//! before *every* kind, and an answer that is finer than yes or no: a `.docx` whose bytes
//! are an MP4 is not merely turned down, it is played by the player a video is played by,
//! and a `.cdr` that is really a PNG gets the picture preview. A file whose content is a
//! format no kind of this app previews — an executable, an audio file, a format nothing
//! here reads — is answered with no preview, which is the answer that leaves every engine
//! unstarted.
//!
//! Three tables answer what a file is, and what is in each of them is what it is for.
//! The common formats — the pictures, the containers a video arrives in, a PDF, the fonts,
//! the archives — are the [`infer`] crate's: a small table of the signatures every tool
//! agrees on, no dependencies of its own, and the same answer any file manager would give.
//! The formats it does not carry but a list of this app's does are this app's own table,
//! asked in the shape the rest of the app already asks signatures in: a transport stream
//! by its sync byte, an MXF by its partition pack, a raw H.264 stream by its sequence
//! parameter set, a WordPerfect document by the signature every file of the family opens
//! with — the same kind of question `office_formats::container_kind` asks an Office
//! document and `has_pdf_header_in` asks a PDF, and one of them, the transport stream, is
//! asked of `video_formats` rather than written again.
//!
//! That second table is the whole of what the two engines' lists hold and the common one
//! does not: every name `video_formats` carries that FFmpeg demuxes, and every name
//! `libre_formats` carries that the render engine imports, less the names [`infer`] has
//! already answered for. What a test is worth differs from format to format and is said
//! with each entry: a magic where the format has one, a start code and a header where the
//! format is a raw stream, a shape where the format's own demuxer scores one instead of
//! reading a magic (see [`SIGNATURES`] for what is in it, where each answer comes from, and
//! for the three kinds of name that are deliberately not in it).
//!
//! What no table answers is answered by the name the file carries, which is the third
//! table: the few names a head cannot settle at all — a container that says nothing about
//! itself inside the probe, a format whose signature is its own text, a name no demuxer
//! reads — each of which keeps the answer it always had. See [`KIND_BY_NAME`] for what is
//! in it and why each entry is there.
//!
//! A table of names is a weaker answer than a table of signatures, and it is asked where
//! the strong ones have nothing to say. What it can answer is only what a name says, so a
//! file that carries one of its names and is not the format that name is written for would
//! be routed by it wrongly — which is why a name whose format is *two* formats, with the
//! engine reading one of them and nothing here previewing the other, is asked a question
//! of its own before it is answered: `.pdb` is a Palm OS database, which the render
//! engine's filters read as an ebook, and it is also the Microsoft program database a
//! compiler writes beside its binaries, which is no document at all — see
//! `palm_ebook_or_program_database` for what is asked of the bytes there. **An engine is
//! handed a file of such a name only where its own header says it is the format that
//! engine reads**, and a file that is the other one shows nothing rather than starting an
//! engine it is not for.
//!
//! What is deliberately in none of the tables matters as much as what is. **A container
//! is not a kind**, so nothing here answers with a box: a zip is what an OpenDocument, an
//! iWork document and an Office package all arrive in, the OLE compound file is what every
//! `.doc`, `.xls` and `.ppt` is, and the file's own name is a better answer than the box
//! is — which is also why a `.docx` that is really a `.doc` is still handed to an engine.
//! The one container that *is* answered for is the package that declares its own type at
//! an offset every file of the format agrees on — an OpenDocument, a StarOffice XML
//! document, a Krita project — because what is read there is the document's own answer
//! rather than the box's (see [`Matcher::Package`]). A format whose signature *is* its
//! text — an SVG, an HTML page, a JSON file, a flat OpenDocument — is left alone for the
//! same reason: what it is is what the text lists decide, and a name another kind already
//! read is not improved by a prefix match.
//!
//! The probe is one read of [`PROBE_BYTES`] and every question asked here is about the
//! head of a file. A file whose content is not on this machine is not opened at all, which
//! is the rule every other read in this app follows (see `cloud_files`).
//!
//! What the tables answer is read as one of three things:
//!
//! * A kind, where the format is one this app's lists claim and it is not the kind the
//!   name claimed: that kind is what the file is previewed as. The third table answers a
//!   kind outright, because there is nothing for a name to disagree with — the name is the
//!   question — which is what makes a format with no signature worth writing down at all.
//! * [`Content::Foreign`], where the format is one no list claims — an executable, an
//!   audio file, a box of some kind this app has no preview for — or where the third
//!   table's name turned out to be the other format's: a `.pdb` that is not the ebook the
//!   engine reads is a file with nothing to show, and no engine is started for it.
//! * [`Content::Unknown`], which is "no opinion" and is the answer in three cases: nothing
//!   in the tables matched and the name is not one of the third table's either (a raw
//!   stream, an exotic container, a text file), the format is one the file's own name
//!   already names (the two agree, so there is nothing to override), and a file with no
//!   extension at all, whose name has nothing to disagree with. The name decides, as it
//!   always has.
//!
//! The answer is held between hovers, keyed by the file and the version of it that was
//! read, because one hover asks this three times — the hook that raises it, the loader
//! that fills it, and the layout that places it — and a file edited in place is read again
//! rather than held to an answer about what it used to be.

use crate::config::config::{AppConfig, PreviewType};
use crate::CONFIG;
use infer::{MatcherType, Type};
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use std::time::SystemTime;

/// How much of a file is read to answer: the front of it, which is where every signature
/// is. Nothing further is ever asked for — the questions here are all about the head of a
/// file — so a hover onto a video costs this and not the video.
const PROBE_BYTES: usize = 4096;

/// How many answers are held between hovers. One hover of one file asks for one answer,
/// and a folder is swept a file at a time, so the list is a session's worth of files
/// rather than a scan's: what it bounds is a pointer dragged across a large folder, and
/// what it costs when it is reached is everything held, the way every other cache of this
/// app's shape answers that.
const ANSWERS_MAX_ENTRIES: usize = 512;

/// What a file is, where the answer is not what its name says it is.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Content {
    /// The format the file holds — or, where no table named the bytes, the format its name
    /// is written for — is one of this app's kinds. Where the two disagree this is the kind
    /// that should have the file; where they do not, it is the kind the name already gave
    /// it.
    Kind(PreviewType),
    /// The content is a format no kind of this app previews — an executable, an audio
    /// file, a format this app has no reader for, or the other format a guarded name is
    /// written under (see [`GUARDS`]). There is nothing to show, and no engine of this
    /// app's is the engine for it.
    Foreign,
    /// No opinion: nothing was recognized, or what was recognized is what the name
    /// already said, or the name has no extension to disagree with. The name decides, as
    /// it always has.
    Unknown,
}

/// The kind a file's content belongs to, where that is not what its name says.
///
/// Asked where a file's kind is decided — the hook's gate, the loader and the layout —
/// and answered from the cache where the same file has been asked about already in this
/// hover. Where the setting is off this is [`Content::Unknown`] without reading anything,
/// which leaves every kind to the lists it has always been decided by.
pub fn of(path: &Path) -> Content {
    if !is_confirmed() {
        return Content::Unknown;
    }

    let key = answer_key(path);

    if let Ok(answers) = ANSWERS.lock() {
        if let Some(answer) = answers.get(&key) {
            return *answer;
        }
    }

    let answer = read(path);

    if let Ok(mut answers) = ANSWERS.lock() {
        if answers.len() >= ANSWERS_MAX_ENTRIES {
            answers.clear();
        }
        answers.insert(key, answer);
    }

    answer
}

/// Whether a file's content is confirmed at all: the tray's `Confirm File Type`, which is
/// what every question in this module is behind.
///
/// It is the setting the picture paths already ask — a picture is decoded by its header
/// rather than by its name where it is on (see `preview_window`) — and it is read from the
/// configuration each time rather than captured, so an edit applies to the next hover
/// rather than to the next start. Every caller that asks it through this module must not
/// be holding the configuration itself: the lock is not reentrant.
pub fn is_confirmed() -> bool {
    CONFIG
        .lock()
        .map(|config| config.confirm_file_type)
        .unwrap_or(false)
}

/// What the file holds, read once and not held.
///
/// The tables are asked in the order they can answer in: the bytes first, through the
/// formats every tool agrees on and then through the ones an engine here reads that no
/// such table carries, and the name the file is under last, for the formats whose own head
/// is nothing either table knows — see [`KIND_BY_NAME`] for what is in that one.
fn read(path: &Path) -> Content {
    // A file whose content is still in the cloud is not opened, because the open is what
    // starts the transfer: the same rule the hook and every reader follow.
    if crate::shell::cloud_files::needs_download(path) {
        return Content::Unknown;
    }

    let Some(probe) = probe(path) else {
        return Content::Unknown;
    };

    if let Some(names) = detected_names(&probe) {
        let Ok(config) = CONFIG.lock() else {
            return Content::Unknown;
        };

        return classify(path, names, &config);
    }

    // Neither table named the front of the file, so the name it carries is what is left to
    // ask. A file with no name to ask about — one with no extension at all — has nothing
    // here to disagree with either, and is left to the lists.
    own_extension(path)
        .and_then(|extension| kind_by_name(&extension, &probe))
        .unwrap_or(Content::Unknown)
}

/// The front of the file, or nothing where it could not be read.
///
/// One read of a bounded buffer rather than a read to the end: a file this app is asked
/// about may be a video of many gigabytes, and what it is is decided by its first
/// kilobytes. What is read is taken up to the bound rather than asked for in one read,
/// because a single read is free to come back with less than was asked for and a signature
/// sits at an offset — a short answer would be read as a file nothing here can name.
fn probe(path: &Path) -> Option<Vec<u8>> {
    let file = File::open(path).ok()?;
    let mut probe = Vec::new();
    file.take(PROBE_BYTES as u64).read_to_end(&mut probe).ok()?;

    Some(probe)
}

/// What the file's own bytes say it is: the names the format is known by to the lists, or
/// nothing where the front of the file is not a format either table names.
fn detected_names(probe: &[u8]) -> Option<&'static [&'static str]> {
    // The formats every table agrees on, in the names this app's lists carry.
    if let Some(kind) = infer::get(probe) {
        if let Some(names) = names_of(&kind) {
            return Some(names);
        }
    }

    // And the formats no such table carries, which is what the table below is for.
    SIGNATURES
        .iter()
        .find(|signature| signature.matches(probe))
        .map(|signature| signature.names)
}

/// What one of the common formats means for this app, in the names its lists carry, or
/// nothing where the format is one the name should still decide.
///
/// A name is not the only way a format is written — a JPEG is a `jpg` and a motion JPEG is
/// an `mjpg`, an ISO base media file is an `mp4` and a `mov` — and every one of those
/// names belongs here, because the answer is not "what is this" but "is this what the file
/// is called": a name the file already carries is the two agreeing, and the reply is that
/// there is nothing to override.
///
/// What is *not* answered here is as deliberate as what is. A **box** — a zip, a 7z, a
/// tar — is not a kind, so nothing is answered for one and the name decides, which is what
/// keeps an OpenDocument, an iWork document and an Office package working. Ogg is the one
/// container that is sometimes a video, and which it is is a question for the table below
/// rather than for a guess. A format whose signature is its text — an SVG, an HTML page —
/// is left to the text lists for the same reason.
fn names_of(kind: &Type) -> Option<&'static [&'static str]> {
    match kind.extension() {
        "zip" | "7z" | "rar" | "tar" | "gz" | "tgz" | "bz2" | "xz" | "cab" | "iso" | "dmg"
        | "ogg" => None,

        // Nothing to show at all: a program, or a sound. No kind of this app previews
        // either, so a document's name that holds one starts no engine — which is the
        // answer this whole module exists to give for the files that were never documents.
        _ if matches!(kind.matcher_type(), MatcherType::App | MatcherType::Audio) => Some(&[]),

        "mp4" | "m4v" => Some(&["mp4", "m4v", "mov", "qt", "3gp", "3g2", "f4v"]),
        "mov" => Some(&["mov", "qt", "mp4"]),
        "mkv" => Some(&["mkv", "mk3d", "mka"]),
        "webm" => Some(&["webm", "mkv"]),
        "avi" => Some(&["avi", "divx"]),
        "flv" => Some(&["flv", "f4v"]),
        "wmv" | "asf" => Some(&["wmv", "asf", "dvr-ms"]),
        "mpg" => Some(&["mpg", "mpeg", "vob", "m2v", "m1v", "mpv"]),
        "swf" => Some(&["swf"]),

        "png" => Some(&["png", "apng"]),
        "jpg" => Some(&["jpg", "jpeg", "jpe", "jfif", "mjpg", "mjpeg"]),
        "gif" => Some(&["gif"]),
        "bmp" => Some(&["bmp"]),
        "tif" => Some(&["tif", "tiff"]),
        "webp" => Some(&["webp"]),
        "ico" => Some(&["ico"]),
        "psd" => Some(&["psd", "psb"]),
        "heic" => Some(&["heic", "heif"]),
        "avif" => Some(&["avif"]),
        // A JPEG XL picture is asked of the codec Windows has for it, and an OpenRaster
        // project is read out of the container it keeps its finished picture in: both are
        // names a list here carries, and the common table is what names them.
        "jxl" => Some(&["jxl"]),
        "ora" => Some(&["ora"]),

        "pdf" => Some(&["pdf"]),
        "ttf" => Some(&["ttf"]),
        "otf" => Some(&["otf"]),
        "woff" => Some(&["woff"]),
        "woff2" => Some(&["woff2"]),

        _ => None,
    }
}

/// One format this app can name from the front of a file.
struct Signature {
    /// The names the format is known by to the lists, which is what a kind is found from.
    ///
    /// A format is written under more than one name — a Matroska file is an `mkv` and a
    /// `webm`, a transport stream is a `ts` and a `m2ts` — and all of them are listed
    /// because the answer is not "what is this" but "is this what the file is called": a
    /// name the file already carries is the two agreeing, and a name no list claims is a
    /// kind this app does not preview.
    names: &'static [&'static str],
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
    fn matches(&self, probe: &[u8]) -> bool {
        match self.matches {
            Matcher::Test(test) => test(probe),
            Matcher::Package(mime) => is_declared_package(probe, mime),
        }
    }
}

/// The formats the common table does not carry, which are the ones an engine here reads
/// and no file manager would name.
///
/// Every name the two engines' lists carry that `infer` has no signature for is answered
/// here, so that the content of a file decides what it is whichever engine the file
/// belongs to: a picture this app's own decoder reads, a drawing the drawing layer plays,
/// a document the render engine imports, and the containers and raw streams FFmpeg plays
/// that no common table names. What a format's test is worth differs, and where it is not
/// a literal the test says what it is: the magic of a format whose header is its own, a
/// start code followed by the header a stream of that codec opens with, or — for the
/// packages that declare a type — the declaration itself.
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
const SIGNATURES: &[Signature] = &[
    // ------------------------------------------------------------------------- pictures
    // A DirectDraw Surface: the four-character code and the size of the header that
    // follows it, which is what tells a texture from a file that begins with a word.
    Signature {
        names: &["dds"],
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
        names: &["pcd"],
        matches: Matcher::Test(is_photo_cd),
    },
    Signature {
        names: &["pcx"],
        matches: Matcher::Test(is_pcx),
    },
    Signature {
        names: &["ras"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0x59, 0xA6, 0x6A, 0x95])),
    },
    // A QuickDraw PICT on disk carries a five-hundred-byte header, so the version operator
    // that opens the picture is at 522 rather than at the front of the file.
    Signature {
        names: &["pct"],
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
        matches: Matcher::Test(|probe| starts_with(probe, &[0x8B, b'J', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A])),
    },
    Signature {
        names: &["mng"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0x8A, b'M', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A])),
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
        matches: Matcher::Test(|probe| at(probe, 0, &[0x00, 0x00, 0x1A, 0x00]) && matches!(probe.get(4..6), Some([0x03, 0x10] | [0x05, 0x10]))),
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
        matches: Matcher::Test(|probe| at(probe, 0, &[0x00, 0x00, 0x02, 0x00, 0x04, 0x04]) || at(probe, 0, &[0xFF, 0x00, 0x02])),
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
        matches: Matcher::Test(|probe| at(probe, 0, &[0x00, 0x00, 0x01, 0x0F]) && probe.get(4).is_some_and(|byte| byte & 0xC0 == 0xC0)),
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
        matches: Matcher::Test(|probe| at(probe, 12, b"xobX") && matches!(read_le32(probe, 16), Some(1..=4))),
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
        matches: Matcher::Test(|probe| starts_with(probe, b".RMF\x00\x00") || starts_with(probe, b".RMP\x00\x00") || starts_with(probe, &[0x2E, 0x72, 0x61, 0xFD])),
    },
    // A recorded RealMedia stream, which Real's own recorder writes rather than the server:
    // the seven bytes of its header, or the three a recording of another kind opens with.
    Signature {
        names: &["ivr"],
        matches: Matcher::Test(|probe| starts_with(probe, b".R1M\x00\x01\x01") || starts_with(probe, b".REC")),
    },
    Signature {
        names: &["cpk"],
        matches: Matcher::Test(|probe| starts_with(probe, b"FILM") && at(probe, 16, b"FDSC")),
    },
    Signature {
        names: &["gxf"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0x00, 0x00, 0x00, 0x00, 0x01, 0xBC]) && at(probe, 10, &[0x00, 0x00, 0x00, 0x00, 0xE1, 0xE2])),
    },
    Signature {
        names: &["ifv"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0x11, 0xD2, 0xD3, 0xAB, 0xBA, 0xA9, 0xCF, 0x11, 0x8E, 0xE6, 0x00, 0xC0, 0x0C, 0x20, 0x53, 0x65, 0x44])),
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
        matches: Matcher::Test(|probe| starts_with(probe, b"pmpm") && at(probe, 4, &[0x01, 0x00, 0x00, 0x00])),
    },
    // Scaleform's video, which FFmpeg has no demuxer for and `file`'s own table names: the
    // four characters the container opens with and the marker thirty-two bytes in.
    Signature {
        names: &["usm"],
        matches: Matcher::Test(|probe| starts_with(probe, b"CRID") && at(probe, 32, b"@UTF")),
    },
    Signature {
        names: &["dav"],
        matches: Matcher::Test(|probe| starts_with(probe, b"DAHUA") || (starts_with(probe, b"DHAV") && matches!(probe.get(4), Some(0xF0 | 0xF1 | 0xFC | 0xFD)))),
    },
    Signature {
        names: &["viv"],
        matches: Matcher::Test(|probe| probe.first() == Some(&0) && (at(probe, 4, b"Version:Vivo/") || at(probe, 5, b"Version:Vivo/"))),
    },
    Signature {
        names: &["rcv"],
        matches: Matcher::Test(is_vc1_test_stream),
    },
    Signature {
        names: &["str"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00])),
    },
    Signature {
        names: &["wtv"],
        matches: Matcher::Test(|probe| starts_with(probe, &[0xB7, 0xD8, 0x00, 0x20, 0x37, 0x49, 0xDA, 0x11, 0xA6, 0x4E, 0x00, 0x07, 0xE9, 0x5E, 0xAD, 0x8D])),
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

/// Whether the front of a file is one of the packages that declare their own type: a zip
/// whose first entry is a stored `mimetype` entry holding `mime`.
///
/// The local file header is thirty bytes and the entry is stored rather than deflated with
/// no extra field, so the eight characters of the entry's name sit at 30 and its content
/// at 38 — an offset every file of every format that writes its type this way agrees on.
/// What follows the type is the next record of the archive, which is to say the letter
/// `P`: asking for it is what keeps a subtype from answering for the type whose name it
/// begins with, the way `…graphics` begins `…graphics-template`.
fn is_declared_package(probe: &[u8], mime: &[u8]) -> bool {
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
fn is_dds_surface(probe: &[u8]) -> bool {
    starts_with(probe, b"DDS ") && at(probe, 4, &[0x7C, 0x00, 0x00, 0x00])
}

/// Whether the front of a file is an OpenEXR picture: the magic, which is the number the
/// format's first version was written as.
fn is_openexr(probe: &[u8]) -> bool {
    starts_with(probe, &[0x76, 0x2F, 0x31, 0x01])
}

/// Whether the front of a file is a Radiance picture, in the header the format has written
/// since it was called RGBE and under the name it has now.
fn is_radiance_picture(probe: &[u8]) -> bool {
    starts_with(probe, b"#?RADIANCE") || starts_with(probe, b"#?RGBE")
}

/// Whether the front of a file is a Netpbm picture.
///
/// `P1` to `P6` are two characters and a separator, which is little enough to be an
/// accident, so the first digit of the size that must follow them is asked for as well.
/// `P7` is a PAM file, whose header names its own fields and ends with a word of its own,
/// and which is asked for those rather than for the two characters.
fn is_netpbm(probe: &[u8]) -> bool {
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
fn is_photo_cd(probe: &[u8]) -> bool {
    at(probe, 2048, b"PCD_IPI") || starts_with(probe, b"PCD_OPA")
}

/// Whether the front of a file is a ZSoft PCX picture: the encoding byte the format is
/// defined by, the version and the encoding it carries, and a palette that is filled in.
fn is_pcx(probe: &[u8]) -> bool {
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
fn is_quickdraw_pict(probe: &[u8]) -> bool {
    at(probe, 522, &[0x00, 0x11, 0x02, 0xFF]) || at(probe, 522, &[0x11, 0x01])
}

/// Whether the front of a file is a JPEG 2000 picture, in either of the two shapes the
/// format is written in.
///
/// The file format is a box structure, and its first box is the signature box — twelve
/// bytes long, so it announces itself at four rather than at the front of the file. What the
/// code stream is instead is the picture with none of the boxes: the marker every codestream
/// opens with, which is the same thing a JPEG 2000 file's `jp2c` box holds.
fn is_jpeg2000(probe: &[u8]) -> bool {
    at(probe, 4, &[b'j', b'P', b' ', b' ', 0x0D, 0x0A, 0x87, 0x0A])
        || starts_with(probe, &[0xFF, 0x4F, 0xFF, 0x51])
}

/// Whether the front of a file is a Flexible Image Transport System picture.
///
/// The header is ASCII card images of eighty bytes each, and the first of them is the record
/// that says a picture starts here. The card that has to follow it is asked for as well: a
/// text file that happens to open with `SIMPLE  =` is not a picture of the sky.
fn is_fits(probe: &[u8]) -> bool {
    starts_with(probe, b"SIMPLE  =") && contains(probe, b"BITPIX")
}

/// Whether the front of a file is a Silicon Graphics picture: the two-byte magic, the
/// storage the picture is held in, and the bytes to a channel — in either byte order, since
/// the format is written both ways.
fn is_silicon_graphics(probe: &[u8]) -> bool {
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
fn is_raw_tiff_variant(probe: &[u8]) -> bool {
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
fn names_the_container_of_a_raw(names: &[&str]) -> bool {
    names
        .iter()
        .any(|name| name.eq_ignore_ascii_case("tif") || name.eq_ignore_ascii_case("tiff"))
}

/// Whether the front of a file is a RIFF container of one of Corel's formats, which say
/// which one they are in the form type that follows the size word.
fn is_riff_form(probe: &[u8], form: &[u8]) -> bool {
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
fn is_enhanced_metafile(probe: &[u8]) -> bool {
    starts_with(probe, &[0x01, 0x00, 0x00, 0x00])
        && at(probe, 40, b" EMF")
        && at(probe, 44, &[0x00, 0x00, 0x01, 0x00])
        && matches!(read_le32(probe, 4), Some(size) if size >= 88)
}

/// Whether the front of a file is a Windows metafile: the placeable header's own key, or a
/// bare metafile, whose header size word is nine for both of the versions it has.
fn is_windows_metafile(probe: &[u8]) -> bool {
    if starts_with(probe, &[0xD7, 0xCD, 0xC6, 0x9A]) {
        return true;
    }

    matches!(probe.first(), Some(0x01 | 0x02)) && at(probe, 2, &[0x09, 0x00])
}

/// Whether the front of a file is a Computer Graphics Metafile, in either of the two
/// encodings the format defines: the binary one, whose first element is BEGIN METAFILE
/// with a class and an identifier of zero and one, or the clear-text one, which spells the
/// same element out.
fn is_cgm(probe: &[u8]) -> bool {
    if starts_with(probe, b"BEGMF") {
        return true;
    }

    read_be16(probe, 0).is_some_and(|element| element & 0xFFE0 == 0x0020)
}

/// Whether the front of a file is a ClarisWorks document: the version byte of the format,
/// and the four characters every file the application wrote carries beside it.
fn is_clarisworks(probe: &[u8]) -> bool {
    probe.first().is_some_and(|version| (1..=6).contains(version))
        && (at(probe, 4, b"BOBO") || at(probe, 4, b"CWKJ"))
}

/// Whether the front of a file is a dBASE table, which carries no magic at all: what is
/// asked is the version byte the format registers, a date in the three bytes after it that
/// is a date, the length of the header the records follow, and the terminator that header
/// ends with.
fn is_dbase_table(probe: &[u8]) -> bool {
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
fn is_dxf(probe: &[u8]) -> bool {
    if starts_with(probe, b"AutoCAD Binary DXF") {
        return true;
    }

    starts_with(probe, b"0\r\nSECTION") || starts_with(probe, b"0\nSECTION")
}

/// Whether the front of a file is a PageMaker document of one version: the tag the format
/// writes six bytes in, and the version word a hundred and ten bytes in.
fn is_pagemaker(probe: &[u8], version: &[u8]) -> bool {
    (at(probe, 6, &[0xFF, 0x99]) || at(probe, 6, &[0x99, 0xFF])) && at(probe, 110, version)
}

/// Whether the front of a file is a WordPerfect container holding the kind of content the
/// caller asks for: the signature every file of the family carries, the product byte that
/// follows it, and the file-type byte that says what is inside.
fn is_wordperfect(probe: &[u8], file_type: u8) -> bool {
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
fn is_hevc_stream(probe: &[u8]) -> bool {
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
fn is_vvc_stream(probe: &[u8]) -> bool {
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
fn is_av1_obu_stream(probe: &[u8]) -> bool {
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
fn is_avs_sequence(probe: &[u8]) -> bool {
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
fn is_dirac_stream(probe: &[u8]) -> bool {
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
        Some(rest) if rest.len() >= 13 => {
            at(rest, 0, b"BBCD") && read_be32(rest, 9) == Some(next)
        }
        _ => true,
    }
}

/// Whether the front of a file is an EVC stream, which is not start-code framed: a run of
/// units, each a four-byte length and a two-byte header — and it is the run that is asked
/// for, because one length and one header are a shape ordinary data has.
fn is_evc_stream(probe: &[u8]) -> bool {
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
fn is_dv_stream(probe: &[u8]) -> bool {
    (0..probe.len().saturating_sub(3))
        .take(1200)
        .any(|offset| at(probe, offset, &[0x1F, 0x07, 0x00, 0x3F]))
}

/// Whether the front of a file is a VC-1 test stream: the byte the format writes four bytes
/// in, and the marker that closes the first frame's header.
fn is_vc1_test_stream(probe: &[u8]) -> bool {
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
fn is_yop(probe: &[u8]) -> bool {
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
fn is_rsd(probe: &[u8]) -> bool {
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
fn is_c93(probe: &[u8]) -> bool {
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
fn is_cdg(probe: &[u8]) -> bool {
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
fn is_cdxl(probe: &[u8]) -> bool {
    if probe.len() < 32 {
        return false;
    }

    let (Some(&kind), Some(&planes), Some(&reserved)) = (probe.first(), probe.get(19), probe.get(18))
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
fn is_moflex(probe: &[u8]) -> bool {
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
fn is_h261_picture(probe: &[u8]) -> bool {
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
fn is_h263_picture(probe: &[u8]) -> bool {
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
fn nal_body(probe: &[u8]) -> Option<&[u8]> {
    let start = probe.iter().position(|byte| *byte != 0)?;

    if start < 2 || probe.get(start) != Some(&1) {
        return None;
    }

    probe.get(start + 1..)
}

/// Whether the front of a file is an H.264 elementary stream: a start code, then the
/// sequence parameter set a stream of one opens with.
fn is_h264_stream(probe: &[u8]) -> bool {
    nal_body(probe).is_some_and(|body| body.first().is_some_and(|header| header & 0x1F == 7))
}

/// `count` bits of `probe` beginning `start` bits into it, most significant bit first.
fn read_bits(probe: &[u8], start: usize, count: usize) -> Option<u32> {
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
fn read_be16(probe: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_be_bytes(
        probe.get(offset..offset + 2)?.try_into().ok()?,
    ))
}

/// And the little-endian one.
fn read_le16(probe: &[u8], offset: usize) -> Option<u16> {
    Some(u16::from_le_bytes(
        probe.get(offset..offset + 2)?.try_into().ok()?,
    ))
}

/// The four bytes at `offset` as a number, and nothing where the probe is too short.
fn read_be32(probe: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_be_bytes(
        probe.get(offset..offset + 4)?.try_into().ok()?,
    ))
}

/// And the little-endian one.
fn read_le32(probe: &[u8], offset: usize) -> Option<u32> {
    Some(u32::from_le_bytes(
        probe.get(offset..offset + 4)?.try_into().ok()?,
    ))
}

/// Whether `probe` holds `needle` at `offset`, and nothing where it is too short to tell.
fn at(probe: &[u8], offset: usize, needle: &[u8]) -> bool {
    probe
        .get(offset..offset + needle.len())
        .is_some_and(|window| window == needle)
}

/// Whether `probe` opens with `prefix`.
fn starts_with(probe: &[u8], prefix: &[u8]) -> bool {
    at(probe, 0, prefix)
}

/// Whether `probe` holds `needle` anywhere in it.
fn contains(probe: &[u8], needle: &[u8]) -> bool {
    probe.windows(needle.len()).any(|window| window == needle)
}

/// One of FFmpeg's own formats: a name the video list carries that no signature above
/// names — a container whose header is not one of the common ones, a raw stream, a
/// capture, game or camera format.
const fn video(extension: &'static str) -> (&'static str, PreviewType) {
    (extension, PreviewType::Videos)
}

/// One of the documents the render engine draws: a name the `[libre]` list carries that
/// no signature above names.
const fn engine(extension: &'static str) -> (&'static str, PreviewType) {
    (extension, PreviewType::Libre)
}

/// The names no table above answers for, and the kind each of them belongs to.
///
/// What is left once the two tables above have answered for everything they can is short,
/// and every entry in it is a name a head cannot settle:
///
/// * a document that is a container and says nothing about itself inside the probe: an
///   iWork document (`key`, `numbers`, `pages`), a StarOffice 5 one (`sda`, `sdc`, `sdd`,
///   `sdw`), a Visio one (`vsd`), a Publisher one, the Pocket Word one, a Zoner one, a
///   gzip-wrapped AbiWord or Gnumeric document, or a Word for the Macintosh file that is
///   one of the older shapes;
/// * a format whose signature is its own text, which the text lists answer for: a flat
///   OpenDocument, a Visio `.vdx`;
/// * a name without any probe of its own — `mvi` and `mxg`, which FFmpeg reads by name,
///   `psp` and `vw`, which no demuxer of its registers, and `cin`, whose demuxer is gone;
/// * and the two TiVo names, whose chunk headers sit a hundred and twenty-eight kilobytes
///   apart, which is not a question the head of a file can be asked.
///
/// A name here answers with the kind, because a name is the only thing left to answer with,
/// and it answers whether or not a list in `config.ini` still holds it: what a format *is*
/// does not change with a setting, and a user who wants none of these previewed has the
/// kind's own switch in the tray's `Preview Types`.
static KIND_BY_NAME: &[(&str, PreviewType)] = &[
    // The video names nothing here can ask the bytes about.
    video("cin"),
    video("flm"),
    video("mvi"),
    video("mxg"),
    video("psp"),
    video("ty"),
    video("ty+"),
    video("vw"),
    // And the documents of the render engine's list that are containers, text, or both.
    engine("abw"),
    engine("fodg"),
    engine("fodp"),
    engine("fodt"),
    engine("gnm"),
    engine("gnumeric"),
    engine("key"),
    engine("mw"),
    engine("numbers"),
    engine("pages"),
    engine("pdb"),
    engine("psw"),
    engine("pub"),
    engine("sda"),
    engine("sdc"),
    engine("sdd"),
    engine("sdw"),
    engine("vdx"),
    engine("vsd"),
    engine("vsdm"),
    engine("vsdx"),
    engine("vstx"),
    engine("zabw"),
    engine("zmf"),
];

/// The names in the table above that are more than one format's, and the question asked of
/// the bytes before any of them is answered.
///
/// A name is here where the format its engine reads shares its spelling with a format
/// nothing here previews, and where the two can be told apart from the front of the file.
/// What such a name answers is the format the engine reads, or nothing at all: a file that
/// is the other format is a file no kind of this app previews, and starting an engine that
/// can only turn it down is the one thing this is for.
static GUARDS: &[(&str, Guard)] = &[("pdb", palm_ebook_or_program_database)];

/// The question a name in [`GUARDS`] is asked of the front of a file, answered in the same
/// three terms the tables answer in: the kind that previews what the bytes hold, nothing
/// where no kind of this app previews them, and no opinion where the bytes do not say.
type Guard = fn(&[u8]) -> Content;

/// What a `.pdb` is, which is the one name in the table that holds two formats.
///
/// The name is written for two things that share nothing. One is the Palm OS database —
/// the AportisDoc and its kin, the ebooks the render engine's own filters read, whose
/// header names the application that wrote them and the records that follow. The other is
/// the Microsoft program database a compiler writes beside its binaries, which is an *MSF*
/// container of debug information and no kind of document at all — and which is the one a
/// developer's folders are full of, and the reason this name is not answered by itself.
///
/// What comes back is the ebook where the file's own header is a Palm OS one, and nothing
/// at all where it is anything else, the program database included. Nothing is left to the
/// name here, deliberately: a `.pdb` that is neither of the two is a file this app has no
/// reader for either way, and the engine is not asked about one.
fn palm_ebook_or_program_database(probe: &[u8]) -> Content {
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
fn is_program_database(probe: &[u8]) -> bool {
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
fn is_palm_database(probe: &[u8]) -> bool {
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
fn is_palm_tag(tag: &[u8]) -> bool {
    tag.first().is_some_and(|byte| byte.is_ascii_alphabetic())
        && tag.iter().all(|byte| byte.is_ascii_graphic())
}

/// What the name a file carries answers for it, where the tables above named nothing.
///
/// It is the last of the three questions this module asks of a file: the bytes first,
/// through the two signature tables, and the name after them, through the table above —
/// whose own exception, a name that is two formats, is asked of the bytes in between (see
/// [`GUARDS`]). `None` where the name is not one the table holds, which is where a file's
/// kind is left to the lists, as it has always been.
fn kind_by_name(extension: &str, probe: &[u8]) -> Option<Content> {
    if let Some((_, guard)) = GUARDS
        .iter()
        .find(|(name, _)| name.eq_ignore_ascii_case(extension))
    {
        return Some(guard(probe));
    }

    KIND_BY_NAME
        .iter()
        .find(|(name, _)| name.eq_ignore_ascii_case(extension))
        .map(|(_, kind)| Content::Kind(*kind))
}

/// What the format the bytes named means for this app, for one file.
fn classify(path: &Path, names: &[&str], config: &AppConfig) -> Content {
    // A file with no extension at all has no name for the content to disagree with, and
    // one whose own name is among the names the content answered with is the two agreeing:
    // there is nothing to override either way.
    let Some(own) = own_extension(path) else {
        return Content::Unknown;
    };

    if names.iter().any(|name| name.eq_ignore_ascii_case(&own)) {
        return Content::Unknown;
    }

    // A name the ImageMagick engine is asked about is answered as that kind where the bytes
    // named the container its format is written in rather than a picture of its own: a
    // camera raw that is a TIFF is a `.nef`, a `.cr2`, an `.arw`, a `.dng` or a `.pef`, and
    // the name is the camera's own answer about what is inside the box (see
    // `names_the_container_of_a_raw`). Every other name is answered below, which is where a
    // `.tif` is taken for the picture it is.
    if names_the_container_of_a_raw(names)
        && crate::formats::magick_formats::matches_magick_list(path, &config.magick_extensions)
    {
        return Content::Kind(PreviewType::Magick);
    }

    match kind_claiming(names, config) {
        Some(kind) => Content::Kind(kind),
        None => Content::Foreign,
    }
}

/// The kind of this app that claims one of the names the content answered with, asked in
/// the order every other classification of a file is asked in.
///
/// Each name is asked of the same list function the hook, the loader and the layout ask,
/// by a name of this module's own making (`content.<extension>`) that stands for the file
/// and is never opened: a list reads the extension and nothing else, so the file the
/// caller holds is not touched a second time. Nothing claimed by any list is a format this
/// app does not preview — see [`Content::Foreign`].
fn kind_claiming(names: &[&str], config: &AppConfig) -> Option<PreviewType> {
    for name in names {
        let named = PathBuf::from(format!("content.{name}"));

        if crate::formats::video_formats::claims_video_name(&named, &config.video_extensions) {
            return Some(PreviewType::Videos);
        }

        // A PDF is a kind of its own rather than a list: the one name is the whole of what
        // the PDF path claims, and what a page is read from is the file's own header.
        if name.eq_ignore_ascii_case("pdf") {
            return Some(PreviewType::Pdf);
        }

        if crate::formats::archive_formats::matches_archive_list(&named, &config.archive_extensions) {
            return Some(PreviewType::Archives);
        }

        if crate::formats::office_formats::matches_office_list(&named, &config.office_extensions) {
            return Some(PreviewType::Office);
        }

        if crate::formats::libre_formats::matches_libre_list(&named, &config.libre_extensions) {
            return Some(PreviewType::Libre);
        }

        if crate::formats::magick_formats::matches_magick_list(&named, &config.magick_extensions) {
            return Some(PreviewType::Magick);
        }

        if crate::formats::design_formats::matches_design_list(&named, &config.design_extensions) {
            return Some(PreviewType::Design);
        }

        if crate::formats::vector_formats::matches_vector_list(&named, &config.vector_extensions) {
            return Some(PreviewType::Vector);
        }

        if crate::formats::text_formats::matches_text_lists(
            &named,
            &config.text_extensions,
            &config.text_names,
        ) {
            return Some(PreviewType::Text);
        }

        if crate::formats::font_formats::matches_font_list(&named, &config.font_extensions) {
            return Some(PreviewType::Fonts);
        }

        if crate::formats::image_formats::matches_image_list(&named, &config.image_extensions) {
            return Some(PreviewType::Images);
        }
    }

    None
}

/// The extension the file is named by, in the form the lists carry.
fn own_extension(path: &Path) -> Option<String> {
    path.extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| extension.to_lowercase())
}

/// The file an answer belongs to: its path, and the version of it that was read.
fn answer_key(path: &Path) -> AnswerKey {
    let metadata = std::fs::metadata(path).ok();

    AnswerKey {
        path: path.to_path_buf(),
        modified: metadata.as_ref().and_then(|metadata| metadata.modified().ok()),
        len: metadata.as_ref().map(|metadata| metadata.len()).unwrap_or(0),
    }
}

/// The file and the version of it an answer was read from. A file saved again is a file to
/// ask about again, and what says so is what says it everywhere else in this app: when it
/// was last written and what it weighed.
#[derive(Clone, PartialEq, Eq, Hash)]
struct AnswerKey {
    path: PathBuf,
    modified: Option<SystemTime>,
    len: u64,
}

/// What has been asked and answered, for the hovers of this run.
static ANSWERS: Lazy<Mutex<HashMap<AnswerKey, Content>>> = Lazy::new(|| Mutex::new(HashMap::new()));

#[cfg(test)]
mod tests {
    use super::*;

    /// A file of one name holding the front of a format, answered the way a hover asks it:
    /// the bytes through the two signature tables, and the name through the third.
    fn classified(name: &str, content: &[u8]) -> Content {
        if let Some(names) = detected_names(content) {
            return classify(Path::new(name), names, &AppConfig::default());
        }

        own_extension(Path::new(name))
            .and_then(|extension| kind_by_name(&extension, content))
            .unwrap_or(Content::Unknown)
    }

    /// The database header a Palm OS document opens with: the name it is filed under, the
    /// four-character type and creator of the application that wrote it, and the count of
    /// the records that follow.
    fn palm_database(name: &str, kind: &[u8; 4], creator: &[u8; 4]) -> Vec<u8> {
        let mut header = vec![0u8; 78 + 8];
        header[..name.len()].copy_from_slice(name.as_bytes());
        header[60..64].copy_from_slice(kind);
        header[64..68].copy_from_slice(creator);
        header[76..78].copy_from_slice(&1u16.to_be_bytes());
        header
    }

    /// A name and a content that disagree are answered with the kind the content belongs
    /// to, which is what a file is handed to an engine by.
    #[test]
    fn a_disagreement_is_answered_with_the_kind_the_content_belongs_to() {
        assert_eq!(
            classified("film.docx", b"\x00\x00\x00\x20ftypisom"),
            Content::Kind(PreviewType::Videos),
            "an MP4 under a document's name is a video"
        );
        assert_eq!(
            classified("drawing.cdr", b"\x89PNG\x0D\x0A\x1A\x0A"),
            Content::Kind(PreviewType::Images),
            "a picture under a drawing's name is a picture"
        );
        assert_eq!(
            classified("sheet.wpd", b"GIF89a"),
            Content::Kind(PreviewType::Images),
            "and a GIF is a picture wherever it is found"
        );
        // A transport stream is one of the formats only this app's table names: the
        // common table has no entry for one at all. What such a file looks like is a sync
        // byte every packet size apart, which is the probe the text lists share.
        let transport_stream = {
            let mut probe = vec![0u8; 188 * 4];
            for packet in 0..4 {
                probe[packet * 188] = 0x47;
            }
            probe
        };
        assert_eq!(
            classified("film.docx", &transport_stream),
            Content::Kind(PreviewType::Videos),
            "a transport stream under a document's name is a video"
        );
        assert_eq!(
            classified("letter.docx", b"%!PS-Adobe-3.0"),
            Content::Kind(PreviewType::Vector),
            "and PostScript is the drawing the vector list reads"
        );
    }

    /// Where the content and the name agree there is nothing to override: the file has
    /// the kind it was always going to have, whether the name it carries is the first the
    /// format is known by or one of the others.
    #[test]
    fn agreement_is_no_opinion() {
        assert_eq!(classified("shot.png", b"\x89PNG\x0D\x0A\x1A\x0A"), Content::Unknown);
        assert_eq!(classified("film.mkv", &[0x1A, 0x45, 0xDF, 0xA3]), Content::Unknown);
        assert_eq!(classified("film.mp4", b"\x00\x00\x00\x20ftypisom"), Content::Unknown);
        assert_eq!(
            classified("picture.mjpeg", &[0xFF, 0xD8, 0xFF, 0xE0]),
            Content::Unknown,
            "a motion JPEG is a video by name as well as a picture by content"
        );
    }

    /// A box is not an answer, which is the half of the signature tables that is
    /// deliberately not in them: an OpenDocument and an Office package are both zips, and
    /// neither table answers for one, so the name the file was given is what decides —
    /// which for a name the third table holds is the engine that draws it, and for every
    /// other name is the list that claims it.
    #[test]
    fn a_box_is_left_to_the_name() {
        for name in ["letter.odt", "report.docx", "bundle.zip"] {
            assert_eq!(
                classified(name, b"PK\x03\x04\x14\x00\x00\x00"),
                Content::Unknown,
                "`{name}` is a zip, and a zip is not a kind"
            );
        }

        // An OLE compound file — every `.doc`, `.xls` and `.ppt` — is a box of the same
        // kind, and the name is what is answered with.
        assert_eq!(
            classified("letter.doc", b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"),
            Content::Unknown
        );
    }

    /// A format whose head is nothing a signature names is answered by the name it is
    /// written under, which is what the table of names is for: the names no probe can be
    /// asked about, and the files whose name is all there is to route them by.
    #[test]
    fn a_format_no_signature_names_is_answered_by_the_name_it_carries() {
        // A TiVo stream, whose chunk headers are a hundred and twenty-eight kilobytes
        // apart, and a Cineon file, whose demuxer is gone: two names no head can settle.
        assert_eq!(
            classified("recording.ty", b"\x00"),
            Content::Kind(PreviewType::Videos)
        );
        assert_eq!(
            classified("clip.ty+", b"\x00"),
            Content::Kind(PreviewType::Videos),
            "a name the lists carry and no signature does is answered as it is written"
        );
        assert_eq!(
            classified("film.cin", b"\x01\x02\x03\x04"),
            Content::Kind(PreviewType::Videos)
        );

        // A flat OpenDocument is a document whose signature is its own markup, and nothing
        // is asked of one: the list the name is written for is what decides.
        assert_eq!(
            classified("sheet.fodt", b"<?xml version=\"1.0\" encoding=\"UTF-8\"?>"),
            Content::Kind(PreviewType::Libre)
        );

        // A WordPerfect document that writes its header is the render engine's by its own
        // bytes. One that writes no header is a file no table here has an opinion about,
        // which is the state every name in the table below is in.
        assert_eq!(
            classified("letter.docx", b"\xFFWPC\x00\x00\x00\x00\x01\x0A"),
            Content::Kind(PreviewType::Libre),
            "a WordPerfect document under a document's name is the engine's"
        );
        assert_eq!(
            classified("letter.wpd", b"nothing in here is a signature"),
            Content::Unknown
        );

        // A name whose format is a package is a zip in the bytes and a document in the
        // name — an iWork deck is one — and the table answers with the engine's kind,
        // which is what the list it is written for would have said.
        assert_eq!(
            classified("deck.key", b"PK\x03\x04\x14\x00\x00\x00"),
            Content::Kind(PreviewType::Libre)
        );
    }

    /// A probe of `needle` at an offset, for the formats whose marker is not at the front
    /// of the file.
    fn padded(offset: usize, needle: &[u8]) -> Vec<u8> {
        let mut probe = vec![0u8; offset + needle.len()];
        probe[offset..].copy_from_slice(needle);

        probe
    }

    /// The head of a package that declares its own type: the first entry of the archive is
    /// a stored `mimetype` entry, and the type itself follows the thirty-byte header of
    /// that entry and the eight characters of its name.
    fn declared_package(mime: &str) -> Vec<u8> {
        let mut probe = vec![0u8; 38];

        probe[..4].copy_from_slice(b"PK\x03\x04");
        probe[26..28].copy_from_slice(&8u16.to_le_bytes());
        probe[30..38].copy_from_slice(b"mimetype");
        probe.extend_from_slice(mime.as_bytes());
        // The next record of the archive, which is what a reader of the declaration finds
        // where the type ends.
        probe.extend_from_slice(b"PK\x03\x04");

        probe
    }

    /// Every format the signature table carries is answered by its own head, under a name
    /// that belongs to another kind. That is what the table is for: a hover meets files
    /// that are not named for what they are, and the bytes are what settle it. The name in
    /// each of these is one whose list would have sent the file somewhere else — to Word,
    /// to the picture decoder, to the text preview — so what is asserted is that the
    /// content wins.
    #[test]
    fn every_format_the_table_carries_is_answered_by_its_own_head() {
        // ---------------------------------------------------------------- pictures
        assert_eq!(
            classified("film.docx", b"DDS \x7C\x00\x00\x00"),
            Content::Kind(PreviewType::Images),
            "a DirectDraw Surface"
        );
        assert_eq!(
            classified("film.docx", &[0x76, 0x2F, 0x31, 0x01]),
            Content::Kind(PreviewType::Images),
            "an OpenEXR picture"
        );
        assert_eq!(
            classified("film.docx", b"#?RGBE\nFORMAT=32-bit_rle_rgbe\n"),
            Content::Kind(PreviewType::Images),
            "a Radiance picture of the older header"
        );
        assert_eq!(
            classified("film.docx", b"#?RADIANCE\nFORMAT=32-bit_rle_rgbe\n"),
            Content::Kind(PreviewType::Images),
            "and one of the newer"
        );
        assert_eq!(
            classified("film.docx", b"farbfeld\x00\x00\x00\x02\x00\x00\x00\x02"),
            Content::Kind(PreviewType::Images),
            "a farbfeld picture"
        );
        assert_eq!(
            classified("film.docx", b"qoif\x00\x00\x00\x02\x00\x00\x00\x02\x03\x00"),
            Content::Kind(PreviewType::Images),
            "a Quite OK Image"
        );
        assert_eq!(
            classified("film.docx", b"P5 2 2 255\n\x00\x00\x00\x00"),
            Content::Kind(PreviewType::Images),
            "a Netpbm picture"
        );
        assert_eq!(
            classified("film.docx", b"P7\nWIDTH 1\nHEIGHT 1\nDEPTH 3\nMAXVAL 255\nENDHDR\n"),
            Content::Kind(PreviewType::Images),
            "a PAM picture, which names its own fields"
        );
        assert_eq!(
            classified("film.docx", &padded(2048, b"PCD_IPI")),
            Content::Kind(PreviewType::Libre),
            "a Photo CD image pac"
        );
        assert_eq!(
            classified("film.docx", &[0x0A, 0x05, 0x01, 0x08, 0, 0, 0, 0, 9, 9, 9, 9]),
            Content::Kind(PreviewType::Libre),
            "a PCX picture"
        );
        assert_eq!(
            classified("film.docx", &[0x59, 0xA6, 0x6A, 0x95]),
            Content::Kind(PreviewType::Libre),
            "a Sun raster"
        );
        assert_eq!(
            classified("film.docx", &padded(522, &[0x00, 0x11, 0x02, 0xFF])),
            Content::Kind(PreviewType::Libre),
            "a QuickDraw PICT, whose version operator is past the header it is drawn with"
        );
        // Two pictures the common table is what names, which this app's own table has
        // nothing to say about: all that is asked is that the answer it gives is used.
        assert_eq!(
            classified("film.docx", &[0xFF, 0x0A, 0x00]),
            Content::Kind(PreviewType::Images),
            "a JPEG XL picture"
        );
        assert_eq!(
            classified("film.docx", &declared_package("image/openraster")),
            Content::Kind(PreviewType::Design),
            "an OpenRaster project, named by the type it declares"
        );

        // ---------------------------------------------------------------- drawings
        assert_eq!(
            classified("film.docx", &{
                let mut probe = vec![0u8; 88];
                probe[..4].copy_from_slice(&[0x01, 0x00, 0x00, 0x00]);
                probe[4..8].copy_from_slice(&88u32.to_le_bytes());
                probe[40..44].copy_from_slice(b" EMF");
                probe[44..48].copy_from_slice(&[0x00, 0x00, 0x01, 0x00]);
                probe
            }),
            Content::Kind(PreviewType::Vector),
            "an enhanced metafile"
        );
        assert_eq!(
            classified("film.docx", &[0xD7, 0xCD, 0xC6, 0x9A, 0x00, 0x00]),
            Content::Kind(PreviewType::Vector),
            "a placeable Windows metafile"
        );
        assert_eq!(
            classified("film.docx", &[0x01, 0x00, 0x09, 0x00, 0x00, 0x01]),
            Content::Kind(PreviewType::Vector),
            "and a bare one"
        );

        // --------------------------------------------------------------- documents
        assert_eq!(
            classified("film.docx", b"@CT "),
            Content::Kind(PreviewType::Libre),
            "a Text602 document"
        );
        assert_eq!(
            classified("film.docx", b"RIFF\x00\x00\x00\x00CDR6"),
            Content::Kind(PreviewType::Libre),
            "a CorelDRAW drawing, in the RIFF container of the older versions"
        );
        assert_eq!(
            classified("film.docx", b"RIFF\x00\x00\x00\x00CMX1"),
            Content::Kind(PreviewType::Libre),
            "a Corel presentation exchange"
        );
        assert_eq!(
            classified("film.docx", b"BEGMF"),
            Content::Kind(PreviewType::Libre),
            "a Computer Graphics Metafile in its clear-text encoding"
        );
        assert_eq!(
            classified("film.docx", &[0x00, 0x20, 0x00, 0x00]),
            Content::Kind(PreviewType::Libre),
            "and one in the binary encoding"
        );
        assert_eq!(
            classified("film.docx", b"\x05\x07\x00\x00BOBO"),
            Content::Kind(PreviewType::Libre),
            "a ClarisWorks document"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = vec![0u8; 32];
                probe[0] = 0x03;
                probe[2] = 0x01;
                probe[3] = 0x0F;
                probe[8] = 0x41;
                probe
            }),
            Content::Kind(PreviewType::Libre),
            "a dBASE table, which is read for its version byte and its date"
        );
        assert_eq!(
            classified("film.docx", b"0\r\nSECTION\r\n"),
            Content::Kind(PreviewType::Libre),
            "a DXF drawing written as text"
        );
        assert_eq!(
            classified("film.docx", b"AutoCAD Binary DXF\r\n\x1a\x00"),
            Content::Kind(PreviewType::Libre),
            "and one written as binary"
        );
        assert_eq!(
            classified("film.docx", b"HWP Document File"),
            Content::Kind(PreviewType::Libre),
            "a Hangul document of the version that writes its name"
        );
        assert_eq!(
            classified("film.docx", b"WordPro"),
            Content::Kind(PreviewType::Libre),
            "a Lotus Word Pro document"
        );
        assert_eq!(
            classified("film.docx", &padded(2, &[0xD3, 0xA8, 0xA8])),
            Content::Kind(PreviewType::Libre),
            "an OS/2 metafile"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = padded(110, &[0x00, 0x06]);
                probe[6] = 0xFF;
                probe[7] = 0x99;
                probe
            }),
            Content::Kind(PreviewType::Libre),
            "a PageMaker 6 document"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = padded(110, &[0x32, 0x06]);
                probe[6] = 0xFF;
                probe[7] = 0x99;
                probe
            }),
            Content::Kind(PreviewType::Libre),
            "and a PageMaker 6.5 one, which the version word tells from it"
        );
        assert_eq!(
            classified("film.docx", b"VCLMTF"),
            Content::Kind(PreviewType::Libre),
            "a StarView metafile"
        );
        assert_eq!(
            classified("film.docx", b"ID;P"),
            Content::Kind(PreviewType::Libre),
            "a SYLK spreadsheet"
        );
        assert_eq!(
            classified("film.docx", b"\xFFWPC\x00\x00\x00\x00\x01\x0A"),
            Content::Kind(PreviewType::Libre),
            "a WordPerfect document"
        );
        assert_eq!(
            classified("film.docx", b"\xFFWPC\x00\x00\x00\x00\x01\x16"),
            Content::Kind(PreviewType::Libre),
            "and a WordPerfect drawing, which the file-type byte tells from it"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = vec![0u8; 22];
                probe[..2].copy_from_slice(&[0x01, 0xFE]);
                probe[20..22].copy_from_slice(&[0xD0, 0x02]);
                probe
            }),
            Content::Kind(PreviewType::Libre),
            "a Microsoft Works document"
        );
        assert_eq!(
            classified("film.docx", b"\x31\xBE\x00\x00"),
            Content::Kind(PreviewType::Libre),
            "a Windows Write document"
        );
        assert_eq!(
            classified("film.docx", b"\xFE\x37\x00\x1C\x00\x00\x00\x00"),
            Content::Kind(PreviewType::Libre),
            "a Word for the Macintosh document"
        );
        assert_eq!(
            classified("film.docx", b"\x09\x08\x08\x00\x00\x05"),
            Content::Kind(PreviewType::Libre),
            "a flat BIFF workbook"
        );
        for (name, head) in [
            ("123", b"\x00\x00\x1A\x00\x03\x10".as_slice()),
            ("wk3", b"\x00\x00\x1A\x00\x00\x10"),
            ("wk4", b"\x00\x00\x1A\x00\x02\x10"),
            ("wk1", b"\x00\x00\x02\x00\x06\x04"),
            ("wks", b"\x00\x00\x02\x00\x04\x04"),
            ("wb2", b"\x00\x00\x02\x00\x02\x10"),
            ("wq1", b"\x00\x00\x02\x00\x20\x51"),
            ("wq2", b"\x00\x00\x02\x00\x21\x51"),
        ] {
            assert_eq!(
                classified("film.docx", head),
                Content::Kind(PreviewType::Libre),
                "a `.{name}` spreadsheet of the years before the current ones"
            );
        }

        // ------------------------------------------------------------------ videos
        assert_eq!(
            classified("film.docx", b"\x00\x00\x00\x20ftypisom"),
            Content::Kind(PreviewType::Videos),
            "an ISO base media file"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = vec![0u8; 188 * 4];
                for packet in 0..4 {
                    probe[packet * 188] = 0x47;
                }
                probe
            }),
            Content::Kind(PreviewType::Videos),
            "a transport stream"
        );
        assert_eq!(
            classified("film.docx", b"\x00\x00\x01\xBA"),
            Content::Kind(PreviewType::Videos),
            "an MPEG program stream"
        );
        assert_eq!(
            classified("film.docx", b"\x00\x00\x01\xB3"),
            Content::Kind(PreviewType::Videos),
            "an MPEG elementary stream"
        );
        assert_eq!(
            classified("film.docx", b"\x00\x00\x00\x01\x67\x42\x00\x1E"),
            Content::Kind(PreviewType::Videos),
            "a raw H.264 stream"
        );
        assert_eq!(
            classified("film.docx", b"\x00\x00\x00\x01\x40\x01\x0C\x01\xFF\xFF"),
            Content::Kind(PreviewType::Videos),
            "a raw H.265 stream, whose video parameter set carries the field reserved to ones"
        );
        assert_eq!(
            classified("film.docx", b"\x00\x00\x00\x01\x00\x79"),
            Content::Kind(PreviewType::Videos),
            "a raw H.266 stream"
        );
        assert_eq!(
            classified("film.docx", b"\x12\x00\x0A\x02\x01\x02"),
            Content::Kind(PreviewType::Videos),
            "an AV1 stream, whose units have to chain"
        );
        assert_eq!(
            classified("film.docx", b"\x00\x00\x01\xB0\x20\x10\x0F\x00\x21\xC0"),
            Content::Kind(PreviewType::Videos),
            "an AVS sequence header, which is told from an MPEG-4 one by its profile"
        );
        assert_eq!(
            classified("film.docx", b"\x00\x00\x01\x0F\xC0"),
            Content::Kind(PreviewType::Videos),
            "a VC-1 sequence header of the advanced profile"
        );
        assert_eq!(
            classified("film.docx", b"\x00\x00\x01\x0F\x80"),
            Content::Unknown,
            "and one of a profile that carries no header of its own is left to the name"
        );
        assert_eq!(
            classified("film.docx", b"BBCD\x00\x00\x00\x00\x0D\x00\x00\x00\x00BBCD\x00\x00\x00\x00\x00\x00\x00\x00\x0D"),
            Content::Kind(PreviewType::Videos),
            "a Dirac or VC-2 stream, whose two units agree about where the first ended"
        );
        assert_eq!(
            classified("film.docx", b"aPv1"),
            Content::Kind(PreviewType::Videos),
            "an APV frame"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = Vec::new();
                for _ in 0..3 {
                    probe.extend_from_slice(&[0x00, 0x00, 0x00, 0x02, 0x32, 0x01]);
                }
                probe
            }),
            Content::Kind(PreviewType::Videos),
            "an EVC stream, whose units are length-prefixed"
        );
        assert_eq!(
            classified("film.docx", &padded(80, &[0x1F, 0x07, 0x00, 0x3F])),
            Content::Kind(PreviewType::Videos),
            "a DV stream"
        );
        assert_eq!(
            classified("film.docx", b"DKIF"),
            Content::Kind(PreviewType::Videos),
            "an IVF stream"
        );
        assert_eq!(
            classified("film.docx", b"\x06\x0E\x2B\x34\x02\x05\x01\x01\x0D\x01\x02\x01\x01\x02"),
            Content::Kind(PreviewType::Videos),
            "an MXF, and the `.imx` essence that is one"
        );
        assert_eq!(
            classified("film.docx", b"NUT/MULTI"),
            Content::Kind(PreviewType::Videos),
            "a NUT stream"
        );
        assert_eq!(
            classified("film.docx", b"YUV4MPEG2 W2 H2 F25:1 Ip A0:0 C420\n"),
            Content::Kind(PreviewType::Videos),
            "a YUV4MPEG2 stream"
        );
        assert_eq!(
            classified("film.docx", b"BIK"),
            Content::Kind(PreviewType::Videos),
            "a Bink movie"
        );
        assert_eq!(
            classified("film.docx", b"\x84\x10\xFF\xFF\xFF\xFF"),
            Content::Kind(PreviewType::Videos),
            "an Id RoQ movie"
        );
        assert_eq!(
            classified("film.docx", b"SMK2"),
            Content::Kind(PreviewType::Videos),
            "a Smacker movie"
        );
        assert_eq!(
            classified("film.docx", b"THP\x00"),
            Content::Kind(PreviewType::Videos),
            "a GameCube THP movie"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = vec![0u8; 36];
                probe[12..16].copy_from_slice(b"xobX");
                probe[16..20].copy_from_slice(&2u32.to_le_bytes());
                probe
            }),
            Content::Kind(PreviewType::Videos),
            "an Xbox XMV movie"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = vec![0u8; 20];
                probe[..2].copy_from_slice(b"YO");
                probe[2] = 1;
                probe[3] = 2;
                probe[6] = 1;
                probe[7] = 1;
                probe[18..20].copy_from_slice(&920u16.to_le_bytes());
                probe
            }),
            Content::Kind(PreviewType::Videos),
            "a YOP movie, whose own fields are what its probe reads"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = vec![0u8; 20];
                probe[..4].copy_from_slice(b"RSD2");
                probe[8..12].copy_from_slice(&1u32.to_le_bytes());
                probe[16..20].copy_from_slice(&8000u32.to_le_bytes());
                probe
            }),
            Content::Kind(PreviewType::Videos),
            "a GameCube RSD stream"
        );
        assert_eq!(
            classified("film.docx", b".RMF\x00\x00"),
            Content::Kind(PreviewType::Videos),
            "a RealMedia stream"
        );
        assert_eq!(
            classified("film.docx", b".R1M\x00\x01\x01"),
            Content::Kind(PreviewType::Videos),
            "a recorded one"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = vec![0u8; 20];
                probe[..4].copy_from_slice(b"FILM");
                probe[16..20].copy_from_slice(b"FDSC");
                probe
            }),
            Content::Kind(PreviewType::Videos),
            "a Sega FILM movie"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = vec![0u8; 16];
                probe[4..6].copy_from_slice(&[0x01, 0xBC]);
                probe[10..16].copy_from_slice(&[0x00, 0x00, 0x00, 0x00, 0xE1, 0xE2]);
                probe
            }),
            Content::Kind(PreviewType::Videos),
            "a GXF stream"
        );
        assert_eq!(
            classified("film.docx", &[0x11, 0xD2, 0xD3, 0xAB, 0xBA, 0xA9, 0xCF, 0x11, 0x8E, 0xE6, 0x00, 0xC0, 0x0C, 0x20, 0x53, 0x65, 0x44]),
            Content::Kind(PreviewType::Videos),
            "an IFV stream"
        );
        assert_eq!(
            classified("film.docx", b"KDK\x00\x00"),
            Content::Kind(PreviewType::Videos),
            "a KUX movie, which is a Flash movie with a header of its own"
        );
        assert_eq!(
            classified("film.docx", b"Interplay MVE File\x1A\x00\x1A\x00"),
            Content::Kind(PreviewType::Videos),
            "an Interplay MVE movie"
        );
        assert_eq!(
            classified("film.docx", b"pmpm\x01\x00\x00\x00"),
            Content::Kind(PreviewType::Videos),
            "a PMP movie"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = vec![0u8; 36];
                probe[..4].copy_from_slice(b"CRID");
                probe[32..36].copy_from_slice(b"@UTF");
                probe
            }),
            Content::Kind(PreviewType::Videos),
            "a Scaleform movie, which FFmpeg has no demuxer for and `file` names"
        );
        assert_eq!(
            classified("film.docx", b"DAHUA"),
            Content::Kind(PreviewType::Videos),
            "a Dahua camera's stream"
        );
        assert_eq!(
            classified("film.docx", b"\x00abcVersion:Vivo/0"),
            Content::Kind(PreviewType::Videos),
            "a Vivo stream"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = vec![0u8; 28];
                probe[3] = 0xC5;
                probe[4..8].copy_from_slice(&8u32.to_le_bytes());
                probe[24..28].copy_from_slice(&[0x0C, 0x00, 0x00, 0x00]);
                probe
            }),
            Content::Kind(PreviewType::Videos),
            "a VC-1 test stream"
        );
        assert_eq!(
            classified("film.docx", &[0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x00]),
            Content::Kind(PreviewType::Videos),
            "a PlayStation STR stream"
        );
        assert_eq!(
            classified("film.docx", &[0xB7, 0xD8, 0x00, 0x20, 0x37, 0x49, 0xDA, 0x11, 0xA6, 0x4E, 0x00, 0x07, 0xE9, 0x5E, 0xAD, 0x8D]),
            Content::Kind(PreviewType::Videos),
            "a Windows recorded television stream"
        );
        assert_eq!(
            classified("film.docx", b"NSVf"),
            Content::Kind(PreviewType::Videos),
            "a Nullsoft stream"
        );

        // The formats whose probe scores a shape rather than a magic, each with the shape
        // FFmpeg's own demuxer scores: a block table, a periodic command byte, a header of
        // offsets, a table that has to chain, a picture start code twice over.
        assert_eq!(
            classified("film.docx", &[0x01, 0x00, 0x04, 0x01, 0x05, 0x00, 0x02, 0x01, 0x07, 0x00, 0x03, 0x01, 0x0A, 0x00, 0x01, 0x01]),
            Content::Kind(PreviewType::Videos),
            "an Interplay C93 movie"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = vec![0u8; 24 * 8];
                for packet in 0..8 {
                    probe[packet * 24] = 0x09;
                }
                probe
            }),
            Content::Kind(PreviewType::Videos),
            "a CD Graphics stream"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = vec![0u8; 32];
                probe[2..6].copy_from_slice(&1388u32.to_be_bytes());
                probe[14..16].copy_from_slice(&320u16.to_be_bytes());
                probe[16..18].copy_from_slice(&200u16.to_be_bytes());
                probe[19] = 6;
                probe[20..22].copy_from_slice(&256u16.to_be_bytes());
                probe[22..24].copy_from_slice(&100u16.to_be_bytes());
                probe[24..26].copy_from_slice(&11025u16.to_be_bytes());
                probe[26] = 15;
                probe
            }),
            Content::Kind(PreviewType::Videos),
            "a Commodore CDXL stream"
        );
        assert_eq!(
            classified("film.docx", &{
                let mut probe = vec![0u8; 26];
                probe[..2].copy_from_slice(b"L2");
                probe[12..14].copy_from_slice(&1u16.to_be_bytes());
                probe[14..18].copy_from_slice(&[0x00, 0x01, 0x00, 0x04]);
                probe
            }),
            Content::Kind(PreviewType::Videos),
            "a Moflex movie"
        );
        assert_eq!(
            classified("film.docx", b"\x00\x01\x00\x11\x22\x00\x01\x00"),
            Content::Kind(PreviewType::Videos),
            "an H.261 stream"
        );
        assert_eq!(
            classified("film.docx", b"\x00\x00\x80\x11\x00\x00\x80"),
            Content::Kind(PreviewType::Videos),
            "an H.263 stream"
        );

        // -------------------------------------------------------------------- text
        assert_eq!(
            classified("film.docx", b"{\\rtf1\\ansi"),
            Content::Kind(PreviewType::Text),
            "an RTF document, which the text preview shows"
        );
        assert_eq!(
            classified("film.docx", b"ttcf\x00\x01\x00\x00"),
            Content::Kind(PreviewType::Fonts),
            "a collection of fonts"
        );
        assert_eq!(
            classified("film.docx", b"%!PS-Adobe-3.0"),
            Content::Kind(PreviewType::Vector),
            "and a PostScript drawing"
        );
    }

    /// The packages that declare their own type are the one shape of container this module
    /// answers for, and what it reads is that declaration — which is what keeps the rule
    /// that a container is not a kind, because the type is the document's own answer rather
    /// than the box's.
    #[test]
    fn a_package_is_answered_by_the_type_it_declares() {
        assert_eq!(
            classified(
                "film.docx",
                &declared_package("application/vnd.oasis.opendocument.graphics")
            ),
            Content::Kind(PreviewType::Libre),
            "a drawing is the render engine's, whatever the file is called"
        );
        assert_eq!(
            classified(
                "film.docx",
                &declared_package("application/vnd.sun.xml.writer")
            ),
            Content::Kind(PreviewType::Libre),
            "and so is a StarOffice document of the XML generation"
        );
        assert_eq!(
            classified("photo.png", &declared_package("application/x-krita")),
            Content::Kind(PreviewType::Design),
            "while a Krita project is a design document"
        );

        // A template's type begins the name of the type it is a template of, and the record
        // that follows the declaration is what settles which of the two a file carries: a
        // `…graphics-template` package is not answered as `…graphics`.
        assert!(!is_declared_package(
            &declared_package("application/vnd.oasis.opendocument.graphics-template"),
            b"application/vnd.oasis.opendocument.graphics"
        ));
        assert!(is_declared_package(
            &declared_package("application/vnd.oasis.opendocument.graphics-template"),
            b"application/vnd.oasis.opendocument.graphics-template"
        ));

        // And a zip that declares nothing is a zip: an Office package, an OpenDocument of
        // the version that says `mimetype` rather than storing it as the first entry, and
        // any other archive are one answer, which is the name.
        assert_eq!(
            classified("letter.docx", b"PK\x03\x04\x14\x00\x00\x00"),
            Content::Unknown
        );
    }

    /// The name `.pdb` is two formats, and only one of them is a document: the Palm OS
    /// database the render engine's own filters read as an ebook, and the Microsoft
    /// program database a compiler writes beside its binaries. The engine is asked about
    /// one of them and never about the other.
    #[test]
    fn the_two_formats_one_name_holds_are_told_apart() {
        let ebook = palm_database("Huckleberry Finn", b"TEXt", b"REAd");
        assert_eq!(
            classified("book.pdb", &ebook),
            Content::Kind(PreviewType::Libre),
            "a Palm OS ebook is the document the engine draws"
        );

        // The program database is the one a developer's folders are full of, and it is no
        // kind of document at all: nothing is shown and no engine is started for it.
        let program_database = {
            let mut probe = b"Microsoft C/C++ MSF 7.00\r\n\x1aDS\x00\x00\x00".to_vec();
            probe.resize(1024, 0);
            probe
        };
        assert_eq!(
            classified("app.pdb", &program_database),
            Content::Foreign,
            "a program database starts no engine"
        );
        assert_eq!(
            classified("app.pdb", b"Microsoft C/C++ program database 2.00\r\n\x1aJG"),
            Content::Foreign,
            "and neither does one of the version before it"
        );

        // What the engine reads is the Palm document, and a file of the name that is
        // nothing of the sort is not left to the name to decide: it is a file this app has
        // no reader for, which is the answer to the program database as well.
        assert_eq!(classified("data.pdb", b"\x00\x01\x02\x03"), Content::Foreign);
        assert_eq!(
            classified("data.pdb", b"not a database of any kind"),
            Content::Foreign
        );
    }

    /// And a name the signature tables *do* answer is not in the table of names: `.ts` and
    /// `.mts` are the two a text list shares, and a file of one of those names that is not
    /// the transport stream is the TypeScript source it is named as — which is a question
    /// only the bytes can settle.
    #[test]
    fn a_name_a_signature_answers_is_not_answered_by_its_name() {
        for (extension, _) in KIND_BY_NAME {
            for signature in SIGNATURES {
                assert!(
                    !signature
                        .names
                        .iter()
                        .any(|name| name.eq_ignore_ascii_case(extension)),
                    "`{extension}` is answered by a signature, so the name must not answer"
                );
            }
        }

        assert_eq!(
            classified("app.ts", b"const x: number = 1;\n"),
            Content::Unknown,
            "a TypeScript file is left to the text list it is named under"
        );
    }

    /// What is in the table is the two engines' own lists, and the kind beside a name is
    /// the kind the lists themselves give it: a video name is a video, a name of the
    /// render engine's is the engine's, and a name neither list carries is not in the
    /// table at all. It is what keeps the table from drifting away from the lists — a name
    /// taken out of `libre_formats`, the way `swf` was, has to be taken out of here with
    /// it — and what keeps `dif`, which both lists carry, answered in the order every
    /// other question about a file is asked in.
    #[test]
    fn the_table_holds_the_names_the_lists_hold() {
        let config = AppConfig::default();

        for (extension, kind) in KIND_BY_NAME {
            let named = PathBuf::from(format!("content.{extension}"));

            if crate::formats::video_formats::claims_video_name(&named, &config.video_extensions) {
                assert_eq!(*kind, PreviewType::Videos, "`{extension}` is FFmpeg's name");
                continue;
            }

            if crate::formats::libre_formats::matches_libre_list(&named, &config.libre_extensions) {
                assert_eq!(*kind, PreviewType::Libre, "`{extension}` is the engine's name");
                continue;
            }

            panic!("`{extension}` is not a name either list carries");
        }
    }

    /// A format no kind of this app previews is a file with nothing to show, and no engine
    /// is started for it.
    #[test]
    fn a_format_no_kind_previews_is_answered_with_nothing() {
        let executable = {
            let mut probe = vec![0u8; 0x80];
            probe[0] = b'M';
            probe[1] = b'Z';
            probe[0x3C] = 0x40;
            probe[0x40..0x44].copy_from_slice(b"PE\x00\x00");
            probe
        };

        assert_eq!(
            classified("letter.docx", &executable),
            Content::Foreign,
            "an executable under a document's name starts no engine"
        );

        // A Unix program is one of those, and the answer for one is the same: nothing is
        // shown and nothing is started.
        let unix_program = {
            let mut probe = vec![0u8; 64];
            probe[..4].copy_from_slice(&[0x7F, b'E', b'L', b'F']);
            probe[4] = 0x02;
            probe[5] = 0x01;
            probe[6] = 0x01;
            probe
        };
        assert_eq!(
            classified("letter.docx", &unix_program),
            Content::Foreign,
            "and neither does a Unix one"
        );
        assert_eq!(
            classified("letter.docx", b"ID3\x04\x00\x00\x00\x00\x00\x00"),
            Content::Foreign,
            "a song is not a document either"
        );
    }

    /// A front nothing in either table names, under a name the third table does not hold
    /// either, is left to the name it has — which is the state a text file is in, and
    /// every format this app reads for itself and confirms nothing about.
    #[test]
    fn what_has_no_signature_is_left_to_the_name_it_has() {
        assert_eq!(
            classified("stream.h264", b"\x00\x01\x02\x03 not a format at all"),
            Content::Unknown,
            "a `.h264` that is not one is answered by its name like any other file"
        );
        assert_eq!(classified("notes.txt", b"just some text\n"), Content::Unknown);
        assert_eq!(
            classified("drawing.svg", b"<svg xmlns=\"http://www.w3.org/2000/svg\">"),
            Content::Unknown,
            "a document whose signature is its own text is left to the lists"
        );
    }

    /// What the tables make of the files in a folder, asked the way a hover asks it —
    /// through the reader, the setting and the cache rather than the rule alone. The
    /// fixtures are the caller's: `RHP_CONTENT_PROBE` names a folder of real files, one of
    /// every format the lists carry and a few named as another kind, which is the check
    /// no table of its own can make.
    #[test]
    #[ignore = "reads the files named in RHP_CONTENT_PROBE"]
    fn content_probe() {
        // The configuration is the machine's own, and a file already written holds the
        // value it was written with — so the setting this module is behind is turned on
        // here rather than taken as it lies: what the probe asks about is the tables.
        if let Ok(mut config) = CONFIG.lock() {
            config.confirm_file_type = true;
        }

        let folder = std::env::var("RHP_CONTENT_PROBE").expect("RHP_CONTENT_PROBE is not set");

        let mut paths: Vec<_> = std::fs::read_dir(&folder)
            .expect("the folder named by RHP_CONTENT_PROBE")
            .filter_map(|entry| entry.ok().map(|entry| entry.path()))
            .collect();
        paths.sort();

        for path in paths {
            println!(
                "{:<16} {:?}",
                path.file_name().unwrap_or_default().to_string_lossy(),
                of(&path)
            );
        }
    }
}

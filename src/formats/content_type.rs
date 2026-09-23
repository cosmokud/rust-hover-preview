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
//! The formats it does not carry but an engine here reads are this app's own, asked in the
//! shape the rest of the app already asks signatures in: a transport stream by its sync
//! byte, an MXF by its partition pack, a raw H.264 stream by its sequence parameter set,
//! and PostScript by the comment that opens it — the same kind of question
//! `office_formats::container_kind` asks an Office document and `has_pdf_header_in` asks a
//! PDF, and one of them, the transport stream, is asked of `video_formats` rather than
//! written again.
//!
//! What neither of those names — a container whose header is not one of the common ones,
//! a document whose bytes are a proprietary format no reader here reads — is answered by
//! the name the file carries, which is the third table: everything FFmpeg demuxes and
//! everything the render engine imports, crossed with the formats [`infer`] carries a
//! signature for, because a format `infer` can name is a format the first table has
//! already answered for. Its names come from the `file_type` data theseus-rs maintains
//! (<https://github.com/theseus-rs/file-type>), which is where a name like `.wpg` or
//! `.ty+` is written down beside the format it is written for, and each entry is a kind of
//! this app's own — see [`KIND_BY_NAME`] for what is in it and what each half of it is.
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
//! A format whose signature *is* its text — an SVG, an HTML page, a JSON file — is left
//! alone for the same reason: what it is is what the text lists decide, and a name another
//! kind already read is not improved by a prefix match.
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
        .find(|signature| (signature.matches)(probe))
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
    /// `webm`, a transport stream is a `ts` and an `m2ts` — and all of them are listed
    /// because the answer is not "what is this" but "is this what the file is called": a
    /// name the file already carries is the two agreeing, and a name no list claims is a
    /// kind this app does not preview.
    names: &'static [&'static str],
    /// Whether the front of a file is one of these.
    matches: fn(&[u8]) -> bool,
}

/// The formats the common table does not carry, which are the ones an engine here reads
/// and no file manager would name: the containers FFmpeg plays that are not the everyday
/// ones, the raw streams a video is stored as, and the drawings a render engine reads.
///
/// A format the table below does not carry either is left to the file's own name, which is
/// what a picture of a format only this app's decoder reads — a `.dds`, a `.qoi`, a `.tga`
/// — and an exotic container like a RoQ file both get: the name they have is the answer,
/// exactly as it was before this module existed.
const SIGNATURES: &[Signature] = &[
    // The ISO base media family, for the names the common table leaves out of it: `3gp`,
    // `3g2` and `m4v` are the same `ftyp` box as an MP4 and a MOV are.
    Signature {
        names: &["3gp", "3g2", "m4v", "mp4", "mov"],
        matches: |probe| at(probe, 4, b"ftyp"),
    },
    // The transport streams, asked with the same probe the text lists ask for the `.ts`
    // they share with TypeScript: the sync byte, run out over the packet size. A `.m2ts`
    // is the same stream with a timestamp in front of every packet, which is one of the
    // sizes that probe knows.
    Signature {
        names: &["ts", "m2ts", "mts"],
        matches: crate::formats::video_formats::has_mpegts_packets,
    },
    // An MPEG elementary stream: a sequence header rather than a program stream, which is
    // what a `.m2v` is and what a bare `.mpg` sometimes turns out to be.
    Signature {
        names: &["m2v", "m1v", "mpg"],
        matches: |probe| starts_with(probe, &[0x00, 0x00, 0x01, 0xB3]),
    },
    // A raw H.264 stream: the start code, then the sequence parameter set one opens with.
    // Which profile it is, is the decoder's business.
    Signature {
        names: &["h264", "264", "h26l", "avc"],
        matches: is_h264_stream,
    },
    Signature {
        names: &["ivf"],
        matches: |probe| starts_with(probe, b"DKIF"),
    },
    Signature {
        names: &["mxf"],
        matches: |probe| starts_with(probe, &[0x06, 0x0E, 0x2B, 0x34]),
    },
    Signature {
        names: &["nut"],
        matches: |probe| starts_with(probe, b"NUT/MULTI"),
    },
    Signature {
        names: &["y4m"],
        matches: |probe| starts_with(probe, b"YUV4MPEG2"),
    },
    Signature {
        names: &["bik", "bk2"],
        matches: |probe| starts_with(probe, b"BIK"),
    },
    // An Ogg stream is a video where what it carries is Theora: the same container holds
    // Vorbis and Opus, and a preview of a sound is not a preview — so the codec is asked
    // for here rather than the container being taken for a picture.
    Signature {
        names: &["ogv", "ogm"],
        matches: |probe| starts_with(probe, b"OggS") && contains(probe, b"theora"),
    },
    // PostScript, which is what an `.eps` is and what an Illustrator document saved
    // without its compatibility layer is. The PDF half of that family needs no entry here:
    // a page is a page, and the common table names it.
    Signature {
        names: &["eps", "epsi", "ai", "ps"],
        matches: |probe| starts_with(probe, b"%!PS-Adobe"),
    },
    // The one text format a document is likely to be mistaken for: an RTF file under a
    // document's name is the text the text lists would have shown anyway.
    Signature {
        names: &["rtf"],
        matches: |probe| starts_with(probe, b"{\\rtf"),
    },
    // A collection of fonts is the one font container the common table does not name,
    // because what it holds is faces rather than a face.
    Signature {
        names: &["ttc"],
        matches: |probe| starts_with(probe, b"ttcf"),
    },
];

/// Whether the front of a file is an H.264 elementary stream: a start code, then the
/// sequence parameter set a stream of one opens with.
fn is_h264_stream(probe: &[u8]) -> bool {
    let Some(start) = probe.iter().position(|byte| *byte != 0) else {
        return false;
    };

    start >= 2
        && probe.get(start) == Some(&1)
        && probe
            .get(start + 1)
            .is_some_and(|header| header & 0x1F == 7)
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

/// The names the tables above have nothing to say about, and the kind each of them belongs
/// to.
///
/// What is in it is the two lists of engines this app drives — `video_formats`'s built-in
/// list, which is what FFmpeg demuxes, and `libre_formats`'s, which is what the render
/// engine imports — less every name a table above already answers for, which is what makes
/// this the last of the three tables rather than a fourth list of formats. The names were
/// read against the file type data theseus-rs maintains
/// (<https://github.com/theseus-rs/file-type>), which is what says what a name like `.wpg`
/// or `.ty+` is an extension *of*, and which is where the one name this table cannot answer
/// by itself is written down under both of the formats it belongs to.
///
/// A name belongs here only where no signature names the format it is written for. The
/// common containers — `mp4`, `mkv`, `avi`, `wmv`, `flv` and the rest — are the first
/// table's; the raw streams, the transport streams, the drawings and the formats FFmpeg
/// plays that no file manager names are the second table's; and what is left is a format
/// whose own head is nothing any table here knows.
///
/// What a name here answers is the kind, because the name is the only thing there is to
/// answer with, and it answers it whether or not a list in `config.ini` still holds the
/// name: what a format *is* does not change with a setting, and a user who wants none of
/// these previewed has the kind's own switch in the tray's `Preview Types`. Where a name
/// is claimed by both engines — `dif` is a DV stream to FFmpeg and a Data Interchange
/// Format spreadsheet to the engine — the answer is the one the lists are asked in
/// everywhere else in this app, which is the video.
static KIND_BY_NAME: &[(&str, PreviewType)] = &[
    // FFmpeg's own: the containers, the raw video streams, and the capture, camera and
    // game formats it demuxes, in the names the video list carries them under.
    video("265"),
    video("266"),
    video("3gpp"),
    video("apv"),
    video("av1"),
    video("avs"),
    video("avs2"),
    video("avs3"),
    video("c93"),
    video("cavs"),
    video("cdg"),
    video("cdxl"),
    video("cin"),
    video("cpk"),
    video("dav"),
    // A DV stream to FFmpeg, and a Data Interchange Format spreadsheet to the engine: the
    // video is what the lists answer with, which is the order every other question about a
    // file is asked in.
    video("dif"),
    video("drc"),
    video("dv"),
    video("evc"),
    video("flm"),
    video("gxf"),
    video("h261"),
    video("h263"),
    video("h265"),
    video("h266"),
    video("hevc"),
    video("ifv"),
    video("imx"),
    video("ismv"),
    video("ivr"),
    video("kux"),
    video("m2p"),
    video("m2t"),
    video("mj2"),
    video("moflex"),
    video("mpe"),
    video("mve"),
    video("mvi"),
    video("mxg"),
    video("nsv"),
    video("obu"),
    video("pmp"),
    video("psp"),
    video("rcv"),
    video("rm"),
    video("rmvb"),
    video("roq"),
    video("rsd"),
    video("smk"),
    video("str"),
    video("thp"),
    video("tod"),
    video("tp"),
    video("tr"),
    video("ty"),
    video("ty+"),
    video("usm"),
    video("vc1"),
    video("vc2"),
    video("viv"),
    video("vro"),
    video("vvc"),
    video("vw"),
    video("wtv"),
    video("xl"),
    video("xmv"),
    video("yop"),
    // The documents the render engine draws: the word processors, spreadsheets,
    // presentations and drawings of the formats that came before the modern ones and of
    // the applications beside them, the open formats themselves, and the drawings no
    // reader here reads — in the names the `[libre]` list carries them under.
    engine("123"),
    engine("602"),
    engine("abw"),
    engine("cdr"),
    engine("cgm"),
    engine("cmx"),
    engine("cwk"),
    engine("dbf"),
    engine("dxf"),
    engine("fodg"),
    engine("fodp"),
    engine("fodt"),
    engine("gnm"),
    engine("gnumeric"),
    engine("hwp"),
    engine("key"),
    engine("lwp"),
    engine("mcw"),
    engine("met"),
    engine("mw"),
    engine("numbers"),
    engine("odb"),
    engine("odc"),
    engine("odf"),
    engine("odg"),
    engine("odm"),
    engine("oth"),
    engine("otg"),
    engine("otm"),
    engine("otp"),
    engine("ots"),
    engine("ott"),
    engine("pages"),
    engine("pcd"),
    engine("pct"),
    engine("pcx"),
    engine("pdb"),
    engine("pm6"),
    engine("pmd"),
    engine("psw"),
    engine("pub"),
    engine("ras"),
    engine("sda"),
    engine("sdc"),
    engine("sdd"),
    engine("sdw"),
    engine("slk"),
    engine("stc"),
    engine("std"),
    engine("sti"),
    engine("stw"),
    engine("svm"),
    engine("sxd"),
    engine("sxg"),
    engine("sxi"),
    engine("sxm"),
    engine("sxw"),
    engine("vdx"),
    engine("vsd"),
    engine("vsdm"),
    engine("vsdx"),
    engine("vstx"),
    engine("wb2"),
    engine("wk1"),
    engine("wk3"),
    engine("wk4"),
    engine("wks"),
    engine("wpg"),
    engine("wq1"),
    engine("wq2"),
    engine("wpd"),
    engine("wps"),
    engine("wri"),
    engine("xlw"),
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
    /// written under, which is what the third table is for: a container FFmpeg demuxes and
    /// a document the render engine draws are both formats a file manager would take for
    /// anything at all, and a name is all there is to route them by.
    #[test]
    fn a_format_no_signature_names_is_answered_by_the_name_it_carries() {
        // A RoQ file: a signature this app's tables carry nothing for, under a name that
        // says what it is.
        assert_eq!(
            classified("film.roq", &[0x84, 0x10, 0xFF, 0xFF]),
            Content::Kind(PreviewType::Videos)
        );
        assert_eq!(
            classified("letter.wpd", b"nothing in here is a signature"),
            Content::Kind(PreviewType::Libre),
            "and a WordPerfect document is the engine's, by the same question"
        );
        assert_eq!(
            classified("clip.ty+", b"\x00"),
            Content::Kind(PreviewType::Videos),
            "a name the lists carry and no signature does is answered as it is written"
        );

        // A name whose format is a package is a zip in the bytes and a document in the
        // name — an iWork deck is one — and the table answers with the engine's kind,
        // which is what the list it is written for would have said.
        assert_eq!(
            classified("deck.key", b"PK\x03\x04\x14\x00\x00\x00"),
            Content::Kind(PreviewType::Libre)
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

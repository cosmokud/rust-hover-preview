//! Which archives are listed by an installed PeaZip rather than by a reader of this app's own.
//!
//! This app reads the archives people meet: a zip, a 7z, a rar, a tar and a `.tar.gz` are each
//! read from their own table of contents by the readers in `archive_listing`, and the page they
//! produce is the same page for all five. Everything else is a format no reader here has ever
//! had — the arj of a 1990s backup, an `lzh` out of old Japanese software, the cab a machine's
//! own driver packages arrive in, an iso, a wim, a vhd, the `.deb` and `.rpm` of a Linux
//! download, the single-stream `.gz`, `.bz2`, `.xz` and `.zst` a file is put through before it
//! is sent — and a hover onto one shows nothing at all today.
//!
//! PeaZip is the tool that opens them. It is a frontend rather than a decoder: the archivers it
//! ships are what actually read these formats, and which of them is asked is a question about the
//! name a file carries — see `Backend`, which answers it for every name. Most are the console
//! archiver's: the 7-Zip console PeaZip carries in `res\bin\7z`, with PeaZip's own extra codecs
//! beside it, which reads more formats than anything else it ships and can be asked for a report
//! meant to be read by something other than a person. The rest belong to the backends PeaZip
//! carries for formats that one cannot open at all: FreeArc's archiver for an `.arc`, zpaq for a
//! `.zpaq`, and Zstandard's own tool for a `.zst` — which the archiver reads too, but without
//! saying how large the stream was before it was compressed, which is the whole of what a hover on
//! one is for.
//!
//! Three of them have no listing to ask for at all. A `.br`, a `.bcm` and an `.lpaq8` are single
//! streams: the tools that write them — Brotli, BCM and LPAQ, each of them carried by PeaZip — put
//! one file into one file and have no command that prints what is inside, because there is nothing
//! inside but the bytes. What a hover on one shows is the member an extraction would write: its
//! name, taken from the file's own name the way every single-stream name here is, and no size,
//! since the only way to learn that one is to decompress. Nothing is started for one — the page is
//! this app's own answer — and what says the name may be previewed at all is the tool being there.
//!
//! So where PeaZip is installed it is asked, and what comes back is the archive's own table of
//! contents, read into the same shape every other listing is and drawn as the same page: a `.cab`
//! is previewed like a `.zip`, because that is what it is — a list of what the file holds, and not
//! a rendering of it. See `peazip_render` for the engines and `archive_listing` for the reading of
//! their answers.
//!
//! The list below is built from what the tools of an installation can be asked about rather than
//! from what PeaZip is said to support. The console archiver's own table settled most of it — the
//! one it prints when it is asked what it reads — and a name belongs here only where that table
//! declares a *format* for it rather than a codec: PeaZip ships codecs (LZ4, LZ5, Lizard,
//! Fast-LZMA2) that its console tool can *unpack inside a container* and cannot open a file of,
//! and nothing beside it opens one either, so none of them is here. Brotli was measured the same
//! way and is here for the opposite reason: the archiver refuses a `.br` too, and `brotli.exe`,
//! which PeaZip carries beside it, is what reads one.
//!
//! What is *not* here is a judgement rather than a gap, and each group is worth naming:
//!
//! * **A name another kind already reads is not here.** `7z`, `zip`, `rar`, `tar`, `zipx`,
//!   `jar`, `apk` and `xpi` are the `[archive]` list's, and they are read by this app
//!   itself — asking an engine for one would be a launch spent on a file this side reads faster
//!   and with nothing to install. `tar.gz` and `tgz` are that list's too, which is why the bare
//!   `gz` here is only ever reached by a file that is *not* a tarball: the archive list is asked
//!   first, so a `.tar.gz` stays the archive it was. `pmd` is the `[libre]` list's (PageMaker's
//!   document, not the PPMd archive the engine's table also spells that way), `swf` and `flv`
//!   are the video list's, `doc`, `xls` and `ppt` — which the engine reads inside its
//!   compound-file format — are the Office list's, `cbz`, `cbr` and `cbc` are the `[ebook]`
//!   list's, which is where a comic is a book rather than a box of files, and `lit` is
//!   `[calibre]`'s, because a Microsoft Reader book is a page the ebook engine draws. A name sits
//!   in exactly one list so that a preview of one cannot come back by two routes.
//! * **A name an engine *can* draw is not always here, and `chm` is the one that is.** A compiled
//!   help file was the ebook engine's for a build and is the archiver's again: the page that engine
//!   draws for one is right, and the two to three seconds it takes to draw it is not — a help file
//!   is a file a pointer crosses on its way somewhere else, and a listing is there before a hover
//!   has finished settling. The judgement is not about whether a page can be drawn but about what
//!   the file is hovered *for*, and it is the only name here that has been on both sides of it.
//! * **A program is not an archive.** The engine lists a `.exe`, a `.dll`, a `.sys`, an `.obj`,
//!   an `.elf`, a Mach-O binary and a firmware capsule too — what it is reading is the resources
//!   inside them — and a hover onto one of those is a hover onto a program, not onto a container
//!   of files. None of them is here, and none of them may be: an archive page for a `.dll` is a
//!   preview nobody asked for.
//! * **A format whose extension is a word rather than a format is not here.** The engine's table
//!   declares `img` for half the disk images it reads and `ext` for the Linux file system — an
//!   `.img` is a disk image in one folder and a camera's picture in the next, and neither name
//!   says what the file is. The disk images that are named here are the ones whose names are
//!   their own: `iso`, `udf`, `dmg`, `vhd`, `vhdx`, `vmdk`, `vdi`, `qcow`, `qcow2`, `squashfs`,
//!   `cramfs`, `apfs`, `hfs` and `hfsx`.
//! * **And the split archive is here only as the name it arrives under.** A file cut into
//!   `.001`, `.002` and the rest is a set rather than a file, and what the engine reads of the
//!   first part is the whole of the archive it was cut from — which is exactly what a hover onto
//!   it should show, so `001` is an entry of the list and nothing would be gained by naming the
//!   rest.
//!
//! Nothing is bundled with this app and nothing is linked against: the tools are the user's own
//! installation of PeaZip, looked for where it installs and beside `config.ini` for a portable
//! copy, and where there is none a hover onto one of these names shows nothing at all — the same
//! answer a camera raw gets on a machine without ImageMagick. A name is answered only where the
//! tool that reads it is there, so a portable copy carrying a few of them shows the names those
//! The list of names this answers for is a row of `crate::formats::lists` — the one table
//! every kind's list is a row of, and the one place a list is written down, the built-in
//! entries and the older lists this app shipped and then changed included.

use crate::config::config::PreviewType;
use crate::formats::{lists, text_formats};
use crate::CONFIG;
use std::path::Path;

/// The tool inside an installed PeaZip that reads a file, and so the tool this app runs to list it.
///
/// The console archiver answers for every name but a handful: it reads more formats than anything
/// else PeaZip ships and can be asked for a report meant to be read by something other than a
/// person. The others are the backends it carries beside it for the formats that one cannot open
/// — and they are not interchangeable with it: FreeArc and zpaq each read their own format and no
/// other, and the three single-stream compressors have no listing at all (see `has_a_listing`).
///
/// A name is routed by the extension it carries and never by its bytes, which is the line TODO.md
/// draws for these formats: a file is not one of theirs by its bytes alone. What that costs is a
/// renamed archive — an `.arc` called a `.dat` is a file none of these tools is asked about, and it
/// shows nothing — and what it buys is that a picture or a document whose bytes a marker of one of
/// these formats happens to resemble is never handed to an archiver on the strength of it.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum Backend {
    /// The console archiver PeaZip carries in `res\bin\7z`, asked for the technical listing of a
    /// name its own format table declares.
    SevenZip,
    /// FreeArc, PeaZip's backend for the `.arc` it has read-write support for, asked for its
    /// verbose listing.
    Arc,
    /// zpaq, the journaling archiver, asked to list the names it holds — the latest version of
    /// each, since that is what a hover is for.
    Zpaq,
    /// Zstandard's own tool, which reports what the archiver leaves blank for a `.zst`: how many
    /// frames the stream holds, and how large it was before it was compressed.
    Zstd,
    /// And the three whose tools cannot answer: Brotli, BCM and LPAQ. Each puts one file into one
    /// stream, so what is previewed is the single member an extraction would write, and nothing is
    /// started to find that out.
    Brotli,
    Bcm,
    Lpaq,
}

impl Backend {
    /// The tool that reads a file, settled by the extension its name carries.
    ///
    /// It is the extension the lists are matched by — `text_formats::lookup_extension`, shared
    /// rather than copied so that routing and the gate cannot disagree about what a file is called
    /// — so a name here is the same one `[peazip] extensions` is written with. A name no arm below
    /// names is the console archiver's, which is what the list is mostly made of and what a user's
    /// own added entry is read by.
    pub fn of(path: &Path) -> Backend {
        match text_formats::lookup_extension(path).as_deref() {
            Some("arc") => Backend::Arc,
            Some("zpaq") => Backend::Zpaq,
            Some("zst" | "tzst" | "zstd") => Backend::Zstd,
            Some("br") => Backend::Brotli,
            Some("bcm") => Backend::Bcm,
            Some("lpaq8") => Backend::Lpaq,
            _ => Backend::SevenZip,
        }
    }

    /// The path this tool keeps inside a PeaZip installation, which is where it is looked for.
    pub fn program(self) -> &'static str {
        match self {
            Backend::SevenZip => r"res\bin\7z\7z.exe",
            Backend::Arc => r"res\bin\arc\Arc.exe",
            Backend::Zpaq => r"res\bin\zpaq\zpaq.exe",
            Backend::Zstd => r"res\bin\zstd\zstd.exe",
            Backend::Brotli => r"res\bin\brotli\brotli.exe",
            Backend::Bcm => r"res\bin\quad\bcm.exe",
            Backend::Lpaq => r"res\bin\lpaq\lpaq8.exe",
        }
    }

    /// Whether this tool has a listing of its own to give.
    ///
    /// The three that do not are single-stream compressors: one file goes in, one comes out, and
    /// no command of theirs prints what is inside, because there is nothing inside but the bytes.
    /// What the app shows for one of their files is derived from the file's own name rather than
    /// read out of an answer (see `peazip_render`), and no process is started for it — so the
    /// presence of the tool, and not a launch, is what says the name may be previewed at all.
    pub fn has_a_listing(self) -> bool {
        !matches!(self, Backend::Brotli | Backend::Bcm | Backend::Lpaq)
    }
}

/// Whether this kind's list claims `path`, without asking whether these previews are switched on,
/// and without claiming a file a reader of this app's own already reads.
///
/// The second half is what keeps a tarball out of the engine's hands. This list carries the bare
/// `gz` a single compressed stream is named by, and the archive list carries the dotted `tar.gz`
/// that names a whole file — so a `.tar.gz` is named by both, by its tail and by its name, and
/// the entry that names the whole file is the one that is right about it. What that costs is one
/// list look per question; what it buys is that a file this app reads itself is never a file an
/// engine is started for, whatever a user's own edit to either list says. The gate is asked
/// beside it by the hook, the way every other kind's is.
///
/// Both halves are asked under one guard rather than under two, which is the shape this had when
/// the archive list's question took the same lock the peazip one had given up: one lock taken
/// twice on a thread is a deadlock, and the two list comparisons are both in memory.
pub fn is_peazip_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| {
            lists::PEAZIP.claims(path, &config)
                && !crate::formats::archive_formats::claims_in(
                    path,
                    lists::ARCHIVE.entries(&config),
                )
        })
        .unwrap_or(false)
}

/// Whether the engine is the one that lists this file at all: a name its own list carries, or the
/// bytes of an archive it reads under a name no list holds — a `.cab` renamed to `.dat`, say.
///
/// It is the question the request side asks before it asks the engine for a listing — a file
/// whose bytes name another kind is not one to start it for, and one whose bytes name *this* kind
/// is the engine's whatever it is called — and it is the same question asked the same way
/// `libre_formats::engine_page_kind` asks it for the render engine and
/// `magick_formats::is_engine_picture` for the image converter: the file's own bytes first, and
/// the name it carries after them. What it is *not* is a gate: whether that kind may be shown is
/// the caller's to ask (`PreviewType::enabled`), because the same question is asked of a preview
/// that is already on screen when a switch is thrown.
pub fn is_engine_archive(path: &Path) -> bool {
    use crate::formats::content_type::Content;

    // The entry is read before the lock, so the guard is not held across the read (see
    // `content_type::of_reaching_config`).
    let content = crate::formats::content_type::of_reaching_config(path);

    match content {
        Content::Kind(PreviewType::Peazip) => true,
        Content::Kind(_) | Content::Foreign => false,
        Content::Unknown => is_peazip_file(path),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// What the list holds is what no reader here and no other list of this app's names: an
    /// archive this app reads itself, a document an engine of its own draws, a video and a
    /// picture are all answered elsewhere, and none of them is here.
    #[test]
    fn holds_no_name_another_kind_already_reads() {
        let config = crate::config::config::AppConfig::default();

        let claimed_elsewhere = |path: &Path| {
            crate::formats::lists::IMAGE.claims(path, &config)
                || crate::formats::lists::VECTOR.claims(path, &config)
                || crate::formats::lists::DESIGN.claims(path, &config)
                || crate::formats::lists::FONT.claims(path, &config)
                || crate::formats::lists::LIBRE.claims(path, &config)
                || crate::formats::lists::MAGICK.claims(path, &config)
                || crate::formats::video_formats::matches_any_video_list(path, &config)
                || crate::formats::lists::ARCHIVE.claims(path, &config)
                || crate::formats::lists::OFFICE.claims(path, &config)
                || crate::formats::lists::EBOOK.claims(path, &config)
                || crate::formats::lists::CALIBRE.claims(path, &config)
                || crate::formats::lists::TEXT.claims(path, &config)
                || crate::formats::lists::NAMES.claims(path, &config)
        };

        for name in [
            "archive.zip",
            "archive.7z",
            "archive.rar",
            "archive.tar",
            "archive.tgz",
            "archive.zipx",
            "package.jar",
            "package.apk",
            "package.xpi",
            "comic.cbz",
            "document.pmd",
            "book.lit",
            "animation.swf",
            "clip.flv",
            "report.doc",
            "sheet.xls",
            "slides.ppt",
        ] {
            let path = Path::new(name);
            assert!(
                claimed_elsewhere(path),
                "`{name}` is read by another kind, so this expectation is written the wrong way round"
            );
            assert!(
                !crate::formats::lists::PEAZIP.claims(path, &config),
                "`{name}` is read by another kind, so the engine is not asked about it"
            );
        }

        // And the one name the two lists can disagree about is checked against the question the
        // app really asks, because the list alone cannot answer it: this list carries the bare
        // `gz` a single compressed stream is named by, and the archive list carries the dotted
        // `tar.gz` that names a whole file, so a `.tar.gz` is named by both — by its tail and by
        // its name — and the entry that names the whole file is the one that wins (see
        // `is_peazip_file`).
        let tarball = Path::new("sources.tar.gz");
        assert!(
            crate::formats::lists::ARCHIVE.claims(tarball, &config),
            "a tarball is the archive list's, by the dotted entry that names it"
        );
        assert!(
            crate::formats::lists::PEAZIP.claims(tarball, &config),
            "and this list names its tail, which is why the difference exists at all"
        );
        assert!(
            !is_peazip_file(tarball),
            "so the file is read here rather than handed to an engine"
        );
    }

    /// And what it does hold is the archives nothing else on the machine opens, the
    /// single-stream compressors, and the disk images whose names are their own.
    #[test]
    fn holds_the_archives_no_reader_here_opens() {
        let config = crate::config::config::AppConfig::default();

        for name in [
            "backup.arj",
            "backup.arc",
            "backup.cab",
            "help.chm",
            "linux.deb",
            "linux.rpm",
            "image.iso",
            "archive.udf",
            "image.dmg",
            "image.vhd",
            "image.vhdx",
            "image.vmdk",
            "package.msi",
            "package.wim",
            "readme.gz",
            "readme.bz2",
            "readme.xz",
            "readme.zst",
            "readme.z",
            "readme.br",
            "readme.bcm",
            "readme.lpaq8",
            "backup.zpaq",
            "archive.001",
            "archive.lzh",
        ] {
            assert!(
                crate::formats::lists::PEAZIP.claims(Path::new(name), &config),
                "`{name}` is one of the engine's archives"
            );
        }
    }

    /// And every name is routed to the tool inside an installation that reads it: the console
    /// archiver for the formats its own table declares, and the backend PeaZip carries beside it
    /// for the names that one cannot open at all.
    #[test]
    fn routes_a_name_to_the_tool_that_reads_it() {
        let routed = |name: &str| Backend::of(Path::new(name));

        assert_eq!(routed("backup.arc"), Backend::Arc);
        assert_eq!(routed("backup.zpaq"), Backend::Zpaq);
        assert_eq!(routed("notes.txt.zst"), Backend::Zstd);
        assert_eq!(routed("notes.tzst"), Backend::Zstd);
        assert_eq!(routed("notes.txt.br"), Backend::Brotli);
        assert_eq!(routed("notes.txt.bcm"), Backend::Bcm);
        assert_eq!(routed("notes.txt.lpaq8"), Backend::Lpaq);

        // The console archiver's is every other name in the list, and every name a user adds to
        // it: it reads the most formats of anything PeaZip ships, so a name no arm above names is
        // its own rather than nothing's.
        for name in [
            "backup.cab",
            "system.iso",
            "readme.bz2",
            "image.vhd",
            "archive.001",
            "something-a-user-added.xyz",
        ] {
            assert_eq!(
                routed(name),
                Backend::SevenZip,
                "`{name}` is the archiver's to read"
            );
        }

        // And a name is routed by the same extension the lists are matched by, which is why that
        // lookup is one function: `ZST` is the `zst` entry, and the tool is the same one.
        assert_eq!(routed("notes.txt.ZST"), Backend::Zstd);
    }

    /// Of those tools, three have no listing of their own to give, and each of them is a
    /// single-stream compressor. What the app does instead of asking one is written where the
    /// routing is; what is asserted here is which of them it is.
    #[test]
    fn knows_which_tools_have_a_listing_to_give() {
        for backend in [Backend::Brotli, Backend::Bcm, Backend::Lpaq] {
            assert!(
                !backend.has_a_listing(),
                "{backend:?} compresses one file to one stream and has nothing to print"
            );
            assert!(
                backend.program().ends_with(".exe"),
                "{backend:?} still names the tool that reads its format"
            );
        }

        for backend in [
            Backend::SevenZip,
            Backend::Arc,
            Backend::Zpaq,
            Backend::Zstd,
        ] {
            assert!(
                backend.has_a_listing(),
                "{backend:?} lists what it reads and is asked for it"
            );
        }
    }

    /// A name that is not a bare extension is dropped rather than matched against, the way every
    /// other list of this app's answers a hand-edited entry.
    #[test]
    fn reads_a_list_of_bare_extensions() {
        let extensions =
            crate::formats::lists::sanitize_extension_list(" .CAB , arj,,archive*.iso ,cab,msi");

        assert_eq!(extensions, vec!["cab", "arj", "msi"]);
    }

    /// What the engine is asked about is asked of the file's own bytes before its name, which is
    /// what makes an archive renamed to a name no list holds the engine's to list: a cabinet file
    /// under a `.dat` is one of its archives, an archive this app reads itself is not, and a name
    /// nothing recognizes is left to the list the name is written in.
    #[test]
    fn asks_the_engine_about_a_file_by_its_bytes_before_its_name() {
        let folder = std::env::temp_dir().join("rust-hover-preview-peazip-engine-archive");
        std::fs::create_dir_all(&folder).expect("a test folder");

        // A cabinet file: `MSCF` and the four zero bytes every one of them carries.
        let renamed = folder.join("package.dat");
        std::fs::write(&renamed, [b'M', b'S', b'C', b'F', 0, 0, 0, 0, 0, 0, 0, 0])
            .expect("a written archive");
        assert!(
            is_engine_archive(&renamed),
            "an archive the engine reads is its to list, whatever it is called"
        );

        // The same bytes under a name this app's own archive list claims: the two agree about
        // what the file is, so a reader here has it and the engine is not asked.
        let named = folder.join("archive.7z");
        std::fs::write(&named, b"7z\xBC\xAF\x27\x1C").expect("a written file");
        assert!(
            !is_engine_archive(&named),
            "an archive this app reads itself is not one to start an engine for"
        );

        // A picture under an archive's name is a picture: what the file's own bytes say is the
        // first answer, and a reader here has it, so no engine is started for it.
        let picture = folder.join("bundle.cab");
        std::fs::write(&picture, [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A])
            .expect("a written file");
        assert!(
            !is_engine_archive(&picture),
            "a file whose bytes are another kind's is not the engine's"
        );

        // And the name's own list, for the archives whose bytes say nothing this table knows.
        std::fs::write(folder.join("backup.cab"), b"a cabinet, of a sort").expect("a written file");
        assert!(
            is_engine_archive(&folder.join("backup.cab")),
            "a name the list carries is the engine's, whatever the bytes say"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }
}

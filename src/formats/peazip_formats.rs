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
//! few read and no others. The list lives in `config.ini` as `[peazip] extensions`, written from
//! the built-in list on first run and read back from there, so a user can add a format a tool of
//! theirs reads and this app does not know, or take one out.
//!
//! The question here is only what a file is *called*, and for most of these names that is the
//! whole of it: what the file *is* — whether a tool can read it at all — is settled by the tool,
//! and a name it cannot read is answered with no preview once and then remembered, so a name put
//! in this list by mistake costs one launch and never another. The three single-stream names are
//! the exception, and the reason is the one written above: nothing is started for one, so there
//! is no answer to remember and nothing to refuse. See `is_engine_archive`, which asks the file's
//! own bytes before its name and is the one question every side asks before a tool is started.

use crate::config::config::PreviewType;
use crate::formats::text_formats;
use crate::CONFIG;
use std::path::Path;

/// The extensions written to `config.ini` on first run: the names the tools of an installed
/// PeaZip can be asked about, that no list of this app's own already claims, and that are
/// containers of files rather than programs.
///
/// Three groups, in the order they are written:
///
/// * The archives and installers nothing else on the machine opens: `001`, `ar`, `arc`, `arj`,
///   `cab`, `chm`, `cpio`, `deb`, `esd`, `hfs`, `hfsx`, `hxs`, `iso`, `lha`, `lzh`, `msi`,
///   `msp`, `pkg`, `ppkg`, `rpm`, `swm`, `udf`, `wim`, `xar`, `xip` and `zpaq`. Some are
///   containers of files in the ordinary sense (an installer, a compiled help file, a Linux package,
///   a disk image), some are the volume of a backup, and all of them are read by a tool of PeaZip's
///   and by nothing this app has. Two of them are that tool's rather than the console archiver's:
///   an `arc` is FreeArc's and a `zpaq` is zpaq's, and neither is read by the archiver at all.
///   A Microsoft Reader book was of this group once and is not any more: the ebook engine draws a
///   page of one, so it is `[calibre]`'s, and `PEAZIP_EXTENSIONS_BEFORE_THE_EBOOKS` below is what
///   takes it out of a file this app wrote before that. A compiled help file went the other way —
///   out of this group and back into it — which is what `PEAZIP_EXTENSIONS_BEFORE_THE_HELP_FILE` is
///   written down for.
/// * The single-stream compressors: `bcm`, `br`, `bz2`, `bzip2`, `gz`, `gzip`, `lpaq8`, `lzma`,
///   `xz`, `z` and `zst`, with the tarball spellings that name them (`taz`, `tbz`, `tbz2`, `tpz`,
///   `tzst`). One of these is a file put through a compressor rather than a container, so the
///   answer is one member, often one whose name is not in the stream at all — and the reading of
///   that answer is what puts a name to it (see `archive_listing`). They are here because a tool
///   of PeaZip's reads them and this app does not, and because what a hover shows for one — the
///   name, the size where the tool knows it, and how much smaller it was made — is the useful
///   part of opening it. `bcm`, `br` and `lpaq8` are the three whose tools cannot say even that
///   much; nothing is started for one of them.
/// * And the disk images whose names are their own: `apfs`, `cramfs`, `dmg`, `qcow`, `qcow2`,
///   `squashfs`, `vdi`, `vhd`, `vhdx`, `vmdk`.
///
/// Deliberately absent, and each for a reason the module documentation above gives: the names
/// another list of this app's already reads (`7z`, `zip`, `rar`, `tar`, `zipx`, `jar`, `apk`,
/// `xpi`, `cbz`, `cbr`, `cbc`, `chm`, `lit`, `tgz`, `pmd`, `swf`, `flv`, `doc`, `xls`, `ppt`),
/// the programs the engine lists as resources (`exe`, `dll`, `sys`, `obj`, `elf`, `macho`, `te`,
/// `b64`, `ihex`, `simg`, `uefif`, `scap`, `lpimg`, `nsis`, `mslz`, `mub`), the extensions that
/// are words rather than formats (`img`, `ext`, `ext2`, `ext3`, `ext4`, `fat`, `ntfs`, `apm`,
/// `mbr`, `gpt`), the codecs its build carries no format for and no tool of its own opens (`lz4`,
/// `lz5`, `lizard`, `flzma2`), the two names this installation carries no tool for at all (`pea`,
/// which no tool of PeaZip's lists, and `paq8`, whose folder holds no executable — both are written
/// down in TODO.md), and the compression formats this app has no need of an engine for (`xz` is
/// here, `lzma86` and `base64` are not).
pub const DEFAULT_PEAZIP_EXTENSIONS: &str =
    "001,apfs,ar,arc,arj,bcm,br,bz2,bzip2,cab,chm,cpio,cramfs,deb,dmg,esd,gz,gzip,\
hfs,hfsx,hxs,iso,lha,lpaq8,lzh,lzma,msi,msp,pkg,ppkg,qcow,qcow2,rpm,squashfs,swm,\
taz,tbz,tbz2,tpz,txz,tzst,udf,udeb,vdi,vhd,vhdx,vmdk,wim,xar,xip,xz,z,zpaq,zst";

/// The built-in `[peazip]` list as it stood while a compiled help file was the ebook engine's.
///
/// A file holding exactly these entries is this app's own earlier list rather than a user's edit, so
/// it is brought up to the built-in list rather than kept as written — which is what gives the name
/// back to the archiver, on an installation that ran the build that had taken it away. `chm` is the
/// only name to have moved twice, and the reason it moved back is worth putting beside it: what the
/// engine draws for one is a page and takes two to three seconds to draw, and what a help file is
/// hovered for is usually nothing at all — so the listing that is there immediately is the better
/// answer, and the page it gives up is one nobody was waiting for (see `calibre_formats`).
pub const PEAZIP_EXTENSIONS_BEFORE_THE_HELP_FILE: &str =
    "001,apfs,ar,arc,arj,bcm,br,bz2,bzip2,cab,cpio,cramfs,deb,dmg,esd,gz,gzip,\
hfs,hfsx,hxs,iso,lha,lpaq8,lzh,lzma,msi,msp,pkg,ppkg,qcow,qcow2,rpm,squashfs,swm,\
taz,tbz,tbz2,tpz,txz,tzst,udf,udeb,vdi,vhd,vhdx,vmdk,wim,xar,xip,xz,z,zpaq,zst";

/// The built-in `[peazip]` list as it stood while `chm` and `lit` were the archiver's names.
///
/// A file holding exactly these entries is the app's own older list rather than a user's edit, so it
/// is brought up to the built-in list rather than kept as written — which is what takes the two
/// names out of every `config.ini` already written. Both are books rather than archives: a
/// compiled help file and a Microsoft Reader book are read by the ebook engine and previewed as a
/// page of one, which needs Calibre installed, where the archiver listed what they hold (see
/// `calibre_formats`). A user who would rather have the listing back adds the name to this list
/// again, and takes it out of `[calibre]` with it.
pub const PEAZIP_EXTENSIONS_BEFORE_THE_EBOOKS: &str =
    "001,apfs,ar,arc,arj,bcm,br,bz2,bzip2,cab,chm,cpio,cramfs,deb,dmg,esd,gz,gzip,\
hfs,hfsx,hxs,iso,lha,lit,lpaq8,lzh,lzma,msi,msp,pkg,ppkg,qcow,qcow2,rpm,squashfs,swm,\
taz,tbz,tbz2,tpz,txz,tzst,udf,udeb,vdi,vhd,vhdx,vmdk,wim,xar,xip,xz,z,zpaq,zst";

/// The list this app wrote before the tools beside the console archiver were driven: the names
/// that list held, which is what tells a file written by that build from one a user has edited
/// (see `crate::config::config::repair_older_lists`).
///
/// A list holding exactly these entries is this app's own — nobody typed it — and is brought up
/// to `DEFAULT_PEAZIP_EXTENSIONS`, which is how an installation that already exists is given the
/// names the archiver's own table never declared: `arc`, `zpaq`, `br`, `bcm` and `lpaq8`. The
/// entries are written in the order they were written then, since that is what a file holds.
pub const PEAZIP_EXTENSIONS_BEFORE_THE_BACKENDS: &str =
    "001,apfs,ar,arj,bz2,bzip2,cab,chm,cpio,cramfs,deb,dmg,esd,gz,gzip,\
hfs,hfsx,hxs,iso,lha,lit,lzh,lzma,msi,msp,pkg,ppkg,qcow,qcow2,rpm,squashfs,swm,\
taz,tbz,tbz2,tpz,txz,tzst,udf,udeb,vdi,vhd,vhdx,vmdk,wim,xar,xip,xz,z,zst";

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

/// Whether the configured list claims `path`.
pub fn matches_peazip_list(path: &Path, extensions: &[String]) -> bool {
    text_formats::matches_configured_extension(path, extensions)
}

/// Read one entry out of the configured list into the lowercase form the lookups use.
///
/// Every name in this list is a bare extension — unlike the archive list, which has to carry the
/// dotted `tar.gz` — so anything that is not one is dropped rather than matched against.
pub fn sanitize_peazip_extensions(list: &str) -> Vec<String> {
    let mut extensions: Vec<String> = Vec::new();

    for entry in list.split(',') {
        let trimmed = entry.trim().trim_start_matches('.').to_lowercase();
        let is_extension = !trimmed.is_empty()
            && !trimmed.starts_with('.')
            && !trimmed.ends_with('.')
            && trimmed
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || matches!(c, '+' | '-' | '_'));

        if is_extension && !extensions.contains(&trimmed) {
            extensions.push(trimmed);
        }
    }

    extensions
}

/// Whether the configured list claims `path`, without asking whether these previews are switched
/// on, and without claiming a file a reader of this app's own already reads.
///
/// The second half is what keeps a tarball out of the engine's hands. This list carries the bare
/// `gz` a single compressed stream is named by, and the archive list carries the dotted `tar.gz`
/// that names a whole file — so a `.tar.gz` is named by both, by its tail and by its name, and
/// the entry that names the whole file is the one that is right about it. What that costs is one
/// list look per question; what it buys is that a file this app reads itself is never a file an
/// engine is started for, whatever a user's own edit to either list says. The gate is asked
/// beside it by the hook, the way every other kind's is.
pub fn is_peazip_file(path: &Path) -> bool {
    let claimed = CONFIG
        .lock()
        .map(|config| matches_peazip_list(path, &config.peazip_extensions))
        .unwrap_or(false);

    // Asked after the lock is given up rather than inside it: the archive list's own question
    // takes the same lock, and one lock taken twice on a thread is a deadlock.
    claimed && !crate::formats::archive_formats::is_archive_file(path)
}

/// Whether a preview may be shown for `path`: the file the configured list claims, and the
/// `Peazip` gate in the tray's `Preview Types` submenu.
///
/// Both halves ask it where a kind can be switched off under a preview that is already on
/// screen: a hover is not sent for a kind that is off, and the layout places nothing for a file
/// whose kind is off, which is how a preview of that kind comes down when the switch does. See
/// `PreviewType::enabled`.
pub fn is_peazip_preview(path: &Path) -> bool {
    is_peazip_file(path) && PreviewType::Peazip.enabled()
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
    use crate::formats::content_type::{self, Content};

    match content_type::of(path) {
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
        let list = sanitize_peazip_extensions(DEFAULT_PEAZIP_EXTENSIONS);
        let config = crate::config::config::AppConfig::default();

        let claimed_elsewhere = |path: &Path| {
            crate::formats::image_formats::matches_image_list(path, &config.image_extensions)
                || crate::formats::vector_formats::matches_vector_list(
                    path,
                    &config.vector_extensions,
                )
                || crate::formats::design_formats::matches_design_list(
                    path,
                    &config.design_extensions,
                )
                || crate::formats::font_formats::matches_font_list(path, &config.font_extensions)
                || crate::formats::libre_formats::matches_libre_list(path, &config.libre_extensions)
                || crate::formats::magick_formats::matches_magick_list(
                    path,
                    &config.magick_extensions,
                )
                || crate::formats::video_formats::matches_video_list(path, &config.video_extensions)
                || crate::formats::archive_formats::matches_archive_list(
                    path,
                    &config.archive_extensions,
                )
                || crate::formats::office_formats::matches_office_list(
                    path,
                    &config.office_extensions,
                )
                || crate::formats::ebook_formats::matches_ebook_list(path, &config.ebook_extensions)
                || crate::formats::calibre_formats::matches_calibre_list(
                    path,
                    &config.calibre_extensions,
                )
                || crate::formats::text_formats::matches_text_lists(
                    path,
                    &config.text_extensions,
                    &config.text_names,
                )
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
                !matches_peazip_list(path, &list),
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
            crate::formats::archive_formats::matches_archive_list(
                tarball,
                &config.archive_extensions
            ),
            "a tarball is the archive list's, by the dotted entry that names it"
        );
        assert!(
            matches_peazip_list(tarball, &list),
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
        let list = sanitize_peazip_extensions(DEFAULT_PEAZIP_EXTENSIONS);

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
                matches_peazip_list(Path::new(name), &list),
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
        let extensions = sanitize_peazip_extensions(" .CAB , arj,,archive*.iso ,cab,msi");

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

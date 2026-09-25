//! Which ebooks are previewed by an installed Calibre rather than by a reader of this app's own.
//!
//! This app previews two things as a book from its own readers — a PDF, drawn from page one of the
//! file, and a comic, whose first plate is a picture inside a container this app reads itself (see
//! `ebook_formats`) — and nothing else. Everything a person actually buys, borrows or is sent is a
//! format neither of those opens: the Mobipocket and Kindle families (`.mobi`, `.azw`, `.azw3`,
//! `.azw4`, `.prc`), the open one (`.epub`), the Russian one (`.fb2`), the ones the dedicated
//! readers of the 2000s used (`.lrf`, `.tcr`, `.pml`, `.rb`, `.snb`), the Palm databases the first
//! ebooks arrived as (`.pdb`), the compiled help file (`.chm`), the Microsoft Reader book (`.lit`),
//! and the scanned book (`.djvu`) — and a hover onto one of them shows nothing at all today.
//!
//! Calibre is the tool that opens them, and it opens more of them than anything else on a Windows
//! machine: its conversion pipeline reads every format above and writes a PDF of what it read,
//! which is the one thing this app already knows how to draw. So where Calibre is installed it is
//! asked, and the page it writes is drawn as a PDF page is — at the `Ebook` kind's scale, over the
//! `Ebook` kind's backdrop, under the `Ebook` kind's switch. See `calibre_render` for the
//! conversion and `document_cache` for the page it is kept as.
//!
//! Nothing is bundled with this app and nothing is linked against: the engine is the user's own
//! installation of Calibre, looked for where it installs — and beside `config.ini` for a portable
//! copy — and run as the user runs it. A name is answered only where the engine is there, and a
//! machine without Calibre shows no preview for one of these names rather than the first page of
//! text the text preview would make of the bytes, which is what a `.fb2` is under its markup. The
//! list lives in `config.ini` as `[calibre] extensions`, written from the built-in list on first
//! run and read back from there, so a user can add a format the engine reads and this app does not
//! know, or take one out.
//!
//! What is listed here is what the engine's own input plugins declare, checked one name at a time,
//! and what no other list of this app's already claims:
//!
//! * The Kindle and Mobipocket family: `azw`, `azw3`, `azw4`, `mobi` and `prc`. All of them are
//!   Palm databases with the same two identifiers inside — see `content_type`, which reads them —
//!   and the engine reads them as one format under five names.
//! * `lit`, the Microsoft Reader book: HTML compressed with LZX inside an OLE compound file, which
//!   is neither a container of pictures nor a format whose text any reader here unpacks, so it is the
//!   engine's. Its neighbour `.chm` is the same kind of box and is *not* here: a compiled help file
//!   is drawn as a page by this engine and takes two to three seconds to draw, and a help file is a
//!   file a pointer crosses rather than one it waits on — so the name is the listing engine's, where
//!   it is answered immediately, and what is given up is a page nobody was waiting for (see
//!   `peazip_formats`).
//! * The open one, `epub`, which is the format most ebooks are sold in.
//! * `fb2`, the FictionBook the Russian ebook sites write.
//! * `djvu`, the scanned book, in the same family as a PDF: what a hover shows is a page of the
//!   scan rather than its text, because a scan has no text to read.
//! * `lrf`, the Sony reader's own container.
//! * `tcr`, `pml`, `snb` and `htmlz`, the formats of the dedicated readers and of the engine
//!   itself.
//!
//! What is deliberately *not* here is a judgement rather than a gap, and each group is worth
//! naming — see TODO.md, where every one of them is written down with what it would take:
//!
//! * **A name another kind already reads is not here.** `docx` is the Office list's, `odt` and
//!   `pdb` are the `[libre]` list's, `html`, `rtf` and `txt` are the text lists', and `pdf` is this
//!   app's own page reader. A name sits in exactly one list so that a preview of one cannot come
//!   back by two routes.
//! * **And the books this app reads itself are not here, because it reads them.** `cbz`, `cbr` and
//!   `cbc` are the `[ebook]` list's: the engine converts a comic the way it converts everything —
//!   unpacking the archive, decoding every plate, rewriting each one and building a document of the
//!   lot — which for a comic of a hundred megabytes is minutes of work for a preview that is one
//!   page, against reading that page out of the container in milliseconds (see `comic_preview`).
//! * **A name that is a programming language is not here.** `rb` is Rocket eBook to the engine and
//!   a Ruby source file to everyone who writes one, and there are far more of the second. A
//!   `.rb` is a text file, which is what the text list already says it is.
//! * **A format the engine does not read is not here.** `tpz` — Amazon's Topaz — is detected by
//!   the engine and refused with a message saying so, and `kfx`, the format the current Kindle
//!   store serves, is read by a third-party plugin rather than by the engine itself. Neither is
//!   asked about, because a name the engine cannot read costs a launch before the answer is
//!   remembered.
//! * **And the formats whose own head is too weak to be worth a signature are answered by their
//!   names alone.** A `.lrx` is the DRM-protected spelling of the Sony container below and is not
//!   declared by the engine's own LRF plugin, so it is left out of the list and written down in
//!   TODO.md; a `.snb` is a SQLite database, a `.tcr` and a `.pml` are text, a `.chm` declares
//!   itself and a `.lit` is an OLE compound file every Office document is as well, and a `.htmlz` is
//!   a zip — for those the name is the whole of the question, and they are the reason the list is
//!   asked at all.
//!
//! The question here is only what a file is *called*, and for most of these names that is the whole
//! of it: what the file *is* — whether the engine can read it at all — is settled by the engine,
//! and a name it cannot read is answered with no preview once and then remembered, so a name put in
//! this list by mistake costs one conversion and never another. The names whose bytes say what they
//! are ask it first, as they do everywhere: a renamed `.mobi`, `.epub`, `.djvu`, `.fb2` or `.lrf`
//! is the engine's whatever it is called (see `content_type::SIGNATURES`).

use crate::config::config::PreviewType;
use crate::formats::text_formats;
use crate::CONFIG;
use std::path::Path;

/// The extensions written to `config.ini` on first run: the ebook formats the engine reads as
/// input that no list of this app's own already claims.
///
/// The four groups are the module documentation's, in the order they are written there: the Kindle
/// and Mobipocket family (`azw`, `azw3`, `azw4`, `mobi`, `prc`), the two books whose text is packed
/// the way no reader here unpacks it (`chm`, `lit`), the open and single-reader formats (`djvu`,
/// `epub`, `fb2`, `lrf`), and the formats of the dedicated readers and of the engine itself
/// (`htmlz`, `pml`, `snb`, `tcr`).
///
/// Deliberately absent, and each for a reason the module documentation above gives: the names
/// another list of this app's already reads (`cbz`, `docx`, `odt`, `pdb`, `html`, `rtf`, `txt`,
/// `pdf`), the comic books, which this app reads itself and needs no engine for (`cbr`, `cbc`), the
/// name that is a programming language (`rb`), the Sony container's protected spelling (`lrx`), and
/// the names the engine does not read at all (`tpz`, and the `kfx` a plugin would be needed for).
pub const DEFAULT_CALIBRE_EXTENSIONS: &str =
    "azw,azw3,azw4,djvu,epub,fb2,htmlz,lit,lrf,mobi,pml,prc,snb,tcr";

/// The built-in `[calibre]` list as it stood while a compiled help file was this engine's to draw.
///
/// `chm` was in this list for one build and is not any more: the page the engine draws for one is
/// right, and the two to three seconds it takes to draw it is not — a help file is a file a pointer
/// crosses on its way somewhere else, and the listing the archiver prints for one is there before a
/// hover has finished settling (see `peazip_formats`, which is where the name is again). A file
/// holding this list is a file this app wrote, so it is brought up to the list of now rather than
/// kept as written, and the name goes back where it came from.
pub const CALIBRE_EXTENSIONS_WITH_THE_HELP_FILE: &str =
    "azw,azw3,azw4,chm,djvu,epub,fb2,htmlz,lit,lrf,mobi,pml,prc,snb,tcr";

/// The built-in `[calibre]` list as it stood before `chm` and `lit` became the engine's names.
///
/// A file holding exactly these entries is this app's own earlier list rather than a user's edit —
/// nobody has typed it — so it is brought up to the built-in list rather than kept as written, which
/// is what gives an installation that already exists the name that stayed. Without it that name would
/// reach a fresh installation only: every `config.ini` already written holds the list as it was, and a
/// list nobody has touched is indistinguishable from one a user edited unless the older spellings of
/// it are written down here (see `config::repair_older_lists`).
///
/// Both names were the `[peazip]` list's until then, so the same change takes them out of that list.
/// `chm` has since gone back — see [`CALIBRE_EXTENSIONS_WITH_THE_HELP_FILE`] — so a file can hold
/// either of the two older spellings of this list, and both are written down for that reason: a list
/// this app shipped is a list it brings up to now, whichever of them it is.
pub const CALIBRE_EXTENSIONS_BEFORE_THE_EBOOKS: &str =
    "azw,azw3,azw4,djvu,epub,fb2,htmlz,lrf,mobi,pml,prc,snb,tcr";

/// Whether the configured list claims `path`.
pub fn matches_calibre_list(path: &Path, extensions: &[String]) -> bool {
    text_formats::matches_configured_extension(path, extensions)
}

/// Read one entry out of the configured list into the lowercase form the lookups use.
///
/// Every name in this list is a bare extension — unlike the archive list, which has to carry the
/// dotted `tar.gz` — so anything that is not one is dropped rather than matched against.
pub fn sanitize_calibre_extensions(list: &str) -> Vec<String> {
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
/// on. The gate is asked beside it by the hook, the way every other kind's is.
pub fn is_calibre_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| matches_calibre_list(path, &config.calibre_extensions))
        .unwrap_or(false)
}

/// Whether a preview may be shown for `path`: the file the configured list claims, and the
/// `Calibre` gate in the tray's `Preview Types` submenu.
///
/// Both halves ask it where a kind can be switched off under a preview that is already on screen:
/// a hover is not sent for a kind that is off, and the layout places nothing for a file whose kind
/// is off, which is how a preview of that kind comes down when the switch does. See
/// `PreviewType::enabled`.
pub fn is_calibre_preview(path: &Path) -> bool {
    is_calibre_file(path) && PreviewType::Calibre.enabled()
}

/// Whether the engine is the one that reads this file at all: a name its own list carries, or the
/// bytes of a book it reads under a name no list holds — a `.mobi` renamed to `.dat`, say.
///
/// It is the question the request side asks before it asks the engine for a page — a book whose
/// bytes name another kind is not one to start it for, and one whose bytes name *this* kind is the
/// engine's whatever it is called — and it is the same question asked the same way
/// `libre_formats::engine_page_kind` asks it for the render engine and
/// `peazip_formats::is_engine_archive` for the listing engine: the file's own bytes first, and the
/// name it carries after them. What it is *not* is a gate: whether that kind may be shown is the
/// caller's to ask (`PreviewType::enabled`), because the same question is asked of a preview that
/// is already on screen when a switch is thrown.
pub fn is_engine_ebook(path: &Path) -> bool {
    use crate::formats::content_type::{self, Content};

    match content_type::of(path) {
        Content::Kind(PreviewType::Calibre) => true,
        Content::Kind(_) | Content::Foreign => false,
        Content::Unknown => is_calibre_file(path),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// What the list holds is what no reader here and no other list of this app's names: an
    /// archive this app reads itself, a document an engine of its own draws, a picture and a text
    /// file are all answered elsewhere, and none of them is here.
    #[test]
    fn holds_no_name_another_kind_already_reads() {
        let list = sanitize_calibre_extensions(DEFAULT_CALIBRE_EXTENSIONS);
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
                || crate::formats::peazip_formats::matches_peazip_list(
                    path,
                    &config.peazip_extensions,
                )
                || crate::formats::text_formats::matches_text_lists(
                    path,
                    &config.text_extensions,
                    &config.text_names,
                )
        };

        for name in [
            "book.rb",
            "comic.cbz",
            "letter.odt",
            "notes.txt",
            "page.html",
            "report.docx",
            "report.pdf",
            "styled.rtf",
        ] {
            let path = Path::new(name);
            assert!(
                claimed_elsewhere(path),
                "`{name}` is read by another kind, so this expectation is written the wrong way round"
            );
            assert!(
                !matches_calibre_list(path, &list),
                "`{name}` is read by another kind, so the engine is not asked about it"
            );
        }

        // A PDF is the one name of that kind's list the engine is still not asked about, because the
        // list holds it for this app's own reader: the PDF is the half of the book kind that is
        // drawn here, and a comic is the half that is read here (see `ebook_formats`).
        let pdf = Path::new("report.pdf");
        assert!(
            crate::readers::pdf_preview::is_pdf_file(pdf),
            "a PDF is this app's own to draw"
        );
        assert!(
            !matches_calibre_list(pdf, &list),
            "so the engine is not asked about one either"
        );

        // And the name this list gained with this build is its own: a Microsoft Reader book is a page
        // only this engine draws. Its neighbour `.chm` was here for a build and is the archiver's
        // again — the page is right and the two to three seconds it takes to draw is not, which is
        // the judgement `peazip_formats` carries the other half of.
        let book = Path::new("book.lit");
        assert!(
            matches_calibre_list(book, &list),
            "a Microsoft Reader book is a page only this engine can draw"
        );
        assert!(
            !crate::formats::peazip_formats::matches_peazip_list(book, &config.peazip_extensions),
            "and it is not the listing engine's any more: one name, one answer"
        );

        let help = Path::new("help.chm");
        assert!(
            crate::formats::peazip_formats::matches_peazip_list(help, &config.peazip_extensions),
            "a compiled help file is the listing engine's, which answers immediately"
        );
        assert!(
            !matches_calibre_list(help, &list),
            "and it is not this engine's: a help file is not a page to wait two seconds for"
        );

        // The Palm database is the one name two engines are about, and it is asked the way the app
        // asks it rather than by the list alone: a `.pdb` is a Palm ebook, which the render engine
        // reads by its own filters, and it is that list's — so the ebook engine is not asked about
        // one as well.
        let palm = Path::new("book.pdb");
        assert!(
            crate::formats::libre_formats::matches_libre_list(palm, &config.libre_extensions),
            "a Palm ebook is the render engine's, by the name that lists it"
        );
        assert!(
            !matches_calibre_list(palm, &list),
            "and not this engine's, so the two cannot both be asked about one file"
        );
    }

    /// And what it does hold is the formats an ebook arrives in that nothing else on the machine
    /// opens: the Kindle and Mobipocket family, the open one, the scanned book, and the containers
    /// of the dedicated readers.
    #[test]
    fn holds_the_ebooks_no_reader_here_opens() {
        let list = sanitize_calibre_extensions(DEFAULT_CALIBRE_EXTENSIONS);

        for name in [
            "book.azw",
            "book.azw3",
            "book.azw4",
            "book.mobi",
            "book.prc",
            "book.epub",
            "book.fb2",
            "book.djvu",
            "book.lrf",
            "book.htmlz",
            "book.pml",
            "book.snb",
            "book.tcr",
        ] {
            assert!(
                matches_calibre_list(Path::new(name), &list),
                "`{name}` is one of the engine's ebooks"
            );
        }
    }

    /// A name that is not a bare extension is dropped rather than matched against, the way every
    /// other list of this app's answers a hand-edited entry.
    #[test]
    fn reads_a_list_of_bare_extensions() {
        let extensions = sanitize_calibre_extensions(" .MOBI , epub,,book*.azw ,epub,tcr");

        assert_eq!(extensions, vec!["mobi", "epub", "tcr"]);
    }

    /// What the engine is asked about is asked of the file's own bytes before its name, which is
    /// what makes a book renamed to a name no list holds the engine's to read: a Mobipocket
    /// database under a `.dat` is one of its books, a picture is not, and a name nothing recognizes
    /// is left to the list the name is written in.
    #[test]
    fn asks_the_engine_about_a_file_by_its_bytes_before_its_name() {
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.calibre_extensions = sanitize_calibre_extensions(DEFAULT_CALIBRE_EXTENSIONS);
        }

        let folder = std::env::temp_dir().join("rust-hover-preview-calibre-engine-ebook");
        std::fs::create_dir_all(&folder).expect("a test folder");

        // A Mobipocket book: the Palm database header with the two identifiers every one of them
        // carries — the type `BOOK` and the creator `MOBI`, sixty and sixty-four bytes in.
        let renamed = folder.join("book.dat");
        std::fs::write(&renamed, mobipocket()).expect("a written book");
        assert!(
            is_engine_ebook(&renamed),
            "a book the engine reads is its to convert, whatever it is called"
        );

        // A picture under a book's name is a picture: what the file's own bytes say is the first
        // answer, and a reader here has it, so no engine is started for it.
        let picture = folder.join("cover.mobi");
        std::fs::write(&picture, [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A])
            .expect("a written picture");
        assert!(
            !is_engine_ebook(&picture),
            "a file whose bytes are another kind's is not the engine's"
        );

        // And the name's own list, for the formats whose bytes say nothing this table knows.
        std::fs::write(folder.join("book.tcr"), b"a compressed text, of a sort")
            .expect("a written file");
        assert!(
            is_engine_ebook(&folder.join("book.tcr")),
            "a name the list carries is the engine's, whatever the bytes say"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// The header of a Mobipocket file: the Palm database every Kindle and Palm ebook is, with the
    /// type and creator the format is defined by.
    fn mobipocket() -> Vec<u8> {
        let mut header = vec![0u8; 78 + 16];
        header[..4].copy_from_slice(b"Book");
        header[60..64].copy_from_slice(b"BOOK");
        header[64..68].copy_from_slice(b"MOBI");
        header[76..78].copy_from_slice(&1u16.to_be_bytes());
        header[78..82].copy_from_slice(b"MOBI");
        header
    }
}

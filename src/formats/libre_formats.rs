//! Which files are previewed by a render engine rather than by a reader of this app's own.
//!
//! LibreOffice imports a hundred-odd formats across four families — word processing,
//! spreadsheets, presentations and drawings — through the Document Liberation Project's
//! libraries and its own filters, and where it is installed this app asks it to draw one
//! as a PDF (see `libreoffice_render`). What is listed here is which of those names this
//! app hands over, and the list is deliberately narrower than the engine's reach:
//!
//! * A name another kind already reads is not here. Pictures, drawings, videos, archives,
//!   fonts, PDFs and the text lists are answered by this app's own readers, and asking an
//!   office suite about a `.png` or a `.md` would be a launch that answers worse than what
//!   is already there. The one exception is a document this app reads only the picture of:
//!   `cdr` is listed both here and in the design list, because what this app reads of one
//!   by itself is the thumbnail its program kept, and the engine draws the drawing.
//! * A name whose file is text a person can read is not here either. `.txt`, `.csv`,
//!   `.html`, `.rtf` and the rest of the text lists are shown by the text preview, which
//!   reads and highlights them; what belongs to an engine is the document that is packed
//!   or encrypted past reading — `.wpd`, `.wps`, `.abw`, `.doc` saved as one of those old
//!   formats, and the spreadsheets and presentations of the same vintage.
//!
//! The list of names this answers for is a row of `crate::formats::lists` — the one table
//! every kind's list is a row of, and the one place a list is written down, the built-in
//! entries and the older lists this app shipped and then changed included.

use crate::config::config::{OfficeEngine, PreviewType};
use crate::formats::lists;
use crate::CONFIG;
use std::path::Path;

/// Which kind a page the render engine draws for this file is shown under, or `None` for a
/// file the engine is not asked about at all.
///
/// Two kinds of document have a page that is the engine's rather than a reader's or an
/// application's. One is this list's own: the documents whose formats this app has no
/// reader for, named here or recognized by their own bytes. The other is an Office document
/// the engine is the one that draws — one whose own application is not installed, so there
/// is no engine of its own to ask for a page, and one the tray has asked the engine for
/// outright, so there is no other engine to ask. Either way the page is shown as the Office
/// document the file is rather than as a document of the engine's kind, which is what
/// `office_formats::page_engine` answers — the one place the choice and the machine are
/// read together, so that both sides here answer the same way.
///
/// The file's own bytes are asked first, as they are wherever a kind is settled: a picture
/// left under a document's name is not a document for this engine. Every Office format is a
/// container — a zip, an OLE compound file — and a container is not a kind here, so the
/// content never names Office and the name is what settles one, below. What the answer is
/// *not* is a gate: whether that kind may be shown is the caller's to ask
/// (`PreviewType::enabled`), because the same question is asked of a preview that is already
/// on screen when a switch is thrown.
pub fn engine_page_kind(path: &Path) -> Option<PreviewType> {
    use crate::formats::content_type::Content;

    // What follows the question asks the machine rather than a list — the name's own row, and
    // which of the two kinds the page belongs to — and takes its own answers. The question
    // itself reads the file's entry before taking the lock, so the guard is not held across
    // the read (see `content_type::of_reaching_config`), and this list comparison is in memory
    // under a guard of its own.
    let content = crate::formats::content_type::of_reaching_config(path);

    match content {
        Content::Kind(PreviewType::Libre) => return Some(PreviewType::Libre),
        Content::Kind(_) | Content::Foreign => return None,
        Content::Unknown => {}
    }

    if CONFIG
        .lock()
        .is_ok_and(|config| lists::LIBRE.claims(path, &config))
    {
        return Some(PreviewType::Libre);
    }

    // An Office document is the engine's only where the engine is the one that draws it: a
    // page that has an application to draw it is that application's, and asking the engine for
    // one as well would be a second rendering of the same document — the one thing this
    // fallback is not for (see `office_formats::page_engine`).
    matches!(
        crate::formats::office_formats::page_engine(path),
        Some(OfficeEngine::LibreOffice)
    )
    .then_some(PreviewType::Document)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// What the list holds is what this app has no reader for: a picture, a drawing, a
    /// font, a PDF and a text file are all answered elsewhere, and none of them is here.
    /// `cdr` is the one name that is in two lists on purpose — this app reads the picture a
    /// CorelDRAW file keeps, and the engine draws the drawing — and it is the reason the
    /// rule is worth stating rather than assuming.
    #[test]
    fn holds_no_name_another_kind_already_reads() {
        let config = crate::config::config::AppConfig::default();

        for name in [
            "photo.png",
            "photo.jpg",
            "texture.tga",
            "drawing.svg",
            "drawing.emf",
            "print.eps",
            "report.pdf",
            "notes.txt",
            "readme.md",
            "page.html",
            "data.csv",
            "styled.rtf",
            "report.docx",
            "book.xlsx",
            "slides.pptx",
            "font.ttf",
            "bundle.zip",
            "animation.swf",
        ] {
            let path = std::path::Path::new(name);
            let claimed = crate::formats::lists::IMAGE.claims(path, &config)
                || crate::formats::lists::VECTOR.claims(path, &config)
                || crate::formats::lists::OFFICE.claims(path, &config)
                || crate::formats::video_formats::matches_any_video_list(path, &config)
                || crate::formats::lists::TEXT.claims(path, &config)
                || crate::formats::lists::NAMES.claims(path, &config);

            if !claimed {
                continue;
            }

            assert!(
                !crate::formats::lists::LIBRE.claims(path, &config),
                "`{name}` is read by another kind, so the engine is not asked about it"
            );
        }

        assert!(
            crate::formats::lists::LIBRE.claims(std::path::Path::new("logo.cdr"), &config),
            "a CorelDRAW document is in this list as well as the design list: the engine \
             draws the drawing where this app reads the picture it keeps"
        );
    }

    /// And the names it does hold are the documents no other kind claims, from the engines
    /// of the nineties to the open formats of today.
    #[test]
    fn holds_the_documents_no_reader_of_this_app_takes() {
        let config = crate::config::config::AppConfig::default();

        for name in [
            "letter.wpd",
            "letter.wps",
            "letter.abw",
            "letter.sxw",
            "letter.lwp",
            "ledger.123",
            "ledger.wk4",
            "sheet.gnumeric",
            "talk.sdd",
            "talk.sxi",
            "poster.pub",
            "plan.vsdx",
            "drawing.dxf",
            "drawing.cdr",
            "document.odt",
            "sheet.ods",
            "slides.odp",
            "drawing.odg",
            "letter.pages",
            "book.odb",
            "drawing.wpg",
            "drawing.vstx",
            "drawing.pm6",
            "drawing.pmd",
        ] {
            assert!(
                crate::formats::lists::LIBRE.claims(std::path::Path::new(name), &config),
                "`{name}` is one of the engine's documents"
            );
        }
    }

    /// And the names the engine has no filter of its own for are not handed to it: a name
    /// like that is a launch that answers nothing, and one of them — a Flash file — is a
    /// launch that does not end. Each of these was read out of an installed engine's own
    /// filter registry, which is the only thing that settles what it answers for; the case
    /// that made it worth checking is beside them.
    #[test]
    fn hands_over_no_name_the_engine_has_no_filter_for() {
        let config = crate::config::config::AppConfig::default();

        for name in [
            "book.swf",
            "book.epub",
            "poster.qxp",
            "poster.pm3",
            "poster.pm4",
            "poster.pm5",
            "plan.vssm",
            "plan.vst",
            "plan.vstm",
            "plan.vtx",
            "plan.vsx",
            "note.agd",
            "note.fhd",
            "letter.jtd",
            "letter.jtt",
            "plot.plt",
            "picture.pxl",
            "text.rl",
            "talk.sdp",
            "drawing.sgf",
            "drawing.sgl",
            "letter.uof",
            "sheet.uop",
            "sheet.uos",
            "letter.uot",
            "drawing.vor",
        ] {
            assert!(
                !crate::formats::lists::LIBRE.claims(std::path::Path::new(name), &config),
                "`{name}` is a name no filter of the engine's declares"
            );
        }
    }

    /// A page the engine draws answers with the kind it is shown under, and the question is
    /// asked of the file's own bytes before its name: a picture left under a document's name
    /// is not a document for this engine, while a CorelDRAW drawing — one whose own container
    /// says what it is, whatever it is called — is.
    #[test]
    fn a_page_the_engine_draws_answers_with_the_kind_it_is_shown_under() {
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.libre_extensions = crate::formats::lists::sanitize_extension_list(
                crate::formats::lists::DEFAULT_LIBRE_EXTENSIONS,
            );
        }

        let folder = std::env::temp_dir().join("rust-hover-preview-engine-page-kind");
        std::fs::create_dir_all(&folder).expect("a test folder");

        // A drawing in CorelDRAW's own RIFF container, under a name no list claims at all:
        // what settles it is the file.
        let renamed = folder.join("drawing.bin");
        std::fs::write(&renamed, b"RIFF\x00\x00\x00\x00CDR6").expect("a written drawing");
        assert_eq!(
            engine_page_kind(&renamed),
            Some(PreviewType::Libre),
            "a CorelDRAW drawing is the engine's to draw, whatever it is called"
        );

        // And under the name its own list holds, which is what a real one is.
        assert_eq!(
            engine_page_kind(&folder.join("drawing.cdr")),
            Some(PreviewType::Libre),
            "a name of the engine's own list is its document too"
        );

        // A picture under a document's name is a picture: no engine is asked about one.
        let picture = folder.join("report.docx");
        std::fs::write(&picture, [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A])
            .expect("a written picture");
        assert_eq!(
            engine_page_kind(&picture),
            None,
            "a picture under a document's name is not a document for this engine"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// An Office document is one engine's or the other's, never both and never neither: the
    /// application that owns the format draws its page where it is installed, and the render
    /// engine beside it draws one where it is not. Both drawing it is what a **smaller
    /// version** of a page in front of the right one was: the engine's page was laid out at
    /// the wait's own size and shown for the second or two the application took to draw the
    /// real one over it. Neither drawing it is a hover that waits for a page nothing was
    /// asked to draw.
    ///
    /// The machine's own answer is what the expectation is written from — whether an
    /// application is installed is not something a test can decide — but the rule is one of
    /// this app's, and it is what this asserts.
    #[test]
    fn an_office_document_is_one_engine_or_the_others() {
        // What the app's own `config.ini` holds is not what this test is about: it asks the
        // machine, and the setting is pinned to the one that asks the machine.
        if let Ok(mut config) = CONFIG.lock() {
            config.office_engine = OfficeEngine::MicrosoftOffice;
        }

        let folder = std::env::temp_dir().join("rust-hover-preview-office-engine-choice");
        std::fs::create_dir_all(&folder).expect("a test folder");

        for name in ["report.docx", "ledger.xlsx", "talk.pptx"] {
            let path = folder.join(name);
            std::fs::write(&path, b"PK\x03\x04\x00\x00\x00\x00").expect("a written document");

            let installed = crate::formats::office_formats::app_installed(&path);
            assert_eq!(
                engine_page_kind(&path),
                (!installed).then_some(PreviewType::Document),
                "`{name}`: the render engine draws it exactly where the application that owns \
                 the format is not installed"
            );
        }

        let _ = std::fs::remove_dir_all(&folder);
    }
}

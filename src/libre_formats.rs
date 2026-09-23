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
//! The list lives in `config.ini` as `[libre] extensions`, written from the built-in list
//! on first run and read back from there, so a user can add a format the engine reads and
//! this app does not know, or take one out. The question here is only what a file is
//! *called*: whether the engine can read it at all is settled by the engine, and a name it
//! cannot read is answered with no preview — once, and then remembered, so a name that was
//! put in this list by mistake costs one conversion and never another.

use crate::text_formats;
use crate::CONFIG;
use std::path::Path;

/// The extensions written to `config.ini` on first run: the document formats the engine
/// reads that this app has no reader of its own for.
///
/// Nothing here is a picture, a drawing, a video, an archive, a font or a text file — those
/// are this app's own kinds — and nothing here is a PDF. What is here is what those kinds
/// leave: the word processors that came before the modern one and the ones beside it
/// (`wpd`, `wps`, `abw`, `lwp`, `cwk`, `hwp`, `602`, `wri`, `qxp`, `pm*`, `p65`), the older
/// spreadsheets (`123`, `wk1`, `wk3`, `wk4`, `wks`, `slk`, `dif`, `dbf`, `wb2`, `wq1`,
/// `wq2`, `gnumeric`, `xlw`), the presentations (`sda`, `sdd`, `sdp`, `sxi`, `sti`), the
/// drawings whose own format this app does not read (`cdr`, `cmx`, `dxf`, `wpg`, `sxd`,
/// `std`, `pub`, `vsd` and the Visio family, `fh*` under its `agd`, `fhd` and `zmf`,
/// `cgm`, `pct`, `plt`, `met`, `svm`, `sgf`) — and the open formats themselves, `odt`,
/// `ods`, `odp`, `odg`, `odc`, `odb`, `odf` and their friends, which no other list claims
/// and which the engine reads exactly.
///
/// The engine imports a few names that are deliberately not here, because there is nothing
/// in them to preview — each is a file *about* a document rather than one:
///
/// * `ase` and `gpl` are colour palettes. LibreOffice reads them to fill a colour picker,
///   and what a page would be drawn from one is nothing at all.
/// * `oxt` is an extension package: a zip of the files that install something into the
///   engine, which is not a document any more than a `.zip` is.
/// * `smf` means StarMath to LibreOffice and a MIDI sequence to everything else that reads
///   the name, and a hover cannot tell which one it has. A MIDI file claimed as a document
///   would be a launch that answers nothing, every time, for a name that is not this app's.
/// * `kth` is a Keynote theme and `iqy` is a web query: the first is a preset, the second
///   is a line of text naming a URL, and neither is a document.
pub const DEFAULT_LIBRE_EXTENSIONS: &str = "123,602,abw,agd,cdr,cgm,cmx,cwk,dbf,dif,dxf,epub,fhd,fodg,fodp,fodt,gnm,gnumeric,hwp,jtd,jtt,key,lwp,mcw,met,mw,numbers,odb,odc,odf,odg,odm,odp,ods,odt,oth,otg,otm,otp,ots,ott,pages,pcd,pct,pcx,pdb,plt,pm3,pm4,pm5,pm6,pmd,psw,pub,pxl,qxp,ras,rl,sda,sdc,sdd,sdp,sdw,sgf,sgl,slk,stc,std,sti,stw,svm,swf,sxd,sxg,sxi,sxm,sxw,uof,uop,uos,uot,vdx,vor,vsd,vsdm,vsdx,vssm,vst,vstm,vstx,vtx,vsx,wb2,wk1,wk3,wk4,wks,wpg,wq1,wq2,wpd,wps,wri,xlw,zabw,zmf";

/// Whether the configured list claims `path`.
pub fn matches_libre_list(path: &Path, extensions: &[String]) -> bool {
    text_formats::matches_configured_extension(path, extensions)
}

/// Read one entry out of the configured list into the lowercase form the lookups use.
///
/// Every name in this list is a bare extension — unlike the archive list, which has to
/// carry the dotted `tar.gz` — so anything that is not one is dropped rather than matched
/// against.
pub fn sanitize_libre_extensions(list: &str) -> Vec<String> {
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

/// Whether the configured list claims `path`, without asking whether these previews are
/// switched on. The gate is asked beside it by the hook, the way every other kind's is.
pub fn is_libre_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| matches_libre_list(path, &config.libre_extensions))
        .unwrap_or(false)
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
        let list = sanitize_libre_extensions(DEFAULT_LIBRE_EXTENSIONS);
        let config = crate::config::AppConfig::default();

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
        ] {
            let path = std::path::Path::new(name);
            let claimed = crate::image_formats::matches_image_list(&path, &config.image_extensions)
                || crate::vector_formats::matches_vector_list(&path, &config.vector_extensions)
                || crate::office_formats::matches_office_list(&path, &config.office_extensions)
                || crate::text_formats::matches_text_lists(
                    &path,
                    &config.text_extensions,
                    &config.text_names,
                );

            if !claimed {
                continue;
            }

            assert!(
                !matches_libre_list(path, &list),
                "`{name}` is read by another kind, so the engine is not asked about it"
            );
        }

        assert!(
            matches_libre_list(std::path::Path::new("logo.cdr"), &list),
            "a CorelDRAW document is in this list as well as the design list: the engine \
             draws the drawing where this app reads the picture it keeps"
        );
    }

    /// And the names it does hold are the documents no other kind claims, from the engines
    /// of the nineties to the open formats of today.
    #[test]
    fn holds_the_documents_no_reader_of_this_app_takes() {
        let list = sanitize_libre_extensions(DEFAULT_LIBRE_EXTENSIONS);

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
            "book.epub",
        ] {
            assert!(
                matches_libre_list(std::path::Path::new(name), &list),
                "`{name}` is one of the engine's documents"
            );
        }
    }
}

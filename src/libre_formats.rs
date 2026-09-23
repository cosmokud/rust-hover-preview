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
/// (`wpd`, `wps`, `abw`, `lwp`, `cwk`, `hwp`, `602`, `wri`), the older spreadsheets (`123`,
/// `wk1`, `wk3`, `wk4`, `wks`, `slk`, `dif`, `dbf`, `wb2`, `wq1`, `wq2`, `gnumeric`, `xlw`),
/// the presentations (`sda`, `sdc`, `sdd`, `sdw`, `sxi`, `sti`), the drawings whose own
/// format this app does not read (`cdr`, `cmx`, `dxf`, `wpg`, `pub`, `zmf`, `cgm`, `pct`,
/// `met`, `svm`, and the Visio names the engine's filter declares) — and the open formats
/// themselves, `odt`, `ods`, `odp`, `odg`, `odc`, `odb`, `odf` and their friends, which no
/// other list claims and which the engine reads exactly.
///
/// What a name is asked about is settled by the engine's own filter registry rather than by
/// the list of formats the engine is *said* to support, and those two are not the same list.
/// A name belongs here only where a filter that imports declares that very extension — read
/// one name at a time out of an installed engine's `share/registry/*.xcd` — because a name
/// no filter declares is a launch that answers nothing: what the engine falls back to is the
/// file's own content, where a filter it cannot use may spin rather than answer (see `swf`),
/// and where it answers at all it answers with nothing. Four groups an earlier list held
/// came out that way:
///
/// * `qxp` is the older QuarkXPress document. The filter reads `qxd` and `qxt`.
/// * `pm3`, `pm4` and `pm5` are PageMaker before 6. The filter reads `pm`, `p65`, `pm6` and
///   `pmd`.
/// * `vssm`, `vst`, `vstm`, `vtx` and `vsx` are Visio stencils and templates. The filter
///   reads `vdx`, `vsd`, `vsdm`, `vsdx` and `vstx`.
/// * `epub` is a name the engine *writes* rather than reads — the one filter that declares
///   it is an export filter, which is the wrong direction for a preview — and `agd`, `fhd`,
///   `jtd`, `jtt`, `plt`, `pxl`, `rl`, `sdp`, `sgf`, `sgl`, `uof`, `uop`, `uos`, `uot` and
///   `vor` are declared by no filter at all.
///
/// A machine whose engine does read one of those can have the name back by adding it: the
/// list is what the user edits, and a name added to it is asked about from the next read.
///
/// The engine imports a few names that are deliberately not here as well, because there is
/// nothing in them to preview — each is a file *about* a document rather than one:
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
/// * `swf` is a Flash animation rather than a document, and the engine does not draw one:
///   asked to convert one, its filter chain spins with a core at a hundred percent and
///   never writes a page — measured on real files, and past every bound a conversion is
///   given. It was in this list once, and a file of that name cost a launch and a core
///   for as long as the engine was left to it. What reads a Flash file is FFmpeg's own
///   SWF demuxer — the drawings, the sounds and the timeline of one — so the name is in
///   the video list, which is where it belongs, and is not here (see `video_formats`).
pub const DEFAULT_LIBRE_EXTENSIONS: &str = "123,602,abw,cdr,cgm,cmx,cwk,dbf,dif,dxf,fodg,fodp,fodt,gnm,gnumeric,hwp,key,lwp,mcw,met,mw,numbers,odb,odc,odf,odg,odm,odp,ods,odt,oth,otg,otm,otp,ots,ott,pages,pcd,pct,pcx,pdb,pm6,pmd,psw,pub,ras,sda,sdc,sdd,sdw,slk,stc,std,sti,stw,svm,sxd,sxg,sxi,sxm,sxw,vdx,vsd,vsdm,vsdx,vstx,wb2,wk1,wk3,wk4,wks,wpg,wq1,wq2,wpd,wps,wri,xlw,zabw,zmf";

/// The built-in `[libre]` list as it stood before the names the engine cannot read were
/// taken out of it: the one `swf` was in, and the four groups above with it.
///
/// A file holding exactly these entries is the app's own older list rather than a user's
/// edit — nobody has touched it — so it is brought up to the built-in list rather than kept
/// as written, which is what takes those names out of every `config.ini` already written.
/// What each of them cost is written beside it in the list above: a launch that answered
/// nothing, and for `swf` a launch that never ended at all.
pub const LIBRE_EXTENSIONS_WITH_THE_NAMES_THE_ENGINE_CANNOT_READ: &str = "123,602,abw,agd,cdr,cgm,cmx,cwk,dbf,dif,dxf,epub,fhd,fodg,fodp,fodt,gnm,gnumeric,hwp,jtd,jtt,key,lwp,mcw,met,mw,numbers,odb,odc,odf,odg,odm,odp,ods,odt,oth,otg,otm,otp,ots,ott,pages,pcd,pct,pcx,pdb,plt,pm3,pm4,pm5,pm6,pmd,psw,pub,pxl,qxp,ras,rl,sda,sdc,sdd,sdp,sdw,sgf,sgl,slk,stc,std,sti,stw,svm,swf,sxd,sxg,sxi,sxm,sxw,uof,uop,uos,uot,vdx,vor,vsd,vsdm,vsdx,vssm,vst,vstm,vstx,vtx,vsx,wb2,wk1,wk3,wk4,wks,wpg,wq1,wq2,wpd,wps,wri,xlw,zabw,zmf";

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

/// Whether a preview may be shown for `path`: the file the configured list claims, and
/// the `Libre` gate in the tray's `Preview Types` submenu.
///
/// Both halves ask it where a kind can be switched off under a preview that is already
/// on screen: a hover is not sent for a kind that is off, and a layout finds no size for
/// a file whose kind is off, which is how a preview of that kind comes down when the
/// switch does. See `PreviewType::enabled`.
pub fn is_libre_preview(path: &Path) -> bool {
    is_libre_file(path) && crate::config::PreviewType::Libre.enabled()
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
            "animation.swf",
        ] {
            let path = std::path::Path::new(name);
            let claimed = crate::image_formats::matches_image_list(&path, &config.image_extensions)
                || crate::vector_formats::matches_vector_list(&path, &config.vector_extensions)
                || crate::office_formats::matches_office_list(&path, &config.office_extensions)
                || crate::video_formats::matches_video_list(&path, &config.video_extensions)
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
            "letter.pages",
            "book.odb",
            "drawing.wpg",
            "drawing.vstx",
            "drawing.pm6",
            "drawing.pmd",
        ] {
            assert!(
                matches_libre_list(std::path::Path::new(name), &list),
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
        let list = sanitize_libre_extensions(DEFAULT_LIBRE_EXTENSIONS);

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
                !matches_libre_list(std::path::Path::new(name), &list),
                "`{name}` is a name no filter of the engine's declares"
            );
        }
    }
}

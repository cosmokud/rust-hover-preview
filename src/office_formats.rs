//! Which files are Office documents.
//!
//! The extension list lives in `config.ini`, written from the built-in list on
//! first run and read back from there, exactly as the text preview's lists and the
//! archive list are — so a user can add a format this list does not name, or take
//! one out, without a rebuild.
//!
//! The question here is only what a file is *called*. What it *is* — an OOXML
//! package or an OLE compound file — is settled by reading the file's own header,
//! and that split is deliberate: the hover gate asks its question of every item the
//! pointer touches, and a synchronizing provider's placeholder is a directory entry
//! that can be answered for but a file that must not be opened, because opening it
//! is what starts the download.

use crate::config::PreviewType;
use crate::text_formats;
use crate::CONFIG;
use std::fs::File;
use std::io::Read;
use std::path::Path;

/// The extensions written to `config.ini` on first run: the Word, Excel and
/// PowerPoint formats a hover is expected to meet, templates and slide shows
/// included.
pub const DEFAULT_OFFICE_EXTENSIONS: &str =
    "doc,docm,docx,dot,dotm,dotx,pot,potm,potx,pps,ppsm,ppsx,ppt,pptm,pptx,xls,xlsb,xlsm,xlsx,xlt,\
xltm,xltx";

/// The bytes a container is recognized by: an OOXML package is a zip, so it starts
/// with the local header of its first part, and a legacy document is an OLE
/// compound file with its own signature.
const CONTAINER_PROBE_BYTES: usize = 8;
const OOXML_MAGIC: [u8; 4] = [b'P', b'K', 0x03, 0x04];
const OLE_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/// How a document is put together, as its own header says.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OfficeContainer {
    /// A zip package holding the OOXML parts.
    Ooxml,
    /// An OLE compound file: Word 97-2003, Excel 97-2003 and PowerPoint 97-2003.
    Ole,
}

/// The application that renders a page for a document of this family, when the
/// render tier is asked for one.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OfficeApp {
    Word,
    Excel,
    PowerPoint,
}

impl OfficeApp {
    /// The automation ProgID the engine is created from.
    pub fn prog_id(self) -> &'static str {
        match self {
            Self::Word => "Word.Application",
            Self::Excel => "Excel.Application",
            Self::PowerPoint => "PowerPoint.Application",
        }
    }

    /// The name the application's processes run under, which is how the engine
    /// tells the instance it started from one the user already had open.
    pub fn image_name(self) -> &'static str {
        match self {
            Self::Word => "WINWORD.EXE",
            Self::Excel => "EXCEL.EXE",
            Self::PowerPoint => "POWERPNT.EXE",
        }
    }
}

/// Whether the configured list claims `path`.
pub fn matches_office_list(path: &Path, extensions: &[String]) -> bool {
    text_formats::matches_configured_extension(path, extensions)
}

/// Read one entry out of the configured list into the lowercase form the lookups
/// use.
///
/// Every name in the office list is a bare extension — unlike the archive list,
/// which has to carry the dotted `tar.gz` — so anything that is not one is dropped
/// rather than matched against.
pub fn sanitize_office_extensions(list: &str) -> Vec<String> {
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

/// Whether the configured list claims `path`, without asking whether Office
/// previews are switched on.
pub fn is_office_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| matches_office_list(path, &config.office_extensions))
        .unwrap_or(false)
}

/// Whether the file is previewed as an Office document under the current
/// configuration. The `Office` gate is checked on top of the list, so turning
/// Office previews off leaves the list alone and turning them back on restores it.
pub fn is_office_preview(path: &Path) -> bool {
    is_office_file(path) && PreviewType::Office.enabled()
}

/// Which application renders a document with this name, by the family its
/// extension belongs to.
///
/// The families are matched on their prefixes rather than on the built-in list, so
/// an extension a user added — `docx2`, say — still finds its engine, and a format
/// no engine claims is answered with no preview at all.
pub fn app_for(path: &Path) -> Option<OfficeApp> {
    let extension = path.extension()?.to_str()?.to_lowercase();

    if extension.starts_with("doc") || extension.starts_with("dot") {
        return Some(OfficeApp::Word);
    }
    if extension.starts_with("xls") || extension.starts_with("xlt") {
        return Some(OfficeApp::Excel);
    }
    if extension.starts_with("ppt") || extension.starts_with("pps") || extension.starts_with("pot")
    {
        return Some(OfficeApp::PowerPoint);
    }

    None
}

/// The box a document's page is measured by when nothing else can measure it, by
/// the shape that family's pages have: a Word page is a portrait sheet, a workbook's
/// first printed page is usually a landscape one, and a slide is a slide.
///
/// What a render is asked for is the room the preview may take rather than this, so
/// that a slide is exported at the width the display can show (see
/// `preview_window::request_office_render`). This is the answer for a document whose
/// page cannot be read — a file whose export was cut short, say — where a page-shaped
/// box still says more than a spinner's own.
pub fn default_page_size(path: &Path) -> (u32, u32) {
    match app_for(path) {
        Some(OfficeApp::Excel) => (1123, 794),
        Some(OfficeApp::PowerPoint) => (1280, 720),
        Some(OfficeApp::Word) | None => (794, 1123),
    }
}

/// The container the file's own header reports, or `None` for a file that is
/// neither — a text file wearing a document's name, say.
///
/// The render tier asks this before a path is handed to Office: what a document is
/// called is the hover gate's question and what it *is* is the engine's, but a file
/// that is not a document at all is not worth an Office start.
pub fn container_kind(path: &Path) -> Option<OfficeContainer> {
    let mut file = File::open(path).ok()?;
    let mut probe = [0u8; CONTAINER_PROBE_BYTES];
    let read = file.read(&mut probe).ok()?;

    if read >= OOXML_MAGIC.len() && probe[..OOXML_MAGIC.len()] == OOXML_MAGIC {
        return Some(OfficeContainer::Ooxml);
    }
    if read >= OLE_MAGIC.len() && probe == OLE_MAGIC {
        return Some(OfficeContainer::Ole);
    }

    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::PathBuf;

    fn list() -> Vec<String> {
        sanitize_office_extensions(DEFAULT_OFFICE_EXTENSIONS)
    }

    #[test]
    fn keeps_only_bare_extensions() {
        let extensions = list();
        assert!(extensions.contains(&"docx".to_string()));
        assert!(extensions.contains(&"xlsx".to_string()));
        assert!(extensions.contains(&"pptx".to_string()));

        let typed = sanitize_office_extensions(".DOCX, xlsm ,nonsense*,,, docx");
        assert_eq!(typed, vec!["docx", "xlsm"]);
    }

    #[test]
    fn matches_names_the_way_the_list_writes_them() {
        let extensions = list();
        assert!(matches_office_list(
            &PathBuf::from(r"C:\docs\report.docx"),
            &extensions
        ));
        assert!(matches_office_list(
            &PathBuf::from(r"C:\docs\Q3.XLSX"),
            &extensions
        ));
        assert!(matches_office_list(
            &PathBuf::from(r"C:\docs\deck.pptm"),
            &extensions
        ));
        assert!(!matches_office_list(
            &PathBuf::from(r"C:\docs\notes.txt"),
            &extensions
        ));
        assert!(!matches_office_list(
            &PathBuf::from(r"C:\docs\archive.zip"),
            &extensions
        ));
    }

    #[test]
    fn finds_an_engine_by_family() {
        assert_eq!(
            app_for(&PathBuf::from(r"C:\docs\report.docx")),
            Some(OfficeApp::Word)
        );
        assert_eq!(
            app_for(&PathBuf::from(r"C:\docs\legacy.doc")),
            Some(OfficeApp::Word)
        );
        assert_eq!(
            app_for(&PathBuf::from(r"C:\docs\book.xlsb")),
            Some(OfficeApp::Excel)
        );
        assert_eq!(
            app_for(&PathBuf::from(r"C:\docs\deck.pptx")),
            Some(OfficeApp::PowerPoint)
        );
        assert_eq!(app_for(&PathBuf::from(r"C:\docs\notes.txt")), None);
    }

    #[test]
    fn reads_the_container_from_the_file_itself() {
        // A folder of this module's own: the tests run beside each other, and one
        // of them clearing its fixtures must not take another's with it.
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-office-tests")
            .join("containers");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let ooxml = folder.join("probe-ooxml.bin");
        std::fs::write(&ooxml, [b'P', b'K', 0x03, 0x04, 0, 0, 0, 0]).expect("a written probe");
        let ole = folder.join("probe-ole.bin");
        std::fs::write(&ole, OLE_MAGIC).expect("a written probe");
        let other = folder.join("probe-other.bin");
        std::fs::write(&other, b"not a document at all").expect("a written probe");

        assert_eq!(container_kind(&ooxml), Some(OfficeContainer::Ooxml));
        assert_eq!(container_kind(&ole), Some(OfficeContainer::Ole));
        assert_eq!(container_kind(&other), None);
        assert_eq!(container_kind(Path::new("Z:\\does\\not\\exist.docx")), None);

        let _ = std::fs::remove_dir_all(&folder);
    }
}

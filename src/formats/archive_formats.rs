//! Which files are archives.
//!
//! The list of names this answers for is a row of `crate::formats::lists` — the one table every
//! kind's list is a row of, and the one place a list is written down, the built-in entries and the
//! older lists this app shipped and then changed included.
//!
//! What is kept here is the one thing about the archive list that is not a list: `tar.gz` is a
//! name rather than an extension. The last dot of `sources.tar.gz` is `gz`, which is not an
//! archive on its own, so an entry containing a dot is matched against the end of the file's name
//! instead of against its extension — and the sanitiser that lets the entry hold a dot at all is
//! the same one, which is why the two are one row rather than two rules that have to agree.
//!
//! The question here is otherwise only what a file is *called*. What it *is* — which of the
//! readers can open it — is settled in `archive_listing` by reading the file's own header, and
//! that split is deliberate: the hover gate asks its question of every item the pointer touches,
//! and a synchronizing provider's placeholder is a directory entry that can be answered for but a
//! file that must not be opened, because opening it is what starts the download.

use crate::config::config::PreviewType;
use crate::formats::text_formats;
use std::path::Path;

/// Whether either form of the configured list claims `path`: its last extension,
/// or a dotted tail of its name for the two-part formats.
pub fn matches_archive_list(path: &Path, extensions: &[String]) -> bool {
    if text_formats::matches_configured_extension(path, extensions) {
        return true;
    }

    let Some(name) = path.file_name().and_then(|name| name.to_str()) else {
        return false;
    };
    let name = name.to_ascii_lowercase();

    extensions
        .iter()
        .any(|extension| extension.contains('.') && name.ends_with(&format!(".{extension}")))
}

/// Whether the configured list claims `path`, without asking whether archive
/// previews are switched on.
pub fn is_archive_file(path: &Path) -> bool {
    crate::CONFIG
        .lock()
        .map(|config| matches_archive_list(path, &config.archive_extensions))
        .unwrap_or(false)
}

/// Whether the file is previewed as an archive under the current configuration.
/// The `Archives` gate is checked on top of the list, so turning archive previews
/// off leaves the list alone and turning them back on restores it.
pub fn is_archive_preview(path: &Path) -> bool {
    is_archive_file(path) && PreviewType::Archives.enabled()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formats::lists::{sanitize_archive_extension_list, DEFAULT_ARCHIVE_EXTENSIONS};
    use std::path::PathBuf;

    fn list() -> Vec<String> {
        sanitize_archive_extension_list(DEFAULT_ARCHIVE_EXTENSIONS)
    }

    #[test]
    fn keeps_the_dotted_entry_the_tarball_needs() {
        let extensions = list();
        assert!(extensions.contains(&"tar.gz".to_string()));
        assert!(extensions.contains(&"zip".to_string()));
        // A leading dot is what a user types; anything that is not an extension
        // is dropped rather than matched against.
        let typed = sanitize_archive_extension_list(".ZIP, tar.gz ,nonsense*,,docx");
        assert_eq!(typed, vec!["zip", "tar.gz", "docx"]);
    }

    #[test]
    fn matches_names_the_way_the_list_writes_them() {
        let extensions = list();
        assert!(matches_archive_list(
            &PathBuf::from(r"C:\downloads\release.zip"),
            &extensions
        ));
        assert!(matches_archive_list(
            &PathBuf::from(r"C:\downloads\sources.tar.gz"),
            &extensions
        ));
        assert!(matches_archive_list(
            &PathBuf::from(r"C:\downloads\sources.TAR.GZ"),
            &extensions
        ));
        assert!(matches_archive_list(
            &PathBuf::from(r"C:\downloads\archive.tgz"),
            &extensions
        ));
        // `tar` is in the list, `gzip` is not, and a bare `.gz` is not an
        // archive: its table is not in the file to read.
        assert!(!matches_archive_list(
            &PathBuf::from(r"C:\downloads\notes.gz"),
            &extensions
        ));
        assert!(!matches_archive_list(
            &PathBuf::from(r"C:\downloads\report.docx"),
            &extensions
        ));
    }
}

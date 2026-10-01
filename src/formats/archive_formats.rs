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

use crate::formats::text_formats;
use std::path::Path;

/// Whether the `[archive]` row's list claims `path`: its last extension, or a dotted tail of its
/// whole name for the two-part formats.
///
/// The list-taking form is what `lists` asks for and what its own test uses, because a caller that
/// has a list in hand — the row, or a test that built one — is a caller asking about that list
/// rather than about the archive kind. Everything else asks `routing::named_as`, which asks the
/// row and so asks this.
pub fn claims_in(path: &Path, extensions: &[String]) -> bool {
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

#[cfg(test)]
mod tests {
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
        let config = crate::config::config::AppConfig::default();

        assert!(crate::formats::lists::ARCHIVE
            .claims(&PathBuf::from(r"C:\downloads\release.zip"), &config));
        assert!(crate::formats::lists::ARCHIVE
            .claims(&PathBuf::from(r"C:\downloads\sources.tar.gz"), &config));
        assert!(crate::formats::lists::ARCHIVE
            .claims(&PathBuf::from(r"C:\downloads\sources.TAR.GZ"), &config));
        assert!(crate::formats::lists::ARCHIVE
            .claims(&PathBuf::from(r"C:\downloads\archive.tgz"), &config));
        // `tar` is in the list, `gzip` is not, and a bare `.gz` is not an
        // archive: its table is not in the file to read.
        assert!(!crate::formats::lists::ARCHIVE
            .claims(&PathBuf::from(r"C:\downloads\notes.gz"), &config));
        assert!(!crate::formats::lists::ARCHIVE
            .claims(&PathBuf::from(r"C:\downloads\report.docx"), &config));
    }
}

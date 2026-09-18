//! Which files are archives.
//!
//! The extension list lives in `config.ini`, written from the built-in list on
//! first run and read back from there, exactly as the text preview's lists are —
//! so a user can add the container formats this list does not name, or take one
//! out, without a rebuild.
//!
//! The question here is only what a file is *called*. What it *is* — which of the
//! readers can open it — is settled in `archive_listing` by reading the file's
//! own header, and that split is deliberate: the hover gate asks its question of
//! every item the pointer touches, and a synchronizing provider's placeholder is
//! a directory entry that can be answered for but a file that must not be opened,
//! because opening it is what starts the download.

use crate::config::PreviewType;
use crate::text_formats;
use crate::CONFIG;
use std::path::Path;

/// The extensions written to `config.ini` on first run: the archive formats a
/// hover is expected to meet, plus the zip containers that cost nothing extra to
/// list (`.jar`, `.apk`, `.xpi`, `.cbz` are all zips).
///
/// `tar.gz` is a name rather than an extension — the last dot of `sources.tar.gz`
/// is `gz`, which is not an archive on its own — so an entry containing a dot is
/// matched against the end of the file's name instead.
pub const DEFAULT_ARCHIVE_EXTENSIONS: &str = "7z,apk,cbz,jar,rar,tar,tar.gz,tgz,xpi,zip,zipx";

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

/// Read one entry out of the configured list into the lowercase form the lookups
/// use.
///
/// The text lists drop an entry that is not a bare extension; here a dot is part
/// of the vocabulary, because the entry that names a tarball has to be `tar.gz` —
/// `gz` alone is a compressed file, not an archive — and is matched against the
/// end of the name rather than against the last extension.
pub fn sanitize_archive_extensions(list: &str) -> Vec<String> {
    let mut extensions: Vec<String> = Vec::new();

    for entry in list.split(',') {
        let trimmed = entry.trim().trim_start_matches('.').to_lowercase();
        let is_extension = !trimmed.is_empty()
            && !trimmed.starts_with('.')
            && !trimmed.ends_with('.')
            && trimmed
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || matches!(c, '.' | '+' | '-' | '_' | '#'));

        if is_extension && !extensions.contains(&trimmed) {
            extensions.push(trimmed);
        }
    }

    extensions
}

/// Whether the configured list claims `path`, without asking whether archive
/// previews are switched on.
pub fn is_archive_file(path: &Path) -> bool {
    CONFIG
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

    fn list() -> Vec<String> {
        sanitize_archive_extensions(DEFAULT_ARCHIVE_EXTENSIONS)
    }

    #[test]
    fn keeps_the_dotted_entry_the_tarball_needs() {
        let extensions = list();
        assert!(extensions.contains(&"tar.gz".to_string()));
        assert!(extensions.contains(&"zip".to_string()));
        // A leading dot is what a user types; anything that is not an extension
        // is dropped rather than matched against.
        let typed = sanitize_archive_extensions(".ZIP, tar.gz ,nonsense*,,docx");
        assert_eq!(typed, vec!["zip", "tar.gz", "docx"]);
    }

    #[test]
    fn matches_names_the_way_the_list_writes_them() {
        use std::path::PathBuf;

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

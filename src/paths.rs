//! Paths in the spelling the Shell will take — and the one folder a run keeps its own files in.
//!
//! The Explorer hook canonicalizes to the verbatim `\\?\` form, which is not a legal
//! thing to hand a Shell call, a browser URL or a media source: a file named in that
//! form is a file none of them can find. The prefix comes off before anything is
//! pointed at a file, and a share keeps its server.

use std::path::{Path, PathBuf};
use std::sync::OnceLock;

/// The folder this run keeps `config.ini`, `theme`, `engines`, `webview`, `libreoffice` and
/// `.env.local` in: the exe's own folder where a `.portable` marker stands beside it — kept in
/// `rust-hover-preview` there, so an unzip does not scatter the app's files among the user's —
/// and the machine's own folders otherwise.
///
/// One root for all of them rather than a folder per kind, because what a portable build promises
/// is that nothing of it is left on the machine — and a folder split across `%APPDATA%` and
/// `%LOCALAPPDATA%` is a promise kept by six call sites agreeing instead of by one.
///
/// A test build never reads the machine's folders here: each of the six keeps its own test folder,
/// and every one of them answers from that before asking this at all.
pub(crate) fn data_root() -> Option<&'static Path> {
    static ROOT: OnceLock<Option<PathBuf>> = OnceLock::new();
    ROOT.get_or_init(portable_root).as_deref()
}

#[cfg(test)]
fn portable_root() -> Option<PathBuf> {
    None
}

#[cfg(not(test))]
fn portable_root() -> Option<PathBuf> {
    let dir = std::env::current_exe().ok()?.parent()?.to_path_buf();
    // `.portable` beside the exe, the way a portable editor says so: its presence is the whole of
    // what it says, whatever it holds.
    dir.join(".portable")
        .is_file()
        .then(|| dir.join("rust-hover-preview"))
}

/// Whether this run is the portable one: a `.portable` file beside the exe, the way a portable
/// editor says so.
pub(crate) fn portable() -> bool {
    data_root().is_some()
}

/// A path with the Shell's verbatim prefix taken off.
pub(crate) fn plain_path(path: &Path) -> String {
    let text = path.to_string_lossy();

    match text.strip_prefix(r"\\?\UNC\") {
        Some(rest) => format!(r"\\{rest}"),
        None => text.strip_prefix(r"\\?\").unwrap_or(&text).to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_run_with_no_portable_marker_keeps_the_machines_own_folders() {
        // Every test answers from a folder of its own before this is asked at all, and the marker
        // is a file beside the exe a test build does not put one beside.
        assert!(data_root().is_none());
        assert!(!portable());
    }

    #[test]
    fn a_verbatim_path_is_read_as_the_spelling_the_shell_will_take() {
        assert_eq!(
            plain_path(Path::new(r"\\?\C:\docs\report.docx")),
            r"C:\docs\report.docx"
        );
        assert_eq!(
            plain_path(Path::new(r"\\?\UNC\server\share\report.docx")),
            r"\\server\share\report.docx"
        );
        // A path that was never verbatim is already the right spelling.
        assert_eq!(
            plain_path(Path::new(r"C:\docs\report.docx")),
            r"C:\docs\report.docx"
        );
        assert_eq!(
            plain_path(Path::new(r"\\server\share\track.mp3")),
            r"\\server\share\track.mp3"
        );
    }
}

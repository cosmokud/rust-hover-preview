//! Paths in the spelling the Shell will take.
//!
//! The Explorer hook canonicalizes to the verbatim `\\?\` form, which is not a legal
//! thing to hand a Shell call, a browser URL or a media source: a file named in that
//! form is a file none of them can find. The prefix comes off before anything is
//! pointed at a file, and a share keeps its server.

use std::path::Path;

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

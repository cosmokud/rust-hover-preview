//! Which files are vector drawings.
//!
//! The extension list lives in `config.ini`, written from the built-in list on first run
//! and read back from there, exactly as the text preview's lists and the archive, office,
//! font and design lists are — so a user can add a format this list does not name, or
//! take one out, without a rebuild.
//!
//! The question here is only what a file is *called*. What it *is* — a Windows metafile,
//! which the drawing layer plays, or an encapsulated PostScript file carrying a preview —
//! is settled in `metafile_image` and `eps_image` by reading the file's own header, and
//! that split is deliberate: the hover gate asks its question of every item the pointer
//! touches, and a synchronizing provider's placeholder is a directory entry that can be
//! answered for but a file that must not be opened, because opening it is what starts
//! the download.

use crate::config::config::PreviewType;
use crate::formats::text_formats;
use crate::CONFIG;
use std::path::Path;

/// The extensions written to `config.ini` on first run: the drawings a hover is expected
/// to meet.
///
/// `svg` and `svgz` are the documents the browser engine draws, and both spellings of the
/// same format; `wmf` and `emf` are Windows' two metafiles — a list of drawing records that
/// the drawing layer plays back, which is why a preview of one is sharp at any size — and
/// `eps` and `epsi` are the two spellings of an encapsulated PostScript file, read here
/// for the preview picture a writer leaves inside it rather than for the PostScript
/// itself: nothing in this app interprets PostScript.
pub const DEFAULT_VECTOR_EXTENSIONS: &str = "emf,eps,epsi,svg,svgz,wmf";

/// The built-in vector list as it stood before `svg` and `svgz` were added to it — which
/// is to say the list the kind had when it was written.
///
/// A file holding exactly these entries is the app's own older list rather than a user's
/// edit — nobody has touched it — so it is brought up to the built-in list rather than kept
/// as written. Without that, the two documents would be left out of every `config.ini`
/// already written, and a list that differs is otherwise the user's own (see
/// `config::configured_list_over_history`).
pub const VECTOR_EXTENSIONS_BEFORE_SVG: &str = "emf,eps,epsi,wmf";

/// Whether the configured list claims `path`.
pub fn matches_vector_list(path: &Path, extensions: &[String]) -> bool {
    text_formats::matches_configured_extension(path, extensions)
}

/// Read one entry out of the configured list into the lowercase form the lookups use.
///
/// Every name in the vector list is a bare extension — unlike the archive list, which has
/// to carry the dotted `tar.gz` — so anything that is not one is dropped rather than
/// matched against.
pub fn sanitize_vector_extensions(list: &str) -> Vec<String> {
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

/// Whether the configured list claims `path`, without asking whether vector previews are
/// switched on.
pub fn is_vector_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| matches_vector_list(path, &config.vector_extensions))
        .unwrap_or(false)
}

/// Whether the file is previewed as a vector drawing under the current configuration. The
/// `Vector` gate is checked on top of the list, so turning vector previews off leaves the
/// list alone and turning them back on restores it.
pub fn is_vector_preview(path: &Path) -> bool {
    is_vector_file(path) && PreviewType::Vector.enabled()
}

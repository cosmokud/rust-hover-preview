// Which files are vector drawings.
//
// The list of names this answers for is a row of `crate::formats::lists` — the one table every
// kind's list is a row of, and the one place a list is written down. What is left here is the
// question the list cannot answer, which is the same question every other kind's file asks.
//
// What a file *is* — a Windows metafile, which the drawing layer plays, or an encapsulated
// PostScript file carrying a preview — is settled in `metafile_image` and `eps_image` by reading
// the file's own header, and that split is deliberate: the hover gate asks its question of every
// item the pointer touches, and a synchronizing provider's placeholder is a directory entry that
// can be answered for but a file that must not be opened, because opening it is what starts the
// download.

use crate::config::config::PreviewType;
use std::path::Path;

/// Whether the configured list claims `path`.
pub fn matches_vector_list(path: &Path, extensions: &[String]) -> bool {
    crate::formats::text_formats::matches_configured_extension(path, extensions)
}

/// Whether the configured list claims `path`, without asking whether vector previews are
/// switched on.
pub fn is_vector_file(path: &Path) -> bool {
    crate::CONFIG
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

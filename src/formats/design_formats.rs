// Which files are design documents and projects.
//
// The list of names this answers for is a row of `crate::formats::lists` — the one table every
// kind's list is a row of, and the one place a list is written down. What is left here is the
// question the list cannot answer, which is the same question every other kind's file asks.
//
// What a file *is* — a Photoshop document with a whole picture at the end of it, a Krita or
// OpenRaster project holding the flattened document as a picture, or a container this app has no
// reader for — is settled in `psd_image` and `project_image` by reading the file's own header,
// and that split is deliberate: the hover gate asks its question of every item the pointer
// touches, and a synchronizing provider's placeholder is a directory entry that can be answered
// for but a file that must not be opened, because opening it is what starts the download.

use crate::config::config::PreviewType;
use std::path::Path;

/// Whether the configured list claims `path`.
pub fn matches_design_list(path: &Path, extensions: &[String]) -> bool {
    crate::formats::text_formats::matches_configured_extension(path, extensions)
}

/// Whether the configured list claims `path`, without asking whether design
/// previews are switched on.
pub fn is_design_file(path: &Path) -> bool {
    crate::CONFIG
        .lock()
        .map(|config| matches_design_list(path, &config.design_extensions))
        .unwrap_or(false)
}

/// Whether the file is previewed as a design document under the current
/// configuration. The `Design` gate is checked on top of the list, so turning design
/// previews off leaves the list alone and turning them back on restores it.
pub fn is_design_preview(path: &Path) -> bool {
    is_design_file(path) && PreviewType::Design.enabled()
}

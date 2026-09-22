//! Which files are design documents and projects.
//!
//! The extension list lives in `config.ini`, written from the built-in list on
//! first run and read back from there, exactly as the text preview's lists and the
//! archive, office and font lists are — so a user can add a format this list does
//! not name, or take one out, without a rebuild.
//!
//! The question here is only what a file is *called*. What it *is* — a Photoshop
//! document with a whole picture at the end of it, a Krita or OpenRaster project
//! holding the flattened document as a picture, or a container this app has no
//! reader for — is settled in `psd_image` and `project_image` by reading the file's
//! own header, and that split is deliberate: the hover gate asks its question of
//! every item the pointer touches, and a synchronizing provider's placeholder is a
//! directory entry that can be answered for but a file that must not be opened,
//! because opening it is what starts the download.

use crate::config::PreviewType;
use crate::text_formats;
use crate::CONFIG;
use std::path::Path;

/// The extensions written to `config.ini` on first run: the layered documents and
/// the project containers a hover is expected to meet.
///
/// `psd` and `psb` are Photoshop's two, read for the merged picture the format keeps
/// at the end of the file rather than for the layers; `kra` and `ora` are Krita's
/// and OpenRaster's, which are zip containers holding that same picture as a file of
/// its own; `cdr` is CorelDRAW's and `procreate` is Procreate's, which are containers
/// of the same kind holding the picture the application wrote for a file manager; and
/// `sketch`, `fig` and `xd` are containers of that kind as well. What each one is
/// read for is in `psd_image`, `project_image`, `cdr_image` and `eps_image`, and a
/// container none of them can open is a file that shows no preview, like any other
/// format this app has no reader for.
///
/// `ai` is the one name here that is usually not this kind's at all: an Illustrator
/// document saved with `Create PDF Compatible File` is a PDF, and the PDF gate claims
/// it before this list is asked. What reaches here under that name is a document saved
/// without that compatibility, which is an encapsulated PostScript file — the artwork
/// as a program, with the preview an older Illustrator left beside it — and what a
/// preview of one can be is what that preview holds. A document that carries none shows
/// nothing, which is the answer every file with no reader gets.
pub const DEFAULT_DESIGN_EXTENSIONS: &str = "ai,cdr,fig,kra,ora,procreate,psb,psd,sketch,xd";

/// The built-in design list as it stood before the two drawing applications that write a
/// container of their own — `cdr`, CorelDRAW's, and `procreate`, Procreate's — were added
/// to it.
///
/// A file holding exactly these entries is the app's own older list rather than a user's
/// edit — nobody has touched it — so it is brought up to the built-in list rather than kept
/// as written. Without that, the two names would reach a fresh installation only: every
/// `config.ini` already written holds a list that differs from the built-in one, and a list
/// that differs is otherwise the user's own (see `config::configured_list_over_history`).
pub const DESIGN_EXTENSIONS_BEFORE_CDR_AND_PROCREATE: &str = "ai,fig,kra,ora,psb,psd,sketch,xd";

/// The built-in design list as it stood before `ai` was added to it.
///
/// A file holding exactly these entries is the app's own older list rather than a user's
/// edit — nobody has touched it — so it is brought up to the built-in list rather than kept
/// as written. Without that, the entry would reach a fresh installation only: every
/// `config.ini` already written holds a list that differs from the built-in one, and a list
/// that differs is otherwise the user's own (see `config::configured_list_over_history`).
pub const DESIGN_EXTENSIONS_BEFORE_AI: &str = "fig,kra,ora,psb,psd,sketch,xd";

/// Whether the configured list claims `path`.
pub fn matches_design_list(path: &Path, extensions: &[String]) -> bool {
    text_formats::matches_configured_extension(path, extensions)
}

/// Read one entry out of the configured list into the lowercase form the lookups
/// use.
///
/// Every name in the design list is a bare extension — unlike the archive list,
/// which has to carry the dotted `tar.gz` — so anything that is not one is dropped
/// rather than matched against.
pub fn sanitize_design_extensions(list: &str) -> Vec<String> {
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

/// Whether the configured list claims `path`, without asking whether design
/// previews are switched on.
pub fn is_design_file(path: &Path) -> bool {
    CONFIG
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

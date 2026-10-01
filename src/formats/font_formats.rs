//! Whether a file is a font, for the three call sites in `webview_preview` that ask it.
//!
//! The list of names this answers for is the `[font]` row of `crate::formats::lists` — the one
//! table every kind's list is a row of, and the one place a list is written down. What is left
//! here is a wrapper over that row, and it is left here rather than asked of `routing` because
//! the browser engine's own three call sites reach it by this path and that file is outside this
//! change's ownership: a re-export would be a module whose only content is a name, which is
//! what every one of the fourteen used to be and what four of them no longer are.
//!
//! What a file *is* — a TrueType or CFF outline font, a collection of faces, or one of the two
//! webfont containers — is settled in `font_preview` by reading the file's own header, and that
//! split is deliberate: the hover gate asks its question of every item the pointer touches, and a
//! synchronizing provider's placeholder is a directory entry that can be answered for but a file
//! that must not be opened, because opening it is what starts the download.

use crate::formats::lists;
use crate::CONFIG;
use std::path::Path;

/// Whether the configured list claims `path`, without asking whether font previews are switched
/// on.
///
/// It takes the configuration's lock itself, which is the shape the rest of this layer has been
/// taken out of and this one has not: a caller that holds the lock asks `routing::named_as`
/// instead, which is a lookup against the row above and a guard taken once rather than by
/// whoever happened to be nearest the file.
pub fn is_font_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| lists::FONT.claims(path, &config))
        .unwrap_or(false)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The font list is matched the way the lists are written, which is the question every row of
    /// the table is asked through — and it is asked here rather than through [`is_font_file`],
    /// because a test that took the global lock to ask about the built-in list would be testing
    /// whatever `config.ini` the developer running it happens to have.
    ///
    /// The last case is the one the lookup's own rule exists for: a file whose *name* is an
    /// extension is a file with no extension at all, and matching it would claim a folder of
    /// extensions as though they were fonts.
    #[test]
    fn matches_names_the_way_the_list_writes_them() {
        let config = crate::config::config::AppConfig::default();

        for name in [
            r"C:\fonts\Inter-Regular.woff2",
            r"C:\fonts\NotoSansJP.OTF",
            r"C:\fonts\msgothic.ttc",
            r"C:\fonts\body.otf",
        ] {
            assert!(
                lists::FONT.claims(std::path::Path::new(name), &config),
                "`{name}` is one of the five formats the specimen can be given"
            );
        }

        for name in [
            r"C:\fonts\readme.txt",
            r"C:\fonts\photo.png",
            r"C:\fonts\Inter-Regular.eot",
            r"C:\fonts\ttf",
        ] {
            assert!(
                !lists::FONT.claims(std::path::Path::new(name), &config),
                "`{name}` is not one of them"
            );
        }
    }
}

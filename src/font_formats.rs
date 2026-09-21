//! Which files are fonts.
//!
//! The extension list lives in `config.ini`, written from the built-in list on
//! first run and read back from there, exactly as the text preview's lists and the
//! archive and office lists are — so a user can add a format this list does not
//! name, or take one out, without a rebuild.
//!
//! The question here is only what a file is *called*. What it *is* — a TrueType or
//! CFF outline font, a collection of faces, or one of the two webfont containers —
//! is settled in `font_preview` by reading the file's own header, and that split is
//! deliberate: the hover gate asks its question of every item the pointer touches,
//! and a synchronizing provider's placeholder is a directory entry that can be
//! answered for but a file that must not be opened, because opening it is what
//! starts the download.

use crate::config::PreviewType;
use crate::text_formats;
use crate::CONFIG;
use std::path::Path;

/// The extensions written to `config.ini` on first run: the font formats a hover is
/// expected to meet.
///
/// The two webfont containers and the three desktop ones, which are the same
/// outlines in different wrappers. What draws one is the browser engine the SVG
/// previews already use — it reads all five, and a `.ttc` is the one of them it
/// cannot be *pointed* at, which is why that face is written out as a font of its
/// own on this side; see `webview_preview` and `font_preview`. What this side reads
/// of them is two tables, the character map and the name, and the drawing is the
/// engine's.
pub const DEFAULT_FONT_EXTENSIONS: &str = "otf,ttc,ttf,woff,woff2";

/// Whether the configured list claims `path`.
pub fn matches_font_list(path: &Path, extensions: &[String]) -> bool {
    text_formats::matches_configured_extension(path, extensions)
}

/// Read one entry out of the configured list into the lowercase form the lookups
/// use.
///
/// Every name in the font list is a bare extension — unlike the archive list, which
/// has to carry the dotted `tar.gz` — so anything that is not one is dropped rather
/// than matched against.
pub fn sanitize_font_extensions(list: &str) -> Vec<String> {
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

/// Whether the configured list claims `path`, without asking whether font previews
/// are switched on.
pub fn is_font_file(path: &Path) -> bool {
    CONFIG
        .lock()
        .map(|config| matches_font_list(path, &config.font_extensions))
        .unwrap_or(false)
}

/// Whether the file is previewed as a font under the current configuration. The
/// `Fonts` gate is checked on top of the list, so turning font previews off leaves
/// the list alone and turning them back on restores it.
pub fn is_font_preview(path: &Path) -> bool {
    is_font_file(path) && PreviewType::Fonts.enabled()
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::PathBuf;

    fn list() -> Vec<String> {
        sanitize_font_extensions(DEFAULT_FONT_EXTENSIONS)
    }

    #[test]
    fn keeps_only_bare_extensions() {
        let extensions = list();
        assert!(extensions.contains(&"ttf".to_string()));
        assert!(extensions.contains(&"otf".to_string()));
        assert!(extensions.contains(&"ttc".to_string()));
        assert!(extensions.contains(&"woff".to_string()));
        assert!(extensions.contains(&"woff2".to_string()));

        let typed = sanitize_font_extensions(".TTF, woff2 ,nonsense*,,ttf,.Font-Regular.otf");
        assert_eq!(typed, vec!["ttf", "woff2"]);
    }

    #[test]
    fn matches_names_the_way_the_list_writes_them() {
        let extensions = list();
        assert!(matches_font_list(
            &PathBuf::from(r"C:\fonts\Inter-Regular.woff2"),
            &extensions
        ));
        assert!(matches_font_list(
            &PathBuf::from(r"C:\fonts\NotoSansJP.OTF"),
            &extensions
        ));
        assert!(matches_font_list(
            &PathBuf::from(r"C:\fonts\msgothic.ttc"),
            &extensions
        ));
        assert!(!matches_font_list(
            &PathBuf::from(r"C:\fonts\readme.txt"),
            &extensions
        ));
        assert!(!matches_font_list(
            &PathBuf::from(r"C:\fonts\photo.png"),
            &extensions
        ));
        // A name that merely ends in one of them is not one of them.
        assert!(!matches_font_list(
            &PathBuf::from(r"C:\fonts\ttf"),
            &extensions
        ));
    }
}

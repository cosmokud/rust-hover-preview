// Which files are fonts.
//
// The list of names this answers for is a row of `crate::formats::lists` — the one table every
// kind's list is a row of, and the one place a list is written down. What is left here is the
// question the list cannot answer, which is the same question every other kind's file asks.
//
// What a file *is* — a TrueType or CFF outline font, a collection of faces, or one of the two
// webfont containers — is settled in `font_preview` by reading the file's own header, and that
// split is deliberate: the hover gate asks its question of every item the pointer touches, and a
// synchronizing provider's placeholder is a directory entry that can be answered for but a file
// that must not be opened, because opening it is what starts the download.

use crate::config::config::PreviewType;
use std::path::Path;

/// Whether the configured list claims `path`.
pub fn matches_font_list(path: &Path, extensions: &[String]) -> bool {
    crate::formats::text_formats::matches_configured_extension(path, extensions)
}

/// Whether the configured list claims `path`, without asking whether font previews
/// are switched on.
pub fn is_font_file(path: &Path) -> bool {
    crate::CONFIG
        .lock()
        .map(|config| matches_font_list(path, &config.font_extensions))
        .unwrap_or(false)
}

/// Whether the file is previewed as a font under the current configuration. The
/// `Fonts` gate is checked on top of the list, so turning font previews off leaves the
/// list alone and turning them back on restores it.
pub fn is_font_preview(path: &Path) -> bool {
    is_font_file(path) && PreviewType::Fonts.enabled()
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::PathBuf;

    fn list() -> Vec<String> {
        crate::formats::lists::sanitize_extension_list(
            crate::formats::lists::DEFAULT_FONT_EXTENSIONS,
        )
    }

    #[test]
    fn keeps_only_bare_extensions() {
        let extensions = list();
        assert!(extensions.contains(&"ttf".to_string()));
        assert!(extensions.contains(&"otf".to_string()));
        assert!(extensions.contains(&"ttc".to_string()));
        assert!(extensions.contains(&"woff".to_string()));
        assert!(extensions.contains(&"woff2".to_string()));

        let typed = crate::formats::lists::sanitize_extension_list(
            ".TTF, woff2 ,nonsense*,,ttf,.Font-Regular.otf",
        );
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

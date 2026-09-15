//! The two bundled color themes, and the scope lookups the renderer uses to ask
//! them for a color.
//!
//! Both files are TextMate themes generated from the VS Code themes of the same
//! name (see `assets/themes/NOTICE.md`) and embedded in the binary, so the
//! preview depends on no file on disk and no installation step. Each is parsed
//! once, on first use, and the parse result — including a failure — is kept for
//! the lifetime of the process.

use crate::config::TextTheme;
use once_cell::sync::Lazy;
use std::io::Cursor;
use syntect::highlighting::{Color, Highlighter, Style, Theme, ThemeSet};
use syntect::parsing::Scope;

const ATOM_ONE_LIGHT: &str = include_str!("../assets/themes/atom-one-light.tmTheme");
const ONE_DARK_PRO: &str = include_str!("../assets/themes/one-dark-pro.tmTheme");

/// A parsed theme plus the scope resolver built from it.
///
/// `Highlighter` borrows the theme it was built from, so the theme is leaked
/// once: it lives as long as the process and is never rebuilt.
pub struct LoadedTheme {
    theme: &'static Theme,
    highlighter: Highlighter<'static>,
}

impl LoadedTheme {
    pub fn theme(&self) -> &'static Theme {
        self.theme
    }

    /// Background color of the preview page.
    pub fn background(&self) -> Color {
        self.theme.settings.background.unwrap_or(Color {
            r: 255,
            g: 255,
            b: 255,
            a: 255,
        })
    }

    /// Color of text that no scope rule claims.
    pub fn foreground(&self) -> Color {
        self.theme.settings.foreground.unwrap_or(Color {
            r: 0,
            g: 0,
            b: 0,
            a: 255,
        })
    }

    /// Style for a TextMate scope stack, such as `markup.heading` followed by
    /// `markup.heading.1.markdown`. A scope the theme says nothing about falls
    /// back to the theme's own foreground, which is what the renderer wants, and
    /// a scope the parser rejects is simply left out of the stack.
    pub fn style_for_scopes(&self, scopes: &[&str]) -> Style {
        let stack: Vec<Scope> = scopes
            .iter()
            .filter_map(|scope| Scope::new(scope).ok())
            .collect();
        self.highlighter.style_for_stack(&stack)
    }
}

fn load(source: &'static str) -> Option<LoadedTheme> {
    let theme = ThemeSet::load_from_reader(&mut Cursor::new(source.as_bytes())).ok()?;
    let theme: &'static Theme = Box::leak(Box::new(theme));

    Some(LoadedTheme {
        theme,
        highlighter: Highlighter::new(theme),
    })
}

static LIGHT: Lazy<Option<LoadedTheme>> = Lazy::new(|| load(ATOM_ONE_LIGHT));
static DARK: Lazy<Option<LoadedTheme>> = Lazy::new(|| load(ONE_DARK_PRO));

/// The bundled theme for `kind`. `None` — which the bundled files make
/// impossible in practice — drops the preview rather than painting it with
/// guessed colors.
pub fn loaded(kind: TextTheme) -> Option<&'static LoadedTheme> {
    match kind {
        TextTheme::Light => LIGHT.as_ref(),
        TextTheme::Dark => DARK.as_ref(),
    }
}

#[cfg(test)]
mod tests {
    use super::{loaded, DARK, LIGHT};
    use crate::config::TextTheme;

    #[test]
    fn both_bundled_themes_parse() {
        assert!(LIGHT.is_some(), "the light theme did not parse");
        assert!(DARK.is_some(), "the dark theme did not parse");
    }

    #[test]
    fn each_theme_carries_its_own_palette() {
        let light = loaded(TextTheme::Light).unwrap();
        let dark = loaded(TextTheme::Dark).unwrap();

        // Atom One Light is a light theme and One Dark Pro is a dark one; a
        // swapped pair would show up here.
        assert!(light.background().r > 200 && light.background().g > 200);
        assert!(dark.background().r < 80 && dark.background().g < 80);
        assert!(light.foreground().r < light.background().r);
        assert!(dark.foreground().r > dark.background().r);
    }

    #[test]
    fn scope_lookups_resolve_both_themes() {
        for kind in [TextTheme::Light, TextTheme::Dark] {
            let theme = loaded(kind).unwrap();
            let comment = theme.style_for_scopes(&["comment"]);
            let plain = theme.style_for_scopes(&["source.rust"]);
            assert_ne!(
                comment.foreground, plain.foreground,
                "comments should not be painted in the default foreground"
            );
        }
    }
}

//! The color themes a text preview can be painted with, and the scope lookups the
//! renderer uses to ask one for a color.
//!
//! The two bundled themes are TextMate themes generated from the VS Code themes of
//! the same name (see `assets/themes/NOTICE.md`) and embedded in the binary, so
//! the preview depends on no file on disk and no installation step. Each is parsed
//! once, on first use, and the parse result — including a failure — is kept for
//! the lifetime of the process.
//!
//! A `.tmTheme` file in the user's `theme` folder is read from disk on the same
//! terms and one step later: the file is read when a theme is asked for, not when
//! the folder is listed, so a folder of files costs nothing until one of them is
//! chosen. A file that is missing, unreadable, or not a theme at all is painted
//! with the default instead (see [`loaded`]), and the files are dropped and read
//! again whenever the tray menu is built, which is how a theme that was edited or
//! repaired since it was chosen is the one the next preview shows.

use crate::config::{read_within_budget, TextTheme};
use crate::theme_files;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::io::Cursor;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Mutex;
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

fn parse(source: &[u8]) -> Option<LoadedTheme> {
    let theme = ThemeSet::load_from_reader(&mut Cursor::new(source)).ok()?;

    // A theme says what its page and its text are; a file that parses and names
    // neither is a plist of some other kind wearing the extension, and guessing
    // colors for it would be worse than the default.
    if theme.settings.background.is_none() && theme.settings.foreground.is_none() {
        return None;
    }

    let theme: &'static Theme = Box::leak(Box::new(theme));

    Some(LoadedTheme {
        theme,
        highlighter: Highlighter::new(theme),
    })
}

static LIGHT: Lazy<Option<LoadedTheme>> = Lazy::new(|| parse(ATOM_ONE_LIGHT.as_bytes()));
static DARK: Lazy<Option<LoadedTheme>> = Lazy::new(|| parse(ONE_DARK_PRO.as_bytes()));

/// The user's themes by the name they were asked for, parsed once each — the ones
/// that did not parse included, which are an answer too: a folder of files that
/// are not themes would otherwise be read again on every repaint.
static CUSTOM: Lazy<Mutex<HashMap<String, Option<&'static LoadedTheme>>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// How many times the user's themes have been read, so that a parsed document can
/// say which reading it was styled against. A theme file edited under a running
/// app is a different theme, and the lines colored with the old one cannot be
/// handed out for it.
static CUSTOM_GENERATION: AtomicU64 = AtomicU64::new(0);

/// The theme for `kind`: its own file for a theme that came from the `theme`
/// folder, the bundled file for the two that ship, and, when a file cannot be
/// read or is not a theme, the default — a broken file is not allowed to take the
/// preview with it.
///
/// `None`, which the bundled files make impossible in practice, drops the preview
/// rather than painting it with guessed colors.
pub fn loaded(kind: TextTheme) -> Option<&'static LoadedTheme> {
    match kind {
        TextTheme::Light => LIGHT.as_ref(),
        TextTheme::Dark => DARK.as_ref(),
        TextTheme::Custom(name) => custom(name).or_else(|| LIGHT.as_ref()),
    }
}

/// The theme a file in the `theme` folder holds, or `None` when the folder has no
/// theme by that name, when the file cannot be read, or when what it holds is not
/// a theme. A file is read whole for the parse, so it is read under the budget every
/// other file is read under; one past it is answered with the default rather than
/// with memory it has no business asking for.
pub fn custom(name: &str) -> Option<&'static LoadedTheme> {
    let mut themes = match CUSTOM.lock() {
        Ok(themes) => themes,
        Err(poisoned) => poisoned.into_inner(),
    };
    if let Some(loaded) = themes.get(name) {
        return *loaded;
    }

    let loaded = theme_files::path_of(name)
        .and_then(|path| read_within_budget(&path))
        .and_then(|source| parse(&source))
        .map(|theme| &*Box::leak(Box::new(theme)));

    themes.insert(name.to_string(), loaded);
    loaded
}

/// Forgets the user's themes and counts a new reading, so that the folder as it
/// is now is what the next menu lists and the next preview paints. Built with the
/// tray menu, which is the one moment a user is looking at the files that back
/// these names.
pub fn refresh_custom() {
    if let Ok(mut themes) = CUSTOM.lock() {
        themes.clear();
    }
    CUSTOM_GENERATION.fetch_add(1, Ordering::Relaxed);
}

/// Which reading of the user's themes is current; see `CUSTOM_GENERATION`.
pub fn generation() -> u64 {
    CUSTOM_GENERATION.load(Ordering::Relaxed)
}

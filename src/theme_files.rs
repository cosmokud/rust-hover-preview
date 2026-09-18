//! The `theme` folder beside `config.ini`: the `.tmTheme` files a user drops into
//! it, and the names the tray menu and `config.ini` know them by.
//!
//! A file is a theme when it carries the extension — in any case, since a Windows
//! file name is case-insensitive — and it goes by its file stem, which is what the
//! menu lists and what `config.ini` writes after its `custom:` marker. Nothing
//! here reads a file: whether what one holds is a theme is `text_theme`'s to
//! decide, and one that turns out not to be is painted with the default rather
//! than kept out of the list.

use std::fs;
use std::path::{Path, PathBuf};

/// The extension a theme file carries. A theme is named by its stem; this is what
/// the stem is joined back to when the file is read.
pub const EXTENSION: &str = ".tmTheme";

/// The folder the user's themes live in.
pub fn dir() -> Option<PathBuf> {
    crate::config::AppConfig::theme_dir()
}

/// Makes sure the folder is there, so there is somewhere to drop a file into
/// without first having to guess a path the app never shows.
pub fn ensure() {
    if let Some(dir) = dir() {
        let _ = fs::create_dir_all(dir);
    }
}

/// The themes the folder holds, by name, in the order the menu lists them.
///
/// A folder that cannot be read is an empty one: there is nothing to list and
/// nothing to paint with, and the bundled themes are unaffected either way.
pub fn names() -> Vec<String> {
    let Some(dir) = dir() else {
        return Vec::new();
    };
    let Ok(entries) = fs::read_dir(dir) else {
        return Vec::new();
    };

    let mut names: Vec<String> = entries
        .flatten()
        .filter(|entry| {
            entry
                .file_type()
                .map(|kind| kind.is_file())
                .unwrap_or(false)
        })
        .filter_map(|entry| Some(theme_name(&entry.path())?.to_string()))
        .collect();

    // Folded, so that a folder holding `Solarized` and `atom` lists the way it
    // reads rather than with every capitalized name ahead of the rest.
    names.sort_by(|a, b| {
        a.to_lowercase()
            .cmp(&b.to_lowercase())
            .then_with(|| a.cmp(b))
    });
    names
}

/// The name the folder holds a theme under, for a name written by hand: matched
/// without regard to case and with the extension optional, so `Solarized`,
/// `solarized` and `solarized.tmTheme` are one theme. `None` when no file goes by
/// that name.
pub fn find(name: &str) -> Option<String> {
    let wanted = stem(name).to_lowercase();
    names()
        .into_iter()
        .find(|held| held.to_lowercase() == wanted)
}

/// Where a theme's file is, resolved through [`find`] so the path is the file's
/// own name rather than the one it was written with — and so a name that is not a
/// file name at all cannot be joined into one.
pub fn path_of(name: &str) -> Option<PathBuf> {
    let name = find(name)?;
    dir().map(|dir| dir.join(format!("{name}{EXTENSION}")))
}

/// The name a file goes by, or `None` when the file is not a theme file.
fn theme_name(path: &Path) -> Option<&str> {
    let extension = path.extension()?.to_str()?;
    if !extension.eq_ignore_ascii_case(EXTENSION.trim_start_matches('.')) {
        return None;
    }

    // A leading dot begins a *name*, not an extension, so `.tmTheme` never
    // reaches this far and `.gitignore` is not a theme.
    let name = path.file_stem()?.to_str()?;
    (!name.is_empty()).then_some(name)
}

/// A written name as a file stem: trimmed, and with the extension dropped when it
/// was written with one.
fn stem(name: &str) -> &str {
    let name = name.trim();
    let Some(start) = name.len().checked_sub(EXTENSION.len()) else {
        return name;
    };

    match name.get(start..) {
        Some(tail) if tail.eq_ignore_ascii_case(EXTENSION) => name[..start].trim(),
        _ => name,
    }
}

//! Custom previews: one external program per file extension, written as `[custom:<name>]`
//! sections of `config.ini`.
//!
//! A preset is the user's own answer to "what should preview this file", and it outranks every
//! list this app ships: a `.rme` a preset names is previewed by that preset's program whatever
//! `[archive]` or `[magick]` would have said about the name (see `routing::custom_preset_for`).
//! The section is the preset's own name — `[custom:Stress Test]` — so the file reads as the
//! window lists it, and the keys under it are the five things the window edits:
//!
//! ```ini
//! [custom:Stress Test]
//! exe=C:\Program Files\AMD\RyzenMaster\bin\Stress.exe
//! extensions=rmc, rmp
//! scaling=fit
//! background=black
//! volume=35
//! ```
//!
//! **Why the section is `custom:` and not a list's name.** Every other section is one of the
//! sixteen lists, keyed by the row that reads it (`formats::lists`); a preset is not a row and
//! is not one per kind — it is a program, and there may be a hundred of them. The prefix is what
//! keeps the two apart in a file both are written to, and it is what makes a preset recognisable
//! to a hand reading the file.
//!
//! **What a preset is not.** It is not a renderer: nothing here draws anything, and the app does
//! not decode a file a preset claims. The program is started with the file, its window is what
//! the preview shows, and a pin opens that program properly rather than pinning a picture of it
//! (see `ui::preview_window::custom_preview`). So `scaling` is how large the program's window is
//! drawn in the hover, `background` is what stands behind it where the window is smaller than the
//! band, and `volume` is the level a preview of it is played at — the same three settings the
//! tray's own rows are, read from the preset rather than from the configuration.

use std::path::Path;

use configparser::ini::Ini;

use crate::formats::lists::sanitize_extension_list;

use super::app_config::AppConfig;
use super::defaults::{
    sanitize_preview_scale_percent, sanitize_volume, DEFAULT_AUDIO_VOLUME,
    DEFAULT_PREVIEW_SCALE_PERCENT,
};
use super::setting_types::{PreviewScale, TransparentBackground, DEFAULT_IMAGE_BACKGROUND};

/// The prefix every preset's section carries: `[custom:<name>]`.
pub const SECTION_PREFIX: &str = "custom:";

/// How many presets the file may hold, which is the window's own ceiling on its list.
pub const MAX_PRESETS: usize = 100;

/// The longest name a preset may have, in characters — a tab has to fit it, and the section
/// name is written from it.
pub const MAX_NAME_CHARS: usize = 48;

/// One custom preview: the program, the names it claims, and the three settings a preview of it
/// is drawn with.
///
/// It is a plain value rather than something with a lock of its own: the configuration holds the
/// list, a read clones the one preset it found, and every reader asks `AppConfig` for it — so a
/// preview on a worker thread never holds the configuration's lock while a program starts.
#[derive(Debug, Clone, PartialEq)]
pub struct CustomPreset {
    /// The name the window lists it by, and the half of its section name after the prefix.
    pub name: String,
    /// The program, as the user typed or browsed to it. Empty is a preset with nothing to run,
    /// which is a preset that claims nothing until one is given.
    pub exe: String,
    /// The extensions it claims, each a bare extension without its dot, lowercased — the same
    /// shape `formats::lists` keeps its own entries in.
    pub extensions: Vec<String>,
    /// How large the program's window is drawn in the hover.
    pub scaling: PreviewScale,
    /// What stands behind it where the window is smaller than the band.
    pub background: TransparentBackground,
    /// The level a preview of it is played at, a percentage.
    pub volume: u32,
}

impl Default for CustomPreset {
    fn default() -> Self {
        Self {
            name: String::new(),
            exe: String::new(),
            extensions: Vec::new(),
            scaling: PreviewScale::Percent(DEFAULT_PREVIEW_SCALE_PERCENT),
            background: DEFAULT_IMAGE_BACKGROUND,
            volume: DEFAULT_AUDIO_VOLUME,
        }
    }
}

impl CustomPreset {
    /// A preset of this name and nothing else: what the window's own `Add` makes.
    pub fn named(name: &str) -> Self {
        Self {
            name: sanitize_name(name).unwrap_or_default(),
            ..Self::default()
        }
    }

    /// The section this preset is written under.
    pub fn section(&self) -> String {
        section_of(&self.name)
    }

    /// The extensions as the window's one input shows them: `jpg, png, gif`.
    pub fn extensions_text(&self) -> String {
        self.extensions.join(", ")
    }

    /// The program, where the preset has one that could be started.
    pub fn exe_path(&self) -> Option<&str> {
        let exe = self.exe.trim();
        (!exe.is_empty()).then_some(exe)
    }

    /// Whether this preset is one a preview can be made of: a program, and at least one name it
    /// claims. A half-filled preset is one the window keeps and nothing previews.
    pub fn is_usable(&self) -> bool {
        self.exe_path().is_some() && !self.extensions.is_empty() && !self.name.is_empty()
    }

    /// Whether this preset claims `path`, by the extension the file ends in.
    ///
    /// The extension alone, lowercased, because that is what the window's input takes and what
    /// the file lists are read by: a preset naming `psd` claims `art.PSD` and does not claim
    /// `art.psd.bak`. A name with no extension is never claimed.
    pub fn claims(&self, path: &Path) -> bool {
        let Some(extension) = path.extension().and_then(|extension| extension.to_str()) else {
            return false;
        };
        let extension = extension.to_ascii_lowercase();
        self.extensions.contains(&extension)
    }

    /// The extensions as the window's input would sanitize them, which is what a save puts back
    /// in the field: a comma-separated list of anything is reduced to the names it may hold.
    pub fn set_extensions_text(&mut self, text: &str) {
        self.extensions = sanitize_extension_list(text);
    }

    /// The preset reduced to what the file and the preview may hold: a name that can be a
    /// section, a program path that is one line, the extensions read through the list's own
    /// sanitiser, and the three settings inside their bounds.
    pub fn sanitized(mut self) -> Option<Self> {
        self.name = sanitize_name(&self.name)?;
        self.exe = sanitize_exe(&self.exe);
        self.extensions = sanitize_extension_list(&self.extensions.join(","));
        self.scaling = match self.scaling {
            PreviewScale::FitToScreen => PreviewScale::FitToScreen,
            PreviewScale::FitToScreenReduced(percent) => {
                PreviewScale::FitToScreenReduced(sanitize_preview_scale_percent(percent))
            }
            PreviewScale::Percent(percent) => {
                PreviewScale::Percent(sanitize_preview_scale_percent(percent))
            }
        };
        self.volume = sanitize_volume(self.volume);
        Some(self)
    }
}

/// The section name a preset's name is written under.
///
/// **Lowercased, because the file's reader lowercases every section it parses** — so a section
/// written as `[custom:Stress Test]` is read back as `custom:stress test`, and a file this app
/// wrote would not be the file it reads. The name as the user typed it is the `name` key inside
/// the section, which is what the window shows and what a save writes back.
pub fn section_of(name: &str) -> String {
    format!("{SECTION_PREFIX}{}", name.to_lowercase())
}

/// A preset name reduced to what may be a section of the file and a tab of the window: letters,
/// digits, spaces, `-`, `_` and `.`, trimmed, and no longer than [`MAX_NAME_CHARS`].
///
/// `None` for a name that holds nothing once those are gone, which is a name the window refuses
/// rather than writes as an empty section.
pub fn sanitize_name(raw: &str) -> Option<String> {
    let cleaned: String = raw
        .trim()
        .chars()
        .filter(|character| {
            character.is_alphanumeric() || matches!(character, ' ' | '-' | '_' | '.')
        })
        .take(MAX_NAME_CHARS)
        .collect();
    let cleaned = cleaned.trim().trim_matches('.').trim().to_string();
    (!cleaned.is_empty()).then_some(cleaned)
}

/// A program path reduced to what one `key=value` line may hold: trimmed, without the quotes a
/// user copies around a path with spaces in it, and without a newline a paste may bring.
fn sanitize_exe(raw: &str) -> String {
    raw.trim()
        .trim_matches('"')
        .replace(['\r', '\n'], "")
        .trim()
        .to_string()
}

/// Every preset written as its own section, in the order the window lists them.
///
/// The order is the list's own rather than the map's, which is what makes a file the app writes
/// and a file it reads back name the same preset where two claim the same extension.
pub fn write_all(config: &AppConfig, ini: &mut Ini) {
    for preset in config.custom_presets.iter().take(MAX_PRESETS) {
        let Some(preset) = preset.clone().sanitized() else {
            continue;
        };
        let section = preset.section();
        ini.set(&section, "name", Some(preset.name.clone()));
        ini.set(&section, "exe", Some(preset.exe.clone()));
        ini.set(&section, "extensions", Some(preset.extensions_text()));
        ini.set(&section, "scaling", Some(preset.scaling.as_str()));
        ini.set(
            &section,
            "background",
            Some(preset.background.as_str().to_string()),
        );
        ini.set(&section, "volume", Some(preset.volume.to_string()));
    }
}

/// Every `[custom:<name>]` section read into the configuration, replacing what it held.
///
/// A section that names nothing, that repeats a name already read, or that is the hundred and
/// first is dropped: the file is the user's to edit, and what the window can show is bounded.
/// A key that is missing or unreadable leaves the preset at where a fresh one starts rather than
/// dropping the preset — the name and the program are the two things worth keeping.
pub fn read_all(ini: &Ini, config: &mut AppConfig) {
    let mut presets: Vec<CustomPreset> = Vec::new();

    let mut sections: Vec<&String> = ini
        .get_map_ref()
        .keys()
        .filter(|section| section.starts_with(SECTION_PREFIX))
        .collect();
    sections.sort_unstable();

    for section in sections {
        if presets.len() >= MAX_PRESETS {
            break;
        }
        // The name as the user typed it, which the section cannot hold: a parsed section is
        // lowercased, so the `name` key is what preserves the case (see `section_of`).
        let named = ini
            .get(section, "name")
            .and_then(|name| sanitize_name(&name))
            .or_else(|| sanitize_name(section.trim_start_matches(SECTION_PREFIX)));
        let Some(name) = named else {
            continue;
        };
        if presets
            .iter()
            .any(|held| held.name.eq_ignore_ascii_case(&name))
        {
            continue;
        }

        let mut preset = CustomPreset::named(&name);
        if let Some(exe) = ini.get(section, "exe") {
            preset.exe = sanitize_exe(&exe);
        }
        if let Some(extensions) = ini.get(section, "extensions") {
            preset.set_extensions_text(&extensions);
        }
        if let Some(scaling) = ini.get(section, "scaling") {
            if let Some(scale) = PreviewScale::from_str(&scaling) {
                preset.scaling = scale;
            }
        }
        if let Some(background) = ini.get(section, "background") {
            if let Some(background) = TransparentBackground::from_str(&background) {
                preset.background = background;
            }
        }
        if let Some(volume) = ini.get(section, "volume") {
            if let Ok(volume) = volume.trim().parse::<u32>() {
                preset.volume = sanitize_volume(volume);
            }
        }
        presets.push(preset);
    }

    config.custom_presets = presets;
}

/// The preset that claims `path`, in the order the configuration holds them.
///
/// It is the question every side asks — the hook, the loader, the layout, the pin — and it is a
/// plain read of the configuration the caller already holds rather than a lock of its own, so a
/// hook callback may ask it (see `routing::custom_preset_for`). A preset with no program or no
/// extension claims nothing: a half-filled tab is one the window keeps and no file is routed by.
pub fn preset_for<'a>(config: &'a AppConfig, path: &Path) -> Option<&'a CustomPreset> {
    config
        .custom_presets
        .iter()
        .find(|preset| preset.is_usable() && preset.claims(path))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn config_with(presets: Vec<CustomPreset>) -> AppConfig {
        AppConfig {
            custom_presets: presets,
            ..AppConfig::default()
        }
    }

    #[test]
    fn a_name_is_reduced_to_what_a_section_and_a_tab_may_hold() {
        assert_eq!(sanitize_name("  Stress Test  "), Some("Stress Test".into()));
        assert_eq!(sanitize_name("a[b]c"), Some("abc".into()));
        assert_eq!(sanitize_name("   "), None);
        assert_eq!(sanitize_name("..."), None);
        assert_eq!(
            sanitize_name(&"x".repeat(80)).unwrap().chars().count(),
            MAX_NAME_CHARS
        );
    }

    #[test]
    fn extensions_are_read_through_the_list_sanitiser() {
        let mut preset = CustomPreset::named("Videos");
        preset.set_extensions_text(" jpg, .PNG, , gif, not an extension!, jpg ");
        assert_eq!(preset.extensions, vec!["jpg", "png", "gif"]);
        assert_eq!(preset.extensions_text(), "jpg, png, gif");
    }

    #[test]
    fn a_preset_claims_the_extension_it_names_and_nothing_else() {
        let mut preset = CustomPreset::named("Art");
        preset.set_extensions_text("psd");
        assert!(preset.claims(Path::new("C:/work/art.PSD")));
        assert!(!preset.claims(Path::new("C:/work/art.psd.bak")));
        assert!(!preset.claims(Path::new("C:/work/psd")));
    }

    #[test]
    fn a_half_filled_preset_claims_nothing() {
        let mut preset = CustomPreset::named("Empty");
        preset.set_extensions_text("psd");
        assert!(!preset.is_usable());
        preset.exe = "C:/tools/art.exe".into();
        assert!(preset.is_usable());

        let config = config_with(vec![preset]);
        assert!(preset_for(&config, Path::new("art.psd")).is_some());
        assert!(preset_for(&config, Path::new("art.png")).is_none());
    }

    #[test]
    fn the_file_writes_a_section_per_preset_and_reads_it_back() {
        let mut preset = CustomPreset::named("Stress Test");
        preset.exe = "\"C:/Program Files/AMD/Stress.exe\"".into();
        preset.set_extensions_text("rmc, rmp");
        preset.scaling = PreviewScale::FitToScreenReduced(40);
        preset.background = TransparentBackground::Black;
        preset.volume = 35;

        let config = config_with(vec![preset.clone()]);
        let mut ini = Ini::new();
        write_all(&config, &mut ini);

        assert_eq!(
            ini.get("custom:stress test", "name").as_deref(),
            Some("Stress Test")
        );
        assert_eq!(
            ini.get("custom:stress test", "exe").as_deref(),
            Some("C:/Program Files/AMD/Stress.exe")
        );
        assert_eq!(
            ini.get("custom:stress test", "extensions").as_deref(),
            Some("rmc, rmp")
        );
        assert_eq!(
            ini.get("custom:stress test", "scaling").as_deref(),
            Some("screen 40")
        );
        assert_eq!(
            ini.get("custom:stress test", "background").as_deref(),
            Some("black")
        );

        let mut read = AppConfig::default();
        read_all(&ini, &mut read);
        assert_eq!(read.custom_presets.len(), 1);
        assert_eq!(read.custom_presets[0].name, "Stress Test");
        assert_eq!(read.custom_presets[0].extensions, vec!["rmc", "rmp"]);
        assert_eq!(
            read.custom_presets[0].scaling,
            PreviewScale::FitToScreenReduced(40)
        );
        assert_eq!(
            read.custom_presets[0].background,
            TransparentBackground::Black
        );
        assert_eq!(read.custom_presets[0].volume, 35);
    }

    #[test]
    fn a_file_that_names_nothing_readable_leaves_a_fresh_preset() {
        let mut ini = Ini::new();
        ini.set("custom:Loose", "volume", Some("900".into()));
        ini.set("custom:Loose", "scaling", Some("enormous".into()));
        ini.set("custom:", "exe", Some("C:/x.exe".into()));

        let mut config = AppConfig::default();
        read_all(&ini, &mut config);

        assert_eq!(config.custom_presets.len(), 1);
        // A section with no `name` key is read by the name the parser gives it, which is the
        // lowercased one (see `section_of`).
        assert_eq!(config.custom_presets[0].name, "loose");
        assert_eq!(config.custom_presets[0].volume, 100);
        assert_eq!(
            config.custom_presets[0].scaling,
            PreviewScale::Percent(DEFAULT_PREVIEW_SCALE_PERCENT)
        );
        assert!(!config.custom_presets[0].is_usable());
    }
}

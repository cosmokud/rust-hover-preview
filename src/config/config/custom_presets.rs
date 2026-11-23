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
//! enabled=true
//! exe=C:\Program Files\AMD\RyzenMaster\bin\Stress.exe
//! extensions=rmc, rmp
//! window_mode=normal
//! scaling=fit
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
//! (see `ui::preview_window::custom_preview`). So `scaling` is how large the preview is drawn in
//! the hover — a percentage of the room the display leaves beside the item, and the fit for the
//! whole of that room, which is what the tray's own menu offers a document: `100%` is the whole of
//! that room, so a 1024x1024 picture on a 1440-tall display is a 1440x1440 preview whatever program
//! draws it. The file gives the shape of the box — a picture's own, and the program's window's where
//! the file declares none — `background` is what stands behind it where the window is smaller than
//! the band, and `volume` is the level a preview of it is played at — the same three settings the
//! tray's own rows are, read from the preset rather than from the configuration.

use std::path::Path;

use configparser::ini::Ini;

use crate::formats::lists::sanitize_extension_list;

use super::app_config::AppConfig;
use super::defaults::{
    sanitize_preview_scale_percent, sanitize_volume, CONFIG_SECTION, DEFAULT_AUDIO_VOLUME,
};
use super::setting_types::PreviewScale;

/// The prefix every preset's section carries: `[custom:<name>]`.
pub const SECTION_PREFIX: &str = "custom:";

/// How many presets the file may hold, which is the window's own ceiling on its list.
pub const MAX_PRESETS: usize = 100;

/// The longest name a preset may have, in characters — a tab has to fit it, and the section
/// name is written from it.
pub const MAX_NAME_CHARS: usize = 48;

/// The settings key holding the order the window lists the presets in: their names, top first.
///
/// Their order is what decides which preset wins an extension two of them name
/// (see `preset_for`), so it is written as a key of its own — the file's sections are a map's
/// order, alphabetical, and a file read back from it would hand `.jpg` to the wrong program.
pub(crate) const PRESET_ORDER_KEY: &str = "custom_preset_order";

/// How a custom preview's target window is captured and framed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CustomWindowMode {
    /// Leave the target window's chrome/styles intact and capture the full window.
    #[default]
    Normal,
    /// Strip standard frame chrome and capture client area only.
    Borderless,
}

impl CustomWindowMode {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Normal => "normal",
            Self::Borderless => "borderless",
        }
    }

    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "normal" => Some(Self::Normal),
            "borderless" => Some(Self::Borderless),
            _ => None,
        }
    }
}

/// How a custom preview captures and displays the target application's window.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CustomCaptureTechnique {
    /// On-monitor `WS_EX_LAYERED` + `LWA_ALPHA = 1` + DWM thumbnail (auto-falls back when unsupported).
    InvisibleLayered,
    /// On-monitor behind opaque `host` at `HWND_BOTTOM` + DWM thumbnail.
    BehindHost,
    /// Parked off-screen + DWM thumbnail.
    OffScreen,
    /// Target window placed directly at preview rect (`HWND_TOPMOST`, click-through, no host/thumbnail).
    DirectWindow,
    /// Staged at `(-10000, -10000)` until settled, then revealed on-monitor with zero flash.
    OffscreenStaged,
    /// In-process `WH_CBT` + `WH_CALLWNDPROC` hook with `DWMWA_CLOAK` + synchronous rect lock.
    #[default]
    InProcessHook,
}

impl CustomCaptureTechnique {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::InvisibleLayered => "invisible_layered",
            Self::BehindHost => "behind_host",
            Self::OffScreen => "off_screen",
            Self::DirectWindow => "direct_window",
            Self::OffscreenStaged => "offscreen_staged",
            Self::InProcessHook => "in_process_hook",
        }
    }

    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "invisible_layered" | "invisible-layered" => Some(Self::InvisibleLayered),
            "behind_host" | "behind-host" => Some(Self::BehindHost),
            "off_screen" | "off-screen" | "offscreen" => Some(Self::OffScreen),
            "direct_window" | "direct-window" => Some(Self::DirectWindow),
            "offscreen_staged" | "offscreen-staged" | "off_screen_staged" | "off-screen-staged" => {
                Some(Self::OffscreenStaged)
            }
            "in_process_hook" | "in-process-hook" | "inprocess_hook" | "inprocess-hook" => {
                Some(Self::InProcessHook)
            }
            _ => None,
        }
    }
}

/// How pinning / opening a custom preview promotes the target application.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CustomPromotionTechnique {
    /// Detach and promote the live preview window in-place (falls back to spawn when cold).
    #[default]
    DirectPromotion,
    /// End the preview and launch a fresh, untouched detached process for the file.
    OpenAsNewProcess,
}

impl CustomPromotionTechnique {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::DirectPromotion => "direct_promotion",
            Self::OpenAsNewProcess => "open_as_new_process",
        }
    }

    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "direct_promotion" | "direct-promotion" => Some(Self::DirectPromotion),
            "open_as_new_process" | "open-as-new-process" => Some(Self::OpenAsNewProcess),
            _ => None,
        }
    }
}

/// How a custom preview is sized: either following `Tray -> Scaling` for the hovered file's kind,
/// or an explicit per-preset override.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CustomScaling {
    #[default]
    FollowTray,
    Override(PreviewScale),
}

impl CustomScaling {
    pub fn as_str(self) -> String {
        match self {
            Self::FollowTray => "tray".to_string(),
            Self::Override(scale) => scale.as_str(),
        }
    }

    pub(crate) fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "tray" | "follow tray" | "follow_tray" | "follow-tray" | "default" | "auto" => {
                Some(Self::FollowTray)
            }
            _ => PreviewScale::from_str(value).map(Self::Override),
        }
    }

    pub fn sanitized(self) -> Self {
        match self {
            Self::FollowTray => Self::FollowTray,
            Self::Override(scale) => Self::Override(match scale {
                PreviewScale::FitToScreen => PreviewScale::FitToScreen,
                PreviewScale::FitToScreenReduced(percent) => {
                    PreviewScale::FitToScreenReduced(sanitize_preview_scale_percent(percent))
                }
                PreviewScale::Percent(percent) => {
                    PreviewScale::Percent(sanitize_preview_scale_percent(percent))
                }
            }),
        }
    }
}

/// One custom preview: the program, the names it claims, and the settings a preview of it
/// is drawn with.
///
/// It is a plain value rather than something with a lock of its own: the configuration holds the
/// list, a read clones the one preset it found, and every reader asks `AppConfig` for it — so a
/// preview on a worker thread never holds the configuration's lock while a program starts.
#[derive(Debug, Clone, PartialEq)]
pub struct CustomPreset {
    /// The name the window lists it by, and the half of its section name after the prefix.
    pub name: String,
    /// Whether this preset is active; disabled presets are skipped by `preset_for`.
    pub enabled: bool,
    /// The program, as the user typed or browsed to it. Empty is a preset with nothing to run,
    /// which is a preset that claims nothing until one is given.
    pub exe: String,
    /// The extensions it claims, each a bare extension without its dot, lowercased — the same
    /// shape `formats::lists` keeps its own entries in.
    pub extensions: Vec<String>,
    /// Whether the target window keeps its normal frame or is captured borderless.
    pub window_mode: CustomWindowMode,
    /// Capture/display technique used for this preset's target window.
    pub capture_technique: CustomCaptureTechnique,
    /// Promotion technique used when pinning / opening this preset's file.
    pub promotion_technique: CustomPromotionTechnique,
    /// How large the preview is drawn in the hover: following `Tray -> Scaling` by default, or an
    /// explicit per-preset share of the room.
    pub scaling: CustomScaling,
    /// Whether to override per-process audio volume via WASAPI (`false` leaves app/Windows Volume Mixer in control).
    pub volume_enabled: bool,
    /// The level a preview of it is played at, a percentage.
    pub volume: u32,
}

impl Default for CustomPreset {
    fn default() -> Self {
        Self {
            name: String::new(),
            enabled: true,
            exe: String::new(),
            extensions: Vec::new(),
            window_mode: CustomWindowMode::default(),
            capture_technique: CustomCaptureTechnique::default(),
            promotion_technique: CustomPromotionTechnique::default(),
            scaling: CustomScaling::default(),
            volume_enabled: false,
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
        self.scaling = self.scaling.sanitized();
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
    sanitize_program_input(raw)
}

/// The program field as a paste hands it over: **Explorer's `Copy as path` (Ctrl+Shift+C) wraps the
/// path in `"`**, so `"C:\Program Files\IrfanView\i_view64.exe"` is the program `C:\Program
/// Files\IrfanView\i_view64.exe`, and a multi-line copy's line break is nobody's path.
pub fn sanitize_program_input(raw: &str) -> String {
    raw.replace(['\r', '\n'], "")
        .trim()
        .trim_matches('"')
        .trim()
        .to_string()
}

fn parse_enabled(value: &str) -> Option<bool> {
    match value.trim().to_ascii_lowercase().as_str() {
        "true" | "1" | "yes" | "on" => Some(true),
        "false" | "0" | "no" | "off" => Some(false),
        _ => None,
    }
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
        ini.set(
            &section,
            "enabled",
            Some(if preset.enabled { "true" } else { "false" }.to_string()),
        );
        ini.set(&section, "exe", Some(preset.exe.clone()));
        ini.set(&section, "extensions", Some(preset.extensions_text()));
        ini.set(
            &section,
            "window_mode",
            Some(preset.window_mode.as_str().to_string()),
        );
        ini.set(
            &section,
            "capture_technique",
            Some(preset.capture_technique.as_str().to_string()),
        );
        ini.set(
            &section,
            "promotion_technique",
            Some(preset.promotion_technique.as_str().to_string()),
        );
        ini.set(&section, "scaling", Some(preset.scaling.as_str()));
        ini.set(
            &section,
            "volume_enabled",
            Some(
                if preset.volume_enabled {
                    "true"
                } else {
                    "false"
                }
                .to_string(),
            ),
        );
        ini.set(&section, "volume", Some(preset.volume.to_string()));
    }

    // The order the window lists them in, which is what the top-to-bottom precedence is
    // (`preset_for`). Read back from the settings section rather than from the order of the
    // sections in the file, which is a map's order: alphabetical, and not the list's. Written
    // even with nothing to name, so a file this build wrote is one that has the key.
    let names: Vec<String> = config
        .custom_presets
        .iter()
        .filter_map(|preset| sanitize_name(&preset.name))
        .collect();
    ini.set(CONFIG_SECTION, PRESET_ORDER_KEY, Some(names.join(", ")));
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
        if let Some(enabled) = ini.get(section, "enabled") {
            if let Some(val) = parse_enabled(&enabled) {
                preset.enabled = val;
            }
        }
        if let Some(exe) = ini.get(section, "exe") {
            preset.exe = sanitize_exe(&exe);
        }
        if let Some(extensions) = ini.get(section, "extensions") {
            preset.set_extensions_text(&extensions);
        }
        if let Some(window_mode) = ini.get(section, "window_mode") {
            if let Some(mode) = CustomWindowMode::from_str(&window_mode) {
                preset.window_mode = mode;
            }
        }
        if let Some(technique) = ini.get(section, "capture_technique") {
            if let Some(parsed) = CustomCaptureTechnique::from_str(&technique) {
                preset.capture_technique = parsed;
            }
        }
        if let Some(technique) = ini.get(section, "promotion_technique") {
            if let Some(parsed) = CustomPromotionTechnique::from_str(&technique) {
                preset.promotion_technique = parsed;
            }
        }
        if let Some(scaling) = ini.get(section, "scaling") {
            if let Some(scale) = CustomScaling::from_str(&scaling) {
                preset.scaling = scale;
            }
        }
        // Legacy `background=` keys in existing config.ini files are silently ignored.
        if let Some(ve) = ini.get(section, "volume_enabled") {
            if let Some(val) = parse_enabled(&ve) {
                preset.volume_enabled = val;
            }
        }
        if let Some(volume) = ini.get(section, "volume") {
            if let Ok(volume) = volume.trim().parse::<u32>() {
                preset.volume = sanitize_volume(volume);
            }
        }
        presets.push(preset);
    }

    // The order the window listed them in, which the sections themselves do not keep: a name the
    // order names but the file does not hold is dropped, and one it does not name — a preset
    // written by hand — keeps the place the sections gave it, at the end.
    if let Some(order) = ini.get(CONFIG_SECTION, PRESET_ORDER_KEY) {
        let names: Vec<String> = order.split(',').filter_map(sanitize_name).collect();
        presets.sort_by_key(|preset| {
            names
                .iter()
                .position(|name| name.eq_ignore_ascii_case(&preset.name))
                .unwrap_or(usize::MAX)
        });
    }

    config.custom_presets = presets;
}

/// The preset that claims `path`: **the first usable, enabled one in the order the window lists
/// them, top to bottom**, which is what lets a preset above another win the extensions both name.
///
/// It is the question every side asks — the hook, the loader, the layout, the pin — and it is a
/// plain read of the configuration the caller already holds rather than a lock of its own, so a
/// hook callback may ask it (see `routing::custom_preset_for`). A preset with no program or no
/// extension claims nothing: a half-filled tab is one the window keeps and no file is routed by.
pub fn preset_for<'a>(config: &'a AppConfig, path: &Path) -> Option<&'a CustomPreset> {
    if !config.custom_preview_enabled {
        return None;
    }
    config
        .custom_presets
        .iter()
        .find(|preset| preset.enabled && preset.is_usable() && preset.claims(path))
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
    fn the_top_preset_wins_an_extension_two_of_them_name() {
        let mut paint = CustomPreset::named("Painted");
        paint.exe = "C:/Windows/System32/mspaint.exe".into();
        paint.set_extensions_text("jpg, png");

        let mut art = CustomPreset::named("Art");
        art.exe = "C:/Art/Art.exe".into();
        art.set_extensions_text("jpg, psd");

        let config = config_with(vec![paint, art]);
        assert_eq!(
            preset_for(&config, Path::new("holiday.JPG")).map(|p| p.name.as_str()),
            Some("Painted")
        );
        assert_eq!(
            preset_for(&config, Path::new("art.psd")).map(|p| p.name.as_str()),
            Some("Art")
        );
    }

    #[test]
    fn the_window_order_is_written_and_read_back() {
        let mut first = CustomPreset::named("Painted");
        first.exe = "C:/mspaint.exe".into();
        first.set_extensions_text("jpg");
        let mut second = CustomPreset::named("Art");
        second.exe = "C:/Art.exe".into();
        second.set_extensions_text("jpg");

        // "Art" sorts above "Painted" as a section; the written order is what puts Painted top.
        let config = config_with(vec![first, second]);
        let mut ini = Ini::new();
        write_all(&config, &mut ini);

        let mut read = AppConfig::default();
        read_all(&ini, &mut read);
        assert_eq!(
            read.custom_presets
                .iter()
                .map(|preset| preset.name.as_str())
                .collect::<Vec<_>>(),
            vec!["Painted", "Art"]
        );
        assert_eq!(
            preset_for(&read, Path::new("holiday.jpg")).map(|p| p.name.as_str()),
            Some("Painted")
        );
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
    fn a_copied_as_path_program_loses_its_quotes() {
        assert_eq!(
            sanitize_program_input("\"C:\\Program Files\\IrfanView\\i_view64.exe\""),
            "C:\\Program Files\\IrfanView\\i_view64.exe"
        );
        assert_eq!(
            sanitize_program_input("  \"C:\\tools\\art.exe\"\r\n"),
            "C:\\tools\\art.exe"
        );
        assert_eq!(sanitize_program_input("  "), "");
    }

    #[test]
    fn the_file_writes_a_section_per_preset_and_reads_it_back() {
        let mut preset = CustomPreset::named("Stress Test");
        preset.enabled = false;
        preset.exe = "\"tools/Stress.exe\"".into();
        preset.set_extensions_text("rmc, rmp");
        preset.window_mode = CustomWindowMode::Borderless;
        preset.scaling = CustomScaling::Override(PreviewScale::FitToScreenReduced(40));
        preset.volume = 35;

        let config = config_with(vec![preset.clone()]);
        let mut ini = Ini::new();
        write_all(&config, &mut ini);

        assert_eq!(
            ini.get("custom:stress test", "name").as_deref(),
            Some("Stress Test")
        );
        assert_eq!(
            ini.get("custom:stress test", "enabled").as_deref(),
            Some("false")
        );
        assert_eq!(
            ini.get("custom:stress test", "exe").as_deref(),
            Some("tools/Stress.exe")
        );
        assert_eq!(
            ini.get("custom:stress test", "extensions").as_deref(),
            Some("rmc, rmp")
        );
        assert_eq!(
            ini.get("custom:stress test", "window_mode").as_deref(),
            Some("borderless")
        );
        assert_eq!(
            ini.get("custom:stress test", "scaling").as_deref(),
            Some("screen 40")
        );
        assert!(ini.get("custom:stress test", "background").is_none());

        let mut read = AppConfig::default();
        read_all(&ini, &mut read);
        assert_eq!(read.custom_presets.len(), 1);
        assert_eq!(read.custom_presets[0].name, "Stress Test");
        assert!(!read.custom_presets[0].enabled);
        assert_eq!(read.custom_presets[0].extensions, vec!["rmc", "rmp"]);
        assert_eq!(
            read.custom_presets[0].window_mode,
            CustomWindowMode::Borderless
        );
        assert_eq!(
            read.custom_presets[0].scaling,
            CustomScaling::Override(PreviewScale::FitToScreenReduced(40))
        );
        assert_eq!(read.custom_presets[0].volume, 35);
    }

    #[test]
    fn enabled_defaults_true_skips_disabled_and_ignores_legacy_background() {
        let mut ini = Ini::new();
        ini.set("custom:legacy", "name", Some("Legacy".into()));
        ini.set("custom:legacy", "exe", Some("viewer.exe".into()));
        ini.set("custom:legacy", "extensions", Some("psd".into()));
        ini.set("custom:legacy", "background", Some("black".into()));

        let mut cfg = AppConfig::default();
        read_all(&ini, &mut cfg);
        assert_eq!(cfg.custom_presets.len(), 1);
        assert!(cfg.custom_presets[0].enabled);

        let mut off = CustomPreset::named("First");
        off.enabled = false;
        off.exe = "first.exe".into();
        off.set_extensions_text("psd");

        let mut on = CustomPreset::named("Second");
        on.enabled = true;
        on.exe = "second.exe".into();
        on.set_extensions_text("psd");

        let cfg = config_with(vec![off.clone(), on]);
        assert_eq!(
            preset_for(&cfg, Path::new("art.psd")).map(|p| p.name.as_str()),
            Some("Second")
        );

        let cfg_only_off = config_with(vec![off]);
        assert!(preset_for(&cfg_only_off, Path::new("art.psd")).is_none());
    }

    #[test]
    fn volume_enabled_defaults_false_and_round_trips() {
        assert!(!CustomPreset::default().volume_enabled);
        assert!(!CustomPreset::named("Fresh").volume_enabled);

        // Missing key in existing config.ini defaults to false.
        let mut ini = Ini::new();
        ini.set("custom:legacy", "name", Some("Legacy".into()));
        ini.set("custom:legacy", "exe", Some("viewer.exe".into()));
        ini.set("custom:legacy", "extensions", Some("mp4".into()));
        ini.set("custom:legacy", "volume", Some("40".into()));
        let mut cfg = AppConfig::default();
        read_all(&ini, &mut cfg);
        assert!(!cfg.custom_presets[0].volume_enabled);
        assert_eq!(cfg.custom_presets[0].volume, 40);

        // Explicit true and false round-trip through write_all / read_all.
        let mut on = CustomPreset::named("Loud");
        on.exe = "player.exe".into();
        on.set_extensions_text("mp4");
        on.volume_enabled = true;
        on.volume = 65;

        let mut off = CustomPreset::named("Default");
        off.exe = "viewer.exe".into();
        off.set_extensions_text("mkv");
        off.volume_enabled = false;
        off.volume = 25;

        let mut out = Ini::new();
        write_all(&config_with(vec![on, off]), &mut out);
        assert_eq!(
            out.get("custom:loud", "volume_enabled").as_deref(),
            Some("true")
        );
        assert_eq!(
            out.get("custom:default", "volume_enabled").as_deref(),
            Some("false")
        );

        let mut read = AppConfig::default();
        read_all(&out, &mut read);
        assert_eq!(read.custom_presets.len(), 2);
        assert!(read.custom_presets[0].volume_enabled);
        assert_eq!(read.custom_presets[0].volume, 65);
        assert!(!read.custom_presets[1].volume_enabled);
        assert_eq!(read.custom_presets[1].volume, 25);
    }

    #[test]
    fn follow_tray_and_normal_window_mode_are_defaults_and_backward_compatible() {
        let mut preset = CustomPreset::named("Viewer");
        preset.exe = "viewer.exe".into();
        preset.set_extensions_text("png");
        assert_eq!(preset.window_mode, CustomWindowMode::Normal);
        assert_eq!(preset.scaling, CustomScaling::FollowTray);

        let config = config_with(vec![preset]);
        let mut ini = Ini::new();
        write_all(&config, &mut ini);
        assert_eq!(
            ini.get("custom:viewer", "window_mode").as_deref(),
            Some("normal")
        );
        assert_eq!(ini.get("custom:viewer", "scaling").as_deref(), Some("tray"));

        let mut read = AppConfig::default();
        read_all(&ini, &mut read);
        assert_eq!(read.custom_presets[0].window_mode, CustomWindowMode::Normal);
        assert_eq!(read.custom_presets[0].scaling, CustomScaling::FollowTray);

        // Legacy ini without `window_mode` and with `fit` / numeric `scaling` still loads cleanly.
        let mut legacy = Ini::new();
        legacy.set("custom:old", "name", Some("Old".into()));
        legacy.set("custom:old", "exe", Some("old.exe".into()));
        legacy.set("custom:old", "extensions", Some("jpg".into()));
        legacy.set("custom:old", "scaling", Some("fit".into()));
        let mut read_legacy = AppConfig::default();
        read_all(&legacy, &mut read_legacy);
        assert_eq!(
            read_legacy.custom_presets[0].window_mode,
            CustomWindowMode::Normal
        );
        assert_eq!(
            read_legacy.custom_presets[0].scaling,
            CustomScaling::Override(PreviewScale::FitToScreen)
        );
    }

    #[test]
    fn a_file_that_names_nothing_readable_leaves_a_fresh_preset() {
        let mut ini = Ini::new();
        ini.set("custom:Loose", "volume", Some("900".into()));
        ini.set("custom:Loose", "scaling", Some("enormous".into()));
        ini.set("custom:Loose", "window_mode", Some("unknown".into()));
        ini.set("custom:", "exe", Some("x.exe".into()));

        let mut config = AppConfig::default();
        read_all(&ini, &mut config);

        assert_eq!(config.custom_presets.len(), 1);
        // A section with no `name` key is read by the name the parser gives it, which is the
        // lowercased one (see `section_of`).
        assert_eq!(config.custom_presets[0].name, "loose");
        assert_eq!(config.custom_presets[0].volume, 100);
        assert_eq!(
            config.custom_presets[0].window_mode,
            CustomWindowMode::Normal
        );
        assert_eq!(config.custom_presets[0].scaling, CustomScaling::FollowTray);
        assert!(!config.custom_presets[0].is_usable());
    }

    #[test]
    fn capture_technique_round_trips_and_defaults_to_in_process_hook() {
        assert_eq!(
            CustomCaptureTechnique::default(),
            CustomCaptureTechnique::InProcessHook
        );
        assert_eq!(
            CustomPreset::named("Fresh").capture_technique,
            CustomCaptureTechnique::InProcessHook
        );
        let cases = [
            (CustomCaptureTechnique::InProcessHook, "in_process_hook"),
            (
                CustomCaptureTechnique::InvisibleLayered,
                "invisible_layered",
            ),
            (CustomCaptureTechnique::BehindHost, "behind_host"),
            (CustomCaptureTechnique::OffScreen, "off_screen"),
            (CustomCaptureTechnique::DirectWindow, "direct_window"),
            (CustomCaptureTechnique::OffscreenStaged, "offscreen_staged"),
        ];
        for (tech, expected_str) in cases {
            let mut preset = CustomPreset::named("TechTest");
            preset.exe = "viewer.exe".into();
            preset.set_extensions_text("png");
            preset.capture_technique = tech;

            let cfg = config_with(vec![preset]);
            let mut ini = Ini::new();
            write_all(&cfg, &mut ini);
            assert_eq!(
                ini.get("custom:techtest", "capture_technique").as_deref(),
                Some(expected_str)
            );

            let mut read = AppConfig::default();
            read_all(&ini, &mut read);
            assert_eq!(read.custom_presets[0].capture_technique, tech);
        }

        // Missing or unknown value falls back to InProcessHook.
        let mut legacy = Ini::new();
        legacy.set("custom:missing", "name", Some("Missing".into()));
        legacy.set("custom:missing", "exe", Some("a.exe".into()));
        legacy.set("custom:missing", "extensions", Some("png".into()));
        legacy.set("custom:unknown", "name", Some("Unknown".into()));
        legacy.set("custom:unknown", "exe", Some("b.exe".into()));
        legacy.set("custom:unknown", "extensions", Some("png".into()));
        legacy.set("custom:unknown", "capture_technique", Some("bogus".into()));

        let mut read = AppConfig::default();
        read_all(&legacy, &mut read);
        assert_eq!(read.custom_presets.len(), 2);
        for p in &read.custom_presets {
            assert_eq!(p.capture_technique, CustomCaptureTechnique::InProcessHook);
        }
    }

    #[test]
    fn promotion_technique_round_trips_and_defaults_to_direct_promotion() {
        assert_eq!(
            CustomPromotionTechnique::default(),
            CustomPromotionTechnique::DirectPromotion
        );
        assert_eq!(
            CustomPreset::named("Fresh").promotion_technique,
            CustomPromotionTechnique::DirectPromotion
        );
        let cases = [
            (
                CustomPromotionTechnique::DirectPromotion,
                "direct_promotion",
            ),
            (
                CustomPromotionTechnique::OpenAsNewProcess,
                "open_as_new_process",
            ),
        ];
        for (tech, expected_str) in cases {
            let mut preset = CustomPreset::named("PromoTest");
            preset.exe = "viewer.exe".into();
            preset.set_extensions_text("png");
            preset.promotion_technique = tech;

            let cfg = config_with(vec![preset]);
            let mut ini = Ini::new();
            write_all(&cfg, &mut ini);
            assert_eq!(
                ini.get("custom:promotest", "promotion_technique")
                    .as_deref(),
                Some(expected_str)
            );

            let mut read = AppConfig::default();
            read_all(&ini, &mut read);
            assert_eq!(read.custom_presets[0].promotion_technique, tech);
        }

        // Missing or unknown value falls back to DirectPromotion.
        let mut legacy = Ini::new();
        legacy.set("custom:missing", "name", Some("Missing".into()));
        legacy.set("custom:missing", "exe", Some("a.exe".into()));
        legacy.set("custom:missing", "extensions", Some("png".into()));
        legacy.set("custom:unknown", "name", Some("Unknown".into()));
        legacy.set("custom:unknown", "exe", Some("b.exe".into()));
        legacy.set("custom:unknown", "extensions", Some("png".into()));
        legacy.set(
            "custom:unknown",
            "promotion_technique",
            Some("bogus".into()),
        );

        let mut read = AppConfig::default();
        read_all(&legacy, &mut read);
        assert_eq!(read.custom_presets.len(), 2);
        for p in &read.custom_presets {
            assert_eq!(
                p.promotion_technique,
                CustomPromotionTechnique::DirectPromotion
            );
        }
    }

    #[test]
    fn master_toggle_off_bypasses_all_presets_and_routing() {
        let preset = CustomPreset {
            name: "Viewer".into(),
            exe: "viewer.exe".into(),
            extensions: vec!["png".into()],
            ..CustomPreset::default()
        };
        let mut cfg = config_with(vec![preset]);
        let path = Path::new("photo.png");

        assert!(preset_for(&cfg, path).is_some());
        assert!(crate::formats::routing::is_custom_preview(path, &cfg));

        cfg.custom_preview_enabled = false;
        assert!(preset_for(&cfg, path).is_none());
        assert!(!crate::formats::routing::is_custom_preview(path, &cfg));
        assert!(crate::formats::routing::custom_preset_for(path, &cfg).is_none());
    }
}

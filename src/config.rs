use configparser::ini::Ini;
use directories::BaseDirs;
use std::fs;
use std::path::PathBuf;

use crate::text_formats::{
    sanitize_extensions, sanitize_names, DEFAULT_TEXT_EXTENSIONS, DEFAULT_TEXT_NAMES,
};

const CONFIG_SECTION: &str = "settings";
/// The text-preview extension list lives in its own section so the one long
/// value stays easy to find and edit by hand.
const TEXT_SECTION: &str = "text";
pub const DEFAULT_WEBP_PLAYBACK_FPS: u32 = 90;
pub const MAX_WEBP_PLAYBACK_FPS: u32 = 90;
pub const DEFAULT_PREVIEW_SCALE_PERCENT: u32 = 100;
pub const MIN_PREVIEW_SCALE_PERCENT: u32 = 1;
pub const MAX_PREVIEW_SCALE_PERCENT: u32 = 1000;
pub const DEFAULT_TEXT_FONT_SCALE_PERCENT: u32 = 125;
pub const MIN_TEXT_FONT_SCALE_PERCENT: u32 = 1;
pub const MAX_TEXT_FONT_SCALE_PERCENT: u32 = 1000;
pub const DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS: f32 = 40.0;
pub const MAX_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS: f32 = 1000.0;

pub fn sanitize_webp_playback_fps(value: u32) -> u32 {
    match value {
        0 => DEFAULT_WEBP_PLAYBACK_FPS,
        1..=MAX_WEBP_PLAYBACK_FPS => value,
        _ => MAX_WEBP_PLAYBACK_FPS,
    }
}

/// The text preview font scale, where `0` and nonsense land back on the default.
/// Anything from 1% to 1000% is honored: the tray offers a handful of steps, but
/// the value is a percentage either way, so a hand-edited one is not rounded to
/// the nearest menu entry.
pub fn sanitize_text_font_scale_percent(value: u32) -> u32 {
    if value == 0 {
        DEFAULT_TEXT_FONT_SCALE_PERCENT
    } else {
        value.clamp(MIN_TEXT_FONT_SCALE_PERCENT, MAX_TEXT_FONT_SCALE_PERCENT)
    }
}

/// How far past the far edge of a text preview the pointer region reaches, in
/// logical pixels at the display's DPI.
///
/// Zero is a distance like any other — the region then ends at the preview, which
/// is what it did before the grace existed — so only a value that is not a number
/// at all falls back to the default.
pub fn sanitize_text_scroll_far_edge_grace_pixels(value: f32) -> f32 {
    if value.is_finite() {
        value.clamp(0.0, MAX_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS)
    } else {
        DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS
    }
}

fn parse_text_font_scale(value: &str) -> Option<u32> {
    let normalized = value.trim().to_ascii_lowercase();
    let normalized = normalized.trim_end_matches('%').trim();
    normalized
        .parse::<u32>()
        .ok()
        .map(sanitize_text_font_scale_percent)
}

fn sanitize_preview_scale_percent(value: u32) -> u32 {
    if value == 0 {
        DEFAULT_PREVIEW_SCALE_PERCENT
    } else {
        value.clamp(MIN_PREVIEW_SCALE_PERCENT, MAX_PREVIEW_SCALE_PERCENT)
    }
}

/// How the preview is sized relative to the media's native pixel dimensions.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PreviewScale {
    /// Scale as large as the available display area allows.
    FitToScreen,
    /// Scale by a percentage of the media's native size.
    Percent(u32),
}

impl PreviewScale {
    pub fn as_str(self) -> String {
        match self {
            Self::FitToScreen => "fit".to_string(),
            Self::Percent(percent) => sanitize_preview_scale_percent(percent).to_string(),
        }
    }

    /// Requested scale relative to the native size, or `None` when the preview
    /// should use the largest scale the display area allows.
    pub fn target_scale(self) -> Option<f32> {
        match self {
            Self::FitToScreen => None,
            Self::Percent(percent) => Some(sanitize_preview_scale_percent(percent) as f32 / 100.0),
        }
    }

    fn from_str(value: &str) -> Option<Self> {
        let normalized = value.trim().to_ascii_lowercase();
        let normalized = normalized.trim_end_matches('%').trim();

        match normalized {
            "fit" | "fit to screen" | "fit-to-screen" | "fit_to_screen" | "fittoscreen" => {
                Some(Self::FitToScreen)
            }
            _ => normalized
                .parse::<u32>()
                .ok()
                .map(|percent| Self::Percent(sanitize_preview_scale_percent(percent))),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TransparentBackground {
    Transparent,
    Black,
    White,
    Checkerboard,
}

impl TransparentBackground {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Transparent => "transparent",
            Self::Black => "black",
            Self::White => "white",
            Self::Checkerboard => "checkerboard",
        }
    }

    fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "transparent" => Some(Self::Transparent),
            "black" => Some(Self::Black),
            "white" => Some(Self::White),
            "checkerboard" => Some(Self::Checkerboard),
            _ => None,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TextTheme {
    /// Atom One Light.
    Light,
    /// One Dark Pro.
    Dark,
}

impl TextTheme {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Light => "light",
            Self::Dark => "dark",
        }
    }

    fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "light" | "atom one light" | "atom-one-light" => Some(Self::Light),
            "dark" | "one dark pro" | "one-dark-pro" => Some(Self::Dark),
            _ => None,
        }
    }
}

/// How a Markdown file is laid out: the document it describes, or the markup
/// itself with the Markdown syntax highlighted like any other source file.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MarkdownMode {
    Rendered,
    Source,
}

impl MarkdownMode {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Rendered => "rendered",
            Self::Source => "source",
        }
    }

    fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "rendered" | "render" | "document" => Some(Self::Rendered),
            "source" | "raw" | "highlighted" | "highlighted source" => Some(Self::Source),
            _ => None,
        }
    }
}

/// What the trigger key does while it is held.
///
/// A preview appears when a file is hovered, so the trigger key is normally what
/// stops that: hold it and nothing previews while it is down. The reverse suits
/// a machine where previews are the exception rather than the rule — hold the key
/// and they appear, let go and they stop — and it is the same gesture either way,
/// which is why one key covers both.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TriggerKeyMode {
    Disable,
    Enable,
}

impl TriggerKeyMode {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Disable => "disable",
            Self::Enable => "enable",
        }
    }

    /// Whether a preview may appear while the trigger key is in this state.
    ///
    /// One question, whichever way round the setting is: in `Disable` the key is
    /// what stops previews, so it has to be up; in `Enable` it is the only thing
    /// that starts them, so it has to be down.
    pub fn allows_previews(self, trigger_key_down: bool) -> bool {
        match self {
            Self::Disable => !trigger_key_down,
            Self::Enable => trigger_key_down,
        }
    }

    fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "disable" | "disables" | "off" | "hold to disable" => Some(Self::Disable),
            "enable" | "enables" | "on" | "hold to enable" => Some(Self::Enable),
            _ => None,
        }
    }
}

#[derive(Debug, Clone)]
pub struct AppConfig {
    pub is_first_run: bool,
    pub run_at_startup: bool,
    pub hover_delay_ms: u64,
    pub preview_enabled: bool,
    /// The modifier the trigger key watches, by name.
    pub trigger_key: String,
    /// What holding it does: stop previews, or allow them.
    pub trigger_key_mode: TriggerKeyMode,
    pub confirm_file_type: bool,
    pub follow_cursor: bool,
    pub same_file_rehover_delay_ms: u64,
    pub webp_playback_fps: u32,
    pub transparent_background: TransparentBackground,
    pub video_volume: u32,
    pub preview_scale: PreviewScale,
    pub theme: TextTheme,
    pub markdown_mode: MarkdownMode,
    /// Whether text files are previewed at all, ahead of the extension list.
    pub text_preview_enabled: bool,
    /// Whether a text preview is more than something to look at: a preview that
    /// scrolls, that can be selected and copied from, and that a pointer can rest
    /// on without closing it. Off by default, because it changes what a preview
    /// does rather than what it shows.
    pub text_preview_full_mode: bool,
    /// Font scale for text previews, as a percentage of the default size.
    pub text_font_scale_percent: u32,
    /// How far past the far edge of a text preview the pointer region reaches, in
    /// logical pixels at the display's DPI, so a hand that overshoots the edge on
    /// its way to the scrollbar does not take the preview down with it.
    pub text_scroll_far_edge_grace_pixels: f32,
    /// Extensions previewed as text, already normalized for lookup.
    pub text_extensions: Vec<String>,
    /// File names previewed as text — the ones with no extension to match, like
    /// `LICENSE` and `Makefile` — already normalized for lookup.
    pub text_names: Vec<String>,
}

impl Default for AppConfig {
    fn default() -> Self {
        Self {
            is_first_run: false,
            run_at_startup: true,
            hover_delay_ms: 0,
            preview_enabled: true,
            trigger_key: "alt".to_string(),
            trigger_key_mode: TriggerKeyMode::Disable,
            confirm_file_type: false,
            follow_cursor: false,
            same_file_rehover_delay_ms: 750,
            webp_playback_fps: DEFAULT_WEBP_PLAYBACK_FPS,
            transparent_background: TransparentBackground::Black,
            video_volume: 0, // Mute by default
            preview_scale: PreviewScale::Percent(DEFAULT_PREVIEW_SCALE_PERCENT),
            theme: TextTheme::Light,
            markdown_mode: MarkdownMode::Rendered,
            text_preview_enabled: true,
            text_preview_full_mode: false,
            text_font_scale_percent: DEFAULT_TEXT_FONT_SCALE_PERCENT,
            text_scroll_far_edge_grace_pixels: DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS,
            text_extensions: sanitize_extensions(DEFAULT_TEXT_EXTENSIONS),
            text_names: sanitize_names(DEFAULT_TEXT_NAMES),
        }
    }
}

impl AppConfig {
    pub fn config_path() -> Option<PathBuf> {
        BaseDirs::new().map(|dirs| {
            dirs.config_dir()
                .join("rust-hover-preview")
                .join("config.ini")
        })
    }

    pub fn load() -> Self {
        let mut config = Self::default();

        if let Some(path) = Self::config_path() {
            config.is_first_run = !path.exists();
            let mut ini = Ini::new();
            if ini.load(path.to_string_lossy().as_ref()).is_ok() {
                config.apply_ini(&ini);
            }
        }

        // Always save to ensure new fields are written to config file
        config.save();
        config
    }

    pub fn reload_from_disk(&mut self) {
        if let Some(path) = Self::config_path() {
            let mut ini = Ini::new();
            if ini.load(path.to_string_lossy().as_ref()).is_ok() {
                self.apply_ini(&ini);
            }
        }
    }

    pub fn save(&self) {
        if let Some(path) = Self::config_path() {
            if let Some(parent) = path.parent() {
                let _ = fs::create_dir_all(parent);
            }
            let mut ini = Ini::new();
            ini.set(
                CONFIG_SECTION,
                "run_at_startup",
                Some(self.run_at_startup.to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "hover_delay_ms",
                Some(self.hover_delay_ms.to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "preview_enabled",
                Some(self.preview_enabled.to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "trigger_key",
                Some(self.trigger_key.clone()),
            );
            ini.set(
                CONFIG_SECTION,
                "trigger_key_mode",
                Some(self.trigger_key_mode.as_str().to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "confirm_file_type",
                Some(self.confirm_file_type.to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "follow_cursor",
                Some(self.follow_cursor.to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "same_file_rehover_delay_ms",
                Some(self.same_file_rehover_delay_ms.to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "webp_playback_fps",
                Some(sanitize_webp_playback_fps(self.webp_playback_fps).to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "transparent_background",
                Some(self.transparent_background.as_str().to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "video_volume",
                Some(self.video_volume.to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "preview_scale",
                Some(self.preview_scale.as_str()),
            );
            ini.set(
                CONFIG_SECTION,
                "theme",
                Some(self.theme.as_str().to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "markdown_mode",
                Some(self.markdown_mode.as_str().to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "text_preview_enabled",
                Some(self.text_preview_enabled.to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "text_preview_full_mode",
                Some(self.text_preview_full_mode.to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "text_font_scale",
                Some(sanitize_text_font_scale_percent(self.text_font_scale_percent).to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "text_scroll_far_edge_grace_pixels",
                Some(
                    sanitize_text_scroll_far_edge_grace_pixels(
                        self.text_scroll_far_edge_grace_pixels,
                    )
                    .to_string(),
                ),
            );
            ini.set(
                TEXT_SECTION,
                "extensions",
                Some(sanitize_extensions(&self.text_extensions.join(",")).join(",")),
            );
            ini.set(
                TEXT_SECTION,
                "names",
                Some(sanitize_names(&self.text_names.join(",")).join(",")),
            );
            let _ = ini.write(path.to_string_lossy().as_ref());
        }
    }

    fn apply_ini(&mut self, ini: &Ini) {
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "run_at_startup") {
            self.run_at_startup = value;
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "hover_delay_ms") {
            self.hover_delay_ms = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "preview_enabled") {
            self.preview_enabled = value;
        }
        // `off_trigger_key` is what the key was called when the trigger could only
        // stop previews; a file written then still names the key the same way.
        if let Some(value) = ini
            .get(CONFIG_SECTION, "trigger_key")
            .or_else(|| ini.get(CONFIG_SECTION, "off_trigger_key"))
        {
            let value = value.trim();
            if !value.is_empty() {
                self.trigger_key = value.to_string();
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "trigger_key_mode") {
            if let Some(mode) = TriggerKeyMode::from_str(&value) {
                self.trigger_key_mode = mode;
            }
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "confirm_file_type") {
            self.confirm_file_type = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "follow_cursor") {
            self.follow_cursor = value;
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "same_file_rehover_delay_ms") {
            self.same_file_rehover_delay_ms = value;
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "webp_playback_fps") {
            if let Ok(value) = u32::try_from(value) {
                self.webp_playback_fps = sanitize_webp_playback_fps(value);
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "transparent_background") {
            if let Some(background) = TransparentBackground::from_str(&value) {
                self.transparent_background = background;
            }
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "video_volume") {
            if let Ok(value) = u32::try_from(value) {
                self.video_volume = value;
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "preview_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.preview_scale = scale;
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "theme") {
            if let Some(theme) = TextTheme::from_str(&value) {
                self.theme = theme;
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "markdown_mode") {
            if let Some(mode) = MarkdownMode::from_str(&value) {
                self.markdown_mode = mode;
            }
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "text_preview_enabled") {
            self.text_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "text_preview_full_mode") {
            self.text_preview_full_mode = value;
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "text_font_scale") {
            if let Some(scale) = parse_text_font_scale(&value) {
                self.text_font_scale_percent = scale;
            }
        }
        if let Ok(Some(value)) = ini.getfloat(CONFIG_SECTION, "text_scroll_far_edge_grace_pixels") {
            self.text_scroll_far_edge_grace_pixels =
                sanitize_text_scroll_far_edge_grace_pixels(value as f32);
        }
        // An empty list means the user removed every extension, and the built-in
        // list comes back only when the key itself is gone.
        if let Some(value) = ini.get(TEXT_SECTION, "extensions") {
            self.text_extensions = sanitize_extensions(&value);
        }
        if let Some(value) = ini.get(TEXT_SECTION, "names") {
            self.text_names = sanitize_names(&value);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{
        parse_text_font_scale, sanitize_text_font_scale_percent,
        sanitize_text_scroll_far_edge_grace_pixels, AppConfig, MarkdownMode, TextTheme,
        TriggerKeyMode,
    };
    use configparser::ini::Ini;

    /// Reading the file is the half of the round trip a test can do without
    /// touching the real `config.ini` in the user's profile.
    #[test]
    fn the_text_keys_are_read_from_the_ini() {
        let mut ini = Ini::new();
        ini.set("settings", "text_preview_enabled", Some("false".into()));
        ini.set("settings", "text_preview_full_mode", Some("true".into()));
        ini.set("settings", "text_font_scale", Some("175%".into()));
        ini.set(
            "settings",
            "text_scroll_far_edge_grace_pixels",
            Some("12.5".into()),
        );
        ini.set("text", "extensions", Some("md, py, .RS, md".into()));
        ini.set(
            "text",
            "names",
            Some(".gitignore, LICENSE , makefile, LICENSE".into()),
        );

        let mut config = AppConfig::default();
        config.apply_ini(&ini);

        assert!(!config.text_preview_enabled);
        assert!(config.text_preview_full_mode);
        assert_eq!(config.text_font_scale_percent, 175);
        assert_eq!(config.text_scroll_far_edge_grace_pixels, 12.5);
        assert_eq!(config.text_extensions, vec!["md", "py", "rs"]);
        assert_eq!(config.text_names, vec!["gitignore", "license", "makefile"]);
    }

    /// Full mode changes what a preview does rather than what it shows, so it waits
    /// to be asked for: a configuration that does not mention it leaves it off.
    #[test]
    fn full_mode_is_off_unless_it_is_turned_on() {
        assert!(!AppConfig::default().text_preview_full_mode);

        let mut ini = Ini::new();
        ini.set("settings", "text_preview_enabled", Some("true".into()));

        let mut config = AppConfig::default();
        config.apply_ini(&ini);
        assert!(!config.text_preview_full_mode);
    }

    /// The trigger key is Alt until another is named, and what it does is to stop
    /// previews until the configuration says otherwise.
    #[test]
    fn the_trigger_key_stops_previews_unless_it_is_told_to_allow_them() {
        let default = AppConfig::default();
        assert_eq!(default.trigger_key, "alt");
        assert_eq!(default.trigger_key_mode, TriggerKeyMode::Disable);

        let mut ini = Ini::new();
        ini.set("settings", "trigger_key", Some("ctrl".into()));
        ini.set("settings", "trigger_key_mode", Some("Enable".into()));

        let mut config = AppConfig::default();
        config.apply_ini(&ini);
        assert_eq!(config.trigger_key, "ctrl");
        assert_eq!(config.trigger_key_mode, TriggerKeyMode::Enable);

        // The name the key had when the trigger could only stop previews still
        // names it, so a file written before the setting changed keeps its key.
        let mut older = Ini::new();
        older.set("settings", "off_trigger_key", Some("shift".into()));

        let mut config = AppConfig::default();
        config.apply_ini(&older);
        assert_eq!(config.trigger_key, "shift");
        assert_eq!(config.trigger_key_mode, TriggerKeyMode::Disable);
    }

    /// The two trigger modes are the same gesture read opposite ways, so each is
    /// the exact inverse of the other.
    #[test]
    fn the_trigger_modes_are_inverses_of_each_other() {
        assert!(TriggerKeyMode::Disable.allows_previews(false));
        assert!(!TriggerKeyMode::Disable.allows_previews(true));

        assert!(!TriggerKeyMode::Enable.allows_previews(false));
        assert!(TriggerKeyMode::Enable.allows_previews(true));

        for key_down in [false, true] {
            assert_ne!(
                TriggerKeyMode::Disable.allows_previews(key_down),
                TriggerKeyMode::Enable.allows_previews(key_down),
                "one mode stops what the other starts"
            );
        }
    }

    #[test]
    fn the_text_font_scale_accepts_percentages_and_resets_on_zero() {
        assert_eq!(sanitize_text_font_scale_percent(0), 125);
        assert_eq!(sanitize_text_font_scale_percent(125), 125);
        assert_eq!(sanitize_text_font_scale_percent(400), 400);
        assert_eq!(sanitize_text_font_scale_percent(5000), 1000);

        assert_eq!(parse_text_font_scale("150"), Some(150));
        assert_eq!(parse_text_font_scale(" 175% "), Some(175));
        assert_eq!(parse_text_font_scale("0"), Some(125));
        assert_eq!(parse_text_font_scale("large"), None);
    }

    /// The grace is a distance, and zero is one of them: it ends the region at the
    /// preview, which is what a user asking for no grace is asking for. Only a
    /// value that is not a distance at all resets to the default.
    #[test]
    fn the_far_edge_grace_is_read_as_a_distance() {
        assert_eq!(AppConfig::default().text_scroll_far_edge_grace_pixels, 40.0);

        assert_eq!(sanitize_text_scroll_far_edge_grace_pixels(40.0), 40.0);
        assert_eq!(sanitize_text_scroll_far_edge_grace_pixels(12.5), 12.5);
        assert_eq!(sanitize_text_scroll_far_edge_grace_pixels(0.0), 0.0);
        assert_eq!(sanitize_text_scroll_far_edge_grace_pixels(-5.0), 0.0);
        assert_eq!(sanitize_text_scroll_far_edge_grace_pixels(5000.0), 1000.0);
        assert_eq!(sanitize_text_scroll_far_edge_grace_pixels(f32::NAN), 40.0);
        assert_eq!(
            sanitize_text_scroll_far_edge_grace_pixels(f32::INFINITY),
            40.0
        );
    }

    #[test]
    fn theme_names_round_trip_and_accept_aliases() {
        assert_eq!(TextTheme::from_str("light"), Some(TextTheme::Light));
        assert_eq!(
            TextTheme::from_str(" Atom One Light "),
            Some(TextTheme::Light)
        );
        assert_eq!(TextTheme::from_str("DARK"), Some(TextTheme::Dark));
        assert_eq!(TextTheme::from_str("one-dark-pro"), Some(TextTheme::Dark));
        assert_eq!(TextTheme::from_str("solarized"), None);
        assert_eq!(TextTheme::Light.as_str(), "light");
        assert_eq!(TextTheme::Dark.as_str(), "dark");
    }

    #[test]
    fn markdown_modes_round_trip_and_accept_aliases() {
        assert_eq!(
            MarkdownMode::from_str("rendered"),
            Some(MarkdownMode::Rendered)
        );
        assert_eq!(MarkdownMode::from_str("Raw"), Some(MarkdownMode::Source));
        assert_eq!(MarkdownMode::from_str("elsewhere"), None);
        assert_eq!(MarkdownMode::Rendered.as_str(), "rendered");
        assert_eq!(MarkdownMode::Source.as_str(), "source");
    }
}

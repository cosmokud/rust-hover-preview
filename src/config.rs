use configparser::ini::Ini;
use directories::BaseDirs;
use once_cell::sync::Lazy;
use std::collections::HashSet;
use std::fs;
use std::path::PathBuf;
use std::sync::Mutex;

use crate::text_formats::{
    sanitize_extensions, sanitize_names, DEFAULT_TEXT_EXTENSIONS, DEFAULT_TEXT_NAMES,
};
use crate::theme_files;

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

/// The marker a theme from the `theme` folder is written after in `config.ini`.
///
/// It is what keeps a file named after a bundled theme apart from the bundled
/// theme of that name: `light.tmTheme` in the folder is `custom:light`, while
/// `light` on its own is Atom One Light however many files the folder holds.
const CUSTOM_THEME_PREFIX: &str = "custom:";

/// Names already interned for [`TextTheme::custom`], so one file name is one
/// `&'static str` however many times a menu or a configuration read hands it over.
static CUSTOM_THEME_NAMES: Lazy<Mutex<HashSet<&'static str>>> =
    Lazy::new(|| Mutex::new(HashSet::new()));

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TextTheme {
    /// Atom One Light.
    Light,
    /// One Dark Pro.
    Dark,
    /// A `.tmTheme` file in the `theme` folder, by the name it is listed under.
    Custom(&'static str),
}

impl TextTheme {
    /// How the theme is written in `config.ini`.
    pub fn as_str(self) -> String {
        match self {
            Self::Light => "light".to_string(),
            Self::Dark => "dark".to_string(),
            Self::Custom(name) => format!("{CUSTOM_THEME_PREFIX}{name}"),
        }
    }

    /// The theme a file of that name in the `theme` folder is, interned so that it
    /// can live in a value the rest of the app copies around.
    ///
    /// The folder is listed and `config.ini` is read back whenever either changes,
    /// and both hand a name over again each time; interning is what keeps the same
    /// theme one value rather than a fresh string per reading.
    pub fn custom(name: &str) -> Self {
        let mut names = match CUSTOM_THEME_NAMES.lock() {
            Ok(names) => names,
            Err(poisoned) => poisoned.into_inner(),
        };
        if let Some(name) = names.get(name).copied() {
            return Self::Custom(name);
        }

        let name: &'static str = Box::leak(name.to_string().into_boxed_str());
        names.insert(name);
        Self::Custom(name)
    }

    fn from_str(value: &str) -> Option<Self> {
        match value.trim().to_ascii_lowercase().as_str() {
            "light" | "atom one light" | "atom-one-light" => Some(Self::Light),
            "dark" | "one dark pro" | "one-dark-pro" => Some(Self::Dark),
            _ => None,
        }
    }

    /// The theme a `config.ini` value names, if this machine has one.
    ///
    /// Three spellings, in the order they are read: a name the tray wrote for a
    /// file (`custom:atom-one-light`), the name of a bundled theme — `light`,
    /// `dark`, and the long spellings accepted as ever — and a file in the `theme`
    /// folder written by hand without the marker, its extension optional. A value
    /// that is none of those leaves the theme as it was, which is what puts the
    /// default back when a selection stops making sense.
    fn resolve(value: &str) -> Option<Self> {
        let value = value.trim();

        if let Some(name) = value.strip_prefix(CUSTOM_THEME_PREFIX) {
            let name = name.trim();
            if name.is_empty() {
                return None;
            }

            // A file the folder holds under another spelling is still that file,
            // and the menu lists the spelling the folder uses; a name it does not
            // hold is kept as written, so the choice survives a file that is
            // missing for the moment.
            let name = theme_files::find(name).unwrap_or_else(|| name.to_string());
            return Some(Self::custom(&name));
        }

        if let Some(built_in) = Self::from_str(value) {
            return Some(built_in);
        }

        theme_files::find(value).map(|name| Self::custom(&name))
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

/// A kind of preview, as the tray's `Toggle Preview Types` submenu lists them.
///
/// A gate is not a file list: it says whether previews of that kind may be shown
/// at all, and the lists that decide *which* files of that kind are previewed are
/// left untouched by it, so switching a kind off and back on restores exactly
/// what was configured.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PreviewType {
    Images,
    Videos,
    Text,
    Pdf,
}

impl PreviewType {
    /// Whether this kind of preview may be shown under the current configuration.
    pub fn enabled(self) -> bool {
        crate::CONFIG
            .lock()
            .map(|config| self.enabled_in(&config))
            .unwrap_or(true)
    }

    /// Whether this kind of preview is switched on in `config`.
    pub fn enabled_in(self, config: &AppConfig) -> bool {
        match self {
            Self::Images => config.image_preview_enabled,
            Self::Videos => config.video_preview_enabled,
            Self::Text => config.text_preview_enabled,
            Self::Pdf => config.pdf_preview_enabled,
        }
    }

    /// Switch this kind of preview on or off.
    pub fn set_enabled_in(self, config: &mut AppConfig, enabled: bool) {
        match self {
            Self::Images => config.image_preview_enabled = enabled,
            Self::Videos => config.video_preview_enabled = enabled,
            Self::Text => config.text_preview_enabled = enabled,
            Self::Pdf => config.pdf_preview_enabled = enabled,
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
    /// Whether image previews may be shown at all.
    pub image_preview_enabled: bool,
    /// Whether video previews may be shown at all.
    pub video_preview_enabled: bool,
    /// Whether text files are previewed at all, ahead of the extension list.
    pub text_preview_enabled: bool,
    /// Whether PDF previews may be shown at all.
    pub pdf_preview_enabled: bool,
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
            image_preview_enabled: true,
            video_preview_enabled: true,
            text_preview_enabled: true,
            pdf_preview_enabled: true,
            text_preview_full_mode: false,
            text_font_scale_percent: DEFAULT_TEXT_FONT_SCALE_PERCENT,
            text_scroll_far_edge_grace_pixels: DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS,
            text_extensions: sanitize_extensions(DEFAULT_TEXT_EXTENSIONS),
            text_names: sanitize_names(DEFAULT_TEXT_NAMES),
        }
    }
}

impl AppConfig {
    /// The app's own folder under the roaming profile, holding `config.ini` and
    /// the `theme` folder beside it.
    fn folder() -> Option<PathBuf> {
        BaseDirs::new().map(|dirs| dirs.config_dir().join("rust-hover-preview"))
    }

    pub fn config_path() -> Option<PathBuf> {
        Self::folder().map(|folder| folder.join("config.ini"))
    }

    /// The folder the user's `.tmTheme` files live in, beside `config.ini`.
    pub fn theme_dir() -> Option<PathBuf> {
        Self::folder().map(|folder| folder.join("theme"))
    }

    pub fn load() -> Self {
        let mut config = Self::default();

        // The folder exists before anything can list it, so that adding a theme is
        // dropping a file into a path the app can name rather than creating one.
        theme_files::ensure();

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
            ini.set(CONFIG_SECTION, "theme", Some(self.theme.as_str()));
            ini.set(
                CONFIG_SECTION,
                "markdown_mode",
                Some(self.markdown_mode.as_str().to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "image_preview_enabled",
                Some(self.image_preview_enabled.to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "video_preview_enabled",
                Some(self.video_preview_enabled.to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "text_preview_enabled",
                Some(self.text_preview_enabled.to_string()),
            );
            ini.set(
                CONFIG_SECTION,
                "pdf_preview_enabled",
                Some(self.pdf_preview_enabled.to_string()),
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
            if let Some(theme) = TextTheme::resolve(&value) {
                self.theme = theme;
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "markdown_mode") {
            if let Some(mode) = MarkdownMode::from_str(&value) {
                self.markdown_mode = mode;
            }
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "image_preview_enabled") {
            self.image_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "video_preview_enabled") {
            self.video_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "text_preview_enabled") {
            self.text_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pdf_preview_enabled") {
            self.pdf_preview_enabled = value;
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

//! The tray icon, and the menu it opens.
//!
//! This is the whole of what a user gets of this app by right-clicking one picture in Explorer:
//! an icon in the notification area, a menu of every setting this build has, and a handler per
//! row of it. It is also where the main thread spends the run (see `run_tray`), so the window
//! behind the icon is the app's own message loop rather than something it shares with a thread.
//!
//! It is split into five parts by what each one is asked for:
//!
//! - `ids` — the command id every row carries, and the table each menu lists. A click is
//!   answered by comparing against these, so a range that ran into another submenu's is a click
//!   read as the wrong setting.
//! - `submenus` — one builder per submenu, the words its items are written as, and the lookup
//!   that turns an id back into the choice it was listed for.
//! - `menus` — the popup itself, assembled from those builders in the order a user sees it.
//! - `commands` — what a click on a row does: one function per command, each writing one setting
//!   and telling whatever else reads it.
//! - `event_loop` — the window procedure behind the icon, and the loop that runs it.
//!
//! The order a change goes through is the order of that list: a menu is built from a table of
//! choices and a range of ids, a click names one of those ids, a handler turns it into a value
//! the setting holds, and the window procedure is what reads the click in the first place.

mod commands;
mod event_loop;
mod ids;
mod menus;
mod submenus;

pub use event_loop::run_tray;

// What the tests read through this module: the ids and tables each menu is built from, and the
// names the submenus and commands answer a click with.
#[cfg(test)]
use crate::config::config::{
    AudioSeek, AvoidMode, EngineIdle, PreviewScale, TransparentBackground, VideoEngine,
    DEFAULT_ANIMATED_SCALE, DEFAULT_AUDIO_SCALE, DEFAULT_AUDIO_SEEK, DEFAULT_AUDIO_VOLUME,
    DEFAULT_AVOID_MODE, DEFAULT_DDS_BACKGROUND, DEFAULT_DESIGN_BACKGROUND, DEFAULT_DOCUMENT_SCALE,
    DEFAULT_EBOOK_SCALE, DEFAULT_FONT_BACKGROUND, DEFAULT_FONT_SCALE, DEFAULT_HOVER_DELAY_MS,
    DEFAULT_HTML_BACKGROUND, DEFAULT_IMAGE_BACKGROUND, DEFAULT_OFFICE_ENGINE_IDLE_SECS,
    DEFAULT_PIN_MODE_AUDIO_SEEK, DEFAULT_PIN_NAV_FILE_TYPES, DEFAULT_PREVIEW_SCALE,
    DEFAULT_RENDER_HTML,
    DEFAULT_SAME_FILE_REHOVER_DELAY_MS, DEFAULT_SETTLING_DELAY_MS, DEFAULT_TRIGGER_KEY_AFFECT_PIN_MODE,
    DEFAULT_VECTOR_BACKGROUND, DEFAULT_VECTOR_SCALE, DEFAULT_VIDEO_ENGINE, DEFAULT_VIDEO_ENGINE_FALLBACK,
    DEFAULT_VIDEO_SCALE, DEFAULT_VIDEO_VOLUME, VOLUME_CHOICES,
};
#[cfg(test)]
use commands::{audio_seek_at, avoid_mode_at, bitmap_scale_at, timing_delay_at};
#[cfg(test)]
use ids::{
    AFK_TIMER_CHOICES_SECS, AUDIO_SCALE_CHOICES, AUDIO_SEEK_CHOICES, AVOID_CHOICES,
    BACKGROUND_CHOICES, BITMAP_SCALE_CHOICES, CACHE_SIZE_CHOICES_MB, CODEC_COMMANDS,
    DDS_BACKGROUND_CHOICES, DOCUMENT_SCALE_CHOICES, ENGINE_IDLE_CHOICES, FONT_SIZE_CHOICES,
    HTML_BACKGROUND_CHOICES, ID_TRAY_AFK_TIMER_BASE, ID_TRAY_ANIMATED_SCALE_BASE,
    ID_TRAY_AUDIO_SCALE_BASE, ID_TRAY_AUDIO_SEEK_BASE, ID_TRAY_AUDIO_VOLUME_BASE,
    ID_TRAY_AVOID_BASE, ID_TRAY_CHECK_UPDATES, ID_TRAY_CODEC_BASE,
    ID_TRAY_DDS_BACKGROUND_BASE, ID_TRAY_DELAY_BASE,
    ID_TRAY_DESIGN_BACKGROUND_BASE, ID_TRAY_DESIGN_SCALE_BASE, ID_TRAY_DOCUMENT_CACHE_BASE,
    ID_TRAY_DOCUMENT_SCALE_BASE, ID_TRAY_EBOOK_SCALE_BASE, ID_TRAY_ENABLE,
    ID_TRAY_ENGINE_IDLE_BASE, ID_TRAY_ENGINE_OFFICE_LIBRE, ID_TRAY_ENGINE_OFFICE_MS,
    ID_TRAY_ENGINE_PERSISTENT_BASE, ID_TRAY_EXIT, ID_TRAY_FONT_100, ID_TRAY_FONT_110,
    ID_TRAY_FONT_BACKGROUND_BASE, ID_TRAY_FONT_SCALE_BASE, ID_TRAY_GENERAL_DISK_CACHE_BASE,
    ID_TRAY_HTML_BACKGROUND_BASE, ID_TRAY_IMAGE_BACKGROUND_BASE, ID_TRAY_IMAGE_CACHE_BASE,
    ID_TRAY_IMAGE_DISK_CACHE_BASE, ID_TRAY_LIBREOFFICE_IDLE_BASE, ID_TRAY_MARKDOWN_RENDERED,
    ID_TRAY_MARKDOWN_SOURCE, ID_TRAY_NORMALIZE_VIDEO_VOLUME, ID_TRAY_NORMALIZE_VOLUME,
    ID_TRAY_OPEN_CONFIG, ID_TRAY_PIN, ID_TRAY_PIN_MODE_AUDIO_SEEK_BASE,
    ID_TRAY_PIN_NAV_ALL, ID_TRAY_PIN_NAV_CATEGORY,
    ID_TRAY_PIN_PAUSE_AUDIO, ID_TRAY_PIN_PAUSE_VIDEO, ID_TRAY_PIN_UPDATE, ID_TRAY_PIN_UPDATE_HOVER,
    ID_TRAY_POSITION_BEST, ID_TRAY_POSITION_FOLLOW, ID_TRAY_PRIORITIZE_KEYBOARD,
    ID_TRAY_REHOVER_DELAY_BASE, ID_TRAY_REMEMBER_VIDEO_VOLUME, ID_TRAY_REMEMBER_VOLUME,
    ID_TRAY_RENDER_HTML, ID_TRAY_RESET_LISTS, ID_TRAY_RESET_SETTINGS, ID_TRAY_SCALE_BASE,
    ID_TRAY_SETTLING_DELAY_BASE, ID_TRAY_STARTUP, ID_TRAY_TEXT_SCALE_BASE,
    ID_TRAY_THEME_CUSTOM_BASE, ID_TRAY_THEME_DARK, ID_TRAY_THEME_LIGHT, ID_TRAY_TICK_BASE,
    ID_TRAY_TRIGGER_AFFECT_PIN, ID_TRAY_TRIGGER_DISABLE, ID_TRAY_TRIGGER_ENABLE,
    ID_TRAY_TRIGGER_ENABLED, ID_TRAY_TYPE_ARCHIVES, ID_TRAY_TYPE_AUDIO, ID_TRAY_TYPE_DESIGN,
    ID_TRAY_TYPE_DOCUMENT, ID_TRAY_TYPE_EBOOK, ID_TRAY_TYPE_FONTS, ID_TRAY_TYPE_IMAGES,
    ID_TRAY_TYPE_TEXT, ID_TRAY_TYPE_VECTOR, ID_TRAY_TYPE_VIDEOS, ID_TRAY_UPDATE,
    ID_TRAY_VECTOR_BACKGROUND_BASE, ID_TRAY_VECTOR_SCALE_BASE, ID_TRAY_VIDEO_ENGINE_BASE,
    ID_TRAY_VIDEO_ENGINE_FALLBACK, ID_TRAY_VIDEO_HW_ACCEL, ID_TRAY_VIDEO_SCALE_BASE,
    ID_TRAY_VIDEO_VOLUME_BASE, ID_TRAY_WEBVIEW_IDLE_BASE, TICK_CHOICES_MS,
    TIMING_DELAY_CHOICES_MS, VIDEO_ENGINE_CHOICES,
};
#[cfg(test)]
use submenus::{
    audio_scale_at, audio_seek_label, avoid_label, background_at, background_label,
    bitmap_scale_label, dds_background_at, document_scale_at, document_scale_label,
    engine_idle_at, engine_idle_label, html_background_at, pin_mode_audio_seek_label,
    pin_nav_label, pin_update_label,
    remember_volume_label, system_menu_label, update_available_label, video_engine_label,
};

#[cfg(test)]
mod tests;

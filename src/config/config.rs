//! Everything the app reads its settings out of, reached from every reader and engine as
//! `config::config`.
//!
//! The path is the one this app has always named this module by — renaming it would touch a
//! hundred import sites to answer a lint about the name itself — so it stays, and what changed
//! is what is behind it. The five modules below are the pieces the configuration is made of,
//! and this one is the whole of it as far as anything outside can tell: every name a reader,
//! an engine or a tray row uses is re-exported here, and none of them can tell which file it
//! came from.
//!
//! - [`defaults`] — what every setting starts at, and the reduction a hand-edited file is read
//!   through: a bound, a default, a clamp.
//! - [`setting_types`] — the shapes a setting's value can take, and the word each is written as.
//! - [`app_config`] — the configuration itself: [`AppConfig`], its fields, and the values a
//!   fresh one starts at.
//! - [`settings_io`] — the file: where it is, how it is read, and how it is written.
//! - [`ini_mapping`] — the file's shape, field by field: what is written, whether the file says
//!   what the app is using, and the read that puts one into the other.

mod app_config;
mod defaults;
mod ini_mapping;
mod setting_types;
mod settings_io;

pub use app_config::AppConfig;
// Every name below is part of the path the app has always reached this module by, whether
// a reader, an engine or a tray row uses it today or not. A `pub use` in a binary crate's
// private module is a re-export the lint would call unused the moment nothing happens to
// call it, which says more about today's callers than about the module's shape.
#[allow(unused_imports)]
pub use defaults::{
    decode_budget_bytes, frame_bytes_within_budget, image_decode_limits, read_within_budget,
    sanitize_decode_budget_gb, sanitize_document_cache_mb, sanitize_general_disk_cache_mb,
    sanitize_hdr_exposure, sanitize_image_cache_mb, sanitize_image_disk_cache_mb,
    sanitize_spinner_delay_ms, sanitize_text_font_scale_percent,
    sanitize_text_scroll_far_edge_grace_pixels, sanitize_tick_ms, sanitize_ttc_face,
    sanitize_volume, sanitize_webp_playback_fps, DEFAULT_AFK_TIMER_SECS,
    DEFAULT_ANIMATED_SCALE_PERCENT, DEFAULT_AUDIO_SCALE_PERCENT, DEFAULT_AUDIO_VOLUME, DEFAULT_DECODE_BUDGET_GB,
    DEFAULT_DOCUMENT_CACHE_MB, DEFAULT_FONT_SCALE_PERCENT, DEFAULT_GENERAL_DISK_CACHE_MB,
    DEFAULT_HDR_EXPOSURE, DEFAULT_HDR_TONE_MAP, DEFAULT_HOVER_DELAY_MS, DEFAULT_IMAGE_CACHE_MB,
    DEFAULT_IMAGE_DISK_CACHE_MB, DEFAULT_LIBREOFFICE_IDLE_SECS, DEFAULT_NORMALIZE_VIDEO_VOLUME,
    DEFAULT_NORMALIZE_VOLUME, DEFAULT_OFFICE_ENGINE, DEFAULT_OFFICE_ENGINE_IDLE_SECS,
    DEFAULT_PREVIEW_SCALE_PERCENT, DEFAULT_REMEMBER_AUDIO_VOLUME, DEFAULT_REMEMBER_VIDEO_VOLUME,
    DEFAULT_RENDER_HTML, DEFAULT_SAME_FILE_REHOVER_DELAY_MS, DEFAULT_SETTLING_DELAY_MS,
    DEFAULT_SPINNER_DELAY_MS, DEFAULT_TEXT_FONT_SCALE_PERCENT,
    DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS, DEFAULT_TICK_MS, DEFAULT_TTC_FACE,
    DEFAULT_VIDEO_ENGINE, DEFAULT_VIDEO_ENGINE_FALLBACK, DEFAULT_VIDEO_HW_ACCEL,
    DEFAULT_VIDEO_SCALE_PERCENT, DEFAULT_VIDEO_VOLUME, DEFAULT_WEBP_PLAYBACK_FPS,
    DEFAULT_WEBVIEW_IDLE_SECS, MAX_AFK_TIMER_SECS, MAX_DECODE_BUDGET_GB, MAX_DOCUMENT_CACHE_MB,
    MAX_GENERAL_DISK_CACHE_MB, MAX_HDR_EXPOSURE, MAX_IMAGE_CACHE_MB, MAX_IMAGE_DISK_CACHE_MB,
    MAX_OFFICE_ENGINE_IDLE_SECS, MAX_PREVIEW_SCALE_PERCENT, MAX_SPINNER_DELAY_MS,
    MAX_TEXT_FONT_SCALE_PERCENT, MAX_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS, MAX_TICK_MS, MAX_TTC_FACE,
    MAX_VOLUME, MAX_WEBP_PLAYBACK_FPS, MIN_DECODE_BUDGET_GB, MIN_HDR_EXPOSURE,
    MIN_PREVIEW_SCALE_PERCENT, MIN_TEXT_FONT_SCALE_PERCENT, MIN_TICK_MS, VIDEO_FFMPEG_ABOVE_PIXELS,
    VOLUME_CHOICES,
};
#[allow(unused_imports)]
pub use setting_types::{
    sanitize_dds_background, sanitize_html_background, AudioSeek, AvoidMode, EngineIdle,
    MarkdownMode, OfficeEngine, PinNavFileTypes, PreviewScale, PreviewType, TextTheme,
    TransparentBackground, TriggerKeyMode, VideoEngine, DEFAULT_ANIMATED_SCALE,
    DEFAULT_AUDIO_SCALE, DEFAULT_AUDIO_SEEK, DEFAULT_AVOID_MODE, DEFAULT_DDS_BACKGROUND,
    DEFAULT_DESIGN_BACKGROUND, DEFAULT_DESIGN_SCALE, DEFAULT_DOCUMENT_SCALE, DEFAULT_EBOOK_SCALE,
    DEFAULT_FOLLOW_CURSOR, DEFAULT_FONT_BACKGROUND, DEFAULT_FONT_SCALE, DEFAULT_HTML_BACKGROUND,
    DEFAULT_IMAGE_BACKGROUND, DEFAULT_PIN_NAV_FILE_TYPES, DEFAULT_PIN_PAUSE_AUDIO,
    DEFAULT_PIN_PAUSE_VIDEO, DEFAULT_PIN_UPDATE_ENABLED, DEFAULT_PIN_UPDATE_ON_HOVER,
    DEFAULT_PREVIEW_SCALE, DEFAULT_TEXT_SCALE, DEFAULT_TRIGGER_KEY_AFFECT_PIN_MODE,
    DEFAULT_VECTOR_BACKGROUND, DEFAULT_VECTOR_SCALE, DEFAULT_VIDEO_SCALE,
};

// The names below are this module's own rather than the app's: the section name the file is
// read under, the shape it is written in, and the handful of things the tests reach through
// `use super::*` rather than naming out of it.
#[cfg(test)]
use crate::formats::lists;
#[cfg(test)]
use crate::readers::tone_map::Curve;
#[cfg(test)]
use configparser::ini::Ini;
#[cfg(test)]
use defaults::CONFIG_SECTION;
#[cfg(test)]
use settings_io::{headings_are_old, ordered_text, take_reset_markers, SETTING_GROUPS};
#[cfg(test)]
use std::fs;
#[cfg(test)]
use std::time::Duration;

#[cfg(test)]
mod tests;

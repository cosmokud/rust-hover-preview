//! What a click on a row of the tray does.
//!
//! One function per command the tray hands out, each of which writes one setting and then
//! tells whatever else reads it. Most of them rebuild nothing: a setting whose truth lives
//! somewhere other than the configuration file is told about the change directly, and one
//! that is read per hover or per frame needs nothing said at all - the next hover answers
//! for itself. The ones that do rebuild are the backdrops, where what stands behind a preview
//! is part of the frame it was composited into, and `toggle_preview_type`, which takes the
//! engine behind a kind of preview with it.
//!
//! The commands are split by which of the two shapes a row has, and this file is the seam
//! between them rather than a list of its own:
//!
//! - `setters` — the rows that name a value, which turn the position an item was listed at
//!   into the choice it stands for and write that. The four lookups that answer the question
//!   the other way live beside them.
//! - `toggles` — the rows that throw a switch, which read the setting and put it the other
//!   way, from the machine itself where the machine is what keeps the truth.
//!
//! Both are re-exported here under the names they had before the split, so the window proc
//! in `event_loop`, the submenu that opens a page, and the tests in `tray::tests` all reach
//! the same paths they always did. The ids that name these are in `tray::ids`.

mod setters;
mod toggles;

pub(super) use setters::{
    open_config_file, open_link, reset_lists_from_tray, reset_settings_from_tray, set_afk_timer,
    set_animated_scale, set_audio_scale, set_audio_seek, set_audio_volume, set_avoid_mode,
    set_dds_background, set_decode_budget_gb, set_design_background, set_design_scale,
    set_document_cache_mb, set_document_scale, set_ebook_scale, set_follow_cursor,
    set_font_background, set_font_scale, set_general_disk_cache_mb, set_hover_delay,
    set_html_background, set_image_background, set_image_cache_mb, set_image_disk_cache_mb,
    set_libreoffice_idle, set_markdown_mode, set_office_engine, set_office_engine_idle,
    set_pin_nav_file_types, set_pin_mode_audio_seek, set_preview_scale, set_same_file_rehover_delay, set_settling_delay,
    set_text_font_scale, set_text_scale, set_theme, set_theme_from_menu, set_tick_ms,
    set_trigger_key_mode, set_vector_background, set_vector_scale, set_video_engine,
    set_video_scale, set_video_volume, set_webview_idle,
};
pub(super) use toggles::{
    toggle_engine_persistent, toggle_normalize_video_volume, toggle_normalize_volume,
    toggle_pin_enabled, toggle_pin_pause_audio, toggle_pin_pause_video, toggle_pin_update_enabled,
    toggle_pin_update_on_hover, toggle_preview_enabled, toggle_preview_type,
    toggle_prioritize_keyboard, toggle_remember_audio_volume, toggle_remember_video_volume,
    toggle_render_html, toggle_startup, toggle_trigger_key_affect_pin_mode,
    toggle_trigger_key_enabled, toggle_video_engine_fallback, toggle_video_hw_accel,
    toggle_video_subtitles,
};

// The four that turn an id back into the choice it was listed for. No window proc asks for
// one of these — a submenu that lists a range of ids knows what its own range means — so the
// only reader is `tray::tests`, which checks a listed table against what this answers with.
#[cfg(test)]
pub(super) use setters::{audio_seek_at, avoid_mode_at, bitmap_scale_at, timing_delay_at};

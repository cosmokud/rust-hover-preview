//! The file's own shape: what `to_ini` writes, whether what a file holds is what the app
//! is using (`differs`), and the read that puts a file into a configuration (`apply_ini`).
//! Every field is named in all three, so a setting added to [`AppConfig`] is a setting the
//! file has to learn in one place rather than three.

use configparser::ini::Ini;

use crate::formats::lists;
use crate::readers::tone_map::Curve;

use super::app_config::AppConfig;
use super::defaults::{
    parse_text_font_scale, sanitize_decode_budget_gb, sanitize_document_cache_mb,
    sanitize_general_disk_cache_mb, sanitize_hdr_exposure, sanitize_image_cache_mb,
    sanitize_image_disk_cache_mb, sanitize_mpv_path, sanitize_pin_zoom_bounce_ms,
    sanitize_pin_zoom_bounce_power, sanitize_pin_zoom_bounce_speed_ms, sanitize_spinner_delay_ms,
    sanitize_text_font_scale_percent,
    sanitize_text_scroll_far_edge_grace_pixels, sanitize_tick_ms, sanitize_ttc_face,
    sanitize_volume, sanitize_webp_playback_fps, CONFIG_SECTION,
};
use super::setting_types::{
    sanitize_afk_timer_secs, sanitize_dds_background, sanitize_html_background,
    sanitize_zoom_background, AudioSeek,
    AvoidMode, BeatHz, EngineIdle, MarkdownMode, OfficeEngine, PinEndOfFile, PinMinimizeTo,
    PinNavFileTypes,
    PreviewScale, TextTheme, TransparentBackground, TriggerKeyMode, VideoEngine,
};

impl AppConfig {
    /// The settings as the file holds them: every key this build writes, with the value it
    /// writes that key as.
    ///
    /// This is what `save` writes to disk, and it is also what a file that has just been read is
    /// held up against, to tell whether it says what the app is using (see `differs`) — which is
    /// why it is a method of its own rather than the body of `save`.
    pub(super) fn to_ini(&self) -> Ini {
        let mut ini = Ini::new();
        ini.set(
            CONFIG_SECTION,
            "run_at_startup",
            Some(self.run_at_startup.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "check_for_updates",
            Some(self.check_for_updates.to_string()),
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
            "trigger_key_enabled",
            Some(self.trigger_key_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "trigger_key_affect_pin_mode",
            Some(self.trigger_key_affect_pin_mode.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_enabled",
            Some(self.pin_enabled.to_string()),
        );
        ini.set(CONFIG_SECTION, "pin_key", Some(self.pin_key.clone()));
        ini.set(
            CONFIG_SECTION,
            "pin_minimize_to",
            Some(self.pin_minimize_to.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_double_click_center",
            Some(self.pin_double_click_center.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_double_click_reset_size",
            Some(self.pin_double_click_reset_size.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_pause_video",
            Some(self.pin_pause_video.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_pause_audio",
            Some(self.pin_pause_audio.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_update_enabled",
            Some(self.pin_update_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_update_on_hover",
            Some(self.pin_update_on_hover.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_nav_video",
            Some(self.pin_nav_video.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_nav_audio",
            Some(self.pin_nav_audio.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_nav_others",
            Some(self.pin_nav_others.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "follow_cursor",
            Some(self.follow_cursor.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "avoid_mode",
            Some(self.avoid_mode.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "same_file_rehover_delay_ms",
            Some(self.same_file_rehover_delay_ms.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "settling_delay_ms",
            Some(self.settling_delay_ms.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "prioritize_keyboard",
            Some(self.prioritize_keyboard.to_string()),
        );
        ini.set(CONFIG_SECTION, "tick_ms", Some(self.tick_ms.to_string()));
        ini.set(
            CONFIG_SECTION,
            "spinner_delay_ms",
            Some(sanitize_spinner_delay_ms(self.spinner_delay_ms).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_zoom_bounce_ms",
            Some(sanitize_pin_zoom_bounce_ms(self.pin_zoom_bounce_ms).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_zoom_bounce_gate",
            Some(self.pin_zoom_bounce_gate.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_zoom_bounce_speed_ms",
            Some(sanitize_pin_zoom_bounce_speed_ms(self.pin_zoom_bounce_speed_ms).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_zoom_bounce_power",
            Some(sanitize_pin_zoom_bounce_power(self.pin_zoom_bounce_power).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "webp_playback_fps",
            Some(sanitize_webp_playback_fps(self.webp_playback_fps).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "image_cache_mb",
            Some(sanitize_image_cache_mb(self.image_cache_mb).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "image_background",
            Some(self.image_background.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "zoom_background",
            Some(self.zoom_background.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "font_background",
            Some(self.font_background.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "dds_background",
            Some(self.dds_background.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "design_background",
            Some(self.design_background.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "html_background",
            Some(self.html_background.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "vector_background",
            Some(self.vector_background.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "video_volume",
            Some(sanitize_volume(self.video_volume).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "audio_volume",
            Some(sanitize_volume(self.audio_volume).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "audio_seek",
            Some(self.audio_seek.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "beat_hz",
            Some(self.beat_hz.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_mode_audio_seek",
            Some(self.pin_mode_audio_seek.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_mode_audio_end_of_file",
            Some(self.pin_mode_audio_end_of_file.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_mode_audio_shuffle",
            Some(self.pin_mode_audio_shuffle.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_mode_video_end_of_file",
            Some(self.pin_mode_video_end_of_file.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_mode_video_shuffle",
            Some(self.pin_mode_video_shuffle.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "pin_mode_image_shuffle",
            Some(self.pin_mode_image_shuffle.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "video_loop",
            Some(self.video_loop.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "audio_loop",
            Some(self.audio_loop.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "normalize_volume",
            Some(self.normalize_volume.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "normalize_video_volume",
            Some(self.normalize_video_volume.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "remember_audio_volume",
            Some(self.remember_audio_volume.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "remember_video_volume",
            Some(self.remember_video_volume.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "preview_scale",
            Some(self.preview_scale.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "video_scale",
            Some(self.video_scale.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "audio_scale",
            Some(self.audio_scale.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "animated_scale",
            Some(self.animated_scale.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "vector_scale",
            Some(self.vector_scale.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "ebook_scale",
            Some(self.ebook_scale.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "document_scale",
            Some(self.document_scale.as_str()),
        );
        ini.set(CONFIG_SECTION, "font_scale", Some(self.font_scale.as_str()));
        ini.set(CONFIG_SECTION, "text_scale", Some(self.text_scale.as_str()));
        ini.set(
            CONFIG_SECTION,
            "design_scale",
            Some(self.design_scale.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "ttc_face",
            Some(sanitize_ttc_face(self.ttc_face).to_string()),
        );
        ini.set(CONFIG_SECTION, "theme", Some(self.theme.as_str()));
        ini.set(
            CONFIG_SECTION,
            "markdown_mode",
            Some(self.markdown_mode.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "render_html",
            Some(self.render_html.to_string()),
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
            "audio_preview_enabled",
            Some(self.audio_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "text_preview_enabled",
            Some(self.text_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "ebook_preview_enabled",
            Some(self.ebook_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "archive_preview_enabled",
            Some(self.archive_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "document_preview_enabled",
            Some(self.document_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "font_preview_enabled",
            Some(self.font_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "design_preview_enabled",
            Some(self.design_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "vector_preview_enabled",
            Some(self.vector_preview_enabled.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "document_cache_mb",
            Some(sanitize_document_cache_mb(self.document_cache_mb).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "image_disk_cache_mb",
            Some(sanitize_image_disk_cache_mb(self.image_disk_cache_mb).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "general_disk_cache_mb",
            Some(sanitize_general_disk_cache_mb(self.general_disk_cache_mb).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "office_engine",
            Some(self.office_engine.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "video_engine",
            Some(self.video_engine.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "video_engine_fallback",
            Some(self.video_engine_fallback.to_string()),
        );
        // The one setting that holds a path: written as it stands, trimmed on the way back in.
        ini.set(
            CONFIG_SECTION,
            "mpv_path",
            Some(self.mpv_path.clone()),
        );
        ini.set(
            CONFIG_SECTION,
            "video_subtitles",
            Some(self.video_subtitles.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "office_engine_idle",
            Some(self.office_engine_idle.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "webview_idle",
            Some(self.webview_idle.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "libreoffice_idle",
            Some(self.libreoffice_idle.as_str()),
        );
        ini.set(
            CONFIG_SECTION,
            "afk_timer_seconds",
            Some(sanitize_afk_timer_secs(self.afk_timer_seconds).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "office_engine_persistent",
            Some(self.office_engine_persistent.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "webview_persistent",
            Some(self.webview_persistent.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "libreoffice_persistent",
            Some(self.libreoffice_persistent.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "decode_budget_gb",
            Some(sanitize_decode_budget_gb(self.decode_budget_gb).to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "video_hw_accel",
            Some(self.video_hw_accel.to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "hdr_tone_map",
            Some(self.hdr_tone_map.as_str().to_string()),
        );
        ini.set(
            CONFIG_SECTION,
            "hdr_exposure",
            Some(sanitize_hdr_exposure(self.hdr_exposure).to_string()),
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
                sanitize_text_scroll_far_edge_grace_pixels(self.text_scroll_far_edge_grace_pixels)
                    .to_string(),
            ),
        );
        // Every list the configuration holds is one row of the table to write (see
        // `formats::lists`).
        lists::write_all(self, &mut ini);
        ini
    }

    /// Whether the file holds something other than what this app would write for the settings
    /// it has just read, which is what makes it a file to write back. The two are held up against
    /// each other both ways round: every key the app writes has to be in the file, holding what
    /// the app writes for it, and every key in the file has to be one the app writes.
    ///
    /// A key the file does not have is a difference, and so is one it holds another value for. A
    /// value the app could not read at all is that second kind, and so is one a setting reduced to
    /// what it allows — a delay past its ceiling, a face of a collection past the last one the
    /// menu offers, a tone map that is not one of them. A key the app does not write is a
    /// difference too: the file is what this app writes and nothing else, so a key of the user's
    /// own, or one left behind by an older or a newer build, is dropped by the write it asks for.
    ///
    /// What is not compared is what is not a key: a comment, a blank line, the order the keys are
    /// written in — so a comment survives until a write happens for another reason, which is when
    /// `save` writes the file out from the settings it knows and the comment goes with it.
    pub(super) fn differs(&self, ini: &Ini) -> bool {
        let wanted = self.to_ini();

        for (section, keys) in wanted.get_map_ref() {
            for (key, value) in keys {
                if ini.get(section, key) != *value {
                    return true;
                }
            }
        }

        for (section, keys) in ini.get_map_ref() {
            for key in keys.keys() {
                if wanted.get(section, key).is_none() {
                    return true;
                }
            }
        }

        false
    }

    /// Read the file into the configuration: what the file says, under the names of now, for
    /// every setting it has. A setting it does not have keeps the value it already holds, which
    /// is the default for a configuration that has just been made — and whether the file needs
    /// writing afterwards is not this read's answer to give (see `differs`).
    pub(super) fn apply_ini(&mut self, ini: &Ini) {
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "run_at_startup") {
            self.run_at_startup = value;
        }
        // Whether the app asks GitHub for a newer release on its own. A file written
        // before the setting existed has no key for it, so a fresh installation checks
        // and a file that says nothing about it is left where it starts.
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "check_for_updates") {
            self.check_for_updates = value;
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "hover_delay_ms") {
            self.hover_delay_ms = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "preview_enabled") {
            self.preview_enabled = value;
        }
        // The key the trigger watches, by name. A file that names it says what it is; a name an
        // older build wrote it under is not a name this app reads at all, and the line goes when
        // the file is written again (see `differs`).
        if let Some(value) = ini.get(CONFIG_SECTION, "trigger_key") {
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
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "trigger_key_enabled") {
            self.trigger_key_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "trigger_key_affect_pin_mode") {
            self.trigger_key_affect_pin_mode = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pin_enabled") {
            self.pin_enabled = value;
        }
        // The key that pins a preview, by name, read the way the trigger key above is: a name
        // that is empty is nothing to bind. A name no key is spelled like is left to the watcher
        // rather than turned away here, which is what lets a build that learns a new spelling
        // read an older file that already named it (see `key_input::key_to_vk`).
        if let Some(value) = ini.get(CONFIG_SECTION, "pin_key") {
            let value = value.trim();
            if !value.is_empty() {
                self.pin_key = value.to_string();
            }
        }
        // Where a pin put away by its own minimize goes — the round bubble on the desktop, or
        // the app's own button in the taskbar. A value that names neither is not one of them:
        // the file is left at the answer a fresh one has, which is the bubble.
        if let Some(value) = ini.get(CONFIG_SECTION, "pin_minimize_to") {
            if let Some(target) = PinMinimizeTo::from_str(&value) {
                self.pin_minimize_to = target;
            }
        }
        // Whether a double-click on a pin puts its window in the middle of the screen it is on.
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pin_double_click_center") {
            self.pin_double_click_center = value;
        }
        // And whether it puts the window back to the size its kind's scaling setting asks for as
        // well, rather than keeping the size a hand gave it.
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pin_double_click_reset_size") {
            self.pin_double_click_reset_size = value;
        }
        // Whether a pin collapsed into its bubble holds what is playing where it is, asked of
        // a video and a sound apart — the two switches of the kind's own `Pause on Minimize`.
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pin_pause_video") {
            self.pin_pause_video = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pin_pause_audio") {
            self.pin_pause_audio = value;
        }
        // Whether a pin follows what the user picks while it is up, and whether the pointer's
        // own hover is one of the ways it does.
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pin_update_enabled") {
            self.pin_update_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pin_update_on_hover") {
            self.pin_update_on_hover = value;
        }
        // The per-anchor answers, read the same way. A file holding only the retired single
        // key seeds all three with it, so a user who picked one walk keeps it; a value
        // that names neither way leaves that anchor at its fresh default. The old key is
        // not written back (see `to_ini`), so the seeding happens once.
        let legacy = ini
            .get(CONFIG_SECTION, "pin_nav_file_types")
            .and_then(|value| PinNavFileTypes::from_str(&value));
        if let Some(value) = ini.get(CONFIG_SECTION, "pin_nav_video") {
            if let Some(mode) = PinNavFileTypes::from_str(&value) {
                self.pin_nav_video = mode;
            }
        } else if let Some(mode) = legacy {
            self.pin_nav_video = mode;
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "pin_nav_audio") {
            if let Some(mode) = PinNavFileTypes::from_str(&value) {
                self.pin_nav_audio = mode;
            }
        } else if let Some(mode) = legacy {
            self.pin_nav_audio = mode;
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "pin_nav_others") {
            if let Some(mode) = PinNavFileTypes::from_str(&value) {
                self.pin_nav_others = mode;
            }
        } else if let Some(mode) = legacy {
            self.pin_nav_others = mode;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "follow_cursor") {
            self.follow_cursor = value;
        }
        // How far a preview is kept off its item, by the name it is written under now. The yes
        // or no question this used to be is not read: an older name is a line the app does not
        // write, and the file is written again without it (see `differs`).
        if let Some(value) = ini.get(CONFIG_SECTION, "avoid_mode") {
            if let Some(mode) = AvoidMode::from_str(&value) {
                self.avoid_mode = mode;
            }
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "same_file_rehover_delay_ms") {
            self.same_file_rehover_delay_ms = value;
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "settling_delay_ms") {
            self.settling_delay_ms = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "prioritize_keyboard") {
            self.prioritize_keyboard = value;
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "tick_ms") {
            self.tick_ms = sanitize_tick_ms(value);
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "spinner_delay_ms") {
            self.spinner_delay_ms = sanitize_spinner_delay_ms(value);
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "pin_zoom_bounce_ms") {
            self.pin_zoom_bounce_ms = sanitize_pin_zoom_bounce_ms(value);
        } else if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "pin_bounce_ms") {
            self.pin_zoom_bounce_ms = sanitize_pin_zoom_bounce_ms(value);
        } else if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "zoom_bounce_ms") {
            self.pin_zoom_bounce_ms = sanitize_pin_zoom_bounce_ms(value);
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pin_zoom_bounce_gate") {
            self.pin_zoom_bounce_gate = value;
        } else if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pin_zoom_bounce_enabled") {
            self.pin_zoom_bounce_gate = value;
        } else if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pin_bounce_gate") {
            self.pin_zoom_bounce_gate = value;
        } else if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pin_bounce_enabled") {
            self.pin_zoom_bounce_gate = value;
        } else if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "bounce_gate") {
            self.pin_zoom_bounce_gate = value;
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "pin_zoom_bounce_speed_ms") {
            self.pin_zoom_bounce_speed_ms = sanitize_pin_zoom_bounce_speed_ms(value);
        } else if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "pin_zoom_bounce_snap_ms") {
            self.pin_zoom_bounce_speed_ms = sanitize_pin_zoom_bounce_speed_ms(value);
        } else if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "pin_bounce_speed_ms") {
            self.pin_zoom_bounce_speed_ms = sanitize_pin_zoom_bounce_speed_ms(value);
        } else if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "bounce_speed_ms") {
            self.pin_zoom_bounce_speed_ms = sanitize_pin_zoom_bounce_speed_ms(value);
        } else if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "bounce_ms") {
            self.pin_zoom_bounce_speed_ms = sanitize_pin_zoom_bounce_speed_ms(value);
        }
        if let Ok(Some(value)) = ini.getfloat(CONFIG_SECTION, "pin_zoom_bounce_power") {
            self.pin_zoom_bounce_power = sanitize_pin_zoom_bounce_power(value as f32);
        } else if let Ok(Some(value)) = ini.getfloat(CONFIG_SECTION, "pin_bounce_power") {
            self.pin_zoom_bounce_power = sanitize_pin_zoom_bounce_power(value as f32);
        } else if let Ok(Some(value)) = ini.getfloat(CONFIG_SECTION, "pin_zoom_bounce_past") {
            self.pin_zoom_bounce_power = sanitize_pin_zoom_bounce_power(value as f32);
        } else if let Ok(Some(value)) = ini.getfloat(CONFIG_SECTION, "bounce_power") {
            self.pin_zoom_bounce_power = sanitize_pin_zoom_bounce_power(value as f32);
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "webp_playback_fps") {
            if let Ok(value) = u32::try_from(value) {
                self.webp_playback_fps = sanitize_webp_playback_fps(value);
            }
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "image_cache_mb") {
            if let Ok(value) = u32::try_from(value) {
                self.image_cache_mb = sanitize_image_cache_mb(value);
            }
        }
        // A picture's backdrop, by its own name. The one setting every preview was drawn over
        // before each kind had one of its own is not read from a file that still names it: the
        // line is one the app does not write, and it goes with the next write (see `differs`).
        if let Some(value) = ini.get(CONFIG_SECTION, "image_background") {
            if let Some(background) = TransparentBackground::from_str(&value) {
                self.image_background = background;
            }
        }
        // The band beside a zoomed picture's backdrop, by its own name: it is the window's
        // answer rather than the file's, so it is a key of its own — and the one backdrop
        // the menu has no item for is read as the one this setting starts at, the way a
        // texture's and a page's are (see `sanitize_zoom_background`).
        if let Some(value) = ini.get(CONFIG_SECTION, "zoom_background") {
            if let Some(background) = TransparentBackground::from_str(&value) {
                self.zoom_background = sanitize_zoom_background(background);
            }
        }
        // The backdrop a vector drawing is drawn over, which is one setting for the two halves
        // of the kind — the documents the browser draws and the metafiles the drawing layer
        // replays — and one name for both.
        if let Some(value) = ini.get(CONFIG_SECTION, "vector_background") {
            if let Some(background) = TransparentBackground::from_str(&value) {
                self.vector_background = background;
            }
        }
        // A font's backdrop is read from its own name, which it always had: fonts are a kind of
        // its own, so there is no earlier spelling of this one to answer.
        if let Some(value) = ini.get(CONFIG_SECTION, "font_background") {
            if let Some(background) = TransparentBackground::from_str(&value) {
                self.font_background = background;
            }
        }
        // A texture's is read the same way, and for the same reason: the DDS kind is one of
        // its own, so a file written before it has nothing under this name and what it wrote
        // about a picture stays what it wrote about a picture. Two of the four backdrops are
        // not offered for a texture — the two that show what stands behind it — so a file
        // that names one of them is read as the backdrop this setting starts at, rather than
        // kept as a value the menu beside it has no item for.
        if let Some(value) = ini.get(CONFIG_SECTION, "dds_background") {
            if let Some(background) = TransparentBackground::from_str(&value) {
                self.dds_background = sanitize_dds_background(background);
            }
        }
        // A design document's is read the same way and for the same reason: the kind is
        // one of its own, so a file written before it has nothing under this name, and
        // what it wrote about a picture stays what it wrote about a picture.
        if let Some(value) = ini.get(CONFIG_SECTION, "design_background") {
            if let Some(background) = TransparentBackground::from_str(&value) {
                self.design_background = background;
            }
        }
        // A page's is read the way the others are, and through the same kind of filter: the
        // backdrop is three of the four backdrops rather than all of them, so a file that
        // names the fourth — transparency, which is what a page was drawn over before the
        // kind had a setting of its own — is read as the backdrop this setting starts at,
        // rather than kept as a value the menu beside it has no item for.
        if let Some(value) = ini.get(CONFIG_SECTION, "html_background") {
            if let Some(background) = TransparentBackground::from_str(&value) {
                self.html_background = sanitize_html_background(background);
            }
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "video_volume") {
            if let Ok(value) = u32::try_from(value) {
                self.video_volume = sanitize_volume(value);
            }
        }
        // And the sound's own, which a file written before the kind existed has no key for: a
        // fresh installation's level is what it plays at until someone says otherwise.
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "audio_volume") {
            if let Ok(value) = u32::try_from(value) {
                self.audio_volume = sanitize_volume(value);
            }
        }
        // Where a sound starts, which is read the way every other named choice is: a file
        // written before the setting existed has no key for it and leaves the setting at the
        // way this build starts — where the sound was left — and a value that names no way of
        // starting one is left there too.
        if let Some(value) = ini.get(CONFIG_SECTION, "audio_seek") {
            if let Some(seek) = AudioSeek::from_str(&value) {
                self.audio_seek = seek;
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "beat_hz") {
            if let Some(hz) = BeatHz::from_str(&value) {
                self.beat_hz = hz;
            }
        }
        // Where a *pinned* sound starts, read the way the hover's own is: a
        // file written before the setting existed has no key for it and leaves
        // the setting at the way this build starts — the beginning — and a
        // value that names no way of starting one is left there too.
        if let Some(value) = ini.get(CONFIG_SECTION, "pin_mode_audio_seek") {
            if let Some(seek) = AudioSeek::from_str(&value) {
                self.pin_mode_audio_seek = seek;
            }
        }
        // What a pinned sound does at the end of its file, and which file a forward step of
        // its walk takes — the pin's own end of file and its shuffle, read the way every
        // named choice and yes-or-no setting is. A file written before either key existed
        // leaves the setting where this build starts, and a value that names no end of file
        // is left there too.
        //
        // A file that still names one of the keys these replaced is seeded from it, once:
        // what is read here is written back under the names this build writes, and the
        // retired keys go with it (see `differs`).
        if let Some(value) = ini.get(CONFIG_SECTION, "pin_mode_audio_end_of_file") {
            if let Some(end) = PinEndOfFile::from_str(&value) {
                self.pin_mode_audio_end_of_file = end;
            }
        } else if let Some(value) = ini.get(CONFIG_SECTION, "pin_mode_audio_jump") {
            if let Some((end, shuffle)) = retired_end_of_file(&value) {
                self.pin_mode_audio_end_of_file = end;
                self.pin_mode_audio_shuffle |= shuffle;
            }
        } else if let Ok(Some(looped)) = ini.getboolcoerce(CONFIG_SECTION, "pin_mode_audio_loop") {
            // The pair older still: a loop left off is the sound stepping to the one
            // beside it, and what the shuffle beside it said is read as the switch it now
            // is, below.
            if !looped {
                self.pin_mode_audio_end_of_file = PinEndOfFile::Continue;
            }
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pin_mode_audio_shuffle") {
            self.pin_mode_audio_shuffle = value;
        }
        // A file from the era of the master switch: a switch that was off is the sound left
        // where its file ended, which is what `Pause` says.
        if let Ok(Some(false)) = ini.getboolcoerce(CONFIG_SECTION, "pin_mode_audio_jump_enabled") {
            self.pin_mode_audio_end_of_file = PinEndOfFile::Pause;
        }
        // The film's own answer to the same two questions, seeded the same way: a file that
        // says nothing about either leaves both where this build starts.
        if let Some(value) = ini.get(CONFIG_SECTION, "pin_mode_video_end_of_file") {
            if let Some(end) = PinEndOfFile::from_str(&value) {
                self.pin_mode_video_end_of_file = end;
            }
        } else if let Some(value) = ini.get(CONFIG_SECTION, "pin_mode_video_jump") {
            if let Some((end, shuffle)) = retired_end_of_file(&value) {
                self.pin_mode_video_end_of_file = end;
                self.pin_mode_video_shuffle |= shuffle;
            }
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pin_mode_video_shuffle") {
            self.pin_mode_video_shuffle = value;
        }
        // And the picture's own walk, read the same way: a file written before the key
        // existed leaves the walk in the folder's own order, where this build starts.
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "pin_mode_image_shuffle") {
            self.pin_mode_image_shuffle = value;
        }
        if let Ok(Some(false)) = ini.getboolcoerce(CONFIG_SECTION, "pin_mode_video_jump_enabled") {
            self.pin_mode_video_end_of_file = PinEndOfFile::Pause;
        }
        // Whether a hovered film or sound loops: on, it goes round; off, it stops at the
        // end. The key it is written under has been renamed once, and the name it had is
        // read here so that a file written before the rename keeps its answer.
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "video_loop") {
            self.video_loop = value;
        } else if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "video_jump_enabled") {
            self.video_loop = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "audio_loop") {
            self.audio_loop = value;
        } else if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "audio_jump_enabled") {
            self.audio_loop = value;
        }
        // Whether a sound's loudness is measured and brought to one level, which a file
        // written before the setting existed has no key for: a fresh installation does not
        // normalize, and a
        // file that says nothing about it is left where it starts.
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "normalize_volume") {
            self.normalize_volume = value;
        }
        // And the video's own, which is off for a file that says nothing about it: a soundtrack is
        // measured only where it was asked for (see `normalize_video_volume`).
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "normalize_video_volume") {
            self.normalize_video_volume = value;
        }
        // Whether a level turned on a pin is the level the next preview is played at. A file written
        // before the setting existed has no key for it, so a fresh installation remembers nothing
        // and a file that says nothing about it is left where it starts.
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "remember_audio_volume") {
            self.remember_audio_volume = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "remember_video_volume") {
            self.remember_video_volume = value;
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "preview_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.preview_scale = scale;
            }
        }
        // A video's scale is written the way a picture's is, and read the same way. A file with no
        // key for it leaves the setting where a fresh installation starts: the picture scale is
        // not an answer to this question, and a file that says nothing about it is a file that
        // says nothing about it (see `differs`).
        if let Some(value) = ini.get(CONFIG_SECTION, "video_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.video_scale = scale;
            }
        }
        // A sound's card is sized by its own share of the room the display has,
        // read through the sound's own bounds rather than the picture's: what the
        // number is a percentage of is the display rather than a size the file
        // asks for, a hand-edited percentage past the whole display is the whole
        // display, and a `0` is the tenth the setting starts at.
        if let Some(value) = ini.get(CONFIG_SECTION, "audio_scale") {
            if let Some(scale) = PreviewScale::from_audio_str(&value) {
                self.audio_scale = scale;
            }
        }
        // And an animation's, the same way again.
        if let Some(value) = ini.get(CONFIG_SECTION, "animated_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.animated_scale = scale;
            }
        }
        // A drawing's scale is the one setting an SVG document and a metafile share — they
        // are one kind in the tray — and it is read from its own name. What the number is a
        // percentage of is the room the display has rather than a size the file asks for.
        if let Some(value) = ini.get(CONFIG_SECTION, "vector_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.vector_scale = scale;
            }
        }
        // A page's scale is read the same way and against the same whole: what the
        // number is a percentage of is the room the display has, one setting per kind
        // of document.
        if let Some(value) = ini.get(CONFIG_SECTION, "ebook_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.ebook_scale = scale;
            }
        }
        // A document's scale, which is the one setting both halves of the `Document` kind
        // answer to: a page of an Office document, and a page the render engine drew.
        if let Some(value) = ini.get(CONFIG_SECTION, "document_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.document_scale = scale;
            }
        }
        // A specimen's scale is read the same way again, against the share of the display
        // the box `font_preview` measures a font at takes.
        if let Some(value) = ini.get(CONFIG_SECTION, "font_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.font_scale = scale;
            }
        }
        // And a design document's, against the room the display has for the picture the
        // file keeps of the document.
        if let Some(value) = ini.get(CONFIG_SECTION, "design_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.design_scale = scale;
            }
        }
        // A page of text's is the one read the same way again, against the room it is
        // measured in rather than against a size of its own — a file that names no share for
        // it leaves it where a fresh installation starts, which is the whole of the display.
        if let Some(value) = ini.get(CONFIG_SECTION, "text_scale") {
            if let Some(scale) = PreviewScale::from_str(&value) {
                self.text_scale = scale;
            }
        }
        // And which face of a collection the specimen is of, in the numbering the tray's
        // `Font Face` submenu offers it in.
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "ttc_face") {
            if let Ok(value) = u32::try_from(value) {
                self.ttc_face = sanitize_ttc_face(value);
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
        // Left where it is where the key is not written at all, which is every file written
        // before the setting existed: the page is the markup until the tray says otherwise.
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "render_html") {
            self.render_html = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "image_preview_enabled") {
            self.image_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "video_preview_enabled") {
            self.video_preview_enabled = value;
        }
        // The sound kind's own switch, read the same way and left where it is where the name
        // is not written at all — which is every file written before the kind existed.
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "audio_preview_enabled") {
            self.audio_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "text_preview_enabled") {
            self.text_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "ebook_preview_enabled") {
            self.ebook_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "archive_preview_enabled") {
            self.archive_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "document_preview_enabled") {
            self.document_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "font_preview_enabled") {
            self.font_preview_enabled = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "design_preview_enabled") {
            self.design_preview_enabled = value;
        }
        // The vector kind's switch is read from its own name, and stays where it is where the
        // name is not written at all.
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "vector_preview_enabled") {
            self.vector_preview_enabled = value;
        }
        // The budget of the folder both engines' pages are kept in, which is one setting where
        // there used to be two — each engine had a cache of its own. A file an older build
        // wrote names one or both of those keys, and what it asks for is honoured once: the
        // larger of the two, since one budget now replaces both, and neither key is written
        // again by this build.
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "document_cache_mb") {
            if let Ok(value) = u32::try_from(value) {
                self.document_cache_mb = sanitize_document_cache_mb(value);
            }
        } else {
            let older = ["libre_cache_mb", "office_cache_mb"]
                .iter()
                .filter_map(|key| ini.getuint(CONFIG_SECTION, key).ok().flatten())
                .filter_map(|value| u32::try_from(value).ok())
                .map(sanitize_document_cache_mb)
                .max();
            if let Some(value) = older {
                self.document_cache_mb = value;
            }
        }
        // The budget of the folder the image converter's developed pictures are kept in, which is
        // a cache of its own and not the one above: what it holds are pictures, and what the
        // documents' budget holds are pages — see `document_cache`.
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "image_disk_cache_mb") {
            if let Ok(value) = u32::try_from(value) {
                self.image_disk_cache_mb = sanitize_image_disk_cache_mb(value);
            }
        }
        // And the budget of the folder a film's own subtitle tracks are copied into, which is a
        // cache of its own for the reason the one above is: what it holds are subtitle files and
        // their fonts rather than developed pictures or the pages engines drew (see `subtitle_files`).
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "general_disk_cache_mb") {
            if let Ok(value) = u32::try_from(value) {
                self.general_disk_cache_mb = sanitize_general_disk_cache_mb(value);
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "office_engine") {
            if let Some(engine) = OfficeEngine::from_str(&value) {
                self.office_engine = engine;
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "video_engine") {
            if let Some(engine) = VideoEngine::from_str(&value) {
                self.video_engine = engine;
            }
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "video_engine_fallback") {
            self.video_engine_fallback = value;
        }
        // A value that cannot be read — a key that is gone, or one that is not text — is nobody's
        // pick, which is the empty path the setting starts at (see `DEFAULT_MPV_PATH`).
        if let Some(value) = ini.get(CONFIG_SECTION, "mpv_path") {
            self.mpv_path = sanitize_mpv_path(&value);
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "video_subtitles") {
            self.video_subtitles = value;
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "office_engine_idle") {
            if let Some(idle) = EngineIdle::from_str(&value) {
                self.office_engine_idle = idle;
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "webview_idle") {
            if let Some(idle) = EngineIdle::from_str(&value) {
                self.webview_idle = idle;
            }
        }
        if let Some(value) = ini.get(CONFIG_SECTION, "libreoffice_idle") {
            if let Some(idle) = EngineIdle::from_str(&value) {
                self.libreoffice_idle = idle;
            }
        }
        if let Ok(Some(value)) = ini.getuint(CONFIG_SECTION, "afk_timer_seconds") {
            self.afk_timer_seconds = sanitize_afk_timer_secs(value);
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "office_engine_persistent") {
            self.office_engine_persistent = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "webview_persistent") {
            self.webview_persistent = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "libreoffice_persistent") {
            self.libreoffice_persistent = value;
        }
        if let Ok(Some(value)) = ini.getboolcoerce(CONFIG_SECTION, "video_hw_accel") {
            self.video_hw_accel = value;
        }
        // A budget is written in gigabytes and may be fractional — `0.5` is a small
        // machine's ceiling — so it is read as the number it is rather than as a
        // count of them.
        if let Ok(Some(value)) = ini.getfloat(CONFIG_SECTION, "decode_budget_gb") {
            self.decode_budget_gb = sanitize_decode_budget_gb(value as f32);
        }
        // A curve is read by name, and a name that is not one of them is answered with the
        // default this field already holds rather than with a curve picked at random: what
        // the setting says has to be a curve for it to be used.
        if let Some(value) = ini.get(CONFIG_SECTION, "hdr_tone_map") {
            if let Some(curve) = Curve::from_str(&value) {
                self.hdr_tone_map = curve;
            }
        }
        if let Ok(Some(value)) = ini.getfloat(CONFIG_SECTION, "hdr_exposure") {
            self.hdr_exposure = sanitize_hdr_exposure(value as f32);
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
        // Every list the configuration holds is read out of the file by the table, one row
        // each (see `formats::lists`).
        //
        // A list is what the file says it is, and a key that is gone is a list the file no
        // longer has: the built-in entries are put back, and the file is written out again
        // because it does not say what the app is using. An empty value is not the same thing
        // — it is a list the user emptied, and it is kept as written.
        //
        // What an older file’s list needs — `svg` and `svgz` given up to the vector list, the
        // formats Windows has a codec for, `dds`, the order the entries are written in, and a
        // list this app has since changed a name in or out of — was seen to by the repair that
        // ran over the file before this did, and which list it walks is a fact about the row
        // rather than about this read.
        lists::read_all(ini, self);
    }
}

/// The end of file a retired `pin_mode_*_jump` value names, and whether it also turned the
/// shuffle on: `shuffle` was a random file of the same kind at the end of a file, which this
/// build spells as `Continue` with the shuffle switch beside the setting — so a file that
/// named it seeds both, and the walk it described is the walk it still gets. A value that
/// names no end of file at all is `None`, which leaves the setting where it was rather than
/// guessing at one.
fn retired_end_of_file(value: &str) -> Option<(PinEndOfFile, bool)> {
    match value.trim().to_ascii_lowercase().as_str() {
        "shuffle" | "random" => Some((PinEndOfFile::Continue, true)),
        _ => PinEndOfFile::from_str(value).map(|end| (end, false)),
    }
}

//! The file the configuration lives in: where it is, how it is read, and how it is
//! written — the headings it is grouped under, the order it is written in, and the
//! repairs a file written by an earlier build is put through.

use std::fs;
use std::path::{Path, PathBuf};

use configparser::ini::Ini;
use directories::BaseDirs;

use crate::config::theme_files;
use crate::formats::lists;

use super::app_config::AppConfig;
use super::defaults::{CONFIG_SECTION, DEFAULT_VIDEO_ENGINE};

/// The headings the settings section is written under, in the order the tray lists its
/// menus: the menu a setting is changed from is the menu it is found under, so the file
/// reads the way the tray does rather than as one alphabetical run of fifty keys.
///
/// The table is also the list of what this app writes, key for key — a test holds `save` to it —
/// so a key in it that a file does not have is a setting the file is missing, and a file missing
/// one is written out again with it (see `differs`). A setting added to the app is therefore a
/// setting that reaches the files that already exist, with nothing else to remember.
///
/// The file lists keep sections of their own below it — `[image]`, `[video]` and the rest
/// — because a list of extensions is a collection rather than a setting, and a section is
/// what the format has for a collection. A heading is a comment instead, which is all the
/// grouping needs to be: nothing reads it back, so the shape of the file stays something
/// the app writes rather than something it has to parse. `Advanced` is where the settings
/// the tray has no menu item for are kept — the ones a hand edit reaches and a menu does
/// not — and a setting `save` writes that is missing from the table is written last of
/// all under `; Ungrouped`, which is where a heading that was forgotten shows up.
///
/// Nothing reads a heading back, and a file grouped by an earlier arrangement of the tray
/// holds exactly the settings of one grouped this way — so the headings a file lists, and the
/// order it lists them in, are the one thing that tells the two apart, and the one thing that
/// brings such a file to be written again (see `headings_are_old`).
pub(super) const SETTING_GROUPS: &[(&str, &[&str])] = &[
    (
        "General",
        &[
            "check_for_updates",
            "pin_enabled",
            "pin_key",
            "pin_nav_file_types",
            "pin_pause_audio",
            "pin_pause_video",
            "pin_update_enabled",
            "pin_update_on_hover",
            "preview_enabled",
            "run_at_startup",
        ],
    ),
    (
        "Preview Types",
        &[
            "archive_preview_enabled",
            "audio_preview_enabled",
            "design_preview_enabled",
            "document_preview_enabled",
            "ebook_preview_enabled",
            "font_preview_enabled",
            "image_preview_enabled",
            "text_preview_enabled",
            "vector_preview_enabled",
            "video_preview_enabled",
        ],
    ),
    (
        "Text Preview",
        &["markdown_mode", "render_html", "text_font_scale", "theme"],
    ),
    (
        "Timing",
        &[
            "hover_delay_ms",
            "prioritize_keyboard",
            "same_file_rehover_delay_ms",
            "settling_delay_ms",
            "trigger_key",
            "trigger_key_affect_pin_mode",
            "trigger_key_enabled",
            "trigger_key_mode",
        ],
    ),
    ("Placement", &["avoid_mode", "follow_cursor"]),
    (
        "Scaling",
        &[
            "animated_scale",
            "audio_scale",
            "design_scale",
            "document_scale",
            "ebook_scale",
            "font_scale",
            "preview_scale",
            "text_scale",
            "vector_scale",
            "video_scale",
        ],
    ),
    (
        "Background",
        &[
            "dds_background",
            "design_background",
            "font_background",
            "html_background",
            "image_background",
            "vector_background",
        ],
    ),
    (
        "Volume",
        &[
            "audio_seek",
            "audio_volume",
            "normalize_video_volume",
            "normalize_volume",
            "pin_mode_audio_loop",
            "pin_mode_audio_seek",
            "pin_mode_audio_shuffle",
            "remember_audio_volume",
            "remember_video_volume",
            "video_subtitles",
            "video_volume",
        ],
    ),
    (
        "Performance",
        &[
            "decode_budget_gb",
            "document_cache_mb",
            "general_disk_cache_mb",
            "image_cache_mb",
            "image_disk_cache_mb",
            "tick_ms",
            "video_hw_accel",
        ],
    ),
    (
        "Engine",
        &[
            "afk_timer_seconds",
            "libreoffice_idle",
            "libreoffice_persistent",
            "office_engine",
            "office_engine_idle",
            "office_engine_persistent",
            "video_engine",
            "video_engine_fallback",
            "webview_idle",
            "webview_persistent",
        ],
    ),
    (
        "Advanced",
        &[
            "hdr_exposure",
            "hdr_tone_map",
            "spinner_delay_ms",
            "text_scroll_far_edge_grace_pixels",
            "ttc_face",
            "webp_playback_fps",
        ],
    ),
];

/// One run of keys as the file writes them: the key, the value when it has one, and the
/// line's newline.
fn write_keys(out: &mut String, keys: &[(&str, &Option<String>)]) {
    for (key, value) in keys {
        out.push_str(key);
        if let Some(value) = value {
            out.push('=');
            out.push_str(value);
        }
        out.push('\n');
    }
}

/// The configuration as the text a person reads: the settings section first, under the
/// headings the tray lists its menus under, then the sections that hold the file lists in
/// alphabetical order, and the keys of every group in alphabetical order within it.
///
/// The `Ini` the values are collected in keeps its sections and keys in hash maps, so what
/// it writes by itself is a different order every time — which is what shuffled a
/// hand-edited file under the person editing it. What is returned here is the same text
/// that writer produces, in an order that does not move.
pub(super) fn ordered_text(ini: &Ini) -> String {
    let map = ini.get_map_ref();

    let mut sections: Vec<&str> = map.keys().map(String::as_str).collect();
    // Settings first, the rest alphabetically: the flag is false only for the
    // settings section, and false sorts ahead of true.
    sections.sort_unstable_by_key(|section| (*section != CONFIG_SECTION, *section));

    let mut out = String::new();
    for section in sections {
        let Some(keys) = map.get(section) else {
            continue;
        };

        // Every section but the first is kept clear of what came before it: the settings
        // section ends with the last heading's last key, and a list section beginning right
        // under it reads as another key of that heading.
        if !out.is_empty() {
            out.push('\n');
        }

        out.push('[');
        out.push_str(section);
        out.push_str("]\n");

        let mut keys: Vec<(&str, &Option<String>)> = keys
            .iter()
            .map(|(key, value)| (key.as_str(), value))
            .collect();

        // A section below the settings one is one list and is written as one run.
        if section != CONFIG_SECTION {
            keys.sort_unstable_by_key(|(key, _)| *key);
            write_keys(&mut out, &keys);
            continue;
        }

        let mut written: Vec<&str> = Vec::new();
        for (heading, heading_keys) in SETTING_GROUPS {
            let mut in_heading: Vec<(&str, &Option<String>)> = keys
                .iter()
                .copied()
                .filter(|(key, _)| heading_keys.contains(key))
                .collect();
            if in_heading.is_empty() {
                continue;
            }
            in_heading.sort_unstable_by_key(|(key, _)| *key);

            // No blank line above the first heading: what it would separate it from is
            // the section's own name.
            if !written.is_empty() {
                out.push('\n');
            }
            out.push_str("; ");
            out.push_str(heading);
            out.push('\n');
            write_keys(&mut out, &in_heading);
            written.extend(in_heading.iter().map(|(key, _)| *key));
        }

        let mut ungrouped: Vec<(&str, &Option<String>)> = keys
            .into_iter()
            .filter(|(key, _)| !written.contains(key))
            .collect();
        if !ungrouped.is_empty() {
            ungrouped.sort_unstable_by_key(|(key, _)| *key);
            out.push_str("\n; Ungrouped\n");
            write_keys(&mut out, &ungrouped);
        }
    }

    out
}

/// Whether a file is one whose settings were grouped the way an earlier build grouped them.
///
/// A heading is a comment, so a file grouped by one arrangement holds exactly the settings of a
/// file grouped by another: nothing is read back from the difference and every key is read the
/// same either way. What the headings a file does list, and the order it lists them in, do say
/// is whether it was written before a heading was added, before a setting moved out of another
/// one, or before the menus themselves were rearranged — and the grouping is the app's to
/// write, so such a file is written again rather than left with the shape it happens to have.
///
/// One write puts every heading back at once, in the order of the table, so a file this has been
/// through lists them as this build lists them and is left alone the next time it is read. That
/// is the whole point of asking the question about the text rather than about the parse: the
/// file is read once a second while it is being watched, and a file the app has just written
/// must not be a file there is something to write. The question is asked of the text for that
/// reason too — a heading is a comment, and a parse keeps everything about a file except the
/// comments (`load`, `reload_from_disk`).
pub(super) fn headings_are_old(text: &str) -> bool {
    // The headings the file lists, in the order it lists them. Compared against the whole line
    // rather than searched for in the file, so a key named after a heading — or a heading
    // written inside a value — is not read as one, and a line's own trailing space is not part
    // of the comparison: an editor that leaves one is not a reason to write the file again.
    let listed: Vec<&str> = text
        .lines()
        .filter_map(|line| {
            SETTING_GROUPS
                .iter()
                .map(|(heading, _)| *heading)
                .find(|heading| line.trim_end().strip_prefix("; ") == Some(*heading))
        })
        .collect();

    // Every heading the table names, in the order the tray lists them. A file missing one, or
    // listing them in another order, is a file written before the settings were arranged this
    // way — and a heading the table does not name at all, `; Ungrouped` among them, is no part
    // of the question: what is compared is the headings this app writes and the places they
    // are written in.
    let written: Vec<&str> = SETTING_GROUPS.iter().map(|(heading, _)| *heading).collect();

    listed != written
}

/// The two files the installer leaves behind for the app to find, taken as it starts.
///
/// The installer never writes `config.ini` — a second writer of that file would carry the
/// defaults of whenever the installer was built, and an installation made that way would keep
/// them for good (see the packaging metadata in `Cargo.toml`) — so a box on its page is a file
/// beside `config.ini` instead, and the app is what reads it. Each is removed as it is read:
/// what it asks for is applied once, and one left behind would be applied again on a later
/// start, after the user had set something of their own.
pub(super) fn take_reset_markers(folder: &Path) -> (bool, bool) {
    let settings = folder.join("reset-settings.marker");
    let extensions = folder.join("reset-extensions.marker");

    let reset_settings = settings.exists();
    let reset_extensions = extensions.exists();

    if reset_settings {
        let _ = fs::remove_file(&settings);
    }
    if reset_extensions {
        let _ = fs::remove_file(&extensions);
    }

    (reset_settings, reset_extensions)
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

    /// The configuration as the file has it, with whatever the file had wrong or missing put
    /// right.
    ///
    /// The file is read in one of two ways: there is none, so a fresh installation is written —
    /// or there is one, and the lists it holds that are this app's own older ones are brought up
    /// before anything is read from it (`formats::lists::repair_older_lists`), and the settings
    /// it holds under
    /// headings an older build wrote are grouped again the way this build groups them
    /// (`headings_are_old`) — since the file is what the user edits and a key left missing would
    /// be repaired again on every load. It is written back where that left something to write,
    /// and where it did not, the file on disk is left exactly as it is: a file this build wrote,
    /// with nothing deleted from it and nothing added to it since, is one there is nothing to say
    /// about.
    pub fn load() -> Self {
        let mut config = Self::default();

        // The folder exists before anything can list it, so that adding a theme is
        // dropping a file into a path the app can name rather than creating one.
        theme_files::ensure();

        let Some(path) = Self::config_path() else {
            return config;
        };

        config.is_first_run = !path.exists();
        if !config.is_first_run {
            let mut ini = Ini::new();

            // A file that cannot be read is left exactly where it is, and what it holds is not
            // guessed at: writing the defaults over it would be a write with nothing to do with
            // the file, which is the one kind of write this app does not make. A file that is
            // missing by then — or was never there — is the fresh installation below.
            //
            // The text is read here rather than by `Ini::load`, which reads and parses the same
            // file: a parse keeps nothing of the comments, and whether the settings sit under the
            // headings this build writes them under is a question about the text (see
            // `headings_are_old`).
            if let Ok(text) = fs::read_to_string(&path) {
                let old_headings = headings_are_old(&text);

                if ini.read(text).is_ok() {
                    let repaired = lists::repair_older_lists(&mut ini) || old_headings;
                    config.apply_ini(&ini);

                    // An engine this machine has not got is put back to `Best` at a start: the
                    // choice names a player that is not here, which the router would ignore anyway,
                    // and a row marked for it is greyed besides (see `VideoEngine::installed`).
                    if !config.video_engine.installed() {
                        config.video_engine = DEFAULT_VIDEO_ENGINE;
                    }

                    // The file is written again where the repair had a list to bring up or a
                    // heading to put back, and where it does not hold what this app writes: a
                    // setting the file does not have, one whose value is not the value the app
                    // reads it back as — which is every value the app could not read at all —
                    // and any key that is not one the app writes, which is where a name this
                    // app no longer uses goes.
                    if repaired || config.differs(&ini) {
                        config.save();
                    }
                }
            }
        }

        if config.is_first_run {
            config.save();
        }

        // What the installer left behind, if it left anything: the two boxes on its page are two
        // files beside this one, and this is where they are read — after the file itself, so what
        // they ask for is what the configuration ends up holding, and before it is handed out, so
        // nothing is ever read out of a configuration that is about to be replaced.
        if let Some(folder) = Self::folder() {
            let (reset_settings, reset_extensions) = take_reset_markers(&folder);

            if reset_settings || reset_extensions {
                if reset_settings {
                    config.reset_to_recommended();
                }
                if reset_extensions {
                    config.reset_extension_lists();
                }

                config.save();
            }
        }

        config
    }

    /// Read the file again, after the watcher saw it change, with the same repairs a start puts
    /// it through and the same question about whether it says what the app is using.
    ///
    /// Both belong here as much as they do at a start: the file is what the user edits, and an
    /// edit that writes a value the app cannot read, adds a key of their own, or names a setting
    /// the way an older build named it, is one to put right there and then rather than at the
    /// next start — as is a file an older build wrote and this one has not read since. It settles
    /// the same way a start does — what it writes is a file that needs nothing, so the write the
    /// watcher sees after it is a file nothing further is done to.
    pub fn reload_from_disk(&mut self) {
        if let Some(path) = Self::config_path() {
            let mut ini = Ini::new();
            if let Ok(text) = fs::read_to_string(&path) {
                let old_headings = headings_are_old(&text);

                if ini.read(text).is_ok() {
                    let repaired = lists::repair_older_lists(&mut ini) || old_headings;
                    self.apply_ini(&ini);

                    // The engine repair a start puts the file through, for the same reason an edit
                    // to the file is put through the others: a choice naming a player this machine
                    // has not got is written back as `Best` (see `VideoEngine::installed`).
                    if !self.video_engine.installed() {
                        self.video_engine = DEFAULT_VIDEO_ENGINE;
                    }

                    if repaired || self.differs(&ini) {
                        self.save();
                    }
                }
            }
        }
    }

    /// Write the configuration out: `to_ini` says what it holds, and `ordered_text` says what
    /// the text of the file is.
    pub fn save(&self) {
        if let Some(path) = Self::config_path() {
            if let Some(parent) = path.parent() {
                let _ = fs::create_dir_all(parent);
            }

            let _ = fs::write(&path, ordered_text(&self.to_ini()));
        }
    }

    /// Every setting back at what this build recommends, with the extension lists left exactly
    /// as they are.
    ///
    /// What is recommended is what a configuration that has just been made holds (`Default`),
    /// so this is the file a fresh installation would be given — written as values rather than
    /// as a removal of the file, since a `config.ini` that is missing is a first run and all
    /// that follows from one. `is_first_run` is not a setting and is carried across: a reset
    /// asked for during a first run is still a first run.
    pub fn reset_to_recommended(&mut self) {
        let taken = lists::held(self);
        let is_first_run = self.is_first_run;

        *self = Self::default();
        self.is_first_run = is_first_run;
        lists::put(self, taken);
    }

    /// Every extension list back at the built-in one, with every other setting left alone.
    pub fn reset_extension_lists(&mut self) {
        lists::reset_built_in(self);
    }

    /// The settings that do not hold what this build recommends, as `(key, now, recommended)`.
    ///
    /// Both sides are read out of `to_ini`, so what is named is what the user would see change
    /// in the file, under the key the file writes it as. The extension lists are not part of the
    /// question — they are the other reset's business — and the answer is sorted by key, so the
    /// dialog built out of it reads the same way every time.
    pub fn settings_apart_from_recommended(&self) -> Vec<(String, String, String)> {
        let recommended = Self::default().to_ini();
        let mine = self.to_ini();
        let mut apart = Vec::new();

        let Some(keys) = recommended.get_map_ref().get(CONFIG_SECTION) else {
            return apart;
        };

        for (key, wanted) in keys {
            let now = mine.get(CONFIG_SECTION, key);

            if now != *wanted {
                apart.push((
                    key.clone(),
                    now.unwrap_or_default(),
                    wanted.clone().unwrap_or_default(),
                ));
            }
        }

        apart.sort();
        apart
    }

    /// The extension lists that are not the built-in ones, named by the section they are written
    /// under — `text` once, though that section holds two lists (`extensions` and `names`).
    pub fn lists_apart_from_built_in(&self) -> Vec<String> {
        let built_in = Self::default().to_ini();
        let mine = self.to_ini();
        let mut apart = Vec::new();

        for (section, keys) in mine.get_map_ref() {
            if section == CONFIG_SECTION {
                continue;
            }

            if keys
                .keys()
                .any(|key| mine.get(section, key) != built_in.get(section, key))
            {
                apart.push(section.clone());
            }
        }

        apart.sort();
        apart
    }
}

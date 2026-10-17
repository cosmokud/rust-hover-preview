//! The file the configuration lives in: where it is, how it is read, and how it is
//! written — the headings it is grouped under, the order it is written in, and the
//! repairs a file written by an earlier build is put through.

use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};

use configparser::ini::Ini;

#[cfg(not(test))]
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
            "pin_double_click_center",
            "pin_double_click_reset_size",
            "pin_enabled",
            "pin_key",
            "pin_nav_audio",
            "pin_nav_others",
            "pin_nav_video",
            "pin_update_enabled",
            "pin_update_on_hover",
            "preview_enabled",
            "run_at_startup",
            "theme",
        ],
    ),
    // Where a pin put away by its own minimize goes — the `Pin Mode → Minimize`
    // submenu, which is why the key sits under one heading rather than beside the pin's
    // own ones: it is what the minimize button does, not how a pin is taken up.
    ("Minimize", &["pin_minimize_to"]),
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
        &["markdown_mode", "render_html", "text_font_scale"],
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
            "zoom_background",
        ],
    ),
    (
        "Volume",
        &[
            "audio_loop",
            "audio_seek",
            "audio_volume",
            "normalize_video_volume",
            "normalize_volume",
            "pin_mode_audio_end_of_file",
            "pin_mode_audio_seek",
            "pin_mode_audio_shuffle",
            "pin_mode_image_shuffle",
            "pin_mode_video_end_of_file",
            "pin_mode_video_shuffle",
            "pin_pause_audio",
            "pin_pause_video",
            "remember_audio_volume",
            "remember_video_volume",
            "video_loop",
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
            "mpv_path",
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
            "beat_hz",
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

/// The text of a file as the app reads it, or `None` for a file this app cannot read: one that
/// holds no setting of its own, which is what a zero-length file left behind by a write that
/// was cut short is, and what a file this app did not write is.
///
/// A file that cannot be read is not guessed at. Reading nothing out of one leaves every
/// setting where a fresh one starts, and the headings and keys a file of this app's own writes
/// and this one does not are then a reason to write those defaults over it — a write with
/// nothing to do with the file, which is the one kind of write this app does not make (see
/// `load`, `reload_from_disk`).
pub(super) fn read_ini(text: &str) -> Option<Ini> {
    let mut ini = Ini::new();
    if ini.read(text.to_string()).is_err() {
        return None;
    }

    ini.get_map_ref()
        .get(CONFIG_SECTION)
        .is_some_and(|settings| !settings.is_empty())
        .then_some(ini)
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

/// The folder a test build of the app keeps its own configuration in, while a test holds one:
/// the folder of a test's own rather than the machine's own.
///
/// The machine's own configuration file is one file that a run of the app reads every second
/// (`main`'s watcher) and writes back what it read, so a test that read or wrote it changed
/// what that run had — and a value no menu offers, a number a hand never chose, stayed.
///
/// The file is one file, so one test at a time has the app's own folder.
#[cfg(test)]
static FILE_TURN: std::sync::Mutex<()> = std::sync::Mutex::new(());

#[cfg(test)]
static HELD_FOLDER: std::sync::Mutex<Option<(std::thread::ThreadId, PathBuf)>> =
    std::sync::Mutex::new(None);

/// The folder of a test's own, held while the test has the app's configuration file. The
/// app's own folder is pointed at it for as long as the guard lives, and the folder goes
/// away with it — including whatever the app wrote in it.
///
/// It is pointed at it for the thread that *took* the guard and for no other, which is the whole of
/// what keeps a test beside this one out of its round trip: the app is asked for its configuration
/// from wherever the app is running, and in a test binary that is a test's own thread — so an
/// unguarded test saving its settings would otherwise be answered with the folder of whichever
/// guarded test happened to be holding one, and its write would land in the middle of that test's
/// save-and-read-back (see `test_folder`).
#[cfg(test)]
pub(crate) struct ConfigFolderOfItsOwn {
    path: PathBuf,
    _turn: std::sync::MutexGuard<'static, ()>,
}

#[cfg(test)]
impl ConfigFolderOfItsOwn {
    /// A folder made for the test now, and the configuration file the app reads and writes
    /// in it. One folder a test rather than one the tests share, so that two tests running
    /// beside each other are not answered with each other's file.
    pub(crate) fn take() -> Self {
        // The turn is taken before the folder is: a test whose guard went away while another
        // still had the app pointed at its folder would clear it out from under it. A test that
        // panicked with the turn leaves it poisoned, and the tests after it still have to have
        // their own file.
        let _turn = FILE_TURN
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        static NEXT: std::sync::atomic::AtomicU32 = std::sync::atomic::AtomicU32::new(0);
        let path = std::env::temp_dir().join(format!(
            "rust-hover-preview-config-{}-{}",
            std::process::id(),
            NEXT.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
        ));
        let _ = fs::create_dir_all(&path);

        *HELD_FOLDER
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner()) =
            Some((std::thread::current().id(), path.clone()));

        Self { path, _turn }
    }

    /// The file the app's own configuration is, in the folder of the test's own.
    pub(crate) fn config_file(&self) -> PathBuf {
        self.path.join("config.ini")
    }
}

#[cfg(test)]
impl Drop for ConfigFolderOfItsOwn {
    fn drop(&mut self) {
        let mut held = HELD_FOLDER
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        *held = None;
        let _ = fs::remove_dir_all(&self.path);
    }
}

/// Where a test build of the app keeps its own configuration while a test does not hold a
/// folder of its own: one folder of the tests' own, never the machine's.
///
/// The held folder is handed to the thread that took the guard and to no other. A test beside this one
/// — a tray menu, a chrome button, a save of its own — is asking from its own thread, and what it must
/// not be answered with is the folder a guarded test is in the middle of a save and a read-back
/// through: that is the round trip this test's own file failed, a file that was there when it was
/// written and gone when it was read (see `HELD_FOLDER`).
#[cfg(test)]
fn test_folder() -> PathBuf {
    let held = HELD_FOLDER
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());

    match held.as_ref() {
        Some((thread, path)) if *thread == std::thread::current().id() => path.clone(),
        _ => std::env::temp_dir().join("rust-hover-preview-tests"),
    }
}

impl AppConfig {
    /// The app's own folder under the roaming profile, holding `config.ini` and
    /// the `theme` folder beside it.
    ///
    /// A test build has a folder of its own for as long as a test holds one, and never the
    /// machine's own: that file is the one a run of the app is watching (see
    /// `ConfigFolderOfItsOwn`).
    fn folder() -> Option<PathBuf> {
        #[cfg(test)]
        {
            return Some(test_folder());
        }

        #[cfg(not(test))]
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
            // A file that cannot be read is left exactly where it is, and what it holds is not
            // guessed at: writing the defaults over it would be a write with nothing to do with
            // the file, which is the one kind of write this app does not make. A file that is
            // missing by then — or was never there — is the fresh installation below.
            //
            // The text is read here rather than by `Ini::load`, which reads and parses the same
            // file: a parse keeps nothing of the comments, and whether the settings sit under the
            // headings this build writes them under is a question about the text (see
            // `headings_are_old`). A file holding no setting of its own at all — a zero-length
            // one among them — is not read at all (see `read_ini`).
            if let Ok(text) = fs::read_to_string(&path) {
                if let Some(mut ini) = read_ini(&text) {
                    let old_headings = headings_are_old(&text);

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

    /// Read the file again, after the watcher saw it change, with the same repairs a start
    /// puts it through — and none of its writing.
    ///
    /// What the file says is taken, and what it has wrong or missing is put right in what the
    /// app is using: a value it cannot read, a key of the user's own, a setting named the way an
    /// older build named it, a list of this app's own from before it grew. The file is left to
    /// say what it said — the next start is what writes it, the same way it always was.
    ///
    /// A write here would be a write the watcher sees, and the write is the whole of the
    /// trouble: the file is the one the user is holding open in the editor this app's own
    /// `Edit Config.ini` row put them in, and an editor writes the file in place — a read a
    /// second later takes half of what they are writing and writes the whole of the app's own
    /// view back over it, which is their edit gone. A write here is also what made an outside
    /// write stay: the app read what some other writer had put there, and the next thing on
    /// disk was the app's own agreeing with it (see `save`).
    pub fn reload_from_disk(&mut self) {
        if let Some(path) = Self::config_path() {
            if let Ok(text) = fs::read_to_string(&path) {
                if let Some(mut ini) = read_ini(&text) {
                    // Read through the repairs rather than writing them back: a list of this
                    // app's own from before it grew, or a file an older build grouped the way
                    // it grouped them, is what the app uses from here on, and the file keeps
                    // what it had until a start writes it.
                    lists::repair_older_lists(&mut ini);
                    self.apply_ini(&ini);

                    // The engine repair a start puts the file through, for the same reason an edit
                    // to the file is put through the others: a choice naming a player this machine
                    // has not got is answered with `Best` (see `VideoEngine::installed`).
                    if !self.video_engine.installed() {
                        self.video_engine = DEFAULT_VIDEO_ENGINE;
                    }
                }
            }
        }
    }

    /// Write the configuration out: `to_ini` says what it holds, and `ordered_text` says what
    /// the text of the file is.
    ///
    /// The text is written beside the file and moved over the one that is there rather than
    /// written into it, so that a write cut short — the process ended while it was happening,
    /// which is what a build does to a running copy so it can link over its binary
    /// (`build.rs`), or a disk that filled — leaves the file as it was, whole, rather than
    /// half-written or empty. A zero-length file is one the next start cannot read, and every
    /// setting it does not find in it is a setting it writes the default back over (see
    /// `read_ini`).
    pub fn save(&self) {
        if let Some(path) = Self::config_path() {
            if let Some(parent) = path.parent() {
                let _ = fs::create_dir_all(parent);
            }

            let text = ordered_text(&self.to_ini());
            let staged = path.with_extension("ini.new");
            // What is moved is on the disk rather than in the write cache when the move puts it
            // under the app's own name: a machine that lost the power between the two would
            // otherwise be left with a file holding nothing, which is the whole of what
            // writing beside and moving is for.
            let staged_written = fs::OpenOptions::new()
                .write(true)
                .create(true)
                .truncate(true)
                .open(&staged)
                .and_then(|mut file| {
                    file.write_all(text.as_bytes())
                        .and_then(|()| file.sync_all())
                });

            if staged_written.is_ok() {
                // The move is what puts the new text in place, so what a reader finds under
                // the app's own name is the whole text or the file that was there.
                if fs::rename(&staged, &path).is_err() {
                    // A move the machine does not allow — the file is held open where it
                    // stands, by a thing that shares it without the delete the move needs —
                    // still leaves the setting to be written down, and a setting left
                    // unwritten is one the next read takes back off the app. The text goes
                    // into the file that is there, the one write this app made before it
                    // wrote beside.
                    let _ = fs::remove_file(&staged);
                    let _ = fs::write(&path, text);
                }
            } else {
                let _ = fs::remove_file(&staged);
            }
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

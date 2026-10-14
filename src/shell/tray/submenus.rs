//! One builder per submenu, and the words each of its items is written as.
//!
//! A submenu that is built by the same function as its neighbours is a submenu whose ids
//! come out of the same shape: a base plus the position the choice was listed at, against a
//! table of the choices. What each one offers is therefore the whole of what the setting
//! holds, and a value a hand-edited `config.ini` asks for that is not one of them is shown
//! with nothing marked rather than rounded to the nearest - which is why each builder takes
//! the setting it marks rather than reading it, and why each table is named for the setting
//! it belongs to (see `tray::ids`).
//!
//! The `*_at` functions are the other half of the same arrangement: the value an id stands
//! for, so a click selects what its item named. An id past the last item the menu offered is
//! one that is not there, and selects nothing.

use super::commands::open_link;
use super::ids::{
    EngineIdleIds, AFK_TIMER_CHOICES_SECS, AUDIO_SCALE_CHOICES, AUDIO_SEEK_CHOICES, AVOID_CHOICES,
    BACKGROUND_CHOICES, BITMAP_SCALE_CHOICES, DDS_BACKGROUND_CHOICES, DOCUMENT_SCALE_CHOICES,
    ENGINE_IDLE_CHOICES, HTML_BACKGROUND_CHOICES, ID_TRAY_AFK_TIMER_BASE,
    ID_TRAY_AUDIO_SEEK_BASE,
    ID_TRAY_CODEC_BASE, ID_TRAY_ENGINE_OFFICE_LIBRE, ID_TRAY_ENGINE_OFFICE_MS, ID_TRAY_MPV_BASE,
    ID_TRAY_PIN_MODE_AUDIO_SEEK_BASE, ID_TRAY_PIN_NAV_AUDIO_ALL, ID_TRAY_PIN_NAV_AUDIO_CATEGORY,
    ID_TRAY_PIN_NAV_OTHERS_ALL, ID_TRAY_PIN_NAV_OTHERS_CATEGORY, ID_TRAY_PIN_NAV_VIDEO_ALL,
    ID_TRAY_PIN_NAV_VIDEO_CATEGORY, ID_TRAY_VIDEO_ENGINE_BASE, ID_TRAY_VIDEO_ENGINE_FALLBACK,
    PIN_END_OF_FILE_CHOICES,
    TICK_CHOICES_MS, TIMING_DELAY_CHOICES_MS, VIDEO_ENGINE_CHOICES,
};

use crate::app::dialogs;
use crate::config::config::{
    AudioSeek, AvoidMode, EngineIdle, OfficeEngine, PinEndOfFile, PinMinimizeTo, PinNavFileTypes,
    PreviewScale,
    TransparentBackground, VideoEngine, DEFAULT_AFK_TIMER_SECS, DEFAULT_AUDIO_SEEK,
    DEFAULT_AVOID_MODE, DEFAULT_DECODE_BUDGET_GB, DEFAULT_OFFICE_ENGINE,
    DEFAULT_PIN_MODE_AUDIO_SEEK,
    DEFAULT_PIN_MINIMIZE_TO, DEFAULT_PIN_NAV_AUDIO, DEFAULT_PIN_NAV_OTHERS, DEFAULT_PIN_NAV_VIDEO,
    DEFAULT_TICK_MS, DEFAULT_VIDEO_ENGINE, DEFAULT_VIDEO_ENGINE_FALLBACK,
};
use crate::engines::libreoffice_render;
use crate::formats::codecs::{self, Row};
use crate::CONFIG;
use std::path::{Path, PathBuf};
use windows::core::{w, PCWSTR};
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CheckMenuRadioItem, CreatePopupMenu, HMENU, MENU_ITEM_FLAGS, MF_BYCOMMAND,
    MF_CHECKED, MF_ENABLED, MF_GRAYED, MF_POPUP, MF_SEPARATOR, MF_STRING, MF_UNCHECKED,
};

/// One `Timing` submenu: an item per delay the setting offers, in the order the table
/// lists them, with the delay the setting is on checked and the one it starts at marked
/// as the default — which is read from the settings rather than written into the labels,
/// so a delay that becomes a default, or stops being one, moves the mark with it.
///
/// The three submenus differ only in what they select, so they are built here rather
/// than one by one: they list the same delays, and a delay added to the table is one
/// every one of them offers.
pub(super) fn timing_delay_menu(base_id: u16, selected_ms: u64, default_ms: u64) -> HMENU {
    let menu = unsafe { CreatePopupMenu().unwrap() };

    for (index, delay_ms) in TIMING_DELAY_CHOICES_MS.iter().enumerate() {
        let flags = MF_STRING
            | if *delay_ms == selected_ms {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            };
        let label = format!("{delay_ms} ms");

        append_labeled_item(
            menu,
            flags,
            base_id + index as u16,
            &default_label(&label, *delay_ms == default_ms),
        );
    }

    menu
}

/// A label with the item the setting starts at marked as the default.
///
/// Every value menu in this tray says two things: which of its items the setting is on,
/// which is the check or radio mark, and which of them a setting nobody has changed would
/// be. The second is not written into the words — a `(Default)` typed into a label is a
/// mark that goes stale the next time the default moves, and a default is a thing this app
/// has moved more than once — so it is derived here from the value the setting itself
/// starts at, which is the `DEFAULT_*` constant every menu hands in.
pub(super) fn default_label(label: &str, is_default: bool) -> String {
    if is_default {
        format!("{label} (Default)")
    } else {
        label.to_string()
    }
}

/// Append one item whose label is built rather than written into the code, which is what a
/// label carrying a default mark is: the text is encoded here and handed over, and the
/// buffer outlives the call it is handed to.
pub(super) fn append_labeled_item(menu: HMENU, flags: MENU_ITEM_FLAGS, id: u16, label: &str) {
    let label: Vec<u16> = label.encode_utf16().chain(std::iter::once(0)).collect();

    let _ = unsafe { AppendMenuW(menu, flags, id as usize, PCWSTR(label.as_ptr())) };
}

/// What a cache size is called in the menu: the size, with the one the cache starts
/// at marked as the default — the caches do not all start at the same one, which is
/// why the default is passed in rather than written into the label. The two sizes at
/// the ceiling are the only ones that are not a plain number of megabytes.
pub(super) fn cache_size_label(megabytes: u32, default_mb: u32) -> String {
    let label = match megabytes {
        1024 => "1 GB".to_string(),
        2048 => "2 GB".to_string(),
        other => format!("{other} MB"),
    };

    default_label(&label, megabytes == default_mb)
}

/// The label a `Decode Budget` item carries: the ceiling it stands for, in the unit
/// that reads best for it, with the one the app starts at marked.
pub(super) fn decode_budget_label(gigabytes: f32) -> String {
    let label = if gigabytes < 1.0 {
        format!("{} MB", (gigabytes * 1024.0).round() as u32)
    } else {
        format!("{gigabytes} GB")
    };

    default_label(&label, gigabytes == DEFAULT_DECODE_BUDGET_GB)
}

/// The `Avoid` submenu: how far a preview is kept off the item it is about, with the
/// way the setting is on marked.
pub(super) fn append_avoid_menu(parent: HMENU, label: PCWSTR, base: u16, avoid_mode: AvoidMode) {
    let menu = unsafe { CreatePopupMenu().unwrap() };

    // The labels are kept for as long as the menu is being filled out, for the same
    // reason the backdrop labels are: `AppendMenuW` is handed a pointer, so the wide
    // strings have to outlive the call that lists them.
    let labels: Vec<Vec<u16>> = AVOID_CHOICES
        .iter()
        .map(|mode| {
            avoid_label(*mode)
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();

    for (index, label) in labels.iter().enumerate() {
        let _ = unsafe {
            AppendMenuW(
                menu,
                MF_STRING,
                (base + index as u16) as usize,
                PCWSTR(label.as_ptr()),
            )
        };
    }

    // One of the ways is the setting, so one of them carries the radio mark; a way the
    // menu does not list is marked by nothing rather than by the wrong one.
    if let Some(index) = AVOID_CHOICES.iter().position(|mode| *mode == avoid_mode) {
        let _ = unsafe {
            CheckMenuRadioItem(
                menu,
                base as u32,
                (base + AVOID_CHOICES.len() as u16 - 1) as u32,
                (base + index as u16) as u32,
                MF_BYCOMMAND.0,
            )
        };
    }

    let _ = unsafe { AppendMenuW(parent, MF_STRING | MF_POPUP, menu.0 as usize, label) };
}

/// What a way of keeping a preview off an item is called in the menu: the words the
/// tray lists it under, with the way this setting starts at marked as the default.
/// The submenu is the `Avoid` one, so each way is named for itself rather than
/// prefixed with the question it answers.
pub(super) fn avoid_label(mode: AvoidMode) -> String {
    let label = match mode {
        AvoidMode::Off => "Nothing",
        AvoidMode::Filename => "Filename",
        AvoidMode::FilenameColumn => "Filename Column",
        AvoidMode::Details => "Details",
    };

    default_label(label, mode == DEFAULT_AVOID_MODE)
}

/// The `Update Preview` submenu's own row: whether a pin that is up is shown
/// the file the user picks next — the one the pointer clicks, or the one the
/// keyboard selects — named for what it follows rather than for that it is on.
/// The row is built in `menus`, where the pin's own submenus are, and the
/// words are kept here with every other label so that what a row is called is
/// asserted in one place (see `tests`).
pub(super) fn pin_update_label() -> &'static str {
    "Follow Selection"
}

/// What the two walks a pin's own **Previous** and **Next** take are called:
/// every file this build can preview, or only the files of the kind of the
/// pinned one.
pub(super) fn pin_nav_label(types: PinNavFileTypes) -> &'static str {
    match types {
        PinNavFileTypes::All => "All Files",
        PinNavFileTypes::Category => "Same Category",
    }
}

/// One `Navigation` anchor submenu: the two walks a pin anchored on this kind of
/// file takes, with the one the setting is on checked and the one it starts at marked.
fn append_pin_nav_anchor(
    parent: HMENU,
    label: PCWSTR,
    all_id: u16,
    category_id: u16,
    selected: PinNavFileTypes,
    default: PinNavFileTypes,
) {
    let menu = unsafe { CreatePopupMenu().unwrap() };
    append_labeled_item(
        menu,
        MF_STRING | if selected == PinNavFileTypes::All { MF_CHECKED } else { MF_UNCHECKED },
        all_id,
        &default_label(pin_nav_label(PinNavFileTypes::All), default == PinNavFileTypes::All),
    );
    append_labeled_item(
        menu,
        MF_STRING | if selected == PinNavFileTypes::Category { MF_CHECKED } else { MF_UNCHECKED },
        category_id,
        &default_label(
            pin_nav_label(PinNavFileTypes::Category),
            default == PinNavFileTypes::Category,
        ),
    );
    let _ = unsafe { AppendMenuW(parent, MF_STRING | MF_POPUP, menu.0 as usize, label) };
}

/// The `Navigation` submenu: three anchors — Video, Audio, Others — each with its
/// own All/Same Category pair. The walk is anchored on the first file pinned, so a
/// video walks videos while a document walks end to end (see `nav_mode_for`).
pub(super) fn append_pin_nav_menu(parent: HMENU, video: PinNavFileTypes, audio: PinNavFileTypes, others: PinNavFileTypes) {
    let menu = unsafe { CreatePopupMenu().unwrap() };
    append_pin_nav_anchor(
        menu,
        w!("Video"),
        ID_TRAY_PIN_NAV_VIDEO_ALL,
        ID_TRAY_PIN_NAV_VIDEO_CATEGORY,
        video,
        DEFAULT_PIN_NAV_VIDEO,
    );
    append_pin_nav_anchor(
        menu,
        w!("Audio"),
        ID_TRAY_PIN_NAV_AUDIO_ALL,
        ID_TRAY_PIN_NAV_AUDIO_CATEGORY,
        audio,
        DEFAULT_PIN_NAV_AUDIO,
    );
    append_pin_nav_anchor(
        menu,
        w!("Others"),
        ID_TRAY_PIN_NAV_OTHERS_ALL,
        ID_TRAY_PIN_NAV_OTHERS_CATEGORY,
        others,
        DEFAULT_PIN_NAV_OTHERS,
    );
    let _ = unsafe { AppendMenuW(parent, MF_STRING | MF_POPUP, menu.0 as usize, w!("Navigation")) };
}

/// What the two places a pinned window's minimize can put it are called, with the one the setting
/// starts at marked as the default (see `default_label` and `PinMinimizeTo`).
pub(super) fn pin_minimize_label(target: PinMinimizeTo) -> String {
    let label = match target {
        PinMinimizeTo::Bubble => "To Bubble",
        PinMinimizeTo::Taskbar => "To Taskbar",
    };

    default_label(label, target == DEFAULT_PIN_MINIMIZE_TO)
}

/// The row under each half of the `Volume` submenu's `Normalize`: whether a
/// level turned on a pinned window's own knob is the level the next preview is
/// played at — named for the level it keeps rather than for that it is
/// remembered.
pub(super) fn remember_volume_label() -> &'static str {
    "Remember Level"
}

/// The row that offers the update a check found, which is the version it is of
/// and nothing else: the whole of what it says is what it offers.
pub(super) fn update_available_label(version: &str) -> String {
    format!("Update Available (v{version})")
}

/// The `System` submenu's own label: the version the build is of, which is
/// where a user looks to see which one is running.
pub(super) fn system_menu_label(version: &str) -> String {
    format!("System (v{version})")
}

/// The `Audio → Seek` submenu: where in a file a sound starts playing — the same
/// four ways the old `Volume → Audio Seek` offered, in the same order. Renamed because it
/// sits under `Audio` now rather than under `Volume` (see `audio_seek_label`).
pub(super) fn append_seek_menu_renamed(parent: HMENU, title: PCWSTR, seek: AudioSeek) {
    append_seek_menu(parent, title, ID_TRAY_AUDIO_SEEK_BASE, audio_seek_label, seek);
}

/// The `Audio → Pin Mode → Seek` submenu: where in a file a *pinned* sound starts
/// playing — the same four ways the old `Volume → Pin Mode Audio Seek` offered, in the
/// same order. Renamed to plain `Seek` because it hangs under `Pin Mode`, beside `Shuffle`
/// and `End of File`, where the question is where a pinned file starts rather than which
/// sound is being played (see `pin_mode_audio_seek_label`).
pub(super) fn append_pin_seek_menu_renamed(parent: HMENU, title: PCWSTR, seek: AudioSeek) {
    append_seek_menu(
        parent,
        title,
        ID_TRAY_PIN_MODE_AUDIO_SEEK_BASE,
        pin_mode_audio_seek_label,
        seek,
    );
}

/// One of the tray's two seek submenus: the four ways of starting a sound,
/// each listed under `title` as `label` words it and carrying the id `base`
/// hands out, the one `seek` names marked, and the submenu hung from
/// `parent`. The two submenus are the one question asked twice — once of the
/// hover's setting and once of the pin's — so they differ in nothing but
/// those three answers (see `append_seek_menu_renamed` and
/// `append_pin_seek_menu_renamed` for what each is).
fn append_seek_menu(
    parent: HMENU,
    title: PCWSTR,
    base: u16,
    label: fn(AudioSeek) -> String,
    seek: AudioSeek,
) {
    let menu = unsafe { CreatePopupMenu().unwrap() };

    // The labels are kept for as long as the menu is being filled out, for the same reason the
    // `Avoid` labels are: `AppendMenuW` is handed a pointer, so the wide strings have to outlive
    // the call that lists them.
    let labels: Vec<Vec<u16>> = AUDIO_SEEK_CHOICES
        .iter()
        .map(|seek| {
            label(*seek)
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();

    for (index, label) in labels.iter().enumerate() {
        let _ = unsafe {
            AppendMenuW(
                menu,
                MF_STRING,
                (base + index as u16) as usize,
                PCWSTR(label.as_ptr()),
            )
        };
    }

    // One of the ways is the setting, so one of them carries the radio mark; a way the menu does
    // not list is marked by nothing rather than by the wrong one.
    if let Some(index) = AUDIO_SEEK_CHOICES.iter().position(|way| *way == seek) {
        let _ = unsafe {
            CheckMenuRadioItem(
                menu,
                base as u32,
                (base + AUDIO_SEEK_CHOICES.len() as u16 - 1) as u32,
                (base + index as u16) as u32,
                MF_BYCOMMAND.0,
            )
        };
    }

    let _ = unsafe {
        AppendMenuW(
            parent,
            MF_STRING | MF_POPUP,
            menu.0 as usize,
            title,
        )
    };
}

/// What a way of starting a sound is called in a menu: the words the tray
/// lists it under, with the way the setting that starts at `default` marked
/// as the default.
///
/// The three that are not a memory are worded as where the sound comes *from* rather than as
/// where it is — `From the Start` rather than `At Start` — because each one answers the question
/// the submenu's own name asks, which is where a hover drops the needle; `Remember` answers it
/// the fourth way, and is left as the one word a user already knows from every player they have
/// used.
fn seek_label(seek: AudioSeek, default: AudioSeek) -> String {
    let label = match seek {
        AudioSeek::Remember => "Remember",
        AudioSeek::Start => "From the Start",
        AudioSeek::Middle => "From the Middle",
        AudioSeek::Random => "Random",
    };

    default_label(label, seek == default)
}

/// What a way of starting a sound is called in the menu: the words the tray lists it under, with
/// the way this setting starts at marked as the default.
pub(super) fn audio_seek_label(seek: AudioSeek) -> String {
    seek_label(seek, DEFAULT_AUDIO_SEEK)
}

/// What a way of starting a *pinned* sound is called in the menu:
/// the words the tray lists it under, with the way *this* setting
/// starts at marked as the default — which is not the way the
/// hover's setting starts, so the mark is the pin's own and not
/// the one `audio_seek_label` puts on the same words.
pub(super) fn pin_mode_audio_seek_label(seek: AudioSeek) -> String {
    seek_label(seek, DEFAULT_PIN_MODE_AUDIO_SEEK)
}

/// The `Video → Loop` / `Audio → Loop` row: the one row that says whether the
/// hovered film or sound goes round at the end of its file, carrying its check. It is a row of
/// the kind's own submenu rather than a submenu of its own — `Loop` is the one answer this app
/// offers for a hover's end — and it takes the id it is handed rather than a range (see
/// `ID_TRAY_VIDEO_LOOP`).
pub(super) fn append_loop_item(menu: HMENU, loop_id: u16, looping: bool, default: bool) {
    append_labeled_item(
        menu,
        MF_STRING | if looping { MF_CHECKED } else { MF_UNCHECKED },
        loop_id,
        &default_label("Loop", default),
    );
}

/// The `Video → Pin Mode → Pause Minimized` / `Audio → Pin Mode → Pause
/// Minimized` row: whether a pin the minimize button put away holds the video or the sound it
/// is playing where it is, carrying its check. It moved here from the `Pin Mode → Minimize →
/// Pause` submenu, one row apiece, because what should go quiet while a pin is away is a
/// question about one kind of file — a film a user wants to go on hearing while a pin is away
/// is not a reason to let a podcast play on — so each kind's own `Pin Mode` answers it for
/// itself (see `toggle_pin_pause_audio`).
pub(super) fn append_pause_minimized_item(menu: HMENU, pause_id: u16, pause: bool) {
    append_labeled_item(
        menu,
        MF_STRING | if pause { MF_CHECKED } else { MF_UNCHECKED },
        pause_id,
        "Pause Minimized",
    );
}

/// The `Video → Pin Mode → Shuffle`, `Audio → Pin Mode → Shuffle` and `Image → Shuffle` rows:
/// whether the next file a *pinned* preview steps on to — the end of a file its `End of
/// File` row calls `Jump`, a caption's next press, a forward key — is a random file of the
/// same kind rather than the one beside it, carrying its check. A hover has no next file,
/// so the row is the pin's own, and a step back is the file before it either way (see
/// `set_pin_shuffle`).
pub(super) fn append_shuffle_item(menu: HMENU, shuffle_id: u16, shuffle: bool, default: bool) {
    append_labeled_item(
        menu,
        MF_STRING | if shuffle { MF_CHECKED } else { MF_UNCHECKED },
        shuffle_id,
        &default_label("Shuffle", default),
    );
}

/// The `Video → Pin Mode → End of File` / `Audio → Pin Mode → End of File`
/// submenu: the three things a pinned film or sound can do at the end of its file, with the
/// one in force carrying its mark.
///
/// The three are one radio group rather than a switch over a list: exactly one of them is in
/// force, and they are the three answers there are — the file left where it ended (`Pause`),
/// the same file again (`Loop`), and the walk's next one (`Jump`, which the kind's own
/// `Shuffle` row decides between the file beside it and a random one of the same kind). One
/// builder serves both kinds because the two are the same question asked twice — once of the
/// film's end and once of the sound's — and each range is what says which setting the click
/// was meant for.
pub(super) fn append_end_of_file_menu(
    parent: HMENU,
    base: u16,
    end_of_file: PinEndOfFile,
    default: PinEndOfFile,
) {
    let menu = unsafe { CreatePopupMenu().unwrap() };

    let labels: Vec<Vec<u16>> = PIN_END_OF_FILE_CHOICES
        .iter()
        .map(|choice| {
            end_of_file_label(*choice, default)
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();
    for (index, label) in labels.iter().enumerate() {
        let _ = unsafe {
            AppendMenuW(
                menu,
                MF_STRING,
                (base + index as u16) as usize,
                PCWSTR(label.as_ptr()),
            )
        };
    }
    if let Some(index) = PIN_END_OF_FILE_CHOICES
        .iter()
        .position(|choice| *choice == end_of_file)
    {
        let _ = unsafe {
            CheckMenuRadioItem(
                menu,
                base as u32,
                (base + PIN_END_OF_FILE_CHOICES.len() as u16 - 1) as u32,
                (base + index as u16) as u32,
                MF_BYCOMMAND.0,
            )
        };
    }

    let _ = unsafe {
        AppendMenuW(
            parent,
            MF_STRING | MF_POPUP,
            menu.0 as usize,
            w!("End of File"),
        )
    };
}

/// The `Video → Pin Mode` / `Audio → Pin Mode` submenu: the rows that are a
/// *pinned* preview's own — whether it holds what it is playing while it is minimized,
/// whether its forward steps shuffle, and what its end of file does — with the sound's own
/// seek row hung under it last. One builder serves both kinds because the two ask the same
/// questions of a film and of a sound (see `append_pause_minimized_item`,
/// `append_shuffle_item`, `append_end_of_file_menu` and `append_pin_seek_menu_renamed`).
///
/// The menu is handed back because the sound hangs one more row inside it: the caller that
/// asked for the submenu is the one that knows whether there is a seek to put in it.
#[allow(clippy::too_many_arguments)] // Each one is a row of the submenu, and a struct of them would be a type for one call.
pub(super) fn append_pin_mode_menu(
    parent: HMENU,
    pause_id: u16,
    pause: bool,
    shuffle_id: u16,
    shuffle: bool,
    shuffle_default: bool,
    end_of_file_base: u16,
    end_of_file: PinEndOfFile,
    end_of_file_default: PinEndOfFile,
) -> HMENU {
    let menu = unsafe { CreatePopupMenu().unwrap() };
    append_pause_minimized_item(menu, pause_id, pause);
    append_shuffle_item(menu, shuffle_id, shuffle, shuffle_default);
    append_end_of_file_menu(menu, end_of_file_base, end_of_file, end_of_file_default);
    let _ = unsafe {
        AppendMenuW(
            parent,
            MF_STRING | MF_POPUP,
            menu.0 as usize,
            w!("Pin Mode"),
        )
    };
    menu
}

/// What an end of file is called in the menu: the words the tray lists it under, with the way
/// the setting starts at marked as the default.
pub(super) fn end_of_file_label(end_of_file: PinEndOfFile, default: PinEndOfFile) -> String {
    let label = match end_of_file {
        PinEndOfFile::Pause => "Pause",
        PinEndOfFile::Loop => "Loop",
        PinEndOfFile::Continue => "Jump",
    };

    default_label(label, end_of_file == default)
}

/// The end of file an item of an `End of File` submenu stands for, by the position it was
/// listed at. An id past the last choice the menu offered is one that is not there.
pub(super) fn end_of_file_at(index: u16) -> Option<PinEndOfFile> {
    PIN_END_OF_FILE_CHOICES.get(index as usize).copied()
}

/// The `Explorer Poll` submenu — the loop's own tick: how often the
/// app looks at the pointer's world while Explorer has focus, with the
/// tick it is at marked.
pub(super) fn append_tick_menu(parent: HMENU, label: PCWSTR, base: u16, tick_ms: u64) {
    let menu = unsafe { CreatePopupMenu().unwrap() };

    // The labels are kept for as long as the menu is being filled out, for the same
    // reason the `Avoid` labels are: `AppendMenuW` is handed a pointer, so the wide
    // strings have to outlive the call that lists them.
    let labels: Vec<Vec<u16>> = TICK_CHOICES_MS
        .iter()
        .map(|tick| {
            tick_label(*tick)
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();

    for (index, label) in labels.iter().enumerate() {
        let _ = unsafe {
            AppendMenuW(
                menu,
                MF_STRING,
                (base + index as u16) as usize,
                PCWSTR(label.as_ptr()),
            )
        };
    }

    // One of the rates is the setting, so one of them carries the radio mark; a tick the
    // menu does not list is marked by nothing rather than by the nearest one.
    if let Some(index) = TICK_CHOICES_MS.iter().position(|tick| *tick == tick_ms) {
        let _ = unsafe {
            CheckMenuRadioItem(
                menu,
                base as u32,
                (base + TICK_CHOICES_MS.len() as u16 - 1) as u32,
                (base + index as u16) as u32,
                MF_BYCOMMAND.0,
            )
        };
    }

    let _ = unsafe { AppendMenuW(parent, MF_STRING | MF_POPUP, menu.0 as usize, label) };
}

/// What a tick is called in the menu: the wait it stands for, in milliseconds, with the
/// one the app starts at marked as the default.
fn tick_label(tick_ms: u64) -> String {
    default_label(&format!("{tick_ms} ms"), tick_ms == DEFAULT_TICK_MS)
}

/// One half of the `Background` submenu: the backdrops a preview can be drawn over,
/// with the one that half is on marked and the one it starts at marked as the default.
/// The halves list the same choices, which is why one builder is handed the base of the
/// ids, the choices and the backdrop to mark rather than the items themselves — the
/// texture's half excepted, which lists two of the four (see `DDS_BACKGROUND_CHOICES`).
pub(super) fn append_background_menu(
    parent: HMENU,
    label: PCWSTR,
    base: u16,
    choices: &[TransparentBackground],
    background: TransparentBackground,
    default: TransparentBackground,
) {
    let menu = unsafe { CreatePopupMenu().unwrap() };

    // The labels are kept for as long as the menu is being filled out, for the same
    // reason the cache labels are: `AppendMenuW` is handed a pointer, so the wide
    // strings have to outlive the call that lists them.
    let labels: Vec<Vec<u16>> = choices
        .iter()
        .map(|choice| {
            background_label(*choice, default)
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();

    for (index, choice) in choices.iter().enumerate() {
        let flags = MF_STRING
            | if *choice == background {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            };
        let _ = unsafe {
            AppendMenuW(
                menu,
                flags,
                (base + index as u16) as usize,
                PCWSTR(labels[index].as_ptr()),
            )
        };
    }

    let _ = unsafe { AppendMenuW(parent, MF_STRING | MF_POPUP, menu.0 as usize, label) };
}

/// What a backdrop is called in the menu: the spelling `config.ini` uses, capitalized, with
/// the one the setting starts at marked as the default. Which of the four a setting starts
/// at is the setting's own — a picture's is the squares, a specimen's is white — which is why
/// the one to mark is handed in rather than written into the words.
pub(super) fn background_label(
    background: TransparentBackground,
    default: TransparentBackground,
) -> String {
    let label = match background {
        TransparentBackground::Transparent => "Transparent",
        TransparentBackground::Black => "Black",
        TransparentBackground::White => "White",
        TransparentBackground::Checkerboard => "Checkerboard",
    };

    default_label(label, background == default)
}

/// The backdrop an item of the `Background` submenu stands for, by the position it
/// was listed at. An id past the last choice the menu offered is one that is not
/// there.
pub(super) fn background_at(index: u16) -> Option<TransparentBackground> {
    BACKGROUND_CHOICES.get(index as usize).copied()
}

/// And the same for an item of the texture's half, which is the one half that offers two
/// backdrops of the four rather than all of them.
pub(super) fn dds_background_at(index: u16) -> Option<TransparentBackground> {
    DDS_BACKGROUND_CHOICES.get(index as usize).copied()
}

/// And the same for an item of a page's half, which is the one half that offers three
/// backdrops of the four — everything but the transparency a page is not drawn over.
pub(super) fn html_background_at(index: u16) -> Option<TransparentBackground> {
    HTML_BACKGROUND_CHOICES.get(index as usize).copied()
}

/// One of the display-share submenus — `Vector`, `Text`, `Ebook`,
/// `Document`, `Font` and `Design`: the shares of the display a
/// document is drawn at, with the one the setting is on marked, and
/// nothing marked for a share the menu does not offer — which is what
/// a hand-edited `config.ini` can ask for. `default` says which of the
/// shares this setting starts at, so the one it names is the one the
/// label marks.
///
/// Each opens with the `By Screen` row between separators, because
/// the first thing a reader of the submenu is told is what the
/// shares below are of.
pub(super) fn append_document_scale_menu(
    parent: HMENU,
    label: PCWSTR,
    base: u16,
    scale: PreviewScale,
    default: PreviewScale,
) {
    let menu = unsafe { CreatePopupMenu().unwrap() };

    append_by_screen_info_row(menu);

    // The labels are kept for as long as the menu is being filled out, for the same
    // reason the cache labels are: `AppendMenuW` is handed a pointer, so the wide
    // strings have to outlive the call that lists them.
    let labels: Vec<Vec<u16>> = DOCUMENT_SCALE_CHOICES
        .iter()
        .map(|choice| {
            document_scale_label(*choice, default)
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();

    for (index, choice) in DOCUMENT_SCALE_CHOICES.iter().enumerate() {
        let flags = MF_STRING
            | if *choice == scale {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            };
        let _ = unsafe {
            AppendMenuW(
                menu,
                flags,
                (base + index as u16) as usize,
                PCWSTR(labels[index].as_ptr()),
            )
        };
    }

    let _ = unsafe { AppendMenuW(parent, MF_STRING | MF_POPUP, menu.0 as usize, label) };
}

/// The `By Screen` row a display-share submenu opens with, with a separator
/// above it and another under it: what the shares below are of, rather than a
/// share itself. The row names a group rather than offering anything, so it is
/// greyed and given no id at all — a click on it lands on command 0, which the
/// window proc ignores (see `append_codecs_menu` for the same id-less treatment
/// the other way round, where an inert row is left normal-looking).
fn append_by_screen_info_row(menu: HMENU) {
    let _ = unsafe { AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null()) };

    append_labeled_item(menu, MF_STRING | MF_GRAYED, 0, "By Screen");

    let _ = unsafe { AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null()) };
}

/// The `By Own Size` row a bitmap's submenu puts its own-size shares
/// under, with a separator above it and another below: the shares
/// that follow are of a bitmap's own size rather than of the display,
/// and the row names the basis the same way `By Screen` names the one
/// above. It is greyed and carries no command either, for the same
/// reason (see `append_by_screen_info_row`).
fn append_by_own_size_info_row(menu: HMENU) {
    let _ = unsafe { AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null()) };

    append_labeled_item(menu, MF_STRING | MF_GRAYED, 0, "By Own Size");

    let _ = unsafe { AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null()) };
}

/// The `Image`, `Video` and `Animated` submenus: the shares a
/// bitmap — a picture, a video's first frame and the player window
/// over it, or a frame that moves — is drawn at, with the one the
/// setting is on marked and nothing marked for a share the menu does
/// not offer, which is what a hand-edited `config.ini` can ask for.
/// One builder serves all three because the three are the same
/// question asked of three kinds of file: the shares are one list,
/// the labels are one function, and what differs is only which
/// setting the submenu writes and the id its items carry.
///
/// The shares are of two bases, and the submenu says which is which
/// around the items themselves: the shares of the display's fitted
/// size open it, under a `By Screen` row, and the shares of a
/// bitmap's own size close it, under a `By Own Size` row, each row
/// bracketed by separators.
pub(super) fn append_bitmap_scale_menu(
    parent: HMENU,
    label: PCWSTR,
    base: u16,
    scale: PreviewScale,
    default: PreviewScale,
) {
    let menu = unsafe { CreatePopupMenu().unwrap() };

    append_by_screen_info_row(menu);

    // Where the own-size group begins: the first share of a size of
    // the file's own, which is the boundary the `By Own Size` row and
    // the separators around it stand at.
    let own_size_begin = BITMAP_SCALE_CHOICES
        .iter()
        .position(|choice| matches!(choice, PreviewScale::Percent(_)))
        .unwrap_or(BITMAP_SCALE_CHOICES.len());

    // The labels are kept for as long as the menu is being filled out, for the same
    // reason the cache labels are: `AppendMenuW` is handed a pointer, so the wide
    // strings have to outlive the call that lists them.
    let labels: Vec<Vec<u16>> = BITMAP_SCALE_CHOICES
        .iter()
        .map(|choice| {
            bitmap_scale_label(*choice, default)
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();

    for (index, label) in labels.iter().enumerate() {
        if index == own_size_begin {
            append_by_own_size_info_row(menu);
        }

        let _ = unsafe {
            AppendMenuW(
                menu,
                MF_STRING,
                (base + index as u16) as usize,
                PCWSTR(label.as_ptr()),
            )
        };
    }

    // One of the shares is the setting, so one of them carries the radio mark; a
    // share the menu does not list is marked by nothing rather than by the wrong
    // one. The mark is over the ids the items carry, which are one run of the
    // base however the two groups and their rows are interleaved.
    if let Some(index) = BITMAP_SCALE_CHOICES.iter().position(|choice| *choice == scale) {
        let _ = unsafe {
            CheckMenuRadioItem(
                menu,
                base as u32,
                (base + BITMAP_SCALE_CHOICES.len() as u16 - 1) as u32,
                (base + index as u16) as u32,
                MF_BYCOMMAND.0,
            )
        };
    }

    let _ = unsafe { AppendMenuW(parent, MF_STRING | MF_POPUP, menu.0 as usize, label) };
}

/// The `Audio` submenu under `Scaling`: the shares of the display a sound's card is
/// laid out over, with the one the setting is on marked, and nothing marked
/// for a share the menu does not offer — which is what a hand-edited
/// `config.ini` can ask for. The shares are of the display rather than of a
/// size of the file's own, because a card holds nothing of the sound to take
/// a share of — what it holds is laid out over the room it is given — so the
/// labels are the document scale's, which name the same question, and the
/// submenu opens with the `By Screen` row, between separators, like those do.
pub(super) fn append_audio_scale_menu(
    parent: HMENU,
    label: PCWSTR,
    base: u16,
    scale: PreviewScale,
    default: PreviewScale,
) {
    let menu = unsafe { CreatePopupMenu().unwrap() };

    append_by_screen_info_row(menu);

    // The labels are kept for as long as the menu is being filled out, for the same
    // reason the cache labels are: `AppendMenuW` is handed a pointer, so the wide
    // strings have to outlive the call that lists them.
    let labels: Vec<Vec<u16>> = AUDIO_SCALE_CHOICES
        .iter()
        .map(|choice| {
            document_scale_label(*choice, default)
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();

    for (index, choice) in AUDIO_SCALE_CHOICES.iter().enumerate() {
        let flags = MF_STRING
            | if *choice == scale {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            };
        let _ = unsafe {
            AppendMenuW(
                menu,
                flags,
                (base + index as u16) as usize,
                PCWSTR(labels[index].as_ptr()),
            )
        };
    }

    let _ = unsafe { AppendMenuW(parent, MF_STRING | MF_POPUP, menu.0 as usize, label) };
}

/// What a share a bitmap is drawn at is called in the menu: the
/// percentage itself, with the one the setting starts at marked as
/// the default. Both bases name a share by the same number — the
/// `By Screen` and `By Own Size` rows say which basis it is of — and
/// taking the whole of either is not a percentage, so it is named
/// for what it does.
pub(super) fn bitmap_scale_label(scale: PreviewScale, default: PreviewScale) -> String {
    let label = match scale {
        PreviewScale::FitToScreenReduced(percent) | PreviewScale::Percent(percent) => {
            format!("{percent}%")
        }
        _ => "Fit to Screen".to_string(),
    };

    default_label(&label, scale == default)
}

/// What a share of the display is called in the menu: the percentage itself, with the
/// one the setting starts at marked as the default. The whole room is not a percentage
/// of it, so it is named for what it is.
pub(super) fn document_scale_label(scale: PreviewScale, default: PreviewScale) -> String {
    let label = match scale {
        PreviewScale::Percent(percent) => format!("{percent}%"),
        _ => "Fit to Screen".to_string(),
    };

    default_label(&label, scale == default)
}

/// The share of the display an item of a `… Scaling` submenu stands for, by the
/// position it was listed at. An id past the last choice the menu offered is one that
/// is not there.
pub(super) fn document_scale_at(index: u16) -> Option<PreviewScale> {
    DOCUMENT_SCALE_CHOICES.get(index as usize).copied()
}

/// The share of the display an item of the `Audio` submenu stands
/// for, by the position it was listed at — the share a sound's card is laid
/// out over. An id past the last choice the menu offered is one that is not
/// there.
pub(super) fn audio_scale_at(index: u16) -> Option<PreviewScale> {
    AUDIO_SCALE_CHOICES.get(index as usize).copied()
}

/// The `Engine → Video Engine` submenu: the `Fallback` switch at the
/// top, then the engines it decides between. An engine this machine
/// does not have is greyed, since there is nothing there to choose;
/// `Best` is the default and carries the mark.
pub(super) fn append_video_engine_menu(parent: HMENU) {
    let (selected, fallback) = CONFIG
        .lock()
        .map(|config| (config.video_engine, config.video_engine_fallback))
        .unwrap_or((DEFAULT_VIDEO_ENGINE, DEFAULT_VIDEO_ENGINE_FALLBACK));

    let menu = unsafe { CreatePopupMenu().unwrap() };

    append_labeled_item(
        menu,
        MF_STRING | if fallback { MF_CHECKED } else { MF_UNCHECKED },
        ID_TRAY_VIDEO_ENGINE_FALLBACK,
        "Fallback",
    );
    let _ = unsafe { AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null()) };

    for (index, engine) in VIDEO_ENGINE_CHOICES.iter().enumerate() {
        // A row is greyed for an engine this machine has not got: `Best` and `Native` are always
        // here, the FFmpeg-based choices need `ffplay`, and mpv needs `mpv.exe` (see
        // `VideoEngine::installed`).
        let unavailable = !engine.installed();

        append_labeled_item(
            menu,
            MF_STRING | if unavailable { MF_GRAYED } else { MF_ENABLED },
            ID_TRAY_VIDEO_ENGINE_BASE + index as u16,
            &video_engine_label(*engine),
        );
    }

    if let Some(index) = VIDEO_ENGINE_CHOICES.iter().position(|e| *e == selected) {
        let _ = unsafe {
            CheckMenuRadioItem(
                menu,
                ID_TRAY_VIDEO_ENGINE_BASE as u32,
                (ID_TRAY_VIDEO_ENGINE_BASE + VIDEO_ENGINE_CHOICES.len() as u16 - 1) as u32,
                (ID_TRAY_VIDEO_ENGINE_BASE + index as u16) as u32,
                MF_BYCOMMAND.0,
            )
        };
    }

    let _ = unsafe { AppendMenuW(parent, MF_STRING | MF_POPUP, menu.0 as usize, w!("Video Engine")) };
}

/// What an engine row is called: the engine's name, with the one the app starts at marked.
pub(super) fn video_engine_label(engine: VideoEngine) -> String {
    let label = match engine {
        VideoEngine::Best => "Best",
        VideoEngine::Mpv => "MPV",
        VideoEngine::Native => "Native",
        VideoEngine::Ffmpeg => "FFmpeg",
        VideoEngine::Hybrid => "Native (FFmpeg above 3.2MP)",
    };
    default_label(label, engine == DEFAULT_VIDEO_ENGINE)
}

/// The `Engine → Office Engine` submenu: which engine an Office document's
/// page is asked of, with the setting's own choice marked.
///
/// There are two engines to ask and both are listed. The row for one this machine
/// has not got is greyed out rather than left out: what it names cannot be
/// started, so the app would fall back to the other engine for as long as that
/// is so — but the choice is the user's, it is remembered where it is made,
/// and the day the engine is installed it is the one that draws
/// (see `office_formats::page_engine`).
pub(super) fn append_office_engine_menu(parent: HMENU) {
    let selected = CONFIG
        .lock()
        .map(|config| config.office_engine)
        .unwrap_or(DEFAULT_OFFICE_ENGINE);

    let office_menu = unsafe { CreatePopupMenu().unwrap() };

    let microsoft_office = default_label(
        "Microsoft Office",
        DEFAULT_OFFICE_ENGINE == OfficeEngine::MicrosoftOffice,
    );
    append_labeled_item(
        office_menu,
        MF_STRING,
        ID_TRAY_ENGINE_OFFICE_MS,
        &microsoft_office,
    );

    // The render engine's row is the one that can be greyed: a document is drawn without
    // either application on a machine that has no Office at all — the other engine is the
    // fallback for every family whose application is missing — while a machine with no
    // LibreOffice has nothing to ask, whatever the setting says.
    let libre_flags = if libreoffice_render::available() {
        MF_STRING
    } else {
        MF_STRING | MF_GRAYED
    };
    append_labeled_item(
        office_menu,
        libre_flags,
        ID_TRAY_ENGINE_OFFICE_LIBRE,
        "LibreOffice",
    );

    let _ = unsafe {
        CheckMenuRadioItem(
            office_menu,
            ID_TRAY_ENGINE_OFFICE_MS as u32,
            ID_TRAY_ENGINE_OFFICE_LIBRE as u32,
            match selected {
                OfficeEngine::MicrosoftOffice => ID_TRAY_ENGINE_OFFICE_MS as u32,
                OfficeEngine::LibreOffice => ID_TRAY_ENGINE_OFFICE_LIBRE as u32,
            },
            MF_BYCOMMAND.0,
        )
    };

    let _ = unsafe {
        AppendMenuW(
            parent,
            MF_STRING | MF_POPUP,
            office_menu.0 as usize,
            w!("Office Engine"),
        )
    };
}

/// The `Away Timer` submenu: how long Explorer may be out of reach before an
/// engine that is not marked `Persistent` is let go. The setting it asks
/// about is `afk_timer_seconds`, which keeps its name.
///
/// It is one setting for every engine this app keeps, because the question it answers is
/// about the user rather than about an engine: nothing this app holds is being looked at
/// while no Explorer window is reachable, and which of those engines is worth holding until
/// the user comes back is what the `Persistent` toggles below are for. It is listed longest
/// first, like every other menu of times here, and it marks the time it is on: a value a
/// hand-edited `config.ini` asks for that is not one of these is shown with nothing marked
/// rather than rounded to the nearest.
pub(super) fn append_afk_timer_menu(parent: HMENU) {
    let seconds = CONFIG
        .lock()
        .map(|config| config.afk_timer_seconds)
        .unwrap_or(DEFAULT_AFK_TIMER_SECS);

    let menu = unsafe { CreatePopupMenu().unwrap() };

    // The labels are kept for as long as the menu is being filled out, for the reason the
    // idle times' are: `AppendMenuW` is handed a pointer, so the wide strings have to
    // outlive the call that lists them.
    let labels: Vec<Vec<u16>> = AFK_TIMER_CHOICES_SECS
        .iter()
        .map(|choice| {
            afk_timer_label(*choice)
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();

    for (index, label) in labels.iter().enumerate() {
        let _ = unsafe {
            AppendMenuW(
                menu,
                MF_STRING,
                (ID_TRAY_AFK_TIMER_BASE + index as u16) as usize,
                PCWSTR(label.as_ptr()),
            )
        };
    }

    if let Some(index) = AFK_TIMER_CHOICES_SECS
        .iter()
        .position(|choice| *choice == seconds)
    {
        let _ = unsafe {
            CheckMenuRadioItem(
                menu,
                ID_TRAY_AFK_TIMER_BASE as u32,
                (ID_TRAY_AFK_TIMER_BASE + AFK_TIMER_CHOICES_SECS.len() as u16 - 1) as u32,
                (ID_TRAY_AFK_TIMER_BASE + index as u16) as u32,
                MF_BYCOMMAND.0,
            )
        };
    }

    let _ = unsafe {
        AppendMenuW(
            parent,
            MF_STRING | MF_POPUP,
            menu.0 as usize,
            w!("Away Timer"),
        )
    };
}

/// One `… Engine TTL` submenu: how far it may go, as the `Persistent` toggle at the top of
/// it and the idle times below that, with the time the engine is on marked and nothing
/// marked for a time the menu does not offer — which is what a hand-edited `config.ini` can
/// ask for.
///
/// The toggle is what decides what the times mean. On, they are the whole of how long the
/// engine is kept, whatever the user is doing; off, `Engine → Away Timer` is what
/// bounds it
/// and the times are not consulted at all, so an engine is kept while Explorer is in front
/// of the user and let go once it has not been for that long. The toggle is at the top and
/// the times below a separator because it is a different kind of answer — a checkmark rather
/// than one of the times — and because nothing below it applies until it is on.
///
/// An engine that is not on the machine at all is greyed out, since there is nothing there
/// to keep.
pub(super) fn append_engine_idle_menu(
    parent: HMENU,
    label: PCWSTR,
    ids: EngineIdleIds,
    idle: EngineIdle,
    persistent: bool,
    default_seconds: u64,
    enabled: bool,
) {
    let menu = unsafe { CreatePopupMenu().unwrap() };

    let persistent_flags = MF_STRING | if persistent { MF_CHECKED } else { MF_UNCHECKED };
    let _ = unsafe {
        AppendMenuW(
            menu,
            persistent_flags,
            ids.persistent as usize,
            w!("Persistent"),
        )
    };
    let _ = unsafe { AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null()) };

    // The labels are kept for as long as the menu is being filled out, for the same
    // reason the cache labels are: `AppendMenuW` is handed a pointer, so the wide
    // strings have to outlive the call that lists them.
    let labels: Vec<Vec<u16>> = ENGINE_IDLE_CHOICES
        .iter()
        .map(|choice| {
            engine_idle_label(*choice, default_seconds)
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();

    for (index, label) in labels.iter().enumerate() {
        let _ = unsafe {
            AppendMenuW(
                menu,
                MF_STRING,
                (ids.times + index as u16) as usize,
                PCWSTR(label.as_ptr()),
            )
        };
    }

    if let Some(index) = ENGINE_IDLE_CHOICES
        .iter()
        .position(|choice| *choice == idle)
    {
        let _ = unsafe {
            CheckMenuRadioItem(
                menu,
                ids.times as u32,
                (ids.times + ENGINE_IDLE_CHOICES.len() as u16 - 1) as u32,
                (ids.times + index as u16) as u32,
                MF_BYCOMMAND.0,
            )
        };
    }

    let flags = if enabled {
        MF_STRING | MF_POPUP
    } else {
        MF_STRING | MF_POPUP | MF_GRAYED
    };

    let _ = unsafe { AppendMenuW(parent, flags, menu.0 as usize, label) };
}

/// The `Codecs` submenu: what this machine has of everything a preview leans on, grouped
/// by what each thing is for.
///
/// A row that is there is marked and reads as any other item does, and picking it does
/// nothing — there is nothing to do about a thing the machine already has. A row that is
/// not there is greyed and cannot be picked at all where there is nowhere to send anyone,
/// and is pickable where the README names a page for it: picking one asks whether to open
/// that page, which is the whole of what this submenu does. The app installs nothing and
/// fetches nothing for it — what a yes hands over is a link, and the browser the user
/// already has is what opens it (see `open_codec_page`).
///
/// A Windows item cannot be both normal-looking and unpickable, so the rows that are
/// present are made inert by having no id at all rather than by being disabled, and the
/// mark they carry is a glyph rather than the checkmark column a menu item can draw —
/// those are the same trade the other way round (see `codec_row_label`).
pub(super) fn append_codecs_menu(menu: HMENU) {
    let codecs_menu = unsafe { CreatePopupMenu().unwrap() };

    let video = codecs::video();
    let images = codecs::images();
    let audio = codecs::audio();

    // The groups are numbered as one list, in the order they are read, so that an id says which
    // row was picked without the menu having to be remembered: the same groups are built again,
    // in the same order, when the click arrives (see `codec_row`).
    let audio_base = ID_TRAY_CODEC_BASE + video.len() as u16;
    let images_base = audio_base + audio.len() as u16;
    let engines_base = images_base + images.len() as u16;

    append_codec_group(codecs_menu, w!("Videos"), video, ID_TRAY_CODEC_BASE);
    append_codec_group(codecs_menu, w!("Audio"), audio, audio_base);
    append_codec_group(codecs_menu, w!("Images"), images, images_base);
    append_codec_group(codecs_menu, w!("Engines"), codecs::engines(), engines_base);

    let _ = unsafe {
        AppendMenuW(
            menu,
            MF_STRING | MF_POPUP,
            codecs_menu.0 as usize,
            w!("Codecs"),
        )
    };
}

/// One group of the `Codecs` submenu, with a row per engine or codec in it.
///
/// `base` is where this group's commands begin in the numbering above. The row named `MPV` is
/// the one that can open a submenu of its own (see `is_mpv_row` and `append_mpv_row`).
fn append_codec_group(parent: HMENU, label: PCWSTR, rows: Vec<Row>, base: u16) {
    let group = unsafe { CreatePopupMenu().unwrap() };

    for (index, row) in rows.iter().enumerate() {
        if is_mpv_row(row) {
            append_mpv_row(group, row, index as u16, base);
        } else {
            append_codec_row(group, row, index as u16, base);
        }
    }

    let _ = unsafe { AppendMenuW(parent, MF_STRING | MF_POPUP, group.0 as usize, label) };
}

/// One row of a codec group: a row the machine has is inert, and so is a row it is missing that
/// has nowhere to be got from — neither is given an id, so picking one closes the menu and does
/// nothing else. The two are told apart by being grey where the second is: a row cannot be both
/// normal-looking and unpickable, and the cross every missing row carries is what says which of
/// them a row is.
fn append_codec_row(group: HMENU, row: &Row, index: u16, base: u16) {
    let pickable = !row.available && row.link.is_some();

    let flags = if pickable || row.available {
        MF_STRING
    } else {
        MF_STRING | MF_GRAYED
    };

    let id = if pickable { base + index } else { 0 };

    append_labeled_item(group, flags, id, &codec_row_label(row));
}

/// Whether a codec row is the one that opens the submenu of mpv installs: the row the tray names
/// `MPV`, matched on the name rather than on the position it holds — the position is a fact of
/// whatever the `Videos` list looks like today, the name is what the row is.
pub(super) fn is_mpv_row(row: &Row) -> bool {
    row.name == "MPV"
}

/// The `MPV` row of a codec group, found by its name rather than by where it is listed: where
/// this machine has more than one install to pick between it opens a submenu of them, and with
/// one, or none, it is the row it has always been (see `is_mpv_row`).
///
/// It is a *normal* popup rather than the inert greyed row an available codec gets, because a
/// greyed popup cannot be opened — the row that offers a choice has to be pickable. The children
/// carry the id range `ID_TRAY_MPV_BASE` hands out, and a click re-enumerates the installs in the
/// same order rather than remembering the menu (see `set_mpv_path`).
fn append_mpv_row(group: HMENU, row: &Row, index: u16, base: u16) {
    let installs = codecs::mpv_installs();

    if installs.len() < 2 {
        append_codec_row(group, row, index, base);
        return;
    }

    let chosen = CONFIG
        .lock()
        .map(|config| config.mpv_path.clone())
        .unwrap_or_default();
    let in_use = codecs::resolve_mpv(
        (!chosen.trim().is_empty()).then_some(Path::new(&chosen)),
        &installs,
    );

    let menu = unsafe { CreatePopupMenu().unwrap() };

    let items = mpv_menu_items(&installs, in_use.as_deref());

    for (index, (label, _)) in items.iter().enumerate() {
        append_labeled_item(menu, MF_STRING, ID_TRAY_MPV_BASE + index as u16, label);
    }

    // One install is in use and the others are not, so one item carries the radio mark — the
    // same shape, and the same `CheckMenuRadioItem` call, as the `Engine → Video Engine` rows
    // (see `append_video_engine_menu`). It clears the others, so a rebuild leaves no stale mark
    // on an install that was in use before.
    if let Some(index) = items.iter().position(|(_, marked)| *marked) {
        let _ = unsafe {
            CheckMenuRadioItem(
                menu,
                ID_TRAY_MPV_BASE as u32,
                (ID_TRAY_MPV_BASE + items.len() as u16 - 1) as u32,
                (ID_TRAY_MPV_BASE + index as u16) as u32,
                MF_BYCOMMAND.0,
            )
        };
    }

    let label: Vec<u16> = codec_row_label(row)
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();

    let _ = unsafe {
        AppendMenuW(
            group,
            MF_STRING | MF_POPUP,
            menu.0 as usize,
            PCWSTR(label.as_ptr()),
        )
    };
}

/// What the `MPV` row's submenu lists: one item per install, in the order they were found, with
/// the one in use marked — and nothing at all where there is no choice to make, which is what
/// leaves the row the row it is today.
///
/// It is pure so the two-install answer is testable without depending on what this machine has
/// installed; the one in use is `resolve_mpv`'s answer rather than a second reading of the
/// setting.
pub(super) fn mpv_menu_items(installs: &[PathBuf], in_use: Option<&Path>) -> Vec<(String, bool)> {
    if installs.len() < 2 {
        return Vec::new();
    }

    installs
        .iter()
        .map(|install| (mpv_row_label(install), Some(install.as_path()) == in_use))
        .collect()
}

/// The most characters an `MPV` row's label may take, the house ellipsis and the separator after
/// it included. A path that does not fit is shortened from the front, whole components kept from
/// the end so no folder is cut in half (`…\SVP 4\mpv64\mpv.exe`).
///
/// The tray has no measuring helper — the pin caption's `longest_prefix_that_fits` measures
/// against a DC — so this is a count of characters: upgrade to measuring if a menu label ever
/// needs one.
pub(super) const MPV_LABEL_BUDGET: usize = 21;

/// What one install reads as in the `MPV` row's submenu: the path as it stands under the budget,
/// and the tail of it after the house ellipsis past it.
pub(super) fn mpv_row_label(path: &Path) -> String {
    let shown = path.to_string_lossy();

    if shown.chars().count() <= MPV_LABEL_BUDGET {
        return shown.into_owned();
    }

    // The `…` and the separator after it take two of the budget; the rest is the tail.
    let room = MPV_LABEL_BUDGET.saturating_sub(2);
    let mut kept: Vec<&str> = Vec::new();
    let mut width = 0;

    for part in shown.split(['\\', '/']).rev() {
        // One character more where the separator before a component has to be kept with it.
        let cost = part.chars().count() + usize::from(!kept.is_empty());

        if width + cost > room && !kept.is_empty() {
            break;
        }

        width += cost;
        kept.push(part);
    }

    let tail: Vec<&str> = kept.into_iter().rev().collect();

    format!("\u{2026}\\{}", tail.join("\\"))
}

/// What one row of the `Codecs` submenu is written as: the mark that says whether it is
/// there, and the name of the engine or format.
///
/// The mark is a glyph in the label rather than the checkmark column a menu item can carry,
/// because the two cannot be told apart for the rows that matter: a menu greys a checked
/// item along with everything else about it, so a present row and a missing one would look
/// the same. A glyph is nothing but text, and it stays legible on the row it is on.
///
/// A row that can be picked carries an ellipsis, which is the Windows way of saying a dialog
/// follows — and one does: the question of whether to open the page the thing is got from.
fn codec_row_label(row: &Row) -> String {
    if row.available {
        format!("\u{2714} {}", row.name)
    } else if row.link.is_some() {
        format!("\u{2716} {}\u{2026}", row.name)
    } else {
        format!("\u{2716} {}", row.name)
    }
}

/// The click on a row of the `Codecs` submenu: ask whether to open the page the thing is
/// installed from, and open it where the answer is yes.
///
/// The row is read again from the machine rather than kept from the menu build, which is
/// what makes an id mean the same thing on both sides: the groups are listed in a fixed
/// order, so the position an id stands for is the position the row is at now. The link is
/// asked of that row for the same reason — a `.heic` arrives in two packages, and which of
/// the two is missing is a fact of this moment rather than of the menu that was built.
pub(super) fn open_codec_page(index: u16) {
    let Some(row) = codec_row(index as usize) else {
        return;
    };

    // A row the machine has offers nothing, and one that was missing when the menu was built
    // and is not anymore is answered the same way: what the row says now is what it is
    // answered with, rather than what it said when the menu was drawn.
    if row.available {
        return;
    }

    let Some(url) = row.link else {
        return;
    };

    if dialogs::confirm_open_page(row.name, url) {
        open_link(url);
    }
}

/// The row the `Codecs` submenu lists at `index`: the groups read as one list, in the order
/// they are listed, which is the order their commands were handed out in.
fn codec_row(index: usize) -> Option<Row> {
    let mut rows = codecs::video();
    rows.extend(codecs::audio());
    rows.extend(codecs::images());
    rows.extend(codecs::engines());

    rows.into_iter().nth(index)
}

/// What an idle time is called in an `… Engine TTL` submenu: the time, with the one an
/// engine is kept for by default marked. The shortest is the only one that is not a
/// whole minute, and the longest is the only one that is not a whole number of
/// minutes.
pub(super) fn engine_idle_label(idle: EngineIdle, default_seconds: u64) -> String {
    let label = match idle {
        EngineIdle::Indefinite => "Indefinitely".to_string(),
        EngineIdle::Seconds(0) => "0 seconds".to_string(),
        EngineIdle::Seconds(60) => "1 minute".to_string(),
        EngineIdle::Seconds(3600) => "1 hour".to_string(),
        EngineIdle::Seconds(seconds) => format!("{} minutes", seconds / 60),
    };

    default_label(&label, idle == EngineIdle::Seconds(default_seconds))
}

/// The idle time an item of an `… Engine TTL` submenu stands for, by the position it
/// was listed at. An id past the last time the menu offered is one that is not there.
pub(super) fn engine_idle_at(index: u16) -> Option<EngineIdle> {
    ENGINE_IDLE_CHOICES.get(index as usize).copied()
}

/// What an away time is called in the `Away Timer` submenu: the time, with
/// the one an engine
/// that is not `Persistent` is bounded by marked as the default.
///
/// Three of the seven are not a whole number of minutes, and each says itself which one it
/// is — a half of a minute and a quarter of one are what the bottom of the list is for, so
/// they are read as seconds rather than as `0 minutes`.
fn afk_timer_label(seconds: u64) -> String {
    let label = match seconds {
        3600 => "1 hour".to_string(),
        60 => "1 minute".to_string(),
        seconds if seconds < 60 => format!("{seconds} seconds"),
        seconds => format!("{} minutes", seconds / 60),
    };

    default_label(&label, seconds == DEFAULT_AFK_TIMER_SECS)
}

/// The away time an item of the `Away Timer` submenu stands for, by the
/// position it was
/// listed at. An id past the last time the menu offered is one that is not there.
pub(super) fn afk_timer_secs_at(index: u16) -> Option<u64> {
    AFK_TIMER_CHOICES_SECS.get(index as usize).copied()
}

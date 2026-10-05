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
    EngineIdleIds, AFK_TIMER_CHOICES_SECS, AUDIO_SEEK_CHOICES, AVOID_CHOICES, BACKGROUND_CHOICES,
    BITMAP_SCALE_CHOICES, DDS_BACKGROUND_CHOICES, DOCUMENT_SCALE_CHOICES, ENGINE_IDLE_CHOICES,
    HTML_BACKGROUND_CHOICES, ID_TRAY_AFK_TIMER_BASE, ID_TRAY_AUDIO_SEEK_BASE, ID_TRAY_CODEC_BASE,
    ID_TRAY_ENGINE_OFFICE_LIBRE, ID_TRAY_ENGINE_OFFICE_MS, ID_TRAY_VIDEO_ENGINE_BASE,
    ID_TRAY_VIDEO_ENGINE_FALLBACK, TICK_CHOICES_MS, TIMING_DELAY_CHOICES_MS, VIDEO_ENGINE_CHOICES,
};

use crate::app::dialogs;
use crate::config::config::{
    AudioSeek, AvoidMode, EngineIdle, OfficeEngine, PreviewScale, TransparentBackground,
    VideoEngine, DEFAULT_AFK_TIMER_SECS, DEFAULT_AUDIO_SEEK, DEFAULT_AVOID_MODE,
    DEFAULT_DECODE_BUDGET_GB, DEFAULT_OFFICE_ENGINE, DEFAULT_TICK_MS, DEFAULT_VIDEO_ENGINE,
    DEFAULT_VIDEO_ENGINE_FALLBACK,
};
use crate::engines::libreoffice_render;
use crate::formats::codecs::{self, Row};
use crate::ui::preview_window::video_engine_installed;
use crate::CONFIG;
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
pub(super) fn avoid_label(mode: AvoidMode) -> String {
    let label = match mode {
        AvoidMode::Off => "Avoid Nothing",
        AvoidMode::Filename => "Avoid Filename",
        AvoidMode::FilenameColumn => "Avoid Filename Column",
        AvoidMode::Details => "Avoid Details",
    };

    default_label(label, mode == DEFAULT_AVOID_MODE)
}

/// The `Volume → Audio Seek` submenu: where in a file a sound starts playing, with the way the
/// setting is on marked.
///
/// It is listed under `Volume` rather than in a menu of its own because it is the same moment
/// and the same question as the level above it — both are read as a player is started, and a
/// change reaches the next hover rather than the sound on screen — and it is a submenu of its
/// own inside that one because a sound's volume is not a video's and neither is a video's start
/// position: a video is looked at from its beginning and nothing else is offered for one.
pub(super) fn append_audio_seek_menu(parent: HMENU, seek: AudioSeek) {
    let menu = unsafe { CreatePopupMenu().unwrap() };

    // The labels are kept for as long as the menu is being filled out, for the same reason the
    // `Avoid` labels are: `AppendMenuW` is handed a pointer, so the wide strings have to outlive
    // the call that lists them.
    let labels: Vec<Vec<u16>> = AUDIO_SEEK_CHOICES
        .iter()
        .map(|seek| {
            audio_seek_label(*seek)
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
                (ID_TRAY_AUDIO_SEEK_BASE + index as u16) as usize,
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
                ID_TRAY_AUDIO_SEEK_BASE as u32,
                (ID_TRAY_AUDIO_SEEK_BASE + AUDIO_SEEK_CHOICES.len() as u16 - 1) as u32,
                (ID_TRAY_AUDIO_SEEK_BASE + index as u16) as u32,
                MF_BYCOMMAND.0,
            )
        };
    }

    let _ = unsafe {
        AppendMenuW(
            parent,
            MF_STRING | MF_POPUP,
            menu.0 as usize,
            w!("Audio Seek"),
        )
    };
}

/// What a way of starting a sound is called in the menu: the words the tray lists it under, with
/// the way this setting starts at marked as the default.
///
/// The three that are not a memory are worded as where the sound comes *from* rather than as
/// where it is — `From the Start` rather than `At Start` — because each one answers the question
/// the submenu's own name asks, which is where a hover drops the needle; `Remember` answers it
/// the fourth way, and is left as the one word a user already knows from every player they have
/// used.
pub(super) fn audio_seek_label(seek: AudioSeek) -> String {
    let label = match seek {
        AudioSeek::Remember => "Remember",
        AudioSeek::Start => "From the Start",
        AudioSeek::Middle => "From the Middle",
        AudioSeek::Random => "Random",
    };

    default_label(label, seek == DEFAULT_AUDIO_SEEK)
}

/// The `Tick` submenu: how often the app looks at the pointer's world while Explorer
/// has focus, with the tick it is at marked.
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

/// One `… Scaling` submenu: the shares of the display a document is drawn at, with the
/// one the setting is on marked, and nothing marked for a share the menu does not
/// offer — which is what a hand-edited `config.ini` can ask for. `default` says which
/// of the shares this setting starts at, so the one it names is the one the label
/// marks.
pub(super) fn append_document_scale_menu(
    parent: HMENU,
    label: PCWSTR,
    base: u16,
    scale: PreviewScale,
    default: PreviewScale,
) {
    let menu = unsafe { CreatePopupMenu().unwrap() };

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

/// The `Image Scaling` and `Video Scaling` submenus: the shares of its own size a bitmap
/// — a picture, or a video's first frame and the player window over it — is drawn at, with
/// the one the setting is on marked and nothing marked for a share the menu does not
/// offer, which is what a hand-edited `config.ini` can ask for. One builder serves both
/// because the two are the same question asked of two kinds of file: the shares are one
/// list, the labels are one function, and what differs is only which setting the submenu
/// writes and the id its items carry.
pub(super) fn append_bitmap_scale_menu(
    parent: HMENU,
    label: PCWSTR,
    base: u16,
    scale: PreviewScale,
    default: PreviewScale,
) {
    let menu = unsafe { CreatePopupMenu().unwrap() };

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

    for (index, choice) in BITMAP_SCALE_CHOICES.iter().enumerate() {
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

/// What a share of a bitmap's own size is called in the menu: the percentage itself, with
/// the one the setting starts at marked as the default. Taking the whole of it is not a
/// percentage, so it is named for what it does.
pub(super) fn bitmap_scale_label(scale: PreviewScale, default: PreviewScale) -> String {
    let label = match scale {
        PreviewScale::Percent(percent) => format!("{percent}%"),
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

/// The `Engine -> Select Engine -> Video` submenu: the `Fallback` switch at the top, then the
/// engines it decides between. An engine this machine does not have is greyed, since there is
/// nothing there to choose; `Best` is the default and carries the mark.
fn append_video_engine_menu(parent: HMENU) {
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
        // `Best` is never greyed: it is not an engine this machine may be without, it is the
        // app's own answer, and it is what a choice the machine cannot supply falls back to
        // anyway (see `resolve_video_engine`).
        let unavailable = *engine != VideoEngine::Best && !video_engine_installed(*engine);

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

    let _ = unsafe { AppendMenuW(parent, MF_STRING | MF_POPUP, menu.0 as usize, w!("Video")) };
}

/// What an engine row is called: the engine's name, with the one the app starts at marked.
pub(super) fn video_engine_label(engine: VideoEngine) -> String {
    let label = match engine {
        VideoEngine::Best => "Best",
        VideoEngine::Native => "Native",
        VideoEngine::Ffmpeg => "FFmpeg",
    };
    default_label(label, engine == DEFAULT_VIDEO_ENGINE)
}

/// The `Engine → Select Engine → Office` submenu: which engine an Office document's page
/// is asked of, with the setting's own choice marked.
///
/// There are two engines to ask and both are listed. The row for one this machine has not got
/// is greyed out rather than left out: what it names cannot be started, so the app would fall
/// back to the other engine for as long as that is so — but the choice is the user's, it is
/// remembered where it is made, and the day the engine is installed it is the one that draws
/// (see `office_formats::page_engine`).
pub(super) fn append_select_engine_menu(parent: HMENU) {
    let selected = CONFIG
        .lock()
        .map(|config| config.office_engine)
        .unwrap_or(DEFAULT_OFFICE_ENGINE);

    let select_engine_menu = unsafe { CreatePopupMenu().unwrap() };
    let office_menu = unsafe { CreatePopupMenu().unwrap() };

    append_video_engine_menu(select_engine_menu);

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
            select_engine_menu,
            MF_STRING | MF_POPUP,
            office_menu.0 as usize,
            w!("Office"),
        )
    };
    let _ = unsafe {
        AppendMenuW(
            parent,
            MF_STRING | MF_POPUP,
            select_engine_menu.0 as usize,
            w!("Select Engine"),
        )
    };
}

/// The `AFK Timer` submenu: how long Explorer may be out of reach before an engine that is
/// not marked `Persistent` is let go.
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
            w!("AFK Timer"),
        )
    };
}

/// One `… Engine TTL` submenu: how far it may go, as the `Persistent` toggle at the top of
/// it and the idle times below that, with the time the engine is on marked and nothing
/// marked for a time the menu does not offer — which is what a hand-edited `config.ini` can
/// ask for.
///
/// The toggle is what decides what the times mean. On, they are the whole of how long the
/// engine is kept, whatever the user is doing; off, `Engine → AFK Timer` is what bounds it
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
/// `base` is where this group's commands begin in the numbering above.
fn append_codec_group(parent: HMENU, label: PCWSTR, rows: Vec<Row>, base: u16) {
    let group = unsafe { CreatePopupMenu().unwrap() };

    // The labels are kept for as long as the group is being filled out, for the reason the
    // cache labels are: `AppendMenuW` is handed a pointer, so the wide strings have to
    // outlive the call that lists them.
    let labels: Vec<Vec<u16>> = rows
        .iter()
        .map(|row| {
            codec_row_label(row)
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();

    for (index, row) in rows.iter().enumerate() {
        // A row the machine has is inert, and so is a row it is missing that has nowhere to
        // be got from: neither is given an id, so picking one closes the menu and does
        // nothing else. The two are told apart by being grey where the second is — a row
        // cannot be both normal-looking and unpickable, and the cross every missing row
        // carries is what says which of them a row is.
        let pickable = !row.available && row.link.is_some();

        let flags = if pickable || row.available {
            MF_STRING
        } else {
            MF_STRING | MF_GRAYED
        };

        let id = if pickable {
            (base + index as u16) as usize
        } else {
            0
        };

        let _ = unsafe { AppendMenuW(group, flags, id, PCWSTR(labels[index].as_ptr())) };
    }

    let _ = unsafe { AppendMenuW(parent, MF_STRING | MF_POPUP, group.0 as usize, label) };
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

/// What an away time is called in the `AFK Timer` submenu: the time, with the one an engine
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

/// The away time an item of the `AFK Timer` submenu stands for, by the position it was
/// listed at. An id past the last time the menu offered is one that is not there.
pub(super) fn afk_timer_secs_at(index: u16) -> Option<u64> {
    AFK_TIMER_CHOICES_SECS.get(index as usize).copied()
}

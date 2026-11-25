//! The menu itself: the popup that opens on a click, built in one pass.
//!
//! This is the whole of what a user sees when they right-click the tray icon, in the order
//! they see it: each row, each submenu hanging off one, and the checkmark or the radio mark
//! that says which of the choices the setting is on. Every value menu says two things about
//! itself - which item is current, and which one a setting nobody has changed would be - and
//! the second is read from the setting rather than written into the label, so a default that
//! moves takes the mark with it (see `submenus::default_label`).
//!
//! The builders this menu is assembled from are one per submenu and live in `submenus`; the
//! handlers a click reaches are in `commands`.

use super::ids::{
    ANIMATED_WEBP_BUDGET_CHOICES_MB, BACKGROUND_CHOICES, DDS_BACKGROUND_CHOICES,
    DECODE_BUDGET_CHOICES_GB, DISK_CACHE_SIZE_CHOICES_MB, FONT_SIZE_CHOICES,
    HTML_BACKGROUND_CHOICES, ID_TRAY_ANIMATED_SCALE_BASE, ID_TRAY_ANIMATED_WEBP_BUDGET_BASE,
    ID_TRAY_AUDIO_LOOP, ID_TRAY_AUDIO_SCALE_BASE, ID_TRAY_AUDIO_SHUFFLE, ID_TRAY_AUDIO_VOLUME_BASE,
    ID_TRAY_AVOID_BASE, ID_TRAY_BROWSE_DIALOG_PREVIEW_ENABLED, ID_TRAY_BYPASS_PLACEMENT,
    ID_TRAY_CHECK_UPDATES, ID_TRAY_CUSTOM_PREVIEW_EDIT, ID_TRAY_CUSTOM_PREVIEW_ENABLE,
    ID_TRAY_DDS_BACKGROUND_BASE, ID_TRAY_DECODE_BUDGET_BASE, ID_TRAY_DELAY_BASE,
    ID_TRAY_DESIGN_BACKGROUND_BASE, ID_TRAY_DESIGN_SCALE_BASE, ID_TRAY_DOCUMENT_CACHE_BASE,
    ID_TRAY_DOCUMENT_SCALE_BASE, ID_TRAY_EBOOK_SCALE_BASE, ID_TRAY_ENABLE,
    ID_TRAY_ENGINE_IDLE_BASE, ID_TRAY_EVERYTHING_PREVIEWS, ID_TRAY_EXIT, ID_TRAY_F3D_IDLE_BASE,
    ID_TRAY_FONT_BACKGROUND_BASE, ID_TRAY_FONT_SCALE_BASE, ID_TRAY_GENERAL_DISK_CACHE_BASE,
    ID_TRAY_HTML_BACKGROUND_BASE, ID_TRAY_IMAGE_BACKGROUND_BASE, ID_TRAY_IMAGE_CACHE_BASE,
    ID_TRAY_IMAGE_DISK_CACHE_BASE, ID_TRAY_IMAGE_SHUFFLE, ID_TRAY_LIBREOFFICE_IDLE_BASE,
    ID_TRAY_MARKDOWN_RENDERED, ID_TRAY_MARKDOWN_SOURCE, ID_TRAY_MODEL_SCALE_BASE,
    ID_TRAY_MPV_IDLE_BASE, ID_TRAY_NORMALIZE_VIDEO_VOLUME, ID_TRAY_NORMALIZE_VOLUME,
    ID_TRAY_OPEN_CONFIG, ID_TRAY_PDF_AUTOFOCUS, ID_TRAY_PDF_FULL_MODE, ID_TRAY_PIN,
    ID_TRAY_PIN_ALWAYS_ON_TOP, ID_TRAY_PIN_MINIMIZE_BUBBLE, ID_TRAY_PIN_MINIMIZE_TASKBAR,
    ID_TRAY_PIN_MODE_AUDIO_CAMERA_FOLLOWS, ID_TRAY_PIN_MODE_AUDIO_END_OF_FILE_BASE,
    ID_TRAY_PIN_MODE_AUDIO_PLAYLIST_ALWAYS_OPEN, ID_TRAY_PIN_MODE_VIDEO_END_OF_FILE_BASE,
    ID_TRAY_PIN_PAUSE_AUDIO, ID_TRAY_PIN_PAUSE_VIDEO, ID_TRAY_PIN_RESTORE,
    ID_TRAY_PIN_SCALING_UNTIL_RESIZED, ID_TRAY_PIN_SCALING_USE_SETTINGS, ID_TRAY_PIN_UPDATE,
    ID_TRAY_PIN_UPDATE_HOVER, ID_TRAY_PIN_UPDATE_TOP, ID_TRAY_POSITION_BEST,
    ID_TRAY_POSITION_FOLLOW, ID_TRAY_PRIORITIZE_KEYBOARD, ID_TRAY_Q_DIR_PREVIEWS,
    ID_TRAY_REHOVER_DELAY_BASE, ID_TRAY_REMEMBER_VIDEO_VOLUME, ID_TRAY_REMEMBER_VOLUME,
    ID_TRAY_RENDER_HTML, ID_TRAY_RESET_LISTS, ID_TRAY_RESET_SETTINGS, ID_TRAY_SCALE_BASE,
    ID_TRAY_SETTLING_DELAY_BASE, ID_TRAY_SILENT_AUTO_UPDATE, ID_TRAY_STARTUP,
    ID_TRAY_TEXT_DONT_PREVIEW_SECRETS, ID_TRAY_TEXT_SCALE_BASE, ID_TRAY_TEXT_SHOW_EXTENSIONLESS,
    ID_TRAY_THEME_CUSTOM_BASE, ID_TRAY_THEME_DARK, ID_TRAY_THEME_LIGHT, ID_TRAY_TICK_BASE,
    ID_TRAY_TRIGGER_AFFECT_PIN, ID_TRAY_TRIGGER_DISABLE, ID_TRAY_TRIGGER_ENABLE,
    ID_TRAY_TRIGGER_ENABLED, ID_TRAY_TYPE_ARCHIVES, ID_TRAY_TYPE_AUDIO, ID_TRAY_TYPE_DESIGN,
    ID_TRAY_TYPE_DOCUMENT, ID_TRAY_TYPE_EBOOK, ID_TRAY_TYPE_FONTS, ID_TRAY_TYPE_IMAGES,
    ID_TRAY_TYPE_MODELS, ID_TRAY_TYPE_TEXT, ID_TRAY_TYPE_VECTOR, ID_TRAY_TYPE_VIDEOS,
    ID_TRAY_UPDATE, ID_TRAY_VECTOR_BACKGROUND_BASE, ID_TRAY_VECTOR_SCALE_BASE,
    ID_TRAY_VIDEO_HW_ACCEL, ID_TRAY_VIDEO_LOOP, ID_TRAY_VIDEO_SCALE_BASE, ID_TRAY_VIDEO_SHUFFLE,
    ID_TRAY_VIDEO_SUBTITLES, ID_TRAY_VIDEO_VOLUME_BASE, ID_TRAY_WEBVIEW_IDLE_BASE,
    ID_TRAY_ZOOM_BACKGROUND_BASE, IMAGE_CACHE_SIZE_CHOICES_MB, MAX_TRAY_CUSTOM_THEMES,
    TRAY_CUSTOM_THEMES, ZOOM_BACKGROUND_CHOICES,
};
use super::submenus::{
    animated_webp_budget_label, append_afk_timer_menu, append_audio_scale_menu, append_avoid_menu,
    append_background_menu, append_bitmap_scale_menu, append_codecs_menu,
    append_document_scale_menu, append_engine_idle_menu, append_labeled_item, append_loop_item,
    append_office_engine_menu, append_pin_mode_menu, append_pin_nav_menu,
    append_pin_seek_menu_renamed, append_playlist_menu, append_seek_menu_renamed,
    append_shuffle_item, append_tick_menu, append_video_engine_menu, cache_size_label,
    decode_budget_label, default_label, enable_preview_label, pin_key_label, pin_minimize_label,
    pin_scaling_label, pin_update_label, remember_volume_label, system_menu_label,
    timing_delay_menu, update_available_label, MenuGuard, MF_RADIOCHECK,
};

use crate::app::seam::{send_preview, PreviewMessage};
use crate::app::updates;
use crate::config::config::{
    sanitize_animated_webp_budget_mb, sanitize_decode_budget_gb, sanitize_text_font_scale_percent,
    MarkdownMode, PinMinimizeTo, PinScalingMode, PreviewType, TextTheme, TriggerKeyMode,
    VideoEngine, DEFAULT_ANIMATED_SCALE, DEFAULT_AUDIO_LOOP, DEFAULT_AUDIO_SCALE,
    DEFAULT_AUDIO_VOLUME, DEFAULT_DDS_BACKGROUND, DEFAULT_DESIGN_BACKGROUND, DEFAULT_DESIGN_SCALE,
    DEFAULT_DOCUMENT_CACHE_MB, DEFAULT_DOCUMENT_SCALE, DEFAULT_EBOOK_SCALE, DEFAULT_FOLLOW_CURSOR,
    DEFAULT_FONT_BACKGROUND, DEFAULT_FONT_SCALE, DEFAULT_GENERAL_DISK_CACHE_MB,
    DEFAULT_HOVER_DELAY_MS, DEFAULT_HTML_BACKGROUND, DEFAULT_IMAGE_BACKGROUND,
    DEFAULT_IMAGE_CACHE_MB, DEFAULT_IMAGE_DISK_CACHE_MB, DEFAULT_MODEL_SCALE,
    DEFAULT_PDF_AUTOFOCUS, DEFAULT_PDF_FULL_MODE, DEFAULT_PIN_MODE_AUDIO_END_OF_FILE,
    DEFAULT_PIN_MODE_AUDIO_SHUFFLE, DEFAULT_PIN_MODE_IMAGE_SHUFFLE,
    DEFAULT_PIN_MODE_VIDEO_END_OF_FILE, DEFAULT_PIN_MODE_VIDEO_SHUFFLE, DEFAULT_PREVIEW_SCALE,
    DEFAULT_SAME_FILE_REHOVER_DELAY_MS, DEFAULT_SETTLING_DELAY_MS, DEFAULT_TEXT_SCALE,
    DEFAULT_VECTOR_BACKGROUND, DEFAULT_VECTOR_SCALE, DEFAULT_VIDEO_LOOP, DEFAULT_VIDEO_SCALE,
    DEFAULT_VIDEO_VOLUME, DEFAULT_ZOOM_BACKGROUND, VOLUME_CHOICES,
};
use crate::config::theme_files;
use crate::engines::libreoffice_render;
use crate::engines::webview_preview;
use crate::formats::codecs::{self, refresh as refresh_codecs};
use crate::text::text_theme;
use crate::{app::startup, CONFIG};
use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CheckMenuRadioItem, GetCursorPos, SetForegroundWindow, TrackPopupMenu, HMENU,
    MF_BYCOMMAND, MF_CHECKED, MF_ENABLED, MF_GRAYED, MF_SEPARATOR, MF_STRING, MF_UNCHECKED,
    TPM_BOTTOMALIGN, TPM_LEFTALIGN,
};

struct MenuOpenGuard;

impl Drop for MenuOpenGuard {
    fn drop(&mut self) {
        updates::note_a_menu_is_open(false);
    }
}

pub(super) unsafe fn show_context_menu(hwnd: HWND) {
    let _ = try_show_context_menu(hwnd);
}

unsafe fn try_show_context_menu(hwnd: HWND) -> Option<()> {
    let root_menu_guard = MenuGuard::new()?;
    let menu = root_menu_guard.handle();
    // The row that asked for the update check is the one a silent install landing mid-click would
    // take away, so a menu is open from here to the moment `TrackPopupMenu` returns — before
    // `DestroyMenu`, since the install it holds back ends this app and a menu whose window is gone
    // is no menu at all (see `updates::note_a_menu_is_open`).
    updates::note_a_menu_is_open(true);
    let open_guard = MenuOpenGuard;

    // **And the startup entry is asked about at every opening of this menu, not only at a start.**
    // It is three registry reads in the ordinary case — the value, the flag Windows keeps beside it,
    // and the command line in it — and what it buys is that the row a user is about to read and the
    // entry a logon will run are the same fact: a `config.ini` hand-edited to `run_at_startup=true`
    // is written here, and an entry a cleaner removed, switched off in Windows, or left pointing at
    // another copy is put right before the checkmark is drawn from it (see
    // `startup::honour_the_configuration`). It is taken before the snapshot below so that the file's
    // own copy of the setting, which this can put right, is read after the correcting rather than
    // before it.
    crate::app::startup::honour_the_configuration();

    // Lock released before the publish: the message a publish sends locks `CONFIG` again on
    // the preview loop. A level this process wrote itself is not an edit, and is not reported
    // as one (see `reload_from_disk_keeping_our_own_levels`).
    let (changed, snap) = {
        let mut config = CONFIG.lock().unwrap_or_else(|e| e.into_inner());
        let changed = config.reload_from_disk_keeping_our_own_levels();
        (changed, config.clone())
    };

    // A level hand-edited into config.ini reaches a pin already playing, which holds a level of
    // its own; an unchanged half is left alone so opening the menu does not push the file's
    // level onto a knob dragged elsewhere (see `tray_volume_moved`).
    if changed.0 {
        send_preview(PreviewMessage::TrayVolume { audio: false });
    }
    if changed.1 {
        send_preview(PreviewMessage::TrayVolume { audio: true });
    }

    updates::request_check();

    populate_context_menu(menu, &snap)?;

    // Get cursor position and show menu
    let mut pt = windows::Win32::Foundation::POINT::default();
    let _ = GetCursorPos(&mut pt);

    let _ = SetForegroundWindow(hwnd).ok();
    let _ = TrackPopupMenu(
        menu,
        TPM_LEFTALIGN | TPM_BOTTOMALIGN,
        pt.x,
        pt.y,
        0,
        hwnd,
        None,
    )
    .ok();
    drop(open_guard);
    drop(root_menu_guard);
    Some(())
}

/// Build the entire tray popup menu from one consistent `AppConfig` snapshot (zero `CONFIG` locks).
pub(super) unsafe fn populate_context_menu(
    menu: HMENU,
    snap: &crate::config::config::AppConfig,
) -> Option<()> {
    // Add "Enable Preview" with checkmark, naming the key that throws the same switch from
    // Explorer so the shortcut is visible where the setting is.
    let preview_enabled = snap.preview_enabled;
    let toggle_key = snap.toggle_key.as_str();
    let enable_flags = MF_STRING
        | if preview_enabled {
            MF_CHECKED
        } else {
            MF_UNCHECKED
        };
    append_labeled_item(
        menu,
        enable_flags,
        ID_TRAY_ENABLE,
        &enable_preview_label(toggle_key),
    );

    // The `Restore Pin Mode` row between the two, and only while a pin is actually alive: clicking
    // it brings the pin's window back over everything, whether it was put away as a bubble, minimized
    // into the taskbar, or left standing under the program the user moved on to. A row that is there
    // with no pin to restore would be a click that does nothing, so it is not appended at all — and
    // `the_pin_is_up` is the seam's published answer rather than the tray's own, which is what makes
    // a pin a detached window of its own count as alive (see `ID_TRAY_PIN_RESTORE`).
    if crate::app::seam::the_pin_is_up() {
        let _ = AppendMenuW(
            menu,
            MF_STRING,
            ID_TRAY_PIN_RESTORE as usize,
            w!("Restore Pin Mode"),
        );
    }

    // And the "Pin Mode" submenu below it, which is the pin's own business: whether the key is
    // watched at all, whether a pin that is up is shown another file while the user picks one,
    // and where a pin the minimize button puts away goes. The key's row names
    // the key that pins the way the trigger key's own submenu names its own: the key is a
    // setting, so the row says which one is watched rather than only whether one is (see
    // `key_input`).
    let pin_enabled = snap.pin_enabled;
    let pin_always_on_top = snap.pin_always_on_top;
    let pin_key = snap.pin_key.as_str();
    let pin_minimize_to = snap.pin_minimize_to;
    let pin_scaling_mode = snap.pin_scaling_mode;
    let pin_update = snap.pin_update_enabled;
    let pin_update_on_hover = snap.pin_update_on_hover;
    let pin_update_move_to_top = snap.pin_update_move_to_top;
    let pin_nav_image = snap.pin_nav_image;
    let pin_nav_video = snap.pin_nav_video;
    let pin_nav_audio = snap.pin_nav_audio;
    let pin_nav_others = snap.pin_nav_others;
    let pin_key_display = pin_key_label(pin_key);
    let pin_label = format!("Enable ({pin_key_display})");
    let pin_label_wide: Vec<u16> = pin_label.encode_utf16().chain(std::iter::once(0)).collect();
    let pin_menu_guard = MenuGuard::new()?;
    let pin_menu = pin_menu_guard.handle();
    let _ = AppendMenuW(
        pin_menu,
        MF_STRING
            | if pin_enabled {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_PIN as usize,
        PCWSTR(pin_label_wide.as_ptr()),
    );
    let _ = AppendMenuW(
        pin_menu,
        MF_STRING
            | if pin_always_on_top {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_PIN_ALWAYS_ON_TOP as usize,
        w!("Always On Top"),
    );

    // The `Update Preview` submenu: whether a pin that is up is shown the file the user picks
    // next — the one the pointer clicks, or the one the keyboard selects — whether being told
    // brings it in front of what it was behind, and whether the pointer's own hover is one of the
    // ways it is told about one. A pin is a window to read or to watch, so all three are off the
    // pin's own row rather than part of it.
    //
    // The two rows below the first are greyed while it is off, because they say nothing then: a
    // pin that follows nothing has no update to come forward for, and one that follows nothing
    // does not follow a hover either. A row that could be ticked without effect would be a setting
    // a user cannot tell from one that does nothing.
    let update_menu_guard = MenuGuard::new()?;
    let update_menu = update_menu_guard.handle();
    append_labeled_item(
        update_menu,
        MF_STRING | if pin_update { MF_CHECKED } else { MF_UNCHECKED },
        ID_TRAY_PIN_UPDATE,
        pin_update_label(),
    );
    // The raise sits right below the follow, because the two are one row read in two parts: the
    // first says the pin is shown the file picked next, the second what it does about standing
    // behind the listing it was picked in — and the user who turns the second off is asking about
    // the very update the two of them describe.
    let mut move_to_top_flags = MF_STRING
        | if pin_update_move_to_top {
            MF_CHECKED
        } else {
            MF_UNCHECKED
        };
    if !pin_update {
        move_to_top_flags |= MF_GRAYED;
    }
    let _ = AppendMenuW(
        update_menu,
        move_to_top_flags,
        ID_TRAY_PIN_UPDATE_TOP as usize,
        w!("Move To Top"),
    );
    // Then a separator, because the row below is not the same kind of row: the pair above is what
    // an update is, and this is a second way the file that update names is chosen.
    let _ = AppendMenuW(update_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let mut hover_flags = MF_STRING
        | if pin_update_on_hover {
            MF_CHECKED
        } else {
            MF_UNCHECKED
        };
    if !pin_update {
        hover_flags |= MF_GRAYED;
    }
    let _ = AppendMenuW(
        update_menu,
        hover_flags,
        ID_TRAY_PIN_UPDATE_HOVER as usize,
        w!("On Hover"),
    );
    let _ = update_menu_guard.append_to(pin_menu, MF_STRING, w!("Update Preview"));

    // The `Minimize` submenu: where a pin the minimize button puts away goes. The two
    // places are one of two rather than a switch — a
    // minimize leaves either the round bubble on the desktop or the app's own button in the
    // taskbar, never both — so they carry a mark apiece and the setting picks between them (see
    // `PinMinimizeTo`), and the bubble is the answer the app starts at.
    //
    // What such a pin holds while it is away is no longer asked here: the two `Pause
    // on Minimize` rows hang under the kind's own `Pin Mode` instead, because a film and a sound
    // are two different things to want quiet and each kind's submenu is the one that knows
    // which of the two is on screen (see `append_pause_minimized_item`).
    let minimize_menu_guard = MenuGuard::new()?;
    let minimize_menu = minimize_menu_guard.handle();
    let minimize_flag = |candidate: PinMinimizeTo| {
        MF_STRING
            | if pin_minimize_to == candidate {
                MF_CHECKED | MF_RADIOCHECK
            } else {
                MF_UNCHECKED
            }
    };
    append_labeled_item(
        minimize_menu,
        minimize_flag(PinMinimizeTo::Bubble),
        ID_TRAY_PIN_MINIMIZE_BUBBLE,
        &pin_minimize_label(PinMinimizeTo::Bubble),
    );
    append_labeled_item(
        minimize_menu,
        minimize_flag(PinMinimizeTo::Taskbar),
        ID_TRAY_PIN_MINIMIZE_TASKBAR,
        &pin_minimize_label(PinMinimizeTo::Taskbar),
    );
    let _ = minimize_menu_guard.append_to(pin_menu, MF_STRING, w!("Minimize"));

    let scaling_menu_guard = MenuGuard::new()?;
    let scaling_menu = scaling_menu_guard.handle();
    let scaling_flag = |candidate: PinScalingMode| {
        MF_STRING
            | if pin_scaling_mode == candidate {
                MF_CHECKED | MF_RADIOCHECK
            } else {
                MF_UNCHECKED
            }
    };
    append_labeled_item(
        scaling_menu,
        scaling_flag(PinScalingMode::UntilResized),
        ID_TRAY_PIN_SCALING_UNTIL_RESIZED,
        &pin_scaling_label(PinScalingMode::UntilResized),
    );
    append_labeled_item(
        scaling_menu,
        scaling_flag(PinScalingMode::UseScalingSettings),
        ID_TRAY_PIN_SCALING_USE_SETTINGS,
        &pin_scaling_label(PinScalingMode::UseScalingSettings),
    );
    let _ = scaling_menu_guard.append_to(pin_menu, MF_STRING, w!("Scaling"));

    // The `Navigation` submenu: four anchors — Image, Video, Audio, Others, in menu
    // order — each with its own All/Same Category pair. The walk is anchored on the first
    // file pinned, so a video walks videos while a document walks end to end (see
    // `nav_mode_for`).
    append_pin_nav_menu(
        pin_menu,
        pin_nav_image,
        pin_nav_video,
        pin_nav_audio,
        pin_nav_others,
    );

    let _ = pin_menu_guard.append_to(menu, MF_STRING, w!("Pin Mode"));

    // Add the "Preview Types" submenu: one gate per kind of preview, on by
    // default, in three groups — what is looked at and heard, what is
    // read, and what is drawn — with a separator between each. A gate
    // is only whether previews of that kind may be shown at all — the
    // lists and settings that decide which files of that kind preview
    // are left alone, so switching one off and back on restores what
    // was configured.
    //
    // One gate covers each pair of a kind this app reads and the kind an engine draws for
    // it: an ImageMagick picture is a picture, a document LibreOffice drew is a document,
    // an archive PeaZip listed is an archive, and a book Calibre converted is a book, so
    // none of the four has a row here.
    let kinds = [
        (PreviewType::Images, ID_TRAY_TYPE_IMAGES, w!("Image")),
        (PreviewType::Videos, ID_TRAY_TYPE_VIDEOS, w!("Video")),
        (PreviewType::Audio, ID_TRAY_TYPE_AUDIO, w!("Audio")),
        (PreviewType::Text, ID_TRAY_TYPE_TEXT, w!("Text")),
        (PreviewType::Ebook, ID_TRAY_TYPE_EBOOK, w!("Ebook")),
        (PreviewType::Archives, ID_TRAY_TYPE_ARCHIVES, w!("Archive")),
        (PreviewType::Document, ID_TRAY_TYPE_DOCUMENT, w!("Document")),
        (PreviewType::Vector, ID_TRAY_TYPE_VECTOR, w!("Vector")),
        (PreviewType::Fonts, ID_TRAY_TYPE_FONTS, w!("Font")),
        (PreviewType::Design, ID_TRAY_TYPE_DESIGN, w!("Design")),
        (PreviewType::Models, ID_TRAY_TYPE_MODELS, w!("3D")),
    ];
    let types_menu_guard = MenuGuard::new()?;
    let types_menu = types_menu_guard.handle();

    for (index, (kind, id, label)) in kinds.into_iter().enumerate() {
        // The three groups — media, things read, and graphics — are
        // separated because a preview is one of those three things,
        // and which of them a kind is, is the first thing the submenu
        // answers about it.
        if index == 3 || index == 7 {
            let _ = AppendMenuW(types_menu, MF_SEPARATOR, 0, PCWSTR::null());
        }

        let flags = MF_STRING
            | if kind.enabled_in(snap) {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            };
        let _ = AppendMenuW(types_menu, flags, id as usize, label);
    }

    let _ = types_menu_guard.append_to(menu, MF_STRING, w!("Preview Types"));

    // The `Preview Settings` submenu, below `Preview Types`: one half per kind that has
    // settings of its own beyond whether it previews at all — what its previews are like, the
    // level it is heard at, whether a hovered one goes round, what a *pinned* one does, and
    // where in a file a sound starts — so a film's settings and a sound's are two menus rather
    // than one (see `append_loop_item` and `append_pin_mode_menu`).
    //
    // `Preview Settings` holds what a preview *is* rather than what it is heard at: the film's own
    // subtitle tracks, whether a hovered one goes round, and what a *pinned* one does — the
    // levels of both kinds are one menu of their own (see `volume_menu`).
    //
    // The film's and the sound's halves used to carry the `Volume` halves the old `Volume` menu
    // held, with the sound's `Audio Seek` renamed `Seek` and its `Pin Mode Audio Seek` renamed
    // `Seek` too, inside `Pin Mode` beside the rest of a pinned sound's own rows.
    //
    // The film's and the sound's `Pause on Minimize` rows are read here, beside the shuffle and
    // the end of file they hang above: all three are what a *pinned* preview of that kind does,
    // and the two switches moved out of `Pin Mode → Minimize` for the same reason the submenus
    // are split in two — a film and a sound are two different things to want quiet.
    let video_volume = snap.video_volume;
    let audio_volume = snap.audio_volume;
    let audio_seek = snap.audio_seek;
    let pin_mode_audio_seek = snap.pin_mode_audio_seek;
    let video_loop = snap.video_loop;
    let audio_loop = snap.audio_loop;
    let video_shuffle = snap.pin_mode_video_shuffle;
    let audio_shuffle = snap.pin_mode_audio_shuffle;
    let image_shuffle = snap.pin_mode_image_shuffle;
    let pin_pause_video = snap.pin_pause_video;
    let pin_pause_audio = snap.pin_pause_audio;
    let pin_video_end_of_file = snap.pin_mode_video_end_of_file;
    let pin_audio_end_of_file = snap.pin_mode_audio_end_of_file;

    let preview_settings_guard = MenuGuard::new()?;
    let preview_settings = preview_settings_guard.handle();
    let append_levels = |levels: HMENU, current: u32, base: u16, default: u32| {
        // Largest first, which is the order a level is looked for in: a pointer crossing a folder
        // of sounds is most often turning one down, and the whole of the scale standing at the top
        // is what the rest of the list is read against. The id a level is listed at stays the
        // table's own position — only the order the items are appended in is turned round — so a
        // click still names its level by the position it holds in `VOLUME_CHOICES` (see
        // `set_audio_volume`).
        for (index, level) in VOLUME_CHOICES.iter().enumerate().rev() {
            let flags = MF_STRING
                | if current == *level {
                    MF_CHECKED | MF_RADIOCHECK
                } else {
                    MF_UNCHECKED
                };

            append_labeled_item(
                levels,
                flags,
                base + index as u16,
                &default_label(&format!("{level}%"), *level == default),
            );
        }
    };

    // Each half carries one row the levels do not answer, and it sits above them: the loudness a file's
    // playing is measured to. A level is asked of the hover and this is asked of the file, and the
    // two are scaled into one another — what the file's own measured loudness is brought to is one
    // level, and the level under it is how much of that is heard (see `Normalize`). Both halves are
    // offered it because both can be heard; they start where they start, the way their levels do.
    let normalize_audio = snap.normalize_volume;
    let normalize_video = snap.normalize_video_volume;
    let remember_audio = snap.remember_audio_volume;
    let remember_video = snap.remember_video_volume;
    let video_subtitles = snap.video_subtitles;
    let video_engine = snap.video_engine;
    // Asked again here for the reason the `Components` rows are asked again: a machine that has just
    // been given FFmpeg is answered from the machine rather than from the hover that cached it
    // (see `codecs::refresh`).
    refresh_codecs();
    let normalize_available = codecs::normalize_available();

    // The two rows the levels do not answer, above them: the loudness a file's playing is measured to,
    // and whether a level turned on a pin is kept. The second is never greyed — it asks about this
    // app's own knob rather than about FFmpeg, so it acts on a machine that has no FFmpeg at all.
    let append_switches =
        |levels: HMENU, normalize: bool, normalize_id: u16, remember: bool, remember_id: u16| {
            let _ = AppendMenuW(
                levels,
                MF_STRING
                    | if normalize && normalize_available {
                        MF_CHECKED
                    } else {
                        MF_UNCHECKED
                    }
                    | if normalize_available {
                        MF_UNCHECKED
                    } else {
                        // A machine without FFmpeg has nothing that measures a loudness or applies
                        // one: the
                        // row is shown as what it is there — a switch that cannot act — rather than as a
                        // click that would do nothing (see `codecs::normalize_available`).
                        MF_GRAYED
                    },
                normalize_id as usize,
                w!("Normalize (Expensive)"),
            );
            append_labeled_item(
                levels,
                MF_STRING | if remember { MF_CHECKED } else { MF_UNCHECKED },
                remember_id,
                remember_volume_label(),
            );
        };

    let video_levels_guard = MenuGuard::new()?;
    let video_levels = video_levels_guard.handle();
    append_switches(
        video_levels,
        normalize_video,
        ID_TRAY_NORMALIZE_VIDEO_VOLUME,
        remember_video,
        ID_TRAY_REMEMBER_VIDEO_VOLUME,
    );
    let _ = AppendMenuW(video_levels, MF_SEPARATOR, 0, PCWSTR::null());
    append_levels(
        video_levels,
        video_volume,
        ID_TRAY_VIDEO_VOLUME_BASE,
        DEFAULT_VIDEO_VOLUME,
    );

    let audio_levels_guard = MenuGuard::new()?;
    let audio_levels = audio_levels_guard.handle();
    append_switches(
        audio_levels,
        normalize_audio,
        ID_TRAY_NORMALIZE_VOLUME,
        remember_audio,
        ID_TRAY_REMEMBER_VOLUME,
    );
    let _ = AppendMenuW(audio_levels, MF_SEPARATOR, 0, PCWSTR::null());
    append_levels(
        audio_levels,
        audio_volume,
        ID_TRAY_AUDIO_VOLUME_BASE,
        DEFAULT_AUDIO_VOLUME,
    );

    // The `Volume` submenu's own two halves, hung from the tray rather than from either
    // kind: what a film and a sound are heard at is one question asked of both, so it is one
    // menu with a half apiece rather than a row of the film's and the sound's halves of
    // `Preview Settings` (see `append_switches` and `append_levels`).
    let volume_menu_guard = MenuGuard::new()?;
    let volume_menu = volume_menu_guard.handle();
    let _ = video_levels_guard.append_to(volume_menu, MF_STRING, w!("Video"));
    let _ = audio_levels_guard.append_to(volume_menu, MF_STRING, w!("Audio"));

    // The picture kind's own half, above the film's: the two questions a picture answers that
    // no other kind's rows answer for it — whether a *pinned* one takes the next file of the
    // folder at random, and what fills the band beside one the wheel has taken past its window.
    // A picture has no level to be heard at, nothing to go round at the end of, and nowhere in
    // it to start, so those two are the whole of it: a hover has no next file, and a step back
    // is the file before it either way (see `append_shuffle_item`).
    //
    // The zoom's backdrop hangs inside `Pin Mode` because a zoom is a pinned picture's — the
    // wheel reaches a hover nowhere — and it is one submenu rather than a `Zoom` holding a
    // `Background`, since a zoom has nothing else to list. It is a question of the window's
    // rather than the file's: what stands behind a picture nobody has touched is the
    // `Background → Image` setting above, while this band is what the zoom opened — so it
    // offers the three backdrops that fill a band rather than the four a preview can be drawn
    // over, and the paint reads it only for a picture with no transparency in it (see
    // `pinned_band_backdrop`).
    let image_pin_mode_guard = MenuGuard::new()?;
    let image_pin_mode = image_pin_mode_guard.handle();
    append_shuffle_item(
        image_pin_mode,
        ID_TRAY_IMAGE_SHUFFLE,
        image_shuffle,
        DEFAULT_PIN_MODE_IMAGE_SHUFFLE,
    );
    let zoom_background = snap.zoom_background;
    append_background_menu(
        image_pin_mode,
        w!("Zoom Background"),
        ID_TRAY_ZOOM_BACKGROUND_BASE,
        &ZOOM_BACKGROUND_CHOICES,
        zoom_background,
        DEFAULT_ZOOM_BACKGROUND,
    );
    let image_menu_guard = MenuGuard::new()?;
    let image_menu = image_menu_guard.handle();
    let _ = image_pin_mode_guard.append_to(image_menu, MF_STRING, w!("Pin Mode"));

    // The `General` half, above the kinds and before the picture's own: whether a preview is made
    // where a **browse dialog** is showing. It is not a question about a kind of file — the same
    // file is previewed or not depending only on which listing it is seen in — so it hangs above
    // the kinds rather than inside one of them, and it is a submenu of its own for the layout's own
    // reason: a `General` heading is where a row about *where* previews happen sits, and the kinds
    // below it each answer what a preview of their kind is like.
    let browse_dialog_preview_enabled = snap.browse_dialog_preview_enabled;
    let everything_previews = snap.everything_previews;
    let q_dir_previews = snap.q_dir_previews;
    let general_guard = MenuGuard::new()?;
    let general = general_guard.handle();
    append_labeled_item(
        general,
        MF_STRING
            | if browse_dialog_preview_enabled {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_BROWSE_DIALOG_PREVIEW_ENABLED,
        "Enable in Browse Dialogs",
    );
    // The same question asked of another kind of listing: Everything's results, where the files its
    // rows name are spread over every folder the search was pointed at. It is a row about *where* a
    // preview happens, which is why it sits here rather than with the kinds below.
    append_labeled_item(
        general,
        MF_STRING
            | if everything_previews {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_EVERYTHING_PREVIEWS,
        "Enable in Everything",
    );
    // And of the fourth kind: a Q-Dir frame's four panes, each showing one folder of its own.
    append_labeled_item(
        general,
        MF_STRING
            | if q_dir_previews {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_Q_DIR_PREVIEWS,
        "Enable in Q-Dir",
    );
    let _ = general_guard.append_to(preview_settings, MF_STRING, w!("General"));

    let _ = image_menu_guard.append_to(preview_settings, MF_STRING, w!("Image"));

    // The film kind's own rows: whether a hovered film goes round at the end of
    // it, and what a *pinned* one does — whether its forward steps shuffle, and what its end
    // does, which is stop there, go round, or step on to the next file that the `Shuffle` row
    // decides between the file beside it and a random one (see `append_loop_item` and
    // `append_pin_mode_menu`).
    let video_preview_guard = MenuGuard::new()?;
    let video_preview = video_preview_guard.handle();
    // The film's own row, above `Loop`: whether its own subtitle tracks are probed for and
    // copied out at all — a switch about a film's picture rather than a level, so it is a row
    // of `Preview Settings → Video` and the film's alone and not the sound's (see
    // `video_subtitles`). Its name is the engine's business rather than one cost for all three
    // (see `subtitles_label`).
    append_labeled_item(
        video_preview,
        MF_STRING
            | if video_subtitles {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_VIDEO_SUBTITLES,
        &subtitles_label(video_engine, VideoEngine::Mpv.installed()),
    );
    append_loop_item(
        video_preview,
        ID_TRAY_VIDEO_LOOP,
        video_loop,
        DEFAULT_VIDEO_LOOP,
    );
    append_pin_mode_menu(
        video_preview,
        ID_TRAY_PIN_PAUSE_VIDEO,
        pin_pause_video,
        ID_TRAY_VIDEO_SHUFFLE,
        video_shuffle,
        DEFAULT_PIN_MODE_VIDEO_SHUFFLE,
        ID_TRAY_PIN_MODE_VIDEO_END_OF_FILE_BASE,
        pin_video_end_of_file,
        DEFAULT_PIN_MODE_VIDEO_END_OF_FILE,
    );
    let _ = video_preview_guard.append_to(preview_settings, MF_STRING, w!("Video"));

    // And the sound kind's own: the same `Loop` row and a `Seek` row of its own, with the
    // `Pin Mode` submenu below them — where in a file a hovered sound starts as the `Seek`
    // row says, and where a *pinned* one does, under `Pin Mode` beside the rest of a pinned
    // sound's own rows. Both are renamed from the old `Volume` menu's `Audio Seek` and `Pin
    // Mode Audio Seek` because they sit under `Audio` now rather than under `Volume`
    // (see `pin_mode_audio_seek`). `Pin Mode` is listed below the kind's own rows, so the
    // settings a hovered sound plays by come before the ones only a pinned sound has.
    // The band a pinned sound's card grows under it is a row of its own kind: whether its camera
    // follows the file the walk stepped onto is read here rather than with the rows above, since
    // it is the band's answer and not a way a sound is played (see `append_playlist_menu`).
    let pin_camera_follows = snap.pin_mode_audio_camera_follows;
    let pin_playlist_always_open = snap.pin_mode_audio_playlist_always_open;
    let audio_preview_guard = MenuGuard::new()?;
    let audio_preview = audio_preview_guard.handle();
    append_loop_item(
        audio_preview,
        ID_TRAY_AUDIO_LOOP,
        audio_loop,
        DEFAULT_AUDIO_LOOP,
    );
    append_seek_menu_renamed(audio_preview, w!("Seek"), audio_seek);
    let audio_pin_mode = append_pin_mode_menu(
        audio_preview,
        ID_TRAY_PIN_PAUSE_AUDIO,
        pin_pause_audio,
        ID_TRAY_AUDIO_SHUFFLE,
        audio_shuffle,
        DEFAULT_PIN_MODE_AUDIO_SHUFFLE,
        ID_TRAY_PIN_MODE_AUDIO_END_OF_FILE_BASE,
        pin_audio_end_of_file,
        DEFAULT_PIN_MODE_AUDIO_END_OF_FILE,
    );
    if let Some(audio_pin_mode) = audio_pin_mode {
        append_pin_seek_menu_renamed(audio_pin_mode, w!("Seek"), pin_mode_audio_seek);
        append_playlist_menu(
            audio_pin_mode,
            ID_TRAY_PIN_MODE_AUDIO_PLAYLIST_ALWAYS_OPEN,
            pin_playlist_always_open,
            ID_TRAY_PIN_MODE_AUDIO_CAMERA_FOLLOWS,
            pin_camera_follows,
        );
    }
    let _ = audio_preview_guard.append_to(preview_settings, MF_STRING, w!("Audio"));

    // Add the "Text" submenu: the size and Markdown mode a text preview is
    // painted with. What a text preview *is* rather than what it shows — the scrollbar, the
    // selection and the keys that copy it — is the pin's business rather than a setting's: a
    // text preview pinned with the key becomes one to work in (see the pin in `preview_window`).
    let text_menu_guard = MenuGuard::new()?;
    let text_menu = text_menu_guard.handle();

    // Add the Theme submenu
    let theme = snap.theme;

    // The folder is read here rather than once at startup: the files these names
    // stand for are the user's to add to and to edit, and this is the moment they
    // are looking at them.
    text_theme::refresh_custom();
    let mut custom_themes = theme_files::names();
    custom_themes.truncate(MAX_TRAY_CUSTOM_THEMES);

    let theme_menu_guard = MenuGuard::new()?;
    let theme_menu = theme_menu_guard.handle();

    // A file theme that no longer loads is painted with the default — see
    // `text_theme::loaded` — so the default is what is marked while the name it
    // was chosen under stays in `config.ini`.
    let marked = match theme {
        TextTheme::Custom(name) if text_theme::custom(name).is_none() => TextTheme::Light,
        theme => theme,
    };
    let theme_flag = |candidate: TextTheme| {
        MF_STRING
            | if marked == candidate {
                MF_CHECKED | MF_RADIOCHECK
            } else {
                MF_UNCHECKED
            }
    };
    let _ = AppendMenuW(
        theme_menu,
        theme_flag(TextTheme::Light),
        ID_TRAY_THEME_LIGHT as usize,
        w!("Atom One Light"),
    );
    let _ = AppendMenuW(
        theme_menu,
        theme_flag(TextTheme::Dark),
        ID_TRAY_THEME_DARK as usize,
        w!("One Dark Pro"),
    );

    if !custom_themes.is_empty() {
        let _ = AppendMenuW(theme_menu, MF_SEPARATOR, 0, PCWSTR::null());
    }

    let mut listed = Vec::with_capacity(custom_themes.len());
    for (index, name) in custom_themes.iter().enumerate() {
        let theme = TextTheme::custom(name);
        // `&&` is how an item asks for a literal ampersand; one on its own
        // underlines the character after it instead of standing for itself.
        let label: Vec<u16> = name
            .replace('&', "&&")
            .encode_utf16()
            .chain(std::iter::once(0))
            .collect();
        let _ = AppendMenuW(
            theme_menu,
            theme_flag(theme),
            ID_TRAY_THEME_CUSTOM_BASE as usize + index,
            PCWSTR(label.as_ptr()),
        );
        listed.push(theme);
    }
    if let Ok(mut items) = TRAY_CUSTOM_THEMES.lock() {
        *items = listed;
    }

    // Add the Font Size submenu, largest first: the steps a size can be picked
    // from, from the largest down to the smallest. A hand-edited size between these
    // steps simply matches none of them, which is why the values are read as
    // written.
    let font_scale = sanitize_text_font_scale_percent(snap.text_font_scale_percent);
    let font_menu_guard = MenuGuard::new()?;
    let font_menu = font_menu_guard.handle();

    let font_flag = |percent: u32| {
        MF_STRING
            | if font_scale == percent {
                MF_CHECKED | MF_RADIOCHECK
            } else {
                MF_UNCHECKED
            }
    };
    // The labels are built once and kept: `AppendMenuW` is handed a pointer, so the
    // wide strings have to outlive the call that lists them.
    let font_labels: Vec<Vec<u16>> = FONT_SIZE_CHOICES
        .iter()
        .map(|(percent, _)| {
            format!("{percent}%")
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();

    for (index, (percent, id)) in FONT_SIZE_CHOICES.iter().enumerate() {
        let _ = AppendMenuW(
            font_menu,
            font_flag(*percent),
            *id as usize,
            PCWSTR(font_labels[index].as_ptr()),
        );
    }
    let _ = font_menu_guard.append_to(text_menu, MF_STRING, w!("Font Size"));

    // Add the Markdown submenu
    let markdown_mode = snap.markdown_mode;
    let markdown_menu_guard = MenuGuard::new()?;
    let markdown_menu = markdown_menu_guard.handle();

    let markdown_flag = |candidate: MarkdownMode| {
        MF_STRING
            | if markdown_mode == candidate {
                MF_CHECKED | MF_RADIOCHECK
            } else {
                MF_UNCHECKED
            }
    };
    let _ = AppendMenuW(
        markdown_menu,
        markdown_flag(MarkdownMode::Rendered),
        ID_TRAY_MARKDOWN_RENDERED as usize,
        w!("Rendered"),
    );
    let _ = AppendMenuW(
        markdown_menu,
        markdown_flag(MarkdownMode::Source),
        ID_TRAY_MARKDOWN_SOURCE as usize,
        w!("Source"),
    );
    let _ = markdown_menu_guard.append_to(text_menu, MF_STRING, w!("Markdown"));

    // Whether a page of HTML is drawn by the browser engine rather than shown as its markup.
    // The whole menu is built from the configuration each time it is opened, so this is read
    // here rather than kept between the click and the row.
    let render_html = snap.render_html;
    let _ = AppendMenuW(
        text_menu,
        MF_STRING
            | if render_html {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_RENDER_HTML as usize,
        w!("Render HTML"),
    );

    let text_show_extensionless = snap.text_show_extensionless;
    let _ = AppendMenuW(
        text_menu,
        MF_STRING
            | if text_show_extensionless {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_TEXT_SHOW_EXTENSIONLESS as usize,
        w!("Preview All Extensionless"),
    );

    let text_dont_preview_secrets = snap.text_dont_preview_secrets;
    let _ = AppendMenuW(
        text_menu,
        MF_STRING
            | if text_dont_preview_secrets {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_TEXT_DONT_PREVIEW_SECRETS as usize,
        w!("Don't Preview Secret Files"),
    );

    let _ = text_menu_guard.append_to(preview_settings, MF_STRING, w!("Text"));

    // The document kind's own half: what a *pinned* document does, which is two rows — whether the
    // pin hands the document to Chromium's own viewer at all, and, where it does, whose the keyboard
    // is once that viewer stands in the band. A document is the one kind whose pin can be either a
    // reader or a picture of a page, and the top row is which (see `pdf_full_mode`); everything below
    // it is about the reader.
    //
    // It hangs under `PDF` rather than under `Ebook`, which is the gate the pages this app reads
    // are listed at: a pinned document is not read by this app at all, and `Ebook` is the name of the
    // reader rather than of the file.
    let pdf_autofocus = snap.pdf_autofocus;
    let pdf_full_mode = snap.pdf_full_mode;
    let pdf_pin_mode_guard = MenuGuard::new()?;
    let pdf_pin_mode = pdf_pin_mode_guard.handle();
    append_labeled_item(
        pdf_pin_mode,
        MF_STRING
            | if pdf_full_mode {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_PDF_FULL_MODE,
        &default_label("Full Mode", DEFAULT_PDF_FULL_MODE),
    );
    append_labeled_item(
        pdf_pin_mode,
        MF_STRING
            // A row about whose the keyboard is is a row about a document nothing shows when the
            // viewer is switched off, so it goes grey rather than being a setting that answers
            // nothing (see `pdf_full_mode` and `pdf_autofocus`).
            | if pdf_full_mode {
                MF_ENABLED
            } else {
                MF_GRAYED
            }
            | if pdf_autofocus {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_PDF_AUTOFOCUS,
        &default_label("Autofocus on PDF", DEFAULT_PDF_AUTOFOCUS),
    );
    let pdf_menu_guard = MenuGuard::new()?;
    let pdf_menu = pdf_menu_guard.handle();
    let _ = pdf_pin_mode_guard.append_to(pdf_menu, MF_STRING, w!("Pin Mode"));
    let _ = pdf_menu_guard.append_to(preview_settings, MF_STRING, w!("PDF"));

    // The theme no longer hangs on the top menu beside the text preview it themes: it is
    // moved into `System`, between the two rows about this app's own place in the machine and
    // the ones about its file (see the `System` submenu below). The `theme_menu` it is built
    // in is filled out here, where the theme's own names are read, and hangs under `System`
    // below.

    // And the whole of it goes on the top menu below `Preview Types`, one row rather than
    // three: the kinds it holds are the three that have settings beyond whether they preview
    // at all, which is what `Preview Types` answers for them.
    let _ = preview_settings_guard.append_to(menu, MF_STRING, w!("Preview Settings"));

    let _ = AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null());

    // Add the "Timing" submenu: whether the keyboard's turn holds the pointer back, how long
    // a hover waits before its preview opens, how long the same file is held off after its
    // preview was dismissed, how long the pointer must be still before anything previews at
    // all, and what the trigger key does.
    let timing_menu_guard = MenuGuard::new()?;
    let timing_menu = timing_menu_guard.handle();

    // Whether the keyboard driving Explorer holds a parked pointer back instead of the file
    // under it previewing: the first row here, the one switch among the submenu's delays and
    // keys, and it is on where the app starts.
    let prioritize_keyboard = snap.prioritize_keyboard;

    let _ = AppendMenuW(
        timing_menu,
        MF_STRING
            | if prioritize_keyboard {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_PRIORITIZE_KEYBOARD as usize,
        w!("Prioritize Keyboard"),
    );

    // Add the "Trigger Key (Alt)" submenu: whether the key is watched at all, the
    // key it watches, and what holding it does. Which of the two modes is active is
    // shown with radio marks, because only one of them can be.
    let trigger_key = snap.trigger_key.as_str();
    let trigger_key_mode = snap.trigger_key_mode;
    let trigger_key_enabled = snap.trigger_key_enabled;
    let trigger_menu_guard = MenuGuard::new()?;
    let trigger_menu = trigger_menu_guard.handle();

    let mut trigger_key_chars = trigger_key.chars();
    let trigger_key_display = match trigger_key_chars.next() {
        Some(first) => first.to_uppercase().collect::<String>() + trigger_key_chars.as_str(),
        None => trigger_key.to_string(),
    };
    let trigger_label = format!("Trigger Key ({trigger_key_display})");
    let trigger_label_wide: Vec<u16> = trigger_label
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();

    // The key itself first: with it off, nothing watches it, and the two items
    // below say what holding it would do.
    let _ = AppendMenuW(
        trigger_menu,
        MF_STRING
            | if trigger_key_enabled {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_TRIGGER_ENABLED as usize,
        w!("Enable Trigger Key"),
    );

    // The second switch: whether the key reaches a pin. Off, which is where it starts, a pinned
    // preview is not the key's to take down — the key is not read while one is up, collapsed into
    // its bubble or not — and the two rows below are about hovers alone.
    let trigger_key_affect_pin_mode = snap.trigger_key_affect_pin_mode;
    let _ = AppendMenuW(
        trigger_menu,
        MF_STRING
            | if trigger_key_affect_pin_mode {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_TRIGGER_AFFECT_PIN as usize,
        w!("Affect Pin Mode"),
    );
    let _ = AppendMenuW(trigger_menu, MF_SEPARATOR, 0, PCWSTR::null());

    let _ = AppendMenuW(
        trigger_menu,
        MF_STRING,
        ID_TRAY_TRIGGER_DISABLE as usize,
        w!("Hold to Disable Preview"),
    );
    let _ = AppendMenuW(
        trigger_menu,
        MF_STRING,
        ID_TRAY_TRIGGER_ENABLE as usize,
        w!("Hold to Enable Preview"),
    );
    let _ = CheckMenuRadioItem(
        trigger_menu,
        ID_TRAY_TRIGGER_DISABLE as u32,
        ID_TRAY_TRIGGER_ENABLE as u32,
        match trigger_key_mode {
            TriggerKeyMode::Disable => ID_TRAY_TRIGGER_DISABLE as u32,
            TriggerKeyMode::Enable => ID_TRAY_TRIGGER_ENABLE as u32,
        },
        MF_BYCOMMAND.0,
    );

    let _ =
        trigger_menu_guard.append_to(timing_menu, MF_STRING, PCWSTR(trigger_label_wide.as_ptr()));

    // Add the "Hover Delay" submenu: how long the pointer rests on a file
    // before a preview is put up for it.
    let hover_delay_ms = snap.hover_delay_ms;
    let delay_menu = timing_delay_menu(ID_TRAY_DELAY_BASE, hover_delay_ms, DEFAULT_HOVER_DELAY_MS)?;
    let _ = delay_menu.append_to(timing_menu, MF_STRING, w!("Hover Delay"));

    // Add the Rehover Delay submenu: how long the same file waits before a preview of it
    // is put up again.
    let same_file_rehover_delay_ms = snap.same_file_rehover_delay_ms;
    let rehover_delay_menu = timing_delay_menu(
        ID_TRAY_REHOVER_DELAY_BASE,
        same_file_rehover_delay_ms,
        DEFAULT_SAME_FILE_REHOVER_DELAY_MS,
    )?;
    let _ = rehover_delay_menu.append_to(timing_menu, MF_STRING, w!("Rehover Delay"));

    // Add the Settling Delay submenu: how long the pointer must be still before a
    // preview may open for anything. It is what a hand crossing a list waits out before
    // the file it comes to rest on is answered, and 0 is that requirement switched off.
    let settling_delay_ms = snap.settling_delay_ms;
    let settling_delay_menu = timing_delay_menu(
        ID_TRAY_SETTLING_DELAY_BASE,
        settling_delay_ms,
        DEFAULT_SETTLING_DELAY_MS,
    )?;
    let _ = settling_delay_menu.append_to(timing_menu, MF_STRING, w!("Settling Delay"));

    let _ = timing_menu_guard.append_to(menu, MF_STRING, w!("Timing"));

    // Add the "Placement" submenu: where a preview lands relative to the cursor or
    // the focused item.
    let placement_menu_guard = MenuGuard::new()?;
    let placement_menu = placement_menu_guard.handle();

    // Add the Position submenu: which side a preview takes. The two placements are one
    // setting shown two ways, so they carry a radio mark each.
    let follow_cursor = snap.follow_cursor;
    let position_menu_guard = MenuGuard::new()?;
    let position_menu = position_menu_guard.handle();

    // The two placements are one setting shown two ways, so the one it starts at carries
    // the default mark — read from that setting, which is what puts it on `Best Position`
    // while `follow_cursor` is false.
    let position_item = |id: u16, label: &str, follows_cursor: bool| {
        append_labeled_item(
            position_menu,
            MF_STRING,
            id,
            &default_label(label, follows_cursor == DEFAULT_FOLLOW_CURSOR),
        );
    };
    position_item(ID_TRAY_POSITION_FOLLOW, "Follow Cursor", true);
    position_item(ID_TRAY_POSITION_BEST, "Best Position", false);
    let _ = CheckMenuRadioItem(
        position_menu,
        ID_TRAY_POSITION_FOLLOW as u32,
        ID_TRAY_POSITION_BEST as u32,
        if follow_cursor {
            ID_TRAY_POSITION_FOLLOW as u32
        } else {
            ID_TRAY_POSITION_BEST as u32
        },
        MF_BYCOMMAND.0,
    );

    let _ = position_menu_guard.append_to(placement_menu, MF_STRING, w!("Position"));

    // Add the Avoid submenu: how far a preview is kept off the item it is about —
    // nothing, the item's name alone, the name with the column it is drawn in, or
    // every column the item draws. The four are one setting, so they carry a radio
    // mark each.
    let avoid_mode = snap.avoid_mode;

    append_avoid_menu(placement_menu, w!("Avoid"), ID_TRAY_AVOID_BASE, avoid_mode);

    let _ = placement_menu_guard.append_to(menu, MF_STRING, w!("Placement"));

    // Add the "Scaling" submenu below it: how large a preview of each kind is drawn,
    // rather than where it lands.
    let scaling_menu_guard = MenuGuard::new()?;
    let scaling_menu = scaling_menu_guard.handle();

    let bypass_placement = snap.bypass_placement;
    let _ = AppendMenuW(
        scaling_menu,
        MF_STRING
            | if bypass_placement {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_BYPASS_PLACEMENT as usize,
        w!("Bypass Placement"),
    );
    let _ = AppendMenuW(scaling_menu, MF_SEPARATOR, 0, PCWSTR::null());

    // Add the Image, Video, Animated and Audio submenus: how large
    // a picture, a video, a frame that moves and a sound's card are
    // drawn. The three that hold a bitmap of the file are each at a
    // share of the display's fitted size or of the file's own size —
    // one builder serves all three, the shares are the same shares, and
    // so is what a click on one means — and the sound's card is at a
    // share of the display, through the builder the document scales
    // use, because a card holds nothing of the sound to take a share
    // of. Each is a submenu of its own because the sizes are settings
    // of their own: what does not move, what plays, what moves inside
    // its frame, and what is heard are four questions.
    let preview_scale = snap.preview_scale;
    let video_scale = snap.video_scale;
    let audio_scale = snap.audio_scale;
    let animated_scale = snap.animated_scale;

    append_bitmap_scale_menu(
        scaling_menu,
        w!("Image"),
        ID_TRAY_SCALE_BASE,
        preview_scale,
        DEFAULT_PREVIEW_SCALE,
    );
    append_bitmap_scale_menu(
        scaling_menu,
        w!("Video"),
        ID_TRAY_VIDEO_SCALE_BASE,
        video_scale,
        DEFAULT_VIDEO_SCALE,
    );
    append_bitmap_scale_menu(
        scaling_menu,
        w!("Animated"),
        ID_TRAY_ANIMATED_SCALE_BASE,
        animated_scale,
        DEFAULT_ANIMATED_SCALE,
    );
    append_audio_scale_menu(
        scaling_menu,
        w!("Audio"),
        ID_TRAY_AUDIO_SCALE_BASE,
        audio_scale,
        DEFAULT_AUDIO_SCALE,
    );

    // Add the Vector, Text, Ebook, Document, Font, Design and 3D submenus:
    // how much of the display each kind of document is drawn over, or
    // measured in. They sit beside the picture scales because they are
    // the same question about other kinds of preview, and each is a
    // submenu of its own because the answers are not the same answers:
    // a picture's percentage is of its own size, a document's is of the
    // display — and a document and a page do not start at the same
    // share of it either. A model is the last of them because a scene is
    // a picture an engine drew rather than a file with a size of its
    // own, so how much of the display to give it is a question of its
    // own.
    //
    // One of them covers both halves of the `Document` kind, since a page
    // the render engine drew is a page like any other: what the setting
    // answers is how much of the display one is given, whichever engine
    // drew it.
    let ebook_scale = snap.ebook_scale;
    let document_scale = snap.document_scale;
    let font_scale = snap.font_scale;
    let design_scale = snap.design_scale;
    let vector_scale = snap.vector_scale;
    let text_scale = snap.text_scale;
    let model_scale = snap.model_scale;

    append_document_scale_menu(
        scaling_menu,
        w!("Vector"),
        ID_TRAY_VECTOR_SCALE_BASE,
        vector_scale,
        DEFAULT_VECTOR_SCALE,
    );
    append_document_scale_menu(
        scaling_menu,
        w!("Text"),
        ID_TRAY_TEXT_SCALE_BASE,
        text_scale,
        DEFAULT_TEXT_SCALE,
    );
    append_document_scale_menu(
        scaling_menu,
        w!("Ebook"),
        ID_TRAY_EBOOK_SCALE_BASE,
        ebook_scale,
        DEFAULT_EBOOK_SCALE,
    );
    append_document_scale_menu(
        scaling_menu,
        w!("Document"),
        ID_TRAY_DOCUMENT_SCALE_BASE,
        document_scale,
        DEFAULT_DOCUMENT_SCALE,
    );
    append_document_scale_menu(
        scaling_menu,
        w!("Font"),
        ID_TRAY_FONT_SCALE_BASE,
        font_scale,
        DEFAULT_FONT_SCALE,
    );
    append_document_scale_menu(
        scaling_menu,
        w!("Design"),
        ID_TRAY_DESIGN_SCALE_BASE,
        design_scale,
        DEFAULT_DESIGN_SCALE,
    );
    // A model is the one kind of preview that is a picture an engine drew of
    // something rather than the file itself, so how much of the display to give it is
    // asked through the same builder the documents above use.
    append_document_scale_menu(
        scaling_menu,
        w!("3D"),
        ID_TRAY_MODEL_SCALE_BASE,
        model_scale,
        DEFAULT_MODEL_SCALE,
    );

    let _ = scaling_menu_guard.append_to(menu, MF_STRING, w!("Scaling"));

    // Add the "Background" submenu: what a preview is drawn over, which is a question
    // a picture and a document answer differently — a picture's transparency is the
    // picture's, while a document is drawn on a page — so each of them has a half
    // of its own, listing the same backdrops.
    let image_background = snap.image_background;
    let vector_background = snap.vector_background;
    let html_background = snap.html_background;
    let font_background = snap.font_background;
    let dds_background = snap.dds_background;
    let design_background = snap.design_background;
    let background_menu_guard = MenuGuard::new()?;
    let background_menu = background_menu_guard.handle();

    // Each half is handed its own default, since the backdrops a half offers are not the same
    // ones everywhere: a picture, a drawing and a design document start at the squares, a
    // specimen, a page and a texture start at a page — and a texture is offered only the two
    // pages, a page all but the transparency. The submenu is named `Background` and each half
    // is named for the kind of preview it belongs to, so a half's label is the kind and
    // nothing else.
    append_background_menu(
        background_menu,
        w!("Image"),
        ID_TRAY_IMAGE_BACKGROUND_BASE,
        &BACKGROUND_CHOICES,
        image_background,
        DEFAULT_IMAGE_BACKGROUND,
    );
    append_background_menu(
        background_menu,
        w!("Vector"),
        ID_TRAY_VECTOR_BACKGROUND_BASE,
        &BACKGROUND_CHOICES,
        vector_background,
        DEFAULT_VECTOR_BACKGROUND,
    );
    append_background_menu(
        background_menu,
        w!("HTML"),
        ID_TRAY_HTML_BACKGROUND_BASE,
        &HTML_BACKGROUND_CHOICES,
        html_background,
        DEFAULT_HTML_BACKGROUND,
    );
    append_background_menu(
        background_menu,
        w!("Font"),
        ID_TRAY_FONT_BACKGROUND_BASE,
        &BACKGROUND_CHOICES,
        font_background,
        DEFAULT_FONT_BACKGROUND,
    );
    append_background_menu(
        background_menu,
        w!("DDS"),
        ID_TRAY_DDS_BACKGROUND_BASE,
        &DDS_BACKGROUND_CHOICES,
        dds_background,
        DEFAULT_DDS_BACKGROUND,
    );
    append_background_menu(
        background_menu,
        w!("Design"),
        ID_TRAY_DESIGN_BACKGROUND_BASE,
        &BACKGROUND_CHOICES,
        design_background,
        DEFAULT_DESIGN_BACKGROUND,
    );

    let _ = background_menu_guard.append_to(menu, MF_STRING, w!("Background"));

    // Add the "Volume" submenu, below `Background` and above `Performance`: what a film and a
    // sound are heard at, one half apiece, with each half's own two switches above its levels.
    // It is a block of its own rather than a row of either kind's half of `Preview Settings`
    // because the level is the whole of what a film and a sound answer the same way, and the
    // other rows of those halves — the film's subtitles, the loop, the pin's own — are not that
    // question at all (see `volume_menu`).
    let _ = volume_menu_guard.append_to(menu, MF_STRING, w!("Volume"));

    // Add the "Performance" submenu: what the app costs while it is working — what it holds on to
    // between hovers, on disk and in memory, and what one file may be answered for. It is the block
    // above `Engine`, where the engines themselves are named: what is here is what is spent while
    // they work. Its two largest children are `Disk Cache` and `Memory`, which is the one
    // distinction that decides which of them a setting belongs under.
    let performance_menu_guard = MenuGuard::new()?;
    let performance_menu = performance_menu_guard.handle();

    // Add the "Disk Cache" submenu: what a preview's own *files* may cost between hovers — the
    // pages a document was drawn as, which are files under the temp folder; the pictures the
    // image converter developed a file into, which are files of their own beside those; and the
    // small subtitle files a film's own embedded tracks were copied into, font attachments
    // included, which is what makes a hover of such a film a read of a few dozen kilobytes rather
    // than a stream of the whole film — each of them listed largest first, with the size its own
    // cache starts at marked. What a preview holds in *memory* is under `Memory` below it.
    let image_cache_mb = snap.image_cache_mb;
    let document_cache_mb = snap.document_cache_mb;
    let image_disk_cache_mb = snap.image_disk_cache_mb;
    let general_disk_cache_mb = snap.general_disk_cache_mb;

    let disk_cache_menu_guard = MenuGuard::new()?;
    let disk_cache_menu = disk_cache_menu_guard.handle();

    // One sizes menu, built for whichever cache is asking: `AppendMenuW` is handed a pointer, so
    // the wide strings have to outlive the call that lists them, and each size's label says the
    // one thing that differs between the caches — the size its own cache starts at.
    let sizes_menu =
        |base: u16, held: u32, default_mb: u32, choices: &[u32]| -> Option<MenuGuard> {
            let menu_guard = MenuGuard::new()?;
            let menu = menu_guard.handle();
            let labels: Vec<Vec<u16>> = choices
                .iter()
                .map(|megabytes| {
                    cache_size_label(*megabytes, default_mb)
                        .encode_utf16()
                        .chain(std::iter::once(0))
                        .collect()
                })
                .collect();

            for (index, _) in choices.iter().enumerate() {
                let _ = unsafe {
                    AppendMenuW(
                        menu,
                        MF_STRING,
                        (base + index as u16) as usize,
                        PCWSTR(labels[index].as_ptr()),
                    )
                };
            }

            // A size the menu does not offer — one a hand-edited `config.ini` asked for — leaves every
            // item unmarked rather than marking the nearest one.
            if let Some(index) = choices.iter().position(|megabytes| *megabytes == held) {
                let _ = unsafe {
                    CheckMenuRadioItem(
                        menu,
                        base as u32,
                        (base + choices.len() as u16 - 1) as u32,
                        (base + index as u16) as u32,
                        MF_BYCOMMAND.0,
                    )
                };
            }

            Some(menu_guard)
        };

    // The frames a decoded picture is shown as, which are held in RAM between hovers: the one
    // cache of the four that is not on disk, and the reason it is listed under `Memory` rather
    // than under `Disk Cache` beside the other three (see `memory_menu` below).
    let image_menu = sizes_menu(
        ID_TRAY_IMAGE_CACHE_BASE,
        image_cache_mb,
        DEFAULT_IMAGE_CACHE_MB,
        &IMAGE_CACHE_SIZE_CHOICES_MB[..],
    )?;

    for (name, base, held, default_mb, choices) in [
        (
            w!("Document"),
            ID_TRAY_DOCUMENT_CACHE_BASE,
            document_cache_mb,
            DEFAULT_DOCUMENT_CACHE_MB,
            &DISK_CACHE_SIZE_CHOICES_MB[..],
        ),
        (
            w!("Image"),
            ID_TRAY_IMAGE_DISK_CACHE_BASE,
            image_disk_cache_mb,
            DEFAULT_IMAGE_DISK_CACHE_MB,
            &DISK_CACHE_SIZE_CHOICES_MB[..],
        ),
        (
            w!("General"),
            ID_TRAY_GENERAL_DISK_CACHE_BASE,
            general_disk_cache_mb,
            DEFAULT_GENERAL_DISK_CACHE_MB,
            &DISK_CACHE_SIZE_CHOICES_MB[..],
        ),
    ] {
        let sizes = sizes_menu(base, held, default_mb, choices)?;
        let _ = sizes.append_to(disk_cache_menu, MF_STRING, name);
    }

    let _ = disk_cache_menu_guard.append_to(performance_menu, MF_STRING, w!("Disk Cache"));

    // Add the "Memory" submenu, which is everything here that is spent in RAM rather than on
    // disk: the frames a decoded image is shown as, which are held between hovers; the ceiling
    // one hover may decode or read *for*; and the room a pinned picture's held run of frames is
    // given, which is what a hand on its track is shown a frame of without walking the file (see
    // `pin_index`). The two ceilings were one item — `Decode Budget` — until a run of frames
    // needed a room of its own, and the name stays with the ceiling it has always meant.
    let decode_budget_gb = sanitize_decode_budget_gb(snap.decode_budget_gb);
    let animated_webp_budget_mb = sanitize_animated_webp_budget_mb(snap.animated_webp_budget_mb);

    let budget_menu_guard = MenuGuard::new()?;
    let budget_menu = budget_menu_guard.handle();

    let budget_labels: Vec<Vec<u16>> = DECODE_BUDGET_CHOICES_GB
        .iter()
        .map(|gigabytes| {
            decode_budget_label(*gigabytes)
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();

    for (index, _) in DECODE_BUDGET_CHOICES_GB.iter().enumerate() {
        let _ = AppendMenuW(
            budget_menu,
            MF_STRING,
            (ID_TRAY_DECODE_BUDGET_BASE + index as u16) as usize,
            PCWSTR(budget_labels[index].as_ptr()),
        );
    }

    // A ceiling the menu does not offer — one a hand-edited `config.ini` asked for — leaves every
    // item unmarked rather than marking the nearest one.
    if let Some(index) = DECODE_BUDGET_CHOICES_GB
        .iter()
        .position(|gigabytes| *gigabytes == decode_budget_gb)
    {
        let _ = CheckMenuRadioItem(
            budget_menu,
            ID_TRAY_DECODE_BUDGET_BASE as u32,
            (ID_TRAY_DECODE_BUDGET_BASE + DECODE_BUDGET_CHOICES_GB.len() as u16 - 1) as u32,
            (ID_TRAY_DECODE_BUDGET_BASE + index as u16) as u32,
            MF_BYCOMMAND.0,
        );
    }

    // And the room a pinned picture's run of frames is given, as a submenu of its own beside the
    // ceiling: it is what a drag of an animated picture spends, and what a machine with less
    // memory to spare turns down first.
    let animated_webp_menu_guard = MenuGuard::new()?;
    let animated_webp_menu = animated_webp_menu_guard.handle();
    let animated_webp_labels: Vec<Vec<u16>> = ANIMATED_WEBP_BUDGET_CHOICES_MB
        .iter()
        .map(|megabytes| {
            animated_webp_budget_label(*megabytes)
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();

    for (index, _) in ANIMATED_WEBP_BUDGET_CHOICES_MB.iter().enumerate() {
        let _ = AppendMenuW(
            animated_webp_menu,
            MF_STRING,
            (ID_TRAY_ANIMATED_WEBP_BUDGET_BASE + index as u16) as usize,
            PCWSTR(animated_webp_labels[index].as_ptr()),
        );
    }

    if let Some(index) = ANIMATED_WEBP_BUDGET_CHOICES_MB
        .iter()
        .position(|megabytes| *megabytes == animated_webp_budget_mb)
    {
        let _ = CheckMenuRadioItem(
            animated_webp_menu,
            ID_TRAY_ANIMATED_WEBP_BUDGET_BASE as u32,
            (ID_TRAY_ANIMATED_WEBP_BUDGET_BASE + ANIMATED_WEBP_BUDGET_CHOICES_MB.len() as u16 - 1)
                as u32,
            (ID_TRAY_ANIMATED_WEBP_BUDGET_BASE + index as u16) as u32,
            MF_BYCOMMAND.0,
        );
    }

    // What is held in memory between hovers and what one file may be answered for drawn into: the
    // ceiling one hover may decode or read *for* first, then a rule, then the two rooms a run of
    // frames is *kept* in — the frames a decoded picture is shown as, held between hovers, and the
    // run a pinned picture's track reads (see `pin_index`). The ceiling is above the two because
    // it decides whether a file is answered at all, which is the question a hand with a machine
    // that struggles reaches for first.
    let memory_menu_guard = MenuGuard::new()?;
    let memory_menu = memory_menu_guard.handle();
    let _ = budget_menu_guard.append_to(memory_menu, MF_STRING, w!("Decode Budget"));
    let _ = AppendMenuW(memory_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = image_menu.append_to(memory_menu, MF_STRING, w!("Image"));
    let _ = animated_webp_menu_guard.append_to(memory_menu, MF_STRING, w!("Animated Image"));

    let _ = memory_menu_guard.append_to(performance_menu, MF_STRING, w!("Memory"));

    // Add the "Explorer Poll" submenu: how often the loop looks at the
    // pointer's world while Explorer has focus. It is the app's own
    // rate rather than a hover's — the one number that trades how soon
    // a move is answered against what the app costs while it works —
    // which is why it is here and not with the delays a hover waits out.
    let tick_ms = snap.tick_ms;

    append_tick_menu(
        performance_menu,
        w!("Explorer Poll"),
        ID_TRAY_TICK_BASE,
        tick_ms,
    );

    let mpv_installed = VideoEngine::Mpv.installed();
    let video_hw_accel = snap.video_hw_accel;

    let _ = AppendMenuW(
        performance_menu,
        MF_STRING
            | if video_hw_accel && mpv_installed {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            }
            | if mpv_installed { MF_ENABLED } else { MF_GRAYED },
        ID_TRAY_VIDEO_HW_ACCEL as usize,
        w!("Hardware Acceleration (MPV Only)"),
    );

    let _ = performance_menu_guard.append_to(menu, MF_STRING, w!("Performance"));

    // Add the "Engine" submenu: which engine a preview is asked of, and how long each engine
    // this app starts is kept — the choice of engine where a kind of document has two of them,
    // and a TTL apiece for the applications this app leaves running between hovers. It is the
    // block below `Performance`, which is about what those engines cost while they are up.
    let engine_menu_guard = MenuGuard::new()?;
    let engine_menu = engine_menu_guard.handle();

    // Add the "Video Engine" and "Office Engine" submenus: which
    // engine each kind of document is asked of, where there is a
    // choice to make. They are the first of the rows that name an
    // engine, above the block that says how long each engine this app
    // starts is kept — the applications they name are the ones those
    // are about. Office is the only kind with two engines to choose
    // between; video has four, of which `Best` is the machine's own
    // answer.
    append_video_engine_menu(engine_menu, snap);
    append_office_engine_menu(engine_menu, snap);

    // Choices above, keeping times below.
    let _ = AppendMenuW(engine_menu, MF_SEPARATOR, 0, PCWSTR::null());

    // How long Explorer has to be out of reach before an engine set to `Away` is let go. It is
    // the second of the three blocks here, below the engine choices and above the TTLs, because
    // it is what the `Away` item of the four TTL submenus below is answered against: a time
    // those offer is what an engine given one is kept by, and this is what bounds the engines
    // that are not (see `app::afk`).
    append_afk_timer_menu(engine_menu, snap);

    // `Away Timer` bounds the engines kept by `Away`; the TTLs below bound the ones given a time.
    let _ = AppendMenuW(engine_menu, MF_SEPARATOR, 0, PCWSTR::null());

    let office_idle = snap.office_engine_idle;

    append_engine_idle_menu(
        engine_menu,
        w!("Microsoft Office TTL"),
        ID_TRAY_ENGINE_IDLE_BASE,
        office_idle,
        true,
    );

    let libreoffice_idle = snap.libreoffice_idle;

    append_engine_idle_menu(
        engine_menu,
        w!("LibreOffice TTL"),
        ID_TRAY_LIBREOFFICE_IDLE_BASE,
        libreoffice_idle,
        libreoffice_render::available(),
    );

    let webview_idle = snap.webview_idle;

    append_engine_idle_menu(
        engine_menu,
        w!("WebView2 TTL"),
        ID_TRAY_WEBVIEW_IDLE_BASE,
        webview_idle,
        webview_preview::is_available(),
    );

    let mpv_idle = snap.video_engine_idle;

    append_engine_idle_menu(
        engine_menu,
        w!("MPV TTL"),
        ID_TRAY_MPV_IDLE_BASE,
        mpv_idle,
        codecs::mpv_available(),
    );

    let f3d_idle = snap.f3d_idle;

    append_engine_idle_menu(
        engine_menu,
        w!("F3D TTL"),
        ID_TRAY_F3D_IDLE_BASE,
        f3d_idle,
        crate::formats::f3d_render::available(),
    );

    let _ = engine_menu_guard.append_to(menu, MF_STRING, w!("Engine"));

    refresh_codecs();
    append_codecs_menu(menu, snap);

    let custom_preview_menu_guard = MenuGuard::new()?;
    let custom_preview_menu = custom_preview_menu_guard.handle();
    let _ = AppendMenuW(
        custom_preview_menu,
        MF_STRING
            | if snap.custom_preview_enabled {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_CUSTOM_PREVIEW_ENABLE as usize,
        w!("Enable"),
    );
    let _ = AppendMenuW(
        custom_preview_menu,
        MF_STRING,
        ID_TRAY_CUSTOM_PREVIEW_EDIT as usize,
        w!("Edit"),
    );
    let _ = custom_preview_menu_guard.append_to(menu, MF_STRING, w!("Custom Preview"));

    let _ = AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null());

    let offered = snap.check_for_updates.then(updates::available).flatten();
    if let Some(version) = offered {
        let label = update_available_label(&version);
        let label_wide: Vec<u16> = label.encode_utf16().chain(std::iter::once(0)).collect();
        let _ = AppendMenuW(
            menu,
            MF_STRING,
            ID_TRAY_UPDATE as usize,
            PCWSTR(label_wide.as_ptr()),
        );
    }

    let system_menu_guard = MenuGuard::new()?;
    let system_menu = system_menu_guard.handle();

    let check_for_updates = snap.check_for_updates;
    let silent_auto_update = snap.silent_auto_update;
    let _ = AppendMenuW(
        system_menu,
        MF_STRING
            | if check_for_updates {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_CHECK_UPDATES as usize,
        w!("Check for Updates"),
    );

    // And whether an update any check finds is put on without asking — but only where
    // `updates::the_app_is_idle` holds, so nothing of the user's is closed by it; what a check finds
    // over a pin or an open menu is owed until the screen is empty. It is greyed while the checks are
    // off rather than hidden: it is a question about a check that is not being made, and a switch
    // that is there and does nothing says why.
    let silent_flags = MF_STRING
        | if silent_auto_update {
            MF_CHECKED
        } else {
            MF_UNCHECKED
        }
        | if check_for_updates {
            MF_ENABLED
        } else {
            MF_GRAYED
        };
    let _ = AppendMenuW(
        system_menu,
        silent_flags,
        ID_TRAY_SILENT_AUTO_UPDATE as usize,
        w!("Silent Auto-Update"),
    );

    // Add "Run at Startup" with checkmark: what starts this app is the entry, and the two can be
    // made to differ from outside this app. **The entry is asked twice, because Windows keeps two
    // facts about it**: the value under `Run` is the entry, and the flag under `StartupApproved\\Run`
    // is whether a logon is allowed to run it — an entry turned off in Task Manager, or in
    // Settings → Apps → Startup, is a value that is there and starts nothing, and a checkmark
    // drawn from the value alone would be a yes on a machine where nothing happens. What is left of
    // that case is said in the label, since that entry is one click away from being put back.
    let startup_present = startup::is_entry_present();
    let startup_enabled = startup_present && startup::is_entry_enabled();
    // And the one setting this app cannot undo from here: an executable marked *Run this program as
    // an administrator* is one Windows will never start from a `Run` value, because it never
    // elevates one. The row says so rather than showing a checkmark over an entry that does nothing.
    let startup_elevated = startup::entry_would_ask_for_elevation();
    let flags = MF_STRING
        | if startup_enabled {
            MF_CHECKED
        } else {
            MF_UNCHECKED
        };
    let label = if startup_elevated {
        "Run at Startup (needs Run as administrator turned off)"
    } else if startup_present && !startup_enabled {
        "Run at Startup (turned off in Windows)"
    } else {
        "Run at Startup"
    };
    let label_wide: Vec<u16> = label.encode_utf16().chain(std::iter::once(0)).collect();
    let _ = AppendMenuW(
        system_menu,
        flags,
        ID_TRAY_STARTUP as usize,
        PCWSTR(label_wide.as_ptr()),
    );

    let _ = AppendMenuW(system_menu, MF_SEPARATOR, 0, PCWSTR::null());

    // And the theme, between two separators of its own: the chrome's own look rather than a
    // text setting, which is what it was a row of the top menu beside the text preview for.
    // It sits here because the two rows above are the app's own place in the machine and the
    // two below are its own file, and neither is about how anything is drawn — the look is
    // the app's own rather than a preview's, so it belongs with the rows about the app.
    let _ = theme_menu_guard.append_to(system_menu, MF_STRING, w!("Theme"));

    let _ = AppendMenuW(system_menu, MF_SEPARATOR, 0, PCWSTR::null());

    // The configuration's own file: it opens with the default app for it,
    // and no one is asked first.
    let _ = AppendMenuW(
        system_menu,
        MF_STRING,
        ID_TRAY_OPEN_CONFIG as usize,
        w!("Edit Config.ini"),
    );

    // What each of the two would change, which is also the answer to whether it is offered:
    // a reset with nothing behind it is greyed rather than shown as a click that would do
    // nothing, and the question it asks is the difference counted here.
    let settings_apart = snap.settings_apart_from_recommended();
    let lists_apart = snap.lists_apart_from_built_in();

    let settings_flags = if settings_apart.is_empty() {
        MF_STRING | MF_GRAYED
    } else {
        MF_STRING
    };
    let _ = AppendMenuW(
        system_menu,
        settings_flags,
        ID_TRAY_RESET_SETTINGS as usize,
        w!("Reset to Recommended Settings..."),
    );

    let lists_flags = if lists_apart.is_empty() {
        MF_STRING | MF_GRAYED
    } else {
        MF_STRING
    };
    let _ = AppendMenuW(
        system_menu,
        lists_flags,
        ID_TRAY_RESET_LISTS as usize,
        w!("Reset Extension Lists..."),
    );

    let system_label = system_menu_label(env!("CARGO_PKG_VERSION"));
    let system_label_wide: Vec<u16> = system_label
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();
    let _ = system_menu_guard.append_to(menu, MF_STRING, PCWSTR(system_label_wide.as_ptr()));

    // Add Exit
    let _ = AppendMenuW(menu, MF_STRING, ID_TRAY_EXIT as usize, w!("Exit"));

    Some(())
}

/// What the `Subtitles` row is called: `Subtitles (MPV)` where mpv is the engine that will play —
/// the pick itself, or `Best` on a machine that has mpv to prefer — and the words the row has
/// always carried for every other answer, because there the app is the one probing a film and
/// copying a track out of it.
///
/// mpv shows a film's own tracks and the `.srt` beside it by itself, so under it the row is not
/// allowed to claim work that is not done. Whether mpv is here is handed in rather than asked
/// here, the way the route's own preference is handed its answers (see
/// `video_hw::resolve_video_engine`), so that both answers are read on this machine's terms.
///
/// shortcut: the configured engine, not the file being hovered; upgrade if a per-file label is ever wanted
pub(super) fn subtitles_label(engine: VideoEngine, mpv_installed: bool) -> String {
    let mpv_plays = engine == VideoEngine::Mpv || (engine == VideoEngine::Best && mpv_installed);

    if mpv_plays {
        "Subtitles (MPV)".to_string()
    } else {
        "Subtitles (Expensive)".to_string()
    }
}

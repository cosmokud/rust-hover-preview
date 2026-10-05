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
    EngineIdleIds, BACKGROUND_CHOICES, CACHE_SIZE_CHOICES_MB, DDS_BACKGROUND_CHOICES,
    DECODE_BUDGET_CHOICES_GB, FONT_SIZE_CHOICES, HTML_BACKGROUND_CHOICES,
    ID_TRAY_ANIMATED_SCALE_BASE, ID_TRAY_AUDIO_VOLUME_BASE, ID_TRAY_AVOID_BASE,
    ID_TRAY_DDS_BACKGROUND_BASE, ID_TRAY_DECODE_BUDGET_BASE, ID_TRAY_DELAY_BASE,
    ID_TRAY_DESIGN_BACKGROUND_BASE, ID_TRAY_DESIGN_SCALE_BASE, ID_TRAY_DOCUMENT_CACHE_BASE,
    ID_TRAY_DOCUMENT_SCALE_BASE, ID_TRAY_EBOOK_SCALE_BASE, ID_TRAY_ENABLE,
    ID_TRAY_ENGINE_IDLE_BASE, ID_TRAY_ENGINE_PERSISTENT_BASE, ID_TRAY_EXIT,
    ID_TRAY_FONT_BACKGROUND_BASE, ID_TRAY_FONT_SCALE_BASE, ID_TRAY_GENERAL_DISK_CACHE_BASE,
    ID_TRAY_HTML_BACKGROUND_BASE, ID_TRAY_IMAGE_BACKGROUND_BASE, ID_TRAY_IMAGE_CACHE_BASE,
    ID_TRAY_IMAGE_DISK_CACHE_BASE, ID_TRAY_LIBREOFFICE_IDLE_BASE, ID_TRAY_MARKDOWN_RENDERED,
    ID_TRAY_MARKDOWN_SOURCE, ID_TRAY_NORMALIZE_VIDEO_VOLUME, ID_TRAY_NORMALIZE_VOLUME,
    ID_TRAY_OPEN_CONFIG, ID_TRAY_PIN, ID_TRAY_PIN_NAV_ALL, ID_TRAY_PIN_NAV_CATEGORY,
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
    ID_TRAY_VECTOR_BACKGROUND_BASE, ID_TRAY_VECTOR_SCALE_BASE, ID_TRAY_VIDEO_HW_ACCEL,
    ID_TRAY_VIDEO_SCALE_BASE, ID_TRAY_VIDEO_VOLUME_BASE, ID_TRAY_WEBVIEW_IDLE_BASE,
    MAX_TRAY_CUSTOM_THEMES, TRAY_CUSTOM_THEMES,
};
use super::submenus::{
    append_afk_timer_menu, append_audio_seek_menu, append_avoid_menu, append_background_menu,
    append_bitmap_scale_menu, append_codecs_menu, append_document_scale_menu,
    append_engine_idle_menu, append_labeled_item, append_select_engine_menu, append_tick_menu,
    cache_size_label, decode_budget_label, default_label, timing_delay_menu,
};

use crate::app::updates;
use crate::config::config::{
    sanitize_decode_budget_gb, sanitize_text_font_scale_percent, EngineIdle, MarkdownMode,
    PinNavFileTypes, PreviewType, TextTheme, TriggerKeyMode, DEFAULT_ANIMATED_SCALE,
    DEFAULT_AUDIO_SEEK, DEFAULT_AUDIO_VOLUME, DEFAULT_AVOID_MODE, DEFAULT_DDS_BACKGROUND,
    DEFAULT_DECODE_BUDGET_GB, DEFAULT_DESIGN_BACKGROUND, DEFAULT_DESIGN_SCALE,
    DEFAULT_DOCUMENT_CACHE_MB, DEFAULT_DOCUMENT_SCALE, DEFAULT_EBOOK_SCALE, DEFAULT_FOLLOW_CURSOR,
    DEFAULT_FONT_BACKGROUND, DEFAULT_FONT_SCALE, DEFAULT_GENERAL_DISK_CACHE_MB,
    DEFAULT_HOVER_DELAY_MS, DEFAULT_HTML_BACKGROUND, DEFAULT_IMAGE_BACKGROUND,
    DEFAULT_IMAGE_CACHE_MB, DEFAULT_IMAGE_DISK_CACHE_MB, DEFAULT_LIBREOFFICE_IDLE_SECS,
    DEFAULT_NORMALIZE_VIDEO_VOLUME, DEFAULT_NORMALIZE_VOLUME, DEFAULT_OFFICE_ENGINE_IDLE_SECS,
    DEFAULT_PIN_NAV_FILE_TYPES, DEFAULT_PIN_PAUSE_AUDIO, DEFAULT_PIN_PAUSE_VIDEO,
    DEFAULT_PIN_UPDATE_ENABLED, DEFAULT_PIN_UPDATE_ON_HOVER, DEFAULT_PREVIEW_SCALE,
    DEFAULT_REMEMBER_AUDIO_VOLUME, DEFAULT_REMEMBER_VIDEO_VOLUME, DEFAULT_RENDER_HTML,
    DEFAULT_SAME_FILE_REHOVER_DELAY_MS, DEFAULT_SETTLING_DELAY_MS, DEFAULT_TEXT_FONT_SCALE_PERCENT,
    DEFAULT_TEXT_SCALE, DEFAULT_TICK_MS, DEFAULT_TRIGGER_KEY_AFFECT_PIN_MODE,
    DEFAULT_VECTOR_BACKGROUND, DEFAULT_VECTOR_SCALE, DEFAULT_VIDEO_HW_ACCEL, DEFAULT_VIDEO_SCALE,
    DEFAULT_VIDEO_VOLUME, DEFAULT_WEBVIEW_IDLE_SECS, VOLUME_CHOICES,
};
use crate::config::theme_files;
use crate::engines::libreoffice_render;
use crate::engines::webview_preview;
use crate::formats::codecs::{self, refresh as refresh_codecs};
use crate::text::text_theme;
use crate::{app::startup, CONFIG};
use windows::core::{w, PCWSTR, PWSTR};
use windows::Win32::Foundation::{BOOL, HWND};
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CheckMenuRadioItem, CreatePopupMenu, DestroyMenu, GetCursorPos, GetMenuItemCount,
    InsertMenuItemW, SetForegroundWindow, TrackPopupMenu, HMENU, MENUITEMINFOW, MFT_STRING,
    MF_BYCOMMAND, MF_CHECKED, MF_GRAYED, MF_POPUP, MF_SEPARATOR, MF_STRING, MF_UNCHECKED, MIIM_ID,
    MIIM_STRING, MIIM_SUBMENU, TPM_BOTTOMALIGN, TPM_LEFTALIGN,
};

pub(super) unsafe fn show_context_menu(hwnd: HWND) {
    let menu = CreatePopupMenu().unwrap();
    if let Ok(mut config) = CONFIG.lock() {
        config.reload_from_disk();
    }

    // The menu is also where the check for a newer release is asked for, after the one
    // the app made as it started: opening the menu is the one moment a user is looking
    // for one, and a check asked for here keeps what is offered current however long this
    // run has been up. It costs nothing here — the check runs on a thread of its own and
    // is answered at most once an hour — and the row above `Run at Startup` reports what
    // the last one found, so an update published since that check is offered on the
    // opening after this one.
    updates::request_check();

    // Add "Enable Preview" with checkmark
    let preview_enabled = CONFIG.lock().map(|c| c.preview_enabled).unwrap_or(true);
    let enable_flags = MF_STRING
        | if preview_enabled {
            MF_CHECKED
        } else {
            MF_UNCHECKED
        };
    let _ = AppendMenuW(
        menu,
        enable_flags,
        ID_TRAY_ENABLE as usize,
        w!("Enable Preview"),
    );

    // And the "Pin Mode" submenu below it, which is the pin's own business: whether the key is
    // watched at all, whether a pin that is up is shown another file while the user picks one,
    // and what a pin collapsed into its bubble does with what it is playing. The key's row names
    // the key that pins the way the trigger key's own submenu names its own: the key is a
    // setting, so the row says which one is watched rather than only whether one is (see
    // `key_input`).
    let (
        pin_enabled,
        pin_key,
        pin_pause_audio,
        pin_pause_video,
        pin_update,
        pin_update_on_hover,
        pin_nav_file_types,
    ) = CONFIG
        .lock()
        .map(|c| {
            (
                c.pin_enabled,
                c.pin_key.clone(),
                c.pin_pause_audio,
                c.pin_pause_video,
                c.pin_update_enabled,
                c.pin_update_on_hover,
                c.pin_nav_file_types,
            )
        })
        .unwrap_or((
            true,
            "space".to_string(),
            DEFAULT_PIN_PAUSE_AUDIO,
            DEFAULT_PIN_PAUSE_VIDEO,
            DEFAULT_PIN_UPDATE_ENABLED,
            DEFAULT_PIN_UPDATE_ON_HOVER,
            DEFAULT_PIN_NAV_FILE_TYPES,
        ));
    let mut pin_key_chars = pin_key.chars();
    let pin_key_display = match pin_key_chars.next() {
        Some(first) => first.to_uppercase().collect::<String>() + pin_key_chars.as_str(),
        None => pin_key,
    };
    let pin_label = format!("Enable ({pin_key_display})");
    let pin_label_wide: Vec<u16> = pin_label.encode_utf16().chain(std::iter::once(0)).collect();
    let pin_menu = CreatePopupMenu().unwrap();
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

    // The `Update Preview` submenu: whether a pin that is up is shown the file the user picks
    // next — the one the pointer clicks, or the one the keyboard selects — and whether the
    // pointer's own hover is one of the ways it is told about one. A pin is a window to read
    // or to watch, so both are off the pin's own row rather than part of it.
    //
    // The second row is greyed while the first is off, because it says nothing then: a pin
    // that follows nothing does not follow a hover either, and a row that could be ticked
    // without effect would be a setting a user cannot tell from one that does nothing.
    let update_menu = CreatePopupMenu().unwrap();
    let _ = AppendMenuW(
        update_menu,
        MF_STRING | if pin_update { MF_CHECKED } else { MF_UNCHECKED },
        ID_TRAY_PIN_UPDATE as usize,
        w!("Enabled"),
    );
    let mut hover_flags = MF_STRING
        | if pin_update_on_hover {
            MF_CHECKED
        } else {
            MF_UNCHECKED
        };
    if !pin_update {
        hover_flags |= MF_GRAYED;
    }
    // A separator between the two, because they are not the same kind of row: the first says
    // whether the pin follows the selection at all, the second only whether a hover is one of
    // the ways that selection is made.
    let _ = AppendMenuW(update_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = AppendMenuW(
        update_menu,
        hover_flags,
        ID_TRAY_PIN_UPDATE_HOVER as usize,
        w!("On Hover"),
    );
    let _ = AppendMenuW(
        pin_menu,
        MF_STRING | MF_POPUP,
        update_menu.0 as usize,
        w!("Update Preview"),
    );

    // The `Pause Preview` submenu: whether a pin collapsed into its bubble holds what it is
    // playing where it is until the pin is put back up again. The two are switches of their own
    // because a video and a sound are two different things to want quiet — a film a user wants to
    // go on hearing while the bubble is up is not a reason to let a podcast play on, and the other
    // way round — and both are on, since a bubble is a pin put away and what it was playing is not
    // what the desktop was asked for. A sound is listed above a video.
    let pause_menu = CreatePopupMenu().unwrap();
    let _ = AppendMenuW(
        pause_menu,
        MF_STRING
            | if pin_pause_audio {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_PIN_PAUSE_AUDIO as usize,
        w!("Audio"),
    );
    let _ = AppendMenuW(
        pause_menu,
        MF_STRING
            | if pin_pause_video {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_PIN_PAUSE_VIDEO as usize,
        w!("Video"),
    );
    let _ = AppendMenuW(
        pin_menu,
        MF_STRING | MF_POPUP,
        pause_menu.0 as usize,
        w!("Pause Preview"),
    );

    // The `Nav File Types` submenu: what the pin's own previous/next buttons step through —
    // every file this build can preview, or only those of the kind of thing the pinned file
    // is. It is a question about the walk rather than about the pin, which is why it hangs
    // under the same row as the two switches above it and not on its own: the buttons are
    // the pin's, and what they are buttons *of* is the whole of the setting.
    //
    // The two answers are one of two rather than a switch, because a folder of mixed work is
    // what a hand is most often looking at and the narrow walk is not obviously the better
    // one — so which walk a pin takes is asked rather than assumed (see `PinNavFileTypes`).
    let nav_types_menu = CreatePopupMenu().unwrap();
    let nav_types_flag = |candidate: PinNavFileTypes| {
        MF_STRING
            | if pin_nav_file_types == candidate {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            }
    };
    let _ = AppendMenuW(
        nav_types_menu,
        nav_types_flag(PinNavFileTypes::All),
        ID_TRAY_PIN_NAV_ALL as usize,
        w!("All"),
    );
    let _ = AppendMenuW(
        nav_types_menu,
        nav_types_flag(PinNavFileTypes::Category),
        ID_TRAY_PIN_NAV_CATEGORY as usize,
        w!("Category"),
    );
    let _ = AppendMenuW(
        pin_menu,
        MF_STRING | MF_POPUP,
        nav_types_menu.0 as usize,
        w!("Nav File Types"),
    );

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        pin_menu.0 as usize,
        w!("Pin Mode"),
    );

    // Add the "Preview Types" submenu: one gate per kind of preview, on by
    // default. A gate is only whether previews of that kind may be shown at all —
    // the lists and settings that decide which files of that kind preview are left
    // alone, so switching one off and back on restores what was configured.
    //
    // One gate covers each pair of a kind this app reads and the kind an engine draws for
    // it: an ImageMagick picture is a picture, a document LibreOffice drew is a document,
    // an archive PeaZip listed is an archive, and a book Calibre converted is a book, so
    // none of the four has a row here.
    let kinds = [
        (PreviewType::Images, ID_TRAY_TYPE_IMAGES, w!("Images")),
        (PreviewType::Videos, ID_TRAY_TYPE_VIDEOS, w!("Videos")),
        (PreviewType::Audio, ID_TRAY_TYPE_AUDIO, w!("Audio")),
        (PreviewType::Text, ID_TRAY_TYPE_TEXT, w!("Text")),
        (PreviewType::Ebook, ID_TRAY_TYPE_EBOOK, w!("Ebook")),
        (PreviewType::Archives, ID_TRAY_TYPE_ARCHIVES, w!("Archives")),
        (PreviewType::Document, ID_TRAY_TYPE_DOCUMENT, w!("Document")),
        (PreviewType::Vector, ID_TRAY_TYPE_VECTOR, w!("Vector")),
        (PreviewType::Fonts, ID_TRAY_TYPE_FONTS, w!("Fonts")),
        (PreviewType::Design, ID_TRAY_TYPE_DESIGN, w!("Design")),
    ];
    let types_menu = CreatePopupMenu().unwrap();

    for (kind, id, label) in kinds {
        let flags = MF_STRING
            | if kind.enabled() {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            };
        let _ = AppendMenuW(types_menu, flags, id as usize, label);
    }

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        types_menu.0 as usize,
        w!("Preview Types"),
    );

    // Add the "Text Preview" submenu: the theme, size and Markdown mode a text preview is
    // painted with. What a text preview *is* rather than what it shows — the scrollbar, the
    // selection and the keys that copy it — is the pin's business rather than a setting's: a
    // text preview pinned with the key becomes one to work in (see the pin in `preview_window`).
    let text_menu = CreatePopupMenu().unwrap();

    // Add the Theme submenu
    let theme = CONFIG.lock().map(|c| c.theme).unwrap_or(TextTheme::Light);

    // The folder is read here rather than once at startup: the files these names
    // stand for are the user's to add to and to edit, and this is the moment they
    // are looking at them.
    text_theme::refresh_custom();
    let mut custom_themes = theme_files::names();
    custom_themes.truncate(MAX_TRAY_CUSTOM_THEMES);

    let theme_menu = CreatePopupMenu().unwrap();

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
                MF_CHECKED
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

    let _ = AppendMenuW(
        text_menu,
        MF_STRING | MF_POPUP,
        theme_menu.0 as usize,
        w!("Theme"),
    );

    // Add the Font Size submenu, largest first: the steps a size can be picked
    // from, from the largest down to the smallest. A hand-edited size between these
    // steps simply matches none of them, which is why the values are read as
    // written.
    let font_scale = CONFIG
        .lock()
        .map(|c| sanitize_text_font_scale_percent(c.text_font_scale_percent))
        .unwrap_or(DEFAULT_TEXT_FONT_SCALE_PERCENT);
    let font_menu = CreatePopupMenu().unwrap();

    let font_flag = |percent: u32| {
        MF_STRING
            | if font_scale == percent {
                MF_CHECKED
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
    let _ = AppendMenuW(
        text_menu,
        MF_STRING | MF_POPUP,
        font_menu.0 as usize,
        w!("Font Size"),
    );

    // Add the Markdown submenu
    let markdown_mode = CONFIG
        .lock()
        .map(|c| c.markdown_mode)
        .unwrap_or(MarkdownMode::Rendered);
    let markdown_menu = CreatePopupMenu().unwrap();

    let markdown_flag = |candidate: MarkdownMode| {
        MF_STRING
            | if markdown_mode == candidate {
                MF_CHECKED
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
    let _ = AppendMenuW(
        text_menu,
        MF_STRING | MF_POPUP,
        markdown_menu.0 as usize,
        w!("Markdown"),
    );

    // Whether a page of HTML is drawn by the browser engine rather than shown as its markup.
    // The whole menu is built from the configuration each time it is opened, so this is read
    // here rather than kept between the click and the row.
    let render_html = CONFIG
        .lock()
        .map(|c| c.render_html)
        .unwrap_or(DEFAULT_RENDER_HTML);
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

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        text_menu.0 as usize,
        w!("Text Preview"),
    );

    let _ = AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null());

    // Add the "Timing" submenu: whether the keyboard's turn holds the pointer back, how long
    // a hover waits before its preview opens, how long the same file is held off after its
    // preview was dismissed, how long the pointer must be still before anything previews at
    // all, and what the trigger key does.
    let timing_menu = CreatePopupMenu().unwrap();

    // Whether the keyboard driving Explorer holds a parked pointer back instead of the file
    // under it previewing: the first row here, the one switch among the submenu's delays and
    // keys, and it is on where the app starts.
    let prioritize_keyboard = CONFIG.lock().map(|c| c.prioritize_keyboard).unwrap_or(true);

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
    let (trigger_key, trigger_key_mode, trigger_key_enabled) = CONFIG
        .lock()
        .map(|c| {
            (
                c.trigger_key.clone(),
                c.trigger_key_mode,
                c.trigger_key_enabled,
            )
        })
        .unwrap_or(("alt".to_string(), TriggerKeyMode::Disable, true));
    let trigger_menu = CreatePopupMenu().unwrap();

    let mut trigger_key_chars = trigger_key.chars();
    let trigger_key_display = match trigger_key_chars.next() {
        Some(first) => first.to_uppercase().collect::<String>() + trigger_key_chars.as_str(),
        None => trigger_key,
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
    let trigger_key_affect_pin_mode = CONFIG
        .lock()
        .map(|c| c.trigger_key_affect_pin_mode)
        .unwrap_or(DEFAULT_TRIGGER_KEY_AFFECT_PIN_MODE);
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

    let _ = AppendMenuW(
        timing_menu,
        MF_STRING | MF_POPUP,
        trigger_menu.0 as usize,
        PCWSTR(trigger_label_wide.as_ptr()),
    );

    // Add the Delay submenu: how long the pointer rests on a file before a preview is
    // put up for it.
    let hover_delay_ms = CONFIG
        .lock()
        .map(|c| c.hover_delay_ms)
        .unwrap_or(DEFAULT_HOVER_DELAY_MS);
    let delay_menu = timing_delay_menu(ID_TRAY_DELAY_BASE, hover_delay_ms, DEFAULT_HOVER_DELAY_MS);

    let _ = AppendMenuW(
        timing_menu,
        MF_STRING | MF_POPUP,
        delay_menu.0 as usize,
        w!("Delay"),
    );

    // Add the Rehover Delay submenu: how long the same file waits before a preview of it
    // is put up again.
    let same_file_rehover_delay_ms = CONFIG
        .lock()
        .map(|c| c.same_file_rehover_delay_ms)
        .unwrap_or(DEFAULT_SAME_FILE_REHOVER_DELAY_MS);
    let rehover_delay_menu = timing_delay_menu(
        ID_TRAY_REHOVER_DELAY_BASE,
        same_file_rehover_delay_ms,
        DEFAULT_SAME_FILE_REHOVER_DELAY_MS,
    );

    let _ = AppendMenuW(
        timing_menu,
        MF_STRING | MF_POPUP,
        rehover_delay_menu.0 as usize,
        w!("Rehover Delay"),
    );

    // Add the Settling Delay submenu: how long the pointer must be still before a
    // preview may open for anything. It is what a hand crossing a list waits out before
    // the file it comes to rest on is answered, and 0 is that requirement switched off.
    let settling_delay_ms = CONFIG
        .lock()
        .map(|c| c.settling_delay_ms)
        .unwrap_or(DEFAULT_SETTLING_DELAY_MS);
    let settling_delay_menu = timing_delay_menu(
        ID_TRAY_SETTLING_DELAY_BASE,
        settling_delay_ms,
        DEFAULT_SETTLING_DELAY_MS,
    );

    let _ = AppendMenuW(
        timing_menu,
        MF_STRING | MF_POPUP,
        settling_delay_menu.0 as usize,
        w!("Settling Delay"),
    );

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        timing_menu.0 as usize,
        w!("Timing"),
    );

    // Add the "Placement" submenu: where a preview lands relative to the cursor or
    // the focused item.
    let placement_menu = CreatePopupMenu().unwrap();

    // Add the Position submenu: which side a preview takes. The two placements are one
    // setting shown two ways, so they carry a radio mark each.
    let follow_cursor = CONFIG
        .lock()
        .map(|c| c.follow_cursor)
        .unwrap_or(DEFAULT_FOLLOW_CURSOR);
    let position_menu = CreatePopupMenu().unwrap();

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

    let _ = AppendMenuW(
        placement_menu,
        MF_STRING | MF_POPUP,
        position_menu.0 as usize,
        w!("Position"),
    );

    // Add the Avoid submenu: how far a preview is kept off the item it is about —
    // nothing, the item's name alone, or the name with the columns a row draws beside
    // it. The three are one setting, so they carry a radio mark each.
    let avoid_mode = CONFIG
        .lock()
        .map(|c| c.avoid_mode)
        .unwrap_or(DEFAULT_AVOID_MODE);

    append_avoid_menu(placement_menu, w!("Avoid"), ID_TRAY_AVOID_BASE, avoid_mode);

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        placement_menu.0 as usize,
        w!("Placement"),
    );

    // Add the "Scaling" submenu below it: how large a preview of each kind is drawn,
    // rather than where it lands.
    let scaling_menu = CreatePopupMenu().unwrap();

    // Add the Image Scaling, Video Scaling and Animated Scaling submenus: how large a
    // picture, a video and an animated picture is drawn, each at a share of its own size
    // rather than of the display. One builder serves all three — the shares are the same
    // shares, and so is what a click on one means — and each is a submenu of its own
    // because the sizes are settings of their own: what does not move, what plays, and
    // what moves inside its frame are three questions.
    let (preview_scale, video_scale, animated_scale) = CONFIG
        .lock()
        .map(|c| (c.preview_scale, c.video_scale, c.animated_scale))
        .unwrap_or((
            DEFAULT_PREVIEW_SCALE,
            DEFAULT_VIDEO_SCALE,
            DEFAULT_ANIMATED_SCALE,
        ));

    append_bitmap_scale_menu(
        scaling_menu,
        w!("Image Scaling"),
        ID_TRAY_SCALE_BASE,
        preview_scale,
        DEFAULT_PREVIEW_SCALE,
    );
    append_bitmap_scale_menu(
        scaling_menu,
        w!("Video Scaling"),
        ID_TRAY_VIDEO_SCALE_BASE,
        video_scale,
        DEFAULT_VIDEO_SCALE,
    );
    append_bitmap_scale_menu(
        scaling_menu,
        w!("Animated Scaling"),
        ID_TRAY_ANIMATED_SCALE_BASE,
        animated_scale,
        DEFAULT_ANIMATED_SCALE,
    );

    // Add the Vector Scaling, Text Scaling, Ebook Scaling, Document Scaling, Font Scaling and
    // Design Scaling submenus: how much of the display each kind of document is drawn over, or
    // measured in. They sit beside the picture scale because they are the same question about
    // other kinds of preview, and each is a submenu of its own because the answers are not the
    // same answers: a picture's percentage is of its own size, a document's is of the display —
    // and a document and a page do not start at the same share of it either.
    //
    // One of them covers both halves of the `Document` kind, since a page the render engine
    // drew is a page like any other: what the setting answers is how much of the display one
    // is given, whichever engine drew it.
    let (ebook_scale, document_scale, font_scale, design_scale, vector_scale, text_scale) = CONFIG
        .lock()
        .map(|c| {
            (
                c.ebook_scale,
                c.document_scale,
                c.font_scale,
                c.design_scale,
                c.vector_scale,
                c.text_scale,
            )
        })
        .unwrap_or((
            DEFAULT_EBOOK_SCALE,
            DEFAULT_DOCUMENT_SCALE,
            DEFAULT_FONT_SCALE,
            DEFAULT_DESIGN_SCALE,
            DEFAULT_VECTOR_SCALE,
            DEFAULT_TEXT_SCALE,
        ));

    append_document_scale_menu(
        scaling_menu,
        w!("Vector Scaling"),
        ID_TRAY_VECTOR_SCALE_BASE,
        vector_scale,
        DEFAULT_VECTOR_SCALE,
    );
    append_document_scale_menu(
        scaling_menu,
        w!("Text Scaling"),
        ID_TRAY_TEXT_SCALE_BASE,
        text_scale,
        DEFAULT_TEXT_SCALE,
    );
    append_document_scale_menu(
        scaling_menu,
        w!("Ebook Scaling"),
        ID_TRAY_EBOOK_SCALE_BASE,
        ebook_scale,
        DEFAULT_EBOOK_SCALE,
    );
    append_document_scale_menu(
        scaling_menu,
        w!("Document Scaling"),
        ID_TRAY_DOCUMENT_SCALE_BASE,
        document_scale,
        DEFAULT_DOCUMENT_SCALE,
    );
    append_document_scale_menu(
        scaling_menu,
        w!("Font Scaling"),
        ID_TRAY_FONT_SCALE_BASE,
        font_scale,
        DEFAULT_FONT_SCALE,
    );
    append_document_scale_menu(
        scaling_menu,
        w!("Design Scaling"),
        ID_TRAY_DESIGN_SCALE_BASE,
        design_scale,
        DEFAULT_DESIGN_SCALE,
    );

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        scaling_menu.0 as usize,
        w!("Scaling"),
    );

    // Add the "Background" submenu: what a preview is drawn over, which is a question
    // a picture and a document answer differently — a picture's transparency is the
    // picture's, while a document is drawn on a page — so each of them has a half
    // of its own, listing the same backdrops.
    let (
        image_background,
        vector_background,
        html_background,
        font_background,
        dds_background,
        design_background,
    ) = CONFIG
        .lock()
        .map(|c| {
            (
                c.image_background,
                c.vector_background,
                c.html_background,
                c.font_background,
                c.dds_background,
                c.design_background,
            )
        })
        .unwrap_or((
            DEFAULT_IMAGE_BACKGROUND,
            DEFAULT_VECTOR_BACKGROUND,
            DEFAULT_HTML_BACKGROUND,
            DEFAULT_FONT_BACKGROUND,
            DEFAULT_DDS_BACKGROUND,
            DEFAULT_DESIGN_BACKGROUND,
        ));
    let background_menu = CreatePopupMenu().unwrap();

    // Each half is handed its own default, since the backdrops a half offers are not the same
    // ones everywhere: a picture, a drawing and a design document start at the squares, a
    // specimen, a page and a texture start at a page — and a texture is offered only the two
    // pages, a page all but the transparency.
    append_background_menu(
        background_menu,
        w!("Image Background"),
        ID_TRAY_IMAGE_BACKGROUND_BASE,
        &BACKGROUND_CHOICES,
        image_background,
        DEFAULT_IMAGE_BACKGROUND,
    );
    append_background_menu(
        background_menu,
        w!("Vector Background"),
        ID_TRAY_VECTOR_BACKGROUND_BASE,
        &BACKGROUND_CHOICES,
        vector_background,
        DEFAULT_VECTOR_BACKGROUND,
    );
    append_background_menu(
        background_menu,
        w!("HTML Background"),
        ID_TRAY_HTML_BACKGROUND_BASE,
        &HTML_BACKGROUND_CHOICES,
        html_background,
        DEFAULT_HTML_BACKGROUND,
    );
    append_background_menu(
        background_menu,
        w!("Font Background"),
        ID_TRAY_FONT_BACKGROUND_BASE,
        &BACKGROUND_CHOICES,
        font_background,
        DEFAULT_FONT_BACKGROUND,
    );
    append_background_menu(
        background_menu,
        w!("DDS Background"),
        ID_TRAY_DDS_BACKGROUND_BASE,
        &DDS_BACKGROUND_CHOICES,
        dds_background,
        DEFAULT_DDS_BACKGROUND,
    );
    append_background_menu(
        background_menu,
        w!("Design Background"),
        ID_TRAY_DESIGN_BACKGROUND_BASE,
        &BACKGROUND_CHOICES,
        design_background,
        DEFAULT_DESIGN_BACKGROUND,
    );

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        background_menu.0 as usize,
        w!("Background"),
    );

    // Add the Volume submenu: two halves, because the two kinds of preview that make a sound
    // are hovered for different things — a video is looked at, and its soundtrack is as likely
    // to be a distraction as anything, while a sound file *is* the sound — so one setting for
    // both would mean turning a film's soundtrack up to hear a song. Each half lists the same
    // ten levels, in the same order, which is what lets one table and one builder serve them;
    // the level each setting stands at carries the default mark (see `VOLUME_CHOICES`). Above
    // each half's levels sits the switch the file itself answers — the loudness it is played
    // at —
    // and below the sound's own sits the other half of the same question, where in a file it
    // starts playing (see `AUDIO_SEEK_CHOICES`).
    let (video_volume, audio_volume, audio_seek) = CONFIG
        .lock()
        .map(|c| (c.video_volume, c.audio_volume, c.audio_seek))
        .unwrap_or((
            DEFAULT_VIDEO_VOLUME,
            DEFAULT_AUDIO_VOLUME,
            DEFAULT_AUDIO_SEEK,
        ));

    let volume_menu = CreatePopupMenu().unwrap();
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
                    MF_CHECKED
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
    let (normalize_audio, normalize_video) = CONFIG
        .lock()
        .map(|config| (config.normalize_volume, config.normalize_video_volume))
        .unwrap_or((DEFAULT_NORMALIZE_VOLUME, DEFAULT_NORMALIZE_VIDEO_VOLUME));
    // Whether a level turned on a pin is kept, asked of the configuration rather than kept beside
    // the switch that reads it: the two rows above the levels are one about the file and one about
    // the level, and neither is a level (see `remember_audio_volume`).
    let (remember_audio, remember_video) = CONFIG
        .lock()
        .map(|config| (config.remember_audio_volume, config.remember_video_volume))
        .unwrap_or((DEFAULT_REMEMBER_AUDIO_VOLUME, DEFAULT_REMEMBER_VIDEO_VOLUME));
    // Asked again here for the reason the `Codecs` rows are asked again: a machine that has just
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
                w!("Normalize"),
            );
            let _ = AppendMenuW(
                levels,
                MF_STRING | if remember { MF_CHECKED } else { MF_UNCHECKED },
                remember_id as usize,
                w!("Remember"),
            );
            let _ = AppendMenuW(levels, MF_SEPARATOR, 0, PCWSTR::null());
        };

    let video_levels = CreatePopupMenu().unwrap();
    append_switches(
        video_levels,
        normalize_video,
        ID_TRAY_NORMALIZE_VIDEO_VOLUME,
        remember_video,
        ID_TRAY_REMEMBER_VIDEO_VOLUME,
    );
    append_levels(
        video_levels,
        video_volume,
        ID_TRAY_VIDEO_VOLUME_BASE,
        DEFAULT_VIDEO_VOLUME,
    );

    let audio_levels = CreatePopupMenu().unwrap();
    append_switches(
        audio_levels,
        normalize_audio,
        ID_TRAY_NORMALIZE_VOLUME,
        remember_audio,
        ID_TRAY_REMEMBER_VOLUME,
    );
    append_levels(
        audio_levels,
        audio_volume,
        ID_TRAY_AUDIO_VOLUME_BASE,
        DEFAULT_AUDIO_VOLUME,
    );

    let _ = AppendMenuW(
        volume_menu,
        MF_STRING | MF_POPUP,
        video_levels.0 as usize,
        w!("Video"),
    );
    let _ = AppendMenuW(
        volume_menu,
        MF_STRING | MF_POPUP,
        audio_levels.0 as usize,
        w!("Audio"),
    );

    append_audio_seek_menu(volume_menu, audio_seek);

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        volume_menu.0 as usize,
        w!("Volume"),
    );

    // Add the "Performance" submenu: what the app costs while it is working — the memory it
    // holds on to between hovers. It is the block above `Engine`, where the engines
    // themselves are named: what is here is what is spent while they work.
    let performance_menu = CreatePopupMenu().unwrap();

    // Add the "Cache" submenu: what a preview's own data may cost between hovers — the frames
    // a decoded image was shown as, which are held in memory; the pages a document was drawn
    // as, which are files under the temp folder; the pictures the image converter developed
    // a file into, which are files of their own beside those; and the small subtitle files a
    // film's own embedded tracks were copied into, font attachments included, which is what
    // makes a hover of such a film a read of a few dozen kilobytes rather than a stream of the
    // whole film — each of them listed largest first, with the size its own cache starts at
    // marked, and each of them saying which of the two kinds of storage it is.
    let (image_cache_mb, document_cache_mb, image_disk_cache_mb, general_disk_cache_mb) = CONFIG
        .lock()
        .map(|c| {
            (
                c.image_cache_mb,
                c.document_cache_mb,
                c.image_disk_cache_mb,
                c.general_disk_cache_mb,
            )
        })
        .unwrap_or((
            DEFAULT_IMAGE_CACHE_MB,
            DEFAULT_DOCUMENT_CACHE_MB,
            DEFAULT_IMAGE_DISK_CACHE_MB,
            DEFAULT_GENERAL_DISK_CACHE_MB,
        ));

    let cache_menu = CreatePopupMenu().unwrap();

    // The labels are built once per cache and kept for as long as its sizes menu is
    // being filled out: `AppendMenuW` is handed a pointer, so the wide strings have
    // to outlive the call that lists them. Which size is the default is the one
    // thing they say that differs between the caches.
    for (name, base, held, default_mb) in [
        (
            w!("Image (RAM)"),
            ID_TRAY_IMAGE_CACHE_BASE,
            image_cache_mb,
            DEFAULT_IMAGE_CACHE_MB,
        ),
        (
            w!("Document (Disk)"),
            ID_TRAY_DOCUMENT_CACHE_BASE,
            document_cache_mb,
            DEFAULT_DOCUMENT_CACHE_MB,
        ),
        (
            w!("Image (Disk)"),
            ID_TRAY_IMAGE_DISK_CACHE_BASE,
            image_disk_cache_mb,
            DEFAULT_IMAGE_DISK_CACHE_MB,
        ),
        (
            w!("General (Disk)"),
            ID_TRAY_GENERAL_DISK_CACHE_BASE,
            general_disk_cache_mb,
            DEFAULT_GENERAL_DISK_CACHE_MB,
        ),
    ] {
        let sizes_menu = CreatePopupMenu().unwrap();

        let cache_labels: Vec<Vec<u16>> = CACHE_SIZE_CHOICES_MB
            .iter()
            .map(|megabytes| {
                cache_size_label(*megabytes, default_mb)
                    .encode_utf16()
                    .chain(std::iter::once(0))
                    .collect()
            })
            .collect();

        for (index, _) in CACHE_SIZE_CHOICES_MB.iter().enumerate() {
            let _ = AppendMenuW(
                sizes_menu,
                MF_STRING,
                (base + index as u16) as usize,
                PCWSTR(cache_labels[index].as_ptr()),
            );
        }

        // A size the menu does not offer — one a hand-edited `config.ini` asked for
        // — leaves every item unmarked rather than marking the nearest one.
        if let Some(index) = CACHE_SIZE_CHOICES_MB
            .iter()
            .position(|megabytes| *megabytes == held)
        {
            let _ = CheckMenuRadioItem(
                sizes_menu,
                base as u32,
                (base + CACHE_SIZE_CHOICES_MB.len() as u16 - 1) as u32,
                (base + index as u16) as u32,
                MF_BYCOMMAND.0,
            );
        }

        let _ = AppendMenuW(
            cache_menu,
            MF_STRING | MF_POPUP,
            sizes_menu.0 as usize,
            name,
        );
    }

    let _ = AppendMenuW(
        performance_menu,
        MF_STRING | MF_POPUP,
        cache_menu.0 as usize,
        w!("Cache"),
    );

    // Add the "Decode Budget" submenu: the ceiling on what one hover may decode or read
    // for rather than on what is kept — a picture's decode, a document's bytes, the page
    // Office exported. It is the one setting that bounds a file rather than a cache, and
    // the only one whose smaller sizes are for a machine with less memory to spare.
    let decode_budget_gb = CONFIG
        .lock()
        .map(|c| sanitize_decode_budget_gb(c.decode_budget_gb))
        .unwrap_or(DEFAULT_DECODE_BUDGET_GB);

    let budget_menu = CreatePopupMenu().unwrap();

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

    // A ceiling the menu does not offer — one a hand-edited `config.ini` asked for —
    // leaves every item unmarked rather than marking the nearest one.
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

    let _ = AppendMenuW(
        performance_menu,
        MF_STRING | MF_POPUP,
        budget_menu.0 as usize,
        w!("Decode Budget"),
    );

    // Add the "Tick" submenu: how often the loop looks at the pointer's world while
    // Explorer has focus. It is the app's own rate rather than a hover's — the one
    // number that trades how soon a move is answered against what the app costs while it
    // works — which is why it is here and not with the delays a hover waits out.
    let tick_ms = CONFIG.lock().map(|c| c.tick_ms).unwrap_or(DEFAULT_TICK_MS);

    append_tick_menu(performance_menu, w!("Tick"), ID_TRAY_TICK_BASE, tick_ms);

    // Add the "Hardware Acceleration" submenu: whether a video is decoded on the graphics card.
    // It is a submenu of its own for the same reason the `Tick` row is one rather than a row of
    // this block: the question is which parts of the app decode anything, and there is more than
    // one part to name — today a video and nothing else, which is why the row inside it is the
    // only one there is (see `video_hw_accel_device`).
    let video_hw_accel = CONFIG
        .lock()
        .map(|c| c.video_hw_accel)
        .unwrap_or(DEFAULT_VIDEO_HW_ACCEL);

    let hardware_menu = unsafe { CreatePopupMenu() }.unwrap();
    let _ = unsafe {
        AppendMenuW(
            hardware_menu,
            MF_STRING
                | if video_hw_accel {
                    MF_CHECKED
                } else {
                    MF_UNCHECKED
                },
            ID_TRAY_VIDEO_HW_ACCEL as usize,
            w!("Video"),
        )
    };
    let _ = unsafe {
        AppendMenuW(
            performance_menu,
            MF_STRING | MF_POPUP,
            hardware_menu.0 as usize,
            w!("Hardware Acceleration"),
        )
    };

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        performance_menu.0 as usize,
        w!("Performance"),
    );

    // Add the "Engine" submenu: which engine a preview is asked of, and how long each engine
    // this app starts is kept — the choice of engine where a kind of document has two of them,
    // and a TTL apiece for the applications this app leaves running between hovers. It is the
    // block below `Performance`, which is about what those engines cost while they are up.
    let engine_menu = CreatePopupMenu().unwrap();

    // How long Explorer has to be out of reach before an engine that is not marked
    // `Persistent` is let go. It is the first row here because it is what the three TTL
    // submenus below are answered against: the times those offer are what a persistent
    // engine is kept by, and this is what bounds one that is not (see `app::afk`).
    append_afk_timer_menu(engine_menu);

    // Add the "Select Engine" submenu: which engine each kind of document is asked of, where
    // there is a choice to make. It is the first of the three that name an engine, above the
    // three that say how long each engine this app starts is kept — the applications it names
    // are the ones those are about. Office is the only kind with two engines to choose
    // between.
    append_select_engine_menu(engine_menu);

    // Microsoft Office TTL: how long the Office engine a family started is kept after
    // that family's last page. Nothing is asked of an engine while it is being kept
    // — it is a process that has already been paid for, and the document it drew a
    // page of is closed — so what the setting buys is the next document of that
    // family not paying for an Office start, and what it costs is an Office
    // application in the process list. It is listed longest first, with the engine
    // that is never let go at the top.
    //
    // The times are what an engine marked `Persistent` is kept by. One that is not is let
    // go by the `AFK Timer` above instead, and its time is not consulted — which is what a
    // user who wants the old behaviour back switches on.
    let office_idle = CONFIG
        .lock()
        .map(|c| c.office_engine_idle)
        .unwrap_or(EngineIdle::Seconds(DEFAULT_OFFICE_ENGINE_IDLE_SECS));
    let office_persistent = CONFIG
        .lock()
        .map(|c| c.office_engine_persistent)
        .unwrap_or(false);

    append_engine_idle_menu(
        engine_menu,
        w!("Microsoft Office TTL"),
        EngineIdleIds {
            times: ID_TRAY_ENGINE_IDLE_BASE,
            persistent: ID_TRAY_ENGINE_PERSISTENT_BASE,
        },
        office_idle,
        office_persistent,
        DEFAULT_OFFICE_ENGINE_IDLE_SECS,
        true,
    );

    // LibreOffice TTL: the same question about the engine the documents beside Office are
    // drawn by, and it is the same shape: what is kept is the application, and what the
    // setting buys is the next document converted without paying for an engine start. The
    // engine holds a stub document of this app's own while it is kept, which is what makes
    // it a running instance a conversion can be handed to (see `libreoffice_render`).
    // Greyed out where no LibreOffice is installed, since there is nothing there to keep.
    let libreoffice_idle = CONFIG
        .lock()
        .map(|c| c.libreoffice_idle)
        .unwrap_or(EngineIdle::Seconds(DEFAULT_LIBREOFFICE_IDLE_SECS));
    let libreoffice_persistent = CONFIG
        .lock()
        .map(|c| c.libreoffice_persistent)
        .unwrap_or(false);

    append_engine_idle_menu(
        engine_menu,
        w!("LibreOffice TTL"),
        EngineIdleIds {
            times: ID_TRAY_LIBREOFFICE_IDLE_BASE,
            persistent: ID_TRAY_ENGINE_PERSISTENT_BASE + 1,
        },
        libreoffice_idle,
        libreoffice_persistent,
        DEFAULT_LIBREOFFICE_IDLE_SECS,
        libreoffice_render::available(),
    );

    // ImageMagick TTL: there is none, and that is the engine's own answer rather than an
    // omission. Every other engine this app starts is one it can keep — an application, a
    // browser, an automation server — and what their TTLs bound is a process left running.
    // ImageMagick is a converter: `magick.exe` reads a file, writes one and exits, so there is
    // no instance to hold and nothing for an idle time to keep. What it develops is a picture
    // of this app's, held in the image cache under the budget pictures already have, and a
    // file whose frame has been given up is developed again — a wait, not a setting
    // (see `imagemagick_render`).
    //
    // PeaZip TTL: and there is none for the same reason, which is the engine's own answer once
    // more. PeaZip is a frontend, and what this app runs of it is the console archiver it
    // carries, which is a converter like the one above: handed an archive it prints the table
    // of contents and exits. There is no instance to keep and nothing an idle time would bound
    // — and what a second hover of the same archive costs is no engine at all, since the
    // listing it produced is held under the file's own key (see `peazip_render`).
    //
    // Calibre TTL: and none again, for the engine's own answer a third time. `ebook-convert.exe`
    // is a converter as well — handed a book it writes a PDF of it and exits, booting a whole
    // Python application to do it — so there is no instance to hold and nothing an idle time
    // could keep warm. It is the engine a TTL would suit best on paper, since a launch costs
    // seconds, and it is the one engine that has nothing to hold: what a second hover of the
    // same book costs is a read of the page it converted, kept in the page cache under the
    // budget documents are kept under (see `calibre_render`).

    // WebView2 TTL: the same question about the browser that draws a document — every
    // document, still or not. It is greyed out on a machine with no WebView2 runtime, since
    // there is nothing there to keep.
    let webview_idle = CONFIG
        .lock()
        .map(|c| c.webview_idle)
        .unwrap_or(EngineIdle::Seconds(DEFAULT_WEBVIEW_IDLE_SECS));
    let webview_persistent = CONFIG.lock().map(|c| c.webview_persistent).unwrap_or(false);

    append_engine_idle_menu(
        engine_menu,
        w!("WebView2 TTL"),
        EngineIdleIds {
            times: ID_TRAY_WEBVIEW_IDLE_BASE,
            persistent: ID_TRAY_ENGINE_PERSISTENT_BASE + 2,
        },
        webview_idle,
        webview_persistent,
        DEFAULT_WEBVIEW_IDLE_SECS,
        webview_preview::is_available(),
    );

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        engine_menu.0 as usize,
        w!("Engine"),
    );

    // Add the "Codecs" submenu: every engine and every codec extension a preview can lean
    // on, each marked with whether this machine has it. The list is what is installed
    // rather than what this app can do, so a row that is greyed is a preview that will not
    // be shown and a package that can be installed to show it — and where the README names
    // a page for that package, picking the row offers to open it.
    //
    // The answers are asked again here, because this is the one moment a user is looking
    // at them: a codec extension installed a minute ago shows up the next time the menu is
    // opened, rather than at the next restart.
    refresh_codecs();
    append_codecs_menu(menu);

    let _ = AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null());

    // A newer release is offered where the app's own settings are, and only while one is
    // waiting to be installed: this menu is built from the state of the world every time
    // it is opened, so the row is there on the first opening after a check found one and
    // gone while there is nothing to say.
    if let Some(version) = updates::available() {
        let label = format!("Update is available! (v{version})");
        let label_wide: Vec<u16> = label.encode_utf16().chain(std::iter::once(0)).collect();
        let _ = AppendMenuW(
            menu,
            MF_STRING,
            ID_TRAY_UPDATE as usize,
            PCWSTR(label_wide.as_ptr()),
        );
    }

    // Add "Run at Startup" with checkmark, which is the registry's answer rather than the
    // configuration's: what starts this app is the entry, and the two can be made to differ
    // from outside this app.
    let startup_enabled = startup::is_startup_enabled();
    let flags = MF_STRING
        | if startup_enabled {
            MF_CHECKED
        } else {
            MF_UNCHECKED
        };
    let _ = AppendMenuW(menu, flags, ID_TRAY_STARTUP as usize, w!("Run at Startup"));

    // Add "Config.ini", the label carrying the version that is running. It is a submenu now,
    // with the two resets under it and the row that opens the file above them, and it is put
    // in with `InsertMenuItemW` rather than `AppendMenuW` for the sake of that row: an item
    // of a menu can carry a command and a submenu at once, and the version label is the row a
    // user looks for to open `config.ini` by hand, so it keeps the command it always had.
    // `AppendMenuW` gives a submenu's item the handle for an id instead, which is a command
    // no one can act on. `Open Config.ini` inside is the same command, so the file is
    // reachable whichever way a click on an item that opens a submenu is answered.
    let config_menu = CreatePopupMenu().unwrap();
    let _ = AppendMenuW(
        config_menu,
        MF_STRING,
        ID_TRAY_OPEN_CONFIG as usize,
        w!("Open Config.ini"),
    );
    let _ = AppendMenuW(config_menu, MF_SEPARATOR, 0, PCWSTR::null());

    // What each of the two would change, which is also the answer to whether it is offered:
    // a reset with nothing behind it is greyed rather than shown as a click that would do
    // nothing, and the question it asks is the difference counted here.
    let (settings_apart, lists_apart) = CONFIG
        .lock()
        .map(|config| {
            (
                config.settings_apart_from_recommended(),
                config.lists_apart_from_built_in(),
            )
        })
        .unwrap_or_default();

    let settings_flags = if settings_apart.is_empty() {
        MF_STRING | MF_GRAYED
    } else {
        MF_STRING
    };
    let _ = AppendMenuW(
        config_menu,
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
        config_menu,
        lists_flags,
        ID_TRAY_RESET_LISTS as usize,
        w!("Reset Extension Lists..."),
    );

    let config_label = format!("Config.ini (v{})", env!("CARGO_PKG_VERSION"));
    let config_label_wide: Vec<u16> = config_label
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();
    let config_item = MENUITEMINFOW {
        cbSize: std::mem::size_of::<MENUITEMINFOW>() as u32,
        fMask: MIIM_ID | MIIM_SUBMENU | MIIM_STRING,
        fType: MFT_STRING,
        wID: ID_TRAY_OPEN_CONFIG as u32,
        hSubMenu: config_menu,
        dwTypeData: PWSTR(config_label_wide.as_ptr() as *mut u16),
        cch: 0,
        ..Default::default()
    };
    let _ = InsertMenuItemW(
        menu,
        GetMenuItemCount(menu).max(0) as u32,
        BOOL(1),
        &config_item,
    );

    // Add Exit
    let _ = AppendMenuW(menu, MF_STRING, ID_TRAY_EXIT as usize, w!("Exit"));

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
    let _ = DestroyMenu(menu);
}

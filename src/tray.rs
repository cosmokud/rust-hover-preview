use crate::config::{
    sanitize_image_cache_mb, sanitize_office_cache_mb, sanitize_pdf_cache_mb,
    sanitize_text_cache_mb, sanitize_text_font_scale_percent, MarkdownMode, OfficeEngineIdle,
    PreviewScale, PreviewType, TextTheme, TransparentBackground, TriggerKeyMode,
    DEFAULT_IMAGE_CACHE_MB, DEFAULT_OFFICE_CACHE_MB, DEFAULT_OFFICE_ENGINE_IDLE_SECS,
    DEFAULT_PDF_CACHE_MB, DEFAULT_PREVIEW_SCALE_PERCENT, DEFAULT_TEXT_CACHE_MB,
    DEFAULT_TEXT_FONT_SCALE_PERCENT,
};
use crate::explorer_hook;
use crate::office_render;
use crate::pdf_preview;
use crate::preview_window::{refresh_preview, refresh_preview_types, trim_image_cache};
use crate::text_preview;
use crate::text_theme;
use crate::theme_files;
use crate::{startup, CONFIG, RUNNING};
use once_cell::sync::Lazy;
use std::os::windows::ffi::OsStrExt;
use std::sync::atomic::Ordering;
use std::sync::Mutex;
use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Shell::{
    ShellExecuteW, Shell_NotifyIconW, NIF_ICON, NIF_MESSAGE, NIF_TIP, NIM_ADD, NIM_DELETE,
    NOTIFYICONDATAW,
};
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CheckMenuRadioItem, CreatePopupMenu, CreateWindowExW, DefWindowProcW, DestroyMenu,
    DispatchMessageW, GetCursorPos, LoadImageW, PeekMessageW, PostQuitMessage, RegisterClassExW,
    RegisterWindowMessageW, SetForegroundWindow, TrackPopupMenu, TranslateMessage, CS_HREDRAW,
    CS_VREDRAW, HICON, IMAGE_ICON, LR_DEFAULTSIZE, LR_SHARED, MF_BYCOMMAND, MF_CHECKED, MF_POPUP,
    MF_SEPARATOR, MF_STRING, MF_UNCHECKED, MSG, PBT_APMRESUMEAUTOMATIC, PBT_APMRESUMESUSPEND,
    PM_REMOVE, SW_SHOWNORMAL, TPM_BOTTOMALIGN, TPM_LEFTALIGN, WM_COMMAND, WM_DESTROY,
    WM_LBUTTONUP, WM_POWERBROADCAST, WM_RBUTTONUP, WM_USER, WNDCLASSEXW, WS_EX_TOOLWINDOW,
    WS_POPUP,
};

const WM_TRAYICON: u32 = WM_USER + 1;
const ID_TRAY_EXIT: u16 = 1001;
const ID_TRAY_STARTUP: u16 = 1002;
const ID_TRAY_ENABLE: u16 = 1003;
const ID_TRAY_CONFIRM_FILE_TYPE: u16 = 1004;
const ID_TRAY_TRIGGER_DISABLE: u16 = 1005; // Hold the trigger key to stop previews
const ID_TRAY_TRIGGER_ENABLE: u16 = 1006; // Hold the trigger key to allow previews
const ID_TRAY_BG_TRANSPARENT: u16 = 1007;
const ID_TRAY_BG_BLACK: u16 = 1008;
const ID_TRAY_BG_WHITE: u16 = 1009;
const ID_TRAY_BG_CHECKERBOARD: u16 = 1016;
const ID_TRAY_VOLUME_MAX: u16 = 1010; // 100%
const ID_TRAY_VOLUME_HIGH: u16 = 1011; // 80%
const ID_TRAY_VOLUME_MEDIUM: u16 = 1012; // 50%
const ID_TRAY_VOLUME_LOW: u16 = 1013; // 25%
const ID_TRAY_VOLUME_VERY_LOW: u16 = 1014; // 10%
const ID_TRAY_VOLUME_MUTE: u16 = 1015; // 0%
const ID_TRAY_POSITION_FOLLOW: u16 = 1020; // Follow cursor
const ID_TRAY_POSITION_BEST: u16 = 1021; // Best position
const ID_TRAY_POSITION_AVOID_NAME: u16 = 1022; // Keep previews off the hovered item's name
const ID_TRAY_DELAY_INSTANT: u16 = 1030; // 0ms
const ID_TRAY_DELAY_VERY_FAST: u16 = 1031; // 200ms
const ID_TRAY_DELAY_MEDIUM: u16 = 1032; // 500ms
const ID_TRAY_DELAY_SLOW: u16 = 1033; // 1000ms
const ID_TRAY_REHOVER_DELAY_INSTANT: u16 = 1034; // 0ms
const ID_TRAY_REHOVER_DELAY_FAST: u16 = 1035; // 200ms
const ID_TRAY_REHOVER_DELAY_MEDIUM: u16 = 1036; // 500ms
const ID_TRAY_REHOVER_DELAY_SLOW: u16 = 1037; // 1000ms
const ID_TRAY_DELAY_FAST_PLUS: u16 = 1038; // 750ms
const ID_TRAY_REHOVER_DELAY_FAST_PLUS: u16 = 1039; // 750ms
const ID_TRAY_OPEN_CONFIG: u16 = 1040;
const ID_TRAY_SCALE_FIT: u16 = 1041;
const ID_TRAY_SCALE_400: u16 = 1042; // 400%
const ID_TRAY_SCALE_300: u16 = 1043; // 300%
const ID_TRAY_SCALE_200: u16 = 1044; // 200%
const ID_TRAY_SCALE_150: u16 = 1045; // 150%
const ID_TRAY_SCALE_100: u16 = 1046; // 100%
const ID_TRAY_SCALE_50: u16 = 1047; // 50%
const ID_TRAY_SCALE_25: u16 = 1048; // 25%
const ID_TRAY_THEME_LIGHT: u16 = 1050; // Atom One Light
const ID_TRAY_THEME_DARK: u16 = 1051; // One Dark Pro
const ID_TRAY_MARKDOWN_RENDERED: u16 = 1052; // Rendered document
const ID_TRAY_MARKDOWN_SOURCE: u16 = 1053; // Highlighted Markdown source
const ID_TRAY_TEXT_FULL_MODE: u16 = 1061; // Text previews scroll/select on/off
/// The `Preview Types` submenu, one command per kind of preview.
const ID_TRAY_TYPE_IMAGES: u16 = 1062;
const ID_TRAY_TYPE_VIDEOS: u16 = 1063;
const ID_TRAY_TYPE_TEXT: u16 = 1064;
const ID_TRAY_TYPE_PDF: u16 = 1065;
const ID_TRAY_TYPE_ARCHIVES: u16 = 1066;
const ID_TRAY_TYPE_OFFICE: u16 = 1067;
/// The `Cache` submenu: one command per size it offers, in the order it lists
/// them, for each of the caches it sizes. They start past the range the `theme`
/// folder's own items occupy (see `ID_TRAY_THEME_CUSTOM_BASE`).
const ID_TRAY_IMAGE_CACHE_BASE: u16 = 1300;
const ID_TRAY_OFFICE_CACHE_BASE: u16 = 1320;
const ID_TRAY_PDF_CACHE_BASE: u16 = 1340;
const ID_TRAY_TEXT_CACHE_BASE: u16 = 1360;
/// The sizes the `Cache` submenu offers, in megabytes, largest first — `2 GB` at
/// the top and a cache that holds nothing at the bottom — and the whole range the
/// settings allow, so a size a hand-edited `config.ini` asks for that is not one of
/// these is shown with nothing checked rather than rounded to one of them.
const CACHE_SIZE_CHOICES_MB: [u32; 9] = [2048, 1024, 512, 256, 128, 64, 32, 16, 0];
const ID_TRAY_FONT_100: u16 = 1072;
const ID_TRAY_FONT_125: u16 = 1073;
const ID_TRAY_FONT_150: u16 = 1074;
const ID_TRAY_FONT_175: u16 = 1075;
const ID_TRAY_FONT_200: u16 = 1076;
const ID_TRAY_FONT_250: u16 = 1077;
const ID_TRAY_FONT_300: u16 = 1078;
const ID_TRAY_FONT_400: u16 = 1079;
const ID_TRAY_FONT_90: u16 = 1080;
const ID_TRAY_FONT_80: u16 = 1081;
const ID_TRAY_FONT_70: u16 = 1082;
/// The `Performance → Keep Office Engine` submenu: one command per idle time it
/// offers, in the order it lists them. The IDs the app used before this ended at
/// 1082 and the `theme` folder's items start at 1100, so this range is the slack
/// between the two.
const ID_TRAY_ENGINE_IDLE_BASE: u16 = 1083;
/// The idle times the `Keep Office Engine` submenu offers, longest first — the
/// order the menu lists them in, so an engine that is never let go is the topmost
/// item and one that is let go as soon as it has drawn a page is the bottom one. A
/// value a hand-edited `config.ini` asks for that is not one of these is shown with
/// nothing marked rather than rounded to the nearest.
const ENGINE_IDLE_CHOICES: [OfficeEngineIdle; 7] = [
    OfficeEngineIdle::Indefinite,
    OfficeEngineIdle::Seconds(3600),
    OfficeEngineIdle::Seconds(1800),
    OfficeEngineIdle::Seconds(600),
    OfficeEngineIdle::Seconds(300),
    OfficeEngineIdle::Seconds(60),
    OfficeEngineIdle::Seconds(0),
];

/// Where the `theme` folder's own items start: one command ID each, in the order
/// the submenu listed them. The IDs the app uses end at the `Keep Office Engine`
/// range above, so these collide with nothing.
const ID_TRAY_THEME_CUSTOM_BASE: u16 = 1100;
/// How many files the theme submenu will list. A menu that long is unusable well
/// before this, and the cap is what keeps a folder of thousands of files from
/// running off the end of the command IDs.
const MAX_TRAY_CUSTOM_THEMES: usize = 200;

const TRAY_CLASS: PCWSTR = w!("RustHoverPreviewTrayClass");

static mut TRAY_HWND: HWND = HWND(std::ptr::null_mut());
static mut TASKBAR_CREATED: u32 = 0;

/// The custom themes the `Theme` submenu last listed, in the order it
/// listed them: a command ID carries a position, and this is what it was a
/// position in. The menu can outlive a change to the folder, so a click has to
/// select the file the item it landed on named rather than whatever is in that
/// place now.
static TRAY_CUSTOM_THEMES: Lazy<Mutex<Vec<TextTheme>>> = Lazy::new(|| Mutex::new(Vec::new()));

unsafe extern "system" fn tray_window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        _ if TASKBAR_CREATED != 0 && msg == TASKBAR_CREATED => {
            // Explorer (taskbar) restarted; re-add tray icon
            remove_tray_icon(hwnd);
            let _ = add_tray_icon(hwnd);
            // The Shell objects the Explorer hook resolves items through are served
            // by explorer.exe, so the ones it holds are proxies into the process
            // that is gone: it has to build them again, or nothing resolves until
            // the app itself is restarted.
            explorer_hook::note_explorer_restart();
            LRESULT(0)
        }
        WM_TRAYICON => {
            let event = lparam.0 as u32;
            if event == WM_RBUTTONUP || event == WM_LBUTTONUP {
                show_context_menu(hwnd);
            }
            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd = (wparam.0 & 0xFFFF) as u16;
            match cmd {
                ID_TRAY_EXIT => {
                    RUNNING.store(false, Ordering::SeqCst);
                    PostQuitMessage(0);
                }
                ID_TRAY_STARTUP => {
                    toggle_startup();
                }
                ID_TRAY_ENABLE => {
                    toggle_preview_enabled();
                }
                ID_TRAY_CONFIRM_FILE_TYPE => {
                    toggle_confirm_file_type();
                }
                ID_TRAY_TRIGGER_DISABLE => set_trigger_key_mode(TriggerKeyMode::Disable),
                ID_TRAY_TRIGGER_ENABLE => set_trigger_key_mode(TriggerKeyMode::Enable),
                ID_TRAY_BG_TRANSPARENT => {
                    set_transparent_background(TransparentBackground::Transparent)
                }
                ID_TRAY_BG_BLACK => set_transparent_background(TransparentBackground::Black),
                ID_TRAY_BG_WHITE => set_transparent_background(TransparentBackground::White),
                ID_TRAY_BG_CHECKERBOARD => {
                    set_transparent_background(TransparentBackground::Checkerboard)
                }
                ID_TRAY_VOLUME_MAX => set_volume(100),
                ID_TRAY_VOLUME_HIGH => set_volume(80),
                ID_TRAY_VOLUME_MEDIUM => set_volume(50),
                ID_TRAY_VOLUME_LOW => set_volume(25),
                ID_TRAY_VOLUME_VERY_LOW => set_volume(10),
                ID_TRAY_VOLUME_MUTE => set_volume(0),
                ID_TRAY_POSITION_FOLLOW => set_follow_cursor(true),
                ID_TRAY_POSITION_BEST => set_follow_cursor(false),
                ID_TRAY_POSITION_AVOID_NAME => toggle_avoid_filename(),
                ID_TRAY_DELAY_INSTANT => set_hover_delay(0),
                ID_TRAY_DELAY_VERY_FAST => set_hover_delay(200),
                ID_TRAY_DELAY_MEDIUM => set_hover_delay(500),
                ID_TRAY_DELAY_FAST_PLUS => set_hover_delay(750),
                ID_TRAY_DELAY_SLOW => set_hover_delay(1000),
                ID_TRAY_REHOVER_DELAY_INSTANT => set_same_file_rehover_delay(0),
                ID_TRAY_REHOVER_DELAY_FAST => set_same_file_rehover_delay(200),
                ID_TRAY_REHOVER_DELAY_MEDIUM => set_same_file_rehover_delay(500),
                ID_TRAY_REHOVER_DELAY_FAST_PLUS => set_same_file_rehover_delay(750),
                ID_TRAY_REHOVER_DELAY_SLOW => set_same_file_rehover_delay(1000),
                ID_TRAY_OPEN_CONFIG => open_config_file(),
                ID_TRAY_SCALE_FIT => set_preview_scale(PreviewScale::FitToScreen),
                ID_TRAY_SCALE_400 => set_preview_scale(PreviewScale::Percent(400)),
                ID_TRAY_SCALE_300 => set_preview_scale(PreviewScale::Percent(300)),
                ID_TRAY_SCALE_200 => set_preview_scale(PreviewScale::Percent(200)),
                ID_TRAY_SCALE_150 => set_preview_scale(PreviewScale::Percent(150)),
                ID_TRAY_SCALE_100 => set_preview_scale(PreviewScale::Percent(100)),
                ID_TRAY_SCALE_50 => set_preview_scale(PreviewScale::Percent(50)),
                ID_TRAY_SCALE_25 => set_preview_scale(PreviewScale::Percent(25)),
                ID_TRAY_THEME_LIGHT => set_theme(TextTheme::Light),
                ID_TRAY_THEME_DARK => set_theme(TextTheme::Dark),
                // The `theme` folder's items, by the position the submenu gave
                // them rather than any position in the folder.
                cmd if (ID_TRAY_THEME_CUSTOM_BASE..ID_TRAY_IMAGE_CACHE_BASE).contains(&cmd) => {
                    set_theme_from_menu((cmd - ID_TRAY_THEME_CUSTOM_BASE) as usize)
                }
                ID_TRAY_MARKDOWN_RENDERED => set_markdown_mode(MarkdownMode::Rendered),
                ID_TRAY_MARKDOWN_SOURCE => set_markdown_mode(MarkdownMode::Source),
                ID_TRAY_TEXT_FULL_MODE => toggle_text_preview_full_mode(),
                ID_TRAY_TYPE_IMAGES => toggle_preview_type(PreviewType::Images),
                ID_TRAY_TYPE_VIDEOS => toggle_preview_type(PreviewType::Videos),
                ID_TRAY_TYPE_TEXT => toggle_preview_type(PreviewType::Text),
                ID_TRAY_TYPE_PDF => toggle_preview_type(PreviewType::Pdf),
                ID_TRAY_TYPE_ARCHIVES => toggle_preview_type(PreviewType::Archives),
                ID_TRAY_TYPE_OFFICE => toggle_preview_type(PreviewType::Office),
                // An Office engine's idle time, by the position it was listed at.
                cmd if (ID_TRAY_ENGINE_IDLE_BASE
                    ..ID_TRAY_ENGINE_IDLE_BASE + ENGINE_IDLE_CHOICES.len() as u16)
                    .contains(&cmd) =>
                {
                    set_office_engine_idle(cmd - ID_TRAY_ENGINE_IDLE_BASE)
                }
                // A cache size, by the position it was listed at.
                cmd if (ID_TRAY_IMAGE_CACHE_BASE..ID_TRAY_OFFICE_CACHE_BASE).contains(&cmd) => {
                    set_image_cache_mb(cmd - ID_TRAY_IMAGE_CACHE_BASE)
                }
                cmd if (ID_TRAY_OFFICE_CACHE_BASE..ID_TRAY_PDF_CACHE_BASE).contains(&cmd) => {
                    set_office_cache_mb(cmd - ID_TRAY_OFFICE_CACHE_BASE)
                }
                cmd if (ID_TRAY_PDF_CACHE_BASE..ID_TRAY_TEXT_CACHE_BASE).contains(&cmd) => {
                    set_pdf_cache_mb(cmd - ID_TRAY_PDF_CACHE_BASE)
                }
                cmd if cmd >= ID_TRAY_TEXT_CACHE_BASE => {
                    set_text_cache_mb(cmd - ID_TRAY_TEXT_CACHE_BASE)
                }
                ID_TRAY_FONT_400 => set_text_font_scale(400),
                ID_TRAY_FONT_300 => set_text_font_scale(300),
                ID_TRAY_FONT_250 => set_text_font_scale(250),
                ID_TRAY_FONT_200 => set_text_font_scale(200),
                ID_TRAY_FONT_175 => set_text_font_scale(175),
                ID_TRAY_FONT_150 => set_text_font_scale(150),
                ID_TRAY_FONT_125 => set_text_font_scale(125),
                ID_TRAY_FONT_100 => set_text_font_scale(100),
                ID_TRAY_FONT_90 => set_text_font_scale(90),
                ID_TRAY_FONT_80 => set_text_font_scale(80),
                ID_TRAY_FONT_70 => set_text_font_scale(70),
                _ => {}
            }
            LRESULT(0)
        }
        WM_DESTROY => {
            remove_tray_icon(hwnd);
            PostQuitMessage(0);
            LRESULT(0)
        }
        WM_POWERBROADCAST => {
            let power_event = wparam.0 as u32;
            if power_event == PBT_APMRESUMEAUTOMATIC || power_event == PBT_APMRESUMESUSPEND {
                // System resumed from sleep — re-add tray icon in case
                // DWM/Explorer restart affected its visibility.
                remove_tray_icon(hwnd);
                let _ = add_tray_icon(hwnd);
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn show_context_menu(hwnd: HWND) {
    let menu = CreatePopupMenu().unwrap();
    if let Ok(mut config) = CONFIG.lock() {
        config.reload_from_disk();
    }

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

    // Add the "Preview Types" submenu: one gate per kind of preview, on by
    // default. A gate is only whether previews of that kind may be shown at all —
    // the lists and settings that decide which files of that kind preview are left
    // alone, so switching one off and back on restores what was configured.
    let kinds = [
        (PreviewType::Images, ID_TRAY_TYPE_IMAGES, w!("Images")),
        (PreviewType::Videos, ID_TRAY_TYPE_VIDEOS, w!("Videos")),
        (PreviewType::Text, ID_TRAY_TYPE_TEXT, w!("Text")),
        (PreviewType::Pdf, ID_TRAY_TYPE_PDF, w!("PDF")),
        (
            PreviewType::Archives,
            ID_TRAY_TYPE_ARCHIVES,
            w!("Archives"),
        ),
        (PreviewType::Office, ID_TRAY_TYPE_OFFICE, w!("Office")),
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

    // Add the "Background" submenu
    let transparent_background = CONFIG
        .lock()
        .map(|c| c.transparent_background)
        .unwrap_or(TransparentBackground::Transparent);
    let background_menu = CreatePopupMenu().unwrap();

    let bg_flag = |background: TransparentBackground| {
        MF_STRING
            | if transparent_background == background {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            }
    };
    let _ = AppendMenuW(
        background_menu,
        bg_flag(TransparentBackground::Transparent),
        ID_TRAY_BG_TRANSPARENT as usize,
        w!("Transparent"),
    );
    let _ = AppendMenuW(
        background_menu,
        bg_flag(TransparentBackground::Black),
        ID_TRAY_BG_BLACK as usize,
        w!("Black"),
    );
    let _ = AppendMenuW(
        background_menu,
        bg_flag(TransparentBackground::White),
        ID_TRAY_BG_WHITE as usize,
        w!("White"),
    );
    let _ = AppendMenuW(
        background_menu,
        bg_flag(TransparentBackground::Checkerboard),
        ID_TRAY_BG_CHECKERBOARD as usize,
        w!("Checkerboard"),
    );
    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        background_menu.0 as usize,
        w!("Background"),
    );

    let _ = AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null());

    // Add "Confirm File Type" with checkmark (content/header sniffing)
    let confirm_file_type = CONFIG.lock().map(|c| c.confirm_file_type).unwrap_or(false);
    let confirm_flags = MF_STRING
        | if confirm_file_type {
            MF_CHECKED
        } else {
            MF_UNCHECKED
        };
    let _ = AppendMenuW(
        menu,
        confirm_flags,
        ID_TRAY_CONFIRM_FILE_TYPE as usize,
        w!("Confirm File Type"),
    );

    // Add the "Trigger Key (Alt)" submenu: the key it watches, and what holding it
    // does. Which of the two is active is shown with radio marks, because only one
    // of them can be.
    let (trigger_key, trigger_key_mode) = CONFIG
        .lock()
        .map(|c| (c.trigger_key.clone(), c.trigger_key_mode))
        .unwrap_or(("alt".to_string(), TriggerKeyMode::Disable));
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
        menu,
        MF_STRING | MF_POPUP,
        trigger_menu.0 as usize,
        PCWSTR(trigger_label_wide.as_ptr()),
    );

    let _ = AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null());

    // Add the "Text Preview" submenu: whether full mode is on, and the theme, size
    // and Markdown mode a text preview is painted with.
    let text_menu = CreatePopupMenu().unwrap();

    // Add "Full Mode" with checkmark
    let text_full_mode = CONFIG
        .lock()
        .map(|c| c.text_preview_full_mode)
        .unwrap_or(true);
    let text_full_flags = MF_STRING
        | if text_full_mode {
            MF_CHECKED
        } else {
            MF_UNCHECKED
        };
    let _ = AppendMenuW(
        text_menu,
        text_full_flags,
        ID_TRAY_TEXT_FULL_MODE as usize,
        w!("Full Mode"),
    );

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
    let font_steps: [(u32, u16); 11] = [
        (400, ID_TRAY_FONT_400),
        (300, ID_TRAY_FONT_300),
        (250, ID_TRAY_FONT_250),
        (200, ID_TRAY_FONT_200),
        (175, ID_TRAY_FONT_175),
        (150, ID_TRAY_FONT_150),
        (125, ID_TRAY_FONT_125),
        (100, ID_TRAY_FONT_100),
        (90, ID_TRAY_FONT_90),
        (80, ID_TRAY_FONT_80),
        (70, ID_TRAY_FONT_70),
    ];

    // The labels are built once and kept: `AppendMenuW` is handed a pointer, so the
    // wide strings have to outlive the call that lists them.
    let font_labels: Vec<Vec<u16>> = font_steps
        .iter()
        .map(|(percent, _)| {
            format!("{percent}%")
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();

    for (index, (percent, id)) in font_steps.iter().enumerate() {
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

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        text_menu.0 as usize,
        w!("Text Preview"),
    );

    // Add the "Timing" submenu: how long a hover waits before its preview opens,
    // and how long the same file is held off after its preview was dismissed.
    let timing_menu = CreatePopupMenu().unwrap();

    // Add the Delay submenu
    let hover_delay_ms = CONFIG.lock().map(|c| c.hover_delay_ms).unwrap_or(0);
    let delay_menu = CreatePopupMenu().unwrap();

    let delay_flag = |delay: u64| {
        MF_STRING
            | if hover_delay_ms == delay {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            }
    };
    let _ = AppendMenuW(
        delay_menu,
        delay_flag(0),
        ID_TRAY_DELAY_INSTANT as usize,
        w!("Instant (0 ms)"),
    );
    let _ = AppendMenuW(
        delay_menu,
        delay_flag(200),
        ID_TRAY_DELAY_VERY_FAST as usize,
        w!("Fast (200 ms)"),
    );
    let _ = AppendMenuW(
        delay_menu,
        delay_flag(500),
        ID_TRAY_DELAY_MEDIUM as usize,
        w!("Medium (500 ms)"),
    );
    let _ = AppendMenuW(
        delay_menu,
        delay_flag(750),
        ID_TRAY_DELAY_FAST_PLUS as usize,
        w!("Relaxed (750 ms)"),
    );
    let _ = AppendMenuW(
        delay_menu,
        delay_flag(1000),
        ID_TRAY_DELAY_SLOW as usize,
        w!("Slow (1000 ms)"),
    );

    let _ = AppendMenuW(
        timing_menu,
        MF_STRING | MF_POPUP,
        delay_menu.0 as usize,
        w!("Delay"),
    );

    let same_file_rehover_delay_ms = CONFIG
        .lock()
        .map(|c| c.same_file_rehover_delay_ms)
        .unwrap_or(750);
    let rehover_delay_menu = CreatePopupMenu().unwrap();

    let rehover_delay_flag = |delay: u64| {
        MF_STRING
            | if same_file_rehover_delay_ms == delay {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            }
    };
    let _ = AppendMenuW(
        rehover_delay_menu,
        rehover_delay_flag(0),
        ID_TRAY_REHOVER_DELAY_INSTANT as usize,
        w!("Instant (0 ms)"),
    );
    let _ = AppendMenuW(
        rehover_delay_menu,
        rehover_delay_flag(200),
        ID_TRAY_REHOVER_DELAY_FAST as usize,
        w!("Fast (200 ms)"),
    );
    let _ = AppendMenuW(
        rehover_delay_menu,
        rehover_delay_flag(500),
        ID_TRAY_REHOVER_DELAY_MEDIUM as usize,
        w!("Medium (500 ms)"),
    );
    let _ = AppendMenuW(
        rehover_delay_menu,
        rehover_delay_flag(750),
        ID_TRAY_REHOVER_DELAY_FAST_PLUS as usize,
        w!("Relaxed (750 ms)"),
    );
    let _ = AppendMenuW(
        rehover_delay_menu,
        rehover_delay_flag(1000),
        ID_TRAY_REHOVER_DELAY_SLOW as usize,
        w!("Slow (1000 ms)"),
    );
    let _ = AppendMenuW(
        timing_menu,
        MF_STRING | MF_POPUP,
        rehover_delay_menu.0 as usize,
        w!("Rehover Delay"),
    );

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        timing_menu.0 as usize,
        w!("Timing"),
    );

    // Add the "Placement" submenu: where a preview lands relative to the cursor or
    // the focused item, and how large it is.
    let placement_menu = CreatePopupMenu().unwrap();

    // Add the Position submenu: which side a preview takes, and whether it is kept
    // off the name of the item it is about. The two placements are one setting shown
    // two ways, so they carry a radio mark each; avoiding the name is a setting of
    // its own and carries a checkmark.
    let (follow_cursor, avoid_filename) = CONFIG
        .lock()
        .map(|c| (c.follow_cursor, c.avoid_filename))
        .unwrap_or((false, true));
    let position_menu = CreatePopupMenu().unwrap();

    let _ = AppendMenuW(
        position_menu,
        MF_STRING,
        ID_TRAY_POSITION_FOLLOW as usize,
        w!("Follow Cursor"),
    );
    let _ = AppendMenuW(
        position_menu,
        MF_STRING,
        ID_TRAY_POSITION_BEST as usize,
        w!("Best Position"),
    );
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

    let _ = AppendMenuW(position_menu, MF_SEPARATOR, 0, PCWSTR::null());
    let _ = AppendMenuW(
        position_menu,
        MF_STRING
            | if avoid_filename {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            },
        ID_TRAY_POSITION_AVOID_NAME as usize,
        w!("Avoid Filename"),
    );

    let _ = AppendMenuW(
        placement_menu,
        MF_STRING | MF_POPUP,
        position_menu.0 as usize,
        w!("Position"),
    );

    // Add the Scaling submenu
    let preview_scale = CONFIG
        .lock()
        .map(|c| c.preview_scale)
        .unwrap_or(PreviewScale::Percent(DEFAULT_PREVIEW_SCALE_PERCENT));
    let scale_menu = CreatePopupMenu().unwrap();

    let scale_flag = |scale: PreviewScale| {
        MF_STRING
            | if preview_scale == scale {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            }
    };
    let _ = AppendMenuW(
        scale_menu,
        scale_flag(PreviewScale::FitToScreen),
        ID_TRAY_SCALE_FIT as usize,
        w!("Fit to Screen"),
    );
    let _ = AppendMenuW(
        scale_menu,
        scale_flag(PreviewScale::Percent(400)),
        ID_TRAY_SCALE_400 as usize,
        w!("400%"),
    );
    let _ = AppendMenuW(
        scale_menu,
        scale_flag(PreviewScale::Percent(300)),
        ID_TRAY_SCALE_300 as usize,
        w!("300%"),
    );
    let _ = AppendMenuW(
        scale_menu,
        scale_flag(PreviewScale::Percent(200)),
        ID_TRAY_SCALE_200 as usize,
        w!("200%"),
    );
    let _ = AppendMenuW(
        scale_menu,
        scale_flag(PreviewScale::Percent(150)),
        ID_TRAY_SCALE_150 as usize,
        w!("150%"),
    );
    let _ = AppendMenuW(
        scale_menu,
        scale_flag(PreviewScale::Percent(100)),
        ID_TRAY_SCALE_100 as usize,
        w!("100%"),
    );
    let _ = AppendMenuW(
        scale_menu,
        scale_flag(PreviewScale::Percent(50)),
        ID_TRAY_SCALE_50 as usize,
        w!("50%"),
    );
    let _ = AppendMenuW(
        scale_menu,
        scale_flag(PreviewScale::Percent(25)),
        ID_TRAY_SCALE_25 as usize,
        w!("25%"),
    );

    let _ = AppendMenuW(
        placement_menu,
        MF_STRING | MF_POPUP,
        scale_menu.0 as usize,
        w!("Scaling"),
    );

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        placement_menu.0 as usize,
        w!("Placement"),
    );

    // Add the Volume submenu
    let current_volume = CONFIG.lock().map(|c| c.video_volume).unwrap_or(0);
    let volume_menu = CreatePopupMenu().unwrap();

    let vol_flag = |vol: u32| {
        MF_STRING
            | if current_volume == vol {
                MF_CHECKED
            } else {
                MF_UNCHECKED
            }
    };
    let _ = AppendMenuW(
        volume_menu,
        vol_flag(100),
        ID_TRAY_VOLUME_MAX as usize,
        w!("Max (100%)"),
    );
    let _ = AppendMenuW(
        volume_menu,
        vol_flag(80),
        ID_TRAY_VOLUME_HIGH as usize,
        w!("High (80%)"),
    );
    let _ = AppendMenuW(
        volume_menu,
        vol_flag(50),
        ID_TRAY_VOLUME_MEDIUM as usize,
        w!("Medium (50%)"),
    );
    let _ = AppendMenuW(
        volume_menu,
        vol_flag(25),
        ID_TRAY_VOLUME_LOW as usize,
        w!("Low (25%)"),
    );
    let _ = AppendMenuW(
        volume_menu,
        vol_flag(10),
        ID_TRAY_VOLUME_VERY_LOW as usize,
        w!("Very Low (10%)"),
    );
    let _ = AppendMenuW(
        volume_menu,
        vol_flag(0),
        ID_TRAY_VOLUME_MUTE as usize,
        w!("Mute (0%)"),
    );

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        volume_menu.0 as usize,
        w!("Volume"),
    );

    // Add the "Performance" submenu: what the app costs while it is working — the
    // applications it starts and keeps, and the memory it holds on to — rather than
    // what a preview looks like.
    let performance_menu = CreatePopupMenu().unwrap();

    // Keep Office Engine: how long the Office engine a family started is kept after
    // that family's last page. Nothing is asked of an engine while it is being kept
    // — it is a process that has already been paid for, and the document it drew a
    // page of is closed — so what the setting buys is the next document of that
    // family not paying for an Office start, and what it costs is an Office
    // application in the process list. It is listed longest first, with the engine
    // that is never let go at the top.
    let engine_idle = CONFIG
        .lock()
        .map(|c| c.office_engine_idle)
        .unwrap_or(OfficeEngineIdle::Seconds(DEFAULT_OFFICE_ENGINE_IDLE_SECS));

    let engine_menu = CreatePopupMenu().unwrap();

    // The labels are kept for as long as the menu is being filled out, for the same
    // reason the cache labels are: `AppendMenuW` is handed a pointer, so the wide
    // strings have to outlive the call that lists them.
    let engine_labels: Vec<Vec<u16>> = ENGINE_IDLE_CHOICES
        .iter()
        .map(|idle| {
            engine_idle_label(*idle)
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect()
        })
        .collect();

    for (index, _) in ENGINE_IDLE_CHOICES.iter().enumerate() {
        let _ = AppendMenuW(
            engine_menu,
            MF_STRING,
            (ID_TRAY_ENGINE_IDLE_BASE + index as u16) as usize,
            PCWSTR(engine_labels[index].as_ptr()),
        );
    }

    // A time the menu does not offer — one a hand-edited `config.ini` asked for —
    // leaves every item unmarked rather than marking the nearest one.
    if let Some(index) = ENGINE_IDLE_CHOICES
        .iter()
        .position(|idle| *idle == engine_idle)
    {
        let _ = CheckMenuRadioItem(
            engine_menu,
            ID_TRAY_ENGINE_IDLE_BASE as u32,
            (ID_TRAY_ENGINE_IDLE_BASE + ENGINE_IDLE_CHOICES.len() as u16 - 1) as u32,
            (ID_TRAY_ENGINE_IDLE_BASE + index as u16) as u32,
            MF_BYCOMMAND.0,
        );
    }

    let _ = AppendMenuW(
        performance_menu,
        MF_STRING | MF_POPUP,
        engine_menu.0 as usize,
        w!("Keep Office Engine"),
    );

    // Add the "Cache" submenu: how much memory a preview's own data may be held in
    // between hovers — the frames a decoded image was shown as, the frames a text
    // preview was painted as, the pages a PDF was drawn as, and the pages Office
    // rendered — each of them listed largest first, with the size its own cache
    // starts at marked. All of it is held in memory and nowhere else, and nothing is
    // written to disk.
    let (image_cache_mb, office_cache_mb, pdf_cache_mb, text_cache_mb) = CONFIG
        .lock()
        .map(|c| {
            (
                c.image_cache_mb,
                c.office_cache_mb,
                c.pdf_cache_mb,
                c.text_cache_mb,
            )
        })
        .unwrap_or((
            DEFAULT_IMAGE_CACHE_MB,
            DEFAULT_OFFICE_CACHE_MB,
            DEFAULT_PDF_CACHE_MB,
            DEFAULT_TEXT_CACHE_MB,
        ));

    let cache_menu = CreatePopupMenu().unwrap();

    // The labels are built once per cache and kept for as long as its sizes menu is
    // being filled out: `AppendMenuW` is handed a pointer, so the wide strings have
    // to outlive the call that lists them. Which size is the default is the one
    // thing they say that differs between the caches.
    for (name, base, held, default_mb) in [
        (w!("Image"), ID_TRAY_IMAGE_CACHE_BASE, image_cache_mb, DEFAULT_IMAGE_CACHE_MB),
        (w!("Text"), ID_TRAY_TEXT_CACHE_BASE, text_cache_mb, DEFAULT_TEXT_CACHE_MB),
        (w!("PDF"), ID_TRAY_PDF_CACHE_BASE, pdf_cache_mb, DEFAULT_PDF_CACHE_MB),
        (w!("Office"), ID_TRAY_OFFICE_CACHE_BASE, office_cache_mb, DEFAULT_OFFICE_CACHE_MB),
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

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        performance_menu.0 as usize,
        w!("Performance"),
    );

    let _ = AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null());

    // Add "Run at Startup" with checkmark
    let startup_enabled = CONFIG.lock().map(|c| c.run_at_startup).unwrap_or(false);
    let flags = MF_STRING
        | if startup_enabled {
            MF_CHECKED
        } else {
            MF_UNCHECKED
        };
    let _ = AppendMenuW(menu, flags, ID_TRAY_STARTUP as usize, w!("Run at Startup"));

    // Add "Config.ini", the label carrying the version that is running
    let config_label = format!("Config.ini (v{})", env!("CARGO_PKG_VERSION"));
    let config_label_wide: Vec<u16> = config_label
        .encode_utf16()
        .chain(std::iter::once(0))
        .collect();
    let _ = AppendMenuW(
        menu,
        MF_STRING,
        ID_TRAY_OPEN_CONFIG as usize,
        PCWSTR(config_label_wide.as_ptr()),
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

fn toggle_startup() {
    if let Ok(mut config) = CONFIG.lock() {
        config.run_at_startup = !config.run_at_startup;
        config.save();

        if config.run_at_startup {
            startup::enable_startup();
        } else {
            startup::disable_startup();
        }
    }
}

fn toggle_preview_enabled() {
    if let Ok(mut config) = CONFIG.lock() {
        config.preview_enabled = !config.preview_enabled;
        config.save();
    }
}

/// What the trigger key does is a setting rather than a view of one, so the preview
/// on screen is rebuilt: in disable mode a held key is what keeps previews away, and
/// switching to enable mode while it is held should show one.
fn set_trigger_key_mode(mode: TriggerKeyMode) {
    if let Ok(mut config) = CONFIG.lock() {
        config.trigger_key_mode = mode;
        config.save();
    }
    refresh_preview();
}

fn toggle_confirm_file_type() {
    if let Ok(mut config) = CONFIG.lock() {
        config.confirm_file_type = !config.confirm_file_type;
        config.save();
    }
}

fn set_transparent_background(background: TransparentBackground) {
    if let Ok(mut config) = CONFIG.lock() {
        config.transparent_background = background;
        config.save();
    }
    refresh_preview();
}

/// A text preview's colors and glyphs are painted into its frame, so the preview
/// on screen is rebuilt rather than only composited again.
fn set_theme(theme: TextTheme) {
    if let Ok(mut config) = CONFIG.lock() {
        config.theme = theme;
        config.save();
    }
    refresh_preview();
}

/// The theme a custom item named, by the position the submenu listed it at. A
/// position there is no item for — a click that outlived its menu — selects
/// nothing rather than the wrong theme.
fn set_theme_from_menu(index: usize) {
    let theme = TRAY_CUSTOM_THEMES
        .lock()
        .ok()
        .and_then(|themes| themes.get(index).copied());

    if let Some(theme) = theme {
        set_theme(theme);
    }
}

/// Switching a kind of preview off drops the one on screen when it is of that
/// kind — the hover that produced it no longer measures — and switching one back
/// on leaves the preview that is up alone, since a preview of another kind has
/// nothing to do with the gate that changed. The Explorer hook reads the gates
/// fresh on every tick, so a change needs no restart and no cache to clear.
fn toggle_preview_type(kind: PreviewType) {
    if let Ok(mut config) = CONFIG.lock() {
        let enabled = kind.enabled_in(&config);
        kind.set_enabled_in(&mut config, !enabled);
        config.save();
    }
    refresh_preview_types();
}

/// What a cache size is called in the menu: the size, with the one the cache starts
/// at marked as the default — the caches do not all start at the same one, which is
/// why the default is passed in rather than written into the label. The two sizes at
/// the ceiling are the only ones that are not a plain number of megabytes.
fn cache_size_label(megabytes: u32, default_mb: u32) -> String {
    let label = match megabytes {
        1024 => "1 GB".to_string(),
        2048 => "2 GB".to_string(),
        other => format!("{other} MB"),
    };

    if megabytes == default_mb {
        format!("{label} (Default)")
    } else {
        label
    }
}

/// What an idle time is called in the `Keep Office Engine` submenu: the time, with
/// the one an engine is kept for by default marked. The shortest is the only one
/// that is not a whole minute, and the longest is the only one that is not a whole
/// number of minutes.
fn engine_idle_label(idle: OfficeEngineIdle) -> String {
    let label = match idle {
        OfficeEngineIdle::Indefinite => "Indefinitely".to_string(),
        OfficeEngineIdle::Seconds(0) => "0 seconds".to_string(),
        OfficeEngineIdle::Seconds(60) => "1 minute".to_string(),
        OfficeEngineIdle::Seconds(3600) => "1 hour".to_string(),
        OfficeEngineIdle::Seconds(seconds) => format!("{} minutes", seconds / 60),
    };

    if idle == OfficeEngineIdle::Seconds(DEFAULT_OFFICE_ENGINE_IDLE_SECS) {
        format!("{label} (Default)")
    } else {
        label
    }
}

/// The idle time an item of the `Keep Office Engine` submenu stands for, by the
/// position it was listed at. An id past the last time the menu offered is one that
/// is not there.
fn engine_idle_at(index: u16) -> Option<OfficeEngineIdle> {
    ENGINE_IDLE_CHOICES.get(index as usize).copied()
}

/// How long the Office engines are kept after their families' last pages.
///
/// Nothing is rebuilt here and nothing on screen changes. An engine that is being
/// let go sooner is let go by the worker the next time it looks — which is twice a
/// second while it is holding one — and one that is being kept longer is a setting
/// the next render already reads, so neither needs waking.
fn set_office_engine_idle(index: u16) {
    let Some(idle) = engine_idle_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.office_engine_idle = idle;
        config.save();
    }
}

/// The size an item of the `Cache` submenu stands for, by the position it was
/// listed at. An id past the last size the menu offered is one that is not there.
fn cache_size_at(index: u16) -> Option<u32> {
    CACHE_SIZE_CHOICES_MB.get(index as usize).copied()
}

/// How much memory the decoded-image cache may hold.
///
/// A preview on screen is not drawn from the cache but from the frame it was
/// loaded as, so nothing on screen changes — what changes is how much is freed, and
/// a smaller size frees it now rather than at the next decode.
fn set_image_cache_mb(index: u16) {
    let Some(megabytes) = cache_size_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.image_cache_mb = sanitize_image_cache_mb(megabytes);
        config.save();
    }

    trim_image_cache();
}

/// How much memory the pages Office rendered may be held in, between hovers.
///
/// Unlike the sizes beside it this does not switch anything off: a page is rendered
/// for the hover that asks for it whatever the size, and a size of nothing means it
/// is dropped when that hover ends. So nothing is rebuilt here either — the next
/// hover answers for itself.
fn set_office_cache_mb(index: u16) {
    let Some(megabytes) = cache_size_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.office_cache_mb = sanitize_office_cache_mb(megabytes);
        config.save();
    }

    office_render::trim_now();
}

/// How much memory the pages a PDF preview was drawn as may be held in, between
/// hovers.
///
/// Like the caches beside it this switches nothing off: a page is rendered for the
/// hover that asks for it whatever the size, and a size of nothing means it is
/// dropped the moment it has been drawn. So nothing on screen changes here —
/// what changes is how much is freed, and a smaller size frees it now rather than
/// at the next render.
fn set_pdf_cache_mb(index: u16) {
    let Some(megabytes) = cache_size_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.pdf_cache_mb = sanitize_pdf_cache_mb(megabytes);
        config.save();
    }

    pdf_preview::trim_now();
}

/// How much memory the frames a text preview was painted as may be held in,
/// between hovers.
///
/// Like the caches beside it this switches nothing off: a frame is painted for the
/// hover that asks for it whatever the size, and a size of nothing means it is
/// dropped the moment it has been painted. So nothing on screen changes here —
/// what changes is how much is freed, and a smaller size frees it now rather than
/// at the next paint.
fn set_text_cache_mb(index: u16) {
    let Some(megabytes) = cache_size_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.text_cache_mb = sanitize_text_cache_mb(megabytes);
        config.save();
    }

    text_preview::trim_now();
}

/// Full mode changes what a text preview *is* rather than what it shows — it
/// scrolls, it can be selected from, and the pointer can rest on it — so the
/// preview on screen is rebuilt the same way a new font size rebuilds it.
fn toggle_text_preview_full_mode() {
    if let Ok(mut config) = CONFIG.lock() {
        config.text_preview_full_mode = !config.text_preview_full_mode;
        config.save();
    }
    refresh_preview();
}

fn set_text_font_scale(percent: u32) {
    if let Ok(mut config) = CONFIG.lock() {
        config.text_font_scale_percent = sanitize_text_font_scale_percent(percent);
        config.save();
    }
    refresh_preview();
}

fn set_markdown_mode(mode: MarkdownMode) {
    if let Ok(mut config) = CONFIG.lock() {
        config.markdown_mode = mode;
        config.save();
    }
    refresh_preview();
}

fn set_volume(volume: u32) {
    if let Ok(mut config) = CONFIG.lock() {
        config.video_volume = volume;
        config.save();
    }
}

fn set_follow_cursor(follow: bool) {
    if let Ok(mut config) = CONFIG.lock() {
        config.follow_cursor = follow;
        config.save();
    }
}

/// Whether a preview keeps off the name of the file it is about is a question the
/// placement asks, and a placement is made when a preview is opened — so, like the
/// position setting beside it, this applies to the next hover rather than moving the
/// preview that is already up.
fn toggle_avoid_filename() {
    if let Ok(mut config) = CONFIG.lock() {
        config.avoid_filename = !config.avoid_filename;
        config.save();
    }
}

fn set_preview_scale(scale: PreviewScale) {
    if let Ok(mut config) = CONFIG.lock() {
        config.preview_scale = scale;
        config.save();
    }
}

fn set_hover_delay(hover_delay_ms: u64) {
    if let Ok(mut config) = CONFIG.lock() {
        config.hover_delay_ms = hover_delay_ms;
        config.save();
    }
}

fn set_same_file_rehover_delay(delay_ms: u64) {
    if let Ok(mut config) = CONFIG.lock() {
        config.same_file_rehover_delay_ms = delay_ms;
        config.save();
    }
}

fn open_config_file() {
    if let Ok(config) = CONFIG.lock() {
        config.save();
    }

    if let Some(path) = crate::config::AppConfig::config_path() {
        let wide_path: Vec<u16> = path
            .as_os_str()
            .encode_wide()
            .chain(std::iter::once(0))
            .collect();

        unsafe {
            let _ = ShellExecuteW(
                HWND(std::ptr::null_mut()),
                w!("open"),
                PCWSTR(wide_path.as_ptr()),
                PCWSTR(std::ptr::null()),
                PCWSTR(std::ptr::null()),
                SW_SHOWNORMAL,
            );
        }
    }
}

unsafe fn add_tray_icon(hwnd: HWND) -> bool {
    // Load the embedded icon resource (assets/icon.ico compiled via build.rs)
    let hicon = if let Ok(hmodule) = GetModuleHandleW(None) {
        let hinstance = HINSTANCE(hmodule.0);
        match LoadImageW(
            hinstance,
            PCWSTR(1 as *const u16),
            IMAGE_ICON,
            0,
            0,
            LR_DEFAULTSIZE | LR_SHARED,
        ) {
            Ok(h) => HICON(h.0),
            Err(_) => HICON::default(),
        }
    } else {
        HICON::default()
    };

    // Fallback to system icon if custom icon failed
    let hicon = if hicon.0.is_null() {
        match LoadImageW(
            None,
            PCWSTR(32512 as *const u16), // IDI_APPLICATION
            IMAGE_ICON,
            0,
            0,
            LR_DEFAULTSIZE | LR_SHARED,
        ) {
            Ok(h) => HICON(h.0),
            Err(_) => HICON::default(),
        }
    } else {
        hicon
    };

    let mut nid = NOTIFYICONDATAW {
        cbSize: std::mem::size_of::<NOTIFYICONDATAW>() as u32,
        hWnd: hwnd,
        uID: 1,
        uFlags: NIF_ICON | NIF_MESSAGE | NIF_TIP,
        uCallbackMessage: WM_TRAYICON,
        hIcon: hicon,
        ..Default::default()
    };

    // Set tooltip
    let tip = "Rust Hover Preview";
    let tip_wide: Vec<u16> = tip.encode_utf16().chain(std::iter::once(0)).collect();
    let len = tip_wide.len().min(nid.szTip.len());
    nid.szTip[..len].copy_from_slice(&tip_wide[..len]);

    Shell_NotifyIconW(NIM_ADD, &nid).as_bool()
}

unsafe fn remove_tray_icon(hwnd: HWND) {
    let nid = NOTIFYICONDATAW {
        cbSize: std::mem::size_of::<NOTIFYICONDATAW>() as u32,
        hWnd: hwnd,
        uID: 1,
        ..Default::default()
    };
    let _ = Shell_NotifyIconW(NIM_DELETE, &nid);
}

pub fn run_tray() {
    unsafe {
        let hinstance = GetModuleHandleW(None).unwrap();

        // Register window class
        let wc = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(tray_window_proc),
            hInstance: hinstance.into(),
            lpszClassName: TRAY_CLASS,
            ..Default::default()
        };

        RegisterClassExW(&wc);

        // Register TaskbarCreated message to detect Explorer restarts
        TASKBAR_CREATED = RegisterWindowMessageW(w!("TaskbarCreated"));

        // Create hidden window for tray messages
        let hwnd = CreateWindowExW(
            WS_EX_TOOLWINDOW,
            TRAY_CLASS,
            w!("Hover Preview Tray"),
            WS_POPUP,
            0,
            0,
            0,
            0,
            None,
            None,
            hinstance,
            None,
        );

        let hwnd = match hwnd {
            Ok(h) => h,
            Err(e) => {
                eprintln!("Failed to create tray window: {:?}", e);
                return;
            }
        };

        TRAY_HWND = hwnd;

        // Add tray icon (retry briefly in case Explorer isn't ready yet)
        let mut added = add_tray_icon(hwnd);
        if !added {
            let mut retries = 20;
            while !added && retries > 0 && RUNNING.load(Ordering::SeqCst) {
                std::thread::sleep(std::time::Duration::from_millis(500));
                added = add_tray_icon(hwnd);
                retries -= 1;
            }
        }

        if !added {
            eprintln!("Failed to add tray icon after retries; exiting.");
            RUNNING.store(false, Ordering::SeqCst);
            remove_tray_icon(hwnd);
            return;
        }

        // Message loop
        let mut msg = MSG::default();
        while RUNNING.load(Ordering::SeqCst) {
            while PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE).as_bool() {
                if msg.message == windows::Win32::UI::WindowsAndMessaging::WM_QUIT {
                    RUNNING.store(false, Ordering::SeqCst);
                    break;
                }
                let _ = TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
            std::thread::sleep(std::time::Duration::from_millis(10));
        }

        // Cleanup
        remove_tray_icon(hwnd);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The order the submenu is built in is the order it is read in, and it is
    /// built top down: the engine that is never let go is the topmost item and the
    /// one let go at once is the bottom one, which is what the choices are listed
    /// in.
    #[test]
    fn engine_idle_is_offered_longest_first() {
        assert_eq!(
            ENGINE_IDLE_CHOICES.map(engine_idle_label),
            [
                "Indefinitely".to_string(),
                "1 hour".to_string(),
                "30 minutes".to_string(),
                "10 minutes (Default)".to_string(),
                "5 minutes".to_string(),
                "1 minute".to_string(),
                "0 seconds".to_string(),
            ]
        );
    }

    /// Every item a person can pick is one the setting can hold, so what the menu
    /// writes is what the menu reads back and marks on the next open.
    #[test]
    fn every_offered_idle_time_is_one_the_setting_keeps() {
        for idle in ENGINE_IDLE_CHOICES {
            assert_eq!(
                OfficeEngineIdle::from_str(&idle.as_str()),
                Some(idle),
                "{idle:?}"
            );
        }

        assert_eq!(
            engine_idle_at(ENGINE_IDLE_CHOICES.len() as u16),
            None,
            "an id past the last item is not one the menu offered"
        );
        assert_eq!(engine_idle_at(0), Some(OfficeEngineIdle::Indefinite));
    }
}

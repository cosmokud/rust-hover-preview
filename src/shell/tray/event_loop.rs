//! The tray window: the message proc behind the icon, and the loop that runs it.
//!
//! One window and one function carry every message the tray gets, which is why the whole
//! match runs under `catch_unwind`: it is `extern "system"`, so a panic unwinding out of it
//! cannot unwind and takes the process with it. One bad menu build would otherwise be a dead
//! app rather than a tray that did not open. The catch is as wide as the proc on purpose,
//! because the menu build, the handlers and the arm that re-adds the icon after Explorer
//! restarts are one code path to a user - and it is safe only because nothing in this codebase
//! takes a poisoned lock fatally (see `tray_window_proc`).
//!
//! `run_tray` is the main thread's run: it registers the class, makes the window, adds the
//! icon, and pumps messages until the window is destroyed.

use super::commands::{
    open_config_file, reset_lists_from_tray, reset_settings_from_tray, set_afk_timer,
    set_animated_scale, set_audio_seek, set_audio_volume, set_avoid_mode, set_dds_background,
    set_decode_budget_gb, set_design_background, set_design_scale, set_document_cache_mb,
    set_document_scale, set_ebook_scale, set_follow_cursor, set_font_background, set_font_scale,
    set_general_disk_cache_mb, set_hover_delay, set_html_background, set_image_background,
    set_image_cache_mb, set_image_disk_cache_mb, set_libreoffice_idle, set_markdown_mode,
    set_office_engine, set_office_engine_idle, set_pin_nav_file_types, set_preview_scale,
    set_same_file_rehover_delay, set_settling_delay, set_text_font_scale, set_text_scale,
    set_theme, set_theme_from_menu, set_tick_ms, set_trigger_key_mode, set_vector_background,
    set_vector_scale, set_video_engine, set_video_scale, set_video_volume, set_webview_idle,
    toggle_engine_persistent, toggle_normalize_video_volume, toggle_normalize_volume,
    toggle_pin_enabled, toggle_pin_pause_audio, toggle_pin_pause_video, toggle_pin_update_enabled,
    toggle_pin_update_on_hover, toggle_preview_enabled, toggle_preview_type,
    toggle_prioritize_keyboard, toggle_remember_audio_volume, toggle_remember_video_volume,
    toggle_render_html, toggle_startup, toggle_trigger_key_affect_pin_mode,
    toggle_trigger_key_enabled, toggle_video_engine_fallback, toggle_video_hw_accel,
};
use super::ids::{
    AFK_TIMER_CHOICES_SECS, AUDIO_SEEK_CHOICES, AVOID_CHOICES, BACKGROUND_CHOICES,
    BITMAP_SCALE_CHOICES, CACHE_SIZE_CHOICES_MB, CODEC_COMMANDS, DDS_BACKGROUND_CHOICES,
    DECODE_BUDGET_CHOICES_GB, DOCUMENT_SCALE_CHOICES, ENGINE_IDLE_CHOICES, HTML_BACKGROUND_CHOICES,
    ID_TRAY_AFK_TIMER_BASE, ID_TRAY_ANIMATED_SCALE_BASE, ID_TRAY_AUDIO_SEEK_BASE,
    ID_TRAY_AUDIO_VOLUME_BASE, ID_TRAY_AVOID_BASE, ID_TRAY_CODEC_BASE, ID_TRAY_DDS_BACKGROUND_BASE,
    ID_TRAY_DECODE_BUDGET_BASE, ID_TRAY_DELAY_BASE, ID_TRAY_DESIGN_BACKGROUND_BASE,
    ID_TRAY_DESIGN_SCALE_BASE, ID_TRAY_DOCUMENT_CACHE_BASE, ID_TRAY_DOCUMENT_SCALE_BASE,
    ID_TRAY_EBOOK_SCALE_BASE, ID_TRAY_ENABLE, ID_TRAY_ENGINE_IDLE_BASE,
    ID_TRAY_ENGINE_OFFICE_LIBRE, ID_TRAY_ENGINE_OFFICE_MS, ID_TRAY_ENGINE_PERSISTENT_BASE,
    ID_TRAY_EXIT, ID_TRAY_FONT_100, ID_TRAY_FONT_110, ID_TRAY_FONT_125, ID_TRAY_FONT_150,
    ID_TRAY_FONT_175, ID_TRAY_FONT_200, ID_TRAY_FONT_250, ID_TRAY_FONT_300, ID_TRAY_FONT_400,
    ID_TRAY_FONT_70, ID_TRAY_FONT_80, ID_TRAY_FONT_90, ID_TRAY_FONT_BACKGROUND_BASE,
    ID_TRAY_FONT_SCALE_BASE, ID_TRAY_GENERAL_DISK_CACHE_BASE, ID_TRAY_HTML_BACKGROUND_BASE,
    ID_TRAY_IMAGE_BACKGROUND_BASE, ID_TRAY_IMAGE_CACHE_BASE, ID_TRAY_IMAGE_DISK_CACHE_BASE,
    ID_TRAY_LIBREOFFICE_IDLE_BASE, ID_TRAY_MARKDOWN_RENDERED, ID_TRAY_MARKDOWN_SOURCE,
    ID_TRAY_NORMALIZE_VIDEO_VOLUME, ID_TRAY_NORMALIZE_VOLUME, ID_TRAY_OPEN_CONFIG, ID_TRAY_PIN,
    ID_TRAY_PIN_NAV_ALL, ID_TRAY_PIN_NAV_CATEGORY, ID_TRAY_PIN_PAUSE_AUDIO,
    ID_TRAY_PIN_PAUSE_VIDEO, ID_TRAY_PIN_UPDATE, ID_TRAY_PIN_UPDATE_HOVER, ID_TRAY_POSITION_BEST,
    ID_TRAY_POSITION_FOLLOW, ID_TRAY_PRIORITIZE_KEYBOARD, ID_TRAY_REHOVER_DELAY_BASE,
    ID_TRAY_REMEMBER_VIDEO_VOLUME, ID_TRAY_REMEMBER_VOLUME, ID_TRAY_RENDER_HTML,
    ID_TRAY_RESET_LISTS, ID_TRAY_RESET_SETTINGS, ID_TRAY_SCALE_BASE, ID_TRAY_SETTLING_DELAY_BASE,
    ID_TRAY_STARTUP, ID_TRAY_TEXT_SCALE_BASE, ID_TRAY_THEME_CUSTOM_BASE, ID_TRAY_THEME_DARK,
    ID_TRAY_THEME_LIGHT, ID_TRAY_TICK_BASE, ID_TRAY_TRIGGER_AFFECT_PIN, ID_TRAY_TRIGGER_DISABLE,
    ID_TRAY_TRIGGER_ENABLE, ID_TRAY_TRIGGER_ENABLED, ID_TRAY_TYPE_ARCHIVES, ID_TRAY_TYPE_AUDIO,
    ID_TRAY_TYPE_DESIGN, ID_TRAY_TYPE_DOCUMENT, ID_TRAY_TYPE_EBOOK, ID_TRAY_TYPE_FONTS,
    ID_TRAY_TYPE_IMAGES, ID_TRAY_TYPE_TEXT, ID_TRAY_TYPE_VECTOR, ID_TRAY_TYPE_VIDEOS,
    ID_TRAY_UPDATE, ID_TRAY_VECTOR_BACKGROUND_BASE, ID_TRAY_VECTOR_SCALE_BASE,
    ID_TRAY_VIDEO_ENGINE_BASE, ID_TRAY_VIDEO_ENGINE_FALLBACK, ID_TRAY_VIDEO_HW_ACCEL,
    ID_TRAY_VIDEO_SCALE_BASE, ID_TRAY_VIDEO_VOLUME_BASE, ID_TRAY_WEBVIEW_IDLE_BASE,
    TASKBAR_CREATED, TICK_CHOICES_MS, TIMING_DELAY_CHOICES_MS, TRAY_CLASS, TRAY_HWND,
    VIDEO_ENGINE_CHOICES, WM_TRAYICON,
};
use super::menus::show_context_menu;
use super::submenus::open_codec_page;

use crate::app::{dialogs, updates};
use crate::config::config::{
    MarkdownMode, OfficeEngine, PinNavFileTypes, PreviewType, TextTheme, TriggerKeyMode,
    VOLUME_CHOICES,
};
use crate::shell::explorer_hook;
use crate::{StartupTrace, RUNNING};
use std::panic::catch_unwind;
use std::sync::atomic::Ordering;
use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Shell::{
    Shell_NotifyIconW, NIF_ICON, NIF_MESSAGE, NIF_TIP, NIM_ADD, NIM_DELETE, NOTIFYICONDATAW,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DispatchMessageW, GetMessageW, LoadImageW, PostQuitMessage,
    RegisterClassExW, RegisterWindowMessageW, TranslateMessage, CS_HREDRAW, CS_VREDRAW, HICON,
    IMAGE_ICON, LR_DEFAULTSIZE, LR_SHARED, MSG, PBT_APMRESUMEAUTOMATIC, PBT_APMRESUMESUSPEND,
    WM_COMMAND, WM_DESTROY, WM_LBUTTONUP, WM_POWERBROADCAST, WM_RBUTTONUP, WNDCLASSEXW,
    WS_EX_TOOLWINDOW, WS_POPUP,
};

/// The one function every message the tray window gets arrives in. It is `extern "system"`,
/// which means a panic unwinding out of it cannot unwind: the ABI boundary aborts the process
/// instead, so one bad menu build is a dead app rather than a tray that did not open. The whole
/// match is therefore run under `catch_unwind`, and a panic returns `LRESULT(0)` and leaves the
/// app running.
///
/// The catch is deliberately as wide as the whole proc, because the tray is one window and one
/// function: the menu build, the `WM_COMMAND` handlers that write a setting, and the arm that
/// re-adds the icon after Explorer restarts are all the same code path to a user, and one guard
/// there covers all three.
///
/// That width is safe because nothing in this codebase takes a poisoned lock fatally —
/// repo-wide there is not one `lock().unwrap()` call. `CONFIG`, `TRAY_CUSTOM_THEMES` and the
/// update caches are all read through `if let Ok` or `.map(..).unwrap_or(..)`, so a panic partway
/// through a build leaves the menu to open with defaults rather than failing the same way twice.
/// Keep that invariant: the day someone adds a `lock().unwrap()`, this catch stops being a
/// safety net and starts being a wedge generator.
///
/// The one ugly consequence, stated plainly: a panic in a `WM_COMMAND` handler unwinds out of
/// the message pump inside `TrackPopupMenu`, so `TrackPopupMenu` never returns to the code that
/// called it, `DestroyMenu` after it is skipped, and the menu is left on screen attached to a
/// leaked handle. It still works, and it goes away when the user dismisses it, but it does not
/// go away on its own. That is a better outcome than a dead process; it is not a clean one.
unsafe extern "system" fn tray_window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    catch_unwind(|| {
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
                // While one of the app's dialogs is on screen the tray answers nothing: a menu
                // opened over it would be a second way into the same question, and a dialog —
                // which owns no window of this app's — is not modal to anything.
                if (event == WM_RBUTTONUP || event == WM_LBUTTONUP) && !dialogs::is_confirming() {
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
                    ID_TRAY_UPDATE => {
                        // The update row answers three ways, and only one of them ends this app:
                        // `Auto` is the installer, which runs silently, replaces this app and
                        // starts it again, so the app ends itself here rather than waiting to be
                        // terminated by the installer it just started; `Manual` is the release
                        // page opened in the user's own browser; `Cancel` is a click that does
                        // nothing and a menu that stays as it was.
                        match updates::ask() {
                            updates::Answer::Auto => {
                                if updates::install() {
                                    RUNNING.store(false, Ordering::SeqCst);
                                    PostQuitMessage(0);
                                }
                            }
                            updates::Answer::Manual => updates::open_release_page(),
                            updates::Answer::Cancel => {}
                        }
                    }
                    ID_TRAY_ENABLE => {
                        toggle_preview_enabled();
                    }
                    ID_TRAY_PIN => {
                        toggle_pin_enabled();
                    }
                    ID_TRAY_PIN_UPDATE => {
                        toggle_pin_update_enabled();
                    }
                    ID_TRAY_PIN_UPDATE_HOVER => {
                        toggle_pin_update_on_hover();
                    }
                    ID_TRAY_PIN_PAUSE_AUDIO => {
                        toggle_pin_pause_audio();
                    }
                    ID_TRAY_PIN_PAUSE_VIDEO => {
                        toggle_pin_pause_video();
                    }
                    ID_TRAY_PIN_NAV_ALL => set_pin_nav_file_types(PinNavFileTypes::All),
                    ID_TRAY_PIN_NAV_CATEGORY => set_pin_nav_file_types(PinNavFileTypes::Category),
                    ID_TRAY_TRIGGER_DISABLE => set_trigger_key_mode(TriggerKeyMode::Disable),
                    ID_TRAY_TRIGGER_ENABLE => set_trigger_key_mode(TriggerKeyMode::Enable),
                    ID_TRAY_TRIGGER_ENABLED => toggle_trigger_key_enabled(),
                    ID_TRAY_TRIGGER_AFFECT_PIN => toggle_trigger_key_affect_pin_mode(),
                    ID_TRAY_PRIORITIZE_KEYBOARD => toggle_prioritize_keyboard(),
                    ID_TRAY_ENGINE_OFFICE_MS => set_office_engine(OfficeEngine::MicrosoftOffice),
                    ID_TRAY_ENGINE_OFFICE_LIBRE => set_office_engine(OfficeEngine::LibreOffice),
                    ID_TRAY_VIDEO_ENGINE_FALLBACK => toggle_video_engine_fallback(),
                    // Which engine plays a video, by the position it was listed at: the
                    // `Video` submenu is the same base-plus-position arrangement every value
                    // menu here is.
                    cmd if (ID_TRAY_VIDEO_ENGINE_BASE
                        ..ID_TRAY_VIDEO_ENGINE_BASE + VIDEO_ENGINE_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_video_engine(
                            VIDEO_ENGINE_CHOICES[(cmd - ID_TRAY_VIDEO_ENGINE_BASE) as usize],
                        )
                    }
                    // A backdrop, by the position it was listed at: an image's or a
                    // document's, whichever half of the `Background` submenu it was in.
                    cmd if (ID_TRAY_IMAGE_BACKGROUND_BASE
                        ..ID_TRAY_IMAGE_BACKGROUND_BASE + BACKGROUND_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_image_background(cmd - ID_TRAY_IMAGE_BACKGROUND_BASE)
                    }
                    cmd if (ID_TRAY_FONT_BACKGROUND_BASE
                        ..ID_TRAY_FONT_BACKGROUND_BASE + BACKGROUND_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_font_background(cmd - ID_TRAY_FONT_BACKGROUND_BASE)
                    }
                    cmd if (ID_TRAY_DDS_BACKGROUND_BASE
                        ..ID_TRAY_DDS_BACKGROUND_BASE + DDS_BACKGROUND_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_dds_background(cmd - ID_TRAY_DDS_BACKGROUND_BASE)
                    }
                    cmd if (ID_TRAY_DESIGN_BACKGROUND_BASE
                        ..ID_TRAY_DESIGN_BACKGROUND_BASE + BACKGROUND_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_design_background(cmd - ID_TRAY_DESIGN_BACKGROUND_BASE)
                    }
                    cmd if (ID_TRAY_VECTOR_BACKGROUND_BASE
                        ..ID_TRAY_VECTOR_BACKGROUND_BASE + BACKGROUND_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_vector_background(cmd - ID_TRAY_VECTOR_BACKGROUND_BASE)
                    }
                    cmd if (ID_TRAY_HTML_BACKGROUND_BASE
                        ..ID_TRAY_HTML_BACKGROUND_BASE + HTML_BACKGROUND_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_html_background(cmd - ID_TRAY_HTML_BACKGROUND_BASE)
                    }
                    // The row above the levels of each half of the `Volume` submenu, the sound's and
                    // the video's: switches rather than levels, and read where a player is started
                    // rather than here.
                    ID_TRAY_NORMALIZE_VOLUME => toggle_normalize_volume(),
                    ID_TRAY_NORMALIZE_VIDEO_VOLUME => toggle_normalize_video_volume(),
                    // And the row under each of those: whether a level turned on a pin is the level
                    // the next preview is played at, which is the same bargain — read by the next
                    // player rather than by anything on screen.
                    ID_TRAY_REMEMBER_VOLUME => toggle_remember_audio_volume(),
                    ID_TRAY_REMEMBER_VIDEO_VOLUME => toggle_remember_video_volume(),
                    // A level of either half of the `Volume` submenu, by the position it was
                    // listed at. The two halves offer the same levels, so one table answers for
                    // both and each range is what says which setting was meant.
                    cmd if (ID_TRAY_VIDEO_VOLUME_BASE
                        ..ID_TRAY_VIDEO_VOLUME_BASE + VOLUME_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_video_volume(cmd - ID_TRAY_VIDEO_VOLUME_BASE)
                    }
                    cmd if (ID_TRAY_AUDIO_VOLUME_BASE
                        ..ID_TRAY_AUDIO_VOLUME_BASE + VOLUME_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_audio_volume(cmd - ID_TRAY_AUDIO_VOLUME_BASE)
                    }
                    // And where a sound starts, by the position the way it names was listed at —
                    // its own range, below both halves of the volume in the menu and past them
                    // here (see `AUDIO_SEEK_CHOICES`).
                    cmd if (ID_TRAY_AUDIO_SEEK_BASE
                        ..ID_TRAY_AUDIO_SEEK_BASE + AUDIO_SEEK_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_audio_seek(cmd - ID_TRAY_AUDIO_SEEK_BASE)
                    }
                    ID_TRAY_POSITION_FOLLOW => set_follow_cursor(true),
                    ID_TRAY_POSITION_BEST => set_follow_cursor(false),
                    // A way of keeping a preview off the hovered item, by the position it
                    // was listed at.
                    cmd if (ID_TRAY_AVOID_BASE
                        ..ID_TRAY_AVOID_BASE + AVOID_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_avoid_mode(cmd - ID_TRAY_AVOID_BASE)
                    }
                    // A delay of one of the three `Timing` submenus, by the position it was
                    // listed at. The three offer the same delays, so one table answers for
                    // all of them and each range is what says which setting was meant.
                    cmd if (ID_TRAY_DELAY_BASE
                        ..ID_TRAY_DELAY_BASE + TIMING_DELAY_CHOICES_MS.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_hover_delay(cmd - ID_TRAY_DELAY_BASE)
                    }
                    cmd if (ID_TRAY_REHOVER_DELAY_BASE
                        ..ID_TRAY_REHOVER_DELAY_BASE + TIMING_DELAY_CHOICES_MS.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_same_file_rehover_delay(cmd - ID_TRAY_REHOVER_DELAY_BASE)
                    }
                    cmd if (ID_TRAY_SETTLING_DELAY_BASE
                        ..ID_TRAY_SETTLING_DELAY_BASE + TIMING_DELAY_CHOICES_MS.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_settling_delay(cmd - ID_TRAY_SETTLING_DELAY_BASE)
                    }
                    // How often the loop looks at the pointer's world, by the position its
                    // item was listed at: the one command here that sets the app's own rate
                    // rather than a delay a hover waits out.
                    cmd if (ID_TRAY_TICK_BASE
                        ..ID_TRAY_TICK_BASE + TICK_CHOICES_MS.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_tick_ms(cmd - ID_TRAY_TICK_BASE)
                    }
                    ID_TRAY_OPEN_CONFIG => open_config_file(),
                    ID_TRAY_RESET_SETTINGS => reset_settings_from_tray(),
                    ID_TRAY_RESET_LISTS => reset_lists_from_tray(),
                    // A row of the `Codecs` submenu, by the position it was listed at. Only the
                    // rows this machine is missing and has a page for carry an id, and the row an
                    // id stands for is read again from the machine rather than kept from the menu
                    // build, so nothing has to be remembered between the click and the row (see
                    // `open_codec_page`).
                    cmd if (ID_TRAY_CODEC_BASE..ID_TRAY_CODEC_BASE + CODEC_COMMANDS)
                        .contains(&cmd) =>
                    {
                        open_codec_page(cmd - ID_TRAY_CODEC_BASE)
                    }
                    // How large a picture is drawn, by the position its item was listed at.
                    cmd if (ID_TRAY_SCALE_BASE
                        ..ID_TRAY_SCALE_BASE + BITMAP_SCALE_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_preview_scale(cmd - ID_TRAY_SCALE_BASE)
                    }
                    ID_TRAY_THEME_LIGHT => set_theme(TextTheme::Light),
                    ID_TRAY_THEME_DARK => set_theme(TextTheme::Dark),
                    // The `theme` folder's items, by the position the submenu gave
                    // them rather than any position in the folder.
                    cmd if (ID_TRAY_THEME_CUSTOM_BASE..ID_TRAY_IMAGE_CACHE_BASE).contains(&cmd) => {
                        set_theme_from_menu((cmd - ID_TRAY_THEME_CUSTOM_BASE) as usize)
                    }
                    ID_TRAY_MARKDOWN_RENDERED => set_markdown_mode(MarkdownMode::Rendered),
                    ID_TRAY_MARKDOWN_SOURCE => set_markdown_mode(MarkdownMode::Source),
                    ID_TRAY_RENDER_HTML => toggle_render_html(),
                    ID_TRAY_TYPE_IMAGES => toggle_preview_type(PreviewType::Images),
                    ID_TRAY_TYPE_VIDEOS => toggle_preview_type(PreviewType::Videos),
                    ID_TRAY_TYPE_AUDIO => toggle_preview_type(PreviewType::Audio),
                    ID_TRAY_TYPE_TEXT => toggle_preview_type(PreviewType::Text),
                    ID_TRAY_TYPE_EBOOK => toggle_preview_type(PreviewType::Ebook),
                    ID_TRAY_TYPE_ARCHIVES => toggle_preview_type(PreviewType::Archives),
                    ID_TRAY_TYPE_DOCUMENT => toggle_preview_type(PreviewType::Document),
                    ID_TRAY_TYPE_FONTS => toggle_preview_type(PreviewType::Fonts),
                    ID_TRAY_TYPE_DESIGN => toggle_preview_type(PreviewType::Design),
                    ID_TRAY_TYPE_VECTOR => toggle_preview_type(PreviewType::Vector),
                    // An Office engine's idle time, by the position it was listed at.
                    cmd if (ID_TRAY_ENGINE_IDLE_BASE
                        ..ID_TRAY_ENGINE_IDLE_BASE + ENGINE_IDLE_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_office_engine_idle(cmd - ID_TRAY_ENGINE_IDLE_BASE)
                    }
                    // The browser engine's idle time, the same way.
                    cmd if (ID_TRAY_WEBVIEW_IDLE_BASE
                        ..ID_TRAY_WEBVIEW_IDLE_BASE + ENGINE_IDLE_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_webview_idle(cmd - ID_TRAY_WEBVIEW_IDLE_BASE)
                    }
                    // And the render engine's, in the range of its own.
                    cmd if (ID_TRAY_LIBREOFFICE_IDLE_BASE
                        ..ID_TRAY_LIBREOFFICE_IDLE_BASE + ENGINE_IDLE_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_libreoffice_idle(cmd - ID_TRAY_LIBREOFFICE_IDLE_BASE)
                    }
                    // An away time for the `AFK Timer`, by the position it was listed at.
                    cmd if (ID_TRAY_AFK_TIMER_BASE
                        ..ID_TRAY_AFK_TIMER_BASE + AFK_TIMER_CHOICES_SECS.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_afk_timer(cmd - ID_TRAY_AFK_TIMER_BASE)
                    }
                    // A `Persistent` toggle, by the submenu it heads: the Office engines, then
                    // the render engine's, then the browser's.
                    cmd if (ID_TRAY_ENGINE_PERSISTENT_BASE..ID_TRAY_ENGINE_PERSISTENT_BASE + 3)
                        .contains(&cmd) =>
                    {
                        toggle_engine_persistent(cmd - ID_TRAY_ENGINE_PERSISTENT_BASE)
                    }
                    // Whether a video is decoded on the graphics card, which is the one row of
                    // the `Hardware Acceleration` submenu and so has an id of its own rather than
                    // a range.
                    ID_TRAY_VIDEO_HW_ACCEL => toggle_video_hw_accel(),
                    // A cache size, by the position it was listed at. Each cache is bounded by the
                    // sizes it offers rather than by the base of the next one, so an item of a
                    // submenu added past these is not read as a cache size.
                    cmd if (ID_TRAY_IMAGE_CACHE_BASE
                        ..ID_TRAY_IMAGE_CACHE_BASE + CACHE_SIZE_CHOICES_MB.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_image_cache_mb(cmd - ID_TRAY_IMAGE_CACHE_BASE)
                    }
                    cmd if (ID_TRAY_DOCUMENT_CACHE_BASE
                        ..ID_TRAY_DOCUMENT_CACHE_BASE + CACHE_SIZE_CHOICES_MB.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_document_cache_mb(cmd - ID_TRAY_DOCUMENT_CACHE_BASE)
                    }
                    cmd if (ID_TRAY_IMAGE_DISK_CACHE_BASE
                        ..ID_TRAY_IMAGE_DISK_CACHE_BASE + CACHE_SIZE_CHOICES_MB.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_image_disk_cache_mb(cmd - ID_TRAY_IMAGE_DISK_CACHE_BASE)
                    }
                    cmd if (ID_TRAY_GENERAL_DISK_CACHE_BASE
                        ..ID_TRAY_GENERAL_DISK_CACHE_BASE + CACHE_SIZE_CHOICES_MB.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_general_disk_cache_mb(cmd - ID_TRAY_GENERAL_DISK_CACHE_BASE)
                    }
                    // The ceiling one hover is answered under, by the position it was
                    // listed at.
                    cmd if (ID_TRAY_DECODE_BUDGET_BASE
                        ..ID_TRAY_DECODE_BUDGET_BASE + DECODE_BUDGET_CHOICES_GB.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_decode_budget_gb(cmd - ID_TRAY_DECODE_BUDGET_BASE)
                    }
                    // A share of the display a document is drawn at, by the position it
                    // was listed at: the vector scale, and the two page scales beside it.
                    cmd if (ID_TRAY_VECTOR_SCALE_BASE
                        ..ID_TRAY_VECTOR_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_vector_scale(cmd - ID_TRAY_VECTOR_SCALE_BASE)
                    }
                    cmd if (ID_TRAY_EBOOK_SCALE_BASE
                        ..ID_TRAY_EBOOK_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_ebook_scale(cmd - ID_TRAY_EBOOK_SCALE_BASE)
                    }
                    cmd if (ID_TRAY_DOCUMENT_SCALE_BASE
                        ..ID_TRAY_DOCUMENT_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_document_scale(cmd - ID_TRAY_DOCUMENT_SCALE_BASE)
                    }
                    cmd if (ID_TRAY_FONT_SCALE_BASE
                        ..ID_TRAY_FONT_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_font_scale(cmd - ID_TRAY_FONT_SCALE_BASE)
                    }
                    // And how much of the display a design document is drawn over, the same
                    // question asked of the picture a file keeps of the whole of one.
                    cmd if (ID_TRAY_DESIGN_SCALE_BASE
                        ..ID_TRAY_DESIGN_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_design_scale(cmd - ID_TRAY_DESIGN_SCALE_BASE)
                    }
                    // And how much of the display a page of text is measured in, the same
                    // question the rows above it are asked of a document.
                    cmd if (ID_TRAY_TEXT_SCALE_BASE
                        ..ID_TRAY_TEXT_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_text_scale(cmd - ID_TRAY_TEXT_SCALE_BASE)
                    }
                    // How large a video is drawn, by the position its item was listed at: the
                    // same shares the pictures above it are offered, in a range of their own
                    // because the two settings are read one each.
                    cmd if (ID_TRAY_VIDEO_SCALE_BASE
                        ..ID_TRAY_VIDEO_SCALE_BASE + BITMAP_SCALE_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_video_scale(cmd - ID_TRAY_VIDEO_SCALE_BASE)
                    }
                    // And how large an animated picture is drawn, the same way again. Which
                    // files that is — a GIF or a WebP that moves, a PNG that does — is settled
                    // by the file's own content when its hover is laid out.
                    cmd if (ID_TRAY_ANIMATED_SCALE_BASE
                        ..ID_TRAY_ANIMATED_SCALE_BASE + BITMAP_SCALE_CHOICES.len() as u16)
                        .contains(&cmd) =>
                    {
                        set_animated_scale(cmd - ID_TRAY_ANIMATED_SCALE_BASE)
                    }
                    ID_TRAY_FONT_400 => set_text_font_scale(400),
                    ID_TRAY_FONT_300 => set_text_font_scale(300),
                    ID_TRAY_FONT_250 => set_text_font_scale(250),
                    ID_TRAY_FONT_200 => set_text_font_scale(200),
                    ID_TRAY_FONT_175 => set_text_font_scale(175),
                    ID_TRAY_FONT_150 => set_text_font_scale(150),
                    ID_TRAY_FONT_125 => set_text_font_scale(125),
                    ID_TRAY_FONT_110 => set_text_font_scale(110),
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
    })
    .unwrap_or(LRESULT(0))
}

unsafe fn add_tray_icon(hwnd: HWND) -> bool {
    // Load the embedded icon resource (assets/icon.ico compiled via build.rs)
    let hicon = if let Ok(hmodule) = GetModuleHandleW(None) {
        let hinstance = HINSTANCE(hmodule.0);
        match LoadImageW(
            hinstance,
            // `MAKEINTRESOURCEW(1)`: with the high word zero, `LoadImageW`
            // reads this value as the ID of the resource to load rather than
            // as the address of a name, so it is an ID here and not a pointer.
            PCWSTR(std::ptr::without_provenance(1)),
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

/// Run the tray window and its message loop, which is where this app's main thread spends the
/// run.
///
/// The `trace` is the start's own (see `StartupTrace`), and what it is handed for is the two
/// steps between the launch and the icon a user is waiting for: the window the icon hangs on
/// and the icon itself. Everything else on this thread comes after that, and nothing of the
/// start is left between them — the housekeeping that used to be is on a thread of its own by
/// the time this is called (see `main`).
pub fn run_tray(trace: &mut StartupTrace) {
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
        trace.step("tray window");

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

        // The icon is in the tray, which is the last thing the start is between the launch and
        // anything the user can see: what the retries above cost shows in this step's own time.
        trace.step("tray icon");

        // Message loop.
        //
        // The loop blocks on the window's own messages rather than polling for them: a
        // window's messages are queued whether or not anyone looks, so there is nothing a
        // poll would find that a wait does not — and every wake this thread has is a
        // message. The icon's own clicks, the menu's commands, Explorer restarting, the
        // system coming back from sleep, and the quit the exit paths post are all of them,
        // and a poll is a hundred wakeups a second spent finding none.
        //
        // Nothing else ends this loop, and nothing has to: the threads that follow it are
        // signalled by the shutdown after it rather than by a flag looked at here.
        let mut msg = MSG::default();
        while RUNNING.load(Ordering::SeqCst) {
            // Zero is `WM_QUIT`, and −1 a failure that leaves the message empty: neither
            // is a message to dispatch, and both mean the loop is over.
            if GetMessageW(&mut msg, None, 0, 0).0 <= 0 {
                RUNNING.store(false, Ordering::SeqCst);
                break;
            }

            let _ = TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        // Cleanup
        remove_tray_icon(hwnd);
    }
}

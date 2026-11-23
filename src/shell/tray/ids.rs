//! The command ids the tray hands out, and the tables each menu lists.
//!
//! Every row and submenu range is allocated sequentially by `define_tray_ids!` starting at 1001,
//! eliminating manual `u16` hole-packing and making ID collisions impossible by construction
//! (verified by `tray::tests`).

use crate::config::config::{
    AudioSeek, AvoidMode, EngineIdle, PinEndOfFile, PreviewScale, TextTheme, TransparentBackground,
    VideoEngine, VOLUME_CHOICES,
};
use once_cell::sync::Lazy;
use std::sync::Mutex;
use windows::core::{w, PCWSTR};
use windows::Win32::UI::WindowsAndMessaging::WM_USER;

pub(super) const WM_TRAYICON: u32 = WM_USER + 1;
pub(super) const WM_TRAY_INSTALL_FINISHED: u32 = WM_USER + 2;
pub(super) const WM_TRAY_UPDATE_INSTALLED: u32 = WM_USER + 3;
#[allow(dead_code)]
pub(super) const WM_TRAY_TICK: u32 = WM_USER + 4;
pub(super) const TIMER_CHECK: usize = 1;
pub(super) const TIMER_DUE: usize = 2;
pub(super) const DUE_TICK_MS: u32 = 1000;

pub(super) const CODEC_COMMANDS: u16 = 64;
pub(super) const MPV_COMMANDS: u16 = 8;
pub(super) const INSTALL_COMMANDS: u16 = 7;
pub(super) const MAX_TRAY_CUSTOM_THEMES: usize = 200;

pub(super) const VIDEO_ENGINE_CHOICES: [VideoEngine; 5] = [
    VideoEngine::Best,
    VideoEngine::Mpv,
    VideoEngine::Ffmpeg,
    VideoEngine::Native,
    VideoEngine::Hybrid,
];

pub(super) const BACKGROUND_CHOICES: [TransparentBackground; 4] = [
    TransparentBackground::Transparent,
    TransparentBackground::Black,
    TransparentBackground::White,
    TransparentBackground::Checkerboard,
];

pub(super) const HTML_BACKGROUND_CHOICES: [TransparentBackground; 3] = [
    TransparentBackground::White,
    TransparentBackground::Black,
    TransparentBackground::Checkerboard,
];

pub(super) const DDS_BACKGROUND_CHOICES: [TransparentBackground; 2] =
    [TransparentBackground::Black, TransparentBackground::White];

pub(super) const ZOOM_BACKGROUND_CHOICES: [TransparentBackground; 3] = [
    TransparentBackground::Black,
    TransparentBackground::White,
    TransparentBackground::Checkerboard,
];

pub(super) const AUDIO_SEEK_CHOICES: [AudioSeek; 4] = [
    AudioSeek::Remember,
    AudioSeek::Start,
    AudioSeek::Middle,
    AudioSeek::Random,
];

pub(super) const AVOID_CHOICES: [AvoidMode; 4] = [
    AvoidMode::Off,
    AvoidMode::Filename,
    AvoidMode::FilenameColumn,
    AvoidMode::Details,
];

pub(super) const TIMING_DELAY_CHOICES_MS: [u64; 15] = [
    0, 25, 50, 100, 150, 200, 250, 300, 400, 500, 600, 700, 800, 900, 1000,
];

pub(super) const TICK_CHOICES_MS: [u64; 5] = [15, 31, 47, 63, 78];

pub(super) const PIN_END_OF_FILE_CHOICES: [PinEndOfFile; 3] = [
    PinEndOfFile::Pause,
    PinEndOfFile::Loop,
    PinEndOfFile::Continue,
];

pub(super) const DOCUMENT_SCALE_CHOICES: [PreviewScale; 5] = [
    PreviewScale::FitToScreen,
    PreviewScale::Percent(75),
    PreviewScale::Percent(50),
    PreviewScale::Percent(25),
    PreviewScale::Percent(10),
];

pub(super) const BITMAP_SCALE_CHOICES: [PreviewScale; 12] = [
    PreviewScale::FitToScreen,
    PreviewScale::FitToScreenReduced(75),
    PreviewScale::FitToScreenReduced(50),
    PreviewScale::FitToScreenReduced(25),
    PreviewScale::FitToScreenReduced(10),
    PreviewScale::Percent(400),
    PreviewScale::Percent(300),
    PreviewScale::Percent(200),
    PreviewScale::Percent(150),
    PreviewScale::Percent(100),
    PreviewScale::Percent(50),
    PreviewScale::Percent(25),
];

pub(super) const AUDIO_SCALE_CHOICES: [PreviewScale; 6] = [
    PreviewScale::Percent(25),
    PreviewScale::Percent(20),
    PreviewScale::Percent(15),
    PreviewScale::Percent(10),
    PreviewScale::Percent(7),
    PreviewScale::Percent(5),
];
/// The sizes `Memory → Image` offers, in megabytes, largest first — `16 GB` at
/// the top and a cache that holds nothing at the bottom — and the whole range the
/// setting allows, so a size a hand-edited `config.ini` asks for that is not one of
/// these is shown with nothing checked rather than rounded to one of them.
pub(super) const IMAGE_CACHE_SIZE_CHOICES_MB: [u32; 12] =
    [16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 0];

pub(super) const DISK_CACHE_SIZE_CHOICES_MB: [u32; 15] = [
    131072, 65536, 32768, 16384, 8192, 4096, 2048, 1024, 512, 256, 128, 64, 32, 16, 0,
];

pub(super) const DECODE_BUDGET_CHOICES_GB: [f32; 6] = [16.0, 8.0, 4.0, 2.0, 1.0, 0.5];

pub(super) const ANIMATED_WEBP_BUDGET_CHOICES_MB: [u32; 8] =
    [8192, 4096, 2048, 1024, 512, 256, 128, 64];

pub(super) const ENGINE_IDLE_CHOICES: [EngineIdle; 9] = [
    EngineIdle::Indefinite,
    EngineIdle::Seconds(3600),
    EngineIdle::Seconds(1800),
    EngineIdle::Seconds(600),
    EngineIdle::Seconds(300),
    EngineIdle::Seconds(60),
    EngineIdle::Seconds(30),
    EngineIdle::Seconds(0),
    EngineIdle::Away,
];

pub(super) const AFK_TIMER_CHOICES_SECS: [u64; 7] = [3600, 1800, 600, 300, 60, 30, 15];

/// Sequential, collision-free allocator for every tray command ID and submenu range.
/// Radio-group pairs used with `CheckMenuRadioItem` (e.g. `TRIGGER_DISABLE`/`TRIGGER_ENABLE`,
/// `POSITION_FOLLOW`/`POSITION_BEST`, `ENGINE_OFFICE_MS`/`ENGINE_OFFICE_LIBRE`) must stay adjacent.
macro_rules! define_tray_ids {
    (
        base = $start:expr;
        $( $name:ident : $width:expr ),* $(,)?
    ) => {
        #[allow(non_camel_case_types)]
        #[repr(usize)]
        enum TrayIdSlot {
            $( $name, )*
            _Count,
        }

        const TRAY_ID_WIDTHS: [u16; TrayIdSlot::_Count as usize] = [
            $( $width as u16, )*
        ];

        const TRAY_ID_STARTS: [u16; TrayIdSlot::_Count as usize] = {
            let mut starts = [0u16; TrayIdSlot::_Count as usize];
            let mut cur: u16 = $start;
            let mut i = 0;
            while i < starts.len() {
                starts[i] = cur;
                cur += TRAY_ID_WIDTHS[i];
                i += 1;
            }
            starts
        };

        $(
            pub(super) const $name: u16 = TRAY_ID_STARTS[TrayIdSlot::$name as usize];
        )*

        #[cfg(test)]
        pub(super) const ALL_COMMAND_SPANS: &[(&str, u16, u16)] = &[
            $( (stringify!($name), $name, $name + ($width as u16)), )*
        ];
    };
}

define_tray_ids! {
    base = 1001;

    // Core & system rows
    ID_TRAY_EXIT : 1,
    ID_TRAY_STARTUP : 1,
    ID_TRAY_ENABLE : 1,
    ID_TRAY_UPDATE : 1,
    ID_TRAY_OPEN_CONFIG : 1,
    ID_TRAY_RESET_SETTINGS : 1,
    ID_TRAY_RESET_LISTS : 1,
    ID_TRAY_CHECK_UPDATES : 1,
    ID_TRAY_SILENT_AUTO_UPDATE : 1,

    // Pin mode rows
    ID_TRAY_PIN : 1,
    ID_TRAY_PIN_RESTORE : 1,
    ID_TRAY_PIN_ALWAYS_ON_TOP : 1,
    ID_TRAY_PIN_UPDATE : 1,
    ID_TRAY_PIN_UPDATE_HOVER : 1,
    ID_TRAY_PIN_UPDATE_TOP : 1,
    ID_TRAY_PIN_PAUSE_AUDIO : 1,
    ID_TRAY_PIN_PAUSE_VIDEO : 1,
    ID_TRAY_PIN_MINIMIZE_BUBBLE : 1,
    ID_TRAY_PIN_MINIMIZE_TASKBAR : 1,
    ID_TRAY_PIN_SCALING_UNTIL_RESIZED : 1,
    ID_TRAY_PIN_SCALING_USE_SETTINGS : 1,
    ID_TRAY_PIN_NAV_IMAGE_ALL : 1,
    ID_TRAY_PIN_NAV_IMAGE_CATEGORY : 1,
    ID_TRAY_PIN_NAV_VIDEO_ALL : 1,
    ID_TRAY_PIN_NAV_VIDEO_CATEGORY : 1,
    ID_TRAY_PIN_NAV_AUDIO_ALL : 1,
    ID_TRAY_PIN_NAV_AUDIO_CATEGORY : 1,
    ID_TRAY_PIN_NAV_OTHERS_ALL : 1,
    ID_TRAY_PIN_NAV_OTHERS_CATEGORY : 1,
    ID_TRAY_PIN_MODE_AUDIO_CAMERA_FOLLOWS : 1,
    ID_TRAY_PIN_MODE_AUDIO_PLAYLIST_ALWAYS_OPEN : 1,

    // Trigger key & placement (radio pairs kept adjacent for CheckMenuRadioItem)
    ID_TRAY_TRIGGER_DISABLE : 1,
    ID_TRAY_TRIGGER_ENABLE : 1,
    ID_TRAY_TRIGGER_ENABLED : 1,
    ID_TRAY_TRIGGER_AFFECT_PIN : 1,
    ID_TRAY_PRIORITIZE_KEYBOARD : 1,
    ID_TRAY_POSITION_FOLLOW : 1,
    ID_TRAY_POSITION_BEST : 1,
    ID_TRAY_BYPASS_PLACEMENT : 1,

    // General preview hosts
    ID_TRAY_BROWSE_DIALOG_PREVIEW_ENABLED : 1,
    ID_TRAY_EVERYTHING_PREVIEWS : 1,
    ID_TRAY_Q_DIR_PREVIEWS : 1,

    // Preview type gates
    ID_TRAY_TYPE_IMAGES : 1,
    ID_TRAY_TYPE_VIDEOS : 1,
    ID_TRAY_TYPE_AUDIO : 1,
    ID_TRAY_TYPE_TEXT : 1,
    ID_TRAY_TYPE_EBOOK : 1,
    ID_TRAY_TYPE_ARCHIVES : 1,
    ID_TRAY_TYPE_DOCUMENT : 1,
    ID_TRAY_TYPE_VECTOR : 1,
    ID_TRAY_TYPE_FONTS : 1,
    ID_TRAY_TYPE_DESIGN : 1,
    ID_TRAY_TYPE_MODELS : 1,

    // Media & text switches
    ID_TRAY_VIDEO_SUBTITLES : 1,
    ID_TRAY_VIDEO_LOOP : 1,
    ID_TRAY_AUDIO_LOOP : 1,
    ID_TRAY_VIDEO_SHUFFLE : 1,
    ID_TRAY_AUDIO_SHUFFLE : 1,
    ID_TRAY_IMAGE_SHUFFLE : 1,
    ID_TRAY_NORMALIZE_VOLUME : 1,
    ID_TRAY_NORMALIZE_VIDEO_VOLUME : 1,
    ID_TRAY_REMEMBER_VOLUME : 1,
    ID_TRAY_REMEMBER_VIDEO_VOLUME : 1,
    ID_TRAY_VIDEO_HW_ACCEL : 1,
    ID_TRAY_THEME_LIGHT : 1,
    ID_TRAY_THEME_DARK : 1,
    ID_TRAY_MARKDOWN_RENDERED : 1,
    ID_TRAY_MARKDOWN_SOURCE : 1,
    ID_TRAY_RENDER_HTML : 1,
    ID_TRAY_TEXT_SHOW_EXTENSIONLESS : 1,
    ID_TRAY_TEXT_DONT_PREVIEW_SECRETS : 1,
    ID_TRAY_PDF_AUTOFOCUS : 1,
    ID_TRAY_PDF_FULL_MODE : 1,
    // The one row of `Preview Settings` that opens a window rather than throwing a switch:
    // the window a user's own previews are made in (see `config::config::custom_presets`).
    ID_TRAY_CUSTOM_PREVIEW : 1,

    // Text font sizes
    ID_TRAY_FONT_100 : 1,
    ID_TRAY_FONT_110 : 1,
    ID_TRAY_FONT_125 : 1,
    ID_TRAY_FONT_150 : 1,
    ID_TRAY_FONT_175 : 1,
    ID_TRAY_FONT_200 : 1,
    ID_TRAY_FONT_250 : 1,
    ID_TRAY_FONT_300 : 1,
    ID_TRAY_FONT_400 : 1,
    ID_TRAY_FONT_90 : 1,
    ID_TRAY_FONT_80 : 1,
    ID_TRAY_FONT_70 : 1,

    // Engine choices (Office radio pair kept adjacent for CheckMenuRadioItem)
    ID_TRAY_ENGINE_OFFICE_MS : 1,
    ID_TRAY_ENGINE_OFFICE_LIBRE : 1,
    ID_TRAY_VIDEO_ENGINE_FALLBACK : 1,

    // Submenu ranges
    ID_TRAY_VIDEO_ENGINE_BASE : VIDEO_ENGINE_CHOICES.len(),
    ID_TRAY_IMAGE_BACKGROUND_BASE : BACKGROUND_CHOICES.len(),
    ID_TRAY_VECTOR_BACKGROUND_BASE : BACKGROUND_CHOICES.len(),
    ID_TRAY_FONT_BACKGROUND_BASE : BACKGROUND_CHOICES.len(),
    ID_TRAY_DESIGN_BACKGROUND_BASE : BACKGROUND_CHOICES.len(),
    ID_TRAY_DDS_BACKGROUND_BASE : DDS_BACKGROUND_CHOICES.len(),
    ID_TRAY_HTML_BACKGROUND_BASE : HTML_BACKGROUND_CHOICES.len(),
    ID_TRAY_ZOOM_BACKGROUND_BASE : ZOOM_BACKGROUND_CHOICES.len(),
    ID_TRAY_VIDEO_VOLUME_BASE : VOLUME_CHOICES.len(),
    ID_TRAY_AUDIO_VOLUME_BASE : VOLUME_CHOICES.len(),
    ID_TRAY_AUDIO_SEEK_BASE : AUDIO_SEEK_CHOICES.len(),
    ID_TRAY_PIN_MODE_AUDIO_SEEK_BASE : AUDIO_SEEK_CHOICES.len(),
    ID_TRAY_PIN_MODE_VIDEO_END_OF_FILE_BASE : PIN_END_OF_FILE_CHOICES.len(),
    ID_TRAY_PIN_MODE_AUDIO_END_OF_FILE_BASE : PIN_END_OF_FILE_CHOICES.len(),
    ID_TRAY_AVOID_BASE : AVOID_CHOICES.len(),
    ID_TRAY_DELAY_BASE : TIMING_DELAY_CHOICES_MS.len(),
    ID_TRAY_REHOVER_DELAY_BASE : TIMING_DELAY_CHOICES_MS.len(),
    ID_TRAY_SETTLING_DELAY_BASE : TIMING_DELAY_CHOICES_MS.len(),
    ID_TRAY_TICK_BASE : TICK_CHOICES_MS.len(),
    ID_TRAY_AFK_TIMER_BASE : AFK_TIMER_CHOICES_SECS.len(),
    ID_TRAY_SCALE_BASE : BITMAP_SCALE_CHOICES.len(),
    ID_TRAY_VIDEO_SCALE_BASE : BITMAP_SCALE_CHOICES.len(),
    ID_TRAY_ANIMATED_SCALE_BASE : BITMAP_SCALE_CHOICES.len(),
    ID_TRAY_AUDIO_SCALE_BASE : AUDIO_SCALE_CHOICES.len(),
    ID_TRAY_VECTOR_SCALE_BASE : DOCUMENT_SCALE_CHOICES.len(),
    ID_TRAY_EBOOK_SCALE_BASE : DOCUMENT_SCALE_CHOICES.len(),
    ID_TRAY_DOCUMENT_SCALE_BASE : DOCUMENT_SCALE_CHOICES.len(),
    ID_TRAY_FONT_SCALE_BASE : DOCUMENT_SCALE_CHOICES.len(),
    ID_TRAY_MODEL_SCALE_BASE : DOCUMENT_SCALE_CHOICES.len(),
    ID_TRAY_DESIGN_SCALE_BASE : DOCUMENT_SCALE_CHOICES.len(),
    ID_TRAY_TEXT_SCALE_BASE : DOCUMENT_SCALE_CHOICES.len(),
    ID_TRAY_IMAGE_CACHE_BASE : IMAGE_CACHE_SIZE_CHOICES_MB.len(),
    ID_TRAY_DOCUMENT_CACHE_BASE : DISK_CACHE_SIZE_CHOICES_MB.len(),
    ID_TRAY_IMAGE_DISK_CACHE_BASE : DISK_CACHE_SIZE_CHOICES_MB.len(),
    ID_TRAY_GENERAL_DISK_CACHE_BASE : DISK_CACHE_SIZE_CHOICES_MB.len(),
    ID_TRAY_DECODE_BUDGET_BASE : DECODE_BUDGET_CHOICES_GB.len(),
    ID_TRAY_ANIMATED_WEBP_BUDGET_BASE : ANIMATED_WEBP_BUDGET_CHOICES_MB.len(),
    ID_TRAY_ENGINE_IDLE_BASE : ENGINE_IDLE_CHOICES.len(),
    ID_TRAY_LIBREOFFICE_IDLE_BASE : ENGINE_IDLE_CHOICES.len(),
    ID_TRAY_WEBVIEW_IDLE_BASE : ENGINE_IDLE_CHOICES.len(),
    ID_TRAY_MPV_IDLE_BASE : ENGINE_IDLE_CHOICES.len(),
    ID_TRAY_F3D_IDLE_BASE : ENGINE_IDLE_CHOICES.len(),
    ID_TRAY_CODEC_BASE : CODEC_COMMANDS,
    ID_TRAY_MPV_BASE : MPV_COMMANDS,
    ID_TRAY_INSTALL_BASE : INSTALL_COMMANDS,
    ID_TRAY_THEME_CUSTOM_BASE : MAX_TRAY_CUSTOM_THEMES,
}

pub(crate) const FONT_SIZE_CHOICES: [(u32, u16); 12] = [
    (400, ID_TRAY_FONT_400),
    (300, ID_TRAY_FONT_300),
    (250, ID_TRAY_FONT_250),
    (200, ID_TRAY_FONT_200),
    (175, ID_TRAY_FONT_175),
    (150, ID_TRAY_FONT_150),
    (125, ID_TRAY_FONT_125),
    (110, ID_TRAY_FONT_110),
    (100, ID_TRAY_FONT_100),
    (90, ID_TRAY_FONT_90),
    (80, ID_TRAY_FONT_80),
    (70, ID_TRAY_FONT_70),
];

pub(super) const TRAY_CLASS: PCWSTR = w!("RustHoverPreviewTrayClass");

pub(super) static TRAY_HWND: std::sync::atomic::AtomicIsize =
    std::sync::atomic::AtomicIsize::new(0);
pub(super) static TASKBAR_CREATED: std::sync::atomic::AtomicU32 =
    std::sync::atomic::AtomicU32::new(0);

/// Custom themes the `Theme` submenu last listed, indexed by offset from `ID_TRAY_THEME_CUSTOM_BASE`.
pub(super) static TRAY_CUSTOM_THEMES: Lazy<Mutex<Vec<TextTheme>>> =
    Lazy::new(|| Mutex::new(Vec::new()));

use crate::app::{dialogs, updates};
use crate::config::config::{
    sanitize_decode_budget_gb, sanitize_document_cache_mb, sanitize_image_cache_mb,
    sanitize_image_disk_cache_mb, sanitize_text_font_scale_percent, sanitize_tick_ms, AudioSeek,
    AvoidMode, EngineIdle, MarkdownMode, OfficeEngine, PreviewScale, PreviewType, TextTheme,
    TransparentBackground, TriggerKeyMode, DEFAULT_AFK_TIMER_SECS, DEFAULT_ANIMATED_SCALE,
    DEFAULT_AUDIO_SEEK, DEFAULT_AVOID_MODE, DEFAULT_DDS_BACKGROUND, DEFAULT_DECODE_BUDGET_GB,
    DEFAULT_DESIGN_BACKGROUND, DEFAULT_DESIGN_SCALE, DEFAULT_DOCUMENT_CACHE_MB,
    DEFAULT_DOCUMENT_SCALE, DEFAULT_EBOOK_SCALE, DEFAULT_FOLLOW_CURSOR, DEFAULT_FONT_BACKGROUND,
    DEFAULT_FONT_SCALE, DEFAULT_HOVER_DELAY_MS, DEFAULT_IMAGE_BACKGROUND, DEFAULT_IMAGE_CACHE_MB,
    DEFAULT_IMAGE_DISK_CACHE_MB, DEFAULT_LIBREOFFICE_IDLE_SECS, DEFAULT_OFFICE_ENGINE,
    DEFAULT_OFFICE_ENGINE_IDLE_SECS, DEFAULT_PREVIEW_SCALE, DEFAULT_SAME_FILE_REHOVER_DELAY_MS,
    DEFAULT_SETTLING_DELAY_MS, DEFAULT_TEXT_FONT_SCALE_PERCENT, DEFAULT_TICK_MS,
    DEFAULT_VECTOR_BACKGROUND, DEFAULT_VECTOR_SCALE, DEFAULT_VIDEO_SCALE, DEFAULT_VIDEO_VOLUME,
    DEFAULT_WEBVIEW_IDLE_SECS, DEFAULT_AUDIO_VOLUME, VOLUME_CHOICES,
};
use crate::config::theme_files;
use crate::engines::document_cache;
use crate::engines::libreoffice_render;
use crate::engines::office_render;
use crate::engines::webview_preview;
use crate::formats::codecs::{self, refresh as refresh_codecs, Row};
use crate::shell::explorer_hook;
use crate::text::text_theme;
use crate::ui::preview_window::{refresh_preview, refresh_preview_types, trim_image_cache};
use crate::{app::startup, StartupTrace, CONFIG, RUNNING};
use once_cell::sync::Lazy;
use std::os::windows::ffi::OsStrExt;
use std::sync::atomic::Ordering;
use std::sync::Mutex;
use windows::core::{w, PCWSTR, PWSTR};
use windows::Win32::Foundation::{BOOL, HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Shell::{
    ShellExecuteW, Shell_NotifyIconW, NIF_ICON, NIF_MESSAGE, NIF_TIP, NIM_ADD, NIM_DELETE,
    NOTIFYICONDATAW,
};
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CheckMenuRadioItem, CreatePopupMenu, CreateWindowExW, DefWindowProcW, DestroyMenu,
    DispatchMessageW, GetCursorPos, GetMenuItemCount, GetMessageW, InsertMenuItemW, LoadImageW,
    PostQuitMessage, RegisterClassExW, RegisterWindowMessageW, SetForegroundWindow, TrackPopupMenu,
    TranslateMessage, CS_HREDRAW, CS_VREDRAW, HICON, HMENU, IMAGE_ICON, LR_DEFAULTSIZE, LR_SHARED,
    MENUITEMINFOW, MENU_ITEM_FLAGS, MFT_STRING, MF_BYCOMMAND, MF_CHECKED, MF_GRAYED, MF_POPUP,
    MF_SEPARATOR, MF_STRING, MF_UNCHECKED, MIIM_ID, MIIM_STRING, MIIM_SUBMENU, MSG,
    PBT_APMRESUMEAUTOMATIC, PBT_APMRESUMESUSPEND, SW_SHOWNORMAL, TPM_BOTTOMALIGN, TPM_LEFTALIGN,
    WM_COMMAND, WM_DESTROY, WM_LBUTTONUP, WM_POWERBROADCAST, WM_RBUTTONUP, WM_USER, WNDCLASSEXW,
    WS_EX_TOOLWINDOW, WS_POPUP,
};

const WM_TRAYICON: u32 = WM_USER + 1;
const ID_TRAY_EXIT: u16 = 1001;
const ID_TRAY_STARTUP: u16 = 1002;
const ID_TRAY_ENABLE: u16 = 1003;
const ID_TRAY_TRIGGER_DISABLE: u16 = 1005; // Hold the trigger key to stop previews
const ID_TRAY_TRIGGER_ENABLE: u16 = 1006; // Hold the trigger key to allow previews
const ID_TRAY_TRIGGER_ENABLED: u16 = 1068; // Whether the trigger key is watched at all
/// The `Engine → Select Engine → Office` pair: which engine an Office document's page is
/// asked of — the application that owns the format, or the render engine beside it. Two ids
/// rather than a range, the way the trigger key's mode has two, and they sit in the gap the
/// update row at 1007 leaves before the volume block at 1010.
const ID_TRAY_ENGINE_OFFICE_MS: u16 = 1008;
const ID_TRAY_ENGINE_OFFICE_LIBRE: u16 = 1009;
/// The `Background` submenu: one command per backdrop it offers, in the order it
/// lists them, for each of the five kinds of preview it keeps apart — a picture's
/// backdrop, a vector drawing's, a font specimen's, a texture's, and a design
/// document's. Each half is a base plus the position the choice was listed at, so one
/// table and one builder serve all of them, and each range is four wide and apart from
/// the others, which is what keeps an item of one from being read as a choice of another.
const ID_TRAY_IMAGE_BACKGROUND_BASE: u16 = 1023;
/// The second half, for a vector drawing — an SVG document the browser draws, or a
/// metafile the drawing layer replays: one page's backdrop for both halves of that kind,
/// since what stands behind a drawing is the same question either way.
const ID_TRAY_VECTOR_BACKGROUND_BASE: u16 = 1054;
/// The third half of the `Background` submenu, for a font specimen — which is a page of its
/// own and so has a backdrop of its own, the same way a document does.
///
/// It sits beside the texture's half rather than in the block the first two are in: every
/// id four wide in that block belongs to something else — 1058 to 1061 is where the text
/// preview's `Full Mode` item is, which is where this range was to begin with, so the last
/// two backdrops of a specimen were answered as full mode being switched on.
const ID_TRAY_FONT_BACKGROUND_BASE: u16 = 1204;
/// The fourth, for a `.dds` texture: a texture's alpha channel is as often a mask or a
/// channel nobody filled in as it is transparency, so what is drawn behind one is a
/// question of its own.
const ID_TRAY_DDS_BACKGROUND_BASE: u16 = 1200;
/// The fifth, for a design document: what its preview is made of is the picture the file
/// keeps of the whole document — a merged image, or the flattened document a project
/// container holds — and that picture's transparency is the document's own, so what stands
/// behind one is a question of its own.
const ID_TRAY_DESIGN_BACKGROUND_BASE: u16 = 1208;
/// The backdrops a half of the `Background` submenu offers, in the order it lists
/// them, with `Transparent` at the top: the whole range the setting holds, so nothing
/// a hand-edited `config.ini` can ask for is left unmarked.
const BACKGROUND_CHOICES: [TransparentBackground; 4] = [
    TransparentBackground::Transparent,
    TransparentBackground::Black,
    TransparentBackground::White,
    TransparentBackground::Checkerboard,
];
/// The backdrops the `DDS Background` half offers, which are two of the four rather than all
/// of them: a texture's alpha channel is as often a mask, a height or a roughness as it is
/// transparency (see `dds_image`), so what is drawn behind one is a page to read the channels
/// against rather than a hole to look through — and the two backdrops that show what stands
/// behind a preview are the two that page has no use for. A file that names one of them
/// anyway, from when this half listed all four, is read as the backdrop the setting starts at
/// rather than kept as a value the menu has no item for (see `sanitize_dds_background`).
const DDS_BACKGROUND_CHOICES: [TransparentBackground; 2] =
    [TransparentBackground::Black, TransparentBackground::White];
/// The two halves of the `Volume` submenu: one command per level it offers, in the order it
/// lists them, the video's range first and the sound's beside it. Each half is a base plus the
/// position a level was listed at, so one table and one builder serve both — the arrangement
/// the `Background`, `Scale` and `Timing` submenus already have — and the two ranges are ten
/// wide and apart from each other and from everything else, which is what keeps a level of one
/// from being read as a level of the other.
///
/// They sit in the stretch between the disk-cache range and the decode budget's, which is the
/// one gap this block had left. The six single ids they replace (1010 to 1015) are gone with
/// them: a volume is a level of one of two lists now rather than a name of its own.
const ID_TRAY_VIDEO_VOLUME_BASE: u16 = 1360;
const ID_TRAY_AUDIO_VOLUME_BASE: u16 = 1370;
/// The `Volume → Audio Seek` submenu: one command per way a sound can be started, in the order
/// it lists them. It is the `Volume` submenu's third item, below the two halves above it,
/// because the two questions belong together — a sound is heard at a level and from a place,
/// and both are answered the moment a hover starts its player rather than while one is playing.
///
/// Its range sits in the slack the `Decode Budget` ceilings leave before the `Vector Scaling`
/// half begins, and it is four wide because there are four ways a sound can be started. That it
/// is past the volume halves on the number line rather than beside them is the whole of what
/// keeps a level of either half from being read as a way of starting a sound — see
/// `the_two_volume_submenus_carry_a_range_apiece`, which holds all three away from the sounds
/// gate and from each other.
const ID_TRAY_AUDIO_SEEK_BASE: u16 = 1390;
/// The ways a sound can be started, in the order the submenu lists them: where it was left the
/// last time it was hovered, which is where the setting starts, then its beginning, its middle,
/// and anywhere in it at all (see `AudioSeek`).
const AUDIO_SEEK_CHOICES: [AudioSeek; 4] = [
    AudioSeek::Remember,
    AudioSeek::Start,
    AudioSeek::Middle,
    AudioSeek::Random,
];
const ID_TRAY_POSITION_FOLLOW: u16 = 1020; // Follow cursor
const ID_TRAY_POSITION_BEST: u16 = 1021; // Best position
/// The `Placement → Avoid` submenu: one command per way a preview is kept off the
/// item it is about, in the order it lists them. The ids sit in a range of their own
/// in the slack the scaling ranges leave, so a way of avoiding is never read as a
/// share of a preview's size.
const ID_TRAY_AVOID_BASE: u16 = 1410;
/// The ways the `Avoid` submenu offers, in the order it lists them: nothing avoided,
/// the name where it is drawn, the column the name is drawn in, and every column the
/// item draws.
const AVOID_CHOICES: [AvoidMode; 4] = [
    AvoidMode::Off,
    AvoidMode::Filename,
    AvoidMode::FilenameColumn,
    AvoidMode::Details,
];
/// The `Timing` submenus that list a delay each — `Delay`, `Rehover Delay` and
/// `Settling Delay`, in the order the menu lists them — one range apiece, each as wide
/// as the delays it offers. They sit past every other range the app hands out, so a
/// delay is never read as a size, a share or a backdrop.
const ID_TRAY_DELAY_BASE: u16 = 1455;
const ID_TRAY_REHOVER_DELAY_BASE: u16 = 1470;
const ID_TRAY_SETTLING_DELAY_BASE: u16 = 1485;
/// The one row of `Timing` that is a switch rather than a delay or a key: whether the
/// keyboard driving Explorer holds a parked pointer back instead of sharing the screen
/// with it. It sits past every range the app hands out up to the tick menu, so it is never
/// read as one of their items.
const ID_TRAY_PRIORITIZE_KEYBOARD: u16 = 1515;
/// The delays those three submenus offer, in the order they list them: no wait at all
/// at the top, where the app starts, and a whole second at the bottom. A delay a
/// hand-edited `config.ini` holds that is not one of these is shown with nothing checked
/// rather than rounded to the nearest.
const TIMING_DELAY_CHOICES_MS: [u64; 15] = [
    0, 25, 50, 100, 150, 200, 250, 300, 400, 500, 600, 700, 800, 900, 1000,
];
/// The `Performance → Tick` submenu: one command per rate the loop may run at, in the
/// order it lists them. It is the app's own rate rather than a delay a hover waits out,
/// and its range sits past every other range the app hands out so that a tick is never
/// read as a delay, a size or a share.
const ID_TRAY_TICK_BASE: u16 = 1500;
/// The ticks the `Tick` submenu offers, in the order it lists them: one system tick at
/// the top, where the app starts, and five of them at the bottom.
///
/// Whole system ticks rather than round numbers, because that is what a wait is
/// honoured in — the loop wakes on the system's clock, so a number that falls between
/// two of its ticks spends the same time as the one below it and reads as a step that
/// changed nothing. A tick a hand-edited `config.ini` holds that is not one of these is
/// shown with nothing checked rather than rounded to the nearest.
const TICK_CHOICES_MS: [u64; 5] = [15, 31, 47, 63, 78];
const ID_TRAY_OPEN_CONFIG: u16 = 1040;
/// The two rows inside the `Config.ini` submenu, beside the one that opens the file: the
/// first puts every setting back at what this build recommends and leaves the extension
/// lists alone, the second puts the lists back and leaves every other setting alone.
///
/// They sit between every range the submenus share and the `Codecs` commands above them, so
/// neither can be read as a click on one of those.
const ID_TRAY_RESET_SETTINGS: u16 = 1520;
const ID_TRAY_RESET_LISTS: u16 = 1521;
/// The rows of the `Codecs` submenu, numbered as one list across its three groups: only the
/// rows this machine is missing and has a page for are given an id at all, and this is where
/// those ids begin. The range is wider than the list is long, so that a row added to any of
/// the three groups is still inside it.
const ID_TRAY_CODEC_BASE: u16 = 1600;
const CODEC_COMMANDS: u16 = 64;
/// The row above `Run at Startup`, which is in the menu only while a newer release is
/// waiting: it puts the installer `updates` fetched on, and the app ends itself as the
/// installer takes over rather than being the copy that has to be terminated.
const ID_TRAY_UPDATE: u16 = 1007;
/// The `Scaling → Image Scaling` submenu: one command per share of its own size a
/// picture is drawn at, in the order it lists them — the first of the two bitmap
/// submenus, each with a range of its own so a click on one is never read as a click
/// on the other.
const ID_TRAY_SCALE_BASE: u16 = 1041;
/// `Vector Scaling`, the second of them, in the range the `Scaling` submenus share: a
/// drawing is asked for a share of the display rather than for a share of a size the file
/// asks for, and both halves of the kind — a document the browser draws and a metafile the
/// drawing layer replays — are asked with this one setting.
const ID_TRAY_VECTOR_SCALE_BASE: u16 = 1400;
/// The `Ebook Scaling` and `Document Scaling` submenus beside it, each listing the same
/// shares: they sit in the slack the `Avoid` items leave, so a share of the display is
/// never read as a way of avoiding the item a preview is about.
const ID_TRAY_EBOOK_SCALE_BASE: u16 = 1405;
const ID_TRAY_DOCUMENT_SCALE_BASE: u16 = 1415;
/// `Font Scaling`, the fourth of them, in the range after the Document one: a specimen is
/// drawn at a share of the display the same way a document is.
const ID_TRAY_FONT_SCALE_BASE: u16 = 1420;
/// `Video Scaling`, the `Image Scaling` submenu's twin below it, in the range directly
/// after the font one: it lists the same shares, so the two share a table and a builder,
/// and it is a range of its own because a click on a video's scale is never a click on a
/// picture's.
const ID_TRAY_VIDEO_SCALE_BASE: u16 = 1425;
/// `Animated Scaling`, the third of them, in the range after the video one: an animated
/// picture is a bitmap like the two above it, so it lists the same shares through the
/// same builder, and the range is its own because what moves has a size apart from what
/// does not.
const ID_TRAY_ANIMATED_SCALE_BASE: u16 = 1435;
/// `Design Scaling`, in the range after the animated one: a design document is previewed
/// from a picture the file keeps of the whole of itself, so it is asked for a share of the
/// display the way a page is rather than for a share of its own size.
const ID_TRAY_DESIGN_SCALE_BASE: u16 = 1445;
/// The shares of the display every `… Scaling` submenu offers, in the order it lists
/// them: the whole room a document can be given at the top, then the shares of it a
/// document is asked for below. What differs between the settings is where they start —
/// `50`, half the display, for a font specimen, and `Fit to Screen` for a drawing, a page
/// and a design document.
const DOCUMENT_SCALE_CHOICES: [PreviewScale; 5] = [
    PreviewScale::FitToScreen,
    PreviewScale::Percent(75),
    PreviewScale::Percent(50),
    PreviewScale::Percent(25),
    PreviewScale::Percent(10),
];
/// The shares the `Image Scaling` and `Video Scaling` submenus offer, in the order they
/// list them: a bitmap is drawn at a share of its own size rather than of the display, so
/// the percentages are the ones that mean something for one. Nothing is marked as the
/// default here — the default is passed to the labels rather than written into the table,
/// because which share a setting starts at is the setting's own business.
const BITMAP_SCALE_CHOICES: [PreviewScale; 8] = [
    PreviewScale::FitToScreen,
    PreviewScale::Percent(400),
    PreviewScale::Percent(300),
    PreviewScale::Percent(200),
    PreviewScale::Percent(150),
    PreviewScale::Percent(100),
    PreviewScale::Percent(50),
    PreviewScale::Percent(25),
];
const ID_TRAY_THEME_LIGHT: u16 = 1050; // Atom One Light
const ID_TRAY_THEME_DARK: u16 = 1051; // One Dark Pro
const ID_TRAY_MARKDOWN_RENDERED: u16 = 1052; // Rendered document
const ID_TRAY_MARKDOWN_SOURCE: u16 = 1053; // Highlighted Markdown source
const ID_TRAY_TEXT_FULL_MODE: u16 = 1061; // Text previews scroll/select on/off
/// The `Preview Types` submenu, one command per kind of preview.
const ID_TRAY_TYPE_IMAGES: u16 = 1062;
const ID_TRAY_TYPE_VIDEOS: u16 = 1063;
const ID_TRAY_TYPE_TEXT: u16 = 1064;
/// The `Ebook` gate: the pages this app reads and draws itself, which is every PDF.
const ID_TRAY_TYPE_EBOOK: u16 = 1065;
const ID_TRAY_TYPE_ARCHIVES: u16 = 1066;
/// The `Document` gate: the pages drawn for a document, whether the application that owns
/// the format drew one or an installed render engine did — the one switch both halves of the
/// kind answer to (see `PreviewType`).
const ID_TRAY_TYPE_DOCUMENT: u16 = 1067;
/// The `Fonts` gate beside it, under the same `Preview Types` submenu.
const ID_TRAY_TYPE_FONTS: u16 = 1070;
/// The `Design` gate beside those, under the same submenu.
const ID_TRAY_TYPE_DESIGN: u16 = 1071;
/// The `Vector` gate, for the drawings that are not pictures: SVG documents, which are the
/// kind SVG documents have always had — the id is the one this gate carried under that name
/// — and the metafiles and encapsulated PostScript files the same kind grew to hold.
const ID_TRAY_TYPE_VECTOR: u16 = 1069;
/// The `Audio` gate: the sounds this app plays, which are the one kind of preview that is
/// heard rather than looked at — and the reason `Volume` has two halves (see
/// `ID_TRAY_AUDIO_VOLUME_BASE`).
///
/// Its id sits in the slack the font sizes leave rather than beside the other gates: 1072 is
/// where the text preview's own `100%` begins, and the block of gates there is two sizes wide.
const ID_TRAY_TYPE_AUDIO: u16 = 1098;
/// The `Cache` submenu: one command per size it offers, in the order it lists
/// them, for each of the three caches it sizes. They start past the range the `theme`
/// folder's own items occupy (see `ID_TRAY_THEME_CUSTOM_BASE`).
const ID_TRAY_IMAGE_CACHE_BASE: u16 = 1300;
/// The `Cache → Document` sizes: how much of what an engine drew is kept between hovers. The
/// pages are files under the temp folder rather than memory, which is what makes it one of the
/// two caches a size is measured in bytes of something on disk.
const ID_TRAY_DOCUMENT_CACHE_BASE: u16 = 1320;
/// The `Cache → Image (Disk)` sizes: how much of what the image converter developed is kept
/// between hovers. The other cache whose size is bytes on disk — pictures of its own, in a
/// folder beside the documents' pages rather than a share of them (see `document_cache`).
const ID_TRAY_IMAGE_DISK_CACHE_BASE: u16 = 1340;
/// The `Performance → Decode Budget` submenu: one command per ceiling it offers, in
/// the order it lists them. It sits in the slack between the `Cache` sizes and the
/// document scale's own range.
const ID_TRAY_DECODE_BUDGET_BASE: u16 = 1380;
/// The sizes the `Cache` submenu offers, in megabytes, largest first — `2 GB` at
/// the top and a cache that holds nothing at the bottom — and the whole range the
/// settings allow, so a size a hand-edited `config.ini` asks for that is not one of
/// these is shown with nothing checked rather than rounded to one of them.
const CACHE_SIZE_CHOICES_MB: [u32; 9] = [2048, 1024, 512, 256, 128, 64, 32, 16, 0];
/// The ceilings the `Decode Budget` submenu offers, in gigabytes, largest first. It is
/// what one hover may decode or read for rather than what is kept, so the range starts
/// far above any file someone meant to hover and ends at the smallest ceiling a large
/// picture still fits in. A value a hand-edited `config.ini` asks for that is not one of
/// these is shown with nothing marked rather than rounded to one of them.
const DECODE_BUDGET_CHOICES_GB: [f32; 6] = [16.0, 8.0, 4.0, 2.0, 1.0, 0.5];
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
/// `110%` is listed between `125%` and `100%` and carries the one id the font sizes
/// have left: the engine-idle ranges take 1083 up to 1096, and the `theme` folder's own
/// items begin at 1100.
const ID_TRAY_FONT_110: u16 = 1097;
/// The sizes the `Text Preview → Font Size` submenu offers, in the order it lists them —
/// largest first — with the id each size carries. A size a hand-edited `config.ini` asks
/// for that is not one of these is shown with nothing marked rather than rounded to the
/// nearest.
const FONT_SIZE_CHOICES: [(u32, u16); 12] = [
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
/// The `Engine → Microsoft Office TTL` submenu: one command per idle time it offers, in
/// the order it lists them. The IDs the app used before this ended at 1082 and the
/// `theme` folder's items start at 1100, so this range is the slack between the two.
const ID_TRAY_ENGINE_IDLE_BASE: u16 = 1083;
/// The `Engine → WebView2 TTL` submenu, the same shape as the Microsoft Office one and in
/// the range after it.
const ID_TRAY_WEBVIEW_IDLE_BASE: u16 = 1090;
/// The `Engine → LibreOffice TTL` submenu, the third of them. It sits in the slack the
/// backdrop halves leave rather than in the block the other two share — 1083 to 1096 is full,
/// one font size at 1097, and the `theme` folder's items begin at 1100 — so its range is the
/// widest run left between the image backdrop's four ids at 1023 and the config row at 1040.
const ID_TRAY_LIBREOFFICE_IDLE_BASE: u16 = 1027;
/// The idle times the three `… TTL` submenus offer, longest first — the
/// order the menus list them in, so an engine that is never let go is the topmost
/// item and one that is let go as soon as it has drawn a page is the bottom one. A
/// value a hand-edited `config.ini` asks for that is not one of these is shown with
/// nothing marked rather than rounded to the nearest.
const ENGINE_IDLE_CHOICES: [EngineIdle; 7] = [
    EngineIdle::Indefinite,
    EngineIdle::Seconds(3600),
    EngineIdle::Seconds(1800),
    EngineIdle::Seconds(600),
    EngineIdle::Seconds(300),
    EngineIdle::Seconds(60),
    EngineIdle::Seconds(0),
];
/// The `Engine → AFK Timer` submenu: one command per away time it offers, in the order it
/// lists them. Both it and the `Persistent` range below sit past every other range the app
/// hands out — the `Tick` range is the last of those and ends at 1504 — so a time is never
/// read as a tick and a toggle is never read as either.
const ID_TRAY_AFK_TIMER_BASE: u16 = 1505;
/// The `Persistent` toggle at the top of each `… TTL` submenu, one command apiece, in the
/// order those submenus are listed: `Microsoft Office TTL`, then `LibreOffice TTL`, then
/// `WebView2 TTL`.
const ID_TRAY_ENGINE_PERSISTENT_BASE: u16 = 1512;
/// The away times the `AFK Timer` submenu offers, in the order it lists them: an hour at the
/// top and a quarter of a minute at the bottom, with the one that bounds an engine by
/// default in the middle. There is no `Indefinitely` here — a time that never comes round is
/// what the `Persistent` toggle beside it is for — and a value a hand-edited `config.ini`
/// asks for that is not one of these is shown with nothing marked rather than rounded to the
/// nearest, the way every other menu of this shape reads one. See `app::afk` for what the
/// time is counted against.
const AFK_TIMER_CHOICES_SECS: [u64; 7] = [3600, 1800, 600, 300, 60, 30, 15];

/// The two command ranges one `… TTL` submenu hands out: the idle times it lists, and the
/// `Persistent` toggle above them. They are one value rather than two because they belong to
/// the same submenu and are always handed over together — a submenu is built for one engine,
/// and both of its ranges are that engine's.
struct EngineIdleIds {
    times: u16,
    persistent: u16,
}

/// Where the `theme` folder's own items start: one command ID each, in the order
/// the submenu listed them. The IDs the app uses end at the `Office Engine TTL`
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
                ID_TRAY_TRIGGER_DISABLE => set_trigger_key_mode(TriggerKeyMode::Disable),
                ID_TRAY_TRIGGER_ENABLE => set_trigger_key_mode(TriggerKeyMode::Enable),
                ID_TRAY_TRIGGER_ENABLED => toggle_trigger_key_enabled(),
                ID_TRAY_PRIORITIZE_KEYBOARD => toggle_prioritize_keyboard(),
                ID_TRAY_ENGINE_OFFICE_MS => set_office_engine(OfficeEngine::MicrosoftOffice),
                ID_TRAY_ENGINE_OFFICE_LIBRE => set_office_engine(OfficeEngine::LibreOffice),
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
                cmd if (ID_TRAY_AVOID_BASE..ID_TRAY_AVOID_BASE + AVOID_CHOICES.len() as u16)
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
                cmd if (ID_TRAY_TICK_BASE..ID_TRAY_TICK_BASE + TICK_CHOICES_MS.len() as u16)
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
                cmd if (ID_TRAY_CODEC_BASE..ID_TRAY_CODEC_BASE + CODEC_COMMANDS).contains(&cmd) => {
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
                ID_TRAY_TEXT_FULL_MODE => toggle_text_preview_full_mode(),
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
}

unsafe fn show_context_menu(hwnd: HWND) {
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

    let _ = AppendMenuW(
        menu,
        MF_STRING | MF_POPUP,
        text_menu.0 as usize,
        w!("Text Preview"),
    );

    let _ = AppendMenuW(menu, MF_SEPARATOR, 0, PCWSTR::null());

    // Add the "Timing" submenu: how long a hover waits before its preview opens, how
    // long the same file is held off after its preview was dismissed, how long the
    // pointer must be still before anything previews at all, and what the trigger key
    // does.
    let timing_menu = CreatePopupMenu().unwrap();

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

    // Whether the keyboard driving Explorer holds a parked pointer back instead of the file
    // under it previewing: the one row here that is a switch rather than a delay or a key,
    // and it starts off, where the pointer's own hover wins.
    let prioritize_keyboard = CONFIG
        .lock()
        .map(|c| c.prioritize_keyboard)
        .unwrap_or(false);

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

    // Add the Vector Scaling, Ebook Scaling, Document Scaling, Font Scaling and Design Scaling
    // submenus: how much of the display each kind of document is drawn over. They sit beside
    // the picture scale because they are the same question about other kinds of preview, and
    // each is a submenu of its own because the answers are not the same answers: a picture's
    // percentage is of its own size, a document's is of the display — and a document and a
    // page do not start at the same share of it either.
    //
    // One of them covers both halves of the `Document` kind, since a page the render engine
    // drew is a page like any other: what the setting answers is how much of the display one
    // is given, whichever engine drew it.
    let (ebook_scale, document_scale, font_scale, design_scale, vector_scale) = CONFIG
        .lock()
        .map(|c| {
            (
                c.ebook_scale,
                c.document_scale,
                c.font_scale,
                c.design_scale,
                c.vector_scale,
            )
        })
        .unwrap_or((
            DEFAULT_EBOOK_SCALE,
            DEFAULT_DOCUMENT_SCALE,
            DEFAULT_FONT_SCALE,
            DEFAULT_DESIGN_SCALE,
            DEFAULT_VECTOR_SCALE,
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
    let (image_background, vector_background, font_background, dds_background, design_background) =
        CONFIG
            .lock()
            .map(|c| {
                (
                    c.image_background,
                    c.vector_background,
                    c.font_background,
                    c.dds_background,
                    c.design_background,
                )
            })
            .unwrap_or((
                DEFAULT_IMAGE_BACKGROUND,
                DEFAULT_VECTOR_BACKGROUND,
                DEFAULT_FONT_BACKGROUND,
                DEFAULT_DDS_BACKGROUND,
                DEFAULT_DESIGN_BACKGROUND,
            ));
    let background_menu = CreatePopupMenu().unwrap();

    // Each half is handed its own default, since the four backdrops a half offers are not
    // the same four everywhere: a picture, a drawing and a document start at the squares, a
    // specimen and a texture start at a page — and a texture is offered only the two pages.
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
    // the level each setting stands at carries the default mark (see `VOLUME_CHOICES`). Below
    // them is the same submenu's other half of the question, which is a sound's alone: where in
    // a file it starts playing (see `AUDIO_SEEK_CHOICES`).
    let (video_volume, audio_volume, audio_seek) = CONFIG
        .lock()
        .map(|c| (c.video_volume, c.audio_volume, c.audio_seek))
        .unwrap_or((
            DEFAULT_VIDEO_VOLUME,
            DEFAULT_AUDIO_VOLUME,
            DEFAULT_AUDIO_SEEK,
        ));

    let volume_menu = CreatePopupMenu().unwrap();
    let levels_menu = |current: u32, base: u16, default: u32| -> HMENU {
        let levels = CreatePopupMenu().unwrap();

        for (index, level) in VOLUME_CHOICES.iter().enumerate() {
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

        levels
    };

    let video_levels = levels_menu(video_volume, ID_TRAY_VIDEO_VOLUME_BASE, DEFAULT_VIDEO_VOLUME);
    let audio_levels = levels_menu(audio_volume, ID_TRAY_AUDIO_VOLUME_BASE, DEFAULT_AUDIO_VOLUME);

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
    // as, which are files under the temp folder; and the pictures the image converter developed
    // a file into, which are files of their own beside those — each of them listed largest
    // first, with the size its own cache starts at marked, and each of them saying which of the
    // two kinds of storage it is.
    let (image_cache_mb, document_cache_mb, image_disk_cache_mb) = CONFIG
        .lock()
        .map(|c| (c.image_cache_mb, c.document_cache_mb, c.image_disk_cache_mb))
        .unwrap_or((
            DEFAULT_IMAGE_CACHE_MB,
            DEFAULT_DOCUMENT_CACHE_MB,
            DEFAULT_IMAGE_DISK_CACHE_MB,
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

fn toggle_startup() {
    // What is flipped is the registry, and what is read to decide which way to flip is the
    // registry too: an entry can be taken away from outside this app, and a toggle that
    // trusted the configuration would then need two clicks to put back — the first turning
    // off something already off. The configuration is the record of the choice, written to
    // agree with what was just done.
    let enable = !startup::is_startup_enabled();

    if enable {
        startup::enable_startup();
    } else {
        startup::disable_startup();
    }

    if let Ok(mut config) = CONFIG.lock() {
        config.run_at_startup = enable;
        config.save();
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

/// Whether the key is watched is a setting rather than a view of one, so the preview
/// on screen is rebuilt: switching the key off while it is held lets a preview
/// through, and switching it back on while it is held takes one away.
fn toggle_trigger_key_enabled() {
    if let Ok(mut config) = CONFIG.lock() {
        config.trigger_key_enabled = !config.trigger_key_enabled;
        config.save();
    }
    refresh_preview();
}

/// Whether the keyboard driving Explorer holds a parked pointer back is a setting rather
/// than a view of one: nothing on screen is rebuilt and nothing is taken down — the switch
/// is read by the hook on its next tick — so a preview that is up when it is thrown is left
/// where it is.
fn toggle_prioritize_keyboard() {
    if let Ok(mut config) = CONFIG.lock() {
        config.prioritize_keyboard = !config.prioritize_keyboard;
        config.save();
    }
}

/// What a picture is drawn over — and every preview that is not a document: a page,
/// a painted frame, a page Office rendered.
///
/// The backdrop is part of the frame a preview was composited into rather than of
/// the file it was drawn from, so the preview on screen is given its frame again
/// rather than left holding the one it has.
fn set_image_background(index: u16) {
    let Some(background) = background_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.image_background = background;
        config.save();
    }
    refresh_preview();
}

/// And the same again for a font specimen, which is drawn on a page of its own: the page's
/// colours are part of what the engine draws, so the preview on screen is rebuilt rather
/// than only composited again.
fn set_font_background(index: u16) {
    let Some(background) = background_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.font_background = background;
        config.save();
    }
    refresh_preview();
}

/// And for a texture, which is a picture like any other on this side of the answer: its
/// frame is composited by this app, so the preview on screen only needs compositing again.
fn set_dds_background(index: u16) {
    let Some(background) = dds_background_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.dds_background = background;
        config.save();
    }
    refresh_preview();
}

/// And for a design document, which is composited by this app the way a texture is: what
/// stands behind the picture the file keeps of the document is this side's to draw, so the
/// preview on screen only needs compositing again.
fn set_design_background(index: u16) {
    let Some(background) = background_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.design_background = background;
        config.save();
    }
    refresh_preview();
}

/// And for a vector drawing, which is drawn over a backdrop of its own: an SVG document is
/// drawn on a page the engine owns, so the page's colours are part of what it draws and the
/// preview on screen is rebuilt rather than only composited again; a metafile is replayed
/// by this side, so a change to it is composited again like a picture's.
fn set_vector_background(index: u16) {
    let Some(background) = background_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.vector_background = background;
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
///
/// A kind switched off also takes the engine behind it: what an engine is for is
/// previews of its own kind, and one kept warm for a kind nobody can be shown is a
/// process — and a licence, and a few hundred megabytes — held for nothing.
fn toggle_preview_type(kind: PreviewType) {
    if let Ok(mut config) = CONFIG.lock() {
        let enabled = kind.enabled_in(&config);
        kind.set_enabled_in(&mut config, !enabled);
        config.save();
    }

    if !kind.enabled() {
        // An engine only one kind of preview is ever started for goes with that
        // kind's gate. The Office processes are ended from here rather than through
        // the worker — a worker inside a call it cannot cut short would hold one of
        // them for good, and this thread may not wait on one — while the browser
        // needs nothing said to it at all: its own thread reads this gate and lets
        // the engine go.
        if let PreviewType::Document = kind {
            office_render::stop_engines();
        }
    }

    refresh_preview_types();
}

/// Put every setting back at what this build recommends, with the extension lists left
/// exactly as they are.
///
/// The reset itself is the configuration's own (`reset_to_recommended`), and everything
/// after it is the rest of the app being told what a reset means: it is every tray toggle
/// at once, so a setting whose truth lives somewhere other than the file has to be put
/// back on the machine as well, and what a setting was measuring has to be asked again.
/// The question comes first and names what would change, which is the same difference the
/// row is offered on.
fn reset_settings_from_tray() {
    let changes = CONFIG
        .lock()
        .map(|config| config.settings_apart_from_recommended())
        .unwrap_or_default();

    if !dialogs::confirm_reset_settings(&changes) {
        return;
    }

    let run_at_startup = {
        let Ok(mut config) = CONFIG.lock() else {
            return;
        };

        config.reset_to_recommended();
        config.save();
        config.run_at_startup
    };

    // The entry is the registry's and the configuration is the record of the choice, the
    // same way round as the toggle beside it: a reset that turns it back on writes the
    // entry, or the file and the machine would disagree about what starts this app.
    if run_at_startup != startup::is_startup_enabled() {
        if run_at_startup {
            startup::enable_startup();
        } else {
            startup::disable_startup();
        }
    }

    // What a reset can turn off as easily as on, and the one setting whose engines are
    // ended from here when it does — the same call its own toggle makes.
    if !PreviewType::Document.enabled() {
        office_render::stop_engines();
    }

    refresh_preview_types();
    refresh_preview();
}

/// Put every extension list back at the built-in one, with every other setting left alone.
///
/// It is the same question over the other half of the configuration. What a kind of
/// preview matches a file against is its list, so what is on screen is asked whether it
/// still measures — and nothing outside the file reads these, which is why this one has no
/// registry to write and no engine to let go.
fn reset_lists_from_tray() {
    let sections = CONFIG
        .lock()
        .map(|config| config.lists_apart_from_built_in())
        .unwrap_or_default();

    if !dialogs::confirm_reset_lists(&sections) {
        return;
    }

    if let Ok(mut config) = CONFIG.lock() {
        config.reset_extension_lists();
        config.save();
    }

    refresh_preview_types();
    refresh_preview();
}

/// One `Timing` submenu: an item per delay the setting offers, in the order the table
/// lists them, with the delay the setting is on checked and the one it starts at marked
/// as the default — which is read from the settings rather than written into the labels,
/// so a delay that becomes a default, or stops being one, moves the mark with it.
///
/// The three submenus differ only in what they select, so they are built here rather
/// than one by one: they list the same delays, and a delay added to the table is one
/// every one of them offers.
fn timing_delay_menu(base_id: u16, selected_ms: u64, default_ms: u64) -> HMENU {
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
fn default_label(label: &str, is_default: bool) -> String {
    if is_default {
        format!("{label} (Default)")
    } else {
        label.to_string()
    }
}

/// Append one item whose label is built rather than written into the code, which is what a
/// label carrying a default mark is: the text is encoded here and handed over, and the
/// buffer outlives the call it is handed to.
fn append_labeled_item(menu: HMENU, flags: MENU_ITEM_FLAGS, id: u16, label: &str) {
    let label: Vec<u16> = label.encode_utf16().chain(std::iter::once(0)).collect();

    let _ = unsafe { AppendMenuW(menu, flags, id as usize, PCWSTR(label.as_ptr())) };
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

    default_label(&label, megabytes == default_mb)
}

/// The label a `Decode Budget` item carries: the ceiling it stands for, in the unit
/// that reads best for it, with the one the app starts at marked.
fn decode_budget_label(gigabytes: f32) -> String {
    let label = if gigabytes < 1.0 {
        format!("{} MB", (gigabytes * 1024.0).round() as u32)
    } else {
        format!("{gigabytes} GB")
    };

    default_label(&label, gigabytes == DEFAULT_DECODE_BUDGET_GB)
}

/// The `Avoid` submenu: how far a preview is kept off the item it is about, with the
/// way the setting is on marked.
fn append_avoid_menu(parent: HMENU, label: PCWSTR, base: u16, avoid_mode: AvoidMode) {
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
fn avoid_label(mode: AvoidMode) -> String {
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
fn append_audio_seek_menu(parent: HMENU, seek: AudioSeek) {
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
fn audio_seek_label(seek: AudioSeek) -> String {
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
fn append_tick_menu(parent: HMENU, label: PCWSTR, base: u16, tick_ms: u64) {
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
fn append_background_menu(
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
fn background_label(background: TransparentBackground, default: TransparentBackground) -> String {
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
fn background_at(index: u16) -> Option<TransparentBackground> {
    BACKGROUND_CHOICES.get(index as usize).copied()
}

/// And the same for an item of the texture's half, which is the one half that offers two
/// backdrops of the four rather than all of them.
fn dds_background_at(index: u16) -> Option<TransparentBackground> {
    DDS_BACKGROUND_CHOICES.get(index as usize).copied()
}

/// One `… Scaling` submenu: the shares of the display a document is drawn at, with the
/// one the setting is on marked, and nothing marked for a share the menu does not
/// offer — which is what a hand-edited `config.ini` can ask for. `default` says which
/// of the shares this setting starts at, so the one it names is the one the label
/// marks.
fn append_document_scale_menu(
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
fn append_bitmap_scale_menu(
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
fn bitmap_scale_label(scale: PreviewScale, default: PreviewScale) -> String {
    let label = match scale {
        PreviewScale::Percent(percent) => format!("{percent}%"),
        _ => "Fit to Screen".to_string(),
    };

    default_label(&label, scale == default)
}

/// What a share of the display is called in the menu: the percentage itself, with the
/// one the setting starts at marked as the default. The whole room is not a percentage
/// of it, so it is named for what it is.
fn document_scale_label(scale: PreviewScale, default: PreviewScale) -> String {
    let label = match scale {
        PreviewScale::Percent(percent) => format!("{percent}%"),
        _ => "Fit to Screen".to_string(),
    };

    default_label(&label, scale == default)
}

/// The share of the display an item of a `… Scaling` submenu stands for, by the
/// position it was listed at. An id past the last choice the menu offered is one that
/// is not there.
fn document_scale_at(index: u16) -> Option<PreviewScale> {
    DOCUMENT_SCALE_CHOICES.get(index as usize).copied()
}

/// The `Engine → Select Engine → Office` submenu: which engine an Office document's page
/// is asked of, with the setting's own choice marked.
///
/// There are two engines to ask and both are listed. The row for one this machine has not got
/// is greyed out rather than left out: what it names cannot be started, so the app would fall
/// back to the other engine for as long as that is so — but the choice is the user's, it is
/// remembered where it is made, and the day the engine is installed it is the one that draws
/// (see `office_formats::page_engine`).
fn append_select_engine_menu(parent: HMENU) {
    let selected = CONFIG
        .lock()
        .map(|config| config.office_engine)
        .unwrap_or(DEFAULT_OFFICE_ENGINE);

    let select_engine_menu = unsafe { CreatePopupMenu().unwrap() };
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
fn append_afk_timer_menu(parent: HMENU) {
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
fn append_engine_idle_menu(
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
fn append_codecs_menu(menu: HMENU) {
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
fn open_codec_page(index: u16) {
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
fn engine_idle_label(idle: EngineIdle, default_seconds: u64) -> String {
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
fn engine_idle_at(index: u16) -> Option<EngineIdle> {
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
fn afk_timer_secs_at(index: u16) -> Option<u64> {
    AFK_TIMER_CHOICES_SECS.get(index as usize).copied()
}

/// How long the Office engines are kept after their families' last pages — for an engine
/// marked `Persistent`. One that is not is let go by the AFK timer instead, and this is not
/// consulted for it (see `app::afk`).
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

/// Which engine Office documents are asked of, from `Engine → Select Engine → Office`.
///
/// Nothing is rebuilt here, and nothing has to be: the choice is read live by the side that
/// asks an engine for a page and by the side that draws one (see `office_formats::page_engine`),
/// so the setting a click leaves behind is the one the next hover is answered by. A preview
/// that is already on screen belongs to the engine that drew it and is replaced the next time
/// a hover is raised — and the pointer has left the file to reach the tray by then.
///
/// What is ended here is this app's own Office applications, at the moment the render engine
/// becomes the one to ask: they were started for pages this choice now takes elsewhere, and
/// the one thing this menu is not for is a process kept warm for work it will not be given.
/// A user's own Word or Excel is not one of these and is never touched (see `office_render`).
fn set_office_engine(engine: OfficeEngine) {
    if let Ok(mut config) = CONFIG.lock() {
        config.office_engine = engine;
        config.save();
    }

    if engine == OfficeEngine::LibreOffice {
        office_render::stop_engines();
    }
}

/// How long the browser engine is kept after the last document it drew — for a browser that
/// is marked `Persistent`; one that is not is let go by the AFK timer instead, and this is
/// not consulted for it. Nothing is rebuilt here either: the engine reads the setting every
/// time it decides whether to let itself go, so a shorter time applies to the engine that is
/// already warm.
fn set_webview_idle(index: u16) {
    let Some(idle) = engine_idle_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.webview_idle = idle;
        config.save();
    }
}

/// How long the LibreOffice engine is kept after the last page it drew — for an engine marked
/// `Persistent`; one that is not is let go by the AFK timer instead, and this is not consulted
/// for it.
///
/// Nothing is rebuilt here either, and nothing has to be: the engine thread reads the setting
/// every second while it waits for documents, so a shorter time applies to the engine that is
/// already running, and `0 seconds` — the bottom of the list — lets go of one within the
/// second. An engine that is kept is a process this app holds and ends itself; a setting of
/// `indefinitely` keeps it for the rest of the run (see `libreoffice_render`).
fn set_libreoffice_idle(index: u16) {
    let Some(idle) = engine_idle_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.libreoffice_idle = idle;
        config.save();
    }
}

/// How long Explorer may be out of reach before an engine that is not marked `Persistent` is
/// let go.
///
/// Nothing is rebuilt here and nothing on screen changes for the same reason the idle times
/// rebuild nothing: an engine that is not persistent reads this on the look it already takes
/// — the Office worker twice a second, the engine thread once a second, the browser every
/// quarter of one while it is up — so a shorter time lets go of an engine that is already
/// warm, and a longer one keeps an engine the next look would have let go of.
fn set_afk_timer(index: u16) {
    let Some(seconds) = afk_timer_secs_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.afk_timer_seconds = seconds;
        config.save();
    }
}

/// Whether one of the three engines is kept whatever the user is doing, from the
/// `Persistent` toggle at the top of its TTL submenu.
///
/// The toggle is which submenu it heads rather than which engine it names, because the three
/// are listed in one order and the ids are handed out in it. Nothing is rebuilt here either:
/// both sides of the setting are read live by the engine that decides with them, so a toggle
/// turned on keeps the engine the next look would have let go of, and one turned off lets go
/// of an engine that is already up as soon as the AFK timer says it may.
fn toggle_engine_persistent(index: u16) {
    let Ok(mut config) = CONFIG.lock() else {
        return;
    };

    match index {
        0 => config.office_engine_persistent = !config.office_engine_persistent,
        1 => config.libreoffice_persistent = !config.libreoffice_persistent,
        2 => config.webview_persistent = !config.webview_persistent,
        _ => return,
    }

    config.save();
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

/// How much of what an engine drew may be kept, between hovers.
///
/// Unlike the image cache beside it this does not switch anything off: a page is drawn for the
/// hover that asks for it whatever the size, and a size of nothing means it is given up when
/// that hover ends. So nothing is rebuilt here either — the next hover answers for itself.
fn set_document_cache_mb(index: u16) {
    let Some(megabytes) = cache_size_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.document_cache_mb = sanitize_document_cache_mb(megabytes);
        config.save();
    }

    document_cache::trim_now();
}

/// How much of what the image converter developed may be kept, between hovers.
///
/// The same shape as the document cache beside it, and for the same reason: a picture is
/// developed for the hover that asks for it whatever the size, and a size of nothing means the
/// hover after it pays for the development again. Nothing is rebuilt here either — the pages are
/// read by the hover that wants one, so what a size does is bound what is left for it to read.
fn set_image_disk_cache_mb(index: u16) {
    let Some(megabytes) = cache_size_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.image_disk_cache_mb = sanitize_image_disk_cache_mb(megabytes);
        config.save();
    }

    document_cache::trim_image_now();
}

/// What one hover may decode or read for, in gigabytes.
///
/// Nothing is rebuilt here, and nothing already on screen changes: every reader asks
/// for the budget as it runs, so the next hover is answered under the new ceiling
/// whatever it is — a smaller one simply refuses more files than a larger one did.
fn set_decode_budget_gb(index: u16) {
    let Some(gigabytes) = DECODE_BUDGET_CHOICES_GB.get(index as usize).copied() else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.decode_budget_gb = sanitize_decode_budget_gb(gigabytes);
        config.save();
    }
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

fn set_video_volume(index: u16) {
    let Some(volume) = VOLUME_CHOICES.get(index as usize).copied() else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.video_volume = volume;
        config.save();
    }
}

/// A level of the `Volume → Audio` submenu, by the position it was listed at: the volume a
/// sound file is played at, which is the setting a card is drawn against as well as the one its
/// player is started with.
fn set_audio_volume(index: u16) {
    let Some(volume) = VOLUME_CHOICES.get(index as usize).copied() else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.audio_volume = volume;
        config.save();
    }
}

/// A way of starting a sound, by the position it was listed at: where in a file a hover drops
/// the needle, which is read as a player is started the way the volume beside it is. An id past
/// the last way the menu offered is one that is not there.
///
/// Nothing on screen is rebuilt: a sound already playing is left where it is, and what a click
/// here changes is where the *next* sound starts — the same bargain the volume beside it makes.
fn set_audio_seek(index: u16) {
    let Some(seek) = audio_seek_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.audio_seek = seek;
        config.save();
    }
}

/// The way an item of the `Volume → Audio Seek` submenu stands for, by the position it was
/// listed at. An id past the last way the menu offered is one that is not there.
fn audio_seek_at(index: u16) -> Option<AudioSeek> {
    AUDIO_SEEK_CHOICES.get(index as usize).copied()
}

fn set_follow_cursor(follow: bool) {
    if let Ok(mut config) = CONFIG.lock() {
        config.follow_cursor = follow;
        config.save();
    }
}

/// The way an item of the `Avoid` submenu stands for, by the position it was listed
/// at. An id past the last way the menu offered is one that is not there.
fn avoid_mode_at(index: u16) -> Option<AvoidMode> {
    AVOID_CHOICES.get(index as usize).copied()
}

/// How far a preview is kept off the item it is about.
///
/// Where that item's name is drawn is read with the hover, so the region is part of
/// the placement that was made when the preview was opened — and, like the position
/// setting beside it, this applies to the next hover rather than moving the preview
/// that is already up.
fn set_avoid_mode(index: u16) {
    let Some(mode) = avoid_mode_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.avoid_mode = mode;
        config.save();
    }
}

/// The tick an item of the `Tick` submenu stands for, by the position it was listed at.
/// An id past the last tick the menu offered is one that is not there.
fn tick_ms_at(index: u16) -> Option<u64> {
    TICK_CHOICES_MS.get(index as usize).copied()
}

/// How often the app looks at the pointer's world while Explorer has focus.
///
/// Nothing on screen changes and nothing is rebuilt: the preview that is up was placed
/// when it was opened, and the tick is what says when the next look happens — so a
/// slower tick is a preview that stays a moment longer after the pointer has left it,
/// and a faster one is an answer that arrives sooner, at the cost of more crossings into
/// Explorer (see `DEFAULT_TICK_MS`).
fn set_tick_ms(index: u16) {
    let Some(tick_ms) = tick_ms_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.tick_ms = sanitize_tick_ms(tick_ms);
        config.save();
    }
}

/// How large a picture is drawn, by the position the item was listed at.
///
/// The size a bitmap is drawn at is part of the placement that was made when the preview
/// was opened — the box is sized, and the frame is scaled into it — so, like the position
/// beside it, this applies to the next hover rather than resizing the preview that is up.
fn set_preview_scale(index: u16) {
    let Some(scale) = bitmap_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.preview_scale = scale;
        config.save();
    }
}

/// The same for a video, at the share `video_scale` names: the frame that stands in for
/// one is a bitmap like a picture, so the shares its submenu offers are the picture's
/// shares, and the setting written is the video's own.
fn set_video_scale(index: u16) {
    let Some(scale) = bitmap_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.video_scale = scale;
        config.save();
    }
}

/// And the same for an animated picture, at the share `animated_scale` names: the frames
/// an animation decodes into are bitmaps like a picture's, so this submenu offers the same
/// shares as the two beside it, and the setting written is the animation's own. Which
/// files are animated is not a setting at all: a GIF, a WebP or a PNG is one when the file
/// itself holds more than a single frame, and a still one keeps the picture scale.
fn set_animated_scale(index: u16) {
    let Some(scale) = bitmap_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.animated_scale = scale;
        config.save();
    }
}

/// How much of the display a PDF page — the `Ebook` kind — is drawn over, by the position
/// the item was listed at.
///
/// The size a document is drawn at is part of the placement that was made when the preview
/// was opened — the box is sized, and the document is drawn into it — so, like the position
/// and the picture scale beside it, this applies to the next hover rather than resizing the
/// preview that is already up.
fn set_ebook_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.ebook_scale = scale;
        config.save();
    }
}

/// How much of the display a page of the `Document` kind is shown over, by the position the
/// item was listed at — the one setting behind both halves of the kind: a page an Office
/// document's own application exported, and a page the render engine drew.
fn set_document_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.document_scale = scale;
        config.save();
    }
}

/// How much of the display a font specimen is drawn over, by the position the item was
/// listed at. The same rule as the three beside it: the next hover, not the one that is up.
fn set_font_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.font_scale = scale;
        config.save();
    }
}

/// How much of the display a design document is drawn over, by the position the item was
/// listed at. The same rule as the four beside it: what a document is previewed from is
/// the picture its own format keeps of the whole thing, so the share is of the display
/// rather than of the document, and it applies to the next hover rather than resizing a
/// preview that is already up.
fn set_design_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.design_scale = scale;
        config.save();
    }
}

/// How much of the display a vector drawing is replayed over, by the position the item was
/// listed at — the same rule as the documents beside it: the share is of the display, and
/// it applies to the next hover rather than resizing a preview that is already up.
fn set_vector_scale(index: u16) {
    let Some(scale) = document_scale_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.vector_scale = scale;
        config.save();
    }
}

/// The share of its own size an item of the `Image Scaling` or `Video Scaling` submenu
/// stands for, by the position it was listed at. An id past the last choice the menu
/// offered is one that is not there.
fn bitmap_scale_at(index: u16) -> Option<PreviewScale> {
    BITMAP_SCALE_CHOICES.get(index as usize).copied()
}

/// The delay an item of a `Timing` submenu stands for, by the position it was listed at.
/// The three submenus list the same delays, so one table answers for all of them — and
/// each submenu's own range is what says which setting the click was meant for. An id
/// past the last item is not one the menu offered.
fn timing_delay_at(index: u16) -> Option<u64> {
    TIMING_DELAY_CHOICES_MS.get(index as usize).copied()
}

/// How long the pointer must rest on a file before a preview is put up for it, by the
/// position its item was listed at.
fn set_hover_delay(index: u16) {
    let Some(delay_ms) = timing_delay_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.hover_delay_ms = delay_ms;
        config.save();
    }
}

/// How long the same file waits before a preview of it is put up again, by the position
/// its item was listed at.
fn set_same_file_rehover_delay(index: u16) {
    let Some(delay_ms) = timing_delay_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.same_file_rehover_delay_ms = delay_ms;
        config.save();
    }
}

/// How long the pointer must be still before a preview may open for anything, by the
/// position its item was listed at.
fn set_settling_delay(index: u16) {
    let Some(delay_ms) = timing_delay_at(index) else {
        return;
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.settling_delay_ms = delay_ms;
        config.save();
    }
}

/// A page, opened in the browser the user already has: the same call the release page is
/// opened with, and nothing is fetched or run here — what becomes of the page is the
/// browser's own business (see `updates::open_release_page`).
fn open_link(url: &str) {
    let wide_url: Vec<u16> = url.encode_utf16().chain(std::iter::once(0)).collect();

    unsafe {
        let _ = ShellExecuteW(
            HWND(std::ptr::null_mut()),
            w!("open"),
            PCWSTR(wide_url.as_ptr()),
            PCWSTR(std::ptr::null()),
            PCWSTR(std::ptr::null()),
            SW_SHOWNORMAL,
        );
    }
}

fn open_config_file() {
    if let Ok(config) = CONFIG.lock() {
        config.save();
    }

    if let Some(path) = crate::config::config::AppConfig::config_path() {
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
            ENGINE_IDLE_CHOICES
                .map(|idle| engine_idle_label(idle, DEFAULT_OFFICE_ENGINE_IDLE_SECS)),
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

    /// The halves of the `Background` submenu list the backdrops their settings hold, in the
    /// order the ids are handed out in, and each id resolves back to the backdrop its item was
    /// listed for — which is what makes a click select what it named. Every half marks the
    /// backdrop its own setting starts at, which is not the same one for every half, and the
    /// texture's half offers two of the four rather than all of them.
    #[test]
    fn every_offered_background_is_one_the_setting_keeps() {
        assert_eq!(
            BACKGROUND_CHOICES.map(|choice| background_label(choice, DEFAULT_IMAGE_BACKGROUND)),
            [
                "Transparent".to_string(),
                "Black".to_string(),
                "White".to_string(),
                "Checkerboard (Default)".to_string(),
            ]
        );

        // The mark is the setting's own answer: one backdrop per half carries it, and which
        // one it is is read from the constant that half's setting starts at.
        for (choices, default) in [
            (&BACKGROUND_CHOICES[..], DEFAULT_IMAGE_BACKGROUND),
            (&BACKGROUND_CHOICES[..], DEFAULT_VECTOR_BACKGROUND),
            (&BACKGROUND_CHOICES[..], DEFAULT_FONT_BACKGROUND),
            (&BACKGROUND_CHOICES[..], DEFAULT_DESIGN_BACKGROUND),
            (&DDS_BACKGROUND_CHOICES[..], DEFAULT_DDS_BACKGROUND),
        ] {
            let marked: Vec<TransparentBackground> = choices
                .iter()
                .copied()
                .filter(|choice| background_label(*choice, default).ends_with(" (Default)"))
                .collect();

            assert_eq!(
                marked,
                [default],
                "the backdrop a half starts at is the one it marks"
            );
        }

        for (index, background) in BACKGROUND_CHOICES.iter().enumerate() {
            assert_eq!(background_at(index as u16), Some(*background));
        }

        assert_eq!(
            background_at(BACKGROUND_CHOICES.len() as u16),
            None,
            "an id past the last item is not one the menu offered"
        );

        // The texture's half is a range and a table of its own, two backdrops wide.
        for (index, background) in DDS_BACKGROUND_CHOICES.iter().enumerate() {
            assert_eq!(dds_background_at(index as u16), Some(*background));
        }

        assert_eq!(
            dds_background_at(DDS_BACKGROUND_CHOICES.len() as u16),
            None,
            "a backdrop the texture's half does not offer is not one of its items"
        );
    }

    /// An item of one half of the submenu is never an item of another, whatever it was
    /// listed at — and never an id something else in the menu hands out either, which is
    /// the failure this range was moved for: a backdrop of a specimen and the text
    /// preview's `Full Mode` item were one id, and the backdrop was read first.
    #[test]
    fn the_five_halves_of_the_background_submenu_carry_different_ids() {
        // Each half is as wide as the choices it offers, which is two for the texture's half
        // and four for the rest: a range that were wider than the items in it would take an
        // id from the half beside it.
        let halves = [
            (
                ID_TRAY_IMAGE_BACKGROUND_BASE,
                BACKGROUND_CHOICES.len() as u16,
            ),
            (
                ID_TRAY_FONT_BACKGROUND_BASE,
                BACKGROUND_CHOICES.len() as u16,
            ),
            (
                ID_TRAY_DDS_BACKGROUND_BASE,
                DDS_BACKGROUND_CHOICES.len() as u16,
            ),
            (
                ID_TRAY_DESIGN_BACKGROUND_BASE,
                BACKGROUND_CHOICES.len() as u16,
            ),
            (
                ID_TRAY_VECTOR_BACKGROUND_BASE,
                BACKGROUND_CHOICES.len() as u16,
            ),
        ];

        for (index, half) in halves.iter().enumerate() {
            for other in &halves[index + 1..] {
                let ours = half.0..half.0 + half.1;
                let theirs = other.0..other.0 + other.1;

                assert!(
                    !ours.contains(&other.0) && !theirs.contains(&half.0),
                    "the ranges {ours:?} and {theirs:?} overlap"
                );
            }
        }

        // The ids around them, of the menus that grew up beside the `Background` one.
        for elsewhere in [
            ID_TRAY_TEXT_FULL_MODE,
            ID_TRAY_THEME_LIGHT,
            ID_TRAY_MARKDOWN_RENDERED,
            ID_TRAY_OPEN_CONFIG,
            ID_TRAY_SCALE_BASE,
            ID_TRAY_VIDEO_SCALE_BASE,
            ID_TRAY_TYPE_IMAGES,
            ID_TRAY_TRIGGER_ENABLED,
        ] {
            for half in halves {
                let ours = half.0..half.0 + half.1;

                assert!(
                    !ours.contains(&elsewhere),
                    "the range {ours:?} contains {elsewhere}, which is another item's id"
                );
            }
        }
    }

    /// The two halves of the `Volume` submenu carry a range apiece, and the levels they offer
    /// are the table's own: silence at the top, the whole of it at the bottom, and every level
    /// above the one before it.
    ///
    /// It is a test about ids for the reason the timing one below is — the two halves share a
    /// table, so a range that ran into the other would hand a click to the wrong setting — and
    /// about the table for the reason the font sizes' own test is: a level no item offers is a
    /// setting a user can reach only by editing the file.
    #[test]
    fn the_two_volume_submenus_carry_a_range_apiece() {
        let bases = [ID_TRAY_VIDEO_VOLUME_BASE, ID_TRAY_AUDIO_VOLUME_BASE];
        let width = VOLUME_CHOICES.len() as u16;

        assert_eq!(VOLUME_CHOICES[0], 0, "the topmost item is silence");
        assert_eq!(
            *VOLUME_CHOICES.last().expect("a last level"),
            100,
            "the bottom one is the whole of it"
        );
        assert!(
            VOLUME_CHOICES.windows(2).all(|pair| pair[0] < pair[1]),
            "every level is louder than the one above it: {VOLUME_CHOICES:?}"
        );
        for default in [DEFAULT_VIDEO_VOLUME, DEFAULT_AUDIO_VOLUME] {
            assert!(
                VOLUME_CHOICES.contains(&default),
                "{default}% is marked as a default and is no item of the menu"
            );
        }

        for (index, base) in bases.iter().enumerate() {
            for above in &bases[..index] {
                assert!(
                    above + width <= *base,
                    "the range at {base} overlaps the one at {above}"
                );
            }
        }

        // The gate for sounds is a command of its own, in the slack the font sizes leave
        // rather than in either range: a level is not a switch, and neither is a switch a
        // level — a collision here is a click that turns previews off where it meant to
        // change a volume.
        for base in bases {
            assert!(
                !(base..base + width).contains(&ID_TRAY_TYPE_AUDIO),
                "the sound gate's id is inside the range at {base}"
            );
        }

        // And the third item of the same submenu — where a sound starts — is a range of its
        // own as well: it shares the menu with both halves rather than a table, and a click on
        // one of its ways is never a click on a level of either (see the test below).
        let seek =
            ID_TRAY_AUDIO_SEEK_BASE..ID_TRAY_AUDIO_SEEK_BASE + AUDIO_SEEK_CHOICES.len() as u16;
        for base in bases {
            let range = base..base + width;
            assert!(
                !range.contains(&seek.start) && !seek.contains(&range.start),
                "the ranges {range:?} and {seek:?} overlap"
            );
        }
    }

    /// The `Volume → Audio Seek` submenu lists a way of starting a sound for every way the
    /// setting has, in the order the ids are handed out in, with the one the setting starts at
    /// marked as the default — and each id resolves back to the way its item was listed for,
    /// which is what makes a click start a sound where it named.
    #[test]
    fn every_offered_way_of_starting_a_sound_is_one_the_setting_keeps() {
        assert_eq!(
            AUDIO_SEEK_CHOICES.map(audio_seek_label),
            [
                "Remember (Default)".to_string(),
                "From the Start".to_string(),
                "From the Middle".to_string(),
                "Random".to_string()
            ]
        );

        let marked: Vec<AudioSeek> = AUDIO_SEEK_CHOICES
            .iter()
            .copied()
            .filter(|seek| audio_seek_label(*seek).ends_with(" (Default)"))
            .collect();

        assert_eq!(
            marked,
            [DEFAULT_AUDIO_SEEK],
            "the way the setting starts at is the way the menu marks"
        );

        for (index, seek) in AUDIO_SEEK_CHOICES.iter().enumerate() {
            assert_eq!(audio_seek_at(index as u16), Some(*seek));
        }

        assert_eq!(
            audio_seek_at(AUDIO_SEEK_CHOICES.len() as u16),
            None,
            "an id past the last item is not one the menu offered"
        );

        // The ways the submenu offers are the ways the setting holds, and each of them is
        // named: a way the menu has no words for is one a user cannot pick, and a way the
        // setting has that the menu does not list is one they cannot reach but by editing
        // `config.ini` — which is the arrangement every other value menu here keeps to.
        assert_eq!(
            AUDIO_SEEK_CHOICES.len(),
            4,
            "every way of starting a sound is offered: {AUDIO_SEEK_CHOICES:?}"
        );
        assert!(
            AUDIO_SEEK_CHOICES.contains(&DEFAULT_AUDIO_SEEK),
            "the way the setting starts at is one of the items"
        );
    }

    /// The three `Timing` submenus list the same delays, one range apiece, and a range
    /// that were wider than the items in it would take an id from the submenu below it —
    /// which is a click selecting a delay that was never listed, for a setting nobody
    /// asked for.
    #[test]
    fn the_three_timing_submenus_carry_a_range_apiece() {
        let bases = [
            ID_TRAY_DELAY_BASE,
            ID_TRAY_REHOVER_DELAY_BASE,
            ID_TRAY_SETTLING_DELAY_BASE,
        ];
        let width = TIMING_DELAY_CHOICES_MS.len() as u16;

        // The table is the steps themselves: no wait at all at the top, a whole second at
        // the bottom, and every delay larger than the one above it.
        assert_eq!(TIMING_DELAY_CHOICES_MS[0], 0, "the topmost item is no wait");
        assert_eq!(
            *TIMING_DELAY_CHOICES_MS.last().expect("a last delay"),
            1000,
            "the bottom one is a whole second"
        );
        assert!(
            TIMING_DELAY_CHOICES_MS
                .windows(2)
                .all(|pair| pair[0] < pair[1]),
            "every delay is larger than the one above it: {TIMING_DELAY_CHOICES_MS:?}"
        );

        // Every value a default mark is read from is one of the items: a default the menu
        // does not offer would leave the setting starting at nothing marked.
        for default_ms in [
            DEFAULT_HOVER_DELAY_MS,
            DEFAULT_SAME_FILE_REHOVER_DELAY_MS,
            DEFAULT_SETTLING_DELAY_MS,
        ] {
            assert!(
                TIMING_DELAY_CHOICES_MS.contains(&default_ms),
                "{default_ms} ms is marked as a default and is no item of the menu"
            );
        }

        for (index, base) in bases.iter().enumerate() {
            for above in &bases[..index] {
                assert!(
                    above + width <= *base,
                    "the range at {base} overlaps the one at {above}"
                );
            }
        }

        for (position, delay_ms) in TIMING_DELAY_CHOICES_MS.iter().enumerate() {
            assert_eq!(
                timing_delay_at(position as u16),
                Some(*delay_ms),
                "an item resolves back to the delay it was listed for"
            );
        }

        assert_eq!(
            timing_delay_at(width),
            None,
            "an id past the last item is not one the menu offered"
        );
    }

    /// The `Avoid` submenu lists every way the setting can be in, in the order the ids
    /// are handed out in, and each id resolves back to the way its item was listed for
    /// — which is what makes a click select what it named.
    #[test]
    fn every_offered_avoid_mode_is_one_the_setting_keeps() {
        assert_eq!(
            AVOID_CHOICES.map(avoid_label),
            [
                "Avoid Nothing".to_string(),
                "Avoid Filename (Default)".to_string(),
                "Avoid Filename Column".to_string(),
                "Avoid Details".to_string()
            ]
        );

        let marked: Vec<AvoidMode> = AVOID_CHOICES
            .iter()
            .copied()
            .filter(|mode| avoid_label(*mode).ends_with(" (Default)"))
            .collect();

        assert_eq!(
            marked,
            [DEFAULT_AVOID_MODE],
            "the way the setting starts at is the way the menu marks"
        );

        for (index, mode) in AVOID_CHOICES.iter().enumerate() {
            assert_eq!(avoid_mode_at(index as u16), Some(*mode));
        }

        assert_eq!(
            avoid_mode_at(AVOID_CHOICES.len() as u16),
            None,
            "an id past the last item is not one the menu offered"
        );
    }

    /// The five `… Scaling` submenus are one range each, and the `Avoid` items sit in
    /// the slack between them: a click on a share of the display is never read as a way
    /// of avoiding the item a preview is about, and the other way round.
    #[test]
    fn the_avoid_submenu_carries_ids_of_its_own() {
        let avoid = ID_TRAY_AVOID_BASE..ID_TRAY_AVOID_BASE + AVOID_CHOICES.len() as u16;
        let document_scales = [
            ID_TRAY_VECTOR_SCALE_BASE,
            ID_TRAY_EBOOK_SCALE_BASE,
            ID_TRAY_DOCUMENT_SCALE_BASE,
            ID_TRAY_FONT_SCALE_BASE,
            ID_TRAY_DESIGN_SCALE_BASE,
        ]
        .map(|base| base..base + DOCUMENT_SCALE_CHOICES.len() as u16);

        for range in &document_scales {
            assert!(
                !range.contains(&avoid.start) && !avoid.contains(&range.start),
                "the ranges {range:?} and {avoid:?} overlap"
            );
        }
        assert!(
            avoid.start > ID_TRAY_POSITION_BEST,
            "the avoid items are listed after the position choices"
        );

        // One range per submenu as well: a click on one scale is never read as a click on
        // the submenu beside it.
        for (index, range) in document_scales.iter().enumerate() {
            for other in document_scales.iter().skip(index + 1) {
                assert!(
                    !range.contains(&other.start) && !other.contains(&range.start),
                    "the ranges {range:?} and {other:?} overlap"
                );
            }
        }
    }

    /// Every `… Scaling` submenu offers the whole room a document can be given and then
    /// the shares of it, in that order, and each id resolves back to the share its item
    /// was listed for — which is what makes a click select what it named. What each
    /// submenu marks as the default is the share its own setting starts at.
    #[test]
    fn every_offered_document_scale_is_one_the_setting_keeps() {
        let drawing_default = DEFAULT_VECTOR_SCALE;

        assert_eq!(
            DOCUMENT_SCALE_CHOICES.map(|scale| document_scale_label(scale, drawing_default)),
            [
                "Fit to Screen (Default)".to_string(),
                "75%".to_string(),
                "50%".to_string(),
                "25%".to_string(),
                "10%".to_string(),
            ]
        );

        for default in [
            drawing_default,
            DEFAULT_EBOOK_SCALE,
            DEFAULT_DOCUMENT_SCALE,
            DEFAULT_FONT_SCALE,
        ] {
            assert_eq!(
                DOCUMENT_SCALE_CHOICES.map(|scale| document_scale_label(scale, default)),
                [
                    if default == PreviewScale::FitToScreen {
                        "Fit to Screen (Default)"
                    } else {
                        "Fit to Screen"
                    },
                    "75%",
                    if default == PreviewScale::Percent(50) {
                        "50% (Default)"
                    } else {
                        "50%"
                    },
                    "25%",
                    "10%",
                ]
                .map(String::from),
                "one share is marked the default at {default:?}"
            );
        }

        for (index, scale) in DOCUMENT_SCALE_CHOICES.iter().enumerate() {
            assert_eq!(document_scale_at(index as u16), Some(*scale));
        }

        assert_eq!(
            document_scale_at(DOCUMENT_SCALE_CHOICES.len() as u16),
            None,
            "an id past the last item is not one the menu offered"
        );
    }

    /// What the menu writes is what the file reads back: every share it offers is one
    /// the setting holds, so a choice made here is still the choice after a restart.
    #[test]
    fn every_offered_document_scale_round_trips_through_the_file() {
        for scale in DOCUMENT_SCALE_CHOICES {
            let written = scale.as_str();

            assert_eq!(
                PreviewScale::from_str(&written),
                Some(scale),
                "`{written}` read back"
            );
        }
    }

    /// A share of the display is never read as a cache size, and no cache size is read as
    /// another's: the three caches are a range each, and the shares added beside them are ranges
    /// of their own.
    #[test]
    fn the_document_scale_ranges_are_not_another_submenus_range() {
        let caches = [
            ID_TRAY_IMAGE_CACHE_BASE..ID_TRAY_IMAGE_CACHE_BASE + CACHE_SIZE_CHOICES_MB.len() as u16,
            ID_TRAY_DOCUMENT_CACHE_BASE
                ..ID_TRAY_DOCUMENT_CACHE_BASE + CACHE_SIZE_CHOICES_MB.len() as u16,
            ID_TRAY_IMAGE_DISK_CACHE_BASE
                ..ID_TRAY_IMAGE_DISK_CACHE_BASE + CACHE_SIZE_CHOICES_MB.len() as u16,
            ID_TRAY_THEME_CUSTOM_BASE..ID_TRAY_IMAGE_CACHE_BASE,
        ];

        for (index, cache) in caches.iter().enumerate() {
            for other in &caches[index + 1..] {
                assert!(
                    !cache.contains(&other.start) && !other.contains(&cache.start),
                    "the ranges {cache:?} and {other:?} overlap"
                );
            }
        }

        for base in [
            ID_TRAY_VECTOR_SCALE_BASE,
            ID_TRAY_EBOOK_SCALE_BASE,
            ID_TRAY_DOCUMENT_SCALE_BASE,
            ID_TRAY_FONT_SCALE_BASE,
            ID_TRAY_DESIGN_SCALE_BASE,
        ] {
            let scales = base..base + DOCUMENT_SCALE_CHOICES.len() as u16;

            for other in &caches {
                assert!(
                    !other.contains(&base) && !scales.contains(&other.start),
                    "the ranges {scales:?} and {other:?} overlap"
                );
            }
        }
    }

    /// The `Image Scaling`, `Video Scaling` and `Animated Scaling` submenus are one
    /// range each, and none of them reaches into another or into the display shares the
    /// document scales beside them hand out: a click on a share of a bitmap is never read
    /// as a click on another setting's share. Each lists every share the setting can be
    /// asked for, in order, and every id resolves back to the share its item was listed
    /// for — which is what makes a click select what it named.
    #[test]
    fn the_bitmap_scaling_submenus_carry_ids_of_their_own() {
        let bitmap_bases = [
            ID_TRAY_SCALE_BASE,
            ID_TRAY_VIDEO_SCALE_BASE,
            ID_TRAY_ANIMATED_SCALE_BASE,
        ];
        let ranges: Vec<(u16, u16)> = bitmap_bases
            .iter()
            .map(|base| (*base, base + BITMAP_SCALE_CHOICES.len() as u16))
            .collect();
        let font_scales = (
            ID_TRAY_FONT_SCALE_BASE,
            ID_TRAY_FONT_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
        );
        let design_scales = (
            ID_TRAY_DESIGN_SCALE_BASE,
            ID_TRAY_DESIGN_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
        );
        let ebook_scales = (
            ID_TRAY_EBOOK_SCALE_BASE,
            ID_TRAY_EBOOK_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
        );
        let document_scales = (
            ID_TRAY_DOCUMENT_SCALE_BASE,
            ID_TRAY_DOCUMENT_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
        );
        let vector_scales = (
            ID_TRAY_VECTOR_SCALE_BASE,
            ID_TRAY_VECTOR_SCALE_BASE + DOCUMENT_SCALE_CHOICES.len() as u16,
        );

        let overlaps = |ours: (u16, u16), theirs: (u16, u16)| {
            (ours.0 < theirs.1 && theirs.0 < ours.1).then_some((ours, theirs))
        };

        for (index, range) in ranges.iter().enumerate() {
            for other in ranges.iter().skip(index + 1) {
                assert_eq!(
                    overlaps(*range, *other),
                    None,
                    "the ranges {range:?} and {other:?} overlap"
                );
            }

            for document in [
                font_scales,
                design_scales,
                ebook_scales,
                document_scales,
                vector_scales,
            ] {
                assert_eq!(
                    overlaps(*range, document),
                    None,
                    "the range {range:?} and the display shares {document:?} overlap"
                );
            }
        }

        for base in bitmap_bases {
            for (index, scale) in BITMAP_SCALE_CHOICES.iter().enumerate() {
                assert_eq!(bitmap_scale_at(index as u16), Some(*scale));
            }

            assert_eq!(
                bitmap_scale_at(BITMAP_SCALE_CHOICES.len() as u16),
                None,
                "an id past the last item of the range at {base} is not one it offered"
            );
        }
    }

    /// The `Image Scaling` and `Video Scaling` submenus offer the shares a bitmap can be
    /// drawn at — the share of its own size, rather than the share of the display the
    /// document scales beside them are — in one order and with one set of labels: what a
    /// share is called does not depend on which of the two is asking, and exactly one
    /// label — the share each setting starts at — reads as the default. Every share an
    /// item can pick is one the setting keeps, so a choice made here is still the choice
    /// after a restart.
    #[test]
    fn every_offered_bitmap_scale_is_one_the_setting_keeps() {
        assert_eq!(
            BITMAP_SCALE_CHOICES.map(|scale| bitmap_scale_label(scale, DEFAULT_PREVIEW_SCALE)),
            [
                "Fit to Screen".to_string(),
                "400%".to_string(),
                "300%".to_string(),
                "200%".to_string(),
                "150%".to_string(),
                "100% (Default)".to_string(),
                "50%".to_string(),
                "25%".to_string(),
            ]
        );

        for default in [
            DEFAULT_PREVIEW_SCALE,
            DEFAULT_VIDEO_SCALE,
            DEFAULT_ANIMATED_SCALE,
        ] {
            let marked: Vec<String> = BITMAP_SCALE_CHOICES
                .iter()
                .map(|scale| bitmap_scale_label(*scale, default))
                .filter(|label| label.ends_with(" (Default)"))
                .collect();

            assert_eq!(
                marked,
                [bitmap_scale_label(default, default)],
                "one share is the default at {default:?}"
            );
        }

        for scale in BITMAP_SCALE_CHOICES {
            let written = scale.as_str();

            assert_eq!(
                PreviewScale::from_str(&written),
                Some(scale),
                "`{written}` read back"
            );
        }
    }

    /// The font sizes are listed largest first, `110%` between the `125%` and `100%` it
    /// sits between, and every size carries an id of its own — so a click selects the
    /// size its item named.
    #[test]
    fn the_font_sizes_are_listed_largest_first() {
        assert!(
            FONT_SIZE_CHOICES
                .windows(2)
                .all(|pair| pair[0].0 > pair[1].0),
            "every size is smaller than the one above it: {FONT_SIZE_CHOICES:?}"
        );
        assert!(
            FONT_SIZE_CHOICES.contains(&(110, ID_TRAY_FONT_110)),
            "110% is offered"
        );

        for (index, (_, id)) in FONT_SIZE_CHOICES.iter().enumerate() {
            let above = &FONT_SIZE_CHOICES[..index];
            assert!(
                !above.iter().any(|(_, earlier)| earlier == id),
                "an id names one size and one only: {id}"
            );
        }
    }

    /// Every item a person can pick is one the setting can hold, so what the menu
    /// writes is what the menu reads back and marks on the next open.
    #[test]
    fn every_offered_idle_time_is_one_the_setting_keeps() {
        for idle in ENGINE_IDLE_CHOICES {
            assert_eq!(EngineIdle::from_str(&idle.as_str()), Some(idle), "{idle:?}");
        }

        assert_eq!(
            engine_idle_at(ENGINE_IDLE_CHOICES.len() as u16),
            None,
            "an id past the last item is not one the menu offered"
        );
        assert_eq!(engine_idle_at(0), Some(EngineIdle::Indefinite));
    }
}

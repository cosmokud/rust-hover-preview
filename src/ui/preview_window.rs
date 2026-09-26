use crate::app::engine_processes;
use crate::config::config::{
    frame_bytes_within_budget, image_decode_limits, read_within_budget, sanitize_image_cache_mb,
    sanitize_spinner_delay_ms, sanitize_webp_playback_fps, AudioSeek, MarkdownMode, OfficeEngine,
    PreviewScale, PreviewType, TextTheme, TransparentBackground, DEFAULT_ANIMATED_SCALE_PERCENT,
    DEFAULT_AUDIO_SEEK, DEFAULT_DDS_BACKGROUND, DEFAULT_DESIGN_BACKGROUND, DEFAULT_DESIGN_SCALE,
    DEFAULT_DOCUMENT_SCALE, DEFAULT_EBOOK_SCALE, DEFAULT_FONT_BACKGROUND, DEFAULT_FONT_SCALE,
    DEFAULT_IMAGE_BACKGROUND, DEFAULT_IMAGE_CACHE_MB, DEFAULT_PREVIEW_SCALE_PERCENT,
    DEFAULT_SPINNER_DELAY_MS, DEFAULT_TEXT_FONT_SCALE_PERCENT,
    DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS, DEFAULT_VECTOR_BACKGROUND, DEFAULT_VECTOR_SCALE,
    DEFAULT_VIDEO_SCALE_PERCENT, DEFAULT_WEBP_PLAYBACK_FPS,
};
use crate::engines::calibre_render;
use crate::engines::imagemagick_render;
use crate::engines::libreoffice_render;
use crate::engines::office_render;
use crate::engines::peazip_render;
use crate::engines::webview_preview;
use crate::formats::archive_formats;
use crate::formats::audio_formats;
use crate::formats::calibre_formats;
use crate::formats::codecs;
use crate::formats::design_formats;
use crate::formats::ebook_formats;
use crate::formats::font_formats;
use crate::formats::libre_formats;
use crate::formats::magick_formats;
use crate::formats::native_formats;
use crate::formats::office_formats;
use crate::formats::peazip_formats;
use crate::formats::vector_formats;
use crate::formats::video_formats;
use crate::readers::audio_seek;
use crate::readers::audio_track::{self, Player, Probed};
use crate::readers::comic_preview;
use crate::readers::dds_image;
use crate::readers::eps_image;
use crate::readers::font_preview;
use crate::readers::metafile_image;
use crate::readers::office_preview;
use crate::readers::pdf_preview;
use crate::readers::project_image;
use crate::readers::psd_image;
use crate::readers::svg_preview;
use crate::readers::tone_map;
use crate::readers::video_player;
use crate::readers::webp_image;
use crate::readers::wic_image;
use crate::shell::cloud_files;
use crate::shell::wheel_input;
use crate::text::archive_preview::{self, ArchivePreviewOptions};
use crate::text::audio_preview::{self, AudioPreviewOptions, Card};
use crate::text::text_preview::{self, TextPreviewOptions};
use crate::{CONFIG, RUNNING};
use gif::DecodeOptions;
use image::{AnimationDecoder, GenericImageView};
use once_cell::sync::Lazy;
use std::cell::RefCell;
use std::collections::{HashMap, VecDeque};
use std::fs::File;
use std::io::BufReader;
use std::os::windows::io::AsRawHandle;
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Output, Stdio};
use std::ptr;
use std::sync::atomic::{AtomicBool, AtomicIsize, AtomicU32, AtomicU64, Ordering};
use std::sync::mpsc::{channel, Receiver, RecvTimeoutError, Sender};
use std::sync::{Arc, Condvar, Mutex, MutexGuard};
use std::time::{Duration, Instant, SystemTime};
use windows::core::{w, PCWSTR, PWSTR};
use windows::Win32::Foundation::{
    CloseHandle, GlobalFree, COLORREF, HANDLE, HWND, LPARAM, LRESULT, POINT, RECT, SIZE,
    WAIT_TIMEOUT, WPARAM,
};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, ClientToScreen, CreateCompatibleDC, CreateDIBSection, DeleteDC, DeleteObject,
    EndPaint, GetMonitorInfoW, MonitorFromPoint, SelectObject, AC_SRC_ALPHA, AC_SRC_OVER,
    BITMAPINFO, BITMAPINFOHEADER, BI_RGB, BLENDFUNCTION, DIB_RGB_COLORS, HBITMAP, HDC, HGDIOBJ,
    MONITORINFO, MONITOR_DEFAULTTONEAREST, PAINTSTRUCT,
};
use windows::Win32::System::DataExchange::{
    CloseClipboard, EmptyClipboard, OpenClipboard, SetClipboardData,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Memory::{GlobalAlloc, GlobalLock, GlobalUnlock, GMEM_MOVEABLE};
use windows::Win32::System::Ole::CF_UNICODETEXT;
use windows::Win32::System::Threading::{
    CreateEventW, OpenProcess, QueryFullProcessImageNameW, SetEvent, TerminateProcess,
    WaitForSingleObject, PROCESS_NAME_WIN32, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_TERMINATE,
};
use windows::Win32::UI::HiDpi::{GetDpiForMonitor, MDT_EFFECTIVE_DPI};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetAsyncKeyState, ReleaseCapture, SetCapture, VK_C, VK_CONTROL,
};
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CreatePopupMenu, CreateWindowExW, DefWindowProcW, DestroyMenu, DispatchMessageW,
    EnumWindows, GetCursorPos, GetSystemMetrics, GetWindow, GetWindowLongPtrW, GetWindowRect,
    GetWindowThreadProcessId, IsWindow, IsWindowVisible, LoadCursorW, MoveWindow, PeekMessageW,
    RegisterClassExW, SetForegroundWindow, SetWindowLongPtrW, SetWindowPos, ShowWindow,
    ShowWindowAsync, SystemParametersInfoW, TrackPopupMenu, TranslateMessage, UpdateLayeredWindow,
    CS_HREDRAW, CS_VREDRAW, GWL_EXSTYLE, GW_OWNER, HWND_TOPMOST, IDC_ARROW, MF_STRING, MSG,
    PBT_APMRESUMEAUTOMATIC, PBT_APMRESUMESUSPEND, PBT_APMSTANDBY, PBT_APMSUSPEND, PM_REMOVE,
    SM_CXVIRTUALSCREEN, SM_CYVIRTUALSCREEN, SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN, SPI_GETWORKAREA,
    SWP_FRAMECHANGED, SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOSIZE, SWP_SHOWWINDOW, SW_HIDE,
    SW_SHOWNOACTIVATE, SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS, TPM_LEFTALIGN, TPM_NONOTIFY,
    TPM_RETURNCMD, TPM_TOPALIGN, ULW_ALPHA, WM_DISPLAYCHANGE, WM_DPICHANGED, WM_LBUTTONDOWN,
    WM_LBUTTONUP, WM_MOUSEMOVE, WM_POWERBROADCAST, WM_RBUTTONUP, WNDCLASSEXW, WS_EX_LAYERED,
    WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_POPUP,
};

const PREVIEW_CLASS: PCWSTR = w!("RustHoverPreviewWindow");

/// Budget for the frames one animation keeps decoded. An animation that fits is
/// decoded once and loops from memory; a larger one plays through a sliding
/// window and its decoder starts over when the animation wraps.
const ANIMATION_RETAINED_BYTES: usize = 128 * 1024 * 1024;
/// Frames the streaming decoder may keep queued ahead of playback before it
/// waits, so decoding can never run away from what is on screen.
const ANIMATION_QUEUE_FRAMES: usize = 6;
/// How much already-played footage piles up behind the playhead before it is
/// released. Releasing in blocks keeps the sliding window from moving per frame.
const ANIMATION_RELEASE_BYTES: usize = 16 * 1024 * 1024;
const MIN_ANIMATION_FRAME_DELAY_MS: u32 = 33;
/// Frames an animation is given before it is handed over.
///
/// Two, rather than a startup buffer: the frame the preview opens on is the first
/// one, and holding a buffer before showing anything is a delay the user reads as
/// the preview not appearing. The decoder runs on ahead of playback once it has
/// been handed over, so it is the queue depth that keeps playback fed, not this —
/// and every frame still reaches the screen, because the handover only decides
/// when the preview opens, not how much of the animation is played.
const ANIMATION_STARTUP_FRAMES: usize = 2;
const STREAMING_SPINNER_MAX_MS: u64 = 1500;
/// How long the preview thread waits when there is nothing on screen: no preview
/// to animate, nothing to repaint, and no pointer region left to keep in step.
/// The wait is the preview channel rather than a sleep, so a hover is answered
/// the moment it arrives and this interval is only a ceiling on how long a window
/// message — a resume, a display change — waits to be noticed.
const IDLE_WAIT_MS: u64 = 500;
const VIDEO_GEOMETRY_CACHE_MAX_ENTRIES: usize = 512;
// Expected executable name of the playback process spawned below, used to
// verify a recorded PID still belongs to that process before killing it.
const VIDEO_PROCESS_IMAGE_NAME: &str = "ffplay.exe";

// Message passing for thread communication
pub static PREVIEW_SENDER: Lazy<Mutex<Option<Sender<PreviewMessage>>>> =
    Lazy::new(|| Mutex::new(None));

// Use AtomicIsize for the HWND pointer (thread-safe)
static PREVIEW_HWND: AtomicIsize = AtomicIsize::new(0);

/// The clock the preview loop's liveness is read on, and when that loop was last
/// seen running.
///
/// One thread owns the answer and another reads it: the Explorer hook is what
/// decides about a preview loop that has stopped answering — the engines such a loop
/// is holding warm are ended from there instead (see `PREVIEW_STALL_MS` in
/// `explorer_hook`) — and it cannot ask the loop itself, because a question put to a
/// loop that has stopped is the one question that would not be answered. So the loop
/// notes its own tick and the hook reads how long ago the last one was. One relaxed
/// store per tick and one relaxed load per look is the whole cost of it.
static PREVIEW_CLOCK: Lazy<Instant> = Lazy::new(Instant::now);
static PREVIEW_ALIVE_MS: AtomicU64 = AtomicU64::new(0);

/// Note that the preview loop has run a tick — whether the tick did anything or not,
/// since a loop waiting on the channel for a hover is a loop that is working (see
/// `preview_stall_ms`).
fn note_preview_alive() {
    PREVIEW_ALIVE_MS.store(
        PREVIEW_CLOCK.elapsed().as_millis() as u64,
        Ordering::Relaxed,
    );
}

/// How long the preview loop has been quiet, in milliseconds: the age of its last
/// tick, growing for as long as the loop is inside work that has not come back.
///
/// A loop that has not ticked since the app started answers with the age of the app,
/// which is the same answer a loop that is gone is owed — nothing here waits on the
/// loop or looks for it; it is only ever a number read.
pub fn preview_stall_ms() -> u64 {
    let alive = PREVIEW_ALIVE_MS.load(Ordering::Relaxed);
    (PREVIEW_CLOCK.elapsed().as_millis() as u64).saturating_sub(alive)
}

/// The number of times the window has been taken down, and the lock that makes a
/// take-down and the reveal it races one step rather than two threads writing the
/// window's visibility at once.
///
/// `hide_preview` moves the count there and then, on the Explorer hook's thread,
/// while the frame that puts it up is installed here when a load lands. A
/// load that lands in the moment after the pointer left would otherwise put the
/// preview back up for a file nobody is on any more, to be taken down again by the
/// next tick: the preview that blinks. A load carries the count it was started
/// under, and a hide moves it, so a reveal whose count has moved is refused. It is
/// a comparison and not a wait — nothing here ever holds a preview back from going
/// up, and the window that comes down with the count is posted rather than sent, so
/// a loop that is busy is never a thread the hide waits on (see `hide_preview`).
static HIDDEN_EPOCH: Mutex<u64> = Mutex::new(0);

/// The hide count a load starting now is under.
fn hidden_epoch() -> u64 {
    match HIDDEN_EPOCH.lock() {
        Ok(epoch) => *epoch,
        Err(poisoned) => *poisoned.into_inner(),
    }
}

/// Whether a load is still the one the pointer asked for, `hidden` being the guard
/// the caller is holding the answer under (see `HIDDEN_EPOCH`): a hide that has run
/// since the load started is the pointer having left it, and the frame it lands
/// with is not to be shown.
fn hover_still_wanted(hidden: &Option<MutexGuard<'static, u64>>, pl: &PendingLoad) -> bool {
    hidden
        .as_ref()
        .map(|epoch| **epoch == pl.hide_epoch)
        .unwrap_or(true)
}

/// The item the hover on screen is about, as the box the view draws it in: the
/// hook's own answer about the pointer, published with every look at the item under
/// it and withdrawn with the window.
///
/// The count above covers a load that lands after a hide has been *sent*. It cannot
/// cover the moment before one, and that is where the flash the eye catches lives:
/// the hook decides a hover on a tick of its own while the frame lands on a tick of
/// this loop's, and a pointer can cross a whole row of the list in between — a file
/// previewed, and taken down again by the hook's very next look. So the reveal is
/// asked one more question, of the pointer itself: the box says which item the hover
/// was resolved from, a pointer outside it is a hover that has moved on, and the
/// frame that lands for one is dropped rather than revealed — the hook's next look
/// answers for whatever the pointer is on instead. The same box is what tells the
/// hook a pointer that has crossed to another item has moved at all, which is not a
/// distance to measure (see `pointer_item_holds`).
///
/// A box that cannot be told — the view reported no item, the hover is the
/// keyboard's, nothing at all is on screen — is no constraint, and the reveal is
/// governed by the count alone.
static HOVER_POINTER_BOX: Mutex<Option<(i32, i32, i32, i32)>> = Mutex::new(None);

/// Note the item the pointer is on, as the box the view draws it in. Published by the
/// Explorer hook wherever it reads the item under the pointer, which is what keeps it
/// the item the pointer is on *now* rather than the one the last preview was for (see
/// `HOVER_POINTER_BOX`).
pub fn publish_pointer_item_box(bounds: (i32, i32, i32, i32)) {
    if let Ok(mut published) = HOVER_POINTER_BOX.lock() {
        *published = Some(bounds);
    }
}

/// Withdraw it: what is on screen is no longer a pointer's hover.
fn clear_pointer_item_box() {
    if let Ok(mut published) = HOVER_POINTER_BOX.lock() {
        *published = None;
    }
}

/// Whether a box on screen holds a point. Half-open, so a point on the box's right or
/// bottom edge is outside it — which is how a window hit-test reads a rectangle, and
/// what the dismissal of a mouse preview is decided by.
fn box_holds(x: i32, y: i32, region: (i32, i32, i32, i32)) -> bool {
    let (left, top, right, bottom) = region;

    x >= left && x < right && y >= top && y < bottom
}

/// Whether a point is still on the item the hover on screen is about — the one
/// question a preview is revealed against, and the one the hook reads a move off (see
/// `HOVER_POINTER_BOX`). A point outside a published box is a pointer that has left
/// the file, and no box at all is nothing to hold a reveal back with.
pub fn pointer_item_holds(x: i32, y: i32) -> bool {
    let Ok(published) = HOVER_POINTER_BOX.lock() else {
        return true;
    };

    (*published)
        .map(|region| box_holds(x, y, region))
        .unwrap_or(true)
}

/// The box the item under the pointer is drawn in, as the last look at it published, or
/// nothing when no look has answered with one (see `HOVER_POINTER_BOX`).
///
/// The hook asks for it where one look has to be told from another: the box a hover's
/// preview was resolved from is the item that preview is about, so a look that answered
/// nothing but left that same box under the pointer is a read that failed, while a look
/// that found another item publishes another box — an item with no preview of its own
/// being an item all the same (see `read_failure_is_the_same_item`).
pub fn pointer_item_box() -> Option<(i32, i32, i32, i32)> {
    HOVER_POINTER_BOX
        .lock()
        .ok()
        .and_then(|published| *published)
}

/// Whether the pointer is on that item this moment, read from the cursor: the
/// reveal's own question. A pointer that cannot be read is not a pointer that has
/// left, so an answer that could not be had holds nothing back either.
fn pointer_on_the_hovered_item() -> bool {
    cursor_position()
        .map(|cursor| pointer_item_holds(cursor.x, cursor.y))
        .unwrap_or(true)
}

/// Lines one wheel notch moves a text preview. Three is the step a text editor
/// takes, and it keeps a screenful to a few notches.
const TEXT_SCROLL_LINES_PER_NOTCH: i64 = 3;
const WHEEL_DELTA: i32 = 120;

/// How far behind the point a preview was opened from the region reaches, in
/// logical pixels. The pointer travels forwards from there, so this is only there
/// to keep the pixel under a hand at rest inside the region.
const TEXT_SCROLL_ANCHOR_SLACK_PIXELS: f32 = 1.0;
/// How far above and below the row the pointer is on the journey to a preview may
/// wander before it is out of it. The journey is made across a row of the list, so
/// the band is the hand's, not the preview's.
const TEXT_SCROLL_CORRIDOR_SLACK_PIXELS: f32 = 12.0;

/// How far either side of the scrollbar's column a press still counts as a press
/// on the bar, in logical pixels. It is deliberately small: the bar is thin, so a
/// hand aiming at it needs some slack, but everything further left is text, and a
/// press in the text is the start of a selection rather than a scroll.
const TEXT_SCROLL_BAR_PRESS_SLACK_PIXELS: f32 = 8.0;

/// A distance written in logical pixels — the pixels of a display at 100% — in the
/// pixels of the display it is drawn on.
///
/// Every margin this app's placement is written around is a distance under a hand
/// rather than a count of pixels, so all of them are scaled this way: a standoff that
/// is comfortable at 100% is a sliver at 200%, and a room worth taking at 100% is a
/// room a preview is squeezed into at 200%. The text, the scrollbar and the grace
/// below are scaled for the same reason.
fn logical_px(dpi: u32, logical_pixels: f32) -> i32 {
    (logical_pixels * dpi as f32 / 96.0).round() as i32
}

/// The grace `text_scroll_far_edge_grace_pixels` asks for past the far edge of a
/// text preview, at the display the preview is on.
///
/// That edge is the one the pointer arrives at last, and the one it can overshoot:
/// crossing the gap to reach the preview is a movement towards it, so the hand is
/// still moving when it gets there, and what usually waits at the end of the
/// journey is the scrollbar — a thin target sitting at the very edge of the frame.
/// A few pixels past it would otherwise take the preview down with it, which is
/// what this is for. It goes on the far side whichever side that is: the preview to
/// the right of the cursor is the common case, and then it is the right edge.
fn far_edge_grace(dpi: u32, configured_pixels: f32) -> i32 {
    logical_px(dpi, configured_pixels)
}

/// The grace as `config.ini` has it, so a hand-edited distance is used as written
/// and a config that names none keeps the default.
fn configured_far_edge_grace_pixels() -> f32 {
    CONFIG
        .lock()
        .map(|config| config.text_scroll_far_edge_grace_pixels)
        .unwrap_or(DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS)
}

/// A region on screen: left, top, right, bottom.
type ScreenRegion = (i32, i32, i32, i32);

/// The regions that keep a preview on screen, in screen coordinates, or `None`
/// when the preview on screen is not one that holds the pointer: the journey to a
/// text preview and the preview itself, or the box a waiting spinner occupies.
static POINTER_HOLD_REGIONS: Lazy<Mutex<Option<Vec<ScreenRegion>>>> =
    Lazy::new(|| Mutex::new(None));

/// Where the preview that is on screen was opened from: the cursor that hovered
/// the file, or the middle of the focused item. The hold region is built from
/// this point and the preview's box, so the whole path between them is inside it.
static TEXT_SCROLL_ANCHOR: Lazy<Mutex<Option<(i32, i32)>>> = Lazy::new(|| Mutex::new(None));

/// Whether the preview on screen is a text preview with more lines than it can
/// show, which is the condition for everything above.
static TEXT_PREVIEW_SCROLLABLE: AtomicBool = AtomicBool::new(false);

/// Whether a pointer is using the preview on screen: any text preview in full
/// mode, which the pointer can rest on to select from. Published beside the region
/// so the check for it stays an atomic read.
static TEXT_PREVIEW_HOLDING: AtomicBool = AtomicBool::new(false);

/// Whether what is on screen is a wait rather than a preview: the spinner a
/// document's page is being rendered behind.
///
/// The pointer is held by this the way it is held by a text preview, and for a
/// reason of its own: there is nothing under the spinner to hand the pointer back
/// to, and a hover that is dismissed while its page is on the way loses the page
/// it was waiting for — the render finishes, but the hover it was for is gone.
///
/// What it holds the pointer *through* is the item the wait is for and not the box
/// the spinner occupies: the spinner is placed a pixel off the pointer and follows
/// it, so a box of its own is one the pointer can never leave — which would be a
/// preview no fast hand could close on its way to somewhere else (see
/// `preview_pointer_hold`).
static WAITING_PREVIEW_HOLDING: AtomicBool = AtomicBool::new(false);

/// The regions that keep a preview alive: the journey to it, and the preview.
///
/// A preview is placed beside what it belongs to rather than over it, so the
/// pointer has to travel to reach it — across a gap, sometimes against the side
/// the placement chose. Joining the two means that journey never leaves the
/// regions, however the preview ended up placed relative to the cursor. The
/// journey is the first of the two regions and the preview is the second.
///
/// The journey reaches the preview's *nearest* point and stops there. Reaching
/// across to its far corner instead — which is what a region built from the two
/// corners of both rectangles does — covers every file in the list beside the
/// preview, from the row the pointer is on down to the preview's bottom: a text
/// preview is as tall as the display allows, and inside the region the item under
/// the pointer is not resolved at all, so those files stop previewing for as long
/// as the preview is up.
fn text_scroll_hold_regions(
    preview: ScreenRegion,
    anchor: (i32, i32),
    far_edge_grace: i32,
    dpi: u32,
) -> [ScreenRegion; 2] {
    let (left, top, right, bottom) = preview;

    // The preview, with the margin a hand that overshot the edge it was travelling
    // towards needs — that edge is the far one, and what usually waits just past it
    // is the scrollbar.
    let preview_region = if anchor.0 < left {
        (left, top, right + far_edge_grace, bottom)
    } else if anchor.0 >= right {
        (left - far_edge_grace, top, right, bottom)
    } else {
        (left, top, right, bottom)
    };

    // The journey: from the point the preview was opened from to the nearest point
    // of the preview, and no further. It is as tall as the journey is and not as
    // tall as the preview is, so a pointer beside a preview crosses a row of the
    // list and nothing else.
    let near_x = anchor.0.clamp(left, right);
    let near_y = anchor.1.clamp(top, bottom);

    let anchor_slack = logical_px(dpi, TEXT_SCROLL_ANCHOR_SLACK_PIXELS);
    let corridor_slack = logical_px(dpi, TEXT_SCROLL_CORRIDOR_SLACK_PIXELS);

    let corridor = (
        anchor.0.min(near_x) - anchor_slack,
        anchor.1.min(near_y) - corridor_slack,
        anchor.0.max(near_x) + anchor_slack,
        anchor.1.max(near_y) + corridor_slack,
    );

    [corridor, preview_region]
}

// Track the ffplay video window HWND for cursor-over-preview detection
static VIDEO_HWND: AtomicIsize = AtomicIsize::new(0);
// Track the ffplay process ID to re-find the window if needed
static VIDEO_PID: AtomicU32 = AtomicU32::new(0);
// Guard to ensure we only run a single style-monitor thread.
static NOACTIVATE_MONITOR_STARTED: AtomicBool = AtomicBool::new(false);
// Flag set when the system resumes from sleep, so the main loop can reset state.
static RESUME_FROM_SLEEP: AtomicBool = AtomicBool::new(false);
// Flag set when the display under the preview changed — a monitor's scale, or the
// desktop rearranged — a frame the window proc can only discard. What was on screen
// is put back by the loop, which is the side that holds the hover it came from.
static DISPLAY_RESET: AtomicBool = AtomicBool::new(false);

static CURRENT_MEDIA: Lazy<Mutex<Option<MediaData>>> = Lazy::new(|| Mutex::new(None));
/// What a probed geometry is only valid for: the file and the version of it that
/// was probed, so a video replaced in place is probed again rather than cropped and
/// sized by the answer about the file it used to be — the same rule every other
/// held picture in this module follows.
#[derive(Clone, PartialEq, Eq, Hash)]
struct VideoGeometryKey {
    path: PathBuf,
    version: FileVersion,
}

#[derive(Clone, Copy)]
enum ProbedGeometry {
    /// The shape the probe read, and the crop the detector settled on.
    Measured(VideoGeometry),
    /// The answer that there is nothing to measure: a file neither FFmpeg nor the media
    /// engine will open. It is an answer like any other and is held like one, so a file
    /// that cannot be measured is not probed again on every hover — and so the hover
    /// that is waiting for a probe can be told that the probe is done (see `video_box`).
    Unmeasurable,
}

static VIDEO_GEOMETRY_CACHE: Lazy<Mutex<HashMap<VideoGeometryKey, ProbedGeometry>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// The geometry cache's own guard, a poisoned lock included.
///
/// What is behind it is a map of facts — a shape, a crop, the answer that there is neither —
/// and there is no invariant a panic could have left half applied, so a lock a panicked probe
/// poisoned is one to read anyway rather than one that reads as an empty cache for the rest
/// of the run: an empty cache is `video_probe_due` true on every hover, which is a probe
/// started, and waited on, for every file that is hovered (see `hidden_epoch` for the same
/// reading of a lock that holds a fact).
fn video_geometry_cache() -> MutexGuard<'static, HashMap<VideoGeometryKey, ProbedGeometry>> {
    match VIDEO_GEOMETRY_CACHE.lock() {
        Ok(cache) => cache,
        Err(poisoned) => poisoned.into_inner(),
    }
}

#[derive(Clone)]
pub enum PreviewMessage {
    /// A preview of the file the pointer hovers, opened from the cursor it was
    /// hovered at.
    ///
    /// The region that comes with it is what the `Avoid` setting measured off the
    /// hovered item — its name, and the columns beside it at `Avoid Details` — which
    /// the placement is kept off at either way of avoiding. See `avoiding_text`. It is
    /// `None` when the setting is off, when the view reported
    /// no text for the item, or when the walk found no item at the cursor.
    Show(PathBuf, i32, i32, Option<ScreenRegion>),
    /// A preview of the focused item, whose box comes with it, and the region the
    /// `Avoid` setting keeps it off — the item's own text, or the name alone, or the
    /// column the name sits in, depending on the way the setting is on. That region is
    /// where its preview is placed from as well as what it is kept off, the way a
    /// hovered item's is, for an item that draws its text as a row of its view: the
    /// last field says whether it does, read off the item's own text by the hook (see
    /// `explorer_hook::ItemText`). See `compute_keyboard_layout`.
    ShowKeyboard(PathBuf, i32, i32, i32, i32, Option<ScreenRegion>, bool),
    Hide,
    Refresh,
    /// A preview type was switched on or off. Only a preview whose own kind is
    /// now off is rebuilt, and it is rebuilt from the hover it came from, so it
    /// goes away on the spot rather than at the next pointer move.
    RefreshTypes,
    /// The render tier is done with a document: a page is waiting in the cache,
    /// or there is no page. The generation is the hover that asked for it, so a
    /// render landing after the pointer has moved on is ignored — the page is
    /// still cached for the next hover either way.
    OfficeRenderReady {
        path: PathBuf,
        generation: u64,
        ok: bool,
    },
    /// A video's probe is done: the geometry is waiting in the cache, or the answer is
    /// that there is none. The generation is the hover that was waiting on it, so a probe
    /// landing after the pointer has moved on is ignored — the answer is held for the next
    /// hover either way (see `video_probe`).
    VideoProbed {
        path: PathBuf,
        generation: u64,
    },
    /// A measure that reads a file is done: the box it answers with is held, or the answer is
    /// that the reader has none for the file — which is not a wait that can be answered, so it
    /// is the one answer a wait comes down on (see `measured_off_the_tick`).
    ///
    /// What was waiting on it is the hover that is on screen, and that is the whole of what
    /// this answer has to be matched to: a box is measured per version of the file, and a hover
    /// of a file whose version has changed is measured again rather than answered with the box
    /// of the version before it.
    MeasureProbed {
        path: PathBuf,
        size: Option<(u32, u32)>,
    },
    /// The ImageMagick engine is done with a file: the picture it developed is in hand, or
    /// there is none — a file it cannot read is remembered as one it will not draw. The
    /// generation is the hover that was waiting on it, so a conversion landing after the
    /// pointer has moved on is ignored; the picture itself is held for the hover that asked
    /// either way (see `magick_render_is_due`).
    MagickReady {
        path: PathBuf,
        generation: u64,
        ok: bool,
    },
    /// The PeaZip engine is done with a file: the archive's table of contents is in hand, read
    /// into the listing cache, or there is none — a file it cannot open is remembered as one it
    /// will not list. The generation is the hover that was waiting on it, so a listing landing
    /// after the pointer has moved on is ignored; the listing itself is held under the file's own
    /// key either way, so the next hover of it is a read rather than a launch (see
    /// `peazip_render_is_due`).
    PeazipReady {
        path: PathBuf,
        generation: u64,
        ok: bool,
    },
}

/// Represents different types of media we can display
enum MediaType {
    StaticImage,
    /// A `.dds` texture, which is a still picture decoded by this app and drawn by this
    /// window like any other — a kind of its own for the one thing about it that differs:
    /// what a preview is drawn over. A texture's alpha channel is as often a mask, a
    /// height or a channel a tool never filled in as it is transparency, so its backdrop
    /// is the tray's own setting rather than a picture's (see the `Background` submenu and
    /// `dds_image`).
    Dds,
    /// An SVG document the engine draws in a window of its own. This side holds no
    /// frame for one: the kind is what the preview loop reads to hand the hover over,
    /// and the media it comes in arrives with nothing in it. It is a kind of its own
    /// rather than a static image because a document is composited over a backdrop of
    /// its own (see the tray's `Background` submenu) — a backdrop the engine is given
    /// rather than one this app composes.
    EngineSvg,
    /// A font the engine draws a specimen of in a window of its own, the same shape as a
    /// document: this side holds no frame for one, and what it answers about a font — the
    /// lines the file's own character map covers, and the face a collection is drawn by —
    /// comes from `font_preview`. It is a kind of its own for the reason above it, and
    /// because what it is drawn over is a page of its own rather than a document's.
    EngineFont,
    AnimatedGif,
    AnimatedApng,
    AnimatedWebP,
    /// A video played by `ffplay`, whose own window is the preview while it is up.
    Video,
    /// A video the media engine Windows has decodes and this window draws — the kind of
    /// video preview a machine without FFmpeg gets (see `video_player`).
    ///
    /// It is a kind of its own rather than a difference inside `Video` because the two are
    /// drawn by different things: the player's window is the preview for one, and the
    /// layered window this app owns is the preview for the other, which is the one
    /// question `render_layered_preview_at` asks of a frame.
    NativeVideo,
    /// A sound: a card of what the file holds, painted into this window's own frame like an
    /// archive's page, with the sound itself played by one of the two engines behind it (see
    /// `audio_preview`). Nothing of the player is drawn — the card is the whole of what is on
    /// screen — which is why a sound is a kind of its own here rather than a video with no
    /// frames in it.
    Audio,
    Pdf,
    Text,
    Archive,
    /// An archive an installed PeaZip listed for this app — a cabinet file, an iso, a disk image,
    /// a Linux package, a single-stream `.gz` or `.zst` — drawn as the same page an archive this
    /// app read itself is drawn as, because that is what it is: a list of what the file holds,
    /// painted by this window into a frame of its own.
    ///
    /// It is a kind of its own for the gate alone, the way the picture an image converter develops
    /// is: the switch over these is not the switch for archives, so a user who wants their isos
    /// left alone is not asking for their zips to be left alone. See `peazip_formats` and
    /// `peazip_render`.
    Peazip,
    Office,
    /// A design document, previewed from the picture its own format keeps of the whole
    /// thing: the merged image at the end of a Photoshop file, or the flattened
    /// document a project container holds beside its layers. It is a kind of its own
    /// for the reason the texture above is — the gate over it is not the gate over
    /// pictures, so a user who wants none of them has a switch that is not the switch
    /// for images — while what draws one is this window, into a frame composed like
    /// any other.
    Design,
    /// A vector drawing: Windows' metafiles, and the preview an encapsulated PostScript
    /// file carries. It is a kind of its own — drawn into this window's frame like a
    /// picture, but drawn by the drawing layer rather than decoded, and composited over a
    /// backdrop of its own — and what makes it worth a kind of its own is the size: the
    /// records are replayed at whatever box the preview is shown at, so a drawing is
    /// sharp at any size the display has.
    Vector,
    /// A document drawn by a render engine rather than read here — CorelDRAW above all:
    /// what comes back is a page, which is sharp at whatever size the preview is shown at,
    /// and what such a file keeps of itself is a thumbnail this app does not show. See
    /// `libre_formats` and `libreoffice_render`.
    Libre,
    /// A picture an installed ImageMagick developed for this app — a camera raw above all:
    /// what comes back is a PNG, which is decoded and drawn like the picture it is, over the
    /// backdrop a picture is drawn over and at the share of its own size a picture is drawn
    /// at. It is the picture kind's second half, and the switch over it is the switch for
    /// pictures; see `magick_formats` and `imagemagick_render`.
    Magick,
    /// A page an installed Calibre converted a book into — a Kindle or Mobipocket file, an EPub,
    /// a FictionBook, a scanned book — drawn exactly as a PDF page is: a frame of this app's own,
    /// made from the first page of the PDF the engine wrote that says anything about the book, at
    /// the share of the display a book is drawn over. It is the book kind's second half, and the
    /// switch over it is the switch for books; see `calibre_formats` and `calibre_render`.
    ///
    /// It is a kind of its own rather than `Pdf` because the two are gated apart: the switch a user
    /// throws over a book the app drew itself is not the switch over one an engine had to convert,
    /// and only one of the two costs a conversion to show.
    Calibre,
    /// The first plate of a comic book, read out of the container it is filed in — a `.cbz`, a
    /// `.cbr` or a `.cbc`: a plate is decoded the way a picture is, and then drawn the way a page
    /// is, at the share of the display a book is drawn over. It is the book kind's *other* half,
    /// and what it has in common with the two above is the whole of why it is here: what a hover on
    /// a book shows is a page, whichever reader drew one. See `ebook_formats` and `comic_preview`,
    /// and note that what draws it is this window rather than an engine — there is nothing to wait
    /// for and nothing to keep.
    Comic,
    Loading,
}

impl MediaType {
    /// The kind of preview this media is, as the tray's gates name them.
    fn kind(&self) -> Option<PreviewType> {
        match self {
            Self::StaticImage | Self::AnimatedGif | Self::AnimatedApng | Self::AnimatedWebP => {
                Some(PreviewType::Images)
            }
            // A texture is a picture as far as the gates go: the list a `.dds` is in is the
            // image list, and the switch for pictures is the switch for it.
            Self::Dds => Some(PreviewType::Images),
            Self::EngineSvg => Some(PreviewType::Vector),
            Self::EngineFont => Some(PreviewType::Fonts),
            Self::Video | Self::NativeVideo => Some(PreviewType::Videos),
            // A sound is its own kind in the tray as well as here: the switch over it is the
            // switch over the sound list, not the video's.
            Self::Audio => Some(PreviewType::Audio),
            Self::Text => Some(PreviewType::Text),
            Self::Pdf => Some(PreviewType::Ebook),
            Self::Archive => Some(PreviewType::Archives),
            // And an archive an engine listed: the page is an archive's page in every way that
            // matters — a frame of this app's own, painted from a listing — and the switch over
            // it is the archive switch, because what a user turns off is archives.
            Self::Peazip => Some(PreviewType::Peazip),
            Self::Office => Some(PreviewType::Document),
            // A design document is drawn into a frame like any picture, and the switch
            // over it is its own: the picture *is* what the file keeps of the document,
            // but a user who wants none of them is not asking for pictures to be off.
            Self::Design => Some(PreviewType::Design),
            // The same for a document an engine drew, at the gate over the document kind.
            Self::Libre => Some(PreviewType::Libre),
            // And for a picture one developed: it is a picture in every way that matters —
            // a frame of this app's own, drawn like any other — and the switch over it is the
            // switch for pictures, because what a user turns off is pictures.
            Self::Magick => Some(PreviewType::Magick),
            // And a book an engine converted: a page like any other, at the gate over books,
            // because what a user turns off is books.
            Self::Calibre => Some(PreviewType::Calibre),
            // And a comic, which is the same kind of preview as a PDF — a page of a book — drawn by
            // a reader of this app's own rather than by an engine, so the switch over it is the
            // switch for books and nothing else.
            Self::Comic => Some(PreviewType::Ebook),
            Self::Vector => Some(PreviewType::Vector),
            Self::Loading => None,
        }
    }

    /// Whether this is the spinner standing in for a preview that is not ready.
    fn is_loading(&self) -> bool {
        matches!(self, Self::Loading)
    }

    /// Whether this is a document the engine draws rather than a frame this app holds.
    /// The preview loop asks before it reaches for a frame, because there is none.
    fn is_engine(&self) -> bool {
        matches!(self, Self::EngineSvg | Self::EngineFont)
    }

    /// Whether this is a video the media engine decodes and this window draws, which is
    /// what the tick asks before it pulls a frame.
    fn is_native_video(&self) -> bool {
        matches!(self, Self::NativeVideo)
    }

    /// Whether this is a sound, whose card is the one painted preview that changes while it is
    /// on screen: the clock and the bar under it are drawn from a player that is running.
    fn is_audio(&self) -> bool {
        matches!(self, Self::Audio)
    }

    /// Whether this preview's appearance is painted into its own frame rather
    /// than recomposited from shared pixels, which is what decides whether a
    /// theme switch means rebuilding it.
    fn is_painted(&self) -> bool {
        matches!(self, Self::Text | Self::Archive | Self::Peazip | Self::Audio)
    }
}

/// A single frame of image data
#[derive(Clone)]
struct ImageFrame {
    pixels: Vec<u8>,
    width: u32,
    height: u32,
    delay_ms: u32, // Delay before next frame (for animations)
    /// Whether every pixel of the frame has an alpha of 255, which is what lets a repaint
    /// copy the frame instead of blending it pixel by pixel.
    ///
    /// A frame that is opaque everywhere is the surface it is drawn on already, whatever the
    /// backdrop behind it is, so composing it is a copy of its bytes — and at the size of a
    /// display that is the difference between a few gigabytes a second and a few hundred
    /// megabytes, on every frame of a video or an animation (see
    /// `compose_preview_pixels_into`).
    ///
    /// It is asked where the pixels are made, off the thread that draws them, and never
    /// guessed: a frame whose producer has not asked keeps the blend, which is the same
    /// picture by a longer road. `false` is therefore always safe and `true` never is — the
    /// one producer that says so without asking is the video path, which forces the alpha of
    /// every pixel it writes (see `copy_locked`).
    opaque: bool,
}

impl ImageFrame {
    /// A frame of `pixels`, with its opacity asked of the pixels themselves.
    ///
    /// The question is a pass over the frame, which is why it is asked here rather than by
    /// whoever composes it: a pass per repaint would cost what the copy it enables saves.
    fn new(pixels: Vec<u8>, width: u32, height: u32, delay_ms: u32) -> Self {
        let opaque = pixels_are_opaque(&pixels);

        Self {
            pixels,
            width,
            height,
            delay_ms,
            opaque,
        }
    }

    /// Replace the pixels of a frame, with the opacity asked of them again.
    fn set_pixels(&mut self, pixels: Vec<u8>) {
        self.opaque = pixels_are_opaque(&pixels);
        self.pixels = pixels;
    }
}

/// Whether every pixel of a frame is opaque, which is what a repaint copies rather than
/// blends (see `ImageFrame`).
fn pixels_are_opaque(pixels: &[u8]) -> bool {
    pixels
        .as_chunks::<4>()
        .0
        .iter()
        .all(|pixel| pixel[3] == 255)
}

/// Frames an animated preview streams while it plays. The decoder appends to the
/// queue and the player drains it; `released` records that the player gave back
/// frames it already showed, after which the animation can no longer loop from
/// memory and the decoder has to start the file over.
///
/// The two are read and written under this one lock, and that is what makes the
/// arrangement sound: a pass that ends with nothing given back leaves every frame
/// in the player's hands and the file needs no decoder again (`decoded`), while a
/// pass that ends after frames were given back is a pass to run again. Settling
/// both under the same lock is what keeps a release and the end of a decode from
/// passing each other, which is a player left holding the last frame of an
/// animation whose decoder has already gone.
struct StreamedFrames {
    queue: VecDeque<ImageFrame>,
    released: bool,
    /// Whether a pass was decoded whole with nothing given back, so the player
    /// holds the file entire and loops it from memory. Nothing is released after
    /// this: a frame dropped out of a file the decoder is done with could never be
    /// decoded again.
    decoded: bool,
}

/// Media data that can be either static or animated
struct MediaData {
    frames: Vec<ImageFrame>,
    /// Shared frame queue for streaming decode (animated formats append here)
    shared_frames: Option<Arc<Mutex<StreamedFrames>>>,
    /// Signal from the background thread that all frames have been decoded
    all_frames_loaded: Option<Arc<AtomicBool>>,
    current_frame: usize,
    last_frame_time: Instant,
    media_type: MediaType,
    /// Cancellation token for background decode work.
    stream_cancel: Option<Arc<AtomicBool>>,
    // For video playback using ffplay
    video_process: Option<Child>,
    loading_start: Option<Instant>,
    /// Where a text preview is scrolled to, when it scrolls at all.
    text_state: Option<TextPreviewState>,
}

/// What a text preview on screen keeps so that it can be worked with: where it is
/// scrolled to, what is selected in it, and the numbers a press is tested against.
///
/// A text preview in full mode always has one of these, whether or not the
/// document is longer than the frame — a selection needs somewhere to live even
/// when there is nothing to scroll. Without full mode a text preview keeps none of
/// it: nothing scrolls, nothing is selected, and a pointer over it dismisses it the
/// way any other preview is dismissed.
struct TextPreviewState {
    path: PathBuf,
    options: TextPreviewOptions,
    dpi: u32,
    width: u32,
    height: u32,
    /// Document line the frame starts at, and how many it shows.
    first_line: usize,
    visible_lines: usize,
    /// Lines the preview can reach, which is also the range its scrollbar is
    /// drawn and dragged in.
    scrollable_lines: usize,
    /// The bar drawn in the frame, kept so a drag can be tested against it.
    scrollbar: Option<text_preview::ScrollBar>,
    /// Whether the pointer is currently dragging the thumb.
    dragging: bool,
    /// The painted lines, which is what a press is turned back into a place in the
    /// text against, and what a selection is copied from.
    lines: Vec<text_preview::FrameLine>,
    /// What is selected in the frame on screen, and whether a drag is extending it.
    selection: Option<text_preview::Selection>,
    selecting: bool,
}

impl TextPreviewState {
    fn max_first_line(&self) -> usize {
        self.scrollable_lines.saturating_sub(self.visible_lines)
    }

    /// The line a scroll of `lines` from here lands on, kept inside the document.
    fn scrolled_by(&self, lines: i64) -> usize {
        (self.first_line as i64 + lines).clamp(0, self.max_first_line() as i64) as usize
    }

    fn can_scroll(&self) -> bool {
        self.max_first_line() > 0
    }
}

#[derive(Clone, Copy)]
struct VideoCrop {
    width: u32,
    height: u32,
    x: u32,
    y: u32,
}

#[derive(Clone, Copy)]
struct VideoGeometry {
    width: u32,
    height: u32,
    crop: Option<VideoCrop>,
}

impl MediaData {
    /// The frame on screen, which a document the engine draws does not have.
    fn current_frame(&self) -> Option<&ImageFrame> {
        self.frames.get(self.current_frame)
    }

    fn current_pixels(&self) -> &[u8] {
        self.current_frame()
            .map(|frame| frame.pixels.as_slice())
            .unwrap_or(&[])
    }

    /// Repaint a sound's card with the clock as it stands, answering whether the frame on
    /// screen changed.
    ///
    /// A painted preview is drawn once and held, so the two things about a card that move are
    /// drawn by asking for the page again — the same arrangement a text preview's scrolling has
    /// (see `repaint_text_preview`), and what keeps the card's own layout in one place: the box
    /// it was painted in is the box it is painted in again, and only the clock, the bar under
    /// it and the scroll of a name the card has no room for differ.
    fn refresh_audio_card(
        &mut self,
        path: &Path,
        elapsed: Option<f64>,
        duration: Option<f64>,
        dpi: u32,
        name_offset: i32,
    ) -> bool {
        let Some(frame) = self.frames.first() else {
            return false;
        };
        let (width, height) = (frame.width, frame.height);

        let Some(card) = audio_card(path, elapsed, duration, name_offset) else {
            return false;
        };
        let Some((pixels, width, height)) =
            audio_preview::render(&card, width, height, dpi, current_audio_options())
        else {
            return false;
        };

        self.frames[0] = ImageFrame::new(pixels, width, height, 0);
        true
    }

    fn current_width(&self) -> u32 {
        self.current_frame().map(|frame| frame.width).unwrap_or(0)
    }

    fn current_height(&self) -> u32 {
        self.current_frame().map(|frame| frame.height).unwrap_or(0)
    }

    /// Whether the frame on screen is opaque everywhere, which is what lets a repaint copy
    /// it rather than blend it (see `ImageFrame::opaque`).
    fn current_frame_is_opaque(&self) -> bool {
        self.current_frame().is_some_and(|frame| frame.opaque)
    }

    /// Check if all frames have finished streaming
    fn is_fully_loaded(&self) -> bool {
        match &self.all_frames_loaded {
            Some(flag) => flag.load(Ordering::Acquire),
            None => true, // No streaming = already complete
        }
    }

    /// Pull newly decoded frames from the shared buffer, then give back the
    /// frames that have already been played.
    ///
    /// Frames are only taken while the retained window has room: the decoder
    /// waits once its queue is full, so this is what keeps a long animation from
    /// decoding itself into memory faster than it is shown. The window stops
    /// meaning anything once the decoder is done with the file — there is nothing
    /// left to hold back, and the frames it finished with are frames of the
    /// animation whatever is in hand.
    fn sync_shared_frames(&mut self) {
        let Some(shared) = self.shared_frames.clone() else {
            return;
        };

        let retained_bytes: usize = self.frames.iter().map(|frame| frame.pixels.len()).sum();
        if retained_bytes < ANIMATION_RETAINED_BYTES || self.is_fully_loaded() {
            let result = shared.lock();
            if let Ok(mut streamed) = result {
                if !streamed.queue.is_empty() {
                    self.frames.extend(streamed.queue.drain(..));
                }
            }
        }

        self.release_played_frames(retained_bytes);
    }

    /// Whether the player has given back frames it already showed.
    fn frames_were_released(&self) -> bool {
        self.shared_frames
            .as_ref()
            .and_then(|shared| shared.lock().ok().map(|streamed| streamed.released))
            .unwrap_or(false)
    }

    /// Give back the frames behind the playhead once enough of them have piled up.
    /// Without this a long animation would either stop part-way at a fixed size
    /// cap or keep its whole decoded length in memory; with it, playback stays
    /// inside a fixed window while the decoder replays the file to loop.
    ///
    /// Two things have to hold before a frame is given back, and both of them are
    /// about a promise the player makes to itself: the frames that are dropped are
    /// frames that have to be decoded again. The window has to be full, which is
    /// the only thing a release is for — an animation that fits in
    /// `ANIMATION_RETAINED_BYTES` is held whole, plays from beginning to end and
    /// wraps back into the frame it started on, and the file is read once for all
    /// of it. And the decoder has to still be working, which is settled under the
    /// same lock the decoder takes as it ends a pass: a decoder that has finished
    /// with a file it decoded whole has nothing to decode again.
    fn release_played_frames(&mut self, retained_bytes: usize) {
        let Some(shared) = self.shared_frames.clone() else {
            return;
        };

        let keep_from = self.current_frame.saturating_sub(1);
        if keep_from == 0 {
            return;
        }

        if retained_bytes < ANIMATION_RETAINED_BYTES {
            return;
        }

        let played_bytes: usize = self.frames[..keep_from]
            .iter()
            .map(|frame| frame.pixels.len())
            .sum();
        if played_bytes < ANIMATION_RELEASE_BYTES {
            return;
        }

        let Ok(mut streamed) = shared.lock() else {
            return;
        };
        if streamed.decoded {
            return;
        }

        self.frames.drain(..keep_from);
        self.current_frame -= keep_from;

        streamed.released = true;
    }

    fn advance_frame(&mut self) -> bool {
        // Pull in any new frames from streaming decode
        self.sync_shared_frames();

        let frame_count = self.frames.len();
        if frame_count <= 1 {
            return false;
        }

        let fully_loaded = self.is_fully_loaded();
        let mut advanced = false;

        // Allow skipping multiple frames per call to keep up with real time.
        for _ in 0..frame_count {
            let delay = Duration::from_millis(effective_frame_delay_ms(
                &self.media_type,
                self.frames[self.current_frame].delay_ms,
            ) as u64);
            if self.last_frame_time.elapsed() >= delay {
                let next = self.current_frame + 1;
                if next < frame_count {
                    // More decoded frames ahead — advance normally
                    self.current_frame = next;
                    self.last_frame_time += delay;
                    advanced = true;
                } else if fully_loaded && !self.frames_were_released() {
                    // All frames decoded and still in memory — safe to loop back
                    // to start
                    self.current_frame = 0;
                    self.last_frame_time += delay;
                    advanced = true;
                } else {
                    // Still streaming, or the start of the animation has been
                    // released to stay inside the memory window: hold this frame
                    // until the next one arrives. Keep the next streamed frame
                    // immediately eligible instead of adding another full-frame
                    // delay.
                    self.last_frame_time = Instant::now()
                        .checked_sub(delay)
                        .unwrap_or_else(Instant::now);
                    break;
                }
            } else {
                break;
            }
        }

        // The playhead is snapped forward when it is late by more than a second
        // past the delay the frame it is on asks for — a machine that slept, a
        // decode that took far longer than the frame it was for — so that catching
        // up cannot run on across several loops.
        //
        // The frame's own delay is part of what late means, and leaving it out is
        // what a frame held for a long time used to cost: a GIF frame that waits
        // two seconds is a frame waiting for two seconds and not a playhead that
        // has fallen behind, so a snap on a flat second reset the clock every tick
        // and a delay of a whole second or more could never be reached — an
        // animation whose first frame is held for 1.2 seconds sat on that frame
        // for good.
        let waiting_for = Duration::from_millis(effective_frame_delay_ms(
            &self.media_type,
            self.frames[self.current_frame].delay_ms,
        ) as u64)
            + Duration::from_secs(1);
        if self.last_frame_time.elapsed() > waiting_for {
            self.last_frame_time = Instant::now();
        }

        advanced
    }

    /// Returns true if this media is an animation still being decoded
    fn is_streaming(&self) -> bool {
        matches!(
            self.media_type,
            MediaType::AnimatedGif | MediaType::AnimatedApng | MediaType::AnimatedWebP
        ) && !self.is_fully_loaded()
    }

    fn should_draw_streaming_overlay(&self) -> bool {
        if !self.is_streaming() || self.frames.len() > 1 {
            return false;
        }

        self.loading_start
            .map(|s| s.elapsed() <= Duration::from_millis(STREAMING_SPINNER_MAX_MS))
            .unwrap_or(false)
    }

    fn update_loading_frame(&mut self) -> bool {
        if !matches!(self.media_type, MediaType::Loading) {
            return false;
        }
        if self.last_frame_time.elapsed() >= Duration::from_millis(33) {
            if !self.frames.is_empty() {
                let width = self.frames[0].width;
                let height = self.frames[0].height;
                if let Some(start) = self.loading_start {
                    let elapsed_secs = start.elapsed().as_secs_f32();
                    let angle = elapsed_secs * 2.0 * std::f32::consts::PI * 1.2;
                    self.frames[0].set_pixels(render_loading_frame(width, height, angle));
                }
            }
            self.last_frame_time = Instant::now();
            return true;
        }
        false
    }

    fn cancel_background_work(&mut self) {
        if let Some(flag) = self.stream_cancel.take() {
            flag.store(true, Ordering::Release);
        }
    }

    /// Take the frame the media engine has ready into the one frame a preview of this kind
    /// is composed of, answering whether there was one to take.
    ///
    /// A video played by the media engine is not a file that is decoded once and drawn
    /// again — it is a new picture every frame — so the frame the placeholder was made of
    /// is the frame every one after it lands in. That is also what keeps a playing video
    /// from allocating: the buffer is written into rather than replaced, and the frames
    /// that arrive between repaints are simply the ones that are never seen.
    fn take_native_video_frame(&mut self) -> bool {
        let Some(frame) = self.frames.first_mut() else {
            return false;
        };

        let Some((width, height)) = video_player::copy_frame_into(&mut frame.pixels) else {
            return false;
        };

        frame.width = width;
        frame.height = height;
        // Every pixel the copy wrote was forced opaque, so the frame a video lands in is one
        // a repaint can copy rather than blend — which is the whole of what makes a video at
        // the size of the display affordable to draw sixty times a second (see
        // `video_player::copy_locked`).
        frame.opaque = true;

        true
    }
}

pub fn show_preview(path: &Path, x: i32, y: i32, avoid: Option<ScreenRegion>) {
    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(ref tx) = *sender {
            let _ = tx.send(PreviewMessage::Show(path.to_path_buf(), x, y, avoid));
        }
    }
}

pub fn show_preview_keyboard(
    path: &Path,
    item_left: i32,
    item_top: i32,
    item_right: i32,
    item_bottom: i32,
    avoid: Option<ScreenRegion>,
    draws_columns: bool,
) {
    // A keyboard preview is not the pointer's, so the item the pointer was last read
    // on has nothing to say about it: the box goes rather than gating a preview the
    // keyboard asked for (see `HOVER_POINTER_BOX`).
    clear_pointer_item_box();

    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(ref tx) = *sender {
            let _ = tx.send(PreviewMessage::ShowKeyboard(
                path.to_path_buf(),
                item_left,
                item_top,
                item_right,
                item_bottom,
                avoid,
                draws_columns,
            ));
        }
    }
}

pub fn hide_preview() {
    // The window coming down, the count of its coming down and the item the pointer
    // was last read on are written under one lock, so a load that is still running —
    // started under the count from before this — cannot put the preview back up after
    // it, and nothing is revealed against a hover that has already gone (see
    // `HIDDEN_EPOCH` and `HOVER_POINTER_BOX`).
    {
        let mut hidden = HIDDEN_EPOCH.lock().ok();
        if let Some(hidden) = hidden.as_mut() {
            **hidden += 1;
        }
        clear_pointer_item_box();
        unsafe {
            let hwnd = HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _);
            if !hwnd.is_invalid() {
                // Posted and not sent: the window belongs to the preview loop, and a
                // send to a loop that is busy is this thread — the hook's — held until
                // it pumps again, which is the one wait a dismissal must not have. What
                // the hide is *for* is the count above, and that is already moved, so
                // the window comes down a moment later rather than now; the loop takes
                // it down itself on the `Hide` below as well, and either of the two is
                // the same window state (see `HIDDEN_EPOCH` and `preview_stall_ms`).
                let _ = ShowWindowAsync(hwnd, SW_HIDE);
            }
        }
    }

    if let Ok(mut current) = CURRENT_MEDIA.try_lock() {
        if let Some(ref mut media) = *current {
            media.cancel_background_work();
            // The player process is ended here and the media engine is not: a session
            // belongs to the thread that made it, which is the preview loop's and not this
            // one (see `video_player`), so a stop asked for from this thread would touch
            // nothing while the sound went on playing. What ends the engine is the loop's
            // own take-down, and the `Hide` below is what calls for it — the media is left
            // where it is for that take-down to find rather than taken away, because a
            // media taken here is one nothing is left to stop.
            kill_player_process(media);
        }
    } else {
        // The media state is locked elsewhere; kill the recorded ffplay by PID
        // instead (verified to still be ffplay before terminating it).
        kill_stray_video_process();
    }

    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(ref tx) = *sender {
            let _ = tx.send(PreviewMessage::Hide);
        }
    }
}

pub fn refresh_preview() {
    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(ref tx) = *sender {
            let _ = tx.send(PreviewMessage::Refresh);
        }
    }
}

/// A preview type was switched on or off in the tray, which is a question only
/// the preview thread can answer: whether what is on screen is of that kind.
pub fn refresh_preview_types() {
    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(ref tx) = *sender {
            let _ = tx.send(PreviewMessage::RefreshTypes);
        }
    }
}

/// The render tier is done with a document. Sent from the engine thread through
/// the same channel every other message arrives on, so the preview loop learns
/// about a page the moment it exists.
pub fn notify_office_render(path: &Path, generation: u64, ok: bool) {
    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(ref tx) = *sender {
            let _ = tx.send(PreviewMessage::OfficeRenderReady {
                path: path.to_path_buf(),
                generation,
                ok,
            });
        }
    }
}

/// A video's probe is done. Sent from the thread the probe ran on, through the same
/// channel every other message arrives on, so the hover that was waiting for it is
/// replayed the moment there is an answer to place it with.
fn notify_video_probed(path: &Path, generation: u64) {
    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(ref tx) = *sender {
            let _ = tx.send(PreviewMessage::VideoProbed {
                path: path.to_path_buf(),
                generation,
            });
        }
    }
}

/// The ImageMagick engine is done with a file. Sent from the engine's own thread, through the
/// same channel every other answer arrives on, so the hover that was waiting is replayed the
/// moment there is a picture to place it with — or taken down, where the answer is that the
/// file is not one the engine can read.
pub fn notify_magick_ready(path: &Path, generation: u64, ok: bool) {
    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(ref tx) = *sender {
            let _ = tx.send(PreviewMessage::MagickReady {
                path: path.to_path_buf(),
                generation,
                ok,
            });
        }
    }
}

/// And the PeaZip engine, on the same terms: the archive has been listed, or it is one the engine
/// will not list. What is waiting on it is a hover that has already asked and is showing the
/// spinner — the page an archive is drawn as cannot be measured before its listing exists — and
/// what the answer does is replay that hover, with the listing in hand or with nothing at all.
pub fn notify_peazip_ready(path: &Path, generation: u64, ok: bool) {
    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(ref tx) = *sender {
            let _ = tx.send(PreviewMessage::PeazipReady {
                path: path.to_path_buf(),
                generation,
                ok,
            });
        }
    }
}

/// Probe a video's geometry on a thread of its own, and tell the preview loop.
///
/// The probe is two external processes and the hover waits for the slower of them, so it
/// is done here rather than on the preview thread: what is on screen while it runs is the
/// waiting spinner, and the hover it belongs to is replayed when the answer lands (see
/// `video_probe_due` and `video_probe` in the preview loop). A probe whose hover has moved
/// on is not wasted — what it answers is held for the next hover of the file — so nothing
/// here is cancelled or waited for.
///
/// What is not left to the probe is whether the hover is answered at all: the wait is not
/// one an engine or a cap will ever end, so the answer is sent whatever the probe did —
/// including a probe that panicked, which is a thread that would otherwise unwind past the
/// notify and leave the spinner standing over a file it has finished with (see
/// `video_probe_due` and `awaiting_engine`).
fn spawn_video_probe(path: PathBuf, generation: u64) {
    std::thread::spawn(move || {
        let _ =
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| probe_video_geometry(&path)));
        notify_video_probed(&path, generation);
    });
}

/// What a box measured off this thread was measured against.
///
/// It is part of what a held box is keyed by because a box is only the answer for what it was
/// measured against: a page's own size is the file's, whatever it is drawn on, while a
/// listing's page is wrapped to the room it is shown in at the text settings it is wrapped
/// for — and a box held for another room is a page laid out to the wrong one.
#[derive(Clone, PartialEq, Eq, Hash)]
enum MeasureScope {
    /// The file's own size: a page's, a plate's, a drawing's declared extent, a specimen's
    /// box.
    File,
    /// A page wrapped to the room it is drawn in, at the text settings it is wrapped for.
    Room {
        cap_width: u32,
        cap_height: u32,
        dpi: u32,
        theme: TextTheme,
        font_scale_percent: u32,
    },
}

/// A box measured off the preview thread, and what it was measured against.
struct MeasuredBox {
    path: PathBuf,
    version: FileVersion,
    scope: MeasureScope,
    /// The box, or `None` for a file its reader has no answer for — which is an answer too,
    /// and one worth holding: a document that will not open is not one to read again on every
    /// hover.
    size: Option<(u32, u32)>,
}

/// The boxes measured off the preview thread, newest first.
///
/// It is a table of its own rather than the readers' own memos, and that is the point: what a
/// reader remembers is dropped wholesale when it fills, and a box that fell out of one of
/// those memos would be measured again the moment it was asked for — on the thread that draws
/// the hover, which is what these measures are kept off. What is held here is answered from
/// here, and the reader's own memo is left to the side that draws the file (see
/// `measured_off_the_tick`).
static MEASURED_BOXES: Lazy<Mutex<Vec<MeasuredBox>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// Entries the table of measured boxes holds before it is emptied.
const MEASURED_BOXES_MAX_ENTRIES: usize = 256;

/// The measures running right now, by the file being measured.
///
/// One thread per file rather than one per hover: a hover that lands on a file whose measure is
/// already running waits for that one instead of starting a second read of the same bytes, and
/// what the layout asks to place that wait is a question about this list (see
/// `measure_waiting`). A file whose version changes while it is being read is measured again by
/// the next hover: this list says a read is running, and what that read answered is held only
/// for the version it was a read of (see `hold_box`).
static MEASURING: Lazy<Mutex<Vec<PathBuf>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// Entries the list of running measures holds before it is emptied. It is a list of reads in
/// flight, so it is never more than a handful long; the ceiling is what a stuck thread would
/// cost.
const MEASURING_MAX_ENTRIES: usize = 64;

/// The box held for this version of this file, measured against `scope` — or `None` when
/// nothing is held for it, which is not the same answer as a held `None`.
fn held_box(
    path: &Path,
    version: &FileVersion,
    scope: &MeasureScope,
) -> Option<Option<(u32, u32)>> {
    let boxes = MEASURED_BOXES.lock().ok()?;

    boxes
        .iter()
        .find(|held| held.path == path && held.version == *version && held.scope == *scope)
        .map(|held| held.size)
}

/// Hold the box a measure answered with.
fn hold_box(path: &Path, version: &FileVersion, scope: &MeasureScope, size: Option<(u32, u32)>) {
    let Ok(mut boxes) = MEASURED_BOXES.lock() else {
        return;
    };

    boxes.retain(|held| !(held.path == path && held.version == *version && held.scope == *scope));
    if boxes.len() >= MEASURED_BOXES_MAX_ENTRIES {
        boxes.clear();
    }

    boxes.insert(
        0,
        MeasuredBox {
            path: path.to_path_buf(),
            version: version.clone(),
            scope: scope.clone(),
            size,
        },
    );
}

/// Say that this file is being measured, answering whether one already was.
fn begin_measure(path: &Path) -> bool {
    let Ok(mut measuring) = MEASURING.lock() else {
        return false;
    };

    if measuring.iter().any(|running| running == path) {
        return false;
    }

    if measuring.len() >= MEASURING_MAX_ENTRIES {
        measuring.clear();
    }

    measuring.push(path.to_path_buf());

    true
}

/// Say that the measure of this file is done with.
fn end_measure(path: &Path) {
    let Ok(mut measuring) = MEASURING.lock() else {
        return;
    };

    measuring.retain(|running| running != path);
}

/// Whether this file is being measured right now: the question the layout places the wait by.
///
/// It is asked of the file the hover is on, straight after the layout measured it, and what it
/// says is whether that measure handed back the wait for a read rather than a box. The two
/// cannot disagree — the wait is placed exactly when the measure the layout has just taken
/// started this file's read (see `measured_off_the_tick`).
fn measure_waiting(path: &Path) -> bool {
    MEASURING
        .lock()
        .map(|measuring| measuring.iter().any(|running| running == path))
        .unwrap_or(false)
}

/// Measure a file on a thread of its own, and tell the preview loop.
///
/// The measure is a read that can be felt — a PDF opened, an archive's table of contents
/// walked, a document parsed, a specimen read — and the hover waits for it, so it is taken
/// here rather than on the preview thread: what is on screen while it runs is the spinner, and
/// the hover it belongs to is replayed when the answer lands, laid out at the box that answer
/// is held under. It is the shape a video's probe has, for the same reason (see
/// `spawn_video_probe`).
///
/// A measure whose hover has moved on is not wasted: what it answered is held for the next
/// hover of the file, so nothing here is cancelled or waited for.
///
/// What a hover waits for is the answer rather than the read, and the answer is sent whatever
/// became of the read — a measure that panicked included. A thread that unwound past the two
/// calls below would leave the file on the list of reads running with nothing to take it off
/// it, and what a hover on that file would be from then on is a spinner nothing ends: the
/// layout goes on placing the wait, since a read for the file is running, and the cap a wait
/// is given is the engines' and does not stand behind a read (see `measured_off_the_tick`
/// and `awaiting_engine`). A read that came apart is answered with the reader's own "nothing
/// for this file" and held like it, so a file whose read comes apart is not read again on
/// every hover either (see `spawn_video_probe`, whose probe answers through the same guard).
fn spawn_measure_probe(
    path: PathBuf,
    version: FileVersion,
    scope: MeasureScope,
    measure: impl FnOnce() -> Option<(u32, u32)> + Send + 'static,
) {
    std::thread::spawn(move || {
        // A read that panicked is no box, and no box is an answer; what it must not be is an
        // unwind past the mark that says the read is done with (see above).
        let size = std::panic::catch_unwind(std::panic::AssertUnwindSafe(measure)).unwrap_or(None);

        // The box is held before the read is marked done, so a hover that arrives while this
        // thread is between the two finds the answer rather than starting a read of its own.
        hold_box(&path, &version, &scope, size);
        end_measure(&path);

        notify_measured(&path, size);
    });
}

/// A measure has been taken off the preview thread, and the box it answered with is held. Sent
/// from the thread the measure ran on, through the same channel every other answer arrives on,
/// so the hover that was waiting for it is replayed the moment there is a box to place it with
/// — or, where the reader has no box for the file at all, told that the wait is over (see
/// `MeasureProbed`).
fn notify_measured(path: &Path, size: Option<(u32, u32)>) {
    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(ref tx) = *sender {
            let _ = tx.send(PreviewMessage::MeasureProbed {
                path: path.to_path_buf(),
                size,
            });
        }
    }
}

/// The box a measure that reads a file answers with: the box this side already holds, or the
/// wait for one that is being measured now.
///
/// `measure` is the reader's own measure — the same call this side would otherwise make on its
/// own thread — and it runs on a thread of this function's own making, once per file version
/// and scope. What comes back to the hovering call meanwhile is the spinner's own box, which
/// is what the layout places a hover at until the answer lands (see `measure_waiting` and
/// `MeasureProbed`).
fn measured_off_the_tick(
    path: &Path,
    scope: MeasureScope,
    measure: impl FnOnce() -> Option<(u32, u32)> + Send + 'static,
) -> Option<(u32, u32)> {
    let version = file_version(path);

    if let Some(held) = held_box(path, &version, &scope) {
        return held;
    }

    if begin_measure(path) {
        spawn_measure_probe(path.to_path_buf(), version, scope, measure);
    }

    Some((office_preview::WAITING_BOX, office_preview::WAITING_BOX))
}

/// Which preview surface the pointer is currently on.
#[derive(Clone, Copy)]
pub struct PreviewCursorHover {
    pub image: bool,
    pub video: bool,
    /// A document the engine draws, in a window of its own: a preview of this app's in every
    /// way but the window it is drawn in, and one the pointer takes the same way — a
    /// document is closed by the pointer arriving at it, exactly as a picture is.
    pub engine: bool,
}

impl PreviewCursorHover {
    pub const NONE: Self = Self {
        image: false,
        video: false,
        engine: false,
    };

    pub fn any(self) -> bool {
        self.image || self.video || self.engine
    }
}

/// Single shared pointer probe for both preview kinds. Callers gate it on
/// "a preview can be under the pointer"; the fast path keeps the cost at a few
/// atomic reads whenever nothing is on screen.
pub fn cursor_preview_hover() -> PreviewCursorHover {
    let preview_hwnd = PREVIEW_HWND.load(Ordering::SeqCst);
    let video_hwnd = VIDEO_HWND.load(Ordering::SeqCst);
    let video_pid = VIDEO_PID.load(Ordering::SeqCst);
    // A document is drawn by the engine, in a window of its own rather than this app's, so
    // its surface is asked for by its own handle: it is the same preview to the pointer.
    let engine_hwnd = webview_preview::showing_hwnd();

    // PREVIEW_HWND is created once at startup and never cleared, so visibility
    // is what tells us whether the layered window is actually on screen.
    let preview_visible =
        preview_hwnd != 0 && unsafe { IsWindowVisible(HWND(preview_hwnd as *mut _)).as_bool() };
    if !preview_visible && video_hwnd == 0 && video_pid == 0 && engine_hwnd == 0 {
        return PreviewCursorHover::NONE;
    }

    unsafe {
        use windows::Win32::UI::WindowsAndMessaging::{GetCursorPos, IsChild, WindowFromPoint};

        let mut cursor_pos = POINT::default();
        if GetCursorPos(&mut cursor_pos).is_err() {
            return PreviewCursorHover::NONE;
        }

        let hwnd_under_cursor = WindowFromPoint(cursor_pos);
        if hwnd_under_cursor.is_invalid() {
            return PreviewCursorHover::NONE;
        }

        let hwnd_ptr = hwnd_under_cursor.0 as isize;
        let image = preview_hwnd != 0 && hwnd_ptr == preview_hwnd;
        // A document is drawn by a browser inside the engine's window, so what the pointer
        // is over is a window of the browser's — a child of the engine's, one or two levels
        // down — and not the engine's own window at all. The engine's window is what the
        // preview is, so the question is whether the window under the pointer is that
        // window or one inside it: comparing the two handles alone never matched, and the
        // touch that closes a document was the one thing that never came from here.
        let engine = engine_hwnd != 0
            && (hwnd_ptr == engine_hwnd
                || IsChild(HWND(engine_hwnd as *mut _), hwnd_under_cursor).as_bool());

        // A hit on the stored HWND is enough; the process-ID fallback covers the
        // race window where ffplay's window exists but VIDEO_HWND isn't stored yet.
        let mut video = video_hwnd != 0 && hwnd_ptr == video_hwnd;
        if !video && video_pid != 0 {
            use windows::Win32::UI::WindowsAndMessaging::GetWindowThreadProcessId;

            let mut window_pid: u32 = 0;
            GetWindowThreadProcessId(hwnd_under_cursor, Some(&mut window_pid));
            video = window_pid == video_pid;
        }

        PreviewCursorHover {
            image,
            video,
            engine,
        }
    }
}

/// Screen-space box of the preview surface that is on screen right now, if any.
/// The Explorer hook uses it to decide whether a keyboard preview was placed
/// over the parked pointer, so a pointer sitting under the preview cannot drive
/// previews or dismiss them.
///
/// An SVG document's preview is the engine's window rather than this app's, so its box
/// is asked for as well — a document is a preview of this app's in every way but the
/// window it is drawn in.
pub fn preview_screen_rect() -> Option<(i32, i32, i32, i32)> {
    if let Some(rect) = webview_preview::screen_rect() {
        return Some(rect);
    }

    unsafe {
        let candidates = [
            PREVIEW_HWND.load(Ordering::SeqCst),
            VIDEO_HWND.load(Ordering::SeqCst),
        ];

        for hwnd_value in candidates {
            if hwnd_value == 0 {
                continue;
            }

            let hwnd = HWND(hwnd_value as *mut _);
            if !IsWindowVisible(hwnd).as_bool() {
                continue;
            }

            let mut rect = RECT::default();
            if GetWindowRect(hwnd, &mut rect).is_ok()
                && rect.right > rect.left
                && rect.bottom > rect.top
            {
                return Some((rect.left, rect.top, rect.right, rect.bottom));
            }
        }
    }

    None
}

/// Decode an image by sniffing its magic bytes rather than by the name it is written under.
///
/// A picture is decoded by its header always: a `.dat` holding a PNG is a picture, and a
/// `.png` holding a container is a container — which is the kind's question, and it has been
/// asked by the time anything here is reached.
fn decode_image_with_header_check(path: &PathBuf) -> Option<image::DynamicImage> {
    let mut reader = image::ImageReader::open(path)
        .ok()?
        .with_guessed_format()
        .ok()?;
    reader.limits(image_decode_limits());

    reader.decode().ok()
}

/// Read image dimensions by sniffing magic bytes instead of trusting the extension.
///
/// A format the `image` crate has no reader for at all is a file it cannot answer for
/// rather than a file that is not a picture, so the codec Windows has is asked before
/// the answer is no: what a hover onto a `.heic`, or onto a still `.webp`, is measured
/// from is its own frame; see `wic_image`.
fn image_dimensions_with_header_check(path: &PathBuf) -> Option<(u32, u32)> {
    image::ImageReader::open(path)
        .ok()?
        .with_guessed_format()
        .ok()?
        .into_dimensions()
        .ok()
        .or_else(|| codec_dimensions(path))
}

/// The size a picture of a format this app's own decoder does not read is measured
/// from: the frame the codec Windows has for it reports, and — where that codec is the
/// WebP one and the machine has none, which is what a Windows 10 machine usually is —
/// libwebp's, which is in the binary; see `wic_image` and `webp_image`. A `.dds` of a
/// format the codec does not read is measured by the decoder this app carries for the
/// rest of that format; see `dds_image`.
fn codec_dimensions(path: &Path) -> Option<(u32, u32)> {
    wic_image::dimensions(path)
        .or_else(|| dds_image::dimensions(path))
        .or_else(|| webp_image::dimensions(path))
}

/// The eight-bit picture a picture whose samples are light is shown as.
///
/// `None` is every picture that holds levels already — a PNG, a JPEG, a BMP, and every
/// other format this app's reader decodes — because a level put through a transfer
/// function a second time is a washed-out picture rather than a corrected one.
///
/// What arrives as one of these is an `.exr` and a Radiance `.hdr`, which are the two
/// float kinds the `image` crate has: three channels for a `.hdr` — and for the `.exr`
/// that was written without an alpha — and four for the one that carries it. A
/// single-channel file is the decoder's to widen, and it arrives as one of the two as
/// well. What the curve does with them is `tone_map`'s; what is done here is the shape,
/// and an alpha is not light and is not put through a curve.
fn tone_mapped_image(img: &image::DynamicImage) -> Option<image::RgbaImage> {
    let (samples, channels, width, height) = match img {
        image::DynamicImage::ImageRgb32F(buffer) => (
            buffer.as_raw().as_slice(),
            3,
            buffer.width(),
            buffer.height(),
        ),
        image::DynamicImage::ImageRgba32F(buffer) => (
            buffer.as_raw().as_slice(),
            4,
            buffer.width(),
            buffer.height(),
        ),
        _ => return None,
    };

    let tone = tone_map::ToneMap::current();
    let count = width as usize * height as usize;
    let mut pixels = vec![0u8; count * 4];

    for index in 0..count {
        let texel = &samples[index * channels..][..channels];

        let alpha = if channels == 4 { texel[3] } else { 1.0 };

        let at = index * 4;
        pixels[at] = tone.encode(texel[0]);
        pixels[at + 1] = tone.encode(texel[1]);
        pixels[at + 2] = tone.encode(texel[2]);
        pixels[at + 3] = tone_mapped_alpha(alpha);
    }

    image::RgbaImage::from_raw(width, height, pixels)
}

/// An alpha as a level: brought into the range and scaled, with no curve applied — what
/// the alpha of a float texture gets as well, and for the same reason: coverage is not
/// light, and a picture composited over a backdrop with a curve put on its alpha is a
/// picture that fades differently from every other one beside it.
fn tone_mapped_alpha(alpha: f32) -> u8 {
    if !alpha.is_finite() {
        return 0;
    }

    (alpha.clamp(0.0, 1.0) * 255.0 + 0.5) as u8
}

/// Convert RGBA pixels to BGRA for Windows GDI
///
/// Shared with the readers that hand back their own pixels rather than going through the
/// `image` crate — a texture this app decodes itself, a codec's own frame — because a
/// frame is composed in one order whatever produced it (see `dds_image`).
pub(crate) fn rgba_to_bgra(rgba: &[u8]) -> Vec<u8> {
    let mut bgra = Vec::with_capacity(rgba.len());
    for chunk in rgba.chunks(4) {
        if chunk.len() == 4 {
            bgra.push(chunk[2]); // B
            bgra.push(chunk[1]); // G
            bgra.push(chunk[0]); // R
            bgra.push(chunk[3]); // A
        }
    }
    bgra
}

/// The backdrop a picture is drawn over — and every other preview that is not a
/// document: a PDF page, a painted frame, a page Office rendered.
fn current_image_background() -> TransparentBackground {
    CONFIG
        .lock()
        .map(|cfg| cfg.image_background)
        .unwrap_or(DEFAULT_IMAGE_BACKGROUND)
}

/// The backdrop a font specimen is drawn over, which is a page of its own: a document's
/// backdrop is the one its shapes are drawn on, and a specimen's is the one its glyphs are.
fn current_font_background() -> TransparentBackground {
    CONFIG
        .lock()
        .map(|cfg| cfg.font_background)
        .unwrap_or(DEFAULT_FONT_BACKGROUND)
}

/// The backdrop a `.dds` texture is drawn over, which the tray keeps apart from a
/// picture's: a texture's alpha channel is as often a mask or a channel nobody filled in as
/// it is transparency, so what is behind one is a question of its own (see `dds_image`).
fn current_dds_background() -> TransparentBackground {
    CONFIG
        .lock()
        .map(|cfg| cfg.dds_background)
        .unwrap_or(DEFAULT_DDS_BACKGROUND)
}

/// The backdrop a design document is drawn over, which the tray keeps apart from a
/// picture's: what is previewed is the picture the file keeps of the whole document, and
/// a designer's transparency is the document's own rather than a photograph's.
fn current_design_background() -> TransparentBackground {
    CONFIG
        .lock()
        .map(|cfg| cfg.design_background)
        .unwrap_or(DEFAULT_DESIGN_BACKGROUND)
}

/// The backdrop a vector drawing is drawn over, which the tray keeps apart from a
/// picture's.
///
/// The kind holds two halves that answer this the same way for different reasons: a
/// metafile says what was drawn and nothing about the sheet under it, so what stands
/// behind the marks is this app's (see `metafile_image`), and an SVG document is drawn on
/// a page of the engine's own, so what it is given is a colour (see `webview_preview`).
fn current_vector_background() -> TransparentBackground {
    CONFIG
        .lock()
        .map(|cfg| cfg.vector_background)
        .unwrap_or(DEFAULT_VECTOR_BACKGROUND)
}

/// How loud a video is played, which is read when one is started rather than when the
/// setting changes: a preview is a few seconds long, and the next one is played at
/// whatever the volume is by then.
fn current_video_volume() -> u32 {
    CONFIG.lock().map(|cfg| cfg.video_volume).unwrap_or(0)
}

/// The volume a sound is previewed at, read the way the video's is: from the configuration at
/// the moment a player is started, so a change in the tray reaches the next hover.
fn current_audio_volume() -> u32 {
    CONFIG.lock().map(|cfg| cfg.audio_volume).unwrap_or(0)
}

/// Where a sound starts, read the way the volume is and at the same moment: from the
/// configuration as a player is started, so a change in the tray reaches the next hover — and
/// a sound already playing is left where it is rather than dropped somewhere else, which is
/// what a seek asked of a running engine would be.
fn current_audio_seek() -> AudioSeek {
    CONFIG
        .lock()
        .map(|cfg| cfg.audio_seek)
        .unwrap_or(DEFAULT_AUDIO_SEEK)
}

/// What a sound's card is built with: the theme and the text size, which are the two settings
/// a painted preview answers to.
fn current_audio_options() -> AudioPreviewOptions {
    CONFIG
        .lock()
        .map(|cfg| AudioPreviewOptions {
            theme: cfg.theme,
            font_scale_percent: cfg.text_font_scale_percent,
        })
        .unwrap_or(AudioPreviewOptions {
            theme: TextTheme::Light,
            font_scale_percent: DEFAULT_TEXT_FONT_SCALE_PERCENT,
        })
}

/// The backdrop an engine-drawn preview of `path` is drawn over: the kind decides it, the
/// same way it decides everything else about a document. The one engine draws both kinds
/// this app hands it — an SVG document, which is a vector drawing, and a font file's
/// specimen — and each has a backdrop of its own.
fn engine_background(path: &Path) -> TransparentBackground {
    if font_formats::is_font_file(path) {
        current_font_background()
    } else {
        current_vector_background()
    }
}

/// The kind of engine-drawn preview `path` would get, when it is one of the two the browser
/// draws: a document, or a font file's specimen.
///
/// It stands in for the file's name where a hover is replayed or taken down: what is on
/// screen for either kind is the engine's window rather than anything this app composed, so
/// what the loop asks about one it asks about the other — the same way the loader asks the
/// name gates in one order.
fn engine_kind_of(path: &Path) -> Option<PreviewType> {
    // What the file's own bytes say comes first, as it does for the loader that draws it: a
    // picture left under a font's name is drawn here as the picture it is rather than by the
    // engine, and a font left under a document's name is the engine's. A drawing is the one
    // kind whose two halves have to be told apart by the name as well — the browser draws a
    // document, and the drawing layer replays a metafile or a PostScript program, which is
    // no engine window at all — so the name answers which half of that kind a file is.
    let content = CONFIG
        .lock()
        .ok()
        .map(|config| crate::formats::content_type::of(path, &config))
        .unwrap_or(crate::formats::content_type::Content::Unknown);

    if let crate::formats::content_type::Content::Kind(kind) = content {
        return match kind {
            PreviewType::Vector if svg_preview::is_svg_file(path) => Some(PreviewType::Vector),
            PreviewType::Fonts => Some(PreviewType::Fonts),
            _ => None,
        };
    }

    if svg_preview::is_svg_file(path) {
        return Some(PreviewType::Vector);
    }

    if font_formats::is_font_file(path) {
        return Some(PreviewType::Fonts);
    }

    None
}

fn current_webp_playback_fps() -> u32 {
    CONFIG
        .lock()
        .map(|cfg| sanitize_webp_playback_fps(cfg.webp_playback_fps))
        .unwrap_or(DEFAULT_WEBP_PLAYBACK_FPS)
}

/// Every scale a hover is laid out by, read together so that a measure and the render
/// that follows it agree on all of them — one read of the configuration rather than a
/// handful of them, and one answer per kind of preview.
fn current_hover_scales() -> HoverScales {
    CONFIG
        .lock()
        .map(|cfg| HoverScales {
            picture: cfg.preview_scale,
            video: cfg.video_scale,
            animated: cfg.animated_scale,
            ebook: cfg.ebook_scale,
            document: cfg.document_scale,
            font: cfg.font_scale,
            design: cfg.design_scale,
            vector: cfg.vector_scale,
        })
        .unwrap_or(HoverScales {
            picture: PreviewScale::Percent(DEFAULT_PREVIEW_SCALE_PERCENT),
            video: PreviewScale::Percent(DEFAULT_VIDEO_SCALE_PERCENT),
            animated: PreviewScale::Percent(DEFAULT_ANIMATED_SCALE_PERCENT),
            ebook: DEFAULT_EBOOK_SCALE,
            document: DEFAULT_DOCUMENT_SCALE,
            font: DEFAULT_FONT_SCALE,
            design: DEFAULT_DESIGN_SCALE,
            vector: DEFAULT_VECTOR_SCALE,
        })
}

/// The theme, Markdown rendering and font size the configuration currently
/// selects, read once per hover so a measure and the render that follows agree.
fn current_text_options() -> TextPreviewOptions {
    CONFIG
        .lock()
        .map(|cfg| TextPreviewOptions {
            theme: cfg.theme,
            markdown_mode: cfg.markdown_mode,
            font_scale_percent: cfg.text_font_scale_percent,
            full_mode: cfg.text_preview_full_mode,
        })
        .unwrap_or(TextPreviewOptions {
            theme: TextTheme::Light,
            markdown_mode: MarkdownMode::Rendered,
            font_scale_percent: DEFAULT_TEXT_FONT_SCALE_PERCENT,
            full_mode: true,
        })
}

/// The options an archive preview is laid out and painted with, read from the
/// configuration the way a text preview's are.
fn current_archive_options() -> ArchivePreviewOptions {
    CONFIG
        .lock()
        .map(|cfg| ArchivePreviewOptions {
            theme: cfg.theme,
            font_scale_percent: cfg.text_font_scale_percent,
        })
        .unwrap_or(ArchivePreviewOptions {
            theme: TextTheme::Light,
            font_scale_percent: DEFAULT_TEXT_FONT_SCALE_PERCENT,
        })
}

/// Whether the preview on screen is one whose appearance is baked into its
/// painted frame rather than recomposited from shared pixels.
fn current_media_is_painted() -> bool {
    CURRENT_MEDIA
        .lock()
        .map(|media| {
            media
                .as_ref()
                .map(|media| media.media_type.is_painted())
                .unwrap_or(false)
        })
        .unwrap_or(false)
}

/// Which kind of preview is on screen, if one is.
///
/// The media the renderer built already knows what it is, so this is the kind of
/// preview that is up rather than a classification of the file it came from. A
/// spinner is no kind: what it is standing in for has not been decided yet.
fn current_media_kind() -> Option<PreviewType> {
    CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| media.as_ref().map(|media| media.media_type.kind()))
        .flatten()
}

/// How long a preview waits for a page before it stops waiting. An engine can be
/// held by a dialog inside Office, and a spinner that never ends is worse than
/// the picture the document saved — so past this the preview comes down and the
/// file is left alone.
const OFFICE_RENDER_WAIT_SECS: u64 = 25;

/// How long the pointer has to have been on a file before an engine is asked to come up for it.
///
/// An engine's launch is a second or more, and an engine that is up is what the first hover of a
/// session on a document of its kind is otherwise paying for: the ask is made while the hover
/// waits, so that the start overlaps what it is waiting for rather than following it (see
/// `warm_engines_for`).
///
/// What the wait is for is the pointer rather than the engine. A hand crossing a folder is on a
/// new file every few dozen milliseconds, and what a hover is about is the file it comes to rest
/// on: an engine asked for on every file a sweep touches would be a machine full of
/// applications for a folder nobody looked at. A sixth of a second is long enough for a hand
/// that was going somewhere else to be somewhere else, and short enough that a hand that stopped
/// has the engine starting before it has finished looking at the file.
const WARM_SETTLE_MS: Duration = Duration::from_millis(150);

/// The file a hover message is about.
fn show_path(show: &PreviewMessage) -> Option<&PathBuf> {
    match show {
        PreviewMessage::Show(path, ..) | PreviewMessage::ShowKeyboard(path, ..) => Some(path),
        _ => None,
    }
}

/// A hover about to be replayed, anchored where it belongs now.
///
/// A mouse hover is replayed where the pointer is rather than where the hover opened:
/// what is being put back is the preview of the file under the hand, and the display
/// it is being put back on is the one the pointer is on now — the point it opened from
/// resolves the display it came from, which is the one that is gone. A keyboard hover
/// is the item's own place and is replayed as it came.
///
/// It is asked of every mouse hover the loop takes up, and not only of the ones being
/// replayed. The point a `Show` arrives with is one the Explorer hook sampled at the top
/// of its own tick — ahead of a walk through the shell and of the look for the `Avoid`
/// region the layout is placed by — so the better part of a frame has passed by the time
/// anything is laid out from it, and a fast hand covers dozens of pixels in that time. The
/// one thing the layout does with the point is keep the box clear of it, which is worth
/// something only while the point is still the hand's.
fn replay_where_the_pointer_is(show: Option<PreviewMessage>) -> Option<PreviewMessage> {
    match (show, cursor_position()) {
        (Some(PreviewMessage::Show(path, _, _, avoid)), Some(cursor)) => {
            Some(PreviewMessage::Show(path, cursor.x, cursor.y, avoid))
        }
        (show, _) => show,
    }
}

/// Whether this hover is owed a render: an Office document with no page in the
/// cache yet — or one whose page was exported narrower than a render asked for this
/// hover's room would be, which is a deck that was first previewed on a smaller
/// display — with the render tier switched on.
fn office_render_is_due(path: &Path, width: u32) -> bool {
    if !office_formats::is_office_preview(path) || !office_render::enabled() {
        return false;
    }

    // A file whose bytes are another kind is not a document to render, whatever it is
    // called: a picture left under a `.docx` name is drawn as the picture it is, and asking
    // Office for a page would start an engine for a file that is not its own — which is the
    // one thing the question of content exists to prevent (see `content_type`).
    if content_names_another_kind(path, PreviewType::Document) {
        return false;
    }

    // Where the render engine is the one that draws this document's page — because the tray
    // has asked it for every Office document, or because the family's application is not
    // installed — there is no engine of this tier's to ask, and the request goes there
    // instead (see `libre_render_is_due`, and `office_formats::page_engine` for the question
    // both sides ask).
    if office_formats::page_engine(path) != Some(OfficeEngine::MicrosoftOffice) {
        return false;
    }

    match office_render::held_page(path) {
        Some(page) => office_render::page_is_narrower_than(path, &page, width),
        None => true,
    }
}

/// Whether the file's own bytes name one of this app's kinds *other* than `kind`.
///
/// It is the question an engine tier asks before it starts anything, and what it says no to
/// is a file that is called what it is: a name and a content that agree are answered with no
/// opinion at all, and only a disagreement — a picture under a document's name — is a file
/// whose engine must not be started.
fn content_names_another_kind(path: &Path, kind: PreviewType) -> bool {
    let content = CONFIG
        .lock()
        .ok()
        .map(|config| crate::formats::content_type::of(path, &config))
        .unwrap_or(crate::formats::content_type::Content::Unknown);

    matches!(
        content,
        crate::formats::content_type::Content::Kind(named) if named != kind
    )
}

/// Ask the render tier for the page a hover needs, at the moment that hover is
/// installed, and answer what is now being waited on.
///
/// Every other format has something to draw the moment a hover is up, because it
/// is read or decoded on this side. A document's page is the one thing that does
/// not exist until Office has drawn it, and none of that work can begin before
/// it is asked for — so asking late is waiting twice, once for the timer and once
/// for the render. A hover is only ever raised for the file the pointer is on, so
/// there is nothing to wait for: the page is asked for as soon as there is a
/// hover to ask for it, and the engine that request starts is kept warm for the
/// documents asked for after it.
fn request_office_render(
    path: &Path,
    generation: u64,
    width: u32,
    height: u32,
) -> Option<(PathBuf, u64)> {
    if !office_render_is_due(path, width) {
        return None;
    }

    office_render::request(path, width, height, generation);
    Some((path.to_path_buf(), generation))
}

/// Whether this hover is owed a page by the render engine: a document the engine draws — one
/// of its own lists, or an Office document whose own application is not installed — with an
/// engine installed to draw it and no page drawn for this version of it yet.
///
/// It is the question `office_render_is_due` asks of a document whose application is there,
/// asked of the documents whose engine is a whole application rather than an automation
/// server. Three things ask it: the layout, which measures a document like this as the wait
/// for a page; the loader, which answers with it that a hover is still waiting rather than
/// failed; and the loop, which asks the engine for the page only where there is one to ask
/// for.
///
/// Which kind the page is shown under is part of the answer rather than a second question:
/// a document of the engine's own is shown under `Libre`, and an Office document the engine
/// draws where Office cannot is shown under `Office`, at that kind's scale and over that
/// kind's backdrop — the file is what it is whichever engine drew it (see
/// `libre_formats::engine_page_kind`).
fn libre_render_is_due(path: &Path) -> bool {
    libre_formats::engine_page_kind(path).is_some_and(PreviewType::enabled)
        && libreoffice_render::available()
        && libreoffice_render::rendered_page(path).is_none()
        && !libreoffice_render::refused(path)
}

/// Ask the engine for the page this hover needs, and answer what is now being waited on.
///
/// The same answer, and for the same reason, as `request_office_render`: a page does not
/// exist until an engine has drawn it, and asking late is waiting twice. Nothing is waited
/// on here either — the conversion runs on the engine's own thread — so what comes back is
/// the wait, and the loop watches the folder the page lands in for it.
fn request_libre_render(path: &Path, generation: u64) -> Option<(PathBuf, u64)> {
    if !libre_render_is_due(path) {
        return None;
    }

    libreoffice_render::request(path);
    Some((path.to_path_buf(), generation))
}

/// Whether this hover is owed a picture by the ImageMagick engine: a file the engine develops,
/// with an engine installed to develop it and nothing developed for this version of it yet —
/// in hand for the hover that asked, or written down for the hovers after it.
///
/// It is the same question `libre_render_is_due` is, asked of an engine that is a converter
/// rather than an application — one that reads a file, writes one and exits, which is why
/// there is a process to wait for and a page rather than an instance to keep. Three things ask
/// it: the layout, which measures a file like this as the wait for a picture; the loader, which
/// answers with it that a hover is still waiting rather than failed; and the loop, which asks
/// the engine for the picture only where there is one to ask for.
fn magick_render_is_due(path: &Path) -> bool {
    // The file's own bytes first, the name after them, exactly as the render engine's own
    // question asks it: a picture renamed to a name no list holds is still the engine's to
    // develop, and one whose bytes are another kind is not a file to start it for (see
    // `magick_formats::is_engine_picture`).
    magick_formats::is_engine_picture(path)
        && PreviewType::Magick.enabled()
        && imagemagick_render::available()
        && !imagemagick_render::refused(path)
        && !imagemagick_render::developed(path)
        && !imagemagick_render::has_page(path)
        // A raw sample dump whose own length does not settle a shape is not a file to ask about:
        // the engine would answer that it must be told a size, which is a launch spent on nothing
        // (see `raw_geometry`).
        && (!imagemagick_render::is_raw_sample(path)
            || imagemagick_render::raw_geometry(path).is_some())
}

/// Ask the engine for the picture this hover needs, and answer what is now being waited on.
///
/// The same answer, and for the same reason, as `request_libre_render`: there is no picture
/// until the engine has developed one, and asking late is waiting twice. Nothing is waited on
/// here either — the conversion runs on the engine's own thread — so what comes back is the
/// wait, and the hover is replayed when the engine answers. The room is part of the request
/// rather than of the wait, and it is the room the display has rather than the one the wait
/// was laid out at: what the engine is told is how large a picture it may write, and a
/// picture written into too small a box is one no later layout can draw any larger (see
/// `PendingLoad::room`).
fn request_magick_render(path: &Path, generation: u64, room: (u32, u32)) -> Option<(PathBuf, u64)> {
    if !magick_render_is_due(path) {
        return None;
    }

    imagemagick_render::request(path, room, generation);
    Some((path.to_path_buf(), generation))
}

/// Whether this hover is owed a listing by the PeaZip engine: an archive the engine lists, with an
/// engine installed to list it and nothing listed for this version of it yet.
///
/// It is the same question `magick_render_is_due` is, asked of an engine that reports rather than
/// draws — one that reads an archive, prints what is inside it and exits, which is why there is a
/// process to wait for and nothing to keep. Four things ask it: the layout, which measures a file
/// like this as the wait for a listing; the loader, which answers with it that a hover is still
/// waiting rather than failed; the loop, which asks the engine for the listing only where there is
/// one to ask for; and the layout's own placement question, which decides whether a hover is a
/// wait for something rather than a preview of it (see `page_is_on_the_way`).
fn peazip_render_is_due(path: &Path) -> bool {
    // The file's own bytes first, the name after them, exactly as the image converter's own
    // question asks it: an archive renamed to a name no list holds is still the engine's to list,
    // and one whose bytes are another kind is not a file to start it for (see
    // `peazip_formats::is_engine_archive`).
    peazip_formats::is_engine_archive(path)
        && PreviewType::Peazip.enabled()
        && peazip_render::available_for(path)
        && !peazip_render::refused(path)
        && !peazip_render::listed(path)
}

/// Ask the engine for the listing this hover needs, and answer what is now being waited on.
///
/// The same answer, and for the same reason, as `request_libre_render` and
/// `request_magick_render`: a listing does not exist until the engine has produced it, and asking
/// late is waiting twice. Nothing is waited on here either — the run happens on the engine's own
/// thread — so what comes back is the wait, and the hover is replayed when the engine answers.
fn request_peazip_render(path: &Path, generation: u64) -> Option<(PathBuf, u64)> {
    if !peazip_render_is_due(path) {
        return None;
    }

    peazip_render::request(path, generation);
    Some((path.to_path_buf(), generation))
}

/// Whether this hover is owed a page by the ebook engine: a book the engine reads, with an engine
/// installed to convert it and nothing converted for this version of it yet.
///
/// It is the same question `libre_render_is_due` is, asked of an engine that converts a whole book
/// rather than drawing a page of one — one that reads a file, writes a PDF and exits, which is why
/// there is a process to wait for and nothing to keep. Four things ask it: the layout, which
/// measures a book like this as the wait for a page; the loader, which answers with it that a hover
/// is still waiting rather than failed; the loop, which asks the engine for the page only where
/// there is one to ask for; and the layout's own placement question, which decides whether a hover
/// is a wait for something rather than a preview of it (see `page_is_on_the_way`).
fn calibre_render_is_due(path: &Path) -> bool {
    // The file's own bytes first, the name after them, exactly as the render engine's own question
    // asks it: a book renamed to a name no list holds is still the engine's to convert, and one
    // whose bytes are another kind is not a file to start it for (see
    // `calibre_formats::is_engine_ebook`).
    calibre_formats::is_engine_ebook(path)
        && PreviewType::Calibre.enabled()
        && calibre_render::available()
        && !calibre_render::refused(path)
        && calibre_render::rendered_page(path).is_none()
}

/// Ask the engine for the page this hover needs, and answer what is now being waited on.
///
/// The same answer, and for the same reason, as `request_libre_render`: a page does not exist until
/// the engine has converted the book, and asking late is waiting twice. Nothing is waited on here
/// either — the conversion runs on the engine's own thread — so what comes back is the wait, and
/// the loop watches the folder the page lands in for it.
fn request_calibre_render(path: &Path, generation: u64) -> Option<(PathBuf, u64)> {
    if !calibre_render_is_due(path) {
        return None;
    }

    calibre_render::request(path);
    Some((path.to_path_buf(), generation))
}

/// Ask the first reader of this file's kind that can answer for it, and answer what is now being
/// waited on.
///
/// It is the chain a hover's page is owed by, walked once rather than spelled out: which readers
/// a kind has is `routing::chain`, which of them can answer for this file is
/// `routing::readers_for`, and this asks them in that order. Nothing is waited on in this thread
/// — every request starts work on an engine's own — so what comes back is the wait the loop
/// watches for, and the first reader that asked for one is the reader being waited on.
fn request_engine_render(path: &Path, generation: u64, room: (u32, u32)) -> Option<(PathBuf, u64)> {
    let kind = CONFIG
        .lock()
        .ok()
        .and_then(|config| crate::formats::routing::kind_of(path, &config))?;

    for reader in crate::formats::routing::readers_for(kind, path) {
        let requested = match reader {
            crate::formats::routing::Reader::Office => {
                request_office_render(path, generation, room.0, room.1)
            }
            crate::formats::routing::Reader::LibreOffice => request_libre_render(path, generation),
            crate::formats::routing::Reader::ImageMagick => {
                request_magick_render(path, generation, room)
            }
            crate::formats::routing::Reader::PeaZip => request_peazip_render(path, generation),
            crate::formats::routing::Reader::Calibre => request_calibre_render(path, generation),

            // A reader of this app's own owes the loop no wait, and neither do the two the loop
            // does not ask: a video's player is started where the video is shown, and a document
            // the browser draws is a hover handed over rather than a page waited on.
            crate::formats::routing::Reader::Native
            | crate::formats::routing::Reader::Ffmpeg
            | crate::formats::routing::Reader::WebView2 => None,
        };

        if requested.is_some() {
            return requested;
        }
    }

    None
}

/// Ask the engine that draws this file's page to be up, where the pointer has settled on it.
///
/// It is the question `request_libre_render` asks a moment later, asked a moment early: a page
/// this file is owed, by an engine that is installed, whose kind is switched on and has not
/// turned the file down. Nothing else is warmed — an engine started for a preview the user has
/// switched off is a process on the machine for nothing, and one started for a file that has its
/// page already is a launch nobody was waiting for. The ask itself is the tier's, and it is a
/// no-op for an engine that is up (see `libreoffice_render::warm`).
///
/// Office is not warmed, and that is not an omission: its worker is asked for a *document* to
/// render — the slot it takes is a `RenderRequest`, and the application is created inside the
/// render it makes — so "start the application and hold it with nothing open" is not a request
/// that tier can be handed without taking it apart, which is not what this is for.
///
/// The ebook engine is not warmed either, and there is nothing of it to warm: every book is a
/// conversion of its own and no instance is kept between them (see `calibre_render`).
fn warm_engines_for(path: &Path) {
    if libre_render_is_due(path) {
        libreoffice_render::warm();
    }
}

/// What an engine that answers by writing a page into the app's own folder has said about this
/// file: `Some(true)` where the page has landed, `Some(false)` where the engine has answered that
/// it will not draw the file at all, and `None` where neither engine is the one being waited on or
/// where one of them is and nothing has come back yet.
///
/// Two engines answer this way — the render engine and the ebook engine — and neither sends a
/// message when it is done: what says a page is there is the page, read out of the cache it was
/// kept in (see `libre_render_is_due` and `calibre_render_is_due`). One question for both, so that
/// the wait and the replay that takes the answer up are one code path whichever engine produced it.
fn engine_page_answer(path: &Path) -> Option<bool> {
    // Which engine owes the page is asked of the file, and it is the same question the request side
    // asked before there was anything to ask for: an engine that was never asked has no page for the
    // file, and waiting on it would be waiting for nothing.
    let (drawn, refused) = if calibre_formats::is_engine_ebook(path) {
        (
            calibre_render::rendered_page(path).is_some(),
            calibre_render::refused(path),
        )
    } else if libre_formats::engine_page_kind(path).is_some() {
        (
            libreoffice_render::rendered_page(path).is_some(),
            libreoffice_render::refused(path),
        )
    } else {
        return None;
    };

    match (drawn, refused) {
        (true, _) => Some(true),
        (false, true) => Some(false),
        (false, false) => None,
    }
}

/// Every scale a hover is laid out by, read from the configuration together so that the
/// measure of a file and the render that follows it cannot disagree about the size.
#[derive(Debug, Clone, Copy)]
struct HoverScales {
    /// The share of its own size a picture is drawn at, which is also the scale every
    /// format that is none of the others below keeps.
    picture: PreviewScale,
    /// The share of its own size a video is drawn at.
    video: PreviewScale,
    /// The share of its own size an animated picture is drawn at.
    ///
    /// It is read apart from the picture scale beside it — and read for an animated
    /// file whether it moves or not, since a single-frame GIF or WebP is a picture —
    /// so that what moves is drawn at the size one wants it at rather than at the size
    /// one wants a photograph at.
    animated: PreviewScale,
    /// The share of the display a PDF — the `Ebook` kind — page is drawn at.
    ebook: PreviewScale,
    /// The share of the display a page of the `Document` kind is shown at: what an installed
    /// engine hands back is a page, so the share is of the room the display has rather than of
    /// a size the file asks for, exactly as a PDF page's is.
    document: PreviewScale,
    /// The share of the display a font specimen is drawn at.
    font: PreviewScale,
    /// The share of the display a design document is drawn at.
    ///
    /// What a design document is previewed from is the picture its own format keeps of
    /// the whole thing, so what the share is of is the room the display has rather than
    /// the size that picture happens to be — the same question the document scales above
    /// answer, and a setting of its own because a drawing wants a different share of the
    /// screen from a page or a specimen.
    design: PreviewScale,
    /// The share of the display a vector drawing is replayed over.
    ///
    /// A drawing is not a bitmap: the records are played again at whatever size the box
    /// asks for, so the room the display has is free quality the way a document's is, and
    /// the share is of that room.
    vector: PreviewScale,
}

/// The scale a preview is laid out and rendered with.
///
/// A PDF page is a vector, so the engine draws it at whatever size it is asked
/// for and a larger preview is sharper text rather than an enlarged raster. The
/// room the display has is therefore the page's size, and a configured percentage
/// below `100%` reduces that size rather than being ignored — the page is no
/// longer enlarged by it either, since enlarging a page is what fit-to-screen
/// already does (see `fit_reduced`). What share of that room a page is drawn at is
/// `ebook_scale`'s to say, and it says the whole of it unless it is asked for less.
///
/// Text is the opposite case: it is drawn at a fixed, display-scaled font size,
/// so enlarging it would only stretch the window around text that stays the same
/// size. `100%` is exactly the rule text wants — never enlarged, reduced only
/// when the space beside the cursor cannot hold it — and the text renderer reads
/// the size it is given as "as many lines and columns as fit".
///
/// An SVG is the same case as a page: it is drawn at whatever size it is asked for,
/// so the room the display has is free quality and the whole of it is what a document
/// is drawn at. What it is asked for is a share of that room rather than a share of
/// the size the file asks for, which is the one thing a document and a picture do not
/// agree on: a picture at `50%` is half of its own size, a document at `50%` is half
/// of the screen. The engine's window is the size that comes out of this and its page
/// fills it, so the setting is the document's size and nothing else: see
/// `webview_preview::frame_page`.
///
/// A page of the `Document` kind is the PDF rule again: whether the application that owns the
/// format exported it or an installed render engine drew it, it is drawn at whatever size it is
/// asked for, at the share of the room `document_scale` names — one setting for both, since it
/// is one question about one shape of preview. The one source that is not a page is the bitmap a
/// workbook is answered with where no page can be exported, and it follows the share the way
/// `bitmap_at_display_scale` reads it.
///
/// A font is the same rule once more, at the share `font_scale` names: the specimen is a
/// page of this app's own — the box `font_preview` measures a font at — and the glyphs are
/// sized from the window the engine draws it in, so a share of the room is a share of the
/// type. A file that will not parse as a font is not measured at all, so a hover onto one
/// never reaches this.
///
/// A design document is the document rule rather than the picture's, at the share
/// `design_scale` names: what its preview is made of is the picture the file keeps of the
/// whole document rather than a picture the file *is*, so the room the display has is what
/// the share is of — the same question a page answers, and a setting of its own because a
/// drawing and a page want different shares of it. See `load_design_preview`.
///
/// A vector drawing is the document rule once more, at the share `vector_scale` names: the
/// records are replayed at whatever size they are asked for, so the room the display has is
/// free quality and a share of it is what the setting means — there is nothing to enlarge
/// and nothing to lose by it. See `load_vector_preview`.
///
/// A video keeps the share of its own size `video_scale` names, which is the picture's
/// rule: what a video's preview is, until the player's window is over it, is its first
/// frame — a bitmap measured the way a picture is — so the share is of the file's own
/// size, and it is a setting of its own because a size that suits a photograph is not
/// always the size one wants to watch a file at. The one thing read beside it is whether
/// the probe has answered yet (see `video_probe_due`).
///
/// An animated picture keeps the share of its own size `animated_scale` names, which is
/// the picture's rule once more — its frames are bitmaps — with the file's own content
/// asked which of the two settings it is under: a `.gif`, `.webp` or `.png` that holds
/// more than one frame is animated and follows `animated_scale`, while one that holds a
/// single frame is a still picture and keeps `preview_scale` like any other. Which one
/// it is is the probe below's answer, and it is asked as one question so that the size
/// a hover is placed at and the size its frames are decoded at are the same answer.
///
/// Every other format keeps the picture scale.
fn effective_preview_scale(path: &Path, scales: HoverScales) -> PreviewScale {
    // What the file's own bytes say it is comes first, as it does for the loader that draws
    // it and for the box the layout places it at: a picture under a video's name is laid out
    // at the picture's share, and one under a document's name at the picture's share too.
    // Where the bytes have nothing to say the name decides below, which is every file that
    // is called what it is.
    let content = CONFIG
        .lock()
        .ok()
        .map(|config| crate::formats::content_type::of(path, &config))
        .unwrap_or(crate::formats::content_type::Content::Unknown);

    if let crate::formats::content_type::Content::Kind(kind) = content {
        return scale_of_kind(kind, path, scales);
    }

    // Two questions that are about the run rather than about the file's kind, and both come
    // before it: a page painted to the frame it is given is not scaled within it, and a video
    // whose probe has not answered yet is the spinner rather than a video. Neither can be asked
    // of the kind, which knows nothing about what the run has done so far.
    if page_is_painted(path) {
        // A listing is a page of text painted to the box it is given, whether this app read the
        // archive itself or an engine listed it, so both are the text rule.
        return scale_of_kind(PreviewType::Text, path, scales);
    }

    if video_probe_due(path) {
        // A video that has not been probed yet is a hover that is waiting, and what is on
        // screen for one is the waiting spinner: a wait is placed at the size it is rather
        // than fitted to the display, and what the probe answers is what the replay that
        // follows it is laid out at (see `video_probe_due`).
        return PreviewScale::Percent(100);
    }

    // And the kind decides the rest, asked of the one table every side asks it in: what a hover
    // is measured at is the answer the hook admitted it under and the loader draws it by (see
    // `formats::routing`). A name no list claims is measured as the picture it ends up being
    // decoded as, which is where the loader's own chain sends one.
    let kind = CONFIG
        .lock()
        .ok()
        .and_then(|config| crate::formats::routing::kind_of(path, &config))
        .unwrap_or(PreviewType::Images);

    scale_of_kind(kind, path, scales)
}

/// The share a preview of one kind is drawn at.
///
/// It is one place per kind rather than a share written into each arm of the chain above,
/// because two questions arrive here now — what the name says a file is, and what its bytes
/// say — and both have to come out at the same share for the same kind: a picture is drawn
/// at the picture's share whether it is called `tomcat.png` or `tomcat.mp4`, or the same
/// bytes would be two sizes depending on the name they were left under.
fn scale_of_kind(kind: PreviewType, path: &Path, scales: HoverScales) -> PreviewScale {
    match kind {
        // A page is a vector, so the room the display has is free quality: the setting is
        // the whole of that room unless it asks for less (see `fit_reduced`).
        PreviewType::Ebook => fit_reduced(scales.ebook),

        // Text is drawn at a fixed, display-scaled font size and a listing is painted to
        // the frame it is given, so neither is enlarged or reduced by a setting: the size
        // the box came out at is the size they are drawn at. An archive an engine listed is
        // the second of those: the same page, painted the same way, from a listing that came
        // back from somewhere else.
        PreviewType::Text | PreviewType::Archives | PreviewType::Peazip | PreviewType::Audio => {
            PreviewScale::Percent(100)
        }

        // A page of the `Document` kind is the Ebook rule at the share `document_scale` names,
        // and both halves of the kind answer to it: the page an Office document's own
        // application exported, and the page an installed render engine drew. The raster
        // picture a workbook is answered with where no printer can export a page is the
        // exception: it is only as good as the pixels it holds, so it follows the configured
        // share the way an image does rather than being enlarged to fit. And a document with
        // nothing drawn for it yet is placed at the spinner's own size, since a page on the way
        // has no shape to fit.
        PreviewType::Document => match office_preview::source_kind(path) {
            office_preview::SourceKind::None => PreviewScale::Percent(100),
            source if source.may_be_enlarged() => fit_reduced(scales.document),
            _ => bitmap_at_display_scale(scales.document),
        },

        // A document an engine draws is the other half of that kind, and it answers to the same
        // setting: what the engine hands back is a page, not a picture with a size of its own to
        // be scaled from, so the share is of the display the way a PDF page's is.
        PreviewType::Libre => fit_reduced(scales.document),

        // A picture an engine developed is the picture's rule: what comes back is a PNG,
        // which is a bitmap with a size of its own — the size the engine wrote it at — so
        // the share is of that size rather than of the display. It is asked here rather than
        // left to the arm below so that a `.nef` the content named and one the name named
        // come out at the same size (see `effective_preview_scale`).
        PreviewType::Magick => scales.picture,

        // And a book an engine converted keeps the book rule, at the share `ebook_scale` names:
        // what the engine hands back is a PDF, which is a page rather than a picture with a size of
        // its own to be scaled from — so the share is of the display, the same question a PDF page
        // and a page an engine drew answer. It is asked here rather than left to an arm of its own
        // because the setting is the same one: a user who wants their books smaller wants them
        // smaller whichever reader drew one.
        PreviewType::Calibre => fit_reduced(scales.ebook),

        // A design document is a document for this question rather than a picture: what is
        // previewed is the picture the file keeps of the whole of itself, at whatever size
        // that is, so the share is of the display the way a page's or a specimen's is.
        PreviewType::Design => fit_reduced(scales.design),

        // Both halves of the drawing kind: what a document costs to draw and what a
        // drawing costs to replay are the same question, and the setting is the same one.
        PreviewType::Vector => fit_reduced(scales.vector),

        // And the specimen once more: the glyphs are sized from the window the engine draws
        // it in, so a share of the room is a share of the type.
        PreviewType::Fonts => fit_reduced(scales.font),

        // A video that is still being probed is a wait rather than a video, and a wait is
        // placed at the size it is; one that has been probed keeps the share of its own
        // size `video_scale` names, which is the picture's rule — what a video's preview is
        // until the player's window is over it is its first frame, a bitmap.
        PreviewType::Videos => {
            if video_probe_due(path) {
                PreviewScale::Percent(100)
            } else {
                scales.video
            }
        }

        // A picture keeps the share of its own size, and an animation one of its own. The
        // animated arm is asked last because asking it is the one thing here that reads the
        // file, and a file that has already answered as another kind never pays for it (see
        // `image_is_animated`).
        PreviewType::Images => animated_scale_for(path, scales).unwrap_or(scales.picture),
    }
}

/// The share of a bitmap's own size an animated picture is drawn at, for a file that
/// holds an animation — or `None` for every file this question is not about: one that
/// is not an animated format, one whose animation scale is the same as its picture
/// scale, and one that turned out to hold a single frame after all.
///
/// The two scales being equal is the cheap early exit, and it is the whole reason this
/// can be asked on every hover: where the answer cannot change what is drawn, the file
/// is not read at all, which is the state a fresh install is in because both settings
/// start at `100%`. Only a user who has taken the trouble to give animations a size of
/// their own pays for the probe, and what that probe costs is the file's own structure
/// rather than its pixels (see `image_is_animated`).
fn animated_scale_for(path: &Path, scales: HoverScales) -> Option<PreviewScale> {
    if scales.animated == scales.picture {
        return None;
    }

    image_is_animated(path).then_some(scales.animated)
}

/// Whether a picture file holds a sequence rather than a single frame: a GIF with more than one
/// frame, a WebP with an animation chunk, a PNG with an animation control chunk, or one of the
/// sequences this app has no reader for.
///
/// Nothing is decoded, and the question is asked of the file's own bytes rather than of its
/// name: a still `.gif` and an animated `.png` are both files a name cannot settle, which is the
/// same reason the loader asks the file and not its extension what it is (see
/// `head::PictureNature`). It is asked of the head, which the router has read already by the
/// time a picture reaches this question, so the answer costs a lookup rather than a read.
fn image_is_animated(path: &Path) -> bool {
    crate::formats::head::picture_nature(path).is_some_and(|nature| nature.moves)
}

/// The scale a bitmap is drawn at, for a share of the display it is asked to follow.
///
/// The room the display has is free quality for a source that is drawn at the size it
/// is asked for, and it is not for a bitmap: enlarging one only stretches the pixels it
/// holds, and a worksheet's corner is a few hundred pixels across rather than a
/// screenful. So a share of the display is read for a bitmap as the same share of its
/// own size, and the whole of the display — a fit — as the bitmap at the size it is,
/// which is what `100%` means for a picture. Nothing here is ever enlarged to fill the
/// room.
fn bitmap_at_display_scale(display_scale: PreviewScale) -> PreviewScale {
    match display_scale {
        PreviewScale::Percent(percent) => PreviewScale::Percent(percent),
        PreviewScale::FitToScreen | PreviewScale::FitToScreenReduced(_) => {
            PreviewScale::Percent(100)
        }
    }
}

/// The room the display has, reduced to the configured share of it where the
/// configuration asks for less than the whole of it.
///
/// A source drawn at any size it is asked for — a PDF page, an SVG document, a page
/// Office rendered — is laid out at fit-to-screen, because the display's room is free
/// quality there. A configured percentage at or above `100%` asks for at least
/// that room and is answered with it, so those settings are one setting for such
/// a source; one below `100%` is a size the user picked, and is answered by
/// reducing the fitted size — `50%` halves it — rather than being ignored.
fn fit_reduced(preview_scale: PreviewScale) -> PreviewScale {
    match preview_scale {
        PreviewScale::Percent(percent) if percent < 100 => {
            PreviewScale::FitToScreenReduced(percent)
        }
        _ => PreviewScale::FitToScreen,
    }
}

/// Whether the preview of `path` is a text preview.
///
/// It is one question rather than a chain of exclusions: what kind a file has is the router's
/// answer, and a file is drawn as text exactly when that answer is text. It used to be written
/// out here as "the text lists claim it and no kind asked earlier does", with the kinds listed
/// one by one — and the list had been left short, so a name written into the text list beside a
/// listing engine's or a picture converter's was measured as text and drawn as the other thing
/// (see `formats::routing`).
///
/// What the file's own bytes say comes first, as it does for the loader that draws it and for
/// the box it is painted into: a file whose content is another kind is not drawn as text
/// whatever it is called, and one whose content is text is drawn as text even where the name
/// is a kind the lists would have claimed first.
fn is_text_preview(path: &Path) -> bool {
    let Ok(config) = CONFIG.lock() else {
        return false;
    };

    match crate::formats::content_type::of(path, &config) {
        crate::formats::content_type::Content::Kind(PreviewType::Text) => return true,
        // Another kind, or a format no kind here previews at all: neither is drawn as text,
        // and the second is drawn as nothing.
        crate::formats::content_type::Content::Kind(_)
        | crate::formats::content_type::Content::Foreign => return false,
        crate::formats::content_type::Content::Unknown => {}
    }

    // The kind the hook called it, asked of the same table the hook asked: a name the text
    // lists hold and an earlier list also claims is that earlier kind, and a preview measured
    // as text would be placed as one and drawn as the other. The switch is part of the
    // question, as it is wherever the text lists are asked — a kind turned off in the tray is
    // not drawn at all.
    PreviewType::Text.enabled_in(&config)
        && crate::formats::routing::kind_of(path, &config) == Some(PreviewType::Text)
}

/// Whether a preview of this file is painted into the box it is given rather than scaled within
/// it: a text file, an archive this app read itself, and an archive an engine listed are pages of
/// one kind — painted at a fixed font size, so the box the layout planned for one is the box it
/// draws into, and the frame that comes back is that box rather than a size to be fitted to a
/// space.
///
/// One question, asked in the two places that have to agree about a kind: the share it is drawn
/// at (`effective_preview_scale`) and the box the loader is handed, which is what the window ends
/// up sized to. Asking it in one place is the point — a kind left out of one of them is a preview
/// that is drawn at the planned size and loaded against the free room of the display, which is a
/// page stretched to the screen, and that is exactly what an archive an engine listed was.
///
/// The engine's own question is the third term, asked the way the engine asks it — the file's
/// bytes first and the name after them — so an archive it lists under a name no list holds (a
/// `.cab` renamed to `.dat`) is a page here too.
fn page_is_painted(path: &Path) -> bool {
    is_text_preview(path)
        || archive_formats::is_archive_file(path)
        || peazip_formats::is_engine_archive(path)
        || drawn_as_audio(path)
}

fn effective_frame_delay_ms(media_type: &MediaType, source_delay_ms: u32) -> u32 {
    match media_type {
        MediaType::AnimatedWebP => {
            let fps = current_webp_playback_fps();
            let min_delay_ms = (1000 / fps).max(1);
            source_delay_ms.max(min_delay_ms)
        }
        _ => source_delay_ms,
    }
}

fn checkerboard_color(x: u32, y: u32) -> (u8, u8, u8) {
    if ((x / 16) + (y / 16)).is_multiple_of(2) {
        (224, 224, 224)
    } else {
        (144, 144, 144)
    }
}

/// Compose `bgra` into `out` (BGRA, top-down) with `background` applied to the
/// alpha channel.
///
/// Writes into the caller's buffer so a repaint can target the layered window's
/// DIB directly, and walks it row by row so the per-pixel background position is
/// a row/column counter instead of a division.
///
/// `opaque` is the frame's own answer to whether every one of its pixels has an alpha of 255
/// (see `ImageFrame`). Where it does, the frame *is* the composed surface and the whole of it
/// is one copy: a pixel the backdrop cannot be seen through is the pixel the blend below would
/// have written, in every backdrop kind, because both the premultiply a transparent backdrop
/// asks for and the blend over an opaque one come to the pixel's own bytes where the alpha is
/// 255. That is the difference between a copy and a division per channel per pixel, on every
/// frame of a video or an animation drawn at the size of the display.
fn compose_preview_pixels_into(
    bgra: &[u8],
    width: u32,
    height: u32,
    background: TransparentBackground,
    opaque: bool,
    out: &mut [u8],
) {
    let width = width as usize;
    let expected = width * height as usize * 4;
    if width == 0 || out.len() < expected {
        return;
    }

    // A short source used to end up zero padded; keep that.
    let usable = bgra.len() / 4 * 4;
    if usable < expected {
        out[usable..expected].fill(0);
    }

    if opaque && usable >= expected {
        out[..expected].copy_from_slice(&bgra[..expected]);
        return;
    }

    let row_bytes = width * 4;
    for (y, (src_row, dst_row)) in bgra
        .chunks_exact(row_bytes)
        .zip(out[..expected].chunks_exact_mut(row_bytes))
        .enumerate()
    {
        compose_preview_row(src_row, dst_row, background, 0, y as u32);
    }
}

/// A box of pixels inside a frame: where it sits in the frame, and how large it is.
///
/// It is what the corner spinner is drawn through — the box is copied out of the frame, drawn
/// into, and composed back at the place it came from — so the four numbers travel together
/// rather than as an argument to everything that touches one (see `render_layered_preview_at`).
#[derive(Clone, Copy)]
struct FrameBox {
    left: u32,
    top: u32,
    width: u32,
    height: u32,
}

impl FrameBox {
    /// The area of the box in bytes, at four bytes to the pixel.
    fn bytes(self) -> usize {
        self.width as usize * self.height as usize * 4
    }
}

/// The same for a box of a frame rather than the whole of it: the box is where it sits in the
/// frame, which is both what the checkerboard's squares are placed by and what the destination's
/// rows are offset by.
///
/// It is what the corner spinner is composed back through: what is drawn over a frame is drawn
/// into a copy of the corner it sits in rather than into a copy of the frame, and the box is
/// what that copy is (see `render_layered_preview_at`). A box carries what was drawn over the
/// frame, so it is blended however opaque the frame under it was — it is a few thousand pixels
/// either way.
fn compose_preview_block_into(
    bgra: &[u8],
    area: FrameBox,
    background: TransparentBackground,
    out: &mut [u8],
    out_width: u32,
) {
    let row_bytes = area.width as usize * 4;
    let out_row_bytes = out_width as usize * 4;
    let expected = area.bytes();

    if area.width == 0 || area.height == 0 || bgra.len() < expected {
        return;
    }

    for (y, src_row) in bgra[..expected].chunks_exact(row_bytes).enumerate() {
        let row = area.top as usize + y;
        let start = row * out_row_bytes + area.left as usize * 4;
        let Some(dst_row) = out.get_mut(start..start + row_bytes) else {
            return;
        };

        compose_preview_row(src_row, dst_row, background, area.left, row as u32);
    }
}

/// A box of a frame's pixels, copied out whole: the corner a spinner is drawn into.
fn copy_frame_box_into(frame: &[u8], frame_width: u32, area: FrameBox, out: &mut Vec<u8>) {
    let row_bytes = area.width as usize * 4;
    let frame_row_bytes = frame_width as usize * 4;
    out.clear();
    out.resize(area.bytes(), 0);

    for y in 0..area.height as usize {
        let from = (area.top as usize + y) * frame_row_bytes + area.left as usize * 4;
        let Some(source) = frame.get(from..from + row_bytes) else {
            return;
        };

        out[y * row_bytes..(y + 1) * row_bytes].copy_from_slice(source);
    }
}

fn compose_preview_row(
    src_row: &[u8],
    dst_row: &mut [u8],
    background: TransparentBackground,
    x_offset: u32,
    y: u32,
) {
    match background {
        TransparentBackground::Transparent => {
            for (px, dst) in src_row
                .as_chunks::<4>()
                .0
                .iter()
                .zip(dst_row.as_chunks_mut::<4>().0.iter_mut())
            {
                let b = px[0] as u32;
                let g = px[1] as u32;
                let r = px[2] as u32;
                let a = px[3] as u32;

                dst[0] = ((b * a + 127) / 255) as u8;
                dst[1] = ((g * a + 127) / 255) as u8;
                dst[2] = ((r * a + 127) / 255) as u8;
                dst[3] = a as u8;
            }
        }
        // The background is settled once per row rather than once per pixel: it
        // cannot change inside the loop, and as a per-pixel match it cost a
        // branch on every pixel of every animation frame.
        TransparentBackground::Black => {
            for (px, dst) in src_row
                .as_chunks::<4>()
                .0
                .iter()
                .zip(dst_row.as_chunks_mut::<4>().0.iter_mut())
            {
                blend_pixel_over(px, dst, 0, 0, 0);
            }
        }
        TransparentBackground::White => {
            for (px, dst) in src_row
                .as_chunks::<4>()
                .0
                .iter()
                .zip(dst_row.as_chunks_mut::<4>().0.iter_mut())
            {
                blend_pixel_over(px, dst, 255, 255, 255);
            }
        }
        TransparentBackground::Checkerboard => {
            for (x, (px, dst)) in src_row
                .as_chunks::<4>()
                .0
                .iter()
                .zip(dst_row.as_chunks_mut::<4>().0.iter_mut())
                .enumerate()
            {
                // Squares are placed by where a pixel is in the *frame*, which for a box is
                // not where it is in the box.
                let (r, g, b) = checkerboard_color(x as u32 + x_offset, y);
                blend_pixel_over(px, dst, b as u32, g as u32, r as u32);
            }
        }
    }
}

/// One pixel blended over an opaque background, written opaquely.
#[inline]
fn blend_pixel_over(px: &[u8], dst: &mut [u8], bg_b: u32, bg_g: u32, bg_r: u32) {
    let a = px[3] as u32;
    let inv_a = 255 - a;

    dst[0] = ((px[0] as u32 * a + bg_b * inv_a + 127) / 255) as u8;
    dst[1] = ((px[1] as u32 * a + bg_g * inv_a + 127) / 255) as u8;
    dst[2] = ((px[2] as u32 * a + bg_r * inv_a + 127) / 255) as u8;
    dst[3] = 255;
}

/// The scale a preview of `orig` size takes inside a room, for the scale the
/// configuration asked for.
///
/// One function answers it for every use — the size a preview is given, and the
/// room the position modes choose between — so the scale the layout plans and the
/// scale the renderer draws at cannot come apart.
fn scale_in_room(
    room_width: f32,
    room_height: f32,
    orig_width: f32,
    orig_height: f32,
    preview_scale: PreviewScale,
) -> f32 {
    let fit_scale = (room_width / orig_width).min(room_height / orig_height);

    // A requested percentage is honored when it fits; anything larger than the
    // available area falls back to the fit scale so nothing is ever clipped. A
    // scale that is a reduction of the fit is a share *of* it, and is applied
    // after it.
    let scale = match preview_scale.target_scale() {
        Some(target_scale) => target_scale.min(fit_scale),
        None => fit_scale,
    };

    scale * preview_scale.fit_share()
}

/// Scale media dimensions to the requested preview scale while never exceeding
/// `max_width`/`max_height`, so the preview always stays fully inside the screen.
fn scale_dimensions(
    orig_width: u32,
    orig_height: u32,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> (u32, u32) {
    let scale = scale_in_room(
        max_width as f32,
        max_height as f32,
        orig_width as f32,
        orig_height as f32,
        preview_scale,
    );

    // Rounded so a fit scale lands on the exact available size, then clamped so
    // float error can never push the preview past the screen edge.
    let new_width = (orig_width as f32 * scale)
        .round()
        .clamp(1.0, max_width.max(1) as f32) as u32;
    let new_height = (orig_height as f32 * scale)
        .round()
        .clamp(1.0, max_height.max(1) as f32) as u32;

    (new_width, new_height)
}

/// Animation frames stream continuously, so shrinking keeps the cheap nearest
/// filter while enlarging uses a smoother filter to avoid blocky previews.
fn frame_resize_filter(
    orig_width: u32,
    orig_height: u32,
    target_width: u32,
    target_height: u32,
) -> image::imageops::FilterType {
    if target_width > orig_width || target_height > orig_height {
        image::imageops::FilterType::Triangle
    } else {
        image::imageops::FilterType::Nearest
    }
}

/// Decode a single GIF frame from canvas to an ImageFrame
fn decode_gif_frame_to_image(
    canvas: &[u8],
    gif_width: u32,
    gif_height: u32,
    target_width: u32,
    target_height: u32,
    delay_ms: u32,
) -> Option<ImageFrame> {
    let scaled = if target_width != gif_width || target_height != gif_height {
        let img = image::RgbaImage::from_raw(gif_width, gif_height, canvas.to_vec())?;
        let resized = image::imageops::resize(
            &img,
            target_width,
            target_height,
            frame_resize_filter(gif_width, gif_height, target_width, target_height),
        );
        resized.into_raw()
    } else {
        canvas.to_vec()
    };

    let bgra = rgba_to_bgra(&scaled);

    Some(ImageFrame::new(bgra, target_width, target_height, delay_ms))
}

/// Composite a GIF frame onto the canvas
fn composite_gif_frame(canvas: &mut [u8], frame: &gif::Frame, gif_width: u32, gif_height: u32) {
    let frame_x = frame.left as usize;
    let frame_y = frame.top as usize;
    let frame_w = frame.width as usize;
    let frame_h = frame.height as usize;

    for y in 0..frame_h {
        for x in 0..frame_w {
            let src_idx = (y * frame_w + x) * 4;
            let dst_x = frame_x + x;
            let dst_y = frame_y + y;
            if dst_x < gif_width as usize && dst_y < gif_height as usize {
                let dst_idx = (dst_y * gif_width as usize + dst_x) * 4;
                if src_idx + 3 < frame.buffer.len() {
                    let alpha = frame.buffer[src_idx + 3];
                    if alpha > 0 {
                        canvas[dst_idx] = frame.buffer[src_idx];
                        canvas[dst_idx + 1] = frame.buffer[src_idx + 1];
                        canvas[dst_idx + 2] = frame.buffer[src_idx + 2];
                        canvas[dst_idx + 3] = alpha;
                    }
                }
            }
        }
    }
}

/// Waits while the player is far enough behind that decoding should pause.
/// Returns false when the preview was cancelled while waiting.
fn await_frame_queue_room(shared: &Arc<Mutex<StreamedFrames>>, cancel: &Arc<AtomicBool>) -> bool {
    while !cancel.load(Ordering::Acquire) {
        let queued = shared
            .lock()
            .map(|streamed| streamed.queue.len())
            .unwrap_or(0);
        if queued < ANIMATION_QUEUE_FRAMES {
            return true;
        }
        std::thread::sleep(Duration::from_millis(5));
    }

    false
}

fn load_animated_gif(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    let file = File::open(path).ok()?;
    let mut decoder = DecodeOptions::new();
    decoder.set_color_output(gif::ColorOutput::RGBA);
    let mut decoder = decoder.read_info(BufReader::new(file)).ok()?;

    let (gif_width, gif_height) = (decoder.width() as u32, decoder.height() as u32);
    let (target_width, target_height) =
        scale_dimensions(gif_width, gif_height, max_width, max_height, preview_scale);

    // The canvas every frame is composited into is the size of the GIF itself, so
    // it is the one allocation here that a file chooses, and it is asked for under
    // the same budget as a decoded picture's.
    let canvas_bytes = frame_bytes_within_budget(gif_width, gif_height, 4)?;

    let mut canvas = vec![0u8; canvas_bytes];
    let mut initial_frames = Vec::new();
    let mut initial_bytes: usize = 0;
    let mut reached_end = false;

    while initial_frames.len() < ANIMATION_STARTUP_FRAMES {
        if cancel.load(Ordering::Acquire) {
            return None;
        }

        let frame = match decoder.read_next_frame() {
            Ok(Some(frame)) => frame,
            Ok(None) => {
                reached_end = true;
                break;
            }
            Err(_) => return None,
        };

        composite_gif_frame(&mut canvas, frame, gif_width, gif_height);
        let delay_ms = (frame.delay as u32 * 10).max(MIN_ANIMATION_FRAME_DELAY_MS);
        let img = decode_gif_frame_to_image(
            &canvas,
            gif_width,
            gif_height,
            target_width,
            target_height,
            delay_ms,
        )?;
        initial_bytes = initial_bytes.saturating_add(img.pixels.len());
        if initial_bytes > ANIMATION_RETAINED_BYTES {
            return None;
        }
        initial_frames.push(img);
    }

    if initial_frames.is_empty() || (reached_end && initial_frames.len() <= 1) {
        return None;
    }

    if reached_end {
        return Some(MediaData {
            frames: initial_frames,
            shared_frames: None,
            all_frames_loaded: None,
            current_frame: 0,
            last_frame_time: Instant::now(),
            media_type: MediaType::AnimatedGif,
            stream_cancel: Some(cancel),
            video_process: None,
            loading_start: None,
            text_state: None,
        });
    }

    let shared = Arc::new(Mutex::new(StreamedFrames {
        queue: VecDeque::new(),
        released: false,
        decoded: false,
    }));
    let shared_clone = Arc::clone(&shared);
    let loaded_flag = Arc::new(AtomicBool::new(false));
    let loaded_flag_clone = Arc::clone(&loaded_flag);
    let skip_frames = initial_frames.len();

    let path_clone = path.clone();
    let cancel_clone = Arc::clone(&cancel);
    std::thread::spawn(move || {
        let mut skip = skip_frames;

        loop {
            let file = match File::open(&path_clone) {
                Ok(f) => f,
                Err(_) => break,
            };
            let mut dec = DecodeOptions::new();
            dec.set_color_output(gif::ColorOutput::RGBA);
            let mut dec = match dec.read_info(BufReader::new(file)) {
                Ok(d) => d,
                Err(_) => break,
            };

            let mut canvas = vec![0u8; (gif_width * gif_height * 4) as usize];
            let mut frame_idx = 0usize;
            let mut cancelled = false;

            while let Ok(Some(frame)) = dec.read_next_frame() {
                if cancel_clone.load(Ordering::Acquire)
                    || !await_frame_queue_room(&shared_clone, &cancel_clone)
                {
                    cancelled = true;
                    break;
                }

                composite_gif_frame(&mut canvas, frame, gif_width, gif_height);
                if frame_idx < skip {
                    frame_idx += 1;
                    continue;
                }

                let delay_ms = (frame.delay as u32 * 10).max(MIN_ANIMATION_FRAME_DELAY_MS);
                if let Some(img) = decode_gif_frame_to_image(
                    &canvas,
                    gif_width,
                    gif_height,
                    target_width,
                    target_height,
                    delay_ms,
                ) {
                    if let Ok(mut streamed) = shared_clone.lock() {
                        streamed.queue.push_back(img);
                    }
                }
                frame_idx += 1;
            }

            // Whether the file has to be decoded again is settled under the same
            // lock the player takes before it gives frames back, so that a release
            // and the end of a pass cannot pass each other: whichever happens
            // first, the other sees it. A pass that ends with nothing given back
            // leaves every frame in the player's hands — the file is its own
            // memory from there, and no frame of it is ever decoded again — while a
            // pass that ends after frames were given back is one to run again.
            let replay = match shared_clone.lock() {
                Ok(mut streamed) => {
                    if streamed.released {
                        true
                    } else {
                        streamed.decoded = true;
                        false
                    }
                }
                // A lock that cannot be taken is nothing left to decode for.
                Err(_) => false,
            };

            if cancelled || !replay {
                break;
            }
            skip = 0;
        }

        loaded_flag_clone.store(true, Ordering::Release);
    });

    Some(MediaData {
        frames: initial_frames,
        shared_frames: Some(shared),
        all_frames_loaded: Some(loaded_flag),
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::AnimatedGif,
        stream_cancel: Some(cancel),
        video_process: None,
        loading_start: Some(Instant::now()),
        text_state: None,
    })
}

/// Open an animated PNG frame iterator for the given path.
fn apng_frames(path: &PathBuf) -> Option<image::Frames<'static>> {
    let file = File::open(path).ok()?;
    // Limits are taken at construction rather than set afterwards: a PNG reader
    // holds the ones it was built with, and every frame of the animation is read
    // through it.
    let decoder =
        image::codecs::png::PngDecoder::with_limits(BufReader::new(file), image_decode_limits())
            .ok()?;
    Some(decoder.apng().ok()?.into_frames())
}

/// APNG delays are exact ratios; clamp to the shared floor so a zero-delay
/// animation cannot spin the render loop.
fn apng_frame_delay_ms(frame: &image::Frame) -> u32 {
    let (numerator, denominator) = frame.delay().numer_denom_ms();
    if denominator == 0 {
        return MIN_ANIMATION_FRAME_DELAY_MS;
    }
    (numerator / denominator).max(MIN_ANIMATION_FRAME_DELAY_MS)
}

/// Convert an APNG frame into an ImageFrame. The decoder already composites
/// blend and dispose operations, so every frame arrives as the full canvas.
fn decode_apng_frame_to_image(
    source: &image::RgbaImage,
    target_width: u32,
    target_height: u32,
    delay_ms: u32,
) -> ImageFrame {
    let (source_width, source_height) = source.dimensions();
    let rgba = if target_width != source_width || target_height != source_height {
        image::imageops::resize(
            source,
            target_width,
            target_height,
            frame_resize_filter(source_width, source_height, target_width, target_height),
        )
        .into_raw()
    } else {
        source.as_raw().clone()
    };

    ImageFrame::new(rgba_to_bgra(&rgba), target_width, target_height, delay_ms)
}

fn load_animated_apng(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    let mut frames_iter = apng_frames(path)?;

    let mut initial_frames = Vec::new();
    let mut initial_bytes: usize = 0;
    let mut reached_end = false;
    let mut target_size: Option<(u32, u32)> = None;

    while initial_frames.len() < ANIMATION_STARTUP_FRAMES {
        if cancel.load(Ordering::Acquire) {
            return None;
        }

        let frame = match frames_iter.next() {
            Some(Ok(frame)) => frame,
            Some(Err(_)) => return None,
            None => {
                reached_end = true;
                break;
            }
        };

        let (target_width, target_height) = match target_size {
            Some(size) => size,
            None => {
                let size = scale_dimensions(
                    frame.buffer().width(),
                    frame.buffer().height(),
                    max_width,
                    max_height,
                    preview_scale,
                );
                target_size = Some(size);
                size
            }
        };

        let delay_ms = apng_frame_delay_ms(&frame);
        let img = decode_apng_frame_to_image(frame.buffer(), target_width, target_height, delay_ms);
        initial_bytes = initial_bytes.saturating_add(img.pixels.len());
        if initial_bytes > ANIMATION_RETAINED_BYTES {
            return None;
        }
        initial_frames.push(img);
    }

    if initial_frames.is_empty() || (reached_end && initial_frames.len() <= 1) {
        return None;
    }

    if reached_end {
        return Some(MediaData {
            frames: initial_frames,
            shared_frames: None,
            all_frames_loaded: None,
            current_frame: 0,
            last_frame_time: Instant::now(),
            media_type: MediaType::AnimatedApng,
            stream_cancel: Some(cancel),
            video_process: None,
            loading_start: None,
            text_state: None,
        });
    }

    let shared = Arc::new(Mutex::new(StreamedFrames {
        queue: VecDeque::new(),
        released: false,
        decoded: false,
    }));
    let shared_clone = Arc::clone(&shared);
    let loaded_flag = Arc::new(AtomicBool::new(false));
    let loaded_flag_clone = Arc::clone(&loaded_flag);
    let skip_frames = initial_frames.len();
    let (target_width, target_height) = target_size?;

    let path_clone = path.clone();
    let cancel_clone = Arc::clone(&cancel);
    std::thread::spawn(move || {
        let mut skip = skip_frames;

        loop {
            let frames = match apng_frames(&path_clone) {
                Some(frames) => frames,
                None => break,
            };

            let mut cancelled = false;
            for frame in frames.skip(skip) {
                let frame = match frame {
                    Ok(frame) => frame,
                    Err(_) => break,
                };

                if cancel_clone.load(Ordering::Acquire)
                    || !await_frame_queue_room(&shared_clone, &cancel_clone)
                {
                    cancelled = true;
                    break;
                }

                let delay_ms = apng_frame_delay_ms(&frame);
                let img = decode_apng_frame_to_image(
                    frame.buffer(),
                    target_width,
                    target_height,
                    delay_ms,
                );
                if let Ok(mut streamed) = shared_clone.lock() {
                    streamed.queue.push_back(img);
                }
            }

            // Whether the file has to be decoded again is settled under the same
            // lock the player takes before it gives frames back, so that a release
            // and the end of a pass cannot pass each other: whichever happens
            // first, the other sees it. A pass that ends with nothing given back
            // leaves every frame in the player's hands — the file is its own
            // memory from there, and no frame of it is ever decoded again — while a
            // pass that ends after frames were given back is one to run again.
            let replay = match shared_clone.lock() {
                Ok(mut streamed) => {
                    if streamed.released {
                        true
                    } else {
                        streamed.decoded = true;
                        false
                    }
                }
                // A lock that cannot be taken is nothing left to decode for.
                Err(_) => false,
            };

            if cancelled || !replay {
                break;
            }
            skip = 0;
        }

        loaded_flag_clone.store(true, Ordering::Release);
    });

    Some(MediaData {
        frames: initial_frames,
        shared_frames: Some(shared),
        all_frames_loaded: Some(loaded_flag),
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::AnimatedApng,
        stream_cancel: Some(cancel),
        video_process: None,
        loading_start: Some(Instant::now()),
        text_state: None,
    })
}

fn decode_webp_animation_frame_to_image(
    bgra: &[u8],
    orig_width: u32,
    orig_height: u32,
    target_width: u32,
    target_height: u32,
    delay_ms: u32,
) -> Option<ImageFrame> {
    let expected_bgra = orig_width as usize * orig_height as usize * 4;
    if bgra.len() != expected_bgra {
        return None;
    }

    let pixels = if target_width == orig_width && target_height == orig_height {
        bgra.to_vec()
    } else {
        let mut rgba = Vec::with_capacity(expected_bgra);
        for chunk in bgra.as_chunks::<4>().0.iter() {
            rgba.push(chunk[2]);
            rgba.push(chunk[1]);
            rgba.push(chunk[0]);
            rgba.push(chunk[3]);
        }
        let img = image::RgbaImage::from_raw(orig_width, orig_height, rgba)?;
        let resized = image::imageops::resize(
            &img,
            target_width,
            target_height,
            frame_resize_filter(orig_width, orig_height, target_width, target_height),
        );
        rgba_to_bgra(&resized.into_raw())
    };

    Some(ImageFrame::new(
        pixels,
        target_width,
        target_height,
        delay_ms,
    ))
}

fn load_animated_webp(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    // The file is read whole — libwebp decodes from bytes rather than from a reader
    // of ours — so it is read under the budget like every other file a hover opens.
    let buffer = Arc::new(read_within_budget(path)?);
    let options = webp_animation::DecoderOptions {
        use_threads: true,
        color_mode: webp_animation::ColorMode::Bgra,
    };
    let decoder = webp_animation::Decoder::new_with_options(buffer.as_slice(), options).ok()?;

    // Frames are decoded at the animation's own size, and this reader is libwebp's
    // rather than the `image` crate's, so the budget is asked for by hand where every
    // decoder above is handed it.
    let (orig_width, orig_height) = decoder.dimensions();
    if orig_width == 0 || orig_height == 0 {
        return None;
    }

    frame_bytes_within_budget(orig_width, orig_height, 4)?;

    let (target_width, target_height) = scale_dimensions(
        orig_width,
        orig_height,
        max_width,
        max_height,
        preview_scale,
    );
    if target_width == 0 || target_height == 0 {
        return None;
    }

    let mut initial_frames = Vec::new();
    let mut initial_bytes: usize = 0;
    let mut previous_timestamp = 0i32;
    let mut reached_end = false;
    let mut iterator = decoder.into_iter();

    while initial_frames.len() < ANIMATION_STARTUP_FRAMES {
        if cancel.load(Ordering::Acquire) {
            return None;
        }

        let frame = match iterator.next() {
            Some(frame) => frame,
            None => {
                reached_end = true;
                break;
            }
        };

        let timestamp = frame.timestamp();
        let delay_ms = (timestamp - previous_timestamp).max(0) as u32;
        previous_timestamp = timestamp;

        let img = decode_webp_animation_frame_to_image(
            frame.data(),
            orig_width,
            orig_height,
            target_width,
            target_height,
            delay_ms,
        )?;
        initial_bytes = initial_bytes.saturating_add(img.pixels.len());
        if initial_bytes > ANIMATION_RETAINED_BYTES {
            return None;
        }
        initial_frames.push(img);
    }

    if initial_frames.is_empty() || (reached_end && initial_frames.len() <= 1) {
        return None;
    }

    if reached_end {
        return Some(MediaData {
            frames: initial_frames,
            shared_frames: None,
            all_frames_loaded: None,
            current_frame: 0,
            last_frame_time: Instant::now(),
            media_type: MediaType::AnimatedWebP,
            stream_cancel: Some(cancel),
            video_process: None,
            loading_start: None,
            text_state: None,
        });
    }

    let shared = Arc::new(Mutex::new(StreamedFrames {
        queue: VecDeque::new(),
        released: false,
        decoded: false,
    }));
    let shared_clone = Arc::clone(&shared);
    let loaded_flag = Arc::new(AtomicBool::new(false));
    let loaded_flag_clone = Arc::clone(&loaded_flag);
    let skip_frames = initial_frames.len();

    drop(iterator);
    let buffer_clone = Arc::clone(&buffer);
    let cancel_clone = Arc::clone(&cancel);
    std::thread::spawn(move || {
        let mut skip = skip_frames;

        loop {
            let options = webp_animation::DecoderOptions {
                use_threads: true,
                color_mode: webp_animation::ColorMode::Bgra,
            };
            let decoder =
                match webp_animation::Decoder::new_with_options(buffer_clone.as_slice(), options) {
                    Ok(decoder) => decoder,
                    Err(_) => break,
                };

            let mut previous_timestamp = 0i32;
            let mut cancelled = false;

            for (frame_idx, frame) in decoder.into_iter().enumerate() {
                if cancel_clone.load(Ordering::Acquire)
                    || !await_frame_queue_room(&shared_clone, &cancel_clone)
                {
                    cancelled = true;
                    break;
                }

                let timestamp = frame.timestamp();
                let delay_ms = (timestamp - previous_timestamp).max(0) as u32;
                previous_timestamp = timestamp;

                if frame_idx < skip {
                    continue;
                }

                if let Some(img) = decode_webp_animation_frame_to_image(
                    frame.data(),
                    orig_width,
                    orig_height,
                    target_width,
                    target_height,
                    delay_ms,
                ) {
                    if let Ok(mut streamed) = shared_clone.lock() {
                        streamed.queue.push_back(img);
                    }
                }
            }

            // Whether the file has to be decoded again is settled under the same
            // lock the player takes before it gives frames back, so that a release
            // and the end of a pass cannot pass each other: whichever happens
            // first, the other sees it. A pass that ends with nothing given back
            // leaves every frame in the player's hands — the file is its own
            // memory from there, and no frame of it is ever decoded again — while a
            // pass that ends after frames were given back is one to run again.
            let replay = match shared_clone.lock() {
                Ok(mut streamed) => {
                    if streamed.released {
                        true
                    } else {
                        streamed.decoded = true;
                        false
                    }
                }
                // A lock that cannot be taken is nothing left to decode for.
                Err(_) => false,
            };

            if cancelled || !replay {
                break;
            }
            skip = 0;
        }
        loaded_flag_clone.store(true, Ordering::Release);
    });

    Some(MediaData {
        frames: initial_frames,
        shared_frames: Some(shared),
        all_frames_loaded: Some(loaded_flag),
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::AnimatedWebP,
        stream_cancel: Some(cancel),
        video_process: None,
        loading_start: Some(Instant::now()),
        text_state: None,
    })
}

/// A frame's place in the cache, and when it was last asked for. The stamp is a
/// counter rather than a clock, so the order frames are dropped in cannot be
/// changed by the system clock moving.
struct ImageCacheEntry {
    frame: ImageFrame,
    bytes: usize,
    last_used: u64,
}

/// The file's modification time and length: what says a file is not the one that
/// was decoded last time.
#[derive(Clone, PartialEq, Eq, Hash)]
struct FileVersion {
    modified: Option<SystemTime>,
    len: u64,
}

/// What a held frame is only valid for.
///
/// The file and its version are the obvious part. The pixel size is there as
/// well because a frame is stored decoded *and scaled*: the same photo shown at
/// 100% and at fit-to-screen really is different pixels, so only the size that
/// was asked for can be handed back for it.
#[derive(Clone, PartialEq, Eq, Hash)]
struct ImageCacheKey {
    path: PathBuf,
    version: FileVersion,
    width: u32,
    height: u32,
}

#[derive(Default)]
struct ImageCache {
    entries: HashMap<ImageCacheKey, ImageCacheEntry>,
    bytes: usize,
    tick: u64,
}

static IMAGE_CACHE: Lazy<Mutex<ImageCache>> = Lazy::new(|| Mutex::new(ImageCache::default()));

fn file_version(path: &Path) -> FileVersion {
    match std::fs::metadata(path) {
        Ok(metadata) => FileVersion {
            modified: metadata.modified().ok(),
            len: metadata.len(),
        },
        Err(_) => FileVersion {
            modified: None,
            len: 0,
        },
    }
}

/// The memory the cache may hold, read from the configuration each time rather
/// than captured, so an edit to `image_cache_mb` applies without a restart.
fn image_cache_limit_bytes() -> usize {
    let megabytes = CONFIG
        .lock()
        .map(|config| sanitize_image_cache_mb(config.image_cache_mb))
        .unwrap_or(DEFAULT_IMAGE_CACHE_MB);

    megabytes as usize * 1024 * 1024
}

/// Drop frames, least recently used first, until the cache fits inside `limit`.
///
/// A limit of zero empties it, which is what makes `image_cache_mb = 0` mean
/// "hold nothing" rather than "hold everything until something else is stored".
fn image_cache_trim(cache: &mut ImageCache, limit: usize) {
    while cache.bytes > limit {
        // Bound to its own statement so the borrow of `entries` has ended before
        // the frame is removed.
        let oldest = cache
            .entries
            .iter()
            .min_by_key(|(_, entry)| entry.last_used)
            .map(|(key, _)| (*key).clone());

        let Some(oldest) = oldest else {
            break;
        };

        if let Some(dropped) = cache.entries.remove(&oldest) {
            cache.bytes -= dropped.bytes;
        }
    }
}

/// Trim the image cache to the configured size now, which is what the tray asks for
/// when a smaller size is chosen: what is over the new budget is freed at the moment
/// it is set rather than at the next decode that happens to pass through here.
pub(crate) fn trim_image_cache() {
    let limit = image_cache_limit_bytes();
    if let Ok(mut cache) = IMAGE_CACHE.lock() {
        image_cache_trim(&mut cache, limit);
    }
}

/// The frame held for `key`, if the cache still has it.
fn image_cache_get(key: &ImageCacheKey) -> Option<ImageFrame> {
    let limit = image_cache_limit_bytes();
    let mut cache = IMAGE_CACHE.lock().ok()?;

    image_cache_trim(&mut cache, limit);

    cache.tick += 1;
    let tick = cache.tick;

    let entry = cache.entries.get_mut(key)?;
    entry.last_used = tick;

    Some(entry.frame.clone())
}

/// Hold `frame` for `key`, dropping whatever no longer fits beside it.
fn image_cache_put(key: ImageCacheKey, frame: ImageFrame) {
    let limit = image_cache_limit_bytes();
    let Ok(mut cache) = IMAGE_CACHE.lock() else {
        return;
    };

    image_cache_trim(&mut cache, limit);

    let bytes = frame.pixels.len();
    // A frame larger than the whole budget would evict everything else and still
    // not fit, so it is simply not held.
    if bytes > limit {
        return;
    }

    cache.tick += 1;
    let tick = cache.tick;

    if let Some(previous) = cache.entries.insert(
        key,
        ImageCacheEntry {
            frame,
            bytes,
            last_used: tick,
        },
    ) {
        cache.bytes -= previous.bytes;
    }
    cache.bytes += bytes;

    image_cache_trim(&mut cache, limit);
}

/// An SVG document as `MediaData`: a kind, and no frame at all, because the engine draws
/// it in a window of its own.
///
/// This is what the loader answers with for a document, and the install path reads it as
/// the signal to hand the hover over: nothing of this app's goes on screen for one, so
/// there is no frame to install and no size to place it by.
fn engine_svg_media() -> MediaData {
    MediaData {
        frames: Vec::new(),
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::EngineSvg,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
    }
}

/// A font as `MediaData`: a kind and nothing else, the same shape as a document — the
/// specimen is the engine's to draw, so there is no frame of this app's to install.
fn engine_font_media() -> MediaData {
    MediaData {
        media_type: MediaType::EngineFont,
        ..engine_svg_media()
    }
}

/// A still image as `MediaData`: one frame, nothing streaming.
///
/// The kind arrives with the frame rather than being decided here: a texture is a still
/// picture like any other, and the one thing that makes it a kind of its own is what it is
/// drawn over — which is the loader's to know, since the loader is what held the file.
fn static_image_media(frame: ImageFrame, kind: MediaType) -> MediaData {
    MediaData {
        frames: vec![frame],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: kind,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
    }
}

/// Load a static image (JPG, PNG, BMP, static WebP, etc.)
///
/// Decoding one costs a full decode, a resample and two whole-buffer conversions,
/// and a file list is a place a pointer is swept back and forth over, so a frame
/// that has already been built is handed back rather than built again.
fn load_static_image(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    // The header carries the image's own size, which is what the layout and the
    // resample are computed from, so it is read first: it names the box the frame
    // would be decoded into, and so the key that frame is held under.
    let dimensions = image_dimensions_with_header_check(path);

    // A texture is a picture to everything below this line, and a kind of its own to what
    // draws it: what a `.dds` preview is composited over is the tray's texture backdrop
    // rather than a picture's (see the `Background` submenu and `dds_image`).
    let kind = if dds_image::is_dds_file(path) {
        MediaType::Dds
    } else {
        MediaType::StaticImage
    };

    let cache_key = dimensions.map(|(width, height)| {
        let (target_width, target_height) =
            scale_dimensions(width, height, max_width, max_height, preview_scale);

        ImageCacheKey {
            path: path.to_path_buf(),
            version: file_version(path),
            width: target_width,
            height: target_height,
        }
    });

    if let Some(key) = cache_key.as_ref() {
        if let Some(frame) = image_cache_get(key) {
            return Some(static_image_media(frame, kind));
        }
    }

    // A header that would not report its dimensions is not a reason to refuse the
    // file: the decoder has the last word on whether it is an image at all, and a
    // frame measured this way is simply not held.
    //
    // A picture of a format this app's own decoder does not read is asked of the codec
    // Windows has for it instead, and that one is handed the box the layout planned
    // rather than the file's own size: what it decodes is the preview, and what it
    // hands back is already in the pixel order the frame is composed in, so the
    // resample and the two conversions are not paid for either; see `wic_image`.
    let (pixels, target_width, target_height) = if wic_image::is_codec_file(path) {
        let (width, height) = match cache_key.as_ref() {
            Some(key) => (key.width, key.height),
            // A picture whose size would not be read has no size to take a share of,
            // so what it is decoded into is the box the layout has.
            None => (max_width, max_height),
        };

        // The WebP codec is the one of them that is a Store package rather than
        // something Windows has, so a machine without it — a Windows 10 machine,
        // usually — is answered by libwebp instead, which is in the binary for the
        // picture that moves; see `webp_image`. A `.dds` is the one of them the codec
        // reads a smaller set of than the format holds, so what it has no answer for —
        // the uncompressed formats, BC4 and BC5 — is asked of a decoder of this app's
        // own; see `dds_image`. Both are guarded by the file's own header, so being
        // asked about a picture that is neither costs a header rather than a file.
        let pixels = wic_image::decode(path, width, height)
            .or_else(|| dds_image::decode(path, width, height))
            .or_else(|| webp_image::decode(path, width, height))?;

        (pixels, width, height)
    } else {
        let img = decode_image_with_header_check(path)?;

        // A picture whose samples are light rather than levels — an EXR, a Radiance HDR —
        // is brought into eight bits before anything else is done with it. It is the one
        // thing that has to see the whole of the file's range, and the order is also the
        // cheaper one: what follows is a resample, and resampling one byte to the channel
        // is a quarter of the memory and a fraction of the time of resampling four bytes
        // of float (see `tone_map`).
        let img = match tone_mapped_image(&img) {
            Some(toned) => image::DynamicImage::ImageRgba8(toned),
            None => img,
        };

        let (orig_width, orig_height) = img.dimensions();
        let (width, height) = match cache_key.as_ref() {
            Some(key) => (key.width, key.height),
            None => scale_dimensions(
                orig_width,
                orig_height,
                max_width,
                max_height,
                preview_scale,
            ),
        };

        let resized = if width != orig_width || height != orig_height {
            img.resize_exact(width, height, image::imageops::FilterType::Triangle)
        } else {
            img
        };

        let rgba = resized.to_rgba8();

        (rgba_to_bgra(rgba.as_raw()), width, height)
    };

    let frame = ImageFrame::new(pixels, target_width, target_height, 0);

    if let Some(key) = cache_key {
        image_cache_put(key, frame.clone());
    }

    Some(static_image_media(frame, kind))
}

/// The picture a design document is previewed from.
///
/// Three answers, in one order. Where an engine is installed the document itself is drawn
/// by it — `libreoffice_render` — and what comes back is a page, sharp at whatever size the
/// preview is shown at. Behind it are the pictures these formats keep of their own work,
/// which is what a machine without the engine is answered with, and what a document the
/// engine cannot read is answered with as well: the planar merged picture Photoshop writes
/// at the end of a file, decoded by `psd_image`, and the picture a project container, a
/// CorelDRAW document or a PostScript one holds, read by `project_image`, `cdr_image` and
/// `eps_image`. Which of those is asked is settled by the file's own header rather than by
/// its name, and each hands back the picture decoded into the box the layout planned rather
/// than at the size it is — which for a layered document can be enormous, and is why the
/// box is what is asked for rather than the size.
///
/// Everywhere else a design preview is a frame of this app's own: it is composed like a
/// picture, held in the image cache like one under the size it was made for, drawn over the
/// backdrop the tray keeps for this kind — `design_background`, a setting of its own — and
/// laid out at the share of the display `design_scale` names, which is what
/// `effective_preview_scale` has already read by the time this is called.
fn load_design_preview(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    let (source_width, source_height) = design_dimensions(path)?;
    let (target_width, target_height) = scale_dimensions(
        source_width,
        source_height,
        max_width,
        max_height,
        preview_scale,
    );

    let key = ImageCacheKey {
        path: path.to_path_buf(),
        version: file_version(path),
        width: target_width,
        height: target_height,
    };

    if let Some(frame) = image_cache_get(&key) {
        return Some(static_image_media(frame, MediaType::Design));
    }

    // The page the engine drew comes first, where there is one; the readers below are the
    // fallback for a machine without the engine and for a document it has not drawn — a
    // name this list and the `[libre]` list both hold is drawn by the engine, and this is
    // only reached for one of those where these previews are what is being asked for. What
    // the engine drew comes back at the size the page fitted into the box rather than at
    // the box, so the frame is built from what was drawn.
    let (pixels, width, height) = if let Some(page) = libreoffice_render::rendered_page(path) {
        pdf_preview::render_first_page(&page, target_width, target_height)?
    } else {
        let pixels = if psd_image::is_psd_file(path) {
            psd_image::decode(path, target_width, target_height)
        } else {
            project_image::decode(path, target_width, target_height)
                .or_else(|| eps_image::decode(path, target_width, target_height))
        }?;

        (pixels, target_width, target_height)
    };

    let frame = ImageFrame::new(pixels, width, height, 0);

    image_cache_put(key, frame.clone());

    Some(static_image_media(frame, MediaType::Design))
}

/// The size of the picture a design document is previewed from: the page an installed
/// render engine drew for it, the document's own size for a Photoshop file, and the size of
/// the picture a project container, a CorelDRAW document or a PostScript one holds.
///
/// What is neither of those is either a CorelDRAW document of the older shape — a RIFF
/// container holding a bitmap rather than a zip holding a file — or the encapsulated
/// PostScript a document was saved as before Illustrator wrote PDFs, and each reader
/// answers for what a file is rather than for what it is called.
///
/// A file none of them will answer for reports no size, which is how a design document
/// this app has no reader for comes to show nothing at all rather than a picture of some
/// other format's making.
fn design_dimensions(path: &Path) -> Option<(u32, u32)> {
    // The page the engine drew is measured first, where there is one: what it holds is the
    // document drawn — a page, sharp at whatever size the preview is shown at — rather than
    // a picture of the document that its application kept at some smaller size. Nothing is
    // asked of the engine here: a page it has not drawn is the preview loop's to ask for.
    if let Some(page) = libreoffice_render::rendered_page(path) {
        return pdf_preview::page_dimensions(&page);
    }

    if psd_image::is_psd_file(path) {
        return psd_image::dimensions(path);
    }

    project_image::dimensions(path).or_else(|| eps_image::dimensions(path))
}

/// The page a render engine drew for a document, as a frame of `kind`.
///
/// What the engine hands back is a PDF, and this is the whole of what is asked of it: the
/// page's own size, the box the layout would place that size in, and page 1 rendered into
/// it — drawn at the size it is shown at rather than scaled up from a picture, which is the
/// reason a document an engine drew is worth asking for at all. The pages are cached by the
/// PDF path itself, under the size they were drawn at.
fn load_engine_page(
    page: &Path,
    kind: MediaType,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    let (page_width, page_height) = pdf_preview::page_dimensions(page)?;
    let (target_width, target_height) = scale_dimensions(
        page_width,
        page_height,
        max_width,
        max_height,
        preview_scale,
    );
    let (pixels, width, height) =
        pdf_preview::render_first_page(page, target_width, target_height)?;

    Some(static_image_media(
        ImageFrame::new(pixels, width, height, 0),
        kind,
    ))
}

/// The box a comic is placed at: the first plate's own size, and nothing at all for a container
/// with no plate in it.
///
/// It is the one box of the book kind that is read out of the file rather than out of a page
/// something drew: the plate is inside the container and this side is the reader, so the two
/// answers are the size of that plate and nothing at all — and nothing is a hover that shows no
/// preview and starts no engine, which is what a box of text under a comic's name gets (see
/// `comic_preview`).
///
/// The size is read once per version of the comic, because it costs a walk of the container's
/// own table of contents and a read of the plate it names rather than a header read of a file.
/// Both of those are felt on a comic of any size, so the read is taken off the preview thread
/// and the hover is laid out as the wait for it (see `measured_off_the_tick`).
fn comic_box(path: &Path) -> Option<(u32, u32)> {
    let source = path.to_path_buf();

    measured_off_the_tick(path, MeasureScope::File, move || {
        comic_preview::dimensions(&source)
    })
}

/// The first plate of a comic, drawn into the box the layout measured it for.
///
/// It is `load_pdf_first_page` for a book that is a container rather than a document: what is drawn
/// is a picture that is already in the file, decoded at the size the box asks for and composited
/// over the backdrop the book kind is drawn over. Nothing is waited on and nothing is converted —
/// the plate is read out of the archive, decoded and scaled in one go — so this is the one book of
/// the kind whose preview is made on the side that shows it.
fn load_comic_page(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    let (page_width, page_height) = comic_preview::dimensions(path)?;

    let (target_width, target_height) = scale_dimensions(
        page_width,
        page_height,
        max_width,
        max_height,
        preview_scale,
    );
    let pixels = comic_preview::decode(path, target_width, target_height)?;

    Some(static_image_media(
        ImageFrame::new(pixels, target_width, target_height, 0),
        MediaType::Comic,
    ))
}

/// The page a converted book is drawn from, drawn into the box the layout measured it for.
///
/// It is `load_engine_page` for the one engine whose page is a whole book rather than a page of a
/// document: what is drawn is the first page of that book which says anything about it, and the box
/// is that page's own size, so a book whose first page is a cover of one colour is previewed from
/// the page behind it rather than as the colour (see `pdf_preview::book_page`).
fn load_book_page(
    page: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    let book = pdf_preview::book_page(page)?;
    let (page_width, page_height) = book.size;

    let (target_width, target_height) = scale_dimensions(
        page_width,
        page_height,
        max_width,
        max_height,
        preview_scale,
    );
    let (pixels, width, height) = pdf_preview::render_book_page(page, target_width, target_height)?;

    Some(static_image_media(
        ImageFrame::new(pixels, width, height, 0),
        MediaType::Calibre,
    ))
}

/// The page the render engine has drawn for an Office document — the fallback for a document
/// whose own application is not installed, and the page itself where the tray has asked the
/// engine for every Office document. Shown as the Office document it is either way.
///
/// Nothing is converted here, and nothing is waited on: a document the engine has not drawn
/// yet is answered with nothing, which is the wait the hover is already in — the loop has
/// asked the engine for the page, and the hover is replayed when it lands (see
/// `libre_render_is_due`). The page that *is* there is read at the share the layout measured
/// it for, and that share is the Office kind's: the file is what it is whichever engine drew
/// it (see `effective_preview_scale`).
///
/// A page the engine drew while it was the one being asked is not read once the choice is the
/// application's: a page is what the engine that is drawn by is asked for, and a document the
/// application is drawing is not shown the other engine's page until it is ready (see
/// `office_formats::page_engine`).
fn load_engine_page_for_office(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    if office_formats::page_engine(path) != Some(OfficeEngine::LibreOffice) {
        return None;
    }

    let page = libreoffice_render::rendered_page(path)?;
    load_engine_page(
        &page,
        MediaType::Office,
        max_width,
        max_height,
        preview_scale,
    )
}

/// The picture the ImageMagick engine developed for a file, drawn as the picture it is.
///
/// Nothing is converted here, and nothing is waited on. Three things answer a hover in the
/// order they are worth asking:
///
/// * the frame this app has already built for this file at this box, which is the image cache
///   every other picture is kept in — a hit costs a lookup and no engine at all, which is what
///   a pointer swept back and forth over a folder of raws is answered with;
/// * the picture the engine developed, which is what the frame is built from: the one held in
///   the engine's hand for the hover that asked, or the page it wrote down for the hovers after
///   it — decoded from the bytes the engine wrote the picture as rather than from the source
///   file, resampled into the box the layout planned, and held in that same cache under the
///   file, its version and that box;
/// * and nothing at all, which is the wait the hover is already in — the loop has asked the
///   engine for the picture, and the hover is replayed when the answer lands (see
///   `magick_render_is_due`).
///
/// What the frame is composited over is `image_background` and what share of its own size it
/// is drawn at is the picture's, because that is what it is: a picture of this app's, in the
/// format a frame is composed in, held like any other.
fn load_magick_picture(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    cancel: &Arc<AtomicBool>,
) -> Option<MediaData> {
    // The size the engine developed this file at — which is the picture's own size rather
    // than the file's, since what the engine writes is a picture fitted into the room — and
    // what the frame is keyed by, together with the box it is drawn in.
    let cache_key = match imagemagick_render::dimensions(path) {
        Some((width, height)) => {
            let (target_width, target_height) =
                scale_dimensions(width, height, max_width, max_height, preview_scale);

            Some(ImageCacheKey {
                path: path.to_path_buf(),
                version: file_version(path),
                width: target_width,
                height: target_height,
            })
        }
        // A file the engine has developed nothing for is a file with no size to key a frame
        // by: what is coming is the engine's answer, and the frame it is drawn as is built
        // from that answer rather than placed in the cache under a size nothing knows.
        None => None,
    };

    if let Some(key) = cache_key.as_ref() {
        if let Some(frame) = image_cache_get(key) {
            return Some(static_image_media(frame, MediaType::Magick));
        }
    }

    // The picture the engine developed: the one in its hand, which is what a hover that asked
    // for it is waiting for — a hover that has moved on leaves it where it is, since what it is
    // waiting for is its own replay — or the page it wrote for the hovers after that one, which
    // is every hover since a restart.
    let developed = match imagemagick_render::take_developed(path) {
        Some(developed) => developed,
        None => imagemagick_render::read_page(path)?,
    };
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    let (target_width, target_height) = scale_dimensions(
        developed.width,
        developed.height,
        max_width,
        max_height,
        preview_scale,
    );
    // A page that is there but does not decode is not a picture: it is given up, so that the
    // file behind it is developed again rather than read into the same nothing on every hover.
    let Some(image) = decode_png(&developed.png) else {
        imagemagick_render::forget_page(path);
        return None;
    };
    let (orig_width, orig_height) = image.dimensions();
    let resized = if target_width != orig_width || target_height != orig_height {
        image.resize_exact(
            target_width,
            target_height,
            image::imageops::FilterType::Triangle,
        )
    } else {
        image
    };

    let rgba = resized.to_rgba8();
    let frame = ImageFrame::new(rgba_to_bgra(rgba.as_raw()), target_width, target_height, 0);

    if let Some(key) = cache_key {
        image_cache_put(key, frame.clone());
    }

    Some(static_image_media(frame, MediaType::Magick))
}

/// The picture the engine wrote, decoded from the bytes it wrote it as — under the budget
/// every other decode of this app is answered under, and with the format asked for rather
/// than taken on trust.
fn decode_png(png: &[u8]) -> Option<image::DynamicImage> {
    let mut reader = image::ImageReader::new(std::io::Cursor::new(png))
        .with_guessed_format()
        .ok()?;
    reader.limits(image_decode_limits());

    reader.decode().ok()
}

/// The drawing a vector file is previewed from.
///
/// Two readers answer for these files and both hand back a frame of the same kind: the
/// records an `.eps` carries, and the records a `.wmf` or an `.emf` is. Both ask the file
/// itself what it is rather than trusting its name, so which is asked first is only a
/// question of which is cheaper to turn down — the metafile reader is asked first for the
/// two names it is written as, and the encapsulated PostScript reader first for everything
/// else, since either of them refuses a file that is not its own after a header.
///
/// What comes back is the drawing replayed at the box the layout planned rather than a
/// picture resampled into it, which is what makes a preview of one sharp at any size the
/// display has. A file carrying a picture instead — a TIFF preview, which is what some
/// writers leave in an `.eps` — is resampled the way every other picture is.
fn load_vector_preview(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    let (source_width, source_height) = vector_dimensions(path)?;
    let (target_width, target_height) = scale_dimensions(
        source_width,
        source_height,
        max_width,
        max_height,
        preview_scale,
    );

    let key = ImageCacheKey {
        path: path.to_path_buf(),
        version: file_version(path),
        width: target_width,
        height: target_height,
    };

    if let Some(frame) = image_cache_get(&key) {
        return Some(static_image_media(frame, MediaType::Vector));
    }

    let pixels = if metafile_image::is_metafile_name(path) {
        metafile_image::decode(path, target_width, target_height)
            .or_else(|| eps_image::decode(path, target_width, target_height))
    } else {
        eps_image::decode(path, target_width, target_height)
            .or_else(|| metafile_image::decode(path, target_width, target_height))
    }?;

    let frame = ImageFrame::new(pixels, target_width, target_height, 0);

    image_cache_put(key, frame.clone());

    Some(static_image_media(frame, MediaType::Vector))
}

/// The size a drawing asks to be shown at: what the preview inside an `.eps` is of, or what
/// a metafile's own header declares.
fn vector_dimensions(path: &Path) -> Option<(u32, u32)> {
    if metafile_image::is_metafile_name(path) {
        return metafile_image::dimensions(path).or_else(|| eps_image::dimensions(path));
    }

    eps_image::dimensions(path).or_else(|| metafile_image::dimensions(path))
}

/// Render the first page of a PDF through the PDF engine built into Windows.
fn load_pdf_first_page(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    let (page_width, page_height) = pdf_preview::page_dimensions(path).unwrap_or((
        pdf_preview::DEFAULT_PAGE_WIDTH,
        pdf_preview::DEFAULT_PAGE_HEIGHT,
    ));
    let (target_width, target_height) = scale_dimensions(
        page_width,
        page_height,
        max_width,
        max_height,
        preview_scale,
    );

    let (pixels, width, height) =
        pdf_preview::render_first_page(path, target_width, target_height)?;

    let frame = ImageFrame::new(pixels, width, height, 0);

    Some(MediaData {
        frames: vec![frame],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::Pdf,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
    })
}

/// Render a page of an Office document into the box the layout planned.
///
/// One source, and it is one this side reads rather than produces: the page the render tier
/// is holding in memory for the document — drawn by Office itself, or by the render engine
/// beside it where the document's own application is not installed, and shown under this
/// kind either way. A document with no page yet is answered with nothing, which is the wait
/// the hover is in, and the page is asked for by the loop the moment the hover is up (see
/// `request_office_render` and `libre_render_is_due`). The source's own aspect ratio is
/// preserved inside the box, so a page that is not the shape the layout assumed is
/// letterboxed instead of stretched.
fn load_office_preview(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    cancel: &Arc<AtomicBool>,
) -> Option<MediaData> {
    let (source_width, source_height) =
        office_preview::measure(path).unwrap_or_else(|| office_formats::default_page_size(path));
    let (target_width, target_height) = scale_dimensions(
        source_width,
        source_height,
        max_width,
        max_height,
        preview_scale,
    );

    let (pixels, width, height) =
        office_preview::render(path, target_width, target_height, Some(cancel))?;

    let frame = ImageFrame::new(pixels, width, height, 0);

    Some(MediaData {
        frames: vec![frame],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::Office,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
    })
}

/// Render a text file into the box the layout planned for it.
///
/// Unlike an image, text is not scaled to the box: the font is a fixed,
/// display-scaled size, and the box decides how many lines and columns are shown.
/// That is why the caller hands over the planned preview size rather than the
/// free space around the cursor — the two are the same thing to this renderer,
/// and using the planned size keeps the painted frame and the window in step.
fn load_text_preview(
    path: &Path,
    width: u32,
    height: u32,
    dpi: u32,
    options: TextPreviewOptions,
) -> Option<MediaData> {
    let frame = text_preview::render_scrolled(path, 0, width, height, dpi, options, None)?;

    // Any text preview in full mode keeps its state, scrollable or not: the
    // pointer rests on it, its text can be selected, and a document that happens to
    // fit is simply one that cannot be scrolled.
    let state = options.full_mode.then(|| TextPreviewState {
        path: path.to_path_buf(),
        options,
        dpi,
        width: frame.width,
        height: frame.height,
        first_line: frame.first_line,
        visible_lines: frame.visible_lines,
        scrollable_lines: frame.scrollable_lines,
        scrollbar: frame.scrollbar,
        dragging: false,
        lines: frame.lines,
        selection: None,
        selecting: false,
    });

    let frame = ImageFrame::new(frame.pixels, frame.width, frame.height, 0);

    Some(MediaData {
        frames: vec![frame],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::Text,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: state,
    })
}

/// Load an archive's contents as a page of its own, the way a text preview is
/// loaded: measured first, then painted into exactly the box the layout planned.
///
/// Which kind the page comes out as is the caller's to say, because the same page is drawn for
/// two of them: an archive this app read itself is shown under `Archives`, and one an engine
/// listed under `Peazip` — the same reader of the same listing and the same painted frame, with
/// only the gate over the preview on screen differing between them.
fn load_archive_preview(
    path: &Path,
    width: u32,
    height: u32,
    dpi: u32,
    options: ArchivePreviewOptions,
    media_type: MediaType,
    cancel: &AtomicBool,
) -> Option<MediaData> {
    let (pixels, width, height) = archive_preview::render(path, width, height, dpi, options)?;
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    Some(MediaData {
        frames: vec![ImageFrame::new(pixels, width, height, 0)],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
    })
}

/// Extract video thumbnail using ffmpeg and create frames for preview
fn load_video_thumbnail(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    // Which engine plays this file is settled here, once, and everything below follows
    // from it: FFmpeg's player when it is installed, and the media engine Windows has when
    // it is not.
    let native = codecs::plays_video_natively();

    let geometry = match probe_video_geometry(path) {
        ProbedGeometry::Measured(geometry) => geometry,
        // No geometry, and the engine that would play the file cannot open it either: a
        // file this machine has no reader for is answered with no preview rather than with
        // a 16:9 box nothing would be drawn into.
        ProbedGeometry::Unmeasurable if native => return None,
        // FFmpeg is the one that would play it, so the box the layout uses is the one it
        // has always used for a file ffprobe could not measure.
        ProbedGeometry::Unmeasurable => VideoGeometry {
            width: 1920,
            height: 1080,
            crop: None,
        },
    };

    let (target_width, target_height) = scale_dimensions(
        geometry.width,
        geometry.height,
        max_width,
        max_height,
        preview_scale,
    );

    // Create a placeholder frame (dark gray) while video plays
    let placeholder_pixels = vec![40u8; (target_width * target_height * 4) as usize];

    let frame = ImageFrame::new(placeholder_pixels, target_width, target_height, 0);

    Some(MediaData {
        frames: vec![frame],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: if native {
            MediaType::NativeVideo
        } else {
            MediaType::Video
        },
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
    })
}

const VIDEO_CROPDETECT_LIMIT: &str = "24";
const VIDEO_CROPDETECT_ROUND: &str = "16";
const VIDEO_CROPDETECT_FRAMES: &str = "48";
const VIDEO_CROP_MAX_AXIS_TRIM_RATIO: f32 = 0.10;
const VIDEO_CROP_MAX_ASYMMETRY_PX: i32 = 12;

/// How long either of a video's two probes is given before it is killed and the file is
/// answered as one that could not be measured.
///
/// The wait this bounds is the one wait in the preview that nothing else ends: the hover is
/// in a `PendingLoad` that is not waiting on an engine, so the cap a page is waited for under
/// is not standing behind it, and a file no probe ever answers for is a spinner that stands
/// at the pointer until the pointer moves. Ten seconds is what a player's own start is given
/// (`VIDEO_START_WAIT_SECS`) and the same order as the work — forty-eight decoded frames of a
/// 4K file is the slow case — and what a file slower than this is answered with is a video
/// placed without its crop, which is what a file `ffprobe` could not read has always been
/// given. If a file of the user's turns out to be slower than this, it is one constant.
const VIDEO_PROBE_TIMEOUT_SECS: u64 = 10;

/// Wait for one of a probe's children, giving up after `timeout`, and answer with what it
/// wrote or with nothing at all.
///
/// A child that is still running at the deadline is killed *before* anything is read of it,
/// because what is being stopped is not the answer but the work: a cropdetect pass over a
/// file whose frames it cannot keep up with runs until the file ends, and a hover that has
/// given up on it is a hover that must not leave it running. The wait that follows the kill
/// is what reaps it, and the output a killed child leaves behind is a partial line — which
/// is the answer arm both callers already have for a probe that answered nothing.
///
/// A child that has finished needs no kill and no second wait: `wait_with_output` drains the
/// pipes and is what closes them, and the handle being signalled is what says there is
/// something to drain. Nothing here is recorded to strike off — the probes are adopted and
/// never written down, a process of a few dozen milliseconds having no file to leave.
fn wait_bounded(mut child: Child, timeout: Duration) -> Option<Output> {
    let handle = HANDLE(child.as_raw_handle());
    let waited = unsafe { WaitForSingleObject(handle, timeout.as_millis() as u32) };

    if waited == WAIT_TIMEOUT {
        let _ = child.kill();
        let _ = child.wait();
        return None;
    }

    child.wait_with_output().ok()
}

/// Get video dimensions using ffprobe
fn get_video_dimensions(path: &PathBuf) -> Option<(u32, u32)> {
    // Spawned rather than run through `Command::output`, which is these two calls
    // under one name, so that the probe is in the job before it is waited on: a
    // probe left behind by a crash would otherwise go on reading a file that nobody
    // is waiting for. The wait that follows is bounded rather than the plain one, so
    // that a file this probe cannot finish with is answered rather than waited on
    // (see `wait_bounded`).
    let child = Command::new("ffprobe")
        .args([
            "-v",
            "error",
            "-err_detect",
            "ignore_err",
            "-fflags",
            "+genpts+discardcorrupt+igndts",
            "-select_streams",
            "v:0",
            "-show_entries",
            "stream=width,height",
            "-of",
            "csv=s=x:p=0",
        ])
        .arg(path)
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .creation_flags(engine_processes::CREATE_NO_WINDOW) // Hide the console window
        .spawn()
        .ok()?;

    engine_processes::adopt(child.id());

    let output = wait_bounded(child, Duration::from_secs(VIDEO_PROBE_TIMEOUT_SECS))?;

    let output_str = String::from_utf8_lossy(&output.stdout);
    let mut parts = output_str.trim().split('x').filter(|part| !part.is_empty());
    let width = parts.next()?.parse().ok()?;
    let height = parts.next()?.parse().ok()?;
    Some((width, height))
}

fn parse_cropdetect_line(line: &str) -> Option<VideoCrop> {
    let idx = line.rfind("crop=")?;
    let token = line[idx + 5..]
        .split_whitespace()
        .next()
        .unwrap_or_default();
    let mut parts = token.split(':');
    let width: u32 = parts.next()?.parse().ok()?;
    let height: u32 = parts.next()?.parse().ok()?;
    let x: u32 = parts.next()?.parse().ok()?;
    let y: u32 = parts.next()?.parse().ok()?;
    if parts.next().is_some() {
        return None;
    }

    Some(VideoCrop {
        width,
        height,
        x,
        y,
    })
}

fn validate_detected_crop(crop: VideoCrop, src_w: u32, src_h: u32) -> bool {
    if crop.width == 0 || crop.height == 0 || crop.width > src_w || crop.height > src_h {
        return false;
    }

    let right = crop.x.saturating_add(crop.width);
    let bottom = crop.y.saturating_add(crop.height);
    if right > src_w || bottom > src_h {
        return false;
    }

    let trim_left = crop.x as i32;
    let trim_top = crop.y as i32;
    let trim_right = src_w.saturating_sub(right) as i32;
    let trim_bottom = src_h.saturating_sub(bottom) as i32;
    let trim_x = src_w.saturating_sub(crop.width);
    let trim_y = src_h.saturating_sub(crop.height);

    if trim_x == 0 && trim_y == 0 {
        return false;
    }

    let trim_x_ratio = trim_x as f32 / src_w as f32;
    let trim_y_ratio = trim_y as f32 / src_h as f32;
    if trim_x_ratio > VIDEO_CROP_MAX_AXIS_TRIM_RATIO
        || trim_y_ratio > VIDEO_CROP_MAX_AXIS_TRIM_RATIO
    {
        return false;
    }

    (trim_left - trim_right).abs() <= VIDEO_CROP_MAX_ASYMMETRY_PX
        && (trim_top - trim_bottom).abs() <= VIDEO_CROP_MAX_ASYMMETRY_PX
}

/// Every crop rectangle ffmpeg's detector reported for the file, with the number
/// of frames that reported it. The source dimensions are not needed to collect
/// them, which is what lets this run alongside the probe that reads them.
fn collect_video_crop_candidates(path: &PathBuf) -> HashMap<(u32, u32, u32, u32), u32> {
    let filter = format!(
        "cropdetect={}:{}:0",
        VIDEO_CROPDETECT_LIMIT, VIDEO_CROPDETECT_ROUND
    );

    // Spawned rather than run through `Command::output` for the same reason the
    // ffprobe next to it is: a probe that is in the job is one a crash cannot leave
    // reading a file with nobody waiting for it.
    let child = match Command::new("ffmpeg")
        .args([
            "-v",
            "info",
            "-err_detect",
            "ignore_err",
            "-fflags",
            "+genpts+discardcorrupt+igndts",
            "-i",
        ])
        .arg(path)
        .args([
            "-frames:v",
            VIDEO_CROPDETECT_FRAMES,
            "-vf",
            &filter,
            "-f",
            "null",
            "-",
        ])
        .stdout(Stdio::null())
        .stderr(Stdio::piped())
        .creation_flags(engine_processes::CREATE_NO_WINDOW)
        .spawn()
    {
        Ok(child) => child,
        Err(_) => return HashMap::new(),
    };

    engine_processes::adopt(child.id());

    let output = match wait_bounded(child, Duration::from_secs(VIDEO_PROBE_TIMEOUT_SECS)) {
        Some(output) => output,
        None => return HashMap::new(),
    };

    let stderr = String::from_utf8_lossy(&output.stderr);
    let mut counts: HashMap<(u32, u32, u32, u32), u32> = HashMap::new();
    for line in stderr.lines() {
        if let Some(crop) = parse_cropdetect_line(line) {
            *counts
                .entry((crop.width, crop.height, crop.x, crop.y))
                .or_insert(0) += 1;
        }
    }

    counts
}

/// The crop the detector was most sure of, among those that hold up against the
/// source dimensions.
fn best_valid_crop(
    counts: HashMap<(u32, u32, u32, u32), u32>,
    src_w: u32,
    src_h: u32,
) -> Option<VideoCrop> {
    let mut best: Option<(VideoCrop, u32)> = None;
    for ((width, height, x, y), count) in counts {
        let crop = VideoCrop {
            width,
            height,
            x,
            y,
        };
        if !validate_detected_crop(crop, src_w, src_h) {
            continue;
        }

        match best {
            Some((existing, existing_count)) => {
                let existing_area = (existing.width as u64) * (existing.height as u64);
                let candidate_area = (crop.width as u64) * (crop.height as u64);
                if count > existing_count
                    || (count == existing_count && candidate_area > existing_area)
                {
                    best = Some((crop, count));
                }
            }
            None => best = Some((crop, count)),
        }
    }

    best.map(|(crop, _)| crop)
}

/// The geometry a video has already been probed for, when it has been probed at all: the
/// answer the probe gave, read from the cache and nothing else.
///
/// This is the lookup the preview thread is allowed to make — it is a mutex and a hash,
/// not two external processes — and it is what tells a hover whether its file has been
/// measured yet (see `video_probe_due`).
fn cached_video_geometry(path: &Path) -> Option<ProbedGeometry> {
    let key = VideoGeometryKey {
        path: path.to_path_buf(),
        version: file_version(path),
    };

    video_geometry_cache().get(&key).copied()
}

/// Probe a video's geometry, from the cache when the file and its version have been
/// probed before.
///
/// This is the one caller that may run the two external processes, so it is only ever
/// called from a thread whose waiting does not matter: the load worker, and the probe a
/// hover is waiting on (see `video_probe_due` in the preview loop). What it answers is
/// held — the failure included — so the next hover of the file is a lookup.
fn probe_video_geometry(path: &PathBuf) -> ProbedGeometry {
    let key = VideoGeometryKey {
        path: path.clone(),
        version: file_version(path),
    };

    if let Some(cached) = video_geometry_cache().get(&key) {
        return *cached;
    }

    // Reading the dimensions and detecting the crop are two external processes,
    // and the detector is the one that decodes frames: neither needs the other's
    // answer until the crop is validated, so they run at once and the hover waits
    // for the slower one rather than for both in turn.
    let (dimensions, candidates) = std::thread::scope(|scope| {
        let dimensions = scope.spawn(|| get_video_dimensions(path));
        let candidates = scope.spawn(|| collect_video_crop_candidates(path));

        (
            dimensions.join().unwrap_or(None),
            candidates.join().unwrap_or_default(),
        )
    });

    // A file FFmpeg is not there for — or one its own probe could not read — is asked of
    // the media engine Windows has, which is also the engine that would play it. That is
    // the whole of the fallback's geometry: there is no crop to detect, because cropdetect
    // is an FFmpeg filter and the engine is handed the frame as the file holds it.
    let Some((src_w, src_h)) = dimensions.or_else(|| video_player::dimensions(path)) else {
        // No picture in the file at all — which leaves two answers, and the one that matters
        // is asked first. A container of a video's name whose streams hold a sound and no
        // picture is a song: the sound is probed for, and a machine that can play it is
        // answered as one from here on, because the router asks this verdict before it asks
        // the video list — so the hover that is replayed for this probe is laid out as the
        // card it is rather than dropped as a video with no shape (see
        // `audio_formats::probed_audio_only`). What is left is a file nothing here can read,
        // which is the answer this arm always gave it.
        let playable = match audio_track::probed(path) {
            Probed::Track(_) => true,
            Probed::Nothing => false,
            Probed::NotAsked => {
                let probed = probe_audio_track(path);
                audio_track::remember(
                    path,
                    match &probed {
                        Some(track) => Probed::Track(track.clone()),
                        None => Probed::Nothing,
                    },
                );
                probed.is_some()
            }
        };

        if playable {
            audio_formats::remember_audio_only(path);
        }

        let mut cache = video_geometry_cache();
        if !cache.contains_key(&key) && cache.len() >= VIDEO_GEOMETRY_CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.insert(key, ProbedGeometry::Unmeasurable);

        return ProbedGeometry::Unmeasurable;
    };
    let crop = best_valid_crop(candidates, src_w, src_h);

    let geometry = if let Some(crop) = crop {
        VideoGeometry {
            width: crop.width,
            height: crop.height,
            crop: Some(crop),
        }
    } else {
        VideoGeometry {
            width: src_w,
            height: src_h,
            crop: None,
        }
    };

    let mut cache = video_geometry_cache();
    if !cache.contains_key(&key) && cache.len() >= VIDEO_GEOMETRY_CACHE_MAX_ENTRIES {
        cache.clear();
    }
    cache.insert(key, ProbedGeometry::Measured(geometry));

    ProbedGeometry::Measured(geometry)
}

/// Data passed to the EnumWindows callback to find ffplay window
struct EnumWindowsData {
    target_pid: u32,
    found_hwnd: HWND,
    best_area: i64,
}

/// Callback for EnumWindows to find a window belonging to a specific process
unsafe extern "system" fn enum_windows_callback(
    hwnd: HWND,
    lparam: LPARAM,
) -> windows::Win32::Foundation::BOOL {
    let data = &mut *(lparam.0 as *mut EnumWindowsData);
    let mut window_pid: u32 = 0;
    GetWindowThreadProcessId(hwnd, Some(&mut window_pid));

    if window_pid != data.target_pid {
        return windows::Win32::Foundation::BOOL(1);
    }

    // Prefer visible, top-level windows (skip hidden and owned/popups behind owners)
    if !IsWindowVisible(hwnd).as_bool() {
        return windows::Win32::Foundation::BOOL(1);
    }

    if let Ok(owner) = GetWindow(hwnd, GW_OWNER) {
        if !owner.is_invalid() {
            return windows::Win32::Foundation::BOOL(1);
        }
    }

    let mut rect = RECT::default();
    if GetWindowRect(hwnd, &mut rect).is_err() {
        return windows::Win32::Foundation::BOOL(1);
    }

    let width = (rect.right - rect.left).max(0) as i64;
    let height = (rect.bottom - rect.top).max(0) as i64;
    let area = width * height;
    if area <= 0 {
        return windows::Win32::Foundation::BOOL(1);
    }

    // Keep the largest candidate; this is typically the real ffplay output window.
    if area > data.best_area {
        data.best_area = area;
        data.found_hwnd = hwnd;
    }

    windows::Win32::Foundation::BOOL(1)
}

/// Style and raise a known ffplay window.
unsafe fn apply_noactivate_to_hwnd(hwnd: HWND) -> bool {
    // Store the video window HWND for cursor-over-preview detection
    VIDEO_HWND.store(hwnd.0 as isize, Ordering::SeqCst);

    // Add WS_EX_NOACTIVATE and WS_EX_TOPMOST to its extended style
    let current_style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
    let new_style = current_style
        | WS_EX_NOACTIVATE.0 as isize
        | WS_EX_TOOLWINDOW.0 as isize
        | WS_EX_TOPMOST.0 as isize;
    SetWindowLongPtrW(hwnd, GWL_EXSTYLE, new_style);

    // Force the video preview window to topmost so it doesn't hide behind Explorer
    let _ = SetWindowPos(
        hwnd,
        HWND_TOPMOST,
        0,
        0,
        0,
        0,
        SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW,
    );
    let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
    true
}

/// Apply WS_EX_NOACTIVATE style to a window
/// Returns true if the window was found and modified
unsafe fn try_apply_noactivate_style(pid: u32) -> bool {
    // Reuse the window already found while it still exists and still belongs to
    // the player. This runs every few milliseconds for as long as a video plays,
    // so the desktop enumeration below is needed once, or again if ffplay
    // recreates its window (the cached handle stops being a window).
    let cached = VIDEO_HWND.load(Ordering::SeqCst);
    if cached != 0 {
        let hwnd = HWND(cached as *mut std::ffi::c_void);
        if IsWindow(hwnd).as_bool() {
            let mut window_pid: u32 = 0;
            GetWindowThreadProcessId(hwnd, Some(&mut window_pid));
            if window_pid == pid {
                return apply_noactivate_to_hwnd(hwnd);
            }
        }
        VIDEO_HWND.store(0, Ordering::SeqCst);
    }

    let mut data = EnumWindowsData {
        target_pid: pid,
        found_hwnd: HWND::default(),
        best_area: 0,
    };

    let _ = EnumWindows(
        Some(enum_windows_callback),
        LPARAM(&mut data as *mut EnumWindowsData as isize),
    );

    if !data.found_hwnd.is_invalid() {
        return apply_noactivate_to_hwnd(data.found_hwnd);
    }

    VIDEO_HWND.store(0, Ordering::SeqCst);
    false
}

/// Set WS_EX_NOACTIVATE on a window belonging to the given process
/// This prevents the window from stealing focus
/// Uses a singleton monitor thread so repeated previews don't spawn extra workers.
///
/// The thread spends its time between players waiting on an event rather than polling: a
/// window belongs to the player that is playing, and while none is, there is nothing to
/// re-assert — so what the wait is for is the hand-over a new player makes (see
/// `set_noactivate_for_process`), and a thread that woke every eighty milliseconds to find
/// that nothing had changed is a thread this app does not need. A wait that is never
/// signalled is not a lost wake-up either: the timeout it is given is the cadence the
/// window is kept in step at, so a hand-over that raced the wait is caught by the next one.
fn ensure_noactivate_monitor() {
    if NOACTIVATE_MONITOR_STARTED
        .compare_exchange(false, true, Ordering::AcqRel, Ordering::Acquire)
        .is_err()
    {
        return;
    }

    std::thread::spawn(|| {
        let mut monitored_pid: u32 = 0;
        let mut pid_started = Instant::now();

        while RUNNING.load(Ordering::Acquire) {
            let pid = VIDEO_PID.load(Ordering::Acquire);

            if pid != monitored_pid {
                monitored_pid = pid;
                pid_started = Instant::now();
                if pid == 0 {
                    VIDEO_HWND.store(0, Ordering::SeqCst);
                }
            }

            if pid != 0 {
                unsafe {
                    let _ = try_apply_noactivate_style(pid);
                }

                let elapsed = pid_started.elapsed();
                let delay_ms = if elapsed < Duration::from_millis(250) {
                    5
                } else if elapsed < Duration::from_secs(2) {
                    20
                } else {
                    100
                };
                wait_for_noactivate_wake(delay_ms);
            } else {
                // Nothing is playing: the wait is the player's own start, which is what
                // signals this thread rather than a clock.
                wait_for_noactivate_wake(NOACTIVATE_IDLE_WAIT_MS);
            }
        }

        NOACTIVATE_MONITOR_STARTED.store(false, Ordering::Release);
    });
}

/// How long the monitor thread waits while nothing is playing, which is a bound on how long
/// it takes to notice the run ending rather than a cadence anything is kept in step at.
const NOACTIVATE_IDLE_WAIT_MS: u64 = 500;

/// The event the monitor thread waits on: signalling it is a player's window having been
/// handed over (see `set_noactivate_for_process`). A handle that could not be created is the
/// null one, and a wait on that fails at once — which leaves the thread polling at the
/// timeout it was given, the way it did before there was an event at all.
///
/// It is kept as the number a handle is rather than as the handle, for the reason
/// `VIDEO_HWND` is: a handle is a raw pointer, and what is shared across threads here is a
/// number that the calls it is handed to take as a handle.
static NOACTIVATE_WAKE: Lazy<isize> = Lazy::new(|| {
    unsafe { CreateEventW(None, false, false, None) }
        .map(|handle| handle.0 as isize)
        .unwrap_or_default()
});

/// The wake event as the handle a Windows call takes.
fn noactivate_wake_handle() -> HANDLE {
    HANDLE(*NOACTIVATE_WAKE as *mut core::ffi::c_void)
}

/// Wait for a player's window to be handed over, or for `timeout_ms` to pass.
fn wait_for_noactivate_wake(timeout_ms: u64) {
    let handle = noactivate_wake_handle();
    if handle.0.is_null() {
        // An event that could not be created is not one to wait on: a failed wait returns
        // at once, and a thread that returned at once every time is a thread spinning on a
        // machine that has no event. What is left is the sleep this was before there was an
        // event at all, which costs the wakeups it always did and nothing more.
        std::thread::sleep(Duration::from_millis(timeout_ms));
        return;
    }

    // A failed wait is the timeout's, and it leaves the caller doing what it would have done
    // anyway: looking at the process it is watching.
    let _ = unsafe { WaitForSingleObject(handle, timeout_ms as u32) };
}

/// Wake the monitor thread: a player is playing, and the window it draws in is the
/// monitor's to keep in step (see `ensure_noactivate_monitor`).
fn wake_noactivate_monitor() {
    let _ = unsafe { SetEvent(noactivate_wake_handle()) };
}

fn set_noactivate_for_process(pid: u32) {
    VIDEO_PID.store(pid, Ordering::SeqCst);

    // First, do a few immediate synchronous checks with very tight timing
    // This minimizes the window where focus can be stolen
    unsafe {
        for _ in 0..10 {
            if try_apply_noactivate_style(pid) {
                // Found and modified - but keep monitoring in case window is recreated
                break;
            }
            // Very short spin-wait for the first attempts
            std::thread::yield_now();
        }
    }

    ensure_noactivate_monitor();
    // The monitor is woken rather than left to notice on its own clock: between players it
    // is waiting on this, and a player whose window appears while that wait runs is one the
    // thread would otherwise look at up to half a second later (see
    // `ensure_noactivate_monitor`).
    wake_noactivate_monitor();
}

/// Start ffplay for video preview with configurable volume
fn start_video_playback(path: &PathBuf, x: i32, y: i32, width: i32, height: i32) -> Option<Child> {
    // Get volume setting from config (0-100)
    let volume = CONFIG.lock().map(|c| c.video_volume).unwrap_or(0);

    // Use ffplay for video playback - borderless, positioned at preview location
    let mut cmd = Command::new("ffplay");

    // If volume is 0, disable audio completely for better performance
    if volume == 0 {
        cmd.arg("-an");
    } else {
        // Convert percentage to ffplay volume filter (0-100 maps to 0.0-1.0)
        let volume_filter = format!("volume={:.2}", volume as f64 / 100.0);
        cmd.args(["-af", &volume_filter]);
    }

    // The geometry is read from the cache and never probed for here: this runs on the
    // preview thread, which is the one thread that must not wait for two external
    // processes — and by the time a player is started for a hover, the probe that sized
    // that hover has already answered (see `probe_video_geometry`). A file whose answer is
    // that there is nothing to measure gets no filter at all, which is the frame as the
    // file holds it.
    let vf = match cached_video_geometry(path) {
        Some(ProbedGeometry::Measured(geometry)) => Some(match geometry.crop {
            Some(crop) => format!(
                "crop={}:{}:{}:{},setsar=1",
                crop.width, crop.height, crop.x, crop.y
            ),
            None => "setsar=1".to_string(),
        }),
        _ => None,
    };
    if let Some(vf) = vf.as_deref() {
        cmd.args(["-vf", vf]);
    }

    let child = cmd
        .args([
            "-err_detect",
            "ignore_err", // Ignore header/stream errors
            "-fflags",
            "+genpts+discardcorrupt+igndts", // Handle missing timestamps & corrupt data
            "-framedrop",                    // Drop undecodable frames instead of stalling
            "-loop",
            "0",         // Loop forever
            "-noborder", // No window border
            "-left",
            &x.to_string(),
            "-top",
            &y.to_string(),
            "-x",
            &width.to_string(),
            "-y",
            &height.to_string(),
            "-autoexit",
            "-loglevel",
            "quiet",
        ])
        .arg(path)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .creation_flags(engine_processes::CREATE_NO_WINDOW) // Hide the console window
        .spawn()
        .ok();

    // After spawning, try to set WS_EX_NOACTIVATE on the ffplay window
    // to prevent it from stealing focus
    if let Some(ref child_process) = child {
        set_noactivate_for_process(child_process.id());

        // The player is this app's own child, and it is taken charge of the way the
        // engines are: one that is up when the app is killed, or crashes, is not left
        // playing with nothing to close it — the job ends it there and then, and the
        // record is what answers for the run that never got to end it. It is recorded
        // as the player rather than as an engine, because what ends a player is its
        // hover ending: a tier being let go of — a preview type switched off, a worker
        // given up on — is not its to receive; see `engine_processes`.
        engine_processes::record_player(VIDEO_PROCESS_IMAGE_NAME, child_process.id());
    }

    child
}

/// A player that has been started and has not put its window up yet.
///
/// A video's preview is the player's own window, and a player is a process: what stands
/// in for the video until that window is there is the waiting spinner, at the pointer it
/// is shown at for every other kind of wait, and the media the player was started for is
/// held here until it is. Holding it here rather than in `CURRENT_MEDIA` is what keeps
/// the spinner on screen: the frame the player will play into would otherwise be what
/// this app's window is showing, and a video's window is the one the player draws.
struct VideoStart {
    /// The frame the player plays into, with the process it was started for.
    media: MediaData,
    path: PathBuf,
    pid: u32,
    started: Instant,
}

/// How long a player is given to put its window up before the wait for it is given up
/// on: a player that has not by then is one that will not, and a spinner that never ends
/// is worse than the desktop it leaves behind.
const VIDEO_START_WAIT_SECS: u64 = 10;

/// How the wait for a player stands: whether the video is on screen now, whether the
/// wait is over some other way, or whether the player is still starting.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum PlayerWait {
    /// The player's window is up: what is on screen is the video from here, and the media
    /// it was started for is the preview.
    Arrived,
    /// The player is not coming: it died, its hover moved on, or it has taken longer than
    /// a start ever does.
    Abandoned,
}

/// How the wait for a player stands. `None` is a player still starting, which is a wait
/// that goes on.
///
/// A window that is up arrives, the cap included: the video is on screen whether or not
/// the start took long, and ending a player that has one would take a picture away. What
/// is read before that is a player that is gone — what a process leaves behind is a
/// handle and not a window, so a window with no player behind it is not a preview — and
/// a start that has run past `VIDEO_START_WAIT_SECS` with no window to show for it is one
/// this app stops watching, because a spinner that never ends is worse than the desktop
/// it leaves behind.
fn player_wait(window_up: bool, player_alive: bool, waited: Duration) -> Option<PlayerWait> {
    if window_up && player_alive {
        return Some(PlayerWait::Arrived);
    }

    if !player_alive || waited >= Duration::from_secs(VIDEO_START_WAIT_SECS) {
        return Some(PlayerWait::Abandoned);
    }

    None
}

/// Stop video playback, on the thread the engine belongs to.
fn stop_video_playback(media: &mut MediaData) {
    // A video the media engine is playing has no process and no window of its own: letting
    // the engine go is the whole of stopping it, and it is done here because this is where
    // every path that ends a video already comes through — the pointer leaving the file,
    // another preview taking its place, the `Videos` gate closing, a display change, a
    // resume from sleep, and the app itself.
    //
    // A sound is stopped here for the same reason and by the same call: the engine that plays
    // one is this app's own, and a sound FFmpeg plays instead is a process in `video_process`
    // below — the same field, killed the same way, because what ends either is the hover
    // ending.
    //
    // The engine's ending is this thread's to perform and no other's — a session belongs to
    // the thread that started it (see `video_player`'s `SESSION`) — so a take-down that runs
    // on some other thread kills the player process and leaves the media engine for the
    // thread that owns it (see `kill_player_process` and `hide_preview`).
    //
    // What a sound had played of its file is written down here rather than where it is heard:
    // a hover that is ending is a sound whose position is settled, and the memory the mode that
    // resumes one reads back is asked to keep what it has before the clock that measured it is
    // taken down (see `audio_seek::flush`).
    audio_seek::flush();

    if media.media_type.is_native_video() || media.media_type.is_audio() {
        video_player::stop();
    }

    kill_player_process(media);
}

/// End the player process a preview holds, if it has one, and forget its window.
///
/// It is the half of a take-down that any thread may perform — a process is ended with a
/// handle, whatever thread holds it — which is what lets the hide on the Explorer hook's
/// thread end the `ffplay` a hover started without reaching for the media engine, whose
/// session is the preview thread's alone (see `stop_video_playback` and `hide_preview`).
fn kill_player_process(media: &mut MediaData) {
    if let Some(ref mut process) = media.video_process {
        // Kill only, never wait: a process stuck in kernel I/O would block the
        // caller (possibly the Explorer hook thread) indefinitely. The leftover
        // process checks confirm death and clear VIDEO_PID.
        let _ = process.kill();
    }
    media.video_process = None;
    // Clear the video window HWND. VIDEO_PID stays recorded until the process
    // is confirmed gone, so a surviving ffplay can still be found and killed.
    VIDEO_HWND.store(0, Ordering::SeqCst);
}

/// Check if the current ffplay process is still running
/// Clears stored state if the process has exited
fn is_video_process_running() -> bool {
    if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
        if let Some(ref mut media) = *media_guard {
            if let Some(ref mut process) = media.video_process {
                match process.try_wait() {
                    Ok(Some(_)) => {
                        let pid = process.id();
                        media.video_process = None;
                        VIDEO_HWND.store(0, Ordering::SeqCst);
                        VIDEO_PID.store(0, Ordering::SeqCst);
                        // The player is confirmed gone, so the record of it goes with
                        // it rather than being left for the next run to look for.
                        engine_processes::forget(pid);
                        return false;
                    }
                    Ok(None) => return true,
                    Err(_) => {
                        media.video_process = None;
                        VIDEO_HWND.store(0, Ordering::SeqCst);
                        // Keep VIDEO_PID: the process is not confirmed dead, so
                        // the leftover-process checks can still find and kill it.
                        return false;
                    }
                }
            }
        }
    }
    false
}

/// True when the handle refers to a process whose executable file name matches
/// `expected_name` (compared case-insensitively).
unsafe fn process_image_matches(handle: HANDLE, expected_name: &str) -> bool {
    let mut buffer = [0u16; 1024];
    let mut len = buffer.len() as u32;
    if QueryFullProcessImageNameW(
        handle,
        PROCESS_NAME_WIN32,
        PWSTR(buffer.as_mut_ptr()),
        &mut len,
    )
    .is_err()
    {
        return false;
    }

    let path = String::from_utf16_lossy(&buffer[..len as usize]);
    std::path::Path::new(&path)
        .file_name()
        .is_some_and(|name| name.eq_ignore_ascii_case(std::ffi::OsStr::new(expected_name)))
}

/// True when `pid` still refers to a live ffplay process.
fn is_ffplay_pid_alive(pid: u32) -> bool {
    unsafe {
        let Ok(handle) = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, pid) else {
            return false;
        };
        let matches = process_image_matches(handle, VIDEO_PROCESS_IMAGE_NAME);
        let _ = CloseHandle(handle);
        matches
    }
}

/// Terminate `pid` when it is still the ffplay process we spawned. Non-blocking:
/// it only requests the termination, it never waits for the process to exit.
fn terminate_ffplay_pid(pid: u32) {
    unsafe {
        let Ok(handle) = OpenProcess(
            PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION,
            false,
            pid,
        ) else {
            return;
        };
        if process_image_matches(handle, VIDEO_PROCESS_IMAGE_NAME) {
            let _ = TerminateProcess(handle, 1);
        }
        let _ = CloseHandle(handle);
    }
}

/// Clear the recorded video process state once `pid` is confirmed gone.
fn clear_video_process_state(pid: u32) {
    if VIDEO_PID
        .compare_exchange(pid, 0, Ordering::SeqCst, Ordering::SeqCst)
        .is_ok()
    {
        VIDEO_HWND.store(0, Ordering::SeqCst);
        // The player is confirmed gone, so the record of it goes with it rather than
        // being left for the next run to look for.
        engine_processes::forget(pid);
    }
}

/// Kill the last spawned ffplay process when it is still alive.
///
/// A process can outlive its `Child` handle: a kill may not take effect, or the
/// handle may be dropped before the process is confirmed gone. Without this a
/// surviving ffplay keeps its window on screen and the next hover would spawn a
/// second one next to it. The PID is only cleared once the process is confirmed
/// gone, so a later call retries instead of losing track of it.
pub fn kill_stray_video_process() {
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    if pid == 0 {
        return;
    }

    if !is_ffplay_pid_alive(pid) {
        clear_video_process_state(pid);
        return;
    }

    terminate_ffplay_pid(pid);

    if !is_ffplay_pid_alive(pid) {
        clear_video_process_state(pid);
    }
}

/// Ensure the ffplay window is topmost and positioned correctly
fn ensure_video_window_topmost(x: i32, y: i32, width: i32, height: i32) -> bool {
    // Re-discover/re-apply style by PID each time to survive ffplay window recreation
    // and keep topmost state resilient over time.
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    if pid != 0 {
        unsafe {
            let _ = try_apply_noactivate_style(pid);
        }
    }

    let hwnd_val = VIDEO_HWND.load(Ordering::SeqCst);
    if hwnd_val == 0 {
        return false;
    }

    unsafe {
        let hwnd = HWND(hwnd_val as *mut std::ffi::c_void);
        if hwnd.is_invalid() {
            VIDEO_HWND.store(0, Ordering::SeqCst);
            return false;
        }

        // Re-assert desired style bits in case ffplay modified them
        let current_style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
        let new_style = current_style
            | WS_EX_NOACTIVATE.0 as isize
            | WS_EX_TOOLWINDOW.0 as isize
            | WS_EX_TOPMOST.0 as isize;
        if new_style != current_style {
            SetWindowLongPtrW(hwnd, GWL_EXSTYLE, new_style);
        }

        let _ = SetWindowPos(
            hwnd,
            HWND_TOPMOST,
            x,
            y,
            width,
            height,
            SWP_NOACTIVATE | SWP_SHOWWINDOW,
        );
        let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
    }

    true
}

/// Load media (image, animated image, text, or video) with appropriate loader
///
/// `max_width` x `max_height` is the box the loader may draw in, and every loader
/// draws its frame inside it — resized, clamped or rasterized at that size. That is
/// the rule rather than a convenience: the preview window is sized to the frame that
/// comes back from here rather than to the box the layout planned (see
/// `render_layered_preview_at`), and the box is what was fitted to the display, so a
/// frame larger than the box is a preview hanging off the edge of the display. A
/// source that draws at whatever size it is asked for — the Windows PDF engine, whose
/// destination is in DIPs and comes back scaled by the display — has to be drawn back
/// into the box it was given; see `pdf_preview::fit_drawn_page`.
///
/// The one source with no frame to draw is an SVG document, which is the engine's: what
/// comes back for one is the kind alone, and the install path hands the hover over
/// rather than putting anything of this app's up (see `MediaType::EngineSvg`).
fn load_media(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    dpi: u32,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    // Every loader below reads the file's bytes, so a file whose content is still
    // in the cloud is answered here rather than after a download the user never
    // asked for. The hook refuses these too; this is the boundary that reads, so
    // it decides for itself rather than trusting that nothing reaches it.
    if cloud_files::needs_download(path) {
        return None;
    }

    // What the file's content says it is comes ahead of what its name does, where the two
    // disagree: a `.docx` whose bytes are an MP4 is loaded as the video it is, and a format
    // no kind of this app previews is loaded as nothing at all — see `content_type` for
    // what settles that, and `load_media_of_kind` for where the kind is handed on.
    let content = CONFIG
        .lock()
        .ok()
        .map(|config| crate::formats::content_type::of(path, &config))
        .unwrap_or(crate::formats::content_type::Content::Unknown);

    match content {
        crate::formats::content_type::Content::Kind(kind) => {
            return load_media_of_kind(
                kind,
                path,
                max_width,
                max_height,
                preview_scale,
                dpi,
                cancel,
            )
        }
        crate::formats::content_type::Content::Foreign => return None,
        crate::formats::content_type::Content::Unknown => {}
    }

    // What the file is, is the router's answer: one order, asked once, and the same one the hook
    // that admitted this hover asked (see `formats::routing`). The configuration is taken and
    // given up around that question alone, so nothing below is holding it.
    let kind = {
        let Ok(config) = CONFIG.lock() else {
            return None;
        };

        crate::formats::routing::kind_of(path, &config)
    };

    let Some(kind) = kind else {
        // A name no list claims has always been the picture path's, and a drawing among those is
        // still the drawing layer's: what it is, is its own header's answer rather than its
        // name's, and an `svg` a hand-edited list no longer names is a document this app can
        // draw. The hook refuses such a file before a hover reaches this far (see
        // `explorer_hook::is_media_file`), so this is the answer for the hover that came the
        // other way — through the content, which named no kind either.
        if svg_preview::is_svg_file(path) {
            return webview_preview::draws(path).then(engine_svg_media);
        }

        return load_picture(path, max_width, max_height, preview_scale, &cancel);
    };

    load_media_of_kind(
        kind,
        path,
        max_width,
        max_height,
        preview_scale,
        dpi,
        cancel,
    )
}

/// The loader for a file whose content named a kind of its own — see `content_type`.
///
/// It is the arm the chain above would have taken had the file been named what its
/// content says it is, reached by the kind rather than by the name: the same loaders, and
/// the same one for each kind. What every one of them reads is the file itself and never
/// the name it is under, which is what makes this a routing rather than a rename.
///
/// The gates are not asked here, exactly as they are not asked by the chain: the hook
/// asked them before a hover could reach this path at all.
fn load_media_of_kind(
    kind: PreviewType,
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    dpi: u32,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    match kind {
        PreviewType::Videos => load_video_thumbnail(path, max_width, max_height, preview_scale),
        // A book is one kind with two readers, and which of the two a file is, is asked of the
        // table that names them rather than assumed here: a page is the PDF engine's, a comic is
        // the first plate read out of the container it is published in, and a comic reached by
        // its *content* would otherwise be read as a page of a PDF that does not exist (see
        // `native_formats`). What cannot be read for — a configuration that will not open — is
        // answered as the page a book most often is.
        PreviewType::Ebook => {
            let job = CONFIG
                .lock()
                .ok()
                .and_then(|config| native_formats::job_for(path, PreviewType::Ebook, &config));

            match job {
                Some(native_formats::NativeJob::Comic) => {
                    load_comic_page(path, max_width, max_height, preview_scale)
                }
                _ => load_pdf_first_page(path, max_width, max_height, preview_scale),
            }
        }
        PreviewType::Archives => load_archive_preview(
            path,
            max_width,
            max_height,
            dpi,
            current_archive_options(),
            MediaType::Archive,
            &cancel,
        ),
        PreviewType::Document => {
            load_office_preview(path, max_width, max_height, preview_scale, &cancel)
                .or_else(|| load_engine_page_for_office(path, max_width, max_height, preview_scale))
        }
        PreviewType::Libre => libreoffice_render::rendered_page(path).and_then(|page| {
            load_engine_page(
                &page,
                MediaType::Libre,
                max_width,
                max_height,
                preview_scale,
            )
        }),
        PreviewType::Magick => {
            load_magick_picture(path, max_width, max_height, preview_scale, &cancel)
        }
        // An archive an engine listed is loaded as an archive: the listing it produced is in the
        // same cache under the same key, so the page is measured and painted from it without this
        // arm knowing where it came from — and a file the engine has not answered for yet is a
        // listing the cache does not hold, which is the wait the hover is already in.
        PreviewType::Peazip => load_archive_preview(
            path,
            max_width,
            max_height,
            dpi,
            current_archive_options(),
            MediaType::Peazip,
            &cancel,
        ),
        PreviewType::Calibre => calibre_render::rendered_page(path)
            .and_then(|page| load_book_page(&page, max_width, max_height, preview_scale)),
        PreviewType::Design => load_design_preview(path, max_width, max_height, preview_scale),
        // Which half of the drawing kind this is, is the name's to say here rather than the
        // content's: a document is drawn by the browser engine and a metafile by the drawing
        // layer, and the content has already answered that the file is a drawing at all.
        PreviewType::Vector => {
            if svg_preview::is_svg_file(path) {
                webview_preview::draws(path).then(engine_svg_media)
            } else {
                load_vector_preview(path, max_width, max_height, preview_scale)
            }
        }
        PreviewType::Text => {
            load_text_preview(path, max_width, max_height, dpi, current_text_options())
        }
        // A sound: a card of what the file holds, painted like an archive's page. The facts are
        // the probe's and are already in hand — the measure that read them is what laid this
        // hover out (see `audio_box`) — and the player the card is drawn against is started by
        // the loop, where every other preview is put up.
        PreviewType::Audio => load_audio_card(path, max_width, max_height, dpi),
        PreviewType::Fonts => (font_preview::probe(path).is_some() && webview_preview::draws(path))
            .then(engine_font_media),
        PreviewType::Images => load_picture(path, max_width, max_height, preview_scale, &cancel),
    }
}

/// The picture path: the animated reader the file's own bytes call for, and everything else as
/// the still picture it is.
///
/// It is where the chain above ends for every name that reaches it, and the arm a picture the
/// content named is loaded by. Which reader is asked about it is the job the router gives a
/// picture (`native_formats::picture_job`), and that job is the file's own bytes first: a `.gif`
/// with one frame in it is a still here, and the animated reader is never asked about one. That
/// job is the file's answer alone, so nothing here takes the configuration or holds its lock
/// across the read of the head.
fn load_picture(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    cancel: &Arc<AtomicBool>,
) -> Option<MediaData> {
    let job = native_formats::picture_job(path);

    match job {
        native_formats::NativeJob::AnimatedGif => {
            if let Some(media) = load_animated_gif(
                path,
                max_width,
                max_height,
                preview_scale,
                Arc::clone(cancel),
            ) {
                return Some(media);
            }
        }
        native_formats::NativeJob::AnimatedWebp => {
            if let Some(media) = load_animated_webp(
                path,
                max_width,
                max_height,
                preview_scale,
                Arc::clone(cancel),
            ) {
                return Some(media);
            }
        }
        native_formats::NativeJob::AnimatedApng => {
            if let Some(media) = load_animated_apng(
                path,
                max_width,
                max_height,
                preview_scale,
                Arc::clone(cancel),
            ) {
                return Some(media);
            }
        }
        native_formats::NativeJob::Picture
        | native_formats::NativeJob::PictureCodec
        | native_formats::NativeJob::Text
        | native_formats::NativeJob::SvgDocument
        | native_formats::NativeJob::Metafile
        | native_formats::NativeJob::Eps
        | native_formats::NativeJob::FontSpecimen
        | native_formats::NativeJob::Pdf
        | native_formats::NativeJob::Comic
        | native_formats::NativeJob::Psd
        | native_formats::NativeJob::Project
        | native_formats::NativeJob::ArchiveZip
        | native_formats::NativeJob::ArchiveSevenZ
        | native_formats::NativeJob::ArchiveRar
        | native_formats::NativeJob::ArchiveTar
        | native_formats::NativeJob::ArchiveTarGz
        | native_formats::NativeJob::VideoMediaFoundation
        | native_formats::NativeJob::AudioMediaFoundation => {}
    }

    // What is left is a still: a picture that never moved, or one whose animated reader
    // answered nothing for it. Nothing is decoded twice to find that out.
    if cancel.load(Ordering::Acquire) {
        return None;
    }
    load_static_image(path, max_width, max_height, preview_scale)
}

/// The box a video hover is placed at.
///
/// A video is measured by a probe — `ffprobe` and a cropdetect pass, two external
/// processes — and the probe is what a hover waits for when its file has not been
/// measured yet: the answer then is the waiting box, which is the box every wait is shown
/// in, and the layout that follows the probe's own replay reads the size here instead
/// (see `video_probe_due`). What the probe answered when it does answer is one of two
/// things, and the file's own lack of an answer is neither: a shape is the shape, and a
/// file with nothing to measure is placed at the 16:9 box FFmpeg's player is handed a
/// file it could not measure at — while the media engine, which plays only what it can
/// open, is answered with no size at all, which is how the layout drops it.
fn video_box(path: &Path) -> Option<(u32, u32)> {
    match cached_video_geometry(path) {
        Some(ProbedGeometry::Measured(geometry)) => Some((geometry.width, geometry.height)),
        Some(ProbedGeometry::Unmeasurable) => {
            (!codecs::plays_video_natively()).then_some((1920, 1080))
        }
        // Not probed yet: the wait for the probe, which is the box the hover is placed
        // in until the answer lands and the hover is replayed.
        None => Some((office_preview::WAITING_BOX, office_preview::WAITING_BOX)),
    }
}

/// Whether this file is a video the probe has not answered for yet.
///
/// A hover for one cannot be laid out as a video — the layout has no shape to place — so
/// it is laid out as the wait for the probe and replayed when the answer lands (see
/// `video_probe` in the preview loop). A file the probe has already answered for is not a
/// wait, whatever the answer was: an unmeasurable video is a video with a fallback box,
/// not one to be probed again on every hover.
fn video_probe_due(path: &Path) -> bool {
    drawn_as_video(path) && cached_video_geometry(path).is_none()
}

/// Whether the preview of `path` is a video: the name the video list carries, or the bytes
/// of a video under a name that list does not have.
///
/// Every question about a video goes through this one answer — whether its shape has to be
/// probed, whether the wait for it is shown, and whether the player takes over the window
/// rather than this app drawing its frames — because the loader plays the file its bytes
/// name, and a hover whose picture is played but whose frames are awaited would sit on a
/// first frame that nothing ever replaces.
fn drawn_as_video(path: &Path) -> bool {
    let content = CONFIG
        .lock()
        .ok()
        .map(|config| crate::formats::content_type::of(path, &config))
        .unwrap_or(crate::formats::content_type::Content::Unknown);

    if matches!(
        content,
        crate::formats::content_type::Content::Kind(PreviewType::Videos)
    ) {
        return PreviewType::Videos.enabled();
    }

    video_formats::is_video_preview(path)
}

/// The box a PDF page asks for, measured off the preview thread.
///
/// A page's size is read out of the document, and the PDF engine opens it whole to read it —
/// which on a book of a thousand pages is a read that can be felt — so it is measured the way a
/// video's shape is: the hover is laid out as the wait and replayed when the answer lands (see
/// `measured_off_the_tick`).
fn pdf_page_box(path: &Path) -> Option<(u32, u32)> {
    let source = path.to_path_buf();

    measured_off_the_tick(path, MeasureScope::File, move || {
        pdf_preview::page_dimensions(&source)
    })
}

/// The box a document the engine draws asks for: the size its own markup declares, read the
/// same way and for the same reason — a document is read and parsed whole to be measured.
fn svg_box(path: &Path) -> Option<(u32, u32)> {
    let source = path.to_path_buf();

    measured_off_the_tick(path, MeasureScope::File, move || {
        svg_preview::measure(&source)
    })
}

/// The box a specimen is drawn in, held for a file that parses as a font.
///
/// The box is this app's own — a font has no size it asks to be drawn at — so what the measure
/// answers is whether the file is a font at all, and that costs a read of it: a collection is
/// read whole to reach the face the specimen shows (see `font_preview::probe`).
fn font_box(path: &Path) -> Option<(u32, u32)> {
    let source = path.to_path_buf();

    measured_off_the_tick(path, MeasureScope::File, move || {
        font_preview::probe(&source)
            .map(|_| (font_preview::SPECIMEN_WIDTH, font_preview::SPECIMEN_HEIGHT))
    })
}

/// The box a vector drawing asks for, measured the same way: what a metafile declares is read
/// out of the whole file, which a drawing of any size takes with it (see
/// `metafile_image::dimensions`).
fn vector_box(path: &Path) -> Option<(u32, u32)> {
    let source = path.to_path_buf();

    measured_off_the_tick(path, MeasureScope::File, move || vector_dimensions(&source))
}

/// The box a listing asks for, measured off the preview thread: an archive's table of contents
/// is a read that can be felt — every entry of a zip walked, a `.tar.gz` inflated to reach one
/// — and the wait for it is the same kind of wait (see `measured_off_the_tick`).
///
/// It is asked of the archives this app reads itself. An archive an engine lists is measured by
/// `archive_box` instead: what that side waits for is the engine's own listing, and a listing
/// that has been remembered is a page measured out of memory rather than out of a file (see
/// `peazip_box`).
fn archive_box_off_the_tick(path: &Path, bounds: ScreenBounds, dpi: u32) -> Option<(u32, u32)> {
    let source = path.to_path_buf();
    let cap_width = (bounds.right - bounds.left).max(1) as u32;
    let cap_height = bounds.height().max(1) as u32;
    let options = current_archive_options();

    measured_off_the_tick(
        path,
        MeasureScope::Room {
            cap_width,
            cap_height,
            dpi,
            theme: options.theme,
            font_scale_percent: options.font_scale_percent,
        },
        move || archive_preview::measure(&source, cap_width, cap_height, dpi, options),
    )
}

/// How long a sound's probe is given — the source reader's own read of a file, or FFmpeg's
/// container probe — before the wait for it is over. The same cap a video's probe has, and for
/// the same reason: a probe that has run this long is a file nothing is coming back from.
const AUDIO_PROBE_TIMEOUT_SECS: u64 = 10;

/// How often a sound's card is painted again while a player is running. The clock changes once
/// a second and the bar creeps by a few pixels in that time, so four times a second is smooth
/// to the eye and a fraction of what a video's own frames cost.
const AUDIO_CARD_REPAINT: Duration = Duration::from_millis(250);

/// How often a card whose name does not fit is painted again, which is the same question asked
/// for the one thing about a card that moves faster than a clock: a name is scrolled across it
/// at about an advance a 100 ms, and a cadence coarse enough for a clock to read smoothly would
/// show that as a slideshow. A card is some four hundred pixels square, so what thirty of them
/// a second costs is a fraction of what the spinner's own overlay costs at the same rate (see
/// `AUDIO_CARD_REPAINT`).
const AUDIO_NAME_REPAINT: Duration = Duration::from_millis(33);

/// The box a sound's card asks for, with the probe that fills it beside it on the same thread.
///
/// Two things a hover on a sound waits for, and both of them are here. The first is the probe:
/// whether this machine has anything that plays the file at all, and what the file says about
/// itself — a source reader for the engine's own decoders, an `ffprobe` pass for FFmpeg's —
/// and a file neither of them can play is a hover answered with nothing rather than with a card
/// of facts nothing will ever play. The second is the card's own layout, which is a page of
/// text wrapped to the room the display has.
///
/// Both are off the preview thread, and what the hover waits in meanwhile is the spinner: a
/// probe is a process, and the box it answers with is what the replayed hover is laid out at
/// (see `measured_off_the_tick`).
fn audio_box(path: &Path, bounds: ScreenBounds, dpi: u32) -> Option<(u32, u32)> {
    let source = path.to_path_buf();
    let cap_width = (bounds.right - bounds.left).max(1) as u32;
    let cap_height = bounds.height().max(1) as u32;
    let options = current_audio_options();

    measured_off_the_tick(
        path,
        MeasureScope::Room {
            cap_width,
            cap_height,
            dpi,
            theme: options.theme,
            font_scale_percent: options.font_scale_percent,
        },
        move || {
            // What the machine has for the file, asked once per file and version and held for
            // the hovers that follow — a file the engine will not play costs one probe rather
            // than one per hover.
            if matches!(audio_track::probed(&source), Probed::NotAsked) {
                let probed = probe_audio_track(&source);
                audio_track::remember(
                    &source,
                    match &probed {
                        Some(track) => Probed::Track(track.clone()),
                        None => Probed::Nothing,
                    },
                );
            }

            let card = audio_card(&source, None, None, 0)?;
            audio_preview::measure(&card, cap_width, cap_height, dpi, options)
        },
    )
}

/// Whether the preview of `path` is a sound: what the file's own bytes say it is — the verdict a
/// probe left behind included — and, for a name no table names, the sound list.
///
/// It is asked the way `drawn_as_video` is asked and for the same reason: a sound is drawn as a
/// card by this app rather than by a player, so the layout has to know one when it sees one —
/// which for a renamed file, or for a container whose streams hold only a song, is a question
/// about the content rather than about the name.
fn drawn_as_audio(path: &Path) -> bool {
    if !PreviewType::Audio.enabled() {
        return false;
    }

    let Some(config) = CONFIG.lock().ok() else {
        return false;
    };

    if matches!(
        crate::formats::content_type::of(path, &config),
        crate::formats::content_type::Content::Kind(PreviewType::Audio)
    ) {
        return true;
    }

    audio_formats::matches_audio_list(path, &config.audio_extensions)
}

/// What a sound's card says, with the clock as it stands — or nothing for a file with no track
/// behind it, which is a file no probe has answered for or one nothing here can play.
///
/// `name_offset` is how far a name the card has no room for has been scrolled: it is nothing
/// for the card a hover is measured with and for the first frame of one, and what the repaints
/// of a moving card hand over (see `audio_preview::NameScroll`).
fn audio_card(
    path: &Path,
    elapsed: Option<f64>,
    duration: Option<f64>,
    name_offset: i32,
) -> Option<Card> {
    let Probed::Track(track) = audio_track::probed(path) else {
        return None;
    };

    Some(Card {
        name: audio_preview::name_of(path),
        facts: audio_preview::facts_of(&track, path),
        duration: duration.or(track.duration),
        elapsed,
        name_offset,
    })
}

/// The card a sound is previewed as, painted into the box the layout settled on.
fn load_audio_card(path: &Path, width: u32, height: u32, dpi: u32) -> Option<MediaData> {
    let card = audio_card(path, None, None, 0)?;
    let (pixels, width, height) =
        audio_preview::render(&card, width, height, dpi, current_audio_options())?;

    Some(MediaData {
        frames: vec![ImageFrame::new(pixels, width, height, 0)],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::Audio,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
    })
}

/// What this machine has for playing a sound: Windows' own decoders where one of them reaches
/// the format, and FFmpeg's player where none does.
///
/// The two are asked in the order the chain names them, and the first that answers is the
/// player the file is played by: the engine answers for a file it can decode — which is the
/// whole of what its probe is for — and everything else is FFmpeg's, where FFmpeg is installed
/// at all. A file neither answers for is a file with no preview.
fn probe_audio_track(path: &Path) -> Option<audio_track::Track> {
    if let Some(track) = video_player::audio_probe(path) {
        return Some(track);
    }

    ffprobe_audio_track(path)
}

/// What FFmpeg's own probe reports about a file, and the player that would play it.
///
/// The container is opened and its streams read rather than the file played to find out, which
/// is the pair of answers this side wants: whether there is a sound in the file at all, and
/// what the card beside it says. A machine without FFmpeg is answered by the engine above or
/// not at all.
fn ffprobe_audio_track(path: &Path) -> Option<audio_track::Track> {
    if !codecs::ffplay_available() {
        return None;
    }

    // Spawned rather than run through `Command::output` so that the probe is in the job before
    // it is waited on, and waited for under a cap rather than for as long as it takes — the
    // same arrangement the video path's own probes have (see `wait_bounded`).
    let child = Command::new("ffprobe")
        .args([
            "-v",
            "error",
            "-err_detect",
            "ignore_err",
            "-fflags",
            "+genpts+discardcorrupt+igndts",
            "-show_entries",
            "format=duration:stream=codec_type,codec_name,sample_rate,channels,bit_rate",
            "-of",
            "default=noprint_wrappers=1",
        ])
        .arg(path)
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .creation_flags(engine_processes::CREATE_NO_WINDOW)
        .spawn()
        .ok()?;

    engine_processes::adopt(child.id());

    let output = wait_bounded(child, Duration::from_secs(AUDIO_PROBE_TIMEOUT_SECS))?;
    let report = String::from_utf8_lossy(&output.stdout);

    audio_track_from_report(&report)
}

/// The track an `ffprobe` report describes, or nothing where the file holds no sound.
///
/// The entries arrive one stream at a time, so what follows a `codec_type=audio` line is that
/// stream's own fields and nothing of the streams before it — which is what makes a film with a
/// soundtrack distinguishable from a song.
fn audio_track_from_report(report: &str) -> Option<audio_track::Track> {
    let mut in_audio = false;
    let mut heard_audio = false;
    let mut codec: Option<String> = None;
    let mut rate: Option<u32> = None;
    let mut channels: Option<u16> = None;
    let mut bitrate: Option<u32> = None;
    let mut duration: Option<f64> = None;

    for line in report.lines() {
        let Some((key, value)) = line.split_once('=') else {
            continue;
        };
        let value = value.trim();

        match key.trim() {
            "codec_type" => {
                in_audio = value == "audio";
                heard_audio |= in_audio;
            }
            "codec_name" if in_audio => codec = Some(codec_label(value)),
            "sample_rate" if in_audio => rate = value.parse().ok(),
            "channels" if in_audio => channels = value.parse().ok(),
            "bit_rate" if in_audio => bitrate = value.parse().ok(),
            "duration" => duration = value.parse().ok(),
            _ => {}
        }
    }

    heard_audio.then_some(audio_track::Track {
        player: Player::Ffmpeg,
        codec,
        rate: rate.filter(|rate| *rate > 0),
        channels: channels.filter(|channels| *channels > 0),
        bitrate: bitrate.filter(|bitrate| *bitrate > 0),
        duration: duration.filter(|duration| *duration > 0.0),
    })
}

/// What a codec is called, by the name FFmpeg writes it under: the words a person reads on a
/// label where the codec has one, and the name itself where it does not.
fn codec_label(name: &str) -> String {
    match name {
        "mp3" => "MP3",
        "flac" => "FLAC",
        "alac" => "ALAC",
        "aac" => "AAC",
        "opus" => "Opus",
        "vorbis" => "Vorbis",
        "speex" => "Speex",
        "wmav1" | "wmav2" | "wmapro" => "WMA",
        "wmalossless" => "WMA Lossless",
        "ac3" => "Dolby Digital",
        "eac3" => "Dolby Digital Plus",
        "dts" => "DTS",
        "ape" => "Monkey's Audio",
        "wavpack" => "WavPack",
        "tta" => "True Audio",
        "musepack" | "mpc7" | "mpc8" => "Musepack",
        "shorten" => "Shorten",
        "tak" => "TAK",
        "amrnb" => "AMR",
        "amrwb" => "AMR-WB",
        "cook" | "atrac3" | "atrac3p" | "sipr" => "RealAudio",
        "dsd_lsbf" | "dsd_msbf" | "dsd_lsbf_planar" | "dsd_msbf_planar" => "DSD",
        name if name.starts_with("pcm_") => "PCM",
        name => return name.to_uppercase(),
    }
    .to_string()
}

/// Start the player a sound's card is drawn against, answering whether a player that was
/// expected arrived.
///
/// A card is drawn whether or not anything plays: at `Volume → Audio` 0% the answer is the card
/// and nothing else, which is what silence looks like and is not a failure. What the caller is
/// told is whether a player that *was* asked for came up — a sound no engine here will actually
/// play is a hover answered with nothing rather than a card whose clock can never move.
///
/// `start` is where in the file the sound is dropped, and it is the caller's answer: it is a
/// question about the file's length and the tray's `Volume → Audio Seek`, both of which are
/// read where the hover is answered (see `audio_seek::start_position`).
fn start_audio_playback(path: &Path, media: &mut MediaData, start: f64) -> bool {
    let Some(track) = audio_track::playable(path) else {
        return false;
    };

    // A player the hover before this one left behind is ended before this one starts, which is
    // the check the video path makes before it spawns its own: a sound must not go on playing
    // over the sound of the file the pointer has moved to. The engine's own session is stopped
    // by `play_audio` rather than here, and this is what answers for the other engine — a
    // player that has not been confirmed gone, whose process handle the hover that started it
    // took with it when it ended.
    kill_stray_video_process();

    let volume = current_audio_volume();
    if volume == 0 {
        return true;
    }

    match track.player {
        Player::Native => {
            // A hover that lands on the file already playing leaves it playing, the same way
            // the FFmpeg path compares the file it last started.
            if video_player::playing_path().as_deref() == Some(path) && video_player::is_playing() {
                return true;
            }

            video_player::play_audio(path, volume, start);
            video_player::is_playing()
        }
        Player::Ffmpeg => {
            media.video_process = start_audio_player(path, volume, start);
            media.video_process.is_some()
        }
    }
}

/// Start FFmpeg's player on a sound: no window at all, which is the whole of what this side asks
/// of it, and the player's own volume scale — `0` to `100`.
///
/// A sound is looped while it is hovered, as a video is: what a hover is for is the file, and a
/// sound that stopped under a pointer that had not moved would be a preview that ended on its
/// own. The card's clock wraps with it (see `audio_clock`).
///
/// Where the sound starts is `-ss`, and it is an option of the *input* rather than of the
/// player: what it does is seek the file before anything of it is read, which is a player that
/// begins at that second rather than one that plays its way there. What it costs is nothing:
/// the seek is the player's own, and nothing is decoded before it.
///
/// A player given such a second is given no loop, and that is `ffplay`'s arrangement rather than
/// this side's: the position its own loop seeks back to *is* the position it was started at —
/// `-ss` is the one value that seek-back reads — so a sound dropped half way into a file would
/// go round the second half of it for as long as it was hovered. That pass is played once
/// instead, `-autoexit` being what ends it at the end of the file, and what this side starts in
/// its place plays the whole file and loops from the beginning of it (see
/// `wrap_audio_player`): every pass after the first goes back to 0:00 whatever the file is and
/// whatever the setting started it at. It is the same rule the engine Windows has is held to,
/// where `SetLoop` restarts the whole presentation rather than a seek into it — one answer for
/// both players, and this side's own hand under the half that has none.
///
/// A sound that starts at the beginning of its file is that second player already, so it is
/// given that player's own loop rather than a pass for this side to restart after.
fn start_audio_player(path: &Path, volume: u32, start: f64) -> Option<Child> {
    let mut command = Command::new("ffplay");
    command.args(["-nodisp", "-autoexit", "-loglevel", "quiet"]);
    command.args(["-volume", &volume.min(100).to_string()]);

    if start.is_finite() && start > 0.0 {
        command.args(["-ss", &format!("{start:.3}")]);
    } else {
        // The whole file, looped by the player itself: with no position given to it, the
        // position its loop returns to is the beginning of the file.
        command.args(["-loop", "0"]);
    }

    let child = command
        .arg(path)
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .creation_flags(engine_processes::CREATE_NO_WINDOW)
        .spawn()
        .ok()?;

    // The player is this app's own child, taken charge of the way every other one is: the job
    // ends it when the app does, and the record answers for a run that never got to end it. It
    // is recorded as the player rather than as an engine, because what ends it is its hover
    // ending.
    engine_processes::record_player(VIDEO_PROCESS_IMAGE_NAME, child.id());
    VIDEO_PID.store(child.id(), Ordering::SeqCst);
    VIDEO_HWND.store(0, Ordering::SeqCst);

    Some(child)
}

/// How long a player has to have lived before its exit is read as the end of the file it was
/// given, where the file's own length says nothing shorter than this.
///
/// A player that has stopped is either one that played its file through or one that never played
/// it at all — a machine with no output device, a decoder that will not have the file — and the
/// second of those stops within a moment of starting. Time is the only thing that tells the two
/// apart, and what it is spent on is the difference between a sound that goes round again and a
/// process spawned a tick apart for as long as the file is hovered. What this sits at is past
/// every failure a player of these files has and under the pass any file is hovered for, and it
/// is a ceiling on what is asked rather than the bar itself: a pass shorter than it — the last
/// moment of a short file, which `Random` can land in — is asked only to have been played (see
/// `reached_the_end`).
const AUDIO_WRAP_MINIMUM: f64 = 1.0;

/// Whether a player that has stopped is one that reached the end of the file it was handed.
///
/// What the player was given is the file from the second the sound was dropped in to its own
/// end, so the length the file says it has is what that pass takes — except that a length is a
/// container's own reading of a header and is a little out for some formats, which is why it is
/// used as a ceiling on what is asked for rather than as the answer itself: a stop is read as the
/// end of the file where the player lived for [`AUDIO_WRAP_MINIMUM`], or for the whole of a pass
/// shorter than that. A player cannot have lived past the end of the pass it was given, and the
/// last moment of a short file is a pass of its own.
fn reached_the_end(played: Duration, length: Option<f64>, offset: f64) -> bool {
    let pass = length
        .map(|length| (length - offset).max(0.0))
        .unwrap_or(AUDIO_WRAP_MINIMUM);

    played.as_secs_f64() >= pass.min(AUDIO_WRAP_MINIMUM)
}

/// Put a sound FFmpeg plays round to the beginning of its file where the player it was given has
/// reached the end of it.
///
/// The player this side starts for a sound dropped into the middle of a file plays that pass and
/// stops — see `start_audio_player` for why the loop cannot be the player's own — so the end of
/// the file is the player's exit, and what is started in its place is a player of the whole file
/// which loops from the beginning of it. Every pass after the first therefore goes back to 0:00,
/// whatever the file is and whatever the setting started it at.
///
/// What the card's clock is drawn from moves with the player, and that is the whole of what a
/// wrap is on this side: the moment the sound was put in at is replaced by the moment the new
/// player started, and the position the clock is counted from by the beginning of the file. A
/// player that reports nothing at all is a clock of this app's, and a clock still counted from
/// the second the *old* pass was dropped in at would have the card say the sound was a minute
/// into a file whose playing was heard to begin.
///
/// A player whose stop is not read as the end of its file by `reached_the_end` — one that never
/// played anything — is not started again: the card is left with no clock rather than with
/// another player, which is the answer a file this machine will not play gets.
fn wrap_audio_player(
    media: &mut MediaData,
    path: &Path,
    started: &mut Option<Instant>,
    offset: &mut f64,
) {
    let Some(process) = media.video_process.as_mut() else {
        return;
    };

    // The field holds the player this hover started and no other, and a process that has been
    // waited on is a process that has ended: a player that is still going is not one to replace.
    if matches!(process.try_wait(), Ok(None)) {
        return;
    }

    let played = started.map(|at| at.elapsed()).unwrap_or_default();
    let length = audio_track::playable(path).and_then(|track| track.duration);

    // The player is confirmed gone — a process that has been waited on has ended — so the record
    // of it goes with it rather than being left for the leftover-process sweep to find, exactly
    // as the end of a video's player is answered for (see `is_video_process_running`).
    let pid = process.id();
    media.video_process = None;
    VIDEO_HWND.store(0, Ordering::SeqCst);
    VIDEO_PID.store(0, Ordering::SeqCst);
    engine_processes::forget(pid);

    if !reached_the_end(played, length, *offset) {
        *started = None;
        return;
    }

    // The pass after this one is the whole file, started the way the hover started the player it
    // replaces: the same volume, the same care about a process left behind, and the same answer
    // as to whether a player arrived at all.
    start_audio_playback(path, media, 0.0);

    if media.video_process.is_some() {
        *started = Some(Instant::now());
        *offset = 0.0;
    } else {
        *started = None;
    }
}

/// Where the sound is and how long it is: the engine's own clock where Windows plays it, and
/// this app's clock over the player's start where FFmpeg does. The whole is the file's own
/// answer either way, and either half is nothing where there is nothing to say it — which is
/// what a card with no player behind it is drawn with.
///
/// A sound loops for as long as it is hovered, so what the clock says is where in the file the
/// sound is *now*: a player that has been going for longer than the file lasts is wrapped back
/// into it, which is what keeps the bar going round rather than standing full. The clock over
/// the player's start is counted from the second the sound was put in at as well — `from` —
/// because `ffplay` reports nothing at all: a file dropped half way into itself draws its clock
/// and its bar at its middle only if this side counts the first half as already played, and
/// what a hover would otherwise show is a sound playing from its middle with a card saying it
/// has just begun.
fn audio_clock(path: &Path, started: Option<Instant>, from: f64) -> (Option<f64>, Option<f64>) {
    let Some(track) = audio_track::playable(path) else {
        return (None, None);
    };

    match track.player {
        // The engine's own clock has the seek in it — it is the engine that was taken to where
        // the sound starts — so what it reports is the position with nothing added to it.
        Player::Native => (video_player::position(), video_player::duration().or(track.duration)),
        Player::Ffmpeg => {
            let elapsed = started.map(|at| from + at.elapsed().as_secs_f64());
            let position = match (elapsed, track.duration) {
                (Some(elapsed), Some(duration)) if duration > 0.0 => Some(elapsed % duration),
                (elapsed, _) => elapsed,
            };

            (position, track.duration)
        }
    }
}

/// Get original dimensions of media for positioning calculations
fn get_media_dimensions(path: &PathBuf) -> Option<(u32, u32)> {
    // What the file's content says it is comes ahead of what its name does, where the two
    // disagree: the box a file is placed at is the box of the kind its content belongs to,
    // and a format no kind previews is placed nowhere at all — see `content_type`.
    let content = CONFIG
        .lock()
        .ok()
        .map(|config| crate::formats::content_type::of(path, &config))
        .unwrap_or(crate::formats::content_type::Content::Unknown);

    match content {
        crate::formats::content_type::Content::Kind(kind) => {
            return media_dimensions_of_kind(kind, path)
        }
        crate::formats::content_type::Content::Foreign => return None,
        crate::formats::content_type::Content::Unknown => {}
    }

    if video_formats::is_video_preview(path) {
        return video_box(path);
    }

    // A PDF is measured from its own first page; one that cannot be read as a
    // PDF reports no dimensions, which drops the preview instead of guessing.
    if pdf_preview::is_pdf_preview(path) {
        return pdf_page_box(path);
    }

    // And a comic, whose page is a picture inside the container: it is measured where the hook asks
    // it, beside the PDF, because the two are the same kind of preview — a page of a book — and are
    // told apart by the reader rather than by the user. A box with no plate in it is measured as
    // nothing, which is the hover that shows no preview at all (see `comic_box`).
    if ebook_formats::is_ebook_preview(path) {
        return comic_box(path);
    }

    // An Office document is measured from the page that has been drawn for it — by Office
    // where its own application is installed, and by the render engine beside it where it is
    // not. A document with neither is measured as the page it is about to get — while a page
    // is coming, which is the only case where one is.
    if office_formats::is_office_preview(path) {
        return office_preview::measure(path);
    }

    // A document this app hands to a render engine is measured from the page that engine
    // drew, and a page not drawn yet is the wait for one: the layout places the spinner's
    // own box, the preview loop asks the engine for the document, and the hover is replayed
    // when the page lands (see `libre_render_is_due`). Nothing is converted here, and
    // nothing is waited on — a launch on this thread is a preview, a tray and a pointer
    // held for as long as the engine takes, which is what a document the engine cannot draw
    // never ends.
    //
    // It is asked where the hook asks it — after the office list, ahead of the design list
    // — because a name can sit in two lists: CorelDRAW is a design document to this app and
    // a drawing to the engine, and it is the engine that draws it (see `libre_formats`).
    if libre_formats::is_libre_preview(path) {
        return libre_box(path);
    }

    // And a picture the ImageMagick engine develops, measured where the hook asks it: after
    // the documents an engine draws, ahead of the design, vector, font and image lists, none
    // of which would have claimed a `.nef` anyway. What is measured is the picture the engine
    // wrote, and one it has not written yet is the wait for it (see `magick_box`).
    if magick_formats::is_magick_preview(path) {
        return magick_box(path);
    }

    // And a book the ebook engine converts, measured where the hook asks it: beside the listing
    // engine and the document engines, none of whose lists would have claimed a `.mobi` anyway.
    // What is measured is the page the engine wrote, and one it has not written yet is the wait
    // for it (see `calibre_box`).
    if calibre_formats::is_calibre_preview(path) {
        return calibre_box(path);
    }

    // A design document is measured from the picture it is previewed from — the merged
    // image at the end of a Photoshop file, or the picture a project container holds —
    // and a file neither reader will answer for reports no size, which is how it comes
    // to show nothing at all rather than a box nothing would be drawn into.
    //
    // It is asked ahead of the two kinds below it because the gate asks it ahead of them:
    // a name written into the design list as well as into the vector or font list is a
    // design document, and the two questions further down would report a size for another
    // kind than the one the tray was asked to switch.
    if design_formats::is_design_preview(path) {
        return design_dimensions(path);
    }

    // An SVG is measured from the document rather than from a header: the size it asks
    // to be drawn at is the size the layout places, and the engine draws it at whatever
    // box comes out of that. It is asked ahead of the readers of the other half of its
    // kind — the vector list names a document beside the metafiles, and neither of those
    // readers would take one — and ahead of the `Images` gate, which is the other list
    // that may have claimed it. A document is its own kind either way: what draws one is
    // not a decoder, and the switch for it is not the switch for pictures. A file of a
    // kind that is switched off reports no size, which is how the layout drops its
    // preview.
    if svg_preview::is_svg_file(path) {
        if !PreviewType::Vector.enabled() {
            return None;
        }

        // The engine is what draws a document, so a machine without one — or a spell
        // the engine has stood down for — has no document preview at all. Reporting no
        // size is what keeps a hover from opening a box nothing would be drawn into,
        // and it costs no read of the file.
        if !webview_preview::can_draw() {
            return None;
        }

        return svg_box(path);
    }

    // A vector drawing is measured from the records it holds: what an `.eps` keeps a
    // preview of, or what a metafile's own header declares its drawing to be.
    if vector_formats::is_vector_preview(path) {
        return vector_box(path);
    }

    // A font is measured at a box of this app's own rather than by anything the file says:
    // a font has no size it asks to be drawn at — what it holds is outlines — so the box is
    // the shape a specimen wants, and the share of the display `font_scale` names decides
    // how large that box is shown. `Fonts` is the gate, and a file that will not parse as a
    // font reports no size at all: that is how a `.ttf` that is something else comes to show
    // nothing rather than a page of another font's glyphs.
    if font_formats::is_font_preview(path) {
        if !webview_preview::can_draw() {
            return None;
        }

        return font_box(path);
    }

    // Whatever is left is a picture, so the `Images` gate is what decides it.
    if !PreviewType::Images.enabled() {
        return None;
    }

    picture_dimensions(path)
}

/// The box for one of the kinds a file's content answered with — see `content_type`.
///
/// It is the arm the chain above would have taken had the file been named what its content
/// says it is, and each arm is what that arm measures: the same reader, asked of the file
/// itself rather than of the name it is under. The gate is asked here as the chain asks it
/// per kind, so a kind switched off has no box and comes down the way it always does.
///
/// Two kinds have no size of their own: an archive listing and a text preview are both laid
/// out from the frame the loader paints rather than from anything their file says.
fn media_dimensions_of_kind(kind: PreviewType, path: &PathBuf) -> Option<(u32, u32)> {
    if !kind.enabled() {
        return None;
    }

    match kind {
        PreviewType::Videos => video_box(path),
        // A sound is measured against the room it is drawn in — the card is a page of text
        // wrapped to the box it is given — so it is asked where the bounds and the DPI are,
        // beside the drawn kinds and not here (see `media_dimensions`).
        PreviewType::Audio => None,
        PreviewType::Ebook => pdf_page_box(path),
        PreviewType::Archives | PreviewType::Text | PreviewType::Peazip => None,
        PreviewType::Document => office_preview::measure(path),
        PreviewType::Libre => libre_box(path),
        PreviewType::Magick => magick_box(path),
        PreviewType::Calibre => calibre_box(path),
        PreviewType::Design => design_dimensions(path),
        PreviewType::Vector => {
            if svg_preview::is_svg_file(path) {
                webview_preview::can_draw().then(|| svg_box(path)).flatten()
            } else {
                vector_box(path)
            }
        }
        PreviewType::Fonts => webview_preview::can_draw()
            .then(|| font_box(path))
            .flatten(),
        PreviewType::Images => picture_dimensions(path),
    }
}

/// The box a document the render engine draws is placed at: the page it has already drawn
/// for this version of the document, the wait for one that is on its way, and nothing at
/// all for a document the engine has turned down or for a machine with no engine to draw
/// one with.
fn libre_box(path: &Path) -> Option<(u32, u32)> {
    if !libreoffice_render::available() {
        // Nothing to draw it with, so there is nothing to show: a machine without the
        // engine shows no preview for these names rather than the thumbnail the file
        // carries, which is the whole reason the name is in this list.
        return None;
    }

    if let Some(page) = libreoffice_render::rendered_page(path) {
        return pdf_preview::page_dimensions(&page);
    }

    // A document the engine has already turned down is not one to wait for.
    if libreoffice_render::refused(path) {
        return None;
    }

    Some((office_preview::WAITING_BOX, office_preview::WAITING_BOX))
}

/// The box a book the ebook engine converts is placed at: the page it has already converted for
/// this version of the book, the wait for one that is on its way, and nothing at all for a book the
/// engine has turned down or for a machine with no engine to convert one with.
///
/// It is the shape `libre_box` has, and it is the same question: what is placed is a page the
/// engine wrote rather than a size the file asks for — a book holds no page of its own, which is
/// the whole reason it is handed to an engine — so the box is the page's own and, until there is
/// one, the spinner's. Nothing is converted here: a book the engine has answered nothing for yet is
/// the wait, and the loop asks for the page the moment there is a hover to ask for it (see
/// `request_calibre_render`).
fn calibre_box(path: &Path) -> Option<(u32, u32)> {
    if !calibre_render::available() {
        // Nothing to convert it with, so there is nothing to show: a machine without the engine
        // shows no preview for these names rather than the first page of the markup a `.fb2` is,
        // which is the whole reason the name is in this list.
        return None;
    }

    if let Some(page) = calibre_render::rendered_page(path) {
        return pdf_preview::book_page(&page).map(|book| book.size);
    }

    // A book the engine has already turned down is not one to wait for.
    if calibre_render::refused(path) {
        return None;
    }

    Some((office_preview::WAITING_BOX, office_preview::WAITING_BOX))
}

/// The box a picture the ImageMagick engine develops is placed at: the size the engine
/// developed it at, the wait for one that is on its way, and nothing at all for a file the
/// engine has turned down or for a machine with no engine to develop one with.
///
/// The size is the picture's own — read from the header of the bytes the engine wrote, which
/// is the size the preview is drawn at under the picture scale — and it is what the layout
/// places and what the frame is keyed by. Nothing is converted here: a file the engine has
/// developed nothing for yet is the wait, and the loop asks for the picture the moment there
/// is a hover to ask for it (see `request_magick_render`).
///
/// The wait is a spinner's box and says nothing about how large the picture will be drawn,
/// which is why it is not the room the engine is asked for: a picture is only this size
/// *after* the conversion, so the room it is developed in is a ceiling on the size its
/// preview can ever be drawn at (see `PendingLoad::room`).
fn magick_box(path: &Path) -> Option<(u32, u32)> {
    if !imagemagick_render::available() {
        // Nothing to develop it with, so there is nothing to show: a machine without the
        // engine shows no preview for these names rather than the picture the camera left
        // inside the file, which is what the shell's own thumbnail is for.
        return None;
    }

    if let Some(size) = imagemagick_render::dimensions(path) {
        return Some(size);
    }

    // A file the engine has already turned down is not one to wait for.
    if imagemagick_render::refused(path) {
        return None;
    }

    // A raw sample dump is the one file the engine cannot measure for itself: what shape it is
    // comes from its own length rather than from a conversion, which is arithmetic this side can
    // do — and a dump whose length does not settle a shape is a file with no preview rather than
    // one to wait for (see `raw_geometry`).
    if imagemagick_render::is_raw_sample(path) {
        return imagemagick_render::raw_geometry(path);
    }

    Some((office_preview::WAITING_BOX, office_preview::WAITING_BOX))
}

/// A picture's own size, read the way this app reads one: its own header, and the codec
/// Windows keeps for the picture formats this app's decoder has no reader for.
fn picture_dimensions(path: &PathBuf) -> Option<(u32, u32)> {
    image_dimensions_with_header_check(path)
}

/// Whether this hover is the wait for a page rather than a preview of one: an Office
/// document, a document the render engine draws or a book the ebook engine converts, with
/// nothing drawn for it yet and a page on the way.
///
/// A hover like that is placed by the spinner's own box, a pointer gap off the hand,
/// rather than by the size a preview would take: it is the wait for the file under the
/// hand, which belongs at the hand. The page is laid out again by the replay that arrives
/// with it, so nothing here has to guess how large it will be.
fn page_is_on_the_way(path: &Path) -> bool {
    // A file whose bytes are another kind is not a document a page is coming for, which is
    // the same question the render tier is asked before it is asked for one: a picture left
    // under a document's name is drawn here, and placing it at the pointer as the wait for a
    // page would be a preview waiting for nothing (see `content_names_another_kind`).
    let office = office_formats::is_office_preview(path)
        && !content_names_another_kind(path, PreviewType::Document)
        && matches!(
            office_preview::source_kind(path),
            office_preview::SourceKind::None
        );

    office
        || libre_render_is_due(path)
        || magick_render_is_due(path)
        || peazip_render_is_due(path)
        || calibre_render_is_due(path)
}

/// The size the layout should place and scale a preview from.
///
/// A text file has no size of its own, so the box its first screenful wants is
/// measured here, bounded by the display it will be shown on. That makes the
/// measurement an intrinsic size in the same sense a PDF page's is: the layout
/// can fit it into the space beside the cursor, and the text renderer is handed
/// the box that comes out of that.
fn media_dimensions(path: &PathBuf, bounds: ScreenBounds, dpi: u32) -> Option<(u32, u32)> {
    // What the file's own bytes say it is comes first, as it does for the loader that draws
    // it and for the share it is laid out at: a `.txt` whose bytes are a picture is measured
    // as the picture it is rather than read as a page of text it is not — which for a file
    // whose bytes are not text is no measurement at all, and a preview that never appears
    // for a file that would otherwise be drawn.
    let content = CONFIG
        .lock()
        .ok()
        .map(|config| crate::formats::content_type::of(path, &config))
        .unwrap_or(crate::formats::content_type::Content::Unknown);

    if let crate::formats::content_type::Content::Kind(kind) = content {
        return match kind {
            // The three kinds measured against the room they are drawn in, which is a question
            // this side has the answer to and `media_dimensions_of_kind` does not.
            PreviewType::Text => text_box(path, bounds, dpi),
            PreviewType::Archives => archive_box_off_the_tick(path, bounds, dpi),
            PreviewType::Peazip => peazip_box(path, bounds, dpi),
            // And the fourth kind that is wrapped to its room: a sound's card is painted at a
            // fixed font size and cut to the box it is given, and what stands in for it until
            // the probe beside it has answered is the waiting spinner.
            PreviewType::Audio => audio_box(path, bounds, dpi),
            _ => media_dimensions_of_kind(kind, path),
        };
    }

    if is_text_preview(path) {
        return text_box(path, bounds, dpi);
    }

    if archive_formats::is_archive_preview(path) {
        return archive_box_off_the_tick(path, bounds, dpi);
    }

    // And an archive an engine lists, measured where the hook asks it: beside the archive list
    // above it, which is where the two are told apart — a name in that list is read by this app
    // itself, and one in this list is read by an engine. What is measured is the page the
    // engine's listing makes, and a listing that has not come back yet is the wait for one (see
    // `peazip_box`).
    if peazip_formats::is_peazip_preview(path) {
        return peazip_box(path, bounds, dpi);
    }

    // A sound is the fourth kind measured against its room, and the last: what a hover on one
    // asks is a card whose facts a probe has to bring back first (see `audio_box`).
    if drawn_as_audio(path) {
        return audio_box(path, bounds, dpi);
    }

    get_media_dimensions(path)
}

/// The box a page of text asks for: as many lines and columns as the room the display has
/// holds, at the font size its DPI gives them.
fn text_box(path: &Path, bounds: ScreenBounds, dpi: u32) -> Option<(u32, u32)> {
    let cap_width = (bounds.right - bounds.left).max(1) as u32;
    let cap_height = bounds.height().max(1) as u32;

    text_preview::measure(path, cap_width, cap_height, dpi, current_text_options())
}

/// And the box a listing asks for, which is the same shape of question asked of an archive's
/// own table of contents.
fn archive_box(path: &Path, bounds: ScreenBounds, dpi: u32) -> Option<(u32, u32)> {
    let cap_width = (bounds.right - bounds.left).max(1) as u32;
    let cap_height = bounds.height().max(1) as u32;

    archive_preview::measure(path, cap_width, cap_height, dpi, current_archive_options())
}

/// The box an archive the PeaZip engine lists is placed at: the box the same listing would be
/// measured at had this app read the archive itself, the spinner's own box while the engine has
/// not answered yet, and nothing at all for a file the engine has turned down or for a machine
/// with no engine to list it.
///
/// The measurement is the archive page's own, and it is the same one either way: a listing is a
/// listing, and where it came from is not something the layout is told (see
/// `archive_listing::listing_for`). What is waited for is a page that does not exist yet — the box
/// an engine's answer has not arrived for says nothing about how large it will be — so it is
/// placed as the spinner, at the pointer's own corner, and laid out again by the replay that
/// arrives with the answer.
fn peazip_box(path: &Path, bounds: ScreenBounds, dpi: u32) -> Option<(u32, u32)> {
    if !peazip_render::available_for(path) {
        // Nothing to list it with, so there is nothing to show: a machine without the engine — or
        // with an installation that does not carry the one tool this name is read by — shows no
        // preview for it rather than a page of something else.
        return None;
    }

    if let Some(size) = archive_box(path, bounds, dpi) {
        return Some(size);
    }

    // A file the engine has already turned down is not one to wait for.
    if peazip_render::refused(path) {
        return None;
    }

    Some((office_preview::WAITING_BOX, office_preview::WAITING_BOX))
}

/// A text preview's box, placed for the width it came out with.
///
/// The box a text preview is placed from is measured at the width the *display* can
/// give, and the layout fits that box beside the cursor or the focused item by
/// shrinking it whole. Text does not shrink with it: the lines wrap at the width the
/// box ends up with, and a narrower box needs *more* rows than the proportional
/// height leaves — a long line that took two rows at the display's width takes three
/// in the box that came back — so the frame could show the first of them and the
/// rest of the line was cut off below it. Measuring the document again at the width
/// the box actually has is what gives it the height the text really takes there, and
/// placing the result again is the same rule that put it there the first time.
///
/// The size that measure answered with is handed back with the layout, because the
/// wait that follows the pointer re-places from the size it was given as well: left
/// at the display's own measurement it would step around the name for a height the
/// wrapped frame does not have (see `HoverPlacement`).
fn text_preview_layout(
    path: &Path,
    layout: PreviewLayout,
    dpi: u32,
    place: impl FnOnce((u32, u32)) -> Option<PreviewLayout>,
) -> (PreviewLayout, Option<(u32, u32)>) {
    if !is_text_preview(path) {
        return (layout, None);
    }

    let Some(size) = text_preview::measure(
        path,
        layout.preview_w,
        layout.max_height,
        dpi,
        current_text_options(),
    ) else {
        return (layout, None);
    };

    match place(size) {
        Some(placed) => (placed, Some(size)),
        None => (layout, None),
    }
}

/// The effective DPI of the display nearest `(x, y)`, which is what a text preview's
/// font size is scaled by — and what every margin a layout is written around is
/// scaled by (see `logical_px`). Falls back to the 96 DPI baseline when no display
/// can be named, the same way the placement falls back to the primary display.
///
/// The display is what is asked, not the window the point happens to be over: a
/// window carries the scale its own process was told about — a UWP one can answer a
/// scale that is not the display's at all — while the display under the point is one
/// question with one answer, whatever is drawn on it.
///
/// The answer is kept per display, because the scale of a display does not change while
/// it is the display: the pointer's own probe asks this every tick and every layout asks
/// it again, and a pointer that has not crossed to another display is answered from here
/// rather than by asking the DPI interface for a number that cannot have moved. A display
/// whose scale does change — a monitor switched to another scaling — is a display whose
/// handle is the same and whose answer is not, so the cache holds one display: the next
/// one named is asked about, and the one after that is asked again.
pub(crate) fn monitor_dpi_from_point(x: i32, y: i32) -> u32 {
    const BASELINE_DPI: u32 = 96;

    unsafe {
        let monitor = MonitorFromPoint(POINT { x, y }, MONITOR_DEFAULTTONEAREST);
        if monitor.is_invalid() {
            return BASELINE_DPI;
        }

        let handle = monitor.0 as isize;
        if let Ok(cached) = MONITOR_DPI.lock() {
            if let Some((cached_monitor, dpi)) = *cached {
                if cached_monitor == handle {
                    return dpi;
                }
            }
        }

        let mut dpi = BASELINE_DPI;
        let mut dpi_x = 0u32;
        let mut dpi_y = 0u32;
        if GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, &mut dpi_x, &mut dpi_y).is_ok() && dpi_x > 0
        {
            dpi = dpi_x;
        }

        if let Ok(mut cached) = MONITOR_DPI.lock() {
            *cached = Some((handle, dpi));
        }

        dpi
    }
}

/// The last display asked about and the scale it answered with — one display's worth, for
/// the reason `monitor_dpi_from_point` gives.
static MONITOR_DPI: Lazy<Mutex<Option<(isize, u32)>>> = Lazy::new(|| Mutex::new(None));

/// Render a single frame of the loading spinner animation (BGRA pixels).
///
/// The frame is the arc and nothing else: the box it is drawn in is transparent,
/// so what a hover that is waiting shows is a spinner rather than a square of its
/// own. The arc is white, with a soft dark halo drawn under it, because a preview
/// goes over whatever Explorer happens to be drawing — the halo is what keeps the
/// spinner visible over a light background, and the white arc is what keeps it
/// visible over a dark one.
fn render_loading_frame(width: u32, height: u32, angle: f32) -> Vec<u8> {
    let total_pixels = (width as usize) * (height as usize);
    let mut pixels = vec![0u8; total_pixels * 4];

    let cx = width as f32 / 2.0;
    let cy = height as f32 / 2.0;

    // Spinner proportional to window size, clamped for aesthetics
    let radius = (width.min(height) as f32 * 0.08).clamp(10.0, 32.0);
    let thickness = (radius * 0.32).clamp(2.5, 7.0);
    // How far the halo reaches past the arc, and how much of it there is where it
    // meets the arc's own edge. Both are drawn from the arc's size, so a large
    // spinner gets a halo in proportion.
    let halo_width = thickness;
    let halo_opacity = 0.85;
    // What the halo is made of: dark enough to read over a light file list, and
    // light enough to read as a shadow rather than a second arc.
    let halo_shade = 12.0;

    let two_pi = std::f32::consts::PI * 2.0;
    let arc_length = std::f32::consts::PI * 1.5; // 270-degree arc

    // Only iterate over the bounding box of the spinner's own ring and halo
    let reach = radius + thickness + halo_width + 1.0;
    let min_x = ((cx - reach).max(0.0)) as u32;
    let max_x = ((cx + reach).min(width as f32 - 1.0)) as u32;
    let min_y = ((cy - reach).max(0.0)) as u32;
    let max_y = ((cy + reach).min(height as f32 - 1.0)) as u32;

    for y in min_y..=max_y {
        for x in min_x..=max_x {
            let dx = x as f32 + 0.5 - cx;
            let dy = y as f32 + 0.5 - cy;
            let dist = (dx * dx + dy * dy).sqrt();

            let ring_dist = (dist - radius).abs();
            if ring_dist > thickness + halo_width {
                continue;
            }

            let pixel_angle = dy.atan2(dx);
            let relative = (pixel_angle - angle).rem_euclid(two_pi);
            if relative > arc_length {
                continue;
            }

            // Smooth gradient: ease-in from tail (transparent) to head (bright)
            let t = relative / arc_length;
            let t_smooth = t * t; // quadratic ease-in

            // Anti-aliased smooth edge, and the halo under it: full where the arc
            // covers it and fading out from the arc's edge, which is what leaves a
            // dark rim past an arc that is white.
            let arc = (1.0 - (ring_dist - thickness + 1.0).max(0.0)).clamp(0.0, 1.0) * t_smooth;
            let halo = (1.0 - (ring_dist - thickness).max(0.0) / halo_width).clamp(0.0, 1.0)
                * halo_opacity
                * t_smooth;

            // The arc over the halo, over nothing at all: the alpha is how much of
            // the two there is, and the colour is what is left of them once that is
            // known — white where the arc covers, the halo's shade where only the
            // halo does.
            let alpha = arc + halo * (1.0 - arc);
            if alpha <= 0.0 {
                continue;
            }
            let shade = ((255.0 * arc + halo_shade * halo * (1.0 - arc)) / alpha).clamp(0.0, 255.0);

            let idx = ((y * width + x) * 4) as usize;
            let shade = shade as u8;
            pixels[idx] = shade; // B
            pixels[idx + 1] = shade; // G
            pixels[idx + 2] = shade; // R
            pixels[idx + 3] = (alpha.clamp(0.0, 1.0) * 255.0) as u8;
        }
    }

    pixels
}

/// Create a loading animation MediaData for the given dimensions
fn create_loading_media(width: u32, height: u32) -> MediaData {
    let pixels = render_loading_frame(width, height, 0.0);
    let frame = ImageFrame::new(pixels, width, height, 33);
    MediaData {
        frames: vec![frame],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::Loading,
        stream_cancel: None,
        video_process: None,
        loading_start: Some(Instant::now()),
        text_state: None,
    }
}

/// The spinner's own geometry, in the frame's coordinates: how large the ring is, how thick
/// it is, and how far its corner sits from the frame's own.
const SPINNER_OVERLAY_RADIUS: f32 = 8.0;
const SPINNER_OVERLAY_THICKNESS: f32 = 2.5;
const SPINNER_OVERLAY_PADDING: f32 = 12.0;

/// The box the corner spinner is drawn in: where it sits in the frame, and how large a box
/// holds it whole.
///
/// One answer for the copying and for the drawing, which have to agree: what the spinner is
/// drawn into is exactly what was copied out of the frame (see `overlay_loading_spinner`).
/// `None` for a frame too small to hold one at all, which is a frame the spinner is not drawn
/// on.
fn spinner_overlay_box(width: u32, height: u32) -> Option<FrameBox> {
    if width < 24 || height < 24 {
        return None;
    }

    // How far the halo reaches past the centre, and one pixel more for the edge it fades
    // out over: the same reach the drawing below walks, so the box is never smaller than
    // what is drawn in it.
    let reach = SPINNER_OVERLAY_RADIUS + SPINNER_OVERLAY_THICKNESS + 4.0 + 1.0;
    let cx =
        width as f32 - SPINNER_OVERLAY_PADDING - SPINNER_OVERLAY_RADIUS - SPINNER_OVERLAY_THICKNESS;
    let cy = height as f32
        - SPINNER_OVERLAY_PADDING
        - SPINNER_OVERLAY_RADIUS
        - SPINNER_OVERLAY_THICKNESS;

    let left = ((cx - reach).max(0.0)) as u32;
    let top = ((cy - reach).max(0.0)) as u32;
    let right = ((cx + reach).min(width as f32 - 1.0)) as u32;
    let bottom = ((cy + reach).min(height as f32 - 1.0)) as u32;

    Some(FrameBox {
        left,
        top,
        width: right - left + 1,
        height: bottom - top + 1,
    })
}

/// Render a small loading spinner overlay onto a box of BGRA pixels copied out of a frame, in
/// place: a spinning arc in the bottom-right corner with a semi-transparent dark backdrop
/// circle.
///
/// The pixels are the box's and the geometry is the frame's — the corner the arc sits in is a
/// padding off the frame's own bottom-right, not off the box's — which is what keeps a spinner
/// drawn into a corner of a frame the same spinner it was when the whole frame was copied to
/// draw it (see `spinner_overlay_box`).
fn overlay_loading_spinner(
    pixels: &mut [u8],
    area: FrameBox,
    frame_width: u32,
    frame_height: u32,
    angle: f32,
) {
    if pixels.len() < area.bytes() {
        return;
    }

    let (left, top) = (area.left, area.top);
    let (width, height) = (area.width, area.height);

    let radius = SPINNER_OVERLAY_RADIUS;
    let thickness = SPINNER_OVERLAY_THICKNESS;
    let backdrop_r = radius + thickness + 4.0;

    // Center of the spinner in the bottom-right corner of the frame
    let cx = frame_width as f32 - SPINNER_OVERLAY_PADDING - radius - thickness;
    let cy = frame_height as f32 - SPINNER_OVERLAY_PADDING - radius - thickness;

    let min_x = (((cx - backdrop_r - 1.0).max(0.0)) as u32).max(left);
    let max_x =
        (((cx + backdrop_r + 1.0).min(frame_width as f32 - 1.0)) as u32).min(left + width - 1);
    let min_y = (((cy - backdrop_r - 1.0).max(0.0)) as u32).max(top);
    let max_y =
        (((cy + backdrop_r + 1.0).min(frame_height as f32 - 1.0)) as u32).min(top + height - 1);

    let two_pi = std::f32::consts::PI * 2.0;
    let arc_length = std::f32::consts::PI * 1.5;

    for y in min_y..=max_y {
        for x in min_x..=max_x {
            let dx = x as f32 + 0.5 - cx;
            let dy = y as f32 + 0.5 - cy;
            let dist = (dx * dx + dy * dy).sqrt();
            let idx = (((y - top) * width + (x - left)) * 4) as usize;
            if idx + 3 >= pixels.len() {
                continue;
            }

            // Semi-transparent dark backdrop circle
            if dist <= backdrop_r {
                let edge = (1.0 - (dist - backdrop_r + 1.0).max(0.0)).clamp(0.0, 1.0);
                let bg_alpha = 0.45 * edge;
                if bg_alpha > 0.0 {
                    pixels[idx] = ((pixels[idx] as f32) * (1.0 - bg_alpha)) as u8;
                    pixels[idx + 1] = ((pixels[idx + 1] as f32) * (1.0 - bg_alpha)) as u8;
                    pixels[idx + 2] = ((pixels[idx + 2] as f32) * (1.0 - bg_alpha)) as u8;
                }
            }

            // Spinner ring
            let ring_dist = (dist - radius).abs();
            if ring_dist > thickness + 1.0 {
                continue;
            }
            let edge_alpha = (1.0 - (ring_dist - thickness + 1.0).max(0.0)).clamp(0.0, 1.0);
            if edge_alpha <= 0.0 {
                continue;
            }
            let pixel_angle = dy.atan2(dx);
            let relative = (pixel_angle - angle).rem_euclid(two_pi);
            if relative <= arc_length {
                let t = relative / arc_length;
                let t_smooth = t * t;
                let alpha = edge_alpha * t_smooth;
                let blend = |bg_c: u8, fg: u8, a: f32| -> u8 {
                    ((bg_c as f32) * (1.0 - a) + (fg as f32) * a).clamp(0.0, 255.0) as u8
                };
                pixels[idx] = blend(pixels[idx], 255, alpha);
                pixels[idx + 1] = blend(pixels[idx + 1], 255, alpha);
                pixels[idx + 2] = blend(pixels[idx + 2], 255, alpha);
            }
        }
    }
}

/// Result from background image loading thread
struct LoadResult {
    generation: u64,
    path: PathBuf,
    media: Option<MediaData>,
    /// Nothing to draw yet, and a page on the way: the preview stays pending
    /// rather than being dropped, so the page has something to replace.
    awaiting_render: bool,
}

/// A decode request consumed by the dedicated loader worker.
struct LoadRequest {
    generation: u64,
    path: PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    dpi: u32,
    cancel: Arc<AtomicBool>,
}

type LoadRequestSlot = Arc<(Mutex<Option<LoadRequest>>, Condvar)>;

fn queue_load_request(slot: &LoadRequestSlot, request: LoadRequest) {
    let (lock, cvar) = &**slot;
    if let Ok(mut pending) = lock.lock() {
        *pending = Some(request);
        cvar.notify_one();
    }
}

fn clear_load_request(slot: &LoadRequestSlot) {
    let (lock, _) = &**slot;
    if let Ok(mut pending) = lock.lock() {
        *pending = None;
    }
}

fn spawn_load_worker(
    request_slot: LoadRequestSlot,
    result_tx: Sender<LoadResult>,
) -> std::thread::JoinHandle<()> {
    std::thread::spawn(move || {
        // A PDF page is rendered through Windows.Data.Pdf and a picture of a format
        // this app has no decoder for is decoded by the codec Windows has, so this
        // thread needs an apartment before the first load asks for either.
        pdf_preview::initialize_apartment();
        wic_image::initialize_apartment();

        while RUNNING.load(Ordering::Acquire) {
            let mut request = {
                let (lock, cvar) = &*request_slot;
                let mut pending = match lock.lock() {
                    Ok(guard) => guard,
                    Err(_) => break,
                };

                while pending.is_none() && RUNNING.load(Ordering::Acquire) {
                    pending = match cvar.wait_timeout(pending, Duration::from_millis(200)) {
                        Ok((guard, _)) => guard,
                        Err(_) => return,
                    };
                }

                if !RUNNING.load(Ordering::Acquire) {
                    break;
                }

                match pending.take() {
                    Some(req) => req,
                    None => continue,
                }
            };

            // Coalesce any queued requests so we decode only the newest target.
            {
                let (lock, _) = &*request_slot;
                if let Ok(mut pending) = lock.lock() {
                    if let Some(newer) = pending.take() {
                        request.cancel.store(true, Ordering::Release);
                        request = newer;
                    }
                } else {
                    break;
                }
            }

            if request.cancel.load(Ordering::Acquire) {
                continue;
            }

            let media = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                load_media(
                    &request.path,
                    request.max_width,
                    request.max_height,
                    request.preview_scale,
                    request.dpi,
                    Arc::clone(&request.cancel),
                )
            }))
            .unwrap_or(None);

            // Every engine a hover can be waiting on is asked here, and a kind left out is
            // not merely a wait that is not shown: the loop reads this as "there is nothing
            // to draw and nothing coming", which is the branch that hides the window and
            // drops the hover — so the engine is never asked for its answer and the file
            // has no preview at all. The listing the PeaZip engine owes is such a wait, and so
            // is the page the ebook engine converts a book into (see `peazip_render_is_due` and
            // `calibre_render_is_due`).
            let awaiting_render = media.is_none()
                && (office_render_is_due(&request.path, request.max_width)
                    || libre_render_is_due(&request.path)
                    || magick_render_is_due(&request.path)
                    || peazip_render_is_due(&request.path)
                    || calibre_render_is_due(&request.path));

            let _ = result_tx.send(LoadResult {
                generation: request.generation,
                path: request.path.clone(),
                media,
                awaiting_render,
            });
        }
    })
}

/// How long a hover's load may run before the spinner is put up for it, as
/// `spinner_delay_ms` in `config.ini` names it.
///
/// Read when a load starts rather than captured once, so an edit applies to the next
/// hover without a restart. The window is hidden while a load runs, so a load that
/// finishes inside the delay has gone straight from nothing to the preview — what the
/// delay is for — and `0` is a spinner that goes up with the load; see `spinner_due`.
fn load_spinner_delay() -> Duration {
    let millis = CONFIG
        .lock()
        .map(|config| sanitize_spinner_delay_ms(config.spinner_delay_ms))
        .unwrap_or(DEFAULT_SPINNER_DELAY_MS);

    Duration::from_millis(millis)
}

/// What placing a hover's preview again needs, kept on a pending load.
///
/// The size is not measured again when the preview follows the pointer: a cursor
/// moving along the item it belongs to finds the same media, and measuring it per
/// tick would re-read a header, a listing or a document sixty times a second to
/// learn what the hover already knew. What is recomputed is the place — that is
/// what the cursor decides.
///
/// A text preview is the one size measured again before the wait starts — its height
/// is the rows its lines wrap into at the width the box came out with — and what is
/// kept here is the size that measure answered with, so a re-placement steps around
/// the name for the rows the frame really has (see `text_preview_layout`).
#[derive(Clone, Copy)]
struct HoverPlacement {
    orig_dims: (u32, u32),
    avoid: Option<ScreenRegion>,
    follow_cursor: bool,
    preview_scale: PreviewScale,
    /// Whether this placement is the waiting spinner's own box rather than a
    /// preview's: it is then placed at the pointer's own corner — a pointer gap off
    /// the cursor, in the quadrant the display has room for it, with the name it
    /// covers left alone — and kept there while the wait runs. See
    /// `compute_mouse_layout` and `waiting_placement`.
    at_the_pointer_corner: bool,
}

/// The placement a hover's waiting spinner is given: the arc's own box at the pointer's
/// own corner — a pointer gap off the hand — whatever the preview it is waiting for is.
///
/// It is the placement a document waiting on a page has always been placed by, and
/// every other kind of load is shown the same way: what a hover shows while it waits
/// is the arc at the hand that hovered the file — which is what says the wait is for
/// the file under it — rather than a preview-sized frame with a spinner in the middle
/// of it, placed where the preview will land and saying nothing about the hand.
///
/// The gap is what keeps the arc off the cursor. The window it is drawn in is the one
/// the pointer's own messages land on, and a hand kept off the arc is a hand still
/// clicking and probing the file the wait is for.
fn waiting_placement(placement: HoverPlacement) -> HoverPlacement {
    HoverPlacement {
        orig_dims: (office_preview::WAITING_BOX, office_preview::WAITING_BOX),
        preview_scale: PreviewScale::Percent(100),
        at_the_pointer_corner: true,
        ..placement
    }
}

/// What placing a pending load again moved: the spinner's own box, which is what is
/// on screen while the load runs, and the preview's, which is what the media lands at
/// and what an engine drawing it is told to draw in.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
struct Followed {
    spinner: bool,
    preview: bool,
}

/// Tracks a pending background load so we can show the spinner while it runs.
struct PendingLoad {
    generation: u64,
    /// The hide count this load was started under. A hide after it is the pointer
    /// having left the file while the load ran, and is what the load's reveal is
    /// checked against so the frame it lands with is dropped rather than shown
    /// (see `HIDDEN_EPOCH`).
    hide_epoch: u64,
    /// The file this load is for, which is what decides whether the engine plays it
    /// and what the engine is pointed at.
    path: PathBuf,
    started: Instant,
    pos_x: i32,
    pos_y: i32,
    width: u32,
    height: u32,
    /// The room the display the hover is on has (`ScreenBounds::room`), which is the box
    /// a page on its way is asked for in and the box a picture the image converter
    /// develops is developed at. A slide is exported at the width the render is asked
    /// for, so the room the display has is the sharpest page that display can show; and
    /// a picture is developed *at* the size it is then drawn at, so a room smaller than
    /// that is a picture that can never be shown any larger.
    ///
    /// It is deliberately not the room this load's own layout came out at. A hover
    /// waiting on an engine is laid out as the spinner's own box at the pointer (see
    /// `waiting_placement`), and the room that layout comes out at is the corner of the
    /// display the spinner was put in — a box that says nothing about how large the
    /// preview will be drawn, and for a picture a ceiling that would leave it thumbnail
    /// sized for good.
    room: (u32, u32),
    spinner_shown: bool,
    /// How long this load may run before the spinner is put up for it, read from
    /// `spinner_delay_ms` when the load started. Every kind of wait is given the same
    /// one — a decode, a page Office is rendering, a browser that has to start.
    spinner_delay: Duration,
    /// Where the spinner is drawn and placed while this load runs — the arc's own box
    /// at the pointer's corner (see `waiting_placement`) — and not the box the preview
    /// will arrive in (`pos_x`, `pos_y`, `width`, `height`).
    spinner_pos: (i32, i32),
    spinner_side: u32,
    /// The mouse hover this load came from, if it was one. A preview that is still
    /// on its way follows the pointer, so it is placed again for a cursor that has
    /// moved along the item since; a keyboard hover's placement belongs to the
    /// item and carries none.
    placement: Option<HoverPlacement>,
    /// Whether this load is replacing what is already on screen — the page that
    /// arrived for the hover that is up — rather than opening a new preview. An
    /// upgrade never shows the spinner: what is there stays where it is, at its own
    /// size, until the page is ready.
    upgrade: bool,
    /// Whether this load is waiting on an engine to produce what no reader here could
    /// draw: the page Office is rendering, the page the render engine beside it is
    /// converting, the picture the image converter is developing, the listing Peazip is
    /// printing. Set where a load comes back with nothing to draw and one of those
    /// engines is owed the file (see `awaiting_render` in the loader).
    ///
    /// It is what the cap on waiting is read against, and it is on the wait rather than
    /// on the request for a reason: an engine answers a file once, so a hover whose
    /// request was folded into one already in flight — the file the engine is working on
    /// is the file this hover asked about — has nothing left that would ever answer it,
    /// and a wait that is not bounded is a spinner that runs for good (see
    /// `OFFICE_RENDER_WAIT_SECS`).
    ///
    /// A video's probe is the other wait marked this way, and for the reason above rather
    /// than a reason of its own: what it waits on is outside this side, and a probe that
    /// answered nothing would leave the hover with nothing here that ends it. The probe is
    /// given a bound of its own (`VIDEO_PROBE_TIMEOUT_SECS`), so the cap behind this is
    /// the second line rather than the first.
    awaiting_engine: bool,
}

impl PendingLoad {
    /// Whether the spinner is due for this load: once it has run for the delay
    /// `spinner_delay_ms` names, which is the same moment for every kind of wait — a
    /// decode, a page Office is rendering, a browser that has to start.
    ///
    /// An upgrade is never due — what is on screen stays where it is, at its own
    /// size, until the page replaces it — and a load that already has its spinner
    /// up is not due again.
    fn spinner_due(&self) -> bool {
        !self.spinner_shown && !self.upgrade && self.started.elapsed() >= self.spinner_delay
    }

    /// Place this load's preview — and the spinner standing in for it — again for
    /// `cursor`, when it is one that follows the pointer: the size the hover measured,
    /// its `Avoid` region and its scale, and the display the pointer is on now — the
    /// room it has (`dpi` is that display's scale, and `cursor` is where it is).
    /// Answers what moved, so the window is only moved when the spinner's own box has
    /// and an engine is only told when the preview's has.
    ///
    /// The two are not the same place. A wait is the arc's own box at the pointer's
    /// corner whatever the preview is (see `waiting_placement`), so a wait that
    /// follows a moving hand is the wait at the hand, while the preview's place is
    /// read for the media that lands at it and for an engine told where to draw.
    ///
    /// A load that answered nothing, or a keyboard hover, is left where it is.
    fn follow_pointer(&mut self, cursor: POINT, dpi: u32) -> Followed {
        let mut followed = Followed {
            spinner: false,
            preview: false,
        };
        let Some(placement) = self.placement else {
            return followed;
        };

        let bounds = monitor_bounds_from_point(cursor.x, cursor.y);

        // The room an engine is asked for follows the hand the preview does, because the
        // display the hand is on is the one the hover is now waiting on: a pointer that has
        // crossed to another display is a picture developed for that display's room.
        self.room = bounds.room();

        if let Some(layout) = compute_mouse_layout(cursor.x, cursor.y, placement, bounds, dpi) {
            followed.preview = (layout.pos_x, layout.pos_y) != (self.pos_x, self.pos_y)
                || (layout.preview_w, layout.preview_h) != (self.width, self.height);
            if followed.preview {
                self.pos_x = layout.pos_x;
                self.pos_y = layout.pos_y;
                self.width = layout.preview_w;
                self.height = layout.preview_h;
            }
        }

        let waiting = waiting_placement(placement);
        if let Some(layout) = compute_mouse_layout(cursor.x, cursor.y, waiting, bounds, dpi) {
            followed.spinner = (layout.pos_x, layout.pos_y) != self.spinner_pos
                || layout.preview_w != self.spinner_side;
            if followed.spinner {
                self.spinner_pos = (layout.pos_x, layout.pos_y);
                self.spinner_side = layout.preview_w;
            }
        }

        followed
    }
}

/// Whether an engine's answer belongs to the wait under the spinner even though it names
/// another hover: the answer is about the file that wait is on, it is an answer that
/// produced what that wait is for, and the wait is one an engine owes the file.
///
/// An engine answers a file once — a request made while it was already working on the
/// same file is folded into that work, and what comes back names the request before it.
/// Read as an answer about a hover that has gone, it would leave the wait with no request
/// that could answer it and no cap that could end it, and the page, picture or listing it
/// names — which is in hand — would not be shown until the file was hovered afresh. It is
/// what `hovered` is not: an answer that is the wait's own without naming its generation
/// (see `awaiting_engine`).
fn answer_belongs_to_the_wait(
    ready_path: &Path,
    ready_ok: bool,
    shown: Option<&Path>,
    pending: Option<&PendingLoad>,
) -> bool {
    ready_ok
        && shown == Some(ready_path)
        && pending.is_some_and(|pl| pl.awaiting_engine && pl.path == ready_path)
}

/// Reusable layered-window surface: one memory DC with one DIB section selected
/// into it, replaced only when the preview dimensions change.
///
/// A repaint used to create and destroy both objects and allocate a fresh
/// `width * height * 4` block for every animation frame. The surface belongs to
/// the preview thread, which owns the window and every repaint.
struct LayeredSurface {
    mem_dc: HDC,
    bitmap: HBITMAP,
    old_bitmap: HGDIOBJ,
    bits: *mut u8,
    width: u32,
    height: u32,
}

impl Drop for LayeredSurface {
    fn drop(&mut self) {
        unsafe {
            if !self.old_bitmap.0.is_null() {
                let _ = SelectObject(self.mem_dc, self.old_bitmap);
            }
            let _ = DeleteObject(self.bitmap);
            let _ = DeleteDC(self.mem_dc);
        }
    }
}

thread_local! {
    static LAYERED_SURFACE: RefCell<Option<LayeredSurface>> = const { RefCell::new(None) };
    /// Mutable box for the corner spinner: the small piece of a frame the arc is drawn into
    /// and then composed back over the frame's own pixels, so that drawing the spinner costs
    /// a copy of the box rather than a copy of the frame — which at the size of a display is
    /// the difference between a few kilobytes and thirty megabytes, every eighty milliseconds
    /// (see `render_layered_preview_at`).
    static OVERLAY_SCRATCH: RefCell<Vec<u8>> = const { RefCell::new(Vec::new()) };
}

/// DIB bits for a `width` x `height` frame, reusing the cached surface when the
/// size is unchanged. `None` means the surface could not be created and the
/// frame is skipped, as a failed `CreateDIBSection` did before.
fn ensure_layered_surface(width: u32, height: u32) -> Option<*mut u8> {
    LAYERED_SURFACE.with(|cell| {
        let mut surface = cell.borrow_mut();

        if let Some(existing) = surface.as_ref() {
            if existing.width == width && existing.height == height {
                return Some(existing.bits);
            }
        }

        // Dropping the previous surface releases its DC and bitmap.
        *surface = None;

        let bmi = BITMAPINFO {
            bmiHeader: BITMAPINFOHEADER {
                biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                biWidth: width as i32,
                biHeight: -(height as i32),
                biPlanes: 1,
                biBitCount: 32,
                biCompression: BI_RGB.0,
                biSizeImage: 0,
                biXPelsPerMeter: 0,
                biYPelsPerMeter: 0,
                biClrUsed: 0,
                biClrImportant: 0,
            },
            bmiColors: [Default::default()],
        };

        unsafe {
            let mem_dc = CreateCompatibleDC(None);
            if mem_dc.0.is_null() {
                return None;
            }

            let mut bits: *mut core::ffi::c_void = ptr::null_mut();
            let Ok(bitmap) = CreateDIBSection(mem_dc, &bmi, DIB_RGB_COLORS, &mut bits, None, 0)
            else {
                let _ = DeleteDC(mem_dc);
                return None;
            };

            if bits.is_null() {
                let _ = DeleteObject(bitmap);
                let _ = DeleteDC(mem_dc);
                return None;
            }

            // Kept selected for the surface's lifetime: `UpdateLayeredWindow`
            // reads the bitmap through this DC.
            let old_bitmap = SelectObject(mem_dc, bitmap);

            *surface = Some(LayeredSurface {
                mem_dc,
                bitmap,
                old_bitmap,
                bits: bits as *mut u8,
                width,
                height,
            });

            Some(bits as *mut u8)
        }
    })
}

unsafe fn render_layered_preview(hwnd: HWND) {
    let mut rect = RECT::default();
    if GetWindowRect(hwnd, &mut rect).is_err() {
        return;
    }

    render_layered_preview_at(hwnd, rect.left, rect.top);
}

/// Paint the frame the window is holding at a given place on screen, sizing the
/// window to the frame.
///
/// The window is the frame's size, which is what makes drawing a frame into the box
/// the layout planned a rule every loader follows rather than a detail of one of them
/// (see `load_media`): the box is what was fitted to the display, so a frame that grew
/// past it takes the preview past the display's edge with it.
///
/// `UpdateLayeredWindow` applies the place, the size and the surface in one call,
/// which is what lets a window that is already on screen take a frame of another
/// size — the spinner's box first, the page's after it — without a moment of the
/// frame it is holding stretched into the new box: between one call and the next, a
/// layered window shows the surface it already has, at whatever size the window
/// has.
unsafe fn render_layered_preview_at(hwnd: HWND, x: i32, y: i32) {
    let Some((width, height)) = (|| {
        let media_guard = CURRENT_MEDIA.lock().ok()?;
        let media = media_guard.as_ref()?;

        // Three kinds are not this window's to draw: a video is played by the player's own
        // window, and a document and a font specimen are drawn by the engine's.
        if matches!(
            media.media_type,
            MediaType::Video | MediaType::EngineSvg | MediaType::EngineFont
        ) {
            return None;
        }

        let width = media.current_width();
        let height = media.current_height();
        let expected_size = width as usize * height as usize * 4;
        if width == 0 || height == 0 || media.current_pixels().len() < expected_size {
            return None;
        }

        // The spinner is nothing but an arc, and what is behind it is the desktop:
        // a backdrop of the configured kind would put back the square its frame is
        // transparent to avoid. Everything else this window draws is composited over
        // the backdrop of its kind — a picture's, the one the tray keeps for a texture, or
        // the one it keeps for the picture a design document is previewed from, each a
        // setting of its own for the reason `dds_image` gives. A document is composited by
        // the engine, over the backdrop of its own, and none of them reaches here.
        let background = if media.media_type.is_loading() {
            TransparentBackground::Transparent
        } else if matches!(media.media_type, MediaType::Dds) {
            current_dds_background()
        } else if matches!(media.media_type, MediaType::Design) {
            current_design_background()
        } else if matches!(media.media_type, MediaType::Vector) {
            current_vector_background()
        } else {
            current_image_background()
        };
        let bits = ensure_layered_surface(width, height)?;
        let out = unsafe { std::slice::from_raw_parts_mut(bits, expected_size) };

        // The frame is composed once, and the spinner — where there is one — is drawn over
        // what came out rather than into a copy of the frame that has to be composed again:
        // what is copied is the box the arc sits in, a few thousand pixels of it, and what is
        // composed back is that same box (see `OVERLAY_SCRATCH`).
        compose_preview_pixels_into(
            media.current_pixels(),
            width,
            height,
            background,
            media.current_frame_is_opaque(),
            out,
        );

        if media.should_draw_streaming_overlay() {
            let elapsed = media
                .loading_start
                .map(|s| s.elapsed().as_secs_f32())
                .unwrap_or(0.0);
            let angle = elapsed * 2.0 * std::f32::consts::PI * 1.2;

            if let Some(area) = spinner_overlay_box(width, height) {
                OVERLAY_SCRATCH.with(|cell| {
                    let mut box_pixels = cell.borrow_mut();
                    copy_frame_box_into(media.current_pixels(), width, area, &mut box_pixels);
                    overlay_loading_spinner(&mut box_pixels, area, width, height, angle);
                    compose_preview_block_into(&box_pixels, area, background, out, width);
                });
            }
        }

        Some((width, height))
    })() else {
        return;
    };

    let Some(mem_dc) =
        LAYERED_SURFACE.with(|cell| cell.borrow().as_ref().map(|surface| surface.mem_dc))
    else {
        return;
    };

    let dst_point = POINT { x, y };
    let size = SIZE {
        cx: width as i32,
        cy: height as i32,
    };
    let src_point = POINT { x: 0, y: 0 };
    let blend = BLENDFUNCTION {
        BlendOp: AC_SRC_OVER as u8,
        BlendFlags: 0,
        SourceConstantAlpha: 255,
        AlphaFormat: AC_SRC_ALPHA as u8,
    };

    let _ = UpdateLayeredWindow(
        hwnd,
        None,
        Some(&dst_point),
        Some(&size),
        mem_dc,
        Some(&src_point),
        COLORREF(0),
        Some(&blend),
        ULW_ALPHA,
    );

    publish_pointer_hold(hwnd);
}

/// Put the loading spinner on screen for a pending load, in the box that load's wait
/// is placed in: the arc's own, at the pointer's corner (see `waiting_placement`).
///
/// A window that is not on screen is moved before the spinner is installed, so a
/// `WM_DPICHANGED` reset from crossing displays cannot discard it, and the spinner
/// is painted before the window is revealed, so the previous preview cannot flash
/// at the new place. It is also what moves a spinner whose box has changed size
/// while it was up: the frame is drawn at the size of the box it goes into.
unsafe fn show_loading_spinner(hwnd: HWND, pl: &PendingLoad) {
    // A wait for a hover the pointer has left is not put up, and nothing of the
    // spinner is built for it — no media, no frame, no window — so a load the hook
    // has already dismissed cannot blink a box on screen in the meantime. The
    // guard is held for the rest of the function, which is what makes this check
    // and the window it writes one step (see `HIDDEN_EPOCH`), and the pointer is
    // asked the same question the reveal is: a wait belongs at the hand that is
    // still on the file, and not at one that has crossed to another item
    // (see `HOVER_POINTER_BOX`).
    let hidden = HIDDEN_EPOCH.lock().ok();
    if !hover_still_wanted(&hidden, pl) || !pointer_on_the_hovered_item() {
        return;
    }

    let (x, y) = pl.spinner_pos;
    let side = pl.spinner_side as i32;

    // A window that is already on screen is not moved ahead of its frame, for the
    // reason the page's install gives: what a layered window shows between one
    // paint and the next is the surface it already has, stretched into whatever
    // box the window has, so a spinner whose box changed under it would be drawn
    // as a bar until the new frame lands. `UpdateLayeredWindow` applies the place
    // and the size with the frame, which is where the move below happens instead.
    let visible = IsWindowVisible(hwnd).as_bool();
    if !visible {
        let _ = MoveWindow(hwnd, x, y, side, side, false);
    }

    let loading = create_loading_media(pl.spinner_side, pl.spinner_side);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(loading);
    }

    if visible {
        render_layered_preview_at(hwnd, x, y);
    } else {
        render_layered_preview(hwnd);
    }
    let _ = SetWindowPos(
        hwnd,
        HWND_TOPMOST,
        x,
        y,
        side,
        side,
        SWP_NOACTIVATE | SWP_SHOWWINDOW,
    );
    let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
}

/// Publish — or withdraw — the regions in which the pointer keeps what is on
/// screen alive, and note which of the two things that hold a pointer is on screen.
///
/// The Explorer hook polls this to decide whether the pointer over the preview
/// means "the user is reading this" or "dismiss it and show what is underneath",
/// and the wheel hook asks the same question before it decides whether the wheel
/// belongs to Explorer or to the preview. One thing holds the pointer through a
/// region: a text preview in full mode, which the pointer can rest on to select from
/// and scroll. The other — the spinner a page is being rendered behind, which has
/// nothing under it to hand the pointer back to and no page yet to be shown in its
/// place — is noted here rather than drawn, because what it holds the pointer through
/// is the item the wait is for (see `preview_pointer_hold`).
unsafe fn publish_pointer_hold(hwnd: HWND) {
    // The pointer can only be on a window that is on screen, and while this one is not —
    // a document is drawn by the engine, in a window of its own — nothing of this app's is
    // under the pointer and there is nothing for a hold to keep alive. Asking the window
    // rather than the media is what keeps the hold from outliving the spinner it was
    // published for: a hold left standing over a document is a preview the pointer can
    // never close, because the pointer arriving at it is exactly what the hold refuses.
    if !IsWindowVisible(hwnd).as_bool() {
        clear_pointer_hold();
        return;
    }

    let (text, waiting) = CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| {
            let media = media.as_ref()?;
            Some((
                media
                    .text_state
                    .as_ref()
                    .map(|state| (state.dpi, state.can_scroll())),
                media.media_type.is_loading(),
            ))
        })
        .unwrap_or((None, false));

    // The wheel only belongs to a preview that can move under it, but the pointer
    // is held by any text preview in full mode: selecting and copying needs a
    // pointer that can rest on the preview whether or not it scrolls.
    TEXT_PREVIEW_SCROLLABLE.store(
        text.map(|(_, can_scroll)| can_scroll).unwrap_or(false),
        Ordering::Release,
    );
    TEXT_PREVIEW_HOLDING.store(text.is_some(), Ordering::Release);
    WAITING_PREVIEW_HOLDING.store(waiting, Ordering::Release);

    let keep_alive = if let Some((dpi, _)) = text {
        let mut rect = RECT::default();
        GetWindowRect(hwnd, &mut rect)
            .ok()
            .map(|_| (rect.left, rect.top, rect.right, rect.bottom))
            .map(|preview| {
                let anchor = TEXT_SCROLL_ANCHOR
                    .lock()
                    .ok()
                    .and_then(|anchor| *anchor)
                    // Without an anchor — a preview that was already on screen when this
                    // started, say — the preview's own corner stands in for it, which
                    // leaves the region as the preview alone.
                    .unwrap_or((preview.0, preview.1));

                text_scroll_hold_regions(
                    preview,
                    anchor,
                    far_edge_grace(dpi, configured_far_edge_grace_pixels()),
                    dpi,
                )
                .to_vec()
            })
    } else {
        // A wait publishes no region of its own. It holds the pointer through the item
        // it is waiting for, which is a question the hook asks of its own item box rather
        // than a rectangle this side hands it (see `preview_pointer_hold`): the spinner
        // is placed at the hand and follows it, so a box of its own would be one the
        // pointer could never leave.
        None
    };

    if let Ok(mut published) = POINTER_HOLD_REGIONS.lock() {
        *published = keep_alive;
    }
}

/// Remember the point a preview was opened from, which is what the hold region
/// stretches back to.
fn set_text_scroll_anchor(x: i32, y: i32) {
    if let Ok(mut anchor) = TEXT_SCROLL_ANCHOR.lock() {
        *anchor = Some((x, y));
    }
}

fn clear_pointer_hold() {
    TEXT_PREVIEW_SCROLLABLE.store(false, Ordering::Release);
    TEXT_PREVIEW_HOLDING.store(false, Ordering::Release);
    WAITING_PREVIEW_HOLDING.store(false, Ordering::Release);
    if let Ok(mut published) = POINTER_HOLD_REGIONS.lock() {
        *published = None;
    }
    if let Ok(mut anchor) = TEXT_SCROLL_ANCHOR.lock() {
        *anchor = None;
    }
}

/// Whether the preview on screen is a text preview that scrolls. Cheap enough for
/// the Explorer hook to ask on every poll tick.
pub fn text_preview_scrollable() -> bool {
    TEXT_PREVIEW_SCROLLABLE.load(Ordering::Acquire)
}

/// Whether the pointer is inside a region that keeps what is on screen alive, without
/// waiting on a lock the preview thread may be holding, so the Explorer hook can ask on
/// every poll tick.
///
/// A text preview holds the pointer through the regions it published — the journey to it
/// and the preview itself — because a hand reading or selecting from one is on its way
/// there or already there.
///
/// A wait holds the pointer through the *file* it is waiting for rather than through the
/// box its spinner occupies. The spinner is placed at the hand and follows it, so a
/// region of its own would be a box the pointer could never leave — and a preview the
/// hook could never close, however fast the hand is moving on to something else. What a
/// wait is owed is the question the item box answers: a pointer still inside the item the
/// hover was resolved from is a pointer still waiting for that file, and one outside it
/// has gone wherever it liked, spinner or no spinner (see `HOVER_POINTER_BOX`).
///
/// An item box nobody could be read for is *not* a hold here, which is the reverse of the
/// rule a reveal follows, and for the reason the hold exists: a hold is the hook leaving
/// the mouse alone, so it is only ever taken on an answer. A wait whose item could not be
/// read is a wait the pointer may still dismiss — what it costs is a page read again from
/// the cache, and what the other reading costs is a preview nothing can close.
pub fn preview_pointer_hold(x: i32, y: i32) -> bool {
    if TEXT_PREVIEW_HOLDING.load(Ordering::Acquire) {
        let Ok(published) = POINTER_HOLD_REGIONS.lock() else {
            return false;
        };

        return (*published)
            .as_ref()
            .map(|regions| {
                regions.iter().any(|(left, top, right, bottom)| {
                    x >= *left && x < *right && y >= *top && y < *bottom
                })
            })
            .unwrap_or(false);
    }

    WAITING_PREVIEW_HOLDING.load(Ordering::Acquire)
        && pointer_item_box().is_some_and(|item| box_holds(x, y, item))
}

/// The published preview region, without blocking. The wheel hook runs inside a
/// system-wide hook procedure, where waiting on a lock held by the preview thread
/// would stall every wheel message on the desktop — so a lock it cannot take
/// immediately means the wheel is not ours to take either.
///
/// This is the preview itself and not the journey to it: the wheel belongs to the
/// preview while the pointer is on (or just past) its edge, not while the pointer
/// is still crossing the row in front of it, where the wheel is Explorer's.
pub fn text_scroll_keep_alive_try() -> Option<(i32, i32, i32, i32)> {
    if !text_preview_scrollable() {
        return None;
    }

    POINTER_HOLD_REGIONS
        .try_lock()
        .ok()
        .and_then(|held| held.as_ref().and_then(|regions| regions.last().copied()))
}

/// Where a scroll of `lines` from the preview's current position lands.
fn text_scroll_target(lines: i64) -> Option<usize> {
    CURRENT_MEDIA.lock().ok().and_then(|media| {
        media
            .as_ref()
            .and_then(|media| media.text_state.as_ref())
            .map(|scroll| scroll.scrolled_by(lines))
    })
}

/// Move the text preview on screen to `first_line` and repaint it.
///
/// The frame is re-rendered rather than slid: only the lines coming into view are
/// styled, the window it is drawn in does not move, and a document that is
/// scrolled through costs a window of highlighting per step rather than a
/// re-read.
unsafe fn scroll_text_preview(hwnd: HWND, first_line: usize) {
    // Taken out of the media so the lock is not held while the lines are styled
    // and painted.
    let Some((path, options, dpi, width, height, current)) =
        CURRENT_MEDIA.lock().ok().and_then(|media| {
            media.as_ref().and_then(|media| {
                media.text_state.as_ref().map(|scroll| {
                    (
                        scroll.path.clone(),
                        scroll.options,
                        scroll.dpi,
                        scroll.width,
                        scroll.height,
                        scroll.first_line,
                    )
                })
            })
        })
    else {
        return;
    };

    if first_line == current {
        return;
    }

    // A selection is a range of the frame that is on screen, so it is dropped
    // rather than left pointing at lines that have moved.
    let Some(frame) =
        text_preview::render_scrolled(&path, first_line, width, height, dpi, options, None)
    else {
        return;
    };

    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        let Some(media) = media.as_mut() else {
            return;
        };

        // The hover may have moved on while this frame was rendered.
        if media.text_state.as_ref().map(|state| &state.path) != Some(&path) {
            return;
        }

        media.frames[0] = ImageFrame::new(frame.pixels, frame.width, frame.height, 0);

        if let Some(state) = media.text_state.as_mut() {
            state.first_line = frame.first_line;
            state.visible_lines = frame.visible_lines;
            state.scrollbar = frame.scrollbar;
            state.lines = frame.lines;
            state.selection = None;
        }
    }

    render_layered_preview(hwnd);
}

/// Repaint the text preview where it is, with whatever is selected in it now.
///
/// A selection is painted into the frame rather than drawn over it, so changing
/// one costs a re-render of the window that is on screen — a screenful of lines,
/// all of which are cached after the first pass.
unsafe fn repaint_text_preview(hwnd: HWND) {
    let Some((path, options, dpi, width, height, first_line, selection)) =
        CURRENT_MEDIA.lock().ok().and_then(|media| {
            media.as_ref().and_then(|media| {
                media.text_state.as_ref().map(|state| {
                    (
                        state.path.clone(),
                        state.options,
                        state.dpi,
                        state.width,
                        state.height,
                        state.first_line,
                        state.selection,
                    )
                })
            })
        })
    else {
        return;
    };

    let Some(frame) =
        text_preview::render_scrolled(&path, first_line, width, height, dpi, options, selection)
    else {
        return;
    };

    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        let Some(media) = media.as_mut() else {
            return;
        };

        if media.text_state.as_ref().map(|state| &state.path) != Some(&path) {
            return;
        }

        media.frames[0] = ImageFrame::new(frame.pixels, frame.width, frame.height, 0);

        if let Some(state) = media.text_state.as_mut() {
            state.lines = frame.lines;
        }
    }

    render_layered_preview(hwnd);
}

/// Whether a press at `(x, y)` in window coordinates lands on the scrollbar, and
/// if so, where the drag starts.
fn text_scroll_drag_target(x: i32, y: i32) -> Option<usize> {
    CURRENT_MEDIA.lock().ok().and_then(|media| {
        let media = media.as_ref()?;
        let scroll = media.text_state.as_ref()?;
        let scrollbar = scroll.scrollbar?;

        // The whole column counts, not just the groove: the bar is thin, and a
        // press a few pixels to its left is a press on the bar as far as the user
        // is concerned.
        let slack = (TEXT_SCROLL_BAR_PRESS_SLACK_PIXELS * scroll.dpi as f32 / 96.0).round() as i32;
        let (left, top, right, bottom) = scrollbar.track;
        if x < left - slack || x >= right + slack || y < top || y >= bottom {
            return None;
        }

        Some(text_preview::scroll_line_at_track_y(
            scrollbar.track,
            scrollbar.thumb,
            y,
            scroll.visible_lines,
            scroll.scrollable_lines,
        ))
    })
}

/// The document line a drag to `y` in window coordinates asks for.
fn drag_target_for_y(y: i32) -> Option<usize> {
    CURRENT_MEDIA.lock().ok().and_then(|media| {
        let scroll = media.as_ref()?.text_state.as_ref()?;
        let scrollbar = scroll.scrollbar?;
        Some(text_preview::scroll_line_at_track_y(
            scrollbar.track,
            scrollbar.thumb,
            y,
            scroll.visible_lines,
            scroll.scrollable_lines,
        ))
    })
}

fn set_text_scroll_dragging(dragging: bool) -> bool {
    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return false;
    };

    media
        .as_mut()
        .and_then(|media| media.text_state.as_mut())
        .map(|scroll| {
            let was = scroll.dragging;
            scroll.dragging = dragging;
            was
        })
        .unwrap_or(false)
}

fn is_text_scroll_dragging() -> bool {
    CURRENT_MEDIA
        .lock()
        .map(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .map(|state| state.dragging)
                .unwrap_or(false)
        })
        .unwrap_or(false)
}

/// Start selecting from a point in the preview. Answers whether there was anything
/// to select, which is what tells a press on a text preview from a press on one of
/// the other formats.
fn begin_text_selection(x: i32, y: i32) -> bool {
    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return false;
    };

    let Some(state) = media.as_mut().and_then(|media| media.text_state.as_mut()) else {
        return false;
    };
    if state.lines.is_empty() {
        return false;
    }

    // A press with no drag behind it is an empty selection, which is what clears
    // whatever was selected before.
    let at = text_preview::position_in(&state.lines, x, y);
    state.selection = Some(text_preview::Selection {
        anchor: at,
        caret: at,
    });
    state.selecting = true;
    true
}

/// Move the end of a selection to a point. Answers whether a drag is running, so
/// the caller knows whether it is the one holding the capture.
fn extend_text_selection(x: i32, y: i32) -> Option<bool> {
    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return None;
    };

    let state = media.as_mut()?.text_state.as_mut()?;
    if !state.selecting {
        return Some(false);
    }

    let at = text_preview::position_in(&state.lines, x, y);
    let changed = state
        .selection
        .map(|selection| selection.caret != at)
        .unwrap_or(false);

    if changed {
        if let Some(selection) = state.selection.as_mut() {
            selection.caret = at;
        }
    } else {
        return Some(false);
    }

    Some(true)
}

/// End a selection drag. Answers whether one was running.
fn end_text_selection() -> bool {
    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return false;
    };

    media
        .as_mut()
        .and_then(|media| media.text_state.as_mut())
        .map(|state| std::mem::replace(&mut state.selecting, false))
        .unwrap_or(false)
}

/// Whether the text preview on screen has something selected, which is what makes
/// a Ctrl+C the preview's to answer.
fn has_text_selection() -> bool {
    CURRENT_MEDIA
        .lock()
        .map(|media| {
            media
                .as_ref()
                .and_then(|media| media.text_state.as_ref())
                .and_then(|state| state.selection)
                .map(|selection| !selection.is_empty())
                .unwrap_or(false)
        })
        .unwrap_or(false)
}

/// Select everything the frame shows: the whole of a document that fits on one
/// page, and the screenful a longer one is showing.
///
/// The range is the one a Copy with nothing selected already takes, so what the
/// highlight covers is what that Copy puts on the clipboard.
unsafe fn select_all_text_preview(hwnd: HWND) {
    let Ok(mut media) = CURRENT_MEDIA.lock() else {
        return;
    };

    let Some(state) = media.as_mut().and_then(|media| media.text_state.as_mut()) else {
        return;
    };
    if state.lines.is_empty() {
        return;
    }

    state.selection = Some(text_preview::Selection {
        anchor: (0, 0),
        caret: (usize::MAX, usize::MAX),
    });

    // The repaint reads the same state through its own lock, so the guard goes
    // back before it is asked to draw.
    drop(media);

    repaint_text_preview(hwnd);
}

/// What a Copy takes from the preview on screen: what is selected, or — with
/// nothing selected — everything the frame shows.
fn text_preview_clipboard_text() -> Option<String> {
    CURRENT_MEDIA.lock().ok().and_then(|media| {
        let media = media.as_ref()?;
        let state = media.text_state.as_ref()?;

        Some(match state.selection {
            Some(selection) if !selection.is_empty() => {
                text_preview::text_in(&state.lines, selection)
            }
            _ => text_preview::frame_text(&state.lines),
        })
    })
}

/// Copy the text preview to the clipboard, answering whether there was anything to
/// copy.
///
/// The clipboard is owned by one window at a time, so this opens it, hands over a
/// movable block of UTF-16, and closes it again; a failure at any step leaves the
/// previous clipboard contents alone.
unsafe fn copy_text_preview(hwnd: HWND) -> bool {
    let Some(text) = text_preview_clipboard_text().filter(|text| !text.is_empty()) else {
        return false;
    };

    if OpenClipboard(hwnd).is_err() {
        return false;
    }

    let _ = EmptyClipboard();

    let units: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();
    let mut copied = false;

    if let Ok(block) = GlobalAlloc(GMEM_MOVEABLE, units.len() * std::mem::size_of::<u16>()) {
        let target = GlobalLock(block) as *mut u16;
        if target.is_null() {
            let _ = GlobalFree(block);
        } else {
            std::ptr::copy_nonoverlapping(units.as_ptr(), target, units.len());
            let _ = GlobalUnlock(block);

            if SetClipboardData(CF_UNICODETEXT.0 as u32, HANDLE(block.0)).is_err() {
                let _ = GlobalFree(block);
            } else {
                // The clipboard owns the block now.
                copied = true;
            }
        }
    }

    let _ = CloseClipboard();
    copied
}

/// Ask for Ctrl+C to copy the preview's selection.
///
/// The preview never takes focus, so a keystroke never arrives as a message: it is
/// read the way the rest of the app reads the keys that belong to it, by asking
/// whether Ctrl is down and C has been pressed since the last time this was asked.
/// Nothing is read at all unless there is a selection, so a Ctrl+C meant for
/// something else is left alone.
fn text_preview_copy_requested() -> bool {
    if !has_text_selection() {
        return false;
    }

    unsafe {
        let control = GetAsyncKeyState(VK_CONTROL.0 as i32) as u16 & 0x8000 != 0;
        let just_pressed = GetAsyncKeyState(VK_C.0 as i32) & 1 != 0;
        control && just_pressed
    }
}

/// What the preview's own menu offers: the whole frame selected, and what is
/// selected put on the clipboard.
const ID_TEXT_PREVIEW_SELECT_ALL: usize = 1;
const ID_TEXT_PREVIEW_COPY: usize = 2;

/// Show the preview's context menu at a point in window coordinates and act on
/// what it returns.
///
/// The menu is asked for its command rather than posting one, so it needs no
/// message loop of its own and the preview never has to take focus to be worked
/// with.
unsafe fn show_text_preview_menu(hwnd: HWND, x: i32, y: i32) {
    let Ok(menu) = CreatePopupMenu() else {
        return;
    };

    let _ = AppendMenuW(
        menu,
        MF_STRING,
        ID_TEXT_PREVIEW_SELECT_ALL,
        w!("Select All"),
    );
    let _ = AppendMenuW(menu, MF_STRING, ID_TEXT_PREVIEW_COPY, w!("Copy"));

    let mut point = POINT { x, y };
    let _ = ClientToScreen(hwnd, &mut point);

    // A menu belongs to the window in front of it, and this one is never in front:
    // asking for the foreground is what lets the menu see the click that chooses
    // from it.
    let _ = SetForegroundWindow(hwnd);

    let command = TrackPopupMenu(
        menu,
        TPM_RETURNCMD | TPM_NONOTIFY | TPM_LEFTALIGN | TPM_TOPALIGN,
        point.x,
        point.y,
        0,
        hwnd,
        None,
    );

    let _ = DestroyMenu(menu);

    match command.0 as usize {
        ID_TEXT_PREVIEW_SELECT_ALL => select_all_text_preview(hwnd),
        ID_TEXT_PREVIEW_COPY => {
            copy_text_preview(hwnd);
        }
        _ => {}
    }
}

unsafe fn reset_preview_after_display_change(hwnd: HWND) {
    let _ = ShowWindow(hwnd, SW_HIDE);
    clear_pointer_hold();

    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        if let Some(ref mut media) = *current {
            media.cancel_background_work();
            stop_video_playback(media);
        }
        *current = None;
    }

    // A document the engine is playing is a window of its own, the way a video is, so
    // it comes down with the rest of the preview rather than being left where the
    // display it was placed for used to be. The hover is replayed and it is put up
    // again at the new one.
    webview_preview::hide();
}

/// The point a mouse message was delivered at, in window coordinates.
fn message_point(lparam: LPARAM) -> (i32, i32) {
    let x = (lparam.0 & 0xFFFF) as u16 as i16 as i32;
    let y = ((lparam.0 >> 16) & 0xFFFF) as u16 as i16 as i32;
    (x, y)
}

unsafe extern "system" fn window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_DISPLAYCHANGE | WM_DPICHANGED => {
            reset_preview_after_display_change(hwnd);
            DISPLAY_RESET.store(true, Ordering::Release);
            LRESULT(0)
        }
        WM_LBUTTONDOWN => {
            // A press on the scrollbar starts a drag from where it landed, so the
            // thumb follows the pointer from the first click. Anywhere else, a
            // press on a text preview starts a selection.
            let (x, y) = message_point(lparam);
            if let Some(first_line) = text_scroll_drag_target(x, y) {
                set_text_scroll_dragging(true);
                let _ = SetCapture(hwnd);
                scroll_text_preview(hwnd, first_line);
            } else if begin_text_selection(x, y) {
                let _ = SetCapture(hwnd);
                repaint_text_preview(hwnd);
            }
            LRESULT(0)
        }
        WM_MOUSEMOVE => {
            let (x, y) = message_point(lparam);
            if is_text_scroll_dragging() {
                if let Some(first_line) = drag_target_for_y(y) {
                    scroll_text_preview(hwnd, first_line);
                }
            } else if extend_text_selection(x, y) == Some(true) {
                repaint_text_preview(hwnd);
            }
            LRESULT(0)
        }
        WM_LBUTTONUP => {
            if set_text_scroll_dragging(false) || end_text_selection() {
                let _ = ReleaseCapture();
            }
            LRESULT(0)
        }
        WM_RBUTTONUP => {
            // The menu is the preview's own, so it opens where it was asked for,
            // and the press that asks for it does not dismiss the preview.
            let (x, y) = message_point(lparam);
            if TEXT_PREVIEW_HOLDING.load(Ordering::Acquire) {
                show_text_preview_menu(hwnd, x, y);
            }
            LRESULT(0)
        }
        WM_POWERBROADCAST => {
            let power_event = wparam.0 as u32;
            match power_event {
                PBT_APMRESUMEAUTOMATIC | PBT_APMRESUMESUSPEND => {
                    // System resumed from sleep — DWM has restarted and the
                    // layered window's composition surface was destroyed.
                    // Reset the preview so the next hover creates everything fresh.
                    reset_preview_after_display_change(hwnd);
                    RESUME_FROM_SLEEP.store(true, Ordering::Release);
                }
                PBT_APMSUSPEND | PBT_APMSTANDBY => {
                    // System is going to sleep. Clean up video playback and
                    // background decoding to avoid resource leaks.
                    reset_preview_after_display_change(hwnd);
                }
                _ => {}
            }
            LRESULT(0)
        }
        windows::Win32::UI::WindowsAndMessaging::WM_PAINT => {
            let mut ps = PAINTSTRUCT::default();
            let _ = BeginPaint(hwnd, &mut ps);
            render_layered_preview(hwnd);
            let _ = EndPaint(hwnd, &ps);
            LRESULT(0)
        }
        windows::Win32::UI::WindowsAndMessaging::WM_DESTROY => LRESULT(0),
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

/// Computed preview window layout
struct PreviewLayout {
    pos_x: i32,
    pos_y: i32,
    max_width: u32,
    max_height: u32,
    preview_w: u32,
    preview_h: u32,
}

#[derive(Clone, Copy)]
struct ScreenBounds {
    left: i32,
    top: i32,
    right: i32,
    bottom: i32,
}

impl ScreenBounds {
    fn height(self) -> i32 {
        self.bottom - self.top
    }

    /// The room the display has: the whole work area, which is the largest box anything
    /// shown on this display can be drawn in — every placement mode takes its own room
    /// out of this one, so a box this size bounds all of them.
    ///
    /// It is what an engine that has to draw a preview *before* the file can be measured
    /// is asked for, rather than the room the hover's own layout came out at: a picture
    /// the image converter develops is developed at the size it is then shown at, so a
    /// room smaller than this is a picture that can never be shown any larger however
    /// much room its preview is given afterwards (see `PendingLoad::room`).
    fn room(self) -> (u32, u32) {
        (
            (self.right - self.left).max(1) as u32,
            (self.bottom - self.top).max(1) as u32,
        )
    }
}

/// The work area of the primary display, for a layout that could not be anchored to
/// the display it belongs to.
///
/// What this stands in for is one display's room, so one display is what it answers
/// with: the whole virtual screen — `SM_XVIRTUALSCREEN` and its width — is the union
/// of every display, and a preview sized to that is a preview that straddles the seam
/// between two of them, which is the thing anchoring a layout to a display is for. A
/// preview asked for on a display that cannot be named is therefore placed on the
/// primary one: somewhere it is wholly visible, rather than somewhere it is not.
fn primary_display_bounds() -> ScreenBounds {
    unsafe {
        let mut work = RECT::default();
        if SystemParametersInfoW(
            SPI_GETWORKAREA,
            0,
            Some(&mut work as *mut RECT as *mut core::ffi::c_void),
            SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS(0),
        )
        .is_ok()
        {
            return ScreenBounds {
                left: work.left,
                top: work.top,
                right: work.right,
                bottom: work.bottom,
            };
        }
    }

    virtual_screen_bounds()
}

/// Every display in one rectangle, for the fallback that cannot do better than the
/// primary display and find it missing.
fn virtual_screen_bounds() -> ScreenBounds {
    unsafe {
        let left = GetSystemMetrics(SM_XVIRTUALSCREEN);
        let top = GetSystemMetrics(SM_YVIRTUALSCREEN);
        let width = GetSystemMetrics(SM_CXVIRTUALSCREEN).max(1);
        let height = GetSystemMetrics(SM_CYVIRTUALSCREEN).max(1);

        ScreenBounds {
            left,
            top,
            right: left + width,
            bottom: top + height,
        }
    }
}

/// Where the pointer is, in screen coordinates.
fn cursor_position() -> Option<POINT> {
    let mut point = POINT::default();
    unsafe { GetCursorPos(&mut point) }.ok()?;
    Some(point)
}

/// Usable bounds of the display nearest to `(x, y)`. Anchoring layout to a
/// single monitor keeps the preview from spilling onto a neighboring display
/// when more than one is attached. Falls back to the primary display if the
/// monitor query fails — a display, rather than the union of them, since a
/// preview sized to the union is one that spills onto a neighbor.
fn monitor_bounds_from_point(x: i32, y: i32) -> ScreenBounds {
    unsafe {
        let monitor = MonitorFromPoint(POINT { x, y }, MONITOR_DEFAULTTONEAREST);
        if !monitor.is_invalid() {
            let mut info = MONITORINFO {
                cbSize: std::mem::size_of::<MONITORINFO>() as u32,
                ..Default::default()
            };
            if GetMonitorInfoW(monitor, &mut info).as_bool() {
                let work = info.rcWork;
                return ScreenBounds {
                    left: work.left,
                    top: work.top,
                    right: work.right,
                    bottom: work.bottom,
                };
            }
        }
    }

    primary_display_bounds()
}

/// The top edge that centers a `height`-tall preview on `center`, kept inside the
/// display.
///
/// Centering is the intent and the screen edge is the limit: a preview that fits
/// under a cursor near the top is placed where the cursor is rather than pushed
/// down the display, and one tall enough to reach an edge is moved only as far as
/// that edge allows.
fn centered_top(center: i32, height: i32, bounds: ScreenBounds) -> i32 {
    let lowest = (bounds.bottom - height).max(bounds.top);
    (center - height / 2).clamp(bounds.top, lowest)
}

/// The least room a way out of the text may leave before a preview is resized into
/// it. Below this the room is a sliver — the tail past a name that fills its row, the
/// last strip of a display under a row at the bottom — and a preview squeezed into it
/// says less than the one left over the name would have.
const MIN_AVOID_ROOM_PIXELS: f32 = 64.0;

/// A layout moved off the region the item it describes draws, as the `Avoid` setting
/// measured it: the name the file is listed under at `Avoid Filename`, and that name
/// with the columns a row writes beside it at `Avoid Details`.
///
/// What a placement is kept clear of, and how far: the region the `Avoid` setting
/// measured off the item, the distance the placement keeps from it, and the pointer a
/// mouse hover's placement is held clear of as well (see `avoiding_text`).
struct Clearance {
    /// The region the `Avoid` setting measured off the hovered item: the name the file
    /// is listed under, that name with the columns a row writes beside it, or `None`
    /// where the setting is off or the view reported no text for the item.
    text: Option<ScreenRegion>,
    /// The distance the placement keeps from `text`, so the text is stepped off rather
    /// than touched at its edge. Already in the pixels of the display the placement is
    /// for — like the least room a way out is worth taking, which `avoiding_text`
    /// scales from the logical distance it is written as.
    gap: i32,
    /// Where the pointer is, for a placement that is a mouse hover's: a way out that
    /// would put the preview over the pointer is not one the hover can use. A mouse
    /// preview is dismissed the moment the cursor touches it, so one placed *under* the
    /// cursor is dismissed the instant it appears — and put back the same way a moment
    /// later, for as long as the pointer sits there, which is a preview that blinks at
    /// the hand rather than one that is read. A pointer clear of the preview is what
    /// makes the dismissal mean what it says: the pointer arriving at the preview is
    /// the user asking for it to go. So a way out that keeps the pointer clear is
    /// preferred over one that does not, whatever size the two offer — and a way out
    /// that covers it is still taken where no other is left, since a preview moved as
    /// far as the display allows is all there is to give. The keyboard's placements
    /// carry no pointer: what they clear is the focused item, and a preview over a
    /// parked pointer is allowed there (see `compute_keyboard_layout`).
    cursor: Option<(i32, i32)>,
}

/// A preview is placed beside what it belongs to rather than over it, and the text of
/// the item it came from is part of what it belongs to: that item stays readable while
/// its preview is up, which is what the `Avoid` setting asks for. The placement the
/// position mode chose is therefore moved by the shortest step that clears that region —
/// past its right edge, past its left, under it or over it, whichever asks the least
/// of the preview — and only a step the display has room for is taken, so a preview
/// moved off one edge is never pushed off another. A step that would put the preview
/// over the pointer is not one a mouse hover can use either, whether or not it is the
/// shortest: the two rules the ways out are held to are the ones `Clearance` names.
///
/// A preview too large for every one of those rooms is *resized* into the roomiest of
/// them rather than left where it covers the text. That is the case a preview filling
/// the display lands in: nothing can be moved into place beside a name while the
/// preview is as wide and as tall as the display, so it is the size that gives, and the
/// preview shows as much as the room beside the name can hold — which is the same rule
/// that sized it in the first place, applied to the room that is left. A room too small
/// to be worth having is not taken at all, so a preview is never squeezed into a sliver
/// to get off a name that a usable preview would have covered anyway.
fn avoiding_text(
    layout: PreviewLayout,
    orig_dims: (u32, u32),
    preview_scale: PreviewScale,
    clearance: Clearance,
    bounds: ScreenBounds,
    dpi: u32,
) -> PreviewLayout {
    let Clearance { text, gap, cursor } = clearance;
    let Some((text_left, text_top, text_right, text_bottom)) = text else {
        return layout;
    };

    let (left, top) = (layout.pos_x, layout.pos_y);
    let (width, height) = (layout.preview_w as i32, layout.preview_h as i32);
    let min_room = logical_px(dpi, MIN_AVOID_ROOM_PIXELS);

    let covers_text = left < text_right
        && left + width > text_left
        && top < text_bottom
        && top + height > text_top;
    if !covers_text {
        return layout;
    }

    // Where a preview clear of the text would sit: past the text's right edge, before
    // its left one, under it and over it, each a gap away from it.
    let past_right = text_right + gap;
    let before_left = text_left - gap;
    let under = text_bottom + gap;
    let over = text_top - gap;

    // The four ways out of the text, each as the room the display has on that side,
    // whether that room is a width — a step to either side — or a height, where the
    // preview sits in it, and whether that anchor is the preview's own far edge. A
    // step right or down grows away from the text from its near edge and a step left
    // or up from its far one.
    let ways_out = [
        (bounds.right - past_right, true, past_right, false),
        (before_left - bounds.left, true, before_left, true),
        (bounds.bottom - under, false, under, false),
        (over - bounds.top, false, over, true),
    ];

    // The ways out are compared by whether they keep the pointer clear, then by the
    // preview's own size, then by how short the step is — see the note on `cursor`.
    let mut best: Option<(bool, i64, i32, PreviewLayout)> = None;
    for (room, along_width, anchor, far_edge) in ways_out {
        let room = room.max(0);

        // The box the way out offers: the room where it constrains the preview, and
        // what the mode allowed it where it does not — a preview moved off the text is
        // never enlarged by the move.
        let (max_width, max_height) = if along_width {
            (room.min(layout.max_width as i32) as u32, layout.max_height)
        } else {
            (layout.max_width, room.min(layout.max_height as i32) as u32)
        };
        if max_width == 0 || max_height == 0 {
            continue;
        }

        let (preview_w, preview_h) = scale_dimensions(
            orig_dims.0,
            orig_dims.1,
            max_width,
            max_height,
            preview_scale,
        );
        if preview_w == 0 || preview_h == 0 {
            continue;
        }

        // A way out that costs the preview its size is only taken where the room left
        // is worth having; one that costs it nothing is taken however little room it
        // leaves, since the preview was already going to be that small.
        let natural_size = if along_width { width } else { height };
        let size_there = if along_width {
            preview_w as i32
        } else {
            preview_h as i32
        };
        if size_there < natural_size && room < min_room {
            continue;
        }

        let placement = PreviewLayout {
            pos_x: if along_width {
                if far_edge {
                    anchor - preview_w as i32
                } else {
                    anchor
                }
            } else {
                left
            },
            pos_y: if along_width {
                top
            } else if far_edge {
                anchor - preview_h as i32
            } else {
                anchor
            },
            max_width,
            max_height,
            preview_w,
            preview_h,
        };

        // A way out that keeps the pointer clear of the preview comes first, then the
        // largest preview, and the shortest move breaks a tie: every way out that fits
        // the preview as it stands offers it the same size, so those are the ones the
        // move decides between, and only a preview that has to shrink is chosen between
        // by what the room holds. See the note on `cursor` above for why the pointer
        // comes ahead of both.
        let clear_of_cursor = cursor.is_none_or(|(x, y)| {
            !box_holds(
                x,
                y,
                (
                    placement.pos_x,
                    placement.pos_y,
                    placement.pos_x + preview_w as i32,
                    placement.pos_y + preview_h as i32,
                ),
            )
        });
        let area = preview_w as i64 * preview_h as i64;
        let step = (placement.pos_x - left).abs() + (placement.pos_y - top).abs();
        let better = match &best {
            Some((best_clear, best_area, best_step, _)) => {
                if clear_of_cursor != *best_clear {
                    // One of the two keeps the pointer clear and the other does not, and
                    // that is the whole of the choice between them.
                    clear_of_cursor
                } else {
                    area > *best_area || (area == *best_area && step < *best_step)
                }
            }
            None => true,
        };
        if better {
            best = Some((clear_of_cursor, area, step, placement));
        }
    }

    match best {
        Some((_, _, _, placement)) => placement,
        None => layout,
    }
}

/// How far off the pointer a preview is placed, in logical pixels: the margin the
/// position modes are written around, and the gap the waiting spinner is kept at as well.
///
/// The spinner is the preview window with an arc in it, so a pointer that lands on it is a
/// pointer that has stopped clicking and probing the file it is waiting on: what a wait is
/// owed is the hand's own corner and not the hand itself, and the same gap every other
/// preview keeps is a gap the arc can be touched across but not spawned in.
const POINTER_STANDOFF_PIXELS: f32 = 20.0;

/// Compute preview layout for mouse hover (relative to cursor position)
///
/// `placement` is what the hover asks for — the size the preview was measured at,
/// the text to keep it off, the position mode it follows and the scale it is drawn
/// with — the same reading a pending load keeps to place its preview again as the
/// pointer moves (see `HoverPlacement`). `dpi` is the display the pointer is on,
/// which is what the margins the placement is written around are scaled by.
///
/// A placement that is the waiting spinner (`at_the_pointer_corner`) is placed by its own
/// rule: the corner nearest the pointer is put a pointer gap off it — the same gap every
/// other preview keeps, so the arc is beside the hand rather than under it — in whichever
/// of the four quadrants the display has room for the spinner, and the name it covers is
/// not stepped around: a spinner waiting on the page of the file under the hand says what
/// it is by being at the hand's own corner, and one placed a row away from it says nothing
/// about what is being waited on. Every other preview keeps the margin its position mode
/// leaves and the room the `Avoid` setting asks for.
fn compute_mouse_layout(
    cursor_x: i32,
    cursor_y: i32,
    placement: HoverPlacement,
    bounds: ScreenBounds,
    dpi: u32,
) -> Option<PreviewLayout> {
    let HoverPlacement {
        orig_dims,
        avoid,
        follow_cursor,
        preview_scale,
        at_the_pointer_corner,
    } = placement;

    let offset = logical_px(dpi, POINTER_STANDOFF_PIXELS);
    let (orig_w, orig_h) = (orig_dims.0 as i32, orig_dims.1 as i32);

    if follow_cursor || at_the_pointer_corner {
        let quadrants = [
            (
                bounds.right - cursor_x - offset,
                bounds.bottom - cursor_y - offset,
                cursor_x + offset,
                cursor_y + offset,
            ), // BR
            (
                cursor_x - bounds.left - offset,
                bounds.bottom - cursor_y - offset,
                bounds.left,
                cursor_y + offset,
            ), // BL
            (
                bounds.right - cursor_x - offset,
                cursor_y - bounds.top - offset,
                cursor_x + offset,
                bounds.top,
            ), // TR
            (
                cursor_x - bounds.left - offset,
                cursor_y - bounds.top - offset,
                bounds.left,
                bounds.top,
            ), // TL
        ];

        let mut best_quadrant = 0;
        let mut best_scale: f32 = 0.0;

        for (i, &(avail_w, avail_h, _, _)) in quadrants.iter().enumerate() {
            if avail_w <= 0 || avail_h <= 0 {
                continue;
            }
            let scale = scale_in_room(
                avail_w as f32,
                avail_h as f32,
                orig_w as f32,
                orig_h as f32,
                preview_scale,
            );
            if scale > best_scale {
                best_scale = scale;
                best_quadrant = i;
            }
        }

        if best_scale <= 0.0 {
            return None;
        }

        let (avail_w, avail_h, _, _) = quadrants[best_quadrant];
        let max_width = avail_w.max(1) as u32;
        let max_height = avail_h.max(1) as u32;

        let (preview_w, preview_h) = scale_dimensions(
            orig_dims.0,
            orig_dims.1,
            max_width,
            max_height,
            preview_scale,
        );
        let media_width = preview_w as i32;
        let media_height = preview_h as i32;

        if media_width <= 0 || media_height <= 0 {
            return None;
        }

        let (pos_x, pos_y) = match best_quadrant {
            0 => (cursor_x + offset, cursor_y + offset),
            1 => (cursor_x - offset - media_width, cursor_y + offset),
            2 => (cursor_x + offset, cursor_y - offset - media_height),
            3 => (
                cursor_x - offset - media_width,
                cursor_y - offset - media_height,
            ),
            _ => (cursor_x + offset, cursor_y + offset),
        };

        let layout = PreviewLayout {
            pos_x,
            pos_y,
            max_width,
            max_height,
            preview_w,
            preview_h,
        };

        // The spinner is placed and left there: the step off the name is what being at
        // the pointer's own corner is instead of.
        if at_the_pointer_corner {
            return Some(layout);
        }

        Some(avoiding_text(
            layout,
            orig_dims,
            preview_scale,
            Clearance {
                text: avoid,
                gap: offset,
                cursor: Some((cursor_x, cursor_y)),
            },
            bounds,
            dpi,
        ))
    } else {
        let left_width = cursor_x - bounds.left - offset;
        let right_width = bounds.right - cursor_x - offset;
        let full_height = bounds.height();

        let left_scale = scale_in_room(
            left_width as f32,
            full_height as f32,
            orig_w as f32,
            orig_h as f32,
            preview_scale,
        );
        let right_scale = scale_in_room(
            right_width as f32,
            full_height as f32,
            orig_w as f32,
            orig_h as f32,
            preview_scale,
        );

        let (use_left, max_width, max_height) = if left_scale > right_scale && left_width > 0 {
            (true, left_width.max(1) as u32, full_height as u32)
        } else if right_width > 0 {
            (false, right_width.max(1) as u32, full_height as u32)
        } else {
            return None;
        };

        let (preview_w, preview_h) = scale_dimensions(
            orig_dims.0,
            orig_dims.1,
            max_width,
            max_height,
            preview_scale,
        );
        let media_width = preview_w as i32;
        let media_height = preview_h as i32;

        if media_width <= 0 || media_height <= 0 {
            return None;
        }

        let pos_x = if use_left {
            cursor_x - offset - media_width
        } else {
            cursor_x + offset
        };
        // Best position means beside the cursor, not in the middle of the
        // display. Centering on the cursor's own line keeps a small preview where
        // the pointer is instead of floating at the screen's center — which for a
        // cursor near the top put most of the preview below it — and the clamp is
        // what still keeps a tall preview inside the display.
        let pos_y = centered_top(cursor_y, media_height, bounds);

        let layout = PreviewLayout {
            pos_x,
            pos_y,
            max_width,
            max_height,
            preview_w,
            preview_h,
        };

        Some(avoiding_text(
            layout,
            orig_dims,
            preview_scale,
            Clearance {
                text: avoid,
                gap: offset,
                cursor: Some((cursor_x, cursor_y)),
            },
            bounds,
            dpi,
        ))
    }
}

/// How far off the item a keyboard preview is placed, in logical pixels.
const KEYBOARD_GAP_PIXELS: f32 = 10.0;

/// The least room beside an item a keyboard preview will squeeze into before it
/// stops treating the item as something to sit beside, in logical pixels. Below this
/// the free space past the item's edge — or past the region a row is kept off — is a
/// sliver, and the preview is placed from the item's middle instead — see
/// `compute_keyboard_layout`.
const MIN_BESIDE_ROOM_PIXELS: f32 = 64.0;

/// What placing a keyboard preview needs: the item it is about — its box, the region
/// the `Avoid` setting keeps it off, and whether its text is drawn as a row — and the
/// size and mode the placement is made at.
///
/// It is the keyboard's answer to `HoverPlacement`, which is the same set of facts
/// read at a cursor rather than at the focused item, and it is asked for in the same
/// way: once as the preview opens, and again with the size a text preview measured
/// itself at, which is the one thing that changes between the two.
#[derive(Clone, Copy)]
struct KeyboardPlacement {
    /// The box the item occupies on screen, which a box item's preview is placed
    /// beside and a row's is not — see `compute_keyboard_layout`.
    item_rect: (i32, i32, i32, i32),
    /// The region the `Avoid` setting keeps a preview off: the item's own text, the
    /// name alone, or the column the name sits in, as the setting has it — or `None`
    /// when nothing is kept off, which a keyboard preview is never asked with: the hook
    /// reads `Avoid Nothing` as `Avoid Filename` and answers with the item's own box
    /// when the view reported no text (see
    /// `explorer_hook::HoveredItem::keyboard_avoid_box`).
    avoid: Option<ScreenRegion>,
    /// Whether the item draws anything beside the piece its name is drawn in, which
    /// with the item's shape is what says it is a row of its view rather than a box —
    /// see `explorer_hook::ItemText`.
    columns: bool,
    /// The media's own size, which the preview is scaled into the room by.
    orig_dims: (u32, u32),
    /// The position mode: whether the preview grows away from the item or is centred
    /// beside it.
    follow_cursor: bool,
    /// The share of the media's own size the preview is drawn at.
    preview_scale: PreviewScale,
}

/// Compute preview layout for keyboard hover (relative to the focused item's box)
/// Positions the preview so it doesn't block the selected file item
///
/// `placement` is what the item asks for: its box, the region the `Avoid` setting
/// keeps a preview off — the boxes each piece of the item's own text is drawn in,
/// which the hook reads off the item's children in one batched call (see
/// `explorer_hook::item_text_box`) — whether the item draws its text as a row of its
/// view, and the size and mode the placement is made at. The region is what the
/// placement is kept clear of *and* where a row's placement is measured from, so a row
/// is only cleared as far as the setting asks; with nothing kept off, an item is
/// placed by the position mode alone. See `avoiding_text` and `KeyboardPlacement`.
///
/// A row of the view is one by the columns it draws beside its name rather than by its
/// box: the row's box is as wide as the *view* the row is drawn in, not as wide as the
/// display, so a `Details` row of a window a quarter of the display across is a row all
/// the same — and read by its box it would be placed past the whole of its columns at
/// every way of avoiding.
///
/// `dpi` is the display the item is on, which is what the margins this is written
/// around — the gap it keeps off the item and the least room beside one that is worth
/// sitting in — are scaled by.
fn compute_keyboard_layout(
    placement: KeyboardPlacement,
    bounds: ScreenBounds,
    dpi: u32,
) -> Option<PreviewLayout> {
    let KeyboardPlacement {
        item_rect,
        avoid,
        columns,
        orig_dims,
        follow_cursor,
        preview_scale,
    } = placement;

    let (item_left, item_top, item_right, item_bottom) = item_rect;
    let gap = logical_px(dpi, KEYBOARD_GAP_PIXELS);
    let min_beside_room = logical_px(dpi, MIN_BESIDE_ROOM_PIXELS);
    let (orig_w, orig_h) = (orig_dims.0 as i32, orig_dims.1 as i32);

    // An item far wider than it is tall whose text is drawn as a row — the name with
    // the columns of a `Details` or `Content` row beside it, which the hook reads off
    // the item's own text — is a row of the list: a box as wide as the view with its
    // text written into the left end of it, whatever the view is doing on the display.
    let item_width = (item_right - item_left).max(0);
    let item_height = (item_bottom - item_top).max(1);
    let row_shaped = columns && item_width >= item_height * 4;

    // What is *beside* a row is not what its edges leave: the room past the row's
    // right edge is the space the view itself is not using — a sliver at the
    // window's edge, which is where a preview squeezed beside a row used to land.
    // The room a row really offers is the empty tail past the region the `Avoid`
    // setting keeps a preview off, and a keyboard preview is placed in it: just past
    // that region's edge, at the row's own line, sized by the tail and the display's
    // height. That is the placement a box item gets past its right edge, with the
    // region's edge standing in for the box's — so at `Avoid Details` it is the
    // placement a `Details` row has always had, past all of its columns, while at
    // `Avoid Filename` the preview is only taken past the name: the columns drawn
    // after it are within what the setting allows a preview to cover.
    //
    // A row with no tail — a narrow view, a name long enough to fill it — and a row
    // with no region to be placed from at all are both left to the placement below,
    // which anchors them at their middle: there is nowhere beside such a row to put a
    // preview, and the display's own room is all there is. (The keyboard path is never
    // asked with no region — see `KeyboardPlacement`.)
    if row_shaped {
        if let Some(tail_right) = avoid
            .map(|(_, _, right, _)| right)
            .filter(|right| *right > item_left && bounds.right - *right - gap >= min_beside_room)
        {
            let max_width = (bounds.right - tail_right - gap).max(1) as u32;
            let room_below = bounds.bottom - item_bottom - gap;
            let room_above = item_top - bounds.top - gap;
            // Follow Cursor grows the preview away from the row — from below it
            // when the larger room is there, from above it when it is not — while
            // Best Position centres it on the row, the way the mouse path centres
            // one on the cursor's line.
            let max_height = if follow_cursor {
                room_below.max(room_above).max(1) as u32
            } else {
                bounds.height().max(1) as u32
            };

            let (preview_w, preview_h) = scale_dimensions(
                orig_dims.0,
                orig_dims.1,
                max_width,
                max_height,
                preview_scale,
            );
            if preview_w == 0 || preview_h == 0 {
                return None;
            }

            let pos_x = tail_right + gap;
            let pos_y = if !follow_cursor {
                centered_top((item_top + item_bottom) / 2, preview_h as i32, bounds)
            } else if room_below >= room_above {
                item_bottom + gap
            } else {
                item_top - gap - preview_h as i32
            };

            let layout = PreviewLayout {
                pos_x,
                pos_y,
                max_width,
                max_height,
                preview_w,
                preview_h,
            };

            return Some(avoiding_text(
                layout,
                orig_dims,
                preview_scale,
                Clearance {
                    text: avoid,
                    gap,
                    cursor: None,
                },
                bounds,
                dpi,
            ));
        }
    }

    // What the placement is anchored at: the item's own edges for a box, and its
    // *middle* for a row that leaves no tail to be placed in (see above). Anchoring
    // at the middle is what the mouse path does with the cursor, so such a row is
    // read the way a hover over it is, and the preview is allowed to cover the rest
    // of it.
    let (anchor_left, anchor_top, anchor_right, anchor_bottom) = if row_shaped {
        let center_x = (item_left + item_right) / 2;
        let center_y = (item_top + item_bottom) / 2;
        (center_x, center_y, center_x, center_y)
    } else {
        (item_left, item_top, item_right, item_bottom)
    };

    if follow_cursor {
        // Quadrant-based positioning relative to what the item is anchored at
        let quadrants = [
            // Bottom-Right of it
            (
                bounds.right - anchor_right - gap,
                bounds.bottom - anchor_bottom - gap,
                anchor_right + gap,
                anchor_bottom + gap,
            ),
            // Bottom-Left of it
            (
                anchor_left - bounds.left - gap,
                bounds.bottom - anchor_bottom - gap,
                bounds.left,
                anchor_bottom + gap,
            ),
            // Top-Right of it
            (
                bounds.right - anchor_right - gap,
                anchor_top - bounds.top - gap,
                anchor_right + gap,
                bounds.top,
            ),
            // Top-Left of it
            (
                anchor_left - bounds.left - gap,
                anchor_top - bounds.top - gap,
                bounds.left,
                bounds.top,
            ),
        ];

        let mut best_quadrant = 0;
        let mut best_scale: f32 = 0.0;

        for (i, &(avail_w, avail_h, _, _)) in quadrants.iter().enumerate() {
            if avail_w <= 0 || avail_h <= 0 {
                continue;
            }
            let scale = scale_in_room(
                avail_w as f32,
                avail_h as f32,
                orig_w as f32,
                orig_h as f32,
                preview_scale,
            );
            if scale > best_scale {
                best_scale = scale;
                best_quadrant = i;
            }
        }

        if best_scale <= 0.0 {
            return None;
        }

        let (avail_w, avail_h, _, _) = quadrants[best_quadrant];
        let max_width = avail_w.max(1) as u32;
        let max_height = avail_h.max(1) as u32;

        let (preview_w, preview_h) = scale_dimensions(
            orig_dims.0,
            orig_dims.1,
            max_width,
            max_height,
            preview_scale,
        );
        let media_width = preview_w as i32;
        let media_height = preview_h as i32;

        if media_width <= 0 || media_height <= 0 {
            return None;
        }

        let (pos_x, pos_y) = match best_quadrant {
            0 => (anchor_right + gap, anchor_bottom + gap),
            1 => (anchor_left - gap - media_width, anchor_bottom + gap),
            2 => (anchor_right + gap, anchor_top - gap - media_height),
            3 => (
                anchor_left - gap - media_width,
                anchor_top - gap - media_height,
            ),
            _ => (anchor_right + gap, anchor_bottom + gap),
        };

        let layout = PreviewLayout {
            pos_x,
            pos_y,
            max_width,
            max_height,
            preview_w,
            preview_h,
        };

        Some(avoiding_text(
            layout,
            orig_dims,
            preview_scale,
            Clearance {
                text: avoid,
                gap,
                cursor: None,
            },
            bounds,
            dpi,
        ))
    } else {
        // Best spot mode: choose the left or right side of what the item is anchored
        // at — its own edges for a box, its middle for a row that leaves no tail to
        // be placed in (see above).
        //
        // The room a side offers is the room past that anchor, which for a box is
        // what keeps the preview off the file it describes. A box that leaves no room
        // on either side is placed from its middle with the display's own room, since
        // a preview squeezed into what is left past its edge is a sliver while one
        // placed from its middle takes the size the display allows; a row without a
        // tail is already anchored there.
        let edge_left_width = anchor_left - bounds.left - gap;
        let edge_right_width = bounds.right - anchor_right - gap;

        let (left_anchor_x, right_anchor_x, left_width, right_width) =
            if edge_left_width < min_beside_room && edge_right_width < min_beside_room {
                let center = ((anchor_left + anchor_right) / 2).clamp(bounds.left, bounds.right);
                (
                    center,
                    center,
                    center - bounds.left - gap,
                    bounds.right - center - gap,
                )
            } else {
                (anchor_left, anchor_right, edge_left_width, edge_right_width)
            };

        let full_height = bounds.height();

        let left_scale = scale_in_room(
            left_width as f32,
            full_height as f32,
            orig_w as f32,
            orig_h as f32,
            preview_scale,
        );
        let right_scale = scale_in_room(
            right_width as f32,
            full_height as f32,
            orig_w as f32,
            orig_h as f32,
            preview_scale,
        );

        let (use_left, max_width, max_height) = if left_scale > right_scale && left_width > 0 {
            (true, left_width.max(1) as u32, full_height as u32)
        } else if right_width > 0 {
            (false, right_width.max(1) as u32, full_height as u32)
        } else {
            return None;
        };

        let (preview_w, preview_h) = scale_dimensions(
            orig_dims.0,
            orig_dims.1,
            max_width,
            max_height,
            preview_scale,
        );
        let media_width = preview_w as i32;
        let media_height = preview_h as i32;

        if media_width <= 0 || media_height <= 0 {
            return None;
        }

        let pos_x = if use_left {
            left_anchor_x - gap - media_width
        } else {
            right_anchor_x + gap
        };
        // The same rule as the mouse path, centered on the line the preview
        // belongs to: beside the item it describes, not adrift in the display.
        let pos_y = centered_top((anchor_top + anchor_bottom) / 2, media_height, bounds);

        let layout = PreviewLayout {
            pos_x,
            pos_y,
            max_width,
            max_height,
            preview_w,
            preview_h,
        };

        Some(avoiding_text(
            layout,
            orig_dims,
            preview_scale,
            Clearance {
                text: avoid,
                gap,
                cursor: None,
            },
            bounds,
            dpi,
        ))
    }
}

pub fn run_preview_window() {
    // Page sizes come from Windows.Data.Pdf and picture sizes from the codec Windows
    // has, so this thread needs an apartment before the first layout asks for one.
    pdf_preview::initialize_apartment();
    wic_image::initialize_apartment();

    let (tx, rx): (Sender<PreviewMessage>, Receiver<PreviewMessage>) = channel();

    // Store sender for other threads to use
    if let Ok(mut sender) = PREVIEW_SENDER.lock() {
        *sender = Some(tx);
    }

    unsafe {
        let hinstance = GetModuleHandleW(None).unwrap();

        // Register window class
        let wc = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(window_proc),
            cbClsExtra: 0,
            cbWndExtra: 0,
            hInstance: hinstance.into(),
            hIcon: Default::default(),
            hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            hbrBackground: Default::default(),
            lpszMenuName: PCWSTR::null(),
            lpszClassName: PREVIEW_CLASS,
            hIconSm: Default::default(),
        };

        RegisterClassExW(&wc);

        // Create the preview window
        let hwnd = CreateWindowExW(
            WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_NOACTIVATE,
            PREVIEW_CLASS,
            w!("Preview"),
            WS_POPUP,
            0,
            0,
            1,
            1,
            None,
            None,
            hinstance,
            None,
        )
        .unwrap();

        // Store HWND as isize
        PREVIEW_HWND.store(hwnd.0 as isize, Ordering::SeqCst);

        // Track current video path to avoid restarting
        let mut current_video_path: Option<PathBuf> = None;
        // The display the hover on screen was measured against: what a sound's card is painted
        // at, and what it is painted at again while its player runs.
        let mut audio_card_dpi = 96u32;
        // When the sound on screen was started, and when its card was last painted. The clock a
        // sound FFmpeg plays is measured from the first, and the second is the cadence its card
        // is drawn at (see `audio_clock` and `AUDIO_CARD_REPAINT`).
        let mut audio_started: Option<Instant> = None;
        // The second of the file the sound on screen was started at, which is nothing for one
        // that started at its beginning: what a card's clock is measured from where the player
        // reports no clock of its own (see `audio_clock`).
        let mut audio_start_offset = 0.0f64;
        // A `Volume → Audio Seek` of `Middle` or `Random` asked of a file whose length nothing
        // had read yet, which is the one start position that cannot be worked out where the
        // sound is started: a share of a length is asked for again the moment a player reports
        // one (see the tick below), and this is the ask, held until then.
        let mut audio_share_seek: Option<AudioSeek> = None;
        let mut audio_repaint_at = Instant::now();
        // The name of the sound on screen, scrolled sideways while the card has no room for it:
        // the scroll is put up with the card it belongs to and advanced by the repaints below,
        // so a hover that changes begins at the beginning again (see `audio_preview::NameScroll`).
        let mut audio_name_scroll: Option<audio_preview::NameScroll> = None;
        // The hover the preview on screen came from, so a theme or Markdown
        // switch can rebuild it without waiting for the next hover.
        let mut current_show: Option<PreviewMessage> = None;
        // Track video position/size for periodic topmost re-assertion
        let mut video_pos: (i32, i32, i32, i32) = (0, 0, 0, 0); // (x, y, w, h)
        let mut last_topmost_check = Instant::now();

        // Background loading support
        let (load_tx, load_rx): (Sender<LoadResult>, Receiver<LoadResult>) = channel();
        let load_request_slot: LoadRequestSlot = Arc::new((Mutex::new(None), Condvar::new()));
        let load_worker = spawn_load_worker(Arc::clone(&load_request_slot), load_tx);
        let mut current_generation: u64 = 0;
        let mut pending_load: Option<PendingLoad> = None;
        let mut pending_load_cancel: Option<Arc<AtomicBool>> = None;
        let mut last_stream_overlay_repaint = Instant::now();
        // The page an engine owes the preview on screen — Office's render tier where the
        // document is one of its own, the render engine where it is one of `[libre]`'s —
        // and whether one has been asked for and is being waited on.
        let mut page_render_pending: Option<(PathBuf, u64)> = None;
        // The hover an engine has already been asked to come up for, so that the ask is made
        // once for a file the pointer has settled on rather than on every tick (see
        // `warm_engines_for`).
        let mut warmed_generation: Option<u64> = None;
        // A page that arrived for the hover already on screen, and is being loaded
        // to replace what is there rather than to open a new preview.
        let mut page_upgrade: Option<PathBuf> = None;
        // The video whose probe a hover is waiting on, and the generation of the hover
        // that is waiting: the geometry is measured on a thread of its own and the hover
        // is replayed when the answer lands (see `VideoProbed`).
        let mut video_probe: Option<(PathBuf, u64)> = None;
        // The hover a video's probe has just answered for. Its replay is the same wait
        // carried on rather than a new preview: what is on screen — the spinner — stays
        // where it is until the video replaces it, the way a page landing on a spinner
        // behaves (see `upgrading`).
        let mut video_replay: Option<PathBuf> = None;
        // The hover a box measured off the preview thread has just answered for. Its replay is
        // the same wait carried on rather than a new preview, the way a video's is (see
        // `measure_replay`).
        let mut measure_replay: Option<PathBuf> = None;
        // A player that has been started and has not put its window up yet: the wait
        // for a video, which the spinner stands in for until the player's window is
        // there (see `VideoStart`).
        let mut video_start: Option<VideoStart> = None;

        // Message loop
        let mut msg = MSG::default();
        // A message the idle wait took off the channel, held for the drain below
        // rather than acted on where it was received.
        let mut carried_preview_msg: Option<PreviewMessage> = None;
        while RUNNING.load(Ordering::SeqCst) {
            // Every tick is noted, whether it does anything or not: what the note is
            // for is the Explorer hook telling a loop that is working from one that has
            // stopped, and a loop waiting on the channel is working (see
            // `preview_stall_ms`).
            note_preview_alive();

            // Check for Windows messages
            while PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE).as_bool() {
                let _ = TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }

            // Detect resume from sleep (WM_POWERBROADCAST handler sets this flag).
            // DWM restarts on resume and destroys the layered window's composition
            // surface, so we must reset all local state to force a fresh start on
            // the next hover.
            if RESUME_FROM_SLEEP.load(Ordering::Acquire) {
                RESUME_FROM_SLEEP.store(false, Ordering::Release);
                current_generation += 1;
                pending_load = None;
                clear_load_request(&load_request_slot);
                if let Some(cancel) = pending_load_cancel.take() {
                    cancel.store(true, Ordering::Release);
                }
                current_video_path = None;
                video_pos = (0, 0, 0, 0);

                // A browser engine is a process that does not survive a suspend in any
                // state worth keeping, so it is let go with everything else and begun
                // again when the next document asks for it.
                webview_preview::shutdown();

                // Re-assert layered window style after DWM restart.
                // DWM is reinitialized during resume and the layered window's
                // per-pixel alpha composition surface may need a fresh anchor.
                let ex_style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
                SetWindowLongPtrW(
                    hwnd,
                    GWL_EXSTYLE,
                    ex_style | WS_EX_LAYERED.0 as isize | WS_EX_TOPMOST.0 as isize,
                );
                let _ = SetWindowPos(
                    hwnd,
                    HWND_TOPMOST,
                    0,
                    0,
                    0,
                    0,
                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED,
                );
            }

            // A player that has been started and has not put its window up yet is being
            // waited for: what is on screen is the spinner standing in for the video, at
            // the pointer it is shown at for every other kind of wait, and the wait ends
            // when the player's window is there — the frame the player was started for is
            // handed over then, and this app's window goes with the wait. A player that
            // never came up — one that died, one that took longer than a start ever does
            // — is ended here rather than left playing behind nothing (see `player_wait`).
            if let Some(start) = video_start.take() {
                let window_up = VIDEO_HWND.load(Ordering::SeqCst) != 0
                    && VIDEO_PID.load(Ordering::SeqCst) == start.pid;
                let alive = is_ffplay_pid_alive(start.pid);

                // The hover this player was started for may have moved on before its
                // window was there. The player is still this app's — nobody else holds
                // it — so it is ended below, but what is on screen and what is pending
                // are another hover's by then and are left alone.
                let current = current_video_path.as_deref() == Some(start.path.as_path());
                let outcome = if current {
                    player_wait(window_up, alive, start.started.elapsed())
                } else {
                    Some(PlayerWait::Abandoned)
                };

                match outcome {
                    Some(PlayerWait::Arrived) => {
                        // The video is what is on screen from here: the frame the player
                        // plays into is the preview, and the spinner — this app's window,
                        // which was standing in for it — comes down with the wait.
                        if let Ok(mut media) = CURRENT_MEDIA.lock() {
                            *media = Some(start.media);
                        }

                        let _ = ShowWindow(hwnd, SW_HIDE);
                        pending_load = None;
                        clear_pointer_hold();
                    }
                    Some(PlayerWait::Abandoned) => {
                        // The player is not coming, or the hover it was started for has
                        // gone: ending it here is what keeps a process this app started
                        // from playing behind a preview that has gone.
                        let mut media = start.media;
                        stop_video_playback(&mut media);

                        if current {
                            let _ = ShowWindow(hwnd, SW_HIDE);
                            pending_load = None;
                            clear_pointer_hold();
                        }
                    }
                    // Still starting: the wait goes on, and the player stays in hand.
                    None => video_start = Some(start),
                }
            }

            // Periodically re-assert topmost on the video window to prevent it
            // from falling behind Explorer or other windows (Bug 2 fix)
            if current_video_path.is_some()
                && last_topmost_check.elapsed() >= Duration::from_millis(200)
            {
                last_topmost_check = Instant::now();
                let _ =
                    ensure_video_window_topmost(video_pos.0, video_pos.1, video_pos.2, video_pos.3);
            }

            // Advance animation frames if needed
            let mut needs_repaint = false;

            // The engine draws a document in a window of its own, and that window is put
            // up only once the page has arrived: what is underneath it — the spinner the
            // wait was shown as — comes down then, and the wait comes down with it. What
            // is on screen is a document, and what is shown next is another hover's.
            //
            // What ends a wait is the document the wait is *for*: the engine is handed one
            // file at a time, and a page that lands names the file it was drawn for — so a
            // landing for a hover the loop has already left ends nothing, and the wait goes
            // on for the file the pointer is actually on. Taking any landing as the end of
            // any wait is what put a file the pointer had left on screen and dropped the
            // wait for the one it was on (see `webview_preview::showing_path`).
            if let Some(shown) = webview_preview::showing_path() {
                if IsWindowVisible(hwnd).as_bool() {
                    let _ = ShowWindow(hwnd, SW_HIDE);
                }

                let waiting_for_this = pending_load.as_ref().is_some_and(|pl| pl.path == shown);

                if waiting_for_this && pending_load.take().is_some() {
                    if let Some(cancel) = pending_load_cancel.take() {
                        cancel.store(true, Ordering::Release);
                    }
                    clear_pointer_hold();
                }
            }

            if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                if let Some(ref mut media) = *media_guard {
                    if media.advance_frame() {
                        needs_repaint = true;
                    }
                    if media.update_loading_frame() {
                        needs_repaint = true;
                    }
                    // A video the media engine plays is a new picture every frame rather
                    // than a file that is decoded once and drawn again, so its frame is
                    // taken here — once a tick, which is as often as one can be shown.
                    // The kind is asked first so that a preview of any other kind pays
                    // for this with one comparison.
                    if media.media_type.is_native_video() && media.take_native_video_frame() {
                        needs_repaint = true;
                    }
                    // A sound's card is the one painted preview that changes while it is on
                    // screen: the clock and the bar under it are drawn from a player that is
                    // running, and a name the card has no room for is scrolled across it. The
                    // clock's own seconds are worth watching four times a second and a scroll
                    // is not, so a card with a name to move is painted at the cadence the
                    // spinner's overlay uses and one whose whole name fits keeps the slower one
                    // (see `AUDIO_CARD_REPAINT` and `AUDIO_NAME_REPAINT`). What a card with no
                    // player behind it — `Volume → Audio` at 0% — costs is its scroll and
                    // nothing else.
                    if media.media_type.is_audio() {
                        // A sound that was asked to start somewhere other than the beginning
                        // and has not been taken there yet is taken there here, on the first
                        // tick the engine will accept a seek on — which is as soon as it has
                        // read the file's own header rather than the next card repaint a
                        // quarter of a second away, and what is heard of the beginning in
                        // between is that much and no more (see `video_player::apply_seek`).
                        video_player::apply_seek();

                        // A sound FFmpeg plays whose pass has ended is put round to the
                        // beginning of its file here, on the tick the player's own exit is
                        // found rather than on the next card repaint a quarter of a second
                        // away: what stands between two passes is the time it takes to start
                        // a player, and nothing of this side is to be added to it (see
                        // `wrap_audio_player`).
                        if let Some(path) = current_show.as_ref().and_then(self::show_path) {
                            wrap_audio_player(
                                media,
                                path,
                                &mut audio_started,
                                &mut audio_start_offset,
                            );
                        }

                        let cadence = match &audio_name_scroll {
                            Some(scroll) if scroll.moves() => AUDIO_NAME_REPAINT,
                            _ => AUDIO_CARD_REPAINT,
                        };

                        if audio_repaint_at.elapsed() >= cadence {
                            audio_repaint_at = Instant::now();

                            let name_offset = match audio_name_scroll.as_mut() {
                                Some(scroll) => {
                                    scroll.advance(Instant::now(), cadence);
                                    scroll.offset()
                                }
                                None => 0,
                            };

                            if let Some(path) = current_show.as_ref().and_then(self::show_path) {
                                let (elapsed, duration) =
                                    audio_clock(path, audio_started, audio_start_offset);

                                // A start position that was a share of a length nothing had
                                // read is asked for here, on the first tick a player says how
                                // long the file is — the ask is one that can be made at any
                                // point in a running sound, since what it is is a seek, and
                                // what it is not is a reason to have left the sound at its
                                // beginning (see `video_player::seek`). A player that says
                                // nothing about its length leaves the ask standing rather than
                                // spending it on a tick it cannot answer.
                                if let (Some(seek), Some(duration)) = (audio_share_seek, duration)
                                {
                                    audio_share_seek = None;

                                    let shared =
                                        audio_seek::start_position(path, seek, Some(duration));
                                    video_player::seek(shared);
                                }

                                // Where the sound had got to is what the mode that resumes one
                                // reads back, so it is written down as the card is repainted —
                                // the only moment anything here knows it. The other three ways
                                // of starting a sound are rules rather than memories and are
                                // not written down at all: a run that is on one of them keeps
                                // nothing, which is what makes the memory the setting's own.
                                if current_audio_seek() == AudioSeek::Remember {
                                    if let Some(elapsed) = elapsed {
                                        audio_seek::remember(path, elapsed);
                                    }
                                }

                                if media.refresh_audio_card(
                                    path,
                                    elapsed,
                                    duration,
                                    audio_card_dpi,
                                    name_offset,
                                ) {
                                    needs_repaint = true;
                                }
                            }
                        }
                    }
                    // While streaming first-frame loading, repaint for spinner animation.
                    if media.should_draw_streaming_overlay()
                        && last_stream_overlay_repaint.elapsed() >= Duration::from_millis(83)
                    {
                        last_stream_overlay_repaint = Instant::now();
                        needs_repaint = true;
                    }
                }
            }
            if needs_repaint {
                render_layered_preview(hwnd);
            }

            // A page an engine has finished with, held apart from the hovers: it is
            // not a hover to act on but an answer about the one on screen. Office's
            // tier and the engines beside it send it as a message, a page the render
            // engine draws is read from the folder it lands in — below — and a load
            // that comes back with nothing to draw because the engine has already
            // finished the work has the answer in hand rather than in a message (see
            // the `awaiting_render` arm of the load below).
            let mut page_ready: Option<(PathBuf, u64, bool)> = None;

            // Check for completed background loads
            while let Ok(result) = load_rx.try_recv() {
                // A load the pointer has left since it was started is not this
                // loop's to take up: the hide that took the window down moved the
                // count it carries (see `HIDDEN_EPOCH`), and a pointer that has
                // crossed to another item since the hover was resolved is the same
                // answer one moment earlier, before any hide has been sent for it
                // (see `HOVER_POINTER_BOX`). Either way the answer is dropped here
                // rather than built — no frame installed, no engine window put up,
                // no player started — for a hover that has already gone.
                //
                // What the wait itself put on screen comes down with it, and that is not
                // the same thing as dropping the answer: a spinner left standing is a wait
                // with nothing left to end it — the load that would have answered it is
                // gone from here, and the pointer cannot dismiss it either, because a wait
                // holds the pointer it is waiting for (see `WAITING_PREVIEW_HOLDING`) and
                // the pointer sitting on that spinner is the one thing the hold refuses. A
                // hand that crossed off the file while the load ran would be left with a
                // preview under it that nothing would ever take down.
                //
                // Asked first is whether this is the hover's own load: a keyboard hover
                // carries no placement, a newer hover is another generation, and a stale
                // answer for a load already dropped is nothing to take down twice
                // (see `HoverPlacement`).
                if pending_load
                    .as_ref()
                    .is_some_and(|pl| pl.hide_epoch != hidden_epoch())
                    || !pointer_on_the_hovered_item()
                {
                    let abandoned = pending_load.as_ref().is_some_and(|pl| {
                        pl.generation == result.generation && pl.placement.is_some()
                    });

                    pending_load = None;

                    if abandoned {
                        // The take-down the `Hide` message performs, for the same reason and
                        // in the same order: the wait is over, so its window, its media, the
                        // hold it published and anything it had asked for go with it.
                        current_generation += 1;
                        clear_load_request(&load_request_slot);
                        if let Some(cancel) = pending_load_cancel.take() {
                            cancel.store(true, Ordering::Release);
                        }

                        let _ = ShowWindow(hwnd, SW_HIDE);
                        clear_pointer_hold();
                        webview_preview::hide();

                        if let Ok(mut current) = CURRENT_MEDIA.lock() {
                            if let Some(ref mut media) = *current {
                                media.cancel_background_work();
                                stop_video_playback(media);
                            }
                            *current = None;
                        }
                        current_video_path = None;
                        video_pos = (0, 0, 0, 0);

                        if let Some(path) = current_show.as_ref().and_then(self::show_path) {
                            office_render::hover_ended(path);
                        }

                        current_show = None;
                        page_render_pending = None;
                        page_upgrade = None;
                    } else {
                        pending_load_cancel = None;
                    }

                    continue;
                }

                if result.generation == current_generation {
                    match result.media {
                        // A document the engine draws. There is no frame of this app's to
                        // install and no window of this app's to put up, so what happens
                        // here is the handover: the engine is told where the wait has
                        // ended up, and the wait stays armed for the document to land on.
                        Some(media) if media.media_type.is_engine() => {
                            let pending = pending_load.take();
                            pending_load_cancel = None;

                            if let Some(pl) = pending.as_ref() {
                                webview_preview::show(
                                    &pl.path,
                                    webview_preview::Area {
                                        x: pl.pos_x,
                                        y: pl.pos_y,
                                        width: pl.width as i32,
                                        height: pl.height as i32,
                                    },
                                    // The engine draws two kinds — a document and a font
                                    // specimen — and each is composited over a backdrop of
                                    // its own rather than over the picture's.
                                    engine_background(&pl.path),
                                );
                            }

                            if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                if let Some(ref mut existing) = *current {
                                    existing.cancel_background_work();
                                }
                                // Nothing of this app's goes on screen for a document:
                                // what is up is the engine's window, and what stands in
                                // for it until that arrives is the spinner.
                                *current = None;
                            }

                            // The wait stays armed for the document to land on: an
                            // engine that has to start a browser is a wait like any
                            // other, shown as one once the delay has run, while one
                            // that draws the document in a few milliseconds is over
                            // before there is anything to show (see `spinner_due`).
                            pending_load = pending;
                        }
                        Some(mut media_data) => {
                            // A video the media engine plays is started here, before its
                            // preview is put up: what this window is about to draw is a
                            // frame of it, and an engine that would not start is a file
                            // with no preview rather than a box of the placeholder pixels
                            // a video preview is opened with.
                            if media_data.media_type.is_native_video() {
                                let (width, height) =
                                    (media_data.current_width(), media_data.current_height());

                                // A hover that lands on the file already playing leaves it
                                // playing, the same way the FFmpeg path compares the file
                                // it last started.
                                if video_player::playing_path().as_deref()
                                    == Some(result.path.as_path())
                                    && video_player::is_playing()
                                {
                                    video_player::resize(width, height);
                                } else {
                                    video_player::play(
                                        &result.path,
                                        width,
                                        height,
                                        current_video_volume(),
                                    );
                                }

                                if !video_player::is_playing() {
                                    let _ = ShowWindow(hwnd, SW_HIDE);
                                    clear_pointer_hold();
                                    pending_load = None;

                                    if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                        if let Some(ref mut existing) = *current {
                                            existing.cancel_background_work();
                                        }
                                        *current = None;
                                    }

                                    continue;
                                }
                            }

                            // A sound is started here, before its card goes up, for the reason
                            // a video's engine is: what the card draws is the clock of a player
                            // that is running. A player that was asked for and did not come up is
                            // a sound with nothing behind it, which is the answer a video's engine
                            // that will not start gets — while a card at `Volume → Audio` 0% asks
                            // for no player at all and is left standing, with its clock still and
                            // its bar empty.
                            if media_data.media_type.is_audio() {
                                // Where the sound is dropped in: a question about the file's
                                // own length and the tray's `Volume → Audio Seek`, and one that
                                // is answered here rather than by the player, which knows
                                // neither. A length nothing has read yet — a container that
                                // does not say, on a machine whose engine may still know it —
                                // leaves the two shares of one unanswered until a player
                                // reports one, which the tick below is what asks again.
                                let seek = current_audio_seek();
                                let length = audio_track::playable(&result.path)
                                    .and_then(|track| track.duration);
                                let start = audio_seek::start_position(&result.path, seek, length);

                                audio_share_seek = (start == 0.0
                                    && matches!(seek, AudioSeek::Middle | AudioSeek::Random)
                                    && length.is_none())
                                .then_some(seek);

                                if !start_audio_playback(&result.path, &mut media_data, start) {
                                    let _ = ShowWindow(hwnd, SW_HIDE);
                                    clear_pointer_hold();
                                    pending_load = None;

                                    if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                        if let Some(ref mut existing) = *current {
                                            existing.cancel_background_work();
                                        }
                                        *current = None;
                                    }

                                    continue;
                                }

                                // The clock a sound FFmpeg plays is this app's own over the
                                // moment the player was started, and there is a player to
                                // measure from exactly where one was started: at
                                // `Volume → Audio` 0% nothing was, and a card whose clock ran
                                // anyway would be a sound it says is playing that is not — and
                                // a position this side would write down as one the file had
                                // been left at (see `audio_seek::remember`).
                                audio_started =
                                    media_data.video_process.is_some().then(Instant::now);
                                audio_start_offset = start;
                                audio_repaint_at = Instant::now();
                                // The marquee the card's name is drawn with, if it needs one:
                                // what a name is scrolled by is the card's own box, which is
                                // the frame that has just arrived, and a scroll is put up with
                                // the card rather than left over from the hover before it (see
                                // `audio_preview::NameScroll`).
                                audio_name_scroll = Some(audio_preview::NameScroll::of(
                                    &audio_preview::name_of(&result.path),
                                    media_data.current_width(),
                                    audio_card_dpi,
                                    current_audio_options(),
                                ));
                            }

                            let mw = media_data.current_width() as i32;
                            let mh = media_data.current_height() as i32;

                            // A load whose spinner was up is placed like any other:
                            // the wait was shown in the spinner's own box at the
                            // pointer, not in the box the preview arrives in (see
                            // `spinner_pos`), so the frame is installed at the
                            // preview's place — the paint below is what takes the
                            // window from one box to the other.
                            let mut pending = pending_load
                                .take()
                                .filter(|pl| pl.generation == result.generation);
                            pending_load_cancel = None;

                            // Placed once more before it is painted, from the pointer as it is
                            // now: the box this frame lands in was decided when the hover was
                            // asked for, and the hand has had the whole load to move on since.
                            // A preview that arrives under the hand is taken down again by the
                            // touch rule at the next tick — the spawn and the dismissal are one
                            // event for the eye — while one that lands clear of it stays. This
                            // is the placement every tick of the wait already makes (see
                            // `PendingLoad::follow_pointer`), one read before the paint, so what
                            // is painted is the preview of the file under the hand rather than
                            // one under the hand itself. A keyboard hover carries no placement
                            // and is left where the item put it.
                            if let Some(pl) = pending.as_mut() {
                                if let Some(cursor) = cursor_position() {
                                    pl.follow_pointer(
                                        cursor,
                                        monitor_dpi_from_point(cursor.x, cursor.y),
                                    );
                                }
                            }

                            // What this load was planned for: the box the layout
                            // came out with, which is the size a slide is
                            // exported at when a page is asked for later.
                            let render_box =
                                pending.as_ref().map(|pl| (pl.width, pl.height)).unwrap_or((
                                    media_data.current_width(),
                                    media_data.current_height(),
                                ));

                            // Move before installing the frame. Crossing between
                            // displays of different scale sends WM_DPICHANGED,
                            // which resets the preview and would otherwise
                            // discard the frame we are about to show, leaving the
                            // other display's image stranded on screen.
                            //
                            // A window that is already on screen is not moved
                            // ahead of the frame, though: a layered window shows
                            // the surface it has at whatever size the window has,
                            // so resizing this one to the page's box before there
                            // is a page to fill it draws the frame it is holding —
                            // the spinner — stretched across that box for as long
                            // as the paint takes. It is moved and resized by the
                            // paint itself, while a hidden window is moved here,
                            // where nothing can show.
                            let visible = IsWindowVisible(hwnd).as_bool();
                            if let Some(ref pl) = pending {
                                if !visible {
                                    let _ = MoveWindow(hwnd, pl.pos_x, pl.pos_y, mw, mh, false);
                                }
                            }

                            if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                if let Some(ref mut existing) = *current {
                                    existing.cancel_background_work();
                                }
                                *current = Some(media_data);
                            }

                            // A layered window keeps its surface while hidden, so
                            // paint the new frame before revealing the window.
                            // Showing first would flash the previous preview at
                            // the new position and size.
                            //
                            // The frame and the window it goes into are written
                            // under the hide count's own lock, so a hide cannot
                            // land between the two — and a load the pointer has
                            // left in the meantime is neither painted nor shown
                            // (see `HIDDEN_EPOCH`).
                            {
                                let hidden = HIDDEN_EPOCH.lock().ok();
                                let wanted = pending
                                    .as_ref()
                                    .map(|pl| hover_still_wanted(&hidden, pl))
                                    .unwrap_or(true);

                                if wanted {
                                    match pending.as_ref().filter(|_| visible) {
                                        Some(pl) => {
                                            render_layered_preview_at(hwnd, pl.pos_x, pl.pos_y)
                                        }
                                        None => render_layered_preview(hwnd),
                                    }

                                    if let Some(pl) = pending {
                                        let _ = SetWindowPos(
                                            hwnd,
                                            HWND_TOPMOST,
                                            pl.pos_x,
                                            pl.pos_y,
                                            mw,
                                            mh,
                                            SWP_NOACTIVATE | SWP_SHOWWINDOW,
                                        );
                                        let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
                                    }
                                }
                            }

                            // A page for this document is one Office start away,
                            // so it is asked for as soon as the hover is up rather
                            // than after the pointer has rested on it.
                            page_render_pending = request_office_render(
                                &result.path,
                                result.generation,
                                render_box.0,
                                render_box.1,
                            );
                        }
                        None if result.awaiting_render => {
                            // Nothing to draw yet and a page on the way: the
                            // pending load stays armed, so the preview is not
                            // dropped and the page has somewhere to land — and the
                            // wait it is given is the one every other kind of wait
                            // gets, so the spinner goes up once the delay has run
                            // (see `spinner_due`). What is asked for is the room the
                            // display has rather than the box this hover's layout
                            // came out at, which for a wait is the spinner's own box
                            // at the pointer and says nothing about how large the
                            // page will be drawn: a slide is exported at the width
                            // the render is asked for, so the room the display has is
                            // the sharpest page that display can show (see
                            // `PendingLoad::room`).
                            let (width, height) = pending_load
                                .as_ref()
                                .map(|pl| pl.room)
                                .unwrap_or_else(|| office_formats::default_page_size(&result.path));

                            // The wait is on an engine from here, whatever was asked
                            // for below: what the loader could not draw is a page, a
                            // picture or a listing that only an engine can produce, and
                            // that is what the cap on waiting is read against — a wait
                            // whose request is answered as part of another's has nothing
                            // else that would ever end it (see `awaiting_engine`).
                            if let Some(pl) = pending_load.as_mut() {
                                pl.awaiting_engine = true;
                            }

                            // A document the render engine draws is asked for the same
                            // way, and in the same breath: neither page exists until an
                            // engine has drawn it, and this is the one hover that is
                            // waiting for one.
                            //
                            // And a picture the image converter develops is the same wait
                            // once more — the engine's answer arrives as a message rather
                            // than as a page in a folder, and what it writes is a picture
                            // at the size it is shown rather than at the size of the file,
                            // which is what makes its room a ceiling rather than a hint: a
                            // picture developed into a box smaller than the display can
                            // never be drawn any larger than that box.
                            // The engines this hover could be owed a page by, asked in the order
                            // the file's kind names them and only where one of them can answer:
                            // whichever starts work is the one being waited on, and the loop
                            // watches for the page or the picture or the listing it will leave
                            // (see `request_engine_render`).
                            let requested = request_engine_render(
                                &result.path,
                                result.generation,
                                (width, height),
                            );

                            // Nothing was asked for because there is nothing left to ask
                            // about: every engine that could owe this file something has
                            // already produced it. The page, the picture, the listing is
                            // in hand — it landed between the load that came back without
                            // it and this check, which is the one moment the two questions
                            // can disagree about — so the answer is read by having the
                            // file loaded again, and the replay below (`page_upgrade` and
                            // all) is the one a page arriving as a message is given. Left
                            // as a wait with nothing behind it, the hover would be a
                            // spinner that no answer and no cap could ever take down.
                            //
                            // A file the engine turned down while the load ran is answered
                            // by that same replay and by nothing else: the load that reads
                            // it finds no page and no wait to be in, which is the branch
                            // that takes the spinner down rather than leaving it up.
                            if requested.is_none() {
                                page_ready = Some((result.path.clone(), result.generation, true));
                            }

                            page_render_pending = requested;
                        }
                        None => {
                            // Loading failed, hide window
                            let _ = ShowWindow(hwnd, SW_HIDE);
                            clear_pointer_hold();
                            if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                if let Some(ref mut existing) = *current {
                                    existing.cancel_background_work();
                                }
                                *current = None;
                            }
                            pending_load = None;
                            pending_load_cancel = None;
                        }
                    }
                }
            }

            // A preview that is still on its way follows the pointer: while a
            // load runs the spinner is the only thing on screen, and a cursor
            // that moves along the item the preview belongs to would otherwise
            // leave it behind. Nothing is measured again — the hover's own size
            // and `Avoid` region are what it is placed with — so this
            // costs a cursor read and a placement per tick, and the window is
            // moved only when the place it comes out at has changed.
            //
            // The spinner follows at its own place rather than the preview's —
            // the arc at the pointer's corner, whatever box the preview will
            // arrive in — so what the hand sees while it waits is the wait at
            // the hand (see `waiting_placement`).
            if let Some(ref mut pl) = pending_load {
                if let Some(cursor) = cursor_position() {
                    let side_before = pl.spinner_side;
                    let dpi = monitor_dpi_from_point(cursor.x, cursor.y);
                    let followed = pl.follow_pointer(cursor, dpi);
                    if followed.spinner && pl.spinner_shown {
                        if pl.spinner_side == side_before {
                            let _ = MoveWindow(
                                hwnd,
                                pl.spinner_pos.0,
                                pl.spinner_pos.1,
                                pl.spinner_side as i32,
                                pl.spinner_side as i32,
                                false,
                            );
                        } else {
                            // The spinner's own box changed size, so its frame is
                            // drawn again at the size it now goes into.
                            show_loading_spinner(hwnd, pl);
                        }
                    }

                    // A wait for an engine-drawn preview takes the engine's window with it:
                    // what the engine is asked for is the box the wait has ended up in, and a
                    // box that moved while the document was on its way is the same document at
                    // the place the hand is — the want is moved rather than the engine being
                    // asked again, which is what keeps a moving pointer from asking for the
                    // same page sixty times a second, and what lets a document that arrives
                    // before its spinner was due land at the hand anyway. Nothing is moved for
                    // an engine that already has it up: that is the preview itself, and it
                    // follows nothing.
                    if followed.preview
                        && engine_kind_of(&pl.path).is_some()
                        && !webview_preview::is_showing()
                    {
                        webview_preview::wanted_here(
                            &pl.path,
                            webview_preview::Area {
                                x: pl.pos_x,
                                y: pl.pos_y,
                                width: pl.width as i32,
                                height: pl.height as i32,
                            },
                        );
                    }
                }
            }

            // Show the loading spinner while a background load runs, once the wait
            // is worth showing: a load that may be about to finish is given the
            // delay `spinner_delay_ms` names (see `spinner_due`).
            if let Some(ref mut pl) = pending_load {
                if pl.spinner_due() {
                    pl.spinner_shown = true;
                    show_loading_spinner(hwnd, pl);
                }
            }

            // A wheel notch over a scrollable text preview belongs to the preview:
            // the wheel hook swallowed it so Explorer does not scroll, and left
            // the ticks here. Rotating the wheel forward scrolls back towards the
            // start of the document, which is the opposite of the sign the
            // message carries.
            let scroll_delta = wheel_input::take_text_scroll_delta();
            if scroll_delta != 0 {
                let notches = (scroll_delta / WHEEL_DELTA) as i64;
                if notches != 0 {
                    let lines = -notches * TEXT_SCROLL_LINES_PER_NOTCH;
                    if let Some(first_line) = text_scroll_target(lines) {
                        scroll_text_preview(hwnd, first_line);
                    }
                }
            }

            // Check for our custom messages. Only the newest hover target matters;
            // collapse stale Show/Hide traffic so we do not spend time computing
            // layouts for files the cursor has already left.
            let mut latest_preview_msg: Option<PreviewMessage> = None;
            let mut refresh_requested = false;
            // A probe's answer, held apart the same way and for the same reason.
            let mut video_probed: Option<(PathBuf, u64)> = None;
            // And a measure's, which is held apart with the box it answered with: what is
            // waiting on it is a hover, and the box is what that hover is replayed with (see
            // `MeasureProbed`).
            let mut measure_probed: Option<(PathBuf, Option<(u32, u32)>)> = None;
            let mut next_preview_msg = carried_preview_msg.take();
            while let Some(preview_msg) = next_preview_msg.or_else(|| rx.try_recv().ok()) {
                next_preview_msg = None;

                match preview_msg {
                    PreviewMessage::Refresh => {
                        if latest_preview_msg.is_none() {
                            // A painted preview's colors are in its frame — the
                            // colors and glyphs of a text preview and the theme of
                            // an archive listing alike — so a theme switch rebuilds
                            // it from the hover it came from; every other preview
                            // only needs the frame composited again.
                            match (current_media_is_painted(), current_show.clone()) {
                                (true, Some(show)) => latest_preview_msg = Some(show),
                                _ => refresh_requested = true,
                            }
                        }
                    }
                    PreviewMessage::RefreshTypes => {
                        // A preview of a kind that was switched off is rebuilt
                        // from the hover it came from, which is what drops it:
                        // the layout finds no size for a file whose kind is off.
                        // Every other preview is left exactly as it is — a
                        // running video is not restarted by a toggle it has
                        // nothing to do with.
                        if latest_preview_msg.is_none() {
                            match (current_media_kind(), current_show.clone()) {
                                (Some(kind), Some(show)) if !kind.enabled() => {
                                    latest_preview_msg = Some(show)
                                }
                                // A preview the engine draws has no media of its own —
                                // the engine's window is the preview — so its kind is
                                // read from the file the hover is about rather than
                                // from what is on screen.
                                (None, Some(show))
                                    if show_path(&show)
                                        .and_then(|path| engine_kind_of(path))
                                        .is_some_and(|kind| !kind.enabled()) =>
                                {
                                    latest_preview_msg = Some(show)
                                }
                                _ => {}
                            }
                        }
                    }
                    PreviewMessage::OfficeRenderReady {
                        path,
                        generation,
                        ok,
                    } => {
                        // The newest hover wins, as it does over every other
                        // message: a page that lands in the same tick as a new
                        // hover is not the answer to it.
                        if latest_preview_msg.is_none() && page_ready.is_none() {
                            page_ready = Some((path, generation, ok));
                        }
                    }
                    PreviewMessage::VideoProbed { path, generation } => {
                        // Held apart the way a render's answer is: it is not a hover to
                        // act on but an answer about the one that is waiting.
                        if latest_preview_msg.is_none() && video_probed.is_none() {
                            video_probed = Some((path, generation));
                        }
                    }
                    PreviewMessage::MeasureProbed { path, size } => {
                        // And a measured box, held apart with the box itself: what is waiting
                        // on it is a hover, and the box is what that hover is laid out with
                        // (see `measure_probed`).
                        if latest_preview_msg.is_none() && measure_probed.is_none() {
                            measure_probed = Some((path, size));
                        }
                    }
                    PreviewMessage::MagickReady {
                        path,
                        generation,
                        ok,
                    } => {
                        // The engine's answer, held apart for the reason the render tier's
                        // is: what the hover that asked is waiting for is a picture to be
                        // placed with, not another hover — and an engine that will not draw
                        // the file is the same wait answered, with nothing in it.
                        if latest_preview_msg.is_none() && page_ready.is_none() {
                            page_ready = Some((path, generation, ok));
                        }
                    }
                    PreviewMessage::PeazipReady {
                        path,
                        generation,
                        ok,
                    } => {
                        // And a listing, which is the same answer once more: what the hover is
                        // waiting for is a table of contents to draw a page from, and an archive
                        // the engine will not list is that wait answered with nothing.
                        if latest_preview_msg.is_none() && page_ready.is_none() {
                            page_ready = Some((path, generation, ok));
                        }
                    }
                    other => {
                        latest_preview_msg = Some(other);
                        refresh_requested = false;
                    }
                }
            }

            // A page the render engine draws is not messaged about the way an Office page
            // is: what that engine writes is a file under the app's own folder, so whether
            // the page has arrived — or whether the engine has answered that it will not
            // draw the document at all — is a read of that folder rather than a message
            // from a thread. What comes of it is the answer an Office page gives, and the
            // same code below takes it up: it is the same wait, in the same box. The page a book
            // is converted into is the second engine that answers this way, and it is read here
            // for the same reason.
            //
            // Which documents are watched for is asked of the place the request was made
            // from, so a hover is never watched for a page nothing was asked to draw: an
            // Office document whose own application is here has a page asked of that
            // application's tier, and is answered by a message rather than by this read (see
            // `libre_render_is_due`).
            if page_ready.is_none() {
                if let Some((path, generation)) = page_render_pending.as_ref() {
                    if let Some(drawn) = engine_page_answer(path) {
                        page_ready = Some((path.clone(), *generation, drawn));
                    }
                }
            }

            // A page the render tier has finished with, for the hover that asked
            // for it: that hover is replayed, which measures the page itself and
            // draws it in place of the spinner it supersedes. A payload from an
            // older hover is dropped here — the page it wrote is kept as far as the
            // cache budget allows, and no further.
            if let Some((ready_path, ready_generation, ready_ok)) = page_ready {
                let shown = current_show.as_ref().and_then(show_path);
                let hovered = ready_generation == current_generation
                    && shown.map(|path| path.as_path()) == Some(ready_path.as_path());

                // Which request the answer belongs to, so that a page landing for a hover
                // that has gone cannot clear a wait another hover is still in.
                let answers_the_request =
                    page_render_pending
                        .as_ref()
                        .is_some_and(|(path, generation)| {
                            *path == ready_path && *generation == ready_generation
                        });

                // An answer for the file the hover on screen is waiting on, naming the
                // hover before it, is that wait's own answer: the request it made was
                // folded into the work this one belongs to, and what it answers is the
                // page, picture or listing the wait is for (see `answer_belongs_to_the_wait`).
                let waiting_for_this_file = !hovered
                    && answer_belongs_to_the_wait(
                        &ready_path,
                        ready_ok,
                        shown.map(|path| path.as_path()),
                        pending_load.as_ref(),
                    );

                if (hovered || waiting_for_this_file) && ready_ok {
                    if latest_preview_msg.is_none() {
                        // The wait is over, so the request stops being the pending
                        // one here — where the page is actually taken up.
                        page_render_pending = None;
                        // The page is there, so the hover is replayed: that measures
                        // the page itself, moves the window to its size and loads it.
                        // It is an upgrade rather than a new preview, though, and what
                        // is on screen stays while it happens — hiding the spinner for
                        // the second a large picture takes to decode is a blink the
                        // user sees and reads as the preview failing.
                        page_upgrade = Some(ready_path.clone());
                        // A mouse hover is replayed where the pointer is now: the
                        // spinner it replaces was kept with the pointer while the
                        // render ran, and a page that jumped back to where the
                        // hover started would jump away from where it was waited
                        // for.
                        latest_preview_msg = replay_where_the_pointer_is(current_show.clone());
                    }
                    // A newer message was in hand, so the page is not shown now. The
                    // wait it was rendered for is left standing rather than cleared
                    // with it: what is being waited on is still that page, so the cap
                    // on waiting keeps something to measure and the spinner comes down
                    // when its time is up instead of hanging there for good.
                } else if hovered {
                    // Nothing was drawn and nothing is coming. A preview that is
                    // not a spinner is kept — it is a preview like any other —
                    // while a spinner has nothing left to stand in for.
                    let showing_spinner = CURRENT_MEDIA
                        .lock()
                        .map(|media| {
                            media
                                .as_ref()
                                .map(|media| media.media_type.is_loading())
                                .unwrap_or(false)
                        })
                        .unwrap_or(false);
                    if showing_spinner {
                        let _ = ShowWindow(hwnd, SW_HIDE);
                        if let Ok(mut current) = CURRENT_MEDIA.lock() {
                            *current = None;
                        }
                    }
                    // The wait is over whether or not its spinner had gone up: an engine
                    // that turns a file down in less than the spinner's delay — a raw
                    // file whose length works out to no picture, a converter that finds
                    // nothing to read — answers before there is anything to take down,
                    // and the pending load is what the spinner would go up for. Left
                    // armed it goes up a moment later for a preview that is not coming
                    // and stays up, which is the one thing the entry above cannot undo:
                    // the answer it belonged to has been read already.
                    pending_load = None;
                    page_render_pending = None;
                } else {
                    // The hover this page was rendered for is over: it landed after
                    // the pointer had moved on, so nothing is waiting for it. What was
                    // rendered is kept as far as the budget allows and no further — at
                    // a size of nothing it is dropped here rather than held for a
                    // hover that has already gone.
                    //
                    // Only the wait this answer belongs to is cleared with it: a hover
                    // that is still waiting for a page of its own keeps the request it
                    // was made for, which is what the cap on waiting is read against.
                    if answers_the_request {
                        page_render_pending = None;
                    }
                    office_render::hover_ended(&ready_path);
                }
            }

            // A video's probe has answered for the hover that was waiting on it: that
            // hover is replayed — this time the layout has the shape the probe cached, so
            // the hover goes on as the video it is — and the wait it is in goes on as the
            // video's own rather than as a new preview (see `video_probe` and
            // `video_replay`). An answer for a hover that has gone is dropped here: what
            // the probe measured is held for the next hover of the file either way.
            if let Some((probed_path, probed_generation)) = video_probed {
                let waiting = video_probe.as_ref().is_some_and(|(path, generation)| {
                    *path == probed_path && *generation == probed_generation
                });

                if waiting {
                    video_probe = None;

                    if probed_generation == current_generation && latest_preview_msg.is_none() {
                        // The hover is replayed where the pointer is now, the way a page
                        // landing on a spinner is: the wait was kept with the pointer
                        // while the probe ran, so the video should be too.
                        video_replay = Some(probed_path);
                        latest_preview_msg = replay_where_the_pointer_is(current_show.clone());
                    }
                }
            }

            // A box measured off the preview thread has answered for the hover that was waiting
            // on it: that hover is replayed — this time the measure reads the box off the table
            // rather than out of the file, so the hover goes on as the preview it is — and the
            // wait it is in goes on as the preview's own rather than as a new one (see
            // `measured_off_the_tick` and `measure_replay`). An answer for a hover that has gone
            // is dropped: what the measure answered is held for the next hover of the file
            // either way.
            if let Some((measured_path, size)) = measure_probed {
                let waiting = current_show.as_ref().and_then(show_path) == Some(&measured_path);

                if waiting && latest_preview_msg.is_none() {
                    if size.is_some() {
                        measure_replay = Some(measured_path);
                        latest_preview_msg = replay_where_the_pointer_is(current_show.clone());
                    } else {
                        // The reader has nothing for this file, which is not a wait that can be
                        // answered: the spinner comes down rather than standing over nothing
                        // until the pointer moves, the way it does for an engine that cannot
                        // draw the file it was asked for.
                        let waiting_on_it = pending_load
                            .as_ref()
                            .is_some_and(|pl| pl.path == measured_path);

                        if waiting_on_it {
                            pending_load = None;
                            pending_load_cancel = None;
                            let _ = ShowWindow(hwnd, SW_HIDE);
                            clear_pointer_hold();

                            if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                *current = None;
                            }
                        }
                    }
                }
            }

            // The display under the preview changed and the frame that was on screen
            // went with it. The window proc could only discard what was drawn; the
            // hover it came from is what knows how to draw it again, at the scale of
            // the display the pointer is on now. A newer message in hand is left to
            // speak for itself, and the flag waits for a tick where none does.
            if latest_preview_msg.is_none() && DISPLAY_RESET.swap(false, Ordering::AcqRel) {
                latest_preview_msg = replay_where_the_pointer_is(current_show.clone());
            }

            // The engine has something to answer for: it could not be had at all, or it
            // could not put the file of the hover that is up into its window. Either way
            // this app draws none of it itself, so there is nothing to fall back to and
            // nothing left to wait for — the wait goes rather than standing as a spinner
            // over nothing.
            if latest_preview_msg.is_none() && webview_preview::take_failure_notice() {
                let waiting_on_the_engine = pending_load
                    .as_ref()
                    .is_some_and(|pl| engine_kind_of(&pl.path).is_some());

                if waiting_on_the_engine && !webview_preview::is_showing() {
                    pending_load = None;
                    pending_load_cancel = None;
                    let _ = ShowWindow(hwnd, SW_HIDE);
                    clear_pointer_hold();

                    if let Ok(mut current) = CURRENT_MEDIA.lock() {
                        *current = None;
                    }
                }
            }

            // Anchored one read before the layout it decides, so the box is placed clear of
            // the hand that asked for this hover rather than of the one it was measured from
            // (see `replay_where_the_pointer_is`).
            if let Some(preview_msg) = replay_where_the_pointer_is(latest_preview_msg) {
                // Common variables for Show/ShowKeyboard - set in match, used after
                let mut show_path: Option<PathBuf> = None;
                let mut show_layout: Option<PreviewLayout> = None;
                let mut show_spinner_layout: Option<PreviewLayout> = None;
                let mut show_placement: Option<HoverPlacement> = None;
                let mut show_is_video: bool = false;
                // Whether this hover is waiting on a video's probe rather than on the
                // video itself (see `video_probe_due`).
                let mut show_video_probe: bool = false;
                // Whether this hover is waiting on a measure that reads the file rather than on
                // the file itself (see `measure_waiting`).
                let mut show_measure_probe: bool = false;
                let mut show_requested = false;
                let mut preview_scale = current_hover_scales().picture;
                let mut show_dpi = 96u32;
                // The room the display the hover is on has, which is what an engine that
                // has to draw the preview before the file can be measured is asked for.
                // It is the display's room rather than the one this hover's layout comes
                // out at, which for a hover waiting on an engine is the corner of the
                // display the spinner was put in (see `PendingLoad::room`).
                let mut show_room: Option<(u32, u32)> = None;
                let show_snapshot = matches!(
                    preview_msg,
                    PreviewMessage::Show(..) | PreviewMessage::ShowKeyboard(..)
                )
                .then(|| preview_msg.clone());

                match preview_msg {
                    PreviewMessage::Show(path, x, y, avoid) => {
                        show_requested = true;
                        // Remember where this preview was opened from: the region
                        // that keeps a scrollable preview alive stretches from
                        // here to the preview, so the pointer can travel between
                        // the two without losing it.
                        set_text_scroll_anchor(x, y);

                        let bounds = monitor_bounds_from_point(x, y);
                        let dpi = monitor_dpi_from_point(x, y);
                        let follow_cursor = CONFIG.lock().map(|c| c.follow_cursor).unwrap_or(true);
                        preview_scale = effective_preview_scale(&path, current_hover_scales());

                        // A document with no page rendered for it yet has nothing to
                        // measure but the wait, so its preview is laid out as the
                        // spinner's own box (see `office_preview::measure` and the
                        // `libre` arm of `get_media_dimensions`) and that box is placed
                        // at the pointer's own corner rather than a preview's way out
                        // beside it: the page it waits for is laid out again by the replay
                        // that arrives with it, so until then the hover is the wait for the
                        // file under the hand, which belongs at the hand.
                        let waiting_spinner = page_is_on_the_way(&path);

                        // A video whose shape the probe has not answered for yet is the
                        // same kind of wait, and for the same reason: there is nothing to
                        // lay out as a video until the probe answers, so the hover is the
                        // wait for it — the spinner at the pointer's own corner — and is
                        // replayed when the answer lands (see `video_probe_due`).
                        let probing = video_probe_due(&path);

                        if let Some(orig_dims) = media_dimensions(&path, bounds, dpi) {
                            // A box that is being read is placed at the pointer's own corner the
                            // way every other wait is: what is on screen is the spinner for a
                            // measure the layout has just started, and it belongs at the hand
                            // that asked (see `measure_waiting`).
                            let measuring = measure_waiting(&path);
                            let is_video = drawn_as_video(&path);
                            let mut placement = HoverPlacement {
                                orig_dims,
                                avoid,
                                follow_cursor,
                                preview_scale,
                                at_the_pointer_corner: waiting_spinner || probing || measuring,
                            };
                            let placed = compute_mouse_layout(x, y, placement, bounds, dpi);
                            if let Some(layout) = placed {
                                let (layout, text_size) =
                                    text_preview_layout(&path, layout, dpi, |size| {
                                        compute_mouse_layout(
                                            x,
                                            y,
                                            HoverPlacement {
                                                orig_dims: size,
                                                ..placement
                                            },
                                            bounds,
                                            dpi,
                                        )
                                    });
                                // A text preview is placed again at the width its box came
                                // out with, and the frame that lands is taller by the rows a
                                // long line wraps into there. The wait re-places from the size
                                // kept here, so it keeps the re-measured one: the display's own
                                // measurement would step the wait around the name for a height
                                // the frame does not have (see `text_preview_layout`).
                                if let Some(size) = text_size {
                                    placement.orig_dims = size;
                                }
                                show_is_video = is_video;
                                show_video_probe = probing;
                                show_measure_probe = measuring;
                                // The display the hover is on, which is what a card of a sound
                                // is painted at and what its clock repaints it at.
                                audio_card_dpi = dpi;
                                show_layout = Some(layout);
                                show_placement = Some(placement);
                                // The wait for this hover is the spinner's own box at
                                // the pointer, whatever the preview's own place came
                                // out at (see `waiting_placement`).
                                show_spinner_layout = compute_mouse_layout(
                                    x,
                                    y,
                                    waiting_placement(placement),
                                    bounds,
                                    dpi,
                                );
                                show_path = Some(path);
                                show_dpi = dpi;
                                show_room = Some(bounds.room());
                            }
                        }
                    }
                    PreviewMessage::ShowKeyboard(path, il, it, ir, ib, avoid, columns) => {
                        show_requested = true;
                        // The focused item lives inside the Explorer window, so
                        // its center resolves to that window's monitor.
                        let center = ((il + ir) / 2, (it + ib) / 2);
                        set_text_scroll_anchor(center.0, center.1);

                        let bounds = monitor_bounds_from_point(center.0, center.1);
                        let dpi = monitor_dpi_from_point(center.0, center.1);
                        let follow_cursor = CONFIG.lock().map(|c| c.follow_cursor).unwrap_or(true);
                        preview_scale = effective_preview_scale(&path, current_hover_scales());

                        if let Some(orig_dims) = media_dimensions(&path, bounds, dpi) {
                            let is_video = drawn_as_video(&path);
                            let placement = KeyboardPlacement {
                                item_rect: (il, it, ir, ib),
                                avoid,
                                columns,
                                orig_dims,
                                follow_cursor,
                                preview_scale,
                            };
                            if let Some(layout) = compute_keyboard_layout(placement, bounds, dpi) {
                                let (layout, _) = text_preview_layout(&path, layout, dpi, |size| {
                                    compute_keyboard_layout(
                                        KeyboardPlacement {
                                            orig_dims: size,
                                            ..placement
                                        },
                                        bounds,
                                        dpi,
                                    )
                                });
                                show_is_video = is_video;
                                show_video_probe = video_probe_due(&path);
                                show_measure_probe = measure_waiting(&path);
                                audio_card_dpi = dpi;
                                show_layout = Some(layout);
                                // A keyboard hover has no pointer for a wait to be
                                // placed at, so its spinner is the arc's own box
                                // beside the item, the way its preview is.
                                show_spinner_layout = compute_keyboard_layout(
                                    KeyboardPlacement {
                                        orig_dims: (
                                            office_preview::WAITING_BOX,
                                            office_preview::WAITING_BOX,
                                        ),
                                        preview_scale: PreviewScale::Percent(100),
                                        ..placement
                                    },
                                    bounds,
                                    dpi,
                                );
                                show_path = Some(path);
                                show_dpi = dpi;
                                show_room = Some(bounds.room());
                            }
                        }
                    }
                    PreviewMessage::Hide => {
                        // Invalidate any pending background loads
                        current_generation += 1;
                        pending_load = None;
                        clear_load_request(&load_request_slot);
                        if let Some(cancel) = pending_load_cancel.take() {
                            cancel.store(true, Ordering::Release);
                        }

                        let _ = ShowWindow(hwnd, SW_HIDE);
                        clear_pointer_hold();

                        // A document the engine is playing goes with the preview: its
                        // window is its own, so nothing else here takes it down.
                        webview_preview::hide();

                        // Stop video playback if any
                        if let Ok(mut current) = CURRENT_MEDIA.lock() {
                            if let Some(ref mut media) = *current {
                                media.cancel_background_work();
                                stop_video_playback(media);
                            }
                            *current = None;
                        }
                        current_video_path = None;
                        video_pos = (0, 0, 0, 0);

                        // The page held for the preview that is going away is no
                        // longer being waited on: at a budget of nothing it is
                        // dropped here rather than kept until something else is
                        // rendered to make room for. What is on screen is read from
                        // the show this message takes down — the message itself is a
                        // hide, and `show_path` names this message's own path here.
                        if let Some(path) = current_show.as_ref().and_then(self::show_path) {
                            office_render::hover_ended(path);
                        }

                        current_show = None;
                        page_render_pending = None;
                        page_upgrade = None;
                    }
                    PreviewMessage::Refresh => {
                        render_layered_preview(hwnd);
                    }
                    // A type toggle is answered by the receive loop above, which
                    // is where the kind of the preview on screen is known; it
                    // replays the hover instead of arriving here as itself.
                    PreviewMessage::RefreshTypes => {}
                    // Likewise answered above: a rendered page replays the hover
                    // it belongs to rather than being handled as a message here.
                    PreviewMessage::OfficeRenderReady { .. } => {}
                    // And a probe's answer, which replays the hover that was waiting
                    // on it the same way.
                    PreviewMessage::VideoProbed { .. } => {}
                    // And a measured box, which replays the hover that was waiting on it in
                    // the same way — or takes its wait down, where the answer is that there is
                    // no box to be had.
                    PreviewMessage::MeasureProbed { .. } => {}
                    // And an engine's, which is the page-shaped answer of a picture
                    // rather than of a page: it replays the hover that was waiting on
                    // it exactly as the render tier's answer does.
                    PreviewMessage::MagickReady { .. } => {}
                    // And the listing an engine produced, which is the same shape of answer
                    // once more: a page's worth of content arriving for the hover that asked
                    // for it, replayed rather than handled as a hover here.
                    PreviewMessage::PeazipReady { .. } => {}
                }

                // Shared load/display logic for Show and ShowKeyboard
                if let (Some(path), Some(layout)) = (show_path, show_layout) {
                    let pos_x = layout.pos_x;
                    let pos_y = layout.pos_y;
                    let media_width = layout.preview_w as i32;
                    let media_height = layout.preview_h as i32;
                    let max_width = layout.max_width;
                    let max_height = layout.max_height;
                    let preview_w = layout.preview_w;
                    let preview_h = layout.preview_h;

                    // The room an engine-drawn preview is asked for is the room the display
                    // has rather than the one this layout came out at: a hover waiting on an
                    // engine is laid out as the spinner's own box at the pointer, and the
                    // room that layout comes out at is the corner the spinner was put in —
                    // which for a picture developed at the size it is shown at would be a
                    // ceiling on the size it could ever be drawn at (see `PendingLoad::room`).
                    let room = show_room.unwrap_or((max_width, max_height));

                    // The box the wait goes in while the load runs: the spinner's own
                    // place, which is not the preview's. A hover with no room for the
                    // spinner anywhere is answered with the preview's own box rather
                    // than with no spinner at all.
                    let (spinner_x, spinner_y, spinner_side) = match show_spinner_layout {
                        Some(spinner) => (spinner.pos_x, spinner.pos_y, spinner.preview_w),
                        None => (pos_x, pos_y, preview_w),
                    };

                    // A page that arrived for the hover already on screen is an
                    // upgrade: what is there — the spinner — stays up while the page
                    // is loaded, and is replaced when it lands. A hover replayed for a
                    // video's probe, or for a box measured off this thread, is the same thing
                    // reached another way: it is the wait it was already in, carried on.
                    let upgrading = page_upgrade.as_deref() == Some(path.as_path())
                        || video_replay.as_deref() == Some(path.as_path())
                        || measure_replay.as_deref() == Some(path.as_path());
                    page_upgrade = None;
                    video_replay = None;
                    measure_replay = None;

                    // A text or archive preview is rendered at the size the
                    // layout planned for it: it is painted at a fixed font
                    // size, so the planned box is the box it draws into rather
                    // than a space to be scaled within — and what the window
                    // is sized to is the frame that comes back. Every other
                    // format is loaded against the free space it may be
                    // scaled within.
                    let painted = page_is_painted(&path);
                    let (load_width, load_height) = if painted {
                        (preview_w, preview_h)
                    } else {
                        (max_width, max_height)
                    };

                    if show_snapshot.is_some() {
                        current_show = show_snapshot.clone();
                    }

                    // A video is played by `ffplay` when FFmpeg is installed, and by the
                    // media engine Windows has when it is not. The two take different
                    // roads from here: FFmpeg's player is its own window, which is what
                    // the branch below puts up, while the engine's frames come back
                    // through the ordinary load and are drawn by this app's own window.
                    let ffplay_plays_video = show_is_video && codecs::ffplay_available();

                    if show_video_probe || show_measure_probe {
                        // The hover is waiting on a probe: nothing of the file can be
                        // laid out or loaded until there is a shape or a box to lay it out with,
                        // so what is put up is the wait every other preview is given —
                        // the spinner at the pointer — and the hover it came from is
                        // replayed when the answer lands, which is when there is a video
                        // to load or a box to lay out (see `video_probe_due`, `video_probe`
                        // and `measure_waiting`). The probe itself runs on a thread of its
                        // own: for a video it is two external processes, and for a box it is a
                        // read of the file, and this is the thread that draws the wait.
                        //
                        // A replay carries the wait it was already in rather than opening
                        // a new one: the same clock — so a probe answered inside
                        // `spinner_delay_ms` does not start that delay over — and the same
                        // spinner, which is on screen already.
                        let waiting = pending_load.take();
                        let (started, spinner_shown, upgrade) = match (upgrading, waiting) {
                            (true, Some(pl)) => (pl.started, pl.spinner_shown, true),
                            _ => (Instant::now(), false, false),
                        };

                        current_generation += 1;
                        let gen = current_generation;
                        clear_load_request(&load_request_slot);
                        if let Some(cancel) = pending_load_cancel.take() {
                            cancel.store(true, Ordering::Release);
                        }

                        if !upgrade {
                            // Another preview takes over here, so what is on screen goes
                            // as the wait for the probe goes up — the frame this app's
                            // window holds, the player a previous hover started, and the
                            // window a document the engine draws was put in.
                            if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                if let Some(ref mut media) = *media_guard {
                                    media.cancel_background_work();
                                    stop_video_playback(media);
                                }
                                // Clear immediately so old pixels never flash while
                                // the new target is being measured.
                                *media_guard = None;
                            }

                            if current_video_path.is_some() {
                                current_video_path = None;
                                video_pos = (0, 0, 0, 0);
                            }

                            webview_preview::hide();

                            let _ = ShowWindow(hwnd, SW_HIDE);
                        }

                        pending_load = Some(PendingLoad {
                            generation: gen,
                            hide_epoch: hidden_epoch(),
                            path: path.clone(),
                            started,
                            pos_x,
                            pos_y,
                            width: preview_w,
                            height: preview_h,
                            room,
                            spinner_shown,
                            spinner_delay: load_spinner_delay(),
                            spinner_pos: (spinner_x, spinner_y),
                            spinner_side,
                            placement: show_placement,
                            upgrade,
                            // The probe a video waits on is marked as an engine's wait is:
                            // what it is waiting for is outside this side, and what is read
                            // against this flag is the cap on a wait that nothing else ends
                            // (see `awaiting_engine`). A box measured off this thread is
                            // not marked — what it waits for is a read that finishes
                            // (see `measured_off_the_tick`).
                            awaiting_engine: show_video_probe,
                        });
                        // The probe a hover is waiting on is the one this message asked for: a
                        // measured box started its own thread where it was measured, so what is
                        // left to start here is a video's (see `measured_off_the_tick`).
                        if show_video_probe {
                            video_probe = Some((path.clone(), gen));
                            spawn_video_probe(path, gen);
                        }
                    } else if ffplay_plays_video {
                        // Cancel any in-flight image load before switching to video.
                        current_generation += 1;
                        pending_load = None;
                        clear_load_request(&load_request_slot);
                        if let Some(cancel) = pending_load_cancel.take() {
                            cancel.store(true, Ordering::Release);
                        }

                        let no_cancel = Arc::new(AtomicBool::new(false));
                        if let Some(media_data) = load_media(
                            &path,
                            load_width,
                            load_height,
                            preview_scale,
                            show_dpi,
                            no_cancel,
                        ) {
                            // Nothing of this app's is on screen for a video — the
                            // player draws it in a window of its own — so what is put
                            // up while the player starts is the wait every other kind
                            // of preview is given: the spinner at the pointer, and it
                            // goes the moment the player's window is there (see
                            // `video_start` and `player_wait`). A hover replayed for a
                            // probe is already that wait, and it stays where it is.
                            if !upgrading {
                                let _ = ShowWindow(hwnd, SW_HIDE);
                            }

                            let process_running = is_video_process_running();
                            let should_start =
                                current_video_path.as_ref() != Some(&path) || !process_running;

                            if should_start {
                                if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                    if let Some(ref mut media) = *media_guard {
                                        media.cancel_background_work();
                                        stop_video_playback(media);
                                    }
                                }

                                // The previous ffplay may have survived its stop
                                // (dropped handle or an unconfirmed kill): kill it
                                // before a new one takes the screen.
                                kill_stray_video_process();

                                let video_process = start_video_playback(
                                    &path,
                                    pos_x,
                                    pos_y,
                                    media_width,
                                    media_height,
                                );
                                let pid =
                                    video_process.as_ref().map(|child| child.id()).unwrap_or(0);

                                current_video_path = Some(path.clone());
                                video_pos = (pos_x, pos_y, media_width, media_height);
                                let _ = ensure_video_window_topmost(
                                    pos_x,
                                    pos_y,
                                    media_width,
                                    media_height,
                                );

                                let mut data = media_data;
                                data.video_process = video_process;

                                if pid != 0 {
                                    // The player is a process with a window to create
                                    // before anything of the file is on screen, and
                                    // there is nothing under the spinner that could
                                    // arrive sooner — a player that starts instantly is
                                    // the only start this wait is not seen for — so it
                                    // is shown from the first tick rather than after
                                    // `spinner_delay_ms`: a frame of the spinner is what
                                    // a video hover would otherwise spend showing the
                                    // desktop.
                                    video_start = Some(VideoStart {
                                        media: data,
                                        path: path.clone(),
                                        pid,
                                        started: Instant::now(),
                                    });
                                    pending_load = Some(PendingLoad {
                                        generation: current_generation,
                                        hide_epoch: hidden_epoch(),
                                        path: path.clone(),
                                        started: Instant::now(),
                                        pos_x,
                                        pos_y,
                                        width: media_width as u32,
                                        height: media_height as u32,
                                        room,
                                        spinner_shown: false,
                                        spinner_delay: Duration::ZERO,
                                        spinner_pos: (spinner_x, spinner_y),
                                        spinner_side,
                                        placement: show_placement,
                                        upgrade: false,
                                        awaiting_engine: false,
                                    });
                                } else if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                    // A player that would not start is no preview:
                                    // nothing of this app's goes up for one.
                                    *current = Some(data);
                                }
                            } else {
                                video_pos = (pos_x, pos_y, media_width, media_height);
                                let _ = ensure_video_window_topmost(
                                    pos_x,
                                    pos_y,
                                    media_width,
                                    media_height,
                                );
                            }
                        }
                    } else {
                        // For images/animations, load async
                        if current_video_path.is_some() {
                            if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                if let Some(ref mut media) = *media_guard {
                                    media.cancel_background_work();
                                    stop_video_playback(media);
                                }
                            }
                            current_video_path = None;
                            video_pos = (0, 0, 0, 0);
                        }

                        // Any preview the engine does not draw is drawn here, so its window
                        // — if one is still up — comes down as this one goes up.
                        if engine_kind_of(&path).is_none() {
                            webview_preview::hide();
                        }

                        if let Some(cancel) = pending_load_cancel.take() {
                            cancel.store(true, Ordering::Release);
                        }
                        if !upgrading {
                            if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                if let Some(ref mut media) = *media_guard {
                                    media.cancel_background_work();
                                    // What is being replaced goes with it, and for a sound
                                    // that is more than bookkeeping: the engine playing one
                                    // is this app's own and this thread's alone, so a media
                                    // dropped here without this call plays on with nothing on
                                    // screen that could stop it. The take-down the `Hide`
                                    // would have performed is not a second chance at it: a
                                    // keyboard preview replacing a hovered one sends its
                                    // `Hide` and its show in the same breath, and the drain
                                    // above keeps only the newest of them (see the collapse
                                    // there). Every other branch that replaces a preview
                                    // stops what it replaces; this one did not.
                                    stop_video_playback(media);
                                }
                                // Clear immediately so old pixels never flash while
                                // the new target is being decoded.
                                *media_guard = None;
                            }
                            let _ = ShowWindow(hwnd, SW_HIDE);
                        }

                        // Start background load; the spinner follows if the wait
                        // turns out to be worth showing (see `spinner_due`).
                        current_generation += 1;
                        let gen = current_generation;
                        let load_cancel = Arc::new(AtomicBool::new(false));
                        pending_load_cancel = Some(Arc::clone(&load_cancel));
                        pending_load = Some(PendingLoad {
                            generation: gen,
                            hide_epoch: hidden_epoch(),
                            path: path.clone(),
                            started: Instant::now(),
                            pos_x,
                            pos_y,
                            width: preview_w,
                            height: preview_h,
                            room,
                            spinner_shown: false,
                            spinner_delay: load_spinner_delay(),
                            spinner_pos: (spinner_x, spinner_y),
                            spinner_side,
                            placement: show_placement,
                            upgrade: upgrading,
                            awaiting_engine: false,
                        });

                        queue_load_request(
                            &load_request_slot,
                            LoadRequest {
                                generation: gen,
                                path,
                                max_width: load_width,
                                max_height: load_height,
                                preview_scale,
                                dpi: show_dpi,
                                cancel: Arc::clone(&load_cancel),
                            },
                        );
                    }
                } else if show_requested {
                    // A newer hover target could not produce a layout/path. Treat it
                    // like a hide so stale async loads cannot resurrect old previews.
                    current_generation += 1;
                    pending_load = None;
                    clear_load_request(&load_request_slot);
                    if let Some(cancel) = pending_load_cancel.take() {
                        cancel.store(true, Ordering::Release);
                    }

                    let _ = ShowWindow(hwnd, SW_HIDE);
                    webview_preview::hide();

                    if let Ok(mut current) = CURRENT_MEDIA.lock() {
                        if let Some(ref mut media) = *current {
                            media.cancel_background_work();
                            stop_video_playback(media);
                        }
                        *current = None;
                    }
                    current_video_path = None;
                    video_pos = (0, 0, 0, 0);
                    current_show = None;
                    page_render_pending = None;
                    page_upgrade = None;
                }
            } else if refresh_requested {
                render_layered_preview(hwnd);
            }

            // Ctrl+C over a text preview copies what is selected in it. The key
            // is polled rather than waited for: the preview never takes focus, so
            // it would never receive the keystroke as a message. Selecting
            // everything is the menu's `Select All`, not a key of its own.
            if text_preview_copy_requested() {
                copy_text_preview(hwnd);
            }

            // The cap on waiting for a page is read against the hover on screen
            // rather than remembered: a render that outlives its hover is dropped
            // here and its page left in the cache — for Office's tier as much as for
            // the render engine, whose conversion runs on and is kept the same way.
            let shown_path = current_show.as_ref().and_then(show_path).cloned();

            // How long the hover on screen has been waiting, which is what the cap is
            // read of, and whether what it is waiting on is an engine at all.
            let (waited, awaiting_engine) = pending_load
                .as_ref()
                .map(|pl| {
                    (
                        pl.started.elapsed() >= Duration::from_secs(OFFICE_RENDER_WAIT_SECS),
                        pl.awaiting_engine,
                    )
                })
                .unwrap_or((false, false));

            // An engine that draws the hover's page is asked to come up once the pointer has
            // settled on the file: what the ask overlaps is the rest of the wait — a launch is
            // a second or more — and a file the pointer is merely crossing costs nothing, since
            // the ask is not made until the hover has outlasted the settle (see `WARM_SETTLE_MS`
            // and `warm_engines_for`). Once per hover: an engine is asked for the file under the
            // hand, and nothing is gained by asking again on every tick.
            if warmed_generation != Some(current_generation) {
                if let Some(pl) = pending_load
                    .as_ref()
                    .filter(|pl| !pl.upgrade && pl.started.elapsed() >= WARM_SETTLE_MS)
                {
                    warmed_generation = Some(current_generation);
                    warm_engines_for(&pl.path);
                }
            }

            let render_wait = page_render_pending.as_ref().map(|(path, generation)| {
                *generation == current_generation && shown_path.as_deref() == Some(path.as_path())
            });

            // A request made for a hover that has gone is not waited for any longer: the
            // page it produces is kept by the engine and the next hover of the file reads
            // it. A wait the hover on screen is still in is left standing.
            if render_wait == Some(false) {
                page_render_pending = None;
            }

            // The preview has waited as long as it waits — a page that has not arrived by
            // now may still be coming, a very large document takes as long as it takes —
            // so the spinner comes down and the hover is left to itself. The render is not
            // abandoned with it: it runs on and its page is cached, so the next hover of
            // that file shows it. Only the engine itself can say a render failed, and it
            // remembers that for the file.
            //
            // What is read here is the wait rather than the request. An engine answers a
            // file once, so a hover whose request was folded into one already in flight is
            // never answered on its own — its request has been read and left, and a cap
            // that was read of the request would be no cap at all: the spinner would stand
            // there for good (see `awaiting_engine`).
            if waited && awaiting_engine {
                page_render_pending = None;
                pending_load = None;
                if let Some(cancel) = pending_load_cancel.take() {
                    cancel.store(true, Ordering::Release);
                }
                let _ = ShowWindow(hwnd, SW_HIDE);
                if let Ok(mut current) = CURRENT_MEDIA.lock() {
                    *current = None;
                }
            }

            // Keep the pointer region in step with the window rather than only
            // with the paints: the window is moved when a frame is installed, and
            // the Explorer hook reads this on every one of its own ticks.
            publish_pointer_hold(hwnd);

            if current_show.is_none() && pending_load.is_none() {
                // Nothing is on screen, which is where this thread used to spend
                // its whole life waking sixty times a second to drain a queue that
                // was empty and publish a region that did not exist. There is
                // nothing to animate, nothing to repaint and nothing left to keep
                // in step, so the wait becomes the preview channel itself: a hover
                // is answered as it arrives rather than on the next tick, and the
                // interval is only a ceiling on how long a window message — a
                // resume, a display change — waits to be noticed.
                carried_preview_msg = match rx.recv_timeout(Duration::from_millis(IDLE_WAIT_MS)) {
                    Ok(message) => Some(message),
                    Err(RecvTimeoutError::Timeout) => None,
                    // Nothing left that could send, so this waits a tick rather
                    // than spinning on a channel no message can arrive on.
                    Err(RecvTimeoutError::Disconnected) => {
                        std::thread::sleep(Duration::from_millis(IDLE_WAIT_MS));
                        None
                    }
                };
            } else {
                std::thread::sleep(Duration::from_millis(16)); // ~60fps loop while something is on screen
            }
        }

        // Signal the dedicated loader worker to stop and wait for shutdown.
        if let Some(cancel) = pending_load_cancel.take() {
            cancel.store(true, Ordering::Release);
        }
        clear_load_request(&load_request_slot);
        let (_, cvar) = &*load_request_slot;
        cvar.notify_all();
        let _ = load_worker.join();
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::config::config::DEFAULT_FONT_SCALE_PERCENT;

    /// A display to place on: 1000 by 800 at its top-left corner.
    fn bounds() -> ScreenBounds {
        ScreenBounds {
            left: 0,
            top: 0,
            right: 1000,
            bottom: 800,
        }
    }

    /// A lock the tests that publish the pointer's own state take, so that one of them
    /// runs at a time: the item box and the hold regions are one set for the whole
    /// process, so two of these tests at once is one test's box answering another test's
    /// question.
    static POINTER_STAND_IN: Mutex<()> = Mutex::new(());

    /// A wait holds the pointer through the item it is waiting for rather than through the
    /// box its spinner occupies: the spinner is placed at the hand and follows it, so a box
    /// of its own is one the pointer can never leave — and the wait would be a preview no
    /// fast hand could close on its way to somewhere else (see `preview_pointer_hold`).
    #[test]
    fn holds_the_pointer_through_the_item_a_wait_is_for() {
        let _stand_in = POINTER_STAND_IN.lock().expect("the pointer's own state");

        // A spinner box at the pointer, as one on screen is, and the item the hover was
        // resolved from a row away from it.
        let spinner = (100, 100, 136, 136);
        *POINTER_HOLD_REGIONS.lock().expect("the published regions") = Some(vec![spinner]);
        publish_pointer_item_box((100, 200, 300, 220));

        WAITING_PREVIEW_HOLDING.store(true, Ordering::Release);
        assert!(
            !preview_pointer_hold(118, 118),
            "the spinner's own box is not a hold"
        );
        assert!(
            preview_pointer_hold(150, 210),
            "the item the wait is for holds the pointer"
        );
        assert!(
            !preview_pointer_hold(150, 230),
            "a pointer below that item has left the file it was waiting for"
        );

        // A wait whose item nobody could be read for is not a hold: a hold is the hook
        // leaving the mouse alone, so it is only ever taken on an answer, and the reading
        // that cannot happen is a preview nothing can close.
        clear_pointer_item_box();
        assert!(
            !preview_pointer_hold(150, 210),
            "a wait with no item box of its own holds nothing"
        );

        // A hold that is no longer noted holds nothing, whatever region was left behind.
        WAITING_PREVIEW_HOLDING.store(false, Ordering::Release);
        assert!(
            !preview_pointer_hold(118, 118),
            "nothing is holding the pointer"
        );

        clear_pointer_hold();
    }

    /// A reveal is held to the item the hover was resolved from: the pointer inside
    /// that item's box is a pointer still on the file, one outside it is a hover that
    /// has moved on, and a hover whose item was never read holds nothing back — which
    /// is what keeps a frame from going up for a file the hand has already left, and
    /// what keeps a keyboard preview or an unknown item from being held to a box that
    /// is not theirs (see `HOVER_POINTER_BOX`).
    #[test]
    fn holds_a_reveal_to_the_item_the_hover_was_resolved_from() {
        let _stand_in = POINTER_STAND_IN.lock().expect("the pointer's own state");

        publish_pointer_item_box((100, 200, 300, 220));

        assert!(pointer_item_holds(100, 200), "the item's own corner holds");
        assert!(pointer_item_holds(299, 219), "and so does its far one");
        assert!(
            !pointer_item_holds(99, 210),
            "a point past its left edge does not"
        );
        assert!(!pointer_item_holds(150, 220), "nor one a row below it");

        clear_pointer_item_box();
        assert!(
            pointer_item_holds(0, 0),
            "an item nothing was read for holds anything"
        );
    }

    fn layout(pos_x: i32, pos_y: i32, width: u32, height: u32) -> PreviewLayout {
        PreviewLayout {
            pos_x,
            pos_y,
            max_width: width,
            max_height: height,
            preview_w: width,
            preview_h: height,
        }
    }

    /// A keyboard placement of one item, at the size and mode the figures are easy to
    /// read in: the media's own size at 100%, and `Best Position`, so the place comes
    /// out of the room beside the item alone. `columns` is whether the item is read as
    /// a row of its view or as a box item — see `KeyboardPlacement`.
    fn keyboard_placement(
        item: (i32, i32, i32, i32),
        avoid: Option<ScreenRegion>,
        columns: bool,
    ) -> KeyboardPlacement {
        KeyboardPlacement {
            item_rect: item,
            avoid,
            columns,
            orig_dims: (400, 300),
            follow_cursor: false,
            preview_scale: PreviewScale::Percent(100),
        }
    }

    /// The display the placement figures are worked out for: 100%, which is what
    /// distances written in logical pixels are the same as the pixels of.
    const TEST_DPI: u32 = 96;

    /// A placement kept off `name`, at the size the media's own scale allows — the
    /// arrangement the figures are easy to read in. It is the step a hover makes, so
    /// the gap is the pointer's standoff at this display's scale, and `cursor` is where
    /// the pointer is for a placement that is a mouse hover's — which the ways out are
    /// held clear of, where it is given (see `avoiding_text`).
    fn placed(
        placement: PreviewLayout,
        media: (u32, u32),
        name: (i32, i32, i32, i32),
        cursor: Option<(i32, i32)>,
        bounds: ScreenBounds,
    ) -> PreviewLayout {
        avoiding_text(
            placement,
            media,
            PreviewScale::Percent(100),
            Clearance {
                text: Some(name),
                gap: logical_px(TEST_DPI, POINTER_STANDOFF_PIXELS),
                cursor,
            },
            bounds,
            TEST_DPI,
        )
    }

    /// Every scale a hover is laid out by, at the shares the app starts at. A test that
    /// is about one of them names that one and leaves the rest where the app has them.
    fn hover_scales() -> HoverScales {
        HoverScales {
            picture: PreviewScale::Percent(DEFAULT_PREVIEW_SCALE_PERCENT),
            video: PreviewScale::Percent(DEFAULT_VIDEO_SCALE_PERCENT),
            animated: PreviewScale::Percent(DEFAULT_ANIMATED_SCALE_PERCENT),
            ebook: DEFAULT_EBOOK_SCALE,
            document: DEFAULT_DOCUMENT_SCALE,
            font: DEFAULT_FONT_SCALE,
            design: DEFAULT_DESIGN_SCALE,
            vector: DEFAULT_VECTOR_SCALE,
        }
    }

    /// What a file's bytes say it is decides the share it is drawn at, the way they decide
    /// the loader that draws it and the box it is placed in: a picture left under a video's
    /// name is drawn at the picture's share, and one left under a document's name or a text
    /// name's at the picture's share too. The shares are given values of their own so that
    /// the answer says which of them was read, and the files are real ones because the
    /// question is asked of the file's own header.
    #[test]
    fn a_share_follows_the_content_rather_than_the_name() {
        let folder = std::env::temp_dir().join("rust-hover-preview-content-share");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let scales = HoverScales {
            picture: PreviewScale::Percent(100),
            video: PreviewScale::Percent(50),
            animated: PreviewScale::Percent(100),
            document: PreviewScale::FitToScreen,
            ebook: PreviewScale::FitToScreen,
            design: PreviewScale::Percent(25),
            vector: PreviewScale::Percent(25),
            font: PreviewScale::Percent(25),
        };

        for name in ["tomcat.mp4", "tomcat.docx", "tomcat.txt"] {
            let path = folder.join(name);
            write_test_png(&path, false);

            assert_eq!(
                effective_preview_scale(&path, scales),
                PreviewScale::Percent(100),
                "`{name}` holds a picture, so the picture's share is what it is drawn at"
            );
        }

        // And a name with nothing behind it keeps the share its name asks for, which is
        // every file this module's other tests are about.
        assert_eq!(
            effective_preview_scale(Path::new(r"C:\docs\report.pdf"), scales),
            PreviewScale::FitToScreen,
            "a page is still laid out at the page's share of the room"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// And the box it is placed in is the picture's too: a picture under a text name is
    /// measured as the picture it is rather than read as a page of text — which for bytes
    /// that are not text is no measurement at all, and a hover that never appears for a file
    /// that would otherwise be drawn.
    #[test]
    fn a_box_follows_the_content_rather_than_the_name() {
        let folder = std::env::temp_dir().join("rust-hover-preview-content-box");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let path = folder.join("tomcat.txt");
        write_test_png(&path, false);

        assert!(
            picture_dimensions(&path).is_some(),
            "the fixture is a picture the header reader can measure"
        );
        assert_eq!(
            media_dimensions(&path, bounds(), TEST_DPI),
            picture_dimensions(&path),
            "so the box is the picture's rather than the text measure's nothing"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// A page is painted into the box the layout planned for it, and an archive an engine listed is
    /// one of those pages rather than a size to fit the room it was given. It is the one question
    /// the share a kind is drawn at and the box the loader is handed both ask, and a kind left out
    /// of either is a listing stretched to the display: what an engine's own archive was, for as
    /// long as the two places listed their kinds by hand.
    #[test]
    fn a_page_is_painted_whether_this_app_read_the_archive_or_an_engine_listed_it() {
        let folder = std::env::temp_dir().join("rust-hover-preview-painted");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let path = |name: &str| folder.join(name);

        // A text file, an archive this app reads itself, and archives an engine lists — a cabinet,
        // a FreeArc archive and a stream its tool weighs: every one of them is a page.
        for name in [
            "notes.txt",
            "photos.zip",
            "backup.cab",
            "backup.arc",
            "readme.bz2",
        ] {
            let file = path(name);
            std::fs::write(&file, b"a file, of a sort").expect("a written file");

            assert!(
                page_is_painted(&file),
                "`{name}` is a page painted into the box it is given"
            );
        }

        // And a picture is not: it is a size of its own, drawn into whatever room it is given.
        let picture = path("tomcat.png");
        write_test_png(&picture, false);
        assert!(
            !page_is_painted(&picture),
            "a picture is a size of its own rather than a page"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// Every question about a file is asked of what its bytes say, which for these three is
    /// the difference between an engine being started and one being left alone: no page is
    /// asked of Office for a picture under a document's name, no wait for one is placed at
    /// the pointer, the player takes over a video under a picture's name, and text under a
    /// document's name is drawn as the text it is.
    #[test]
    fn a_foreign_engine_is_never_started_for_a_file_that_is_not_its_own() {
        if let Ok(mut config) = CONFIG.lock() {
            config.document_preview_enabled = true;
            config.video_preview_enabled = true;
        }

        let folder = std::env::temp_dir().join("rust-hover-preview-content-engines");
        std::fs::create_dir_all(&folder).expect("a test folder");

        // A picture under a document's name: Word is never asked for a page, and the hover
        // is not placed as the wait for one.
        let renamed = folder.join("report.docx");
        write_test_png(&renamed, false);

        assert!(
            office_formats::is_office_preview(&renamed),
            "the name is the document list's, which is what answered before this"
        );
        assert!(
            !office_render_is_due(&renamed, 800),
            "and the bytes are a picture's, so no page is asked of Office"
        );
        assert!(
            !page_is_on_the_way(&renamed),
            "nor is the hover placed at the pointer as the wait for one"
        );

        // A video under a picture's name: the probe has an answer to fetch, the wait for it
        // is shown, and the player takes the window over.
        let renamed = folder.join("clip.png");
        std::fs::write(&renamed, b"\x00\x00\x00\x20ftypisom").expect("a written video");

        assert!(
            !crate::formats::video_formats::is_video_file(&renamed),
            "the name is the picture list's, which is what answered before this"
        );
        assert!(
            drawn_as_video(&renamed),
            "and the bytes are a video's, so a video is what is drawn"
        );
        assert!(
            video_probe_due(&renamed),
            "whose shape is probed like any other video's"
        );

        // And text under a document's name, which is the box it is painted into and the
        // frame its lines wrap in.
        let renamed = folder.join("letter.docx");
        std::fs::write(&renamed, b"{\\rtf1\\ansi\\deff0 hello}").expect("a written document");

        let listed_as_text = {
            let config = CONFIG.lock().expect("the configuration");
            crate::formats::text_formats::matches_text_lists(
                &renamed,
                &config.text_extensions,
                &config.text_names,
            )
        };

        assert!(!listed_as_text, "the name is not one the text lists carry");
        assert!(
            is_text_preview(&renamed),
            "and the bytes are text, so text is what draws it"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// An Office document is asked of one engine and not the other: the application that owns
    /// the format draws the page where it is installed, and the render engine beside it draws
    /// one where it is not.
    ///
    /// Both drawing it is what a **smaller version** of a preview in front of the right one
    /// was: the engine's page is a page of the document's own size, and the hover it arrived
    /// in had been laid out as the wait for a page — the spinner's box at the pointer — so it
    /// was drawn at the spinner's share of the display and shown there for the second or two
    /// the application took to draw the real page over it. Neither drawing it is a hover
    /// waiting for a page nothing was asked to draw.
    ///
    /// Which application is installed is the machine's answer rather than this app's, so the
    /// expectation is written from it; the rule is this app's, and it is what is asserted.
    #[test]
    fn an_office_document_is_asked_of_one_engine_and_not_the_other() {
        if let Ok(mut config) = CONFIG.lock() {
            config.document_preview_enabled = true;
            // What the app's own `config.ini` holds is not what this test is about: it asks
            // the machine, and the setting is pinned to the one that asks the machine.
            config.office_engine = OfficeEngine::MicrosoftOffice;
            config.office_extensions = office_formats::sanitize_office_extensions(
                office_formats::DEFAULT_OFFICE_EXTENSIONS,
            );
        }

        let folder = std::env::temp_dir().join("rust-hover-preview-office-engines");
        std::fs::create_dir_all(&folder).expect("a test folder");

        // A document its own application draws, on a machine that has one: the tier is asked
        // for the page, and the render engine is not — it would be a second rendering of the
        // same document, at a size and a place the layout had never measured a page for.
        let named = folder.join("report.docx");
        std::fs::write(&named, b"PK\x03\x04\x00\x00\x00\x00").expect("a written document");
        let installed = office_formats::app_installed(&named);

        assert!(
            office_formats::is_office_preview(&named),
            "the name is the document list's, which is what a document is known by: every \
             format of Office's is a container, and a container says nothing about itself"
        );
        assert_eq!(
            office_render_is_due(&named, 800),
            installed,
            "the application that owns the format is asked for a page exactly where it is here"
        );
        assert_eq!(
            libre_formats::engine_page_kind(&named).is_some(),
            !installed,
            "and the render engine beside it is what draws one exactly where it is not, so a \
             document is never drawn twice and never left undrawn"
        );

        // And a name no family claims is nobody's: an application that could not be resolved
        // from it is not asked for a page, and the render engine draws a document of its own
        // kinds rather than one whose name it was never shown (see `engine_page_kind`).
        let unclaimed = folder.join("report.bin");
        std::fs::write(&unclaimed, b"PK\x03\x04\x00\x00\x00\x00").expect("a written document");

        assert!(
            !office_formats::app_installed(&unclaimed),
            "no family answers for a name like this one"
        );
        assert_eq!(
            libre_formats::engine_page_kind(&unclaimed),
            None,
            "and the render engine is not asked about it either"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// A little GIF of one pixel, written with `frames` frame blocks in it: an
    /// animation when there is more than one frame and a still picture when there is
    /// one, which is the whole of what tells the two apart. The bytes are a real file
    /// rather than a shape, because that is what the probe and the decoder both read —
    /// two colours, one pixel, and the frames of it.
    fn write_test_gif(path: &Path, frames: usize) {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(b"GIF89a");
        bytes.extend_from_slice(&[1, 0, 1, 0]); // one pixel square
        bytes.push(0xF0); // a global colour table, two entries
        bytes.push(0); // background colour index
        bytes.push(0); // pixel aspect ratio
        bytes.extend_from_slice(&[0, 0, 0, 255, 255, 255]); // black and white

        for _ in 0..frames {
            bytes.push(0x2C); // an image descriptor
            bytes.extend_from_slice(&[0, 0, 0, 0, 1, 0, 1, 0, 0]); // at 0,0, one pixel
            bytes.push(2); // the LZW code size
            bytes.extend_from_slice(&[2, 0x4C, 0x01, 0]); // one block of one code, then the end
        }

        bytes.push(0x3B); // the trailer
        std::fs::write(path, bytes).expect("a written GIF");
    }

    /// A little PNG, with an `acTL` chunk ahead of its image data when `animated`: the
    /// chunk is the whole of what the probe reads, and the pixel chunks after it are a
    /// real still frame so that the file is a PNG either way.
    fn write_test_png(path: &Path, animated: bool) {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&[0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A]);

        // The header: one pixel, eight bits, colour type six.
        push_png_chunk(
            &mut bytes,
            *b"IHDR",
            &[0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0],
        );

        if animated {
            push_png_chunk(&mut bytes, *b"acTL", &[0, 0, 0, 2, 0, 0, 0, 0]);
            push_png_chunk(
                &mut bytes,
                *b"fcTL",
                &[
                    0, 0, 0, 0, // the first frame's sequence number
                    0, 0, 0, 1, 0, 0, 0, 1, // its width and height
                    0, 0, 0, 0, 0, 0, // where it sits
                    0, 0, 0, 1, 0, 0, 0, 1, // the frame's own delay
                    0, 0, // how it replaces what it is drawn over
                ],
            );
        }

        push_png_chunk(
            &mut bytes,
            *b"IDAT",
            &[0x78, 0x01, 0x63, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01],
        );
        push_png_chunk(&mut bytes, *b"IEND", &[]);

        std::fs::write(path, bytes).expect("a written PNG");
    }

    /// One PNG chunk: its body's length, its type, its body, and the CRC of the type and
    /// the body together, which is the whole of the container.
    fn push_png_chunk(bytes: &mut Vec<u8>, kind: [u8; 4], body: &[u8]) {
        bytes.extend_from_slice(&(body.len() as u32).to_be_bytes());
        bytes.extend_from_slice(&kind);
        bytes.extend_from_slice(body);
        bytes.extend_from_slice(&crc32(&[&kind, body].concat()).to_be_bytes());
    }

    /// The CRC every PNG chunk is checked with: a table-less, bit-at-a-time take on the
    /// standard polynomial, which is all a test fixture needs.
    fn crc32(bytes: &[u8]) -> u32 {
        let mut crc = 0xFFFF_FFFFu32;
        for byte in bytes {
            crc ^= *byte as u32;
            for _ in 0..8 {
                crc = if crc & 1 != 0 {
                    (crc >> 1) ^ 0xEDB8_8320
                } else {
                    crc >> 1
                };
            }
        }
        !crc
    }

    /// A keyboard preview of a row is placed in the room past the region the `Avoid`
    /// setting keeps it off, so a row is cleared only as far as the setting asks: past
    /// every column at `Avoid Details`, and only past the name at `Avoid Filename`,
    /// where the columns drawn after the name are the room the preview takes.
    #[test]
    fn a_keyboard_rows_tail_begins_at_the_region_it_is_kept_off() {
        let row = (0, 100, 1000, 140);

        let past_the_name = compute_keyboard_layout(
            keyboard_placement(row, Some((20, 104, 120, 136)), true),
            bounds(),
            TEST_DPI,
        )
        .expect("a placement past the name");

        let past_every_column = compute_keyboard_layout(
            keyboard_placement(row, Some((20, 104, 900, 136)), true),
            bounds(),
            TEST_DPI,
        )
        .expect("a placement past the row's columns");

        assert_eq!(past_the_name.pos_x, 130, "just past the name");
        assert_eq!(past_every_column.pos_x, 910, "just past the columns");
        assert!(
            past_the_name.pos_x < past_every_column.pos_x,
            "a narrower region leaves more of the row to be covered"
        );
    }

    /// A row is a row whatever the window is doing: a `Details` row of a window
    /// narrower than half the display still takes its placement from the region the
    /// `Avoid` setting keeps it off — past the name, past the `Name` column, past the
    /// row's columns — rather than from its own right edge, which is where the same
    /// row read as a box would put its preview at every one of those settings. The row
    /// below is a sixth of the display across, its text written into the left end of
    /// it, as `Details` draws one.
    #[test]
    fn a_rows_placement_is_measured_from_the_region_whatever_the_window_is() {
        let row = (0, 100, 420, 124);
        let gap = logical_px(TEST_DPI, KEYBOARD_GAP_PIXELS);

        // The region each way of avoiding keeps off, and the edge it leaves the
        // preview past: the name's own width, the `Name` column, and the columns.
        let regions = [
            ((12, 103, 180, 122), 180, "past the name"),
            ((12, 103, 220, 122), 220, "past the `Name` column"),
            ((12, 103, 418, 122), 418, "past the row's columns"),
        ];

        for (region, region_right, what) in regions {
            let layout = compute_keyboard_layout(
                keyboard_placement(row, Some(region), true),
                bounds(),
                TEST_DPI,
            )
            .unwrap_or_else(|| panic!("a placement {what}"));

            assert_eq!(layout.pos_x, region_right + gap, "{what}");
        }

        // Which leaves the name's placement inside the row it belongs to, where the
        // row's own right edge would have put it had the row been read as a box.
        let past_the_name = compute_keyboard_layout(
            keyboard_placement(row, Some((12, 103, 180, 122)), true),
            bounds(),
            TEST_DPI,
        )
        .expect("a placement past the name");

        assert!(
            past_the_name.pos_x < row.2,
            "the name's tail is inside the row: {} against the row's own edge {}",
            past_the_name.pos_x,
            row.2
        );
    }

    /// An item that draws nothing beside its name — the label under an icon — is a box
    /// item: its preview is placed beside the box itself whichever region is kept off,
    /// since the label is a piece drawn inside that box and not a row of the view.
    #[test]
    fn a_box_items_tail_is_its_own_edge_and_not_its_labels() {
        let tile = (0, 100, 120, 220);
        let label = (25, 190, 95, 206);
        let gap = logical_px(TEST_DPI, KEYBOARD_GAP_PIXELS);

        for avoid in [Some(label), None] {
            let layout =
                compute_keyboard_layout(keyboard_placement(tile, avoid, false), bounds(), TEST_DPI)
                    .expect("a placement beside the tile");

            assert_eq!(layout.pos_x, tile.2 + gap, "just past the tile's own edge");
        }
    }

    /// A row the placement was given no region for is placed by the position mode
    /// alone: it is anchored at its middle, the way a hover over it is read, and the
    /// preview is allowed to cover it. No keyboard preview is asked this way — the hook
    /// reads `Avoid Nothing` as `Avoid Filename` (see `KeyboardPlacement`) — and the
    /// answer is what a caller with no region to be placed from gets.
    #[test]
    fn a_keyboard_row_with_nothing_kept_off_is_placed_by_position_alone() {
        let layout = compute_keyboard_layout(
            keyboard_placement((0, 100, 1000, 140), None, true),
            bounds(),
            TEST_DPI,
        )
        .expect("a placement");

        assert_eq!(layout.pos_x, 510, "half the row's width, and the gap");
    }

    /// A page — a PDF's, or one a document was drawn as — is drawn at the room the display
    /// has, because the room is free quality there. Every setting at or above
    /// `100%` asks for at least that room, so they are one setting for a page; only
    /// a setting below it is a size the user picked, and it is answered by
    /// reducing the fitted size rather than by ignoring it. Which page setting is
    /// read is the kind's own: a PDF page follows `ebook_scale` and a page drawn for
    /// a document follows `document_scale`, and neither moves for the picture scale.
    #[test]
    fn a_page_takes_the_room_the_display_has() {
        let pdf = PathBuf::from(r"C:\docs\report.pdf");
        let vector_scale = DEFAULT_VECTOR_SCALE;
        let document_scale = PreviewScale::Percent(75);

        for configured in [
            PreviewScale::FitToScreen,
            PreviewScale::Percent(100),
            PreviewScale::Percent(200),
            PreviewScale::Percent(400),
        ] {
            assert_eq!(
                effective_preview_scale(
                    &pdf,
                    HoverScales {
                        picture: configured,
                        ebook: configured,
                        vector: vector_scale,
                        document: document_scale,
                        ..hover_scales()
                    }
                ),
                PreviewScale::FitToScreen
            );
        }

        assert_eq!(
            effective_preview_scale(
                &pdf,
                HoverScales {
                    picture: PreviewScale::Percent(400),
                    ebook: PreviewScale::Percent(50),
                    vector: vector_scale,
                    document: document_scale,
                    ..hover_scales()
                }
            ),
            PreviewScale::FitToScreenReduced(50)
        );
        assert_eq!(
            effective_preview_scale(
                &pdf,
                HoverScales {
                    picture: PreviewScale::Percent(400),
                    ebook: PreviewScale::Percent(25),
                    vector: vector_scale,
                    document: document_scale,
                    ..hover_scales()
                }
            ),
            PreviewScale::FitToScreenReduced(25),
            "a page follows its own share of the display, whatever the picture scale says"
        );
        assert_eq!(
            effective_preview_scale(
                &pdf,
                HoverScales {
                    picture: PreviewScale::Percent(400),
                    ebook: PreviewScale::FitToScreen,
                    vector: vector_scale,
                    document: document_scale,
                    ..hover_scales()
                }
            ),
            PreviewScale::FitToScreen
        );
    }

    /// A document's scale means the room the display has rather than the size the file
    /// asks for, so `Fit to Screen` is the whole of that room and a percentage is a
    /// share of it — half the display at `50%`, a tenth at `10%` — whatever scale the
    /// pictures beside it are drawn at. A share the configuration can hold but the menu
    /// does not offer is honored rather than rounded to a menu entry.
    #[test]
    fn a_document_is_drawn_at_its_share_of_the_room() {
        let svg = PathBuf::from(r"C:\art\clock.svg");
        let configured = PreviewScale::Percent(100);
        let ebook = PreviewScale::FitToScreen;

        assert_eq!(
            effective_preview_scale(
                &svg,
                HoverScales {
                    picture: configured,
                    vector: PreviewScale::FitToScreen,
                    ebook,
                    ..hover_scales()
                }
            ),
            PreviewScale::FitToScreen
        );
        for percent in [75, 50, 25, 10] {
            assert_eq!(
                effective_preview_scale(
                    &svg,
                    HoverScales {
                        picture: configured,
                        vector: PreviewScale::Percent(percent),
                        ebook,
                        ..hover_scales()
                    }
                ),
                PreviewScale::FitToScreenReduced(percent),
                "{percent}% of the room"
            );
        }

        assert_eq!(
            effective_preview_scale(
                &svg,
                HoverScales {
                    picture: configured,
                    vector: PreviewScale::Percent(60),
                    ebook,
                    ..hover_scales()
                }
            ),
            PreviewScale::FitToScreenReduced(60),
            "a share the menu does not offer is the share it is"
        );

        for configured in [
            PreviewScale::FitToScreen,
            PreviewScale::Percent(100),
            PreviewScale::Percent(400),
        ] {
            assert_eq!(
                effective_preview_scale(
                    &svg,
                    HoverScales {
                        picture: configured,
                        vector: PreviewScale::Percent(100),
                        ebook,
                        ..hover_scales()
                    }
                ),
                PreviewScale::FitToScreen,
                "asking for the whole room or more is the whole room"
            );
        }
    }

    /// A font is drawn at a share of the room the way a document is, and it is the fourth
    /// setting of its own: the specimen is a page of this app's making, so the share decides
    /// how large the type is drawn — and what a PDF, a document and a page are configured at
    /// leaves it where it is.
    #[test]
    fn a_specimen_is_drawn_at_its_share_of_the_room() {
        let font = PathBuf::from(r"C:\fonts\Inter-Regular.woff2");
        let picture = PreviewScale::Percent(400);
        let ebook = PreviewScale::FitToScreen;
        let vector = PreviewScale::Percent(75);

        assert_eq!(
            effective_preview_scale(
                &font,
                HoverScales {
                    picture,
                    vector,
                    ebook,
                    ..hover_scales()
                }
            ),
            PreviewScale::FitToScreenReduced(DEFAULT_FONT_SCALE_PERCENT),
            "a specimen starts at half the room, whatever the other kinds are set to"
        );

        for percent in [75, 50, 25, 10] {
            assert_eq!(
                effective_preview_scale(
                    &font,
                    HoverScales {
                        picture,
                        vector,
                        ebook,
                        font: PreviewScale::Percent(percent),
                        ..hover_scales()
                    }
                ),
                PreviewScale::FitToScreenReduced(percent),
                "{percent}% of the room"
            );
        }

        for configured in [
            PreviewScale::FitToScreen,
            PreviewScale::Percent(100),
            PreviewScale::Percent(400),
        ] {
            assert_eq!(
                effective_preview_scale(
                    &font,
                    HoverScales {
                        picture,
                        vector,
                        ebook,
                        font: configured,
                        ..hover_scales()
                    }
                ),
                PreviewScale::FitToScreen,
                "asking for the whole room or more is the whole room"
            );
        }

        // A file that is not a font keeps the picture scale, whatever the specimen's is.
        let png = PathBuf::from(r"C:\art\photo.png");
        assert_eq!(
            effective_preview_scale(
                &png,
                HoverScales {
                    picture: PreviewScale::Percent(200),
                    vector,
                    ebook,
                    font: PreviewScale::Percent(10),
                    ..hover_scales()
                }
            ),
            PreviewScale::Percent(200)
        );
    }

    /// The picture scale is about another kind of preview, so a document's size does
    /// not move when it does: the two settings are read one each rather than one for
    /// both.
    #[test]
    fn a_documents_size_is_not_the_pictures_setting() {
        let svg = PathBuf::from(r"C:\art\clock.svg");
        let picture = PreviewScale::Percent(100);

        assert_eq!(
            effective_preview_scale(
                &svg,
                HoverScales {
                    picture,
                    vector: DEFAULT_VECTOR_SCALE,
                    ebook: PreviewScale::Percent(25),
                    document: PreviewScale::Percent(75),
                    font: PreviewScale::Percent(10),
                    ..hover_scales()
                }
            ),
            PreviewScale::FitToScreen,
            "a document follows its own share, not a page's"
        );

        // And a picture is still drawn at the picture scale, whatever the document
        // settings say.
        let png = PathBuf::from(r"C:\art\photo.png");
        assert_eq!(
            effective_preview_scale(
                &png,
                HoverScales {
                    picture: PreviewScale::Percent(200),
                    vector: PreviewScale::FitToScreen,
                    ebook: PreviewScale::Percent(25),
                    document: PreviewScale::Percent(75),
                    font: PreviewScale::Percent(10),
                    ..hover_scales()
                }
            ),
            PreviewScale::Percent(200)
        );
    }

    /// A video's size is a setting of its own: a video the probe has answered for keeps
    /// the share `video_scale` names whatever the pictures beside it are drawn at. A
    /// video the probe has not answered for is not laid out at any share yet — what is
    /// on screen is the wait for the probe — so both settings leave that hover in the
    /// spinner's own box (see `video_probe_due`).
    #[test]
    fn a_video_follows_its_own_scale() {
        let video = PathBuf::from(r"C:\clips\holiday.mp4");
        let unprobed = PathBuf::from(r"C:\clips\not-probed-yet.mp4");

        assert_eq!(
            effective_preview_scale(
                &unprobed,
                HoverScales {
                    video: PreviewScale::Percent(50),
                    picture: PreviewScale::Percent(400),
                    ..hover_scales()
                }
            ),
            PreviewScale::Percent(100),
            "a video the probe has not answered for waits in the spinner's own box"
        );

        // The probe's answer is held per file and version, which is the state a hover
        // the probe has already answered for is in by the time its replay is laid out.
        video_geometry_cache().insert(
            VideoGeometryKey {
                path: video.clone(),
                version: file_version(&video),
            },
            ProbedGeometry::Measured(VideoGeometry {
                width: 1920,
                height: 1080,
                crop: None,
            }),
        );

        assert_eq!(
            effective_preview_scale(
                &video,
                HoverScales {
                    video: PreviewScale::Percent(50),
                    picture: PreviewScale::Percent(400),
                    ..hover_scales()
                }
            ),
            PreviewScale::Percent(50),
            "a measured video follows its own share, not the picture's"
        );

        assert_eq!(
            effective_preview_scale(
                &video,
                HoverScales {
                    video: PreviewScale::FitToScreen,
                    ..hover_scales()
                }
            ),
            PreviewScale::FitToScreen,
            "and taking the whole of its own size is a share it is offered too"
        );
    }

    /// A GIF that moves, a WebP that moves and a PNG that moves are one kind of preview —
    /// the kind whose size `animated_scale` names — while a GIF or a PNG that holds a
    /// single frame is a picture like any other and keeps the picture scale. What tells
    /// them apart is the file's own content, not its name: an animated PNG is very often
    /// called `.png`, and a `.gif` written by a still encoder is a picture.
    #[test]
    fn an_animation_follows_its_own_scale_and_a_still_keeps_the_pictures() {
        let folder = std::env::temp_dir().join("rhp-animated-scale-fixtures");
        std::fs::create_dir_all(&folder).expect("a fixture folder");

        let animated_gif = folder.join("animated.gif");
        write_test_gif(&animated_gif, 2);
        let still_gif = folder.join("still.gif");
        write_test_gif(&still_gif, 1);
        let animated_apng = folder.join("animated.png");
        write_test_png(&animated_apng, true);
        let still_png = folder.join("still.png");
        write_test_png(&still_png, false);

        let scales = HoverScales {
            picture: PreviewScale::Percent(100),
            animated: PreviewScale::Percent(25),
            ..hover_scales()
        };

        assert_eq!(
            effective_preview_scale(&animated_gif, scales),
            PreviewScale::Percent(25),
            "a GIF with two frames is an animation"
        );
        assert_eq!(
            effective_preview_scale(&still_gif, scales),
            PreviewScale::Percent(100),
            "a GIF with one frame is a picture"
        );
        assert_eq!(
            effective_preview_scale(&animated_apng, scales),
            PreviewScale::Percent(25),
            "a PNG whose chunks say it animates is an animation"
        );
        assert_eq!(
            effective_preview_scale(&still_png, scales),
            PreviewScale::Percent(100),
            "a PNG without the animation control chunk is a picture"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// A WebP's animation is a chunk of its container, and the probe reads it there: an
    /// `ANIM` chunk ahead of the picture data is an animation, while a plain `VP8 ` or
    /// `VP8L` WebP is a picture whatever else follows it. Nothing checksummed is written
    /// here — the probe walks chunk lengths, and the frames those chunks hold are the
    /// loader's business rather than this question's.
    #[test]
    fn a_webp_is_an_animation_when_its_container_says_so() {
        let folder = std::env::temp_dir().join("rhp-webp-probe-fixtures");
        std::fs::create_dir_all(&folder).expect("a fixture folder");

        let chunk = |kind: &[u8; 4], body: &[u8]| {
            let mut bytes = kind.to_vec();
            bytes.extend_from_slice(&(body.len() as u32).to_le_bytes());
            bytes.extend_from_slice(body);
            if body.len() % 2 == 1 {
                bytes.push(0); // chunks are padded to an even length
            }
            bytes
        };
        let riff = |chunks: Vec<u8>| {
            let mut bytes = b"RIFF".to_vec();
            bytes.extend_from_slice(&((chunks.len() + 4) as u32).to_le_bytes());
            bytes.extend_from_slice(b"WEBP");
            bytes.extend_from_slice(&chunks);
            bytes
        };

        let mut animated_chunks = chunk(b"VP8X", &[0x02, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
        animated_chunks.extend_from_slice(&chunk(b"ANIM", &[0; 6]));
        animated_chunks.extend_from_slice(&chunk(b"ANMF", &[0; 16]));
        let animated_path = folder.join("animated.webp");
        std::fs::write(&animated_path, riff(animated_chunks)).expect("a written WebP");

        let still_path = folder.join("still.webp");
        std::fs::write(&still_path, riff(chunk(b"VP8 ", &[0; 4]))).expect("a written WebP");

        assert!(image_is_animated(&animated_path), "an `ANIM` chunk");
        assert!(!image_is_animated(&still_path), "a plain picture");
        assert!(
            !image_is_animated(&folder.join("missing.webp")),
            "a file that is not there is not an animation"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }

    #[test]
    fn only_a_gif_with_a_second_frame_counts_as_one() {
        let folder = std::env::temp_dir().join("rhp-animated-probe-fixtures");
        std::fs::create_dir_all(&folder).expect("a fixture folder");

        let gif = folder.join("animation.gif");
        write_test_gif(&gif, 2);
        let still = folder.join("still.gif");
        write_test_gif(&still, 1);

        assert!(image_is_animated(&gif), "two frame blocks in the file");
        assert!(!image_is_animated(&still), "one frame block in the file");

        // The same file the loader sees, so the size a hover is placed at is the size its
        // frames are decoded at: two frames are what makes it an animation there as well.
        let loaded = load_animated_gif(
            &gif,
            64,
            64,
            PreviewScale::Percent(100),
            Arc::new(AtomicBool::new(false)),
        )
        .expect("the two-frame file loads as an animation");
        assert_eq!(
            loaded.media_type.kind(),
            Some(PreviewType::Images),
            "it is a preview of the `Images` kind"
        );

        assert!(
            load_animated_gif(
                &still,
                64,
                64,
                PreviewScale::Percent(100),
                Arc::new(AtomicBool::new(false)),
            )
            .is_none(),
            "a single frame is not an animation, so the loader turns it down"
        );

        let _ = std::fs::remove_dir_all(&folder);
    }

    /// A share of the fitted size is not a share of the media's own size, and the
    /// difference is the point: `50%` is half the media — the same size whatever
    /// room it has — while a reduced fit is half of what the room would have
    /// allowed, so it grows with the room.
    #[test]
    fn reduces_the_fitted_size_by_the_configured_share() {
        // 800 by 600 in a 400 by 300 room: the fit is half the media's size, so
        // half of the fit is a quarter of it.
        assert_eq!(
            scale_dimensions(800, 600, 400, 300, PreviewScale::FitToScreen),
            (400, 300)
        );
        assert_eq!(
            scale_dimensions(800, 600, 400, 300, PreviewScale::FitToScreenReduced(50)),
            (200, 150)
        );
        assert_eq!(
            scale_dimensions(800, 600, 400, 300, PreviewScale::FitToScreenReduced(25)),
            (100, 75)
        );

        // The same media in a room that allows all of it: 50% of the media is 400
        // by 300, while half of the fitted 800 by 600 is not.
        assert_eq!(
            scale_dimensions(800, 600, 1900, 1000, PreviewScale::Percent(50)),
            (400, 300)
        );
        assert_eq!(
            scale_dimensions(800, 600, 1900, 1000, PreviewScale::FitToScreenReduced(50)),
            (667, 500)
        );
    }

    /// A document whose page has not been rendered yet is a spinner, and the
    /// spinner is placed at its own size: fitted to the display it would be a
    /// screen-sized square with a spinner drawn in the middle of it.
    #[test]
    fn a_document_with_no_page_yet_waits_in_the_spinners_own_box() {
        let waiting = PathBuf::from(r"C:\docs\not-rendered-yet.docx");
        let scale = effective_preview_scale(
            &waiting,
            HoverScales {
                picture: PreviewScale::Percent(400),
                vector: DEFAULT_VECTOR_SCALE,
                ebook: PreviewScale::Percent(25),
                document: PreviewScale::Percent(10),
                font: PreviewScale::Percent(DEFAULT_FONT_SCALE_PERCENT),
                ..hover_scales()
            },
        );

        assert_eq!(scale, PreviewScale::Percent(100));
        assert_eq!(
            scale_dimensions(
                office_preview::WAITING_BOX,
                office_preview::WAITING_BOX,
                1920,
                1080,
                scale,
            ),
            (office_preview::WAITING_BOX, office_preview::WAITING_BOX)
        );
    }

    /// A page's share of the display is read for a bitmap — a workbook's corner where
    /// no page can be exported — as the same share of its own size, and the whole of
    /// the display as the bitmap at the size it is: a quarter of the room is a quarter
    /// of the picture, and a fit never stretches one to fill the screen.
    #[test]
    fn a_bitmap_follows_a_pages_share_without_being_enlarged() {
        for percent in [75, 50, 25, 10] {
            assert_eq!(
                bitmap_at_display_scale(PreviewScale::Percent(percent)),
                PreviewScale::Percent(percent),
                "{percent}% of the picture"
            );
        }

        assert_eq!(
            bitmap_at_display_scale(PreviewScale::FitToScreen),
            PreviewScale::Percent(100),
            "the whole room is the picture at its own size"
        );
        assert_eq!(
            bitmap_at_display_scale(PreviewScale::FitToScreenReduced(50)),
            PreviewScale::Percent(100)
        );
    }

    /// The spinner is the arc alone: the box it is drawn in is transparent, so a
    /// hover that is waiting shows a spinner rather than a square of its own.
    #[test]
    fn draws_the_spinner_without_a_box_around_it() {
        let frame = render_loading_frame(64, 64, 0.0);
        assert_eq!(frame.len(), 64 * 64 * 4);

        // The corners — and so the box the arc is drawn in — are nothing at all.
        for corner in [0usize, 63, 64 * 63, 64 * 64 - 1] {
            assert_eq!(&frame[corner * 4..corner * 4 + 4], &[0, 0, 0, 0]);
        }

        // The arc is drawn, and the halo under it and the arc's own tail are
        // partly there rather than filled in.
        let alphas: Vec<u8> = frame
            .as_chunks::<4>()
            .0
            .iter()
            .map(|pixel| pixel[3])
            .collect();
        assert!(alphas.iter().any(|&alpha| alpha > 200), "the arc is drawn");
        assert!(
            alphas.iter().any(|&alpha| (1..=200).contains(&alpha)),
            "the halo and the tail fade rather than fill"
        );
    }

    /// The waiting frame is the size of the spinner in it: over a turn of the
    /// spinner, the arc and its halo reach each of the box's edges, so a spinner
    /// placed at the pointer's corner is the spinner at the pointer rather than an
    /// empty frame around one.
    #[test]
    fn the_waiting_frame_is_the_size_of_the_spinner_in_it() {
        let side = office_preview::WAITING_BOX as usize;
        let (mut min_x, mut min_y) = (side, side);
        let (mut max_x, mut max_y) = (0, 0);

        // The arc ends in a tail that fades to nothing, so one frame has fewer of
        // the box's edges inked than the next: the frame is measured against the
        // whole turn rather than against the arc's position in it.
        for quarter in 0..4 {
            let angle = quarter as f32 * std::f32::consts::FRAC_PI_2;
            let frame = render_loading_frame(side as u32, side as u32, angle);
            for (index, pixel) in frame.as_chunks::<4>().0.iter().enumerate() {
                if pixel[3] == 0 {
                    continue;
                }
                let (x, y) = (index % side, index / side);
                min_x = min_x.min(x);
                min_y = min_y.min(y);
                max_x = max_x.max(x);
                max_y = max_y.max(y);
            }
        }

        // Nothing but the fade the halo ends in is left between the spinner and
        // the frame it is placed in.
        assert!(
            min_x <= 2 && min_y <= 2,
            "the spinner reaches its frame's near edges, {min_x} and {min_y} of {side}"
        );
        assert!(
            max_x >= side - 3 && max_y >= side - 3,
            "and its far ones, {max_x} and {max_y} of {side}"
        );
    }

    /// A load that may be about to finish is given the delay `spinner_delay_ms` names
    /// before the spinner goes up — the same one for every kind of wait — while a delay
    /// of nothing is a spinner that goes up with the load.
    #[test]
    fn puts_the_spinner_up_once_the_load_has_run_for_the_delay() {
        let load = |age: Duration, delay: Duration, upgrade: bool| PendingLoad {
            generation: 1,
            hide_epoch: 0,
            path: PathBuf::new(),
            started: Instant::now() - age,
            pos_x: 0,
            pos_y: 0,
            width: 64,
            height: 64,
            room: (1920, 1040),
            spinner_shown: false,
            spinner_delay: delay,
            spinner_pos: (0, 0),
            spinner_side: office_preview::WAITING_BOX,
            placement: None,
            upgrade,
            awaiting_engine: false,
        };
        let default_delay = Duration::from_millis(DEFAULT_SPINNER_DELAY_MS);

        // A load that may be about to finish: not yet, and due once it has run for
        // the delay.
        assert!(!load(Duration::from_millis(20), default_delay, false).spinner_due());
        assert!(load(default_delay, default_delay, false).spinner_due());

        // A delay of nothing puts the spinner up with the load, and a load given a
        // longer one than it has run for is not due yet.
        assert!(load(Duration::ZERO, Duration::ZERO, false).spinner_due());
        assert!(!load(Duration::from_secs(1), Duration::from_secs(5), false).spinner_due());

        // An upgrade never: what is on screen stays until the page replaces it.
        assert!(!load(Duration::from_secs(10), default_delay, true).spinner_due());

        // And a spinner that is already up is not put up a second time.
        let mut showing = load(Duration::from_secs(10), default_delay, false);
        showing.spinner_shown = true;
        assert!(!showing.spinner_due());
    }

    /// An engine answers a file once: the answer to a request that was folded into work
    /// already in flight names the hover before the one that is waiting, and it is still
    /// that wait's own answer. Read as an answer about a hover that has gone, it would
    /// leave the spinner standing over work the engine has already finished.
    #[test]
    fn an_engine_answer_names_the_hover_before_the_one_waiting_on_it() {
        let page = PathBuf::from(r"C:\docs\notes.docx");
        let other = PathBuf::from(r"C:\docs\other.docx");

        let load = |path: PathBuf, awaiting_engine: bool| PendingLoad {
            generation: 2,
            hide_epoch: 0,
            path,
            started: Instant::now(),
            pos_x: 0,
            pos_y: 0,
            width: 64,
            height: 64,
            room: (1920, 1040),
            spinner_shown: true,
            spinner_delay: Duration::from_millis(DEFAULT_SPINNER_DELAY_MS),
            spinner_pos: (0, 0),
            spinner_side: office_preview::WAITING_BOX,
            placement: None,
            upgrade: false,
            awaiting_engine,
        };

        let waiting = load(page.clone(), true);

        // The answer to the hover before this one, for the same file, is this wait's own.
        assert!(answer_belongs_to_the_wait(
            &page,
            true,
            Some(page.as_path()),
            Some(&waiting)
        ));

        // An engine that drew nothing has nothing to replay — the load that reads the
        // file again is what takes the spinner down, so this is not an answer to hand on.
        assert!(!answer_belongs_to_the_wait(
            &page,
            false,
            Some(page.as_path()),
            Some(&waiting)
        ));

        // Another file's page is not this wait's, and a wait that is not on an engine at
        // all — a decode, a probe — is not one an engine's answer belongs to.
        assert!(!answer_belongs_to_the_wait(
            &other,
            true,
            Some(page.as_path()),
            Some(&waiting)
        ));
        assert!(!answer_belongs_to_the_wait(
            &page,
            true,
            Some(page.as_path()),
            Some(&load(page.clone(), false))
        ));

        // And nothing waiting is nothing to answer: a page that lands after the pointer
        // has left belongs to the hover that has gone.
        assert!(!answer_belongs_to_the_wait(
            &page,
            true,
            Some(page.as_path()),
            None
        ));
        assert!(!answer_belongs_to_the_wait(
            &page,
            true,
            None,
            Some(&waiting)
        ));
    }

    /// A preview that is still on its way follows the pointer: it is placed again
    /// for a cursor that has moved along the item, kept where it is when the
    /// cursor has not moved, and left alone when the hover it came from was the
    /// keyboard's rather than the pointer's.
    ///
    /// Two places follow the cursor, and they are not the same one: the preview's,
    /// which is where the media lands, and the spinner's own, which is the arc at the
    /// pointer's corner whatever box the preview will arrive in.
    #[test]
    fn a_pending_preview_follows_the_pointer() {
        let pending = |placement: Option<HoverPlacement>| PendingLoad {
            generation: 1,
            hide_epoch: 0,
            path: PathBuf::new(),
            started: Instant::now(),
            pos_x: 0,
            pos_y: 0,
            width: 0,
            height: 0,
            room: (1920, 1040),
            spinner_shown: true,
            spinner_delay: Duration::from_millis(DEFAULT_SPINNER_DELAY_MS),
            spinner_pos: (0, 0),
            spinner_side: 0,
            placement,
            upgrade: false,
            awaiting_engine: false,
        };
        let placement = HoverPlacement {
            orig_dims: (800, 600),
            avoid: None,
            follow_cursor: false,
            preview_scale: PreviewScale::FitToScreen,
            at_the_pointer_corner: false,
        };

        let mut pl = pending(Some(placement));
        let followed = pl.follow_pointer(POINT { x: 300, y: 300 }, TEST_DPI);
        assert!(followed.preview, "placed for the cursor");
        let placed = (pl.pos_x, pl.pos_y, pl.width, pl.height);

        // The cursor has not moved, so neither place has changed: nothing to move.
        assert_eq!(
            pl.follow_pointer(POINT { x: 300, y: 300 }, TEST_DPI),
            Followed {
                spinner: false,
                preview: false
            }
        );
        assert_eq!((pl.pos_x, pl.pos_y, pl.width, pl.height), placed);

        // The cursor has moved: the preview is placed again, somewhere else.
        let followed = pl.follow_pointer(POINT { x: 200, y: 300 }, TEST_DPI);
        assert!(followed.preview);
        assert_ne!((pl.pos_x, pl.pos_y, pl.width, pl.height), placed);

        // A wait is the spinner's own box rather than the preview's, whatever the
        // preview is: an 800 by 600 picture is waited for in the arc's own box at the
        // pointer's own corner, a pointer gap off it — the box a document waiting on a
        // page is placed in — while the preview keeps the place and the size its own
        // layout gave it.
        let mut waiting = pending(Some(placement));
        let followed = waiting.follow_pointer(POINT { x: 300, y: 300 }, TEST_DPI);
        assert!(followed.spinner);
        let gap = logical_px(TEST_DPI, POINTER_STANDOFF_PIXELS);
        assert_eq!(
            waiting.spinner_pos,
            (300 + gap, 300 + gap),
            "the spinner's own corner is a gap off the cursor, not under it"
        );
        assert_eq!(waiting.spinner_side, office_preview::WAITING_BOX);
        let preview = compute_mouse_layout(
            300,
            300,
            placement,
            monitor_bounds_from_point(300, 300),
            TEST_DPI,
        )
        .expect("a placed preview");
        assert_eq!(
            (waiting.pos_x, waiting.pos_y),
            (preview.pos_x, preview.pos_y),
            "the preview keeps the place it will arrive at"
        );
        assert_eq!(
            (waiting.width, waiting.height),
            (preview.preview_w, preview.preview_h),
            "and the box it will arrive in"
        );

        // A cursor moving along the item takes the spinner with it — still the arc's
        // own box, still a gap off the hand — and the preview's place with it.
        let followed = waiting.follow_pointer(POINT { x: 500, y: 300 }, TEST_DPI);
        assert!(followed.spinner);
        assert_eq!(waiting.spinner_pos, (500 + gap, 300 + gap));
        assert_eq!(waiting.spinner_side, office_preview::WAITING_BOX);

        // A keyboard hover's placement is the item's own: it follows nothing.
        let mut keyboard = pending(None);
        assert_eq!(
            keyboard.follow_pointer(POINT { x: 640, y: 400 }, TEST_DPI),
            Followed {
                spinner: false,
                preview: false
            }
        );
        assert_eq!((keyboard.pos_x, keyboard.pos_y), (0, 0));
    }

    /// The layout applies the reduction the same way the size it plans does, so a
    /// preview placed for a page is the reduction of the one fit-to-screen would
    /// have placed — not the room's own size with a percentage applied to it
    /// somewhere else.
    #[test]
    fn a_layout_reduces_a_page_by_the_configured_share() {
        let page = |preview_scale: PreviewScale| HoverPlacement {
            orig_dims: (800, 600),
            avoid: None,
            follow_cursor: false,
            preview_scale,
            at_the_pointer_corner: false,
        };

        let full = compute_mouse_layout(
            300,
            300,
            page(PreviewScale::FitToScreen),
            bounds(),
            TEST_DPI,
        )
        .expect("a placed page");
        let half = compute_mouse_layout(
            300,
            300,
            page(PreviewScale::FitToScreenReduced(50)),
            bounds(),
            TEST_DPI,
        )
        .expect("a placed page");

        assert_eq!(
            (full.preview_w, full.preview_h),
            (half.preview_w * 2, half.preview_h * 2)
        );
    }

    /// The spinner a hover is waiting on is placed at the pointer's own corner — the one
    /// of the four the display has room for, a pointer gap off the cursor so the window is
    /// not under it — with no step off the name it covers, so the wait stays at the hand
    /// that is waiting on it while leaving that hand free to click and probe the file
    /// underneath (see `waiting_placement`).
    #[test]
    fn places_the_waiting_spinner_at_the_pointers_own_corner() {
        let side = office_preview::WAITING_BOX;
        let name = (100, 300, 400, 320);
        let gap = logical_px(TEST_DPI, POINTER_STANDOFF_PIXELS);
        let spinner = |cursor_x: i32, cursor_y: i32| {
            compute_mouse_layout(
                cursor_x,
                cursor_y,
                HoverPlacement {
                    orig_dims: (side, side),
                    avoid: Some(name),
                    follow_cursor: false,
                    preview_scale: PreviewScale::Percent(100),
                    at_the_pointer_corner: true,
                },
                bounds(),
                TEST_DPI,
            )
            .expect("a placed spinner")
        };

        // Room in every quadrant: the spinner sits in the pointer's own corner, over the
        // name it is waiting on, one pointer gap off the cursor rather than in it.
        let placement = spinner(300, 300);
        assert_eq!((placement.pos_x, placement.pos_y), (300 + gap, 300 + gap));
        assert_eq!((placement.preview_w, placement.preview_h), (side, side));

        // And the room that layout comes out at is that corner of the display rather than
        // the display itself. It is the spinner's own room and says nothing about how large
        // the preview it is waiting for will be drawn, which is why it is not the room the
        // engine that draws that preview is asked for (see `PendingLoad::room`).
        assert_eq!(
            (placement.max_width, placement.max_height),
            (
                (bounds().right - 300 - gap) as u32,
                (bounds().bottom - 300 - gap) as u32,
            ),
            "the room is the corner the spinner was put in, not the display"
        );

        // With the display ending just past the pointer there is no room in the
        // quadrant it grows into, so the spinner takes the corner that is visible:
        // its box placed to the pointer's left, the same gap off it.
        let placement = spinner(980, 300);
        assert_eq!(
            (placement.pos_x, placement.pos_y),
            (980 - gap - side as i32, 300 + gap)
        );
    }

    #[test]
    fn leaves_a_placement_that_is_already_clear_of_the_text() {
        // The name is drawn to the left of the cursor's column, so the preview beside
        // the cursor is already off it.
        let name = (100, 300, 400, 320);
        let placement = placed(layout(420, 300, 300, 300), (300, 300), name, None, bounds());

        assert_eq!((placement.pos_x, placement.pos_y), (420, 300));
    }

    #[test]
    fn moves_a_preview_out_of_the_name_it_covers() {
        // A row of a Details view: the pointer's item draws its name in a band twenty
        // pixels tall, and the preview came out beside the cursor with its top inside
        // that band. Down is the shortest way out, so it ends up just under the name,
        // where its own column already was.
        let name = (100, 300, 400, 320);
        let placement = placed(layout(120, 300, 300, 300), (300, 300), name, None, bounds());

        assert_eq!((placement.pos_x, placement.pos_y), (120, 340));
        assert_eq!((placement.preview_w, placement.preview_h), (300, 300));
    }

    /// A way out that would put the preview over the pointer is not one a mouse hover
    /// can use. A mouse preview is dismissed the moment the cursor touches it, so one
    /// placed over the cursor is dismissed the instant it appears — and put back the
    /// same way a moment later, for as long as the pointer sits there, which is a
    /// preview that blinks at the hand rather than one that is read. The tile figures
    /// are the case that reaches it: the pointer is on a thumbnail with the label below
    /// it, and the step that clears the label by rising over it is the thumbnail the
    /// pointer is standing on — see `avoiding_text`.
    #[test]
    fn keeps_a_placement_off_the_pointer_it_is_for() {
        let label = (400, 600, 700, 620);
        let thumbnail = (500, 450);

        // With nothing said about the pointer, the shortest step that clears the label
        // is the one over it — which lands on the thumbnail the pointer is on.
        let blind = placed(
            layout(420, 400, 300, 300),
            (300, 300),
            label,
            None,
            bounds(),
        );
        assert_eq!((blind.pos_x, blind.pos_y), (420, 280));
        assert!(
            box_holds(
                thumbnail.0,
                thumbnail.1,
                (
                    blind.pos_x,
                    blind.pos_y,
                    blind.pos_x + blind.preview_w as i32,
                    blind.pos_y + blind.preview_h as i32,
                ),
            ),
            "the step over the label is the thumbnail the pointer is on"
        );

        // Held to the pointer, the same hover takes the way out that clears both: the
        // room before the label, which the thumbnail's own column has.
        let placement = placed(
            layout(420, 400, 300, 300),
            (300, 300),
            label,
            Some(thumbnail),
            bounds(),
        );
        assert_eq!((placement.pos_x, placement.pos_y), (80, 400));
        assert!(
            !box_holds(
                thumbnail.0,
                thumbnail.1,
                (
                    placement.pos_x,
                    placement.pos_y,
                    placement.pos_x + placement.preview_w as i32,
                    placement.pos_y + placement.preview_h as i32,
                ),
            ),
            "the preview is clear of the pointer"
        );
        assert!(
            placement.pos_x + placement.preview_w as i32 <= label.0,
            "and still clear of the label it was moved off"
        );
    }

    #[test]
    fn takes_the_side_when_the_display_has_no_room_under_the_name() {
        // The same placement on a display that ends below it: the preview cannot drop
        // under the name at its own size, so it goes past the name's right edge, which
        // the display has room for.
        let name = (100, 300, 400, 320);
        let short = ScreenBounds {
            bottom: 500,
            ..bounds()
        };
        let placement = placed(layout(120, 300, 300, 300), (300, 300), name, None, short);

        assert_eq!((placement.pos_x, placement.pos_y), (420, 300));
        assert_eq!((placement.preview_w, placement.preview_h), (300, 300));
    }

    #[test]
    fn resizes_a_preview_too_large_to_get_off_the_name() {
        // A preview as large as the display allows, beside a row whose name spans most
        // of it: no way out holds the preview as it stands, so the size is what gives
        // and the roomiest way out is taken. Under the name, that is the full width the
        // mode allowed and the height the display leaves below the row.
        let name = (100, 300, 700, 320);
        let placement = placed(layout(0, 0, 1000, 800), (1000, 800), name, None, bounds());

        assert_eq!((placement.pos_x, placement.pos_y), (0, 340));
        assert_eq!((placement.preview_w, placement.preview_h), (575, 460));
        // The name it was moved off is clear of it: the top edge is the row's bottom
        // plus the gap the claim above was measured with.
        assert!(placement.pos_y >= 320 + 20);
    }

    #[test]
    fn leaves_a_preview_alone_when_every_way_out_is_a_sliver() {
        // A display with no room worth having on any side of the name: the placement is
        // what the mode chose, since a preview squeezed into a sliver says less than the
        // one left over the name.
        let name = (40, 80, 360, 100);
        let tight = ScreenBounds {
            left: 0,
            top: 0,
            right: 400,
            bottom: 180,
        };
        let placement = placed(layout(150, 80, 200, 90), (200, 90), name, None, tight);

        assert_eq!((placement.pos_x, placement.pos_y), (150, 80));
    }

    #[test]
    fn keeps_the_preview_inside_the_display_it_moves_on() {
        // Clearing the name to the right falls short of what the preview needs here, so
        // the step under the name is taken instead — resized into the room it leaves
        // when even that is not enough for it as it stands.
        let name = (100, 300, 700, 320);
        let narrow = ScreenBounds {
            right: 900,
            ..bounds()
        };
        let placement = placed(layout(690, 100, 200, 300), (200, 300), name, None, narrow);

        // 720 is the name's right edge plus the gap, and 720 + 200 leaves the display;
        // 340 is its bottom plus the gap, and 340 + 300 does not.
        assert_eq!((placement.pos_x, placement.pos_y), (690, 340));
    }

    /// The room the display has is its whole work area: the largest box anything shown on
    /// it can be drawn in, and the box an engine that has to draw a preview before the file
    /// can be measured is asked for (see `PendingLoad::room`). A work area with no room in
    /// it is a pixel rather than nothing, because a box of nothing is an engine asked to
    /// write nothing.
    #[test]
    fn a_display_room_is_the_whole_of_its_work_area() {
        assert_eq!(bounds().room(), (1000, 800));

        assert_eq!(
            ScreenBounds {
                left: 100,
                top: 40,
                right: 1800,
                bottom: 1040,
            }
            .room(),
            (1700, 1000),
            "the room is the work area whatever corner it is anchored at"
        );

        assert_eq!(
            ScreenBounds {
                left: 0,
                top: 0,
                right: 0,
                bottom: 0,
            }
            .room(),
            (1, 1),
            "a display with no room is asked for a pixel rather than for nothing"
        );
    }

    /// A planned preview is inside the display it was planned for: the box it takes is
    /// within the room it was given, that room is within the room the display itself has,
    /// and the whole of it — place and size, both axes — is within that display's work area.
    fn assert_inside_the_display(layout: PreviewLayout, bounds: ScreenBounds) {
        assert!(
            layout.preview_w <= layout.max_width && layout.preview_h <= layout.max_height,
            "the preview takes {} by {} of the {} by {} it was given",
            layout.preview_w,
            layout.preview_h,
            layout.max_width,
            layout.max_height
        );

        // The room a placement hands its preview is taken out of the display — the side of
        // the pointer it goes beside, the quadrant it grows into, the room left past a name —
        // so the display's own room bounds every one of them. That is what makes it the box
        // to ask an engine for when the file cannot be measured before it is drawn: a picture
        // developed into a smaller box than this could never be drawn at the size a layout
        // asks for (see `PendingLoad::room`).
        let room = bounds.room();
        assert!(
            layout.max_width <= room.0 && layout.max_height <= room.1,
            "the room a layout was given, {} by {}, leaves the display's own room of {} by {}",
            layout.max_width,
            layout.max_height,
            room.0,
            room.1
        );

        assert!(
            layout.pos_x >= bounds.left
                && layout.pos_y >= bounds.top
                && layout.pos_x + layout.preview_w as i32 <= bounds.right
                && layout.pos_y + layout.preview_h as i32 <= bounds.bottom,
            "the preview at {},{} is {} by {}, which leaves the display of {} by {} at {},{}",
            layout.pos_x,
            layout.pos_y,
            layout.preview_w,
            layout.preview_h,
            bounds.right - bounds.left,
            bounds.bottom - bounds.top,
            bounds.left,
            bounds.top
        );
    }

    /// A display and the scale it is drawn at, which is one case for a placement test.
    type ScaledDisplay = (ScreenBounds, u32);

    /// What the placement matrices are made of.
    struct PlacementCases {
        displays: Vec<ScaledDisplay>,
        media: [(u32, u32); 7],
        scales: [PreviewScale; 4],
    }

    /// The displays, the display scales and the media every placement test is run
    /// over: a 1080p display, a 4K one, a 4K one that is the second display rather than
    /// the first, and a portrait one — each at 100%, 150% and 200%, because the margins
    /// a placement is written around are scaled by the display it is on and the ones
    /// that came apart at a scale are what the scaling is for — against the shapes media
    /// comes in: pages both ways up, a slide, the waiting spinner, a picture, and things
    /// far wider and far taller than any display.
    fn placement_cases() -> PlacementCases {
        let geometries = [
            ScreenBounds {
                left: 0,
                top: 0,
                right: 1920,
                bottom: 1040,
            },
            ScreenBounds {
                left: 0,
                top: 0,
                right: 3840,
                bottom: 2090,
            },
            ScreenBounds {
                left: 1920,
                top: 0,
                right: 5760,
                bottom: 2090,
            },
            ScreenBounds {
                left: 0,
                top: 0,
                right: 1440,
                bottom: 2500,
            },
        ];
        let media = [
            (794, 1123),
            (1123, 794),
            (1920, 1080),
            (36, 36),
            (100, 100),
            (8000, 120),
            (120, 8000),
        ];
        let scales = [
            PreviewScale::FitToScreen,
            PreviewScale::FitToScreenReduced(50),
            PreviewScale::Percent(100),
            PreviewScale::Percent(400),
        ];
        // Every display at every scale, as one list: the margins a placement is
        // written around are scaled by the display it is on, so a display at 150% is
        // not the display at 100%, and the pair is what a case is.
        let displays = geometries
            .iter()
            .flat_map(|bounds| [TEST_DPI, 144, 192].map(|dpi| (*bounds, dpi)))
            .collect();

        PlacementCases {
            displays,
            media,
            scales,
        }
    }

    /// The one thing a hovered preview may never do: leave the display it was planned
    /// for. Every position mode, every scale, every shape of media, anchored at each
    /// corner of the display and at its middle, with a name under the pointer and with a
    /// row across the display's top to be kept off — because a preview that grows past
    /// its display takes an edge and a room that disagree to find, and the ones that
    /// disagree are not the ones anyone hovers over on purpose.
    #[test]
    fn places_every_hover_inside_its_display() {
        let PlacementCases {
            displays,
            media,
            scales,
        } = placement_cases();
        let mut placed = 0usize;

        for (bounds, dpi) in displays {
            let points = [
                (bounds.left, bounds.top),
                (bounds.right - 1, bounds.top),
                (bounds.left, bounds.bottom - 1),
                (bounds.right - 1, bounds.bottom - 1),
                (
                    (bounds.left + bounds.right) / 2,
                    (bounds.top + bounds.bottom) / 2,
                ),
            ];

            for (orig_width, orig_height) in media {
                for preview_scale in scales {
                    for follow_cursor in [true, false] {
                        for (cursor_x, cursor_y) in points {
                            let names = [
                                None,
                                // The name of the item the pointer is on.
                                Some((cursor_x - 100, cursor_y - 8, cursor_x + 100, cursor_y + 8)),
                                // A row's text across the whole display's top, which a
                                // preview above the middle of the display covers.
                                Some((bounds.left, bounds.top, bounds.right, bounds.top + 40)),
                            ];

                            for avoid in names {
                                let placement = HoverPlacement {
                                    orig_dims: (orig_width, orig_height),
                                    avoid,
                                    follow_cursor,
                                    preview_scale,
                                    // Only the waiting spinner is placed at the
                                    // pointer's own corner, and only it is ever that size.
                                    at_the_pointer_corner: (orig_width, orig_height) == (36, 36),
                                };

                                let layout = compute_mouse_layout(
                                    cursor_x, cursor_y, placement, bounds, dpi,
                                );
                                let Some(layout) = layout else {
                                    continue;
                                };

                                assert_inside_the_display(layout, bounds);
                                placed += 1;
                            }
                        }
                    }
                }
            }
        }

        // A matrix that answers nothing checks nothing: nearly every combination of
        // point, media, scale and setting has a layout to check.
        assert!(
            placed >= 3000,
            "{placed} of the matrix's layouts were placed, which is too few to have checked the rest"
        );
    }

    /// The same promise for the preview a keyboard hover raises, which is placed from
    /// the item rather than from the pointer: a row of a list, a box item, and an item
    /// the display has scrolled half off its top and half off its bottom — the cases
    /// where the room beside an item and the item's own edges disagree.
    #[test]
    fn places_every_keyboard_hover_inside_its_display() {
        let PlacementCases {
            displays,
            media,
            scales,
        } = placement_cases();
        let mut placed = 0usize;

        for (bounds, dpi) in displays {
            let items = [
                // A row of a list, drawn across the view at its middle.
                (
                    bounds.left,
                    (bounds.top + bounds.bottom) / 2,
                    bounds.right,
                    (bounds.top + bounds.bottom) / 2 + 40,
                ),
                // A box item, the way an icon view draws one.
                (
                    bounds.left + 100,
                    bounds.top + 100,
                    bounds.left + 220,
                    bounds.top + 220,
                ),
                // The first row of a scrolled list, half off the display's top.
                (bounds.left, bounds.top - 20, bounds.right, bounds.top + 20),
                // And the last one, half off its bottom.
                (
                    bounds.left,
                    bounds.bottom - 20,
                    bounds.right,
                    bounds.bottom + 20,
                ),
            ];

            for (item_left, item_top, item_right, item_bottom) in items {
                for (orig_width, orig_height) in media {
                    for preview_scale in scales {
                        for follow_cursor in [true, false] {
                            let avoids = [
                                None,
                                // The name the row is listed under, at its left end.
                                Some((
                                    item_left + 8,
                                    item_top + 8,
                                    item_left + 220,
                                    item_bottom - 8,
                                )),
                                // And the row's whole width, as `Avoid Details` reads it.
                                Some((item_left, item_top, item_right, item_bottom)),
                            ];

                            // Read as a row of the view and as a box item, since what
                            // the item draws beside its name decides which of the two
                            // placements it gets and both are promises of their own:
                            // the item is placed from the region it is kept off or
                            // from its own edges, whichever its text says it is.
                            for columns in [true, false] {
                                for avoid in avoids {
                                    let Some(layout) = compute_keyboard_layout(
                                        KeyboardPlacement {
                                            item_rect: (
                                                item_left,
                                                item_top,
                                                item_right,
                                                item_bottom,
                                            ),
                                            avoid,
                                            columns,
                                            orig_dims: (orig_width, orig_height),
                                            follow_cursor,
                                            preview_scale,
                                        },
                                        bounds,
                                        dpi,
                                    ) else {
                                        continue;
                                    };

                                    assert_inside_the_display(layout, bounds);
                                    placed += 1;
                                }
                            }
                        }
                    }
                }
            }
        }

        assert!(
            placed >= 2000,
            "{placed} of the matrix's layouts were placed, which is too few to have checked the rest"
        );
    }

    /// The whole app path for one file, without Explorer: the preview loop is started,
    /// the file is shown the way a hover shows it, and what happens next is reported.
    /// Ignored, and driven by `RHP_APP_PROBE` —
    /// `$env:RHP_APP_PROBE = "C:\art\clock.svg"; cargo test -- --ignored --nocapture app_hover_probe`
    /// — for a document whose preview does not appear.
    #[test]
    #[ignore = "shows a preview window"]
    fn app_hover_probe() {
        let Ok(path) = std::env::var("RHP_APP_PROBE") else {
            println!("set RHP_APP_PROBE to a path");
            return;
        };
        let path = PathBuf::from(path);

        println!(
            "engine available: {}",
            crate::engines::webview_preview::is_available()
        );
        println!(
            "document: {}",
            crate::engines::webview_preview::draws(&path)
        );
        // Which engine would play a video here, which is the whole of the fallback's
        // routing: `ffplay` when it is installed, and the media engine Windows has when it
        // is not. Reported for every file rather than only for a video, because a probe is
        // run to find out what the machine is doing.
        println!(
            "video: played natively = {}",
            crate::formats::codecs::plays_video_natively()
        );

        std::thread::spawn(run_preview_window);
        std::thread::sleep(Duration::from_millis(500));

        show_preview(&path, 200, 200, None);

        for step in 0..25 {
            std::thread::sleep(Duration::from_millis(200));

            let media = CURRENT_MEDIA.lock().ok().and_then(|media| {
                media.as_ref().map(|media| {
                    (
                        media.current_width(),
                        media.current_height(),
                        media.frames.len(),
                        media.media_type.is_native_video(),
                        // Whether the placeholder a video preview is opened with has been
                        // replaced by a frame of it, which is the one thing a video that
                        // plays and a video that does not look different in.
                        media
                            .current_pixels()
                            .as_chunks::<4>()
                            .0
                            .iter()
                            .any(|pixel| pixel[..3] != [40, 40, 40]),
                    )
                })
            });

            println!(
                "{:>5} ms: engine_showing={} (w, h, frames, native video, framed)={:?}",
                (step + 1) * 200,
                crate::engines::webview_preview::is_showing(),
                media
            );

            if crate::engines::webview_preview::is_showing() {
                break;
            }
        }

        // Kept up long enough for the screen to be looked at, for the probes that
        // measure what is on it rather than how long it took to get there.
        let hold = std::env::var("RHP_APP_PROBE_HOLD_MS")
            .ok()
            .and_then(|ms| ms.trim().parse().ok())
            .unwrap_or(0);
        if hold > 0 {
            std::thread::sleep(Duration::from_millis(hold));
        }

        hide_preview();
        std::thread::sleep(Duration::from_millis(300));
        println!(
            "after hide: engine_showing={}",
            crate::engines::webview_preview::is_showing()
        );
    }

    /// A video replaced in place is probed again rather than cropped and sized by
    /// the answer about the file it used to be: the version is part of what a probed
    /// geometry is held for.
    #[test]
    fn keys_a_probed_geometry_by_the_files_version() {
        // A folder of this module's own: the tests run beside each other, and one
        // of them clearing its fixtures must not take another's with it.
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-video-tests")
            .join("geometry");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let path = folder.join("keyed.mp4");
        std::fs::write(&path, b"one").expect("a written file");

        let key = |path: &PathBuf| VideoGeometryKey {
            path: path.clone(),
            version: file_version(path),
        };

        let first = key(&path);
        assert!(first == key(&path), "the same file is the same key");

        std::fs::write(&path, b"a longer file").expect("a rewritten file");
        assert!(first != key(&path), "a rewritten file is another key");

        let _ = std::fs::remove_file(&path);
    }

    /// The report an `ffprobe` pass writes is read for the sound in it and nothing of the
    /// picture before it: a film with a soundtrack is not a sound, and the fields a card is
    /// drawn with are the ones that follow the sound's own stream.
    #[test]
    fn reads_a_sound_out_of_an_ffprobe_report() {
        let report = "codec_type=video\ncodec_name=h264\nwidth=1920\n\
                      codec_type=audio\ncodec_name=flac\nsample_rate=44100\nchannels=2\n\
                      bit_rate=1006000\nduration=562.31\n";

        let track = audio_track_from_report(report).expect("a sound in the report");
        assert_eq!(track.player, Player::Ffmpeg);
        assert_eq!(track.codec.as_deref(), Some("FLAC"));
        assert_eq!(track.rate, Some(44_100));
        assert_eq!(track.channels, Some(2));
        assert_eq!(track.bitrate, Some(1_006_000));
        assert_eq!(track.duration, Some(562.31));

        assert_eq!(
            audio_track_from_report("codec_type=video\ncodec_name=h264\nduration=10.0\n"),
            None,
            "a file with no sound stream in it is not a sound, however it is named"
        );
    }

    /// What this machine has for the sounds named in `RHP_AUDIO_PROBE`, and what their cards
    /// come out as.
    ///
    /// Ignored by default, like every other probe here: it reads real files, it starts no
    /// player and it plays nothing — what it asks is the question a hover asks before a card is
    /// drawn, and what it prints is the answer beside the size the card was painted at.
    ///
    /// ```text
    /// $env:RHP_AUDIO_PROBE = "C:\music\track.flac;C:\music\podcast.opus"
    /// cargo test audio_probe -- --ignored --nocapture
    /// ```
    #[test]
    #[ignore = "reads the files named in RHP_AUDIO_PROBE"]
    fn audio_probe() {
        let Ok(list) = std::env::var("RHP_AUDIO_PROBE") else {
            println!("set RHP_AUDIO_PROBE to one or more paths, separated by ';'");
            return;
        };

        for path in list
            .split(';')
            .map(str::trim)
            .filter(|path| !path.is_empty())
            .map(PathBuf::from)
        {
            println!("\n--- {} ---", path.display());
            println!(
                "the lists call it a sound: {}, and the preview is shown: {}",
                crate::formats::audio_formats::matches_audio_list(
                    &path,
                    &CONFIG.lock().map(|config| config.audio_extensions.clone()).unwrap_or_default(),
                ),
                drawn_as_audio(&path)
            );

            let started = Instant::now();
            let probed = probe_audio_track(&path);
            println!("probe: {probed:?} ({} ms)", started.elapsed().as_millis());

            // A container of a video's name is the case the whole verdict exists for: what the
            // video probe finds in it, and what the router says once that probe has answered.
            let named_video = CONFIG
                .lock()
                .map(|config| {
                    video_formats::matches_video_list(&path, &config.video_extensions)
                })
                .unwrap_or(false);
            if named_video {
                println!(
                    "the video probe answered {}",
                    match probe_video_geometry(&path) {
                        ProbedGeometry::Measured(_) => "a shape",
                        ProbedGeometry::Unmeasurable => "nothing to measure",
                    }
                );
                println!(
                    "and the router calls it {:?}",
                    CONFIG
                        .lock()
                        .ok()
                        .and_then(|config| crate::formats::routing::kind_of(&path, &config))
                );
                println!("its card is drawn: {}", drawn_as_audio(&path));
            }

            let Some(track) = probed else {
                println!("nothing here plays this file");
                continue;
            };

            println!("the card says: {:?}", audio_preview::facts_of(&track, &path));
            for (elapsed, duration) in [(None, None), (Some(67.0), track.duration), (Some(0.5), None)] {
                let Some(card) = audio_card(&path, elapsed, duration, 0) else {
                    continue;
                };
                let options = current_audio_options();
                let (width, height) =
                    audio_preview::measure(&card, 4096, 2160, 96, options).expect("a measured card");
                let painted =
                    audio_preview::render(&card, width, height, 96, options).expect("a painted card");

                println!(
                    "card at {elapsed:?} / {duration:?}: {width}x{height}, {} bytes of frame",
                    painted.0.len()
                );
            }

            // And what the engine actually does with the file, which is the question the probe
            // beside it does not answer: a probe says this machine has a decoder, and what a
            // hover needs is a player that gets somewhere. The two came apart once — a name the
            // engine resolved as a URL was refused after `Play` had already answered, so a file
            // probed as playable drew a card whose clock never moved (see `Session::begin`) —
            // and what tells that apart from a file that plays is the position below: a session
            // that failed reports a state of not playing and a position stuck at nothing, and
            // one that is playing reports both moving. Only the engine is asked: FFmpeg's
            // player is a process of its own and nothing on this side is drawn from it.
            if track.player == Player::Native {
                let volume = current_audio_volume().max(1);
                let seek = current_audio_seek();
                let start = audio_seek::start_position(&path, seek, track.duration);
                println!("starting at {start:.3}s, by `Volume → Audio Seek`");

                video_player::play_audio(&path, volume, start);
                std::thread::sleep(Duration::from_millis(500));

                println!(
                    "playing natively: {}, at {:?} of {:?} ({})",
                    video_player::is_playing(),
                    video_player::position(),
                    video_player::duration(),
                    if video_player::playing_path().as_deref() == Some(path.as_path()) {
                        "the file that was asked for"
                    } else {
                        "not the file that was asked for"
                    },
                );
                video_player::stop();
            }
        }
    }

    /// A video that has not been probed yet is a hover that is waiting, so its box is the
    /// wait's — and the probe's answer, whatever it is, is what the box becomes: a shape
    /// is the shape, and a file with nothing to measure is the box FFmpeg's player is
    /// given rather than one the probe is asked for again.
    #[test]
    fn a_video_that_has_not_been_probed_waits_in_the_waiting_box() {
        // A folder of this module's own, and a name no other test uses: the cache is
        // shared, and what this test puts in it must be its own key.
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-video-tests")
            .join("probe-box");
        std::fs::create_dir_all(&folder).expect("a test folder");
        let path = folder.join("waited-for.mp4");
        std::fs::write(&path, b"a file the probe has not seen").expect("a written file");

        let key = VideoGeometryKey {
            path: path.clone(),
            version: file_version(&path),
        };
        assert!(
            cached_video_geometry(&path).is_none(),
            "nothing has probed this file"
        );
        assert_eq!(
            video_box(&path),
            Some((office_preview::WAITING_BOX, office_preview::WAITING_BOX)),
            "an unprobed video is placed as the wait for its probe"
        );

        // The answer the probe gives is what the hover is placed at, and it is read from
        // the cache rather than measured again.
        video_geometry_cache().insert(
            key.clone(),
            ProbedGeometry::Measured(VideoGeometry {
                width: 640,
                height: 360,
                crop: None,
            }),
        );
        assert_eq!(video_box(&path), Some((640, 360)));

        // A file the probe could not measure is not a file to probe again: it is the box
        // the player that would try the file anyway is given, and no preview at all where
        // the engine that plays it cannot open it either.
        video_geometry_cache().insert(key, ProbedGeometry::Unmeasurable);
        assert_eq!(
            video_box(&path),
            (!codecs::plays_video_natively()).then_some((1920, 1080))
        );

        let _ = std::fs::remove_file(&path);
    }

    /// The wait for a video's player ends one of two ways and never by itself: a player
    /// whose window is up has arrived, a player that is gone is not coming, and a start
    /// that has run past the cap is given up on — while a player that is alive with no
    /// window yet is a start that is still going.
    #[test]
    fn waits_for_a_player_only_until_it_is_there_or_gone() {
        let cap = Duration::from_secs(VIDEO_START_WAIT_SECS);

        assert_eq!(
            player_wait(true, true, Duration::ZERO),
            Some(PlayerWait::Arrived),
            "a player with a window up is the video"
        );
        assert_eq!(
            player_wait(false, true, Duration::from_secs(1)),
            None,
            "a player still starting is a wait that goes on"
        );
        assert_eq!(
            player_wait(true, false, Duration::ZERO),
            Some(PlayerWait::Abandoned),
            "what a dead player leaves behind is a handle, not a preview"
        );
        assert_eq!(
            player_wait(false, false, Duration::ZERO),
            Some(PlayerWait::Abandoned),
            "a player that is gone is not coming back to put a window up"
        );
        assert_eq!(
            player_wait(false, true, cap),
            Some(PlayerWait::Abandoned),
            "a start past the cap is not watched any longer"
        );
        assert_eq!(
            player_wait(true, true, cap),
            Some(PlayerWait::Arrived),
            "and a window that is up has arrived, cap or no cap"
        );
    }

    /// A player that has stopped is read as one that reached the end of the file it was handed or
    /// as one that never played it, and the difference is time: a stop well into the pass the
    /// player was given is the end of the file, and a stop within a moment of starting is a
    /// player that failed rather than a file that ended. What the file says about its length
    /// bounds the pass and never overrules the moment, so the last of a short file — a pass
    /// shorter than a moment — is asked only to have been played.
    #[test]
    fn a_player_that_has_stopped_is_read_as_the_end_of_its_file_or_not() {
        let length = Some(200.0);

        assert!(
            reached_the_end(Duration::from_secs(200), length, 0.0),
            "a player that played the whole file reached the end of it"
        );
        assert!(
            reached_the_end(Duration::from_secs(31), length, 170.0),
            "and so did one given only the last half minute of it"
        );
        assert!(
            !reached_the_end(Duration::from_millis(80), length, 0.0),
            "a player that stopped the moment it started never played the file, long or not"
        );
        assert!(
            !reached_the_end(Duration::from_millis(80), Some(2_000.0), 0.0),
            "and a file of any length is not read from a stop that short"
        );
        assert!(
            reached_the_end(Duration::from_millis(900), Some(0.5), 0.0),
            "a pass shorter than the moment is read by its own length: it cannot be lived past"
        );
        assert!(
            reached_the_end(Duration::from_secs(2), None, 0.0),
            "a file that says nothing about its length is read by the moment alone"
        );
        assert!(
            !reached_the_end(Duration::from_millis(200), None, 0.0),
            "which is still a moment a player that failed has stopped inside of"
        );
    }

    /// A probe's child is waited for with a deadline, and what a child that has outrun it
    /// leaves behind is an answer of its own rather than a wait that goes on: the process is
    /// ended where it stands and the caller is told there is nothing from this one, which is
    /// the arm both probes already have for a child that answered nothing.
    ///
    /// What stands in for the two children is an ordinary console program, since what the
    /// helper does with a child is the same whatever the child is and a test is not going to
    /// start ffmpeg. Neither stand-in reaches the network: a command that prints a line
    /// finishes however loaded the machine is, and a ping of thirty replies is still there a
    /// moment later whatever it is told.
    #[test]
    fn waits_for_a_probes_child_only_until_the_deadline() {
        let quick = Command::new("cmd")
            .args(["/C", "echo", "a line from the child"])
            .stdout(Stdio::piped())
            .stderr(Stdio::null())
            .spawn()
            .expect("a child that finishes");

        let answered = wait_bounded(quick, Duration::from_secs(VIDEO_PROBE_TIMEOUT_SECS))
            .expect("a child that ends inside the deadline answers with its output");
        assert!(answered.status.success());
        assert!(
            String::from_utf8_lossy(&answered.stdout).contains("a line from the child"),
            "and what it wrote is in the answer, which is the pipe that was drained"
        );

        let slow = Command::new("ping")
            .args(["-n", "30", "127.0.0.1"])
            .stdout(Stdio::piped())
            .stderr(Stdio::null())
            .spawn()
            .expect("a child that does not finish");
        let pid = slow.id();

        assert!(
            wait_bounded(slow, Duration::from_millis(50)).is_none(),
            "a child that has outrun the deadline is not waited for any longer"
        );
        assert!(
            !engine_processes::is_running(pid),
            "and it is ended rather than left reading a file nobody is waiting for"
        );
    }

    /// The whole path a hover takes for a document that is already on disk: measure
    /// it, place it, load it. Ignored, and driven by `RHP_OFFICE_PROBE` —
    /// `$env:RHP_OFFICE_PROBE = "C:\docs\one.xlsx"; cargo test -- --ignored --nocapture office_hover_probe`
    /// — for a document whose preview does not appear.
    #[test]
    #[ignore = "reads the files named in RHP_OFFICE_PROBE"]
    fn office_hover_probe() {
        // The drawing half of the app runs in a multithreaded apartment.
        pdf_preview::initialize_apartment();

        let Ok(list) = std::env::var("RHP_OFFICE_PROBE") else {
            println!("set RHP_OFFICE_PROBE to one or more paths, separated by ';'");
            return;
        };

        // A display with a cursor on it, which is what a hover arrives with.
        let bounds = ScreenBounds {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        let (cursor_x, cursor_y) = (900, 500);
        let dpi = 96;

        for path in list
            .split(';')
            .map(str::trim)
            .filter(|path| !path.is_empty())
        {
            let path = PathBuf::from(path);
            println!("\n--- {} ---", path.display());

            // Which engine draws a page for this document, which is the question a preview
            // that arrives smaller than it should turn on: a page the render engine beside
            // Office drew is a page of the document's own size, and a hover laid out as the
            // wait for one places it at the spinner's box (see `libre_formats`).
            println!(
                "engines: office tier = {}, chosen engine = {:?}, render engine = {:?}, \
                 application installed = {}",
                office_render_is_due(&path, 800),
                office_formats::page_engine(&path),
                libre_formats::engine_page_kind(&path),
                office_formats::app_installed(&path),
            );
            println!(
                "render engine page: {}, refused = {}",
                libreoffice_render::rendered_page(&path)
                    .map(|page| page.display().to_string())
                    .unwrap_or_else(|| "none".to_string()),
                libreoffice_render::refused(&path)
            );

            let configured = current_hover_scales().picture;
            let scale = effective_preview_scale(&path, current_hover_scales());
            println!("scale: configured {configured:?}, effective {scale:?}");

            let Some(dimensions) = media_dimensions(&path, bounds, dpi) else {
                println!("measure: nothing — the hover shows no preview");
                continue;
            };
            println!("measure: {dimensions:?}");

            let Some(layout) = compute_mouse_layout(
                cursor_x,
                cursor_y,
                HoverPlacement {
                    orig_dims: dimensions,
                    avoid: None,
                    follow_cursor: false,
                    preview_scale: scale,
                    at_the_pointer_corner: false,
                },
                bounds,
                dpi,
            ) else {
                println!("layout: none — the hover shows no preview");
                continue;
            };
            println!(
                "layout: {}x{} at ({}, {}), free room {}x{}",
                layout.preview_w,
                layout.preview_h,
                layout.pos_x,
                layout.pos_y,
                layout.max_width,
                layout.max_height
            );

            let cancel = Arc::new(AtomicBool::new(false));
            let started = Instant::now();
            match load_media(
                &path,
                layout.max_width,
                layout.max_height,
                scale,
                dpi,
                Arc::clone(&cancel),
            ) {
                Some(media) => println!(
                    "loaded: {}x{}, {} frame(s), in {:?}",
                    media.current_width(),
                    media.current_height(),
                    media.frames.len(),
                    started.elapsed()
                ),
                None => println!(
                    "loaded: nothing — the hover blinks, in {:?}",
                    started.elapsed()
                ),
            }
        }
    }

    /// The whole path a hover takes for a document an installed render engine draws, which
    /// is the one path with a wait in the middle of it: measured as the wait for a page, the
    /// page asked for, the wait watched, and then — when the engine has drawn it — measured
    /// and loaded from the page. It also answers the order of the lists, which is what
    /// decides whether the engine is asked about a file at all. Ignored, and driven by
    /// `RHP_LIBRE_PROBE` —
    /// `$env:RHP_LIBRE_PROBE = "C:\art\logo.cdr"; cargo test -- --ignored --nocapture libre_hover_probe`
    /// — for a document whose preview does not appear.
    #[test]
    #[ignore = "reads the files named in RHP_LIBRE_PROBE and starts the installed LibreOffice"]
    fn libre_hover_probe() {
        pdf_preview::initialize_apartment();

        let Ok(list) = std::env::var("RHP_LIBRE_PROBE") else {
            println!("set RHP_LIBRE_PROBE to one or more paths, separated by ';'");
            return;
        };

        // A name the app would not hand the engine can be forced into this run's list,
        // which is how the give-up itself is watched: the files that show it are the ones
        // the engine cannot draw at all, and no name it answers for does that. This is the
        // list as it is held in memory, so the file on disk is not touched.
        if let Ok(forced) = std::env::var("RHP_LIBRE_PROBE_FORCE") {
            if let Ok(mut config) = crate::CONFIG.lock() {
                for name in forced
                    .split(',')
                    .map(str::trim)
                    .filter(|name| !name.is_empty())
                {
                    let name = name.trim_start_matches('.').to_lowercase();
                    if !config.libre_extensions.contains(&name) {
                        config.libre_extensions.push(name.clone());
                    }
                    println!("forced `{name}` into [libre] for this run");
                }
            }
        }

        let bounds = ScreenBounds {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        let dpi = 96;
        let (cursor_x, cursor_y) = (900, 500);

        for path in list
            .split(';')
            .map(str::trim)
            .filter(|path| !path.is_empty())
        {
            let path = PathBuf::from(path);
            println!("\n--- {} ---", path.display());
            println!("engine available: {}", libreoffice_render::available());
            let text_lists = {
                let config = crate::CONFIG.lock().expect("the configuration");
                crate::formats::text_formats::matches_text_lists(
                    &path,
                    &config.text_extensions,
                    &config.text_names,
                )
            };
            println!(
                "kinds: video = {}, pdf = {}, office = {}, libre = {}, design = {}, vector = {}, text = {}",
                crate::formats::video_formats::is_video_file(&path),
                pdf_preview::is_pdf_file(&path),
                office_formats::is_office_file(&path),
                libre_formats::is_libre_file(&path),
                design_formats::is_design_file(&path),
                vector_formats::is_vector_file(&path),
                text_lists,
            );

            let scale = effective_preview_scale(&path, current_hover_scales());
            println!("scale: {scale:?}");
            println!("render due: {}", libre_render_is_due(&path));
            println!(
                "measure before the page: {:?} — the spinner's own box is {}",
                media_dimensions(&path, bounds, dpi),
                office_preview::WAITING_BOX
            );

            // The hover the page is asked for, which is what the loop does the moment a
            // document like this is missed. A file this kind does not hold — one another
            // list reads, or one the engine has already turned down — has nothing to ask
            // for and nothing to wait on.
            let Some(_waiting) = request_libre_render(&path, 0) else {
                println!("request: nothing — the engine is not asked about this file");
                continue;
            };

            let started = Instant::now();
            let mut page = None;
            while page.is_none()
                && !libreoffice_render::refused(&path)
                && started.elapsed() < Duration::from_secs(90)
            {
                std::thread::sleep(Duration::from_millis(250));
                page = libreoffice_render::rendered_page(&path);
            }
            println!(
                "engine: page = {}, refused = {}, after {:?}",
                page.as_ref()
                    .map(|page| page.display().to_string())
                    .unwrap_or_else(|| "none".to_string()),
                libreoffice_render::refused(&path),
                started.elapsed()
            );

            if page.is_none() {
                continue;
            }

            // The replay, which is what the loop does when the page lands: the hover is
            // measured again — this time from the page — placed, and loaded.
            let Some(dimensions) = media_dimensions(&path, bounds, dpi) else {
                println!("measure after the page: nothing — the hover shows no preview");
                continue;
            };
            println!("measure after the page: {dimensions:?}");

            let Some(layout) = compute_mouse_layout(
                cursor_x,
                cursor_y,
                HoverPlacement {
                    orig_dims: dimensions,
                    avoid: None,
                    follow_cursor: false,
                    preview_scale: scale,
                    at_the_pointer_corner: false,
                },
                bounds,
                dpi,
            ) else {
                println!("layout: none — the hover shows no preview");
                continue;
            };

            let cancel = Arc::new(AtomicBool::new(false));
            let started = Instant::now();
            match load_media(
                &path,
                layout.max_width,
                layout.max_height,
                scale,
                dpi,
                Arc::clone(&cancel),
            ) {
                Some(media) => println!(
                    "loaded: {}x{}, {} frame(s), in {:?}",
                    media.current_width(),
                    media.current_height(),
                    media.frames.len(),
                    started.elapsed()
                ),
                None => println!(
                    "loaded: nothing — the hover blinks, in {:?}",
                    started.elapsed()
                ),
            }
        }
    }

    /// The whole path a hover takes for an archive an installed PeaZip lists, which is the
    /// third path with a wait in the middle of it: measured as the wait for a listing, the
    /// listing asked for, the wait watched, and then — when the engine has answered —
    /// measured and loaded from the listing itself. Ignored, and driven by `RHP_PEAZIP_PROBE` —
    /// `$env:RHP_PEAZIP_PROBE = "C:\downloads\backup.cab"; cargo test -- --ignored --nocapture peazip_hover_probe`
    /// — for an archive whose preview does not appear.
    ///
    /// It also answers the routing, which is what such a preview is usually about: the name the
    /// configured list carries, what the file's bytes and its name together make of it, whether
    /// the engine is installed to list it, and whether this hover is the wait for one at all.
    /// That last question is asked here of the loader's own predicate rather than of the list,
    /// because the list is only half of it: a file the engine will be asked about is still a
    /// file with no preview if the load that missed it is not read as a wait (see
    /// `awaiting_render` in the loader worker).
    #[test]
    #[ignore = "reads the files named in RHP_PEAZIP_PROBE and starts the installed PeaZip"]
    fn peazip_hover_probe() {
        let Ok(list) = std::env::var("RHP_PEAZIP_PROBE") else {
            println!("set RHP_PEAZIP_PROBE to one or more paths, separated by ';'");
            return;
        };

        let bounds = ScreenBounds {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        let dpi = 96;
        let (cursor_x, cursor_y) = (900, 500);

        for path in list
            .split(';')
            .map(str::trim)
            .filter(|path| !path.is_empty())
        {
            let path = PathBuf::from(path);
            println!("\n--- {} ---", path.display());

            let claimed = crate::CONFIG
                .lock()
                .map(|config| peazip_formats::matches_peazip_list(&path, &config.peazip_extensions))
                .unwrap_or(false);

            println!(
                "kinds: peazip list = {claimed}, engine archive = {}, engine available = {}",
                peazip_formats::is_engine_archive(&path),
                peazip_render::available()
            );
            println!(
                "state: refused = {}, listed = {}, render due = {}",
                peazip_render::refused(&path),
                peazip_render::listed(&path),
                peazip_render_is_due(&path)
            );

            let scale = effective_preview_scale(&path, current_hover_scales());
            println!("scale: {scale:?}");
            println!(
                "measure before the listing: {:?} — the spinner's own box is {}",
                media_dimensions(&path, bounds, dpi),
                office_preview::WAITING_BOX
            );

            // The hover the listing is asked for, which is what the loop does the moment an
            // archive like this is missed. A file this kind does not hold — one another list
            // reads, one whose bytes are another kind, or one the engine has already turned
            // down — has nothing to ask for and nothing to wait on.
            let Some(_waiting) = request_peazip_render(&path, 0) else {
                println!("request: nothing — the engine is not asked about this file");
                continue;
            };

            let started = Instant::now();
            let mut listing = None;
            while listing.is_none()
                && !peazip_render::refused(&path)
                && started.elapsed() < Duration::from_secs(90)
            {
                std::thread::sleep(Duration::from_millis(250));
                listing = crate::readers::archive_listing::listing_for(&path, None);
            }

            println!(
                "engine: listed = {}, refused = {}, after {:?}",
                listing.is_some(),
                peazip_render::refused(&path),
                started.elapsed()
            );

            if let Some(listing) = &listing {
                println!(
                    "listing: {} entries, {} bytes over them, packed {:?}, file {}",
                    listing.entries.len(),
                    listing.total_size,
                    listing.packed_total,
                    listing.file_size
                );
                for entry in listing.entries.iter().take(4) {
                    println!(
                        "  {} ({} bytes, packed {:?}, dir {})",
                        entry.name, entry.size, entry.packed, entry.is_dir
                    );
                }
            }

            // The replay, which is what the loop does when the listing lands: the hover is
            // measured again — this time from the listing — placed, and loaded.
            let Some(dimensions) = media_dimensions(&path, bounds, dpi) else {
                println!("measure after the listing: nothing — the hover shows no preview");
                continue;
            };
            println!("measure after the listing: {dimensions:?}");

            let Some(layout) = compute_mouse_layout(
                cursor_x,
                cursor_y,
                HoverPlacement {
                    orig_dims: dimensions,
                    avoid: None,
                    follow_cursor: false,
                    preview_scale: scale,
                    at_the_pointer_corner: false,
                },
                bounds,
                dpi,
            ) else {
                println!("layout: none — the hover shows no preview");
                continue;
            };
            println!(
                "layout: {}x{} at ({}, {}), free room {}x{}",
                layout.preview_w,
                layout.preview_h,
                layout.pos_x,
                layout.pos_y,
                layout.max_width,
                layout.max_height
            );

            let cancel = Arc::new(AtomicBool::new(false));
            let started = Instant::now();
            match load_media(
                &path,
                layout.max_width,
                layout.max_height,
                scale,
                dpi,
                Arc::clone(&cancel),
            ) {
                Some(media) => println!(
                    "loaded: {}x{}, {} frame(s), in {:?}",
                    media.current_width(),
                    media.current_height(),
                    media.frames.len(),
                    started.elapsed()
                ),
                None => println!(
                    "loaded: nothing — the hover blinks, in {:?}",
                    started.elapsed()
                ),
            }
        }
    }

    /// The whole path a hover takes for a comic, which is the book kind's fastest path: what the
    /// file's name and its bytes together make of it, which plate the reader chooses out of the
    /// container, the size it is placed at, and the frame it is drawn into. Ignored, and driven by
    /// `RHP_COMIC_PROBE` —
    /// `$env:RHP_COMIC_PROBE = "F:\manga\vol1.cbz;F:\manga\vol2.cbr"; cargo test -- --ignored --nocapture comic_hover_probe`
    /// — for a comic whose preview does not appear, and for working through a folder of them one at
    /// a time.
    ///
    /// Nothing here waits on anything: a comic is read rather than converted, so the whole of what a
    /// hover does is a listing and one plate — and the timings printed are what says so.
    #[test]
    #[ignore = "reads the files named in RHP_COMIC_PROBE"]
    fn comic_hover_probe() {
        let Ok(list) = std::env::var("RHP_COMIC_PROBE") else {
            println!("set RHP_COMIC_PROBE to one or more paths, separated by ';'");
            return;
        };

        let bounds = ScreenBounds {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        let dpi = 96;
        let (cursor_x, cursor_y) = (900, 500);

        for path in list
            .split(';')
            .map(str::trim)
            .filter(|path| !path.is_empty())
            .map(PathBuf::from)
        {
            println!("\n--- {} ---", path.display());

            let (claimed, page_name) = crate::CONFIG
                .lock()
                .map(|config| {
                    (
                        ebook_formats::matches_ebook_list(&path, &config.ebook_extensions),
                        ebook_formats::matches_page_name(&path, &config.ebook_extensions),
                    )
                })
                .unwrap_or((false, false));

            println!(
                "kinds: ebook list = {claimed}, page name = {page_name}, comic name = {}, preview = {}",
                claimed && !page_name,
                ebook_formats::is_ebook_preview(&path)
            );

            let started = Instant::now();
            let Some(plate) = comic_preview::first_page_name(&path) else {
                println!(
                    "plate: none — the container holds no page this app can draw, after {:?}",
                    started.elapsed()
                );
                continue;
            };
            println!("plate: `{plate}` chosen in {:?}", started.elapsed());

            let started = Instant::now();
            let dimensions = comic_preview::dimensions(&path);
            println!("size: {dimensions:?} read in {:?}", started.elapsed());

            let scale = effective_preview_scale(&path, current_hover_scales());
            println!("scale: {scale:?}");

            let Some(measured) = media_dimensions(&path, bounds, dpi) else {
                println!("measure: nothing — the hover shows no preview");
                continue;
            };
            println!("measure: {measured:?}");

            let Some(layout) = compute_mouse_layout(
                cursor_x,
                cursor_y,
                HoverPlacement {
                    orig_dims: measured,
                    avoid: None,
                    follow_cursor: false,
                    preview_scale: scale,
                    at_the_pointer_corner: false,
                },
                bounds,
                dpi,
            ) else {
                println!("layout: none — the hover shows no preview");
                continue;
            };
            println!(
                "layout: {}x{} at ({}, {}), free room {}x{}",
                layout.preview_w,
                layout.preview_h,
                layout.pos_x,
                layout.pos_y,
                layout.max_width,
                layout.max_height
            );

            let cancel = Arc::new(AtomicBool::new(false));
            let started = Instant::now();
            match load_media(
                &path,
                layout.max_width,
                layout.max_height,
                scale,
                dpi,
                Arc::clone(&cancel),
            ) {
                Some(media) => println!(
                    "loaded: {}x{}, {} frame(s), in {:?}",
                    media.current_width(),
                    media.current_height(),
                    media.frames.len(),
                    started.elapsed()
                ),
                None => println!(
                    "loaded: nothing — the hover blinks, in {:?}",
                    started.elapsed()
                ),
            }
        }
    }

    /// The whole path a hover takes for a book an installed Calibre converts, which is the
    /// render engine's path with a slower engine behind it: measured as the wait for a page, the
    /// conversion asked for, the page watched for, and then — when the engine has answered —
    /// measured and loaded from the page itself. Ignored, and driven by `RHP_CALIBRE_PROBE` —
    /// `$env:RHP_CALIBRE_PROBE = "C:\books\book.mobi;C:\books\book.epub"; cargo test -- --ignored --nocapture calibre_hover_probe`
    /// — for a book whose preview does not appear, and for working through a list of the formats
    /// this engine is asked about one at a time.
    ///
    /// It also answers the routing, which is what such a preview is usually about: the name the
    /// configured list carries, what the file's bytes and its name together make of it, whether
    /// the engine is installed to convert it, and whether this hover is the wait for one at all.
    /// That last question is asked here of the loader's own predicate rather than of the list,
    /// because the list is only half of it: a file the engine will be asked about is still a file
    /// with no preview if the load that missed it is not read as a wait (see `awaiting_render` in
    /// the loader worker).
    #[test]
    #[ignore = "reads the files named in RHP_CALIBRE_PROBE and starts the installed Calibre"]
    fn calibre_hover_probe() {
        let Ok(list) = std::env::var("RHP_CALIBRE_PROBE") else {
            println!("set RHP_CALIBRE_PROBE to one or more paths, separated by ';'");
            return;
        };

        let bounds = ScreenBounds {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        let dpi = 96;
        let (cursor_x, cursor_y) = (900, 500);

        // The threads a preview really runs on are in an apartment before they ask for a page
        // (see `spawn_load_worker` and `run_preview_window`), and a probe asks for one from a test
        // thread that is not: a page's size is read through `Windows.Data.Pdf`, so the probe has to
        // put itself in one first, exactly as the two threads do.
        pdf_preview::initialize_apartment();
        wic_image::initialize_apartment();

        for path in list
            .split(';')
            .map(str::trim)
            .filter(|path| !path.is_empty())
        {
            let path = PathBuf::from(path);
            println!("\n--- {} ---", path.display());

            let claimed = crate::CONFIG
                .lock()
                .map(|config| {
                    calibre_formats::matches_calibre_list(&path, &config.calibre_extensions)
                })
                .unwrap_or(false);

            println!(
                "kinds: calibre list = {claimed}, engine ebook = {}, engine available = {}",
                calibre_formats::is_engine_ebook(&path),
                calibre_render::available()
            );
            println!(
                "state: refused = {}, page = {:?}, render due = {}",
                calibre_render::refused(&path),
                calibre_render::rendered_page(&path),
                calibre_render_is_due(&path)
            );

            let scale = effective_preview_scale(&path, current_hover_scales());
            println!("scale: {scale:?}");
            println!(
                "measure before the page: {:?} — the spinner's own box is {}",
                media_dimensions(&path, bounds, dpi),
                office_preview::WAITING_BOX
            );

            // The hover the page is asked for, which is what the loop does the moment a book like
            // this is missed. A file this kind does not hold — one another list reads, one whose
            // bytes are another kind, or one the engine has already turned down — has nothing to
            // ask for and nothing to wait on.
            let Some(_waiting) = request_calibre_render(&path, 0) else {
                println!("request: nothing — the engine is not asked about this file");
                continue;
            };

            let started = Instant::now();
            let mut page = None;
            while page.is_none()
                && !calibre_render::refused(&path)
                && started.elapsed() < Duration::from_secs(300)
            {
                std::thread::sleep(Duration::from_millis(250));
                page = calibre_render::rendered_page(&path);
            }

            println!(
                "engine: page = {:?}, refused = {}, after {:?}",
                page,
                calibre_render::refused(&path),
                started.elapsed()
            );

            if let Some(page) = &page {
                println!(
                    "page: {} bytes, {:?} — previewed from {:?}",
                    std::fs::metadata(page).map(|meta| meta.len()).unwrap_or(0),
                    pdf_preview::page_dimensions(page),
                    pdf_preview::book_page(page)
                );
            }

            // The replay, which is what the loop does when the page lands: the hover is measured
            // again — this time from the page — placed, and loaded.
            let Some(dimensions) = media_dimensions(&path, bounds, dpi) else {
                println!("measure after the page: nothing — the hover shows no preview");
                continue;
            };
            println!("measure after the page: {dimensions:?}");

            let Some(layout) = compute_mouse_layout(
                cursor_x,
                cursor_y,
                HoverPlacement {
                    orig_dims: dimensions,
                    avoid: None,
                    follow_cursor: false,
                    preview_scale: scale,
                    at_the_pointer_corner: false,
                },
                bounds,
                dpi,
            ) else {
                println!("layout: none — the hover shows no preview");
                continue;
            };
            println!(
                "layout: {}x{} at ({}, {}), free room {}x{}",
                layout.preview_w,
                layout.preview_h,
                layout.pos_x,
                layout.pos_y,
                layout.max_width,
                layout.max_height
            );

            let cancel = Arc::new(AtomicBool::new(false));
            let started = Instant::now();
            match load_media(
                &path,
                layout.max_width,
                layout.max_height,
                scale,
                dpi,
                Arc::clone(&cancel),
            ) {
                Some(media) => println!(
                    "loaded: {}x{}, {} frame(s), in {:?}",
                    media.current_width(),
                    media.current_height(),
                    media.frames.len(),
                    started.elapsed()
                ),
                None => println!(
                    "loaded: nothing — the hover blinks, in {:?}",
                    started.elapsed()
                ),
            }
        }
    }

    #[test]
    #[ignore = "builds an animation larger than the retained window of its own"]
    fn sliding_window_probe() {
        use image::codecs::gif::{GifEncoder, Repeat};
        use image::{Delay, Frame, Rgba, RgbaImage};
        use std::io::BufWriter;

        let path = std::env::temp_dir().join("rhp-sliding-window-probe.gif");
        let (side, frame_count, delay_ms) = (1024u32, 40u32, 40u32);

        {
            let file = File::create(&path).expect("a file to write the probe animation to");
            let mut encoder = GifEncoder::new_with_speed(BufWriter::new(file), 30);
            encoder.set_repeat(Repeat::Infinite).expect("looping");
            for index in 0..frame_count {
                let shade = (index * 251 / frame_count) as u8;
                let image = RgbaImage::from_pixel(side, side, Rgba([shade, 40, 255 - shade, 255]));
                encoder
                    .encode_frame(Frame::from_parts(
                        image,
                        0,
                        0,
                        Delay::from_numer_denom_ms(delay_ms, 1),
                    ))
                    .expect("a frame");
            }
        }

        let decoded_mb = u64::from(side) * u64::from(side) * 4 * u64::from(frame_count) / 1048576;
        println!(
            "\n--- {} ({} frames, {} MB decoded) ---",
            path.display(),
            frame_count,
            decoded_mb
        );

        let cancel = Arc::new(AtomicBool::new(false));
        let Some(mut media) = load_animated_gif(
            &path,
            1920,
            1080,
            PreviewScale::Percent(100),
            Arc::clone(&cancel),
        ) else {
            println!("load: None");
            return;
        };

        let start = Instant::now();
        let mut last_frame = media.current_frame;
        let mut advances = 0usize;
        let mut wraps = 0usize;
        let mut next_log = Instant::now();

        while start.elapsed() < Duration::from_secs(12) {
            if media.advance_frame() {
                advances += 1;
                if media.current_frame <= last_frame {
                    wraps += 1;
                }
                last_frame = media.current_frame;
            }

            if Instant::now() >= next_log {
                next_log = Instant::now() + Duration::from_millis(1000);
                println!(
                    "  t={:>5.2}s frame={:<4} held={:<4} loaded={:<5} released={:<5} advances={} wraps={}",
                    start.elapsed().as_secs_f32(),
                    media.current_frame,
                    media.frames.len(),
                    media.is_fully_loaded(),
                    media.frames_were_released(),
                    advances,
                    wraps
                );
            }

            std::thread::sleep(Duration::from_millis(4));
        }

        println!(
            "after 12s: advances={} wraps={} frame={} held={}",
            advances,
            wraps,
            media.current_frame,
            media.frames.len()
        );
        media.cancel_background_work();
        let _ = std::fs::remove_file(&path);
    }
}

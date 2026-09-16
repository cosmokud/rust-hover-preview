use crate::config::{
    sanitize_webp_playback_fps, MarkdownMode, PreviewScale, TextTheme, TransparentBackground,
    DEFAULT_PREVIEW_SCALE_PERCENT, DEFAULT_TEXT_FONT_SCALE_PERCENT,
    DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS, DEFAULT_WEBP_PLAYBACK_FPS,
};
use crate::pdf_preview;
use crate::text_formats;
use crate::text_preview::{self, TextPreviewOptions};
use crate::video_formats::is_video_file;
use crate::wheel_input;
use crate::{CONFIG, RUNNING};
use gif::DecodeOptions;
use image::{AnimationDecoder, GenericImageView};
use once_cell::sync::Lazy;
use std::cell::RefCell;
use std::collections::{HashMap, VecDeque};
use std::env;
use std::fs::{File, OpenOptions};
use std::io::{BufReader, Write};
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
use std::ptr;
use std::sync::atomic::{AtomicBool, AtomicIsize, AtomicU32, Ordering};
use std::sync::mpsc::{channel, Receiver, Sender};
use std::sync::{Arc, Condvar, Mutex};
use std::time::{Duration, Instant};
use windows::core::{w, PCWSTR, PWSTR};
use windows::Win32::Foundation::{
    CloseHandle, GlobalFree, COLORREF, HANDLE, HWND, LPARAM, LRESULT, POINT, RECT, SIZE, WPARAM,
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
    OpenProcess, QueryFullProcessImageNameW, TerminateProcess, PROCESS_NAME_WIN32,
    PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_TERMINATE,
};
use windows::Win32::UI::HiDpi::{GetDpiForMonitor, MDT_EFFECTIVE_DPI};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetAsyncKeyState, ReleaseCapture, SetCapture, VK_C, VK_CONTROL,
};
use windows::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CreatePopupMenu, CreateWindowExW, DefWindowProcW, DestroyMenu, DispatchMessageW,
    EnumWindows, GetSystemMetrics, GetWindow, GetWindowLongPtrW, GetWindowRect,
    GetWindowThreadProcessId, IsWindow, IsWindowVisible, LoadCursorW, MoveWindow, PeekMessageW,
    RegisterClassExW, SetForegroundWindow, SetWindowLongPtrW, SetWindowPos, ShowWindow,
    TrackPopupMenu, TranslateMessage, UpdateLayeredWindow, CS_HREDRAW, CS_VREDRAW, GWL_EXSTYLE,
    GW_OWNER, HWND_TOPMOST, IDC_ARROW, MF_STRING, MSG, PBT_APMRESUMEAUTOMATIC,
    PBT_APMRESUMESUSPEND, PBT_APMSTANDBY, PBT_APMSUSPEND, PM_REMOVE, SM_CXVIRTUALSCREEN,
    SM_CYVIRTUALSCREEN, SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN, SWP_FRAMECHANGED, SWP_NOACTIVATE,
    SWP_NOMOVE, SWP_NOSIZE, SWP_SHOWWINDOW, SW_HIDE, SW_SHOWNOACTIVATE, TPM_LEFTALIGN,
    TPM_NONOTIFY, TPM_RETURNCMD, TPM_TOPALIGN, ULW_ALPHA, WM_DISPLAYCHANGE, WM_DPICHANGED,
    WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MOUSEMOVE, WM_POWERBROADCAST, WM_RBUTTONUP, WNDCLASSEXW,
    WS_EX_LAYERED, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_POPUP,
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
const ANIMATION_STARTUP_PREBUFFER_FRAMES: usize = 12;
const ANIMATION_STARTUP_PREBUFFER_MS: u32 = 500;
const STREAMING_SPINNER_MAX_MS: u64 = 1500;
const VIDEO_GEOMETRY_CACHE_MAX_ENTRIES: usize = 512;
// Expected executable name of the playback process spawned below, used to
// verify a recorded PID still belongs to that process before killing it.
const VIDEO_PROCESS_IMAGE_NAME: &str = "ffplay.exe";

// Message passing for thread communication
pub static PREVIEW_SENDER: Lazy<Mutex<Option<Sender<PreviewMessage>>>> =
    Lazy::new(|| Mutex::new(None));

// Use AtomicIsize for the HWND pointer (thread-safe)
static PREVIEW_HWND: AtomicIsize = AtomicIsize::new(0);

/// Lines one wheel notch moves a text preview. Three is the step a text editor
/// takes, and it keeps a screenful to a few notches.
const TEXT_SCROLL_LINES_PER_NOTCH: i64 = 3;
const WHEEL_DELTA: i32 = 120;

/// How far behind the point a preview was opened from the region reaches, in
/// logical pixels. The pointer travels forwards from there, so this is only there
/// to keep the pixel under a hand at rest inside the region.
const TEXT_SCROLL_ANCHOR_SLACK_PIXELS: i32 = 1;

/// How far either side of the scrollbar's column a press still counts as a press
/// on the bar, in logical pixels. It is deliberately small: the bar is thin, so a
/// hand aiming at it needs some slack, but everything further left is text, and a
/// press in the text is the start of a selection rather than a scroll.
const TEXT_SCROLL_BAR_PRESS_SLACK_PIXELS: f32 = 8.0;

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
///
/// The configured distance is in logical pixels, so it is the same distance under a
/// hand on any display: a margin that is comfortable at 100% is a sliver at 200%,
/// and the scrollbar it is there for scales with the text.
fn far_edge_grace(dpi: u32, configured_pixels: f32) -> i32 {
    (configured_pixels * dpi as f32 / 96.0).round() as i32
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

/// The region that keeps a scrollable text preview on screen, in screen
/// coordinates, or `None` when the preview on screen does not scroll.
static TEXT_SCROLL_KEEP_ALIVE: Lazy<Mutex<Option<ScreenRegion>>> = Lazy::new(|| Mutex::new(None));

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

/// The region that keeps a preview alive: the line from the point it was opened
/// from to the preview, joined to the preview itself.
///
/// A preview is placed beside what it belongs to rather than over it, so the
/// pointer has to travel to reach it — across a gap, sometimes against the side
/// the placement chose. Joining the two means that journey never leaves the
/// region, however the preview ended up placed relative to the cursor, and it
/// keeps the region off everything else: one pixel behind the point the preview
/// was opened from and no slack beyond the preview, apart from `far_edge_grace`
/// past the edge the pointer was travelling towards — so what it covers away from
/// the preview is a single row of the file list.
fn text_scroll_hold_region(
    preview: ScreenRegion,
    anchor: (i32, i32),
    far_edge_grace: i32,
) -> ScreenRegion {
    let (left, top, right, bottom) = preview;

    let (region_left, region_right) = if anchor.0 < left {
        // The preview is to the right of where the pointer was, so the journey is
        // rightwards and the edge it ends at is the right one.
        (
            anchor.0 - TEXT_SCROLL_ANCHOR_SLACK_PIXELS,
            right + far_edge_grace,
        )
    } else if anchor.0 >= right {
        // …or to its left, where the journey ends at the left edge.
        (
            left - far_edge_grace,
            anchor.0 + TEXT_SCROLL_ANCHOR_SLACK_PIXELS,
        )
    } else {
        // The pointer is already in the preview's own column, so there is no edge
        // it travelled towards and the preview is all the region needs to be.
        (left, right)
    };

    (
        region_left,
        top.min(anchor.1),
        region_right,
        bottom.max(anchor.1),
    )
}

// Track the ffplay video window HWND for cursor-over-preview detection
static VIDEO_HWND: AtomicIsize = AtomicIsize::new(0);
// Track the ffplay process ID to re-find the window if needed
static VIDEO_PID: AtomicU32 = AtomicU32::new(0);
// Guard to ensure we only run a single style-monitor thread.
static NOACTIVATE_MONITOR_STARTED: AtomicBool = AtomicBool::new(false);
// Flag set when the system resumes from sleep, so the main loop can reset state.
static RESUME_FROM_SLEEP: AtomicBool = AtomicBool::new(false);

static CURRENT_MEDIA: Lazy<Mutex<Option<MediaData>>> = Lazy::new(|| Mutex::new(None));
static VIDEO_GEOMETRY_CACHE: Lazy<Mutex<HashMap<PathBuf, VideoGeometry>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

#[derive(Clone)]
pub enum PreviewMessage {
    Show(PathBuf, i32, i32),
    ShowKeyboard(PathBuf, i32, i32, i32, i32),
    Hide,
    Refresh,
}

/// Represents different types of media we can display
enum MediaType {
    StaticImage,
    AnimatedGif,
    AnimatedApng,
    AnimatedWebP,
    Video,
    Pdf,
    Text,
    Loading,
}

/// A single frame of image data
struct ImageFrame {
    pixels: Vec<u8>,
    width: u32,
    height: u32,
    delay_ms: u32, // Delay before next frame (for animations)
}

/// Frames an animated preview streams while it plays. The decoder appends to the
/// queue and the player drains it; `released` records that the player gave back
/// frames it already showed, after which the animation can no longer loop from
/// memory and the decoder has to start the file over.
struct StreamedFrames {
    queue: VecDeque<ImageFrame>,
    released: bool,
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
    fn current_pixels(&self) -> &[u8] {
        &self.frames[self.current_frame].pixels
    }

    fn current_width(&self) -> u32 {
        self.frames[self.current_frame].width
    }

    fn current_height(&self) -> u32 {
        self.frames[self.current_frame].height
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
    /// decoding itself into memory faster than it is shown.
    fn sync_shared_frames(&mut self) {
        let Some(shared) = self.shared_frames.clone() else {
            return;
        };

        let retained_bytes: usize = self.frames.iter().map(|frame| frame.pixels.len()).sum();
        if retained_bytes < ANIMATION_RETAINED_BYTES {
            let result = shared.lock();
            if let Ok(mut streamed) = result {
                if !streamed.queue.is_empty() {
                    self.frames.extend(streamed.queue.drain(..));
                }
            }
        }

        self.release_played_frames();
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
    fn release_played_frames(&mut self) {
        let Some(shared) = self.shared_frames.clone() else {
            return;
        };

        let keep_from = self.current_frame.saturating_sub(1);
        if keep_from == 0 {
            return;
        }

        let played_bytes: usize = self.frames[..keep_from]
            .iter()
            .map(|frame| frame.pixels.len())
            .sum();
        if played_bytes < ANIMATION_RELEASE_BYTES {
            return;
        }

        self.frames.drain(..keep_from);
        self.current_frame -= keep_from;

        mark_streamed_frames_released(&shared);
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

        // Safety: if last_frame_time drifted too far behind (e.g. >1s),
        // snap it forward to avoid perpetual catch-up across multiple loops
        if self.last_frame_time.elapsed() > Duration::from_secs(1) {
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
                    self.frames[0].pixels = render_loading_frame(width, height, angle);
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
}

pub fn show_preview(path: &PathBuf, x: i32, y: i32) {
    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(ref tx) = *sender {
            let _ = tx.send(PreviewMessage::Show(path.clone(), x, y));
        }
    }
}

pub fn show_preview_keyboard(
    path: &PathBuf,
    item_left: i32,
    item_top: i32,
    item_right: i32,
    item_bottom: i32,
) {
    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(ref tx) = *sender {
            let _ = tx.send(PreviewMessage::ShowKeyboard(
                path.clone(),
                item_left,
                item_top,
                item_right,
                item_bottom,
            ));
        }
    }
}

pub fn hide_preview() {
    unsafe {
        let hwnd = HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _);
        if !hwnd.is_invalid() {
            let _ = ShowWindow(hwnd, SW_HIDE);
        }
    }

    if let Ok(mut current) = CURRENT_MEDIA.try_lock() {
        if let Some(ref mut media) = *current {
            media.cancel_background_work();
            stop_video_playback(media);
        }
        *current = None;
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

/// Which preview surface the pointer is currently on.
#[derive(Clone, Copy)]
pub struct PreviewCursorHover {
    pub image: bool,
    pub video: bool,
}

impl PreviewCursorHover {
    pub const NONE: Self = Self {
        image: false,
        video: false,
    };

    pub fn any(self) -> bool {
        self.image || self.video
    }
}

/// Single shared pointer probe for both preview kinds. Callers gate it on
/// "a preview can be under the pointer"; the fast path keeps the cost at a few
/// atomic reads whenever nothing is on screen.
pub fn cursor_preview_hover() -> PreviewCursorHover {
    let preview_hwnd = PREVIEW_HWND.load(Ordering::SeqCst);
    let video_hwnd = VIDEO_HWND.load(Ordering::SeqCst);
    let video_pid = VIDEO_PID.load(Ordering::SeqCst);

    // PREVIEW_HWND is created once at startup and never cleared, so visibility
    // is what tells us whether the layered window is actually on screen.
    let preview_visible =
        preview_hwnd != 0 && unsafe { IsWindowVisible(HWND(preview_hwnd as *mut _)).as_bool() };
    if !preview_visible && video_hwnd == 0 && video_pid == 0 {
        return PreviewCursorHover::NONE;
    }

    unsafe {
        use windows::Win32::UI::WindowsAndMessaging::{GetCursorPos, WindowFromPoint};

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

        // A hit on the stored HWND is enough; the process-ID fallback covers the
        // race window where ffplay's window exists but VIDEO_HWND isn't stored yet.
        let mut video = video_hwnd != 0 && hwnd_ptr == video_hwnd;
        if !video && video_pid != 0 {
            use windows::Win32::UI::WindowsAndMessaging::GetWindowThreadProcessId;

            let mut window_pid: u32 = 0;
            GetWindowThreadProcessId(hwnd_under_cursor, Some(&mut window_pid));
            video = window_pid == video_pid;
        }

        PreviewCursorHover { image, video }
    }
}

/// Screen-space box of the preview surface that is on screen right now, if any.
/// The Explorer hook uses it to decide whether a keyboard preview was placed
/// over the parked pointer, so a pointer sitting under the preview cannot drive
/// previews or dismiss them.
pub fn preview_screen_rect() -> Option<(i32, i32, i32, i32)> {
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

fn is_gif_file(path: &PathBuf) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_lowercase() == "gif")
        .unwrap_or(false)
}

fn is_webp_file(path: &PathBuf) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_lowercase() == "webp")
        .unwrap_or(false)
}

/// True when a `.png` file carries an animation control chunk. An APNG is an
/// ordinary PNG plus an `acTL` chunk ahead of its first `IDAT`, so the chunk list
/// is walked instead of decoding anything.
fn png_has_animation_control_chunk(path: &PathBuf) -> bool {
    use std::io::{Read, Seek, SeekFrom};

    const SIGNATURE: [u8; 8] = [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];

    let Ok(file) = File::open(path) else {
        return false;
    };
    let mut reader = BufReader::new(file);

    let mut signature = [0u8; 8];
    if reader.read_exact(&mut signature).is_err() || signature != SIGNATURE {
        return false;
    }

    let mut header = [0u8; 8];
    while reader.read_exact(&mut header).is_ok() {
        let length = u32::from_be_bytes([header[0], header[1], header[2], header[3]]) as i64;
        let chunk_type = &header[4..8];

        if chunk_type == b"acTL" {
            return true;
        }

        // The animation chunks precede the image data, so once the pixels start
        // there is nothing left to find.
        if chunk_type == b"IDAT" || chunk_type == b"IEND" {
            return false;
        }

        // Step over the chunk body and its CRC.
        if reader.seek(SeekFrom::Current(length + 4)).is_err() {
            return false;
        }
    }

    false
}

fn is_apng_file(path: &PathBuf) -> bool {
    match path
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_lowercase())
        .as_deref()
    {
        Some("apng") => true,
        // An animated PNG keeps the `.png` extension whenever it was written by
        // an ordinary PNG encoder, so those are decided by content.
        Some("png") => png_has_animation_control_chunk(path),
        _ => false,
    }
}

fn is_confirm_file_type_enabled() -> bool {
    CONFIG
        .lock()
        .map(|cfg| cfg.confirm_file_type)
        .unwrap_or(false)
}

/// Guess image format from header bytes instead of file extension.
fn guessed_image_format(path: &PathBuf) -> Option<image::ImageFormat> {
    image::ImageReader::open(path)
        .ok()?
        .with_guessed_format()
        .ok()?
        .format()
}

/// Decode an image by sniffing magic bytes instead of trusting the extension.
fn decode_image_with_header_check(path: &PathBuf) -> Option<image::DynamicImage> {
    image::ImageReader::open(path)
        .ok()?
        .with_guessed_format()
        .ok()?
        .decode()
        .ok()
}

/// Read image dimensions by sniffing magic bytes instead of trusting the extension.
fn image_dimensions_with_header_check(path: &PathBuf) -> Option<(u32, u32)> {
    image::ImageReader::open(path)
        .ok()?
        .with_guessed_format()
        .ok()?
        .into_dimensions()
        .ok()
}

/// Convert RGBA pixels to BGRA for Windows GDI
fn rgba_to_bgra(rgba: &[u8]) -> Vec<u8> {
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

fn current_transparent_background() -> TransparentBackground {
    CONFIG
        .lock()
        .map(|cfg| cfg.transparent_background)
        .unwrap_or(TransparentBackground::Transparent)
}

fn current_webp_playback_fps() -> u32 {
    CONFIG
        .lock()
        .map(|cfg| sanitize_webp_playback_fps(cfg.webp_playback_fps))
        .unwrap_or(DEFAULT_WEBP_PLAYBACK_FPS)
}

fn current_preview_scale() -> PreviewScale {
    CONFIG
        .lock()
        .map(|cfg| cfg.preview_scale)
        .unwrap_or(PreviewScale::Percent(DEFAULT_PREVIEW_SCALE_PERCENT))
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

/// Whether the preview on screen is a text preview, whose appearance is baked
/// into its painted frame rather than recomposited from shared pixels.
fn current_media_is_text() -> bool {
    CURRENT_MEDIA
        .lock()
        .map(|media| {
            matches!(
                media.as_ref().map(|media| &media.media_type),
                Some(MediaType::Text)
            )
        })
        .unwrap_or(false)
}

/// The scale a preview is laid out and rendered with.
///
/// A PDF page is a vector, so the engine draws it at whatever size it is asked
/// for and a larger preview is sharper text rather than an enlarged raster. The
/// configured scale could only hold that back, so a PDF always takes the space
/// the display allows.
///
/// Text is the opposite case: it is drawn at a fixed, display-scaled font size,
/// so enlarging it would only stretch the window around text that stays the same
/// size. `100%` is exactly the rule text wants — never enlarged, reduced only
/// when the space beside the cursor cannot hold it — and the text renderer reads
/// the size it is given as "as many lines and columns as fit".
///
/// Every other format keeps the configured scale.
fn effective_preview_scale(path: &Path, preview_scale: PreviewScale) -> PreviewScale {
    if pdf_preview::is_pdf_file(path) {
        PreviewScale::FitToScreen
    } else if text_formats::is_text_file(path) {
        PreviewScale::Percent(100)
    } else {
        preview_scale
    }
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
    if ((x / 16) + (y / 16)) % 2 == 0 {
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
fn compose_preview_pixels_into(
    bgra: &[u8],
    width: u32,
    height: u32,
    background: TransparentBackground,
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

    let row_bytes = width * 4;
    for (y, (src_row, dst_row)) in bgra
        .chunks_exact(row_bytes)
        .zip(out[..expected].chunks_exact_mut(row_bytes))
        .enumerate()
    {
        compose_preview_row(src_row, dst_row, background, y as u32);
    }
}

fn compose_preview_row(
    src_row: &[u8],
    dst_row: &mut [u8],
    background: TransparentBackground,
    y: u32,
) {
    match background {
        TransparentBackground::Transparent => {
            for (px, dst) in src_row.chunks_exact(4).zip(dst_row.chunks_exact_mut(4)) {
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
        TransparentBackground::Black
        | TransparentBackground::White
        | TransparentBackground::Checkerboard => {
            for (x, (px, dst)) in src_row
                .chunks_exact(4)
                .zip(dst_row.chunks_exact_mut(4))
                .enumerate()
            {
                let b = px[0] as u32;
                let g = px[1] as u32;
                let r = px[2] as u32;
                let a = px[3] as u32;

                let (bg_b, bg_g, bg_r) = match background {
                    TransparentBackground::Black => (0u32, 0u32, 0u32),
                    TransparentBackground::White => (255u32, 255u32, 255u32),
                    TransparentBackground::Checkerboard => {
                        let (cr, cg, cb) = checkerboard_color(x as u32, y);
                        (cb as u32, cg as u32, cr as u32)
                    }
                    TransparentBackground::Transparent => unreachable!(),
                };
                let inv_a = 255 - a;

                dst[0] = ((b * a + bg_b * inv_a + 127) / 255) as u8;
                dst[1] = ((g * a + bg_g * inv_a + 127) / 255) as u8;
                dst[2] = ((r * a + bg_r * inv_a + 127) / 255) as u8;
                dst[3] = 255;
            }
        }
    }
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
    let fit_scale =
        (max_width as f32 / orig_width as f32).min(max_height as f32 / orig_height as f32);

    // A requested percentage is honored when it fits; anything larger than the
    // available area falls back to the fit scale so nothing is ever clipped.
    let scale = match preview_scale.target_scale() {
        Some(target_scale) => target_scale.min(fit_scale),
        None => fit_scale,
    };

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

    Some(ImageFrame {
        pixels: bgra,
        width: target_width,
        height: target_height,
        delay_ms,
    })
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

/// Whether the player has released the frames it already showed, which means the
/// file has to be decoded again to play the animation another time.
fn streamed_frames_released(shared: &Arc<Mutex<StreamedFrames>>) -> bool {
    shared
        .lock()
        .map(|streamed| streamed.released)
        .unwrap_or(false)
}

/// Records that the player gave back frames it already showed.
fn mark_streamed_frames_released(shared: &Arc<Mutex<StreamedFrames>>) {
    let result = shared.lock();
    if let Ok(mut streamed) = result {
        streamed.released = true;
    }
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

    let mut canvas = vec![0u8; (gif_width * gif_height * 4) as usize];
    let mut initial_frames = Vec::new();
    let mut initial_bytes: usize = 0;
    let mut buffered_ms: u32 = 0;
    let mut reached_end = false;

    while initial_frames.len() < ANIMATION_STARTUP_PREBUFFER_FRAMES
        && (initial_frames.len() < 2 || buffered_ms < ANIMATION_STARTUP_PREBUFFER_MS)
    {
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
        buffered_ms = buffered_ms.saturating_add(delay_ms);
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

            // The player gave back the frames it already showed, so the file is
            // decoded again to play the animation another time.
            if cancelled || !streamed_frames_released(&shared_clone) {
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
    let decoder = image::codecs::png::PngDecoder::new(BufReader::new(file)).ok()?;
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

    ImageFrame {
        pixels: rgba_to_bgra(&rgba),
        width: target_width,
        height: target_height,
        delay_ms,
    }
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
    let mut buffered_ms: u32 = 0;
    let mut reached_end = false;
    let mut target_size: Option<(u32, u32)> = None;

    while initial_frames.len() < ANIMATION_STARTUP_PREBUFFER_FRAMES
        && (initial_frames.len() < 2 || buffered_ms < ANIMATION_STARTUP_PREBUFFER_MS)
    {
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
        buffered_ms = buffered_ms.saturating_add(delay_ms);
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

            // The player gave back the frames it already showed, so the file is
            // decoded again to play the animation another time.
            if cancelled || !streamed_frames_released(&shared_clone) {
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
        for chunk in bgra.chunks_exact(4) {
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

    Some(ImageFrame {
        pixels,
        width: target_width,
        height: target_height,
        delay_ms,
    })
}

fn load_animated_webp(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    let buffer = Arc::new(std::fs::read(path).ok()?);
    let options = webp_animation::DecoderOptions {
        use_threads: true,
        color_mode: webp_animation::ColorMode::Bgra,
    };
    let decoder = webp_animation::Decoder::new_with_options(buffer.as_slice(), options).ok()?;

    let (orig_width, orig_height) = decoder.dimensions();
    if orig_width == 0 || orig_height == 0 || orig_width > 16384 || orig_height > 16384 {
        return None;
    }

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
    let mut buffered_ms: u32 = 0;
    let mut previous_timestamp = 0i32;
    let mut reached_end = false;
    let mut iterator = decoder.into_iter();

    while initial_frames.len() < ANIMATION_STARTUP_PREBUFFER_FRAMES
        && (initial_frames.len() < 2 || buffered_ms < ANIMATION_STARTUP_PREBUFFER_MS)
    {
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
        buffered_ms = buffered_ms.saturating_add(delay_ms);
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

            // The player gave back the frames it already showed, so the file is
            // decoded again to play the animation another time.
            if cancelled || !streamed_frames_released(&shared_clone) {
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

/// Load a static image (JPG, PNG, BMP, static WebP, etc.)
fn load_static_image(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    let img = if is_confirm_file_type_enabled() {
        decode_image_with_header_check(path)?
    } else {
        image::open(path).ok()?
    };
    let (orig_width, orig_height) = img.dimensions();
    let (target_width, target_height) = scale_dimensions(
        orig_width,
        orig_height,
        max_width,
        max_height,
        preview_scale,
    );

    let resized = if target_width != orig_width || target_height != orig_height {
        img.resize_exact(
            target_width,
            target_height,
            image::imageops::FilterType::Triangle,
        )
    } else {
        img
    };

    let rgba = resized.to_rgba8();
    let bgra = rgba_to_bgra(rgba.as_raw());

    let frame = ImageFrame {
        pixels: bgra,
        width: target_width,
        height: target_height,
        delay_ms: 0,
    };

    Some(MediaData {
        frames: vec![frame],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::StaticImage,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
    })
}

/// Render the first page of a PDF through the PDF engine built into Windows.
fn load_pdf_first_page(
    path: &PathBuf,
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

    let frame = ImageFrame {
        pixels,
        width,
        height,
        delay_ms: 0,
    };

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

    let frame = ImageFrame {
        pixels: frame.pixels,
        width: frame.width,
        height: frame.height,
        delay_ms: 0,
    };

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

/// Extract video thumbnail using ffmpeg and create frames for preview
fn load_video_thumbnail(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
) -> Option<MediaData> {
    let geometry = get_video_geometry(path).unwrap_or(VideoGeometry {
        width: 1920,
        height: 1080,
        crop: None,
    });
    let (target_width, target_height) = scale_dimensions(
        geometry.width,
        geometry.height,
        max_width,
        max_height,
        preview_scale,
    );

    // Create a placeholder frame (dark gray) while video plays
    let placeholder_pixels = vec![40u8; (target_width * target_height * 4) as usize];

    let frame = ImageFrame {
        pixels: placeholder_pixels,
        width: target_width,
        height: target_height,
        delay_ms: 0,
    };

    Some(MediaData {
        frames: vec![frame],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::Video,
        stream_cancel: None,
        video_process: None,
        loading_start: None,
        text_state: None,
    })
}

// Windows constant for hiding console window
const CREATE_NO_WINDOW: u32 = 0x08000000;
const VIDEO_CROPDETECT_LIMIT: &str = "24";
const VIDEO_CROPDETECT_ROUND: &str = "16";
const VIDEO_CROPDETECT_FRAMES: &str = "48";
const VIDEO_CROP_MAX_AXIS_TRIM_RATIO: f32 = 0.10;
const VIDEO_CROP_MAX_ASYMMETRY_PX: i32 = 12;

/// Get video dimensions using ffprobe
fn get_video_dimensions(path: &PathBuf) -> Option<(u32, u32)> {
    let output = Command::new("ffprobe")
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
        .creation_flags(CREATE_NO_WINDOW) // Hide the console window
        .output()
        .ok()?;

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

fn detect_video_crop(path: &PathBuf, src_w: u32, src_h: u32) -> Option<VideoCrop> {
    let filter = format!(
        "cropdetect={}:{}:0",
        VIDEO_CROPDETECT_LIMIT, VIDEO_CROPDETECT_ROUND
    );

    let output = Command::new("ffmpeg")
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
        .creation_flags(CREATE_NO_WINDOW)
        .output()
        .ok()?;

    let stderr = String::from_utf8_lossy(&output.stderr);
    let mut counts: HashMap<(u32, u32, u32, u32), u32> = HashMap::new();
    for line in stderr.lines() {
        if let Some(crop) = parse_cropdetect_line(line) {
            *counts
                .entry((crop.width, crop.height, crop.x, crop.y))
                .or_insert(0) += 1;
        }
    }

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

fn get_video_geometry(path: &PathBuf) -> Option<VideoGeometry> {
    if let Ok(cache) = VIDEO_GEOMETRY_CACHE.lock() {
        if let Some(cached) = cache.get(path) {
            return Some(*cached);
        }
    }

    let (src_w, src_h) = get_video_dimensions(path)?;
    let crop = detect_video_crop(path, src_w, src_h);

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

    if let Ok(mut cache) = VIDEO_GEOMETRY_CACHE.lock() {
        if !cache.contains_key(path) && cache.len() >= VIDEO_GEOMETRY_CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.insert(path.clone(), geometry);
    }

    Some(geometry)
}

fn log_video_preview(
    path: &PathBuf,
    x: i32,
    y: i32,
    width: i32,
    height: i32,
    vf: Option<&str>,
    geometry: Option<VideoGeometry>,
) {
    let log_path = env::temp_dir().join("rust-hover-preview-video.log");
    let Ok(mut file) = OpenOptions::new().create(true).append(true).open(log_path) else {
        return;
    };

    let crop = geometry
        .and_then(|g| g.crop)
        .map(|c| format!("{}:{}:{}:{}", c.width, c.height, c.x, c.y))
        .unwrap_or_else(|| "none".to_string());
    let geom = geometry
        .map(|g| format!("{}x{}", g.width, g.height))
        .unwrap_or_else(|| "none".to_string());
    let vf = vf.unwrap_or("none");

    let _ = writeln!(
        file,
        "path=\"{}\" pos={}x{} window={}x{} geometry={} crop={} vf=\"{}\"",
        path.display(),
        x,
        y,
        width,
        height,
        geom,
        crop,
        vf
    );
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
                std::thread::sleep(Duration::from_millis(delay_ms));
            } else {
                std::thread::sleep(Duration::from_millis(80));
            }
        }

        NOACTIVATE_MONITOR_STARTED.store(false, Ordering::Release);
    });
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

    let geometry = get_video_geometry(path);
    let vf = geometry.map(|geometry| {
        if let Some(crop) = geometry.crop {
            format!(
                "crop={}:{}:{}:{},setsar=1",
                crop.width, crop.height, crop.x, crop.y
            )
        } else {
            "setsar=1".to_string()
        }
    });
    log_video_preview(path, x, y, width, height, vf.as_deref(), geometry);
    if let Some(vf) = vf.as_deref() {
        cmd.args(["-vf", &vf]);
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
        .creation_flags(CREATE_NO_WINDOW) // Hide the console window
        .spawn()
        .ok();

    // After spawning, try to set WS_EX_NOACTIVATE on the ffplay window
    // to prevent it from stealing focus
    if let Some(ref child_process) = child {
        set_noactivate_for_process(child_process.id());
    }

    child
}

/// Stop video playback process
fn stop_video_playback(media: &mut MediaData) {
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
                        media.video_process = None;
                        VIDEO_HWND.store(0, Ordering::SeqCst);
                        VIDEO_PID.store(0, Ordering::SeqCst);
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

    if is_video_file(path) {
        return load_video_thumbnail(path, max_width, max_height, preview_scale);
    }

    if pdf_preview::is_pdf_file(path) {
        return load_pdf_first_page(path, max_width, max_height, preview_scale);
    }

    if text_formats::is_text_file(path) {
        return load_text_preview(path, max_width, max_height, dpi, current_text_options());
    }

    let guessed_format = if is_confirm_file_type_enabled() {
        guessed_image_format(path)
    } else {
        None
    };

    if matches!(guessed_format, Some(image::ImageFormat::Gif)) || is_gif_file(path) {
        // Try animated GIF first
        if let Some(media) = load_animated_gif(
            path,
            max_width,
            max_height,
            preview_scale,
            Arc::clone(&cancel),
        ) {
            return Some(media);
        }
        if cancel.load(Ordering::Acquire) {
            return None;
        }
        // Fall back to static for single-frame GIFs
        return load_static_image(path, max_width, max_height, preview_scale);
    }

    if matches!(guessed_format, Some(image::ImageFormat::WebP)) || is_webp_file(path) {
        // Try animated WebP first
        if let Some(media) = load_animated_webp(
            path,
            max_width,
            max_height,
            preview_scale,
            Arc::clone(&cancel),
        ) {
            return Some(media);
        }
        if cancel.load(Ordering::Acquire) {
            return None;
        }
        // Fall back to static for non-animated WebP
        return load_static_image(path, max_width, max_height, preview_scale);
    }

    if is_apng_file(path) {
        // Try animated APNG first
        if let Some(media) = load_animated_apng(
            path,
            max_width,
            max_height,
            preview_scale,
            Arc::clone(&cancel),
        ) {
            return Some(media);
        }
        if cancel.load(Ordering::Acquire) {
            return None;
        }
        // Fall back to static for single-frame APNGs
        return load_static_image(path, max_width, max_height, preview_scale);
    }

    // Default to static image
    if cancel.load(Ordering::Acquire) {
        return None;
    }
    load_static_image(path, max_width, max_height, preview_scale)
}

/// Get original dimensions of media for positioning calculations
fn get_media_dimensions(path: &PathBuf) -> Option<(u32, u32)> {
    if is_video_file(path) {
        return get_video_geometry(path)
            .map(|g| (g.width, g.height))
            .or(Some((1920, 1080)));
    }

    // A PDF is measured from its own first page; one that cannot be read as a
    // PDF reports no dimensions, which drops the preview instead of guessing.
    if pdf_preview::is_pdf_file(path) {
        return pdf_preview::page_dimensions(path);
    }

    if is_confirm_file_type_enabled() {
        image_dimensions_with_header_check(path)
    } else {
        image::image_dimensions(path).ok()
    }
}

/// The size the layout should place and scale a preview from.
///
/// A text file has no size of its own, so the box its first screenful wants is
/// measured here, bounded by the display it will be shown on. That makes the
/// measurement an intrinsic size in the same sense a PDF page's is: the layout
/// can fit it into the space beside the cursor, and the text renderer is handed
/// the box that comes out of that.
fn media_dimensions(path: &PathBuf, bounds: ScreenBounds, dpi: u32) -> Option<(u32, u32)> {
    if text_formats::is_text_file(path) {
        let cap_width = (bounds.right - bounds.left).max(1) as u32;
        let cap_height = bounds.height().max(1) as u32;
        return text_preview::measure(path, cap_width, cap_height, dpi, current_text_options());
    }

    get_media_dimensions(path)
}

/// Effective DPI of the display nearest `(x, y)`, which is what a text preview's
/// font size is scaled by. Falls back to the 96 DPI baseline when the monitor
/// query fails, the same way the placement falls back to the virtual screen.
fn monitor_dpi_from_point(x: i32, y: i32) -> u32 {
    unsafe {
        let monitor = MonitorFromPoint(POINT { x, y }, MONITOR_DEFAULTTONEAREST);
        if !monitor.is_invalid() {
            let mut dpi_x = 0u32;
            let mut dpi_y = 0u32;
            if GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, &mut dpi_x, &mut dpi_y).is_ok()
                && dpi_x > 0
            {
                return dpi_x;
            }
        }
    }

    96
}

/// Render a single frame of the loading spinner animation (BGRA pixels)
fn render_loading_frame(width: u32, height: u32, angle: f32) -> Vec<u8> {
    let total_pixels = (width as usize) * (height as usize);
    let mut pixels = vec![0u8; total_pixels * 4];

    let cx = width as f32 / 2.0;
    let cy = height as f32 / 2.0;

    // Spinner proportional to window size, clamped for aesthetics
    let radius = (width.min(height) as f32 * 0.08).clamp(10.0, 32.0);
    let thickness = (radius * 0.32).clamp(2.5, 7.0);

    // Background color (dark charcoal)
    let bg: [u8; 3] = [30, 30, 30];

    // Fill background
    for pixel in pixels.chunks_exact_mut(4) {
        pixel[0] = bg[0]; // B
        pixel[1] = bg[1]; // G
        pixel[2] = bg[2]; // R
        pixel[3] = 255; // A
    }

    let two_pi = std::f32::consts::PI * 2.0;
    let arc_length = std::f32::consts::PI * 1.5; // 270-degree arc

    // Only iterate over the bounding box of the spinner ring
    let min_x = ((cx - radius - thickness - 2.0).max(0.0)) as u32;
    let max_x = ((cx + radius + thickness + 2.0).min(width as f32 - 1.0)) as u32;
    let min_y = ((cy - radius - thickness - 2.0).max(0.0)) as u32;
    let max_y = ((cy + radius + thickness + 2.0).min(height as f32 - 1.0)) as u32;

    for y in min_y..=max_y {
        for x in min_x..=max_x {
            let dx = x as f32 + 0.5 - cx;
            let dy = y as f32 + 0.5 - cy;
            let dist = (dx * dx + dy * dy).sqrt();

            let ring_dist = (dist - radius).abs();
            if ring_dist > thickness + 1.0 {
                continue;
            }

            // Anti-aliased smooth edge
            let edge_alpha = (1.0 - (ring_dist - thickness + 1.0).max(0.0)).clamp(0.0, 1.0);
            if edge_alpha <= 0.0 {
                continue;
            }

            let pixel_angle = dy.atan2(dx);
            let relative = (pixel_angle - angle).rem_euclid(two_pi);

            if relative <= arc_length {
                // Smooth gradient: ease-in from tail (transparent) to head (bright)
                let t = relative / arc_length;
                let t_smooth = t * t; // quadratic ease-in
                let alpha = edge_alpha * t_smooth;

                let idx = ((y * width + x) * 4) as usize;
                let blend = |bg_c: u8, fg: u8, a: f32| -> u8 {
                    ((bg_c as f32) * (1.0 - a) + (fg as f32) * a).clamp(0.0, 255.0) as u8
                };

                pixels[idx] = blend(bg[0], 255, alpha); // B
                pixels[idx + 1] = blend(bg[1], 255, alpha); // G
                pixels[idx + 2] = blend(bg[2], 255, alpha); // R
                pixels[idx + 3] = 255;
            }
        }
    }

    pixels
}

/// Create a loading animation MediaData for the given dimensions
fn create_loading_media(width: u32, height: u32) -> MediaData {
    let pixels = render_loading_frame(width, height, 0.0);
    let frame = ImageFrame {
        pixels,
        width,
        height,
        delay_ms: 33,
    };
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

/// Render a small loading spinner overlay onto an existing BGRA pixel buffer (in-place).
/// Draws a spinning arc in the bottom-right corner with a semi-transparent dark backdrop circle.
fn overlay_loading_spinner(pixels: &mut [u8], width: u32, height: u32, angle: f32) {
    if width < 24 || height < 24 {
        return;
    }

    let radius = 8.0_f32;
    let thickness = 2.5_f32;
    let padding = 12.0_f32;
    let backdrop_r = radius + thickness + 4.0;

    // Center of the spinner in the bottom-right corner
    let cx = width as f32 - padding - radius - thickness;
    let cy = height as f32 - padding - radius - thickness;

    let min_x = ((cx - backdrop_r - 1.0).max(0.0)) as u32;
    let max_x = ((cx + backdrop_r + 1.0).min(width as f32 - 1.0)) as u32;
    let min_y = ((cy - backdrop_r - 1.0).max(0.0)) as u32;
    let max_y = ((cy + backdrop_r + 1.0).min(height as f32 - 1.0)) as u32;

    let two_pi = std::f32::consts::PI * 2.0;
    let arc_length = std::f32::consts::PI * 1.5;

    for y in min_y..=max_y {
        for x in min_x..=max_x {
            let dx = x as f32 + 0.5 - cx;
            let dy = y as f32 + 0.5 - cy;
            let dist = (dx * dx + dy * dy).sqrt();
            let idx = ((y * width + x) * 4) as usize;
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
    media: Option<MediaData>,
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
        // Rendering a PDF page goes through Windows.Data.Pdf on this thread.
        pdf_preview::initialize_apartment();

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

            let _ = result_tx.send(LoadResult {
                generation: request.generation,
                media,
            });
        }
    })
}

/// Tracks a pending background load so we can show the spinner after a delay
struct PendingLoad {
    generation: u64,
    started: Instant,
    pos_x: i32,
    pos_y: i32,
    width: u32,
    height: u32,
    spinner_shown: bool,
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
    /// Mutable frame copy for the loading spinner, which is overlaid before the
    /// frame is composed.
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
    let Some((width, height)) = (|| {
        let media_guard = CURRENT_MEDIA.lock().ok()?;
        let media = media_guard.as_ref()?;

        if matches!(media.media_type, MediaType::Video) {
            return None;
        }

        let width = media.current_width();
        let height = media.current_height();
        let expected_size = width as usize * height as usize * 4;
        if width == 0 || height == 0 || media.current_pixels().len() < expected_size {
            return None;
        }

        let background = current_transparent_background();
        let bits = ensure_layered_surface(width, height)?;
        let out = unsafe { std::slice::from_raw_parts_mut(bits, expected_size) };

        if media.should_draw_streaming_overlay() {
            let elapsed = media
                .loading_start
                .map(|s| s.elapsed().as_secs_f32())
                .unwrap_or(0.0);
            let angle = elapsed * 2.0 * std::f32::consts::PI * 1.2;
            OVERLAY_SCRATCH.with(|cell| {
                let mut buf = cell.borrow_mut();
                buf.clear();
                buf.extend_from_slice(&media.current_pixels()[..expected_size]);
                overlay_loading_spinner(&mut buf, width, height, angle);
                compose_preview_pixels_into(&buf, width, height, background, out);
            });
        } else {
            compose_preview_pixels_into(media.current_pixels(), width, height, background, out);
        }

        Some((width, height))
    })() else {
        return;
    };

    let mut rect = RECT::default();
    if GetWindowRect(hwnd, &mut rect).is_err() {
        return;
    }

    let Some(mem_dc) =
        LAYERED_SURFACE.with(|cell| cell.borrow().as_ref().map(|surface| surface.mem_dc))
    else {
        return;
    };

    let dst_point = POINT {
        x: rect.left,
        y: rect.top,
    };
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

    publish_text_scroll_keep_alive(hwnd);
}

/// Publish — or withdraw — the region in which the pointer keeps a scrollable
/// text preview alive.
///
/// The Explorer hook polls this to decide whether the pointer over the preview
/// means "the user is reading this" or "dismiss it and show what is underneath",
/// and the wheel hook asks the same question before it decides whether the wheel
/// belongs to Explorer or to the preview.
unsafe fn publish_text_scroll_keep_alive(hwnd: HWND) {
    let state = CURRENT_MEDIA.lock().ok().and_then(|media| {
        let media = media.as_ref()?;
        if !matches!(media.media_type, MediaType::Text) {
            return None;
        }
        let state = media.text_state.as_ref()?;
        Some((state.dpi, state.can_scroll()))
    });

    // The wheel only belongs to a preview that can move under it, but the pointer
    // is held by any text preview in full mode: selecting and copying needs a
    // pointer that can rest on the preview whether or not it scrolls.
    TEXT_PREVIEW_SCROLLABLE.store(
        state.map(|(_, can_scroll)| can_scroll).unwrap_or(false),
        Ordering::Release,
    );
    TEXT_PREVIEW_HOLDING.store(state.is_some(), Ordering::Release);

    let keep_alive = state.and_then(|(dpi, _)| {
        let mut rect = RECT::default();
        if GetWindowRect(hwnd, &mut rect).is_err() {
            return None;
        }

        let preview = (rect.left, rect.top, rect.right, rect.bottom);
        let anchor = TEXT_SCROLL_ANCHOR
            .lock()
            .ok()
            .and_then(|anchor| *anchor)
            // Without an anchor — a preview that was already on screen when this
            // started, say — the preview's own corner stands in for it, which
            // leaves the region as the preview alone.
            .unwrap_or((preview.0, preview.1));

        Some(text_scroll_hold_region(
            preview,
            anchor,
            far_edge_grace(dpi, configured_far_edge_grace_pixels()),
        ))
    });

    if let Ok(mut published) = TEXT_SCROLL_KEEP_ALIVE.lock() {
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

fn clear_text_scroll_keep_alive() {
    TEXT_PREVIEW_SCROLLABLE.store(false, Ordering::Release);
    TEXT_PREVIEW_HOLDING.store(false, Ordering::Release);
    if let Ok(mut published) = TEXT_SCROLL_KEEP_ALIVE.lock() {
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

/// Whether the pointer is inside the region that keeps a text preview alive.
/// Answered from a published rectangle, so the Explorer hook can ask on every poll
/// tick.
pub fn text_scroll_pointer_hold(x: i32, y: i32) -> bool {
    if !TEXT_PREVIEW_HOLDING.load(Ordering::Acquire) {
        return false;
    }

    let Ok(published) = TEXT_SCROLL_KEEP_ALIVE.lock() else {
        return false;
    };

    published
        .map(|(left, top, right, bottom)| x >= left && x < right && y >= top && y < bottom)
        .unwrap_or(false)
}

/// The published region, without blocking. The wheel hook runs inside a
/// system-wide hook procedure, where waiting on a lock held by the preview thread
/// would stall every wheel message on the desktop — so a lock it cannot take
/// immediately means the wheel is not ours to take either.
pub fn text_scroll_keep_alive_try() -> Option<(i32, i32, i32, i32)> {
    if !text_preview_scrollable() {
        return None;
    }

    TEXT_SCROLL_KEEP_ALIVE
        .try_lock()
        .ok()
        .and_then(|held| *held)
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

        media.frames[0] = ImageFrame {
            pixels: frame.pixels,
            width: frame.width,
            height: frame.height,
            delay_ms: 0,
        };

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

        media.frames[0] = ImageFrame {
            pixels: frame.pixels,
            width: frame.width,
            height: frame.height,
            delay_ms: 0,
        };

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

/// The one thing the preview's own menu does: put what is selected on the
/// clipboard.
const ID_TEXT_PREVIEW_COPY: usize = 1;

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

    if command.0 as usize == ID_TEXT_PREVIEW_COPY {
        copy_text_preview(hwnd);
    }
}

unsafe fn reset_preview_after_display_change(hwnd: HWND) {
    let _ = ShowWindow(hwnd, SW_HIDE);
    clear_text_scroll_keep_alive();

    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        if let Some(ref mut media) = *current {
            media.cancel_background_work();
            stop_video_playback(media);
        }
        *current = None;
    }
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
}

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

/// Usable bounds of the display nearest to `(x, y)`. Anchoring layout to a
/// single monitor keeps the preview from spilling onto a neighboring display
/// when more than one is attached. Falls back to the whole virtual screen if
/// the monitor query fails.
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

    virtual_screen_bounds()
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

/// Compute preview layout for mouse hover (relative to cursor position)
fn compute_mouse_layout(
    cursor_x: i32,
    cursor_y: i32,
    orig_dims: (u32, u32),
    follow_cursor: bool,
    preview_scale: PreviewScale,
    bounds: ScreenBounds,
) -> Option<PreviewLayout> {
    let offset = 20;
    let (orig_w, orig_h) = (orig_dims.0 as i32, orig_dims.1 as i32);
    let desired_scale = preview_scale.target_scale().unwrap_or(f32::INFINITY);

    if follow_cursor {
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
            let scale_x = avail_w as f32 / orig_w as f32;
            let scale_y = avail_h as f32 / orig_h as f32;
            let scale = scale_x.min(scale_y).min(desired_scale);
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

        Some(PreviewLayout {
            pos_x,
            pos_y,
            max_width,
            max_height,
            preview_w,
            preview_h,
        })
    } else {
        let left_width = cursor_x - bounds.left - offset;
        let right_width = bounds.right - cursor_x - offset;
        let full_height = bounds.height();

        let left_scale_x = left_width as f32 / orig_w as f32;
        let left_scale_y = full_height as f32 / orig_h as f32;
        let left_scale = left_scale_x.min(left_scale_y).min(desired_scale);

        let right_scale_x = right_width as f32 / orig_w as f32;
        let right_scale_y = full_height as f32 / orig_h as f32;
        let right_scale = right_scale_x.min(right_scale_y).min(desired_scale);

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

        Some(PreviewLayout {
            pos_x,
            pos_y,
            max_width,
            max_height,
            preview_w,
            preview_h,
        })
    }
}

/// Compute preview layout for keyboard hover (relative to item bounding rect)
/// Positions the preview so it doesn't block the selected file item
fn compute_keyboard_layout(
    item_rect: (i32, i32, i32, i32),
    orig_dims: (u32, u32),
    follow_cursor: bool,
    preview_scale: PreviewScale,
    bounds: ScreenBounds,
) -> Option<PreviewLayout> {
    let (item_left, item_top, item_right, item_bottom) = item_rect;
    let gap = 10;
    let (orig_w, orig_h) = (orig_dims.0 as i32, orig_dims.1 as i32);
    let desired_scale = preview_scale.target_scale().unwrap_or(f32::INFINITY);

    if follow_cursor {
        // Quadrant-based positioning relative to item rect edges
        let quadrants = [
            // Bottom-Right of item
            (
                bounds.right - item_right - gap,
                bounds.bottom - item_bottom - gap,
                item_right + gap,
                item_bottom + gap,
            ),
            // Bottom-Left of item
            (
                item_left - bounds.left - gap,
                bounds.bottom - item_bottom - gap,
                bounds.left,
                item_bottom + gap,
            ),
            // Top-Right of item
            (
                bounds.right - item_right - gap,
                item_top - bounds.top - gap,
                item_right + gap,
                bounds.top,
            ),
            // Top-Left of item
            (
                item_left - bounds.left - gap,
                item_top - bounds.top - gap,
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
            let scale_x = avail_w as f32 / orig_w as f32;
            let scale_y = avail_h as f32 / orig_h as f32;
            let scale = scale_x.min(scale_y).min(desired_scale);
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
            0 => (item_right + gap, item_bottom + gap),
            1 => (item_left - gap - media_width, item_bottom + gap),
            2 => (item_right + gap, item_top - gap - media_height),
            3 => (item_left - gap - media_width, item_top - gap - media_height),
            _ => (item_right + gap, item_bottom + gap),
        };

        Some(PreviewLayout {
            pos_x,
            pos_y,
            max_width,
            max_height,
            preview_w,
            preview_h,
        })
    } else {
        // Best spot mode: choose left or right side of item
        let left_width = item_left - bounds.left - gap;
        let right_width = bounds.right - item_right - gap;
        let full_height = bounds.height();

        let left_scale_x = left_width as f32 / orig_w as f32;
        let left_scale_y = full_height as f32 / orig_h as f32;
        let left_scale = left_scale_x.min(left_scale_y).min(desired_scale);

        let right_scale_x = right_width as f32 / orig_w as f32;
        let right_scale_y = full_height as f32 / orig_h as f32;
        let right_scale = right_scale_x.min(right_scale_y).min(desired_scale);

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
            item_left - gap - media_width
        } else {
            item_right + gap
        };
        // The same rule as the mouse path, centered on the row the preview
        // belongs to: beside the item it describes, not adrift in the display.
        let pos_y = centered_top((item_top + item_bottom) / 2, media_height, bounds);

        Some(PreviewLayout {
            pos_x,
            pos_y,
            max_width,
            max_height,
            preview_w,
            preview_h,
        })
    }
}

pub fn run_preview_window() {
    // Page sizes come from Windows.Data.Pdf, so this thread needs an apartment
    // before the first layout asks for one.
    pdf_preview::initialize_apartment();

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

        // Message loop
        let mut msg = MSG::default();
        while RUNNING.load(Ordering::SeqCst) {
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
            if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                if let Some(ref mut media) = *media_guard {
                    if media.advance_frame() {
                        needs_repaint = true;
                    }
                    if media.update_loading_frame() {
                        needs_repaint = true;
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

            // Check for completed background loads
            while let Ok(result) = load_rx.try_recv() {
                if result.generation == current_generation {
                    match result.media {
                        Some(media_data) => {
                            let mw = media_data.current_width() as i32;
                            let mh = media_data.current_height() as i32;

                            let pending = pending_load.take().filter(|pl| {
                                pl.generation == result.generation && !pl.spinner_shown
                            });
                            pending_load_cancel = None;

                            // Move before installing the frame. Crossing between
                            // displays of different scale sends WM_DPICHANGED,
                            // which resets the preview and would otherwise
                            // discard the frame we are about to show, leaving the
                            // other display's image stranded on screen.
                            if let Some(ref pl) = pending {
                                let _ = MoveWindow(hwnd, pl.pos_x, pl.pos_y, mw, mh, false);
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
                            render_layered_preview(hwnd);

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
                        None => {
                            // Loading failed, hide window
                            let _ = ShowWindow(hwnd, SW_HIDE);
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

            // Show loading spinner if a background load has been pending for 3+ seconds
            if let Some(ref mut pl) = pending_load {
                if !pl.spinner_shown && pl.started.elapsed() >= Duration::from_secs(2) {
                    pl.spinner_shown = true;
                    let loading = create_loading_media(pl.width, pl.height);
                    let _ = MoveWindow(
                        hwnd,
                        pl.pos_x,
                        pl.pos_y,
                        pl.width as i32,
                        pl.height as i32,
                        false,
                    );
                    // Install the spinner after the move so a WM_DPICHANGED
                    // reset from crossing displays cannot discard it.
                    if let Ok(mut current) = CURRENT_MEDIA.lock() {
                        *current = Some(loading);
                    }
                    // Paint the spinner before revealing the window so the
                    // previous preview cannot flash at the new position.
                    render_layered_preview(hwnd);
                    let _ = SetWindowPos(
                        hwnd,
                        HWND_TOPMOST,
                        pl.pos_x,
                        pl.pos_y,
                        pl.width as i32,
                        pl.height as i32,
                        SWP_NOACTIVATE | SWP_SHOWWINDOW,
                    );
                    let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
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
            while let Ok(preview_msg) = rx.try_recv() {
                match preview_msg {
                    PreviewMessage::Refresh => {
                        if latest_preview_msg.is_none() {
                            // A text preview's colors are in the painted frame,
                            // so a theme or Markdown switch rebuilds it from the
                            // hover it came from; every other preview only needs
                            // the frame composited again.
                            match (current_media_is_text(), current_show.clone()) {
                                (true, Some(show)) => latest_preview_msg = Some(show),
                                _ => refresh_requested = true,
                            }
                        }
                    }
                    other => {
                        latest_preview_msg = Some(other);
                        refresh_requested = false;
                    }
                }
            }

            if let Some(preview_msg) = latest_preview_msg {
                // Common variables for Show/ShowKeyboard - set in match, used after
                let mut show_path: Option<PathBuf> = None;
                let mut show_layout: Option<PreviewLayout> = None;
                let mut show_is_video: bool = false;
                let mut show_requested = false;
                let mut preview_scale = current_preview_scale();
                let mut show_dpi = 96u32;
                let show_snapshot = matches!(
                    preview_msg,
                    PreviewMessage::Show(..) | PreviewMessage::ShowKeyboard(..)
                )
                .then(|| preview_msg.clone());

                match preview_msg {
                    PreviewMessage::Show(path, x, y) => {
                        show_requested = true;
                        // Remember where this preview was opened from: the region
                        // that keeps a scrollable preview alive stretches from
                        // here to the preview, so the pointer can travel between
                        // the two without losing it.
                        set_text_scroll_anchor(x, y);

                        let bounds = monitor_bounds_from_point(x, y);
                        let dpi = monitor_dpi_from_point(x, y);
                        let follow_cursor = CONFIG.lock().map(|c| c.follow_cursor).unwrap_or(true);
                        preview_scale = effective_preview_scale(&path, preview_scale);

                        if let Some(orig_dims) = media_dimensions(&path, bounds, dpi) {
                            let is_video = is_video_file(&path);
                            if let Some(layout) = compute_mouse_layout(
                                x,
                                y,
                                orig_dims,
                                follow_cursor,
                                preview_scale,
                                bounds,
                            ) {
                                show_is_video = is_video;
                                show_layout = Some(layout);
                                show_path = Some(path);
                                show_dpi = dpi;
                            }
                        }
                    }
                    PreviewMessage::ShowKeyboard(path, il, it, ir, ib) => {
                        show_requested = true;
                        // The focused item lives inside the Explorer window, so
                        // its center resolves to that window's monitor.
                        let center = ((il + ir) / 2, (it + ib) / 2);
                        set_text_scroll_anchor(center.0, center.1);

                        let bounds = monitor_bounds_from_point(center.0, center.1);
                        let dpi = monitor_dpi_from_point(center.0, center.1);
                        let follow_cursor = CONFIG.lock().map(|c| c.follow_cursor).unwrap_or(true);
                        preview_scale = effective_preview_scale(&path, preview_scale);

                        if let Some(orig_dims) = media_dimensions(&path, bounds, dpi) {
                            let is_video = is_video_file(&path);
                            if let Some(layout) = compute_keyboard_layout(
                                (il, it, ir, ib),
                                orig_dims,
                                follow_cursor,
                                preview_scale,
                                bounds,
                            ) {
                                show_is_video = is_video;
                                show_layout = Some(layout);
                                show_path = Some(path);
                                show_dpi = dpi;
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
                        clear_text_scroll_keep_alive();

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
                        current_show = None;
                    }
                    PreviewMessage::Refresh => {
                        render_layered_preview(hwnd);
                    }
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

                    // A text preview is rendered at the size the layout planned
                    // for it: text is never scaled to fill a box, so the planned
                    // box is the box it draws into. Every other format is loaded
                    // against the free space it may be scaled within.
                    let (load_width, load_height) = if text_formats::is_text_file(&path) {
                        (preview_w, preview_h)
                    } else {
                        (max_width, max_height)
                    };

                    if show_snapshot.is_some() {
                        current_show = show_snapshot.clone();
                    }

                    if show_is_video {
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
                            // For video, hide our window and use ffplay
                            let _ = ShowWindow(hwnd, SW_HIDE);

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

                                if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                    let mut data = media_data;
                                    data.video_process = video_process;
                                    *current = Some(data);
                                }

                                current_video_path = Some(path.clone());
                                video_pos = (pos_x, pos_y, media_width, media_height);
                                let _ = ensure_video_window_topmost(
                                    pos_x,
                                    pos_y,
                                    media_width,
                                    media_height,
                                );
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

                        if let Some(cancel) = pending_load_cancel.take() {
                            cancel.store(true, Ordering::Release);
                        }
                        if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                            if let Some(ref mut media) = *media_guard {
                                media.cancel_background_work();
                            }
                            // Clear immediately so old pixels never flash while
                            // the new target is being decoded.
                            *media_guard = None;
                        }
                        let _ = ShowWindow(hwnd, SW_HIDE);

                        // Start background load; spinner will appear after 2s
                        current_generation += 1;
                        let gen = current_generation;
                        let load_cancel = Arc::new(AtomicBool::new(false));
                        pending_load_cancel = Some(Arc::clone(&load_cancel));
                        pending_load = Some(PendingLoad {
                            generation: gen,
                            started: Instant::now(),
                            pos_x,
                            pos_y,
                            width: preview_w,
                            height: preview_h,
                            spinner_shown: false,
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
                }
            } else if refresh_requested {
                render_layered_preview(hwnd);
            }

            // Ctrl+C over a text preview copies what is selected in it. The key is
            // polled rather than waited for: the preview never takes focus, so it
            // would never receive the keystroke as a message.
            if text_preview_copy_requested() {
                copy_text_preview(hwnd);
            }

            // Keep the pointer region in step with the window rather than only
            // with the paints: the window is moved when a frame is installed, and
            // the Explorer hook reads this on every one of its own ticks.
            publish_text_scroll_keep_alive(hwnd);

            std::thread::sleep(std::time::Duration::from_millis(16)); // ~60fps loop is enough and lowers idle CPU
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
    use super::{
        centered_top, compute_keyboard_layout, compute_mouse_layout, text_scroll_hold_region,
        ScreenBounds,
    };
    use crate::config::PreviewScale;
    use std::path::PathBuf;

    /// A 1920x1040 work area at the origin, which is all the placement math needs.
    fn screen() -> ScreenBounds {
        ScreenBounds {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1040,
        }
    }

    #[test]
    fn a_centered_preview_follows_the_cursor_and_stops_at_the_edge() {
        let cursor_y = 520;
        assert_eq!(centered_top(cursor_y, 300, screen()), cursor_y - 150);
        assert_eq!(centered_top(30, 800, screen()), 0);
        assert_eq!(centered_top(1000, 800, screen()), 240);
        // A preview taller than the display is pinned to its top edge rather than
        // clamped into a range that does not exist.
        assert_eq!(centered_top(520, 2000, screen()), 0);
    }

    #[test]
    fn best_position_puts_the_preview_on_the_cursors_own_line() {
        let layout = compute_mouse_layout(
            600,
            520,
            (400, 300),
            false,
            PreviewScale::Percent(100),
            screen(),
        )
        .expect("a layout");

        // Beside the cursor horizontally, centered on it vertically.
        assert_eq!((layout.pos_x, layout.pos_y), (620, 370));
        assert_eq!((layout.preview_w, layout.preview_h), (400, 300));
    }

    #[test]
    fn best_position_moves_a_small_preview_up_to_the_top_edge() {
        // A cursor near the top: the preview is centered on it as far as the edge
        // allows, instead of being centered on the display below the cursor.
        let layout = compute_mouse_layout(
            600,
            30,
            (400, 800),
            false,
            PreviewScale::Percent(100),
            screen(),
        )
        .expect("a layout");

        assert_eq!(layout.pos_y, 0);
        assert!(layout.pos_y + layout.preview_h as i32 <= screen().bottom);

        // And near the bottom it stops against that edge instead.
        let layout = compute_mouse_layout(
            600,
            1000,
            (400, 800),
            false,
            PreviewScale::Percent(100),
            screen(),
        )
        .expect("a layout");

        assert_eq!(layout.pos_y, 240);
    }

    #[test]
    fn follow_cursor_mode_keeps_its_quadrant_placement() {
        let layout = compute_mouse_layout(
            600,
            520,
            (400, 300),
            true,
            PreviewScale::Percent(100),
            screen(),
        )
        .expect("a layout");

        // Still offset below the cursor rather than centered on it.
        assert_eq!((layout.pos_x, layout.pos_y), (620, 540));
    }

    #[test]
    fn a_keyboard_preview_centers_on_the_row_it_describes() {
        let layout = compute_keyboard_layout(
            (200, 300, 300, 320),
            (400, 300),
            false,
            PreviewScale::Percent(100),
            screen(),
        )
        .expect("a layout");

        assert_eq!(layout.pos_x, 310);
        assert_eq!(layout.pos_y, 160);
    }

    /// The region that keeps a preview alive joins the point it was opened from
    /// to the preview: one pixel behind that point along the way, and the far edge
    /// — the one the pointer arrives at — extended by the grace a hand that
    /// overshoots it needs.
    #[test]
    fn the_hold_region_stretches_from_the_file_to_the_preview() {
        // A preview placed to the right of the cursor: the region starts one
        // pixel behind the pointer — a step further takes the pointer out of it —
        // and reaches past the preview on the right, the side it was reached from.
        let region = text_scroll_hold_region((620, 300, 1020, 700), (600, 500), 30);
        assert_eq!(region, (599, 300, 1050, 700));

        // Placed to the left instead: the mirror image, with the grace on the left.
        let region = text_scroll_hold_region((200, 100, 600, 500), (620, 300), 30);
        assert_eq!(region, (170, 100, 621, 500));

        // A preview in the pointer's own column is just the preview, joined to the
        // pointer's row: it was not travelled to from either side.
        let region = text_scroll_hold_region((300, 100, 500, 400), (400, 600), 30);
        assert_eq!(region, (300, 100, 500, 600));

        // With no grace asked for, the region ends at the preview.
        let region = text_scroll_hold_region((620, 300, 1020, 700), (600, 500), 0);
        assert_eq!(region, (599, 300, 1020, 700));

        // The point the preview was opened from is inside its own region, and one
        // pixel further back is not.
        let anchor = (900, 400);
        let region = text_scroll_hold_region((920, 300, 1300, 700), anchor, 30);
        assert!(region.0 <= anchor.0 && region.2 >= anchor.0);
        assert!(region.1 <= anchor.1 && region.3 >= anchor.1);
        assert!(region.0 > anchor.0 - 2);
    }

    /// The configured grace is measured in logical pixels, so it is the same
    /// distance under a hand on any display; the scrollbar it is there for scales
    /// with the text.
    #[test]
    fn the_far_edge_grace_follows_the_display_dpi() {
        assert_eq!(super::far_edge_grace(96, 40.0), 40);
        assert_eq!(super::far_edge_grace(120, 40.0), 50);
        assert_eq!(super::far_edge_grace(144, 40.0), 60);
        assert_eq!(super::far_edge_grace(192, 40.0), 80);

        // Any distance the configuration names takes the same road, and the one
        // that ends the grace where the preview does stays zero.
        assert_eq!(super::far_edge_grace(96, 12.5), 13);
        assert_eq!(super::far_edge_grace(192, 12.5), 25);
        assert_eq!(super::far_edge_grace(192, 0.0), 0);
    }

    /// Full mode is what a preview keeps its state for: with it off there is
    /// nothing to select into, nothing to scroll and no region to hold a pointer
    /// with, so a text preview is only ever a picture of a file.
    #[test]
    fn only_full_mode_gives_a_text_preview_a_state() {
        let mut path = std::env::temp_dir();
        path.push(format!(
            "rust-hover-preview-state-{}.txt",
            std::process::id()
        ));
        std::fs::write(&path, b"one\ntwo\nthree\n").unwrap();

        let options = |full_mode| crate::text_preview::TextPreviewOptions {
            theme: crate::config::TextTheme::Light,
            markdown_mode: crate::config::MarkdownMode::Rendered,
            font_scale_percent: 100,
            full_mode,
        };

        let normal = super::load_text_preview(&path, 400, 300, 96, options(false)).unwrap();
        assert!(
            normal.text_state.is_none(),
            "a preview that cannot be worked with keeps nothing to work with"
        );
        assert_eq!(normal.frames.len(), 1);

        let full = super::load_text_preview(&path, 400, 300, 96, options(true)).unwrap();
        let state = full.text_state.expect("full mode keeps its state");
        assert!(state.selection.is_none());
        assert!(!state.selecting);
        assert!(!state.dragging);
        assert!(
            !state.lines.is_empty(),
            "the painted lines are what a press is read against"
        );
        assert!(
            !state.can_scroll(),
            "a three-line file has nothing to scroll, but is still selectable"
        );

        let _ = std::fs::remove_file(&path);
    }

    /// The wheel and a drag both move a preview through this: a step is counted
    /// from where it is now and stops at either end of the document.
    #[test]
    fn scrolling_a_text_preview_stops_at_both_ends() {
        let scroll = super::TextPreviewState {
            path: PathBuf::from("preview.txt"),
            options: crate::text_preview::TextPreviewOptions {
                theme: crate::config::TextTheme::Light,
                markdown_mode: crate::config::MarkdownMode::Rendered,
                font_scale_percent: 100,
                full_mode: true,
            },
            dpi: 96,
            width: 800,
            height: 600,
            first_line: 100,
            visible_lines: 40,
            scrollable_lines: 200,
            scrollbar: None,
            dragging: false,
            lines: Vec::new(),
            selection: None,
            selecting: false,
        };

        assert!(scroll.can_scroll());
        assert_eq!(scroll.max_first_line(), 160);
        assert_eq!(scroll.scrolled_by(3), 103);
        assert_eq!(scroll.scrolled_by(-3), 97);
        assert_eq!(scroll.scrolled_by(1000), 160);
        assert_eq!(scroll.scrolled_by(-1000), 0);

        // A document that fits has nowhere to go.
        let fits = super::TextPreviewState {
            visible_lines: 200,
            scrollable_lines: 200,
            ..scroll
        };
        assert!(!fits.can_scroll());
        assert_eq!(fits.scrolled_by(3), 0);
    }
}

use crate::archive_formats;
use crate::archive_preview::{self, ArchivePreviewOptions};
use crate::cloud_files;
use crate::config::{
    decode_budget_bytes, image_decode_limits, read_within_budget, sanitize_image_cache_mb,
    sanitize_webp_playback_fps, MarkdownMode, PreviewScale, PreviewType, TextTheme,
    TransparentBackground, DEFAULT_IMAGE_CACHE_MB, DEFAULT_PREVIEW_SCALE_PERCENT,
    DEFAULT_SVG_SCALE_PERCENT, DEFAULT_TEXT_FONT_SCALE_PERCENT,
    DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS, DEFAULT_WEBP_PLAYBACK_FPS,
};
use crate::engine_processes;
use crate::office_formats;
use crate::office_preview;
use crate::office_render;
use crate::pdf_preview;
use crate::svg_animation;
use crate::svg_preview;
use crate::text_formats;
use crate::text_preview::{self, TextPreviewOptions};
use crate::video_formats::{self, is_video_file};
use crate::webview_preview;
use crate::wheel_input;
use crate::{CONFIG, RUNNING};
use gif::DecodeOptions;
use image::{AnimationDecoder, GenericImageView};
use once_cell::sync::Lazy;
use std::cell::RefCell;
use std::collections::{HashMap, VecDeque};
use std::fs::File;
use std::io::BufReader;
use std::os::windows::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
use std::ptr;
use std::sync::atomic::{AtomicBool, AtomicIsize, AtomicU32, Ordering};
use std::sync::mpsc::{channel, Receiver, RecvTimeoutError, Sender};
use std::sync::{Arc, Condvar, Mutex};
use std::time::{Duration, Instant, SystemTime};
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
    EnumWindows, GetCursorPos, GetSystemMetrics, GetWindow, GetWindowLongPtrW, GetWindowRect,
    GetWindowThreadProcessId, IsWindow, IsWindowVisible, LoadCursorW, MoveWindow, PeekMessageW,
    RegisterClassExW, SetForegroundWindow, SetWindowLongPtrW, SetWindowPos, ShowWindow,
    SystemParametersInfoW, TrackPopupMenu, TranslateMessage, UpdateLayeredWindow, CS_HREDRAW,
    CS_VREDRAW, GWL_EXSTYLE, GW_OWNER, HWND_TOPMOST, IDC_ARROW, MF_STRING, MSG,
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

/// Lines one wheel notch moves a text preview. Three is the step a text editor
/// takes, and it keeps a screenful to a few notches.
const TEXT_SCROLL_LINES_PER_NOTCH: i64 = 3;
const WHEEL_DELTA: i32 = 120;

/// How far behind the point a preview was opened from the region reaches, in
/// logical pixels. The pointer travels forwards from there, so this is only there
/// to keep the pixel under a hand at rest inside the region.
const TEXT_SCROLL_ANCHOR_SLACK_PIXELS: i32 = 1;
/// How far above and below the row the pointer is on the journey to a preview may
/// wander before it is out of it. The journey is made across a row of the list, so
/// the band is the hand's, not the preview's.
const TEXT_SCROLL_CORRIDOR_SLACK_PIXELS: i32 = 12;

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
/// it was waiting for — the render finishes, but the hover it was for is gone. The
/// spinner is placed a pixel off the pointer and follows it, which is what keeps
/// the pointer on the file it is waiting on; a pointer that moves into the box
/// anyway — a hand settling, or a move the box has not caught up with — is a
/// pointer still waiting for that file, not one leaving it.
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

    let corridor = (
        anchor.0.min(near_x) - TEXT_SCROLL_ANCHOR_SLACK_PIXELS,
        anchor.1.min(near_y) - TEXT_SCROLL_CORRIDOR_SLACK_PIXELS,
        anchor.0.max(near_x) + TEXT_SCROLL_ANCHOR_SLACK_PIXELS,
        anchor.1.max(near_y) + TEXT_SCROLL_CORRIDOR_SLACK_PIXELS,
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

static VIDEO_GEOMETRY_CACHE: Lazy<Mutex<HashMap<VideoGeometryKey, VideoGeometry>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

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
    /// hovered item's is. See `compute_keyboard_layout`.
    ShowKeyboard(PathBuf, i32, i32, i32, i32, Option<ScreenRegion>),
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
}

/// Represents different types of media we can display
enum MediaType {
    StaticImage,
    /// A still SVG document, drawn rather than decoded. It is a kind of its own
    /// rather than a static image because a document is composited over a backdrop
    /// of its own (see the tray's `Background` submenu).
    StaticSvg,
    AnimatedGif,
    AnimatedApng,
    AnimatedWebP,
    AnimatedSvg,
    Video,
    Pdf,
    Text,
    Archive,
    Office,
    Loading,
}

impl MediaType {
    /// The kind of preview this media is, as the tray's gates name them.
    fn kind(&self) -> Option<PreviewType> {
        match self {
            Self::StaticImage
            | Self::AnimatedGif
            | Self::AnimatedApng
            | Self::AnimatedWebP => Some(PreviewType::Images),
            Self::StaticSvg | Self::AnimatedSvg => Some(PreviewType::Svg),
            Self::Video => Some(PreviewType::Videos),
            Self::Text => Some(PreviewType::Text),
            Self::Pdf => Some(PreviewType::Pdf),
            Self::Archive => Some(PreviewType::Archives),
            Self::Office => Some(PreviewType::Office),
            Self::Loading => None,
        }
    }

    /// Whether this is the spinner standing in for a preview that is not ready.
    fn is_loading(&self) -> bool {
        matches!(self, Self::Loading)
    }

    /// Whether this is an SVG document rather than a picture. A document is drawn
    /// over a backdrop of its own, which is why the renderer asks.
    fn is_svg(&self) -> bool {
        matches!(self, Self::StaticSvg | Self::AnimatedSvg)
    }

    /// Whether this preview's appearance is painted into its own frame rather
    /// than recomposited from shared pixels, which is what decides whether a
    /// theme switch means rebuilding it.
    fn is_painted(&self) -> bool {
        matches!(self, Self::Text | Self::Archive)
    }
}

/// A single frame of image data
#[derive(Clone)]
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
            MediaType::AnimatedGif
                | MediaType::AnimatedApng
                | MediaType::AnimatedWebP
                | MediaType::AnimatedSvg
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
) {
    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(ref tx) = *sender {
            let _ = tx.send(PreviewMessage::ShowKeyboard(
                path.to_path_buf(),
                item_left,
                item_top,
                item_right,
                item_bottom,
                avoid,
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

fn is_gif_file(path: &Path) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_lowercase() == "gif")
        .unwrap_or(false)
}

fn is_webp_file(path: &Path) -> bool {
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
    let mut reader = image::ImageReader::open(path)
        .ok()?
        .with_guessed_format()
        .ok()?;
    reader.limits(image_decode_limits());

    reader.decode().ok()
}

/// Decode an image by the extension it is named with, for the files whose own
/// bytes are not asked what they are. Read under the same budget as every other
/// decoder, so the path is chosen by the setting and not by what it costs.
fn decode_image_by_extension(path: &PathBuf) -> Option<image::DynamicImage> {
    let mut reader = image::ImageReader::open(path).ok()?;
    reader.limits(image_decode_limits());

    reader.decode().ok()
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

/// The backdrop a picture is drawn over — and every other preview that is not a
/// document: a PDF page, a painted frame, a page Office rendered.
fn current_image_background() -> TransparentBackground {
    CONFIG
        .lock()
        .map(|cfg| cfg.image_background)
        .unwrap_or(TransparentBackground::Transparent)
}

/// The backdrop an SVG document is drawn over, which the tray keeps apart from a
/// picture's.
fn current_svg_background() -> TransparentBackground {
    CONFIG
        .lock()
        .map(|cfg| cfg.svg_background)
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

/// The share of the display an SVG document is drawn at, read from the configuration
/// the way the picture scale beside it is.
fn current_svg_scale() -> PreviewScale {
    CONFIG
        .lock()
        .map(|cfg| cfg.svg_scale)
        .unwrap_or(PreviewScale::Percent(DEFAULT_SVG_SCALE_PERCENT))
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

/// The file a hover message is about.
fn show_path(show: &PreviewMessage) -> Option<&PathBuf> {
    match show {
        PreviewMessage::Show(path, ..) | PreviewMessage::ShowKeyboard(path, ..) => Some(path),
        _ => None,
    }
}

/// Whether this hover is owed a render: an Office document with no page in the
/// cache yet, with the render tier switched on.
fn office_render_is_due(path: &Path) -> bool {
    office_formats::is_office_preview(path)
        && office_render::enabled()
        && office_render::cached_render(path).is_none()
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
    if !office_render_is_due(path) {
        return None;
    }

    office_render::request(path, width, height, generation);
    Some((path.to_path_buf(), generation))
}

/// The scale a preview is laid out and rendered with.
///
/// A PDF page is a vector, so the engine draws it at whatever size it is asked
/// for and a larger preview is sharper text rather than an enlarged raster. The
/// room the display has is therefore the page's size, and a configured percentage
/// below `100%` reduces that size rather than being ignored — the page is no
/// longer enlarged by it either, since enlarging a page is what fit-to-screen
/// already does (see `fit_reduced`).
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
/// of the screen. Where a document is played by the engine rather than drawn here, the
/// window is the size that came out of this and the page fills it, so the setting means
/// the same thing in both readers: see `webview_preview::frame_page`.
///
/// Every other format keeps the configured scale.
fn effective_preview_scale(
    path: &Path,
    preview_scale: PreviewScale,
    svg_scale: PreviewScale,
) -> PreviewScale {
    if pdf_preview::is_pdf_file(path) {
        fit_reduced(preview_scale)
    } else if is_text_preview(path) || archive_formats::is_archive_file(path) {
        PreviewScale::Percent(100)
    } else if office_formats::is_office_file(path) {
        // A page Office rendered is vector, so the room the display has is free
        // quality — the rule a PDF follows. The raster picture a workbook is
        // answered with where no printer can export a page is the exception: it
        // is only as good as the pixels it holds, so it follows the configured
        // scale the way an image does rather than being enlarged to fit.
        match office_preview::source_kind(path) {
            // Nothing to draw and a page on the way: what is on screen is the
            // spinner in a box of its own, and a box that small is placed at the
            // size it is rather than fitted to the display the way a page is.
            office_preview::SourceKind::None => PreviewScale::Percent(100),
            source => {
                if source.may_be_enlarged() {
                    fit_reduced(preview_scale)
                } else {
                    preview_scale
                }
            }
        }
    } else if svg_preview::is_svg_file(path) {
        fit_reduced(svg_scale)
    } else {
        preview_scale
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
/// The video gate is asked first, because it is the one that settles the
/// extensions the text list shares with it — `.ts` and `.mts` — by content, and
/// only it can tell a TypeScript source from the transport stream that goes by the
/// same name. A file it turns down falls through to the text gate; one it accepts
/// is a video, which the text renderer would find nothing readable in.
fn is_text_preview(path: &Path) -> bool {
    !is_video_file(path) && text_formats::is_text_file(path)
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
        // The background is settled once per row rather than once per pixel: it
        // cannot change inside the loop, and as a per-pixel match it cost a
        // branch on every pixel of every animation frame.
        TransparentBackground::Black => {
            for (px, dst) in src_row.chunks_exact(4).zip(dst_row.chunks_exact_mut(4)) {
                blend_pixel_over(px, dst, 0, 0, 0);
            }
        }
        TransparentBackground::White => {
            for (px, dst) in src_row.chunks_exact(4).zip(dst_row.chunks_exact_mut(4)) {
                blend_pixel_over(px, dst, 255, 255, 255);
            }
        }
        TransparentBackground::Checkerboard => {
            for (x, (px, dst)) in src_row
                .chunks_exact(4)
                .zip(dst_row.chunks_exact_mut(4))
                .enumerate()
            {
                let (r, g, b) = checkerboard_color(x as u32, y);
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

/// The bytes a frame of this shape is, when they fit what one hover may decode for
/// — the question for the readers that allocate a canvas of their own rather than
/// going through a decoder's limits.
///
/// A picture's shape is itself unbounded: a seven-thousand by ten-thousand
/// illustration is an ordinary thing to hover and is decoded at the size it is, so
/// what a file may ask for is the budget rather than a cap on its dimensions (see
/// `config::decode_budget_bytes`). The product is taken in `u64`, so a shape whose
/// frame overflows a `u32` cannot wrap into a size that would pass.
fn frame_bytes_within_budget(width: u32, height: u32, bytes_per_pixel: u64) -> Option<usize> {
    let bytes = width as u64 * height as u64 * bytes_per_pixel;
    (bytes <= decode_budget_bytes()).then_some(bytes as usize)
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

/// A still SVG document as `MediaData`: one frame, nothing streaming, and the kind
/// its backdrop is chosen by.
fn static_svg_media(frame: ImageFrame) -> MediaData {
    MediaData {
        media_type: MediaType::StaticSvg,
        ..static_image_media(frame)
    }
}

/// A still image as `MediaData`: one frame, nothing streaming.
fn static_image_media(frame: ImageFrame) -> MediaData {
    MediaData {
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
            return Some(static_image_media(frame));
        }
    }

    // A header that would not report its dimensions is not a reason to refuse the
    // file: the decoder has the last word on whether it is an image at all, and a
    // frame measured this way is simply not held.
    let img = if is_confirm_file_type_enabled() {
        decode_image_with_header_check(path)?
    } else {
        decode_image_by_extension(path)?
    };

    let (orig_width, orig_height) = img.dimensions();
    let (target_width, target_height) = match cache_key.as_ref() {
        Some(key) => (key.width, key.height),
        None => scale_dimensions(
            orig_width,
            orig_height,
            max_width,
            max_height,
            preview_scale,
        ),
    };

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
    let frame = ImageFrame {
        pixels: rgba_to_bgra(rgba.as_raw()),
        width: target_width,
        height: target_height,
        delay_ms: 0,
    };

    if let Some(key) = cache_key {
        image_cache_put(key, frame.clone());
    }

    Some(static_image_media(frame))
}

/// Draw an SVG into the box the layout planned.
///
/// Three things can end up playing a document that moves, in this order: the engine,
/// which is Chromium and plays the whole of SMIL and CSS; this app's own reader, for a
/// machine with no engine on it; and the still first frame, for a document whose
/// animation is outside what the reader can follow. A document that does not move is
/// drawn here and costs none of that.
///
/// Where the engine is going to play it, what is drawn here is the still frame the
/// engine's window lands on top of: a hover shows the document rather than a gap while
/// a browser starts, and a document whose page never arrives simply stays still.
fn load_svg_preview(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    cancel: &Arc<AtomicBool>,
) -> Option<MediaData> {
    let (document_width, document_height) = svg_preview::measure(path)?;
    let (target_width, target_height) = scale_dimensions(
        document_width,
        document_height,
        max_width,
        max_height,
        preview_scale,
    );

    if webview_preview::moves(path) {
        // The engine's window is put up by the preview loop, which is what knows where
        // the preview belongs and when it is on screen.
        let (pixels, width, height) =
            svg_preview::render(path, target_width, target_height, Some(cancel))?;

        return Some(static_svg_media(ImageFrame {
            pixels,
            width,
            height,
            delay_ms: 0,
        }));
    }

    if let Some(media) = load_animated_svg(
        path,
        max_width,
        max_height,
        preview_scale,
        Arc::clone(cancel),
    ) {
        return Some(media);
    }

    if cancel.load(Ordering::Acquire) {
        return None;
    }

    let (pixels, width, height) =
        svg_preview::render(path, target_width, target_height, Some(cancel))?;

    Some(static_svg_media(ImageFrame {
        pixels,
        width,
        height,
        delay_ms: 0,
    }))
}

/// Play an SVG that moves, frame by frame, into the queue every animation streams
/// through.
///
/// The renderer cannot play one itself — usvg drops animation, so a document's
/// declarations are worked out here and written back into it, once per frame, and each
/// frame is a parse and a rasterization of its own. What keeps that affordable is where
/// the frames go: the same streaming queue a GIF or an animated WebP fills, whose
/// playback holds one frame per thirty-third of a second and lets the producer run only
/// as far ahead as the queue has room for. A pass that fits in the player's window is
/// drawn once and looped from memory; one that does not is drawn again for each loop,
/// which is what the other animated formats do too.
///
/// Nothing is played for a document that does not move, or whose frame count is one:
/// the answer is `None` and the caller draws it as the still it is.
fn load_animated_svg(
    path: &Path,
    max_width: u32,
    max_height: u32,
    preview_scale: PreviewScale,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    let source = svg_preview::source(path)?;
    let document = roxmltree::Document::parse(source.as_ref()).ok()?;
    let playback = svg_animation::Playback::parse(&document)?;

    if playback.frames() <= 1 {
        return None;
    }

    let (document_width, document_height) = svg_preview::measure(path)?;
    let (target_width, target_height) = scale_dimensions(
        document_width,
        document_height,
        max_width,
        max_height,
        preview_scale,
    );
    if target_width == 0 || target_height == 0 {
        return None;
    }

    // The first frames are drawn before the preview is handed over, so the animation
    // opens on motion rather than on a spinner — the same head start the other
    // animated formats take.
    let mut initial_frames = Vec::new();
    let mut initial_bytes: usize = 0;
    let startup_frames = ANIMATION_STARTUP_FRAMES.min(playback.frames() as usize);

    while initial_frames.len() < startup_frames {
        if cancel.load(Ordering::Acquire) {
            return None;
        }

        let text = playback.document_at(&document, initial_frames.len() as u32);
        let Some((pixels, width, height)) =
            svg_preview::render_text(&text, target_width, target_height)
        else {
            break;
        };

        initial_bytes = initial_bytes.saturating_add(pixels.len());
        if initial_bytes > ANIMATION_RETAINED_BYTES {
            return None;
        }

        initial_frames.push(ImageFrame {
            pixels,
            width,
            height,
            delay_ms: svg_animation::FRAME_MS,
        });
    }

    if initial_frames.is_empty() {
        return None;
    }

    let rendered = initial_frames.len() as u32;
    let shared = Arc::new(Mutex::new(StreamedFrames {
        queue: VecDeque::new(),
        released: false,
    }));
    let shared_clone = Arc::clone(&shared);
    let loaded_flag = Arc::new(AtomicBool::new(false));
    let loaded_flag_clone = Arc::clone(&loaded_flag);
    let cancel_clone = Arc::clone(&cancel);
    let source = Arc::clone(&source);

    std::thread::spawn(move || {
        // The animation is worked out again here rather than handed over: what a
        // playback is made of are places in a parsed document, and this thread parses
        // its own from the text it owns.
        let Ok(document) = roxmltree::Document::parse(source.as_ref()) else {
            loaded_flag_clone.store(true, Ordering::Release);
            return;
        };
        let Some(playback) = svg_animation::Playback::parse(&document) else {
            loaded_flag_clone.store(true, Ordering::Release);
            return;
        };

        let repeats = playback.repeats_its_pass();
        let mut pass: u32 = 0;

        loop {
            let mut cancelled = false;
            let first = if pass == 0 { rendered } else { 0 };

            for index in first..playback.frames() {
                if cancel_clone.load(Ordering::Acquire)
                    || !await_frame_queue_room(&shared_clone, &cancel_clone)
                {
                    cancelled = true;
                    break;
                }

                // A pass that holds everything the document does is played again from
                // its start; one that was cut short takes the next stretch of the
                // document instead, so its animation keeps going forward.
                let frame = if repeats {
                    index
                } else {
                    pass * playback.frames() + index
                };
                let text = playback.document_at(&document, frame);
                let Some((pixels, width, height)) =
                    svg_preview::render_text(&text, target_width, target_height)
                else {
                    continue;
                };

                if let Ok(mut streamed) = shared_clone.lock() {
                    streamed.queue.push_back(ImageFrame {
                        pixels,
                        width,
                        height,
                        delay_ms: svg_animation::FRAME_MS,
                    });
                }
            }

            // The player gave back the frames it already showed, so the pass is drawn
            // again to play the animation another time. A document that is played
            // forward never stops until the hover does.
            if cancelled || (repeats && !streamed_frames_released(&shared_clone)) {
                break;
            }

            pass += 1;
        }

        loaded_flag_clone.store(true, Ordering::Release);
    });

    Some(MediaData {
        frames: initial_frames,
        shared_frames: Some(shared),
        all_frames_loaded: Some(loaded_flag),
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::AnimatedSvg,
        stream_cancel: Some(cancel),
        video_process: None,
        loading_start: Some(Instant::now()),
        text_state: None,
    })
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

/// Render a page of an Office document into the box the layout planned.
///
/// Two sources are tried in the order of what they are worth: the page Office
/// rendered in the background, when there is one, and the picture the document
/// saved inside itself. The renderer is asked for the source's own aspect ratio
/// inside the box, so a page that is not the shape the layout assumed is
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

/// Load an archive's contents as a page of its own, the way a text preview is
/// loaded: measured first, then painted into exactly the box the layout planned.
fn load_archive_preview(
    path: &Path,
    width: u32,
    height: u32,
    dpi: u32,
    options: ArchivePreviewOptions,
    cancel: &AtomicBool,
) -> Option<MediaData> {
    let (pixels, width, height) = archive_preview::render(path, width, height, dpi, options)?;
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    Some(MediaData {
        frames: vec![ImageFrame {
            pixels,
            width,
            height,
            delay_ms: 0,
        }],
        shared_frames: None,
        all_frames_loaded: None,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::Archive,
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

/// Every crop rectangle ffmpeg's detector reported for the file, with the number
/// of frames that reported it. The source dimensions are not needed to collect
/// them, which is what lets this run alongside the probe that reads them.
fn collect_video_crop_candidates(path: &PathBuf) -> HashMap<(u32, u32, u32, u32), u32> {
    let filter = format!(
        "cropdetect={}:{}:0",
        VIDEO_CROPDETECT_LIMIT, VIDEO_CROPDETECT_ROUND
    );

    let output = match Command::new("ffmpeg")
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
    {
        Ok(output) => output,
        Err(_) => return HashMap::new(),
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

/// The geometry a video preview is sized and cropped by, from the cache when the
/// file and its version have been probed before.
fn get_video_geometry(path: &PathBuf) -> Option<VideoGeometry> {
    let key = VideoGeometryKey {
        path: path.clone(),
        version: file_version(path),
    };

    if let Ok(cache) = VIDEO_GEOMETRY_CACHE.lock() {
        if let Some(cached) = cache.get(&key) {
            return Some(*cached);
        }
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

    let (src_w, src_h) = dimensions?;
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

    if let Ok(mut cache) = VIDEO_GEOMETRY_CACHE.lock() {
        if !cache.contains_key(&key) && cache.len() >= VIDEO_GEOMETRY_CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.insert(key, geometry);
    }

    Some(geometry)
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
        .creation_flags(CREATE_NO_WINDOW) // Hide the console window
        .spawn()
        .ok();

    // After spawning, try to set WS_EX_NOACTIVATE on the ffplay window
    // to prevent it from stealing focus
    if let Some(ref child_process) = child {
        set_noactivate_for_process(child_process.id());

        // The player is this app's own child, so it goes in the job with the engines:
        // a video preview that is up when the app is killed, or crashes, is not left
        // playing with nothing to close it.
        engine_processes::adopt(child_process.id());
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

    if is_video_file(path) {
        return load_video_thumbnail(path, max_width, max_height, preview_scale);
    }

    if pdf_preview::is_pdf_file(path) {
        return load_pdf_first_page(path, max_width, max_height, preview_scale);
    }

    if archive_formats::is_archive_file(path) {
        return load_archive_preview(
            path,
            max_width,
            max_height,
            dpi,
            current_archive_options(),
            &cancel,
        );
    }

    if office_formats::is_office_file(path) {
        return load_office_preview(path, max_width, max_height, preview_scale, &cancel);
    }

    if text_formats::is_text_file(path) {
        return load_text_preview(path, max_width, max_height, dpi, current_text_options());
    }

    // An SVG is drawn rather than decoded, and it is asked after the text lists for
    // the same reason the hook asks them in that order: a file is whichever kind
    // claims it first, and a name a user has put in the text list is a text file.
    if svg_preview::is_svg_file(path) {
        return load_svg_preview(path, max_width, max_height, preview_scale, &cancel);
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
    if video_formats::is_video_preview(path) {
        return get_video_geometry(path)
            .map(|g| (g.width, g.height))
            .or(Some((1920, 1080)));
    }

    // A PDF is measured from its own first page; one that cannot be read as a
    // PDF reports no dimensions, which drops the preview instead of guessing.
    if pdf_preview::is_pdf_preview(path) {
        return pdf_preview::page_dimensions(path);
    }

    // An Office document is measured from the page that has been rendered for it,
    // or from the picture the document saved. A document with neither is measured
    // as the page it is about to get — while the render tier is on, which is the
    // only case where one is coming.
    if office_formats::is_office_preview(path) {
        return office_preview::measure(path);
    }

    // An SVG is measured from the document rather than from a header: the size it
    // asks to be drawn at is the size the layout places, and the renderer draws it
    // at whatever box comes out of that. It is asked ahead of the `Images` gate —
    // which is what gets a document here, its name being an entry of the image list
    // — because a document is its own kind: what draws one is not a decoder, and
    // the switch for it is not the switch for pictures. A file of a kind that is
    // switched off reports no size, which is how the layout drops its preview.
    if svg_preview::is_svg_file(path) {
        if !PreviewType::Svg.enabled() {
            return None;
        }

        return svg_preview::measure(path);
    }

    // Whatever is left is a picture, so the `Images` gate is what decides it.
    if !PreviewType::Images.enabled() {
        return None;
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
    if is_text_preview(path) {
        let cap_width = (bounds.right - bounds.left).max(1) as u32;
        let cap_height = bounds.height().max(1) as u32;
        return text_preview::measure(path, cap_width, cap_height, dpi, current_text_options());
    }

    if archive_formats::is_archive_preview(path) {
        let cap_width = (bounds.right - bounds.left).max(1) as u32;
        let cap_height = bounds.height().max(1) as u32;
        return archive_preview::measure(
            path,
            cap_width,
            cap_height,
            dpi,
            current_archive_options(),
        );
    }

    get_media_dimensions(path)
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
fn text_preview_layout(
    path: &Path,
    layout: PreviewLayout,
    dpi: u32,
    place: impl FnOnce((u32, u32)) -> Option<PreviewLayout>,
) -> PreviewLayout {
    if !is_text_preview(path) {
        return layout;
    }

    let Some(size) = text_preview::measure(
        path,
        layout.preview_w,
        layout.max_height,
        dpi,
        current_text_options(),
    ) else {
        return layout;
    };

    place(size).unwrap_or(layout)
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

            let awaiting_render = media.is_none() && office_render_is_due(&request.path);

            let _ = result_tx.send(LoadResult {
                generation: request.generation,
                path: request.path.clone(),
                media,
                awaiting_render,
            });
        }
    })
}

/// How long a background load may run before the spinner is put up for it.
///
/// The window is hidden while a load runs, so one that finishes inside this has
/// gone straight from nothing to the preview: the delay is what keeps a decode
/// that takes a few milliseconds from flashing a spinner on the way past. A load
/// that came back waiting on the render tier is not given it — it has nothing to
/// show and seconds of work ahead of it — which is what `spinner_due` decides.
const LOAD_SPINNER_DELAY_SECS: u64 = 2;

/// What placing a hover's preview again needs, kept on a pending load.
///
/// The size is not measured again when the preview follows the pointer: a cursor
/// moving along the item it belongs to finds the same media, and measuring it per
/// tick would re-read a header, a listing or a document sixty times a second to
/// learn what the hover already knew. What is recomputed is the place — that is
/// what the cursor decides.
#[derive(Clone, Copy)]
struct HoverPlacement {
    orig_dims: (u32, u32),
    avoid: Option<ScreenRegion>,
    follow_cursor: bool,
    preview_scale: PreviewScale,
    /// Whether this hover is the waiting spinner, which is placed flush at the
    /// pointer's own corner and kept there while the wait runs. See
    /// `compute_mouse_layout`.
    flush_at_cursor: bool,
}

/// Tracks a pending background load so we can show the spinner while it runs.
struct PendingLoad {
    generation: u64,
    /// The file this load is for, which is what decides whether the engine plays it
    /// and what the engine is pointed at.
    path: PathBuf,
    started: Instant,
    pos_x: i32,
    pos_y: i32,
    width: u32,
    height: u32,
    spinner_shown: bool,
    /// The mouse hover this load came from, if it was one. A preview that is still
    /// on its way follows the pointer, so it is placed again for a cursor that has
    /// moved along the item since; a keyboard hover's placement belongs to the
    /// item and carries none.
    placement: Option<HoverPlacement>,
    /// Whether the load came back with nothing to draw and a page on the way.
    /// Nothing can be shown until that page lands, and asking for it is seconds
    /// of work, so the spinner goes up at once rather than after the delay a load
    /// that might finish in milliseconds is given.
    awaiting_render: bool,
    /// Whether this load is replacing what is already on screen — the page that
    /// arrived for the hover that is up — rather than opening a new preview. An
    /// upgrade never shows the spinner: what is there stays where it is, at its own
    /// size, until the page is ready.
    upgrade: bool,
}

impl PendingLoad {
    /// Whether the spinner is due for this load: at once where the wait is a
    /// render's, and once it has run for `LOAD_SPINNER_DELAY_SECS` otherwise.
    ///
    /// An upgrade is never due — what is on screen stays where it is, at its own
    /// size, until the page replaces it — and a load that already has its spinner
    /// up is not due again.
    fn spinner_due(&self) -> bool {
        !self.spinner_shown
            && !self.upgrade
            && (self.awaiting_render
                || self.started.elapsed() >= Duration::from_secs(LOAD_SPINNER_DELAY_SECS))
    }

    /// Place this load's preview again for `cursor`, when it is one that follows
    /// the pointer: the size the hover measured, its `Avoid` region and
    /// its scale, and the display the pointer is on now. Answers whether the place
    /// it came out at is a new one, so the window is only moved when it is.
    ///
    /// A load that answered nothing, or a keyboard hover, is left where it is.
    fn follow_pointer(&mut self, cursor: POINT) -> bool {
        let Some(placement) = self.placement else {
            return false;
        };

        let Some(layout) = compute_mouse_layout(
            cursor.x,
            cursor.y,
            placement,
            monitor_bounds_from_point(cursor.x, cursor.y),
        ) else {
            return false;
        };

        let moved = (layout.pos_x, layout.pos_y) != (self.pos_x, self.pos_y)
            || (layout.preview_w, layout.preview_h) != (self.width, self.height);
        if moved {
            self.pos_x = layout.pos_x;
            self.pos_y = layout.pos_y;
            self.width = layout.preview_w;
            self.height = layout.preview_h;
        }

        moved
    }
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

        if matches!(media.media_type, MediaType::Video) {
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
        // transparent to avoid. Every other preview is composited over that
        // backdrop as it always was — a document over the one of its own, and
        // everything else over the picture's.
        let background = if media.media_type.is_loading() {
            TransparentBackground::Transparent
        } else if media.media_type.is_svg() {
            current_svg_background()
        } else {
            current_image_background()
        };
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

/// Put the loading spinner on screen for a pending load, at the box that load is
/// planned for.
///
/// A window that is not on screen is moved before the spinner is installed, so a
/// `WM_DPICHANGED` reset from crossing displays cannot discard it, and the spinner
/// is painted before the window is revealed, so the previous preview cannot flash
/// at the new place. It is also what moves a spinner whose box has changed size
/// while it was up: the frame is drawn at the size of the box it goes into.
unsafe fn show_loading_spinner(hwnd: HWND, pl: &PendingLoad) {
    // A window that is already on screen is not moved ahead of its frame, for the
    // reason the page's install gives: what a layered window shows between one
    // paint and the next is the surface it already has, stretched into whatever
    // box the window has, so a spinner whose box changed under it would be drawn
    // as a bar until the new frame lands. `UpdateLayeredWindow` applies the place
    // and the size with the frame, which is where the move below happens instead.
    let visible = IsWindowVisible(hwnd).as_bool();
    if !visible {
        let _ = MoveWindow(
            hwnd,
            pl.pos_x,
            pl.pos_y,
            pl.width as i32,
            pl.height as i32,
            false,
        );
    }

    let loading = create_loading_media(pl.width, pl.height);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(loading);
    }

    if visible {
        render_layered_preview_at(hwnd, pl.pos_x, pl.pos_y);
    } else {
        render_layered_preview(hwnd);
    }
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

/// Publish — or withdraw — the regions in which the pointer keeps what is on
/// screen alive.
///
/// The Explorer hook polls this to decide whether the pointer over the preview
/// means "the user is reading this" or "dismiss it and show what is underneath",
/// and the wheel hook asks the same question before it decides whether the wheel
/// belongs to Explorer or to the preview. Two things hold the pointer: a text
/// preview in full mode, which the pointer can rest on to select from and scroll,
/// and the spinner a page is being rendered behind, which has nothing under it to
/// hand the pointer back to and no page yet to be shown in its place.
unsafe fn publish_pointer_hold(hwnd: HWND) {
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
                )
                .to_vec()
            })
    } else if waiting {
        // The box the spinner itself occupies: the one place on screen where the
        // pointer is over this app's own window rather than over the file, and so
        // the one place where what is under the pointer cannot say whether the
        // pointer is still on the file it is waiting for. It is placed a pixel off
        // the pointer, so a pointer that ends up inside this box has not gone
        // anywhere.
        let mut rect = RECT::default();
        GetWindowRect(hwnd, &mut rect)
            .ok()
            .map(|_| vec![(rect.left, rect.top, rect.right, rect.bottom)])
    } else {
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

/// Whether the pointer is inside a region that keeps what is on screen alive.
/// Answered from a published rectangle, so the Explorer hook can ask on every poll
/// tick.
pub fn preview_pointer_hold(x: i32, y: i32) -> bool {
    if !TEXT_PREVIEW_HOLDING.load(Ordering::Acquire)
        && !WAITING_PREVIEW_HOLDING.load(Ordering::Acquire)
    {
        return false;
    }

    let Ok(published) = POINTER_HOLD_REGIONS.lock() else {
        return false;
    };

    (*published)
        .as_ref()
        .map(|regions| {
            regions.iter().any(|(left, top, right, bottom)| {
                x >= *left && x < *right && y >= *top && y < *bottom
            })
        })
        .unwrap_or(false)
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
const MIN_AVOID_ROOM_PX: i32 = 64;

/// A layout moved off the region the item it describes draws, as the `Avoid` setting
/// measured it: the name the file is listed under at `Avoid Filename`, and that name
/// with the columns a row writes beside it at `Avoid Details`.
///
/// A preview is placed beside what it belongs to rather than over it, and the text of
/// the item it came from is part of what it belongs to: that item stays readable while
/// its preview is up, which is what the `Avoid` setting asks for. The placement the
/// position mode chose is therefore moved by the shortest step that clears that region —
/// past its right edge, past its left, under it or over it, whichever asks the least
/// of the preview — and only a step the display has room for is taken, so a preview
/// moved off one edge is never pushed off another.
///
/// A preview too large for every one of those rooms is *resized* into the roomiest of
/// them rather than left where it covers the text. That is the case a preview filling
/// the display lands in: nothing can be moved into place beside a name while the
/// preview is as wide and as tall as the display, so it is the size that gives, and the
/// preview shows as much as the room beside the name can hold — which is the same rule
/// that sized it in the first place, applied to the room that is left. A room too small
/// to be worth having is not taken at all, so a preview is never squeezed into a sliver
/// to get off a name that a usable preview would have covered anyway.
///
/// `gap` is the distance the placement keeps from what it is beside, so the text is
/// cleared by that much rather than touched at its edge.
fn avoiding_text(
    layout: PreviewLayout,
    orig_dims: (u32, u32),
    preview_scale: PreviewScale,
    avoid: Option<ScreenRegion>,
    gap: i32,
    bounds: ScreenBounds,
) -> PreviewLayout {
    let Some((text_left, text_top, text_right, text_bottom)) = avoid else {
        return layout;
    };

    let (left, top) = (layout.pos_x, layout.pos_y);
    let (width, height) = (layout.preview_w as i32, layout.preview_h as i32);

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

    let mut best: Option<(i64, i32, PreviewLayout)> = None;
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
        if size_there < natural_size && room < MIN_AVOID_ROOM_PX {
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

        // The largest preview wins, and the shortest move breaks a tie: every way out
        // that fits the preview as it stands offers it the same size, so those are the
        // ones the move decides between, and only a preview that has to shrink is
        // chosen between by what the room holds.
        let area = preview_w as i64 * preview_h as i64;
        let step = (placement.pos_x - left).abs() + (placement.pos_y - top).abs();
        let better = match &best {
            Some((best_area, best_step, _)) => {
                area > *best_area || (area == *best_area && step < *best_step)
            }
            None => true,
        };
        if better {
            best = Some((area, step, placement));
        }
    }

    match best {
        Some((_, _, placement)) => placement,
        None => layout,
    }
}

/// Compute preview layout for mouse hover (relative to cursor position)
///
/// `placement` is what the hover asks for — the size the preview was measured at,
/// the text to keep it off, the position mode it follows and the scale it is drawn
/// with — the same reading a pending load keeps to place its preview again as the
/// pointer moves (see `HoverPlacement`).
///
/// A placement that is the waiting spinner (`flush_at_cursor`) is placed by its own
/// rule: the corner nearest the pointer is put at the pointer — a single pixel off
/// it, which is what `offset` below is — in whichever of the four quadrants the
/// display has room for the spinner, and the name it covers is not stepped around:
/// a spinner waiting on the page of the file under the hand says what it is by being
/// at the hand, and one placed a row away from it says nothing about what is being
/// waited on. Every other preview keeps the margin its position mode leaves and the
/// room the `Avoid` setting asks for.
fn compute_mouse_layout(
    cursor_x: i32,
    cursor_y: i32,
    placement: HoverPlacement,
    bounds: ScreenBounds,
) -> Option<PreviewLayout> {
    let HoverPlacement {
        orig_dims,
        avoid,
        follow_cursor,
        preview_scale,
        flush_at_cursor,
    } = placement;

    // How far off the pointer a preview is placed. The spinner touches the pointer —
    // one pixel of standoff, because the preview window is what a mouse message at
    // the pointer lands on, and the pointer has to keep clicking and probing the file
    // it is waiting on rather than the spinner that is waiting on it. Every other
    // preview keeps the margin the position modes are written around.
    let offset = if flush_at_cursor { 1 } else { 20 };
    let (orig_w, orig_h) = (orig_dims.0 as i32, orig_dims.1 as i32);

    if follow_cursor || flush_at_cursor {
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

        // The spinner is placed and left there: the step off the name is what the
        // flush rule is instead of.
        if flush_at_cursor {
            return Some(layout);
        }

        Some(avoiding_text(
            layout,
            orig_dims,
            preview_scale,
            avoid,
            offset,
            bounds,
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
            avoid,
            offset,
            bounds,
        ))
    }
}

/// The least room beside an item a keyboard preview will squeeze into before it
/// stops treating the item as something to sit beside. Below this the free space
/// past the item's edge — or past the region a row is kept off — is a sliver, and the
/// preview is placed from the item's middle instead — see `compute_keyboard_layout`.
const MIN_BESIDE_ROOM_PX: i32 = 64;

/// Compute preview layout for keyboard hover (relative to item bounding rect)
/// Positions the preview so it doesn't block the selected file item
///
/// `avoid` is the region the `Avoid` setting keeps a preview off — the boxes each
/// piece of the item's own text is drawn in, which the hook reads off the item's
/// children in one batched call (see `explorer_hook::item_text_box`). It is what the
/// placement is kept clear of *and* where a row's placement is measured from, so a row
/// is only cleared as far as the setting asks; with nothing kept off, an item is
/// placed by the position mode alone. See `avoiding_text`.
fn compute_keyboard_layout(
    item_rect: (i32, i32, i32, i32),
    orig_dims: (u32, u32),
    follow_cursor: bool,
    avoid: Option<ScreenRegion>,
    preview_scale: PreviewScale,
    bounds: ScreenBounds,
) -> Option<PreviewLayout> {
    let (item_left, item_top, item_right, item_bottom) = item_rect;
    let gap = 10;
    let (orig_w, orig_h) = (orig_dims.0 as i32, orig_dims.1 as i32);

    // An item far wider than it is tall and at least half the display across is a
    // row of the list — Content view draws every item that way, as a box as wide as
    // the view with the name and the columns written into its left end.
    let item_width = (item_right - item_left).max(0);
    let item_height = (item_bottom - item_top).max(1);
    let display_width = (bounds.right - bounds.left).max(1);
    let row_shaped = item_width >= item_height * 4 && item_width * 2 >= display_width;

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
    // with nothing kept off at all are both left to the placement below, which anchors
    // them at their middle: there is nowhere beside such a row to put a preview, and
    // the display's own room is all there is.
    if row_shaped {
        if let Some(tail_right) = avoid
            .map(|(_, _, right, _)| right)
            .filter(|right| *right > item_left && bounds.right - *right - gap >= MIN_BESIDE_ROOM_PX)
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
                avoid,
                gap,
                bounds,
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
            avoid,
            gap,
            bounds,
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
            if edge_left_width < MIN_BESIDE_ROOM_PX && edge_right_width < MIN_BESIDE_ROOM_PX {
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
            avoid,
            gap,
            bounds,
        ))
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
        // The page the render tier owes the preview on screen, and whether one
        // has been asked for and is being waited on.
        let mut office_render_pending: Option<(PathBuf, u64)> = None;
        // A page that arrived for the hover already on screen, and is being loaded
        // to replace what is there rather than to open a new preview.
        let mut office_upgrade: Option<PathBuf> = None;

        // Message loop
        let mut msg = MSG::default();
        // A message the idle wait took off the channel, held for the drain below
        // rather than acted on where it was received.
        let mut carried_preview_msg: Option<PreviewMessage> = None;
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

            // The engine plays a document in a window of its own, and that window is
            // put up only once the page has arrived: what is underneath it — the still
            // frame this app drew — comes down then, and is put back by whatever
            // preview is shown next.
            if webview_preview::is_showing() && IsWindowVisible(hwnd).as_bool() {
                let _ = ShowWindow(hwnd, SW_HIDE);
            }

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
                            match pending.as_ref().filter(|_| visible) {
                                Some(pl) => render_layered_preview_at(hwnd, pl.pos_x, pl.pos_y),
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

                                // A document the engine plays is handed over once its
                                // still frame is up: the engine's window lands on top
                                // of that when the page has arrived, and until then
                                // what is on screen is the document itself.
                                if webview_preview::moves(&pl.path) {
                                    let area = webview_preview::Area {
                                        x: pl.pos_x,
                                        y: pl.pos_y,
                                        width: mw,
                                        height: mh,
                                    };
                                    // The document is an SVG — that is what the engine
                                    // plays — so the backdrop is the one the tray keeps
                                    // for documents rather than the picture's.
                                    webview_preview::show(&pl.path, area, current_svg_background());
                                }
                            }

                            // A page for this document is one Office start away,
                            // so it is asked for as soon as the hover is up rather
                            // than after the pointer has rested on it.
                            office_render_pending = request_office_render(
                                &result.path,
                                result.generation,
                                render_box.0,
                                render_box.1,
                            );
                        }
                        None if result.awaiting_render => {
                            // Nothing to draw yet and a page on the way: the
                            // pending load stays armed, so the preview is not
                            // dropped and the page has somewhere to land — and it
                            // is told what it is waiting on, which is what puts
                            // the spinner up at once rather than after the delay
                            // a load that may be about to finish is given. The
                            // page is asked for in the box its family's pages
                            // have: the spinner's own box is a spinner's, and
                            // says nothing about how large the page will be
                            // drawn.
                            if let Some(pl) = pending_load.as_mut() {
                                pl.awaiting_render = true;
                            }
                            let (width, height) = office_formats::default_page_size(&result.path);
                            office_render_pending = request_office_render(
                                &result.path,
                                result.generation,
                                width,
                                height,
                            );
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
            if let Some(ref mut pl) = pending_load {
                if let Some(cursor) = cursor_position() {
                    let box_before = (pl.width, pl.height);
                    if pl.follow_pointer(cursor) && pl.spinner_shown {
                        if (pl.width, pl.height) == box_before {
                            let _ = MoveWindow(
                                hwnd,
                                pl.pos_x,
                                pl.pos_y,
                                pl.width as i32,
                                pl.height as i32,
                                false,
                            );
                        } else {
                            // The spinner's own box changed size, so its frame is
                            // drawn again at the size it now goes into.
                            show_loading_spinner(hwnd, pl);
                        }
                    }
                }
            }

            // Show the loading spinner while a background load runs, once the
            // wait is worth showing — at once for one that is waiting on a page
            // to be rendered, and after a moment for one that may be about to
            // finish (see `spinner_due`).
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
            // The render tier's own message, held apart from the hovers: it is
            // not a hover to act on but an answer about the one on screen.
            let mut office_render_ready: Option<(PathBuf, u64, bool)> = None;
            let mut next_preview_msg = carried_preview_msg.take();
            loop {
                let Some(preview_msg) = next_preview_msg.or_else(|| rx.try_recv().ok()) else {
                    break;
                };
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
                        if latest_preview_msg.is_none() && office_render_ready.is_none() {
                            office_render_ready = Some((path, generation, ok));
                        }
                    }
                    other => {
                        latest_preview_msg = Some(other);
                        refresh_requested = false;
                    }
                }
            }

            // A page the render tier has finished with, for the hover that asked
            // for it: that hover is replayed, which measures the page itself and
            // draws it in place of the spinner it supersedes. A payload from an
            // older hover is dropped here — the page it wrote is kept as far as the
            // cache budget allows, and no further.
            if let Some((ready_path, ready_generation, ready_ok)) = office_render_ready {
                let shown = current_show.as_ref().and_then(show_path);
                let hovered = ready_generation == current_generation
                    && shown.map(|path| path.as_path()) == Some(ready_path.as_path());

                if hovered && ready_ok {
                    if latest_preview_msg.is_none() {
                        // The wait is over, so the request stops being the pending
                        // one here — where the page is actually taken up.
                        office_render_pending = None;
                        // The page is there, so the hover is replayed: that measures
                        // the page itself, moves the window to its size and loads it.
                        // It is an upgrade rather than a new preview, though, and what
                        // is on screen stays while it happens — hiding the spinner for
                        // the second a large picture takes to decode is a blink the
                        // user sees and reads as the preview failing.
                        office_upgrade = Some(ready_path.clone());
                        // A mouse hover is replayed where the pointer is now: the
                        // spinner it replaces was kept with the pointer while the
                        // render ran, and a page that jumped back to where the
                        // hover started would jump away from where it was waited
                        // for. A keyboard hover is the item's own place and is
                        // replayed as it came.
                        latest_preview_msg = match (current_show.clone(), cursor_position()) {
                            (Some(PreviewMessage::Show(path, _, _, avoid)), Some(cursor)) => {
                                Some(PreviewMessage::Show(path, cursor.x, cursor.y, avoid))
                            }
                            (show, _) => show,
                        };
                    }
                    // A newer message was in hand, so the page is not shown now. The
                    // wait it was rendered for is left standing rather than cleared
                    // with it: what is being waited on is still that page, so the cap
                    // on waiting keeps something to measure and the spinner comes down
                    // when its time is up instead of hanging there for good.
                } else if hovered {
                    office_render_pending = None;
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
                        pending_load = None;
                        let _ = ShowWindow(hwnd, SW_HIDE);
                        if let Ok(mut current) = CURRENT_MEDIA.lock() {
                            *current = None;
                        }
                    }
                } else {
                    office_render_pending = None;
                    // The hover this page was rendered for is over: it landed after
                    // the pointer had moved on, so nothing is waiting for it. What was
                    // rendered is kept as far as the budget allows and no further — at
                    // a size of nothing it is dropped here rather than held for a
                    // hover that has already gone.
                    office_render::hover_ended(&ready_path);
                }
            }

            // The display under the preview changed and the frame that was on screen
            // went with it. The window proc could only discard what was drawn; the
            // hover it came from is what knows how to draw it again, at the scale of
            // the display the pointer is on now. A newer message in hand is left to
            // speak for itself, and the flag waits for a tick where none does.
            if latest_preview_msg.is_none() && DISPLAY_RESET.swap(false, Ordering::AcqRel) {
                latest_preview_msg = current_show.clone();
            }

            // The engine could not be had for the preview that is up — the folder it
            // keeps its state in is held by a browser that is not this app's — so the
            // hover is replayed and laid out again, where the document is played by
            // this app's own reader instead of being left as the still frame the
            // engine's window was going to land on.
            if latest_preview_msg.is_none() && webview_preview::take_failure_notice() {
                latest_preview_msg = current_show.clone();
            }

            if let Some(preview_msg) = latest_preview_msg {
                // Common variables for Show/ShowKeyboard - set in match, used after
                let mut show_path: Option<PathBuf> = None;
                let mut show_layout: Option<PreviewLayout> = None;
                let mut show_placement: Option<HoverPlacement> = None;
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
                        preview_scale =
                            effective_preview_scale(&path, preview_scale, current_svg_scale());

                        // A document with no page rendered for it yet is answered with
                        // the waiting spinner, and that spinner is placed flush at the
                        // pointer rather than a margin away from it: it is the wait for
                        // the file under the hand, and the hand is where it belongs.
                        let waiting_spinner = office_formats::is_office_preview(&path)
                            && matches!(
                                office_preview::source_kind(&path),
                                office_preview::SourceKind::None
                            );

                        if let Some(orig_dims) = media_dimensions(&path, bounds, dpi) {
                            let is_video = is_video_file(&path);
                            let placement = HoverPlacement {
                                orig_dims,
                                avoid,
                                follow_cursor,
                                preview_scale,
                                flush_at_cursor: waiting_spinner,
                            };
                            if let Some(layout) = compute_mouse_layout(x, y, placement, bounds) {
                                let layout = text_preview_layout(&path, layout, dpi, |size| {
                                    compute_mouse_layout(
                                        x,
                                        y,
                                        HoverPlacement {
                                            orig_dims: size,
                                            ..placement
                                        },
                                        bounds,
                                    )
                                });
                                show_is_video = is_video;
                                show_layout = Some(layout);
                                show_placement = Some(placement);
                                show_path = Some(path);
                                show_dpi = dpi;
                            }
                        }
                    }
                    PreviewMessage::ShowKeyboard(path, il, it, ir, ib, avoid) => {
                        show_requested = true;
                        // The focused item lives inside the Explorer window, so
                        // its center resolves to that window's monitor.
                        let center = ((il + ir) / 2, (it + ib) / 2);
                        set_text_scroll_anchor(center.0, center.1);

                        let bounds = monitor_bounds_from_point(center.0, center.1);
                        let dpi = monitor_dpi_from_point(center.0, center.1);
                        let follow_cursor = CONFIG.lock().map(|c| c.follow_cursor).unwrap_or(true);
                        preview_scale =
                            effective_preview_scale(&path, preview_scale, current_svg_scale());

                        if let Some(orig_dims) = media_dimensions(&path, bounds, dpi) {
                            let is_video = is_video_file(&path);
                            if let Some(layout) = compute_keyboard_layout(
                                (il, it, ir, ib),
                                orig_dims,
                                follow_cursor,
                                avoid,
                                preview_scale,
                                bounds,
                            ) {
                                let layout = text_preview_layout(&path, layout, dpi, |size| {
                                    compute_keyboard_layout(
                                        (il, it, ir, ib),
                                        size,
                                        follow_cursor,
                                        avoid,
                                        preview_scale,
                                        bounds,
                                    )
                                });
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
                        office_render_pending = None;
                        office_upgrade = None;
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

                    // A page that arrived for the hover already on screen is an
                    // upgrade: what is there — the spinner — stays up while the page
                    // is loaded, and is replaced when it lands.
                    let upgrading = office_upgrade.as_deref() == Some(path.as_path());
                    office_upgrade = None;

                    // A text or archive preview is rendered at the size the
                    // layout planned for it: both are painted at a fixed font
                    // size, so the planned box is the box they draw into rather
                    // than a space to be scaled within. Every other format is
                    // loaded against the free space it may be scaled within.
                    let painted = is_text_preview(&path) || archive_formats::is_archive_file(&path);
                    let (load_width, load_height) = if painted {
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

                        // A preview that is not the engine's is drawn here, so the
                        // engine's window — if one is still up — comes down as this
                        // one goes up.
                        if !webview_preview::moves(&path) {
                            webview_preview::hide();
                        }

                        if let Some(cancel) = pending_load_cancel.take() {
                            cancel.store(true, Ordering::Release);
                        }
                        if !upgrading {
                            if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                if let Some(ref mut media) = *media_guard {
                                    media.cancel_background_work();
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
                            path: path.clone(),
                            started: Instant::now(),
                            pos_x,
                            pos_y,
                            width: preview_w,
                            height: preview_h,
                            spinner_shown: false,
                            placement: show_placement,
                            awaiting_render: false,
                            upgrade: upgrading,
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
                    office_render_pending = None;
                    office_upgrade = None;
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
            // here and its page left in the cache.
            let shown_path = current_show.as_ref().and_then(show_path).cloned();

            let render_wait = office_render_pending.as_ref().map(|(path, generation)| {
                let hovered = *generation == current_generation
                    && shown_path.as_deref() == Some(path.as_path());
                let waited = pending_load
                    .as_ref()
                    .map(|pl| pl.started.elapsed() >= Duration::from_secs(OFFICE_RENDER_WAIT_SECS))
                    .unwrap_or(false);

                (hovered, waited)
            });

            match render_wait {
                Some((false, _)) => office_render_pending = None,
                Some((true, true)) => {
                    // The preview has waited as long as it waits — a page that has
                    // not arrived by now may still be coming, a very large document
                    // takes as long as it takes — so the spinner comes down and the
                    // hover is left to itself. The render is not abandoned with it:
                    // it runs on and its page is cached, so the next hover of that
                    // file shows it. Only the engine itself can say a render failed,
                    // and it remembers that for the file.
                    let abandoned = office_render_pending
                        .take()
                        .map(|(path, _)| path)
                        .filter(|path| shown_path.as_deref() == Some(path.as_path()));

                    if abandoned.is_some() {
                        pending_load = None;
                        if let Some(cancel) = pending_load_cancel.take() {
                            cancel.store(true, Ordering::Release);
                        }
                        let _ = ShowWindow(hwnd, SW_HIDE);
                        if let Ok(mut current) = CURRENT_MEDIA.lock() {
                            *current = None;
                        }
                    }
                }
                _ => {}
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

    /// A display to place on: 1000 by 800 at its top-left corner.
    fn bounds() -> ScreenBounds {
        ScreenBounds {
            left: 0,
            top: 0,
            right: 1000,
            bottom: 800,
        }
    }

    /// A waiting spinner holds the pointer the way a scrollable text preview does:
    /// inside the box it published, and only while it says it is holding — which is
    /// what keeps a pointer that drifts onto the spinner from dismissing the hover
    /// whose page is on its way.
    #[test]
    fn holds_the_pointer_over_a_waiting_spinner() {
        let spinner = (100, 100, 136, 136);
        *POINTER_HOLD_REGIONS.lock().expect("the published regions") = Some(vec![spinner]);

        WAITING_PREVIEW_HOLDING.store(true, Ordering::Release);
        assert!(preview_pointer_hold(118, 118), "inside the spinner's box");
        assert!(!preview_pointer_hold(99, 99), "outside it");

        // A published region with nothing holding the pointer is not a hold: it is
        // what one left behind by a preview that has gone would look like.
        WAITING_PREVIEW_HOLDING.store(false, Ordering::Release);
        assert!(
            !preview_pointer_hold(118, 118),
            "nothing is holding the pointer"
        );

        clear_pointer_hold();
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

    /// A placement kept off `name`, at the size the media's own scale allows — the
    /// arrangement the figures are easy to read in.
    fn placed(
        placement: PreviewLayout,
        media: (u32, u32),
        name: (i32, i32, i32, i32),
        bounds: ScreenBounds,
    ) -> PreviewLayout {
        avoiding_text(
            placement,
            media,
            PreviewScale::Percent(100),
            Some(name),
            20,
            bounds,
        )
    }

    /// A keyboard preview of a row is placed in the room past the region the `Avoid`
    /// setting keeps it off, so a row is cleared only as far as the setting asks: past
    /// every column at `Avoid Details`, and only past the name at `Avoid Filename`,
    /// where the columns drawn after the name are the room the preview takes.
    #[test]
    fn a_keyboard_rows_tail_begins_at_the_region_it_is_kept_off() {
        let row = (0, 100, 1000, 140);
        let media = (400, 300);

        let past_the_name = compute_keyboard_layout(
            row,
            media,
            false,
            Some((20, 104, 120, 136)),
            PreviewScale::Percent(100),
            bounds(),
        )
        .expect("a placement past the name");

        let past_every_column = compute_keyboard_layout(
            row,
            media,
            false,
            Some((20, 104, 900, 136)),
            PreviewScale::Percent(100),
            bounds(),
        )
        .expect("a placement past the row's columns");

        assert_eq!(past_the_name.pos_x, 130, "just past the name");
        assert_eq!(past_every_column.pos_x, 910, "just past the columns");
        assert!(
            past_the_name.pos_x < past_every_column.pos_x,
            "a narrower region leaves more of the row to be covered"
        );
    }

    /// With nothing kept off — `Don't Avoid` — a row is placed by the position mode
    /// alone: it is anchored at its middle, the way a hover over it is read, and the
    /// preview is allowed to cover it.
    #[test]
    fn a_keyboard_row_with_nothing_kept_off_is_placed_by_position_alone() {
        let layout = compute_keyboard_layout(
            (0, 100, 1000, 140),
            (400, 300),
            false,
            None,
            PreviewScale::Percent(100),
            bounds(),
        )
        .expect("a placement");

        assert_eq!(layout.pos_x, 510, "half the row's width, and the gap");
    }

    /// A page — a PDF's, or one Office rendered — is drawn at the room the display
    /// has, because the room is free quality there. Every setting at or above
    /// `100%` asks for at least that room, so they are one setting for a page; only
    /// a setting below it is a size the user picked, and it is answered by
    /// reducing the fitted size rather than by ignoring it.
    #[test]
    fn a_page_takes_the_room_the_display_has() {
        let pdf = PathBuf::from(r"C:\docs\report.pdf");
        let svg_scale = PreviewScale::Percent(DEFAULT_SVG_SCALE_PERCENT);

        for configured in [
            PreviewScale::FitToScreen,
            PreviewScale::Percent(100),
            PreviewScale::Percent(200),
            PreviewScale::Percent(400),
        ] {
            assert_eq!(
                effective_preview_scale(&pdf, configured, svg_scale),
                PreviewScale::FitToScreen
            );
        }

        assert_eq!(
            effective_preview_scale(&pdf, PreviewScale::Percent(50), svg_scale),
            PreviewScale::FitToScreenReduced(50)
        );
        assert_eq!(
            effective_preview_scale(&pdf, PreviewScale::Percent(25), svg_scale),
            PreviewScale::FitToScreenReduced(25)
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

        assert_eq!(
            effective_preview_scale(&svg, configured, PreviewScale::FitToScreen),
            PreviewScale::FitToScreen
        );
        for percent in [75, 50, 25, 10] {
            assert_eq!(
                effective_preview_scale(&svg, configured, PreviewScale::Percent(percent)),
                PreviewScale::FitToScreenReduced(percent),
                "{percent}% of the room"
            );
        }

        assert_eq!(
            effective_preview_scale(&svg, configured, PreviewScale::Percent(60)),
            PreviewScale::FitToScreenReduced(60),
            "a share the menu does not offer is the share it is"
        );

        for configured in [
            PreviewScale::FitToScreen,
            PreviewScale::Percent(100),
            PreviewScale::Percent(400),
        ] {
            assert_eq!(
                effective_preview_scale(&svg, configured, PreviewScale::Percent(100)),
                PreviewScale::FitToScreen,
                "asking for the whole room or more is the whole room"
            );
        }
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
                picture,
                PreviewScale::Percent(DEFAULT_SVG_SCALE_PERCENT)
            ),
            PreviewScale::FitToScreenReduced(DEFAULT_SVG_SCALE_PERCENT)
        );

        // And a picture is still drawn at the picture scale, whatever the document
        // setting says.
        let png = PathBuf::from(r"C:\art\photo.png");
        assert_eq!(
            effective_preview_scale(&png, PreviewScale::Percent(200), PreviewScale::FitToScreen),
            PreviewScale::Percent(200)
        );
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
            PreviewScale::Percent(400),
            PreviewScale::Percent(DEFAULT_SVG_SCALE_PERCENT),
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
        let alphas: Vec<u8> = frame.chunks_exact(4).map(|pixel| pixel[3]).collect();
        assert!(alphas.iter().any(|&alpha| alpha > 200), "the arc is drawn");
        assert!(
            alphas.iter().any(|&alpha| (1..=200).contains(&alpha)),
            "the halo and the tail fade rather than fill"
        );
    }

    /// The waiting frame is the size of the spinner in it: over a turn of the
    /// spinner, the arc and its halo reach each of the box's edges, so a spinner
    /// placed flush at the pointer is the spinner at the pointer rather than an
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
            for (index, pixel) in frame.chunks_exact(4).enumerate() {
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

    /// A load that may be about to finish is given a moment before the spinner
    /// goes up, while one that came back waiting on a render is not: it has
    /// nothing to show, and the wait ahead of it is seconds of Office's time.
    #[test]
    fn puts_the_spinner_up_at_once_only_for_a_wait_that_is_known() {
        let load = |awaiting_render: bool, age: Duration, upgrade: bool| PendingLoad {
            generation: 1,
            path: PathBuf::new(),
            started: Instant::now() - age,
            pos_x: 0,
            pos_y: 0,
            width: 64,
            height: 64,
            spinner_shown: false,
            placement: None,
            awaiting_render,
            upgrade,
        };

        // A page on the way, a moment into the wait: the spinner is due already.
        assert!(load(true, Duration::from_millis(20), false).spinner_due());

        // A load that may be about to finish: not yet, and due once it has run
        // for the delay.
        assert!(!load(false, Duration::from_millis(20), false).spinner_due());
        assert!(load(false, Duration::from_secs(LOAD_SPINNER_DELAY_SECS), false).spinner_due());

        // An upgrade never: what is on screen stays until the page replaces it.
        assert!(!load(true, Duration::from_secs(10), true).spinner_due());

        // And a spinner that is already up is not put up a second time.
        let mut showing = load(false, Duration::from_secs(10), false);
        showing.spinner_shown = true;
        assert!(!showing.spinner_due());
    }

    /// A preview that is still on its way follows the pointer: it is placed again
    /// for a cursor that has moved along the item, kept where it is when the
    /// cursor has not moved, and left alone when the hover it came from was the
    /// keyboard's rather than the pointer's.
    #[test]
    fn a_pending_preview_follows_the_pointer() {
        let pending = |placement: Option<HoverPlacement>| PendingLoad {
            generation: 1,
            path: PathBuf::new(),
            started: Instant::now(),
            pos_x: 0,
            pos_y: 0,
            width: 0,
            height: 0,
            spinner_shown: true,
            placement,
            awaiting_render: false,
            upgrade: false,
        };
        let placement = HoverPlacement {
            orig_dims: (800, 600),
            avoid: None,
            follow_cursor: false,
            preview_scale: PreviewScale::FitToScreen,
            flush_at_cursor: false,
        };

        let mut pl = pending(Some(placement));
        assert!(
            pl.follow_pointer(POINT { x: 300, y: 300 }),
            "placed for the cursor"
        );
        let placed = (pl.pos_x, pl.pos_y, pl.width, pl.height);

        // The cursor has not moved, so the place has not changed: nothing to move.
        assert!(!pl.follow_pointer(POINT { x: 300, y: 300 }));
        assert_eq!((pl.pos_x, pl.pos_y, pl.width, pl.height), placed);

        // The cursor has moved: the preview is placed again, somewhere else.
        assert!(pl.follow_pointer(POINT { x: 200, y: 300 }));
        assert_ne!((pl.pos_x, pl.pos_y, pl.width, pl.height), placed);

        // A spinner box is the size the hover measured and stays that size while
        // the cursor moves along the item — the scale it was planned with is the
        // effective one, which for a document waiting on a render is `100%` — and
        // it is placed flush at the pointer's corner it follows.
        let mut waiting = pending(Some(HoverPlacement {
            orig_dims: (office_preview::WAITING_BOX, office_preview::WAITING_BOX),
            preview_scale: PreviewScale::Percent(100),
            flush_at_cursor: true,
            ..placement
        }));
        assert!(waiting.follow_pointer(POINT { x: 300, y: 300 }));
        assert_eq!(
            (waiting.pos_x, waiting.pos_y),
            (301, 301),
            "the spinner's own corner is at the cursor"
        );
        assert!(waiting.follow_pointer(POINT { x: 500, y: 300 }));
        assert_eq!(
            (waiting.width, waiting.height),
            (office_preview::WAITING_BOX, office_preview::WAITING_BOX)
        );

        // A keyboard hover's placement is the item's own: it follows nothing.
        let mut keyboard = pending(None);
        assert!(!keyboard.follow_pointer(POINT { x: 640, y: 400 }));
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
            flush_at_cursor: false,
        };

        let full = compute_mouse_layout(300, 300, page(PreviewScale::FitToScreen), bounds())
            .expect("a placed page");
        let half = compute_mouse_layout(
            300,
            300,
            page(PreviewScale::FitToScreenReduced(50)),
            bounds(),
        )
        .expect("a placed page");

        assert_eq!(
            (full.preview_w, full.preview_h),
            (half.preview_w * 2, half.preview_h * 2)
        );
    }

    /// The spinner a hover is waiting on is placed flush at the pointer's own
    /// corner — the one of the four the display has room for, a pixel off the
    /// pointer so the window is not under it — with no margin between the two and
    /// no step off the name it covers, so the wait stays at the hand that is
    /// waiting on it.
    #[test]
    fn places_the_waiting_spinner_flush_at_the_pointers_corner() {
        let side = office_preview::WAITING_BOX;
        let name = (100, 300, 400, 320);
        let spinner = |cursor_x: i32, cursor_y: i32| {
            compute_mouse_layout(
                cursor_x,
                cursor_y,
                HoverPlacement {
                    orig_dims: (side, side),
                    avoid: Some(name),
                    follow_cursor: false,
                    preview_scale: PreviewScale::Percent(100),
                    flush_at_cursor: true,
                },
                bounds(),
            )
            .expect("a placed spinner")
        };

        // Room in every quadrant: the spinner sits in the pointer's own corner,
        // over the name it is waiting on, rather than a gap away from the cursor.
        let placement = spinner(300, 300);
        assert_eq!((placement.pos_x, placement.pos_y), (301, 301));
        assert_eq!((placement.preview_w, placement.preview_h), (side, side));

        // With the display ending just past the pointer there is no room in the
        // quadrant it grows into, so the spinner takes the corner that is visible:
        // its box placed to the pointer's left, still touching it.
        let placement = spinner(980, 300);
        assert_eq!(
            (placement.pos_x, placement.pos_y),
            (980 - side as i32 - 1, 301)
        );
    }

    #[test]
    fn leaves_a_placement_that_is_already_clear_of_the_text() {
        // The name is drawn to the left of the cursor's column, so the preview beside
        // the cursor is already off it.
        let name = (100, 300, 400, 320);
        let placement = placed(layout(420, 300, 300, 300), (300, 300), name, bounds());

        assert_eq!((placement.pos_x, placement.pos_y), (420, 300));
    }

    #[test]
    fn moves_a_preview_out_of_the_name_it_covers() {
        // A row of a Details view: the pointer's item draws its name in a band twenty
        // pixels tall, and the preview came out beside the cursor with its top inside
        // that band. Down is the shortest way out, so it ends up just under the name,
        // where its own column already was.
        let name = (100, 300, 400, 320);
        let placement = placed(layout(120, 300, 300, 300), (300, 300), name, bounds());

        assert_eq!((placement.pos_x, placement.pos_y), (120, 340));
        assert_eq!((placement.preview_w, placement.preview_h), (300, 300));
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
        let placement = placed(layout(120, 300, 300, 300), (300, 300), name, short);

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
        let placement = placed(layout(0, 0, 1000, 800), (1000, 800), name, bounds());

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
        let placement = placed(layout(150, 80, 200, 90), (200, 90), name, tight);

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
        let placement = placed(layout(690, 100, 200, 300), (200, 300), name, narrow);

        // 720 is the name's right edge plus the gap, and 720 + 200 leaves the display;
        // 340 is its bottom plus the gap, and 340 + 300 does not.
        assert_eq!((placement.pos_x, placement.pos_y), (690, 340));
    }

    /// A planned preview is inside the display it was planned for: the box it takes is
    /// within the room it was given, and the whole of it — place and size, both axes —
    /// is within that display's work area.
    fn assert_inside_the_display(layout: PreviewLayout, bounds: ScreenBounds) {
        assert!(
            layout.preview_w <= layout.max_width && layout.preview_h <= layout.max_height,
            "the preview takes {} by {} of the {} by {} it was given",
            layout.preview_w,
            layout.preview_h,
            layout.max_width,
            layout.max_height
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

    /// The displays and the media every placement test is run over: a 1080p one, a 4K
    /// one, a 4K one that is the second display rather than the first, and a portrait
    /// one, against the shapes media comes in — pages both ways up, a slide, the
    /// waiting spinner, a picture, and things far wider and far taller than any display.
    fn displays_and_media() -> ([ScreenBounds; 4], [(u32, u32); 7], [PreviewScale; 4]) {
        let displays = [
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

        (displays, media, scales)
    }

    /// The one thing a hovered preview may never do: leave the display it was planned
    /// for. Every position mode, every scale, every shape of media, anchored at each
    /// corner of the display and at its middle, with a name under the pointer and with a
    /// row across the display's top to be kept off — because a preview that grows past
    /// its display takes an edge and a room that disagree to find, and the ones that
    /// disagree are not the ones anyone hovers over on purpose.
    #[test]
    fn places_every_hover_inside_its_display() {
        let (displays, media, scales) = displays_and_media();
        let mut placed = 0usize;

        for bounds in displays {
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
                                    // Only the waiting spinner is placed flush at the
                                    // pointer, and only it is ever that size.
                                    flush_at_cursor: (orig_width, orig_height) == (36, 36),
                                };

                                let Some(layout) =
                                    compute_mouse_layout(cursor_x, cursor_y, placement, bounds)
                                else {
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
        let (displays, media, scales) = displays_and_media();
        let mut placed = 0usize;

        for bounds in displays {
            let items = [
                // A row of a list, drawn across the view at its middle.
                (
                    bounds.left,
                    (bounds.top + bounds.bottom) / 2,
                    bounds.right,
                    (bounds.top + bounds.bottom) / 2 + 40,
                ),
                // A box item, the way Content view draws one.
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

                            for avoid in avoids {
                                let Some(layout) = compute_keyboard_layout(
                                    (item_left, item_top, item_right, item_bottom),
                                    (orig_width, orig_height),
                                    follow_cursor,
                                    avoid,
                                    preview_scale,
                                    bounds,
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

        assert!(
            placed >= 2000,
            "{placed} of the matrix's layouts were placed, which is too few to have checked the rest"
        );
    }

    /// The whole path a document that moves takes: worked out from its declarations,
    /// opened on drawn frames, and streamed from there. Ignored, and driven by
    /// `RHP_SVG_PROBE` — `$env:RHP_SVG_PROBE = "C:\art\spinner.svg"; cargo test -- --ignored --nocapture animated_svg_probe`
    /// — for a document whose animation does not play.
    #[test]
    #[ignore = "reads the file named in RHP_SVG_PROBE"]
    fn animated_svg_probe() {
        let Ok(path) = std::env::var("RHP_SVG_PROBE") else {
            println!("set RHP_SVG_PROBE to a path");
            return;
        };

        let path = PathBuf::from(path);
        let cancel = Arc::new(AtomicBool::new(false));
        let started = Instant::now();

        match load_animated_svg(
            &path,
            800,
            800,
            PreviewScale::FitToScreen,
            Arc::clone(&cancel),
        ) {
            Some(media) => {
                println!(
                    "opened on {} frame(s) of {}x{}, streamed from there, in {:?}",
                    media.frames.len(),
                    media.current_width(),
                    media.current_height(),
                    started.elapsed()
                );
            }
            None => println!("nothing played, in {:?}", started.elapsed()),
        }

        cancel.store(true, Ordering::Release);

        // A pass has to keep up with a frame every thirty-three milliseconds, so what
        // one frame costs is the number worth having here.
        let Some(source) = svg_preview::source(&path) else {
            return;
        };
        let Ok(document) = roxmltree::Document::parse(source.as_ref()) else {
            return;
        };
        let Some(playback) = svg_animation::Playback::parse(&document) else {
            return;
        };

        let started = Instant::now();
        let mut drawn = Vec::new();

        for index in 0..playback.frames().min(12) {
            let text = playback.document_at(&document, index);

            if let Some((pixels, _, _)) = svg_preview::render_text(&text, 800, 800) {
                drawn.push(pixels);
            }
        }

        if let Some(first) = drawn.first() {
            println!(
                "drew {} frame(s) in {:?} — {:?} each",
                drawn.len(),
                started.elapsed(),
                started.elapsed() / drawn.len() as u32
            );

            // A document that plays but does not move is a document whose declarations
            // were read and whose frames came out the same, which is the difference
            // between "nothing played" and "nothing moved".
            let moved = drawn.last().is_some_and(|last| last != first);

            println!("frames differ: {moved}");
        }
    }

    /// A document that moves is played: it opens on frames that were drawn before the
    /// preview was handed over, and the rest of the pass is streamed. One that stands
    /// still is not played at all, and is left to the still renderer.
    #[test]
    fn plays_a_document_that_moves_and_not_one_that_stands_still() {
        let folder = std::env::temp_dir()
            .join("rust-hover-preview-svg-tests")
            .join("animated");
        std::fs::create_dir_all(&folder).expect("a test folder");

        let moving = folder.join("spinner.svg");
        std::fs::write(
            &moving,
            br##"<svg xmlns="http://www.w3.org/2000/svg" width="40" height="40"><rect x="16" y="2" width="8" height="14" fill="#2b5fd9"><animateTransform attributeName="transform" type="rotate" from="0 20 20" to="360 20 20" dur="1s" repeatCount="indefinite"/></rect></svg>"##,
        )
        .expect("a written document");

        let cancel = Arc::new(AtomicBool::new(false));
        let media = load_animated_svg(
            &moving,
            200,
            200,
            PreviewScale::FitToScreen,
            Arc::clone(&cancel),
        )
        .expect("a document that moves");

        assert!(matches!(media.media_type, MediaType::AnimatedSvg));
        assert!(!media.frames.is_empty(), "the pass opens on drawn frames");
        assert!(
            media.frames[0].pixels.iter().any(|byte| *byte != 0),
            "the first frame was drawn"
        );
        assert!(
            media.shared_frames.is_some(),
            "the rest of the pass is streamed rather than drawn up front"
        );

        cancel.store(true, Ordering::Release);

        let still = folder.join("standing.svg");
        std::fs::write(
            &still,
            br##"<svg xmlns="http://www.w3.org/2000/svg" width="40" height="40"><rect width="40" height="40" fill="#2b5fd9"/></svg>"##,
        )
        .expect("a written document");

        let cancel = Arc::new(AtomicBool::new(false));
        assert!(
            load_animated_svg(
                &still,
                200,
                200,
                PreviewScale::FitToScreen,
                Arc::clone(&cancel)
            )
            .is_none(),
            "a document that stands still is not played"
        );
        cancel.store(true, Ordering::Release);

        let _ = std::fs::remove_file(&moving);
        let _ = std::fs::remove_file(&still);
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
            crate::webview_preview::is_available()
        );
        println!("document moves: {}", crate::webview_preview::moves(&path));

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
                    )
                })
            });

            println!(
                "{:>5} ms: engine_showing={} media={:?}",
                (step + 1) * 200,
                crate::webview_preview::is_showing(),
                media
            );

            if crate::webview_preview::is_showing() {
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
            crate::webview_preview::is_showing()
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

            let configured = current_preview_scale();
            let scale = effective_preview_scale(&path, configured, current_svg_scale());
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
                    flush_at_cursor: false,
                },
                bounds,
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
}

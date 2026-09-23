use crate::archive_formats;
use crate::archive_preview::{self, ArchivePreviewOptions};
use crate::cloud_files;
use crate::codecs;
use crate::config::{
    frame_bytes_within_budget, image_decode_limits, read_within_budget, sanitize_image_cache_mb,
    sanitize_spinner_delay_ms, sanitize_webp_playback_fps, MarkdownMode, PreviewScale, PreviewType,
    TextTheme, TransparentBackground, DEFAULT_ANIMATED_SCALE_PERCENT, DEFAULT_DDS_BACKGROUND,
    DEFAULT_DESIGN_BACKGROUND, DEFAULT_DESIGN_SCALE, DEFAULT_FONT_BACKGROUND, DEFAULT_FONT_SCALE,
    DEFAULT_IMAGE_BACKGROUND, DEFAULT_IMAGE_CACHE_MB, DEFAULT_LIBRE_SCALE, DEFAULT_OFFICE_SCALE,
    DEFAULT_PDF_SCALE, DEFAULT_PREVIEW_SCALE_PERCENT, DEFAULT_SPINNER_DELAY_MS,
    DEFAULT_TEXT_FONT_SCALE_PERCENT, DEFAULT_TEXT_SCROLL_FAR_EDGE_GRACE_PIXELS,
    DEFAULT_VECTOR_BACKGROUND, DEFAULT_VECTOR_SCALE, DEFAULT_VIDEO_SCALE_PERCENT,
    DEFAULT_WEBP_PLAYBACK_FPS,
};
use crate::dds_image;
use crate::design_formats;
use crate::engine_processes;
use crate::eps_image;
use crate::font_formats;
use crate::font_preview;
use crate::libre_formats;
use crate::libreoffice_render;
use crate::metafile_image;
use crate::office_formats;
use crate::office_preview;
use crate::office_render;
use crate::pdf_preview;
use crate::project_image;
use crate::psd_image;
use crate::svg_preview;
use crate::text_formats;
use crate::text_preview::{self, TextPreviewOptions};
use crate::tone_map;
use crate::vector_formats;
use crate::video_formats::{self, is_video_file};
use crate::video_player;
use crate::webp_image;
use crate::webview_preview;
use crate::wheel_input;
use crate::wic_image;
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
    Pdf,
    Text,
    Archive,
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
            Self::Text => Some(PreviewType::Text),
            Self::Pdf => Some(PreviewType::Pdf),
            Self::Archive => Some(PreviewType::Archives),
            Self::Office => Some(PreviewType::Office),
            // A design document is drawn into a frame like any picture, and the switch
            // over it is its own: the picture *is* what the file keeps of the document,
            // but a user who wants none of them is not asking for pictures to be off.
            Self::Design => Some(PreviewType::Design),
            // The same for a document an engine drew, at the gate over the engine's kind.
            Self::Libre => Some(PreviewType::Libre),
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

    fn current_width(&self) -> u32 {
        self.current_frame().map(|frame| frame.width).unwrap_or(0)
    }

    fn current_height(&self) -> u32 {
        self.current_frame().map(|frame| frame.height).unwrap_or(0)
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

/// Probe a video's geometry on a thread of its own, and tell the preview loop.
///
/// The probe is two external processes and the hover waits for the slower of them, so it
/// is done here rather than on the preview thread: what is on screen while it runs is the
/// waiting spinner, and the hover it belongs to is replayed when the answer lands (see
/// `video_probe_due` and `video_probe` in the preview loop). A probe whose hover has moved
/// on is not wasted — what it answers is held for the next hover of the file — so nothing
/// here is cancelled or waited for.
fn spawn_video_probe(path: PathBuf, generation: u64) {
    std::thread::spawn(move || {
        let _ = probe_video_geometry(&path);
        notify_video_probed(&path, generation);
    });
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
fn png_has_animation_control_chunk(path: &Path) -> bool {
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

fn is_apng_file(path: &Path) -> bool {
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

/// Whether a GIF holds more than its first frame, which is the whole of what makes
/// one an animation rather than a picture: the frames are the file's own blocks, and
/// whether there is a second one is a thing its structure says.
///
/// Nothing is decoded here. The blocks are walked by their own lengths — an
/// extension's sub-blocks are stepped over the same way, since they carry lengths
/// too — so what this costs is a few seeks and no pixels: the same rule the PNG
/// probe above follows, and the one the decoder follows when it does decode the
/// file for real.
fn gif_is_animated(path: &Path) -> bool {
    use std::io::{Read, Seek, SeekFrom};

    let Ok(file) = File::open(path) else {
        return false;
    };
    let mut reader = BufReader::new(file);

    // The signature and the logical screen descriptor: `GIF87a` or `GIF89a`, then
    // seven bytes of screen size, colour table and background.
    let mut header = [0u8; 13];
    if reader.read_exact(&mut header).is_err() || &header[..3] != b"GIF" {
        return false;
    }

    // A global colour table follows the descriptor, and the frame blocks after it:
    // `0x2C` opens one and carries its own bounds and local table, `0x21` opens an
    // extension whose sub-blocks carry lengths, and `0x3B` ends the file.
    if header[10] & 0x80 != 0 {
        let table_bytes = 3 * (1u64 << ((header[10] & 0x07) + 1));
        if reader.seek(SeekFrom::Current(table_bytes as i64)).is_err() {
            return false;
        }
    }

    let mut frames = 0usize;
    let mut block = [0u8; 1];

    while reader.read_exact(&mut block).is_ok() {
        match block[0] {
            0x2C => {
                frames += 1;
                if frames > 1 {
                    return true;
                }

                // The frame's bounds and flags: the local colour table, where it has
                // one, sits between the flags and the frame's pixels.
                let mut descriptor = [0u8; 9];
                if reader.read_exact(&mut descriptor).is_err() {
                    return false;
                }
                if descriptor[8] & 0x80 != 0 {
                    let table_bytes = 3 * (1u64 << ((descriptor[8] & 0x07) + 1));
                    if reader.seek(SeekFrom::Current(table_bytes as i64)).is_err() {
                        return false;
                    }
                }

                // The LZW code size byte, then the pixel data as sub-blocks.
                if reader.read_exact(&mut block).is_err() || !skip_gif_sub_blocks(&mut reader) {
                    return false;
                }
            }
            0x21 => {
                // An extension's label byte, then its own sub-blocks.
                if reader.read_exact(&mut block).is_err() || !skip_gif_sub_blocks(&mut reader) {
                    return false;
                }
            }
            0x3B => return false,
            // Anything else is not where the next block can be, so the walk is over.
            _ => return false,
        }
    }

    false
}

/// Step a GIF reader over one run of sub-blocks: a length byte per block, ending at
/// a zero-length one. The pixels and the extensions a specimen of either is skipped
/// by are both held this way, so the one walk serves both.
fn skip_gif_sub_blocks(reader: &mut BufReader<File>) -> bool {
    use std::io::{Read, Seek, SeekFrom};

    let mut length = [0u8; 1];
    loop {
        if reader.read_exact(&mut length).is_err() {
            return false;
        }
        if length[0] == 0 {
            return true;
        }
        if reader
            .seek(SeekFrom::Current(length[0] as i64))
            .is_err()
        {
            return false;
        }
    }
}

/// Whether a WebP file holds an animation, which its container says: an extended
/// WebP that animates carries an `ANIM` chunk, and its frames are `ANMF` chunks
/// after it.
///
/// The chunks are walked by their own lengths for the same reason the GIF's blocks
/// are: what is asked is what the file holds, and nothing has to be decoded to
/// answer it. A plain `VP8 ` or `VP8L` WebP has no chunk list at all — its picture
/// data is the chunk itself — so the walk stops where the picture starts.
fn webp_has_animation(path: &Path) -> bool {
    use std::io::{Read, Seek, SeekFrom};

    let Ok(file) = File::open(path) else {
        return false;
    };
    let mut reader = BufReader::new(file);

    // `RIFF`, the file's own length, and the form type every WebP carries.
    let mut header = [0u8; 12];
    if reader.read_exact(&mut header).is_err()
        || &header[..4] != b"RIFF"
        || &header[8..] != b"WEBP"
    {
        return false;
    }

    let mut chunk = [0u8; 8];
    while reader.read_exact(&mut chunk).is_ok() {
        let length = u32::from_le_bytes([chunk[4], chunk[5], chunk[6], chunk[7]]) as i64;

        if &chunk[..4] == b"ANIM" {
            return true;
        }

        // The picture data is the end of anything worth walking: an extended WebP
        // that animates names its animation ahead of its frames, and a plain one is
        // nothing but the picture.
        if &chunk[..4] == b"VP8 " || &chunk[..4] == b"VP8L" {
            return false;
        }

        // Chunks are padded to an even length.
        if reader
            .seek(SeekFrom::Current(length + length % 2))
            .is_err()
        {
            return false;
        }
    }

    false
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
fn codec_dimensions(path: &PathBuf) -> Option<(u32, u32)> {
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
            page: cfg.pdf_scale,
            office: cfg.office_scale,
            font: cfg.font_scale,
            design: cfg.design_scale,
            libre: cfg.libre_scale,
            vector: cfg.vector_scale,
        })
        .unwrap_or(HoverScales {
            picture: PreviewScale::Percent(DEFAULT_PREVIEW_SCALE_PERCENT),
            video: PreviewScale::Percent(DEFAULT_VIDEO_SCALE_PERCENT),
            animated: PreviewScale::Percent(DEFAULT_ANIMATED_SCALE_PERCENT),
            page: DEFAULT_PDF_SCALE,
            office: DEFAULT_OFFICE_SCALE,
            font: DEFAULT_FONT_SCALE,
            design: DEFAULT_DESIGN_SCALE,
            libre: DEFAULT_LIBRE_SCALE,
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

    match office_render::cached_render(path) {
        Some(cached) => office_render::page_is_narrower_than(&cached, width),
        None => true,
    }
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

/// Whether this hover is owed a page by the render engine: a document the `[libre]` list
/// holds, an engine installed to draw it, and no page drawn for this version of it yet.
///
/// It is the question `office_render_is_due` asks of an Office document, asked of the
/// documents whose engine is a whole application rather than an automation server. Three
/// things ask it: the layout, which measures a document like this as the wait for a page;
/// the loader, which answers with it that a hover is still waiting rather than failed; and
/// the loop, which asks the engine for the page only where there is one to ask for.
fn libre_render_is_due(path: &Path) -> bool {
    libre_formats::is_libre_preview(path)
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
    /// The share of the display a PDF page is drawn at.
    page: PreviewScale,
    /// The share of the display a document an engine drew is shown at: what the engine hands
    /// back is a page, so the share is of the room the display has rather than of a size the
    /// file asks for, exactly as a PDF page's is.
    libre: PreviewScale,
    /// The share of the display the page an Office document is drawn as is shown at.
    office: PreviewScale,
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
/// `pdf_scale`'s to say, and it says the whole of it unless it is asked for less.
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
/// A page Office rendered is the PDF rule again: it is drawn at whatever size it is
/// asked for, at the share of the room `office_scale` names. The one source that is not
/// a page is the bitmap a workbook is answered with where no page can be exported, and
/// it follows the share the way `bitmap_at_display_scale` reads it.
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
    if pdf_preview::is_pdf_file(path) {
        fit_reduced(scales.page)
    } else if is_text_preview(path) || archive_formats::is_archive_file(path) {
        PreviewScale::Percent(100)
    } else if video_probe_due(path) {
        // A video that has not been probed yet is a hover that is waiting, and what is on
        // screen for one is the waiting spinner: a wait is placed at the size it is rather
        // than fitted to the display, and what the probe answers is what the replay that
        // follows it is laid out at (see `video_probe_due`).
        PreviewScale::Percent(100)
    } else if office_formats::is_office_file(path) {
        // A page Office rendered is vector, so the room the display has is free
        // quality — the rule a PDF follows, at the share `office_scale` names. The
        // raster picture a workbook is answered with where no printer can export a
        // page is the exception: it is only as good as the pixels it holds, so it
        // follows the configured share the way an image does rather than being
        // enlarged to fit.
        match office_preview::source_kind(path) {
            // Nothing to draw and a page on the way: what is on screen is the
            // spinner in a box of its own, and a box that small is placed at the
            // size it is rather than fitted to the display the way a page is.
            office_preview::SourceKind::None => PreviewScale::Percent(100),
            source if source.may_be_enlarged() => fit_reduced(scales.office),
            _ => bitmap_at_display_scale(scales.office),
        }
    } else if libre_formats::is_libre_file(path) {
        // A document an engine draws follows a scale of its own, and the share is of the
        // display the way a PDF page's is: what the engine hands back is a page, not a
        // picture with a size of its own to be scaled from.
        fit_reduced(scales.libre)
    } else if design_formats::is_design_file(path) {
        // A design document is a document for this question rather than a picture: what is
        // previewed is the picture the file keeps of the whole of itself, at whatever size
        // that is, so the share is of the display the way a page's or a specimen's is.
        //
        // It is asked ahead of the kind below it because the gate asks it ahead of that
        // kind as well: a name written into both lists is a design document, and the share
        // the vector list would give it is not the share its own setting names.
        fit_reduced(scales.design)
    } else if svg_preview::is_svg_file(path) || vector_formats::is_vector_file(path) {
        // Both halves of the kind, asked as one: what a document costs to draw and what a
        // drawing costs to replay are the same question, and the setting is the same one.
        fit_reduced(scales.vector)
    } else if font_formats::is_font_file(path) {
        fit_reduced(scales.font)
    } else if is_video_file(path) {
        scales.video
    } else {
        // The animated arm is asked last because asking it is the one thing here that
        // reads the file, and a file that has already answered as another kind never
        // pays for it (see `image_is_animated`).
        animated_scale_for(path, scales).unwrap_or(scales.picture)
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

/// Whether a picture file holds an animation rather than a single frame: a GIF with
/// more than one frame, a WebP with an animation chunk, or a PNG with an animation
/// control chunk.
///
/// Nothing is decoded, and the question is asked of the file's own structure rather
/// than of its name alone: a still `.gif` and an animated `.png` are both files a name
/// cannot settle, which is the same reason the loader asks the file and not its
/// extension what it is. A file that is none of the three formats — whatever it is
/// called — is answered by the bytes its reader sees, and a format whose containers do
/// not animate, a JPEG or a `.bmp`, is answered without opening anything at all.
fn image_is_animated(path: &Path) -> bool {
    let extension = path
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_ascii_lowercase());

    match extension.as_deref() {
        Some("gif") => gif_is_animated(path),
        Some("webp") => webp_has_animation(path),
        Some("apng") => true,
        Some("png") => png_has_animation_control_chunk(path),
        _ => false,
    }
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
/// Every kind the gate asks about before the text lists is asked about here too, and for
/// the reason the gate asks them first: a file is whichever kind claims it, and a file the
/// hook claimed as a drawing or a document has been measured, laid out and gated as that
/// kind by the time this is asked. A name written into the text list as well as an earlier
/// list is that earlier kind, and rendering it as text would be a preview of another kind
/// than the one the tray was asked to switch.
///
/// The video gate is the one that settles the extensions the text list shares with it —
/// `.ts` and `.mts` — by content, since only it can tell a TypeScript source from the
/// transport stream that goes by the same name; the rest are bare names, so asking them
/// costs a lookup each and no read of the file.
fn is_text_preview(path: &Path) -> bool {
    if !text_formats::is_text_file(path) {
        return false;
    }

    // Every kind the hook asks ahead of the text lists is asked ahead of them here too, so
    // that a name in two lists is measured as the kind the hook called it: the text lists
    // reach further than the others — a `.md` in the `[libre]` list is a Markdown document
    // the engine would be asked to draw — and a preview measured as text would be placed
    // as one and drawn as the other. A name in the video, PDF, archive or office list is
    // excluded the same way, and `libre` is the newest of them.
    !is_video_file(path)
        && !pdf_preview::is_pdf_file(path)
        && !archive_formats::is_archive_file(path)
        && !office_formats::is_office_file(path)
        && !libre_formats::is_libre_file(path)
        && !design_formats::is_design_file(path)
        && !vector_formats::is_vector_file(path)
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
        let img = if is_confirm_file_type_enabled() {
            decode_image_with_header_check(path)?
        } else {
            decode_image_by_extension(path)?
        };

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

    let frame = ImageFrame {
        pixels,
        width: target_width,
        height: target_height,
        delay_ms: 0,
    };

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

    let frame = ImageFrame {
        pixels,
        width,
        height,
        delay_ms: 0,
    };

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
    let (target_width, target_height) =
        scale_dimensions(page_width, page_height, max_width, max_height, preview_scale);
    let (pixels, width, height) = pdf_preview::render_first_page(page, target_width, target_height)?;

    Some(static_image_media(
        ImageFrame {
            pixels,
            width,
            height,
            delay_ms: 0,
        },
        kind,
    ))
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

    let frame = ImageFrame {
        pixels,
        width: target_width,
        height: target_height,
        delay_ms: 0,
    };

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

    VIDEO_GEOMETRY_CACHE.lock().ok()?.get(&key).copied()
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

    if let Ok(cache) = VIDEO_GEOMETRY_CACHE.lock() {
        if let Some(cached) = cache.get(&key) {
            return *cached;
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

    // A file FFmpeg is not there for — or one its own probe could not read — is asked of
    // the media engine Windows has, which is also the engine that would play it. That is
    // the whole of the fallback's geometry: there is no crop to detect, because cropdetect
    // is an FFmpeg filter and the engine is handed the frame as the file holds it.
    let Some((src_w, src_h)) = dimensions.or_else(|| video_player::dimensions(path)) else {
        if let Ok(mut cache) = VIDEO_GEOMETRY_CACHE.lock() {
            if !cache.contains_key(&key) && cache.len() >= VIDEO_GEOMETRY_CACHE_MAX_ENTRIES {
                cache.clear();
            }
            cache.insert(key, ProbedGeometry::Unmeasurable);
        }

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

    if let Ok(mut cache) = VIDEO_GEOMETRY_CACHE.lock() {
        if !cache.contains_key(&key) && cache.len() >= VIDEO_GEOMETRY_CACHE_MAX_ENTRIES {
            cache.clear();
        }
        cache.insert(key, ProbedGeometry::Measured(geometry));
    }

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

/// Stop video playback process
fn stop_video_playback(media: &mut MediaData) {
    // A video the media engine is playing has no process and no window of its own: letting
    // the engine go is the whole of stopping it, and it is done here because this is where
    // every path that ends a video already comes through — the pointer leaving the file,
    // another preview taking its place, the `Videos` gate closing, a display change, a
    // resume from sleep, and the app itself.
    if media.media_type.is_native_video() {
        video_player::stop();
    }

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
        // Where no Office is installed to draw a page, the render engine beside it draws one
        // instead: the same page, shown as an Office document rather than as a document of
        // the engine's own kind — the file is what it is, whichever engine drew it.
        return load_office_preview(path, max_width, max_height, preview_scale, &cancel).or_else(
            || {
                let page = libreoffice_render::pdf_for_office(path)?;
                load_engine_page(
                    &page,
                    MediaType::Office,
                    max_width,
                    max_height,
                    preview_scale,
                )
            },
        );
    }

    // A design document is read for the picture its own format keeps of the whole
    // thing, and it is asked where the hook asks it: after the office list, ahead of the
    // text lists and the picture path — neither of which would have claimed one of these
    // names anyway. What the reader refuses is refused outright rather than falling
    // through to the decoder the picture path ends in, because that decoder is for the
    // names it names and this is not one of them.
    //
    // The gate is not asked here, exactly as it is not asked for a PDF or a font: the
    // hook asks it before a hover can reach this path at all.
    // A document this app hands to a render engine is the engine's to draw, and there is
    // nothing behind it: what such a file keeps of itself is a thumbnail, and a thumbnail is
    // not shown — see `libre_formats` — so a machine without the engine shows nothing at all.
    if libre_formats::is_libre_file(path) {
        // The page is the engine's to draw and it is not drawn here: what is loaded is the
        // page that engine has already written. A document without one is a wait rather
        // than a failure — the loop has asked for it, and the hover is replayed when the
        // page lands (see `libre_render_is_due`) — so what it gets here is nothing, which
        // is the spinner it is already showing.
        return libreoffice_render::rendered_page(path).and_then(|page| {
            load_engine_page(
                &page,
                MediaType::Libre,
                max_width,
                max_height,
                preview_scale,
            )
        });
    }

    if design_formats::is_design_file(path) {
        return load_design_preview(path, max_width, max_height, preview_scale);
    }

    // A vector drawing is the drawing layer's to replay rather than a decoder's to read,
    // and it is asked beside the design documents for the same reason they are: what it
    // is, is its own header's answer rather than its name's.
    //
    // The list names both halves of the kind, though, so the document half is asked here
    // first: an `.svg` is an entry of the vector list beside the metafiles, and neither
    // reader below would take one — the hover would show nothing at all, which is not
    // what a document the gate has already claimed may come to.
    if vector_formats::is_vector_file(path) {
        return if svg_preview::is_svg_file(path) {
            webview_preview::draws(path).then(engine_svg_media)
        } else {
            load_vector_preview(path, max_width, max_height, preview_scale)
        };
    }

    if text_formats::is_text_file(path) {
        return load_text_preview(path, max_width, max_height, dpi, current_text_options());
    }

    // A document is the engine's to draw — this app rasterizes none of them — so there is
    // nothing to make here: what comes back is the kind alone, and the install path reads
    // it as the hover to hand over. An engine that cannot draw it is no preview, which is
    // the answer a file that will not decode gets.
    //
    // It is asked here for a document the vector list does not name and the image list
    // does, which is where the hook asks the same question: its own document check sits
    // inside the image block, after the list that would have claimed a picture. The text
    // lists are asked ahead of it for the reason the hook asks them there too — a name a
    // user has put in the text list is a text file.
    if svg_preview::is_svg_file(path) {
        return webview_preview::draws(path).then(engine_svg_media);
    }

    // A font is drawn the same way, and asks two questions before it is: whether the engine
    // can draw anything at all, and whether the file is a font this side can describe. A
    // `.ttf` holding something else is answered with nothing rather than with a page of
    // another font's glyphs, which is what the probe settles — and what it answers is what
    // the specimen is made of, so the second question is asked first.
    if font_formats::is_font_file(path) {
        return (font_preview::probe(path).is_some() && webview_preview::draws(path))
            .then(engine_font_media);
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
        // What is left is a still, and a still is not decoded here either: it is the
        // picture the codec Windows has for WebP draws, at the box the layout planned,
        // which is what `load_static_image` asks for; see `wic_image`.
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
    video_formats::is_video_preview(path) && cached_video_geometry(path).is_none()
}

/// Get original dimensions of media for positioning calculations
fn get_media_dimensions(path: &PathBuf) -> Option<(u32, u32)> {
    if video_formats::is_video_preview(path) {
        return video_box(path);
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

        return Some((office_preview::WAITING_BOX, office_preview::WAITING_BOX));
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

        return svg_preview::measure(path);
    }

    // A vector drawing is measured from the records it holds: what an `.eps` keeps a
    // preview of, or what a metafile's own header declares its drawing to be.
    if vector_formats::is_vector_preview(path) {
        return vector_dimensions(path);
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

        return font_preview::probe(path)
            .map(|_| (font_preview::SPECIMEN_WIDTH, font_preview::SPECIMEN_HEIGHT));
    }

    // Whatever is left is a picture, so the `Images` gate is what decides it.
    if !PreviewType::Images.enabled() {
        return None;
    }

    if is_confirm_file_type_enabled() {
        image_dimensions_with_header_check(path)
    } else {
        // The same two readers, asked the way this path asks them: the name first,
        // since a file whose content is not confirmed is taken for what it is called,
        // and then the codec Windows has — which is the only reader there is for the
        // picture formats this app's decoder cannot read at all.
        image::image_dimensions(path)
            .ok()
            .or_else(|| codec_dimensions(path))
    }
}

/// Whether this hover is the wait for a page rather than a preview of one: an Office
/// document or a document the render engine draws, with nothing drawn for it yet and a page
/// on the way.
///
/// A hover like that is placed by the spinner's own box, flush at the pointer, rather than
/// by the size a preview would take: it is the wait for the file under the hand, which
/// belongs at the hand. The page is laid out again by the replay that arrives with it, so
/// nothing here has to guess how large it will be.
fn page_is_on_the_way(path: &Path) -> bool {
    let office = office_formats::is_office_preview(path)
        && matches!(
            office_preview::source_kind(path),
            office_preview::SourceKind::None
        );

    office || libre_render_is_due(path)
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
/// font size is scaled by — and what every margin a layout is written around is
/// scaled by (see `logical_px`). Falls back to the 96 DPI baseline when no display
/// can be named, the same way the placement falls back to the primary display.
///
/// The display is what is asked, not the window the point happens to be over: a
/// window carries the scale its own process was told about — a UWP one can answer a
/// scale that is not the display's at all — while the display under the point is one
/// question with one answer, whatever is drawn on it.
pub(crate) fn monitor_dpi_from_point(x: i32, y: i32) -> u32 {
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

            let awaiting_render = media.is_none()
                && (office_render_is_due(&request.path, request.max_width)
                    || libre_render_is_due(&request.path));

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
#[derive(Clone, Copy)]
struct HoverPlacement {
    orig_dims: (u32, u32),
    avoid: Option<ScreenRegion>,
    follow_cursor: bool,
    preview_scale: PreviewScale,
    /// Whether this placement is the waiting spinner's own box rather than a
    /// preview's: it is then placed flush at the pointer's own corner and kept there
    /// while the wait runs. See `compute_mouse_layout` and `waiting_placement`.
    flush_at_cursor: bool,
}

/// The placement a hover's waiting spinner is given: the arc's own box, flush at the
/// pointer's corner, whatever the preview it is waiting for is.
///
/// It is the placement a document waiting on a page has always been placed by, and
/// every other kind of load is shown the same way: what a hover shows while it waits
/// is the arc at the hand that hovered the file — which is what says the wait is for
/// the file under it — rather than a preview-sized frame with a spinner in the middle
/// of it, placed where the preview will land and saying nothing about the hand.
fn waiting_placement(placement: HoverPlacement) -> HoverPlacement {
    HoverPlacement {
        orig_dims: (office_preview::WAITING_BOX, office_preview::WAITING_BOX),
        preview_scale: PreviewScale::Percent(100),
        flush_at_cursor: true,
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
    /// The file this load is for, which is what decides whether the engine plays it
    /// and what the engine is pointed at.
    path: PathBuf,
    started: Instant,
    pos_x: i32,
    pos_y: i32,
    width: u32,
    height: u32,
    /// The room the layout allowed this preview, which is the box a page on its way
    /// is asked for in: a slide is exported at the width the render is asked for, so
    /// the room the display has is the sharpest page that display can show.
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

/// Put the loading spinner on screen for a pending load, in the box that load's wait
/// is placed in: the arc's own, at the pointer's corner (see `waiting_placement`).
///
/// A window that is not on screen is moved before the spinner is installed, so a
/// `WM_DPICHANGED` reset from crossing displays cannot discard it, and the spinner
/// is painted before the window is revealed, so the previous preview cannot flash
/// at the new place. It is also what moves a spinner whose box has changed size
/// while it was up: the frame is drawn at the size of the box it goes into.
unsafe fn show_loading_spinner(hwnd: HWND, pl: &PendingLoad) {
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
                    dpi,
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
const MIN_AVOID_ROOM_PIXELS: f32 = 64.0;

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
/// cleared by that much rather than touched at its edge. It is already in the pixels
/// of the display this placement is for, like the least room a way out is worth
/// taking, which is scaled here from the logical distance it is written as.
fn avoiding_text(
    layout: PreviewLayout,
    orig_dims: (u32, u32),
    preview_scale: PreviewScale,
    avoid: Option<ScreenRegion>,
    gap: i32,
    bounds: ScreenBounds,
    dpi: u32,
) -> PreviewLayout {
    let Some((text_left, text_top, text_right, text_bottom)) = avoid else {
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

/// How far off the pointer a preview is placed, in logical pixels: the margin the
/// position modes are written around.
const POINTER_STANDOFF_PIXELS: f32 = 20.0;

/// How far off the pointer the waiting spinner is placed, in logical pixels: a
/// single pixel, because the preview window is what a mouse message at the pointer
/// lands on, and the pointer has to keep clicking and probing the file it is waiting
/// on rather than the spinner that is waiting on it.
const SPINNER_STANDOFF_PIXELS: f32 = 1.0;

/// Compute preview layout for mouse hover (relative to cursor position)
///
/// `placement` is what the hover asks for — the size the preview was measured at,
/// the text to keep it off, the position mode it follows and the scale it is drawn
/// with — the same reading a pending load keeps to place its preview again as the
/// pointer moves (see `HoverPlacement`). `dpi` is the display the pointer is on,
/// which is what the margins the placement is written around are scaled by.
///
/// A placement that is the waiting spinner (`flush_at_cursor`) is placed by its own
/// rule: the corner nearest the pointer is put at the pointer — a single logical
/// pixel off it, which is what `offset` below is — in whichever of the four
/// quadrants the display has room for the spinner, and the name it covers is not
/// stepped around: a spinner waiting on the page of the file under the hand says
/// what it is by being at the hand, and one placed a row away from it says nothing
/// about what is being waited on. Every other preview keeps the margin its position
/// mode leaves and the room the `Avoid` setting asks for.
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
        flush_at_cursor,
    } = placement;

    let offset = if flush_at_cursor {
        logical_px(dpi, SPINNER_STANDOFF_PIXELS)
    } else {
        logical_px(dpi, POINTER_STANDOFF_PIXELS)
    };
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
            avoid,
            offset,
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
    /// when nothing is kept off, or the view reported no text for the item.
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
    // with nothing kept off at all are both left to the placement below, which anchors
    // them at their middle: there is nowhere beside such a row to put a preview, and
    // the display's own room is all there is.
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
                avoid,
                gap,
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
            avoid,
            gap,
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
            avoid,
            gap,
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
            // is on screen is the document, and what is shown next is another hover's.
            if webview_preview::is_showing() {
                if IsWindowVisible(hwnd).as_bool() {
                    let _ = ShowWindow(hwnd, SW_HIDE);
                }

                if pending_load.take().is_some() {
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
                        Some(media_data) => {
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

                            let mw = media_data.current_width() as i32;
                            let mh = media_data.current_height() as i32;

                            // A load whose spinner was up is placed like any other:
                            // the wait was shown in the spinner's own box at the
                            // pointer, not in the box the preview arrives in (see
                            // `spinner_pos`), so the frame is installed at the
                            // preview's place — the paint below is what takes the
                            // window from one box to the other.
                            let pending = pending_load
                                .take()
                                .filter(|pl| pl.generation == result.generation);
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
                            // (see `spinner_due`). The page is asked for in the room
                            // this preview may take, which is what the page is drawn
                            // at the size of: the spinner's own box is a spinner's
                            // and says nothing about how large the page will be
                            // drawn, while a slide is exported at the width the
                            // render is asked for — so the room the display has is
                            // the sharpest page that display can show.
                            let (width, height) = pending_load
                                .as_ref()
                                .map(|pl| pl.room)
                                .unwrap_or_else(|| office_formats::default_page_size(&result.path));
                            // A document the render engine draws is asked for the same
                            // way, and in the same breath: neither page exists until an
                            // engine has drawn it, and this is the one hover that is
                            // waiting for one.
                            page_render_pending = request_office_render(
                                &result.path,
                                result.generation,
                                width,
                                height,
                            )
                            .or_else(|| request_libre_render(&result.path, result.generation));
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

                    // A wait for an engine-drawn preview takes the engine's window with
                    // it: the same file asked for again is that window moved rather
                    // than the page navigated to again, so what it is about to draw in
                    // is where the wait ended up. Nothing is asked of an engine that
                    // already has it up — that is the preview itself, and it follows
                    // nothing.
                    if followed.preview
                        && pl.spinner_shown
                        && engine_kind_of(&pl.path).is_some()
                        && !webview_preview::is_showing()
                    {
                        webview_preview::show(
                            &pl.path,
                            webview_preview::Area {
                                x: pl.pos_x,
                                y: pl.pos_y,
                                width: pl.width as i32,
                                height: pl.height as i32,
                            },
                            engine_background(&pl.path),
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
            // A page an engine has finished with, held apart from the hovers: it is
            // not a hover to act on but an answer about the one on screen. Office's
            // tier is the one that sends it; a page the render engine draws is read
            // from the folder it lands in, below, and lands in the same place.
            let mut page_ready: Option<(PathBuf, u64, bool)> = None;
            // A probe's answer, held apart the same way and for the same reason.
            let mut video_probed: Option<(PathBuf, u64)> = None;
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
            // same code below takes it up: it is the same wait, in the same box.
            if page_ready.is_none() {
                if let Some((path, generation)) = page_render_pending
                    .as_ref()
                    .filter(|(path, _)| libre_formats::is_libre_file(path))
                {
                    let drawn = libreoffice_render::rendered_page(path).is_some();
                    let refused = libreoffice_render::refused(path);

                    if drawn || refused {
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

                if hovered && ready_ok {
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
                    page_render_pending = None;
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
                    page_render_pending = None;
                    // The hover this page was rendered for is over: it landed after
                    // the pointer had moved on, so nothing is waiting for it. What was
                    // rendered is kept as far as the budget allows and no further — at
                    // a size of nothing it is dropped here rather than held for a
                    // hover that has already gone.
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

            if let Some(preview_msg) = latest_preview_msg {
                // Common variables for Show/ShowKeyboard - set in match, used after
                let mut show_path: Option<PathBuf> = None;
                let mut show_layout: Option<PreviewLayout> = None;
                let mut show_spinner_layout: Option<PreviewLayout> = None;
                let mut show_placement: Option<HoverPlacement> = None;
                let mut show_is_video: bool = false;
                // Whether this hover is waiting on a video's probe rather than on the
                // video itself (see `video_probe_due`).
                let mut show_video_probe: bool = false;
                let mut show_requested = false;
                let mut preview_scale = current_hover_scales().picture;
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
                        preview_scale = effective_preview_scale(&path, current_hover_scales());

                        // A document with no page rendered for it yet has nothing to
                        // measure but the wait, so its preview is laid out as the
                        // spinner's own box (see `office_preview::measure` and the
                        // `libre` arm of `get_media_dimensions`) and that box is placed
                        // flush at the pointer rather than a margin away from it: the
                        // page it waits for is laid out again by the replay that arrives
                        // with it, so until then the hover is the wait for the file under
                        // the hand, which belongs at the hand.
                        let waiting_spinner = page_is_on_the_way(&path);

                        // A video whose shape the probe has not answered for yet is the
                        // same kind of wait, and for the same reason: there is nothing to
                        // lay out as a video until the probe answers, so the hover is the
                        // wait for it — the spinner at the pointer — and is replayed when
                        // the answer lands (see `video_probe_due`).
                        let probing = video_probe_due(&path);

                        if let Some(orig_dims) = media_dimensions(&path, bounds, dpi) {
                            let is_video = is_video_file(&path);
                            let placement = HoverPlacement {
                                orig_dims,
                                avoid,
                                follow_cursor,
                                preview_scale,
                                flush_at_cursor: waiting_spinner || probing,
                            };
                            let placed = compute_mouse_layout(x, y, placement, bounds, dpi);
                            if let Some(layout) = placed {
                                let layout = text_preview_layout(&path, layout, dpi, |size| {
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
                                show_is_video = is_video;
                                show_video_probe = probing;
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
                            let is_video = is_video_file(&path);
                            let placement = KeyboardPlacement {
                                item_rect: (il, it, ir, ib),
                                avoid,
                                columns,
                                orig_dims,
                                follow_cursor,
                                preview_scale,
                            };
                            if let Some(layout) = compute_keyboard_layout(placement, bounds, dpi) {
                                let layout = text_preview_layout(&path, layout, dpi, |size| {
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
                    // video's probe is the same thing reached another way: it is the
                    // wait it was already in, carried on.
                    let upgrading = page_upgrade.as_deref() == Some(path.as_path())
                        || video_replay.as_deref() == Some(path.as_path());
                    page_upgrade = None;
                    video_replay = None;

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

                    // A video is played by `ffplay` when FFmpeg is installed, and by the
                    // media engine Windows has when it is not. The two take different
                    // roads from here: FFmpeg's player is its own window, which is what
                    // the branch below puts up, while the engine's frames come back
                    // through the ordinary load and are drawn by this app's own window.
                    let ffplay_plays_video = show_is_video && codecs::ffplay_available();

                    if show_video_probe {
                        // The hover is waiting on the probe: nothing of the file can be
                        // laid out or loaded until there is a shape to lay it out with,
                        // so what is put up is the wait every other preview is given —
                        // the spinner at the pointer — and the hover it came from is
                        // replayed when the answer lands, which is when there is a video
                        // to load (see `video_probe_due` and `video_probe`). The probe
                        // itself runs on a thread of its own: it is two external
                        // processes, and this is the thread that draws the wait.
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
                            path: path.clone(),
                            started,
                            pos_x,
                            pos_y,
                            width: preview_w,
                            height: preview_h,
                            room: (max_width, max_height),
                            spinner_shown,
                            spinner_delay: load_spinner_delay(),
                            spinner_pos: (spinner_x, spinner_y),
                            spinner_side,
                            placement: show_placement,
                            upgrade,
                        });
                        video_probe = Some((path.clone(), gen));
                        spawn_video_probe(path, gen);
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
                                        path: path.clone(),
                                        started: Instant::now(),
                                        pos_x,
                                        pos_y,
                                        width: media_width as u32,
                                        height: media_height as u32,
                                        room: (max_width, max_height),
                                        spinner_shown: false,
                                        spinner_delay: Duration::ZERO,
                                        spinner_pos: (spinner_x, spinner_y),
                                        spinner_side,
                                        placement: show_placement,
                                        upgrade: false,
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
                            room: (max_width, max_height),
                            spinner_shown: false,
                            spinner_delay: load_spinner_delay(),
                            spinner_pos: (spinner_x, spinner_y),
                            spinner_side,
                            placement: show_placement,
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

            let render_wait = page_render_pending.as_ref().map(|(path, generation)| {
                let hovered = *generation == current_generation
                    && shown_path.as_deref() == Some(path.as_path());
                let waited = pending_load
                    .as_ref()
                    .map(|pl| pl.started.elapsed() >= Duration::from_secs(OFFICE_RENDER_WAIT_SECS))
                    .unwrap_or(false);

                (hovered, waited)
            });

            match render_wait {
                Some((false, _)) => page_render_pending = None,
                Some((true, true)) => {
                    // The preview has waited as long as it waits — a page that has
                    // not arrived by now may still be coming, a very large document
                    // takes as long as it takes — so the spinner comes down and the
                    // hover is left to itself. The render is not abandoned with it:
                    // it runs on and its page is cached, so the next hover of that
                    // file shows it. Only the engine itself can say a render failed,
                    // and it remembers that for the file.
                    let abandoned = page_render_pending
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
    use crate::config::DEFAULT_FONT_SCALE_PERCENT;

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
    /// the gap is the pointer's standoff at this display's scale.
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
            logical_px(TEST_DPI, POINTER_STANDOFF_PIXELS),
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
            page: DEFAULT_PDF_SCALE,
            office: DEFAULT_OFFICE_SCALE,
            font: DEFAULT_FONT_SCALE,
            design: DEFAULT_DESIGN_SCALE,
            libre: DEFAULT_LIBRE_SCALE,
            vector: DEFAULT_VECTOR_SCALE,
        }
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
        push_png_chunk(&mut bytes, *b"IHDR", &[0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0]);

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
            let layout = compute_keyboard_layout(
                keyboard_placement(tile, avoid, false),
                bounds(),
                TEST_DPI,
            )
            .expect("a placement beside the tile");

            assert_eq!(layout.pos_x, tile.2 + gap, "just past the tile's own edge");
        }
    }

    /// With nothing kept off — `Avoid Nothing` — a row is placed by the position mode
    /// alone: it is anchored at its middle, the way a hover over it is read, and the
    /// preview is allowed to cover it.
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

    /// A page — a PDF's, or one Office rendered — is drawn at the room the display
    /// has, because the room is free quality there. Every setting at or above
    /// `100%` asks for at least that room, so they are one setting for a page; only
    /// a setting below it is a size the user picked, and it is answered by
    /// reducing the fitted size rather than by ignoring it. Which page setting is
    /// read is the kind's own: a PDF page follows `pdf_scale` and a page Office
    /// rendered follows `office_scale`, and neither moves for the picture scale.
    #[test]
    fn a_page_takes_the_room_the_display_has() {
        let pdf = PathBuf::from(r"C:\docs\report.pdf");
        let vector_scale = DEFAULT_VECTOR_SCALE;
        let office_scale = PreviewScale::Percent(75);

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
                        page: configured,
                        vector: vector_scale,
                        office: office_scale,
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
                    page: PreviewScale::Percent(50),
                    vector: vector_scale,
                    office: office_scale,
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
                    page: PreviewScale::Percent(25),
                    vector: vector_scale,
                    office: office_scale,
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
                    page: PreviewScale::FitToScreen,
                    vector: vector_scale,
                    office: office_scale,
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
        let page = PreviewScale::FitToScreen;

        assert_eq!(
            effective_preview_scale(
                &svg,
                HoverScales {
                    picture: configured,
                    vector: PreviewScale::FitToScreen,
                    page,
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
                        page,
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
                    page,
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
                        page,
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
        let page = PreviewScale::FitToScreen;
        let vector = PreviewScale::Percent(75);

        assert_eq!(
            effective_preview_scale(
                &font,
                HoverScales {
                    picture,
                    vector,
                    page,
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
                        page,
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
                        page,
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
                    page,
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
                    page: PreviewScale::Percent(25),
                    office: PreviewScale::Percent(75),
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
                    page: PreviewScale::Percent(25),
                    office: PreviewScale::Percent(75),
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
        if let Ok(mut cache) = VIDEO_GEOMETRY_CACHE.lock() {
            cache.insert(
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
        }

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

        assert!(webp_has_animation(&animated_path), "an `ANIM` chunk");
        assert!(!webp_has_animation(&still_path), "a plain picture");
        assert!(
            !webp_has_animation(&folder.join("missing.webp")),
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

        assert!(gif_is_animated(&gif), "two frame blocks in the file");
        assert!(!gif_is_animated(&still), "one frame block in the file");

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
                page: PreviewScale::Percent(25),
                office: PreviewScale::Percent(10),
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

    /// A load that may be about to finish is given the delay `spinner_delay_ms` names
    /// before the spinner goes up — the same one for every kind of wait — while a delay
    /// of nothing is a spinner that goes up with the load.
    #[test]
    fn puts_the_spinner_up_once_the_load_has_run_for_the_delay() {
        let load = |age: Duration, delay: Duration, upgrade: bool| PendingLoad {
            generation: 1,
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
        };
        let placement = HoverPlacement {
            orig_dims: (800, 600),
            avoid: None,
            follow_cursor: false,
            preview_scale: PreviewScale::FitToScreen,
            flush_at_cursor: false,
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
        // pointer's corner — the box a document waiting on a page is placed in — while
        // the preview keeps the place and the size its own layout gave it.
        let mut waiting = pending(Some(placement));
        let followed = waiting.follow_pointer(POINT { x: 300, y: 300 }, TEST_DPI);
        assert!(followed.spinner);
        assert_eq!(
            waiting.spinner_pos,
            (301, 301),
            "the spinner's own corner is at the cursor"
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
        // own box, still at the hand — and the preview's place with it.
        let followed = waiting.follow_pointer(POINT { x: 500, y: 300 }, TEST_DPI);
        assert!(followed.spinner);
        assert_eq!(waiting.spinner_pos, (501, 301));
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
            flush_at_cursor: false,
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
                TEST_DPI,
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
                                    // Only the waiting spinner is placed flush at the
                                    // pointer, and only it is ever that size.
                                    flush_at_cursor: (orig_width, orig_height) == (36, 36),
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
            crate::webview_preview::is_available()
        );
        println!("document: {}", crate::webview_preview::draws(&path));
        // Which engine would play a video here, which is the whole of the fallback's
        // routing: `ffplay` when it is installed, and the media engine Windows has when it
        // is not. Reported for every file rather than only for a video, because a probe is
        // run to find out what the machine is doing.
        println!(
            "video: played natively = {}",
            crate::codecs::plays_video_natively()
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
                            .chunks_exact(4)
                            .any(|pixel| pixel[..3] != [40, 40, 40]),
                    )
                })
            });

            println!(
                "{:>5} ms: engine_showing={} (w, h, frames, native video, framed)={:?}",
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
        if let Ok(mut cache) = VIDEO_GEOMETRY_CACHE.lock() {
            cache.insert(
                key.clone(),
                ProbedGeometry::Measured(VideoGeometry {
                    width: 640,
                    height: 360,
                    crop: None,
                }),
            );
        }
        assert_eq!(video_box(&path), Some((640, 360)));

        // A file the probe could not measure is not a file to probe again: it is the box
        // the player that would try the file anyway is given, and no preview at all where
        // the engine that plays it cannot open it either.
        if let Ok(mut cache) = VIDEO_GEOMETRY_CACHE.lock() {
            cache.insert(key, ProbedGeometry::Unmeasurable);
        }
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
                    flush_at_cursor: false,
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
            println!(
                "kinds: video = {}, pdf = {}, office = {}, libre = {}, design = {}, vector = {}, text = {}",
                is_video_file(&path),
                pdf_preview::is_pdf_file(&path),
                office_formats::is_office_file(&path),
                libre_formats::is_libre_file(&path),
                design_formats::is_design_file(&path),
                vector_formats::is_vector_file(&path),
                text_formats::is_text_file(&path),
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
                    flush_at_cursor: false,
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

use crate::config::TransparentBackground;
use crate::{CONFIG, RUNNING};
use gif::DecodeOptions;
use image::GenericImageView;
use once_cell::sync::Lazy;
use std::collections::{HashMap, VecDeque};
use std::env;
use std::fs::{File, OpenOptions};
use std::io::{BufReader, Write};
use std::os::windows::process::CommandExt;
use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::ptr;
use std::sync::atomic::{AtomicBool, AtomicIsize, AtomicU32, Ordering};
use std::sync::mpsc::{channel, Receiver, Sender};
use std::sync::{Arc, Condvar, Mutex};
use std::time::{Duration, Instant};
use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{COLORREF, HWND, LPARAM, LRESULT, POINT, RECT, SIZE, WPARAM};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, CreateCompatibleDC, CreateDIBSection, DeleteDC, DeleteObject, EndPaint,
    SelectObject, AC_SRC_ALPHA, AC_SRC_OVER, BITMAPINFO, BITMAPINFOHEADER, BI_RGB, BLENDFUNCTION,
    DIB_RGB_COLORS, PAINTSTRUCT,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DispatchMessageW, EnumWindows, GetSystemMetrics, GetWindow,
    GetWindowLongPtrW, GetWindowRect, GetWindowThreadProcessId, IsWindowVisible, LoadCursorW,
    MoveWindow, PeekMessageW, RegisterClassExW, SetWindowLongPtrW, SetWindowPos, ShowWindow,
    TranslateMessage, UpdateLayeredWindow, CS_HREDRAW, CS_VREDRAW, GWL_EXSTYLE, GW_OWNER,
    HWND_TOPMOST, IDC_ARROW, MSG, PM_REMOVE, SM_CXSCREEN, SM_CYSCREEN, SWP_NOACTIVATE, SWP_NOMOVE,
    SWP_NOSIZE, SWP_SHOWWINDOW, SW_HIDE, SW_SHOWNOACTIVATE, ULW_ALPHA, WNDCLASSEXW, WS_EX_LAYERED,
    WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_POPUP,
};

const PREVIEW_CLASS: PCWSTR = w!("RustHoverPreviewWindow");

// Video extensions for detection
const VIDEO_EXTENSIONS: &[&str] = &["mp4", "webm", "mkv", "avi", "mov", "wmv", "flv", "m4v"];
const MAX_STREAMED_ANIMATION_FRAMES: usize = 300;
const MAX_STREAMED_ANIMATION_BYTES: usize = 256 * 1024 * 1024;
const MIN_ANIMATION_FRAME_DELAY_MS: u32 = 33;
const ANIMATION_STARTUP_PREBUFFER_FRAMES: usize = 12;
const ANIMATION_STARTUP_PREBUFFER_MS: u32 = 500;
const CONFIG_RELOAD_INTERVAL_MS: u64 = 250;
const STREAMING_SPINNER_MAX_MS: u64 = 1500;

// Message passing for thread communication
pub static PREVIEW_SENDER: Lazy<Mutex<Option<Sender<PreviewMessage>>>> =
    Lazy::new(|| Mutex::new(None));

// Use AtomicIsize for the HWND pointer (thread-safe)
static PREVIEW_HWND: AtomicIsize = AtomicIsize::new(0);

// Track the ffplay video window HWND for cursor-over-preview detection
static VIDEO_HWND: AtomicIsize = AtomicIsize::new(0);
// Track the ffplay process ID to re-find the window if needed
static VIDEO_PID: AtomicU32 = AtomicU32::new(0);
// Guard to ensure we only run a single style-monitor thread.
static NOACTIVATE_MONITOR_STARTED: AtomicBool = AtomicBool::new(false);

static CURRENT_MEDIA: Lazy<Mutex<Option<MediaData>>> = Lazy::new(|| Mutex::new(None));
static VIDEO_GEOMETRY_CACHE: Lazy<Mutex<HashMap<PathBuf, VideoGeometry>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

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
    AnimatedWebP,
    Video,
    Loading,
}

/// A single frame of image data
struct ImageFrame {
    pixels: Vec<u8>,
    width: u32,
    height: u32,
    delay_ms: u32, // Delay before next frame (for animations)
}

/// Media data that can be either static or animated
struct MediaData {
    frames: Vec<ImageFrame>,
    /// Shared frame queue for streaming decode (animated formats append here)
    shared_frames: Option<Arc<Mutex<VecDeque<ImageFrame>>>>,
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

    /// Pull any newly decoded frames from the shared buffer
    fn sync_shared_frames(&mut self) {
        if let Some(ref shared) = self.shared_frames {
            if let Ok(mut shared_frames) = shared.lock() {
                if !shared_frames.is_empty() {
                    self.frames.extend(shared_frames.drain(..));
                }
            }
        }
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
            let delay = Duration::from_millis(self.frames[self.current_frame].delay_ms as u64);
            if self.last_frame_time.elapsed() >= delay {
                let next = self.current_frame + 1;
                if next < frame_count {
                    // More decoded frames ahead — advance normally
                    self.current_frame = next;
                    self.last_frame_time += delay;
                    advanced = true;
                } else if fully_loaded {
                    // All frames decoded — safe to loop back to start
                    self.current_frame = 0;
                    self.last_frame_time += delay;
                    advanced = true;
                } else {
                    // Still streaming — pause on this frame until the next
                    // one arrives. Keep the next streamed frame immediately
                    // eligible instead of adding another full-frame delay.
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
            MediaType::AnimatedGif | MediaType::AnimatedWebP
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

/// Check if cursor is currently over the IMAGE preview window only
pub fn is_cursor_over_image_preview() -> bool {
    unsafe {
        use windows::Win32::Foundation::POINT;
        use windows::Win32::UI::WindowsAndMessaging::{GetCursorPos, WindowFromPoint};

        let mut cursor_pos = POINT::default();
        if GetCursorPos(&mut cursor_pos).is_err() {
            return false;
        }

        let hwnd_under_cursor = WindowFromPoint(cursor_pos);
        if hwnd_under_cursor.is_invalid() {
            return false;
        }

        let hwnd_ptr = hwnd_under_cursor.0 as isize;

        let preview_hwnd = PREVIEW_HWND.load(Ordering::SeqCst);
        preview_hwnd != 0 && hwnd_ptr == preview_hwnd
    }
}

/// Check if cursor is currently over the VIDEO preview window (ffplay)
/// Also checks by process ID to handle the race condition where the ffplay
/// window exists but VIDEO_HWND hasn't been stored yet.
pub fn is_cursor_over_video_preview() -> bool {
    unsafe {
        use windows::Win32::Foundation::POINT;
        use windows::Win32::UI::WindowsAndMessaging::{
            GetCursorPos, GetWindowThreadProcessId, WindowFromPoint,
        };

        let mut cursor_pos = POINT::default();
        if GetCursorPos(&mut cursor_pos).is_err() {
            return false;
        }

        let hwnd_under_cursor = WindowFromPoint(cursor_pos);
        if hwnd_under_cursor.is_invalid() {
            return false;
        }

        let hwnd_ptr = hwnd_under_cursor.0 as isize;

        // Check by stored HWND
        let video_hwnd = VIDEO_HWND.load(Ordering::SeqCst);
        if video_hwnd != 0 && hwnd_ptr == video_hwnd {
            return true;
        }

        // Also check by process ID — covers the race window where ffplay's
        // window exists but VIDEO_HWND hasn't been discovered yet
        let video_pid = VIDEO_PID.load(Ordering::SeqCst);
        if video_pid != 0 {
            let mut window_pid: u32 = 0;
            GetWindowThreadProcessId(hwnd_under_cursor, Some(&mut window_pid));
            if window_pid == video_pid {
                return true;
            }
        }

        false
    }
}

fn is_video_file(path: &PathBuf) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| VIDEO_EXTENSIONS.contains(&ext.to_lowercase().as_str()))
        .unwrap_or(false)
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

fn checkerboard_color(x: u32, y: u32) -> (u8, u8, u8) {
    if ((x / 16) + (y / 16)) % 2 == 0 {
        (224, 224, 224)
    } else {
        (144, 144, 144)
    }
}

fn compose_preview_pixels(
    bgra: &[u8],
    width: u32,
    height: u32,
    background: TransparentBackground,
) -> Vec<u8> {
    let mut out = Vec::with_capacity(bgra.len());

    for (idx, px) in bgra.chunks(4).enumerate() {
        if px.len() != 4 {
            continue;
        }

        let b = px[0] as u32;
        let g = px[1] as u32;
        let r = px[2] as u32;
        let a = px[3] as u32;

        match background {
            TransparentBackground::Transparent => {
                out.push(((b * a + 127) / 255) as u8);
                out.push(((g * a + 127) / 255) as u8);
                out.push(((r * a + 127) / 255) as u8);
                out.push(a as u8);
            }
            TransparentBackground::Black
            | TransparentBackground::White
            | TransparentBackground::Checkerboard => {
                let x = (idx as u32) % width;
                let y = (idx as u32) / width;
                let (bg_b, bg_g, bg_r) = match background {
                    TransparentBackground::Black => (0, 0, 0),
                    TransparentBackground::White => (255, 255, 255),
                    TransparentBackground::Checkerboard => checkerboard_color(x, y),
                    TransparentBackground::Transparent => unreachable!(),
                };
                let inv_a = 255 - a;

                out.push(((b * a + (bg_b as u32) * inv_a + 127) / 255) as u8);
                out.push(((g * a + (bg_g as u32) * inv_a + 127) / 255) as u8);
                out.push(((r * a + (bg_r as u32) * inv_a + 127) / 255) as u8);
                out.push(255);
            }
        }
    }

    let expected = width as usize * height as usize * 4;
    if out.len() < expected {
        out.resize(expected, 0);
    }

    out
}

/// Scale image dimensions to fit within max bounds while maintaining aspect ratio
fn scale_dimensions(
    orig_width: u32,
    orig_height: u32,
    max_width: u32,
    max_height: u32,
) -> (u32, u32) {
    if orig_width <= max_width && orig_height <= max_height {
        return (orig_width, orig_height);
    }

    let scale_x = max_width as f32 / orig_width as f32;
    let scale_y = max_height as f32 / orig_height as f32;
    let scale = scale_x.min(scale_y);

    let new_width = (orig_width as f32 * scale).max(1.0) as u32;
    let new_height = (orig_height as f32 * scale).max(1.0) as u32;

    (new_width, new_height)
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
            image::imageops::FilterType::Nearest,
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

fn load_animated_gif(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
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
        scale_dimensions(gif_width, gif_height, max_width, max_height);

    let mut canvas = vec![0u8; (gif_width * gif_height * 4) as usize];
    let mut initial_frames = Vec::new();
    let mut initial_bytes: usize = 0;
    let mut buffered_ms: u32 = 0;
    let mut reached_end = false;

    while initial_frames.len() < MAX_STREAMED_ANIMATION_FRAMES
        && initial_frames.len() < ANIMATION_STARTUP_PREBUFFER_FRAMES
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
        if initial_bytes > MAX_STREAMED_ANIMATION_BYTES {
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
        });
    }

    let shared = Arc::new(Mutex::new(VecDeque::new()));
    let shared_clone = Arc::clone(&shared);
    let loaded_flag = Arc::new(AtomicBool::new(false));
    let loaded_flag_clone = Arc::clone(&loaded_flag);
    let skip_frames = initial_frames.len();

    let path_clone = path.clone();
    let cancel_clone = Arc::clone(&cancel);
    std::thread::spawn(move || {
        let file = match File::open(&path_clone) {
            Ok(f) => f,
            Err(_) => {
                loaded_flag_clone.store(true, Ordering::Release);
                return;
            }
        };
        let mut dec = DecodeOptions::new();
        dec.set_color_output(gif::ColorOutput::RGBA);
        let mut dec = match dec.read_info(BufReader::new(file)) {
            Ok(d) => d,
            Err(_) => {
                loaded_flag_clone.store(true, Ordering::Release);
                return;
            }
        };

        let mut canvas = vec![0u8; (gif_width * gif_height * 4) as usize];
        let mut frame_idx = 0usize;
        let mut streamed_count = skip_frames;
        let mut streamed_bytes = initial_bytes;

        while let Ok(Some(frame)) = dec.read_next_frame() {
            if cancel_clone.load(Ordering::Acquire)
                || streamed_count >= MAX_STREAMED_ANIMATION_FRAMES
            {
                break;
            }

            composite_gif_frame(&mut canvas, frame, gif_width, gif_height);
            if frame_idx < skip_frames {
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
                let frame_bytes = img.pixels.len();
                if streamed_bytes.saturating_add(frame_bytes) > MAX_STREAMED_ANIMATION_BYTES {
                    break;
                }
                if let Ok(mut frames) = shared_clone.lock() {
                    frames.push_back(img);
                }
                streamed_count += 1;
                streamed_bytes = streamed_bytes.saturating_add(frame_bytes);
            }
            frame_idx += 1;
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
    })
}

fn decode_webp_frame_to_image(
    buf: &[u8],
    has_alpha: bool,
    orig_width: u32,
    orig_height: u32,
    target_width: u32,
    target_height: u32,
    delay_ms: u32,
) -> Option<ImageFrame> {
    let expected_src = orig_width as usize * orig_height as usize * if has_alpha { 4 } else { 3 };
    if buf.len() != expected_src {
        return None;
    }

    if target_width == orig_width && target_height == orig_height {
        let mut bgra = Vec::with_capacity(orig_width as usize * orig_height as usize * 4);
        if has_alpha {
            for chunk in buf.chunks_exact(4) {
                bgra.push(chunk[2]);
                bgra.push(chunk[1]);
                bgra.push(chunk[0]);
                bgra.push(chunk[3]);
            }
        } else {
            for chunk in buf.chunks_exact(3) {
                bgra.push(chunk[2]);
                bgra.push(chunk[1]);
                bgra.push(chunk[0]);
                bgra.push(255);
            }
        }

        return Some(ImageFrame {
            pixels: bgra,
            width: target_width,
            height: target_height,
            delay_ms,
        });
    }

    let rgba = if has_alpha {
        buf.to_vec()
    } else {
        let mut rgba = Vec::with_capacity(orig_width as usize * orig_height as usize * 4);
        for chunk in buf.chunks_exact(3) {
            rgba.push(chunk[0]);
            rgba.push(chunk[1]);
            rgba.push(chunk[2]);
            rgba.push(255);
        }
        rgba
    };

    let img = image::RgbaImage::from_raw(orig_width, orig_height, rgba)?;
    let resized = image::imageops::resize(
        &img,
        target_width,
        target_height,
        image::imageops::FilterType::Nearest,
    );
    let bgra = rgba_to_bgra(&resized.into_raw());

    Some(ImageFrame {
        pixels: bgra,
        width: target_width,
        height: target_height,
        delay_ms,
    })
}

fn load_animated_webp(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    let file = File::open(path).ok()?;
    let reader = BufReader::new(file);
    let mut decoder = image_webp::WebPDecoder::new(reader).ok()?;

    if !decoder.is_animated() {
        return None;
    }

    let (orig_width, orig_height) = decoder.dimensions();
    if orig_width == 0 || orig_height == 0 || orig_width > 16384 || orig_height > 16384 {
        return None;
    }

    let (target_width, target_height) =
        scale_dimensions(orig_width, orig_height, max_width, max_height);
    if target_width == 0 || target_height == 0 {
        return None;
    }

    let has_alpha = decoder.has_alpha();
    let num_frames = decoder.num_frames();
    if !(2..=10000).contains(&num_frames) {
        return None;
    }

    let bytes_per_pixel: usize = if has_alpha { 4 } else { 3 };
    let buf_size = orig_width as usize * orig_height as usize * bytes_per_pixel;
    if buf_size > 100_000_000 {
        return None;
    }

    let mut buf = vec![0u8; buf_size];
    let mut initial_frames = Vec::new();
    let mut initial_bytes: usize = 0;
    let mut buffered_ms: u32 = 0;

    for _ in 0..num_frames {
        if cancel.load(Ordering::Acquire) {
            return None;
        }

        let delay_ms = match decoder.read_frame(&mut buf) {
            Ok(delay_ms) => delay_ms.max(MIN_ANIMATION_FRAME_DELAY_MS),
            Err(_) => break,
        };
        let img = decode_webp_frame_to_image(
            &buf,
            has_alpha,
            orig_width,
            orig_height,
            target_width,
            target_height,
            delay_ms,
        )?;
        initial_bytes = initial_bytes.saturating_add(img.pixels.len());
        if initial_bytes > MAX_STREAMED_ANIMATION_BYTES {
            return None;
        }
        buffered_ms = buffered_ms.saturating_add(delay_ms);
        initial_frames.push(img);

        if initial_frames.len() >= MAX_STREAMED_ANIMATION_FRAMES
            || (initial_frames.len() >= ANIMATION_STARTUP_PREBUFFER_FRAMES
                || (initial_frames.len() >= 2 && buffered_ms >= ANIMATION_STARTUP_PREBUFFER_MS))
        {
            break;
        }
    }

    if initial_frames.len() <= 1 {
        return None;
    }

    if initial_frames.len() >= num_frames as usize {
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
        });
    }

    let shared = Arc::new(Mutex::new(VecDeque::new()));
    let shared_clone = Arc::clone(&shared);
    let loaded_flag = Arc::new(AtomicBool::new(false));
    let loaded_flag_clone = Arc::clone(&loaded_flag);
    let skip_frames = initial_frames.len();

    let path_clone = path.clone();
    let cancel_clone = Arc::clone(&cancel);
    std::thread::spawn(move || {
        let file = match File::open(&path_clone) {
            Ok(f) => f,
            Err(_) => {
                loaded_flag_clone.store(true, Ordering::Release);
                return;
            }
        };
        let mut dec = match image_webp::WebPDecoder::new(BufReader::new(file)) {
            Ok(d) => d,
            Err(_) => {
                loaded_flag_clone.store(true, Ordering::Release);
                return;
            }
        };

        let bpp: usize = if dec.has_alpha() { 4 } else { 3 };
        let mut buf = vec![0u8; orig_width as usize * orig_height as usize * bpp];
        let has_alpha = dec.has_alpha();
        let total = dec.num_frames();
        let mut streamed_count = skip_frames;
        let mut streamed_bytes = initial_bytes;

        for i in 0..total {
            if cancel_clone.load(Ordering::Acquire)
                || streamed_count >= MAX_STREAMED_ANIMATION_FRAMES
            {
                break;
            }

            let delay_ms = match dec.read_frame(&mut buf) {
                Ok(delay_ms) => delay_ms.max(MIN_ANIMATION_FRAME_DELAY_MS),
                Err(_) => break,
            };
            if (i as usize) < skip_frames {
                continue;
            }

            if let Some(img) = decode_webp_frame_to_image(
                &buf,
                has_alpha,
                orig_width,
                orig_height,
                target_width,
                target_height,
                delay_ms,
            ) {
                let frame_bytes = img.pixels.len();
                if streamed_bytes.saturating_add(frame_bytes) > MAX_STREAMED_ANIMATION_BYTES {
                    break;
                }
                if let Ok(mut frames) = shared_clone.lock() {
                    frames.push_back(img);
                }
                streamed_count += 1;
                streamed_bytes = streamed_bytes.saturating_add(frame_bytes);
            }
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
    })
}

/// Load a static image (JPG, PNG, BMP, static WebP, etc.)
fn load_static_image(path: &PathBuf, max_width: u32, max_height: u32) -> Option<MediaData> {
    let img = if is_confirm_file_type_enabled() {
        decode_image_with_header_check(path)?
    } else {
        image::open(path).ok()?
    };
    let (orig_width, orig_height) = img.dimensions();
    let (target_width, target_height) =
        scale_dimensions(orig_width, orig_height, max_width, max_height);

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
    })
}

/// Extract video thumbnail using ffmpeg and create frames for preview
fn load_video_thumbnail(path: &PathBuf, max_width: u32, max_height: u32) -> Option<MediaData> {
    let geometry = get_video_geometry(path).unwrap_or(VideoGeometry {
        width: 1920,
        height: 1080,
        crop: None,
    });
    let (target_width, target_height) =
        scale_dimensions(geometry.width, geometry.height, max_width, max_height);

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

/// Apply WS_EX_NOACTIVATE style to a window
/// Returns true if the window was found and modified
unsafe fn try_apply_noactivate_style(pid: u32) -> bool {
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
        // Store the video window HWND for cursor-over-preview detection
        VIDEO_HWND.store(data.found_hwnd.0 as isize, Ordering::SeqCst);

        // Found the window, add WS_EX_NOACTIVATE and WS_EX_TOPMOST to its extended style
        let current_style = GetWindowLongPtrW(data.found_hwnd, GWL_EXSTYLE);
        let new_style = current_style
            | WS_EX_NOACTIVATE.0 as isize
            | WS_EX_TOOLWINDOW.0 as isize
            | WS_EX_TOPMOST.0 as isize;
        SetWindowLongPtrW(data.found_hwnd, GWL_EXSTYLE, new_style);

        // Force the video preview window to topmost so it doesn't hide behind Explorer
        let _ = SetWindowPos(
            data.found_hwnd,
            HWND_TOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW,
        );
        let _ = ShowWindow(data.found_hwnd, SW_SHOWNOACTIVATE);
        return true;
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
        let _ = process.kill();
        let _ = process.wait();
    }
    media.video_process = None;
    // Clear the video window HWND
    VIDEO_HWND.store(0, Ordering::SeqCst);
    VIDEO_PID.store(0, Ordering::SeqCst);
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
                        VIDEO_PID.store(0, Ordering::SeqCst);
                        return false;
                    }
                }
            }
        }
    }
    false
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

/// Load media (image, animated image, or video) with appropriate loader
fn load_media(
    path: &PathBuf,
    max_width: u32,
    max_height: u32,
    cancel: Arc<AtomicBool>,
) -> Option<MediaData> {
    if cancel.load(Ordering::Acquire) {
        return None;
    }

    if is_video_file(path) {
        return load_video_thumbnail(path, max_width, max_height);
    }

    let guessed_format = if is_confirm_file_type_enabled() {
        guessed_image_format(path)
    } else {
        None
    };

    if matches!(guessed_format, Some(image::ImageFormat::Gif)) || is_gif_file(path) {
        // Try animated GIF first
        if let Some(media) = load_animated_gif(path, max_width, max_height, Arc::clone(&cancel)) {
            return Some(media);
        }
        if cancel.load(Ordering::Acquire) {
            return None;
        }
        // Fall back to static for single-frame GIFs
        return load_static_image(path, max_width, max_height);
    }

    if matches!(guessed_format, Some(image::ImageFormat::WebP)) || is_webp_file(path) {
        // Try animated WebP first
        if let Some(media) = load_animated_webp(path, max_width, max_height, Arc::clone(&cancel)) {
            return Some(media);
        }
        if cancel.load(Ordering::Acquire) {
            return None;
        }
        // Fall back to static for non-animated WebP
        return load_static_image(path, max_width, max_height);
    }

    // Default to static image
    if cancel.load(Ordering::Acquire) {
        return None;
    }
    load_static_image(path, max_width, max_height)
}

/// Get original dimensions of media for positioning calculations
fn get_media_dimensions(path: &PathBuf) -> Option<(u32, u32)> {
    if is_video_file(path) {
        return get_video_geometry(path)
            .map(|g| (g.width, g.height))
            .or(Some((1920, 1080)));
    }

    if is_confirm_file_type_enabled() {
        image_dimensions_with_header_check(path)
    } else {
        image::image_dimensions(path).ok()
    }
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

unsafe fn render_layered_preview(hwnd: HWND) {
    let Some((width, height, pixels)) = (|| {
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
        let pixels = if media.should_draw_streaming_overlay() {
            let elapsed = media
                .loading_start
                .map(|s| s.elapsed().as_secs_f32())
                .unwrap_or(0.0);
            let angle = elapsed * 2.0 * std::f32::consts::PI * 1.2;
            let mut buf = media.current_pixels().to_vec();
            overlay_loading_spinner(&mut buf, width, height, angle);
            compose_preview_pixels(&buf, width, height, background)
        } else {
            compose_preview_pixels(media.current_pixels(), width, height, background)
        };

        Some((width, height, pixels))
    })() else {
        return;
    };

    let mut rect = RECT::default();
    if GetWindowRect(hwnd, &mut rect).is_err() {
        return;
    }

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

    let mem_dc = CreateCompatibleDC(None);
    if mem_dc.0.is_null() {
        return;
    }

    let mut bits: *mut core::ffi::c_void = ptr::null_mut();
    let Ok(bitmap) = CreateDIBSection(mem_dc, &bmi, DIB_RGB_COLORS, &mut bits, None, 0) else {
        let _ = DeleteDC(mem_dc);
        return;
    };

    if bits.is_null() {
        let _ = DeleteObject(bitmap);
        let _ = DeleteDC(mem_dc);
        return;
    }

    ptr::copy_nonoverlapping(pixels.as_ptr(), bits as *mut u8, pixels.len());

    let old_bitmap = SelectObject(mem_dc, bitmap);
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

    if !old_bitmap.0.is_null() {
        let _ = SelectObject(mem_dc, old_bitmap);
    }
    let _ = DeleteObject(bitmap);
    let _ = DeleteDC(mem_dc);
}

unsafe extern "system" fn window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
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

/// Compute preview layout for mouse hover (relative to cursor position)
fn compute_mouse_layout(
    cursor_x: i32,
    cursor_y: i32,
    orig_dims: (u32, u32),
    follow_cursor: bool,
    screen_width: i32,
    screen_height: i32,
) -> Option<PreviewLayout> {
    let offset = 20;
    let (orig_w, orig_h) = (orig_dims.0 as i32, orig_dims.1 as i32);

    if follow_cursor {
        let quadrants = [
            (
                screen_width - cursor_x - offset,
                screen_height - cursor_y - offset,
                cursor_x + offset,
                cursor_y + offset,
            ), // BR
            (
                cursor_x - offset,
                screen_height - cursor_y - offset,
                0,
                cursor_y + offset,
            ), // BL
            (
                screen_width - cursor_x - offset,
                cursor_y - offset,
                cursor_x + offset,
                0,
            ), // TR
            (cursor_x - offset, cursor_y - offset, 0, 0), // TL
        ];

        let mut best_quadrant = 0;
        let mut best_scale: f32 = 0.0;

        for (i, &(avail_w, avail_h, _, _)) in quadrants.iter().enumerate() {
            if avail_w <= 0 || avail_h <= 0 {
                continue;
            }
            let scale_x = avail_w as f32 / orig_w as f32;
            let scale_y = avail_h as f32 / orig_h as f32;
            let scale = scale_x.min(scale_y).min(1.0);
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

        let (preview_w, preview_h) =
            scale_dimensions(orig_dims.0, orig_dims.1, max_width, max_height);
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
        let left_width = cursor_x - offset;
        let right_width = screen_width - cursor_x - offset;
        let full_height = screen_height;

        let left_scale_x = left_width as f32 / orig_w as f32;
        let left_scale_y = full_height as f32 / orig_h as f32;
        let left_scale = left_scale_x.min(left_scale_y).min(1.0);

        let right_scale_x = right_width as f32 / orig_w as f32;
        let right_scale_y = full_height as f32 / orig_h as f32;
        let right_scale = right_scale_x.min(right_scale_y).min(1.0);

        let (use_left, max_width, max_height) = if left_scale > right_scale && left_width > 0 {
            (true, left_width.max(1) as u32, full_height as u32)
        } else if right_width > 0 {
            (false, right_width.max(1) as u32, full_height as u32)
        } else {
            return None;
        };

        let (preview_w, preview_h) =
            scale_dimensions(orig_dims.0, orig_dims.1, max_width, max_height);
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
        let pos_y = (screen_height - media_height) / 2;

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
    item_left: i32,
    item_top: i32,
    item_right: i32,
    item_bottom: i32,
    orig_dims: (u32, u32),
    follow_cursor: bool,
    screen_width: i32,
    screen_height: i32,
) -> Option<PreviewLayout> {
    let gap = 10;
    let (orig_w, orig_h) = (orig_dims.0 as i32, orig_dims.1 as i32);

    if follow_cursor {
        // Quadrant-based positioning relative to item rect edges
        let quadrants = [
            // Bottom-Right of item
            (
                screen_width - item_right - gap,
                screen_height - item_bottom - gap,
                item_right + gap,
                item_bottom + gap,
            ),
            // Bottom-Left of item
            (
                item_left - gap,
                screen_height - item_bottom - gap,
                0,
                item_bottom + gap,
            ),
            // Top-Right of item
            (
                screen_width - item_right - gap,
                item_top - gap,
                item_right + gap,
                0,
            ),
            // Top-Left of item
            (item_left - gap, item_top - gap, 0, 0),
        ];

        let mut best_quadrant = 0;
        let mut best_scale: f32 = 0.0;

        for (i, &(avail_w, avail_h, _, _)) in quadrants.iter().enumerate() {
            if avail_w <= 0 || avail_h <= 0 {
                continue;
            }
            let scale_x = avail_w as f32 / orig_w as f32;
            let scale_y = avail_h as f32 / orig_h as f32;
            let scale = scale_x.min(scale_y).min(1.0);
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

        let (preview_w, preview_h) =
            scale_dimensions(orig_dims.0, orig_dims.1, max_width, max_height);
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
        let left_width = item_left - gap;
        let right_width = screen_width - item_right - gap;
        let full_height = screen_height;

        let left_scale_x = left_width as f32 / orig_w as f32;
        let left_scale_y = full_height as f32 / orig_h as f32;
        let left_scale = left_scale_x.min(left_scale_y).min(1.0);

        let right_scale_x = right_width as f32 / orig_w as f32;
        let right_scale_y = full_height as f32 / orig_h as f32;
        let right_scale = right_scale_x.min(right_scale_y).min(1.0);

        let (use_left, max_width, max_height) = if left_scale > right_scale && left_width > 0 {
            (true, left_width.max(1) as u32, full_height as u32)
        } else if right_width > 0 {
            (false, right_width.max(1) as u32, full_height as u32)
        } else {
            return None;
        };

        let (preview_w, preview_h) =
            scale_dimensions(orig_dims.0, orig_dims.1, max_width, max_height);
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
        let pos_y = (screen_height - media_height) / 2;

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
        let mut last_config_reload = Instant::now();

        // Message loop
        let mut msg = MSG::default();
        while RUNNING.load(Ordering::SeqCst) {
            if last_config_reload.elapsed() >= Duration::from_millis(CONFIG_RELOAD_INTERVAL_MS) {
                last_config_reload = Instant::now();
                if let Ok(mut config) = CONFIG.lock() {
                    config.reload_from_disk();
                }
            }

            // Check for Windows messages
            while PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE).as_bool() {
                let _ = TranslateMessage(&msg);
                DispatchMessageW(&msg);
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

                            // If window wasn't shown yet (fast load), show it now
                            if let Some(ref pl) = pending_load {
                                if pl.generation == result.generation && !pl.spinner_shown {
                                    let _ = MoveWindow(hwnd, pl.pos_x, pl.pos_y, mw, mh, false);
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

                            if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                if let Some(ref mut existing) = *current {
                                    existing.cancel_background_work();
                                }
                                *current = Some(media_data);
                            }
                            pending_load = None;
                            pending_load_cancel = None;
                            render_layered_preview(hwnd);
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
                    if let Ok(mut current) = CURRENT_MEDIA.lock() {
                        *current = Some(loading);
                    }
                    let _ = MoveWindow(
                        hwnd,
                        pl.pos_x,
                        pl.pos_y,
                        pl.width as i32,
                        pl.height as i32,
                        false,
                    );
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
                    render_layered_preview(hwnd);
                }
            }

            // Check for our custom messages
            while let Ok(preview_msg) = rx.try_recv() {
                // Common variables for Show/ShowKeyboard - set in match, used after
                let mut show_path: Option<PathBuf> = None;
                let mut show_layout: Option<PreviewLayout> = None;
                let mut show_is_video: bool = false;

                match preview_msg {
                    PreviewMessage::Show(path, x, y) => {
                        let screen_width = GetSystemMetrics(SM_CXSCREEN);
                        let screen_height = GetSystemMetrics(SM_CYSCREEN);
                        let follow_cursor = CONFIG.lock().map(|c| c.follow_cursor).unwrap_or(true);

                        if let Some(orig_dims) = get_media_dimensions(&path) {
                            let is_video = is_video_file(&path);
                            if let Some(layout) = compute_mouse_layout(
                                x,
                                y,
                                orig_dims,
                                follow_cursor,
                                screen_width,
                                screen_height,
                            ) {
                                show_is_video = is_video;
                                show_layout = Some(layout);
                                show_path = Some(path);
                            }
                        }
                    }
                    PreviewMessage::ShowKeyboard(path, il, it, ir, ib) => {
                        let screen_width = GetSystemMetrics(SM_CXSCREEN);
                        let screen_height = GetSystemMetrics(SM_CYSCREEN);
                        let follow_cursor = CONFIG.lock().map(|c| c.follow_cursor).unwrap_or(true);

                        if let Some(orig_dims) = get_media_dimensions(&path) {
                            let is_video = is_video_file(&path);
                            if let Some(layout) = compute_keyboard_layout(
                                il,
                                it,
                                ir,
                                ib,
                                orig_dims,
                                follow_cursor,
                                screen_width,
                                screen_height,
                            ) {
                                show_is_video = is_video;
                                show_layout = Some(layout);
                                show_path = Some(path);
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

                    if show_is_video {
                        // Cancel any in-flight image load before switching to video.
                        current_generation += 1;
                        pending_load = None;
                        clear_load_request(&load_request_slot);
                        if let Some(cancel) = pending_load_cancel.take() {
                            cancel.store(true, Ordering::Release);
                        }

                        let no_cancel = Arc::new(AtomicBool::new(false));
                        if let Some(media_data) =
                            load_media(&path, max_width, max_height, no_cancel)
                        {
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
                                max_width,
                                max_height,
                                cancel: Arc::clone(&load_cancel),
                            },
                        );
                    }
                }
            }

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

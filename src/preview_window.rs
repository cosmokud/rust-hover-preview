use crate::{CONFIG, RUNNING};
use gif::DecodeOptions;
use image::GenericImageView;
use once_cell::sync::Lazy;
use std::fs::File;
use std::io::BufReader;
use std::os::windows::process::CommandExt;
use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::sync::atomic::{AtomicIsize, AtomicU32, Ordering};
use std::sync::mpsc::{channel, Receiver, Sender};
use std::sync::Mutex;
use std::time::{Duration, Instant};
use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{COLORREF, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, EndPaint, InvalidateRect, SetStretchBltMode, StretchDIBits, BITMAPINFO,
    BITMAPINFOHEADER, BI_RGB, DIB_RGB_COLORS, HALFTONE, PAINTSTRUCT, SRCCOPY,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DispatchMessageW, EnumWindows, GetSystemMetrics,
    GetWindowLongPtrW, GetWindowThreadProcessId, LoadCursorW, MoveWindow, PeekMessageW,
    RegisterClassExW, SetLayeredWindowAttributes, SetWindowLongPtrW, SetWindowPos, ShowWindow,
    TranslateMessage, CS_HREDRAW, CS_VREDRAW, GWL_EXSTYLE, HWND_TOPMOST, IDC_ARROW, LWA_ALPHA,
    MSG, PM_REMOVE, SM_CXSCREEN, SM_CYSCREEN, SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOSIZE,
    SWP_SHOWWINDOW, SW_HIDE, SW_SHOWNOACTIVATE, WNDCLASSEXW, WS_EX_LAYERED, WS_EX_NOACTIVATE,
    WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_POPUP,
};

const PREVIEW_CLASS: PCWSTR = w!("RustHoverPreviewWindow");

// Video extensions for detection
const VIDEO_EXTENSIONS: &[&str] = &["mp4", "webm", "mkv", "avi", "mov", "wmv", "flv", "m4v"];

// Message passing for thread communication
pub static PREVIEW_SENDER: Lazy<Mutex<Option<Sender<PreviewMessage>>>> =
    Lazy::new(|| Mutex::new(None));

// Use AtomicIsize for the HWND pointer (thread-safe)
static PREVIEW_HWND: AtomicIsize = AtomicIsize::new(0);

// Track the ffplay video window HWND for cursor-over-preview detection
static VIDEO_HWND: AtomicIsize = AtomicIsize::new(0);
// Track the ffplay process ID to re-find the window if needed
static VIDEO_PID: AtomicU32 = AtomicU32::new(0);

static CURRENT_MEDIA: Lazy<Mutex<Option<MediaData>>> = Lazy::new(|| Mutex::new(None));

pub enum PreviewMessage {
    Show(PathBuf, i32, i32),
    Hide,
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
    current_frame: usize,
    last_frame_time: Instant,
    media_type: MediaType,
    // For video playback using ffplay
    video_process: Option<Child>,
    loading_start: Option<Instant>,
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

    fn advance_frame(&mut self) -> bool {
        if self.frames.len() <= 1 {
            return false;
        }

        let delay = Duration::from_millis(self.frames[self.current_frame].delay_ms as u64);
        if self.last_frame_time.elapsed() >= delay {
            self.current_frame = (self.current_frame + 1) % self.frames.len();
            self.last_frame_time = Instant::now();
            return true;
        }
        false
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
}

pub fn show_preview(path: &PathBuf, x: i32, y: i32) {
    if let Ok(sender) = PREVIEW_SENDER.lock() {
        if let Some(ref tx) = *sender {
            let _ = tx.send(PreviewMessage::Show(path.clone(), x, y));
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

/// Check if cursor is currently over any preview window (image or video)
pub fn is_cursor_over_preview() -> bool {
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
        
        // Check image preview window
        let preview_hwnd = PREVIEW_HWND.load(Ordering::SeqCst);
        if preview_hwnd != 0 && hwnd_ptr == preview_hwnd {
            return true;
        }
        
        // Check video preview window (ffplay)
        let video_hwnd = VIDEO_HWND.load(Ordering::SeqCst);
        if video_hwnd != 0 && hwnd_ptr == video_hwnd {
            return true;
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

/// Load an animated GIF file
fn load_animated_gif(path: &PathBuf, max_width: u32, max_height: u32) -> Option<MediaData> {
    let file = File::open(path).ok()?;
    let mut decoder = DecodeOptions::new();
    decoder.set_color_output(gif::ColorOutput::RGBA);
    let mut decoder = decoder.read_info(BufReader::new(file)).ok()?;

    let (gif_width, gif_height) = (decoder.width() as u32, decoder.height() as u32);
    let (target_width, target_height) = scale_dimensions(gif_width, gif_height, max_width, max_height);

    let mut frames = Vec::new();
    let mut canvas = vec![0u8; (gif_width * gif_height * 4) as usize];

    while let Some(frame) = decoder.read_next_frame().ok()? {
        // Composite frame onto canvas
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

        // Scale canvas to target size
        let scaled = if target_width != gif_width || target_height != gif_height {
            let img =
                image::RgbaImage::from_raw(gif_width, gif_height, canvas.clone())?;
            let resized = image::imageops::resize(
                &img,
                target_width,
                target_height,
                image::imageops::FilterType::Triangle,
            );
            resized.into_raw()
        } else {
            canvas.clone()
        };

        let bgra = rgba_to_bgra(&scaled);

        // GIF delay is in centiseconds, convert to milliseconds (minimum 20ms for smooth playback)
        let delay_ms = (frame.delay as u32 * 10).max(20);

        frames.push(ImageFrame {
            pixels: bgra,
            width: target_width,
            height: target_height,
            delay_ms,
        });
    }

    if frames.is_empty() {
        return None;
    }

    Some(MediaData {
        frames,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::AnimatedGif,
        video_process: None,
        loading_start: None,
    })
}

/// Load an animated WebP file using image_webp directly for reliable frame decoding.
/// This bypasses the image crate's AnimationDecoder wrapper which has a frame iterator
/// state bug: if any frame errors, subsequent frames all retry the same broken position.
/// Using image_webp directly also avoids a latent RIFF padding bug in next_frame_start
/// calculation (unrounded anmf_size used instead of rounded).
fn load_animated_webp(path: &PathBuf, max_width: u32, max_height: u32) -> Option<MediaData> {
    let file = File::open(path).ok()?;
    let reader = BufReader::new(file);
    let mut decoder = image_webp::WebPDecoder::new(reader).ok()?;

    // Check if it's animated
    if !decoder.is_animated() {
        return None; // Not animated, use static loader
    }

    let (orig_width, orig_height) = decoder.dimensions();

    // Validate dimensions to prevent excessive allocation or crashes
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

    // Sanity check on frame count
    if num_frames == 0 || num_frames > 10000 {
        return None;
    }

    let bytes_per_pixel: usize = if has_alpha { 4 } else { 3 };
    let buf_size = orig_width as usize * orig_height as usize * bytes_per_pixel;

    // Prevent excessive memory allocation (100MB per frame buffer)
    if buf_size > 100_000_000 {
        return None;
    }

    let mut buf = vec![0u8; buf_size];

    let mut frames = Vec::new();

    for _ in 0..num_frames {
        match decoder.read_frame(&mut buf) {
            Ok(delay_ms) => {
                let delay_ms = delay_ms.max(20); // Minimum 20ms

                // Validate buffer was properly filled
                if buf.len() != buf_size {
                    break;
                }

                // Convert to RGBA if the image doesn't have alpha
                let rgba = if has_alpha {
                    buf.clone()
                } else {
                    let mut rgba =
                        Vec::with_capacity(orig_width as usize * orig_height as usize * 4);
                    for chunk in buf.chunks_exact(3) {
                        rgba.push(chunk[0]);
                        rgba.push(chunk[1]);
                        rgba.push(chunk[2]);
                        rgba.push(255);
                    }
                    rgba
                };

                // Validate RGBA buffer size
                let expected_rgba = orig_width as usize * orig_height as usize * 4;
                if rgba.len() != expected_rgba {
                    break;
                }

                let img = match image::RgbaImage::from_raw(orig_width, orig_height, rgba) {
                    Some(img) => img,
                    None => break,
                };

                let scaled = if target_width != orig_width || target_height != orig_height {
                    let resized = image::imageops::resize(
                        &img,
                        target_width,
                        target_height,
                        image::imageops::FilterType::Triangle,
                    );
                    resized.into_raw()
                } else {
                    img.into_raw()
                };

                let bgra = rgba_to_bgra(&scaled);

                // Final validation of output buffer
                let expected_bgra = target_width as usize * target_height as usize * 4;
                if bgra.len() != expected_bgra {
                    break;
                }

                frames.push(ImageFrame {
                    pixels: bgra,
                    width: target_width,
                    height: target_height,
                    delay_ms,
                });
            }
            Err(_) => break, // Stop on first error
        }
    }

    if frames.is_empty() {
        return None;
    }

    Some(MediaData {
        frames,
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::AnimatedWebP,
        video_process: None,
        loading_start: None,
    })
}

/// Load a static image (JPG, PNG, BMP, static WebP, etc.)
fn load_static_image(path: &PathBuf, max_width: u32, max_height: u32) -> Option<MediaData> {
    let img = image::open(path).ok()?;
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
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::StaticImage,
        video_process: None,
        loading_start: None,
    })
}

/// Extract video thumbnail using ffmpeg and create frames for preview
fn load_video_thumbnail(path: &PathBuf, max_width: u32, max_height: u32) -> Option<MediaData> {
    // Try to use ffprobe to get video dimensions
    let dimensions = get_video_dimensions(path).unwrap_or((1920, 1080));
    let (target_width, target_height) =
        scale_dimensions(dimensions.0, dimensions.1, max_width, max_height);

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
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::Video,
        video_process: None,
        loading_start: None,
    })
}

// Windows constant for hiding console window
const CREATE_NO_WINDOW: u32 = 0x08000000;

/// Get video dimensions using ffprobe
fn get_video_dimensions(path: &PathBuf) -> Option<(u32, u32)> {
    let output = Command::new("ffprobe")
        .args([
            "-v",
            "error",
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
        .creation_flags(CREATE_NO_WINDOW)  // Hide the console window
        .output()
        .ok()?;

    let output_str = String::from_utf8_lossy(&output.stdout);
    let parts: Vec<&str> = output_str.trim().split('x').collect();
    if parts.len() == 2 {
        let width = parts[0].parse().ok()?;
        let height = parts[1].parse().ok()?;
        return Some((width, height));
    }
    None
}

/// Data passed to the EnumWindows callback to find ffplay window
struct EnumWindowsData {
    target_pid: u32,
    found_hwnd: HWND,
}

/// Callback for EnumWindows to find a window belonging to a specific process
unsafe extern "system" fn enum_windows_callback(hwnd: HWND, lparam: LPARAM) -> windows::Win32::Foundation::BOOL {
    let data = &mut *(lparam.0 as *mut EnumWindowsData);
    let mut window_pid: u32 = 0;
    GetWindowThreadProcessId(hwnd, Some(&mut window_pid));
    
    if window_pid == data.target_pid {
        data.found_hwnd = hwnd;
        return windows::Win32::Foundation::BOOL(0); // Stop enumeration
    }
    windows::Win32::Foundation::BOOL(1) // Continue enumeration
}

/// Apply WS_EX_NOACTIVATE style to a window
/// Returns true if the window was found and modified
unsafe fn try_apply_noactivate_style(pid: u32) -> bool {
    let mut data = EnumWindowsData {
        target_pid: pid,
        found_hwnd: HWND::default(),
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
        return true;
    }
    false
}

/// Set WS_EX_NOACTIVATE on a window belonging to the given process
/// This prevents the window from stealing focus
/// Uses aggressive polling to minimize the race condition window
fn set_noactivate_for_process(pid: u32) {
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
    
    // Continue monitoring in background thread for longer period
    // The window might appear later, be recreated, or lose topmost
    std::thread::spawn(move || {
        unsafe {
            for i in 0..200 {
                let _ = try_apply_noactivate_style(pid);

                // Gradually increase delay as we wait longer
                let delay = if i < 20 { 1 } else if i < 60 { 5 } else { 25 };
                std::thread::sleep(Duration::from_millis(delay));
            }
        }
    });
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
    
    let child = cmd.args([
            "-loop",
            "0",                   // Loop forever
            "-noborder",           // No window border
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
        .creation_flags(CREATE_NO_WINDOW)  // Hide the console window
        .spawn()
        .ok();
    
    // After spawning, try to set WS_EX_NOACTIVATE on the ffplay window
    // to prevent it from stealing focus
    if let Some(ref child_process) = child {
        VIDEO_PID.store(child_process.id(), Ordering::SeqCst);
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
                        return false;
                    }
                    Ok(None) => return true,
                    Err(_) => {
                        media.video_process = None;
                        VIDEO_HWND.store(0, Ordering::SeqCst);
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
    let hwnd_val = VIDEO_HWND.load(Ordering::SeqCst);
    let mut hwnd_val = hwnd_val;
    if hwnd_val == 0 {
        let pid = VIDEO_PID.load(Ordering::SeqCst);
        if pid == 0 {
            return false;
        }

        unsafe {
            let _ = try_apply_noactivate_style(pid);
        }
        hwnd_val = VIDEO_HWND.load(Ordering::SeqCst);
        if hwnd_val == 0 {
            return false;
        }
    }

    unsafe {
        let hwnd = HWND(hwnd_val as *mut std::ffi::c_void);
        if hwnd.is_invalid() {
            return false;
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
    }

    true
}

/// Load media (image, animated image, or video) with appropriate loader
fn load_media(path: &PathBuf, max_width: u32, max_height: u32) -> Option<MediaData> {
    if is_video_file(path) {
        return load_video_thumbnail(path, max_width, max_height);
    }

    if is_gif_file(path) {
        // Try animated GIF first
        if let Some(media) = load_animated_gif(path, max_width, max_height) {
            if media.frames.len() > 1 {
                return Some(media);
            }
        }
        // Fall back to static for single-frame GIFs
        return load_static_image(path, max_width, max_height);
    }

    if is_webp_file(path) {
        // Try animated WebP first
        if let Some(media) = load_animated_webp(path, max_width, max_height) {
            return Some(media);
        }
        // Fall back to static for non-animated WebP
        return load_static_image(path, max_width, max_height);
    }

    // Default to static image
    load_static_image(path, max_width, max_height)
}

/// Get original dimensions of media for positioning calculations
fn get_media_dimensions(path: &PathBuf) -> Option<(u32, u32)> {
    if is_video_file(path) {
        return get_video_dimensions(path).or(Some((1920, 1080)));
    }

    image::image_dimensions(path).ok()
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
        pixel[3] = 255;   // A
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

                pixels[idx] = blend(bg[0], 255, alpha);     // B
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
        current_frame: 0,
        last_frame_time: Instant::now(),
        media_type: MediaType::Loading,
        video_process: None,
        loading_start: Some(Instant::now()),
    }
}

/// Result from background image loading thread
struct LoadResult {
    generation: u64,
    media: Option<MediaData>,
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
            let hdc = BeginPaint(hwnd, &mut ps);

            if let Ok(media_guard) = CURRENT_MEDIA.lock() {
                if let Some(ref media) = *media_guard {
                    // Don't paint for video - ffplay handles its own window
                    if !matches!(media.media_type, MediaType::Video) {
                        // Validate pixel buffer before painting
                        let expected_size = (media.current_width() as usize)
                            * (media.current_height() as usize)
                            * 4;
                        if media.current_width() == 0
                            || media.current_height() == 0
                            || media.current_pixels().len() < expected_size
                        {
                            let _ = EndPaint(hwnd, &ps);
                            return LRESULT(0);
                        }

                        // Create bitmap info
                        let bmi = BITMAPINFO {
                            bmiHeader: BITMAPINFOHEADER {
                                biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                                biWidth: media.current_width() as i32,
                                biHeight: -(media.current_height() as i32), // Negative for top-down
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

                        SetStretchBltMode(hdc, HALFTONE);

                        StretchDIBits(
                            hdc,
                            0,
                            0,
                            media.current_width() as i32,
                            media.current_height() as i32,
                            0,
                            0,
                            media.current_width() as i32,
                            media.current_height() as i32,
                            Some(media.current_pixels().as_ptr() as *const _),
                            &bmi,
                            DIB_RGB_COLORS,
                            SRCCOPY,
                        );
                    }
                }
            }

            let _ = EndPaint(hwnd, &ps);
            LRESULT(0)
        }
        windows::Win32::UI::WindowsAndMessaging::WM_DESTROY => LRESULT(0),
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
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

        // Set window fully opaque (255 = no transparency)
        SetLayeredWindowAttributes(hwnd, COLORREF(0), 255, LWA_ALPHA).ok();

        // Store HWND as isize
        PREVIEW_HWND.store(hwnd.0 as isize, Ordering::SeqCst);

        // Track current video path to avoid restarting
        let mut current_video_path: Option<PathBuf> = None;

        // Background loading support
        let (load_tx, load_rx): (Sender<LoadResult>, Receiver<LoadResult>) = channel();
        let mut current_generation: u64 = 0;

        // Message loop
        let mut msg = MSG::default();
        while RUNNING.load(Ordering::SeqCst) {
            // Check for Windows messages
            while PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE).as_bool() {
                let _ = TranslateMessage(&msg);
                DispatchMessageW(&msg);
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
                }
            }
            if needs_repaint {
                let _ = InvalidateRect(hwnd, None, false);
            }

            // Check for completed background loads
            while let Ok(result) = load_rx.try_recv() {
                if result.generation == current_generation {
                    match result.media {
                        Some(media_data) => {
                            if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                *current = Some(media_data);
                            }
                            let _ = InvalidateRect(hwnd, None, true);
                        }
                        None => {
                            // Loading failed, hide window
                            let _ = ShowWindow(hwnd, SW_HIDE);
                            if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                *current = None;
                            }
                        }
                    }
                }
            }

            // Check for our custom messages
            while let Ok(preview_msg) = rx.try_recv() {
                match preview_msg {
                    PreviewMessage::Show(path, x, y) => {
                        // Get screen dimensions
                        let screen_width = GetSystemMetrics(SM_CXSCREEN);
                        let screen_height = GetSystemMetrics(SM_CYSCREEN);
                        let offset = 20; // Gap between cursor and preview

                        // Get config for positioning mode
                        let follow_cursor =
                            CONFIG.lock().map(|c| c.follow_cursor).unwrap_or(true);

                        // Get original media dimensions first
                        let orig_dims = match get_media_dimensions(&path) {
                            Some(dims) => dims,
                            None => continue,
                        };
                        let (orig_w, orig_h) = (orig_dims.0 as i32, orig_dims.1 as i32);

                        let is_video = is_video_file(&path);

                        if follow_cursor {
                            // Follow cursor mode: use 4 quadrants around cursor
                            let quadrants = [
                                (
                                    screen_width - x - offset,
                                    screen_height - y - offset,
                                    x + offset,
                                    y + offset,
                                ), // BR
                                (x - offset, screen_height - y - offset, 0, y + offset), // BL
                                (screen_width - x - offset, y - offset, x + offset, 0),  // TR
                                (x - offset, y - offset, 0, 0),                          // TL
                            ];

                            // Find the best quadrant
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
                                continue;
                            }

                            let (avail_w, avail_h, _, _) = quadrants[best_quadrant];
                            let max_width = avail_w.max(1) as u32;
                            let max_height = avail_h.max(1) as u32;

                            // Pre-calculate preview dimensions for positioning
                            let (preview_w, preview_h) = scale_dimensions(
                                orig_dims.0, orig_dims.1, max_width, max_height,
                            );
                            let media_width = preview_w as i32;
                            let media_height = preview_h as i32;

                            if media_width <= 0 || media_height <= 0 {
                                continue;
                            }

                            let (pos_x, pos_y) = match best_quadrant {
                                0 => (x + offset, y + offset),
                                1 => (x - offset - media_width, y + offset),
                                2 => (x + offset, y - offset - media_height),
                                3 => (x - offset - media_width, y - offset - media_height),
                                _ => (x + offset, y + offset),
                            };

                            if is_video {
                                if let Some(media_data) = load_media(&path, max_width, max_height) {
                                    // For video, hide our window and use ffplay
                                    let _ = ShowWindow(hwnd, SW_HIDE);

                                    let process_running = is_video_process_running();
                                    let should_start = current_video_path.as_ref() != Some(&path)
                                        || !process_running;

                                    if should_start {
                                        if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                            if let Some(ref mut media) = *media_guard {
                                                stop_video_playback(media);
                                            }
                                        }

                                        let video_process = start_video_playback(
                                            &path, pos_x, pos_y, media_width, media_height,
                                        );

                                        if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                            let mut data = media_data;
                                            data.video_process = video_process;
                                            *current = Some(data);
                                        }

                                        current_video_path = Some(path.clone());
                                        let _ = ensure_video_window_topmost(
                                            pos_x, pos_y, media_width, media_height,
                                        );
                                    } else {
                                        let _ = ensure_video_window_topmost(
                                            pos_x, pos_y, media_width, media_height,
                                        );
                                    }
                                }
                            } else {
                                // For images/animations, show loading then load async
                                if current_video_path.is_some() {
                                    if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                        if let Some(ref mut media) = *media_guard {
                                            stop_video_playback(media);
                                        }
                                    }
                                    current_video_path = None;
                                }

                                // Show loading animation immediately
                                current_generation += 1;
                                let gen = current_generation;
                                let loading = create_loading_media(preview_w, preview_h);
                                if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                    *current = Some(loading);
                                }

                                let _ = MoveWindow(
                                    hwnd, pos_x, pos_y, media_width, media_height, false,
                                );
                                let _ = SetWindowPos(
                                    hwnd, HWND_TOPMOST, pos_x, pos_y,
                                    media_width, media_height,
                                    SWP_NOACTIVATE | SWP_SHOWWINDOW,
                                );
                                let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
                                let _ = InvalidateRect(hwnd, None, true);

                                // Load media in background thread
                                let tx = load_tx.clone();
                                let path_clone = path.clone();
                                std::thread::spawn(move || {
                                    let media = std::panic::catch_unwind(
                                        std::panic::AssertUnwindSafe(|| {
                                            load_media(&path_clone, max_width, max_height)
                                        }),
                                    )
                                    .unwrap_or(None);
                                    let _ = tx.send(LoadResult {
                                        generation: gen,
                                        media,
                                    });
                                });
                            }
                        } else {
                            // Best spot mode: choose left or right side of cursor for maximum size
                            let left_width = x - offset;
                            let right_width = screen_width - x - offset;
                            let full_height = screen_height;

                            // Calculate which side can show the media larger
                            let left_scale_x = left_width as f32 / orig_w as f32;
                            let left_scale_y = full_height as f32 / orig_h as f32;
                            let left_scale = left_scale_x.min(left_scale_y).min(1.0);

                            let right_scale_x = right_width as f32 / orig_w as f32;
                            let right_scale_y = full_height as f32 / orig_h as f32;
                            let right_scale = right_scale_x.min(right_scale_y).min(1.0);

                            let (use_left, max_width, max_height) =
                                if left_scale > right_scale && left_width > 0 {
                                    (true, left_width.max(1) as u32, full_height as u32)
                                } else if right_width > 0 {
                                    (false, right_width.max(1) as u32, full_height as u32)
                                } else {
                                    continue;
                                };

                            // Pre-calculate preview dimensions for positioning
                            let (preview_w, preview_h) = scale_dimensions(
                                orig_dims.0, orig_dims.1, max_width, max_height,
                            );
                            let media_width = preview_w as i32;
                            let media_height = preview_h as i32;

                            if media_width <= 0 || media_height <= 0 {
                                continue;
                            }

                            // Position: center vertically, left or right side
                            let pos_x = if use_left {
                                x - offset - media_width
                            } else {
                                x + offset
                            };
                            let pos_y = (screen_height - media_height) / 2;

                            if is_video {
                                if let Some(media_data) = load_media(&path, max_width, max_height) {
                                    // For video, hide our window and use ffplay
                                    let _ = ShowWindow(hwnd, SW_HIDE);

                                    let process_running = is_video_process_running();
                                    let should_start = current_video_path.as_ref() != Some(&path)
                                        || !process_running;

                                    if should_start {
                                        if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                            if let Some(ref mut media) = *media_guard {
                                                stop_video_playback(media);
                                            }
                                        }

                                        let video_process = start_video_playback(
                                            &path, pos_x, pos_y, media_width, media_height,
                                        );

                                        if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                            let mut data = media_data;
                                            data.video_process = video_process;
                                            *current = Some(data);
                                        }

                                        current_video_path = Some(path.clone());
                                        let _ = ensure_video_window_topmost(
                                            pos_x, pos_y, media_width, media_height,
                                        );
                                    } else {
                                        let _ = ensure_video_window_topmost(
                                            pos_x, pos_y, media_width, media_height,
                                        );
                                    }
                                }
                            } else {
                                // For images/animations, show loading then load async
                                if current_video_path.is_some() {
                                    if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                        if let Some(ref mut media) = *media_guard {
                                            stop_video_playback(media);
                                        }
                                    }
                                    current_video_path = None;
                                }

                                // Show loading animation immediately
                                current_generation += 1;
                                let gen = current_generation;
                                let loading = create_loading_media(preview_w, preview_h);
                                if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                    *current = Some(loading);
                                }

                                let _ = MoveWindow(
                                    hwnd, pos_x, pos_y, media_width, media_height, false,
                                );
                                let _ = SetWindowPos(
                                    hwnd, HWND_TOPMOST, pos_x, pos_y,
                                    media_width, media_height,
                                    SWP_NOACTIVATE | SWP_SHOWWINDOW,
                                );
                                let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
                                let _ = InvalidateRect(hwnd, None, true);

                                // Load media in background thread
                                let tx = load_tx.clone();
                                let path_clone = path.clone();
                                std::thread::spawn(move || {
                                    let media = std::panic::catch_unwind(
                                        std::panic::AssertUnwindSafe(|| {
                                            load_media(&path_clone, max_width, max_height)
                                        }),
                                    )
                                    .unwrap_or(None);
                                    let _ = tx.send(LoadResult {
                                        generation: gen,
                                        media,
                                    });
                                });
                            }
                        }
                    }
                    PreviewMessage::Hide => {
                        // Invalidate any pending background loads
                        current_generation += 1;

                        let _ = ShowWindow(hwnd, SW_HIDE);

                        // Stop video playback if any
                        if let Ok(mut current) = CURRENT_MEDIA.lock() {
                            if let Some(ref mut media) = *current {
                                stop_video_playback(media);
                            }
                            *current = None;
                        }
                        current_video_path = None;
                    }
                }
            }

            std::thread::sleep(std::time::Duration::from_millis(8)); // ~120fps for responsive preview
        }
    }
}

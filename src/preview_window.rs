use crate::{CONFIG, RUNNING};
use gif::DecodeOptions;
use image::codecs::webp::WebPDecoder;
use image::{AnimationDecoder, GenericImageView, ImageDecoder};
use once_cell::sync::Lazy;
use std::fs::File;
use std::io::BufReader;
use std::os::windows::process::CommandExt;
use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::sync::atomic::{AtomicIsize, Ordering};
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
    CreateWindowExW, DefWindowProcW, DispatchMessageW, GetSystemMetrics, LoadCursorW, MoveWindow,
    PeekMessageW, RegisterClassExW, SetLayeredWindowAttributes, SetWindowPos, ShowWindow,
    TranslateMessage, CS_HREDRAW, CS_VREDRAW, HWND_TOPMOST, IDC_ARROW, LWA_ALPHA, MSG, PM_REMOVE,
    SM_CXSCREEN, SM_CYSCREEN, SWP_NOACTIVATE, SWP_SHOWWINDOW, SW_HIDE, SW_SHOWNOACTIVATE,
    WNDCLASSEXW, WS_EX_LAYERED, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_POPUP,
};

const PREVIEW_CLASS: PCWSTR = w!("RustHoverPreviewWindow");

// Video extensions for detection
const VIDEO_EXTENSIONS: &[&str] = &["mp4", "webm", "mkv", "avi", "mov", "wmv", "flv", "m4v"];

// Message passing for thread communication
pub static PREVIEW_SENDER: Lazy<Mutex<Option<Sender<PreviewMessage>>>> =
    Lazy::new(|| Mutex::new(None));

// Use AtomicIsize for the HWND pointer (thread-safe)
static PREVIEW_HWND: AtomicIsize = AtomicIsize::new(0);

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
    })
}

/// Load an animated WebP file
fn load_animated_webp(path: &PathBuf, max_width: u32, max_height: u32) -> Option<MediaData> {
    let file = File::open(path).ok()?;
    let reader = BufReader::new(file);
    let decoder = WebPDecoder::new(reader).ok()?;

    // Check if it's animated
    if !decoder.has_animation() {
        return None; // Not animated, use static loader
    }

    let (orig_width, orig_height) = decoder.dimensions();
    let (target_width, target_height) =
        scale_dimensions(orig_width, orig_height, max_width, max_height);

    let mut frames = Vec::new();

    for frame_result in decoder.into_frames() {
        if let Ok(frame) = frame_result {
            let (numer, denom) = frame.delay().numer_denom_ms();
            let delay_ms = if denom > 0 { numer / denom } else { 100 };
            let delay_ms = delay_ms.max(20); // Minimum 20ms

            let img = frame.into_buffer();
            let (w, h) = (img.width(), img.height());

            let scaled = if target_width != w || target_height != h {
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

            frames.push(ImageFrame {
                pixels: bgra,
                width: target_width,
                height: target_height,
                delay_ms,
            });
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

/// Start ffplay for video preview (muted, looped)
fn start_video_playback(path: &PathBuf, x: i32, y: i32, width: i32, height: i32) -> Option<Child> {
    // Use ffplay for video playback - muted, borderless, positioned at preview location
    Command::new("ffplay")
        .args([
            "-an",                  // No audio (muted)
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
        .ok()
}

/// Stop video playback process
fn stop_video_playback(media: &mut MediaData) {
    if let Some(ref mut process) = media.video_process {
        let _ = process.kill();
        let _ = process.wait();
    }
    media.video_process = None;
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
                }
            }
            if needs_repaint {
                let _ = InvalidateRect(hwnd, None, false);
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

                            if let Some(media_data) = load_media(&path, max_width, max_height) {
                                let media_width = media_data.current_width() as i32;
                                let media_height = media_data.current_height() as i32;

                                let (pos_x, pos_y) = match best_quadrant {
                                    0 => (x + offset, y + offset),
                                    1 => (x - offset - media_width, y + offset),
                                    2 => (x + offset, y - offset - media_height),
                                    3 => (x - offset - media_width, y - offset - media_height),
                                    _ => (x + offset, y + offset),
                                };

                                if is_video {
                                    // For video, hide our window and use ffplay
                                    let _ = ShowWindow(hwnd, SW_HIDE);

                                    // Check if same video is already playing
                                    let should_start = current_video_path.as_ref() != Some(&path);

                                    if should_start {
                                        // Stop any existing video
                                        if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                            if let Some(ref mut media) = *media_guard {
                                                stop_video_playback(media);
                                            }
                                        }

                                        // Start new video playback
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
                                    }
                                } else {
                                    // For images/animations, use our preview window
                                    // Stop any video if switching from video
                                    if current_video_path.is_some() {
                                        if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                            if let Some(ref mut media) = *media_guard {
                                                stop_video_playback(media);
                                            }
                                        }
                                        current_video_path = None;
                                    }

                                    let _ = MoveWindow(
                                        hwnd,
                                        pos_x,
                                        pos_y,
                                        media_width,
                                        media_height,
                                        false,
                                    );
                                    let _ = SetWindowPos(
                                        hwnd,
                                        HWND_TOPMOST,
                                        pos_x,
                                        pos_y,
                                        media_width,
                                        media_height,
                                        SWP_NOACTIVATE | SWP_SHOWWINDOW,
                                    );

                                    if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                        *current = Some(media_data);
                                    }

                                    let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
                                    let _ = InvalidateRect(hwnd, None, true);
                                }
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

                            if let Some(media_data) = load_media(&path, max_width, max_height) {
                                let media_width = media_data.current_width() as i32;
                                let media_height = media_data.current_height() as i32;

                                // Position: center vertically, left or right side
                                let pos_x = if use_left {
                                    x - offset - media_width
                                } else {
                                    x + offset
                                };
                                let pos_y = (screen_height - media_height) / 2; // Center vertically

                                if is_video {
                                    // For video, hide our window and use ffplay
                                    let _ = ShowWindow(hwnd, SW_HIDE);

                                    let should_start = current_video_path.as_ref() != Some(&path);

                                    if should_start {
                                        if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                            if let Some(ref mut media) = *media_guard {
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
                                    }
                                } else {
                                    // For images/animations
                                    if current_video_path.is_some() {
                                        if let Ok(mut media_guard) = CURRENT_MEDIA.lock() {
                                            if let Some(ref mut media) = *media_guard {
                                                stop_video_playback(media);
                                            }
                                        }
                                        current_video_path = None;
                                    }

                                    let _ = MoveWindow(
                                        hwnd,
                                        pos_x,
                                        pos_y,
                                        media_width,
                                        media_height,
                                        false,
                                    );
                                    let _ = SetWindowPos(
                                        hwnd,
                                        HWND_TOPMOST,
                                        pos_x,
                                        pos_y,
                                        media_width,
                                        media_height,
                                        SWP_NOACTIVATE | SWP_SHOWWINDOW,
                                    );

                                    if let Ok(mut current) = CURRENT_MEDIA.lock() {
                                        *current = Some(media_data);
                                    }

                                    let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
                                    let _ = InvalidateRect(hwnd, None, true);
                                }
                            }
                        }
                    }
                    PreviewMessage::Hide => {
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

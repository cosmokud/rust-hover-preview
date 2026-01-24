use crate::{CONFIG, RUNNING};
use image::GenericImageView;
use once_cell::sync::Lazy;
use std::path::PathBuf;
use std::sync::atomic::{AtomicIsize, Ordering};
use std::sync::mpsc::{channel, Receiver, Sender};
use std::sync::Mutex;
use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{COLORREF, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, EndPaint, InvalidateRect, SetStretchBltMode, StretchDIBits, BITMAPINFO, BITMAPINFOHEADER, BI_RGB,
    DIB_RGB_COLORS, HALFTONE, PAINTSTRUCT, SRCCOPY,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DispatchMessageW, GetSystemMetrics,
    LoadCursorW, MoveWindow, PeekMessageW, RegisterClassExW, SetLayeredWindowAttributes,
    SetWindowPos, ShowWindow, TranslateMessage, CS_HREDRAW, CS_VREDRAW, HWND_TOPMOST, IDC_ARROW,
    LWA_ALPHA, MSG, PM_REMOVE, SM_CXSCREEN, SM_CYSCREEN, SWP_NOACTIVATE, SWP_SHOWWINDOW, SW_HIDE,
    SW_SHOWNOACTIVATE, WNDCLASSEXW, WS_EX_LAYERED, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW,
    WS_EX_TOPMOST, WS_POPUP,
};

const PREVIEW_CLASS: PCWSTR = w!("RustHoverPreviewWindow");

// Message passing for thread communication
pub static PREVIEW_SENDER: Lazy<Mutex<Option<Sender<PreviewMessage>>>> =
    Lazy::new(|| Mutex::new(None));

// Use AtomicIsize for the HWND pointer (thread-safe)
static PREVIEW_HWND: AtomicIsize = AtomicIsize::new(0);

static CURRENT_IMAGE: Lazy<Mutex<Option<ImageData>>> = Lazy::new(|| Mutex::new(None));

pub enum PreviewMessage {
    Show(PathBuf, i32, i32),
    Hide,
}

struct ImageData {
    pixels: Vec<u8>,
    width: u32,
    height: u32,
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

fn load_image(path: &PathBuf, max_width: u32, max_height: u32) -> Option<ImageData> {
    let img = image::open(path).ok()?;

    // Calculate scaled dimensions to fit within screen bounds while maintaining aspect ratio
    let (orig_width, orig_height) = img.dimensions();
    
    // If image fits, show at 100%
    if orig_width <= max_width && orig_height <= max_height {
        let rgba = img.to_rgba8();
        let mut pixels: Vec<u8> = Vec::with_capacity((orig_width * orig_height * 4) as usize);
        for pixel in rgba.pixels() {
            pixels.push(pixel[2]); // B
            pixels.push(pixel[1]); // G
            pixels.push(pixel[0]); // R
            pixels.push(pixel[3]); // A
        }
        return Some(ImageData {
            pixels,
            width: orig_width,
            height: orig_height,
        });
    }
    
    // Scale down to fit while maintaining aspect ratio
    let scale_x = max_width as f32 / orig_width as f32;
    let scale_y = max_height as f32 / orig_height as f32;
    let scale = scale_x.min(scale_y);
    let new_width = (orig_width as f32 * scale) as u32;
    let new_height = (orig_height as f32 * scale) as u32;

    // Resize image - use resize_exact to get exact dimensions we calculated
    // (img.resize preserves aspect ratio and may return different dimensions)
    let resized = img.resize_exact(new_width, new_height, image::imageops::FilterType::Triangle);
    let rgba = resized.to_rgba8();

    // Convert to BGRA for Windows
    let mut pixels: Vec<u8> = Vec::with_capacity((new_width * new_height * 4) as usize);
    for pixel in rgba.pixels() {
        pixels.push(pixel[2]); // B
        pixels.push(pixel[1]); // G
        pixels.push(pixel[0]); // R
        pixels.push(pixel[3]); // A
    }

    Some(ImageData {
        pixels,
        width: new_width,
        height: new_height,
    })
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

            if let Ok(image_guard) = CURRENT_IMAGE.lock() {
                if let Some(ref img) = *image_guard {
                    // Create bitmap info
                    let bmi = BITMAPINFO {
                        bmiHeader: BITMAPINFOHEADER {
                            biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                            biWidth: img.width as i32,
                            biHeight: -(img.height as i32), // Negative for top-down
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
                        img.width as i32,
                        img.height as i32,
                        0,
                        0,
                        img.width as i32,
                        img.height as i32,
                        Some(img.pixels.as_ptr() as *const _),
                        &bmi,
                        DIB_RGB_COLORS,
                        SRCCOPY,
                    );
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

        // Message loop
        let mut msg = MSG::default();
        while RUNNING.load(Ordering::SeqCst) {
            // Check for Windows messages
            while PeekMessageW(&mut msg, None, 0, 0, PM_REMOVE).as_bool() {
                let _ = TranslateMessage(&msg);
                DispatchMessageW(&msg);
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
                        let follow_cursor = CONFIG.lock().map(|c| c.follow_cursor).unwrap_or(true);
                        
                        // Get original image dimensions first
                        let orig_dims = match image::image_dimensions(&path) {
                            Ok(dims) => dims,
                            Err(_) => continue,
                        };
                        let (orig_w, orig_h) = (orig_dims.0 as i32, orig_dims.1 as i32);
                        
                        if follow_cursor {
                            // Follow cursor mode: use 4 quadrants around cursor
                            // Quadrant 0: Bottom-Right of cursor
                            // Quadrant 1: Bottom-Left of cursor  
                            // Quadrant 2: Top-Right of cursor
                            // Quadrant 3: Top-Left of cursor
                            let quadrants = [
                                (screen_width - x - offset, screen_height - y - offset, x + offset, y + offset),      // BR
                                (x - offset, screen_height - y - offset, 0, y + offset),                              // BL
                                (screen_width - x - offset, y - offset, x + offset, 0),                               // TR
                                (x - offset, y - offset, 0, 0),                                                        // TL
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
                            
                            if let Some(img_data) = load_image(&path, max_width, max_height) {
                                let img_width = img_data.width as i32;
                                let img_height = img_data.height as i32;
                                
                                let (pos_x, pos_y) = match best_quadrant {
                                    0 => (x + offset, y + offset),
                                    1 => (x - offset - img_width, y + offset),
                                    2 => (x + offset, y - offset - img_height),
                                    3 => (x - offset - img_width, y - offset - img_height),
                                    _ => (x + offset, y + offset),
                                };
                                
                                // IMPORTANT: Resize window BEFORE setting image data to prevent
                                // race condition where WM_PAINT fires with old window size but new image
                                let _ = MoveWindow(hwnd, pos_x, pos_y, img_width, img_height, false);
                                let _ = SetWindowPos(hwnd, HWND_TOPMOST, pos_x, pos_y, img_width, img_height, SWP_NOACTIVATE | SWP_SHOWWINDOW);
                                
                                // Now set the image data after window is properly sized
                                if let Ok(mut current) = CURRENT_IMAGE.lock() {
                                    *current = Some(img_data);
                                }
                                
                                let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
                                let _ = InvalidateRect(hwnd, None, true);
                            }
                        } else {
                            // Best spot mode: choose left or right side of cursor for maximum size
                            let left_width = x - offset;
                            let right_width = screen_width - x - offset;
                            let full_height = screen_height;
                            
                            // Calculate which side can show the image larger
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
                                continue;
                            };
                            
                            if let Some(img_data) = load_image(&path, max_width, max_height) {
                                let img_width = img_data.width as i32;
                                let img_height = img_data.height as i32;
                                
                                // Position: center vertically, left or right side
                                let pos_x = if use_left {
                                    x - offset - img_width
                                } else {
                                    x + offset
                                };
                                let pos_y = (screen_height - img_height) / 2; // Center vertically
                                
                                // IMPORTANT: Resize window BEFORE setting image data to prevent
                                // race condition where WM_PAINT fires with old window size but new image
                                let _ = MoveWindow(hwnd, pos_x, pos_y, img_width, img_height, false);
                                let _ = SetWindowPos(hwnd, HWND_TOPMOST, pos_x, pos_y, img_width, img_height, SWP_NOACTIVATE | SWP_SHOWWINDOW);
                                
                                // Now set the image data after window is properly sized
                                if let Ok(mut current) = CURRENT_IMAGE.lock() {
                                    *current = Some(img_data);
                                }
                                
                                let _ = ShowWindow(hwnd, SW_SHOWNOACTIVATE);
                                let _ = InvalidateRect(hwnd, None, true);
                            }
                        }
                    }
                    PreviewMessage::Hide => {
                        let _ = ShowWindow(hwnd, SW_HIDE);
                        if let Ok(mut current) = CURRENT_IMAGE.lock() {
                            *current = None;
                        }
                    }
                }
            }

            std::thread::sleep(std::time::Duration::from_millis(8)); // ~120fps for responsive preview
        }
    }
}

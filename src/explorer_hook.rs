use crate::preview_window::{hide_preview, show_preview};
use crate::{CONFIG, RUNNING};
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use std::path::PathBuf;
use std::sync::atomic::Ordering;
use std::time::{Duration, Instant};
use windows::core::{Interface, VARIANT};
use windows::Win32::Foundation::{HWND, POINT, RECT, SHANDLE_PTR};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CLSCTX_ALL, COINIT_APARTMENTTHREADED,
};
use windows::Win32::UI::Shell::{IShellWindows, ShellWindows};
use windows::Win32::UI::WindowsAndMessaging::{
    GetClassNameW, GetCursorPos, GetForegroundWindow, GetWindowPlacement, GetWindowRect,
    GetWindowThreadProcessId, IsIconic, IsWindowVisible, WindowFromPoint, WINDOWPLACEMENT,
    SW_SHOWMAXIMIZED,
};

// Supported image extensions
const IMAGE_EXTENSIONS: &[&str] = &[
    "jpg", "jpeg", "png", "gif", "bmp", "ico", "tiff", "tif", "webp",
];

// Supported video extensions
const VIDEO_EXTENSIONS: &[&str] = &[
    "mp4", "webm", "mkv", "avi", "mov", "wmv", "flv", "m4v",
];

fn is_image_file(path: &PathBuf) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| IMAGE_EXTENSIONS.contains(&ext.to_lowercase().as_str()))
        .unwrap_or(false)
}

fn is_video_file(path: &PathBuf) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| VIDEO_EXTENSIONS.contains(&ext.to_lowercase().as_str()))
        .unwrap_or(false)
}

fn is_media_file(path: &PathBuf) -> bool {
    is_image_file(path) || is_video_file(path)
}

/// Get current folder path from an Explorer window
fn get_explorer_folder_path(hwnd: HWND) -> Option<String> {
    unsafe {
        let shell_windows: IShellWindows =
            CoCreateInstance(&ShellWindows, None, CLSCTX_ALL).ok()?;

        let count = shell_windows.Count().ok()?;

        for i in 0..count {
            let variant = VARIANT::from(i);
            if let Ok(disp) = shell_windows.Item(&variant) {
                // Get the IWebBrowser2 interface
                if let Ok(browser) =
                    disp.cast::<windows::Win32::UI::Shell::IWebBrowser2>()
                {
                    // Check if this is the window we're looking for
                    if let Ok(browser_hwnd) = browser.HWND() {
                        if browser_hwnd == SHANDLE_PTR(hwnd.0 as isize) {
                            // Get the location URL
                            if let Ok(url) = browser.LocationURL() {
                                let url_str = url.to_string();
                                if url_str.starts_with("file:///") {
                                    let path = url_str
                                        .strip_prefix("file:///")
                                        .unwrap_or(&url_str)
                                        .replace('/', "\\");
                                    // URL decode
                                    let path = urlencoding_decode(&path);
                                    return Some(path);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    None
}

/// Simple URL decoding for file paths
fn urlencoding_decode(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    let mut chars = s.chars().peekable();

    while let Some(c) = chars.next() {
        if c == '%' {
            let hex: String = chars.by_ref().take(2).collect();
            if let Ok(byte) = u8::from_str_radix(&hex, 16) {
                result.push(byte as char);
            } else {
                result.push('%');
                result.push_str(&hex);
            }
        } else {
            result.push(c);
        }
    }

    result
}

/// Get all Explorer windows and their current folder paths
fn get_all_explorer_folders() -> Vec<(HWND, String)> {
    let mut result = Vec::new();

    unsafe {
        if let Ok(shell_windows) =
            CoCreateInstance::<_, IShellWindows>(&ShellWindows, None, CLSCTX_ALL)
        {
            if let Ok(count) = shell_windows.Count() {
                for i in 0..count {
                    let variant = VARIANT::from(i);
                    if let Ok(disp) = shell_windows.Item(&variant) {
                        if let Ok(browser) =
                            disp.cast::<windows::Win32::UI::Shell::IWebBrowser2>()
                        {
                            if let Ok(browser_hwnd) = browser.HWND() {
                                let hwnd = HWND(browser_hwnd.0 as *mut _);
                                if let Ok(url) = browser.LocationURL() {
                                    let url_str = url.to_string();
                                    if url_str.starts_with("file:///") {
                                        let path = url_str
                                            .strip_prefix("file:///")
                                            .unwrap_or(&url_str)
                                            .replace('/', "\\");
                                        let path = urlencoding_decode(&path);
                                        result.push((hwnd, path));
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    result
}

/// Find which Explorer folder the cursor is currently over
fn get_current_explorer_folder() -> Option<String> {
    unsafe {
        let mut cursor_pos = POINT::default();
        if GetCursorPos(&mut cursor_pos).is_err() {
            return None;
        }

        // Get window under cursor
        let hwnd = WindowFromPoint(cursor_pos);
        if hwnd.0.is_null() {
            return None;
        }

        // Walk up parent windows to find Explorer window
        let mut current_hwnd = hwnd;
        let folders = get_all_explorer_folders();

        // Check if any parent window is an Explorer window
        for _ in 0..20 {
            // Limit iterations
            for (explorer_hwnd, folder) in &folders {
                if current_hwnd == *explorer_hwnd {
                    return Some(folder.clone());
                }
            }

            // Get parent
            if let Ok(parent) = windows::Win32::UI::WindowsAndMessaging::GetParent(current_hwnd) {
                if parent.0.is_null() || parent == current_hwnd {
                    break;
                }
                current_hwnd = parent;
            } else {
                break;
            }
        }

        // Also check if the foreground window is an Explorer
        let foreground = GetForegroundWindow();
        for (explorer_hwnd, folder) in &folders {
            if foreground == *explorer_hwnd {
                return Some(folder.clone());
            }
        }
    }

    None
}

/// Get the filename under cursor using accessibility - try multiple approaches
fn get_item_name_under_cursor() -> Option<String> {
    unsafe {
        let mut cursor_pos = POINT::default();
        if GetCursorPos(&mut cursor_pos).is_err() {
            return None;
        }

        // Use accessibility to get the item name
        let mut accessible: Option<windows::Win32::UI::Accessibility::IAccessible> = None;
        let mut child_variant = VARIANT::default();

        if windows::Win32::UI::Accessibility::AccessibleObjectFromPoint(
            cursor_pos,
            &mut accessible,
            &mut child_variant,
        )
        .is_ok()
        {
            if let Some(ref acc) = accessible {
                // Try with the child variant first
                if let Ok(name) = acc.get_accName(&child_variant) {
                    let name_str = name.to_string();
                    if !name_str.is_empty() && name_str != "Items View" && name_str != "Folder View" {
                        return Some(name_str);
                    }
                }
                
                // Try with default variant
                if let Ok(name) = acc.get_accName(&VARIANT::default()) {
                    let name_str = name.to_string();
                    if !name_str.is_empty() && name_str != "Items View" && name_str != "Folder View" {
                        return Some(name_str);
                    }
                }

                // Try get_accValue
                if let Ok(value) = acc.get_accValue(&child_variant) {
                    let value_str = value.to_string();
                    if !value_str.is_empty() {
                        return Some(value_str);
                    }
                }

                // Try getting help text which sometimes has the filename
                if let Ok(help) = acc.get_accHelp(&child_variant) {
                    let help_str = help.to_string();
                    if !help_str.is_empty() {
                        return Some(help_str);
                    }
                }
            }
        }
    }

    None
}

/// Try to find an image or video file in a specific folder by item name
fn find_media_in_folder(folder: &str, item_name: &str) -> Option<PathBuf> {
    let folder_path = PathBuf::from(folder);

    // First try: item_name as-is
    let full_path = folder_path.join(item_name);
    if full_path.exists() && is_media_file(&full_path) {
        return Some(full_path);
    }

    // Second try: search for files that match this name (with or without extension)
    if let Ok(entries) = std::fs::read_dir(&folder_path) {
        for entry in entries.flatten() {
            let path = entry.path();
            
            // Check full filename match (e.g., "image.jpg" or "video.mp4")
            if let Some(file_name) = path.file_name().and_then(|s| s.to_str()) {
                if file_name == item_name && is_media_file(&path) {
                    return Some(path);
                }
            }
            
            // Check file stem match (e.g., "image" matches "image.jpg")
            if let Some(file_stem) = path.file_stem().and_then(|s| s.to_str()) {
                if file_stem == item_name && is_media_file(&path) {
                    return Some(path);
                }
            }
            
            // Check case-insensitive match
            if let Some(file_name) = path.file_name().and_then(|s| s.to_str()) {
                if file_name.to_lowercase() == item_name.to_lowercase() && is_media_file(&path) {
                    return Some(path);
                }
            }
        }
    }

    None
}

/// Try to find an image or video file under the cursor
fn get_file_under_cursor() -> Option<PathBuf> {
    // Get the item name under cursor
    let item_name = get_item_name_under_cursor()?;

    // Get ALL Explorer folders (all windows and tabs)
    let all_folders = get_all_explorer_folders();

    // Try to find the file in ANY of the open Explorer folders
    for (_, folder) in &all_folders {
        if let Some(path) = find_media_in_folder(folder, &item_name) {
            return Some(path);
        }
    }

    None
}

/// Quick check if foreground window is Explorer (cheap, no COM)
fn is_foreground_explorer() -> bool {
    unsafe {
        let foreground = GetForegroundWindow();
        if foreground.is_invalid() {
            return false;
        }
        is_explorer_window(foreground)
    }
}

/// Check if a window is maximized
fn is_window_maximized(hwnd: HWND) -> bool {
    unsafe {
        let mut placement = WINDOWPLACEMENT::default();
        placement.length = std::mem::size_of::<WINDOWPLACEMENT>() as u32;
        if GetWindowPlacement(hwnd, &mut placement).is_ok() {
            return placement.showCmd == SW_SHOWMAXIMIZED.0 as u32;
        }
    }
    false
}

/// Check if a window is fullscreen (covers entire screen)
fn is_window_fullscreen(hwnd: HWND) -> bool {
    unsafe {
        let mut window_rect = RECT::default();
        if GetWindowRect(hwnd, &mut window_rect).is_err() {
            return false;
        }
        
        // Get screen dimensions
        let screen_width = windows::Win32::UI::WindowsAndMessaging::GetSystemMetrics(
            windows::Win32::UI::WindowsAndMessaging::SM_CXSCREEN,
        );
        let screen_height = windows::Win32::UI::WindowsAndMessaging::GetSystemMetrics(
            windows::Win32::UI::WindowsAndMessaging::SM_CYSCREEN,
        );
        
        // Check if window covers entire screen (with small tolerance for borders)
        let width = window_rect.right - window_rect.left;
        let height = window_rect.bottom - window_rect.top;
        
        width >= screen_width && height >= screen_height
    }
}

/// Check if foreground window is maximized or fullscreen AND is not Explorer
/// Returns true if we should sleep (Explorer is hidden behind a maximized/fullscreen window)
fn is_explorer_hidden_by_foreground() -> bool {
    unsafe {
        let foreground = GetForegroundWindow();
        if foreground.is_invalid() {
            return false;
        }
        
        // If foreground IS Explorer, it's not hidden
        if is_explorer_window(foreground) {
            return false;
        }
        
        // Check if foreground is maximized or fullscreen
        is_window_maximized(foreground) || is_window_fullscreen(foreground)
    }
}

/// Check if a window is minimized
fn is_window_minimized(hwnd: HWND) -> bool {
    unsafe { IsIconic(hwnd).as_bool() }
}

/// Get count of Explorer windows and count of visible (not minimized) ones
/// Returns (total_count, visible_count)
fn get_explorer_window_counts() -> (usize, usize) {
    unsafe {
        let mut total = 0;
        let mut visible = 0;
        
        if let Ok(shell_windows) =
            CoCreateInstance::<_, IShellWindows>(&ShellWindows, None, CLSCTX_ALL)
        {
            if let Ok(count) = shell_windows.Count() {
                for i in 0..count {
                    let variant = VARIANT::from(i);
                    if let Ok(disp) = shell_windows.Item(&variant) {
                        if let Ok(browser) =
                            disp.cast::<windows::Win32::UI::Shell::IWebBrowser2>()
                        {
                            if let Ok(browser_hwnd) = browser.HWND() {
                                let hwnd = HWND(browser_hwnd.0 as *mut _);
                                total += 1;
                                
                                // Check if window is visible and not minimized
                                if IsWindowVisible(hwnd).as_bool() && !is_window_minimized(hwnd) {
                                    visible += 1;
                                }
                            }
                        }
                    }
                }
            }
        }
        
        (total, visible)
    }
}

/// Enum representing the state of Explorer windows for CPU optimization
#[derive(Debug, Clone, Copy, PartialEq)]
enum ExplorerState {
    /// No Explorer windows open at all - longest sleep
    NoExplorerWindows,
    /// All Explorer windows are minimized - long sleep
    AllMinimized,
    /// A non-Explorer window is maximized/fullscreen, hiding Explorer - long sleep
    HiddenByForeground,
    /// Explorer is visible but not in focus - medium sleep
    VisibleNotFocused,
    /// Explorer is in focus and cursor might be over it - active polling
    ActiveFocus,
}

/// Determine the current state of Explorer for CPU optimization
fn get_explorer_state() -> ExplorerState {
    // Quick check: is foreground Explorer? (cheapest check)
    if is_foreground_explorer() {
        return ExplorerState::ActiveFocus;
    }
    
    // Check if foreground is maximized/fullscreen (cheap check)
    if is_explorer_hidden_by_foreground() {
        return ExplorerState::HiddenByForeground;
    }
    
    // Need to check Explorer window states (more expensive, uses COM)
    let (total, visible) = get_explorer_window_counts();
    
    if total == 0 {
        return ExplorerState::NoExplorerWindows;
    }
    
    if visible == 0 {
        return ExplorerState::AllMinimized;
    }
    
    // Explorer windows exist and are visible, but not in foreground
    ExplorerState::VisibleNotFocused
}

/// Check if cursor is currently over an Explorer window (regardless of foreground)
/// This is the expensive check that uses COM
fn is_cursor_over_explorer_full() -> bool {
    unsafe {
        let mut cursor_pos = POINT::default();
        if GetCursorPos(&mut cursor_pos).is_err() {
            return false;
        }

        // Get window under cursor
        let hwnd = WindowFromPoint(cursor_pos);
        if hwnd.is_invalid() {
            return false;
        }

        // Walk up parent windows to find Explorer window
        let mut current_hwnd = hwnd;
        let folders = get_all_explorer_folders();

        for _ in 0..20 {
            // Check if this window is an Explorer window
            for (explorer_hwnd, _) in &folders {
                if current_hwnd == *explorer_hwnd {
                    return true;
                }
            }
            
            // Also check by class/process
            if is_explorer_window(current_hwnd) {
                return true;
            }

            // Get parent
            if let Ok(parent) = windows::Win32::UI::WindowsAndMessaging::GetParent(current_hwnd) {
                if parent.is_invalid() || parent == current_hwnd {
                    break;
                }
                current_hwnd = parent;
            } else {
                break;
            }
        }
    }
    false
}

fn is_explorer_window(hwnd: HWND) -> bool {
    unsafe {
        let mut class_name = [0u16; 256];
        let len = GetClassNameW(hwnd, &mut class_name);
        let class_str = if len > 0 {
            OsString::from_wide(&class_name[..len as usize])
                .to_string_lossy()
                .to_lowercase()
        } else {
            String::new()
        };

        // Check for common Explorer window classes
        if class_str.contains("cabinetwclass") || class_str.contains("explorerwclass") {
            return true;
        }

        // Fallback: check process name
        let mut process_id: u32 = 0;
        GetWindowThreadProcessId(hwnd, Some(&mut process_id));

        if let Ok(handle) = windows::Win32::System::Threading::OpenProcess(
            windows::Win32::System::Threading::PROCESS_QUERY_LIMITED_INFORMATION,
            false,
            process_id,
        ) {
            let mut buffer = [0u16; 260];
            let mut size = buffer.len() as u32;
            if windows::Win32::System::Threading::QueryFullProcessImageNameW(
                handle,
                windows::Win32::System::Threading::PROCESS_NAME_WIN32,
                windows::core::PWSTR(buffer.as_mut_ptr()),
                &mut size,
            )
            .is_ok()
            {
                let path = OsString::from_wide(&buffer[..size as usize]);
                let path_str = path.to_string_lossy().to_lowercase();
                let _ = windows::Win32::Foundation::CloseHandle(handle);
                return path_str.contains("explorer.exe");
            }
            let _ = windows::Win32::Foundation::CloseHandle(handle);
        }
    }
    false
}

/// Main loop for explorer hook
pub fn run_explorer_hook() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

    let mut last_file: Option<PathBuf> = None;
    let mut hover_start: Option<Instant> = None;
    let mut last_cursor_pos = POINT::default();
    
    // State for optimized polling
    let mut last_state_check = Instant::now();
    let mut current_state = ExplorerState::NoExplorerWindows;
    
    // Polling intervals based on state
    const DEEP_SLEEP_MS: u64 = 1000;   // No Explorer windows - check once per second
    const LONG_SLEEP_MS: u64 = 500;    // All minimized or hidden - check twice per second
    const MEDIUM_SLEEP_MS: u64 = 150;  // Visible but not focused - moderate checking
    const ACTIVE_POLL_MS: u64 = 30;    // Active focus - responsive polling
    
    // How often to re-evaluate the state when in sleep modes
    const STATE_RECHECK_DEEP_MS: u64 = 2000;    // When no Explorer windows
    const STATE_RECHECK_LONG_MS: u64 = 1000;    // When minimized/hidden
    const STATE_RECHECK_MEDIUM_MS: u64 = 300;   // When visible but not focused
    const STATE_RECHECK_ACTIVE_MS: u64 = 100;   // When active

    while RUNNING.load(Ordering::SeqCst) {
        // Check if preview is enabled
        let (preview_enabled, hover_delay_ms) = CONFIG
            .lock()
            .map(|c| (c.preview_enabled, c.hover_delay_ms))
            .unwrap_or((true, 0));
        
        if !preview_enabled {
            if last_file.is_some() {
                hide_preview();
                last_file = None;
                hover_start = None;
            }
            // Sleep longer when disabled
            std::thread::sleep(Duration::from_millis(LONG_SLEEP_MS));
            continue;
        }
        
        let hover_delay = Duration::from_millis(hover_delay_ms);

        // Determine sleep duration and whether to recheck state based on current state
        let (sleep_ms, state_recheck_ms) = match current_state {
            ExplorerState::NoExplorerWindows => (DEEP_SLEEP_MS, STATE_RECHECK_DEEP_MS),
            ExplorerState::AllMinimized => (LONG_SLEEP_MS, STATE_RECHECK_LONG_MS),
            ExplorerState::HiddenByForeground => (LONG_SLEEP_MS, STATE_RECHECK_LONG_MS),
            ExplorerState::VisibleNotFocused => (MEDIUM_SLEEP_MS, STATE_RECHECK_MEDIUM_MS),
            ExplorerState::ActiveFocus => (ACTIVE_POLL_MS, STATE_RECHECK_ACTIVE_MS),
        };
        
        // Periodically re-evaluate the state
        if last_state_check.elapsed() > Duration::from_millis(state_recheck_ms) {
            current_state = get_explorer_state();
            last_state_check = Instant::now();
        }
        
        // If Explorer is not accessible, hide preview and sleep
        match current_state {
            ExplorerState::NoExplorerWindows 
            | ExplorerState::AllMinimized 
            | ExplorerState::HiddenByForeground => {
                if last_file.is_some() {
                    hide_preview();
                    last_file = None;
                    hover_start = None;
                }
                std::thread::sleep(Duration::from_millis(sleep_ms));
                continue;
            }
            ExplorerState::VisibleNotFocused => {
                // Explorer is visible but not focused - do a quick cursor check
                // Only activate full polling if cursor is actually over Explorer
                if !is_cursor_over_explorer_full() {
                    if last_file.is_some() {
                        hide_preview();
                        last_file = None;
                        hover_start = None;
                    }
                    std::thread::sleep(Duration::from_millis(sleep_ms));
                    continue;
                }
                // Cursor is over Explorer, switch to active state
                current_state = ExplorerState::ActiveFocus;
            }
            ExplorerState::ActiveFocus => {
                // Continue with active polling below
            }
        }

        // Explorer is active - use faster polling
        std::thread::sleep(Duration::from_millis(ACTIVE_POLL_MS));

        unsafe {
            // Get cursor position
            let mut cursor_pos = POINT::default();
            if GetCursorPos(&mut cursor_pos).is_err() {
                continue;
            }

            // If cursor moved significantly, check what's under it
            let moved = (cursor_pos.x - last_cursor_pos.x).abs() > 5
                || (cursor_pos.y - last_cursor_pos.y).abs() > 5;

            if moved {
                last_cursor_pos = cursor_pos;
                
                // When cursor moves, check immediately what file is under it
                if let Some(file_path) = get_file_under_cursor() {
                    if last_file.as_ref() == Some(&file_path) {
                        // Same file - keep preview
                        continue;
                    } else {
                        // Different file - hide and start new hover timer
                        hide_preview();
                        last_file = None;
                        hover_start = Some(Instant::now());
                    }
                } else {
                    // No file under cursor - hide preview
                    if last_file.is_some() {
                        hide_preview();
                        last_file = None;
                    }
                    hover_start = Some(Instant::now());
                }
                continue;
            }

            // Check if we've hovered long enough
            if let Some(start) = hover_start {
                if start.elapsed() >= hover_delay {
                    // Try to get file under cursor
                    if let Some(file_path) = get_file_under_cursor() {
                        if last_file.as_ref() != Some(&file_path) {
                            last_file = Some(file_path.clone());
                            show_preview(&file_path, cursor_pos.x, cursor_pos.y);
                        }
                    } else {
                        // No file found, hide preview
                        if last_file.is_some() {
                            hide_preview();
                            last_file = None;
                        }
                    }
                }
            } else {
                // Initialize hover_start if not moving
                hover_start = Some(Instant::now());
            }
        }
    }
}

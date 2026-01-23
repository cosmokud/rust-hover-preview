use crate::preview_window::{hide_preview, show_preview};
use crate::RUNNING;
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use std::path::PathBuf;
use std::sync::atomic::Ordering;
use windows::core::{Interface, VARIANT};
use windows::Win32::Foundation::{HWND, POINT, SHANDLE_PTR};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CLSCTX_ALL, COINIT_APARTMENTTHREADED,
};
use windows::Win32::UI::Shell::{
    IShellWindows, ShellWindows,
};
use windows::Win32::UI::WindowsAndMessaging::{
    GetClassNameW, GetCursorPos, GetForegroundWindow, GetWindowThreadProcessId, WindowFromPoint,
};

// Supported image extensions
const IMAGE_EXTENSIONS: &[&str] = &[
    "jpg", "jpeg", "png", "gif", "bmp", "ico", "tiff", "tif", "webp",
];

fn is_image_file(path: &PathBuf) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| IMAGE_EXTENSIONS.contains(&ext.to_lowercase().as_str()))
        .unwrap_or(false)
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

/// Try to find an image file under the cursor
fn get_file_under_cursor() -> Option<PathBuf> {
    // Get the current Explorer folder
    let folder = get_current_explorer_folder();
    
    // Get the item name under cursor
    let item_name = get_item_name_under_cursor();

    // Debug logging to a file
    if let Ok(mut file) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open("C:\\temp\\hover_preview_debug.log")
    {
        use std::io::Write;
        let _ = writeln!(file, "Folder: {:?}, Item: {:?}", folder, item_name);
    }

    let folder = folder?;
    let item_name = item_name?;

    // Try to construct the full path
    // The item_name might be just the filename or include the extension
    let folder_path = PathBuf::from(&folder);

    // First try: item_name as-is
    let full_path = folder_path.join(&item_name);
    
    // Debug: log the full path attempt
    if let Ok(mut file) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open("C:\\temp\\hover_preview_debug.log")
    {
        use std::io::Write;
        let _ = writeln!(file, "Trying path: {:?}, exists: {}, is_image: {}", 
            full_path, full_path.exists(), is_image_file(&full_path));
    }
    
    if full_path.exists() && is_image_file(&full_path) {
        return Some(full_path);
    }

    // Second try: search for files that match this name (with or without extension)
    if let Ok(entries) = std::fs::read_dir(&folder_path) {
        for entry in entries.flatten() {
            let path = entry.path();
            
            // Check full filename match (e.g., "image.jpg")
            if let Some(file_name) = path.file_name().and_then(|s| s.to_str()) {
                if file_name == item_name && is_image_file(&path) {
                    return Some(path);
                }
            }
            
            // Check file stem match (e.g., "image" matches "image.jpg")
            if let Some(file_stem) = path.file_stem().and_then(|s| s.to_str()) {
                if file_stem == item_name && is_image_file(&path) {
                    return Some(path);
                }
            }
            
            // Check case-insensitive match
            if let Some(file_name) = path.file_name().and_then(|s| s.to_str()) {
                if file_name.to_lowercase() == item_name.to_lowercase() && is_image_file(&path) {
                    return Some(path);
                }
            }
        }
    }

    None
}

/// Check if cursor is currently over an Explorer window (regardless of foreground)
fn is_cursor_over_explorer() -> bool {
    unsafe {
        let mut cursor_pos = POINT::default();
        if GetCursorPos(&mut cursor_pos).is_err() {
            return false;
        }

        // Get window under cursor
        let hwnd = WindowFromPoint(cursor_pos);
        if hwnd.0.is_null() {
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
                if parent.0.is_null() || parent == current_hwnd {
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

        // Log for debugging
        if let Ok(mut file) = std::fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open("C:\\temp\\hover_preview_debug.log")
        {
            use std::io::Write;
            let _ = writeln!(file, "Window class: '{}'", class_str);
        }

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
                
                // Log process name
                if let Ok(mut file) = std::fs::OpenOptions::new()
                    .create(true)
                    .append(true)
                    .open("C:\\temp\\hover_preview_debug.log")
                {
                    use std::io::Write;
                    let _ = writeln!(file, "Process: '{}'", path_str);
                }
                
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
    let mut hover_start: Option<std::time::Instant> = None;
    let hover_delay = std::time::Duration::from_millis(500);
    let mut last_cursor_pos = POINT::default();
    let mut log_counter = 0u32;

    while RUNNING.load(Ordering::SeqCst) {
        std::thread::sleep(std::time::Duration::from_millis(50));

        unsafe {
            // Log every ~2 seconds for debugging
            log_counter += 1;
            if log_counter % 40 == 0 {
                let over_explorer = is_cursor_over_explorer();
                if let Ok(mut file) = std::fs::OpenOptions::new()
                    .create(true)
                    .append(true)
                    .open("C:\\temp\\hover_preview_debug.log")
                {
                    use std::io::Write;
                    let _ = writeln!(file, "Cursor over explorer: {}", over_explorer);
                }
            }
            
            // Check if cursor is over any Explorer window (not just foreground)
            if !is_cursor_over_explorer() {
                if last_file.is_some() {
                    hide_preview();
                    last_file = None;
                    hover_start = None;
                }
                continue;
            }

            // Check cursor position
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
                // If it's the same file, keep the preview; if different or none, handle accordingly
                if let Some(file_path) = get_file_under_cursor() {
                    if last_file.as_ref() == Some(&file_path) {
                        // Same file - keep preview, no need to reset timer
                        continue;
                    } else {
                        // Different file - hide and start new hover timer
                        hide_preview();
                        last_file = None;
                        hover_start = Some(std::time::Instant::now());
                    }
                } else {
                    // No file under cursor - hide preview
                    if last_file.is_some() {
                        hide_preview();
                        last_file = None;
                    }
                    hover_start = Some(std::time::Instant::now());
                }
                continue;
            }

            // Check if we've hovered long enough
            if let Some(start) = hover_start {
                if start.elapsed() >= hover_delay {
                    // Try to get file under cursor
                    if let Some(file_path) = get_file_under_cursor() {
                        // Debug log
                        if let Ok(mut logfile) = std::fs::OpenOptions::new()
                            .create(true)
                            .append(true)
                            .open("C:\\temp\\hover_preview_debug.log")
                        {
                            use std::io::Write;
                            let _ = writeln!(logfile, "Found file: {:?}, last_file: {:?}, match: {}", 
                                file_path, last_file, last_file.as_ref() == Some(&file_path));
                        }
                        
                        if last_file.as_ref() != Some(&file_path) {
                            last_file = Some(file_path.clone());
                            // Debug log - about to call show_preview
                            if let Ok(mut logfile) = std::fs::OpenOptions::new()
                                .create(true)
                                .append(true)
                                .open("C:\\temp\\hover_preview_debug.log")
                            {
                                use std::io::Write;
                                let _ = writeln!(logfile, "Calling show_preview!");
                            }
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
                hover_start = Some(std::time::Instant::now());
            }
        }
    }
}

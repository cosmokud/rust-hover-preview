use crate::preview_window::{
    hide_preview, is_cursor_over_image_preview, is_cursor_over_video_preview, show_preview,
    show_preview_keyboard,
};
use crate::{CONFIG, RUNNING};
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use std::path::{Path, PathBuf};
use std::sync::{atomic::Ordering, Mutex};
use std::time::{Duration, Instant};
use windows::core::{Interface, VARIANT};
use windows::Win32::Foundation::{HWND, POINT, RECT};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL, COINIT_APARTMENTTHREADED,
};
use windows::Win32::System::Variant::VariantClear;
use windows::Win32::UI::Accessibility::{CUIAutomation, IUIAutomation};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetAsyncKeyState, VK_DOWN, VK_END, VK_HOME, VK_LEFT, VK_NEXT, VK_PRIOR, VK_RIGHT, VK_UP,
};
use windows::Win32::UI::Shell::{IShellFolderViewDual, IShellWindows, ShellWindows};
use windows::Win32::UI::WindowsAndMessaging::{
    GetClassNameW, GetCursorPos, GetForegroundWindow, GetWindowPlacement, GetWindowRect,
    GetWindowThreadProcessId, IsIconic, IsWindowVisible, WindowFromPoint, SW_SHOWMAXIMIZED,
    WINDOWPLACEMENT,
};

// Supported image extensions
const IMAGE_EXTENSIONS: &[&str] = &[
    "jpg", "jpeg", "jpe", "jfif", "png", "gif", "bmp", "ico", "tiff", "tif", "webp",
];

// Supported video extensions
const VIDEO_EXTENSIONS: &[&str] = &["mp4", "webm", "mkv", "avi", "mov", "wmv", "flv", "m4v"];

struct FolderMediaIndex {
    folder: String,
    built_at: Instant,
    by_file_name: HashMap<String, PathBuf>,
    by_stem: HashMap<String, PathBuf>,
}

struct ExplorerFoldersCache {
    built_at: Instant,
    folders: Vec<(isize, String)>,
}

struct ShellViewMediaIndex {
    hwnd: isize,
    built_at: Instant,
    by_display_name: HashMap<String, PathBuf>,
    by_file_name: HashMap<String, PathBuf>,
}

const FOLDER_INDEX_TTL_MS: u64 = 60000;
const EXPLORER_FOLDERS_CACHE_TTL_MS: u64 = 250;
const SHELL_VIEW_INDEX_TTL_MS: u64 = 500;
const SHELL_VIEW_INDEX_MAX_ITEMS: i32 = 5000;

static FOLDER_MEDIA_INDEX: Lazy<Mutex<Option<FolderMediaIndex>>> = Lazy::new(|| Mutex::new(None));
static EXPLORER_FOLDERS_CACHE: Lazy<Mutex<Option<ExplorerFoldersCache>>> =
    Lazy::new(|| Mutex::new(None));
static SHELL_VIEW_MEDIA_INDEX: Lazy<Mutex<Option<ShellViewMediaIndex>>> =
    Lazy::new(|| Mutex::new(None));

fn is_jpeg_extension(ext: &str) -> bool {
    matches!(ext, "jpg" | "jpeg" | "jpe" | "jfif")
}

fn clear_variant(variant: &mut VARIANT) {
    unsafe {
        let _ = VariantClear(variant as *mut VARIANT);
    }
}

fn build_folder_media_index(folder_path: &PathBuf, folder_key: &str) -> Option<FolderMediaIndex> {
    let mut by_file_name = HashMap::new();
    let mut by_stem = HashMap::new();

    let entries = std::fs::read_dir(folder_path).ok()?;
    for entry in entries.flatten() {
        let path = entry.path();
        if !path.is_file() || !is_media_file(&path) {
            continue;
        }

        if let Some(file_name) = path.file_name().and_then(|s| s.to_str()) {
            by_file_name
                .entry(file_name.to_ascii_lowercase())
                .or_insert_with(|| path.clone());
        }

        if let Some(stem) = path.file_stem().and_then(|s| s.to_str()) {
            by_stem
                .entry(stem.to_ascii_lowercase())
                .or_insert(path.clone());
        }
    }

    Some(FolderMediaIndex {
        folder: folder_key.to_string(),
        built_at: Instant::now(),
        by_file_name,
        by_stem,
    })
}

fn lookup_media_in_folder_index(
    folder_path: &PathBuf,
    folder_key: &str,
    item_name: &str,
) -> Option<PathBuf> {
    let item_name = item_name.trim();
    if item_name.is_empty() {
        return None;
    }

    let item_name_lower = item_name.to_ascii_lowercase();
    let item_stem_lower = Path::new(item_name)
        .file_stem()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());
    let item_ext_lower = Path::new(item_name)
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());

    let mut cache = FOLDER_MEDIA_INDEX.lock().ok()?;
    let needs_rebuild = match cache.as_ref() {
        Some(index) => {
            index.folder != folder_key
                || index.built_at.elapsed() > Duration::from_millis(FOLDER_INDEX_TTL_MS)
        }
        None => true,
    };

    if needs_rebuild {
        *cache = build_folder_media_index(folder_path, folder_key);
    }

    let index = cache.as_ref()?;

    if let Some(path) = index.by_file_name.get(&item_name_lower) {
        return Some(path.clone());
    }

    if let Some(stem_key) = item_stem_lower.as_ref() {
        if let Some(path) = index.by_stem.get(stem_key) {
            if let Some(item_ext) = item_ext_lower.as_deref() {
                if let Some(candidate_ext) = path.extension().and_then(|s| s.to_str()) {
                    let candidate_ext_lower = candidate_ext.to_ascii_lowercase();
                    if candidate_ext_lower == item_ext
                        || (is_jpeg_extension(&candidate_ext_lower) && is_jpeg_extension(item_ext))
                    {
                        return Some(path.clone());
                    }
                }
            } else {
                return Some(path.clone());
            }
        }
    }

    if item_ext_lower.is_none() {
        if let Some(path) = index.by_stem.get(&item_name_lower) {
            return Some(path.clone());
        }
    }

    None
}

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
    if let Ok(cache) = EXPLORER_FOLDERS_CACHE.lock() {
        if let Some(cache_entry) = cache.as_ref() {
            if cache_entry.built_at.elapsed()
                <= Duration::from_millis(EXPLORER_FOLDERS_CACHE_TTL_MS)
            {
                return cache_entry
                    .folders
                    .iter()
                    .map(|(hwnd, folder)| (HWND(*hwnd as *mut _), folder.clone()))
                    .collect();
            }
        }
    }

    let mut result: Vec<(isize, String)> = Vec::new();

    unsafe {
        if let Ok(shell_windows) =
            CoCreateInstance::<_, IShellWindows>(&ShellWindows, None, CLSCTX_ALL)
        {
            if let Ok(count) = shell_windows.Count() {
                for i in 0..count {
                    let variant = VARIANT::from(i);
                    if let Ok(disp) = shell_windows.Item(&variant) {
                        if let Ok(browser) = disp.cast::<windows::Win32::UI::Shell::IWebBrowser2>()
                        {
                            if let Ok(browser_hwnd) = browser.HWND() {
                                let hwnd = browser_hwnd.0 as isize;
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

    if let Ok(mut cache) = EXPLORER_FOLDERS_CACHE.lock() {
        *cache = Some(ExplorerFoldersCache {
            built_at: Instant::now(),
            folders: result.clone(),
        });
    }

    result
        .into_iter()
        .map(|(hwnd, folder)| (HWND(hwnd as *mut _), folder))
        .collect()
}

fn get_explorer_hwnd_under_cursor_or_foreground() -> Option<HWND> {
    unsafe {
        let mut cursor_pos = POINT::default();
        if GetCursorPos(&mut cursor_pos).is_err() {
            return None;
        }

        let hwnd = WindowFromPoint(cursor_pos);
        if !hwnd.0.is_null() {
            let mut current_hwnd = hwnd;
            for _ in 0..20 {
                if is_explorer_window(current_hwnd) {
                    return Some(current_hwnd);
                }

                if let Ok(parent) = windows::Win32::UI::WindowsAndMessaging::GetParent(current_hwnd)
                {
                    if parent.0.is_null() || parent == current_hwnd {
                        break;
                    }
                    current_hwnd = parent;
                } else {
                    break;
                }
            }
        }

        let foreground = GetForegroundWindow();
        if !foreground.is_invalid() && is_explorer_window(foreground) {
            return Some(foreground);
        }
    }

    None
}

fn build_shell_view_media_index(hwnd: HWND) -> Option<ShellViewMediaIndex> {
    let hwnd_key = hwnd.0 as isize;
    let mut by_display_name = HashMap::new();
    let mut by_file_name = HashMap::new();

    unsafe {
        let shell_windows =
            CoCreateInstance::<_, IShellWindows>(&ShellWindows, None, CLSCTX_ALL).ok()?;
        let count = shell_windows.Count().ok()?;

        for i in 0..count {
            let variant = VARIANT::from(i);
            let disp = match shell_windows.Item(&variant) {
                Ok(disp) => disp,
                Err(_) => continue,
            };
            let browser = match disp.cast::<windows::Win32::UI::Shell::IWebBrowser2>() {
                Ok(browser) => browser,
                Err(_) => continue,
            };
            let browser_hwnd = match browser.HWND() {
                Ok(browser_hwnd) => browser_hwnd,
                Err(_) => continue,
            };
            if browser_hwnd.0 != hwnd_key as isize {
                continue;
            }

            let document = browser.Document().ok()?;
            let shell_view = document.cast::<IShellFolderViewDual>().ok()?;
            let folder = shell_view.Folder().ok()?;
            let items = folder.Items().ok()?;
            let item_count = items.Count().ok()?.min(SHELL_VIEW_INDEX_MAX_ITEMS);

            for item_index in 0..item_count {
                let item_variant = VARIANT::from(item_index);
                let item = match items.Item(&item_variant) {
                    Ok(item) => item,
                    Err(_) => continue,
                };

                let path_str = match item.Path() {
                    Ok(path) => path.to_string(),
                    Err(_) => continue,
                };
                let path = PathBuf::from(path_str);
                if !path.exists() || !is_media_file(&path) {
                    continue;
                }

                if let Ok(name) = item.Name() {
                    by_display_name
                        .entry(name.to_string().to_ascii_lowercase())
                        .or_insert_with(|| path.clone());
                }

                if let Some(file_name) = path.file_name().and_then(|s| s.to_str()) {
                    by_file_name
                        .entry(file_name.to_ascii_lowercase())
                        .or_insert_with(|| path.clone());
                }
            }

            return Some(ShellViewMediaIndex {
                hwnd: hwnd_key,
                built_at: Instant::now(),
                by_display_name,
                by_file_name,
            });
        }
    }

    None
}

fn find_media_in_shell_view(hwnd: HWND, item_name: &str) -> Option<PathBuf> {
    let item_name = item_name.trim();
    if item_name.is_empty() {
        return None;
    }

    let hwnd_key = hwnd.0 as isize;
    if let Ok(cache) = SHELL_VIEW_MEDIA_INDEX.lock() {
        if let Some(index) = cache.as_ref() {
            if index.hwnd == hwnd_key
                && index.built_at.elapsed() <= Duration::from_millis(SHELL_VIEW_INDEX_TTL_MS)
            {
                let item_name_lower = item_name.to_ascii_lowercase();
                return index
                    .by_file_name
                    .get(&item_name_lower)
                    .or_else(|| index.by_display_name.get(&item_name_lower))
                    .cloned();
            }
        }
    }

    let index = build_shell_view_media_index(hwnd)?;
    let item_name_lower = item_name.to_ascii_lowercase();
    let result = index
        .by_file_name
        .get(&item_name_lower)
        .or_else(|| index.by_display_name.get(&item_name_lower))
        .cloned();

    if let Ok(mut cache) = SHELL_VIEW_MEDIA_INDEX.lock() {
        *cache = Some(index);
    }

    result
}

fn find_media_in_current_shell_view(item_name: &str) -> Option<PathBuf> {
    let hwnd = get_explorer_hwnd_under_cursor_or_foreground()?;
    find_media_in_shell_view(hwnd, item_name)
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

/// Names that indicate container elements, not actual files
const CONTAINER_NAMES: &[&str] = &[
    "Items View",
    "Folder View",
    "Shell Folder View",
    "ShellView",
    "UIItemsView",
    "DirectUIHWND",
    "Search Results",
    "File list",
    "Name",
    "Date modified",
    "Type",
    "Size",
    "Date",
    "Date created",
    "Details",
    "List",
    "Content",
    "Tiles",
    "Large icons",
    "Medium icons",
    "Small icons",
    "Extra large icons",
    "Item",
    "Group",
    "Header",
];

/// Patterns that suggest a value might be a folder path rather than a file
const FOLDER_PATTERNS: &[&str] = &["search-ms:", "shell:", "::{"];

/// Check if a name is a container/UI element name rather than an actual file
fn is_container_name(name: &str) -> bool {
    if name.is_empty() {
        return true;
    }
    CONTAINER_NAMES
        .iter()
        .any(|&c| name.eq_ignore_ascii_case(c))
}

/// Check if a value looks like a valid file path (not a shell special path)
fn is_valid_file_path(s: &str) -> bool {
    if s.is_empty() {
        return false;
    }
    // Skip shell special paths
    for pattern in FOLDER_PATTERNS {
        if s.to_lowercase().contains(pattern) {
            return false;
        }
    }
    // Check if it looks like a file path
    let path = PathBuf::from(s);
    path.is_absolute()
}

/// Result from accessibility query - can be a filename or a full path
#[derive(Debug, Clone)]
enum AccessibilityResult {
    /// Just a filename (need to find in folder)
    FileName(String),
    /// Full path to file (from search results)
    FullPath(PathBuf),
}

/// Check whether an accessibility element/child variant contains the current cursor point.
/// This prevents fallbacks from returning focused/default items that are not truly hovered.
fn is_variant_under_cursor(
    acc: &windows::Win32::UI::Accessibility::IAccessible,
    variant: &VARIANT,
    cursor_pos: &POINT,
) -> bool {
    unsafe {
        let mut left = 0;
        let mut top = 0;
        let mut width = 0;
        let mut height = 0;

        if acc
            .accLocation(&mut left, &mut top, &mut width, &mut height, variant)
            .is_err()
        {
            return false;
        }

        if width <= 0 || height <= 0 {
            return false;
        }

        let right = left.saturating_add(width);
        let bottom = top.saturating_add(height);

        cursor_pos.x >= left && cursor_pos.x < right && cursor_pos.y >= top && cursor_pos.y < bottom
    }
}

/// Get the filename or full path under cursor using accessibility - try multiple approaches
fn get_item_under_cursor() -> Option<AccessibilityResult> {
    unsafe {
        let mut cursor_pos = POINT::default();
        if GetCursorPos(&mut cursor_pos).is_err() {
            return None;
        }

        // Use accessibility to get the item info
        let mut accessible: Option<windows::Win32::UI::Accessibility::IAccessible> = None;
        let mut child_variant = VARIANT::default();

        let result = (|| -> Option<AccessibilityResult> {
            if windows::Win32::UI::Accessibility::AccessibleObjectFromPoint(
                cursor_pos,
                &mut accessible,
                &mut child_variant,
            )
            .is_err()
            {
                return None;
            }

            if let Some(ref acc) = accessible {
                // First, try to get the value - this often contains the full path in search results
                if is_variant_under_cursor(acc, &child_variant, &cursor_pos) {
                    if let Ok(value) = acc.get_accValue(&child_variant) {
                        let value_str = value.to_string();
                        // Check if it's a valid full path (not shell:, search-ms:, etc.)
                        if is_valid_file_path(&value_str) {
                            let path = PathBuf::from(&value_str);
                            if path.exists() && is_media_file(&path) {
                                return Some(AccessibilityResult::FullPath(path));
                            }
                        }
                    }
                }

                // Try with the child variant first for name
                if is_variant_under_cursor(acc, &child_variant, &cursor_pos) {
                    if let Ok(name) = acc.get_accName(&child_variant) {
                        let name_str = name.to_string();
                        if !is_container_name(&name_str) {
                            // Check if the name itself is a full path (can happen in search)
                            if is_valid_file_path(&name_str) {
                                let path = PathBuf::from(&name_str);
                                if path.exists() && is_media_file(&path) {
                                    return Some(AccessibilityResult::FullPath(path));
                                }
                            }
                            return Some(AccessibilityResult::FileName(name_str));
                        }
                    }
                }

                // Try with default variant
                let default_variant = VARIANT::default();
                if is_variant_under_cursor(acc, &default_variant, &cursor_pos) {
                    if let Ok(name) = acc.get_accName(&default_variant) {
                        let name_str = name.to_string();
                        if !is_container_name(&name_str) {
                            // Check if the name itself is a full path
                            if is_valid_file_path(&name_str) {
                                let path = PathBuf::from(&name_str);
                                if path.exists() && is_media_file(&path) {
                                    return Some(AccessibilityResult::FullPath(path));
                                }
                            }
                            return Some(AccessibilityResult::FileName(name_str));
                        }
                    }
                }

                // Try navigating parent chain to find item name (for list/details views)
                if let Some(result) = try_get_item_from_parent(acc, &child_variant, &cursor_pos) {
                    return Some(result);
                }

                // Try getting help text which sometimes has info
                if is_variant_under_cursor(acc, &child_variant, &cursor_pos) {
                    if let Ok(help) = acc.get_accHelp(&child_variant) {
                        let help_str = help.to_string();
                        if !help_str.is_empty() && !is_container_name(&help_str) {
                            return Some(AccessibilityResult::FileName(help_str));
                        }
                    }
                }

                // Try description which may have path info
                if is_variant_under_cursor(acc, &child_variant, &cursor_pos) {
                    if let Ok(desc) = acc.get_accDescription(&child_variant) {
                        let desc_str = desc.to_string();
                        // Check for path in description
                        if is_valid_file_path(&desc_str) {
                            let path = PathBuf::from(&desc_str);
                            if path.exists() && is_media_file(&path) {
                                return Some(AccessibilityResult::FullPath(path));
                            }
                        }
                    }
                }

                // Try to walk up parent hierarchy more aggressively (for details view text cells)
                if let Some(result) = try_deep_parent_search(acc, &cursor_pos) {
                    return Some(result);
                }
            }

            None
        })();

        clear_variant(&mut child_variant);
        return result;
    }
}

/// Try to get item info by navigating the accessibility parent chain
/// This helps with List/Details views where hovering over filename text doesn't directly give the name
fn try_get_item_from_parent(
    acc: &windows::Win32::UI::Accessibility::IAccessible,
    _child_variant: &VARIANT,
    cursor_pos: &POINT,
) -> Option<AccessibilityResult> {
    unsafe {
        // Try to get parent accessible object
        if let Ok(parent_disp) = acc.accParent() {
            if let Ok(parent_acc) =
                parent_disp.cast::<windows::Win32::UI::Accessibility::IAccessible>()
            {
                let default_variant = VARIANT::default();

                // Try to get name from parent
                if is_variant_under_cursor(&parent_acc, &default_variant, cursor_pos) {
                    if let Ok(name) = parent_acc.get_accName(&default_variant) {
                        let name_str = name.to_string();
                        if !is_container_name(&name_str) {
                            return Some(AccessibilityResult::FileName(name_str));
                        }
                    }
                }

                // Try to get value (path) from parent
                if is_variant_under_cursor(&parent_acc, &default_variant, cursor_pos) {
                    if let Ok(value) = parent_acc.get_accValue(&default_variant) {
                        let value_str = value.to_string();
                        if !value_str.is_empty() {
                            let path = PathBuf::from(&value_str);
                            if path.is_absolute() && path.exists() && is_media_file(&path) {
                                return Some(AccessibilityResult::FullPath(path));
                            }
                        }
                    }
                }

                // Try child enumeration to find focused/selected item
                if let Some(result) = try_get_focused_child(&parent_acc, cursor_pos) {
                    return Some(result);
                }
            }
        }

        // Try getting focused element within the accessible object
        if let Ok(mut focus) = acc.accFocus() {
            let focus_result = (|| -> Option<AccessibilityResult> {
                // If focus returns a variant with child ID
                let vt = focus.as_raw().Anonymous.Anonymous.vt;
                if vt == windows::Win32::System::Variant::VT_I4.0 {
                    if is_variant_under_cursor(acc, &focus, cursor_pos) {
                        if let Ok(name) = acc.get_accName(&focus) {
                            let name_str = name.to_string();
                            if !is_container_name(&name_str) {
                                return Some(AccessibilityResult::FileName(name_str));
                            }
                        }
                    }
                }

                None
            })();

            clear_variant(&mut focus);
            if focus_result.is_some() {
                return focus_result;
            }
        }
    }
    None
}

/// Try to find focused/hot-tracked child in accessibility tree
fn try_get_focused_child(
    acc: &windows::Win32::UI::Accessibility::IAccessible,
    cursor_pos: &POINT,
) -> Option<AccessibilityResult> {
    unsafe {
        // Get child count
        if let Ok(count) = acc.accChildCount() {
            // Limit iteration to prevent hanging
            let max_check = (count as i32).min(100);

            for i in 1..=max_check {
                let child_var = VARIANT::from(i);

                if !is_variant_under_cursor(acc, &child_var, cursor_pos) {
                    continue;
                }

                // Check state for focus/hot tracking
                if let Ok(state) = acc.get_accState(&child_var) {
                    let state_val = state.as_raw().Anonymous.Anonymous.Anonymous.uintVal;
                    // STATE_SYSTEM_HOTTRACKED = 0x80, STATE_SYSTEM_FOCUSED = 0x4
                    if (state_val & 0x80) != 0 || (state_val & 0x4) != 0 {
                        if let Ok(name) = acc.get_accName(&child_var) {
                            let name_str = name.to_string();
                            if !is_container_name(&name_str) {
                                return Some(AccessibilityResult::FileName(name_str));
                            }
                        }

                        // Also try value for full path
                        if let Ok(value) = acc.get_accValue(&child_var) {
                            let value_str = value.to_string();
                            if !value_str.is_empty() {
                                let path = PathBuf::from(&value_str);
                                if path.is_absolute() && path.exists() && is_media_file(&path) {
                                    return Some(AccessibilityResult::FullPath(path));
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

/// Deep parent search - walk up the accessibility tree to find file information
/// This is especially useful for Details/List views where clicking on a text cell
/// gives us the cell, not the row/item
fn try_deep_parent_search(
    acc: &windows::Win32::UI::Accessibility::IAccessible,
    cursor_pos: &POINT,
) -> Option<AccessibilityResult> {
    unsafe {
        let mut current_acc = acc.clone();

        // Walk up to 5 levels of parent hierarchy
        for _ in 0..5 {
            // Try to get parent
            if let Ok(parent_disp) = current_acc.accParent() {
                if let Ok(parent_acc) =
                    parent_disp.cast::<windows::Win32::UI::Accessibility::IAccessible>()
                {
                    let default_variant = VARIANT::default();

                    // Try getting name from parent
                    if is_variant_under_cursor(&parent_acc, &default_variant, cursor_pos) {
                        if let Ok(name) = parent_acc.get_accName(&default_variant) {
                            let name_str = name.to_string();
                            if !is_container_name(&name_str) {
                                // Check if it's a full path
                                if is_valid_file_path(&name_str) {
                                    let path = PathBuf::from(&name_str);
                                    if path.exists() && is_media_file(&path) {
                                        return Some(AccessibilityResult::FullPath(path));
                                    }
                                }
                                // It's a filename
                                return Some(AccessibilityResult::FileName(name_str));
                            }
                        }
                    }

                    // Try getting value from parent (may contain path)
                    if is_variant_under_cursor(&parent_acc, &default_variant, cursor_pos) {
                        if let Ok(value) = parent_acc.get_accValue(&default_variant) {
                            let value_str = value.to_string();
                            if is_valid_file_path(&value_str) {
                                let path = PathBuf::from(&value_str);
                                if path.exists() && is_media_file(&path) {
                                    return Some(AccessibilityResult::FullPath(path));
                                }
                            }
                        }
                    }

                    // Try to find selected/focused child of this parent
                    if let Some(result) = try_get_focused_child(&parent_acc, cursor_pos) {
                        return Some(result);
                    }

                    current_acc = parent_acc;
                } else {
                    break;
                }
            } else {
                break;
            }
        }
    }
    None
}

/// Try to find an image or video file in a specific folder by item name
fn find_media_in_folder(folder: &str, item_name: &str) -> Option<PathBuf> {
    let item_name = item_name.trim();
    if item_name.is_empty() {
        return None;
    }

    let folder_path = PathBuf::from(folder);
    let folder_key = folder_path.to_string_lossy().into_owned();

    // First try: item_name as-is
    let full_path = folder_path.join(item_name);
    if full_path.exists() && is_media_file(&full_path) {
        return Some(full_path);
    }

    // JPEG extension aliases can differ between Explorer labels and on-disk names.
    // Try sibling JPEG aliases before consulting the folder index.
    if let Some(item_ext) = Path::new(item_name)
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase())
    {
        if is_jpeg_extension(&item_ext) {
            if let Some(stem) = Path::new(item_name).file_stem().and_then(|s| s.to_str()) {
                for alt in ["jpg", "jpeg", "jpe", "jfif"] {
                    if alt == item_ext {
                        continue;
                    }
                    let candidate = folder_path.join(format!("{}.{}", stem, alt));
                    if candidate.exists() && is_media_file(&candidate) {
                        return Some(candidate);
                    }
                }
            }
        }
    }

    // Fallback: use a short-lived folder index so large folders are scanned once
    // instead of once per hover poll.
    lookup_media_in_folder_index(&folder_path, &folder_key, item_name)
}

fn get_file_under_cursor() -> Option<PathBuf> {
    // Get the item info under cursor
    let item_info = get_item_under_cursor()?;

    match item_info {
        AccessibilityResult::FullPath(path) => {
            // Already have full path (from search results), verify it's a media file
            if is_media_file(&path) {
                return Some(path);
            }
            None
        }
        AccessibilityResult::FileName(item_name) => {
            // First try the Explorer folder currently under cursor.
            // This avoids accidental matches from other windows/tabs.
            let current_folder = get_current_explorer_folder();
            if let Some(folder) = current_folder.as_ref() {
                if let Some(path) = find_media_in_folder(folder, &item_name) {
                    return Some(path);
                }
            }

            // Search-result views expose virtual/search folders through ShellWindows,
            // while FolderItems still keep each result's real filesystem path.
            if current_folder.is_none() {
                if let Some(path) = find_media_in_current_shell_view(&item_name) {
                    return Some(path);
                }
            }

            // Fallback: search all Explorer folders (all windows and tabs)
            let all_folders = get_all_explorer_folders();

            // Try to find the file in ANY of the open Explorer folders
            for (_, folder) in &all_folders {
                if let Some(path) = find_media_in_folder(folder, &item_name) {
                    return Some(path);
                }
            }

            // Also try treating item_name as a potential full path
            let potential_path = PathBuf::from(&item_name);
            if potential_path.is_absolute()
                && potential_path.exists()
                && is_media_file(&potential_path)
            {
                return Some(potential_path);
            }

            None
        }
    }
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
                        if let Ok(browser) = disp.cast::<windows::Win32::UI::Shell::IWebBrowser2>()
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

/// Detect whether the user is actively navigating Explorer with keyboard.
/// We treat both current-down and "pressed since last check" states as input.
fn is_keyboard_navigation_input_detected() -> bool {
    unsafe {
        let navigation_keys = [
            VK_UP, VK_DOWN, VK_LEFT, VK_RIGHT, VK_HOME, VK_END, VK_PRIOR, VK_NEXT,
        ];

        navigation_keys.iter().any(|&key| {
            let state = GetAsyncKeyState(key.0 as i32) as u16;
            (state & 0x8000) != 0 || (state & 0x0001) != 0
        })
    }
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

/// Information about a focused item from Explorer's accessibility tree
struct FocusedItemInfo {
    result: AccessibilityResult,
    rect: RECT,
}

/// Get the currently focused/selected file in Explorer using UI Automation.
/// UI Automation is far more reliable than MSAA IAccessible for modern Explorer.
fn get_focused_explorer_item(automation: &IUIAutomation) -> Option<FocusedItemInfo> {
    unsafe {
        // Only works when Explorer is the foreground window
        let foreground = GetForegroundWindow();
        if foreground.is_invalid() || !is_explorer_window(foreground) {
            return None;
        }

        // Get the currently focused UI element via UI Automation
        let focused = automation.GetFocusedElement().ok()?;

        // Get the element name (this is the filename in Explorer)
        let name = focused.CurrentName().ok()?.to_string();
        if name.is_empty() || is_container_name(&name) {
            return None;
        }

        // Get bounding rectangle (screen coordinates)
        let rect = focused.CurrentBoundingRectangle().ok()?;
        if rect.right <= rect.left || rect.bottom <= rect.top {
            return None;
        }

        // Check if name is a full path (can happen in search results)
        if is_valid_file_path(&name) {
            let path = PathBuf::from(&name);
            if path.exists() && is_media_file(&path) {
                return Some(FocusedItemInfo {
                    result: AccessibilityResult::FullPath(path),
                    rect,
                });
            }
        }

        Some(FocusedItemInfo {
            result: AccessibilityResult::FileName(name),
            rect,
        })
    }
}

fn resolve_focused_item_to_path(item: &FocusedItemInfo) -> Option<PathBuf> {
    match &item.result {
        AccessibilityResult::FullPath(path) => {
            if is_media_file(path) {
                Some(path.clone())
            } else {
                None
            }
        }
        AccessibilityResult::FileName(item_name) => {
            // Search-result views are virtual folders; resolve through Shell view
            // items before trying normal folder joins.
            if let Some(path) = find_media_in_current_shell_view(item_name) {
                return Some(path);
            }

            // Search all open Explorer folder paths
            let all_folders = get_all_explorer_folders();
            for (_, folder) in &all_folders {
                if let Some(path) = find_media_in_folder(folder, item_name) {
                    return Some(path);
                }
            }
            // Try as a potential full path
            let potential_path = PathBuf::from(item_name);
            if potential_path.is_absolute()
                && potential_path.exists()
                && is_media_file(&potential_path)
            {
                return Some(potential_path);
            }
            None
        }
    }
}

/// Main loop for explorer hook
pub fn run_explorer_hook() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

    // Create UI Automation instance for keyboard focus detection (cached for the lifetime of the loop)
    let uia: Option<IUIAutomation> =
        unsafe { CoCreateInstance(&CUIAutomation, None, CLSCTX_ALL).ok() };

    let mut last_file: Option<PathBuf> = None;
    let mut hover_start: Option<Instant> = None;
    let mut last_cursor_pos = POINT::default();

    // Keyboard hover state
    let mut keyboard_file: Option<PathBuf> = None;
    let mut last_focused_name: Option<String> = None;
    let mut is_keyboard_hover = false;
    // Short grace after starting a video preview to avoid instant self-dismiss
    // while ffplay window is still initializing under the cursor.
    let mut video_hover_guard_until: Option<Instant> = None;
    // Folder/input gate state: suppress preview after folder changes until explicit user input.
    let mut last_cursor_folder: Option<String> = None;
    let mut suspend_preview_until_user_input = false;
    let mut allow_keyboard_preview_on_first_observation = false;
    let mut folder_change_time: Option<Instant> = None;
    let mut suspended_initial_focus: Option<String> = None;
    let mut last_folder_probe = Instant::now();
    let mut last_hover_probe = Instant::now();
    let mut last_keyboard_focus_probe = Instant::now();

    // State for optimized polling
    let mut last_state_check = Instant::now();
    let mut current_state = ExplorerState::NoExplorerWindows;

    // Polling intervals based on state
    const DEEP_SLEEP_MS: u64 = 1000; // No Explorer windows - check once per second
    const LONG_SLEEP_MS: u64 = 500; // All minimized or hidden - check twice per second
    const MEDIUM_SLEEP_MS: u64 = 150; // Visible but not focused - moderate checking
    const ACTIVE_POLL_MS: u64 = 30; // Active focus - responsive polling
    const VIDEO_HOVER_DISMISS_GRACE_MS: u64 = 350;
    const FOLDER_PROBE_MS: u64 = 200;
    const HOVER_PROBE_MS: u64 = 120;
    const KEYBOARD_FOCUS_PROBE_MS: u64 = 80;

    // How often to re-evaluate the state when in sleep modes
    const STATE_RECHECK_DEEP_MS: u64 = 2000; // When no Explorer windows
    const STATE_RECHECK_LONG_MS: u64 = 1000; // When minimized/hidden
    const STATE_RECHECK_MEDIUM_MS: u64 = 300; // When visible but not focused
    const STATE_RECHECK_ACTIVE_MS: u64 = 100; // When active

    while RUNNING.load(Ordering::SeqCst) {
        // Check if preview is enabled
        let (preview_enabled, hover_delay_ms) = CONFIG
            .lock()
            .map(|c| (c.preview_enabled, c.hover_delay_ms))
            .unwrap_or((true, 0));

        if !preview_enabled {
            if last_file.is_some() || keyboard_file.is_some() {
                hide_preview();
                last_file = None;
                hover_start = None;
            }
            keyboard_file = None;
            last_focused_name = None;
            is_keyboard_hover = false;
            video_hover_guard_until = None;
            suspend_preview_until_user_input = false;
            allow_keyboard_preview_on_first_observation = false;
            last_cursor_folder = None;
            folder_change_time = None;
            suspended_initial_focus = None;
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
                if last_file.is_some() || keyboard_file.is_some() {
                    hide_preview();
                    last_file = None;
                    hover_start = None;
                    keyboard_file = None;
                    last_focused_name = None;
                    is_keyboard_hover = false;
                    video_hover_guard_until = None;
                }
                std::thread::sleep(Duration::from_millis(sleep_ms));
                continue;
            }
            ExplorerState::VisibleNotFocused => {
                // Explorer is visible but not focused - do a quick cursor check
                // Only activate full polling if cursor is actually over Explorer
                if !is_cursor_over_explorer_full() {
                    if last_file.is_some() || keyboard_file.is_some() {
                        hide_preview();
                        last_file = None;
                        hover_start = None;
                        keyboard_file = None;
                        last_focused_name = None;
                        is_keyboard_hover = false;
                        video_hover_guard_until = None;
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

            // Detect folder navigation/opening and suspend preview until user input.
            if last_folder_probe.elapsed() >= Duration::from_millis(FOLDER_PROBE_MS) {
                last_folder_probe = Instant::now();
                if let Some(folder) = get_current_explorer_folder() {
                    if last_cursor_folder.as_ref() != Some(&folder) {
                        last_cursor_folder = Some(folder);
                        suspend_preview_until_user_input = true;
                        allow_keyboard_preview_on_first_observation = false;
                        folder_change_time = Some(Instant::now());
                        suspended_initial_focus = None;
                        hover_start = None;
                        last_focused_name = None;
                        // Reset cursor baseline so we don't mistake stale delta for movement.
                        last_cursor_pos = cursor_pos;
                        // Drain stale GetAsyncKeyState flags from prior navigation
                        let _ = is_keyboard_navigation_input_detected();

                        if last_file.is_some() || keyboard_file.is_some() || is_keyboard_hover {
                            hide_preview();
                        }
                        last_file = None;
                        keyboard_file = None;
                        is_keyboard_hover = false;
                        video_hover_guard_until = None;
                    }
                }
            }

            // If cursor moved significantly, check what's under it
            let moved = (cursor_pos.x - last_cursor_pos.x).abs() > 5
                || (cursor_pos.y - last_cursor_pos.y).abs() > 5;

            // Hard gate: after folder change, do not preview until explicit user input.
            if suspend_preview_until_user_input {
                // Cooldown: ignore all input for 150ms after folder change to let
                // COM/accessibility settle and to avoid stale keyboard state.
                if let Some(change_time) = folder_change_time {
                    if change_time.elapsed() < Duration::from_millis(150) {
                        continue;
                    }
                }

                if moved {
                    suspend_preview_until_user_input = false;
                    allow_keyboard_preview_on_first_observation = false;
                    hover_start = Some(Instant::now());
                    suspended_initial_focus = None;
                    folder_change_time = None;
                } else {
                    // Detect real keyboard navigation by observing UI Automation
                    // focus changes, which is far more reliable than GetAsyncKeyState
                    // (whose "pressed since last check" bit can carry stale state
                    // from the navigation that opened this folder).
                    let mut keyboard_unlocked = false;
                    if is_foreground_explorer()
                        && last_keyboard_focus_probe.elapsed()
                            >= Duration::from_millis(KEYBOARD_FOCUS_PROBE_MS)
                    {
                        last_keyboard_focus_probe = Instant::now();
                        if let Some(focused_info) =
                            uia.as_ref().and_then(|a| get_focused_explorer_item(a))
                        {
                            let focused_name = match &focused_info.result {
                                AccessibilityResult::FileName(name) => name.clone(),
                                AccessibilityResult::FullPath(path) => {
                                    path.to_string_lossy().to_string()
                                }
                            };

                            if suspended_initial_focus.is_none() {
                                // Record the auto-focused first item
                                // (set by Windows when folder opens)
                                suspended_initial_focus = Some(focused_name);
                            } else if suspended_initial_focus.as_ref() != Some(&focused_name) {
                                // Focus actually changed — user pressed a navigation key
                                keyboard_unlocked = true;
                            }
                        }
                    }

                    if keyboard_unlocked {
                        suspend_preview_until_user_input = false;
                        allow_keyboard_preview_on_first_observation = true;
                        suspended_initial_focus = None;
                        folder_change_time = None;
                    } else {
                        continue;
                    }
                }
            }

            if moved {
                last_cursor_pos = cursor_pos;

                // Mouse movement always takes priority - dismiss keyboard hover
                if is_keyboard_hover {
                    hide_preview();
                    keyboard_file = None;
                    is_keyboard_hover = false;
                    video_hover_guard_until = None;
                }
                // Reset focused name tracking so keyboard navigation can be re-detected
                // after mouse stops moving
                last_focused_name = None;
                allow_keyboard_preview_on_first_observation = false;

                // Dismiss preview only when the mouse has actually moved onto it.
                // This avoids blocking keyboard navigation when the cursor is static.
                if !is_keyboard_hover {
                    let over_image_preview = is_cursor_over_image_preview();
                    let over_video_preview = is_cursor_over_video_preview();

                    if over_image_preview || over_video_preview {
                        let guard_active = video_hover_guard_until
                            .map(|until| Instant::now() < until)
                            .unwrap_or(false);

                        // For video, keep the short spawn grace to prevent instant close
                        // right after ffplay appears under the cursor.
                        if !over_video_preview || !guard_active {
                            if last_file.is_some() {
                                hide_preview();
                                last_file = None;
                            }
                            video_hover_guard_until = None;
                        }

                        hover_start = Some(Instant::now());
                        continue;
                    }
                }

                // While moving (including list scrolling), avoid heavy accessibility
                // resolution and wait until hover is stable before probing media.
                if last_file.is_some() {
                    hide_preview();
                    last_file = None;
                    video_hover_guard_until = None;
                }
                hover_start = Some(Instant::now());
                continue;
            }

            // Mouse is stationary - check for keyboard navigation
            // Only when Explorer is the foreground window (keyboard input goes there)
            if is_foreground_explorer()
                && last_keyboard_focus_probe.elapsed()
                    >= Duration::from_millis(KEYBOARD_FOCUS_PROBE_MS)
            {
                last_keyboard_focus_probe = Instant::now();
                if let Some(focused_info) = uia.as_ref().and_then(|a| get_focused_explorer_item(a))
                {
                    let focused_name = match &focused_info.result {
                        AccessibilityResult::FileName(name) => name.clone(),
                        AccessibilityResult::FullPath(path) => path.to_string_lossy().to_string(),
                    };

                    if last_focused_name.is_none() {
                        if allow_keyboard_preview_on_first_observation {
                            // User explicitly used keyboard right after folder change.
                            // Allow first observed focused item to trigger preview.
                            last_focused_name = Some(focused_name.clone());
                            allow_keyboard_preview_on_first_observation = false;

                            // Dismiss any active mouse hover
                            if last_file.is_some() && !is_keyboard_hover {
                                hide_preview();
                                last_file = None;
                                hover_start = None;
                            }

                            // Resolve to a media file and show keyboard preview
                            if let Some(path) = resolve_focused_item_to_path(&focused_info) {
                                if keyboard_file.as_ref() != Some(&path) {
                                    // Hide previous preview before showing new one
                                    if is_keyboard_hover {
                                        hide_preview();
                                    }
                                    keyboard_file = Some(path.clone());
                                    is_keyboard_hover = true;
                                    video_hover_guard_until = if is_video_file(&path) {
                                        Some(
                                            Instant::now()
                                                + Duration::from_millis(
                                                    VIDEO_HOVER_DISMISS_GRACE_MS,
                                                ),
                                        )
                                    } else {
                                        None
                                    };
                                    show_preview_keyboard(
                                        &path,
                                        focused_info.rect.left,
                                        focused_info.rect.top,
                                        focused_info.rect.right,
                                        focused_info.rect.bottom,
                                    );
                                }
                            } else {
                                // Not a media file - hide any keyboard preview
                                if is_keyboard_hover {
                                    hide_preview();
                                }
                                keyboard_file = None;
                                is_keyboard_hover = false;
                                video_hover_guard_until = None;
                            }
                            continue;
                        }

                        // First observation after mouse stopped - just record, don't trigger
                        last_focused_name = Some(focused_name);
                    } else if last_focused_name.as_ref() != Some(&focused_name) {
                        // Focused item changed - keyboard navigation detected
                        last_focused_name = Some(focused_name);
                        allow_keyboard_preview_on_first_observation = false;

                        // Dismiss any active mouse hover
                        if last_file.is_some() && !is_keyboard_hover {
                            hide_preview();
                            last_file = None;
                            hover_start = None;
                        }

                        // Resolve to a media file and show keyboard preview
                        if let Some(path) = resolve_focused_item_to_path(&focused_info) {
                            if keyboard_file.as_ref() != Some(&path) {
                                // Hide previous preview before showing new one
                                if is_keyboard_hover {
                                    hide_preview();
                                }
                                keyboard_file = Some(path.clone());
                                is_keyboard_hover = true;
                                video_hover_guard_until = if is_video_file(&path) {
                                    Some(
                                        Instant::now()
                                            + Duration::from_millis(VIDEO_HOVER_DISMISS_GRACE_MS),
                                    )
                                } else {
                                    None
                                };
                                show_preview_keyboard(
                                    &path,
                                    focused_info.rect.left,
                                    focused_info.rect.top,
                                    focused_info.rect.right,
                                    focused_info.rect.bottom,
                                );
                            }
                        } else {
                            // Not a media file - hide any keyboard preview
                            if is_keyboard_hover {
                                hide_preview();
                            }
                            keyboard_file = None;
                            is_keyboard_hover = false;
                            video_hover_guard_until = None;
                        }
                        continue;
                    }
                }
            }

            // If keyboard hover is active, skip mouse hover delay logic
            if is_keyboard_hover {
                continue;
            }

            // Check if we've hovered long enough (mouse hover)
            if let Some(start) = hover_start {
                if start.elapsed() >= hover_delay {
                    if last_hover_probe.elapsed() < Duration::from_millis(HOVER_PROBE_MS) {
                        continue;
                    }
                    last_hover_probe = Instant::now();

                    // Try to get file under cursor
                    if let Some(file_path) = get_file_under_cursor() {
                        if last_file.as_ref() != Some(&file_path) {
                            last_file = Some(file_path.clone());
                            video_hover_guard_until = if is_video_file(&file_path) {
                                Some(
                                    Instant::now()
                                        + Duration::from_millis(VIDEO_HOVER_DISMISS_GRACE_MS),
                                )
                            } else {
                                None
                            };
                            show_preview(&file_path, cursor_pos.x, cursor_pos.y);
                        }
                    } else {
                        // No file found while mouse is stationary.
                        // Keep current preview state; dismissal should happen only on mouse move.
                    }
                }
            } else {
                // Initialize hover_start if not moving
                hover_start = Some(Instant::now());
            }
        }
    }

    unsafe {
        CoUninitialize();
    }
}

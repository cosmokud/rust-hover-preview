use crate::pdf_preview::is_pdf_file;
use crate::preview_window::{
    cursor_preview_hover, hide_preview, kill_stray_video_process, preview_screen_rect,
    show_preview, show_preview_keyboard, text_scroll_pointer_hold, PreviewCursorHover,
};
use crate::text_formats::is_text_file;
use crate::video_formats::is_video_file;
use crate::wheel_input;
use crate::{CONFIG, RUNNING};
use once_cell::sync::Lazy;
use std::collections::{HashMap, HashSet};
use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use std::path::{Path, PathBuf};
use std::sync::{atomic::Ordering, Arc, Mutex};
use std::time::{Duration, Instant};
use windows::core::{Interface, VARIANT};
use windows::Win32::Foundation::{BOOL, HWND, LPARAM, POINT, RECT};
use windows::Win32::Graphics::Gdi::ScreenToClient;
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoTaskMemFree, CoUninitialize, IServiceProvider, CLSCTX_ALL,
    COINIT_APARTMENTTHREADED,
};
use windows::Win32::System::Variant::VariantClear;
use windows::Win32::UI::Accessibility::{
    CUIAutomation, IUIAutomation, IUIAutomationElement, IUIAutomationLegacyIAccessiblePattern,
    UIA_DataItemControlTypeId, UIA_LegacyIAccessiblePatternId, UIA_ListItemControlTypeId,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetAsyncKeyState, VK_DOWN, VK_END, VK_HOME, VK_LBUTTON, VK_LEFT, VK_MBUTTON, VK_NEXT, VK_PRIOR,
    VK_RBUTTON, VK_RETURN, VK_RIGHT, VK_UP, VK_XBUTTON1, VK_XBUTTON2,
};
use windows::Win32::UI::Shell::{
    IFolderView, INameSpaceTreeControl, IPersistFolder2, IShellBrowser, IShellFolder,
    IShellFolderViewDual, IShellItem, IShellView, IShellWindows, SHCreateItemFromIDList,
    SHCreateItemWithParent, SID_STopLevelBrowser, ShellWindows, SIGDN_DESKTOPABSOLUTEPARSING,
};
use windows::Win32::UI::WindowsAndMessaging::{
    EnumWindows, GetClassNameW, GetCursorPos, GetForegroundWindow, GetSystemMetrics,
    GetWindowPlacement, GetWindowRect, GetWindowThreadProcessId, IsIconic, IsWindowVisible,
    WindowFromPoint, SM_CXVIRTUALSCREEN, SM_CYVIRTUALSCREEN, SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN,
    SW_SHOWMAXIMIZED, WINDOWPLACEMENT,
};

// Supported image extensions
const IMAGE_EXTENSIONS: &[&str] = &[
    "jpg", "jpeg", "jpe", "jfif", "png", "apng", "gif", "bmp", "ico", "tiff", "tif", "webp", "tga",
    "pbm", "pgm", "ppm", "pam", "pnm", "hdr", "exr", "qoi", "ff",
];

struct FolderMediaIndex {
    built_at: Instant,
    by_file_name: HashMap<String, PathBuf>,
    by_stem: HashMap<String, PathBuf>,
}

struct ExplorerFoldersCache {
    built_at: Instant,
    folders: Arc<Vec<(isize, String)>>,
}

struct ShellViewMediaIndex {
    built_at: Instant,
    by_display_name: HashMap<String, PathBuf>,
    by_file_name: HashMap<String, PathBuf>,
    by_stem: HashMap<String, PathBuf>,
    root_folder: Option<String>,
}

struct SearchRootMediaIndex {
    built_at: Instant,
    by_file_name: HashMap<String, Vec<PathBuf>>,
    by_stem: HashMap<String, Vec<PathBuf>>,
}

struct ExplorerWindowCounts {
    total: usize,
    visible: usize,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct DisplaySignature {
    left: i32,
    top: i32,
    width: i32,
    height: i32,
}

struct ActiveShellViewContext {
    shell_view_hwnd: isize,
    location_url: Option<String>,
    folder_path: Option<String>,
    shell_view: IShellView,
    client_point: POINT,
}

#[derive(Clone, Default)]
struct HoverResolverHints {
    current_folder: Option<String>,
    location_url: Option<String>,
    is_search_view: bool,
    search_root: Option<String>,
    shell_view_hwnd: Option<isize>,
}

/// "Do not preview this file" latch, shared by the mouse hover path.
#[derive(Default)]
struct SuppressedHover {
    file: Option<PathBuf>,
    started_at: Option<Instant>,
    /// Set when a keyboard preview was dismissed by mouse movement: the latch
    /// never expires for that file and is released only when the cursor
    /// resolves a different file, or when a keyboard preview takes over.
    sticky: bool,
}

impl SuppressedHover {
    fn clear(&mut self) {
        self.file = None;
        self.started_at = None;
        self.sticky = false;
    }

    fn suppress(&mut self, file: PathBuf, sticky: bool) {
        self.file = Some(file);
        self.started_at = Some(Instant::now());
        self.sticky = sticky;
    }

    fn matches(&self, path: &PathBuf) -> bool {
        self.file
            .as_ref()
            .map(|file| same_path(file, path))
            .unwrap_or(false)
    }

    /// Regular latches keep the same-file rehover delay; a sticky one never lets
    /// the suppressed file back onto the mouse path.
    fn rehover_allowed(&self, required_delay_ms: u64) -> bool {
        if self.sticky {
            return false;
        }

        self.started_at
            .map(|started| started.elapsed() >= Duration::from_millis(required_delay_ms))
            .unwrap_or(true)
    }
}

/// Pointer freeze for keyboard previews. A keyboard preview is placed next to
/// the focused item, which can put it right over the parked cursor; the pointer
/// must not take over in that case. The freeze is decided from the preview's own
/// box at spawn time and released by a real mouse move, so cursor jitter cannot
/// end a keyboard preview.
#[derive(Default)]
struct KeyboardPointerPause {
    armed: bool,
    /// Set on every keyboard preview spawn while the preview thread has not
    /// published a box yet.
    box_watch_until: Option<Instant>,
    /// Last observed box, so a box that is still being replaced is not trusted.
    last_box: Option<(i32, i32, i32, i32)>,
}

impl KeyboardPointerPause {
    /// Called on every keyboard preview spawn: the box arrives once the preview
    /// is actually on screen, and the freeze is decided from it.
    fn watch_for_box(&mut self) {
        self.armed = false;
        self.last_box = None;
        self.box_watch_until =
            Some(Instant::now() + Duration::from_millis(KEYBOARD_PREVIEW_BOX_WATCH_MS));
    }

    fn clear(&mut self) {
        self.armed = false;
        self.box_watch_until = None;
        self.last_box = None;
    }

    fn is_watching(&self) -> bool {
        self.box_watch_until.is_some()
    }

    /// True while the pointer must not drive previews, probe them, or dismiss
    /// them: the preview box is either known to cover the cursor, or still being
    /// waited on.
    fn freezes_pointer(&self) -> bool {
        self.armed || self.box_watch_until.is_some()
    }

    /// Cursor movement that hands control back to the mouse. Small movements are
    /// ignored on purpose so a parked mouse cannot cancel a keyboard preview.
    fn move_threshold_px(&self) -> i32 {
        if self.freezes_pointer() {
            KEYBOARD_POINTER_MOVE_TOLERANCE_PX
        } else {
            MOUSE_MOVE_PX
        }
    }

    /// Decide the freeze from the preview's on-screen box. A box only counts
    /// once it survives a second observation, because a preview that is being
    /// replaced can still report the outgoing window. `None` keeps waiting until
    /// the watch expires, so a preview that never appears cannot freeze the
    /// pointer forever.
    fn evaluate_box(&mut self, cursor: POINT, preview_box: Option<(i32, i32, i32, i32)>) {
        let Some(watch_until) = self.box_watch_until else {
            return;
        };

        let Some(box_rect) = preview_box else {
            self.last_box = None;
            if Instant::now() >= watch_until {
                self.box_watch_until = None;
            }
            return;
        };

        if self.last_box != Some(box_rect) {
            self.last_box = Some(box_rect);
            return;
        }

        let (left, top, right, bottom) = box_rect;
        self.box_watch_until = None;
        self.last_box = None;
        self.armed = cursor.x >= left && cursor.x < right && cursor.y >= top && cursor.y < bottom;
    }
}

/// Wheel-scroll settle probe. A scroll moves the list under a parked pointer,
/// and Explorer can still be animating the scroll when the loop looks, so a
/// scroll-driven probe only acts on an item that survives a second observation —
/// the same rule the keyboard preview box uses. Without a scroll in flight the
/// probe is pass-through and the single-probe-per-parked-cursor behavior stands.
#[derive(Default)]
struct ScrollSettleProbe {
    pending: bool,
    last: Option<PathBuf>,
}

impl ScrollSettleProbe {
    /// A wheel tick arrived: the next stationary probes belong to this gesture.
    fn arm(&mut self) {
        self.pending = true;
        self.last = None;
    }

    /// A real mouse move supersedes the gesture.
    fn disarm(&mut self) {
        self.pending = false;
        self.last = None;
    }

    fn is_pending(&self) -> bool {
        self.pending
    }

    /// Records the item resolved under the cursor and reports whether the probe
    /// may act on it. A miss settles immediately: it can only dismiss, and the
    /// mouse-move path dismisses on a miss just the same.
    fn observe(&mut self, resolved: Option<&PathBuf>) -> bool {
        if !self.pending {
            return true;
        }

        let settled = match (self.last.as_ref(), resolved) {
            (Some(previous), Some(current)) => same_path(previous, current),
            (None, None) => true,
            _ => false,
        };

        self.last = resolved.cloned();
        if settled {
            self.pending = false;
        }
        settled
    }
}

const FOLDER_INDEX_TTL_MS: u64 = 60000;
const EXPLORER_FOLDERS_CACHE_TTL_MS: u64 = 250;
const SHELL_VIEW_INDEX_TTL_MS: u64 = 5000;
const SHELL_VIEW_INDEX_MAX_ITEMS: i32 = 50000;
const SHELL_VIEW_INDEX_SYNC_ITEM_LIMIT: i32 = 1000;
const SEARCH_ROOT_INDEX_TTL_MS: u64 = 60000;
const SEARCH_ROOT_INDEX_MAX_DIRS: usize = 20000;
const SEARCH_ROOT_INDEX_MAX_FILES: usize = 50000;
const EXPLORER_PROBE_SLOW_MS: u64 = 700;
const EXPLORER_WINDOW_CACHE_TTL_MS: u64 = 1000;
const FOLDER_INDEX_CACHE_MAX_ENTRIES: usize = 16;
const EXPLORER_REAL_FOLDER_CACHE_MAX_ENTRIES: usize = 256;
const SEARCH_ROOT_CACHE_MAX_ENTRIES: usize = 8;
const FOLDER_PROBE_MS: u64 = 200;
const IDLE_FOLDER_PROBE_MS: u64 = 750;
const FOLDER_PROBE_TRIGGER_MS: u64 = 400;
const DISPLAY_CHANGE_BACKOFF_MS: u64 = 1500;
const KEYBOARD_FOCUS_INPUT_GRACE_MS: u64 = 500;
const HOVER_RESOLVER_INPUT_GRACE_MS: u64 = 1500;
const WHEEL_SCROLL_SETTLE_MS: u64 = 150;
const MOUSE_MOVE_PX: i32 = 5;
const KEYBOARD_POINTER_MOVE_TOLERANCE_PX: i32 = 20;
const KEYBOARD_PREVIEW_BOX_WATCH_MS: u64 = 2500;
const VK_BACK_CODE: i32 = 0x08;
const VK_CONTROL_CODE: i32 = 0x11;
const VK_MENU_CODE: i32 = 0x12;
const VK_T_CODE: i32 = 0x54;

static FOLDER_MEDIA_INDEX: Lazy<Mutex<HashMap<String, FolderMediaIndex>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));
static FOLDER_INDEX_BUILDING: Lazy<Mutex<HashSet<String>>> =
    Lazy::new(|| Mutex::new(HashSet::new()));
static EXPLORER_FOLDERS_CACHE: Lazy<Mutex<Option<ExplorerFoldersCache>>> =
    Lazy::new(|| Mutex::new(None));
static SHELL_VIEW_MEDIA_INDEX: Lazy<Mutex<HashMap<isize, ShellViewMediaIndex>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));
static LEGACY_SEARCH_SHELL_VIEW_MEDIA_INDEX: Lazy<Mutex<HashMap<isize, ShellViewMediaIndex>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));
static SEARCH_ROOT_MEDIA_INDEX: Lazy<Mutex<HashMap<String, SearchRootMediaIndex>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));
static SEARCH_ROOT_INDEX_BUILDING: Lazy<Mutex<HashSet<String>>> =
    Lazy::new(|| Mutex::new(HashSet::new()));
static EXPLORER_LAST_REAL_FOLDERS: Lazy<Mutex<HashMap<isize, String>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));
static EXPLORER_WINDOW_CACHE: Lazy<Mutex<HashMap<isize, (bool, Instant)>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

fn clear_shell_view_probe_caches() {
    if let Ok(mut cache) = SHELL_VIEW_MEDIA_INDEX.lock() {
        cache.clear();
    }
    if let Ok(mut cache) = LEGACY_SEARCH_SHELL_VIEW_MEDIA_INDEX.lock() {
        cache.clear();
    }
    if let Ok(mut cache) = EXPLORER_FOLDERS_CACHE.lock() {
        *cache = None;
    }
    if let Ok(mut cache) = EXPLORER_WINDOW_CACHE.lock() {
        cache.clear();
    }
    if let Ok(mut cache) = EXPLORER_LAST_REAL_FOLDERS.lock() {
        cache.clear();
    }
}

fn current_display_signature() -> Option<DisplaySignature> {
    unsafe {
        let width = GetSystemMetrics(SM_CXVIRTUALSCREEN);
        let height = GetSystemMetrics(SM_CYVIRTUALSCREEN);
        if width <= 0 || height <= 0 {
            return None;
        }

        Some(DisplaySignature {
            left: GetSystemMetrics(SM_XVIRTUALSCREEN),
            top: GetSystemMetrics(SM_YVIRTUALSCREEN),
            width,
            height,
        })
    }
}

fn display_signature_changed(
    previous: Option<DisplaySignature>,
    current: DisplaySignature,
) -> bool {
    previous
        .map(|signature| signature != current)
        .unwrap_or(false)
}

fn recent_elapsed_within(elapsed: Option<Duration>, limit_ms: u64) -> bool {
    elapsed
        .map(|elapsed| elapsed <= Duration::from_millis(limit_ms))
        .unwrap_or(false)
}

fn should_probe_keyboard_focus(recent_navigation_elapsed: Option<Duration>) -> bool {
    recent_elapsed_within(recent_navigation_elapsed, KEYBOARD_FOCUS_INPUT_GRACE_MS)
}

fn should_probe_hover_resolver(
    preview_active: bool,
    cursor_moved: bool,
    recent_input_elapsed: Option<Duration>,
) -> bool {
    preview_active
        || cursor_moved
        || recent_elapsed_within(recent_input_elapsed, HOVER_RESOLVER_INPUT_GRACE_MS)
}

fn should_probe_stationary_hover(already_probed: bool) -> bool {
    !already_probed
}

/// The pointer probe only matters while a mouse preview can be under the
/// pointer. Keyboard previews and a frozen pointer never trigger it.
fn should_probe_preview_hover(
    pointer_frozen: bool,
    mouse_preview_active: bool,
    suppress_until_cursor_leaves: bool,
) -> bool {
    !pointer_frozen && (mouse_preview_active || suppress_until_cursor_leaves)
}

fn is_jpeg_extension(ext: &str) -> bool {
    matches!(ext, "jpg" | "jpeg" | "jpe" | "jfif")
}

fn clear_variant(variant: &mut VARIANT) {
    unsafe {
        let _ = VariantClear(variant as *mut VARIANT);
    }
}

fn build_folder_media_index(folder_path: &PathBuf, _folder_key: &str) -> Option<FolderMediaIndex> {
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
        built_at: Instant::now(),
        by_file_name,
        by_stem,
    })
}

fn trim_folder_index_cache(cache: &mut HashMap<String, FolderMediaIndex>) {
    if cache.len() <= FOLDER_INDEX_CACHE_MAX_ENTRIES {
        return;
    }

    let mut oldest_key: Option<String> = None;
    let mut oldest_age = Duration::ZERO;
    for (folder, index) in cache.iter() {
        let age = index.built_at.elapsed();
        if oldest_key.is_none() || age > oldest_age {
            oldest_key = Some(folder.clone());
            oldest_age = age;
        }
    }

    if let Some(oldest_key) = oldest_key {
        cache.remove(&oldest_key);
    }
}

fn queue_folder_index_build(folder_path: PathBuf, folder_key: String) {
    let should_build = {
        let cache_is_fresh = FOLDER_MEDIA_INDEX
            .lock()
            .ok()
            .and_then(|cache| cache.get(&folder_key).map(|index| index.built_at.elapsed()))
            .map(|age| age <= Duration::from_millis(FOLDER_INDEX_TTL_MS))
            .unwrap_or(false);
        if cache_is_fresh {
            false
        } else if let Ok(mut building) = FOLDER_INDEX_BUILDING.lock() {
            if building.contains(&folder_key) {
                false
            } else {
                building.insert(folder_key.clone());
                true
            }
        } else {
            false
        }
    };

    if !should_build {
        return;
    }

    std::thread::spawn(move || {
        let built_index = build_folder_media_index(&folder_path, &folder_key);
        if let Some(index) = built_index {
            if let Ok(mut cache) = FOLDER_MEDIA_INDEX.lock() {
                cache.insert(folder_key.clone(), index);
                trim_folder_index_cache(&mut cache);
            }
        }
        if let Ok(mut building) = FOLDER_INDEX_BUILDING.lock() {
            building.remove(&folder_key);
        }
    });
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

    if let Ok(mut cache) = FOLDER_MEDIA_INDEX.lock() {
        cache.retain(|_, index| {
            index.built_at.elapsed() <= Duration::from_millis(FOLDER_INDEX_TTL_MS)
        });

        if let Some(index) = cache.get(folder_key) {
            if let Some(path) = index.by_file_name.get(&item_name_lower) {
                return Some(path.clone());
            }

            if let Some(stem_key) = item_stem_lower.as_ref() {
                if let Some(path) = index.by_stem.get(stem_key) {
                    if let Some(item_ext) = item_ext_lower.as_deref() {
                        if let Some(candidate_ext) = path.extension().and_then(|s| s.to_str()) {
                            let candidate_ext_lower = candidate_ext.to_ascii_lowercase();
                            if candidate_ext_lower == item_ext
                                || (is_jpeg_extension(&candidate_ext_lower)
                                    && is_jpeg_extension(item_ext))
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
        }
    }

    // Never block hover polling on a huge folder scan. Queue async build instead.
    queue_folder_index_build(folder_path.clone(), folder_key.to_string());
    None
}

fn is_image_file(path: &PathBuf) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| IMAGE_EXTENSIONS.contains(&ext.to_lowercase().as_str()))
        .unwrap_or(false)
}

fn is_media_file(path: &PathBuf) -> bool {
    is_image_file(path) || is_video_file(path) || is_pdf_file(path) || is_text_file(path)
}

fn same_path(a: &PathBuf, b: &PathBuf) -> bool {
    a == b
        || a.to_string_lossy()
            .eq_ignore_ascii_case(&b.to_string_lossy())
}

fn urlencoding_decode(s: &str) -> String {
    let mut bytes = Vec::with_capacity(s.len());
    let mut chars = s.as_bytes().iter().copied().peekable();

    while let Some(byte) = chars.next() {
        if byte == b'%' {
            let hi = chars.next();
            let lo = chars.next();
            if let (Some(hi), Some(lo)) = (hi, lo) {
                let hex = [hi, lo];
                if let Ok(hex_str) = std::str::from_utf8(&hex) {
                    if let Ok(decoded) = u8::from_str_radix(hex_str, 16) {
                        bytes.push(decoded);
                        continue;
                    }
                }
                bytes.push(b'%');
                bytes.push(hi);
                bytes.push(lo);
            } else {
                bytes.push(b'%');
                if let Some(hi) = hi {
                    bytes.push(hi);
                }
            }
        } else if byte == b'+' {
            bytes.push(b' ');
        } else {
            bytes.push(byte);
        }
    }

    String::from_utf8_lossy(&bytes).into_owned()
}

fn urlencoding_decode_repeated(s: &str) -> String {
    let mut current = s.to_string();
    for _ in 0..3 {
        let decoded = urlencoding_decode(&current);
        if decoded == current {
            break;
        }
        current = decoded;
    }
    current
}

fn is_search_ms_url(url_str: &str) -> bool {
    url_str
        .trim_start()
        .to_ascii_lowercase()
        .starts_with("search-ms:")
}

fn normalize_file_url_path(url_str: &str) -> Option<String> {
    let path = if let Some(path) = url_str.strip_prefix("file:///") {
        path.replace('/', "\\")
    } else if let Some(path) = url_str.strip_prefix("file://") {
        format!("\\\\{}", path.replace('/', "\\"))
    } else {
        return None;
    };

    Some(urlencoding_decode(&path))
}

fn normalize_search_location(location: &str) -> Option<String> {
    let decoded = urlencoding_decode_repeated(location);
    let location = decoded.trim();
    if location.is_empty() {
        return None;
    }

    let path = normalize_file_url_path(location).unwrap_or_else(|| location.replace('/', "\\"));
    let path = path.trim().trim_matches('"').to_string();
    if path.is_empty() {
        None
    } else {
        Some(path)
    }
}

fn search_ms_location_from_url(url_str: &str) -> Option<String> {
    if !is_search_ms_url(url_str) {
        return None;
    }

    let decoded_url = urlencoding_decode_repeated(url_str);
    for part in decoded_url.split('&') {
        let decoded_part = urlencoding_decode_repeated(part.trim());
        let part_lower = decoded_part.to_ascii_lowercase();

        for prefix in ["crumb=location:", "crumb=folder:"] {
            if let Some(index) = part_lower.find(prefix) {
                let location_start = index + prefix.len();
                return normalize_search_location(&decoded_part[location_start..]);
            }
        }
    }

    None
}

fn is_usable_folder_path(path: &str) -> bool {
    let path = path.trim();
    !path.is_empty() && PathBuf::from(path).is_dir()
}

fn resolve_media_path_candidate(text: &str) -> Option<PathBuf> {
    let candidate = text.trim().trim_matches(|c| c == '"' || c == '\'');
    if candidate.is_empty() {
        return None;
    }

    let normalized = normalize_file_url_path(candidate).unwrap_or_else(|| candidate.to_string());
    if !is_valid_file_path(&normalized) {
        return None;
    }

    let path = PathBuf::from(normalized);
    if path.exists() && is_media_file(&path) {
        Some(path)
    } else {
        None
    }
}

fn resolve_media_path_from_text(text: &str) -> Option<PathBuf> {
    let text = text.trim();
    if text.is_empty() {
        return None;
    }

    if let Some(path) = resolve_media_path_candidate(text) {
        return Some(path);
    }

    let chars: Vec<(usize, char)> = text.char_indices().collect();
    for window in chars.windows(3) {
        let [(start, drive), (_, colon), (_, slash)] = window else {
            continue;
        };
        if !drive.is_ascii_alphabetic() || *colon != ':' || (*slash != '\\' && *slash != '/') {
            continue;
        }

        let candidate = &text[*start..];
        if let Some(path) = resolve_media_path_candidate(candidate) {
            return Some(path);
        }

        let mut ends: Vec<usize> = candidate.char_indices().map(|(idx, _)| idx).collect();
        ends.push(candidate.len());
        for end in ends.into_iter().rev() {
            if end <= 3 {
                continue;
            }
            if let Some(path) = resolve_media_path_candidate(&candidate[..end]) {
                return Some(path);
            }
        }
    }

    None
}

fn cache_explorer_real_folder(hwnd: isize, folder: &str) {
    if !is_usable_folder_path(folder) {
        return;
    }

    if let Ok(mut cache) = EXPLORER_LAST_REAL_FOLDERS.lock() {
        if !cache.contains_key(&hwnd) && cache.len() >= EXPLORER_REAL_FOLDER_CACHE_MAX_ENTRIES {
            if let Some(old_key) = cache.keys().next().copied() {
                cache.remove(&old_key);
            }
        }
        cache.insert(hwnd, folder.to_string());
    }
}

fn get_cached_explorer_real_folder(hwnd: isize) -> Option<String> {
    EXPLORER_LAST_REAL_FOLDERS
        .lock()
        .ok()
        .and_then(|cache| cache.get(&hwnd).cloned())
        .filter(|folder| is_usable_folder_path(folder))
}

fn resolve_explorer_location_folder(hwnd: isize, url_str: &str) -> Option<String> {
    if let Some(path) = normalize_file_url_path(url_str) {
        if is_usable_folder_path(&path) {
            cache_explorer_real_folder(hwnd, &path);
            return Some(path);
        }
    }

    if is_search_ms_url(url_str) {
        if let Some(path) = search_ms_location_from_url(url_str) {
            if is_usable_folder_path(&path) {
                cache_explorer_real_folder(hwnd, &path);
                return Some(path);
            }
        }

        // Win11 search can stop exposing a parseable root. Use the last normal
        // folder seen for this exact Explorer window, which matches the second
        // same-folder-window workaround without needing a real second window.
        return get_cached_explorer_real_folder(hwnd);
    }

    None
}

fn merge_common_folder_root(current: Option<PathBuf>, path: &Path) -> Option<PathBuf> {
    let candidate_dir = path.parent()?.to_path_buf();

    match current {
        Some(existing) => existing
            .ancestors()
            .find(|ancestor| candidate_dir.starts_with(ancestor))
            .map(|ancestor| ancestor.to_path_buf()),
        None => Some(candidate_dir),
    }
}

fn get_shell_view_search_root(view_hwnd_key: isize) -> Option<String> {
    let mut cache = SHELL_VIEW_MEDIA_INDEX.lock().ok()?;
    cache.retain(|_, index| {
        index.built_at.elapsed() <= Duration::from_millis(SHELL_VIEW_INDEX_TTL_MS)
    });

    if !cache.contains_key(&view_hwnd_key) {
        let index = build_shell_view_media_index(view_hwnd_key)?;
        cache.insert(view_hwnd_key, index);
    }

    cache
        .get(&view_hwnd_key)
        .and_then(|index| index.root_folder.clone())
}

fn resolve_search_root_from_context(context: &ActiveShellViewContext) -> Option<String> {
    if let Some(url) = context.location_url.as_deref() {
        if let Some(root) = resolve_explorer_location_folder(context.shell_view_hwnd, url) {
            return Some(root);
        }
    }

    if let Some(folder) = context
        .folder_path
        .as_deref()
        .filter(|folder| is_usable_folder_path(folder))
    {
        cache_explorer_real_folder(context.shell_view_hwnd, folder);
        return Some(folder.to_string());
    }

    if let Some(root) = get_shell_view_search_root(context.shell_view_hwnd) {
        cache_explorer_real_folder(context.shell_view_hwnd, &root);
        return Some(root);
    }

    get_cached_explorer_real_folder(context.shell_view_hwnd)
}

fn is_probable_search_view_context(context: &ActiveShellViewContext) -> bool {
    context
        .location_url
        .as_deref()
        .map(is_search_ms_url)
        .unwrap_or(false)
        || (context.folder_path.is_none()
            && get_shell_view_search_root(context.shell_view_hwnd).is_some())
}

/// Explorer windows and their folders. The cached snapshot is shared through an
/// `Arc` so a poll tick does not clone the list and every folder string.
fn get_all_explorer_folders() -> Arc<Vec<(isize, String)>> {
    if let Ok(cache) = EXPLORER_FOLDERS_CACHE.lock() {
        if let Some(cache_entry) = cache.as_ref() {
            if cache_entry.built_at.elapsed()
                <= Duration::from_millis(EXPLORER_FOLDERS_CACHE_TTL_MS)
            {
                return Arc::clone(&cache_entry.folders);
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
                                let folder = browser
                                    .LocationURL()
                                    .ok()
                                    .and_then(|url| {
                                        let url_str = url.to_string();
                                        resolve_explorer_location_folder(hwnd, &url_str)
                                    })
                                    .or_else(|| get_cached_explorer_real_folder(hwnd));

                                if let Some(folder) = folder {
                                    result.push((hwnd, folder));
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    let result = Arc::new(result);

    if let Ok(mut cache) = EXPLORER_FOLDERS_CACHE.lock() {
        *cache = Some(ExplorerFoldersCache {
            built_at: Instant::now(),
            folders: Arc::clone(&result),
        });
    }

    result
}

fn get_explorer_hwnd_under_cursor_or_foreground() -> Option<HWND> {
    unsafe {
        let folders = get_all_explorer_folders();

        let mut cursor_pos = POINT::default();
        if GetCursorPos(&mut cursor_pos).is_ok() {
            let hwnd = WindowFromPoint(cursor_pos);
            if !hwnd.0.is_null() {
                let mut current_hwnd = hwnd;
                let mut top_hwnd = hwnd;

                for _ in 0..20 {
                    for (explorer_hwnd, _) in folders.iter() {
                        let explorer_hwnd = HWND(*explorer_hwnd as *mut _);
                        if current_hwnd == explorer_hwnd {
                            return Some(explorer_hwnd);
                        }
                    }

                    match windows::Win32::UI::WindowsAndMessaging::GetParent(current_hwnd) {
                        Ok(parent) if !parent.0.is_null() && parent != current_hwnd => {
                            top_hwnd = parent;
                            current_hwnd = parent;
                        }
                        _ => break,
                    }
                }

                // Only fall back to the top-level parent. Child controls also run
                // in explorer.exe, so process-name checks on the original child
                // can return a handle ShellWindows cannot resolve.
                if is_explorer_window(top_hwnd) {
                    return Some(top_hwnd);
                }
            }
        }

        let foreground = GetForegroundWindow();
        if !foreground.is_invalid() {
            for (explorer_hwnd, _) in folders.iter() {
                let explorer_hwnd = HWND(*explorer_hwnd as *mut _);
                if foreground == explorer_hwnd {
                    return Some(explorer_hwnd);
                }
            }

            if is_explorer_window(foreground) {
                return Some(foreground);
            }
        }
    }

    None
}

fn get_explorer_hwnd_under_cursor_or_foreground_legacy() -> Option<HWND> {
    unsafe {
        let folders = get_all_explorer_folders();

        let mut cursor_pos = POINT::default();
        if GetCursorPos(&mut cursor_pos).is_ok() {
            let hwnd = WindowFromPoint(cursor_pos);
            if !hwnd.0.is_null() {
                let mut current_hwnd = hwnd;
                let mut top_hwnd = hwnd;

                for _ in 0..20 {
                    for (explorer_hwnd, _) in folders.iter() {
                        let explorer_hwnd = HWND(*explorer_hwnd as *mut _);
                        if current_hwnd == explorer_hwnd {
                            return Some(explorer_hwnd);
                        }
                    }

                    match windows::Win32::UI::WindowsAndMessaging::GetParent(current_hwnd) {
                        Ok(parent) if !parent.0.is_null() && parent != current_hwnd => {
                            top_hwnd = parent;
                            current_hwnd = parent;
                        }
                        _ => break,
                    }
                }

                if is_explorer_window(top_hwnd) {
                    return Some(top_hwnd);
                }
            }
        }

        let foreground = GetForegroundWindow();
        if !foreground.is_invalid() {
            for (explorer_hwnd, _) in folders.iter() {
                let explorer_hwnd = HWND(*explorer_hwnd as *mut _);
                if foreground == explorer_hwnd {
                    return Some(explorer_hwnd);
                }
            }

            if is_explorer_window(foreground) {
                return Some(foreground);
            }
        }
    }

    None
}

fn get_explorer_location_url(hwnd: HWND) -> Option<String> {
    let hwnd_key = hwnd.0 as isize;

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
            if browser_hwnd.0 != hwnd_key {
                continue;
            }

            return browser.LocationURL().ok().map(|url| url.to_string());
        }
    }

    None
}

fn get_current_explorer_location_url_legacy() -> Option<String> {
    if let Some(context) = get_active_shell_view_context_at_cursor() {
        if let Some(url) = context.location_url {
            return Some(url);
        }
    }

    let hwnd = get_explorer_hwnd_under_cursor_or_foreground_legacy()?;
    get_explorer_location_url(hwnd)
}

fn get_active_shell_view_context(screen_point: &POINT) -> Option<ActiveShellViewContext> {
    unsafe {
        let shell_windows =
            CoCreateInstance::<_, IShellWindows>(&ShellWindows, None, CLSCTX_ALL).ok()?;
        let count = shell_windows.Count().ok()?;
        let cursor_hwnd = WindowFromPoint(*screen_point);
        let mut rect_candidate: Option<ActiveShellViewContext> = None;
        let mut foreground_candidate: Option<ActiveShellViewContext> = None;
        let foreground = GetForegroundWindow();

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
            let service_provider = match browser.cast::<IServiceProvider>() {
                Ok(service_provider) => service_provider,
                Err(_) => continue,
            };
            let shell_browser: IShellBrowser =
                match service_provider.QueryService(&SID_STopLevelBrowser) {
                    Ok(shell_browser) => shell_browser,
                    Err(_) => continue,
                };
            let shell_view = match shell_browser.QueryActiveShellView() {
                Ok(shell_view) => shell_view,
                Err(_) => continue,
            };
            let shell_view_hwnd = match shell_view.GetWindow() {
                Ok(hwnd) if !hwnd.is_invalid() => hwnd,
                _ => continue,
            };
            if !IsWindowVisible(shell_view_hwnd).as_bool() || is_window_minimized(shell_view_hwnd) {
                continue;
            }
            let shell_view_hwnd_key = shell_view_hwnd.0 as isize;

            let mut rect = RECT::default();
            if GetWindowRect(shell_view_hwnd, &mut rect).is_err() {
                continue;
            }

            let location_url = browser.LocationURL().ok().map(|url| url.to_string());
            let folder_path = get_shell_view_folder_path(&shell_view);
            let mut client_point = *screen_point;
            if !ScreenToClient(shell_view_hwnd, &mut client_point).as_bool() {
                continue;
            }

            let context = ActiveShellViewContext {
                shell_view_hwnd: shell_view_hwnd_key,
                location_url,
                folder_path,
                shell_view,
                client_point,
            };

            if !cursor_hwnd.is_invalid() && hwnd_is_same_or_ancestor(cursor_hwnd, shell_view_hwnd) {
                return Some(context);
            }

            if point_in_rect(screen_point, &rect) && rect_candidate.is_none() {
                rect_candidate = Some(context);
                continue;
            }

            if foreground == HWND(browser_hwnd.0 as *mut _) {
                foreground_candidate = Some(context);
            }
        }

        rect_candidate.or(foreground_candidate)
    }
}

fn get_active_shell_view_context_at_cursor() -> Option<ActiveShellViewContext> {
    unsafe {
        let mut cursor_pos = POINT::default();
        if GetCursorPos(&mut cursor_pos).is_err() {
            return None;
        }
        get_active_shell_view_context(&cursor_pos)
    }
}

fn get_current_explorer_location_url() -> Option<String> {
    if let Some(context) = get_active_shell_view_context_at_cursor() {
        if let Some(url) = context.location_url {
            return Some(url);
        }
    }

    let hwnd = get_explorer_hwnd_under_cursor_or_foreground()?;
    get_explorer_location_url(hwnd)
}

fn hwnd_is_same_or_ancestor(child: HWND, ancestor: HWND) -> bool {
    if child.is_invalid() || ancestor.is_invalid() {
        return false;
    }

    let mut current = child;
    for _ in 0..32 {
        if current == ancestor {
            return true;
        }

        match unsafe { windows::Win32::UI::WindowsAndMessaging::GetParent(current) } {
            Ok(parent) if !parent.is_invalid() && parent != current => {
                current = parent;
            }
            _ => break,
        }
    }

    false
}

fn is_current_search_view_legacy() -> bool {
    get_current_explorer_location_url_legacy()
        .as_deref()
        .map(is_search_ms_url)
        .unwrap_or(false)
}

fn get_current_hover_resolver_hints() -> HoverResolverHints {
    let mut hints = HoverResolverHints::default();

    if let Some(context) = get_active_shell_view_context_at_cursor() {
        hints.current_folder = context.folder_path.clone();
        hints.location_url = context.location_url.clone();
        hints.shell_view_hwnd = Some(context.shell_view_hwnd);

        hints.is_search_view = is_probable_search_view_context(&context);
        if hints.is_search_view {
            hints.search_root = resolve_search_root_from_context(&context);
            if hints.current_folder.is_none() {
                hints.current_folder = hints.search_root.clone();
            }
        }
    }

    if hints.current_folder.is_none() {
        hints.current_folder = get_current_explorer_folder();
    }
    if hints.is_search_view && hints.search_root.is_none() {
        hints.search_root = get_current_explorer_search_root();
    }

    hints
}

fn point_in_rect(point: &POINT, rect: &RECT) -> bool {
    point.x >= rect.left && point.x < rect.right && point.y >= rect.top && point.y < rect.bottom
}

fn normalize_existing_path(path: PathBuf) -> Option<PathBuf> {
    if !path.exists() {
        return None;
    }

    std::fs::canonicalize(&path).ok().or(Some(path))
}

fn normalize_media_path(path: PathBuf) -> Option<PathBuf> {
    if !path.exists() || !is_media_file(&path) {
        return None;
    }

    normalize_existing_path(path)
}

fn get_active_shell_view_under_cursor(screen_point: &POINT) -> Option<(IShellView, POINT)> {
    let context = get_active_shell_view_context(screen_point)?;
    Some((context.shell_view, context.client_point))
}

fn shell_item_to_path(shell_item: &IShellItem) -> Option<PathBuf> {
    unsafe {
        let display_name = shell_item
            .GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING)
            .ok()?;
        let path_string = display_name.to_string().ok();
        CoTaskMemFree(Some(display_name.0 as *const core::ffi::c_void));
        let path_string = path_string?;

        normalize_existing_path(PathBuf::from(path_string))
    }
}

fn shell_item_to_media_path(shell_item: &IShellItem) -> Option<PathBuf> {
    shell_item_to_path(shell_item).and_then(normalize_media_path)
}

fn get_shell_view_folder_path(shell_view: &IShellView) -> Option<String> {
    let folder_view = shell_view.cast::<IFolderView>().ok()?;

    unsafe {
        let persist_folder = folder_view.GetFolder::<IPersistFolder2>().ok()?;
        let pidl = persist_folder.GetCurFolder().ok()?;
        let shell_item = SHCreateItemFromIDList::<IShellItem>(pidl).ok();
        CoTaskMemFree(Some(pidl as *const core::ffi::c_void));
        shell_item
            .and_then(|shell_item| shell_item_to_path(&shell_item))
            .filter(|path| path.is_dir())
            .map(|path| path.to_string_lossy().into_owned())
    }
}

fn shell_item_from_view_pidl(
    folder_view: &IFolderView,
    pidl: *mut windows::Win32::UI::Shell::Common::ITEMIDLIST,
) -> Option<IShellItem> {
    unsafe {
        let direct = SHCreateItemFromIDList::<IShellItem>(pidl).ok();
        if direct.is_some() {
            return direct;
        }

        let shell_folder = folder_view.GetFolder::<IShellFolder>().ok()?;
        SHCreateItemWithParent::<_, IShellItem>(None, &shell_folder, pidl).ok()
    }
}

fn get_shell_data_model_file_from_context(context: &ActiveShellViewContext) -> Option<PathBuf> {
    if let Ok(hit_test) = context.shell_view.cast::<INameSpaceTreeControl>() {
        unsafe {
            if let Ok(shell_item) = hit_test.HitTest(&context.client_point) {
                if let Some(path) = shell_item_to_media_path(&shell_item) {
                    return Some(path);
                }
            }
        }
    }

    None
}

fn get_shell_data_model_file_under_cursor_fast() -> Option<PathBuf> {
    let context = get_active_shell_view_context_at_cursor()?;
    get_shell_data_model_file_from_context(&context)
}

fn get_current_explorer_search_root() -> Option<String> {
    if let Some(context) = get_active_shell_view_context_at_cursor() {
        if is_probable_search_view_context(&context) {
            return resolve_search_root_from_context(&context);
        }
    }

    let hwnd = get_explorer_hwnd_under_cursor_or_foreground()?;
    let hwnd_key = hwnd.0 as isize;
    let url = get_explorer_location_url(hwnd)?;
    if is_search_ms_url(&url) {
        resolve_explorer_location_folder(hwnd_key, &url)
    } else {
        None
    }
}

fn get_current_explorer_search_root_legacy() -> Option<String> {
    if let Some(context) = get_active_shell_view_context_at_cursor() {
        if let Some(url) = context.location_url.as_deref() {
            if is_search_ms_url(url) {
                return resolve_explorer_location_folder(context.shell_view_hwnd, url)
                    .or_else(|| get_shell_view_search_root(context.shell_view_hwnd))
                    .or_else(|| get_cached_explorer_real_folder(context.shell_view_hwnd));
            }
        }
    }

    let hwnd = get_explorer_hwnd_under_cursor_or_foreground_legacy()?;
    let hwnd_key = hwnd.0 as isize;
    let url = get_explorer_location_url(hwnd)?;
    if is_search_ms_url(&url) {
        resolve_explorer_location_folder(hwnd_key, &url)
    } else {
        None
    }
}

fn get_focused_shell_view_media_path(item: &FocusedItemInfo) -> Option<PathBuf> {
    let focus_point = POINT {
        x: item.rect.left + (item.rect.right - item.rect.left) / 2,
        y: item.rect.top + (item.rect.bottom - item.rect.top) / 2,
    };
    let (shell_view, _) = get_active_shell_view_under_cursor(&focus_point)?;
    let folder_view = shell_view.cast::<IFolderView>().ok()?;

    unsafe {
        for item_index in [
            folder_view.GetFocusedItem().ok(),
            folder_view.GetSelectionMarkedItem().ok(),
        ]
        .into_iter()
        .flatten()
        {
            let pidl = match folder_view.Item(item_index) {
                Ok(pidl) if !pidl.is_null() => pidl,
                _ => continue,
            };
            let shell_item = shell_item_from_view_pidl(&folder_view, pidl);
            CoTaskMemFree(Some(pidl as *const core::ffi::c_void));

            if let Some(path) =
                shell_item.and_then(|shell_item| shell_item_to_media_path(&shell_item))
            {
                return Some(path);
            }
        }
    }

    None
}

fn build_shell_view_media_index(view_hwnd_key: isize) -> Option<ShellViewMediaIndex> {
    let mut by_display_name = HashMap::new();
    let mut by_file_name = HashMap::new();
    let mut by_stem = HashMap::new();
    let mut root_folder: Option<PathBuf> = None;

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
            let service_provider = match browser.cast::<IServiceProvider>() {
                Ok(service_provider) => service_provider,
                Err(_) => continue,
            };
            let shell_browser: IShellBrowser =
                match service_provider.QueryService(&SID_STopLevelBrowser) {
                    Ok(shell_browser) => shell_browser,
                    Err(_) => continue,
                };
            let shell_view = match shell_browser.QueryActiveShellView() {
                Ok(shell_view) => shell_view,
                Err(_) => continue,
            };
            let shell_view_hwnd = match shell_view.GetWindow() {
                Ok(hwnd) if !hwnd.is_invalid() => hwnd,
                _ => continue,
            };
            if shell_view_hwnd.0 as isize != view_hwnd_key {
                continue;
            }

            let document = browser.Document().ok()?;
            let shell_view = document.cast::<IShellFolderViewDual>().ok()?;
            let folder = shell_view.Folder().ok()?;
            let items = folder.Items().ok()?;
            let item_count = items.Count().ok()?;
            if item_count > SHELL_VIEW_INDEX_SYNC_ITEM_LIMIT {
                return None;
            }
            let item_count = item_count.min(SHELL_VIEW_INDEX_MAX_ITEMS);

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

                if let Some(stem) = path.file_stem().and_then(|s| s.to_str()) {
                    by_stem
                        .entry(stem.to_ascii_lowercase())
                        .or_insert_with(|| path.clone());
                }

                root_folder = merge_common_folder_root(root_folder, &path);
            }

            return Some(ShellViewMediaIndex {
                built_at: Instant::now(),
                by_display_name,
                by_file_name,
                by_stem,
                root_folder: root_folder.map(|path| path.to_string_lossy().into_owned()),
            });
        }
    }

    None
}

fn build_legacy_search_shell_view_media_index(
    browser_hwnd_key: isize,
) -> Option<ShellViewMediaIndex> {
    let mut by_display_name = HashMap::new();
    let mut by_file_name = HashMap::new();
    let mut by_stem = HashMap::new();
    let mut root_folder: Option<PathBuf> = None;

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
            if browser_hwnd.0 as isize != browser_hwnd_key {
                continue;
            }

            let document = browser.Document().ok()?;
            let shell_view = document.cast::<IShellFolderViewDual>().ok()?;
            let folder = shell_view.Folder().ok()?;
            let items = folder.Items().ok()?;
            let item_count = items.Count().ok()?;
            if item_count > SHELL_VIEW_INDEX_SYNC_ITEM_LIMIT {
                return None;
            }
            let item_count = item_count.min(SHELL_VIEW_INDEX_MAX_ITEMS);

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

                if let Some(stem) = path.file_stem().and_then(|s| s.to_str()) {
                    by_stem
                        .entry(stem.to_ascii_lowercase())
                        .or_insert_with(|| path.clone());
                }

                root_folder = merge_common_folder_root(root_folder, &path);
            }

            return Some(ShellViewMediaIndex {
                built_at: Instant::now(),
                by_display_name,
                by_file_name,
                by_stem,
                root_folder: root_folder.map(|path| path.to_string_lossy().into_owned()),
            });
        }
    }

    None
}

fn lookup_path_in_shell_view_index(
    index: &ShellViewMediaIndex,
    item_name: &str,
) -> Option<PathBuf> {
    let item_name_lower = item_name.to_ascii_lowercase();
    let item_stem_lower = Path::new(item_name)
        .file_stem()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());
    let item_ext_lower = Path::new(item_name)
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());

    if let Some(path) = index
        .by_file_name
        .get(&item_name_lower)
        .or_else(|| index.by_display_name.get(&item_name_lower))
    {
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

fn find_media_in_shell_view(view_hwnd_key: isize, item_name: &str) -> Option<PathBuf> {
    let item_name = item_name.trim();
    if item_name.is_empty() {
        return None;
    }

    let mut cache = SHELL_VIEW_MEDIA_INDEX.lock().ok()?;
    cache.retain(|_, index| {
        index.built_at.elapsed() <= Duration::from_millis(SHELL_VIEW_INDEX_TTL_MS)
    });

    if !cache.contains_key(&view_hwnd_key) {
        let index = build_shell_view_media_index(view_hwnd_key)?;
        cache.insert(view_hwnd_key, index);
    }

    let index = cache.get(&view_hwnd_key)?;
    lookup_path_in_shell_view_index(index, item_name)
}

fn find_media_in_shell_view_legacy(browser_hwnd_key: isize, item_name: &str) -> Option<PathBuf> {
    let item_name = item_name.trim();
    if item_name.is_empty() {
        return None;
    }

    let mut cache = LEGACY_SEARCH_SHELL_VIEW_MEDIA_INDEX.lock().ok()?;
    cache.retain(|_, index| {
        index.built_at.elapsed() <= Duration::from_millis(SHELL_VIEW_INDEX_TTL_MS)
    });

    if !cache.contains_key(&browser_hwnd_key) {
        let index = build_legacy_search_shell_view_media_index(browser_hwnd_key)?;
        cache.insert(browser_hwnd_key, index);
    }

    let index = cache.get(&browser_hwnd_key)?;
    lookup_path_in_shell_view_index(index, item_name)
}

fn find_media_in_current_shell_view(item_name: &str) -> Option<PathBuf> {
    let context = get_active_shell_view_context_at_cursor()?;
    find_media_in_shell_view(context.shell_view_hwnd, item_name)
}

fn find_media_in_current_shell_view_legacy(item_name: &str) -> Option<PathBuf> {
    if let Some(context) = get_active_shell_view_context_at_cursor() {
        if let Some(path) = find_media_in_shell_view(context.shell_view_hwnd, item_name) {
            return Some(path);
        }
    }

    let hwnd = get_explorer_hwnd_under_cursor_or_foreground_legacy()?;
    find_media_in_shell_view_legacy(hwnd.0 as isize, item_name)
}

fn lookup_media_in_hover_folder(
    item_name: &str,
    current_folder_hint: Option<&str>,
) -> Option<PathBuf> {
    if let Some(folder) = current_folder_hint {
        if let Some(path) = find_media_in_folder(folder, item_name) {
            return Some(path);
        }
    }

    get_current_explorer_folder().and_then(|folder| find_media_in_folder(&folder, item_name))
}

fn add_search_index_path(index: &mut SearchRootMediaIndex, path: &PathBuf) {
    if let Some(file_name) = path.file_name().and_then(|s| s.to_str()) {
        index
            .by_file_name
            .entry(file_name.to_ascii_lowercase())
            .or_default()
            .push(path.clone());
    }

    if let Some(stem) = path.file_stem().and_then(|s| s.to_str()) {
        index
            .by_stem
            .entry(stem.to_ascii_lowercase())
            .or_default()
            .push(path.clone());
    }
}

fn build_search_root_media_index(root: &str) -> Option<SearchRootMediaIndex> {
    let root_path = PathBuf::from(root);
    if !root_path.is_dir() {
        return None;
    }

    let mut index = SearchRootMediaIndex {
        built_at: Instant::now(),
        by_file_name: HashMap::new(),
        by_stem: HashMap::new(),
    };
    let mut dirs = vec![root_path];
    let mut scanned_dirs = 0usize;
    let mut indexed_files = 0usize;

    while let Some(dir) = dirs.pop() {
        if scanned_dirs >= SEARCH_ROOT_INDEX_MAX_DIRS
            || indexed_files >= SEARCH_ROOT_INDEX_MAX_FILES
        {
            break;
        }
        scanned_dirs += 1;

        let entries = match std::fs::read_dir(&dir) {
            Ok(entries) => entries,
            Err(_) => continue,
        };

        for entry in entries.flatten() {
            let file_type = match entry.file_type() {
                Ok(file_type) => file_type,
                Err(_) => continue,
            };
            let path = entry.path();

            if file_type.is_dir() {
                dirs.push(path);
                continue;
            }

            if file_type.is_file() && is_media_file(&path) {
                add_search_index_path(&mut index, &path);
                indexed_files += 1;
                if indexed_files >= SEARCH_ROOT_INDEX_MAX_FILES {
                    break;
                }
            }
        }
    }

    Some(index)
}

fn lookup_path_in_search_root_index(
    index: &SearchRootMediaIndex,
    item_name: &str,
) -> Option<PathBuf> {
    let item_name_lower = item_name.to_ascii_lowercase();
    let item_stem_lower = Path::new(item_name)
        .file_stem()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());
    let item_ext_lower = Path::new(item_name)
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| s.to_ascii_lowercase());

    if let Some(paths) = index.by_file_name.get(&item_name_lower) {
        if let Some(path) = paths.first() {
            return Some(path.clone());
        }
    }

    if let Some(stem_key) = item_stem_lower.as_ref() {
        if let Some(paths) = index.by_stem.get(stem_key) {
            for path in paths {
                if let Some(item_ext) = item_ext_lower.as_deref() {
                    if let Some(candidate_ext) = path.extension().and_then(|s| s.to_str()) {
                        let candidate_ext_lower = candidate_ext.to_ascii_lowercase();
                        if candidate_ext_lower == item_ext
                            || (is_jpeg_extension(&candidate_ext_lower)
                                && is_jpeg_extension(item_ext))
                        {
                            return Some(path.clone());
                        }
                    }
                } else {
                    return Some(path.clone());
                }
            }
        }
    }

    None
}

fn queue_search_root_index_build(root: String) {
    let root_path = PathBuf::from(&root);
    if !root_path.is_dir() {
        return;
    }

    let should_build = {
        let cache_is_fresh = SEARCH_ROOT_MEDIA_INDEX
            .lock()
            .ok()
            .and_then(|cache| cache.get(&root).map(|index| index.built_at.elapsed()))
            .map(|age| age <= Duration::from_millis(SEARCH_ROOT_INDEX_TTL_MS))
            .unwrap_or(false);

        if cache_is_fresh {
            false
        } else if let Ok(mut building) = SEARCH_ROOT_INDEX_BUILDING.lock() {
            if building.contains(&root) {
                false
            } else {
                building.insert(root.clone());
                true
            }
        } else {
            false
        }
    };

    if !should_build {
        return;
    }

    std::thread::spawn(move || {
        let built_index = build_search_root_media_index(&root);
        if let Some(index) = built_index {
            if let Ok(mut cache) = SEARCH_ROOT_MEDIA_INDEX.lock() {
                cache.insert(root.clone(), index);
                if cache.len() > SEARCH_ROOT_CACHE_MAX_ENTRIES {
                    let mut oldest_key: Option<String> = None;
                    let mut oldest_age = Duration::ZERO;
                    for (cache_root, cache_index) in cache.iter() {
                        let age = cache_index.built_at.elapsed();
                        if oldest_key.is_none() || age > oldest_age {
                            oldest_key = Some(cache_root.clone());
                            oldest_age = age;
                        }
                    }
                    if let Some(oldest_key) = oldest_key {
                        cache.remove(&oldest_key);
                    }
                }
            }
        }
        if let Ok(mut building) = SEARCH_ROOT_INDEX_BUILDING.lock() {
            building.remove(&root);
        }
    });
}

fn lookup_media_in_search_root_index(root: &str, item_name: &str) -> Option<PathBuf> {
    let item_name = item_name.trim();
    if item_name.is_empty() {
        return None;
    }

    let mut has_fresh_index = false;

    if let Ok(mut cache) = SEARCH_ROOT_MEDIA_INDEX.lock() {
        cache.retain(|_, index| {
            index.built_at.elapsed() <= Duration::from_millis(SEARCH_ROOT_INDEX_TTL_MS)
        });

        if let Some(index) = cache.get(root) {
            has_fresh_index = true;
            if let Some(path) = lookup_path_in_search_root_index(index, item_name) {
                return Some(path);
            }
        }
    }

    if !has_fresh_index {
        queue_search_root_index_build(root.to_string());
    }

    None
}

/// Find which Explorer folder the cursor is currently over
fn get_current_explorer_folder() -> Option<String> {
    if let Some(context) = get_active_shell_view_context_at_cursor() {
        if let Some(folder) = context.folder_path {
            return Some(folder);
        }
        if is_probable_search_view_context(&context) {
            if let Some(folder) = resolve_search_root_from_context(&context) {
                return Some(folder);
            }
        }
        if let Some(url) = context.location_url.as_deref() {
            if let Some(folder) = resolve_explorer_location_folder(context.shell_view_hwnd, url) {
                return Some(folder);
            }
        }
        if let Some(folder) = get_cached_explorer_real_folder(context.shell_view_hwnd) {
            return Some(folder);
        }
    }

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
            for (explorer_hwnd, folder) in folders.iter() {
                if current_hwnd == HWND(*explorer_hwnd as *mut _) {
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
        for (explorer_hwnd, folder) in folders.iter() {
            if foreground == HWND(*explorer_hwnd as *mut _) {
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
                        if let Some(path) = resolve_media_path_from_text(&value_str) {
                            return Some(AccessibilityResult::FullPath(path));
                        }
                    }
                }

                // Try with the child variant first for name
                if is_variant_under_cursor(acc, &child_variant, &cursor_pos) {
                    if let Ok(name) = acc.get_accName(&child_variant) {
                        let name_str = name.to_string();
                        if !is_container_name(&name_str) {
                            if let Some(path) = resolve_media_path_from_text(&name_str) {
                                return Some(AccessibilityResult::FullPath(path));
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
                            if let Some(path) = resolve_media_path_from_text(&name_str) {
                                return Some(AccessibilityResult::FullPath(path));
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
                        if let Some(path) = resolve_media_path_from_text(&desc_str) {
                            return Some(AccessibilityResult::FullPath(path));
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
                            if let Some(path) = resolve_media_path_from_text(&name_str) {
                                return Some(AccessibilityResult::FullPath(path));
                            }
                            return Some(AccessibilityResult::FileName(name_str));
                        }
                    }
                }

                // Try to get value (path) from parent
                if is_variant_under_cursor(&parent_acc, &default_variant, cursor_pos) {
                    if let Ok(value) = parent_acc.get_accValue(&default_variant) {
                        let value_str = value.to_string();
                        if let Some(path) = resolve_media_path_from_text(&value_str) {
                            return Some(AccessibilityResult::FullPath(path));
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
                                if let Some(path) = resolve_media_path_from_text(&name_str) {
                                    return Some(AccessibilityResult::FullPath(path));
                                }
                                return Some(AccessibilityResult::FileName(name_str));
                            }
                        }

                        // Also try value for full path
                        if let Ok(value) = acc.get_accValue(&child_var) {
                            let value_str = value.to_string();
                            if let Some(path) = resolve_media_path_from_text(&value_str) {
                                return Some(AccessibilityResult::FullPath(path));
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
                                if let Some(path) = resolve_media_path_from_text(&name_str) {
                                    return Some(AccessibilityResult::FullPath(path));
                                }
                                return Some(AccessibilityResult::FileName(name_str));
                            }
                        }
                    }

                    // Try getting value from parent (may contain path)
                    if is_variant_under_cursor(&parent_acc, &default_variant, cursor_pos) {
                        if let Ok(value) = parent_acc.get_accValue(&default_variant) {
                            let value_str = value.to_string();
                            if let Some(path) = resolve_media_path_from_text(&value_str) {
                                return Some(AccessibilityResult::FullPath(path));
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

fn accessibility_result_from_name(name: String) -> Option<AccessibilityResult> {
    let name = name.trim().to_string();
    if name.is_empty() || is_container_name(&name) {
        return None;
    }

    if let Some(path) = resolve_media_path_from_text(&name) {
        return Some(AccessibilityResult::FullPath(path));
    }

    Some(AccessibilityResult::FileName(name))
}

fn accessibility_result_from_legacy_pattern(
    pattern: &IUIAutomationLegacyIAccessiblePattern,
) -> Option<AccessibilityResult> {
    unsafe {
        for candidate in [
            pattern.CurrentValue().ok().map(|s| s.to_string()),
            pattern.CurrentDescription().ok().map(|s| s.to_string()),
            pattern.CurrentName().ok().map(|s| s.to_string()),
        ]
        .into_iter()
        .flatten()
        {
            if let Some(result) = accessibility_result_from_name(candidate) {
                return Some(result);
            }
        }
    }

    None
}

fn accessibility_result_from_uia_element(
    element: &IUIAutomationElement,
) -> Option<AccessibilityResult> {
    unsafe {
        if let Ok(name) = element.CurrentName() {
            if let Some(result) = accessibility_result_from_name(name.to_string()) {
                return Some(result);
            }
        }

        let control_type = element.CurrentControlType().ok();
        if control_type == Some(UIA_ListItemControlTypeId)
            || control_type == Some(UIA_DataItemControlTypeId)
        {
            if let Ok(pattern) = element
                .GetCurrentPatternAs::<IUIAutomationLegacyIAccessiblePattern>(
                    UIA_LegacyIAccessiblePatternId,
                )
            {
                if let Some(result) = accessibility_result_from_legacy_pattern(&pattern) {
                    return Some(result);
                }
            }
        }
    }

    None
}

fn get_item_under_cursor_uia(automation: &IUIAutomation) -> Option<AccessibilityResult> {
    unsafe {
        let mut cursor_pos = POINT::default();
        if GetCursorPos(&mut cursor_pos).is_err() {
            return None;
        }

        let mut element = automation.ElementFromPoint(cursor_pos).ok()?;
        let walker = automation.ControlViewWalker().ok()?;

        for _ in 0..8 {
            if let Ok(rect) = element.CurrentBoundingRectangle() {
                if rect.left <= cursor_pos.x
                    && cursor_pos.x <= rect.right
                    && rect.top <= cursor_pos.y
                    && cursor_pos.y <= rect.bottom
                {
                    if let Some(result) = accessibility_result_from_uia_element(&element) {
                        return Some(result);
                    }
                }
            }

            element = walker.GetParentElement(&element).ok()?;
        }
    }

    None
}

fn get_accessibility_item_under_cursor(
    automation: Option<&IUIAutomation>,
) -> Option<AccessibilityResult> {
    get_item_under_cursor().or_else(|| automation.and_then(get_item_under_cursor_uia))
}

fn get_file_under_cursor_normal(
    automation: Option<&IUIAutomation>,
    hints: &HoverResolverHints,
) -> Option<PathBuf> {
    if let Some(path) = get_shell_data_model_file_under_cursor_fast() {
        return Some(path);
    }

    let item_info = get_accessibility_item_under_cursor(automation)?;

    match item_info {
        AccessibilityResult::FullPath(path) => {
            if is_media_file(&path) {
                Some(path)
            } else {
                None
            }
        }
        AccessibilityResult::FileName(item_name) => {
            if let Some(path) = resolve_media_path_from_text(&item_name) {
                return Some(path);
            }

            if let Some(folder) = hints
                .current_folder
                .as_deref()
                .map(str::to_string)
                .or_else(get_current_explorer_folder)
            {
                if let Some(path) = find_media_in_folder(&folder, &item_name) {
                    return Some(path);
                }
            }

            let potential_path = PathBuf::from(&item_name);
            if potential_path.is_absolute()
                && potential_path.exists()
                && is_media_file(&potential_path)
            {
                return Some(potential_path);
            }

            // Last-resort fallback only when we have no active-view context.
            if hints.current_folder.is_none() {
                let all_folders = get_all_explorer_folders();
                for (_, folder) in all_folders.iter() {
                    if let Some(path) = find_media_in_folder(folder, &item_name) {
                        return Some(path);
                    }
                }
            }

            None
        }
    }
}

fn get_file_under_cursor_search_legacy(
    automation: Option<&IUIAutomation>,
    hints: &HoverResolverHints,
) -> Option<PathBuf> {
    if let Some(path) = get_shell_data_model_file_under_cursor_fast() {
        return Some(path);
    }

    let item_info = get_accessibility_item_under_cursor(automation)?;

    match item_info {
        AccessibilityResult::FullPath(path) => {
            if is_media_file(&path) {
                Some(path)
            } else {
                None
            }
        }
        AccessibilityResult::FileName(item_name) => {
            if let Some(path) = resolve_media_path_from_text(&item_name) {
                return Some(path);
            }

            let current_is_search_view = hints.is_search_view || is_current_search_view_legacy();
            let current_search_root = hints.search_root.clone().or_else(|| {
                if current_is_search_view {
                    get_current_explorer_search_root_legacy()
                } else {
                    None
                }
            });

            if let Some(root) = current_search_root.as_deref() {
                if let Some(path) = find_media_in_folder(root, &item_name) {
                    return Some(path);
                }
            }

            if let Some(path) = find_media_in_current_shell_view_legacy(&item_name) {
                return Some(path);
            }
            if let Some(root) = current_search_root.as_deref() {
                if let Some(path) = lookup_media_in_search_root_index(root, &item_name) {
                    return Some(path);
                }
            }

            if current_is_search_view {
                return None;
            }

            if let Some(folder) = get_current_explorer_folder() {
                if let Some(path) = find_media_in_folder(&folder, &item_name) {
                    return Some(path);
                }
            }

            let all_folders = get_all_explorer_folders();
            for (_, folder) in all_folders.iter() {
                if let Some(path) = find_media_in_folder(folder, &item_name) {
                    return Some(path);
                }
            }

            None
        }
    }
}

fn get_file_under_cursor(
    automation: Option<&IUIAutomation>,
    hints: &HoverResolverHints,
) -> Option<PathBuf> {
    if hints.is_search_view || is_current_search_view_legacy() {
        return get_file_under_cursor_search_legacy(automation, hints);
    }

    get_file_under_cursor_normal(automation, hints)
}

fn get_file_under_cursor_checked(
    automation: Option<&IUIAutomation>,
    hints: &HoverResolverHints,
    slow_probe_count: &mut u32,
) -> Option<PathBuf> {
    let started = Instant::now();
    let result = get_file_under_cursor(automation, hints);

    if started.elapsed() >= Duration::from_millis(EXPLORER_PROBE_SLOW_MS) {
        *slow_probe_count = slow_probe_count.saturating_add(1);
    } else {
        *slow_probe_count = 0;
    }

    result
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

/// Allocation-free equivalent of lowercasing the class name and searching for an
/// ASCII needle. A `u16` outside ASCII becomes U+FFFD under `to_string_lossy`
/// and can never match, so it is skipped.
fn utf16_contains_ascii_ignore_case(haystack: &[u16], needle: &[u8]) -> bool {
    if needle.is_empty() || haystack.len() < needle.len() {
        return false;
    }

    haystack.windows(needle.len()).any(|window| {
        window
            .iter()
            .zip(needle)
            .all(|(ch, byte)| *ch < 0x80 && (*ch as u8).eq_ignore_ascii_case(byte))
    })
}

fn explorer_browser_class_matches(hwnd: HWND) -> bool {
    unsafe {
        let mut class_name = [0u16; 256];
        let len = GetClassNameW(hwnd, &mut class_name);
        if len <= 0 {
            return false;
        }

        let class_name = &class_name[..len as usize];
        utf16_contains_ascii_ignore_case(class_name, b"cabinetwclass")
            || utf16_contains_ascii_ignore_case(class_name, b"explorerwclass")
    }
}

unsafe extern "system" fn enum_explorer_windows_callback(hwnd: HWND, lparam: LPARAM) -> BOOL {
    let counts = &mut *(lparam.0 as *mut ExplorerWindowCounts);

    if explorer_browser_class_matches(hwnd) {
        counts.total += 1;
        if IsWindowVisible(hwnd).as_bool() && !is_window_minimized(hwnd) {
            counts.visible += 1;
        }
    }

    BOOL(1)
}

/// Get count of Explorer windows and count of visible (not minimized) ones.
/// Uses top-level HWND enumeration instead of ShellWindows COM to avoid
/// making Explorer's shell automation providers allocate during idle polling.
fn get_explorer_window_counts() -> (usize, usize) {
    let mut counts = ExplorerWindowCounts {
        total: 0,
        visible: 0,
    };

    unsafe {
        let _ = EnumWindows(
            Some(enum_explorer_windows_callback),
            LPARAM(&mut counts as *mut ExplorerWindowCounts as isize),
        );
    }

    (counts.total, counts.visible)
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

    // Check Explorer browser windows without ShellWindows COM.
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

/// Keyboard navigation state: `active` is true while a navigation key is held
/// or was pressed since the previous poll, and `pressed` is the fresh press
/// transition alone. A held key keeps reporting `active` forever, so a folder
/// change uses `pressed` to tell a new key press apart from state left over
/// from the navigation that opened the folder.
fn keyboard_navigation_input_state() -> (bool, bool) {
    key_input_state(&[
        VK_UP, VK_DOWN, VK_LEFT, VK_RIGHT, VK_HOME, VK_END, VK_PRIOR, VK_NEXT,
    ])
}

/// Left, right and middle buttons are deliberate input even when the cursor
/// never moves: the folder a double-click opens is user navigation, not a
/// background change the preview has to wait out.
fn mouse_button_input_state() -> (bool, bool) {
    key_input_state(&mouse_press_buttons())
}

/// Enter opens the focused item, so it drives folder changes without any
/// pointer input at all.
fn activation_key_input_state() -> (bool, bool) {
    key_input_state(&[VK_RETURN])
}

fn key_input_state(
    keys: &[windows::Win32::UI::Input::KeyboardAndMouse::VIRTUAL_KEY],
) -> (bool, bool) {
    unsafe {
        let mut active = false;
        let mut pressed = false;
        for &key in keys {
            let state = GetAsyncKeyState(key.0 as i32) as u16;
            if is_pressed_or_down_state(state) {
                active = true;
            }
            if (state & 0x0001) != 0 {
                pressed = true;
            }
        }

        (active, pressed)
    }
}

fn folder_probe_interval_ms(preview_active: bool) -> u64 {
    if preview_active {
        FOLDER_PROBE_MS
    } else {
        IDLE_FOLDER_PROBE_MS
    }
}

fn is_key_down_state(state: u16) -> bool {
    (state & 0x8000) != 0
}

fn is_key_down(vk: i32) -> bool {
    unsafe { is_key_down_state(GetAsyncKeyState(vk) as u16) }
}

fn is_explorer_navigation_shortcut_key(key_vk: i32, alt_down: bool, ctrl_down: bool) -> bool {
    key_vk == VK_BACK_CODE
        || (alt_down
            && matches!(
                key_vk,
                key if key == VK_LEFT.0 as i32 || key == VK_RIGHT.0 as i32 || key == VK_UP.0 as i32
            ))
        || (ctrl_down && key_vk == VK_T_CODE)
}

fn is_explorer_navigation_shortcut_detected() -> bool {
    let alt_down = is_key_down(VK_MENU_CODE);
    let ctrl_down = is_key_down(VK_CONTROL_CODE);
    let shortcut_keys = [
        VK_BACK_CODE,
        VK_LEFT.0 as i32,
        VK_RIGHT.0 as i32,
        VK_UP.0 as i32,
        VK_T_CODE,
    ];

    shortcut_keys.iter().any(|&key_vk| unsafe {
        is_pressed_or_down_state(GetAsyncKeyState(key_vk) as u16)
            && is_explorer_navigation_shortcut_key(key_vk, alt_down, ctrl_down)
    })
}

fn mouse_navigation_buttons() -> [windows::Win32::UI::Input::KeyboardAndMouse::VIRTUAL_KEY; 2] {
    [VK_XBUTTON1, VK_XBUTTON2]
}

fn is_mouse_navigation_button_detected() -> bool {
    mouse_navigation_buttons()
        .iter()
        .any(|&key| unsafe { is_pressed_or_down_state(GetAsyncKeyState(key.0 as i32) as u16) })
}

fn mouse_press_buttons() -> [windows::Win32::UI::Input::KeyboardAndMouse::VIRTUAL_KEY; 3] {
    [VK_LBUTTON, VK_RBUTTON, VK_MBUTTON]
}

fn hover_location_key(hints: &HoverResolverHints) -> Option<String> {
    hints
        .current_folder
        .as_ref()
        .map(|folder| format!("folder:{folder}"))
        .or_else(|| {
            hints
                .search_root
                .as_ref()
                .map(|root| format!("search:{root}"))
        })
        .or_else(|| hints.location_url.as_ref().map(|url| format!("url:{url}")))
        .or_else(|| hints.shell_view_hwnd.map(|hwnd| format!("view:{hwnd}")))
}

fn is_pressed_or_down_state(state: u16) -> bool {
    (state & 0x8000) != 0 || (state & 0x0001) != 0
}

fn off_trigger_key_to_vk(key: &str) -> Option<i32> {
    let key = key.trim().to_ascii_lowercase();
    let vk = match key.as_str() {
        "alt" | "menu" => 0x12,
        "shift" => 0x10,
        "ctrl" | "control" => 0x11,
        "win" | "windows" | "meta" => 0x5B,
        "space" => 0x20,
        "tab" => 0x09,
        "enter" | "return" => 0x0D,
        "esc" | "escape" => 0x1B,
        "backspace" => 0x08,
        "capslock" | "caps_lock" => 0x14,
        "left" => 0x25,
        "up" => 0x26,
        "right" => 0x27,
        "down" => 0x28,
        "insert" | "ins" => 0x2D,
        "delete" | "del" => 0x2E,
        "home" => 0x24,
        "end" => 0x23,
        "pageup" | "pgup" => 0x21,
        "pagedown" | "pgdn" => 0x22,
        "lshift" | "leftshift" => 0xA0,
        "rshift" | "rightshift" => 0xA1,
        "lctrl" | "leftctrl" | "lcontrol" | "leftcontrol" => 0xA2,
        "rctrl" | "rightctrl" | "rcontrol" | "rightcontrol" => 0xA3,
        "lalt" | "leftalt" => 0xA4,
        "ralt" | "rightalt" => 0xA5,
        key if key.len() == 1 => {
            let byte = key.as_bytes()[0];
            if byte.is_ascii_alphabetic() {
                byte.to_ascii_uppercase() as i32
            } else if byte.is_ascii_digit() {
                byte as i32
            } else {
                return None;
            }
        }
        key if key.starts_with('f') => {
            let n = key[1..].parse::<i32>().ok()?;
            if (1..=24).contains(&n) {
                0x70 + (n - 1)
            } else {
                return None;
            }
        }
        _ => return None,
    };

    Some(vk)
}

/// Whether the resolved off-trigger virtual key is currently held.
fn key_is_down(vk: i32) -> bool {
    unsafe {
        let state = GetAsyncKeyState(vk) as u16;
        (state & 0x8000) != 0
    }
}

/// Whether the pointer is inside the region a scrollable text preview published,
/// reading the cursor position here: the caller runs before the polling loop has
/// read it for this iteration. Answers `false` without a syscall when nothing on
/// screen scrolls.
fn text_scroll_pointer_hold_now() -> bool {
    let mut cursor_pos = POINT::default();
    unsafe {
        if GetCursorPos(&mut cursor_pos).is_err() {
            return false;
        }
    }

    text_scroll_pointer_hold(cursor_pos.x, cursor_pos.y)
}

/// Check if cursor is currently over an Explorer window (regardless of foreground).
/// Keep this HWND/class based; calling ShellWindows here caused Explorer-side
/// COM providers to allocate while we were merely checking cursor position.
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

        for _ in 0..20 {
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
    let hwnd_key = hwnd.0 as isize;
    if hwnd_key == 0 {
        return false;
    }

    if let Ok(cache) = EXPLORER_WINDOW_CACHE.lock() {
        if let Some((is_explorer, cached_at)) = cache.get(&hwnd_key) {
            if cached_at.elapsed() <= Duration::from_millis(EXPLORER_WINDOW_CACHE_TTL_MS) {
                return *is_explorer;
            }
        }
    }

    let mut is_explorer = false;
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
            is_explorer = true;
        } else {
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
                    is_explorer = path_str.contains("explorer.exe");
                }
                let _ = windows::Win32::Foundation::CloseHandle(handle);
            }
        }
    }

    if let Ok(mut cache) = EXPLORER_WINDOW_CACHE.lock() {
        if cache.len() >= 512 {
            cache.retain(|_, (_, cached_at)| {
                cached_at.elapsed() <= Duration::from_millis(EXPLORER_WINDOW_CACHE_TTL_MS)
            });
        }
        cache.insert(hwnd_key, (is_explorer, Instant::now()));
    }

    is_explorer
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
            // Keyboard focus has a direct Shell view focused item even in search-ms
            // results. Prefer that full path when Explorer exposes it.
            if let Some(path) = get_focused_shell_view_media_path(item) {
                return Some(path);
            }

            // Try as a potential full path
            let potential_path = PathBuf::from(item_name);
            if potential_path.is_absolute()
                && potential_path.exists()
                && is_media_file(&potential_path)
            {
                return Some(potential_path);
            }

            if let Some(path) = lookup_media_in_hover_folder(item_name, None) {
                return Some(path);
            }

            let current_url = get_current_explorer_location_url();
            let current_is_search_view = current_url
                .as_deref()
                .map(is_search_ms_url)
                .unwrap_or(false);

            if !current_is_search_view {
                let all_folders = get_all_explorer_folders();
                for (_, folder) in all_folders.iter() {
                    if let Some(path) = find_media_in_folder(folder, item_name) {
                        return Some(path);
                    }
                }
                return None;
            }

            let current_search_root = get_current_explorer_search_root();

            // Search mode: keep the special resolver intact.
            if let Some(root) = current_search_root.as_deref() {
                if let Some(path) = find_media_in_folder(root, item_name) {
                    return Some(path);
                }
            }
            if let Some(path) = find_media_in_current_shell_view(item_name) {
                return Some(path);
            }
            if let Some(root) = current_search_root.as_deref() {
                if let Some(path) = lookup_media_in_search_root_index(root, item_name) {
                    return Some(path);
                }
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
    let mut suppressed = SuppressedHover::default();
    let mut pointer_pause = KeyboardPointerPause::default();
    let mut hover_start: Option<Instant> = None;
    let mut last_cursor_pos = POINT::default();

    // Keyboard hover state
    let mut keyboard_file: Option<PathBuf> = None;
    let mut last_focused_name: Option<String> = None;
    let mut is_keyboard_hover = false;
    let mut suppress_preview_until_cursor_leaves_preview = false;
    let mut stationary_search_miss_started_at: Option<Instant> = None;
    // Short grace after starting a video preview to avoid instant self-dismiss
    // while ffplay window is still initializing under the cursor.
    let mut video_hover_guard_until: Option<Instant> = None;
    // Folder/input gate state: suppress preview after folder changes until explicit user input.
    let mut last_cursor_location: Option<String> = None;
    let mut hover_resolver_hints = HoverResolverHints::default();
    let mut suspend_preview_until_user_input = false;
    let mut allow_keyboard_preview_on_first_observation = false;
    let mut folder_change_time: Option<Instant> = None;
    let mut suspended_initial_focus: Option<String> = None;
    // A folder change that follows recent input is user navigation: its gate
    // lifts on its own once the view has settled, so the item under a parked
    // cursor previews without a mouse move.
    let mut folder_change_user_initiated = false;
    // Navigation-key press transitions seen so far, and the value when the
    // current suspension started. Only a later press may lift that
    // suspension, which keeps key state left over from the navigation that
    // opened the folder from counting as new input.
    let mut keyboard_navigation_press_seq: u64 = 0;
    let mut keyboard_press_seq_at_suspend: u64 = 0;
    let mut last_folder_probe = Instant::now();
    let mut last_hover_probe = Instant::now();
    let mut last_keyboard_focus_probe = Instant::now();
    let mut last_user_input_at: Option<Instant> = None;
    let mut last_keyboard_navigation_input_at: Option<Instant> = None;
    // Set on a click, Enter or a navigation key press: the folder probe runs at
    // the faster cadence for a moment afterwards, so a folder that press opens
    // is noticed before the user's next key press instead of up to a full idle
    // interval later.
    let mut last_navigation_trigger_at: Option<Instant> = None;
    let mut stationary_hover_probe_done = false;

    // Safety net for a ffplay process that survived a stop (failed or
    // unconfirmed kill): while nothing is hovered it is re-checked and killed.
    let mut last_video_process_sweep = Instant::now();

    // Wheel scrolling moves the list under a stationary pointer, so the wheel
    // tick counter is the only signal that the hovered item changed (see
    // `wheel_input`).
    let mut consumed_wheel_ticks = wheel_input::wheel_tick_count();
    let mut last_wheel_tick_at: Option<Instant> = None;
    let mut scroll_probe = ScrollSettleProbe::default();

    // State for optimized polling
    let mut last_state_check = Instant::now();
    let mut current_state = ExplorerState::NoExplorerWindows;

    // Polling intervals based on state
    const DEEP_SLEEP_MS: u64 = 1000; // No Explorer windows - check once per second
    const LONG_SLEEP_MS: u64 = 500; // All minimized or hidden - check twice per second
    const MEDIUM_SLEEP_MS: u64 = 150; // Visible but not focused - moderate checking
    const ACTIVE_POLL_MS: u64 = 30; // Active focus - responsive polling
    const VIDEO_HOVER_DISMISS_GRACE_MS: u64 = 350;
    const HOVER_PROBE_MS: u64 = 60;
    const KEYBOARD_FOCUS_PROBE_MS: u64 = 80;
    const STATIONARY_SEARCH_MISS_HIDE_MS: u64 = 180;
    const EXPLORER_SLOW_PROBE_LIMIT: u32 = 3;
    const EXPLORER_PROBE_BACKOFF_MS: u64 = 1500;
    const VIDEO_PROCESS_SWEEP_MS: u64 = 1000;

    // How often to re-evaluate the state when in sleep modes
    const STATE_RECHECK_DEEP_MS: u64 = 2000; // When no Explorer windows
    const STATE_RECHECK_LONG_MS: u64 = 1000; // When minimized/hidden
    const STATE_RECHECK_MEDIUM_MS: u64 = 300; // When visible but not focused
    const STATE_RECHECK_ACTIVE_MS: u64 = 100; // When active

    let (mut config_snapshot, mut off_trigger_vk) = CONFIG
        .lock()
        .map(|c| {
            let snapshot = (
                c.preview_enabled,
                c.hover_delay_ms,
                c.enable_off_trigger_key,
                c.same_file_rehover_delay_ms,
            );
            // Resolved once per config change instead of once per tick.
            let vk = if c.enable_off_trigger_key {
                off_trigger_key_to_vk(&c.off_trigger_key)
            } else {
                None
            };
            (snapshot, vk)
        })
        .unwrap_or(((true, 0, true, 750), Some(0x12)));
    let mut slow_explorer_probe_count = 0u32;
    let mut explorer_probe_backoff_until: Option<Instant> = None;
    let mut last_display_signature = current_display_signature();

    while RUNNING.load(Ordering::SeqCst) {
        // Nothing is hovered, so any ffplay still alive is a leftover from a
        // stop that did not take effect: kill it before it lingers on screen.
        if last_file.is_none()
            && keyboard_file.is_none()
            && !is_keyboard_hover
            && last_video_process_sweep.elapsed() >= Duration::from_millis(VIDEO_PROCESS_SWEEP_MS)
        {
            last_video_process_sweep = Instant::now();
            kill_stray_video_process();
        }

        if let Some(display_signature) = current_display_signature() {
            if display_signature_changed(last_display_signature, display_signature) {
                last_display_signature = Some(display_signature);
                clear_shell_view_probe_caches();
                hide_preview();
                last_file = None;
                keyboard_file = None;
                is_keyboard_hover = false;
                suppressed.clear();
                pointer_pause.clear();
                stationary_search_miss_started_at = None;
                hover_start = None;
                video_hover_guard_until = None;
                stationary_hover_probe_done = false;
                suspend_preview_until_user_input = true;
                allow_keyboard_preview_on_first_observation = false;
                folder_change_user_initiated = false;
                folder_change_time = Some(Instant::now());
                suspended_initial_focus = None;
                keyboard_press_seq_at_suspend = keyboard_navigation_press_seq;
                hover_resolver_hints = HoverResolverHints::default();
                last_cursor_location = None;
                slow_explorer_probe_count = 0;
                explorer_probe_backoff_until =
                    Some(Instant::now() + Duration::from_millis(DISPLAY_CHANGE_BACKOFF_MS));
            } else {
                last_display_signature = Some(display_signature);
            }
        }

        if slow_explorer_probe_count >= EXPLORER_SLOW_PROBE_LIMIT
            && explorer_probe_backoff_until.is_none()
        {
            explorer_probe_backoff_until =
                Some(Instant::now() + Duration::from_millis(EXPLORER_PROBE_BACKOFF_MS));
            clear_shell_view_probe_caches();
            hide_preview();
            last_file = None;
            keyboard_file = None;
            is_keyboard_hover = false;
            suppressed.clear();
            pointer_pause.clear();
            stationary_search_miss_started_at = None;
            hover_start = None;
            video_hover_guard_until = None;
            stationary_hover_probe_done = false;
        }

        if let Some(until) = explorer_probe_backoff_until {
            if Instant::now() < until {
                if last_file.is_some() || keyboard_file.is_some() || is_keyboard_hover {
                    hide_preview();
                }
                last_file = None;
                keyboard_file = None;
                is_keyboard_hover = false;
                hover_start = None;
                video_hover_guard_until = None;
                stationary_hover_probe_done = false;
                std::thread::sleep(Duration::from_millis(MEDIUM_SLEEP_MS));
                continue;
            }

            explorer_probe_backoff_until = None;
            slow_explorer_probe_count = 0;
            current_state = get_explorer_state();
            last_state_check = Instant::now();
        }

        if let Ok(config) = CONFIG.lock() {
            config_snapshot = (
                config.preview_enabled,
                config.hover_delay_ms,
                config.enable_off_trigger_key,
                config.same_file_rehover_delay_ms,
            );
            off_trigger_vk = if config.enable_off_trigger_key {
                off_trigger_key_to_vk(&config.off_trigger_key)
            } else {
                None
            };
        }

        let preview_enabled = config_snapshot.0;
        let hover_delay_ms = config_snapshot.1;
        let enable_off_trigger_key = config_snapshot.2;
        let same_file_rehover_delay_ms = config_snapshot.3;

        let off_trigger_active = enable_off_trigger_key && off_trigger_vk.is_some_and(key_is_down);

        if off_trigger_active {
            if last_file.is_some() || keyboard_file.is_some() {
                hide_preview();
            }
            keyboard_file = None;
            last_file = None;
            suppressed.clear();
            pointer_pause.clear();
            stationary_search_miss_started_at = None;
            hover_start = None;
            last_focused_name = None;
            is_keyboard_hover = false;
            video_hover_guard_until = None;
            std::thread::sleep(Duration::from_millis(ACTIVE_POLL_MS));
            continue;
        }

        if !preview_enabled {
            if last_file.is_some() || keyboard_file.is_some() {
                hide_preview();
                last_file = None;
                suppressed.clear();
                pointer_pause.clear();
                stationary_search_miss_started_at = None;
                hover_start = None;
            }
            keyboard_file = None;
            last_focused_name = None;
            is_keyboard_hover = false;
            video_hover_guard_until = None;
            suspend_preview_until_user_input = false;
            allow_keyboard_preview_on_first_observation = false;
            folder_change_user_initiated = false;
            last_cursor_location = None;
            hover_resolver_hints = HoverResolverHints::default();
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
                    stationary_search_miss_started_at = None;
                    hover_start = None;
                    keyboard_file = None;
                    last_focused_name = None;
                    is_keyboard_hover = false;
                    video_hover_guard_until = None;
                    pointer_pause.clear();
                }
                std::thread::sleep(Duration::from_millis(sleep_ms));
                continue;
            }
            ExplorerState::VisibleNotFocused => {
                // Explorer is visible but not focused - do a quick cursor check
                // Only activate full polling if cursor is actually over Explorer.
                // A pointer that is using a text preview is not evidence that the
                // user has left: the preview is on top of Explorer, so the check
                // below cannot see Explorer under it.
                if !is_cursor_over_explorer_full() && !text_scroll_pointer_hold_now() {
                    if last_file.is_some() || keyboard_file.is_some() {
                        hide_preview();
                        last_file = None;
                        stationary_search_miss_started_at = None;
                        hover_start = None;
                        keyboard_file = None;
                        last_focused_name = None;
                        is_keyboard_hover = false;
                        video_hover_guard_until = None;
                        pointer_pause.clear();
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

            // Whether the pointer is using a text preview that scrolls: on the
            // preview, or inside the margin around it — which covers the gap it
            // crosses on its way from the file it belongs to. While that holds,
            // the preview is something the user is reading rather than something
            // in the way, so it is not dismissed, and the file under the pointer
            // is not resolved, so it cannot be replaced by whatever it covers.
            //
            // Read once here because more than one path in this loop asks: the
            // dismissal below, the hover resolver, the mouse hover delay, and the
            // "Explorer is visible but not focused" branch, where a pointer that
            // is not over Explorer is not a reason to close a preview either.
            let loop_now = Instant::now();

            // Read straight from the published region, with nothing held over from
            // the last tick: the moment the pointer is out of it, the preview is
            // treated the way it was before the pointer ever touched it.
            let text_scroll_hold = text_scroll_pointer_hold(cursor_pos.x, cursor_pos.y);

            let move_threshold = pointer_pause.move_threshold_px();
            let moved = (cursor_pos.x - last_cursor_pos.x).abs() > move_threshold
                || (cursor_pos.y - last_cursor_pos.y).abs() > move_threshold;
            // Read the navigation keys first: GetAsyncKeyState's "pressed since
            // the previous call" bit goes away with the first read of a key in
            // an iteration, and that fresh press is what a folder change has to
            // tell apart from a held key.
            let (keyboard_navigation_active, keyboard_navigation_press) =
                keyboard_navigation_input_state();
            let explorer_navigation_shortcut_input = is_explorer_navigation_shortcut_detected();
            let keyboard_navigation_input =
                explorer_navigation_shortcut_input || keyboard_navigation_active;
            let mouse_navigation_input = is_mouse_navigation_button_detected();
            let (mouse_button_input, mouse_button_press) = mouse_button_input_state();
            let (activation_key_input, activation_key_press) = activation_key_input_state();

            // A wheel tick only counts while the wheel is driving Explorer: the
            // pointer is over it, or over a keyboard preview that covers the
            // pointer while Explorer still receives the wheel. The counter is
            // always consumed so a tick seen over something else cannot be
            // replayed.
            let wheel_ticks = wheel_input::wheel_tick_count();
            let wheel_tick = wheel_ticks != consumed_wheel_ticks;
            if wheel_tick {
                consumed_wheel_ticks = wheel_ticks;
            }
            let keyboard_owns_pointer = is_keyboard_hover || pointer_pause.freezes_pointer();
            let wheel_scroll = wheel_tick
                && (is_cursor_over_explorer_full()
                    || (keyboard_owns_pointer && cursor_preview_hover().any()));
            if wheel_scroll {
                last_wheel_tick_at = Some(loop_now);
                scroll_probe.arm();

                if keyboard_owns_pointer {
                    // The wheel is the mouse taking over from the keyboard: close
                    // the keyboard preview, release the pointer, and end the
                    // recent-keyboard-input window so the keyboard cannot
                    // re-establish its preview while the wheel is driving. The file
                    // the keyboard showed is not latched, so the mouse may preview
                    // it again where the cursor ends up.
                    hide_preview();
                    keyboard_file = None;
                    is_keyboard_hover = false;
                    video_hover_guard_until = None;
                    pointer_pause.clear();
                    last_focused_name = None;
                    allow_keyboard_preview_on_first_observation = true;
                    last_keyboard_navigation_input_at = None;
                }
            }
            let scrolling = recent_elapsed_within(
                last_wheel_tick_at.map(|at| at.elapsed()),
                WHEEL_SCROLL_SETTLE_MS,
            );

            if moved
                || keyboard_navigation_input
                || mouse_navigation_input
                || mouse_button_input
                || activation_key_input
                || wheel_scroll
            {
                last_user_input_at = Some(loop_now);
            }
            if keyboard_navigation_input {
                last_keyboard_navigation_input_at = Some(loop_now);
            }
            if keyboard_navigation_press {
                keyboard_navigation_press_seq = keyboard_navigation_press_seq.wrapping_add(1);
                // The item a fresh press selects is the user's own choice, so it
                // must not be swallowed as a fresh baseline, even while no
                // baseline is stored (folder change, mouse move, startup).
                allow_keyboard_preview_on_first_observation = true;
            }
            if mouse_button_press || activation_key_press || keyboard_navigation_press {
                last_navigation_trigger_at = Some(loop_now);
            }

            if explorer_navigation_shortcut_input || mouse_navigation_input {
                if last_file.is_some() || keyboard_file.is_some() || is_keyboard_hover {
                    hide_preview();
                }
                last_file = None;
                keyboard_file = None;
                is_keyboard_hover = false;
                suppressed.clear();
                pointer_pause.clear();
                stationary_search_miss_started_at = None;
                hover_start = None;
                last_focused_name = None;
                video_hover_guard_until = None;
                suspend_preview_until_user_input = true;
                allow_keyboard_preview_on_first_observation = false;
                folder_change_user_initiated = false;
                // History navigation (Backspace, Alt+arrows, the mouse
                // back/forward buttons) changes the folder too: keep the faster
                // probe cadence through the hold and past the release so the new
                // location is recognized before the user's next key press.
                last_navigation_trigger_at = Some(loop_now);
                folder_change_time = Some(Instant::now());
                suspended_initial_focus = None;
                keyboard_press_seq_at_suspend = keyboard_navigation_press_seq;
                last_cursor_pos = cursor_pos;
                stationary_hover_probe_done = false;
                continue;
            }

            // A keyboard preview is placed next to the focused item, which can put
            // it right over the parked cursor. Decide from the preview's own box
            // whether the pointer is under it, and freeze every pointer-driven
            // trigger until the mouse is moved on purpose.
            if pointer_pause.is_watching() {
                pointer_pause.evaluate_box(cursor_pos, preview_screen_rect());
            }

            // Close as soon as the cursor touches the preview window. Keep
            // suppressing preview until the cursor leaves so a delayed spinner
            // or background load result cannot resurrect a stuck preview under
            // the pointer. Keyboard previews own the screen: they may cover the
            // parked cursor and are never dismissed by it.
            let preview_hover = if should_probe_preview_hover(
                is_keyboard_hover || pointer_pause.freezes_pointer(),
                last_file.is_some(),
                suppress_preview_until_cursor_leaves_preview,
            ) {
                cursor_preview_hover()
            } else {
                PreviewCursorHover::NONE
            };
            let over_image_preview = preview_hover.image;
            let over_video_preview = preview_hover.video;
            let over_any_preview = preview_hover.any();

            // A text preview is the exception to that rule: while the pointer is
            // using one, the preview is kept — see `text_scroll_hold` above for
            // what "using it" covers. A preview the pointer cannot work with keeps
            // the old behaviour, which is to close as soon as it is touched.
            let guard_active = video_hover_guard_until
                .map(|until| Instant::now() < until)
                .unwrap_or(false);
            let should_dismiss_for_preview_hover =
                (over_image_preview && !text_scroll_hold) || (over_video_preview && !guard_active);

            if should_dismiss_for_preview_hover
                || (suppress_preview_until_cursor_leaves_preview && over_any_preview)
            {
                suppress_preview_until_cursor_leaves_preview = true;
                if let Some(file) = last_file.clone() {
                    suppressed.suppress(file, false);
                }
                hide_preview();
                last_file = None;
                keyboard_file = None;
                is_keyboard_hover = false;
                stationary_search_miss_started_at = None;
                video_hover_guard_until = None;
                stationary_hover_probe_done = false;
                hover_start = Some(Instant::now());
                continue;
            }

            if suppress_preview_until_cursor_leaves_preview {
                suppress_preview_until_cursor_leaves_preview = false;
                stationary_hover_probe_done = false;
                hover_start = Some(Instant::now());
                last_cursor_pos = cursor_pos;
                continue;
            }

            // Detect folder/navigation changes and suspend preview until user input.
            // Probe at active-poll cadence only while a preview is visible; idle
            // polling keeps the slower cadence to avoid extra COM work. A click,
            // Enter or navigation key press takes that place for a moment: the
            // folder it opens has to be recognized before the user's next key
            // press, which otherwise lands in the folder-change gate below. While
            // the keyboard drives, this cursor-based probe stays off: it resolves
            // the window under the pointer, which is the keyboard preview itself.
            let preview_active =
                last_file.is_some() || keyboard_file.is_some() || is_keyboard_hover;
            let navigation_trigger_active = recent_elapsed_within(
                last_navigation_trigger_at.map(|at| at.elapsed()),
                FOLDER_PROBE_TRIGGER_MS,
            );
            if last_folder_probe.elapsed()
                >= Duration::from_millis(if navigation_trigger_active {
                    FOLDER_PROBE_MS
                } else {
                    folder_probe_interval_ms(preview_active)
                })
                && !is_keyboard_hover
                && !pointer_pause.freezes_pointer()
                && !text_scroll_hold
                && should_probe_hover_resolver(
                    preview_active,
                    moved,
                    last_user_input_at.map(|at| at.elapsed()),
                )
            {
                last_folder_probe = Instant::now();
                hover_resolver_hints = get_current_hover_resolver_hints();
                if let Some(location_key) = hover_location_key(&hover_resolver_hints) {
                    if last_cursor_location.as_ref() != Some(&location_key) {
                        if let Some(folder) = hover_resolver_hints.current_folder.clone() {
                            queue_folder_index_build(PathBuf::from(&folder), folder);
                        }
                        // A change that follows recent input is user navigation:
                        // the file under the parked cursor may preview as soon as
                        // the new view has settled, without a mouse move.
                        let user_navigation = recent_elapsed_within(
                            last_user_input_at.map(|at| at.elapsed()),
                            HOVER_RESOLVER_INPUT_GRACE_MS,
                        );
                        last_cursor_location = Some(location_key);
                        suspend_preview_until_user_input = true;
                        allow_keyboard_preview_on_first_observation = false;
                        folder_change_user_initiated = user_navigation;
                        folder_change_time = Some(Instant::now());
                        suspended_initial_focus = None;
                        hover_start = None;
                        last_focused_name = None;
                        // Reset cursor baseline so we don't mistake stale delta for movement.
                        last_cursor_pos = cursor_pos;
                        stationary_hover_probe_done = false;
                        // Drain stale GetAsyncKeyState flags from prior navigation,
                        // then remember the press count: only a later navigation key
                        // press may lift this suspension, so key state left over
                        // from the navigation that opened the folder cannot.
                        let _ = keyboard_navigation_input_state();
                        keyboard_press_seq_at_suspend = keyboard_navigation_press_seq;

                        if last_file.is_some() || keyboard_file.is_some() || is_keyboard_hover {
                            hide_preview();
                        }
                        last_file = None;
                        suppressed.clear();
                        pointer_pause.clear();
                        stationary_search_miss_started_at = None;
                        keyboard_file = None;
                        is_keyboard_hover = false;
                        video_hover_guard_until = None;
                    }
                }
            }

            // Hard gate: after folder change, do not preview until explicit user input.
            if suspend_preview_until_user_input {
                // Cooldown: ignore all input for 150ms after folder change to let
                // COM/accessibility settle and to avoid stale keyboard state.
                if let Some(change_time) = folder_change_time {
                    if change_time.elapsed() < Duration::from_millis(150) {
                        continue;
                    }
                }

                let navigation_press =
                    keyboard_navigation_press_seq != keyboard_press_seq_at_suspend;

                // A scroll is deliberate pointer input, so it releases the
                // suspension exactly like a mouse move. The armed probe is used
                // instead of the raw tick because the cooldown above bails out
                // before this check, which would swallow a one-notch scroll.
                // A folder opened by a click, Enter or a navigation key is user
                // navigation too, so its suspension lifts once the new view has
                // settled and the item under the parked cursor can preview
                // without a mouse move. A navigation key press after the change
                // releases it as well: that press is the user asking for the
                // keyboard preview and must not be swallowed as a baseline.
                if moved
                    || scroll_probe.is_pending()
                    || folder_change_user_initiated
                    || navigation_press
                {
                    suspend_preview_until_user_input = false;
                    allow_keyboard_preview_on_first_observation = navigation_press;
                    folder_change_user_initiated = false;
                    hover_start = Some(Instant::now());
                    stationary_hover_probe_done = false;
                    suspended_initial_focus = None;
                    folder_change_time = None;
                } else {
                    // Nothing released the gate yet. Keep watching UI Automation
                    // focus changes as a fallback for focus moves that no counted
                    // press explains, such as Explorer restoring focus while the
                    // new view is being built.
                    let mut keyboard_unlocked = false;
                    if should_probe_keyboard_focus(
                        last_keyboard_navigation_input_at.map(|at| at.elapsed()),
                    ) && is_foreground_explorer()
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
                        folder_change_user_initiated = false;
                        suspended_initial_focus = None;
                        folder_change_time = None;
                    } else {
                        continue;
                    }
                }
            }

            // A move is "the mouse driving Explorer": resolve the item under the
            // cursor, and drop the preview when that is no longer the file it
            // shows. Another file always takes over, even while the pointer is
            // inside a scrollable preview's region — the region is only about the
            // pointer being on its way to the preview or on it, and the block
            // below tells those two apart by whether anything is under the pointer
            // at all.
            if moved {
                last_cursor_pos = cursor_pos;
                stationary_search_miss_started_at = None;
                stationary_hover_probe_done = false;
                // A real move hands control back to the mouse and ends any
                // scroll gesture that was still settling.
                pointer_pause.clear();
                scroll_probe.disarm();

                // Mouse movement always takes priority - dismiss keyboard hover.
                // The keyboard preview may have been covering the cursor, so the
                // file it showed stays latched until the cursor reaches another
                // file, or the user navigates with the keyboard again.
                if is_keyboard_hover {
                    if let Some(file) = keyboard_file.clone() {
                        suppressed.suppress(file, true);
                    }
                    hide_preview();
                    keyboard_file = None;
                    is_keyboard_hover = false;
                    video_hover_guard_until = None;
                }
                // A mouse move leaves the keyboard focus baseline unknown, so the
                // next focus observed while the user is driving with the keyboard
                // acts immediately. Recording it as a fresh baseline instead would
                // swallow the first key press and keep the mouse preview on screen.
                last_focused_name = None;
                allow_keyboard_preview_on_first_observation = true;

                if let Some(suppressed_file) = suppressed.file.clone() {
                    if let Some(current_file) = get_file_under_cursor_checked(
                        uia.as_ref(),
                        &hover_resolver_hints,
                        &mut slow_explorer_probe_count,
                    ) {
                        if same_path(&suppressed_file, &current_file) {
                            hover_start = Some(Instant::now());
                            continue;
                        }
                        suppressed.clear();
                        stationary_search_miss_started_at = None;
                    }
                }

                // While moving (including list scrolling), avoid heavy accessibility
                // resolution and wait until hover is stable before probing media.
                if last_file.is_some() {
                    let mut keep_while_scrolling_preview = false;
                    if let Some(current_file) = get_file_under_cursor_checked(
                        uia.as_ref(),
                        &hover_resolver_hints,
                        &mut slow_explorer_probe_count,
                    ) {
                        if last_file
                            .as_ref()
                            .map(|last| same_path(last, &current_file))
                            .unwrap_or(false)
                        {
                            hover_start = Some(Instant::now());
                            continue;
                        }
                        // Another file is under the pointer, so this is a hover like
                        // any other and the preview gives way to it — even while the
                        // pointer is inside a scrollable preview's region, which is
                        // why the check below is the "no file at all" case only.
                        suppressed.clear();
                        stationary_search_miss_started_at = None;
                    } else if text_scroll_hold {
                        // Nothing under the pointer: it is on its way to (or on) the
                        // scrollable preview, which is not the user leaving the file
                        // it shows.
                        keep_while_scrolling_preview = true;
                    } else if let Some(file) = last_file.clone() {
                        suppressed.suppress(file, false);
                    }

                    if !keep_while_scrolling_preview {
                        hide_preview();
                        last_file = None;
                        video_hover_guard_until = None;
                    }
                }
                hover_start = Some(Instant::now());
                continue;
            }

            // The wheel is still turning, so any preview on screen belongs to a
            // file that has scrolled away. Hold the stability window open and
            // skip the focus/hover probes: the item that lands under the cursor
            // is resolved below once the list stops moving.
            if scrolling {
                hover_start = Some(loop_now);
                stationary_search_miss_started_at = None;
                stationary_hover_probe_done = false;
                continue;
            }

            // Mouse is stationary - check for keyboard navigation
            // Only when Explorer is the foreground window (keyboard input goes there)
            if should_probe_keyboard_focus(last_keyboard_navigation_input_at.map(|at| at.elapsed()))
                && is_foreground_explorer()
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
                            // The focus baseline is unknown — the mouse just moved, or
                            // the user unlocked a folder change with the keyboard — so
                            // this first observed item acts immediately instead of
                            // being recorded and waiting for a second key press.
                            last_focused_name = Some(focused_name.clone());
                            allow_keyboard_preview_on_first_observation = false;

                            // Dismiss any active mouse hover
                            if last_file.is_some() && !is_keyboard_hover {
                                hide_preview();
                                last_file = None;
                                suppressed.clear();
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
                                    suppress_preview_until_cursor_leaves_preview = false;
                                    pointer_pause.watch_for_box();
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

                        // Nothing to compare against yet and no keyboard input has
                        // claimed this focus, so record it as the baseline only.
                        last_focused_name = Some(focused_name);
                    } else if last_focused_name.as_ref() != Some(&focused_name) {
                        // Focused item changed - keyboard navigation detected
                        last_focused_name = Some(focused_name);
                        allow_keyboard_preview_on_first_observation = false;

                        // Dismiss any active mouse hover
                        if last_file.is_some() && !is_keyboard_hover {
                            hide_preview();
                            last_file = None;
                            suppressed.clear();
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
                                suppress_preview_until_cursor_leaves_preview = false;
                                pointer_pause.watch_for_box();
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

            // If keyboard hover is active, or the pointer is frozen under a
            // keyboard preview, skip mouse hover delay logic entirely. The same
            // goes for a pointer inside a scrollable text preview: the file under
            // it is not what the user is looking at, so nothing is hovered over.
            if is_keyboard_hover || pointer_pause.freezes_pointer() || text_scroll_hold {
                continue;
            }

            // Check if we've hovered long enough (mouse hover)
            if let Some(start) = hover_start {
                if start.elapsed() >= hover_delay {
                    if !should_probe_stationary_hover(stationary_hover_probe_done) {
                        if let Some(miss_started) = stationary_search_miss_started_at {
                            if miss_started.elapsed()
                                >= Duration::from_millis(STATIONARY_SEARCH_MISS_HIDE_MS)
                            {
                                match last_file.clone() {
                                    Some(file) => suppressed.suppress(file, false),
                                    None => suppressed.clear(),
                                }
                                hide_preview();
                                last_file = None;
                                stationary_search_miss_started_at = None;
                                video_hover_guard_until = None;
                            }
                        }
                        continue;
                    }
                    if last_hover_probe.elapsed() < Duration::from_millis(HOVER_PROBE_MS) {
                        continue;
                    }
                    last_hover_probe = Instant::now();

                    // A scroll gesture that has not settled yet owns the probe.
                    let scroll_driven = scroll_probe.is_pending();

                    // Try to get file under cursor
                    let resolved = get_file_under_cursor_checked(
                        uia.as_ref(),
                        &hover_resolver_hints,
                        &mut slow_explorer_probe_count,
                    );
                    // One probe per parked cursor, except while a scroll gesture is
                    // still settling: there the latch stays open until two
                    // consecutive probes agree on the item under the cursor.
                    stationary_hover_probe_done = scroll_probe.observe(resolved.as_ref());
                    if scroll_driven && !stationary_hover_probe_done {
                        continue;
                    }

                    if let Some(file_path) = resolved {
                        if !last_file
                            .as_ref()
                            .map(|last| same_path(last, &file_path))
                            .unwrap_or(false)
                        {
                            if suppressed.matches(&file_path) {
                                let required_delay = hover_delay_ms.max(same_file_rehover_delay_ms);
                                if !suppressed.rehover_allowed(required_delay) {
                                    continue;
                                }
                            }
                            suppressed.clear();
                            stationary_search_miss_started_at = None;
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
                        if scroll_driven {
                            // The list settled on something that is not a media
                            // file: drop the preview that scrolled away, the same
                            // way the mouse-move path does.
                            match last_file.clone() {
                                Some(file) => suppressed.suppress(file, false),
                                None => suppressed.clear(),
                            }
                            hide_preview();
                            last_file = None;
                            stationary_search_miss_started_at = None;
                            video_hover_guard_until = None;
                            hover_start = Some(Instant::now());
                            continue;
                        }

                        let search_view_active =
                            hover_resolver_hints.is_search_view || is_current_search_view_legacy();
                        if search_view_active {
                            let miss_started =
                                stationary_search_miss_started_at.get_or_insert_with(Instant::now);
                            if miss_started.elapsed()
                                >= Duration::from_millis(STATIONARY_SEARCH_MISS_HIDE_MS)
                            {
                                match last_file.clone() {
                                    Some(file) => suppressed.suppress(file, false),
                                    None => suppressed.clear(),
                                }
                                hide_preview();
                                last_file = None;
                                stationary_search_miss_started_at = None;
                                video_hover_guard_until = None;
                                hover_start = Some(Instant::now());
                                continue;
                            }
                        } else {
                            stationary_search_miss_started_at = None;
                        }
                    }
                }
            } else {
                // Initialize hover_start if not moving
                stationary_search_miss_started_at = None;
                stationary_hover_probe_done = false;
                hover_start = Some(Instant::now());
            }
        }
    }

    unsafe {
        CoUninitialize();
    }
}

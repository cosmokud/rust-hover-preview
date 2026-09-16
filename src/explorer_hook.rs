use crate::config::{PreviewType, TriggerKeyMode};
use crate::pdf_preview::is_pdf_file;
use crate::preview_window::{
    cursor_preview_hover, hide_preview, kill_stray_video_process, preview_screen_rect,
    show_preview, show_preview_keyboard, text_scroll_pointer_hold, PreviewCursorHover,
};
use crate::text_formats::matches_text_lists;
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
use windows::core::{w, Interface, IUnknown, VARIANT};
use windows::Win32::Foundation::{BOOL, HWND, LPARAM, POINT, RECT};
use windows::Win32::Graphics::Gdi::{ClientToScreen, ScreenToClient};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoTaskMemFree, CoUninitialize, IDataObject, IServiceProvider,
    CLSCTX_ALL, CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED,
};
use windows::Win32::System::Variant::{VariantClear, VT_I4};
use windows::Win32::UI::Accessibility::{
    CUIAutomation, CUIAutomationRegistrar, IUIAutomation, IUIAutomationCacheRequest,
    IUIAutomationElement, IUIAutomationLegacyIAccessiblePattern, IUIAutomationRegistrar,
    IUIAutomationSelectionPattern, IUIAutomationTreeWalker, IUIAutomationValuePattern,
    TreeScope_Descendants, TreeScope_Element, UIAutomationPropertyInfo, UIAutomationType_Int,
    UIA_BoundingRectanglePropertyId, UIA_ControlTypePropertyId, UIA_DataItemControlTypeId,
    UIA_LegacyIAccessiblePatternId, UIA_ListItemControlTypeId, UIA_NamePropertyId,
    UIA_NativeWindowHandlePropertyId, UIA_PROPERTY_ID, UIA_SelectionPatternId, UIA_ValuePatternId,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetAsyncKeyState, VK_DOWN, VK_END, VK_HOME, VK_LBUTTON, VK_LEFT, VK_MBUTTON, VK_NEXT, VK_PRIOR,
    VK_RBUTTON, VK_RETURN, VK_RIGHT, VK_UP, VK_XBUTTON1, VK_XBUTTON2,
};
use windows::Win32::UI::Shell::{
    IFolderView, IFolderView2, INameSpaceTreeControl, IPersistFolder2, IShellBrowser, IShellFolder,
    IShellFolderViewDual, IShellItem, IShellItemArray, IShellView, IShellWindows,
    ItemIndex_Property_GUID, SHCreateItemFromIDList, SHCreateItemWithParent,
    SHCreateShellItemArrayFromDataObject, SID_STopLevelBrowser, SIGDN_DESKTOPABSOLUTEPARSING,
    SIGDN_FILESYSPATH, SIGDN_NORMALDISPLAY, ShellWindows, SVGIO_ALLVIEW, SVGIO_SELECTION,
};
use windows::Win32::UI::WindowsAndMessaging::{
    EnumWindows, GetAncestor, GetClassNameW, GetCursorPos, GetForegroundWindow, GetSystemMetrics,
    GetWindowPlacement, GetWindowRect, GetWindowThreadProcessId, IsIconic, IsWindowVisible,
    WindowFromPoint, GA_ROOT, SM_CXVIRTUALSCREEN, SM_CYVIRTUALSCREEN, SM_XVIRTUALSCREEN,
    SM_YVIRTUALSCREEN, SW_SHOWMAXIMIZED, WINDOWPLACEMENT,
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
    /// Whether the walk read every item of the view. One that did not is replaced
    /// by a walk that runs off the hook thread, so a file the short walk did not
    /// reach is only missing for as long as that takes.
    complete: bool,
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

/// Everything the mouse path needs to resolve the item under the pointer the way
/// the Shell itself knows it: one UI Automation client whose property reads are
/// batched into a single round trip per element, Explorer's own item-position
/// property, the view the pointer is in, and the answer the last probe produced.
///
/// What it holds is as telling as what it does not: there is no folder index, no
/// view index and no search root here. A search across folders is answered by the
/// item the pointer is on rather than by a name, so nothing has to be walked,
/// remembered or kept warm for the pointer to have an answer.
struct MouseResolver {
    automation: Option<IUIAutomation>,
    /// The batched property request every element under the pointer is read
    /// with, so an element costs one crossing into the view's provider rather
    /// than one per property.
    cache: Option<IUIAutomationCacheRequest>,
    walker: Option<IUIAutomationTreeWalker>,
    /// Explorer's own `ItemIndex` property, registered once per process. `None`
    /// when the registrar refuses it, which leaves the raw route skipped and the
    /// routes after it to answer exactly as they did before.
    item_index_property: Option<UIA_PROPERTY_ID>,
    /// The Shell window collection, created once and kept: building it is the one
    /// call every lookup would otherwise repeat.
    shell_windows: Option<IShellWindows>,
    view: Option<CachedFolderView>,
    probe: Option<ProbeMemo>,
}

/// The view the pointer is in, kept while it is still the same view.
struct CachedFolderView {
    browser_hwnd: isize,
    shell_browser: IShellBrowser,
    /// The shell view's own identity, so a window that navigated is a different
    /// view even though the window is the same one.
    view_identity: *mut core::ffi::c_void,
    folder_view: IFolderView2,
}

/// The answer one point produced, kept for the rest of the loop tick.
struct ProbeMemo {
    point: POINT,
    answer: Option<PathBuf>,
}

/// The item the pointer is over, as the view's accessibility provider reports it.
struct HoveredItem {
    /// The item's position in the view, one-based as Explorer's own `ItemIndex`
    /// reports it — the one fact about a search result that a shared name cannot
    /// take away, because two results may share a name and only one of them is
    /// at this position.
    index: Option<i32>,
    /// The name the item goes by in the view, which may hide the extension.
    name: String,
    /// The item's legacy accessible value: for a file a search has surfaced this
    /// is normally the file's own path, which is what lets a view that reports no
    /// position still be answered without looking anything up.
    value: Option<String>,
}

impl HoveredItem {
    /// Whether a second look at the same point found the same item. What the view
    /// says about an item is only true of the item — a list can move under a
    /// parked pointer — so an answer is only taken when both looks agree.
    fn same_item(&self, other: &HoveredItem) -> bool {
        self.index == other.index && self.name == other.name
    }
}

impl MouseResolver {
    fn new(automation: Option<IUIAutomation>) -> Self {
        let item_index_property = register_item_index_property();

        let (cache, walker) = match automation.as_ref() {
            Some(automation) => unsafe {
                let cache = automation.CreateCacheRequest().ok();
                if let Some(cache) = cache.as_ref() {
                    let _ = cache.SetTreeScope(TreeScope_Element);
                    for property in [
                        UIA_ControlTypePropertyId,
                        UIA_BoundingRectanglePropertyId,
                        UIA_NativeWindowHandlePropertyId,
                        UIA_NamePropertyId,
                    ] {
                        let _ = cache.AddProperty(property);
                    }
                    if let Some(item_index) = item_index_property {
                        let _ = cache.AddProperty(item_index);
                    }
                    let _ = cache.AddPattern(UIA_LegacyIAccessiblePatternId);
                }
                (cache, automation.ControlViewWalker().ok())
            },
            None => (None, None),
        };

        Self {
            automation,
            cache,
            walker,
            item_index_property,
            shell_windows: unsafe {
                CoCreateInstance::<_, IShellWindows>(&ShellWindows, None, CLSCTX_ALL).ok()
            },
            view: None,
            probe: None,
        }
    }

    /// Drop what describes the view, because what describes the last one describes
    /// the wrong place once the window has navigated.
    fn forget_view(&mut self) {
        self.view = None;
    }

    /// One answer per loop tick: what a tick learned is not carried into the next
    /// one, where the list under a parked pointer may have moved on.
    fn forget_probe(&mut self) {
        self.probe = None;
    }

    fn probed_at(&self, point: POINT) -> Option<Option<PathBuf>> {
        self.probe
            .as_ref()
            .filter(|probe| probe.point.x == point.x && probe.point.y == point.y)
            .map(|probe| probe.answer.clone())
    }

    fn remember_probe(&mut self, point: POINT, answer: Option<PathBuf>) {
        self.probe = Some(ProbeMemo { point, answer });
    }
}

impl CachedFolderView {
    /// Whether the cached view is still the one the window is showing. A window
    /// that navigated is showing a different view, and what the old one knew about
    /// its items describes the place the window has left.
    fn is_current(&self) -> bool {
        unsafe {
            let browser_window = HWND(self.browser_hwnd as *mut core::ffi::c_void);
            if !IsWindowVisible(browser_window).as_bool() || IsIconic(browser_window).as_bool() {
                return false;
            }

            match self.shell_browser.QueryActiveShellView() {
                Ok(view) => view
                    .cast::<IUnknown>()
                    .map(|identity| Interface::as_raw(&identity) == self.view_identity)
                    .unwrap_or(false),
                Err(_) => false,
            }
        }
    }
}

/// Explorer's own `ItemIndex` property, asked of the UI Automation registrar.
///
/// The id a registered property is read under is a runtime value, so the GUID
/// Explorer publishes under the name `ItemIndex` has to be exchanged for it once
/// before the property can be read at all. Everything about this is optional: a
/// registrar that refuses the property leaves the poke points above unread, and
/// the pointer is answered by the routes that do not need a position.
fn register_item_index_property() -> Option<UIA_PROPERTY_ID> {
    unsafe {
        let registrar: IUIAutomationRegistrar =
            CoCreateInstance(&CUIAutomationRegistrar, None, CLSCTX_INPROC_SERVER).ok()?;
        let property = registrar
            .RegisterProperty(&UIAutomationPropertyInfo {
                guid: ItemIndex_Property_GUID,
                pProgrammaticName: w!("ItemIndex"),
                r#type: UIAutomationType_Int,
            })
            .ok()?;

        (property > 0).then_some(UIA_PROPERTY_ID(property))
    }
}

/// "Do not preview this file" latch, shared by the mouse hover path. It is a
/// delay, not a verdict: the file it names is held off the mouse path until the
/// same-file rehover delay has passed, and released sooner when the cursor
/// resolves another file. A keyboard preview that a mouse move dismissed latches
/// the file it showed the same way — long enough that the handover cannot flash
/// the file straight back, and no longer, because a pointer parked on that file
/// afterwards is a user asking for it and not a repeat of the handover.
#[derive(Default)]
struct SuppressedHover {
    file: Option<PathBuf>,
    started_at: Option<Instant>,
}

impl SuppressedHover {
    fn clear(&mut self) {
        self.file = None;
        self.started_at = None;
    }

    fn suppress(&mut self, file: PathBuf) {
        self.file = Some(file);
        self.started_at = Some(Instant::now());
    }

    fn matches(&self, path: &PathBuf) -> bool {
        self.file
            .as_ref()
            .map(|file| same_path(file, path))
            .unwrap_or(false)
    }

    fn rehover_allowed(&self, required_delay_ms: u64) -> bool {
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
    /// ignored on purpose so a parked mouse cannot cancel a keyboard preview — and
    /// the wider tolerance holds for the whole of the keyboard's turn, not only
    /// while the preview's box happens to cover the cursor.
    fn move_threshold_px(&self, keyboard_owns_screen: bool) -> i32 {
        if keyboard_owns_screen || self.freezes_pointer() {
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
/// A Shell view index that saw every item cost a walk of the whole view, so it is
/// kept longer than a short one — long enough that the walk is paid for by the
/// hovers that read it rather than by the ones that follow it.
const SHELL_VIEW_INDEX_COMPLETE_TTL_MS: u64 = 30000;
const SHELL_VIEW_INDEX_MAX_ITEMS: i32 = 50000;
const SHELL_VIEW_INDEX_SYNC_ITEM_LIMIT: i32 = 1000;
/// How long a shell view may be walked for an index before the walk is left where
/// it is. The walk happens on the hook thread — a COM call, a file check and a
/// media-gate test per item — so a view of hundreds of items would otherwise stop
/// the loop for as long as it takes, and every route that does not need an index
/// is asked after this one anyway. What has been collected by the deadline is
/// still an index.
const SHELL_VIEW_INDEX_BUILD_BUDGET_MS: u64 = 150;
const SEARCH_ROOT_INDEX_TTL_MS: u64 = 60000;
/// How long a lookup may walk the folders below one while the hook loop waits for
/// it. The walk answers the question a search asks — where below this folder does
/// that name live — and it is given only as long as a probe can afford: a small
/// tree is walked to the end inside it, a large one is answered by the index that
/// walks it in the background.
const SEARCH_DESCEND_BUDGET_MS: u64 = 40;
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
/// How far up from the element under the pointer the item that holds it is looked
/// for. The item is the nearest list row or data item; what lies between it and
/// the element under the pointer is the view's own chrome — an icon, a label, a
/// row's text — so the walk is short by nature, and it is bounded here anyway.
const POINTER_ITEM_ANCESTOR_LIMIT: usize = 8;
/// The most Shell windows that will be asked which one the pointer is in. A
/// collection that reports more than this is not one to walk for every probe.
const SHELL_WINDOW_LIMIT: i32 = 64;

static FOLDER_MEDIA_INDEX: Lazy<Mutex<HashMap<String, FolderMediaIndex>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));
static FOLDER_INDEX_BUILDING: Lazy<Mutex<HashSet<String>>> =
    Lazy::new(|| Mutex::new(HashSet::new()));
static EXPLORER_FOLDERS_CACHE: Lazy<Mutex<Option<ExplorerFoldersCache>>> =
    Lazy::new(|| Mutex::new(None));
static SHELL_VIEW_MEDIA_INDEX: Lazy<Mutex<HashMap<isize, ShellViewMediaIndex>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));
static SHELL_VIEW_INDEX_BUILDING: Lazy<Mutex<HashSet<isize>>> =
    Lazy::new(|| Mutex::new(HashSet::new()));
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

    let mut stale = false;

    if let Ok(cache) = FOLDER_MEDIA_INDEX.lock() {
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

            stale = index.built_at.elapsed() > Duration::from_millis(FOLDER_INDEX_TTL_MS);
        } else {
            stale = true;
        }
    }

    // Never block hover polling on a huge folder scan: the names the folder is
    // already known by are read while a walk of it that is due runs behind, so the
    // names do not disappear for as long as that walk takes.
    if stale {
        queue_folder_index_build(folder_path.clone(), folder_key.to_string());
    }
    None
}

fn is_image_file(path: &PathBuf) -> bool {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| IMAGE_EXTENSIONS.contains(&ext.to_lowercase().as_str()))
        .unwrap_or(false)
}

/// Whether a preview may be shown for `path`: the kind of preview it would get,
/// and whether that kind is switched on in the tray's `Toggle Preview Types`
/// submenu.
///
/// The kinds are asked the way the renderer asks them — a video first, then a
/// PDF, then the text lists, then the image extensions — so the two cannot
/// disagree about what a file is. A video goes first because only its content
/// settles the extensions it shares with text: a `.ts` carrying MPEG-TS packets
/// is a video however the gates stand, and one that does not is the TypeScript
/// source the text lists claim.
fn is_media_file(path: &PathBuf) -> bool {
    let Ok(config) = CONFIG.lock() else {
        return false;
    };

    if is_video_file(path) {
        return PreviewType::Videos.enabled_in(&config);
    }
    if is_pdf_file(path) {
        return PreviewType::Pdf.enabled_in(&config);
    }
    if matches_text_lists(path, &config.text_extensions, &config.text_names) {
        return PreviewType::Text.enabled_in(&config);
    }

    is_image_file(path) && PreviewType::Images.enabled_in(&config)
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
    let root = {
        let mut cache = SHELL_VIEW_MEDIA_INDEX.lock().ok()?;
        cache.retain(|_, index| shell_view_index_is_fresh(index));

        if !cache.contains_key(&view_hwnd_key) {
            let index = build_shell_view_media_index(
                view_hwnd_key,
                Some(Duration::from_millis(SHELL_VIEW_INDEX_BUILD_BUDGET_MS)),
            )?;
            cache.insert(view_hwnd_key, index);
        }

        let index = cache.get(&view_hwnd_key)?;
        let root = index.root_folder.clone();

        if !index.complete {
            // The folders of some of the items are only the folders those have in
            // common, which is a narrower root than the one the results really
            // share: the walk is finished off the hook thread, and the root read
            // after it is the one of them all.
            drop(cache);
            queue_shell_view_index_completion(view_hwnd_key);
        }

        root
    };

    root
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

/// The media files the view has selected, asked of the view itself.
///
/// The accessibility tree hands a search result over as a name, and a name is
/// ambiguous the moment two results share one — which is what a search across
/// folders produces, the same file name sitting in any number of them. The view's
/// own selection is not ambiguous: it names the item the keyboard is on, by
/// identity, wherever that item's file lives. A keyboard preview asks this first
/// for that reason, and takes an answer from it only when exactly one of the
/// selected files goes by the name of the item the focus is on.
fn shell_view_selected_media_paths(context: &ActiveShellViewContext) -> Vec<PathBuf> {
    let mut paths = Vec::new();

    unsafe {
        let Ok(data_object) = context
            .shell_view
            .GetItemObject::<IDataObject>(SVGIO_SELECTION)
        else {
            return paths;
        };
        let Ok(items) = SHCreateShellItemArrayFromDataObject::<_, IShellItemArray>(&data_object)
        else {
            return paths;
        };
        let Ok(count) = items.GetCount() else {
            return paths;
        };

        for index in 0..count.min(16) {
            if let Ok(item) = items.GetItemAt(index) {
                if let Some(path) = shell_item_to_media_path(&item) {
                    paths.push(path);
                }
            }
        }
    }

    paths
}

/// The path of the item whose cell holds a point, asked of the view's own grid.
///
/// This is the one answer a name cannot give: every item a view shows has a cell
/// of its own, and the cell the point falls in belongs to exactly one item,
/// whatever that item is called. It is what resolves a search result whose file
/// name another result shares, for the pointer and for the keyboard alike.
///
/// The cells are laid out top to bottom, so the row a point is in is found by
/// bisecting their positions — an item's y never decreases as its index grows —
/// and the item within that row by walking the few positions that share it. A
/// cell runs from its own position to the next one's, and the last one in a row
/// to the view's edge.
fn view_item_media_path_at_point(
    context: &ActiveShellViewContext,
    screen_point: POINT,
) -> Option<PathBuf> {
    let folder_view = context.shell_view.cast::<IFolderView>().ok()?;

    unsafe {
        let count = folder_view.ItemCount(SVGIO_ALLVIEW).ok()?;
        if count <= 0 {
            return None;
        }

        let mut origin = POINT::default();
        let view_window = HWND(context.shell_view_hwnd as *mut core::ffi::c_void);
        if !ClientToScreen(view_window, &mut origin).as_bool() {
            return None;
        }
        let local = POINT {
            x: screen_point.x - origin.x,
            y: screen_point.y - origin.y,
        };

        // An item's position is asked for by its own pidl, so each probe is the item
        // and the position it sits at; the pidl is given back right away, since only
        // the position matters here.
        let position_of = |index: i32| -> Option<POINT> {
            let pidl = folder_view.Item(index).ok()?;
            if pidl.is_null() {
                return None;
            }
            let position = folder_view.GetItemPosition(pidl).ok();
            CoTaskMemFree(Some(pidl as *const core::ffi::c_void));
            position
        };

        // The last item whose top is at or above the point: the row it is in, or the
        // row before it.
        let mut low = 0i32;
        let mut high = count - 1;
        let mut row_start = 0i32;
        while low <= high {
            let middle = low + (high - low) / 2;
            match position_of(middle) {
                Some(position) if position.y <= local.y => {
                    row_start = middle;
                    low = middle + 1;
                }
                Some(_) => high = middle - 1,
                None => return None,
            }
        }

        // Back to the first item of that row, then across it.
        let row_y = position_of(row_start)?.y;
        let mut first = row_start;
        for _ in 0..64 {
            match position_of(first - 1) {
                Some(position) if position.y == row_y => first -= 1,
                _ => break,
            }
        }

        for index in first..count.min(first + 64) {
            let position = match position_of(index) {
                Some(position) => position,
                None => break,
            };
            if position.y != row_y {
                break;
            }

            let cell_right = position_of(index + 1)
                .filter(|next| next.y == row_y)
                .map(|next| next.x)
                .unwrap_or(i32::MAX);

            if local.x >= position.x && local.x < cell_right {
                let pidl = match folder_view.Item(index) {
                    Ok(pidl) if !pidl.is_null() => pidl,
                    _ => return None,
                };
                let shell_item = shell_item_from_view_pidl(&folder_view, pidl);
                CoTaskMemFree(Some(pidl as *const core::ffi::c_void));
                return shell_item.and_then(|item| shell_item_to_media_path(&item));
            }
        }
    }

    None
}

/// The Shell view's own focused item, as the path it resolves to.
///
/// This is the route a keyboard preview asks first, and it is the view itself
/// that is asked — what it says is focused, or marked as the selection — rather
/// than anything the accessibility tree reports.
fn shell_view_focused_media_path(context: &ActiveShellViewContext) -> Option<PathBuf> {
    let folder_view = context.shell_view.cast::<IFolderView>().ok()?;

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

/// The point in the middle of a focused item's box.
fn focused_item_center(item: &FocusedItemInfo) -> POINT {
    POINT {
        x: item.rect.left + (item.rect.right - item.rect.left) / 2,
        y: item.rect.top + (item.rect.bottom - item.rect.top) / 2,
    }
}

/// Whether a path found for a keyboard preview is the file the focused item
/// stands for — the check that makes a path found from the item's box safe to
/// take, because a box is a place and the list can move under it.
fn focused_path_is_item(item: &FocusedItemInfo, path: &PathBuf) -> bool {
    match &item.result {
        AccessibilityResult::FullPath(focused_path) => same_path(focused_path, path),
        AccessibilityResult::FileName(name) => path_matches_item_name(path, name),
    }
}

/// The file a focused item stands for, asked at the item's own box the way the
/// pointer asks it about the item under the cursor.
///
/// A result from a search that spans folders is why the box is asked at all: the
/// file is in none of the folders the results are gathered under, so no lookup by
/// name can reach it, and the item itself is the only thing that knows where it
/// is. The pointer reads that answer from the shell data model for the item under
/// the cursor; the keyboard reads the same answer for the item the focus is on,
/// which is what makes a result that spans folders preview exactly as a hovered
/// one does.
///
/// The accessibility provider is asked at the item before this, in
/// `get_focused_explorer_item` — see `focused_item_probe_points`, which is where a
/// search result states the file it stands for. What is left for the box is the
/// view and the data model. Every answer is taken only when it names the item: a
/// list that scrolled under a rect read a moment earlier can put another file in
/// that box.
fn focused_item_path_at_box(item: &FocusedItemInfo) -> Option<PathBuf> {
    let point = focused_item_center(item);

    if let Some(context) = get_active_shell_view_context(&point) {
        // What the view has selected, asked before anything else: the keyboard's
        // focus is the selection, and the selection names the item by identity —
        // the answer a name cannot give when two results share one. Only a
        // selection that matches the focused item's name exactly once is taken from
        // it, so several selected files that share the name are left to the cells
        // below, which are not ambiguous at all.
        let mut matching = shell_view_selected_media_paths(&context)
            .into_iter()
            .filter(|path| focused_path_is_item(item, path));
        if let (Some(path), None) = (matching.next(), matching.next()) {
            return Some(path);
        }

        // The item whose cell holds the focused item's middle: the view's grid says
        // which file that is, by identity, however many results share its name.
        if let Some(path) = view_item_media_path_at_point(&context, point) {
            if focused_path_is_item(item, &path) {
                return Some(path);
            }
        }

        // The view's own focused item, then the item that sits in the box — the same
        // order the pointer tries them in.
        if let Some(path) = shell_view_focused_media_path(&context) {
            if focused_path_is_item(item, &path) {
                return Some(path);
            }
        }

        if let Some(path) = get_shell_data_model_file_from_context(&context) {
            if focused_path_is_item(item, &path) {
                return Some(path);
            }
        }
    }

    None
}

/// The points inside a focused item's box that are worth asking the accessibility
/// provider about, in the order they are worth asking.
///
/// A search result states the file it stands for in the value of the text it shows,
/// and that value lives on the text rather than on the item containing it — so the
/// text is what has to be asked. A list item exposes it as a child, an icon view
/// hangs it below the icon, and a content row writes it a line down from the top.
/// The item's own subtree is therefore walked and each element found is asked at
/// its middle: descendants rather than children, because the text can sit a level
/// or two down, and only for an element that *is* an item, so a list cannot hand
/// over its whole contents to be probed.
///
/// What the subtree does not expose is covered by points in the item's own box,
/// taken in bands across it — a fifth of the way down, three fifths, and four
/// fifths — because a view keeps its text in one of them: a details row centres it
/// on the row, a content row writes its name on the first line and its details on
/// the second, and an icon view hangs the label below the icon. Each band is asked
/// just inside the name column — a hand's width in, then a name's width in — and at
/// the item's middle, which is where a centred label is. Every point is only asked
/// when the ones before it answered nothing, so breadth costs a call only where a
/// single point would have spent one and found nothing.
fn focused_item_probe_points(
    automation: &IUIAutomation,
    element: &IUIAutomationElement,
    rect: &RECT,
) -> Vec<POINT> {
    let mut points = Vec::new();

    unsafe {
        let is_item = element
            .CurrentControlType()
            .map(|control_type| {
                control_type == UIA_ListItemControlTypeId
                    || control_type == UIA_DataItemControlTypeId
            })
            .unwrap_or(false);

        if is_item {
            if let Ok(condition) = automation.CreateTrueCondition() {
                if let Ok(descendants) = element.FindAll(TreeScope_Descendants, &condition) {
                    if let Ok(count) = descendants.Length() {
                        for index in 0..count.min(6) {
                            if let Ok(child) = descendants.GetElement(index) {
                                if let Ok(child_rect) = child.CurrentBoundingRectangle() {
                                    if child_rect.right > child_rect.left
                                        && child_rect.bottom > child_rect.top
                                    {
                                        points.push(POINT {
                                            x: child_rect.left
                                                + (child_rect.right - child_rect.left) / 2,
                                            y: child_rect.top
                                                + (child_rect.bottom - child_rect.top) / 2,
                                        });
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    let width = (rect.right - rect.left).max(1);
    let height = (rect.bottom - rect.top).max(1);
    let clamp = |x: i32| x.min(rect.right - 2).max(rect.left + 1);
    let center_x = clamp(rect.left + width / 2);

    let band_x = [
        clamp(rect.left + (height / 2).max(2)),
        clamp(rect.left + height * 2),
        center_x,
    ];

    for fraction in [1, 3, 4] {
        let y = rect.top + height * fraction / 5;
        for x in band_x {
            points.push(POINT { x, y });
        }
    }

    points
}

/// The path the accessibility provider gives up for a focused item, asked at each
/// of its points until one of them answers with a file the item names.
fn focused_item_provider_path(item_name: &str, points: &[POINT]) -> Option<PathBuf> {
    for point in points {
        if let Some(AccessibilityResult::FullPath(path)) = get_item_at_point(*point) {
            if path_matches_item_name(&path, item_name) {
                return Some(path);
            }
        }
    }

    None
}

/// Read a Shell view into an index, optionally under a time budget.
///
/// The walk costs a COM call, a file check and a media-gate test per item, all of
/// it on the thread that asks for it, so the caller that runs on the hook loop
/// gives it a budget and gets a short index back — one the file it wanted may not
/// be in. `complete` says which it is, and a short one is replaced by a walk with
/// no budget on a thread of its own, which is what keeps a missing file a matter
/// of waiting a moment rather than of never being found.
fn build_shell_view_media_index(
    view_hwnd_key: isize,
    budget: Option<Duration>,
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
            let build_deadline = budget.map(|budget| Instant::now() + budget);
            let mut complete = true;

            for item_index in 0..item_count {
                // The walk is left where it is once its budget is out, and says so:
                // a view of hundreds of items costs a COM call, a file check and a
                // media-gate test per item, all of it with the caller's thread
                // stopped, and the caller that cannot wait asks for a complete index
                // behind it instead.
                if let Some(deadline) = build_deadline {
                    if Instant::now() >= deadline {
                        complete = false;
                        break;
                    }
                }

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
                complete,
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
    cache.retain(|_, index| shell_view_index_is_fresh(index));

    if !cache.contains_key(&view_hwnd_key) {
        let index = build_shell_view_media_index(
            view_hwnd_key,
            Some(Duration::from_millis(SHELL_VIEW_INDEX_BUILD_BUDGET_MS)),
        )?;
        cache.insert(view_hwnd_key, index);
    }

    let index = cache.get(&view_hwnd_key)?;
    let found = lookup_path_in_shell_view_index(index, item_name);
    let needs_completion = !index.complete;

    if needs_completion {
        // A short index is only ever a stop-gap: the walk is finished off the hook
        // thread whether or not this lookup found what it asked for, so a file that
        // is not in it yet is a moment away rather than a mystery.
        queue_shell_view_index_completion(view_hwnd_key);
    }

    found
}

/// Whether a Shell view index is still worth reading. One that saw every item
/// cost a walk of the whole view and is kept longer than a short one, which the
/// next lookup will replace anyway.
fn shell_view_index_is_fresh(index: &ShellViewMediaIndex) -> bool {
    let ttl = if index.complete {
        SHELL_VIEW_INDEX_COMPLETE_TTL_MS
    } else {
        SHELL_VIEW_INDEX_TTL_MS
    };

    index.built_at.elapsed() <= Duration::from_millis(ttl)
}

/// Finish off a Shell view index the hook loop could not wait for, on a thread of
/// its own, and put it where the next lookup will read it.
///
/// The walk is the same one, without a budget: nothing waits for it, and the
/// hook loop is free while it runs. A view is built once at a time — a second
/// request for one already being built is dropped rather than raced.
fn queue_shell_view_index_completion(view_hwnd_key: isize) {
    let should_build = match SHELL_VIEW_INDEX_BUILDING.lock() {
        Ok(mut keys) => keys.insert(view_hwnd_key),
        Err(_) => false,
    };
    if !should_build {
        return;
    }

    std::thread::spawn(move || {
        // The walk is Shell COM, and this thread is not one of the app's own: the
        // apartment is claimed here, the way every other thread that talks to the
        // shell claims it, or every call the walk makes is refused with "not
        // initialized" and the completion quietly does nothing — which is the one
        // thing it must not do. It is the multi-threaded apartment, which needs no
        // message pump and so is the right one for a thread that has none.
        unsafe {
            let _ = CoInitializeEx(None, COINIT_MULTITHREADED);
        }

        if let Some(index) = build_shell_view_media_index(view_hwnd_key, None) {
            if let Ok(mut cache) = SHELL_VIEW_MEDIA_INDEX.lock() {
                cache.insert(view_hwnd_key, index);
            }
        }

        unsafe {
            CoUninitialize();
        }

        if let Ok(mut keys) = SHELL_VIEW_INDEX_BUILDING.lock() {
            keys.remove(&view_hwnd_key);
        }
    });
}

fn find_media_in_current_shell_view(item_name: &str) -> Option<PathBuf> {
    let context = get_active_shell_view_context_at_cursor()?;
    find_media_in_shell_view(context.shell_view_hwnd, item_name)
}

fn lookup_media_in_hover_folder(
    item_name: &str,
    current_folder_hint: Option<&str>,
) -> Option<PathBuf> {
    if let Some(folder) = current_folder_hint {
        if let Some(path) = find_media_in_folder(folder, item_name) {
            return Some(path);
        }
        if let Some(path) = lookup_media_below_folder(folder, item_name) {
            return Some(path);
        }
    }

    get_current_explorer_folder().and_then(|folder| {
        find_media_in_folder(&folder, item_name)
            .or_else(|| lookup_media_below_folder(&folder, item_name))
    })
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

/// Read the files under a folder into an index, optionally under a time budget.
///
/// The walk is the one a search's results are found by: every folder below the
/// one it starts at, and every media file in them, by name. It is offered a
/// budget by the callers that run on the hook loop, which get a short index back
/// and leave the rest to the walk that runs in the background.
fn build_search_root_media_index(
    root: &str,
    budget: Option<Duration>,
) -> Option<SearchRootMediaIndex> {
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
    let deadline = budget.map(|budget| Instant::now() + budget);

    while let Some(dir) = dirs.pop() {
        if scanned_dirs >= SEARCH_ROOT_INDEX_MAX_DIRS
            || indexed_files >= SEARCH_ROOT_INDEX_MAX_FILES
        {
            break;
        }
        if let Some(deadline) = deadline {
            if Instant::now() >= deadline {
                break;
            }
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

/// Find a file by name below a folder, walking the folders under it now.
///
/// This is the question a search answers: a search started in a folder returns
/// results from any folder below it, and the shell hands a result over as its
/// name — so when the name is not *in* the folder the search began at, the only
/// folders that can hold it are the ones underneath. The walk is the same one the
/// search-root index does, given a budget small enough for the hook loop to spend
/// between probes; what it does not reach in time is left to the index, which
/// walks the same tree in the background.
fn lookup_media_below_folder(folder: &str, item_name: &str) -> Option<PathBuf> {
    let index = build_search_root_media_index(
        folder,
        Some(Duration::from_millis(SEARCH_DESCEND_BUDGET_MS)),
    )?;

    lookup_path_in_search_root_index(&index, item_name)
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
        let built_index = build_search_root_media_index(&root, None);
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

    let mut has_index = false;
    let mut stale = false;

    if let Ok(cache) = SEARCH_ROOT_MEDIA_INDEX.lock() {
        if let Some(index) = cache.get(root) {
            has_index = true;
            stale = index.built_at.elapsed() > Duration::from_millis(SEARCH_ROOT_INDEX_TTL_MS);
            if let Some(path) = lookup_path_in_search_root_index(index, item_name) {
                return Some(path);
            }
        }
    }

    if !has_index || stale {
        // An index being old is a reason to walk the tree again, not a reason to
        // stop reading the names it holds: it is read while the walk that refreshes
        // it runs, so a name that was found a minute ago is not unfindable for as
        // long as that walk takes. Nothing else retires it — the cache bounds its
        // entries, and a fresh walk replaces the one it was built from.
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

        get_item_at_point(cursor_pos)
    }
}

/// Get the filename or full path of the item at a point, the way a pointer reads
/// it.
///
/// The point is what makes this shared: an item under the cursor and the item a
/// keyboard preview is about are the same question asked at two places, and every
/// answer below is Explorer's — a search result gives up the file it stands for in
/// its accessible value, whatever folder that file is in.
fn get_item_at_point(point: POINT) -> Option<AccessibilityResult> {
    unsafe {
        // Use accessibility to get the item info
        let mut accessible: Option<windows::Win32::UI::Accessibility::IAccessible> = None;
        let mut child_variant = VARIANT::default();

        let result = (|| -> Option<AccessibilityResult> {
            if windows::Win32::UI::Accessibility::AccessibleObjectFromPoint(
                point,
                &mut accessible,
                &mut child_variant,
            )
            .is_err()
            {
                return None;
            }

            if let Some(ref acc) = accessible {
                // First, try to get the value - this often contains the full path in search results
                if is_variant_under_cursor(acc, &child_variant, &point) {
                    if let Ok(value) = acc.get_accValue(&child_variant) {
                        let value_str = value.to_string();
                        if let Some(path) = resolve_media_path_from_text(&value_str) {
                            return Some(AccessibilityResult::FullPath(path));
                        }
                    }
                }

                // Try with the child variant first for name. A view can report text
                // that is not a file's name here — see `name_could_be_previewed` —
                // and the walk goes on to the parent item when it does.
                if is_variant_under_cursor(acc, &child_variant, &point) {
                    if let Ok(name) = acc.get_accName(&child_variant) {
                        let name_str = name.to_string();
                        if !is_container_name(&name_str) {
                            if let Some(path) = resolve_media_path_from_text(&name_str) {
                                return Some(AccessibilityResult::FullPath(path));
                            }
                            if name_could_be_previewed(&name_str) {
                                return Some(AccessibilityResult::FileName(name_str));
                            }
                        }
                    }
                }

                // Try with default variant
                let default_variant = VARIANT::default();
                if is_variant_under_cursor(acc, &default_variant, &point) {
                    if let Ok(name) = acc.get_accName(&default_variant) {
                        let name_str = name.to_string();
                        if !is_container_name(&name_str) {
                            if let Some(path) = resolve_media_path_from_text(&name_str) {
                                return Some(AccessibilityResult::FullPath(path));
                            }
                            if name_could_be_previewed(&name_str) {
                                return Some(AccessibilityResult::FileName(name_str));
                            }
                        }
                    }
                }

                // Try navigating parent chain to find item name (for list/details views)
                if let Some(result) = try_get_item_from_parent(acc, &child_variant, &point) {
                    return Some(result);
                }

                // Try getting help text which sometimes has info
                if is_variant_under_cursor(acc, &child_variant, &point) {
                    if let Ok(help) = acc.get_accHelp(&child_variant) {
                        let help_str = help.to_string();
                        if !help_str.is_empty()
                            && !is_container_name(&help_str)
                            && name_could_be_previewed(&help_str)
                        {
                            return Some(AccessibilityResult::FileName(help_str));
                        }
                    }
                }

                // Try description which may have path info
                if is_variant_under_cursor(acc, &child_variant, &point) {
                    if let Ok(desc) = acc.get_accDescription(&child_variant) {
                        let desc_str = desc.to_string();
                        if let Some(path) = resolve_media_path_from_text(&desc_str) {
                            return Some(AccessibilityResult::FullPath(path));
                        }
                    }
                }

                // Try to walk up parent hierarchy more aggressively (for details view text cells)
                if let Some(result) = try_deep_parent_search(acc, &point) {
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
                            if name_could_be_previewed(&name_str) {
                                return Some(AccessibilityResult::FileName(name_str));
                            }
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

/// Whether a name could be a file this app previews: one whose extension is in a
/// list the app knows, or a name the text lists carry.
///
/// A view can put text under the pointer that is not a file's name — the metadata
/// line of a Content view row, the value of a column — and the accessibility tree
/// reports it the same way it reports a file name. Taking it for the file under the
/// pointer sends every lookup below after a file called `12 KB` and finds nothing,
/// which is a preview that never appears even though a file *is* under the pointer;
/// skipping it lets the walk go on to the item that carries the real name. A name
/// that fails this is not a file the app could preview even if a folder held it, so
/// nothing is lost by passing it over.
fn name_could_be_previewed(name: &str) -> bool {
    let name = name.trim();
    if name.is_empty() {
        return false;
    }

    let path = PathBuf::from(name);
    if is_image_file(&path) || is_video_file(&path) || is_pdf_file(&path) {
        return true;
    }

    match CONFIG.lock() {
        Ok(config) => matches_text_lists(&path, &config.text_extensions, &config.text_names),
        // Without the text lists the answer is unknown, and a name that is not a
        // file is harmless: the lookups that follow simply find nothing.
        Err(_) => true,
    }
}

fn accessibility_result_from_name(name: String) -> Option<AccessibilityResult> {
    let name = name.trim().to_string();
    if name.is_empty() || is_container_name(&name) {
        return None;
    }

    if let Some(path) = resolve_media_path_from_text(&name) {
        return Some(AccessibilityResult::FullPath(path));
    }

    if !name_could_be_previewed(&name) {
        return None;
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

/// The item the pointer is over, as the view's accessibility provider reports it,
/// or nothing when the pointer is over no item at all.
fn uia_item_under_cursor(resolver: &MouseResolver, point: POINT) -> Option<HoveredItem> {
    let automation = resolver.automation.as_ref()?;
    let cache = resolver.cache.as_ref()?;
    let walker = resolver.walker.as_ref()?;

    unsafe {
        let mut element = automation.ElementFromPointBuildCache(point, cache).ok()?;

        for _ in 0..=POINTER_ITEM_ANCESTOR_LIMIT {
            if let Some(item) =
                hovered_item_from_element(&element, resolver.item_index_property, point)
            {
                return Some(item);
            }

            element = walker.GetParentElementBuildCache(&element, cache).ok()?;
        }
    }

    None
}

/// The item an element is, when the pointer is inside it.
///
/// The box is the whole test. A row is an item from its left edge to its right
/// one, and what a pointer anywhere on that row belongs to is the file the row
/// stands for — the same thing Explorer highlights when it is selected. An item
/// whose box does not hold the point is not an answer even though the walk passed
/// through it, and since the walk starts at the element under the pointer, the
/// first item that holds the point is the one the pointer is on.
fn hovered_item_from_element(
    element: &IUIAutomationElement,
    item_index_property: Option<UIA_PROPERTY_ID>,
    point: POINT,
) -> Option<HoveredItem> {
    unsafe {
        let control_type = element.CachedControlType().ok()?;
        if control_type != UIA_ListItemControlTypeId && control_type != UIA_DataItemControlTypeId {
            return None;
        }

        let bounds = element.CachedBoundingRectangle().ok()?;
        let holds_point = bounds.left < bounds.right
            && bounds.top < bounds.bottom
            && point.x >= bounds.left
            && point.x < bounds.right
            && point.y >= bounds.top
            && point.y < bounds.bottom;
        if !holds_point {
            return None;
        }

        let name = element
            .CachedName()
            .map(|value| value.to_string())
            .unwrap_or_default();
        let value = element
            .GetCachedPatternAs::<IUIAutomationLegacyIAccessiblePattern>(
                UIA_LegacyIAccessiblePatternId,
            )
            .ok()
            .and_then(|pattern| pattern.CachedValue().ok())
            .map(|value| value.to_string())
            .filter(|value| !value.trim().is_empty());

        Some(HoveredItem {
            index: cached_item_index(element, item_index_property),
            name: name.trim().to_string(),
            value,
        })
    }
}

/// The position the view holds an element at, as Explorer reports it.
///
/// The property is one-based, and zero is what the provider answers when it has
/// nothing to say about the element's position at all — so only a positive value
/// is a position. A missing one is not a failure: the item's name and its
/// accessible value still stand on their own.
fn cached_item_index(
    element: &IUIAutomationElement,
    item_index_property: Option<UIA_PROPERTY_ID>,
) -> Option<i32> {
    let property = item_index_property?;

    unsafe {
        let value = element.GetCachedPropertyValue(property).ok()?;
        let raw = value.as_raw().Anonymous.Anonymous;
        if raw.vt != VT_I4.0 {
            return None;
        }

        let index = raw.Anonymous.lVal;
        (index > 0).then_some(index)
    }
}

/// The root window the pointer is over, which is the window a Shell view has to
/// belong to for the items it draws to be the ones under the pointer.
fn root_window_at(point: POINT) -> Option<HWND> {
    unsafe {
        let window = WindowFromPoint(point);
        if window.is_invalid() {
            return None;
        }

        let root = GetAncestor(window, GA_ROOT);
        (!root.is_invalid()).then_some(root)
    }
}

/// Every Shell view registered for a window, deduplicated by the view's own
/// identity.
///
/// A window that holds several tabs registers one Shell window per tab, and every
/// one of them answers with the frame's own window — so the frame names a set of
/// views rather than one, and something else has to say which of them the pointer
/// is in. The tabs that are not showing are not skipped here: they are what the
/// item settles, and a view skipped here could be the one holding it.
fn folder_views_for_window(resolver: &MouseResolver, root_key: isize) -> Vec<CachedFolderView> {
    let mut candidates: Vec<CachedFolderView> = Vec::new();
    let Some(shell_windows) = resolver.shell_windows.as_ref() else {
        return candidates;
    };

    unsafe {
        let Ok(count) = shell_windows.Count() else {
            return candidates;
        };

        for index in 0..count.min(SHELL_WINDOW_LIMIT) {
            let Ok(dispatch) = shell_windows.Item(&VARIANT::from(index)) else {
                continue;
            };
            let Ok(browser) = dispatch.cast::<windows::Win32::UI::Shell::IWebBrowser2>() else {
                continue;
            };
            let Ok(handle) = browser.HWND() else {
                continue;
            };
            let browser_window = HWND(handle.0 as *mut core::ffi::c_void);
            if browser_window.0 as isize != root_key {
                continue;
            }
            if !IsWindowVisible(browser_window).as_bool() || IsIconic(browser_window).as_bool() {
                continue;
            }

            let Ok(service_provider) = browser.cast::<IServiceProvider>() else {
                continue;
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
            let Ok(identity) = shell_view.cast::<IUnknown>() else {
                continue;
            };
            let view_identity = Interface::as_raw(&identity);
            if candidates
                .iter()
                .any(|candidate| candidate.view_identity == view_identity)
            {
                continue;
            }
            let Ok(folder_view) = shell_view.cast::<IFolderView2>() else {
                continue;
            };

            candidates.push(CachedFolderView {
                browser_hwnd: root_key,
                shell_browser,
                view_identity,
                folder_view,
            });
        }
    }

    candidates
}

/// How many Shell windows are registered, which is what says whether a window the
/// pointer is in could be holding tabs.
fn shell_window_count(resolver: &MouseResolver) -> Option<i32> {
    unsafe { resolver.shell_windows.as_ref()?.Count().ok() }
}

/// The file the pointer's item stands for, asked of the view the pointer is in.
///
/// The view is the one the window under the pointer is showing, and a window that
/// holds tabs registers one Shell window per tab — all of them answering with the
/// frame's own window — so the frame names a set of views and not one. Which of
/// them it is, is settled by the item: each candidate is asked for the file at the
/// item's position under the item's name, and a view holding some other folder's
/// items has nothing there to answer with. What two of them do answer has to
/// agree: two tabs showing the same folder are one answer and either of them is
/// the right one, while two tabs whose items differ are a question the item cannot
/// settle — and no answer is better than the wrong tab's file.
fn pointer_item_file_path(
    resolver: &mut MouseResolver,
    point: POINT,
    item: &HoveredItem,
) -> Option<PathBuf> {
    let index = item.index? - 1;
    let root_key = root_window_at(point)?.0 as isize;

    // The view that answered last is asked first, but only while it is the *only*
    // registration for the window: with one view there is no second one for it to
    // disagree with, and a window holding tabs is never answered from the cache —
    // a cached tab is one of several, and the cache cannot say which of them is
    // showing.
    if shell_window_count(resolver) == Some(1) {
        if let Some(cached) = resolver.view.as_ref() {
            if cached.browser_hwnd == root_key && cached.is_current() {
                if let Some(path) = view_item_file_path(&cached.folder_view, index, &item.name) {
                    return Some(path);
                }
            }
        }
    }

    let mut answer: Option<(CachedFolderView, PathBuf)> = None;

    for candidate in folder_views_for_window(resolver, root_key) {
        let Some(path) = view_item_file_path(&candidate.folder_view, index, &item.name) else {
            continue;
        };

        match &answer {
            None => answer = Some((candidate, path)),
            Some((_, existing)) if !same_path(existing, &path) => return None,
            // Another tab showing the same folder is the same answer.
            Some(_) => {}
        }
    }

    let (view, path) = answer?;
    resolver.view = Some(view);
    Some(path)
}

/// The file system path the Shell holds for an item — the path the item *is*,
/// rather than a name it is shown under. An item that stands for no file on disk
/// (a library, a drive, a search root) has none, which is the Shell's own answer
/// that there is nothing here to preview.
fn shell_item_filesystem_path(item: &IShellItem) -> Option<PathBuf> {
    unsafe {
        let display_name = item.GetDisplayName(SIGDN_FILESYSPATH).ok()?;
        let path_string = display_name.to_string().ok();
        CoTaskMemFree(Some(display_name.0 as *const core::ffi::c_void));
        let path_string = path_string?;
        let path = PathBuf::from(path_string);

        (!path.as_os_str().is_empty()).then_some(path)
    }
}

/// Whether the name a view shows an item under is the name that was asked about.
/// Explorer labels a file with the name its view displays — the extension hidden
/// when the user has chosen to hide it — so the item's own display name is what
/// the asked-about name is compared with.
fn item_display_name_matches(item: &IShellItem, expected_name: &str) -> bool {
    unsafe {
        let Ok(display_name) = item.GetDisplayName(SIGDN_NORMALDISPLAY) else {
            return false;
        };
        let name = display_name.to_string().ok();
        CoTaskMemFree(Some(display_name.0 as *const core::ffi::c_void));

        name.map(|name| name.trim().eq_ignore_ascii_case(expected_name.trim()))
            .unwrap_or(false)
    }
}

/// The file the view's item at `index` stands for, taken from the Shell item the
/// view itself hands over.
///
/// The path is the Shell's own answer for the item at that position
/// (`SIGDN_FILESYSPATH`), not something put together from what the item is
/// called: a search whose results come from many folders holds any number of
/// files that share a name, and the position an item holds is the one thing a
/// shared name cannot take away. The name is still checked against the item's own,
/// because a list that moved between the two calls would otherwise hand over its
/// neighbour's path — and an answer that does not agree is no answer at all.
fn view_item_file_path(
    folder_view: &IFolderView2,
    index: i32,
    expected_name: &str,
) -> Option<PathBuf> {
    if index < 0 {
        return None;
    }

    unsafe {
        let item = folder_view.GetItem::<IShellItem>(index).ok()?;
        let path = shell_item_filesystem_path(&item)?;

        if !expected_name.is_empty()
            && name_could_be_previewed(expected_name)
            && !item_display_name_matches(&item, expected_name)
            && !path_matches_item_name(&path, expected_name)
        {
            return None;
        }

        normalize_media_path(path).filter(|path| path.is_file())
    }
}

/// The file the pointer is over.
///
/// The pointer asks one question — what is under me — and the view under it
/// answers by identity: the item the accessibility provider says the point is
/// inside, the position that item holds in the view, and the file that position
/// stands for. Nothing is looked up by name for the pointer, because a search
/// across folders is full of names that belong to more than one file and a name
/// is the one thing the view does not need. What follows the raw route is the
/// same answer asked of another witness, in the order that costs least: the
/// item's own accessible value when it carries the file's path, the view's own
/// cell under the point, and — for folder views, where a name really is
/// unambiguous — the item's name in the folder the pointer is in.
fn get_file_under_cursor(
    resolver: &mut MouseResolver,
    hints: &HoverResolverHints,
) -> Option<PathBuf> {
    let mut point = POINT::default();
    if unsafe { GetCursorPos(&mut point) }.is_err() {
        return None;
    }

    // A tick asks about the same point more than once — the move path asks for
    // the file it latched and then for the one on screen — and the answer is the
    // same both times. Nothing is carried past the tick: the list under a parked
    // pointer may have moved on by the next one.
    if let Some(answer) = resolver.probed_at(point) {
        return answer;
    }

    let answer = resolve_file_under_cursor(resolver, hints, point);
    resolver.remember_probe(point, answer.clone());
    answer
}

fn resolve_file_under_cursor(
    resolver: &mut MouseResolver,
    hints: &HoverResolverHints,
    point: POINT,
) -> Option<PathBuf> {
    if let Some(item) = uia_item_under_cursor(resolver, point) {
        if let Some(path) = pointer_item_file_path(resolver, point, &item) {
            // The view is asked twice: a wheel turns the list under a parked
            // pointer, and an item that is no longer at the point the first answer
            // described is not what that answer is about.
            if uia_item_under_cursor(resolver, point)
                .map(|again| again.same_item(&item))
                .unwrap_or(false)
            {
                return Some(path);
            }
        }

        // What the item says about itself: a search result carries the file's own
        // path in its accessible value.
        if let Some(value) = item.value.as_deref() {
            if let Some(path) = resolve_media_path_from_text(value) {
                return Some(path);
            }
        }
    }

    let item_info = get_accessibility_item_under_cursor(resolver.automation.as_ref())?;

    match item_info {
        AccessibilityResult::FullPath(path) => is_media_file(&path).then_some(path),
        AccessibilityResult::FileName(item_name) => {
            if let Some(path) = resolve_media_path_from_text(&item_name) {
                return Some(path);
            }

            // The cell the point falls in, asked of the view's own grid: the one
            // answer a name cannot give, for a view that reports no position.
            if let Some(context) = get_active_shell_view_context(&point) {
                if let Some(path) = view_item_media_path_at_point(&context, point) {
                    if path_matches_item_name(&path, &item_name) {
                        return Some(path);
                    }
                }
            }

            // A name is a witness only where it is unambiguous, and that is a
            // folder view: the search's results span folders, so a name that is in
            // any one of them is not evidence about the item under the pointer —
            // two results sharing a name is what the identity routes above exist
            // for, and a name lookup here would undo them.
            if !hints.is_search_view {
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

                // Last resort, and only when no view said which folder it is: the
                // folders the open Explorer windows are showing.
                if hints.current_folder.is_none() {
                    let all_folders = get_all_explorer_folders();
                    for (_, folder) in all_folders.iter() {
                        if let Some(path) = find_media_in_folder(folder, &item_name) {
                            return Some(path);
                        }
                    }
                }
            }

            None
        }
    }
}

fn get_file_under_cursor_checked(
    resolver: &mut MouseResolver,
    hints: &HoverResolverHints,
    slow_probe_count: &mut u32,
) -> Option<PathBuf> {
    let started = Instant::now();
    let result = get_file_under_cursor(resolver, hints);

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

/// What tells one keyboard focus observation from the next.
///
/// The name alone does not. A search whose results come from several folders can
/// hold the same file name in more than one of them — a `README.md` here and a
/// `README.md` there — and moving between the two would read as no change at all,
/// which leaves the preview on the file the keyboard came from and never asks the
/// new one up. The item's box is taken with the name for that reason: two files
/// that share a name do not share a row, and an item that did not move keeps its
/// box.
#[derive(Clone, PartialEq)]
struct FocusedItemKey {
    name: String,
    rect: (i32, i32, i32, i32),
}

impl FocusedItemKey {
    fn new(name: String, rect: &RECT) -> Self {
        Self {
            name,
            rect: (rect.left, rect.top, rect.right, rect.bottom),
        }
    }
}

/// Whether a path is the file an Explorer item's name stands for. The name a view
/// shows is compared whole first, and — when the name carries no extension of its
/// own, because the view hides it — to the path's stem. Two files that share a
/// stem but differ in extension are two files, so a name that does carry one only
/// matches the same extension, or a sibling in the JPEG family, which is the one
/// Explorer can label a file with instead of the extension it is stored under.
fn path_matches_item_name(path: &Path, item_name: &str) -> bool {
    let item_name = item_name.trim();
    if item_name.is_empty() {
        return false;
    }

    let full_name_matches = path
        .file_name()
        .and_then(|s| s.to_str())
        .map(|file_name| file_name.eq_ignore_ascii_case(item_name))
        .unwrap_or(false);
    if full_name_matches {
        return true;
    }

    let (Some(item_stem), Some(path_stem)) = (
        Path::new(item_name).file_stem().and_then(|s| s.to_str()),
        path.file_stem().and_then(|s| s.to_str()),
    ) else {
        return false;
    };
    if !item_stem.eq_ignore_ascii_case(path_stem) {
        return false;
    }

    match Path::new(item_name).extension().and_then(|s| s.to_str()) {
        None => true,
        Some(item_ext) => {
            let item_ext = item_ext.to_ascii_lowercase();
            path.extension()
                .and_then(|s| s.to_str())
                .map(|path_ext| {
                    let path_ext = path_ext.to_ascii_lowercase();
                    item_ext == path_ext
                        || (is_jpeg_extension(&item_ext) && is_jpeg_extension(&path_ext))
                })
                .unwrap_or(false)
        }
    }
}

/// The path a focused Explorer element carries, when what it carries is the item
/// it names.
///
/// A result from a search that spans folders is why this exists: its name is only
/// what the file is called, and every lookup that could turn a name into a path
/// searches the folders the results have in common, which is not the folder the
/// file is in. The element does carry the path — in the same accessible value the
/// pointer path reads a result's path from — so the keyboard path asks for it
/// there too, and takes it only when it names the item: a value belongs to the
/// element, and an element that is not the item cannot claim it. Both the value a
/// provider exposes as a pattern of its own and the one the legacy bridge carries
/// are asked: Explorer answers with either, depending on the view.
fn focused_item_media_path(element: &IUIAutomationElement, item_name: &str) -> Option<PathBuf> {
    let mut candidates: Vec<String> = Vec::new();

    unsafe {
        if let Ok(pattern) =
            element.GetCurrentPatternAs::<IUIAutomationValuePattern>(UIA_ValuePatternId)
        {
            if let Ok(value) = pattern.CurrentValue() {
                candidates.push(value.to_string());
            }
        }

        if let Ok(pattern) = element.GetCurrentPatternAs::<IUIAutomationLegacyIAccessiblePattern>(
            UIA_LegacyIAccessiblePatternId,
        ) {
            if let Ok(value) = pattern.CurrentValue() {
                candidates.push(value.to_string());
            }
            if let Ok(value) = pattern.CurrentDescription() {
                candidates.push(value.to_string());
            }
        }
    }

    for candidate in candidates {
        if let Some(path) = resolve_media_path_from_text(&candidate) {
            if path_matches_item_name(&path, item_name) {
                return Some(path);
            }
        }
    }

    None
}

/// Whether a UIA element names a file: an Explorer item does, and a list, a header
/// or a toolbar does not.
fn element_names_an_item(element: &IUIAutomationElement) -> bool {
    match unsafe { element.CurrentName() } {
        Ok(name) => {
            let name = name.to_string();
            !name.is_empty() && !is_container_name(&name)
        }
        Err(_) => false,
    }
}

/// The item a view has selected, for a focused element that is the view itself.
///
/// The search results view is why this exists: it can report the list as the
/// focused element while the focus it draws is on one of the results, and a list's
/// name stands for no file, so nothing downstream can be resolved from it — which
/// is why a keyboard preview in such a view found nothing at all. The list answers
/// for its own selection, and the item of that selection with the keyboard focus,
/// or the first one it holds, is the item a keyboard preview is about.
fn selected_item_of_focused_list(element: &IUIAutomationElement) -> Option<IUIAutomationElement> {
    unsafe {
        let pattern = element
            .GetCurrentPatternAs::<IUIAutomationSelectionPattern>(UIA_SelectionPatternId)
            .ok()?;
        let selection = pattern.GetCurrentSelection().ok()?;
        let count = selection.Length().ok()?;

        let mut first_named: Option<IUIAutomationElement> = None;

        for index in 0..count.min(16) {
            let Ok(candidate) = selection.GetElement(index) else {
                continue;
            };
            if !element_names_an_item(&candidate) {
                continue;
            }

            let has_keyboard_focus = candidate
                .CurrentHasKeyboardFocus()
                .map(|focused| focused.as_bool())
                .unwrap_or(false);
            if has_keyboard_focus {
                return Some(candidate);
            }

            if first_named.is_none() {
                first_named = Some(candidate);
            }
        }

        first_named
    }
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

        // A view can report the list itself as focused while the focus it draws is
        // on one of its items — the search results view does — and a list's name is
        // not a file's, so the selection is where the item has to be taken from. For
        // a focused element that already is an item this is the element itself, so
        // nothing changes for a folder view.
        let focused = if element_names_an_item(&focused) {
            focused
        } else {
            selected_item_of_focused_list(&focused)?
        };

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

        // A result whose file is not in the folder it was found under: its name
        // cannot name it, so the item itself is asked. The element's own accessible
        // value is the first place to look — see `focused_item_media_path` — and
        // the provider is asked at the item's box next.
        if let Some(path) = focused_item_media_path(&focused, &name) {
            return Some(FocusedItemInfo {
                result: AccessibilityResult::FullPath(path),
                rect,
            });
        }

        // The item's box, asked the way the pointer asks the item under the cursor:
        // the accessibility provider at the item — at the text it shows before
        // anywhere else — is where a search result states the file it stands for,
        // whatever folder that file is in. A name is only findable in a folder that
        // holds it, which is the one thing a search across folders never offers, so
        // this is the question that has to be asked, and the one hovering the same
        // result already gets an answer from.
        let probe_points = focused_item_probe_points(automation, &focused, &rect);
        if let Some(path) = focused_item_provider_path(&name, &probe_points) {
            return Some(FocusedItemInfo {
                result: AccessibilityResult::FullPath(path),
                rect,
            });
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
            // The item's own box, asked what the pointer asks the window under the
            // cursor: which item of the view sits there, and which item the view
            // itself says is focused. The provider was asked for the item already —
            // see `focused_item_probe_points` — and these are the two answers that
            // can still name a result no folder holds.
            if let Some(path) = focused_item_path_at_box(item) {
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

                // A search that began at this root returns results from the folders
                // below it as well, and those are handed over as names: the name is
                // walked for under the root rather than only looked for in it.
                if let Some(path) = lookup_media_below_folder(root, item_name) {
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

    // What the mouse path resolves the item under the pointer with: the same UI
    // Automation client, asked through a batched request, plus the view it last
    // found the pointer in.
    let mut mouse = MouseResolver::new(uia.clone());

    let mut last_file: Option<PathBuf> = None;
    let mut suppressed = SuppressedHover::default();
    let mut pointer_pause = KeyboardPointerPause::default();
    let mut hover_start: Option<Instant> = None;
    let mut last_cursor_pos = POINT::default();

    // Keyboard hover state
    let mut keyboard_file: Option<PathBuf> = None;
    let mut last_focused_key: Option<FocusedItemKey> = None;
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
    let mut suspended_initial_focus: Option<FocusedItemKey> = None;
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
    // Whether the keyboard is the one driving Explorer: set on a navigation key
    // press and kept across the keyboard previews that follow, cleared only by
    // deliberate pointer input — a move past the pointer tolerance or a wheel
    // tick — or by a reset that ends the keyboard's turn outright (previews
    // switched off, a display change, Explorer leaving the foreground). While
    // it holds, the parked pointer may neither raise a preview nor take one
    // over, which is what keeps the two previews from fighting over a list the
    // keyboard is walking: a focused item with no preview to give must not hand
    // the pointer the screen. It shows worst on a large search-result view,
    // where the items the keyboard walks are the ones still being resolved.
    let mut keyboard_screen_owner = false;
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

    let (mut config_snapshot, mut trigger_key_vk) = CONFIG
        .lock()
        .map(|c| {
            let snapshot = (
                c.preview_enabled,
                c.hover_delay_ms,
                c.trigger_key_mode,
                c.same_file_rehover_delay_ms,
            );
            // Resolved once per config change instead of once per tick.
            let vk = off_trigger_key_to_vk(&c.trigger_key);
            (snapshot, vk)
        })
        .unwrap_or(((true, 0, TriggerKeyMode::Disable, 750), Some(0x12)));
    let mut slow_explorer_probe_count = 0u32;
    let mut explorer_probe_backoff_until: Option<Instant> = None;
    let mut last_display_signature = current_display_signature();

    while RUNNING.load(Ordering::SeqCst) {
        // The pointer's answer belongs to the tick that produced it: the list under
        // a parked pointer can have moved on by the next one.
        mouse.forget_probe();

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
                mouse.forget_view();
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
                keyboard_screen_owner = false;
                hover_resolver_hints = HoverResolverHints::default();
                last_cursor_location = None;
                slow_explorer_probe_count = 0;
                explorer_probe_backoff_until =
                    Some(Instant::now() + Duration::from_millis(DISPLAY_CHANGE_BACKOFF_MS));
            } else {
                last_display_signature = Some(display_signature);
            }
        }

        // The shell answered too slowly, too often: stop asking it about the file
        // under the cursor for a moment. What the pause must not do is what it used
        // to — take the preview away and clear everything with it. At the size
        // where this fires, in a search view of several hundred results, the probes
        // are slow *every* time, so the pause re-armed itself before the next
        // preview could appear and the view looked like it had no previews at all.
        // What is on screen belongs to the file the cursor was on, and the pause
        // only stops the asking.
        if slow_explorer_probe_count >= EXPLORER_SLOW_PROBE_LIMIT
            && explorer_probe_backoff_until.is_none()
        {
            explorer_probe_backoff_until =
                Some(Instant::now() + Duration::from_millis(EXPLORER_PROBE_BACKOFF_MS));
            // The slow probe left a hover window half-open; the next probe after
            // the pause starts one again rather than inheriting it.
            stationary_search_miss_started_at = None;
            stationary_hover_probe_done = false;
        }

        if let Some(until) = explorer_probe_backoff_until {
            if Instant::now() < until {
                // The one thing still watched for is the cursor leaving the file
                // the preview is of: that would leave a preview describing a file
                // the pointer is no longer on, which is worse than no preview. A
                // pointer that stays is answered with what it already has.
                unsafe {
                    let mut cursor_pos = POINT::default();
                    if GetCursorPos(&mut cursor_pos).is_ok()
                        && ((cursor_pos.x - last_cursor_pos.x).abs() > MOUSE_MOVE_PX
                            || (cursor_pos.y - last_cursor_pos.y).abs() > MOUSE_MOVE_PX)
                    {
                        last_cursor_pos = cursor_pos;

                        if last_file.is_some() {
                            hide_preview();
                            last_file = None;
                            hover_start = None;
                            video_hover_guard_until = None;
                        }
                    }
                }

                std::thread::sleep(Duration::from_millis(MEDIUM_SLEEP_MS));
                continue;
            }

            explorer_probe_backoff_until = None;
            slow_explorer_probe_count = 0;
            // The pause is not a place a hover resumes from: whatever the cursor is
            // over when it ends has to be probed as something new.
            hover_start = Some(Instant::now());
            stationary_hover_probe_done = false;
            current_state = get_explorer_state();
            last_state_check = Instant::now();
        }

        if let Ok(config) = CONFIG.lock() {
            config_snapshot = (
                config.preview_enabled,
                config.hover_delay_ms,
                config.trigger_key_mode,
                config.same_file_rehover_delay_ms,
            );
            trigger_key_vk = off_trigger_key_to_vk(&config.trigger_key);
        }

        let preview_enabled = config_snapshot.0;
        let hover_delay_ms = config_snapshot.1;
        let trigger_key_mode = config_snapshot.2;
        let same_file_rehover_delay_ms = config_snapshot.3;

        // One question, two settings: the key either stops previews while it is
        // held, or is the only thing that lets them happen. Either way, what is left
        // to do when they are not allowed is the same as when they are turned off.
        let trigger_key_down = trigger_key_vk.is_some_and(key_is_down);
        let previews_allowed = trigger_key_mode.allows_previews(trigger_key_down);

        if !previews_allowed || !preview_enabled {
            if last_file.is_some() || keyboard_file.is_some() {
                hide_preview();
                suppressed.clear();
                pointer_pause.clear();
                stationary_search_miss_started_at = None;
                hover_start = None;
            }
            keyboard_file = None;
            last_file = None;
            last_focused_key = None;
            is_keyboard_hover = false;
            keyboard_screen_owner = false;
            video_hover_guard_until = None;
            suspend_preview_until_user_input = false;
            allow_keyboard_preview_on_first_observation = false;
            folder_change_user_initiated = false;
            last_cursor_location = None;
            hover_resolver_hints = HoverResolverHints::default();
            folder_change_time = None;
            suspended_initial_focus = None;
            // A held key has to be noticed the moment it is released, so the poll
            // stays quick while the trigger is what is holding previews back, and
            // slows down only when previews are turned off outright.
            std::thread::sleep(Duration::from_millis(if previews_allowed {
                LONG_SLEEP_MS
            } else {
                ACTIVE_POLL_MS
            }));
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
                    last_focused_key = None;
                    is_keyboard_hover = false;
                    video_hover_guard_until = None;
                    pointer_pause.clear();
                    keyboard_screen_owner = false;
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
                        last_focused_key = None;
                        is_keyboard_hover = false;
                        video_hover_guard_until = None;
                        pointer_pause.clear();
                        keyboard_screen_owner = false;
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

            let move_threshold =
                pointer_pause.move_threshold_px(is_keyboard_hover || keyboard_screen_owner);
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
                    keyboard_screen_owner = false;
                    last_focused_key = None;
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
                // The press is also the keyboard taking the screen: the pointer
                // stays parked, and nothing it sits on previews, until it takes
                // its turn back with a move or the wheel.
                keyboard_screen_owner = true;
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
                last_focused_key = None;
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
                keyboard_screen_owner = false;
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
                    suppressed.suppress(file);
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
                        // The view this location describes is a different one now —
                        // another folder, or the same window searched again — so what
                        // was cached about the last one describes the wrong place: a
                        // folder remembered for the window, an index of the items it
                        // was showing, and the view the pointer was last resolved in.
                        // A name looked up against those resolves to something that is
                        // not there, or to nothing at all.
                        clear_shell_view_probe_caches();
                        mouse.forget_view();
                        suspend_preview_until_user_input = true;
                        allow_keyboard_preview_on_first_observation = false;
                        folder_change_user_initiated = user_navigation;
                        folder_change_time = Some(Instant::now());
                        suspended_initial_focus = None;
                        hover_start = None;
                        last_focused_key = None;
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
                            let focused_key = FocusedItemKey::new(focused_name, &focused_info.rect);

                            if suspended_initial_focus.is_none() {
                                // Record the auto-focused first item
                                // (set by Windows when folder opens)
                                suspended_initial_focus = Some(focused_key);
                            } else if suspended_initial_focus.as_ref() != Some(&focused_key) {
                                // Focus actually changed — user pressed a navigation key
                                keyboard_unlocked = true;
                            }
                        }
                    }

                    if keyboard_unlocked {
                        suspend_preview_until_user_input = false;
                        allow_keyboard_preview_on_first_observation = true;
                        keyboard_screen_owner = true;
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
                keyboard_screen_owner = false;

                // Mouse movement always takes priority - dismiss keyboard hover.
                // The keyboard preview may have been covering the cursor, so the
                // file it showed is latched the way any dismissed hover is: held
                // off the mouse path for the same-file rehover delay, so the
                // handover cannot flash it straight back, and previewable again
                // after that — a pointer left sitting on the file is a user asking
                // for it.
                if is_keyboard_hover {
                    if let Some(file) = keyboard_file.clone() {
                        suppressed.suppress(file);
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
                last_focused_key = None;
                allow_keyboard_preview_on_first_observation = true;

                if let Some(suppressed_file) = suppressed.file.clone() {
                    if let Some(current_file) = get_file_under_cursor_checked(
                        &mut mouse,
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
                        &mut mouse,
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
                        suppressed.suppress(file);
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
                    // The name alone cannot tell two observations apart: a search
                    // can hold the same name in more than one folder, and the box is
                    // what says which of them the keyboard is on.
                    let focused_key = FocusedItemKey::new(focused_name, &focused_info.rect);

                    if last_focused_key.is_none() {
                        if allow_keyboard_preview_on_first_observation {
                            // The focus baseline is unknown — the mouse just moved, or
                            // the user unlocked a folder change with the keyboard — so
                            // this first observed item acts immediately instead of
                            // being recorded and waiting for a second key press.
                            last_focused_key = Some(focused_key);
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
                                    keyboard_screen_owner = true;
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
                        last_focused_key = Some(focused_key);
                    } else if last_focused_key.as_ref() != Some(&focused_key) {
                        // Focused item changed - keyboard navigation detected
                        last_focused_key = Some(focused_key);
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
                                keyboard_screen_owner = true;
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
            //
            // A pointer parked since the keyboard last drove is in the same
            // position for the same reason: the keyboard owns the screen, so a
            // file it left behind — or one it landed on that has no preview to
            // give — is not a reason for the pointer to raise a preview of
            // whatever it happens to sit on. That is the whole of the fight, and
            // it is worst in a large search-result view, where the file the
            // keyboard walks away from is still whatever is under the cursor.
            if is_keyboard_hover
                || pointer_pause.freezes_pointer()
                || text_scroll_hold
                || keyboard_screen_owner
            {
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
                                    Some(file) => suppressed.suppress(file),
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
                        &mut mouse,
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
                                    // The latch is a delay and not a verdict: the
                                    // probe is left open so the file previews as
                                    // soon as the delay has passed, rather than
                                    // leaving a pointer parked on it unanswered.
                                    stationary_hover_probe_done = false;
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
                                Some(file) => suppressed.suppress(file),
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
                                    Some(file) => suppressed.suppress(file),
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

    // Everything COM handed the loop is released while the apartment that owns it
    // is still initialized. An interface released after `CoUninitialize` belongs to
    // an apartment that has already been torn down, which faults — and the mouse
    // resolver holds the Shell window collection, the batched property request and
    // the view it last resolved in, not just the automation client.
    drop(mouse);
    drop(uia);

    unsafe {
        CoUninitialize();
    }
}

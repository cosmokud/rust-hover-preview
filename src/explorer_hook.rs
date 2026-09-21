use crate::archive_formats::matches_archive_list;
use crate::cloud_files;
use crate::config::{AvoidMode, PreviewType, TriggerKeyMode};
use crate::font_formats::matches_font_list;
use crate::image_formats::matches_image_list;
use crate::office_formats::matches_office_list;
use crate::pdf_preview::is_pdf_file;
use crate::preview_window::{
    cursor_preview_hover, hide_preview, kill_stray_video_process, monitor_dpi_from_point,
    preview_pointer_hold, preview_screen_rect, show_preview, show_preview_keyboard,
    PreviewCursorHover,
};
use crate::svg_preview;
use crate::text_formats::matches_text_lists;
use crate::video_formats::{is_video_file, matches_video_list};
use crate::wheel_input;
use crate::{CONFIG, RUNNING};
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::ffi::OsString;
use std::os::windows::ffi::{OsStrExt, OsStringExt};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Mutex;
use std::time::{Duration, Instant};
use windows::core::{w, IUnknown, Interface, VARIANT};
use windows::Win32::Foundation::{BOOL, HWND, LPARAM, POINT, RECT, SIZE};
use windows::Win32::Graphics::Gdi::{
    CreateCompatibleDC, CreateFontIndirectW, DeleteDC, DeleteObject, EnumDisplayMonitors,
    GetMonitorInfoW, GetTextExtentPoint32W, MonitorFromWindow, SelectObject, HDC, HMONITOR,
    LOGFONTW, MONITORINFO, MONITOR_DEFAULTTONEAREST,
};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoTaskMemFree, CoUninitialize, IServiceProvider, CLSCTX_ALL,
    CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED,
};
use windows::Win32::System::Variant::VT_I4;
use windows::Win32::UI::Accessibility::{
    CUIAutomation, CUIAutomation8, CUIAutomationRegistrar, IUIAutomation, IUIAutomation2,
    IUIAutomationCacheRequest, IUIAutomationElement, IUIAutomationLegacyIAccessiblePattern,
    IUIAutomationRegistrar, IUIAutomationSelectionPattern, IUIAutomationTreeWalker,
    TreeScope_Children, TreeScope_Element, UIA_BoundingRectanglePropertyId,
    UIA_ControlTypePropertyId, UIA_DataItemControlTypeId, UIA_EditControlTypeId,
    UIA_LegacyIAccessiblePatternId, UIA_ListItemControlTypeId, UIA_NamePropertyId,
    UIA_NativeWindowHandlePropertyId, UIA_SelectionPatternId, UIA_TextControlTypeId,
    UIAutomationPropertyInfo, UIAutomationType_Int, UIA_CONTROLTYPE_ID, UIA_PROPERTY_ID,
};
use windows::Win32::UI::HiDpi::{GetDpiForMonitor, GetDpiForSystem, MDT_EFFECTIVE_DPI};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetAsyncKeyState, VK_DOWN, VK_END, VK_HOME, VK_LBUTTON, VK_LEFT, VK_MBUTTON, VK_NEXT, VK_PRIOR,
    VK_RBUTTON, VK_RETURN, VK_RIGHT, VK_UP, VK_XBUTTON1, VK_XBUTTON2,
};
use windows::Win32::UI::Shell::{
    IFolderView, IFolderView2, IPersistFolder2, IShellBrowser, IShellItem, IShellView,
    IShellWindows, ItemIndex_Property_GUID, SHCreateItemFromIDList, SID_STopLevelBrowser,
    ShellWindows, SIGDN_DESKTOPABSOLUTEPARSING, SIGDN_FILESYSPATH, SIGDN_NORMALDISPLAY,
};
use windows::Win32::UI::WindowsAndMessaging::{
    EnumWindows, GetAncestor, GetClassNameW, GetCursorPos, GetForegroundWindow, GetWindowPlacement,
    GetWindowRect, GetWindowThreadProcessId, IsIconic, IsWindowVisible, SystemParametersInfoW,
    WindowFromPoint, GA_ROOT, SPI_GETICONTITLELOGFONT, SW_SHOWMAXIMIZED,
    SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS, WINDOWPLACEMENT,
};

struct ExplorerWindowCounts {
    total: usize,
    visible: usize,
}

/// What the desktop looks like from here, which is what decides whether everything
/// the hook remembers about a view — the window the pointer is in, the folder it has
/// open, the item an observation was made of — still describes anything.
///
/// Each display is in it, rather than the union of them: a display rescaled from 100%
/// to 150% rearranges everything drawn on that display and leaves the union exactly
/// where it was, and a display becoming the primary one moves where new windows open
/// and what the system's own metrics are measured against, without moving the union
/// either.
#[derive(Clone, Debug, Eq, PartialEq)]
struct DisplaySignature {
    displays: Vec<DisplayEntry>,
}

/// One display as the signature sees it: where it is, what it is scaled to, and
/// whether it is the primary one.
#[derive(Clone, Debug, Eq, PartialEq)]
struct DisplayEntry {
    rect: (i32, i32, i32, i32),
    dpi: u32,
    primary: bool,
}

/// The view under a point, as the shell describes it: the window it is drawn in,
/// the URL it was opened with, and the folder it has open.
struct ActiveShellViewContext {
    shell_view_hwnd: isize,
    location_url: Option<String>,
    folder_path: Option<String>,
}

#[derive(Clone, Default)]
struct HoverResolverHints {
    current_folder: Option<String>,
    location_url: Option<String>,
    is_search_view: bool,
    search_root: Option<String>,
    shell_view_hwnd: Option<isize>,
}

/// What both paths need to resolve an item to the file it stands for: one UI
/// Automation client whose property reads are batched into a single round trip per
/// element, Explorer's own item-position property, and the view that last answered
/// for a window.
///
/// What it holds is as telling as what it does not: there is no folder index, no
/// view index and no search root here. A file is found from the item that stands
/// for it rather than from its name, so nothing has to be walked, remembered or
/// kept warm for either path to have an answer.
struct ItemResolver {
    automation: Option<IUIAutomation>,
    /// The batched property request every element is read with, so an element
    /// costs one crossing into the view's provider rather than one per property.
    cache: Option<IUIAutomationCacheRequest>,
    walker: Option<IUIAutomationTreeWalker>,
    /// Explorer's own `ItemIndex` property, registered once per process. `None`
    /// when the registrar refuses it, which leaves the identity route skipped and
    /// the item's own value to answer.
    item_index_property: Option<UIA_PROPERTY_ID>,
    /// The Shell window collection, created once and kept: building it is the one
    /// call every lookup would otherwise repeat.
    shell_windows: Option<IShellWindows>,
    /// The view that answered for a window, kept while the window holds only that
    /// one view: a window with tabs has several and none of them is the cache's to
    /// choose between.
    view: Option<AnsweredView>,
    probe: Option<ProbeMemo>,
}

/// The view that answered for a window.
struct AnsweredView {
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

/// The text an item draws inside its own box, as the item's own children report it.
///
/// One walk reads the pieces of it for every way of avoiding that asks for one, so it
/// answers with both boxes rather than being asked twice: the whole of what the item
/// draws, and the piece its name is drawn in.
#[derive(Clone, Copy)]
struct ItemText {
    /// Every piece of the item's text taken as one box — a row's name with the columns
    /// beside it, the label under an icon — which is the region `Avoid Details` keeps a
    /// preview off. Its right edge is where the item's content stops.
    all: RECT,
    /// The piece the name is drawn in: the leftmost of them, which in the views that
    /// draw their items as rows is the `Name` column of `Details` — the name above the
    /// path of `Content` — and the label itself under an icon. It is the room the name
    /// is *given*: the region `Avoid Filename Column` keeps a preview off, and the box
    /// `Avoid Filename` narrows to the width the name itself is drawn at — see
    /// [`HoveredItem::name_box`].
    name: RECT,
}

/// The item a file is resolved from, as the view's accessibility provider reports
/// it — read at the cursor for the pointer and at the focused item for the
/// keyboard, so both paths get their answer from the same facts.
struct HoveredItem {
    /// The item's position in the view, one-based as Explorer's own `ItemIndex`
    /// reports it — the one fact about a search result that a shared name cannot
    /// take away, because two results may share a name and only one of them is
    /// at this position.
    index: Option<i32>,
    /// The name the item goes by in the view, which may hide the extension.
    name: String,
    /// The item's legacy accessible value: for a file a search has surfaced this
    /// is normally the file's own path, which is the second way an item is
    /// answered when it reports no position.
    value: Option<String>,
    /// The box the item occupies on screen, which is what says the pointer is
    /// inside it and where a keyboard preview is placed.
    bounds: RECT,
    /// The text the item draws — its name, and the columns a view that draws its
    /// items as rows writes beside it — or `None` when the view reported no text, or
    /// the walk that would have read one was not asked for it, which is what a walk
    /// with nothing to keep a preview off asks. What it holds is the region the
    /// `Avoid` setting places a preview from, the keyboard's as much as the pointer's.
    text: Option<ItemText>,
    /// The window the item is drawn in, whose frame is the window the item's view
    /// belongs to.
    native_window: isize,
}

impl HoveredItem {
    /// Whether a second look found the same item. What the view says about an item
    /// is only true of the item — a list can move under a parked pointer — so an
    /// answer is only taken when both looks agree.
    fn same_item(&self, other: &HoveredItem) -> bool {
        self.index == other.index && self.name == other.name
    }

    /// The region a preview of this item is kept off, as the `Avoid` setting has it:
    /// the name where it is drawn at `Filename`, the box the view gives the name at
    /// `FilenameColumn`, that box with the columns a row writes beside it at `Details`,
    /// and nothing at all at `Off` — or the item's own box at any of the first three
    /// for a view that reports no text, the name being drawn inside that box whatever
    /// the view says about it, so an item whose text cannot be measured is avoided as
    /// the whole of itself.
    fn avoid_box(&self) -> Option<(i32, i32, i32, i32)> {
        let region = match avoid_mode() {
            AvoidMode::Off => return None,
            AvoidMode::Filename => self.name_box(),
            AvoidMode::FilenameColumn => self.text.map(|text| text.name),
            AvoidMode::Details => self.text.map(|text| text.all),
        }
        .unwrap_or(self.bounds);

        Some((region.left, region.top, region.right, region.bottom))
    }

    /// The box the item's name is drawn in, cut to the width the name itself takes:
    /// the region `Avoid Filename` keeps a preview off, where the box as the view
    /// reported it is the one `Avoid Filename Column` keeps it off.
    ///
    /// A view reports the room an item's name is *given* — the `Name` column of a
    /// `Details` row is one width for every file in it, a long name and a short one
    /// alike — so the width the name is drawn at, at the larger of the two sizes a view
    /// may draw it at, is measured and the box is narrowed to it. It is only ever
    /// narrowed: a name that fills the room it was given, or is drawn truncated to it,
    /// is left as the view reported it.
    fn name_box(&self) -> Option<RECT> {
        let text = self.text?;

        let Some(width) = drawn_name_width(&self.name, region_display_dpi(&text.name)) else {
            return Some(text.name);
        };

        Some(RECT {
            right: (text.name.left + width).min(text.name.right),
            ..text.name
        })
    }
}

/// The display a region is drawn on, as the DPI its pixels are in.
///
/// The window an item is reported through is not a display that can always be asked:
/// the shell's own item provider answers with no window at all, and a display taken
/// from the system's instead is the wrong one for an item on a scaled display — a name
/// drawn at 200% would be measured half its width. Where the region is, is what says
/// which display it is drawn on, so the middle of it is what is looked up.
fn region_display_dpi(region: &RECT) -> u32 {
    monitor_dpi_from_point(
        (region.left + region.right) / 2,
        (region.top + region.bottom) / 2,
    )
}

/// The share of the icon font the shell draws an item's name at in the views that give
/// the name a line of its own: `Content` draws the name at 125% of the icon font, where
/// a `Details` row is drawn at the font itself. A name is measured at the larger of the
/// two so that one region clears the name in either view — which costs a `Details` row
/// a quarter more room than its name takes, and is what keeps a `Content` row's name,
/// extension and all, from being covered.
///
/// The share was read off the pixels: at 100% scaling, `aa.txt`,
/// `mid-length-name.txt` and `a-very-long-file-name-here.txt` are drawn 36, 136 and 202
/// pixels wide in `Content` against 28, 110 and 162 in `Details`.
const NAME_FONT_SCALE: (i64, i64) = (5, 4);

/// The width a name is drawn at, in the pixels of the display the item is on, or
/// `None` when it cannot be measured — which leaves the name as the view reported it.
///
/// The shell draws an item's name in the icon font, the one folder views are given, at
/// the size the view gives it — see `NAME_FONT_SCALE` — so that is what a name is
/// measured with: a font and a memory DC are made for the one call and let go again. It
/// is asked beside the walk that read the item — once per preview, not once per probe —
/// and the font is the shell's own, written for the display the system is at, so it is
/// scaled here to the display the item is drawn on — `dpi`, the one the region is in
/// (see [`region_display_dpi`]) — which is the scale the drawn name is in.
///
/// A display that is not known is not measured at another one's scale: the name is left
/// as the view reported it, the answer a name with nothing in it gets, because a region
/// that clears too much room is one a preview is placed too far away, while a region cut
/// at the wrong scale is one that covers the name it was cut from.
fn drawn_name_width(name: &str, dpi: u32) -> Option<i32> {
    let wide: Vec<u16> = name.encode_utf16().collect();
    if wide.is_empty() || dpi == 0 {
        return None;
    }

    unsafe {
        let mut logfont = LOGFONTW::default();
        SystemParametersInfoW(
            SPI_GETICONTITLELOGFONT,
            0,
            Some(&mut logfont as *mut LOGFONTW as *mut core::ffi::c_void),
            SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS(0),
        )
        .ok()?;

        let system_dpi = GetDpiForSystem();
        if system_dpi != 0 && dpi != system_dpi {
            let height = logfont.lfHeight as i64 * dpi as i64;
            logfont.lfHeight = (height / system_dpi as i64) as i32;
        }

        // The name is drawn at one of two sizes and the region has to clear whichever
        // it is, so the larger is what is measured — see `NAME_FONT_SCALE`.
        let height = logfont.lfHeight as i64 * NAME_FONT_SCALE.0;
        logfont.lfHeight = (height / NAME_FONT_SCALE.1) as i32;

        let font = CreateFontIndirectW(&logfont);
        if font.0.is_null() {
            return None;
        }

        let dc = CreateCompatibleDC(None);
        if dc.0.is_null() {
            let _ = DeleteObject(font);
            return None;
        }

        let previous = SelectObject(dc, font);
        let mut extent = SIZE::default();
        let measured = GetTextExtentPoint32W(dc, &wide, &mut extent).as_bool();
        let _ = SelectObject(dc, previous);
        let _ = DeleteObject(font);
        let _ = DeleteDC(dc);

        (measured && extent.cx > 0).then_some(extent.cx)
    }
}

impl ItemResolver {
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

        let mut resolver = Self {
            automation,
            cache,
            walker,
            item_index_property,
            shell_windows: None,
            view: None,
            probe: None,
        };
        resolver.rebuild_shell();
        resolver
    }

    /// Build the Shell window collection again, and drop the view that answered
    /// through it. These two are what the resolver holds that another process
    /// serves: the collection and every view are Explorer's own, so once Explorer
    /// is not the process it was, both are proxies into one that is gone — and a
    /// proxy into a gone process fails for good rather than reconnecting.
    fn rebuild_shell(&mut self) {
        self.shell_windows =
            unsafe { CoCreateInstance::<_, IShellWindows>(&ShellWindows, None, CLSCTX_ALL).ok() };
        self.view = None;
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

impl AnsweredView {
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

/// How long a UI Automation call may take before it is abandoned. Every probe
/// crosses into the shell's own thread, so a shell that has stopped answering
/// would otherwise hold this thread inside a probe for as long as it likes — and
/// the slow-probe backoff cannot see a probe that has not returned yet, which
/// leaves an unbounded wait with nothing watching it.
const UIA_TIMEOUT_MS: u32 = 500;

/// The UI Automation client both paths resolve items with, with every call
/// bounded.
///
/// `CUIAutomation8` is the client that carries `IUIAutomation2`, which is where
/// the timeouts live: the legacy `CUIAutomation` object does not answer for that
/// interface at all, so asking it to be bounded is what failed — and, when that
/// was treated as fatal, what left every hover without a preview. The legacy
/// client is still the fallback, and a client that cannot be bounded is still
/// used, because a client that cannot be bounded still resolves items: what it
/// costs is the wait the timeouts were meant to remove, which is the behavior the
/// app had before, and not a dead app.
fn automation_client() -> Option<IUIAutomation> {
    let automation: IUIAutomation = unsafe {
        match CoCreateInstance(&CUIAutomation8, None, CLSCTX_ALL) {
            Ok(automation) => automation,
            Err(_) => CoCreateInstance(&CUIAutomation, None, CLSCTX_ALL).ok()?,
        }
    };

    if let Ok(bounded) = automation.cast::<IUIAutomation2>() {
        unsafe {
            let _ = bounded.SetConnectionTimeout(UIA_TIMEOUT_MS);
            let _ = bounded.SetTransactionTimeout(UIA_TIMEOUT_MS);
        }
    }

    Some(automation)
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
    fn move_threshold_px(&self, keyboard_owns_screen: bool, dpi: u32) -> i32 {
        let logical_pixels = if keyboard_owns_screen || self.freezes_pointer() {
            KEYBOARD_POINTER_MOVE_TOLERANCE_PIXELS
        } else {
            MOUSE_MOVE_PIXELS
        };

        (logical_pixels * dpi as f32 / 96.0).round() as i32
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

const EXPLORER_PROBE_SLOW_MS: u64 = 700;
const EXPLORER_WINDOW_CACHE_TTL_MS: u64 = 1000;
const EXPLORER_REAL_FOLDER_CACHE_MAX_ENTRIES: usize = 256;
const FOLDER_PROBE_MS: u64 = 200;
const IDLE_FOLDER_PROBE_MS: u64 = 750;
const FOLDER_PROBE_TRIGGER_MS: u64 = 400;
const DISPLAY_CHANGE_BACKOFF_MS: u64 = 1500;
/// How often the desktop is walked and compared with what it was: the displays a
/// machine has, each one's scale, and which of them is the primary one (see
/// `current_display_signature`). A display change is acted on within this of it
/// happening, which is a fifth of a second against the backoff the rebuild waits out
/// anyway — and walking every display on every tick would be the most expensive thing
/// in a loop that runs two or three dozen times a second.
const DISPLAY_CHECK_MS: u64 = 200;
const KEYBOARD_FOCUS_INPUT_GRACE_MS: u64 = 500;
const HOVER_RESOLVER_INPUT_GRACE_MS: u64 = 1500;
const WHEEL_SCROLL_SETTLE_MS: u64 = 150;
/// How far the pointer has to move before it counts as moved at all, in logical
/// pixels — and the wider distance that counts while the keyboard owns the screen, so
/// a pointer resting on a desk cannot cancel a keyboard preview. Logical distances,
/// scaled by the display the pointer is on: a hand moves the same distance whatever
/// the display it is over is scaled to.
const MOUSE_MOVE_PIXELS: f32 = 5.0;
const KEYBOARD_POINTER_MOVE_TOLERANCE_PIXELS: f32 = 20.0;
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

static EXPLORER_LAST_REAL_FOLDERS: Lazy<Mutex<HashMap<isize, String>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));
static EXPLORER_WINDOW_CACHE: Lazy<Mutex<HashMap<isize, (bool, Instant)>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// Explorer restarts seen, counted rather than flagged: the hook loop is the one
/// that has to notice one, since it is the thread holding what a restart
/// invalidates, and a count cannot be missed by a reset that would have landed
/// between two of its ticks. Bumped by the tray thread from the `TaskbarCreated`
/// broadcast the shell sends when its taskbar is built again.
static EXPLORER_RESTARTS: AtomicU64 = AtomicU64::new(0);

/// How long a Shell collection that could not be built is left alone before it is
/// built again. Building it fails while the shell is still coming up — at logon,
/// and in the moment after a restart — and a collection left missing answers every
/// later lookup the way a dead one does.
const SHELL_COLLECTION_RETRY_MS: u64 = 1000;

/// How long previews wait after Explorer has restarted: every window the pointer
/// could be over is the new shell's, and it is still putting them up.
const EXPLORER_RESTART_BACKOFF_MS: u64 = 1500;

/// Note that Explorer is not the process it was. Every Shell object this app holds
/// is served by explorer.exe, and the ones taken before a restart are proxies into
/// the process that is gone. Called on the `TaskbarCreated` broadcast, which
/// reaches every top-level window when the shell's taskbar is created again.
pub fn note_explorer_restart() {
    EXPLORER_RESTARTS.fetch_add(1, Ordering::SeqCst);
}

/// Explorer restarts seen so far; the tray thread is the only writer.
fn explorer_restart_count() -> u64 {
    EXPLORER_RESTARTS.load(Ordering::SeqCst)
}

/// Drop what describes a view: the folder the last probe remembered for a window,
/// and whether a window is one of Explorer's. The window the pointer is in is
/// resolved again the next time it is asked about.
fn clear_shell_view_probe_caches() {
    if let Ok(mut cache) = EXPLORER_WINDOW_CACHE.lock() {
        cache.clear();
    }
    if let Ok(mut cache) = EXPLORER_LAST_REAL_FOLDERS.lock() {
        cache.clear();
    }
}

/// `MONITORINFOF_PRIMARY`, which the Windows bindings do not name: the flag
/// `GetMonitorInfo` sets on the display new windows open on.
const MONITORINFOF_PRIMARY_FLAG: u32 = 0x1;

/// One display, as `EnumDisplayMonitors` walks them, added to the list the signature
/// is being built from. The list is what the caller's `LPARAM` points at.
unsafe extern "system" fn collect_display(
    monitor: HMONITOR,
    _dc: HDC,
    _rect: *mut RECT,
    data: LPARAM,
) -> BOOL {
    let displays = &mut *(data.0 as *mut Vec<DisplayEntry>);

    let mut info = MONITORINFO {
        cbSize: std::mem::size_of::<MONITORINFO>() as u32,
        ..Default::default()
    };
    if !GetMonitorInfoW(monitor, &mut info).as_bool() {
        return BOOL(1);
    }

    let mut dpi_x = 0u32;
    let mut dpi_y = 0u32;
    let dpi = if GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, &mut dpi_x, &mut dpi_y).is_ok() {
        dpi_x
    } else {
        0
    };
    let display = info.rcMonitor;

    displays.push(DisplayEntry {
        rect: (display.left, display.top, display.right, display.bottom),
        dpi,
        primary: info.dwFlags & MONITORINFOF_PRIMARY_FLAG != 0,
    });

    BOOL(1)
}

fn current_display_signature() -> Option<DisplaySignature> {
    let mut displays: Vec<DisplayEntry> = Vec::new();

    unsafe {
        let data = LPARAM(&mut displays as *mut Vec<DisplayEntry> as isize);
        if !EnumDisplayMonitors(None, None, Some(collect_display), data).as_bool() {
            return None;
        }
    }

    (!displays.is_empty()).then_some(DisplaySignature { displays })
}

fn display_signature_changed(
    previous: Option<&DisplaySignature>,
    current: &DisplaySignature,
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

/// Whether a preview may be shown for `path`: the kind of preview it would get,
/// and whether that kind is switched on in the tray's `Preview Types`
/// submenu.
///
/// The kinds are asked the way the renderer asks them — a video first, then a
/// PDF, then the archive list, then the office list, then the text lists, then
/// the image list — so the two cannot disagree about what a file is. A video goes
/// first because only its content settles the extensions it shares with text: a
/// `.ts` carrying MPEG-TS packets is a video however the gates stand, and one that
/// does not is the TypeScript source the text lists claim. The lists come from the
/// configuration this already holds rather than from the gates' own lookups, which
/// would take the same lock again.
fn is_media_file(path: &Path) -> bool {
    let Ok(config) = CONFIG.lock() else {
        return false;
    };

    if matches_video_list(path, &config.video_extensions) {
        return PreviewType::Videos.enabled_in(&config);
    }
    if is_pdf_file(path) {
        return PreviewType::Pdf.enabled_in(&config);
    }
    if matches_archive_list(path, &config.archive_extensions) {
        return PreviewType::Archives.enabled_in(&config);
    }
    if matches_office_list(path, &config.office_extensions) {
        return PreviewType::Office.enabled_in(&config);
    }
    if matches_text_lists(path, &config.text_extensions, &config.text_names) {
        return PreviewType::Text.enabled_in(&config);
    }

    // A font is its own kind rather than an entry of the image list — what draws one is the
    // browser engine, the way a document's is — so it is asked here, ahead of the list that
    // would have turned a `.ttf` down: the renderer asks the same question in the same
    // place, so the two cannot disagree about which gate a file is under.
    if matches_font_list(path, &config.font_extensions) {
        return PreviewType::Fonts.enabled_in(&config);
    }

    if !matches_image_list(path, &config.image_extensions) {
        return false;
    }

    // A document is its own kind even though its name is an entry of the image list:
    // what draws one is not a decoder, and the switch for it is not the switch for
    // pictures. Asked here the way the renderer asks it — the list first, then the
    // document — so the two cannot disagree about which gate a file is under.
    if svg_preview::is_svg_file(path) {
        return PreviewType::Svg.enabled_in(&config);
    }

    PreviewType::Images.enabled_in(&config)
}

/// How far a preview is placed clear of the item it is about, as the tray's `Avoid`
/// submenu and `config.ini` have it.
fn avoid_mode() -> AvoidMode {
    CONFIG
        .lock()
        .map(|config| config.avoid_mode)
        .unwrap_or(AvoidMode::Off)
}

fn same_path(a: &PathBuf, b: &PathBuf) -> bool {
    a == b
        || a.as_os_str()
            .encode_wide()
            .map(ascii_lower)
            .eq(b.as_os_str().encode_wide().map(ascii_lower))
}

/// Fold one UTF-16 unit the way `eq_ignore_ascii_case` folds a character, so a
/// path can be compared without being turned into a string first.
fn ascii_lower(unit: u16) -> u16 {
    if (b'A' as u16..=b'Z' as u16).contains(&unit) {
        unit + 32
    } else {
        unit
    }
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

    get_cached_explorer_real_folder(context.shell_view_hwnd)
}

fn is_probable_search_view_context(context: &ActiveShellViewContext) -> bool {
    context
        .location_url
        .as_deref()
        .map(is_search_ms_url)
        .unwrap_or(false)
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

            let context = ActiveShellViewContext {
                shell_view_hwnd: shell_view_hwnd_key,
                location_url,
                folder_path,
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

/// Whether the view under the pointer is a search's results. The hints carry the
/// same answer, but they are read at the folder probe's cadence: a search opened a
/// moment ago is seen here before it is seen there.
fn is_current_search_view_legacy() -> bool {
    match get_active_shell_view_context_at_cursor() {
        Some(context) => context
            .location_url
            .as_deref()
            .map(is_search_ms_url)
            .unwrap_or(false),
        None => false,
    }
}

/// What the view under the pointer is showing, for the probes that need to know a
/// location has changed.
///
/// It is answered by the view itself — the folder it has open and the URL it was
/// opened with — and by nothing else: the resolution of a file does not depend on
/// it, so a view the shell does not describe leaves the hints empty rather than
/// sending the hook looking for another witness.
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
    if !path.exists() || !is_media_file(&path) || cloud_files::needs_download(&path) {
        return None;
    }

    normalize_existing_path(path)
}

/// The path a Shell item stands for, in the form the rest of the app works in.
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

/// The item the pointer is over, as the view's accessibility provider reports it,
/// or nothing when the pointer is over no item at all.
///
/// `measure_content` asks the walk for the item's own text as well, which a preview
/// is placed from. A plain probe passes `false`: it answers which file the pointer is
/// on, and a preview that keeps off that file's item asks for the text when it is
/// about to be shown — see `avoid_box_under_cursor`.
fn uia_item_from_point(
    resolver: &ItemResolver,
    point: POINT,
    measure_content: bool,
) -> Option<HoveredItem> {
    let automation = resolver.automation.as_ref()?;
    let cache = resolver.cache.as_ref()?;
    let element = unsafe { automation.ElementFromPointBuildCache(point, cache) }.ok()?;

    walk_to_item(resolver, &element, Some(point), measure_content)
}

/// The item the keyboard is on: the element Explorer says holds the focus, or the
/// item of that element's list when the view reports the list itself as focused.
fn uia_item_from_focus(resolver: &ItemResolver) -> Option<HoveredItem> {
    let automation = resolver.automation.as_ref()?;
    let cache = resolver.cache.as_ref()?;
    let focused = unsafe { automation.GetFocusedElementBuildCache(cache) }.ok()?;

    // A view can report the list itself as focused while the focus it draws is on
    // one of its items — the search results view does — and a list's name is not a
    // file's, so the selection is where the item has to be taken from. For a
    // focused element that already is an item this is the element itself.
    let start = if element_names_an_item(&focused) {
        focused
    } else {
        selected_item_of_focused_list(&focused)?
    };

    // A keyboard preview is placed from the item alone, so the text the region is made
    // of is read with it — see `item_text_box`.
    walk_to_item(resolver, &start, None, true)
}

/// The nearest item at or above an element, as the view reports it.
///
/// The walk starts where the caller's evidence does — the element under the
/// pointer, or the focused element — and goes up to the first list row or data
/// item, because what the view says about an item is kept on the item and not on
/// the text it draws inside it.
fn walk_to_item(
    resolver: &ItemResolver,
    start: &IUIAutomationElement,
    point: Option<POINT>,
    measure_content: bool,
) -> Option<HoveredItem> {
    let cache = resolver.cache.as_ref()?;
    let walker = resolver.walker.as_ref()?;
    let mut element = start.clone();

    for _ in 0..=POINTER_ITEM_ANCESTOR_LIMIT {
        if let Some(item) = item_from_element(resolver, &element, point, measure_content) {
            return Some(item);
        }

        element = unsafe { walker.GetParentElementBuildCache(&element, cache) }.ok()?;
    }

    None
}

/// The item an element is, when the pointer's box test says it is the one asked
/// about — or, for the keyboard, whenever it is an item at all.
///
/// The box is what says the pointer is on an item: a row is an item from its left
/// edge to its right one, and a pointer anywhere on that row belongs to the file
/// the row stands for, which is the same thing Explorer highlights when it is
/// selected. An item whose box does not hold the point is not an answer even
/// though the walk passed through it, and since the walk starts at the element
/// under the pointer, the first item that holds the point is the one it is on. A
/// keyboard preview has no point to test — the focused element is the evidence —
/// so there it is the item itself that answers.
fn item_from_element(
    resolver: &ItemResolver,
    element: &IUIAutomationElement,
    point: Option<POINT>,
    measure_content: bool,
) -> Option<HoveredItem> {
    let control_type = element_control_type(element)?;
    if control_type != UIA_ListItemControlTypeId && control_type != UIA_DataItemControlTypeId {
        return None;
    }

    let bounds = element_bounds(element)?;
    if bounds.left >= bounds.right || bounds.top >= bounds.bottom {
        return None;
    }
    if let Some(point) = point {
        let holds_point = point.x >= bounds.left
            && point.x < bounds.right
            && point.y >= bounds.top
            && point.y < bounds.bottom;
        if !holds_point {
            return None;
        }
    }

    // The item's own text is read only where it can answer, and only when the caller
    // wants it: what the view draws is what a way of avoiding is measured from, and a
    // walk with nothing to keep a preview off has no region to read. See
    // `item_text_box`.
    let text = (measure_content && avoid_mode() != AvoidMode::Off)
        .then(|| item_text_box(resolver, element, &bounds))
        .flatten();

    Some(HoveredItem {
        index: element_item_index(element, resolver.item_index_property),
        name: element_name(element).unwrap_or_default().trim().to_string(),
        value: element_value(element),
        bounds,
        text,
        native_window: element_native_window(element),
    })
}

/// The text an item draws inside its own box, or `None` when it draws none.
///
/// A view gives every item the box it occupies, and what it draws inside that box
/// is reported the way a view reports everything: each piece of an item's text — a
/// Content row's name and path, a Details row's name, type, modified date and size,
/// the label under an icon — is an element of its own carrying the box it is drawn
/// in, child of the item. Reading them answers two things a box cannot: the whole of
/// what the item draws, whose right edge is where a row's content stops — the room a
/// keyboard preview is placed in — and the leftmost piece, which in the views that
/// draw their items as rows is the `Name` column of `Details` or the name above the
/// path of `Content`, and which is the region a way of avoiding is measured from.
///
/// It is measured from the item's own children for the same reason: a view reports
/// what it draws as children, and the rightmost of them is the edge the row's content
/// stops at. The read is batched into one round trip with the properties the rest of
/// the walk already asks for. A view that draws its columns some other way reports no
/// text at all, and is answered with `None`, which leaves the item measured by its box
/// — see [`HoveredItem::avoid_box`].
fn item_text_box(
    resolver: &ItemResolver,
    element: &IUIAutomationElement,
    bounds: &RECT,
) -> Option<ItemText> {
    let automation = resolver.automation.as_ref()?;
    let cache = resolver.cache.as_ref()?;

    unsafe {
        let condition = automation.CreateTrueCondition().ok()?;
        let children = element
            .FindAllBuildCache(TreeScope_Children, &condition, cache)
            .ok()?;
        let count = children.Length().ok()?;

        let mut all: Option<RECT> = None;
        let mut name: Option<RECT> = None;
        for index in 0..count {
            let Ok(child) = children.GetElement(index) else {
                continue;
            };
            if !element_is_drawn_text(&child) {
                continue;
            }
            let Ok(rect) = child.CachedBoundingRectangle() else {
                continue;
            };
            if rect.right <= rect.left {
                continue;
            }
            // The name is the leftmost piece, which is the one the views that draw
            // their items as rows put first; two pieces drawn from the same edge — the
            // name above the path of `Content` — are told apart by taking the higher.
            let is_name = match name {
                None => true,
                Some(current) => {
                    rect.left < current.left
                        || (rect.left == current.left && rect.top < current.top)
                }
            };
            if is_name {
                name = Some(rect);
            }
            all = Some(match all {
                Some(union) => RECT {
                    left: union.left.min(rect.left),
                    top: union.top.min(rect.top),
                    right: union.right.max(rect.right),
                    bottom: union.bottom.max(rect.bottom),
                },
                None => rect,
            });
        }

        // Text reported outside the item's box is reported wrong, and the box is what
        // answers for an item whose text the view does not place.
        let inside = |rect: &RECT| rect.right > bounds.left && rect.left < bounds.right;
        let all = all.filter(inside)?;

        Some(ItemText {
            all,
            name: name.filter(inside).unwrap_or(all),
        })
    }
}

/// Whether an element is text the view draws, which is what an item's own content
/// is made of. A row of a file list reports its name and its columns that way, and
/// anything else it may report — the file's icon, the row's own container — is not
/// part of the text whose end is being measured.
fn element_is_drawn_text(element: &IUIAutomationElement) -> bool {
    match element_control_type(element) {
        Some(control_type) => {
            control_type == UIA_EditControlTypeId || control_type == UIA_TextControlTypeId
        }
        None => false,
    }
}

/// The control type of an element, from the batched cache when the element came
/// from one and from the provider itself when it did not: the item a keyboard
/// preview is resolved from can be handed over by a selection pattern, which is
/// not a cached read of the element under a pointer.
fn element_control_type(element: &IUIAutomationElement) -> Option<UIA_CONTROLTYPE_ID> {
    unsafe {
        element
            .CachedControlType()
            .or_else(|_| element.CurrentControlType())
            .ok()
    }
}

fn element_name(element: &IUIAutomationElement) -> Option<String> {
    unsafe {
        element
            .CachedName()
            .or_else(|_| element.CurrentName())
            .ok()
            .map(|name| name.to_string())
    }
}

fn element_bounds(element: &IUIAutomationElement) -> Option<RECT> {
    unsafe {
        element
            .CachedBoundingRectangle()
            .or_else(|_| element.CurrentBoundingRectangle())
            .ok()
    }
}

fn element_native_window(element: &IUIAutomationElement) -> isize {
    unsafe {
        element
            .CachedNativeWindowHandle()
            .or_else(|_| element.CurrentNativeWindowHandle())
            .map(|window| window.0 as isize)
            .unwrap_or_default()
    }
}

/// The value the item's legacy accessible pattern carries — for a file a search
/// has surfaced, the file's own path.
fn element_value(element: &IUIAutomationElement) -> Option<String> {
    unsafe {
        let pattern = element
            .GetCachedPatternAs::<IUIAutomationLegacyIAccessiblePattern>(
                UIA_LegacyIAccessiblePatternId,
            )
            .or_else(|_| {
                element.GetCurrentPatternAs::<IUIAutomationLegacyIAccessiblePattern>(
                    UIA_LegacyIAccessiblePatternId,
                )
            })
            .ok()?;

        pattern
            .CachedValue()
            .or_else(|_| pattern.CurrentValue())
            .ok()
            .map(|value| value.to_string())
            .filter(|value| !value.trim().is_empty())
    }
}

/// The position the view holds an element at, as Explorer reports it.
///
/// The property is one-based, and zero is what the provider answers when it has
/// nothing to say about the element's position at all — so only a positive value
/// is a position. A missing one is not a failure: the item's own value still
/// stands on its own.
fn element_item_index(
    element: &IUIAutomationElement,
    item_index_property: Option<UIA_PROPERTY_ID>,
) -> Option<i32> {
    let property = item_index_property?;

    unsafe {
        let value = element
            .GetCachedPropertyValue(property)
            .or_else(|_| element.GetCurrentPropertyValue(property))
            .ok()?;
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
fn folder_views_for_window(resolver: &ItemResolver, root_key: isize) -> Vec<AnsweredView> {
    let mut candidates: Vec<AnsweredView> = Vec::new();
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

            candidates.push(AnsweredView {
                browser_hwnd: root_key,
                shell_browser,
                view_identity,
                folder_view,
            });
        }
    }

    candidates
}

/// How many Shell windows are registered, which is what says whether a window
/// could be holding tabs.
fn shell_window_count(resolver: &ItemResolver) -> Option<i32> {
    unsafe { resolver.shell_windows.as_ref()?.Count().ok() }
}

/// The file an item stands for, asked of the view that is showing it.
///
/// The view belongs to a window — the frame the pointer is over, or the one the
/// focused item is drawn in — and a window that holds tabs registers one Shell
/// window per tab, all of them answering with the frame's own window, so the frame
/// names a set of views and not one. Which of them it is, is settled in two steps
/// that must not be confused with each other.
///
/// First the item: a candidate view is asked whether the item at that position is
/// the item we are on, by name and nothing else. That is an identity question, and
/// a folder answers it exactly as a file does — what the item *is* says nothing
/// about which view holds it. Then, and only for the views that claimed the item,
/// the file: the path the Shell hands over, gated to a file this app previews. A
/// view that holds the item but has no file to show it (a folder, an archive, a
/// document) is a *match* with nothing to preview, not a view that failed to match
/// — treating it as the latter is how another tab's file gets shown while a folder
/// is hovered. What several matches do has to agree: two tabs showing the same
/// folder are one answer, while tabs that disagree — about the file, or about
/// whether there is one at all — are a question the item cannot settle, and no
/// answer is better than the wrong tab's file.
fn item_file_path(
    resolver: &mut ItemResolver,
    root_key: isize,
    item: &HoveredItem,
) -> Option<PathBuf> {
    let index = item.index? - 1;
    let registrations = shell_window_count(resolver);

    // An item with no name cannot be told from another tab's item, so only a window
    // showing a single view can be answered without one.
    if item.name.is_empty() && registrations != Some(1) {
        return None;
    }

    // The view that answered last is asked first, but only while it is the *only*
    // registration for the window: with one view there is no second one for it to
    // disagree with, and a window holding tabs is never answered from the cache —
    // a cached tab is one of several, and the cache cannot say which of them is
    // showing.
    if registrations == Some(1) {
        if let Some(answered) = resolver.view.as_ref() {
            if answered.browser_hwnd == root_key && answered.is_current() {
                if let Some(path) = view_item_media_path(&answered.folder_view, index) {
                    return Some(path);
                }
            }
        }
    }

    let mut matches = 0usize;
    let mut matched_without_file = 0usize;
    let mut answer: Option<(AnsweredView, PathBuf)> = None;
    let mut disagreed = false;

    for candidate in folder_views_for_window(resolver, root_key) {
        if !view_item_holds(&candidate.folder_view, index, &item.name) {
            continue;
        }
        matches += 1;

        let Some(path) = view_item_media_path(&candidate.folder_view, index) else {
            // This view holds the item, and the item is not a file to preview.
            matched_without_file += 1;
            continue;
        };

        match &answer {
            None => answer = Some((candidate, path)),
            Some((_, existing)) if !same_path(existing, &path) => disagreed = true,
            // Another tab showing the same folder is the same answer.
            Some(_) => {}
        }
    }

    if matches == 0 {
        return None;
    }

    // Tabs that do not agree about what the item is — one holding a file, another
    // holding a folder, a name that is in both at that position — cannot be told
    // apart, and either answer could be the wrong tab's.
    if matched_without_file > 0 && (matches > 1 || disagreed) {
        return None;
    }
    if disagreed {
        return None;
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

/// Whether the item at a position in a view is the item that was asked about.
///
/// This is the identity question and nothing else: the name the view shows the item
/// under against the name the accessibility tree reports for it. What the item *is*
/// is not part of it — a folder goes by its name exactly as a file does, and a view
/// that holds a folder has to be seen as holding the item, or the tab that owns it
/// abstains and another tab's file answers in its place. An item with no name to ask
/// about is taken as held, which the caller has already established can only be
/// asked of a window showing one view.
fn view_item_holds(folder_view: &IFolderView2, index: i32, expected_name: &str) -> bool {
    if index < 0 || expected_name.is_empty() {
        return true;
    }

    unsafe {
        let Ok(item) = folder_view.GetItem::<IShellItem>(index) else {
            return false;
        };

        item_display_name_matches(&item, expected_name)
    }
}

/// The path a view's item at a position stands for, when it is a file this app
/// previews.
///
/// The path is the Shell's own answer for the item at that position
/// (`SIGDN_FILESYSPATH`) — for a search result, the real file wherever it lives,
/// and for a folder or an item that stands for no file at all, no answer. It is the
/// *second* question, asked only of the views that first claimed the item: a folder
/// reaches this point and stops here, which is what keeps the preview of a folder
/// from being another tab's file.
fn view_item_media_path(folder_view: &IFolderView2, index: i32) -> Option<PathBuf> {
    if index < 0 {
        return None;
    }

    unsafe {
        let item = folder_view.GetItem::<IShellItem>(index).ok()?;
        let path = shell_item_filesystem_path(&item)?;

        normalize_media_path(path).filter(|path| path.is_file())
    }
}

/// The file the pointer is over, resolved once per point.
///
/// The pointer asks one question — what is under me — and the view under it
/// answers by identity: the item the accessibility provider says the point is
/// inside, the position that item holds in the view, and the file that position
/// stands for. Nothing is looked up by name for the pointer, because a search
/// across folders is full of names that belong to more than one file and a name
/// is the one thing the view does not need. What follows the identity route is the
/// same answer asked of the item itself: the accessible value it carries, when
/// that value is a whole path.
fn get_file_under_cursor(resolver: &mut ItemResolver) -> Option<PathBuf> {
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

    let answer = resolve_file_under_cursor(resolver, point);
    resolver.remember_probe(point, answer.clone());
    answer
}

/// The file the pointer is over.
///
/// Two witnesses and no third: the item the pointer is on, turned into a file by
/// the view that is showing it, and the item's own accessible value when that
/// value is a whole path. Nothing is looked up by name, nothing is walked, and a
/// name that nothing can vouch for is left unanswered rather than guessed at — a
/// search across folders is full of names that belong to more than one file.
fn resolve_file_under_cursor(resolver: &mut ItemResolver, point: POINT) -> Option<PathBuf> {
    let item = uia_item_from_point(resolver, point, false)?;

    if let Some(root_key) = root_window_at(point).map(|window| window.0 as isize) {
        if let Some(path) = item_file_path(resolver, root_key, &item) {
            // The view is asked twice: a wheel turns the list under a parked
            // pointer, and an item that is no longer at the point the first answer
            // described is not what that answer is about.
            if uia_item_from_point(resolver, point, false)
                .map(|again| again.same_item(&item))
                .unwrap_or(false)
            {
                return Some(path);
            }
        }
    }

    // What the item says about itself: a search result carries the file's own path
    // in its accessible value, and a name that is a whole path was come by the same
    // way.
    if let Some(value) = item.value.as_deref() {
        if let Some(path) = resolve_media_path_from_text(value) {
            return Some(path);
        }
    }

    if let Some(path) = resolve_media_path_from_text(&item.name) {
        return Some(path);
    }

    None
}

fn get_file_under_cursor_checked(
    resolver: &mut ItemResolver,
    slow_probe_count: &mut u32,
) -> Option<PathBuf> {
    let started = Instant::now();
    let result = get_file_under_cursor(resolver);

    if started.elapsed() >= Duration::from_millis(EXPLORER_PROBE_SLOW_MS) {
        *slow_probe_count = slow_probe_count.saturating_add(1);
    } else {
        *slow_probe_count = 0;
    }

    result
}

/// The region a preview of the file under the pointer is kept off, as the `Avoid`
/// setting has it for the item that file is: the name the item draws at `Filename`,
/// that name with the columns beside it at `Details`, or the item's own box at either
/// setting when the view reports no text for it.
///
/// Asked when a preview is about to be shown rather than with every probe. A probe
/// answers which file the pointer is on, and that answer is what the rest of the loop
/// runs on; where that file's name is drawn is a question only a preview asks, and it
/// is asked here so that a pointer merely sweeping over a list pays nothing for it. A
/// walk that finds no item at all leaves the preview placed as it always was.
fn avoid_box_under_cursor(resolver: &ItemResolver, point: POINT) -> Option<(i32, i32, i32, i32)> {
    // Read before the walk as well as inside `avoid_box`: with the setting off there
    // is no region to be had, so the item is never asked for one.
    if avoid_mode() == AvoidMode::Off {
        return None;
    }

    let item = uia_item_from_point(resolver, point, true)?;
    item.avoid_box()
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
        let mut placement = WINDOWPLACEMENT {
            length: std::mem::size_of::<WINDOWPLACEMENT>() as u32,
            ..Default::default()
        };
        if GetWindowPlacement(hwnd, &mut placement).is_ok() {
            return placement.showCmd == SW_SHOWMAXIMIZED.0 as u32;
        }
    }
    false
}

/// Check if a window is fullscreen (covers the display it is on)
///
/// The display is the one the window is on, not the primary one that `SM_CXSCREEN`
/// reports: on a 4K display beside a 1080p primary, every window larger than 1920 by
/// 1080 used to read as fullscreen although most of that display was still showing,
/// and a game covering a 1080p display beside a 4K one was missed although it covered
/// the display it was on completely. Both answers are wrong in the way that matters
/// here — Explorer is read as hidden behind a window that does not reach it, or as
/// reachable behind one that covers the display it is on.
fn is_window_fullscreen(hwnd: HWND) -> bool {
    // How far past its display a window may reach and still be that display's
    // fullscreen window. A window that covers a display is that display to the pixel
    // or within a border drawn outside it.
    const FULLSCREEN_TOLERANCE_PX: i32 = 2;

    unsafe {
        let mut window_rect = RECT::default();
        if GetWindowRect(hwnd, &mut window_rect).is_err() {
            return false;
        }

        let monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
        if monitor.is_invalid() {
            return false;
        }

        let mut info = MONITORINFO {
            cbSize: std::mem::size_of::<MONITORINFO>() as u32,
            ..Default::default()
        };
        if !GetMonitorInfoW(monitor, &mut info).as_bool() {
            return false;
        }

        let display = info.rcMonitor;
        window_rect.left <= display.left + FULLSCREEN_TOLERANCE_PX
            && window_rect.top <= display.top + FULLSCREEN_TOLERANCE_PX
            && window_rect.right >= display.right - FULLSCREEN_TOLERANCE_PX
            && window_rect.bottom >= display.bottom - FULLSCREEN_TOLERANCE_PX
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

/// Whether the pointer is inside a region a preview published as holding it,
/// reading the cursor position here: the caller runs before the polling loop has
/// read it for this iteration. Answers `false` without a syscall when nothing on
/// screen holds the pointer.
fn preview_pointer_hold_now() -> bool {
    let mut cursor_pos = POINT::default();
    unsafe {
        if GetCursorPos(&mut cursor_pos).is_err() {
            return false;
        }
    }

    preview_pointer_hold(cursor_pos.x, cursor_pos.y)
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

/// What a keyboard preview is resolved from: the item Explorer says holds the
/// focus, and the window its view belongs to.
struct FocusedItemInfo {
    item: HoveredItem,
    /// The frame of the window the item is drawn in — the window whose views can
    /// be the one holding it.
    root_window: Option<isize>,
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

/// The item Explorer says holds the keyboard focus, as the view reports it.
fn get_focused_explorer_item(resolver: &ItemResolver) -> Option<FocusedItemInfo> {
    // Only works when Explorer is the foreground window
    let foreground = unsafe { GetForegroundWindow() };
    if foreground.is_invalid() || !is_explorer_window(foreground) {
        return None;
    }

    let item = uia_item_from_focus(resolver)?;
    if item.name.is_empty() || is_container_name(&item.name) {
        return None;
    }
    if item.bounds.right <= item.bounds.left || item.bounds.bottom <= item.bounds.top {
        return None;
    }

    Some(FocusedItemInfo {
        root_window: root_window_of_item(&item),
        item,
    })
}

/// The frame of the window an item is drawn in, which is the window whose views
/// can be the one holding it. A provider that reports no window of its own leaves
/// the foreground window, which is Explorer's while the keyboard is driving.
fn root_window_of_item(item: &HoveredItem) -> Option<isize> {
    let window = HWND(item.native_window as *mut core::ffi::c_void);
    if !window.is_invalid() {
        let root = unsafe { GetAncestor(window, GA_ROOT) };
        if !root.is_invalid() {
            return Some(root.0 as isize);
        }
    }

    let foreground = unsafe { GetForegroundWindow() };
    (!foreground.is_invalid()).then_some(foreground.0 as isize)
}

/// The file a keyboard preview is about.
///
/// The item the focus is on is turned into a file the same way the pointer's item
/// is: by the view that is showing it, which is what resolves a search result whose
/// file lives in another folder — a name has no folder to be found in when the
/// results span folders, and two results may share one. The item's own accessible
/// value answers when its position is not reported, and when the name itself is a
/// path. Nothing else is tried: a name that nothing can vouch for is left
/// unanswered rather than guessed at.
fn resolve_focused_item_to_path(
    resolver: &mut ItemResolver,
    focused: &FocusedItemInfo,
) -> Option<PathBuf> {
    if let Some(root_key) = focused.root_window {
        if let Some(path) = item_file_path(resolver, root_key, &focused.item) {
            return Some(path);
        }
    }

    if let Some(path) = resolve_media_path_from_text(&focused.item.name) {
        return Some(path);
    }

    focused
        .item
        .value
        .as_deref()
        .and_then(resolve_media_path_from_text)
}

/// Main loop for explorer hook
pub fn run_explorer_hook() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
    }

    // What both paths resolve an item to a file with: one UI Automation client
    // whose property reads are batched into a round trip per element, plus the
    // view that answered last for a window.
    let mut resolver = ItemResolver::new(automation_client());

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
                c.trigger_key_enabled,
            );
            // Resolved once per config change instead of once per tick.
            let vk = off_trigger_key_to_vk(&c.trigger_key);
            (snapshot, vk)
        })
        .unwrap_or(((true, 0, TriggerKeyMode::Disable, 750, true), Some(0x12)));
    let mut slow_explorer_probe_count = 0u32;
    let mut explorer_probe_backoff_until: Option<Instant> = None;
    let mut last_display_signature = current_display_signature();
    let mut last_display_check = Instant::now();
    // What the Shell objects in hand were built against, and when they were last
    // built: a restart is what invalidates them, and a build that came out missing
    // is tried again rather than left missing for the rest of the run.
    let mut last_explorer_restarts = explorer_restart_count();
    let mut last_shell_build = Instant::now();

    while RUNNING.load(Ordering::SeqCst) {
        // Explorer restarting is not something the resolver can recover from by
        // itself: the window collection it holds and the view that answered through
        // it are served by explorer.exe, and a proxy into a process that is gone
        // does not reconnect — its calls fail for good, which is what left every
        // hover after a restart with no answer at all. So the collection is built
        // again when a restart has been counted, and a collection that could not be
        // built when it was first asked for is built again on a slow retry, since
        // leaving it missing answers exactly as a dead one does.
        let explorer_restarts = explorer_restart_count();
        if explorer_restarts != last_explorer_restarts
            || (resolver.shell_windows.is_none()
                && last_shell_build.elapsed() >= Duration::from_millis(SHELL_COLLECTION_RETRY_MS))
        {
            last_explorer_restarts = explorer_restarts;
            last_shell_build = Instant::now();
            resolver.rebuild_shell();
            // Both caches are keyed by window handles, and those belonged to the
            // shell that is gone.
            clear_shell_view_probe_caches();

            // Nothing on screen is about a shell that exists any more, and the
            // windows the pointer could be over are the new shell's — which is
            // still putting them up, so the probes wait a moment before they ask.
            hide_preview();
            last_file = None;
            keyboard_file = None;
            is_keyboard_hover = false;
            last_focused_key = None;
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
                Some(Instant::now() + Duration::from_millis(EXPLORER_RESTART_BACKOFF_MS));
            current_state = get_explorer_state();
            last_state_check = Instant::now();
        }

        // The pointer's answer belongs to the tick that produced it: the list under
        // a parked pointer can have moved on by the next one.
        resolver.forget_probe();

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

        // Walking every display is not something to do on every tick for a change that
        // is noticed within this of happening anyway, and rebuilding what it
        // invalidates already waits out `DISPLAY_CHANGE_BACKOFF_MS`.
        if last_display_check.elapsed() >= Duration::from_millis(DISPLAY_CHECK_MS) {
            last_display_check = Instant::now();

            if let Some(display_signature) = current_display_signature() {
                if display_signature_changed(last_display_signature.as_ref(), &display_signature) {
                    last_display_signature = Some(display_signature);
                    clear_shell_view_probe_caches();
                    resolver.forget_view();
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
                    if GetCursorPos(&mut cursor_pos).is_ok() {
                        let dpi = monitor_dpi_from_point(cursor_pos.x, cursor_pos.y);
                        let threshold = pointer_pause.move_threshold_px(false, dpi);

                        if (cursor_pos.x - last_cursor_pos.x).abs() > threshold
                            || (cursor_pos.y - last_cursor_pos.y).abs() > threshold
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
                config.trigger_key_enabled,
            );
            trigger_key_vk = off_trigger_key_to_vk(&config.trigger_key);
        }

        let preview_enabled = config_snapshot.0;
        let hover_delay_ms = config_snapshot.1;
        let trigger_key_mode = config_snapshot.2;
        let same_file_rehover_delay_ms = config_snapshot.3;
        let trigger_key_enabled = config_snapshot.4;

        // One question, two settings: the key either stops previews while it is
        // held, or is the only thing that lets them happen. Either way, what is left
        // to do when they are not allowed is the same as when they are turned off.
        // A key that is switched off is not asked about, and holds nothing back.
        let trigger_key_down = trigger_key_enabled && trigger_key_vk.is_some_and(key_is_down);
        let previews_allowed =
            !trigger_key_enabled || trigger_key_mode.allows_previews(trigger_key_down);

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
                // A pointer that a preview is holding is not evidence that the
                // user has left: the preview is on top of Explorer, so the check
                // below cannot see Explorer under it.
                if !is_cursor_over_explorer_full() && !preview_pointer_hold_now() {
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

            // Whether the pointer is on a preview that holds it: a text preview the
            // user can read or select from — on the preview, or inside the margin
            // around it, which covers the gap it crosses on its way from the file it
            // belongs to — or the spinner a page is being rendered behind, which the
            // pointer reaches only by drifting into its box, and which has no page
            // yet to hand the pointer back to. While that holds, what is on screen
            // is what the user is waiting on rather than something in the way, so it
            // is not dismissed, and the file under the pointer is not resolved, so
            // it cannot be replaced by whatever it covers.
            //
            // Read once here because more than one path in this loop asks: the
            // dismissal below, the hover resolver, the mouse hover delay, and the
            // "Explorer is visible but not focused" branch, where a pointer that
            // is not over Explorer is not a reason to close a preview either.
            let loop_now = Instant::now();

            // Read straight from the published region, with nothing held over from
            // the last tick: the moment the pointer is out of it, the preview is
            // treated the way it was before the pointer ever touched it.
            let pointer_hold = preview_pointer_hold(cursor_pos.x, cursor_pos.y);

            let move_threshold = pointer_pause.move_threshold_px(
                is_keyboard_hover || keyboard_screen_owner,
                monitor_dpi_from_point(cursor_pos.x, cursor_pos.y),
            );
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

            // A preview that holds the pointer is the exception to that rule: while
            // it does, the preview is kept — see `pointer_hold` above for what
            // holding it covers. A preview the pointer cannot work with keeps the
            // old behaviour, which is to close as soon as it is touched.
            let guard_active = video_hover_guard_until
                .map(|until| Instant::now() < until)
                .unwrap_or(false);
            let should_dismiss_for_preview_hover =
                (over_image_preview && !pointer_hold) || (over_video_preview && !guard_active);

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
                && !pointer_hold
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
                        // folder remembered for the window and the view that answered
                        // for it. A name looked up against those resolves to something
                        // that is not there, or to nothing at all.
                        clear_shell_view_probe_caches();
                        resolver.forget_view();
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
                        if let Some(focused_info) = get_focused_explorer_item(&resolver) {
                            let focused_key = FocusedItemKey::new(
                                focused_info.item.name.clone(),
                                &focused_info.item.bounds,
                            );

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
                    if let Some(current_file) =
                        get_file_under_cursor_checked(&mut resolver, &mut slow_explorer_probe_count)
                    {
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
                    let mut keep_while_pointer_held = false;
                    if let Some(current_file) =
                        get_file_under_cursor_checked(&mut resolver, &mut slow_explorer_probe_count)
                    {
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
                    } else if pointer_hold {
                        // Nothing under the pointer: it is on its way to (or on) the
                        // preview, or on the spinner standing in for one, which is
                        // not the user leaving the file it shows.
                        keep_while_pointer_held = true;
                    } else if let Some(file) = last_file.clone() {
                        suppressed.suppress(file);
                    }

                    if !keep_while_pointer_held {
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
                if let Some(focused_info) = get_focused_explorer_item(&resolver) {
                    // The name alone cannot tell two observations apart: a search
                    // can hold the same name in more than one folder, and the box is
                    // what says which of them the keyboard is on.
                    let focused_key = FocusedItemKey::new(
                        focused_info.item.name.clone(),
                        &focused_info.item.bounds,
                    );

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
                            if let Some(path) =
                                resolve_focused_item_to_path(&mut resolver, &focused_info)
                            {
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
                                        focused_info.item.bounds.left,
                                        focused_info.item.bounds.top,
                                        focused_info.item.bounds.right,
                                        focused_info.item.bounds.bottom,
                                        focused_info.item.avoid_box(),
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
                        if let Some(path) =
                            resolve_focused_item_to_path(&mut resolver, &focused_info)
                        {
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
                                    focused_info.item.bounds.left,
                                    focused_info.item.bounds.top,
                                    focused_info.item.bounds.right,
                                    focused_info.item.bounds.bottom,
                                    focused_info.item.avoid_box(),
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
            // goes for a pointer held by what is on screen — inside a scrollable
            // text preview, or on the spinner it is waiting behind: the file under
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
                || pointer_hold
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
                        &mut resolver,
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
                            show_preview(
                                &file_path,
                                cursor_pos.x,
                                cursor_pos.y,
                                avoid_box_under_cursor(&resolver, cursor_pos),
                            );
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
    // an apartment that has already been torn down, which faults — and the resolver
    // holds the Shell window collection, the batched property request and the view
    // that answered last, not just the automation client.
    drop(resolver);

    unsafe {
        CoUninitialize();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A signature of one display, which is all the tests here need: what is compared
    /// is a display's own place and scale, not what the desktop adds up to.
    fn signed(displays: &[(u32, bool)]) -> DisplaySignature {
        DisplaySignature {
            displays: displays
                .iter()
                .map(|(dpi, primary)| DisplayEntry {
                    rect: (0, 0, 1920, 1080),
                    dpi: *dpi,
                    primary: *primary,
                })
                .collect(),
        }
    }

    /// What the signature is for: a display rescaled, or another one made the primary,
    /// rearranges everything Explorer is drawing without moving the desktop's own
    /// bounds — so both are changes, and a desktop that did not change is not.
    #[test]
    fn a_display_change_is_a_rescale_or_another_primary_display() {
        let before = signed(&[(96, true), (96, false)]);

        assert!(
            !display_signature_changed(Some(&before), &signed(&[(96, true), (96, false)])),
            "the same desktop is not a change"
        );
        assert!(
            display_signature_changed(Some(&before), &signed(&[(144, true), (96, false)])),
            "the primary display rescaled to 150% is one"
        );
        assert!(
            display_signature_changed(Some(&before), &signed(&[(96, true), (144, false)])),
            "and so is the second display rescaled, which the desktop's bounds do not show"
        );
        assert!(
            display_signature_changed(Some(&before), &signed(&[(96, false), (96, true)])),
            "and another display becoming the primary one"
        );

        assert!(
            !display_signature_changed(None, &before),
            "the first look at a desktop is not a change"
        );
    }

    /// The width a name is drawn at is the name's own: a longer name measures wider
    /// than a short one, which is what makes the region a preview is kept off the name
    /// rather than the column it sits in.
    #[test]
    fn a_name_is_measured_by_what_it_takes() {
        let short = drawn_name_width("aa.txt", 96).expect("a short name measures");
        let middle = drawn_name_width("mid-length-name.txt", 96).expect("a middle name measures");
        let long =
            drawn_name_width("a-very-long-file-name-here.txt", 96).expect("a long name measures");

        assert!(short > 0, "a name takes some room: {short}");
        assert!(
            short < middle && middle < long,
            "a longer name takes more room: {short} < {middle} < {long}"
        );
    }

    /// A name is measured in the pixels of the display it is drawn on, which is what a
    /// scaled display draws it twice as wide in. Measuring one at the system's scale on
    /// another display's is what covered the name: a provider that hands out no window
    /// left the display unknown, and the region came out at half the name on a display
    /// at 200% — see `region_display_dpi`.
    #[test]
    fn a_name_is_measured_at_the_scale_of_its_display() {
        let plain = drawn_name_width("mid-length-name.txt", 96).expect("a name measures");
        let scaled = drawn_name_width("mid-length-name.txt", 192).expect("a name measures");

        let ratio = scaled as f32 / plain as f32;
        assert!(
            (ratio - 2.0).abs() < 0.05,
            "twice the display draws twice the name: {plain} vs {scaled}"
        );
    }

    /// A display that cannot be told apart from another is no display to measure
    /// against, so a name without one is left as the view reported it rather than
    /// measured at a scale that is not its own — see `drawn_name_width`.
    #[test]
    fn a_name_without_a_display_is_not_measured() {
        assert_eq!(drawn_name_width("report.txt", 0), None);
    }

    /// A name with nothing in it has no width to be kept off, so the box the view
    /// reported is left as it is — see `HoveredItem::name_box`.
    #[test]
    fn a_name_with_nothing_in_it_is_not_measured() {
        assert_eq!(drawn_name_width("", 96), None);
    }

    /// The name measured is the name the view shows, extension and all: what a
    /// preview is kept off is the whole of what is drawn, not the stem it starts with.
    #[test]
    fn a_name_is_measured_with_its_extension() {
        let listed = drawn_name_width("report.txt", 96).expect("a listed name measures");
        let stem = drawn_name_width("report", 96).expect("a stem measures");

        assert!(
            listed > stem,
            "the extension takes room of its own: {stem} vs {listed}"
        );
    }
}

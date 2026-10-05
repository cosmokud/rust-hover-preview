//! The hover hook: the thread that watches Explorer's listing and puts a preview of
//! whatever the pointer comes to rest on in front of it.
//!
//! It is one loop over one pointer. Everything here exists to answer a question that loop
//! asks, and the parts are split by what each one is asked for:
//!
//! - `views` — the shapes a look comes back in: the view that answered, the item the
//!   pointer is over, the file that item names.
//! - `resolver` — the UI Automation client, the Shell window collection, and the view that
//!   answered last. What holds a COM object for the rest of this.
//! - `view_sort` — how long the hook waits between looks, and the order a folder is listed
//!   in.
//! - `probe_trace` — what a run cost, when `RHP_HOOK_TRACE` asks for it, and whether a
//!   probe is worth making at all.
//! - `locations` — an Explorer URL turned into a path, and the place a view describes.
//! - `ui_elements` — the UI Automation walks: the item under a point, and the boxes its
//!   text is drawn in.
//! - `shell_views` — the same question asked of the Shell instead: the views a window
//!   holds.
//! - `window_state` — how many Explorer windows there are, and what they are doing.
//! - `input` — the keys and the buttons, read once per tick.
//! - `hover_location` — where the pointer is, and which listing it is over.
//! - `pin_watch` — what a pinned preview watches while it waits to be given the next file.
//! - `focus_items` — the item the keyboard is on rather than the one under the pointer.
//! - `main_loop` — `run_explorer_hook` itself.
//!
//! The order a hover goes through is the order of that list: the pointer is read, the
//! state of Explorer's windows decides whether a look is worth making, the look produces a
//! view and an item, the item names a file, and the loop decides what to do with it.
//! `main_loop` is last because it is the only part that drives the rest: it holds the
//! state between ticks and asks each of them the one question it is for.
//!
//! The parts share one namespace through this module — each reads the others as one set of
//! items rather than naming them through their own — so that moving one question across a
//! boundary is not a change to it.

mod focus_items;
mod hover_location;
mod input;
mod locations;
mod main_loop;
mod pin_watch;
mod probe_trace;
mod resolver;
mod shell_views;
mod ui_elements;
mod view_sort;
mod views;
mod window_state;

use crate::app::engine_processes;
use crate::config::config::{
    AvoidMode, TriggerKeyMode, DEFAULT_HOVER_DELAY_MS, DEFAULT_PIN_UPDATE_ENABLED,
    DEFAULT_PIN_UPDATE_ON_HOVER, DEFAULT_SAME_FILE_REHOVER_DELAY_MS, DEFAULT_SETTLING_DELAY_MS,
    DEFAULT_TICK_MS, DEFAULT_TRIGGER_KEY_AFFECT_PIN_MODE,
};
use crate::engines::webview_preview;
use crate::formats::video_formats::is_video_file;
use crate::shell::wheel_input;
use crate::ui::preview_window::{
    cursor_preview_hover, hide_preview, kill_stray_video_process, monitor_dpi_from_point,
    note_engine_page_drag, pinned, pinned_path, pointer_item_box, pointer_item_holds,
    preview_pointer_hold, preview_screen_rect, preview_stall_ms, publish_pin_media_press,
    publish_pointer_item_box, request_pin_end, show_preview, show_preview_keyboard,
    take_pin_resumed, update_pinned_preview, PreviewCursorHover,
};
use crate::{CONFIG, RUNNING};
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::os::windows::ffi::OsStrExt;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Mutex;
use std::time::{Duration, Instant};
use windows::core::{w, IUnknown, Interface, GUID, VARIANT};
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
    UIA_GroupControlTypeId, UIA_LegacyIAccessiblePatternId, UIA_ListItemControlTypeId,
    UIA_NamePropertyId, UIA_NativeWindowHandlePropertyId, UIA_SelectionPatternId,
    UIA_TextControlTypeId, UIAutomationPropertyInfo, UIAutomationType_Int, UIA_CONTROLTYPE_ID,
    UIA_PROPERTY_ID,
};
use windows::Win32::UI::HiDpi::{GetDpiForMonitor, GetDpiForSystem, MDT_EFFECTIVE_DPI};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetAsyncKeyState, GetDoubleClickTime, VK_DELETE, VK_DOWN, VK_END, VK_HOME, VK_LBUTTON, VK_LEFT,
    VK_MBUTTON, VK_NEXT, VK_PRIOR, VK_RBUTTON, VK_RETURN, VK_RIGHT, VK_TAB, VK_UP, VK_XBUTTON1,
    VK_XBUTTON2,
};
use windows::Win32::UI::Shell::PropertiesSystem::PROPERTYKEY;
use windows::Win32::UI::Shell::{
    IFolderView, IFolderView2, IPersistFolder2, IShellBrowser, IShellItem, IShellView,
    IShellWindows, IWebBrowser2, ItemIndex_Property_GUID, SHCreateItemFromIDList,
    SID_STopLevelBrowser, ShellWindows, FWF_AUTOARRANGE, SIGDN_DESKTOPABSOLUTEPARSING,
    SIGDN_FILESYSPATH, SIGDN_NORMALDISPLAY, SORTCOLUMN, SORTDIRECTION, SORT_DESCENDING,
};
use windows::Win32::UI::WindowsAndMessaging::{
    EnumWindows, GetAncestor, GetClassNameW, GetCursorPos, GetForegroundWindow, GetWindowPlacement,
    GetWindowRect, IsChild, IsIconic, IsWindow, IsWindowVisible, SystemParametersInfoW,
    WindowFromPoint, GA_ROOT, SPI_GETICONTITLELOGFONT, SW_SHOWMAXIMIZED,
    SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS, WINDOWPLACEMENT,
};

// What the rest of the crate asks for by name.
pub use main_loop::run_explorer_hook;
pub use view_sort::note_explorer_restart;
pub(crate) use view_sort::{view_sort_of, SortKey, ViewSort};
pub(crate) use window_state::is_foreground_explorer;

// The parts, in one namespace (see the note above). The macro is named rather than globbed:
// it is written from two parts and lives with the counter it is written beside.
use focus_items::*;
use hover_location::*;
use input::*;
use locations::*;
use pin_watch::*;
pub(crate) use probe_trace::note_pin_click;
use probe_trace::*;
use resolver::*;
use shell_views::*;
use ui_elements::*;
use view_sort::*;
use views::*;
use window_state::*;

#[cfg(test)]
mod tests;

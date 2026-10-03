//! The item the keyboard is on rather than the one under the pointer: the focused list,
//! the item it holds selected, and the path that item names.
//!
//! A walk from the focus rather than from the point, asked when the pin's own keys
//! move the listing and when the keyboard is the only thing that has — and a listing
//! answers for its own selection whoever has the keyboard, which is what makes it the
//! reading a click behind a focused pin is followed by.

use super::*;

/// The file the view under the pointer has selected, read from the view's own selection
/// rather than from the point or the keyboard focus.
///
/// The list is what answers for its own selection, through the same pattern the focused
/// list is asked with when the view reports the list itself as focused (see
/// `selected_item_of_focused_list`) — and a list answers it whoever has the keyboard, which
/// is what makes this the reading a click behind a focused pin is followed by. The walk
/// starts at what the point is on and climbs to the list the way the item walks climb to
/// the item (see `walk_to_item`), and the item it names is resolved the way the pointer's
/// own item is: by the view that is showing it (see `item_file_path`).
///
/// A selection nothing names a file for is an answer rather than a reason to ask the
/// element above: the list holds this item, and nothing above it holds another selection.
/// Nothing at all is answered where the point is on no list, or where the shell will not
/// describe the item the list holds.
pub(super) fn selected_file_in_view(
    resolver: &mut ItemResolver,
    pointer: &PointerTick,
) -> Option<PathBuf> {
    let automation = resolver.automation.as_ref()?;
    let cache = resolver.cache.as_ref()?;
    let walker = resolver.walker.as_ref()?;
    let window = item_window_of(pointer.window)?;

    let mut element =
        unsafe { automation.ElementFromPointBuildCache(pointer.point, cache) }.ok()?;
    for _ in 0..POINTER_ITEM_ANCESTOR_LIMIT {
        if let Ok(pattern) = unsafe {
            element.GetCurrentPatternAs::<IUIAutomationSelectionPattern>(UIA_SelectionPatternId)
        } {
            if let Ok(selection) = unsafe { pattern.GetCurrentSelection() } {
                if let Ok(count) = unsafe { selection.Length() } {
                    for index in 0..count.min(16) {
                        let Ok(candidate) = (unsafe { selection.GetElement(index) }) else {
                            continue;
                        };
                        if !element_names_an_item(&candidate) {
                            continue;
                        }
                        if let Some(item) = walk_to_item(resolver, &candidate, None, false) {
                            return item_file_path(resolver, &window, &item);
                        }
                        return None;
                    }
                }
                return None;
            }
        }
        element = unsafe { walker.GetParentElementBuildCache(&element, cache) }.ok()?;
    }

    None
}

/// Whether the window under the pointer is an Explorer window or one inside one.
///
/// The window is handed in rather than read here: what a tick has under the pointer is one
/// window, and the walk up from it to a class the shell is known by is the same walk
/// whoever starts it.
/// Keep this HWND/class based; calling ShellWindows here caused Explorer-side
/// COM providers to allocate while we were merely checking cursor position.
pub(super) fn is_cursor_over_explorer_full(window: HWND) -> bool {
    unsafe {
        if window.is_invalid() {
            return false;
        }

        // Walk up parent windows to find Explorer window
        let mut current_hwnd = window;

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

/// Whether a window is one of Explorer's folder windows.
///
/// The class is the whole of the test, and only the class: a folder view is drawn in the
/// browser frame, and this is the same answer the walk over Explorer's windows is made
/// with (`explorer_browser_class_matches`), so the loop cannot count a window as Explorer's
/// in one place and not in the other. What it keeps out is the rest of the shell, which is
/// explorer.exe as well — the desktop (`Progman`, `WorkerW`), the taskbar, the Start menu,
/// the search box. Reading one of those as an Explorer window is what made the foreground
/// test answer yes with the desktop in front, which is what Show Desktop leaves there: the
/// state never left `ActiveFocus`, the counts that would have said every Explorer window
/// was minimized were never asked for, and a preview that was already up was never taken
/// down.
///
/// The ask is one class lookup, kept for the window (see `EXPLORER_WINDOW_CACHE`). The
/// process behind the window used to be read as well, for any window that was not one of
/// the two classes — an `OpenProcess` and an image name for every window the pointer
/// walked over — and it answered yes for every piece of the shell that is explorer.exe,
/// which is the answer this test must never give.
pub(super) fn is_explorer_window(hwnd: HWND) -> bool {
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

    let is_explorer = explorer_browser_class_matches(hwnd);

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
pub(super) struct FocusedItemInfo {
    pub(super) item: HoveredItem,
    /// The frame of the window the item is drawn in — the window whose views can
    /// be the one holding it.
    root_window: Option<isize>,
}

impl FocusedItemInfo {
    /// The window the item is resolved in, as the item path asks for one: the frame whose
    /// views can hold it, and the window the item's own provider says it is drawn in —
    /// which is what says which of that frame's views drew it (see `ItemWindow`).
    ///
    /// The keyboard has no pointer to take a window from, so where the provider reports
    /// none — the shell's item provider often does not — the frame's views are told apart
    /// by what they answer, which is how the keyboard path resolved its item before the
    /// window the item is drawn in was read at all.
    fn item_window(&self) -> Option<ItemWindow> {
        Some(ItemWindow {
            frame: self.root_window?,
            drawn_in: self.item.native_window,
        })
    }
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
pub(super) struct FocusedItemKey {
    pub(super) name: String,
    pub(super) rect: (i32, i32, i32, i32),
}

impl FocusedItemKey {
    pub(super) fn new(name: String, rect: &RECT) -> Self {
        Self {
            name,
            rect: (rect.left, rect.top, rect.right, rect.bottom),
        }
    }
}

/// Whether a UIA element names a file: an Explorer item does, and a list, a header
/// or a toolbar does not.
pub(super) fn element_names_an_item(element: &IUIAutomationElement) -> bool {
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
pub(super) fn selected_item_of_focused_list(
    element: &IUIAutomationElement,
) -> Option<IUIAutomationElement> {
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
pub(super) fn get_focused_explorer_item(resolver: &ItemResolver) -> Option<FocusedItemInfo> {
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
pub(super) fn root_window_of_item(item: &HoveredItem) -> Option<isize> {
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
pub(super) fn resolve_focused_item_to_path(
    resolver: &mut ItemResolver,
    focused: &FocusedItemInfo,
) -> Option<PathBuf> {
    if let Some(window) = focused.item_window() {
        if let Some(path) = item_file_path(resolver, &window, &focused.item) {
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

/// The place the item the keyboard is on was read in: the view that drew it, and what that view
/// was showing, in the form places are compared in (see `place_of_window`).
pub(super) fn focused_item_location(
    resolver: &mut ItemResolver,
    focused: &FocusedItemInfo,
) -> Option<HoverLocation> {
    place_of_window(resolver, focused.item_window()?)
}

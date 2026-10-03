//! What Explorer's own windows are doing: how many there are, how many are showing, how
//! many of those the window in front does not hide, and what that count means for how
//! often the hook looks at all.
//!
//! The cheap checks are here too — whether the window in front is Explorer at all,
//! whether a click landed on a listing — because they are the answers that keep a walk
//! from being made to find out none of it.

use super::*;

/// Quick check if foreground window is Explorer (cheap, no COM)
///
/// The pin key asks it as well, because a key pressed while another program is in front is a press
/// about that program and not about the pin. A pin the user has pressed is not that case: it is
/// the window in front itself, and the keyboard it answers is a question asked of the pin's own
/// window procedure rather than of this (see `pinned_key_command`).
pub(crate) fn is_foreground_explorer() -> bool {
    unsafe {
        let foreground = GetForegroundWindow();
        if foreground.is_invalid() {
            return false;
        }
        is_explorer_window(foreground)
    }
}

/// Whether the engine's window — the window a page of HTML is drawn in — owns the keyboard
/// right now, which is what a click into a page that runs makes it do. Nothing else the
/// engine draws can be asked the question: a document and a specimen refuse activation, so
/// a foreground window of this family is a page the user clicked into and had keys arrive
/// at, and the pin key needs no case of its own for the same reason — a press while this is
/// true is already answered as another program's key (`is_foreground_explorer`).
///
/// The question is asked of the foreground window rather than of the pointer, because the
/// pointer can be anywhere: a page the user has clicked into and then left the hand resting
/// beside still has the keyboard, and the keys that belong to it are the keys the app's own
/// readers must leave alone. The engine's own window is not the only one of its family in
/// front — a page the user has activated focuses a child of it — so the question covers the
/// children the way `cursor_preview_hover` does, in the other direction and for the same
/// reason: the surface is the engine's, and so is everything inside it.
pub(super) fn engine_owns_the_keyboard() -> bool {
    let engine = webview_preview::showing_hwnd();
    if engine == 0 {
        return false;
    }

    unsafe {
        let engine = HWND(engine as *mut _);
        let foreground = GetForegroundWindow();
        if foreground.is_invalid() {
            return false;
        }

        (foreground.0 as isize) == engine.0 as isize || IsChild(engine, foreground).as_bool()
    }
}

/// Check if a window is maximized
pub(super) fn is_window_maximized(hwnd: HWND) -> bool {
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
pub(super) fn is_window_fullscreen(hwnd: HWND) -> bool {
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

/// The region the window in front hides what is behind, where that window is one
/// that hides anything at all: a maximized or fullscreen window that is not
/// Explorer's. `None` is a foreground window nothing is hidden behind — an ordinary
/// one, Explorer itself, or no window at all.
///
/// What it is for is asking whether an Explorer window is behind it. Taking the
/// foreground window alone for the answer — a maximized window, so Explorer must be
/// hidden behind it — is wrong on an extended desktop, and was: a maximized window
/// hides what is on the display it covers and says nothing about an Explorer window
/// on the display beside it, so the state went to `HiddenByForeground` — which hides
/// the preview and never asks where the cursor is — and a pointer that crossed over
/// to the Explorer window on the second display showed nothing at all until the
/// click that made Explorer the foreground window. What a window hides is what its
/// own rectangle holds, which is what the walk over Explorer's windows is asked.
pub(super) fn foreground_cover_rect() -> Option<RECT> {
    unsafe {
        let foreground = GetForegroundWindow();
        if foreground.is_invalid() || is_explorer_window(foreground) {
            return None;
        }

        // Only a maximized or fullscreen window covers anything whole.
        if !(is_window_maximized(foreground) || is_window_fullscreen(foreground)) {
            return None;
        }

        let mut rect = RECT::default();
        GetWindowRect(foreground, &mut rect).ok()?;
        Some(rect)
    }
}

/// Whether a window is hidden behind the region the window in front covers.
///
/// A window with no region over it is hidden by nothing, and one whose own rectangle
/// cannot be read is answered as reachable for the same reason: what follows from
/// hidden is a state that stops asking where the cursor is, so where the two cannot
/// be told apart the answer that keeps the previews is the one to give.
pub(super) fn window_is_behind_cover(hwnd: HWND, cover: Option<RECT>) -> bool {
    let Some(cover) = cover else {
        return false;
    };

    let mut rect = RECT::default();
    if unsafe { GetWindowRect(hwnd, &mut rect) }.is_err() {
        return false;
    }

    rect_contains(cover, rect)
}

/// Whether one rectangle holds another whole.
pub(super) fn rect_contains(outer: RECT, inner: RECT) -> bool {
    inner.left >= outer.left
        && inner.top >= outer.top
        && inner.right <= outer.right
        && inner.bottom <= outer.bottom
}

/// Check if a window is minimized
pub(super) fn is_window_minimized(hwnd: HWND) -> bool {
    unsafe { IsIconic(hwnd).as_bool() }
}

/// Allocation-free equivalent of lowercasing the class name and searching for an
/// ASCII needle. A `u16` outside ASCII becomes U+FFFD under `to_string_lossy`
/// and can never match, so it is skipped.
pub(super) fn utf16_contains_ascii_ignore_case(haystack: &[u16], needle: &[u8]) -> bool {
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

pub(super) fn explorer_browser_class_matches(hwnd: HWND) -> bool {
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

pub(super) unsafe extern "system" fn enum_explorer_windows_callback(
    hwnd: HWND,
    lparam: LPARAM,
) -> BOOL {
    let counts = &mut *(lparam.0 as *mut ExplorerWindowCounts);

    if explorer_browser_class_matches(hwnd) {
        counts.total += 1;
        if IsWindowVisible(hwnd).as_bool() && !is_window_minimized(hwnd) {
            counts.visible += 1;
            if !window_is_behind_cover(hwnd, counts.cover) {
                counts.reachable += 1;
            }
        }
    }

    BOOL(1)
}

/// What the walk over Explorer's windows found: how many there are, how many are
/// showing, and how many of those are out from behind the region the window in front
/// covers. Uses top-level HWND enumeration instead of ShellWindows COM to avoid
/// making Explorer's shell automation providers allocate during idle polling.
pub(super) fn get_explorer_window_counts(cover: Option<RECT>) -> ExplorerWindowCounts {
    let mut counts = ExplorerWindowCounts {
        total: 0,
        visible: 0,
        reachable: 0,
        cover,
    };

    unsafe {
        let _ = EnumWindows(
            Some(enum_explorer_windows_callback),
            LPARAM(&mut counts as *mut ExplorerWindowCounts as isize),
        );
    }

    counts
}

/// Enum representing the state of Explorer windows for CPU optimization
#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) enum ExplorerState {
    /// No Explorer windows open at all - longest sleep
    NoExplorerWindows,
    /// All Explorer windows are minimized - long sleep
    AllMinimized,
    /// Every window Explorer has showing is behind the region a maximized or
    /// fullscreen window in front of it covers - long sleep
    HiddenByForeground,
    /// Explorer is visible but not in focus - medium sleep
    VisibleNotFocused,
    /// Explorer is in focus and cursor might be over it - active polling
    ActiveFocus,
}

/// Determine the current state of Explorer for CPU optimization
pub(super) fn get_explorer_state() -> ExplorerState {
    // Quick check: is foreground Explorer? (cheapest check)
    if is_foreground_explorer() {
        return ExplorerState::ActiveFocus;
    }

    // Everything past here is about the Explorer windows themselves: how many there
    // are, and whether the window in front leaves any of them out from behind itself.
    let counts = get_explorer_window_counts(foreground_cover_rect());
    explorer_state_from_counts(&counts, pinned())
}

/// Whether `window` is one of this app's own rather than Explorer's: the preview window
/// itself, or something inside the engine's window where a document is drawn.
///
/// A pinned window stands over the listing it was taken from, and it stands over it where
/// the files are. A click that lands on it is a click on the listing underneath — the user
/// is looking at a folder and clicking a file in it — so it is read as a click on a listing
/// rather than dropped for having landed on us. What the shell will name for such a point is
/// this window, which is why a click is read a second way wherever its own look at the point
/// answers nothing: out of the item the view has the focus on, which is where a click's own
/// selection is answered from the view's side (see `PinUpdateWatch::click_picked_item`).
pub(super) fn is_our_own_window(window: HWND) -> bool {
    if preview_window_is_at(window) {
        return true;
    }

    let engine = webview_preview::showing_hwnd();
    engine != 0
        && unsafe {
            let engine = HWND(engine as *mut _);
            window == engine || IsChild(engine, window).as_bool()
        }
}

/// Whether `window` is the preview window — the pinned one where a pin is up, and the
/// hover's where one is not. Both are this app's own, and both stand over a listing.
pub(super) fn preview_window_is_at(window: HWND) -> bool {
    !window.is_invalid() && crate::ui::preview_window::is_preview_window(window.0 as isize)
}

/// Whether a click where the pointer is can be read out of a listing behind it.
///
/// One of Explorer's own windows, or one of this app's own standing over one. Anything else
/// — the desktop, another program, a browser — is a click there is nothing behind to read,
/// and a click on it is left to expire rather than answered out of whatever happens to be
/// drawn underneath.
pub(super) fn click_is_over_a_listing(over_explorer: bool, over_our_own: bool) -> bool {
    over_explorer || over_our_own
}

/// Whether a press read on this tick is one a listing can answer for: the window it landed
/// on is Explorer's or this app's own, it is still there, and it is not a popup covering the
/// listing.
///
/// The last half is what `View ▸ Tiles` and `Sort ▸ Name` fail: those are popups over rows,
/// the press that picks one dismisses it, and the hand does not move — so the point is right
/// back over the row the popup was covering and the lookup answers with a file the user never
/// clicked. Whether the popup was destroyed by the press (`aimed_at_a_gone_window`) or merely
/// hidden, which is what XAML popups do, it is chrome on one tick or the other: a press read
/// while it is still up lands on it (`current_is_chrome`), and a press read on the poll after
/// it closed finds the listing already there but was aimed at the popup
/// (`previous_was_chrome`). Both are refused. Both ticks are asked because the press bit is
/// read on whichever tick finds it, and a shell that shows the popup for one tick longer or
/// takes it down sooner must not change what the press means.
pub(super) fn press_is_a_listing(
    over_explorer: bool,
    over_our_own: bool,
    aimed_at_a_gone_window: bool,
    current_is_chrome: bool,
    previous_was_chrome: bool,
) -> bool {
    click_is_over_a_listing(over_explorer, over_our_own)
        && !aimed_at_a_gone_window
        && !current_is_chrome
        && !previous_was_chrome
}

/// Whether a window under the pointer is a menu or flyout rather than anything a listing can
/// be read under: the popups Explorer puts over its own files.
///
/// A popup is Explorer's own as far as `is_cursor_over_explorer_full` is concerned — its owner
/// chain reaches the frame — so a press on it reads as a press on the listing. The class name is
/// what says otherwise. Every one of them carries a substring of its own kind: the XAML shell
/// hosts them in `Microsoft.UI.Content.PopupWindowSiteBridge` and
/// `Microsoft.UI.Content.DesktopChildSiteBridge`, and a classic menu window is `#32768`. The
/// three bare words are the older shells' naming, and none of the windows a listing is read
/// under has one: the list is `SysListView32` under `DirectUIHWND`, the frame and the view are
/// `ExplorerWClass` and `CabinetWClass`, and this app's own window is `RustHoverPreviewWindow`.
///
/// Read per call rather than cached: two `GetClassNameW` calls on every pinned tick.
pub(super) fn window_is_menu_popup_chrome(window: HWND) -> bool {
    class_is_menu_popup_chrome(&window_class_of(window))
}

/// The matching half of `window_is_menu_popup_chrome`, on the class name rather than on the
/// window, so that the list of what is and is not chrome is a fact about strings that can be
/// written down rather than one that needs a window put on somebody's screen.
pub(super) fn class_is_menu_popup_chrome(class: &str) -> bool {
    let class = class.to_ascii_lowercase();

    ["#32768", "sitebridge", "popup", "flyout", "menu"]
        .iter()
        .any(|chrome| class.contains(chrome))
}

/// Whether the window under the pointer on the tick before has been destroyed. Nothing on the
/// first tick of a watch, and nothing where that tick named no window at all: a tick that read no
/// window is a gap in the reading rather than a window that went away.
pub(super) fn window_is_gone(window: Option<HWND>) -> bool {
    window.is_some_and(|window| !window.is_invalid() && !unsafe { IsWindow(window) }.as_bool())
}

/// The state the counts come out as: which sleep the loop takes, and whether the
/// cursor is asked about at all.
///
/// `pin_up` is handed in rather than read here so that the one answer a pin changes is
/// a decision this function makes rather than a fact it goes and looks up — it is the whole
/// of what `PinUpdateWatch` depends on, and a rule that can only be tested by putting a window
/// on somebody's screen is a rule that goes untested.
pub(super) fn explorer_state_from_counts(
    counts: &ExplorerWindowCounts,
    pin_up: bool,
) -> ExplorerState {
    if counts.total == 0 {
        return ExplorerState::NoExplorerWindows;
    }

    if counts.visible == 0 {
        return ExplorerState::AllMinimized;
    }

    // Nothing Explorer has showing is out from behind the region the window in front
    // covers, on any display, so the pointer cannot reach one of them.
    if counts.reachable == 0 {
        return ExplorerState::HiddenByForeground;
    }

    // A window this app has put up and the user has not closed is a window of its own, and a pin
    // takes the focus off Explorer on purpose — `pin_take_focus`, so that the keys the pin answers
    // are the user's own rather than the listing's. Reading that arrangement as "Explorer is
    // showing but nobody is in it" is what dropped the loop to the medium cadence behind a
    // preview that is on screen and being worked in. Minimized, and behind a window that covers
    // it, are left where they are: there is no listing under the pointer in either, so a pin has
    // nothing to be shown another file from (see `ExplorerState`).
    if pin_up {
        return ExplorerState::ActiveFocus;
    }

    // Explorer windows exist and are visible, but not in foreground
    ExplorerState::VisibleNotFocused
}

/// The sleep a state is answered with, and how often that state is read again while it is being
/// answered with — the whole of the ladder, in one place because two branches of the loop are
/// paced by it: the hover machinery below, and the pinned tick above it, which is a reader only
/// because a pin in front of a listing is `ActiveFocus` however the keyboard is arranged (see
/// `explorer_state_from_counts`).
pub(super) fn explorer_pace(state: ExplorerState, tick_ms: u64) -> (u64, u64) {
    match state {
        ExplorerState::NoExplorerWindows => (DEEP_SLEEP_MS, STATE_RECHECK_DEEP_MS),
        ExplorerState::AllMinimized => (LONG_SLEEP_MS, STATE_RECHECK_LONG_MS),
        ExplorerState::HiddenByForeground => (LONG_SLEEP_MS, STATE_RECHECK_LONG_MS),
        ExplorerState::VisibleNotFocused => (MEDIUM_SLEEP_MS, STATE_RECHECK_MEDIUM_MS),
        ExplorerState::ActiveFocus => (tick_ms, STATE_RECHECK_ACTIVE_MS),
    }
}

/// Read the state of Explorer, and record it for the engines' away timer.
///
/// Whether an Explorer window is left reachable is the whole of what an engine that is not
/// marked `Persistent` is let go by (see `app::afk`), and the answer is one this loop already
/// works out for its own sleeps — so the recording is done here, on the read, rather than
/// anywhere the state is used: a state this function did not answer is a state the clock has
/// not been told about, and the engines keep reading the last thing it was told.
pub(super) fn read_explorer_state() -> ExplorerState {
    let state = get_explorer_state();

    crate::app::afk::note_explorer_reachable(matches!(
        state,
        ExplorerState::VisibleNotFocused | ExplorerState::ActiveFocus
    ));

    state
}

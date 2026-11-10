//! What Explorer's own windows are doing: how many there are, how many are showing, how
//! many of those the window in front does not hide, and what that count means for how
//! often the hook looks at all.
//!
//! The cheap checks are here too — whether the window in front is Explorer at all,
//! whether a click landed on a listing — because they are the answers that keep a walk
//! from being made to find out none of it.

use super::*;
use crate::shell::window_classes;

/// Whether the window in front is a listing this hook answers for — Explorer's folder window or a
/// Common Item Dialog (see `is_explorer_window`) — asked of the cached class lookup rather than of
/// COM, which is what makes it the cheap half of the state machine.
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

/// Whether the window in front is a listing this hook answers for — Explorer's folder window, a browse
/// dialog while `Preview Settings → General → Enable in Browse Dialogs` is on, or Everything's results
/// while `… → Enable in Everything` is — asked of the class alone: one `GetForegroundWindow` and the class
/// compares of the three class predicates, with no lock and no allocation. Each setting is its own
/// published fact, which costs one atomic load and no lock either (see `app::seam::a_listing_is_previewed`).
///
/// The question the pin's own placement asks, and the one `is_foreground_explorer` is not: a pin is
/// owed the front of its band where the window it was picked over is *either* kind, and a file picked
/// in a program's own `Ctrl`+`O` must leave the pin in the front of that dialog exactly as a file picked
/// in a folder view leaves it in the front of the folder — the press recorded the frame it was made in,
/// and a dialog is no less a listing for being somebody else's window (see `file_dialog`).
///
/// Not the cached form, either: the class of a window does not change under the answer, but a dialog is
/// a frame that appears and goes with one keystroke, and a memo read before it appeared would answer for
/// the window it replaced.
pub(crate) fn the_front_is_a_listing() -> bool {
    unsafe {
        let foreground = GetForegroundWindow();
        !foreground.is_invalid() && crate::app::seam::a_listing_is_previewed(foreground)
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
///
/// The engine's window is asked of the fact the viewer publishes — one atomic, no lock and no stale
/// tick, written where the handle itself moves (see `app::seam::ENGINE_HWND`): the window belongs to a
/// process this one does not share, so asking the engine here answers nothing at all in a run whose
/// previews are drawn elsewhere.
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
/// one, a listing itself (a folder window or a Common Item Dialog, which is what
/// `is_explorer_window` answers for both), the desktop, or no window at all.
///
/// **A listing never covers anything, whichever kind it is.** The class is asked before the
/// rectangle, so a maximized dialog cannot reach `window_is_behind_cover` either: a dialog is the
/// user working in a listing, and answering it as a window that hides what is behind would sleep the
/// loop out of the very state its own items are read in.
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
///
/// The window in front is handed in rather than read here, because the caller reads it once and asks
/// it both this and whether it is a listing: two reads of `GetForegroundWindow` on one tick are two
/// windows as far as an activation landing between them goes.
pub(super) fn foreground_cover_rect(foreground: HWND) -> Option<RECT> {
    unsafe {
        if foreground.is_invalid() || is_explorer_window(foreground) {
            return None;
        }

        // And the class before the rectangle: a window of the desktop's in front covers nothing,
        // whatever its own box says (see `a_foreground_of_this_class_covers_what_is_behind`).
        let mut name = [0u16; 32];
        let len = GetClassNameW(foreground, &mut name).max(0) as usize;
        let class = String::from_utf16_lossy(&name[..len.min(name.len())]);
        if !a_foreground_of_this_class_covers_what_is_behind(&class) {
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

/// Whether a window in front of this class may hide a listing behind it at all.
///
/// **The desktop never does**, whatever its rectangle says: `Progman` and `WorkerW` are drawn over
/// the display to the pixel, so read as ordinary windows they are a fullscreen window hiding every
/// Explorer window on it — the state that leaves the loop sleeping and the previews dead until some
/// other window is clicked. The desktop is the bottom of the desktop's own list and hides nothing; a
/// foreground that has been *left* on it — by a player of this app's killed in front, or by a press
/// on the wallpaper — is not a window over the listing. The two classes are kept out of Explorer's
/// own windows for the same reason (see `is_explorer_window`), and this is the same fact asked of a
/// class rather than of a window, so that it can be answered without a desktop underneath it.
pub(super) fn a_foreground_of_this_class_covers_what_is_behind(class: &str) -> bool {
    !(class.eq_ignore_ascii_case("Progman") || class.eq_ignore_ascii_case("WorkerW"))
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

pub(super) unsafe extern "system" fn enum_explorer_windows_callback(
    hwnd: HWND,
    lparam: LPARAM,
) -> BOOL {
    let counts = &mut *(lparam.0 as *mut ExplorerWindowCounts);

    if window_classes::explorer_class_matches(hwnd) {
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
    // The window in front, read once for both questions the state asks of it: whether it is a listing
    // (answered before the walk, which is the cheap half), and the region it hides what is behind
    // where it hides anything at all.
    let foreground = unsafe { GetForegroundWindow() };

    // Cheapest answer there is, and the one most hovers are read on: Explorer's own frame in front is
    // the active arrangement, and the walk over its windows is not made for it.
    if !foreground.is_invalid() && crate::shell::window_classes::explorer_class_matches(foreground)
    {
        return ExplorerState::ActiveFocus;
    }

    // **A Common Item Dialog and Everything's result window are listings and are nobody's Explorer
    // windows**, so they fall through to the counts rather than to the answer above: the walk does not
    // count them — counting them would move the pace and the covered-window rule, which are Explorer's
    // own — and a run with no folder window open would read `NoExplorerWindows` and stop asking where the
    // pointer is while the user picks a file in Chrome's `Ctrl`+`O` or hovers a row in Everything. So the
    // fact is handed to the rule, which is where the row it answers for is decided (see
    // `is_explorer_window` and `explorer_state_from_counts`).
    let foreground_is_a_listing = !foreground.is_invalid() && is_explorer_window(foreground);

    // Everything past here is about the Explorer windows themselves: how many there
    // are, and whether the window in front leaves any of them out from behind itself.
    let counts = get_explorer_window_counts(foreground_cover_rect(foreground));
    explorer_state_from_counts(
        &counts,
        crate::app::seam::the_pin_is_up(),
        foreground_is_a_listing,
    )
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
    if preview_window_is_at(window) || press_is_on_the_pins_own_menu(window) {
        return true;
    }

    // The engine's window is read from the fact the viewer publishes: the window itself is answered for
    // by `preview_window_is_at` above, being the viewer's, but a page inside it is not — a page is drawn by
    // the browser's own windows, which belong to the browser's process rather than to the viewer's, and
    // `IsChild` is the half that answers for it (see `app::seam::ENGINE_HWND` and
    // `webview_preview::browser_pid`).
    let engine = webview_preview::showing_hwnd();
    engine != 0
        && unsafe {
            let engine = HWND(engine as *mut _);
            window == engine || IsChild(engine, window).as_bool()
        }
}

/// Whether `window` is the preview's — the hover's where there is no pin, the pinned one where there
/// is, and every window a pin draws through besides: its film's host, the caption windows, the bubble,
/// its menu, the stand-in wearing its taskbar button. They are all windows of a viewer process, they are
/// all this app's own, and they all stand over a listing.
///
/// **The process is the whole of the answer**, and that is deliberate: the preview side makes windows as
/// it needs them, so a list of the ones it made would be a leak every time one was made and a list this
/// side had to be told about (see `app::seam::window_is_a_viewers`).
pub(super) fn preview_window_is_at(window: HWND) -> bool {
    !window.is_invalid() && crate::app::seam::window_is_a_viewers(window.0 as isize)
}

/// Whether a press outside every window of ours puts the standing window of ours behind the one it
/// landed in: the whole of what `Always On Top` off means.
///
/// **The press is the whole of the question, and nothing here asks whether Explorer is in front.** A
/// press inside the window that is already active — the whole of what a listing is made of: its tree,
/// its rows, the gap between them, its toolbar — moves no foreground at all. The frame stays the
/// foreground window and only a child's focus moves, so a demotion wired to that transition never saw
/// the press that made it (see `lower_pinned_preview`, which is where that history is written down).
/// What a press does do is name the window the user is working in, and a window of ours owes that
/// window its place whatever it is: Explorer, one of Explorer's children, another program, the
/// desktop.
///
/// `ours` is the answer `press_is_inside_ours` gives for the window under the press. Handed in rather
/// than read here because the same rule is answered for a point no window of ours can be made at (see
/// `press_is_outside_the_pin_window`, which is handed its answer for the same reason).
///
/// **`the_listing_answers` is what keeps a pick from flashing the listing over the pin.** A press in
/// the listing behind a pin that follows that listing is the press of a pick: the hook reads the file it
/// landed on and asks for the pin to be shown it, and the raise that comes of that ask is the answer the
/// press gets. Putting the pin behind the listing first is the same press answered twice, and what the
/// user sees between the two is the Explorer window over the pin for as long as the pick takes to
/// resolve — the flash a pick with the mouse has and a pick with the keyboard does not, because a key
/// moves the focus and presses nothing (see `note_a_press` and `app::seam::PIN_FOLLOWS_THE_LISTING`).
///
/// A press that names no file is answered all the same, and the pin still goes behind for it: the
/// hook's own `pin_press_was_not_a_pick` is the answer that press is owed, and it is made where the
/// file under the point is known to be one the pin cannot show.
pub(crate) fn a_press_puts_us_behind(
    pressed: bool,
    ours: bool,
    ours_stands: bool,
    always_on_top: bool,
    the_listing_answers: bool,
) -> bool {
    pressed && ours_stands && !always_on_top && !ours && !the_listing_answers
}

/// Whether the press just read is one the pick road is going to answer: a press in a listing, behind a
/// pin that follows the listing (see `a_press_puts_us_behind` and `app::seam::PIN_FOLLOWS_THE_LISTING`).
///
/// **Both halves are asked, and each is the half it is.** The listing is what makes a press a pick at
/// all — a press on another program, on the desktop, on the taskbar or on one of Explorer's own menus is
/// nobody's pick — and the switch is what makes the pick the pin's business. A pin that follows nothing
/// is a pin a press in the listing *is* the answer for, and it goes behind as it always has.
///
/// Asked of the class of the frame the press landed in rather than of the process behind it, and with no
/// lock and no allocation, because this is a low-level callback and every press on the machine waits for
/// it (see `window_classes::explorer_class_matches`). A Common Item Dialog is a listing a pick can be made
/// in exactly as a folder window is — the pin is standing beside the file the user is choosing in a
/// program's Open dialog, and the press that chooses it is the press that hands it over (see
/// `explorer_hook::file_dialog`).
pub(crate) fn the_pick_road_answers(window: HWND) -> bool {
    crate::app::seam::the_pin_follows_the_listing()
        && crate::app::seam::a_listing_is_previewed(window)
}

/// Whether the point a press landed on is inside any window of ours: the broad answer, and the one the
/// press road has to ask.
///
/// **`is_our_own_window` alone is not it, and that is the whole of what this adds.** That function
/// answers the listing's question — a press on a window of ours is a press on the listing it stands
/// over — so it answers for the windows of a viewer process and for the browser's page drawn inside the
/// engine's window. A pin's film is neither: mpv draws it in a window of its own process, and a press on
/// it read as somebody else's is a press that puts the pin behind its own picture (see
/// `the_pin_of_the_window`, which is the second half).
///
/// **And the hit test cannot be the whole of the question either.** A film's host answers
/// `HTTRANSPARENT` so the press reaches the pin's chrome, which means `WindowFromPoint` over the pin's own
/// picture names the *listing the pin was fitted over* — so the box the pin stands in is asked as well,
/// the same one the pick road asks (see `preview_window_rect` and `press_is_on_the_pin`).
pub(crate) fn press_is_inside_ours(window: HWND, point: POINT) -> bool {
    is_our_own_window(window)
        || (!window.is_invalid()
            && crate::app::seam::the_pin_of_the_window(window.0 as u64).is_some())
        || crate::app::seam::preview_window_rect().is_some_and(|(left, top, right, bottom)| {
            point.x >= left && point.x < right && point.y >= top && point.y < bottom
        })
}

/// Whether a window of ours is standing to be put anywhere: a pin's own, or the window a pin detached
/// into.
///
/// Asked of what the preview loop publishes every tick rather than of the pin's state, because the press
/// hook asks it and may take no lock. And the published answer is the one a window of its own has: a
/// detached window is no pin of the run's any more — `PIN_UP` is cleared with the detach — while it is
/// still a window of ours that owes a press elsewhere its place (see `PIN_KEYS_OWNER`).
pub(crate) fn a_window_of_ours_stands() -> bool {
    crate::app::seam::PIN_KEYS_OWNER.load(Ordering::Acquire) != 0
}

/// Whether that window stands above every other one: the setting a press must not overrule (see
/// `PIN_KEYS_TOPMOST`).
pub(crate) fn our_window_stands_over_everything() -> bool {
    crate::app::seam::PIN_KEYS_TOPMOST.load(Ordering::Acquire)
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

/// Whether `window` is the bubble's: the circle a collapsed pin leaves on screen, which is
/// the only window of the pin's own a press can land on while the pin's window is hidden (see
/// `pin_menu_dismiss_ask`).
///
/// **It carries the bubble's own menu too.** The window grows down to hold the panel beneath the
/// circle when a right-click opens it (see `open_bubble_menu`), so the window under a hand on one of
/// its rows is the bubble whether the panel is up or — on the tick after a row was pressed — the
/// listing the panel and the bubble were standing over (see `press_is_on_the_pins_own_menu`).
pub(super) fn bubble_window_is_at(window: HWND) -> bool {
    let bubble = crate::app::seam::the_bubble_window();
    window_is_the_slot(window, bubble)
}

/// Whether `window` is the window `slot` holds, where a slot of nothing holds no window: how the
/// handles this app keeps in process-wide slots are asked about, the hovers' and the pin's own (see
/// `PIN_BUBBLE_HWND` and `PIN_MENU_HWND`).
pub(super) fn window_is_the_slot(window: HWND, slot: isize) -> bool {
    !window.is_invalid() && slot != 0 && window.0 as isize == slot
}

/// Whether a press read on a pinned tick is one the Explorer hook publishes so the card's
/// menu is put away: a press that landed anywhere but the pin's own window. The pin window is
/// `WS_EX_NOACTIVATE` and holds no capture, so a press on another window never reaches its
/// procedure — the hook's poll is the only side that sees it — and this is the rule that
/// publishes it over the pin-end seam (see `request_pin_menu_dismiss`). On the pin's own
/// window the menu has its own answer (`pinned_menu_press`), so a press there is not
/// published. The answer is handed in rather than read here, because the rule that can only
/// be tested by putting a window on somebody's screen is a rule that goes untested (see
/// `preview_window_is_at`).
pub(super) fn press_is_outside_the_pin_window(pressed: bool, on_pin_window: bool) -> bool {
    pressed && !on_pin_window
}

/// The dismissal ask for one pinned tick: whether a press the tick's input
/// pass read puts the card's menu away. Both halves of the ask are facts of
/// the tick's one reading — the press from the pass, the window from the
/// pointer reading the tick opened with — so a press is judged against the
/// window as that reading named it, and never against a window a reading
/// made later in the tick would name: the first press on a window is the
/// press that activates it, and the two readings of one tick can name
/// different windows. The rule itself, with the window's identity handed in,
/// is `press_is_outside_the_pin_window`.
///
/// The bubble's own window and the pin menu's own window count as the pin's own here, and a
/// press on them is not published: a collapsed pin leaves no window on screen but the
/// bubble, the bubble carries the same menu a card does, and a press on either or on the
/// menu itself (e.g. clamped under the cursor near screen edges) is answered there rather
/// than putting the menu away (see `bubble_window_is_at`, `pin_menu_window_is_at`, and
/// `bubble_menu_press`).
pub(super) fn pin_menu_dismiss_ask(clicked: bool, tick_pointer: Option<&PointerTick>) -> bool {
    let on_the_pins_own_window = clicked
        && tick_pointer.is_some_and(|pointer| {
            preview_window_is_at(pointer.window)
                || bubble_window_is_at(pointer.window)
                || pin_menu_window_is_at(pointer.window)
        });

    press_is_outside_the_pin_window(clicked, on_the_pins_own_window)
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

/// Whether a press the hook read is one the pin's own buttons answered rather than the
/// listing the pin stands over: the buttons are the pin's, and the file drawn under one is
/// not one the user picked — the minimize that press asks for must not arrive with that file (see
/// `PinUpdateWatch::follow` and `pin_button_covers`).
pub(super) fn press_is_the_pins_chrome(pressed: bool, on_a_pin_button: bool) -> bool {
    pressed && on_a_pin_button
}

/// Whether a press was read on the very tick this app's own window moved: the pin's box changed
/// under a hand that did not move, so the point is over whatever the move uncovered (see
/// `PinUpdateWatch::note_pin_box`).
///
/// Such a press is the pin's own gesture rather than the listing's, and it is neither a pick nor a
/// no-pick: a resize, a maximize or a double-clicked picture moves the window off the point the hand
/// was on, the hit test there names the listing the pin has just moved across, and read as that
/// listing's the press picks no file — which lowers the pin behind the listing it moved over (see
/// `press_is_answered` and `pin_press_was_not_a_pick`). The point need not be inside the pin's new
/// box for this: the move is what refuses it, not the box (see `press_is_on_the_pin`).
///
/// The move is handed in rather than read here, because it is the same reading the watch's own
/// gesture state is dropped on (see `PinUpdateWatch::note_pin_box` and `the_pin_window_moved`).
pub(super) fn press_is_the_pins_own_gesture(pressed: bool, pin_moved: bool) -> bool {
    pressed && pin_moved
}

/// Whether a press is one the pin itself took: its point is inside the pin's own box, whatever the
/// window under that point turned out to be.
///
/// **The hit test cannot answer this one**, which is what the box is asked instead for: a pin's film
/// stands in a host that answers `HTTRANSPARENT` so the press reaches the pin's chrome, and a hit
/// test that skips the host names the *listing* the pin was fitted over (see `pin_film_host`). Read
/// as that listing's, a press on the pin's own picture is a press that picked no file — the point
/// holds the gap between the rows — and the pin is put back behind the listing for it (see
/// `pin_press_was_not_a_pick` and `lower_pinned_preview`).
///
/// The box is handed in rather than read here, because the watch already holds the pin's box and it
/// is the same reading its own move is measured against (see `PinUpdateWatch::note_pin_box`).
pub(super) fn press_is_on_the_pin(
    pressed: bool,
    point: POINT,
    pin_box: Option<(i32, i32, i32, i32)>,
) -> bool {
    pressed && pin_box.is_some_and(|pin_box| point_in_box(point, pin_box))
}

/// Whether a press is one the pin's own menu answered: either recorded by the menu's own press
/// handler, or read while the pointer is over the menu window (or was on the tick before).
pub(super) fn press_is_on_the_pins_menu(
    menu_pressed: bool,
    pressed: bool,
    window: HWND,
    previous_window: Option<HWND>,
) -> bool {
    menu_pressed
        || (pressed
            && (press_is_on_the_pins_own_menu(window)
                || previous_window.is_some_and(press_is_on_the_pins_own_menu)))
}

/// Whether a press is one the pin answers at all, whatever it turned out to name.
///
/// The pin owes the listing an answer for a press in that listing's own window — the frame, its view,
/// its toolbar and path bar and tabs, the menus it puts over its own rows — because every one of
/// those presses is either a file picked or a press that picked no file, and the two answers are
/// the pin staying in front and the pin going behind. A press on this app's own window is not the
/// listing's: that is the pin being used, and what it stands over does not change. A press that
/// destroyed the window it landed on names no place to go behind.
///
/// Separate from the pick rule because the two answer different questions: that one says whether a
/// file was picked, and this one says whether there is anything to answer. A press on the path bar
/// is no pick and is still answered — it was the press the pin was left in front of the listing by.
pub(super) fn press_is_answered(
    pressed: bool,
    over_explorer: bool,
    over_our_own: bool,
    aimed_at_a_gone_window: bool,
) -> bool {
    pressed && over_explorer && !over_our_own && !aimed_at_a_gone_window
}

/// Whether the rest a hand is making is a hover of the file under it: it is a rest at all only where
/// the setting asks for hovers, the hand has arrived, the delay a preview is given has passed since it
/// last moved, and this rest has not been answered already.
///
/// A rest is answered by the thing that read it. A pick named the file under the point, so there is
/// nothing left to ask. A press on a row of this app's own menu was answered by that row, and it left
/// the hand over the listing the row was covering — a file nobody chose. Asked again, either of them
/// answers with a file the user never came to rest on, and a pin put away by that press is brought
/// back on it by the update a hover offers (`bring_a_put_away_pin_back`). The answer holds until the
/// hand moves, which is when there is a new rest to ask about.
///
/// The menus are this app's own and not Explorer's: a popup of the shell's is the listing being
/// worked in, and what it leaves under the hand is a file the user may well mean to see.
pub(super) fn a_rest_is_a_hover(on_hover: bool, arrived: bool, spent: bool, settled: bool) -> bool {
    on_hover && arrived && !spent && settled
}

/// Whether the window under the pointer is a menu or flyout rather than anything a listing can
/// be read under: the popups Explorer puts over its own files, and the one panel this app puts
/// over them.
///
/// A popup is Explorer's own as far as `is_cursor_over_explorer_full` is concerned — its owner
/// chain reaches the frame — so a press on it reads as a press on the listing. Two answers say
/// otherwise, and either one is enough.
///
/// The class name is the first: every popup of Explorer's carries a substring of its own kind —
/// the XAML shell hosts them in `Microsoft.UI.Content.PopupWindowSiteBridge` and
/// `Microsoft.UI.Content.DesktopChildSiteBridge`, and a classic menu window is `#32768`. The
/// three bare words are the older shells' naming, and none of the windows a listing is read
/// under has one: the list is `SysListView32` under `DirectUIHWND`, the frame and the view are
/// `ExplorerWClass` and `CabinetWClass`, and this app's own window is `RustHoverPreviewWindow`.
///
/// The second is the window itself, for the two panels this app puts over a listing: the menu a
/// pinned window draws in a window of its own, and the menu a collapsed pin draws in its bubble (see
/// `press_is_on_the_pins_own_menu`). They are asked about by their handles rather than by their
/// names, so that the pin is free to rename its own windows without quietly making a press on a row
/// of a menu a press on the file behind the panel. That press is the pin's own while the menu is up
/// — the window under it is the panel, which is no listing — but the press is read on whichever tick
/// the shell spends the press bit on, and the tick after the menu closes finds the point reading as
/// the listing the panel was covering. The tick before is asked about for exactly that, and this
/// rule has to answer for the panel for it to be of any use.
///
/// Read per call rather than cached: up to two `GetClassNameW` calls on a pinned tick, and none at
/// all where one of the app's own panels is what answered.
pub(super) fn window_is_menu_popup_chrome(window: HWND) -> bool {
    press_is_on_the_pins_own_menu(window) || class_is_menu_popup_chrome(&window_class_of(window))
}

/// Whether a press on `window` is a press on one of this app's own menus rather than on the file the
/// panel covers: the panel a pinned window draws in a window of its own, and the panel a collapsed
/// pin draws in its bubble (see `pin_menu_window_is_at` and `bubble_window_is_at`).
///
/// These are the two panels of this app that answer a press with a row rather than with the listing
/// underneath, and answering one leaves the hand resting over the file the row covered — a file
/// nobody chose. So the rest of that hand is spent, and is not read as a hover of that file (see
/// `a_rest_is_a_hover`). A popup of Explorer's is deliberately not one of them: it is the listing
/// being worked in, and what it leaves under the hand is a file the user may well mean to see.
pub(super) fn press_is_on_the_pins_own_menu(window: HWND) -> bool {
    pin_menu_window_is_at(window) || bubble_window_is_at(window)
}

/// Whether `window` is a standing pin's own menu: the panel a pinned window draws in a window of
/// its own over the listing (see `pin_menu_window`).
///
/// A window where a press is not a press on the file behind it: the card holds the pointer while
/// its menu is up, and a press the hook reads off the panel is a row rather than the file that row
/// covers (see `press_is_on_the_pins_own_menu`, which the bubble's own panel answers beside it).
pub(super) fn pin_menu_window_is_at(window: HWND) -> bool {
    let menu = crate::app::seam::the_menu_window();
    window_is_the_slot(window, menu)
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
/// on somebody's screen is a rule that goes untested. `foreground_is_a_listing` is handed in for the
/// same reason: it is read once, beside the walk that fills `counts`.
///
/// **A listing in front is answered before the counts at all**, and it is the row the pointer is read
/// on: Explorer's folder window, or a Common Item Dialog some other program put up. Neither number
/// describes that arrangement — a dialog is no Explorer window, and a run with none open is one whose
/// counts would say there is nothing to look at while the user is picking a file (see
/// `a_listing_is_previewed` and `get_explorer_state`).
///
/// A foreground that is no listing leaves the rows it had, because what it is instead is not something
/// this loop can act on: where Explorer is reachable beside it the pointer may still cross to that
/// window, so the row stays `VisibleNotFocused`, and `HiddenByForeground` is already the row that asks
/// nothing of the pointer — a modal dialog of another program's covers nothing (see
/// `foreground_cover_rect`), so the state it reaches is the one the pointer can still be right in.
pub(super) fn explorer_state_from_counts(
    counts: &ExplorerWindowCounts,
    pin_up: bool,
    foreground_is_a_listing: bool,
) -> ExplorerState {
    if foreground_is_a_listing {
        return ExplorerState::ActiveFocus;
    }

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
/// Whether an Explorer window is left reachable is the whole of what an engine set to
/// `Until away` is let go by (see `app::afk`), and the answer is one this loop already
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

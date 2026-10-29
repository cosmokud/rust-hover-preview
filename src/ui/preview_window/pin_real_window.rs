//! The pin's own window: a window of this app's that Windows treats as a window, for a pin that is
//! to be one.
//!
//! **Why it is a second window rather than the preview's.** The window a hover is drawn in is
//! created as a non-activating tool window, and that is not a defect to be fixed later: a hover
//! appears under a pointer that may be in the middle of renaming a file, and a window that took the
//! focus would swallow the name being typed. The pin is the opposite case — the hand pressed it, the
//! user is in it, and it is meant to be raised, listed, minimized and restored the way every other
//! window is. What Windows decides from the styles of a window (does it hold a taskbar button, is it
//! in Alt+Tab, may a click activate it) is asked of a window once, when it is made, so the two uses
//! cannot share one window: a pin made out of the hover's window is a pin that Windows will never
//! raise on a click, which is the whole reason this app once had to raise it with the topmost band
//! and take that band off again.
//!
//! So a pin that is to be a window gets this one: not a tool window, not non-activating, layered
//! because its edges are alpha, and no band in its styles at all — the setting asks for the band when
//! the pin is shown, and asks for the ordinary band when it is not (see `pin_lifecycle::band_place`
//! for the band the setting names).
//!
//! What is *not* here is as much the point: no message of its own for presses, because Windows
//! activates the window a press lands on and a press on this window is a press on the pin; no
//! stand-in in the taskbar, because this window brings its own button; no mirror, because the Shell
//! draws a thumbnail of the window itself. A pin that is a window needs none of that, and the first
//! ticket of this module is the window alone so that the rest can be proven one kind at a time.

use std::sync::atomic::{AtomicIsize, Ordering};
use std::sync::Once;

use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, GetWindowLongPtrW, IsWindowVisible, MoveWindow, RegisterClassExW,
    SendMessageW, SetWindowLongPtrW, SetWindowPos, ShowWindow, GWL_EXSTYLE, ICON_BIG, ICON_SMALL,
    SW_HIDE, SW_SHOWNOACTIVATE, SWP_FRAMECHANGED, SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOSIZE,
    SWP_NOZORDER, WINDOW_EX_STYLE, WINDOW_STYLE, WNDCLASSEXW, WS_CLIPCHILDREN, WS_EX_APPWINDOW,
    WS_EX_LAYERED, WS_EX_TOOLWINDOW, WS_MINIMIZEBOX, WS_POPUP, WS_SYSMENU, WM_SETICON,
};

use super::pin_lifecycle;
use super::window_proc;
use super::ScreenRegion;

/// The class the pin's own window is made from, kept apart from the preview's so that a window the
/// Shell is asked about is a window of the pin's rather than a hover's.
const REAL_PIN_CLASS: PCWSTR = w!("RustHoverPreviewPin");

/// What the window says on the taskbar while it has no file to name. A pin names its own file as soon
/// as it is up (see `pin_taskbar::set_pin_taskbar_title`, which is the same name written here).
const REAL_PIN_TITLE: PCWSTR = w!("Rust Preview");

/// The styles a pin's own window is made with.
///
/// `WS_POPUP`: the pin's chrome is drawn by this app, not by Windows, so there is no caption of the
/// platform's to wear around the picture. `WS_SYSMENU | WS_MINIMIZEBOX`: the Shell's own menu on the
/// window's taskbar button — Close, Minimize — is a gesture on a window the user is meant to be able
/// to treat like any other, and a window without them wears a button that refuses both.
///
/// `WS_EX_LAYERED`: the pin's corners and the shadow around its chrome are alpha, and a layered
/// window is a window in every other respect, including activation (see `real_pin_style`).
///
/// **`WS_EX_APPWINDOW` or `WS_EX_TOOLWINDOW`, and the setting picks.** With `Pin Mode → Always On Top`
/// off the pin is an ordinary window of the desktop's and the button is its own, asked for by the
/// window rather than inferred from whatever owns it — which is the bit that makes the pin's button
/// survive a film's host going down for a swap, and the reason the pair no longer hands one between
/// them (see `set_the_pin_window_listed`). With the setting on the pin stands over every program and
/// owes no button at all, so it wears the tool bit instead: not in the taskbar, not in Alt+Tab.
///
/// **Never both, and never neither.** A window with neither bit is listed by the ownerless-window
/// rule, which is the button a pin over everything must not have.
/// `WS_CLIPCHILDREN`: the chrome's strips are this window's children on every road (see
/// `pin_chrome_strips`). Without this bit this window's paint covers the client area whole, children
/// included, so the chrome would be wiped over at the pin's own repaint cadence — measured: the band
/// showed the pin's flat colour with nothing of the chrome in it. With it, the paint stops at the
/// strips and they are left standing.
///
/// **What is missing is the answer.** `WS_EX_NOACTIVATE` is what refuses the window the focus;
/// `WS_EX_TOPMOST` is a band, and the band is the setting's (see `band_place`).
///
/// `on_the_taskbar` is the place the setting gives the pin in the Shell's lists (see the notes above).
fn real_pin_style(on_the_taskbar: bool) -> (WINDOW_STYLE, WINDOW_EX_STYLE) {
    (
        WS_POPUP | WS_SYSMENU | WS_MINIMIZEBOX | WS_CLIPCHILDREN,
        WS_EX_LAYERED | listed_bit(on_the_taskbar),
    )
}

/// The bit that puts a pin in the Shell's lists, or keeps it out of them: `WS_EX_APPWINDOW` for a pin
/// that owes a taskbar button, `WS_EX_TOOLWINDOW` for one that does not.
fn listed_bit(on_the_taskbar: bool) -> WINDOW_EX_STYLE {
    if on_the_taskbar {
        WS_EX_APPWINDOW
    } else {
        WS_EX_TOOLWINDOW
    }
}

/// Put the pin in the place the setting gives it in the Shell's lists, or take it out of them: the one
/// road that writes `WS_EX_APPWINDOW`/`WS_EX_TOOLWINDOW` on the window that wears the pin's button.
///
/// **The button is the pin's own, and the setting is the whole of the question** (see `real_pin_style`
/// and `pin_film_host::the_button_follows_the_setting`). Nothing of the pair hands a bit to the other:
/// the film's host is a tool window with no button whether it is up or down, so a swap — which hides
/// the host for the wait of the next player — cannot take the pin's button away, and the pin itself is
/// never hidden by this road while it stands.
///
/// **A button is the Shell's answer to a window being *shown*, and a bit written behind a visible
/// window's back is not re-read for it.** Measured: a window that stays up wears the button it was
/// shown with, however the bit is re-read — which is why the pair's button was intermittent. So a pin
/// that is on the screen when its place changes pays exactly one hide and one show: down, the two bits
/// flipped, the style re-read, and up again (see `ask_the_window_to_re_read_its_style`). A pin that is
/// down is left alone: the show it is owed is the one that brings it, and that show reads the bits
/// written here.
///
/// The two bits are written together and never separately: a window with neither is listed by the
/// ownerless-window rule, which is the button a pin over everything must not have.
pub(super) unsafe fn set_the_pin_window_listed(hwnd: HWND, on_the_taskbar: bool) {
    let style = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) };
    let wanted = the_listed_style(style, on_the_taskbar);
    if wanted == style {
        return;
    }

    // **Down before the bits change, and up after the style has been re-read.** The hide is what makes
    // the show that follows the one the Shell reads the bits with.
    let visible = unsafe { IsWindowVisible(hwnd) }.as_bool();
    if visible {
        let _ = unsafe { ShowWindow(hwnd, SW_HIDE) };
    }

    let _ = unsafe { SetWindowLongPtrW(hwnd, GWL_EXSTYLE, wanted) };
    unsafe { ask_the_window_to_re_read_its_style(hwnd) };

    if visible {
        #[cfg(test)]
        if let Ok(mut shown) = APPWINDOW_SHOWS.lock() {
            shown.push(hwnd.0 as isize);
        }
        let _ = unsafe { ShowWindow(hwnd, SW_SHOWNOACTIVATE) };
    }
}

/// The extended style a pin of `style` wears for the place the setting gives it: the app-window bit for
/// a pin that owes a button, the tool bit for one that does not.
fn the_listed_style(style: isize, on_the_taskbar: bool) -> isize {
    let app_window = WS_EX_APPWINDOW.0 as isize;
    let tool_window = WS_EX_TOOLWINDOW.0 as isize;
    if on_the_taskbar {
        (style | app_window) & !tool_window
    } else {
        (style | tool_window) & !app_window
    }
}

/// The windows this run has put down and up again so that the Shell reads a taskbar bit: the one
/// thing this road does that a style write behind a visible window's back cannot, and what a test
/// holds a road to where the Shell itself cannot be asked (see `set_the_pin_window_appwindow`).
#[cfg(test)]
static APPWINDOW_SHOWS: std::sync::Mutex<Vec<isize>> = std::sync::Mutex::new(Vec::new());

/// The windows whose taskbar bit this run has paid a show for (see `APPWINDOW_SHOWS`).
#[cfg(test)]
pub(super) fn the_appwindow_shows() -> Vec<isize> {
    APPWINDOW_SHOWS
        .lock()
        .map(|shown| shown.clone())
        .unwrap_or_default()
}

/// Forget the shows paid before a test (see `APPWINDOW_SHOWS`).
#[cfg(test)]
pub(super) fn forget_the_appwindow_shows() {
    if let Ok(mut shown) = APPWINDOW_SHOWS.lock() {
        shown.clear();
    }
}

/// A style is not a thing a window is told about live: without being asked to re-read it, a window
/// keeps presenting the surface it was made with, and the Shell goes on listing it as it was.
unsafe fn ask_the_window_to_re_read_its_style(hwnd: HWND) {
    let _ = unsafe {
        SetWindowPos(
            hwnd,
            None,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
        )
    };
}

/// The pin's own window, where one has been made.
static REAL_PIN_HWND: AtomicIsize = AtomicIsize::new(0);

static REAL_PIN_CLASS_REGISTERED: Once = Once::new();

/// The class is registered the first time a pin asks for a window of its own, rather than at the
/// loop's start: a run that never pins a picture has no reason to have one, and a class registered
/// on the road that needs it cannot be out of step with the styles its window is made with.
unsafe fn register_real_pin_class() {
    REAL_PIN_CLASS_REGISTERED.call_once(|| {
        let hinstance = GetModuleHandleW(None).expect("the module's own handle");
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(window_proc),
            hInstance: hinstance.into(),
            lpszClassName: REAL_PIN_CLASS,
            ..Default::default()
        };
        RegisterClassExW(&class);
    });
}

/// Stand a window of the test's own as the pin's own window: what a test needs to ask a road which
/// window it acts on without putting a real pin on the user's desktop.
#[cfg(test)]
pub(super) fn stand_real_pin_window_for_a_test(hwnd: Option<HWND>) {
    REAL_PIN_HWND.store(hwnd.map_or(0, |hwnd| hwnd.0 as isize), Ordering::SeqCst);
}

/// The pin's own window, or nothing where it has not been made: asked by a road that only wants to
/// know whether there is one.
pub(super) fn real_pin_window() -> Option<HWND> {
    let hwnd = REAL_PIN_HWND.load(Ordering::SeqCst);
    (hwnd != 0).then_some(HWND(hwnd as *mut _))
}

/// The pin's own window, made where there is none and put at the box the pin says, off the screen.
///
/// The window is made hidden and at the pin's box, so that the first show paints into a window that
/// is already the right size: a layered window shown before its first paint is a window with nothing
/// in it, and what it would show is the surface it does not have (see `show_pinned_window` for the
/// same bargain made by the preview's window).
pub(super) unsafe fn real_pin_window_at(window: ScreenRegion) -> Option<HWND> {
    let width = (window.2 - window.0).max(1);
    let height = (window.3 - window.1).max(1);

    if let Some(hwnd) = real_pin_window() {
        let _ = MoveWindow(hwnd, window.0, window.1, width, height, false);
        // **A window kept between pins is given the place the setting names now**, before the show it is
        // about to be given: the row may have been changed while it was down, and a pin that comes up
        // wearing the place of the pin before it is a window with a button it does not owe — or without
        // the one it does (see `set_the_pin_window_listed`). A window that is down pays no show for it.
        unsafe { set_the_pin_window_listed(hwnd, !pin_lifecycle::pin_always_on_top()) };
        // Kept with the window rather than made once: a take-up that finds them gone — a test
        // that stood the window without them — stands them again (see `take_the_strips_up`).
        unsafe { super::pin_chrome_strips::take_the_strips_up(hwnd) };
        return Some(hwnd);
    }

    register_real_pin_class();
    let (style, ex_style) = real_pin_style(!pin_lifecycle::pin_always_on_top());
    let hinstance = GetModuleHandleW(None).ok()?;
    let hwnd = CreateWindowExW(
        ex_style,
        REAL_PIN_CLASS,
        REAL_PIN_TITLE,
        style,
        window.0,
        window.1,
        width,
        height,
        None,
        None,
        hinstance,
        None,
    )
    .ok()?;

    REAL_PIN_HWND.store(hwnd.0 as isize, Ordering::SeqCst);
    // **The icon and the stills a taskbar button is drawn from are written before the first show**, so
    // that the button a pin appears with is the app's own from its first frame rather than the Shell's
    // fallback for a window that has neither (see `set_the_pins_icon` and
    // `pin_taskbar::supply_the_stills_of_this_window`).
    unsafe { set_the_pins_icon(hwnd) };
    unsafe { super::pin_taskbar::supply_the_stills_of_this_window(hwnd) };
    // The chrome the pin is carried in is made with the window that carries it, on every road —
    // a hosted film, a park, a cover, FFmpeg's player, or this app's own picture — and kept for
    // the window's life (see `pin_chrome_strips`).
    unsafe { super::pin_chrome_strips::take_the_strips_up(hwnd) };
    Some(hwnd)
}

/// Give the pin's window the app's own icon, so that the Shell draws the button this window wears with
/// the app's picture rather than the one it falls back on for a window with none.
///
/// Asked before the window's first show, and of both sizes the Shell asks a window for: the taskbar's
/// button and the Alt+Tab frame are the small one, the file's own picture beside a hover the large
/// (see `app_icon`).
unsafe fn set_the_pins_icon(hwnd: HWND) {
    let icon = unsafe { crate::shell::app_icon::app_icon() };
    if icon.0.is_null() {
        return;
    }

    for which in [ICON_BIG, ICON_SMALL] {
        unsafe {
            SendMessageW(
                hwnd,
                WM_SETICON,
                WPARAM(which as usize),
                LPARAM(icon.0 as isize),
            )
        };
    }
}

/// Take the pin's own window down: hidden but kept, the way the preview's window is kept between
/// hovers. A window that is kept holds its class, its styles and its place, and a pin taken up again
/// is the same window rather than one made and shown in a frame.
pub(super) unsafe fn hide_real_pin_window() {
    if let Some(hwnd) = real_pin_window() {
        let _ = ShowWindow(hwnd, SW_HIDE);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows::Win32::Foundation::LRESULT;
    use windows::Win32::UI::WindowsAndMessaging::{
        DestroyWindow, GetWindowLongPtrW, IsIconic, GWL_EXSTYLE, WM_GETICON, WS_EX_NOACTIVATE,
        WS_EX_TOOLWINDOW, WS_EX_TOPMOST,
    };

    /// The app's own icon, as a handle a `WM_GETICON` answer is compared against (see
    /// `set_the_pins_icon`).
    fn icon_from_the_app() -> isize {
        unsafe { crate::shell::app_icon::app_icon() }.0 as isize
    }

    /// What makes a window a window as far as the Shell is concerned, and what makes it not one: the
    /// styles are the whole answer, and they are the reason a pin cannot be made out of the hover's
    /// window after the fact.
    ///
    /// **The two listed bits are the setting's, and a pin wears one of them and never both.** A pin that
    /// is not over everything is an ordinary window with a button of its own; a pin that is owes no
    /// button and is kept out of the Shell's lists entirely (see `real_pin_style`).
    #[test]
    fn the_pins_own_window_is_listed_and_may_be_activated() {
        for on_the_taskbar in [true, false] {
            let (style, ex) = real_pin_style(on_the_taskbar);

            assert_eq!(
                style & WS_POPUP,
                WS_POPUP,
                "the pin's chrome is this app's, so there is nothing of the platform's to wear"
            );
            assert_eq!(
                style & WS_CLIPCHILDREN,
                WS_CLIPCHILDREN,
                "the chrome's strips are this window's children, and an unclipped paint wipes them out"
            );
            assert_eq!(
                ex & WS_EX_LAYERED,
                WS_EX_LAYERED,
                "the pin's edges are alpha"
            );
            assert_eq!(
                ex & WS_EX_APPWINDOW,
                if on_the_taskbar {
                    WS_EX_APPWINDOW
                } else {
                    WINDOW_EX_STYLE::default()
                },
                "the button is the window's own where the setting says it owes one"
            );
            assert_eq!(
                ex & WS_EX_TOOLWINDOW,
                if on_the_taskbar {
                    WINDOW_EX_STYLE::default()
                } else {
                    WS_EX_TOOLWINDOW
                },
                "and a pin over everything is out of Alt+Tab and out of the taskbar"
            );
            assert_ne!(
                ex & (WS_EX_APPWINDOW | WS_EX_TOOLWINDOW),
                WINDOW_EX_STYLE::default(),
                "neither bit is a window the ownerless rule would list anyway"
            );
            assert_eq!(
                ex & WS_EX_NOACTIVATE,
                WINDOW_EX_STYLE::default(),
                "a pin may be activated by the press that lands on it"
            );
            assert_eq!(
                ex & WS_EX_TOPMOST,
                WINDOW_EX_STYLE::default(),
                "the band is the setting's to put on, not the style's to start in"
            );
        }
    }

    /// A pin that is taken down is a window taken off the screen, and the style a take-down writes is
    /// the hover's: `WS_EX_NOACTIVATE` is what keeps a preview from taking the caret out of a name being
    /// typed, and it is the one thing a pin's own window must never wear. The window is kept between
    /// pins rather than made again (see `hide_real_pin_window`), so a take-down that stamped it left
    /// every pin of the run after the first refusing the very press that is meant to raise it: the style
    /// is asked of the window rather than of the asker, and a window Windows will not activate is a
    /// window no click brings to the front.
    #[test]
    fn a_taken_down_pin_leaves_its_own_window_one_a_click_may_activate() {
        use super::super::pin_model::pin_set_focusable;

        let _one = super::super::pin_window::PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let box_ = (-20000, -20000, -19680, -19760);
        let hwnd = unsafe { real_pin_window_at(box_) }.expect("the pin's own window");

        let keeps_out_of_the_keyboard = |hwnd: HWND| unsafe {
            let ex = GetWindowLongPtrW(hwnd, GWL_EXSTYLE) as u32;
            WINDOW_EX_STYLE(ex) & WS_EX_NOACTIVATE
        };

        assert_eq!(keeps_out_of_the_keyboard(hwnd), WINDOW_EX_STYLE::default());

        // What `end_pin` asks of the window it takes down: the style goes on and the window is kept.
        unsafe { pin_set_focusable(hwnd, false) };
        assert_eq!(
            keeps_out_of_the_keyboard(hwnd),
            WINDOW_EX_STYLE::default(),
            "the pin's own window is the Shell's to activate, so a take-down leaves it as it was made"
        );

        // And the style is still the hover's to ask for: the same window, stood as no pin's at all,
        // wears it, which is what makes the answer above about the window rather than about the call.
        stand_real_pin_window_for_a_test(None);
        unsafe { pin_set_focusable(hwnd, false) };
        assert_ne!(
            keeps_out_of_the_keyboard(hwnd),
            WINDOW_EX_STYLE::default(),
            "a window that is no pin's own is the hover's to keep out of the keyboard"
        );

        unsafe {
            let _ = DestroyWindow(hwnd);
        }
        REAL_PIN_HWND.store(0, Ordering::SeqCst);
    }

    /// The styles above are the styles of the window that is actually made: a plain `STATIC` window
    /// created with them would prove the constants and nothing about this module. A pin with no
    /// hosted FFmpeg film over it is the layered window it always was.
    ///
    /// **And the window the module makes wears the listed bit the setting names at the moment it is
    /// made**, which is where the pin's button comes from on the very first pin of a run (see
    /// `real_pin_window_at`).
    #[test]
    fn the_window_the_module_makes_wears_them() {
        // The window is the process's one pin window and the slot it stands in is cleared at the end:
        // another test holding the lock while it asks for the same window is a test handed a handle
        // this one took away, so the two are one at a time like every other stand of a pin
        // (see `PIN_TESTS_ONE_AT_A_TIME`).
        let _one = super::super::pin_window::PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let previous = super::super::pin_lifecycle::pin_always_on_top();
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.pin_always_on_top = false;
        }

        let box_ = (-20000, -20000, -19680, -19760);
        let hwnd = unsafe { real_pin_window_at(box_) }.expect("the pin's own window");

        let ex = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32;
        let ex = WINDOW_EX_STYLE(ex);

        assert_eq!(
            ex & WS_EX_APPWINDOW,
            WS_EX_APPWINDOW,
            "a pin with the setting off is an ordinary window with a button"
        );
        assert_eq!(ex & WS_EX_LAYERED, WS_EX_LAYERED);
        assert_eq!(ex & WS_EX_TOOLWINDOW, WINDOW_EX_STYLE::default());
        assert_eq!(ex & WS_EX_NOACTIVATE, WINDOW_EX_STYLE::default());
        assert_eq!(
            unsafe { SendMessageW(hwnd, WM_GETICON, WPARAM(ICON_BIG as usize), LPARAM(0)) },
            LRESULT(icon_from_the_app()),
            "the button the window wears is drawn with the app's own icon from its first frame"
        );
        assert_eq!(
            unsafe { SendMessageW(hwnd, WM_GETICON, WPARAM(ICON_SMALL as usize), LPARAM(0)) },
            LRESULT(icon_from_the_app()),
            "and the small one the Shell puts beside a hover is the same icon"
        );

        unsafe {
            let _ = DestroyWindow(hwnd);
        }
        REAL_PIN_HWND.store(0, Ordering::SeqCst);

        // And the other way about: the same road with the setting on makes a window that is out of the
        // Shell's lists and asks for no button at all.
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.pin_always_on_top = true;
        }
        let hwnd = unsafe { real_pin_window_at(box_) }.expect("the pin's own window");
        let ex = WINDOW_EX_STYLE(unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32);
        assert_eq!(
            ex & WS_EX_APPWINDOW,
            WINDOW_EX_STYLE::default(),
            "a pin that is Always On Top owes no button"
        );
        assert_eq!(ex & WS_EX_TOOLWINDOW, WS_EX_TOOLWINDOW);

        unsafe {
            let _ = DestroyWindow(hwnd);
        }
        REAL_PIN_HWND.store(0, Ordering::SeqCst);
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.pin_always_on_top = previous;
        }
    }

    /// **The setting's row pays one hide and one show, and the two listed bits move together.**
    ///
    /// A button is the Shell's answer to a window being shown, so a pin that is already on the screen
    /// has to be put down and up again for a row that changes what it ought to wear — exactly once, and
    /// with the bits written in between (see `set_the_pin_window_listed`). A pin that is down pays
    /// nothing: the show it is owed is the one that brings it, and that show reads the bits already
    /// written.
    #[test]
    fn the_setting_pays_one_hide_and_one_show_and_never_both_bits() {
        let _one = super::super::pin_window::PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let box_ = (-20000, -20000, -19680, -19760);
        let hwnd = unsafe { real_pin_window_at(box_) }.expect("the pin's own window");
        let listed = |hwnd: HWND| WINDOW_EX_STYLE(unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32);

        // A pin that is not on the screen is never put down and up for a bit: the show it is owed is
        // the one that brings it.
        forget_the_appwindow_shows();
        unsafe { set_the_pin_window_listed(hwnd, true) };
        assert_eq!(
            listed(hwnd) & WS_EX_APPWINDOW,
            WS_EX_APPWINDOW,
            "a pin that owes a button asks for one"
        );
        assert_eq!(
            listed(hwnd) & WS_EX_TOOLWINDOW,
            WINDOW_EX_STYLE::default(),
            "and is not a tool window at the same time"
        );
        assert!(
            the_appwindow_shows().is_empty(),
            "a window that is down is left alone: the show it is owed is the one that brings it"
        );

        // Up, and then the row: one hide and one show, and the two bits swapped together.
        let _ = unsafe { ShowWindow(hwnd, SW_SHOWNOACTIVATE) };
        forget_the_appwindow_shows();
        unsafe { set_the_pin_window_listed(hwnd, false) };
        assert_eq!(
            listed(hwnd) & WS_EX_APPWINDOW,
            WINDOW_EX_STYLE::default(),
            "a pin that is Always On Top gives the bit up"
        );
        assert_eq!(
            listed(hwnd) & WS_EX_TOOLWINDOW,
            WS_EX_TOOLWINDOW,
            "and takes the tool bit in the same pass"
        );
        assert_eq!(
            the_appwindow_shows(),
            vec![hwnd.0 as isize],
            "a visible pin pays exactly one show for the row"
        );
        assert!(
            unsafe { IsWindowVisible(hwnd) }.as_bool(),
            "and it is put straight back up, which is the show the Shell reads the bits with"
        );
        assert!(
            !unsafe { IsIconic(hwnd) }.as_bool(),
            "a row that changes a bit neither minimizes nor hides the pin the user is looking at"
        );

        // A row that asks for the place the window is already in asks for nothing at all.
        forget_the_appwindow_shows();
        unsafe { set_the_pin_window_listed(hwnd, false) };
        assert!(
            the_appwindow_shows().is_empty(),
            "a bit that has not moved is no show and no flicker"
        );

        unsafe {
            let _ = DestroyWindow(hwnd);
        }
        REAL_PIN_HWND.store(0, Ordering::SeqCst);
    }
}

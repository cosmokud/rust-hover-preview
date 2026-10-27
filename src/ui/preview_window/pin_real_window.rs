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
use windows::Win32::Foundation::HWND;
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, GetWindowLongPtrW, MoveWindow, RegisterClassExW, SetWindowLongPtrW,
    SetWindowPos, ShowWindow, GWL_EXSTYLE, SW_HIDE, SWP_FRAMECHANGED, SWP_NOACTIVATE,
    SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER, WINDOW_EX_STYLE, WINDOW_STYLE, WNDCLASSEXW,
    WS_CLIPCHILDREN, WS_EX_APPWINDOW, WS_EX_LAYERED, WS_MINIMIZEBOX, WS_POPUP, WS_SYSMENU,
};

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
/// window is a window in every other respect, including activation. **A pin that holds a hosted
/// FFmpeg film drops it**, because that film is presented through the composed renderer into this
/// window's own DC rather than through a flip-model swapchain (see `real_pin_style`).
///
/// `WS_EX_APPWINDOW`: a taskbar button asked for by the window itself rather than inferred from what
/// owns it, which is the button this window owes and the one the stand-in used to fake.
///
/// `WS_CLIPCHILDREN`: a hosted FFmpeg film is a **child** window standing in this one, and the
/// chrome's strips are this one's children on every road (see `pin_film_host` and
/// `pin_chrome_strips`), and a pin that holds an FFmpeg film paints its own surface with `GDI`
/// rather than by `UpdateLayeredWindow` (see `paint::paints_through_the_window_dc`). Without this
/// bit that paint covers the client area whole, child included, so the film would be wiped over at
/// the pin's own repaint cadence — measured: the band showed the pin's flat colour with nothing of
/// the film in it. With it, the pin's own DC stops at the film's box and the film is left
/// standing.
///
/// **What is missing is the answer.** `WS_EX_TOOLWINDOW` is what keeps a window out of Alt+Tab and
/// out of the taskbar; `WS_EX_NOACTIVATE` is what refuses it the focus; `WS_EX_TOPMOST` is a band,
/// and the band is the setting's (see `band_place`).
fn real_pin_style(layered: bool) -> (WINDOW_STYLE, WINDOW_EX_STYLE) {
    (
        WS_POPUP | WS_SYSMENU | WS_MINIMIZEBOX | WS_CLIPCHILDREN,
        if layered {
            WS_EX_LAYERED | WS_EX_APPWINDOW
        } else {
            WS_EX_APPWINDOW
        },
    )
}

/// Put the pin's own window in the layered or the non-layered form its presentation asks for, now.
/// The window is reused across pins, so the form is written on every take-up: a hosted FFmpeg film
/// is composed through this window's own DC and must not wear `WS_EX_LAYERED`, and every other pin
/// keeps the layered surface it has always had (see `pin_film_host`).
pub(super) unsafe fn set_the_pin_window_layered(hwnd: HWND, layered: bool) {
    if unsafe { set_an_ex_style(hwnd, WS_EX_LAYERED, layered) } {
        unsafe { ask_the_window_to_re_read_its_style(hwnd) };
    }
}

/// Ask for or give back the pin's own taskbar button and Alt+Tab entry: what a pin whose chrome is
/// owned by its film host gives up, and what it takes back where no host owns it (see
/// `pin_film_host::own_the_pin_window`). A window's place in the Shell's lists is asked of the window
/// rather than of the asker, so it is written with the road and not with the setting.
pub(super) unsafe fn set_the_pin_window_appwindow(hwnd: HWND, app_window: bool) {
    if unsafe { set_an_ex_style(hwnd, WS_EX_APPWINDOW, app_window) } {
        unsafe { ask_the_window_to_re_read_its_style(hwnd) };
    }
}

/// Put a bit of the pin window's extended style on or off, answering whether it changed.
///
/// **A style of zero is a style, and not a window that has gone.** The pin's own window wears one
/// extended bit at a time on some roads — the app-window bit alone, the layered bit alone — so
/// clearing it leaves nothing at all, and an early answer for that case is a bit nothing ever puts
/// back (see `set_the_pin_window_appwindow`).
unsafe fn set_an_ex_style(hwnd: HWND, bit: WINDOW_EX_STYLE, on: bool) -> bool {
    let style = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) };
    let wanted = if on {
        style | bit.0 as isize
    } else {
        style & !(bit.0 as isize)
    };
    if wanted == style {
        return false;
    }

    let _ = unsafe { SetWindowLongPtrW(hwnd, GWL_EXSTYLE, wanted) };
    true
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

/// Write the form the pin now asks for, wherever its own window exists: asked on a take-up and on
/// a file swapped into a standing pin, so a film stepped to or away from changes presentation with
/// the file rather than waiting for the next show.
pub(super) unsafe fn apply_the_pin_window_presentation() {
    if let Some(hwnd) = real_pin_window() {
        let layered = !super::pin_film_host::the_pin_presents_a_hosted_film_composed();
        unsafe { set_the_pin_window_layered(hwnd, layered) };
    }
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
        // Kept with the window rather than made once: a take-up that finds them gone — a test
        // that stood the window without them — stands them again (see `take_the_strips_up`).
        unsafe { super::pin_chrome_strips::take_the_strips_up(hwnd) };
        return Some(hwnd);
    }

    register_real_pin_class();
    let (style, ex_style) =
        real_pin_style(!super::pin_film_host::the_pin_presents_a_hosted_film_composed());
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
    // The chrome the pin is carried in is made with the window that carries it, on every road —
    // a hosted film, a park, a cover, FFmpeg's player, or this app's own picture — and kept for
    // the window's life (see `pin_chrome_strips`).
    unsafe { super::pin_chrome_strips::take_the_strips_up(hwnd) };
    Some(hwnd)
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
    use windows::Win32::UI::WindowsAndMessaging::{
        DestroyWindow, GetWindowLongPtrW, GWL_EXSTYLE, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW,
        WS_EX_TOPMOST,
    };

    /// What makes a window a window as far as the Shell is concerned, and what makes it not one: the
    /// styles are the whole answer, and they are the reason a pin cannot be made out of the hover's
    /// window after the fact.
    #[test]
    fn the_pins_own_window_is_listed_and_may_be_activated() {
        let (style, ex) = real_pin_style(true);
        let (bare_style, bare) = real_pin_style(false);

        assert_eq!(
            style & WS_POPUP,
            WS_POPUP,
            "the pin's chrome is this app's, so there is nothing of the platform's to wear"
        );
        assert_eq!(
            style & WS_CLIPCHILDREN,
            WS_CLIPCHILDREN,
            "the film is a child of this window, and a composed paint that is not clipped over it \
             wipes the picture out at the pin's own repaint cadence"
        );
        assert_eq!(
            bare_style & WS_CLIPCHILDREN,
            WS_CLIPCHILDREN,
            "and the pin that composes a film through its own DC clips it the same way"
        );
        assert_eq!(
            ex & WS_EX_LAYERED,
            WS_EX_LAYERED,
            "the pin's edges are alpha"
        );
        assert_eq!(
            bare & WS_EX_LAYERED,
            WINDOW_EX_STYLE::default(),
            "but a pin that composes a hosted film through its own DC is not layered"
        );
        assert_eq!(
            bare & WS_EX_APPWINDOW,
            WS_EX_APPWINDOW,
            "and it is still the window the taskbar button belongs to"
        );
        assert_eq!(
            bare & WS_EX_TOOLWINDOW,
            WINDOW_EX_STYLE::default(),
            "still listed, still activatable, either way"
        );
        assert_eq!(
            ex & WS_EX_APPWINDOW,
            WS_EX_APPWINDOW,
            "the button is the window's own, not a stand-in's"
        );
        assert_eq!(
            ex & WS_EX_TOOLWINDOW,
            WINDOW_EX_STYLE::default(),
            "a tool window is not in Alt+Tab, and a pin must be"
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
    #[test]
    fn the_window_the_module_makes_wears_them() {
        // The window is the process's one pin window and the slot it stands in is cleared at the end:
        // another test holding the lock while it asks for the same window is a test handed a handle
        // this one took away, so the two are one at a time like every other stand of a pin
        // (see `PIN_TESTS_ONE_AT_A_TIME`).
        let _one = super::super::pin_window::PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let box_ = (-20000, -20000, -19680, -19760);
        let hwnd = unsafe { real_pin_window_at(box_) }.expect("the pin's own window");

        let ex = unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) } as u32;
        let ex = WINDOW_EX_STYLE(ex);

        assert_eq!(ex & WS_EX_APPWINDOW, WS_EX_APPWINDOW);
        assert_eq!(ex & WS_EX_LAYERED, WS_EX_LAYERED);
        assert_eq!(ex & WS_EX_TOOLWINDOW, WINDOW_EX_STYLE::default());
        assert_eq!(ex & WS_EX_NOACTIVATE, WINDOW_EX_STYLE::default());

        unsafe {
            let _ = DestroyWindow(hwnd);
        }
        REAL_PIN_HWND.store(0, Ordering::SeqCst);
    }
}

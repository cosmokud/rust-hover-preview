//! The engine's page, nested in the pin's own window, taking the presses that land on its band.
//!
//! A layered window is hit-tested on its composed surface before its children, so a band painted
//! transparent is a band no child of the pin is reachable through: the pin claims the band with the
//! least coverage a pixel can have while the page really stands in it, and hands it what lands there
//! (see `paint::film_band_is_hit_testable`). The page is a window of the browser's, on another
//! thread, so nothing arrives at it by itself.
//!
//! Coordinates are the message's own: a press and a move carry the pin's client point, a wheel
//! carries the screen's, and a cursor message carries neither. What is sent on is the target's
//! client point, mapped by Windows rather than by arithmetic so a display scale is not this file's
//! business.

use windows::Win32::Foundation::{HWND, LPARAM, POINT, WPARAM};
use windows::Win32::Graphics::Gdi::MapWindowPoints;
use windows::Win32::System::Threading::{AttachThreadInput, GetCurrentThreadId};
use windows::Win32::UI::Input::KeyboardAndMouse::{ReleaseCapture, SetCapture, SetFocus};
use windows::Win32::UI::WindowsAndMessaging::{
    GetCursorPos, GetWindowThreadProcessId, IsChild, IsWindow, SendMessageW, WindowFromPoint,
    HTCLIENT, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MBUTTONDOWN, WM_MBUTTONUP, WM_MOUSEHWHEEL,
    WM_MOUSEMOVE, WM_MOUSEWHEEL, WM_RBUTTONDOWN, WM_RBUTTONUP, WM_SETCURSOR,
};

use super::media_types::MediaType;
use crate::engines::webview_preview;

/// Whether the page's band is the pin's own pixel: only the engine's two kinds, and only while the
/// page really stands nested in the pin. A claim on a band with nothing in it is a band that eats
/// the presses of whatever stands behind the pin.
pub(super) fn page_claims_the_band(media: MediaType, nested: bool) -> bool {
    nested && matches!(media, MediaType::EngineSvg | MediaType::EngineFont)
}

/// Whether a screen point is on a band, which is a box's left, top, width and height.
pub(super) fn band_holds(band: (i32, i32, i32, i32), point: (i32, i32)) -> bool {
    let (x, y, width, height) = band;
    point.0 >= x && point.0 < x + width && point.1 >= y && point.1 < y + height
}

/// Whether a mouse message is one the pin hands to the page rather than its own.
pub(super) fn carried(msg: u32) -> bool {
    matches!(
        msg,
        WM_LBUTTONDOWN
            | WM_LBUTTONUP
            | WM_RBUTTONDOWN
            | WM_RBUTTONUP
            | WM_MBUTTONDOWN
            | WM_MBUTTONUP
            | WM_MOUSEMOVE
            | WM_MOUSEWHEEL
            | WM_MOUSEHWHEEL
            | WM_SETCURSOR
    )
}

/// Whether a message is a button going down, which is what makes the page the caret's.
pub(super) fn a_press(msg: u32) -> bool {
    matches!(msg, WM_LBUTTONDOWN | WM_RBUTTONDOWN | WM_MBUTTONDOWN)
}

/// Whether a message is a button coming back up, which gives the pointer back.
pub(super) fn a_release(msg: u32) -> bool {
    matches!(msg, WM_LBUTTONUP | WM_RBUTTONUP | WM_MBUTTONUP)
}

/// Whether a wheel over the nested page's band is the page's to take: a page scrolls, and the pin's
/// own zoom claim over the same box would swallow every notch (see `wheel_input`).
pub(crate) fn the_wheel_is_the_pages(point: (i32, i32)) -> bool {
    webview_preview::nested_page_band().is_some_and(|band| band_holds(band, point))
}

/// Hand `msg` to the page where it landed on the page's band, answering whether it was the page's.
///
/// The pin takes the pointer for a press, so the moves and the release of a drag that leaves the
/// band still arrive here and are still the page's: the page is what the press was for.
pub(super) unsafe fn forward(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> bool {
    if !carried(msg) {
        return false;
    }

    let Some(band) = webview_preview::nested_page_band() else {
        return false;
    };

    let page = HWND(webview_preview::showing_hwnd() as *mut _);
    if page.0.is_null() || !unsafe { IsWindow(page) }.as_bool() {
        return false;
    }

    let screen = match msg {
        WM_MOUSEWHEEL | WM_MOUSEHWHEEL => point_of(lparam),
        WM_SETCURSOR => {
            let mut point = POINT::default();
            if unsafe { GetCursorPos(&mut point) }.is_err() {
                return false;
            }
            (point.x, point.y)
        }
        _ => {
            let (x, y) = point_of(lparam);
            super::pin_input::client_to_screen(hwnd, x, y)
        }
    };

    if !band_holds(band, screen) {
        return false;
    }

    // The window under the point rather than the page's own: a page is drawn by a window of the
    // browser's inside the host, and that is the one a real press is delivered to. Anything that
    // is not the page or inside it is not a window to hand a press to.
    let target = unsafe {
        WindowFromPoint(POINT {
            x: screen.0,
            y: screen.1,
        })
    };
    if target.0.is_null() || (target != page && !unsafe { IsChild(page, target) }.as_bool()) {
        return false;
    }

    let (mut point, wparam) = match msg {
        WM_MOUSEWHEEL | WM_MOUSEHWHEEL => (
            POINT {
                x: screen.0,
                y: screen.1,
            },
            wparam,
        ),
        WM_SETCURSOR => (
            POINT::default(),
            WPARAM((HTCLIENT as usize) | ((WM_MOUSEMOVE as usize) << 16)),
        ),
        _ => {
            let (x, y) = point_of(lparam);
            (POINT { x, y }, wparam)
        }
    };

    if msg != WM_MOUSEWHEEL && msg != WM_MOUSEHWHEEL && msg != WM_SETCURSOR {
        unsafe { MapWindowPoints(hwnd, target, std::slice::from_mut(&mut point)) };
    }

    if a_press(msg) {
        // A wheel or a move would take the caret from the pin's own caption: only a press is what
        // makes the page the window the keys belong to (see `window_state::engine_owns_the_keyboard`).
        unsafe { focus(page) };
        unsafe { SetCapture(hwnd) };
    } else if a_release(msg) {
        let _ = unsafe { ReleaseCapture() };
    }

    let cursor = if msg == WM_SETCURSOR {
        WPARAM(target.0 as usize)
    } else {
        wparam
    };

    let _ = unsafe {
        SendMessageW(
            target,
            msg,
            cursor,
            LPARAM((((point.y as u32 & 0xFFFF) << 16) | (point.x as u32 & 0xFFFF)) as isize),
        )
    };

    true
}

/// Make the page the window the caret is in, which is the pin's thread asking for a window of the
/// engine's: a thread can only focus a window of its own queue, so the two queues are joined for
/// the one call and parted again.
unsafe fn focus(page: HWND) {
    let page_thread = unsafe { GetWindowThreadProcessId(page, None) };
    let mine = unsafe { GetCurrentThreadId() };
    let joined = page_thread != 0 && page_thread != mine;

    if joined {
        let _ = unsafe { AttachThreadInput(mine, page_thread, true) };
    }
    let _ = unsafe { SetFocus(page) };
    if joined {
        let _ = unsafe { AttachThreadInput(mine, page_thread, false) };
    }
}

/// A message's two coordinates, signed, which is how every mouse message carries a point.
fn point_of(lparam: LPARAM) -> (i32, i32) {
    let packed = lparam.0 as u32;
    (
        (packed & 0xFFFF) as u16 as i16 as i32,
        (packed >> 16) as u16 as i16 as i32,
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The band is the page's box, and the pixel at its left and top is on it while the one at the
    /// far edge is the next window's: the same half-open rule every region of this app's follows.
    #[test]
    fn a_point_is_on_the_band_up_to_but_not_at_its_far_edges() {
        let band = (100, 200, 640, 480);

        assert!(band_holds(band, (100, 200)), "the top-left pixel is on it");
        assert!(band_holds(band, (739, 679)), "and the last one inside it");
        assert!(!band_holds(band, (740, 600)), "the far edge is not");
        assert!(!band_holds(band, (700, 680)), "neither is the bottom one");
        assert!(!band_holds(band, (99, 300)));
        assert!(!band_holds(band, (300, 199)));
    }

    /// The claim on the band belongs to the engine's two kinds and to nothing else, and to them
    /// only while the page is nested: a claim on a band with no page in it eats the presses of
    /// whatever stands behind the pin.
    #[test]
    fn only_a_nested_page_claims_its_band() {
        assert!(page_claims_the_band(MediaType::EngineSvg, true));
        assert!(page_claims_the_band(MediaType::EngineFont, true));
        assert!(
            !page_claims_the_band(MediaType::EngineSvg, false),
            "a page on the desktop is a hole the pin must not claim"
        );
        assert!(!page_claims_the_band(MediaType::StaticImage, true));
        assert!(!page_claims_the_band(MediaType::Video, true));
    }

    /// What the pin hands the page is the mouse messages and nothing of its own: a key, a paint or
    /// a capture change is the pin's.
    #[test]
    fn the_pin_hands_the_page_the_pointer_and_not_its_own_messages() {
        for msg in [
            WM_LBUTTONDOWN,
            WM_LBUTTONUP,
            WM_RBUTTONDOWN,
            WM_RBUTTONUP,
            WM_MBUTTONDOWN,
            WM_MBUTTONUP,
            WM_MOUSEMOVE,
            WM_MOUSEWHEEL,
            WM_MOUSEHWHEEL,
            WM_SETCURSOR,
        ] {
            assert!(carried(msg), "{msg:#x} is the page's");
        }

        assert!(!carried(
            windows::Win32::UI::WindowsAndMessaging::WM_KEYDOWN
        ));
        assert!(!carried(windows::Win32::UI::WindowsAndMessaging::WM_PAINT));
        assert!(!carried(
            windows::Win32::UI::WindowsAndMessaging::WM_CAPTURECHANGED
        ));

        assert!(a_press(WM_LBUTTONDOWN));
        assert!(a_release(WM_LBUTTONUP));
        assert!(!a_press(WM_MOUSEMOVE));
        assert!(!a_release(WM_MOUSEMOVE));
    }

    /// A message's point is two signed halves, and a point left of the pin's origin is negative
    /// rather than a large positive one.
    #[test]
    fn a_messages_point_is_signed() {
        let packed = LPARAM(((300isize) << 16) | ((-20isize) & 0xFFFF));
        assert_eq!(point_of(packed), (-20, 300));
    }
}

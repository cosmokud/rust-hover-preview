//! System-wide mouse-wheel watcher.
//!
//! The Explorer hook resolves hovered files from the cursor position, so a wheel
//! scroll that moves the list under a parked pointer is invisible to it: nothing
//! changes until the mouse is moved. A low-level mouse hook reports wheel
//! messages without touching Explorer, and the hook thread publishes a monotonic
//! tick counter that the polling loop consumes to re-resolve the item under the
//! cursor once the list settles.
//!
//! The same hook is what lets a scrollable text preview take the wheel: a preview the
//! user has not pressed is a window the keyboard is not in, so Windows would deliver
//! its wheel messages to Explorer instead. When the pointer is inside the region a
//! scrollable preview published, the message is swallowed here and counted for the
//! preview thread.

use std::sync::atomic::{AtomicI32, AtomicU32, AtomicU64, Ordering};
use windows::Win32::Foundation::{LPARAM, LRESULT, POINT, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::{
    CallNextHookEx, HC_ACTION, MSLLHOOKSTRUCT, WH_MOUSE_LL, WM_MOUSEHWHEEL, WM_MOUSEWHEEL,
};

/// Wheel messages seen since startup; the hook thread is the only writer.
static WHEEL_TICKS: AtomicU64 = AtomicU64::new(0);

/// Wheel notches handed to a text preview. Signed: up is positive. Swapped to
/// zero by the preview thread, so a notch is counted once.
static TEXT_SCROLL_DELTA: AtomicI32 = AtomicI32::new(0);

/// Thread running the hook's message pump, published so shutdown can wake it.
static WHEEL_THREAD_ID: AtomicU32 = AtomicU32::new(0);

/// Number of wheel messages seen so far, counted across the whole desktop.
pub fn wheel_tick_count() -> u64 {
    WHEEL_TICKS.load(Ordering::Relaxed)
}

/// Wheel notches a text preview has been given since this was last called, in
/// wheel units (120 to a notch).
pub fn take_text_scroll_delta() -> i32 {
    TEXT_SCROLL_DELTA.swap(0, Ordering::AcqRel)
}

/// Starts the hook thread; `main` joins the returned handle after `request_stop`.
pub fn spawn_wheel_watcher() -> std::thread::JoinHandle<()> {
    super::hook_thread::spawn(
        &WHEEL_THREAD_ID,
        WH_MOUSE_LL,
        Some(wheel_hook_proc),
        "mouse wheel",
    )
}

/// Wakes the hook thread's message pump so `main` can join it.
pub fn request_stop() {
    super::hook_thread::request_stop(&WHEEL_THREAD_ID)
}

unsafe extern "system" fn wheel_hook_proc(code: i32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    // Keep this callback cheap: every wheel message on the desktop waits for it
    // to return, so it reads one rectangle without blocking and otherwise touches
    // a single atomic.
    if code == HC_ACTION as i32
        && (wparam.0 == WM_MOUSEWHEEL as usize || wparam.0 == WM_MOUSEHWHEEL as usize)
    {
        if let Some(region) = crate::ui::preview_window::text_scroll_keep_alive_try() {
            let info = &*(lparam.0 as *const MSLLHOOKSTRUCT);
            let point = info.pt;
            if point_in_region(point, region) {
                // The wheel is the preview's, not Explorer's: count it for the
                // preview thread and swallow the message so the list behind the
                // preview does not move as well. A low-level hook carries the
                // delta in the high word of `mouseData` — the message's own
                // `wParam` is not filled in here, which is why reading it there
                // would count every notch as zero.
                if wparam.0 == WM_MOUSEWHEEL as usize {
                    let notches = (info.mouseData >> 16) as u16 as i16 as i32;
                    TEXT_SCROLL_DELTA.fetch_add(notches, Ordering::AcqRel);
                }
                return LRESULT(1);
            }
        }

        WHEEL_TICKS.fetch_add(1, Ordering::Relaxed);
    }

    CallNextHookEx(None, code, wparam, lparam)
}

fn point_in_region(point: POINT, region: (i32, i32, i32, i32)) -> bool {
    let (left, top, right, bottom) = region;
    point.x >= left && point.x < right && point.y >= top && point.y < bottom
}

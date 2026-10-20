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
//!
//! It is the same hook again for a pinned picture, whose window is one nobody is in
//! either: a notch over it is the picture being zoomed, and the count is left for the
//! window procedure to take (see `pin_zoom_keep_alive_try`).
//!
//! And for a pinned page, where the notch is its own text size rather than its place:
//! a `Ctrl` held over one steps the size, and a page short enough to fit its box is
//! resized even though it has no scroll to give (see `text_size_keep_alive_try`).
//!
//! And it is the same hook that has to *stop* there, because a region is a box on the
//! screen and not a claim on what is in front of it: a pin can be behind another
//! program — that is what `Pin Mode → Always On Top` off means — and a notch turned
//! while the user is working in that program is that program's. So a notch belongs to
//! this app only while the preview is the surface the pointer is standing on, which is
//! the question the preview thread answers once a tick and publishes here: a pin raised
//! over the listing the user is working in is exactly that surface, with the keyboard
//! still in the listing behind it (see `publish_preview_under_the_pointer`), and a pin
//! behind another program is not (see `key_input::pin_owns_foreground`).

use std::sync::atomic::{AtomicI32, AtomicU32, AtomicU64, Ordering};
use windows::Win32::Foundation::{LPARAM, LRESULT, POINT, WPARAM};
use windows::Win32::UI::Input::KeyboardAndMouse::{GetAsyncKeyState, VK_CONTROL};
use windows::Win32::UI::WindowsAndMessaging::{
    CallNextHookEx, HC_ACTION, MSLLHOOKSTRUCT, WH_MOUSE_LL, WM_MOUSEHWHEEL, WM_MOUSEWHEEL,
};

/// Wheel messages seen since startup; the hook thread is the only writer.
static WHEEL_TICKS: AtomicU64 = AtomicU64::new(0);

/// Wheel notches handed to a text preview. Signed: up is positive. Swapped to
/// zero by the preview thread, so a notch is counted once.
static TEXT_SCROLL_DELTA: AtomicI32 = AtomicI32::new(0);

/// Wheel notches handed to a pinned picture, in the same units and on the same terms as
/// the text preview's own count above: signed, swapped to zero by the reader, and added
/// to only where the pointer is inside the box the pin published.
static PIN_ZOOM_DELTA: AtomicI32 = AtomicI32::new(0);

/// Wheel notches handed to a pinned page's own text size, on the same terms again: a notch turned
/// over a page with `Ctrl` held is how large that page is drawn rather than where it is (see
/// `text_size_after_notches`).
static PIN_TEXT_SIZE_DELTA: AtomicI32 = AtomicI32::new(0);

/// Thread running the hook's message pump, published so shutdown can wake it.
static WHEEL_THREAD_ID: AtomicU32 = AtomicU32::new(0);

/// Whether the preview on screen is the surface the pointer is standing on, published by the
/// preview thread once a tick.
///
/// It is what makes a pin *scrollable the moment it is raised*: a notch over a preview used to be
/// this app's only while the pin was the window the keyboard was in — Always On Top, or the window
/// the user had clicked into — so a pin raised over the listing they were working in took no notch
/// and had to be clicked once before it would scroll. The pointer is over the pin the whole time;
/// the keyboard is in the listing behind it, which is the point of the raise (see
/// `Pin Mode → Update Preview → Move To Top`).
///
/// It is published rather than asked here because a `WH_MOUSE_LL` callback may not walk the
/// desktop to find out what is in front — the same reason the pin's key owners are published by
/// the same tick (see `key_input::publish_pin_key_owner`). One tick of staleness is all it can be
/// wrong by, and a notch is turned by a hand that is not moving.
static PREVIEW_UNDER_POINTER: AtomicU32 = AtomicU32::new(0);

/// Whether the preview on screen is the surface the pointer is standing on, as the preview thread
/// last saw it (see `PREVIEW_UNDER_POINTER`).
fn preview_is_under_the_pointer() -> bool {
    PREVIEW_UNDER_POINTER.load(Ordering::Acquire) != 0
}

/// Publish what the pointer is on, from the tick that is allowed to ask the desktop (see
/// `PREVIEW_UNDER_POINTER`).
pub(crate) fn publish_preview_under_the_pointer(under: bool) {
    PREVIEW_UNDER_POINTER.store(u32::from(under), Ordering::Release);
}

/// Whether a notch inside a region this app published is this app's to take (see
/// `wheel_hook_proc`).
///
/// Two questions and either one is enough. The keyboard's is the older of the two: a pin that is
/// Always On Top, or one the user has clicked into, is a preview standing over the pointer wherever
/// the keyboard is, and a hover — no pin up — is a window nothing can be above. The pointer's is
/// the one that made a raised pin scrollable without a click: the preview is what the pointer is
/// standing on, which the preview thread read off the desktop once a tick (see
/// `preview_is_under_the_pointer`).
fn preview_takes_the_wheel() -> bool {
    crate::shell::key_input::pin_owns_foreground() || preview_is_under_the_pointer()
}

/// Number of wheel messages seen so far, counted across the whole desktop.
pub fn wheel_tick_count() -> u64 {
    WHEEL_TICKS.load(Ordering::Relaxed)
}

/// Wheel notches a text preview has been given since this was last called, in
/// wheel units (120 to a notch).
pub fn take_text_scroll_delta() -> i32 {
    TEXT_SCROLL_DELTA.swap(0, Ordering::AcqRel)
}

/// Wheel notches a pinned page has been given for its own text size since this was last called, in
/// wheel units (120 to a notch). Signed, up being larger.
pub fn take_text_size_delta() -> i32 {
    PIN_TEXT_SIZE_DELTA.swap(0, Ordering::AcqRel)
}

/// Whether `Ctrl` is held, asked of the keyboard rather than read from the message: a wheel message
/// carries no modifiers of its own, and this is a read of one number where a low-level callback may
/// do nothing else (see `wheel_hook_proc`).
fn control_is_down() -> bool {
    unsafe { GetAsyncKeyState(VK_CONTROL.0 as i32) as u16 & 0x8000 != 0 }
}

/// Wheel notches a pinned picture has been given since this was last called, in wheel
/// units (120 to a notch). Signed, up being positive — which is a zoom in.
pub fn take_pin_zoom_delta() -> i32 {
    PIN_ZOOM_DELTA.swap(0, Ordering::AcqRel)
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

/// The hook's own procedure, `pub(crate)` so that a test can hand it a wheel message the way the
/// system does — three numbers and a struct — since what it decides is which of the three owners a
/// notch belongs to, and whether the loop that reads the count is woken for it.
pub(crate) unsafe extern "system" fn wheel_hook_proc(
    code: i32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    // Keep this callback cheap: every wheel message on the desktop waits for it
    // to return, so it reads one rectangle without blocking and otherwise touches
    // a single atomic.
    if code == HC_ACTION as i32
        && (wparam.0 == WM_MOUSEWHEEL as usize || wparam.0 == WM_MOUSEHWHEEL as usize)
    {
        // A low-level hook carries the delta in the high word of `mouseData` — the
        // message's own `wParam` is not filled in here, which is why reading it there
        // would count every notch as zero. A horizontal wheel reports a delta here too
        // and is passed through rather than counted: nothing this app owns moves sideways.
        let info = &*(lparam.0 as *const MSLLHOOKSTRUCT);
        let notches = (info.mouseData >> 16) as u16 as i16 as i32;

        if let Some(region) = crate::ui::preview_window::text_scroll_keep_alive_try() {
            // The box the preview published is asked first, and what owns the wheel second: in that
            // order, so that the second — a few calls into Windows — is a cost neither every wheel
            // message on the desktop nor a notch that is outside every preview pays (see
            // `preview_takes_the_wheel`).
            if point_in_region(info.pt, region) && preview_takes_the_wheel() {
                // The wheel is the preview's, not Explorer's: count it for the
                // preview thread and swallow the message so the list behind the
                // preview does not move as well. `Ctrl` held makes it the page's own size rather
                // than its place, which is the one grip a reader has on how large a pinned page is
                // drawn (see `PIN_TEXT_SIZE_DELTA`).
                if wparam.0 == WM_MOUSEWHEEL as usize {
                    if control_is_down() {
                        PIN_TEXT_SIZE_DELTA.fetch_add(notches, Ordering::AcqRel);
                    } else {
                        TEXT_SCROLL_DELTA.fetch_add(notches, Ordering::AcqRel);
                    }
                }
                return LRESULT(1);
            }
        }

        // And the page's size for a page with no scroll to give: one that fits its box is a page
        // the branch above never reaches, because there is no scroll there for Explorer to lose —
        // so a `Ctrl` and a wheel over one is asked for here instead (see
        // `text_size_keep_alive_try`).
        if wparam.0 == WM_MOUSEWHEEL as usize && control_is_down() {
            if let Some(region) = crate::ui::preview_window::text_size_keep_alive_try() {
                if point_in_region(info.pt, region) && preview_takes_the_wheel() {
                    PIN_TEXT_SIZE_DELTA.fetch_add(notches, Ordering::AcqRel);
                    return LRESULT(1);
                }
            }
        }

        // And the same answer for a pinned picture, which is a window the keyboard is not in for
        // the same reason: a hover the user pinned is a window nobody has pressed, so Explorer
        // would take the notch and scroll the list under it — and the whole of what a pin does not
        // get for free is the other half of that: a pin raised over the listing is still a window
        // the keyboard is not in, and it is the surface the pointer is on all the same (see
        // `preview_takes_the_wheel`). One that is behind another program is that program's.
        //
        // The loop that reads the count is woken here as well, because the count is an
        // atomic and nothing else: a notch turned while that loop is in its own wait for a
        // pin would be read at the end of it — 62 ms, measured, which is four frames of a
        // 144 Hz display and the whole of what makes a zoom start unevenly (see
        // `wake_for_a_wheel_notch`).
        if let Some(region) = crate::ui::preview_window::pin_zoom_keep_alive_try() {
            if point_in_region(info.pt, region) && preview_takes_the_wheel() {
                if wparam.0 == WM_MOUSEWHEEL as usize && notches != 0 {
                    PIN_ZOOM_DELTA.fetch_add(notches, Ordering::AcqRel);
                    crate::ui::preview_window::wake_for_a_wheel_notch();
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

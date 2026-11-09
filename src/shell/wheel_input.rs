//! System-wide mouse-wheel watcher.
//!
//! The Explorer hook resolves hovered files from the cursor position, so a wheel
//! scroll that moves the list under a parked pointer is invisible to it: nothing
//! changes until the mouse is moved. A low-level mouse hook reports wheel
//! messages without touching Explorer, and the hook thread publishes a monotonic
//! tick counter that the polling loop consumes to re-resolve the item under the
//! cursor once the list settles.
//!
//! That counter is the run's rather than this process's, because the loop that consumes it is the
//! tray's and the hook may be a viewer's (see `app::seam::WHEEL_TICKS`).
//!
//! The same hook is what lets a scrollable text preview take the wheel: a preview the
//! user has not pressed is a window the keyboard is not in, so Windows would deliver
//! its wheel messages to Explorer instead. When the pointer is inside the region a
//! scrollable preview published, the message is swallowed here and counted for the
//! preview thread.
//!
//! It is the same hook again for a pinned picture, whose window is one nobody is in
//! either: a notch over it is the picture being zoomed, and the count is left for the
//! window procedure to take (see `PIN_ZOOM_REGION`).
//!
//! And for a pinned page, where the notch is its own text size rather than its place:
//! a `Ctrl` held over one steps the size, and a page short enough to fit its box is
//! resized even though it has no scroll to give (see `TEXT_SIZE_REGION`).
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

use crate::app::seam::{
    ask_the_pin_to_go_behind, MODELS_BAND, PIN_PLAYLIST_BAND, PIN_PLAYLIST_DELTA,
    PIN_TEXT_SIZE_DELTA, PIN_VOLUME_DELTA, PIN_VOLUME_REGION, PIN_ZOOM_DELTA, PIN_ZOOM_REGION,
    TEXT_SCROLL_DELTA, TEXT_SCROLL_REGION, TEXT_SIZE_REGION,
};
use crate::app::seam::{post_a_notch_to_the_model, wake_preview_loop};
use crate::shell::explorer_hook::{
    a_press_puts_us_behind, a_window_of_ours_stands, our_window_stands_over_everything,
    press_is_inside_ours, the_pick_road_answers, window_class_of,
};
use once_cell::sync::Lazy;
use std::sync::atomic::{AtomicU32, Ordering};
use windows::Win32::Foundation::{LPARAM, LRESULT, POINT, WPARAM};
use windows::Win32::UI::Input::KeyboardAndMouse::{GetAsyncKeyState, VK_CONTROL};
use windows::Win32::UI::WindowsAndMessaging::{
    CallNextHookEx, WindowFromPoint, HC_ACTION, MSLLHOOKSTRUCT, WH_MOUSE_LL, WM_LBUTTONDOWN,
    WM_MBUTTONDOWN, WM_MBUTTONUP, WM_MOUSEHWHEEL, WM_MOUSEWHEEL, WM_RBUTTONDOWN, WM_XBUTTONDOWN,
    WM_XBUTTONUP,
};

/// Thread running the hook's message pump, published so shutdown can wake it.
static WHEEL_THREAD_ID: AtomicU32 = AtomicU32::new(0);

/// Whether a notch inside a region this app published is this app's to take (see
/// `wheel_hook_proc`).
///
/// Two questions and either one is enough. The keyboard's is the older of the two: a pin that is
/// Always On Top, or one the user has clicked into, is a preview standing over the pointer wherever
/// the keyboard is, and a hover — no pin up — is a window nothing can be above. The pointer's is
/// the one that made a raised pin scrollable without a click: the preview is what the pointer is
/// standing on, which the preview thread read off the desktop once a tick (see
/// `the_pointer_is_on`).
///
/// **It is asked of the one block whose box the notch was turned in**, and of no other: a detached
/// window's box is its own, and the viewer that took its place in the ring is not the window the
/// pointer is standing on.
fn preview_takes_the_wheel(block: &'static crate::app::seam::ViewerFacts) -> bool {
    crate::shell::key_input::pin_owns_foreground(block)
        || crate::app::seam::the_pointer_is_on(block)
}

/// Number of wheel messages seen so far, counted across the whole desktop by whichever process of the
/// run holds this hook (see `app::seam::WHEEL_TICKS`).
pub fn wheel_tick_count() -> u64 {
    crate::app::seam::wheel_ticks_seen()
}

/// Whether `Ctrl` is held, asked of the keyboard rather than read from the message: a wheel message
/// carries no modifiers of its own, and this is a read of one number where a low-level callback may
/// do nothing else (see `wheel_hook_proc`).
fn control_is_down() -> bool {
    unsafe { GetAsyncKeyState(VK_CONTROL.0 as i32) as u16 & 0x8000 != 0 }
}

/// Starts the hook thread; `main` joins the returned handle after `request_stop`.
pub fn spawn_wheel_watcher() -> std::thread::JoinHandle<()> {
    // The press road's trace is read here rather than on a press: a low-level callback may not read the
    // environment, which allocates, while every mouse message on the desktop waits for it (see
    // `note_a_press`).
    Lazy::force(&PRESS_TRACE);

    super::hook_thread::spawn(
        &WHEEL_THREAD_ID,
        WH_MOUSE_LL,
        Some(wheel_hook_proc),
        "mouse wheel",
    )
}

/// Whether the press road says what it decided, read once (see `RHP_SEAM_TRACE`).
static PRESS_TRACE: Lazy<bool> = Lazy::new(|| std::env::var_os("RHP_SEAM_TRACE").is_some());

/// Put the standing window of ours behind the window a press landed in, or ask for nothing.
///
/// **It is the press and not a foreground transition**, which is the difference that took three attempts
/// to find: a press inside the window that is already active — every press a listing is made of — moves
/// no foreground, so the road that watched the foreground never saw one, and the pin stayed over the
/// listing it was pressed in (see `lower_pinned_preview`, which carries the history).
///
/// Nothing here places a window. A placement is a call into the window manager and this callback runs
/// inside the OS input path, so the ask is published and the preview loop makes it on its own tick: the
/// same road a pick's press takes (see `ask_the_pin_to_go_behind`).
///
/// **A press the pick road is going to answer is no press of this one.** The press that lands in the
/// listing behind a pin that follows that listing is the press of a pick: the file under it is read by
/// the hook's own tick and the pin is asked to show it, and asking the pin to go behind the listing
/// first is one press answered twice — the listing over the pin until the pick's raise lands, which is
/// the flash a mouse pick has and a keyboard pick does not (see `the_pick_road_answers` and
/// `a_press_puts_us_behind`). A press that names no file is still answered, and answered behind: the
/// hook's `pin_press_was_not_a_pick` is that press's own.
///
/// **Every process that draws a window of ours holds one of these hooks**, which is what makes the rule
/// the same for a detached window as for a pin: each asks about the press, and each answers with the
/// window it holds.
fn note_a_press(point: POINT) {
    // A press on a desktop with nothing of ours standing is no press of ours: one atomic, which is what
    // every press costs on a run with no pin up.
    if !a_window_of_ours_stands() {
        return;
    }

    // Safety: asks the desktop for the window under a point it was given; dereferences nothing.
    let window = unsafe { WindowFromPoint(point) };
    let ours = press_is_inside_ours(window, point);
    let always_on_top = our_window_stands_over_everything();
    let the_listing_answers = !ours && the_pick_road_answers(window);

    if !a_press_puts_us_behind(true, ours, true, always_on_top, the_listing_answers) {
        if *PRESS_TRACE {
            eprintln!(
                "seam: press at {},{} on {} ({}), ours {} -> nothing asked (always on top {}, the \
                 listing answers {})",
                point.x,
                point.y,
                window.0 as isize,
                window_class_of(window),
                ours as u8,
                always_on_top as u8,
                the_listing_answers as u8,
            );
        }
        return;
    }

    // The frame and not the window: a press inside a listing is a press on one of its children, and a
    // window put behind a child is a placement the desktop does nothing with (see `the_frame_of`).
    let behind = crate::shell::window_classes::the_frame_of(window);
    ask_the_pin_to_go_behind(behind.0 as isize);
    // The ask is an atomic and nothing else, so the loop that reads it is woken for it: the press is a
    // window being put over ours, and a lowering one tick late is the listing flashing over the pin.
    wake_preview_loop();

    if *PRESS_TRACE {
        eprintln!(
            "seam: press at {},{} on {} ({}) is not ours -> asked behind {} ({})",
            point.x,
            point.y,
            window.0 as isize,
            window_class_of(window),
            behind.0 as isize,
            window_class_of(behind),
        );
    }
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
    // Keep this callback cheap: every mouse message on the desktop waits for it
    // to return, so it reads one rectangle without blocking and otherwise touches
    // a single atomic.
    // The pin key may be a mouse button, and a button never reaches the keyboard hook: its edge is
    // counted here instead, so a click shorter than the preview loop's wait is still a press. The
    // message is passed on like every other — the button is another window's as much as this app's.
    if code == HC_ACTION as i32 {
        match wparam.0 as u32 {
            // A press, and not the pin key's button: the button a pin key names is a press that raises or
            // lowers the pin, and a demotion on it would be the key answering twice (see `note_a_press`).
            WM_LBUTTONDOWN | WM_RBUTTONDOWN => {
                // Safety: the desktop passes the event's own `MSLLHOOKSTRUCT` in `lParam` of a
                // `HC_ACTION` call, and it outlives this callback.
                let info = unsafe { &*(lparam.0 as *const MSLLHOOKSTRUCT) };
                note_a_press(info.pt);
            }
            WM_MBUTTONDOWN | WM_MBUTTONUP => {
                crate::shell::key_input::note_mouse_button(
                    0x04, // VK_MBUTTON
                    wparam.0 == WM_MBUTTONDOWN as usize,
                );
            }
            WM_XBUTTONDOWN | WM_XBUTTONUP => {
                // The button id is the high word of mouseData, the same place a notch's delta
                // sits for the wheel: 1 is XBUTTON1 (mouse4), 2 is XBUTTON2 (mouse5).
                let info = &*(lparam.0 as *const MSLLHOOKSTRUCT);
                let vk = match (info.mouseData >> 16) as u16 {
                    1 => 0x05, // VK_XBUTTON1
                    2 => 0x06, // VK_XBUTTON2
                    _ => 0,
                };
                if vk != 0 {
                    crate::shell::key_input::note_mouse_button(
                        vk,
                        wparam.0 == WM_XBUTTONDOWN as usize,
                    );
                }
            }
            _ => {}
        }
    }

    if code == HC_ACTION as i32
        && (wparam.0 == WM_MOUSEWHEEL as usize || wparam.0 == WM_MOUSEHWHEEL as usize)
    {
        // A low-level hook carries the delta in the high word of `mouseData` — the
        // message's own `wParam` is not filled in here, which is why reading it there
        // would count every notch as zero. A horizontal wheel reports a delta here too
        // and is passed through rather than counted: nothing this app owns moves sideways.
        let info = &*(lparam.0 as *const MSLLHOOKSTRUCT);
        let notches = (info.mouseData >> 16) as u16 as i16 as i32;

        // A notch over the volume popup or the button that opens it is the level's, and it is
        // asked before the two regions below: the button sits inside the pin's own box, which
        // the zoom road claims whole, and the popup can float over a page that scrolls.
        if wparam.0 == WM_MOUSEWHEEL as usize && notches != 0 {
            if let Some(block) = PIN_VOLUME_REGION.the_block_over((info.pt.x, info.pt.y)) {
                if preview_takes_the_wheel(block) {
                    PIN_VOLUME_DELTA
                        .in_block(block)
                        .fetch_add(notches, Ordering::AcqRel);
                    wake_preview_loop();
                    return LRESULT(1);
                }
            }
        }

        if let Some(block) = TEXT_SCROLL_REGION.the_block_over((info.pt.x, info.pt.y)) {
            // The box the preview published is asked first, and what owns the wheel second: in that
            // order, so that the second — a few calls into Windows — is a cost neither every wheel
            // message on the desktop nor a notch that is outside every preview pays (see
            // `preview_takes_the_wheel`).
            if preview_takes_the_wheel(block) {
                // The wheel is the preview's, not Explorer's: count it for the
                // preview thread and swallow the message so the list behind the
                // preview does not move as well. `Ctrl` held makes it the page's own size rather
                // than its place, which is the one grip a reader has on how large a pinned page is
                // drawn (see `PIN_TEXT_SIZE_DELTA`).
                //
                // The loop that reads the count is woken for it, because the count is an atomic and
                // nothing else: a notch turned while that loop is in its own wait for a pin would be
                // read at the end of it — 62 ms, measured, which is four frames of a 144 Hz display
                // and the whole of what makes a scroll start unevenly (see `wake_preview_loop`). Both
                // counts are owed the same wake: a `Ctrl` notch is the page resized rather than moved,
                // which waits on the same tick.
                if wparam.0 == WM_MOUSEWHEEL as usize {
                    if control_is_down() {
                        PIN_TEXT_SIZE_DELTA
                            .in_block(block)
                            .fetch_add(notches, Ordering::AcqRel);
                    } else {
                        TEXT_SCROLL_DELTA
                            .in_block(block)
                            .fetch_add(notches, Ordering::AcqRel);
                    }
                    if notches != 0 {
                        wake_preview_loop();
                    }
                }
                return LRESULT(1);
            }
        }

        // And the page's size for a page with no scroll to give: one that fits its box is a page
        // the branch above never reaches, because there is no scroll there for Explorer to lose —
        // so a `Ctrl` and a wheel over one is asked for here instead (see
        // `TEXT_SIZE_REGION`).
        if wparam.0 == WM_MOUSEWHEEL as usize && control_is_down() {
            if let Some(block) = TEXT_SIZE_REGION.the_block_over((info.pt.x, info.pt.y)) {
                if preview_takes_the_wheel(block) {
                    PIN_TEXT_SIZE_DELTA
                        .in_block(block)
                        .fetch_add(notches, Ordering::AcqRel);
                    return LRESULT(1);
                }
            }
        }

        // **A pinned model's notch is its own scene's**, and it is asked before the pictures' zoom below:
        // f3d's wheel is its camera's and not a picture of this app's, so the notch is posted to the
        // engine's window and swallowed here — the kinds a notch zooms are refused there rather than here
        // (see `pin_zoom_view`), and no notch falls through to the listing under the pin.
        if wparam.0 == WM_MOUSEWHEEL as usize {
            if let Some(block) = MODELS_BAND.the_block_over((info.pt.x, info.pt.y)) {
                if preview_takes_the_wheel(block) {
                    if notches != 0 {
                        post_a_notch_to_the_model(block, info.pt.x, info.pt.y, notches);
                    }
                    return LRESULT(1);
                }
            }
        }

        // The playlist band under a pinned sound's card: a notch over it is the list being stepped
        // rather than the card being zoomed, and it is asked before the pictures' own notch because
        // a band stands inside the box that road claims whole (see `PIN_PLAYLIST_BAND`).
        if let Some(block) = PIN_PLAYLIST_BAND.the_block_over((info.pt.x, info.pt.y)) {
            if preview_takes_the_wheel(block) {
                if wparam.0 == WM_MOUSEWHEEL as usize && notches != 0 {
                    PIN_PLAYLIST_DELTA
                        .in_block(block)
                        .fetch_add(notches, Ordering::AcqRel);
                    wake_preview_loop();
                }
                return LRESULT(1);
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
        // `wake_preview_loop`).
        if let Some(block) = PIN_ZOOM_REGION.the_block_over((info.pt.x, info.pt.y)) {
            if preview_takes_the_wheel(block) {
                if wparam.0 == WM_MOUSEWHEEL as usize && notches != 0 {
                    PIN_ZOOM_DELTA
                        .in_block(block)
                        .fetch_add(notches, Ordering::AcqRel);
                    wake_preview_loop();
                }
                return LRESULT(1);
            }
        }

        crate::app::seam::count_a_wheel_notch();
    }

    CallNextHookEx(None, code, wparam, lparam)
}

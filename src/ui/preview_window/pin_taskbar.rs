//! The pin's taskbar stand-in: the window a pin leaves behind when it is put away into the
//! taskbar rather than into the round bubble.
//!
//! Every other window this app owns is a `WS_EX_TOOLWINDOW` popup, which is what keeps previews
//! out of the taskbar and out of Alt+Tab — nothing of this app is expected to be a window the
//! desktop lists. A user who asks for **Pin Mode → Minimize → To Taskbar** is asking for the
//! opposite for one state: a pin put away should leave the app in the taskbar the way any other
//! window that is minimized does, so it can be brought back from there.
//!
//! Rather than mutate the shared preview window's extended style — which would risk the hover
//! and pin surfaces leaking into the taskbar, the very thing previews are kept out of — this is
//! a window of its own, created once and reused, exactly as the bubble is (see `pin_bubble`).
//! It is a plain overlapped window with an app icon and title, kept 1×1 and off the screen, and
//! it is *shown minimized* rather than merely shown: the taskbar's own restore then reaches this
//! window, which answers it by putting the pin back where it stood and taking itself away again.

use super::*;

use windows::Win32::UI::WindowsAndMessaging::{
    SC_CLOSE, SC_RESTORE, SetWindowTextW, SW_SHOWMINNOACTIVE, WA_ACTIVE, WM_COMMAND, WM_QUERYOPEN,
    WS_CAPTION, WS_EX_APPWINDOW, WS_MINIMIZEBOX, WS_OVERLAPPED, WS_SYSMENU,
};
/// The class the taskbar stand-in is created from.
pub(super) const PIN_TASKBAR_CLASS: PCWSTR = w!("RustHoverPreviewPinTaskbar");

/// The stand-in's window, or zero while it has never been shown.
pub(super) static PIN_TASKBAR_HWND: AtomicIsize = AtomicIsize::new(0);

/// Where the stand-in is kept: far enough off every display that even a momentary un-minimize
/// cannot put a window on screen. It is never meant to be seen — only its taskbar button is.
const TASKBAR_OFFSCREEN: i32 = -32000;

/// The title the taskbar button carries. The app's own name rather than the pinned file's, since
/// the button stands for the app that has a window put away, not for a document.
const TASKBAR_TITLE: PCWSTR = w!("Rust Preview");

/// Put the stand-in up: the app appears in the taskbar, minimized and off the screen, and the pin
/// stays hidden until the button is clicked.
///
/// It is *shown minimized* rather than shown and then minimized, because a window that is up has
/// a place on the desktop and this one has none: `SW_SHOWMINNOACTIVE` asks the Shell for the
/// taskbar button without taking the focus, which is the whole of what a pin put away wants.
pub(super) unsafe fn show_pin_taskbar() {
    let Some(hwnd) = pin_taskbar_window() else {
        return;
    };

    // The title is set every time the stand-in is shown, for the reason a preview is painted
    // before it is revealed: a window that is already up is not asked again for the name it was
    // shown with, and a handle kept across a Shell restart is not the window it was.
    let _ = SetWindowTextW(hwnd, TASKBAR_TITLE);
    let _ = MoveWindow(hwnd, TASKBAR_OFFSCREEN, TASKBAR_OFFSCREEN, 1, 1, false);
    let _ = ShowWindow(hwnd, SW_SHOWMINNOACTIVE);
}

/// Take the stand-in down, if one is up.
pub(super) unsafe fn hide_pin_taskbar() {
    let hwnd = PIN_TASKBAR_HWND.load(Ordering::SeqCst);
    if hwnd == 0 {
        return;
    }

    let _ = ShowWindow(HWND(hwnd as *mut _), SW_HIDE);
}

/// Whether the stand-in is showing, which is the same question as whether the pin was put away
/// into the taskbar rather than into the bubble.
///
/// It is read before a restore does anything, because it is what tells the two apart: a restore
/// from the bubble places the window beside the bubble (see `placed_pin_box`), and a restore from
/// the taskbar simply puts it back where it stood.
pub(super) fn pin_taskbar_is_showing() -> bool {
    showing_taskbar_window().is_some()
}

/// The stand-in's window where it is up, or nothing where it is not — asked by a caller that has
/// something to give the window rather than a question about the pin, such as the media buttons a
/// taskbar hover draws (see `pin_taskbar_buttons`).
///
/// A window that has never been shown is not created here: what this answers is whether there is a
/// taskbar button, and asking must not be the thing that puts one there.
pub(super) fn showing_taskbar_window() -> Option<HWND> {
    let hwnd = PIN_TASKBAR_HWND.load(Ordering::SeqCst);
    if hwnd == 0 {
        return None;
    }

    let hwnd = HWND(hwnd as *mut _);
    unsafe { IsWindowVisible(hwnd).as_bool().then_some(hwnd) }
}

/// Take whichever stand-in a collapsed pin left — the round bubble, the taskbar button, or both
/// (only ever one is up) — down together, so a restore and a take-down need not know which one it
/// was.
pub(super) fn hide_minimized_pin() {
    hide_pin_bubble();
    if pin_always_on_top() {
        unsafe { hide_pin_taskbar() };
    }
}

/// The stand-in's window, created once for the run.
///
/// It is a plain overlapped window rather than a layered popup, and it is deliberately *not*
/// `WS_EX_TOOLWINDOW`: that bit is what keeps every other window of this app out of the taskbar,
/// and this is the one window whose whole purpose is to be in it. `WS_EX_APPWINDOW` states that
/// outright rather than leaving it to the ownerless-window rule.
pub(super) unsafe fn pin_taskbar_window() -> Option<HWND> {
    let existing = PIN_TASKBAR_HWND.load(Ordering::SeqCst);
    if existing != 0 {
        return Some(HWND(existing as *mut _));
    }

    let hinstance = GetModuleHandleW(None).ok()?;
    let hwnd = CreateWindowExW(
        WS_EX_APPWINDOW,
        PIN_TASKBAR_CLASS,
        TASKBAR_TITLE,
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        TASKBAR_OFFSCREEN,
        TASKBAR_OFFSCREEN,
        1,
        1,
        None,
        None,
        hinstance,
        None,
    )
    .ok()?;

    PIN_TASKBAR_HWND.store(hwnd.0 as isize, Ordering::SeqCst);
    Some(hwnd)
}

unsafe fn handle_taskbar_restore() {
    if pin_always_on_top() {
        hide_pin_taskbar();
    }
    if pin_is_collapsed() {
        ask_pin(PinCommand::Restore);
    } else {
        let preview_hwnd = HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _);
        if !preview_hwnd.is_invalid() {
            let _ = SetForegroundWindow(preview_hwnd);
            let _ = SetWindowPos(
                preview_hwnd,
                HWND_TOP,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE,
            );
            place_pinned_siblings();
        }
    }
}

/// The stand-in's own window procedure: the two things its taskbar button can be asked for, and
/// nothing else, because it is a window with no surface and no content.
///
/// A button click restores the window the Shell holds — this one — and that is the moment the pin
/// is asked for instead, so the click is answered here rather than passed on: the window is
/// hidden and the pin comes back where it stood. A close, from the button's own menu or from
/// `Alt+F4`, ends the pin. Both are the same door the bubble's click and right-click use (see
/// `pin_bubble_proc`), so there is still one path out of a pin whichever stand-in asked for it.
pub(super) unsafe extern "system" fn pin_taskbar_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        // A press in the hover's toolbar, which the Shell sends here because it is this window's own
        // taskbar button the hover belongs to. What each button means is the pin's own command,
        // queued for the loop the way a press on the caption is (see `pin_taskbar_buttons`): this
        // procedure knows which number is which button and nothing about what a walk or a hold does.
        WM_COMMAND => {
            if let Some(command) = pin_media_button_click(wparam) {
                ask_pin(command);
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_SYSCOMMAND => match (wparam.0 & 0xFFF0) as u32 {
            SC_RESTORE => {
                handle_taskbar_restore();
                LRESULT(0)
            }
            SC_CLOSE => {
                ask_pin(PinCommand::Close);
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        },
        // The Shell asks an iconized window whether it may be opened before it opens it; this one
        // never may, because what it stands in for is not on the screen: the ask is answered by
        // putting the pin back and taking the stand-in away, and the button is never brought up.
        WM_QUERYOPEN => {
            handle_taskbar_restore();
            LRESULT(0)
        }
        WM_ACTIVATE => {
            // A click on the taskbar button may activate the window rather than ask it to
            // restore; the pin is asked for either way. A deactivation says nothing about the pin.
            if (wparam.0 & 0xFFFF) as u32 == WA_ACTIVE {
                handle_taskbar_restore();
            }
            LRESULT(0)
        }
        WM_CLOSE => {
            ask_pin(PinCommand::Close);
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

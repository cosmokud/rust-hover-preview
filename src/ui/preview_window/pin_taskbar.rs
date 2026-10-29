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
//! It is a borderless top-level window with an app icon and a title, kept 1×1 and off the screen,
//! and it is *shown minimized* rather than merely shown: the taskbar's own restore then reaches this
//! window, which answers it by putting the pin back where it stood and taking itself away again. The
//! borderless part is not a detail: the Shell draws the button's hover from the window itself, so
//! whatever the window wears is drawn over the pin's own frame in it (see `stand_in_style`).

use super::*;

use windows::Win32::Foundation::BOOL;
use windows::Win32::Graphics::Dwm::{
    DwmSetWindowAttribute, DWMWA_FORCE_ICONIC_REPRESENTATION, DWMWA_HAS_ICONIC_BITMAP,
};
use windows::Win32::UI::WindowsAndMessaging::{
    DestroyWindow, IsIconic, IsWindow, SetWindowTextW, SC_CLOSE, SC_MINIMIZE, SC_RESTORE,
    SHOW_WINDOW_CMD, SW_MINIMIZE, SW_SHOWMINNOACTIVE, SW_SHOWNOACTIVATE, WA_ACTIVE,
    WINDOW_EX_STYLE, WINDOW_STYLE, WM_COMMAND, WM_DWMSENDICONICTHUMBNAIL, WM_QUERYOPEN,
    WS_EX_APPWINDOW, WS_MINIMIZEBOX, WS_POPUP, WS_SYSMENU,
};
/// The class the taskbar stand-in is created from.
pub(super) const PIN_TASKBAR_CLASS: PCWSTR = w!("RustHoverPreviewPinTaskbar");

/// The stand-in's window, or zero while it has never been shown.
pub(super) static PIN_TASKBAR_HWND: AtomicIsize = AtomicIsize::new(0);

/// Whether this app is asking its own stand-in for a form right now.
///
/// Windows asks an iconic window whether it may be opened — `WM_QUERYOPEN` — for the app's own
/// `ShowWindow` calls as much as for the Shell's gesture on the taskbar button, and the two are
/// told apart by nothing else: the same question arrives either way, from inside the call. What
/// this flag is for is answering the first as "yes, come up" and the second as the pin's own
/// restore (see `pin_taskbar_proc`).
static ASKING_THE_STAND_IN_UP: AtomicBool = AtomicBool::new(false);

/// Ask the stand-in for a form, saying that anything arriving during the ask is this app's own.
///
/// A stand-in that is iconic is *not* shown by `ShowWindow` alone: the call asks the window whether
/// it may open first, and an answer of "no" leaves it iconic — which is how the mirroring step got
/// no picture to take (see `show_pin_taskbar_mirroring`).
unsafe fn ask_the_stand_in(hwnd: HWND, form: SHOW_WINDOW_CMD) {
    ASKING_THE_STAND_IN_UP.store(true, Ordering::SeqCst);
    let _ = ShowWindow(hwnd, form);
    ASKING_THE_STAND_IN_UP.store(false, Ordering::SeqCst);
}

/// Where the stand-in is kept: far enough off every display that even a momentary un-minimize
/// cannot put a window on screen. It is never meant to be seen — only its taskbar button is, and the
/// picture the Shell mirrors into it (see `pin_taskbar_mirror`).
pub(super) const TASKBAR_OFFSCREEN: i32 = -32000;

/// The title the taskbar button carries while there is no file to name: the app's own name, which
/// is what the button stands for until a pin has been read (see `pin_taskbar_title`).
const TASKBAR_TITLE: PCWSTR = w!("Rust Preview");

/// The name the stand-in's window carries: the pinned file's own name, so that the Shell's own
/// label beside the hover says which file the button stands for rather than which app it belongs
/// to. The app's own name while there is no file — a pin whose media is still being read, and a
/// stand-in asked for by a road that has no pin of its own.
fn pin_taskbar_title() -> Vec<u16> {
    taskbar_title_for(pinned_path().as_deref())
}

/// The name a file gives the button, or the app's own name where there is no file to name: the
/// file's own name rather than the path, which is what the Shell's label has room for, and what a
/// hand reads it for.
fn taskbar_title_for(path: Option<&Path>) -> Vec<u16> {
    let name = path
        .and_then(|path| path.file_name())
        .map(|name| name.to_string_lossy().into_owned())
        .filter(|name| !name.is_empty())
        .unwrap_or_else(|| "Rust Preview".to_string());

    name.encode_utf16().chain(std::iter::once(0)).collect()
}

/// Write the name the button is to say, and remember none of it: the title is asked for from the
/// pin's own state every time it is written, so a file stepped onto while the stand-in is up is a
/// file the button names after the step (see `refresh_pin_taskbar_title`).
pub(super) unsafe fn set_pin_taskbar_title(hwnd: HWND) {
    let title = pin_taskbar_title();
    let _ = SetWindowTextW(hwnd, PCWSTR(title.as_ptr()));
}

/// Put the pinned file's own name on the stand-in's window again, for a pin whose file changed
/// while the button was already up — a step of the walk, which is the one road that changes what a
/// pin is showing without taking the stand-in down.
pub(super) fn refresh_pin_taskbar_title() {
    let Some(hwnd) = showing_taskbar_window() else {
        return;
    };

    // Safety: the handle is this app's own stand-in window, and a title is a string of this app's
    // own written onto it.
    unsafe { set_pin_taskbar_title(hwnd) };
}

/// Put the stand-in up: the app appears in the taskbar, off the screen, in whichever of its two forms
/// the pin asks for.
///
/// A pin that is *on the screen* is shown in it: the window is up and the Shell is composing it, which
/// is the whole of what makes the hover a picture of the pin rather than a still frame of it. A pin
/// that has been put away into the taskbar is the other form — the window is minimized, and the frame
/// the Shell cached on the way down is the last frame it was showing (see `pin_taskbar_mirror`).
///
/// **And a button is shown only where one is owed.** Every road here is a road into a button for a pin
/// that is there, but they are not all the same road: the setting's own toggle is one that cannot tell
/// a pin put away into the taskbar from one put away into the round bubble, and the button it would
/// show over the bubble is a button for a pin that is not in the taskbar and that holds no picture of
/// one — every window of the pin's is off the screen. So the rule the tick reads a button back
/// against is asked here too, and the stand-in goes down where none is owed rather than up (see
/// `a_button_is_owed` and `keep_the_stand_in_up`).
///
pub(super) unsafe fn show_pin_taskbar() {
    if !a_button_is_owed() {
        hide_pin_taskbar();
        return;
    }

    let Some(hwnd) = pin_taskbar_window() else {
        return;
    };

    // The name is set every time the stand-in is shown, for the reason a preview is painted
    // before it is revealed: a window that is already up is not asked again for the name it was
    // shown with, and a handle kept across a Shell restart is not the window it was. What it names
    // is the file the pin is of, so the Shell's own label beside the hover says which file the
    // button stands for (see `pin_taskbar_title`).
    set_pin_taskbar_title(hwnd);

    if pin_is_collapsed() {
        // The pin is put away, so the button is all there is of it: what a hover shows of a window
        // that has gone down is the frame the Shell cached on the way down, and the frame it caches
        // is the one the window was last showing — so it is minimized here rather than hidden, and
        // what it was last showing it keeps (see `pin_taskbar_mirror`).
        //
        // And the still of it is the app's own picture rather than whatever the Shell would compose
        // itself, which for a window that is off every display is nothing (see
        // `the_buttons_still_is_the_apps_own`).
        the_buttons_still_is_the_apps_own(hwnd, true);
        ask_the_stand_in(hwnd, minimized_form(hwnd));
        return;
    }

    show_pin_taskbar_mirroring();
}

/// Tell the Shell whether the stills of this window's taskbar button are the bitmap this app supplies,
/// or the window's own composition.
///
/// **It is the app's own only while the pin is put away.** A window the Shell is showing a button for
/// is drawn from its own composition unless it is asked to be drawn from a bitmap instead, and for a
/// put-away pin that composition is nothing at all — the stand-in is off every display, so its rows
/// are whatever the Shell's surface for it last had, and a film host's rows are the black it is erased
/// to behind a picture the Shell composes over it — while the bitmap is the pin (see
/// `pin_taskbar_mirror`). A pin that is *on the screen* is the other way about: its hover is the live
/// picture of the pin, and forcing the bitmap on it would freeze the film the user is watching in the
/// button.
///
/// Asked of both windows a pin can be put away as: the stand-in (see `show_pin_taskbar`), and the film
/// host that owns a pin's chrome on a film a player draws (see `pin_film_host`).
pub(super) unsafe fn the_buttons_still_is_the_apps_own(hwnd: HWND, the_apps_own: bool) {
    #[cfg(test)]
    if the_apps_own {
        STILLS_FORCED.fetch_add(1, Ordering::SeqCst);
    }
    let forced = BOOL(the_apps_own as i32);
    let _ = DwmSetWindowAttribute(
        hwnd,
        DWMWA_FORCE_ICONIC_REPRESENTATION,
        &forced as *const BOOL as *const core::ffi::c_void,
        std::mem::size_of::<BOOL>() as u32,
    );
}

/// How many times this run has told the Shell that a button's still is the app's own.
///
/// **The attribute is write-only** — `DwmGetWindowAttribute` refuses `DWMWA_FORCE_ICONIC_REPRESENTATION`
/// — so a test can see that the ask was made and not what it left behind (see
/// `the_buttons_still_is_the_apps_own`).
#[cfg(test)]
static STILLS_FORCED: AtomicU32 = AtomicU32::new(0);

/// The asks this run has made for a button's still to be the app's own (see `STILLS_FORCED`).
#[cfg(test)]
pub(super) fn stills_forced() -> u32 {
    STILLS_FORCED.load(Ordering::SeqCst)
}

/// Put the stand-in up in the form that shows the pin: off the screen, and *shown* — a window the
/// Shell composes, because a window it is not composing has no picture in it to put in a hover.
///
/// The size comes with the picture rather than being set here: the box a mirror is drawn into is the
/// source's own shape, and only the mirror knows what it is mirroring (see `pin_taskbar_mirror`).
pub(super) unsafe fn show_pin_taskbar_mirroring() {
    let Some(hwnd) = pin_taskbar_window() else {
        return;
    };

    set_pin_taskbar_title(hwnd);
    // A pin that is on the screen is mirrored rather than supplied: what its button's hover shows is
    // the Shell's drawing of the pin as it is, moving with the film (see
    // `the_buttons_still_is_the_apps_own`).
    the_buttons_still_is_the_apps_own(hwnd, false);
    ask_the_stand_in(hwnd, SW_SHOWNOACTIVATE);
}

/// How a stand-in is taken down to its put-away form: minimized, whichever state of up it is in.
///
/// A window that is already up — one a pin on the screen was being shown in — is minimized rather
/// than shown minimized, since a stand-in that is up is one the Shell is composing and the frame it
/// caches is the one it was composing a moment ago. `SW_SHOWMINNOACTIVE` is the other half of that:
/// it asks the Shell for a taskbar button *and* for the minimized state in one call, which is what a
/// stand-in that has never been up needs.
fn minimized_form(hwnd: HWND) -> SHOW_WINDOW_CMD {
    // Safety: asks a window whether it is up, which is a property rather than a change.
    if unsafe { IsWindowVisible(hwnd).as_bool() } {
        SW_MINIMIZE
    } else {
        SW_SHOWMINNOACTIVE
    }
}

/// Take the stand-in down, if one is up.
pub(super) unsafe fn hide_pin_taskbar() {
    let hwnd = PIN_TASKBAR_HWND.load(Ordering::SeqCst);
    if hwnd == 0 {
        return;
    }

    // Nothing of the pin is left for the button to be a picture of, so the Shell is told to stop
    // drawing one: a registration behind a window nobody is showing is a registration that outlives
    // the pin it was for (see `pin_taskbar_mirror`).
    pin_taskbar_mirror::stop_pin_mirror();
    let _ = ShowWindow(HWND(hwnd as *mut _), SW_HIDE);
}

/// Bring the stand-in down for good, for the Shell's own Close on its button.
///
/// The Shell's Close is answered by ending the pin (see `pin_taskbar_proc`), and what this window
/// owes beside that is to be *destroyed* rather than hidden. A hidden window keeps its button and
/// its toolbar to itself, and a toolbar can neither be removed nor added a second time on the
/// window it hangs on (see `pin_taskbar_buttons`): a closed button would come back wearing the
/// marks of the pin that was closed and refuse the next pin's. So the window goes, with the marks
/// it was drawn from, and the next pin makes another one — which is also what takes the button off
/// the taskbar rather than leaving one for a window with nothing in it.
pub(super) unsafe fn close_pin_taskbar() {
    let hwnd = PIN_TASKBAR_HWND.swap(0, Ordering::SeqCst);
    if hwnd == 0 {
        return;
    }

    pin_taskbar_mirror::stop_pin_mirror();
    pin_taskbar_mirror::forget_the_pin_picture();
    pin_taskbar_buttons::forget_media_toolbar();
    let _ = DestroyWindow(HWND(hwnd as *mut _));
}

/// Whether the stand-in is up, in either of the two forms it has: the one a pin that is on the screen
/// is shown in, and the put-away one a hover holds the last frame in.
///
/// It is read before a restore does anything, because it is what tells the two places a pin can be put
/// away into apart: a restore from the bubble places the window beside the bubble (see
/// `placed_pin_box`), and a restore from the taskbar simply puts it back where it stood.
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
    // **The stand-in goes where no button is owed**, which is the rule every other road into and out
    // of it asks (see `a_button_is_owed` and `show_pin_taskbar`). It used to be hidden exactly where
    // the setting said a button *was* owed and left standing where it said none was, which is the
    // button the user saw sometimes and not others — a stand-in under a pin with `Always On Top` on,
    // and no button at all for a pin that should have had one.
    if !a_button_is_owed() {
        unsafe { hide_pin_taskbar() };
    }
}

/// The two style words the stand-in is created with, so that the window is the app's own button in
/// the taskbar and nothing else.
///
/// **No caption, and that is what a hover is read for.** The Shell draws a window's hover from the
/// window itself, frame and all, and this window's picture is the DWM thumbnail of the pin in its
/// client area — so a caption is a native title bar drawn over the pin's own frame, which is what
/// the hover showed before there was one (see `pin_taskbar_mirror`). A borderless popup's client
/// area is the whole of it, so the hover is the pin and nothing of Windows around it.
///
/// `WS_SYSMENU` and `WS_MINIMIZEBOX` are kept for what the Shell reads them for rather than for
/// what is drawn of them: a window with no system menu is given a taskbar button with no icon on it
/// (see `app_icon`), and the buttons are what a minimize and a restore from the button's own menu
/// arrive as (see `pin_taskbar_proc`). `WS_EX_APPWINDOW` is what asks for the button at all.
fn stand_in_style() -> (WINDOW_STYLE, WINDOW_EX_STYLE) {
    (WS_POPUP | WS_SYSMENU | WS_MINIMIZEBOX, WS_EX_APPWINDOW)
}

/// Whether the pin wants a taskbar button at all: on the screen with `Pin Mode → Always On Top`
/// off — the setting whose button this stands in for — or put away into the taskbar rather than
/// into the bubble.
fn a_button_is_owed() -> bool {
    if !pinned() {
        return false;
    }

    // A pin standing in a window of its own holds a taskbar button of its own, named after its file
    // and drawn by the Shell from the window itself — a stand-in beside it would be a second button for
    // one pin, and a hover of a picture of a picture (see `pin_real_window`). Where that window is a
    // film's host, the still the hover is drawn from is this app's own picture of the host rather than
    // the host's rows (see `pin_film_host::keep_the_pins_picture_for_the_taskbar`).
    if pin_is_its_own_window() {
        return false;
    }

    if pin_is_collapsed() {
        // Put away: into the taskbar, or into the round bubble. Only the taskbar's collapse leaves
        // no bubble on the screen, and the bubble that is up is what the two are told apart by (see
        // `collapse_pin_to`).
        return pin_bubble::showing_bubble_window().is_none();
    }

    !pin_always_on_top()
}

/// Put the stand-in in the form the pin asks for, and make it where the Shell has taken it away.
///
/// A stand-in is a window the Shell has a taskbar button for, and the Shell has gestures of its own
/// on that button: a second click, `Win+M`, the button's own `Minimize`. Every one of them
/// minimizes or restores *this* window rather than the pin, and none of the pin's own roads runs on
/// that account — so the form is read back here, on the tick, which is the one thing that runs
/// whether or not a gesture of the Shell's reached this app. What a stand-in left in the wrong form
/// costs is the picture the Shell draws the button's hover from: a minimized window is not one the
/// Shell is composing, so a pin that is on the screen under a minimized stand-in is a hover frozen
/// on the frame it was put down with.
///
/// And it is the same read that brings a button the app has lost back: a handle to a window that is
/// not there is not a window, so the ask below makes another one — which is the whole of what the
/// user's "the button is gone until the app is restarted" was, and what a run must not be.
pub(super) fn keep_the_stand_in_up() {
    if !a_button_is_owed() {
        // **A pin that is gone owes no button**, and a stand-in left standing for one is a window of
        // this app's in the Shell's list with nothing behind it: it goes down with the pin it was
        // for. Asked here rather than only on the close road, because the close is one road out of a
        // pin — the watchdog's end is another, and the Shell has gestures of its own — and this is the
        // tick that runs after all of them (see `hide_pin_taskbar`).
        if showing_taskbar_window().is_some() {
            unsafe { hide_pin_taskbar() };
        }

        return;
    }

    let Some(hwnd) = showing_taskbar_window() else {
        // No button is up at all: the pin's own place asks for one, and the form of it is the one
        // that place asks for (see `show_pin_taskbar`).
        //
        // Safety: the stand-in is this app's own window, made here where the handle the run holds
        // is one to a window that has gone.
        unsafe { show_pin_taskbar() };
        return;
    };

    // A pin the Shell has minimized with the rest of the desktop is a pin whose button is to stay
    // down with it: what is owed then is the frame the Shell cached for a minimized window, and not
    // a button standing over a desktop nobody is working in.
    let preview = HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _);
    if !preview.is_invalid() && unsafe { IsIconic(preview).as_bool() } {
        return;
    }

    // Safety: asks this app's own stand-in which form it is in, which is a property rather than a
    // change to it.
    let minimized = unsafe { IsIconic(hwnd).as_bool() };
    if minimized == pin_is_collapsed() {
        return;
    }

    // Safety: the same window, asked for the form the pin's own place wants — minimized for a pin
    // put away into the taskbar, up and mirroring for one on the screen.
    unsafe { show_pin_taskbar() };
}

/// The stand-in's window, created once for the run.
///
/// It is a borderless top-level window rather than a layered popup — nothing about it is drawn by
/// this app; the picture in it is the Shell's own thumbnail — and it is deliberately *not*
/// `WS_EX_TOOLWINDOW`: that bit is what keeps every other window of this app out of the taskbar,
/// and this is the one window whose whole purpose is to be in it. `WS_EX_APPWINDOW` states that
/// outright rather than leaving it to the ownerless-window rule (see `stand_in_style`).
///
/// **A window that has gone is made again rather than handed back.** The Shell closes a taskbar
/// button's window from its own menu, a Shell that restarts takes the window with it, and a handle
/// kept on a window that is not there is a run whose roads that show the stand-in all show nothing:
/// no button, no hover, and nothing to ask about either until the app is restarted.
pub(super) unsafe fn pin_taskbar_window() -> Option<HWND> {
    let existing = PIN_TASKBAR_HWND.load(Ordering::SeqCst);
    if existing != 0 {
        let existing = HWND(existing as *mut _);
        if IsWindow(existing).as_bool() {
            return Some(existing);
        }

        // The window the slot holds has gone: a Shell that closed it from the button's own menu, a
        // Shell that has restarted, or whatever took the thread that made it. It is let go of the
        // way its own close lets go of it — marks, mirror and slot together — so that the window
        // made below is a fresh one: a handle kept on a window that is not there is a run whose next
        // pin has no button at all and nothing to ask about it.
        close_pin_taskbar();
    }

    let (style, ex_style) = stand_in_style();
    let hinstance = GetModuleHandleW(None).ok()?;
    let hwnd = CreateWindowExW(
        ex_style,
        PIN_TASKBAR_CLASS,
        TASKBAR_TITLE,
        style,
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

    // And the Shell is told that this window will supply the still its button's hover is drawn of, so
    // that a put-away pin's button holds the pin rather than the icon the Shell falls back on for a
    // window with no still of its own (see `pin_taskbar_mirror`). It is set once, where the window is
    // made: what the still *is* is answered when the Shell asks for it, since that is the only moment
    // the app is asked to supply one, and whether the Shell is to use it at all is read from where the
    // pin is (see `the_buttons_still_is_the_apps_own`).
    unsafe { supply_the_stills_of_this_window(hwnd) };

    Some(hwnd)
}

/// Tell the Shell that this window supplies the still its taskbar button's hover is drawn of.
///
/// **Every window of this app's that can wear a button is made wearing this**, because a window whose
/// picture is not in its own rows has nothing the Shell could freeze for a hover: the stand-in beside a
/// put-away pin is off every display and holds a mirror composed over it, and the film host that owns a
/// pin's chrome holds a DWM thumbnail composed over its client (see `the_buttons_still_is_the_apps_own`,
/// `pin_film_host::keep_the_pins_picture_for_the_taskbar`).
///
/// Safety: `hwnd` is a window of this app's own.
pub(super) unsafe fn supply_the_stills_of_this_window(hwnd: HWND) {
    let supplies_its_own_still = BOOL(1);
    let _ = unsafe {
        DwmSetWindowAttribute(
            hwnd,
            DWMWA_HAS_ICONIC_BITMAP,
            &supplies_its_own_still as *const BOOL as *const core::ffi::c_void,
            std::mem::size_of::<BOOL>() as u32,
        )
    };
}

unsafe fn handle_taskbar_restore() {
    // **Nothing is hidden before the pin is asked.** A restore from the taskbar puts the pin back
    // where it stood, and which place that is is read from whether the stand-in is up
    // (`pin_taskbar_is_showing`, in `restore_pin`) — so a stand-in taken down here would be the one
    // fact the pin's own restore needs, and the pin would come back beside the pointer on the
    // taskbar instead of where the user left it. The pin's restore takes the stand-in down itself
    // where the setting says there is to be none (see `hide_minimized_pin`).
    if pin_is_collapsed() {
        ask_pin(PinCommand::Restore);
    } else {
        // **The window asked for is the pin's own**, whichever window the Shell's button belongs to:
        // a film's host carries the button and owns the pin's chrome, and what a press on it asks for
        // is the pin rather than the window that draws the film. Asked of `PREVIEW_HWND` it was the
        // hover's window that came forward, which is put down for the whole life of a pin that stands
        // in one of its own (see `pinned_window`).
        let pin_hwnd = pinned_window();
        if !pin_hwnd.is_invalid() {
            // A press on the pin's stand-in is the user asking for the pin, and it owes no band for
            // it: the window the hand is in is the window Windows activates, and an activated window
            // is already the one in front (see `band_place`). What it is owed is itself — a pin the
            // Shell hid with the rest of the desktop is a window that is not on the screen, and a
            // window that is not on the screen is put up by its paint and not by a raise: what a
            // layered window shows between two paints is the surface it already holds, stretched
            // into whatever box it is given (see `show_pinned_window`).
            show_pinned_window(pin_hwnd);
            let _ = SetForegroundWindow(pin_hwnd);
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
            // The Shell's own `Minimize` on this window, which is its button's gesture rather than
            // the pin's: what the user asked a button for is the pin, so it is asked for here rather
            // than left to `DefWindowProcW`, which would minimize the app's button out from under the
            // pin and leave the hover a picture of a minimized window (see `keep_the_stand_in_up`).
            SC_MINIMIZE => {
                handle_taskbar_restore();
                LRESULT(0)
            }
            SC_CLOSE => {
                // The Shell's Close, which is the pin's end and the button's own going: asking the
                // pin to close is not enough, because what the Shell sent the ask to is this window
                // (see `close_pin_taskbar`).
                close_pin_taskbar();
                ask_pin(PinCommand::Close);
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wparam, lparam),
        },
        // The Shell's ask for the still of a button that is down — its own thumbnail of the window it
        // falls back on otherwise is the window's icon — answered with the pin's own picture (see
        // `pin_taskbar_mirror`'s `answer_a_thumbnail_ask`).
        WM_DWMSENDICONICTHUMBNAIL => pin_taskbar_mirror::answer_a_thumbnail_ask(hwnd, lparam),
        // The Shell asks an iconized window whether it may be opened before it opens it; this one
        // never may, because what it stands in for is not on the screen: the ask is answered by
        // putting the pin back and taking the stand-in away, and the button is never brought up.
        //
        // **Except when this app is the one asking.** The question arrives for the mirroring step's
        // own showing of a put-away stand-in as much as for the user's click, and answered as the
        // click it asked the pin back out of the taskbar it had just been put into — which is what
        // left every minimize after a restore doing nothing at all: the window stayed iconic and
        // composed nothing, so the picture could not be taken, and the restore the query queued put
        // the pin back on the screen (see `ASKING_THE_STAND_IN_UP`).
        WM_QUERYOPEN => {
            if ASKING_THE_STAND_IN_UP.load(Ordering::SeqCst) {
                return LRESULT(1);
            }
            handle_taskbar_restore();
            LRESULT(0)
        }
        WM_ACTIVATE => {
            // A click on the taskbar button may activate a stand-in that is *up* rather than ask it
            // to restore; the pin is asked for either way. A deactivation says nothing about the
            // pin — and neither does an activation of a stand-in whose pin is put away.
            //
            // **A stand-in that is put away is brought back by the Shell's own restore and by
            // nothing else.** What carries the user's click to a window that is down is the
            // question an iconic window is sent before it may open (`WM_QUERYOPEN`) and the button's
            // own `Restore` (`SC_RESTORE`); an activation of a put-away stand-in is otherwise this
            // app's own doing — a collapse takes the pin's window off the screen, Windows hands the
            // keyboard on to the next window, and the stand-in a click a moment before had activated
            // is the one it picks — and read as that click it asked for the pin's restore behind the
            // collapse that had just put the pin away (see `collapse_pin_to`).
            if (wparam.0 & 0xFFFF) as u32 == WA_ACTIVE && !pin_is_collapsed() {
                handle_taskbar_restore();
            }
            LRESULT(0)
        }
        WM_CLOSE => {
            // The Shell's own Close on the button, and the same two things as `SC_CLOSE` above.
            close_pin_taskbar();
            ask_pin(PinCommand::Close);
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The name as a hand reads it out of the field the Shell is given.
    fn named(path: Option<&Path>) -> String {
        let units = taskbar_title_for(path);
        let end = units
            .iter()
            .position(|unit| *unit == 0)
            .unwrap_or(units.len());
        String::from_utf16_lossy(&units[..end])
    }

    /// The button names the file the pin is of, and never the folder it came out of: what the
    /// Shell's label has room for is the name, and a path on it would be the same name cut short
    /// for the sake of where the file happens to live.
    #[test]
    fn the_button_names_the_file_the_pin_is_of() {
        assert_eq!(
            named(Some(Path::new("films/shark tale.mp4"))),
            "shark tale.mp4"
        );
        assert_eq!(
            named(Some(Path::new(""))),
            "Rust Preview",
            "a path with no name of its own is no file to name"
        );
        assert_eq!(named(None), "Rust Preview");
        assert_eq!(
            taskbar_title_for(None).last(),
            Some(&0),
            "and the field the Shell reads is terminated"
        );
    }
}

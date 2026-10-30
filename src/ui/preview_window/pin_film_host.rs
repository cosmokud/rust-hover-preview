//! The window a pinned film is drawn in: a window of the desktop's that owns the pin's chrome and
//! shows the film through the Shell's thumbnail of the player's own window.
//!
//! **Both engines are on one road.** mpv is launched with no `--wid` and makes a window of its own,
//! and FFmpeg's player is an SDL program with no embed switch at all, so it makes one too. The host
//! asks the Shell for a thumbnail of that window (`DwmRegisterThumbnail`) and draws it into its own
//! client area (see `pin_film_thumbnail`). The host is then a **window of the desktop's that owns
//! the pin's chrome**: Windows keeps the chrome above the film, hides it with the host, keeps it out
//! of Alt+Tab, and moves neither for this app (see `the_pins_chrome_is_owned` and
//! `own_the_pin_window`). A resize step moves this app's own window and re-points one rectangle;
//! nothing of the player's is asked for a box. **Why that is the only road, and what a resize may
//! move, is `Pin mode → The film road` in `ARCHITECTURE.md`.**
//!
//! **Why it is not the pin's window itself.** The pin's window is the chrome and the transport as
//! well as the band, and the film would be drawn over the caption. So the picture gets a window of
//! its own at the band's box, and the pin keeps every other pixel (see `the_film_host_box`).
//!
//! **Why the presses are kept off it, and why that is the hit test rather than a disabled window.**
//! A press on the picture is a press on the pin: the pin is dragged by its film, and the band's
//! alpha-1 claim is already hit-testable (see `fill_band_hit_testable`). The host answers
//! `WM_NCHITTEST` with `HTTRANSPARENT`, which is the documented road for "the window under this one,
//! in this thread": the pin's chrome and its strips stand under the host, so a press on the band is
//! answered by the pin, exactly as `WindowFromPoint` skipping a disabled window used to give it. **A
//! disabled window was the wrong tool for that job**, because it is also a window the Shell can
//! neither activate nor minimize — and this window is the one the pair's taskbar button belongs to
//! (see `own_the_pin_window`). A real click was traced delivering nothing to a disabled host at all
//! while the Shell gave up on activating it and re-activated the pin's chrome instead (its owner),
//! so the button could only ever deactivate and re-activate the app.
//!
//! **So the host is an ordinary window of the desktop's**: the Shell lists it, activates it, and puts
//! it down, and that button is the toggle. An activation of it gives the *pin* the caret and leaves
//! the foreground on the host — the window whose button was pressed, and so the window a second
//! press minimizes (see `answer_an_activation_of_the_host`) — and a minimize of it is the pin going
//! away into the taskbar (see `the_shell_put_the_host_down`). The Shell sometimes lists the pin's
//! chrome as well, so the chrome's own minimize is the same put-away (see
//! `the_shell_put_the_chrome_down`), and the Shell's restore of either window is `the_pin_came_back`.
//!
//! **And it is the window a put-away pin's taskbar button belongs to.** An owned window carries no
//! button of its own, so a pin put away into the taskbar is *the host minimized*, brought back by the
//! Shell's own restore (see `the_pin_came_back`), and its hover is the still this app takes of it
//! before it goes down — the host's own rows are the black behind the film, so a still of those rows
//! is no picture at all (see `keep_the_pins_picture_for_the_taskbar`). **So that one button is the
//! whole of the toggle**: a minimize of the host is the pin going away into the taskbar, on the road
//! the pin's own minimize takes, and the Shell's restore of the window is the pin coming back (see
//! `the_shell_put_the_host_down`).
//!
//! The host outlives the player once begun into it: a seek is mpv's own IPC and a step begins
//! another player behind the same host. It goes down with the pin (see
//! `take_down_the_pin_film_host`, called from the state a road out of a pin clears).

use std::sync::atomic::{AtomicIsize, AtomicU32, Ordering};
use std::sync::Once;

use windows::core::w;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, EndPaint, FillRect, GetStockObject, SetBrushOrgEx, SetStretchBltMode, StretchBlt,
    BLACK_BRUSH, HBRUSH, HALFTONE, NULL_BRUSH, SRCCOPY,
};
use windows::Win32::Graphics::Gdi::PAINTSTRUCT;
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::EnableWindow;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, GetClientRect, GetWindowLongPtrW,
    RegisterClassExW, SendMessageW, SetWindowLongPtrW, SetWindowPos, ShowWindow, GWLP_HWNDPARENT,
    GWL_EXSTYLE, GWL_STYLE, HTTRANSPARENT, HWND_NOTOPMOST, HWND_TOPMOST, ICON_BIG, ICON_SMALL,
    SC_CLOSE, SC_MINIMIZE, SIZE_MINIMIZED, SWP_NOACTIVATE, SWP_NOCOPYBITS, SWP_NOMOVE, SWP_NOSIZE,
    SWP_NOZORDER, SW_HIDE, SW_SHOWNOACTIVATE, WA_ACTIVE, WA_CLICKACTIVE, WM_ACTIVATE, WM_CLOSE,
    WM_DWMSENDICONICTHUMBNAIL, WM_ERASEBKGND, WM_NCHITTEST, WM_PAINT, WM_SETICON, WM_SETFOCUS,
    WM_SIZE, WM_SYSCOMMAND, WNDCLASSEXW, WS_CLIPCHILDREN, WS_EX_APPWINDOW, WS_MINIMIZEBOX, WS_POPUP,
    WS_SYSMENU, WINDOW_EX_STYLE,
};

use super::pin_trace::{note_a_message_of_a_window, note_pin_trace, window_name};
use super::*;

/// The class a film's host is made from. A class of its own rather than the pin window's: what a
/// `WS_CHILD` host answers to is a window that draws nothing and takes nothing, and the pin's
/// window procedure is the pin's (see `pin_real_window`).
const FILM_HOST_CLASS: windows::core::PCWSTR = w!("RustHoverPreviewPinFilmHost");

/// The default answer is most of what the host does: it draws nothing of the film's own and no
/// press lands on it. Everything both classes a host can be made from answer is one function (see
/// `answer_a_host`), because a hover's host that a pin took up is the pin's host from that moment on.
/// This bindings version exposes `DefWindowProcW` as a Rust-callable function, so the class still
/// needs a system-ABI shim around it.
unsafe extern "system" fn the_default_answer(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    if let Some(answer) = unsafe { answer_a_host(hwnd, msg, wparam, lparam) } {
        return answer;
    }

    unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
}

/// What either class a host is made from answers before `DefWindowProcW`: the host this module makes,
/// and the hover's host a pin took up rather than beginning a second player (see
/// `promote_the_hover_film_host`).
///
/// **One answer rather than two procedures' worth, because the window changes hands under a take-up**:
/// a hover's host that a pin promoted *is* the pin's host from then on, so everything the pin's own
/// class owes — the band painted behind the film, the cover a relaunch left standing, the Shell's own
/// close, the Shell's minimize of the button or the button's own second click, the window coming back
/// out of the taskbar, and the Shell's ask for the button's still — has to be answered by whichever
/// class the window happened to be made from.
///
/// Safety: the answers are those of a window of this app's own, asked from its own window procedure.
pub(super) unsafe fn answer_a_host(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> Option<LRESULT> {
    // **What a hand on the pair's taskbar button delivers is written down before anything answers
    // it**, because the whole of the toggle that button was given assumes which messages arrive —
    // an activation of this window, the Shell's `SC_MINIMIZE`, or the `SIZE_MINIMIZED` a window
    // reports as it goes down (see `the_shell_put_the_host_down`) — and a click that still does
    // nothing on a real desktop says which of the three it really is, and which window it reaches
    // (see `note_a_message_of_a_window`). Nothing is answered differently for the line.
    //
    // **A hover's host is written about by nobody**, though it is made from a class answered by this
    // same procedure: it wears no taskbar button and is nobody's click (see
    // `pin_taskbar::a_button_is_owed`), so the line is written only for the window the pair's button
    // belongs to, while the pin it belongs to is up.
    if pinned() && film_host() == Some(hwnd) {
        note_a_message_of_a_window("film host", hwnd, msg, wparam, lparam);
    }

    // **A press on the picture is a press on the pin, and the hit test is what keeps it off this
    // window.** `HTTRANSPARENT` is the documented answer for "route this mouse message to the window
    // under this one, in this thread": the pin's chrome and its strips stand under the host, so a
    // press on the band is delivered to the pin and the band is dragged by its film (see
    // `fill_band_hit_testable`).
    //
    // **What it replaces, and why that was the wrong tool.** This window used to be *disabled* for
    // this, because `WindowFromPoint` skips a disabled window. A disabled window is also one the
    // Shell can neither activate nor minimize, and this is the window the pair's taskbar button
    // belongs to (see `own_the_pin_window`): a real click on that button was traced delivering
    // nothing here at all, with the Shell re-activating the pin's chrome instead — the owner it falls
    // back to when the window it was asked for is disabled — so the button could only ever deactivate
    // and re-activate the app. The hit test does the one thing that was wanted and leaves activation
    // and the Shell's own minimize to this window.
    //
    // **The hover's host is never this window.** It is made from the same class and answered here
    // too, but it is not the pin's host, and it stays disabled and out of everything a button is (see
    // `hover_film_host::create`).
    if msg == WM_NCHITTEST && film_host() == Some(hwnd) {
        return Some(LRESULT(HTTRANSPARENT as isize));
    }

    // **The band behind the picture is painted here rather than erased by the class**, and the class
    // wears `NULL_BRUSH` so that this is the only hand that draws it (see
    // `paint_the_host_behind_the_film` and `create`). `1` is "handled": it is what keeps
    // `DefWindowProcW` from erasing the same pixels a second time.
    if msg == WM_ERASEBKGND {
        unsafe { paint_the_host_behind_the_film(hwnd, HDC(wparam.0 as *mut _)) };
        return Some(LRESULT(1));
    }

    // A cover this app painted stands where the film was, and this paint is its own (see
    // `answer_the_paint_of_a_cover`).
    if let Some(answer) = unsafe { answer_the_paint_of_a_cover(hwnd, msg) } {
        return Some(answer);
    }

    if let Some(answer) = answer_the_hosts_close(msg, wparam) {
        return Some(answer);
    }

    // **The Shell's own minimize of this window is the pin going away into the taskbar**, and `0` is
    // what keeps `DefWindowProcW` from putting the window down behind the pin's state (see
    // `the_shell_put_the_host_down`). **The collapse itself is the loop's, never this procedure's**: a
    // window procedure that collapsed the pin would take the pin's lock, give a keyboard back and
    // take a still of this very window from inside a message about it, so the command is queued and
    // the pin's own `Minimize → To Taskbar` road puts the window down with the state that says so.
    if the_shell_put_the_host_down(msg, wparam) {
        // **The decision the log is read for**: whether this minimize is the pair's button or a
        // window of the user's own going down with the desktop (see `the_host_wears_the_pins_button`).
        let wears_the_button = the_host_wears_the_pins_button(hwnd);
        note_pin_trace!(
            "film host decision the shell's minimize: {}",
            if wears_the_button {
                "the host wears the button, the taskbar minimize is queued"
            } else {
                "refused, this is not the window the button is worn by"
            }
        );
        if wears_the_button {
            ask_pin(PinCommand::MinimizeTo(PinMinimizeTo::Taskbar));
            return Some(LRESULT(0));
        }
    }

    // **A click on the button of a pin the user is already in is that pin's minimize**, and it is
    // asked before the handover below, which is what every other click on that button is (see
    // `answer_a_second_click_on_the_hosts_button`).
    if let Some(answer) = answer_a_second_click_on_the_hosts_button(hwnd, msg, wparam) {
        return Some(answer);
    }

    // **An activation of the host is a request for the pin.** A pair that is not Always On Top wears its
    // taskbar button on the host, and the host is the window that button activates — so the activation
    // a click on that button delivers has to be answered for the pin, whose window is the one the user
    // asked the button for (see `answer_an_activation_of_the_host`).
    if let Some(answer) =
        unsafe { answer_an_activation_of_the_host(hwnd, msg, wparam, &Win32PinWindow) }
    {
        return Some(answer);
    }

    // **A host the Shell brings back from the taskbar is the pin coming back.** The host is the window
    // the Shell's button belongs to on this road, and a pin left believing it is away refuses
    // everything a pin is for (see `the_pin_came_back`, `own_the_pin_window`).
    the_pin_came_back(hwnd, msg, wparam);

    // The Shell's ask for the still of a button that is down, answered with the picture the collapse
    // took of this very window (see `keep_the_pins_picture_for_the_taskbar`).
    if msg == WM_DWMSENDICONICTHUMBNAIL {
        return Some(unsafe { pin_taskbar_mirror::answer_a_thumbnail_ask(hwnd, lparam) });
    }

    None
}

/// **A host that is the Shell's window is the Shell's close too.** Its taskbar button's Close arrives
/// here as a `WM_CLOSE` or an `SC_CLOSE`, and taken it would destroy the window the pin's chrome is
/// owned by — the pin with it, under a state that still believes a pin is up. What a close means for the
/// pin is the pin's own road's answer (see `window_proc`), so it is asked for the pin and the window is
/// left standing for the teardown that ends it.
///
/// A hover's host is a window nobody is given a close for, and its close is swallowed for the same
/// reason: destroying it would take the picture of a hover (or of the pin it was promoted into) with it.
fn answer_the_hosts_close(msg: u32, wparam: WPARAM) -> Option<LRESULT> {
    if msg != WM_CLOSE && !(msg == WM_SYSCOMMAND && (wparam.0 & 0xFFF0) as u32 == SC_CLOSE) {
        return None;
    }

    if pinned() {
        request_pin_end();
    }
    Some(LRESULT(0))
}

/// Whether this message is **the Shell putting the host — the window the pair's taskbar button
/// belongs to — down**: its own system command, or the size a window is handed as it is minimized.
///
/// **Both readings are asked, because either is what arrives.** The Shell's ask for a window to go
/// away is `WM_SYSCOMMAND`/`SC_MINIMIZE` — the press on the button of a window the user is already in,
/// and what `Win`+`D` delivers to every window that offers a minimize box — while a
/// `ShowWindow(SW_MINIMIZE)` from anywhere at all minimizes the window and reports it as
/// `WM_SIZE`/`SIZE_MINIMIZED`. Neither is a restore, which is the other half of the same button and is
/// read already (see `the_pin_came_back`).
///
/// **`SC_MINIMIZE` is read with its low four bits masked away**, because those are the Shell's own
/// flags rather than part of the command — the same reading the close above makes of `SC_CLOSE`.
fn the_shell_put_the_host_down(msg: u32, wparam: WPARAM) -> bool {
    (msg == WM_SYSCOMMAND && (wparam.0 & 0xFFF0) as u32 == SC_MINIMIZE)
        || (msg == WM_SIZE && wparam.0 as u32 == SIZE_MINIMIZED)
}

/// Whether a minimize of this window is the pin's own going away: **this is the host the pair's button
/// is worn by, and there is a pin on the screen for the button to put away.**
///
/// **`Pin Mode → Always On Top` is the answer that must not be left out.** With the setting on the
/// pair wears no button at all (see `own_the_pin_window`), so nothing the Shell can press belongs to
/// it, and a minimize arriving on the host is the user's own — `Win`+`D`, or a call of something
/// else's — which is a window going down with the desktop and coming back with it. Turned into a
/// collapse, it would leave a pin over everything with no button to bring it back (see
/// `the_button_follows_the_setting`).
///
/// **And a pin already away, or on its way, is nothing to put away.** A collapse sets `collapsed`
/// before it asks the window down, and `the_pin_is_being_put_down()` is that same road's note that the
/// messages arriving now are its own (see `collapse_pin_to` and `put_the_pins_window_in_the_taskbar`):
/// either answer says this minimize is not a new ask, and a second collapse of a pin that is already
/// away is a command the loop would spend on nothing.
fn the_host_wears_the_pins_button(hwnd: HWND) -> bool {
    film_host() == Some(hwnd) && the_chrome_wears_the_pins_button()
}

/// Whether a minimize of the pin's **chrome** is the pair's button being pressed: the pin's chrome is
/// owned by a film host, and the pair is one that wears the button at all (see
/// `the_host_wears_the_pins_button`, which is this and the window).
///
/// **The Shell sometimes lists the chrome too, and drives it.** The pin's window is unowned and not a
/// tool window on this road (the host owns it), so it can be listed beside the host, and a real click
/// has been seen deactivating and re-activating it. Whichever window of the pair the Shell puts down,
/// the pin's state and the collapse have to be the same, so the chrome's minimize is read as the same
/// put-away (see `the_shell_put_the_chrome_down`).
pub(super) fn the_chrome_wears_the_pins_button() -> bool {
    film_host().is_some()
        && pinned()
        && pin_is_its_own_window()
        && !pin_lifecycle::pin_always_on_top()
        && !pin_is_collapsed()
        && !the_pin_is_being_put_down()
}

/// Whether this message is **the Shell putting the chrome of the pair down**, in the words the chrome's
/// own procedure reads: the same command and the same size a host is put down by, for the window the
/// Shell may be driving instead of the host (see `the_chrome_wears_the_pins_button`).
///
/// The answer is the host's own — the taskbar minimize — so that whichever window of the pair the Shell
/// puts down, the pin goes away once, by the state that says so (see `the_shell_put_the_host_down`).
pub(super) fn the_shell_put_the_chrome_down(msg: u32, wparam: WPARAM) -> bool {
    the_shell_put_the_host_down(msg, wparam) && the_chrome_wears_the_pins_button()
}

/// What the **second** click on the pair's taskbar button asks for: **the pin, put away into the
/// taskbar**.
///
/// The Shell activates the window its button belongs to on every click, and with `Pin Mode → Always On
/// Top` off that window is the film host (see `own_the_pin_window`). The **first** click arrives while
/// the user is in something else — `pin_is_focused()` is false — and it is the keyboard handover this
/// window has always made (see `answer_an_activation_of_the_host`): the pin is the window the button
/// was asked for, so it is given the caret.
///
/// The second click arrives the same way, while the pin already holds the keyboard the first one gave
/// it — and **a window whose own button is clicked while the user is in it is a window being
/// minimized**. That is what every other window on this desktop does with its second click, and it is
/// the toggle this pair was missing: the click puts the pin away by the road its own `Minimize → To
/// Taskbar` takes, and the same button brings it back by the Shell's restore (see
/// `the_pin_came_back`).
///
/// **To the taskbar, whatever `Pin Mode → Minimize to` says.** The button is the taskbar's: a pin
/// collapsed into its bubble has its host hidden with it, and a hidden host has no button at all (see
/// `the_button_follows_the_screen`), so a bubble collapse would be a click that took away the very
/// button it came from.
///
/// **And a click that arrives while the pin holds no keyboard is left alone.** The user is in another
/// program, so the click is the handover it has always been and not the start of a toggle: answered
/// here, the pin would go away instead of coming to the front.
fn answer_a_second_click_on_the_hosts_button(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
) -> Option<LRESULT> {
    if msg != WM_ACTIVATE
        || !matches!((wparam.0 & 0xFFFF) as u32, WA_ACTIVE | WA_CLICKACTIVE)
        || !pin_is_focused()
        || !the_host_wears_the_pins_button(hwnd)
    {
        // **Which check answered is what makes a log of one click decisive**, and only the message
        // the toggle is is written: this procedure is asked about every message the host is handed,
        // and a line for each of them is a flood rather than a trace (see
        // `why_the_second_click_was_not_the_toggle`).
        if msg == WM_ACTIVATE {
            note_pin_trace!(
                "film host decision a second click on the button: refused, {}",
                why_the_second_click_was_not_the_toggle(hwnd, wparam)
            );
        }
        return None;
    }

    note_pin_trace!(
        "film host decision a second click on the button: the taskbar minimize is queued"
    );
    ask_pin(PinCommand::MinimizeTo(PinMinimizeTo::Taskbar));
    // Nothing is asked of the window: the pin is already the window the user is in, and this click
    // asked for that window to go away rather than for the caret to be put in it.
    Some(LRESULT(0))
}

/// **Which check of `answer_a_second_click_on_the_hosts_button` answered `no`**, in the words a trace
/// line can be read by: an activation of some other kind, a pin that holds no keyboard, a window that
/// is not the one the button is worn by, or a pin that is already away or over everything.
///
/// **It decides nothing.** The gate is that function's and this is only asked where it has already
/// refused a click, so that a log says which of the four it was; the two ask the same facts in the
/// same order, and the day they disagree the gate is the answer.
fn why_the_second_click_was_not_the_toggle(hwnd: HWND, wparam: WPARAM) -> &'static str {
    if !matches!((wparam.0 & 0xFFFF) as u32, WA_ACTIVE | WA_CLICKACTIVE) {
        return "not an activation";
    }

    if !pin_is_focused() {
        return "the pin holds no keyboard";
    }

    if film_host() != Some(hwnd) || !pin_is_its_own_window() {
        return "not the button's host";
    }

    "the pin is away/always-on-top"
}

/// What an activation of the pin's host asks for: **the caret in the pin, and nothing else**.
///
/// With `Pin Mode → Always On Top` off the pair's taskbar button belongs to the film host — the host
/// owns the pin's chrome, so it is the window the Shell lists for the pair (see `own_the_pin_window`).
/// The host is an ordinary window now, so the Shell activates it: a click's activation delivers
/// `WM_ACTIVATE` and `WM_SETFOCUS` to it, and what the user asked the button for is the pin. So the
/// pin is made focusable and given the caret, and the window procedure that answers this is the one
/// the pin's own keys arrive at from that moment.
///
/// **And the foreground is left where it is, which is the host.** Handing it to the pin would put the
/// Shell back where the disabled host had it: `GetForegroundWindow()` would no longer be the window
/// whose button the user pressed, and the next press of that button could not minimize it — the
/// Shell's own toggle would be dead again. The caret is also all the pin needs: `GetFocus()` is the
/// pin, so its own `WM_KEYDOWN` answers its keys, and the low-level key hook leaves them alone because
/// `key_input::pin_owns_caret` counts the pin's own window as the pin's (see `pin_take_the_caret`).
///
/// **Only the pin's own host, and only while a pin stands.** A hover's host is made from the same class
/// and answered by the same procedure, and a hover is nobody's button and no pin's window (see
/// `answer_a_host`); a pin that has ended has nothing to hand a keyboard to. **And only once per
/// activation**: a caret already in the pin is where the user asked the button for it, and asking again
/// is a second ask for a question Windows has answered. **But the click that arrives
/// once the pin has that keyboard is the button's own minimize** — a window the user is in whose button
/// is pressed is a window going away — and it is answered before this (see
/// `answer_a_second_click_on_the_hosts_button`). What is left to the default answer is every other
/// activation of a window the pin is already in.
///
/// **The host's own `WM_SETFOCUS` is the same road**, and it really is delivered now that the host is
/// enabled — an activation puts the caret on it before this answer moves it.
///
/// `window` is the seam the ask is made through, so a test can read the ask back rather than make it on
/// the user's desktop.
pub(super) unsafe fn answer_an_activation_of_the_host(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    window: &dyn PinWindow,
) -> Option<LRESULT> {
    let activated = match msg {
        // The low word is the activation state; the high word says whether the window is minimized,
        // and a minimized window coming back is the same request (see the message's own layout).
        WM_ACTIVATE => matches!(
            (wparam.0 & 0xFFFF) as u32,
            WA_ACTIVE | WA_CLICKACTIVE
        ),
        WM_SETFOCUS => true,
        _ => false,
    };
    if !activated || film_host() != Some(hwnd) || !pinned() || !pin_is_its_own_window() {
        return None;
    }

    // **A caret already in the pin is where the user asked the button for it**, so this activation is
    // not the handover: the click that arrives once the pin has that keyboard is the button's own
    // minimize, and it is answered before this (see
    // `answer_a_second_click_on_the_hosts_button`). Written down as the refusal it is: a log of a
    // click that does nothing is the whole of why these lines exist.
    if pin_is_focused() {
        note_pin_trace!(
            "film host decision an activation of the host: refused, the pin already holds the keyboard"
        );
        return None;
    }

    note_pin_trace!(
        "film host decision an activation of the host: the pin takes the caret, the foreground is left \
         on the host"
    );
    unsafe { pin_take_the_caret(pinned_window(), window) };
    Some(LRESULT(0))
}

/// Whether a resize drag is in flight: the one gesture the thumbnail road must never place a source
/// under, so that a film whose size arrives late is not sized while the hand is on an edge (see
/// `pin_film_thumbnail::place_the_source_when_the_player_answers`).
///
/// **Asked without waiting for the pin's own lock**, because the asker may run inside a call that
/// holds it — the tick's own pass — and a lock that cannot be taken here is answered with the road
/// as it was rather than waited on (see `pin_state_try`).
pub(super) fn a_resize_drag_is_live() -> bool {
    pin_window::pin_state_try()
        .and_then(|pinned| pinned.pin().map(|pin| pin.dragging))
        .flatten()
        .is_some_and(|drag| matches!(drag.action, PinDragAction::Resize(_)))
}

/// **A cover this app painted is the host's picture until the replacement's registration takes the
/// box** (see `put_the_cover_in`), so the two messages that would otherwise draw the host's black
/// over it answer differently: the erase fills nothing, and the paint draws the retained frame
/// again, re-cropped to the client this paint is for.
///
/// It is one function rather than two window procedures' worth, because a pin's host can be a
/// hover's promoted into it (see `promote_the_hover_film_host`), and the cover is owed by whichever
/// class the host was made from.
pub(super) unsafe fn answer_the_paint_of_a_cover(hwnd: HWND, msg: u32) -> Option<LRESULT> {
    if msg != WM_PAINT || !the_cover_stands() {
        return None;
    }
    let mut paint = PAINTSTRUCT::default();
    // Safety: the context is the one Windows hands this paint with, and it is given back below.
    let dc = unsafe { BeginPaint(hwnd, &mut paint) };
    if !dc.is_invalid() {
        unsafe { paint_the_cover(hwnd, dc) };
    }
    unsafe { let _ = EndPaint(hwnd, &paint); };
    Some(LRESULT(0))
}

/// Paint the host's whole client black: the film is the Shell's thumbnail (see `pin_film_thumbnail`)
/// or FFmpeg's child, which fills the client box itself.
///
/// **The host's own erase is what closes the hole a film fitted into its box once showed as three
/// bands**: the Shell composes the thumbnail *over* this app's content, so a client painted in the
/// bars alone is a pixel nobody drew — and one of those reads as the desktop seen through the pin.
/// The picture fills its destination, which is the whole client, so nothing of this fill shows
/// through once the Shell has composed; what it covers is the client the Shell has not composed yet,
/// or one whose source is not drawing (see `pin_film_thumbnail::settle_the_registration`). Both hosts
/// a pin can show its film through are erased through here — the pin's own class, and the hover's a
/// pin took up (see `hover_film_host`).
pub(super) unsafe fn paint_the_host_behind_the_film(host: HWND, dc: HDC) {
    // **Nothing paints the host's black over a cover it owes the screen**, and the count above says
    // the same thing: an erase that fills here is the black a still must never show (see
    // `the_cover_stands`).
    if the_cover_stands() {
        return;
    }
    #[cfg(test)]
    HOST_ERASES.fetch_add(1, Ordering::SeqCst);
    let mut client = RECT::default();
    if unsafe { GetClientRect(host, &mut client) }.is_err() {
        return;
    }

    // SAFETY: the brush is a stock object this app neither made nor owns, and the rectangle is the
    // client box read above; GDI reports rather than faults for a device context that has gone.
    let brush = HBRUSH(unsafe { GetStockObject(BLACK_BRUSH) }.0);
    unsafe { FillRect(dc, &client, brush) };
}

/// The frame a still of this app's painted into the host, kept so the host can go on standing on it
/// for the whole of a relaunch's wait: BGRA bytes at the size FFmpeg answered (see
/// `pin_film_thumbnail::stand_the_still`).
///
/// It is bytes rather than a `DibSurface` because a static has to be `Send` and a device context is
/// not; the surface a paint needs is made, filled and given back inside that paint.
struct CoverFrame {
    pixels: Vec<u8>,
    width: u32,
    height: u32,
}

/// The cover this app painted, or nothing while the host owes no picture of its own.
static COVER: Mutex<Option<CoverFrame>> = Mutex::new(None);

/// Whether a cover this app painted is what the host owes the screen (see `COVER`).
pub(super) fn the_cover_stands() -> bool {
    COVER.lock().is_ok_and(|held| held.is_some())
}

/// Keep a landed still as the host's cover until the picture changes hands (see
/// `put_the_cover_away`).
///
/// **The frame goes with the host it was decoded for**: a paint that lands after that host has gone
/// is nobody's picture, and a cover left standing would keep the next pin's host from erasing.
///
/// A frame that is not the size it claims is refused rather than kept: what arrives here came out of
/// a child process, and a short frame is a smear rather than a picture.
pub(super) fn put_the_cover_in(host: HWND, pixels: Vec<u8>, width: u32, height: u32) -> bool {
    if film_host() != Some(host)
        || width == 0
        || height == 0
        || pixels.len() < width as usize * height as usize * 4
    {
        return false;
    }
    match COVER.lock() {
        Ok(mut held) => {
            *held = Some(CoverFrame {
                pixels,
                width,
                height,
            });
            true
        }
        Err(_) => false,
    }
}

/// Give the cover up: the replacement's registration has taken the box, so the Shell draws the
/// picture and the host owes nothing of its own (see `pin_film_thumbnail::hand_the_picture_over`).
pub(super) fn put_the_cover_away() {
    if let Ok(mut held) = COVER.lock() {
        *held = None;
    }
}

/// Draw the cover into `dc`, **cover-cropped to the host's client as it is now**: the same centred
/// crop the registration wears, so a resize or an invalidation in the middle of a gap re-crops the
/// picture instead of showing the host's black (see `pin_film_thumbnail::the_pictures_crop`).
///
/// Safety: `dc` is a context for `host`, as the erase and the paint are handed theirs.
pub(super) unsafe fn paint_the_cover(host: HWND, dc: HDC) -> Option<(i32, i32, i32, i32)> {
    let mut client = RECT::default();
    if unsafe { GetClientRect(host, &mut client) }.is_err() {
        return None;
    }
    let (width, height) = (client.right - client.left, client.bottom - client.top);
    if width <= 0 || height <= 0 {
        return None;
    }
    let held = COVER.lock().ok()?;
    let cover = held.as_ref()?;
    let (left, top, crop_width, crop_height) = super::pin_film_thumbnail::the_pictures_crop(
        (cover.width as i32, cover.height as i32),
        (width, height),
        Some(width as f64 / height as f64),
    )
    .0;
    let surface = DibSurface::create(cover.width, cover.height)?;
    let length = cover.width as usize * cover.height as usize * 4;
    // Safety: the surface is this call's own of exactly these pixels, and the cover was checked for
    // them when it was put in.
    unsafe { std::ptr::copy_nonoverlapping(cover.pixels.as_ptr(), surface.bits(), length) };
    // Safety: the two contexts are this call's and the host's, and one scaling copy is asked of them.
    let painted = unsafe {
        let _ = SetStretchBltMode(dc, HALFTONE);
        // The brush origin travels with the mode: halftone measures the destination from it.
        let _ = SetBrushOrgEx(dc, 0, 0, None);
        StretchBlt(
            dc,
            0,
            0,
            width,
            height,
            surface.dc,
            left,
            top,
            crop_width,
            crop_height,
            SRCCOPY,
        )
        .as_bool()
    };
    #[cfg(test)]
    if let Ok(mut held) = THE_COVERS_CROP.lock() {
        *held = Some((left, top, crop_width, crop_height));
    }
    painted.then_some((left, top, crop_width, crop_height))
}

/// The crop the cover was last drawn with, for a test that reads the road rather than the pixels
/// (see `paint_the_cover`).
#[cfg(test)]
static THE_COVERS_CROP: Mutex<Option<(i32, i32, i32, i32)>> = Mutex::new(None);

/// The crop the cover was last drawn with (see `THE_COVERS_CROP`).
#[cfg(test)]
pub(super) fn the_covers_crop() -> Option<(i32, i32, i32, i32)> {
    THE_COVERS_CROP.lock().ok().and_then(|held| *held)
}

/// Registered once per run: the class a film's host is made from (see `create`).
static FILM_HOST_CLASS_REGISTERED: Once = Once::new();

/// How many erases this run has painted into a film host: the seam that says **nothing paints over a
/// still a relaunch left standing** (see `paint_the_host_behind_the_film` and
/// `pin_film_thumbnail::stand_the_outgoing_picture_from_its_dc`).
#[cfg(test)]
static HOST_ERASES: AtomicU32 = AtomicU32::new(0);

/// The erases this run has painted into a film host (see `HOST_ERASES`).
#[cfg(test)]
pub(super) fn host_erases() -> u32 {
    HOST_ERASES.load(Ordering::SeqCst)
}

/// The host of the pin that is up, and the player this app has attached to it.
static PIN_FILM_HOST: AtomicIsize = AtomicIsize::new(0);
static PIN_FILM_HOST_PLAYER: AtomicU32 = AtomicU32::new(0);
/// The hosted player and the file it was begun for: the player's own answer to what it is showing,
/// which is not the pin's answer while a swap is waiting for its picture (see `note_hosted_player`).
static PIN_FILM_HOST_FILE: Lazy<Mutex<Option<(u32, PathBuf)>>> = Lazy::new(|| Mutex::new(None));
/// The hosted player whose frame this app has proved on screen, so the pipe read that proves it is
/// asked once and not on every tick (see `the_hosted_frame_is_up`).
static PRESENTED_PID: AtomicU32 = AtomicU32::new(0);

/// The host the pin is showing its film through, or nothing while there is none.
pub(super) fn film_host() -> Option<HWND> {
    let host = PIN_FILM_HOST.load(Ordering::SeqCst);
    (host != 0).then_some(HWND(host as *mut _))
}

/// Whether the pin's band is standing over this host at all: taken up at the band, or shown with a
/// picture in it (see `the_band_is_a_standing_host`).
///
/// **It is what keeps a band filled black from being permanent.** The picture's road takes the film
/// off the band by forgetting the registration, putting the cover away and hiding the host
/// (`take_the_film_off_the_band`), and the band painted for a host that is no longer there is a
/// black rectangle this app has claimed on pixels nothing stands behind — the very fill the survey
/// of a swap used to leave on screen. A host that has been taken off the band has no claim on it
/// until a picture is put back in one (see `paint`'s host arm and `set_the_pin_film_host_visible`).
static HOST_HOLDS_THE_BAND: AtomicBool = AtomicBool::new(false);

/// Whether the pin's band must be painted as this app's own claim on the pixel because a host with
/// nothing in it stands there: the one state in which the black is the truth (see
/// `the_band_is_a_standing_host`).
pub(super) fn the_band_is_a_standing_host() -> bool {
    HOST_HOLDS_THE_BAND.load(Ordering::SeqCst)
        && film_host().is_some()
        && !the_cover_stands()
        && !the_host_is_drawing()
}

/// Note that the pin's band is standing over this host again — taken up at the band, shown with a
/// picture in it, or handed a picture by the registration (see `HOST_HOLDS_THE_BAND`).
fn claim_the_band_for_the_host() {
    HOST_HOLDS_THE_BAND.store(true, Ordering::SeqCst);
}

/// Note that the registration in the pin's host is a picture of a player's, so the band standing over
/// it is the film's rather than this app's black (see `the_band_is_a_standing_host`).
pub(super) fn note_the_band_has_a_picture() {
    if !HOST_HOLDS_THE_BAND.swap(true, Ordering::SeqCst) {
        // A band that was not the host's a moment ago may be painted black, and the picture that has
        // just arrived is the one thing that would never be painted over it (see
        // `the_band_owes_a_repaint`).
        the_band_owes_a_repaint();
    }
}

/// Note that the film has left the pin's band: the host is not the band's picture any more, so the
/// band owes its own paint rather than a claim on a host that is not on the screen (see
/// `take_the_film_off_the_band`).
fn release_the_band() {
    HOST_HOLDS_THE_BAND.store(false, Ordering::SeqCst);
}

/// Whether a repaint of the pinned band is owed because what stands in its host has changed — a
/// picture handed over, a host shown, a film taken off the band (see `take_the_band_repaint`).
static BAND_REPAINT_OWED: AtomicBool = AtomicBool::new(false);

/// Owe the pinned band exactly one repaint.
///
/// **The band's pixels are a function of what stands in the host**, so every transition of that has
/// to be painted or the last paint is the last word — and the film road is the one road on which
/// nothing else asks for a paint: a film a player draws moves no frame of this app's, so its band
/// would keep whatever the paint during the wait wrote (see `paint`'s host arm). This is the debt
/// that closes the gap, and it is a flag rather than a message so that a road on another thread can
/// leave it without knowing anything about the loop's channel.
pub(super) fn the_band_owes_a_repaint() {
    BAND_REPAINT_OWED.store(true, Ordering::SeqCst);
}

/// Take the band's repaint debt, if one is owed: the loop's tick is what spends it (see
/// `the_band_owes_a_repaint`).
pub(super) fn take_the_band_repaint() -> bool {
    BAND_REPAINT_OWED.swap(false, Ordering::SeqCst)
}

/// Whether a picture is standing on the pin's band that a film begun over it can be held for: the
/// registration the Shell draws through, or a still of one this app painted (see
/// `pin_film_thumbnail::the_picture_is_standing`).
///
/// **It is asked of the picture rather than of the newest player**, which is the difference a step
/// taken while a step is still waiting makes: `VIDEO_PID` and the hosted player both name the
/// player that has not drawn anything yet, while what the user is looking at is the window the
/// registration still draws (see `swap_pinned_media` and `the_player_on_the_band`).
pub(super) fn the_picture_stands_to_keep() -> bool {
    film_host().is_some() && (the_host_is_drawing() || the_cover_stands())
}

/// The player that is the picture on the pin's band, or `0` where the picture is a still of one this
/// app already ended: what a step holds, and what it retires once its own film has taken the picture
/// over (see `swap_pinned_media` and `pin_film_thumbnail::the_standing_source_process`).
pub(super) fn the_player_on_the_band() -> u32 {
    film_host()
        .and_then(pin_film_thumbnail::the_standing_source_process)
        .unwrap_or(0)
}

/// The player this app has attached to the host, recorded beside it so that a window of mpv's that
/// is not the pin's — a hover's, which is a window of the desktop — is never mistaken for it.
pub(super) fn hosted_player_pid() -> Option<u32> {
    let player = PIN_FILM_HOST_PLAYER.load(Ordering::SeqCst);
    (player != 0).then_some(player)
}

/// Record the player begun into the host, and the file it was begun for: called where the player is
/// spawned, before its window exists, so that the wait for its first frame is a wait for *its* child
/// and not for the last film's.
///
/// **The file is recorded with the player rather than read off the pin's media**, because the two are
/// not the same fact while a swap is waiting for its picture: the pin is still showing the file it is
/// leaving, and the placement of the incoming source asks what the *player* is showing. Answered off
/// the pin's media it named the outgoing film, or nothing at all — and a source that cannot be placed
/// cannot be drawn, so the box and the picture waiting for it would each be waiting for the other (see
/// `the_hosted_players_file` and `pin_film_thumbnail::the_films_own_size`).
///
/// A record that names no player names no file either, which is what a take-down leaves.
pub(super) fn note_hosted_player(pid: u32, path: Option<&Path>) {
    PIN_FILM_HOST_PLAYER.store(pid, Ordering::SeqCst);
    if let Ok(mut held) = PIN_FILM_HOST_FILE.lock() {
        *held = path
            .filter(|_| pid != 0)
            .map(|path| (pid, path.to_path_buf()));
    }
    PRESENTED_PID.store(0, Ordering::SeqCst);
    // A hosted player has no window of the desktop's to publish, and this one is the film now.
    VIDEO_HWND.store(0, Ordering::SeqCst);
}

/// The file the hosted player `pid` was begun for, where this app has that player on record.
///
/// It is the answer to "what is this player showing", which is a different question from "what is the
/// pin showing" for the whole of a swap's wait: the player is begun before the file it plays is
/// installed, and the placement of its source has to be made at that file's own size (see
/// `note_hosted_player` and `pin_film_thumbnail::the_film_the_player_plays`).
pub(super) fn the_hosted_players_file(pid: u32) -> Option<PathBuf> {
    let held = PIN_FILM_HOST_FILE.lock().ok()?;
    held.as_ref()
        .filter(|(player, _)| *player == pid)
        .map(|(_, path)| path.clone())
}

/// The host an FFmpeg film at this pin's band is kept for, where this is the pin's own window and
/// the pin stands in it: the film is drawn through it by the thumbnail, exactly as mpv's is (see
/// `pin_film_thumbnail`). `None` is every other launch, where the player keeps its own window of the
/// desktop's over the preview.
///
/// **Whether the film is drawn through a host is a fact of the launch rather than of the host's
/// existence**: a player begun while the host is still being made is one the monitor adopts as soon
/// as the host is there (see `host_for_a_hosted_player`), and nothing about the command line
/// depends on which came first.
pub(super) fn the_pins_film_is_kept_in_its_host(path: &Path, kind: Option<MediaType>) -> bool {
    the_film_is_shown_in_a_host(path, kind) && video_route(path) == VideoRoute::Ffplay
}

/// Whether the pin's chrome is owned by its film host, and so is not the window Windows lists and
/// gives a taskbar button to: the host is that window on this road (see `own_the_pin_window`).
pub(super) fn the_pins_chrome_is_owned() -> bool {
    film_host().is_some()
}

/// The window of a pin that holds the band the pair belongs in: the host that owns the pin's chrome,
/// on the road that has a pair, and the pin's own window on every other.
///
/// **One window of the pair owns the band, and it is the host.** Windows keeps an owned window in its
/// owner's band — the owner put in the topmost band puts the chrome in it, and the owner taken out
/// takes the chrome out with it (measured) — so a band asked of the chrome alone leaves the pair half
/// over another program's window and half behind it: the chrome in front, answering every press in
/// the band it is nearly transparent in, with the film it belongs to still under that program. That
/// is a pin that looks behind a window and yet takes the hand, and it is the one thing this window is
/// asked for (see `place_the_pin_in_the_band`).
pub(super) fn the_pin_window_that_holds_the_band(pin: HWND) -> HWND {
    the_pins_chrome_is_owned()
        .then(film_host)
        .flatten()
        .unwrap_or(pin)
}

/// Put a pin in a band — the topmost one, the ordinary one, or the place behind another window — as
/// **one window**, so that every road that raises, demotes or lowers a pin moves the film and the
/// chrome together (see `the_pin_window_that_holds_the_band`).
///
/// **A topmost band the desktop did not give is asked for again, from outside it**, and this is the one
/// place that answer lives rather than in each raise (see `raise_pinned_preview`, which used to ask it
/// for a held band): the desktop answers a band assignment as made and can still not make it for a
/// window it has not laid out yet — one shown this moment — and a pair the desktop banded on one side
/// only is the split this window exists to end (measured: the chrome in the band and the host out of
/// it).
///
/// **And the band is asked of both windows of the pair, whatever one of them holds it.** An owned
/// window out of its owner's band is a window the desktop can leave behind an ordinary one — the pin
/// behind the listing it was asked for from, with the film it belongs to in front of the listing — and
/// what the pairing was thought to buy is an assumption about the desktop rather than a bit this app
/// set. So the pin's own window is put in the band the pair belongs in beside the window that holds it,
/// which is what `Always On Top` means for a pin: measured on the real screen, a pin promoted out of a
/// hover had its own band written `HWND_NOTOPMOST` at the promotion and was never asked again.
///
/// Both windows are this app's own: nothing here is a call into another process.
pub(super) unsafe fn place_the_pin_in_the_band(pin: HWND, after: HWND) {
    let window = the_pin_window_that_holds_the_band(pin);

    let place = |window: HWND| {
        let _ = SetWindowPos(
            window,
            after,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
        // One line per placement, so that a run — and a test holding the order of a road to what it
        // asked the desktop for — reads the band write rather than inferring it (see
        // `the_host_has_appeared`).
        note_pin_trace!("band {:?} asked of {:?}", after, window);

        if after == HWND_TOPMOST && !pin_lifecycle::window_holds_the_topmost_band(window) {
            let _ = SetWindowPos(
                window,
                HWND_NOTOPMOST,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
            );
            let _ = SetWindowPos(
                window,
                HWND_TOPMOST,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
            );
        }
    };

    place(window);
    if window != pin && !pin.is_invalid() {
        place(pin);
    }
}

/// Ask for the pair's taskbar button on the window of it the Shell can list, or give it up with the
/// setting.
///
/// **`Pin Mode → Always On Top` is the row that decides whether the pair has a button at all**: a pin
/// over everything is one the user finds without one, and a pin that is not is an ordinary window with
/// an ordinary window's button (see `pin_taskbar::a_button_is_owed`, which is the same rule for the
/// stand-in of a pin standing in this app's own window).
///
/// **With the setting off the button belongs to the window the Shell lists for the pair**, which is the
/// host while the host is on the screen: the chrome is owned by it and never listed beside it (see
/// `own_the_pin_window`). **And while the host is down the chrome carries it**, because that is the
/// window the user still has: a swap hides the host for the whole wait of the next player (see
/// `take_the_film_off_the_band`), and a pair whose only window on the screen is owned by a hidden one
/// wears no button at all — a pin that cannot be brought back from the taskbar for as long as that
/// player takes to draw, or for good where it never draws (see `the_host_is_drawing` and
/// `video_playback::ensure_video_window_topmost`, whose show is the only road back).
///
/// `the_host_is_up` is the visibility the host is to have — the show that is about to be made, or the
/// hide that has just been taken — because a bit written *before* the show it belongs to is the bit the
/// Shell reads with it, which is what keeps a visible window from being put down and up again for a
/// button it could have been given for free (see `pin_real_window::set_the_pin_window_appwindow`).
pub(super) unsafe fn the_button_follows_the_screen(the_host_is_up: bool) {
    let Some(pin) = pin_real_window::real_pin_window() else {
        return;
    };
    let app_window = !pin_lifecycle::pin_always_on_top();
    match film_host() {
        Some(host) => {
            unsafe { pin_real_window::set_the_pin_window_appwindow(host, app_window) };
            unsafe {
                pin_real_window::set_the_pin_window_appwindow(pin, app_window && !the_host_is_up)
            };
        }
        // A pair with no host is the chrome's own window, and it wears the button the setting names.
        None => unsafe { pin_real_window::set_the_pin_window_appwindow(pin, app_window) },
    }
}

/// Ask for the pair's taskbar button, or give it up, as the setting now reads.
///
/// **`Pin Mode → Always On Top` is the row that decides whether a pin has a button at all**: a pin
/// over everything is one the user finds without one, and a pin that is not is an ordinary window with
/// an ordinary window's button (see `own_the_pin_window` and `pin_taskbar::a_button_is_owed`, which is
/// the same rule for the stand-in of a pin standing in this app's own window).
///
/// The bit is the Shell's to read, and it reads it as a window is shown (see
/// `pin_real_window::set_the_pin_window_appwindow`), so this is the call a row that changes the setting
/// makes: a pin already on the screen is the one the user is looking at, and it is the one that has to
/// gain or lose the button.
pub(super) unsafe fn the_button_follows_the_setting() {
    let host_is_up = film_host().is_some_and(|host| unsafe { IsWindowVisible(host) }.as_bool());
    unsafe { the_button_follows_the_screen(host_is_up) };
}

/// Whether the pin's host holds an FFmpeg player whose window has not been found yet: the take-up
/// notes the player it is keeping so that the monitor adopts its window as the source the moment it
/// exists (see `host_for_a_hosted_player`).
pub(super) fn adopt_a_stranded_ffplay_into_the_pin_host(pid: u32, path: &Path) -> bool {
    if pid == 0 || video_route(path) != VideoRoute::Ffplay {
        return false;
    }
    note_hosted_player(pid, Some(path));
    true
}

/// Whether this file is shown in a host this app owns: a film pinned in a window of the pin's own,
/// whichever engine plays it. Asked of the route rather than of the player's process because it is
/// the *launch* that decides where the picture is drawn (see `start_video_playback`).
pub(super) fn the_film_is_shown_in_a_host(path: &Path, kind: Option<MediaType>) -> bool {
    // A sound's player is FFmpeg's too, and a sound's route would match the film routes below: the
    // media on screen has to be a film before any host road is asked (see `MediaType::is_audio`).
    pin_is_its_own_window()
        && kind.or_else(current_media_type) == Some(MediaType::Video)
        && matches!(video_route(path), VideoRoute::Mpv | VideoRoute::Ffplay)
}

/// The box the host is placed at, from the band the pin says its picture stands in. The band *is* a
/// screen region (see `PinnedPreview::content`), and a host of the desktop's is placed at it.
fn the_film_host_box(content: ScreenRegion) -> (i32, i32, i32, i32) {
    (
        content.0,
        content.1,
        (content.2 - content.0).max(1),
        (content.3 - content.1).max(1),
    )
}

/// Make the pin's chrome window owned by its film host, or give it back to the Shell, and take the
/// taskbar button and the Alt+Tab entry with it.
///
/// **Who Windows lists and activates is asked of a window once, when it is made, and ownership is the
/// answer that can still be written later.** An owned window stands above its owner, hides with it,
/// and is out of Alt+Tab and the taskbar — which is the whole of what the film road needs from the OS,
/// and the whole of what the pin owes the host for holding its film (see
/// `the_pins_chrome_is_owned`). `WS_EX_APPWINDOW` is what asks for a taskbar
/// button regardless of ownership, so it goes with the road that owns the pin and comes back with the
/// one that does not (see `pin_real_window::real_pin_style`).
///
/// Both are same-process calls into a window of this app's own: there is no cross-process send here,
/// and nothing of this app's is locked across either.
unsafe fn own_the_pin_window(owned: bool) {
    let Some(pin) = pin_real_window::real_pin_window() else {
        return;
    };
    let owner = if owned { film_host() } else { None }.map_or(0, |owner| owner.0 as isize);
    if unsafe { GetWindowLongPtrW(pin, GWLP_HWNDPARENT) } != owner {
        unsafe { SetWindowLongPtrW(pin, GWLP_HWNDPARENT, owner) };
    }

    // **The button is the setting's, and whose window wears it is the host's.** A pair owes no button
    // at all while `Pin Mode → Always On Top` says the pin is over everything already (see
    // `a_button_is_owed`, which is the same rule for the stand-in of a pin standing in this app's own
    // window); with the setting off the host is the window that wears it, because it is the window the
    // Shell lists for the pair, and the chrome is never listed beside it *while the host is on the
    // screen* — the one road that hands it to the chrome is the host going down, and that is asked at
    // the hide rather than here (see `the_button_follows_the_screen`).
    //
    // **And the chrome takes it back the moment no host owns it**, because the pin's window is kept
    // between pins: a film's pin closes and the next pin of a file this app draws the whole way stands
    // in that same window, so a chrome left without a button is a file with no way back from the
    // taskbar. The two calls are one pass on the same thread, so the show the Shell reads for it is
    // the show the next pin makes.
    let app_window = !pin_lifecycle::pin_always_on_top();
    unsafe { pin_real_window::set_the_pin_window_appwindow(pin, !owned && app_window) };
    if let Some(host) = film_host() {
        unsafe { pin_real_window::set_the_pin_window_appwindow(host, app_window) };
    }

    // **The chrome is put in the band the setting names, worn or not.** A window given an owner or
    // given back keeps whatever band it wore before (measured): a bit left on it by the pin before this
    // one, or the band it held while its last host owned it, is a pair that disagrees about where it
    // stands (see `the_pin_window_that_holds_the_band`). The host is put in that band with it on the
    // same pass, and every placement after this asks the setting again (see `band_place` and
    // `place_the_pin_in_the_band`).
    //
    // **It used to be `HWND_NOTOPMOST` for an owned pin, on the assumption that Windows keeps an owned
    // window in its owner's band.** That is not a bit this app can rely on: the pin's own window was
    // then never topmost however the setting read, which is the pin that appears behind Explorer with
    // `Always On Top` on.
    let band = band_place();
    let _ = unsafe {
        SetWindowPos(
            pin,
            band,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        )
    };
}

/// Whether the tick owes the player's window the front: a film a hover is following, or a pin's
/// own hosted film. The pin's half is asked of the host rather than of the hover's bookkeeping,
/// because a hover's `Hide` can empty that bookkeeping while the pin still owns the window (see
/// `a_pin_owns_the_window_a_hide_would_take_down`).
pub(super) fn the_video_window_is_owed_the_front(current_video_path: Option<&Path>) -> bool {
    current_video_path.is_some() || (pinned() && pin_is_its_own_window() && film_host().is_some())
}

/// Whether the player this app is holding is the one it began into the host — the question every
/// road that used to find a top-level mpv window by pid asks instead (see
/// `ensure_video_window_topmost`).
///
/// It says the pairing and not that the film is drawing: a player begun into the host has not put
/// a child of its own there for a moment or more, and that wait is `the_host_is_drawing`.
pub(super) fn the_pinned_player_was_begun_in_the_host() -> bool {
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    pid != 0 && film_host().is_some() && hosted_player_pid() == Some(pid)
}

/// How often an unchanged `the_host_is_drawing` answer is written down: the tick asks it hundreds of
/// times a second, and on-change plus this heartbeat is what makes the trace readable (see
/// `trace_the_hosted_drawing`).
const HOSTED_DRAWING_TRACE_EVERY: Duration = Duration::from_secs(1);

/// One `the_host_is_drawing` answer as the trace last wrote it: the player whose source the host is
/// drawing, the host (zero for none), the answer, and when it was written.
type HostedDrawingTrace = (u32, isize, bool, Instant);

/// The last answer `the_host_is_drawing` wrote down (see `trace_the_hosted_drawing`).
static LAST_HOSTED_DRAWING: Lazy<Mutex<Option<HostedDrawingTrace>>> =
    Lazy::new(|| Mutex::new(None));

/// Write the answer of one look at the hosted film's window, on change and at most once a second.
///
/// **The line is asked for by every tick, several times, and used to write some 3 MB a run** — which
/// is a trace nothing can be read out of. A state transition is the interesting part and is never
/// missed; an answer that has not moved is repeated only as the heartbeat that says the tick is
/// still asking it.
fn trace_the_hosted_drawing(pid: u32, host: Option<HWND>, drawing: bool) {
    let Ok(mut last) = LAST_HOSTED_DRAWING.lock() else {
        return;
    };
    let now = Instant::now();
    let host_number = host.map_or(0, |host| host.0 as isize);
    let answer = (pid, host_number, drawing, now);
    if !the_drawing_trace_is_due(*last, answer, now) {
        return;
    }
    *last = Some(answer);
    match host {
        Some(host) => note_pin_trace!(
            "hosted drawing pid {pid} host {host:?} drawing {}",
            drawing as u8
        ),
        None => note_pin_trace!("hosted drawing pid {pid} host none false"),
    }
}

/// Whether an answer is due a line: a change of player, host or answer, and otherwise no more than
/// the heartbeat `HOSTED_DRAWING_TRACE_EVERY` after the last one (see `trace_the_hosted_drawing`).
fn the_drawing_trace_is_due(
    last: Option<HostedDrawingTrace>,
    answer: HostedDrawingTrace,
    now: Instant,
) -> bool {
    last.is_none_or(|(pid, host, drawing, at)| {
        (pid, host, drawing) != (answer.0, answer.1, answer.2)
            || now.duration_since(at) >= HOSTED_DRAWING_TRACE_EVERY
    })
}

/// **Whether the pin's host has a picture in it at all**, whichever player began it: the registration
/// standing in that host, which is the frame on screen (see
/// `pin_film_thumbnail::the_standing_source_process`).
///
/// It is the question every road that would show, raise or order the host asks, and it is asked of the
/// registration rather than of the newest player because of the whole of a relaunch's wait: the
/// replacement is recorded before it has made a window, while the player before it is **covered** and is
/// what the host is drawing (see `pin_playback::relaunch_pinned_player`). Asked of the newest pid
/// instead, a seek's instant is a pin declared empty — the z-order re-assert skipped, the band treated
/// as a hole — with a still standing in it. A host with no registration in it yet is the player's
/// own window, which is all there is while the source is being adopted and the Shell has not been handed
/// it (see `pin_film_thumbnail::adopt_the_source`). **A still this app painted is the one picture a host
/// can have with no registration in it**, and it is not this question's: the host stands where it is and
/// the tick's show has nothing to add to it (see
/// `pin_film_thumbnail::stand_the_outgoing_picture_from_its_dc`).
pub(super) fn the_host_is_drawing() -> bool {
    let Some(host) = film_host() else {
        trace_the_hosted_drawing(0, None, false);
        return false;
    };
    if let Some(standing) = pin_film_thumbnail::the_standing_source_process(host) {
        trace_the_hosted_drawing(standing, Some(host), true);
        return true;
    }
    // A host with no registration in it yet is the pass between a source being adopted and the
    // Shell being handed it, and the player's own window is all there is to show for it.
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    let drawing = the_hosted_player_is_drawing(pid);
    trace_the_hosted_drawing(pid, Some(host), drawing);
    drawing
}

/// Whether **this** player is up as the pin's picture: the source the Shell draws into the host is the
/// window this player made, adopted and still live (see `pin_film_thumbnail::the_source_is_up`).
///
/// **It is asked of the process rather than of the host alone**, because the host outlives the player
/// it holds: a file step begins a new player for the same host, and the old source would answer "a
/// film is drawing" for a film about to be killed. The source records the player it was adopted for,
/// so the pid and the window are one question. What it is *not* is "has the host a picture": a
/// replacement that has made no window yet answers no here while the frozen picture it replaced is
/// still drawing (see `the_host_is_drawing`).
pub(super) fn the_hosted_player_is_drawing(pid: u32) -> bool {
    film_host().is_some()
        && pid != 0
        && hosted_player_pid() == Some(pid)
        && pin_film_thumbnail::the_source_is_up(pid)
}

/// Whether this player has a window on screen: a source the Shell is drawing through a host of this
/// app's, or a window of the desktop's of its own. It is one answer because the settlement roads
/// have one question — has this player arrived — and a film drawn through a host publishes no window
/// of the desktop's to be found by (see `video_window_for`, `settle_video_retirement` and
/// `PinSwapHold::settle`).
pub(super) fn the_players_window_is_up(pid: u32) -> bool {
    video_window_for(pid).is_some() || the_hosted_player_is_drawing(pid)
}

/// Make the host where the pin has none, or put the one it has at the band it stands at now.
///
/// Asked of the pin rather than of the caller's box, because the two roads that reach it — a show
/// and a box change — have the pin's state written before either of them asks, and the band the
/// host is placed at is the same band the window was painted around.
///
/// **A host made here is not shown unless the player is already drawing in it.** It is a black
/// window until the player this app begins has put a child of its own in it, and a black rectangle
/// over the band is the film the hover was showing, blanked. A host a restore takes up is drawing
/// already and is shown here; one a fresh launch made is shown by the tick's own road once that
/// source is there (see `ensure_video_window_topmost`).
///
/// `adopted` is the hover's host on the road that takes an already-running film up: its player is
/// the one drawing the hover, so the window is the pin's rather than replaced (see
/// `promote_the_hover_film_host`). Every other road passes `None` and a fresh host is made where
/// there is none.
pub(super) unsafe fn take_up_the_pin_film_host(adopted: Option<HWND>) {
    let Some(path) = pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.path.clone())) else {
        take_down_the_pin_film_host();
        return;
    };

    if !the_film_is_shown_in_a_host(&path, None) {
        take_down_the_pin_film_host();
        return;
    }

    let Some(content) = pinned_content() else {
        return;
    };

    if let Some(host) = adopted.filter(|host| unsafe { IsWindow(*host) }.as_bool()) {
        // **The hover's host is already the window this road wants**, so it is the pin's host rather
        // than a window moved into one: the player keeps drawing into the same source and not a frame
        // of the film leaves the screen (see `promote_the_hover_film_host`).
        unsafe { promote_the_hover_film_host(host, content) };
        return;
    }

    let host = match film_host() {
        Some(host) => host,
        None => {
            let Some(host) = create() else {
                return;
            };
            note_pin_trace!(
                "host created {:?} pid {} source {:?} client {:?}",
                host,
                VIDEO_PID.load(Ordering::SeqCst),
                pin_film_thumbnail::the_source_window(),
                a_client_box(host),
            );
            host
        }
    };

    // **The pin's chrome is the host's to own on this road** (see `own_the_pin_window`). Written here
    // because this is the road that knows both windows.
    unsafe { own_the_pin_window(true) };

    let (x, y, width, height) = the_film_host_box(content);
    let _ = SetWindowPos(
        host,
        HWND::default(),
        x,
        y,
        width,
        height,
        SWP_NOZORDER | SWP_NOACTIVATE,
    );
    // A picture the Shell is already drawing into this host goes with the box it was placed at, even
    // where the host was made and the film's window found in the same pass (see `pin_film_thumbnail`).
    // **The source itself is never asked for a box again**, so this is a destination rectangle and
    // nothing else (see `pin_film_thumbnail::the_band_was_placed`).
    pin_film_thumbnail::the_band_was_placed(host);
    pin_chrome_strips::place_the_chrome_strips();

    // **A hosted player has no window of the desktop's to be found by pid**, so any handle still
    // published for an older player is not this film's picture and must not be read as one (see
    // `the_players_window_is_up`).
    if the_pinned_player_was_begun_in_the_host() {
        VIDEO_HWND.store(0, Ordering::SeqCst);
    }

    // Only a host the player is already drawing in is shown: a fresh one is a black rectangle over
    // the band, and what the band showed a moment ago is the film the hover was playing (see the
    // module note on `create`).
    claim_the_band_for_the_host();
    if the_hosted_player_is_drawing(VIDEO_PID.load(Ordering::SeqCst)) {
        // **A host a restore takes up is drawing already, and showing it is an appearance like any
        // other** — the one the swap's hide left off the band, put back on it with the band write the
        // show is owed (see `the_host_has_appeared`).
        let was_up = unsafe { IsWindowVisible(host) }.as_bool();
        // And the pair's button goes with the appearance, written before the show it belongs to (see
        // `the_button_follows_the_screen`).
        unsafe { the_button_follows_the_screen(true) };
        let _ = ShowWindow(host, SW_SHOWNOACTIVATE);
        if !was_up {
            unsafe { the_host_has_appeared(host) };
        }
        pin_chrome_strips::sync_the_chrome_strips();
    }
}

/// Take a hover's already-running host up as the pin's own: **the same window, the same player, the
/// same source and registration, and not one frame of the film off the screen.**
///
/// **Why it can be done at all.** A hover's film is given a host only so that it can be handed to a
/// pin without a second launch, and that host is already the one shape a pin's host has to be: a
/// top-level, non-layered window of this app's — `WS_POPUP | WS_CLIPCHILDREN`, the film a thumbnail
/// of the player's own window (see `hover_film_host`). What a hover's host is *for a hover* is the
/// rest of its style — the topmost band nothing may cover, transparent to the press, out of the
/// Shell's lists — and the pin's is none of that (see `create`).
///
/// **Why it is not a re-parent.** The host is a window of the desktop's so that the Shell composes it
/// and the thumbnail of the film in it: moving it into the pin's layered window would take that away,
/// so the host takes the pin's chrome as its own window instead (see `own_the_pin_window`).
///
/// **Nothing is begun, frozen or read back**, so there is no pause between the hover and the pin and
/// no second decode of the file: what changes is a style, an owner, a band and a box, and what does
/// not is the player inside it.
unsafe fn promote_the_hover_film_host(host: HWND, content: ScreenRegion) {
    let pid = hover_film_host::hover_film_host_player().unwrap_or(0);

    // **A hover's player the pin has just taken is the resident, where a resident was serving it.**
    // The process is the pin's from here — the same window and the same player, not a copy — so its
    // record as this app's hover player goes, and a fresh one is begun once this pin has closed (see
    // `mpv_resident::a_pin_took_it_over`).
    mpv_resident::a_pin_took_it_over();

    // The pin's host from here, written before the roads that ask the host questions through the
    // pin's own lock can run: the player the hover began is the pin's own, and the frame it is
    // drawing is already on the screen — the hover had proved it.
    PIN_FILM_HOST.store(host.0 as isize, Ordering::SeqCst);
    PIN_FILM_HOST_PLAYER.store(pid, Ordering::SeqCst);
    PRESENTED_PID.store(pid, Ordering::SeqCst);
    VIDEO_HWND.store(0, Ordering::SeqCst);
    // The promoted window is the pin's band from here, picture and all (see
    // `the_band_is_a_standing_host`).
    claim_the_band_for_the_host();
    unsafe { pin_taskbar::set_pin_taskbar_title(host) };

    // **The two style words the Shell reads the pair's button for come on with the pin's bits**, on
    // a host that was made from the *hover's* class and never went through `create`: a promoted host
    // is the window the pair's button belongs to from here, and it is given them, and the app's
    // icon with them, before the show this road pays for below — that show being the one a button is
    // drawn from (see `the_button_shells_reads_its_gestures_from`).
    unsafe { the_button_shells_reads_its_gestures_from(host) };

    // **The hover's band and the press that falls through it go, and the pin's bits come on**: a host
    // the Shell lists wears neither `WS_EX_TOOLWINDOW` nor `WS_EX_NOACTIVATE` (see `create`).
    //
    // **The app-window bit is not written here.** `WS_EX_APPWINDOW` is what asks the Shell for a
    // button, and the Shell reads it as a window is shown: a promotion is a host that is already on the
    // screen, so the bit goes through the one road that puts a window down and up again for it —
    // `pin_real_window::set_the_pin_window_appwindow`, which is the whole of why the pair's button was
    // intermittent (see `the_button_follows_the_screen`).
    let ex = unsafe { GetWindowLongPtrW(host, GWL_EXSTYLE) };
    let ex = ex
        & !(WS_EX_TOPMOST.0 as isize)
        & !(WS_EX_TRANSPARENT.0 as isize)
        & !(WS_EX_TOOLWINDOW.0 as isize)
        & !(WS_EX_NOACTIVATE.0 as isize)
        & !(WS_EX_APPWINDOW.0 as isize);
    let _ = unsafe { SetWindowLongPtrW(host, GWL_EXSTYLE, ex) };
    // A style written behind a window's back is not one it presents with until it is asked to read it,
    // and the band it was in goes with the same call.
    let _ = unsafe {
        SetWindowPos(
            host,
            HWND_NOTOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED,
        )
    };
    // **The hover's host was made disabled, and the pin's host is an ordinary window**: a disabled one
    // is skipped by `WindowFromPoint`, but it is also a window the Shell can neither activate nor put
    // down, and this is the window the pair's button belongs to from here (see `create` and the module
    // note). The presses are the hit test's on this road too (see `answer_a_host`).
    let _ = unsafe { EnableWindow(host, true) };
    // The pair's button from here is the setting's, read with the show this road pays for (see
    // `the_button_follows_the_screen`).
    unsafe { the_button_follows_the_screen(true) };

    // The chrome is the host's own window from here, and the pair is put in the band the setting
    // names in one call — the host taking it and the chrome following it (see
    // `place_the_pin_in_the_band`). The host keeps the picture: it was the hover's window on the
    // screen and it is the pin's window on the screen, moved to the band the pin was painted around.
    unsafe { own_the_pin_window(true) };
    unsafe { place_the_pin_film_host(content) };
    let pin = pin_real_window::real_pin_window().unwrap_or(host);
    unsafe { place_the_pin_in_the_band(pin, band_place()) };

    // The hover's handle is given up rather than left to destroy the pin's host (see
    // `forget_the_hover_film_host_that_was_handed_over`), and the strips are put back over the
    // window that is the film's now.
    hover_film_host::forget_the_hover_film_host_that_was_handed_over(host);
    unsafe { pin_chrome_strips::raise_the_chrome_strips_over_the_host() };

    note_pin_trace!(
        "host promoted {:?} pid {} source {:?} client {:?} band {:?}",
        host,
        pid,
        pin_film_thumbnail::the_source_window(),
        a_client_box(host),
        band_place(),
    );
}

/// Put the host at the pin's band, without showing it: the band-change road, where the window is
/// either already up (see `ensure_pinned_sibling_box`).
pub(super) unsafe fn place_the_pin_film_host(content: ScreenRegion) {
    let Some(host) = film_host() else {
        return;
    };

    let (x, y, width, height) = the_film_host_box(content);
    // **A host has nothing of this app's to carry across a step**, so its bits are not copied: the
    // picture is the Shell's thumbnail or FFmpeg's child, and the only pixels this app paints in the
    // host are its whole client erased black behind that picture — a fresh surface every step is what
    // keeps a bar of an earlier box from standing black over the picture the box now draws (see
    // `pin_film_thumbnail::settle_the_registration`).
    let flags = SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS;
    // The temporary `RHP_RESIZE_TRACE` diagnostic's own clock for this placement, taken only where a
    // resize step is being measured (see `pin_resize_probe`).
    let started = pin_resize_probe::a_step_started();
    let _ = SetWindowPos(host, HWND::default(), x, y, width, height, flags);
    // **What a resize step costs on the thumbnail road is this one call**: the destination is the
    // host's new client area, and the film's window is not touched at all (see `pin_film_thumbnail`).
    pin_film_thumbnail::the_band_was_placed(host);
    pin_resize_probe::note_the_host_placement(started);
    pin_chrome_strips::place_the_chrome_strips();
}

/// Put the pair where it belongs the moment the window that carries the film appears: the band the
/// setting names, asked again of both windows of it.
///
/// **Every band write of a pin's arrival lands while the host is still hidden** — the take-up's
/// placement, the appearance raise `take_the_band_and_give_it_back`, and the placement
/// `show_pinned_window` makes before it shows the chrome. A fresh host is hidden until its player
/// draws (see `create` and `video_playback::ensure_video_window_topmost`), and a swap hides the host
/// for the wait of the next player (see `take_the_film_off_the_band`), so with
/// `Pin Mode → Always On Top` off the pair came up behind whatever was in front of the listing and
/// stayed there until the user picked the file again: a second pick only "worked" because by then
/// the host was on the screen and the same call landed on live windows (see
/// `place_the_pin_in_the_band`).
///
/// **A band asked of a window that is already in the ordinary one is no raise, and the show of a
/// hidden window is no raise either** — both measured on this desktop: `HWND_NOTOPMOST` answers
/// success and leaves the window exactly where it was, and a window shown back into the place it went
/// down in comes back behind everything opened over it in the meantime. What does raise a window of a
/// process without the foreground is the topmost band taken for one call and given straight back,
/// which is the road the pin's arrival already takes (see `take_the_band_and_give_it_back`) and the
/// reason `HWND_TOP` is not used anywhere here.
///
/// **A band that is held is not this road's to give back.** Where a press's grace or a step's wait
/// names the topmost band, `band_place` answers it and that is the band the pair is left in: neither
/// hold is cleared here, so the band is given back by the clock that took it and not by an
/// appearance.
unsafe fn the_host_has_appeared(host: HWND) {
    let pin = pinned_window();
    let band = band_place();
    if band != HWND_TOPMOST {
        note_pin_trace!("host appeared {:?}: the band taken to bring the pair forward", host);
        unsafe { place_the_pin_in_the_band(pin, HWND_TOPMOST) };
    }
    unsafe { place_the_pin_in_the_band(pin, band) };
}

/// Hide or show the host the film is drawn in.
///
/// **Nothing but the film leaving the pin hides it** (see `take_the_film_off_the_band`), and a hide
/// with a pin standing over a source the host is drawing is written out as the defect it would be:
/// the picture is the Shell's thumbnail of the player's window, and there is no frame of this app's
/// to stand in for it (see `pin_film_thumbnail`). The chrome stays where the chrome state puts it:
/// the strips are the pin's window's rather than the host's, so they are not hidden with it.
///
/// The last show this host was asked for, and the host it was asked for: the tick asks on every
/// pass, so the trace's `host visible` line is written where the answer changes rather than on every
/// ask (938 of a 114 s run's lines were this family and `topmost shown`).
static LAST_HOST_SHOW: Lazy<Mutex<Option<(isize, bool)>>> = Lazy::new(|| Mutex::new(None));

/// Whether a show is one the trace has not written yet: the answer the last pass wrote is the
/// polling line a normal session's trace does not need (see `trace_the_host_show`).
fn the_show_is_new(last: Option<(isize, bool)>, host: isize, visible: bool) -> bool {
    last != Some((host, visible))
}

/// Write the show this host was asked for, where it is not the one the last pass asked for.
fn trace_the_host_show(host: HWND, visible: bool) {
    let Ok(mut last) = LAST_HOST_SHOW.lock() else {
        return;
    };
    if !the_show_is_new(*last, host.0 as isize, visible) {
        return;
    }
    *last = Some((host.0 as isize, visible));
    note_pin_trace!("host visible {} {:?}", visible as u8, host);
}

/// How many times this run has asked the host to hide, for the tests that hold a relaunch to never
/// taking the picture off the screen (see `set_the_pin_film_host_visible`).
#[cfg(test)]
pub(super) static HOST_HIDES_ASKED: AtomicU32 = AtomicU32::new(0);

/// The hides this run has asked for.
#[cfg(test)]
pub(super) fn host_hides_asked() -> u32 {
    HOST_HIDES_ASKED.load(Ordering::SeqCst)
}

pub(super) unsafe fn set_the_pin_film_host_visible(visible: bool) {
    #[cfg(test)]
    if !visible {
        HOST_HIDES_ASKED.fetch_add(1, Ordering::SeqCst);
    }
    if let Some(host) = film_host() {
        // **A show is the band standing over this host again**, and a hide is not: the band's
        // pixels are a function of that, so each transition is painted once (see
        // `the_band_is_a_standing_host`).
        let was_up = unsafe { IsWindowVisible(host) }.as_bool();
        // **A host the Shell has minimized is the Shell's to bring back.** Showing a minimized
        // window restores it, so a tick that showed the host it is watching would undo the user's own
        // minimize of it every tick — and the button that minimized it is the host's own on this road
        // (see `own_the_pin_window`).
        if visible && unsafe { IsIconic(host) }.as_bool() {
            return;
        }
        // **A picture hidden under a pin is the blank this road exists not to make.** The one road
        // that may hide the host is the film being taken off the pin, which lets the registration go
        // before it asks for the hide; anything else is a relaunch's step taken over a standing
        // source, and a run says so here rather than in what the user sees.
        if !visible
            && pinned()
            && !pin_is_collapsed()
            && pin_film_thumbnail::the_source_is_standing(hosted_player_pid().unwrap_or(0))
        {
            note_pin_trace!(
                "film host hidden while a pin lives {:?}: the picture was the standing source",
                host,
            );
        }
        if visible {
            // **The button is the appearing host's from the show that makes it the pair's listed
            // window**, written before the show so that the Shell reads it with it (see
            // `the_button_follows_the_screen`).
            unsafe { the_button_follows_the_screen(true) };
        }
        let _ = ShowWindow(host, if visible { SW_SHOWNOACTIVATE } else { SW_HIDE });
        // The line is written after the show rather than before it, so that a run's trace reads in the
        // order the desktop was asked (see `trace_the_host_show`).
        trace_the_host_show(host, visible);
        if visible {
            claim_the_band_for_the_host();
            // **The band is asked again of a host that has just appeared**, because every band write
            // of an arrival landed while this window was still hidden (see `the_host_has_appeared`).
            if !was_up {
                unsafe { the_host_has_appeared(host) };
            }
        } else {
            // **And a hidden host's button goes back to the chrome**, which is the window the user is
            // left with once this one is down (see `the_button_follows_the_screen`).
            unsafe { the_button_follows_the_screen(false) };
        }
        // **The film follows the host off the screen and back**, which is the one road a source is
        // hidden on: the Shell stops drawing the thumbnail of a hidden window, and there is no host
        // to draw it in while the host is down (see `pin_film_thumbnail`).
        pin_film_thumbnail::set_the_source_visible(visible);
        // A host that was down and is up is a picture the pin's band has never been painted over
        // (see `the_band_owes_a_repaint`).
        if visible && !was_up {
            the_band_owes_a_repaint();
        }
    }
    // The strips carry the chrome on every road, whatever the host is doing: the sync answers
    // what the chrome state asks for over the picture or over the cover (see `pin_chrome_strips`).
    pin_chrome_strips::sync_the_chrome_strips();
}

/// Take the film off the band, because the pin is about to show the host no film: the registration
/// goes first, so the hide that follows is a window going down rather than a picture going black.
///
/// It is one call rather than two at each road that ends a film, because the order is the whole of
/// it — a host hidden with the source the Shell draws still registered in it is the blank frame this
/// road is a film road to avoid (see `set_the_pin_film_host_visible` and
/// `pin_film_thumbnail::forget_the_thumbnail`).
pub(super) unsafe fn take_the_film_off_the_band() {
    pin_film_thumbnail::forget_the_thumbnail_because("the film is leaving the pin");
    // A cover standing in the host is that film's, and the host is about to show another one (see
    // `put_the_cover_away`).
    put_the_cover_away();
    unsafe { set_the_pin_film_host_visible(false) };
    // **And the band is nobody's claim any more**, which is the half of this that has to be written:
    // a host with its registration gone and its cover away has nothing to show, and a band still
    // claiming that pixel would be the opaque black rectangle that stands where the film was and
    // never goes (see `the_band_is_a_standing_host`).
    release_the_band();
    the_band_owes_a_repaint();
}

/// Take the picture the host's own taskbar button is to hold, while the host is still on the screen to
/// take it of: the last frame of the film, which is what the hover of a put-away pin shows.
///
/// **The host is the picture, and not the source the Shell draws in it.** A film in a host is a DWM
/// thumbnail composed *over* the host's client — it is no part of the host's own rows — so a still the
/// Shell took of those rows is the black the host is erased to (see
/// `paint_the_host_behind_the_film`, `pin_taskbar_mirror`). What is taken instead is the host's client
/// as the desktop is showing it, which is the film and the chrome strips drawn over it, and it is that
/// picture the Shell is handed when it asks for the still of a button that is down.
///
/// **And the button is told to use it for as long as the pin is away**, because the host has no
/// composition of its own for the Shell to freeze: a forced still is what keeps the hover of a
/// minimized host the pin rather than the icon the Shell falls back on (see
/// `the_buttons_still_is_the_apps_own`).
///
/// **The picture is taken while the host is still on the screen, and the button is told to hold it
/// whether or not this row could take one.** The row that puts a pin away into the taskbar takes it
/// here; the row that asks for the taskbar from the *bubble* has had no window of the pin's on the
/// screen since the bubble's own collapse, and the picture that collapse took — the last frame the
/// user saw — is what the button is handed (see `take_the_pins_picture`, `collapse_pin_to`). A row
/// that forced nothing here left the hover of such a pin blank, because a window that is down has no
/// composition of its own for the Shell to freeze.
///
/// Safety: `host` is the pin's own film host, asked by the collapse before the pin leaves the screen.
pub(super) unsafe fn keep_the_pins_picture_for_the_taskbar(host: HWND) {
    unsafe { take_the_pins_picture(host) };
    pin_taskbar_mirror::the_stills_are_out_of_date(host);
    unsafe { pin_taskbar::the_buttons_still_is_the_apps_own(host, true) };
}

/// Take the host's picture, to be handed to the Shell as the still of its button: the film and the
/// chrome as the desktop is showing them, which is the last frame of a pin that is put away.
///
/// **It is taken at every collapse and not only at a put-away into the taskbar**, because the last
/// moment the host is on the screen is the collapse — a pin moved on to the taskbar from the bubble
/// wears the host's button too, and every window of the pin's has been off the screen since (see
/// `keep_the_pins_picture_for_the_taskbar`).
///
/// A window that is not drawn where it stands is no picture at all, and what the button already holds
/// is the picture to keep (see `pin_taskbar_mirror::take_the_picture_of`).
///
/// Safety: `host` is the pin's own film host, asked by the collapse before the pin leaves the screen.
pub(super) unsafe fn take_the_pins_picture(host: HWND) {
    if the_pin_window_is_down(host) {
        return;
    }

    pin_taskbar_mirror::take_the_picture_of(host);
}

/// Put the film back on the band when the pin comes back up: the host shown again with the source the
/// Shell draws through it, and the chrome strips synced over it.
///
/// **Nothing else brings the pair back on this road.** A pin put away as a window takes its film off the
/// screen with it — the host is hidden and the source hidden with it (see `hide_pinned_siblings`,
/// `pin_film_thumbnail::set_the_source_visible`) — and the window the Shell restores is the host, which
/// no road of this app's shows: `place_the_pin_film_host` moves a window and the tick holds its own
/// show back for a pin whose chrome is owned (see `the_pin_owes_a_show`). What a restore that showed
/// neither left was a black band where the film had been.
///
/// It is asked where the pin's *state* says it is up — the restore's own road — rather than on the
/// tick, so that the film comes back with the pin in the same pass rather than a frame later, and so
/// that the still taken for the button goes with the pin's being on the screen again (see
/// `keep_the_pins_picture_for_the_taskbar`).
///
/// Safety: asked by the pin's own restore, with the pin up and its host still standing.
pub(super) unsafe fn bring_the_film_back() {
    let Some(host) = film_host() else {
        return;
    };

    // **A pin the user asked for is up with its film, whatever the Shell has not finished doing.** The
    // window the Shell restores is the host, and the size that says it is back can arrive before the
    // iconic bit has gone — the gate in `set_the_pin_film_host_visible` holds a *tick* back from
    // showing a host the user minimized, and it must not hold this road back: the pin's state is up
    // here because the Shell brought the window back, and a host left iconic is the black band with
    // the pin's own window standing over it, recovered only by whatever pass happens to show it later.
    // So the host is asked up rather than waited for, without activating it — the pin's own window is
    // the one the restore gave the front to.
    if unsafe { IsIconic(host) }.as_bool() {
        let _ = unsafe { ShowWindow(host, SW_SHOWNOACTIVATE) };
    }

    unsafe { set_the_pin_film_host_visible(true) };

    // **A pin that is up is drawn live again**: the still that stood for it while it was away is the
    // pin's no longer, and a forced one would freeze the hover of a film the user is watching.
    unsafe { pin_taskbar::the_buttons_still_is_the_apps_own(host, false) };
    pin_taskbar_mirror::forget_the_pin_picture();
}

/// Destroy the host and forget the player it held: the pin it was showing a film for is ending, and
/// a window left standing would keep the desktop's band it owns. The strips stay: they are the pin's
/// window's, and the next take-up places them again (see `pin_chrome_strips`).
pub(super) unsafe fn take_down_the_pin_film_host() {
    // **A band a step held goes back before the host that owns it does.** The hold is a placement
    // of the pair, and one cleared after the host is gone would have the pin placed topmost by its
    // own teardown instead (see `PIN_STEP_BAND_HELD` and `own_the_pin_window`).
    unsafe { pin_lifecycle::give_the_step_band_back() };
    // **The picture the Shell is drawing goes before the window it is drawn in.** The film's window is
    // the player's own, so a registration left behind a host that is gone is the Shell's resource
    // leaking, and the source is taken off the screen with the host it was seen in (see
    // `pin_film_thumbnail`).
    pin_film_thumbnail::set_the_source_visible(false);
    // A pin's own end is the last settle of the held frames it counted, written before the record
    // that counted them goes (see `pin_film_thumbnail::settle_the_held_frames`).
    pin_film_thumbnail::settle_the_held_frames();
    pin_film_thumbnail::forget_the_thumbnail_because("the pin is ending");
    // The band has no host of its own any more (see `the_band_is_a_standing_host`).
    release_the_band();
    // The cover went with the picture: there is no host left for it to be re-cropped into (see
    // `put_the_cover_away`).
    put_the_cover_away();
    // And the still taken for the host's own button goes with it: a picture of a window that is gone is
    // no pin's hover (see `keep_the_pins_picture_for_the_taskbar`).
    pin_taskbar_mirror::forget_the_pin_picture();
    // **The pin's chrome is given back to the Shell before the host that owns it goes.** Destroying
    // a window destroys the windows it owns, and the pin's window is kept between pins — so a
    // teardown that destroyed the owner first would take the pin's own window with it (see
    // `own_the_pin_window`).
    unsafe { own_the_pin_window(false) };
    // **The player is ended before the window it draws in is destroyed.** A player whose surface was
    // destroyed under it is a kill the player did not get to answer; the record below is what the
    // ordinary kill road retries to confirmation (see `terminate_player_pid`).
    if let Some(pid) = hosted_player_pid() {
        terminate_player_pid(pid);
    }
    if let Some(host) = film_host() {
        let _ = DestroyWindow(host);
    }
    PIN_FILM_HOST.store(0, Ordering::SeqCst);
    PIN_FILM_HOST_PLAYER.store(0, Ordering::SeqCst);
    PRESENTED_PID.store(0, Ordering::SeqCst);
}

/// The window the pin's film is drawn from, where the Shell is drawing it: the player's own window
/// the host holds a thumbnail of, which is the third window a press can land in and the one the
/// keyboard hook has to know about (see `key_input::pin_owns_caret`).
///
/// `None` where no host is up, and the player's window of the desktop's is the hook's own to know.
pub(super) fn the_films_own_window() -> Option<HWND> {
    film_host()
        .is_some()
        .then(pin_film_thumbnail::the_source_window)
        .flatten()
}

/// A client box as a trace reads it, or nothing where the window has already gone.
fn a_client_box(hwnd: HWND) -> Option<(i32, i32)> {
    let mut rect = RECT::default();
    (unsafe { GetClientRect(hwnd, &mut rect) }.is_ok()).then_some((rect.right, rect.bottom))
}

unsafe fn create() -> Option<HWND> {
    FILM_HOST_CLASS_REGISTERED.call_once(|| {
        let hinstance = GetModuleHandleW(None).expect("the module's own handle");
        // **The icon this class's windows are given a taskbar button with.** A class takes an icon
        // once, as it registers, and a window whose class has none is drawn with the default exe
        // rectangle rather than with the app's own face (see `app_icon`). The host a hover is
        // promoted into is made from the *hover's* class and is given the same icon on the window
        // instead (see `the_button_shells_reads_its_gestures_from`).
        let icon = unsafe { crate::shell::app_icon::app_icon() };
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(the_default_answer),
            hInstance: hinstance.into(),
            hIcon: icon,
            hIconSm: icon,
            // **No background brush, because the host paints its own erase** (see
            // `paint_the_host_behind_the_film`): the class's brush is what erased a grown band
            // black, and the one erase road is the window procedure's.
            hbrBackground: HBRUSH(GetStockObject(NULL_BRUSH).0),
            lpszClassName: FILM_HOST_CLASS,
            ..Default::default()
        };
        RegisterClassExW(&class);
    });

    let hinstance = GetModuleHandleW(None).ok()?;
    // **Made without `WS_VISIBLE`, and that is what keeps the band the film's until the Shell has
    // composed the picture.** A host put on the screen empty is a black rectangle where the hover's
    // picture was, for as long as the player takes to open the file. What shows it is the take-up of a
    // host the player is already drawing in, and the tick's own road (see
    // `take_up_the_pin_film_host`, `ensure_video_window_topmost`).
    //
    // **`WS_CLIPCHILDREN`** so that the host's own erase is clipped to the area outside the chrome
    // strips when a pin's window is not the one painting (see `pin_real_window::real_pin_style`).
    //
    // **`WS_SYSMENU | WS_MINIMIZEBOX` are the Shell's read of the taskbar button this window wears,
    // and nothing drawn of them**: there is no caption here to draw a system menu on, so what the
    // bits are for is the button — a window with no system menu is given one with no icon on it, and
    // the button's own minimize and restore are the `SC_MINIMIZE` and `SIZE_MINIMIZED` this app
    // answers (see `the_button_shells_reads_its_gestures_from`, and
    // `pin_taskbar::stand_in_style`, which is the same two words for the same reason).
    //
    // **A pin that is Always On Top owes no taskbar button, and the host is where a pair's button
    // lives** — so the bit is asked of the setting here rather than always added (see
    // `own_the_pin_window`, which is the road that keeps it in step afterwards).
    let ex_style = if pin_lifecycle::pin_always_on_top() {
        WINDOW_EX_STYLE(0)
    } else {
        WS_EX_APPWINDOW
    };
    let host = CreateWindowExW(
        ex_style,
        FILM_HOST_CLASS,
        w!("film host"),
        WS_POPUP | WS_SYSMENU | WS_MINIMIZEBOX | WS_CLIPCHILDREN,
        0,
        0,
        1,
        1,
        // A window of the desktop's has no parent — it is the one that has an *owner*, and the owner
        // is the pin's chrome (see `own_the_pin_window`).
        HWND::default(),
        None,
        hinstance,
        None,
    )
    .ok()?;

    // The window is made *with* the two words the Shell reads its button for rather than given them
    // after its first show, and the helper is asked here for the icon and so that the two roads to a
    // pin's host are one (see `the_button_shells_reads_its_gestures_from`).
    unsafe { the_button_shells_reads_its_gestures_from(host) };

    // **An ordinary window of the desktop's, enabled and activatable.** Its presses are kept for the
    // pin by its own hit test rather than by `WS_DISABLED`, so that the Shell can activate and put
    // down the window the pair's taskbar button belongs to (see `answer_a_host` and the module note).
    // Disabling it here is what made a click on that button deliver nothing to it at all.
    PIN_FILM_HOST.store(host.0 as isize, Ordering::SeqCst);
    // **A log has a starting point**: what the host was made as, which every line about it is read
    // against — whether it wears the Shell's own button (the `WS_EX_APPWINDOW` this app asks for
    // with `Always On Top` off), which band the pair stands in, and who owns whom. The owner is
    // written as it stands here: the host is made with none, and the pin's chrome is given to it by
    // `own_the_pin_window` afterwards (see it for the road that writes the relation).
    note_pin_trace!(
        "film host created appwindow={} always-on-top={} owner={} pin-owner={}",
        ((GetWindowLongPtrW(host, GWL_EXSTYLE) as u32 & WS_EX_APPWINDOW.0) != 0) as u8,
        pin_lifecycle::pin_always_on_top() as u8,
        window_name(HWND(GetWindowLongPtrW(host, GWLP_HWNDPARENT) as *mut _)),
        window_name(HWND(
            GetWindowLongPtrW(pinned_window(), GWLP_HWNDPARENT) as *mut _
        )),
    );
    // **The host is the window the Shell lists**, so the button the pin used to carry names the file
    // the pin is of (see `pin_taskbar::set_pin_taskbar_title`), and it is made supplying the still its
    // button's hover is drawn of — this window's own rows are the black it erases to, not the film the
    // Shell composes over them (see `keep_the_pins_picture_for_the_taskbar`).
    unsafe { pin_taskbar::set_pin_taskbar_title(host) };
    unsafe { pin_taskbar::supply_the_stills_of_this_window(host) };
    // And the chrome is put back over it: making the host puts it on top of the strips, which are
    // the pin's window's and were standing before it, so the order is re-asserted here (see
    // `pin_chrome_strips`).
    unsafe { pin_chrome_strips::raise_the_chrome_strips_over_the_host() };
    Some(host)
}

/// **The two style words the Shell reads the pair's taskbar button for, and the icon it is drawn
/// with**, asked of the one window that wears that button: the pin's film host, whichever road made
/// it.
///
/// **Why the words.** A window with no system menu is given a taskbar button with no icon on it, and
/// the minimize and restore on a button's own menu are what the Shell sends a window it is able to
/// put down — the same two words the stand-in and the pin's own window are made with (see
/// `pin_taskbar::stand_in_style` and `pin_real_window::real_pin_style`). They draw nothing here: this
/// host is a borderless popup with no `WS_CAPTION`, so there is no caption for a system menu to be
/// drawn on. What they are is the Shell's own read of a button.
///
/// **Why the icon.** A button is drawn from the window's own icon, and a window whose class carries
/// none is drawn with the default exe rectangle. The pin's host class carries the app's own (see
/// `create`); a host promoted out of a hover was made from the *hover's* class, which carries none
/// and must not, so the window is given it here, both sizes. Setting the icon a window already has is
/// no change, which is what makes this safe to ask at every road.
///
/// **Why a re-read after the style write, and what it is for.** The rule the rest of this app holds
/// styles to is that one written behind a window's back is not one it presents with until it is asked
/// to read it again (see `pin_real_window::set_the_pin_window_appwindow`), so the write is followed by
/// the same `SWP_FRAMECHANGED` call its ex-style road makes. There is no non-client frame here for
/// that word to recalculate — no caption, no border — so this call is the rule's price rather than a
/// frame's fix. **The read that matters for the button is the show**, and both roads have one: a host
/// this app made is created with the words before its first show, and a promotion writes them before
/// the show `the_button_follows_the_screen` pays for in the same pass (see
/// `pin_real_window::set_the_pin_window_appwindow`, which is why a show is what a button is read on).
unsafe fn the_button_shells_reads_its_gestures_from(host: HWND) {
    let style = unsafe { GetWindowLongPtrW(host, GWL_STYLE) };
    let wanted = style | (WS_SYSMENU | WS_MINIMIZEBOX).0 as isize;
    if wanted != style {
        let _ = unsafe { SetWindowLongPtrW(host, GWL_STYLE, wanted) };
        let _ = unsafe {
            SetWindowPos(
                host,
                None,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
            )
        };
    }

    // The window's own icon, both sizes: the small one is what a taskbar button draws, and the
    // large one what anything listing the window as a window draws.
    let icon = unsafe { crate::shell::app_icon::app_icon() };
    let icon = LPARAM(icon.0 as isize);
    let _ = unsafe { SendMessageW(host, WM_SETICON, WPARAM(ICON_BIG as usize), icon) };
    let _ = unsafe { SendMessageW(host, WM_SETICON, WPARAM(ICON_SMALL as usize), icon) };
}

/// When the take-up that is waiting on its hosted film began waiting, while one is.
///
/// A pin taken up on a hover's mpv film keeps the hover up and its own window down until the
/// hosted player draws (see `show_pinned_window`): the wait is this instant, and the tick's
/// handover ends it. Nothing else reads it — the wait is bounded by the player-start timeout
/// the handover settles on, and a pin that ends or collapses disarms it by going away.
static PIN_HANDOVER_SINCE: Lazy<Mutex<Option<Instant>>> = Lazy::new(|| Mutex::new(None));

/// The hover a take-up froze, and the second its picture stands at.
///
/// **It is the picture on screen until the pin can show the same one.** The take-up holds the
/// hover's player where it is rather than ending it, so the frame the user is looking at is mpv's own
/// last frame and nothing this app has to paint or cover: the wait ends when the pin's player reports
/// that frame drawn, and the swap is two windows showing one frame (see `settle_pin_handover`).
#[derive(Clone, Copy)]
struct FrozenHover {
    /// The hover's player, still running and held on the frame it was showing.
    pid: u32,
    /// The second that frame is at — the second the pin's player was begun on, and the one it has to
    /// be standing at before the swap is made.
    seconds: f64,
}

/// The hover a take-up is holding while the wait for its hosted player's frame is on.
static PIN_HANDOVER_HOVER: Lazy<Mutex<Option<FrozenHover>>> = Lazy::new(|| Mutex::new(None));

/// Hold the hover a take-up froze: the player still drawing the frame on screen, and the second it
/// stands at. Written by the take-up that begins the pin's player on that frame, and read by the
/// handover that swaps the two (see `FrozenHover`).
pub(super) fn hold_the_hover_for_the_handover(pid: u32, seconds: f64) {
    if let Ok(mut frozen) = PIN_HANDOVER_HOVER.lock() {
        *frozen = Some(FrozenHover { pid, seconds });
    }
}

/// The hover a take-up is holding, while the wait for its hosted player's frame is on.
fn held_hover() -> Option<FrozenHover> {
    PIN_HANDOVER_HOVER.lock().ok().and_then(|frozen| *frozen)
}

/// The hover a take-up froze, as a test reads it: the player still drawing the picture on screen and
/// the second the pin's player was begun on (see `FrozenHover`).
#[cfg(test)]
pub(super) fn held_hover_for_a_test() -> Option<(u32, f64)> {
    held_hover().map(|frozen| (frozen.pid, frozen.seconds))
}

/// End the hover a take-up froze, where one is still held: it is the picture no longer, and nothing
/// else holds it — the player was taken out of the media at the take-up, so this is the only call
/// that will ever end it. A kill not confirmed on the spot is parked with the orphans the loop
/// retries rather than dropped, which is the same road a pin's own teardown uses (see
/// `retire_orphaned_player`).
fn retire_the_frozen_hover(frozen: Option<FrozenHover>) {
    if let Some(frozen) = frozen {
        // The host the hover was drawn in is the picture no longer, and the pin it was waiting
        // for has already been shown: destroying the host ends mpv with it, so the retire below
        // is the confirmation rather than the picture's own end (see `hover_film_host`).
        unsafe { hover_film_host::take_down_the_hover_film_host_if_it_held(frozen.pid) };
        retire_orphaned_player(frozen.pid);
    }
}

/// Whether a pin taken up on a hover's film is still waiting on its hosted player's first frame:
/// its own window down, the hover up as the placeholder, and no chrome over either.
///
/// Asked by the roads that must leave a waiting pin alone: the tick's restore, which would put a
/// hidden-on-purpose window back up, and the chrome refresh, which would spend the arrival over
/// nothing (see `the_pin_owes_a_show` and `refresh_pin_chrome`).
///
/// **An arm cannot outlive the handover it was taken for, and that is a property of this answer
/// rather than of every road that ends one.** The wait is a pin that is not on the screen yet: the
/// moment the pin's own window is up — the handover, a box change that showed it, any road at all —
/// there is nothing left to wait for, whatever the arm's own bookkeeping says. And a player that has
/// not drawn by the start's own bound is a wait that is over the other way, which the tick settles
/// (see `settle_pin_handover`): an arm read here past that bound is a wait nothing is going to end,
/// and a strip held down by it is a caption the pointer can never ask back.
///
/// It is one predicate because it is one fact, and every reader asks it here: a road that forgot to
/// disarm cannot leave the chrome held down, the film paused, or the pin unshown for longer than the
/// event that ends the handover itself.
pub(super) fn pin_handover_armed() -> bool {
    let Some(since) = PIN_HANDOVER_SINCE.lock().ok().and_then(|since| *since) else {
        return false;
    };

    !the_handover_has_run_out(since) && !the_pins_own_window_is_on_the_screen()
}

/// Whether the wait a take-up began has run out: a player that has not drawn inside the start's own
/// bound is one that will not (see `VIDEO_START_WAIT_SECS`).
fn the_handover_has_run_out(since: Instant) -> bool {
    since.elapsed() >= Duration::from_secs(VIDEO_START_WAIT_SECS)
}

/// Whether the pin the handover is being waited for is on the screen, **which is the event that ends
/// a handover**: the pin's own window is what the wait exists to show (see `show_pinned_window`).
///
/// A pin standing in the preview's window is no window of its own at all, and a film is only ever
/// taken up in one of its own (see `pin_should_stand_in_its_own_window`), so there is nothing to ask
/// about there and the wait is the arm's alone.
fn the_pins_own_window_is_on_the_screen() -> bool {
    pin_real_window::real_pin_window()
        .is_some_and(|pin| unsafe { IsWindowVisible(pin) }.as_bool())
}

/// Arm the handover a take-up on a hover's film waits on, from the moment the hosted player is
/// begun in the background (see `show_pinned_window`).
pub(super) fn arm_pin_handover() {
    if let Ok(mut since) = PIN_HANDOVER_SINCE.lock() {
        *since = Some(Instant::now());
    }
}

/// Disarm the handover: the film drew and the pin went up, the wait failed and the pin is being
/// torn down, or the pin went away down another road and there is nothing left to wait on. The hover
/// the take-up froze goes with it: a wait that is over has no picture to hold (see `FrozenHover`).
pub(super) fn disarm_pin_handover() {
    if let Ok(mut since) = PIN_HANDOVER_SINCE.lock() {
        *since = None;
    }
    if let Ok(mut frozen) = PIN_HANDOVER_HOVER.lock() {
        *frozen = None;
    }
}

/// End the wait a take-up began and the hover it froze, for a pin that is going away before the
/// picture could change hands.
///
/// **What the hover was being kept for is the pin's**, so the pin's own end is the other end of
/// that wait: the second the pin's player was begun on is of no use to a window that is not going to
/// appear, and the hover it froze is a whole player of this app's held on one frame with nothing
/// left that would ever let go of it — the player was taken out of the media at the take-up, so the
/// handover is the only call that would have ended it. The loop's own settle disarms it too, on the
/// first tick that finds no pin up, but the close is where that state has to be given back rather
/// than a tick after it — the same reason a step's parked player is ended from the same exit (see
/// `settle_the_retirement_for_a_pin_that_is_going` and `settle_pin_handover`).
pub(super) unsafe fn settle_the_handover_for_a_pin_that_is_going() {
    let frozen = held_hover();
    disarm_pin_handover();
    retire_the_frozen_hover(frozen);
}

/// Stand a handover wait's beginning in place of the take-up's own, or give the take-up's own
/// answer back with `None` — a test of the timeout the wait is bounded by, which cannot wait out
/// the start timeout itself (see `settle_pin_handover`).
#[cfg(test)]
pub(super) fn stand_pin_handover_armed_at(since: Option<Instant>) {
    if let Ok(mut armed) = PIN_HANDOVER_SINCE.lock() {
        *armed = since;
    }
}

/// What the wait for a hosted film's first frame has come to.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum PinHandover {
    /// The film is not on screen yet and the player is still starting: the pin stays down and
    /// the hover stays up, and the wait goes on.
    Waiting,
    /// The hosted player drew: the pin went up over its live film with the chrome arriving,
    /// and the hover came down in the same pass.
    HandedOver,
    /// The player errored or never drew inside the player-start timeout: the hidden pin is for
    /// the teardown, and the hover it never took down is what is kept.
    Failed,
}

/// How far from the second it was begun on a hosted player may be and still be standing on the frame
/// the hover froze.
///
/// A frame's own timestamp is not the second it was sought to, and a film whose frames are a second
/// apart is a film whose frame at 12.345 is stamped 12.000. It is not a tolerance on the swap — the
/// player is begun paused, so it cannot be anywhere but on that frame — it is the width of the gap
/// between a seek target and the frame that answers it (see `the_hosted_frame_is_presented`).
const PIN_HANDOVER_FRAME_SECS: f64 = 1.0;

/// Whether this hosted player has a frame of its own on screen: the source the Shell draws it from
/// *and*, where the player can be asked, a video output it says is configured for a frame.
///
/// A window is not a picture (see `the_host_is_drawing`), and `vo-configured` is the
/// player's own word that its video output has been set up for a frame. The answer is latched per
/// player because the read is bounded by a pipe timeout rather than free: once this player has
/// proved a frame it has one, and every later tick answers at once. A step begins a new player, so
/// `note_hosted_player` clears the latch for it.
///
/// **A player with no pipe to answer through is one whose picture is only ever its window**, and the
/// source being up is the whole of the proof for it: FFmpeg's player reports nothing at all (see
/// `mpv_ipc::answers_properties`), and mpv answers the same question with its own video output.
pub(super) fn the_hosted_frame_is_up(pid: u32) -> bool {
    if pid != 0 && PRESENTED_PID.load(Ordering::SeqCst) == pid {
        return true;
    }
    if !the_hosted_player_is_drawing(pid) {
        return false;
    }
    if !the_players_output_is_configured(pid) {
        return false;
    }
    PRESENTED_PID.store(pid, Ordering::SeqCst);
    true
}

/// Whether this player's own video output is configured for a frame, where the player answers over a
/// pipe: **the half of the presented predicate that does not need to know which second it is on**.
///
/// It is asked where a player has just been begun and the second it stands on is the caller's own
/// (see `pin_film_thumbnail::hand_the_picture_over`), and it is the same fact `the_hosted_frame_is_up`
/// proves a frame with, so the two cannot answer differently about the same player.
pub(super) fn the_players_output_is_configured(pid: u32) -> bool {
    !mpv_ipc::answers_properties(pid)
        || mpv_ipc::get_property(pid, "vo-configured").as_deref() == Some("true")
}

/// Whether the hosted player has the frame the hover froze on screen, rather than only a window.
///
/// **A window is not a picture.** A player makes its window before it has decoded anything, and a
/// source adopted for it is the most `the_hosted_player_is_drawing` can see — which leaves the band
/// black until the first frame is drawn into it, and is the empty stage the whole wait exists to keep
/// behind the hover. The two facts that separate a window from a picture are the player's own, read
/// over its pipe: `time-pos` has no value until a file is loaded and positioned, and
/// `vo-configured` is true once the video output is configured for a frame. **Both**, because either
/// alone is a weaker question — a position is where the file is and not what has been drawn, and a
/// configured video output is a size and a format and not a frame.
///
/// The player is begun paused on that second (see `video_launch::mpv_args`), so the pair is not a
/// race: it is true exactly while the frame at that second is the one on screen, and the unpause at
/// the handover is the first thing that moves it. A read that does not answer is a player that has
/// not proved anything, so the wait goes on — bounded by the player-start timeout, never by a timer
/// of its own (see `VIDEO_START_WAIT_SECS`).
pub(super) fn the_hosted_frame_is_presented(pid: u32, at: f64) -> bool {
    let Some(seconds) =
        mpv_ipc::get_property(pid, "time-pos").and_then(|answer| answer.parse::<f64>().ok())
    else {
        return false;
    };

    the_hosted_frame_is_up(pid) && (seconds - at).abs() <= PIN_HANDOVER_FRAME_SECS
}

/// Hand a pin taken up on a hover's film over once its hosted player has the frame on screen, and
/// tear the hidden pin down where that player never does.
///
/// The only visible swap, and it is one frame: the pin's window appears with the film already showing
/// the frame the hover was frozen on — no cover, no black stage — the arrival starts over that live
/// picture, the hover comes down, the hosted player is let go of, and the frozen hover is ended, all
/// in the same pass (see `show_pinned_window`). What the wait waits on is the player's own word that
/// the frame is drawn and not a window that has been made (see `the_hosted_frame_is_presented`), and
/// nothing here is a timer: a player that takes its time is waited for, up to the player-start bound.
/// Where the player errored or the start timeout ran out, the wait is over the other way: the pin
/// never appears, the hover it never took down is what is kept, and the hover's own player is ended
/// rather than left frozen for a handover that is never coming (see `Reason::MediaGone`).
///
/// One shot: the arm goes down however the wait ends, so a handover never fires twice and a
/// failure never sticks behind a hidden pin.
pub(super) unsafe fn settle_pin_handover() -> PinHandover {
    let Some(since) = PIN_HANDOVER_SINCE.lock().ok().and_then(|since| *since) else {
        return PinHandover::Waiting;
    };
    let frozen = held_hover();

    // A pin that went away down another road — closed, or collapsed into its bubble — leaves no
    // wait behind it: the close owns the teardown and the collapse owns the pin now. The hover the
    // take-up froze is ended with it, because the pin that was going to take the picture from it is
    // not going to.
    if !pinned() || pin_is_collapsed() {
        disarm_pin_handover();
        retire_the_frozen_hover(frozen);
        return PinHandover::Waiting;
    }

    // The film's own frame on screen is the handover: the wait is over, and what is on screen from
    // here is the pin over that picture rather than the hover as its placeholder. A frame this app
    // cannot prove is a frame it waits for — the read failing is not an error, it is not-yet.
    //
    // **And a pin whose own window is already on the screen is a handover that has been overtaken**,
    // which is the same wait ending one step early: whatever road showed it — a box change, a
    // restore, the Shell — has put the pin up before its film's first frame was proved, and what is
    // left to do is the hover's own ending rather than a show. The arm's readers answer this case
    // already (see `pin_handover_armed`); this is the wait behind them, so that the hover the take-up
    // froze is let go of here rather than a start's bound later.
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    let overtaken = the_pins_own_window_is_on_the_screen();
    if overtaken
        || (the_hosted_player_is_drawing(pid)
            && frozen.is_some_and(|frozen| the_hosted_frame_is_presented(pid, frozen.seconds)))
    {
        note_pin_trace!(
            "handover handed over after {}ms{}",
            since.elapsed().as_millis(),
            if overtaken { " (the pin was already up)" } else { "" },
        );
        disarm_pin_handover();
        if !overtaken {
            // The handover is the pin's first frame, not an arrival: the chrome stays hidden over it
            // until the pointer asks for a strip (see `refresh_pin_chrome`). A pin that is already on
            // the screen keeps the chrome it has: the strips are the pointer's answer, and this wait
            // is not a road that may take one out from under a hand.
            with_pin(|pin| pin.chrome = PinChrome::hidden());
            unsafe { show_pinned_window(pinned_window()) };
        }

        // **And the film's own window is not asked for anything here**: it was placed once, at the
        // film's own size, when it was adopted, and a box the pin has now is a destination rectangle
        // (see `pin_film_thumbnail`).

        // **The hover is ended before the pin's player is let go of.** The two are the same sound at
        // the same second, and the unpause is a post to one process while the kill is a termination of
        // another: asked in the other order, both are audible until the kill lands, which is heard as
        // an echo of the frame the handover is standing on. The pin's window is already up over that
        // frame (see `show_pinned_window`), so the hover's going is the picture no longer and nothing
        // is left on screen for the order to change.
        retire_the_frozen_hover(frozen);

        // **And the film is playing from here, unless the user had held it.** The player was begun
        // paused so that the frame the hover froze is the frame it has; letting it go of that hold is
        // this moment and no earlier one — and the clock the pass is measured by moves onto the
        // second it is let go at, because a clock begun at the launch has counted the whole wait
        // (see `PinTransport::released` and `rebase_video_loop_clock_to`).
        if let Some(frozen) = frozen.filter(|_| !pinned_is_held()) {
            mpv_ipc::set_paused(pid, false);
            video_hw::rebase_video_loop_clock_to(pid, frozen.seconds);
            update_pin_transport(|transport| transport.released(frozen.seconds));
        }
        return PinHandover::HandedOver;
    }

    // No player behind the wait is a player that errored rather than one still starting: the
    // pin it was begun for never appears, and the hover it never took down is what is kept.
    if !is_video_process_running() {
        disarm_pin_handover();
        retire_the_frozen_hover(frozen);
        return PinHandover::Failed;
    }

    // And a player that never draws inside the player-start timeout is one that will not: a
    // hidden pin waiting past a start ever takes is a stuck one (see `VIDEO_START_WAIT_SECS`).
    if since.elapsed() >= Duration::from_secs(VIDEO_START_WAIT_SECS) {
        note_pin_trace!(
            "handover failed: no frame after {}ms",
            since.elapsed().as_millis()
        );
        disarm_pin_handover();
        retire_the_frozen_hover(frozen);
        return PinHandover::Failed;
    }

    PinHandover::Waiting
}

#[cfg(test)]
mod tests {
    use super::*;

    /// **A show the trace already wrote is not written again**: the tick asks every pass, so a host
    /// shown, a host hidden and a host shown again are three lines, and a host shown once is one.
    #[test]
    fn a_host_show_is_written_when_it_changes_and_not_on_every_pass() {
        assert!(
            the_show_is_new(None, 0x2000, true),
            "the first show of a run is the state it began in"
        );
        assert!(
            !the_show_is_new(Some((0x2000, true)), 0x2000, true),
            "and an answer that has not moved is not written on every tick"
        );
        assert!(
            the_show_is_new(Some((0x2000, true)), 0x2000, false),
            "a hide is a change"
        );
        assert!(
            the_show_is_new(Some((0x2000, true)), 0x3000, true),
            "and so is another host"
        );
    }

    /// **The per-tick line is written on change and at most once a second otherwise**, which is what
    /// keeps a run's trace readable: the tick asks `the_hosted_player_is_drawing` hundreds of times a
    /// second and each answer it has not moved is one line an unchanged answer owes nobody.
    #[test]
    fn a_hosted_drawing_answer_is_traced_on_change_and_once_a_second_otherwise() {
        let now = Instant::now();
        let answer = (7, 0x2000, true, now);
        assert!(
            the_drawing_trace_is_due(None, answer, now),
            "the first answer of a run is the state it began in"
        );
        assert!(
            !the_drawing_trace_is_due(Some(answer), answer, now + Duration::from_millis(200)),
            "and an answer that has not moved is not written on every tick"
        );
        assert!(
            the_drawing_trace_is_due(Some(answer), answer, now + HOSTED_DRAWING_TRACE_EVERY),
            "the heartbeat is what says the tick is still asking it"
        );
        assert!(
            the_drawing_trace_is_due(
                Some(answer),
                (7, 0x2000, false, now + Duration::from_millis(1)),
                now + Duration::from_millis(1)
            ),
            "a transition is never missed, however soon after the last line it is: a player that
             draws and stops and draws again is three answers and not one"
        );
        assert!(
            the_drawing_trace_is_due(
                Some(answer),
                (8, 0x2000, true, now + Duration::from_millis(1)),
                now + Duration::from_millis(1)
            ),
            "and another player is another answer"
        );
        assert!(
            the_drawing_trace_is_due(
                Some(answer),
                (7, 0x3000, true, now + Duration::from_millis(1)),
                now + Duration::from_millis(1)
            ),
            "as is the same player in another host"
        );
    }
}

//! The FFPLAY video pin's own road, written down when `RHP_PIN_TRACE` asks for it: the take-up,
//! the window and host it draws through, every decision to show or skip it, and what a hover's
//! `Hide` does to a pin that is already standing.
//!
//! **And what a window of the pin's is handed as a hand arrives on it**: the messages a click on the
//! pair's taskbar button delivers, named by the window they arrive at rather than by a handle
//! nothing can be read out of (see `window_name` and `note_a_message_of_a_window`).
//!
//! The gate is one atomic read: a run without the variable writes nothing and builds no line.

use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Once;

use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::UI::Input::KeyboardAndMouse::{GetFocus, IsWindowEnabled};
use windows::Win32::UI::WindowsAndMessaging::{
    GetForegroundWindow, IsIconic, IsWindowVisible, SC_CLOSE, SC_MINIMIZE, SC_RESTORE,
    SIZE_MAXIMIZED, SIZE_MINIMIZED, SIZE_RESTORED, WA_ACTIVE, WA_CLICKACTIVE, WA_INACTIVE,
    WM_ACTIVATE, WM_ACTIVATEAPP, WM_CLOSE, WM_ENABLE, WM_KILLFOCUS, WM_MOUSEACTIVATE,
    WM_NCACTIVATE, WM_QUERYOPEN, WM_SETFOCUS, WM_SHOWWINDOW, WM_SIZE, WM_SYSCOMMAND,
};

static READ: Once = Once::new();
static ASKED: AtomicBool = AtomicBool::new(false);

/// Whether `RHP_PIN_TRACE` asks for the pin's road to be written out.
pub(super) fn asked() -> bool {
    READ.call_once(|| {
        ASKED.store(
            std::env::var_os("RHP_PIN_TRACE").is_some(),
            Ordering::Relaxed,
        );
    });
    ASKED.load(Ordering::Relaxed)
}

/// Write one timestamped line to `%TEMP%\rhp-pin-trace.log`, appended and flushed so a run that
/// ends badly leaves what it wrote.
pub(super) fn write(line: &str) {
    let stamp = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|elapsed| format!("{}.{:03}", elapsed.as_secs(), elapsed.subsec_millis()))
        .unwrap_or_else(|_| "-".to_string());
    if let Ok(mut file) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(std::env::temp_dir().join("rhp-pin-trace.log"))
    {
        use std::io::Write;
        let _ = writeln!(file, "{stamp} {line}");
        let _ = file.flush();
    }
}

/// Whether a line is going anywhere at all: the log `RHP_PIN_TRACE` names, or — in a test build —
/// the seam that reads a line back out of a run (see `NOTED`).
///
/// It is what a road asks *before* it works out what a line says. The macro is lazy about its own
/// arguments, but a line whose parts are read into locals first (a state, a word, `GetFocus`) is read
/// for nothing when nothing is going anywhere — and `asked` alone would keep a test from reading a
/// line at all, because the tests run without the variable.
pub(super) fn wanted() -> bool {
    asked() || cfg!(test)
}

/// Which of this app's windows a handle names, in the one word a line can be read by.
///
/// **An `HWND` is a number nothing can be read out of**, and the whole of what a line about one asks
/// is *which* window a message arrived at: the pin's film host, the window the pin stands in, the
/// preview's own window, the bubble a collapsed pin becomes, or the stand-in wearing its button.
///
/// **Asked by comparing handles, never by asking Windows about the window.** This runs from inside
/// window procedures, where the window a message is about may be on its way out, and where a
/// question that cost a call would be paid on every message of a hand on a button. Nothing here waits
/// for the pin's lock — every handle read is a published one — so a procedure that is answering a
/// message about the pin can still ask it. A handle that names none of these is `other`, and one that
/// names nothing at all is `none`.
pub(super) fn window_name(hwnd: HWND) -> &'static str {
    if hwnd.0.is_null() {
        return "none";
    }

    if super::pin_film_host::film_host() == Some(hwnd) {
        return "host";
    }

    // **The pin's window is asked of the pin first**, because a pin that stands in this app's own
    // window answers `pinned_window` with that window rather than with one of its own: named the
    // other way round, the window a pin is in would read as a hover's (see `pinned_window`).
    if super::pinned() && super::pinned_window() == hwnd {
        return "pin";
    }

    if HWND(super::PREVIEW_HWND.load(Ordering::SeqCst) as *mut _) == hwnd {
        return "preview";
    }

    if super::pin_bubble::showing_bubble_window() == Some(hwnd) {
        return "bubble";
    }

    if super::pin_taskbar::showing_taskbar_window() == Some(hwnd) {
        return "stand-in";
    }

    "other"
}

/// Whether a message a pin's window is handed is one a hand on the pair's taskbar button could
/// arrive as, and so is worth a line: the activations, the system command and the size a window is
/// put down by, the show and the enable, and the close. **Paint, erase, cursor, pointer, thumbnail
/// and placement traffic is left out** — it is most of what a run of messages is, and none of it is a
/// hand on a button. `WM_WINDOWPOSCHANGED` in particular is the placement tick's own: the pair is
/// re-positioned constantly, and it was 538 of the 637 lines a traced click wrote.
fn a_click_could_be_this(msg: u32) -> bool {
    matches!(
        msg,
        WM_ACTIVATE
            | WM_SETFOCUS
            | WM_KILLFOCUS
            | WM_ACTIVATEAPP
            | WM_NCACTIVATE
            | WM_SYSCOMMAND
            | WM_SIZE
            | WM_SHOWWINDOW
            | WM_QUERYOPEN
            | WM_MOUSEACTIVATE
            | WM_ENABLE
            | WM_CLOSE
    )
}

/// A message's name as Windows' own documentation gives it, so that a log can be read against it.
fn the_messages_name(msg: u32) -> &'static str {
    match msg {
        WM_ACTIVATE => "WM_ACTIVATE",
        WM_SETFOCUS => "WM_SETFOCUS",
        WM_KILLFOCUS => "WM_KILLFOCUS",
        WM_ACTIVATEAPP => "WM_ACTIVATEAPP",
        WM_NCACTIVATE => "WM_NCACTIVATE",
        WM_SYSCOMMAND => "WM_SYSCOMMAND",
        WM_SIZE => "WM_SIZE",
        WM_SHOWWINDOW => "WM_SHOWWINDOW",
        WM_QUERYOPEN => "WM_QUERYOPEN",
        WM_MOUSEACTIVATE => "WM_MOUSEACTIVATE",
        WM_ENABLE => "WM_ENABLE",
        WM_CLOSE => "WM_CLOSE",
        _ => "other",
    }
}

/// The word for what a message carries in its `wParam`, where the message has one: the activation
/// state, the system command, or the size a window is reporting. `-` for every message whose
/// `wParam` is a flag or a handle rather than one of those — the number itself is written beside it,
/// and a name for something this app does not read would be a word nobody asked for.
fn the_messages_word(msg: u32, wparam: WPARAM) -> String {
    match msg {
        // The low word is the activation state; the high word says whether the window arrived
        // minimized, which is the same request (see `answer_an_activation_of_the_host`).
        WM_ACTIVATE => match (wparam.0 & 0xFFFF) as u32 {
            WA_ACTIVE => "WA_ACTIVE".to_string(),
            WA_CLICKACTIVE => "WA_CLICKACTIVE".to_string(),
            WA_INACTIVE => "WA_INACTIVE".to_string(),
            other => other.to_string(),
        },
        // The low four bits of a system command belong to whoever sent one, so the command is the
        // word without them — the same mask the road that reads a minimize makes of it (see
        // `pin_film_host::the_shell_put_the_host_down`).
        WM_SYSCOMMAND => match (wparam.0 & 0xFFF0) as u32 {
            SC_MINIMIZE => "SC_MINIMIZE".to_string(),
            SC_RESTORE => "SC_RESTORE".to_string(),
            SC_CLOSE => "SC_CLOSE".to_string(),
            other => other.to_string(),
        },
        WM_SIZE => match wparam.0 as u32 {
            SIZE_MINIMIZED => "SIZE_MINIMIZED".to_string(),
            SIZE_RESTORED => "SIZE_RESTORED".to_string(),
            SIZE_MAXIMIZED => "SIZE_MAXIMIZED".to_string(),
            other => other.to_string(),
        },
        _ => "-".to_string(),
    }
}

/// Write one line about a message a window of the pin's is handed: **what a hand on the pair's
/// taskbar button actually delivers**, which is the one thing a run of the toggle it was added for
/// cannot answer from inside.
///
/// **Three roads were assumed for that click** — an activation of the film host, the Shell's
/// `SC_MINIMIZE`, and the `SIZE_MINIMIZED` a window reports as it is put down (see
/// `pin_film_host::the_shell_put_the_host_down`) — and a click that still does nothing on a real
/// desktop says which of them the mouse really takes, and which window answers it.
///
/// **What is written beside the message is the state every reading of one is a reading of**: whether
/// this window is the pin's own host, what the pin is, where the window is, and which window the
/// foreground and the caret are in — because the same `WM_ACTIVATE` is the keyboard handover on one
/// state and the minimize on the other (see `pin_film_host::answer_a_second_click_on_the_hosts_button`).
///
/// `subject` is what the window is called in the line's own words (`film host`, `pin`).
pub(super) fn note_a_message_of_a_window(
    subject: &str,
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) {
    // Nothing is built where nothing is going, and the message is asked first: it is the answer a
    // window procedure gives most of the time (see `wanted`).
    if !a_click_could_be_this(msg) || !wanted() {
        return;
    }

    // Safety: both are questions about the desktop's own bookkeeping. Neither can fault, and a
    // window that is not this app's is `other` (see `window_name`).
    let foreground = unsafe { window_name(GetForegroundWindow()) };
    let focus = unsafe { window_name(GetFocus()) };
    note_pin_trace!(
        "{subject} msg {} arg={} wparam={} lparam={} own-host={} window={} pinned={} collapsed={} \
         pin-window={} always-on-top={} visible={} iconic={} enabled={} foreground={foreground} \
         focus={focus}",
        the_messages_name(msg),
        the_messages_word(msg, wparam),
        wparam.0,
        lparam.0,
        (super::pin_film_host::film_host() == Some(hwnd)) as u8,
        window_name(hwnd),
        super::pinned() as u8,
        super::pin_is_collapsed() as u8,
        super::pin_is_its_own_window() as u8,
        super::pin_always_on_top() as u8,
        // Safety: the window is one of this app's own, and these only read what Windows keeps of it.
        unsafe { IsWindowVisible(hwnd) }.as_bool() as u8,
        unsafe { IsIconic(hwnd) }.as_bool() as u8,
        unsafe { IsWindowEnabled(hwnd) }.as_bool() as u8,
    );
}

/// A path as the name of the file in it: the folder is the same on both sides of every question a
/// line asks, and a path long enough to be worth reading is a path too long to read.
pub(super) fn path_name(path: &std::path::Path) -> String {
    path.file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_else(|| "-".to_string())
}

/// The lines a test run wrote, so a test can hold the road to what it said rather than to what the
/// log happened to be configured to keep: the same seam `mpv_ipc` keeps its command lines with, and
/// for the same reason (see `note_pin_trace`).
#[cfg(test)]
static NOTED: std::sync::Mutex<Vec<String>> = std::sync::Mutex::new(Vec::new());

/// Write one line down for a test that reads it (see `NOTED`).
#[cfg(test)]
pub(super) fn note_a_line_for_a_test(line: String) {
    if let Ok(mut noted) = NOTED.lock() {
        noted.push(line);
    }
}

/// The lines this run has written (see `NOTED`).
#[cfg(test)]
pub(super) fn noted_lines() -> Vec<String> {
    NOTED
        .lock()
        .map(|noted| noted.clone())
        .unwrap_or_default()
}

/// Forget the lines written before a test (see `NOTED`).
#[cfg(test)]
pub(super) fn clear_noted() {
    if let Ok(mut noted) = NOTED.lock() {
        noted.clear();
    }
}

/// Write one line, and build nothing at all where the trace is off. A test run keeps every line it
/// wrote, whether the log is on or not (see `NOTED`).
macro_rules! note_pin_trace {
    ($($line:tt)*) => {{
        if super::pin_trace::asked() {
            super::pin_trace::write(&format!($($line)*));
        }
        #[cfg(test)]
        super::pin_trace::note_a_line_for_a_test(format!($($line)*));
    }};
}

pub(super) use note_pin_trace;

#[cfg(test)]
mod tests {
    use super::*;

    /// **A window is named and never numbered**: a line about a message says which window of this
    /// app's answered it, and a handle names nothing at all — one that is no window is `none` and one
    /// that is none of these windows is `other`. Both are asked with no pin state, because that is
    /// what the naming has to answer from: the handles are published, and a window procedure asks
    /// this about a window Windows may already be taking apart.
    #[test]
    fn a_window_is_named_by_what_this_app_knows_of_it() {
        assert_eq!(
            window_name(HWND(std::ptr::null_mut())),
            "none",
            "the handle that names no window is named none rather than other"
        );
        assert_eq!(
            window_name(HWND(0x5EED_1EAF as *mut _)),
            "other",
            "and a window of nobody's that this app knows nothing of is other"
        );
    }
}

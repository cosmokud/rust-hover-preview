//! The app's own questions, and the one thing all of them need from the platform: a
//! message box asked with this app's words on its buttons.
//!
//! `MessageBoxW` is what asks. It is the platform's own dialog, and this app has no
//! window to put one of its own in — the one window it owns is the tray's, which is
//! never shown — so the box is given no owner and is set to the foreground instead.
//! What it is not told is what its buttons mean: `Yes` and `No` are the platform's
//! names for a question this app is not asking, and the names it does ask for are put
//! on them as the dialog is created (see `hook`).
//!
//! Whether one of these dialogs is on screen is answered here rather than by whoever
//! asked, because the tray has to know about all of them: a menu opened over a dialog
//! would be a second way into the same question, and a message box owns no window of
//! this app's and so is modal to nothing.

use std::cell::RefCell;
use std::ffi::c_void;
use std::sync::atomic::{AtomicBool, Ordering};
use windows::core::PCWSTR;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::WindowsAndMessaging::{
    CallNextHookEx, GetDlgItem, MessageBoxW, SetWindowTextW, SetWindowsHookExW,
    UnhookWindowsHookEx, HCBT_ACTIVATE, IDNO, IDYES, MB_DEFBUTTON2, MB_ICONINFORMATION,
    MB_ICONWARNING, MB_SETFOREGROUND, MB_YESNO, MESSAGEBOX_STYLE, WH_CBT,
};

/// The caption every dialog of this app carries.
const CAPTION: &str = "Rust Hover Preview";

/// How many settings the reset question names before it stops counting them out. A
/// list of sixty keys is not a question anybody reads; the count that is left over is
/// said in one line instead.
const NAMED_SETTINGS: usize = 8;

/// Whether one of this app's dialogs is on screen right now.
static CONFIRMING: AtomicBool = AtomicBool::new(false);

thread_local! {
    /// The names the two buttons of the dialog being asked are to carry, the
    /// affirmative one first.
    static BUTTON_NAMES: RefCell<Option<(String, String)>> = const { RefCell::new(None) };
}

/// Whether one of this app's dialogs is on screen right now.
pub(crate) fn is_confirming() -> bool {
    CONFIRMING.load(Ordering::SeqCst)
}

/// Note that a dialog is about to be shown, and that it is gone.
///
/// The two are the caller's to pair: a dialog is a call that does not come back until it
/// is answered, so what is between them is one question and nothing else.
pub(crate) fn begin() {
    CONFIRMING.store(true, Ordering::SeqCst);
}

pub(crate) fn end() {
    CONFIRMING.store(false, Ordering::SeqCst);
}

/// Name the buttons of the dialog about to be shown. The two are the platform's names
/// for a question this app is not asking, so what it does ask is put here instead.
pub(crate) fn set_button_names(affirmative: &str, negative: &str) {
    BUTTON_NAMES.with(|names| {
        *names.borrow_mut() = Some((affirmative.to_string(), negative.to_string()));
    });
}

/// Name the two buttons of a dialog as it is created.
///
/// The hook belongs to one thread, the one asking, and its life is one dialog: the
/// message box is created by the thread that calls for it, so what it names is a window
/// of this process's own, and no other process can meet it or be reached by it.
/// `HCBT_ACTIVATE` is the moment the dialog is whole — every button of it exists by the
/// time it is about to be shown — and a window that is not this dialog has no button
/// under these ids, which is how the two lookups answer nothing and it is left exactly
/// as it is.
pub(crate) unsafe extern "system" fn hook(code: i32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    if code == HCBT_ACTIVATE as i32 {
        let dialog = HWND(wparam.0 as *mut c_void);

        BUTTON_NAMES.with(|names| {
            if let Some((affirmative, negative)) = names.borrow().as_ref() {
                for (id, label) in [(IDYES, affirmative), (IDNO, negative)] {
                    if let Ok(button) = GetDlgItem(dialog, id.0) {
                        let text = wide(label);
                        let _ = SetWindowTextW(button, PCWSTR(text.as_ptr()));
                    }
                }
            }
        });
    }

    CallNextHookEx(None, code, wparam, lparam)
}

/// Ask a question that has two answers, and answer with what was said.
///
/// The second button is the default one, which is the whole of what makes a question
/// safe to put in front of a hand: what Enter or a stray space answers is the answer
/// that does nothing.
///
/// The icon is the caller's, because the two ends of what this app asks are not the
/// same kind of question: a reset changes a file and says so with a warning, while
/// the question about a page changes nothing and is told as what it is.
fn confirm(text: &str, affirmative: &str, icon: MESSAGEBOX_STYLE) -> bool {
    let caption = wide(CAPTION);
    let text = wide(text);

    set_button_names(affirmative, "Cancel");
    begin();

    let hook = unsafe { SetWindowsHookExW(WH_CBT, Some(hook), None, GetCurrentThreadId()) };

    let answer = unsafe {
        MessageBoxW(
            HWND::default(),
            PCWSTR(text.as_ptr()),
            PCWSTR(caption.as_ptr()),
            MB_YESNO | icon | MB_DEFBUTTON2 | MB_SETFOREGROUND,
        )
    };

    if let Ok(hook) = hook {
        unsafe {
            let _ = UnhookWindowsHookEx(hook);
        }
    }

    end();

    answer == IDYES
}

/// The question the `Reset to Recommended Settings` row asks, naming what would change.
///
/// What it names is the file's own keys, in the file's own words, so what a user is
/// agreeing to is the change they would see in `config.ini` — and the lists are said
/// out loud to be untouched, since the other row is the one that touches those.
pub(crate) fn confirm_reset_settings(changes: &[(String, String, String)]) -> bool {
    let mut text = String::from("Reset these settings to the values this build recommends?\n\n");

    for (key, now, recommended) in changes.iter().take(NAMED_SETTINGS) {
        text.push_str(&format!("{key}: {now} -> {recommended}\n"));
    }

    if changes.len() > NAMED_SETTINGS {
        text.push_str(&format!("and {} more\n", changes.len() - NAMED_SETTINGS));
    }

    text.push_str("\nYour extension lists are not touched.");

    confirm(&text, "Reset", MB_ICONWARNING)
}

/// The question the `Reset Extension Lists` row asks, naming the lists that would go.
pub(crate) fn confirm_reset_lists(sections: &[String]) -> bool {
    let text = format!(
        "Replace these lists with the built-in ones?\n\n{}\n\n\
         Every name you added is lost, and every name you removed comes back.\n\
         Your other settings are not touched.",
        sections.join(", ")
    );

    confirm(&text, "Reset", MB_ICONWARNING)
}

/// The question a missing row of the tray's `Codecs` submenu asks: whether to open the page
/// the engine or codec is got from.
///
/// Nothing is installed from here and nothing is fetched: a yes hands the page to the
/// browser the user already has, which is the page the README names for the thing, and a no
/// leaves the machine exactly as it was — so nothing about this question is a warning, and
/// the address is in it so that what a yes opens is read before it is opened.
pub(crate) fn confirm_open_page(name: &str, url: &str) -> bool {
    let text = format!(
        "Open the download page for {name}?\n\n{url}\n\n\
         The page opens in your browser. Nothing is installed by this app."
    );

    confirm(&text, "Open", MB_ICONINFORMATION)
}

/// A string as a message box wants it: UTF-16, terminated.
fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}

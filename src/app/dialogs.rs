//! The app's own questions, and the one thing all of them need from the platform: a
//! message box asked with this app's words on its buttons.
//!
//! `MessageBoxW` is what asks. It is the platform's own dialog, and this app has no
//! window to put one of its own in for the tray — the one window it owns is the tray's,
//! which is never shown — so a box asked from the tray is given no owner and is set to
//! the foreground instead. What it is not told is what its buttons mean: `Yes` and `No`
//! are the platform's names for a question this app is not asking, and the names it does
//! ask for are put on them as the dialog is created (see `hook`).
//!
//! **The one question that is about a window of this app's own is not asked here.** A delete
//! from the `Custom Preview` editor is asked in a card of this app's own, drawn from the same
//! palette as the window it belongs to and with no sound from the platform at all (see
//! `ui::ask_card`) — a message box cannot be themed, and a delete is the one question here that
//! has to look like the window it destroys something of.
//!
//! Whether one of these dialogs is on screen is answered here rather than by whoever
//! asked, because the tray has to know about all of them: a menu opened over a dialog
//! would be a second way into the same question, and a message box owns no window of
//! this app's and so is modal to nothing. A custom card is announced through the same two
//! calls (`begin`/`end`), so the tray's rule holds for it as well.

use std::cell::RefCell;
use std::ffi::c_void;
use std::sync::atomic::{AtomicBool, Ordering};
use windows::core::PCWSTR;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::WindowsAndMessaging::{
    CallNextHookEx, GetDlgItem, MessageBoxW, SetWindowTextW, SetWindowsHookExW,
    UnhookWindowsHookEx, HCBT_ACTIVATE, IDNO, IDYES, MB_DEFBUTTON2, MB_ICONINFORMATION,
    MB_ICONWARNING, MB_OK, MB_SETFOREGROUND, MB_YESNO, MESSAGEBOX_STYLE, WH_CBT,
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
///
/// The box is given no owner: it is asked from a tray whose window is never shown, so the
/// platform's own is the only window there is to be modal to, and it is set to the foreground
/// instead. The one question this app asks *about* a window of its own is not asked here (see
/// `ui::ask_card`).
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

/// The question a missing row of the tray's `Components` submenu asks: whether to open the page
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

/// The question a row of the `Components → Install` submenu asks: whether to install the engine with
/// `winget`, which is a download, an installer and possibly a UAC prompt rather than a page.
///
/// The package id is in the question for the same reason a page's address is: what a yes runs is
/// read before it runs, and an engine whose `winget` package is not what a user expected is one
/// they can refuse here. What a yes opens is a console window, and it is said so: the window is
/// where the install is watched and where a failure is read.
pub(crate) fn confirm_install(name: &str, id: &str) -> bool {
    let text = format!(
        "Install {name} with winget?\n\n{id}\n\n\
         A console window opens and shows the install, without administrator rights. The machine's \
         installers for LibreOffice, calibre, ImageMagick and PeaZip may raise their own UAC prompt \
         — that is theirs, and it is expected."
    );

    confirm(&text, "Install", MB_ICONINFORMATION)
}

/// What an install that failed says: the code `winget` left, what was left alone, and where the
/// reason is — the console window that is still up, showing what winget printed.
pub(crate) fn report_install_failure(name: &str, code: i32) {
    let text = format!(
        "{name} could not be installed.\n\nwinget exited with code {code}.\n\n\
         The winget window that was open shows why. Its page is still on the row above, and \
         nothing else was changed."
    );

    report(&text);
}

/// What the `Run at Startup` row says when the machine would not take the entry.
///
/// It is said out loud rather than left to a checkmark that flips back: a `Run` key a policy holds,
/// a hive this process may not write, and an entry Windows refuses to run all read the same way
/// from the tray — as a switch that does nothing — and the one thing a user can act on is what the
/// machine refused.
pub(crate) fn report_startup_failure(outcome: crate::app::startup::Outcome) {
    // The mark on the file is not a startup entry, and a box headed "the startup entry could not be
    // written" over it would be naming the wrong thing: what the machine refused is the one flag
    // Windows reads for a launch that is not a logon's.
    if let crate::app::startup::Outcome::MarkRefused { on } = outcome {
        report_mark_failure(on);
        return;
    }

    // And a task that stayed is the row half in place rather than a write that failed: the entry is
    // written and a logon starts the app elevated beside it, which is the one thing turning the row
    // off was about.
    if let crate::app::startup::Outcome::TaskStayed = outcome {
        report_task_stayed();
        return;
    }

    let state = match outcome {
        // The machine refused the write and still holds the opposite of what was asked for.
        crate::app::startup::Outcome::NotWritten { held: false } => {
            "The app will not start when you sign in until this is fixed."
        }
        crate::app::startup::Outcome::NotWritten { held: true } => {
            "The old setting is still there, so the app still starts when you sign in."
        }
        crate::app::startup::Outcome::NeedsElevation => {
            "Changing the administrator startup needs administrator permission, and this app does \
             not have it right now, so nothing was changed."
        }
        crate::app::startup::Outcome::TaskRefused => {
            "Windows would not accept the administrator startup, so nothing was changed."
        }
        _ => "Another startup setting is still in place, so the app still starts when you sign in.",
    };

    let text = format!(
        "Windows did not accept the startup setting.\n\n\
         The app could not be added to the list of programs that start when you sign in. {state}\n\n\
         This usually happens because of a security policy or antivirus software, or because the \
         administrator prompt was answered with No. The rest of the app still works."
    );

    report(&text);
}

/// What the `Run as administrator` row says when the mark on the file did not follow it, in the
/// direction it was going.
///
/// **The two directions leave two different machines, and the box says which.** On, the task is in
/// place and only a launch by hand is left unelevated — the row half in effect. Off, the ordinary
/// entry is in place *beside* a mark that stops it, because Windows starts no `Run` value whose file
/// asks for elevation — the row's removal leaving a machine where nothing starts at a logon at all,
/// which is the one state a user has to be told about rather than left to find at the next start.
fn report_mark_failure(on: bool) {
    let text = match on {
        true => "The app will start as an administrator when you sign in, but starting it by hand \
                 will not ask for administrator rights.\n\n\
                 Windows would not accept the 'Run this program as an administrator' setting on the \
                 app file. You can set it yourself in the file's Properties, on the Compatibility \
                 tab.\n\n\
                 This usually happens because of a security policy or antivirus software. The rest \
                 of the app still works.",
        false => "The app still asks for administrator rights when you start it by hand.\n\n\
                  Windows would not remove the 'Run this program as an administrator' setting from \
                  the app file, and while that setting is there the app will not start when you \
                  sign in. You can remove it in the file's Properties, on the Compatibility tab, or \
                  by turning this setting on and off again.\n\n\
                  This usually happens because of a security policy or antivirus software. The rest \
                  of the app still works.",
    };

    report(text);
}

/// What the `Run as administrator` row says when the ordinary entry was written and the task that
/// starts this app elevated could not be taken away.
///
/// **A logon would start the app twice.** The value starts it unelevated and the task starts it
/// elevated, and which of the two is left standing is whatever Windows begins second — so the row
/// reads as off while a logon still starts an administrator's copy. Nothing else is written over:
/// a task this token may not delete is one the same row, run from an elevated copy of this app, can.
fn report_task_stayed() {
    report(
        "The app is set to start when you sign in, but the administrator startup could not be \
         removed.\n\n\
         Windows would now start the app twice when you sign in. To fix it, turn this setting on and \
         off again, or delete the 'RustHoverPreview' task in the Windows Task Scheduler.\n\n\
         This usually happens because of a security policy or antivirus software. The rest of the \
         app still works.",
    );
}

/// What the tray says when the copy it started to replace it never stood up.
///
/// **The row is in place; only this process's replacement is not.** The machine holds what the row
/// asked for — the task and the mark for an elevated start, the ordinary entry without them for an
/// unelevated one — and what did not happen is this tray being replaced by the copy that would run
/// under it. Leaving anyway would be the app gone with nothing to come back to, so this one stays
/// and says so: the same change is had by exiting and starting the app again.
pub(crate) fn report_restart_failure() {
    report(
        "The setting is saved, but the app could not restart itself.\n\n\
         Close the app from the tray menu and open it again to use the new setting.\n\n\
         This usually happens when Windows will not start a new copy of the app, for example because \
         of a security policy or antivirus software. Nothing else was changed.",
    );
}

/// A box that reports rather than asks, for what has already happened and cannot be refused.
///
/// It is a dialog of this app's like the questions are — the tray answers no menu while one is on
/// screen — and it is shown from the thread that needs it, which for an install is the worker's.
fn report(text: &str) {
    let caption = wide(CAPTION);
    let text = wide(text);

    begin();

    unsafe {
        MessageBoxW(
            HWND::default(),
            PCWSTR(text.as_ptr()),
            PCWSTR(caption.as_ptr()),
            MB_OK | MB_ICONWARNING | MB_SETFOREGROUND,
        );
    }

    end();
}

/// A string as a message box wants it: UTF-16, terminated.
fn wide(text: &str) -> Vec<u16> {
    text.encode_utf16().chain(std::iter::once(0)).collect()
}

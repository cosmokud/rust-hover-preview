//! The one row that asks the pin for something rather than writing a setting: the tray's
//! `Restore Pin Mode` (see `ID_TRAY_PIN_RESTORE`).
//!
//! A pin is the preview thread's — its window, the band it stands in and the caret it answers keys
//! with are all that thread's to ask — so the row sends the ask across the seam rather than reaching
//! into the window itself (see `PreviewMessage::PinRestore` and `pin_lifecycle::restore_the_pin_from_the_tray`).
//!
//! It is behind the same catch as every other row (`tray_window_proc`), and it writes no setting:
//! the row is a command, not a switch, which is why it hangs beside `Enable Preview` and `Pin Mode`
//! rather than inside either of them.

use windows::Win32::UI::WindowsAndMessaging::{AllowSetForegroundWindow, ASFW_ANY};

use crate::app::seam::{send_preview, PreviewMessage};

/// Bring the pin that is up back over everything it can be: out of its bubble or off the taskbar,
/// up from the state the Shell left it in, and in front with the keyboard.
///
/// **The foreground is granted before the ask crosses.** The window is the preview's, and the
/// preview is a process of its own whenever this run has a viewer (see `app::viewer`) — a process
/// that did not just receive a click, so Windows refuses it the foreground for asking (see
/// `pin_model::pin_take_focus`). What did receive the click is this one: the user answered a row of
/// this app's own menu, and the foreground the press earned is handed on where the restore will be
/// made. A refusal is still ordinary — the tray's own window may have lost the foreground between
/// the menu closing and the message landing — and the restore answers it with the band instead.
pub(in super::super) fn restore_pin_mode() {
    // Safety: asks the desktop to allow one process of this run to bring a window to the front,
    // which is the grant the press that opened the menu earned rather than a call into anything.
    unsafe {
        let _ = AllowSetForegroundWindow(ASFW_ANY);
    }

    send_preview(PreviewMessage::PinRestore);
}

//! The preview window's own procedure, and the keys it answers: a message read, a press
//! taken, and what each of them means before it reaches the tick.

use super::*;

pub(super) unsafe fn reset_preview_after_display_change(hwnd: HWND) {
    let _ = ShowWindow(hwnd, SW_HIDE);
    clear_pointer_hold();

    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        if let Some(ref mut media) = *current {
            media.cancel_background_work();
            stop_video_playback(media);
        }
        *current = None;
    }

    // A document the engine is playing is a window of its own, the way a video is, so
    // it comes down with the rest of the preview rather than being left where the
    // display it was placed for used to be. The hover is replayed and it is put up
    // again at the new one.
    webview_preview::hide();
}

/// Whether a key-down is Windows repeating a key that is already down rather than a second press
/// of it: the flag is bit 30 of a `WM_KEYDOWN` message's `lParam`, which is set on every
/// auto-repeat and clear on the first one (see the pinned window's own procedure).
pub(super) const KEY_REPEAT: isize = 1 << 30;

/// What one key pressed on a pinned window that has the focus means, as the command it asks for.
///
/// A key the pin is given the keyboard for is the pin's alone — it was not sent on to Explorer,
/// and nothing here is asked of the window in front of the pin. Which is the whole of what makes
/// an arrow a walk of the pin's own folder: while the pin is the window the user is in, there is
/// no listing in the keyboard for the arrow to move.
///
/// The keys here are the pin's own window answering as a window in the foreground does, and
/// that window is the gate: a key reaches it because Windows routed it there, which it does only
/// while the pin is the window the keyboard is in. There is no reading of the keyboard behind
/// this, and nothing to switch off — a pin the user has not clicked is a window nobody is in,
/// and a window nobody is in is sent no keystrokes at all.
///
/// A Space is the one key that is a play/pause, and it is a play/pause of whatever is playing:
/// a card is drawn with no buttons on it and a card with no way to hold a sound is a sound that
/// can only be listened to from beginning to end, so the key holds it; a video's bar already
/// offers the same hold under the picture, and the key is that hold asked by the keyboard rather
/// than by a press on the bar. Which of the two a Space is, is the file in the window's answer
/// rather than this mapping's (see `pin_toggle_target`). It used to hide or swap the window,
/// which read as a pin that got in the way of a Space typed anywhere near it; a key this window
/// is in front of is not a key to be dismissed with, and the pin is not dismissed by this one any
/// more. What a Space means to the file on screen is the loop's answer rather than this one's —
/// a picture, a page, and a sound at no volume at all are all swallowed (see
/// `pin_command_request`).
pub(super) fn pinned_key_command(vk: i32) -> Option<PinCommand> {
    // The keys are the virtual-key codes of the message's `wParam`, which are constants
    // rather than patterns, so the two directions are told apart by guards rather than by
    // arms: `VK_LEFT` and `VK_UP` both mean back, and `VK_RIGHT` and `VK_DOWN` both mean on.
    if vk == VK_LEFT.0 as i32 || vk == VK_UP.0 as i32 {
        return Some(PinCommand::Previous);
    }

    if vk == VK_RIGHT.0 as i32 || vk == VK_DOWN.0 as i32 {
        return Some(PinCommand::Next);
    }

    if vk == VK_ESCAPE.0 as i32 {
        return Some(PinCommand::Close);
    }

    if vk == VK_SPACE.0 as i32 {
        return Some(PinCommand::TogglePlayback);
    }

    // `T` is FFmpeg's own key for the next subtitle track, which is why it is this letter and
    // not one invented here — a user who has used the player elsewhere presses the key they
    // already know. It is answered as a relaunch rather than as a key, because the choice has to
    // outlive the relaunch (see `PinCommand::NextSubtitle`).
    //
    // A repeat is let through here rather than swallowed as a Space is, and the reason is that
    // every step of it is a relaunch and a relaunch takes long enough that a hand resting on
    // the key would otherwise do nothing visible: one press asks for one track.
    if vk == b'T' as i32 {
        return Some(PinCommand::NextSubtitle);
    }

    None
}

/// What a command the keyboard hook counted for a standing pin means, in the pin's own words.
///
/// The same answers `pinned_key_command` gives and in the same order, but read from the hook's
/// numbers rather than from this app's own mapping: a `WH_KEYBOARD_LL` callback may not take the
/// pin's lock to ask, so that mapping is a table of numbers over there (see
/// `key_input::PIN_KEY_COMMANDS`) and this is the one place that turns them back into commands.
/// The two must be kept in step — a key answered by one and not the other is a key this app acts
/// on and also passes on, which is a double action rather than a missing one.
pub(super) fn hook_pin_key_command(command: u8) -> Option<PinCommand> {
    match command {
        0 => Some(PinCommand::Previous),
        1 => Some(PinCommand::Next),
        2 => Some(PinCommand::Close),
        3 => Some(PinCommand::TogglePlayback),
        4 => Some(PinCommand::NextSubtitle),
        _ => None,
    }
}

/// What one key-down on a pinned window leaves for the loop, if anything.
///
/// It is the mapping above, and one more question: is this message Windows repeating a key that
/// is already down rather than pressing it again? A Space is the one command that cannot answer
/// one — a preview flickering between playing and held for as long as the hand is on the key —
/// while a walk is a thing a hand can sensibly hold down, so the arrows still repeat.
pub(super) fn pinned_key_down_command(vk: i32, lparam: isize) -> Option<PinCommand> {
    let command = pinned_key_command(vk)?;

    (!matches!(command, PinCommand::TogglePlayback) || lparam & KEY_REPEAT == 0).then_some(command)
}

/// What one key message on a pinned window leaves for the loop, if anything, with the class of the
/// message asked as part of the question.
///
/// A chord is a key the window eats without answering — it must not be closed by `Alt+F4` or left
/// by `Alt+Tab`, and none of them is a walk or a hold — and the keyboard hook, which stands in for
/// this window whenever the caret is in the player's rather than here, has to agree that it is a key
/// it too must not answer. The rule is therefore asked of the hook's one rather than repeated here,
/// because a repeat is where the two copies of one arrangement drift apart (see
/// `key_input::pin_key_message_acts`).
pub(super) fn pinned_key_message_command(
    message: u32,
    vk: i32,
    lparam: isize,
) -> Option<PinCommand> {
    if !crate::shell::key_input::pin_key_message_acts(message) {
        return None;
    }

    pinned_key_down_command(vk, lparam)
}

/// A mouse message's point in the coordinates the media of a pinned window is drawn in: the same
/// point, less the caption that has been added above it — and the same point exactly for a kind
/// whose chrome is drawn over its media, whose media begins at the window's own top (see
/// `pin_overlay_chrome`).
///
/// A hover's own window has no caption, so a point is already in its frame's coordinates and is
/// handed back unchanged — which is what lets one set of handlers serve both.
pub(super) fn media_point(x: i32, y: i32) -> (i32, i32) {
    // Nothing pinned leaves the point as it is, and that answer does not need the pin's lock:
    // this is asked on every mouse message the window is sent, and a hover's own window has no
    // caption (see `PIN_ACTIVE`).
    if !pinned() {
        return (x, y);
    }

    // A sound is given no caption (see `pinned_caption_height`), so its card begins at the
    // window's own top and a point on it is already where the card is drawn.
    let band = pin_state().and_then(|pinned| {
        let pin = pinned.pin()?;
        (!pin.collapsed && !pin.overlay).then_some(pin.caption)
    });

    match band {
        Some(band) => (x, y - band),
        None => (x, y),
    }
}

pub(super) unsafe extern "system" fn window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_DISPLAYCHANGE | WM_DPICHANGED => {
            reset_preview_after_display_change(hwnd);
            DISPLAY_RESET.store(true, Ordering::Release);
            LRESULT(0)
        }
        WM_PIN_RELEASE_POINTER => {
            // The watchdog gave up on a pin and cleared its state from outside the loop, and
            // this is the loop letting go of what its own window was still holding. The pin is
            // already gone by the time this arrives, so nothing here asks about one — the state
            // has been taken and there is no release coming to take the pointer back (see
            // `spawn_pin_watchdog`).
            release_pin_capture(hwnd);
            LRESULT(0)
        }
        WM_ACTIVATE => {
            // Windows is the only thing that can say which window the user is in, and it says it
            // here: a pin that has lost the focus is a pin the user has clicked away from, and
            // the keys belong to whatever is in front of it now. Nothing is asked of the pin on
            // the way out — a walk the user started with the pin in front and finished with it
            // behind is half a walk, not one.
            if wparam.0 as u32 == WA_INACTIVE {
                pin_release_focus();
            }
            LRESULT(0)
        }
        WM_KEYDOWN | WM_SYSKEYDOWN => {
            // A key that arrives here is one Windows decided belonged to this window, which is the
            // whole of the gate: while the pin is the window the user is in, an arrow walks the
            // pin's own folder and is not sent on to the listing behind, and while it is not, no
            // key arrives here at all and the arrow belongs to whatever the user is working in.
            //
            // **The class of the message is asked of the same rule the keyboard hook asks**, because
            // the hook stands in for this window whenever the caret is in the player's rather than
            // here (see `key_input::pin_key_message_acts`). One answer to a press rather than two: a
            // chord counted over there and swallowed here would walk this window's folder for a
            // `Ctrl`+Left this procedure answers with nothing but a swallow — and swallowing it is
            // also what keeps the system chords away from `DefWindowProcW`, which is where Alt+F4
            // turns into a close and Alt+Tab into an application switcher.
            //
            // A key this window *does* answer is left as the same command a caption button leaves,
            // rather than acted on here, because that is where the walk is answered (see `ask_pin`
            // and `pin_command_request`).
            if let Some(command) = pinned_key_message_command(msg, wparam.0 as i32, lparam.0) {
                ask_pin(command);
            }

            // Nothing is beeped at, and nothing is forwarded. A Space that reaches a pinned sound
            // is answered by the loop rather than here, and a Space that reaches a pin of any
            // other kind is answered by doing nothing at all, and letting `DefWindowProcW` ring
            // for the one key that is meant to be swallowed is a beep out of a window nobody can
            // see (see `pinned_key_command`).
            LRESULT(0)
        }
        WM_CLOSE | WM_SYSCOMMAND => {
            // The window is never to be destroyed. It is created once, for the life of the
            // process, and put up and taken down a thousand times over as previews come and go;
            // letting a close reach `DestroyWindow` would end every preview from here on with
            // nothing to report it, which is what Alt+F4 against a pin the user has clicked
            // would otherwise do.
            //
            // A close is answered the way a close button is — by ending the pin, and with it the
            // media under it and the claim on the keyboard. On a pin, a close is a close. On a
            // hover, of which there is no window to close and no claim to drop, it is nothing at
            // all, exactly as it was before this window could be activated (see
            // `pin_window::Reason::Closed`).
            if pinned() {
                request_pin_end();
            }
            LRESULT(0)
        }
        WM_CHAR => {
            // The character a key press has already been answered as. It is swallowed for the
            // same reason the key-down is: the pin has had the keyboard, and what it does with a
            // key is decided there — a character left to the default procedure would beep for
            // every key the pin answered silently, and would type into a caret there is none of.
            LRESULT(0)
        }
        WM_SETCURSOR => {
            // What the pointer is over on a pinned window, said with the pointer itself: an edge
            // of one is a resize, and a resize is the one thing a window has no other way of
            // announcing (see `pinned_set_cursor`). The point is not in this message: its `lParam`
            // is the hit-test code with the mouse message above it, so the cursor is asked where
            // it is — which is what the function below does.
            if pinned() && pinned_set_cursor(hwnd) {
                return LRESULT(1);
            }

            // And everywhere else it is the arrow, set here rather than left to whatever the
            // pointer was last told to be. A window that answers this message without setting a
            // cursor is a window that keeps the shape the last one left behind, which is how a
            // hover came to wear the diagonal that a pinned window's edge had put up.
            if let Ok(cursor) = LoadCursorW(None, IDC_ARROW) {
                SetCursor(cursor);
            }
            LRESULT(1)
        }
        WM_LBUTTONDOWN => {
            // A press on a pinned window is what makes it the window the user is in. A pin is
            // behind everything, so the press is the only thing that says the user means it rather
            // than the file behind it — and once the keyboard is here, the keys the pin answers are
            // keys the listing behind it is not sent (see `pin_take_focus`).
            if pinned() && !pin_is_focused() {
                pin_take_focus(hwnd);
            }

            // A press on a pinned window is the pin's before it is anything else's: the volume
            // popup over the picture, the caption's buttons, the caption itself, an edge, and the
            // media under the hand are all things a window does with a pointer (see `pinned_press`).
            let (x, y) = message_point(lparam);
            if pinned() && (pinned_volume_press(hwnd, x, y) || pinned_press(hwnd, x, y)) {
                return LRESULT(0);
            }

            // A press on the scrollbar starts a drag from where it landed, so the
            // thumb follows the pointer from the first click. Anywhere else, a
            // press on a text preview starts a selection. What the media's own handlers are
            // given is the point inside the frame, which for a pinned window begins below the
            // caption (see `media_point`).
            let (media_x, media_y) = media_point(x, y);
            if let Some(first_line) = text_scroll_drag_target(media_x, media_y) {
                set_text_scroll_dragging(true);
                let _ = SetCapture(hwnd);
                scroll_text_preview(hwnd, first_line);
            } else if begin_text_selection(media_x, media_y) {
                let _ = SetCapture(hwnd);
                repaint_text_preview(hwnd);
            }
            LRESULT(0)
        }
        WM_MOUSEMOVE => {
            let (x, y) = message_point(lparam);
            if pinned() {
                pinned_mouse_move(hwnd, x, y);
            }
            let (media_x, media_y) = media_point(x, y);
            if is_text_scroll_dragging() {
                if let Some(first_line) = drag_target_for_y(media_y) {
                    scroll_text_preview(hwnd, first_line);
                }
            } else if extend_text_selection(media_x, media_y) == Some(true) {
                repaint_text_preview(hwnd);
            }
            LRESULT(0)
        }
        WM_LBUTTONUP => {
            let (x, y) = message_point(lparam);
            if pinned() && pinned_release(hwnd, x, y, &Win32PinWindow) {
                return LRESULT(0);
            }
            if set_text_scroll_dragging(false) || end_text_selection() {
                let _ = ReleaseCapture();
            }
            // **The last resort, and the only writer on this road.** The arms above own the presses
            // they answer, and an arm with nothing armed does not touch the pointer at all — so
            // this is where the capture is given back, once, for a press no arm owns: the pin
            // rebuilt under the hand, the teardown, the watchdog, a press that reached a window
            // whose pin is gone. It runs after `pinned_release` has returned, which is what keeps
            // it from preempting an arm, and it is guarded by `GetCapture`, so it is a no-op
            // unless this window really is holding the pointer — and if it is, every mouse message
            // on the desktop is arriving here instead of at whatever it was aimed at, which is the
            // dead window a leaked capture looks like (see `release_pin_capture`).
            release_pin_capture(hwnd);
            LRESULT(0)
        }
        windows::Win32::UI::WindowsAndMessaging::WM_CAPTURECHANGED => {
            // Whatever a pinned window was doing with the pointer is over: a capture lost to
            // another window is a drag that is not coming back, and one this app released itself
            // is its own road's to finish (see `pin_capture_lost`).
            pin_capture_lost(&Win32PinWindow);
            LRESULT(0)
        }
        WM_RBUTTONUP => {
            // The menu is the preview's own, so it opens where it was asked for,
            // and the press that asks for it does not dismiss the preview.
            let (x, y) = message_point(lparam);
            if TEXT_PREVIEW_HOLDING.load(Ordering::Acquire) {
                show_text_preview_menu(hwnd, x, y);
            }
            LRESULT(0)
        }
        WM_POWERBROADCAST => {
            let power_event = wparam.0 as u32;
            match power_event {
                PBT_APMRESUMEAUTOMATIC | PBT_APMRESUMESUSPEND => {
                    // System resumed from sleep — DWM has restarted and the
                    // layered window's composition surface was destroyed.
                    // Reset the preview so the next hover creates everything fresh.
                    reset_preview_after_display_change(hwnd);
                    RESUME_FROM_SLEEP.store(true, Ordering::Release);
                }
                PBT_APMSUSPEND | PBT_APMSTANDBY => {
                    // System is going to sleep. Clean up video playback and
                    // background decoding to avoid resource leaks.
                    reset_preview_after_display_change(hwnd);
                }
                _ => {}
            }

            // A pinned preview does not survive either half of it: a suspension has already
            // taken down every surface and every session the pin was a window onto, and what it
            // holds is not something a resume can put back. The hook is told, and the loop takes
            // the pin down on its next tick, which is the same path its close button takes.
            if pinned() {
                request_pin_end();
            }
            LRESULT(0)
        }
        windows::Win32::UI::WindowsAndMessaging::WM_PAINT => {
            let mut ps = PAINTSTRUCT::default();
            let _ = BeginPaint(hwnd, &mut ps);
            render_layered_preview(hwnd);
            let _ = EndPaint(hwnd, &ps);
            LRESULT(0)
        }
        windows::Win32::UI::WindowsAndMessaging::WM_DESTROY => LRESULT(0),
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

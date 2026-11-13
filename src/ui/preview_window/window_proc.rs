//! The preview window's own procedure, and the keys it answers: a message read, a press
//! taken, and what each of them means before it reaches the tick.

use super::pin_trace::note_pin_trace;
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
    // arms: `VK_LEFT` is back and `VK_RIGHT` is on, and the two the page shares with them are
    // their own commands because a page has somewhere for them to go (see `PinCommand::PageBack`).
    if vk == VK_LEFT.0 as i32 {
        return Some(PinCommand::Previous);
    }

    if vk == VK_RIGHT.0 as i32 {
        return Some(PinCommand::Next);
    }

    // An up arrow is a line back where the window holds a page of text and the previous file where
    // it holds anything else, which is the same question a Space is: what the key means is the file
    // in the window's answer rather than this mapping's (see `pin_command_request`). A `PageUp` is
    // the same walk for a window with no page and a screenful for one with a page, and it is its own
    // command for exactly that reason: the two used to be one, and a caret needs them apart (see
    // `PinCommand::CaretUp`).
    if vk == VK_UP.0 as i32 {
        return Some(PinCommand::CaretUp);
    }

    if vk == VK_PRIOR.0 as i32 {
        return Some(PinCommand::PageBack);
    }

    if vk == VK_DOWN.0 as i32 {
        return Some(PinCommand::CaretDown);
    }

    if vk == VK_NEXT.0 as i32 {
        return Some(PinCommand::PageOn);
    }

    // The two ends of the pin's own list, and of a page's line where the page has the reader's focus
    // and of its document under `Ctrl` — which of the two a key is is the file in the window's answer
    // and the caret's, asked where both are known (see `pin_command_request` and
    // `text_caret::page_has_focus`).
    if vk == VK_HOME.0 as i32 {
        return Some(PinCommand::PageTop);
    }

    if vk == VK_END.0 as i32 {
        return Some(PinCommand::PageBottom);
    }

    // `F3` is the key the system has used for "find the next one" since before this app: a user who
    // has looked for a word in anything else presses the key they already know (see `PinCommand::FindNext`).
    if vk == VK_F3.0 as i32 {
        return Some(PinCommand::FindNext);
    }

    if vk == VK_ESCAPE.0 as i32 {
        return Some(PinCommand::Close);
    }

    if vk == VK_SPACE.0 as i32 {
        return Some(PinCommand::TogglePlayback);
    }

    // `Return` is a list's own key wherever a hand already is in the list, and a sound's card with its
    // band up is a list the hand is inside: the row it has picked is the row it plays. Nothing else on
    // a pin is a list, so a `Return` elsewhere is answered by the command and not by the mapping — a
    // band that is down or holds no pick has nothing to play (see `PinCommand::PlayRow`).
    if vk == VK_RETURN.0 as i32 {
        return Some(PinCommand::PlayRow);
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
        5 => Some(PinCommand::CaretUp),
        6 => Some(PinCommand::PageBack),
        7 => Some(PinCommand::CaretDown),
        8 => Some(PinCommand::PageOn),
        9 => Some(PinCommand::PageTop),
        10 => Some(PinCommand::PageBottom),
        11 => Some(PinCommand::FindNext),
        12 => Some(PinCommand::PlayRow),
        _ => None,
    }
}

/// What a Win chord the keyboard hook counted for a standing pin means, in the pin's own words.
///
/// The same one table `key_input::PinWinChord` counts into, read from its numbers in the order
/// `app::seam::PIN_WIN_CHORDS` lists them — and the same rule the key table beside it keeps: the two must
/// be kept in step, because a chord counted into a number nothing here names is a key swallowed for
/// nothing at all.
pub(super) fn hook_pin_win_chord(chord: u8) -> Option<PinCommand> {
    match chord {
        0 => Some(PinCommand::Minimize),
        1 => Some(PinCommand::MoveToDisplay(PinSide::Left)),
        2 => Some(PinCommand::MoveToDisplay(PinSide::Right)),
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
///
/// **A Win chord asks the pin nothing at all**, and that is the hook's rule as well: the pin has its own
/// arrangement for one — its minimize, and a window carried to the display beside this one — and the rest
/// are the Shell's, not something to walk a folder with (see `key_input::pin_win_chord`). The hook takes
/// them wherever it stands, and this is where one that reached the window anyway is answered: with nothing.
pub(super) fn pinned_key_message_command(
    message: u32,
    vk: i32,
    lparam: isize,
    win_down: bool,
) -> Option<PinCommand> {
    if !crate::app::seam::pin_key_message_acts(message) {
        return None;
    }

    if crate::shell::key_input::a_win_chord_takes(vk, win_down) {
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

/// Answer a close asked of one of this app's windows: the Shell's **Close** on its taskbar button, its
/// group's **Close all windows**, and `Alt`+`F4` where the key is not swallowed.
///
/// **The tray's ask is a different road, and this one exists because it is.** `request_pin_end` is the tray
/// speaking to the pin it stood over, and a window of its own spends that without acting on it (see
/// `PIN_END_REQUESTED`) — so a close pressed on the window standing on the screen would be swallowed by the
/// very window it was pressed on. A close of the window's own is answered as the pin's own close instead,
/// the road its caption's cross takes, which is what a window of its own needs: its window *is* it, and the
/// loop's own take-down is what carries the process away with the window (see `detach_open::settle`).
pub(super) fn close_asked_of_the_window() {
    if crate::app::detach::is_detached() {
        ask_pin(PinCommand::Close);
        return;
    }

    request_pin_end();
}

pub(super) unsafe extern "system" fn window_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    // **This window is the one the Shell may be answering a click with instead of the film host**, so
    // the same set of messages is written down for it — and only while a pin stands, because this is
    // the hover's own procedure as well and a hover is nobody's button (see
    // `pin_trace::note_a_message_of_a_window`). Nothing is answered differently for the line.
    if pinned() {
        pin_trace::note_a_message_of_a_window("pin", hwnd, msg, wparam, lparam);
    }

    // **A press on the pair's button is the Shell taking this app away and giving it straight back**, and
    // this is where the first half of that blink is noted: the chrome is the window the Shell drives when
    // it cannot put the host down, because it is the window already in front (see
    // `pin_film_host::note_the_app_was_taken_away_while_the_pin_held_the_keyboard`). The note is what
    // tells that click from every activation this app makes of itself, and it carries the caret too: only
    // a taking made while the pin holds the keyboard is noted, because by the time the activation arrives
    // Windows has taken the caret away with the app and not yet given it back. **And the thread the message
    // names is asked as well**, because `WM_ACTIVATEAPP`'s `lParam` is the other side of the switch: a
    // taking the taskbar did is a click on its button, and a taking any other program did — an Alt-Tab — is
    // that program's hand on the app and no click at all.
    pin_film_host::note_the_app_was_taken_away_while_the_pin_held_the_keyboard(msg, wparam, lparam);

    match msg {
        WM_DISPLAYCHANGE | WM_DPICHANGED => {
            // A pin is not replayed from the pointer and its media is not something the pointer
            // can be asked for again, so it survives the display it crossed to: nothing is put
            // away, nothing is stopped and nothing is hidden, and the tick re-boxes it for the
            // display it landed on (see `replace_pinned_window`). Hover is put away and drawn
            // again, as it always has been.
            if !pinned() {
                reset_preview_after_display_change(hwnd);
            }
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
        WM_DROPFILES => {
            // **A file dragged onto this window is a file the pin is asked to show.** This procedure
            // is the hover's, the pin's own window's and the bubble's, so the answer is the pin's or
            // nobody's: a hover is not a window a file is dropped onto and nothing of a drop before
            // there is a pin is anyone's to answer (see `pin_drop`).
            //
            // The list is read and given back here rather than kept, because the handle belongs to
            // this thread and the tick that plans a pick is the loop's own: what crosses between the
            // two is a path and nothing else (see `pin_drop::take_drop`).
            //
            // **A drop changes nothing here, and the loop is what changes it.** The window is the
            // furthest this procedure may go: a pick is planned, loaded and installed off a tick that
            // waits in `MsgWaitForMultipleObjectsEx`, so the loop is woken the way every other ask
            // made from a message is (see `crate::app::seam::wake_preview_loop`).
            //
            // Nothing is passed on to the default procedure either: a `WM_DROPFILES` left to it is a
            // drop nobody answers, and the list it carries is the Shell's to finish (see `DragFinish`).
            if pinned() {
                let path = pin_drop::take_drop(HDROP(wparam.0 as _));
                if let Some(path) = path {
                    if let Ok(mut slot) = pin_drop::PIN_DROP.lock() {
                        *slot = Some(pin_drop::DroppedOn {
                            path,
                            put_away: pin_is_collapsed(),
                        });
                    }
                    crate::app::seam::wake_preview_loop();
                }
            }
            LRESULT(0)
        }
        WM_ACTIVATE => {
            // **A press on the pair's taskbar button while the pin is in front is this pin going away**,
            // and on the road the Shell takes for it the click arrives at this window rather than at the
            // host the button belongs to (see `pin_film_host::answer_a_click_on_the_chromes_button`).
            // Asked before the arm's own work: a deactivation is not an activation, so it falls through to
            // what this arm has always answered it with.
            if let Some(answer) = pin_film_host::answer_a_click_on_the_chromes_button(msg, wparam) {
                return answer;
            }

            // Windows is the only thing that can say which window the user is in, and it says it
            // here: a pin that has lost the focus is a pin the user has clicked away from, and
            // the keys belong to whatever is in front of it now. Nothing is asked of the pin on
            // the way out — a walk the user started with the pin in front and finished with it
            // behind is half a walk, not one.
            if wparam.0 as u32 == WA_INACTIVE {
                pin_release_focus();
                // Whatever is in front now was clicked into rather than
                // pressed here, so a menu left up over the card goes away
                // with the activation: the capture path answers presses, and
                // this answers the activation change that came without one.
                close_pin_menu(hwnd);
            }

            // **Alt+Tab, the taskbar button and a click on the chrome all activate this window
            // rather than the page**, and the page is a window of its own that Windows does not
            // bring along. It is placed again last, which shows it and puts it above the pin the
            // activation has just raised (see `place_pinned_siblings` and `Host::put_window`).
            if matches!((wparam.0 & 0xFFFF) as u32, WA_ACTIVE | WA_CLICKACTIVE)
                && engine_page_in_the_band().is_some()
            {
                place_pinned_siblings();
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
            // A key that reaches the window puts the card's menu away before it is answered: the
            // panel is drawn over the card, and a hand that has gone back to the keyboard has left
            // it. Escape — a close everywhere else — is the menu's while it is up, as it is the
            // bubble's own, so it puts the panel away rather than the pin under it (see
            // `close_pin_menu` and `pin_bubble`).
            let menu_was_up = close_pin_menu(hwnd);
            if menu_was_up && wparam.0 as i32 == VK_ESCAPE.0 as i32 {
                return LRESULT(0);
            }

            // **A find bar that is up owns the keyboard, and this arm is the only owner there is.**
            // The keyboard hook stands down from the pin's keys for as long as a bar is up (see
            // `key_input::publish_find_bar_up`), because a key it swallowed would never become the
            // `WM_CHAR` of the arm below — the characters a query is typed from. So every key
            // arrives here, and the bar's one table reads the ones it owns: a return, a backspace,
            // the two arrows and the two page keys, the two ends, `Esc` and `F3` (see
            // `text_find::key_edit_with`). Every other key is swallowed whole — a `Space`, a `T` and
            // a `Tab` mean nothing to a query — and none of them may reach the page's own table
            // below or `DefWindowProcW`'s.
            if text_find::is_open() {
                if let Some(edit) = text_find::key_edit(wparam.0 as i32) {
                    text_find::note(edit);
                }
                return LRESULT(0);
            }

            // **And the band's own search field is the second text field a pin carries.** It is asked
            // of the same table for the same reason, and it answers the keys it names — a backspace, a
            // delete, the two arrows, the two ends, `Esc` and `Enter` — while a key it does not name is
            // the pin's, so the arrows that step the list still do (see
            // `pin_playlist::key_while_the_search_has_focus`). The characters are the `WM_CHAR` below,
            // which exists only because the hook stood down for them (see
            // `key_input::publish_find_bar_up` and `app::seam::PIN_PLAYLIST_SEARCH_UP`).
            if pin_playlist::key_while_the_search_has_focus(wparam.0 as i32) {
                render_layered_preview(hwnd);
                return LRESULT(0);
            }

            // **`Ctrl`+`F` over a band that is up is the band's search field.** The chord is answered
            // here rather than as a command of the pin's, because it is a caret moving into a field and
            // not a walk of the folder: the command road ends in a `PinCommand`, and none of the pin's
            // commands names a caret. It is swallowed whole, so the character the loop would have
            // translated from it never reaches the field it has just focused.
            if wparam.0 as i32 == VK_F.0 as i32
                && text_selection::control_is_down()
                && pin_playlist::focus_the_search()
            {
                render_layered_preview(hwnd);
                return LRESULT(0);
            }

            if let Some(command) = pinned_key_message_command(
                msg,
                wparam.0 as i32,
                lparam.0,
                crate::shell::explorer_hook::a_win_key_is_down(),
            ) {
                ask_pin(command);
            }

            // Nothing is beeped at, and nothing is forwarded. A Space that reaches a pinned sound
            // is answered by the loop rather than here, and a Space that reaches a pin of any
            // other kind is answered by doing nothing at all, and letting `DefWindowProcW` ring
            // for the one key that is meant to be swallowed is a beep out of a window nobody can
            // see (see `pinned_key_command`).
            LRESULT(0)
        }
        WM_CLOSE => {
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
                close_asked_of_the_window();
            }
            LRESULT(0)
        }
        WM_SYSCOMMAND => {
            // **Not every system command is a close**, and answering them all as one is what made
            // `Win`+`D` end Pin Mode. Show Desktop is the Shell minimizing every window that offers
            // a minimize box, and that ask arrives here as `SC_MINIMIZE` on the pin's own window: the
            // gesture meant to clear the screen took the pin away with it, though nothing had closed
            // and nothing had been collapsed (see `pin_real_window` for the styles that make the
            // window one Show Desktop takes).
            //
            // Only `SC_CLOSE` is a close: the window's own button, and the row the Shell's menu
            // beside its taskbar button puts on it. The rest of what a system command can be —
            // minimizing with the desktop and coming back with it above all — is the desktop's
            // business with the window and belongs to `DefWindowProcW`, the road every other window
            // in the system takes for it.
            //
            // A stand-in of `pin_taskbar`'s window answers those two commands itself, because that
            // window is not the pin and has nothing to minimize: the pin is put away into its bubble
            // instead (see `handle_taskbar_restore`).
            // The low four bits of a system command belong to the application that sends one, so the
            // command is the word without them — the mask `pin_taskbar`'s own procedure reads its two
            // with.
            if (wparam.0 & 0xFFF0) as u32 == SC_CLOSE {
                if pinned() {
                    close_asked_of_the_window();
                }
                return LRESULT(0);
            }

            // **A minimize of this window is the pair being put down, where this window is the pair's
            // second one.** A film pin's chrome is unowned and not a tool window, so the Shell
            // sometimes lists it beside its film host and a real click has been seen deactivating and
            // re-activating it. Whichever window of the pair the Shell puts down, the pin's state and
            // the collapse have to be the same, so the command is queued and `DefWindowProcW` is not
            // reached: a bare minimize here would take the window down behind a pin that still
            // believes it is up (see `pin_film_host::the_shell_put_the_chrome_down`).
            if pin_film_host::the_shell_put_the_chrome_down(msg, wparam) {
                note_pin_trace!(
                    "pin decision the shell's minimize of the chrome: the taskbar minimize is queued"
                );
                ask_pin(PinCommand::MinimizeTo(PinMinimizeTo::Taskbar));
                return LRESULT(0);
            }

            // **A maximize or a restore asked of this window is the pin's own box, not the
            // window's.** A double-click on the caption arrives here as `SC_MAXIMIZE`, and the box
            // `DefWindowProcW` would give the window natively is one the pin's state never hears
            // about: no `pin.restore` to come back from and no `pin.bound` for the files that follow
            // (see `toggle_pin_maximized`). So each is answered by the pin's own toggle — one
            // command, because the state it pivots on is the same one either way — and the command
            // is not passed on, a native box being a box nothing else records. Every other system
            // command is still the desktop's business and reaches the default procedure below.
            let maximized = pin_state()
                .and_then(|pinned| pinned.pin().map(|pin| pin.restore.is_some()))
                .unwrap_or(false);
            match (wparam.0 & 0xFFF0) as u32 {
                SC_MAXIMIZE if !maximized => {
                    note_pin_trace!(
                        "pin decision the window was asked to maximize: the pin's own maximize is queued"
                    );
                    ask_pin(PinCommand::Maximize);
                    return LRESULT(0);
                }
                SC_RESTORE if maximized => {
                    note_pin_trace!(
                        "pin decision the window was asked to restore: the pin's own restore is queued"
                    );
                    ask_pin(PinCommand::Maximize);
                    return LRESULT(0);
                }
                _ => {}
            }

            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_SIZE => {
            // **A pin put away as a window comes back as a window, and the pin's state has to follow it.**
            // The gate and the ask are one implementation shared with the film host that owns the pin's
            // chrome on the other road a pin is put away as a window (see `the_pin_came_back`).
            the_pin_came_back(hwnd, msg, wparam);

            // **And the size a minimize reports is the same put-away the command above is**, for the
            // same reason: the Shell may be putting this window down rather than the host it is owned
            // by, and a pin whose chrome went down believing it was up is a pin that answers no
            // minimize and holds everything a collapse holds back (see
            // `pin_film_host::the_shell_put_the_chrome_down`).
            if pin_film_host::the_shell_put_the_chrome_down(msg, wparam) {
                note_pin_trace!(
                    "pin decision the shell's minimize of the chrome: the taskbar minimize is queued"
                );
                ask_pin(PinCommand::MinimizeTo(PinMinimizeTo::Taskbar));
                return LRESULT(0);
            }

            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_MOVE | WM_WINDOWPOSCHANGED => {
            // **The pin's own motion is the only road the caption's popup can follow it on**: a window
            // that moves repaints nothing, so nothing of the paint road runs and the popup would be
            // left standing where the pin was. A maximize, a restore and a minimize answer here too —
            // and a pin that cannot show its popup hides it rather than leaving it on the desktop (see
            // `follow_the_pin`).
            pin_chrome_strips::follow_the_pin(hwnd);
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_CHAR => {
            // **A character is the find bar's while it is up.** It arrives here as one UTF-16 code
            // unit because the loop translates the key before dispatching it, and the bar is the
            // one caret a pinned page has (see `text_find::typed_character`). The page's own key
            // table is below this arm, so nothing of a typed character reaches it.
            if text_find::is_open() {
                if let Some(character) = text_find::typed_character(wparam.0) {
                    text_find::note(FindEdit::Typed(character));
                }
                return LRESULT(0);
            }

            // And a character is the band's search field's while that has the caret: the field is the
            // second text field this window carries, and its query is typed from the same messages the
            // page's is — dropped where the code unit is half a character the field does not hold (see
            // `pin_playlist::type_into_the_search`).
            if pin_playlist::search_has_focus() {
                if let Some(character) = text_find::typed_character(wparam.0) {
                    pin_playlist::type_into_the_search(character);
                    render_layered_preview(hwnd);
                }
                return LRESULT(0);
            }

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
            let (x, y) = message_point(lparam);
            // The point on the screen the press landed at, which is what the card's menu is
            // asked in: its panels are drawn in a window of their own over the screen rather
            // than inside this one (see `pin_menu_window`).
            let (screen_x, screen_y) = client_to_screen(hwnd, x, y);
            // A captured press aimed elsewhere — another app, the taskbar —
            // while the card's menu is up: the menu goes away with it, and the
            // press goes no further. No focus is taken for it and no drag is
            // begun from it: the hand is on another window, and answering it
            // as a press on the card would carry this window for that hand.
            if pin_menu_is_open() && pin_menu_point_is_outside(screen_x, screen_y) {
                close_pin_menu(hwnd);
                return LRESULT(0);
            }

            // A press on a pinned window is what makes it the window the user is in. A pin is
            // behind everything, so the press is the only thing that says the user means it rather
            // than the file behind it — and once the keyboard is here, the keys the pin answers are
            // keys the listing behind it is not sent (see `pin_take_focus`).
            //
            // **And the press is answered before the keyboard is asked for.** A press on a pinned
            // window is the pin's before it is anything else's: the volume popup over the picture,
            // the caption's buttons, the caption itself, an edge, and the media under the hand are
            // all things a window does with a pointer — and the drag or the resize one of them
            // begins is the pointer taken for this window. That has to happen before any of the
            // window-manager work the keyboard costs, or the first press of a hand that is already
            // moving has nothing to carry it (see `press_the_pin`).
            // **What the press was answered with, written down.** Every road a press can take on a
            // pin is silent — a caption button, a panel, the transport bar and a drag all look the
            // same from outside the window — and a press that carries nothing is the one fault no
            // other line in this trace can name: the band a drag asks for on its way in is the only
            // thing `begin_pin_drag` writes, so a press with no band line after it is a press some
            // arm above the media took. The point, the chrome it was read against, the answer and
            // what the pin is holding afterwards are the whole of that answer.
            let (answered, dragging, chrome) = {
                let answered = pinned()
                    && unsafe { press_the_pin(hwnd, x, y, &Win32PinWindow, !pin_is_focused()) };
                let (dragging, chrome) = pin_state()
                    .and_then(|pinned| {
                        pinned.pin().map(|pin| {
                            (pin.dragging.is_some(), (pin.chrome.caption, pin.chrome.bar))
                        })
                    })
                    .unwrap_or((false, (false, false)));
                (answered, dragging, chrome)
            };
            note_pin_trace!(
                "pin press {},{} pinned {} answered {} dragging {} chrome {}/{}",
                x,
                y,
                pinned() as u8,
                answered as u8,
                dragging as u8,
                chrome.0 as u8,
                chrome.1 as u8,
            );
            if answered {
                return LRESULT(0);
            }

            // A find bar that is up is the first thing a hand on the page can land on: the card floats
            // over the page's top-right corner, so a press inside it is the card's rather than the
            // page's, and nothing under it was what the hand meant (see `text_find::press`).
            if text_find::press(x, y) {
                // A press on one of the card's controls is held until it is let go of, and the pointer
                // is taken for as long as that lasts: a hand that slides off the card still lets go of
                // the button it pressed, rather than leaving it lit under a pointer that is somewhere
                // else (see `text_find::release`). A press in the query is held for the same reason: a
                // drag grows the selection and the release is what ends it (see `text_find::drag`).
                if text_find::a_control_is_held() {
                    let _ = SetCapture(hwnd);
                }
                // A press in the query has moved the caret without asking for a search, so nothing
                // else on this tick is going to paint it: the card is drawn again here rather than
                // left showing the caret it had before the hand landed.
                render_layered_preview(hwnd);
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
            // The rest of a drag a model's band owns: the window that took the down posts the moves, and
            // ends it on the level of the button where the release was the engine's (see
            // `f3d_player::answer_a_models_band`).
            if pinned() {
                if let Some(answer) =
                    unsafe { f3d_player::answer_a_models_band(hwnd, msg, wparam, lparam) }
                {
                    return answer;
                }
            }
            let (x, y) = message_point(lparam);
            if pinned() {
                pinned_mouse_move(hwnd, x, y);
            }
            // A drag in the find bar's query is answered before the page's own selection, and only
            // where the card really took the press: the selection grows to the place under the
            // pointer, and the card is drawn again only where that place changed (see
            // `text_find::drag`).
            if text_find::drag(x) {
                render_layered_preview(hwnd);
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
            // The release that ends a model's drag is posted on and answered here, whatever window holds
            // the pointer, so the pin's own arms never see the end of a press it did not take (see
            // `f3d_player::answer_a_models_band`).
            if pinned() {
                if let Some(answer) =
                    unsafe { f3d_player::answer_a_models_band(hwnd, msg, wparam, lparam) }
                {
                    return answer;
                }
            }
            let (x, y) = message_point(lparam);
            // A press the find bar's card answered is released here, before the pin's own arms: the
            // card owns the press from the moment it takes it, and a chevron held down stops where it
            // got to. The capture the press took is given back here rather than left to the arm below,
            // which this answer returns before — the same `GetCapture` guard, for the same reason (see
            // `release_pin_capture`).
            if text_find::release() {
                release_pin_capture(hwnd);
                render_layered_preview(hwnd);
                return LRESULT(0);
            }
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
            // A press the find bar's card took is over when the pointer is taken away from this window:
            // a chevron that never sees its release would step for as long as the bar is up, a drag
            // would go on growing a selection with nobody holding it, and the card would draw a button
            // held under a hand that is no longer on it (see `text_find::release`).
            if text_find::release() {
                render_layered_preview(hwnd);
            }
            // **Who took the pointer, while a hand still held it**: a drag that neither moved nor
            // clicked is a drag nobody's messages were carrying, and the window named here is the one
            // that took them (see `pin_capture_lost`). `lParam` is the window gaining the capture.
            if pinned() && pin_is_dragging() {
                note_pin_trace!(
                    "the pointer was taken from the pin while a drag was in hand, by {:?} (button down {})",
                    HWND(lparam.0 as *mut _),
                    left_button_down() as u8
                );
            }
            // Whatever a pinned window was doing with the pointer is over: a capture lost to
            // another window is a drag that is not coming back, and one this app released itself
            // is its own road's to finish (see `pin_capture_lost`).
            pin_capture_lost(&Win32PinWindow);
            LRESULT(0)
        }
        WM_RBUTTONDOWN => {
            // A right-click on a pinned window asks for the window's own menu, at
            // the point it was asked for: a sound's card carries the card's own rows
            // and every other kind the rows a bubble's menu holds (see
            // `pinned_menu_paint`). The menu is the preview's own, so the press does
            // not dismiss the preview; a right-click on anything else is left to the
            // default procedure — a hover's window has no menu, and the text
            // preview's own opens on the release below.
            if pinned() {
                let (x, y) = message_point(lparam);
                // The point on the screen, which is the window the card's menu's panels are
                // drawn in and the point one is asked of at.
                let (menu_x, menu_y) = client_to_screen(hwnd, x, y);
                // A captured right-click aimed elsewhere while the menu is up
                // puts it away rather than moving it: the panel belongs to the
                // screen rather than to this window, and a point it holds no
                // panel at is a hand that has left.
                if pin_menu_is_open() && pin_menu_point_is_outside(menu_x, menu_y) {
                    close_pin_menu(hwnd);
                    return LRESULT(0);
                }
                // And a row the menu is not for — a film's transport bar — is left under the hand
                // rather than answered here: what is there is the player's own strip (see
                // `pin_right_click_page`). A right-click in the text of a pinned page is the
                // window's own, with the page's two commands above its rows.
                //
                // **A model's band is neither**: it is f3d's pointer, so the press, the drag and
                // the release are the engine's and this window holds no menu there at all (see
                // `pin_band_is_the_models`). What keeps the menu reachable is the caption above
                // the picture, which is the chrome's and answers below.
                if pin_band_is_the_models(y) {
                    unsafe { f3d_player::press_the_models_band_with(hwnd, x, y, msg) };
                    return LRESULT(0);
                }
                if let Some(page) = pin_right_click_page(y) {
                    open_pin_menu(hwnd, menu_x, menu_y, page);
                }
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_MBUTTONDOWN => {
            // A middle press while the menu is up puts it away wherever it
            // landed: the menu is answered press by press, and no middle press
            // is a row's own.
            if pin_menu_is_open() {
                close_pin_menu(hwnd);
                return LRESULT(0);
            }
            // **And over a model's band it is the engine's**, as every press there is: f3d's own
            // style pans by it, and this app answers nothing for a middle press on a picture but
            // this (see `pin_band_is_the_models`).
            if pinned() {
                let (x, y) = message_point(lparam);
                if pin_band_is_the_models(y) {
                    unsafe { f3d_player::press_the_models_band_with(hwnd, x, y, msg) };
                    return LRESULT(0);
                }
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_MBUTTONUP => {
            // The release that ends a middle press on a model's band, on the same road a left one
            // ends (see `f3d_player::answer_a_models_band`).
            if pinned() {
                if let Some(answer) =
                    unsafe { f3d_player::answer_a_models_band(hwnd, msg, wparam, lparam) }
                {
                    return answer;
                }
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_RBUTTONUP => {
            // The release that ends a right press on a model's band, which is the engine's own
            // menu and its own drag rather than this window's (see
            // `f3d_player::answer_a_models_band`).
            if pinned() {
                if let Some(answer) =
                    unsafe { f3d_player::answer_a_models_band(hwnd, msg, wparam, lparam) }
                {
                    return answer;
                }
            }
            // The menu is the preview's own, so it opens where it was asked for,
            // and the press that asks for it does not dismiss the preview.
            let (x, y) = message_point(lparam);
            // A page's own `Select All` / `Copy` menu is for a text preview with no window
            // around it, and opens on the release rather than the press: a pinned page asks
            // the window's own menu for the same rows and the ones under them, and that menu
            // is already up from the press (see `pin_right_click_page`).
            if TEXT_PREVIEW_HOLDING.load(Ordering::Acquire) && !pinned() {
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
        // **The Shell's ask for the still of this window's taskbar button, while the window is down.** A
        // pin that stands in a window of its own wears its own button, and a button whose window has been
        // put away has no composition for the Shell to freeze — the picture the collapse took is what is
        // handed over instead (see `pin_taskbar_mirror::answer_a_thumbnail_ask`). The same procedure is the
        // hover's, and a hover answers nothing here: no picture is ever held for one, and an ask with
        // nothing to hand it leaves the Shell its own fallback rather than anything this app draws.
        windows::Win32::UI::WindowsAndMessaging::WM_DWMSENDICONICTHUMBNAIL => {
            pin_taskbar_mirror::answer_a_thumbnail_ask(hwnd, lparam)
        }
        windows::Win32::UI::WindowsAndMessaging::WM_DESTROY => LRESULT(0),
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

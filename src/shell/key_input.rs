//! System-wide keyboard watcher for the pin key, and for the keys a standing pin answers.
//!
//! A low-level keyboard hook counts the pin key wherever the pointer happens to be,
//! and publishes a monotonic press counter the preview loop consumes, exactly as
//! the wheel watcher beside it publishes wheel ticks. That key is never swallowed:
//! `Space` is Explorer's own, and what the loop does with the count is its business
//! (see `PIN_PRESSES`). Only the key that puts a pin up, and the one that brings a
//! collapsed one back, is answered here.
//!
//! The keys a pin already up answers — the arrows, `Esc`, `Space`, `T` — used to arrive as
//! messages to the pin's own window, and are answered here as well, and *are* swallowed. That is a
//! reversal of the rule above, and it is deliberate. The pin's window eats all four whenever it
//! holds the caret (`WM_KEYDOWN` returns without reaching `DefWindowProcW`), so swallowing them
//! costs a key nothing; what it buys is that they still work when it does not hold the caret, and
//! a pinned video is exactly when that happens. Stepping to the next file begins a new player
//! (`restart_pinned_player`), that player's window exists for a moment before
//! `WS_EX_NOACTIVATE` is applied to it (`apply_noactivate_to_hwnd`), and if it takes the caret in
//! that window the arrows stop being navigation and become FFmpeg's own: left and right seek by
//! ten seconds, up and down by a minute, and a user pressing "next file" twice gets a film
//! scrubbing back and forth over itself instead. The swallow is what stops FFmpeg ever seeing
//! them.
//!
//! It is gated on the caret being in one of *this pin's* two windows, so a key typed into whatever
//! is behind the pin is still that window's key — the hook is not a global remap, and a pin up
//! over Explorer does not steal the arrows from Explorer.
//!
//! **A find bar is the one state in which none of that happens.** A bar floats over a pinned page and
//! the query it holds is a text field, and a field's characters are the `WM_CHAR` that
//! `TranslateMessage` makes *after* every low-level hook has returned: a `Space` or a `T` swallowed
//! here is a character the query can never hold. So while a bar is up over the focused pin, that
//! pin's keys are not taken at all — nothing is counted and nothing swallowed, and the pin's own
//! window answers them, which is where the bar's one key table lives (see `PIN_FIND_BAR_UP` and
//! `text_find::key_edit_with`).

#[cfg(test)]
use crate::app::seam::PIN_KEY_COUNTED;
use crate::app::seam::{
    PIN_FIND_BAR_UP, PIN_KEYS_OWNER, PIN_KEYS_TOPMOST, PIN_KEYS_VIDEO, PIN_KEY_PRESSES,
};
use crate::CONFIG;
use std::sync::atomic::{AtomicBool, AtomicU32, AtomicU64, Ordering};
use windows::Win32::Foundation::{LPARAM, LRESULT, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::{
    CallNextHookEx, GetForegroundWindow, HC_ACTION, KBDLLHOOKSTRUCT, WH_KEYBOARD_LL, WM_KEYDOWN,
    WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP,
};

/// The chords the pin key is bound to, packed as `key_binding` packs them, or zero
/// when nothing is watched. Read by the hook instead of a lock, which it may not take.
static PIN_CHORDS: [AtomicU64; crate::shell::key_binding::MAX_ALTERNATIVES] =
    [const { AtomicU64::new(0) }; crate::shell::key_binding::MAX_ALTERNATIVES];

/// The chord the preview shortcut is bound to, packed the same way and read by the hook
/// for the same reason.
static TOGGLE_CHORDS: [AtomicU64; crate::shell::key_binding::MAX_ALTERNATIVES] =
    [const { AtomicU64::new(0) }; crate::shell::key_binding::MAX_ALTERNATIVES];

/// Pin-key presses seen since startup, swapped to zero by the preview loop; the preview shortcut's own
/// count beside it, and whether each chord's main key is held — a bit per `PIN_CHORDS` slot, which is
/// what tells a press from a repeat.
///
/// All three are counted here and read on the other side, so they are published in `app::seam` and these
/// are the names they are counted by.
pub(crate) use crate::app::seam::PIN_MAIN_DOWN;
pub(crate) use crate::app::seam::PIN_PRESSES;
pub(crate) use crate::app::seam::TOGGLE_PRESSES;

/// The same bits for the preview shortcut's own chords.
static TOGGLE_MAIN_DOWN: AtomicU32 = AtomicU32::new(0);

/// The keys a standing pin answers, as the commands they mean, and the two questions asked of that
/// table — which command a virtual key is, and whether a message about one is a message the pin's own
/// window would act on.
///
/// All three are published in `app::seam`, because the pin's own window procedure reads them too and the
/// two sides have to agree about the table (see `PIN_KEY_COMMANDS` there). The order of the arms is the
/// order of the pin's own commands, and the two must be kept in step: a key listed in one and not the
/// other is a key this app acts on and also passes on, which is a double action rather than a missing
/// one.
pub(crate) use crate::app::seam::{pin_key_command, pin_key_message_acts};
#[cfg(test)]
use crate::app::seam::{PIN_COMMAND_COUNT, PIN_KEY_COMMANDS};

/// The command that means `PinCommand::TogglePlayback`, which is the one row of `PIN_KEY_COMMANDS` whose
/// repeats are not commands.
///
/// It is the pin's own arrangement (`pinned_key_down_command`), and it is the whole of what the flag
/// beside it is for: a Space is a toggle, so a hand resting on it would otherwise leave a film flickering
/// between playing and held at Windows' repeat rate, while a held arrow is a hand walking the folder and a
/// held `T` is a hand asking for the next track.
const PIN_COMMAND_TOGGLE_PLAYBACK: usize = 3;

/// Whether the toggle is held, which is what tells a Space press from a Space still held down.
///
/// One flag rather than one per key of `PIN_KEY_COMMANDS`, because the pin's own arrangement
/// answers every repeat of every one of them but the toggle — so a flag per key would be a flag
/// nothing reads.
static PIN_KEY_DOWN: AtomicBool = AtomicBool::new(false);

/// Whether a key-down is a command, from the command it means and whether the toggle was already
/// held when it arrived.
///
/// The pin's own arrangement and not this hook's: every repeat of every key of
/// `PIN_KEY_COMMANDS` is a command but the toggle's, because a Space is a play/pause and a hand
/// resting on it would otherwise leave a film flickering between playing and held, while a held
/// arrow is a hand walking the folder (see `pinned_key_down_command`).
///
/// The toggle's flag is only read for the toggle, because a flag read for an arrow would be a flag
/// a left press cleared out from under a Space that was still down.
fn pin_key_press_is_a_command(command: u8, toggle_was_down: bool) -> bool {
    command as usize != PIN_COMMAND_TOGGLE_PLAYBACK || !toggle_was_down
}

/// What the hook does with one message about a key of `PIN_KEY_COMMANDS`, once the caret is known to
/// be in one of the pin's two windows.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum PinKeyAnswer {
    /// Let through to the window the caret is in, counting nothing.
    PassedOn,
    /// Swallowed, and counted as one press of this command.
    Command(u8),
    /// Swallowed and counted by nothing: a chord the pin's own window throws away, or a release that
    /// lets a held key go.
    Swallowed,
}

/// What the pin does with one message about a key of its own, asked once so that the callback below
/// is the only place the answer is acted on.
///
/// **A find bar that is up answers `PassedOn` to every message, the key that holds a film included.**
/// The bar's field is a text field, its characters arrive as `WM_CHAR`, and `TranslateMessage` only
/// makes one from a key-down nothing swallowed: a `Space` or a `T` taken here is a query that can
/// hold neither (see `PIN_FIND_BAR_UP`). Everything else is the arrangement it always was — a key the
/// pin answers is counted and swallowed, a chord and a release are swallowed and counted by nothing
/// (see `pin_key_message_acts` and `pin_key_press_is_a_command`).
fn the_pin_answers_key(
    command: u8,
    message: u32,
    find_bar_up: bool,
    toggle_was_down: bool,
) -> PinKeyAnswer {
    if find_bar_up {
        return PinKeyAnswer::PassedOn;
    }

    if pin_key_message_acts(message) {
        return if pin_key_press_is_a_command(command, toggle_was_down) {
            PinKeyAnswer::Command(command)
        } else {
            PinKeyAnswer::Swallowed
        };
    }

    if matches!(message, WM_SYSKEYDOWN | WM_KEYUP | WM_SYSKEYUP) {
        return PinKeyAnswer::Swallowed;
    }

    PinKeyAnswer::PassedOn
}

/// What the flag saying the pin's own play/pause key is held stands once one message about one of
/// the pin's keys has been read, from what it stood before.
///
/// **The caret's loss is the release this hook will never see, and it is the only thing that clears
/// a flag the caret took with it.** The flag is read and written where the caret is this pin's, so a
/// Space let go of over the window behind — where a user types, where they rename a file — is
/// delivered to that window and never arrives here. Left standing, the flag says a Space nobody is
/// holding is down, and every Space pressed over the pin afterwards is swallowed as its repeat: the
/// film could not be held or let go of from the keyboard until the key was pressed and released
/// where the hook could see both halves again. So a caret that is not this pin's lets the flag go,
/// and the next Space over the pin is a press.
///
/// A Space genuinely still held while the caret is taken elsewhere is the price, and it is the
/// cheaper one: the user got there with the mouse while the key was down, and the cost of that is at
/// most one Space answered where a held one was expected.
fn pin_toggle_held_after(
    command: u8,
    message: u32,
    owns_caret: bool,
    toggle_was_down: bool,
) -> bool {
    // The flag is the toggle's own (see `pin_key_press_is_a_command`), so a message about any other
    // key of the table says nothing about it — whether or not this hook took the key.
    if command as usize != PIN_COMMAND_TOGGLE_PLAYBACK {
        return toggle_was_down;
    }

    if !owns_caret {
        return false;
    }

    // A release is the only message the caret's own window lets go of the flag on, and a key-down is
    // the only one that puts it down: the class this hook acts on at all is one class, so there is
    // nothing else to answer here.
    match message {
        WM_KEYDOWN => true,
        WM_KEYUP => false,
        _ => toggle_was_down,
    }
}

/// Whether a caret in one of this pin's own windows, which is the only place these keys are taken.
///
/// **It is also the whole of who owns them, and the hook is the owner: there is one answer to a
/// press rather than two.** A press taken here is swallowed, so it never becomes a message and the
/// pin's own window procedure never sees it — a swallowed key is not also a key
/// `pinned_key_down_command` maps. A press the caret is elsewhere for is passed on, and neither
/// side sees it either: Windows routes it to whatever is in front of the pin, which is a window of
/// its own. The tick does not reclaim it — the loop drains the counters this fills and acts on what
/// it finds (see `take_pin_key_presses`), which is the same press's *continuation* rather than a
/// second owner of it: nothing here counts a press and nothing there re-counts one, so there is no
/// tick that can answer a key the hook already answered.
///
/// Which of the two answers a press is decided by where the caret is, rather than by which arrived
/// first, and that is the whole of what makes the swallow safe rather than a double action. The pin's
/// own procedure is kept because it is the same five answers read from a message's `wParam` instead
/// of from this file's numbers, and because a pin holding the caret answers its keys without asking
/// a thread of another kind; it is shadowed for every key `PIN_KEY_COMMANDS` names, since a caret
/// in that window is a caret this file has already swallowed.
///
/// The pin's own window is what decides it: a zero owner means there is no pin, and a pin that has
/// just been closed has not been told that yet, so the answer has to be no for a tick after the
/// window has gone. The video handle cannot stand on its own, because it outlives the pin by a
/// tick — a player this app retired keeps its handle for the length of the hold that lets its
/// window be replaced (see `retire_replaced_player`) — and a stale video handle with no pin behind
/// it is a window of FFmpeg's that this app has no business taking keys from.
fn pin_owns_caret(owner: u64, video: u64, foreground: u64) -> bool {
    owner != 0 && foreground != 0 && (foreground == owner || (video != 0 && foreground == video))
}

/// The same question asked of the published handles rather than of numbers: whether the pin is the
/// window the user is in right now — or whether no pin is up at all.
///
/// Zero for both handles is the answer that is not a pin's: what is on screen is a hover preview,
/// which is put up under the pointer and stands in front of everything, so there is no window of
/// anybody else's for a key or a wheel notch to be taken from. A pin is the opposite — it can be
/// sitting behind another program, which is what `Pin Mode → Always On Top` off means — and while
/// the user is in that program, what arrives there is that program's, not the pin's.
///
/// It is the wheel hook's question as well as the keyboard hook's, which is why it is answered here
/// rather than beside either of them: both run inside a low-level callback, and what neither may do
/// there is walk the desktop to ask.
pub(crate) fn pin_owns_foreground() -> bool {
    let owner = PIN_KEYS_OWNER.load(Ordering::Acquire);
    if owner == 0 {
        return true;
    }

    // A pin that is Always On Top is the window the pointer is in wherever the keyboard is: it
    // stands over every other window, so the box the preview published is the box on screen and
    // a notch turned in it is the pin's own (see `Pin Mode → Always On Top`). Without this the
    // setting's own name was a lie for the wheel — the pin hovered, was seen, and was not
    // scrolled until the user clicked it, which is a click a topmost window never needed.
    if PIN_KEYS_TOPMOST.load(Ordering::Acquire) {
        return true;
    }

    // Safety: asks Windows which window is in front. It is a handle looked up, not a pointer
    // dereferenced, and it cannot fail.
    let foreground = unsafe { GetForegroundWindow().0 as u64 };
    pin_owns_caret(owner, PIN_KEYS_VIDEO.load(Ordering::Acquire), foreground)
}

/// Thread running the hook's message pump, published so shutdown can wake it.
static KEY_THREAD_ID: AtomicU32 = AtomicU32::new(0);

/// The virtual key a configured name stands for, or `None` for a name no key of
/// this app's is spelled like.
///
/// The spelling is the one `config.ini` and the tray write: a named key, a single
/// letter or digit, or `f1` through `f24`, in any case. Left and right variants of
/// the modifiers are their own names, so a user can bind the right Alt and keep the
/// left one for the trigger.
pub(crate) fn key_to_vk(key: &str) -> Option<i32> {
    let key = key.trim().to_ascii_lowercase();
    // An underscore is the same name spelled two ways: `page_down` beside `pagedown`.
    spelled_key_to_vk(&key).or_else(|| {
        key.contains('_')
            .then(|| spelled_key_to_vk(&key.replace('_', "")))
            .flatten()
    })
}

fn spelled_key_to_vk(key: &str) -> Option<i32> {
    let vk = match key {
        "alt" | "menu" => 0x12,
        "shift" => 0x10,
        "ctrl" | "control" => 0x11,
        "win" | "windows" | "meta" => 0x5B,
        "space" => 0x20,
        "tab" => 0x09,
        "enter" | "return" => 0x0D,
        "esc" | "escape" => 0x1B,
        "backspace" => 0x08,
        "capslock" | "caps_lock" => 0x14,
        "left" => 0x25,
        "up" => 0x26,
        "right" => 0x27,
        "down" => 0x28,
        "insert" | "ins" => 0x2D,
        "delete" | "del" => 0x2E,
        "home" => 0x24,
        "end" => 0x23,
        "pageup" | "pgup" => 0x21,
        "pagedown" | "pgdn" => 0x22,
        "lshift" | "leftshift" => 0xA0,
        "rshift" | "rightshift" => 0xA1,
        "lctrl" | "leftctrl" | "lcontrol" | "leftcontrol" => 0xA2,
        "rctrl" | "rightctrl" | "rcontrol" | "rightcontrol" => 0xA3,
        "lalt" | "leftalt" => 0xA4,
        "ralt" | "rightalt" => 0xA5,
        key if key.len() == 1 => {
            let byte = key.as_bytes()[0];
            if byte.is_ascii_alphabetic() {
                byte.to_ascii_uppercase() as i32
            } else if byte.is_ascii_digit() {
                byte as i32
            } else {
                return None;
            }
        }
        key if key.starts_with('f') => {
            let n = key[1..].parse::<i32>().ok()?;
            if (1..=24).contains(&n) {
                0x70 + (n - 1)
            } else {
                return None;
            }
        }
        _ => return None,
    };

    Some(vk)
}

/// Bind the hook to what the configuration currently says: the pin key where the
/// feature is on and the name is one a key is spelled like, and nothing at all
/// otherwise. Called as the app starts, by the config watcher after a reload, and
/// by the tray's own toggle — a low-level hook may not take a lock, so what it
/// reads is this number rather than the configuration itself.
pub(crate) fn refresh() {
    let (wanted, toggle_wanted) = CONFIG
        .lock()
        .map(|config| {
            let pin = if config.pin_enabled {
                let binding = crate::shell::key_binding::parse_or_default(
                    &config.pin_key,
                    crate::shell::key_binding::DEFAULT_PIN_KEY,
                );
                crate::shell::key_binding::packed_binding(&binding)
            } else {
                [0; crate::shell::key_binding::MAX_ALTERNATIVES]
            };
            let toggle = crate::shell::key_binding::parse_or_default(
                &config.toggle_key,
                crate::config::config::DEFAULT_TOGGLE_KEY,
            );
            (pin, crate::shell::key_binding::packed_binding(&toggle))
        })
        .unwrap_or((
            [0; crate::shell::key_binding::MAX_ALTERNATIVES],
            [0; crate::shell::key_binding::MAX_ALTERNATIVES],
        ));

    install_chords(&PIN_CHORDS, &PIN_MAIN_DOWN, wanted);
    install_chords(&TOGGLE_CHORDS, &TOGGLE_MAIN_DOWN, toggle_wanted);
}

/// Put a packed binding into the table the hook reads, clearing the down bits when a key
/// changed under a held key: the release of the old one is never seen once the watch has
/// moved, and a press counted for it would be a press with no finger behind it.
fn install_chords(
    chords: &[AtomicU64],
    main_down: &AtomicU32,
    wanted: [u64; crate::shell::key_binding::MAX_ALTERNATIVES],
) {
    let mut changed = false;
    for (slot, packed) in chords.iter().zip(wanted) {
        if slot.swap(packed, Ordering::AcqRel) != packed {
            changed = true;
        }
    }
    if changed {
        main_down.store(0, Ordering::Release);
    }
}

/// Count a mouse pin button's edge the way the keyboard hook counts a key's.
///
/// Called from the mouse hook, so it touches atomics only: no lock, no desktop call.
pub(crate) fn note_mouse_button(vk: i32, down: bool) {
    if note_key_edge(&PIN_CHORDS, &PIN_MAIN_DOWN, &PIN_PRESSES, vk, down, || true) {
        crate::app::seam::wake_preview_loop();
    }
}

/// Whether a key of the pin's own is down right now — a mouse button it is bound to included.
///
/// **The press a pin is taken up by is a hand and not a clock.** Explorer answers a middle press on a
/// row by selecting that row and activating its window, and when that activation is dispatched is
/// Explorer's business: it lands after this app's raise as often as before it. A band given back
/// before it is a pin put under the listing it was just taken up over, which is why the take-up asks
/// this of the hand rather than of a duration (see `pin_lifecycle::A_PRESS_BAND_IS_HELD_FOR`).
///
/// Record a main key's down or up edge against a chord table, answering whether it is a
/// press that was counted.
///
/// The count is gated by the caller rather than after the fact: the preview shortcut is
/// answered only while Explorer is in front, and a press made in another program must
/// still mark its key down so that holding it there cannot become a press when Explorer
/// comes forward.
fn note_key_edge(
    chords: &[AtomicU64],
    main_down: &AtomicU32,
    presses: &AtomicU32,
    vk: i32,
    down: bool,
    count_press: impl FnOnce() -> bool,
) -> bool {
    let slots = chords.iter().enumerate().fold(0, |slots, (index, slot)| {
        let packed = slot.load(Ordering::Acquire);
        if packed != 0 && crate::shell::key_binding::packed_main(packed) == vk {
            slots | (1u32 << index)
        } else {
            slots
        }
    });
    if slots == 0 {
        return false;
    }

    if !down {
        main_down.fetch_and(!slots, Ordering::AcqRel);
        return false;
    }

    // Marked down even when the companions are not held, so an out-of-order press cannot
    // turn a later repeat into a press (see `key_hook_proc`).
    let first = main_down.fetch_or(slots, Ordering::AcqRel) & slots == 0;
    let matched = chords.iter().enumerate().any(|(index, slot)| {
        slots & (1u32 << index) != 0
            && crate::shell::key_binding::packed_press_matches(slot.load(Ordering::Acquire), vk)
    });
    if first && matched && count_press() {
        presses.fetch_add(1, Ordering::AcqRel);
        return true;
    }
    false
}

/// The `PIN_CHORDS` slots whose main key is `vk`, as one bit each.
fn pin_main_slots(vk: i32) -> u32 {
    PIN_CHORDS
        .iter()
        .enumerate()
        .fold(0, |slots, (index, slot)| {
            let packed = slot.load(Ordering::Acquire);
            if packed != 0 && crate::shell::key_binding::packed_main(packed) == vk {
                slots | (1u32 << index)
            } else {
                slots
            }
        })
}

/// Mark these chords' main keys down and answer whether none of them already was.
fn pin_first_press(slots: u32) -> bool {
    PIN_MAIN_DOWN.fetch_or(slots, Ordering::AcqRel) & slots == 0
}

/// Whether any of these chords has `vk` as its main key with its companions held.
fn pin_chord_matches(vk: i32, slots: u32) -> bool {
    PIN_CHORDS.iter().enumerate().any(|(index, slot)| {
        slots & (1u32 << index) != 0
            && crate::shell::key_binding::packed_press_matches(slot.load(Ordering::Acquire), vk)
    })
}

/// Starts the hook thread; `main` joins the returned handle after `request_stop`.
pub(crate) fn spawn_key_watcher() -> std::thread::JoinHandle<()> {
    super::hook_thread::spawn(
        &KEY_THREAD_ID,
        WH_KEYBOARD_LL,
        Some(key_hook_proc),
        "keyboard",
    )
}

/// Wakes the hook thread's message pump so `main` can join it.
pub(crate) fn request_stop() {
    super::hook_thread::request_stop(&KEY_THREAD_ID)
}

pub(crate) unsafe extern "system" fn key_hook_proc(
    code: i32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    // Keep this callback cheap: every keystroke on the desktop waits for it to
    // return, and a hook that answers too slowly is removed by Windows. It reads a
    // few numbers, touches a few atomics, and mostly passes the message on — the
    // pin key is a key another window is entitled to as much as this app is.
    if code != HC_ACTION as i32 {
        return CallNextHookEx(None, code, wparam, lparam);
    }

    let event = &*(lparam.0 as *const KBDLLHOOKSTRUCT);
    let vk = event.vkCode as i32;
    let message = wparam.0 as u32;

    // The preview shortcut, answered only while Explorer is in front. The key is still passed
    // on whatever happens: another program's F7 is that program's, and one pressed in Explorer
    // goes on to Explorer as well. A repeat is not a press, and a key held in another program is
    // marked down so it cannot become one when Explorer comes forward.
    if matches!(message, WM_KEYDOWN | WM_SYSKEYDOWN | WM_KEYUP | WM_SYSKEYUP) {
        let down = matches!(message, WM_KEYDOWN | WM_SYSKEYDOWN);
        if note_key_edge(
            &TOGGLE_CHORDS,
            &TOGGLE_MAIN_DOWN,
            &TOGGLE_PRESSES,
            vk,
            down,
            crate::shell::window_classes::the_foreground_is_explorer,
        ) {
            crate::app::seam::wake_preview_loop();
        }
    }

    let main_slots = pin_main_slots(vk);
    if main_slots != 0 {
        match message {
            // A held key repeats, and every repeat arrives as another key-down: what
            // makes one a press is that its main key was not already down (see
            // `PIN_MAIN_DOWN`). It is marked down even when its companions are not
            // held, so that a later repeat cannot turn a chord pressed in the wrong
            // order into a press.
            WM_KEYDOWN | WM_SYSKEYDOWN if pin_first_press(main_slots) => {
                if pin_chord_matches(vk, main_slots) {
                    PIN_PRESSES.fetch_add(1, Ordering::AcqRel);
                }
            }
            WM_KEYUP | WM_SYSKEYUP => {
                PIN_MAIN_DOWN.fetch_and(!main_slots, Ordering::AcqRel);
            }
            _ => {}
        }

        return CallNextHookEx(None, code, wparam, lparam);
    }

    // A key a standing pin answers, taken here and passed on to nobody, but only while the caret is
    // in one of the pin's own two windows. The pin's window eats these itself, so the swallow costs
    // a key nothing there; FFmpeg's window does not eat them, and an arrow that reaches it is a
    // seek, which is the whole of what the swallow is for.
    if let Some(command) = pin_key_command(vk) {
        let owner = PIN_KEYS_OWNER.load(Ordering::Acquire);
        let video = PIN_KEYS_VIDEO.load(Ordering::Acquire);
        let owns_caret = pin_owns_caret(owner, video, GetForegroundWindow().0 as u64);

        // The toggle's flag is read and written once for the whole of this message, and it is the
        // flag and not the counter that needs the caret asked of first: a Space released over a
        // window that is not this pin's is a release this hook is never given (see
        // `pin_toggle_held_after`).
        let toggle_was_down = if command as usize == PIN_COMMAND_TOGGLE_PLAYBACK {
            PIN_KEY_DOWN.load(Ordering::Acquire)
        } else {
            false
        };
        if command as usize == PIN_COMMAND_TOGGLE_PLAYBACK {
            let held = pin_toggle_held_after(command, message, owns_caret, toggle_was_down);
            if held != toggle_was_down {
                PIN_KEY_DOWN.store(held, Ordering::Release);
            }
        }

        if owns_caret {
            // The answer is one value because the gate that matters is the first arm: a key the pin
            // would have answered is passed on whole while a find bar is up, so the field's
            // `WM_CHAR` can exist at all (see `the_pin_answers_key`).
            match the_pin_answers_key(
                command,
                message,
                PIN_FIND_BAR_UP.load(Ordering::Acquire),
                toggle_was_down,
            ) {
                PinKeyAnswer::PassedOn => {}
                PinKeyAnswer::Command(counted) => {
                    PIN_KEY_PRESSES[counted as usize].fetch_add(1, Ordering::AcqRel);
                    #[cfg(test)]
                    PIN_KEY_COUNTED.fetch_add(1, Ordering::Relaxed);
                    // The count is an atomic and nothing else, so the loop is woken for it: a press
                    // read at the end of a pinned static page's wait is 50 ms late, rounded up to
                    // the system timer's own 15.6 ms tick, and the caret is a mark the reader who
                    // pressed the key is watching for (see `wake_preview_loop`).
                    crate::app::seam::wake_preview_loop();
                    return LRESULT(1);
                }
                PinKeyAnswer::Swallowed => return LRESULT(1),
            }
        }
    }

    CallNextHookEx(None, code, wparam, lparam)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_key_the_settings_can_name_is_read_the_way_the_file_writes_it() {
        assert_eq!(key_to_vk("space"), Some(0x20));
        assert_eq!(key_to_vk(" Space "), Some(0x20));
        assert_eq!(key_to_vk("alt"), Some(0x12));
        assert_eq!(key_to_vk("ralt"), Some(0xA5));
        assert_eq!(key_to_vk("a"), Some(0x41));
        assert_eq!(key_to_vk("7"), Some(0x37));
        assert_eq!(key_to_vk("f8"), Some(0x77));
        assert_eq!(key_to_vk("F24"), Some(0x87));

        assert_eq!(key_to_vk("nonsense"), None);
        assert_eq!(key_to_vk("f25"), None);
        assert_eq!(key_to_vk(""), None);
    }

    /// The keys a standing pin answers, and the command each is read as, are the pin's own
    /// (`pinned_key_command`) rather than this file's own arrangement of them. A key in one list
    /// and not the other is a key this app acts on and also passes on, which is a double action
    /// rather than a missing one.
    #[test]
    fn the_keys_a_standing_pin_answers_are_the_keys_the_pin_window_answers() {
        // Every virtual key the hook reads as a command, in the command the pin's own window would
        // have given it, numbered as `hook_pin_key_command` numbers them over there.
        assert_eq!(
            pin_key_command(0x25),
            Some(0),
            "VK_LEFT is the previous file"
        );
        assert_eq!(
            pin_key_command(0x26),
            Some(5),
            "VK_UP is the page's caret a line back, which is a walk of the folder to a window with \
             no page in it"
        );
        assert_eq!(pin_key_command(0x27), Some(1), "VK_RIGHT is the next file");
        assert_eq!(
            pin_key_command(0x28),
            Some(7),
            "VK_DOWN is the page's caret a line on, and the same window walks on for it"
        );
        assert_eq!(pin_key_command(0x1B), Some(2), "VK_ESCAPE closes the pin");
        assert_eq!(
            pin_key_command(0x20),
            Some(3),
            "VK_SPACE holds the file or lets it go"
        );
        assert_eq!(
            pin_key_command(0x54),
            Some(4),
            "VK_T is the next subtitle track"
        );
        assert_eq!(
            pin_key_command(0x21),
            Some(6),
            "VK_PRIOR is the page key: a screenful back, where an arrow is a line"
        );
        assert_eq!(
            pin_key_command(0x22),
            Some(8),
            "and VK_NEXT is the screenful on"
        );
        assert_eq!(
            pin_key_command(0x24),
            Some(9),
            "VK_HOME is the beginning of the page in the window"
        );
        assert_eq!(
            pin_key_command(0x23),
            Some(10),
            "and VK_END is the last screenful of it"
        );
        assert_eq!(
            pin_key_command(0x72),
            Some(11),
            "VK_F3 takes the page to the next place its selection appears"
        );

        // A key this app does not answer is passed on untouched, or the hook would eat a letter of a
        // file the user is renaming in the window behind the pin.
        assert_eq!(pin_key_command(0x41), None, "VK_A is not this app's key");
        assert_eq!(
            pin_key_command(0x0D),
            None,
            "VK_RETURN is left to the pin's own window, which ignores it"
        );

        // Every command the table names has a counter counted into it, because a command counted
        // into a counter that was never reserved for it is a command acted on at random.
        assert!(
            PIN_KEY_COMMANDS
                .iter()
                .all(|(_, command)| (*command as usize) < PIN_COMMAND_COUNT),
            "a command of `PIN_KEY_COMMANDS` is outside the counters reserved for them"
        );

        // The toggle is named where the hook reads it rather than numbered, because the flag that
        // tells its repeats from its presses is asked about by name.
        assert_eq!(
            pin_key_command(0x20),
            Some(PIN_COMMAND_TOGGLE_PLAYBACK as u8),
            "the toggle is the key a pinned window's own play/pause is bound to"
        );
        assert_eq!(
            PIN_KEY_COMMANDS
                .iter()
                .filter(|(_, command)| *command as usize == PIN_COMMAND_TOGGLE_PLAYBACK)
                .count(),
            1,
            "exactly one key of the table is the toggle, which is `Space`"
        );
    }

    /// A key-down is a command for every one of the pin's own keys but the toggle, so a held arrow
    /// keeps walking the folder and only a held Space is swallowed.
    #[test]
    fn only_the_toggle_turns_a_repeat_into_no_command() {
        assert!(
            pin_key_press_is_a_command(0, false),
            "a first press of a key that walks the folder is one walk"
        );
        assert!(
            pin_key_press_is_a_command(0, true),
            "a held arrow walks the folder at Windows' repeat rate, which is what holding one is for"
        );
        assert!(
            pin_key_press_is_a_command(4, true),
            "a held `T` asks for the next subtitle track per repeat, and a step of it is a relaunch"
        );
        assert!(
            pin_key_press_is_a_command(PIN_COMMAND_TOGGLE_PLAYBACK as u8, false),
            "a Space press is one hold of the film"
        );
        assert!(
            !pin_key_press_is_a_command(PIN_COMMAND_TOGGLE_PLAYBACK as u8, true),
            "a Space still held is not another hold, which would leave a film flickering"
        );
    }

    /// A Space released over a window that is not the pin's is a release the hook is never given, so
    /// a flag left standing across it is a film that cannot be held or let go of from the keyboard
    /// until the key is pressed and released where both halves can be seen.
    ///
    /// The caret's loss is the notice the flag is reconciled on, because it is the only one that is
    /// guaranteed to arrive: a Space let go of over the listing behind the pin — where a user types,
    /// where a file is renamed — is delivered to that window, and this hook is not told (see
    /// `pin_toggle_held_after`).
    #[test]
    fn a_space_released_over_another_window_does_not_stick_down_in_the_table() {
        let space = PIN_COMMAND_TOGGLE_PLAYBACK as u8;

        assert!(
            !pin_toggle_held_after(space, WM_KEYUP, false, true),
            "the release went to the window the caret is in, so nothing this hook is given clears a \
             flag it is still standing on — and the next Space over the pin would be swallowed as \
             the repeat of one nobody is holding"
        );
        assert!(
            pin_toggle_held_after(space, WM_KEYDOWN, true, true),
            "while the caret is still the pin's, a Space down is a Space held, and its repeats are \
             swallowed rather than answered"
        );
        assert!(
            pin_toggle_held_after(space, WM_KEYDOWN, true, false),
            "and a first Space down is a hold of the film, which is the press the counter takes"
        );
        assert!(
            !pin_toggle_held_after(space, WM_KEYUP, true, true),
            "a release where the hook can see it is the ordinary end of the hold"
        );

        // The flag is the toggle's own, so a message about any other key of the table leaves it
        // exactly as it found it — the caret's loss included, because an arrow cannot have told the
        // flag anything either way.
        assert!(
            pin_toggle_held_after(0, WM_KEYUP, false, true),
            "a walk's release says nothing about whether a Space is held, and the caret's loss is \
             reconciled on the toggle's own messages rather than on every key of the table"
        );
        assert!(
            !pin_toggle_held_after(0, WM_KEYDOWN, false, false),
            "and a walk's press leaves the flag exactly as it found it, which here is nothing \
             standing"
        );
    }

    /// A key pressed plainly is the pin's to answer and a chord is not, which the hook reads from
    /// the pin's own rule rather than from the arm it happens to match on.
    ///
    /// The pin's own window swallows `Alt+F4` and `Alt+Tab` without answering them, so counting one
    /// here is a key acted on for a caret this app is in and let by for one it is not: a `Ctrl`+Left
    /// walked the pin's folder, and a `Ctrl`+Space held its film, each of them for a chord the pin
    /// itself had just thrown away (see `pin_key_message_acts`).
    #[test]
    fn a_chord_is_not_a_key_the_pin_answers() {
        assert!(
            pin_key_message_acts(WM_KEYDOWN),
            "a key pressed plainly is what a standing pin answers, and what the hook counts"
        );
        assert!(
            !pin_key_message_acts(WM_SYSKEYDOWN),
            "a key held with a modifier is not a walk and not a hold, so nothing is counted for it"
        );
        assert!(
            !pin_key_message_acts(WM_KEYUP),
            "and a release is not a command either — it is what lets a held one go"
        );
        assert!(
            !pin_key_message_acts(WM_SYSKEYUP),
            "nor is the release of a chord"
        );
    }

    /// **A find bar that is up takes every key the pin answers by passing it on.** A key swallowed
    /// here never becomes the `WM_CHAR` a text field types from, so `Space` and `T` — a play/pause
    /// and a subtitle step to a film — used to be characters a query could not hold. The bar's own
    /// table reads them where they land instead (see `text_find::key_edit_with`).
    #[test]
    fn a_find_bar_that_is_up_passes_on_every_key_the_pin_answers() {
        for (vk, command) in PIN_KEY_COMMANDS {
            assert_eq!(
                the_pin_answers_key(command, WM_KEYDOWN, true, false),
                PinKeyAnswer::PassedOn,
                "`{vk:#x}` is the field's while a bar is up, so the key-down has to reach the \
                 window the field lives in or the character it types never exists"
            );
            assert_eq!(
                the_pin_answers_key(command, WM_KEYDOWN, false, false),
                PinKeyAnswer::Command(command),
                "and with no bar up `{vk:#x}` is swallowed and counted exactly as it always was"
            );
        }

        // Which is not the same as the pin's other two arms: a chord and a release are swallowed
        // by the pin and let through by the bar, and a Space still held counts nothing either way.
        for message in [WM_SYSKEYDOWN, WM_KEYUP, WM_SYSKEYUP] {
            assert_eq!(
                the_pin_answers_key(0, message, true, false),
                PinKeyAnswer::PassedOn,
                "{message:#x} is the field's too: a bar cannot leave the key that let go of it \
                 behind in a window it never saw"
            );
            assert_eq!(
                the_pin_answers_key(0, message, false, false),
                PinKeyAnswer::Swallowed,
                "while the pin's own window throws it away, and the hook does the same"
            );
        }

        assert_eq!(
            the_pin_answers_key(
                PIN_COMMAND_TOGGLE_PLAYBACK as u8,
                WM_KEYDOWN,
                true,
                true
            ),
            PinKeyAnswer::PassedOn,
            "a Space typed into a query is a Space, whether or not a film would have counted it as \
             the repeat of one held"
        );
        assert_eq!(
            the_pin_answers_key(PIN_COMMAND_TOGGLE_PLAYBACK as u8, WM_KEYDOWN, false, true),
            PinKeyAnswer::Swallowed,
            "and with the bar away the held toggle is still swallowed rather than answered twice"
        );
        assert_eq!(
            the_pin_answers_key(PIN_COMMAND_TOGGLE_PLAYBACK as u8, WM_KEYDOWN, false, false),
            PinKeyAnswer::Command(PIN_COMMAND_TOGGLE_PLAYBACK as u8),
            "and a first Space is the hold it always was"
        );
    }

    /// The hook takes a key only while the caret is in one of *this* pin's two windows, so a pin up
    /// over a listing does not steal the arrows from it — and a pin that has just gone does not
    /// keep a stale handle's worth of them.
    #[test]
    fn a_key_is_only_taken_from_the_caret_while_this_pin_owns_the_window_it_is_in() {
        let owner = 0x1000;
        let video = 0x2000;

        assert!(
            pin_owns_caret(owner, video, owner),
            "the pin's own window is this app's, and a key in it is the pin's"
        );
        assert!(
            pin_owns_caret(owner, video, video),
            "FFmpeg's window belongs to the pin for as long as the pin is up, which is the whole of \
             what stops an arrow from reaching the player as a seek"
        );

        // A pin showing a picture has no player, and the zero the loop publishes for it is a handle
        // to nothing rather than a handle to something.
        assert!(
            !pin_owns_caret(owner, 0, video),
            "with no player published there is no window of this pin's to take a key in"
        );

        // Whatever is behind the pin — the listing the pin is standing over, a name being typed into,
        // another program's window — is not one of the pin's, and the hook is not a global remap.
        assert!(
            !pin_owns_caret(owner, video, 0x3000),
            "a window that is not this pin's keeps its own keys, whatever the pin is showing"
        );
        assert!(
            !pin_owns_caret(owner, video, 0),
            "no window in front at all is not the pin's window"
        );

        // A pin that has just been closed leaves a stale handle published for one tick, and
        // comparing against it must not hand a stranger's keys to a window that is gone — nor let a
        // player that outlived its pin by a tick keep them.
        assert!(
            !pin_owns_caret(0, 0, owner),
            "with no pin published there is no pin that could own the caret"
        );
        assert!(
            !pin_owns_caret(0, video, video),
            "a player that outlived its pin by a tick does not own the caret either"
        );
    }

    /// A mouse pin button is counted where the mouse hook sees it rather than on a later tick:
    /// the edge rule is the keyboard hook's own, one press per physical press and a release that
    /// clears it.
    #[test]
    fn a_mouse_button_press_is_counted_once_per_edge() {
        let binding = crate::shell::key_binding::parse_binding("mouse3,mouse3").expect("two slots");
        let packed = crate::shell::key_binding::packed_binding(&binding);
        let chords: [AtomicU64; crate::shell::key_binding::MAX_ALTERNATIVES] =
            std::array::from_fn(|index| AtomicU64::new(packed[index]));
        let down = AtomicU32::new(0);
        let presses = AtomicU32::new(0);
        let note = |vk, is_down| note_key_edge(&chords, &down, &presses, vk, is_down, || true);

        assert!(
            !note(0x05, true),
            "a button the binding does not name is not a press"
        );
        assert!(note(0x04, true), "the middle button is the binding");
        assert_eq!(presses.load(Ordering::Acquire), 1);
        assert_eq!(
            down.load(Ordering::Acquire),
            0b11,
            "every slot for that main key is marked down, not just the one that counted"
        );

        assert!(
            !note(0x04, true),
            "a second down with no release in between is not a press"
        );
        assert_eq!(presses.load(Ordering::Acquire), 1);

        assert!(!note(0x04, false), "a release counts nothing of its own");
        assert_eq!(down.load(Ordering::Acquire), 0);

        assert!(
            note(0x04, true),
            "and the next press after it is a press again"
        );
        assert_eq!(presses.load(Ordering::Acquire), 2);
    }

    /// The preview shortcut is edge-triggered like the pin key, and a press it is not allowed to
    /// answer is still marked down: a key held in another program must not become a press when
    /// Explorer comes forward, and a press it is allowed to answer is one toggle.
    #[test]
    fn the_preview_shortcut_toggles_once_and_only_where_it_is_allowed() {
        let binding = crate::shell::key_binding::parse_binding("f7").expect("a key");
        let packed = crate::shell::key_binding::packed_binding(&binding);
        let chords: [AtomicU64; crate::shell::key_binding::MAX_ALTERNATIVES] =
            std::array::from_fn(|index| AtomicU64::new(packed[index]));
        let down = AtomicU32::new(0);
        let presses = AtomicU32::new(0);
        let note = |is_down, in_explorer| {
            note_key_edge(&chords, &down, &presses, 0x76, is_down, || in_explorer)
        };

        assert!(!note(true, false), "a press in another program is ignored");
        assert_eq!(presses.load(Ordering::Acquire), 0);
        assert!(!note(true, true), "and its repeat is not a press either");
        assert_eq!(
            down.load(Ordering::Acquire),
            1,
            "the key was still marked down where it was pressed"
        );
        assert!(!note(false, true), "a release counts nothing");
        assert!(note(true, true), "a fresh press in Explorer toggles once");
        assert!(!note(true, true), "and holding it does not toggle again");
        assert_eq!(presses.load(Ordering::Acquire), 1);
    }
}

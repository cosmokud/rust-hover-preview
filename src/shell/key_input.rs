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

use crate::CONFIG;
use std::sync::atomic::{AtomicBool, AtomicI32, AtomicU32, AtomicU64, Ordering};
use windows::Win32::Foundation::{LPARAM, LRESULT, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::{
    CallNextHookEx, GetForegroundWindow, HC_ACTION, KBDLLHOOKSTRUCT, WH_KEYBOARD_LL, WM_KEYDOWN,
    WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP,
};

/// The virtual key the pin is bound to, or `0` when nothing is watched.
static PIN_VK: AtomicI32 = AtomicI32::new(0);

/// Pin-key presses seen since startup, swapped to zero by the preview loop.
static PIN_PRESSES: AtomicU32 = AtomicU32::new(0);

/// Whether the watched key is held, which is what tells a press from a repeat.
static PIN_DOWN: AtomicBool = AtomicBool::new(false);

/// The two window handles a standing pin answers keys for, as raw pointers, or zero.
///
/// Published by the preview loop on every tick rather than asked for from inside the hook, which
/// may not take a lock and may not walk the desktop: two atomics and a compare are all a
/// `WH_KEYBOARD_LL` callback may afford, because every keystroke on the machine waits for it to
/// return and Windows removes a hook that is slow.
static PIN_KEYS_OWNER: AtomicU64 = AtomicU64::new(0);
static PIN_KEYS_VIDEO: AtomicU64 = AtomicU64::new(0);

/// The keys a standing pin answers, as the commands they mean.
///
/// The set is the pin's own (`pinned_key_command`), read here as numbers because the hook thread
/// cannot take the pin's lock to ask. The order of the arms is the order of the pin's own, and the
/// two must be kept in step: a key listed in one and not the other is a key this app acts on and
/// also passes on, which is a double action rather than a missing one.
///
/// `VK_LEFT`/`VK_UP` are the same command and `VK_RIGHT`/`VK_DOWN` the same one, which is the pin's
/// own arrangement and not a coincidence: in a window with a film in it, up and down have nowhere
/// to go, so they walk the list with left and right rather than doing nothing.
pub(crate) const PIN_KEY_COMMANDS: [(i32, u8); 7] = [
    (0x25, 0), // VK_LEFT  -> Previous
    (0x26, 0), // VK_UP    -> Previous
    (0x27, 1), // VK_RIGHT -> Next
    (0x28, 1), // VK_DOWN  -> Next
    (0x1B, 2), // VK_ESCAPE -> Close
    (0x20, 3), // VK_SPACE -> TogglePlayback
    (0x54, 4), // VK_T     -> NextSubtitle
];

/// How many commands `PIN_KEY_COMMANDS` names, which is how many counters are kept for them.
///
/// Named rather than written as the `5` of an array length so that a command added to the table
/// above is a command this count has to be told about, rather than a command whose presses are
/// counted into a counter that was never reserved for it.
const PIN_COMMAND_COUNT: usize = 5;

/// The command that means `PinCommand::TogglePlayback`, which is the one row of
/// `PIN_KEY_COMMANDS` whose repeats are not commands.
///
/// It is the pin's own arrangement (`pinned_key_down_command`), and it is the whole of what the
/// flag below is for: a Space is a toggle, so a hand resting on it would otherwise leave a film
/// flickering between playing and held at Windows' repeat rate, while a held arrow is a hand
/// walking the folder and a held `T` is a hand asking for the next track.
const PIN_COMMAND_TOGGLE_PLAYBACK: usize = 3;

/// Which command of `PIN_KEY_COMMANDS` a virtual key is listed under, or `None` for a key this app
/// does not answer.
///
/// Split out of the callback below so that the table is one list of facts and this is the one place
/// it is read.
pub(crate) fn pin_key_command(vk: i32) -> Option<u8> {
    PIN_KEY_COMMANDS
        .iter()
        .find(|(key, _)| *key == vk)
        .map(|(_, command)| *command)
}

/// Tell the hook which windows a standing pin is answering for, or that there is no such pin.
///
/// Called by the preview loop on the tick it settles the pin's chrome, which is the same tick that
/// already knows whether a pin is up and which player it is showing. Passing zero for the video
/// handle is the normal case — a pin showing an image has no player to lose keys to.
pub(crate) fn publish_pin_key_owner(owner: u64, video: u64) {
    PIN_KEYS_OWNER.store(owner, Ordering::Release);
    PIN_KEYS_VIDEO.store(video, Ordering::Release);
}

/// Pin-key commands seen since this was last called, one per press, in the order pressed.
///
/// Drained by the preview loop on the same tick that drains the pin key's own counter, so a
/// command travels the same path a press of the pin's own caption button does: it is queued, and
/// the loop acts on it (see `ask_pin`).
static PIN_KEY_PRESSES: [AtomicU32; PIN_COMMAND_COUNT] = [
    AtomicU32::new(0),
    AtomicU32::new(0),
    AtomicU32::new(0),
    AtomicU32::new(0),
    AtomicU32::new(0),
];

/// Whether the toggle is held, which is what tells a Space press from a Space still held down.
///
/// One flag rather than one per key of `PIN_KEY_COMMANDS`, because the pin's own arrangement
/// answers every repeat of every one of them but the toggle — so a flag per key would be a flag
/// nothing reads.
static PIN_KEY_DOWN: AtomicBool = AtomicBool::new(false);

/// Take the pin's own key commands, one count per command and per press since the last call.
pub(crate) fn take_pin_key_presses() -> Vec<u8> {
    PIN_KEY_PRESSES
        .iter()
        .enumerate()
        .flat_map(|(command, count)| {
            std::iter::repeat_n(command as u8, count.swap(0, Ordering::AcqRel) as usize)
        })
        .collect()
}

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

/// Whether a message about a key a standing pin answers is one the pin's own window would act on.
///
/// The class of the message is part of the question rather than an accident of which arm of a
/// window procedure it arrived on: a pinned window answers a key pressed plainly and answers no
/// chord at all (`WM_SYSKEYDOWN` is swallowed there so that `Alt+F4` cannot close it and `Alt+Tab`
/// cannot leave it, and nothing is asked of it), so **this hook must not count a system key either**.
/// It did, and the difference was a chord acted on for a caret this app is in: `Ctrl`+Left walked
/// the pin's folder and `Ctrl`+Space held its film, each of them for a chord the pin's own window
/// had just thrown away.
///
/// It is one function rather than a pattern in each file because the pin's window asks it too — the
/// answer belongs to the arrangement of the pin rather than to the hook, and a hook standing in for
/// that arrangement is a hook that has to be told what it is standing in for (see
/// `preview_window::pinned_key_message_command`).
pub(crate) fn pin_key_message_acts(message: u32) -> bool {
    message == WM_KEYDOWN
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
    let vk = match key.as_str() {
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
    let wanted = CONFIG
        .lock()
        .ok()
        .filter(|config| config.pin_enabled)
        .and_then(|config| key_to_vk(&config.pin_key))
        .unwrap_or(0);

    // A key that was changed under a held key is not still held: the release of
    // the old one is never seen once the watch has moved, and a press counted for
    // it would be a press the preview loop acts on with no finger behind it.
    if PIN_VK.swap(wanted, Ordering::AcqRel) != wanted {
        PIN_DOWN.store(false, Ordering::Release);
    }
}

/// Pin-key presses since this was last called. The preview loop polls it once a
/// tick, the way the Explorer hook polls the wheel's own counter.
pub(crate) fn take_presses() -> u32 {
    PIN_PRESSES.swap(0, Ordering::AcqRel)
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

unsafe extern "system" fn key_hook_proc(code: i32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
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

    let watched = PIN_VK.load(Ordering::Acquire);
    if watched != 0 && vk == watched {
        match message {
            // A held key repeats, and every repeat arrives as another key-down: what
            // makes one a press is that the key was not already down (see `PIN_DOWN`).
            WM_KEYDOWN | WM_SYSKEYDOWN if !PIN_DOWN.swap(true, Ordering::AcqRel) => {
                PIN_PRESSES.fetch_add(1, Ordering::AcqRel);
            }
            WM_KEYUP | WM_SYSKEYUP => PIN_DOWN.store(false, Ordering::Release),
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
        let toggle_was_down = (command as usize == PIN_COMMAND_TOGGLE_PLAYBACK)
            .then(|| PIN_KEY_DOWN.load(Ordering::Acquire))
            .unwrap_or(false);
        if command as usize == PIN_COMMAND_TOGGLE_PLAYBACK {
            let held = pin_toggle_held_after(command, message, owns_caret, toggle_was_down);
            if held != toggle_was_down {
                PIN_KEY_DOWN.store(held, Ordering::Release);
            }
        }

        if owns_caret {
            // The one class the pin's own window acts on, read from its own rule rather than from
            // the arm this hook happens to match on (see `pin_key_message_acts`).
            if pin_key_message_acts(message) {
                // The flag is only read for the toggle, so an arrow held down cannot clear a Space
                // that is still held (see `pin_key_press_is_a_command`).
                if pin_key_press_is_a_command(command, toggle_was_down) {
                    PIN_KEY_PRESSES[command as usize].fetch_add(1, Ordering::AcqRel);
                }
                return LRESULT(1);
            }

            // A chord and a release are both swallowed and neither is counted. The chord is
            // swallowed exactly as the pin's own window swallows it — it is what keeps `Alt+F4` and
            // `Alt+Tab` away from a window that must not be closed or left — and it is counted by
            // nothing, because that window answers no chord either and a `Ctrl`+Left answered here
            // is a walk of a folder the user did not ask to walk. The release is taken so that
            // letting go over the pin and pressing Space again somewhere else leaves no Space stuck
            // down in the table.
            if matches!(message, WM_SYSKEYDOWN | WM_KEYUP | WM_SYSKEYUP) {
                return LRESULT(1);
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
            Some(0),
            "VK_UP is the previous file too"
        );
        assert_eq!(pin_key_command(0x27), Some(1), "VK_RIGHT is the next file");
        assert_eq!(
            pin_key_command(0x28),
            Some(1),
            "VK_DOWN is the next file too"
        );
        assert_eq!(pin_key_command(0x1B), Some(2), "VK_ESCAPE closes the pin");
        assert_eq!(pin_key_command(0x20), Some(3), "VK_SPACE holds the file or lets it go");
        assert_eq!(pin_key_command(0x54), Some(4), "VK_T is the next subtitle track");

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
}

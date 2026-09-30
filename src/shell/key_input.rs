//! System-wide keyboard watcher for the pin key.
//!
//! A low-level keyboard hook counts the pin key wherever the pointer happens to be,
//! and publishes a monotonic press counter the preview loop consumes, exactly as
//! the wheel watcher beside it publishes wheel ticks. The key is never swallowed:
//! `Space` is Explorer's own, and what the loop does with the count is its business
//! (see `PIN_PRESSES`). Only the key that puts a pin up, and the one that brings a
//! collapsed one back, is answered here; the keys a pressed pin answers arrive as
//! messages to that window instead (see `pinned_key_command`).

use crate::CONFIG;
use std::sync::atomic::{AtomicBool, AtomicI32, AtomicU32, Ordering};
use windows::Win32::Foundation::{LPARAM, LRESULT, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::{
    CallNextHookEx, HC_ACTION, KBDLLHOOKSTRUCT, WH_KEYBOARD_LL, WM_KEYDOWN, WM_KEYUP,
    WM_SYSKEYDOWN, WM_SYSKEYUP,
};

/// The virtual key the pin is bound to, or `0` when nothing is watched.
static PIN_VK: AtomicI32 = AtomicI32::new(0);

/// Pin-key presses seen since startup, swapped to zero by the preview loop.
static PIN_PRESSES: AtomicU32 = AtomicU32::new(0);

/// Whether the watched key is held, which is what tells a press from a repeat.
static PIN_DOWN: AtomicBool = AtomicBool::new(false);

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
    // return, and a hook that answers too slowly is removed by Windows. It reads
    // one number, touches two atomics, and always passes the message on — the pin
    // key is a key another window is entitled to as much as this app is.
    let watched = PIN_VK.load(Ordering::Acquire);
    if watched != 0 && code == HC_ACTION as i32 {
        let event = &*(lparam.0 as *const KBDLLHOOKSTRUCT);
        if event.vkCode as i32 != watched {
            return CallNextHookEx(None, code, wparam, lparam);
        }

        match wparam.0 as u32 {
            // A held key repeats, and every repeat arrives as another key-down: what
            // makes one a press is that the key was not already down (see `PIN_DOWN`).
            WM_KEYDOWN | WM_SYSKEYDOWN if !PIN_DOWN.swap(true, Ordering::AcqRel) => {
                PIN_PRESSES.fetch_add(1, Ordering::AcqRel);
            }
            WM_KEYUP | WM_SYSKEYUP => PIN_DOWN.store(false, Ordering::Release),
            _ => {}
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
}

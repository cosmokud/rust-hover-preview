//! System-wide keyboard watcher for the pin key.
//!
//! A preview is pinned by pressing a key while it is on screen, and the key has to
//! be seen wherever the pointer happens to be — the preview window never has
//! focus, and Explorer owns the keyboard whenever it is the active window. A
//! low-level keyboard hook reports the key without touching either of them, and
//! the hook thread publishes a monotonic press counter that the preview loop
//! consumes, exactly as the wheel watcher beside it publishes wheel ticks.
//!
//! The key is never swallowed: `Space` is Explorer's own key, and pinning a
//! preview does not make it this app's. What the hook does is count, and what the
//! preview loop does with the count is its own business (see `PIN_PRESSES`).
//!
//! This module is also where the spelling of a key name lives — `alt`, `f8`, `a` —
//! because two settings are written that way: the trigger key the Explorer hook
//! watches, and the pin key watched here.

use crate::CONFIG;
use std::sync::atomic::{AtomicBool, AtomicI32, AtomicU32, Ordering};
use std::time::Duration;
use windows::Win32::Foundation::{HINSTANCE, LPARAM, LRESULT, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::WindowsAndMessaging::{
    CallNextHookEx, GetMessageW, PostThreadMessageW, SetWindowsHookExW, UnhookWindowsHookEx,
    HC_ACTION, KBDLLHOOKSTRUCT, MSG, WH_KEYBOARD_LL, WM_KEYDOWN, WM_KEYUP, WM_QUIT, WM_SYSKEYDOWN,
    WM_SYSKEYUP,
};

/// The virtual key the pin is bound to, or `0` when nothing is watched — the
/// feature is off, or the configured name is not one this app knows. Written by
/// `refresh`, read by the hook procedure.
static PIN_VK: AtomicI32 = AtomicI32::new(0);

/// Pin-key presses seen since startup; the hook thread is the only writer and the
/// preview loop is the only reader, so a press is counted once.
static PIN_PRESSES: AtomicU32 = AtomicU32::new(0);

/// Whether the watched key is held. The hook reports a press for every key-down
/// message and a held key repeats, so this is what tells a press from a repeat.
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

/// Starts the hook thread and waits until it is ready to be stopped, so a
/// shutdown racing with startup can still end it. `main` joins the returned
/// handle after `request_stop`.
pub(crate) fn spawn_key_watcher() -> std::thread::JoinHandle<()> {
    let handle = std::thread::spawn(run_key_watcher);

    for _ in 0..200 {
        if KEY_THREAD_ID.load(Ordering::SeqCst) != 0 || handle.is_finished() {
            break;
        }
        std::thread::sleep(Duration::from_millis(5));
    }

    handle
}

/// Wakes the hook thread's message pump so `main` can join it.
pub(crate) fn request_stop() {
    let thread_id = KEY_THREAD_ID.load(Ordering::SeqCst);
    if thread_id != 0 {
        unsafe {
            let _ = PostThreadMessageW(thread_id, WM_QUIT, WPARAM(0), LPARAM(0));
        }
    }
}

/// Owns the low-level hook and the message pump it needs: the system calls the
/// hook procedure on the thread that installed it, so that thread must keep
/// dispatching messages for the life of the process.
fn run_key_watcher() {
    unsafe {
        KEY_THREAD_ID.store(GetCurrentThreadId(), Ordering::SeqCst);

        let module = GetModuleHandleW(None)
            .map(|handle| HINSTANCE(handle.0))
            .unwrap_or_default();
        let hook = match SetWindowsHookExW(WH_KEYBOARD_LL, Some(key_hook_proc), module, 0) {
            Ok(hook) => hook,
            Err(error) => {
                eprintln!("Failed to install the keyboard hook: {:?}", error);
                KEY_THREAD_ID.store(0, Ordering::SeqCst);
                return;
            }
        };

        let mut msg = MSG::default();
        while GetMessageW(&mut msg, None, 0, 0).as_bool() {}

        let _ = UnhookWindowsHookEx(hook);
        KEY_THREAD_ID.store(0, Ordering::SeqCst);
    }
}

unsafe extern "system" fn key_hook_proc(code: i32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    // Keep this callback cheap: every keystroke on the desktop waits for it to
    // return, and a hook that answers too slowly is removed by Windows. It reads
    // one number, touches two atomics, and always passes the message on — the pin
    // key is a key another window is entitled to as much as this app is.
    if code == HC_ACTION as i32 {
        let watched = PIN_VK.load(Ordering::Acquire);
        if watched != 0 {
            let event = &*(lparam.0 as *const KBDLLHOOKSTRUCT);
            if event.vkCode as i32 == watched {
                match wparam.0 as u32 {
                    WM_KEYDOWN | WM_SYSKEYDOWN => {
                        // A held key repeats, and every repeat arrives as another
                        // key-down: what makes one a press is that the key was not
                        // already down (see `PIN_DOWN`).
                        if !PIN_DOWN.swap(true, Ordering::AcqRel) {
                            PIN_PRESSES.fetch_add(1, Ordering::AcqRel);
                        }
                    }
                    WM_KEYUP | WM_SYSKEYUP => PIN_DOWN.store(false, Ordering::Release),
                    _ => {}
                }
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
        // The spellings `config.ini` is written in: a named key, a single letter or digit, and the
        // function keys — the same table the trigger key is read by, since the two are the same
        // kind of setting. A name that is not one of them is nothing to watch, which is a key
        // that is simply not bound rather than one bound to something else.
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

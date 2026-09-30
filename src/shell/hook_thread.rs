//! The thread a low-level hook lives on, shared by the watchers that install one.
//!
//! The system calls a low-level hook procedure on the thread that installed it, so
//! that thread must keep dispatching messages for the life of the process. The
//! watchdog apps use has no other pump, and the message loop is the whole of what
//! the thread does once the hook is in.

use std::sync::atomic::{AtomicU32, Ordering};
use std::time::Duration;
use windows::Win32::Foundation::{HINSTANCE, LPARAM, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::WindowsAndMessaging::{
    GetMessageW, PostThreadMessageW, SetWindowsHookExW, UnhookWindowsHookEx, HOOKPROC, MSG,
    WINDOWS_HOOK_ID, WM_QUIT,
};

/// Starts the hook thread and waits until it is ready to be stopped, so a shutdown
/// racing with startup can still end it. `main` joins the returned handle after
/// `request_stop`.
pub(crate) fn spawn(
    thread_id: &'static AtomicU32,
    hook: WINDOWS_HOOK_ID,
    proc: HOOKPROC,
    what: &'static str,
) -> std::thread::JoinHandle<()> {
    let handle = std::thread::spawn(move || run(thread_id, hook, proc, what));

    for _ in 0..200 {
        if thread_id.load(Ordering::SeqCst) != 0 || handle.is_finished() {
            break;
        }
        std::thread::sleep(Duration::from_millis(5));
    }

    handle
}

/// Wakes the hook thread's message pump so `main` can join it.
pub(crate) fn request_stop(thread_id: &AtomicU32) {
    let id = thread_id.load(Ordering::SeqCst);
    if id != 0 {
        unsafe {
            let _ = PostThreadMessageW(id, WM_QUIT, WPARAM(0), LPARAM(0));
        }
    }
}

fn run(thread_id: &AtomicU32, hook_id: WINDOWS_HOOK_ID, proc: HOOKPROC, what: &str) {
    unsafe {
        thread_id.store(GetCurrentThreadId(), Ordering::SeqCst);

        let module = GetModuleHandleW(None)
            .map(|handle| HINSTANCE(handle.0))
            .unwrap_or_default();
        let hook = match SetWindowsHookExW(hook_id, proc, module, 0) {
            Ok(hook) => hook,
            Err(error) => {
                eprintln!("Failed to install the {what} hook: {error:?}");
                thread_id.store(0, Ordering::SeqCst);
                return;
            }
        };

        let mut msg = MSG::default();
        while GetMessageW(&mut msg, None, 0, 0).as_bool() {}

        let _ = UnhookWindowsHookEx(hook);
        thread_id.store(0, Ordering::SeqCst);
    }
}

//! Named-mutex guard that keeps the app to a single running instance.

use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{CloseHandle, GetLastError, BOOL, ERROR_ALREADY_EXISTS, HANDLE};
use windows::Win32::System::Threading::CreateMutexW;

/// Session-local (`Local\`) so each signed-in user gets their own instance and
/// tray icon.
const INSTANCE_MUTEX: PCWSTR = w!(r"Local\rust-hover-preview-single-instance");

/// Keeps the instance mutex alive for as long as it is held. Dropping it
/// releases the name so the next launch becomes the primary instance.
pub struct InstanceGuard {
    /// `None` only when the mutex could not be created; the app then runs
    /// unguarded rather than refusing to start.
    handle: Option<HANDLE>,
}

impl Drop for InstanceGuard {
    fn drop(&mut self) {
        if let Some(handle) = self.handle {
            unsafe {
                let _ = CloseHandle(handle);
            }
        }
    }
}

/// Claims the single-instance guard, returning `None` when another instance
/// already holds it and this process should exit.
pub fn acquire() -> Option<InstanceGuard> {
    claim(INSTANCE_MUTEX)
}

fn claim(name: PCWSTR) -> Option<InstanceGuard> {
    let handle = match unsafe { CreateMutexW(None, BOOL(0), name) } {
        Ok(handle) => handle,
        Err(error) => {
            eprintln!("Failed to create the single-instance mutex: {error}");
            return Some(InstanceGuard { handle: None });
        }
    };

    // `GetLastError` must be read immediately after `CreateMutexW`, before any
    // other call can overwrite it.
    if unsafe { GetLastError() } == ERROR_ALREADY_EXISTS {
        unsafe {
            let _ = CloseHandle(handle);
        }
        return None;
    }

    Some(InstanceGuard {
        handle: Some(handle),
    })
}

//! Named-mutex guard that keeps the app to a single running instance.

use std::sync::Mutex;

#[cfg(test)]
use std::os::windows::ffi::OsStrExt;

use once_cell::sync::Lazy;
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

/// Safety: a guard is a handle and nothing else — no thread affinity, no borrowed state — so the
/// mutex that holds it may live in a static and be closed by whichever thread is ending the run.
unsafe impl Send for InstanceGuard {}

/// The guard this process holds, kept where a hand-over can give it up **before** the copy that
/// replaces this one starts looking for it.
///
/// A tray that ends itself for a successor — the elevated start being turned on or off — would
/// otherwise leave the successor racing a name this process still holds: the copy the task or the
/// Shell starts would find the mutex taken, decide it was the second instance, and exit, leaving no
/// tray at all. What gives it up is `let_go_early`, and what it costs is a moment in which this
/// process is still running and no longer the app's only instance — which is exactly the state it is
/// in on its way out.
static HELD: Lazy<Mutex<Option<InstanceGuard>>> = Lazy::new(|| Mutex::new(None));

impl Drop for InstanceGuard {
    fn drop(&mut self) {
        if let Some(handle) = self.handle {
            unsafe {
                let _ = CloseHandle(handle);
            }
        }
    }
}

/// Claims the single-instance guard, answering `false` when another instance already holds it and
/// this process should exit.
///
/// The guard is kept here rather than returned, because the one thing a caller would do with it —
/// hold it until the process ends — is what this already does, and the one thing a hand-over needs
/// is to be able to give it up early (see `let_go_early`).
pub fn acquire() -> bool {
    let guard = claim(INSTANCE_MUTEX);

    if guard.is_none() {
        return false;
    }

    if let Ok(mut held) = HELD.lock() {
        *held = guard;
    }

    true
}

/// Give up the single-instance name **while this process is still running**, which a tray that is
/// about to be replaced by a copy of itself does so that the copy can take it.
///
/// It is called immediately before the successor is started and nothing after it claims again: a
/// process that has given the name up is committed to leaving, and the window it opens is the one
/// Windows itself has when a program is closed and started again.
pub(crate) fn let_go_early() {
    if let Ok(mut held) = HELD.lock() {
        // Dropped here rather than at the end of the process: `InstanceGuard`'s own `Drop` closes
        // the handle, and closing the last handle is what releases the name.
        let _ = held.take();
    }
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

#[cfg(test)]
mod tests {
    use super::*;
    use std::ffi::OsString;

    /// A name of the test's own, so that the check does not depend on whether a copy of this app is
    /// running on the machine — which is the one thing the real name cannot tell a test.
    fn a_name_of_its_own(tag: &str) -> Vec<u16> {
        let name = format!(
            "Local\\rust-hover-preview-test-{tag}-{}",
            std::process::id()
        );
        OsString::from(name)
            .encode_wide()
            .chain(std::iter::once(0))
            .collect()
    }

    /// The rule the hand-over depends on: a name that is released is one the next claim takes, and
    /// releasing it twice is the same as releasing it once.
    #[test]
    fn a_name_is_taken_again_once_it_is_released() {
        let _folder = crate::config::config::ConfigFolderOfItsOwn::take();

        let name = a_name_of_its_own("hand-over");
        let first = claim(PCWSTR(name.as_ptr())).expect("a name nobody holds");
        assert!(
            claim(PCWSTR(name.as_ptr())).is_none(),
            "a name that is held is not taken a second time"
        );

        // Dropping the guard is what the hand-over does before it starts its successor.
        drop(first);
        drop(
            claim(PCWSTR(name.as_ptr()))
                .expect("a name given up early is one the copy begun right after it can take"),
        );

        // And the process-wide entry point answers honestly about the name it holds, whatever the
        // machine has: it either has it or it does not, and a release that happens either way is not
        // a panic.
        let _ = acquire();
        let_go_early();
        let_go_early();
    }
}

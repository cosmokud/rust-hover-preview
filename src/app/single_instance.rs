//! Named-mutex guard that keeps the app to a single running instance.

use std::sync::Mutex;
use std::time::{Duration, Instant};

#[cfg(test)]
use std::os::windows::ffi::OsStrExt;

use once_cell::sync::Lazy;
use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{CloseHandle, GetLastError, BOOL, ERROR_ALREADY_EXISTS, HANDLE};
use windows::Win32::System::Threading::{CreateMutexW, OpenMutexW, SYNCHRONIZATION_SYNCHRONIZE};

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

/// How long a copy that has taken the app's name is watched before it is believed to be the app:
/// a copy that dies in the first moments of its life — on a machine under a scanner's thumb, or on
/// a start that fails before its first window — is a copy that never stood, and a tray that left
/// for one would be the app gone with nothing to come back to.
const A_SUCCESSOR_SETTLES_IN: Duration = Duration::from_millis(250);

/// How often the name is asked about while a hand-over waits for its successor.
const THE_NAME_IS_ASKED_THIS_OFTEN: Duration = Duration::from_millis(25);

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

    // A poisoned lock is still the slot this guard belongs in: a guard dropped here because a panic
    // happened to be unwinding elsewhere would give the name up while this process is running, and
    // a second copy would stand beside it.
    let mut held = HELD.lock().unwrap_or_else(|poisoned| poisoned.into_inner());
    *held = guard;

    true
}

/// Give up the single-instance name **while this process is still running**, which a tray that is
/// about to be replaced by a copy of itself does so that the copy can take it.
///
/// It is called immediately before the successor is started, and the window it opens is the one
/// Windows itself has when a program is closed and started again. The one thing that may claim
/// after it is `acquire` from the hand-over's own failure path: a successor that could not be
/// started at all is a copy that never took the name, so the copy still standing takes it back
/// rather than staying as a second app beside a name nobody holds.
pub(crate) fn let_go_early() {
    // Dropped here rather than at the end of the process: `InstanceGuard`'s own `Drop` closes the
    // handle, and closing the last handle is what releases the name. A poisoned lock is taken the
    // same way round — the guard is in it, and a hand-over that skipped it would start a successor
    // that finds the app already running and leaves.
    let mut held = HELD.lock().unwrap_or_else(|poisoned| poisoned.into_inner());
    let _ = held.take();
}

/// Wait for the copy a hand-over started to take the app's name, and answer whether it did.
///
/// **A start that was accepted is not a copy that stood.** The Shell answers that it took a
/// request, and a task whose own instance is already running starts nothing at all; what proves a
/// successor is the app's own name, which only a copy that got as far as being the app holds. And a
/// copy that takes the name and dies a moment later is not one either, so the name is watched for
/// `A_SUCCESSOR_SETTLES_IN` before the answer is yes.
pub(crate) fn wait_for_a_successor(bound: Duration) -> bool {
    wait_for(INSTANCE_MUTEX, bound, A_SUCCESSOR_SETTLES_IN)
}

fn wait_for(name: PCWSTR, bound: Duration, settles_in: Duration) -> bool {
    let deadline = Instant::now() + bound;
    let mut taken_at: Option<Instant> = None;

    loop {
        match (held(name), taken_at) {
            (true, None) => taken_at = Some(Instant::now()),
            (true, Some(since)) if since.elapsed() >= settles_in => return true,
            // A name that was taken and given up again is a copy that did not stand: what is waited
            // for is the next one, until the bound runs out.
            (false, Some(_)) => taken_at = None,
            _ => {}
        }

        if Instant::now() >= deadline {
            return taken_at.is_some();
        }

        std::thread::sleep(THE_NAME_IS_ASKED_THIS_OFTEN);
    }
}

/// Whether the name is held by some process, asked without claiming it.
fn held(name: PCWSTR) -> bool {
    // Safety: the name is opened for the least a handle can be asked of it, and the handle this
    // call answers with is closed here. A name nobody holds is an error rather than a handle.
    match unsafe { OpenMutexW(SYNCHRONIZATION_SYNCHRONIZE, false, name) } {
        Ok(handle) => {
            // Safety: the handle `OpenMutexW` made and nothing else holds.
            unsafe {
                let _ = CloseHandle(handle);
            }
            true
        }
        Err(_) => false,
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

    /// The rule a hand-over's wait is made of: a name nobody holds is not a successor, a name that is
    /// held is, and a copy that takes the name and gives it up again is not one either.
    #[test]
    fn a_successor_is_a_copy_that_holds_the_name() {
        let name = a_name_of_its_own("successor");

        assert!(
            !wait_for(
                PCWSTR(name.as_ptr()),
                Duration::from_millis(120),
                Duration::ZERO
            ),
            "a name nobody holds is a copy that never came"
        );

        let held = claim(PCWSTR(name.as_ptr())).expect("a name nobody holds");
        assert!(
            wait_for(
                PCWSTR(name.as_ptr()),
                Duration::from_millis(500),
                Duration::ZERO
            ),
            "a copy that took the name is one that stood"
        );

        // A copy that dies before it settles is not one: the name is watched for a moment, and what
        // it is watched for is a holder that is still there at the end of it.
        let dying = std::thread::spawn(move || {
            std::thread::sleep(Duration::from_millis(40));
            drop(held);
        });

        assert!(
            !wait_for(
                PCWSTR(name.as_ptr()),
                Duration::from_millis(600),
                Duration::from_millis(200)
            ),
            "a copy that gave the name up again is a copy that did not stand"
        );

        let _ = dying.join();
    }
}

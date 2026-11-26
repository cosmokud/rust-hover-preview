//! Detach Mode: a pin that is a window of its own rather than a pin of the tray's.
//!
//! A pin is one window and one whole set of process-wide facts — its state, its bubble, its film
//! host, its chrome strips, the player it plays through, the thumbnail it is drawn by — and those
//! facts are singletons by construction (see `ui::preview_window::pin_window::PIN_STATE`, and the
//! statics of every `pin_*.rs` beside it). Detaching a pin is not a second pin inside the same
//! process: it is *this* pin, and the viewer process that was drawing hovers stops serving hovers
//! and stands as a window of its own while the tray spawns a fresh viewer (see `app::viewer` and
//! `pin_detach::detach_the_pinned_window`).

use std::ffi::OsString;
use std::sync::atomic::{AtomicBool, Ordering};

use windows::Win32::Foundation::{CloseHandle, BOOL};

/// The switch that says this process draws for a run of the tray app's that is already going, and the pid
/// of that run: every name the seam is known by is made of it (see `app::seam::the_seam_name`).
pub(crate) const VIEWER: &str = "--viewer";

/// How this process was started.
#[derive(PartialEq, Debug)]
pub(crate) enum Startup {
    /// The tray app, which is what a launch with no switch of ours on it is.
    Tray,
    /// The elevated copy of the app that makes the one write a process without a token of its own
    /// may not make — the scheduled task that starts this app at a logon (see `app::startup`). It
    /// registers or removes that task and exits: no tray, no viewer, no instance.
    Elevated(crate::app::startup::Applied),
    /// The viewer of the run named by the pid that comes with it: the process every preview is drawn in,
    /// which owns no listing and watches nothing of one (see `app::viewer` and `run_a_viewer`).
    Viewer(u32),
    /// `--viewer` with no valid tray pid behind it. Nothing is started rather than a second tray icon.
    Refused,
}

/// Whether this process is a detached window, published so the loop's tick may ask lock-free.
static DETACHED_HERE: AtomicBool = AtomicBool::new(false);

/// How this process was started, read from its own command line.
pub(crate) fn startup_mode() -> Startup {
    parse(std::env::args_os().skip(1))
}

fn parse(args: impl IntoIterator<Item = OsString>) -> Startup {
    // The elevated helper is a launch of its own kind rather than a mode of this one: it writes one
    // thing and exits, and every road below is a road it does not take.
    let arguments: Vec<OsString> = args.into_iter().collect();

    if arguments
        .iter()
        .any(|arg| arg.to_str() == Some(crate::app::startup::ELEVATE_SWITCH))
    {
        // A launch that was begun for one write and cannot say which is one that must not stand as
        // a second tray icon: nothing is started rather than the wrong thing.
        return crate::app::startup::helper_request(arguments.iter().cloned())
            .map_or(Startup::Refused, Startup::Elevated);
    }

    let mut args = arguments.into_iter();
    let mut a_viewer_was_asked_for = false;
    let mut viewer: Option<u32> = None;

    while let Some(arg) = args.next() {
        if arg.to_str() == Some(VIEWER) {
            a_viewer_was_asked_for = true;
            viewer = args
                .next()
                .and_then(|value| value.into_string().ok())
                .and_then(|value| value.parse().ok());
        }
    }

    if a_viewer_was_asked_for {
        return viewer.map_or(Startup::Refused, Startup::Viewer);
    }

    Startup::Tray
}

/// Whether this process's pin has been detached into a window of its own.
pub(crate) fn is_detached() -> bool {
    DETACHED_HERE.load(Ordering::Acquire)
}

/// Mark this viewer's pin as a standalone detached window for the rest of the process lifetime.
pub(crate) fn this_window_is_its_own() {
    DETACHED_HERE.store(true, Ordering::Release);
}

#[cfg(test)]
pub(crate) fn forget_this_window_is_its_own() {
    DETACHED_HERE.store(false, Ordering::Release);
}

/// End this process once its detached window closes.
pub(crate) fn the_process_may_go() {
    crate::RUNNING.store(false, Ordering::SeqCst);
}

/// End this process when `parent` (the tray app) exits, fallback when job adoption fails.
pub(crate) fn watch_the_process(parent: u32) {
    use windows::Win32::System::Threading::{
        OpenProcess, WaitForSingleObject, PROCESS_SYNCHRONIZE,
    };

    std::thread::spawn(move || {
        let process = unsafe { OpenProcess(PROCESS_SYNCHRONIZE, BOOL(0), parent) };
        let Ok(process) = process else {
            the_process_may_go();
            return;
        };
        unsafe {
            let _ = WaitForSingleObject(process, u32::MAX);
            let _ = CloseHandle(process);
        }
        the_process_may_go();
    });
}

#[cfg(test)]
mod tests {
    use super::*;

    fn arguments(line: &[&str]) -> Vec<OsString> {
        line.iter().map(OsString::from).collect()
    }

    #[test]
    fn a_launch_with_no_switch_of_ours_on_it_is_the_tray_app() {
        assert_eq!(parse(arguments(&[])), Startup::Tray);
        assert_eq!(parse(arguments(&["--verbose"])), Startup::Tray);
    }

    #[test]
    fn the_elevated_helper_is_a_launch_of_its_own_kind() {
        use crate::app::startup::{Applied, ELEVATE_SWITCH};

        assert_eq!(
            parse(arguments(&[ELEVATE_SWITCH, "on"])),
            Startup::Elevated(Applied::TurnOn)
        );
        assert_eq!(
            parse(arguments(&[ELEVATE_SWITCH, "off"])),
            Startup::Elevated(Applied::TurnOff)
        );
        // A word this app does not know is not a request to write anything, and it is not the
        // tray either: a launch that was begun for one write and cannot say which is one that
        // must not stand as a second tray icon.
        assert_eq!(parse(arguments(&[ELEVATE_SWITCH])), Startup::Refused);
        assert_eq!(
            parse(arguments(&[ELEVATE_SWITCH, "maybe"])),
            Startup::Refused
        );
        // And every other launch is what it always was.
        assert_eq!(parse(arguments(&[])), Startup::Tray);
        assert_eq!(parse(arguments(&["--viewer", "7"])), Startup::Viewer(7));
    }

    #[test]
    fn a_viewer_names_the_run_it_draws_for() {
        assert_eq!(
            parse(arguments(&["--viewer", "4242"])),
            Startup::Viewer(4242)
        );
    }

    #[test]
    fn a_viewer_that_names_no_run_begins_nothing() {
        assert_eq!(parse(arguments(&["--viewer"])), Startup::Refused);
        assert_eq!(
            parse(arguments(&["--viewer", "not a pid"])),
            Startup::Refused
        );
        assert_eq!(parse(arguments(&["--viewer", ""])), Startup::Refused);
    }
}

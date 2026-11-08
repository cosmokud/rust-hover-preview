//! The viewer: the one process of this executable a run draws every preview in.
//!
//! **The tray renders nothing and the viewer renders everything.** A run is two roles of one executable,
//! decided once at start-up and never again: the tray owns the listing — the Explorer hook, the two key
//! watches, the tray icon, the sweeps of what earlier runs left behind — and the viewer owns every pixel —
//! the preview loop, the low-level wheel hook, the media, the engines. What crosses between them is the
//! seam's two rings and the facts beside them (see `app::seam`), and what that buys is a preview whose
//! media is already open in a process that can hand its window to another one without reading the file a
//! second time.
//!
//! A viewer is begun once, at start-up, and replaced when it dies: a viewer that is gone is a hover nobody
//! draws, so the next show the tray is about to send begins another one (see
//! `make_sure_a_viewer_is_standing`). A run that cannot begin one at all is not left with a tray and no
//! previews — it draws them itself, which is exactly the run this app was before the split (see `main`).

use std::os::windows::io::IntoRawHandle;
use std::sync::atomic::{AtomicIsize, AtomicU64, Ordering};
use std::time::{Duration, Instant};

use once_cell::sync::Lazy;
use windows::Win32::Foundation::{CloseHandle, HANDLE, WAIT_TIMEOUT};
use windows::Win32::System::Threading::WaitForSingleObject;

/// How long a viewer that was just begun is given before another may be. A viewer that dies as it starts
/// would otherwise be begun once per hover, and a hover comes at pointer pace.
const RESPITE: Duration = Duration::from_secs(1);

/// The viewer's process, kept as the number a handle is rather than as the handle for the reason every
/// handle in this tree is: it is asked whether the viewer is still there on the hover path, which may not
/// take a lock.
///
/// It is never closed while it is the recorded one. A pid the machine has handed to another process would
/// answer for the viewer that is gone, and holding the handle is also what keeps that pid from being
/// handed out again.
static THE_VIEWER: AtomicIsize = AtomicIsize::new(0);

/// The clock the pause between two attempts is measured on, and when the next one is due.
static THE_CLOCK: Lazy<Instant> = Lazy::new(Instant::now);
static NEXT_ATTEMPT: AtomicU64 = AtomicU64::new(0);

/// Begin the viewer this run draws through, and answer whether one is standing.
///
/// Called once, by `main`, before the threads that publish are begun: where it answers false the run draws
/// its own previews, and no viewer of its own is ever asked for again.
pub(crate) fn begin_the_viewer() -> bool {
    let Some(child) = begin_one() else {
        return false;
    };

    keep_the_handle(child);
    true
}

/// Make sure a viewer is standing, beginning one where the one there was has gone.
///
/// Asked before a show is put in the ring, and by nothing else (see `seam::send_preview`). Nothing is said
/// to the user about a viewer's death: the tray has no window of its own that could say anything, and the
/// answer is a replacement begun on the next hover that would otherwise have been drawn by nobody.
pub(crate) fn make_sure_a_viewer_is_standing() {
    if the_viewer_is_standing() || !an_attempt_is_due() {
        return;
    }

    let Some(child) = begin_one() else {
        return;
    };

    note_a_replacement(child.id());
    keep_the_handle(child);
}

/// One viewer, begun and recorded as the process this run's previews are drawn in.
fn begin_one() -> Option<std::process::Child> {
    let now = THE_CLOCK.elapsed().as_millis() as u64;
    NEXT_ATTEMPT.store(now + RESPITE.as_millis() as u64, Ordering::Release);

    let exe = std::env::current_exe().ok()?;
    let child = crate::app::engine_processes::hidden_command(exe)
        // The run it draws for: the section, the events and the facts are all named for that pid (see
        // `app::seam::the_seam_name`), and a viewer is told it rather than made to guess.
        .arg(crate::app::detach::VIEWER)
        .arg(std::process::id().to_string())
        .spawn()
        .ok()?;

    // The job every process of this app's is in, which ends the viewer with the tray app: a viewer left
    // standing with no tray to ask it for anything is a window nothing would ever close.
    crate::app::engine_processes::adopt(child.id());
    // And the one fact about it the tray's own windows are judged by (see `seam::window_is_a_viewers`).
    crate::app::seam::note_the_viewer_process(child.id());

    Some(child)
}

/// Keep the viewer's process handle, and close the one a viewer that has gone left behind.
fn keep_the_handle(child: std::process::Child) {
    let handle = child.into_raw_handle() as isize;
    let before = THE_VIEWER.swap(handle, Ordering::AcqRel);
    if before == 0 {
        return;
    }

    // Safety: the handle recorded before this one, which this process owns and nothing asks about any
    // more: a viewer that has gone is not looked for again.
    unsafe {
        let _ = CloseHandle(HANDLE(before as *mut core::ffi::c_void));
    }
}

fn the_viewer_is_standing() -> bool {
    let handle = THE_VIEWER.load(Ordering::Acquire);
    if handle == 0 {
        return false;
    }

    // Safety: a process handle this process owns, waited on for no time at all: the call asks whether the
    // process is still there rather than waiting for it to stop being there.
    unsafe { WaitForSingleObject(HANDLE(handle as *mut core::ffi::c_void), 0) == WAIT_TIMEOUT }
}

/// Whether it is time to try again (see `RESPITE`).
fn an_attempt_is_due() -> bool {
    (THE_CLOCK.elapsed().as_millis() as u64) >= NEXT_ATTEMPT.load(Ordering::Acquire)
}

/// Say that a viewer had to be begun again, on the seam's own trace: a viewer that keeps dying is the one
/// thing this road cannot say out loud, and a run that asked for the trace asked to be told.
fn note_a_replacement(pid: u32) {
    if std::env::var_os("RHP_SEAM_TRACE").is_some() {
        eprintln!("seam: the viewer had gone, {pid} begun in its place");
    }
}

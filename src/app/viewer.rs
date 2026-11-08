//! The viewer: the one process of this executable a run draws every preview in.
//!
//! **The tray renders nothing and the viewer renders everything.** A run is two roles of one executable,
//! decided once at start-up and never again: the tray owns the listing — the Explorer hook, the two key
//! watches, the tray icon, the sweeps of what earlier runs left behind — and the viewer owns every pixel —
//! the preview loop, the low-level wheel hook, the media, the engines. What crosses between them is the
//! seam's two rings and the facts beside them (see `app::seam`), and what that buys is a preview whose
//! media is already open in a process that keeps drawing it when its pin is detached, rather than being
//! handed to a process that would have to read the file a second time (see `app::detach`).
//!
//! **A run outlives its viewer, and may have more than one.** A viewer is begun at start-up and replaced
//! when it dies: one that is gone is a hover nobody draws, so the next show the tray is about to send
//! begins another (see `make_sure_a_viewer_is_standing`). And it is begun again, at once, when the one
//! there was becomes a window of its own, because the pin that was detached was drawn in the process the
//! hovers were drawn in — the run's hovers have nobody else (see `another_viewer_is_owed`). A run that
//! cannot begin one at all is not left with a tray and no previews — it draws them itself, which is exactly
//! the run this app was before the split (see `main`).

use std::os::windows::io::IntoRawHandle;
use std::sync::atomic::{AtomicIsize, AtomicU64, Ordering};
use std::time::{Duration, Instant};

use once_cell::sync::Lazy;
use windows::Win32::Foundation::{CloseHandle, HANDLE, WAIT_OBJECT_0, WAIT_TIMEOUT};
use windows::Win32::System::Threading::WaitForSingleObject;

/// How long a viewer that was just begun is given to say its preview loop is running (see
/// `begin_the_viewer`).
///
/// It is a bound on a start rather than on a preview: a viewer that has not answered by the end of it is one
/// this run will never talk to, and the road it falls back to is the one this app took before the split — so
/// a machine slow enough to miss it loses the second process, never the previews.
const THE_VIEWER_ANSWERS_IN: Duration = Duration::from_secs(3);

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
///
/// **A viewer is standing when it says so, not when it was spawned.** A process is begun whether or not it
/// ever joins the section and turns its loop, and a run that took the rings' road on the strength of a
/// spawn alone would publish every hover of the session into a ring nobody reads — no previews at all, from
/// a start-up that reported nothing wrong. So the answer is the readiness event the viewer sets once it is
/// up, waited on for a bound: a viewer that does not answer is a process nothing here will ever talk to, and
/// this run is the single-process app instead (see `main` and `seam::preview_ready_event`).
pub(crate) fn begin_the_viewer() -> bool {
    let Some(mut child) = begin_one() else {
        return false;
    };

    if !wait_for_a_viewer(child.id(), THE_VIEWER_ANSWERS_IN) {
        // Ended rather than left standing: a viewer that answered nothing draws nothing, and a second
        // process holding a preview loop of its own is one this run has no use for. And this process's own
        // windows are once again the ones the hook judges a press by (see `seam::window_is_a_viewers`).
        let _ = child.kill();
        crate::app::seam::note_the_viewer_process(0);
        return false;
    }

    keep_the_handle(child);
    true
}

/// Wait for the viewer `viewer` to answer that its preview loop is running (see
/// `seam::preview_ready_event`).
///
/// The event is the viewer's own, so the handle is this call's to close: the tray waits on it once, and the
/// process that made the name keeps it for as long as it is the answer (see `begin_one`).
fn wait_for_a_viewer(viewer: u32, bound: Duration) -> bool {
    let event = crate::app::seam::preview_ready_event(viewer);
    if event.is_invalid() {
        return false;
    }

    // Safety: an event of this run's own, waited on for a bound and nothing else: nothing here is
    // dereferenced.
    let answered = unsafe { WaitForSingleObject(event, bound.as_millis() as u32) == WAIT_OBJECT_0 };

    // Safety: the handle this call made and nothing else holds, closed by the call that made it.
    unsafe {
        let _ = CloseHandle(event);
    }
    answered
}

/// Begin the viewer that draws the hovers of a run whose viewer has become a window of its own, and make it
/// the one the rings are read by (see `app::detach`).
///
/// **The wait is a viewer's start, so it is a thread's own.** Called from the tray side that read the
/// notice, which is the hook's loop: a loop that sat out a process's start-up would be a hook that stopped
/// answering Explorer for as long as the launch took.
pub(crate) fn another_viewer_is_owed(pid: u32) {
    // **A run that draws its own previews owes no viewer.** The window that left was drawn by this process,
    // which is still drawing the hovers, and a viewer begun here would read a ring nothing publishes into
    // (see `main` and `seam::the_rings_are_the_road`).
    if !crate::app::seam::the_rings_are_the_road() {
        return;
    }

    // The window that left is still this app's own, standing over the desktop: the run keeps it in the
    // record a press on it is judged by (see `seam::window_is_a_viewers`).
    crate::app::seam::note_a_detached_viewer(pid);

    std::thread::spawn(begin_the_viewer_again);
}

fn begin_the_viewer_again() {
    // A hover may have got there first and begun the viewer this run needs (see
    // `make_sure_a_viewer_is_standing`): what is owed then is nothing, and a second one begun here would be
    // two processes reading one preview ring.
    if the_viewer_is_standing() && !crate::app::seam::the_viewer_has_left() {
        return;
    }

    // `begin_one` notes the new process as this run's viewer before it can answer, which is what makes the
    // retarget the first thing that happens: every hover that lands while the new one is starting is a
    // message in the preview ring, waiting for the one process that reads it.
    let Some(mut child) = begin_one() else {
        return;
    };

    if wait_for_a_viewer(child.id(), THE_VIEWER_ANSWERS_IN) {
        keep_the_handle(child);
        return;
    }

    // A viewer that answered nothing draws nothing, and the road back is the one every lost viewer's is:
    // the next show begins another (see `make_sure_a_viewer_is_standing`).
    let _ = child.kill();
    crate::app::seam::note_the_viewer_process(0);
}

/// Say which road this run took, on the trace a run asks for when what is under the seam is measured: the
/// rings, where a viewer answered, or this process's own channel and preview thread, where none did (see
/// `begin_the_viewer` and `main`).
pub(crate) fn note_the_road(two_processes: bool) {
    if std::env::var_os("RHP_SEAM_TRACE").is_none() {
        return;
    }

    if two_processes {
        eprintln!("seam: a viewer answered, the rings are the road");
    } else {
        eprintln!("seam: no viewer answered, this process draws its own previews");
    }
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

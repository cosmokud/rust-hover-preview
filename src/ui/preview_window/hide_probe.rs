//! **What the pin is actually showing while a relaunch's player is off the screen**, read off the
//! desktop rather than off this app's own bookkeeping.
//!
//! It exists for one question and one road: `RHP_FREEZE=hide` freezes the outgoing film by hiding the
//! player's window, which rests entirely on one unmeasured sentence — that the Shell stops drawing a
//! hidden window's thumbnail and leaves it frozen at its last frame (see `pin_film_thumbnail`'s module
//! note). A road that "looks" frozen and is black is the failure every earlier cover was measured as,
//! so the pin's own box is sampled as composed on the desk: a percentage of lit points, twenty times a
//! second, from the moment of the hide to a moment past the swap.
//!
//! The read is `surfaces::screen_pixels`, the same `BitBlt` off the screen a pin's taskbar picture is
//! taken with — **not** `PrintWindow`, which answers a layered or D3D window with black and would
//! measure the sampler rather than the pin.
//!
//! The gate is `RHP_PIN_TRACE`, because a probe that runs is a screen read twenty times a second: a
//! run without it spawns no thread and reads nothing.

use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Mutex;
use std::time::{Duration, Instant};

use windows::Win32::Foundation::{HWND, RECT};

use super::pin_trace::{asked, note_pin_trace};

/// How often the pin is read off the desk: a gap of 350-600 ms is what a relaunch costs, so twenty
/// samples a second resolves both ends of it and the black inside it.
const SAMPLE_EVERY: Duration = Duration::from_millis(50);

/// How long the pin is watched past the swap, so a picture that only arrives with the swap is seen
/// arriving rather than assumed.
const AFTER_SWAP: Duration = Duration::from_millis(300);

/// The bound on a probe that never sees a swap at all (a relaunch that died): the player-start bound
/// with room over it, so a thread is never left reading the screen for a pin that has gone.
const THE_BOUND: Duration = Duration::from_secs(12);

/// The points a percentage is counted over, on each side of the grid.
const GRID: u32 = 16;

/// One relaunch's measurement.
struct Probe {
    /// The player whose window was hidden, as the trace names it.
    pid: u32,
    /// That window, as a number, for the `source` field.
    source: isize,
    /// The pin's own box on the desk, read while it was still a box.
    area: RECT,
    /// The moment of the hide, which every `t=+<N>ms` is counted from.
    started: Instant,
    /// The last sample, for the cadence a run is read against.
    last: Instant,
    /// The first sample that found nothing lit, or `None` where the pin never went black.
    black_at: Option<u128>,
    /// The longest wait between two samples, which says whether the cadence held.
    max_gap_ms: u128,
    /// Whether the replacement's own source has drawn yet (see `note_drawn`).
    drawn: bool,
    /// The moment the picture changed hands, or `None` while the swap is still owed (see
    /// `note_swapped`).
    swapped_at: Option<Instant>,
}

/// The relaunch being measured, or nothing between relaunches: one at a time, because a pin has one
/// film on screen and a second seek restarts the measurement rather than stacking one on another.
static PROBE: Mutex<Option<Probe>> = Mutex::new(None);

/// Whether the sampling thread is up, so one probe is one thread however many relaunches it spans.
static RUNNING: AtomicBool = AtomicBool::new(false);

/// Begin measuring a relaunch: the player whose window was just hidden, and the pin it drew.
///
/// Nothing is measured without `RHP_PIN_TRACE`, and nothing is measured for a pin with no host to be
/// read from — a hide with no box to sample is a run that has nothing to answer.
pub(super) fn start(pid: u32) {
    if !asked() {
        return;
    }
    let Some(host) = super::pin_film_host::film_host() else {
        note_pin_trace!("hide probe unavailable pid {pid}: the pin has no host to be read from");
        return;
    };
    // Safety: the host is this app's own window, and what is asked of it is where its client stands
    // on the screen.
    let Some(area) = (unsafe { super::pin_taskbar_mirror::client_area_on_the_screen(host) }) else {
        note_pin_trace!("hide probe unavailable pid {pid}: the pin's box is not on the screen");
        return;
    };
    let source = super::pin_film_thumbnail::the_source_window().map_or(0, |source| source.0 as isize);
    let now = Instant::now();
    let Ok(mut held) = PROBE.lock() else {
        return;
    };
    if held.is_some() {
        note_pin_trace!("hide probe restarted pid {pid}: a newer relaunch is the one measured");
    }
    *held = Some(Probe {
        pid,
        source,
        area,
        started: now,
        last: now,
        black_at: None,
        max_gap_ms: 0,
        drawn: false,
        swapped_at: None,
    });
    drop(held);

    if RUNNING.swap(true, Ordering::AcqRel) {
        return;
    }
    std::thread::spawn(run);
}

/// The replacement's source has drawn what it was placed for, which is the last step before the
/// swap (see `pin_film_thumbnail::hand_the_picture_over`).
pub(super) fn note_drawn(pid: u32) {
    with_probe(pid, |probe| probe.drawn = true);
}

/// The picture has changed hands, so the gap is over: the result is written once, here, and the
/// sampling goes on for a moment past it.
pub(super) fn note_swapped(pid: u32) {
    let Ok(mut held) = PROBE.lock() else {
        return;
    };
    let Some(probe) = held.as_mut().filter(|probe| probe.pid == pid) else {
        return;
    };
    if probe.swapped_at.is_some() {
        return;
    }
    probe.swapped_at = Some(Instant::now());
    let (black_at, max_gap_ms) = (probe.black_at, probe.max_gap_ms);
    drop(held);
    note_pin_trace!(
        "hide probe result frozen={} black_at={} max_gap_ms={}",
        if black_at.is_none() { "yes" } else { "no" },
        match black_at {
            Some(t) => format!("+{t}ms"),
            None => "never".to_string(),
        },
        max_gap_ms,
    );
}

/// Do one thing to the probe measuring `pid`, and to no other: a relaunch restarts the measurement,
/// so a step of the one it replaced is dropped rather than written over the newer one's.
fn with_probe(pid: u32, change: impl FnOnce(&mut Probe)) {
    let Ok(mut held) = PROBE.lock() else {
        return;
    };
    if let Some(probe) = held.as_mut().filter(|probe| probe.pid == pid) {
        change(probe);
    }
}

/// The sampling thread: one line a sample, and the pin read off the desk rather than asked of itself.
fn run() {
    loop {
        let now = Instant::now();
        let step = {
            let Ok(held) = PROBE.lock() else {
                RUNNING.store(false, Ordering::Release);
                return;
            };
            let Some(probe) = held.as_ref() else {
                RUNNING.store(false, Ordering::Release);
                return;
            };
            let gap = now.saturating_duration_since(probe.last);
            (
                probe.pid,
                probe.source,
                probe.area,
                probe.started,
                probe.drawn,
                probe.swapped_at,
                gap.as_millis(),
            )
        };
        let (pid, source, area, started, drawn, swapped_at, gap_ms) = step;
        let t = started.elapsed().as_millis();
        let lit = the_pin_is_lit(area);
        let state = if swapped_at.is_some() {
            "swapped"
        } else if drawn {
            "drawn"
        } else {
            "hidden"
        };
        note_pin_trace!(
            "hide probe t=+{t}ms pin lit={lit}% source={:?} pid={pid} state={state}",
            HWND(source as *mut _),
        );
        if let Ok(mut held) = PROBE.lock() {
            if let Some(probe) = held.as_mut().filter(|probe| probe.pid == pid) {
                probe.last = now;
                probe.max_gap_ms = probe.max_gap_ms.max(gap_ms);
                // **Only the gap counts as the gap**: a sample after the swap is the picture the
                // replacement put there, and a film that opens on black is not a pin that went black.
                if swapped_at.is_none() && probe.black_at.is_none() && lit == 0 {
                    probe.black_at = Some(t);
                }
                if probe
                    .swapped_at
                    .is_some_and(|at| now.duration_since(at) >= AFTER_SWAP)
                    || started.elapsed() >= THE_BOUND
                {
                    *held = None;
                    RUNNING.store(false, Ordering::Release);
                    return;
                }
            }
        }
        std::thread::sleep(SAMPLE_EVERY);
    }
}

/// The percentage of a grid inside the pin's own box that is lit, read off the desktop as it is
/// composed.
///
/// The grid sits in the middle of the box rather than across all of it, because the pin's chrome — the
/// caption's strip and the transport bar — is drawn into the band's edges, and a lit bar over a black
/// film would read as a picture that is still there.
///
/// shortcut: a film that is black at every sampled point — a fade, black leader, a dark scene — reads
/// as `0%` exactly like a black hole, so a run is read against a bright film, which is what the run
/// book asks for. Upgrade it by comparing the samples rather than counting lit pixels, if a road ever
/// needs to be measured against a dark one.
fn the_pin_is_lit(area: RECT) -> u32 {
    let Some((pixels, width, height)) = super::surfaces::screen_pixels(area) else {
        return 0;
    };
    let (width, height) = (width as u32, height as u32);
    if width == 0 || height == 0 {
        return 0;
    }
    let stride = width as usize * 4;
    let (left, right) = (width / 4, width * 3 / 4);
    let (top, bottom) = (height / 4, height * 3 / 4);
    let mut lit = 0;
    for row in 0..GRID {
        let y = top + (bottom - top) * row / GRID;
        for column in 0..GRID {
            let x = left + (right - left) * column / GRID;
            let at = y as usize * stride + x as usize * 4;
            if pixels[at] | pixels[at + 1] | pixels[at + 2] != 0 {
                lit += 1;
            }
        }
    }
    // A point is lit or it is not; the count is over the grid, so this is the percentage of it.
    lit * 100 / (GRID * GRID) as u32
}

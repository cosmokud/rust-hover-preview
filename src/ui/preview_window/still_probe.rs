//! **What the pin is actually showing while a relaunch has no source of its own drawing in it**,
//! read off the desktop rather than off this app's own bookkeeping.
//!
//! It began as the measurement of one sentence and it is now the measurement of the road that ships.
//! A seek press asks FFmpeg for the frame on screen while the hand is still on the bar, and the
//! release paints it into the pin's host, lets the registration go and ends the outgoing player at
//! once (see `pin_film_thumbnail::stand_the_pressed_still`). **Whether that still is what the pin
//! then shows is not a question this app can answer about itself** — a DWM thumbnail cannot outlive
//! the drawing of its source, so the still is what replaces it — and the pin's own box is sampled as
//! composed on the desk instead: a percentage of lit points, twenty times a second, from the
//! unregister to a moment past the swap.
//!
//! The same sample convicted `RHP_FREEZE=hide`: measured on glass, a hidden source's thumbnail went
//! black **18 to 71 ms** after the hide and stayed black until the swap — gaps of 415, 500, 520, 555
//! and 700 ms across five seeks. No road that freezes the outgoing player can work, and the numbers
//! are in `ARCHITECTURE.md` with the road that replaced it (see `pin_playback::FreezeRoad`).
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

use windows::Win32::Foundation::RECT;

use super::pin_trace::{asked, note_pin_trace};

/// How often the pin is read off the desk: a gap of 400-700 ms is what a relaunch costs, so twenty
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
    /// The player the pin was showing when the measurement began: the one a still was painted from
    /// and a hide took off the screen, whose end is what `outgoing=` reads.
    pid: u32,
    /// The pin's own box on the desk, read while it was still a box.
    area: RECT,
    /// The moment the still was painted and the registration let go — or, on the hide road, the
    /// moment of the hide: every `t=+<N>ms` is counted from it, and so is `covered_gap_ms`.
    started: Instant,
    /// The first sample that found nothing lit, or `None` where the pin never went black.
    black_at: Option<u128>,
    /// Whether a still this app painted was what the pin was meant to be showing: the road's whole
    /// claim, and what `painted=` answers (see `announce`).
    painted: bool,
    /// The moment the picture changed hands, or `None` while the swap is still owed (see
    /// `note_swapped`).
    swapped_at: Option<Instant>,
    /// Whether the result has been written, so one measurement is one `still probe result` line
    /// however many of the ends it reaches (see `announce`).
    result: bool,
}

impl Probe {
    /// **One result line, and one only**: the answer the measurement exists for.
    ///
    /// `painted=yes` is the road proven — a still of this app's was painted and no sample inside the
    /// gap found the pin black — and anything else is written as what it was rather than as a
    /// failure of the paint: a `black_at=` inside the gap is a pin that went black, and a hide road
    /// writes `painted=no` because no still was ever painted for it.
    ///
    /// A measurement that is given up before its swap — a newer relaunch, the pin going, the bound —
    /// writes `note` first and then the same line, so a run is never left with a probe that said
    /// nothing at all.
    fn announce(&mut self, note: Option<&str>) {
        if self.result {
            return;
        }
        self.result = true;
        if let Some(note) = note {
            note_pin_trace!("still probe ended without a swap: {note}");
        }
        note_pin_trace!(
            "still probe result painted={} black_at={} covered_gap_ms={}",
            if self.painted && self.black_at.is_none() {
                "yes"
            } else {
                "no"
            },
            match self.black_at {
                Some(t) => format!("+{t}ms"),
                None => "never".to_string(),
            },
            self.swapped_at
                .unwrap_or_else(Instant::now)
                .duration_since(self.started)
                .as_millis(),
        );
    }
}

/// The relaunch being measured, or nothing between relaunches: one at a time, because a pin has one
/// film on screen and a second seek restarts the measurement rather than stacking one on another.
static PROBE: Mutex<Option<Probe>> = Mutex::new(None);

/// Whether the sampling thread is up, so one probe is one thread however many relaunches it spans.
static RUNNING: AtomicBool = AtomicBool::new(false);

/// Begin measuring a relaunch: the player the pin was showing, and whether a still of this app's was
/// painted under the registration (`true`) or the film is waiting on a source that is off the screen
/// or gone (`false`, the hide road's answer).
///
/// Nothing is measured without `RHP_PIN_TRACE`, and nothing is measured for a pin with no host to be
/// read from — a measurement with no box to sample has nothing to answer.
pub(super) fn start(pid: u32, painted: bool) {
    if !asked() {
        return;
    }
    let Some(host) = super::pin_film_host::film_host() else {
        note_pin_trace!("still probe unavailable pid {pid}: the pin has no host to be read from");
        return;
    };
    // Safety: the host is this app's own window, and what is asked of it is where its client stands
    // on the screen.
    let Some(area) = (unsafe { super::pin_taskbar_mirror::client_area_on_the_screen(host) }) else {
        note_pin_trace!("still probe unavailable pid {pid}: the pin's box is not on the screen");
        return;
    };
    let now = Instant::now();
    let Ok(mut held) = PROBE.lock() else {
        return;
    };
    if let Some(mut older) = held.take() {
        older.announce(Some("a newer relaunch is the one measured"));
    }
    *held = Some(Probe {
        pid,
        area,
        started: now,
        black_at: None,
        painted,
        swapped_at: None,
        result: false,
    });
    drop(held);

    if RUNNING.swap(true, Ordering::AcqRel) {
        return;
    }
    std::thread::spawn(run);
}

/// The picture has changed hands, so the gap is over: the result is written once, here, and the
/// sampling goes on for a moment past it so the picture arriving is seen arriving.
pub(super) fn note_swapped() {
    let Ok(mut held) = PROBE.lock() else {
        return;
    };
    let Some(probe) = held.as_mut() else {
        return;
    };
    if probe.swapped_at.is_some() {
        return;
    }
    probe.swapped_at = Some(Instant::now());
    probe.announce(None);
}

/// Give the measurement up — a newer relaunch, or the pin going — and write what it had.
///
/// **A probe is never left with no answer**: what would otherwise be a thread reading the desktop for
/// a pin that has gone is one result line with the reason beside it.
pub(super) fn forget(why: &str) {
    let Ok(mut held) = PROBE.lock() else {
        return;
    };
    if let Some(mut probe) = held.take() {
        probe.announce(Some(why));
    }
}

/// One sample's record, copied out under the lock so the pixels are read without it held.
fn the_step() -> Option<(u32, RECT, Instant, bool, Option<Instant>)> {
    let held = PROBE.lock().ok()?;
    let probe = held.as_ref()?;
    Some((
        probe.pid,
        probe.area,
        probe.started,
        probe.painted,
        probe.swapped_at,
    ))
}

/// Whether there is nothing left to measure, and the flag says so.
///
/// **The flag is cleared before the probe is looked at, and a measurement that arrived in the gap is
/// taken up rather than left without a thread**: `start` spawns only when it finds the flag down, so a
/// start that found it up just as this thread was finishing would otherwise leave a probe nobody
/// samples. Answers `true` where this thread is done, and `false` where it owns the newer probe.
fn the_measurement_is_over() -> bool {
    RUNNING.store(false, Ordering::Release);
    let waiting = PROBE.lock().ok().is_some_and(|held| held.is_some());
    !(waiting && !RUNNING.swap(true, Ordering::AcqRel))
}

/// The sampling thread: one line a sample, and the pin read off the desk rather than asked of itself.
fn run() {
    loop {
        let Some((pid, area, started, painted, swapped_at)) = the_step() else {
            if the_measurement_is_over() {
                return;
            }
            continue;
        };
        let t = started.elapsed().as_millis();
        let lit = the_pin_is_lit(area);
        // **Both facts are this app's own to answer honestly**: whether the player that was on screen
        // is still running, and whether a registration is still standing in the host. A road that
        // claims a still is proven must be read with `outgoing=ended thumb=unregistered`.
        let outgoing = if super::engine_processes::is_running(pid) {
            "alive"
        } else {
            "ended"
        };
        let thumb = if super::pin_film_thumbnail::a_registration_stands() {
            "registered"
        } else {
            "unregistered"
        };
        let state = if swapped_at.is_some() {
            "swapped"
        } else if painted {
            "painted"
        } else {
            "waiting"
        };
        note_pin_trace!(
            "still probe t=+{t}ms pin lit={lit}% outgoing={outgoing} thumb={thumb} state={state}"
        );

        let done = {
            let Ok(mut held) = PROBE.lock() else {
                RUNNING.store(false, Ordering::Release);
                return;
            };
            let Some(probe) = held.as_mut() else {
                RUNNING.store(false, Ordering::Release);
                return;
            };
            // **Only the gap counts as the gap**: a sample after the swap is the picture the
            // replacement put there, and a film that opens on black is not a pin that went black.
            if probe.swapped_at.is_none() && probe.black_at.is_none() && lit == 0 {
                probe.black_at = Some(t);
            }
            let done = probe
                .swapped_at
                .is_some_and(|at| at.elapsed() >= AFTER_SWAP)
                || probe.started.elapsed() >= THE_BOUND;
            if done {
                if let Some(mut probe) = held.take() {
                    probe.announce(None);
                }
            }
            done
        };
        if done {
            if the_measurement_is_over() {
                return;
            }
            continue;
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

#[cfg(test)]
mod tests {
    use super::*;

    fn a_probe(painted: bool) -> Probe {
        let now = Instant::now();
        Probe {
            pid: 0,
            area: RECT::default(),
            started: now,
            black_at: None,
            painted,
            swapped_at: None,
            result: false,
        }
    }

    /// **The result is written once, and it is the whole of the road's claim**: a still painted and a
    /// pin never read black is `painted=yes black_at=never`, one black sample inside the gap is
    /// `painted=no black_at=+Nms` — and a second end arriving after the first writes nothing.
    #[test]
    fn the_still_probe_answers_once_and_says_what_it_saw() {
        super::super::pin_trace::clear_noted();

        let mut proven = a_probe(true);
        proven.announce(None);
        proven.announce(Some("the pin is going"));
        let lines = super::super::pin_trace::noted_lines();
        assert_eq!(
            lines.len(),
            1,
            "a still that was never read black is the road proven, and one end is one line: \
             {lines:?}"
        );
        assert!(
            lines[0].starts_with("still probe result painted=yes black_at=never covered_gap_ms="),
            "{lines:?}"
        );

        super::super::pin_trace::clear_noted();
        let mut black = a_probe(true);
        black.black_at = Some(142);
        black.swapped_at = Some(Instant::now());
        black.announce(None);
        let lines = super::super::pin_trace::noted_lines();
        assert_eq!(lines.len(), 1, "one result for one measurement: {lines:?}");
        assert!(
            lines[0].starts_with("still probe result painted=no black_at=+142ms covered_gap_ms="),
            "a pin read black inside the gap is not the road proven: {lines:?}"
        );

        super::super::pin_trace::clear_noted();
        let mut hidden = a_probe(false);
        hidden.announce(Some("the pin is going"));
        let lines = super::super::pin_trace::noted_lines();
        assert_eq!(
            lines.len(),
            2,
            "a measurement given up says why and then what it had: {lines:?}"
        );
        assert_eq!(lines[0], "still probe ended without a swap: the pin is going");
        assert!(
            lines[1].starts_with("still probe result painted=no black_at=never"),
            "and a road with no still in it never claims one was painted: {lines:?}"
        );
        super::super::pin_trace::clear_noted();
    }
}

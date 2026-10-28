//! **What the pin is actually showing while a relaunch has no source of its own drawing in it**,
//! read off the desktop rather than off this app's own bookkeeping.
//!
//! A seek's release takes the frame out of the source's own window, paints the pin's host, flushes,
//! lets the registration go and ends the outgoing player at once (see
//! `pin_film_thumbnail::stand_the_still`). **Whether that still is what the pin then shows is not a
//! question this app can answer about itself** — a DWM thumbnail cannot outlive the drawing of its
//! source, so a picture of this app's own is what replaces it — and the pin's own box is sampled as
//! composed on the desk instead: a percentage of lit points, twenty times a second, from the release to
//! a moment past the swap.
//!
//! **Every measurement ends with one result line**, whatever end it reaches: the swap, a newer
//! relaunch, the pin going, or the bound.
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

use windows::Win32::Foundation::{HWND, RECT};

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

/// The points a percentage is counted over, on each side of the grid — the grid a captured frame is
/// gated on as well, so the pin and the still it is covered by are measured the same way.
pub(super) const GRID: u32 = 16;

/// One relaunch's measurement.
struct Probe {
    /// The player the pin was showing when the measurement began: the one a still is painted from and
    /// a hide took off the screen, whose end is what `outgoing=` reads.
    pid: u32,
    /// The host the pin's box is drawn through, as a handle number: what `z=` asks the registration
    /// for, since the window actually drawn is the Shell's to answer and the record is this app's.
    host: isize,
    /// The pin's own box on the desk, read while it was still a box.
    area: RECT,
    /// The moment the measurement began — the release that owes the cover, or the hide — so every
    /// `t=+<N>ms` is counted from it.
    started: Instant,
    /// The moment the still was painted, or `None` while none has been: **the gap that counts is from
    /// here**, because before it the pin is the outgoing player's own picture (see `note_painted`).
    painted_at: Option<Instant>,
    /// The first sample that found nothing lit between the paint and the swap, or `None` where the
    /// pin never went black after the cover went up.
    black_at: Option<u128>,
    /// The first sample that found **the outgoing source's own client** black, or `None` where the
    /// player drew its own picture all the way to the swap — the field that tells a player that
    /// stopped drawing from a picture the compositor stopped drawing (see `the_source_reading`).
    src_first_black_at: Option<u128>,
    /// The last percentage the outgoing source was read at, which is what `src_lit_at_swap` is taken
    /// from at the moment the picture changes hands (see `note_swapped`).
    last_src: Option<u32>,
    /// What the outgoing source was read at when the picture changed hands.
    src_lit_at_swap: Option<u32>,
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
    /// gap found the pin black — and `painted=no` is what it was instead: a `black_at=` inside the gap
    /// is a pin that went black, and no still painted at all is a relaunch on the `off` road.
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
        // **`painted=yes` is the road proven**: a still of this app's went up and no sample between it
        // and the swap found the pin black. `no` is a pin that went black under the cover, or a relaunch
        // that painted nothing — a frame that came back too dark to be a picture.
        let painted = match self.painted_at {
            Some(_) if self.black_at.is_none() => "yes",
            _ => "no",
        };
        note_pin_trace!(
            "still probe result painted={painted} black_at={} src_first_black_at={} src_lit_at_swap={}",
            match self.black_at {
                Some(t) => format!("+{t}ms"),
                None => "never".to_string(),
            },
            match self.src_first_black_at {
                Some(t) => format!("+{t}ms"),
                None => "never".to_string(),
            },
            match self.src_lit_at_swap {
                Some(pct) => format!("{pct}%"),
                None => "-".to_string(),
            },
        );
    }
}

/// The relaunch being measured, or nothing between relaunches: one at a time, because a pin has one
/// film on screen and a second seek restarts the measurement rather than stacking one on another.
static PROBE: Mutex<Option<Probe>> = Mutex::new(None);

/// Whether the sampling thread is up, so one probe is one thread however many relaunches it spans.
static RUNNING: AtomicBool = AtomicBool::new(false);

/// Begin measuring a relaunch: the player the pin was showing.
///
/// It is begun at the **release**, on the way to the capture, so the paint that follows and the swap
/// that ends it are one relaunch's. The gap that counts is the stretch from the paint to the swap
/// (see `note_painted`).
///
/// Nothing is measured without `RHP_PIN_TRACE`, and nothing is measured for a pin with no host to be
/// read from — a measurement with no box to sample has nothing to answer.
pub(super) fn start(pid: u32) {
    if !asked() {
        return;
    }
    // **What the pin's box is drawn through as this app believes it**, written at the release beside
    // the measurement it is read against: a run can then tell a registration that stood from one this
    // app thought stood (see `pin_film_thumbnail::note_the_covering_registration`).
    super::pin_film_thumbnail::note_the_covering_registration();
    let Some(host) = super::pin_film_host::film_host() else {
        announce_unmeasured(pid, "the pin has no host to be read from");
        return;
    };
    // Safety: the host is this app's own window, and what is asked of it is where its client stands
    // on the screen.
    let Some(area) = (unsafe { super::pin_taskbar_mirror::client_area_on_the_screen(host) }) else {
        announce_unmeasured(pid, "the pin's box is not on the screen");
        return;
    };
    let now = Instant::now();
    let Ok(mut held) = PROBE.lock() else {
        return;
    };
    // **A newer relaunch does not swallow the older one's answer**: what the last one measured is
    // written out in full before the record is replaced.
    if let Some(mut older) = held.take() {
        older.announce(Some("a newer relaunch is the one measured"));
    }
    *held = Some(Probe {
        pid,
        host: host.0 as isize,
        area,
        started: now,
        painted_at: None,
        black_at: None,
        src_first_black_at: None,
        last_src: None,
        src_lit_at_swap: None,
        swapped_at: None,
        result: false,
    });
    drop(held);

    if RUNNING.swap(true, Ordering::AcqRel) {
        return;
    }
    std::thread::spawn(run);
}

/// A measurement that has nothing to measure still answers, **once and with the line every other end
/// writes**: a run is never left with a relaunch whose measurement said nothing at all (see
/// `announce`).
fn announce_unmeasured(pid: u32, why: &str) {
    note_pin_trace!("still probe unavailable pid {pid}: {why}");
    note_pin_trace!(
        "still probe result painted=no black_at=never src_first_black_at=never src_lit_at_swap=-"
    );
}

/// The still this measurement was waiting on has been painted: **the gap the probe answers for begins
/// here**, since before it the pin is the outgoing player's own picture and after it a still of this
/// app's own (see `pin_film_thumbnail::stand_the_still`).
pub(super) fn note_painted() {
    let Ok(mut held) = PROBE.lock() else {
        return;
    };
    let Some(probe) = held.as_mut() else {
        return;
    };
    if probe.painted_at.is_some() {
        return;
    }
    probe.painted_at = Some(Instant::now());
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
    // **The last reading taken of the outgoing source before the picture changed hands** is what the
    // swap is answered with: the samples after it are the replacement's (see `the_source_reading`).
    probe.src_lit_at_swap = probe.last_src;
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
fn the_step() -> Option<(u32, isize, RECT, Instant, Option<Instant>, Option<Instant>)> {
    let held = PROBE.lock().ok()?;
    let probe = held.as_ref()?;
    Some((
        probe.pid,
        probe.host,
        probe.area,
        probe.started,
        probe.painted_at,
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
        let Some((pid, host, area, started, painted_at, swapped_at)) = the_step() else {
            if the_measurement_is_over() {
                return;
            }
            continue;
        };
        let t = started.elapsed().as_millis();
        let lit = the_pin_is_lit(area);
        // **The outgoing window's own client is read beside the pin's box**, because the pin going
        // black cannot say which of the two stopped: a source still lit under a black pin is a
        // picture the compositor stopped drawing, and a source read black with it is a player that
        // stopped drawing (see `pin_film_thumbnail::the_sources_window_facts`).
        let source = super::pin_film_thumbnail::the_sources_picture_is_lit(pid);
        let facts = super::pin_film_thumbnail::the_sources_window_facts(pid);
        // **Whose window the box is drawn through, as the record answers it**: the outgoing player is
        // `pid` and any other is `other`, which is a registration re-pointed rather than one gone.
        let drawn =
            super::pin_film_thumbnail::the_standing_source_process(HWND(host as *mut _));
        let z = match drawn {
            Some(drawn) if drawn == pid => "pid",
            Some(_) => "other",
            None => "none",
        };
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
        } else if painted_at.is_some() {
            "painted"
        } else {
            "waiting"
        };
        note_pin_trace!(
            "still probe t=+{t}ms pin lit={lit}% src lit={} vis={} layered={} topmost={} z={z} outgoing={outgoing} thumb={thumb} state={state}",
            the_source_reading(source),
            a_flag(facts.map(|facts| facts.0)),
            a_flag(facts.map(|facts| facts.1)),
            a_flag(facts.map(|facts| facts.2)),
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
            // **Only the covered gap counts as the gap**: the samples before the paint are the
            // outgoing player's own picture and the ones after the swap are the replacement's, and a
            // film that opens on black is not a pin that went black (see `note_painted`).
            let covered = probe.painted_at.is_some() && probe.swapped_at.is_none();
            if covered && probe.black_at.is_none() && lit == 0 {
                probe.black_at = Some(t);
            }
            // **The source's own reading is kept from the release**, not only inside the covered gap:
            // where it first went black is the whole of the answer this field exists for.
            if probe.src_first_black_at.is_none() && source == Some(false) {
                probe.src_first_black_at = Some(t);
            }
            if let Some(pct) = source.map(|lit| u32::from(lit) * 100) {
                probe.last_src = Some(pct);
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

/// The outgoing source's own client as one field: the percentage a reading is worth, and which
/// reading it was — because a window that answered nothing and a window that answered black are two
/// different facts, and a run must never read one as the other.
fn the_source_reading(lit: Option<bool>) -> &'static str {
    match lit {
        Some(true) => "100%|ok",
        Some(false) => "0%|black",
        None => "-|none",
    }
}

/// One window fact as a field: `1` or `0` where there was a window to ask, `-` where there was none.
fn a_flag(asked: Option<bool>) -> &'static str {
    match asked {
        Some(true) => "1",
        Some(false) => "0",
        None => "-",
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn a_probe(painted: bool) -> Probe {
        let now = Instant::now();
        Probe {
            pid: 0,
            host: 0,
            area: RECT::default(),
            started: now,
            painted_at: painted.then_some(now),
            black_at: None,
            src_first_black_at: None,
            last_src: None,
            src_lit_at_swap: None,
            swapped_at: None,
            result: false,
        }
    }

    /// **The result is written once, and it is the whole of the road's claim**: a still painted and a
    /// pin never read black is `painted=yes black_at=never` with the gap measured from the paint, one
    /// black sample inside the gap is `painted=no black_at=+Nms`, and a second end arriving after the
    /// first writes nothing.
    #[test]
    fn the_still_probe_answers_once_and_says_what_it_saw() {
        super::super::pin_trace::clear_noted();

        let mut proven = a_probe(true);
        proven.swapped_at = Some(Instant::now());
        proven.announce(None);
        proven.announce(Some("the pin is going"));
        let lines = super::super::pin_trace::noted_lines();
        assert_eq!(
            lines.len(),
            1,
            "a still that was never read black is the road proven, and one end is one line:              {lines:?}"
        );
        assert!(
            lines[0].starts_with("still probe result painted=yes black_at=never src_first_black_at="),
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
            lines[0].starts_with("still probe result painted=no black_at=+142ms "),
            "a pin read black inside the gap is not the road proven: {lines:?}"
        );

        super::super::pin_trace::clear_noted();
        let mut unpainted = a_probe(false);
        unpainted.announce(Some("the pin is going"));
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

    /// **The one field this measurement was extended for**: a pin read black says nothing about why,
    /// so the outgoing source's own client is written beside it — and the difference between a source
    /// that stayed lit (`100%|ok`, a picture the compositor stopped drawing) and one that went black
    /// with the pin (`0%|black`, a player that stopped drawing) is the whole of what it is read for. A
    /// window that answered nothing is `-|none`, and is never counted as a black one.
    #[test]
    fn the_source_is_read_beside_the_pin_and_says_which_of_the_two_went_black() {
        assert_eq!(the_source_reading(Some(true)), "100%|ok");
        assert_eq!(the_source_reading(Some(false)), "0%|black");
        assert_eq!(the_source_reading(None), "-|none");
        assert_eq!(
            (a_flag(Some(true)), a_flag(Some(false)), a_flag(None)),
            ("1", "0", "-"),
            "a fact no window answered is not a zero"
        );

        // The first black reading of the source is where it is counted, and the last reading before
        // the swap is what the picture changed hands on.
        super::super::pin_trace::clear_noted();
        let mut probe = a_probe(true);
        probe.src_first_black_at = Some(72);
        probe.last_src = Some(0);
        probe.swapped_at = Some(Instant::now());
        probe.src_lit_at_swap = probe.last_src;
        probe.announce(None);
        let lines = super::super::pin_trace::noted_lines();
        let result = lines
            .iter()
            .find(|line| line.starts_with("still probe result"))
            .unwrap_or_else(|| panic!("the source's readings are written with the pin's: {lines:?}"));
        assert!(
            result.ends_with("src_first_black_at=+72ms src_lit_at_swap=0%"),
            "where the source first went black, and what it was at the swap: {result}"
        );
        super::super::pin_trace::clear_noted();
    }
}

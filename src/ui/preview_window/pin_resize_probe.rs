//! Two temporary diagnostics for the one thing a pinned mpv film still does wrong: its box shakes
//! while the edge under the hand is dragged, and nothing structural has explained it.
//!
//! **`RHP_HIDE_FILM_ON_RESIZE`, present, answers "is it the film or the chrome?"** — the film's own
//! window is taken off the screen for the length of a resize drag and put back when the drag ends,
//! and nothing else changes. A pin that resizes smoothly with no film on it has the film's own
//! present to blame, and no arrangement of windows will fix that; one that is still shaky with no
//! film on it has this app's chrome paint to blame. Either answer is decisive, which is why the knob
//! moves no geometry, changes no band, takes no ownership and paints nothing differently (see
//! `hide_the_film_for_a_resize_drag`).
//!
//! **`RHP_RESIZE_TRACE`, present, answers "where did the step's time go?"** — one line per resize
//! step to the pin trace's own file, carrying the box the pointer asked for, the direction the step
//! went in, and the three costs a step pays: the chrome's paint, the film host's placement, and the
//! fit of mpv's child into it. The box the child was asked for is written down with it and looked at
//! on a later tick, because the question there is when the film's window *arrived* rather than how
//! long asking for it took (see `observe_the_film`). **And the line carries the box the child was
//! standing at as the step's own commits ended, with the strip of its host that box did not cover** —
//! the lag itself, which is what the leading ask of `pin_film_host` is measured by (see
//! `measure_the_film` and `the_uncovered_strip`).
//!
//! **Neither is a setting.** Both are presence-style `RHP_*` variables read once, a run with neither
//! set builds no line, hides nothing and pays one thread-local read per paint, and a pin's road is
//! the road it was with neither named. To remove them: delete this module and its `mod` line, and the
//! calls to `begin_step`, `end_step`, `a_step_started`, `note_the_paint`, `note_the_host_placement`,
//! `note_the_child_fit`, `settle_on_the_tick`, `hide_the_film_for_a_resize_drag`,
//! `show_the_film_after_a_resize_drag`, `the_film_is_hidden` and `forget_the_hidden_film` — the
//! compiler names every one of them.

use std::cell::Cell;
use std::fmt::Write as _;

use windows::Win32::UI::WindowsAndMessaging::{GetClientRect, SWP_ASYNCWINDOWPOS};

use super::*;

/// The variable that asks for the pin's film to be taken off the screen for the length of a resize
/// drag.
const HIDE_FILM_VARIABLE: &str = "RHP_HIDE_FILM_ON_RESIZE";
/// The variable that asks for one line per resize step, and for the box the film's window was asked
/// for to be looked at on a later tick.
const RESIZE_TRACE_VARIABLE: &str = "RHP_RESIZE_TRACE";

/// What the two variables asked for, read once: `(hide the film, trace the steps)`.
///
/// One read for the process, because both are facts about the run rather than about a pin: a variable
/// set after the app has started changes nothing, which is what makes two runs comparable and not two
/// halves of one.
static ASKED: Lazy<(bool, bool)> = Lazy::new(|| {
    (
        std::env::var_os(HIDE_FILM_VARIABLE).is_some(),
        std::env::var_os(RESIZE_TRACE_VARIABLE).is_some(),
    )
});

#[cfg(test)]
thread_local! {
    /// The answer a test stands in place of the machine's own (see `stand_the_knobs_in`).
    static STAND_IN: Cell<Option<(bool, bool)>> = const { Cell::new(None) };
}

/// Stand answers for the two diagnostics on this thread in place of the machine's own, or give the
/// machine's own back with `None`: per thread because the tests run at once, and because every road
/// that asks reads it on the thread the test is on.
#[cfg(test)]
pub(super) fn stand_the_knobs_in(answer: Option<(bool, bool)>) {
    STAND_IN.with(|held| held.set(answer));
}

/// What the two variables asked for, as the machine's own answer unless a test stood one in.
fn asked() -> (bool, bool) {
    #[cfg(test)]
    if let Some(answer) = STAND_IN.with(|held| held.get()) {
        return answer;
    }
    *ASKED
}

/// A box as a window's rectangle is written, in screen coordinates: left, top, width, height.
pub(super) type ScreenBox = (i32, i32, i32, i32);

/// One resize step as it is measured: the box the pointer asked for, the direction it went in, the
/// three costs it paid, and the box the film's window was asked for.
pub(super) struct Step {
    pub(super) box_: ScreenRegion,
    pub(super) grew: bool,
    /// The chrome's own paint: `render_pinned_preview_at` and the `UpdateLayeredWindow` in it.
    pub(super) paint: Option<Duration>,
    /// Placing the film host at the band, which is this app's own `SetWindowPos`.
    pub(super) host: Option<Duration>,
    /// Fitting mpv's child into the host, which is the one cost of a step that crosses to another
    /// process — and the one the box of the film is waiting on.
    pub(super) fit: Option<Duration>,
    /// Whether that fit asked for the box asynchronously rather than waiting for the box to be taken.
    pub(super) fit_is_async: bool,
    /// The film window the fit asked for a box, that box in screen coordinates, and the instant it was
    /// asked for: what a later tick looks at (see `observe_the_film`).
    pub(super) asked_for: Option<(HWND, ScreenBox, Instant)>,
    /// The host and the child the step fitted, so that the child's own box and the strip of its host it
    /// leaves uncovered are read when the step's commits have all run (see `measure_the_film`).
    pub(super) filmed: Option<(HWND, HWND)>,
    /// The child's box as the step left it, and the strip of the host that box does not cover: the lag
    /// itself, measured at the end of the step rather than argued about (see `the_uncovered_strip`).
    pub(super) uncovered: Option<(ScreenBox, (i32, i32))>,
}

thread_local! {
    /// The step being measured, while a resize step is in flight and the trace asked for one.
    static STEP: RefCell<Option<Step>> = const { RefCell::new(None) };
    /// The film window this diagnostic hid, so that the drag's end can put it back.
    static HIDDEN_FILM: Cell<isize> = const { Cell::new(0) };
    /// The box a step asked the film's window for, waiting for a later tick to look at it.
    static OBSERVATION: RefCell<Option<Observation>> = const { RefCell::new(None) };
}

/// One step as it begins: the box and the direction known, every cost still to be measured.
pub(super) fn a_step_at(box_: ScreenRegion, grew: bool) -> Step {
    Step {
        box_,
        grew,
        paint: None,
        host: None,
        fit: None,
        fit_is_async: false,
        asked_for: None,
        filmed: None,
        uncovered: None,
    }
}

/// Begin measuring one resize step, at the box the pointer asked for and the direction it moved in.
/// Nothing at all is stored where `RHP_RESIZE_TRACE` was not set.
pub(super) fn begin_step(box_: ScreenRegion, grew: bool) {
    if !asked().1 {
        return;
    }
    STEP.with(|step| *step.borrow_mut() = Some(a_step_at(box_, grew)));
}

/// Close the step: one line to the pin trace's own file, and the box it asked the film's window for
/// left to a later tick to look at.
///
/// **The child's own box is read here, at the step's end and not when it was asked for**, because the
/// ask is the thing that lands late: a step whose numbers were taken at the ask would say the film is
/// where it was told to be before mpv's thread had a chance to move it.
pub(super) fn end_step() {
    let Some(mut step) = STEP.with(|step| step.borrow_mut().take()) else {
        return;
    };
    step.uncovered = step.filmed.and_then(|(host, child)| measure_the_film(host, child));
    pin_trace::write(&a_step_line(&step));
    if let Some((child, want, since)) = step.asked_for {
        OBSERVATION.with(|held| {
            *held.borrow_mut() = Some(Observation {
                child,
                want,
                since,
                tries: 0,
            })
        });
    }
}

/// The instant something inside a step began, or nothing where no step is being measured: the one
/// question the paint and the two placements ask before they pay for an `Instant`.
pub(super) fn a_step_started() -> Option<Instant> {
    STEP.with(|step| step.borrow().as_ref().map(|_| Instant::now()))
}

/// Note how long the chrome's paint took, where one is being measured.
pub(super) fn note_the_paint(started: Option<Instant>) {
    if let Some(started) = started {
        note_in_the_step(|step| step.paint = Some(started.elapsed()));
    }
}

/// Note how long placing the film host took, where a step is being measured.
pub(super) fn note_the_host_placement(started: Option<Instant>) {
    if let Some(started) = started {
        note_in_the_step(|step| step.host = Some(started.elapsed()));
    }
}

/// Note the fit of the film's window into the host, the box it asked for, and how it asked for it: the
/// one measurement a later tick has to finish (see `observe_the_film`).
pub(super) fn note_the_child_fit(
    host: HWND,
    child: HWND,
    rect: (i32, i32, i32, i32),
    flags: SET_WINDOW_POS_FLAGS,
    started: Instant,
) {
    note_in_the_step(|step| {
        step.fit = Some(started.elapsed());
        step.fit_is_async = flags.contains(SWP_ASYNCWINDOWPOS);
        step.asked_for = Some((child, the_screen_box(host, rect), started));
        step.filmed = Some((host, child));
    });
}

/// Note one measurement in the step that is being measured, and do nothing at all where none is.
fn note_in_the_step(note: impl FnOnce(&mut Step)) {
    STEP.with(|step| {
        if let Some(step) = step.borrow_mut().as_mut() {
            note(step);
        }
    });
}

/// The line one step writes, built once and handed to the trace's own file: the box the pointer asked
/// for, the direction it went in, and each cost the step paid.
pub(super) fn a_step_line(step: &Step) -> String {
    let (left, top, right, bottom) = step.box_;
    let mut line = String::with_capacity(160);
    let _ = write!(
        line,
        "resize step box {left},{top},{}x{} {} paint ",
        right - left,
        bottom - top,
        if step.grew { "grew" } else { "shrank" },
    );
    write_a_cost(&mut line, step.paint);
    line.push_str(" host ");
    write_a_cost(&mut line, step.host);
    line.push_str(" fit ");
    write_a_cost(&mut line, step.fit);
    if step.fit.is_some() {
        line.push_str(if step.fit_is_async { " async" } else { " sync" });
    }
    match step.asked_for {
        Some((child, want, _)) => {
            let _ = write!(
                line,
                " child {child:?} want {},{} {}x{}",
                want.0, want.1, want.2, want.3
            );
        }
        None => line.push_str(" child none"),
    }
    // **What the film's window was standing at as the step ended, and the strip of its host it did not
    // cover** — the number the lag is: what the host has to paint in the strip, whether that is the
    // frame a drag holds, the brush, or the live picture of a child held over the band.
    if let Some((child, (strip_w, strip_h))) = step.uncovered {
        let _ = write!(
            line,
            " now {},{} {}x{} strip {strip_w}x{strip_h}",
            child.0, child.1, child.2, child.3
        );
    }
    line
}

/// The child's own box and the strip of its host it does not cover, read at the end of a step: three
/// reads of rectangles, on windows of this app's own, so this is no message to anybody and no wait on
/// the player's thread (see `observe_the_film`, which reads the same fact a tick later).
fn measure_the_film(host: HWND, child: HWND) -> Option<(ScreenBox, (i32, i32))> {
    let mut client = RECT::default();
    let mut got = RECT::default();
    let mut origin = POINT::default();
    unsafe {
        GetClientRect(host, &mut client).ok()?;
        GetWindowRect(child, &mut got).ok()?;
        let _ = ClientToScreen(host, &mut origin);
    }
    let band = (origin.x, origin.y, client.right, client.bottom);
    let child_box = (got.left, got.top, got.right - got.left, got.bottom - got.top);
    Some((child_box, the_uncovered_strip(band, child_box)))
}

/// How much of a `band` box a `child` box does not cover: the width of the strip left over on the
/// right and the height of the strip left over at the bottom, each zero where the child covers the
/// band and each the whole extent where the child is not in it at all.
///
/// **Two strips and not a difference of areas**, because that is the shape of what the host paints: it
/// fills its whole client box, and a child that is a step behind on two edges leaves a band on each.
pub(super) fn the_uncovered_strip(band: ScreenBox, child: ScreenBox) -> (i32, i32) {
    let covered = |start: i32, extent: i32, from: i32, size: i32| {
        ((start + extent).min(from + size) - start.max(from)).max(0)
    };
    (
        (band.2 - covered(band.0, band.2, child.0, child.2)).max(0),
        (band.3 - covered(band.1, band.3, child.1, child.3)).max(0),
    )
}

/// One cost as the line carries it, in milliseconds, or a dash where the step did not pay it.
fn write_a_cost(line: &mut String, took: Option<Duration>) {
    match took {
        Some(took) => {
            let _ = write!(line, "{:.3}ms", took.as_secs_f64() * 1000.0);
        }
        None => line.push('-'),
    }
}

/// The box a child window of `host` is asked for when it is given `rect` in the host's own client
/// coordinates — the origin as well as the size, because a child held over its band is asked for a
/// negative one: what a later tick compares `GetWindowRect` against, which answers in screen
/// coordinates for a child as well as for a window of the desktop's.
fn the_screen_box(host: HWND, rect: (i32, i32, i32, i32)) -> ScreenBox {
    let mut origin = POINT::default();
    let _ = unsafe { ClientToScreen(host, &mut origin) };
    (
        origin.x + rect.0,
        origin.y + rect.1,
        rect.2,
        rect.3,
    )
}

/// What a step asked the film's window for, and what a later tick has to say about it.
struct Observation {
    child: HWND,
    want: ScreenBox,
    since: Instant,
    tries: u32,
}

/// How many ticks the box a step asked for is waited for before the window is written down as still
/// not there: a bound rather than a wait, and the last look is written either way.
const OBSERVATION_TICKS: u32 = 8;

/// Look once at the box the film's window is standing at.
///
/// **On a tick and never in the step**, which is the whole of what makes the number mean anything: the
/// fit a step pays for is asynchronous on the road a pin's film is on, so the box it names arrives
/// later, and how much later is the measurement. One `GetWindowRect` per tick — a read of a stored
/// rectangle rather than a message to anybody — and the look is given up after `OBSERVATION_TICKS` of
/// them rather than waited on.
fn observe_the_film() {
    let Some(mut observation) = OBSERVATION.with(|held| held.borrow_mut().take()) else {
        return;
    };
    observation.tries += 1;

    let mut got = RECT::default();
    let read = unsafe { GetWindowRect(observation.child, &mut got) };
    let found = read.is_ok().then(|| {
        (
            got.left,
            got.top,
            got.right - got.left,
            got.bottom - got.top,
        )
    });
    let off = found.map(|got| the_delta(observation.want, got));
    if off != Some((0, 0, 0, 0)) && observation.tries < OBSERVATION_TICKS {
        OBSERVATION.with(|held| *held.borrow_mut() = Some(observation));
        return;
    }

    let mut line = String::with_capacity(140);
    let _ = write!(line, "resize observe child {:?}", observation.child);
    let want = observation.want;
    match (found, off) {
        (Some(got), Some((dx, dy, dw, dh))) => {
            let _ = write!(
                line,
                " want {},{} {}x{} got {},{},{}x{} off {dx},{dy},{dw},{dh}",
                want.0, want.1, want.2, want.3, got.0, got.1, got.2, got.3
            );
        }
        _ => line.push_str(" gone"),
    }
    let _ = write!(
        line,
        " after {:.3}ms on {} ticks",
        observation.since.elapsed().as_secs_f64() * 1000.0,
        observation.tries
    );
    pin_trace::write(&line);
}

/// How far the box the film's window was asked for is from the one it was found at, on each of the
/// four edges a box is written in: left, top, width, height.
pub(super) fn the_delta(want: ScreenBox, got: ScreenBox) -> ScreenBox {
    (
        got.0 - want.0,
        got.1 - want.1,
        got.2 - want.2,
        got.3 - want.3,
    )
}

/// Take the pin's film off the screen for the length of a resize drag, where
/// `RHP_HIDE_FILM_ON_RESIZE` asked for it.
///
/// **What is hidden is the film's own window and not the host it stands in**, which is the whole of
/// why this is not the park's road: on the road a pin's mpv film is drawn on, the host *owns* the
/// pin's chrome, so hiding the host hides the pin — and a drag with nothing on the screen answers the
/// question this exists for with nothing at all. The window taken off the screen is the child the
/// player draws into, so the chrome goes on moving over the host's own black band and the only thing
/// this changes is whether the film is seen.
///
/// **Posted rather than waited for**, `ShowWindowAsync` as the hover's own host road uses it: the
/// window belongs to mpv's process, and a synchronous show is this thread waiting on mpv's.
pub(super) fn hide_the_film_for_a_resize_drag() {
    if !asked().0 {
        return;
    }
    let Some(host) = pin_film_host::film_host() else {
        return;
    };
    let Some(pid) = pin_film_host::hosted_player_pid() else {
        return;
    };
    let Some(child) = pin_film_host::a_child_of_the_host(host, pid) else {
        return;
    };
    if HIDDEN_FILM.with(|hidden| hidden.get()) == child.0 as isize {
        return;
    }
    let _ = unsafe { ShowWindowAsync(child, SW_HIDE) };
    HIDDEN_FILM.with(|hidden| hidden.set(child.0 as isize));
    if asked().1 {
        pin_trace::write(&format!("resize film hidden {child:?}"));
    }
}

/// Put the film back where the drag left it.
///
/// The box is not this call's to ask for: the fit ran on every step the film was off the screen for
/// (see `the_film_is_hidden`, which is what keeps it off), so the window is already the band's size
/// and what is owed here is only the show. A window that has gone with its player is forgotten rather
/// than shown: the flag may not outlive it, because the fit of the *next* film asks it (see
/// `fit_flags`).
pub(super) fn show_the_film_after_a_resize_drag() {
    let hidden = HIDDEN_FILM.with(|hidden| hidden.replace(0));
    if hidden == 0 {
        return;
    }
    let child = HWND(hidden as *mut _);
    if unsafe { IsWindow(child) }.as_bool() {
        let _ = unsafe { ShowWindowAsync(child, SW_SHOWNOACTIVATE) };
        if asked().1 {
            pin_trace::write(&format!("resize film shown {child:?}"));
        }
    }
}

/// Forget the film this diagnostic hid, without showing it: the window it named is going away with
/// the host that held it, and nothing may be left saying a film is off the screen that no longer
/// exists (see `take_down_the_pin_film_host`).
pub(super) fn forget_the_hidden_film() {
    HIDDEN_FILM.with(|hidden| hidden.set(0));
}

/// Whether this diagnostic has the film off the screen, which is what the fit that follows the host's
/// own box asks before it shows anything: a fit that showed the child is a diagnostic that lasts one
/// step (see `fit_flags`).
pub(super) fn the_film_is_hidden() -> bool {
    HIDDEN_FILM.with(|hidden| hidden.get() != 0)
}

/// Settle both diagnostics on a tick that is not a resize step: the film comes back once the drag is
/// over — whichever way it ended, which is the whole reason this is asked of the tick rather than of
/// the release — and the box the last step asked the film's window for is looked at (see
/// `observe_the_film`).
pub(super) fn settle_on_the_tick(dragging: bool) {
    let (hides, traces) = asked();
    if !hides && !traces {
        return;
    }
    if hides && !dragging {
        show_the_film_after_a_resize_drag();
    }
    if traces {
        observe_the_film();
    }
}

/// Drop the step a test opened, so a thread the suite reuses is not left measuring one.
#[cfg(test)]
pub(super) fn forget_the_step() {
    STEP.with(|step| *step.borrow_mut() = None);
}

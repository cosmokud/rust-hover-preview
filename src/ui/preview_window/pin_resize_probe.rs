//! One temporary diagnostic for a resize step: **`RHP_RESIZE_TRACE`, present, answers "where did the
//! step's time go?"**
//!
//! One line per resize step to the pin trace's own file, carrying the box the pointer asked for, the
//! direction the step went in, and the two costs the surviving film road pays for it: the chrome's
//! paint and the film host's placement. The child fit and the strip it left uncovered were the third
//! and fourth numbers, and they went with the child road — nothing of the player's is asked for a box
//! across a step any more, so there is nothing there to measure.
//!
//! **Not a setting.** The variable is presence-style and read once, a run with it unset builds no
//! line and pays one thread-local read per paint, and a pin's road is the road it was with it unset.
//! To remove it: delete this module and its `mod` line, and the calls to `begin_step`, `end_step`,
//! `a_step_started`, `note_the_paint` and `note_the_host_placement` — the compiler names every one
//! of them.

#[cfg(test)]
use std::cell::Cell;
use std::fmt::Write as _;

use super::*;

/// The variable that asks for one line per resize step.
const RESIZE_TRACE_VARIABLE: &str = "RHP_RESIZE_TRACE";

/// Whether the trace was asked for, read once: a fact about the run rather than about a pin, so a
/// variable set after the app has started changes nothing and two runs stay comparable.
static ASKED: Lazy<bool> = Lazy::new(|| std::env::var_os(RESIZE_TRACE_VARIABLE).is_some());

#[cfg(test)]
thread_local! {
    /// The answer a test stands in place of the machine's own (see `stand_the_trace_in`).
    static STAND_IN: Cell<Option<bool>> = const { Cell::new(None) };
}

/// Stand an answer for the trace on this thread in place of the machine's own, or give the machine's
/// own back with `None`: per thread because the tests run at once.
#[cfg(test)]
pub(super) fn stand_the_trace_in(answer: Option<bool>) {
    STAND_IN.with(|held| held.set(answer));
}

/// Whether the machine asked for the trace, unless a test stood an answer in.
fn asked() -> bool {
    #[cfg(test)]
    if let Some(answer) = STAND_IN.with(|held| held.get()) {
        return answer;
    }
    *ASKED
}

/// One resize step as it is measured: the box the pointer asked for, the direction it went in, and
/// the two costs the surviving road pays.
pub(super) struct Step {
    pub(super) box_: ScreenRegion,
    pub(super) grew: bool,
    /// The chrome's own paint: `render_pinned_preview_at` and the `UpdateLayeredWindow` in it.
    pub(super) paint: Option<Duration>,
    /// Placing the film host at the band, which is this app's own `SetWindowPos`.
    pub(super) host: Option<Duration>,
}

thread_local! {
    /// The step being measured, while a resize step is in flight and the trace asked for one.
    static STEP: RefCell<Option<Step>> = const { RefCell::new(None) };
}

/// One step as it begins: the box and the direction known, every cost still to be measured.
pub(super) fn a_step_at(box_: ScreenRegion, grew: bool) -> Step {
    Step {
        box_,
        grew,
        paint: None,
        host: None,
    }
}

/// Begin measuring one resize step, at the box the pointer asked for and the direction it moved in.
/// Nothing at all is stored where `RHP_RESIZE_TRACE` was not set.
pub(super) fn begin_step(box_: ScreenRegion, grew: bool) {
    if !asked() {
        return;
    }
    STEP.with(|step| *step.borrow_mut() = Some(a_step_at(box_, grew)));
}

/// Close the step: one line to the pin trace's own file, with the costs the step paid.
pub(super) fn end_step() {
    let Some(step) = STEP.with(|step| step.borrow_mut().take()) else {
        return;
    };
    pin_trace::write(&a_step_line(&step));
}

/// The instant something inside a step began, or nothing where no step is being measured: the one
/// question the paint and the placement ask before they pay for an `Instant`.
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
    let mut line = String::with_capacity(96);
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
    line
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

/// Drop the step a test opened, so a thread the suite reuses is not left measuring one.
#[cfg(test)]
pub(super) fn forget_the_step() {
    STEP.with(|step| *step.borrow_mut() = None);
}

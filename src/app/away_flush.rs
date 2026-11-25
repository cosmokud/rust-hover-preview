//! Giving back what a sweep of a folder left in memory, once the user has been away long enough —
//! the `Performance → Memory → Respect Away Timer` row.
//!
//! **`Engine → Away Timer` was a rule about processes and is now a rule about memory too.** An
//! engine set to `Until away` is a process let go once no Explorer window has been reachable for
//! that long (`app::afk`), which is the app deciding it is no longer being used. The caches a hover
//! leaves behind are the other half of that: a decoded frame, a run of an animation, a parsed page,
//! an archive's listing are all held for a hover nobody is there to make, and every one of them is
//! bounded by a budget under `Memory` that says how much may be held — not how long. With this row
//! on, the same moment the engines are let go is the one those caches are given back at.
//!
//! **It fires once per away period and not on every tick.** What it is for is a machine left alone:
//! nothing accumulates while nobody is at the keyboard, so the flush that matters is the first, and
//! the latch is re-armed the moment Explorer is reachable again. A flush on every look would be the
//! app doing work for the whole of an evening to keep a cache that is already empty.
//!
//! **Nothing on screen is ever taken by it.** A window of this app's that is up — pinned, minimized
//! into the taskbar, put away as a bubble, or detached into a process of its own — is a window its
//! process keeps drawing, and what it is drawing is held by the window rather than by these caches;
//! but a window that is *open* is also a window the user can come back to in a second, and the
//! frames it may be shown next are the ones this would take. So a process with anything on screen
//! flushes nothing, and — because the user is still away — keeps the flush it owed rather than
//! losing it: the latch lives in the process that holds the clock, and the flush is made on the
//! first tick after that process is empty.
//!
//! The predicate is a function of its own rather than an `if` inside the tick for the reason this
//! app writes every such rule down: a rule about what may be given back can only be tested by
//! giving something back, and the ticks that would test it are the ones with a window on screen.

use std::sync::atomic::{AtomicBool, Ordering};

/// Whether this process has already flushed for the away period it is in.
///
/// One per process, because one per process is what flushes: a viewer that has become a window of
/// its own holds its own caches, and the tray's own preview thread holds its own while it draws the
/// hovers. A viewer begun *after* the flush starts with nothing held, so the latch is only ever
/// about the one process it lives in.
static FLUSHED_THIS_AWAY: AtomicBool = AtomicBool::new(false);

/// Whether a process that is `away`, asks for the flush (`respects_the_timer`), has nothing on
/// screen (`something_is_on_screen`) and has not yet flushed for this away period owes the flush.
///
/// Four answers, and each of them is one road out:
///
/// * a user who is present owes nothing — the whole of what a cache is for is the next hover;
/// * a row turned off owes nothing, which is what makes this the same app it was before the row
///   existed;
/// * a window that is up owes nothing *and keeps what it owes*, because the away period is still
///   the one it is in and the first empty tick of it is still to come;
/// * and a period that has already been flushed for owes nothing, because nothing has been held
///   since.
pub(crate) fn the_away_flush_is_owed(
    away: bool,
    respects_the_timer: bool,
    something_is_on_screen: bool,
    already_flushed: bool,
) -> bool {
    away && respects_the_timer && !something_is_on_screen && !already_flushed
}

/// Whether this process owes the flush now, asked of the run's own away clock and the file
/// (see `the_away_flush_is_owed`).
///
/// The setting is read from the configuration each time rather than captured, the way every budget
/// and idle time here is: a row clicked while the machine is away is honoured on the next tick
/// rather than at the next start.
pub(crate) fn the_flush_is_owed(something_is_on_screen: bool) -> bool {
    let respects_the_timer = crate::CONFIG
        .lock()
        .map(|config| config.respect_away_timer)
        .unwrap_or(crate::config::config::DEFAULT_RESPECT_AWAY_TIMER);

    the_away_flush_is_owed(
        crate::app::afk::the_run_is_away(),
        respects_the_timer,
        something_is_on_screen,
        FLUSHED_THIS_AWAY.load(Ordering::Acquire),
    )
}

/// Give back every cache this flush is about, and note that this away period has been answered.
///
/// What is given back is everything under `Performance → Memory` that is *held* rather than being
/// drawn: the decoded stills, the run a pinned picture was walked into and the one a drag is
/// kept in, the timings of a picture that plays, and the three caches under `Others`. The metadata
/// memos beside them — a film's probed geometry, a file's content type, a sound's track, an
/// engine's refusal — are left where they are: they are tens of kilobytes rather than tens of
/// megabytes, and the refusal markers are the one thing here whose loss costs a process launch
/// rather than a read.
pub(crate) fn flush_the_caches() {
    crate::ui::preview_window::flush_the_memory_caches();
    crate::others_cache::flush();
    crate::formats::animation_frames::clear_timelines();

    FLUSHED_THIS_AWAY.store(true, Ordering::Release);
}

/// Note that a flush is owed again — which is what a user coming back is.
///
/// Called from the tick that finds Explorer reachable: the away period is over, whatever it
/// answered has been spent, and the next one is a flush of its own.
pub(crate) fn the_user_is_back() {
    FLUSHED_THIS_AWAY.store(false, Ordering::Release);
}

#[cfg(test)]
mod tests {
    use super::*;

    /// **All four answers are needed for the flush to be owed**, and each of them on its own is
    /// enough to say it is not — which is the whole of what the row means.
    #[test]
    fn the_flush_needs_the_timer_the_row_the_empty_screen_and_a_period_not_yet_answered() {
        assert!(the_away_flush_is_owed(true, true, false, false));

        for (away, respects, on_screen, flushed) in [
            (false, true, false, false),
            (true, false, false, false),
            (true, true, true, false),
            (true, true, false, true),
        ] {
            assert!(
                !the_away_flush_is_owed(away, respects, on_screen, flushed),
                "away={away}, respects={respects}, on_screen={on_screen}, flushed={flushed} \
                 owes no flush"
            );
        }
    }

    /// **A window that is up keeps the flush it owes rather than losing it.** The user is still
    /// away, so the period is still the one this process is in: the tick that takes the window
    /// down is the one that flushes, and a process that answered "nothing to do" because the
    /// screen was not empty would go on holding a folder's worth of frames for the evening.
    #[test]
    fn a_window_that_is_up_owes_the_flush_when_it_goes() {
        assert!(!the_away_flush_is_owed(true, true, true, false));
        assert!(
            the_away_flush_is_owed(true, true, false, false),
            "and it is owed the moment the screen is empty"
        );
    }

    /// **One flush per away period**, which is what the latch is: a period already answered owes
    /// nothing more until the user comes back and leaves again.
    #[test]
    fn one_flush_is_made_per_away_period() {
        the_user_is_back();
        assert!(!FLUSHED_THIS_AWAY.load(Ordering::Acquire));

        FLUSHED_THIS_AWAY.store(true, Ordering::Release);
        assert!(
            !the_away_flush_is_owed(true, true, false, FLUSHED_THIS_AWAY.load(Ordering::Acquire)),
            "a period already answered owes nothing"
        );

        the_user_is_back();
        assert!(
            the_away_flush_is_owed(true, true, false, FLUSHED_THIS_AWAY.load(Ordering::Acquire)),
            "and a user who came back and left again owes the next one"
        );

        the_user_is_back();
    }
}

//! How long Explorer has been out of reach, which is what bounds an engine that is not
//! marked `Persistent` — the `Engine → AFK Timer` setting.
//!
//! The count is kept here rather than by the engines because only the Explorer hook can
//! answer it, and it answers it already: out of reach is every Explorer window minimized, or
//! every one of them behind the region a maximized or fullscreen window in front covers, on
//! any display — the same classification the hook makes for its own sleeps, which is where
//! the multimonitor rule lives (`ExplorerState`). So the hook writes what it found once a
//! pass and the engines read it on the look they already take, and nothing here polls, wakes
//! anything or starts a thread of its own: this decides when an engine that is not being used
//! is let go, and nothing else.
//!
//! What it is not is a measure of input. A user who is sitting still in Explorer is a user
//! whose engines are being kept for hovers they may yet make, and a mouse that has not moved
//! says nothing about the window in front of it; the window state above says all of it.
//!
//! Nothing is written to disk and nothing outlives the run. A run starts with the user
//! present, and whatever is warm when it ends is ended by the run itself.

use std::sync::Mutex;
use std::time::{Duration, Instant};

use once_cell::sync::Lazy;

use crate::config::config::DEFAULT_AFK_TIMER_SECS;
use crate::CONFIG;

/// When the last Explorer window stopped being reachable, or `None` while one is.
///
/// `None` is also how a run starts and how a machine whose hook has not looked yet reads:
/// the app never lets an engine go on a clock it has not watched start.
static AWAY_SINCE: Lazy<Mutex<Option<Instant>>> = Lazy::new(|| Mutex::new(None));

/// Record whether any Explorer window is reachable, as the hook found it.
///
/// The instant is kept across calls rather than restarted on every one, so what the clock
/// measures is the whole time away and not the time since the last look.
pub(crate) fn note_explorer_reachable(reachable: bool) {
    if let Ok(mut away_since) = AWAY_SINCE.lock() {
        if reachable {
            *away_since = None;
        } else if away_since.is_none() {
            *away_since = Some(Instant::now());
        }
    }
}

/// Whether Explorer has been out of reach for longer than `afk_timer_seconds` names.
///
/// The setting is read from the configuration each time rather than captured, so an edit
/// applies to engines that are already warm — the same reason the idle times are read live.
/// A run that has not been told anything yet, or a user who is present, has not been away:
/// both answer `false`.
pub(crate) fn expired() -> bool {
    let since = AWAY_SINCE.lock().ok().and_then(|away_since| *away_since);

    let Some(since) = since else {
        return false;
    };

    let seconds = CONFIG
        .lock()
        .map(|config| config.afk_timer_seconds)
        .unwrap_or(DEFAULT_AFK_TIMER_SECS);

    since.elapsed() >= Duration::from_secs(seconds)
}

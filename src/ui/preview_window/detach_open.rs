//! Detached viewer exit latch: ends a detached viewer process once its pin closes.

use std::sync::atomic::{AtomicBool, Ordering};

use super::requests::pinned;
use crate::app::detach;

/// Latches true once a detached viewer's pin has been observed up, so closing it exits the process.
static THIS_WINDOW_WAS_UP: AtomicBool = AtomicBool::new(false);

/// End a detached viewer process once its standalone pin closes.
pub(super) fn settle() {
    if !detach::is_detached() {
        return;
    }

    if pinned() {
        THIS_WINDOW_WAS_UP.store(true, Ordering::Release);
    } else if THIS_WINDOW_WAS_UP.load(Ordering::Acquire) {
        detach::the_process_may_go();
    }
}

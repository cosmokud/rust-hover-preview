//! The open a detached window begins with, and the two facts its own tick owes.
//!
//! A detached window is this app in a mode where the file is not *hovered* and not *picked* — it is
//! handed over by the tray app that was standing in the window this one becomes (see
//! `crate::app::detach`). So it is opened the way every window of this app's is opened: a preview of
//! it is put up, and what is on screen is then taken up as a pin. That is not a shortcut taken for
//! convenience — it is the only road that draws every kind this app shows. A film needs its player
//! begun, a document needs an engine asked for its page and waited on, a model needs its viewer —
//! and all of that is the hover's own work, driven by the loop's own state, which a hand-off that
//! read a file and built a window by itself would have to reproduce a second time, kind by kind.
//!
//! The hover is the one part of that road a detached window does not show: it is nobody's preview,
//! and a card appearing beside the window the user just asked to detach is a card they did not ask
//! for. So its window is kept off the screen while it loads — moved past the corner of every display,
//! which is a position rather than a hidden window, because the preview is laid out and painted
//! against the box it was given and hiding it would be a window the loop has stopped drawing.
//!
//! What is left over is what the tick below answers, and both answers are the same fact seen from
//! its two sides:
//!
//! - **The window is up.** Not when the message was sent and not when the pin's state was written:
//!   when the window has been drawn. Only then does the tray app's own pin come down, so the file is
//!   never in neither window and the two are never both on the screen for longer than the time
//!   between the two being drawn. The maximize the hand-off carried is put back before that answer
//!   is given, for the same reason — a window that comes up at its restore size and *then* fills the
//!   screen is a window the user watches jump.
//! - **The window is gone.** A detached window is not a process with a tray icon and an Exit row: its
//!   window *is* it, so the moment that window ends this process has nothing left to be. The end is
//!   the pin's own end — the caption's Close, Escape, the menu — and it is read here rather than
//!   written into `end_pin`, because the loop that draws the window is the only thing that can end
//!   the process that owns it.
//!
//! **A viewer whose pin was detached answers the second of those and nothing else.** It was handed no
//! file, so there is no open of a window to wait on: the one window it has is the pin it was already
//! drawing, and the latch below is what says that window has been on the screen — so its going is this
//! process's own end, and a closed window does not leave a process behind `main` would keep forever.
//!
//! Nothing here runs in the tray app, and nothing of the first half runs in a viewer: `is_detached` is
//! false there and the whole module is a function call that answers at once.

use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Mutex;
use std::time::{Duration, Instant};

use once_cell::sync::Lazy;
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{
    SetWindowPos, SWP_NOACTIVATE, SWP_NOSIZE, SWP_NOZORDER,
};

use super::media_types::PreviewMessage;
use super::pin_geometry::PinFrame;
use super::pin_lifecycle::pin_what_is_on_screen;
use super::pin_window::{ask_pin, pin_state};
use super::requests::pinned;
use super::tick::{send_preview, PREVIEW_HWND};
use super::PinCommand;
use crate::app::detach;

/// How many ticks a pin has to have been up before the window is taken to be drawn.
///
/// The take-up itself draws the window (a take-up is a paint), so the tick after it is a window that
/// is there; the second is the margin for a tick whose paint the paint path answered from a frame
/// already composed. What it costs is two frames of the loop's own cadence, which is the smallest
/// wait in this module and the one a hand cannot see.
const PAINTED_BEFORE_TELLING: u32 = 2;

/// Where a preview nobody is to see is kept: past the corner of every display, which is the place
/// a window goes when it is a window rather than a session (see `pin_taskbar`).
const OFF_THE_SCREEN: i32 = -32000;

/// How long the file is given to arrive before the window stops waiting for it.
///
/// A document handed to an Office engine spends up to fifteen seconds being rendered and still
/// arrives, so this is generous rather than tight. What it is for is the case that would otherwise
/// leave a process of this app's standing about for ever with no window in it: a file this build
/// cannot show at all, and a preview the configuration has switched off. The pin it was handed over
/// from stays either way — it comes down only on the answer this gives (see
/// `crate::app::detach::wait_for_the_hand_over`).
const OPEN_WAIT: Duration = Duration::from_secs(90);

/// And how long nothing at all is allowed to be on screen before that is read as the same answer:
/// a preview that was asked for and answered with nothing leaves neither a load in flight nor
/// anything on screen, and nothing more is coming.
const NOTHING_IS_COMING_FOR: u32 = 180;

/// Where a detached window's open has got to. One process has one of these because one process has
/// one window.
#[derive(Default)]
struct Opening {
    /// Whether the preview has been asked for.
    begun: bool,
    /// When it was asked for, which is the clock `OPEN_WAIT` is measured on.
    since: Option<Instant>,
    /// Whether the preview has been taken up as a pin.
    asked: bool,
    /// How many ticks the pin has been up for.
    up_for: u32,
    /// Whether the maximize the hand-off carried has been dealt with — asked for, or not this
    /// window's to ask (a sound's card is not a window that can be maximized).
    maximized: bool,
    /// Whether the tray app has been told the window is up.
    told: bool,
    /// How long nothing has been on screen and nothing has been loading.
    nothing_for: u32,
}

static OPENING: Lazy<Mutex<Opening>> = Lazy::new(|| Mutex::new(Opening::default()));

/// Whether this process's own window has been on the screen, which is what makes its going the process's
/// own end for a viewer whose pin was detached (see `settle`).
static THIS_WINDOW_WAS_UP: AtomicBool = AtomicBool::new(false);

/// Open the window this process was handed over: a preview of the file, kept out of sight, which the
/// tick below takes up as a pin the moment it is settled.
///
/// Called once, by `run_preview_window`, after the channel and the window classes are in place and
/// before the loop turns. A process that is not a detached window asks nothing and does nothing.
pub(super) fn begin() {
    let Some(request) = detach::open_request() else {
        return;
    };

    let Ok(mut opening) = OPENING.lock() else {
        return;
    };
    if opening.begun {
        return;
    }
    opening.begun = true;
    opening.since = Some(Instant::now());
    let point = (request.region.0, request.region.1);
    drop(opening);

    // The very message a hover of this file arrives as — the file, a point for the layout to be
    // placed around, and no item to be kept clear of, because there is no listing here to keep off
    // (see `PreviewMessage::Show`). Where it lands is not the point: the window this becomes is
    // placed by the hand-off's own box (see `settle`).
    send_preview(PreviewMessage::Show(request.path, point.0, point.1, None));
}

/// Answer the tick's questions: has the file been opened, is the window drawn, and is it still there.
///
/// Called at the end of every turn of the loop, where the preview and its paints have just been
/// looked at — `current_show` and `pending` are the loop's own record of what is on screen and what
/// is still coming, which is what the pin key's own question is asked of.
pub(super) fn settle(current_show: Option<&PreviewMessage>, pending: bool) {
    if !detach::is_detached() {
        return;
    }

    let Some(request) = detach::open_request() else {
        // **A viewer whose pin became a window of its own was handed nothing**, so it has none of the road
        // below: it has the pin it was already drawing, and that pin is the only window this process has.
        // Once that window has gone there is nothing left for the process to be — nobody is waiting on it and
        // no other window of its own is on the screen — so it goes the way a handed-over window's does (see
        // `the_process_may_go`). Said with a latch rather than with "nothing is pinned", because the tick a
        // hover is taken up on has nothing pinned either.
        if pinned() {
            THIS_WINDOW_WAS_UP.store(true, Ordering::Release);
        } else if THIS_WINDOW_WAS_UP.load(Ordering::Acquire) {
            detach::the_process_may_go();
        }
        return;
    };

    let Ok(mut opening) = OPENING.lock() else {
        return;
    };
    if !opening.begun {
        return;
    }

    if pinned() {
        // The maximize the hand-off carried: the box a maximized window fills is a box chosen for
        // the whole window, so it is asked of the pin that now exists rather than by the take-up.
        // The answer is one the pin gives itself on its next turn, which is why this turn is spent
        // asking rather than counting.
        if !opening.maximized {
            opening.maximized = true;
            if a_maximize_is_owed() {
                ask_pin(PinCommand::Maximize);
                return;
            }
        }

        opening.up_for = opening.up_for.saturating_add(1);
        if opening.up_for >= PAINTED_BEFORE_TELLING && !opening.told {
            opening.told = true;
            drop(opening);
            // The moment the tray app's own pin may come down: the file is in a window that is
            // drawn, in the box it was drawn in, at the size it was drawn at.
            detach::the_window_is_up();
        }
        return;
    }

    // A window that was up and is not, or one that has been told about and is not: either way this
    // process has done the one thing it was begun for, and a detached window has no tray icon to
    // sit under afterwards.
    if opening.told || opening.up_for > 0 {
        drop(opening);
        detach::the_process_may_go();
        return;
    }

    // The take-up has been asked for and has not landed yet: nothing to look at until it does, and
    // the preview it came from is kept out of sight until the window replaces it.
    keep_the_preview_off_the_screen();

    if opening.asked {
        if opening
            .since
            .is_some_and(|since| since.elapsed() > OPEN_WAIT)
        {
            drop(opening);
            detach::the_process_may_go();
        }
        return;
    }

    // A preview two cards of this app's wide, standing where a hand would see it, is not what a
    // detached window is: it is drawn at a place past every display until it becomes the window
    // the hand-off named.

    // Nothing on screen and nothing in flight is a file that was asked for and answered with
    // nothing — a kind this build cannot show — so the wait ends rather than standing here for it.
    if !pending && current_show.is_none() {
        opening.nothing_for = opening.nothing_for.saturating_add(1);
        if opening.nothing_for > NOTHING_IS_COMING_FOR {
            drop(opening);
            detach::the_process_may_go();
            return;
        }
    } else {
        opening.nothing_for = 0;
    }

    if opening
        .since
        .is_some_and(|since| since.elapsed() > OPEN_WAIT)
    {
        drop(opening);
        detach::the_process_may_go();
        return;
    }

    // The pin key's own question, asked by the window that replaced the key: what is on screen, is it
    // settled, and is it this file. A spinner is not settled and a page still being drawn is not
    // either, which is what the load in flight is carried in for (see `pin_what_is_on_screen`).
    let current = current_show.cloned();
    if let Some(PreviewMessage::Pin { path, .. }) = pin_what_is_on_screen(&current, None) {
        if path != request.path {
            return;
        }
        opening.asked = true;
        drop(opening);
        // The box is the hand-off's rather than the hover's: what the tray app's pin was showing the
        // file at, which is what makes this window the same window. The media itself is the hover's,
        // which is what has just been loaded, decoded and — for a film, a document, a model — begun.
        send_preview(PreviewMessage::Pin {
            path,
            rect: request.region,
            arriving: true,
        });
    }
}

/// Keep the preview nobody is to see past the corner of every display.
///
/// A position rather than a hidden window, and that is the whole of the choice: the box a preview is
/// laid out, decoded and painted at is the box the window is standing at, and a hidden window is one
/// the loop has stopped drawing altogether. Moving it costs a `SetWindowPos` on a window that is not
/// composed, and the take-up after it is given the box the hand-off named rather than this one.
fn keep_the_preview_off_the_screen() {
    let hwnd = HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut std::ffi::c_void);
    if hwnd.0.is_null() {
        return;
    }

    // Safety: a window of this process's own, moved without being resized, raised or activated.
    unsafe {
        let _ = SetWindowPos(
            hwnd,
            None,
            OFF_THE_SCREEN,
            OFF_THE_SCREEN,
            0,
            0,
            SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE,
        );
    }
}

/// Whether the hand-off asked for a maximized window and the window can be one.
///
/// A sound's card is the kind that cannot: it is drawn at the size its own setting names in the
/// middle of whatever box it is given, and it draws no maximize to press — so a pin of one has no
/// restore box and no maximize to ask for, which is exactly the answer `restore` gives here.
fn a_maximize_is_owed() -> bool {
    if !detach::open_request().is_some_and(|request| request.maximized) {
        return false;
    }

    let Some(pinned) = pin_state() else {
        return false;
    };
    let Some(pin) = pinned.pin() else {
        return false;
    };

    pin.frame != PinFrame::None && pin.restore.is_none()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_window_nobody_asked_to_maximize_owes_no_maximize() {
        // `a_maximize_is_owed` reads a request that only a detached process has, and a test process
        // is the tray app: the answer is no before anything else is looked at.
        assert!(!a_maximize_is_owed());
    }
}

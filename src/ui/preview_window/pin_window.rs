//! The window a road out of a pin acts on, and the two ends that can act on it.
//!
//! The pin's *state* had one teardown before this (`end_pin_state_guards`); its *window*
//! work had two, written out by hand at each end, and the watchdog's was the copy that
//! drifted. That copy kept the pin's pointer and its keyboard, so a pin the watchdog killed
//! for being hung left a desktop nothing could be clicked on — the same defect ADR 4 records
//! for the state list, one level out, and the same repair: one list, called from every exit.
//!
//! So the window work is one function over two roads. The road is the value that says which
//! of the two ends may act, and it is `PinDrag::delivered`'s shape: one field deciding which
//! of two ends may do the thing. The loop's own tick owns every window this app has and is
//! turning, so it may act on them. The watchdog is a thread that has given up on the loop, so
//! it may not make the loop answer it and must not wait on it: the one request only the loop
//! can make of the window is posted and left for a loop that comes back, and the hide is put
//! on a thread of the watchdog's own rather than run on the one that has given up.
//!
//! The seam is seven operations and nothing wider. `move` and `blit` are not here because
//! nothing on a road out of a pin moves or paints the preview window: a take-down hands the
//! window to the loop's ordinary take-down, and the box a pin is put up at was chosen long
//! before any of this (see `placed_pin_box`, and ADR 6, which keeps a box and the paint that
//! stands it up in one call). Widening a seam with operations nothing calls is the
//! speculative half of this decision; the half that earns its keep is that the two roads now
//! share one list, and that the list can be read back on a machine with no display and no pin
//! in it.

#[cfg(test)]
use std::sync::Mutex;

/// The message a pin's own window procedure answers to let the pointer go, which only the
/// watchdog ever asks for.
///
/// Posted and never sent, because the loop being given up on is the thing that must not be
/// waited on: a loop that was slow rather than gone does come back to a window that is
/// still holding the pointer, and a window that comes back still holding it is a desktop
/// nothing can be clicked on (see `spawn_pin_watchdog`).
pub(crate) const WM_PIN_RELEASE_POINTER: u32 = windows::Win32::UI::WindowsAndMessaging::WM_APP + 4;

/// Which of the two ends is taking a pin down, and therefore what it may ask of the window.
///
/// Take up, take down and end are the pin's three lifecycle transitions and the three are
/// not the same event; this is the second of them seen from the window's side. What each end
/// may do of the window is the value and nothing else, which is the point: the window work
/// used to be written out once per hand, and the copies drifted.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum PinExit {
    /// The preview loop's own tick, which owns the windows and is turning. The take-down it
    /// performs is the ordinary one a hover's dismissal goes through, so the windows come
    /// down as part of that and this end has only the pointer to let go of.
    Loop,
    /// The watchdog's thread, which has given up on the loop. The loop will not be answering
    /// anything, so this end asks by message and hides on a thread of its own.
    Watchdog,
}

impl PinExit {
    /// Whether this end may let the pointer go itself rather than ask the loop to.
    ///
    /// Only the loop's own tick may. A capture is held per thread, and only the thread the
    /// press was delivered on can let it go; a watchdog that called `ReleaseCapture` would
    /// release a capture it does not hold — a press belonging to another window — and hand
    /// that press to a window the hand is not aimed at.
    pub(crate) fn may_release_pointer(self) -> bool {
        self == PinExit::Loop
    }
}

/// What a road out of a pin is allowed to ask of a window.
///
/// Each method is one Win32 operation, or the pair Windows requires to happen together, and
/// none of them is a policy: *which* of them a road asks for, and in what order, is the
/// decision this seam exists to make testable. A handle is `isize` for the reason
/// `NOACTIVATE_WAKE` is one — a handle is a raw pointer and these cross threads — and zero
/// is no window, which is what `HWND(0)`, a failed `GetCapture` and `SetFocus(None)` each
/// already mean.
pub(crate) trait PinWindow {
    /// The window a pin is drawn in, or zero where there is none.
    fn hwnd(&self) -> isize;

    /// Let go of the pointer, if this window is the one holding it.
    fn release_capture(&self, hwnd: isize);

    /// Make this window one that can be focused, or one that cannot.
    fn set_focusable(&self, hwnd: isize, focusable: bool);

    /// Put the keyboard on a window, or take it off every window where `hwnd` is zero.
    fn set_focus(&self, hwnd: isize);

    /// Bring a window to the front, so a keyboard handed back lands where it came from.
    fn set_foreground(&self, hwnd: isize);

    /// Take down the pinned window and everything of somebody else's standing in it.
    ///
    /// The player's window belongs to the engine and the browser's to the browser, so this
    /// is more than one `ShowWindow` and is named for what it is rather than for the first
    /// call in it.
    fn hide_pin_windows(&self);

    /// Leave a message for the loop to answer, rather than making it answer now.
    fn post(&self, hwnd: isize, message: u32);
}

/// One thing a road out of a pin was seen to do, in the order it did it.
///
/// Recorded rather than returned so a test can assert the whole of a road's list with the
/// order inside it, which is where the roads differ and where a hand-written copy drifts
/// first: the pointer is let go *before* the state goes down, because `ReleaseCapture`
/// delivers `WM_CAPTURECHANGED` back into the window procedure this same thread is running,
/// and a window procedure that finds the pin already gone has nothing to release a drag for.
#[cfg(test)]
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum PinWindowCall {
    ReleaseCapture,
    SetFocusable {
        focusable: bool,
    },
    SetFocus,
    SetForeground,
    /// The pinned window and everything of somebody else's standing in it taken off the
    /// screen. Recorded rather than done inside the teardown, because the thread that can run
    /// it is the caller's to choose and the two callers choose differently.
    HidePinWindows,
    Post(u32),
}

/// A window that records what it was asked to do instead of doing it, so a road out of a
/// pin can be read back.
///
/// The second adapter, and the reason this is a seam rather than a signature. With a window
/// that remembers, each road is a list that can be asserted whole on a machine with no
/// display and no pin in it. It also *answers*: a recorder can be told there is no window at
/// all, so the branch that only fires where the handle is missing — the one a live desktop
/// almost never takes, and the one where a teardown that assumes a window has nothing to let
/// go of — is reachable at all.
#[cfg(test)]
pub(crate) struct RecordedPinWindow {
    hwnd: isize,
    calls: Mutex<Vec<PinWindowCall>>,
}

#[cfg(test)]
impl RecordedPinWindow {
    /// A recorder standing in for a window that is there, or for one that is not.
    pub(crate) fn new(hwnd: isize) -> Self {
        Self {
            hwnd,
            calls: Mutex::new(Vec::new()),
        }
    }

    /// Everything the road did, in the order it did it.
    pub(crate) fn calls(&self) -> Vec<PinWindowCall> {
        self.calls
            .lock()
            .map(|calls| calls.clone())
            .unwrap_or_default()
    }

    fn record(&self, call: PinWindowCall) {
        if let Ok(mut calls) = self.calls.lock() {
            calls.push(call);
        }
    }
}

#[cfg(test)]
impl PinWindow for RecordedPinWindow {
    fn hwnd(&self) -> isize {
        self.hwnd
    }

    fn release_capture(&self, _hwnd: isize) {
        self.record(PinWindowCall::ReleaseCapture);
    }

    fn set_focusable(&self, _hwnd: isize, focusable: bool) {
        self.record(PinWindowCall::SetFocusable { focusable });
    }

    fn set_focus(&self, _hwnd: isize) {
        self.record(PinWindowCall::SetFocus);
    }

    fn set_foreground(&self, _hwnd: isize) {
        self.record(PinWindowCall::SetForeground);
    }

    fn hide_pin_windows(&self) {
        self.record(PinWindowCall::HidePinWindows);
    }

    fn post(&self, _hwnd: isize, message: u32) {
        self.record(PinWindowCall::Post(message));
    }
}

/// What one road out of a pin still owes the window, and on which thread.
///
/// The hide is owed rather than done inside the teardown for one reason: the thread that can
/// run it is the caller's to choose, and the two callers choose differently for the same
/// reason they choose everything else differently. The loop's tick hides the pin on its own
/// turn, as part of the ordinary take-down; the watchdog has given up on that turn and puts
/// the hide on a thread of its own. Both go through the same list.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum PinHide {
    /// The window comes down as part of the ordinary take-down, on the loop's own tick.
    WithTheTakeDown,
    /// The loop will not be answering, so the hide runs on a thread of the watchdog's own.
    OnItsOwnThread,
}

/// Take a pin down, from one of the two ends, on the window that end may act on.
///
/// This is the window half of the teardown ADR 4 made one list, and it is one list for the
/// same reason. What the loop's tick and the watchdog's thread have in common they do here,
/// once; what they may not do differently is the road, and that is the only thing that says
/// which of them they are doing.
///
/// `settle` is the pin's own state teardown, called in the middle of this rather than around
/// it, and it is handed the same window because the keyboard handover in the middle of it is
/// window work: a pin that took the focus has to put it back where it came from, and that
/// was one of the items the watchdog's hand-written copy of this list left out.
pub(crate) fn take_pin_down(
    exit: PinExit,
    window: &dyn PinWindow,
    settle: &mut dyn FnMut(&dyn PinWindow),
) -> PinHide {
    let hwnd = window.hwnd();

    // The pointer goes before the state does, and only where this thread is the one holding
    // it: `ReleaseCapture` delivers `WM_CAPTURECHANGED` back into this thread's own window
    // procedure, and that procedure has to find a pin still standing to end a drag for.
    if exit.may_release_pointer() && hwnd != 0 {
        window.release_capture(hwnd);
    }

    settle(window);

    match exit {
        PinExit::Loop => PinHide::WithTheTakeDown,
        PinExit::Watchdog => {
            if hwnd != 0 {
                window.post(hwnd, WM_PIN_RELEASE_POINTER);
            }
            PinHide::OnItsOwnThread
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The pointer is let go before the state settles, and only where this thread is the one
    /// holding it.
    ///
    /// The order is load-bearing rather than incidental. `ReleaseCapture` delivers
    /// `WM_CAPTURECHANGED` back into the window procedure this same thread is running, and
    /// that procedure ends a drag by looking the pin up: a release taken after the state has
    /// gone finds nothing to end. A capture release moved to the other side of the settle is a
    /// drag that is never ended, and a window whose caption then answers nothing at all.
    ///
    /// And only the loop may do it at all. A capture belongs to the thread that took it, so a
    /// watchdog asking for one releases nothing — or, if some other window on that thread is
    /// holding one, releases that window's press and hands it to a pin the hand is not aimed
    /// at. That is why the watchdog's road asks by message instead, and the message is posted
    /// rather than sent because the loop it is asking is the loop it has given up on.
    #[test]
    fn the_pointer_goes_before_the_state_and_only_where_this_thread_holds_it() {
        let loop_window = RecordedPinWindow::new(0x1000);
        let mut settle = |_: &dyn PinWindow| {};
        assert_eq!(
            take_pin_down(PinExit::Loop, &loop_window, &mut settle),
            PinHide::WithTheTakeDown
        );
        assert_eq!(
            loop_window.calls(),
            vec![PinWindowCall::ReleaseCapture],
            "the release is taken while this thread can still end the drag it belongs to"
        );

        let watchdog_window = RecordedPinWindow::new(0x1000);
        assert_eq!(
            take_pin_down(PinExit::Watchdog, &watchdog_window, &mut settle),
            PinHide::OnItsOwnThread
        );
        assert_eq!(
            watchdog_window.calls(),
            vec![PinWindowCall::Post(WM_PIN_RELEASE_POINTER)],
            "a watchdog has no capture of its own to release, and cannot wait on the loop's"
        );
    }

    /// The whole of what each road owes the window, for a pin that is holding all of it.
    ///
    /// One test for the whole list rather than a sample, because the defect this seam exists
    /// for is a list that had become two: the watchdog's hand-written copy of the window work
    /// had four of the five items below missing or wrong, and every one of them was invisible
    /// until a pin was killed for being hung. A test that asserts the first item and moves on
    /// would have passed against the copy.
    #[test]
    fn each_road_owes_the_window_the_whole_of_the_list() {
        // A settle that is the pin's own teardown: everything about the state, and the keyboard
        // handover in the middle of it, which is the item the copy dropped.
        let mut settle = |window: &dyn PinWindow| {
            window.set_focus(0);
            window.set_focusable(0x1000, false);
            window.set_foreground(0x2000);
        };

        let loop_window = RecordedPinWindow::new(0x1000);
        take_pin_down(PinExit::Loop, &loop_window, &mut settle);
        assert_eq!(
            loop_window.calls(),
            vec![
                PinWindowCall::ReleaseCapture,
                PinWindowCall::SetFocus,
                PinWindowCall::SetFocusable { focusable: false },
                PinWindowCall::SetForeground,
            ],
            "the loop releases the pointer and then settles the pin, which gives the keyboard back"
        );

        let watchdog_window = RecordedPinWindow::new(0x1000);
        take_pin_down(PinExit::Watchdog, &watchdog_window, &mut settle);
        assert_eq!(
            watchdog_window.calls(),
            vec![
                PinWindowCall::SetFocus,
                PinWindowCall::SetFocusable { focusable: false },
                PinWindowCall::SetForeground,
                PinWindowCall::Post(WM_PIN_RELEASE_POINTER),
            ],
            "the watchdog settles the pin first — the keyboard handover is not optional — and asks \
             for the pointer afterwards, because by then the drag it belonged to is gone"
        );
    }

    /// A road out of a pin with no window behind it asks it for nothing.
    ///
    /// Zero is a handle this app has not been given yet, which is what the window procedure
    /// sees before the window exists and what the watchdog sees if the app is on its way down.
    /// A teardown that assumed a window would post a message to handle zero and try to release
    /// a capture of zero, both of which are calls on a handle that may since be somebody else's.
    #[test]
    fn a_road_with_no_window_asks_it_for_nothing() {
        for exit in [PinExit::Loop, PinExit::Watchdog] {
            let window = RecordedPinWindow::new(0);
            let mut settle = |window: &dyn PinWindow| window.set_focus(0);
            take_pin_down(exit, &window, &mut settle);

            assert_eq!(
                window.calls(),
                vec![PinWindowCall::SetFocus],
                "{exit:?}: the state still settles, and the window still has nothing asked of it"
            );
        }
    }
}

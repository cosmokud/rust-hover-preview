//! What a road out of a pin is allowed to ask of a window, and the window that remembers being asked.
//!
//! The seam and its one adapter, which is why neither is the pin's lifecycle: what a road may
//! ask of the window is the value (`PinExit` picks the list, `end_pin` in the parent runs it),
//! and `RecordedPinWindow` is the window that answers instead of acting, so a road can be read
//! back whole. The lifecycle and the state it acts on stay in the parent.

#[cfg(test)]
use std::sync::Mutex;

use super::ScreenRegion;

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

    /// Where the pointer is, in screen coordinates.
    ///
    /// Read rather than taken from a mouse message, and that is the whole of why: a window being
    /// resized moves its own origin out from under the coordinates a message would have carried,
    /// so the delta a message reports is measured in a frame the window has itself moved. A drag
    /// asks the pointer where it is instead, which is a question with one answer however late the
    /// message that should have carried it arrives.
    fn pointer(&self) -> Option<(i32, i32)>;

    /// The box the window is standing at on the screen, or nothing where it has none.
    ///
    /// The screen's own box and not the one the pin remembers: a window the hand has carried
    /// since it was maximized is no longer standing where the maximize left it, and a resize
    /// begun from the remembered box is begun from the screen's top border rather than from where
    /// the hand left the window — which is the window snapping back the moment an edge is pulled
    /// (see `apply_pin_drag`).
    fn window_box(&self, hwnd: isize) -> Option<ScreenRegion>;

    /// Take the pointer for this window, so every message about it arrives here.
    fn capture(&self, hwnd: isize);

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

    /// Take down the round bubble a collapsed pin leaves standing for it.
    ///
    /// A window of this app's own, and the one thing a take-down owes that the ordinary
    /// take-down does not do: the loop's own hide brings the pin's windows down as part of the
    /// ordinary message, and a bubble is not the pin's window and is not in that message.
    fn hide_pin_bubble(&self);

    /// Paint the pin's own window as it stands.
    ///
    /// Named rather than being folded into `hide_pin_windows` because it is the one operation
    /// here that has a caller on a road *into* a window as well as out of one: a drag that has
    /// just been let go of has to be drawn at where the hand left it, and the drawing is GDI on
    /// a handle that has to be a real window. Left outside the seam, a test could only exercise
    /// a drag's end by naming a handle it did not have — which is a repaint of somebody else's
    /// memory rather than a test.
    fn repaint(&self);

    /// Put the window FFmpeg draws the picture into back where the pin's media band is, and put it
    /// up there — which is the whole of what undoes a drag's park.
    ///
    /// One method rather than a show and a place, because on a window of another process the show
    /// *is* the place: `ShowWindow` un-hides a window where it stands, and both halves of undoing
    /// a park are one `SetWindowPos` carrying `SWP_SHOWWINDOW` (see `ensure_video_window_topmost`).
    /// Split into two, the second is the half that can be left out, and a show without the place
    /// is a film on screen at the box the drag began at — the failure this call exists to be able
    /// to assert about.
    ///
    /// The band is read at the moment the park ends rather than remembered from before it, which
    /// is the whole of what a move changes: the hand carries the window somewhere else and every
    /// place that would have told the player about it was answered out of the hand while the park
    /// stood, so the rect the player's window is still at is the one the drag started from.
    ///
    /// `None` for a pin with no band to go back into — collapsed, or already gone — where there is
    /// no rect to be told and only the window's own hiddenness left to take back. That is a real
    /// answer rather than a degenerate one, and the recorder can be given it.
    fn unpark_player_window(&self, band: Option<ScreenRegion>);

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
    /// Where the pointer is, and the refusal where the machine would not say.
    Pointer(Option<(i32, i32)>),
    /// Where the window stands, and the refusal where it has no box to stand at.
    WindowBox(Option<ScreenRegion>),
    /// The pointer taken for this window, which is every mouse message on the desktop arriving
    /// here from that moment on.
    Capture,
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
    /// The round bubble a collapsed pin left standing for it, taken down with the take-down.
    HidePinBubble,
    /// The pin's own window drawn as it stands, which is what a drag's end owes after letting
    /// go of the pointer: the window is where the hand left it and the screen has to be told.
    Repaint,
    /// The player's own window put back where the pin's band is and put up there — the band, or
    /// nothing at all where there is no band to put it back into.
    ///
    /// Recorded whole rather than as a show and a place, because the two are one `SetWindowPos` on
    /// a window of another process and a test that could only see one of them would pass against
    /// the defect this records: a picture on screen at the box the drag began at.
    UnparkPlayerWindow(Option<ScreenRegion>),
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
    /// Where the pointer is, for a drag measured off it.
    pointer: Option<(i32, i32)>,
    /// Where the window stands, for a drag begun from the screen's own box rather than the pin's.
    box_: Option<ScreenRegion>,
    calls: Mutex<Vec<PinWindowCall>>,
}

#[cfg(test)]
impl RecordedPinWindow {
    /// A recorder standing in for a window that is there, or for one that is not.
    pub(crate) fn new(hwnd: isize) -> Self {
        Self::with(hwnd, None, None)
    }

    /// A recorder for a window the pointer is on and that stands at a box of its own.
    ///
    /// Both halves are given rather than invented because they are the two things a drag is
    /// measured against, and a recorder that made them up would be testing a drag against a
    /// pointer this test invented rather than the one it says it is standing at. `None` for
    /// either is a real answer a real desktop gives — `GetCursorPos` and `GetWindowRect` are both
    /// refusable — and it is the answer that leaves a window holding the pointer with nothing
    /// that will ever let go of it.
    pub(crate) fn with(
        hwnd: isize,
        pointer: Option<(i32, i32)>,
        box_: Option<ScreenRegion>,
    ) -> Self {
        Self {
            hwnd,
            pointer,
            box_,
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

    fn pointer(&self) -> Option<(i32, i32)> {
        let pointer = self.pointer;
        self.record(PinWindowCall::Pointer(pointer));
        pointer
    }

    fn window_box(&self, _hwnd: isize) -> Option<ScreenRegion> {
        self.record(PinWindowCall::WindowBox(self.box_));
        self.box_
    }

    fn capture(&self, _hwnd: isize) {
        self.record(PinWindowCall::Capture);
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

    fn hide_pin_bubble(&self) {
        self.record(PinWindowCall::HidePinBubble);
    }

    fn repaint(&self) {
        self.record(PinWindowCall::Repaint);
    }

    fn unpark_player_window(&self, band: Option<ScreenRegion>) {
        self.record(PinWindowCall::UnparkPlayerWindow(band));
    }

    fn post(&self, _hwnd: isize, message: u32) {
        self.record(PinWindowCall::Post(message));
    }
}

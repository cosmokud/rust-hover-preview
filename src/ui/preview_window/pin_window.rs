//! The pin's lifecycle in one value, and the window a road out of it acts on.
//!
//! The pin's *state* had one teardown before this (`end_pin_state_guards`); its *window* work
//! had two, written out by hand at each end, and the watchdog's was the copy that drifted. That
//! copy kept the pin's pointer and its keyboard, so a pin the watchdog killed for being hung
//! left a desktop nothing could be clicked on — the same defect ADR 4 records for the state
//! list, one level out, and the same repair: one list, called from every exit.
//!
//! So this module is the pin's lifecycle and the window a road out of it acts on, and the
//! second is what the first has to be asked about. `PinState` is the whole of what a pin is:
//! the window and what it holds, in one value under one lock. The road is the value that says
//! which of the two ends may act, and it is `PinExit`'s shape: one field deciding which of two
//! ends may do the thing. The loop's own tick owns every window this app has and is turning,
//! so it may act on them. The watchdog is a thread that has given up on the loop, so it may not
//! make the loop answer it and must not wait on it: the one request only the loop can make of
//! the window is posted and left for a loop that comes back, and the hide is put on a thread of
//! the watchdog's own rather than run on the one that has given up.
//!
//! What the type buys is that the illegal orders have no spelling. A pin the watchdog killed
//! used to be able to keep `PIN_FOCUSED` and the foreground it had stolen, because those were
//! two globals with nothing saying they belonged to a pin that had gone; here they are a field
//! of a pin that exists, so `Ending` cannot be holding a keyboard. A walk queued before a kill
//! used to be answerable after it, for the same reason; here the queue is a field of the pin
//! too, and `end` empties it in the same move that takes the window away.
//!
//! Four things are deliberately *not* in here, and each was written down rather than left to be
//! found. The pin's own geometry and press handling stay where they were — this is the
//! lifecycle, not the drawing. The walk the planner is working on is the planner's, and the
//! planner holds its lock across a `Condvar::wait_timeout`, so folding it puts the pin's lock
//! on a wait path. The loop's liveness clock stays an atomic, because the watchdog reads it
//! from a thread that has given up on the loop and must not be able to block on the loop's lock
//! to learn the loop is gone (see `note_pin_alive`). And what the Explorer hook asks about a
//! pin — whether one is up, and whether one has been closed since it last asked — is published
//! rather than held, for the same reason: the hook asks on every tick, and a hook blocked
//! behind a preview thread that has stopped turning is a hook that stops answering Explorer,
//! which is a worse failure than the one the publication exists to avoid.
//!
//! The seam is seven operations and nothing wider. `move` and `blit` are not here because
//! nothing on a road out of a pin moves or paints the preview window: a take-down hands the
//! window to the loop's ordinary take-down, and the box a pin is put up at was chosen long
//! before any of this (see `placed_pin_box`, and ADR 6, which keeps a box and the paint that
//! stands it up in one call). Widening a seam with operations nothing calls is the speculative
//! half of this decision; the half that earns its keep is that the two roads now share one list,
//! and that the list can be read back on a machine with no display and no pin in it.

use std::collections::VecDeque;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Mutex, MutexGuard};

use super::{PinCommand, PinnedPreview};

/// The message a pin's own window procedure answers to let the pointer go, which only the
/// watchdog ever asks for.
///
/// Posted and never sent, because the loop being given up on is the thing that must not be
/// waited on: a loop that was slow rather than gone does come back to a window that is
/// still holding the pointer, and a window that comes back still holding it is a desktop
/// nothing can be clicked on (see `spawn_pin_watchdog`).
pub(super) const WM_PIN_RELEASE_POINTER: u32 = windows::Win32::UI::WindowsAndMessaging::WM_APP + 4;

/// Which of the two ends is taking a pin down, and therefore what it may ask of the window.
///
/// Take up, take down and end are the pin's three lifecycle transitions and the three are
/// not the same event; this is the second of them seen from the window's side. What each end
/// may do of the window is the value and nothing else, which is the point: the window work
/// used to be written out once per hand, and the copies drifted.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum PinExit {
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
    pub(super) fn may_release_pointer(self) -> bool {
        self == PinExit::Loop
    }
}

/// Everything a pin is holding while it is up.
///
/// The three fields are three things that used to be globals with nothing saying they belonged
/// together, and each of them is a way a pin has outlived what it was holding: a keyboard
/// claim that outlived the window (the watchdog's stranded caret), a command queue answered
/// into whatever pin came next (the queue that stopped being a slot), and a collapsed flag read
/// as `pin_is_collapsed` while the pin itself had gone.
pub(super) struct PinUp {
    /// The window and what it is showing: the geometry, the chrome, the pointer's business
    /// and the player's, all of which is the pin's own and outlives every flag beside it.
    pin: PinnedPreview,
    /// The keyboard this app took for this pin, or nothing while it holds none.
    ///
    /// One field and not a bit and a handle, because those are what it was: `PIN_FOCUSED` said
    /// the keyboard had been taken and `PIN_PREVIOUS_FOREGROUND` said where it came from, and
    /// either could be set without the other. Ending a pin needs both or neither, so a pin that
    /// took the keyboard from nowhere has no spelling here.
    keyboard: Option<PinKeyboard>,
    /// The commands the chrome has left for the loop, oldest first.
    ///
    /// A queue rather than a slot, and the reason is a double-click: a slot holds one command
    /// and the next write over it, so a double-click on `Next` — two presses, the second within
    /// the system double-click time, both delivered as separate `WM_LBUTTONUP`s — left one
    /// command where the user asked for two, and the second file was not walked to. Two `Next`
    /// clicks are the ordinary way to move two files along, so the loss was on the commonest
    /// button.
    ///
    /// Bounded because the loop that drains it is the one that would have to be stopped for it to
    /// grow: a caption clicked faster than the loop turns, which is a hand drumming on a button.
    /// At that rate the loop is behind anyway, and what is dropped is the oldest, so what
    /// survives is the user's latest intent rather than the first thing they asked for.
    ///
    /// It is a field of the pin rather than a global because of what a pin's end has to do with
    /// it: a command left behind is a command about a window that is no longer there, fired
    /// against whatever pin comes next — the chrome is drawn for the file now on screen, so a
    /// `Close` that belonged to the last one closes this one. `end` empties it in the same move
    /// that takes the window, so there is no order in which one happens without the other.
    commands: VecDeque<PinCommand>,
}

/// How many commands may be waiting before the oldest is dropped. A loop turns on the order of
/// sixty times a second and a caption button is a press, so a queue this deep is already a
/// loop that is not keeping up rather than a hand that is ahead of it.
const PIN_COMMANDS_MAX: usize = 16;

/// The keyboard a pin has taken, and where it has to be put back to.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
struct PinKeyboard {
    /// The window that was in front when the keyboard was taken, or zero where there was none
    /// to put it back to.
    ///
    /// Zero is `HWND(0)`'s own meaning rather than a sentinel this module invented: a
    /// `WS_POPUP` with no parent has no `GW_OWNER` at all, so there is sometimes no window to
    /// hand back to and the handover is the focus going to nothing.
    behind: isize,
}

/// Which of the three places a pin is in.
///
/// `Down` and `Ending` are both "no pin is up", and they are separate because a pin that is
/// *over* still has to be answerable for having been, which is the same thing said after a
/// road out of it as before one. The pin's own value is only in `Up`, so there is no way to
/// write a state that is over and still holding a pin.
///
/// The variants are a window and two words, and that is deliberate rather than boxed: the value
/// lives behind one lock and a lock hands out a pointer, so the size of the largest variant is
/// paid once when the lock is taken rather than on every read of it, and boxing `PinUp` would
/// add an allocation and a pointer chase to every mouse message the window is sent.
#[allow(clippy::large_enum_variant)]
pub(super) enum PinState {
    /// No pin is up, and no pin has been taken down since the last one was taken up.
    Down,
    /// A pin is up, and this is everything it is holding.
    Up(PinUp),
    /// A pin is over. Nothing is held: the state that could hold something has been taken by
    /// the move that got here.
    Ending,
}

/// The pin's lifecycle, in one value and behind one lock.
///
/// Written in the past tense and naming what it exists to stop, as everything else in this
/// tree is: the loop took a pin up by writing four globals and a window's extended style, and
/// took it down again by clearing four globals, from three hand-written lists that had drifted
/// apart. A watchdog that killed a pin left the keyboard it had claimed and the foreground it
/// had stolen, because nothing said those belonged to the pin rather than to the app.
static PIN_STATE: Mutex<PinState> = Mutex::new(PinState::Down);

/// The pin's state, or nothing where the lock cannot be taken at all.
///
/// A poisoned lock is read as no pin rather than read through, which is what every reader of
/// `PINNED` did: a pin whose state cannot be read is a pin nothing can be asked about, and
/// "there is no pin" is the answer that has always been given. The teardown is the one thing
/// that reads through it, and says why (see `end_pin`).
pub(super) fn pin_state() -> Option<MutexGuard<'static, PinState>> {
    PIN_STATE.lock().ok()
}

/// Whether a pin is up: the published copy, and the answer every thread may take without a lock.
///
/// The Explorer hook asks this on every tick, and it cannot be answered off the lock: a hook
/// that blocks behind a preview thread that has stopped turning is a hook that stops answering
/// Explorer, which is a worse failure than the one this answers. So the phase is *published* as
/// well as held, and this is the published copy — written after the state under the lock at
/// both ends, so a reader that finds a pin up finds the pin rather than the tail of one being
/// taken down (see `install` and `end_pin`).
static PIN_UP: AtomicBool = AtomicBool::new(false);

/// Whether a pin has been closed since this was last asked, as a published copy for the same
/// reason `PIN_UP` is one.
///
/// The Explorer hook reads it once per tick to know that what is under the pointer is a hover it
/// has not answered yet: the file it was on when the pin went up is not a file the pointer has
/// left and come back to, and treating it as one would hold the next preview back for the
/// re-hover delay. What a pin is over writes it (see `end_pin`).
static PIN_RESUMED: AtomicBool = AtomicBool::new(false);

/// Whether a preview is pinned, answered without a lock and without waiting for one (see
/// [`PIN_UP`]).
pub(super) fn pin_is_up() -> bool {
    PIN_UP.load(Ordering::Acquire)
}

/// Whether a pin has been closed since this was last asked (see [`PIN_RESUMED`]).
pub(super) fn take_pin_resumed() -> bool {
    PIN_RESUMED.swap(false, Ordering::AcqRel)
}

/// Take a pin up over whatever was on screen: the window, what it is showing, and everything
/// it starts holding.
///
/// The publish happens after the state is written and before the caller is told, which is the
/// whole of what a reader of [`pin_is_up`] is owed. A pin taken up over another one — the file
/// it was showing was picked while it was up — is the same window showing another file, so what
/// belongs to the window rather than to the file is carried over by the caller before this is
/// called. A pin's keyboard claim is not carried over: it was the old file's, and the new pin
/// has pressed nothing.
pub(super) fn install(pin: PinnedPreview) {
    if let Some(mut state) = pin_state() {
        *state = PinState::Up(PinUp {
            pin,
            keyboard: None,
            commands: VecDeque::new(),
        });
        PIN_UP.store(true, Ordering::Release);
    }
}

/// What a reader of the pin is owed while it is up: a window and its own state, and nothing
/// else. Everything that wanted a whole pin takes this instead.
impl PinState {
    /// The pin's own state, or nothing where there is no pin.
    pub(super) fn pin(&self) -> Option<&PinnedPreview> {
        match self {
            PinState::Up(up) => Some(&up.pin),
            _ => None,
        }
    }

    /// The pin's own state, mutably, or nothing where there is no pin.
    ///
    /// Mutable and not `&`: everything a repaint, a drag or a tick changes about a pin is
    /// changed through here, and a reader that wanted to hold a borrow across a window call
    /// would be holding the loop's own lock across a message pump.
    pub(super) fn pin_mut(&mut self) -> Option<&mut PinnedPreview> {
        match self {
            PinState::Up(up) => Some(&mut up.pin),
            _ => None,
        }
    }

    /// Whether a pin is up at all, which is the phase and nothing about what it holds.
    #[cfg(test)]
    fn is_up(&self) -> bool {
        self.pin().is_some()
    }

    /// The claim on the keyboard itself, for the one caller that takes it away while the pin
    /// stays up.
    fn keyboard_mut(&mut self) -> Option<&mut Option<PinKeyboard>> {
        match self {
            PinState::Up(up) => Some(&mut up.keyboard),
            _ => None,
        }
    }

    /// Take the command the chrome left, if one was left.
    fn take_command(&mut self) -> Option<PinCommand> {
        match self {
            PinState::Up(up) => up.commands.pop_front(),
            _ => None,
        }
    }

    /// Take this pin down, handing back what the window work still needs.
    ///
    /// The keyboard claim is *returned* rather than read afterwards, which is what makes a pin
    /// that is over unable to still be holding one: the claim leaves with the pin, and
    /// `Ending` has nowhere to put it. The command queue goes the same way — it is a field
    /// of the pin, so a walk queued before a kill cannot be answered after it, because there
    /// is nothing left to answer it into.
    fn end(&mut self) -> Option<PinKeyboard> {
        let PinState::Up(mut up) = std::mem::replace(self, PinState::Ending) else {
            // A pin that is not up is already over, and the end it is being given is recorded
            // rather than refused: two roads racing on the same pin is normal (the watchdog and
            // the loop's own tick both watch for a hung one), and the second one must not be the
            // one that leaves the keyboard claimed.
            *self = PinState::Ending;
            return None;
        };

        up.commands.clear();
        up.keyboard
    }
}

/// Give the pin the keyboard: the window that is in front, and the pin's claim to it.
///
/// The claim and the window it came from are written together, which is what the two globals
/// this replaced did not do: `PIN_PREVIOUS_FOREGROUND` was written before the `SetFocus` and
/// `PIN_FOCUSED` after it, so a window procedure re-entered by either call — and both of them
/// deliver messages — could see one without the other.
pub(super) fn take_keyboard(behind: isize) {
    if let Some(mut state) = pin_state() {
        if let PinState::Up(up) = &mut *state {
            up.keyboard = Some(PinKeyboard { behind });
        }
    }
}

/// The pin has lost the keyboard to the window in front: drop the claim and forget where it
/// came from, because there is nothing to hand back to.
pub(super) fn release_keyboard() {
    if let Some(mut state) = pin_state() {
        if let PinState::Up(up) = &mut *state {
            up.keyboard = None;
        }
    }
}

/// Leave a command for the preview loop to act on (see [`PinUp::commands`]).
pub(super) fn ask_pin(command: PinCommand) {
    if let Some(mut state) = pin_state() {
        if let PinState::Up(up) = &mut *state {
            if up.commands.len() >= PIN_COMMANDS_MAX {
                up.commands.pop_front();
            }
            up.commands.push_back(command);
        }
    }
}

/// Take the command the chrome left, if one was left.
///
/// `None` says the queue was already empty, which is the whole of what the loop asks for
/// beyond the command itself: the read is destructive, so a command nobody wanted any more is
/// gone rather than acted on by a later tick.
pub(super) fn take_pin_command() -> Option<PinCommand> {
    pin_state()?.take_command()
}

/// The whole of what a road out of a pin owes the window, and on which thread.
///
/// The hide is owed rather than done inside the teardown for one reason: the thread that can
/// run it is the caller's to choose, and the two callers choose differently for the same
/// reason they choose everything else differently. The loop's tick hides the pin on its own
/// turn, as part of the ordinary take-down; the watchdog has given up on that turn and puts
/// the hide on a thread of its own. Both go through the same list.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum PinHide {
    /// The window comes down as part of the ordinary take-down, on the loop's own tick.
    WithTheTakeDown,
    /// The loop will not be answering, so the hide runs on a thread of the watchdog's own.
    OnItsOwnThread,
}

/// Take a pin down, from one of the two ends, on the window that end may act on.
///
/// This is the window half of the teardown, and it is one list for the same reason the state
/// half is: what the loop's tick and the watchdog's thread have in common they do here once,
/// and what they may not do differently is the road.
///
/// `settle` is the pin's own state teardown, called in the middle of this rather than around
/// it, and it is handed the same window because the keyboard handover in the middle of it is
/// window work: a pin that took the focus has to put it back where it came from.
pub(super) fn take_pin_down(
    exit: PinExit,
    window: &dyn PinWindow,
    settle: &mut dyn FnMut(&dyn PinWindow),
) -> PinHide {
    let hwnd = window.hwnd();

    // The pointer goes before the state does, and only where this thread is the one holding
    // it: `ReleaseCapture` delivers `WM_CAPTURECHANGED` back into this thread's own window
    // procedure, and that procedure ends a drag by looking the pin up — a release taken after
    // the state has gone finds nothing to end.
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

/// Put a pin down for good, whatever is holding it, and publish that none is up.
///
/// Everything the pin was holding goes with it in the one move: the keyboard claim and the
/// command queue are fields of the pin, so `Ending` is a state in which neither can be read
/// (see `PinState::end`).
pub(super) fn put_the_pin_down() {
    if let Some(mut state) = pin_state() {
        state.end();
    }
    PIN_UP.store(false, Ordering::Release);
}

/// Note that a pin is over, for the hook's own answer about what is under the pointer.
///
/// Published rather than held, like `PIN_UP`: the Explorer hook reads it once per tick and
/// cannot be answered off the pin's lock.
pub(super) fn note_pin_ended() {
    PIN_RESUMED.store(true, Ordering::Release);
}

/// Put the keyboard back where it came from, while the pin itself stays up.
///
/// This is a collapse, not an end: the pin is still a pin and is still holding its window, and
/// what it has stopped holding is the keyboard — a round bubble standing in for a window cannot
/// be a window the user types into. It is the same handover `end_pin` performs, taken out under
/// the lock and done outside it, for the same reason; what differs is that the pin is left up and
/// holding no keyboard, which is the state it is in for the rest of its life until the hand
/// presses it again.
pub(super) fn give_the_keyboard_back(window: &dyn PinWindow) {
    let keyboard = pin_state().and_then(|mut state| match state.keyboard_mut() {
        Some(keyboard) => keyboard.take(),
        None => None,
    });

    if let Some(keyboard) = keyboard {
        hand_keyboard_back(&keyboard, window.hwnd(), window);
    }
}

/// Put the keyboard back where it came from, on the window this app's own.
///
/// The focus is given up before the style is changed, because a window that has been made
/// non-activating while it still holds the focus holds it in a state Windows does not expect.
/// The window behind is one that may since have gone, and the call is refused by Windows in
/// that case rather than acted on.
fn hand_keyboard_back(keyboard: &PinKeyboard, hwnd: isize, window: &dyn PinWindow) {
    if hwnd == 0 {
        return;
    }

    window.set_focus(0);
    window.set_focusable(hwnd, false);

    if keyboard.behind != 0 {
        window.set_foreground(keyboard.behind);
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
pub(super) trait PinWindow {
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
pub(super) enum PinWindowCall {
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
pub(super) struct RecordedPinWindow {
    hwnd: isize,
    calls: Mutex<Vec<PinWindowCall>>,
}

#[cfg(test)]
impl RecordedPinWindow {
    /// A recorder standing in for a window that is there, or for one that is not.
    pub(super) fn new(hwnd: isize) -> Self {
        Self {
            hwnd,
            calls: Mutex::new(Vec::new()),
        }
    }

    /// Everything the road did, in the order it did it.
    pub(super) fn calls(&self) -> Vec<PinWindowCall> {
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

/// Whether the pin holds a claim on the keyboard, for the tests on this side of the module,
/// which cannot ask the desktop where the caret is.
#[cfg(test)]
pub(super) fn pin_holds_a_keyboard() -> bool {
    pin_state().is_some_and(|mut state| state.keyboard_mut().is_some_and(|held| held.is_some()))
}

/// Take the pin that is up away and hand it back, for a test that means to put one where it was.
///
/// The other half of [`stand_pin`], and shaped like it: a test that reached into the state to
/// stash a pin would be asserting on the type rather than on a pin, and a test that could only
/// put one back could not put *none* back — which is the state most of them need afterwards,
/// because a test that leaves a pin up is a test every test after it is standing inside.
#[cfg(test)]
pub(super) fn take_pin_for_a_test() -> Option<PinnedPreview> {
    pin_state().and_then(|mut state| {
        let PinState::Up(up) = std::mem::replace(&mut *state, PinState::Down) else {
            return None;
        };
        PIN_UP.store(false, Ordering::Release);
        Some(up.pin)
    })
}

/// Put a pin up, or take the last one down, for a test that stands one in place of whatever
/// was there.
///
/// Deliberately whole rather than a setter for each part: a test that reaches past this into
/// the state is a test asserting on the type rather than on a pin, and the transitions are
/// what this module is for.
#[cfg(test)]
pub(super) fn stand_pin(pin: Option<PinnedPreview>) {
    match pin {
        Some(pin) => install(pin),
        None => {
            if let Some(mut state) = pin_state() {
                *state = PinState::Down;
                PIN_UP.store(false, Ordering::Release);
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The lifecycle's whole vocabulary, in the order a pin goes through it, so a test that
    /// wants a pin up has one call for it and a test that wants it down has one for that.
    ///
    /// The state is a process-wide value, so these tests are one at a time: a suite of
    /// transitions over a shared machine is a suite where one test's pin is another's.
    static ONE_AT_A_TIME: Mutex<()> = Mutex::new(());

    /// The pin's own value, with nothing asked of it and nothing answered about it.
    fn a_pin() -> PinnedPreview {
        PinnedPreview::for_test()
    }

    /// A pin, up, with the keyboard taken from `behind` and one command waiting.
    fn a_pin_holding_the_keyboard(behind: isize) {
        install(a_pin());
        take_keyboard(behind);
        ask_pin(PinCommand::Next);
    }

    /// The keyboard the pin is holding, and where it came from, or nothing while it holds none.
    fn keyboard() -> Option<PinKeyboard> {
        match pin_state()?.keyboard_mut() {
            Some(keyboard) => *keyboard,
            None => None,
        }
    }

    /// The window a teardown settled, standing in for this app's own.
    fn settle_pin(behind: isize) -> RecordedPinWindow {
        let window = RecordedPinWindow::new(0x1000);
        let keyboard = pin_state().and_then(|mut state| match state.keyboard_mut() {
            Some(keyboard) => keyboard.take(),
            None => None,
        });
        if let Some(keyboard) = keyboard {
            if window.hwnd() != 0 {
                window.set_focus(0);
                window.set_focusable(window.hwnd(), false);
                if keyboard.behind != 0 {
                    window.set_foreground(keyboard.behind);
                }
            }
        }
        PIN_UP.store(false, Ordering::Release);
        PIN_RESUMED.store(true, Ordering::Release);
        stand_pin(None);
        let _ = behind;
        window
    }

    /// The pointer is let go before the state settles, and only where this thread is the one
    /// holding it.
    ///
    /// The order is load-bearing rather than incidental. `ReleaseCapture` delivers
    /// `WM_CAPTURECHANGED` back into the window procedure this same thread is running, and
    /// that procedure ends a drag by looking the pin up: a release taken after the state has
    /// gone finds nothing to end, and a window whose caption then answers nothing at all.
    ///
    /// And only the loop's own tick may do it at all. A capture belongs to the thread that
    /// took it, so a watchdog asking for one releases nothing — or, if some other window on
    /// that thread is holding one, releases that window's press and hands it to a pin the
    /// hand is not aimed at.
    #[test]
    fn the_pointer_goes_before_the_state_and_only_where_this_thread_holds_it() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
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

        install(a_pin());
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

    /// A pin that has just come up answers that it is up, and a pin that is over answers that
    /// it is not — and the two copies of that answer never disagree.
    ///
    /// The Explorer hook asks on every tick and cannot be answered off the lock, so the phase
    /// is published as well as held, and the two are written in one order at both ends.
    #[test]
    fn the_published_flag_and_the_state_never_disagree() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        assert!(pin_is_up(), "installed: a pin is published as up");
        assert!(
            pin_state().is_some_and(|state| state.is_up()),
            "and the state agrees"
        );

        let window = RecordedPinWindow::new(0x1000);
        let mut settle = |_: &dyn PinWindow| {
            settle_pin(0);
        };
        take_pin_down(PinExit::Loop, &window, &mut settle);
        assert!(!pin_is_up(), "ended: and as down");
        assert!(
            !pin_state().is_some_and(|state| state.is_up()),
            "and the state agrees"
        );
    }

    /// A pin taken up over another one is the new one whole.
    ///
    /// The pin it was is the same window showing another file, so what belongs to the window
    /// rather than to the file is carried over by the caller before the take-up. What does not
    /// carry is the claim on the keyboard: that was the old file's, and the new pin has pressed
    /// nothing.
    #[test]
    fn a_pin_taken_up_over_another_one_is_the_new_one_whole() {
        let _one = ONE_AT_A_TIME.lock();

        a_pin_holding_the_keyboard(0x2000);
        install(a_pin());

        assert!(pin_is_up(), "the swap is a pin up like any other");
        assert_eq!(
            keyboard(),
            None,
            "a pin taken up over another one has taken no keyboard, and remembers no window to \
             hand one back to"
        );
        assert_eq!(
            take_pin_command(),
            None,
            "and the queue is the new pin's, not the old one's"
        );
    }

    /// Every command a caption's button asks for survives the trip out of the window procedure
    /// and back, and a pin's end leaves none of them behind.
    #[test]
    fn every_command_a_caption_asks_for_comes_back_to_the_loop() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        for command in [
            PinCommand::Previous,
            PinCommand::Next,
            PinCommand::Minimize,
            PinCommand::Maximize,
            PinCommand::Close,
            PinCommand::Restore,
            PinCommand::TogglePlayback,
        ] {
            ask_pin(command);
            assert_eq!(
                take_pin_command(),
                Some(command),
                "{command:?} does not survive being left for the loop"
            );
        }

        // A command left in the queue when a pin ends is a command about a window that is gone,
        // fired against whatever pin comes next.
        install(a_pin());
        ask_pin(PinCommand::Close);
        stand_pin(None);
        assert_eq!(
            take_pin_command(),
            None,
            "a pin that is over leaves none behind"
        );
    }

    /// A command queue drops the oldest rather than growing without end.
    ///
    /// A loop that cannot keep up with a hand drumming on a button must not be the reason a
    /// session ends. What goes is the oldest, so what survives is what was last asked for.
    #[test]
    fn a_command_queue_drops_the_oldest_rather_than_growing_without_end() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        for _ in 0..(PIN_COMMANDS_MAX + 4) {
            ask_pin(PinCommand::Next);
        }

        let mut drained = Vec::new();
        while let Some(command) = take_pin_command() {
            drained.push(command);
        }

        assert_eq!(
            drained.len(),
            PIN_COMMANDS_MAX,
            "the queue is bounded, however many asks are made of it"
        );
    }

    /// No road out of a pin leaves the keyboard claimed.
    ///
    /// This is the defect the type exists for. `PIN_FOCUSED` and `PIN_PREVIOUS_FOREGROUND` were
    /// two globals with nothing saying they belonged to a pin that had gone, so a pin the
    /// watchdog killed kept the keyboard it had claimed and the foreground it had stolen: a
    /// desktop with no caret in it.
    #[test]
    fn no_road_out_of_a_pin_leaves_the_keyboard_claimed() {
        let _one = ONE_AT_A_TIME.lock();

        for exit in [PinExit::Loop, PinExit::Watchdog] {
            a_pin_holding_the_keyboard(0x2000);
            let window = RecordedPinWindow::new(0x1000);
            let mut settle = |_: &dyn PinWindow| {
                settle_pin(0x2000);
            };
            take_pin_down(exit, &window, &mut settle);

            assert_eq!(
                keyboard(),
                None,
                "{exit:?}: a pin that is over holds no keyboard, so there is nothing for a \
                 later pin to inherit and nothing for a reader to find"
            );
        }
    }

    /// The note that the keyboard was taken is dropped with the focus.
    ///
    /// Windows taking the focus away is the user clicking into something else: the window now
    /// in front holds the keyboard, so there is nothing to hand over and the claim goes.
    #[test]
    fn the_note_that_the_keyboard_was_taken_is_dropped_with_the_focus() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        take_keyboard(0x2000);
        release_keyboard();
        assert_eq!(keyboard(), None, "losing the focus drops the claim with it");
    }

    /// A pin with nothing on the keyboard owes nobody a handover.
    ///
    /// A pin nobody pressed holds no keyboard to give back, so a teardown of one does not go
    /// near the focus at all.
    #[test]
    fn a_pin_with_nothing_on_the_keyboard_owes_nobody_a_handover() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        let window = settle_pin(0x2000);

        assert!(
            window.calls().is_empty(),
            "nothing is asked of a window this pin never took anything from"
        );
    }

    /// The pin's own state is readable while it is up and gone once it is over.
    #[test]
    fn the_pin_s_own_state_is_readable_while_it_is_up_and_gone_once_it_is_over() {
        let _one = ONE_AT_A_TIME.lock();
        stand_pin(None);

        assert!(
            pin_state().is_none_or(|state| state.pin().is_none()),
            "no pin to read at first"
        );

        install(a_pin());
        assert!(
            pin_state().is_some_and(|state| state.pin().is_some()),
            "a pin that is up has a window and a state of its own to read"
        );

        let window = RecordedPinWindow::new(0x1000);
        let mut settle = |_: &dyn PinWindow| {
            settle_pin(0);
        };
        take_pin_down(PinExit::Loop, &window, &mut settle);
        assert!(
            pin_state().is_some_and(|state| state.pin().is_none()),
            "and a pin that is over has neither"
        );
    }

    /// A collapsed pin is still a pin, and a collapse still gives the keyboard back.
    ///
    /// The bubble is a window of this app's own and does not hold the keyboard, so a pin
    /// collapsed with the caret on it still has to hand it back.
    #[test]
    fn a_collapsed_pin_gives_the_keyboard_back_without_going_down() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        take_keyboard(0x2000);
        if let Some(mut state) = pin_state() {
            if let Some(pin) = state.pin_mut() {
                pin.collapsed = true;
            }
        }

        let window = RecordedPinWindow::new(0x1000);
        give_the_keyboard_back(&window);

        assert!(pin_is_up(), "a collapsed pin is still a pin");
        assert_eq!(keyboard(), None, "and it holds no keyboard");
        assert!(
            window.calls().contains(&PinWindowCall::SetForeground),
            "the window behind is put back in front, so a keyboard handed back lands where it \
             came from"
        );
    }

    /// A pin that took the keyboard from nothing hands it back to nothing.
    ///
    /// There is sometimes no window behind: a `WS_POPUP` with no parent has no `GW_OWNER` at
    /// all, so the handover is the focus going to nothing.
    #[test]
    fn a_pin_that_took_the_keyboard_from_nothing_hands_it_back_to_nothing() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        take_keyboard(0);
        let window = RecordedPinWindow::new(0x1000);
        give_the_keyboard_back(&window);

        assert!(
            window
                .calls()
                .iter()
                .all(|call| !matches!(call, PinWindowCall::SetForeground)),
            "there is no window behind a pin that took the keyboard from nothing"
        );
    }

    /// What a pin's end publishes for the Explorer hook, and what it does not.
    ///
    /// A pin that is over is a pointer that is on something new: the file the pin was of is
    /// not a hover the hook has already answered, and one is due the moment the pin is gone
    /// rather than after the delay a re-hover of the same file is given.
    #[test]
    fn a_pin_s_end_is_published_to_the_hook_once() {
        let _one = ONE_AT_A_TIME.lock();

        take_pin_resumed();
        install(a_pin());
        note_pin_ended();

        assert!(
            take_pin_resumed(),
            "a pin that is over is a hover the hook has not answered yet"
        );
        assert!(
            !take_pin_resumed(),
            "and it is said once, because a second true would be a file the pointer never left"
        );
    }
}

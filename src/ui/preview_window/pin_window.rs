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
//! of a pin that exists, so `Ending(_)` cannot be holding a keyboard. A walk queued before a
//! kill used to be answerable after it, for the same reason; here the queue is a field of the
//! pin too, and `end` empties it in the same move that takes the window away.
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
//! The seam is what the pin's own window is asked, and nothing wider. `move` and `blit` are not
//! here because nothing on a road out of a pin moves or paints the preview window: a take-down
//! hands the window to the loop's ordinary take-down, and the box a pin is put up at was chosen
//! long before any of this (see `placed_pin_box`, and ADR 6, which keeps a box and the paint
//! that stands it up in one call). Widening a seam with operations nothing calls is the
//! speculative half of this decision; the half that earns its keep is that the two roads now
//! share one list, and that the list can be read back on a machine with no display and no pin
//! in it.
//!
//! It grew by five for a reason worth recording, because the five are not lifecycle operations
//! at all: they are *a drag*, which is the other thing this window is asked of a pin and the
//! only one that had no test surface. `pointer`, `window_box`, `capture`, `repaint` and
//! `unpark_player_window` are the whole of what "a press becomes a carried drag, and the drag is
//! let go of" costs the machine. They are here rather than in a second trait over the same window
//! for the reason the two ends of the capture are here: they are two halves of one handoff, and a
//! handoff whose halves are asked of two different objects is a handoff that can disagree with
//! itself — which is what a window left holding the pointer for the whole desktop was (see
//! `begin_pin_drag`).
//!
//! The last of the five is a window of *another* process, and it earns its place the same way the
//! others do: the defect it covers is a film on screen at a box nobody put it at, and a test can
//! only see that by being shown the call. Every other way of undoing a drag's park was invisible
//! from here — the flag came down and nothing was recorded, so an unpark that showed the window
//! without placing it read exactly like one that placed it (see `unpark_pinned_player`).
//!
//! What is deliberately still *not* here is the rest of a drag: `apply_pin_drag` moves the
//! window and resizes it, and that is a `SetWindowPos` and an `UpdateLayeredWindow` on
//! arguments a whole frame's worth of pixels. Putting it behind this seam would mean a seam
//! wide enough to carry a paint, and the decision worth testing there — which box a drag at
//! these deltas produces — is arithmetic on four numbers and is tested as arithmetic (see
//! `resize_pinned_window`).

use std::collections::VecDeque;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Mutex, MutexGuard};

use super::{PinCommand, PinnedPreview, ScreenRegion};

mod window;

pub(super) use window::PinWindow;
#[cfg(test)]
pub(super) use window::{PinWindowCall, RecordedPinWindow};

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
    fn may_release_pointer(self) -> bool {
        self == PinExit::Loop
    }
}

/// Why a pin is being taken down, and therefore which of the two ends is taking it down.
///
/// This is what every caller of `end_pin` now says instead of choosing a road for itself. The
/// road was a free choice before, which is how the watchdog's copy of the teardown came to
/// exist: three callers, three hand-written lists, and the copy that drifted was nobody's
/// mistake at the point it was written down — the caller could not tell which list it was
/// supposed to be writing. Naming the reason and deriving the road from it makes the two the
/// same question, and it is a question with an answer: every reason but [`Reason::Hung`] is
/// the loop's own tick.
///
/// [`Reason::Hung`] is the one that is not, and it is not a choice anybody makes: it is the
/// watchdog's own. That is the whole of why the two ends are different roads — a loop that
/// has stopped turning cannot be asked for anything, so a kill made from outside it can only
/// post and hand the hide to a thread of its own.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(crate) enum Reason {
    /// The pin's own close: the caption's cross, a close sent to the window, or the key that
    /// closes it. Answered by the loop, because the window is the loop's and so is the player
    /// and the browser under it.
    Closed,
    /// Something outside the loop asked for it: previews were turned off, the trigger key is
    /// holding them back, the tray's `Pin Mode → Enable` row was switched off, or the machine
    /// resumed from sleep. The request is not the take-down — what a pin *is* belongs to the
    /// loop — so the loop is what ends it, on its next tick.
    Asked,
    /// The kind behind the pin, or the pin's own key, was switched off in the tray.
    SwitchedOff,
    /// The thing the pin is a window onto came apart: the player's process is gone, or the
    /// engine took a document and never drew it. This is the half of "until it is closed, or it
    /// comes apart" that is not a button.
    MediaGone,
    /// The loop stopped turning while a window the user is looking at was up, and a thread
    /// that has given up on the loop ended the pin anyway. The only reason whose road is not
    /// the loop's own.
    Hung,
}

impl Reason {
    /// Which of the two ends is taking this pin down.
    ///
    /// Derived rather than asked, so a caller cannot put the watchdog's road on a close or
    /// the loop's road on a kill — the mistake that let the two teardowns drift in the first
    /// place was each caller choosing.
    pub(super) fn exit(self) -> PinExit {
        match self {
            Reason::Hung => PinExit::Watchdog,
            _ => PinExit::Loop,
        }
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
/// *over* still has to be answerable for having been: `Ending` remembers why, which is what
/// makes the reason a caller passed testable rather than asserted. The pin's own value is
/// only in `Up`, so there is no way to write a state that is over and still holding a pin.
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
    /// A pin is over and this is the end that took it. Nothing is held: the state that could
    /// hold something has been taken by the move that got here.
    Ending(#[cfg_attr(not(test), allow(dead_code))] Reason),
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
/// called, and so is what the chrome had left for the loop: a command is a press on that
/// window's caption, and the window is the same one.
///
/// A pin's keyboard claim is *not* carried over. It was the old file's, and the new pin has
/// pressed nothing — so a swap taken over a pin with the caret on it leaves the caret on the
/// window and stops the pin claiming it, which is what `pin_take_focus` decides again on the
/// next press.
pub(super) fn install(pin: PinnedPreview) {
    if let Some(mut state) = pin_state() {
        let commands = match &mut *state {
            PinState::Up(up) => std::mem::take(&mut up.commands),
            _ => VecDeque::new(),
        };

        *state = PinState::Up(PinUp {
            pin,
            keyboard: None,
            commands,
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

    /// Why the last pin ended, where one has ended.
    #[cfg(test)]
    fn reason(&self) -> Option<Reason> {
        match self {
            PinState::Ending(reason) => Some(*reason),
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

    /// Take this pin down for `reason`, handing back what the window work still needs.
    ///
    /// The keyboard claim is *returned* rather than read afterwards, which is what makes a pin
    /// that is over unable to still be holding one: the claim leaves with the pin, and
    /// `Ending(_)` has nowhere to put it. The command queue goes the same way — it is a field
    /// of the pin, so a walk queued before a kill cannot be answered after it, because there
    /// is nothing left to answer it into.
    ///
    /// Whether there was a pin to take down goes back beside the claim, because the caller owes
    /// the window different work in the two cases and the claim alone cannot tell them apart (see
    /// [`PinEnded`]).
    fn end(&mut self, reason: Reason) -> PinEnded {
        let PinState::Up(mut up) = std::mem::replace(self, PinState::Ending(reason)) else {
            // A pin that is not up is already over, and the end it is being given is recorded
            // rather than refused: two roads racing on the same pin is normal (the watchdog and
            // the loop's own tick both watch for a hung one), and the second one must not be the
            // one that leaves the keyboard claimed.
            *self = PinState::Ending(reason);
            return PinEnded {
                was_up: false,
                keyboard: None,
            };
        };

        up.commands.clear();
        PinEnded {
            was_up: true,
            keyboard: up.keyboard,
        }
    }
}

/// What a road out of a pin took away, in the two shapes the window work has to tell apart.
///
/// Two fields rather than the one optional this replaced, because *a pin that held no keyboard* and
/// *no pin was up* are different answers and the road does different work for each. The first owes
/// the window its `WS_EX_NOACTIVATE` back — the style is only ever taken off in
/// [`hand_keyboard_back`], which runs for a claim and for nothing else — and the second owes it
/// nothing at all, because a road that finds the pin already over is a road that was told nothing
/// to begin with: the road that ended it has already ended it.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
struct PinEnded {
    /// Whether a pin was up when this was called, and this is the call that took it down.
    was_up: bool,
    /// The keyboard that pin was holding, if it was holding one.
    keyboard: Option<PinKeyboard>,
}

/// Give the pin the keyboard, if the keyboard arrived in it: the window that was in front, and
/// the pin's claim to it.
///
/// Whether it arrived is the caller's to report and not this module's to assume, because Windows
/// is what decides it — the foreground lock hands the foreground to whoever had the last input,
/// and a process that has not had it is refused when it asks — and because no test on this side of
/// the module can put a caret into a window that does not exist here. A claim written on the
/// *asking* is a pin that believes it holds a keyboard it was never given, holding a note of a
/// window the user is no longer in: the handover that note exists for is then a foreground asked
/// back from a window the user left, and it lands there.
///
/// The claim and the window it came from are written together, which is what the two globals
/// this replaced did not do: `PIN_PREVIOUS_FOREGROUND` was written before the `SetFocus` and
/// `PIN_FOCUSED` after it, so a window procedure re-entered by either call — and both of them
/// deliver messages — could see one without the other.
pub(super) fn take_keyboard(behind: isize, arrived: bool) {
    if !arrived {
        return;
    }

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
/// This is the whole of a road out of a pin. The state goes first and with it everything the
/// pin was holding — the keyboard claim, the command queue — because the three callers that
/// used to write this out by hand had drifted apart by four items, and the copy that had
/// drifted was the one made off the loop. Then the window work happens, and what that work is
/// says only which of the two ends is doing it.
///
/// Nothing here takes the pin's lock across a window call. That is not tidiness: the window
/// calls here deliver messages back into this same thread's window procedure, which asks for
/// the pin's lock on every one of them (see `ReleaseCapture` and `WM_CAPTURECHANGED`), so a
/// guard held across them is a thread waiting on a lock it owns. The claim is taken out under
/// the lock and the keyboard is handed back outside it, which is the same order the old
/// `pin_drop_focus` was written in and for the same reason.
///
/// The lock is read *through* where it is poisoned, which is the one place it is, and the
/// difference is the whole of what a poisoning is: a thread that unwound mid-teardown left the
/// pin up with its keyboard claimed, and a pin that cannot be taken down because the lock it is
/// in is poisoned is the defect this module was written to remove.
pub(super) fn end_pin(reason: Reason, window: &dyn PinWindow) -> PinHide {
    let hwnd = window.hwnd();
    let exit = reason.exit();

    // The pointer goes before the state does, and only where this thread is the one holding
    // it: `ReleaseCapture` delivers `WM_CAPTURECHANGED` back into this thread's own window
    // procedure, and that procedure ends a drag by looking the pin up — a release taken after
    // the state has gone finds nothing to end.
    if exit.may_release_pointer() && hwnd != 0 {
        window.release_capture(hwnd);
    }

    // The state, and everything the pin was holding, in one move under one lock. This is the
    // list the watchdog's hand-written copy of had four items missing from.
    let ended = PIN_STATE
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
        .end(reason);

    if ended.was_up {
        match ended.keyboard {
            // The keyboard goes back to the window it came from, outside the lock: `SetFocus` and
            // `SetForegroundWindow` are messages to this app's own window and to the one behind
            // it, and both arrive back here.
            Some(keyboard) => hand_keyboard_back(&keyboard, hwnd, window),
            // And a pin that was holding none still owes the window its style back. `SetFocus` is
            // not asked for first here because there is no focus of ours to give up: the claim is
            // written on the answer Windows gave rather than on the asking (see `take_keyboard`),
            // so a pin holding none is not the window the keyboard is in.
            None => make_the_window_unfocusable_again(hwnd, window),
        }
    }
    // A road that finds the pin already over is told nothing about the window: the road that ended
    // it has already ended it, style and all, and the second road is a race the watchdog and the
    // loop's own tick are both entitled to lose (see `two_roads_racing_on_one_pin_end_it_once`).

    PIN_UP.store(false, Ordering::Release);
    // A pin that is over is a pointer that is on something new: the file the pin was of is not
    // a hover the hook has already answered, and one is due the moment the pin is gone rather
    // than after the delay a re-hover of the same file is given (see `PIN_RESUMED`).
    PIN_RESUMED.store(true, Ordering::Release);

    // What the pin left behind that is not the pin: the walk the planner is working on, the
    // bubble's drag latch and the box a drag had left the window at. Each is somebody else's
    // state and each has its own owner, so the list is theirs — but it is called from here, so
    // there is still one place a road out of a pin is written down.
    super::end_pin_beside_the_state();

    match exit {
        PinExit::Loop => {
            // The loop's own take-down is the ordinary one a hover's dismissal goes through, so
            // the bubble — which is a window of this app's own, standing in for a pin that has
            // gone — goes with it here rather than with the message below. A collapsed pin
            // whose state has been cleared but whose bubble is still on screen is a window
            // nothing will ever take away again.
            window.hide_pin_bubble();
            PinHide::WithTheTakeDown
        }
        PinExit::Watchdog => {
            // Asked by message rather than taken: the loop being given up on is the thing that
            // must not be waited on, and a loop that was slow rather than gone does come back
            // to a window that is still holding the pointer.
            if hwnd != 0 {
                window.post(hwnd, WM_PIN_RELEASE_POINTER);
            }
            PinHide::OnItsOwnThread
        }
    }
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
///
/// This is the only place the style goes back on, which used to leave the pin nobody pressed
/// focusable after it came down — see [`make_the_window_unfocusable_again`] for what that costs
/// and why the road out of a claim-less pin asks for it of itself.
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

/// Put `WS_EX_NOACTIVATE` back on a window a pin took it off, for a pin that held no keyboard.
///
/// The style is only ever taken off in [`hand_keyboard_back`], which runs for a claim and for
/// nothing else, so a road out of a pin that was holding none went down leaving a window of this
/// app's own focusable while it was showing nothing. It is a `WS_EX_TOOLWINDOW` popup with
/// `WS_EX_TOPMOST` and no owner, so an alt-tab, a `SetForegroundWindow` asked for from somewhere
/// else, or a click landing on the area of a window that is hidden all move the caret out of
/// whatever the user is typing in — which is the one thing the style is on that window for (see
/// `pin_set_focusable`).
///
/// Claim-less rather than *pin-gone*, and the collapse is the difference. A pin that is put away
/// to its bubble still holds its claim until [`give_the_keyboard_back`] takes it, and that road
/// already goes through [`hand_keyboard_back`] and so already puts the style back: the pinned
/// window is hidden, and the bubble standing in for it is a window of its own with the style
/// already on it. So "no claim is held" is the condition that means *this window is not one the
/// user is in* on both roads, where "no pin is up" is not a condition at all — it is true of a
/// collapsed pin, which is a pin.
///
/// No focus is given up first, and there is none to give up: a claim is written on what Windows
/// answered rather than on what was asked for, so a pin holding no claim is a pin the keyboard
/// never arrived in (see `take_keyboard`).
fn make_the_window_unfocusable_again(hwnd: isize, window: &dyn PinWindow) {
    if hwnd == 0 {
        return;
    }

    window.set_focusable(hwnd, false);
}

/// Whether the pin holds a claim on the keyboard, which is this app's own record of a press that
/// handed it the keyboard — and not the question of where the caret is now, which is
/// `preview_window::pin_is_focused` and which can be answered without this module at all.
///
/// The two are read together for exactly one question, which is whether a pin that was given the
/// keyboard still has it: a claim that stands while the pin has been activated away from is a claim
/// that has stopped being true, and the pin's own arrows are then the keys of whatever is in front
/// of it (see `preview_window::pin_keyboard_wanted_back`).
pub(super) fn pin_claims_the_keyboard() -> bool {
    pin_state().is_some_and(|mut state| state.keyboard_mut().is_some_and(|held| held.is_some()))
}

/// Whether the pin holds a claim on the keyboard, for the tests on this side of the module,
/// which cannot ask the desktop where the caret is.
#[cfg(test)]
pub(super) fn pin_holds_a_keyboard() -> bool {
    pin_claims_the_keyboard()
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

/// The lock every test that shares the preview's own state takes, so that one of them runs at a
/// time: what is shared is process-wide and there is one copy of it, so two of these at once is
/// one test's press answered by another's window.
///
/// That state is the media the preview is showing (`CURRENT_MEDIA`), the pin's slot and the
/// published copy of whether one is up (`PIN_STATE` and `PIN_UP`, reached through `stand_pin`
/// and `pin_state`), and the press the Explorer hook publishes for the pin's drag
/// (`publish_pin_media_press`). A pin is the clearest case of what goes wrong without this. A
/// test stands a pin up through `stand_pin`, which writes the slot whole, and a second test
/// doing the same at the same moment does not merge with it — it replaces it, so the first goes
/// on to assert against the second's pin and is answered by it. There is no torn read to appeal
/// to here: both takes are whole, and both take the slot's own lock. Two tests each holding a
/// lock, and neither holding the other's answer.
///
/// So the lock is wanted of every test that reaches that state and not only of the ones that
/// write it. A test that merely asks `pinned()`, or reads `CURRENT_MEDIA` to see what kind of
/// media is up, is still asserting on what it found, and what it found is whatever the test
/// beside it left standing: a hover asked while a pin is up is answered by `show_preview`'s
/// early return, and reports no hover at all rather than a wrong one. Reading is the half that
/// is easy to leave out, and it fails the same way writing does.
///
/// There is one lock for this state, not one per module, and it is here because this module owns
/// the state: the slot is the one declared above, and the tests that stand a pin up stand it from
/// both sides of the module wall. A second lock over the same slot is not a second line of
/// defence, it is no defence at all — a test in `preview_window` holding that one and a test here
/// holding this one each believe they are alone with the pin, and are not. The fix for that is
/// not a second mutex, which is what this file had alongside the one in `preview_window`, but
/// this one mutex with both modules' tests reaching it by path.
///
/// A poisoned lock is read through rather than refused. This mutex holds no state, only the order
/// of the tests around it, so a panic inside one of them leaves nothing behind to recover but
/// that order — and refusing the lock would turn one failed assertion into every later test of
/// the group, each turned away at the door over a hold it never needed. What is being protected
/// is the order, so the order is taken back.
///
/// The rule for a test that reaches this state, then: does it stand a pin up, take one down,
/// write the media, or ask any of the above? Take this lock on the first line of the body, before
/// anything touches that state. A test that touches none of it takes nothing — most of both
/// modules is arithmetic over values handed to it — so the suite is not serialised, only the part
/// of it that shares a machine. A lock over some other process-wide thing, like the pointer's own
/// stand-in in `preview_window`, is a different lock and neither stands in for this one.
///
/// And holding the lock is half of what is owed to the test that comes next: a test that stands a
/// pin up puts it down again before it lets the lock go. Serialising the tests puts them one at a
/// time, which is what stops them answering one another mid-test, but it does nothing about what
/// one of them leaves standing for the next one to walk into — and a test that asserts "no pin to
/// read at first" is asserting about whatever the test before it did not clean up. That is an
/// order dependence rather than a race, so it hides from a suite that happens to run in a lucky
/// order and shows up as an occasional red under load, where thread scheduling decides who runs
/// next. Three tests here stood a pin up to ask about its queue and left it up, and the read of
/// the empty slot is what caught it.
#[cfg(test)]
pub(super) static PIN_TESTS_ONE_AT_A_TIME: Mutex<()> = Mutex::new(());

#[cfg(test)]
mod tests;

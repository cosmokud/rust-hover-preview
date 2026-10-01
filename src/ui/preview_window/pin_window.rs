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
//! It grew by four for a reason worth recording, because the four are not lifecycle operations
//! at all: they are *a drag*, which is the other thing this window is asked of a pin and the
//! only one that had no test surface. `pointer`, `window_box`, `capture` and `repaint` are the
//! whole of what "a press becomes a carried drag, and the drag is let go of" costs the machine.
//! They are here rather than in a second trait over the same window for the reason the two ends
//! of the capture are here: they are two halves of one handoff, and a handoff whose halves are
//! asked of two different objects is a handoff that can disagree with itself — which is what a
//! window left holding the pointer for the whole desktop was (see `begin_pin_drag`).
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
    fn end(&mut self, reason: Reason) -> Option<PinKeyboard> {
        let PinState::Up(mut up) = std::mem::replace(self, PinState::Ending(reason)) else {
            // A pin that is not up is already over, and the end it is being given is recorded
            // rather than refused: two roads racing on the same pin is normal (the watchdog and
            // the loop's own tick both watch for a hung one), and the second one must not be the
            // one that leaves the keyboard claimed.
            *self = PinState::Ending(reason);
            return None;
        };

        up.commands.clear();
        up.keyboard
    }
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
    let keyboard = PIN_STATE
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
        .end(reason);

    // The keyboard goes back to the window it came from, outside the lock: `SetFocus` and
    // `SetForegroundWindow` are messages to this app's own window and to the one behind it,
    // and both arrive back here.
    if let Some(keyboard) = keyboard {
        hand_keyboard_back(&keyboard, hwnd, window);
    }

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
    /// Where the pointer is, for a drag measured off it.
    pointer: Option<(i32, i32)>,
    /// Where the window stands, for a drag begun from the screen's own box rather than the pin's.
    box_: Option<ScreenRegion>,
    calls: Mutex<Vec<PinWindowCall>>,
}

#[cfg(test)]
impl RecordedPinWindow {
    /// A recorder standing in for a window that is there, or for one that is not.
    pub(super) fn new(hwnd: isize) -> Self {
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
    pub(super) fn with(
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

    /// Every reason there is, which is what a road out of a pin is asked for.
    const EVERY_REASON: [Reason; 5] = [
        Reason::Closed,
        Reason::Asked,
        Reason::SwitchedOff,
        Reason::MediaGone,
        Reason::Hung,
    ];

    /// The pin's own value, with nothing asked of it and nothing answered about it.
    fn a_pin() -> PinnedPreview {
        PinnedPreview::for_test()
    }

    /// A pin, up, with the keyboard taken from `behind` and one command waiting.
    fn a_pin_holding_the_keyboard(behind: isize) {
        install(a_pin());
        take_keyboard(behind, true);
        ask_pin(PinCommand::Next);
    }

    /// The keyboard the pin is holding, and where it came from, or nothing while it holds none.
    ///
    /// Asked of the claim itself rather than of the pin, because a pin that is over holds none
    /// and a reader that could only see a pin could not tell "ended, and gave it back" from
    /// "ended, and kept it" — which is the whole of what these tests are asking.
    fn keyboard() -> Option<PinKeyboard> {
        match pin_state()?.keyboard_mut() {
            Some(keyboard) => *keyboard,
            None => None,
        }
    }

    /// Why the last pin ended, where one has.
    fn why() -> Option<Reason> {
        pin_state()?.reason()
    }

    /// The pointer goes before the state and only where this thread holds it.
    ///
    /// The order is load-bearing rather than incidental. `ReleaseCapture` delivers
    /// `WM_CAPTURECHANGED` back into the window procedure this same thread is running, and
    /// that procedure ends a drag by looking the pin up: a release taken after the state has
    /// gone finds nothing to end, and a window whose caption then answers nothing at all.
    ///
    /// And only the loop's own tick may do it at all. A capture belongs to the thread that
    /// took it, so a watchdog asking for one releases nothing — or, if some other window on
    /// that thread is holding one, releases that window's press and hands it to a pin the
    /// hand is not aimed at. That is why the watchdog's road asks by message instead.
    #[test]
    fn the_pointer_goes_before_the_state_and_only_where_this_thread_holds_it() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        let loop_window = RecordedPinWindow::new(0x1000);
        assert_eq!(
            end_pin(Reason::Closed, &loop_window),
            PinHide::WithTheTakeDown
        );
        assert_eq!(
            loop_window.calls(),
            vec![PinWindowCall::ReleaseCapture, PinWindowCall::HidePinBubble],
            "the release is taken while this thread can still end the drag it belongs to"
        );

        install(a_pin());
        let watchdog_window = RecordedPinWindow::new(0x1000);
        assert_eq!(
            end_pin(Reason::Hung, &watchdog_window),
            PinHide::OnItsOwnThread
        );
        assert_eq!(
            watchdog_window.calls(),
            vec![PinWindowCall::Post(WM_PIN_RELEASE_POINTER)],
            "a watchdog has no capture of its own to release, and cannot wait on the loop's"
        );
    }

    /// The whole of what each road owes the window, for a pin holding all of it.
    ///
    /// One test for the whole list rather than a sample, because the defect this seam exists
    /// for is a list that had become two: the watchdog's hand-written copy of the window work
    /// had four of the six items below missing or wrong, and every one of them was invisible
    /// until a pin was killed for being hung. A test that asserts the first item and moves on
    /// would have passed against the copy.
    #[test]
    fn the_whole_of_what_each_road_owes_the_window() {
        let _one = ONE_AT_A_TIME.lock();

        a_pin_holding_the_keyboard(0x2000);
        let loop_window = RecordedPinWindow::new(0x1000);
        end_pin(Reason::Closed, &loop_window);
        assert_eq!(
            loop_window.calls(),
            vec![
                PinWindowCall::ReleaseCapture,
                PinWindowCall::SetFocus,
                PinWindowCall::SetFocusable { focusable: false },
                PinWindowCall::SetForeground,
                PinWindowCall::HidePinBubble,
            ],
            "the loop releases the pointer, hands the keyboard back, and takes the bubble down \
             itself — the ordinary take-down that follows brings the pinned window with it"
        );

        a_pin_holding_the_keyboard(0x2000);
        let watchdog_window = RecordedPinWindow::new(0x1000);
        end_pin(Reason::Hung, &watchdog_window);
        assert_eq!(
            watchdog_window.calls(),
            vec![
                PinWindowCall::SetFocus,
                PinWindowCall::SetFocusable { focusable: false },
                PinWindowCall::SetForeground,
                PinWindowCall::Post(WM_PIN_RELEASE_POINTER),
            ],
            "the watchdog hands the keyboard back first — that item is not optional — and asks \
             for the pointer afterwards, because by then the drag it belonged to is gone. The \
             bubble it leaves to the thread of its own that will do the hiding."
        );
    }

    /// A road out of a pin with no window behind it asks it for nothing of that window.
    ///
    /// Zero is a handle this app has not been given yet, which is what the window procedure
    /// sees before the window exists and what the watchdog sees if the app is on its way down.
    /// A teardown that assumed a window would post a message to handle zero and try to release
    /// a capture of zero, both of which are calls on a handle that may since be somebody else's.
    ///
    /// The hide is asked for either way, because the bubble is a window of this app's own read
    /// from its own slot rather than from the handle the road was given: a take-down that
    /// skipped it where the pin's own window had already gone would leave a bubble standing that
    /// nothing would ever take away.
    ///
    /// The state settles regardless, which is the half that does not need a window at all: a pin
    /// is over whether or not there was anything to take down.
    #[test]
    fn a_road_with_no_window_asks_it_for_nothing() {
        let _one = ONE_AT_A_TIME.lock();

        for reason in EVERY_REASON {
            install(a_pin());
            take_keyboard(0x2000, true);
            let window = RecordedPinWindow::new(0);
            end_pin(reason, &window);

            let mut expected = Vec::new();
            if reason.exit() == PinExit::Loop {
                expected.push(PinWindowCall::HidePinBubble);
            }

            assert_eq!(
                window.calls(),
                expected,
                "{reason:?}: nothing is asked of a handle \
                that is not there"
            );
            assert!(!pin_is_up(), "{reason:?}: and the pin is still over");
        }
    }

    /// The pin's own lock is not held across a window call.
    ///
    /// The window calls a teardown makes deliver messages back into this same thread's
    /// window procedure, which asks for the pin's lock on every one of them. A guard held
    /// across them is a thread waiting on a lock it owns — so the claim is taken out under
    /// the lock and the keyboard is handed back outside it, and the order is observable: by
    /// the time the focus is given up, the pin is already gone.
    #[test]
    fn the_pin_s_own_lock_is_not_held_across_a_window_call() {
        let _one = ONE_AT_A_TIME.lock();

        a_pin_holding_the_keyboard(0x2000);
        let window = LookingWindow::new();
        end_pin(Reason::Closed, &window);

        assert_eq!(
            window.looked(),
            vec![true, false, false, false],
            "the pointer is let go while the pin is still there to end the drag for, and every \
             window call after it is made with the state already gone and the lock let go — \
             which is what lets a window procedure re-entered by one of them find no pin and \
             end nothing twice"
        );
    }

    /// A recorder that looks the pin up on every call it is asked for, standing in for the
    /// window procedure every one of these calls is delivered back into.
    struct LookingWindow {
        inner: RecordedPinWindow,
        /// Whether a pin was up at each call, in the order the calls were made.
        looked: Mutex<Vec<bool>>,
    }

    impl LookingWindow {
        fn new() -> Self {
            Self {
                inner: RecordedPinWindow::new(0x1000),
                looked: Mutex::new(Vec::new()),
            }
        }

        /// Whether a pin was up, for each call, in the order the calls were made.
        fn looked(&self) -> Vec<bool> {
            self.looked
                .lock()
                .map(|looked| looked.clone())
                .unwrap_or_default()
        }

        /// What a window procedure does on every message it is sent: look the pin up.
        fn looking(&self) {
            if let Ok(mut looked) = self.looked.lock() {
                looked.push(pin_state().is_some_and(|state| state.is_up()));
            }
        }
    }

    impl PinWindow for LookingWindow {
        fn hwnd(&self) -> isize {
            self.inner.hwnd()
        }

        fn pointer(&self) -> Option<(i32, i32)> {
            self.inner.pointer()
        }

        fn window_box(&self, hwnd: isize) -> Option<ScreenRegion> {
            self.inner.window_box(hwnd)
        }

        fn capture(&self, hwnd: isize) {
            self.inner.capture(hwnd)
        }

        fn release_capture(&self, hwnd: isize) {
            self.looking();
            self.inner.release_capture(hwnd);
        }

        fn set_focusable(&self, hwnd: isize, focusable: bool) {
            self.looking();
            self.inner.set_focusable(hwnd, focusable);
        }

        fn set_focus(&self, hwnd: isize) {
            self.looking();
            self.inner.set_focus(hwnd);
        }

        fn set_foreground(&self, hwnd: isize) {
            self.looking();
            self.inner.set_foreground(hwnd);
        }

        fn hide_pin_windows(&self) {
            self.inner.hide_pin_windows();
        }

        fn hide_pin_bubble(&self) {
            self.inner.hide_pin_bubble();
        }

        fn repaint(&self) {
            self.inner.repaint();
        }

        fn post(&self, hwnd: isize, message: u32) {
            self.inner.post(hwnd, message);
        }
    }

    /// A pin that has just come up answers that it is up, and a pin that is over answers that
    /// it is not — and the two copies of that answer never disagree.
    ///
    /// The Explorer hook asks on every tick and cannot be answered off the lock, so the phase
    /// is published as well as held, and the two are written in one order at both ends. That
    /// pairing is only checkable from here: a flag that ran ahead of the state would have
    /// previews held back over no window, and one that ran behind would have hovers answered
    /// over a window nothing would ever take down.
    #[test]
    fn the_published_flag_and_the_state_never_disagree() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        assert!(pin_is_up(), "installed: a pin is published as up");
        assert_eq!(
            pin_is_up(),
            pin_state().is_some_and(|state| state.is_up()),
            "and the state agrees"
        );

        end_pin(Reason::Closed, &RecordedPinWindow::new(0x1000));
        assert!(!pin_is_up(), "ended: and as down");
        assert_eq!(
            pin_is_up(),
            pin_state().is_some_and(|state| state.is_up()),
            "and the state agrees"
        );
    }

    /// A pin taken up over another one is the same window, showing another file.
    ///
    /// The pin it was is one window, so what belongs to the window rather than to the file — a
    /// maximized box, a level, chrome, a bound — is carried over by the caller before the
    /// take-up, and so is what the chrome had left for the loop: a command is a press on that
    /// window's caption, and the window is the same one.
    ///
    /// What does not carry is the claim on the keyboard. It was the old file's, and the new pin
    /// has pressed nothing — so a swap taken over a pin with the caret on it stops that pin
    /// claiming it, rather than leaving a claim on a file the window is no longer showing.
    #[test]
    fn a_pin_taken_up_over_another_one_keeps_the_window_s_own_queue_and_drops_its_claim() {
        let _one = ONE_AT_A_TIME.lock();

        a_pin_holding_the_keyboard(0x2000);
        install(a_pin());

        assert!(pin_is_up(), "the swap is a pin up like any other");
        assert_eq!(
            take_pin_command(),
            Some(PinCommand::Next),
            "the command the caption had left is a press on this window, and this window is \
             still here"
        );
        assert_eq!(
            keyboard(),
            None,
            "but the keyboard claim is not carried: it was the old file's, and the new pin has \
             pressed nothing"
        );
    }

    /// Every command a caption's button asks for survives the trip out of the window procedure
    /// and back.
    ///
    /// The codes are the whole of the crossing — the loop and the window procedure share
    /// nothing else — so a button whose code nothing reads back is a button that does nothing.
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

        // Two asks are two commands, in the order they were made. A single slot dropped the
        // first: a double-click on `Next` is two `WM_LBUTTONUP`s, and moving two files along is
        // what double-clicking a `Next` is for.
        ask_pin(PinCommand::Previous);
        ask_pin(PinCommand::Next);
        assert_eq!(take_pin_command(), Some(PinCommand::Previous));
        assert_eq!(take_pin_command(), Some(PinCommand::Next));
        assert_eq!(take_pin_command(), None, "a command is taken once");
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

    /// The keyboard goes back to the window it came from.
    ///
    /// A window that is hidden while it still holds the focus leaves Windows to pick what to
    /// activate next, and for a `WS_EX_TOOLWINDOW` popup that is not reliably the Explorer
    /// window that was in front a moment ago.
    #[test]
    fn the_keyboard_goes_back_to_the_window_it_came_from() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        take_keyboard(0x2000, true);
        let window = RecordedPinWindow::new(0x1000);
        end_pin(Reason::Closed, &window);
        assert_eq!(
            window.calls(),
            vec![
                PinWindowCall::ReleaseCapture,
                PinWindowCall::SetFocus,
                PinWindowCall::SetFocusable { focusable: false },
                PinWindowCall::SetForeground,
                PinWindowCall::HidePinBubble,
            ],
            "the window behind is put back in front, so a keyboard handed back lands where it \
             came from"
        );
    }

    /// A pin that took the keyboard from nothing hands it back to nothing.
    ///
    /// There is sometimes no window behind: a `WS_POPUP` with no parent has no `GW_OWNER` at
    /// all, so the handover is the focus going to nothing rather than to a window that was
    /// never there.
    #[test]
    fn a_pin_that_took_the_keyboard_from_nothing_hands_it_back_to_nothing() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        take_keyboard(0, true);
        let window = RecordedPinWindow::new(0x1000);
        end_pin(Reason::Closed, &window);

        assert!(
            window
                .calls()
                .iter()
                .all(|call| !matches!(call, PinWindowCall::SetForeground)),
            "there is no window behind a pin that took the keyboard from nothing"
        );
    }

    /// A pin with nothing on the keyboard owes nobody a handover.
    ///
    /// A pin nobody pressed holds no keyboard to give back, so a teardown of one does not go
    /// near the focus at all — which is the item the old `pin_drop_focus` asked `PIN_FOCUSED`
    /// about before doing anything.
    #[test]
    fn a_pin_with_nothing_on_the_keyboard_owes_nobody_a_handover() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        let window = RecordedPinWindow::new(0x1000);
        end_pin(Reason::Closed, &window);

        assert_eq!(
            window.calls(),
            vec![PinWindowCall::ReleaseCapture, PinWindowCall::HidePinBubble],
            "nothing is asked of a window this pin never took anything from"
        );
    }

    /// The note that the keyboard was taken is dropped with the pin, or with the focus.
    ///
    /// Windows taking the focus away is the user clicking into something else: the window now
    /// in front holds the keyboard, so there is nothing to hand over and the claim goes. That
    /// is the one road out that does not do the handover, and it is the same shape as ending
    /// the pin — one field that says whether there is a keyboard to give back, rather than a
    /// bit and a handle that could be set one without the other.
    #[test]
    fn the_note_that_the_keyboard_was_taken_is_dropped_with_the_focus() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        take_keyboard(0x2000, true);
        release_keyboard();
        assert_eq!(keyboard(), None, "losing the focus drops the claim with it");

        a_pin_holding_the_keyboard(0x2000);
        end_pin(Reason::Closed, &RecordedPinWindow::new(0x1000));
        assert_eq!(keyboard(), None, "a pin that is over holds no keyboard");
    }

    /// A press the window was refused writes no claim, and a pin holding none hands nothing back.
    ///
    /// The foreground lock refuses the press to a process that did not have the last input, which
    /// is what a press on a pin looks like while something else owns the input rather than a fault
    /// in anything. The claim used to be written on the *asking*, so that pin came away from the
    /// press believing it held a keyboard and holding a note of a window the user is no longer in
    /// — and the take-down then asked that window for the foreground, taking the focus from
    /// whatever the user had moved on to. The arrival is the caller's to report and this module's
    /// to insist on: the claim is written where the keyboard demonstrably is and nowhere else, and
    /// a caller that has not asked Windows cannot write one.
    #[test]
    fn a_press_the_windows_refused_leaves_no_claim_to_hand_over() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        take_keyboard(0x2000, false);
        assert_eq!(
            keyboard(),
            None,
            "a keyboard that never arrived is not one to hold a note of a window for"
        );

        let window = RecordedPinWindow::new(0x1000);
        end_pin(Reason::Closed, &window);
        assert_eq!(
            window.calls(),
            vec![PinWindowCall::ReleaseCapture, PinWindowCall::HidePinBubble],
            "and a pin with no claim asks the window the user left for nothing, rather than \
             bringing it back to the front"
        );
    }

    /// A pin taken for a reason is ended by the loop's own tick, and a pin the loop has given
    /// up on is ended by a thread of its own — and no other pairing exists.
    #[test]
    fn every_reason_is_the_loop_s_own_tick_except_the_watchdog_s() {
        for (reason, exit) in [
            (Reason::Closed, PinExit::Loop),
            (Reason::Asked, PinExit::Loop),
            (Reason::SwitchedOff, PinExit::Loop),
            (Reason::MediaGone, PinExit::Loop),
            (Reason::Hung, PinExit::Watchdog),
        ] {
            assert_eq!(reason.exit(), exit, "{reason:?} is {exit:?}'s road");
        }
    }

    /// Two roads racing on one pin end it once.
    ///
    /// The watchdog and the loop's own tick both watch for a hung loop, and either may be
    /// first. The second one must find nothing to take down rather than refusing, and must not
    /// be the one that leaves the keyboard claimed.
    #[test]
    fn two_roads_racing_on_one_pin_end_it_once() {
        let _one = ONE_AT_A_TIME.lock();

        a_pin_holding_the_keyboard(0x2000);
        end_pin(Reason::Hung, &RecordedPinWindow::new(0x1000));

        let second = RecordedPinWindow::new(0x1000);
        end_pin(Reason::Hung, &second);

        assert_eq!(keyboard(), None, "the pin is over either way");
        assert_eq!(
            second.calls(),
            vec![PinWindowCall::Post(WM_PIN_RELEASE_POINTER)],
            "and the second road is told nothing about a window that is already gone, because \
             the pointer release it asks for is a request and there is nothing to release"
        );
    }

    /// Every reason reaches the same teardown, and nothing is left standing after it.
    ///
    /// One list, called from every exit: the state, the keyboard claim, the command queue and
    /// the bubble, whatever asked for it. This is the locality the review asked for — one
    /// teardown, several callers — and it is the property that fails first if a road is written
    /// without it, which is not a hypothetical: the watchdog's hand-written copy of this list
    /// cleared the pin and its flags and left the keyboard it had claimed and the walk a caption
    /// button had queued, so a pin killed for being hung ended still holding a caret and a walk
    /// to be answered into whatever pin came next.
    ///
    /// Every reason and not one of them, and the whole of what a pin was holding and not a
    /// sample of it, because each of those four items was one the copy had wrong.
    #[test]
    fn every_reason_reaches_the_same_teardown() {
        let _one = ONE_AT_A_TIME.lock();

        for reason in EVERY_REASON {
            a_pin_holding_the_keyboard(0x2000);
            end_pin(reason, &RecordedPinWindow::new(0x1000));

            assert!(!pin_is_up(), "{reason:?}: the pin is over");
            assert_eq!(
                keyboard(),
                None,
                "{reason:?}: the keyboard went with it, so there is nothing for a later pin to \
                 inherit and nothing for a reader to find"
            );
            assert_eq!(
                take_pin_command(),
                None,
                "{reason:?}: so did the queue — a command left behind is a command about a window \
                 that is gone, fired against whatever pin comes next"
            );
            assert_eq!(
                why(),
                Some(reason),
                "{reason:?}: and the end says which end it was, which is the whole of what a \
                 reason handed to a teardown is for"
            );
        }
    }

    /// A collapsed pin is still a pin, and still goes down the same way.
    ///
    /// A pin that has been put away is not up but is not gone: the loop's own tick still has
    /// to be able to end it, and the end has to settle everything an un-collapsed one holds.
    /// The bubble is a window of this app's own and does not hold the keyboard, so a pin
    /// collapsed with the caret on it still has to hand it back.
    #[test]
    fn a_collapsed_pin_is_still_a_pin_and_still_goes_down_the_same_way() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        take_keyboard(0x2000, true);
        ask_pin(PinCommand::Close);
        if let Some(mut state) = pin_state() {
            if let Some(pin) = state.pin_mut() {
                pin.collapsed = true;
            }
        }

        let window = RecordedPinWindow::new(0x1000);
        end_pin(Reason::Closed, &window);

        assert!(!pin_is_up());
        assert_eq!(keyboard(), None);
        assert_eq!(take_pin_command(), None);
        assert!(
            window.calls().contains(&PinWindowCall::SetForeground),
            "a bubble does not hold the keyboard, so a pin collapsed with the caret on it still \
             has to hand it back"
        );
    }

    /// A pin whose window is a bubble takes the bubble with it, on the road that owns the
    /// thread.
    ///
    /// A collapsed pin whose state has been cleared but whose bubble is still on screen is a
    /// window nothing will ever take away again, so the loop's own road takes it down itself
    /// rather than leaving it for the ordinary take-down, which does not know about it. The
    /// watchdog's road cannot, and its own thread's hide covers it (see `hide_pin_windows`).
    #[test]
    fn a_pin_whose_window_is_a_bubble_takes_the_bubble_with_it() {
        let _one = ONE_AT_A_TIME.lock();

        install(a_pin());
        let window = RecordedPinWindow::new(0x1000);
        end_pin(Reason::Closed, &window);

        assert!(
            window.calls().contains(&PinWindowCall::HidePinBubble),
            "the bubble goes with the pin, on the road that owns the thread"
        );
    }

    /// The pin's own value is readable while it is up and gone once it is over.
    #[test]
    fn the_pin_s_own_state_is_readable_while_it_is_up_and_gone_once_it_is_over() {
        let _one = ONE_AT_A_TIME.lock();

        assert!(
            pin_state().is_none_or(|state| state.pin().is_none()),
            "no pin to read at first"
        );

        install(a_pin());
        assert!(
            pin_state().is_some_and(|state| state.pin().is_some()),
            "a pin that is up has a window and a state of its own to read"
        );

        end_pin(Reason::Closed, &RecordedPinWindow::new(0x1000));
        assert!(
            pin_state().is_some_and(|state| state.pin().is_none()),
            "and a pin that is over has neither, which is the whole of what Ending means"
        );
    }

    /// What a pin's end publishes for the Explorer hook, and what it does not.
    ///
    /// A pin that is over is a pointer that is on something new: the file the pin was of is
    /// not a hover the hook has already answered, and one is due the moment the pin is gone
    /// rather than after the delay a re-hover of the same file is given. That is a fact about
    /// the pin's end and nothing else, so it is published here rather than left to the caller
    /// that happened to remember to do it.
    #[test]
    fn a_pin_s_end_is_published_to_the_hook_once() {
        let _one = ONE_AT_A_TIME.lock();

        // Whatever a test elsewhere on the machine left published is drained first, since the
        // flag is process-wide and this one is about what an end does rather than about the
        // flag having been clear to begin with.
        take_pin_resumed();

        install(a_pin());
        end_pin(Reason::Closed, &RecordedPinWindow::new(0x1000));
        assert!(
            take_pin_resumed(),
            "a pin that is over is a hover the hook has not answered yet"
        );
        assert!(
            !take_pin_resumed(),
            "and it is said once, because the hook reads it once per tick and a second true \
             would be a file the pointer never left"
        );
    }
}

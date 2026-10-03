use super::*;

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
///
/// The style in each list is the window being made unfocusable again on a pin that was holding
/// no keyboard at all (see `make_the_window_unfocusable_again`); it is on the list because it
/// is owed on every one of these roads, and on the loop's it comes after the release for the
/// same reason the state does — the release is a message back into this thread.
#[test]
fn the_pointer_goes_before_the_state_and_only_where_this_thread_holds_it() {
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    install(a_pin());
    let loop_window = RecordedPinWindow::new(0x1000);
    assert_eq!(
        end_pin(Reason::Closed, &loop_window),
        PinHide::WithTheTakeDown
    );
    assert_eq!(
        loop_window.calls(),
        vec![
            PinWindowCall::ReleaseCapture,
            PinWindowCall::SetFocusable { focusable: false },
            PinWindowCall::HidePinBubble,
        ],
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
        vec![
            PinWindowCall::SetFocusable { focusable: false },
            PinWindowCall::Post(WM_PIN_RELEASE_POINTER),
        ],
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
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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

    fn unpark_player_window(&self, band: Option<ScreenRegion>) {
        self.inner.unpark_player_window(band);
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
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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

    stand_pin(None);
}

/// Every command a caption's button asks for survives the trip out of the window procedure
/// and back.
///
/// The codes are the whole of the crossing — the loop and the window procedure share
/// nothing else — so a button whose code nothing reads back is a button that does nothing.
#[test]
fn every_command_a_caption_asks_for_comes_back_to_the_loop() {
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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

    stand_pin(None);
}

/// A command queue drops the oldest rather than growing without end.
///
/// A loop that cannot keep up with a hand drumming on a button must not be the reason a
/// session ends. What goes is the oldest, so what survives is what was last asked for.
#[test]
fn a_command_queue_drops_the_oldest_rather_than_growing_without_end() {
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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

    stand_pin(None);
}

/// The keyboard goes back to the window it came from.
///
/// A window that is hidden while it still holds the focus leaves Windows to pick what to
/// activate next, and for a `WS_EX_TOOLWINDOW` popup that is not reliably the Explorer
/// window that was in front a moment ago.
#[test]
fn the_keyboard_goes_back_to_the_window_it_came_from() {
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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
/// about before doing anything. The style is the one thing it still owes, because a window
/// left focusable is a window an alt-tab can take the caret with (see
/// `make_the_window_unfocusable_again`).
#[test]
fn a_pin_with_nothing_on_the_keyboard_owes_nobody_a_handover() {
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    install(a_pin());
    let window = RecordedPinWindow::new(0x1000);
    end_pin(Reason::Closed, &window);

    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::ReleaseCapture,
            PinWindowCall::SetFocusable { focusable: false },
            PinWindowCall::HidePinBubble,
        ],
        "the focus is not touched, because there was never a claim on it to give up — but the \
         style goes back, because a pin the hand never pressed never took it off and must not \
         leave it off"
    );
}

/// Every road out of a pin that was holding no keyboard still puts the style back, in the place
/// on that road where the handover would have gone.
///
/// The bug this pins down is a pin that came down focusable. `hand_keyboard_back` is the only
/// place the style goes back and it runs only for a claim, so every pin the hand never pressed
/// — and, since the claim is written on what Windows answered rather than on what was asked,
/// every pin whose press Windows refused — went down leaving a window of this app's own
/// activatable while it was showing nothing. On a `WS_EX_TOOLWINDOW` popup with
/// `WS_EX_TOPMOST` and no owner that is an alt-tab, a `SetForegroundWindow` asked for from
/// elsewhere, or a click landing on the hidden window's area, each moving the caret out of
/// whatever the user is typing in.
///
/// Every reason and not one of them, and the whole recorded list rather than a search of it,
/// because where on each road the style belongs is the part that differs between them and the
/// part a hand-written copy of this list gets wrong first. On the loop's road it goes after the
/// pointer release and before the bubble comes down; on the watchdog's there is no pointer to
/// release, so it goes first and the request for the loop's comes after — the same place the
/// handover itself would have gone on each.
#[test]
fn a_road_out_of_a_pin_holding_no_keyboard_still_puts_the_style_back() {
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    for reason in EVERY_REASON {
        // A pin the hand never pressed, which is the ordinary one: nothing is claimed, so the
        // handover `hand_keyboard_back` performs never runs at all.
        install(a_pin());
        assert!(!pin_holds_a_keyboard(), "{reason:?}: a pin nobody pressed");

        let mut expected = Vec::new();
        if reason.exit() == PinExit::Loop {
            expected.push(PinWindowCall::ReleaseCapture);
        }
        expected.push(PinWindowCall::SetFocusable { focusable: false });
        if reason.exit() == PinExit::Loop {
            expected.push(PinWindowCall::HidePinBubble);
        } else {
            expected.push(PinWindowCall::Post(WM_PIN_RELEASE_POINTER));
        }

        let window = RecordedPinWindow::new(0x1000);
        end_pin(reason, &window);
        assert_eq!(
            window.calls(),
            expected,
            "{reason:?}: the style goes back where the handover would have gone, and nothing \
             else is asked of a pin that took nothing"
        );

        // And the other road into the same state: a press Windows refused writes no claim
        // (see `take_keyboard`), so this pin came down through the branch above too.
        install(a_pin());
        take_keyboard(0x2000, false);
        assert!(!pin_holds_a_keyboard(), "{reason:?}: a refused press");

        let refused = RecordedPinWindow::new(0x1000);
        end_pin(reason, &refused);
        assert_eq!(
            refused.calls(),
            expected,
            "{reason:?}: and a press Windows refused is the same road as never pressing at all"
        );
    }
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
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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
///
/// It is also the road that made the style worth restoring here: the press took
/// `WS_EX_NOACTIVATE` off the window before it asked for the foreground, so a refused press
/// left a pin that was focusable and holding nothing, and its take-down had nothing to restore
/// it from (see `make_the_window_unfocusable_again`).
#[test]
fn a_press_the_windows_refused_leaves_no_claim_to_hand_over() {
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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
        vec![
            PinWindowCall::ReleaseCapture,
            PinWindowCall::SetFocusable { focusable: false },
            PinWindowCall::HidePinBubble,
        ],
        "and a pin with no claim asks the window the user left for nothing — not the \
         foreground, and not the focus — while still putting the style it took off back"
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
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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
    let _one = PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

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

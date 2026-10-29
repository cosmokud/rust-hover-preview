//! Carrying a pinned window with the pointer: taking the pointer for the drag, following a hand
//! whose messages arrive elsewhere, and letting go at either end of it.

use super::*;

/// Whether the left mouse button is down, as the Explorer hook's last read of the buttons
/// found it.
///
/// Published rather than read here, because `GetAsyncKeyState` cannot be asked for this
/// without also spending the press bit that tells the hook a click has happened — and the
/// hook is what `Pin Mode → Update Preview` follows a pick by. See `left_button_down` for
/// the whole of it, and `PIN_MEDIA_LEFT_PRESSES` for the press beside the state.
pub(super) static PIN_MEDIA_LEFT_DOWN: AtomicBool = AtomicBool::new(false);

/// How many left-button presses the Explorer hook has read, as a count rather than a bit.
///
/// The preview thread polls for a pin's press handling more slowly than the hook publishes,
/// so a press published as a bit would be overwritten by the next tick's "no press" before
/// this side ever looked and would read as no press at all. A count cannot be missed that
/// way: what this side asks is whether the count has moved, which is the question a press is.
pub(super) static PIN_MEDIA_LEFT_PRESSES: AtomicU64 = AtomicU64::new(0);

/// Whether a pinned window's pointer is carrying it or pulling it to a new size right now.
///
/// What a press, a drag and their ends do to the film is **nothing**: the film is a thumbnail of a
/// source window no gesture moves, so a plain click and a drag that carries the window are the
/// same answer to the player — it plays on untouched (see `begin_pin_drag`). This is read for the
/// window's own work and by a test that asks whether a hand is on it.
#[cfg(test)]
pub(super) fn pin_is_dragging() -> bool {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.dragging.is_some()))
        .unwrap_or(false)
}

/// Set the cursor for where the pointer is on a pinned window, answering whether this app set it.
///
/// A resize is the one gesture a window gives away with the shape of the pointer rather than with
/// anything drawn, and a pinned window has no other way of saying its edges are its own: an edge
/// that does not say so is an edge found by trying. A side names the side it is and a corner the
/// diagonal it lies on, which is where the edge is rather than which way the box will move — a
/// kind whose pixels are their own shape moves both sides whichever edge is taken (see `PinFrame`),
/// and the shape of the pointer is still about the hand.
///
/// The point is read from the cursor rather than taken from a message, because the message that
/// asks the question does not carry one: `WM_SETCURSOR`'s `lParam` is the hit-test code with the
/// mouse message above it, and read as a point it is a hand a pixel off the left of every window
/// and five hundred and twelve pixels down — which is the diagonal cursor drawn over the whole of
/// a pin, since a band is a band whatever the pointer is really on.
pub(super) unsafe fn pinned_set_cursor(hwnd: HWND) -> bool {
    let Some(point) = cursor_screen_point() else {
        return false;
    };
    let Some(origin) = window_origin(hwnd) else {
        return false;
    };

    let (frame, hovering, dragging) = {
        let Some(pinned) = pin_state() else {
            return false;
        };
        let Some(pin) = pinned.pin() else {
            return false;
        };
        if pin.collapsed {
            return false;
        }

        (
            pin.frame,
            pin.resize_edge(point.0 - origin.0, point.1 - origin.1),
            pin.dragging.map(|drag| drag.action),
        )
    };

    // What a drag that is going says about the pointer stands until it lets go, whatever the
    // pointer has since been dragged over: a hand that is carrying a window is not asking where
    // the edges of it are, and a hand that is pulling one should keep the shape it began with.
    // A pan carries the picture rather than the window and is given the same shape for the same
    // reason: what the hand is holding has moved, whatever it is holding it by.
    let edge = match dragging {
        Some(PinDragAction::Move) | Some(PinDragAction::Pan { .. }) => {
            if let Ok(cursor) = LoadCursorW(None, IDC_SIZEALL) {
                SetCursor(cursor);
            }
            return true;
        }
        Some(PinDragAction::Resize(edge)) => edge,
        None if frame != PinFrame::None => match hovering {
            Some(edge) => edge,
            None => return false,
        },
        None => return false,
    };

    let named = if edge.horizontal() && edge.vertical() {
        // The two diagonals: the one that runs down to the right, and the one that runs up to it.
        if edge.left == edge.top {
            IDC_SIZENWSE
        } else {
            IDC_SIZENESW
        }
    } else if edge.horizontal() {
        IDC_SIZEWE
    } else {
        IDC_SIZENS
    };

    if let Ok(cursor) = LoadCursorW(None, named) {
        SetCursor(cursor);
    }
    true
}

/// Answer a press that has landed on the window standing in a pin's media band.
///
/// A document the engine draws is not this app's pixels and not this app's window: the band the
/// pin leaves for it is filled by a browser's window, which is *over* the pin's — a press there
/// is delivered to the browser, and the pinned window never sees it at all. So a hand on the
/// drawing that means to carry the window, which is what a hand anywhere on the media of a pin
/// means (see `pinned_content_is_the_pins`), would begin no drag: the window would not move,
/// and the document the hand was on would not either. What is read here is the press itself —
/// the one reading the engine's window cannot be asked for, since it is not this window — and
/// what it begins is the drag the point means, by the same geometry the window procedure uses
/// for a press on the parts of the pin the engine does not cover.
///
/// It is a transition and not a state, and the latch is what makes it one: the press that is
/// answered sets it, a button that is not down clears it, and a hand held down on the band is
/// one drag rather than one per tick. The drag it begins takes the pointer for the pin (`SetCapture`
/// in `begin_pin_drag`), so the rest of it is the one `pinned_mouse_move` and `pinned_release` carry.
///
/// Three kinds of press are deliberately left where they landed: a document that *runs*, whose own
/// rectangle is where its own clicks belong (see `page_runs`); a press with no pin up or on one down
/// to its bubble, which is not on this window at all; and a press that did not reach the engine's
/// window, which has already come to some window's to answer.
///
/// The press is recognised by the count the Explorer hook publishes having moved rather than by
/// the button being down: a press and a release inside the gap between two ticks leaves the button
/// up on both of them, so a latch on the level alone never opens. The count cannot be missed that
/// way, and the button's own state is still what says the hand is on it now (see
/// `pin_media_press_count`).
pub(super) unsafe fn settle_pinned_engine_press(hwnd: HWND, seen_presses: &mut u64) {
    let presses = pin_media_press_count();

    if presses == *seen_presses {
        return;
    }
    *seen_presses = presses;

    // A press the hook has seen and this window has not answered, with the button already up
    // again, is a click rather than a hand on the drawing: there is nothing to drag, and the
    // click is the page's own.
    if !left_button_down() {
        return;
    }

    // Nothing of anybody else's stands in a bubble, and nothing of the engine's does either: what
    // a collapsed pin leaves is the round bubble and the listing under it.
    if pin_is_collapsed() {
        return;
    }

    // The document that is really on screen, read from the engine rather than from the pin: a pin
    // mid-swap still names the file it is showing, and it is the engine's window that took the
    // press either way.
    let Some(shown) = webview_preview::showing_path() else {
        return;
    };

    if webview_preview::page_runs(&shown) {
        return;
    }

    // What took the press: the engine's window, or one of the browser's own child windows inside
    // it — the same question the pointer is asked in `cursor_preview_hover`, asked here about a
    // press in the same spot.
    let Some(point) = cursor_screen_point() else {
        return;
    };
    if !engine_window_is_at(point.0, point.1) {
        return;
    }

    let Some((window, dpi, frame)) = pinned_window_frame() else {
        return;
    };

    // **A press on a pinned window is what makes the pin the window the user is in, whichever part
    // of it was pressed** — and what it begins is the drag its place means. The drag is begun first
    // for the reason the pin's own procedure begins it first: the pointer is what carries the hand,
    // and the window work the keyboard costs must not stand between the press and it (see
    // `pin_input::press_the_pin`).
    begin_pin_drag(
        hwnd,
        &Win32PinWindow,
        pinned_engine_press_action(window, dpi, frame, point),
        false,
    );
    pin_take_focus(hwnd, &Win32PinWindow);
}

/// Bring a pinned window to the front of its band for a press the user made on the preview,
/// whichever window of the preview's took it.
///
/// A press is read from the count the Explorer hook publishes rather than from this window's own
/// procedure, because the procedure is not always the one that gets it: a page an engine of this
/// app's draws stands over the pin's band in a window of its own, and a film can stand outside the
/// band it was given, so a press can be one Windows hands to a window this one never sees. The
/// pointer says which window that was — the pin's own, the engine's page, or the film's — and any of
/// them is the preview being asked for.
///
/// What the raise answers is the refusal that made a press look unanswered: the foreground is
/// refused to a process that has just been put behind, which is the first press after a click
/// anywhere else, while the keyboard arrives and the window stays where it was (see
/// `pin_take_focus`). This raise is not refused (see `put_pin_at_the_front`).
///
/// What a press *does* — a move, a resize, the card's own buttons — is not answered here. A press
/// this window's procedure never saw begins its drag from the road that owns it: the engine's page
/// has one of its own (see `settle_pinned_engine_press`).
pub(super) unsafe fn settle_pinned_press(hwnd: HWND, seen_presses: &mut u64) {
    let presses = pin_media_press_count();

    if presses == *seen_presses {
        return;
    }
    *seen_presses = presses;

    if !pinned() || pin_is_collapsed() {
        return;
    }

    let Some(point) = cursor_screen_point() else {
        return;
    };

    // The window the press went to, asked of the point rather than of the procedure that answered
    // it: a press on a film or on the engine's page is answered by a window this one never sees.
    use windows::Win32::UI::WindowsAndMessaging::WindowFromPoint;

    let under = WindowFromPoint(POINT {
        x: point.0,
        y: point.1,
    })
    .0 as isize;
    let film = VIDEO_HWND.load(Ordering::SeqCst);

    let pressed_the_preview = is_preview_window(under)
        || (film != 0 && film == under)
        || engine_window_is_at(point.0, point.1);

    if pressed_the_preview {
        put_pin_at_the_front(hwnd);
    }
}

/// Whether the window under a point is the one a document the engine draws is standing in.
///
/// What the pointer is over *inside* that window is a browser's own child window — one or two
/// levels down — so the question is whether the window under the point is the engine's or one
/// inside it, exactly as `cursor_preview_hover` asks it: comparing the two handles alone never
/// answers yes.
///
/// **A page is also the browser's window rather than the host's**, and a window of the browser's
/// that the host's tree does not answer for is still the page standing under the pointer: the
/// process that drew it is the one the host started, and nothing else on the desktop is. Asked of
/// the tree alone, a press on a document's picture was read as a press on somebody else's window,
/// and the pin was not raised by it (see `settle_pinned_press`).
pub(super) fn engine_window_is_at(x: i32, y: i32) -> bool {
    unsafe {
        use windows::Win32::UI::WindowsAndMessaging::{
            GetWindowThreadProcessId, IsChild, WindowFromPoint,
        };

        let under = WindowFromPoint(POINT { x, y });
        if under.is_invalid() {
            return false;
        }

        let engine = webview_preview::showing_hwnd();
        let mut pid = 0u32;
        let _ = GetWindowThreadProcessId(under, Some(&mut pid));

        the_engines_window(
            engine != 0 && under.0 as isize == engine,
            engine != 0 && IsChild(HWND(engine as *mut _), under).as_bool(),
            pid,
            webview_preview::browser_pid(),
        )
    }
}

/// The rule `engine_window_is_at` decides by, apart from the desktop it reads: a window is the
/// page the engine draws where it *is* the host window, where it is a child of it, or where the
/// process that owns it is the browser the host started — and never otherwise, an unknown
/// browser process included.
pub(super) fn the_engines_window(
    is_host: bool,
    is_child: bool,
    pid: u32,
    browser_pid: u32,
) -> bool {
    is_host || is_child || (browser_pid != 0 && pid == browser_pid)
}

/// Whether the left button is down, as the Explorer hook published it.
///
/// `GetAsyncKeyState`'s low-order bit — the one that says a key has been pressed since the
/// previous call — is *spent* by any call for that key, whichever thread makes it. Asking for it
/// here therefore took the click out from under the Explorer hook, which is what tells
/// `Pin Mode → Update Preview` that the user picked a file: the two loops poll at different rates
/// on different threads, so a click was answered by whichever read it first. The hook reads the
/// buttons once per tick and publishes the left button's state; this is that state.
pub(super) fn left_button_down() -> bool {
    PIN_MEDIA_LEFT_DOWN.load(Ordering::Acquire)
}

/// Publish the left mouse button's state for the pin's own press handling, from the one place
/// in this app that reads the buttons.
///
/// The hook calls it on every tick, pinned or not, so the published state is never older than
/// a tick and never absent. Ordering is release/acquire because the only thing crossing here
/// is the button's state and not one reading of another's memory (see `PIN_MEDIA_LEFT_DOWN`).
pub fn publish_pin_media_press(down: bool, pressed: bool) {
    PIN_MEDIA_LEFT_DOWN.store(down, Ordering::Release);

    if pressed {
        PIN_MEDIA_LEFT_PRESSES.fetch_add(1, Ordering::AcqRel);
    }
}

/// How many left presses the hook has seen, as a count rather than as a bit, because the
/// preview thread polls slower than the hook writes: a bit published and published again a
/// tick later reads as no press at all from here, and a press is a transition that has to be
/// noticed to have happened. The count is what makes it survivable, and it wraps at a rate no
/// session reaches (see `PIN_MEDIA_LEFT_PRESSES`).
pub(super) fn pin_media_press_count() -> u64 {
    PIN_MEDIA_LEFT_PRESSES.load(Ordering::Acquire)
}

/// The box, display scale and frame of the pin that is up, for a press that landed on the window
/// standing in its media band — read in one look, so that the lock is not held across the answer.
pub(super) fn pinned_window_frame() -> Option<(ScreenRegion, u32, PinFrame)> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    Some((pin.window_box(), pin.dpi, pin.frame))
}

/// What a press that has landed on the window standing in a pin's media band begins: a resize
/// where the point is on an edge of the pin's own window, and a move anywhere else.
///
/// The two are asked in the order the window procedure asks them of a press that came to it — an
/// edge first, and the media after it (`pinned_press`) — and of the same geometry, because a
/// press on the pin is a press on the pin whichever window took it: the bands an edge is found
/// in are the ones a hand on the opaque parts of the pin finds them in (`pin_resize_edge`), and
/// everything left of them is the media, which is the handle a window's body is.
///
/// The point arrives in screen coordinates, since it is read from the pointer rather than from a
/// message of this window's, and the pin's own window is where its coordinates begin.
pub(super) fn pinned_engine_press_action(
    window: ScreenRegion,
    dpi: u32,
    frame: PinFrame,
    point: (i32, i32),
) -> PinDragAction {
    if frame != PinFrame::None {
        let size = ((window.2 - window.0).max(1), (window.3 - window.1).max(1));
        if let Some(edge) = pin_resize_edge(size, dpi, point.0 - window.0, point.1 - window.1) {
            return PinDragAction::Resize(edge);
        }
    }

    PinDragAction::Move
}

/// What a press becomes: a drag of the window, and the pointer taken for it.
///
/// *When a press becomes a carried drag* is the whole of this function, and it is the decision
/// that had no tests: the pointer is taken only once there is a drag for it, and the two were
/// not connected. A lock this thread could not take, or a pin that had been taken down since
/// the press was read, left the window holding the pointer for the whole desktop with nothing
/// that would ever release it. Every mouse message then went to this window rather than to
/// whatever the pointer was aimed at, and the cursor kept whichever shape the last edge gave
/// it — a desktop that looked broken until some other window took the pointer for itself,
/// which is why clicking elsewhere appeared to bring it back.
///
/// It is behind `PinWindow` because all three of the things it asks of the machine are: where
/// the pointer is, where the window stands, and whether the pointer may be taken. Only the
/// middle step — *whether there is a pin for the drag to live in* — is this file's, and it is
/// the step that decides which of the other two is asked at all.
pub(super) fn begin_pin_drag(
    hwnd: HWND,
    window: &dyn PinWindow,
    action: PinDragAction,
    delivered: bool,
) {
    let hwnd = hwnd.0 as isize;
    let Some(from) = window.pointer() else {
        return;
    };
    // The box the window is standing at on screen, rather than the one the pin remembers: a
    // window the hand has carried since it was maximized is no longer standing where the
    // maximize left it, and a resize begun from the remembered box is begun from the screen's
    // own top border rather than from the place the hand left the window at — which is the
    // window snapping back to the top the moment an edge is pulled (see `apply_pin_drag`).
    let Some(window_box) = window.window_box(hwnd) else {
        return;
    };

    let installed = pin_state()
        .and_then(|mut pinned| {
            let pin = pinned.pin_mut()?;
            pin.dragging = Some(PinDrag {
                from,
                window: window_box,
                action,
                delivered,
                // A drag has not been carried out to anywhere yet, and `from` is a place it has
                // not been: the window has not moved by a single pixel until the pointer does.
                carried: (i32::MIN, i32::MIN),
            });
            Some(pin.dragging)
        })
        .is_some();

    if installed {
        window.capture(hwnd);
        // **Nothing of the film is asked for here, and there is nothing of it to ask.** The film
        // is a thumbnail of a source window no gesture moves, so a press, a move, a resize and
        // their ends are this app's own window work and one rectangle: no key, no kill, no
        // relaunch — the picture plays on through the gesture (see `pin_film_thumbnail`).
        //
        // A press that never moves is a click, and a click is a press and a release and nothing
        // else: there is nothing here for the pointer's travel to decide either way.
        // **The band is asserted once per drag rather than on every step of it.** Which band a pin
        // belongs in is a question about the setting (`band_place`), and the box a drag is applying is
        // not a fact it is asked about — so re-asking it on every pointer move was a DWM reorder per
        // move for a window already where it names (see `apply_pin_drag` and `finish_pin_drag`).
        //
        // SAFETY: the handle is the window the drag is carried in — the pin's own, or the one it
        // stands in — and `SetWindowPos` takes any window and reports rather than faults.
        unsafe {
            pin_film_host::place_the_pin_in_the_band(HWND(hwnd as *mut _), band_place());
        }
        // **The source is not asked for anything here, or anywhere else in the gesture.** Its box was
        // fixed at adoption, at the film's own size, and the Shell stretches it into whatever
        // destination the pin's host reports — so a drag of either kind is this app's own window and
        // one rectangle (see `pin_film_thumbnail`).
    } else {
        window.release_capture(hwnd);
    }
}

/// Let go of the pointer a drag took.
///
/// It used to mark the notice the release raises as one this app asked for, and the marking is
/// gone with the road that read it: `WM_CAPTURECHANGED` arrives inside this call, and the window
/// procedure answering one used to have to tell this app's own release from a capture somebody else
/// took — because a release was a moment the road letting go of the pointer was already handling,
/// and a stolen capture was not. Nothing of a player is asked of either end now, so the two settle
/// identically and there is nothing left for the marking to keep apart. What the call still has to
/// be is the one place the pointer
/// is let go of: `ReleaseCapture` answers a window procedure on this same thread before it returns
/// (see `release_pin_capture`), so a drag's record is dropped from inside this call either way.
pub(super) fn release_the_pointer(window: &dyn PinWindow, hwnd: isize) {
    window.release_capture(hwnd);
}

/// Show the window FFmpeg is drawing the picture into, or take a condemned one off the screen.
///
/// A window of this app's player is a window of another process, and the only show this app asks of
/// it is the one the style monitor makes to put it up without activating it (see
/// `ensure_video_window_topmost`). The hide is the supersede road's alone: a replacement the loop
/// has condemned is un-shown rather than placed (see `revalidate_shown_replacement`).
pub(super) fn set_pinned_player_window_visible(visible: bool) {
    // **A film drawn into the pin's host has no window of the desktop's to hide.** The host is the
    // picture's window and the band is the hole it is seen through, so there is nothing of the
    // player's to take off the screen: a host hidden here is the standing registration blanked,
    // which is the one thing no road may do to a pin (see `pin_film_host`). The question is asked
    // of the pairing, because a pin that took up a hover's top-level player has a host standing
    // empty until a relaunch fills it.
    if pin_film_host::the_pinned_player_was_begun_in_the_host() {
        // SAFETY: the host is this app's own window, and `ShowWindow` takes any window and reports
        // rather than faults. A hidden host is a hidden film, so only a show is asked of it.
        if visible {
            unsafe { pin_film_host::set_the_pin_film_host_visible(true) };
        }
        return;
    }

    let Some(hwnd) = video_window_for(VIDEO_PID.load(Ordering::SeqCst)) else {
        return;
    };

    // SAFETY: `hwnd` is a window this app found by enumerating the desktop and matching it to the
    // player it started, so it is a live handle to this process's own window; `ShowWindow` takes
    // any window and reports rather than faults.
    unsafe {
        let _ = ShowWindow(hwnd, if visible { SW_SHOWNOACTIVATE } else { SW_HIDE });
    }
}

pub(super) fn hide_pinned_player_window() {
    set_pinned_player_window_visible(false);
}

/// Carry on, and let go of, a drag whose press this window was never given.
///
/// A drag begun from the hook's published press count cannot be ended by a message, and there is
/// nothing else to end it with: the engine's window is what a press over a drawing is delivered
/// to, and nothing promises the release is ever delivered here.
///
/// So the end is read from the same place the beginning was, and the button's published state is
/// a level rather than a transition: it cannot be missed the way a message can. The drag is carried
/// on while the button is down — from the pointer's position rather than from a move message, and
/// because `GetCursorPos` is not subject to whatever is holding the pointer.
///
/// A drag this window *was* given its release for is left alone: it ends in `pinned_release`, and
/// its capture is this window's own to hold until that message arrives.
pub(super) unsafe fn settle_pinned_engine_drag(hwnd: HWND) {
    let Some(drag) = pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin.dragging))
    else {
        return;
    };

    if drag.delivered {
        return;
    }

    if left_button_down() {
        apply_pin_drag(hwnd);
        return;
    }

    // No final apply on the way out: a dragged pin's release is delivered to the pin's own window,
    // where `pinned_release` applies the pointer's last place (see `finish_pin_drag`).
    finish_pin_drag(hwnd, &Win32PinWindow);
}

/// Whether a drag is being carried on from what the hook publishes rather than from a message: one
/// begun off the hook's press count, whose end is the same reading and not a `WM_LBUTTONUP` that
/// is not promised to arrive (see `settle_pinned_engine_drag`).
pub(super) fn pin_drag_is_carried() -> bool {
    carried_drag().is_some_and(|drag| !drag.delivered)
}

/// The place a carried drag has last put its window, which is the pointer's position when it did
/// and is how a pass is told from one that has moved anything.
pub(super) fn pin_drag_carried_to() -> (i32, i32) {
    carried_drag()
        .map(|drag| drag.carried)
        .unwrap_or((i32::MIN, i32::MIN))
}

/// The drag of a pinned drawing the loop is carrying, read in one look: what a pass needs both
/// the liveness and the last place from, so the two are not read under the lock separately.
pub(super) fn carried_drag() -> Option<PinDrag> {
    pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin.dragging))
}

/// Follow a drag of a pinned drawing for as long as the hand is going.
///
/// This is the whole of what a drag of a drawing is missing. Every other kind of pinned window has
/// its drag carried by `WM_MOUSEMOVE` messages delivered to it, and the wait at the end of the loop
/// wakes the moment one is queued, so it follows the hand at the pointer's own rate. A drawing's
/// press is taken by the engine's window, on a thread of its own, and the moves went down with it
/// to the browser's own child window — so this window is sent nothing at all, and the loop is
/// left carrying the drag (see `settle_pinned_engine_drag`). A loop carries it at the loop's pace,
/// which is one place per tick, and a tick of a pinned static document is `STATIC_PIN_WAIT_MS` away:
/// eleven places a second is what that costs, which a hand on a 144 Hz display reads as a window
/// being thrown after the pointer rather than carried by it.
///
/// So while a hand is going, the loop does not wait at all. It puts the window where the pointer
/// is, gives the window procedure its turn, and asks again. A pass costs a `GetCursorPos` and a
/// comparison when the pointer has not shifted, so following costs what following costs.
///
/// Two things bound it, both about not being a thread that never gives the tick back: a hand that
/// stops moving is handed back to the loop's own wait, but only once it has sat still for a while,
/// and every pass notes the loop alive (see `note_pin_alive`).
pub(super) unsafe fn carry_pin_drag_with_the_hand(
    hwnd: HWND,
    rx: &Receiver<PreviewMessage>,
) -> CarryOut {
    let mut carry = Carry::begin(PIN_DRAG_WAIT_MS);
    let mut moved_at = Instant::now();

    while pin_drag_is_carried() {
        if carry.is_exhausted() {
            break;
        }
        note_pin_alive();

        let carried_to = pin_drag_carried_to();
        settle_pinned_engine_drag(hwnd);
        if pin_drag_carried_to() != carried_to {
            moved_at = Instant::now();
        } else if moved_at.elapsed() >= PIN_DRAG_HAND_RESTED {
            break;
        }

        // The window's own queue is drained here so that a window being carried goes on
        // answering everything a window answers — and the drain is refused if this thread is
        // already inside one, because a window procedure that pumps the same queue re-enters
        // it over those messages, and a nested drain is how a caption button ends up answered
        // twice for one click.
        if !carry.pump_window_messages() {
            break;
        }

        carry.drain(rx);
    }

    carry.finish()
}

/// What a carry found on the loop's channel, for the tick to take up before anything else.
pub(super) type CarryOut = Vec<PreviewMessage>;

/// One window's drag being carried by the loop instead of by its own messages.
///
/// The two loops that carry a pointer — this one and the bubble's — were two hand-written
/// loops with two different bounds and neither of them bounded in the way that matters. This
/// one followed a hand that never stopped moving, for as long as the hand did, giving the tick
/// back only when the pointer had been still for `PIN_DRAG_HAND_RESTED`; and while it was
/// inside, the loop's own channel was never polled at all, so a `Close` on the caption, a
/// media kind switched off in the tray, or an engine that had died all went unacted-on for the
/// length of the drag. A caption that stops answering during a drag is the livelock this
/// bounds, and the bound is what makes the loop a guest in its own tick rather than a
/// replacement for it.
///
/// Three things are held to, and each is a way this used to be a thread that never gave the
/// tick back:
///
/// - A deadline, so a hand that moves for minutes is a window that follows for minutes rather
///   than a loop that is gone. Past it the drag is still carried — by the tick, at the tick's
///   pace — so nothing is lost but the smoothness.
/// - A refusal to pump the window queue from inside a pump, so a re-entrant dispatch cannot
///   drain the same queue twice.
/// - A drain of the preview channel on the way through, so a message sent while the hand is
///   moving is held rather than lost, and is taken up by the tick that resumes.
pub(super) struct Carry {
    /// When this carry is over, whatever the hand is doing. Past it the drag is carried by the
    /// tick instead, which is slower and answers everything.
    pub(super) deadline: Instant,
    /// Whether a pump is already in progress on this thread. Set by `pump_window_messages` for
    /// the duration of the drain, so a window procedure that pumps re-enters without a second
    /// drain over the same queue.
    pub(super) pumping: bool,
    /// Messages the drain found, in the order they arrived, for the tick to take up.
    pub(super) carried: Vec<PreviewMessage>,
}

impl Carry {
    /// Begin carrying, bounded by `wait_ms`.
    pub(super) fn begin(wait_ms: u64) -> Self {
        Self {
            deadline: Instant::now() + Duration::from_millis(wait_ms),
            pumping: false,
            carried: Vec::new(),
        }
    }

    /// Whether this carry is over and the tick should be given the drag back.
    pub(super) fn is_exhausted(&self) -> bool {
        Instant::now() >= self.deadline
    }

    /// Drain the window queue, answering what is on it where it is.
    ///
    /// `false` means the queue is already being drained on this thread, and the carry is over:
    /// a nested drain would dispatch the same message twice, and a caption button answered
    /// twice for one click is a click that lands in the wrong place.
    pub(super) fn pump_window_messages(&mut self) -> bool {
        if self.pumping {
            return false;
        }
        self.pumping = true;
        pump_window_messages();
        self.pumping = false;
        true
    }

    /// Take off whatever the loop has been sent, holding it for the tick.
    ///
    /// Non-blocking, and every pass: this is a look on the way past rather than a wait, since
    /// the wait is what the carry exists to avoid. The messages go to `carried` rather than to
    /// the loop's own drain, because the tick is the thing that acts on them.
    pub(super) fn drain(&mut self, rx: &Receiver<PreviewMessage>) {
        while let Ok(message) = rx.try_recv() {
            self.carried.push(message);
        }
    }

    /// Hand back what was taken off the channel, in the order it arrived.
    ///
    /// A tick that resumes with these in hand acts on them before it does anything else, so a
    /// pin asked to close while it was being dragged closes when the drag ends rather than
    /// after the next hover.
    pub(super) fn finish(self) -> CarryOut {
        self.carried
    }
}

/// Let go of a drag, from whichever of the two ends has arrived.
///
/// The tail of `pinned_release` and the whole of what a drag read out of the hook's published
/// button state owes the same three things, so they are one function: the pointer this window
/// took for the drag is let go of (it is taken outside the lock, and marked as this app's own
/// release, because `ReleaseCapture` delivers `WM_CAPTURECHANGED` into this same thread's window
/// procedure before it returns — see `release_the_pointer`), a resize asks for its media to be
/// laid out again at the box it ended up with, and the window is painted at where it stands. Which
/// of the two ends is *asked* is the caller's: a message ends a drag this window was given, and
/// the tick ends a drag begun out of what the hook published (see `settle_pinned_engine_drag`).
///
/// Returns whether there was a drag to let go of, which is what a caller that has other release
/// work to do needs to know.
///
/// The window work is the one call `begin_pin_drag` conditions, read back off the seam rather
/// than off `GetCapture` — the two ends of a capture are one list, and a list that had to be
/// read out of the machine in two places is a list whose two halves can disagree.
pub(super) fn finish_pin_drag(hwnd: HWND, window: &dyn PinWindow) -> bool {
    let drag =
        pin_state().and_then(|mut pinned| pinned.pin_mut().and_then(|pin| pin.dragging.take()));
    let Some(drag) = drag else {
        // A drag the pin was rebuilt over is a drag that is not coming back, and the capture its
        // press took is not something this window may keep: a pinned window holding the pointer
        // eats every mouse message on the desktop and keeps whichever shape the last edge gave the
        // cursor, for the rest of the process (see `release_the_pointer`, and
        // `reconcile_swap_take_up` for the rebuild that takes the drag away).
        release_the_pointer(window, hwnd.0 as isize);
        return false;
    };

    release_the_pointer(window, hwnd.0 as isize);

    // **A resize asks for its media to be laid out again at the box it settled on**, and for a
    // film that is the host fitted to the new box and nothing else — no player is ended for it
    // (see `relayout_pinned_media`).
    if matches!(drag.action, PinDragAction::Resize(_)) {
        if let Some(content) = pinned_content() {
            if let Ok(mut request) = PIN_BOX_REQUEST.lock() {
                *request = Some(content);
            }
        }
    }

    // The held frames the drag's steps were counted in are summed here, at the settle that ends it
    // (see `pin_film_thumbnail::settle_the_held_frames`).
    pin_film_thumbnail::settle_the_held_frames();

    // **Nothing of the player is shown from here, by either kind of drag.** The film is a
    // thumbnail of a source window the gesture never moved, so it was never hidden and there is
    // nothing to hand back: a show from here used to be the second half of the flash a drag's end
    // had, where the band went transparent over a window whose first composited frame had not
    // landed yet.
    window.repaint();
    // The drag's other end asks for the band its beginning asked for, which is the whole of what a
    // drag re-asserts: a step of it does not ask at all (see `begin_pin_drag` and `apply_pin_drag`).
    // SAFETY: `hwnd` is the window the drag was carried in, and `SetWindowPos` takes any window and
    // reports rather than faults.
    unsafe {
        pin_film_host::place_the_pin_in_the_band(hwnd, band_place());
    }
    true
}

/// Let go of everything a pinned window was doing with the pointer, now that the pointer is a
/// window's other than this one's to be doing it with.
///
/// Windows delivers this message to the window that *lost* the capture rather than a release to the
/// one that took it, and it is the only notice a gesture gets when it did not end at this window's
/// own release: a click into another window, a dialog of somebody else's, or the watchdog taking a
/// pin down mid-drag all end here rather than on a `WM_LBUTTONUP` (see `release_pin_capture`). So
/// the drag's record goes, and with it the gesture the tick is watching for.
///
/// **A capture ends whatever gesture was running — it does not create one, and it does not ask a
/// player for anything.** A relaunch the user asked for is the relaunch road's, and the gesture's
/// own end settles nothing about it. What there is no `hwnd` for is a capture this app has already
/// lost:
/// Windows has taken it, and releasing it would be a courtesy owed to nobody.
///
/// Windows delivers `WM_CAPTURECHANGED` to the window that lost the capture whether it released it
/// or another window took it, and it says in no way the two can be told apart afterwards. Both
/// leave the same thing behind, which is a drag that has ended.
pub(super) fn pin_capture_lost(window: &dyn PinWindow) {
    // A capture taken from under a drag is the drag's other settle, and the held frames it counted are
    // summed here as they are at a release (see `pin_film_thumbnail::settle_the_held_frames`).
    pin_film_thumbnail::settle_the_held_frames();

    // The seek's aim is taken with the capture: a scrub abandoned here never relaunches, so
    // nothing is left standing to aim a player at.
    if let Some(mut pinned) = pin_state() {
        if let Some(pin) = pinned.pin_mut() {
            pin.dragging = None;
            pin.pressed = None;
            pin.transport.pressed = None;
            pin.transport.seeking = None;
            pin.transport.hovered = None;
            // A knob that was being held is let go of with the capture: a pointer that has gone
            // elsewhere is not a hand still on the level, and the popup is put away by the tick that
            // finds the pointer away from it (see `refresh_pin_volume`).
            pin.volume.dragging = false;
        }
    }

    // And the card's menu, which the pointer is held for exactly the way a drag is: a capture
    // somebody else took is a menu nothing can be answered in, and the panels are no longer
    // inside the card's own window — so unlike a drag, a menu left standing by one is not a
    // menu a press on the card can still put away (see `close_pin_menu` and
    // `pin_menu_window`).
    unsafe { close_pin_menu(HWND(window.hwnd() as *mut _)) };
}

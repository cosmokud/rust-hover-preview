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
/// A press that has not moved is not yet a drag, which is the whole of what this asks over the
/// drag itself: `PinnedPreview::dragging` is written when the press becomes a drag and not when the
/// pointer goes down, so a click on the caption is not a drag and a click does not stop the film.
pub(super) fn pin_is_dragging() -> bool {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.dragging.is_some()))
        .unwrap_or(false)
}

/// Whether the player standing behind a pinned window's media band has been put away for the length
/// of a drag.
///
/// **It is asked by every place that would put that window up again, and it is the whole of what
/// makes the park hold.** A drag hides the player's window on its own first pointer message and
/// paints the band flat over the hole that leaves (see `park_pinned_player`), and a hide something
/// else undoes is not a park at all: both `SetWindowPos(SWP_SHOWWINDOW)` and
/// `ShowWindow(SW_SHOWNOACTIVATE)` are shows, so every place that puts the window where the band is
/// has been putting it back on screen as well.
///
/// Three of those places run many times a second, and they are the three a drag spends its life
/// inside: a resize places the player on every pointer move (see `apply_pin_drag`), the tick
/// re-asserts it every couple of hundred milliseconds, and the style monitor's own raise goes every
/// hundred milliseconds whatever anybody asked for (see `apply_noactivate_to_hwnd`). The tick is what
/// holds a hand that has stopped moving — a hand resting on an edge is a drag like any other and the
/// longest one a user spends — and the monitor is the one no caller can hold off, so the film was
/// back on screen a frame after it was hidden, at the pointer's own pace: the exact cost the park
/// exists to pay, with the stutter it was written for coming back along with the picture.
///
/// **The band cannot cover for the park, and the flat paint is why.** The band is transparent
/// everywhere else so that the window underneath shows through it, and the player is topmost and
/// nearer the front of the band than the pin itself — so a re-shown player is a window over an
/// opaque rectangle, which is worse on screen than the hole the paint was hiding, not better (see
/// `render_pinned_preview_at`).
///
/// So the window is left exactly where the park left it: hidden, and at the place the drag had put
/// it before the park, which costs nothing to be out of date because the tick places it the moment
/// the park is undone (see `unpark_pinned_player`).
pub(super) fn pin_player_is_parked() -> bool {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.parked))
        .unwrap_or(false)
}

/// Whether the player on screen is being held for the length of a gesture.
///
/// It is one flag rather than a question asked of the player, because the player reports nothing:
/// this is the same arrangement `PinTransport` keeps for its own hold (see `transport_playing`), and
/// it is read here so that a gesture which is begun and ended between two ticks costs nothing and
/// one that is in flight is settled on the tick that finds it.
pub(super) static VIDEO_DRAG_HOLDING: AtomicBool = AtomicBool::new(false);

/// Whether the player on screen is being held for the length of a gesture.
pub(super) fn video_drag_holding() -> bool {
    VIDEO_DRAG_HOLDING.load(Ordering::Acquire)
}

/// Remember which way the hold is, answering whether it changed — so that the key is only posted
/// on the tick a gesture began or ended and not on every tick in between.
pub(super) fn video_drag_hold_set(dragging: bool) -> bool {
    VIDEO_DRAG_HOLDING.swap(dragging, Ordering::AcqRel) != dragging
}

/// What a drag of the window does to the film in it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum VideoDragHold {
    /// Nothing at all — the gesture is over a film this app did not stop, or is not over one.
    Leave,
    /// Hold the player where it stands.
    Hold,
    /// Let the player go on from where it stands.
    Release,
}

/// What a drag does to a film that is `playing`, on a hold that is `ours`.
///
/// The three facts are the whole question and two of them are refusals. **The key this posts is a
/// toggle**, which is what makes the gesture the wrong thing to consult: a toggle on a film that is
/// already held *resumes* it, so a drag begun over a film the pause button had stopped and then let
/// go of is a film the user started by moving the window they were watching it in. And a hold this
/// app did not put there is not ours to take back at the end of a gesture, or the same thing happens
/// a second time round.
///
/// So a film that is not playing is left alone, and a hold that is not ours is left alone, and only
/// a film that is playing and a hold that is ours are answered. The two refusals are not the same
/// refusal: the first is about there being nothing to hold and the second about not having been the
/// one to hold it, and the second is the one a test on the film alone would miss.
///
/// **The third fact is read out of the film on screen rather than kept beside it**, which is the
/// whole of what makes the second refusal reliable: `ours` is the transport's own record of a hold
/// this app put there (see `PinTransport::drag_held`), so it is reconciled by every path that ends
/// the player it was made against — a pin taken down, a file swapped, a player that died — and a
/// gesture that has ended cannot leave one standing for the next film to be dragged.
pub(super) fn video_drag_hold_decision(dragging: bool, playing: bool, ours: bool) -> VideoDragHold {
    match (dragging, ours) {
        (true, false) if playing => VideoDragHold::Hold,
        (false, true) => VideoDragHold::Release,
        _ => VideoDragHold::Leave,
    }
}

/// The claim over the film on screen, after a tick that has found a gesture in flight or has not.
///
/// **A claim ends with the gesture that made it, and with nothing else**, which is what makes it
/// safe to ask this before anything at all is known about the film on screen: a claim that outlives
/// its gesture is a claim the *next* film to be dragged is answered with, and that film was never
/// held by anybody — so the drag is refused the hold it exists for, and its release then posts the
/// toggle onto a film nobody stopped.
///
/// It is asked of the gesture and the claim alone rather than of either film, because the two
/// refusals further down are about what a drag does to a player and a claim is not about a player:
/// one of them is a kind of pin whose media this app does not hold and the other is no pin at all,
/// and behind either of them a claim used to be left standing for the rest of the run (see
/// `video_drag_hold_apply`).
pub(super) fn video_drag_hold_claim(dragging: bool, ours: bool) -> bool {
    dragging && ours
}

/// Begin or end the hold, answering whether a player was told.
///
/// It is the pause key, and it is the same key the transport bar's own hold is (see
/// `ffplay_key_pause`), posted to the same window: a player asked to hold does hold, and a player
/// that is holding is still there to be let go of — which is what makes the end of the gesture a
/// key rather than a relaunch. A relaunch would put the film back at the second the gesture began
/// at, so a window dragged for a second and released would lose the second it was watching.
///
/// **It is asked of the film rather than of the gesture, because the key is a toggle.** That is the
/// fault this used to have: the gesture was the only thing consulted, so a drag begun over a film
/// the pause button had already held posted a pause key onto a held film and *resumed* it — a film
/// started by the user moving the window they were watching it in. So both ends ask whether there
/// is anything playing to hold and whether this app is the one holding it, and a hold that is not
/// ours is left exactly as it was found.
///
/// **The hold is also written down**, or rather it is written through the transport, because a hold
/// this app has only asserted in a static flag is a hold nothing else can see: `pin_is_playing`
/// would still say the film is playing, the loop would still count its clock up, and the bar would
/// still be drawn at a playhead racing away from a frozen picture. It goes in through the same
/// `held`/`released` a press of the pause button uses, so the second written is the second the
/// picture is at and the clock under it is rebased when the hand lets go (see `PinTransport::held`
/// and `PinTransport::released`).
///
/// A hold that is begun and never ended is the one fault worth guarding: the picture would stay
/// frozen with the window moving under it for ever. The hold is ended again by every path that ends
/// a drag — the release, a capture lost to another window, the pin going down (which kills the
/// player outright, so there is nothing left to be holding) — and the tick is what notices, so the
/// only way to be stuck is a tick that stops running, which takes the whole window with it.
///
/// **And the claim over the hold is taken back before any of that is asked.** Two things were true
/// of it where it used to be written, and together they were the fault. It was a flag beside the
/// pin, so it outlived the film it was made against — a pin taken down, a file swapped and a player
/// that died all left it standing — and the release arm answered out of it only after two refusals
/// that both return early, so a gesture that ended over either of them left it standing as well. The
/// next film to be dragged was then answered by a claim belonging to a film that was not there: the
/// drag refused the hold it exists for, and its release posted the toggle onto a film nobody had
/// stopped — a frozen picture with a pause glyph over it and a loop still counting a clock.
///
/// So the claim is the transport's own field, which every path that begins, holds, lets go of or
/// loses a player reconciles (see `PinTransport::drag_held`), and the take-back runs ahead of the
/// refusals rather than behind them. A key that could not be posted is a player with no window, and
/// a player with no window is holding nothing — so the claim goes whether or not the key lands,
/// because leaving it standing would refuse every *later* gesture the film for the rest of the run.
pub(super) fn video_drag_hold_apply(dragging: bool) -> bool {
    // The film in front of the window — in flight, or the one the last gesture was over — read
    // before anything is refused below, because the claim has to be reconciled whether or not there
    // is still a film of its own to reconcile it in. A pin that is not up has taken the claim with
    // it, since there is nothing left to reconcile, and a pin that is up over something else is
    // carrying a claim about that something else rather than about this.
    //
    // It is read from a copy rather than from the pin: a lock held across the keys below would be a
    // lock held while this thread posts to another program's window.
    let Some((_, _, transport, _)) = pinned_playback_state() else {
        return false;
    };

    // Reconciled before anything is asked about the film on screen, and written only when it has
    // actually moved: see `video_drag_hold_claim`.
    let reconciled = video_drag_hold_claim(dragging, transport.drag_held);
    if reconciled != transport.drag_held {
        update_pin_transport(|state| state.drag_held = reconciled);
    }

    // A drag on a pin of a kind whose media this app draws is not a drag on a player at all: the
    // engine's window is not FFmpeg's and there is nothing to hold. The claim is already back where
    // it belongs by this point, which is the whole of what it was reconciled for.
    if current_media_type() != Some(MediaType::Video) {
        return false;
    }

    // What the decision is given is the claim as it stood when the tick began rather than the one
    // reconciled above, because that is what the gesture in flight did: a claim made while a gesture
    // was in flight is the one that gesture is still to let go of.
    match video_drag_hold_decision(dragging, pin_is_playing(&transport), transport.drag_held) {
        VideoDragHold::Leave => false,

        VideoDragHold::Hold => {
            if !ffplay_key_pause() {
                return false;
            }

            update_pin_transport(|state| {
                state.held(pin_playhead(&transport).unwrap_or(0.0));
                // The claim is written with the hold rather than beside it, because this is the
                // only place one is ever made and it is the records of this hold that take it back.
                state.drag_held = true;
            });
            true
        }

        VideoDragHold::Release => {
            if !ffplay_key_pause() {
                return false;
            }

            update_pin_transport(|state| state.released(pin_playhead(&transport).unwrap_or(0.0)));
            true
        }
    }
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
    let edge = match dragging {
        Some(PinDragAction::Move) => {
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

    // A press on a pinned window is what makes the pin the window the user is in, whichever part
    // of it was pressed, and what it begins is the drag its place means.
    pin_take_focus(hwnd);
    begin_pin_drag(
        hwnd,
        &Win32PinWindow,
        pinned_engine_press_action(window, dpi, frame, point),
        false,
    );
}

/// Whether the window under a point is the one a document the engine draws is standing in.
///
/// What the pointer is over *inside* that window is a browser's own child window — one or two
/// levels down — so the question is whether the window under the point is the engine's or one
/// inside it, exactly as `cursor_preview_hover` asks it: comparing the two handles alone never
/// answers yes.
pub(super) fn engine_window_is_at(x: i32, y: i32) -> bool {
    let engine = webview_preview::showing_hwnd();
    if engine == 0 {
        return false;
    }

    unsafe {
        use windows::Win32::UI::WindowsAndMessaging::{IsChild, WindowFromPoint};

        let under = WindowFromPoint(POINT { x, y });
        !under.is_invalid()
            && (under.0 as isize == engine || IsChild(HWND(engine as *mut _), under).as_bool())
    }
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
        park_pinned_player();
    } else {
        window.release_capture(hwnd);
    }
}

/// Put the player's window back where it was parked from, at where the band is *now*.
///
/// **The place is part of putting it back, and it is not tidiness.** A park is a hide and nothing
/// else: the window is left standing exactly where the drag last put it, and every place that
/// would move it while the park stands is answered out of the hand (see `pin_player_is_parked`).
/// So on a move — the one drag with no relaunch to bring a window of its own back — the rect the
/// player's window is still at is the rect the drag *began* from, and a bare `ShowWindow` puts
/// the film on screen there: a picture at the old box, behind a band that has moved, for as long
/// as the tick takes to re-assert the real one. Two hundred milliseconds of a film in the wrong
/// place is what the hand sees at the moment it lets go, so the place is asked for here rather
/// than waited for — and it is asked for through `ensure_pinned_sibling_box`, which is the same
/// call the tick makes, so a park that ends and a tick that re-asserts cannot disagree about where
/// the picture belongs. The band is read now, out of the pin, rather than remembered from before
/// the drag; a rect read at the wrong moment is the whole of the defect.
///
/// The film is *not* let go of here: that is the tick's half of the pair (see
/// `video_drag_hold_apply`), and the window is shown before that tick rather than after it, which
/// is the order worth having — a player unpaused while its window is hidden decodes into nothing,
/// so the first frame the hand sees is whichever one it decodes after being shown, a stutter on the
/// single frame the drag ended on.
///
/// It is asked of the pin and not of the player's window, because a pin taken down mid-drag has
/// ended the player outright and there is no window of its own left to find (see `park_pinned_player`).
pub(super) fn unpark_pinned_player(window: &dyn PinWindow) -> bool {
    let unparked = pin_state().is_some_and(|mut pinned| {
        pinned
            .pin_mut()
            .is_some_and(|pin| std::mem::take(&mut pin.parked))
    });

    if !unparked {
        return false;
    }

    // The band is read after the flag rather than with it, and both are read before the window is
    // asked for anything — the flag says the park is over and the band says where the window is
    // to go, and a pin that has been taken down between the two has no band to go to.
    window.unpark_player_window(pinned_content());
    true
}

/// Whether the capture being lost is one this app asked for, which Windows itself will not say.
///
/// `WM_CAPTURECHANGED` is delivered to the window that lost the capture whether it released it or
/// another window took it, and by the time a window procedure is looking at it, `GetCapture()`
/// gives the same answer either way: not this window. Only this app's own record of what it asked
/// for can tell the two ends apart, and a drag's end needs them apart — its own release is a moment
/// the road that let go of the pointer is already handling, and a capture stolen is a drag that is
/// not coming back (see `pin_capture_lost`).
///
/// A static rather than a field of anything, because the writer and the reader are not the same
/// function: the release is asked of a window and the notice is answered inside the call it raises.
pub(super) static PIN_CAPTURE_OURS: AtomicBool = AtomicBool::new(false);

/// Let go of the pointer, marking the notice the release raises as one this app asked for.
///
/// The marking is around the call and not after it, because the notice arrives *inside* it:
/// `ReleaseCapture` sends `WM_CAPTURECHANGED` back into this same thread's own window procedure
/// before it returns (see `release_pin_capture`), so the window procedure answering one has to be
/// able to see that this app is the one who let go. Unmarked, the two ends of a drag are one end,
/// and the notice this app raised itself is answered by a road with no relaunch decision in hand —
/// which puts a player that is about to be taken down back on screen at the box its drag began at.
pub(super) fn release_the_pointer(window: &dyn PinWindow, hwnd: isize) {
    PIN_CAPTURE_OURS.store(true, Ordering::Release);
    window.release_capture(hwnd);
    PIN_CAPTURE_OURS.store(false, Ordering::Release);
}

/// Whether the capture being lost is one this app released on purpose.
pub(super) fn pin_capture_is_ours() -> bool {
    PIN_CAPTURE_OURS.load(Ordering::Acquire)
}

/// Hide or show the window FFmpeg is drawing the picture into.
///
/// It is hidden with `ShowWindow(SW_HIDE)` rather than by moving it off the screen or making it
/// zero-sized: this is a window of another process, and the two cheapest ways to hide a window
/// this app does not own are to lie about its size or to lie about where it is. `ShowWindow` is the
/// one that says what it means, and it is the same call the style monitor already makes to put the
/// window up without activating it (see `ensure_video_window_topmost`).
pub(super) fn set_pinned_player_window_visible(visible: bool) {
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

pub(super) fn show_pinned_player_window() {
    set_pinned_player_window_visible(true);
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
        return false;
    };

    release_the_pointer(window, hwnd.0 as isize);

    // A resize relaunches at the size the drag settled on (see `relayout_pinned_media`), and a
    // relaunch is a player being ended and another begun — so it is the relaunch that puts the
    // parked window back, and putting it back here as well would show a window that is about to be
    // taken down again, for as long as the relaunch takes. A move has no relaunch, so it is the
    // move's own release, which is also the drag that has to place the picture rather than only
    // show it (see `unpark_pinned_player`).
    let relaunching = matches!(drag.action, PinDragAction::Resize(_))
        && pinned_content().is_some_and(|content| {
            if let Ok(mut request) = PIN_BOX_REQUEST.lock() {
                *request = Some(content);
            }
            true
        });
    if !relaunching {
        unpark_pinned_player(window);
    }

    window.repaint();
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
/// **The park goes with it, which is the whole of what this path was missing.** A drag of a video
/// parks the player's window and paints the band flat over the hole that leaves (see
/// `park_pinned_player`), and both halves were undone by the drag's own release and by nothing else —
/// so a capture taken from under a drag ended the drag and left the park standing. The film went
/// back to playing regardless, because the tick watches the gesture and not the park, and what was
/// left on screen was an opaque band with no picture behind it for the rest of the pin's life and not
/// for the length of the drag. Only a relaunch took it, and a relaunch is not something a user does
/// to fix a window that has gone black.
///
/// So this is the release's work without the release's moment: the pointer's own facts are dropped,
/// the picture is put back before the tick lets the film go — that order being the one worth having
/// for the same reason it is there (see `finish_pin_drag`) — and the band is transparent again
/// because the flag it reads is down. What there is no `hwnd` for is a capture this app has already
/// lost: Windows has taken it, and releasing it would be a courtesy owed to nobody.
///
/// **Only a capture somebody else took gives the park up, and that is the whole of what this path
/// gets wrong by not saying.** Windows delivers `WM_CAPTURECHANGED` to the window that lost the
/// capture whether it released it or another window took it, and it says in no way the two can be
/// told apart afterwards — `GetCapture()` answers the same thing for both, which is that this
/// window is not the one holding anything. But this app's own `ReleaseCapture` is how a drag ends,
/// and the road letting go of the pointer is already deciding on the way out what happens to the
/// park: a resize is a relaunch, and a relaunch is a player with a window of its own that puts
/// that window up and takes the flag down (see `restart_pinned_player`). A capture-loss answering
/// its own app's release by putting the parked window back undoes that deliberate hold-off — the
/// outgoing player comes up at the box the drag began at, over a band the flag has already stopped
/// painting flat, for as long as the relaunch takes. So a notice this app raised itself is answered
/// as a notice and not as a loss: the pointer's facts still go, and the park is left standing for
/// the road that let go of it to take back (see `release_the_pointer`).
pub(super) fn pin_capture_lost(window: &dyn PinWindow) {
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

    // This app's own release, and the road that asked for it is still inside the call: the park is
    // that road's to take back, with the relaunch decision in hand that a release delivered to a
    // window procedure has not got (see `finish_pin_drag`).
    if pin_capture_is_ours() {
        return;
    }

    // Asked of the pin rather than taken from the lock above, because putting another process's
    // window up is not work to be done while this app's own lock is held — the same seam
    // `finish_pin_drag` keeps its `ReleaseCapture` outside of, for the same reason.
    unpark_pinned_player(window);
}

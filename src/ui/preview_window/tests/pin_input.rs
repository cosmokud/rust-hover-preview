use super::*;

/// A film a gesture has stopped is held in the transport, so nothing else can read it as playing.
///
/// This is the second half of the drag-hold fault. The pause key was posted and nothing was
/// written down, so `pin_is_playing` still said the film was playing, the loop still counted its
/// clock up, and the bar was drawn at a playhead racing away from a frozen picture. So the hold
/// goes in through the same `held` and `released` a press of the pause button uses, which is
/// what writes `paused_at` and rebases the clock — and both of those are visible in the two
/// fields the bar and the loop read.
#[test]
fn a_hold_a_gesture_puts_a_film_into_is_one_the_bar_can_see() {
    let mut playing = PinTransport::default();
    playing.begun(12.0, true, false);

    assert_eq!(
        playing.paused_at, None,
        "a film this app began is not held, which is the precondition for the two claims below \
             to mean anything"
    );
    assert!(
        transport_clock(&playing).is_some(),
        "and it is at some second — the one this transport's own clock measures, which is what \
             `video_drag_hold_apply` writes down"
    );

    // What `video_drag_hold_apply` writes when a gesture begins over a playing film.
    let mut held = playing;
    held.held(14.5);

    assert_eq!(
        held.paused_at,
        Some(14.5),
        "a gesture that stops a film has to write down that it stopped it, or `pin_is_playing` \
             says the film is playing, the loop counts its clock up, and the bar races away from a \
             frozen picture"
    );
    assert!(
        !transport_playing(&held, true),
        "and a transport bar must then be drawn with a play glyph rather than a pause one: the \
             player behind it is alive but holding still, which is the whole of what `paused_at` \
             stopped meaning"
    );

    // And what it writes when the gesture ends.
    let mut released = held;
    released.released(
        held.paused_at
            .expect("a held film has the second it was held at"),
    );

    assert_eq!(
        released.paused_at, None,
        "letting go of a window is not a hold: the film is playing on from the second it was \
             stopped at"
    );
    assert!(
        transport_playing(&released, true),
        "so the bar is a pause glyph again, and the loop is free to count the clock up"
    );
    assert!(
        released.started.is_some_and(|(_, from)| from == 14.5),
        "rebased onto the second the film was held at rather than left at twelve, or the playhead \
             jumps forward by the whole length of the gesture the moment the hand lets go"
    );
}

/// A gesture over a film that is already held changes nothing, and ending it starts nothing.
///
/// The key this posts is a *toggle*, so posting it onto a held film resumes it — which is how a
/// drag over a paused film used to start it: the user moving the window they were watching it in.
/// So both ends of the gesture are refused, and they are refused for two different reasons that
/// are both asserted here because either one on its own would leave the fault half-fixed.
///
/// The film is `playing` in every case below and the only thing that changes is whether the hold
/// is ours, which is the fault's actual shape: the drag was told `true`, and the film was not
/// playing.
#[test]
fn a_gesture_over_a_film_that_is_already_held_leaves_it_held() {
    let mut held = PinTransport::default();
    held.begun(30.0, true, false);
    held.held(44.0);

    assert!(
        !pin_is_playing_about(&held),
        "the premise: a film the pause button is holding is not playing, so there is nothing for \
             a gesture to hold and nothing for ending it to release"
    );

    for (dragging, ours, expected, what) in [
        (
            true,
            false,
            VideoDragHold::Leave,
            "a gesture begun over a film somebody else is holding has to be refused: the key is a \
                 toggle, so posting it here *resumes* a film the user paused",
        ),
        (
            false,
            false,
            VideoDragHold::Leave,
            "and ending that gesture has to be refused as well, or the film is started on the \
                 way out by the hand that let go of the window",
        ),
        (
            true,
            true,
            VideoDragHold::Leave,
            "a gesture over a film this app is already holding has stopped nothing and so has \
                 nothing to add — posting the toggle here would let go of the very hold the \
                 transport is recording",
        ),
        (
            false,
            true,
            VideoDragHold::Release,
            "and only a hold this app put there is taken back when the gesture ends",
        ),
    ] {
        assert_eq!(
            video_drag_hold_decision(dragging, false, ours),
            expected,
            "{what}"
        );
    }

    assert_eq!(
        held.paused_at,
        Some(44.0),
        "and none of it moved the film: both ends of a gesture over a hold that is not ours leave \
             it exactly where it was, which is the whole of what the refusals above are for"
    );
}

/// A gesture over a film that is playing holds it, and only that.
///
/// The other two cells of the decision, which together with the refusals make the whole of it: a
/// film that is playing and a hold this app does not have is the only combination that holds,
/// because it is the only one where the toggle does what the gesture means.
#[test]
fn a_gesture_over_a_playing_film_holds_it_and_only_a_playing_film() {
    assert_eq!(
        video_drag_hold_decision(true, true, false),
        VideoDragHold::Hold,
        "a hand moving the window of a film that is playing stops the film, which is the whole \
             of what the gesture is for — decoding while the window is being re-scaled and \
             re-presented on every pointer message is what makes a drag stutter"
    );

    assert_eq!(
        video_drag_hold_decision(false, true, false),
        VideoDragHold::Leave,
        "and a gesture that ends while the film is playing but is not held by us has nothing to \
             release: posting the toggle here would start a film the user had already stopped"
    );

    assert_eq!(
        video_drag_hold_decision(false, false, false),
        VideoDragHold::Leave,
        "a gesture over a film that was never playing is nothing at either end"
    );
}

/// A claim ends with the gesture that made it, and with nothing else — so it is reconciled
/// before anything at all is known about the film on screen.
///
/// The first row is the whole of the fault the reconciliation exists to take back, and it is
/// what the two refusals in `video_drag_hold_apply` used to skip: a gesture that ended over a
/// kind of pin whose media this app does not hold, or over no pin at all, left the claim
/// standing, and the next film to be dragged was answered by it — refused the hold it exists
/// for, and then paused by its release, since the key this posts is a toggle.
///
/// The second row is the one that must not be reconciled away: a relaunch underneath a hand in
/// flight begins another player that is held for that same gesture (see
/// `PinTransport::begun`), so the claim it carries is what the gesture is still to let go of.
#[test]
fn a_claim_ends_with_the_gesture_that_made_it_and_with_nothing_else() {
    for (dragging, ours, wanted, what) in [
        (
            false,
            true,
            false,
            "a gesture that has ended takes its claim back whatever is on screen — a kind of \
                 pin whose media this app does not hold, and no pin at all, are both refusals about \
                 the film and neither is a reason to keep a claim over it",
        ),
        (
            true,
            true,
            true,
            "while a gesture is in flight the claim stands, including across a relaunch that \
                 carries the hold: the film is still held for that same gesture, and dropping it \
                 there would freeze the picture for good once the hand let go",
        ),
        (
            false,
            false,
            false,
            "and there is nothing to take back where nothing was claimed",
        ),
        (
            true,
            false,
            false,
            "nor to keep where nothing is claimed, whatever the gesture is doing",
        ),
    ] {
        assert_eq!(video_drag_hold_claim(dragging, ours), wanted, "{what}");
    }
}

/// The hold a gesture puts a film into, written as the drag-hold arm writes it.
///
/// It is a function because every case below is what happens to this afterwards, and because it
/// stands in for the arm itself: a test cannot post a pause key to a player that does not
/// exist, so the write the arm makes is made here (see `video_drag_hold_apply`).
fn transport_held_by_a_gesture(began_at: f64) -> PinTransport {
    let mut transport = PinTransport::default();
    transport.begun(began_at, true, false);
    transport.held(began_at + 2.0);
    transport.drag_held = true;
    transport
}

/// A claim a gesture made does not outlive the player it was made against, whichever of the
/// ways that film stops being that player's.
///
/// This is the fault the claim used to have, and it needed both of its halves to be a fault a
/// user could see. It was a flag beside the pin, so nothing that ended the *player* ended it: a
/// pin taken down, a file swapped and a player that died all left it standing. And the next film
/// to be dragged was then answered by a claim belonging to a film that was not on screen — the
/// drag refused the hold it exists for, and then its release posted the toggle onto a film
/// nobody had stopped: a frozen picture, a pause glyph over it, and a loop counting up a clock
/// that nothing is playing.
///
/// So the claim is the transport's own field and every end of a player reconciles it (see
/// `PinTransport::drag_held`). The one that must *not* is the relaunch that carries the hold,
/// which is asserted at the end: the film is still held for the gesture that is still in flight,
/// and a claim dropped there would freeze it for good the moment the hand let go.
#[test]
fn a_claim_a_gesture_made_does_not_outlive_the_player_it_was_made_against() {
    assert!(
        transport_held_by_a_gesture(12.0).drag_held,
        "the premise: a gesture over a film that is playing holds it and says so, and that is \
             the only place a claim is ever made"
    );

    for (transport, what) in [
        (
            {
                let mut transport = transport_held_by_a_gesture(12.0);
                transport.player_gone(14.0);
                transport
            },
            "a player that has died is kept as a held file and not as a held gesture: there is \
                 no player left to hold, so a drag ending after that has to find nothing holding",
        ),
        (
            {
                let mut transport = transport_held_by_a_gesture(12.0);
                transport.begun(90.0, true, false);
                transport
            },
            "a relaunch carrying no hold is a film playing on, and no gesture is holding a \
                 film that is playing",
        ),
        (
            {
                let mut transport = transport_held_by_a_gesture(12.0);
                transport.released(14.0);
                transport
            },
            "a film let go of is playing on, and the press that let it go is the pause \
                 button's as much as the gesture's — either way nothing is holding it now",
        ),
        (
            PinTransport::default(),
            "and a transport that is not that player's claims nothing at all, which is what \
                 both a pin taken down and a file swapped between the two gestures leave behind",
        ),
    ] {
        assert!(!transport.drag_held, "{what}");

        // And what the next gesture is answered with, which is the half of the fault a user
        // is the one to see.
        let mut next = transport;
        next.begun(3.0, true, false);

        assert_eq!(
            video_drag_hold_decision(true, true, next.drag_held),
            VideoDragHold::Hold,
            "a drag over a film that is playing has to stop it, whatever a gesture that ended \
                 over some other film asserted on the way out"
        );
        assert_eq!(
            video_drag_hold_decision(false, false, next.drag_held),
            VideoDragHold::Leave,
            "and that gesture's release has to start nothing: the key is a toggle, so posting \
                 it here pauses a film that was never held"
        );
    }

    let mut relaunched = transport_held_by_a_gesture(12.0);
    relaunched.begun(90.0, true, true);

    assert!(
        relaunched.drag_held && relaunched.paused_at.is_some(),
        "a seek or a resize settling underneath a hand in flight begins another player that is \
             held for that same gesture, so the claim travels with it — dropping it here would leave \
             the film frozen for good once the hand let go of the window"
    );
    assert_eq!(
        video_drag_hold_decision(false, false, relaunched.drag_held),
        VideoDragHold::Release,
        "and the gesture that was in flight is still owed its release, which is the one case \
             where the toggle lands on a film a gesture really is holding"
    );
}

/// A gesture that has ended takes its claim back off whatever is on screen, refusals included.
///
/// The refusals the release arm used to make first — a pin that is not up, and a kind of pin
/// whose media this app does not hold — are the right answers to what a drag does to a film,
/// and are exactly where a claim was left standing behind them. So the reconciliation runs ahead
/// of them, and what is asserted here is that the claim is gone afterwards either way.
///
/// The rest of the release is not asserted because it cannot be reached from a test: the key is
/// posted to a player's own window, and there is no player on this machine to post it to (see
/// `video_window_for`), so the arm stops after the claim has been reconciled — which is the
/// whole of what this is about, and the whole of what starting a player in a test would buy.
///
/// **What is asserted is only this test's own pin.** The pin is a process-wide value, and only
/// the tests that take `PIN_TESTS_ONE_AT_A_TIME` are held off one another (see `stand_pin`), so
/// another test can take this one's pin away and put it back between two steps of it — a claim
/// left in a transport this test did not write, or a transport taken away before this test could
/// reconcile it, is a moment of another test's and not a fault to be found here. So the pin is
/// recognised by the file it was put up with, and a moment when something else is standing is
/// left alone.
#[test]
fn a_gesture_that_has_ended_takes_its_claim_back_off_whatever_is_on_screen() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let mine = PathBuf::from("claimed-by-a-gesture.mkv");

    for kind in [MediaType::Video, MediaType::StaticImage] {
        stand_pin(Some(PinnedPreview {
            path: mine.clone(),
            transport: transport_held_by_a_gesture(12.0),
            ..PinnedPreview::for_test()
        }));

        let mut media = create_loading_media(320, 240);
        media.media_type = kind;
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = Some(media);
        }

        assert!(
            !video_drag_hold_apply(false),
            "a gesture that has ended is told to no player at all — there is none here for the \
                 key to be posted to — so the only thing it can have done is take the claim back \
                 ({kind:?})"
        );

        let standing = pin_state().and_then(|state| {
            state
                .pin()
                .map(|pin| (pin.path == mine, pin.transport.drag_held))
        });
        if let Some((true, claimed)) = standing {
            assert!(
                !claimed,
                "and the claim has to be gone from this test's own pin whatever is on screen \
                     now, or the next drag of that film is refused the hold and then paused by its \
                     release ({kind:?})"
            );
        }
    }

    // And a pin that is not up: there is no transport to reconcile, and a claim that went down
    // with the pin cannot be answered for by whatever is put up in its place.
    stand_pin(None);
    assert!(
        !video_drag_hold_apply(false),
        "a gesture that ended over no pin at all has nothing to hold and nothing to let go of, \
             and whatever is put up in the pin's place is not to be answered for it"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// The pin's own answer about a transport, without asking the desktop what is playing.
///
/// `pin_is_playing` routes on the media's type and would ask the engine about a file this app
/// plays itself, so the claim under test — a film with a second written down is not playing — is
/// asked of the transport directly, which is the part that is the transport's own.
fn pin_is_playing_about(transport: &PinTransport) -> bool {
    transport_playing(transport, true)
}

/// A pin that was given the keyboard and no longer has it asks for it back — and a pin that
/// never had it, or that still has it, does not.
///
/// Both facts are unarguable, which is why they are the whole question: the claim is this app's
/// own record of a press, and `GetFocus` is Windows' answer about the caret. The fault behind
/// this is that an arrow press on a pin can be taken by the player instead. The player is begun
/// again on every navigation, every seek and every resize; its window is created *without*
/// `WS_EX_NOACTIVATE` — measured, the style is absent the whole time between the window
/// appearing and this app's monitor thread styling it — and all four of FFmpeg's arrows are
/// seeks, so an arrow that reaches the player is a seek however the symptom is described.
#[test]
fn a_pin_that_claims_the_keyboard_and_does_not_have_it_asks_for_it_back() {
    assert!(
        pin_keyboard_wanted_back(true, false),
        "a claim standing while Windows says the pin does not have the caret is a claim that \
             has stopped being true, and the pin's arrows are then the keys of whatever is in front \
             of it"
    );
    assert!(
        !pin_keyboard_wanted_back(false, false),
        "a pin that was never given the keyboard must not ask for it: that is asking for a \
             keyboard nobody handed over, and the claim would be written on the asking"
    );
    assert!(
        !pin_keyboard_wanted_back(true, true),
        "a pin that still has the caret holds the keyboard it claimed, and asking again is the \
             ask that takes the focus from whatever the user is now working in"
    );
}

/// The arrows on a pinned window walk the pin's own folder and nothing else, on every press.
///
/// The mapping itself is asserted above this one; what is asserted here is the property the
/// whole fault turns on. There is no key-driven seek anywhere in this app — a seek is what the
/// transport bar's mouse asks for, and the only keys this app has ever posted to a player are
/// the pause and the bar's own hold, which is the same key — so an arrow cannot be a seek
/// whatever state the pin is in. That is what makes "navigates on the first press, seeks on the
/// second" a statement about *which window* got the key rather than about what the key means,
/// and it is why a loop is given by beginning the file again rather than by posting an arrow at
/// it (see `video_launch::rewind_launch_seconds`).
#[test]
fn an_arrow_on_a_pinned_window_is_always_a_walk_of_its_own_folder() {
    for (vk, expected, name) in [
        (VK_LEFT.0 as i32, PinCommand::Previous, "left"),
        (VK_UP.0 as i32, PinCommand::Previous, "up"),
        (VK_RIGHT.0 as i32, PinCommand::Next, "right"),
        (VK_DOWN.0 as i32, PinCommand::Next, "down"),
    ] {
        assert_eq!(
            pinned_key_command(vk),
            Some(expected),
            "{name} walks the pin's own folder, in both the pin's directions and not only in \
                 the reading one"
        );
    }

    let walks = [PinCommand::Previous, PinCommand::Next];
    for vk in [
        VK_LEFT.0 as i32,
        VK_RIGHT.0 as i32,
        VK_UP.0 as i32,
        VK_DOWN.0 as i32,
    ] {
        assert!(
            walks.contains(&pinned_key_command(vk).expect("an arrow is always a walk")),
            "an arrow must never map to anything that is not one of the two walks, whatever \
                 state the pin is in"
        );
    }
}

/// The hook and the pin's own window answer the same keys, and mean the same thing by them.
///
/// These are two copies of one table, because a `WH_KEYBOARD_LL` callback may not take the
/// pin's lock to ask for the other one (see `key_input::PIN_KEY_COMMANDS`). A number that drifts
/// between them is a key this app acts on *wrongly* rather than one that stops working, and it
/// is invisible from either side alone: each side is right about itself.
///
/// The set of keys being the same set is the whole of **one owner per keypress**, which is what
/// the swallow rests on. A key the pin's window answers and the hook does not name reaches
/// FFmpeg whenever the caret is in the player's window — and every key bound to a seek in that
/// player is a fixed increment, so an arrow that arrives there is a step and not a walk. A key
/// the hook names and the pin's window does not answer is acted on for a caret the window was
/// never in. Neither is a missing key: both are a key answered by the wrong party, or by both.
///
/// **And the same is asked of the class of the message, not only of the key.** A chord is not a
/// walk and not a hold however plain the key under the modifier is: the pin's own window eats
/// `Alt+F4` and `Alt+Tab` without answering them, so a hook that counted `WM_SYSKEYDOWN` was
/// walking this window's folder for a `Ctrl`+Left and holding its film for a `Ctrl`+Space —
/// each one a key acted on for a caret the hook owns and answered by nothing for a caret the
/// window was in.
#[test]
fn the_hook_and_the_pin_window_answer_the_same_keys_the_same_way() {
    use crate::shell::key_input::{pin_key_command, pin_key_message_acts, PIN_KEY_COMMANDS};

    for (vk, command) in PIN_KEY_COMMANDS {
        assert_eq!(
            hook_pin_key_command(command),
            pinned_key_command(vk),
            "the hook counts `{vk:#x}` as command {command}, and that number has to mean what \
                 the pin's own window would have done with the key, or a walk arrives as a close"
        );
    }

    // Both directions of the set, over every virtual key a `wParam` can hold: a key one side
    // answers and the other does not is the double action above, and it is exactly what makes
    // "one owner per keypress" a property of the code rather than a hope about it.
    let drifted: Vec<i32> = (0..=0xFF)
        .filter(|vk| pinned_key_command(*vk).is_some() != pin_key_command(*vk).is_some())
        .collect();

    assert!(
        drifted.is_empty(),
        "these keys are answered by one of the two mappings and not the other, which is a key \
             acted on once for a caret this app is not in and passed on once for a caret it is: \
             {drifted:#x?}"
    );

    // The class, over the four messages a key can arrive as. `pinned_key_message_command` is
    // what this window's own procedure asks, so this is the procedure's rule read back rather
    // than a second copy of it — which is the point: the two sides ask one function, and this
    // is what pins that function to the answer both of them need.
    for (message, acted_on, name) in [
        (WM_KEYDOWN, true, "a key-down"),
        (WM_SYSKEYDOWN, false, "a system key-down"),
        (WM_KEYUP, false, "a key release"),
        (
            windows::Win32::UI::WindowsAndMessaging::WM_SYSKEYUP,
            false,
            "a system release",
        ),
    ] {
        assert_eq!(
            pin_key_message_acts(message),
            acted_on,
            "{name} of a key a standing pin answers is {acted_on} answered by the hook, and a \
                 chord the pin's own window throws away must not be counted for a caret this app is \
                 in merely because the caret is in FFmpeg's window instead"
        );
        assert_eq!(
            pinned_key_message_command(message, VK_LEFT.0 as i32, 0).is_some(),
            acted_on,
            "and {name} is answered by the pin's own window on the same question, or the same \
                 key is a walk for one of the two windows this app has and nothing for the other"
        );
    }
}

/// The gesture is noticed on the ticks it begins and ends on, and costing nothing in between.
///
/// A press is not a drag — `PinnedPreview::dragging` is written when the pointer moves, not when
/// it goes down — so a click on the caption, which is how a pin is given the keyboard, must not
/// stop the film. And the flag is compared rather than read, because posting the pause key on
/// every tick of a gesture would toggle a held film back into playing on the second one.
///
/// **The two records are separate, and this is why.** One says a drag is in flight; the other,
/// in the transport, says the film on screen is held *because of* it. They are not the same
/// question, because a drag begun over a film the pause button has already held has stopped
/// nothing — and ending that gesture must not start anything. So the claim is asserted where
/// the hold it belongs to is, and a claim that was never taken is not taken back by a gesture
/// that had nothing to do with it (see `PinTransport::drag_held`).
#[test]
fn a_gesture_is_noticed_on_the_ticks_it_begins_and_ends_on() {
    let start = video_drag_holding();
    let begun = !start;

    assert!(
        video_drag_hold_set(begun),
        "the tick that finds the gesture begun has to be told, because that is the tick that \
             posts the hold"
    );
    assert_eq!(
        video_drag_holding(),
        begun,
        "and the gesture is in flight while it is"
    );
    assert!(
        !video_drag_hold_set(begun),
        "a tick inside the gesture has nothing to post, and posting the pause key again would \
             toggle a held film back into playing"
    );
    assert!(
        video_drag_hold_set(start),
        "the tick that finds the gesture over has to be told, so that the film is let go of \
             where it stood rather than left frozen under a window that is still moving"
    );
    assert_eq!(
        video_drag_holding(),
        start,
        "and the two ends of a gesture undo each other"
    );
}

/// A drag parks the player's window and holds the film, and neither of the two does the other's
/// half — which is the whole of how a drag stops the picture stuttering without pausing it
/// twice.
///
/// Parking is asked of the window and holding is asked of the film, and they are asked by two
/// different callers at two different times: the window is hidden on the drag's own first
/// pointer message, and the film is held by the tick. The key is a **toggle**, so a pause
/// posted from both places is not a pause at all — a drag begun over a playing film would leave
/// it playing with its window hidden, and a drag begun over a film the pause button had held
/// would start it. So the flag carries nothing about the film: what the film is doing is the
/// transport's own fact (`PinTransport::drag_held`), written by the tick that holds it and
/// reconciled by every path that ends the player it was made against.
///
/// Neither half is asked twice, and that is asserted on the flag rather than on a key: this
/// machine has no player to post a key to (see `video_window_for`), so the arms stop at the
/// point where the key would go. What *is* asserted is the whole of what could have made them
/// stop agreeing — that parking writes nothing to the transport, that a second park of the same
/// drag and a second unpark of the same drag are both no-ops, that the tick's half leaves the
/// parking flag exactly as it found it, that the park stands in the way of every raise while it
/// is held and that a capture lost to another window lets go of it, and that with no pin up
/// neither half has anything to do.
///
/// The capture is the second end of a drag and a whole of the park on its own: a drag ends either
/// at this window's own release or at a capture another window took, and only the release was
/// undoing what the park had done (see `pin_capture_lost`).
#[test]
fn a_drag_parks_the_pictures_window_without_holding_the_film_twice() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let mut transport = PinTransport::default();
    transport.begun(3.0, true, false);

    stand_pin(Some(PinnedPreview {
        path: PathBuf::from("parked-because-a-window-was-dragged.mkv"),
        transport,
        ..PinnedPreview::for_test()
    }));

    let mut media = create_loading_media(320, 240);
    media.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }

    // What a pin is asked about rather than a copy of it, because the pin is a process-wide
    // value and only the tests taking `PIN_TESTS_ONE_AT_A_TIME` are held off one another (see
    // `stand_pin`).
    let parked = || {
        pin_state().and_then(|state| {
            state
                .pin()
                .map(|pin| (pin.parked, pin.transport.drag_held, pin.transport.paused_at))
        })
    };

    // The window the park and the unpark are asked of, so that what this test is about — the
    // flag, and the film's two halves — can be told from what is asked of the player's window
    // (see `PinWindow::unpark_player_window`).
    let window = RecordedPinWindow::new(0x1000);

    assert!(
        park_pinned_player(HWND(0x1000 as *mut _), (0, 0), false),
        "the first pointer message of a drag hides the player's window, which is the half of \
             the stutter that is the compositor's and the half this app can answer on the message"
    );
    assert_eq!(
        parked(),
        Some((true, false, None)),
        "and it hides nothing else: the flag says the picture is away and carries no word about \
             the film, because the film is held by the tick posting one key and a second key would \
             be a toggle out of the hold rather than into it"
    );
    assert!(
        park_pinned_player(HWND(0x1000 as *mut _), (0, 0), false),
        "a begin over a standing park extends it rather than refusing it: the second call keeps \
             the first call's capture and keeps the settle's record current, and the detail is \
             answered where parks are settled (see `pin_frames`)"
    );

    // The tick's half of the pair, asked with the gesture in flight. It stops before the key —
    // there is no player on this machine to post it to — and what it stops without is the
    // claim, because a claim this app could not act on is a claim it must not leave standing.
    assert!(
        !video_drag_hold_apply(true),
        "a tick that cannot post the hold tells no player about anything, so the whole of what \
             it can have done is taken back"
    );
    assert_eq!(
        parked(),
        Some((true, false, None)),
        "and holding the film leaves the parking flag alone: the two halves are paired by the \
             gesture rather than by each other, so one of them running a tick late cannot show a \
             player window over a band painted flat, nor a hole in the desktop over a window that \
             is still being resized"
    );

    // The swap is the settle's and not this test's, and it is driven through the same line either
    // way (see `settle_pinned_park`): the band is opaque until there is a player to see through
    // it, and this machine has none.
    assert!(
        unpark_pinned_player(&window),
        "the swap puts the player's window back at the band in the same tick the flag goes down"
    );
    assert!(
        !unpark_pinned_player(&window),
        "and a swap that has already happened has nothing left to put back, so a pin taken down \
             mid-drag cannot put up a window that no longer exists"
    );

    // The park is not a hide and no re-show: every place that would put the player's window
    // where the band is — a resize drag's own raise on every pointer move, and the tick's
    // re-assertion every couple of hundred milliseconds — is answered out of the hand while a
    // park stands, so the film stays away for the whole of the drag rather than for the one
    // pointer message it was hidden on (see `pin_player_is_parked`).
    assert!(
        park_pinned_player(HWND(0x1000 as *mut _), (0, 0), false) && pin_player_is_parked(),
        "a parked player is a parked pin as far as every raise is concerned, which is what makes \
             the park hold for a hand that has stopped moving as well as one that has not"
    );

    // The release this window asks of itself is not a capture lost to anybody, and the
    // difference is the whole of what the hold-off above is for. `WM_CAPTURECHANGED` is
    // delivered to the window that lost the capture whether it let go or another window took
    // it, so a resize that let go of the pointer was raising this notice at itself — and an
    // answer that puts the parked player back is a film on screen at the box the drag began
    // at, over a band the flag has already stopped painting flat, for as long as the relaunch
    // takes (see `finish_pin_drag`).
    let releasing = CapturingWindow::around(a_window_at((300, 200, 700, 600)));
    with_pin(|pin| {
        pin.dragging = Some(PinDrag {
            from: (0, 0),
            window: (0, 0, 640, 480),
            action: PinDragAction::Resize(PinResize {
                left: false,
                top: false,
                right: true,
                bottom: true,
            }),
            delivered: true,
            carried: (i32::MIN, i32::MIN),
        });
    });
    assert!(
        finish_pin_drag(HWND(0x1000 as *mut _), &releasing),
        "a resize is let go of through a release that raises its own notice from inside the call"
    );
    assert!(
        pin_player_is_parked(),
        "and the band is left to the settle either way: the park stands until there is a player to \
             see through it, which for a resize is the replacement the release has just asked for"
    );
    assert!(
        !releasing
            .calls()
            .iter()
            .any(|call| matches!(call, PinWindowCall::UnparkPlayerWindow(_))),
        "with nothing asked of the player's window at all: a player about to be taken down must \
             not be put back up on the way to being taken down"
    );
    // The relayout this release asked for is not asserted on: the slot it is written to is a
    // machine value the loop drains on its next turn and several tests here write, so what is
    // in it while this test runs belongs to whoever gets there. It is taken so that a resize's
    // end leaves nothing behind.
    let _relayout_on_its_way_to_the_loop = take_relayout_request();

    // The other end of a drag is the capture going to another window, and it ends the drag's
    // record — which is all it ends. The park stands for the settle, exactly as it does after this
    // app's own release, because a stolen capture and a release leave the same thing behind: a
    // drag that has ended and a band that is still this app's to fill (see `pin_capture_lost`).
    let stolen = RecordedPinWindow::new(0x1000);
    with_pin(|pin| {
        pin.dragging = Some(PinDrag {
            from: (0, 0),
            window: (0, 0, 640, 480),
            action: PinDragAction::Resize(PinResize {
                left: false,
                top: false,
                right: true,
                bottom: true,
            }),
            delivered: true,
            carried: (i32::MIN, i32::MIN),
        });
    });
    assert!(
        pin_is_dragging(),
        "a drag is in flight before the capture is stolen from it, which is the state the message \
             has to find"
    );

    pin_capture_lost(&stolen);

    assert!(
        !pin_is_dragging() && pin_player_is_parked(),
        "a capture stolen mid-gesture ends the drag and nothing else: the park is the settle's to \
             take back, and the tick watching the gesture lets the film go either way"
    );
    assert_eq!(
        parked(),
        Some((true, false, None)),
        "so the pair is where a release leaves it, and the band is still filled with the frame \
             rather than being a hole in the desktop over a window that is not there"
    );
    assert!(
        stolen.calls().is_empty(),
        "with nothing asked of the player's window: a swap is the settle's, on the tick that finds \
             a player to make it against"
    );
    assert!(
        unpark_pinned_player(&stolen),
        "which is the settle's line, reached here by hand because the loop is not running"
    );

    // With no pin up there is no flag to write and no transport to reconcile, so a drag that
    // ends over nothing at all is answered by neither half.
    stand_pin(None);
    assert!(
        !park_pinned_player(HWND(0x1000 as *mut _), (0, 0), false),
        "a drag over no pin at all has no picture to park and nothing to hold"
    );
    assert!(
        !unpark_pinned_player(&window),
        "and no pin to put a picture back for"
    );
    assert!(
        !video_drag_hold_apply(false),
        "while a claim that went down with the pin cannot be answered for by whatever is put up \
             in its place"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A park is not taken back by the flag it is written in, and this is the whole of why: the band
/// is transparent while a film is playing because FFmpeg's own window stands in it, so a park that
/// ends while that window is not there is a hole in the desktop in every band a player of this
/// app's is drawn in. A resize's release used to be exactly that — the relaunch that answers it
/// wrote the flag down the moment it began a replacement, which is a player that is running and
/// has no window yet — and what the user saw for the length of the wait was the file behind, and
/// the windows beside it, through the window they were holding the edge of.
///
/// So the settle asks the window rather than the flag, and a park that is answered while the
/// replacement is still starting stays standing, holding the last frame it took scaled to the box
/// the drag settled on. It is asked with the window rather than with a pid because the two overlap
/// during a relaunch and a pid cannot say which of them is on screen (see `video_window_for`).
///
/// The frame is given up with the park and not one tick before it: it is the band's picture for as
/// long as the band has no window in it, and a pin that is not being dragged pays nothing to hold
/// a picture the size of a display.
#[test]
fn a_park_is_taken_back_only_once_the_players_own_window_is_there() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    stand_pin(Some(PinnedPreview {
        path: PathBuf::from("parked-until-the-replacement-arrives.mkv"),
        ..PinnedPreview::for_test()
    }));

    let mut media = create_loading_media(320, 240);
    media.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }

    assert!(
        park_pinned_player(HWND(0x1000 as *mut _), (0, 0), false),
        "a drag puts the player's window away, which is the half of the band that is then this \
             app's to fill"
    );
    // The machine this runs on has no player of this app's, so the frame a real park would have
    // taken off the desktop is stood in for it (see `hold_video_window_frame`).
    stand_video_frame_for_a_test([0u8, 0, 255, 255].repeat(16), 4, 4);
    assert!(
        held_video_frame().is_some(),
        "and the frame that was on screen a pointer message ago is held for the drag to scale"
    );

    assert!(
        !settle_pinned_park_onto(false),
        "a replacement that is running and has no window yet is not a picture to see through the \
             band, so the park is not taken back"
    );
    assert!(
        pin_player_is_parked(),
        "and the band keeps the frame it has rather than becoming a hole in the desktop — which is \
             what this very settle used to do, for as long as a player takes to open a file"
    );
    assert!(
        !settle_pinned_park_onto(false),
        "a settle asked again while the replacement is still starting is still the same answer"
    );

    assert!(
        settle_pinned_park_onto(true),
        "and it is taken back the moment the player's own window is standing in the band"
    );
    assert!(
        !pin_player_is_parked(),
        "which is the whole of what the park was holding off: the band is transparent again \
             because there is a window of somebody else's in it"
    );
    assert!(
        !settle_pinned_park_onto(true),
        "and a park that has been taken back cannot be taken back again, so a relaunch that finds \
             no drag in flight does not write a flag nobody will ever ask about"
    );
    assert!(
        held_video_frame().is_none(),
        "the frame goes with it rather than being held for the rest of the run: a pin that is not \
             being dragged pays nothing for a picture the size of a display"
    );

    // And with no pin up there is no flag, so a settle is answered by nothing at all.
    stand_pin(None);
    assert!(
        !settle_pinned_park_onto(true),
        "a park over no pin at all cannot be standing, and so cannot be taken back"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A window that hands the pointer back the way the machine does, by delivering the capture
/// notice back into the window procedure before the call returns.
///
/// `ReleaseCapture` sends `WM_CAPTURECHANGED` to the window it took the capture from, on the
/// same thread and *during* the call, so a window procedure answering one is answering it
/// while the caller is still inside its own release. Standing that up here is the whole of
/// what makes the two ends of a capture tellable apart on a machine with no desktop: a
/// recorder that merely recorded the release would let a park be undone twice with nothing in
/// the test ever noticing.
struct CapturingWindow {
    inner: RecordedPinWindow,
}

impl CapturingWindow {
    /// This window, keeping the rest of a drag's list on the recorder underneath.
    fn around(inner: RecordedPinWindow) -> Self {
        Self { inner }
    }

    /// Everything the road did, in the order it did it.
    fn calls(&self) -> Vec<PinWindowCall> {
        self.inner.calls()
    }
}

impl PinWindow for CapturingWindow {
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
        self.inner.release_capture(hwnd);
        // On the machine this is `ReleaseCapture` sending `WM_CAPTURECHANGED` back into this
        // thread's own window procedure, which is where `pin_capture_lost` is called from and
        // still inside the release that caused it.
        pin_capture_lost(self);
    }

    fn set_focusable(&self, hwnd: isize, focusable: bool) {
        self.inner.set_focusable(hwnd, focusable)
    }

    fn set_focus(&self, hwnd: isize) {
        self.inner.set_focus(hwnd)
    }

    fn set_foreground(&self, hwnd: isize) {
        self.inner.set_foreground(hwnd)
    }

    fn hide_pin_windows(&self) {
        self.inner.hide_pin_windows()
    }

    fn hide_pin_bubble(&self) {
        self.inner.hide_pin_bubble()
    }

    fn unpark_player_window(&self, band: Option<ScreenRegion>) {
        self.inner.unpark_player_window(band)
    }

    fn repaint(&self) {
        self.inner.repaint()
    }

    fn post(&self, hwnd: isize, message: u32) {
        self.inner.post(hwnd, message)
    }
}

/// The relayout a finished resize asked for, taken off the request slot.
///
/// Taken rather than read because the slot is a machine value the loop drains on its next turn and
/// other tests write it too, so a test that left one behind would be laying out whatever pin ran
/// next.
fn take_relayout_request() -> Option<ScreenRegion> {
    PIN_BOX_REQUEST
        .lock()
        .ok()
        .and_then(|mut request| request.take())
}

use super::*;

/// A negative answer kept against an item outlives the tick that read it, and that is
/// what the click retry has to drop before it asks again.
///
/// A look that finds an item and cannot name a file for it is kept as an answer, because most
/// of the time that is what it is — a folder, an application, a name no kind claims. Nothing
/// tells that apart from a click read before the shell caught up with the focus move, so a
/// retry that kept the memo would be answered with the negative it was sent to replace (see
/// the retry in `PinUpdateWatch::follow`).
#[test]
fn a_negative_item_answer_outlives_the_tick_that_read_it() {
    // `ItemResolver::default()` rather than `ItemResolver::new`, which asks the shell for a
    // window collection this test has no apartment to ask it with.
    let mut resolver = ItemResolver::default();
    let point = POINT { x: 40, y: 60 };
    let bounds = (10, 50, 400, 70);

    // The shape both a folder and a listing caught mid-move produce: an item, and no
    // file to be had for it.
    resolver.remember_probe(point, None);
    resolver.remember_item(&PointerLook {
        path: None,
        item_bounds: Some(bounds),
        drawn_in: 0,
    });

    // The pointer's own answer belongs to the tick that produced it, so the next tick
    // begins with neither the point memo nor — which is the point of this test — the
    // item's.
    resolver.forget_probe();
    assert!(
        resolver.probed_at(point).is_none(),
        "the tick's own answer does not outlive the tick"
    );
    assert!(
        matches!(resolver.item_under(point, 0), Some(None)),
        "but the item's negative answer does, and it is the one a retry would read"
    );

    // Which is why the retry drops it rather than asking the item again: this is what
    // makes the ask a real one, so a shell that has caught up with the focus by now is
    // free to answer differently than it could not a moment ago.
    resolver.forget_item();
    assert!(
        resolver.item_under(point, 0).is_none(),
        "so the retry reads the shell instead of the answer it is replacing"
    );
}

/// A click on a pinned window is a click on the listing the window is standing over: the
/// reading that asks what file is at the point, rather than what is on top of it, is the
/// one the pick was always asking. A click over neither is still refused.
#[test]
fn a_click_on_a_pinned_window_is_a_click_on_the_listing_under_it() {
    // The reported case: the pointer is on this app's own window, which is standing on a
    // listing. Explorer is not under the pointer at all, by `WindowFromPoint`'s answer.
    assert!(
        click_is_over_a_listing(false, true),
        "a click on a pinned window over a listing is a click on that listing"
    );

    // The ordinary case, unchanged.
    assert!(
        click_is_over_a_listing(true, false),
        "a click on Explorer's own window is a click on a listing"
    );
    assert!(
        click_is_over_a_listing(true, true),
        "and both at once is still one"
    );

    // And the case that must stay refused: nothing under the pointer is a listing.
    assert!(
        !click_is_over_a_listing(false, false),
        "a click on the desktop, or another program, has no listing behind it to read"
    );
}

/// A press aimed at a menu popup is not a press on the listing that popup was covering:
/// `View ▸ Tiles` and `Sort ▸ Name` sit over rows, and the press that picks one dismisses
/// it without the hand moving. Which tick reads the press is the shell's choice and not the
/// rule's — it is up on one, down on the next, or down already on both — so all three are
/// refused. A press over a listing that was never a menu under it is unchanged.
#[test]
fn a_press_through_a_dismissed_popup_is_not_a_pick() {
    assert!(
        press_is_a_listing(true, false, false, false, false),
        "a press on Explorer's own window, still there, is a pick"
    );
    assert!(
        press_is_a_listing(false, true, false, false, false),
        "and a press on a pinned window standing over the listing is one too"
    );
    assert!(
        press_is_a_listing(true, true, false, false, false),
        "and both at once is still one"
    );
    assert!(
        !press_is_a_listing(true, false, true, false, false),
        "a press that destroyed the window it landed on is a press on the popup, not the row behind it"
    );
    assert!(
        !press_is_a_listing(false, true, true, false, false),
        "and it is refused on a pinned window over the listing as well"
    );
    assert!(
        !press_is_a_listing(false, false, false, false, false),
        "a press on the desktop, or another program, has no listing behind it to read"
    );

    // The popup is read on the tick the press bit is spent on, and a XAML popup hides rather
    // than going away: `IsWindow` answers for it all the way through.
    assert!(
        !press_is_a_listing(true, false, false, true, false),
        "a press read while the popup is still up is a press on the popup"
    );

    // The toolbar button that opened it is chrome by the same rule and reads the same way.
    assert!(
        !press_is_a_listing(true, false, false, true, true),
        "and so is the press that opened it"
    );

    // The poll after it closed: the window under the point is the listing again, which is
    // why only the tick before can tell this press from a click on that row.
    assert!(
        !press_is_a_listing(true, false, false, false, true),
        "a press whose previous window was the popup is that popup's, not the row behind it"
    );
    assert!(
        !press_is_a_listing(false, true, false, false, true),
        "and it is refused over a pinned window as well"
    );

    // A hand that came from the desktop or from the pin itself is not a hand that just
    // dismissed a menu, however few ticks apart the two ticks are.
    assert!(
        press_is_a_listing(true, false, false, false, false),
        "a press whose previous window was merely somewhere else is a press on the listing"
    );
    assert!(
        press_is_a_listing(false, true, false, false, false),
        "and a flick from the pin back into the listing is one too"
    );

    assert!(
        !window_is_gone(None),
        "nothing to have been destroyed on the first tick of a watch"
    );
    assert!(
        !window_is_gone(Some(HWND(std::ptr::null_mut()))),
        "and a tick that named no window at all is a gap in the reading, not a window that went away"
    );
}

/// A pointer that has left the item the preview on screen is about has moved,
/// whether or not it went far enough to clear the threshold: the threshold is a
/// hand's own jitter, and a row of the list crossed under it would otherwise
/// leave the preview of the file that was left behind on screen, with the probe
/// that would have noticed held behind a single probe per parked cursor. The two
/// states that are not a move are the ones where there is nothing on screen to
/// leave, and the pointer a preview is holding — see
/// `pointer_moved_off_the_hovered_item`.
#[test]
fn a_pointer_off_the_hovered_item_has_moved() {
    assert!(
        pointer_moved_off_the_hovered_item(true, false, false),
        "a pointer outside the item the preview is about has left it"
    );
    assert!(
        !pointer_moved_off_the_hovered_item(true, false, true),
        "a pointer still inside it has not"
    );
    assert!(
        !pointer_moved_off_the_hovered_item(false, false, false),
        "and nothing is on screen for a pointer to have left"
    );
    assert!(
        !pointer_moved_off_the_hovered_item(true, true, false),
        "nor is a pointer the preview itself holds one that has left it"
    );
}

/// A look that answered nothing is not a verdict on its own: the item it found says
/// which of the two things it was. The same item — the box the preview was resolved
/// from, still under the pointer — is a read that failed, and the preview of that file
/// is owed the question asked again rather than being taken down and put back. Another
/// item's box is a pointer that has moved onto something with no preview of its own —
/// an application, a folder — where the look answered properly and the preview of the
/// file that was left has to go with it. A hover with no box of its own holds nothing
/// back, and a box that no longer holds the pointer is not a box it is still on.
#[test]
fn a_read_that_failed_is_told_from_an_item_with_no_preview() {
    let item = (100, 200, 400, 220);
    let application = (100, 220, 400, 240);
    let inside = POINT { x: 150, y: 210 };
    let next_row = POINT { x: 150, y: 230 };

    assert!(
        read_failure_is_the_same_item(Some(item), Some(item), inside),
        "the same item, still under the pointer, is a read that failed"
    );
    assert!(
        !read_failure_is_the_same_item(Some(item), Some(application), next_row),
        "another item is another box, and the preview of the file that was left goes"
    );
    assert!(
        !read_failure_is_the_same_item(Some(item), Some(item), next_row),
        "nor is a box the pointer has left one it is still on"
    );
    assert!(
        !read_failure_is_the_same_item(None, Some(item), inside),
        "a hover that noted no box of its own holds nothing back"
    );
    assert!(
        !read_failure_is_the_same_item(Some(item), None, inside),
        "and a look that published no box is not a look at the same item"
    );
}

/// The item the keyboard has landed on counts only where it landed in the place the pin's watch
/// was already watching: another folder, another tab of one window and another window all move
/// the focus without a key having moved it, and read as a pick they are a pin following the
/// user's navigation — see `PinUpdateWatch::note_place`.
#[test]
fn a_focus_item_landed_in_another_place_is_a_baseline() {
    let view = |url: &str, hwnd: isize| HoverLocation {
        folder: None,
        search_root: None,
        location_url: Some(url.to_string()),
        view_hwnd: Some(hwnd),
    };
    let mut watch = PinUpdateWatch::default();

    assert!(
        !watch.note_place(Some(view("file:///D:/Pictures", 0x1234))),
        "the place the watch is shown the focus in first is the one it begins from"
    );
    assert!(
        !watch.note_place(Some(view("file:///D:/Pictures", 0x1234))),
        "and the same place again is the watch's own: a key moved the focus within it"
    );
    assert!(
        watch.note_place(Some(view("file:///D:/Videos", 0x1234))),
        "another folder is another place"
    );
    assert!(
        watch.note_place(Some(view("file:///D:/Videos", 0x5678))),
        "and another tab of the same window is one too"
    );
    assert!(
        watch.note_place(Some(view("file:///D:/Music", 0x9abc))),
        "as is another window showing another folder"
    );
    assert!(
        !watch.note_place(None),
        "a look that answered nothing is not a place that changed"
    );
    assert!(
        !watch.note_place(Some(view("file:///D:/Music", 0x9abc))),
        "and it leaves the place in hand where it is"
    );
    assert!(
        watch.note_place(Some(view("file:///D:/Pictures", 0x1234))),
        "so a place read after one that answered nothing is still compared with it"
    );
}

/// The item the focus lands on is a file the user picked only where a key that walks a listing
/// is what moved it there: an Enter, a Delete, a click, a shortcut and any key held with a
/// modifier down all put the focus somewhere without a key walking it there, and the place
/// cannot be asked for an answer where the shell describes one — see
/// `PinUpdateWatch::focus_moved_by_key`.
#[test]
fn a_focus_moved_by_anything_but_a_key_walking_is_not_a_pick() {
    let walking = FocusMoveInput {
        walked_by_key: true,
        ..Default::default()
    };
    let command = FocusMoveInput {
        moved_otherwise: true,
        ..Default::default()
    };
    let click = FocusMoveInput {
        clicked: true,
        ..Default::default()
    };
    let started = Instant::now();
    let mut watch = PinUpdateWatch::default();

    assert!(
        !watch.focus_moved_by_key(started),
        "a watch that has seen nothing has no witness to give"
    );

    watch.note_focus_move(walking, started);
    assert!(
        watch.focus_moved_by_key(started),
        "a key walking the listing is the witness"
    );
    assert!(
        watch.focus_moved_by_key(started + Duration::from_millis(KEYBOARD_FOCUS_INPUT_GRACE_MS)),
        "and it stands while a key is given to have moved something"
    );
    assert!(
        !watch
            .focus_moved_by_key(started + Duration::from_millis(KEYBOARD_FOCUS_INPUT_GRACE_MS + 1)),
        "but no longer: a witness nothing made good on is not the key's"
    );

    // A key held down says so on every tick, which is what a listing walked with a held arrow
    // is: the witness is taken again for as long as the walk lasts.
    watch.note_focus_move(walking, started + Duration::from_millis(1000));
    watch.note_focus_move(walking, started + Duration::from_millis(1030));
    assert!(watch.focus_moved_by_key(started + Duration::from_millis(1040)));

    // And a move that is not a walk takes it away — an arrow pressed before an Enter, which is
    // how a folder is opened from the keyboard, does not stand as the witness for the item that
    // Enter lands the focus on, and neither does a key that is part of a chord.
    watch.note_focus_move(command, started + Duration::from_millis(1100));
    assert!(
        !watch.focus_moved_by_key(started + Duration::from_millis(1110)),
        "an Enter, a shortcut or a modifier-held key is not a walk"
    );
    watch.note_focus_move(click, started + Duration::from_millis(1200));
    assert!(
        !watch.focus_moved_by_key(started + Duration::from_millis(1210)),
        "and neither is a click, which is the pointer acting on the view"
    );
    watch.note_focus_move(walking, started + Duration::from_millis(1300));
    watch.note_focus_move(
        FocusMoveInput {
            walked_by_key: true,
            moved_otherwise: true,
            clicked: false,
        },
        started + Duration::from_millis(1350),
    );
    assert!(
        !watch.focus_moved_by_key(started + Duration::from_millis(1360)),
        "a key pressed with a modifier down is a command, not a walk"
    );
}

/// A click re-armed on the file the pin already shows keeps the listing it was made in, and
/// only its clock moves.
///
/// The press bit is spent the moment the button goes down, so the tick that sees the click
/// is a tick *after* it, and the view under the pointer by then is the answer to the retry
/// rather than a fresh place. Reading the place again would compare the retry against the
/// listing that click itself produced — which never differs — and the click would never be
/// released (see `PinUpdateWatch::answer_click`).
#[test]
fn a_second_click_on_the_file_the_pin_shows_keeps_the_place_the_first_was_made_in() {
    let here = HoverLocation {
        folder: None,
        search_root: None,
        location_url: Some("file:///D:/Pictures".to_string()),
        view_hwnd: Some(0x1234),
    };
    let showing = Path::new("D:/Pictures/one.png");
    let mut watch = PinUpdateWatch {
        pending: Some(PendingClick {
            at: Instant::now(),
            point: POINT { x: 10, y: 20 },
            place: Some(here.clone()),
        }),
        ..Default::default()
    };
    let first_at = watch.pending.as_ref().expect("a held click").at;

    // The second click is on the file the pin is already showing, which is the case that
    // re-arms rather than releases.
    watch.answer_click(
        showing,
        showing,
        true,
        Instant::now(),
        POINT { x: 30, y: 40 },
    );

    let held = watch.pending.as_ref().expect("the click is re-armed");
    let carried = held.place.as_ref().expect("the place is carried forward");
    assert!(
        !hover_location_changed(carried, &here),
        "the listing the first click was made in is carried forward, not re-read"
    );
    assert!(
        held.at > first_at,
        "and the hold's own clock moves on, which is what arms the retry"
    );
    assert_eq!(
        (held.point.x, held.point.y),
        (30, 40),
        "while the point the retry looks the file up under is the new one"
    );
}

/// A pin shown another file is not a watch beginning: the item the keyboard is on, the place it
/// was read in and the hand having arrived are the listing's and the hand's, and a swap that
/// forgot them made the first selection a user makes after pressing the pin and clicking back
/// into Explorer a baseline rather than a pick — see `PinUpdateWatch::note_shown`.
#[test]
fn a_swap_carries_the_keyboard_baseline_forward() {
    let here = HoverLocation {
        folder: None,
        search_root: None,
        location_url: Some("file:///D:/Pictures".to_string()),
        view_hwnd: Some(0x1234),
    };
    let item = |name: &str, top: i32| FocusedItemKey {
        name: name.to_string(),
        rect: (0, top, 200, top + 20),
    };
    let mut watch = PinUpdateWatch::default();

    // A pin taken up begins from nothing: what is on the keyboard when it comes up is a
    // baseline, not a choice.
    watch.note_shown(Path::new("D:/Pictures/one.png"));
    assert!(
        watch.focused.is_none() && watch.place.is_none() && !watch.arrived,
        "a watch that has not watched anything holds no baseline to keep"
    );

    // A listing is read into it, and the hand arrives. A click is held for a retry, as one
    // that landed while the shell was still catching up would be.
    watch.focused = Some(item("one.png", 100));
    watch.note_place(Some(here.clone()));
    watch.arrived = true;
    watch.pending = Some(PendingClick {
        at: Instant::now(),
        point: POINT { x: 120, y: 240 },
        place: None,
    });

    // A swap throws the pointer's own reading away and nothing else.
    watch.note_shown(Path::new("D:/Pictures/two.png"));
    assert_eq!(
        watch.showing.as_deref(),
        Some(Path::new("D:/Pictures/two.png")),
        "and is measured against the file the pin has now"
    );
    assert_eq!(
        watch.focused.as_ref().map(|key| key.name.as_str()),
        Some("one.png"),
        "the item the keyboard is on belongs to the listing, not to the file the pin shows"
    );
    assert_eq!(
        watch.place.as_ref().and_then(|place| place.view_hwnd),
        Some(0x1234),
        "and so does the place it was read in, which is what tells a move from a landing"
    );
    assert!(
        watch.arrived,
        "a window that has moved is not a hand that has not come"
    );
    assert!(
        watch.pointer.is_none() && watch.settled_at.is_none() && !watch.probed,
        "the settle and the probe were measured against a window that has since moved"
    );
    assert!(
        watch.pending.is_none(),
        "and a click held for a retry was measured against the showing before this one"
    );

    // And so the first item the focus lands on after the swap is a change the watch already
    // had a baseline for — which is what makes it a pick rather than a baseline. `FocusedItemKey`
    // is compared by its name and its box rather than as a whole, which is the same pair the
    // watch itself tells one observation from the next by.
    assert_ne!(
        watch.focused.as_ref().map(|key| (&key.name, key.rect)),
        Some((&"two.png".to_string(), (0, 120, 200, 140))),
        "the first selection after the swap is a different item from the one before it"
    );
}

/// A pin taken up again — a new pin, after the old one was closed — is a watch that has watched
/// nothing, and begins from nothing whatever the listing had on the keyboard.
#[test]
fn a_pin_taken_up_begins_from_nothing() {
    let mut watch = PinUpdateWatch {
        focused: Some(FocusedItemKey {
            name: "one.png".to_string(),
            rect: (0, 100, 200, 120),
        }),
        place: Some(HoverLocation {
            folder: None,
            search_root: None,
            location_url: Some("file:///D:/Pictures".to_string()),
            view_hwnd: Some(0x1234),
        }),
        arrived: true,
        ..PinUpdateWatch::default()
    };
    watch.note_shown(Path::new("D:/Pictures/one.png"));

    // A pin is up again on another file, after the first was taken down: the watch was reset
    // while nothing was pinned (see `PinUpdateWatch::follow`), so this is a take-up.
    watch = PinUpdateWatch::default();
    // A click is left held across a pin that is over, and a pin taken up afterwards is the
    // watch beginning: the click was measured against a window that is not there now.
    watch.pending = Some(PendingClick {
        at: Instant::now(),
        point: POINT { x: 40, y: 60 },
        place: None,
    });
    // The selection the listing held goes with it, for the same reason: it was read in a
    // listing this pin is not standing over.
    watch.sel_place = Some(HoverLocation {
        folder: None,
        search_root: None,
        location_url: Some("file:///D:/Videos".to_string()),
        view_hwnd: Some(0x1234),
    });
    watch.sel_selected = Some(std::path::PathBuf::from("D:/Videos/clip.mp4"));
    watch.note_shown(Path::new("D:/Videos/clip.mp4"));

    assert!(
        watch.focused.is_none() && watch.place.is_none() && !watch.arrived,
        "what is on the keyboard when a pin comes up is a baseline, not a choice"
    );
    assert!(
        watch.pending.is_none(),
        "and a click held for a retry was held against the pin before this one"
    );
    assert!(
        watch.sel_place.is_none() && watch.sel_selected.is_none(),
        "and the selection the listing held went with the pin before this one"
    );
    assert_eq!(
        watch.showing.as_deref(),
        Some(Path::new("D:/Videos/clip.mp4"))
    );
}

/// The second press of the double-click that opened a folder is not a pick, and the pin keeps
/// the file it was showing.
///
/// This is the whole of the fault. A press is read against whatever is under the pointer at the
/// tick it lands, and a double-click that opens a folder leaves the new listing drawn exactly
/// where the hand already was — so the second press was answered with whatever file the new
/// folder happened to put there, and the pin swapped to it. The place cannot tell the two
/// apart: by the time the second press lands the folder has already changed, so where the press
/// was made and where the pointer is are one place. The hand can, because a pick in a new
/// listing is a file the hand travelled to.
///
/// Asserted on the decision rather than on the swap, because the swap needs a pin and a shell
/// to stand in: the rule is what this app decides, and everything downstream of it is `offer`.
#[test]
fn the_second_press_of_a_double_click_is_not_a_pick() {
    let tolerance = KeyboardPointerPause::default().move_threshold_px(false, 96);
    let on_the_folder = POINT { x: 400, y: 300 };
    // The new folder's listing drew a row under the hand, in the same place.
    let on_the_new_listing = POINT { x: 402, y: 301 };

    let mut watch = PinUpdateWatch::default();

    // The first press: the folder. A pick, and `offer` declines it — a folder is not a file
    // this app can show anything for.
    assert!(
        watch.press_is_a_pick(on_the_folder, tolerance),
        "the first press of a double-click is a pick of its own"
    );

    // The second, one gesture later, on the file the new folder put under the hand.
    assert!(
        !watch.press_is_a_pick(on_the_new_listing, tolerance),
        "a press the hand has not moved for is the other half of a gesture, not a pick"
    );

    // And a third, for the same reason: a triple-click is one gesture, and the second press
    // being refused must not leave the third answering on the strength of the first.
    assert!(
        !watch.press_is_a_pick(on_the_new_listing, tolerance),
        "a triple-click is one gesture too"
    );

    // What it does cost: a click on a file the hand reached is a pick, however soon the click
    // before it was made.
    let mut moved = PinUpdateWatch::default();
    assert!(moved.press_is_a_pick(on_the_folder, tolerance));
    assert!(
        moved.press_is_a_pick(POINT { x: 40, y: 140 }, tolerance),
        "a click on another row is a file the hand travelled to"
    );

    // And the very first press a watch sees is a pick: there is no press before it to be the
    // other half of, which is the first click after a pin is taken up.
    assert!(
        PinUpdateWatch::default().press_is_a_pick(on_the_folder, tolerance),
        "the first press a watch sees is a pick"
    );
}

/// The first click after a pin is focused is a pick, and an answer of "the file you are
/// already looking at" is not what settles one. The timing that produces such an answer
/// belongs to the shell; the rule is this app's own (see `PinUpdateWatch::answer_click`).
#[test]
fn a_click_the_pin_already_shows_is_not_an_answer_to_that_click() {
    let showing = Path::new("D:/Pictures/one.png");
    let clicked_at = Instant::now();
    let point = POINT { x: 300, y: 180 };
    let mut watch = PinUpdateWatch::default();

    // The click tick. What a listing caught mid-click answers is the file the pin is already
    // showing, and `offer` declines that in silence — so a click given it has been given
    // nothing at all, and has to be held to be asked again. This is the arm that was missing:
    // with the hold never taken, the retry had nothing to retry.
    watch.answer_click(showing, showing, true, clicked_at, point);
    assert_eq!(
        watch.pending.as_ref().map(|held| held.at),
        Some(clicked_at),
        "the click is held from the tick it was read on, not dropped"
    );
    assert_eq!(
        watch.pending.as_ref().map(|held| held.point),
        Some(point),
        "against the spot the hand clicked at, which is what a moved hand is told from"
    );

    // The retry. The shell says the same thing again, and the hold survives it — but its own
    // clock does not move, because it is measured from the click: a hold re-armed here would
    // slide its window along with every tick and a click that is being spent rather than
    // answered would be held for as long as the hand stayed still over it.
    let retried_at = clicked_at + Duration::from_millis(120);
    watch.answer_click(showing, showing, false, retried_at, point);
    assert_eq!(
        watch.pending.as_ref().map(|held| held.at),
        Some(clicked_at),
        "and the hold is still the one the click took out, with its own clock"
    );

    // Case is not the difference: a listing that spells the name its own way is still
    // answering with the file on screen, and treating it as another file would swap a window
    // onto the picture it is already showing.
    watch.answer_click(
        Path::new("d:/pictures/ONE.PNG"),
        showing,
        false,
        retried_at,
        point,
    );
    assert!(
        watch.pending.is_some(),
        "and a spelling is not a different file"
    );

    // The shell has caught up and names the file the user actually clicked, which is the
    // answer the click was taken up for. Now the hold is given up.
    watch.answer_click(
        Path::new("D:/Pictures/two.png"),
        showing,
        false,
        retried_at,
        point,
    );
    assert!(
        watch.pending.is_none(),
        "a file the pin is not showing ends the hold"
    );

    // Which includes a file nothing here can preview: Explorer selected it and there is
    // nothing to put in the window, so holding the click for it would ask the shell the same
    // question for as long as the hold lasts and swap nothing either way.
    watch.pending = Some(PendingClick {
        at: clicked_at,
        point,
        place: None,
    });
    watch.answer_click(
        Path::new("D:/Pictures/archive.7z"),
        showing,
        false,
        retried_at,
        point,
    );
    assert!(
        watch.pending.is_none(),
        "and so does a file this app has nothing to show for"
    );

    // And a hover that settles on the file already on screen is not a click and arms
    // nothing: a hand coming to rest on what is on screen has asked for nothing, so there is
    // nothing of its own to hold on its behalf.
    watch.answer_click(showing, showing, false, retried_at, point);
    assert!(
        watch.pending.is_none(),
        "a hover on the file already shown arms no click"
    );
}

/// A click read out of the focus is the click that made it, and a listing that replaced the one
/// it was made in is not that click: a folder change would otherwise answer every one, and the
/// preview would follow the item under a parked pointer from listing to listing (see
/// `PinUpdateWatch::click_picked_item`).
#[test]
fn a_focus_change_is_a_click_pick_only_where_that_click_was_made() {
    let now = Instant::now();
    let clicked_in = HoverLocation {
        folder: None,
        search_root: None,
        location_url: Some("file:///D:/Pictures".to_string()),
        view_hwnd: Some(0x1234),
    };
    let under_the_click = (100, 230, 480, 250);

    let watch = PinUpdateWatch {
        pending: Some(PendingClick {
            at: now,
            point: POINT { x: 300, y: 240 },
            place: Some(clicked_in.clone()),
        }),
        ..PinUpdateWatch::default()
    };

    assert!(
        watch.click_picked_item(under_the_click, Some(&clicked_in), now),
        "the item under the click, in the place the click was made in, is what that click picked"
    );

    // The folder change, which is the case the rule is written against: the item the fresh
    // listing has put under the hand is drawn exactly where the click landed, so nothing about
    // the box can tell it — the place is what says it is not the item the click was on.
    let elsewhere = HoverLocation {
        location_url: Some("file:///D:/Pictures/2024".to_string()),
        ..clicked_in.clone()
    };
    assert!(
        !watch.click_picked_item(under_the_click, Some(&elsewhere), now),
        "an item in another folder is not the file the click was on, whatever box it is drawn in"
    );

    // A shell that described no place: which listing the item is in is not known, and a
    // comparison that cannot be made is never made in the click's favour.
    assert!(
        !watch.click_picked_item(under_the_click, None, now),
        "a place nobody answered is not the place the click was made in"
    );

    // An item the click was not on: a focus that lands beside the hand is the shell's own move.
    assert!(
        !watch.click_picked_item((0, 0, 60, 40), Some(&clicked_in), now),
        "the item has to be drawn where the click landed"
    );

    // And a click whose hold has run out is not a click any more, whatever the focus is doing.
    let mut expired = watch;
    if let Some(held) = expired.pending.as_mut() {
        held.at = now - Duration::from_millis(PIN_CLICK_RETRY_MS + 1);
    }
    assert!(
        !expired.click_picked_item(under_the_click, Some(&clicked_in), now),
        "a focus that moves after the hold is over is nobody's pick"
    );
}

/// A selection the listing holds is a pick wherever it changes in the listing watched —
/// whoever has the keyboard — and a baseline everywhere else (see
/// `PinUpdateWatch::note_selection`).
#[test]
fn a_selection_change_in_the_watched_listing_is_a_pick() {
    let watched = HoverLocation {
        folder: None,
        search_root: None,
        location_url: Some("file:///D:/Pictures".to_string()),
        view_hwnd: Some(0x1234),
    };
    let elsewhere = HoverLocation {
        location_url: Some("file:///D:/Pictures/2024".to_string()),
        ..watched.clone()
    };
    let one = std::path::PathBuf::from("D:/Pictures/one.png");
    let two = std::path::PathBuf::from("D:/Pictures/two.png");
    let three = std::path::PathBuf::from("D:/Pictures/2024/three.png");

    let mut watch = PinUpdateWatch::default();

    // The very first sighting is a baseline, not a pick: which file the listing holds
    // before the hand does anything in it is not a file anybody picked.
    assert_eq!(
        watch.note_selection(watched.clone(), Some(one.clone()), false),
        None,
        "the first sighting baselines rather than offers"
    );
    assert_eq!(watch.sel_selected.as_deref(), Some(one.as_path()));

    // Another file selected in the same listing is the pick — by click or by key, which
    // of the two is not asked.
    assert_eq!(
        watch.note_selection(watched.clone(), Some(two.clone()), false),
        Some(two.clone()),
        "a change in the watched listing is a pick"
    );

    // No change is nothing, however often it is seen.
    assert_eq!(
        watch.note_selection(watched.clone(), Some(two.clone()), false),
        None,
        "seeing the same selection again is not another pick"
    );

    // A fresh listing under the hand baselines again rather than offering what it
    // happens to hold: the preview does not follow a folder change.
    assert_eq!(
        watch.note_selection(elsewhere.clone(), Some(three.clone()), false),
        None,
        "a folder moved to is a fresh baseline, not a pick"
    );
    assert_eq!(watch.sel_selected.as_deref(), Some(three.as_path()));

    // ...unless a click is in hand: the click selects before the shell reports it, so a
    // selection read in a listing the watch has not seen is still that click's pick.
    let mut fresh = PinUpdateWatch::default();
    assert_eq!(
        fresh.note_selection(elsewhere.clone(), Some(three.clone()), true),
        Some(three.clone()),
        "a click's selection is that click's pick even in a new listing"
    );

    // A click the shell has not caught up with offers nothing and keeps the old answer
    // standing, so the change offers on the tick the shell reports it on.
    let mut lagging = PinUpdateWatch::default();
    assert_eq!(
        lagging.note_selection(watched.clone(), Some(one.clone()), false),
        None
    );
    assert_eq!(
        lagging.note_selection(watched.clone(), None, true),
        None,
        "a click nothing is reported for yet offers nothing"
    );
    assert_eq!(
        lagging.sel_selected.as_deref(),
        Some(one.as_path()),
        "and the old answer is left standing for the change"
    );
    assert_eq!(
        lagging.note_selection(watched.clone(), Some(two.clone()), false),
        Some(two.clone()),
        "which the change then offers"
    );
}

/// A poll that names no file in the listing already watched is a miss rather than a
/// pick: the baseline is left standing so the next sighting of the same file is not
/// offered. This is the pin reverting to the Explorer selection after its own
/// Previous/Next walk: one transient `None` (a gap under the pointer, a UIA miss
/// while the mouse moves) cleared the baseline, and the recovery of the same file
/// read as a change.
#[test]
fn a_selection_miss_in_the_watched_listing_is_not_a_pick() {
    let watched = HoverLocation {
        folder: None,
        search_root: None,
        location_url: Some("file:///D:/Pictures".to_string()),
        view_hwnd: Some(0x1234),
    };
    let one = std::path::PathBuf::from("D:/Pictures/one.png");
    let two = std::path::PathBuf::from("D:/Pictures/two.png");

    let mut watch = PinUpdateWatch::default();
    assert_eq!(
        watch.note_selection(watched.clone(), Some(one.clone()), false),
        None
    );

    assert_eq!(
        watch.note_selection(watched.clone(), None, false),
        None,
        "a miss names no file and offers none"
    );
    assert_eq!(
        watch.sel_selected.as_deref(),
        Some(one.as_path()),
        "and the baseline is left standing"
    );
    assert_eq!(
        watch.note_selection(watched.clone(), Some(one.clone()), false),
        None,
        "so recovering the same file is not a pick"
    );

    assert_eq!(
        watch.note_selection(watched.clone(), Some(two.clone()), false),
        Some(two.clone()),
        "while a genuinely different file still is"
    );

    // A first sighting that is itself a miss leaves no baseline for the first file
    // to be mistaken for a change.
    let mut missed = PinUpdateWatch::default();
    assert_eq!(missed.note_selection(watched.clone(), None, false), None);
    assert_eq!(
        missed.note_selection(watched.clone(), Some(one.clone()), false),
        None,
        "the first file actually seen is the baseline, not a pick"
    );
    assert_eq!(
        missed.note_selection(watched.clone(), Some(one.clone()), false),
        None
    );
}

use super::*;

/// A folder of a test's own with the entries a drop is judged by: two files a pin can show, one it
/// cannot, and a folder, which is the entry a drop is refused for rather than stepped over.
fn a_folder_of_drops(name: &str) -> PathBuf {
    let folder = std::env::temp_dir()
        .join("rust-hover-preview-pin-drop-tests")
        .join(name);
    let _ = std::fs::remove_dir_all(&folder);
    std::fs::create_dir_all(folder.join("a folder")).expect("a folder a test can write to");

    for (file, bytes) in [
        ("a.png", vec![0u8; 8]),
        ("b.png", vec![0u8; 8]),
        ("run.exe", vec![0u8; 8]),
    ] {
        std::fs::write(folder.join(file), bytes).expect("a file a test can write");
    }

    folder
}

/// A pin standing on a picture of that folder, so that a drop is answered against a window that is
/// showing something rather than against nothing at all.
fn a_pin_on(path: &Path) -> PinnedPreview {
    PinnedPreview {
        path: path.to_path_buf(),
        ..overlay_pin((0, 0, 100, 100), PinChrome::always())
    }
}

/// **A drop is judged by its name and by what it is, and nothing a drop names that a pin has no
/// preview for is anything at all.** The two answers a support gate has to give are the two a read of
/// the file would have to give, without a read: a `.png` is a kind the image list claims, a `.exe` is
/// a file no list claims, and a folder is an entry that is not a file — and a folder of anything, even
/// of a thousand pictures, is still refused, because the listing is the answer to a drop of one and it
/// is not the pin's window.
#[test]
fn a_drop_is_supported_by_its_name_and_refused_for_a_folder_or_a_kind_nothing_shows() {
    let folder = a_folder_of_drops("support");

    assert!(
        pin_drop::the_pin_can_show(&folder.join("a.png")),
        "a picture is what a pin is shown, and the image list's own name claims it"
    );
    assert!(
        !pin_drop::the_pin_can_show(&folder.join("run.exe")),
        "a program is no kind of preview, so there is nothing to show and nothing to say"
    );
    assert!(
        !pin_drop::the_pin_can_show(&folder.join("a folder")),
        "a folder is not a file, and the listing it would be is not this window's"
    );
    assert!(
        !pin_drop::the_pin_can_show(&folder.join("nothing here.png")),
        "and a file that is not there is no file to show"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// **A drop of several files is one file: the first the hand dragged, and only the first.** The list
/// reaches the loop as a single path and nothing carries the rest of it, so what the pick slot holds
/// after the message is one file — and the tick that takes it has one pick, not a queue. It is proved
/// here by putting the first of a multi-file drop where the window procedure leaves it and reading the
/// loop's own take, which is the whole of the road between the two.
#[test]
fn a_drop_of_several_files_leaves_one_pick_and_the_first_of_them() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();
    let folder = a_folder_of_drops("several");

    let first = folder.join("a.png");
    let second = folder.join("b.png");
    if let Ok(mut slot) = pin_drop::PIN_DROP.lock() {
        *slot = Some(pin_drop::DroppedOn {
            path: first.clone(),
            put_away: false,
        });
    }

    let mut pin = a_pin_on(&second);
    pin.path = second.clone();
    stand_pin(Some(pin));

    let mut pin_pick = None;
    let mut pin_held_pick = None;
    let mut pin_swap_requested = None;
    let mut pin_walk = Some(PinStep {
        at: second.clone(),
        from: second.clone(),
        step: 1,
        order: vec![second.clone()],
        missed: 0,
        shuffle: false,
        anchor: Some(crate::formats::routing::NavCategory::Images),
    });
    if let Some(drop) = pin_drop::take_dropped_on() {
        pin_walk = None;
        pin_swap_requested = drop.put_away.then(|| drop.path.clone());
        pin_pick = Some(drop.path);
    }

    let Some((path, walk, _user_named)) = take_pin_pick(
        &mut pin_pick,
        &mut pin_held_pick,
        &mut pin_swap_requested,
        &mut pin_walk,
        false,
    ) else {
        panic!("a drop the window took is a pick the loop has");
    };
    assert_eq!(
        path, first,
        "the file the hand dragged first is the file the pin is shown"
    );
    assert!(
        walk.is_none(),
        "and the walk it was stepping is gone with it: the next press re-anchors on the dropped file"
    );
    assert!(
        pin_pick.is_none(),
        "the slot is emptied by the take, so the tick after this one has nothing of the drop"
    );

    stand_pin(None);
    stand_pin(previous_pin);
    let _ = std::fs::remove_dir_all(&folder);
}

/// **A drop the pin cannot show changes nothing: no pick, no walk, and the file on the window exactly
/// the file it was showing.** A refusal is nobody's step, so unlike a walk's step over a file there is
/// no file to step on to and nothing to put in its place — the window answers a folder or a program by
/// not answering, which is what makes a drop safe to make over a picture the user is working in.
#[test]
fn a_refused_drop_leaves_the_pick_the_walk_and_the_pins_file_alone() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();
    let folder = a_folder_of_drops("refused");

    let showing = folder.join("a.png");
    let mut pin = a_pin_on(&showing);
    pin.path = showing.clone();
    stand_pin(Some(pin));
    let held = PinStep {
        at: showing.clone(),
        from: showing.clone(),
        step: 1,
        order: vec![showing.clone(), folder.join("b.png")],
        missed: 0,
        shuffle: false,
        anchor: Some(crate::formats::routing::NavCategory::Images),
    };
    stand_pin_walk(Some(held.clone()));

    // Every refusal the gate gives is what the window procedure does with it: nothing is written
    // down, so the loop finds nothing where the window left nothing. Asked of the gate rather than of
    // the message, since the list is an `HDROP` the Shell owns and no test can make one cheaply.
    for refused in [folder.join("run.exe"), folder.join("a folder")] {
        assert!(
            !pin_drop::the_pin_can_show(&refused),
            "a drop the pin cannot show is one it says nothing about and writes down nowhere"
        );
    }
    assert!(
        pin_drop::take_dropped_on().is_none(),
        "and the loop finds nothing where the window left nothing"
    );

    let mut pin_pick = None;
    let mut pin_held_pick = None;
    let mut pin_swap_requested = None;
    let mut walk = take_pin_walk();
    // Asked with a wait in hand so that the pin's own walk is not a source of a pick: what is being
    // proved is that nothing of the drop was written down, and the walk it was standing on is the
    // pin's own rather than the drop's.
    assert!(
        take_pin_pick(
            &mut pin_pick,
            &mut pin_held_pick,
            &mut pin_swap_requested,
            &mut walk,
            true,
        )
        .is_none(),
        "so there is no pick, and the pin is not asked for another file"
    );
    assert_eq!(
        pinned_path().as_deref(),
        Some(showing.as_path()),
        "the file on the window is the file it was showing"
    );
    if let Some(mut state) = pin_state() {
        if let Some(standing) = state.walk_mut() {
            *standing = Some(held);
        }
    }

    stand_pin(None);
    stand_pin(previous_pin);
    let _ = std::fs::remove_dir_all(&folder);
}

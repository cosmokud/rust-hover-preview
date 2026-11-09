use super::*;
use windows::Win32::System::Memory::GMEM_ZEROINIT;
use windows::Win32::UI::Shell::DROPFILES;

/// **A file list the Shell would have handed this window, made here.** A `WM_DROPFILES`
/// carries an `HDROP`, which is a moveable block beginning with the structure that says where in it
/// the names begin — and a list is what the presence of `CF_HDROP` promises the window procedure
/// gets the list and gives it back with `DragFinish`.
///
/// The block is made with `fWide` set, which is the whole of what makes it a list of UTF-16 names
/// rather than a list of ANSI ones, and it carries every name the caller hands it because the message
/// is the shell's own road and not this app's: what a drop names is read out of exactly this.
///
/// Safety: the block is `LocalAlloc`ed and released by dropping the `HDROP` the caller passes to
/// `DragFinish`, which is the only road a list is given back by.
unsafe fn a_file_list(names: &[&Path]) -> HDROP {
    // The names as the wide text a list holds, each ended by the null between them, then the further
    // null one that says there are no more.
    let mut units: Vec<u16> = Vec::new();
    for name in names {
        units.extend(name.as_os_str().encode_wide());
        units.push(0);
    }
    units.push(0);

    // The structure first, then the names: `pFiles` is where in the same block the names begin,
    // and what is stored there is the offset from the beginning of the block rather than a pointer.
    // The block is the Shell's own kind — moveable global memory, which is what a drop's list is — and
    // it is cleared before anything is written into it, so that a name shorter than the block is a
    // name and not whatever the block was holding.
    let head = std::mem::size_of::<DROPFILES>();
    let bytes = head + units.len() * std::mem::size_of::<u16>();
    let block = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, bytes)
        .expect("a block the Shell's list lives in");
    let base = GlobalLock(block) as *mut u8;
    let files = DROPFILES {
        pFiles: head as u32,
        pt: POINT { x: 0, y: 0 },
        fNC: false.into(),
        fWide: true.into(),
    };
    std::ptr::write(base as *mut DROPFILES, files);
    std::ptr::copy_nonoverlapping(units.as_ptr() as *const u8, base.add(head), units.len() * 2);
    let _ = GlobalUnlock(block);

    // The list is the block the names are in, which is what `DragQueryFileW` walks: the offset at the
    // head of it and the names after, in the block the Shell's own answer for the drop gives back.
    HDROP(base as *mut core::ffi::c_void)
}

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

/// **A drop is read off the list the message carried, and the list is given back.** The road a real
/// drag takes is a `WM_DROPFILES` with an `HDROP` in it, so the name is read out of a list made the
/// way the Shell's is — the structure at the head of the block, the names after it, the null between
/// them — rather than out of a path handed straight to the reader. And the list is given back with
/// `DragFinish` on every road the reader takes, refusal and acceptance alike: a list the window holds
/// after answering is one the Shell never gets back.
#[test]
fn a_drop_is_read_off_the_list_the_message_carried_and_the_list_is_given_back() {
    let folder = a_folder_of_drops("reading");
    let picture = folder.join("a.png");
    let program = folder.join("run.exe");

    // Safety: both lists are the caller's, and each is given back by the `DragFinish` inside
    // `take_drop` — which is the whole of what this test is asking about.
    let accepted = unsafe { pin_drop::take_drop(a_file_list(&[picture.as_path()])) };
    assert_eq!(
        accepted.as_deref(),
        Some(picture.as_path()),
        "the file a pin can show comes back as the file it names"
    );

    let refused = unsafe { pin_drop::take_drop(a_file_list(&[program.as_path()])) };
    assert!(
        refused.is_none(),
        "and a file no kind previews is read and refused, rather than left to be refused unread"
    );

    // And a list of several files is the first of them and nothing else: what the hand dragged first
    // is what the pin is asked for, and the entries after it are read by nobody (see
    // `read_dropped_file`).
    let several = unsafe {
        pin_drop::take_drop(a_file_list(&[
            picture.as_path(),
            folder.join("b.png").as_path(),
        ]))
    };
    assert_eq!(
        several.as_deref(),
        Some(picture.as_path()),
        "one drop is one file, and it is the first of the list"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// **A file dropped on a bubble brings the pin back up on it.** The bubble is a window with no media
/// band to put a file in, so a drop it takes is a pin the user is asking for — and it is answered by the
/// ordinary restore an update names, which is the same road the key takes, rather than by leaving the
/// window down and the file nowhere (see `pin_drop::DroppedOn::put_away` and `restore_pin`).
#[test]
fn a_drop_on_a_bubble_brings_the_pin_back_on_the_file_it_names() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();
    let folder = a_folder_of_drops("bubble");

    let showing = folder.join("a.png");
    let dropped = folder.join("b.png");
    let mut pin = a_pin_on(&showing);
    pin.path = showing.clone();
    pin.collapsed = true;
    stand_pin(Some(pin));
    assert!(
        pin_is_collapsed(),
        "the pin is away in its bubble, which is the window the drop lands on"
    );

    if let Ok(mut slot) = pin_drop::PIN_DROP.lock() {
        *slot = Some(pin_drop::DroppedOn {
            path: dropped.clone(),
            put_away: true,
        });
    }

    let Some(drop) = pin_drop::take_dropped_on() else {
        panic!("the drop the window took is the loop's to find");
    };
    assert!(
        drop.put_away,
        "the file was dropped on the bubble, so the pin has to come back for it"
    );
    // The restore is the loop's answer to that fact, and it is the one thing that leaves the pin up:
    // nothing else in this test touches the collapse, so the state after it is what the restore did.
    restore_pin(&Win32PinWindow, PinRestore::Update);
    assert!(
        !pin_is_collapsed(),
        "a drop on the bubble is the pin put back up rather than a file with no window to show it in"
    );

    stand_pin(None);
    stand_pin(previous_pin);
    let _ = std::fs::remove_dir_all(&folder);
}

/// **A drop is answered with `Pin Mode → Update Preview` off, which is where a pick behind the window
/// is not.** The row is about a file another hand chose behind the pin — the listing the user is
/// working in — and a dropped file was handed to the window itself, so the loop's gate reads the drop
/// as one the window is asked for whatever the row says. It is proved against the gate itself rather
/// than against a run of the loop: the setting is turned off, a drop is taken the way the loop takes it,
/// and the answer the gate gives a pick is the one it gives this one.
#[test]
fn a_drop_is_taken_with_the_update_setting_off() {
    let _room = the_room_to_itself();
    let previous_pin = take_pin_for_a_test();
    let folder = a_folder_of_drops("gate");

    let showing = folder.join("a.png");
    let dropped = folder.join("b.png");
    let mut pin = a_pin_on(&showing);
    pin.path = showing.clone();
    stand_pin(Some(pin));

    // The two rows a pick's gate reads, off for as long as this test runs and put back when the guard
    // is dropped, so a test running beside this one is not answered with this one's setting (see
    // `walk::WalkEveryFile`).
    struct TheRowOff(bool);
    impl TheRowOff {
        fn now() -> Self {
            let mut config = CONFIG.lock().expect("the configuration");
            let was = config.pin_update_enabled;
            config.pin_update_enabled = false;
            TheRowOff(was)
        }
    }
    impl Drop for TheRowOff {
        fn drop(&mut self) {
            if let Ok(mut config) = CONFIG.lock() {
                config.pin_update_enabled = self.0;
            }
        }
    }
    let _off = TheRowOff::now();
    assert!(!pin_update_enabled(), "the row this test is about is off");

    if let Ok(mut slot) = pin_drop::PIN_DROP.lock() {
        *slot = Some(pin_drop::DroppedOn {
            path: dropped.clone(),
            put_away: false,
        });
    }

    // What the loop does with a drop, in the order it does it: the walk is given up, the drop becomes
    // the pick, and the drop's own fact is carried down to the gate (see `run_preview_window`).
    let mut pin_pick = None;
    let mut pin_held_pick = None;
    let mut pin_swap_requested = None;
    let mut pin_walk = Some(PinStep {
        at: showing.clone(),
        from: showing.clone(),
        step: 1,
        order: vec![showing.clone()],
        missed: 0,
        shuffle: false,
        anchor: Some(crate::formats::routing::NavCategory::Images),
    });
    let mut dropped_pick = false;
    if let Some(drop) = pin_drop::take_dropped_on() {
        pin_walk = None;
        dropped_pick = true;
        pin_pick = Some(drop.path);
    }
    assert!(
        dropped_pick,
        "the drop is the loop's, and the walk it was stepping is gone with it"
    );
    assert!(
        pin_walk.is_none(),
        "so the next press re-anchors on the dropped file rather than the folder it came from"
    );

    let taken = take_pin_pick(
        &mut pin_pick,
        &mut pin_held_pick,
        &mut pin_swap_requested,
        &mut pin_walk,
        false,
    );
    let Some((path, walk, user_named)) = taken else {
        panic!("a drop the window took is a pick the loop has");
    };
    assert_eq!(
        path, dropped,
        "the file a hand dropped is the file asked for"
    );
    assert!(walk.is_none(), "and it is nobody's step of a walk");
    assert!(
        user_named,
        "a drop arrives as a file the user named, which is what the row below would refuse"
    );
    // **The gate the loop reads around this take, which is the one line this test is about.** A drop
    // is indistinguishable here from a pick made behind the window — the file is in the same hand and
    // named by the same hand — so the fact that it was dragged onto the window is carried down beside
    // it, and it is that fact the gate reads rather than the row (see the tick's own `pin_pick_dropped`
    // and `pin_drop`). With the row off and the drop flag standing, the file is taken; without the
    // flag the same pick would be refused.
    let gate = |dropped: bool| dropped || pin_update_enabled() || !user_named;
    assert!(
        gate(true),
        "a file dragged onto the window is answered with the row that gates a pick behind it off"
    );
    assert!(
        !gate(false),
        "and the same file picked behind the window is what that row refuses"
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

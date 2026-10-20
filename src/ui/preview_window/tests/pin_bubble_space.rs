use super::*;
use windows::core::w;
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DestroyWindow, GetWindow, IsWindowVisible, ShowWindow, GW_HWNDPREV,
    SW_SHOWNOACTIVATE, WINDOW_EX_STYLE, WS_POPUP,
};

/// A SPACE with a focused bubble holds a film or a sound, and nothing else.
#[test]
fn space_on_a_focused_bubble_holds_only_a_film_or_a_sound() {
    for kind in [MediaType::Video, MediaType::NativeVideo, MediaType::Audio] {
        assert!(
            bubble_space_toggles(Some(kind)),
            "a {kind:?} bubble has playback for SPACE to hold"
        );
    }

    for kind in [
        MediaType::StaticImage,
        MediaType::Text,
        MediaType::Pdf,
        MediaType::Office,
        MediaType::Unplayable,
    ] {
        assert!(
            !bubble_space_toggles(Some(kind)),
            "a {kind:?} bubble has no playback, so SPACE is not a hold"
        );
    }

    assert!(
        !bubble_space_toggles(None),
        "with nothing on screen SPACE is not a hold either"
    );
}

/// Only a click that moved nothing on a bubble the hand was already in puts the window back: the
/// focusing click stays a bubble, so SPACE has a window to hold.
#[test]
fn only_a_click_on_an_already_focused_bubble_puts_the_window_back() {
    assert!(
        bubble_click_restores(true, false),
        "a click on a focused bubble restores it"
    );
    assert!(
        !bubble_click_restores(false, false),
        "the first click only focuses, so the bubble stays up for SPACE"
    );
    assert!(
        !bubble_click_restores(true, true),
        "a drag carries the bubble rather than restoring it"
    );
    assert!(
        !bubble_click_restores(false, true),
        "a drag that also focused carries the bubble too"
    );
}

/// The bubble's menu holds the hold for a film or a sound, the two walks of the folder it
/// stands in, the switch that decides what a minimize holds back, the two places a minimize can
/// put the pin, and the two ways out of any bubble.
#[test]
fn the_bubble_menu_holds_only_what_has_playback() {
    for kind in [MediaType::Video, MediaType::NativeVideo, MediaType::Audio] {
        assert_eq!(
            bubble_menu_rows(Some(kind), false, true),
            vec![
                BubbleMenuRow::Toggle,
                BubbleMenuRow::NextFile,
                BubbleMenuRow::PreviousFile,
                BubbleMenuRow::Separator,
                BubbleMenuRow::PauseOnMinimize,
                BubbleMenuRow::Separator,
                BubbleMenuRow::MinimizeToBubble,
                BubbleMenuRow::MinimizeToTaskbar,
                BubbleMenuRow::Separator,
                BubbleMenuRow::ToTaskbar,
                BubbleMenuRow::Close,
            ],
            "a {kind:?} bubble offers the hold first"
        );
    }

    for kind in [Some(MediaType::StaticImage), Some(MediaType::Text), None] {
        assert_eq!(
            bubble_menu_rows(kind, false, true),
            vec![
                BubbleMenuRow::NextFile,
                BubbleMenuRow::PreviousFile,
                BubbleMenuRow::Separator,
                BubbleMenuRow::MinimizeToBubble,
                BubbleMenuRow::MinimizeToTaskbar,
                BubbleMenuRow::Separator,
                BubbleMenuRow::ToTaskbar,
                BubbleMenuRow::Close,
            ],
            "a {kind:?} bubble offers no hold, only the walk, the two places and the two ways out"
        );
    }
}

/// The menu names the hold for what it will do, the walk for the file it steps to, and the way
/// out for where it goes: Pause while playing and Play while held, and the taskbar spelled the
/// way the tray spells it.
#[test]
fn the_bubble_menu_names_what_each_row_will_do() {
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::Maximize, false, false),
        "Maximize",
        "the caption's maximize is named for what it will do"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::Maximize, false, true),
        "Restore",
        "and for the other way it will do it once the window is maximized"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::Toggle, true, false),
        "Pause",
        "a hold on playing media pauses it"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::Toggle, false, false),
        "Play",
        "a hold on held media lets it go"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::NextFile, true, false),
        "Next File",
        "the walk forward is the file after this one"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::PreviousFile, false, false),
        "Previous File",
        "and the walk back the one before it"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::Separator, true, false),
        "",
        "a rule carries nothing: there is nothing on one to read"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::PauseOnMinimize, true, false),
        "Pause on Minimize",
        "the switch is spelled the way the tray spells it"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::MinimizeToBubble, true, false),
        "Minimize to Bubble",
        "the place a minimize can put the pin is named for it"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::MinimizeToTaskbar, false, false),
        "Minimize to Taskbar",
        "and so is the other"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::ToBubble, true, false),
        "To Bubble",
        "the bubble a pin still standing is put away into is named for it, the way the taskbar is"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::ToTaskbar, true, false),
        "To Taskbar",
        "the stand-in is the taskbar's, not the tray's"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::Close, false, false),
        "Close",
        "the way out is a close"
    );
}

/// The painted menu carries the same rows in the same order, labelled for the playback behind
/// them — the panel the chrome paints cannot disagree with the rows the press answers.
#[test]
fn the_painted_menu_carries_the_rows_the_press_answers() {
    let painted: Vec<String> =
        bubble_menu_chrome_rows(Some(MediaType::Audio), true, false, true, false)
            .iter()
            .map(|row| row.label.clone())
            .collect();
    assert_eq!(
        painted,
        vec![
            "Pause",
            "Next File",
            "Previous File",
            "",
            "Pause on Minimize",
            "",
            "Minimize to Bubble",
            "Minimize to Taskbar",
            "",
            "To Taskbar",
            "Close"
        ],
        "a playing sound's menu pauses first"
    );

    let painted: Vec<String> =
        bubble_menu_chrome_rows(Some(MediaType::Video), false, false, true, false)
            .iter()
            .map(|row| row.label.clone())
            .collect();
    assert_eq!(
        painted,
        vec![
            "Play",
            "Next File",
            "Previous File",
            "",
            "Pause on Minimize",
            "",
            "Minimize to Bubble",
            "Minimize to Taskbar",
            "",
            "To Taskbar",
            "Close"
        ],
        "a held film's menu lets it go"
    );

    let painted: Vec<String> =
        bubble_menu_chrome_rows(Some(MediaType::Text), true, false, true, false)
            .iter()
            .map(|row| row.label.clone())
            .collect();
    assert_eq!(
        painted,
        vec![
            "Next File",
            "Previous File",
            "",
            "Minimize to Bubble",
            "Minimize to Taskbar",
            "",
            "To Taskbar",
            "Close"
        ],
        "a page's menu offers only the walk, the two places and the two ways out"
    );
}

/// The bubble's menu carries no flyout arrow: the shared placement points one wherever a Seek
/// row could open, and there is no Seek row here — so Close, the last row, stands unmarked.
#[test]
fn the_bubble_menu_points_no_arrow_at_any_row() {
    let rows = bubble_menu_chrome_rows(Some(MediaType::Audio), true, false, true, false);
    let popup = bubble_menu_popup((0, 44), 500, 500, 96, &rows);
    assert_eq!(
        popup.arrow,
        pin_chrome::MenuArrow::None,
        "no row of the bubble's menu opens anything"
    );
}

/// The two places a minimize can put a pin are one group rather than two switches: exactly one of
/// them carries the disc, and which one it is is the setting the tray's own `Pin Mode → Minimize`
/// writes — so a menu opened over a pick that has just been made shows it (see `PinMinimizeTo`).
#[test]
fn exactly_one_of_the_two_minimize_picks_carries_the_disc() {
    let _target = PinMinimizeToPutBack::put(PinMinimizeTo::Taskbar);

    let rows = bubble_menu_chrome_rows(Some(MediaType::Text), false, false, true, false);
    let marked: Vec<&str> = rows
        .iter()
        .filter(|row| matches!(row.mark, pin_chrome::MenuMark::Bullet))
        .map(|row| row.label.as_str())
        .collect();
    assert_eq!(
        marked,
        ["Minimize to Taskbar"],
        "the place a minimize puts the pin is the one the disc is on, and it is alone"
    );

    let _target = PinMinimizeToPutBack::put(PinMinimizeTo::Bubble);
    let rows = bubble_menu_chrome_rows(Some(MediaType::Text), false, false, true, false);
    let marked: Vec<&str> = rows
        .iter()
        .filter(|row| matches!(row.mark, pin_chrome::MenuMark::Bullet))
        .map(|row| row.label.as_str())
        .collect();
    assert_eq!(
        marked,
        ["Minimize to Bubble"],
        "and the other pick has it where the setting names the other place"
    );
}

/// The check the `Pause on Minimize` row carries is the switch of the kind the bubble stands
/// for: a film reads the video's and a sound reads the audio's, and a kind with no playback
/// carries no row at all.
#[test]
fn the_pause_switch_carries_the_check_of_the_kind_it_stands_for() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _switches = PausePreviewSettings::stood_at(false, true);

    let film = bubble_menu_chrome_rows(Some(MediaType::Video), false, false, true, false);
    let row = film
        .iter()
        .position(|row| row.label == "Pause on Minimize")
        .expect("a film's bubble carries the switch");
    assert!(
        matches!(film[row].mark, pin_chrome::MenuMark::Check(false)),
        "the video's switch is off, so the film's row is unchecked"
    );

    let sound = bubble_menu_chrome_rows(Some(MediaType::Audio), false, false, true, false);
    let row = sound
        .iter()
        .position(|row| row.label == "Pause on Minimize")
        .expect("a sound's bubble carries the switch");
    assert!(
        matches!(sound[row].mark, pin_chrome::MenuMark::Check(true)),
        "the audio's switch is on, so the sound's row is checked"
    );

    let page = bubble_menu_chrome_rows(Some(MediaType::Text), false, false, true, false);
    assert!(
        page.iter().all(|row| row.label != "Pause on Minimize"),
        "a page's bubble has no playback for a minimize to hold"
    );
}

/// A pin put back up out of a bubble the hand never carried comes back at its own box: the same
/// place and the same size it went down at.
///
/// The fault this is a regression test for: every restore placed the window beside the bubble, and
/// the bubble stands on the minimize button of the window the restore has just brought back — so a
/// pin cycled through the bubble with nothing moved walked sideways by most of its own width on
/// every click, shrank as it ran out of room, and ended up against the edge of the screen.
#[test]
fn a_pin_restored_from_a_bubble_that_was_never_carried_comes_back_at_its_own_box() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_bubble = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    let previous_carried = PIN_BUBBLE_CARRIED.load(Ordering::Acquire);

    // A bubble standing where the pin's own box is not: a restore that placed the window beside it
    // would have to leave a different box behind.
    // Safety: a window of this test's own, created and destroyed under the one-at-a-time lock.
    let bubble = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            700,
            500,
            120,
            120,
            None,
            None,
            None,
            None,
        )
        .expect("a window for the bubble")
    };
    PIN_BUBBLE_HWND.store(bubble.0 as isize, Ordering::SeqCst);

    let left_at = (100, 80, 420, 320);
    let mut pin = overlay_pin(left_at, PinChrome::always());
    pin.collapsed = true;
    stand_pin(Some(pin));
    PIN_BUBBLE_CARRIED.store(false, Ordering::Release);

    restore_pin(&a_restored_window(), PinRestore::Pressed);

    let restored = pin_state().and_then(|pinned| {
        pinned
            .pin()
            .map(|pin| (pin.collapsed, pin.content, pin.window_size()))
    });
    assert_eq!(
        restored,
        Some((false, left_at, (320, 240))),
        "a window placed beside the bubble would be a different box, and a pin put back up is the \
         window that went down"
    );

    // Safety: the window is this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(bubble);
    }
    PIN_BUBBLE_HWND.store(previous_bubble, Ordering::SeqCst);
    PIN_BUBBLE_CARRIED.store(previous_carried, Ordering::Release);
    stand_pin(previous_pin);
}

/// Whether `above` stands over `below` in the z-order right now: an ask of the desktop rather than a
/// reading of what this app meant to ask for.
///
/// Safety: both handles are windows of this test's or of this app's own, and the question is a
/// measurement of them rather than a change.
unsafe fn stands_over(above: HWND, below: HWND) -> bool {
    let mut walk = GetWindow(below, GW_HWNDPREV);
    while let Ok(window) = walk {
        if window == above {
            return true;
        }
        if window.is_invalid() {
            break;
        }
        walk = GetWindow(window, GW_HWNDPREV);
    }

    false
}

/// The offset this file's windows stand at: no display is there, so a run of the suite is no window
/// of the user's and no place the hand can land in.
const NOWHERE: i32 = -32000;

/// A window of the test's own, opened where no display is.
///
/// Safety: the window is the caller's to destroy.
unsafe fn a_window_of_the_tests_own() -> HWND {
    CreateWindowExW(
        WINDOW_EX_STYLE::default(),
        w!("STATIC"),
        w!(""),
        WS_POPUP,
        NOWHERE,
        NOWHERE,
        320,
        240,
        None,
        None,
        None,
        None,
    )
    .expect("a window of the test's own")
}

/// The pin put away into its bubble, with its own window on the desktop and a window of the test's
/// opened over it — the arrangement a restore is asked in, with `Pin Mode → Always On Top` off.
///
/// The second window is made after the pin's, which is what puts it in front, and it is still there
/// after the restore, so that a window which came back where a hidden one sits is a window that came
/// back underneath it.
///
/// The caller holds the pin tests' lock and puts back every global written here: the handle the pin
/// is drawn by, the bubble's, the one saying a bubble was carried, the setting, and the pin itself.
///
/// Safety: both windows are the caller's to destroy.
unsafe fn a_pin_put_away_into_its_bubble() -> (HWND, HWND) {
    let pin_hwnd = a_window_of_the_tests_own();
    PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);
    let in_front = a_window_of_the_tests_own();
    let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
    let _ = ShowWindow(in_front, SW_SHOWNOACTIVATE);
    assert!(
        stands_over(in_front, pin_hwnd),
        "the window opened after the pin's is the one in front of it"
    );

    // The setting the user's own report is under: a pin that is not Always On Top.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
    }

    // A pin put away into its bubble: collapsed, its own window gone from the screen, and no stand-in
    // for it but the bubble this test holds no window for.
    PIN_BUBBLE_HWND.store(0, Ordering::SeqCst);
    PIN_BUBBLE_CARRIED.store(false, Ordering::Release);
    let mut pin = overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    );
    pin.collapsed = true;
    stand_pin(Some(pin));
    // What the collapse into a bubble leaves: the window off the screen, and nothing standing in for
    // it but the bubble.
    let _ = ShowWindow(pin_hwnd, SW_HIDE);
    assert!(
        !IsWindowVisible(pin_hwnd).as_bool(),
        "a pin put away into its bubble has no window on the screen"
    );

    (pin_hwnd, in_front)
}

/// A pin brought back out of its bubble comes back in front of the windows that were opened over it,
/// and takes the keyboard with it, with `Pin Mode → Always On Top` off.
///
/// The fault this is a regression test for: a click on the bubble brought the window back *behind*
/// whatever the user had been working in. What a restore asked of the window was a place — `HWND_TOP`
/// for one that was off the screen, and `HWND_NOTOPMOST`, which asks for the place a window already
/// has, for one that was not — and a place is all it got: an ordinary window that is raised is a
/// window the next activation is free to stand over, and the window the hide left in front is the one
/// Windows activated. A press on the bubble is the one thing that says which window the user means, so
/// the window the restore answers it with is the one that comes to the front and the one that holds
/// the caret, the way a press on the window itself leaves it (see `restore_pin`, `pin_take_focus`).
#[test]
fn a_pin_brought_back_by_a_press_on_its_bubble_comes_back_in_front_and_takes_the_keyboard() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_bubble = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    let previous_carried = PIN_BUBBLE_CARRIED.load(Ordering::Acquire);
    let before = pin_always_on_top();

    // Safety: the windows are this test's own, made and destroyed under the one-at-a-time lock, and
    // the pin is stood up under it too.
    let (pin_hwnd, in_front) = unsafe { a_pin_put_away_into_its_bubble() };

    // What is asked of the pin's window is recorded rather than done: a window put in front of the
    // user is not something a test may do to their desktop. The window the desktop answers for is the
    // one this test made, because the restore asks about the handle the app holds.
    let window = RecordedPinWindow::new(pin_hwnd.0 as isize);
    restore_pin(&window, PinRestore::Pressed);

    assert!(
        !pin_is_collapsed(),
        "the restore is what puts the pin back up"
    );
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "a pin the bubble brought back is put above the windows the user had in front of it: what a \
         hide leaves behind is a place at the bottom of the z-order, and a showing that asks for the \
         place a window already has is a pin that comes back behind everything"
    );
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::SetFocusable { focusable: true },
            PinWindowCall::SetForeground,
            PinWindowCall::SetFocus,
        ],
        "the window the press brought back is made focusable again — the collapse made it \
         non-activating — and is then brought to the front and given the caret"
    );

    // Safety: both windows are this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(in_front);
        let _ = DestroyWindow(pin_hwnd);
        close_pin_taskbar();
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_BUBBLE_HWND.store(previous_bubble, Ordering::SeqCst);
    PIN_BUBBLE_CARRIED.store(previous_carried, Ordering::Release);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before;
    }
    stand_pin(previous_pin);
}

/// A pin brought back by its key while the listing holds the keyboard leaves the keyboard where the
/// user is working: only a press on this app's own stand-in asks for the front and the caret.
///
/// The key is pressed over the listing the file is picked in, and the file is what it asks for — a
/// window that took the caret with it would take the listing the user was stepping through with it
/// (see `PinRestore::Key`).
#[test]
fn a_pin_brought_back_by_its_key_leaves_the_keyboard_in_the_listing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_bubble = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    let previous_carried = PIN_BUBBLE_CARRIED.load(Ordering::Acquire);
    let before = pin_always_on_top();

    // Safety: the windows are this test's own, made and destroyed under the one-at-a-time lock, and
    // the pin is stood up under it too.
    let (pin_hwnd, in_front) = unsafe { a_pin_put_away_into_its_bubble() };

    let window = RecordedPinWindow::new(pin_hwnd.0 as isize);
    restore_pin(&window, PinRestore::Key);

    assert!(!pin_is_collapsed(), "the key is what puts the pin back up");
    assert_eq!(
        window.calls(),
        Vec::new(),
        "a pin brought back by its key asks nothing of the window: the listing keeps the keyboard"
    );

    // Safety: both windows are this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(in_front);
        let _ = DestroyWindow(pin_hwnd);
        close_pin_taskbar();
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    PIN_BUBBLE_HWND.store(previous_bubble, Ordering::SeqCst);
    PIN_BUBBLE_CARRIED.store(previous_carried, Ordering::Release);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = before;
    }
    stand_pin(previous_pin);
}

/// A preview update brings a pin in front of whatever was standing over it, without taking the
/// keyboard; with `Pin Mode → Update Preview → Move To Top` off it leaves the pin where the user's
/// own clicks put it.
///
/// The z-order is read off the desktop's own list because that is the one part of a raise nothing
/// records: what the call was asked for is a place among other windows, and a window put behind the
/// listing it is following is the whole of the fault this row answers (see `raise_pinned_preview`).
#[test]
fn a_preview_update_brings_a_pin_in_front_and_leaves_it_when_the_row_is_off() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);

    // Safety: both windows are this test's own, opened where no display is. The second is made
    // after the pin's, which is what puts it in front, and it is still there after each raise, so
    // that a pin which did not come forward is a pin still under it.
    let (pin_hwnd, in_front) = unsafe {
        let pin_hwnd = a_window_of_the_tests_own();
        PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);
        let in_front = a_window_of_the_tests_own();
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(in_front, SW_SHOWNOACTIVATE);
        assert!(
            stands_over(in_front, pin_hwnd),
            "the window opened after the pin's is the one in front of it"
        );
        (pin_hwnd, in_front)
    };

    // A pin that is up and on the screen, with `Always On Top` off: an ordinary window, which is
    // the one that can be standing behind the listing the user picks a file in.
    stand_pin(Some(overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    )));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }

    // Safety: the raise under test, asked of a window this test made.
    unsafe { raise_pinned_preview(pin_hwnd) };
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "an update puts the pin above the window that was standing over it: the file the user \
         picked is the one they are looking at next"
    );
    // Safety: the raise left `Always On Top` as it found it, which is the half of the raise that is
    // not about the window's place: the band it ends in is the one the setting names.
    let raised_style = unsafe { GetWindowLongPtrW(pin_hwnd, GWL_EXSTYLE) } as u32;
    assert_eq!(
        raised_style & WS_EX_TOPMOST.0,
        0,
        "a pin that is not Always On Top is raised through the topmost band rather than into it, \
         and left behind the windows that really are topmost"
    );
    assert_ne!(
        unsafe { GetForegroundWindow() },
        pin_hwnd,
        "the pin is raised without taking the keyboard: the pick it answers may be one made with \
         the arrows in the listing, and a raise that took the focus would end that walk"
    );

    // The row switched off: the window in front is put back there, and the update leaves it alone.
    unsafe {
        let _ = SetWindowPos(
            in_front,
            HWND_TOP,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    };
    assert!(
        unsafe { stands_over(in_front, pin_hwnd) },
        "the window in front is the one the setting off leaves the pin behind"
    );
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_update_move_to_top = false;
    }

    // Safety: the raise under test, asked of a window this test made.
    unsafe { raise_pinned_preview(pin_hwnd) };
    assert!(
        unsafe { stands_over(in_front, pin_hwnd) },
        "with the row off a preview update is a file swapped into the pin where it stands, and the \
         window stays where the user's own clicks put it"
    );
    // Safety: both windows are this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(in_front);
        let _ = DestroyWindow(pin_hwnd);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// A press that names no file gives the front back: the pin a pick raised is put behind the window
/// the press was made in, which is what a click on the tree, the toolbar or the gap between the
/// rows is answered with (see `lower_pinned_preview`), and a pin that is Always On Top keeps the
/// band it is in.
///
/// The z-order is read off the desktop's own list for the same reason the raise test reads it
/// there: what the call was asked for is a place among other windows, and a pin left standing over
/// the listing it was clicked in is the fault this answers.
#[test]
fn a_press_that_picks_no_file_puts_the_pin_back_behind_the_listing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);

    // Safety: both windows are this test's own, opened where no display is. The second is made
    // after the pin's, which is what puts it in front — the listing, in the arrangement the press
    // is read in.
    let (pin_hwnd, in_front) = unsafe {
        let pin_hwnd = a_window_of_the_tests_own();
        PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);
        let in_front = a_window_of_the_tests_own();
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(in_front, SW_SHOWNOACTIVATE);
        (pin_hwnd, in_front)
    };

    // A pin that is up and on the screen, with `Always On Top` off.
    stand_pin(Some(overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    )));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }

    // The pick raises the pin first, as an update does, and the press that follows names no file.
    unsafe { raise_pinned_preview(pin_hwnd) };
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "an update puts the pin above the listing the file was picked in"
    );
    unsafe { lower_pinned_preview(pin_hwnd, in_front) };
    assert!(
        unsafe { stands_over(in_front, pin_hwnd) },
        "the pin goes back behind the window the press was made in, which is what a click that \
         picked no file is answered with"
    );
    // Safety: the lower left `Always On Top` as it found it, which is the half of it that is not
    // about the window's place.
    let lowered_style = unsafe { GetWindowLongPtrW(pin_hwnd, GWL_EXSTYLE) } as u32;
    assert_eq!(
        lowered_style & WS_EX_TOPMOST.0,
        0,
        "a pin put back is still in the ordinary band, not the topmost one"
    );

    // The next pick brings it forward again, so a press that named no file is not a pin left
    // behind for good.
    unsafe { raise_pinned_preview(pin_hwnd) };
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "a later pick raises the pin again"
    );

    // `Always On Top` keeps its own band: a pin there is not the listing's to put behind.
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
    }
    unsafe { raise_pinned_preview(pin_hwnd) };
    let topmost_style = unsafe { GetWindowLongPtrW(pin_hwnd, GWL_EXSTYLE) } as u32;
    assert_ne!(
        topmost_style & WS_EX_TOPMOST.0,
        0,
        "a pin that is Always On Top is raised into the topmost band"
    );
    unsafe { lower_pinned_preview(pin_hwnd, in_front) };
    assert!(
        unsafe { stands_over(pin_hwnd, in_front) },
        "a pin that is Always On Top is left where it stands"
    );
    let still_topmost = unsafe { GetWindowLongPtrW(pin_hwnd, GWL_EXSTYLE) } as u32;
    assert_ne!(
        still_topmost & WS_EX_TOPMOST.0,
        0,
        "and a press that picks no file does not take its band away from it"
    );

    // Safety: both windows are this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(in_front);
        let _ = DestroyWindow(pin_hwnd);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// The raise a pick makes is asked for again on the ticks right after it, because the click that
/// made the pick reaches the listing's window after it reaches this app: the listing is activated
/// then, which is that window being put back over the pin. A pin that has been covered within the
/// grace is raised again, and one that is covered after it is not.
///
/// The window the user is working in is handed in rather than read, so the answer is asked of two
/// windows this test made (see `keep_pin_at_the_front`).
#[test]
fn a_pin_raised_for_a_pick_answers_for_the_front_while_the_grace_lasts() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);

    // Safety: both windows are this test's own, opened where no display is.
    let (pin_hwnd, in_front) = unsafe {
        let pin_hwnd = a_window_of_the_tests_own();
        PREVIEW_HWND.store(pin_hwnd.0 as isize, Ordering::SeqCst);
        let in_front = a_window_of_the_tests_own();
        let _ = ShowWindow(pin_hwnd, SW_SHOWNOACTIVATE);
        let _ = ShowWindow(in_front, SW_SHOWNOACTIVATE);
        (pin_hwnd, in_front)
    };

    stand_pin(Some(overlay_pin(
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240),
        PinChrome::always(),
    )));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
        config.pin_update_move_to_top = true;
    }

    // A pick raises the pin, and the click that made it then activates the listing — the window
    // this test made, put back where the click puts it.
    unsafe {
        raise_pinned_preview(pin_hwnd);
        let _ = SetWindowPos(
            in_front,
            HWND_TOP,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }
    assert!(
        !unsafe { pin_stands_over(pin_hwnd, in_front) },
        "the listing's window is activated after the pick, which is it standing over the pin"
    );

    // Safety: the grace under test, asked of windows this test made.
    let mut deadline = Some(Instant::now() + PIN_RAISE_GRACE);
    unsafe { keep_pin_at_the_front(pin_hwnd, in_front, &mut deadline, Instant::now()) };
    assert!(
        unsafe { pin_stands_over(pin_hwnd, in_front) },
        "while the grace lasts a pin covered by the click that picked its file is raised again"
    );

    // The same covering, now that the grace has run out: the pin is left where the user's own
    // clicks put it.
    unsafe {
        let _ = SetWindowPos(
            in_front,
            HWND_TOP,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }
    let mut deadline = Some(Instant::now());
    unsafe { keep_pin_at_the_front(pin_hwnd, in_front, &mut deadline, Instant::now()) };
    assert!(
        unsafe { stands_over(in_front, pin_hwnd) },
        "past the grace an update no longer answers for the front: the window stays where it is"
    );
    assert_eq!(
        deadline, None,
        "a grace that has run out is dropped rather than asked about on every tick after it"
    );

    // Safety: both windows are this test's own, and the handles the app holds are put back as they
    // were.
    unsafe {
        let _ = DestroyWindow(in_front);
        let _ = DestroyWindow(pin_hwnd);
    }
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

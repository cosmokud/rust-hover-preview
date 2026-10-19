use super::*;
use windows::core::w;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DestroyWindow, WINDOW_EX_STYLE, WS_POPUP,
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
            bubble_menu_rows(Some(kind), false),
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
            bubble_menu_rows(kind, false),
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
        bubble_menu_row_label(BubbleMenuRow::Toggle, true),
        "Pause",
        "a hold on playing media pauses it"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::Toggle, false),
        "Play",
        "a hold on held media lets it go"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::NextFile, true),
        "Next File",
        "the walk forward is the file after this one"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::PreviousFile, false),
        "Previous File",
        "and the walk back the one before it"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::Separator, true),
        "",
        "a rule carries nothing: there is nothing on one to read"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::PauseOnMinimize, true),
        "Pause on Minimize",
        "the switch is spelled the way the tray spells it"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::MinimizeToBubble, true),
        "Minimize to Bubble",
        "the place a minimize can put the pin is named for it"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::MinimizeToTaskbar, false),
        "Minimize to Taskbar",
        "and so is the other"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::ToBubble, true),
        "To Bubble",
        "the bubble a pin still standing is put away into is named for it, the way the taskbar is"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::ToTaskbar, true),
        "To Taskbar",
        "the stand-in is the taskbar's, not the tray's"
    );
    assert_eq!(
        bubble_menu_row_label(BubbleMenuRow::Close, false),
        "Close",
        "the way out is a close"
    );
}

/// The painted menu carries the same rows in the same order, labelled for the playback behind
/// them — the panel the chrome paints cannot disagree with the rows the press answers.
#[test]
fn the_painted_menu_carries_the_rows_the_press_answers() {
    let painted: Vec<String> = bubble_menu_chrome_rows(Some(MediaType::Audio), true, false)
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

    let painted: Vec<String> = bubble_menu_chrome_rows(Some(MediaType::Video), false, false)
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

    let painted: Vec<String> = bubble_menu_chrome_rows(Some(MediaType::Text), true, false)
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
    let rows = bubble_menu_chrome_rows(Some(MediaType::Audio), true, false);
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

    let rows = bubble_menu_chrome_rows(Some(MediaType::Text), false, false);
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
    let rows = bubble_menu_chrome_rows(Some(MediaType::Text), false, false);
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

    let film = bubble_menu_chrome_rows(Some(MediaType::Video), false, false);
    let row = film
        .iter()
        .position(|row| row.label == "Pause on Minimize")
        .expect("a film's bubble carries the switch");
    assert!(
        matches!(film[row].mark, pin_chrome::MenuMark::Check(false)),
        "the video's switch is off, so the film's row is unchecked"
    );

    let sound = bubble_menu_chrome_rows(Some(MediaType::Audio), false, false);
    let row = sound
        .iter()
        .position(|row| row.label == "Pause on Minimize")
        .expect("a sound's bubble carries the switch");
    assert!(
        matches!(sound[row].mark, pin_chrome::MenuMark::Check(true)),
        "the audio's switch is on, so the sound's row is checked"
    );

    let page = bubble_menu_chrome_rows(Some(MediaType::Text), false, false);
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

    restore_pin();

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

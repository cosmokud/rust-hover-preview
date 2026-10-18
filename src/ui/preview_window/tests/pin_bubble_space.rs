use super::*;

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
/// stands in, the switch that decides what a minimize holds back, and the two ways out of any
/// bubble.
#[test]
fn the_bubble_menu_holds_only_what_has_playback() {
    for kind in [MediaType::Video, MediaType::NativeVideo, MediaType::Audio] {
        assert_eq!(
            bubble_menu_rows(Some(kind)),
            vec![
                BubbleMenuRow::Toggle,
                BubbleMenuRow::NextFile,
                BubbleMenuRow::PreviousFile,
                BubbleMenuRow::Separator,
                BubbleMenuRow::PauseOnMinimize,
                BubbleMenuRow::Separator,
                BubbleMenuRow::ToTaskbar,
                BubbleMenuRow::Close,
            ],
            "a {kind:?} bubble offers the hold first"
        );
    }

    for kind in [Some(MediaType::StaticImage), Some(MediaType::Text), None] {
        assert_eq!(
            bubble_menu_rows(kind),
            vec![
                BubbleMenuRow::NextFile,
                BubbleMenuRow::PreviousFile,
                BubbleMenuRow::Separator,
                BubbleMenuRow::ToTaskbar,
                BubbleMenuRow::Close,
            ],
            "a {kind:?} bubble offers no hold, only the walk and the two ways out"
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
    let painted: Vec<String> = bubble_menu_chrome_rows(Some(MediaType::Audio), true)
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
            "To Taskbar",
            "Close"
        ],
        "a playing sound's menu pauses first"
    );

    let painted: Vec<String> = bubble_menu_chrome_rows(Some(MediaType::Video), false)
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
            "To Taskbar",
            "Close"
        ],
        "a held film's menu lets it go"
    );

    let painted: Vec<String> = bubble_menu_chrome_rows(Some(MediaType::Text), true)
        .iter()
        .map(|row| row.label.clone())
        .collect();
    assert_eq!(
        painted,
        vec!["Next File", "Previous File", "", "To Taskbar", "Close"],
        "a page's menu offers only the walk and the two ways out"
    );
}

/// The bubble's menu carries no flyout arrow: the shared placement points one wherever a Seek
/// row could open, and there is no Seek row here — so Close, the last row, stands unmarked.
#[test]
fn the_bubble_menu_points_no_arrow_at_any_row() {
    let rows = bubble_menu_chrome_rows(Some(MediaType::Audio), true);
    let popup = bubble_menu_popup((0, 44), 500, 500, 96, &rows);
    assert_eq!(
        popup.arrow,
        pin_chrome::MenuArrow::None,
        "no row of the bubble's menu opens anything"
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

    let film = bubble_menu_chrome_rows(Some(MediaType::Video), false);
    let row = film
        .iter()
        .position(|row| row.label == "Pause on Minimize")
        .expect("a film's bubble carries the switch");
    assert!(
        matches!(film[row].mark, pin_chrome::MenuMark::Check(false)),
        "the video's switch is off, so the film's row is unchecked"
    );

    let sound = bubble_menu_chrome_rows(Some(MediaType::Audio), false);
    let row = sound
        .iter()
        .position(|row| row.label == "Pause on Minimize")
        .expect("a sound's bubble carries the switch");
    assert!(
        matches!(sound[row].mark, pin_chrome::MenuMark::Check(true)),
        "the audio's switch is on, so the sound's row is checked"
    );

    let page = bubble_menu_chrome_rows(Some(MediaType::Text), false);
    assert!(
        page.iter().all(|row| row.label != "Pause on Minimize"),
        "a page's bubble has no playback for a minimize to hold"
    );
}

use super::*;

/// A SPACE with a focused bubble holds a film or a sound, and nothing else.
#[test]
fn space_on_a_focused_bubble_holds_only_a_film_or_a_sound() {
    for kind in [
        MediaType::Video,
        MediaType::NativeVideo,
        MediaType::Audio,
    ] {
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

/// The bubble's menu holds the hold for a film or a sound, and the two ways out of any bubble.
#[test]
fn the_bubble_menu_holds_only_what_has_playback() {
    for kind in [
        MediaType::Video,
        MediaType::NativeVideo,
        MediaType::Audio,
    ] {
        assert_eq!(
            bubble_menu_rows(Some(kind)),
            vec![
                BubbleMenuRow::Toggle,
                BubbleMenuRow::ToTaskbar,
                BubbleMenuRow::Close,
            ],
            "a {kind:?} bubble offers the hold first"
        );
    }

    for kind in [
        Some(MediaType::StaticImage),
        Some(MediaType::Text),
        None,
    ] {
        assert_eq!(
            bubble_menu_rows(kind),
            vec![BubbleMenuRow::ToTaskbar, BubbleMenuRow::Close],
            "a {kind:?} bubble offers only the two ways out"
        );
    }
}

use super::*;

/// What a bubble carries for each kind, read as the mark it is drawn with.
///
/// One question with two halves and both of them are asked here: which kinds carry a playback
/// glyph, and which of the two glyphs that is. A mark that answered only the first would be a
/// bubble over a paused film drawn with a play triangle, which is a bubble saying the opposite of
/// what the file is doing (see `MediaType::bubble_mark`).
#[test]
fn a_kinds_bubble_mark_is_the_glyph_of_what_that_kind_is_doing() {
    // The two that play: the glyph follows the playback.
    for kind in [MediaType::Video, MediaType::NativeVideo, MediaType::Audio] {
        assert_eq!(kind.bubble_mark(true), pin_chrome::BubbleMark::Play);
        assert_eq!(kind.bubble_mark(false), pin_chrome::BubbleMark::Pause);
        assert!(kind.bubble_mark(true).is_playback());
    }

    // A page: nothing about a page plays, so nothing about it changes with a playback it does not
    // have.
    for kind in [MediaType::Text, MediaType::Archive, MediaType::Peazip] {
        assert_eq!(kind.bubble_mark(true), pin_chrome::BubbleMark::Page);
        assert_eq!(kind.bubble_mark(false), pin_chrome::BubbleMark::Page);
    }

    // And everything else is a picture — including a file the pin was shown and could not draw,
    // whose bubble shows the failure frame itself (see `has_bubble_picture`).
    for kind in [
        MediaType::StaticImage,
        MediaType::Pdf,
        MediaType::Office,
        MediaType::Unplayable,
    ] {
        assert_eq!(kind.bubble_mark(true), pin_chrome::BubbleMark::Picture);
        assert_eq!(kind.bubble_mark(false), pin_chrome::BubbleMark::Picture);
        assert!(!kind.bubble_mark(true).is_playback());
    }
}

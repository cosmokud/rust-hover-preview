//! The pin's road out of being a pin: the caption's **Detach Window** button, the same row in both
//! menus, and the sound card's own button. Every one of them funnels into
//! [`detach_the_pinned_window`].
//!
//! **Nothing is reopened.** The file is already open here, the player that plays it is already begun,
//! and a sound is already at its own second — so detaching reads no file, starts no player and seeks
//! nothing. The pin becomes a window of this process's own, the tray is told that the viewer which
//! serves the hovers has left, and the process the tray began begins another (see `crate::app::detach`
//! and `app::viewer`). The notifications here are two published facts and one seam message, which is
//! why this is answered on whatever thread pressed the button.
//!
//! What is left of the old road — the hand-off, the box, the second a sound had got to — is kept for the
//! `--detach` process to be deleted with (see `crate::app::detach`).

use std::path::{Path, PathBuf};
use std::sync::atomic::Ordering;
use std::sync::Mutex;

use once_cell::sync::Lazy;

use super::pin_window::{pin_state, PIN_UP};

/// Where the sound on screen had got to, as the card last drew it, and the file it is a position in.
///
/// A sound is the one kind a hand-off cannot open the way it opens every other kind: a picture, a
/// film's first frame and a page are the same wherever they are read, while a second into a sound
/// is where the player on this side of the process boundary has got to, and nothing in the file
/// holds it. So it is kept here, written by the tick that draws the card's clock and read by the
/// hand-off, which is the only reason anything outside the loop's own state knows it (see
/// `note_the_sounds_position` and `the_pin_to_hand_over`).
static THE_SOUNDS_POSITION: Lazy<Mutex<Option<(PathBuf, f64)>>> = Lazy::new(|| Mutex::new(None));

/// Write down where the sound on screen had got to, which is the second a hand-off of that file is
/// taken up at.
///
/// A card that has no clock behind it — a sound no player would take — answers nothing, and a
/// position that is not a position is written as none rather than as the position of the file
/// before it.
pub(super) fn note_the_sounds_position(path: &Path, position: Option<f64>) {
    let Ok(mut kept) = THE_SOUNDS_POSITION.lock() else {
        return;
    };

    match position.filter(|seconds| seconds.is_finite() && *seconds >= 0.0) {
        Some(position) => {
            let changed = kept.as_ref().is_none_or(|(file, _)| file.as_path() != path);
            if changed {
                *kept = Some((path.to_path_buf(), position));
                return;
            }
            if let Some((_, at)) = kept.as_mut() {
                *at = position;
            }
        }
        // A file whose card has no clock is a file nothing is known about, and the position kept
        // for the file before it is not that answer.
        None => *kept = None,
    }
}

/// Where the sound of `path` had got to, or nothing where the card on screen is of another file or
/// is a card with no clock behind it.
#[allow(dead_code)]
fn the_sounds_position(path: &Path) -> Option<f64> {
    THE_SOUNDS_POSITION
        .lock()
        .ok()
        .and_then(|kept| kept.as_ref().map(|(file, at)| (file.clone(), *at)))
        .filter(|(file, _)| file.as_path() == path)
        .map(|(_, at)| at)
}

/// Detach the pinned window: it stops being a pin of the tray's and becomes a window of its own, here,
/// with everything it is holding left as it is.
///
/// Does nothing where there is no pin to detach — every road that reaches here after the pin it was drawn
/// for has gone — nothing where this window is one of its own already, and nothing where the run has no
/// block for the viewer the detach would leave it owing.
pub(super) fn detach_the_pinned_window() {
    // A window of its own has nothing left to become: it *is* the window a pin is detached to be, and
    // every road that offers the button — the caption's mark, a sound card's own button and both menus —
    // draws itself saying so. This is the one place that is true whatever is drawn.
    if crate::app::detach::is_detached() || !a_pin_is_standing() {
        return;
    }

    // **A window of its own is one the run has to be able to draw in.** The viewer that becomes it keeps the
    // block it claimed, and the tray begins another to serve the hovers it leaves, so a run with no block
    // left for that one would leave a window standing and the hovers with nobody — every hover after it put
    // in a ring nothing reads. Refused here, before the flip: the pin stays a pin and the hovers keep working
    // (see `seam::SEAM_VIEWERS` and `seam::a_block_can_be_had`).
    if !crate::app::seam::a_block_can_be_had() {
        // shortcut: a block claimed between this question and the tray's spawn still costs the replacement,
        // which the next hover begins again once a block is had; upgrade to one reserved for it.
        if std::env::var_os("RHP_SEAM_TRACE").is_some() {
            eprintln!("seam: no block for the window this pin would become; it stays a pin");
        }

        return;
    }

    // **The notice goes first, because the rings are the road only until the flip below.** It is what
    // begins the viewer that draws the hovers from here on, and this process stops reading that ring the
    // moment it is a window of its own (see `app::seam::the_preview_ring_is_ours`).
    crate::app::seam::send_tray(crate::app::seam::TrayMessage::ViewerDetached(
        std::process::id(),
    ));

    // The tray's half: the file is no longer what a click in the listing would be sent to, and no pin is up —
    // so a hover is admitted again over a window that keeps standing. Written through the published copies
    // and not through `pin_window::publish_pin_up`, because this process's own copy is the window this call
    // keeps (see `pin_window::PIN_UP_HERE`).
    crate::app::seam::publish_the_pinned_path(None);
    PIN_UP.store(false, Ordering::Release);

    // And the window's own, last: nothing is taken up, nothing is read, no box is rebuilt, and every press
    // the chrome draws from here on reads the window as one of its own (see
    // `crate::app::detach::this_window_is_its_own`).
    crate::app::detach::this_window_is_its_own();
}

/// Whether there is a pin on this screen to detach: the state rather than either published copy, because
/// the copies answer the chrome's question and this one is the road's own.
fn a_pin_is_standing() -> bool {
    pin_state().is_some_and(|state| state.pin().is_some())
}

/// What is handed over: the file, the media box it is drawn in, whether the window is maximized,
/// and where the sound of it had got to.
///
/// The last of those is unanswered for every kind but a sound: a file has no second of its own to
/// be resumed at, and the window that comes up reads one wherever it is a kind that plays.
///
/// The box is the *media's*, which is what a take-up is given — hand it the window's box and the
/// window that comes up is the window this one is with a second frame of chrome around it. And the
/// maximize is asked for rather than carried: `restore` is written only while a window is maximized,
/// so a window that has one is a window whose restore box exists and whose box is the display's.
#[allow(dead_code)]
type HandedOver = (std::path::PathBuf, (i32, i32, i32, i32), bool, Option<f64>);

// What the `--detach` road carries: kept with it and unused until it is deleted (see the module doc).
#[allow(dead_code)]
fn the_pin_to_hand_over() -> Option<HandedOver> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;

    // Where the sound on screen had got to, which is read here rather than by the window that takes
    // it up: what the hand-off becomes is the window that was standing, and a sound that came up at
    // the second its own file alone would name — its beginning, or wherever a hover of it last left
    // off — would be the one thing about it that moved. A pin of any other kind has no second to
    // hand over, and the answer for it is nothing (see `note_the_sounds_position`).
    let at = the_sounds_position(&pin.path);

    Some((pin.path.clone(), pin.content, pin.restore.is_some(), at))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ui::preview_window::pin_model::PinnedPreview;
    use crate::ui::preview_window::pin_window::{end_pin, install, Reason, RecordedPinWindow};
    use crate::ui::preview_window::tests::the_room_to_itself;

    /// The flip is a process's own for the rest of its life, and a test process outlives more windows than
    /// one: the flag is turned back off when this goes out of scope, including when a body panics.
    struct AWindowOfItsOwnNoLonger;

    impl Drop for AWindowOfItsOwnNoLonger {
        fn drop(&mut self) {
            crate::app::detach::forget_this_window_is_its_own();
        }
    }

    #[test]
    fn detaching_makes_the_pin_a_window_of_its_own_and_tells_the_tray_to_begin_another() {
        let _room = the_room_to_itself();

        let mut pin = PinnedPreview::for_test();
        pin.path = PathBuf::from("sound.flac");
        install(pin);
        assert!(
            crate::app::seam::the_pin_is_up(),
            "a pin the tray can drive is up to begin with"
        );
        assert_eq!(
            crate::app::seam::pinned_path(),
            Some(PathBuf::from("sound.flac")),
            "and the tray knows the file a click would be sent to"
        );

        detach_the_pinned_window();
        let _flipped = AWindowOfItsOwnNoLonger;
        let told = crate::app::seam::take_tray_messages();

        // **The notice the tray begins another viewer from**, and it names this process: that is the one
        // which is standing on the screen from here on and serving no hovers.
        assert!(
            told.iter().any(|message| matches!(
                message,
                crate::app::seam::TrayMessage::ViewerDetached(pid) if *pid == std::process::id()
            )),
            "the tray is told which viewer has left: {} notices",
            told.len()
        );
        assert!(
            crate::app::detach::is_detached(),
            "and every press the chrome draws from here on reads the window as one of its own"
        );

        // **The pin stops being a pin, and the window keeps standing.** No pin of the tray's is up — so a
        // hover is admitted again over this window — and no file is named where a click would be sent, while
        // this process still has a window to draw and still has the file in it.
        assert!(
            !crate::app::seam::the_pin_is_up(),
            "the tray has no pin any more"
        );
        assert_eq!(
            crate::app::seam::pinned_path(),
            None,
            "and no file to send to"
        );
        assert!(
            super::super::requests::pinned(),
            "while the window itself stands: nothing was taken down"
        );
        assert_eq!(
            pin_state().and_then(|state| state.pin().map(|pin| pin.path.clone())),
            Some(PathBuf::from("sound.flac")),
            "and nothing was read again: the window is still on the file it was showing"
        );

        // **And every road out is drawn as spent**, from the same flag: the caption's mark, the rows of both
        // menus, and the button in a sound card's own top margin.
        assert_eq!(
            super::super::paint::pinned_paint().map(|paint| paint.detachable),
            Some(false),
            "the caption draws the detach mark as one a window of its own has spent"
        );
        let rows = super::super::pin_menu::menu_rows(crate::app::detach::is_detached());
        assert!(
            rows[row_at(&rows, "Detach Window")].disabled,
            "the card's menu keeps the row it has always listed, and asks it as one spent"
        );
        let bubble = super::super::pin_bubble::bubble_menu_chrome_rows(
            Some(super::super::media_types::MediaType::Audio),
            false,
            false,
            true,
            false,
            crate::app::detach::is_detached(),
        );
        assert!(
            bubble[row_at(&bubble, "Detach Window")].disabled,
            "and so does the menu a standing window and a bubble share"
        );
        assert_eq!(
            super::super::pin_lifecycle::pinned_audio_chrome(None).map(|chrome| chrome.detached),
            Some(true),
            "and the sound card's own button says it in its own place"
        );

        // The window's own end goes on working in a window of its own, and ends nothing else.
        end_pin(Reason::Closed, &RecordedPinWindow::new(0x4100));
        assert!(
            !super::super::requests::pinned(),
            "the window came down when it was closed"
        );
        assert!(
            !crate::app::seam::the_pin_is_up(),
            "and published no pin of the tray's on its way out"
        );
    }

    fn row_at(rows: &[crate::text::pin_chrome::MenuRow], label: &str) -> usize {
        rows.iter()
            .position(|row| row.label == label)
            .unwrap_or_else(|| panic!("the menu holds the {label} row"))
    }

    #[test]
    fn nothing_is_handed_over_where_no_pin_is_standing() {
        let _room = the_room_to_itself();

        assert_eq!(the_pin_to_hand_over(), None);
    }

    #[test]
    fn a_pin_standing_hands_over_its_file_its_media_box_and_whether_it_is_maximized() {
        let _room = the_room_to_itself();

        let mut pin = PinnedPreview::for_test();
        // The box it is drawn in and the box a restore would put it back at: a pin that has one of
        // the second is a pin that is maximized.
        pin.content = (40, 60, 240, 160);
        pin.restore = Some((0, 0, 400, 300));
        install(pin);

        let (path, region, maximized, at) = the_pin_to_hand_over().expect("a pin is standing");
        assert_eq!(path, std::path::PathBuf::from("picture.png"));
        assert_eq!(region, (40, 60, 240, 160));
        assert!(maximized);
        // A picture has no second to be resumed at, and nothing was ever drawn for one.
        assert_eq!(at, None);

        end_pin(Reason::Closed, &RecordedPinWindow::new(0x4000));
    }

    #[test]
    fn a_window_that_is_not_maximized_hands_over_no_maximize() {
        let _room = the_room_to_itself();

        let pin = PinnedPreview::for_test();
        assert_eq!(pin.restore, None);
        install(pin);

        let (_, region, maximized, _) = the_pin_to_hand_over().expect("a pin is standing");
        assert_eq!(region, (0, 0, 100, 100));
        assert!(!maximized);

        end_pin(Reason::Closed, &RecordedPinWindow::new(0x4001));
    }

    #[test]
    fn a_sound_hands_over_the_second_its_card_had_drawn() {
        let _room = the_room_to_itself();
        *THE_SOUNDS_POSITION
            .lock()
            .unwrap_or_else(|held| held.into_inner()) = None;

        let mut pin = PinnedPreview::for_test();
        pin.path = PathBuf::from("sound.flac");
        install(pin);
        note_the_sounds_position(Path::new("sound.flac"), Some(100.5));

        let (_, _, _, at) = the_pin_to_hand_over().expect("a pin is standing");
        assert_eq!(at, Some(100.5));

        end_pin(Reason::Closed, &RecordedPinWindow::new(0x4002));
    }

    #[test]
    fn a_second_drawn_for_another_file_is_not_this_window_s() {
        let _room = the_room_to_itself();
        *THE_SOUNDS_POSITION
            .lock()
            .unwrap_or_else(|held| held.into_inner()) = None;

        let mut pin = PinnedPreview::for_test();
        pin.path = PathBuf::from("sound.flac");
        install(pin);
        // The card on screen is of the file after this one — a step taken while the hand-off was
        // asked for — and its second is not a second in the file the pin is handed over on.
        note_the_sounds_position(Path::new("another.flac"), Some(220.0));

        let (_, _, _, at) = the_pin_to_hand_over().expect("a pin is standing");
        assert_eq!(at, None);

        end_pin(Reason::Closed, &RecordedPinWindow::new(0x4003));
    }

    #[test]
    fn a_card_with_no_clock_writes_down_no_second() {
        let _room = the_room_to_itself();
        *THE_SOUNDS_POSITION
            .lock()
            .unwrap_or_else(|held| held.into_inner()) = None;

        note_the_sounds_position(Path::new("sound.flac"), Some(100.5));
        note_the_sounds_position(Path::new("sound.flac"), None);
        assert_eq!(the_sounds_position(Path::new("sound.flac")), None);

        // A second that is not one — a player that has gone and answered with a NaN — is written
        // down as none rather than kept as the position of the card before it.
        note_the_sounds_position(Path::new("sound.flac"), Some(100.5));
        note_the_sounds_position(Path::new("sound.flac"), Some(f64::NAN));
        assert_eq!(the_sounds_position(Path::new("sound.flac")), None);

        // And a second is kept for the file it is a second in, and for no other.
        note_the_sounds_position(Path::new("sound.flac"), Some(12.0));
        assert_eq!(the_sounds_position(Path::new("sound.flac")), Some(12.0));
        assert_eq!(the_sounds_position(Path::new("another.flac")), None);
    }
}

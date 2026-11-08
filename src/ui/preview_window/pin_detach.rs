//! The pin's road out of this process: the caption's **Detach Window** button, the same row in both
//! menus, and the hand-off they make.
//!
//! What it hands over is the whole of what a window needs to be the same window somewhere else —
//! the file, the box it is drawn in, and whether it is maximized — and nothing about the pin's own
//! state, because there is no state to hand over: the window the hand-off becomes is a take-up of
//! the same file at the same box (see `crate::app::detach`).
//!
//! Ask stays here rather than in the window's own procedure for one reason: the hand-off begins a
//! process, and a process is not something the thread that draws a window and pumps its own messages
//! should be waiting on. The ask is made on a thread of its own, and the pin it came from is taken
//! down by the answer — the new window's first paint — rather than by this call.

use std::path::{Path, PathBuf};
use std::sync::Mutex;

use once_cell::sync::Lazy;

use super::pin_window::pin_state;

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
fn the_sounds_position(path: &Path) -> Option<f64> {
    THE_SOUNDS_POSITION
        .lock()
        .ok()
        .and_then(|kept| kept.as_ref().map(|(file, at)| (file.clone(), *at)))
        .filter(|(file, _)| file.as_path() == path)
        .map(|(_, at)| at)
}

/// Hand the pinned window to a process of its own.
///
/// Does nothing where there is no pin to hand over, which is every road that reaches here after the
/// pin it was drawn for has gone.
pub(super) fn detach_the_pinned_window() {
    // A window of its own has nothing left to hand on: it *is* the window a pin was handed over to
    // become, and a second hand-off would be one process handing its own window to a third for no
    // choice the user made (see `app::detach::is_detached`). Every road that offers the hand-off —
    // the caption's button, a sound card's window button and both menus — draws itself saying the
    // same, and this is the one place that is true whatever is drawn.
    if crate::app::detach::is_detached() {
        return;
    }

    let Some((path, region, maximized, at)) = the_pin_to_hand_over() else {
        return;
    };

    // The hook reads the end this brings as a hand-off rather than as a close, so no hover of the
    // file under the pointer is opened over the window that has just come up — said where that end
    // is asked for rather than here, because a hand-off this machine refuses ends no pin at all (see
    // `crate::app::detach::wait_for_the_hand_over`).

    std::thread::spawn(move || {
        // A hand-off this machine refuses — a second process that cannot be begun at all — is a
        // hand-off that never happened: nothing was noted as handed over, no window came up, and
        // the pin it came from is still standing with the file in it.
        let _ = crate::app::detach::hand_off(&path, region, maximized, at);
    });
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
type HandedOver = (std::path::PathBuf, (i32, i32, i32, i32), bool, Option<f64>);

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

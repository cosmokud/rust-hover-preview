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

use super::pin_window::pin_state;

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

    let Some((path, region, maximized)) = the_pin_to_hand_over() else {
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
        let _ = crate::app::detach::hand_off(&path, region, maximized);
    });
}

/// What is handed over: the file, the media box it is drawn in, and whether the window is maximized.
///
/// The box is the *media's*, which is what a take-up is given — hand it the window's box and the
/// window that comes up is the window this one is with a second frame of chrome around it. And the
/// maximize is asked for rather than carried: `restore` is written only while a window is maximized,
/// so a window that has one is a window whose restore box exists and whose box is the display's.
type HandedOver = (std::path::PathBuf, (i32, i32, i32, i32), bool);

fn the_pin_to_hand_over() -> Option<HandedOver> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;

    Some((pin.path.clone(), pin.content, pin.restore.is_some()))
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

        let (path, region, maximized) = the_pin_to_hand_over().expect("a pin is standing");
        assert_eq!(path, std::path::PathBuf::from("picture.png"));
        assert_eq!(region, (40, 60, 240, 160));
        assert!(maximized);

        end_pin(Reason::Closed, &RecordedPinWindow::new(0x4000));
    }

    #[test]
    fn a_window_that_is_not_maximized_hands_over_no_maximize() {
        let _room = the_room_to_itself();

        let pin = PinnedPreview::for_test();
        assert_eq!(pin.restore, None);
        install(pin);

        let (_, region, maximized) = the_pin_to_hand_over().expect("a pin is standing");
        assert_eq!(region, (0, 0, 100, 100));
        assert!(!maximized);

        end_pin(Reason::Closed, &RecordedPinWindow::new(0x4001));
    }
}

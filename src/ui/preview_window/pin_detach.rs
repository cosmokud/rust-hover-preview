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

use std::sync::atomic::Ordering;

use super::pin_window::{pin_state, PIN_UP};

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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ui::preview_window::pin_model::PinnedPreview;
    use crate::ui::preview_window::pin_window::{end_pin, install, Reason, RecordedPinWindow};
    use crate::ui::preview_window::tests::the_room_to_itself;
    use std::path::PathBuf;

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
}

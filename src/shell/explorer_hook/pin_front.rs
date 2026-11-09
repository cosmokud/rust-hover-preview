//! Where the front stands on a pinned tick, as the hook reads it: whether the foreground changed, which
//! is what the item and view caches are dropped on.
//!
//! **A foreground transition is not a demotion, and this file no longer pretends it is.** A press inside
//! the window that is already active — the whole of what a listing is made of — moves no foreground at
//! all: the frame stays the foreground window and only a child's focus moves, so a lowering wired to
//! "Explorer has come to the front" waited for a transition that press cannot make. The pin is put behind
//! the window a press landed in, by the press road itself (see `window_state::a_press_puts_us_behind`).
//!
//! What is left is the one thing the reading was always good for. The caches the pick is answered from
//! are frozen while a pin is up, and a pin that has just taken the focus or given it back is exactly when
//! the user is about to pick — so they are dropped on the transition, which is the one tick it happens on.
//!
//! **Nothing reads the foreground while no pin is up**, so the reading a pin comes up on is its baseline
//! rather than a transition: a baseline left over from the pin before is a drop made for a window that has
//! not been looked at for the whole of the hover between them (see `PinFront::read`).

/// Where the front stands: whether a pin has been read before, and where Explorer stood on that reading.
#[derive(Default)]
pub(super) struct PinFront {
    up: bool,
    explorer_foreground: bool,
}

/// One pinned tick's reading of the foreground (see `PinFront::read`).
#[derive(Clone, Copy, Default)]
pub(super) struct FrontReading {
    /// The foreground is not the window it was on the last reading while a pin was up: what the item and
    /// view caches are dropped on, since the pin taking the focus and giving it back moves neither.
    pub(super) changed: bool,
}

impl PinFront {
    /// No pin is up any more, so the next reading of a pin is that pin's own baseline.
    pub(super) fn pin_is_down(&mut self) {
        self.up = false;
    }

    /// One reading of the foreground, taken on a tick whose pin is up.
    pub(super) fn read(&mut self, explorer_foreground: bool) -> FrontReading {
        let changed = self.up && explorer_foreground != self.explorer_foreground;

        self.up = true;
        self.explorer_foreground = explorer_foreground;

        FrontReading { changed }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// **A pin raised over a listing is not a transition.** Explorer is the foreground window the whole
    /// time — the pin is `WS_EX_NOACTIVATE` — so the reading the pin comes up on is its baseline and every
    /// press a listing is made of leaves it exactly where it was: no drop, and nothing here for a lowering
    /// to hang on.
    #[test]
    fn a_pin_raised_over_a_listing_reads_its_baseline() {
        let mut front = PinFront::default();

        let raised = front.read(true);
        assert!(!raised.changed, "a pin's first reading is its baseline");

        for _ in 0..5 {
            assert!(!front.read(true).changed, "and a press in it moves it not");
        }
    }

    /// The transition is still the one thing read here: the pin took the focus and gave it back, which is
    /// when the caches below the pick are stale — the view the shell answers a click from has moved, and
    /// the click that follows is answered out of it.
    #[test]
    fn the_pin_taking_the_focus_and_giving_it_back_is_the_drop() {
        let mut front = PinFront::default();
        front.read(true);

        let pin_held_it = front.read(false);
        assert!(
            pin_held_it.changed,
            "the caches go with the pin taking the focus"
        );

        let listing = front.read(true);
        assert!(listing.changed, "and with the listing taking it back");
    }

    /// The baseline is the reading of the pin that is up and of no other: a pin taken down with the
    /// foreground wherever the next one finds it starts again from that, not from the pin before it.
    #[test]
    fn the_baseline_belongs_to_the_pin_that_is_up() {
        let mut front = PinFront::default();
        front.read(true);
        front.read(false);

        front.pin_is_down();
        let raised = front.read(true);
        assert!(
            !raised.changed,
            "a pin that has just come up is asked its baseline, not the pin before's"
        );
    }
}

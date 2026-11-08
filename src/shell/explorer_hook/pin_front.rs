//! Where a standing pin's front stands, as the pinned tick reads it: whether Explorer has just taken
//! that front back from the pin, which is the one thing the demotion ask is for (see
//! `app::seam::ask_the_pin_to_go_behind`).
//!
//! **A pin is in front of the listing it was raised over, and the listing keeps the keyboard.** A pin
//! is `WS_EX_NOACTIVATE`, so the raise leaves Explorer the foreground window, and "Explorer is in front"
//! is not the question: a raise is the pin in front of a listing that already had the foreground, and an
//! ask made on that state is the pin taken down the tick it went up.
//!
//! What is asked instead is the pin's own history: a pin *held* the foreground (the user clicked into it,
//! or it activated itself on its way up) and Explorer has it back now. Only that answer is a pin to put
//! behind the listing, and it is the one a press in the listing cannot give — a click in the window that
//! is already active reorders nothing, so nothing takes the pin down while Explorer is the window the
//! user came back to.
//!
//! **Nothing reads the foreground while no pin is up**, so the reading a pin comes up on is its baseline
//! rather than a transition: a baseline left over from the pin before is an answer about a window that
//! has not been looked at for the whole of the hover between them, and read as a transition it is the
//! raise demoted by the state it was raised over (see `PinFront::read`).

/// Where the front stands: whether a pin has been read before, whether Explorer held the foreground on
/// that reading, and whether the pin itself has held it since it came up.
#[derive(Default)]
pub(super) struct PinFront {
    up: bool,
    explorer_foreground: bool,
    held: bool,
}

/// One pinned tick's reading of the foreground (see `PinFront::read`).
#[derive(Clone, Copy, Default)]
pub(super) struct FrontReading {
    /// The foreground is not the window it was on the last reading while a pin was up: what the item and
    /// view caches are dropped on, since the pin taking the focus and giving it back moves neither.
    pub(super) changed: bool,
    /// Explorer has just taken the front back from the pin, which is the ask that puts the pin behind it.
    pub(super) explorer_took_the_front_back: bool,
}

impl PinFront {
    /// No pin is up any more, so the next reading of a pin is that pin's own baseline.
    pub(super) fn pin_is_down(&mut self) {
        self.up = false;
    }

    /// One reading of the foreground, taken on a tick whose pin is up.
    pub(super) fn read(&mut self, explorer_foreground: bool) -> FrontReading {
        let changed = self.up && explorer_foreground != self.explorer_foreground;
        // Only a pin that has held the foreground can have it taken back: a pin whose raise left the
        // listing in front was in front of that listing already, and putting it behind the listing it
        // was raised over is the fault this whole file is about.
        let explorer_took_the_front_back = self.up && explorer_foreground && self.held;

        self.up = true;
        self.explorer_foreground = explorer_foreground;
        // A pin that came up holding the foreground keeps the front it has to lose; one that came up over
        // a listing only holds it once the user has clicked into it.
        self.held = !explorer_foreground;

        FrontReading {
            changed,
            explorer_took_the_front_back,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// **A pin raised over a listing stays in front.** Explorer is the foreground window the whole time
    /// — the pin is `WS_EX_NOACTIVATE` — so what the ask must not read as a transition is the state the
    /// pin was raised over, on the tick it comes up and on every one after it.
    #[test]
    fn a_pin_raised_over_a_listing_is_not_asked_to_go_behind() {
        let mut front = PinFront::default();

        let raised = front.read(true);
        assert!(!raised.changed, "a pin's first reading is its baseline");
        assert!(!raised.explorer_took_the_front_back);

        for _ in 0..5 {
            let later = front.read(true);
            assert!(!later.changed);
            assert!(!later.explorer_took_the_front_back);
        }
    }

    /// **And the ask is still made for the case it was added for**: the pin had the front — the user
    /// clicked into it — and Explorer has it back. This is the one edge that no press in the listing can
    /// answer, since a click in the window that is already active reorders nothing.
    #[test]
    fn explorer_taking_the_front_back_from_the_pin_is_the_ask() {
        let mut front = PinFront::default();
        front.read(true);

        let pin_held_it = front.read(false);
        assert!(
            pin_held_it.changed,
            "the caches go with the pin taking the focus"
        );
        assert!(
            !pin_held_it.explorer_took_the_front_back,
            "the pin holding the foreground is the pin in front"
        );

        let listing = front.read(true);
        assert!(listing.changed);
        assert!(listing.explorer_took_the_front_back);
    }

    /// A pin whose raise took the foreground itself is the same debt, read the other way round: it holds
    /// the front from its first reading, and the click that gives it back asks for the lowering.
    #[test]
    fn a_pin_that_came_up_in_front_loses_it_to_explorer() {
        let mut front = PinFront::default();
        let raised = front.read(false);
        assert!(!raised.changed);
        assert!(!raised.explorer_took_the_front_back);

        assert!(front.read(true).explorer_took_the_front_back);
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
            !raised.changed && !raised.explorer_took_the_front_back,
            "a pin that has just come up is asked its baseline, not the pin before's"
        );
    }
}

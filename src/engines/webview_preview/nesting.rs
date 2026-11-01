//! Where the engine's window stands: a window of the desktop's, or a child of the window the
//! preview is in.
//!
//! **A page is a child of the pin while a pin is up, and that is what a taskbar peek needs.**
//! The Shell composes a hover of a window from the window itself — its children included — and a
//! page left as a window of the desktop's is no part of the pin: the peek showed the chrome and
//! the band's hole. Made a child of the pin, the page is drawn into the same surface DWM freezes,
//! which is the same arrangement the chrome strips already have. While no pin is up the page goes
//! back to the window of the desktop's it was: a hover window refuses activation, and a page that
//! runs needs to be able to be the foreground.
//!
//! A `WS_CHILD` cannot be activated on its own, so the redirection of the keyboard into a page that
//! runs is the pin's (see `mouse_activate_answers` and `Host::put_window`).

use windows::Win32::Foundation::{HWND, POINT};
use windows::Win32::Graphics::Gdi::ClientToScreen;
use windows::Win32::UI::WindowsAndMessaging::{WS_CHILD, WS_EX_TOPMOST, WS_POPUP};

use super::api::Area;
use super::environment::ex_style_for;

/// Where the page's window stands, and the style it wore before it was a child.
///
/// The old style is kept because a page that goes back to being a window of the desktop's has to
/// be exactly the window it was — a hover's topmost, non-activating tool window — rather than a
/// second copy of that decision written out at the unnest.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub(super) struct Nesting {
    parent: Option<HWND>,
    desktop_ex_style: isize,
}

/// What asking for a parent changed.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) enum NestingChange {
    /// It was on the desktop and is now a child.
    Entered,
    /// It was a child and is now on the desktop.
    Left,
    /// It was a child of another window and is now a child of this one.
    Moved,
    /// It is already where it was asked to be: nothing for the window to do.
    Stayed,
}

impl Nesting {
    pub(super) const fn new() -> Self {
        Self {
            parent: None,
            desktop_ex_style: 0,
        }
    }

    pub(super) fn parent(self) -> Option<HWND> {
        self.parent
    }

    /// The extended style to put back when the page leaves the desktop's band.
    pub(super) fn desktop_ex_style(self) -> isize {
        self.desktop_ex_style
    }

    /// Record the parent the page is asked to stand in, answering what the window has to do.
    ///
    /// `current_ex_style` is read from the window only where it is leaving the desktop, because
    /// that is the one moment the style it wore is not the one this stores. Asking for the same
    /// parent again is `Stayed`, which is what keeps a placement made every pointer move from
    /// reparenting the page every pointer move.
    pub(super) fn settle(
        &mut self,
        parent: Option<HWND>,
        current_ex_style: isize,
    ) -> NestingChange {
        if self.parent == parent {
            return NestingChange::Stayed;
        }

        let was = self.parent;
        if was.is_none() && parent.is_some() {
            self.desktop_ex_style = current_ex_style;
        }
        self.parent = parent;

        match (was, parent) {
            (None, Some(_)) => NestingChange::Entered,
            (Some(_), None) => NestingChange::Left,
            (Some(_), Some(_)) => NestingChange::Moved,
            (None, None) => NestingChange::Stayed,
        }
    }
}

/// The style a page's window carries as a child: `WS_CHILD` in place of `WS_POPUP`.
///
/// The same rule the chrome strips are made by — a child is drawn into its parent's surface (see
/// `pin_chrome_strips`).
pub(super) fn child_style(style: isize) -> isize {
    (style & !(WS_POPUP.0 as isize)) | WS_CHILD.0 as isize
}

/// And the style it carries as a window of the desktop's again.
pub(super) fn popup_style(style: isize) -> isize {
    (style & !(WS_CHILD.0 as isize)) | WS_POPUP.0 as isize
}

/// The extended style a page carries as a child: a child has no band of its own, so `WS_EX_TOPMOST`
/// is taken off, and the activation refusal stays the document's (see `ex_style_for`).
pub(super) fn child_ex_style(current: isize, runs: bool) -> isize {
    ex_style_for(current & !(WS_EX_TOPMOST.0 as isize), runs)
}

/// The page's box in its parent's client coordinates, from the screen box the pin's band was
/// measured in and the parent's client origin.
///
/// A child is positioned in client units at every display scale, so this is the whole of the DPI
/// arithmetic: both boxes are the pixels the app is already working in.
pub(super) fn client_band(area: Area, parent_origin: (i32, i32)) -> (i32, i32) {
    (area.x - parent_origin.0, area.y - parent_origin.1)
}

/// The screen point of a window's client origin.
pub(super) fn parent_origin(parent: HWND) -> (i32, i32) {
    let mut origin = POINT::default();
    // Safety: the handle is a window this app holds, and this asks where its client starts.
    unsafe {
        let _ = ClientToScreen(parent, &mut origin);
    }
    (origin.x, origin.y)
}

/// What a nested placement does, in the order it must happen.
///
/// The page is put first, which puts it above the chrome the pin already has standing; the chrome
/// is then raised back over it. Held as values rather than written inline so the suite can hold the
/// order to the one thing it must be.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) enum NestedCall {
    PlacePage,
    RaiseChrome,
}

/// Whether a placement has anything to do: the page is placed again for a box that changed and for
/// a page that changed which window it stands in, and for nothing else.
///
/// A placement arrives on every pointer move of a drag, and a page put in the box it already has —
/// and the chrome raised over it again — is four z-order changes and a recomposited layered pin per
/// tick (see `Host::put_window`).
pub(super) fn the_page_is_placed(
    change: NestingChange,
    last_placed: Option<(i32, i32, i32, i32)>,
    box_: (i32, i32, i32, i32),
) -> bool {
    change != NestingChange::Stayed || last_placed != Some(box_)
}

pub(super) fn nested_calls(over_pin: bool) -> [Option<NestedCall>; 2] {
    [
        Some(NestedCall::PlacePage),
        over_pin.then_some(NestedCall::RaiseChrome),
    ]
}

/// What putting the page back on the desktop does, in the order it must happen.
///
/// **The parent is given up before the style is.** A window styled `WS_POPUP` while it is still the
/// owned child of the pin is a popup the pin owns, and it hides with the pin for the rest of the
/// run: the parent goes first, the popup style after it, and the style it wore on the desktop last.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(super) enum UnnestCall {
    LeaveParent,
    PopupStyle,
    DesktopExStyle,
}

pub(super) fn unnest_calls() -> [Option<UnnestCall>; 3] {
    [
        Some(UnnestCall::LeaveParent),
        Some(UnnestCall::PopupStyle),
        Some(UnnestCall::DesktopExStyle),
    ]
}

/// The parent a page is to be settled in, and whether the one it stands in was taken away under it.
///
/// A pin's window is destroyed when the display arrangement changes, and a page left a child of a
/// destroyed window is a page nobody shows again: what is recorded is forgotten rather than stood
/// in, and the page is put back on the desktop (see `Nesting::settle`). `alive` is `IsWindow`, taken
/// as a parameter so the decision is a thing a test can hold.
pub(super) fn settle_parents(
    recorded: Option<HWND>,
    requested: Option<HWND>,
    alive: impl Fn(HWND) -> bool,
) -> (Option<HWND>, bool) {
    let orphaned = recorded.is_some_and(|parent| !alive(parent));
    (requested.filter(|parent| alive(*parent)), orphaned)
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows::Win32::UI::WindowsAndMessaging::{
        WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST,
    };

    /// A fake handle is enough for the state machine: nothing here asks the system anything.
    fn a_window(number: isize) -> HWND {
        HWND(number as *mut _)
    }

    /// The parent the page stands in is the pin while one is up and the hover's window when it is
    /// not, and asking for the same parent again is nothing to do — a placement arrives on every
    /// pointer move, and a reparent on every pointer move is a page that flickers.
    #[test]
    fn a_page_is_reparented_only_when_the_window_it_stands_in_changes() {
        let pin = a_window(0x1000);
        let hover = a_window(0x2000);
        let mut nesting = Nesting::new();

        assert_eq!(
            nesting.settle(Some(pin), 0x1234),
            NestingChange::Entered,
            "a page on the desktop is made a child of the pin"
        );
        assert_eq!(nesting.parent(), Some(pin));
        assert_eq!(
            nesting.settle(Some(pin), 0x9999),
            NestingChange::Stayed,
            "and the next placement in the same pin reparents nothing"
        );

        assert_eq!(
            nesting.settle(Some(hover), 0x9999),
            NestingChange::Moved,
            "unpinning puts the page in the hover's window"
        );
        assert_eq!(nesting.parent(), Some(hover));

        assert_eq!(
            nesting.settle(None, 0x9999),
            NestingChange::Left,
            "and with no preview window it stands on the desktop again"
        );
        assert_eq!(nesting.parent(), None);
        assert_eq!(
            nesting.settle(None, 0x9999),
            NestingChange::Stayed,
            "asking for the desktop twice is asking once"
        );
    }

    /// The style a child wears is a child's, and the style it goes back to is the one it wore
    /// before it was one — the pin's window tree and the hover's are the same window.
    #[test]
    fn the_page_is_a_child_while_nested_and_a_popup_again_afterwards() {
        let (popup, popup_ex) = (WS_POPUP.0 as isize, 0x1234);
        let child = child_style(popup);

        assert_eq!(child & WS_CHILD.0 as isize, WS_CHILD.0 as isize);
        assert_eq!(child & WS_POPUP.0 as isize, 0);
        assert_eq!(
            popup_style(child) & WS_POPUP.0 as isize,
            WS_POPUP.0 as isize
        );
        assert_eq!(popup_style(child) & WS_CHILD.0 as isize, 0);

        let ex = child_ex_style(
            WS_EX_TOPMOST.0 as isize | WS_EX_TOOLWINDOW.0 as isize | WS_EX_NOACTIVATE.0 as isize,
            false,
        );
        assert_eq!(ex & WS_EX_TOPMOST.0 as isize, 0, "a child has no band");
        assert_eq!(
            ex & WS_EX_TOOLWINDOW.0 as isize,
            WS_EX_TOOLWINDOW.0 as isize
        );
        assert_eq!(
            ex & WS_EX_NOACTIVATE.0 as isize,
            WS_EX_NOACTIVATE.0 as isize,
            "a document is looked at, not typed into, as a child too"
        );
        assert_eq!(
            child_ex_style(WS_EX_NOACTIVATE.0 as isize, true) & WS_EX_NOACTIVATE.0 as isize,
            0,
            "and a page that runs is still allowed the keyboard"
        );

        let mut nesting = Nesting::new();
        nesting.settle(Some(a_window(0x1000)), popup_ex);
        assert_eq!(
            nesting.desktop_ex_style(),
            popup_ex,
            "the window it was before nesting is the window it goes back to"
        );
    }

    /// A child is placed in its parent's client units, and the band the pin measured is in screen
    /// units: the same band is another pair of numbers inside a pin that has been carried.
    #[test]
    fn the_band_is_the_pins_client_band_and_not_its_screen_box() {
        let area = Area {
            x: 1200,
            y: 340,
            width: 640,
            height: 480,
        };

        assert_eq!(client_band(area, (1000, 100)), (200, 240));
        assert_eq!(
            client_band(area, (0, 0)),
            (area.x, area.y),
            "a parent at the desktop's origin needs no conversion"
        );
        assert_eq!(
            client_band(area, (-2000, -1000)),
            (3200, 1340),
            "and one carried off the top-left is converted the other way"
        );
    }

    /// The page is put above the pin's existing children, and the chrome is then raised back over
    /// it — the one order that leaves the caption over the page rather than under it.
    #[test]
    fn the_page_is_placed_before_the_chrome_is_raised_back_over_it() {
        assert_eq!(
            nested_calls(true),
            [Some(NestedCall::PlacePage), Some(NestedCall::RaiseChrome)]
        );
        assert_eq!(
            nested_calls(false),
            [Some(NestedCall::PlacePage), None],
            "a hover's window has no chrome strips to raise"
        );
    }

    /// Coming out is the other way round from going in: a page whose style says `WS_POPUP` while it
    /// is still the pin's owned child hides with the pin for the rest of the run.
    #[test]
    fn the_parent_is_given_up_before_the_popup_style_is_put_back() {
        assert_eq!(
            unnest_calls(),
            [
                Some(UnnestCall::LeaveParent),
                Some(UnnestCall::PopupStyle),
                Some(UnnestCall::DesktopExStyle),
            ]
        );
    }

    /// A parent Windows no longer has is never stood in, and the page it held is a page of the
    /// desktop's again — asked for the same dead parent or for nothing at all.
    #[test]
    fn a_destroyed_parent_is_forgotten_rather_than_stood_in() {
        let pin = a_window(0x1000);
        let alive = |window: HWND| window != pin;

        assert_eq!(
            settle_parents(Some(pin), Some(pin), alive),
            (None, true),
            "a page in a pin that is gone is put back on the desktop"
        );
        assert_eq!(
            settle_parents(Some(pin), None, alive),
            (None, true),
            "and leaving a parent that is gone is still its parent being gone"
        );
        assert_eq!(
            settle_parents(None, Some(pin), alive),
            (None, false),
            "a pin asked for that is not there is not a parent either"
        );
        assert_eq!(
            settle_parents(Some(pin), Some(pin), |_| true),
            (Some(pin), false),
            "and a pin that is still there is the parent the page stands in"
        );
        assert_eq!(
            settle_parents(None, None, |_| true),
            (None, false),
            "a page already on the desktop is nobody's orphan"
        );
    }

    /// A tick that asks for the box the page is already in places nothing, and one that moved the
    /// page does: a drag of any length costs one placement per box, not one per tick.
    #[test]
    fn a_stayed_placement_in_the_same_box_places_nothing() {
        let box_ = (100, 200, 640, 480);

        assert!(!the_page_is_placed(NestingChange::Stayed, Some(box_), box_));
        assert!(the_page_is_placed(
            NestingChange::Stayed,
            Some(box_),
            (101, 200, 640, 480)
        ));
        assert!(the_page_is_placed(NestingChange::Stayed, None, box_));
        assert!(the_page_is_placed(NestingChange::Entered, Some(box_), box_));
        assert!(the_page_is_placed(NestingChange::Moved, Some(box_), box_));
    }

    /// Settling into the parent it is already in is answered as `Stayed`, which is what makes a
    /// placement that only moved nothing to do — the tick asks again on every pointer move.
    #[test]
    fn a_settle_into_the_same_parent_with_a_live_one_changes_nothing() {
        let pin = a_window(0x1000);
        let mut nesting = Nesting::new();

        nesting.settle(Some(pin), 0x1234);
        for _ in 0..10 {
            let (parent, orphaned) = settle_parents(nesting.parent(), Some(pin), |_| true);
            assert!(!orphaned);
            assert_eq!(
                nesting.settle(parent, 0x1234),
                NestingChange::Stayed,
                "the same parent asked for again is nothing to do"
            );
        }
        assert_eq!(nesting.parent(), Some(pin));
    }
}

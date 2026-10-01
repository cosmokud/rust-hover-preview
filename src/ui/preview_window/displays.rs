//! Which display a point is on, and what that display has room for.
//!
//! Every box this app puts on the screen is placed against one display's work area and drawn at
//! one display's scale, and both of those were two Win32 calls made from wherever the box
//! happened to be computed. That is why placement was the one decision that could not be tested
//! at all: `compute_mouse_layout` takes its bounds as a parameter and has a hundred and thirty
//! geometry tests behind it, while the three lines above it that choose the bounds reach
//! straight into the display driver. `clamp_pinned_box` — a pure function with six callers, the
//! one that stops a dragged pin being stranded half off a monitor — was untestable for exactly
//! this reason and no other: it worked the display out for itself, inside its own body, so the
//! only way to ask it anything was to attach a screen.
//!
//! So this module is the seam those two lines go through, and the decision is *above* it rather
//! than inside it: the trait knows how to name a display and `display_at` decides what happens
//! where it cannot. That ordering is what makes the branch testable. A fallback that lived in
//! the Win32 adapter would be a branch no recorder could reach, because a recorder is the adapter
//! standing in for it — so the adapter answers `None` and the decision, which is shared, places
//! the preview on the primary display instead. This is the same arrangement `pin_window` uses,
//! and for the same reason: `PinWindow` answers `HWND(0)` for a window that is not there and
//! `end_pin` decides what a road out of a pin may do about it.
//!
//! Two adapters, so the seam is real rather than hypothetical. `Desktop` is this machine's own
//! displays, and it is the one the loop and the window procedure both call through `DESKTOPS`.
//! `RecordedDisplays` stands in for a machine with a display to name and for a machine with none
//! to name at all, and remembers what it was asked and in what order — which is what lets a
//! placement be asserted as *the preview was kept on the display the hand was on*, rather than
//! as a rectangle that happened to come out right.
//!
//! It is a module of its own rather than more of `PinWindow` because the two seams are about
//! different things and merging them would make a lie of both. `PinWindow` is the pin's own
//! window and what a road out of a pin's lifecycle may ask of it; this is the desk the window is
//! standing on, asked about by a hover before any window exists at all. Putting `work_area` on a
//! trait whose subject is one pin's window would make every one of the twenty-odd callers here a
//! caller of the pin's window, which is the shape the review's own note warns about: widening a
//! seam with operations nothing on the pin calls is the speculative half of the decision.
//!
//! Nothing else moved here. A display's *arrangement* — which monitor is left of which, what the
//! taskbar has taken off the bottom — is a question about a machine and is answered by Windows;
//! there is no decision in it to test. What is decided is *which of those displays this point
//! belongs to*, and that is the two functions below.

use once_cell::sync::Lazy;
use std::sync::Mutex;
use windows::Win32::Foundation::{POINT, RECT};
use windows::Win32::Graphics::Gdi::{
    GetMonitorInfoW, MonitorFromPoint, HMONITOR, MONITORINFO, MONITOR_DEFAULTTONEAREST,
};
use windows::Win32::UI::HiDpi::{GetDpiForMonitor, MDT_EFFECTIVE_DPI};
use windows::Win32::UI::WindowsAndMessaging::{
    GetSystemMetrics, SystemParametersInfoW, SM_CXVIRTUALSCREEN, SM_CYVIRTUALSCREEN,
    SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN, SPI_GETWORKAREA, SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS,
};

use super::ScreenBounds;

/// The scale Windows draws at where nothing has told us otherwise.
///
/// A display whose scale cannot be read is a display at 100%, which is what the driver reports
/// for a monitor it has no `MDT_EFFECTIVE_DPI` for, and it is also the number every margin in
/// this tree is written against. Guessing high would draw a preview that overflows the display
/// it was placed on; guessing low would draw one that is needlessly small. The baseline is the
/// only answer that is wrong the same way everywhere.
const BASELINE_DPI: u32 = 96;

/// One display, as far as anything placing a box on the screen is concerned.
///
/// A work area and a scale, and nothing about the display's own identity or arrangement: a
/// placement cannot ask which monitor this is, and the reason it used to be able to is the seam
/// this module draws.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) struct Display {
    /// The room this display has, in the pixels of the machine rather than of the scale: the
    /// work area, which is the display with its taskbar and its own edges taken off it.
    pub(super) work_area: ScreenBounds,
    /// How many machine pixels this display draws one of its own logical pixels at, which every
    /// margin a placement is written in is scaled by.
    pub(super) dpi: u32,
}

/// What a caller is allowed to ask of the desk.
///
/// Two questions and nothing else, because they are the only two a decision about where a box
/// goes has ever needed answered, and both were answered by a `MonitorFromPoint` in the middle
/// of whatever function happened to be placing something.
///
/// `nearest` answers `None` rather than a default, and that is the whole of what makes the
/// fallback above the seam testable: a machine whose driver cannot name the display under a point
/// is a machine a recorder can stand in for, and a trait that invented a rectangle there would
/// make that unreachable from any side but the real adapter's.
pub(super) trait Displays {
    /// The display nearest to a point, or nothing where the machine cannot name one.
    ///
    /// *Nearest* rather than *containing*, and it is deliberate: a point between two displays,
    /// or beyond both, still belongs to one of them, and a layout that gave up there would be a
    /// layout that had to be given a rule for a seam nobody can see.
    fn nearest(&self, x: i32, y: i32) -> Option<Display>;

    /// One display's room for a point no display could be named for.
    ///
    /// The primary display, and not the union of them: the whole virtual screen is every display
    /// at once, and a preview sized to that straddles the seam between two of them, which is
    /// the thing anchoring a layout to a display is for.
    fn primary(&self) -> ScreenBounds;
}

/// The display a point belongs to, and what may be done where it cannot be named.
///
/// The decision, and it is the whole of this module that is one: a point with a display is placed
/// on that display, and a point without one is placed on the primary one at the baseline scale —
/// somewhere it is wholly visible, rather than somewhere it is not. That is what the two
/// adapters above share, which is why it is here and not in either of them.
pub(super) fn display_at(displays: &dyn Displays, x: i32, y: i32) -> Display {
    displays.nearest(x, y).unwrap_or(Display {
        work_area: displays.primary(),
        dpi: BASELINE_DPI,
    })
}

/// This machine's own displays, which is what every caller that is not a test asks.
///
/// A `const` rather than a `static`: the adapter holds nothing, and the one piece of state on
/// this side of the seam — the cache of the last display's scale — belongs to the adapter that
/// reads it, so a caller that wants a different machine passes a different `&dyn Displays`
/// rather than swapping a global (see `RecordedDisplays`).
pub(super) const DESKTOPS: Desktop = Desktop;

/// The work area of the display a point is on, for the two thirds of the callers that want only
/// this half of it.
pub(super) fn work_area_at(x: i32, y: i32) -> ScreenBounds {
    display_at(&DESKTOPS, x, y).work_area
}

/// The scale of the display a point is on, for the callers that size rather than place.
pub(super) fn dpi_at(x: i32, y: i32) -> u32 {
    display_at(&DESKTOPS, x, y).dpi
}

/// This machine's displays, asked through Windows.
///
/// The first of the two adapters, and a unit struct for the reason `pin_window::Win32PinWindow`
/// is one: there is one of these and it is never built, and a value that carries no state
/// cannot be a state the two callers disagree about.
#[derive(Clone, Copy)]
pub(super) struct Desktop;

impl Desktop {
    /// The work area of the primary display, for a point no display could be named for.
    ///
    /// `SystemParametersInfo` first and the whole virtual screen only if it fails, because the
    /// two answer different questions: the first is one display with its taskbar taken off, the
    /// second is every display with no taskbar taken off any of them. The second is a worse
    /// answer than no answer — a preview sized to it straddles two monitors — which is why it is
    /// the fallback rather than the first thing tried.
    fn primary_work_area(&self) -> ScreenBounds {
        unsafe {
            let mut work = RECT::default();
            if SystemParametersInfoW(
                SPI_GETWORKAREA,
                0,
                Some(&mut work as *mut RECT as *mut core::ffi::c_void),
                SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS(0),
            )
            .is_ok()
            {
                return ScreenBounds {
                    left: work.left,
                    top: work.top,
                    right: work.right,
                    bottom: work.bottom,
                };
            }
        }

        self.virtual_screen()
    }

    /// Every display in one rectangle, for the fallback that cannot do better than the primary
    /// display and find it missing.
    fn virtual_screen(&self) -> ScreenBounds {
        unsafe {
            let left = GetSystemMetrics(SM_XVIRTUALSCREEN);
            let top = GetSystemMetrics(SM_YVIRTUALSCREEN);
            let width = GetSystemMetrics(SM_CXVIRTUALSCREEN).max(1);
            let height = GetSystemMetrics(SM_CYVIRTUALSCREEN).max(1);

            ScreenBounds {
                left,
                top,
                right: left + width,
                bottom: top + height,
            }
        }
    }

    /// The scale of the display a point is on.
    ///
    /// The display is asked and not the window, which is the whole of what this question is
    /// for: a window carries the scale its own process was told about — a UWP one can answer a
    /// scale that is not the display's at all — while the display under the point is one
    /// question with one answer, whatever is drawn on it.
    ///
    /// The answer is kept per display, because the scale of a display does not change while it
    /// is the display: the pointer's own probe asks this every tick and every layout asks it
    /// again, and a pointer that has not crossed to another display is answered from here rather
    /// than by asking the DPI interface for a number that cannot have moved. A display whose
    /// scale does change — a monitor switched to another scaling — is a display whose handle is
    /// the same and whose answer is not, so the cache holds one display: the next one named is
    /// asked about, and the one after that is asked again.
    fn scale_of(&self, monitor: isize) -> u32 {
        if let Ok(cached) = MONITOR_DPI.lock() {
            if let Some((cached_monitor, dpi)) = *cached {
                if cached_monitor == monitor {
                    return dpi;
                }
            }
        }

        let mut dpi = BASELINE_DPI;
        let mut dpi_x = 0u32;
        let mut dpi_y = 0u32;
        unsafe {
            if GetDpiForMonitor(
                monitor_handle(monitor),
                MDT_EFFECTIVE_DPI,
                &mut dpi_x,
                &mut dpi_y,
            )
            .is_ok()
                && dpi_x > 0
            {
                dpi = dpi_x;
            }
        }

        if let Ok(mut cached) = MONITOR_DPI.lock() {
            *cached = Some((monitor, dpi));
        }

        dpi
    }
}

impl Displays for Desktop {
    fn nearest(&self, x: i32, y: i32) -> Option<Display> {
        let monitor = unsafe { MonitorFromPoint(POINT { x, y }, MONITOR_DEFAULTTONEAREST) };
        if monitor.is_invalid() {
            return None;
        }

        let mut work = None;
        unsafe {
            let mut info = MONITORINFO {
                cbSize: std::mem::size_of::<MONITORINFO>() as u32,
                ..Default::default()
            };
            if GetMonitorInfoW(monitor, &mut info).as_bool() {
                let area = info.rcWork;
                work = Some(ScreenBounds {
                    left: area.left,
                    top: area.top,
                    right: area.right,
                    bottom: area.bottom,
                });
            }
        }

        // A display that was named but whose work area could not be read still has a scale, and
        // the two used to be asked about separately: the room fell back to the primary display's
        // and the scale was read from this display regardless. Keeping that here rather than
        // pushing the whole thing up to `display_at` is what keeps the fallback *only* for a
        // display that could not be named at all — a much narrower branch, and the one the
        // decision above exists for.
        Some(Display {
            work_area: work.unwrap_or_else(|| self.primary_work_area()),
            dpi: self.scale_of(monitor.0 as isize),
        })
    }

    fn primary(&self) -> ScreenBounds {
        self.primary_work_area()
    }
}

/// The last display asked about and the scale it answered with — one display's worth, for the
/// reason `Desktop::scale_of` gives.
static MONITOR_DPI: Lazy<Mutex<Option<(isize, u32)>>> = Lazy::new(|| Mutex::new(None));

/// The handle a display's scale is read from, which is the driver's rather than a window's.
///
/// A handle is a raw pointer, which is why it is an `isize` here as it is everywhere else in this
/// window layer (see `pin_window::PinWindow`) and never an `HWND`: nothing in this file may
/// dereference it.
fn monitor_handle(monitor: isize) -> HMONITOR {
    HMONITOR(monitor as *mut _)
}

/// A desk that remembers what it was asked and answers with whatever it has been given.
///
/// The second adapter, and the reason this is a seam rather than a signature. It is told which
/// display the machine has, or that it has none to name, so the branch a live desktop almost
/// never takes — the driver refusing to name the display under the pointer, which is what
/// happens on a machine mid-display-change — is reachable at all, and the places a decision
/// *looks* for a display are visible as a list rather than inferred from the box that came out.
///
/// It is not a mock of Windows and does not pretend to be one: it answers with a rectangle and a
/// scale, which is all any caller of `Displays` ever asked for, and it refuses to invent either
/// where the machine it stands in for has neither.
#[cfg(test)]
pub(super) struct RecordedDisplays {
    /// The display this machine has, or nothing for a machine that cannot name one.
    named: Option<ScreenBounds>,
    /// The scale that display answers with, and the primary's room.
    named_dpi: u32,
    primary: ScreenBounds,
    asks: Mutex<Vec<DisplayAsk>>,
}

/// One question asked of the desk, and the point it was asked about.
#[cfg(test)]
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum DisplayAsk {
    /// Where the display nearest this point is, and its scale.
    Nearest((i32, i32)),
    /// The primary display's room, asked for because the nearest one could not be named.
    Primary,
}

#[cfg(test)]
impl RecordedDisplays {
    /// A machine with one display at `(left, top)`, drawn at `dpi`.
    pub(super) fn one_display(bounds: ScreenBounds, dpi: u32) -> Self {
        Self {
            named: Some(bounds),
            named_dpi: dpi,
            primary: bounds,
            asks: Mutex::new(Vec::new()),
        }
    }

    /// A machine whose driver will not name the display under a point, and whose primary
    /// display's room is the only display there is to answer with.
    pub(super) fn no_display_to_name(primary: ScreenBounds) -> Self {
        Self {
            named: None,
            named_dpi: BASELINE_DPI,
            primary,
            asks: Mutex::new(Vec::new()),
        }
    }

    /// Everything that was asked of the desk, in the order it was asked.
    pub(super) fn asks(&self) -> Vec<DisplayAsk> {
        self.asks
            .lock()
            .map(|asks| asks.clone())
            .unwrap_or_default()
    }

    /// Just the points a display was asked about, in order, for a caller that wants to say
    /// *which point* a decision anchored a box to and has nothing to say about the fallback.
    ///
    /// The narrow form because the anchor is the question nearly every caller of the seam is
    /// asking — a window is kept on the display its middle is on, a preview is placed on the
    /// display the pointer is on — and a test that had to match against `DisplayAsk::Nearest` to
    /// read a coordinate would be asserting on the adapter's vocabulary rather than on the
    /// decision.
    pub(super) fn points_asked_about(&self) -> Vec<(i32, i32)> {
        self.asks()
            .into_iter()
            .filter_map(|ask| match ask {
                DisplayAsk::Nearest(point) => Some(point),
                DisplayAsk::Primary => None,
            })
            .collect()
    }

    fn record(&self, ask: DisplayAsk) {
        if let Ok(mut asks) = self.asks.lock() {
            asks.push(ask);
        }
    }
}

#[cfg(test)]
impl Displays for RecordedDisplays {
    fn nearest(&self, x: i32, y: i32) -> Option<Display> {
        self.record(DisplayAsk::Nearest((x, y)));
        self.named.map(|work_area| Display {
            work_area,
            dpi: self.named_dpi,
        })
    }

    fn primary(&self) -> ScreenBounds {
        self.record(DisplayAsk::Primary);
        self.primary
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The display most tests stand on: 1000 by 800 at its own top-left corner, at 96 dpi.
    fn the_desk() -> ScreenBounds {
        ScreenBounds {
            left: 0,
            top: 0,
            right: 1000,
            bottom: 800,
        }
    }

    /// A display to the right of the first, so the two can be told apart by what they answer.
    fn the_second_display() -> ScreenBounds {
        ScreenBounds {
            left: 1000,
            top: 0,
            right: 2000,
            bottom: 800,
        }
    }

    /// A point is placed on the display it is on, and the display it is on is the one nearest
    /// it — not the primary, and not the union of them.
    ///
    /// The union is the failure this rules out, and it is invisible to a test that only ever
    /// runs on one monitor: a preview sized to the whole virtual screen straddles the seam
    /// between two displays, which is the thing anchoring a layout to a display is for, and on a
    /// single-display machine the two rectangles are the same rectangle.
    #[test]
    fn a_point_is_placed_on_the_display_it_is_on() {
        for (point, expected) in [
            ((10, 10), the_desk()),
            ((1500, 400), the_second_display()),
            ((10, 700), the_desk()),
            ((1990, 790), the_second_display()),
        ] {
            let desk = RecordedDisplays::one_display(expected, 96);
            assert_eq!(
                display_at(&desk, point.0, point.1).work_area,
                expected,
                "({point:?}) belongs to the display it is on"
            );
        }
    }

    /// A point the machine cannot place is placed on the primary display, and the primary
    /// display is asked for rather than a rectangle invented here.
    ///
    /// This is the branch a live desktop almost never takes — a driver that will not name the
    /// display under the pointer, which is what a machine mid-display-change looks like — and it
    /// is the whole of why the decision sits above the seam rather than inside the adapter that
    /// would have to refuse. Note the order the two asks are recorded in: the display is looked
    /// for first, and the fallback is only reached because that came back empty.
    #[test]
    fn a_point_no_display_can_be_named_for_is_placed_on_the_primary_one() {
        let desk = RecordedDisplays::no_display_to_name(the_desk());

        assert_eq!(
            display_at(&desk, 400, 400),
            Display {
                work_area: the_desk(),
                dpi: BASELINE_DPI,
            },
            "somewhere it is wholly visible, rather than somewhere it is not"
        );
        assert_eq!(
            desk.asks(),
            vec![
                DisplayAsk::Nearest((400, 400)),
                DisplayAsk::Primary,
            ],
            "the display is looked for first and the fallback is only reached because it was not found"
        );
    }

    /// The scale of a display is asked together with its room, so a caller cannot place a box on
    /// one display at the margins of another.
    ///
    /// These were two functions on two calls — `monitor_bounds_from_point` and
    /// `monitor_dpi_from_point` — and a box sized by one and clamped by the other is a box whose
    /// margins are in the pixels of a different display than its edges are.
    #[test]
    fn the_room_and_the_scale_come_from_the_same_display() {
        let desk = RecordedDisplays::one_display(the_second_display(), 144);

        assert_eq!(
            display_at(&desk, 1200, 300),
            Display {
                work_area: the_second_display(),
                dpi: 144,
            },
            "both halves are the display the point is on"
        );
    }

    /// The scale is remembered for one display, and asked about again for the next one.
    ///
    /// A display's scale does not change while it is the display, and the pointer asks this
    /// every tick: the answer is kept so a pointer that has not crossed to another display is
    /// not asked the DPI interface the same number every sixteen milliseconds. It is kept for
    /// *one* display deliberately — a monitor switched to another scaling has the same handle
    /// and a different answer, so the next display named is asked about afresh rather than the
    /// second question after it being answered by the first one's value.
    ///
    /// Asserted against the machine's own displays rather than a recorder's, because the cache
    /// is inside the adapter and a recorder that kept one would be testing a second copy of the
    /// memo rather than the one the loop runs on. That is the one thing here that is *not* a
    /// decision, and it is asserted only as far as it can be without inventing a display.
    #[test]
    fn a_display_s_scale_is_remembered_for_that_display_and_no_other() {
        if let Ok(mut cached) = MONITOR_DPI.lock() {
            *cached = None;
        }

        let first = DESKTOPS.scale_of(0x1);
        assert_eq!(
            DESKTOPS.scale_of(0x1),
            first,
            "the same display answered from the cache rather than from the driver"
        );
        let cached = MONITOR_DPI.lock().map(|cached| *cached).unwrap_or(None);
        assert_eq!(
            cached.map(|(_, dpi)| dpi),
            Some(first),
            "and the cached answer is the one that was asked for"
        );

        DESKTOPS.scale_of(0x2);
        let cached = MONITOR_DPI.lock().map(|cached| *cached).unwrap_or(None);
        assert_eq!(
            cached.map(|(monitor, _)| monitor),
            Some(0x2),
            "a display named after the one that was cached is asked about rather than read out"
        );

        if let Ok(mut cached) = MONITOR_DPI.lock() {
            *cached = None;
        }
    }

    /// A machine that cannot name a display for the point can still say what its primary
    /// display has room for, and the two are asked separately.
    ///
    /// The fallback is one display and not the union: `SM_XVIRTUALSCREEN` is every display at
    /// once, so a preview sized to it is a preview that straddles two of them. That is why the
    /// adapter answers a *primary* and not a "screen", and why the union is only what a
    /// machine whose primary cannot be read is given.
    #[test]
    fn the_fallback_is_one_display_and_not_the_union_of_them() {
        let desk = RecordedDisplays::no_display_to_name(the_desk());
        let Display { work_area, .. } = display_at(&desk, -5_000, -5_000);

        assert_eq!(work_area, the_desk());
        assert!(
            work_area
                != ScreenBounds {
                    left: 0,
                    top: 0,
                    right: 2000,
                    bottom: 800,
                },
            "a point no display can be named for is placed on one display rather than on the union"
        );
    }
}

//! The page a pin frames, drawn into the pin by the Shell's own thumbnail of the page's window.
//!
//! **What a page's band has to be is a picture, and the pin's window is a layered one.** The band is a
//! hole in the pin's surface, so the page's window is what fills it — and a window of somebody else's
//! that is *in front of* the pin is a band that lags a drag (it is carried by a cross-process
//! `SetWindowPos` that lands after the pixels do), a window that eats the press the pin's own drag needs,
//! and a window that raises itself over the pin's band. That is the film road's problem, and this is the
//! film road's answer, one picture at a time: the page is a *source* on the desktop — invisible at one
//! part in 255 and click-through, lowered once, placed at the band once and only moved — and the pin
//! draws it through `DwmRegisterThumbnail` of the source's client area, into the band.
//!
//! **And the registration is what makes the pin's taskbar hover a page rather than a hole.** The Shell
//! pictures a window from what the compositor composed for it, so a pin whose band is a hole is a peek
//! with nothing in it. A registration composed *into* the pin is part of what is composed for it, so the
//! peek is the page — live while the pin is up, and the last thing the Shell composed for it once it is
//! put away. Nothing here takes a picture of the screen or asks the Shell for an iconic thumbnail: those
//! roads are the ones a layered window answers with black (see `pin_taskbar_mirror`, and the film road's
//! table of them in `ARCHITECTURE.md`).
//!
//! **The road follows `pin_film_thumbnail`'s rules and has this one difference: the destination is the
//! pin, not a host of its own.** The pin's window is the caption and the transport as well as the band,
//! so the destination rectangle is the band's own rect inside the pin's client, and everything else the
//! pin draws is untouched by it. The source's client is a constant for as long as it is the pin's
//! picture (the one placement, see `webview_preview`'s `put_window`), so a frame can only change with
//! the pin's band — which on a drag changes in *screen* coordinates only, and so is the same rectangle
//! in the pin's client: **a drag re-points nothing at all**.

use super::*;

use std::sync::Mutex;

use crate::ui::preview_window::pin_trace::note_pin_trace;

use windows::Win32::Foundation::{HWND, RECT};
use windows::Win32::Graphics::Dwm::{
    DwmRegisterThumbnail, DwmUnregisterThumbnail, DwmUpdateThumbnailProperties,
    DWM_THUMBNAIL_PROPERTIES, DWM_TNP_RECTDESTINATION, DWM_TNP_SOURCECLIENTAREAONLY,
    DWM_TNP_VISIBLE,
};

/// The band a registration is pointed at: `left, top, width, height` in the client coordinates of the
/// window the Shell draws into.
type Band = (i32, i32, i32, i32);

/// The registration standing, the two windows it was made for and the rectangle pair it was last
/// handed: `NOTHING` while no page is drawn.
static PAGE_THUMBNAIL: Mutex<Held> = Mutex::new(Held::NOTHING);

#[derive(Clone, Copy, PartialEq, Eq)]
struct Held {
    /// The pin window the Shell draws into, and the page window it draws.
    pin: isize,
    page: isize,
    registration: isize,
    /// The destination last handed, in the pin's client coordinates, or nothing where none stands: a
    /// placement that asks for the same one is the Shell asked nothing (see `the_page_was_placed`).
    shown: Option<Band>,
}

impl Held {
    const NOTHING: Self = Self {
        pin: 0,
        page: 0,
        registration: 0,
        shown: None,
    };
}

/// Draw the pin's page into it, for `content` — the media band the pin is showing, in screen
/// coordinates — or let the registration go where the pin has no page in its band.
///
/// It is asked wherever the band moves or changes (see `pin_lifecycle::place_pinned_siblings`) and it is
/// cheap to ask: a page that is not the pin's picture is a registration let go of, and a pair that did not
/// change hands the Shell nothing at all — **which is every step of a drag**, the destination being the
/// band in the pin's own client coordinates rather than on the screen.
pub(super) fn the_page_was_placed(content: Option<ScreenRegion>) {
    let Some((pin, page, destination)) = content.and_then(the_pair) else {
        let_registration_go("the pin has no page in its band");
        return;
    };

    // The band handed, the pin's client it is a rectangle in and the two windows it was made for: a
    // band drawn off its source is one of these three disagreeing with the others (see the module note).
    let client = unsafe { pin_film_thumbnail::client_size(HWND(pin as *mut _)) };
    note_pin_trace!(
        "page band: destination {}x{}+{}+{} pin client {client:?} pin={:#x} page={:#x}",
        destination.2,
        destination.3,
        destination.0,
        destination.1,
        pin,
        page
    );

    let Ok(mut held) = PAGE_THUMBNAIL.lock() else {
        return;
    };

    // **A registration is for one pair of windows**, so another pin, or a page the engine has replaced,
    // is the picture given up and made again rather than a rectangle re-pointed at a window the Shell
    // was never asked about.
    if held.registration != 0 && (held.pin != pin || held.page != page) {
        let registration = held.registration;
        *held = Held::NOTHING;
        let _ = unsafe { DwmUnregisterThumbnail(registration) };
    }
    held.pin = pin;
    held.page = page;
    if held.registration == 0 {
        match unsafe { DwmRegisterThumbnail(HWND(pin as *mut _), HWND(page as *mut _)) } {
            Ok(registration) => held.registration = registration,
            Err(error) => {
                // The Shell refuses a source that has made its window and not yet presented one, and nothing
                // here asks a second time: the next placement of the band does (see `the_page_was_placed`).
                *held = Held::NOTHING;
                note_pin_trace!("page band: register refused {error:?}");
                return;
            }
        }
    }

    if held.shown == Some(destination) {
        note_pin_trace!(
            "page band: unchanged {}x{}+{}+{}",
            destination.2,
            destination.3,
            destination.0,
            destination.1
        );
        return;
    }

    let (x, y, width, height) = destination;
    let properties = DWM_THUMBNAIL_PROPERTIES {
        dwFlags: DWM_TNP_RECTDESTINATION | DWM_TNP_SOURCECLIENTAREAONLY | DWM_TNP_VISIBLE,
        rcDestination: RECT {
            left: x,
            top: y,
            right: x + width,
            bottom: y + height,
        },
        rcSource: RECT::default(),
        opacity: 255,
        fVisible: true.into(),
        // The page is its client area and nothing around it: the window is a popup a browser draws in,
        // and a frame sampled with it would be a page drawn off its own box by the frame's width.
        fSourceClientAreaOnly: true.into(),
    };
    if let Err(error) = unsafe { DwmUpdateThumbnailProperties(held.registration, &properties) } {
        let registration = held.registration;
        *held = Held::NOTHING;
        let _ = unsafe { DwmUnregisterThumbnail(registration) };
        note_pin_trace!("page band: update refused {error:?}");
        return;
    }
    held.shown = Some(destination);
    note_pin_trace!("page band: shown {width}x{height}+{x}+{y}");
}

/// Let the registration go, keeping the pin and the page it was made for: what a pin going down, a pin
/// swapping the page for a picture, and a pin whose band no longer holds a page all owe the Shell.
pub(super) fn forget_the_page_thumbnail_because(reason: &str) {
    let_registration_go(reason);
}

fn let_registration_go(reason: &str) {
    let Ok(mut held) = PAGE_THUMBNAIL.lock() else {
        return;
    };
    let registration = std::mem::replace(&mut *held, Held::NOTHING).registration;
    if registration != 0 {
        let _ = unsafe { DwmUnregisterThumbnail(registration) };
        note_pin_trace!("page thumbnail let go (reason={reason})");
    }
}

/// The pin window the Shell draws into, the page window it draws, and the band's rectangle in the pin's
/// own client coordinates — or nothing where either window or the band is not there.
fn the_pair(content: ScreenRegion) -> Option<(isize, isize, Band)> {
    let page = webview_preview::showing_hwnd();
    if page == 0 {
        return None;
    }
    let pin = pin_real_window::real_pin_window()?;

    // The pin's client origin and area, which are the coordinates the Shell is handed the rectangle in: a
    // layered pin is drawn where its client is, so the band is the same rectangle in either coordinates —
    // and the *difference* is what makes a drag re-point nothing (see the module note).
    let origin = unsafe { pin_film_thumbnail::client_origin(pin) }?;
    let client = unsafe { pin_film_thumbnail::client_size(pin) }?;

    Some((
        pin.0 as isize,
        page,
        the_band_in_the_client(content, origin, client),
    ))
}

/// A band, in the client coordinates of the window that draws it: the pin's own rectangle of itself, cut
/// to the client so that a band a display change has left stale is a rectangle inside the window rather
/// than one reaching past the surface the Shell samples.
fn the_band_in_the_client(content: ScreenRegion, origin: (i32, i32), client: (i32, i32)) -> Band {
    let (client_width, client_height) = (client.0.max(1), client.1.max(1));
    let left = (content.0 - origin.0).clamp(0, client_width);
    let top = (content.1 - origin.1).clamp(0, client_height);
    let right = (content.2 - origin.0).clamp(left, client_width);
    let bottom = (content.3 - origin.1).clamp(top, client_height);

    (left, top, (right - left).max(1), (bottom - top).max(1))
}

#[cfg(test)]
mod tests {
    use super::*;

    /// **The band is the same rectangle in the pin's client coordinates wherever the pin stands**, which
    /// is the whole of why a drag costs the Shell nothing: the band and the client move together. And a
    /// band is never a rectangle reaching past the client the Shell samples (see `the_band_in_the_client`).
    #[test]
    fn a_band_is_the_same_rectangle_wherever_its_pin_stands() {
        let content = (500, 400, 1140, 700);
        let client = (500, 400, 640, 300);
        let placed = the_band_in_the_client(content, (500, 400), (client.2, client.3));
        assert_eq!(placed, (0, 0, 640, 300), "the band is the pin's own client");

        // The same pin carried a hundred pixels each way: the band and the client moved together, so
        // the pair handed to the Shell did not move at all.
        assert_eq!(
            the_band_in_the_client((600, 500, 1240, 800), (600, 500), (client.2, client.3)),
            placed,
            "a drag re-points nothing"
        );

        // A pin's caption above the band: the band is the rectangle below it, in the client's units.
        assert_eq!(
            the_band_in_the_client((500, 428, 1140, 728), (500, 400), (640, 300)),
            (0, 28, 640, 272),
            "the chrome above the band is the pin's own pixels"
        );

        // A band the geometry has left outside the window is cut to the client rather than handed to the
        // Shell as a rectangle sampled past the surface it is sampling.
        assert_eq!(
            the_band_in_the_client((560, 400, 1180, 712), (500, 400), (640, 300)),
            (60, 0, 580, 300),
            "a band reaching past the client is cut to it"
        );
    }

    /// A registration is for one pair of windows and one band: a band already handed is the Shell asked
    /// nothing, and a pair that changed is the picture given up rather than a rectangle re-pointed.
    #[test]
    fn a_pair_is_re_pointed_only_where_it_changed() {
        let held = Held {
            pin: 1,
            page: 2,
            registration: 3,
            shown: Some((0, 28, 640, 272)),
        };

        assert_eq!(held.shown, Some((0, 28, 640, 272)));
        assert_eq!((held.pin, held.page), (1, 2));
        assert!(
            held.shown != Some((0, 28, 640, 271)),
            "a band that changed is a pair to hand over again"
        );
    }
}

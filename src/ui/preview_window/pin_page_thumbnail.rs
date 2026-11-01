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
//! so the destination rectangle is the band's own rows inside the pin's client, and everything else the
//! pin draws is untouched by it. The source's client follows the band — the Shell fits the thumbnail of
//! that client into the band, so a client of another shape is a page with bars of the desktop around it
//! (see `webview_preview`'s `put_window`) — while a frame changes with the *pin's* band alone: on a drag
//! that is a rectangle of the same size in the same client, so **a drag re-points nothing at all**.

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

/// Draw the pin's page into it — or let the registration go where the pin has no page in its band.
///
/// It is asked wherever the band moves or changes (see `pin_lifecycle::place_pinned_siblings`) and it is
/// cheap to ask: a page that is not the pin's picture is a registration let go of, and a pair that did not
/// change hands the Shell nothing at all — **which is every step of a drag**, the destination being the
/// pin's own band rather than anything on the screen.
pub(super) fn the_page_was_placed() {
    let Some((pin, page, destination)) = the_pair() else {
        let_registration_go("the pin has no page in its band");
        return;
    };

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
        let Ok(registration) =
            (unsafe { DwmRegisterThumbnail(HWND(pin as *mut _), HWND(page as *mut _)) })
        else {
            // The Shell refuses a source that has made its window and not yet presented one, and nothing
            // here asks a second time: the next placement of the band does (see `the_page_was_placed`).
            *held = Held::NOTHING;
            return;
        };
        held.registration = registration;
    }

    if held.shown == Some(destination) {
        return;
    }

    let (x, y, width, height) = destination;
    note_pin_trace!(
        "page band {}x{}+{}+{} in the pin {:?}",
        width,
        height,
        x,
        y,
        HWND(pin as *mut _)
    );
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
    if unsafe { DwmUpdateThumbnailProperties(held.registration, &properties) }.is_err() {
        let registration = held.registration;
        *held = Held::NOTHING;
        let _ = unsafe { DwmUnregisterThumbnail(registration) };
        return;
    }
    held.shown = Some(destination);
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
///
/// **The destination is the band the pin is drawing, read off the window it is drawn in**, not a rectangle
/// worked out from where the page stands on the screen: the Shell fits this thumbnail of the page's client
/// into it, so a band that was cut to fit a window that had not resized yet would be a page drawn small
/// inside a window of the desktop's own alpha — which is the desktop seen through it. What the pin's state
/// says the band is, and the band the window is, are one answer here or they are said out loud (see
/// `the_band_in_the_client`).
fn the_pair() -> Option<(isize, isize, Band)> {
    let page = webview_preview::showing_hwnd();
    if page == 0 {
        return None;
    }
    let pin = pin_real_window::real_pin_window()?;

    // The pin's client area, which is the surface the Shell samples and the coordinates it is handed the
    // rectangle in — the same rectangle the paint fills from the band's own rows.
    let client = unsafe { pin_film_thumbnail::client_size(pin) }?;
    let rows = pin_lifecycle::pinned_band_rows_of_the_client(client)?;

    Some((pin.0 as isize, page, the_band_in_the_client(client, rows)))
}

/// A band, in the client coordinates of the window that draws it: the rows that window's own paint leaves
/// for the media, which is the whole of what the page stands in.
///
/// It cannot be smaller than the band or reach past it, since it is the band — a pin's window is fitted to
/// its media, so there is no third rectangle between the two for a stale one to be cut to.
fn the_band_in_the_client(client: (i32, i32), rows: (i32, i32)) -> Band {
    let (top, height) = rows;
    (0, top, client.0.max(1), height.max(1))
}

#[cfg(test)]
mod tests {
    use super::*;

    /// **The destination is the pin's own band, so it is the client's rows wherever the pin stands and
    /// whatever the pin was before**: a drag changes no side and so re-points nothing, and a resize is a
    /// rectangle the Shell is handed again — never a band cut to fit a window that had not caught up.
    #[test]
    fn a_band_is_the_pins_own_rows_wherever_it_stands() {
        let client = (640, 300);
        // A caption above the band and nothing below it: the band is the rows under the caption.
        let placed = the_band_in_the_client(client, pinned_band_rows(client.1, 28, 0, false));
        assert_eq!(placed, (0, 28, 640, 272), "the band is below the caption");

        // The same client on a pin carried a hundred pixels each way: the band is the same rectangle,
        // which is the whole of why a step of a drag asks the Shell nothing.
        assert_eq!(
            the_band_in_the_client(client, pinned_band_rows(client.1, 28, 0, false)),
            placed,
            "a drag re-points nothing"
        );

        // A window the hand pulled out: the band is the new client's own rows, edge to edge.
        let resized = (900, 500);
        assert_eq!(
            the_band_in_the_client(resized, pinned_band_rows(resized.1, 28, 0, false)),
            (0, 28, 900, 472),
            "a resize is a band of the client it was made"
        );

        // A chrome drawn over the media is no rows of the band's own: the whole client is the band.
        assert_eq!(
            the_band_in_the_client(client, pinned_band_rows(client.1, 28, 0, true)),
            (0, 0, 640, 300),
            "chrome over the media leaves the whole client to it"
        );

        // **And it is the paint's own answer**, which is what makes the registration and the alpha-1 claim
        // the pin paints the same rectangle: a transport bar below the band is rows neither of them has.
        for (height, caption, transport, overlay) in [
            (300, 28, 0, false),
            (300, 28, 40, false),
            (300, 28, 40, true),
        ] {
            let client = (640, height);
            let rows = pinned_band_rows(height, caption, transport, overlay);
            let (_, top, width, band) = the_band_in_the_client(client, rows);
            assert_eq!(
                (top, band),
                rows,
                "the destination is the rows the paint fills"
            );
            assert_eq!(width, client.0, "the band is the client's whole width");
        }
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

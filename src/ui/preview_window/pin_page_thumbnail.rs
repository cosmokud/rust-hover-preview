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
//! so the destination rectangle is the rows the pin's own paint leaves for the media, and everything else
//! the pin draws is untouched by it. The source's client follows those rows — the Shell fits the thumbnail
//! of that client into the destination, so a client of another shape is a page drawn at a shape the band is
//! not (see `webview_preview`'s `put_window`) — and a resize and a new document are both a size while a
//! move is not: a drag changes the pin's band in *screen* coordinates only, so the destination is the same
//! rectangle in the pin's client and **a drag re-points nothing at all**.
//!
//! **And nothing here is ever registered for a pin that does not frame an engine page.** The page's window
//! is the engine's for the session rather than any pin's — hidden between documents, its handle still
//! published — and the Shell answers a thumbnail of a hidden window with this app's icon, so a band that
//! took one would draw an icon into a pin of a kind that has no page at all (see `the_pair`). Every road in
//! here is reached through the pin's page (see `pin_lifecycle::engine_page_in_the_band`).

use super::*;

use std::sync::Mutex;

use crate::ui::preview_window::pin_trace::note_pin_trace;

use windows::Win32::Foundation::{HWND, RECT};
use windows::Win32::Graphics::Dwm::{
    DwmRegisterThumbnail, DwmUnregisterThumbnail, DwmUpdateThumbnailProperties,
    DWM_THUMBNAIL_PROPERTIES, DWM_TNP_RECTDESTINATION, DWM_TNP_SOURCECLIENTAREAONLY,
    DWM_TNP_VISIBLE,
};
use windows::Win32::UI::WindowsAndMessaging::IsWindowVisible;

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

/// Draw the pin's page into it — the pin's own band being the destination — or let the registration go
/// where the pin has no page in its band.
///
/// `content` is the media band the pin is showing, in screen coordinates, and the destination is not
/// worked out from it (see `the_band_in_the_client`): it is what a pin whose window and media box
/// disagree is said out loud against. **A pin with no page in its band is the registration let go and
/// nothing else**, and that gate is the whole of why this is only ever asked on the page's own road —
/// the engine's window outlives the pin's page (see the module note).
///
/// It is asked wherever the band moves or changes (see `pin_lifecycle::place_pinned_siblings`) and it is
/// cheap to ask: a page that is not the pin's picture is a registration let go of, and a pair that did not
/// change hands the Shell nothing at all — **which is every step of a drag**, the destination being the
/// pin's own band rather than anything on the screen.
pub(super) fn the_page_was_placed(content: Option<ScreenRegion>) {
    let Some((pin, page, destination)) = content.and_then(the_pair) else {
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
/// **A page window that is not on screen is no page at all.** The engine's window is the session's rather
/// than the pin's, and a hide leaves its handle published — so a registration made against a hidden one is
/// a thumbnail the Shell answers with this app's icon, which is a pin that peeks itself instead of its
/// page. That is the one fault this refuses rather than reports, and refusing costs nothing: the next
/// placement of the band asks again, exactly as it does after the `DwmRegisterThumbnail` below has failed.
fn the_pair(content: ScreenRegion) -> Option<(isize, isize, Band)> {
    let page = webview_preview::showing_hwnd();
    if page == 0 {
        return None;
    }
    if !unsafe { IsWindowVisible(HWND(page as *mut _)) }.as_bool() {
        return None;
    }
    let pin = pin_real_window::real_pin_window()?;

    // The pin's client area, which is the surface the Shell samples and the coordinates it is handed the
    // rectangle in, and the rows its own paint leaves for the media, which are the band those rows are.
    let client = unsafe { pin_film_thumbnail::client_size(pin) }?;
    let rows = pin_lifecycle::pinned_band_rows_of_the_client(client)?;

    Some((
        pin.0 as isize,
        page,
        the_band_in_the_client(content, client, rows),
    ))
}

/// A band, in the client coordinates of the window that draws it: the rows that window's own paint leaves
/// for the media, edge to edge — what the pin is drawing, and so the one rectangle the page is fitted into.
///
/// **It is never cut to a smaller rectangle**, because there is no third rectangle to cut to: the pin's
/// window is fitted to its media, so a destination clamped to the client would be a band the pin paints and
/// the page does not cover — and the rows it left out are the pin's own alpha-1 claim, which is the desktop
/// seen through it. `content` is used only to say out loud the one fault this can hide, a pin whose window
/// has not been resized for the media box it is showing: the band is the window's rows either way, so it is
/// covered rather than cut short.
fn the_band_in_the_client(content: ScreenRegion, client: (i32, i32), rows: (i32, i32)) -> Band {
    let (top, height) = rows;
    let height = height.max(1);

    if (content.2 - content.0, content.3 - content.1) != (client.0, height) {
        note_pin_trace!(
            "page band: pin window {}x{} is not its media {}x{}",
            client.0,
            client.1,
            content.2 - content.0,
            content.3 - content.1
        );
    }

    (0, top, client.0.max(1), height)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// **The destination is the pin's own band, so it is the client's rows wherever the pin stands and
    /// whatever the pin was before**: a drag changes no side and so re-points nothing, and a resize is a
    /// rectangle the Shell is handed again — never a band cut to fit a window that had not caught up.
    #[test]
    fn a_band_is_the_pins_own_rows_wherever_it_stands() {
        let content = (500, 400, 1140, 700);
        let client = (640, 300);
        // A caption above the band and no transport below it: the band is the rows under the caption.
        let placed =
            the_band_in_the_client(content, client, pinned_band_rows(client.1, 28, 0, false));
        assert_eq!(placed, (0, 28, 640, 272), "the band is below the caption");

        // The same client on a pin carried a hundred pixels each way: the band is the same rectangle,
        // which is the whole of why a step of a drag asks the Shell nothing.
        assert_eq!(
            the_band_in_the_client(content, client, pinned_band_rows(client.1, 28, 0, false)),
            placed,
            "a drag re-points nothing"
        );

        // A window the hand pulled out: the band is the new client's own rows, edge to edge.
        let resized = (900, 500);
        assert_eq!(
            the_band_in_the_client(content, resized, pinned_band_rows(resized.1, 28, 0, false)),
            (0, 28, 900, 472),
            "a resize is a band of the client it was made"
        );

        // A media box smaller than the window is a trace line and not a destination: what is handed to
        // the Shell is still the whole band, since a band cut to the box would leave the window's own
        // alpha-1 rows painting the desktop.
        assert_eq!(
            the_band_in_the_client(
                (500, 400, 760, 594),
                client,
                pinned_band_rows(client.1, 28, 0, false)
            ),
            placed,
            "a media box that disagrees with the window is not a band cut to it"
        );

        // A kind whose chrome is drawn over the media: the band is the whole client.
        assert_eq!(
            the_band_in_the_client(content, client, pinned_band_rows(client.1, 28, 0, true)),
            (0, 0, 640, 300),
            "a band drawn under the chrome is the whole client"
        );
    }

    /// **The fatal regression, as a test.** The engine's window outlives the pin's page with its handle
    /// still published, so a pin that frames no page is a pin this road must say nothing about: a
    /// registration made for one is a thumbnail of a hidden window, which the Shell answers with this
    /// app's icon, in every kind of pin's band and on every step of its drag.
    #[test]
    fn nothing_is_registered_for_a_pin_with_no_page() {
        the_page_was_placed(None);

        let held = *PAGE_THUMBNAIL
            .lock()
            .expect("the held registration is a lock of this module's own");
        assert_eq!(
            (held.pin, held.page, held.registration, held.shown),
            (0, 0, 0, None),
            "a pin with no page in its band is a registration let go of"
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

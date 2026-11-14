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
//! **The destination is the pin itself, and the caption over the band rides in a window of its own.**
//! The Shell composes a thumbnail **over its destination window's own content, child windows included**
//! (`ARCHITECTURE.md`), so nothing this app draws into a child of the pin can stand over the picture —
//! which is why the caption is a popup owned by the pin rather than a strip inside it (see
//! `pin_chrome_strips`). Cropping the source instead is not an option: the Shell ignores `DWM_TNP_RECTSOURCE`
//! here and flattens the page by the rows it was asked to cut (measured; see `ac91362`). Nothing is
//! cropped and nothing is scaled: the destination is the band in the pin's client coordinates and the
//! source is the page's whole client.
//!
//! **A drag re-points nothing at all.** The source's client is a constant for as long as it is the
//! pin's picture (the one placement, see `webview_preview`'s `put_window`), and the band is placed in the
//! pin's *client* coordinates, which move with the pin: a band that moves on the screen is the same
//! rectangle in its pin, and the Shell is handed the same pair either way.

use super::*;

use std::sync::Mutex;

use crate::ui::preview_window::pin_trace::note_pin_trace;

use windows::Win32::Foundation::HWND;
use windows::Win32::Graphics::Dwm::{
    DwmRegisterThumbnail, DwmUnregisterThumbnail, DwmUpdateThumbnailProperties,
    DWM_THUMBNAIL_PROPERTIES, DWM_TNP_RECTDESTINATION, DWM_TNP_SOURCECLIENTAREAONLY,
    DWM_TNP_VISIBLE,
};

/// The band a registration is pointed at: `left, top, width, height` in the client coordinates of the
/// pin the Shell draws into.
type Band = (i32, i32, i32, i32);

/// The registration standing, the two windows it was made for and the band it was last handed:
/// `NOTHING` while no page is drawn.
static PAGE_THUMBNAIL: Mutex<Held> = Mutex::new(Held::NOTHING);

#[derive(Clone, Copy, PartialEq, Eq)]
struct Held {
    /// The pin window the Shell draws into, and the page window it draws.
    pin: isize,
    page: isize,
    registration: isize,
    /// The band last handed, in the pin's client coordinates: a placement that asks for the same
    /// rectangle and page is the Shell asked nothing (see `the_page_was_placed`).
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
/// change hands the Shell nothing at all — **which is every step of a drag**, the destination riding the
/// pin's own client coordinates rather than the screen.
pub(super) fn the_page_was_placed(content: Option<ScreenRegion>) {
    let Some((pin, page, band)) = content.and_then(the_pair) else {
        let_registration_go("the pin has no page in its band");
        return;
    };

    let Ok(mut held) = PAGE_THUMBNAIL.lock() else {
        return;
    };

    // **A registration is for one pair of windows**, so another pin, or a page the engine has replaced,
    // is the picture given up and made again rather than a rectangle re-pointed at a window the Shell
    // was never asked about.
    if held.registration != 0 && (held.pin != pin.0 as isize || held.page != page) {
        let registration = held.registration;
        *held = Held::NOTHING;
        let _ = unsafe { DwmUnregisterThumbnail(registration) };
    }
    held.pin = pin.0 as isize;
    held.page = page;

    // **A registration made against a page that has not been drawn is a band the Shell pictures with
    // the source's own icon** — the default application pictogram, stretched — until the page presents
    // something. So the Shell is not asked until the page is on the screen (see
    // `pin_lifecycle::the_pins_page_is_on_the_screen`).
    //
    // A registration already standing is never dropped for this: a walk between two pages keeps the page
    // it is showing up until the next one commits, exactly as the PDF viewer's walk does (see
    // `engines::pdf_viewer::a_walk_between_documents_keeps_the_viewer`).
    if !registration_is_owed(
        held.registration == 0,
        pin_lifecycle::the_pins_page_is_on_the_screen(),
    ) {
        note_pin_trace!("page band: waiting for the page before it is registered");
        return;
    }

    if held.registration == 0 {
        match unsafe { DwmRegisterThumbnail(pin, HWND(page as *mut _)) } {
            Ok(registration) => held.registration = registration,
            Err(error) => {
                // The Shell refuses a source that has made its window and not yet presented one, and
                // nothing here asks a second time: the next placement of the band does (see above).
                *held = Held::NOTHING;
                note_pin_trace!("page band: register refused {error:?}");
                return;
            }
        }
    }

    // A band already handed is the Shell asked nothing, and the hand-over's own line is
    // `point_the_thumbnail`'s.
    if held.shown == Some(band) {
        note_pin_trace!(
            "page band: unchanged {}x{}+{}+{}",
            band.2,
            band.3,
            band.0,
            band.1
        );
        return;
    }

    point_the_thumbnail(&mut held, band);
}

/// Hand the Shell the band to draw into, and keep it. False where the Shell refused: nothing here asks
/// a second time, so the registration is let go (see `the_page_was_placed`).
fn point_the_thumbnail(held: &mut Held, band: Band) -> bool {
    // The whole client of the source, and never a rectangle of it: the Shell's source crop is ignored
    // and flattens the picture (see the module note).
    let properties = DWM_THUMBNAIL_PROPERTIES {
        dwFlags: DWM_TNP_RECTDESTINATION | DWM_TNP_SOURCECLIENTAREAONLY | DWM_TNP_VISIBLE,
        rcDestination: RECT {
            left: band.0,
            top: band.1,
            right: band.0 + band.2,
            bottom: band.1 + band.3,
        },
        opacity: 255,
        fVisible: true.into(),
        // The page is its client area and nothing around it: the window is a popup a browser draws in,
        // and a frame sampled with it would be a page drawn off its own box by the frame's width.
        fSourceClientAreaOnly: true.into(),
        ..Default::default()
    };

    if let Err(error) = unsafe { DwmUpdateThumbnailProperties(held.registration, &properties) } {
        let registration = held.registration;
        *held = Held::NOTHING;
        let _ = unsafe { DwmUnregisterThumbnail(registration) };
        note_pin_trace!("page band: update refused {error:?}");
        return false;
    }
    held.shown = Some(band);

    let pin =
        unsafe { pin_film_thumbnail::client_size(HWND(held.pin as *mut _)) }.unwrap_or((0, 0));
    let source =
        unsafe { pin_film_thumbnail::client_size(HWND(held.page as *mut _)) }.unwrap_or((0, 0));
    note_pin_trace!(
        "page band: shown pin {:#x} client {}x{} page {:#x} client {}x{} dest {}x{}+{}+{}",
        held.pin,
        pin.0,
        pin.1,
        held.page,
        source.0,
        source.1,
        band.2,
        band.3,
        band.0,
        band.1
    );
    true
}

/// Whether the pin frames an engine page the Shell is not drawing into it: a page in the band whose
/// registration is missing, made for another pair, or never pointed at a band.
///
/// **A refused registration is a page that never asks again unless this is asked**: the failure is
/// the Shell's answer to a source that has not presented, which the next placement would answer —
/// and a pin that is not moved, resized or swapped makes no next placement. The tick reads this and
/// asks (see `place_pinned_siblings`, and `pin_lifecycle::show_pinned_window` for the take-up).
pub(super) fn the_band_is_owed_a_page() -> bool {
    if pin_lifecycle::engine_page_in_the_band().is_none() {
        return false;
    }
    let page = webview_preview::showing_hwnd();
    if page == 0 {
        return false;
    }
    let Some(pin) = pin_real_window::real_pin_window() else {
        return false;
    };
    let Ok(held) = PAGE_THUMBNAIL.lock() else {
        return false;
    };
    held.registration == 0
        || held.shown.is_none()
        || held.pin != pin.0 as isize
        || held.page != page
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
fn the_pair(content: ScreenRegion) -> Option<(HWND, isize, Band)> {
    let page = webview_preview::showing_hwnd();
    if page == 0 {
        return None;
    }
    let pin = pin_real_window::real_pin_window()?;

    // The pin's client origin and area, which are the coordinates the band is placed in: a dragged pin
    // moves its client and the band together, so the rectangle handed over does not move at all.
    let origin = unsafe { pin_film_thumbnail::client_origin(pin) }?;
    let client = unsafe { pin_film_thumbnail::client_size(pin) }?;

    Some((pin, page, the_band_in_the_client(content, origin, client)))
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

/// Whether the Shell is owed a registration this time: a *new* one (`held_registration_is_zero`) waits
/// for the page it is to draw, and a standing one is kept whatever the answer is (see
/// `the_page_was_placed`).
fn registration_is_owed(held_registration_is_zero: bool, page_is_on_the_screen: bool) -> bool {
    !held_registration_is_zero || page_is_on_the_screen
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A fresh registration waits for the page it is to draw, and a standing one is kept: the Shell pictures
    /// a source that has presented nothing with its window's icon (see `registration_is_owed`).
    #[test]
    fn a_registration_waits_for_the_page_and_a_standing_one_is_kept() {
        assert!(
            !registration_is_owed(true, false),
            "a band with no picture behind it yet is no registration"
        );
        assert!(
            registration_is_owed(true, true),
            "and the page on the screen is the registration the band is owed"
        );
        assert!(
            registration_is_owed(false, false),
            "a standing registration is never dropped for a page that has not landed"
        );
        assert!(registration_is_owed(false, true), "nor for one that has");
    }

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
        // the rectangle the Shell is handed did not move at all.
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
            "a band that changed size is a band to hand over again"
        );
        assert!(
            held.pin != 4,
            "a pin that changed is a registration made again"
        );
        assert!(
            held.page != 5,
            "a page that changed is a registration made again"
        );
    }
}

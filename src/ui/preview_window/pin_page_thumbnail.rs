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
//! so the destination rectangle is the band's own rect inside the pin's client, less the rows the
//! caption stands in over it: a thumbnail composites over the destination window's own content, child
//! windows included, so a caption whose rows were handed over would be a caption under the picture —
//! invisible while its buttons still took the press (see `ARCHITECTURE.md`). The crop is a crop and not
//! a scale: the source loses the same rows, scaled by the factor it already had. Everything else the
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
    DWM_THUMBNAIL_PROPERTIES, DWM_TNP_RECTDESTINATION, DWM_TNP_RECTSOURCE,
    DWM_TNP_SOURCECLIENTAREAONLY, DWM_TNP_VISIBLE,
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
    /// The destination last handed, in the pin's client coordinates, and the caption rows cropped off
    /// its top, or nothing where none stands: a placement that asks for the same pair is the Shell
    /// asked nothing (see `the_page_was_placed`).
    shown: Option<(Band, i32)>,
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

    // The caption rows the registration already keeps out of the picture: a band that moved is the
    // same crop re-pointed, and the chrome road is what changes the crop itself
    // (see `the_caption_rows_are`).
    let caption = held.shown.map_or(0, |(_, caption)| caption);
    if held.shown == Some((destination, caption)) {
        note_pin_trace!(
            "page band: unchanged {}x{}+{}+{}",
            destination.2,
            destination.3,
            destination.0,
            destination.1
        );
        return;
    }

    if point_the_thumbnail(&mut held, destination, caption) {
        note_pin_trace!(
            "page band: shown {}x{}+{}+{}",
            destination.2,
            destination.3,
            destination.0,
            destination.1
        );
    }
}

/// Keep `caption` rows off the top of the page's picture, or draw all of it again, for the pin whose
/// chrome has just come or gone (see `pin_chrome_strips::sync_the_chrome_strips`).
///
/// A pin with no page registration — every pin that is not a browser's — is nothing at all here, and so
/// is a caption number the registration was handed already: this is asked after every paint.
pub(super) fn the_caption_rows_are(caption: i32) {
    let Ok(mut held) = PAGE_THUMBNAIL.lock() else {
        return;
    };
    // `shown` stands only with a registration: a page let go is nothing to crop.
    let Some((band, was)) = held.shown else {
        return;
    };
    if was == caption {
        return;
    }

    if !point_the_thumbnail(&mut held, band, caption) {
        note_pin_trace!("page band: caption {caption} refused, the registration is let go");
    }
}

/// Hand the Shell the destination and the source for a band and a caption, and keep the pair. False
/// where the Shell refused: nothing here asks a second time, so the registration is let go (see
/// `the_page_was_placed`).
fn point_the_thumbnail(held: &mut Held, band: Band, caption: i32) -> bool {
    // The source's client, which is a constant for as long as it is the pin's picture: a page that has
    // gone answers nothing, and the band's own size is then the factor it would have had anyway.
    let source = unsafe { pin_film_thumbnail::client_size(HWND(held.page as *mut _)) };
    let (destination, cropped) = the_rectangles(band, caption, source.unwrap_or((band.2, band.3)));

    let properties = DWM_THUMBNAIL_PROPERTIES {
        dwFlags: DWM_TNP_RECTDESTINATION
            | DWM_TNP_SOURCECLIENTAREAONLY
            | DWM_TNP_VISIBLE
            | cropped.map_or(0, |_| DWM_TNP_RECTSOURCE),
        rcDestination: destination,
        rcSource: cropped.unwrap_or_default(),
        opacity: 255,
        fVisible: true.into(),
        // The page is its client area and nothing around it: the window is a popup a browser draws in,
        // and a frame sampled with it would be a page drawn off its own box by the frame's width.
        fSourceClientAreaOnly: true.into(),
    };

    if let Some(source) = cropped {
        note_pin_trace!(
            "page band: caption {caption} rows cropped: dest {}x{}+{}+{} src {}x{}+{}+{}",
            destination.right - destination.left,
            destination.bottom - destination.top,
            destination.left,
            destination.top,
            source.right - source.left,
            source.bottom - source.top,
            source.left,
            source.top
        );
    }

    if let Err(error) = unsafe { DwmUpdateThumbnailProperties(held.registration, &properties) } {
        let registration = held.registration;
        *held = Held::NOTHING;
        let _ = unsafe { DwmUnregisterThumbnail(registration) };
        note_pin_trace!("page band: update refused {error:?}");
        return false;
    }
    held.shown = Some((band, caption));
    true
}

/// The two rectangles a registration is handed: the band, less the rows its caption stands in over
/// its top, and the same rows off the source's client — `None` where the source is the whole of it.
///
/// Cropped and never scaled: the source loses `rows * source height / band height`, which is the row
/// itself while the page is drawn one to one, and the same picture with its top cut off when it is not.
/// Nothing is cropped where the band already starts below the caption (a kind whose chrome stands in
/// bands around its media), nor where the crop would leave no rows at all.
fn the_rectangles(band: Band, caption: i32, source: (i32, i32)) -> (RECT, Option<RECT>) {
    let (x, y, width, height) = band;
    let whole = RECT {
        left: x,
        top: y,
        right: x + width,
        bottom: y + height,
    };

    let rows = (caption - y).clamp(0, height);
    if rows == 0 || rows >= height {
        return (whole, None);
    }

    let (source_width, source_height) = (source.0.max(1), source.1.max(1));
    let cropped =
        ((rows as i64 * source_height as i64 / height as i64) as i32).clamp(0, source_height - 1);

    (
        RECT {
            left: x,
            top: y + rows,
            right: x + width,
            bottom: y + height,
        },
        Some(RECT {
            left: 0,
            top: cropped,
            right: source_width,
            bottom: source_height,
        }),
    )
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
            shown: Some(((0, 28, 640, 272), 0)),
        };

        assert_eq!(held.shown, Some(((0, 28, 640, 272), 0)));
        assert_eq!((held.pin, held.page), (1, 2));
        assert!(
            held.shown != Some(((0, 28, 640, 271), 0)),
            "a band that changed is a pair to hand over again"
        );
        assert!(
            held.shown != Some(((0, 28, 640, 272), 28)),
            "a caption that changed is a pair to hand over again"
        );
    }

    /// **The caption's rows are given up, not scaled over**: a thumbnail composites over the pin's own
    /// content and its children with it, so a strip under it is a caption nobody can see whose buttons
    /// still take the press. The source loses the same rows at the factor it already had.
    #[test]
    fn the_caption_rows_are_cropped_off_both_rectangles() {
        let band = (0, 0, 640, 300);
        let source = (640, 300);

        // No caption: the whole band and the whole of the source, as the road has always handed it.
        assert_eq!(
            the_rectangles(band, 0, source),
            (
                RECT {
                    left: 0,
                    top: 0,
                    right: 640,
                    bottom: 300
                },
                None
            ),
            "no caption is no crop"
        );

        // A caption over the band: the same crop of the page's client, one to one.
        assert_eq!(
            the_rectangles(band, 28, source),
            (
                RECT {
                    left: 0,
                    top: 28,
                    right: 640,
                    bottom: 300
                },
                Some(RECT {
                    left: 0,
                    top: 28,
                    right: 640,
                    bottom: 300
                })
            ),
            "the caption's rows off the client as well"
        );

        // A source drawn at another size keeps the factor it had: the crop is the same fraction of it.
        assert_eq!(
            the_rectangles(band, 28, (1280, 600)),
            (
                RECT {
                    left: 0,
                    top: 28,
                    right: 640,
                    bottom: 300
                },
                Some(RECT {
                    left: 0,
                    top: 56,
                    right: 1280,
                    bottom: 600
                })
            ),
            "the crop is scaled by the factor the source already had"
        );

        // A band that already starts below the caption — the chrome in bands around its media — is
        // nothing to crop, the picture being under no strip of it.
        assert_eq!(
            the_rectangles((0, 28, 640, 272), 28, source),
            the_rectangles((0, 28, 640, 272), 0, source),
            "a caption the band starts under crops nothing"
        );

        // A caption taller than the band leaves no rows to draw in, and the band is handed over whole
        // rather than scaled or cropped into nothing.
        assert_eq!(
            the_rectangles(band, 300, source),
            the_rectangles(band, 0, source),
            "a crop that leaves no rows is no crop"
        );
    }
}

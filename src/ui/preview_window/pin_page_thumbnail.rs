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
//! **The road follows `pin_film_thumbnail`'s rules and has this one difference: the destination is a
//! child surface of the pin, not the pin.** The Shell composes a thumbnail **over its destination
//! window's own content, child windows included**, so a pin as the destination is a pin whose layered
//! surface — the caption, the transport, the alpha-1 claim — is painted *under* the page's picture,
//! invisible while its pixels still took the press. The destination is therefore a child window of the
//! pin, kept at the bottom of the pin's child order and under the chrome strips: the pin's own client,
//! then the picture, then the caption, the transport and the volume popup, which float over a page
//! because they are siblings above it. Nothing is cropped and nothing is scaled: the destination is the
//! picture child's whole client and the source is the page's whole client.
//!
//! **A drag re-points nothing at all.** The source's client is a constant for as long as it is the
//! pin's picture (the one placement, see `webview_preview`'s `put_window`), and the picture child is
//! placed in the pin's *client* coordinates, which move with the pin: a band that moves on the screen is
//! the same rectangle in its pin, and the Shell is handed the child's whole client either way.

use super::*;

use std::sync::atomic::{AtomicIsize, Ordering};
use std::sync::{Mutex, Once};

use crate::ui::preview_window::pin_trace::note_pin_trace;

use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Dwm::{
    DwmRegisterThumbnail, DwmUnregisterThumbnail, DwmUpdateThumbnailProperties,
    DWM_THUMBNAIL_PROPERTIES, DWM_TNP_RECTDESTINATION, DWM_TNP_SOURCECLIENTAREAONLY,
    DWM_TNP_VISIBLE,
};
use windows::Win32::Graphics::Gdi::{GetStockObject, HBRUSH, NULL_BRUSH};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, GetWindowLongPtrW, IsWindow, RegisterClassExW,
    SetWindowPos, ShowWindow, GWL_STYLE, HWND_BOTTOM, SWP_NOACTIVATE, SW_HIDE, SW_SHOWNOACTIVATE,
    WNDCLASSEXW, WS_CHILD, WS_DISABLED, WS_VISIBLE,
};

/// The band a picture child is placed at: `left, top, width, height` in the client coordinates of the
/// pin it belongs to.
type Band = (i32, i32, i32, i32);

/// The class the picture child is made from: a window that draws nothing of its own and takes nothing,
/// and whose whole surface is the Shell's thumbnail of the page (see the module note). Not layered — a
/// layered destination is refused by the Shell — so it needs a class of its own rather than the pin's.
const PICTURE_CLASS: PCWSTR = w!("RustHoverPreviewPinPagePicture");
static PICTURE_CLASS_REGISTERED: Once = Once::new();

/// The picture child standing, and the pin window it was made for: kept between placements like the
/// chrome strips, and made again where the pin it belongs to has changed (see `the_picture_at`).
static PICTURE: AtomicIsize = AtomicIsize::new(0);
static PICTURE_PARENT: AtomicIsize = AtomicIsize::new(0);

/// The registration standing, the two windows it was made for and the client size it was last handed:
/// `NOTHING` while no page is drawn.
static PAGE_THUMBNAIL: Mutex<Held> = Mutex::new(Held::NOTHING);

#[derive(Clone, Copy, PartialEq, Eq)]
struct Held {
    /// The picture child the Shell draws into, and the page window it draws.
    picture: isize,
    page: isize,
    registration: isize,
    /// The picture child's client size last handed — a placement that asks for the same size and page
    /// is the Shell asked nothing (see `the_page_was_placed`).
    shown: Option<(i32, i32)>,
}

impl Held {
    const NOTHING: Self = Self {
        picture: 0,
        page: 0,
        registration: 0,
        shown: None,
    };
}

/// The default answer is the whole of what a picture child does: it draws nothing and takes nothing, and
/// the pixels it shows are the Shell's thumbnail. The bindings expose `DefWindowProcW` as a
/// Rust-callable function, so the class still needs a system-ABI shim around it.
unsafe extern "system" fn the_default_answer(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
}

/// Draw the pin's page into it, for `content` — the media band the pin is showing, in screen
/// coordinates — or let the registration go where the pin has no page in its band.
///
/// It is asked wherever the band moves or changes (see `pin_lifecycle::place_pinned_siblings`) and it is
/// cheap to ask: a page that is not the pin's picture is a registration let go of, and a pair that did not
/// change hands the Shell nothing at all — **which is every step of a drag**, the picture child riding
/// the pin's own client coordinates rather than the screen.
pub(super) fn the_page_was_placed(content: Option<ScreenRegion>) {
    let Some((pin, page, band)) = content.and_then(the_pair) else {
        let_registration_go("the pin has no page in its band");
        return;
    };

    // The picture child is placed at the band before anything is handed over, so the destination it
    // offers is the band's own rectangle: under the strips, which are siblings kept above it.
    let Some(picture) = the_picture_at(pin, band) else {
        let_registration_go("the pin has no picture child for its page");
        return;
    };
    let Some(client) = (unsafe { pin_film_thumbnail::client_size(picture) }) else {
        let_registration_go("the picture child has left");
        return;
    };

    let Ok(mut held) = PAGE_THUMBNAIL.lock() else {
        return;
    };

    // **A registration is for one pair of windows**, so another picture child, or a page the engine has
    // replaced, is the picture given up and made again rather than a rectangle re-pointed at a window
    // the Shell was never asked about.
    if held.registration != 0 && (held.picture != picture.0 as isize || held.page != page) {
        let registration = held.registration;
        *held = Held::NOTHING;
        let _ = unsafe { DwmUnregisterThumbnail(registration) };
        hide_the_picture();
    }
    held.picture = picture.0 as isize;
    held.page = page;
    if held.registration == 0 {
        match unsafe { DwmRegisterThumbnail(picture, HWND(page as *mut _)) } {
            Ok(registration) => held.registration = registration,
            Err(error) => {
                // The Shell refuses a source that has made its window and not yet presented one, and nothing
                // here asks a second time: the next placement of the band does (see `the_page_was_placed`).
                *held = Held::NOTHING;
                hide_the_picture();
                note_pin_trace!("page band: register refused {error:?}");
                return;
            }
        }
    }

    // The destination is the picture child's whole client and its size is the band's: a band that moved
    // inside the pin hands the Shell the same pair, and only a band that changed size is a new one.
    if held.shown == Some(client) {
        note_pin_trace!(
            "page band: unchanged {}x{} picture {:#x} page={:#x}",
            client.0,
            client.1,
            picture.0 as isize,
            page
        );
        return;
    }

    point_the_thumbnail(&mut held, client, page);
}

/// Hand the Shell the destination and the source for a picture child, and keep the client size handed.
/// False where the Shell refused: nothing here asks a second time, so the registration is let go (see
/// `the_page_was_placed`).
fn point_the_thumbnail(held: &mut Held, client: (i32, i32), page: isize) -> bool {
    // The whole client of either window, and never a rectangle of it: the Shell's source crop is one the
    // picture child makes unnecessary (see the module note).
    let properties = DWM_THUMBNAIL_PROPERTIES {
        dwFlags: DWM_TNP_RECTDESTINATION | DWM_TNP_SOURCECLIENTAREAONLY | DWM_TNP_VISIBLE,
        rcDestination: RECT {
            left: 0,
            top: 0,
            right: client.0,
            bottom: client.1,
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
        hide_the_picture();
        note_pin_trace!("page band: update refused {error:?}");
        return false;
    }
    held.shown = Some(client);

    let source = unsafe { pin_film_thumbnail::client_size(HWND(page as *mut _)) }.unwrap_or((0, 0));
    note_pin_trace!(
        "page band: shown picture {:#x} client {}x{} source {:#x} client {}x{}",
        held.picture,
        client.0,
        client.1,
        page,
        source.0,
        source.1
    );
    true
}

/// The pin's picture child, placed at `band` in the pin's client coordinates and kept at the bottom of
/// its child order — under the chrome strips, which stay above it by being ordered over it.
///
/// Made once and kept, like a strip, and made again where the pin window it was made for has changed: a
/// pin window that went takes its children with it, and a child of another pin is a picture drawn into
/// somebody else's window.
fn the_picture_at(pin: HWND, band: Band) -> Option<HWND> {
    let mut picture = HWND(PICTURE.load(Ordering::SeqCst) as *mut _);
    let made_for = PICTURE_PARENT.load(Ordering::SeqCst);
    let standing = picture.0 as isize != 0
        && made_for == pin.0 as isize
        && unsafe { IsWindow(picture) }.as_bool();

    if !standing {
        if unsafe { IsWindow(picture) }.as_bool() {
            // A destination that goes takes its registration with it: let it go before the window does.
            let_registration_go("the picture child's pin changed");
            let _ = unsafe { DestroyWindow(picture) };
            note_pin_trace!("page picture: destroyed {:#x}", picture.0 as isize);
        }
        PICTURE.store(0, Ordering::SeqCst);
        PICTURE_PARENT.store(0, Ordering::SeqCst);
        picture = make_the_picture(pin)?;
        PICTURE.store(picture.0 as isize, Ordering::SeqCst);
        PICTURE_PARENT.store(pin.0 as isize, Ordering::SeqCst);
        note_pin_trace!(
            "page picture: made {:#x} for pin {:#x}",
            picture.0 as isize,
            pin.0 as isize
        );
    }

    // Bottom of the child order and at the band in one call: the strips are made on top and placed with
    // `SWP_NOZORDER`, so this is the order that sticks (see `pin_chrome_strips`).
    let _ = unsafe {
        SetWindowPos(
            picture,
            HWND_BOTTOM,
            band.0,
            band.1,
            band.2,
            band.3,
            SWP_NOACTIVATE,
        )
    };

    let up = unsafe { GetWindowLongPtrW(picture, GWL_STYLE) } & WS_VISIBLE.0 as isize != 0;
    if !up {
        let _ = unsafe { ShowWindow(picture, SW_SHOWNOACTIVATE) };
    }
    Some(picture)
}

/// Make the picture child: a child of the pin, disabled so that a press on the band keeps landing on the
/// pin, and with no layered style — the Shell refuses a layered destination.
fn make_the_picture(pin: HWND) -> Option<HWND> {
    PICTURE_CLASS_REGISTERED.call_once(|| unsafe {
        let hinstance = GetModuleHandleW(None).expect("the module's own handle");
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(the_default_answer),
            hInstance: hinstance.into(),
            hbrBackground: HBRUSH(GetStockObject(NULL_BRUSH).0),
            lpszClassName: PICTURE_CLASS,
            ..Default::default()
        };
        RegisterClassExW(&class);
    });

    let hinstance = unsafe { GetModuleHandleW(None) }.ok()?;
    unsafe {
        CreateWindowExW(
            Default::default(),
            PICTURE_CLASS,
            w!("page picture"),
            WS_CHILD | WS_DISABLED,
            0,
            0,
            1,
            1,
            pin,
            None,
            hinstance,
            None,
        )
    }
    .ok()
}

/// Take the picture child off the screen, keeping it: nothing is drawn into a child with no registration,
/// and a hidden one carries the last page into the next pin that finds it standing.
fn hide_the_picture() {
    let picture = HWND(PICTURE.load(Ordering::SeqCst) as *mut _);
    if picture.0 as isize == 0 || !unsafe { IsWindow(picture) }.as_bool() {
        return;
    }
    let up = unsafe { GetWindowLongPtrW(picture, GWL_STYLE) } & WS_VISIBLE.0 as isize != 0;
    if up {
        let _ = unsafe { ShowWindow(picture, SW_HIDE) };
        note_pin_trace!("page picture: hidden {:#x}", picture.0 as isize);
    }
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
        hide_the_picture();
    }
}

/// The pin window whose child draws the page, the page window it draws, and the band's rectangle in the
/// pin's own client coordinates — or nothing where either window or the band is not there.
fn the_pair(content: ScreenRegion) -> Option<(HWND, isize, Band)> {
    let page = webview_preview::showing_hwnd();
    if page == 0 {
        return None;
    }
    let pin = pin_real_window::real_pin_window()?;

    // The pin's client origin and area, which are the coordinates the picture child is placed in: a
    // layered pin is drawn where its client is, so the band is the same rectangle in either coordinates —
    // and the *difference* is what makes a drag re-point nothing (see the module note).
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
        // the rectangle the picture child is placed at did not move at all.
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
            picture: 1,
            page: 2,
            registration: 3,
            shown: Some((640, 272)),
        };

        assert_eq!(held.shown, Some((640, 272)));
        assert_eq!((held.picture, held.page), (1, 2));
        assert!(
            held.shown != Some((640, 271)),
            "a band that changed size is a pair to hand over again"
        );
        assert!(
            held.picture != 4,
            "a picture child that changed is a registration made again"
        );
        assert!(
            held.page != 5,
            "a page that changed is a registration made again"
        );
    }
}

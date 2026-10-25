//! The windows the pin's chrome and overlays are carried in.
//!
//! **Why they exist.** The film's host is a child of the pin's window standing in the whole of the
//! client rectangle (see `band_in_the_window`), and a child composites over the layered surface its
//! parent paints. So the moment mpv began to draw, the film stood over the caption and the
//! transport. The pin's bitmap therefore paints no chrome at all — only its alpha-1 claim on the
//! chrome's rows, which keeps every press on them the pin's (see `render_pinned_preview_at`) — and
//! the chrome has windows of its own: layered children of the pin's window, placed at the strips
//! the chrome is painted in, which composite over whatever stands in the band — a hosted film, the
//! cover, FFmpeg's player, or this app's own picture.
//!
//! **The volume popup needs one for the same reason.** It is painted into the pin's own surface too,
//! but over a hosted film that surface is under the child, so the panel was invisible while its
//! pixels still took the press. Its strip is a third one, covering only the popup's footprint, and
//! it stands over the film while the popup is open.
//!
//! **What they carry, and when.** Their surfaces are what the pin paints — the caption's strip, the
//! transport's, and the volume popup's footprint, taken as they were painted (see `CHROME_SURFACE`,
//! `TRANSPORT_SURFACE` and `VOLUME_SURFACE`) — and each strip is on the screen while the chrome
//! state asks for it, on every road. They belong to the pin's window rather than to any host: made
//! with it, they stand wherever the pin's band is, whatever is drawing in it.
//!
//! **Why they take no input.** They are made disabled, and a disabled child is skipped by
//! `WindowFromPoint`, so a press on the chrome keeps landing on the pin and being answered by its
//! coordinates — the same answer it had when the chrome was only paint (see the note over the
//! host, which is disabled for the same reason).
//!
//! **Why `UpdateLayeredWindow` and never `SetLayeredWindowAttributes`.** The strips are carried by
//! per-pixel alpha, because the pixels they carry are the pin's own — alpha and all — and the two
//! calls are exclusive on a layered window: once SLWA has been asked of one, ULW fails on it until
//! the layered bit is cleared and set again (measured, as the child-and-layered pair itself was).
//! Nothing here ever asks for SLWA.

use std::sync::atomic::{AtomicIsize, Ordering};
use std::sync::Once;

use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{GetStockObject, HBRUSH, NULL_BRUSH};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, GetWindowLongPtrW, RegisterClassExW, SetWindowPos, ShowWindow,
    GWL_STYLE, HWND_TOP, SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER, SW_HIDE,
    SW_SHOWNOACTIVATE, WNDCLASSEXW, WS_CHILD, WS_DISABLED, WS_EX_LAYERED, WS_VISIBLE,
};

use super::*;

/// The class the strips are made from: two windows that draw nothing of their own and take
/// nothing, painted by the pin's paint the way the pin's own surface is. A class of their own
/// rather than the pin window's, for the reason the host has one (see `pin_film_host`).
const STRIP_CLASS: PCWSTR = w!("RustHoverPreviewPinChromeStrip");

/// The default answer is the whole of what a strip does: it draws nothing and takes nothing, and
/// the pixels it shows are this app's own `UpdateLayeredWindow` (see the module doc). The
/// bindings expose `DefWindowProcW` as a Rust-callable function, so the class still needs a
/// system-ABI shim around it.
unsafe extern "system" fn the_default_answer(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
}

/// The caption's strip, the transport's, and the volume popup's, where the pin's chrome has
/// windows of its own, and nothing while the pin's window has not taken them up.
static CAPTION_STRIP: AtomicIsize = AtomicIsize::new(0);
static TRANSPORT_STRIP: AtomicIsize = AtomicIsize::new(0);
static VOLUME_STRIP: AtomicIsize = AtomicIsize::new(0);
static STRIP_CLASS_REGISTERED: Once = Once::new();

pub(super) fn caption_strip() -> Option<HWND> {
    let hwnd = CAPTION_STRIP.load(Ordering::SeqCst);
    (hwnd != 0).then_some(HWND(hwnd as *mut _))
}

pub(super) fn transport_strip() -> Option<HWND> {
    let hwnd = TRANSPORT_STRIP.load(Ordering::SeqCst);
    (hwnd != 0).then_some(HWND(hwnd as *mut _))
}

pub(super) fn volume_strip() -> Option<HWND> {
    let hwnd = VOLUME_STRIP.load(Ordering::SeqCst);
    (hwnd != 0).then_some(HWND(hwnd as *mut _))
}

/// What a strip is carried by, read out of the pin in one look: the box the strips are placed at,
/// and whether each strip is the chrome the pin is painting now. It is the same read the paint
/// makes (`pinned_paint`), so what a strip is shown at is what the chrome is drawn at, and what a
/// strip shows is what the chrome is drawn as.
struct TheStripsState {
    width: i32,
    height: i32,
    caption_height: i32,
    transport_height: i32,
    caption: bool,
    bar: bool,
    /// The footprint the volume popup is drawn in, or nothing while it is closed.
    volume: Option<RECT>,
}

fn the_strips_state() -> Option<TheStripsState> {
    let pinned = pin_state()?;
    let pin = pinned.pin()?;
    if pin.collapsed {
        return None;
    }

    let (width, height) = pin.window_size();
    let caption_height = pin.caption;
    let transport_height = pinned_transport_height(pin.dpi, pin.transport_bar);
    let caption = caption_height > 0
        && pin.chrome.caption
        && pin_chrome::caption_fits(width, pin.dpi, pin.frame != PinFrame::None);
    let bar = transport_height > 0 && pin.chrome.bar;
    let volume = pin.volume.open.then(|| {
        pinned_volume_panel(pin).map(|popup| pin_chrome::volume_popup_bounds(&popup, width, height))
    });
    let volume = volume.flatten();

    Some(TheStripsState {
        width,
        height,
        caption_height,
        transport_height,
        caption,
        bar,
        volume,
    })
}

/// Make the strips with the pin's own window: children of it, layered from the moment they
/// are made (a layered child is the one thing a film's host is allowed to have over it, and it is
/// measured to work — see the module doc over `pin_film_host`), disabled, and off the screen until
/// something is carried in them.
///
/// The class is registered here the way the host's is, and its style is nothing at all: no
/// `CS_CLASSDC`, which breaks the layered calls, and no `CS_OWNDC` either, which nothing here
/// needs — a strip's surface is this app's own DIB, not a GDI device (see `ensure_layered_surface`).
///
/// Asked wherever the pin's window is taken up, and a take-up that finds them standing finds them
/// kept: they belong to the window rather than to any road, and a strip left standing carries the
/// chrome into the next take-up's band rather than leaving it without any.
///
/// Above the host, and left there: the host is placed with `SWP_NOZORDER` wherever the band
/// goes, so this is the order that sticks — until the host is made again, which puts it on top
/// and is answered by `raise_the_chrome_strips_over_the_host` (see `pin_film_host`).
pub(super) unsafe fn take_the_strips_up(parent: HWND) {
    STRIP_CLASS_REGISTERED.call_once(|| {
        let hinstance = GetModuleHandleW(None).expect("the module's own handle");
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(the_default_answer),
            hInstance: hinstance.into(),
            hbrBackground: HBRUSH(GetStockObject(NULL_BRUSH).0),
            lpszClassName: STRIP_CLASS,
            ..Default::default()
        };
        RegisterClassExW(&class);
    });

    let Ok(hinstance) = GetModuleHandleW(None) else {
        return;
    };
    let make = |name: PCWSTR| -> Option<HWND> {
        let strip = CreateWindowExW(
            WS_EX_LAYERED,
            STRIP_CLASS,
            name,
            WS_CHILD | WS_DISABLED,
            0,
            0,
            1,
            1,
            parent,
            None,
            hinstance,
            None,
        )
        .ok()?;
        // Above the host, and left there: the host is placed with `SWP_NOZORDER` wherever the band
        // goes, so this is the order that sticks — until the host is made again (see
        // `raise_the_chrome_strips_over_the_host`).
        let _ = SetWindowPos(
            strip,
            HWND_TOP,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
        Some(strip)
    };

    // One at a time: a take-up that finds one strip standing and the other gone keeps the one,
    // and a parent a test stood in answers the make with nothing (see
    // `stand_real_pin_window_for_a_test`).
    if !caption_strip().is_some_and(|strip| unsafe { IsWindow(strip) }.as_bool()) {
        if let Some(strip) = make(w!("caption strip")) {
            CAPTION_STRIP.store(strip.0 as isize, Ordering::SeqCst);
        }
    }
    if !transport_strip().is_some_and(|strip| unsafe { IsWindow(strip) }.as_bool()) {
        if let Some(strip) = make(w!("transport strip")) {
            TRANSPORT_STRIP.store(strip.0 as isize, Ordering::SeqCst);
        }
    }
    if !volume_strip().is_some_and(|strip| unsafe { IsWindow(strip) }.as_bool()) {
        if let Some(strip) = make(w!("volume strip")) {
            VOLUME_STRIP.store(strip.0 as isize, Ordering::SeqCst);
        }
    }
}

/// Put the strips back over the host a take-up has just made: making it puts it on top of them,
/// and the chrome stands over the film rather than under it (see `pin_film_host`).
///
/// The caption first, the transport after it, and the volume popup last, which is the order the
/// take-up made them in: the popup stands nearest the top, the transport under it, the caption
/// under that, and the host under them all.
pub(super) unsafe fn raise_the_chrome_strips_over_the_host() {
    for strip in [caption_strip(), transport_strip(), volume_strip()]
        .into_iter()
        .flatten()
    {
        if unsafe { IsWindow(strip) }.as_bool() {
            let _ = SetWindowPos(
                strip,
                HWND_TOP,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
            );
        }
    }
}

/// Take the strips down: the pin's window they are children of is going, and a strip left
/// standing would carry the last pin's chrome into the next pin's band. Windows itself destroys
/// them with their parent; what is forgotten here are the handles.
///
/// Only tests take the pin's window down — the run keeps it between pins — so only tests ask
/// for this (see `StandingFilm`).
#[cfg(test)]
pub(super) unsafe fn take_the_strips_down() {
    use windows::Win32::UI::WindowsAndMessaging::DestroyWindow;

    if let Some(strip) = caption_strip() {
        let _ = DestroyWindow(strip);
    }
    if let Some(strip) = transport_strip() {
        let _ = DestroyWindow(strip);
    }
    if let Some(strip) = volume_strip() {
        let _ = DestroyWindow(strip);
    }
    CAPTION_STRIP.store(0, Ordering::SeqCst);
    TRANSPORT_STRIP.store(0, Ordering::SeqCst);
    VOLUME_STRIP.store(0, Ordering::SeqCst);
}

/// Put the strips where the chrome is painted: the caption across the window's first rows and
/// the transport across its last (see `pinned_band_rows`, which is where those rows are read
/// from by the paint).
///
/// It is asked wherever the pin's box is placed, so the strips never stand anywhere the band is
/// not — and a strip whose rows do not exist (a pin with no caption, or none playing) is left
/// where it is, because the sync never shows it.
pub(super) unsafe fn place_the_chrome_strips() {
    let Some(state) = the_strips_state() else {
        return;
    };

    // The row a resize pushed the pin's picture down by, which is the row the window was moved up by:
    // a strip is placed in the client's own coordinates and so rides that move, and adding the row
    // here is what keeps the chrome on the box the pin's picture landed on (see
    // `the_slack_a_resize_paints_with`). The sizes are the pin's own and not the surface's, so only
    // the place moves.
    let row = paint::the_row_a_resize_paints_the_picture_down();

    if state.caption_height > 0 {
        if let Some(strip) = caption_strip() {
            let _ = SetWindowPos(
                strip,
                HWND::default(),
                0,
                row,
                state.width,
                state.caption_height,
                SWP_NOZORDER | SWP_NOACTIVATE,
            );
        }
    }

    if state.transport_height > 0 {
        if let Some(strip) = transport_strip() {
            let _ = SetWindowPos(
                strip,
                HWND::default(),
                0,
                state.height - state.transport_height + row,
                state.width,
                state.transport_height,
                SWP_NOZORDER | SWP_NOACTIVATE,
            );
        }
    }

    if let Some(volume) = state.volume {
        if let Some(strip) = volume_strip() {
            let _ = SetWindowPos(
                strip,
                HWND::default(),
                volume.left,
                volume.top + row,
                (volume.right - volume.left).max(1),
                (volume.bottom - volume.top).max(1),
                SWP_NOZORDER | SWP_NOACTIVATE,
            );
        }
    }
}

/// Carry the strip's pixels into its own layered surface, and answer whether they were carried.
///
/// The surface asked for is the strip's own, kept between paints like every other layered surface
/// (see `ensure_layered_surface`). The pixels are copied verbatim from the surface the chrome was
/// painted on — they are already the form a layered window is composited from (see `pin_chrome`).
unsafe fn paint_a_strip(strip: HWND, surface: &DibSurface) -> bool {
    let (width, height) = (surface.width, surface.height);
    if width == 0 || height == 0 {
        return false;
    }

    let Some(bits) = ensure_layered_surface(strip.0 as isize, width, height) else {
        return false;
    };
    unsafe {
        std::ptr::copy_nonoverlapping(surface.bits(), bits, width as usize * height as usize * 4);
    }

    let Some(mem_dc) = layered_surface_dc(strip.0 as isize) else {
        return false;
    };
    let src_point = POINT { x: 0, y: 0 };
    let size = SIZE {
        cx: width as i32,
        cy: height as i32,
    };
    let blend = BLENDFUNCTION {
        BlendOp: AC_SRC_OVER as u8,
        BlendFlags: 0,
        SourceConstantAlpha: 255,
        AlphaFormat: AC_SRC_ALPHA as u8,
    };
    // The place is not given: the strip stands where `place_the_chrome_strips` put it, and
    // `UpdateLayeredWindow` with no place changes the pixels and the size and nothing else.
    unsafe {
        UpdateLayeredWindow(
            strip,
            None,
            None,
            Some(&size),
            mem_dc,
            Some(&src_point),
            COLORREF(0),
            Some(&blend),
            ULW_ALPHA,
        )
        .is_ok()
    }
}

/// Show or hide one strip as the chrome and the host ask, carrying its pixels up first so that a
/// strip is never on the screen without what belongs in it.
///
/// A strip whose pixels cannot be carried — the surface the chrome was painted on is not the size
/// of the strip's rows now, which is a box the paint has not caught up with yet — is kept off the
/// screen rather than shown carrying what it last had: the paint that follows the box change
/// repaints it and the sync after that paint shows it (see `sync_the_chrome_strips`).
unsafe fn set_the_strip_at(
    strip: HWND,
    surface: Option<&DibSurface>,
    wanted: (u32, u32),
    shown: bool,
) {
    // The strip's own bit, not `IsWindowVisible`: that answer is false under a hidden pin, and a
    // strip left up by one pin would otherwise be revealed with the next (see `show_pinned_window`).
    let up = unsafe { GetWindowLongPtrW(strip, GWL_STYLE) } & WS_VISIBLE.0 as isize != 0;
    if !shown {
        if up {
            let _ = unsafe { ShowWindow(strip, SW_HIDE) };
        }
        return;
    }

    let carries = surface
        .map(|surface| (surface.width, surface.height) == wanted)
        .unwrap_or(false);
    if carries && unsafe { paint_a_strip(strip, surface.expect("the surface the strip carries")) } {
        if !up {
            let _ = unsafe { ShowWindow(strip, SW_SHOWNOACTIVATE) };
        }
    } else if up {
        let _ = unsafe { ShowWindow(strip, SW_HIDE) };
    }
}

/// Bring the strips in line with the pin and its chrome, which is the one applier every road
/// answers: the tick's chrome-change branch, the paint the chrome's own show and hide are answered
/// on, and wherever the host's own visibility changes (see `set_the_pin_film_host_visible`, whose
/// show and hide are every road a park, a swap and the film's own arrival take). A strip is on the
/// screen while the chrome asks for it, and what it carries is the pixels the pin painted into
/// those strips the last time it painted — the pin's bitmap carries none of them (see
/// `render_pinned_preview_at`).
///
/// The place is put first: a box the paint has not caught up with yet is a size the strips'
/// surfaces do not have, and a strip shown carrying what it last had is kept off the screen until
/// the paint after the box change repaints it (see `set_the_strip_at`).
pub(super) unsafe fn sync_the_chrome_strips() {
    // The strips stand where the chrome is painted before anything is carried into them, so a
    // resize answers them on the paint it is painted on rather than on the place after it.
    place_the_chrome_strips();

    let (Some(caption), Some(transport), Some(volume)) =
        (caption_strip(), transport_strip(), volume_strip())
    else {
        return;
    };

    let Some(state) = the_strips_state() else {
        // No pin to read, or a pin collapsed into its bubble: whatever the strips are carrying,
        // nothing on screen is asking for any of it.
        let _ = unsafe { ShowWindow(caption, SW_HIDE) };
        let _ = unsafe { ShowWindow(transport, SW_HIDE) };
        let _ = unsafe { ShowWindow(volume, SW_HIDE) };
        return;
    };

    let caption_wanted = (
        state.width.max(1) as u32,
        state.caption_height.max(1) as u32,
    );
    CHROME_SURFACE.with(|cell| {
        let surface = cell.borrow();
        unsafe {
            set_the_strip_at(caption, surface.as_ref(), caption_wanted, state.caption);
        }
    });

    let transport_wanted = (
        state.width.max(1) as u32,
        state.transport_height.max(1) as u32,
    );
    TRANSPORT_SURFACE.with(|cell| {
        let surface = cell.borrow();
        unsafe {
            set_the_strip_at(transport, surface.as_ref(), transport_wanted, state.bar);
        }
    });

    let volume_wanted = state
        .volume
        .map(|rect| {
            (
                (rect.right - rect.left).max(1) as u32,
                (rect.bottom - rect.top).max(1) as u32,
            )
        })
        .unwrap_or((1, 1));
    VOLUME_SURFACE.with(|cell| {
        let surface = cell.borrow();
        unsafe {
            set_the_strip_at(
                volume,
                surface.as_ref(),
                volume_wanted,
                state.volume.is_some(),
            );
        }
    });
}

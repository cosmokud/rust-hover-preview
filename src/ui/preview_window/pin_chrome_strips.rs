//! The two windows a hosted film's chrome is carried over it in.
//!
//! **Why they exist.** The film's host is a child of the pin's window standing in the whole of the
//! client rectangle (see `band_in_the_window`), and a child composites over the layered surface its
//! parent paints. So the moment mpv began to draw, the film stood over the caption and the
//! transport: the pin went on painting both into its own rows and answering presses on them by
//! their coordinates, and neither was ever seen again. The chrome therefore has windows of its own
//! — layered children of the pin's window, placed at the strips the chrome is painted in — which
//! composite over the host and so over the film, carrying the very pixels the pin paints.
//!
//! **What they carry, and when.** Their surfaces are the strips the pin paints — the caption's
//! surface and the transport's, taken as they were painted (see `CHROME_SURFACE` and
//! `TRANSPORT_SURFACE`) — and a strip is on the screen only while the host is *and* the chrome
//! state asks for that strip. A park hides the host, so the strips go with it; the chrome's own
//! show and hide is answered on the repaint it is answered by everywhere else (see `PinChrome`).
//! The pin goes on painting both strips into its own rows exactly as it always has: that surface
//! is what stands over the film on every road that has no host in it, and these windows only ever
//! stand where the pin's own pixels already are.
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
    CreateWindowExW, DefWindowProcW, DestroyWindow, RegisterClassExW, SetWindowPos, ShowWindow,
    HWND_TOP, SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER, SW_HIDE, SW_SHOWNOACTIVATE,
    WNDCLASSEXW, WS_CHILD, WS_DISABLED, WS_EX_LAYERED,
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

/// The caption's strip and the transport's, where a hosted film's chrome has windows of its own,
/// and nothing while there is no host for them to stand over.
static CAPTION_STRIP: AtomicIsize = AtomicIsize::new(0);
static TRANSPORT_STRIP: AtomicIsize = AtomicIsize::new(0);
static STRIP_CLASS_REGISTERED: Once = Once::new();

pub(super) fn caption_strip() -> Option<HWND> {
    let hwnd = CAPTION_STRIP.load(Ordering::SeqCst);
    (hwnd != 0).then_some(HWND(hwnd as *mut _))
}

pub(super) fn transport_strip() -> Option<HWND> {
    let hwnd = TRANSPORT_STRIP.load(Ordering::SeqCst);
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

    Some(TheStripsState {
        width,
        height,
        caption_height,
        transport_height,
        caption,
        bar,
    })
}

/// Make the two strips with the host they stand over: children of the pin's own window, layered
/// from the moment they are made (a layered child is the one thing a film's host is allowed to
/// have over it, and it is measured to work — see the module doc over `pin_film_host`), disabled,
/// and off the screen until something is carried in them.
///
/// The class is registered here the way the host's is, and its style is nothing at all: no
/// `CS_CLASSDC`, which breaks the layered calls, and no `CS_OWNDC` either, which nothing here
/// needs — a strip's surface is this app's own DIB, not a GDI device (see `ensure_layered_surface`).
pub(super) unsafe fn take_the_strips_up_with_the_host(parent: HWND) {
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
        // goes, so this is the order that sticks (see `pin_film_host`).
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

    if let Some(strip) = make(w!("caption strip")) {
        CAPTION_STRIP.store(strip.0 as isize, Ordering::SeqCst);
    }
    if let Some(strip) = make(w!("transport strip")) {
        TRANSPORT_STRIP.store(strip.0 as isize, Ordering::SeqCst);
    }
}

/// Take the two strips down with the host: the pin they are children of is ending or has stopped
/// showing a film, and a strip left standing would carry the last film's chrome into the next
/// pin's band (see `take_down_the_pin_film_host`).
pub(super) unsafe fn take_the_strips_down() {
    if let Some(strip) = caption_strip() {
        let _ = DestroyWindow(strip);
    }
    if let Some(strip) = transport_strip() {
        let _ = DestroyWindow(strip);
    }
    CAPTION_STRIP.store(0, Ordering::SeqCst);
    TRANSPORT_STRIP.store(0, Ordering::SeqCst);
}

/// Put the strips where the chrome is painted, at the box the host has just been placed at: the
/// caption across the window's first rows and the transport across its last (see
/// `pinned_band_rows`, which is where those rows are read from by the paint).
///
/// It is asked wherever the host is placed, so the strips never stand anywhere the film is not —
/// and a strip whose rows do not exist (a pin with no caption, or none playing) is left where it
/// is, because the sync never shows it.
pub(super) unsafe fn place_the_chrome_strips() {
    let Some(state) = the_strips_state() else {
        return;
    };

    if state.caption_height > 0 {
        if let Some(strip) = caption_strip() {
            let _ = SetWindowPos(
                strip,
                HWND::default(),
                0,
                0,
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
                state.height - state.transport_height,
                state.width,
                state.transport_height,
                SWP_NOZORDER | SWP_NOACTIVATE,
            );
        }
    }
}

/// Carry the strip's pixels into its own layered surface, and answer whether they were carried.
///
/// The surface asked for is the strip's own, kept between paints like every other layered surface
/// (see `ensure_layered_surface`). The pixels are copied verbatim from the surface the chrome was
/// painted on — they are already the form a layered window is composited from, which is the form
/// `copy_surface_rows_into` carries them to the pin's surface in (see `pin_chrome`).
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
    let up = unsafe { IsWindowVisible(strip) }.as_bool();
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

/// Bring the strips in line with the host and the chrome: a strip is on the screen while the host
/// is *and* the chrome asks for that strip, and what it carries is the pixels the pin painted into
/// those rows the last time it painted.
///
/// It is asked wherever the host's own visibility changes (see `set_the_pin_film_host_visible`,
/// whose show and hide are every road a park, a swap and the film's own arrival take) and after
/// every paint of the pin, which is where the chrome's own show and hide are answered (see
/// `PinChrome`). It is asked for nothing where there is no host: the strips are the host's, made
/// and taken down with it (see `take_the_strips_up_with_the_host`).
pub(super) unsafe fn sync_the_chrome_strips() {
    let Some(host) = pin_film_host::film_host() else {
        return;
    };
    let (Some(caption), Some(transport)) = (caption_strip(), transport_strip()) else {
        return;
    };

    let Some(state) = the_strips_state() else {
        // No pin to read, or a pin collapsed into its bubble: whatever the strips are carrying,
        // the pin's own surface is not being asked for any of it.
        let _ = unsafe { ShowWindow(caption, SW_HIDE) };
        let _ = unsafe { ShowWindow(transport, SW_HIDE) };
        return;
    };

    let host_is_up = unsafe { IsWindowVisible(host) }.as_bool();
    let caption_wanted = (
        state.width.max(1) as u32,
        state.caption_height.max(1) as u32,
    );
    CHROME_SURFACE.with(|cell| {
        let surface = cell.borrow();
        unsafe {
            set_the_strip_at(
                caption,
                surface.as_ref(),
                caption_wanted,
                host_is_up && state.caption,
            );
        }
    });

    let transport_wanted = (
        state.width.max(1) as u32,
        state.transport_height.max(1) as u32,
    );
    TRANSPORT_SURFACE.with(|cell| {
        let surface = cell.borrow();
        unsafe {
            set_the_strip_at(
                transport,
                surface.as_ref(),
                transport_wanted,
                host_is_up && state.bar,
            );
        }
    });
}

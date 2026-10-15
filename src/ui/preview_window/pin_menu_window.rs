//! The window a card's menu floats in: a panel of this app's own over the screen,
//! rather than rows drawn inside the card the menu came out of.
//!
//! The menu is the card's own and is drawn the way the card's own panels are drawn
//! (see `pin_chrome::menu`), but the window it is drawn in is not the card's: a panel
//! held inside a card's own box is a panel the card's own edges cut rows off — the
//! card is as tall as its media and no taller, and the menu is taller than that. So the
//! menu has a window of its own, put where the panel is and sized to the panel, and it
//! is the display's own work area rather than the card's box that says where a panel
//! may go (see `pinned_menu_paint`).
//!
//! It is a layered popup that never takes focus and is never in the taskbar, exactly
//! as the card's own window and the bubble are, and it takes no input of its own: the
//! card's window holds the pointer while its menu is up, so a press anywhere goes
//! there and is answered against the panel's own rows in the screen's own coordinates
//! (see `pin_menu::pinned_menu_press`).
//!
//! There is no such window where the card's own is a stand-in, which is what a test
//! has, so every road here answers nothing and builds nothing without one
//! (`pin_menu_window`).

use super::*;

/// The class the menu window is created from.
pub(super) const PIN_MENU_CLASS: PCWSTR = w!("RustHoverPreviewPinMenu");

/// The menu window that is up, or nothing while no menu is.
static PIN_MENU_HWND: AtomicIsize = AtomicIsize::new(0);

/// Put the card's menu where it stands, in a window of its own over the screen, and
/// take that window away where no menu is up.
///
/// It is asked of the state rather than told about a change, so that every road a menu
/// has out of — a press that answered a row, a press that left it, the ask the Explorer
/// hook left behind, a pin taken down — leaves nothing on screen: the window is the
/// paint, and the paint is read back from where the menu's own state is held (see
/// `pinned_menu_geometry`).
///
/// The window is put above the card's own while it is up, and asked for again on the
/// loop's clock for that reason among others: the card is raised on its own roads (see
/// `pin_chrome_up`), and a menu left under the card it came out of is rows a hand
/// cannot read.
pub(super) unsafe fn render_pin_menu() {
    let Some(paint) = pinned_menu_geometry() else {
        hide_pin_menu_window();
        return;
    };
    let Some(hwnd) = pin_menu_window() else {
        return;
    };

    // The scale the panels were laid out at, which is the scale they are painted at: it
    // is the one the card itself was asked at, so that a panel's own pads, rows and marks are
    // the size its placement measured them (see `pin_chrome::menu_panel_sizes`).
    let scale = paint.dpi as f32 / 96.0;

    // The window is the two panels and the shadow under them, and no more: a window
    // bigger than its panel is a window whose transparent room is drawn over whatever
    // is under it for nothing.
    let shadow = (pin_chrome::MENU_SHADOW_PIXELS * scale).round().max(1.0) as i32;
    let panels = [
        Some(&paint.popup),
        paint.flyout.as_ref().map(|flyout| &flyout.popup),
    ];
    let left = panels
        .iter()
        .flatten()
        .map(|panel| panel.panel.left)
        .min()
        .unwrap_or(0);
    let right = panels
        .iter()
        .flatten()
        .map(|panel| panel.panel.right)
        .max()
        .unwrap_or(0);
    let top = panels
        .iter()
        .flatten()
        .map(|panel| panel.panel.top)
        .min()
        .unwrap_or(0);
    let bottom = panels
        .iter()
        .flatten()
        .map(|panel| panel.panel.bottom)
        .max()
        .unwrap_or(0)
        + shadow;
    let width = (right - left).max(1);
    let height = (bottom - top).max(1);

    let _ = SetWindowPos(
        hwnd,
        HWND_TOPMOST,
        left,
        top,
        width,
        height,
        SWP_NOACTIVATE | SWP_SHOWWINDOW,
    );

    let Some(bits) = ensure_layered_surface(hwnd.0 as isize, width as u32, height as u32) else {
        return;
    };
    let out = unsafe { std::slice::from_raw_parts_mut(bits, width as usize * height as usize * 4) };
    // The window is nothing but the panels in it: the rest of it is the desktop the
    // menu floats over, which is what a blank buffer is.
    out.fill(0);

    let Some(palette) = pin_chrome::ChromePalette::current() else {
        return;
    };

    MENU_WINDOW_SURFACE.with(|cell| {
        let mut surface = cell.borrow_mut();
        if surface
            .as_ref()
            .map(|surface| (surface.width, surface.height))
            != Some((width as u32, height as u32))
        {
            *surface = DibSurface::create(width as u32, height as u32);
        }
        let Some(surface) = surface.as_ref() else {
            return;
        };

        let blanked = unsafe {
            std::slice::from_raw_parts_mut(
                surface.bits(),
                surface.width as usize * surface.height as usize * 4,
            )
        };
        blanked.fill(0);

        // The panel is held on the screen and painted into this window's own rows, so
        // it is moved by the window's own top-left before it is drawn (see
        // `menu_popup_shifted`).
        let popup = pin_chrome::menu_popup_shifted(&paint.popup, -left, -top);
        pin_chrome::paint_menu_popup(
            out,
            width,
            &palette,
            &popup,
            &paint.rows,
            paint.hover,
            surface,
            scale,
        );

        // And the flyout, beside the menu it came from: the two panels do not overlap,
        // so what the first one carried across the surface left the second's own rows
        // alone.
        if let Some(flyout) = paint.flyout.as_ref() {
            let flyout_popup = pin_chrome::menu_popup_shifted(&flyout.popup, -left, -top);
            pin_chrome::paint_menu_popup(
                out,
                width,
                &palette,
                &flyout_popup,
                &flyout.rows,
                flyout.hover,
                surface,
                scale,
            );
        }
    });

    let Some(mem_dc) = layered_surface_dc(hwnd.0 as isize) else {
        return;
    };
    let dst_point = POINT { x: left, y: top };
    let size = SIZE {
        cx: width,
        cy: height,
    };
    let src_point = POINT { x: 0, y: 0 };
    let blend = BLENDFUNCTION {
        BlendOp: AC_SRC_OVER as u8,
        BlendFlags: 0,
        SourceConstantAlpha: 255,
        AlphaFormat: AC_SRC_ALPHA as u8,
    };

    let _ = UpdateLayeredWindow(
        hwnd,
        None,
        Some(&dst_point),
        Some(&size),
        mem_dc,
        Some(&src_point),
        COLORREF(0),
        Some(&blend),
        ULW_ALPHA,
    );
}

/// The card's menu window, created the first time a menu is put up and kept until it
/// is taken away: a window is not made for every right-click, and a menu is not the
/// only thing the window is asked to be sized to.
///
/// `None` where the app has no preview window of its own, which is what a test stands
/// the card in: the menu window is that window's sibling, and there is nothing to be a
/// sibling of in a test.
fn pin_menu_window() -> Option<HWND> {
    if PREVIEW_HWND.load(Ordering::SeqCst) == 0 {
        return None;
    }

    let existing = PIN_MENU_HWND.load(Ordering::SeqCst);
    if existing != 0 {
        return Some(HWND(existing as *mut _));
    }

    let hwnd = unsafe {
        let hinstance = GetModuleHandleW(None).ok()?;
        CreateWindowExW(
            WS_EX_LAYERED | WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_NOACTIVATE | WS_EX_TRANSPARENT,
            PIN_MENU_CLASS,
            w!("Pinned menu"),
            WS_POPUP,
            0,
            0,
            1,
            1,
            None,
            None,
            hinstance,
            None,
        )
        .ok()?
    };

    PIN_MENU_HWND.store(hwnd.0 as isize, Ordering::SeqCst);
    Some(hwnd)
}

/// Take the card's menu window away, if one is up: the menu is over, and a window
/// nothing will ever paint again is a panel the desktop keeps showing.
///
/// It is hidden rather than destroyed, which is what the bubble's own hide is: a window of
/// this app's own, asked to be gone from whichever road found the pin was, and the
/// watchdog's road is a thread that does not own the windows (see `hide_pin_bubble`).
pub(super) unsafe fn hide_pin_menu_window() {
    let hwnd = PIN_MENU_HWND.load(Ordering::SeqCst);
    if hwnd == 0 {
        return;
    }

    let _ = ShowWindow(HWND(hwnd as *mut _), SW_HIDE);
}

/// The menu window's own procedure, which answers nothing: the card's window holds the
/// pointer while its menu is up, so a press anywhere on the desktop is delivered there
/// and answered against the panel's own rows (see `pin_menu::pinned_menu_press`).
pub(super) unsafe extern "system" fn pin_menu_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    DefWindowProcW(hwnd, msg, wparam, lparam)
}

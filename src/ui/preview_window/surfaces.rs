//! The layered surfaces a preview is painted onto: the two windows of somebody else's memory a
//! frame is written into, and the device contexts they are read back through.

use super::*;

/// Reusable layered-window surface: one memory DC with one DIB section selected
/// into it, replaced only when the preview dimensions change.
///
/// A repaint used to create and destroy both objects and allocate a fresh
/// `width * height * 4` block for every animation frame. The surface belongs to
/// the preview thread, which owns the window and every repaint.
pub(super) struct LayeredSurface {
    pub(super) mem_dc: HDC,
    pub(super) bitmap: HBITMAP,
    pub(super) old_bitmap: HGDIOBJ,
    pub(super) bits: *mut u8,
    pub(super) width: u32,
    pub(super) height: u32,
}

impl Drop for LayeredSurface {
    fn drop(&mut self) {
        unsafe {
            if !self.old_bitmap.0.is_null() {
                let _ = SelectObject(self.mem_dc, self.old_bitmap);
            }
            let _ = DeleteObject(self.bitmap);
            let _ = DeleteDC(self.mem_dc);
        }
    }
}

thread_local! {
    /// The surfaces this thread paints layered windows on, one per window and kept between
    /// repaints. The windows at most are six, and all of them are painted by the preview
    /// thread: the preview — which is the pinned window while a pin is up — the round bubble a
    /// collapsed pin leaves, the window of the card's own over the screen, and the three strips a
    /// hosted film's pin carries over its film — the caption's, the transport's, and the volume
    /// popup's while it is open (see `pin_chrome_strips`). A surface rebuilt per repaint would be
    /// a display's worth of pixels allocated for every frame of a video.
    pub(super) static LAYERED_SURFACES: RefCell<Vec<(isize, LayeredSurface)>> = const { RefCell::new(Vec::new()) };
    /// Mutable box for the corner spinner: the small piece of a frame the arc is drawn into
    /// and then composed back over the frame's own pixels, so that drawing the spinner costs
    /// a copy of the box rather than a copy of the frame — which at the size of a display is
    /// the difference between a few kilobytes and thirty megabytes, every eighty milliseconds
    /// (see `render_layered_preview_at`).
    pub(super) static OVERLAY_SCRATCH: RefCell<Vec<u8>> = const { RefCell::new(Vec::new()) };
    /// The membrane a pinned window's caption is painted on: the strip is drawn by
    /// `pin_chrome`, which is GDI's business, into a surface of its own and copied onto the
    /// window's, which is this module's. It is kept between paints for the reason the windows'
    /// surfaces are — a pinned video repaints sixty times a second.
    pub(super) static CHROME_SURFACE: RefCell<Option<DibSurface>> = const { RefCell::new(None) };
    /// The membrane a pinned window's transport bar is painted on, kept for the reason the
    /// caption's is: a pinned video repaints while it plays, and the two strips are painted one
    /// after the other — so they cannot share one surface without one of them being rebuilt.
    pub(super) static TRANSPORT_SURFACE: RefCell<Option<DibSurface>> = const { RefCell::new(None) };
    /// The membrane a pinned window's volume popup is painted on, for the reason the caption's and
    /// the transport's are: it is carried by a strip of its own over a hosted film, and it is only
    /// the popup's footprint rather than the whole window. It is cleared before each use, because
    /// what is carried off it is the popup on a transparency the strip shows through.
    pub(super) static VOLUME_SURFACE: RefCell<Option<DibSurface>> = const { RefCell::new(None) };
    /// The membrane a caption's tooltip is written on. It is the size of the whole window rather
    /// than of a strip, because a name is measured and drawn through GDI and needs a device
    /// context of its own, and it is kept between paints for the reason the surfaces above are:
    /// a pinned video repaints sixty times a second, and a name is up for as long as a pointer
    /// rests on a button.
    ///
    /// It is cleared before each use, because what is carried off it is decided by the alpha
    /// byte GDI leaves behind — a run's box is sealed opaque and the rest of the surface is
    /// nothing, and a name from the last frame left opaque on this one would be a second name
    /// drawn over the first.
    pub(super) static TOOLTIP_SURFACE: RefCell<Option<DibSurface>> = const { RefCell::new(None) };
    /// The membrane a card's menu is written on, for the reason the tooltip's is: the
    /// rows' labels are measured and drawn through GDI, which needs a device context of
    /// its own, and it is the size of the whole window for that reason rather than of a
    /// strip. It is kept between paints for the reason the surfaces above are, and
    /// cleared before each use for the reason the tooltip's is.
    pub(super) static MENU_SURFACE: RefCell<Option<DibSurface>> = const { RefCell::new(None) };
    /// The membrane the card's menu is written on in the window of its own over the screen, for
    /// the reason the tooltip's membrane is. It is the menu window's rather than the card's
    /// (`MENU_SURFACE`) because the two windows are painted in turn, and a surface rebuilt for
    /// each of them is one made and thrown away on every repaint of either. It is cleared
    /// before each use, for the reason the tooltip's is.
    pub(super) static MENU_WINDOW_SURFACE: RefCell<Option<DibSurface>> = const { RefCell::new(None) };
}

/// The last frame a pinned film's player had on screen, held as a frame of its own rather than
/// as the window it was drawn into.
///
/// A drag puts the player's window away for the whole of its own length, so what is left of the
/// pin's band is this: the picture that was on screen a pointer message ago, scaled to the box the
/// window is being dragged to. It is a copy of what was on screen rather than a decode of the
/// file, because the file can only be decoded by the player that was decoding it, and a drag is
/// not a moment a decode fits into (see `compose_parked_band`).
pub(super) struct HeldVideoFrame {
    pub(super) pixels: Vec<u8>,
    pub(super) width: u32,
    pub(super) height: u32,
}

/// The one held frame, behind a lock rather than a thread-local because the two ends of it are
/// not the same thread by construction: it is written on the pointer message that begins a drag
/// and read by every repaint of that drag, which is a message per move and a repaint per resize
/// box, and only one of those is the pinned window's own procedure.
///
/// It is held for the length of a park and given up when the park is over, so a pin that is not
/// being dragged pays nothing for it (see `settle_pinned_park`).
pub(super) static HELD_VIDEO_FRAME: Lazy<Mutex<Option<HeldVideoFrame>>> =
    Lazy::new(|| Mutex::new(None));

/// The held frame, behind its lock, so that a paint scaling a picture the size of a display does
/// not copy one out from under itself (see [`HELD_VIDEO_FRAME`]).
pub(super) fn held_video_frame() -> MutexGuard<'static, Option<HeldVideoFrame>> {
    HELD_VIDEO_FRAME
        .lock()
        .unwrap_or_else(|held| held.into_inner())
}

/// Stand a held frame in, for a test about what the band's picture is rather than about how it was
/// taken off the desktop — which needs a player of this app's, and this machine has none.
#[cfg(test)]
pub(super) fn stand_video_frame_for_a_test(pixels: Vec<u8>, width: u32, height: u32) {
    if let Ok(mut held) = HELD_VIDEO_FRAME.lock() {
        *held = Some(HeldVideoFrame {
            pixels,
            width,
            height,
        });
    }
}

/// Give the held frame up: the player's own window is on screen again, so the band's picture is
/// that window's rather than anything kept here.
pub(super) fn forget_video_frame() {
    if let Ok(mut held) = HELD_VIDEO_FRAME.lock() {
        *held = None;
    }
}

/// Copy the player's window into the frame a drag paints from, answering whether there was a
/// picture to take.
///
/// It is read off the desktop with `BitBlt` rather than asked of the window with `PrintWindow`,
/// and the reason is what is being taken: this is the frame the user is looking at, so it is read
/// where the frame is. A window that draws through D3D — which FFmpeg's player does whenever the
/// decoder is hardware-accelerated — answers `PrintWindow` with a black rectangle or with nothing,
/// which is the one thing this cannot be.
///
/// A failure answers false rather than a blank frame: a black band painted over a player's window
/// that is still there would be worse than the frame, and the park below knows the difference
/// between a picture and no picture (see `compose_parked_band`).
pub(super) fn hold_video_window_frame() -> bool {
    // **A film the pin draws through its host has no window of the desktop's of its own to find by
    // pid**: the picture is the Shell's thumbnail of the player's window, composed into the host,
    // so what is on screen in the band is the host itself and that is the rect the frame is read
    // off. Everything else is a player's own window over the preview, found the way it always was
    // (see `pin_film_host`).
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    let window = if pin_film_host::the_pinned_player_was_begun_in_the_host() {
        pin_film_host::film_host()
    } else if hover_film_host::the_hover_player_was_begun_in_the_host(pid) {
        // The hover's host is the visible surface, with the film composed into it (see
        // `hover_film_host`).
        hover_film_host::hover_film_host()
    } else {
        video_window_for(pid)
    };
    let Some(hwnd) = window else {
        return false;
    };

    let mut rect = RECT::default();
    // SAFETY: `hwnd` came out of `video_window_for`, which settles `IsWindow` on the handle it
    // publishes, and `RECT` is this frame's own. `GetWindowRect` reports rather than faults for
    // a window that is on its way out, which is what a player ending under a drag is.
    unsafe {
        if GetWindowRect(hwnd, &mut rect).is_err() {
            return false;
        }
    }

    // A player begun with `-noborder` has no frame of its own around the picture, so the window's
    // rect is the picture's — and the band is that rect, which is what makes this the frame of the
    // band rather than of something near it.
    let Some((mut pixels, width, height)) = screen_pixels(rect) else {
        return false;
    };

    // GDI writes colour and not alpha, and a layered window is composited from premultiplied
    // coverage: this frame is opaque everywhere because a player's window has nothing behind
    // it to be transparent over.
    for pixel in pixels.as_chunks_mut::<4>().0 {
        pixel[3] = 255;
    }

    if let Ok(mut held) = HELD_VIDEO_FRAME.lock() {
        *held = Some(HeldVideoFrame {
            pixels,
            width,
            height,
        });
    }

    true
}

/// The pixels the desktop is showing in a screen rectangle, or nothing where there is no picture to
/// take.
///
/// It is read off the desktop with `BitBlt` rather than asked of a window with `PrintWindow`, and the
/// reason is what is being taken: a picture the user is looking at is read where it is. A window that
/// draws through D3D — which FFmpeg's player does whenever the decoder is hardware-accelerated, which
/// a page's own window does, and which a layered window like a preview does whatever it holds —
/// answers `PrintWindow` with a black rectangle or with nothing, and that is the one thing a picture
/// of a pin cannot be (see `hold_video_window_frame`, `take_the_pin_picture`).
///
/// GDI writes colour and not alpha, since the desktop has nothing behind it to be transparent over: a
/// caller that composites these pixels premultiplied has to fill it in.
pub(super) fn screen_pixels(area: RECT) -> Option<(Vec<u8>, u32, u32)> {
    let width = (area.right - area.left).max(0) as u32;
    let height = (area.bottom - area.top).max(0) as u32;
    if width == 0 || height == 0 {
        return None;
    }

    let surface = DibSurface::create(width, height)?;

    // SAFETY: the surface is this call's own, of `width` by `height` pixels, and the screen DC is
    // released on the way out of the block whatever `BitBlt` answers.
    unsafe {
        let screen = windows::Win32::Graphics::Gdi::GetDC(HWND(std::ptr::null_mut()));
        let copied = windows::Win32::Graphics::Gdi::BitBlt(
            surface.dc,
            0,
            0,
            width as i32,
            height as i32,
            screen,
            area.left,
            area.top,
            SRCCOPY,
        )
        .is_ok();
        let _ = windows::Win32::Graphics::Gdi::ReleaseDC(HWND(std::ptr::null_mut()), screen);
        // GDI batches its calls and what is read next is not a GDI call.
        let _ = GdiFlush();

        if !copied {
            return None;
        }

        let pixels =
            std::slice::from_raw_parts(surface.bits(), width as usize * height as usize * 4)
                .to_vec();
        Some((pixels, width, height))
    }
}

/// How many windows keep a surface of their own: the preview — pinned or not — the bubble,
/// the card's menu, and the three strips a hosted film's pin carries over it.
pub(super) const LAYERED_SURFACE_WINDOWS: usize = 6;

/// DIB bits for a `width` x `height` frame in one window's own surface, reusing what
/// that window already has when the size is unchanged. `None` means the surface could
/// not be created and the frame is skipped, as a failed `CreateDIBSection` did before.
pub(super) fn ensure_layered_surface(window: isize, width: u32, height: u32) -> Option<*mut u8> {
    LAYERED_SURFACES.with(|cell| {
        let mut surfaces = cell.borrow_mut();

        if let Some((_, existing)) = surfaces.iter().find(|(key, _)| *key == window) {
            if existing.width == width && existing.height == height {
                return Some(existing.bits);
            }
        }

        // Dropping the previous surface releases its DC and bitmap — this window's only,
        // since the surface of the other one is left where it is.
        surfaces.retain(|(key, _)| *key != window);
        while surfaces.len() >= LAYERED_SURFACE_WINDOWS {
            surfaces.remove(0);
        }

        let bmi = BITMAPINFO {
            bmiHeader: BITMAPINFOHEADER {
                biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
                biWidth: width as i32,
                biHeight: -(height as i32),
                biPlanes: 1,
                biBitCount: 32,
                biCompression: BI_RGB.0,
                biSizeImage: 0,
                biXPelsPerMeter: 0,
                biYPelsPerMeter: 0,
                biClrUsed: 0,
                biClrImportant: 0,
            },
            bmiColors: [Default::default()],
        };

        unsafe {
            let mem_dc = CreateCompatibleDC(None);
            if mem_dc.0.is_null() {
                return None;
            }

            let mut bits: *mut core::ffi::c_void = ptr::null_mut();
            let Ok(bitmap) = CreateDIBSection(mem_dc, &bmi, DIB_RGB_COLORS, &mut bits, None, 0)
            else {
                let _ = DeleteDC(mem_dc);
                return None;
            };

            if bits.is_null() {
                let _ = DeleteObject(bitmap);
                let _ = DeleteDC(mem_dc);
                return None;
            }

            // Kept selected for the surface's lifetime: `UpdateLayeredWindow`
            // reads the bitmap through this DC.
            let old_bitmap = SelectObject(mem_dc, bitmap);

            surfaces.push((
                window,
                LayeredSurface {
                    mem_dc,
                    bitmap,
                    old_bitmap,
                    bits: bits as *mut u8,
                    width,
                    height,
                },
            ));

            Some(bits as *mut u8)
        }
    })
}

/// The memory DC of the surface this thread last ensured for a window, which is what
/// `UpdateLayeredWindow` reads the bits through.
pub(super) fn layered_surface_dc(window: isize) -> Option<HDC> {
    LAYERED_SURFACES.with(|cell| {
        cell.borrow()
            .iter()
            .find(|(key, _)| *key == window)
            .map(|(_, surface)| surface.mem_dc)
    })
}

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
    /// repaints. There are two windows at most, and both are painted by the preview thread:
    /// the preview — which is the pinned window while a pin is up — and the round bubble a
    /// collapsed pin leaves. A surface rebuilt per repaint would be a display's worth of
    /// pixels allocated for every frame of a video.
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
    /// The frame a pinned window is being dragged to is scaled out of: the media's own pixels, in
    /// a surface GDI can read, which is what lets a band of a size the frame is not be filled by
    /// `StretchBlt` rather than by a pixel at a time. It is kept between repaints for the reason
    /// the two strips above are — an edge under a hand is a repaint per pointer move — and its
    /// size follows the frame, not the box being dragged to.
    pub(super) static BAND_SOURCE: RefCell<Option<DibSurface>> = const { RefCell::new(None) };
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
    let Some(hwnd) = video_window_for(VIDEO_PID.load(Ordering::SeqCst)) else {
        return false;
    };

    let mut rect = RECT::default();
    // SAFETY: `hwnd` came out of `video_window_for`, which settles `IsWindow` on the handle it
    // publishes, and `RECT` is this frame's own. `GetWindowRect` reports rather than faults for
    // a window that is on its way out, which is what a player ending under a drag is.
    let (width, height) = unsafe {
        if GetWindowRect(hwnd, &mut rect).is_err() {
            return false;
        }
        (
            (rect.right - rect.left).max(0) as u32,
            (rect.bottom - rect.top).max(0) as u32,
        )
    };

    // A player begun with `-noborder` has no frame of its own around the picture, so the window's
    // rect is the picture's — and the band is that rect, which is what makes this the frame of the
    // band rather than of something near it.
    if width == 0 || height == 0 {
        return false;
    }

    let Some(surface) = DibSurface::create(width, height) else {
        return false;
    };

    let wanted = width as usize * height as usize * 4;
    // SAFETY: the surface is this frame's own, of `width` by `height` pixels, and the screen DC is
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
            rect.left,
            rect.top,
            SRCCOPY,
        )
        .is_ok();
        let _ = windows::Win32::Graphics::Gdi::ReleaseDC(HWND(std::ptr::null_mut()), screen);
        // GDI batches its calls and what is read next is not a GDI call.
        let _ = GdiFlush();

        if !copied {
            return false;
        }

        let mut pixels = std::slice::from_raw_parts(surface.bits(), wanted).to_vec();
        // GDI writes colour and not alpha, and a layered window is composited from premultiplied
        // coverage: this frame is opaque everywhere because a player's window has nothing behind
        // it to be transparent over (see `stretch_into_band` for the same correction).
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
    }

    true
}

/// How many windows keep a surface of their own: the preview — pinned or not — and the bubble.
pub(super) const LAYERED_SURFACE_WINDOWS: usize = 2;

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

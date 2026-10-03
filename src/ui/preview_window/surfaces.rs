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

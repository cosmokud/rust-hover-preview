//! What a taskbar button's hover shows of the pin itself.
//!
//! The hover of a window is a live picture of that window: the Shell takes what the Desktop Window
//! Manager composed for it and draws it small. That is why a playing player's hover is the film and
//! a minimized one's is the last frame it showed — nothing the player does, and nothing this app can
//! do either, unless there is a window of the pin's to be a picture of.
//!
//! There is one. The stand-in the taskbar button belongs to is given the pin's own content by a DWM
//! thumbnail — the Shell drawing one window inside another (`DwmRegisterThumbnail`) — so a pin's
//! hover is the picture the Shell would have drawn had the pin itself been the window in the
//! taskbar, and a film inside it moves without this app drawing a frame.
//!
//! Which window is asked first, because it is not always the preview window: a film FFmpeg plays and
//! a document the engine draws have windows of their own standing in the pin's band, and the preview
//! window is a hole there with a caption and a bar around it (see `place_pinned_siblings`).
//!
//! A mirror is registered once per source and left alone. What is asked of the Shell again is asked
//! when the source, its process or the box changes — a pin swapped to another file, a window resized,
//! a player relaunched into the very number of the window it replaced — and the answer is a window,
//! a process and a box, so a mirror costs this app three comparisons a tick.
//!
//! A sound is pictured by the pin's own window, which is where its card is drawn: the peek of a
//! sound is its card, the way the peek of a picture is the picture. The three buttons beside it are
//! what a sound put away has to *offer*; they are not a reason for its hover to hold nothing (see
//! `mirror_source` and `pin_taskbar_buttons`).
//!
//! shortcut: the pin's own window is a layered one, and what the Shell composes for a layered window
//! as a thumbnail source is not something this app can ask without a screen to look at. If a picture,
//! a page or a card turns out to be a hover with nothing in it, the fallback is to draw the pin's
//! own surface into the stand-in instead of mirroring it — the surface is already a DIB this app
//! owns (see `DibSurface::bitmap` and `render_layered_preview`), so no mirror is needed for it.
//!
//! **And a hover of a *minimized* window is not a mirror at all.** What the Shell holds of a window
//! that has gone down is a still of the window's own rows — the last thing the app drew — and a mirror
//! this app registered *into* the stand-in is no part of those rows: it is composed over them while
//! the window is up, and it goes with the window's going down. The stand-in therefore *supplies* that
//! still itself: it is made wearing `DWMWA_HAS_ICONIC_BITMAP` (see `pin_taskbar`), and when the Shell
//! asks for the thumbnail of a button that is down, the pin's own picture is handed over — a picture
//! taken off the desktop while the pin's windows were still on the screen (see `take_the_pin_picture`,
//! `answer_the_thumbnail_request`). What the Shell holds otherwise is its own default for a window
//! with nothing of its own: the window's icon, stretched into whatever box the stand-in was sized to,
//! which is what a put-away pin's hover showed for every kind of pin.

use super::*;

use windows::Win32::Foundation::RECT;
use windows::Win32::Graphics::Dwm::{
    DwmFlush, DwmInvalidateIconicBitmaps, DwmRegisterThumbnail, DwmSetIconicThumbnail,
    DwmUnregisterThumbnail, DwmUpdateThumbnailProperties, DWM_THUMBNAIL_PROPERTIES,
    DWM_TNP_RECTDESTINATION, DWM_TNP_SOURCECLIENTAREAONLY, DWM_TNP_VISIBLE,
};
use windows::Win32::Graphics::Gdi::{SetBrushOrgEx, SetStretchBltMode, StretchBlt, HALFTONE};
use windows::Win32::UI::WindowsAndMessaging::{GetClientRect, IsWindow, IsWindowVisible};

/// How large a mirrored pin may be on its longest side.
///
/// The Shell draws a hover at its own size whatever the window holds, so this is the ceiling on what
/// a hover costs in pixels: a display at 100% draws a thumbnail a few hundred pixels wide, one at
/// 300% wants three times that, and 640 covers the last of them with room for a pin that is not a
/// widescreen picture.
const PIN_MIRROR_MAX_PIXELS: i32 = 640;

/// The stand-in's mirror: the window the Shell is drawing into it, the size its own window was left
/// at, and the registration the Shell answered with.
///
/// The three travel together — a new source is a new registration, and a new size is a new box for
/// the picture to be drawn in — so they are one value rather than three flags that could disagree.
static PIN_MIRROR: Mutex<PinMirror> = Mutex::new(PinMirror::NOTHING);

/// What the stand-in is showing, which is `NOTHING` for a stand-in with no pin in it.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
struct PinMirror {
    /// The source window as the number of its handle, or zero while nothing is mirrored.
    source: isize,
    /// The process the source window belongs to, which is what tells a window that has been made
    /// again from the one this registration was taken with: a handle is a slot the system hands out
    /// again, and a player ended and begun inside one tick can be given the very number of the window
    /// it replaced (see `window_process`).
    source_process: u32,
    /// The size the stand-in's *client area* was left at, which is the box the picture goes in.
    client: (i32, i32),
    /// What the Shell answered with when it took the registration, or zero.
    thumbnail: isize,
}

impl PinMirror {
    /// A mirror of nothing: no source, no registration, and the square the stand-in is created in —
    /// which is the square it keeps until there is a picture to put in it, since a picture once
    /// taken is never given up for a box with nothing in it (see `sync_pin_mirror`).
    const NOTHING: Self = Self {
        source: 0,
        source_process: 0,
        client: (1, 1),
        thumbnail: 0,
    };

    /// Whether this is already that window, of that process, in that box, which is the whole of what
    /// the Shell has been told — so the whole of what deciding to ask it again costs.
    fn already_shows(&self, source: isize, source_process: u32, client: (i32, i32)) -> bool {
        self.source == source && self.source_process == source_process && self.client == client
    }
}

/// Give the stand-in the pin's own picture again, if anything about it has changed.
///
/// Called on every tick of the preview loop, and answered with a window, its process and the size of
/// its client area: nothing is asked of the Shell unless one of them is not what it was told, which
/// is what makes a moving mirror cost this app nothing per frame.
///
/// **A picture once taken is the picture to hold.** The window the pin is drawn by is off the screen
/// on roads of its own — a display or DPI change takes the preview down before the hover it came from
/// is replayed, a player is ended and begun again, a media is swapped for another — and settling a
/// mirror for a window that is not there takes down the registration the picture stands in *and*
/// re-sizes the stand-in for a box with nothing in it: what the stand-in holds after that is its own
/// surface, which this app has never painted, so its hover is whatever the Shell's surface for that
/// window last had rather than the pin. What that costs is not only the hover of the moment: the
/// frame a put-away pin's button holds is taken off *that* surface when the stand-in goes down, so a
/// collapse in the middle of one of those roads is a button whose peek is empty for the rest of the
/// run. So a picture already taken is left standing until another can be taken in its place — which
/// for a pin put away is the same rule at its narrowest, since every window the pin's picture could
/// be taken from is off the screen by then (see `keep_the_pin_for_the_taskbar`, `mirror_source`).
pub(super) fn sync_pin_mirror(kind: Option<MediaType>) {
    let Some(hwnd) = pin_taskbar::showing_taskbar_window() else {
        return;
    };

    // Nothing of the pin's is on the screen to be pictured, so what the stand-in holds is the picture
    // to hold: it is left exactly as it stands, in the box it stands in.
    let Some(source) = mirror_source(kind) else {
        return;
    };

    // Safety: the window came out of `mirror_source`, which settles `IsWindow` on the handle it
    // answers with, and what is asked of it is the process it belongs to: a question about a window
    // rather than a change to it.
    let process = unsafe { window_process(source) };

    // Safety: the same window, asked for the size of its client area: a measurement rather than a
    // change, and nothing at all for a window that has gone.
    let Some(area) = (unsafe { client_size(source) }) else {
        return;
    };
    let client = mirror_size(area);

    let Ok(mut mirror) = PIN_MIRROR.lock() else {
        return;
    };

    if mirror.already_shows(source.0 as isize, process, client) {
        return;
    }

    // Safety: the destination is this app's own stand-in window, the source is a window this app is
    // showing the pin in, and the registration taken before this one is this module's own.
    unsafe { settle_mirror(hwnd, &mut mirror, source, process, client) };
}

/// Take the mirror down: there is no pin left whose button could be hovered.
///
/// It is the stand-in going away that asks for this, and not the pin: a registration left behind a
/// hidden window is a registration the Shell is still drawing, and one of those per pin is a leak of
/// the Shell's own resources.
pub(super) fn stop_pin_mirror() {
    let Ok(mut mirror) = PIN_MIRROR.lock() else {
        return;
    };

    if mirror.thumbnail == 0 {
        return;
    }

    // Safety: the number is the registration this module was answered with and has held since, and
    // it is taken down once and put out of the state in the same move.
    let _ = unsafe { DwmUnregisterThumbnail(mirror.thumbnail) };
    *mirror = PinMirror::NOTHING;
}

/// Keep the pin in its taskbar button: take the picture the button is to hold *before* the windows it
/// is a picture of are taken off the screen.
///
/// The whole of this function is where it is called from, which is the collapse into the taskbar
/// before the pin's windows go (see `collapse_pin_to`): a picture of a window that has been taken off
/// the screen is a picture of nothing, so it is taken while there is still one to take.
pub(super) unsafe fn keep_the_pin_for_the_taskbar() {
    pin_taskbar::show_pin_taskbar_mirroring();
    sync_pin_mirror(current_media_type());
    take_the_pin_picture();
    let _ = DwmFlush();
}

/// The picture the button holds of a pin that is put away: what the Shell's still of that button is
/// made of.
///
/// It is a picture of this app's own rather than a mirror, because the Shell draws the hover of a
/// window that is *down* from the bitmap the app hands it rather than from the window's composition —
/// and a mirror is composed over the stand-in rather than in it (see this module's header).
struct PinPicture {
    /// The picture, as the desktop was showing it (see `screen_pixels`).
    pixels: Vec<u8>,
    width: u32,
    height: u32,
}

/// The one picture the button holds: taken where a pin is put away, and handed to the Shell when it
/// asks what the button of a window that is down looks like.
///
/// Behind a lock rather than a thread-local because the two ends of it are not the same thread by
/// construction: one is the collapse, and the other is a message from the Shell's own clock.
static PIN_PICTURE: Mutex<Option<PinPicture>> = Mutex::new(None);

/// Take the pin's picture, while there is still a window of the pin's on the screen to take it of.
///
/// **This is what the hover of a put-away pin needs and never had.** What the Shell holds of a window
/// that has gone down is its own default for a window with nothing of its own — the window's icon,
/// stretched into the box the stand-in was sized to — because a stand-in whose only content is a
/// mirror of another window has no still of its own: the mirror is composed over it while it is up and
/// goes with the window's going down. The picture taken here is what the Shell is handed instead,
/// every time it asks (see `answer_the_thumbnail_request`).
///
/// It is taken of the same window the mirror is of, and at the same moment: while a pin is being put
/// away and its own windows are still on the screen to be read (see `keep_the_pin_for_the_taskbar`). A
/// window of the pin's that is not drawn where it stands — a player on its way out, a preview a
/// display change has taken down, a window the Shell has minimized — is no picture at all, and the
/// picture already taken is the one to hold, which is the rule the mirror itself keeps (see
/// `sync_pin_mirror`).
pub(super) fn take_the_pin_picture() {
    let Some(source) = mirror_source(current_media_type()) else {
        return;
    };

    // Safety: the window came out of `mirror_source`, which settles `IsWindow` on the handle it
    // answers with, and what is asked of it is where its client area stands on the screen — a
    // question about the window rather than a change to it.
    let Some(area) = (unsafe { client_area_on_the_screen(source) }) else {
        return;
    };

    let Some((pixels, width, height)) = screen_pixels(area) else {
        return;
    };

    if let Ok(mut picture) = PIN_PICTURE.lock() {
        *picture = Some(PinPicture {
            pixels,
            width,
            height,
        });
    }

    // The Shell keeps the bitmap it was last handed, and it does not know that this one is of another
    // file: it is told that its stills are out of date, and it asks for this one the next time a button
    // of this window is looked at.
    let Some(hwnd) = pin_taskbar::showing_taskbar_window() else {
        return;
    };

    // Safety: the window is this app's own stand-in, and what is asked of it is a renewal of its own
    // stills rather than a change to what it is showing.
    let _ = unsafe { DwmInvalidateIconicBitmaps(hwnd) };
}

/// Give the picture up: there is no pin left whose button could hold it.
///
/// It is the button going that asks for this and not the pin: a window made again for the next pin is
/// a button the picture of the pin before it must not be handed for (see `close_pin_taskbar`,
/// `pin_taskbar_window`).
pub(super) fn forget_the_pin_picture() {
    if let Ok(mut picture) = PIN_PICTURE.lock() {
        *picture = None;
    }
}

/// Answer the Shell's ask for the still of a button that is down, with the pin's own picture.
///
/// The Shell asks only while it is showing an *iconic* representation of the stand-in — which is every
/// moment a pin is put away into the taskbar — and it asks because the window told it there would be a
/// bitmap to be had (see `pin_taskbar`'s `pin_taskbar_window`). The box it asked in is `requested`: a
/// bitmap past either side of it is refused outright, and a bitmap it is never given leaves the button
/// with the Shell's own default, which is the window's icon (see `take_the_pin_picture`).
///
/// Safety: the window is this app's own stand-in, asked from its own window procedure.
pub(super) unsafe fn answer_the_thumbnail_request(hwnd: HWND, requested: (i32, i32)) {
    let Ok(picture) = PIN_PICTURE.lock() else {
        return;
    };
    let Some(picture) = picture.as_ref() else {
        return;
    };

    let Some((width, height)) = thumbnail_size(requested, (picture.width, picture.height)) else {
        return;
    };
    let Some(bitmap) = thumbnail_bitmap(picture, width, height) else {
        return;
    };

    // Safety: the bitmap is this call's own, made for this ask, and the window is the stand-in the ask
    // arrived at. The Shell copies what it is given, so the bitmap goes with this call's own surface
    // rather than being kept (see `DibSurface`).
    let _ = DwmSetIconicThumbnail(hwnd, bitmap.bitmap(), 0);
}

/// What the button holds, for a test asking which picture a put-away pin's button is showing.
#[cfg(test)]
pub(super) fn the_pin_picture() -> Option<(Vec<u8>, u32, u32)> {
    let picture = PIN_PICTURE.lock().ok()?;
    let picture = picture.as_ref()?;
    Some((picture.pixels.clone(), picture.width, picture.height))
}

/// The box a thumbnail is made in: the picture's own shape, brought inside the box the Shell asked
/// for and never past it.
///
/// A picture already smaller than the ask is handed over at its own size, since the Shell draws what
/// it is given into the room it has. Nothing is ever made larger: the Shell refuses a thumbnail with a
/// dimension past the box it asked in.
fn thumbnail_size(requested: (i32, i32), picture: (u32, u32)) -> Option<(u32, u32)> {
    let (asked_width, asked_height) = requested;
    if asked_width <= 0 || asked_height <= 0 {
        return None;
    }

    let (asked_width, asked_height) = (asked_width as u32, asked_height as u32);
    let (width, height) = (picture.0.max(1), picture.1.max(1));
    if width <= asked_width && height <= asked_height {
        return Some((width, height));
    }

    let share = (asked_width as f32 / width as f32).min(asked_height as f32 / height as f32);
    let scaled = |side: u32| (side as f32 * share) as u32;
    Some((scaled(width).max(1), scaled(height).max(1)))
}

/// The bitmap a picture is handed over in: the picture scaled into `width` x `height`, opaque
/// everywhere.
///
/// What the Shell is handed is a bitmap of this app's own making, which the Shell copies, so it goes
/// with this call's own surface rather than being kept (see `DibSurface`).
fn thumbnail_bitmap(picture: &PinPicture, width: u32, height: u32) -> Option<DibSurface> {
    let source = DibSurface::create(picture.width, picture.height)?;

    // Safety: the surface was made by this call, of the picture's own size, and nothing else has a
    // pointer to what is behind it.
    unsafe {
        std::ptr::copy_nonoverlapping(picture.pixels.as_ptr(), source.bits(), picture.pixels.len());
    }

    let target = DibSurface::create(width, height)?;
    // Safety: both surfaces were made by this call, and what is asked of the device contexts is a
    // scaling copy between them.
    unsafe {
        // GDI's own filtering rather than a nearest-neighbour copy: a picture taken of a display's
        // worth of pixels is brought down to a box a few hundred wide, and a hover is read at a
        // glance. The brush origin travels with the mode: halftone measures the destination from it.
        let _ = SetStretchBltMode(target.dc, HALFTONE);
        let _ = SetBrushOrgEx(target.dc, 0, 0, None);
        if !StretchBlt(
            target.dc,
            0,
            0,
            width as i32,
            height as i32,
            source.dc,
            0,
            0,
            picture.width as i32,
            picture.height as i32,
            SRCCOPY,
        )
        .as_bool()
        {
            return None;
        }

        // GDI writes colour and not alpha, and what the Shell composites is premultiplied: a peek of a
        // window is opaque everywhere, which is what a picture read off the desktop is (see
        // `screen_pixels`).
        let rows =
            std::slice::from_raw_parts_mut(target.bits(), width as usize * height as usize * 4);
        for pixel in rows.as_chunks_mut::<4>().0 {
            pixel[3] = 255;
        }
    }

    Some(target)
}

/// Where a window's client area stands on the screen, or nothing for a window with no area to be a
/// picture of.
///
/// The *client* area rather than the window's own rect, because that is what the mirror is of: what a
/// player's window wears around the picture is not the pin (see `settle_mirror`). A window the Shell
/// has minimized answers with no area at all, which is the answer that keeps the picture already
/// taken — a minimized window is not drawn where it stands, so nothing read off the desk there would
/// be the pin.
/// The same area for a caller outside this module's mirror, which is the relaunch measurement that
/// reads the pin off the desktop (see `hide_probe`).
pub(super) unsafe fn client_area_on_the_screen(hwnd: HWND) -> Option<RECT> {
    let mut area = RECT::default();
    GetClientRect(hwnd, &mut area).ok()?;
    let (width, height) = (area.right - area.left, area.bottom - area.top);
    if width <= 0 || height <= 0 {
        return None;
    }

    let mut origin = POINT {
        x: area.left,
        y: area.top,
    };
    if !ClientToScreen(hwnd, &mut origin).as_bool() {
        return None;
    }

    Some(RECT {
        left: origin.x,
        top: origin.y,
        right: origin.x + width,
        bottom: origin.y + height,
    })
}

/// Register the mirror of a window in the stand-in's client area, taking down the registration this
/// window wore before it.
///
/// Every road out leaves the state saying what the stand-in was left showing, including the roads
/// where the Shell refused: a registration that was not given is a hover with the desktop showing
/// through it, and that is what it was before this module existed rather than a state to retry from.
unsafe fn settle_mirror(
    hwnd: HWND,
    mirror: &mut PinMirror,
    source: HWND,
    source_process: u32,
    client: (i32, i32),
) {
    if mirror.thumbnail != 0 {
        let _ = DwmUnregisterThumbnail(mirror.thumbnail);
    }
    *mirror = PinMirror::NOTHING;

    // The window is the box the picture is drawn in, so it is moved before the picture is asked for:
    // the destination is the client area as it stands, and a box measured before a move is the box
    // the picture would have been drawn in a size ago.
    fit_mirror_window(hwnd, client);

    let Ok(thumbnail) = DwmRegisterThumbnail(hwnd, source) else {
        *mirror = PinMirror {
            source: 0,
            source_process: 0,
            client,
            thumbnail: 0,
        };
        return;
    };

    let mut area = RECT::default();
    if GetClientRect(hwnd, &mut area).is_err() {
        let _ = DwmUnregisterThumbnail(thumbnail);
        return;
    }

    let properties = DWM_THUMBNAIL_PROPERTIES {
        dwFlags: DWM_TNP_RECTDESTINATION | DWM_TNP_SOURCECLIENTAREAONLY | DWM_TNP_VISIBLE,
        // The whole of the stand-in's client area, which is why the window was sized to the source's
        // own shape: what the Shell draws is the source stretched into this box.
        rcDestination: area,
        rcSource: RECT::default(),
        opacity: 255,
        fVisible: true.into(),
        // A player's window is a frame around a picture, and the picture is its client area: what is
        // mirrored is what is being played rather than what the window is wearing.
        fSourceClientAreaOnly: true.into(),
    };

    if DwmUpdateThumbnailProperties(thumbnail, &properties).is_err() {
        let _ = DwmUnregisterThumbnail(thumbnail);
        return;
    }

    *mirror = PinMirror {
        source: source.0 as isize,
        source_process,
        client,
        thumbnail,
    };
}

/// Move the stand-in so that its *client area* is the size the picture needs, off every display.
///
/// The client area and not the window, because the destination of a thumbnail is measured in that
/// area: what the stand-in wears around it is measured here rather than assumed from its style, and
/// the arithmetic leaves the client area the source's own shape for a borderless window — which is
/// what the stand-in is (see `pin_taskbar::stand_in_style`).
unsafe fn fit_mirror_window(hwnd: HWND, client: (i32, i32)) {
    let mut window = RECT::default();
    let mut area = RECT::default();
    let _ = GetWindowRect(hwnd, &mut window);
    let _ = GetClientRect(hwnd, &mut area);

    // What the window wears around its client area, as it stands: measured rather than worked out
    // from the style, which is the same answer for less code. A window that answers nothing is a
    // window with nothing around it, which is the frame of a window that has gone.
    let frame = (
        (window.right - window.left - (area.right - area.left)).max(0),
        (window.bottom - window.top - (area.bottom - area.top)).max(0),
    );

    let _ = MoveWindow(
        hwnd,
        pin_taskbar::TASKBAR_OFFSCREEN,
        pin_taskbar::TASKBAR_OFFSCREEN,
        client.0 + frame.0,
        client.1 + frame.1,
        false,
    );
}

/// The window the pin's picture is in, or nothing where the pin has none to take.
///
/// **Which window that is, is the kind's own answer first.** A film FFmpeg is playing and a document
/// or a specimen the engine is drawing have windows of their own standing in the pin's band — it is
/// those windows that hold the picture, and the band is a hole where they were — so a kind that
/// draws in a window of its own is pictured by that window or by nothing at all, and never by the
/// window the hole is in (see `MediaType::draws_in_a_window_of_its_own`). Everything else is drawn by
/// the pin's own window: a picture, a page, a card, a listing, and a film the media engine decodes
/// into this window's own frame. For those, a player or a browser still standing in the band is a
/// better picture than the window behind them is, so they are asked first (see
/// `place_pinned_siblings` and `render_layered_preview`).
///
/// **And a window that is not on the screen is no picture at all.** The Shell composes nothing for a
/// window that is not there, so a registration taken of one is a mirror the Shell draws nothing
/// into — and taking it is what takes down the registration the picture already standing in the
/// stand-in was taken with (see `sync_pin_mirror`). A preview a display change has taken down, a
/// player being relaunched and a handle this app has not cleared yet are all this case, and all of
/// them are answered by nothing here so that the picture already standing is the one that stays.
///
/// **A pin put away into the *bubble* is answered by the bubble**, which is the only window of the
/// pin's the collapse leaves on the screen, and is why the hover the bubble's own `To Taskbar`
/// belongs to shows the pin rather than the desk behind it (see `pin_bubble::showing_bubble_window`).
/// A pin put away into the *taskbar* has every window of its own off the screen by then, and the
/// picture its button holds is the one the collapse took while they were still up (see
/// `keep_the_pin_for_the_taskbar`).
pub(super) fn mirror_source(kind: Option<MediaType>) -> Option<HWND> {
    if pin_is_collapsed() {
        if let Some(bubble) = pin_bubble::showing_bubble_window() {
            return Some(bubble);
        }
    }

    let player = VIDEO_HWND.load(Ordering::SeqCst);
    let document = webview_preview::showing_hwnd();
    let preview = PREVIEW_HWND.load(Ordering::SeqCst);

    if let Some(kind) = kind.filter(|kind| {
        pinned_media_owner().is_some_and(|(path, _)| kind.draws_in_a_window_of_its_own(&path))
    }) {
        let own = if kind.is_engine() { document } else { player };
        return on_the_screen(HWND(own as *mut _));
    }

    [player, document, preview]
        .into_iter()
        .find_map(|window| on_the_screen(HWND(window as *mut _)))
}

/// A window the Shell is drawing, or nothing where it is not one: a handle that is no window at all
/// — a player that has ended, a handle this app has not cleared yet — and a window that has been
/// hidden both compose nothing, and both are answers here rather than handles to be mirrored.
fn on_the_screen(hwnd: HWND) -> Option<HWND> {
    if hwnd.is_invalid() {
        return None;
    }

    // Safety: the handle is one this app holds — its own window, its player's, or the engine's — and
    // both questions are about the window rather than changes to it.
    unsafe { (IsWindow(hwnd).as_bool() && IsWindowVisible(hwnd).as_bool()).then_some(hwnd) }
}

/// The size the stand-in's client area is wanted at for a source's own size: the source's shape,
/// brought down to `PIN_MIRROR_MAX_PIXELS` on its longest side.
///
/// The shape is kept exactly rather than the picture being given a box of this app's own choosing,
/// because what the Shell draws is the source stretched into the box it is given: a box of another
/// shape is a picture of another shape.
fn mirror_size(client: (i32, i32)) -> (i32, i32) {
    let (width, height) = client;
    let (width, height) = (width.max(1), height.max(1));
    let longest = width.max(height);
    if longest <= PIN_MIRROR_MAX_PIXELS {
        return (width, height);
    }

    let share = PIN_MIRROR_MAX_PIXELS as f32 / longest as f32;
    (
        ((width as f32 * share).round() as i32).max(1),
        ((height as f32 * share).round() as i32).max(1),
    )
}

/// A window's client area, or nothing for a window that has gone.
unsafe fn client_size(hwnd: HWND) -> Option<(i32, i32)> {
    let mut client = RECT::default();
    GetClientRect(hwnd, &mut client).ok()?;
    Some((client.right - client.left, client.bottom - client.top))
}

/// The process a window belongs to, which is half of what says *which* window a mirror is of.
///
/// A handle on its own is not that answer: a handle is a slot the system hands out again, so a player
/// that was ended and begun inside one tick — a walk stepping onto another film, a seek, a relaunch —
/// can be given the very number of the window it replaced, at the very box that window stood at. A
/// mirror compared by handle and box alone is then left registered to a window that is gone, and the
/// Shell draws nothing into the stand-in for it (see `PinMirror::already_shows`).
unsafe fn window_process(hwnd: HWND) -> u32 {
    let mut process: u32 = 0;
    // Safety: the handle is a window `on_the_screen` has already settled `IsWindow` on, and what is
    // asked of it is the process it belongs to: a question about the window, not a change to it.
    let _ = GetWindowThreadProcessId(hwnd, Some(&mut process));
    process
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A pin the size of a display is brought down to the ceiling with its shape kept: what the Shell
    /// draws is stretched into the box it is given, so a box of another shape is another picture.
    #[test]
    fn a_mirror_is_the_shape_of_what_it_shows_and_never_larger_than_the_ceiling() {
        assert_eq!(
            mirror_size((320, 200)),
            (320, 200),
            "a pin smaller than the ceiling is mirrored at its own size"
        );
        assert_eq!(
            mirror_size((PIN_MIRROR_MAX_PIXELS, 100)),
            (PIN_MIRROR_MAX_PIXELS, 100),
            "and one exactly at it is not touched"
        );

        let (width, height) = mirror_size((3840, 2160));
        assert_eq!(width, PIN_MIRROR_MAX_PIXELS, "a larger one is brought down");
        assert_eq!(height, 360, "and keeps the shape it was brought down from");

        let (width, height) = mirror_size((2160, 3840));
        assert_eq!(height, PIN_MIRROR_MAX_PIXELS);
        assert_eq!(width, 360, "in either direction, since a pin may be tall");

        // A source with no size at all — a window that has been created and not yet laid out — is
        // still a box, because a thumbnail drawn into a box of no width is a division nobody wants.
        assert_eq!(mirror_size((0, 0)), (1, 1));
    }

    /// The comparison the tick makes: a mirror asked for twice is asked for once.
    #[test]
    fn a_mirror_is_only_settled_again_when_what_it_shows_or_the_box_it_is_in_changes() {
        let holding = PinMirror {
            source: 0x1000,
            source_process: 40,
            client: (400, 300),
            thumbnail: 7,
        };

        assert!(holding.already_shows(0x1000, 40, (400, 300)));
        assert!(
            !holding.already_shows(0x2000, 40, (400, 300)),
            "another pin is another picture"
        );
        assert!(
            !holding.already_shows(0x1000, 41, (400, 300)),
            "and a window made again in the number of the one before it is another window: the \
             registration this stands for is one the Shell draws nothing into, so the picture has \
             to be taken again"
        );
        assert!(
            !holding.already_shows(0x1000, 40, (401, 300)),
            "and the same pin in another box is another picture too"
        );
        assert!(
            !PinMirror::NOTHING.already_shows(0x1000, 40, (400, 300)),
            "a stand-in showing nothing is never already showing the pin"
        );
    }
}

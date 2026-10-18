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
//! when the source or the box changes — a pin swapped to another file, a window resized, a player
//! relaunched — and the answer is two numbers, so a mirror costs this app two comparisons a tick.
//!
//! A sound is shown as nothing, deliberately: its card is a list of the file's own facts, and the
//! three buttons beside it are the whole of what a sound put away has to offer (see
//! `pin_taskbar_buttons`).
//!
//! shortcut: the pin's own window is a layered one, and what the Shell composes for a layered window
//! as a thumbnail source is not something this app can ask without a screen to look at. If a picture,
//! a page or a card turns out to be a hover with nothing in it, the fallback is to draw the pin's
//! own surface into the stand-in instead of mirroring it — the surface is already a DIB this app
//! owns (see `DibSurface::bitmap` and `render_layered_preview`), so no mirror is needed for it.

use super::*;

use windows::Win32::Foundation::RECT;
use windows::Win32::Graphics::Dwm::{
    DwmFlush, DwmRegisterThumbnail, DwmUnregisterThumbnail, DwmUpdateThumbnailProperties,
    DWM_THUMBNAIL_PROPERTIES, DWM_TNP_RECTDESTINATION, DWM_TNP_SOURCECLIENTAREAONLY,
    DWM_TNP_VISIBLE,
};
use windows::Win32::UI::WindowsAndMessaging::{GetClientRect, IsWindowVisible};

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
    /// The size the stand-in's *client area* was left at, which is the box the picture goes in.
    client: (i32, i32),
    /// What the Shell answered with when it took the registration, or zero.
    thumbnail: isize,
}

impl PinMirror {
    /// A mirror of nothing: no source, no registration, and the square the stand-in is kept in while
    /// there is nothing of the pin to show.
    const NOTHING: Self = Self {
        source: 0,
        client: (1, 1),
        thumbnail: 0,
    };

    /// Whether this is already that window in that box, which is the whole of what the Shell has
    /// been told — so the whole of what deciding to ask it again costs.
    fn already_shows(&self, source: isize, client: (i32, i32)) -> bool {
        self.source == source && self.client == client
    }
}

/// Give the stand-in the pin's own picture again, if anything about it has changed.
///
/// Called on every tick of the preview loop, and answered with two comparisons: what the pin is
/// drawn by and the size of that window. Nothing is asked of the Shell unless one of them is not
/// what it was told, which is what makes a moving mirror cost this app nothing per frame.
pub(super) fn sync_pin_mirror(kind: Option<MediaType>) {
    let Some(hwnd) = pin_taskbar::showing_taskbar_window() else {
        return;
    };

    let source = mirror_source(kind);

    // **A pin put away is pictured once, and by the collapse.** Every window the pin's picture could
    // be taken from after that is off the screen — the collapse puts them all away — and the Shell
    // composes nothing for a window that is not on the screen, so settling again takes the
    // registration the collapse took down, re-sizes the stand-in for a picture that is not there, and
    // leaves the hover with the stand-in's own rows to show. What the collapse took is the whole of
    // the picture there is (see `keep_the_pin_for_the_taskbar`, `mirror_source`).
    if source.is_none() && pin_is_collapsed() {
        return;
    }

    let handle = source.map_or(0, |window| window.0 as isize);

    // Safety: the handle is a window this app is showing the pin in — its own, its player's, or the
    // engine's — and what is asked of it is the size of its client area: a measurement rather than a
    // change, and nothing at all for a window that has gone.
    let client = mirror_size(source.and_then(|window| unsafe { client_size(window) }));

    let Ok(mut mirror) = PIN_MIRROR.lock() else {
        return;
    };

    if mirror.already_shows(handle, client) {
        return;
    }

    // Safety: the destination is this app's own stand-in window, the source is a window this app is
    // showing the pin in, and the registration taken before this one is this module's own.
    unsafe { settle_mirror(hwnd, &mut mirror, source, client) };
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

/// Keep the pin in its taskbar button: give the hover the picture it is to hold *before* the windows
/// it is a picture of are taken off the screen.
///
/// The whole of this function is where it is called from, which is the collapse into the taskbar
/// before the pin's windows go (see `collapse_pin_to`). What a hover shows of a window that has been
/// put away is the frame the Shell cached as the button's window was minimized, and that frame is
/// composed of windows that are about to be hidden — so the picture is taken while there is still
/// one to take, and the composition is waited for rather than raced: the Shell draws a thumbnail on
/// its own clock, and this is the one moment the app cannot wait for the next tick.
pub(super) unsafe fn keep_the_pin_for_the_taskbar() {
    pin_taskbar::show_pin_taskbar_mirroring();
    sync_pin_mirror(current_media_type());
    let _ = DwmFlush();
}

/// Register the mirror of a window in the stand-in's client area, or take the old one down when
/// there is no window to mirror.
///
/// Every road out leaves the state saying what the stand-in was left showing, including the roads
/// where the Shell refused: a registration that was not given is a hover with the desktop showing
/// through it, and that is what it was before this module existed rather than a state to retry from.
unsafe fn settle_mirror(
    hwnd: HWND,
    mirror: &mut PinMirror,
    source: Option<HWND>,
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

    let Some(source) = source else {
        return;
    };

    let Ok(thumbnail) = DwmRegisterThumbnail(hwnd, source) else {
        *mirror = PinMirror {
            source: 0,
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

/// The window the pin's picture is in, or nothing where the pin has none to mirror.
///
/// A film FFmpeg is playing and a document the engine is drawing are windows of their own standing
/// in the pin's band — it is those windows that hold the picture — and everything else is drawn by
/// the preview window: a picture, a page, a card, and a film the media engine decodes into this
/// window's own frame (see `place_pinned_siblings` and `render_layered_preview`).
///
/// **And a pin put away is answered by a window that is on the screen, and by nothing else.** The
/// bubble a collapse leaves is the picture of a pin that is *put away* — it is composed, it is
/// showing the pin's own picture, and it is why the hover the bubble's own `To Taskbar` belongs to
/// shows the pin rather than the desk behind it (see `pin_bubble::showing_bubble_window`). The
/// pin's own windows are a picture for the one instant they are still up, which is the instant the
/// collapse takes it (see `keep_the_pin_for_the_taskbar`), and not after: the collapse hides them,
/// and the Shell composes nothing for a window that is not on the screen — so a thumbnail of one is
/// a registration the Shell draws nothing into, and asking for it takes down the registration the
/// picture was taken with (see `sync_pin_mirror`).
///
/// A sound has no picture to mirror, and a pin that is showing nothing yet has nothing to show.
pub(super) fn mirror_source(kind: Option<MediaType>) -> Option<HWND> {
    match kind {
        None | Some(MediaType::Audio) => return None,
        Some(_) => {}
    }

    if pin_is_collapsed() {
        if let Some(bubble) = pin_bubble::showing_bubble_window() {
            return Some(bubble);
        }
    }

    let player = VIDEO_HWND.load(Ordering::SeqCst);
    let document = webview_preview::showing_hwnd();
    let preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let window = if player != 0 {
        HWND(player as *mut _)
    } else if document != 0 {
        HWND(document as *mut _)
    } else if preview != 0 {
        HWND(preview as *mut _)
    } else {
        return None;
    };

    // A pin put away has nothing of its own left on the screen but the bubble above, and that is not
    // up: what a window that went down with the collapse is a picture of is nothing.
    if pin_is_collapsed() && !unsafe { IsWindowVisible(window).as_bool() } {
        return None;
    }

    Some(window)
}

/// The size the stand-in's client area is wanted at for a source's own size: the source's shape,
/// brought down to `PIN_MIRROR_MAX_PIXELS` on its longest side — and the square a stand-in with
/// nothing to show is kept at.
///
/// The shape is kept exactly rather than the picture being given a box of this app's own choosing,
/// because what the Shell draws is the source stretched into the box it is given: a box of another
/// shape is a picture of another shape.
fn mirror_size(client: Option<(i32, i32)>) -> (i32, i32) {
    let Some((width, height)) = client else {
        return PinMirror::NOTHING.client;
    };

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

#[cfg(test)]
mod tests {
    use super::*;

    /// A pin the size of a display is brought down to the ceiling with its shape kept: what the Shell
    /// draws is stretched into the box it is given, so a box of another shape is another picture.
    #[test]
    fn a_mirror_is_the_shape_of_what_it_shows_and_never_larger_than_the_ceiling() {
        assert_eq!(
            mirror_size(Some((320, 200))),
            (320, 200),
            "a pin smaller than the ceiling is mirrored at its own size"
        );
        assert_eq!(
            mirror_size(Some((PIN_MIRROR_MAX_PIXELS, 100))),
            (PIN_MIRROR_MAX_PIXELS, 100),
            "and one exactly at it is not touched"
        );

        let (width, height) = mirror_size(Some((3840, 2160)));
        assert_eq!(width, PIN_MIRROR_MAX_PIXELS, "a larger one is brought down");
        assert_eq!(height, 360, "and keeps the shape it was brought down from");

        let (width, height) = mirror_size(Some((2160, 3840)));
        assert_eq!(height, PIN_MIRROR_MAX_PIXELS);
        assert_eq!(width, 360, "in either direction, since a pin may be tall");

        // A source with no size at all — a window that has been created and not yet laid out — is
        // still a box, because a thumbnail drawn into a box of no width is a division nobody wants.
        assert_eq!(mirror_size(Some((0, 0))), (1, 1));
        assert_eq!(mirror_size(None), PinMirror::NOTHING.client);
    }

    /// The comparison the tick makes: a mirror asked for twice is asked for once.
    #[test]
    fn a_mirror_is_only_settled_again_when_what_it_shows_or_the_box_it_is_in_changes() {
        let holding = PinMirror {
            source: 0x1000,
            client: (400, 300),
            thumbnail: 7,
        };

        assert!(holding.already_shows(0x1000, (400, 300)));
        assert!(
            !holding.already_shows(0x2000, (400, 300)),
            "another pin is another picture"
        );
        assert!(
            !holding.already_shows(0x1000, (401, 300)),
            "and the same pin in another box is another picture too"
        );
        assert!(
            !PinMirror::NOTHING.already_shows(0x1000, (400, 300)),
            "a stand-in showing nothing is never already showing the pin"
        );
    }
}

//! The pin's film drawn by the Desktop Window Manager rather than nested in a window of this app's.
//!
//! **A prototype, behind `RHP_MPV_THUMBNAIL_PIN`, and it should be deleted once the answer is in.**
//! mpv is launched with no `--wid` and makes a window of the desktop's of its own; the pin's film
//! host asks the Shell to draw that window into itself (`DwmRegisterThumbnail`), and what follows a
//! resize is one rectangle rather than two windows agreeing about a box.
//!
//! **Why the film has to be a window of the desktop's.** A thumbnail is taken of a window's
//! composition surface, and a `WS_CHILD` has none of its own — it is drawn into its parent's — so a
//! film that is to be thumbnailed cannot be mpv's child under the host. That is the whole of why
//! this road launches mpv without `--wid` (see `video_launch::mpv_args`).
//!
//! **Why a resize step asks nothing of mpv.** The Shell stretches the source's client area into the
//! destination rectangle whatever the two sizes are, so a step is this app moving its own two
//! windows and re-pointing the destination at the host's new client area (see
//! `the_band_was_placed`). mpv's window is asked for a box once, when the drag ends (see
//! `settle_the_source_at_the_band`), and that ask is `SWP_ASYNCWINDOWPOS` like every other call made
//! on a window of another process's here.
//!
//! **Why mpv's window is kept on the screen behind the host.** The Shell draws nothing for a source
//! that is hidden — the thumbnail of a `SW_HIDE` window is frozen at its last frame — so the source
//! is never hidden while the pin is up: it is put at the bottom of the z-order instead, where the
//! host and the pin are over it (see `adopt_the_source`), and it is taken off the screen only where
//! the host it is seen in is (see `set_the_source_visible`).
//!
//! shortcut: the source keeps the size a drag began at while the hand is on the edge, and the Shell
//! draws it into a destination of any size — so a hand that *shrinks* the box leaves the part of the
//! source the host no longer covers standing on the desktop until the drag ends. Upgrade to a coarse
//! resize per step, or to a source clipped by something this app owns, if that shows.

use std::sync::atomic::{AtomicIsize, AtomicU32, Ordering};
use std::sync::Mutex;

use windows::Win32::Foundation::{HWND, POINT, RECT};
use windows::Win32::Graphics::Gdi::ClientToScreen;
use windows::Win32::Graphics::Dwm::{
    DwmRegisterThumbnail, DwmUnregisterThumbnail, DwmUpdateThumbnailProperties,
    DWM_THUMBNAIL_PROPERTIES, DWM_TNP_RECTDESTINATION, DWM_TNP_SOURCECLIENTAREAONLY,
    DWM_TNP_VISIBLE,
};
use windows::Win32::UI::WindowsAndMessaging::{
    GetClientRect, GetForegroundWindow, GetWindowLongPtrW, GetWindowRect, GetWindowThreadProcessId,
    IsWindow, IsWindowVisible, SetForegroundWindow, SetWindowLongPtrW, SetWindowPos,
    ShowWindowAsync, GWL_EXSTYLE, HWND_BOTTOM, SW_HIDE, SWP_ASYNCWINDOWPOS, SWP_NOACTIVATE,
    SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER, SW_SHOWNOACTIVATE, WS_EX_APPWINDOW, WS_EX_NOACTIVATE,
    WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_EX_TRANSPARENT,
};

use super::*;

/// The source window as the number of its handle, read without the lock by the questions the tick
/// asks on every pass (see `the_source_is_up`); zero while there is none.
static SOURCE: AtomicIsize = AtomicIsize::new(0);
/// The process the source window belongs to, which is what tells a window given the number of one it
/// replaced from the one this registration was taken with.
static SOURCE_PROCESS: AtomicU32 = AtomicU32::new(0);
/// The host whose client area the thumbnail is drawn in, as a number read without the lock by the
/// host's own erase (see `the_host_shows_a_thumbnail`); settled with the registration below.
static REGISTERED_HOST: AtomicIsize = AtomicIsize::new(0);

/// The registration, and the box it was last told the picture goes in.
static THUMBNAIL: Mutex<Thumbnail> = Mutex::new(Thumbnail::NOTHING);

#[derive(Clone, Copy)]
struct Thumbnail {
    /// The host the Shell is drawing into, and the registration it answered with.
    host: isize,
    registration: isize,
    /// The size of the host's client area the registration was last given, as a width and a height.
    destination: (i32, i32),
}

impl Thumbnail {
    const NOTHING: Self = Self {
        host: 0,
        registration: 0,
        destination: (0, 0),
    };
}

/// Take a player's own window as the film the pin's host shows: out of the Shell's lists, at the
/// bottom of the desktop's order, and registered in the host it stands behind.
///
/// **The style is put on once, at adoption.** `WS_EX_TOOLWINDOW` is the window's leaving the taskbar
/// and Alt+Tab, `WS_EX_NOACTIVATE` is it never taking the keyboard from the file being renamed, and
/// `WS_EX_TRANSPARENT` is a press falling through it should it ever be in front of anything. The
/// app-window and topmost bits mpv opens with are cleared: this window is behind the pin and must
/// never be held over it (see `video_launch::mpv_args`, which no longer asks for `--ontop` here).
///
/// **The z-order insert is one post rather than a wait**, `SWP_ASYNCWINDOWPOS` to the bottom: the
/// window belongs to mpv's process, and a synchronous reorder is this thread waiting on mpv's — the
/// road that deadlocked this app once already (see `fit_the_hosted_mpv_to_the_host`). What a post
/// costs is the window standing on its own for the moment it takes mpv to pump it, which is behind
/// the hover the pin is being taken up from.
///
/// Safety: `source` is a window the caller found by `pid` and this app did not make.
pub(super) unsafe fn adopt_the_source(source: HWND, pid: u32) {
    let ex = unsafe { GetWindowLongPtrW(source, GWL_EXSTYLE) };
    let ex = (ex & !(WS_EX_APPWINDOW.0 as isize) & !(WS_EX_TOPMOST.0 as isize))
        | WS_EX_TOOLWINDOW.0 as isize
        | WS_EX_NOACTIVATE.0 as isize
        | WS_EX_TRANSPARENT.0 as isize;
    let _ = unsafe { SetWindowLongPtrW(source, GWL_EXSTYLE, ex) };
    let _ = unsafe {
        SetWindowPos(
            source,
            HWND_BOTTOM,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS,
        )
    };

    // **A window that made itself the foreground gives it back.** mpv shows its own window when it has
    // opened the file, which is after the click that made the pin, and the style above stops it taking
    // the keyboard but not from having taken it already: the pin is asked for it here where the Shell
    // says the player's window is the one in front. A window of this app's own was the foreground a
    // moment ago, so the ask is one Windows grants.
    if unsafe { GetForegroundWindow() } == source {
        if let Some(pin) = pin_real_window::real_pin_window() {
            let _ = unsafe { SetForegroundWindow(pin) };
        }
    }

    note_the_source(source, pid);
    if let Some(host) = pin_film_host::film_host() {
        the_band_was_placed(host);
    }
}

/// Record the window the Shell is to draw: another player, or the same player's window made again,
/// is another picture and goes down with the registration the last one had.
pub(super) fn note_the_source(source: HWND, pid: u32) {
    let number = source.0 as isize;
    if SOURCE.load(Ordering::SeqCst) == number && SOURCE_PROCESS.load(Ordering::SeqCst) == pid {
        return;
    }

    if let Ok(mut held) = THUMBNAIL.lock() {
        if held.registration != 0 {
            let _ = unsafe { DwmUnregisterThumbnail(held.registration) };
        }
        *held = Thumbnail::NOTHING;
    }
    REGISTERED_HOST.store(0, Ordering::SeqCst);
    SOURCE.store(number, Ordering::SeqCst);
    SOURCE_PROCESS.store(pid, Ordering::SeqCst);
}

/// Whether this player's window is up as the pin's picture: still a window, and still one of the
/// player this app records rather than one given the number of a window it replaced.
pub(super) fn the_source_is_up(pid: u32) -> bool {
    let number = SOURCE.load(Ordering::SeqCst);
    if number == 0 || pid == 0 || SOURCE_PROCESS.load(Ordering::SeqCst) != pid {
        return false;
    }
    let source = HWND(number as *mut _);
    // Safety: the handle is one this module adopted from a window of mpv's, and the two questions are
    // about the window rather than changes to it.
    unsafe {
        if !IsWindow(source).as_bool() {
            return false;
        }
        let mut owner = 0u32;
        GetWindowThreadProcessId(source, Some(&mut owner));
        owner == pid
    }
}

/// Whether `host` is the window the Shell is drawing the picture in, which is the one fact the host's
/// own erase is answered from: a host wearing a thumbnail has nothing of this app's to fill.
pub(super) fn the_host_shows_a_thumbnail(host: HWND) -> bool {
    REGISTERED_HOST.load(Ordering::SeqCst) == host.0 as isize
}

/// Follow the host: register the picture where there is none, point it at the host's whole client
/// area, and move the source so its corner rides with the host's.
///
/// **The source's size is not asked for here**, which is the whole of what makes a resize step cheap:
/// the destination is the host's box and the Shell stretches whatever the source's box is into it.
/// The size is asked for once, when the hand lets go (see `settle_the_source_at_the_band`).
pub(super) fn the_band_was_placed(host: HWND) {
    if SOURCE.load(Ordering::SeqCst) == 0 {
        return;
    }
    let Some(origin) = (unsafe { client_origin(host) }) else {
        return;
    };
    let Some(destination) = (unsafe { client_size(host) }) else {
        return;
    };

    let source = HWND(SOURCE.load(Ordering::SeqCst) as *mut _);
    move_the_source(source, origin);

    let Ok(mut held) = THUMBNAIL.lock() else {
        return;
    };
    if held.host == host.0 as isize
        && held.registration != 0
        && held.destination == destination
    {
        return;
    }
    settle_the_registration(&mut held, host, destination);
}

/// Ask mpv's window for the box the host is at now: the one cross-process resize a drag makes, made
/// when the hand has let go of the edge rather than on every step of it.
pub(super) fn settle_the_source_at_the_band(host: HWND) {
    if SOURCE.load(Ordering::SeqCst) == 0 {
        return;
    }
    let Some(origin) = (unsafe { client_origin(host) }) else {
        return;
    };
    let Some((width, height)) = (unsafe { client_size(host) }) else {
        return;
    };
    let source = HWND(SOURCE.load(Ordering::SeqCst) as *mut _);
    if !unsafe { IsWindow(source) }.as_bool() {
        return;
    }

    let mut current = RECT::default();
    if unsafe { GetWindowRect(source, &mut current) }.is_err() {
        return;
    }
    if (current.left, current.top, current.right - current.left, current.bottom - current.top)
        == (origin.0, origin.1, width, height)
    {
        return;
    }
    let _ = unsafe {
        SetWindowPos(
            source,
            HWND::default(),
            origin.0,
            origin.1,
            width.max(1),
            height.max(1),
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS,
        )
    };
}

/// Take the picture off the screen with the host it is seen in, or put it back.
///
/// **A hidden source is a thumbnail the Shell stops drawing**, so this is only ever asked where the
/// host goes down with it — a pin put away, a park over the band — and never while the host is up.
/// The call is posted rather than made, for the reason every call on this window is (see
/// `adopt_the_source`).
pub(super) fn set_the_source_visible(visible: bool) {
    let number = SOURCE.load(Ordering::SeqCst);
    if number == 0 {
        return;
    }
    let source = HWND(number as *mut _);
    // Safety: the handle is a window this module adopted, and the question is about it rather than a
    // change to it. A window already where it is wanted is left alone: a host's show runs on every
    // tick of a pin that is drawing, and a `ShowWindow` a tick is a message mpv's thread owes nothing.
    if unsafe { IsWindowVisible(source) }.as_bool() != visible {
        let _ = unsafe {
            ShowWindowAsync(source, if visible { SW_SHOWNOACTIVATE } else { SW_HIDE })
        };
    }
}

/// Take the registration down: the host it was drawn in is going, and a registration left behind a
/// window that is gone is a leak of the Shell's own resources.
pub(super) fn forget_the_thumbnail() {
    if let Ok(mut held) = THUMBNAIL.lock() {
        if held.registration != 0 {
            let _ = unsafe { DwmUnregisterThumbnail(held.registration) };
        }
        *held = Thumbnail::NOTHING;
    }
    REGISTERED_HOST.store(0, Ordering::SeqCst);
    SOURCE.store(0, Ordering::SeqCst);
    SOURCE_PROCESS.store(0, Ordering::SeqCst);
}

/// Take the registration down where it is drawn in `host`: a host of the hover's that goes while it
/// is showing a film of its own, which is every road this prototype does not take.
pub(super) fn forget_the_thumbnail_of(host: HWND) {
    let drawn_in = REGISTERED_HOST.load(Ordering::SeqCst) == host.0 as isize;
    if drawn_in {
        forget_the_thumbnail();
    }
}

/// Register the picture in the host's client area, or point the registration the host already wears at
/// the box it is in now.
fn settle_the_registration(held: &mut Thumbnail, host: HWND, destination: (i32, i32)) {
    if held.registration != 0 && held.host != host.0 as isize {
        let _ = unsafe { DwmUnregisterThumbnail(held.registration) };
        held.registration = 0;
    }
    held.host = host.0 as isize;
    REGISTERED_HOST.store(held.host, Ordering::SeqCst);

    if held.registration == 0 {
        let source = HWND(SOURCE.load(Ordering::SeqCst) as *mut _);
        let Ok(registration) = (unsafe { DwmRegisterThumbnail(host, source) }) else {
            // Nothing to draw in until a later placement asks again: a source that has made its
            // window and not yet presented is a registration the Shell can refuse.
            REGISTERED_HOST.store(0, Ordering::SeqCst);
            return;
        };
        held.registration = registration;
    }

    let properties = DWM_THUMBNAIL_PROPERTIES {
        dwFlags: DWM_TNP_RECTDESTINATION | DWM_TNP_SOURCECLIENTAREAONLY | DWM_TNP_VISIBLE,
        // The whole of the host's client area: what the Shell draws is the source stretched into it.
        rcDestination: RECT {
            left: 0,
            top: 0,
            right: destination.0.max(1),
            bottom: destination.1.max(1),
        },
        rcSource: RECT::default(),
        opacity: 255,
        fVisible: true.into(),
        // The picture is the player's client area and not the frame mpv puts around it.
        fSourceClientAreaOnly: true.into(),
    };
    if unsafe { DwmUpdateThumbnailProperties(held.registration, &properties) }.is_err() {
        let _ = unsafe { DwmUnregisterThumbnail(held.registration) };
        *held = Thumbnail::NOTHING;
        REGISTERED_HOST.store(0, Ordering::SeqCst);
        return;
    }
    held.destination = destination;
}

/// Move the source so its own corner is the host's, keeping the size it has: a step of a move drag
/// and a step of a resize alike, and the one thing that keeps a pure move from leaving the film
/// standing where the pin was.
fn move_the_source(source: HWND, origin: (i32, i32)) {
    let mut current = RECT::default();
    // Safety: the handle is a window this module adopted, and both questions are about the window.
    if unsafe { GetWindowRect(source, &mut current) }.is_err() || !unsafe { IsWindow(source) }.as_bool()
    {
        return;
    }
    if (current.left, current.top) == origin {
        return;
    }
    let _ = unsafe {
        SetWindowPos(
            source,
            HWND::default(),
            origin.0,
            origin.1,
            0,
            0,
            SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS,
        )
    };
}

/// The corner of a window's client area, on the screen.
unsafe fn client_origin(hwnd: HWND) -> Option<(i32, i32)> {
    let mut origin = POINT::default();
    (unsafe { ClientToScreen(hwnd, &mut origin) }.as_bool()).then_some((origin.x, origin.y))
}

/// The size of a window's client area, or nothing for a window that has gone.
unsafe fn client_size(hwnd: HWND) -> Option<(i32, i32)> {
    let mut client = RECT::default();
    (unsafe { GetClientRect(hwnd, &mut client) }.is_ok())
        .then_some((client.right - client.left, client.bottom - client.top))
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A registration is pointed at the host's box again only when the box or the host has changed,
    /// which is the whole cost a resize step is meant to be: a step that moved nothing asks nothing.
    #[test]
    fn a_thumbnail_is_settled_again_only_when_its_host_or_the_box_it_is_drawn_in_changes() {
        let holding = Thumbnail {
            host: 0x2000,
            registration: 9,
            destination: (320, 240),
        };

        assert!(holding.host == 0x2000 && holding.destination == (320, 240));
        assert!(
            !(holding.host == 0x2001 && holding.destination == (320, 240)),
            "another host is another registration"
        );
        assert!(
            !(holding.host == 0x2000 && holding.destination == (321, 240)),
            "and the same host in another box is another destination"
        );
    }

    /// The registration is nobody's where none was taken, and a host wearing one is the only host the
    /// Shell is drawing into — which is the answer the host's own erase is given.
    #[test]
    fn a_host_shows_a_thumbnail_only_while_a_registration_stands_in_it() {
        assert!(!the_host_shows_a_thumbnail(HWND(0x2000 as *mut _)));
        assert_eq!(Thumbnail::NOTHING.registration, 0);
        assert_eq!(Thumbnail::NOTHING.host, 0);
    }
}

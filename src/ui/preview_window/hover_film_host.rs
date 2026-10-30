//! The window a hovered film is drawn in: a top-level host of this app's own, so that the
//! hover's picture can be moved into a pin's window later without beginning a second player.
//!
//! **Why it is not the player's own window.** A window of the desktop's cannot be turned into the
//! pin's chrome window without losing it, and the hover wants a destination the Shell can draw a
//! thumbnail into anyway. Both engines make a window of their own — mpv is launched with no `--wid`
//! and FFmpeg's player has no embed switch — and this host is the Shell's destination for a
//! thumbnail of it rather than its parent (see `pin_film_thumbnail`).
//!
//! **Why it is top-level and hidden.** Until the Shell has composed the picture in it, the host
//! draws nothing but black — so it is made off the screen, shown only once the player's source is
//! adopted, and placed at the hover's own box.
//!
//! **Why it is disabled.** A press on the picture is the hover's press, and a disabled window is
//! skipped by `WindowFromPoint`. (The pin's host keeps its presses with a `WM_NCHITTEST` answer
//! instead, because a disabled window cannot be activated or put down and the pair's taskbar button
//! is worn by it — see `pin_film_host::answer_a_host`.) It also never activates: a hover must never
//! take the focus from a file being renamed.
//!
//! The host is still the thing the hover is, and the thing a pin takes up: it is promoted into the
//! pin's host, so the registration the source is drawn with is already in the window that becomes
//! the pin's and a take-up asks neither the player nor the source for anything (see
//! `pin_film_host::promote_the_hover_film_host`).
//!
//! The host belongs to the hover, not to the player: a hover that replaces this one takes the same
//! host when the old player is still in it, and every road that ends the hover destroys it (see
//! `kill_player_process`, `hide_preview`, and the loop's `Hide`). A host a pin took up belongs to
//! the pin from then on (see `pin_film_host`).

use std::sync::atomic::{AtomicIsize, AtomicU32, Ordering};
use std::sync::Once;

use windows::core::w;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{GetStockObject, HBRUSH, NULL_BRUSH};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::Input::KeyboardAndMouse::EnableWindow;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, RegisterClassExW, ShowWindow, ShowWindowAsync,
    SW_HIDE, SW_SHOWNOACTIVATE, WNDCLASSEXW, WS_CLIPCHILDREN, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW,
    WS_EX_TOPMOST, WS_EX_TRANSPARENT, WS_POPUP,
};

use super::*;

/// The class a hover's film host is made from. A class of its own rather than the pin host's: the
/// two hosts have different parents, different styles and different owners, and only the pin's
/// takes a `WS_CHILD` box.
const HOVER_FILM_HOST_CLASS: windows::core::PCWSTR = w!("RustHoverPreviewHoverFilmHost");

unsafe extern "system" fn the_default_answer(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    // **This window can be a pin's host, so it answers everything a pin's host does.** A pin taken up
    // on a hover's running film keeps this window rather than making another one (see `pin_film_host`,
    // `promote_the_hover_film_host`), so the band behind the picture, the cover a relaunch left
    // standing, the Shell's close, the window coming back out of the taskbar and the still its button is
    // drawn of are all one answer shared with the pin's own class (see `answer_a_host`).
    if let Some(answer) = unsafe { pin_film_host::answer_a_host(hwnd, msg, wparam, lparam) } {
        return answer;
    }

    unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
}

/// The host a hovered film is drawn in, and the player this app has begun into it.
static HOVER_FILM_HOST: AtomicIsize = AtomicIsize::new(0);
static HOVER_FILM_HOST_PLAYER: AtomicU32 = AtomicU32::new(0);
/// The file the player in the host was begun for, so that a pin taking the host up can prove the
/// player is the one drawing the film it was given rather than another hover's.
static HOVER_FILM_HOST_PATH: Lazy<Mutex<Option<PathBuf>>> = Lazy::new(|| Mutex::new(None));
/// The thread that made the host, which is the only thread Windows lets destroy it.
static HOVER_FILM_HOST_THREAD: AtomicU32 = AtomicU32::new(0);
static HOVER_FILM_HOST_CLASS_REGISTERED: Once = Once::new();

/// The host the hovered film is drawn in, or nothing while there is none.
pub(super) fn hover_film_host() -> Option<HWND> {
    let host = HOVER_FILM_HOST.load(Ordering::SeqCst);
    (host != 0).then_some(HWND(host as *mut _))
}

/// The player this app has begun into the hover's host.
pub(super) fn hover_film_host_player() -> Option<u32> {
    let player = HOVER_FILM_HOST_PLAYER.load(Ordering::SeqCst);
    (player != 0).then_some(player)
}

/// The file the player this host holds was begun for: the path beside the hover's player, or nothing
/// where the host holds another player (see `the_film_the_player_plays`).
pub(super) fn the_hover_players_file(pid: u32) -> Option<PathBuf> {
    if hover_film_host_player() != Some(pid) {
        return None;
    }
    HOVER_FILM_HOST_PATH
        .lock()
        .ok()
        .and_then(|held| held.clone())
}

/// Record the player begun into the hover's host, before it has a window of its own in there.
pub(super) fn note_hover_film_host_player(pid: u32, path: &Path) {
    HOVER_FILM_HOST_PLAYER.store(pid, Ordering::SeqCst);
    if let Ok(mut held) = HOVER_FILM_HOST_PATH.lock() {
        *held = Some(path.to_path_buf());
    }
}

/// The hover's host, where this player is the one begun into it for this file and is already
/// drawing there: the host a pin can move into its own window instead of beginning a second player.
pub(super) fn the_hover_film_to_adopt(path: &Path) -> Option<HWND> {
    if current_media_type() != Some(MediaType::Video)
        || !matches!(video_route(path), VideoRoute::Mpv | VideoRoute::Ffplay)
    {
        return None;
    }

    // **The registration follows the host rather than a picture being begun again**: the host is the
    // destination of the Shell's thumbnail of the player's own window, whichever engine made it (see
    // `pin_film_thumbnail` and `pin_film_host::promote_the_hover_film_host`). A hover whose source is
    // not adopted yet is not taken up: there is no picture to move.
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    if !the_hover_player_was_begun_in_the_host(pid) || !the_hover_player_is_drawing(pid) {
        return None;
    }

    let matches = HOVER_FILM_HOST_PATH
        .lock()
        .ok()
        .and_then(|held| held.as_deref().map(|held| held == path))
        .unwrap_or(false);
    matches.then(hover_film_host).flatten()
}

/// Forget the hover's host after handing its window to the pin: the window is the pin's from here,
/// so the hover's own take-down must not destroy it (see `pin_film_host`).
pub(super) fn forget_the_hover_film_host_that_was_handed_over(host: HWND) {
    if hover_film_host() != Some(host) {
        return;
    }
    HOVER_FILM_HOST.store(0, Ordering::SeqCst);
    HOVER_FILM_HOST_PLAYER.store(0, Ordering::SeqCst);
    HOVER_FILM_HOST_THREAD.store(0, Ordering::SeqCst);
    if let Ok(mut held) = HOVER_FILM_HOST_PATH.lock() {
        *held = None;
    }
}

/// Whether this file is a hovered film that a host should hold: mpv's, and FFmpeg's player's while
/// Pin Mode is on. A film pinned from a hover is drawn in a host by then, so a pin's take-up can take
/// the running picture up without beginning a second player; every other route — the engine's own
/// frame, and a pin standing in the preview's window — keeps the window it has.
///
/// It is asked on a video launch only, which is what keeps a sound out: the audio road begins its
/// player through `start_audio_player` and never asks here (see `the_hover_film_to_adopt` for the
/// media-type gate every take-up asks, where the media on screen is the one that decides).
pub(super) fn the_hover_film_is_kept_in_a_host(path: &Path) -> bool {
    !pinned() && pin_enabled() && matches!(video_route(path), VideoRoute::Mpv | VideoRoute::Ffplay)
}

/// Whether the player this app is holding was begun into the hover's host.
pub(super) fn the_hover_player_was_begun_in_the_host(pid: u32) -> bool {
    pid != 0 && hover_film_host().is_some() && hover_film_host_player() == Some(pid)
}

/// Whether this player has put a picture in the hover's host — the fact a hover's start waits on,
/// and the one that says there is a picture to show rather than a black rectangle.
///
/// **The picture is the source the Shell draws into the host**: the player's own window, adopted and
/// still live (see `pin_film_thumbnail::the_source_is_up`), whichever engine made it.
pub(super) fn the_hover_player_is_drawing(pid: u32) -> bool {
    hover_film_host().is_some()
        && the_hover_player_was_begun_in_the_host(pid)
        && pin_film_thumbnail::the_source_is_up(pid)
}

/// Make the hover's host where there is none, put it at `rect`, and leave it off the screen until
/// the player has drawn into it: the window FFmpeg's child is adopted into, and the destination of
/// mpv's thumbnail.
pub(super) unsafe fn take_up_the_hover_film_host(rect: ScreenRegion) -> Option<HWND> {
    // **A handle to nothing is not answered as the host**: a window belongs to the thread that made
    // it, and one whose thread has gone is destroyed with it (see `take_down_the_hover_film_host`,
    // which forgets the same handle for the same reason).
    let host = match hover_film_host().filter(|host| unsafe { IsWindow(*host) }.as_bool()) {
        Some(host) => host,
        None => unsafe { create()? },
    };
    unsafe { place_the_hover_film_host(rect) };
    Some(host)
}

/// Put the host at the hover's box, without showing it: a move, a relayout or a hover re-shown at
/// another box, with the Shell's picture cropped to it.
pub(super) unsafe fn place_the_hover_film_host(rect: ScreenRegion) {
    let Some(host) = hover_film_host() else {
        return;
    };

    let width = (rect.2 - rect.0).max(1);
    let height = (rect.3 - rect.1).max(1);
    let _ = SetWindowPos(
        host,
        HWND::default(),
        rect.0,
        rect.1,
        width,
        height,
        SWP_NOZORDER | SWP_NOACTIVATE,
    );
    // **A picture the Shell is drawing into this host follows the box it is drawn in**, which is the
    // thumbnail's crop re-asked of the film's own source rectangle and nothing of the player's (see
    // `pin_film_thumbnail::the_band_was_placed`). A host with no such picture has nothing to re-point,
    // and nothing of the player's is ever asked for a box on this road.
    pin_film_thumbnail::the_band_was_placed(host);
}

/// Show or hide the hover's host: what a player's first frame is revealed by, and what a hover
/// that is going away (or a player that has not drawn) leaves off the screen.
pub(super) unsafe fn set_the_hover_film_host_visible(visible: bool) {
    if let Some(host) = hover_film_host() {
        let _ = ShowWindow(host, if visible { SW_SHOWNOACTIVATE } else { SW_HIDE });
    }
}

/// The hover's box as the region its host is placed at: the loop keeps a video's place as
/// `(x, y, width, height)` and every window placed from it takes `(left, top, right, bottom)`.
pub(super) fn the_hover_box_as_a_region(box_: (i32, i32, i32, i32)) -> ScreenRegion {
    let (x, y, width, height) = box_;
    (x, y, x + width, y + height)
}

/// Put the host at `rect` and show it, where this player is the one begun into it and has a child
/// of its own drawing: the hover's first real frame, on the tick the start wait ends.
pub(super) unsafe fn reveal_the_hover_film_host(pid: u32, rect: ScreenRegion) {
    if !the_hover_player_is_drawing(pid) {
        return;
    }
    unsafe { place_the_hover_film_host(rect) };
    unsafe { set_the_hover_film_host_visible(true) };
}

/// Destroy the hover's host and forget the player it held. The player's own window outlives the host,
/// so it is also asked to end where the caller has one (see `kill_player_process`).
pub(super) unsafe fn take_down_the_hover_film_host() {
    // A window belongs to the thread that made it. A hover ended from the Explorer hook's thread
    // leaves the host to the loop's own `Hide`, which runs on the host's thread (see
    // `hide_preview`); clearing the handle here would lose the only way to find it again.
    if HOVER_FILM_HOST_THREAD.load(Ordering::SeqCst) != unsafe { GetCurrentThreadId() } {
        // It can still be taken off the screen from here: the hide is posted to the host's own
        // thread, so the film does not linger between this road and the loop's `Hide`.
        if let Some(host) = hover_film_host() {
            if unsafe { IsWindow(host) }.as_bool() {
                let _ = ShowWindowAsync(host, SW_HIDE);
            } else {
                // **The thread that made it has gone, and the window with it**: a handle to nothing
                // is forgotten rather than answered as the host by every road that asks.
                forget_the_hover_film_host_that_was_handed_over(host);
            }
        }
        return;
    }
    if let Some(host) = hover_film_host() {
        let _ = DestroyWindow(host);
        // A registration the Shell is drawing into this window goes with it, where this road ever
        // had one: a host of a hover's is a thumbnail destination on no road this prototype takes,
        // and a registration left behind a destroyed window is the Shell's resource leaking (see
        // `pin_film_thumbnail`).
        pin_film_thumbnail::forget_the_thumbnail_of(host);
    }
    HOVER_FILM_HOST.store(0, Ordering::SeqCst);
    HOVER_FILM_HOST_PLAYER.store(0, Ordering::SeqCst);
    HOVER_FILM_HOST_THREAD.store(0, Ordering::SeqCst);
    if let Ok(mut held) = HOVER_FILM_HOST_PATH.lock() {
        *held = None;
    }
}

/// Destroy the host only where this player is the one it holds: the owner-aware teardown, so a
/// pin's teardown never takes a hover's host with it and a hover's end never takes the pin's.
pub(super) unsafe fn take_down_the_hover_film_host_if_it_held(pid: u32) {
    if the_hover_player_was_begun_in_the_host(pid) {
        unsafe { take_down_the_hover_film_host() };
    }
}

unsafe fn create() -> Option<HWND> {
    HOVER_FILM_HOST_CLASS_REGISTERED.call_once(|| {
        let hinstance = GetModuleHandleW(None).expect("the module's own handle");
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(the_default_answer),
            hInstance: hinstance.into(),
            // **No background brush, because the window procedure paints its own erase** — black
            // for a hover's band, and the picture a drag holds where a pin took this window up (see
            // `the_default_answer`).
            hbrBackground: HBRUSH(GetStockObject(NULL_BRUSH).0),
            lpszClassName: HOVER_FILM_HOST_CLASS,
            ..Default::default()
        };
        RegisterClassExW(&class);
    });

    let hinstance = GetModuleHandleW(None).ok()?;
    // A black window behind the film and nothing else: no border, no taskbar entry, no activation,
    // no topmost band of its own to give back, and no press of its own — the picture is the Shell's
    // thumbnail of the player's window and a press on it is the hover's press (see the module doc).
    // Made off every display and without `WS_VISIBLE`: it is shown by the tick that finds the
    // player's source adopted.
    //
    // `WS_CLIPCHILDREN` for the same reason the pin's own window wears it: the window procedure
    // fills the band behind the film on every erase, and an erase that is not clipped to the area
    // outside the player's window is that fill drawn over the film at every show and every resize of
    // the band (see `pin_real_window::real_pin_style`).
    let host = CreateWindowExW(
        WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_TOPMOST | WS_EX_TRANSPARENT,
        HOVER_FILM_HOST_CLASS,
        w!("hover film host"),
        WS_POPUP | WS_CLIPCHILDREN,
        -32000,
        -32000,
        1,
        1,
        None,
        None,
        hinstance,
        None,
    )
    .ok()?;

    let _ = EnableWindow(host, false);
    // The still its button's hover is drawn of is this app's own picture, taken when a pin puts it away
    // into the taskbar: this window is composed of a thumbnail the Shell draws *over* it, so its own
    // rows are no picture of the film (see `pin_taskbar::supply_the_stills_of_this_window`).
    unsafe { pin_taskbar::supply_the_stills_of_this_window(host) };
    HOVER_FILM_HOST.store(host.0 as isize, Ordering::SeqCst);
    HOVER_FILM_HOST_PLAYER.store(0, Ordering::SeqCst);
    HOVER_FILM_HOST_THREAD.store(unsafe { GetCurrentThreadId() }, Ordering::SeqCst);
    Some(host)
}

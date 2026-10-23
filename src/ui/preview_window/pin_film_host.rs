//! The window a pinned mpv film is drawn in: a child of the pin's own window, standing in the
//! band its picture is shown through.
//!
//! **Why it is not the pin's window itself.** mpv's `--wid` fills the window it is given, and the
//! pin's window is the chrome and the transport as well as the band; pointed at the pin, the film
//! would be drawn over the caption. So the picture gets a window of its own at the band's client
//! rectangle, and the pin keeps every other pixel (see `band_in_the_window`, which is the box the
//! host is given).
//!
//! **Why it is a child.** The film has to travel with the pin — on a move, on a resize, and with
//! the window when it is hidden or put away — and a child does that in the compositor rather than
//! by this app moving two windows and then repairing their order. That repair is what the topmost
//! band was for, and it is removed for the host (see `ensure_video_window_topmost`).
//!
//! **Why it is disabled.** A press on the picture is a press on the pin: the pin is dragged by its
//! film, and the band's alpha-1 claim is already hit-testable (see `fill_band_hit_testable`). A
//! disabled child is skipped by `WindowFromPoint`, so the press arrives at the pin rather than at
//! mpv's window — the same answer the prototype measured, and no style of mpv's to rely on.
//!
//! The host outlives the players begun into it: a seek is mpv's own IPC, a resize is the host's
//! box, and a file step begins another player into the same host. It goes down with the pin (see
//! `take_down_the_pin_film_host`, called from the state a road out of a pin clears).

use std::sync::atomic::{AtomicIsize, AtomicU32, Ordering};
use std::sync::Once;

use windows::core::w;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{GetStockObject, BLACK_BRUSH, HBRUSH};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::EnableWindow;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, GetWindow, GetWindowThreadProcessId,
    RegisterClassExW, ShowWindow, GW_CHILD, GW_HWNDNEXT, SW_HIDE, SW_SHOWNOACTIVATE,
    WINDOW_EX_STYLE, WNDCLASSEXW, WS_CHILD,
};

use super::*;

/// The class a film's host is made from. A class of its own rather than the pin window's: what a
/// `WS_CHILD` host answers to is a window that draws nothing and takes nothing, and the pin's
/// window procedure is the pin's (see `pin_real_window`).
const FILM_HOST_CLASS: windows::core::PCWSTR = w!("RustHoverPreviewPinFilmHost");

/// The default answer is the whole of what the host does: it draws nothing of its own and no press
/// lands on it, because the picture is mpv's child (see the module doc). This bindings version
/// exposes `DefWindowProcW` as a Rust-callable function, so the class still needs a system-ABI shim
/// around it.
unsafe extern "system" fn the_default_answer(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
}

/// The host of the pin that is up, and the player this app has attached to it.
static PIN_FILM_HOST: AtomicIsize = AtomicIsize::new(0);
static PIN_FILM_HOST_PLAYER: AtomicU32 = AtomicU32::new(0);
static FILM_HOST_CLASS_REGISTERED: Once = Once::new();

/// The host the pin is showing its film through, or nothing while there is none.
pub(super) fn film_host() -> Option<HWND> {
    let host = PIN_FILM_HOST.load(Ordering::SeqCst);
    (host != 0).then_some(HWND(host as *mut _))
}

/// The player this app has attached to the host, recorded beside it so that a window of mpv's that
/// is not the pin's — a hover's, which is a window of the desktop — is never mistaken for it.
pub(super) fn hosted_player_pid() -> Option<u32> {
    let player = PIN_FILM_HOST_PLAYER.load(Ordering::SeqCst);
    (player != 0).then_some(player)
}

/// Record the player begun into the host: called where the player is spawned, before its window
/// exists, so that the wait for its first frame is a wait for *its* child and not for the last
/// film's.
pub(super) fn note_hosted_player(pid: u32) {
    PIN_FILM_HOST_PLAYER.store(pid, Ordering::SeqCst);
}

/// The host to give a player begun for `path`, where that player is a pinned mpv and the pin
/// stands in a window of its own. `None` is every other launch: a hover's film, FFmpeg's player,
/// and the media engine's own frame.
pub(super) fn host_for_the_pins_film(path: &Path) -> Option<HWND> {
    the_film_is_shown_in_a_host(path).then(film_host).flatten()
}

/// Whether this file is shown in a host this app owns: a film mpv plays, pinned in a window of the
/// pin's own. Asked of the route rather than of the player's process because it is the *launch*
/// that decides where the picture is drawn (see `start_mpv_video_playback`).
fn the_film_is_shown_in_a_host(path: &Path) -> bool {
    pin_is_its_own_window() && video_route(path) == VideoRoute::Mpv
}

/// Whether the player this app is holding is the one it began into the host — the question every
/// road that used to find a top-level mpv window by pid asks instead (see
/// `ensure_video_window_topmost`, `hold_video_window_frame`, `settle_pinned_mpv_ontop`).
///
/// It says the pairing and not that the film is drawing: a player begun into the host has not put
/// a child of its own there for a moment or more, and that wait is `the_hosted_player_is_drawing`.
pub(super) fn the_pinned_player_was_begun_in_the_host() -> bool {
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    pid != 0 && film_host().is_some() && hosted_player_pid() == Some(pid)
}

/// Whether this player has put a child of its own under the host, which is what a film's window
/// arriving *is* for a player that has no window of the desktop's.
///
/// **It is asked of the process, not of the window alone**, because the host carries the last
/// player's child for as long as that process is alive: a file step begins a new player into the
/// same host, and the older child would answer "a film is drawing" for a film about to be killed.
/// A child's owning process is the player that made it — mpv's window is created by mpv's own
/// process under `--wid` — so the two facts are one question (see `GetWindowThreadProcessId`).
pub(super) fn the_hosted_player_is_drawing(pid: u32) -> bool {
    pid != 0 && hosted_player_pid() == Some(pid) && host_has_a_child_of(pid)
}

/// Whether this player has a window on screen: a window of the desktop's of its own, or the child
/// a hosted player draws under the pin's own host. It is one answer because the settlement roads
/// have one question — has this player arrived — and a hosted player never publishes a `VIDEO_HWND`
/// to be found by (see `video_window_for`, `settle_video_retirement` and `PinSwapHold::settle`).
pub(super) fn the_players_window_is_up(pid: u32) -> bool {
    video_window_for(pid).is_some() || the_hosted_player_is_drawing(pid)
}

fn host_has_a_child_of(pid: u32) -> bool {
    let Some(host) = film_host() else {
        return false;
    };

    unsafe {
        let mut child = GetWindow(host, GW_CHILD).unwrap_or_default();
        while !child.is_invalid() {
            let mut owner = 0u32;
            GetWindowThreadProcessId(child, Some(&mut owner));
            if owner == pid {
                return true;
            }
            child = GetWindow(child, GW_HWNDNEXT).unwrap_or_default();
        }
    }

    false
}

/// The band's rectangle in the pin window's own client coordinates: the whole window for a film,
/// whose chrome is drawn over it, and the content box below the caption where it is not.
fn band_in_the_window(content: ScreenRegion) -> (i32, i32, i32, i32) {
    let (overlay, caption) = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| (pin.overlay, pin.caption)))
        .unwrap_or((true, 0));

    let width = (content.2 - content.0).max(1);
    let height = (content.3 - content.1).max(1);

    (0, if overlay { 0 } else { caption }, width, height)
}

/// Make the host where the pin has none, or put the one it has at the band it stands at now.
///
/// Asked of the pin rather than of the caller's box, because the two roads that reach it — a show
/// and a box change — have the pin's state written before either of them asks, and the band the
/// host is placed at is the same band the window was painted around.
///
/// **A host made here is not shown unless the player is already drawing in it.** It is a black
/// window until the player this app begins has put a child of its own in it, and a black rectangle
/// over the band is the film the hover was showing, blanked. A host a restore takes up is drawing
/// already and is shown here; one a fresh launch made is shown by the tick's own road once that
/// child is there (see `ensure_video_window_topmost`), and the park's settle that shows it again
/// after a park.
pub(super) unsafe fn take_up_the_pin_film_host() {
    let Some(path) = pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.path.clone())) else {
        take_down_the_pin_film_host();
        return;
    };

    if !the_film_is_shown_in_a_host(&path) {
        take_down_the_pin_film_host();
        return;
    }

    let Some(content) = pinned_content() else {
        return;
    };

    let host = match film_host() {
        Some(host) => host,
        None => {
            let Some(parent) = pin_real_window::real_pin_window() else {
                return;
            };
            let Some(host) = create(parent) else {
                return;
            };
            host
        }
    };

    let (x, y, width, height) = band_in_the_window(content);
    let _ = SetWindowPos(
        host,
        HWND::default(),
        x,
        y,
        width,
        height,
        SWP_NOZORDER | SWP_NOACTIVATE,
    );
    pin_chrome_strips::place_the_chrome_strips();

    // **A hosted player has no window of the desktop's to be found by pid**, so any handle still
    // published for an older player is not this film's picture and must not be read as one (see
    // `the_players_window_is_up`). A player the pin took up from a hover is not this case yet: it
    // is a top-level window until the relaunch below it takes the film into the host, and the
    // park that relaunch raises goes looking for that window (see
    // `begin_the_adopted_film_in_the_host`).
    if the_pinned_player_was_begun_in_the_host() {
        VIDEO_HWND.store(0, Ordering::SeqCst);
    }

    // Only a host the player is already drawing in is shown: a fresh one is a black rectangle over
    // the band, and what the band showed a moment ago is the film the hover was playing (see the
    // module note on `create`).
    if the_hosted_player_is_drawing(VIDEO_PID.load(Ordering::SeqCst)) {
        let _ = ShowWindow(host, SW_SHOWNOACTIVATE);
        pin_chrome_strips::sync_the_chrome_strips();
    }
}

/// Put the host at the pin's band, without showing it: the band-change road, where the window is
/// either already up or is the park's to hand back (see `ensure_pinned_sibling_box`).
pub(super) unsafe fn place_the_pin_film_host(content: ScreenRegion) {
    let Some(host) = film_host() else {
        return;
    };

    let (x, y, width, height) = band_in_the_window(content);
    let _ = SetWindowPos(
        host,
        HWND::default(),
        x,
        y,
        width,
        height,
        SWP_NOZORDER | SWP_NOACTIVATE,
    );
    pin_chrome_strips::place_the_chrome_strips();
}

/// Hide or show the host for a park, at the box it already has: a park hides the window behind the
/// band, and the pin paints its held frame in its place (see `park_pinned_player_inner`). The
/// chrome stays where the chrome state puts it: the strips are the pin's window's rather than the
/// host's, and a park hides the picture rather than the chrome (see `pin_chrome_strips`).
pub(super) unsafe fn set_the_pin_film_host_visible(visible: bool) {
    if let Some(host) = film_host() {
        let _ = ShowWindow(host, if visible { SW_SHOWNOACTIVATE } else { SW_HIDE });
    }
    // The strips carry the chrome on every road, whatever the host is doing: the sync answers
    // what the chrome state asks for over the picture or over the cover (see `pin_chrome_strips`).
    pin_chrome_strips::sync_the_chrome_strips();
}

/// Destroy the host and forget the player it held: the pin it was a child of is ending, and a
/// window left in the pin's window would stand in the next pin's band as a black rectangle. The
/// strips stay: they are the pin's window's, and the next take-up places them again (see
/// `pin_chrome_strips`).
pub(super) unsafe fn take_down_the_pin_film_host() {
    if let Some(host) = film_host() {
        let _ = DestroyWindow(host);
    }
    PIN_FILM_HOST.store(0, Ordering::SeqCst);
    PIN_FILM_HOST_PLAYER.store(0, Ordering::SeqCst);
}

unsafe fn create(parent: HWND) -> Option<HWND> {
    FILM_HOST_CLASS_REGISTERED.call_once(|| {
        let hinstance = GetModuleHandleW(None).expect("the module's own handle");
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(the_default_answer),
            hInstance: hinstance.into(),
            hbrBackground: HBRUSH(GetStockObject(BLACK_BRUSH).0),
            lpszClassName: FILM_HOST_CLASS,
            ..Default::default()
        };
        RegisterClassExW(&class);
    });

    let hinstance = GetModuleHandleW(None).ok()?;
    // **Made without `WS_VISIBLE`, and that is what keeps the band the film's until the player has
    // drawn into it.** A host put on the screen empty is a black rectangle where the hover's picture
    // was, for as long as the player takes to open the file — and for a pin that took up a hover's
    // player it is the picture on screen that the park's held frame is read from (see
    // `begin_the_adopted_film_in_the_host`). What shows it is the take-up of a host the player is
    // already drawing in, the tick's own road, and the settle of a park (see
    // `take_up_the_pin_film_host`, `ensure_video_window_topmost`).
    let host = CreateWindowExW(
        WINDOW_EX_STYLE::default(),
        FILM_HOST_CLASS,
        w!("film host"),
        WS_CHILD,
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

    // Disabled before it is ever shown: mpv's child is the picture, and the press that lands on
    // the picture has to reach the pin (see the module doc).
    let _ = EnableWindow(host, false);
    PIN_FILM_HOST.store(host.0 as isize, Ordering::SeqCst);
    // And the chrome is put back over it: making the host puts it on top of the strips, which are
    // the pin's window's and were standing before it, so the order is re-asserted here (see
    // `pin_chrome_strips`).
    unsafe { pin_chrome_strips::raise_the_chrome_strips_over_the_host() };
    Some(host)
}

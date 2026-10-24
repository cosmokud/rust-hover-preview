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

use std::sync::atomic::{AtomicBool, AtomicIsize, AtomicU32, Ordering};
use std::sync::Once;

use windows::core::w;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{GetStockObject, BLACK_BRUSH, HBRUSH};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::EnableWindow;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, GetWindow, GetWindowLongPtrW, GetWindowRect,
    GetWindowThreadProcessId, RegisterClassExW, SetParent, SetWindowLongPtrW, SetWindowPos,
    ShowWindow, GWL_EXSTYLE, GWL_STYLE, GW_CHILD, GW_HWNDNEXT, HWND_NOTOPMOST, HWND_TOP,
    SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER, SWP_SHOWWINDOW, SW_HIDE,
    SW_SHOWNOACTIVATE, WINDOW_EX_STYLE, WNDCLASSEXW, WS_CHILD, WS_EX_TOPMOST, WS_EX_TRANSPARENT,
    WS_POPUP, WS_VISIBLE,
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
/// The hosted player whose frame this app has proved on screen, so the pipe read that proves it is
/// asked once and not on every tick (see `the_hosted_frame_is_up`).
static PRESENTED_PID: AtomicU32 = AtomicU32::new(0);
/// The move cover: the held frame the pin paints in its own bitmap while a hover's host is hidden,
/// restyled and re-parented into the pin (see `raise_the_move_cover` and `show_pinned_window`).
static PIN_MOVE_COVER: AtomicBool = AtomicBool::new(false);
/// Whether the pin's film is FFmpeg's, written where the pin is taken up and where a film is begun
/// into the host. An atomic rather than a read of the pin's state because the chrome's own road asks
/// it with the pin's lock already held (see `the_hosted_frame_is_up`).
static PIN_FILM_IS_FFPLAY: AtomicBool = AtomicBool::new(false);
/// The second the hover's picture stood at when that cover was raised, so that the cover is only
/// given up once a frame has been drawn into the moved window (see `settle_the_move_cover`).
static PIN_MOVE_COVER_SECONDS: Lazy<Mutex<Option<f64>>> = Lazy::new(|| Mutex::new(None));
static FILM_HOST_CLASS_REGISTERED: Once = Once::new();

/// Raise the move cover: the pin's next paint stands on the held frame (or on black where none
/// could be taken) rather than on an alpha-1 band, so the desktop cannot show through while the
/// hover's host leaves the compositor as a top-level window and arrives as a child of the pin.
pub(super) fn raise_the_move_cover() {
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    let seconds =
        mpv_ipc::get_property(pid, "time-pos").and_then(|answer| answer.parse::<f64>().ok());
    if let Ok(mut held) = PIN_MOVE_COVER_SECONDS.lock() {
        *held = seconds;
    }
    PIN_MOVE_COVER.store(true, Ordering::SeqCst);
}

/// Whether the move cover is standing (see `raise_the_move_cover`).
pub(super) fn the_move_cover_is_up() -> bool {
    PIN_MOVE_COVER.load(Ordering::SeqCst)
}

/// Stand the cover down without painting: a park owns the band from its first capture, and a
/// take-down owns it by going away (see `park_pinned_player_inner` and `take_down_the_pin_film_host`).
pub(super) fn clear_the_move_cover() {
    PIN_MOVE_COVER.store(false, Ordering::SeqCst);
    if let Ok(mut held) = PIN_MOVE_COVER_SECONDS.lock() {
        *held = None;
    }
}

/// Give the move cover up once the moved player has drawn a frame into its new window: the child
/// window composites above the pin's bitmap, so the cover is invisible from then on, but it must go
/// so that a later park or step does not open on a stale frame. A paused film's clock never moves,
/// which keeps the cover standing until a take-down, a park or a step clears it (see
/// `clear_the_move_cover`).
pub(super) fn settle_the_move_cover() -> bool {
    if !the_move_cover_is_up() {
        return false;
    }

    let pid = VIDEO_PID.load(Ordering::SeqCst);
    if !the_hosted_player_is_drawing(pid) {
        return false;
    }
    // FFplay has no pipe to say its clock moved: its child kept drawing through the reparent, and
    // the adopt road latched the frame it was drawing (see `adopt_the_hover_film_host`).
    if !the_pins_film_is_ffplay() {
        let captured = PIN_MOVE_COVER_SECONDS.lock().ok().and_then(|held| *held);
        let moved =
            mpv_ipc::get_property(pid, "time-pos").and_then(|answer| answer.parse::<f64>().ok());
        if !captured
            .zip(moved)
            .is_some_and(|(captured, moved)| moved > captured)
        {
            return false;
        }
    }

    clear_the_move_cover();
    if !pin_player_is_parked() {
        forget_video_frame();
    }
    if let Some((window, _, _)) = pinned_window_box() {
        // SAFETY: the handle is the pin's own window and the paint is a layered-window blit out of
        // a surface this app drew for exactly this window (see `render_pinned_preview_at`).
        unsafe { render_pinned_preview_at(pinned_window(), window.0, window.1) };
    }
    true
}

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
    PRESENTED_PID.store(0, Ordering::SeqCst);
    // A hosted player has no window of the desktop's to publish, and this one is the film now.
    VIDEO_HWND.store(0, Ordering::SeqCst);
}

/// The host to give a player begun for `path`, where that player is a pinned mpv and the pin
/// stands in a window of its own. `None` is every other launch: a hover's film, FFmpeg's player,
/// and the media engine's own frame. FFmpeg's player is not given a window — it is adopted into
/// the host once it makes its own (see `host_for_the_pins_ffplay`).
pub(super) fn host_for_the_pins_film(path: &Path) -> Option<HWND> {
    (the_film_is_shown_in_a_host(path) && video_route(path) == VideoRoute::Mpv)
        .then(film_host)
        .flatten()
}

/// The host an FFmpeg film begun for `path` is to be adopted into, where this is the pin's own
/// window and the pin stands in it. `None` is every other launch, where the player keeps its own
/// window of the desktop's (see `draws_in_a_window_of_its_own`).
pub(super) fn host_for_the_pins_ffplay(path: &Path) -> Option<HWND> {
    (the_film_is_shown_in_a_host(path) && video_route(path) == VideoRoute::Ffplay)
        .then(film_host)
        .flatten()
}

/// Whether this file is shown in a host this app owns: a film mpv plays or FFmpeg's player draws,
/// pinned in a window of the pin's own. Asked of the route rather than of the player's process
/// because it is the *launch* that decides where the picture is drawn (see `start_video_playback`).
pub(super) fn the_film_is_shown_in_a_host(path: &Path) -> bool {
    pin_is_its_own_window() && matches!(video_route(path), VideoRoute::Mpv | VideoRoute::Ffplay)
}

/// Whether the pin's film is FFmpeg's, whose child is fitted to the host's box on every place:
/// mpv's child sizes itself and is left alone (see `place_the_pin_film_host`).
pub(super) fn the_pins_film_is_ffplay() -> bool {
    PIN_FILM_IS_FFPLAY.load(Ordering::SeqCst)
}

/// Write down which player the pin's film is drawn by, for the questions asked of the host while the
/// pin's own lock is held (see `the_pins_film_is_ffplay`).
pub(super) fn note_the_pins_film_route(path: &Path) {
    PIN_FILM_IS_FFPLAY.store(video_route(path) == VideoRoute::Ffplay, Ordering::SeqCst);
}

/// Fit the player's child to the host's client box, where the pin's film is FFmpeg's and the
/// player is the one the host holds.
pub(super) unsafe fn fit_the_hosted_ffplay_to_the_host(host: HWND) {
    if !the_pins_film_is_ffplay() {
        return;
    }
    let Some(pid) = hosted_player_pid() else {
        return;
    };
    if let Some(child) = a_child_of_the_host(host, pid) {
        unsafe { ffplay_window::fit_the_hosted_player_to_the_host(child, host) };
    }
}

/// Whether the tick owes the player's window the front: a film a hover is following, or a pin's
/// own hosted film. The pin's half is asked of the host rather than of the hover's bookkeeping,
/// because a hover's `Hide` can empty that bookkeeping while the pin still owns the window (see
/// `a_pin_owns_the_window_a_hide_would_take_down`).
pub(super) fn the_video_window_is_owed_the_front(current_video_path: Option<&Path>) -> bool {
    current_video_path.is_some()
        || (pinned() && pin_is_its_own_window() && film_host().is_some())
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
///
/// **And only a shown child with a size counts, never a hidden or empty one.** A player that has
/// just been begun makes its window before it has decoded anything, so existence alone is true
/// for the whole of the decode — which is the empty host the cover was dropped for. What the
/// engine's own report of a first frame is, is a window it shows: visible and with a box, which
/// is what the host is shown for and the cover handed over on (see `host_has_a_child_of`). No
/// timer poll invents it and nothing shown is hidden again.
pub(super) fn the_hosted_player_is_drawing(pid: u32) -> bool {
    let Some(host) = film_host() else {
        return false;
    };
    pid != 0 && hosted_player_pid() == Some(pid) && a_child_of_the_host(host, pid).is_some()
}

/// Whether this player has a window on screen: a window of the desktop's of its own, or the child
/// a hosted player draws under the pin's own host. It is one answer because the settlement roads
/// have one question — has this player arrived — and a hosted player never publishes a `VIDEO_HWND`
/// to be found by (see `video_window_for`, `settle_video_retirement` and `PinSwapHold::settle`).
pub(super) fn the_players_window_is_up(pid: u32) -> bool {
    video_window_for(pid).is_some() || the_hosted_player_is_drawing(pid)
}

/// The window a player of `pid`'s has made under `host`, where it is drawing: the shown child with
/// a box, which is the engine's own report of a first frame (see `the_hosted_player_is_drawing`).
/// The style is read rather than `IsWindowVisible`, because a host is hidden until that frame — a
/// shown child under a hidden host is still the film arriving.
pub(super) fn a_child_of_the_host(host: HWND, pid: u32) -> Option<HWND> {
    unsafe {
        let mut child = GetWindow(host, GW_CHILD).unwrap_or_default();
        while !child.is_invalid() {
            let mut owner = 0u32;
            GetWindowThreadProcessId(child, Some(&mut owner));
            // A made window is not a drawn one: a child without the shown style or with no box
            // is a player still opening, so only the engine's own shown window with a box is taken
            // as its first frame.
            // SAFETY: `child` is a live window found under this app's own host, and the style is
            // a read of its own bits that reports rather than faults for a window on its way out.
            let shown = GetWindowLongPtrW(child, GWL_STYLE) & WS_VISIBLE.0 as isize != 0;
            if owner == pid && shown {
                let mut rect = RECT::default();
                if GetWindowRect(child, &mut rect).is_ok()
                    && rect.right > rect.left
                    && rect.bottom > rect.top
                {
                    return Some(child);
                }
            }
            child = GetWindow(child, GW_HWNDNEXT).unwrap_or_default();
        }
    }

    None
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
///
/// `adopted` is the hover's top-level host on the road that moves an already-running film into the
/// pin: its player is the one drawing the hover, and the window is re-styled and re-parented rather
/// than replaced (see `adopt_the_hover_film_host`). Every other road passes `None` and a fresh child
/// is made where there is none.
pub(super) unsafe fn take_up_the_pin_film_host(adopted: Option<HWND>) {
    let Some(path) = pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.path.clone())) else {
        take_down_the_pin_film_host();
        return;
    };

    if !the_film_is_shown_in_a_host(&path) {
        take_down_the_pin_film_host();
        return;
    }

    // Written before any road that asks the host a question under the pin's own lock can run (see
    // `the_pins_film_is_ffplay`).
    note_the_pins_film_route(&path);

    let Some(content) = pinned_content() else {
        return;
    };

    if let Some(host) = adopted.filter(|host| unsafe { IsWindow(*host) }.as_bool()) {
        unsafe { adopt_the_hover_film_host(host, content) };
        return;
    }

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
    unsafe { fit_the_hosted_ffplay_to_the_host(host) };
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

/// Move the hover's already-running top-level host into the pin's window.
///
/// **The player is not begun again.** Its child keeps drawing where it is; only the window around
/// it changes parent and style. The order is the one Windows needs and the spike proved: clear the
/// topmost band while the host is still top-level, turn `WS_POPUP` into `WS_CHILD`, `SetParent`,
/// clear the topmost bit `SetParent` puts back, then place the host at the pin's band and show it
/// (the pin was shown first, so the host is never revealed with a hidden parent). The host stays
/// disabled, so a press on the picture is still a press on the pin.
unsafe fn adopt_the_hover_film_host(host: HWND, content: ScreenRegion) {
    let Some(parent) = pin_real_window::real_pin_window() else {
        return;
    };
    let pid = hover_film_host::hover_film_host_player().unwrap_or(0);
    let (x, y, width, height) = band_in_the_window(content);

    // Off the screen before the restyle: a visible top-level window changing style and parent is a
    // surface the compositor has to present again, and the pin's cover is the picture until the
    // final show below (see `raise_the_move_cover`). Hiding costs nothing visually because the
    // cover is already on screen.
    let _ = ShowWindow(host, SW_HIDE);

    // The topmost band goes before the style does, while the window is still a window of the
    // desktop's: a child cannot hold it, and SetParent would only put it back.
    let ex = GetWindowLongPtrW(host, GWL_EXSTYLE);
    let ex = ex & !(WS_EX_TOPMOST.0 as isize) & !(WS_EX_TRANSPARENT.0 as isize);
    let _ = SetWindowLongPtrW(host, GWL_EXSTYLE, ex);
    let _ = SetWindowPos(
        host,
        HWND_NOTOPMOST,
        0,
        0,
        0,
        0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
    );

    let style = GetWindowLongPtrW(host, GWL_STYLE);
    let style = (style & !(WS_POPUP.0 as isize)) | WS_CHILD.0 as isize;
    let _ = SetWindowLongPtrW(host, GWL_STYLE, style);
    let _ = SetParent(host, parent);

    // `SetParent` restores the topmost bit for a window that had one; the pin's host wears none.
    let _ = SetWindowLongPtrW(host, GWL_EXSTYLE, ex);
    let _ = SetWindowPos(
        host,
        HWND_NOTOPMOST,
        0,
        0,
        0,
        0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
    );
    let _ = SetWindowPos(
        host,
        HWND_TOP,
        x,
        y,
        width,
        height,
        SWP_NOACTIVATE | SWP_SHOWWINDOW,
    );
    // The child inside the host keeps the box it was fitted to while this was a hover; the pin's
    // band is another box, and FFmpeg's child does not follow its parent (see
    // `fit_the_hosted_player_to_the_host`).
    unsafe { fit_the_hosted_ffplay_to_the_host(host) };
    let _ = EnableWindow(host, false);

    PIN_FILM_HOST.store(host.0 as isize, Ordering::SeqCst);
    PIN_FILM_HOST_PLAYER.store(pid, Ordering::SeqCst);
    // The hover had already proved its frame: keep the latch rather than making the pin ask again.
    PRESENTED_PID.store(pid, Ordering::SeqCst);
    hover_film_host::forget_the_hover_film_host_that_was_handed_over(host);
    VIDEO_HWND.store(0, Ordering::SeqCst);

    pin_chrome_strips::place_the_chrome_strips();
    unsafe { pin_chrome_strips::raise_the_chrome_strips_over_the_host() };
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
    unsafe { fit_the_hosted_ffplay_to_the_host(host) };
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
    // **The player is ended before the window it draws in is destroyed.** Destroying a parent takes
    // its children with it, and a player whose surface was destroyed under it is a kill the player
    // did not get to answer; the record below is what the ordinary kill road retries to confirmation
    // (see `terminate_player_pid`).
    PIN_FILM_IS_FFPLAY.store(false, Ordering::SeqCst);
    if let Some(pid) = hosted_player_pid() {
        terminate_player_pid(pid);
    }
    if let Some(host) = film_host() {
        let _ = DestroyWindow(host);
    }
    PIN_FILM_HOST.store(0, Ordering::SeqCst);
    PIN_FILM_HOST_PLAYER.store(0, Ordering::SeqCst);
    PRESENTED_PID.store(0, Ordering::SeqCst);
    clear_the_move_cover();
    // A cover left over from a move must not open the next pin or park on this film's frame. A
    // park's own frame is its to give up, so it is left standing for the park.
    if !pin_player_is_parked() {
        forget_video_frame();
    }
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

/// When the take-up that is waiting on its hosted film began waiting, while one is.
///
/// A pin taken up on a hover's mpv film keeps the hover up and its own window down until the
/// hosted player draws (see `show_pinned_window`): the wait is this instant, and the tick's
/// handover ends it. Nothing else reads it — the wait is bounded by the player-start timeout
/// the handover settles on, and a pin that ends or collapses disarms it by going away.
static PIN_HANDOVER_SINCE: Lazy<Mutex<Option<Instant>>> = Lazy::new(|| Mutex::new(None));

/// The hover a take-up froze, and the second its picture stands at.
///
/// **It is the picture on screen until the pin can show the same one.** The take-up holds the
/// hover's player where it is rather than ending it, so the frame the user is looking at is mpv's own
/// last frame and nothing this app has to paint or cover: the wait ends when the pin's player reports
/// that frame drawn, and the swap is two windows showing one frame (see `settle_pin_handover`).
#[derive(Clone, Copy)]
struct FrozenHover {
    /// The hover's player, still running and held on the frame it was showing.
    pid: u32,
    /// The second that frame is at — the second the pin's player was begun on, and the one it has to
    /// be standing at before the swap is made.
    seconds: f64,
}

/// The hover a take-up is holding while the wait for its hosted player's frame is on.
static PIN_HANDOVER_HOVER: Lazy<Mutex<Option<FrozenHover>>> = Lazy::new(|| Mutex::new(None));

/// Hold the hover a take-up froze: the player still drawing the frame on screen, and the second it
/// stands at. Written by the take-up that begins the pin's player on that frame, and read by the
/// handover that swaps the two (see `FrozenHover`).
pub(super) fn hold_the_hover_for_the_handover(pid: u32, seconds: f64) {
    if let Ok(mut frozen) = PIN_HANDOVER_HOVER.lock() {
        *frozen = Some(FrozenHover { pid, seconds });
    }
}

/// The hover a take-up is holding, while the wait for its hosted player's frame is on.
fn held_hover() -> Option<FrozenHover> {
    PIN_HANDOVER_HOVER.lock().ok().and_then(|frozen| *frozen)
}

/// The hover a take-up froze, as a test reads it: the player still drawing the picture on screen and
/// the second the pin's player was begun on (see `FrozenHover`).
#[cfg(test)]
pub(super) fn held_hover_for_a_test() -> Option<(u32, f64)> {
    held_hover().map(|frozen| (frozen.pid, frozen.seconds))
}

/// End the hover a take-up froze, where one is still held: it is the picture no longer, and nothing
/// else holds it — the player was taken out of the media at the take-up, so this is the only call
/// that will ever end it. A kill not confirmed on the spot is parked with the orphans the loop
/// retries rather than dropped, which is the same road a pin's own teardown uses (see
/// `retire_orphaned_player`).
fn retire_the_frozen_hover(frozen: Option<FrozenHover>) {
    if let Some(frozen) = frozen {
        // The host the hover was drawn in is the picture no longer, and the pin it was waiting
        // for has already been shown: destroying the host ends mpv with it, so the retire below
        // is the confirmation rather than the picture's own end (see `hover_film_host`).
        unsafe { hover_film_host::take_down_the_hover_film_host_if_it_held(frozen.pid) };
        retire_orphaned_player(frozen.pid);
    }
}

/// Whether a pin taken up on a hover's film is still waiting on its hosted player's first frame:
/// its own window down, the hover up as the placeholder, and no chrome over either.
///
/// Asked by the roads that must leave a waiting pin alone: the tick's restore, which would put a
/// hidden-on-purpose window back up, and the chrome refresh, which would spend the arrival over
/// nothing (see `the_pin_owes_a_show` and `refresh_pin_chrome`).
pub(super) fn pin_handover_armed() -> bool {
    PIN_HANDOVER_SINCE
        .lock()
        .ok()
        .and_then(|since| *since)
        .is_some()
}

/// Arm the handover a take-up on a hover's film waits on, from the moment the hosted player is
/// begun in the background (see `show_pinned_window`).
pub(super) fn arm_pin_handover() {
    if let Ok(mut since) = PIN_HANDOVER_SINCE.lock() {
        *since = Some(Instant::now());
    }
}

/// Disarm the handover: the film drew and the pin went up, the wait failed and the pin is being
/// torn down, or the pin went away down another road and there is nothing left to wait on. The hover
/// the take-up froze goes with it: a wait that is over has no picture to hold (see `FrozenHover`).
pub(super) fn disarm_pin_handover() {
    if let Ok(mut since) = PIN_HANDOVER_SINCE.lock() {
        *since = None;
    }
    if let Ok(mut frozen) = PIN_HANDOVER_HOVER.lock() {
        *frozen = None;
    }
}

/// Stand a handover wait's beginning in place of the take-up's own, or give the take-up's own
/// answer back with `None` — a test of the timeout the wait is bounded by, which cannot wait out
/// the start timeout itself (see `settle_pin_handover`).
#[cfg(test)]
pub(super) fn stand_pin_handover_armed_at(since: Option<Instant>) {
    if let Ok(mut armed) = PIN_HANDOVER_SINCE.lock() {
        *armed = since;
    }
}

/// What the wait for a hosted film's first frame has come to.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum PinHandover {
    /// The film is not on screen yet and the player is still starting: the pin stays down and
    /// the hover stays up, and the wait goes on.
    Waiting,
    /// The hosted player drew: the pin went up over its live film with the chrome arriving,
    /// and the hover came down in the same pass.
    HandedOver,
    /// The player errored or never drew inside the player-start timeout: the hidden pin is for
    /// the teardown, and the hover it never took down is what is kept.
    Failed,
}

/// How far from the second it was begun on a hosted player may be and still be standing on the frame
/// the hover froze.
///
/// A frame's own timestamp is not the second it was sought to, and a film whose frames are a second
/// apart is a film whose frame at 12.345 is stamped 12.000. It is not a tolerance on the swap — the
/// player is begun paused, so it cannot be anywhere but on that frame — it is the width of the gap
/// between a seek target and the frame that answers it (see `the_hosted_frame_is_presented`).
const PIN_HANDOVER_FRAME_SECS: f64 = 1.0;

/// Whether this hosted player has a frame of its own on screen: a shown child of the host's *and* a
/// video output the player says is configured for a frame.
///
/// A window is not a picture (see `the_hosted_player_is_drawing`), and `vo-configured` is the
/// player's own word that its video output has been set up for a frame. The answer is latched per
/// player because the read is bounded by a pipe timeout rather than free: once this player has
/// proved a frame it has one, and every later tick answers at once. A step begins a new player, so
/// `note_hosted_player` clears the latch for it.
///
/// **FFmpeg's player has no pipe to answer that word.** For it the frame's proof is its adopted
/// child: a window that is the player's whole picture, adopted only after the player made and
/// showed it, is the fact — and a pipe read is mpv's own answer (see `adopt_the_hover_film_host`,
/// which latches on the same fact). The MPV road is left exactly as it was.
pub(super) fn the_hosted_frame_is_up(pid: u32) -> bool {
    if pid != 0 && PRESENTED_PID.load(Ordering::SeqCst) == pid {
        return true;
    }
    if !the_hosted_player_is_drawing(pid) {
        return false;
    }
    if the_pins_film_is_ffplay() {
        PRESENTED_PID.store(pid, Ordering::SeqCst);
        return true;
    }
    if mpv_ipc::get_property(pid, "vo-configured").as_deref() != Some("true") {
        return false;
    }
    PRESENTED_PID.store(pid, Ordering::SeqCst);
    true
}

/// Whether the hosted player has the frame the hover froze on screen, rather than only a window.
///
/// **A window is not a picture.** mpv makes its child under the host before it has decoded anything,
/// and a shown child with a box is the most `the_hosted_player_is_drawing` can see — which leaves the
/// band black until the first frame is drawn into it, and is the empty stage the whole wait exists to
/// keep behind the hover. The two facts that separate a window from a picture are the player's own,
/// read over its pipe: `time-pos` has no value until a file is loaded and positioned, and
/// `vo-configured` is true once the video output is configured for a frame. **Both**, because either
/// alone is a weaker question — a position is where the file is and not what has been drawn, and a
/// configured video output is a size and a format and not a frame.
///
/// The player is begun paused on that second (see `video_launch::mpv_args`), so the pair is not a
/// race: it is true exactly while the frame at that second is the one on screen, and the unpause at
/// the handover is the first thing that moves it. A read that does not answer is a player that has
/// not proved anything, so the wait goes on — bounded by the player-start timeout, never by a timer
/// of its own (see `VIDEO_START_WAIT_SECS`).
pub(super) fn the_hosted_frame_is_presented(pid: u32, at: f64) -> bool {
    let Some(seconds) =
        mpv_ipc::get_property(pid, "time-pos").and_then(|answer| answer.parse::<f64>().ok())
    else {
        return false;
    };

    the_hosted_frame_is_up(pid) && (seconds - at).abs() <= PIN_HANDOVER_FRAME_SECS
}

/// Hand a pin taken up on a hover's film over once its hosted player has the frame on screen, and
/// tear the hidden pin down where that player never does.
///
/// The only visible swap, and it is one frame: the pin's window appears with the film already showing
/// the frame the hover was frozen on — no cover, no black stage — the arrival starts over that live
/// picture, the hover comes down, the hosted player is let go of, and the frozen hover is ended, all
/// in the same pass (see `show_pinned_window`). What the wait waits on is the player's own word that
/// the frame is drawn and not a window that has been made (see `the_hosted_frame_is_presented`), and
/// nothing here is a timer: a player that takes its time is waited for, up to the player-start bound.
/// Where the player errored or the start timeout ran out, the wait is over the other way: the pin
/// never appears, the hover it never took down is what is kept, and the hover's own player is ended
/// rather than left frozen for a handover that is never coming (see `Reason::MediaGone`).
///
/// One shot: the arm goes down however the wait ends, so a handover never fires twice and a
/// failure never sticks behind a hidden pin.
pub(super) unsafe fn settle_pin_handover() -> PinHandover {
    let Some(since) = PIN_HANDOVER_SINCE.lock().ok().and_then(|since| *since) else {
        return PinHandover::Waiting;
    };
    let frozen = held_hover();

    // A pin that went away down another road — closed, or collapsed into its bubble — leaves no
    // wait behind it: the close owns the teardown and the collapse owns the pin now. The hover the
    // take-up froze is ended with it, because the pin that was going to take the picture from it is
    // not going to.
    if !pinned() || pin_is_collapsed() {
        disarm_pin_handover();
        retire_the_frozen_hover(frozen);
        return PinHandover::Waiting;
    }

    // The film's own frame on screen is the handover: the wait is over, and what is on screen from
    // here is the pin over that picture rather than the hover as its placeholder. A frame this app
    // cannot prove is a frame it waits for — the read failing is not an error, it is not-yet.
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    if the_hosted_player_is_drawing(pid)
        && frozen.is_some_and(|frozen| the_hosted_frame_is_presented(pid, frozen.seconds))
    {
        disarm_pin_handover();
        // The handover is the pin's first frame, not an arrival: the chrome stays hidden over it
        // until the pointer asks for a strip (see `refresh_pin_chrome`).
        with_pin(|pin| pin.chrome = PinChrome::hidden());
        unsafe { show_pinned_window(pinned_window()) };

        // **And the film is playing from here, unless the user had held it.** The player was begun
        // paused so that the frame the hover froze is the frame it has; letting it go of that hold is
        // this moment and no earlier one — and the clock the pass is measured by moves onto the
        // second it is let go at, because a clock begun at the launch has counted the whole wait
        // (see `PinTransport::released` and `rebase_video_loop_clock_to`).
        if let Some(frozen) = frozen.filter(|_| !pinned_is_held()) {
            mpv_ipc::set_paused(pid, false);
            video_hw::rebase_video_loop_clock_to(pid, frozen.seconds);
            update_pin_transport(|transport| transport.released(frozen.seconds));
        }

        // The pin has the picture, so the hover's player is the picture no longer: ended now that its
        // window is behind the pin's, rather than left frozen and unowned (see `FrozenHover`).
        retire_the_frozen_hover(frozen);
        return PinHandover::HandedOver;
    }

    // No player behind the wait is a player that errored rather than one still starting: the
    // pin it was begun for never appears, and the hover it never took down is what is kept.
    if !is_video_process_running() {
        disarm_pin_handover();
        retire_the_frozen_hover(frozen);
        return PinHandover::Failed;
    }

    // And a player that never draws inside the player-start timeout is one that will not: a
    // hidden pin waiting past a start ever takes is a stuck one (see `VIDEO_START_WAIT_SECS`).
    if since.elapsed() >= Duration::from_secs(VIDEO_START_WAIT_SECS) {
        disarm_pin_handover();
        retire_the_frozen_hover(frozen);
        return PinHandover::Failed;
    }

    PinHandover::Waiting
}

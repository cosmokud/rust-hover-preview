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
use windows::Win32::Graphics::Gdi::{
    FillRect, GetStockObject, SetBrushOrgEx, SetStretchBltMode, StretchDIBits, BLACK_BRUSH,
    HALFTONE, HBRUSH, NULL_BRUSH, STRETCH_BLT_MODE,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::EnableWindow;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, GetClientRect, GetParent, GetWindow,
    GetWindowLongPtrW, GetWindowRect, GetWindowThreadProcessId, RegisterClassExW, SetParent,
    SetWindowLongPtrW, SetWindowPos, ShowWindow, GWLP_HWNDPARENT, GWL_EXSTYLE, GWL_STYLE, GW_CHILD,
    GW_HWNDNEXT, HWND_NOTOPMOST, HWND_TOP, HWND_TOPMOST, SC_CLOSE, SWP_ASYNCWINDOWPOS,
    SWP_NOACTIVATE, SWP_NOCOPYBITS, SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER, SWP_SHOWWINDOW, SW_HIDE,
    SW_SHOWNOACTIVATE, WINDOW_EX_STYLE, WM_CLOSE, WM_ERASEBKGND, WM_SIZE, WM_SYSCOMMAND,
    WNDCLASSEXW, WS_CHILD, WS_CLIPCHILDREN, WS_EX_APPWINDOW, WS_EX_TOPMOST, WS_EX_TRANSPARENT,
    WS_POPUP, WS_VISIBLE,
};

use super::pin_trace::note_pin_trace;
use super::*;

/// The class a film's host is made from. A class of its own rather than the pin window's: what a
/// `WS_CHILD` host answers to is a window that draws nothing and takes nothing, and the pin's
/// window procedure is the pin's (see `pin_real_window`).
const FILM_HOST_CLASS: windows::core::PCWSTR = w!("RustHoverPreviewPinFilmHost");

/// The default answer is most of what the host does: it draws nothing of the film's own and no
/// press lands on it, because the picture is mpv's child (see the module doc) — and on `WM_SIZE` it
/// puts that child at its own client box, and on `WM_ERASEBKGND` it fills the band the child leaves
/// (see `paint_the_host_behind_the_film`). This bindings version exposes `DefWindowProcW` as a
/// Rust-callable function, so the class still needs a system-ABI shim around it.
unsafe extern "system" fn the_default_answer(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    // **The strip a grown band has gained is painted here rather than erased to black**, and the
    // class wears `NULL_BRUSH` so that this is the only hand that draws it (see
    // `paint_the_host_behind_the_film` and `create`). `1` is "handled": it is what keeps
    // `DefWindowProcW` from erasing the same pixels a second time.
    if msg == WM_ERASEBKGND {
        unsafe { paint_the_host_behind_the_film(hwnd, HDC(wparam.0 as *mut _)) };
        return LRESULT(1);
    }

    // **The child's box is given in the message that gave the host its own**, so the film is the
    // band's size before this app composites the pin over it rather than a size or more behind (see
    // `fit_the_hosted_mpv_to_the_host`).
    if msg == WM_SIZE {
        fit_the_hosted_mpv_to_the_host(hwnd);
    }

    // **A host that is the Shell's window is the Shell's close too.** Its taskbar button's Close
    // arrives here as a `WM_CLOSE` or an `SC_CLOSE`, and taken it would destroy the window the pin's
    // chrome is owned by — the pin with it, under a state that still believes a pin is up. What a
    // close means for the pin is the pin's own road's answer (see `window_proc`).
    if PIN_FILM_HOST_IS_TOP_LEVEL.load(Ordering::SeqCst)
        && (msg == WM_CLOSE || (msg == WM_SYSCOMMAND && (wparam.0 & 0xFFF0) as u32 == SC_CLOSE))
    {
        request_pin_end();
        return LRESULT(0);
    }

    unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
}

/// Paint the part of the host the film's child does not cover: the frame a drag is holding,
/// stretched into the box, or black where there is none in hand.
///
/// **A band that has just grown is this app's to fill until the child is its size.** mpv gives its
/// child the host's new box from the message that gave the host its own, and does it on its own
/// thread (see `fit_the_hosted_mpv_to_the_host`), so the strip the band has gained is drawn by
/// nobody until then — and it used to be erased black, which is the flash this fills over.
///
/// **The destination is the whole client box and not the strip**, because the frame is a picture of
/// the band and is scaled *into* the band; the host's `WS_CLIPCHILDREN` clips the part of it the
/// live child stands over, which is what keeps the film untouched (see `create`).
///
/// Both hosts a pin can be showing its film through are erased through here — the pin's own class,
/// and the hover's a pin took up (see `hover_film_host`) — and a host that is not the pin's is
/// erased black, which is what the class brush did before it.
pub(super) unsafe fn paint_the_host_behind_the_film(host: HWND, dc: HDC) {
    let mut client = RECT::default();
    if unsafe { GetClientRect(host, &mut client) }.is_err() {
        return;
    }

    if film_host() == Some(host)
        && !the_pins_film_is_ffplay()
        && unsafe { stretch_the_held_frame(dc, client.right, client.bottom) }
    {
        return;
    }

    // SAFETY: the brush is a stock object this app neither made nor owns, and the rectangle is the
    // client box read above; GDI reports rather than faults for a device context that has gone.
    let brush = HBRUSH(unsafe { GetStockObject(BLACK_BRUSH) }.0);
    unsafe { FillRect(dc, &client, brush) };
}

/// Stretch the frame a drag is holding over a `width` by `height` client box, answering whether GDI
/// drew it.
///
/// The frame is stretched out of its own bits rather than through a surface made for it: it is
/// already a block of BGRA pixels, and a device context and bitmap made for the width of one erase
/// and thrown away after it is two handles and a display's worth of memory per pointer move.
unsafe fn stretch_the_held_frame(dc: HDC, width: i32, height: i32) -> bool {
    let held = held_video_frame();
    let Some(frame) = held.as_ref() else {
        return false;
    };
    if frame.width == 0
        || frame.height == 0
        || frame.pixels.len() != frame.width as usize * frame.height as usize * 4
    {
        return false;
    }

    let info = BITMAPINFO {
        bmiHeader: BITMAPINFOHEADER {
            biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
            biWidth: frame.width as i32,
            biHeight: -(frame.height as i32),
            biPlanes: 1,
            biBitCount: 32,
            biCompression: BI_RGB.0,
            ..Default::default()
        },
        ..Default::default()
    };

    // SAFETY: the bits are the held frame's own, of the size checked above, and `dc` is the one the
    // host was asked to erase with; GDI reports rather than faults for a DC that has gone.
    unsafe {
        // GDI's own filtering rather than a nearest-neighbour stretch, the way the taskbar's peek
        // is scaled: the strip is a few pixels of growth at a time, and a soft edge is the picture
        // where a hard one is a stair (see `pin_taskbar_mirror`). The brush origin travels with the
        // mode, because halftone measures the destination from it.
        let mode = SetStretchBltMode(dc, HALFTONE);
        let _ = SetBrushOrgEx(dc, 0, 0, None);
        let drawn = StretchDIBits(
            dc,
            0,
            0,
            width,
            height,
            0,
            0,
            frame.width as i32,
            frame.height as i32,
            Some(frame.pixels.as_ptr().cast()),
            &info,
            DIB_RGB_COLORS,
            SRCCOPY,
        );
        if mode != 0 {
            let _ = SetStretchBltMode(dc, STRETCH_BLT_MODE(mode));
        }
        drawn > 0
    }
}

/// Whether the resize drag in flight took the frame the host paints with, so that its own end gives
/// it up and no road's frame is dropped with it (see `hold_the_frame_a_growing_host_paints`).
static HOST_FILL_HELD: AtomicBool = AtomicBool::new(false);

/// Capture the frame a growing host paints the strip it gains with.
///
/// **Taken on the pointer message that begins a resize drag**, before the band has moved: what is
/// wanted is the picture that was on screen when the hand took the edge. It is the held-frame road
/// a park paints its band from rather than a capture of this app's own, because a frame read any
/// other way — `PrintWindow`, a decode — is a black rectangle or a whole decode on the message a
/// drag begins on (see `hold_video_window_frame`).
///
/// Only mpv's film, and only where the host is on screen: FFmpeg's child is given the host's new box
/// by the call that gives the host its own (see `fit_the_hosted_ffplay_to_the_host`), and a host
/// that is not on screen has no picture of the film's to read — the desktop where it stands is the
/// pin's chrome or somebody else's window.
pub(super) fn hold_the_frame_a_growing_host_paints() {
    HOST_FILL_HELD.store(false, Ordering::SeqCst);
    if the_pins_film_is_ffplay() || !the_pinned_player_was_begun_in_the_host() {
        return;
    }
    if !film_host().is_some_and(|host| unsafe { IsWindowVisible(host) }.as_bool()) {
        return;
    }

    let taken = hold_video_window_frame();
    HOST_FILL_HELD.store(taken, Ordering::SeqCst);
}

/// Give up the frame a growing host painted with, where this drag took one: the band is the film's
/// again by the time a drag ends, and a frame kept past it is a picture a later erase would paint
/// the strip with.
///
/// A frame a park or the move cover is painting from is theirs and is left standing — their roads
/// give it up when the band has a player in it again (see `settle_the_move_cover`,
/// `take_the_park_down`).
pub(super) fn give_up_the_frame_a_growing_host_paints() {
    if !HOST_FILL_HELD.swap(false, Ordering::SeqCst) {
        return;
    }
    if !the_move_cover_is_up() && !pin_player_is_parked() {
        forget_video_frame();
    }
}

/// The host of the pin that is up, and the player this app has attached to it.
static PIN_FILM_HOST: AtomicIsize = AtomicIsize::new(0);
static PIN_FILM_HOST_PLAYER: AtomicU32 = AtomicU32::new(0);
/// The hosted player whose frame this app has proved on screen, so the pipe read that proves it is
/// asked once and not on every tick (see `the_hosted_frame_is_up`).
static PRESENTED_PID: AtomicU32 = AtomicU32::new(0);
/// mpv's own child under the host, and the host and player it was found under and for: a resize step
/// asks it on every pointer move, and the handle is only ever a hint (see `the_hosted_mpv_child`).
static HOSTED_MPV_CHILD: AtomicIsize = AtomicIsize::new(0);
static HOSTED_MPV_CHILD_HOST: AtomicIsize = AtomicIsize::new(0);
static HOSTED_MPV_CHILD_PID: AtomicU32 = AtomicU32::new(0);
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
    // Another player draws another window, whatever handle the last one left behind.
    forget_the_hosted_mpv_child();
    // A hosted player has no window of the desktop's to publish, and this one is the film now.
    VIDEO_HWND.store(0, Ordering::SeqCst);
}

/// The host to give a player begun for `path`, where that player is a pinned mpv and the pin
/// stands in a window of its own. `None` is every other launch: a hover's film, FFmpeg's player,
/// and the media engine's own frame. FFmpeg's player is not given a window — it is adopted into
/// the host once it makes its own (see `host_for_the_pins_ffplay`).
pub(super) fn host_for_the_pins_film(path: &Path, kind: Option<MediaType>) -> Option<HWND> {
    (the_film_is_shown_in_a_host(path, kind) && video_route(path) == VideoRoute::Mpv)
        .then(film_host)
        .flatten()
}

/// The host an FFmpeg film begun for `path` is to be adopted into, where this is the pin's own
/// window and the pin stands in it. `None` is every other launch, where the player keeps its own
/// window of the desktop's (see `draws_in_a_window_of_its_own`).
pub(super) fn host_for_the_pins_ffplay(path: &Path, kind: Option<MediaType>) -> Option<HWND> {
    (the_film_is_shown_in_a_host(path, kind) && video_route(path) == VideoRoute::Ffplay)
        .then(film_host)
        .flatten()
}

/// **Which road a pin's mpv film is drawn on: the default is the film's host owning the pin's
/// chrome, and `RHP_MPV_CHILD_HOST_PIN`, present, asks for the hosted child host instead.**
///
/// A `WS_CHILD` is always drawn above its parent's client area, and a `WS_EX_LAYERED` parent — the
/// pin's chrome, painted through `UpdateLayeredWindow` — cannot clip or synchronise a hardware
/// swapchain child. So with nothing set the film is given not a child of the pin but a **window of
/// the desktop's of its own, which owns the pin's chrome**: Windows then keeps the chrome above the
/// film, hides it with the host, keeps it out of Alt+Tab, and moves neither for this app (see
/// `the_pin_film_host_is_a_window_of_the_desktop`). **Not a setting**: the variable brings the child
/// host this app stood in before it back for an A/B that needs no rebuild.
const MPV_CHILD_HOST_PIN_VARIABLE: &str = "RHP_MPV_CHILD_HOST_PIN";

/// Whether the film host owns the pin's chrome, which is the road unless a machine asked for the
/// child host (see [`MPV_CHILD_HOST_PIN_VARIABLE`]). The environment is read here and nowhere else,
/// and the answer is a fact of the process: flipping the variable needs the app restarted, not the
/// film re-pinned.
pub(super) fn the_film_host_owns_the_pin() -> bool {
    #[cfg(test)]
    if let Some(answer) = the_owned_road_stand_in() {
        return answer;
    }

    std::env::var_os(MPV_CHILD_HOST_PIN_VARIABLE).is_none()
}

/// Stand an answer for [`the_film_host_owns_the_pin`] on this thread in place of the machine's own, or
/// give the machine's own answer back with `None`.
///
/// **Per thread because the tests run at once**, and because every road that asks the question in a
/// test asks it from the thread the test is on: one test standing the road off is not another — and a
/// run whose environment set the opt-out would otherwise hand every test written against the owned
/// road the child host (see `pin_film_host::stand_the_mpv_pin_hosted_in` for the process-wide
/// form, which is the right one only where the answer is read from no thread in particular).
#[cfg(test)]
pub(super) fn stand_the_owned_road_in(answer: Option<bool>) {
    OWNED_ROAD_STAND_IN.with(|held| held.set(answer));
}

/// The answer [`the_film_host_owns_the_pin`] is given while a test stands one in.
#[cfg(test)]
fn the_owned_road_stand_in() -> Option<bool> {
    OWNED_ROAD_STAND_IN.with(|held| held.get())
}

#[cfg(test)]
thread_local! {
    static OWNED_ROAD_STAND_IN: std::cell::Cell<Option<bool>> = const { std::cell::Cell::new(None) };
}

/// Whether a film of `route` at a pin's band is drawn in a window of the desktop's that owns the
/// pin's chrome, with the machine's own answer given as a parameter: so a test is a call rather
/// than an environment write every other test in the process would see.
///
/// mpv's only. FFmpeg's player is drawn into a host of the pin's for its renderer's sake and keeps
/// the road it has (see `a_hosted_film_is_presented_composed`).
pub(super) fn a_film_route_is_drawn_in_a_host_of_its_own(route: VideoRoute, owned: bool) -> bool {
    owned && route == VideoRoute::Mpv
}

/// Whether the film at this pin's band is drawn in a window of the desktop's that owns the pin's
/// chrome: the pin stands in a window of its own, its media is a film, and the film is mpv's.
pub(super) fn the_pin_film_host_is_a_window_of_the_desktop(path: &Path) -> bool {
    pin_is_its_own_window()
        && a_film_route_is_drawn_in_a_host_of_its_own(
            video_route(path),
            the_film_host_owns_the_pin(),
        )
}

/// Whether the host that is up is a window of the desktop's — written where it is made, because a
/// host's level is a fact about the window and not about the setting that asked for it: a file
/// stepped between mpv and FFmpeg is another level, and the host is made again for it.
static PIN_FILM_HOST_IS_TOP_LEVEL: AtomicBool = AtomicBool::new(false);

/// Whether the pin's chrome is owned by its film host, and so is not the window Windows lists and
/// gives a taskbar button to: the host is that window on this road (see `own_the_pin_window`).
pub(super) fn the_pins_chrome_is_owned() -> bool {
    PIN_FILM_HOST_IS_TOP_LEVEL.load(Ordering::SeqCst) && film_host().is_some()
}

/// The window of a pin that holds the band the pair belongs in: the host that owns the pin's chrome,
/// on the road that has a pair, and the pin's own window on every other.
///
/// **One window of the pair owns the band, and it is the host.** Windows keeps an owned window in its
/// owner's band — the owner put in the topmost band puts the chrome in it, and the owner taken out
/// takes the chrome out with it (measured) — so a band asked of the chrome alone leaves the pair half
/// over another program's window and half behind it: the chrome in front, answering every press in
/// the band it is nearly transparent in, with the film it belongs to still under that program. That
/// is a pin that looks behind a window and yet takes the hand, and it is the one thing this window is
/// asked for (see `place_the_pin_in_the_band`).
pub(super) fn the_pin_window_that_holds_the_band(pin: HWND) -> HWND {
    the_pins_chrome_is_owned()
        .then(film_host)
        .flatten()
        .unwrap_or(pin)
}

/// Put a pin in a band — the topmost one, the ordinary one, or the place behind another window — as
/// **one window**, so that every road that raises, demotes or lowers a pin moves the film and the
/// chrome together (see `the_pin_window_that_holds_the_band`).
///
/// **A topmost band the desktop did not give is asked for again, from outside it**, and this is the one
/// place that answer lives rather than in each raise (see `raise_pinned_preview`, which used to ask it
/// for a held band): the desktop answers a band assignment as made and can still not make it for a
/// window it has not laid out yet — one shown this moment — and a pair the desktop banded on one side
/// only is the split this window exists to end (measured: the chrome in the band and the host out of
/// it).
///
/// Both windows are this app's own: nothing here is a call into another process.
pub(super) unsafe fn place_the_pin_in_the_band(pin: HWND, after: HWND) {
    let window = the_pin_window_that_holds_the_band(pin);
    let _ = SetWindowPos(
        window,
        after,
        0,
        0,
        0,
        0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
    );

    if after == HWND_TOPMOST && !pin_lifecycle::window_holds_the_topmost_band(window) {
        let _ = SetWindowPos(
            window,
            HWND_NOTOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
        let _ = SetWindowPos(
            window,
            HWND_TOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }
}

/// Whether the pin's film host is a child of the pin's own window, which is where a host made or
/// adopted by the pin stands. A host still a window of the desktop's is a handover's on its way in,
/// or the host of a pin whose chrome is owned by it (see `the_pin_film_host_is_a_window_of_the_desktop`).
pub(super) fn the_pin_film_host_stands_in_the_pins_own_window() -> bool {
    if PIN_FILM_HOST_IS_TOP_LEVEL.load(Ordering::SeqCst) {
        return false;
    }
    let (Some(host), Some(parent)) = (film_host(), pin_real_window::real_pin_window()) else {
        return false;
    };
    unsafe { GetParent(host) }.is_ok_and(|found| found == parent)
}

/// Take a pinned FFmpeg player that was begun as a window of the desktop's into the pin's host,
/// rather than beginning it again: the take-up that finds the pin's own player already running
/// adopts it by the same road the monitor uses once the window exists.
pub(super) fn adopt_a_stranded_ffplay_into_the_pin_host(pid: u32, path: &Path) -> bool {
    if pid == 0 || video_route(path) != VideoRoute::Ffplay {
        return false;
    }
    let Some(host) = film_host() else {
        return false;
    };
    let window = ffplay_window::video_window_for(pid);
    note_the_pins_film_route(path);
    note_hosted_player(pid);
    if let Some(window) = window {
        unsafe { ffplay_window::adopt_the_players_window(window, host) };
    }
    true
}

/// Whether FFmpeg's player is presented through the composed road — SDL's Direct3D 9 renderer
/// blitting into a window DC — rather than through a flip-model DXGI swapchain.
///
/// **It is the launch's own decision, asked here rather than made here** (see
/// `video_playback::the_renderer_a_film_is_begun_with`, which is where the flicker and the reason for
/// the composed road are written out).
///
/// **And it is one answer for the hover and the pin.** A hosted film is begun once, for whatever
/// window will hold it; the pin it is later moved into has to present it the same way, or the picture
/// would be drawn into a layered surface its renderer cannot reach. The pin's own form and the
/// player's renderer are therefore one decision (see `the_pin_presents_a_hosted_film_composed`).
pub(super) fn ffplay_is_presented_composed() -> bool {
    video_playback::the_film_is_presented_composed()
}

/// **A temporary diagnostic, and not a setting: `RHP_MPV_COMPOSED_PIN`, present means on.**
///
/// A pinned mpv film is drawn into a child of a **layered** pin window, so every paint of the pin
/// re-sends the whole surface through `UpdateLayeredWindow` while the player's swapchain presents
/// its own frames on top of it. A pinned FFmpeg film takes the composed road instead — a pin window
/// with no `WS_EX_LAYERED`, painted with `BitBlt` — and is smooth under a drag-resize, while the mpv
/// pin is shaky. This asks, by hand, whether the road is the whole of that difference: with the
/// variable set, a hosted mpv film takes the composed road exactly as ffplay's does.
///
/// **It should be deleted once the answer is in** — either the composed road is what fixes the
/// shake and mpv is moved onto it for good, or it is not and this goes. The environment is read
/// here and nowhere else, and the answer is a fact of the process: flipping the variable needs the
/// app restarted, not the film re-pinned.
const MPV_COMPOSED_PIN_VARIABLE: &str = "RHP_MPV_COMPOSED_PIN";

/// Whether the machine asked a pinned mpv film to take the composed road (see
/// [`MPV_COMPOSED_PIN_VARIABLE`]).
pub(super) fn the_mpv_pin_is_composed() -> bool {
    std::env::var_os(MPV_COMPOSED_PIN_VARIABLE).is_some()
}

/// Whether the pin on screen holds a hosted film presented through the composed road. It is
/// the one predicate behind three things that must never disagree: the player's renderer (the composed
/// one, unless a machine named another by hand), the pin window's style (no `WS_EX_LAYERED`), and the
/// paint road (`GDI` into the window DC rather than `UpdateLayeredWindow`, which is chosen from the
/// style itself — see `paint::paints_through_the_window_dc`).
pub(super) fn the_pin_presents_a_hosted_film_composed() -> bool {
    if !pin_is_its_own_window() {
        return false;
    }

    let Some(path) = pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.path.clone())) else {
        return false;
    };

    // **A pin whose film is drawn in a window of the desktop's is the layered surface over it,
    // whatever else a machine asked for.** The picture stands below this one, so the band has to be
    // an alpha the host shows through and a press falls on — a composed pin is blitted through its
    // own DC and is opaque over the whole of itself, which is the film gone. Asked of the road rather
    // than of the host that is up, because the form is written at a take-up, before there is one (see
    // `the_pin_film_host_is_a_window_of_the_desktop`).
    if the_pin_film_host_is_a_window_of_the_desktop(&path) {
        return false;
    }

    // A sound's player is FFmpeg's too: only a film at this pin's band is presented this way, and
    // a placeholder counts because the pin is taken up before its film is installed.
    current_media_type().is_none_or(|kind| {
        matches!(
            kind,
            MediaType::Video | MediaType::Loading | MediaType::Unplayable
        )
    }) && a_hosted_film_is_presented_composed(
        video_route(&path),
        ffplay_is_presented_composed(),
        the_mpv_pin_is_composed(),
    )
}

/// The decision itself, with each road's own answer as a parameter: the environment is read in one
/// place ([`the_mpv_pin_is_composed`]), so a test is a call rather than an environment write every
/// other launch in the process would see.
///
/// FFmpeg's player is composed unless a machine named a flip-model renderer by hand; mpv is composed
/// only where the temporary diagnostic asks for it.
pub(super) fn a_hosted_film_is_presented_composed(
    route: VideoRoute,
    ffplay_presented_composed: bool,
    mpv_asked_by_hand: bool,
) -> bool {
    match route {
        VideoRoute::Ffplay => ffplay_presented_composed,
        VideoRoute::Mpv => mpv_asked_by_hand,
        _ => false,
    }
}

/// Whether this file is shown in a host this app owns: a film mpv plays or FFmpeg's player draws,
/// pinned in a window of the pin's own. Asked of the route rather than of the player's process
/// because it is the *launch* that decides where the picture is drawn (see `start_video_playback`).
pub(super) fn the_film_is_shown_in_a_host(path: &Path, kind: Option<MediaType>) -> bool {
    // A sound's player is FFmpeg's too, and a sound's route would match the film routes below: the
    // media on screen has to be a film before any host road is asked (see `MediaType::is_audio`).
    pin_is_its_own_window()
        && kind.or_else(current_media_type) == Some(MediaType::Video)
        && matches!(video_route(path), VideoRoute::Mpv | VideoRoute::Ffplay)
}

/// Whether a pinned video FFmpeg's player draws stands in the pin's own host rather than a window
/// of the desktop's. The narrow gate every host road asks, so that a hover and an audio card keep
/// the windows they always had (see `MediaType::draws_in_a_window_of_its_own`).
pub(super) fn a_pinned_ffplay_film_is_drawn_in_the_pins_host(path: &Path) -> bool {
    pin_is_its_own_window() && video_route(path) == VideoRoute::Ffplay
}

/// Whether a pinned film is drawn in a host rather than a window of the desktop's: mpv always, and
/// FFmpeg's player only when this pin hosts it. The drag and relayout roads reuse mpv's host arm.
pub(super) fn the_pinned_film_is_drawn_in_a_host(path: &Path) -> bool {
    video_route(path) == VideoRoute::Mpv || a_pinned_ffplay_film_is_drawn_in_the_pins_host(path)
}

/// Whether the pin's film is FFmpeg's, whose child is fitted to the host's box by the road that
/// placed the host: mpv's child is fitted by the host's own `WM_SIZE`, because mpv sizes it on its
/// own thread and at its own pace (see `fit_the_hosted_mpv_to_the_host`).
pub(super) fn the_pins_film_is_ffplay() -> bool {
    PIN_FILM_IS_FFPLAY.load(Ordering::SeqCst)
}

/// Write down which player the pin's film is drawn by, for the questions asked of the host while the
/// pin's own lock is held (see `the_pins_film_is_ffplay`).
pub(super) fn note_the_pins_film_route(path: &Path) {
    PIN_FILM_IS_FFPLAY.store(video_route(path) == VideoRoute::Ffplay, Ordering::SeqCst);
    // The film a step is on is another film: its player may have made itself another child.
    forget_the_hosted_mpv_child();
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

/// Forget the child mpv was drawing in: called where this app writes down which player draws the
/// pin's film, because a film a step is on — or another player begun into the same host — is another
/// window of mpv's (see `the_hosted_mpv_child`).
fn forget_the_hosted_mpv_child() {
    HOSTED_MPV_CHILD.store(0, Ordering::SeqCst);
    HOSTED_MPV_CHILD_HOST.store(0, Ordering::SeqCst);
    HOSTED_MPV_CHILD_PID.store(0, Ordering::SeqCst);
}

/// The window mpv draws the pin's film in under `host`, where this host holds a player of mpv's.
///
/// **It is found the way FFmpeg's child is found** — a shown child of the host owned by the player
/// this app began into it (`a_child_of_the_host`) — and the shown-with-a-box rule is what picks the
/// picture rather than a window mpv has made and not yet drawn in. The handle is then kept, because
/// a resize step asks for it on every pointer move and looking it up is a walk of the host's
/// children.
///
/// **The kept handle is only ever a hint.** mpv makes itself a new child when its video output is
/// reinitialised, so a handle that is no longer a window is re-found, and the host and player it was
/// found for are kept beside it: an adopted host, a host made again, and a player begun over the
/// last one are all other answers, and none of them inherits the box of the film before it.
fn the_hosted_mpv_child(host: HWND) -> Option<HWND> {
    // Only a film mpv was begun into this host for has a child of mpv's under it: FFmpeg's is fitted
    // by the road that places the host (see `fit_the_hosted_ffplay_to_the_host`), and a picture, a
    // preview and a park hold no player at all.
    if the_pins_film_is_ffplay() || !the_pinned_player_was_begun_in_the_host() {
        return None;
    }

    let pid = hosted_player_pid()?;
    let kept = HOSTED_MPV_CHILD.load(Ordering::SeqCst);
    if kept != 0
        && HOSTED_MPV_CHILD_HOST.load(Ordering::SeqCst) == host.0 as isize
        && HOSTED_MPV_CHILD_PID.load(Ordering::SeqCst) == pid
    {
        let child = HWND(kept as *mut _);
        if unsafe { IsWindow(child) }.as_bool() {
            return Some(child);
        }
    }

    let child = a_child_of_the_host(host, pid)?;
    HOSTED_MPV_CHILD_PID.store(pid, Ordering::SeqCst);
    HOSTED_MPV_CHILD_HOST.store(host.0 as isize, Ordering::SeqCst);
    HOSTED_MPV_CHILD.store(child.0 as isize, Ordering::SeqCst);
    Some(child)
}

/// Whether a resize of this host would place mpv's child at its box, as a test reads the gate the
/// host's `WM_SIZE` asks before it places anything (see `the_hosted_mpv_child`).
#[cfg(test)]
pub(super) fn a_host_resize_would_place_the_child_for_a_test(host: HWND) -> bool {
    the_hosted_mpv_child(host).is_some()
}

/// The flags a host's `WM_SIZE` fits mpv's child with.
///
/// **A host that is a window of the desktop's asks for the box rather than waiting for it.** Its
/// child is still mpv's, but the message doing the asking is the host's own `WM_SIZE`, which is the
/// window being moved — a synchronous fit there is this app's thread waiting on mpv's, on the step
/// path, which is the hang this road exists to not have. What it costs is the box landing a message
/// later, and the film is clipped by the band rather than the band a hole.
pub(super) fn fit_flags() -> SET_WINDOW_POS_FLAGS {
    let mut flags = SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS;
    // **A film taken off the screen for a resize drag is not put back by the fit that follows its
    // host's own box**, which is the one thing the temporary `RHP_HIDE_FILM_ON_RESIZE` diagnostic
    // asked for: a fit that showed the child is a diagnostic that lasts one step (see
    // `pin_resize_probe`).
    if !pin_resize_probe::the_film_is_hidden() {
        flags |= SWP_SHOWWINDOW;
    }
    if PIN_FILM_HOST_IS_TOP_LEVEL.load(Ordering::SeqCst) {
        flags |= SWP_ASYNCWINDOWPOS;
    }
    flags
}

/// The box the host is placed at, from the band the pin says its picture stands in.
///
/// A child host is placed in the pin window's own client coordinates, and a host that is a window of
/// the desktop's is placed at the band itself — the band *is* a screen region (see
/// `PinnedPreview::content`), so the two roads differ in nothing but which coordinates the number is
/// in.
fn the_film_host_box(content: ScreenRegion) -> (i32, i32, i32, i32) {
    if PIN_FILM_HOST_IS_TOP_LEVEL.load(Ordering::SeqCst) {
        (
            content.0,
            content.1,
            (content.2 - content.0).max(1),
            (content.3 - content.1).max(1),
        )
    } else {
        band_in_the_window(content)
    }
}

/// Make the pin's chrome window owned by its film host, or give it back to the Shell, and take the
/// taskbar button and the Alt+Tab entry with it.
///
/// **Who Windows lists and activates is asked of a window once, when it is made, and ownership is the
/// answer that can still be written later.** An owned window stands above its owner, hides with it,
/// and is out of Alt+Tab and the taskbar — which is the whole of what the new road needs from the OS,
/// and the whole of what the pin owes the host for holding its film (see
/// `the_pin_film_host_is_a_window_of_the_desktop`). `WS_EX_APPWINDOW` is what asks for a taskbar
/// button regardless of ownership, so it goes with the road that owns the pin and comes back with the
/// one that does not (see `pin_real_window::real_pin_style`).
///
/// Both are same-process calls into a window of this app's own: there is no cross-process send here,
/// and nothing of this app's is locked across either.
unsafe fn own_the_pin_window(owned: bool) {
    let Some(pin) = pin_real_window::real_pin_window() else {
        return;
    };
    let owner = if owned { film_host() } else { None }.map_or(0, |owner| owner.0 as isize);
    if unsafe { GetWindowLongPtrW(pin, GWLP_HWNDPARENT) } != owner {
        unsafe { SetWindowLongPtrW(pin, GWLP_HWNDPARENT, owner) };
    }
    unsafe { pin_real_window::set_the_pin_window_appwindow(pin, !owned) };

    // **The chrome carries no band of its own**, and a window given an owner or given back keeps
    // whatever band it wore before (measured): a bit left on it by the pin before this one, or the
    // band it held while its last host owned it, is a pair that disagrees about where it stands (see
    // `the_pin_window_that_holds_the_band`). A host that owns the chrome bands it with itself, so what
    // is written here is the band a window of its own belongs in, and every placement after this
    // asks the setting again (see `band_place`).
    let band = if owned { HWND_NOTOPMOST } else { band_place() };
    let _ = unsafe {
        SetWindowPos(
            pin,
            band,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        )
    };
}

/// Put the film mpv draws at the host's client box, from the message that gave the host its own.
///
/// **This is what makes a hosted mpv film follow a resize instead of trailing it.** FFmpeg's child is
/// fitted in the same call that places the host (see `fit_the_hosted_ffplay_to_the_host`); mpv's is
/// sized by mpv's own thread, and at the pace of a fast hand it arrives several sizes behind the
/// band it stands in.
///
/// **The call blocks on purpose.** `SetWindowPos` on a window of another process's is a `SendMessage`
/// to the thread that owns it and returns once that thread has taken the box, which is the whole
/// point: the child is the band's size *before* the pin's own surface is composited over it, where a
/// box applied later (`SWP_ASYNCWINDOWPOS`) would be the defect still there for a frame. Nothing of
/// this app's is locked across it — the pin's state is read and given up in `band_in_the_window`
/// before the host is placed — so what is waited on is mpv pumping its own queue. How long it took
/// is written to the pin trace, so a step that stalls is a number to report rather than a guess.
fn fit_the_hosted_mpv_to_the_host(host: HWND) -> bool {
    let Some(child) = the_hosted_mpv_child(host) else {
        return false;
    };

    let mut client = RECT::default();
    if unsafe { GetClientRect(host, &mut client) }.is_err() {
        return false;
    }
    let (width, height) = (client.right.max(1), client.bottom.max(1));
    let flags = fit_flags();
    let started = Instant::now();
    let placed =
        unsafe { SetWindowPos(child, HWND::default(), 0, 0, width, height, flags) }.is_ok();

    note_pin_trace!(
        "hosted mpv fit host {host:?} child {child:?} {width}x{height} {:.3}ms placed {}",
        started.elapsed().as_secs_f64() * 1000.0,
        placed as u8,
    );
    // What the temporary `RHP_RESIZE_TRACE` diagnostic writes down, and the box a later tick looks for
    // (see `pin_resize_probe`).
    pin_resize_probe::note_the_child_fit(host, child, (width, height), flags, started);
    placed
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
        note_pin_trace!("hosted drawing pid {pid} host none false");
        return false;
    };
    let drawing =
        pid != 0 && hosted_player_pid() == Some(pid) && a_child_of_the_host(host, pid).is_some();
    note_pin_trace!(
        "hosted drawing pid {pid} host {host:?} drawing {}",
        drawing as u8
    );
    drawing
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
///
/// **The row a resize pushed the picture down by is added to it**, because the window was moved up
/// by that row (see `the_slack_a_resize_paints_with`): a child is placed in the client's own
/// coordinates and so rides that move, and adding the row here is what leaves the band on the box
/// the pin's picture landed on instead of bobbing with the slack.
pub(super) fn band_in_the_window(content: ScreenRegion) -> (i32, i32, i32, i32) {
    let (overlay, caption) = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| (pin.overlay, pin.caption)))
        .unwrap_or((true, 0));

    let width = (content.2 - content.0).max(1);
    let height = (content.3 - content.1).max(1);
    let row = paint::the_row_a_resize_paints_the_picture_down();

    (
        0,
        (if overlay { 0 } else { caption }) + row,
        width,
        height,
    )
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

    if !the_film_is_shown_in_a_host(&path, None) {
        take_down_the_pin_film_host();
        return;
    }

    // Written before any road that asks the host a question under the pin's own lock can run (see
    // `the_pins_film_is_ffplay`).
    note_the_pins_film_route(&path);

    let Some(content) = pinned_content() else {
        return;
    };

    // **The road's level is the film's, and a host of the other level is not this host.** A file
    // stepped between mpv and FFmpeg keeps the host on the road that draws into it, and the two
    // levels cannot be reached through one window: "mpv is drawn in a window of the desktop's"
    // answers true for one film and false for the next, so a host already made for the other road
    // is taken down and made again (see `the_pin_film_host_is_a_window_of_the_desktop`).
    let a_window_of_the_desktop = the_pin_film_host_is_a_window_of_the_desktop(&path);
    if film_host().is_some()
        && PIN_FILM_HOST_IS_TOP_LEVEL.load(Ordering::SeqCst) != a_window_of_the_desktop
    {
        unsafe { take_down_the_pin_film_host() };
    }

    if let Some(host) = adopted.filter(|host| unsafe { IsWindow(*host) }.as_bool()) {
        let pid = hover_film_host::hover_film_host_player().unwrap_or(0);
        note_pin_trace!(
            "host adopted {:?} pid {} child {:?} client {:?} desktop {}",
            host,
            pid,
            a_child_of_the_host(host, pid),
            a_client_box(host),
            a_window_of_the_desktop as u8,
        );
        if a_window_of_the_desktop {
            // **The hover's host is already the window this road wants**, so it is the pin's host
            // rather than a window moved into one: the player keeps drawing into the same child and
            // not a frame of the film leaves the screen (see `promote_the_hover_film_host`).
            unsafe { promote_the_hover_film_host(host, content) };
        } else {
            unsafe { adopt_the_hover_film_host(host, content) };
        }
        return;
    }

    let host = match film_host() {
        Some(host) => host,
        None => {
            let Some(parent) = pin_real_window::real_pin_window() else {
                return;
            };
            let Some(host) = create(parent, a_window_of_the_desktop) else {
                return;
            };
            note_pin_trace!(
                "host created {:?} desktop {} parent {:?} pid {} child {:?} client {:?}",
                host,
                a_window_of_the_desktop as u8,
                parent,
                VIDEO_PID.load(Ordering::SeqCst),
                a_child_of_the_host(host, VIDEO_PID.load(Ordering::SeqCst)),
                a_client_box(host),
            );
            host
        }
    };

    // **The pin's chrome is the host's to own on this road**, and the Shell's to list on the other
    // (see `own_the_pin_window`). Written here because this is the road that knows both windows.
    unsafe { own_the_pin_window(PIN_FILM_HOST_IS_TOP_LEVEL.load(Ordering::SeqCst)) };

    let (x, y, width, height) = the_film_host_box(content);
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

/// Take a hover's already-running top-level host up as the pin's own: **the same window, the same
/// player, the same child, and not one frame of the film off the screen.**
///
/// **Why it can be done at all.** A hover's film is given a host only so that it can be handed to a
/// pin without a second launch, and that host is made the one shape a pin's host has to be: a
/// top-level, non-layered window of this app's — `WS_POPUP | WS_CLIPCHILDREN`, the picture a child
/// of mpv's inside it (see `hover_film_host`). What a hover's host is *for a hover* is the rest of
/// its style — the topmost band nothing may cover, transparent to the press, out of the Shell's
/// lists — and the pin's is none of that (see `create`).
///
/// **Why it is not the re-parent the child road does.** A `WS_EX_LAYERED` chrome cannot clip or
/// synchronise a hardware swapchain child, which is the whole reason the pin's host is a window of
/// the desktop's on this road (see `the_pin_film_host_is_a_window_of_the_desktop`). Moving the
/// hover's host into the pin's window would undo exactly that; the host takes the pin's chrome as
/// its own window instead (see `own_the_pin_window`).
///
/// **Nothing is begun, frozen or read back**, so there is no pause between the hover and the pin and
/// no second decode of the file: what changes is a style, an owner, a band and a box, and what does
/// not is the player inside it.
unsafe fn promote_the_hover_film_host(host: HWND, content: ScreenRegion) {
    let pid = hover_film_host::hover_film_host_player().unwrap_or(0);

    // The pin's host from here, written before the roads that ask the host questions through the
    // pin's own lock can run: the player the hover began is the pin's own, and the frame it is
    // drawing is already on the screen — the hover had proved it.
    PIN_FILM_HOST.store(host.0 as isize, Ordering::SeqCst);
    PIN_FILM_HOST_IS_TOP_LEVEL.store(true, Ordering::SeqCst);
    PIN_FILM_HOST_PLAYER.store(pid, Ordering::SeqCst);
    PRESENTED_PID.store(pid, Ordering::SeqCst);
    VIDEO_HWND.store(0, Ordering::SeqCst);
    unsafe { pin_taskbar::set_pin_taskbar_title(host) };

    // **The hover's band and the press that falls through it go, and the pin's bits come on**: a host
    // the Shell lists wears neither `WS_EX_TOOLWINDOW` nor `WS_EX_NOACTIVATE`, and the app-window bit
    // is what gives the pair its button (see `create` and `own_the_pin_window`).
    let ex = unsafe { GetWindowLongPtrW(host, GWL_EXSTYLE) };
    let ex = (ex
        & !(WS_EX_TOPMOST.0 as isize)
        & !(WS_EX_TRANSPARENT.0 as isize)
        & !(WS_EX_TOOLWINDOW.0 as isize)
        & !(WS_EX_NOACTIVATE.0 as isize))
        | WS_EX_APPWINDOW.0 as isize;
    let _ = unsafe { SetWindowLongPtrW(host, GWL_EXSTYLE, ex) };
    // A style written behind a window's back is not one it presents with until it is asked to read it,
    // and the band it was in goes with the same call.
    let _ = unsafe {
        SetWindowPos(
            host,
            HWND_NOTOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_FRAMECHANGED,
        )
    };

    // The chrome is the host's own window from here, and the pair is put in the band the setting
    // names in one call — the host taking it and the chrome following it (see
    // `place_the_pin_in_the_band`). The host keeps the picture: it was the hover's window on the
    // screen and it is the pin's window on the screen, moved to the band the pin was painted around.
    unsafe { own_the_pin_window(true) };
    unsafe { place_the_pin_film_host(content) };
    let pin = pin_real_window::real_pin_window().unwrap_or(host);
    unsafe { place_the_pin_in_the_band(pin, band_place()) };

    // The hover's handle is given up rather than left to destroy the pin's host (see
    // `forget_the_hover_film_host_that_was_handed_over`), and the strips are put back over the
    // window that is the film's now.
    hover_film_host::forget_the_hover_film_host_that_was_handed_over(host);
    unsafe { pin_chrome_strips::raise_the_chrome_strips_over_the_host() };

    note_pin_trace!(
        "host promoted {:?} pid {} child {:?} client {:?} band {:?}",
        host,
        pid,
        a_child_of_the_host(host, pid),
        a_client_box(host),
        band_place(),
    );
}

/// Put the host at the pin's band, without showing it: the band-change road, where the window is
/// either already up or is the park's to hand back (see `ensure_pinned_sibling_box`).
pub(super) unsafe fn place_the_pin_film_host(content: ScreenRegion) {
    let Some(host) = film_host() else {
        return;
    };

    let (x, y, width, height) = the_film_host_box(content);
    // The temporary `RHP_RESIZE_TRACE` diagnostic's own clock for this placement, taken only where a
    // resize step is being measured (see `pin_resize_probe`).
    let started = pin_resize_probe::a_step_started();
    let _ = SetWindowPos(
        host,
        HWND::default(),
        x,
        y,
        width,
        height,
        SWP_NOZORDER | SWP_NOACTIVATE,
    );
    pin_resize_probe::note_the_host_placement(started);
    unsafe { fit_the_hosted_ffplay_to_the_host(host) };
    pin_chrome_strips::place_the_chrome_strips();
}

/// Hide or show the host for a park, at the box it already has: a park hides the window behind the
/// band, and the pin paints its held frame in its place (see `park_pinned_player_inner`). The
/// chrome stays where the chrome state puts it: the strips are the pin's window's rather than the
/// host's, and a park hides the picture rather than the chrome (see `pin_chrome_strips`).
pub(super) unsafe fn set_the_pin_film_host_visible(visible: bool) {
    if let Some(host) = film_host() {
        // **A host the Shell has minimized is the Shell's to bring back.** Showing a minimized
        // window restores it, so a tick that showed the host it is watching would undo the user's own
        // minimize of it every tick — and the button that minimized it is the host's own on this road
        // (see `own_the_pin_window`). A host the app hid for a park is not iconic, and is shown here as
        // it always was.
        if visible && unsafe { IsIconic(host) }.as_bool() {
            return;
        }
        note_pin_trace!("host visible {} {:?}", visible as u8, host);
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
    // **The pin's chrome is given back to the Shell before the host that owns it goes.** Destroying
    // a window destroys the windows it owns, and the pin's window is kept between pins — so a
    // teardown that destroyed the owner first would take the pin's own window with it (see
    // `own_the_pin_window`).
    unsafe { own_the_pin_window(false) };
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
    PIN_FILM_HOST_IS_TOP_LEVEL.store(false, Ordering::SeqCst);
    PRESENTED_PID.store(0, Ordering::SeqCst);
    forget_the_hosted_mpv_child();
    // The temporary `RHP_HIDE_FILM_ON_RESIZE` diagnostic's film goes with the host that held it, and a
    // flag left standing is the next film fitted without a show (see `pin_resize_probe`).
    pin_resize_probe::forget_the_hidden_film();
    clear_the_move_cover();
    // A cover left over from a move must not open the next pin or park on this film's frame. A
    // park's own frame is its to give up, so it is left standing for the park.
    if !pin_player_is_parked() {
        forget_video_frame();
    }
}

/// A client box as a trace reads it, or nothing where the window has already gone.
fn a_client_box(hwnd: HWND) -> Option<(i32, i32)> {
    let mut rect = RECT::default();
    (unsafe { GetClientRect(hwnd, &mut rect) }.is_ok()).then_some((rect.right, rect.bottom))
}

unsafe fn create(parent: HWND, a_window_of_the_desktop: bool) -> Option<HWND> {
    FILM_HOST_CLASS_REGISTERED.call_once(|| {
        let hinstance = GetModuleHandleW(None).expect("the module's own handle");
        let class = WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(the_default_answer),
            hInstance: hinstance.into(),
            // **No background brush, because the host paints its own erase** (see
            // `paint_the_host_behind_the_film`): the class's brush is what erased a grown band
            // black, and the one erase road is the window procedure's.
            hbrBackground: HBRUSH(GetStockObject(NULL_BRUSH).0),
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
    //
    // **`WS_CLIPCHILDREN` because the picture is this window's child and this window paints its own
    // erase.** The band is filled with the frame a drag holds (or black), and an erase that is not
    // clipped to the area outside the player's window is that fill drawn over the film at every show
    // and every resize. The pin's own window is clipped for the same reason, one level up (see
    // `pin_real_window::real_pin_style`).
    let host = CreateWindowExW(
        if a_window_of_the_desktop {
            WS_EX_APPWINDOW
        } else {
            WINDOW_EX_STYLE::default()
        },
        FILM_HOST_CLASS,
        w!("film host"),
        if a_window_of_the_desktop {
            WS_POPUP | WS_CLIPCHILDREN
        } else {
            WS_CHILD | WS_CLIPCHILDREN
        },
        0,
        0,
        1,
        1,
        // A window of the desktop's has no parent — it is the one that has an *owner*, and the
        // owner is the pin's chrome (see `own_the_pin_window`). Nothing here is a `WS_CHILD`.
        if a_window_of_the_desktop {
            HWND::default()
        } else {
            parent
        },
        None,
        hinstance,
        None,
    )
    .ok()?;

    // Disabled before it is ever shown: mpv's child is the picture, and the press that lands on
    // the picture has to reach the pin (see the module doc).
    let _ = EnableWindow(host, false);
    PIN_FILM_HOST.store(host.0 as isize, Ordering::SeqCst);
    PIN_FILM_HOST_IS_TOP_LEVEL.store(a_window_of_the_desktop, Ordering::SeqCst);
    // **A host that is a window of the desktop's is the one the Shell lists**, so the button the pin
    // used to carry names the file the pin is of (see `pin_taskbar::set_pin_taskbar_title`).
    if a_window_of_the_desktop {
        unsafe { pin_taskbar::set_pin_taskbar_title(host) };
    }
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

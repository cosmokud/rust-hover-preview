//! The window a pinned film is drawn in: a window of the desktop's that owns the pin's chrome and
//! shows the film through the Shell's thumbnail of the player's own window.
//!
//! **Both engines are on one road.** mpv is launched with no `--wid` and makes a window of its own,
//! and FFmpeg's player is an SDL program with no embed switch at all, so it makes one too. The host
//! asks the Shell for a thumbnail of that window (`DwmRegisterThumbnail`) and draws it into its own
//! client area (see `pin_film_thumbnail`). The host is then a **window of the desktop's that owns
//! the pin's chrome**: Windows keeps the chrome above the film, hides it with the host, keeps it out
//! of Alt+Tab, and moves neither for this app (see `the_pins_chrome_is_owned` and
//! `own_the_pin_window`). A resize step moves this app's own window and re-points one rectangle;
//! nothing of the player's is asked for a box. **Why that is the only road, and what a resize may
//! move, is `Pin mode → The film road` in `ARCHITECTURE.md`.**
//!
//! **Why it is not the pin's window itself.** The pin's window is the chrome and the transport as
//! well as the band, and the film would be drawn over the caption. So the picture gets a window of
//! its own at the band's box, and the pin keeps every other pixel (see `the_film_host_box`).
//!
//! **Why it is disabled.** A press on the picture is a press on the pin: the pin is dragged by its
//! film, and the band's alpha-1 claim is already hit-testable (see `fill_band_hit_testable`). A
//! disabled window is skipped by `WindowFromPoint`, so the press arrives at the pin rather than at
//! the host.
//!
//! The host outlives the player once begun into it: a seek is mpv's own IPC and a step begins
//! another player behind the same host. It goes down with the pin (see
//! `take_down_the_pin_film_host`, called from the state a road out of a pin clears).

use std::sync::atomic::{AtomicIsize, AtomicU32, Ordering};
use std::sync::Once;

use windows::core::w;
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, RECT, WPARAM};
use windows::Win32::Graphics::Gdi::{
    BeginPaint, EndPaint, FillRect, GetStockObject, SetBrushOrgEx, SetStretchBltMode, StretchBlt,
    BLACK_BRUSH, HBRUSH, HALFTONE, NULL_BRUSH, SRCCOPY,
};
use windows::Win32::Graphics::Gdi::PAINTSTRUCT;
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::EnableWindow;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, GetClientRect, GetWindowLongPtrW,
    RegisterClassExW, SetWindowLongPtrW, SetWindowPos, ShowWindow, GWLP_HWNDPARENT, GWL_EXSTYLE,
    HWND_NOTOPMOST, HWND_TOPMOST, SC_CLOSE, SWP_NOACTIVATE, SWP_NOCOPYBITS, SWP_NOMOVE, SWP_NOSIZE,
    SWP_NOZORDER, SW_HIDE, SW_SHOWNOACTIVATE, WM_CLOSE, WM_ERASEBKGND, WM_PAINT, WM_SYSCOMMAND,
    WNDCLASSEXW, WS_CLIPCHILDREN, WS_EX_APPWINDOW, WS_POPUP,
};

use super::pin_trace::note_pin_trace;
use super::*;

/// The class a film's host is made from. A class of its own rather than the pin window's: what a
/// `WS_CHILD` host answers to is a window that draws nothing and takes nothing, and the pin's
/// window procedure is the pin's (see `pin_real_window`).
const FILM_HOST_CLASS: windows::core::PCWSTR = w!("RustHoverPreviewPinFilmHost");

/// The default answer is most of what the host does: it draws nothing of the film's own and no
/// press lands on it. On `WM_ERASEBKGND` it fills the band the picture leaves (see
/// `paint_the_host_behind_the_film`). This bindings version exposes `DefWindowProcW` as a
/// Rust-callable function, so the class still needs a system-ABI shim around it.
unsafe extern "system" fn the_default_answer(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    // **The band behind the picture is painted here rather than erased by the class**, and the class
    // wears `NULL_BRUSH` so that this is the only hand that draws it (see
    // `paint_the_host_behind_the_film` and `create`). `1` is "handled": it is what keeps
    // `DefWindowProcW` from erasing the same pixels a second time.
    if msg == WM_ERASEBKGND {
        unsafe { paint_the_host_behind_the_film(hwnd, HDC(wparam.0 as *mut _)) };
        return LRESULT(1);
    }

    // A cover this app painted stands where the film was, and this paint is its own (see
    // `answer_the_paint_of_a_cover`).
    if let Some(answer) = unsafe { answer_the_paint_of_a_cover(hwnd, msg) } {
        return answer;
    }

    // **A host that is the Shell's window is the Shell's close too.** Its taskbar button's Close
    // arrives here as a `WM_CLOSE` or an `SC_CLOSE`, and taken it would destroy the window the pin's
    // chrome is owned by — the pin with it, under a state that still believes a pin is up. What a
    // close means for the pin is the pin's own road's answer (see `window_proc`).
    if msg == WM_CLOSE || (msg == WM_SYSCOMMAND && (wparam.0 & 0xFFF0) as u32 == SC_CLOSE) {
        request_pin_end();
        return LRESULT(0);
    }

    unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
}

/// Whether a resize drag is in flight: the one gesture the thumbnail road must never place a source
/// under, so that a film whose size arrives late is not sized while the hand is on an edge (see
/// `pin_film_thumbnail::place_the_source_when_the_player_answers`).
///
/// **Asked without waiting for the pin's own lock**, because the asker may run inside a call that
/// holds it — the tick's own pass — and a lock that cannot be taken here is answered with the road
/// as it was rather than waited on (see `pin_state_try`).
pub(super) fn a_resize_drag_is_live() -> bool {
    pin_window::pin_state_try()
        .and_then(|pinned| pinned.pin().map(|pin| pin.dragging))
        .flatten()
        .is_some_and(|drag| matches!(drag.action, PinDragAction::Resize(_)))
}

/// **A cover this app painted is the host's picture until the replacement's registration takes the
/// box** (see `put_the_cover_in`), so the two messages that would otherwise draw the host's black
/// over it answer differently: the erase fills nothing, and the paint draws the retained frame
/// again, re-cropped to the client this paint is for.
///
/// It is one function rather than two window procedures' worth, because a pin's host can be a
/// hover's promoted into it (see `promote_the_hover_film_host`), and the cover is owed by whichever
/// class the host was made from.
pub(super) unsafe fn answer_the_paint_of_a_cover(hwnd: HWND, msg: u32) -> Option<LRESULT> {
    if msg != WM_PAINT || !the_cover_stands() {
        return None;
    }
    let mut paint = PAINTSTRUCT::default();
    // Safety: the context is the one Windows hands this paint with, and it is given back below.
    let dc = unsafe { BeginPaint(hwnd, &mut paint) };
    if !dc.is_invalid() {
        unsafe { paint_the_cover(hwnd, dc) };
    }
    unsafe { let _ = EndPaint(hwnd, &paint); };
    Some(LRESULT(0))
}

/// Paint the host's whole client black: the film is the Shell's thumbnail (see `pin_film_thumbnail`)
/// or FFmpeg's child, which fills the client box itself.
///
/// **The host's own erase is what closes the hole a film fitted into its box once showed as three
/// bands**: the Shell composes the thumbnail *over* this app's content, so a client painted in the
/// bars alone is a pixel nobody drew — and one of those reads as the desktop seen through the pin.
/// The picture fills its destination, which is the whole client, so nothing of this fill shows
/// through once the Shell has composed; what it covers is the client the Shell has not composed yet,
/// or one whose source is not drawing (see `pin_film_thumbnail::settle_the_registration`). Both hosts
/// a pin can show its film through are erased through here — the pin's own class, and the hover's a
/// pin took up (see `hover_film_host`).
pub(super) unsafe fn paint_the_host_behind_the_film(host: HWND, dc: HDC) {
    // **Nothing paints the host's black over a cover it owes the screen**, and the count above says
    // the same thing: an erase that fills here is the black a still must never show (see
    // `the_cover_stands`).
    if the_cover_stands() {
        return;
    }
    #[cfg(test)]
    HOST_ERASES.fetch_add(1, Ordering::SeqCst);
    let mut client = RECT::default();
    if unsafe { GetClientRect(host, &mut client) }.is_err() {
        return;
    }

    // SAFETY: the brush is a stock object this app neither made nor owns, and the rectangle is the
    // client box read above; GDI reports rather than faults for a device context that has gone.
    let brush = HBRUSH(unsafe { GetStockObject(BLACK_BRUSH) }.0);
    unsafe { FillRect(dc, &client, brush) };
}

/// The frame a still of this app's painted into the host, kept so the host can go on standing on it
/// for the whole of a relaunch's wait: BGRA bytes at the size FFmpeg answered (see
/// `pin_film_thumbnail::stand_the_still`).
///
/// It is bytes rather than a `DibSurface` because a static has to be `Send` and a device context is
/// not; the surface a paint needs is made, filled and given back inside that paint.
struct CoverFrame {
    pixels: Vec<u8>,
    width: u32,
    height: u32,
}

/// The cover this app painted, or nothing while the host owes no picture of its own.
static COVER: Mutex<Option<CoverFrame>> = Mutex::new(None);

/// Whether a cover this app painted is what the host owes the screen (see `COVER`).
pub(super) fn the_cover_stands() -> bool {
    COVER.lock().is_ok_and(|held| held.is_some())
}

/// Keep a landed still as the host's cover until the picture changes hands (see
/// `put_the_cover_away`).
///
/// **The frame goes with the host it was decoded for**: a paint that lands after that host has gone
/// is nobody's picture, and a cover left standing would keep the next pin's host from erasing.
///
/// A frame that is not the size it claims is refused rather than kept: what arrives here came out of
/// a child process, and a short frame is a smear rather than a picture.
pub(super) fn put_the_cover_in(host: HWND, pixels: Vec<u8>, width: u32, height: u32) -> bool {
    if film_host() != Some(host)
        || width == 0
        || height == 0
        || pixels.len() < width as usize * height as usize * 4
    {
        return false;
    }
    match COVER.lock() {
        Ok(mut held) => {
            *held = Some(CoverFrame {
                pixels,
                width,
                height,
            });
            true
        }
        Err(_) => false,
    }
}

/// Give the cover up: the replacement's registration has taken the box, so the Shell draws the
/// picture and the host owes nothing of its own (see `pin_film_thumbnail::hand_the_picture_over`).
pub(super) fn put_the_cover_away() {
    if let Ok(mut held) = COVER.lock() {
        *held = None;
    }
}

/// Draw the cover into `dc`, **cover-cropped to the host's client as it is now**: the same centred
/// crop the registration wears, so a resize or an invalidation in the middle of a gap re-crops the
/// picture instead of showing the host's black (see `pin_film_thumbnail::the_pictures_crop`).
///
/// Safety: `dc` is a context for `host`, as the erase and the paint are handed theirs.
pub(super) unsafe fn paint_the_cover(host: HWND, dc: HDC) -> Option<(i32, i32, i32, i32)> {
    let mut client = RECT::default();
    if unsafe { GetClientRect(host, &mut client) }.is_err() {
        return None;
    }
    let (width, height) = (client.right - client.left, client.bottom - client.top);
    if width <= 0 || height <= 0 {
        return None;
    }
    let held = COVER.lock().ok()?;
    let cover = held.as_ref()?;
    let (left, top, crop_width, crop_height) = super::pin_film_thumbnail::the_pictures_crop(
        (cover.width as i32, cover.height as i32),
        (width, height),
        Some(width as f64 / height as f64),
    )
    .0;
    let surface = DibSurface::create(cover.width, cover.height)?;
    let length = cover.width as usize * cover.height as usize * 4;
    // Safety: the surface is this call's own of exactly these pixels, and the cover was checked for
    // them when it was put in.
    unsafe { std::ptr::copy_nonoverlapping(cover.pixels.as_ptr(), surface.bits(), length) };
    // Safety: the two contexts are this call's and the host's, and one scaling copy is asked of them.
    let painted = unsafe {
        let _ = SetStretchBltMode(dc, HALFTONE);
        // The brush origin travels with the mode: halftone measures the destination from it.
        let _ = SetBrushOrgEx(dc, 0, 0, None);
        StretchBlt(
            dc,
            0,
            0,
            width,
            height,
            surface.dc,
            left,
            top,
            crop_width,
            crop_height,
            SRCCOPY,
        )
        .as_bool()
    };
    #[cfg(test)]
    if let Ok(mut held) = THE_COVERS_CROP.lock() {
        *held = Some((left, top, crop_width, crop_height));
    }
    painted.then_some((left, top, crop_width, crop_height))
}

/// The crop the cover was last drawn with, for a test that reads the road rather than the pixels
/// (see `paint_the_cover`).
#[cfg(test)]
static THE_COVERS_CROP: Mutex<Option<(i32, i32, i32, i32)>> = Mutex::new(None);

/// The crop the cover was last drawn with (see `THE_COVERS_CROP`).
#[cfg(test)]
pub(super) fn the_covers_crop() -> Option<(i32, i32, i32, i32)> {
    THE_COVERS_CROP.lock().ok().and_then(|held| *held)
}

/// Registered once per run: the class a film's host is made from (see `create`).
static FILM_HOST_CLASS_REGISTERED: Once = Once::new();

/// How many erases this run has painted into a film host: the seam that says **nothing paints over a
/// still a relaunch left standing** (see `paint_the_host_behind_the_film` and
/// `pin_film_thumbnail::stand_the_outgoing_picture_from_its_dc`).
#[cfg(test)]
static HOST_ERASES: AtomicU32 = AtomicU32::new(0);

/// The erases this run has painted into a film host (see `HOST_ERASES`).
#[cfg(test)]
pub(super) fn host_erases() -> u32 {
    HOST_ERASES.load(Ordering::SeqCst)
}

/// The host of the pin that is up, and the player this app has attached to it.
static PIN_FILM_HOST: AtomicIsize = AtomicIsize::new(0);
static PIN_FILM_HOST_PLAYER: AtomicU32 = AtomicU32::new(0);
/// The hosted player whose frame this app has proved on screen, so the pipe read that proves it is
/// asked once and not on every tick (see `the_hosted_frame_is_up`).
static PRESENTED_PID: AtomicU32 = AtomicU32::new(0);

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

/// The host an FFmpeg film at this pin's band is kept for, where this is the pin's own window and
/// the pin stands in it: the film is drawn through it by the thumbnail, exactly as mpv's is (see
/// `pin_film_thumbnail`). `None` is every other launch, where the player keeps its own window of the
/// desktop's over the preview.
///
/// **Whether the film is drawn through a host is a fact of the launch rather than of the host's
/// existence**: a player begun while the host is still being made is one the monitor adopts as soon
/// as the host is there (see `host_for_a_hosted_player`), and nothing about the command line
/// depends on which came first.
pub(super) fn the_pins_film_is_kept_in_its_host(path: &Path, kind: Option<MediaType>) -> bool {
    the_film_is_shown_in_a_host(path, kind) && video_route(path) == VideoRoute::Ffplay
}

/// Whether the pin's chrome is owned by its film host, and so is not the window Windows lists and
/// gives a taskbar button to: the host is that window on this road (see `own_the_pin_window`).
pub(super) fn the_pins_chrome_is_owned() -> bool {
    film_host().is_some()
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

/// Whether the pin's host holds an FFmpeg player whose window has not been found yet: the take-up
/// notes the player it is keeping so that the monitor adopts its window as the source the moment it
/// exists (see `host_for_a_hosted_player`).
pub(super) fn adopt_a_stranded_ffplay_into_the_pin_host(pid: u32, path: &Path) -> bool {
    if pid == 0 || video_route(path) != VideoRoute::Ffplay {
        return false;
    }
    note_hosted_player(pid);
    true
}

/// Whether this file is shown in a host this app owns: a film pinned in a window of the pin's own,
/// whichever engine plays it. Asked of the route rather than of the player's process because it is
/// the *launch* that decides where the picture is drawn (see `start_video_playback`).
pub(super) fn the_film_is_shown_in_a_host(path: &Path, kind: Option<MediaType>) -> bool {
    // A sound's player is FFmpeg's too, and a sound's route would match the film routes below: the
    // media on screen has to be a film before any host road is asked (see `MediaType::is_audio`).
    pin_is_its_own_window()
        && kind.or_else(current_media_type) == Some(MediaType::Video)
        && matches!(video_route(path), VideoRoute::Mpv | VideoRoute::Ffplay)
}

/// The box the host is placed at, from the band the pin says its picture stands in. The band *is* a
/// screen region (see `PinnedPreview::content`), and a host of the desktop's is placed at it.
fn the_film_host_box(content: ScreenRegion) -> (i32, i32, i32, i32) {
    (
        content.0,
        content.1,
        (content.2 - content.0).max(1),
        (content.3 - content.1).max(1),
    )
}

/// Make the pin's chrome window owned by its film host, or give it back to the Shell, and take the
/// taskbar button and the Alt+Tab entry with it.
///
/// **Who Windows lists and activates is asked of a window once, when it is made, and ownership is the
/// answer that can still be written later.** An owned window stands above its owner, hides with it,
/// and is out of Alt+Tab and the taskbar — which is the whole of what the film road needs from the OS,
/// and the whole of what the pin owes the host for holding its film (see
/// `the_pins_chrome_is_owned`). `WS_EX_APPWINDOW` is what asks for a taskbar
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

/// Whether the tick owes the player's window the front: a film a hover is following, or a pin's
/// own hosted film. The pin's half is asked of the host rather than of the hover's bookkeeping,
/// because a hover's `Hide` can empty that bookkeeping while the pin still owns the window (see
/// `a_pin_owns_the_window_a_hide_would_take_down`).
pub(super) fn the_video_window_is_owed_the_front(current_video_path: Option<&Path>) -> bool {
    current_video_path.is_some() || (pinned() && pin_is_its_own_window() && film_host().is_some())
}

/// Whether the player this app is holding is the one it began into the host — the question every
/// road that used to find a top-level mpv window by pid asks instead (see
/// `ensure_video_window_topmost`).
///
/// It says the pairing and not that the film is drawing: a player begun into the host has not put
/// a child of its own there for a moment or more, and that wait is `the_host_is_drawing`.
pub(super) fn the_pinned_player_was_begun_in_the_host() -> bool {
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    pid != 0 && film_host().is_some() && hosted_player_pid() == Some(pid)
}

/// How often an unchanged `the_host_is_drawing` answer is written down: the tick asks it hundreds of
/// times a second, and on-change plus this heartbeat is what makes the trace readable (see
/// `trace_the_hosted_drawing`).
const HOSTED_DRAWING_TRACE_EVERY: Duration = Duration::from_secs(1);

/// One `the_host_is_drawing` answer as the trace last wrote it: the player whose source the host is
/// drawing, the host (zero for none), the answer, and when it was written.
type HostedDrawingTrace = (u32, isize, bool, Instant);

/// The last answer `the_host_is_drawing` wrote down (see `trace_the_hosted_drawing`).
static LAST_HOSTED_DRAWING: Lazy<Mutex<Option<HostedDrawingTrace>>> =
    Lazy::new(|| Mutex::new(None));

/// Write the answer of one look at the hosted film's window, on change and at most once a second.
///
/// **The line is asked for by every tick, several times, and used to write some 3 MB a run** — which
/// is a trace nothing can be read out of. A state transition is the interesting part and is never
/// missed; an answer that has not moved is repeated only as the heartbeat that says the tick is
/// still asking it.
fn trace_the_hosted_drawing(pid: u32, host: Option<HWND>, drawing: bool) {
    let Ok(mut last) = LAST_HOSTED_DRAWING.lock() else {
        return;
    };
    let now = Instant::now();
    let host_number = host.map_or(0, |host| host.0 as isize);
    let answer = (pid, host_number, drawing, now);
    if !the_drawing_trace_is_due(*last, answer, now) {
        return;
    }
    *last = Some(answer);
    match host {
        Some(host) => note_pin_trace!(
            "hosted drawing pid {pid} host {host:?} drawing {}",
            drawing as u8
        ),
        None => note_pin_trace!("hosted drawing pid {pid} host none false"),
    }
}

/// Whether an answer is due a line: a change of player, host or answer, and otherwise no more than
/// the heartbeat `HOSTED_DRAWING_TRACE_EVERY` after the last one (see `trace_the_hosted_drawing`).
fn the_drawing_trace_is_due(
    last: Option<HostedDrawingTrace>,
    answer: HostedDrawingTrace,
    now: Instant,
) -> bool {
    last.is_none_or(|(pid, host, drawing, at)| {
        (pid, host, drawing) != (answer.0, answer.1, answer.2)
            || now.duration_since(at) >= HOSTED_DRAWING_TRACE_EVERY
    })
}

/// **Whether the pin's host has a picture in it at all**, whichever player began it: the registration
/// standing in that host, which is the frame on screen (see
/// `pin_film_thumbnail::the_standing_source_process`).
///
/// It is the question every road that would show, raise or order the host asks, and it is asked of the
/// registration rather than of the newest player because of the whole of a relaunch's wait: the
/// replacement is recorded before it has made a window, while the player before it is **covered** and is
/// what the host is drawing (see `pin_playback::relaunch_pinned_player`). Asked of the newest pid
/// instead, a seek's instant is a pin declared empty — the z-order re-assert skipped, the band treated
/// as a hole — with a still standing in it. A host with no registration in it yet is the player's
/// own window, which is all there is while the source is being adopted and the Shell has not been handed
/// it (see `pin_film_thumbnail::adopt_the_source`). **A still this app painted is the one picture a host
/// can have with no registration in it**, and it is not this question's: the host stands where it is and
/// the tick's show has nothing to add to it (see
/// `pin_film_thumbnail::stand_the_outgoing_picture_from_its_dc`).
pub(super) fn the_host_is_drawing() -> bool {
    let Some(host) = film_host() else {
        trace_the_hosted_drawing(0, None, false);
        return false;
    };
    if let Some(standing) = pin_film_thumbnail::the_standing_source_process(host) {
        trace_the_hosted_drawing(standing, Some(host), true);
        return true;
    }
    // A host with no registration in it yet is the pass between a source being adopted and the
    // Shell being handed it, and the player's own window is all there is to show for it.
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    let drawing = the_hosted_player_is_drawing(pid);
    trace_the_hosted_drawing(pid, Some(host), drawing);
    drawing
}

/// Whether **this** player is up as the pin's picture: the source the Shell draws into the host is the
/// window this player made, adopted and still live (see `pin_film_thumbnail::the_source_is_up`).
///
/// **It is asked of the process rather than of the host alone**, because the host outlives the player
/// it holds: a file step begins a new player for the same host, and the old source would answer "a
/// film is drawing" for a film about to be killed. The source records the player it was adopted for,
/// so the pid and the window are one question. What it is *not* is "has the host a picture": a
/// replacement that has made no window yet answers no here while the frozen picture it replaced is
/// still drawing (see `the_host_is_drawing`).
pub(super) fn the_hosted_player_is_drawing(pid: u32) -> bool {
    film_host().is_some()
        && pid != 0
        && hosted_player_pid() == Some(pid)
        && pin_film_thumbnail::the_source_is_up(pid)
}

/// Whether this player has a window on screen: a source the Shell is drawing through a host of this
/// app's, or a window of the desktop's of its own. It is one answer because the settlement roads
/// have one question — has this player arrived — and a film drawn through a host publishes no window
/// of the desktop's to be found by (see `video_window_for`, `settle_video_retirement` and
/// `PinSwapHold::settle`).
pub(super) fn the_players_window_is_up(pid: u32) -> bool {
    video_window_for(pid).is_some() || the_hosted_player_is_drawing(pid)
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
/// source is there (see `ensure_video_window_topmost`).
///
/// `adopted` is the hover's host on the road that takes an already-running film up: its player is
/// the one drawing the hover, so the window is the pin's rather than replaced (see
/// `promote_the_hover_film_host`). Every other road passes `None` and a fresh host is made where
/// there is none.
pub(super) unsafe fn take_up_the_pin_film_host(adopted: Option<HWND>) {
    let Some(path) = pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.path.clone())) else {
        take_down_the_pin_film_host();
        return;
    };

    if !the_film_is_shown_in_a_host(&path, None) {
        take_down_the_pin_film_host();
        return;
    }

    let Some(content) = pinned_content() else {
        return;
    };

    if let Some(host) = adopted.filter(|host| unsafe { IsWindow(*host) }.as_bool()) {
        // **The hover's host is already the window this road wants**, so it is the pin's host rather
        // than a window moved into one: the player keeps drawing into the same source and not a frame
        // of the film leaves the screen (see `promote_the_hover_film_host`).
        unsafe { promote_the_hover_film_host(host, content) };
        return;
    }

    let host = match film_host() {
        Some(host) => host,
        None => {
            let Some(host) = create() else {
                return;
            };
            note_pin_trace!(
                "host created {:?} pid {} source {:?} client {:?}",
                host,
                VIDEO_PID.load(Ordering::SeqCst),
                pin_film_thumbnail::the_source_window(),
                a_client_box(host),
            );
            host
        }
    };

    // **The pin's chrome is the host's to own on this road** (see `own_the_pin_window`). Written here
    // because this is the road that knows both windows.
    unsafe { own_the_pin_window(true) };

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
    // A picture the Shell is already drawing into this host goes with the box it was placed at, even
    // where the host was made and the film's window found in the same pass (see `pin_film_thumbnail`).
    // **The source itself is never asked for a box again**, so this is a destination rectangle and
    // nothing else (see `pin_film_thumbnail::the_band_was_placed`).
    pin_film_thumbnail::the_band_was_placed(host);
    pin_chrome_strips::place_the_chrome_strips();

    // **A hosted player has no window of the desktop's to be found by pid**, so any handle still
    // published for an older player is not this film's picture and must not be read as one (see
    // `the_players_window_is_up`).
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

/// Take a hover's already-running host up as the pin's own: **the same window, the same player, the
/// same source and registration, and not one frame of the film off the screen.**
///
/// **Why it can be done at all.** A hover's film is given a host only so that it can be handed to a
/// pin without a second launch, and that host is already the one shape a pin's host has to be: a
/// top-level, non-layered window of this app's — `WS_POPUP | WS_CLIPCHILDREN`, the film a thumbnail
/// of the player's own window (see `hover_film_host`). What a hover's host is *for a hover* is the
/// rest of its style — the topmost band nothing may cover, transparent to the press, out of the
/// Shell's lists — and the pin's is none of that (see `create`).
///
/// **Why it is not a re-parent.** The host is a window of the desktop's so that the Shell composes it
/// and the thumbnail of the film in it: moving it into the pin's layered window would take that away,
/// so the host takes the pin's chrome as its own window instead (see `own_the_pin_window`).
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
        "host promoted {:?} pid {} source {:?} client {:?} band {:?}",
        host,
        pid,
        pin_film_thumbnail::the_source_window(),
        a_client_box(host),
        band_place(),
    );
}

/// Put the host at the pin's band, without showing it: the band-change road, where the window is
/// either already up (see `ensure_pinned_sibling_box`).
pub(super) unsafe fn place_the_pin_film_host(content: ScreenRegion) {
    let Some(host) = film_host() else {
        return;
    };

    let (x, y, width, height) = the_film_host_box(content);
    // **A host has nothing of this app's to carry across a step**, so its bits are not copied: the
    // picture is the Shell's thumbnail or FFmpeg's child, and the only pixels this app paints in the
    // host are its whole client erased black behind that picture — a fresh surface every step is what
    // keeps a bar of an earlier box from standing black over the picture the box now draws (see
    // `pin_film_thumbnail::settle_the_registration`).
    let flags = SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS;
    // The temporary `RHP_RESIZE_TRACE` diagnostic's own clock for this placement, taken only where a
    // resize step is being measured (see `pin_resize_probe`).
    let started = pin_resize_probe::a_step_started();
    let _ = SetWindowPos(host, HWND::default(), x, y, width, height, flags);
    // **What a resize step costs on the thumbnail road is this one call**: the destination is the
    // host's new client area, and the film's window is not touched at all (see `pin_film_thumbnail`).
    pin_film_thumbnail::the_band_was_placed(host);
    pin_resize_probe::note_the_host_placement(started);
    pin_chrome_strips::place_the_chrome_strips();
}

/// Hide or show the host the film is drawn in.
///
/// **Nothing but the film leaving the pin hides it** (see `take_the_film_off_the_band`), and a hide
/// with a pin standing over a source the host is drawing is written out as the defect it would be:
/// the picture is the Shell's thumbnail of the player's window, and there is no frame of this app's
/// to stand in for it (see `pin_film_thumbnail`). The chrome stays where the chrome state puts it:
/// the strips are the pin's window's rather than the host's, so they are not hidden with it.
///
/// The last show this host was asked for, and the host it was asked for: the tick asks on every
/// pass, so the trace's `host visible` line is written where the answer changes rather than on every
/// ask (938 of a 114 s run's lines were this family and `topmost shown`).
static LAST_HOST_SHOW: Lazy<Mutex<Option<(isize, bool)>>> = Lazy::new(|| Mutex::new(None));

/// Whether a show is one the trace has not written yet: the answer the last pass wrote is the
/// polling line a normal session's trace does not need (see `trace_the_host_show`).
fn the_show_is_new(last: Option<(isize, bool)>, host: isize, visible: bool) -> bool {
    last != Some((host, visible))
}

/// Write the show this host was asked for, where it is not the one the last pass asked for.
fn trace_the_host_show(host: HWND, visible: bool) {
    let Ok(mut last) = LAST_HOST_SHOW.lock() else {
        return;
    };
    if !the_show_is_new(*last, host.0 as isize, visible) {
        return;
    }
    *last = Some((host.0 as isize, visible));
    note_pin_trace!("host visible {} {:?}", visible as u8, host);
}

/// How many times this run has asked the host to hide, for the tests that hold a relaunch to never
/// taking the picture off the screen (see `set_the_pin_film_host_visible`).
#[cfg(test)]
pub(super) static HOST_HIDES_ASKED: AtomicU32 = AtomicU32::new(0);

/// The hides this run has asked for.
#[cfg(test)]
pub(super) fn host_hides_asked() -> u32 {
    HOST_HIDES_ASKED.load(Ordering::SeqCst)
}

pub(super) unsafe fn set_the_pin_film_host_visible(visible: bool) {
    #[cfg(test)]
    if !visible {
        HOST_HIDES_ASKED.fetch_add(1, Ordering::SeqCst);
    }
    if let Some(host) = film_host() {
        // **A host the Shell has minimized is the Shell's to bring back.** Showing a minimized
        // window restores it, so a tick that showed the host it is watching would undo the user's own
        // minimize of it every tick — and the button that minimized it is the host's own on this road
        // (see `own_the_pin_window`).
        if visible && unsafe { IsIconic(host) }.as_bool() {
            return;
        }
        // **A picture hidden under a pin is the blank this road exists not to make.** The one road
        // that may hide the host is the film being taken off the pin, which lets the registration go
        // before it asks for the hide; anything else is a relaunch's step taken over a standing
        // source, and a run says so here rather than in what the user sees.
        if !visible
            && pinned()
            && !pin_is_collapsed()
            && pin_film_thumbnail::the_source_is_standing(hosted_player_pid().unwrap_or(0))
        {
            note_pin_trace!(
                "film host hidden while a pin lives {:?}: the picture was the standing source",
                host,
            );
        }
        trace_the_host_show(host, visible);
        let _ = ShowWindow(host, if visible { SW_SHOWNOACTIVATE } else { SW_HIDE });
        // **The film follows the host off the screen and back**, which is the one road a source is
        // hidden on: the Shell stops drawing the thumbnail of a hidden window, and there is no host
        // to draw it in while the host is down (see `pin_film_thumbnail`).
        pin_film_thumbnail::set_the_source_visible(visible);
    }
    // The strips carry the chrome on every road, whatever the host is doing: the sync answers
    // what the chrome state asks for over the picture or over the cover (see `pin_chrome_strips`).
    pin_chrome_strips::sync_the_chrome_strips();
}

/// Take the film off the band, because the pin is about to show the host no film: the registration
/// goes first, so the hide that follows is a window going down rather than a picture going black.
///
/// It is one call rather than two at each road that ends a film, because the order is the whole of
/// it — a host hidden with the source the Shell draws still registered in it is the blank frame this
/// road is a film road to avoid (see `set_the_pin_film_host_visible` and
/// `pin_film_thumbnail::forget_the_thumbnail`).
pub(super) unsafe fn take_the_film_off_the_band() {
    pin_film_thumbnail::forget_the_thumbnail();
    // A cover standing in the host is that film's, and the host is about to show another one (see
    // `put_the_cover_away`).
    put_the_cover_away();
    unsafe { set_the_pin_film_host_visible(false) };
}

/// Destroy the host and forget the player it held: the pin it was showing a film for is ending, and
/// a window left standing would keep the desktop's band it owns. The strips stay: they are the pin's
/// window's, and the next take-up places them again (see `pin_chrome_strips`).
pub(super) unsafe fn take_down_the_pin_film_host() {
    // **A band a step held goes back before the host that owns it does.** The hold is a placement
    // of the pair, and one cleared after the host is gone would have the pin placed topmost by its
    // own teardown instead (see `PIN_STEP_BAND_HELD` and `own_the_pin_window`).
    unsafe { pin_lifecycle::give_the_step_band_back() };
    // **The picture the Shell is drawing goes before the window it is drawn in.** The film's window is
    // the player's own, so a registration left behind a host that is gone is the Shell's resource
    // leaking, and the source is taken off the screen with the host it was seen in (see
    // `pin_film_thumbnail`).
    pin_film_thumbnail::set_the_source_visible(false);
    // A pin's own end is the last settle of the held frames it counted, written before the record
    // that counted them goes (see `pin_film_thumbnail::settle_the_held_frames`).
    pin_film_thumbnail::settle_the_held_frames();
    pin_film_thumbnail::forget_the_thumbnail();
    // The cover went with the picture: there is no host left for it to be re-cropped into (see
    // `put_the_cover_away`).
    put_the_cover_away();
    // **The pin's chrome is given back to the Shell before the host that owns it goes.** Destroying
    // a window destroys the windows it owns, and the pin's window is kept between pins — so a
    // teardown that destroyed the owner first would take the pin's own window with it (see
    // `own_the_pin_window`).
    unsafe { own_the_pin_window(false) };
    // **The player is ended before the window it draws in is destroyed.** A player whose surface was
    // destroyed under it is a kill the player did not get to answer; the record below is what the
    // ordinary kill road retries to confirmation (see `terminate_player_pid`).
    if let Some(pid) = hosted_player_pid() {
        terminate_player_pid(pid);
    }
    if let Some(host) = film_host() {
        let _ = DestroyWindow(host);
    }
    PIN_FILM_HOST.store(0, Ordering::SeqCst);
    PIN_FILM_HOST_PLAYER.store(0, Ordering::SeqCst);
    PRESENTED_PID.store(0, Ordering::SeqCst);
}

/// The window the pin's film is drawn from, where the Shell is drawing it: the player's own window
/// the host holds a thumbnail of, which is the third window a press can land in and the one the
/// keyboard hook has to know about (see `key_input::pin_owns_caret`).
///
/// `None` where no host is up, and the player's window of the desktop's is the hook's own to know.
pub(super) fn the_films_own_window() -> Option<HWND> {
    film_host()
        .is_some()
        .then(pin_film_thumbnail::the_source_window)
        .flatten()
}

/// A client box as a trace reads it, or nothing where the window has already gone.
fn a_client_box(hwnd: HWND) -> Option<(i32, i32)> {
    let mut rect = RECT::default();
    (unsafe { GetClientRect(hwnd, &mut rect) }.is_ok()).then_some((rect.right, rect.bottom))
}

unsafe fn create() -> Option<HWND> {
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
    // **Made without `WS_VISIBLE`, and that is what keeps the band the film's until the Shell has
    // composed the picture.** A host put on the screen empty is a black rectangle where the hover's
    // picture was, for as long as the player takes to open the file. What shows it is the take-up of a
    // host the player is already drawing in, and the tick's own road (see
    // `take_up_the_pin_film_host`, `ensure_video_window_topmost`).
    //
    // **`WS_CLIPCHILDREN`** so that the host's own erase is clipped to the area outside the chrome
    // strips when a pin's window is not the one painting (see `pin_real_window::real_pin_style`).
    let host = CreateWindowExW(
        WS_EX_APPWINDOW,
        FILM_HOST_CLASS,
        w!("film host"),
        WS_POPUP | WS_CLIPCHILDREN,
        0,
        0,
        1,
        1,
        // A window of the desktop's has no parent — it is the one that has an *owner*, and the owner
        // is the pin's chrome (see `own_the_pin_window`).
        HWND::default(),
        None,
        hinstance,
        None,
    )
    .ok()?;

    // Disabled before it is ever shown: the picture is the Shell's thumbnail, and the press that
    // lands on it has to reach the pin (see the module doc).
    let _ = EnableWindow(host, false);
    PIN_FILM_HOST.store(host.0 as isize, Ordering::SeqCst);
    // **The host is the window the Shell lists**, so the button the pin used to carry names the file
    // the pin is of (see `pin_taskbar::set_pin_taskbar_title`).
    unsafe { pin_taskbar::set_pin_taskbar_title(host) };
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

/// Whether this hosted player has a frame of its own on screen: the source the Shell draws it from
/// *and*, where the player can be asked, a video output it says is configured for a frame.
///
/// A window is not a picture (see `the_host_is_drawing`), and `vo-configured` is the
/// player's own word that its video output has been set up for a frame. The answer is latched per
/// player because the read is bounded by a pipe timeout rather than free: once this player has
/// proved a frame it has one, and every later tick answers at once. A step begins a new player, so
/// `note_hosted_player` clears the latch for it.
///
/// **A player with no pipe to answer through is one whose picture is only ever its window**, and the
/// source being up is the whole of the proof for it: FFmpeg's player reports nothing at all (see
/// `mpv_ipc::answers_properties`), and mpv answers the same question with its own video output.
pub(super) fn the_hosted_frame_is_up(pid: u32) -> bool {
    if pid != 0 && PRESENTED_PID.load(Ordering::SeqCst) == pid {
        return true;
    }
    if !the_hosted_player_is_drawing(pid) {
        return false;
    }
    if !the_players_output_is_configured(pid) {
        return false;
    }
    PRESENTED_PID.store(pid, Ordering::SeqCst);
    true
}

/// Whether this player's own video output is configured for a frame, where the player answers over a
/// pipe: **the half of the presented predicate that does not need to know which second it is on**.
///
/// It is asked where a player has just been begun and the second it stands on is the caller's own
/// (see `pin_film_thumbnail::hand_the_picture_over`), and it is the same fact `the_hosted_frame_is_up`
/// proves a frame with, so the two cannot answer differently about the same player.
pub(super) fn the_players_output_is_configured(pid: u32) -> bool {
    !mpv_ipc::answers_properties(pid)
        || mpv_ipc::get_property(pid, "vo-configured").as_deref() == Some("true")
}

/// Whether the hosted player has the frame the hover froze on screen, rather than only a window.
///
/// **A window is not a picture.** A player makes its window before it has decoded anything, and a
/// source adopted for it is the most `the_hosted_player_is_drawing` can see — which leaves the band
/// black until the first frame is drawn into it, and is the empty stage the whole wait exists to keep
/// behind the hover. The two facts that separate a window from a picture are the player's own, read
/// over its pipe: `time-pos` has no value until a file is loaded and positioned, and
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
        note_pin_trace!(
            "handover handed over after {}ms",
            since.elapsed().as_millis()
        );
        disarm_pin_handover();
        // The handover is the pin's first frame, not an arrival: the chrome stays hidden over it
        // until the pointer asks for a strip (see `refresh_pin_chrome`).
        with_pin(|pin| pin.chrome = PinChrome::hidden());
        unsafe { show_pinned_window(pinned_window()) };

        // **And the film's own window is not asked for anything here**: it was placed once, at the
        // film's own size, when it was adopted, and a box the pin has now is a destination rectangle
        // (see `pin_film_thumbnail`).

        // **The hover is ended before the pin's player is let go of.** The two are the same sound at
        // the same second, and the unpause is a post to one process while the kill is a termination of
        // another: asked in the other order, both are audible until the kill lands, which is heard as
        // an echo of the frame the handover is standing on. The pin's window is already up over that
        // frame (see `show_pinned_window`), so the hover's going is the picture no longer and nothing
        // is left on screen for the order to change.
        retire_the_frozen_hover(frozen);

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
        note_pin_trace!(
            "handover failed: no frame after {}ms",
            since.elapsed().as_millis()
        );
        disarm_pin_handover();
        retire_the_frozen_hover(frozen);
        return PinHandover::Failed;
    }

    PinHandover::Waiting
}

#[cfg(test)]
mod tests {
    use super::*;

    /// **A show the trace already wrote is not written again**: the tick asks every pass, so a host
    /// shown, a host hidden and a host shown again are three lines, and a host shown once is one.
    #[test]
    fn a_host_show_is_written_when_it_changes_and_not_on_every_pass() {
        assert!(
            the_show_is_new(None, 0x2000, true),
            "the first show of a run is the state it began in"
        );
        assert!(
            !the_show_is_new(Some((0x2000, true)), 0x2000, true),
            "and an answer that has not moved is not written on every tick"
        );
        assert!(
            the_show_is_new(Some((0x2000, true)), 0x2000, false),
            "a hide is a change"
        );
        assert!(
            the_show_is_new(Some((0x2000, true)), 0x3000, true),
            "and so is another host"
        );
    }

    /// **The per-tick line is written on change and at most once a second otherwise**, which is what
    /// keeps a run's trace readable: the tick asks `the_hosted_player_is_drawing` hundreds of times a
    /// second and each answer it has not moved is one line an unchanged answer owes nobody.
    #[test]
    fn a_hosted_drawing_answer_is_traced_on_change_and_once_a_second_otherwise() {
        let now = Instant::now();
        let answer = (7, 0x2000, true, now);
        assert!(
            the_drawing_trace_is_due(None, answer, now),
            "the first answer of a run is the state it began in"
        );
        assert!(
            !the_drawing_trace_is_due(Some(answer), answer, now + Duration::from_millis(200)),
            "and an answer that has not moved is not written on every tick"
        );
        assert!(
            the_drawing_trace_is_due(Some(answer), answer, now + HOSTED_DRAWING_TRACE_EVERY),
            "the heartbeat is what says the tick is still asking it"
        );
        assert!(
            the_drawing_trace_is_due(
                Some(answer),
                (7, 0x2000, false, now + Duration::from_millis(1)),
                now + Duration::from_millis(1)
            ),
            "a transition is never missed, however soon after the last line it is: a player that
             draws and stops and draws again is three answers and not one"
        );
        assert!(
            the_drawing_trace_is_due(
                Some(answer),
                (8, 0x2000, true, now + Duration::from_millis(1)),
                now + Duration::from_millis(1)
            ),
            "and another player is another answer"
        );
        assert!(
            the_drawing_trace_is_due(
                Some(answer),
                (7, 0x3000, true, now + Duration::from_millis(1)),
                now + Duration::from_millis(1)
            ),
            "as is the same player in another host"
        );
    }
}

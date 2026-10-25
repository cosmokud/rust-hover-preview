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
//! `the_band_was_placed`). mpv's window is asked for a box once at each end of a drag and never
//! between (see `place_the_source`), and that ask is `SWP_ASYNCWINDOWPOS` like every other call made
//! on a window of another process's here.
//!
//! **Where the source stands is a knob, and the answer is measured: it has to be on the screen.**
//! `RHP_THUMBNAIL_SOURCE` names three placements, and two of them are known dead:
//!
//! - `behind` (the default) is the road this started as: the source stands at the pin's own box,
//!   under the host, and the host covers it. **It is the only placement that draws.** Measured on
//!   this machine: a resize drag through it is perfectly smooth with live video in the pin.
//! - `offscreen` puts the whole source outside the virtual screen, a desktop's width to its left and
//!   a desktop's height below it. **Known dead, kept for archaeology.** The registration succeeds and
//!   nothing is drawn: traced as `source adopted … placed by mpv 1920x1080+-4440+3880 placement
//!   Offscreen`, with the picture adopted and no error anywhere.
//! - `edge` leaves exactly 2 px of it over the virtual screen's bottom-right corner. **Known dead as
//!   well, and for the same reason** — `… 1920x1080+3438+1438 placement Edge`, registered, nothing
//!   drawn. Two pixels in bounds are not enough to count as a window the Shell is drawing; the
//!   composition surface of a source that is not genuinely on screen is not composed, so its
//!   thumbnail is empty.
//!
//! So the answer to the question the knob was built to ask is **no: the Shell does not keep a
//! thumbnail live for a window the user cannot see.** A source of a pin's film must be a window the
//! compositor is drawing, which is why the default is the one that stands under the host and why the
//! other two stay only as the evidence. A misspelt value is answered with `behind` too — a value that
//! names none of the three must never fall back to a placement that cannot draw.
//!
//! **A placement is made once at each end of a drag and never between.** The source takes the box the
//! pin began the drag at, is left exactly there while the hand is on the edge, and is asked for the
//! box the drag settled at when the hand lets go. What the two off-screen placements buy is the
//! shrink a `behind` source cannot follow: the destination shrinks, the source does not, and there is
//! nothing of a source that is not on the screen for the host to leave behind.
//!
//! **Only `behind` is moved with the host between those ends** (see `the_band_moves_the_source`),
//! because it is the only placement that has to stay under the pin to stay unseen; the two off-screen
//! placements have nothing to follow, and not asking is the point.
//!
//! **The source is never hidden.** The Shell draws nothing for a source it considers hidden — the
//! thumbnail of a `SW_HIDE` window is frozen at its last frame — so the source is *left on the
//! screen* at a placement where nothing of it shows, and it is taken off only where the host it is
//! seen in is going away for good (see `set_the_source_visible`). That is also why an `offscreen` or
//! `edge` source is **launched** where it will live rather than moved there after it is styled: mpv
//! shows its window when it opens the file, which is before this app has found it (see
//! `the_source_launch_origin`).
//!
//! shortcut: a `behind` source keeps the size a drag began at while the hand is on the edge, so a hand
//! that *shrinks* the box leaves the part of the source the host no longer covers standing on the
//! desktop until the drag ends. That was the leak the two off-screen placements were wanted for, and
//! they do not draw at all — so this is the shipped default and the leak is the price of a picture;
//! upgrade it by clipping the source where a window of this app's can clip it, not by moving it off
//! the screen.

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

use super::pin_trace::note_pin_trace;
use super::*;

/// Where the source window stands while the pin is on the screen: what `RHP_THUMBNAIL_SOURCE` names
/// (see [`the_source_placement`] and the module note on the three).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum SourcePlacement {
    /// Outside the virtual screen entirely: nothing of it can ever be seen — **and nothing of it is
    /// composed either, so its thumbnail is empty (measured, see the module note).**
    Offscreen,
    /// Exactly two of its pixels over the virtual screen's bottom-right corner: **not enough to be
    /// composed, so its thumbnail is empty (measured, see the module note).**
    Edge,
    /// At the pin's own box, under the host that covers it: the one placement the Shell draws.
    Behind,
}

/// **A prototype's second knob, and not a setting: `RHP_THUMBNAIL_SOURCE`, absent meaning
/// `behind`.**
///
/// The three were an A/B of the one thing this road cannot ask Windows without a screen to look at:
/// whether the Shell keeps the thumbnail of a source it is not drawing anywhere. It does not (see the
/// module note), so the one that draws is the default. A misspelt value is traced and answered with
/// `behind` rather than read as a road it does not name, and rather than falling back to a placement
/// the Shell will not draw.
const SOURCE_PLACEMENT_VARIABLE: &str = "RHP_THUMBNAIL_SOURCE";
/// How far beyond the virtual screen's left and bottom an `offscreen` source is put: a whole desktop
/// worth, which no monitor arrangement reaches, and well inside the ±32 767 that a window message's
/// packed coordinates can carry — past which Windows silently clamps the position (see
/// `the_source_launch_origin`).
const OFFSCREEN_GAP: i32 = 1_000;
/// How much of an `edge` source is inside the virtual screen: the 2×2 pixels of its corner that stand
/// over the screen's bottom-right one.
const EDGE_INSIDE_PIXELS: i32 = 2;

/// The placement a `RHP_THUMBNAIL_SOURCE` value names, or nothing for a value that names none.
///
/// `None` is the variable being absent, which is `behind` — the placement that draws.
fn source_placement_from(value: Option<&std::ffi::OsStr>) -> Option<SourcePlacement> {
    match value.and_then(|value| value.to_str()) {
        None | Some("behind") => Some(SourcePlacement::Behind),
        Some("offscreen") => Some(SourcePlacement::Offscreen),
        Some("edge") => Some(SourcePlacement::Edge),
        Some(_) => None,
    }
}

/// Where the machine asked the source to stand (see [`SOURCE_PLACEMENT_VARIABLE`]).
///
/// Read by the shipped binary and not by a test run, for the reason the thumbnail road's own knob is
/// (see `pin_film_host::the_pin_shows_its_film_by_thumbnail`): a suite that read it would answer
/// every test written about one placement with another's. A test stands a placement in with
/// [`stand_the_source_placement_in`].
fn the_source_placement() -> SourcePlacement {
    #[cfg(test)]
    {
        PLACEMENT_STAND_IN.with(|held| held.get())
    }
    #[cfg(not(test))]
    {
        let asked = std::env::var_os(SOURCE_PLACEMENT_VARIABLE);
        let placement = source_placement_from(asked.as_deref());
        if placement.is_none() {
            note_pin_trace!(
                "thumbnail source placement {:?} names none of behind/offscreen/edge; behind",
                asked
            );
        }
        placement.unwrap_or(SourcePlacement::Behind)
    }
}

/// Stand a placement in place of the machine's own answer on this thread (see
/// `the_source_placement`), which is the default where nothing is stood in.
#[cfg(test)]
pub(super) fn stand_the_source_placement_in(placement: SourcePlacement) {
    PLACEMENT_STAND_IN.with(|held| held.set(placement));
}

#[cfg(test)]
thread_local! {
    static PLACEMENT_STAND_IN: std::cell::Cell<SourcePlacement> =
        const { std::cell::Cell::new(SourcePlacement::Behind) };
}

/// The corner a placement puts the source at, given the virtual screen's box (`left, top, right,
/// bottom`), where the host's client area stands, and the box the source is wanted at.
///
/// The answer is in the virtual desktop's coordinates, which is the frame both things that place a
/// window here speak: `SetWindowPos` directly and `--geometry` through `the_source_launch_origin`.
fn source_origin(
    placement: SourcePlacement,
    virtual_box: (i32, i32, i32, i32),
    host_origin: (i32, i32),
    size: (i32, i32),
) -> (i32, i32) {
    let (left, top, right, bottom) = virtual_box;
    let (width, height) = (size.0.max(1), size.1.max(1));
    match placement {
        // Under the host, which is the only placement that has anything over it to keep it unseen.
        SourcePlacement::Behind => host_origin,
        // A desktop's width left of the virtual screen and a desktop's height below it, so that no
        // arrangement of monitors can reach any part of it.
        SourcePlacement::Offscreen => (
            left - (right - left).max(width) - OFFSCREEN_GAP,
            bottom + (bottom - top).max(height) + OFFSCREEN_GAP,
        ),
        // The screen's bottom-right corner with the source's own corner over it: the Shell sees a
        // window on the screen, and what it sees of it is `EDGE_INSIDE_PIXELS` square.
        SourcePlacement::Edge => (right - EDGE_INSIDE_PIXELS, bottom - EDGE_INSIDE_PIXELS),
    }
}

/// Whether a host placement asks the source to follow its corner, which is `behind`'s whole job and
/// neither off-screen placement's (see `the_band_was_placed`).
fn the_band_moves_the_source(placement: SourcePlacement) -> bool {
    placement == SourcePlacement::Behind
}

/// Where a source is launched: the placement its knob names, for the box the pin is at.
///
/// **`--geometry` is measured from the primary display's work area rather than its monitor**, which
/// mpv adds to the position it is given (`options/m_option.c`'s `m_geometry_apply` and
/// `video/out/win_state.c`, which offsets by the work area's origin). That origin is (0, 0) unless
/// the taskbar is docked along the primary's left or top edge, where an uncorrected `edge` placement
/// would be pushed off the screen by the taskbar's own thickness — and where a `behind` launch would
/// be a box away from the pin's. The offset is taken off every placement, so the three differ only in
/// the box they name.
///
/// **Nothing here is clamped by mpv.** Its geometry parser takes the numbers as integers and its
/// Windows backend applies them exactly where `--geometry` names an x and a y — the first placement
/// is forced rather than fitted to a screen (`window_resize`, whose `force_pos` is true until the
/// window's bounds are initialized), so a negative or a far-off-screen box lands where it says. The
/// one clamp left is the *size*: mpv limits it to 100% of the screen it picks, which a pin's box is
/// at or under.
///
/// **Measured, mpv 0.41.0, on a 3440×1440 virtual screen.** `--geometry=640x360+-4440+3880` and
/// `--geometry=640x360+3438+1438` are both applied exactly: Windows reads the windows back at
/// `-4440,3880` and `3438,1438`, 640×360, so mpv neither refuses a negative box nor fits it back onto
/// a screen. What is clamped is the *distance* — `--geometry=640x360+-1000000+-1000000` is handed to
/// `SetWindowPos` by mpv and read back at `-32768,-32768`, the signed 16-bit floor a window message's
/// packed coordinates carry. That is why an off-screen placement is one virtual desktop plus
/// `OFFSCREEN_GAP` rather than a megapixel: past ±32 767 the window is somewhere other than the box
/// that was asked for, and mpv's own idea of where its window is goes with it.
pub(super) fn the_source_launch_origin(at: (i32, i32), size: (i32, i32)) -> (i32, i32) {
    let screen = displays::the_virtual_screen();
    let (x, y) = source_origin(
        the_source_placement(),
        (screen.left, screen.top, screen.right, screen.bottom),
        at,
        size,
    );
    let work = displays::the_primary_work_area();
    (x - work.left, y - work.top)
}

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
/// bottom of the desktop's order, and registered in the host that shows it.
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
/// the hover the pin is being taken up from — or nowhere at all, on the two off-screen placements (see
/// `the_source_placement`).
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

    // **Where mpv put the window it made for itself, before this app asks it for anything.** The
    // source is launched at its placement (see `the_source_launch_origin`), and this is the line that
    // says whether mpv honoured it — the box a placement is corrected to below is asked for beside it
    // in the trace.
    let mut created = RECT::default();
    if unsafe { GetWindowRect(source, &mut created) }.is_ok() {
        note_pin_trace!(
            "source adopted {:?} pid {} placed by mpv {}x{}+{}+{} placement {:?}",
            source,
            pid,
            created.right - created.left,
            created.bottom - created.top,
            created.left,
            created.top,
            the_source_placement(),
        );
    }

    if let Some(host) = pin_film_host::film_host() {
        // **The placement is made in the adoption pass, before the registration draws anything.**
        // Styling, the bottom insert and this are one pass, and for the two off-screen placements the
        // launch already put the window where nothing of it can be seen — this only corrects the box
        // to the host's (see `place_the_source`).
        place_the_source(host);
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

/// Follow the host: register the picture where there is none, and point it at the host's whole client
/// area.
///
/// **The source is not resized here**, which is the whole of what makes a resize step cheap: the
/// destination is the host's box and the Shell stretches whatever the source's box is into it. It is
/// asked for a box once at each end of a drag and never between (see `place_the_source`).
///
/// **Nor is it moved here, on the two off-screen placements** — the point of those is that a step of a
/// drag asks mpv's process for nothing at all. A `behind` source does follow the host's corner, which
/// is the only thing keeping it under the host rather than standing where the pin used to be (see
/// `the_band_moves_the_source`).
///
/// **What a drag step costs is what follows**: one comparison, and a `DwmUpdateThumbnailProperties`
/// only where the host's client box changed.
pub(super) fn the_band_was_placed(host: HWND) {
    let number = SOURCE.load(Ordering::SeqCst);
    if number == 0 {
        return;
    }
    let Some(destination) = (unsafe { client_size(host) }) else {
        return;
    };

    if the_band_moves_the_source(the_source_placement()) {
        if let Some(origin) = unsafe { client_origin(host) } {
            move_the_source(HWND(number as *mut _), origin);
        }
    }

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

/// Put the source where its placement wants it, at the box the host is at: the one ask a drag makes of
/// a window of mpv's, made at the drag's two ends and at every box that settles without one.
///
/// **Its size is the box the pin is at, and it is not asked for again until the box settles.** A drag
/// that shrinks the pin leaves a `behind` source at the box the drag began at, and the part of it the
/// host no longer covers stands on the desktop until the drag ends — which is the leak the two
/// off-screen placements exist to leave behind (see the module note and `source_origin`).
///
/// Nothing is asked of a window already at the box the placement names, which is what makes the call
/// at a drag's beginning free where the source has not left its placement since the last one.
pub(super) fn place_the_source(host: HWND) {
    let number = SOURCE.load(Ordering::SeqCst);
    if number == 0 {
        return;
    }
    let source = HWND(number as *mut _);
    let Some(origin) = (unsafe { client_origin(host) }) else {
        return;
    };
    let Some(size) = (unsafe { client_size(host) }) else {
        return;
    };
    let size = (size.0.max(1), size.1.max(1));
    let placement = the_source_placement();
    let screen = displays::the_virtual_screen();
    let (x, y) = source_origin(
        placement,
        (screen.left, screen.top, screen.right, screen.bottom),
        origin,
        size,
    );

    // Safety: the handle is a window this module adopted, and both questions are about the window.
    if !unsafe { IsWindow(source) }.as_bool() {
        return;
    }
    let mut current = RECT::default();
    if unsafe { GetWindowRect(source, &mut current) }.is_err() {
        return;
    }
    if (current.left, current.top, current.right - current.left, current.bottom - current.top)
        == (x, y, size.0, size.1)
    {
        return;
    }

    let placed = unsafe {
        SetWindowPos(
            source,
            HWND::default(),
            x,
            y,
            size.0,
            size.1,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS,
        )
    };
    // The ask and its answer, because a placement is the one thing on this road that is a guess about
    // another process's window: a `--geometry` mpv did not honour shows up as a `placed` of zero here.
    note_pin_trace!(
        "source placed {:?} {:?} {}x{}+{}+{} host {:?} ok {}",
        source,
        placement,
        size.0,
        size.1,
        x,
        y,
        host,
        placed.is_ok() as u8,
    );
}

/// Take the picture off the screen with the host it is seen in, or put it back.
///
/// **A hidden source is a thumbnail the Shell stops drawing**, so this is only ever asked where the
/// host goes down with it — a pin put away, a park over the band — and never while the host is up.
/// The call is posted rather than made, for the reason every call on this window is (see
/// `adopt_the_source`).
///
/// **The two off-screen placements are not touched at all**: a source that is nowhere the user can
/// see has nothing to take off the screen, and hiding it would freeze the thumbnail the host it is
/// seen in is about to want back (see `the_source_placement`). A source at the bottom of the z-order
/// is the only one there is a reason to hide, so `behind` is the only placement that asks.
pub(super) fn set_the_source_visible(visible: bool) {
    let number = SOURCE.load(Ordering::SeqCst);
    if number == 0 || the_source_placement() != SourcePlacement::Behind {
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
            // window and not yet presented is a registration the Shell can refuse. **A refusal is the
            // answer this road is asking for where the source is off the screen**, which is why it is
            // traced rather than swallowed (see `the_source_placement`).
            note_pin_trace!(
                "thumbnail register failed host {:?} source {:?} placement {:?}",
                host,
                source,
                the_source_placement(),
            );
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
        note_pin_trace!(
            "thumbnail properties failed host {:?} destination {:?} placement {:?}",
            host,
            destination,
            the_source_placement(),
        );
        let _ = unsafe { DwmUnregisterThumbnail(held.registration) };
        *held = Thumbnail::NOTHING;
        REGISTERED_HOST.store(0, Ordering::SeqCst);
        return;
    }
    held.destination = destination;
}

/// Move the source so its own corner is the host's, keeping the size it has: the one thing that keeps
/// a `behind` source under the host as the pin is carried. It is asked of a host placement and never
/// of the other two placements (see `the_band_moves_the_source`), because neither of those has a
/// corner it needs to ride with.
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

    /// The knob names three roads, and the one that draws is the one a machine that says nothing
    /// gets. A misspelling is answered by the default above and traced, rather than read as a
    /// placement it does not name.
    #[test]
    fn the_source_placement_knob_names_three_roads_and_defaults_to_behind() {
        assert_eq!(
            SOURCE_PLACEMENT_VARIABLE, "RHP_THUMBNAIL_SOURCE",
            "the knob is the name the branch's notes and the report give it"
        );
        let named = |value: Option<&str>| {
            source_placement_from(value.map(std::ffi::OsStr::new))
        };
        assert_eq!(named(None), Some(SourcePlacement::Behind));
        assert_eq!(named(Some("behind")), Some(SourcePlacement::Behind));
        assert_eq!(named(Some("offscreen")), Some(SourcePlacement::Offscreen));
        assert_eq!(named(Some("edge")), Some(SourcePlacement::Edge));
        assert_eq!(named(Some("Edge")), None);
        assert_eq!(named(Some("on-screen")), None);
        assert_eq!(named(Some("")), None);
        // A value that names no placement is answered with the one that draws, never with a
        // placement the Shell will not compose.
        assert_eq!(
            source_placement_from(Some(std::ffi::OsStr::new("bhind")))
                .unwrap_or(SourcePlacement::Behind),
            SourcePlacement::Behind,
        );
    }

    /// Each placement puts the source exactly where its name says, and the two off-screen ones leave
    /// nothing of it on a screen however the monitors are arranged.
    #[test]
    fn every_placement_puts_the_source_where_its_name_says() {
        let screen = (0, 0, 3440, 1440);
        let size = (640, 360);

        assert_eq!(
            source_origin(SourcePlacement::Behind, screen, (100, 200), size),
            (100, 200),
            "behind is the host's own corner and nothing else"
        );

        let (x, y) = source_origin(SourcePlacement::Offscreen, screen, (100, 200), size);
        assert!(x + size.0 <= screen.0 && y >= screen.3, "fully off the screen: {x},{y}");
        assert!(x < -3000 && y > 2400, "and far enough off it to be worth the name");
        // A source wider than any display still starts past the screen's left edge.
        let (wide, _) = source_origin(SourcePlacement::Offscreen, screen, (100, 200), (8000, 360));
        assert!(wide + 8000 <= screen.0, "a source wider than the desktop is still not on it");

        assert_eq!(
            source_origin(SourcePlacement::Edge, screen, (100, 200), size),
            (3438, 1438),
            "2 px of its corner over the screen's bottom-right one, and no more"
        );
        let (x, y) = source_origin(SourcePlacement::Edge, (-1920, 0, 1920, 1080), (0, 0), size);
        assert_eq!((x, y), (1918, 1078), "a screen that starts left of the primary is its own box");
    }

    /// A test stands the placement it is about in, and is answered with the one that draws otherwise
    /// — the shipped binary's own answer when the variable says nothing.
    #[test]
    fn a_test_stands_the_placement_it_is_about_in() {
        assert_eq!(the_source_placement(), SourcePlacement::Behind);
        stand_the_source_placement_in(SourcePlacement::Offscreen);
        assert_eq!(the_source_placement(), SourcePlacement::Offscreen);
        stand_the_source_placement_in(SourcePlacement::Edge);
        assert_eq!(the_source_placement(), SourcePlacement::Edge);
        stand_the_source_placement_in(SourcePlacement::Behind);
    }

    /// A step of a drag asks nothing of the source on the two off-screen placements: what follows the
    /// host is the destination rectangle and nothing else (see `the_band_was_placed`).
    #[test]
    fn only_a_source_behind_the_host_follows_it() {
        assert!(the_band_moves_the_source(SourcePlacement::Behind));
        assert!(!the_band_moves_the_source(SourcePlacement::Offscreen));
        assert!(!the_band_moves_the_source(SourcePlacement::Edge));
    }
}

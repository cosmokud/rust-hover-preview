//! The pin's film drawn by the Desktop Window Manager rather than nested in a window of this app's.
//!
//! **This is the road a pin's film is drawn on, whichever engine plays it, and the only one.** mpv is
//! launched with no `--wid` and makes a window of the desktop's of its own; FFmpeg's player has no
//! embed switch to pass and makes one too. The pin's film host asks the Shell to draw that window into
//! itself (`DwmRegisterThumbnail`), and what follows a resize is one rectangle rather than two windows
//! agreeing about a box. **The rules this file implements, the roads that lost and the measured
//! evidence are `Pin mode → The film road` in `ARCHITECTURE.md`.**
//!
//! **Why the film has to be a window of the desktop's.** A thumbnail is taken of a window's
//! composition surface, and a `WS_CHILD` has none of its own — it is drawn into its parent's — so a
//! film that is to be thumbnailed cannot be a child of the host. That is the whole of why this road
//! launches mpv without `--wid` (see `video_launch::mpv_args`), and the whole of why FFmpeg's player
//! is no longer `SetParent`ed into one.
//!
//! **The source is placed once, at the film's own size, and never again.** That is the whole of this
//! road now, and it is where two flashes were: a 19.4 s trace of the road that resized the source at
//! each end of a drag had 326 of 2,821 thumbnail frames whose crop's window disagreed with the
//! source's live client — every one of them within a frame or two of a re-place, 1 frame (7 ms) at
//! the press with the crop 2.1–2.4× too small, and 72–130 frames (145 ms) at the release with the
//! crop sampling 2.9–3.3× past the live client, which is the zoom-out with black around it. The
//! middle of a drag was perfect, because nothing moved there. So nothing moves here: the source takes
//! its box as soon as the film's own size is known — at adoption, or on the pass that watches the
//! player where the player had not opened the file yet — the picture adoption read is a constant for
//! the pin's life, and a resize, a maximize, a restore, a zoom, a walk, a press, a release and a
//! settle are all **one `DwmUpdateThumbnailProperties`** — the destination and the crop, and nothing
//! cross-process at all (see `place_the_source`, `place_the_source_when_the_player_answers` and
//! `the_band_was_placed`).
//!
//! **The size is the film's own picture size, never the pin's box.** The source is placed at the size
//! the film would be drawn at were it not fitted to a window, from the player's own `video-params` —
//! or from the probe's measurement of the same film where the player answers nothing at all, which
//! is waited for rather than guessed at. **FFmpeg's player answers nothing**, so the probe's
//! measurement is the whole of what its source is placed at (see `the_films_native_size`). A pin at or under that size is a *downscale* the
//! Shell makes, which is sharp at every box the hand pulls it to; a pin over it is the upscale the
//! player itself would have to make, from the film's own pixels and not from a box three times a
//! small pin. **`RHP_THUMBNAIL_SOURCE` used to offer the pin's box as the other arm, and it is gone**:
//! `box` only ever meant one placement at the box the pin was adopted at, which nothing measures and
//! nothing wants, so the film's own size is the only size there is.
//!
//! **What a thumbnail shows is the source's picture, not its whole client**: the rectangle the film
//! is drawn in is cropped out of the source to **the box's own shape**, and the whole box is handed
//! over as the destination, so the film fills its pin (see `the_thumbnail_box`).
//! **`RHP_THUMBNAIL_PICTURE` used to offer `client` as the other arm, and it is gone**: stretching the
//! whole client, margins included, into the box is the pillar of black that read as a pin zoomed out
//! (measured: `client 916x515 … picture 2595x1460 margins l311t0r311b0` — 311 px a side stretched
//! into the picture) and the upscale that read as one zoomed in.
//!
//! **One source per live film, and one registration per destination host.** The source is the live
//! film's own window, placed once at the film's own size, and the registration that draws it lives in
//! the host showing it: the hover's while a hover is up, and the pin's once the hover is taken up.
//! **A take-up begins no player and places nothing**: the hover's host is promoted into the pin's, so
//! the destination is the very window the registration is already in, and the pin's box is one
//! re-crop of the same picture (see `pin_film_host::promote_the_hover_film_host` and
//! `hover_film_host`). A hover that is never pinned takes its source and its registration down with
//! it (see `forget_the_thumbnail_of`).
//!
//! **Filling the box is not a detail of the crop, it is what makes a bar impossible.** The road this
//! replaces fitted the picture into the box at the picture's own shape and left the rest to the
//! host's black: a destination *smaller* than the client whenever the box's aspect and the film's
//! were not the same, read as the film squashed into the middle of the pin with a pillar either side.
//! The destination is the client and the crop is the part of the picture that has the client's
//! shape, so there is no rectangle left over to read as the wrong thing — whatever order anything is
//! ordered in, and whatever the two shapes are.
//!
//! **A shape is one film's.** The client and the shape the crop is cut by are invalidated with the
//! film — teardown forgets them, adoption of another window resets them, and every frame checks the
//! record still belongs to the source it was placed for (see `forget_the_shape_of_another_film`) —
//! and the probe answers only for the film the player was begun for (see `the_films_own_size`), so no
//! frame of the film that follows can be cropped by the shape of the one that ended.
//!
//! **The composition invariant is one line**: the destination is the host's whole client, so the
//! Shell draws every pixel of the pin, and the host's own erase paints that same client under it —
//! which is what a client the Shell has not composed yet, or one whose source is not drawing, shows
//! instead of the desktop. **The source's side of the pair is a constant**, so a frame can only
//! change with the host's box: the trace's `cli` (the source's live client) and `pic` (the picture
//! the crop was built from) are the same number for the pin's whole life, and a frame where they
//! disagree says something other than this app resized the source (see `note_the_thumbnail_pair`).
//!
//! **The invariant has teeth in this module, not only in prose.** Three calls name the source's
//! geometry and there are no others: `size_the_source` — the one placement, reached only through
//! `place_the_source` — `move_the_source`, which carries `SWP_NOSIZE` and so cannot change a size,
//! and `lower_the_source`, which carries `SWP_NOMOVE | SWP_NOSIZE` and only changes the z-order. A
//! second placement is refused out loud (`source re-placed`) and counted, so a path that reaches one
//! is a line in the trace rather than a frame nobody explained.
//!
//! **Where the source stands, and what that was measured to be.** It stands at the pin's corner,
//! under the host, and it is **moved when the piece of it the shell samples would otherwise leave the
//! screen** (see `keep_the_crop_on_the_screen` and below). The placement knob this road once had named
//! two other placements and both are dead: `offscreen` (a desktop's width left and a desktop's height
//! below the virtual screen) traced as `source adopted … 1920x1080+-4440+3880` with the registration
//! succeeding and the pin empty, and `edge` (2 px over the screen's bottom-right corner) the same at
//! `… 1920x1080+3438+1438`. **The Shell does not keep a thumbnail live for a window the user cannot
//! see**: the composition surface of a source that is not genuinely on screen is not composed, so its
//! thumbnail is empty. The knob is gone with them — a placement that cannot draw is not an arm of an
//! A/B — and the source is launched where it will stand rather than moved there after it is styled,
//! because mpv shows its window as soon as it opens the file (see `the_source_launch_origin`).
//!
//! **The source is moved, and never resized, to keep the sampled rectangle on the screen.** The
//! source is native-sized and no longer follows the pin, so a pin at an edge or a corner of the
//! desktop has its crop hanging off it — and that is the black the dead placements measured, from the
//! other side. A move cannot change the content the crop describes (the crop is relative to the
//! source's client, and its shape is the box's), while a resize is the flash this road exists to not
//! have; so when a frame's crop rect in screen coordinates would leave the work area of the monitor
//! the pin is on, the source is slid the shortest way back onto it, on a 64 px grid so a slow drag at
//! an edge moves it once every 64 px rather than every frame. **A source that already fits is not
//! touched**, which is every ordinary centred drag (see `keep_the_crop_on_the_screen`).
//!
//! **A box over the film's own size upscales the picture whole.** The crop of a box the picture does
//! not fill is the picture, cut to nothing, and the Shell stretches it into the box — cropping a
//! picture that is already being magnified would throw pixels away for a softer pin. That is the
//! upscale a very large pin is drawn with, and the trace's `clamp` field is where a run says so (see
//! `the_crop_filling`).
//!
//! **The film's own size is `video-params`' `dw`/`dh`, and not `osd-dimensions`.** `dwidth`/`dheight`
//! are the film's *display* size — what it would be drawn at were it not fitted to a window — which
//! is the number this road wants: measured, they read `1920x1080` whatever box the source is in.
//! `osd-dimensions` is the player's answer about its *output*: on mpv 0.41.0 a 1920×1080 film in a
//! 640×300 window answers `{"w":640,"h":300,"par":1.000000,"aspect":2.133333,"mt":0,"mb":0,"ml":53,
//! "mr":54}` — the window the output configured for and the letterbox it leaves, which is what the
//! box made the picture rather than what the picture is (see `the_players_picture_size`). A player
//! that has made its window and not yet opened the file is asked a bounded handful of times, and the
//! probe's own measurement of the same film answers for it (see `the_films_native_size`).
//!
//! **The structural answer to a source mid-reconfigure is a flip-model swapchain**, which the
//! compositor scales to the window's client area by construction instead of relying on mpv catching
//! a `WM_SIZE`: `RHP_THUMBNAIL_FLIP=1` passes `--d3d11-flip=yes` to a pinned mpv on this road and
//! this road only, so the user can A/B the two runs. It is **never asked for on a source made
//! invisible**: a layered window has no flip-model swapchain, so `RHP_THUMBNAIL_FLIP` is read,
//! overruled and traced as overruled rather
//! than dropped on the floor (see `video_launch::the_thumbnail_swapchain_flip`).
//!
//! **The source is never hidden and never resized.** Hidden, the Shell stops drawing it: the thumbnail
//! goes black **18 to 71 ms** after the `SW_HIDE` and stays black until something draws over it —
//! measured on glass, five seeks, gaps of 415, 500, 520, 555 and 700 ms (the `hide` road, see
//! `still_probe` and the seek road table in `ARCHITECTURE.md`). **The sentence this file used to
//! carry here — that a hidden window's thumbnail is frozen at its last frame — was measured and is
//! false**, and believing it cost four cover roads that each froze the outgoing player. A DWM
//! thumbnail cannot outlive the drawing of its source, so a cover has to be a picture of this app's
//! own (see `stand_the_outgoing_picture_from_its_dc`). The source is therefore taken off the screen
//! only where the host it is seen in is going away for good (see `set_the_source_visible`).
//! Resized is the
//! flash this road was measured with. Moved it is, but only to keep the crop on the screen and never
//! with a new size: one placement, and then a slide at an edge and nothing else until teardown (see
//! `keep_the_crop_on_the_screen`).
//!
//! **On this road it is also made invisible rather than merely covered** (`RHP_THUMBNAIL_INVISIBLE`,
//! on by default and off with `=0`): the source wears `WS_EX_LAYERED` with
//! `SetLayeredWindowAttributes(…, 1, LWA_ALPHA)`, so the compositor still draws it — a source the
//! Shell is not drawing has no picture to thumbnail at all (measured, see the placements above) —
//! while nothing of it reaches the screen. That is the same one-step-from-nothing band this app's own
//! chrome used to wear, and it is what lets the source stand at the film's own size rather than the
//! pin's: a window nobody can see cannot spill past a small pin. It is put on in the adoption pass
//! and re-asserted by the style monitor if mpv's window comes back without it (see
//! `keep_the_source_invisible`).
//!
//! **The bet this road is:** the Shell's thumbnail of a layered window is the window's own content at
//! full opacity, so the film still draws in the pin with the source itself invisible. It has two ways
//! to lose, and the run says which: a pin that goes black while the player is still alive is a Shell
//! sampling the blended (1/255) result instead of the window's content, while a pin that goes black
//! and takes the player with it is a present that cannot target a layered window at all. That is also
//! why nothing here hides the window to achieve the invisibility: the two failures are told apart by
//! whether the player is still there, and a hidden source would answer that question with itself.
//!
//! shortcut: a source only ever slides far enough to keep its crop on the screen, so on
//! `RHP_THUMBNAIL_INVISIBLE=0` — the visible road, where the film is covered rather than composited
//! away — a pin carried across the screen leaves its film behind and a pin smaller than its film has
//! the film standing around it. That road is an A/B of the invisibility and nothing else. Upgrade it
//! by clipping the source in a window of this app's, not by following the pin with a window of mpv's.

use std::sync::atomic::{AtomicIsize, AtomicU32, Ordering};
use std::sync::Mutex;
use std::time::{Duration, Instant};

use windows::Win32::Foundation::{COLORREF, HWND, POINT, RECT};
use windows::Win32::Graphics::Dwm::{
    DwmFlush, DwmRegisterThumbnail, DwmUnregisterThumbnail, DwmUpdateThumbnailProperties,
    DWM_THUMBNAIL_PROPERTIES, DWM_TNP_RECTDESTINATION, DWM_TNP_RECTSOURCE,
    DWM_TNP_SOURCECLIENTAREAONLY, DWM_TNP_VISIBLE,
};
use windows::Win32::Graphics::Gdi::{
    BitBlt, ClientToScreen, GetDC, GetPixel, InvalidateRect, ReleaseDC, UpdateWindow, CLR_INVALID,
};
use windows::Win32::UI::WindowsAndMessaging::{
    GetClientRect, GetSystemMetrics, GetWindowLongPtrW, GetWindowRect, IsWindow, IsWindowVisible,
    SetLayeredWindowAttributes, SetWindowLongPtrW, SetWindowPos, ShowWindowAsync, GWL_EXSTYLE,
    HWND_BOTTOM, LWA_ALPHA, SM_CXMAXTRACK, SM_CYMAXTRACK, SWP_ASYNCWINDOWPOS, SWP_NOACTIVATE,
    SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER, SW_HIDE, SW_SHOWNOACTIVATE, WS_EX_APPWINDOW,
    WS_EX_LAYERED, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_EX_TRANSPARENT,
};
// Only the binary asks a window who owns it: a test has one process (see `the_source_is_up`).
#[cfg(not(test))]
use windows::Win32::UI::WindowsAndMessaging::GetWindowThreadProcessId;

use super::pin_trace::note_pin_trace;
use super::*;

// **`RHP_THUMBNAIL_SOURCE` lost and is gone.** It named `native` (one placement at the film's own
// picture size) or `box` (one placement at the pin's box the source was adopted at), and nothing
// measures or wants the second: the film's own size is the only size a source is placed at (see
// `the_one_placement`).

/// **A third knob, and not a setting either: `RHP_THUMBNAIL_INVISIBLE`, on unless it says `0`.**
///
/// It is the one variable this road's experiment is run by: with it on the source wears layered
/// alpha 1, so nothing of it reaches the screen — which is what lets it stand at the film's own size
/// rather than the pin's (see the module note); with it off the source is the covered window it always
/// was, and it is only ever covered where the pin is bigger than the film. Only `0` turns it off — a
/// value that names no road is not read as one it does not name — and the launch line says which of
/// the two a run was.
const INVISIBLE_VARIABLE: &str = "RHP_THUMBNAIL_INVISIBLE";
/// The alpha the invisible source is drawn at: 1 of 255, the one step that is composited rather than
/// absent, which is what keeps the Shell drawing it (see the module note).
const INVISIBLE_ALPHA: u8 = 1;

// **`RHP_THUMBNAIL_PICTURE` lost and is gone.** It named `picture` (the rectangle of the film's own
// picture, cropped to the box's shape) or `client` (the source's whole client stretched into the
// box, margins and all), and the second is the pillar of black the module note measures — so the
// picture is the only thing a thumbnail of the source shows (see `the_thumbnail_box`).

/// Whether the machine asked this road to make its source invisible (see [`INVISIBLE_VARIABLE`]).
///
/// Read by the shipped binary and stood in by a test, for the reason the size knob above is.
pub(super) fn the_source_is_invisible() -> bool {
    #[cfg(test)]
    {
        INVISIBLE_STAND_IN.with(|held| held.get())
    }
    #[cfg(not(test))]
    {
        source_is_invisible_from(std::env::var_os(INVISIBLE_VARIABLE).as_deref())
    }
}

/// The knob's own reading: invisible unless the value is exactly `0`.
fn source_is_invisible_from(value: Option<&std::ffi::OsStr>) -> bool {
    value.and_then(|value| value.to_str()) != Some("0")
}

/// Stand the invisible road (or the visible one) in place of the machine's own answer on this
/// thread, which is the default where nothing is stood in (see `the_source_is_invisible`).
#[cfg(test)]
pub(super) fn stand_the_invisible_road_in(invisible: bool) {
    INVISIBLE_STAND_IN.with(|held| held.set(invisible));
}

#[cfg(test)]
thread_local! {
    static INVISIBLE_STAND_IN: std::cell::Cell<bool> = const { std::cell::Cell::new(true) };
}

/// The extended styles this road puts on a player's own window (see `adopt_the_source`).
///
/// `WS_EX_TOOLWINDOW` is the window's leaving the taskbar and Alt+Tab, `WS_EX_NOACTIVATE` is it never
/// taking the keyboard from the file being renamed, and `WS_EX_TRANSPARENT` is a press falling through
/// it should it ever be in front of anything. `WS_EX_LAYERED` is the invisibility itself, and the app
/// and topmost bits mpv opens with are cleared: this window is behind the pin and must never be held
/// over it (see `video_launch::mpv_args`, which no longer asks for `--ontop` here).
fn source_ex_style(current: isize, invisible: bool) -> isize {
    let style = (current & !(WS_EX_APPWINDOW.0 as isize) & !(WS_EX_TOPMOST.0 as isize))
        | WS_EX_TOOLWINDOW.0 as isize
        | WS_EX_NOACTIVATE.0 as isize
        | WS_EX_TRANSPARENT.0 as isize;
    if invisible {
        style | WS_EX_LAYERED.0 as isize
    } else {
        style & !(WS_EX_LAYERED.0 as isize)
    }
}

/// How many times adoption asks the player for the film's own picture size before it leaves the
/// placement to the road that watches the player: a handful of bounded reads a few milliseconds
/// apart, which is what a player that has made its window and not yet opened the file needs to
/// answer (see `the_films_native_size`).
const NATIVE_SIZE_ASKS: u8 = 4;
/// How long adoption waits between two of those reads. **This is the only wait this road makes and it
/// is not on a drag's road**: it happens once, in the adoption pass, and it is bounded by
/// `NATIVE_SIZE_ASKS` reads that are each bounded themselves (see `adopt_the_source`).
const NATIVE_SIZE_WAIT: Duration = Duration::from_millis(3);

/// How long a film that was not sized at adoption is given to answer before the road stops asking:
/// the player answers when it has opened the file, and a placement that waits for that answer is a
/// placement at the film's own size rather than at whatever box the player made for itself (see
/// `place_the_source_when_the_player_answers`).
const LATE_NATIVE_WAIT: Duration = Duration::from_secs(1);
/// How often the road asks a player that has not answered yet: a paced read rather than a read on
/// every pass of the monitor that asks it (see `a_late_placement_is_due`).
const LATE_NATIVE_EVERY: Duration = Duration::from_millis(64);

/// How often a source whose player has no pipe is sampled for a picture: **one DC and a handful of
/// pixel reads, paced to the monitor that asks it and never on a frame** (see `the_source_drew`).
/// It is the same pace a player that has not answered its own size is asked at, because it is the
/// same kind of wait.
const SOURCE_DREW_EVERY: Duration = Duration::from_millis(64);

/// Whether a source with no pipe has drawn, as a test stands it in: `Some(true)` and `Some(false)` are
/// a source that drew and one that never does, and `None` reads the source's own pixels (see
/// `the_source_has_lit_pixels`).
///
/// It is `Some(true)` until a test says otherwise: the tests that are about a swap are about the swap,
/// and the sampled road has tests of its own that stand both answers in.
#[cfg(test)]
static DREW_STAND_IN: Mutex<Option<bool>> = Mutex::new(Some(true));

/// Stand an answer for the drawn sample in place of the machine's own (see `the_source_drew`).
#[cfg(test)]
pub(super) fn stand_the_drawn_probe_in(answer: Option<bool>) {
    if let Ok(mut held) = DREW_STAND_IN.lock() {
        *held = answer;
    }
    // The pace is a cost on the machine's own sample, and a test that changes the answer is not made to
    // wait it out: the answer it stood in is the next sample the pass would have taken.
    if let Ok(mut pending) = PENDING_SOURCE.lock() {
        if let Some(waiting) = pending.as_mut() {
            waiting.sampled = None;
        }
    }
}

/// Stand the moment the source waiting to be handed the picture over was adopted, or give its own
/// answer back with `None` — a test of the bound that sample is waited for, which cannot wait out the
/// start bound itself (see `hand_the_picture_over`).
#[cfg(test)]
pub(super) fn stand_the_pending_source_since(since: Option<Instant>) {
    if let Ok(mut pending) = PENDING_SOURCE.lock() {
        if let Some(waiting) = pending.as_mut() {
            waiting.since = since.unwrap_or_else(Instant::now);
        }
    }
}

/// Where a source is launched: the pin's own corner.
///
/// **The source is launched where it will stand rather than at wherever mpv would put it**: mpv shows
/// its own window as soon as it has opened the file — before this app has found the window and adopted
/// it — and the one placement this road makes puts it under the host (see the module note).
///
/// **`--geometry` is measured from the primary display's work area rather than its monitor**, which
/// mpv adds to the position it is given (`options/m_option.c`'s `m_geometry_apply` and
/// `video/out/win_state.c`, which offsets by the work area's origin). That origin is (0, 0) unless
/// the taskbar is docked along the primary's left or top edge, where an uncorrected launch would be a
/// taskbar's thickness away from the corner of the box showing the film. The offset is taken off the
/// corner and nothing else.
///
/// **Nothing here is clamped by mpv.** Its geometry parser takes the numbers as integers and its
/// Windows backend applies them exactly where `--geometry` names an x and a y — the first placement
/// is forced rather than fitted to a screen (`window_resize`, whose `force_pos` is true until the
/// window's bounds are initialized), so a negative box lands where it says. The one clamp is the
/// *size*: mpv limits it to 100% of the screen it picks, which a pin's box is at or under — and the
/// source's own size is asked for after the launch, by this app and not by an argument (see
/// `place_the_source`).
pub(super) fn the_source_launch_origin(at: (i32, i32)) -> (i32, i32) {
    let work = displays::the_primary_work_area();
    (at.0 - work.left, at.1 - work.top)
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

/// How many times a source has been placed since a test last cleared it: the seam that says a take-up
/// re-crops the one picture rather than asking the player for anything (see `place_the_source`).
#[cfg(test)]
static SOURCE_PLACEMENTS: AtomicU32 = AtomicU32::new(0);
/// How many times a second placement has been refused since a test last cleared it: the seam that
/// says no path reached `source re-placed` silently (see `place_the_source`).
#[cfg(test)]
static SOURCE_REPLACEMENTS: AtomicU32 = AtomicU32::new(0);

/// How many registrations the Shell has been asked for since a test last cleared it: one per source
/// window, so a source that replaced itself is a second registration rather than a second thumbnail
/// over the first (see `settle_the_registration`).
#[cfg(test)]
static REGISTRATIONS: AtomicU32 = AtomicU32::new(0);

/// How many registrations have been let go of since a test last cleared it: a handover is one swap and
/// **one** old registration, and never a picture dropped before the next one is there (see
/// `hand_the_picture_over`).
#[cfg(test)]
static UNREGISTRATIONS: AtomicU32 = AtomicU32::new(0);

/// The one host this road has drawn a picture into, or nothing while no registration is recorded: a
/// pin's host and a hover's are the same one at the take-up that promotes it (see
/// `settle_the_registration`).
#[cfg(test)]
pub(super) fn the_registration_host() -> Option<isize> {
    THUMBNAIL
        .lock()
        .ok()
        .and_then(|held| (held.host != 0).then_some(held.host))
}

/// How many placements this run has made (see `place_the_source`).
#[cfg(test)]
pub(super) fn the_source_placements() -> u32 {
    SOURCE_PLACEMENTS.load(Ordering::SeqCst)
}

/// How many second placements this run has refused (see `place_the_source`).
#[cfg(test)]
pub(super) fn the_source_replacements() -> u32 {
    SOURCE_REPLACEMENTS.load(Ordering::SeqCst)
}

/// How many registrations the Shell has been asked for (see `settle_the_registration`).
#[cfg(test)]
pub(super) fn the_registrations() -> u32 {
    REGISTRATIONS.load(Ordering::SeqCst)
}

/// Forget the registrations counted before a test (see `settle_the_registration`).
#[cfg(test)]
pub(super) fn clear_the_registrations() {
    REGISTRATIONS.store(0, Ordering::SeqCst);
}

/// Forget the placements counted before a test (see `place_the_source`).
#[cfg(test)]
pub(super) fn clear_the_source_placements() {
    SOURCE_PLACEMENTS.store(0, Ordering::SeqCst);
    SOURCE_REPLACEMENTS.store(0, Ordering::SeqCst);
}

/// How many registrations this run has let go of (see `let_the_registration_go`).
#[cfg(test)]
pub(super) fn the_unregistrations() -> u32 {
    UNREGISTRATIONS.load(Ordering::SeqCst)
}

/// Forget the unregistrations counted before a test (see `let_the_registration_go`).
#[cfg(test)]
pub(super) fn clear_the_unregistrations() {
    UNREGISTRATIONS.store(0, Ordering::SeqCst);
}

/// Whether the picture standing in a host is the source this road has adopted — the question a test
/// asks between a source's adoption and the swap that makes it the picture (see
/// `hand_the_picture_over`).
#[cfg(test)]
pub(super) fn the_standing_source_is_the_current_one() -> bool {
    let number = SOURCE.load(Ordering::SeqCst);
    THUMBNAIL
        .lock()
        .ok()
        .is_some_and(|held| held.registration != 0 && held.registered == number && number != 0)
}

/// Whether a source adopted to take the picture over is still waiting to be ready (see
/// `PendingSource`).
#[cfg(test)]
pub(super) fn a_source_is_waiting_to_take_the_picture() -> bool {
    PENDING_SOURCE
        .lock()
        .ok()
        .is_some_and(|pending| pending.is_some())
}

/// The client the picture standing is cropped by: the one placement's, never a live read of the source
/// (see `settle_the_registration`). It is what a seek's own line is read against (see
/// `note_the_seek_picture`).
#[cfg(test)]
pub(super) fn the_cropped_client() -> Option<(i32, i32)> {
    THUMBNAIL
        .lock()
        .ok()
        .and_then(|held| held.source)
        .map(|shape| shape.client)
}

/// A box as a trace and a placement carry it: origin, width and height.
type SourceBox = (i32, i32, i32, i32);

/// What a pin's source shows, fixed at the one placement: the client the ask named, and the film's own
/// shape fitted inside it (see `the_one_placement` and `the_thumbnail_box`).
#[derive(Clone, Copy, PartialEq, Debug)]
struct SourceShape {
    /// The window and the player the shape was placed for. **A shape is one film's**: the record out
    /// of which a pair is built is invalidated with the film, so no frame of the film that follows
    /// can be cropped by the shape of the one that ended (see `forget_the_shape_of_another_film`).
    source: isize,
    process: u32,
    /// The source's client, as the placement asked for it. **Never read again**: nothing asks the
    /// window for another box, so the crop's window is a constant for the pin's whole life.
    client: (i32, i32),
    /// The film's own shape, width over height, from the picture's own size — or nothing where it
    /// could not be read, which is the whole client into the whole box.
    aspect: Option<f64>,
}

impl SourceShape {
    /// Whether this shape was placed for the window and player of `source`: a shape is never read
    /// for a film other than the one it was placed for.
    fn belongs_to(self, source: HWND, process: u32) -> bool {
        self.source == source.0 as isize && self.process == process
    }
}

/// The film's own size at the moment a placement is asked for, and which of the two roads is asking.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum NativeSize {
    /// The picture's own size, from the player or the probe, as adoption asked it.
    Known((i32, i32)),
    /// The picture's own size as the road that watches the player answered it, after adoption had
    /// given up waiting: the same one placement, made late (see
    /// `place_the_source_when_the_player_answers`).
    Late((i32, i32)),
    /// Nothing yet: the placement waits for the player rather than guessing a box that would have to
    /// be corrected later.
    NotYet,
    /// Nothing the player or the probe will answer: the source is placed once at the largest box the
    /// road can justify and nothing asks the window for another one.
    Unobtainable,
}

impl NativeSize {
    /// The size a placement is asked for, or nothing where the road is still waiting.
    fn box_size(self, fallback: (i32, i32)) -> Option<(i32, i32)> {
        match self {
            NativeSize::Known(size) | NativeSize::Late(size) => Some(size),
            NativeSize::Unobtainable => Some(fallback),
            NativeSize::NotYet => None,
        }
    }

    /// The film's shape as the source's picture is fitted into its client, or nothing where the size
    /// is not the picture's own — an unobtainable size is a box the road guessed, not a shape.
    fn aspect(self) -> Option<f64> {
        match self {
            NativeSize::Known((width, height)) | NativeSize::Late((width, height)) => {
                Some(width as f64 / height.max(1) as f64)
            }
            NativeSize::NotYet | NativeSize::Unobtainable => None,
        }
    }
}

#[derive(Clone, Copy)]
struct Thumbnail {
    /// The host the Shell is drawing into, and the registration it answered with.
    host: isize,
    registration: isize,
    /// The window that registration draws. **It is not always the current source**: a source adopted
    /// to take the picture over draws nothing until it is ready, and the picture before it keeps the
    /// host until then (see `hand_the_picture_over`).
    registered: isize,
    /// The player whose window `registered` is. It is kept beside the window rather than read off the
    /// shape, because a source is registered before its one placement in the corner where the film's
    /// own size is not known yet — and the host is drawing that player either way (see
    /// `the_standing_source_process`).
    registered_process: u32,
    /// The rectangle pair the registration was last handed — the host's client it is drawn in, the
    /// picture's rectangle in the source and the source's own client — or nothing where none stands.
    /// **A change of either is a re-point** (see `settle_the_registration`).
    shown: Option<ThumbnailBox>,
    /// The picture the registered source shows: `None` until its one placement is made, and the same
    /// value for every frame after it (see `place_the_source`).
    source: Option<SourceShape>,
}

impl Thumbnail {
    const NOTHING: Self = Self {
        host: 0,
        registration: 0,
        registered: 0,
        registered_process: 0,
        shown: None,
        source: None,
    };
}

/// A film the road adopted without a size of its own, and the asks it owes the player until the size
/// is known (see `place_the_source_when_the_player_answers`).
#[derive(Clone, Copy)]
struct LatePlacement {
    source: isize,
    process: u32,
    /// When adoption found the source was not sized: what the wait `LATE_NATIVE_WAIT` is measured
    /// from.
    since: Instant,
    /// When the player was last asked, so a monitor that passes every few milliseconds asks a paced
    /// handful of times rather than on every pass.
    asked: Option<Instant>,
}

/// The film the road is still waiting to place, or nothing while none is.
static LATE_PLACEMENT: Mutex<Option<LatePlacement>> = Mutex::new(None);

/// Wait for the size of a film this adoption could not size (see `LatePlacement`).
fn arm_the_late_placement(source: HWND, process: u32) {
    if let Ok(mut armed) = LATE_PLACEMENT.lock() {
        *armed = Some(LatePlacement {
            source: source.0 as isize,
            process,
            since: Instant::now(),
            asked: None,
        });
    }
}

/// Give up waiting for a size: the film was placed, or there is no film of this one left.
fn forget_the_late_placement() {
    if let Ok(mut armed) = LATE_PLACEMENT.lock() {
        *armed = None;
    }
}

/// A source this road has adopted that is **not the picture in the host yet**.
///
/// **A new source may not take over before it is ready.** A relaunch — FFmpeg's player on every seek,
/// the copy of a film's subtitles landing late, mpv's own window made again — is another window, and a
/// registration moved onto it the moment it exists is a picture cropped by a client the one placement
/// did not ask for, drawn over the host's black while the player opens the file. So the registration
/// that is standing goes on drawing the source before it, and the swap is one step made when this one
/// can draw: its live client is the client the placement asked for, and its player's own video output
/// is configured for a frame (see `hand_the_picture_over`). The wait is bounded by the player-start
/// wait, and the trace says which of the two ended it.
#[derive(Clone, Copy)]
struct PendingSource {
    source: isize,
    process: u32,
    /// The one placement this source was given, once the film's own size is known.
    placed: Option<SourceShape>,
    /// When the source was adopted: the wait the swap is bounded by.
    since: Instant,
    /// The `not ready` answer last written and when, so a pass every hundred milliseconds is a line a
    /// run can be read from rather than a flood (see `note_the_source_not_ready`).
    not_ready: Option<((i32, i32), Instant)>,
    /// When the source was last sampled for a picture, and whether one has ever been seen: the sample
    /// is paced and a source that has drawn is not asked again (see `the_source_drew`).
    sampled: Option<Instant>,
    drew: bool,
}

/// The source adopted and not yet the picture, if one is (see `PendingSource`).
static PENDING_SOURCE: Mutex<Option<PendingSource>> = Mutex::new(None);

/// Wait for a source to be ready to take the picture over (see `PendingSource`).
fn arm_the_pending_source(source: HWND, process: u32) {
    if let Ok(mut pending) = PENDING_SOURCE.lock() {
        *pending = Some(PendingSource {
            source: source.0 as isize,
            process,
            placed: None,
            since: Instant::now(),
            not_ready: None,
            sampled: None,
            drew: false,
        });
    }
}

/// The source this road is waiting on, where it is the one named.
fn the_pending_source_for(source: isize, process: u32) -> Option<PendingSource> {
    PENDING_SOURCE
        .lock()
        .ok()
        .and_then(|pending| *pending)
        .filter(|pending| pending.source == source && pending.process == process)
}

/// Give up waiting for a source: the picture changed hands, or there is nothing left to hand it to.
fn forget_the_pending_source() {
    if let Ok(mut pending) = PENDING_SOURCE.lock() {
        *pending = None;
    }
}

/// Forget the wait only where it is still `source`'s and `process`'s: a source adopted while this one
/// was being handed over is the next wait, and it is not this call's to give up (see
/// `hand_the_picture_over`).
fn forget_the_pending_source_of(source: isize, process: u32) {
    if let Ok(mut pending) = PENDING_SOURCE.lock() {
        if pending.is_some_and(|pending| pending.source == source && pending.process == process) {
            *pending = None;
        }
    }
}

/// Keep the one placement a source that is waiting to take the picture over was given.
fn note_the_placement_of_a_pending_source(source: isize, process: u32, placed: SourceShape) {
    if let Ok(mut pending) = PENDING_SOURCE.lock() {
        if let Some(waiting) = pending
            .as_mut()
            .filter(|waiting| waiting.source == source && waiting.process == process)
        {
            waiting.placed = Some(placed);
        }
    }
}

/// Whether a host has a picture in it at all — the question `note_the_source` asks before it lets the
/// picture before this source go (see `hand_the_picture_over`).
///
/// **A still this app painted is a picture too**, and it is the one a relaunch's DC capture
/// leaves standing (see `stand_the_outgoing_picture_from_its_dc`): a source adopted over it waits its
/// turn exactly as it would behind a registration, rather than being registered before it can draw.
fn the_picture_is_standing() -> bool {
    a_still_stands()
        || THUMBNAIL
            .lock()
            .ok()
            .is_some_and(|held| held.registration != 0 && held.registered != 0)
}

/// Whether a source adopted to take the picture over is still waiting to be ready, for the player it
/// belongs to: the question the loop's retirement asks before it ends the player that is still the
/// picture on screen (see `video_retire::settle_video_retirement`).
///
/// **It is the same question the swap asks**, so a record nothing is waiting on any more — a source
/// this road has moved off, or one a registration already draws — is not a wait and answers no.
pub(super) fn a_handover_is_pending_for(process: u32) -> bool {
    let Some(pending) = PENDING_SOURCE.lock().ok().and_then(|pending| *pending) else {
        return false;
    };
    process != 0
        && pending.process == process
        && pending.source != 0
        && SOURCE.load(Ordering::SeqCst) == pending.source
        && THUMBNAIL
            .lock()
            .ok()
            .is_some_and(|held| held.registration != 0 && held.registered != pending.source)
}

/// Whether this player's source is the window the host is drawing: **the one step of a handover**, and
/// the answer that says the player it replaced has nothing of the pin's on screen any more.
pub(super) fn the_source_is_standing(pid: u32) -> bool {
    if !the_source_is_up(pid) {
        return false;
    }
    let number = SOURCE.load(Ordering::SeqCst);
    THUMBNAIL
        .lock()
        .ok()
        .is_some_and(|held| the_registration_draws(&held, number))
}

/// The same question asked by a reader that may not wait for it: the tick.
///
/// **The record is held by the pass that watches the player, and that pass asks the host for its own
/// client origin while it holds it** — a call into this very thread — so a tick that waited on the
/// record would be the loop stopped dead behind the thread it is meant to be driving (see
/// `place_the_source_when_the_player_answers` and `client_origin`). A record that cannot be taken at
/// once is therefore answered as a picture that has not changed hands yet, which is the answer that
/// costs one more tick and never one that acts: what asks this settles the same wait again on the
/// next turn (see `PinSwapHold::settle`).
pub(super) fn the_source_is_standing_try(pid: u32) -> bool {
    if !the_source_is_up(pid) {
        return false;
    }
    let number = SOURCE.load(Ordering::SeqCst);
    THUMBNAIL
        .try_lock()
        .ok()
        .is_some_and(|held| the_registration_draws(&held, number))
}

/// Whether the registration in the record draws the source `number`: the one step of a handover,
/// shared by the two ways of asking it (see `the_source_is_standing`).
fn the_registration_draws(held: &Thumbnail, number: isize) -> bool {
    held.registration != 0 && held.registered == number && number != 0
}

/// Whether any registration is standing in the pin's host, whichever player it draws: the state the
/// still probe reads beside the pixels, because `thumb=unregistered` is half of what proving the
/// still road takes (see `still_probe`).
pub(super) fn a_registration_stands() -> bool {
    THUMBNAIL
        .lock()
        .ok()
        .is_some_and(|held| held.registration != 0)
}

/// **The registration the pin's box is being drawn through, as this app believes it**: the source
/// window, the player it belongs to, the destination it was last handed, whether that window is on
/// the screen and whether a registration stands at all — one line, written at a release.
///
/// It is the same picture the still probe's own samples are read beside, at the one moment that
/// matters: a pin read black with this line saying a registration stood over a visible source is a
/// box nothing was drawn in, while the same black with `standing=0` is a registration this app never
/// had. `standing` is the record's own answer and not the Shell's (see `still_probe`).
pub(super) fn note_the_covering_registration() {
    let (registered, process, destination, standing) = THUMBNAIL
        .lock()
        .ok()
        .map(|held| {
            (
                held.registered,
                held.registered_process,
                held.shown.map(|shown| shown.destination),
                held.registration != 0 && held.registered != 0,
            )
        })
        .unwrap_or((0, 0, None, false));
    // Safety: the handle is the source this module adopted from a player's own window, and the
    // question is about the window rather than a change to it.
    let visible =
        registered != 0 && unsafe { IsWindowVisible(HWND(registered as *mut _)) }.as_bool();
    let destination = match destination {
        Some((left, top, right, bottom)) => {
            format!("{}x{}+{}+{}", right - left, bottom - top, left, top)
        }
        None => "none".to_string(),
    };
    note_pin_trace!(
        "covering registration: source {:?} pid {process} dest={destination} visible={} standing={}",
        HWND(registered as *mut _),
        u8::from(visible),
        u8::from(standing),
    );
}

/// **The player whose window the registration standing in `host` is drawing**, where one stands: the
/// picture that host has on screen, whoever began it.
///
/// It is deliberately not the newest pid. A relaunch records its replacement before that player has
/// made a window, and the player before it is **frozen** and is what the host goes on drawing until the
/// swap (see `hand_the_picture_over`), so a road that reads the newest pid as the picture sees a host
/// with a still frame in it and calls it empty — which is a seek's instant declared a hole, with the
/// z-order re-assert taken off the pin for the whole wait that follows (see
/// `pin_film_host::the_host_is_drawing`).
pub(super) fn the_standing_source_process(host: HWND) -> Option<u32> {
    THUMBNAIL.lock().ok().and_then(|held| {
        let standing = held.registration != 0
            && held.registered != 0
            && held.host == host.0 as isize
            // **A registration outlives the window it draws.** The record is what the host is asked
            // whether it has a picture in it (see `pin_film_host::the_host_is_drawing`), and a player
            // ended under a standing registration leaves it naming a window that is gone: the Shell
            // composes nothing, and the band painted for a picture that is there is a claim with
            // nothing behind it — the desktop showing through the card for the ticks between one
            // film's source going and the next one's registration standing.
            && unsafe { IsWindow(HWND(held.registered as *mut _)) }.as_bool();
        standing
            .then_some(held.registered_process)
            .filter(|pid| *pid != 0)
    })
}

/// Take a player's own window as the film the pin's host shows: out of the Shell's lists, at the
/// bottom of the desktop's order, and registered in the host that shows it.
///
/// **The style is put on once, at adoption** (see `source_ex_style`), and **the invisibility with it**:
/// `WS_EX_LAYERED` and `SetLayeredWindowAttributes(…, 1, LWA_ALPHA)` in the same pass, so the window is
/// composited at one step from nothing rather than covered by the host. Nothing is hidden to do it and
/// `WS_VISIBLE` is never cleared — a hidden source is a thumbnail the Shell stops drawing at all (see
/// the module note).
///
/// **The z-order insert is one post rather than a wait**, `SWP_ASYNCWINDOWPOS` to the bottom: the
/// window belongs to mpv's process, and a synchronous reorder is this thread waiting on mpv's — the
/// road that deadlocked this app once already on the child road. What a post
/// costs is the window standing on its own for the moment it takes mpv to pump it, which is behind
/// the hover the pin is being taken up from.
///
/// **`host` is the window the picture is drawn in**: the pin's film host, or the hover's on the road
/// a hover is shown through a thumbnail of its own (see `hover_film_host`). One source per live film,
/// and the destination is whatever host is showing it — the source's box does not depend on which.
///
/// Safety: `source` is a window the caller found by `pid` and this app did not make.
pub(super) unsafe fn adopt_the_source(source: HWND, pid: u32, host: HWND) {
    // **A window the still was taken of is not adopted again**: its player is ended the moment the
    // still is painted, and this pass can arrive in between with the window still a window (see
    // `STILL_PID`).
    if STILL_PID.load(Ordering::SeqCst) == pid {
        return;
    }
    let ex = unsafe { GetWindowLongPtrW(source, GWL_EXSTYLE) };
    let invisible = the_source_is_invisible();
    let _ = unsafe { SetWindowLongPtrW(source, GWL_EXSTYLE, source_ex_style(ex, invisible)) };
    if invisible {
        // Safety: as above, and the alpha is this app's own to give a window of another process's. The
        // one step from nothing, `LWA_ALPHA`, is what keeps the window drawn (and therefore
        // thumbnailable) while nothing of it reaches the screen.
        let applied =
            unsafe { SetLayeredWindowAttributes(source, COLORREF(0), INVISIBLE_ALPHA, LWA_ALPHA) };
        note_pin_trace!(
            "source invisible {:?} layered alpha {} applied ok {} (RHP_THUMBNAIL_INVISIBLE)",
            source,
            INVISIBLE_ALPHA,
            applied.is_ok() as u8,
        );
    } else {
        note_pin_trace!(
            "source invisible {:?} off: the source is a covered window (RHP_THUMBNAIL_INVISIBLE=0)",
            source,
        );
    }
    lower_the_source(source);
    // **The new source has taken the picture, so the band a step held is the setting's again**:
    // the outgoing film is behind this one from here, and the pin goes back to being the ordinary
    // window a pin that is not Always On Top is (see `PIN_STEP_BAND_HELD`).
    unsafe { pin_lifecycle::give_the_step_band_back() };

    // **A window that made itself the foreground gives it back.** mpv shows its own window when it has
    // opened the file, which is after the click that made the pin, and the style above stops it taking
    // the keyboard but not from having taken it already: the window the launch recorded as the one in
    // front is asked for it here — the pin's own window where the user was in the pin, the listing
    // where the user was working in it, and, where that window has gone, the ordinary window behind
    // (see `pin_lifecycle::a_players_foreground_goes_to`). A window of this app's own was the
    // foreground a moment ago, so the ask is one Windows grants.
    unsafe { pin_lifecycle::give_the_foreground_back_from(source) };

    note_the_source(source, pid);

    // **Where the player put the window it made for itself, before this app asks it for anything.**
    // The source is launched on the desktop (see `the_source_launch_origin`) and this is the line
    // that says where it was born and which player made it; the one placement's own line follows it
    // with what this app asked for instead. **`player` is the engine a run is read by** — the image
    // this app recorded when it began the player — so an FFmpeg film and an mpv one are told apart
    // here and nowhere else in the trace's per-frame lines.
    let mut created = RECT::default();
    if unsafe { GetWindowRect(source, &mut created) }.is_ok() {
        note_pin_trace!(
            "source adopted {:?} pid {} player {} born {}x{}+{}+{}",
            source,
            pid,
            video_process_image(),
            created.right - created.left,
            created.bottom - created.top,
            created.left,
            created.top,
        );
    }

    // **The one placement the source ever takes is made here**, in the adoption pass and before the
    // registration draws anything: the film's own picture size at the destination host's corner, asked
    // once and never asked for again (see `place_the_source`).
    //
    // **A film whose own size the player has not answered is waited for rather than guessed at**, and
    // the road that watches the player goes on asking: the placement is made as soon as the player
    // answers, still once and never a second time (see `place_the_source_when_the_player_answers`).
    // The wait happens before the lock is taken, so a read and a sleep are never spent holding what
    // the pin's own resize steps want.
    let placed = THUMBNAIL
        .lock()
        .ok()
        .and_then(|held| held.source)
        // A shape recorded for another film is not this one's: the placement is this film's first.
        .filter(|shape| shape.belongs_to(source, pid));
    let native = the_films_native_size(pid).map_or(NativeSize::NotYet, NativeSize::Known);
    let placed = place_the_source(source, host, pid, placed, native);
    let pending = the_pending_source_for(source.0 as isize, pid).is_some();
    if let Some(placed) = placed.filter(|_| pending) {
        // **The picture standing keeps drawing**: this source's placement is kept for the swap and not
        // for this pass (see `hand_the_picture_over`).
        note_the_placement_of_a_pending_source(source.0 as isize, pid, placed);
    }
    if let Ok(mut held) = THUMBNAIL.lock() {
        if !pending {
            if let Some(placed) = placed {
                held.source = Some(placed);
            }
        }
        // **A film adopted without a size is one the road is still waiting for**, and what waits is
        // the pass that watches the player rather than a second placement made here.
        let sized = if pending {
            the_pending_source_for(source.0 as isize, pid)
                .is_some_and(|pending| pending.placed.is_some())
        } else {
            held.source
                .is_some_and(|shape| shape.belongs_to(source, pid))
        };
        if sized {
            forget_the_late_placement();
        } else {
            arm_the_late_placement(source, pid);
        }
        // **A still this app painted is the picture until a source can draw over it**: registering
        // this one here would compose a window that has drawn nothing over that still, which is the
        // blank the paint exists to keep out (see `stand_the_outgoing_picture_from_its_dc`).
        if !(pending && a_still_stands()) {
            settle_the_registration(&mut held, host);
        }
    }
}

/// Record the window the Shell is to draw: another player, or the same player's window made again,
/// is another picture — which **takes over only when it can draw it**.
///
/// **The registration that is standing is not taken down here.** A relaunch produces a new window
/// while the one it replaces is still the picture on screen (that is what the player being parked
/// rather than ended is for), and unregistering before the new source can draw is the host's own black
/// erase seen through the pin for as long as the player takes to open the file — the flash a subtitle
/// copy and a seek both reported. So the new source is recorded, the old registration keeps drawing,
/// and `hand_the_picture_over` makes the swap in one step when the new one is ready. With nothing
/// standing there is nothing to keep, and the record is reset as it always was.
pub(super) fn note_the_source(source: HWND, pid: u32) {
    let number = source.0 as isize;
    if SOURCE.load(Ordering::SeqCst) == number && SOURCE_PROCESS.load(Ordering::SeqCst) == pid {
        return;
    }

    if the_picture_is_standing() {
        let drawn = THUMBNAIL
            .lock()
            .ok()
            .map(|held| held.registered)
            .unwrap_or(0);
        arm_the_pending_source(source, pid);
        SOURCE.store(number, Ordering::SeqCst);
        SOURCE_PROCESS.store(pid, Ordering::SeqCst);
        note_where_the_source_stands(source);
        if drawn != 0 {
            note_pin_trace!(
                "thumbnail source waiting {:?} pid {}: the picture standing is {:?} and draws until it \
                 is ready",
                source,
                pid,
                HWND(drawn as *mut _),
            );
        } else {
            note_pin_trace!(
                "thumbnail source waiting {:?} pid {pid}: the still painted for a relaunch draws until \
                 it is ready",
                source,
            );
        }
        return;
    }

    let registration = THUMBNAIL.lock().ok().map(|held| held.registration);
    if let Ok(mut held) = THUMBNAIL.lock() {
        *held = Thumbnail::NOTHING;
    }
    if let Some(registration) = registration {
        let_the_registration_go(registration);
    }
    forget_the_pending_source();
    REGISTERED_HOST.store(0, Ordering::SeqCst);
    SOURCE.store(number, Ordering::SeqCst);
    SOURCE_PROCESS.store(pid, Ordering::SeqCst);
    note_where_the_source_stands(source);
}

/// Write where a source this road has just begun to draw stands, as the ask that follows is compared
/// against: **a window recorded here has nothing in flight**, so its own answer is the truth rather
/// than a message that may not have landed (see `set_the_source_visible`).
fn note_where_the_source_stands(source: HWND) {
    // Safety: the handle is one a player made and this module has just recorded, and the question is
    // about the window rather than a change to it.
    let up = unsafe { IsWindowVisible(source) }.as_bool();
    if let Ok(mut held) = THE_SOURCE_SHOW.lock() {
        *held = Some((source.0 as isize, up));
    }
}

/// The window the Shell is drawing the pin's film from, where there is one: the window of mpv's the
/// pin's picture is a thumbnail of, which is the third window a press can land in on this road (see
/// `the_band_was_placed` and `key_input::pin_owns_caret`).
pub(super) fn the_source_window() -> Option<HWND> {
    let number = SOURCE.load(Ordering::SeqCst);
    (number != 0).then_some(HWND(number as *mut _))
}

/// Whether this player's window is up as the pin's picture: still a window, and still one of the
/// player this app records rather than one given the number of a window it replaced.
///
/// The record is the player, so the live window's own owner is asked only of the binary: a test has
/// one process and the player it stands for is a number, so its source is a real window of its own
/// (see `the_source_window`).
pub(super) fn the_source_is_up(pid: u32) -> bool {
    let number = SOURCE.load(Ordering::SeqCst);
    if number == 0 || pid == 0 || SOURCE_PROCESS.load(Ordering::SeqCst) != pid {
        return false;
    }
    let source = HWND(number as *mut _);
    // Safety: the handle is one this module adopted from a player's own window, and the questions are
    // about the window rather than changes to it.
    unsafe {
        if !IsWindow(source).as_bool() {
            return false;
        }
        #[cfg(test)]
        return true;
        #[cfg(not(test))]
        {
            let mut owner = 0u32;
            GetWindowThreadProcessId(source, Some(&mut owner));
            owner == pid
        }
    }
}

/// Put the invisibility back where something took it off the source: the style monitor's pass over a
/// player's window asks this, and it is the only caller.
///
/// **A video output reinitialised is a window that can come back without the style** — mpv recreates
/// its window on a VO reinit, and a new window is one this app has not adopted (so it is found and
/// adopted by the monitor's own pass) while a window that is the same handle can still have its
/// extended style reset under it. Either way what has to be true afterwards is one fact, and it is
/// cheap to ask: a window without `WS_EX_LAYERED` is a film standing over the pin. Every re-application
/// is traced, so a run says how often the player un-did it instead of leaving it to be guessed.
///
/// Safety: the handle is the source this module adopted, if there is one, and the two calls only put a
/// style on a window of another process's — the same pair the adoption pass makes.
pub(super) unsafe fn keep_the_source_invisible() {
    if !the_source_is_invisible() {
        return;
    }
    let Some(source) = the_source_window() else {
        return;
    };
    let ex = unsafe { GetWindowLongPtrW(source, GWL_EXSTYLE) };
    if ex & (WS_EX_LAYERED.0 as isize) != 0 {
        return;
    }
    let _ = unsafe { SetWindowLongPtrW(source, GWL_EXSTYLE, source_ex_style(ex, true)) };
    let applied =
        unsafe { SetLayeredWindowAttributes(source, COLORREF(0), INVISIBLE_ALPHA, LWA_ALPHA) };
    note_pin_trace!(
        "source invisible {:?} re-applied: the layered style had gone, alpha {} ok {}",
        source,
        INVISIBLE_ALPHA,
        applied.is_ok() as u8,
    );
}

/// Point the Shell's thumbnail at the host's client area — **the whole of it, filled by the picture**
/// (see `the_thumbnail_box`) — and nothing else.
///
/// **The source is not touched here, which is the whole of what makes a step cheap.** Its box, its
/// client and the film's shape were fixed at the one placement (see `place_the_source`), so a resize,
/// a maximize, a restore, a zoom and a walk are all the same one thing: the destination rectangle and
/// the crop the box's shape asks for. Nothing is asked of a window of mpv's, nothing is read of the
/// player, and nothing on this road is cross-process at all.
///
/// **A frame is only a re-point where the pair changed**: a step of a drag moves the host's client
/// every time, and a tick that finds the same box hands the Shell nothing (see `the_registration_wants`).
pub(super) fn the_band_was_placed(host: HWND) {
    if SOURCE.load(Ordering::SeqCst) == 0 || a_still_stands() {
        return;
    }
    let Ok(mut held) = THUMBNAIL.lock() else {
        return;
    };
    settle_the_registration(&mut held, host);
}

/// **The one placement a source ever takes**: the film's own picture size at the corner of the host
/// showing it — the pin's, or the hover's the pin is taken up from — asked for once, at adoption, or
/// by the road that watches the player as soon as the film's size is known (see
/// `place_the_source_when_the_player_answers`), and never asked for again until the window is taken
/// down.
///
/// **What the record it leaves says is why a resize cannot flash.** The client the ask named and the
/// shape the film came with are kept beside the registration (see [`SourceShape`]), and every frame
/// after this one is the destination rectangle and a crop of *that* constant — no `SetWindowPos` on a
/// window of mpv's, no read of a player that has just been asked to resize, and nothing to disagree
/// with the surface the Shell is sampling (see the module note).
///
/// **A second ask is the bug, so it is refused and traced rather than made.** `source re-placed` is a
/// line that must never appear in a run: it means something on the pin's road thought the source's box
/// was its own to change again.
///
/// **The ask is posted**, like every other call on a window of another process's here. What is
/// recorded is the box the ask names, so the crop is the picture that ask describes whether or not
/// the player has taken it yet — which is the one thing that keeps this road out of the asynchronous
/// gap between a window taking a box and its output configuring for it.
///
/// **Nothing is held while the player is waited for**: the reads and the sleep that may follow them
/// are this function's own, and the record is written once when the answer is in (see
/// `adopt_the_source`).
///
/// `native` is what the road knows about the film's own size, and it is also what tells the trace
/// which of the roads this placement is: the adoption pass, the late one after the player answered,
/// a film still to be sized, or a film the road gave its wait to and was answered nothing.
fn place_the_source(
    source: HWND,
    host: HWND,
    pid: u32,
    placed: Option<SourceShape>,
    native: NativeSize,
) -> Option<SourceShape> {
    if placed.is_some() {
        // **The line that must never appear in a run**, and the counter a test reads when it checks
        // that no road can reach a second placement without it being said out loud.
        note_pin_trace!(
            "source re-placed {:?} host {:?}: a source takes one placement and this is a second ask",
            source,
            host,
        );
        #[cfg(test)]
        SOURCE_REPLACEMENTS.fetch_add(1, Ordering::SeqCst);
        return None;
    }
    let origin = (unsafe { client_origin(host) })?;
    let wanted = native.box_size(the_max_track());
    let Some((x, y, width, height)) =
        the_one_placement(None, Some(origin), wanted, the_max_track())
    else {
        if native == NativeSize::NotYet {
            note_pin_trace!(
                "thumbnail source native unknown {:?}: no box asked yet, the player has not answered",
                source,
            );
        }
        return None;
    };
    let asked = size_the_source(source, (x, y, width, height));
    let aspect = native.aspect();
    let (left, top, picture_width, picture_height) = a_picture_fitted_to(aspect, (width, height));
    match native {
        // **A film the road could not size is placed once at the largest box it can justify**: the
        // picture is upscaled into that box and cropped by the box's shape like every other, and the
        // line says the size was never read so a run cannot mistake the box for the film's own.
        NativeSize::Unobtainable => note_pin_trace!(
            "thumbnail source native unobtainable {:?}: no answer from the player or the probe, \
             placed once at the largest box {}x{}+{}+{} ok {}",
            source,
            width,
            height,
            x,
            y,
            asked as u8,
        ),
        // **The answer that came after adoption gave up waiting** is this road's own line: the same
        // one placement, and the film's own size it names.
        NativeSize::Late(_) => note_pin_trace!(
            "thumbnail source native late {}x{}+{}+{} aspect {:.4} placed once {:?} at {}x{}+{}+{} \
             ok {}",
            picture_width,
            picture_height,
            left,
            top,
            aspect.unwrap_or(0.0),
            source,
            width,
            height,
            x,
            y,
            asked as u8,
        ),
        NativeSize::Known(_) => note_pin_trace!(
            "thumbnail source native {}x{}+{}+{} aspect {:.4} placed once {:?} at {}x{}+{}+{} ok {}",
            picture_width,
            picture_height,
            left,
            top,
            aspect.unwrap_or(0.0),
            source,
            width,
            height,
            x,
            y,
            asked as u8,
        ),
        // A film still waiting for its size never reaches the ask above.
        NativeSize::NotYet => return None,
    }
    #[cfg(test)]
    SOURCE_PLACEMENTS.fetch_add(1, Ordering::SeqCst);
    Some(SourceShape {
        source: source.0 as isize,
        process: pid,
        client: (width, height),
        aspect,
    })
}

/// **The one placement a film that adoption could not size is owed, made as soon as the player
/// answers** — and the source placed once at the largest box the road can justify where the film has
/// had its wait and was answered nothing.
///
/// The monitor that keeps the player's window styled passes here every few milliseconds (see
/// `ffplay_window::ensure_noactivate_monitor`), which is the road that already waits on the player: it
/// asks once per `LATE_NATIVE_EVERY` rather than on every pass, and it asks nothing at all while a
/// resize gesture is live — the one invariant is that no resize gesture may have the source sized
/// under it (see `pin_film_host::a_resize_drag_is_live`). A drag that ends before the size is known
/// does not stop the placement; a drag in flight does.
///
/// **It is the same placement and not a second one.** A source that has a shape is left alone, and
/// the wait is given up with the placement, so nothing asks this window for another box again.
pub(super) fn place_the_source_when_the_player_answers(host: HWND) {
    if pin_film_host::a_resize_drag_is_live() {
        return;
    }
    let number = SOURCE.load(Ordering::SeqCst);
    let process = SOURCE_PROCESS.load(Ordering::SeqCst);
    if number == 0 || process == 0 {
        forget_the_late_placement();
        return;
    }
    let Ok(mut held) = THUMBNAIL.lock() else {
        return;
    };
    // **The source this pass is about is the one the road is on**: the source waiting to take the
    // picture over where a swap is waiting, and the standing one otherwise. A source that has its one
    // placement already is left alone.
    let pending = the_pending_source_for(number, process);
    let sized = match pending {
        Some(pending) => pending.placed.is_some(),
        None => held
            .source
            .is_some_and(|shape| shape.belongs_to(HWND(number as *mut _), process)),
    };
    if sized {
        forget_the_late_placement();
        return;
    }
    let Ok(mut armed) = LATE_PLACEMENT.lock() else {
        return;
    };
    let Some(waiting) =
        (*armed).filter(|waiting| waiting.source == number && waiting.process == process)
    else {
        return;
    };
    let waited = waiting.since.elapsed();
    if !a_late_placement_is_due(waited, waiting.asked.map(|asked| asked.elapsed())) {
        return;
    }
    if let Some(waiting) = armed.as_mut() {
        waiting.asked = Some(Instant::now());
    }
    drop(armed);

    let source = HWND(number as *mut _);
    let native = match the_players_picture_size(process).or_else(|| the_films_own_size(process)) {
        Some(size) => NativeSize::Late(size),
        // **A film that has had its wait and answered nothing is placed once at the largest box the
        // road can justify**, rather than left at the box the player chose for itself: that box is
        // not the film's picture and nothing on this road ever reads it again.
        None if waited >= LATE_NATIVE_WAIT => NativeSize::Unobtainable,
        None => return,
    };
    let placed = place_the_source(source, host, process, None, native);
    if let Some(placed) = placed {
        if pending.is_some() {
            // The picture standing keeps drawing; this placement is the swap's (see
            // `hand_the_picture_over`).
            note_the_placement_of_a_pending_source(number, process, placed);
        } else {
            held.source = Some(placed);
        }
        forget_the_late_placement();
    }
    settle_the_registration(&mut held, host);
}

/// Hand the picture from the source that is drawing it to the one that replaced it, **once the new one
/// can draw the frame it was placed for**.
///
/// **The one visible swap of a relaunch, and it is one step.** Until now the registration standing in
/// the host draws the window before this one, so the pin has a picture for the whole of the wait and
/// none of it is the host's black erase — and no frame is cropped by a client the source was not placed
/// at. Ready is `the_source_can_draw` and `the_source_drew`: the client the one placement asked for is
/// the window's live client, and it has a frame to show — its player's video output is configured for
/// one where the player answers over a pipe, and a few points of its own window have been read where it
/// answers nothing at all.
///
/// **It is paced by the pass that already watches the player** (`ffplay_window`), so it costs a client
/// read and at most one bounded property read per pass while a swap is waiting, and nothing at all once
/// one is not. A resize gesture holds it off, the way it holds off the placement, so no gesture ever
/// has the picture change hands under it.
///
/// **The new registration is handed its pair before the old one is let go of**, so the Shell is never
/// asked to draw the host with nothing in it. The player being replaced has nothing of the pin's on
/// screen from the line after, and the loop's own retirement ends it (see
/// `video_retire::settle_video_retirement`, which waits for this swap rather than for a window).
///
/// **A source that never reaches its own client is not allowed to blank the pin**: the old picture goes
/// on drawing, `thumbnail source not ready` is written, and at the player-start bound the picture is
/// handed over anyway with the line that says so — the one answer that never exposes the host's black
/// for a handover.
pub(super) fn hand_the_picture_over(host: HWND) {
    if pin_film_host::a_resize_drag_is_live() {
        return;
    }
    let Some(pending) = PENDING_SOURCE.lock().ok().and_then(|pending| *pending) else {
        return;
    };
    let current = SOURCE.load(Ordering::SeqCst);
    if pending.source == 0 || current != pending.source {
        // The record names a window this road has left: there is nothing to hand over.
        forget_the_pending_source_of(pending.source, pending.process);
        return;
    }
    let source = HWND(pending.source as *mut _);
    // Safety: the handle is one a player made and this module adopted, and the question is about it.
    if !unsafe { IsWindow(source) }.as_bool() {
        forget_the_pending_source_of(pending.source, pending.process);
        return;
    }
    // A source still waiting for its own size has nothing to be ready for yet (see
    // `place_the_source_when_the_player_answers`).
    let Some(shape) = pending.placed else {
        return;
    };
    let live = unsafe { client_size(source) };
    let sized = the_source_can_draw(live, shape, pending.process);
    // **A source with a pipe is asked its own video output and a source without one is asked its own
    // pixels**, which is the whole of what "it has drawn" can mean for a player that answers nothing
    // (see `the_source_drew`).
    let ready = sized && the_source_drew(pending, source, shape.client);
    let waited = pending.since.elapsed();
    if !ready {
        note_the_source_not_ready(pending, live);
        if waited < Duration::from_secs(VIDEO_START_WAIT_SECS) {
            return;
        }
        if sized {
            // **The box is the one thing that is right**: this is a source the Shell would draw if
            // there were a picture in it, and the picture has not come (a film whose opening frame is
            // black, or a window whose surface a sample cannot reach). Changing hands anyway is the
            // one answer that never leaves the old picture to a player that is ending.
            note_pin_trace!(
                "thumbnail source never drew after {}ms, swapping anyway",
                waited.as_millis(),
            );
        } else {
            note_pin_trace!(
                "thumbnail source still not ready after {}ms: the picture changes hands anyway rather \
                 than leaving the one before it to a player that is ending",
                waited.as_millis(),
            );
        }
    }

    let Ok(mut held) = THUMBNAIL.lock() else {
        return;
    };
    let old = *held;
    // The new source's pair first: nothing of the old registration is touched until the Shell has
    // taken the new one.
    held.registration = 0;
    held.registered = 0;
    held.shown = None;
    held.source = Some(shape);
    settle_the_registration(&mut held, host);
    if held.registration == 0 {
        // **A registration the Shell refused is not a swap**: the picture standing stays, and the pass
        // asks again (see `settle_the_registration`).
        *held = old;
        REGISTERED_HOST.store(old.host, Ordering::SeqCst);
        return;
    }
    let_the_registration_go(old.registration);
    let swapped = held.registered;
    drop(held);
    // **The still is covered from the line above**, and it goes with the registration that took the
    // picture over: the swap needs no clear, because the Shell composes this registration over the
    // still this app painted (see `stand_the_outgoing_picture_from_its_dc`).
    STILL_HOST.store(0, Ordering::SeqCst);
    STILL_PID.store(0, Ordering::SeqCst);
    forget_the_pending_source_of(pending.source, pending.process);
    // **The line names both players, not just the two windows**: a handover made onto a dead window is
    // invisible in a pair of handles, and a run should never have to read a pid off anything else to
    // see it.
    note_pin_trace!(
        "thumbnail swapped old {:?} pid {} new {:?} pid {} after {}ms",
        HWND(old.registered as *mut _),
        old.source.map_or(0, |shape| shape.process),
        HWND(swapped as *mut _),
        pending.process,
        waited.as_millis(),
    );
    // **The host's retained cover goes with the registration that covers it**: from here the Shell
    // draws the replacement over that box, and a cover left standing would be painted again by the
    // next erase or resize (see `pin_film_host::put_the_cover_away`).
    pin_film_host::put_the_cover_away();
    // **The band is painted for the picture it now stands over, and this is the one moment that
    // owes it**: no frame of this app's moves behind a film a player draws, so a band painted for
    // the wait before this step would keep that paint until some other kind of redraw happened to
    // come along — which for a film is never (see `pin_film_host::the_band_owes_a_repaint`).
    pin_film_host::note_the_band_has_a_picture();
    pin_film_host::the_band_owes_a_repaint();
    // **And the still probe's measurement closes on the same step**: one thing happened, so the swap
    // and the answer to what the pin showed waiting for it are adjacent lines (see `still_probe`).
    super::still_probe::note_swapped();
    // **The replacement is the picture, so the relaunch that began it is over.** Until this line the
    // record is a player a newer gesture may kill before it publishes; from here it is the source the
    // host draws, and killing it is the black band (see `pin_playback::settle_the_published_relaunch`).
    pin_playback::settle_the_published_relaunch(pending.process);
    // **And the seek's own numbers where a seek is what asked for this source**, so the frame a seek
    // lands on is one line rather than an argument (see `note_the_seek_picture`). The reason is taken
    // rather than read, so a swap for a file step is not written as a seek that never happened.
    if pin_playback::take_the_last_restart_reason() == Some(pin_playback::RestartReason::Seek) {
        note_the_seek_picture(
            pending.process,
            pin_playback::pinned_playhead().unwrap_or(0.0),
        );
    }
}

/// Whether a source adopted to take the picture over can draw the frame it was placed for: the client
/// the one placement asked for is the client the window has, and — where the player answers over a
/// pipe — its video output is configured for a frame.
///
/// **This is the presented predicate without the second**, and the difference is what makes it usable
/// here: a replacement is begun at the second a seek or a cue asked for, and this pass has no second to
/// hold it to (see `pin_film_host::the_hosted_frame_is_presented`). A player with no pipe answers
/// nothing and its window is the whole of its picture, exactly as `the_hosted_frame_is_up` reads it.
fn the_source_can_draw(live: Option<(i32, i32)>, shape: SourceShape, process: u32) -> bool {
    live == Some(shape.client) && pin_film_host::the_players_output_is_configured(process)
}

/// Whether a source with no pipe has drawn anything into the client it was placed at: **the second
/// half of readiness for the one engine that cannot report a frame**, sampled and latched.
///
/// FFmpeg's player answers no property at all (see `mpv_ipc::answers_properties`), so `the_source_can_draw`
/// is only "the window is the size we asked for" for it — and that window is the right size and empty for
/// the whole of a decode, which is the swap onto a blank band a seek and a subtitle copy were measured
/// with. What separates a sized window from a drawn one is what is in it, so a few points inside the
/// client are read through the window's own DC until one of them is not black.
///
/// **It is paced and it is latched**: the monitor asks every few milliseconds and a read of another
/// process's window per pass is the per-frame cross-process work this road refuses, so the sample is one
/// per `SOURCE_DREW_EVERY` and a source that has drawn is never sampled again for its life. A drag holds
/// the whole handover off, so nothing here runs under one (see `hand_the_picture_over`).
///
/// The answer is written down once, when a source draws and nowhere else: a run reads how long the
/// picture before it stood from one line rather than from a gap between two.
fn the_source_drew(pending: PendingSource, source: HWND, client: (i32, i32)) -> bool {
    // A player that answers over a pipe is asked its own video output instead of its pixels (see
    // `the_source_can_draw`), and a source that has drawn once has drawn.
    if pending.drew || mpv_ipc::answers_properties(pending.process) {
        return true;
    }
    let now = Instant::now();
    if pending
        .sampled
        .is_some_and(|sampled| now.duration_since(sampled) < SOURCE_DREW_EVERY)
    {
        return false;
    }
    let drew = the_source_sample(source, client);
    if let Ok(mut record) = PENDING_SOURCE.lock() {
        if let Some(waiting) = record.as_mut().filter(|waiting| {
            waiting.source == pending.source && waiting.process == pending.process
        }) {
            waiting.sampled = Some(now);
            waiting.drew |= drew;
        }
    }
    if drew {
        note_pin_trace!(
            "thumbnail source drew after {}ms",
            pending.since.elapsed().as_millis(),
        );
    }
    drew
}

/// The sample itself: the machine's own read of the source's pixels, or the answer a test stood in for
/// it (see `DREW_STAND_IN`).
fn the_source_sample(source: HWND, client: (i32, i32)) -> bool {
    #[cfg(test)]
    if let Ok(held) = DREW_STAND_IN.lock() {
        if let Some(answer) = *held {
            return answer;
        }
    }
    // Safety: the handle is one this module adopted from a player's own window.
    unsafe { the_source_has_lit_pixels(source, client) }
}

/// Whether the picture the source of this player holds is lit, read through the player's own window
/// now rather than at the monitor's pace.
///
/// It is the drawn probe (`the_source_has_lit_pixels`) asked at one moment, and its one caller is a
/// relaunch's freeze: after a player is suspended, **this is what says whether the frame it was
/// showing is still there for the Shell to draw** — `ok` for a window that kept its picture, `black`
/// for one with nothing in it, which no registration of that window can bring back (see
/// `pin_playback::freeze_the_outgoing_picture`). `None` where this player has no source window to
/// read, which is a film that has drawn nothing yet.
///
/// It reads a few pixels of the window's own DC, which is not a wait on the player's threads: a
/// suspended player is asked nothing, and the answer is the surface it left behind.
pub(super) fn the_sources_picture_is_lit(pid: u32) -> Option<bool> {
    if !the_source_is_up(pid) {
        return None;
    }
    let source = the_source_window()?;
    // Safety: the handle is the source this module adopted from a player's own window, and the calls
    // read a few of its pixels or report that they cannot.
    let client = unsafe { client_size(source) }.unwrap_or((0, 0));
    Some(the_source_sample(
        source,
        (client.0.max(1), client.1.max(1)),
    ))
}

/// The same window read as what it is rather than what it draws: **whether it is on the screen, and
/// the two extended styles that decide whether the Shell can compose it at all** — `visible`,
/// `layered`, `topmost` — or nothing where this player has no source window, which is a reading no
/// window answered.
///
/// It is `the_sources_picture_is_lit`'s companion and nothing else asks it: a pin read black with its
/// source still lit is a picture the compositor stopped drawing, and the same black with the source's
/// own window black is a player that stopped drawing — two answers the pin's pixels alone cannot tell
/// apart (see `still_probe`).
pub(super) fn the_sources_window_facts(pid: u32) -> Option<(bool, bool, bool)> {
    if !the_source_is_up(pid) {
        return None;
    }
    let source = the_source_window()?;
    // Safety: the handle is the source this module adopted from a player's own window, and both
    // questions are about the window rather than changes to it.
    unsafe {
        let visible = IsWindowVisible(source).as_bool();
        let style = GetWindowLongPtrW(source, GWL_EXSTYLE) as u32;
        Some((
            visible,
            style & WS_EX_LAYERED.0 != 0,
            style & WS_EX_TOPMOST.0 != 0,
        ))
    }
}

/// Whether a few pixels of the source's own window are lit: **the only "it has drawn" there is for a
/// player that answers nothing**, and what `the_source_drew` paces.
///
/// The points are a handful well inside the client — corners are where a player's own letterbox or
/// border stands — and one lit sample is enough, so a film's first frame answers on the first pass it
/// is drawn. A frame that really is black everywhere (a fade, black leader) is the bound's case rather
/// than this one's, and the bound swaps anyway (see `hand_the_picture_over`).
///
/// shortcut: a window whose picture is presented through a swapchain whose redirection surface is not
/// updated answers black here, so such a player waits the start bound and is swapped at it rather than
/// on its first frame. Confirm the sample reaches a live FFmpeg window before trusting a run's
/// `thumbnail source drew` timing; upgrade it to a read of the composed picture only if it does not.
pub(super) unsafe fn the_source_has_lit_pixels(source: HWND, client: (i32, i32)) -> bool {
    let (width, height) = (client.0.max(1), client.1.max(1));
    let points = [
        (width / 4, height / 4),
        (width / 2, height / 4),
        (width / 2, height / 2),
        (width / 4, height * 3 / 4),
        (width * 3 / 4, height * 3 / 4),
    ];
    // Safety: the handle is a window this module adopted from a player's own, and the calls read a few
    // of its pixels or report that they cannot; the context is given back before this answers.
    unsafe {
        let dc = GetDC(source);
        if dc.is_invalid() {
            return false;
        }
        let lit = points.iter().any(|&(x, y)| {
            let pixel = GetPixel(dc, x, y);
            pixel.0 != CLR_INVALID && pixel.0 & 0x00FF_FFFF != 0
        });
        ReleaseDC(source, dc);
        lit
    }
}

/// Write a handover's `not ready` answer: the source's live client, the client the one placement asked
/// for, and the same line once a second while it stands. Written on change and never per frame.
fn note_the_source_not_ready(pending: PendingSource, live: Option<(i32, i32)>) {
    let live = live.unwrap_or((0, 0));
    let asked = pending.placed.map_or((0, 0), |shape| shape.client);
    let now = Instant::now();
    let due = {
        let Ok(mut record) = PENDING_SOURCE.lock() else {
            return;
        };
        let Some(waiting) = record.as_mut().filter(|waiting| {
            waiting.source == pending.source && waiting.process == pending.process
        }) else {
            return;
        };
        let due = waiting.not_ready.is_none_or(|(seen, at)| {
            seen != live || now.duration_since(at) >= Duration::from_secs(1)
        });
        if due {
            waiting.not_ready = Some((live, now));
        }
        due
    };
    if due {
        note_pin_trace!(
            "thumbnail source not ready cli {}x{} asked {}x{}",
            live.0,
            live.1,
            asked.0,
            asked.1,
        );
    }
}

/// **One bounded read of what the engine reports drawing, after a seek** — and the line a run is
/// convicted or cleared by: the source window, the client the Shell is sampling (`cli`), the picture
/// the player says it is drawing (`pic`), and the client the one placement asked for (`asked`).
///
/// **The three numbers are the whole of the smaller-than-the-box frame.** A `cli` that is not `asked`
/// is a surface the crop was not built for: the Shell scales the film out of a window of another size,
/// which is a next frame drawn smaller than the pin's film box. A `pic` that *was read* and is not
/// `asked` is a film the player is fitting inside its own window. Either is a difference the line
/// names, so a seek that flashed and a seek that did not are told apart by what they wrote and not by
/// what was remembered.
///
/// **`pic none` is nothing read, and never a mismatch.** FFmpeg's player answers no property at all,
/// so the picture of the frame it has drawn cannot be read from it; a seek onto one was written as a
/// `pic!=asked` failure on every run until this asked whether anything had been read. What can be read
/// is read from **the source whose registration stands** — the frame on screen, which for the whole of a
/// relaunch's wait is the player before the replacement, and after the swap is the one that just drew
/// (see `the_standing_source_process`) — so a line written during a wait cannot describe a window that
/// has drawn nothing.
///
/// The line is written once per seek — in place for mpv, and at the swap for FFmpeg's player, whose
/// seek is a new window — and never per frame.
pub(super) fn note_the_seek_picture(pid: u32, seconds: f64) {
    let asked = THUMBNAIL
        .lock()
        .ok()
        .and_then(|held| held.source)
        .map(|shape| shape.client);
    let standing = the_standing_source();
    let source = standing
        .map(|(window, _)| window)
        .or_else(the_source_window);
    let cli = source.and_then(|source| unsafe { client_size(source) });
    // The picture belongs to the player whose window is on screen; the caller's pid is the answer only
    // where nothing is registered yet.
    let pid = standing.map(|(_, process)| process).unwrap_or(pid);
    let pic = the_players_picture_size(pid);
    let differs = match asked {
        Some(asked) if cli != Some(asked) => " cli!=asked",
        Some(asked) if pic.is_some_and(|pic| pic != asked) => " pic!=asked",
        _ => "",
    };
    note_pin_trace!(
        "seek picture cli {} pic {} asked {} source {:?}{differs} second {seconds:.3}",
        a_box_as_text(cli),
        a_box_as_text(pic),
        a_box_as_text(asked),
        source,
    );
}

/// The window the registration standing draws, and the player it belongs to: **the picture on screen**,
/// where one stands (see `the_standing_source_process`).
///
/// It takes no host, because the line that reads it is about the film rather than about a box: a seek is
/// the same seek in a hover's host and in a pin's.
fn the_standing_source() -> Option<(HWND, u32)> {
    let held = THUMBNAIL.lock().ok()?;
    if held.registration == 0 || held.registered == 0 || held.registered_process == 0 {
        return None;
    }
    Some((HWND(held.registered as *mut _), held.registered_process))
}

/// A box as a trace line spells it, and `none` where there is no answer at all.
fn a_box_as_text(size: Option<(i32, i32)>) -> String {
    size.map_or_else(
        || "none".to_string(),
        |(width, height)| format!("{width}x{height}"),
    )
}

/// Whether the road owes the player an ask: the first one as soon as the film is waiting, then one
/// per `LATE_NATIVE_EVERY`, and the last one at the end of the wait (which is the one that places the
/// film at the largest box where the player never answered).
fn a_late_placement_is_due(waited: Duration, since_the_last_ask: Option<Duration>) -> bool {
    if waited >= LATE_NATIVE_WAIT {
        return true;
    }
    since_the_last_ask.is_none_or(|since| since >= LATE_NATIVE_EVERY)
}

/// The one box a source is ever asked for — the film's own picture size at the corner of the host
/// showing it, clamped to what the window system will take — and nothing at all for a source that has
/// been placed already, for a host that is not there, or for a film whose own size is not known yet.
///
/// **Every `None` is a promise.** A placed source is a source whose client, whose picture and whose
/// shape cannot change under the Shell again, so a press, a drag, a release and a settle have nothing
/// to ask for (see `the_band_was_placed`); and a film that has not been sized is waited for rather
/// than guessed at, because placing a guess and correcting it later is exactly the flash this road
/// exists to not have (see `the_films_native_size`).
fn the_one_placement(
    placed: Option<SourceShape>,
    host_origin: Option<(i32, i32)>,
    native: Option<(i32, i32)>,
    max_track: (i32, i32),
) -> Option<SourceBox> {
    if placed.is_some() {
        return None;
    }
    let (x, y) = host_origin?;
    let wanted = native?;
    Some((
        x,
        y,
        wanted.0.clamp(1, max_track.0.max(1)),
        wanted.1.clamp(1, max_track.1.max(1)),
    ))
}

/// The largest window the window system will take, per axis: `SM_CXMAXTRACK` and `SM_CYMAXTRACK`
/// measured on the desktop, which is the clamp a source asked for a film larger than the screen is
/// held to (see `the_one_placement`).
fn the_max_track() -> (i32, i32) {
    // Safety: two metrics questions about the desktop, which need no window.
    unsafe {
        (
            GetSystemMetrics(SM_CXMAXTRACK).max(1),
            GetSystemMetrics(SM_CYMAXTRACK).max(1),
        )
    }
}

/// The film's own picture size as the player answers it, or the probe's measurement of the same film
/// where the player has not answered: the size the film would be drawn at were it not fitted to a
/// window, which is the picture's own size and not the box it is shown in (see `the_players_picture_size`
/// and `the_films_own_size`).
///
/// **A player that has just made its window has not opened the file yet**, so adoption waits for the
/// answer rather than placing the box it could guess: a handful of bounded reads a few milliseconds
/// apart, and then the size the same film was probed at before it was launched. The wait is on the
/// adoption pass and nowhere near a drag (see `adopt_the_source`), and a film neither the player nor
/// the probe can answer for is left to the road that watches the player (see
/// `place_the_source_when_the_player_answers`).
fn the_films_native_size(pid: u32) -> Option<(i32, i32)> {
    for ask in 0..NATIVE_SIZE_ASKS {
        if let Some(size) = the_players_picture_size(pid) {
            return Some(size);
        }
        // **A player with no pipe to answer through never answers, however long it is waited for**,
        // and the probe's measurement of the same film is the answer instead (see
        // `mpv_ipc::answers_properties`).
        if ask + 1 < NATIVE_SIZE_ASKS && mpv_ipc::answers_properties(pid) {
            std::thread::sleep(NATIVE_SIZE_WAIT);
        }
    }
    the_films_own_size(pid)
}

/// The film's own picture size as the probe measured it when the film was loaded: read out of the
/// geometry cache and never asked of the player, which is the fallback for a player that has not
/// opened the file by the time its window is adopted (see `the_films_native_size`).
///
/// **The probe only answers for the film the player was begun for**, which the player's own record
/// names: the pin is still showing the file it is leaving for the whole of a swap's wait, so a probe of
/// *that* file would place the new film's source at the old film's size and shape, which is the
/// previous film's shape cropping the next one's first frames. A player this app cannot name a file
/// for is answered nothing here, and the player's own `video-params` is what places it (see
/// `the_film_the_player_plays`).
fn the_films_own_size(pid: u32) -> Option<(i32, i32)> {
    let path = the_film_the_player_plays(pid)?;
    match cached_video_geometry(&path)? {
        ProbedGeometry::Measured(geometry) => Some((geometry.width as i32, geometry.height as i32)),
        // A file the probe could not measure is not a size: the player's answer is what is waited
        // for, and nothing is placed on the shape of a file nothing measured.
        ProbedGeometry::Unmeasurable => None,
    }
}

/// The file the player `pid` was begun for, where this app has it on record: the hover's own file
/// beside the hover's player, the pin's beside the player begun into the pin's host, and the pin's own
/// media only while the player the pin holds is that very one.
///
/// **The record is what keeps a step's outgoing film from answering for the incoming one**, and the
/// player's own record is what answers for a film the pin has not installed yet: a swap's player is
/// begun before the file it plays is the pin's, and a path read off the pin's media without the player
/// beside it is the old film's (see `the_films_own_size` and `pin_film_host::the_hosted_players_file`).
fn the_film_the_player_plays(pid: u32) -> Option<PathBuf> {
    if pid == 0 {
        return None;
    }
    if hover_film_host::hover_film_host_player() == Some(pid) {
        return hover_film_host::the_hover_players_file(pid);
    }
    if let Some(path) = pin_film_host::the_hosted_players_file(pid) {
        return Some(path);
    }
    let (path, _) = pinned_media_owner()?;
    let held = CURRENT_MEDIA.lock().ok().and_then(|media| {
        media
            .as_ref()
            .and_then(|media| media.video_process.as_ref().map(|process| process.id()))
    });
    (held == Some(pid)).then_some(path)
}

/// The film's own picture size out of the player's `video-params`, in one read: `dw`/`dh` are the
/// size the film is displayed at once its pixel aspect is counted, and `w`/`h` the coded size, which
/// is the same number for the square pixels every film this road is handed has. **A film the player
/// turns on its side is as tall as its codec says it is wide**, so a `rotate` of 90 or 270 swaps them:
/// the size asked for is the shape the player draws, not the shape the frame is stored in.
///
/// **`osd-dimensions` is not this.** Measured on mpv 0.41.0 it answers
/// `{"w":640,"h":300,"par":1.000000,"aspect":2.133333,"mt":0,"mb":0,"ml":53,"mr":54}` — the window the
/// output is configured for and the letterbox the film leaves inside it, which is what the *box* made
/// the picture rather than what the picture is (see the module note).
fn the_players_picture_size(pid: u32) -> Option<(i32, i32)> {
    if pid == 0 {
        return None;
    }
    let params = mpv_ipc::get_property(pid, "video-params")?;
    let (width, height) = (
        a_node_field(&params, "dw").or_else(|| a_node_field(&params, "w"))?,
        a_node_field(&params, "dh").or_else(|| a_node_field(&params, "h"))?,
    );
    if width <= 0 || height <= 0 {
        return None;
    }
    let turned = a_node_field(&params, "rotate").unwrap_or(0).rem_euclid(180) == 90;
    Some(if turned {
        (height, width)
    } else {
        (width, height)
    })
}

/// One field out of a node property's answer: mpv writes an object as `{"w":640,"h":300,…}`, so the
/// field is the number after its own key and up to the field beside it (see `mpv_ipc::reply_data`).
fn a_node_field(object: &str, name: &str) -> Option<i32> {
    let key = format!("\"{name}\":");
    let after = object.find(&key)? + key.len();
    a_pixel_count(object[after..].split([',', '}']).next()?)
}

/// A property's answer as a pixel count, or nothing where it is not a number this app can use: mpv
/// answers a scalar with a bare number, and may write a whole one as a float.
fn a_pixel_count(answer: &str) -> Option<i32> {
    answer
        .trim()
        .parse::<f64>()
        .ok()
        .map(|value| value.round() as i32)
}

/// Take the picture off the screen with the host it is seen in, or put it back.
///
/// **A hidden source is a thumbnail the Shell stops drawing**, so this is only ever asked where the
/// host goes down with it — a pin put away, a park over the band — and never while the host is up.
/// The call is posted rather than made, for the reason every call on this window is (see
/// `adopt_the_source`).
///
/// **The ask is remembered, because the window's own answer is a message that may still be on its
/// way.** `ShowWindowAsync` to a player's window is a post to another process, so a hide asked for at
/// a collapse can still be in flight when the restore asks for the show — and a show that read
/// `IsWindowVisible` would see the window still up, decline to ask for anything, and let the queued
/// hide land last: the source down, the host up, and the pin a black band over a film that is playing.
/// What is compared against here is therefore the last ask (see `THE_SOURCE_SHOW`), not the window,
/// and a host's show on every tick of a pin that is drawing still costs one comparison and no message.
pub(super) fn set_the_source_visible(visible: bool) {
    let Some(source) = the_source_window() else {
        return;
    };
    let number = source.0 as isize;
    let asked = THE_SOURCE_SHOW.lock().ok().and_then(|held| *held);
    if asked == Some((number, visible)) {
        return;
    }
    if let Ok(mut held) = THE_SOURCE_SHOW.lock() {
        *held = Some((number, visible));
    }
    // Safety: the handle is a window this module adopted, and the form asked for is the one the pin's
    // own state wants.
    let _ = unsafe { ShowWindowAsync(source, if visible { SW_SHOWNOACTIVATE } else { SW_HIDE }) };
}

/// The source this road last asked a show or a hide for, and which of the two: the record that keeps a
/// hide still in flight from being read as "already hidden" by the restore that follows it (see
/// `set_the_source_visible`).
///
/// It is written where the source is recorded — a player's window is up when it is adopted, and it is
/// that window the next ask is about — and cleared with the source it names.
static THE_SOURCE_SHOW: Mutex<Option<(isize, bool)>> = Mutex::new(None);

/// Take the registration down: the host it was drawn in is going, and a registration left behind a
/// window that is gone is a leak of the Shell's own resources.
///
/// `reason` is the road that took it, written out because **the band's pixels are a function of
/// whether this registration stands**: a road that lets it go and does not paint what it left is
/// the black rectangle that stays where the film was (see `pin_film_host::the_band_owes_a_repaint`).
pub(super) fn forget_the_thumbnail_because(reason: &str) {
    note_pin_trace!("forget thumbnail (reason={reason})");
    let_the_picture_go();
    STILL_HOST.store(0, Ordering::SeqCst);
    STILL_PID.store(0, Ordering::SeqCst);
    // The measurement of what the pin was left showing goes with the pin (see `still_probe`).
    super::still_probe::forget("the pin is going");
    if let Ok(mut held) = HELD_FRAMES.lock() {
        // The pin that counted these is gone, so the next one opens its own first second.
        *held = HeldFrames::default();
    }
    // **A registration standing is a picture, and letting it go is a band that has changed**: the
    // fill that follows this is the pin's own claim, and it must be painted rather than left over
    // whatever the last paint wrote (see `pin_film_host::the_band_owes_a_repaint`).
    pin_film_host::the_band_owes_a_repaint();
}

/// The same, named for the pin's own take-down — **and the name the suite knows**, which is why it is
/// a test-only door: every road of the app's has to say which one it is (see
/// `forget_the_thumbnail_because`).
#[cfg(test)]
pub(super) fn forget_the_thumbnail() {
    forget_the_thumbnail_because("the caller let the picture go");
}

/// **Let one picture off the screen where the pin is not going**: the registration and the record of
/// the source it drew are given up, and the pin's own state — what it counted, what it measured, the
/// still a press decoded — is left alone.
///
/// It is the step a still's paint makes room with: the outgoing player's registration goes so this
/// app's own cover can stand where it was, and nothing of the pin goes with it (see `stand_the_still`).
fn let_the_picture_go() {
    if let Ok(mut held) = THUMBNAIL.lock() {
        let registration = held.registration;
        *held = Thumbnail::NOTHING;
        let_the_registration_go(registration);
    }
    // The film this road was waiting to size went with the window it was going to place, and so did
    // the source it was waiting to hand the picture to (see `arm_the_pending_source`).
    forget_the_late_placement();
    forget_the_pending_source();
    REGISTERED_HOST.store(0, Ordering::SeqCst);
    SOURCE.store(0, Ordering::SeqCst);
    SOURCE_PROCESS.store(0, Ordering::SeqCst);
    if let Ok(mut held) = THE_SOURCE_SHOW.lock() {
        // The window this was an ask about is nobody's source now (see `set_the_source_visible`).
        *held = None;
    }
}

/// **Let the outgoing player's registration go so a cover can stand where it was, without taking a
/// replacement's place in the queue with it.**
///
/// A replacement may already be adopted and placed, waiting for this very swap: it is not the picture
/// being covered, and its one placement asked for a second time is the flash this road exists not to
/// have (see `place_the_source`). So a pending source keeps its record and its placement, and only the
/// registration and the host's own bookkeeping go. With nothing adopted yet the source record is
/// cleared too, and the pass that watches the replacement adopts it afresh — through the cover, which
/// is what `a_still_stands` holds the registration behind (see `adopt_the_source`).
fn let_the_registration_go_for_a_still() {
    if let Ok(mut held) = THUMBNAIL.lock() {
        let registration = held.registration;
        *held = Thumbnail::NOTHING;
        let_the_registration_go(registration);
    }
    forget_the_late_placement();
    REGISTERED_HOST.store(0, Ordering::SeqCst);
    if !PENDING_SOURCE.lock().is_ok_and(|pending| pending.is_some()) {
        forget_the_pending_source();
        SOURCE.store(0, Ordering::SeqCst);
        SOURCE_PROCESS.store(0, Ordering::SeqCst);
    }
}

/// Take the registration down where it is drawn in `host`: a host of the hover's that goes while it
/// is showing a film of its own, which is every road that is not this one.
pub(super) fn forget_the_thumbnail_of(host: HWND) {
    let drawn_in = REGISTERED_HOST.load(Ordering::SeqCst) == host.0 as isize;
    if drawn_in {
        forget_the_thumbnail_because("the host it was drawn in has gone");
    }
}

/// The host a relaunch has painted a still into, or nothing while none stands: **a picture of this
/// app's own where the registration was drawing**, and the state that keeps a new source from taking
/// the box before it can draw (see `the_picture_is_standing`). It goes with the registration, the
/// host, and the pin (see `forget_the_thumbnail`).
static STILL_HOST: AtomicIsize = AtomicIsize::new(0);

/// The player the still standing was taken of, or nothing while none stands: **its window is not
/// adopted again**. The player is ended a moment after the still is painted, and the pass that watches
/// a player the pin is drawing would otherwise find the window it has just been let go of and register
/// a dying source over the still (see `adopt_the_source`).
static STILL_PID: AtomicU32 = AtomicU32::new(0);

/// Whether a still this app painted for a relaunch is the picture on screen (see `STILL_HOST`).
pub(super) fn a_still_stands() -> bool {
    STILL_HOST.load(Ordering::SeqCst) != 0
}

/// **The still a relaunch takes out of the source window's own device context** — the cover
/// `RHP_FREEZE=dc` is.
///
/// **Why the source's own DC, and neither the screen nor a window's last presented frame.** A pinned
/// film's source wears `WS_EX_LAYERED` at alpha 1/255 so that nothing of it reaches the screen, which
/// makes a `BitBlt` off the screen black, and a window that draws through D3D answers an ask for its own
/// frame with black or with a remake of it rather than the frame (see `surfaces`). What is left is the
/// surface the player itself draws into — the one the readiness sample already reads lit twenty times a
/// second **throughout the very gap this cover is for** (see `the_source_has_lit_pixels`), which is what
/// makes a copy of it worth a relaunch.
///
/// **It costs about a millisecond, so it is taken at the release**: nothing is decoded, nothing waits,
/// and the frame it paints is the frame that was on screen a moment ago.
///
/// **A frame that cannot be a picture is never painted** (see `capture_the_sources_dc`): `false` is
/// `off` for that one relaunch — nothing painted, nothing unregistered, nothing ended — so the outgoing
/// player goes on drawing until its replacement does, and the pin is the picture it was.
///
/// The order from here on is `stand_the_still`'s, which is the whole correctness of the cover:
/// the frame into the host's retained cover, one repaint without an erase, `DwmFlush`, the registration
/// let go, and only then the player ended.
pub(super) fn stand_the_outgoing_picture_from_its_dc(pid: u32, seconds: f64) -> bool {
    let Some(host) = pin_film_host::film_host() else {
        note_pin_trace!(
            "placeholder unavailable pid {pid}: the pin has no host to paint a still in"
        );
        return false;
    };
    // The probe that reads the pin throughout the wait begins where the cover does (see `still_probe`).
    super::still_probe::start(pid);
    let Some(((width, height), pixels)) = capture_the_sources_dc(pid) else {
        note_pin_trace!("placeholder unavailable pid {pid}: the old picture plays on");
        return false;
    };
    stand_the_still(host.0 as isize, pid, (width, height), pixels, seconds)
}

/// The file the captured frame is dumped to, beside the trace log in the temp folder, so a run can be
/// held against a frame decoded at the same second (see `dump_the_frame`).
const THE_DC_STILL_PNG: &str = "rhp-pin-dc-still.png";

/// **How much of the sampled grid must be lit for a captured frame to be a picture at all**: 2% of a
/// 16×16 grid, which is six of its points.
///
/// It rejects the all-zero surface an empty, dead or never-drawn window answers with and accepts a
/// genuinely dark scene, because a dark frame's pixels are small numbers rather than zeros. It is a
/// fraction rather than one lit point for the reason the readiness sample's single point is not enough
/// here: a stale or garbage surface can hold one bright pixel, and a cover painted from one is a black
/// band with a speck in it.
const THE_DC_STILL_MIN_LIT_SHARE: u32 = 2;

/// The frame in the source's own window as the BGRA pixels and the size a still is painted from, or
/// nothing where the window holds no picture (see `stand_the_outgoing_picture_from_its_dc`).
fn capture_the_sources_dc(pid: u32) -> Option<((u32, u32), Vec<u8>)> {
    let begun = Instant::now();
    if !the_source_is_up(pid) {
        return None;
    }
    let source = the_source_window()?;
    // Safety: the handle is the source this module adopted from a player's own window.
    let (width, height) = unsafe { client_size(source) }.filter(|(w, h)| *w > 0 && *h > 0)?;
    let surface = DibSurface::create(width as u32, height as u32)?;
    // Safety: the two contexts are the source's own client and this call's surface, and what is asked
    // of them is one `SRCCOPY` between them; the source's is given back before this answers.
    let copied = unsafe {
        let dc = GetDC(source);
        if dc.is_invalid() {
            return None;
        }
        let copied = BitBlt(surface.dc, 0, 0, width, height, dc, 0, 0, SRCCOPY).is_ok();
        ReleaseDC(source, dc);
        // GDI batches its calls and what is read next is not a GDI call.
        let _ = GdiFlush();
        copied
    };
    if !copied {
        note_pin_trace!("still captured from the source dc pid {pid}: the copy failed");
        return None;
    }
    // `DibSurface::pixels` is BGRA at the size of the surface, which is the shape the cover is put in
    // as, so the crop and the paint below are the ones a still of this app's always has.
    let pixels = surface.pixels();
    let lit = the_frames_lit_share(&pixels, width as u32, height as u32);
    note_pin_trace!(
        "still captured from the source dc pid {pid} {width}x{height} lit={lit}% in {}ms",
        begun.elapsed().as_millis(),
    );
    if lit < THE_DC_STILL_MIN_LIT_SHARE {
        note_pin_trace!(
            "still captured from the source dc pid {pid} rejected: lit={lit}% is under {THE_DC_STILL_MIN_LIT_SHARE}%, so it is not a picture"
        );
        return None;
    }
    // **A run holds the frame to something other than a judge's eye**: the checksum and the size are
    // what a frame decoded at the same second is compared against, and the PNG beside the trace log is
    // the picture itself (see `dump_the_frame`).
    if super::pin_trace::asked() {
        let dumped = dump_the_frame(&pixels, width as u32, height as u32);
        note_pin_trace!(
            "still captured frame pid {pid} {width}x{height} checksum={:016x} png={}",
            the_frames_checksum(&pixels),
            if dumped { THE_DC_STILL_PNG } else { "-" },
        );
    }
    Some(((width as u32, height as u32), pixels))
}

/// The percentage of a 16×16 grid over the inner half of a BGRA frame that is lit: **the pin's own
/// probe's grid**, applied to the captured frame rather than to the desk (see `still_probe`).
///
/// The grid sits in the middle of the frame for the reason it does there: a player's own border and
/// letterbox are at the edges, and a lit border over a black frame would read as a picture that is
/// there.
fn the_frames_lit_share(pixels: &[u8], width: u32, height: u32) -> u32 {
    let grid = super::still_probe::GRID;
    if width == 0 || height == 0 || pixels.len() < width as usize * height as usize * 4 {
        return 0;
    }
    let stride = width as usize * 4;
    let (left, right) = (width / 4, width * 3 / 4);
    let (top, bottom) = (height / 4, height * 3 / 4);
    let mut lit = 0;
    for row in 0..grid {
        let y = top + (bottom - top) * row / grid;
        for column in 0..grid {
            let x = left + (right - left) * column / grid;
            let at = y as usize * stride + x as usize * 4;
            if pixels[at] | pixels[at + 1] | pixels[at + 2] != 0 {
                lit += 1;
            }
        }
    }
    lit * 100 / (grid * grid)
}

/// FNV-1a over the frame's bytes: the one number a run can hold against a frame decoded at the same
/// second (see `dump_the_frame`).
fn the_frames_checksum(pixels: &[u8]) -> u64 {
    pixels.iter().fold(0xcbf2_9ce4_8422_2325, |hash, byte| {
        (hash ^ u64::from(*byte)).wrapping_mul(0x0000_0100_0000_01b3)
    })
}

/// **Write the captured frame beside the trace log**, so the next run compares a picture rather than a
/// percentage: BGRA from GDI carries no alpha, so it is written opaque.
///
/// It is a dump for a measurement and not a feature of the road: one file, overwritten by each capture,
/// and nothing reads it back.
fn dump_the_frame(pixels: &[u8], width: u32, height: u32) -> bool {
    let rgba: Vec<u8> = pixels
        .as_chunks::<4>()
        .0
        .iter()
        .flat_map(|bgra| [bgra[2], bgra[1], bgra[0], 0xFF])
        .collect();
    image::save_buffer(
        std::env::temp_dir().join(THE_DC_STILL_PNG),
        &rgba,
        width,
        height,
        image::ExtendedColorType::Rgba8,
    )
    .is_ok()
}

/// **Paint a captured frame as the pin's cover, then let the registration go and end the player it was
/// taken from — in that order, which is the correctness of the whole road.**
///
/// A DWM thumbnail composes *above* the host's own content, so a paint under a standing registration
/// shows nothing at all; and an unregister before the paint has committed shows the host's black
/// erase. So the steps are: the frame into the host's retained cover, one repaint (no erase), a
/// `DwmFlush` so the paint and the composition are committed, the registration let go, and only then
/// the outgoing player ended — freeing its file, sound device and GPU device before the replacement
/// needs them (see the module note in `pin_film_host` and `ARCHITECTURE.md`).
///
/// One trace line per step, so a wrong order is a line in the run rather than a black frame nobody
/// can explain.
fn stand_the_still(
    host: isize,
    pid: u32,
    (width, height): (u32, u32),
    pixels: Vec<u8>,
    at: f64,
) -> bool {
    let expected = width as usize * height as usize * 4;
    // **A frame that is not the size it says it is is not painted**: a still cut from a short frame is
    // a smear rather than a picture.
    if width == 0 || height == 0 || pixels.len() < expected {
        note_pin_trace!("placeholder unavailable pid {pid}: the frame came back short");
        return false;
    }
    let host = HWND(host as *mut _);
    // Safety: the handle is the host this still was captured for, a window of this app's.
    if (unsafe { client_size(host) }).is_none_or(|(width, height)| width <= 0 || height <= 0) {
        note_pin_trace!("placeholder unavailable pid {pid}: the pin has no box to fill");
        return false;
    }
    if !pin_film_host::put_the_cover_in(host, pixels, width, height) {
        note_pin_trace!(
            "placeholder unavailable pid {pid}: the host is not the one it was decoded for"
        );
        return false;
    }
    // **One repaint and no erase**: the host's own erase is the black a still must never show, and a
    // repaint without it is what draws the cover (see `pin_film_host::answer_the_paint_of_a_cover`).
    unsafe {
        let _ = InvalidateRect(host, None, false);
        let _ = UpdateWindow(host);
    }
    note_pin_trace!("placeholder painted from a still pid {pid} at {at:.3}");
    note_pin_trace!("still painted pid {pid} {width}x{height}");
    // The composition is asked to commit before the registration that covers this paint goes, so the
    // step from the Shell's picture to this app's own is one the user never sees the inside of.
    unsafe {
        let _ = DwmFlush();
    }
    note_pin_trace!("dwm flushed");
    // **Let go of before the host's record is written**, so the pass that watches the outgoing player
    // cannot adopt it back in between (see `STILL_PID`). This is the picture only: the pin is not
    // going, and what it counted and measured — and a replacement already waiting for the swap —
    // stays with it.
    let_the_registration_go_for_a_still();
    STILL_PID.store(pid, Ordering::SeqCst);
    STILL_HOST.store(host.0 as isize, Ordering::SeqCst);
    note_pin_trace!("thumbnail let go for a still pid {pid}");
    super::still_probe::note_painted();
    // **The player is ended at once**, now that nothing of the screen is drawn from it: the file, the
    // sound device and the GPU device are free before the replacement needs them. A player this road
    // had already parked for the loop goes with it rather than being ended twice.
    retire_orphaned_player(pid);
    forget_the_retirement_of(pid);
    if pid != 0 {
        note_pin_trace!("outgoing player ended at once pid {pid}");
    }
    true
}

/// The rectangle a picture of `aspect` is drawn in inside a `box_`: fitted to the box's own shape and
/// centred in it, as `x, y, width, height` — and the whole box where the shape is not one this app can
/// use.
///
/// **Integer, inside the box and never negative**, which is what makes it a rectangle the Shell will
/// take: one axis is the box's and the other is the box's times or over the picture's shape, rounded so
/// a film is a pixel out at most and clamped so it is never a pixel over.
///
/// **Adoption asks this of the source's client and the film's own shape**, once, and what it gives is
/// the picture's rectangle for the pin's whole life: a source placed at the size of its film has no
/// margins to leave, and one clamped by the window system letterboxes exactly like the player's own
/// fit (see `place_the_source` and `the_thumbnail_box`).
fn a_picture_fitted_to(aspect: Option<f64>, box_: (i32, i32)) -> (i32, i32, i32, i32) {
    let (box_width, box_height) = (box_.0.max(1), box_.1.max(1));
    let (width, height) = match aspect {
        Some(aspect) if aspect.is_finite() && aspect > 0.0 => {
            if (box_width as f64 / box_height as f64) > aspect {
                let width = (box_height as f64 * aspect)
                    .round()
                    .clamp(1.0, box_width as f64);
                (width as i32, box_height)
            } else {
                let height = (box_width as f64 / aspect)
                    .round()
                    .clamp(1.0, box_height as f64);
                (box_width, height as i32)
            }
        }
        _ => (box_width, box_height),
    };
    (
        (box_width - width) / 2,
        (box_height - height) / 2,
        width,
        height,
    )
}

/// What one registration of the Shell's thumbnail shows: the box the picture is drawn in, the rectangle
/// of the source's client it is taken from, and the client that rectangle was fitted into.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
struct ThumbnailBox {
    /// The rectangle of the host's client area the picture is drawn in: **the whole of it, always** (see
    /// `the_thumbnail_box`).
    destination: (i32, i32, i32, i32),
    /// The picture's rectangle in the source's own client area, as left, top, right and bottom — or
    /// nothing where the whole client is what is shown (see `the_thumbnail_box`).
    source: Option<(i32, i32, i32, i32)>,
    /// The client the source is at, as a width and a height.
    client: (i32, i32),
    /// Whether the crop had to be cut to the picture rather than being the extent the box's shape asks
    /// for — a box the picture does not fill, whose crop is the picture whole and the Shell's to
    /// upscale — which a trace reads and nothing else does (see `the_crop_filling`).
    clamped: bool,
}

/// **What a thumbnail of the source shows: the box, filled.** The destination is the host's whole client
/// area and the source is the rectangle of the picture that *is* the box's shape — so the film covers its
/// box and there is nothing left for the host's black to show through (see `the_crop_filling`).
///
/// `aspect` is the film's own width over its height and `client` the source's client, and **both are
/// constants of the pin's life**: they are what the one placement left behind, never a box the window
/// has just been asked for (see `SourceShape` and `place_the_source`). Nothing here reads a window or a
/// player, so a frame is a function of the box alone.
///
/// **A destination fitted into the box was a bar of black waiting for the two numbers to disagree**, and
/// they disagreed at exactly the two moments the old road re-read its picture: a drag's beginning and
/// its release, where a source re-asked for another box cropped a surface the Shell was not sampling
/// (see the module note). **Filling the box is bars made impossible by construction**: whatever the two
/// shapes are, the destination is the client and the crop is the part of the picture that has the
/// client's shape, so there is no rectangle left over to read as the wrong thing.
///
/// At rest the two shapes agree — a pin is built at its film's shape and a drag keeps the shape it began
/// with (`pin_geometry::resize_pinned_content`) — so the crop is the picture whole and the pair handed to
/// the Shell is the one the containing road handed it.
///
/// **A picture nothing has read yet is the whole client into the whole box** — the fallback for a
/// shape this app cannot crop by.
fn the_thumbnail_box(box_: (i32, i32), client: (i32, i32), aspect: Option<f64>) -> ThumbnailBox {
    let client = (client.0.max(1), client.1.max(1));
    let box_ = (box_.0.max(1), box_.1.max(1));
    let (source, clamped) = match aspect {
        Some(_) => {
            let ((left, top, width, height), clamped) = the_pictures_crop(client, box_, aspect);
            (Some((left, top, left + width, top + height)), clamped)
        }
        None => (None, false),
    };
    ThumbnailBox {
        destination: (0, 0, box_.0, box_.1),
        source,
        client,
        clamped,
    }
    .checked(box_)
}

/// **A pair is a promise the Shell keeps or cannot keep**: it draws the destination with the crop, so
/// this is where the two are held to their invariants rather than hoped into them — the destination is
/// the box, and the crop is a non-empty rectangle inside the window it was fitted into. Nothing is
/// clamped here: a crop that failed this test is a rect the Shell is asked to sample from outside the
/// surface it is sampling, which is the band of desktop the module note describes, so it is a program
/// error and not a frame to paint.
impl ThumbnailBox {
    fn checked(self, box_: (i32, i32)) -> Self {
        debug_assert_eq!(
            self.destination,
            (0, 0, box_.0.max(1), box_.1.max(1)),
            "the destination is the whole box the pair was built for"
        );
        debug_assert!(
            self.source.is_none_or(|(left, top, right, bottom)| {
                left >= 0
                    && top >= 0
                    && right > left
                    && bottom > top
                    && right <= self.client.0.max(1)
                    && bottom <= self.client.1.max(1)
            }),
            "the crop {:?} is a non-empty rect inside the client {:?}",
            self.source,
            self.client
        );
        self
    }
}

/// **The crop a host's client is filled by a picture in a source's own client**: the picture fitted
/// inside that client and then cut to the client's shape — the one rectangle the registration samples
/// and the one a captured still is painted with, so a relaunch's cover is the frame it replaced rather
/// than a second guess at its shape (see `the_thumbnail_box` and
/// `stand_the_outgoing_picture_from_its_dc`).
pub(super) fn the_pictures_crop(
    client: (i32, i32),
    box_: (i32, i32),
    aspect: Option<f64>,
) -> ((i32, i32, i32, i32), bool) {
    the_crop_filling(a_picture_fitted_to(aspect, client), box_)
}

/// The rectangle of `picture` — `left, top, width, height` inside the source's client — that has the
/// shape of `box_`: the film cropped to fill the box and centred in it, and whether the crop had to be
/// cut to the picture rather than being the extent the shape asks for.
///
/// **A picture wider than the box keeps its height and gives up its sides; one no wider keeps its width
/// and gives up its top and bottom.** Integer and clamped, so the rectangle is always inside the picture
/// and never empty, whatever the two shapes are and however the rounding falls.
///
/// **A box larger than the picture on either axis is the picture whole**, which is the one case where
/// cropping would be throwing pixels away: the Shell is scaling the crop *up* into the box either way,
/// so a crop of a picture that is already being magnified is a softer picture for nothing. That is the
/// upscale a big pin is drawn with — from the film's own size and not from a box three times a small
/// pin — and the trace's `clamp` field is where it is read (see the module note).
fn the_crop_filling(
    picture: (i32, i32, i32, i32),
    box_: (i32, i32),
) -> ((i32, i32, i32, i32), bool) {
    let (left, top, width, height) = picture;
    let (width, height) = (width.max(1), height.max(1));
    let (box_width, box_height) = (box_.0.max(1), box_.1.max(1));
    if box_width > width || box_height > height {
        return ((left, top, width, height), true);
    }
    let (crop_width, crop_height, clamped) =
        if width as f64 / height as f64 > box_width as f64 / box_height as f64 {
            let asked = (height as f64 * box_width as f64 / box_height as f64).round() as i32;
            let crop = asked.clamp(1, width);
            (crop, height, crop != asked)
        } else {
            let asked = (width as f64 * box_height as f64 / box_width as f64).round() as i32;
            let crop = asked.clamp(1, height);
            (width, crop, crop != asked)
        };
    (
        (
            left + (width - crop_width) / 2,
            top + (height - crop_height) / 2,
            crop_width,
            crop_height,
        ),
        clamped,
    )
}

/// The client box the source is at, read off the window for the trace's `cli` — the source is launched
/// `--no-border`, so its window box and its client area are one rectangle, and the host's own box stands
/// in for a source that has gone.
///
/// **This is a question Windows answers about a window, not a read of the player**, and nothing decides
/// anything by it: the crop is built from the client the one placement asked for (see `place_the_source`),
/// so this is the honest live answer printed beside `pic` and never a box this app has re-read its way to.
fn the_source_client(source: HWND, host: (i32, i32)) -> (i32, i32) {
    match unsafe { client_size(source) } {
        Some(client) if client.0 > 0 && client.1 > 0 => client,
        _ => host,
    }
}

/// **Give up the shape of the film before this one.** A [`SourceShape`] is the placement of one
/// source window in one player, and it is invalidated with that film: teardown forgets the whole
/// record, adoption of another window resets it (see `note_the_source`), and this is the check every
/// frame makes — a record left by a film that has gone is dropped rather than cropping the first
/// frames of the film that followed it by the old film's shape.
///
/// The line is written once, where the record goes, so a run says the shape was given up rather than
/// leaving the missing crop to be guessed at.
fn forget_the_shape_of_another_film(held: &mut Thumbnail) {
    // **The window the shape has to belong to is the one the registration draws**, which is the
    // current source until a handover replaces it: a shape whose source is waiting to take the picture
    // over is the swap's and is not this record's to give up (see `hand_the_picture_over`).
    let drawn = if held.registration != 0 {
        held.registered
    } else {
        SOURCE.load(Ordering::SeqCst)
    };
    let Some(shape) = held.source.filter(|shape| shape.source != drawn) else {
        return;
    };
    held.source = None;
    note_pin_trace!(
        "thumbnail source shape forgotten {:?}: it was placed for another film",
        HWND(shape.source as *mut _),
    );
}

/// Let one registration go, and count it: the Shell's own resource, and the number the road's tests
/// hold a handover to (see `the_unregistrations`).
fn let_the_registration_go(registration: isize) {
    if registration == 0 {
        return;
    }
    let _ = unsafe { DwmUnregisterThumbnail(registration) };
    #[cfg(test)]
    UNREGISTRATIONS.fetch_add(1, Ordering::SeqCst);
}

/// Whether a registration wearing `held.shown` has to be pointed at `wanted`: **nothing else is a
/// reason not to re-point.** Another host client is another destination and another box shape is
/// another crop, and a registration left wearing the old pair is a picture cropped off the surface the
/// Shell is sampling (see `settle_the_registration`).
fn the_registration_wants(held: &Thumbnail, wanted: ThumbnailBox) -> bool {
    held.registration == 0 || held.shown != Some(wanted)
}

/// Register the picture in the host's client area, or point the registration the host already wears at
/// the rectangles it is drawn with **now**: the host's live client area, and the crop the picture
/// adoption fixed gives the box's shape (see `the_thumbnail_box`).
///
/// **The invariant is one line long**: the destination rectangle is the host's whole client, so every
/// pixel of the host is drawn by the Shell and a bar is a rectangle nothing could have left over, and
/// the host's own erase paints that client under it — which is what a client the Shell has not yet
/// composed, or one whose source is not drawing, shows rather than the desktop. What breaks it is a
/// pair built against a box the host is no longer in, or a source rectangle reaching past the surface
/// the Shell is sampling, so both are read here and the pair is traced whether it moved or not (see
/// `note_the_thumbnail_pair`).
///
/// **The source's side of the pair is a constant.** Its client and the film's shape were fixed by the
/// one placement (see `place_the_source`), so a frame can only change with the host's box: `cli` in the
/// trace is the window's own live answer and `pic` the constant the pair was cropped against, and the
/// two are the same number for the pin's whole life — a frame where they disagree says something other
/// than this app resized the source.
fn settle_the_registration(held: &mut Thumbnail, host: HWND) {
    if held.registration != 0 && held.host != host.0 as isize {
        let registration = held.registration;
        held.registration = 0;
        held.registered = 0;
        held.shown = None;
        let_the_registration_go(registration);
    }
    held.host = host.0 as isize;
    REGISTERED_HOST.store(held.host, Ordering::SeqCst);

    // **The window this registration draws, and not necessarily the current source.** A source adopted
    // to take the picture over is not the picture until it is ready, so what a standing registration
    // draws is read out of the record rather than out of the source this road has just adopted (see
    // `hand_the_picture_over`).
    let (number, process) = if held.registration != 0 && held.registered != 0 {
        (held.registered, held.registered_process)
    } else {
        (
            SOURCE.load(Ordering::SeqCst),
            SOURCE_PROCESS.load(Ordering::SeqCst),
        )
    };
    if number == 0 {
        return;
    }
    let source = HWND(number as *mut _);
    let Some(client) = (unsafe { client_size(host) }) else {
        return;
    };
    // **The one read of the source at this call.** `cli` in the trace is this number and `pic` the one
    // the pair was cropped against, so a frame where the two disagree says which window a flash would
    // have to come from (see `note_the_thumbnail_pair`).
    let live = the_source_client(source, client);
    // **The shape is the film's and is invalidated with it**: a record left by the film that ended is
    // given up here rather than used to crop the first frames of the one that followed it (see
    // `forget_the_shape_of_another_film`).
    forget_the_shape_of_another_film(held);
    let (source_client, aspect) = held
        .source
        .map_or((live, None), |source| (source.client, source.aspect));
    let shown = the_thumbnail_box(client, source_client, aspect);
    let a_re_point = the_registration_wants(held, shown);
    // **The crop the Shell samples is kept on the screen before anything is handed over**: a held
    // frame can need the source slid too, because a pin dragged to another monitor has the same pair
    // and a different work area (see `keep_the_crop_on_the_screen`).
    keep_the_crop_on_the_screen(host, source, shown, a_re_point);
    if !a_re_point {
        // **A held frame is counted for a pin, and only while one is up**: a hover's host is asked the
        // same question on every tick and is not what the heartbeat and the summary are about (see
        // `HeldFrames`).
        if pin_film_host::film_host() == Some(host) {
            note_a_held_frame(live != shown.client);
        }
        return;
    }

    if held.registration == 0 {
        #[cfg(test)]
        REGISTRATIONS.fetch_add(1, Ordering::SeqCst);
        let Ok(registration) = (unsafe { DwmRegisterThumbnail(host, source) }) else {
            // Nothing to draw in until the registration is asked for again: a source that has made
            // its window and not yet presented is a registration the Shell can refuse. A refusal is
            // traced rather than swallowed, because nothing on this road asks a second time.
            note_pin_trace!(
                "thumbnail register failed host {:?} source {:?}",
                host,
                source,
            );
            REGISTERED_HOST.store(0, Ordering::SeqCst);
            return;
        };
        held.registration = registration;
        held.registered = number;
        held.registered_process = process;
    }

    let (x, y, width, height) = shown.destination;
    let mut properties = DWM_THUMBNAIL_PROPERTIES {
        dwFlags: DWM_TNP_RECTDESTINATION | DWM_TNP_SOURCECLIENTAREAONLY | DWM_TNP_VISIBLE,
        // The host's whole client area, unconditionally: the film fills the box it is drawn in (see
        // `the_thumbnail_box`).
        rcDestination: RECT {
            left: x,
            top: y,
            right: x + width,
            bottom: y + height,
        },
        rcSource: RECT::default(),
        opacity: 255,
        fVisible: true.into(),
        // The picture is the player's client area and not the frame mpv puts around it.
        fSourceClientAreaOnly: true.into(),
    };
    // **The source is taken from its picture and not from its whole client**: the letterbox margins mpv
    // leaves around a film of another shape are not part of the film, and stretching them into the box
    // is the pillar of black that read as a pin zoomed out (see `the_thumbnail_box`).
    if let Some((left, top, right, bottom)) = shown.source {
        properties.dwFlags |= DWM_TNP_RECTSOURCE;
        properties.rcSource = RECT {
            left,
            top,
            right,
            bottom,
        };
    }
    if unsafe { DwmUpdateThumbnailProperties(held.registration, &properties) }.is_err() {
        note_pin_trace!(
            "thumbnail properties failed host {:?} box {:?}",
            host,
            client,
        );
        let registration = held.registration;
        // The picture adoption placed is kept: the next registration is cropped by the same one, and a
        // failure of the Shell's is not a reason to fall back to the whole client.
        let source = held.source;
        *held = Thumbnail::NOTHING;
        held.source = source;
        let_the_registration_go(registration);
        REGISTERED_HOST.store(0, Ordering::SeqCst);
        return;
    }
    held.shown = Some(shown);
    // **A host the pin is standing over with a registration in it is the band's picture**, whatever
    // the band was painted for the wait before it (see `pin_film_host::the_band_is_a_standing_host`).
    if pin_film_host::film_host() == Some(host) && held.registered != 0 {
        pin_film_host::note_the_band_has_a_picture();
    }
    // **The host is asked to erase its whole client behind the picture**, which is the other half of
    // the invariant above: the Shell composes the thumbnail over this app's content, so a client the
    // host never erased is a pixel nobody drew — and one of those is the desktop seen through the pin.
    let _ = unsafe { InvalidateRect(host, None, true) };
    note_the_thumbnail_pair(host, client, live, shown);
}

/// The one line that makes a hole visible at a glance: the host the Shell is drawing into (its own box
/// and where it stands), the host client the pair was built from, the source's live client (`cli`), the
/// client the crop was built from (`pic`) and the exact rectangles handed over. **`host` and `box`
/// disagree only where something stands on the host's client** — a border, a frame style — which is a
/// picture that cannot line up with its pin's box, and **`cli` and `pic` disagree only where the source
/// was resized by something that is not this app**, which is the one way a crop can be of a surface the
/// Shell is not sampling.
fn note_the_thumbnail_pair(host: HWND, client: (i32, i32), live: (i32, i32), shown: ThumbnailBox) {
    let (host_x, host_y, host_width, host_height) =
        unsafe { window_box(host) }.unwrap_or((0, 0, client.0, client.1));
    let (x, y, width, height) = shown.destination;
    let (left, top, right, bottom) = match shown.source {
        Some((left, top, right, bottom)) => (left, top, right, bottom),
        None => (0, 0, shown.client.0, shown.client.1),
    };
    note_pin_trace!(
        "thumbnail update host {}x{}+{}+{} box {}x{} dest {}x{}+{}+{} cli {}x{} pic {}x{} src {}x{}+{}+{}",
        host_width,
        host_height,
        host_x,
        host_y,
        client.0,
        client.1,
        width,
        height,
        x,
        y,
        live.0,
        live.1,
        shown.client.0,
        shown.client.1,
        right - left,
        bottom - top,
        left,
        top,
    );
}

/// **A held frame: counted, not printed.** A tick that found the pair unchanged says almost nothing —
/// the source's side is a constant, so the only thing a held line could carry is that the pin has not
/// moved — and a run of 2,375 of them against 446 updates is 84% of the trace reading the same thing.
/// So a held frame is counted, a **heartbeat** is written at most once a second while a pin lives (the
/// frames of that second and whether any of them disagreed with the pair), and one **summary** line is
/// written at each settle with every frame held since the last summary (`thumbnail held n=2375
/// disagrees=0`). The updates are untouched: they are the measurement (see `note_the_thumbnail_pair`).
#[derive(Default)]
struct HeldFrames {
    /// The instant the second being counted opened, and the frames counted in it. Nothing until the
    /// first held frame, which opens the first second and prints nothing itself.
    heartbeat: Option<Instant>,
    in_the_second: u32,
    /// The frames held since the last summary, and whether any of them found the source's live client
    /// disagreeing with the client the crop was built from (see `note_the_thumbnail_pair`).
    since_summary: u32,
    disagrees: bool,
}

impl HeldFrames {
    /// Count one held frame, answering the heartbeat line where a whole second has passed since the
    /// last one: at most one line a second, and nothing at all for the first second's opening frame.
    fn a_frame(&mut self, now: Instant, disagrees: bool) -> Option<String> {
        self.in_the_second += 1;
        self.since_summary += 1;
        self.disagrees |= disagrees;
        let opened = *self.heartbeat.get_or_insert(now);
        if now.duration_since(opened) < Duration::from_secs(1) {
            return None;
        }
        let line = format!(
            "thumbnail held heartbeat n={} disagrees={}",
            self.in_the_second, self.disagrees as u8,
        );
        self.heartbeat = Some(now);
        self.in_the_second = 0;
        Some(line)
    }

    /// The summary line a settle is owed: every frame held since the last summary, and whether any of
    /// them disagreed with the pair. Nothing at all where no frame has been held since then.
    fn a_summary(&mut self) -> Option<String> {
        if self.since_summary == 0 {
            return None;
        }
        let line = format!(
            "thumbnail held n={} disagrees={}",
            self.since_summary, self.disagrees as u8,
        );
        self.since_summary = 0;
        self.disagrees = false;
        Some(line)
    }
}

/// The held frames this pin has seen (see [`HeldFrames`]), which is one pin's at a time: a registration
/// belongs to the pin that took it up, and its teardown settles what it counted.
static HELD_FRAMES: Mutex<HeldFrames> = Mutex::new(HeldFrames {
    heartbeat: None,
    in_the_second: 0,
    since_summary: 0,
    disagrees: false,
});

/// Count one held frame, and write the heartbeat it is owed where a second of them has passed.
///
/// **Nothing is counted while the trace is off**: this is a trace's own bookkeeping rather than a fact
/// about the pin, so a run without `RHP_PIN_TRACE` pays one read of a flag (see `pin_trace::asked`).
fn note_a_held_frame(disagrees: bool) {
    if !super::pin_trace::asked() {
        return;
    }
    if let Ok(mut held) = HELD_FRAMES.lock() {
        if let Some(line) = held.a_frame(Instant::now(), disagrees) {
            note_pin_trace!("{}", line);
        }
    }
}

/// Write the one summary line a settle is owed, and start counting again from it (see [`HeldFrames`]).
///
/// **The settle is a gesture's** — a drag's release and the capture lost under one are the two ends
/// this road has — and a pin's teardown is the last of them: what a pin held before it went is still
/// worth one line, and a pin's own life is a settle of its own.
pub(super) fn settle_the_held_frames() {
    if let Ok(mut held) = HELD_FRAMES.lock() {
        if let Some(line) = held.a_summary() {
            note_pin_trace!("{}", line);
        }
    }
}

/// How far a source may slide at an edge: the grid a moved position lands on, so that a slow drag
/// towards a screen edge moves the source once every 64 px rather than on every frame.
const SOURCE_STEP: i32 = 64;

/// Why a source's position was asked for: the two reasons a crop can need the source moved, and the
/// word the trace writes for it (see `the_sources_position`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum SourceFit {
    /// The crop left the work area and the source slides the shortest way back onto it.
    Moved,
    /// The crop is longer than the work area on an axis, so no position holds it whole and the source
    /// is put where the crop starts with the work area instead.
    Cramped,
}

/// **Keep the rectangle the Shell actually samples on the screen, by moving the source and never
/// resizing it.** The source is native-sized and does not follow the pin, so a pin at an edge or a
/// corner of the desktop leaves its crop hanging off it — and DWM does not compose a source that is
/// off the screen, which is black where the pin is (see the module note and the two placements it
/// measured dead). A move cannot change the content the crop describes, because the crop is relative
/// to the source's client and its shape is the box's; a resize could, which is why this only ever
/// moves.
///
/// **Nothing moves unless it has to.** The position asked for is the one the source already stands
/// at, unless the crop at that position leaves the work area of the monitor the pin is on: an
/// ordinary centred drag never asks for anything, and neither does a pin standing still. A move is
/// posted, like every other call on a window of mpv's here, and it never carries a size: a moved
/// source is a slid source, not a resized one.
///
/// `a_re_point` is whether this frame is handing over a new pair: the one thing a crop that cannot
/// be held on the screen is traced for, rather than on every tick it stays that way.
fn keep_the_crop_on_the_screen(host: HWND, source: HWND, shown: ThumbnailBox, a_re_point: bool) {
    let Some((at_x, at_y)) = (unsafe { client_origin(source) }) else {
        return;
    };
    // The rectangle DWM samples, in the source's own client coordinates: the crop where there is one,
    // and the whole client where a picture nothing has read hands it that (see `the_thumbnail_box`).
    let crop = shown.source.map_or(
        (0, 0, shown.client.0, shown.client.1),
        |(left, top, right, bottom)| (left, top, right - left, bottom - top),
    );
    let Some((host_x, host_y, host_width, host_height)) = (unsafe { window_box(host) }) else {
        return;
    };
    // **The monitor the pin is on** is the one its band's middle is on, which is the display the user
    // is dragging the window across rather than whichever one a corner is nearest.
    let work = displays::work_area_at(host_x + host_width / 2, host_y + host_height / 2);
    let Some(((to_x, to_y), fit)) = the_sources_position((at_x, at_y), crop, work.region()) else {
        return;
    };

    let (width, height) = (work.right - work.left, work.height());
    let why = match fit {
        SourceFit::Moved => "the crop left the work area",
        SourceFit::Cramped => "the crop is longer than the work area on an axis",
    };
    if (to_x, to_y) == (at_x, at_y) {
        // **A crop that cannot be held on the screen is said out loud** rather than left to draw black
        // in silence, once per settle and not once per tick.
        if fit == SourceFit::Cramped && a_re_point {
            note_pin_trace!(
                "thumbnail source cannot hold the crop on screen at {},{}: {} (work {}x{}+{}+{})",
                at_x,
                at_y,
                why,
                width,
                height,
                work.left,
                work.top,
            );
        }
        return;
    }

    let asked = move_the_source(source, (to_x, to_y));
    note_pin_trace!(
        "thumbnail source moved {},{} -> {},{} because {} (work {}x{}+{}+{} ok {})",
        at_x,
        at_y,
        to_x,
        to_y,
        why,
        width,
        height,
        work.left,
        work.top,
        asked as u8,
    );
}

/// The position the source must stand at for its sampled rectangle to be wholly on `work`, and why —
/// or nothing at all where that rectangle is already there (see `keep_the_crop_on_the_screen`).
///
/// `crop` is the sampled rectangle in the source's own client coordinates, as left, top, width and
/// height, and `work` the work area in screen coordinates. The size is never part of the answer.
fn the_sources_position(
    at: (i32, i32),
    crop: SourceBox,
    work: ScreenRegion,
) -> Option<((i32, i32), SourceFit)> {
    let (left, top, width, height) = crop;
    let (x, x_cramped) = a_source_axis(at.0, (left, width), (work.0, work.2));
    let (y, y_cramped) = a_source_axis(at.1, (top, height), (work.1, work.3));
    let cramped = x_cramped || y_cramped;
    if (x, y) == at && !cramped {
        return None;
    }
    Some((
        (x, y),
        if cramped {
            SourceFit::Cramped
        } else {
            SourceFit::Moved
        },
    ))
}

/// One axis of `the_sources_position`: where the source stood when the crop was inside the work area,
/// the shortest slide back when it was not, and the work area's own start when the crop is longer than
/// the work area and no position can hold it — with whether that is the case.
///
/// **The slide lands on a 64 px grid, away from the edge it was over**, which is what keeps a slow
/// drag's many frames from each moving the source: a step of the drag that asks for the same side of
/// the same grid is answered with the same position and moves nothing. Rounding towards the inside is
/// what keeps a grid step from putting the crop back out of the work area, and the one case it cannot
/// reach — a work area shorter than a grid step — is the bound itself.
fn a_source_axis(at: i32, crop: (i32, i32), work: (i32, i32)) -> (i32, bool) {
    let (offset, extent) = (crop.0, crop.1.max(1));
    let low = work.0 - offset;
    let high = work.1 - offset - extent;
    if at >= low && at <= high {
        return (at, false);
    }
    if high < low {
        // No position holds the whole crop: start it where the work area starts, so the pin draws from
        // the picture's own beginning rather than from a band of desktop (the trace says so).
        return (low, true);
    }
    let wanted = at.clamp(low, high);
    let rounded = wanted.div_euclid(SOURCE_STEP) * SOURCE_STEP;
    let rounded = if wanted == low && rounded != wanted {
        rounded + SOURCE_STEP
    } else {
        rounded
    };
    (rounded.clamp(low, high), false)
}

/// **The only call in this module that ever sizes the source.** One source per film is placed once,
/// at the film's own size, and **never resized**: a resize gesture may move the box and the source's
/// position, and nothing else (see `place_the_source` and the module note). The flags say so — no
/// `SWP_NOSIZE` — and the only caller is the one placement.
fn size_the_source(source: HWND, box_: SourceBox) -> bool {
    let (x, y, width, height) = box_;
    debug_assert!(
        width >= 1 && height >= 1,
        "a source is never asked for an empty box"
    );
    unsafe {
        SetWindowPos(
            source,
            HWND::default(),
            x,
            y,
            width,
            height,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS,
        )
    }
    .is_ok()
}

/// **The only other call in this module that touches the source's geometry, and it cannot change a
/// size**: `SWP_NOSIZE` is the whole of that promise. What a gesture does to the source is a slide
/// back onto the screen, and a slide is not a resize (see `keep_the_crop_on_the_screen`).
fn move_the_source(source: HWND, at: (i32, i32)) -> bool {
    unsafe {
        SetWindowPos(
            source,
            HWND::default(),
            at.0,
            at.1,
            0,
            0,
            SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS,
        )
    }
    .is_ok()
}

/// Put the source at the bottom of the desktop's order without touching its geometry: a source that
/// is behind the pin rather than merely layered at it (see `adopt_the_source`).
fn lower_the_source(source: HWND) {
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
}

/// The corner of a window's client area, on the screen.
unsafe fn client_origin(hwnd: HWND) -> Option<(i32, i32)> {
    let mut origin = POINT::default();
    (unsafe { ClientToScreen(hwnd, &mut origin) }.as_bool()).then_some((origin.x, origin.y))
}

/// The size of a window's client area, or nothing for a window that has gone.
pub(super) unsafe fn client_size(hwnd: HWND) -> Option<(i32, i32)> {
    let mut client = RECT::default();
    (unsafe { GetClientRect(hwnd, &mut client) }.is_ok())
        .then_some((client.right - client.left, client.bottom - client.top))
}

/// A window's whole box, as the `SetWindowPos` that places it names: origin, width and height.
unsafe fn window_box(hwnd: HWND) -> Option<(i32, i32, i32, i32)> {
    let mut rect = RECT::default();
    (unsafe { GetWindowRect(hwnd, &mut rect) }.is_ok()).then_some((
        rect.left,
        rect.top,
        rect.right - rect.left,
        rect.bottom - rect.top,
    ))
}

#[cfg(test)]
mod tests {
    use super::super::tests::the_room_to_itself;
    use super::*;

    /// **A captured frame is gated on the same grid the pin is measured with, and a dark scene passes
    /// it while an empty surface does not** — the gate that keeps a `dc` capture from painting the black
    /// band it exists to prevent (see `the_frames_lit_share`).
    #[test]
    fn a_captured_frame_is_a_picture_or_it_is_not_painted() {
        let (width, height) = (64u32, 64u32);
        let frame = |shade: u8| vec![shade; width as usize * height as usize * 4];

        assert_eq!(
            the_frames_lit_share(&frame(0), width, height),
            0,
            "an all-zero surface is what an empty or dead window answers with"
        );
        const {
            assert!(
                0 < THE_DC_STILL_MIN_LIT_SHARE,
                "and it must be rejected rather than painted as a still"
            )
        };
        assert_eq!(
            the_frames_lit_share(&frame(4), width, height),
            100,
            "a genuinely dark scene is small numbers rather than zeros, so it passes the gate"
        );

        // **One bright point is not a picture**: the grid is a fraction for exactly this reason.
        let mut speck = frame(0);
        speck[32 * 64 * 4 + 32 * 4] = 0xFF;
        assert!(
            the_frames_lit_share(&speck, width, height) < THE_DC_STILL_MIN_LIT_SHARE,
            "a single lit pixel of a stale surface does not pass for a frame"
        );
        assert_eq!(
            the_frames_lit_share(&frame(0), 0, 0),
            0,
            "and a frame that is not a frame has no share at all"
        );

        // The checksum is the number a run compares against a frame decoded at the same second.
        assert_ne!(
            the_frames_checksum(&frame(0)),
            the_frames_checksum(&frame(4)),
            "two different pictures are two different checksums"
        );
        assert_eq!(
            the_frames_checksum(&frame(0)),
            the_frames_checksum(&frame(0)),
            "and the same picture is the same one"
        );
    }

    /// **A step re-points the registration only where the host's box changed**: the source's client and
    /// the film's shape are constants of the pin's life, fixed by the one placement, so a pair is a
    /// function of the box alone (see `settle_the_registration` and `the_thumbnail_box`).
    #[test]
    fn a_step_re_points_a_registration_only_where_the_box_changed() {
        let client = (1000, 560);
        let aspect = Some(2595.0 / 1460.0);
        let source = SourceShape {
            source: 0x1000,
            process: 0x10,
            client: (2731, 1536),
            aspect,
        };
        let shown_at = |box_: (i32, i32)| the_thumbnail_box(box_, source.client, source.aspect);
        let shown = shown_at(client);
        let mut holding = Thumbnail {
            host: 0x2000,
            registration: 9,
            registered: source.source,
            registered_process: source.process,
            shown: Some(shown),
            source: Some(source),
        };

        assert!(
            !the_registration_wants(&holding, shown),
            "a tick that found the same box asks nothing"
        );
        assert!(
            the_registration_wants(&holding, shown_at((1200, 560))),
            "the host's own client is a new destination, and the pair is handed over again"
        );
        assert!(
            the_registration_wants(&holding, shown_at((1001, 560))),
            "one pixel of the box is one pixel the Shell has not drawn"
        );

        holding.shown = Some(shown_at((1200, 560)));
        assert!(
            !the_registration_wants(&holding, shown_at((1200, 560))),
            "and the pair it was just handed is not handed again"
        );
        holding.registration = 0;
        assert!(
            the_registration_wants(&holding, shown),
            "where no registration stands, the pair is what has to be drawn"
        );
    }

    /// **No pair of shapes can leave a bar**: the destination is the host's whole client and the source is
    /// the part of the picture that has that client's shape, so the Shell fills the box from edge to edge
    /// whatever the two shapes are — which is what the containing road could not say, and what the flash
    /// at each end of a drag was (see `the_thumbnail_box`).
    #[test]
    fn the_box_is_filled_and_no_pair_of_shapes_can_leave_a_bar() {
        let shapes = [
            None,
            Some(16.0 / 9.0),
            Some(9.0 / 16.0),
            Some(1.0),
            Some(2.39),
            Some(2595.0 / 1460.0),
            Some(1.0 + 1e-9),
            Some(0.0),
            Some(f64::NAN),
        ];
        for (box_width, box_height) in [
            (1, 1),
            (2, 3),
            (641, 361),
            (999, 1001),
            (1000, 560),
            (1920, 1080),
            (2100, 900),
            (2560, 1080),
            (3840, 2160),
        ] {
            for shape in shapes {
                let picture = a_picture_fitted_to(shape, (box_width, box_height));
                let ((x, y, width, height), _) = the_crop_filling(picture, (box_width, box_height));
                assert!(
                    x >= picture.0
                        && y >= picture.1
                        && width >= 1
                        && height >= 1
                        && x + width <= picture.0 + picture.2
                        && y + height <= picture.1 + picture.3,
                    "{width}x{height}+{x}+{y} is inside the picture {picture:?} at {shape:?}"
                );
                let shown =
                    the_thumbnail_box((box_width, box_height), (box_width, box_height), shape);
                assert_eq!(
                    shown.destination,
                    (0, 0, box_width, box_height),
                    "the destination is the box and never a picture fitted into one, at {shape:?}"
                );
            }
        }
    }

    /// **The invariant the Shell is owed, over every shape pair a broadcast or a probe can produce**:
    /// the destination is the box and the crop is a non-empty rectangle inside the window it was fitted
    /// into, for every client, cached aspect and box aspect — including a shape that is nothing, a
    /// nonsense one, and a client of one pixel (see `ThumbnailBox::checked`).
    #[test]
    fn every_pair_is_a_non_empty_crop_inside_the_window_it_was_fitted_into() {
        let shapes = [
            None,
            Some(16.0 / 9.0),
            Some(9.0 / 16.0),
            Some(1.0),
            Some(2.39),
            Some(2595.0 / 1460.0),
            Some(1.0 + 1e-9),
            Some(0.0),
            Some(-1.0),
            Some(f64::NAN),
            Some(f64::INFINITY),
        ];
        let clients = [
            (1, 1),
            (2, 3),
            (1332, 749),
            (1920, 1080),
            (2880, 1460),
            (3217, 1460),
            (999, 1001),
            (3840, 2160),
        ];
        let boxes = [
            (1, 1),
            (2, 3),
            (641, 361),
            (1000, 560),
            (1332, 749),
            (1920, 1080),
            (2100, 900),
            (3840, 2160),
        ];
        for client in clients {
            for shape in shapes {
                for box_ in boxes {
                    let shown = the_thumbnail_box(box_, client, shape);
                    assert_eq!(
                        shown.destination,
                        (0, 0, box_.0, box_.1),
                        "the destination is the box, at {client:?} {shape:?} {box_:?}"
                    );
                    let Some((left, top, right, bottom)) = shown.source else {
                        assert!(
                            shape.is_none(),
                            "a shape nothing has read is the whole client"
                        );
                        continue;
                    };
                    assert!(
                        right > left && bottom > top,
                        "{right}x{bottom}+{left}+{top} is not empty"
                    );
                    assert!(
                        left >= 0 && top >= 0 && right <= client.0 && bottom <= client.1,
                        "{right}x{bottom}+{left}+{top} is inside the client {client:?}, at \
                         {shape:?} {box_:?}"
                    );
                }
            }
        }
    }

    /// **A held frame prints nothing on its own, and a second of standing still prints exactly one
    /// line.** The trace's own collapse, asked of the counter rather than of a run: a step that finds
    /// the pair unchanged opens the second and says nothing, a whole second of them is one heartbeat,
    /// and the frames are all still summed for the settle (see `HeldFrames`).
    #[test]
    fn a_held_frame_prints_nothing_and_a_second_of_standing_still_prints_one_line() {
        let mut held = HeldFrames::default();
        let opened = Instant::now();
        let at = |ms: u64| opened + Duration::from_millis(ms);

        assert_eq!(
            held.a_frame(at(0), false),
            None,
            "a step that does not move the pair prints no line"
        );
        assert_eq!(
            held.a_frame(at(999), false),
            None,
            "and neither does the rest of the second it opened"
        );
        assert_eq!(
            held.a_frame(at(1000), false).as_deref(),
            Some("thumbnail held heartbeat n=3 disagrees=0"),
            "a second of standing still is exactly one line, and it carries that second's frames"
        );
        assert_eq!(
            held.a_frame(at(1500), true),
            None,
            "the next second is still inside its first half, and a disagreement is not a line"
        );
        assert_eq!(
            held.a_frame(at(2000), false).as_deref(),
            Some("thumbnail held heartbeat n=2 disagrees=1"),
            "one heartbeat a second, and whether any of them disagreed with the pair"
        );

        assert_eq!(
            held.a_summary().as_deref(),
            Some("thumbnail held n=5 disagrees=1"),
            "the settle sums every frame held since the last summary, not just the last second"
        );
        assert_eq!(
            held.a_summary(),
            None,
            "and a settle with nothing held since the last one prints nothing at all"
        );
    }

    /// The crop the pair would hand the Shell, as the rectangle DWM samples in the source's own
    /// client: the crop where there is one, and the whole client where there is not.
    fn a_crop_of(box_: (i32, i32), client: (i32, i32), aspect: Option<f64>) -> SourceBox {
        let shown = the_thumbnail_box(box_, client, aspect);
        shown.source.map_or(
            (0, 0, shown.client.0, shown.client.1),
            |(left, top, right, bottom)| (left, top, right - left, bottom - top),
        )
    }

    /// The rectangle the Shell actually samples, on the screen: where the source stands plus the crop
    /// inside its client.
    fn a_sampled_rect(at: (i32, i32), crop: SourceBox) -> ScreenRegion {
        (
            at.0 + crop.0,
            at.1 + crop.1,
            at.0 + crop.0 + crop.2,
            at.1 + crop.1 + crop.3,
        )
    }

    /// **A pin flush with any edge or corner of the desktop has its crop slid back onto the work
    /// area**, and the same on a monitor left of the primary one whose work area does not start at
    /// zero — the two shapes a screen edge comes in (see `the_sources_position`).
    #[test]
    fn a_crop_at_any_edge_or_corner_ends_up_inside_the_work_area() {
        let film = (1280, 720);
        let box_ = (640, 360);
        let crop = a_crop_of(box_, film, Some(16.0 / 9.0));
        for work in [(0, 0, 1920, 1080), (-1920, 0, 0, 1040)] {
            let xs = [work.0, (work.0 + work.2 - box_.0) / 2, work.2 - box_.0];
            let ys = [work.1, (work.1 + work.3 - box_.1) / 2, work.3 - box_.1];
            for x in xs {
                for y in ys {
                    // The source was adopted at the pin's own corner, which is the box's.
                    let at = (x, y);
                    let (to, _) =
                        the_sources_position(at, crop, work).unwrap_or((at, SourceFit::Moved));
                    let sampled = a_sampled_rect(to, crop);
                    assert!(
                        sampled.0 >= work.0
                            && sampled.1 >= work.1
                            && sampled.2 <= work.2
                            && sampled.3 <= work.3,
                        "a crop {sampled:?} is inside the work area {work:?} for a pin at {at:?}"
                    );
                }
            }
        }
    }

    /// **A centred drag asks the source for no position at all.** The source stands centred on the
    /// desktop with its whole picture inside the work area, and every box a drag passes through crops
    /// inside that picture — so nothing the hand does to the pin's shape moves the source (see
    /// `keep_the_crop_on_the_screen`).
    #[test]
    fn a_centred_drag_asks_the_source_for_no_position() {
        let film = (1280, 720);
        let work = (0, 0, 1920, 1080);
        let at = ((work.2 - film.0) / 2, (work.3 - film.1) / 2);
        for box_ in [
            (640, 360),
            (700, 394),
            (800, 450),
            (1000, 560),
            (1280, 720),
            (1200, 675),
            (2263, 1274),
        ] {
            let crop = a_crop_of(box_, film, Some(16.0 / 9.0));
            assert_eq!(
                the_sources_position(at, crop, work),
                None,
                "a centred drag's {box_:?} moves the source not at all"
            );
        }
    }

    /// **A move lands on the grid, and a source that has been moved is a source that stays put**: the
    /// position asked for is inside the work area, and the next frame at that same position asks for
    /// nothing — which is what keeps a slow drag at an edge from moving the source on every frame.
    #[test]
    fn a_move_lands_on_the_grid_and_is_then_stable() {
        let work = (0, 0, 1920, 1080);
        let crop = (97, 41, 1000, 600);
        for at in [(-400, -400), (1800, 900), (-97, 500), (900, 0)] {
            let (to, _) =
                the_sources_position(at, crop, work).expect("the crop left the work area");
            let sampled = a_sampled_rect(to, crop);
            assert!(
                sampled.0 >= work.0
                    && sampled.1 >= work.1
                    && sampled.2 <= work.2
                    && sampled.3 <= work.3,
                "the moved crop {sampled:?} is on the work area {work:?}"
            );
            for (from, to) in [(at.0, to.0), (at.1, to.1)] {
                if from != to {
                    assert_eq!(
                        to % SOURCE_STEP,
                        0,
                        "an axis that slides lands on the {SOURCE_STEP} px grid, at {to}"
                    );
                }
            }
            assert_eq!(
                the_sources_position(to, crop, work),
                None,
                "and a source already there is not asked to move again"
            );
        }
    }

    /// **A crop longer than the work area cannot be held inside it, and the trace is told so rather
    /// than the pin drawing black in silence**: the source is put where the crop starts with the work
    /// area, and the axis that does fit is still slid onto it (a 4K film on a smaller desktop).
    #[test]
    fn a_crop_longer_than_the_work_area_is_cramped_and_says_so() {
        let work = (0, 0, 1920, 1040);
        assert_eq!(
            the_sources_position((300, 200), (0, 0, 3840, 2160), work),
            Some(((0, 0), SourceFit::Cramped)),
            "both axes are longer than the work area, so the crop starts where it does"
        );
        assert_eq!(
            the_sources_position((0, 0), (0, 0, 3840, 2160), work),
            Some(((0, 0), SourceFit::Cramped)),
            "and a source already there is still answered as cramped, so the trace can say so"
        );
        let Some(((x, y), fit)) = the_sources_position((300, 900), (0, 0, 3840, 720), work) else {
            panic!("the crop is longer than the work area on one axis");
        };
        assert_eq!(
            (x, fit),
            (0, SourceFit::Cramped),
            "the axis longer than the work area is aligned with its start, and the trace says so"
        );
        assert!(
            y >= work.1 && y + 720 <= work.3,
            "and the axis that fits is still moved onto the work area, at {y}"
        );
    }

    /// **A second placement is refused out loud and counted**, so no road can reach one silently:
    /// the trace says `source re-placed`, and the seam a test reads counts it (see
    /// `place_the_source`).
    #[test]
    fn a_second_placement_is_refused_out_loud_and_counted() {
        let _room = the_room_to_itself();
        clear_the_source_placements();
        let placed = SourceShape {
            source: 0x1000,
            process: 0x10,
            client: (1920, 1080),
            aspect: Some(16.0 / 9.0),
        };
        assert_eq!(
            place_the_source(
                HWND(0x1000 as *mut _),
                HWND(0x2000 as *mut _),
                0x10,
                Some(placed),
                NativeSize::Known((1920, 1080)),
            ),
            None,
            "a source that has been placed is not placed again"
        );
        assert_eq!(
            the_source_replacements(),
            1,
            "and the refusal is counted rather than swallowed"
        );
        clear_the_source_placements();
    }

    /// The registration is nobody's where none was taken: a thumbnail record with no pair stands in no
    /// host, which is what a host that has not been pointed at yet reads as.
    #[test]
    fn nothing_stands_in_a_host_until_a_pair_is_handed_over() {
        assert_eq!(Thumbnail::NOTHING.registration, 0);
        assert_eq!(Thumbnail::NOTHING.host, 0);
        assert_eq!(Thumbnail::NOTHING.shown, None);
    }

    /// **A film adopted after another is never cropped by the shape the last one left.** A shape is the
    /// one placement of one source window in one player (see `SourceShape`), and the record out of
    /// which a pair is built is invalidated with the film: a left-over shape is given up before any
    /// frame of the next film is built, so that film's first frames are either its own shape or no
    /// crop at all — never the previous film's crop.
    #[test]
    fn a_film_adopted_after_another_is_never_cropped_by_the_old_films_shape() {
        let _room = the_room_to_itself();

        // Film A: a 2.39:1 picture placed once at its own size.
        let a = SourceShape {
            source: 0xA1,
            process: 0xAA,
            client: (2595, 1460),
            aspect: Some(2595.0 / 1460.0),
        };

        // Film B is the next player's window, and its own size has not been read yet.
        SOURCE.store(0xB2, Ordering::SeqCst);
        SOURCE_PROCESS.store(0xBB, Ordering::SeqCst);

        // A record left as it was has no claim on B's frames: it is given up where they are built.
        let mut held = Thumbnail {
            host: 0x2000,
            registration: 0,
            registered: 0,
            registered_process: 0,
            shown: None,
            source: Some(a),
        };
        forget_the_shape_of_another_film(&mut held);
        assert_eq!(
            held.source, None,
            "the shape A's placement left is given up with A"
        );
        let first = the_thumbnail_box((1000, 560), (1687, 949), held.source.and_then(|s| s.aspect));
        assert_eq!(
            first.source, None,
            "B's first frame is not cropped by a shape nobody has read for B"
        );
        assert_eq!(
            first.destination,
            (0, 0, 1000, 560),
            "and the whole box is still what B is drawn in"
        );

        // When B's own size is known, the one placement is B's own and the crop is B's shape.
        let b_size = (1920, 1080);
        assert_eq!(
            the_one_placement(None, Some((0, 0)), Some(b_size), (3440, 1440)),
            Some((0, 0, 1920, 1080)),
            "the film's own size, and not the shape the film before it was placed at"
        );
        let b = SourceShape {
            source: 0xB2,
            process: 0xBB,
            client: b_size,
            aspect: NativeSize::Known(b_size).aspect(),
        };
        let second = the_thumbnail_box((1000, 560), b.client, b.aspect);
        let (left, top, right, bottom) = second
            .source
            .expect("B's own picture is what its frames are cropped by");
        let crop_aspect = (right - left) as f64 / (bottom - top) as f64;
        assert!(
            (crop_aspect - 16.0 / 9.0).abs() < 0.01,
            "B is cropped by B's 16:9 picture, not by A's 2.39 one: {crop_aspect}"
        );

        // And the adoption of B resets the record it is made into, whatever the last film left.
        if let Ok(mut recorded) = THUMBNAIL.lock() {
            recorded.source = Some(a);
        }
        SOURCE.store(0, Ordering::SeqCst);
        SOURCE_PROCESS.store(0, Ordering::SeqCst);
        note_the_source(HWND(0xB2 as *mut _), 0xBB);
        assert_eq!(
            THUMBNAIL.lock().ok().and_then(|held| held.source),
            None,
            "a source adopted is a record emptied of the film that ended"
        );

        forget_the_thumbnail();
    }

    /// **A probe measures the film the player was begun for and no other.** The pin's file is the
    /// outgoing one for the moment a step's source is adopted, and a probe read without the player
    /// beside it would place the incoming film at the outgoing film's size and shape — the previous
    /// film's shape cropping the next one's first frames (see `the_films_own_size`).
    #[test]
    fn a_probe_answers_only_for_the_film_the_player_was_begun_for() {
        let _room = the_room_to_itself();
        let path = PathBuf::from("a-probed-film.mkv");
        let key = VideoGeometryKey {
            path: path.clone(),
            version: file_version(&path),
        };
        video_geometry_cache().insert(
            key.clone(),
            ProbedGeometry::Measured(VideoGeometry {
                width: 2595,
                height: 1460,
                frame_width: 2595,
                frame_height: 1460,
                crop: None,
                duration: None,
                subtitles: SubtitleStreams::default(),
                sidecar: None,
                derived: None,
                subtitle_codecs: Vec::new(),
                attachment_codecs: Vec::new(),
                subtitle_extraction_failed: false,
            }),
        );

        let pid = 0x5eed;
        hover_film_host::note_hover_film_host_player(pid, &path);
        assert_eq!(
            the_films_own_size(pid),
            Some((2595, 1460)),
            "the film the player was begun for is the film the probe answers for"
        );
        assert_eq!(
            the_films_own_size(pid + 1),
            None,
            "and another player — the next film of a step — is answered nothing rather than the \
             size of the film before it"
        );

        unsafe { hover_film_host::take_down_the_hover_film_host() };
        video_geometry_cache().remove(&key);
    }

    /// **A film neither the player nor the probe answers for is placed once at the largest box the road
    /// can justify**, and the road gives the player a bounded wait to answer in before it does: a first
    /// ask at once, one per `LATE_NATIVE_EVERY` after it, and the placement at the end of
    /// `LATE_NATIVE_WAIT`. A film still waiting is placed at nothing at all.
    #[test]
    fn a_film_with_no_answer_is_placed_at_the_largest_box_after_its_wait() {
        assert_eq!(
            NativeSize::NotYet.box_size((3440, 1440)),
            None,
            "a film still waiting is not placed on a guess"
        );
        assert_eq!(
            NativeSize::Known((1920, 1080)).box_size((3440, 1440)),
            Some((1920, 1080)),
            "the player's answer is the box"
        );
        assert_eq!(
            NativeSize::Late((1920, 1080)).box_size((3440, 1440)),
            Some((1920, 1080)),
            "and the answer that came after adoption is the same box, placed once"
        );
        assert_eq!(
            NativeSize::Unobtainable.box_size((3440, 1440)),
            Some((3440, 1440)),
            "and a film answered nothing is the largest box the road can justify"
        );
        assert_eq!(
            NativeSize::Unobtainable.aspect(),
            None,
            "which is a box and not the film's shape"
        );
        assert_eq!(
            NativeSize::Known((1920, 1080)).aspect(),
            Some(16.0 / 9.0),
            "while the film's own size is the shape its crop is cut by"
        );

        assert!(
            a_late_placement_is_due(Duration::ZERO, None),
            "the first pass after adoption asks the player"
        );
        assert!(
            !a_late_placement_is_due(Duration::from_millis(10), Some(Duration::ZERO)),
            "and the next passes wait out the pace rather than asking on every one"
        );
        assert!(
            a_late_placement_is_due(Duration::from_millis(100), Some(LATE_NATIVE_EVERY)),
            "one ask per pace"
        );
        assert!(
            a_late_placement_is_due(LATE_NATIVE_WAIT, Some(Duration::from_millis(1))),
            "and the end of the wait is the last ask, which places the box where nothing answers"
        );
    }

    /// **The invisibility knob, and the style it is the whole of.** It is on unless it says exactly `0`;
    /// on, the source wears `WS_EX_LAYERED` beside the three bits this road always asserts and keeps the
    /// app-window and topmost bits off; off, the layered bit is not on it at all (see
    /// `source_is_invisible_from` and `source_ex_style`).
    #[test]
    fn the_invisible_knob_is_on_unless_it_says_zero_and_is_the_layered_bit() {
        assert_eq!(INVISIBLE_VARIABLE, "RHP_THUMBNAIL_INVISIBLE");
        let named = |value: Option<&str>| source_is_invisible_from(value.map(std::ffi::OsStr::new));
        assert!(
            named(None),
            "this road's default is the source nothing of which is seen"
        );
        assert!(named(Some("1")));
        assert!(
            !named(Some("0")),
            "and `0` is the visible road, byte for byte"
        );
        assert!(
            named(Some("")),
            "a value that names nothing is not read as off"
        );
        assert!(
            named(Some("off")),
            "only the one value turns it off, and it is `0`"
        );

        let inert = WS_EX_APPWINDOW.0 as isize | WS_EX_TOPMOST.0 as isize;
        let invisible = source_ex_style(inert, true);
        assert_eq!(
            invisible & (WS_EX_LAYERED.0 as isize),
            WS_EX_LAYERED.0 as isize
        );
        for bit in [WS_EX_TOOLWINDOW, WS_EX_NOACTIVATE, WS_EX_TRANSPARENT] {
            assert_eq!(
                invisible & (bit.0 as isize),
                bit.0 as isize,
                "{bit:?} is kept"
            );
        }
        assert_eq!(
            invisible & (WS_EX_APPWINDOW.0 as isize),
            0,
            "and the app-window bit mpv opens with is gone: this film is not a window of the user's"
        );
        assert_eq!(
            invisible & (WS_EX_TOPMOST.0 as isize),
            0,
            "nor is it held over the pin it is seen in"
        );

        let visible = source_ex_style(WS_EX_LAYERED.0 as isize | inert, false);
        assert_eq!(
            visible & (WS_EX_LAYERED.0 as isize),
            0,
            "the visible road's source is not layered, whatever it was wearing"
        );
    }

    /// The invisible road is stood in by a test, and is the default where nothing is stood in (see
    /// `the_source_is_invisible`).
    #[test]
    fn a_test_stands_the_invisible_road_it_is_about_in() {
        assert!(the_source_is_invisible());
        stand_the_invisible_road_in(false);
        assert!(!the_source_is_invisible());
        stand_the_invisible_road_in(true);
    }

    /// **The film's own size is one read of the player's `video-params`**: `dw`/`dh` are the picture's
    /// own size and `w`/`h` the coded one, which is the fallback for a player that writes only those;
    /// and a player this app does not hold is not asked at all (see `the_players_picture_size`).
    #[test]
    fn the_picture_size_is_one_read_of_the_video_params_node() {
        let _room = the_room_to_itself();
        let pid = 0x5eed;
        mpv_ipc::remember_pipe(pid, "rhp-test-picture-size", false);
        mpv_ipc::clear_reads();

        // An anamorphically coded film: the display size is what a source is placed at, because that is
        // the shape the player would draw it in.
        mpv_ipc::answer_a_read(
            pid,
            "video-params",
            Some(
                r#"{"pixelformat":"yuv420p","w":1440,"h":1080,"dw":1920,"dh":1080,"aspect":1.777778}"#,
            ),
        );
        assert_eq!(the_players_picture_size(pid), Some((1920, 1080)));
        assert_eq!(
            mpv_ipc::noted_reads(),
            vec![format!("{pid}:video-params")],
            "asked once, and for the node: a read per field is more round trips at adoption"
        );

        mpv_ipc::answer_a_read(pid, "video-params", Some(r#"{"w":640,"h":360}"#));
        assert_eq!(
            the_players_picture_size(pid),
            Some((640, 360)),
            "a player that writes no display size is read at its coded one"
        );

        // A phone's film: the frame is stored landscape and drawn portrait.
        mpv_ipc::answer_a_read(
            pid,
            "video-params",
            Some(r#"{"w":1920,"h":1080,"dw":1920,"dh":1080,"rotate":90}"#),
        );
        assert_eq!(
            the_players_picture_size(pid),
            Some((1080, 1920)),
            "a film the player turns on its side is asked for at the shape it is drawn"
        );

        mpv_ipc::answer_a_read(pid, "video-params", Some(r#"{"w":0,"h":0,"dw":0,"dh":0}"#));
        assert_eq!(
            the_players_picture_size(pid),
            None,
            "a player with no picture in its output has no size to place anything at"
        );
        mpv_ipc::answer_a_read(pid, "video-params", None);
        assert_eq!(the_players_picture_size(pid), None);
        assert_eq!(
            the_players_picture_size(0),
            None,
            "a player this app does not hold is not asked"
        );

        mpv_ipc::forget_pipe(pid);
        mpv_ipc::clear_reads();
    }

    /// **A film nobody has sized yet is waited for, and nothing is placed on a guess**: adoption asks
    /// the player a bounded number of times and then takes the probe's own measurement of the same film
    /// (see `the_films_native_size`).
    #[test]
    fn a_film_that_is_not_yet_sized_is_waited_for_and_not_guessed() {
        let _room = the_room_to_itself();
        let pid = 0x5eed;
        mpv_ipc::remember_pipe(pid, "rhp-test-unsized-film", false);
        mpv_ipc::clear_reads();
        mpv_ipc::answer_a_read(pid, "video-params", None);

        assert_eq!(
            the_films_native_size(pid),
            the_films_own_size(pid),
            "a player that never answers is the probe's own measurement, not a box this app guessed"
        );
        assert_eq!(
            mpv_ipc::noted_reads().len(),
            NATIVE_SIZE_ASKS as usize,
            "and the wait is bounded: one read per ask and no more"
        );
        assert_eq!(
            the_one_placement(None, Some((100, 200)), None, (3440, 1440),),
            None,
            "a film of no known size is never placed at the pin's box to be corrected later"
        );

        mpv_ipc::forget_pipe(pid);
        mpv_ipc::clear_reads();
    }

    /// **The one placement: the film's own size at the pin's corner, clamped to what the window system
    /// will take — and nothing at all the second time.** A press, a drag, a release, a settle, a
    /// maximize and a zoom all arrive here with a source that has been placed, and every one of them is
    /// answered with nothing: what the crop is cut from is a constant for the pin's life (see
    /// `place_the_source`).
    #[test]
    fn a_source_takes_one_placement_and_never_a_second_one() {
        let native = Some((1920, 1080));
        let max_track = (3440, 1440);

        assert_eq!(
            the_one_placement(None, Some((100, 200)), native, max_track,),
            Some((100, 200, 1920, 1080)),
            "the film's own size, at the pin's corner"
        );

        let placed = Some(SourceShape {
            source: 0x1000,
            process: 0x10,
            client: (1920, 1080),
            aspect: Some(16.0 / 9.0),
        });
        for native in [None, Some((1920, 1080)), Some((3840, 2160))] {
            assert_eq!(
                the_one_placement(placed, Some((100, 200)), native, max_track),
                None,
                "a placement already taken is not taken again, whatever the film is"
            );
        }
        assert_eq!(
            the_one_placement(None, None, native, max_track),
            None,
            "no host is no corner to place at"
        );
        assert_eq!(
            the_one_placement(None, Some((0, 0)), None, max_track),
            None,
            "a native size that is not known yet is not the box: it is waited for"
        );
        assert_eq!(
            the_one_placement(None, Some((0, 0)), Some((8192, 4608)), max_track),
            Some((0, 0, 3440, 1440)),
            "a film bigger than the desktop is clamped to what a window may be"
        );
        assert_eq!(
            the_one_placement(None, Some((0, 0)), Some((0, 0)), max_track),
            Some((0, 0, 1, 1)),
            "and never to nothing at all"
        );
    }

    /// **A picture is fitted into a client of another shape and centred in it, and the box a pin draws it
    /// in is filled by cropping that picture instead**: 16:9 in a 4:3 client is 60 px of letterbox above
    /// and below, 16:9 in a wider one is pillars either side, and a client of the picture's own shape is
    /// the picture whole (see `a_picture_fitted_to`, which is the picture's rectangle and not the box's).
    #[test]
    fn a_picture_is_fitted_into_its_client_and_centred() {
        let wide = Some(16.0 / 9.0);

        assert_eq!(a_picture_fitted_to(wide, (640, 480)), (0, 60, 640, 360));
        assert_eq!(a_picture_fitted_to(wide, (2100, 900)), (250, 0, 1600, 900));
        assert_eq!(
            a_picture_fitted_to(wide, (1920, 1080)),
            (0, 0, 1920, 1080),
            "a client the film's own shape is the film and nothing beside it"
        );
    }

    /// **A crop fills the box, so the picture gives up the side the box's shape does not want**: a picture
    /// wider than the box loses its sides, one no wider loses its top and bottom, a box of the picture's
    /// own shape loses nothing, and the strip the rounding could not split in two is one pixel of crop
    /// rather than a bar (see `the_crop_filling`).
    #[test]
    fn a_crop_fills_the_box_whatever_shape_the_box_is() {
        // Picture wider than the box: 3200x1000 kept whole down, cropped across to 1600x1000 in a
        // 1600x1000 box, and to 800x1000 in a 800x1000 one.
        assert_eq!(
            the_crop_filling((0, 0, 3200, 1000), (1600, 1000)).0,
            (800, 0, 1600, 1000)
        );
        assert_eq!(
            the_crop_filling((0, 0, 3200, 1000), (800, 1000)).0,
            (1200, 0, 800, 1000)
        );
        // No wider than the box: the width is kept and the top and bottom go.
        assert_eq!(
            the_crop_filling((0, 0, 1000, 4000), (1000, 1000)).0,
            (0, 1500, 1000, 1000)
        );
        // A box the picture's own shape is the picture whole, and says so.
        assert_eq!(
            the_crop_filling((311, 0, 2595, 1460), (2595, 1460)),
            ((311, 0, 2595, 1460), false),
            "nothing is cropped where the two shapes agree"
        );
        // A crop is centred in the picture, wherever the picture stands in the client.
        assert_eq!(
            the_crop_filling((311, 0, 2595, 1460), (1000, 560)).0,
            (311, 3, 2595, 1453),
            "the odd pixel of the rounding is the picture's own and never a bar"
        );
        // A box of one pixel is the picture's largest square, and never nothing.
        assert_eq!(
            the_crop_filling((0, 0, 1440, 1080), (1, 1)).0,
            (180, 0, 1080, 1080)
        );
        assert_eq!(the_crop_filling((0, 0, 1, 4000), (1, 1)).0, (0, 1999, 1, 1));
        // **A box larger than the picture on either axis is the picture whole**, upscaled by the Shell,
        // and the trace says so: cropping a picture that is already being magnified only throws pixels
        // away, and the destination is the box whatever the crop is, so there is no bar either way.
        for box_ in [(4000, 1), (1920, 1080), (1921, 1080), (5, 100), (4, 101)] {
            let ((left, top, width, height), upscaled) = the_crop_filling((0, 0, 4, 100), box_);
            assert_eq!(
                (left, top, width, height),
                (0, 0, 4, 100),
                "{box_:?} is over the picture, so the whole picture is what fills it"
            );
            assert!(
                upscaled,
                "and the trace says the crop was cut to the picture"
            );
        }
    }

    /// **A still a relaunch paints is cut by the box's own shape, exactly as the registration's crop is**:
    /// the film fitted into the source's client and then the rectangle of it that has the host's shape, so
    /// the cover the user sees is the frame that was just on screen and not a second guess at its shape
    /// (see `the_pictures_crop`).
    #[test]
    fn a_captured_still_is_cropped_by_the_boxs_shape() {
        let wide = Some(16.0 / 9.0);
        // A 1920x1080 frame on a square band: the height is kept and the sides go, centred.
        assert_eq!(
            the_pictures_crop((1920, 1080), (400, 400), wide).0,
            (420, 0, 1080, 1080)
        );
        // The same frame on the 320x240 band a pin is built at: the largest 4:3 rectangle of it.
        assert_eq!(
            the_pictures_crop((1920, 1080), (320, 240), wide).0,
            (240, 0, 1440, 1080)
        );
        // The film's own shape inside its client: a 4:3 film in a 16:9 client is pillars either side,
        // and the crop is of the picture rather than of the client it stands in.
        assert_eq!(
            the_pictures_crop((1920, 1080), (400, 400), Some(0.75)).0,
            (555, 135, 810, 810)
        );
        // A box the film's own shape is the picture whole.
        assert_eq!(
            the_pictures_crop((1920, 1080), (1920, 1080), wide).0,
            (0, 0, 1920, 1080)
        );
    }

    /// **A shape that rounds is a picture a pixel out and never a pixel over its client**: the measured
    /// 2595x1460 picture in a 1000x560 client is 995 px wide in it, and a client of one pixel is the
    /// client itself (see `a_picture_fitted_to`).
    #[test]
    fn a_fitted_picture_rounds_inside_its_client_and_never_out_of_it() {
        let measured = Some(2595.0 / 1460.0);
        let (x, y, width, height) = a_picture_fitted_to(measured, (1000, 560));
        assert_eq!((x, y, width, height), (2, 0, 995, 560));
        assert!(
            x + width <= 1000 && y + height <= 560,
            "never a pixel over the client"
        );
        assert!(x >= 0 && y >= 0, "and never one outside it");

        assert_eq!(a_picture_fitted_to(Some(1.7773973), (1, 1)), (0, 0, 1, 1));
        assert_eq!(a_picture_fitted_to(Some(16.0 / 9.0), (1, 1)), (0, 0, 1, 1));
        assert_eq!(
            a_picture_fitted_to(Some(16.0 / 9.0), (0, 0)),
            (0, 0, 1, 1),
            "a client of nothing is asked for what a window can be"
        );
        assert_eq!(a_picture_fitted_to(None, (800, 600)), (0, 0, 800, 600));
        assert_eq!(
            a_picture_fitted_to(Some(f64::NAN), (800, 600)),
            (0, 0, 800, 600)
        );
    }

    /// **The thumbnail shows the picture's rectangle and never the whole client**: the crop is the part
    /// of the picture that fills the box, and a picture nothing has read is the whole client into the
    /// whole box — the fallback for a shape this app cannot crop by (see `the_thumbnail_box`).
    #[test]
    fn a_thumbnail_shows_the_pictures_rectangle() {
        // The measured source: a 16:9 film in a 3217x1460 client, 311 px of pillar either side.
        let aspect = Some(2595.0 / 1460.0);
        let shown = the_thumbnail_box((1000, 560), (3217, 1460), aspect);
        assert_eq!(shown.client, (3217, 1460));
        assert_eq!(
            shown.source,
            Some((311, 3, 2906, 1456)),
            "the picture is taken out of the source's margins and cropped to the box's shape, not \
             stretched with them"
        );
        assert_eq!(
            shown.destination,
            (0, 0, 1000, 560),
            "and the box is filled by it: no pillar, and no letterbox above or below"
        );

        let nothing_read = the_thumbnail_box((1000, 560), (3217, 1460), None);
        assert_eq!(
            nothing_read.source, None,
            "a picture nothing has read is not cropped by"
        );
        assert_eq!(nothing_read.destination, (0, 0, 1000, 560));
    }

    /// **A press, a drag, a release and the settle after it ask the source for no box at all, and every
    /// one of them crops by the one picture adoption fixed.** The box is the only thing that changes, so
    /// the crop is a function of it alone, the crop never reaches outside the picture, and a box over the
    /// film's own size is the whole picture upscaled by the Shell (see `place_the_source`,
    /// `the_one_placement` and `the_thumbnail_box`).
    #[test]
    fn a_press_a_drag_a_release_and_a_settle_ask_the_source_for_no_box() {
        // The one placement, as adoption left it: a 1920x1080 film, placed at its own size.
        let source = SourceShape {
            source: 0x1000,
            process: 0x10,
            client: (1920, 1080),
            aspect: Some(16.0 / 9.0),
        };
        let picture = a_picture_fitted_to(source.aspect, source.client);
        let placed = Some(source);
        let boxes = [
            (640, 360),
            (867, 487),
            (1000, 560),
            (1332, 749),
            (1920, 1080),
            (2263, 1274),
            (3840, 2160),
        ];

        for box_ in boxes {
            assert_eq!(
                the_one_placement(placed, Some((100, 200)), Some((1920, 1080)), (3440, 1440),),
                None,
                "{box_:?} is a moment of the gesture, and the gesture asks for no box"
            );
            let shown = the_thumbnail_box(box_, source.client, source.aspect);
            assert_eq!(
                shown.destination,
                (0, 0, box_.0, box_.1),
                "the box is filled, edge to edge, at {box_:?}"
            );
            assert_eq!(
                shown.client, source.client,
                "the crop's window is the one placement's, never a box the pin was dragged to"
            );
            let (left, top, right, bottom) =
                shown.source.expect("the picture is what is cropped by");
            assert!(
                left >= picture.0
                    && top >= picture.1
                    && right <= picture.0 + picture.2
                    && bottom <= picture.1 + picture.3,
                "{box_:?} crops inside the film's own picture"
            );
            assert_eq!(
                shown,
                the_thumbnail_box(box_, source.client, source.aspect),
                "and the same box is always the same pair: nothing else is an input"
            );
        }

        let over = the_thumbnail_box((3840, 2160), source.client, source.aspect);
        assert_eq!(
            over.source,
            Some((
                picture.0,
                picture.1,
                picture.0 + picture.2,
                picture.1 + picture.3
            )),
            "a box over the film's own size is the picture whole, upscaled by the Shell"
        );
        assert!(over.clamped, "which the trace's clamp field says");
        assert_eq!(
            over.destination,
            (0, 0, 3840, 2160),
            "and the box is still filled, so there is no bar either way"
        );
    }
}

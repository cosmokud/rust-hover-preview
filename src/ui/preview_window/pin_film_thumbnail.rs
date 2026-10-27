//! The pin's film drawn by the Desktop Window Manager rather than nested in a window of this app's.
//!
//! **This is the road a pin's mpv film is drawn on**: `RHP_MPV_THUMBNAIL_PIN=0` is the escape hatch
//! back to the child road, and it is byte for byte the road this app stood in before (see
//! `pin_film_host::MPV_THUMBNAIL_PIN_VARIABLE`). mpv is launched with no `--wid` and makes a window
//! of the desktop's of its own; the pin's film host asks the Shell to draw that window into itself
//! (`DwmRegisterThumbnail`), and what follows a resize is one rectangle rather than two windows
//! agreeing about a box.
//!
//! **Why the film has to be a window of the desktop's.** A thumbnail is taken of a window's
//! composition surface, and a `WS_CHILD` has none of its own — it is drawn into its parent's — so a
//! film that is to be thumbnailed cannot be mpv's child under the host. That is the whole of why
//! this road launches mpv without `--wid` (see `video_launch::mpv_args`).
//!
//! **The source is placed once, at the film's own size, and never again.** That is the whole of this
//! road now, and it is where two flashes were: a 19.4 s trace of the road that resized the source at
//! each end of a drag had 326 of 2,821 thumbnail frames whose crop's window disagreed with the
//! source's live client — every one of them within a frame or two of a re-place, 1 frame (7 ms) at
//! the press with the crop 2.1–2.4× too small, and 72–130 frames (145 ms) at the release with the
//! crop sampling 2.9–3.3× past the live client, which is the zoom-out with black around it. The
//! middle of a drag was perfect, because nothing moved there. So nothing moves here: the source takes
//! its box in the adoption pass, the picture adoption read is a constant for the pin's life, and a
//! resize, a maximize, a restore, a zoom, a walk, a press, a release and a settle are all **one
//! `DwmUpdateThumbnailProperties`** — the destination and the crop, and nothing cross-process at all
//! (see `place_the_source` and `the_band_was_placed`).
//!
//! **The size is the film's own picture size, never the pin's box.** `native` (the default, see
//! [`SOURCE_SIZE_VARIABLE`]) places the source at the size the film would be drawn at were it not
//! fitted to a window, from the player's own `video-params` — or from the probe's measurement of the
//! same film where the player has not opened the file yet, which is waited for rather than guessed
//! at. A pin at or under that size is a *downscale* the Shell makes, which is sharp at every box the
//! hand pulls it to; a pin over it is the upscale the player itself would have to make, from the
//! film's own pixels and not from a box three times a small pin. `box` is the A/B's other arm and
//! places the same one source at the pin's box instead.
//!
//! **What a thumbnail shows is the source's picture, not its whole client** (`RHP_THUMBNAIL_PICTURE`,
//! `picture` by default): the rectangle the film is drawn in is cropped out of the source to **the
//! box's own shape**, and the whole box is handed over as the destination, so the film fills its pin
//! (see `the_thumbnail_box`). `client` is the road this began on and it is dead: it stretched the
//! whole client, margins included, into the box, which is the pillar of black that read as a pin
//! zoomed out (measured: `client 916x515 … picture 2595x1460 margins l311t0r311b0` — 311 px a side
//! stretched into the picture) and the upscale that read as one zoomed in.
//!
//! **Filling the box is not a detail of the crop, it is what makes a bar impossible.** The road this
//! replaces fitted the picture into the box at the picture's own shape and left the rest to the
//! host's black: a destination *smaller* than the client whenever the box's aspect and the film's
//! were not the same, read as the film squashed into the middle of the pin with a pillar either side.
//! The destination is the client and the crop is the part of the picture that has the client's
//! shape, so there is no rectangle left over to read as the wrong thing — whatever order anything is
//! ordered in, and whatever the two shapes are.
//!
//! **The composition invariant is one line**: the destination is the host's whole client, so the
//! Shell draws every pixel of the pin, and the host's own erase paints that same client under it —
//! which is what a client the Shell has not composed yet, or one whose source is not drawing, shows
//! instead of the desktop. **The source's side of the pair is a constant**, so a frame can only
//! change with the host's box: the trace's `cli` (the source's live client) and `pic` (the picture
//! the crop was built from) are the same number for the pin's whole life, and a frame where they
//! disagree says something other than this app resized the source (see `note_the_thumbnail_pair`).
//!
//! **Where the source stands, and what that was measured to be.** It stands at the pin's corner,
//! under the host. `RHP_THUMBNAIL_SOURCE` used to name two other placements and both are dead:
//! `offscreen` (a desktop's width left and a desktop's height below the virtual screen) traced as
//! `source adopted … 1920x1080+-4440+3880` with the registration succeeding and the pin empty, and
//! `edge` (2 px over the screen's bottom-right corner) the same at `… 1920x1080+3438+1438`. **The
//! Shell does not keep a thumbnail live for a window the user cannot see**: the composition surface
//! of a source that is not genuinely on screen is not composed, so its thumbnail is empty. The knob
//! is gone with them — a placement that cannot draw is not an arm of an A/B — and the source is
//! launched where it will stand rather than moved there after it is styled, because mpv shows its
//! window as soon as it opens the file (see `the_source_launch_origin`).
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
//! this road only, so the user can A/B the two runs. It is never asked for on the `--wid` road,
//! where a flip-model swapchain cannot be a child of another window's (see
//! `video_launch::mpv_args`), **and never on a source made invisible**: a layered window has no
//! flip-model swapchain, so `RHP_THUMBNAIL_FLIP` is read, overruled and traced as overruled rather
//! than dropped on the floor (see `video_launch::the_thumbnail_swapchain_flip`).
//!
//! **The source is never hidden, moved or resized.** Hidden the Shell stops drawing it — the
//! thumbnail of a `SW_HIDE` window is frozen at its last frame — so it is taken off the screen only
//! where the host it is seen in is going away for good (see `set_the_source_visible`). Moved it
//! would be a second window writing the same fact as the pin's own drag, and the crop does not care
//! where the source stands (the pair is relative to its client). Resized is the flash this road was
//! measured with. One placement, and then nothing until teardown.
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
//! shortcut: a source never moved is a source that stays where it was placed, so on
//! `RHP_THUMBNAIL_INVISIBLE=0` — the visible road, where the film is covered rather than composited
//! away — a pin carried across the screen leaves its film behind and a pin smaller than its film has
//! the film standing around it. That road is an A/B of the invisibility and nothing else. Upgrade it
//! by clipping the source in a window of this app's, not by moving a window of mpv's again.

use std::sync::atomic::{AtomicIsize, AtomicU32, Ordering};
use std::sync::Mutex;
use std::time::Duration;

use windows::Win32::Foundation::{COLORREF, HWND, POINT, RECT};
use windows::Win32::Graphics::Gdi::{ClientToScreen, InvalidateRect};
use windows::Win32::Graphics::Dwm::{
    DwmRegisterThumbnail, DwmUnregisterThumbnail, DwmUpdateThumbnailProperties,
    DWM_THUMBNAIL_PROPERTIES, DWM_TNP_RECTDESTINATION, DWM_TNP_RECTSOURCE,
    DWM_TNP_SOURCECLIENTAREAONLY, DWM_TNP_VISIBLE,
};
use windows::Win32::UI::WindowsAndMessaging::{
    GetClientRect, GetForegroundWindow, GetSystemMetrics, GetWindowLongPtrW, GetWindowRect,
    GetWindowThreadProcessId, IsWindow, IsWindowVisible, SetForegroundWindow,
    SetLayeredWindowAttributes, SetWindowLongPtrW, SetWindowPos, ShowWindowAsync, GWL_EXSTYLE,
    HWND_BOTTOM, LWA_ALPHA, SM_CXMAXTRACK, SM_CYMAXTRACK, SW_HIDE, SWP_ASYNCWINDOWPOS,
    SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER, SW_SHOWNOACTIVATE, WS_EX_APPWINDOW,
    WS_EX_LAYERED, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_EX_TRANSPARENT,
};

use super::pin_trace::note_pin_trace;
use super::*;

/// **The one A/B this road's size has: `RHP_THUMBNAIL_SOURCE`, absent meaning `native`.**
///
/// `native` places the source, once, at the film's own picture size (see the module note): a pin at or
/// under that size is a downscale the Shell makes, which is sharp, and a pin over it is the upscale
/// the player itself would have had to make. `box` places it, the same once, at the pin's own box
/// instead — the arm that shows what the other one is worth, because a pin grown past the box it was
/// adopted at is drawn from that box.
///
/// **The placement knob this variable used to be is dead, and the generous-and-re-place road it
/// belonged to is not kept behind a word of its own.** A source asked for a new box at each end of a
/// drag is the flash the module note measures, and a run of it is what the trace of this change
/// already is. `offscreen` and `edge`, the two placements measured dead, are gone with it; the source
/// stands under the host, which is the one placement the Shell composes. A value that names neither of
/// the two sizes is `native`.
const SOURCE_SIZE_VARIABLE: &str = "RHP_THUMBNAIL_SOURCE";

/// What `RHP_THUMBNAIL_SOURCE` names (see [`SOURCE_SIZE_VARIABLE`]).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum SourceSize {
    /// The film's own picture size, clamped to what the window system will take.
    Native,
    /// The pin's box at the moment the source was adopted, and nothing else.
    Box,
}

/// The road a `RHP_THUMBNAIL_SOURCE` value names: `box` is the only word that changes anything, and
/// nothing set — or a word this app does not know — is `native`, the road that keeps the picture sharp.
fn source_size_from(value: Option<&std::ffi::OsStr>) -> SourceSize {
    match value.and_then(|value| value.to_str()) {
        Some("box") => SourceSize::Box,
        _ => SourceSize::Native,
    }
}

#[cfg(not(test))]
fn read_the_source_size() -> SourceSize {
    source_size_from(std::env::var_os(SOURCE_SIZE_VARIABLE).as_deref())
}

#[cfg(test)]
fn read_the_source_size() -> SourceSize {
    // The suite is the default road whatever the machine was started with (see `THE_SOURCE_SIZE`).
    SourceSize::Native
}

/// The size this run was started with, read once: a variable set after the app has started changes
/// nothing, so two runs are two comparable arms rather than two halves of one (see
/// [`SOURCE_SIZE_VARIABLE`]).
static THE_SOURCE_SIZE: Lazy<SourceSize> = Lazy::new(read_the_source_size);

/// The size road this run is on (see [`SOURCE_SIZE_VARIABLE`]).
fn the_source_size() -> SourceSize {
    *THE_SOURCE_SIZE
}

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

/// **`RHP_THUMBNAIL_PICTURE`: what a thumbnail of the source shows.** `picture` (the default) crops the
/// source to the rectangle mpv draws the film in — the letterbox margins a film of another shape leaves
/// in its client are not part of the film — and crops that rectangle again to the box's own shape, so a
/// pin's film always fills its box and has no bar to be found in (see `the_thumbnail_box`).
///
/// **`client` is the road this began on, and it is dead.** It stretched the source's whole client,
/// margins included, into the box: measured on the glass as `client 916x515 drawn 3217x1460 picture
/// 2595x1460 margins l311t0r311b0`, where 311 px a side of black pillar was stretched into the pin —
/// the picture small in a large black field that read as *zoomed out* — and the same in reverse as a
/// source smaller than the box, the upscaled blur that read as *zoomed in*. It stays as the other arm of
/// the A/B (see `the_thumbnail_box`).
const PICTURE_VARIABLE: &str = "RHP_THUMBNAIL_PICTURE";

/// What `RHP_THUMBNAIL_PICTURE` names.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum PictureRoad {
    /// The source's picture rectangle, cropped to the box's own shape: the default road, and the one a
    /// pin's film fills its box on.
    Picture,
    /// The source's whole client, stretched into the box: today's road, and the A/B's other arm.
    Client,
}

impl PictureRoad {
    /// The word the trace writes for this road: the value the variable takes.
    fn named(self) -> &'static str {
        match self {
            PictureRoad::Picture => "picture",
            PictureRoad::Client => "client",
        }
    }
}

/// The road a `RHP_THUMBNAIL_PICTURE` value names: `client` is the only word that changes anything, and
/// nothing set — or a word this app does not know — is `picture`, the road that draws the film itself.
fn picture_road_from(value: Option<&std::ffi::OsStr>) -> PictureRoad {
    match value.and_then(|value| value.to_str()) {
        Some("client") => PictureRoad::Client,
        _ => PictureRoad::Picture,
    }
}

#[cfg(not(test))]
fn read_the_picture_road() -> PictureRoad {
    picture_road_from(std::env::var_os(PICTURE_VARIABLE).as_deref())
}

#[cfg(test)]
fn read_the_picture_road() -> PictureRoad {
    // The suite is the default road whatever the machine was started with (see `THE_PICTURE_ROAD`).
    PictureRoad::Picture
}

/// The road this run was started with, read once: a variable set after the app has started changes
/// nothing, so two runs are two comparable arms rather than two halves of one (see [`PICTURE_VARIABLE`]).
static THE_PICTURE_ROAD: Lazy<PictureRoad> = Lazy::new(read_the_picture_road);

/// The picture road this run is on (see [`PICTURE_VARIABLE`]).
fn the_picture_road() -> PictureRoad {
    *THE_PICTURE_ROAD
}

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

/// How many times adoption asks the player for the film's own picture size before it takes the
/// probe's measurement of the same film instead: a handful of bounded reads a few milliseconds apart,
/// which is what a player that has made its window and not yet opened the file needs to answer (see
/// `the_films_native_size`).
const NATIVE_SIZE_ASKS: u8 = 4;
/// How long adoption waits between two of those reads. **This is the only wait this road makes and it
/// is not on a drag's road**: it happens once, in the adoption pass, and it is bounded by
/// `NATIVE_SIZE_ASKS` reads that are each bounded themselves (see `adopt_the_source`).
const NATIVE_SIZE_WAIT: Duration = Duration::from_millis(3);

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
/// taskbar's thickness away from the pin's corner. The offset is taken off the corner and nothing
/// else.
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

/// A box as a trace and a placement carry it: origin, width and height.
type SourceBox = (i32, i32, i32, i32);

/// What a pin's source shows, fixed at the one placement: the client the ask named, and the film's own
/// shape fitted inside it (see `the_one_placement` and `the_thumbnail_box`).
#[derive(Clone, Copy, PartialEq, Debug)]
struct SourceShape {
    /// The source's client, as the placement asked for it. **Never read again**: nothing asks the
    /// window for another box, so the crop's window is a constant for the pin's whole life.
    client: (i32, i32),
    /// The film's own shape, width over height, from the picture's own size — or nothing where it
    /// could not be read, which is the whole client into the whole box.
    aspect: Option<f64>,
}

#[derive(Clone, Copy)]
struct Thumbnail {
    /// The host the Shell is drawing into, and the registration it answered with.
    host: isize,
    registration: isize,
    /// The rectangle pair the registration was last handed — the host's client it is drawn in, the
    /// picture's rectangle in the source and the source's own client — or nothing where none stands.
    /// **A change of either is a re-point** (see `settle_the_registration`).
    shown: Option<ThumbnailBox>,
    /// The picture the source shows: `None` until the one placement is made, and the same value for
    /// every frame after it (see `place_the_source`).
    source: Option<SourceShape>,
}

impl Thumbnail {
    const NOTHING: Self = Self {
        host: 0,
        registration: 0,
        shown: None,
        source: None,
    };
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
/// road that deadlocked this app once already (see `fit_the_hosted_mpv_to_the_host`). What a post
/// costs is the window standing on its own for the moment it takes mpv to pump it, which is behind
/// the hover the pin is being taken up from.
///
/// Safety: `source` is a window the caller found by `pid` and this app did not make.
pub(super) unsafe fn adopt_the_source(source: HWND, pid: u32) {
    let ex = unsafe { GetWindowLongPtrW(source, GWL_EXSTYLE) };
    let invisible = the_source_is_invisible();
    let _ = unsafe { SetWindowLongPtrW(source, GWL_EXSTYLE, source_ex_style(ex, invisible)) };
    if invisible {
        // Safety: as above, and the alpha is this app's own to give a window of another process's. The
        // one step from nothing, `LWA_ALPHA`, is what keeps the window drawn (and therefore
        // thumbnailable) while nothing of it reaches the screen.
        let applied = unsafe {
            SetLayeredWindowAttributes(source, COLORREF(0), INVISIBLE_ALPHA, LWA_ALPHA)
        };
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
    // source is launched at the pin's corner (see `the_source_launch_origin`) and this is the line
    // that says whether mpv honoured it; the one placement's own line follows it with what this app
    // asked for instead.
    let mut created = RECT::default();
    if unsafe { GetWindowRect(source, &mut created) }.is_ok() {
        note_pin_trace!(
            "source adopted {:?} pid {} placed by mpv {}x{}+{}+{}",
            source,
            pid,
            created.right - created.left,
            created.bottom - created.top,
            created.left,
            created.top,
        );
    }

    if let Some(host) = pin_film_host::film_host() {
        // **The one placement the source ever takes is made here**, in the adoption pass and before
        // the registration draws anything: the film's own picture size at the host's corner, asked
        // once and never asked for again (see `place_the_source`).
        //
        // **A film whose own size the player has not answered is waited for rather than guessed at**,
        // and where even the probe has nothing the source is left where mpv put it — which is still a
        // box this app never asks again, and the crop is of the picture that is there (see
        // `the_films_native_size`). The wait happens before the lock is taken, so a read and a sleep
        // are never spent holding what the pin's own resize steps want.
        let placed = THUMBNAIL.lock().ok().and_then(|held| held.source);
        let placed = place_the_source(source, host, pid, placed);
        if let Ok(mut held) = THUMBNAIL.lock() {
            if let Some(placed) = placed {
                held.source = Some(placed);
            }
            settle_the_registration(&mut held, host);
        }
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

/// The window the Shell is drawing the pin's film from, where there is one: the window of mpv's the
/// pin's picture is a thumbnail of, which is the third window a press can land in on this road (see
/// `the_band_was_placed` and `key_input::pin_owns_caret`).
pub(super) fn the_source_window() -> Option<HWND> {
    let number = SOURCE.load(Ordering::SeqCst);
    (number != 0).then_some(HWND(number as *mut _))
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
    if SOURCE.load(Ordering::SeqCst) == 0 {
        return;
    }
    let Ok(mut held) = THUMBNAIL.lock() else {
        return;
    };
    settle_the_registration(&mut held, host);
}

/// **The one placement a pin's source ever takes**: the film's own picture size at the pin's corner,
/// asked for once, at adoption, and never asked for again until the window is taken down.
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
fn place_the_source(
    source: HWND,
    host: HWND,
    pid: u32,
    placed: Option<SourceShape>,
) -> Option<SourceShape> {
    let origin = (unsafe { client_origin(host) })?;
    let client = (unsafe { client_size(host) })?;
    let native = the_films_native_size(pid);
    let Some((x, y, width, height)) = the_one_placement(
        placed,
        Some(origin),
        native,
        the_max_track(),
        the_source_size(),
        client,
    ) else {
        if placed.is_some() {
            note_pin_trace!(
                "source re-placed {:?} host {:?}: a source takes one placement and this is a second ask",
                source,
                host,
            );
        } else {
            note_pin_trace!(
                "thumbnail source native unknown {:?}: no box asked, and none will be",
                source,
            );
        }
        return None;
    };
    let asked = unsafe {
        SetWindowPos(
            source,
            HWND::default(),
            x,
            y,
            width,
            height,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS,
        )
    };
    let aspect = native.map(|native| native.0 as f64 / native.1.max(1) as f64);
    let (left, top, picture_width, picture_height) = a_picture_fitted_to(aspect, (width, height));
    note_pin_trace!(
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
        asked.is_ok() as u8,
    );
    Some(SourceShape {
        client: (width, height),
        aspect,
    })
}

/// The one box a pin's source is ever asked for — the film's own picture size at the pin's corner,
/// clamped to what the window system will take — and nothing at all for a source that has been placed
/// already, for a host that is not there, or for a film whose own size is not known yet.
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
    size: SourceSize,
    box_: (i32, i32),
) -> Option<SourceBox> {
    if placed.is_some() {
        return None;
    }
    let (x, y) = host_origin?;
    let wanted = match size {
        SourceSize::Native => native?,
        SourceSize::Box => box_,
    };
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
/// adoption pass and nowhere near a drag (see `adopt_the_source`), and what it buys is the one
/// placement this road makes.
fn the_films_native_size(pid: u32) -> Option<(i32, i32)> {
    for ask in 0..NATIVE_SIZE_ASKS {
        if let Some(size) = the_players_picture_size(pid) {
            return Some(size);
        }
        if ask + 1 < NATIVE_SIZE_ASKS {
            std::thread::sleep(NATIVE_SIZE_WAIT);
        }
    }
    the_films_own_size()
}

/// The film's own picture size as the probe measured it when the film was loaded: read out of the
/// geometry cache and never asked of the player, which is the fallback for a player that has not
/// opened the file by the time its window is adopted (see `the_films_native_size` and
/// `dimensions::video_box`).
fn the_films_own_size() -> Option<(i32, i32)> {
    let (path, _) = pinned_media_owner()?;
    let (width, height) = dimensions::video_box(&path)?;
    Some((width as i32, height as i32))
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
    Some(if turned { (height, width) } else { (width, height) })
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
pub(super) fn set_the_source_visible(visible: bool) {
    let Some(source) = the_source_window() else {
        return;
    };
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
/// is showing a film of its own, which is every road that is not this one.
pub(super) fn forget_the_thumbnail_of(host: HWND) {
    let drawn_in = REGISTERED_HOST.load(Ordering::SeqCst) == host.0 as isize;
    if drawn_in {
        forget_the_thumbnail();
    }
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
                let width = (box_height as f64 * aspect).round().clamp(1.0, box_width as f64);
                (width as i32, box_height)
            } else {
                let height = (box_width as f64 / aspect).round().clamp(1.0, box_height as f64);
                (box_width, height as i32)
            }
        }
        _ => (box_width, box_height),
    };
    ((box_width - width) / 2, (box_height - height) / 2, width, height)
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
/// **The `client` road, and a picture nothing has read yet, is the whole client into the whole box** —
/// today's stretch, and the fallback for a shape this app cannot crop by.
fn the_thumbnail_box(
    box_: (i32, i32),
    client: (i32, i32),
    aspect: Option<f64>,
    road: PictureRoad,
) -> ThumbnailBox {
    let aspect = (road == PictureRoad::Picture).then_some(aspect).flatten();
    let client = (client.0.max(1), client.1.max(1));
    let box_ = (box_.0.max(1), box_.1.max(1));
    let (source, clamped) = match aspect {
        Some(_) => {
            let ((left, top, width, height), clamped) =
                the_crop_filling(a_picture_fitted_to(aspect, client), box_);
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
        let _ = unsafe { DwmUnregisterThumbnail(held.registration) };
        held.registration = 0;
        held.shown = None;
    }
    held.host = host.0 as isize;
    REGISTERED_HOST.store(held.host, Ordering::SeqCst);

    let source = HWND(SOURCE.load(Ordering::SeqCst) as *mut _);
    let Some(client) = (unsafe { client_size(host) }) else {
        return;
    };
    // **The one read of the source at this call.** `cli` in the trace is this number and `pic` the one
    // the pair was cropped against, so a frame where the two disagree says which window a flash would
    // have to come from (see `note_the_thumbnail_pair`).
    let live = the_source_client(source, client);
    let (source_client, aspect) = held
        .source
        .map_or((live, None), |source| (source.client, source.aspect));
    let shown = the_thumbnail_box(client, source_client, aspect, the_picture_road());
    if !the_registration_wants(held, shown) {
        note_the_thumbnail_pair("held", host, client, live, shown);
        return;
    }

    if held.registration == 0 {
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
    // is the pillar of black that read as a pin zoomed out (see `PICTURE_VARIABLE`).
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
        let _ = unsafe { DwmUnregisterThumbnail(held.registration) };
        // The picture adoption placed is kept: the next registration is cropped by the same one, and a
        // failure of the Shell's is not a reason to fall back to the whole client.
        let source = held.source;
        *held = Thumbnail::NOTHING;
        held.source = source;
        REGISTERED_HOST.store(0, Ordering::SeqCst);
        return;
    }
    held.shown = Some(shown);
    // **The host is asked to erase its whole client behind the picture**, which is the other half of
    // the invariant above: the Shell composes the thumbnail over this app's content, so a client the
    // host never erased is a pixel nobody drew — and one of those is the desktop seen through the pin.
    let _ = unsafe { InvalidateRect(host, None, true) };
    note_the_thumbnail_pair("update", host, client, live, shown);
}

/// The one line that makes a hole visible at a glance: the host the Shell is drawing into (its own box
/// and where it stands), the host client the pair was built from, the source's live client (`cli`), the
/// client the crop was built from (`pic`) and the exact rectangles handed over. **`host` and `box`
/// disagree only where something stands on the host's client** — a border, a frame style — which is a
/// picture that cannot line up with its pin's box, and **`cli` and `pic` disagree only where the source
/// was resized by something that is not this app**, which is the one way a crop can be of a surface the
/// Shell is not sampling.
fn note_the_thumbnail_pair(
    what: &str,
    host: HWND,
    client: (i32, i32),
    live: (i32, i32),
    shown: ThumbnailBox,
) {
    let (host_x, host_y, host_width, host_height) =
        unsafe { window_box(host) }.unwrap_or((0, 0, client.0, client.1));
    let (x, y, width, height) = shown.destination;
    let (left, top, right, bottom) = match shown.source {
        Some((left, top, right, bottom)) => (left, top, right, bottom),
        None => (0, 0, shown.client.0, shown.client.1),
    };
    note_pin_trace!(
        "thumbnail {} host {}x{}+{}+{} box {}x{} dest {}x{}+{}+{} cli {}x{} pic {}x{} src {}x{}+{}+{} ({}={})",
        what,
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
        PICTURE_VARIABLE,
        the_picture_road().named(),
    );
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
    use super::*;

    /// **A step re-points the registration only where the host's box changed**: the source's client and
    /// the film's shape are constants of the pin's life, fixed by the one placement, so a pair is a
    /// function of the box alone (see `settle_the_registration` and `the_thumbnail_box`).
    #[test]
    fn a_step_re_points_a_registration_only_where_the_box_changed() {
        let client = (1000, 560);
        let aspect = Some(2595.0 / 1460.0);
        let source = SourceShape {
            client: (2731, 1536),
            aspect,
        };
        let shown_at = |box_: (i32, i32)| {
            the_thumbnail_box(box_, source.client, source.aspect, PictureRoad::Picture)
        };
        let shown = shown_at(client);
        let mut holding = Thumbnail {
            host: 0x2000,
            registration: 9,
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
                let ((x, y, width, height), _) =
                    the_crop_filling(picture, (box_width, box_height));
                assert!(
                    x >= picture.0
                        && y >= picture.1
                        && width >= 1
                        && height >= 1
                        && x + width <= picture.0 + picture.2
                        && y + height <= picture.1 + picture.3,
                    "{width}x{height}+{x}+{y} is inside the picture {picture:?} at {shape:?}"
                );
                let shown = the_thumbnail_box(
                    (box_width, box_height),
                    (box_width, box_height),
                    shape,
                    PictureRoad::Picture,
                );
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
                    let shown = the_thumbnail_box(box_, client, shape, PictureRoad::Picture);
                    assert_eq!(
                        shown.destination,
                        (0, 0, box_.0, box_.1),
                        "the destination is the box, at {client:?} {shape:?} {box_:?}"
                    );
                    let Some((left, top, right, bottom)) = shown.source else {
                        assert!(shape.is_none(), "a shape nothing has read is the whole client");
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

    /// The registration is nobody's where none was taken: a thumbnail record with no pair stands in no
    /// host, which is what a host that has not been pointed at yet reads as.
    #[test]
    fn nothing_stands_in_a_host_until_a_pair_is_handed_over() {
        assert_eq!(Thumbnail::NOTHING.registration, 0);
        assert_eq!(Thumbnail::NOTHING.host, 0);
        assert_eq!(Thumbnail::NOTHING.shown, None);
    }

    /// **The size knob names the film's own size by default, and `box` for the one A/B**: a value this
    /// app does not know — including the placement words the knob used to carry — is `native` (see
    /// `source_size_from`).
    #[test]
    fn the_source_size_knob_is_native_unless_it_says_box() {
        assert_eq!(
            SOURCE_SIZE_VARIABLE, "RHP_THUMBNAIL_SOURCE",
            "the knob is the name the branch's notes and the report give it"
        );
        let named = |value: Option<&str>| source_size_from(value.map(std::ffi::OsStr::new));
        assert_eq!(named(None), SourceSize::Native);
        assert_eq!(named(Some("native")), SourceSize::Native);
        assert_eq!(named(Some("box")), SourceSize::Box);
        assert_eq!(
            named(Some("Box")),
            SourceSize::Native,
            "the other arm is not guessed at from another case"
        );
        assert_eq!(named(Some("")), SourceSize::Native);
        assert_eq!(
            named(Some("behind")),
            SourceSize::Native,
            "the placement words this knob used to carry are not roads it knows now"
        );
        assert_eq!(
            the_source_size(),
            SourceSize::Native,
            "and the suite is the sharp one, whatever the machine was started with"
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
        assert!(named(None), "this road's default is the source nothing of which is seen");
        assert!(named(Some("1")));
        assert!(!named(Some("0")), "and `0` is the visible road, byte for byte");
        assert!(named(Some("")), "a value that names nothing is not read as off");
        assert!(named(Some("off")), "only the one value turns it off, and it is `0`");

        let inert = WS_EX_APPWINDOW.0 as isize | WS_EX_TOPMOST.0 as isize;
        let invisible = source_ex_style(inert, true);
        assert_eq!(invisible & (WS_EX_LAYERED.0 as isize), WS_EX_LAYERED.0 as isize);
        for bit in [WS_EX_TOOLWINDOW, WS_EX_NOACTIVATE, WS_EX_TRANSPARENT] {
            assert_eq!(invisible & (bit.0 as isize), bit.0 as isize, "{bit:?} is kept");
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
        let _one = super::super::pin_window::PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|held| held.into_inner());
        let pid = 0x5eed;
        mpv_ipc::remember_pipe(pid, "rhp-test-picture-size", false);
        mpv_ipc::clear_reads();

        // An anamorphically coded film: the display size is what a source is placed at, because that is
        // the shape the player would draw it in.
        mpv_ipc::answer_a_read(
            pid,
            "video-params",
            Some(r#"{"pixelformat":"yuv420p","w":1440,"h":1080,"dw":1920,"dh":1080,"aspect":1.777778}"#),
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
        let _one = super::super::pin_window::PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|held| held.into_inner());
        let pid = 0x5eed;
        mpv_ipc::remember_pipe(pid, "rhp-test-unsized-film", false);
        mpv_ipc::clear_reads();
        mpv_ipc::answer_a_read(pid, "video-params", None);

        assert_eq!(
            the_films_native_size(pid),
            the_films_own_size(),
            "a player that never answers is the probe's own measurement, not a box this app guessed"
        );
        assert_eq!(
            mpv_ipc::noted_reads().len(),
            NATIVE_SIZE_ASKS as usize,
            "and the wait is bounded: one read per ask and no more"
        );
        assert_eq!(
            the_one_placement(
                None,
                Some((100, 200)),
                None,
                (3440, 1440),
                SourceSize::Native,
                (640, 360)
            ),
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
            the_one_placement(
                None,
                Some((100, 200)),
                native,
                max_track,
                SourceSize::Native,
                (640, 360)
            ),
            Some((100, 200, 1920, 1080)),
            "the film's own size, at the pin's corner"
        );
        assert_eq!(
            the_one_placement(
                None,
                Some((100, 200)),
                native,
                max_track,
                SourceSize::Box,
                (640, 360)
            ),
            Some((100, 200, 640, 360)),
            "and the A/B's other arm is the pin's box, once, at the same corner"
        );

        let placed = Some(SourceShape {
            client: (1920, 1080),
            aspect: Some(16.0 / 9.0),
        });
        for box_ in [(640, 360), (1920, 1080), (3840, 2160)] {
            for native in [None, Some((1920, 1080)), Some((3840, 2160))] {
                for size in [SourceSize::Native, SourceSize::Box] {
                    assert_eq!(
                        the_one_placement(placed, Some((100, 200)), native, max_track, size, box_),
                        None,
                        "a placement already taken is not taken again, whatever the box is"
                    );
                }
            }
        }
        assert_eq!(
            the_one_placement(None, None, native, max_track, SourceSize::Native, (640, 360)),
            None,
            "no host is no corner to place at"
        );
        assert_eq!(
            the_one_placement(
                None,
                Some((0, 0)),
                None,
                max_track,
                SourceSize::Native,
                (10_000, 10_000)
            ),
            None,
            "a native size that is not known yet is not the box: it is waited for"
        );
        assert_eq!(
            the_one_placement(
                None,
                Some((0, 0)),
                Some((8192, 4608)),
                max_track,
                SourceSize::Native,
                (640, 360)
            ),
            Some((0, 0, 3440, 1440)),
            "a film bigger than the desktop is clamped to what a window may be"
        );
        assert_eq!(
            the_one_placement(
                None,
                Some((0, 0)),
                Some((0, 0)),
                max_track,
                SourceSize::Native,
                (640, 360)
            ),
            Some((0, 0, 1, 1)),
            "and never to nothing at all"
        );
    }

    /// **The picture road is the default and the only other word is `client`**: a value this app does
    /// not know is not a second way to ask for a stretch of the whole client (see `picture_road_from`).
    #[test]
    fn the_picture_knob_names_the_picture_road_and_the_client_one() {
        assert_eq!(
            PICTURE_VARIABLE, "RHP_THUMBNAIL_PICTURE",
            "the knob is the name the report and the run's own trace give it"
        );
        let named = |value: Option<&str>| picture_road_from(value.map(std::ffi::OsStr::new));
        assert_eq!(named(None), PictureRoad::Picture);
        assert_eq!(named(Some("picture")), PictureRoad::Picture);
        assert_eq!(named(Some("client")), PictureRoad::Client);
        assert_eq!(
            named(Some("Client")),
            PictureRoad::Picture,
            "the dead road is not guessed at from another case"
        );
        assert_eq!(named(Some("")), PictureRoad::Picture);
        assert_eq!(named(Some("stretch")), PictureRoad::Picture);
        assert_eq!(PictureRoad::Picture.named(), "picture");
        assert_eq!(PictureRoad::Client.named(), "client");
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
        assert_eq!(the_crop_filling((0, 0, 1440, 1080), (1, 1)).0, (180, 0, 1080, 1080));
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
            assert!(upscaled, "and the trace says the crop was cut to the picture");
        }
    }

    /// **A shape that rounds is a picture a pixel out and never a pixel over its client**: the measured
    /// 2595x1460 picture in a 1000x560 client is 995 px wide in it, and a client of one pixel is the
    /// client itself (see `a_picture_fitted_to`).
    #[test]
    fn a_fitted_picture_rounds_inside_its_client_and_never_out_of_it() {
        let measured = Some(2595.0 / 1460.0);
        let (x, y, width, height) = a_picture_fitted_to(measured, (1000, 560));
        assert_eq!((x, y, width, height), (2, 0, 995, 560));
        assert!(x + width <= 1000 && y + height <= 560, "never a pixel over the client");
        assert!(x >= 0 && y >= 0, "and never one outside it");

        assert_eq!(a_picture_fitted_to(Some(1.7773973), (1, 1)), (0, 0, 1, 1));
        assert_eq!(a_picture_fitted_to(Some(16.0 / 9.0), (1, 1)), (0, 0, 1, 1));
        assert_eq!(
            a_picture_fitted_to(Some(16.0 / 9.0), (0, 0)),
            (0, 0, 1, 1),
            "a client of nothing is asked for what a window can be"
        );
        assert_eq!(a_picture_fitted_to(None, (800, 600)), (0, 0, 800, 600));
        assert_eq!(a_picture_fitted_to(Some(f64::NAN), (800, 600)), (0, 0, 800, 600));
    }

    /// **The picture road crops the source to the part of its picture that fills the box; the client road
    /// and a picture nothing has read are the whole client into the whole box** — the two arms of the
    /// A/B, and the fallback that is byte for byte today's road (see `the_thumbnail_box`).
    #[test]
    fn a_thumbnail_shows_the_pictures_rectangle_or_the_whole_client() {
        // The measured source: a 16:9 film in a 3217x1460 client, 311 px of pillar either side.
        // The measured source: a 16:9 film in a 3217x1460 client, 311 px of pillar either side.
        let aspect = Some(2595.0 / 1460.0);
        let shown = the_thumbnail_box((1000, 560), (3217, 1460), aspect, PictureRoad::Picture);
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

        let whole = the_thumbnail_box((1000, 560), (3217, 1460), aspect, PictureRoad::Client);
        assert_eq!(whole.source, None, "the dead road hands the Shell the whole client");
        assert_eq!(whole.destination, (0, 0, 1000, 560));

        let nothing_read = the_thumbnail_box((1000, 560), (3217, 1460), None, PictureRoad::Picture);
        assert_eq!(nothing_read.source, None, "a picture nothing has read is not cropped by");
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
                the_one_placement(
                    placed,
                    Some((100, 200)),
                    Some((1920, 1080)),
                    (3440, 1440),
                    SourceSize::Native,
                    box_
                ),
                None,
                "{box_:?} is a moment of the gesture, and the gesture asks for no box"
            );
            let shown =
                the_thumbnail_box(box_, source.client, source.aspect, PictureRoad::Picture);
            assert_eq!(
                shown.destination,
                (0, 0, box_.0, box_.1),
                "the box is filled, edge to edge, at {box_:?}"
            );
            assert_eq!(
                shown.client, source.client,
                "the crop's window is the one placement's, never a box the pin was dragged to"
            );
            let (left, top, right, bottom) = shown.source.expect("the picture is what is cropped by");
            assert!(
                left >= picture.0
                    && top >= picture.1
                    && right <= picture.0 + picture.2
                    && bottom <= picture.1 + picture.3,
                "{box_:?} crops inside the film's own picture"
            );
            assert_eq!(
                shown,
                the_thumbnail_box(box_, source.client, source.aspect, PictureRoad::Picture),
                "and the same box is always the same pair: nothing else is an input"
            );
        }

        let over = the_thumbnail_box((3840, 2160), source.client, source.aspect, PictureRoad::Picture);
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

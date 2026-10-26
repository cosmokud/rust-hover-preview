//! The pin's film drawn by the Desktop Window Manager rather than nested in a window of this app's.
//!
//! **This is the road a pin's mpv film is drawn on**: `RHP_MPV_THUMBNAIL_PIN=0` is the escape hatch
//! back to the child road, and it is byte for byte the road this app stood in before (see
//! `pin_film_host::MPV_THUMBNAIL_PIN_VARIABLE`). mpv is launched with no `--wid` and makes a window
//! of the desktop's of its own; the pin's film
//! host asks the Shell to draw that window into itself (`DwmRegisterThumbnail`), and what follows a
//! resize is one rectangle rather than two windows agreeing about a box.
//!
//! **Why the film has to be a window of the desktop's.** A thumbnail is taken of a window's
//! composition surface, and a `WS_CHILD` has none of its own — it is drawn into its parent's — so a
//! film that is to be thumbnailed cannot be mpv's child under the host. That is the whole of why
//! this road launches mpv without `--wid` (see `video_launch::mpv_args`).
//!
//! **A resize step asks little of mpv.** The Shell stretches the source's picture into the destination
//! rectangle whatever the two sizes are, so a step is this app moving its own two windows and
//! re-pointing the destination at the host's new client area (see `the_band_was_placed`). The source is
//! resized only as a coarse catch-up while the hand moves and exactly at each end of the drag (see
//! `follow_the_band_while_dragging` and `place_the_source`), and every ask is `SWP_ASYNCWINDOWPOS` like
//! every other call made on a window of another process's here.
//!
//! **What a thumbnail shows is the source's picture, not its whole client** (`RHP_THUMBNAIL_PICTURE`,
//! `picture` by default): the rectangle mpv draws the film in — the client less the letterbox margins a
//! film of another shape leaves — is cropped out of the source to **the box's own shape**, and the whole
//! box is handed over as the destination, so the film fills its pin (see `the_thumbnail_box`). `client`
//! is the road this began on and it is dead: it stretched the whole client, margins included, into the
//! box, which is the pillar of black that read as a pin zoomed out (measured: `client 916x515 … picture
//! 2595x1460 margins l311t0r311b0` — 311 px a side stretched into the picture) and the upscale that read
//! as one zoomed in. The shape is read at a settle and kept, so a step of a drag still reads nothing of
//! the player's.
//!
//! **Filling the box is not a detail of the crop, it is what makes a flash impossible.** The road this
//! replaces fitted the picture into the box at the picture's own shape and left the rest to the host's
//! black: a destination *smaller* than the client whenever the box's aspect and the film's were not the
//! same. The drift between them is what the hand sees as the film squashed into the middle of the pin
//! with a pillar either side, twice per drag — at the press and at the release — because those two
//! moments are where the picture's shape is read again (see `place_the_source`), while every step between
//! them is re-pointed from the shape that read gave and a box whose aspect a resize holds still (see
//! `pin_geometry::resize_pinned_content`). No pair of shapes can produce a bar now, whatever order the
//! box, the source and the shape are re-pointed in.
//!
//! **The composition invariant is one line**: the destination is the host's whole client, so the Shell
//! draws every pixel of the pin, and the host's own erase paints that same client under it — which is
//! what a client the Shell has not composed yet, or one whose source is not drawing, shows instead of the
//! desktop. Both rectangles are built at the moment of the `DwmUpdateThumbnailProperties` call from the
//! host's live client and the window the picture is *in*: the source's live client, except where the
//! settle that is pointing the pair knows the content is somewhere else still — the box it has just asked
//! away from, or the window the player's own answer named (see `the_window_the_picture_is_at`). **A
//! change of any of them is a re-point**, and a settle re-points even where the box did not change,
//! because the settle's own ask is what changes the source's client. A pair built against a box the host
//! is no longer in — or a source rectangle reaching past the surface the Shell is sampling, which the
//! Shell answers with nothing — is the picture cropped and a band of desktop showing through the pin.
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
//! **A placement is made exactly at each end of a drag**, and between them only the coarse catch-up
//! the hand's own resizing is owed — on the visible placement (see above, and
//! `follow_the_band_while_dragging`). The source takes the box the pin began the drag at and is asked
//! for the box the drag settled at when the hand lets go. What the two off-screen placements were
//! wanted for is the shrink a `behind` source cannot follow: the destination shrinks, the source does
//! not, and there is nothing of a source that is not on the screen for the host to leave behind —
//! which is the one thing they are no good for, since a source that is not on the screen is not drawn
//! at all.
//!
//! **Only `behind` is moved with the host** (see `the_band_moves_the_source`), because it is the only
//! placement that has to stay under the pin to stay unseen; the two off-screen placements have nothing
//! to follow, and not asking is the point.
//!
//! **The invisible road's drag asks the source nothing between its two ends.** One ask while the Shell
//! is thumbnailing a window is a window mid-reconfigure — part of its surface holds the previous scale
//! and part the next, and the pin shows both at once, which is the flicker measured on the glass. So on
//! that road a resize asks once at its beginning for a box nothing of which can be seen and which the
//! whole gesture can be drawn down from (see `the_box_a_placement_asks_for`), then asks nothing until
//! the settle at its end, and neither a corner nor a size follows the hand in between (see
//! `the_band_was_placed`). `0 asks while dragging` in the settle line is the proof, and
//! `RHP_THUMBNAIL_INVISIBLE=0` is the run it is compared with.
//!
//! **The picture's fill is read at a settle, never during a drag**: the read is a bounded pipe
//! request, and the middle of a resize is the one moment this road must not spend ten milliseconds
//! waiting on another process (see `settle_the_fill`).
//!
//! **A settle is verified, not assumed.** One ask is not a proof that mpv kept the box: the player
//! makes its window, configures its output for the frame, and can drop or overwrite a placement made
//! under it (measured: a source adopted at a pin's band still read back at mpv's own default box).
//! The Shell stretches whatever box the source is really in into the destination, so a source left at
//! another box is a picture stretched out of its shape — which is what a pin that is *zoomed in* or
//! *small in a field of black* after a resize drag is looking at, and it stands until the next box
//! settles. So the tick compares the source's own box with the pin's and asks again where they differ:
//! throttled by `SETTLE_RETRY`, capped at `SETTLE_ASKS` per settle, traced per look, and with the
//! player's `vo-configured` read once at adoption rather than per tick (see `ask_the_source_for`,
//! `keep_the_source_at_the_band` and `the_players_output_is_configured`).
//!
//! **But a box that matches is not yet a picture that fills it.** Measured on the glass: every settle
//! succeeded on its first ask, the source's window box always equalled the host's, and the pin still
//! came out zoomed in (cropped) or small in a field of black. The mismatch is inside mpv — the output
//! it is presenting is not the size of the window it is presenting into. So a settle also **reads the
//! player's own output over the existing pipe, once per settle**, and treats the settle as taken only
//! when the box matches *and* the output fills it (see `the_picture_fills`). Where the box matched and
//! the output did not, the source is **nudged**: the same box is asked for one pixel larger and then
//! exactly, because a repeated identical size produces no `WM_SIZE` and a VO stuck at an old buffer has
//! nothing else to make it reconfigure. One nudge per settle, capped, and a source that still
//! disagrees says `source never filled …` once and is left alone — and not asked again for that box at
//! all (see `settle_the_fill`, `nudge_the_source` and `place_the_source`).
//!
//! **The numbers are `osd-dimensions`, and the first attempt at them was wrong in a way worth writing
//! down.** This settle used to read `dwidth`/`dheight` and call them the drawn picture; they are the
//! film's *display* size — what it would be drawn at were it not fitted to a window — so they read
//! `1920x1080` for every box this road ever asked for, the comparison could never pass, and the nudge
//! fired on every settle in the run. `osd-dimensions` is the player's own answer about its output:
//! `w`/`h` are the window the VO configured for, and `mt`/`mb`/`ml`/`mr` are the margins the film's own
//! shape leaves, so the picture drawn inside them is `w` less the two horizontal margins by `h` less
//! the two vertical ones. Measured on mpv 0.41.0, a 1920×1080 film in a 640×300 window:
//! `{"w":640,"h":300,"par":1.000000,"aspect":2.133333,"mt":0,"mb":0,"ml":53,"mr":54}` — 53 + 533 +
//! 54 = 640, which is the letterbox of a film the box does not share the shape of. **The invariant is
//! therefore the picture plus its margins covering the client area, and that same configured window
//! being the client area**: a film wider than the box passes because its margins are the black it is
//! honestly shown with, and only an output configured for another size fails (see `Drawn` and
//! `the_drawn_picture`).
//!
//! **The structural version of that invariant is a flip-model swapchain**, which the compositor
//! scales to the window's client area by construction instead of relying on mpv catching a
//! `WM_SIZE`: `RHP_THUMBNAIL_FLIP=1` passes `--d3d11-flip=yes` to a pinned mpv on this road and
//! this road only, so the user can A/B the two runs. It is never asked for on the `--wid` road,
//! where a flip-model swapchain cannot be a child of another window's (see
//! `video_launch::mpv_args`), **and never on a source made invisible**: a layered window has no
//! flip-model swapchain, so `RHP_THUMBNAIL_FLIP` is read, overruled and traced as overruled rather than
//! dropped on the floor (see `video_launch::the_thumbnail_swapchain_flip`).
//!
//! **A visible source follows the box while a drag resizes it**, because a frozen source grown from
//! 680 px to 2263 px is a 3.3× upscale no swapchain model can hide. It follows asymmetrically: a
//! *growing* box is chased only once it is a tenth larger than the source (a coarse catch-up, a handful
//! of asks over a long drag, soft in between), while a *shrinking* one is asked for ahead of itself by a
//! small margin so a step of lag cannot leave the source spilling out from under the host (see
//! `follow_the_band_while_dragging`). **The invisible road buys the same picture the other way round** —
//! one generous box at the drag's beginning instead of a handful of catch-ups — for the reason the next
//! paragraph gives.
//!
//! **The source is never hidden.** The Shell draws nothing for a source it considers hidden — the
//! thumbnail of a `SW_HIDE` window is frozen at its last frame — so the source is *left on the
//! screen* at a placement where nothing of it shows, and it is taken off only where the host it is
//! seen in is going away for good (see `set_the_source_visible`). That is also why an `offscreen` or
//! `edge` source is **launched** where it will live rather than moved there after it is styled: mpv
//! shows its window when it opens the file, which is before this app has found it (see
//! `the_source_launch_origin`).
//!
//! **On this road it is also made invisible rather than merely covered** (`RHP_THUMBNAIL_INVISIBLE`,
//! on by default and off with `=0`): the source wears `WS_EX_LAYERED` with
//! `SetLayeredWindowAttributes(…, 1, LWA_ALPHA)`, so the compositor still draws it — a source the
//! Shell is not drawing has no picture to thumbnail at all (measured, see the three placements above)
//! — while nothing of it reaches the screen. That is the same one-step-from-nothing band this app's
//! own chrome used to wear, and it is what lets the source be *larger than the box* without spilling
//! (see `the_generous_source_size`): a window nobody can see cannot stand outside a pin, so a resize
//! does not have to chase it, and a source that is not chased is a source that is not mid-reconfigure
//! while the Shell is thumbnailing it. It is put on in the adoption pass and re-asserted by the style
//! monitor if mpv's window comes back without it (see `keep_the_source_invisible`).
//!
//! **The bet this road is:** the Shell's thumbnail of a layered window is the window's own content at
//! full opacity, so the film still draws in the pin with the source itself invisible. It has two ways to
//! lose, and the run says which: a pin that goes black while the player is still alive — still answering
//! `source fills` with the source's own size — is a Shell sampling the blended (1/255) result instead of
//! the window's content, while a pin that goes black and takes the player with it is a present that
//! cannot target a layered window at all. That is also why nothing here hides the window to achieve the
//! invisibility: the two failures are told apart by whether the player is still there, and a hidden
//! source would answer that question with itself.
//!
//! shortcut: a shrink is led by a small inset and a grow lags a tenth, so between two asks a shrinking
//! source can still stand a step outside the host and a growing one is stretched and soft; the two
//! off-screen placements that would have hidden the spill do not draw at all. That is the *visible*
//! placement's price and the invisible road pays none of it, which is what `RHP_THUMBNAIL_INVISIBLE=0`
//! is for: the two behaviours are one variable apart. Upgrade the visible road by clipping the source
//! where a window of this app's can clip it, not by moving it off the screen.

use std::sync::atomic::{AtomicIsize, AtomicU32, Ordering};
use std::sync::Mutex;
use std::time::{Duration, Instant};

use windows::Win32::Foundation::{COLORREF, HWND, POINT, RECT};
use windows::Win32::Graphics::Gdi::{ClientToScreen, InvalidateRect};
use windows::Win32::Graphics::Dwm::{
    DwmRegisterThumbnail, DwmUnregisterThumbnail, DwmUpdateThumbnailProperties,
    DWM_THUMBNAIL_PROPERTIES, DWM_TNP_RECTDESTINATION, DWM_TNP_RECTSOURCE,
    DWM_TNP_SOURCECLIENTAREAONLY, DWM_TNP_VISIBLE,
};
use windows::Win32::UI::WindowsAndMessaging::{
    GetClientRect, GetForegroundWindow, GetWindowLongPtrW, GetWindowRect, GetWindowThreadProcessId,
    IsWindow, IsWindowVisible, SetForegroundWindow, SetLayeredWindowAttributes, SetWindowLongPtrW,
    SetWindowPos, ShowWindowAsync, GWL_EXSTYLE, HWND_BOTTOM, LWA_ALPHA, SW_HIDE, SWP_ASYNCWINDOWPOS,
    SWP_NOACTIVATE, SWP_NOMOVE, SWP_NOSIZE, SWP_NOZORDER, SW_SHOWNOACTIVATE, WS_EX_APPWINDOW,
    WS_EX_LAYERED, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_EX_TRANSPARENT,
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

/// **A third knob, and not a setting either: `RHP_THUMBNAIL_INVISIBLE`, on unless it says `0`.**
///
/// It is the one variable this road's experiment is run by: with it on the source wears layered
/// alpha 1, so nothing of it reaches the screen and it may stand larger than the box (see the module
/// note and `the_generous_source_size`); with it off the source is the covered window it always was and
/// a drag chases it as it always did. Only `0` turns it off — a value that names no road is not read as
/// one it does not name — and the launch line says which of the two a run was.
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
/// Read by the shipped binary and stood in by a test, for the reason the placement knob above is.
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

/// How far beyond the virtual screen's left and bottom an `offscreen` source is put: a whole desktop
/// worth, which no monitor arrangement reaches, and well inside the ±32 767 that a window message's
/// packed coordinates can carry — past which Windows silently clamps the position (see
/// `the_source_launch_origin`).
const OFFSCREEN_GAP: i32 = 1_000;
/// How much of an `edge` source is inside the virtual screen: the 2×2 pixels of its corner that stand
/// over the screen's bottom-right one.
const EDGE_INSIDE_PIXELS: i32 = 2;

/// How long a settle waits between one ask at the source and the next one: long enough for the
/// `SWP_ASYNCWINDOWPOS` post to have been pumped and for a player that is still configuring its
/// output to have finished, and short enough that a source left at another box is not on the screen
/// for a hand's notice.
const SETTLE_RETRY: Duration = Duration::from_millis(150);
/// How many asks one settle makes of the source.
///
/// **One ask is not a proof that a window went where it was told.** mpv makes its window, configures
/// its output for the frame and only then keeps the box it was given, so a placement made into that
/// is one it can drop or overwrite (measured: a source adopted at a pin's band still read back
/// 1920x1080 at mpv's own corner). A handful of asks covers a configure; a source that has not taken
/// its box after them is one the player or the Shell is overruling, which is a fact for the trace and
/// not something the tick may spin on.
const SETTLE_ASKS: u8 = 4;

/// The most a picture may be off its window's client box and still count as filling it: one pixel of
/// rounding either way, which is where an aspect worked out in integers lands (see
/// `the_picture_fills`).
const FILL_SLACK: i32 = 1;

/// How often a drag may ask the source for a new box: often enough that a hand's own pace is not what
/// the picture lags by, and rare enough that a long drag is a handful of asks rather than one per
/// pointer move (see `follow_the_band_while_dragging`). **The visible placement's only**: the invisible
/// road asks the source nothing at all between the two ends of a drag.
const DRAG_ASK_EVERY: Duration = Duration::from_millis(120);

/// A growing box is chased once it is more than this many tenths larger than the source, in integers
/// (`box * 10 > source * 11`): a coarse catch-up, so the softness of a frozen source is what a drag
/// pays for between asks rather than what it pays for all of them (see `the_drag_catch_up`).
const GROWTH_OUT_OF_TENTHS: i32 = 11;

/// How far ahead of a shrinking box the source is asked for, as pixels of inset per this many pixels
/// of the box: 8 px per axis at a 680 px box, scaled from the box itself so a large pin gets a lead a
/// step can land in, and never under [`SHRINK_LEAD_MIN`] (see `the_shrink_lead`).
const SHRINK_LEAD_PER: i32 = 85;
const SHRINK_LEAD_MIN: i32 = 8;

/// The most the source may stand larger than the pin's box on the invisible road, as a multiple of it:
/// a pin that will be dragged from a small box to a large one is the one that needs the room, but a pin
/// that stays small must not allocate a 4K surface for a film nobody will ever see at that size (see
/// `the_generous_source_size`).
const GENEROUS_CAP: i32 = 3;

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
        // arrangement of monitors can reach any part of it. **Measured dead as well**: nothing of it
        // is composed, so there is no picture to show (see the module note).
        SourcePlacement::Offscreen => (
            left - (right - left).max(width) - OFFSCREEN_GAP,
            bottom + (bottom - top).max(height) + OFFSCREEN_GAP,
        ),
        // The screen's bottom-right corner with the source's own corner over it. **Measured dead**:
        // 2 px in bounds is not a window the Shell is drawing, so the thumbnail of it is empty (see
        // the module note).
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

/// A box as a placement, a trace and an ask carry it: origin, width and height.
type SourceBox = (i32, i32, i32, i32);

/// The last box a drag asked the source for, and when: the throttle, and the test that the ask
/// before it has landed (see `follow_the_band_while_dragging`).
static DRAG_ASK: Mutex<Option<(Instant, SourceBox)>> = Mutex::new(None);

/// How many asks a drag has made at the source since the last `source settled` line was written.
///
/// **It exists so that `0 asks while dragging` is provable from a run rather than asserted here.** The
/// invisible road makes no ask at all between the two ends of a drag (see `the_band_was_placed`), and
/// the line the drag's end settles on carries this count so the log says so; the visible road's same
/// line carries the handful of catch-ups it spent, which is the other half of the A/B.
static ASKS_WHILE_DRAGGING: AtomicU32 = AtomicU32::new(0);

/// What a settle knows about the player's output filling its own window (see `the_picture_fills`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Fill {
    /// Not looked at yet for this settle.
    Unlooked,
    /// The picture, with the margins it is drawn in, covers the client area, and the window the output
    /// configured for is that client area — letterboxed films included.
    Filled,
    /// The player did not answer the question: nothing to judge and nothing to nudge.
    Unknown,
    /// The nudge was made at this instant; the picture is looked at again after `SETTLE_RETRY`.
    Nudged(Instant),
    /// Nudged and still not filled: the settle's one line has been said and nothing more is asked for
    /// this box, however many settles name it again.
    NeverFilled,
}

/// What the fill side of a settle is owed on one look (see `the_fill_look`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum FillLook {
    /// Read the player's picture size: this settle has not looked yet, or a nudge is due its look.
    Look,
    /// The nudge is still inside `SETTLE_RETRY`.
    Wait,
    /// Nothing left to look at: the settle has an answer.
    Done,
}

#[derive(Clone, Copy)]
struct Thumbnail {
    /// The host the Shell is drawing into, and the registration it answered with.
    host: isize,
    registration: isize,
    /// The rectangle pair the registration was last handed — the host's client it is drawn in, the
    /// picture's rectangle in the source and the source's own client — or nothing where none stands.
    /// **A change of any of the three is a re-point** (see `settle_the_registration`).
    shown: Option<ThumbnailBox>,
    /// The box the source was last asked for at a settle, how many asks of it have been made, when the
    /// last of them was, and whether that settle has already reported the box as never taken: what
    /// the tick reads to tell a placement that landed from one that did not (see `ask_the_source_for`).
    asked: (i32, i32, i32, i32),
    asks: u8,
    asked_at: Option<Instant>,
    gave_up: bool,
    /// How many asks the settle's box took when it landed, for the line the fill proof writes (the
    /// box look resets its own count the moment the box matches).
    matched_after: u8,
    /// The fill invariant of this settle (see `settle_the_fill`).
    fill: Fill,
    /// The shape of the film's picture, as the player's last `osd-dimensions` answer gave it — a
    /// settle's read, kept because a step of a drag crops the source by it and a step may not read the
    /// player (see `the_thumbnail_box`).
    picture: Option<f64>,
    /// The window the player last said its picture is in, as that same answer gave it: the surface the
    /// Shell is sampling, which is a box behind the source's own window after an ask the output has not
    /// taken yet (see `the_window_the_picture_is_at`).
    presented: Option<(i32, i32)>,
}

impl Thumbnail {
    const NOTHING: Self = Self {
        host: 0,
        registration: 0,
        shown: None,
        asked: (0, 0, 0, 0),
        asks: 0,
        asked_at: None,
        gave_up: false,
        matched_after: 0,
        fill: Fill::Unlooked,
        picture: None,
        presented: None,
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
/// the hover the pin is being taken up from — or nowhere at all, on the two off-screen placements (see
/// `the_source_placement`).
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
        //
        // **And it counts against the settle only where the player has said its output is
        // configured.** mpv makes its window and then configures its output for the first frame, and a
        // placement made into that is one the configure can drop or overwrite — measured: a source
        // adopted at a pin's band still read back 1920x1080 at mpv's own corner, with the ask answered
        // as made. The ask is made anyway, because it is what puts the film under the host in the
        // meantime; where the word is not yet given it is a free try and the tick asks again at once
        // instead of trusting it (see `keep_the_source_at_the_band`). The word is read here, once — a
        // pipe read per tick is not a thing this road may spend.
        if let Some(box_) = the_box_the_placement_names(host) {
            ask_the_source_for(source, host, box_, the_players_output_is_configured(pid));
        }
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

/// Follow the host: register the picture where there is none, and point it at the host's client area —
/// **the whole of it, filled by the picture** (see `the_thumbnail_box`).
///
/// **The source is not resized on every step here**, which is the whole of what makes a step cheap:
/// the destination is the host's box and the Shell stretches whatever the source's box is into it.
/// The source is resized exactly at each end of a drag (see `place_the_source`) and, between them on a
/// *visible* placement, only by the coarse catch-up a hand that is resizing it is owed (see
/// `follow_the_band_while_dragging`).
///
/// **Nor is it moved here, on the two off-screen placements** — the point of those is that a step of a
/// drag asks mpv's process for nothing at all. A `behind` source does follow the host's corner, which
/// is the only thing keeping it under the host rather than standing where the pin used to be (see
/// `the_band_moves_the_source`).
///
/// **What a drag step costs is what follows**: a comparison, a `DwmUpdateThumbnailProperties` only
/// where the rectangle pair changed, and — on a visible placement, a handful of times over a long
/// drag — one posted ask for a new source size (see `follow_the_band_while_dragging`). On the invisible
/// road a step is the destination rectangle and nothing else at all.
pub(super) fn the_band_was_placed(host: HWND) {
    let number = SOURCE.load(Ordering::SeqCst);
    if number == 0 {
        return;
    }

    if the_band_moves_the_source(the_source_placement()) {
        if the_source_is_invisible() {
            // **A drag step asks an invisible source for nothing — not a corner and not a size.** One
            // ask while the Shell is thumbnailing a window is a window *mid-reconfigure*: part of its
            // surface holds the previous scale and part the next, and the pin shows both at once. That
            // is the flicker measured on the glass, and it is the catch-up asks that cause it, so this
            // road makes none: the source is stood where the drag began (nobody can see it, so it cannot
            // spill) at the generous size that same first ask gave it, and the only thing that follows
            // the hand is the destination rectangle (see `the_box_a_placement_asks_for`).
            //
            // shortcut: a source the drag outgrows is the Shell's to stretch, so a long drag to a box
            // bigger than the generous one is soft by however far it outgrew it, and a source of a
            // shape the box does not share is the picture cropped by the box and upscaled to fill it.
            // Upgrade it by asking once more at a step where the drag is *idle* (a hand held still
            // past `SETTLE_RETRY`), not by catching up on every step.
        } else {
            if let Some(origin) = unsafe { client_origin(host) } {
                move_the_source(HWND(number as *mut _), origin);
            }
            // **And how big the source is, where a hand is pulling the box.** A step is this app's two
            // windows and a destination rectangle, but a source left at the size a drag began at is a
            // picture the Shell stretches by however far the hand has come — so a resize is followed,
            // coarsely and ahead of itself on the way down (see `follow_the_band_while_dragging`).
            follow_the_band_while_dragging(host);
        }
    }

    let Ok(mut held) = THUMBNAIL.lock() else {
        return;
    };
    // **A step crops by the client the source is at now**, never by where a settle last saw the
    // picture: the film follows the box it is drawn in (see `the_window_the_picture_is_at`).
    settle_the_registration(&mut held, host, None);
}

/// Put the source where its placement wants it, at the box the host is at: the ask a drag makes at its
/// two ends and at every box that settles without one.
///
/// **Its size is the box the pin is at**, except at the beginning of a resize on the invisible road,
/// where it is a generous box the whole drag can be drawn down from and nothing between the two ends
/// asks again (see `the_box_a_placement_asks_for`). A visible drag between two placements leaves a
/// `behind` source at the size the follow last asked for rather than at the box the drag began at (see
/// `follow_the_band_while_dragging`), so what the settle here corrects is the lead it was asked for
/// ahead of the box and the softness of a catch-up still behind it.
///
/// Nothing is posted to a window already at the box the placement names, which is what makes the call
/// at a drag's beginning cheap where the source has not left its placement since the last one; the
/// one thing it does spend is the fill read, because a placement is a fresh settle even where it names
/// the same box (see `settle_the_fill`).
///
/// `what` names the moment this is, for the one line a flash is read from: a drag's beginning, its
/// release, or a settle that is no drag's (see `note_the_cover`).
pub(super) fn place_the_source(host: HWND, what: &'static str) {
    let Some(source) = the_source_window() else {
        return;
    };
    let Some(box_) = the_box_the_placement_names(host) else {
        return;
    };
    let box_ = the_box_a_placement_asks_for(source, box_);
    // **A placement is a fresh settle even where it names the box the last settle did**: a drag
    // between the two may have left the player's output at another size, so the picture is looked at
    // again rather than taken on trust from the settle before it (see `settle_the_fill`).
    //
    // **Except where that settle already ended on `source never filled` for this very box.** The nudge
    // was made, it did not fix it, and the trace has said so — so a source a player or the Shell is
    // overruling is left alone until the box changes (a drag that ends at another box, a maximize, a
    // zoom) or the window is made again, which resets this whole record (see `note_the_source`). That
    // is what makes a repeated nudge impossible by construction rather than by the cap alone.
    if let Ok(mut held) = THUMBNAIL.lock() {
        if the_settle_is_fresh(&held, box_) {
            held.fill = Fill::Unlooked;
        }
    }
    ask_the_source_for(source, host, box_, true);
    // The box this moment's pair is drawn in, the shape it is drawn with and the crop that fills it —
    // the three numbers a flash is answerable by, at the three moments a flash can be in (see
    // `note_the_cover`).
    note_the_cover(what, host);
    // **The settle is the drag follow's previous ask**: nothing follows a box that has not landed, and
    // this is the ask whose landing the next drag's first step waits on (see
    // `follow_the_band_while_dragging`).
    if let Ok(mut last) = DRAG_ASK.lock() {
        *last = Some((Instant::now(), box_));
    }
}

/// Whether a placement naming `box_` judges the player's picture again: always, unless the last
/// settle ended on `source never filled` for that same box — the one state a nudge has already failed
/// in, and the one state a repeated nudge must not be made from (see `place_the_source`).
fn the_settle_is_fresh(held: &Thumbnail, box_: SourceBox) -> bool {
    held.fill != Fill::NeverFilled || held.asked != box_
}

/// The box a placement asks the source for: the pin's own box, or — on the invisible road, at the
/// beginning of a resize — one big enough for the whole drag to be drawn down from.
///
/// **That one ask is the whole of what a drag does to the source on this road.** Nothing between the
/// two ends asks again (see `the_band_was_placed`), so this is the last chance to give the player's
/// output a size the gesture can be drawn from: the Shell scales the source into the destination, so a
/// source at least as large as the drag will get is a picture only ever scaled *down*, and a source
/// nobody can see may be as large as that needs (see `the_generous_source_size`). Everything else — a
/// move, a maximize, a zoom, a walk, either end of any drag — is the placement's own box.
fn the_box_a_placement_asks_for(source: HWND, box_: SourceBox) -> SourceBox {
    if !(the_source_is_invisible() && pin_drag::pin_is_resizing()) {
        return box_;
    }
    let own = the_films_own_size();
    let (width, height) = the_generous_source_size((box_.2, box_.3), own);
    note_pin_trace!(
        "source generous {:?} {}x{}+{}+{} reason generous box {}x{} own {:?} cap {}x",
        source,
        width,
        height,
        box_.0,
        box_.1,
        box_.2,
        box_.3,
        own,
        GENEROUS_CAP,
    );
    (box_.0, box_.1, width, height)
}

/// How big the source may be asked to be where nothing of it can be seen: the larger of the pin's box
/// and a half again, and the film's own display size, per axis — never less than the box and never more
/// than `GENEROUS_CAP` times it.
///
/// **The film's own size is the point of it.** A box smaller than the film is a picture the Shell is
/// scaling down anyway; asking the player for the film's own size instead means a drag from a small pin
/// to a large one is drawn from a source at least as big as the film, and never upscaled at the source
/// before being scaled again by the Shell. The cap is the other half: a 160 px pin must not allocate a
/// 4K source for a film nobody is going to see at that size, so a small box simply gets three times
/// itself and is soft for the length of a drag it grows past that.
fn the_generous_source_size(box_: (i32, i32), own: Option<(i32, i32)>) -> (i32, i32) {
    let generous = |extent: i32, own: Option<i32>| {
        let half_again = (extent * 3 / 2).max(1);
        own.unwrap_or(0)
            .max(half_again)
            .clamp(1, extent.saturating_mul(GENEROUS_CAP).max(1))
    };
    (
        generous(box_.0, own.map(|own| own.0)),
        generous(box_.1, own.map(|own| own.1)),
    )
}

/// The film's own display size, as the probe answered it when the film was measured: read out of the
/// geometry cache and never asked of the player, because this runs on the first pointer message of a
/// drag and a pipe read there is a stall in the hand (see `dimensions::video_box`).
fn the_films_own_size() -> Option<(i32, i32)> {
    let (path, _) = pinned_media_owner()?;
    let (width, height) = dimensions::video_box(&path)?;
    Some((width as i32, height as i32))
}

/// Whether the player whose window is the pin's picture has said its video output is configured for a
/// frame, which is the one fact that says a placement made now will be kept.
///
/// **Read once, at adoption, where the answer is cheap and means something** (see `adopt_the_source`).
/// Nothing on the tick's road asks it: a pipe read per tick is not a thing this road may spend, and the
/// tick's own question is where the source's box actually is (see `keep_the_source_at_the_band`).
pub(super) fn the_players_output_is_configured(pid: u32) -> bool {
    pid != 0 && mpv_ipc::get_property(pid, "vo-configured").as_deref() == Some("true")
}

/// Ask again for a box the source did not take: the tick's proof that a settle landed rather than its
/// assumption that it did.
///
/// **A hand on an edge asks nothing here.** The drag's own follow is what a resize asks of the player
/// (see `follow_the_band_while_dragging`), and this check would be a second writer asking for the same
/// resize on every tick — the thrash this road exists not to have. A drag between the two ends is a
/// destination rectangle and a handful of catch-ups, and the box it settled at is this function's on
/// the tick after it (see the module note).
///
/// **Nothing here is a cross-process call.** The source's corner and size are Windows' answer about a
/// window, and the one ask that reaches the player is posted (see `ask_the_source_for`).
pub(super) fn keep_the_source_at_the_band(host: HWND) {
    if pin_drag::pin_is_dragging() {
        return;
    }
    let Some(source) = the_source_window() else {
        return;
    };
    let Some(box_) = the_box_the_placement_names(host) else {
        return;
    };
    ask_the_source_for(source, host, box_, true);
}

/// The box a placement wants the source at: the corner the placement puts it at, at the size of the
/// host's client area, both in the virtual desktop's coordinates. `None` for a host that has gone.
fn the_box_the_placement_names(host: HWND) -> Option<(i32, i32, i32, i32)> {
    let origin = unsafe { client_origin(host) }?;
    let size = unsafe { client_size(host) }?;
    let size = (size.0.max(1), size.1.max(1));
    let screen = displays::the_virtual_screen();
    let (x, y) = source_origin(
        the_source_placement(),
        (screen.left, screen.top, screen.right, screen.bottom),
        origin,
        size,
    );
    Some((x, y, size.0, size.1))
}

/// What one look at the source answers, as a rule with no window in it so that the timing and the cap
/// can be asked without one (see `ask_the_source_for`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum SettleLook {
    /// At the box the settle asks for, after this many asks of it.
    Matched(u8),
    /// Not at the box, and the last ask is too recent to make another.
    TooSoon,
    /// Not at the box, and this settle's next ask is due.
    Due,
    /// Not at the box after every ask one settle makes: the box is not being taken.
    Never,
    /// A settle that has already said the box is not being taken: nothing more to say.
    Nothing,
}

/// The look one settle makes at a source: the counts advanced where a new box begins a new settle, and
/// whether this look may ask.
fn the_settle_look(
    held: &mut Thumbnail,
    box_: (i32, i32, i32, i32),
    now_at: (i32, i32, i32, i32),
    now: Instant,
) -> SettleLook {
    if held.asked != box_ {
        held.asked = box_;
        held.asks = 0;
        held.asked_at = None;
        held.gave_up = false;
        held.matched_after = 0;
        held.fill = Fill::Unlooked;
    }
    if now_at == box_ {
        let asks = held.asks;
        held.asks = 0;
        held.asked_at = None;
        return SettleLook::Matched(asks);
    }
    if held.gave_up {
        return SettleLook::Nothing;
    }
    if held.asks >= SETTLE_ASKS {
        held.gave_up = true;
        return SettleLook::Never;
    }
    if held
        .asked_at
        .is_some_and(|at| now.duration_since(at) < SETTLE_RETRY)
    {
        return SettleLook::TooSoon;
    }
    SettleLook::Due
}

/// Ask the source for `box_`, at most as often as a settle allows, and say in the trace what the look
/// answered: the placement, the count it is at, the box the window is actually in, and the one line
/// that says the box was taken.
///
/// **A placement is a guess about another process's window**, so this is the one place that reads the
/// answer back: mpv can drop or overwrite a box it was given before its output was configured, and the
/// Shell stretches whatever box the source is really in into the destination — a stale source is a
/// picture stretched in the wrong shape, not a picture that is late (see the module note).
///
/// The ask is posted and never waited on: the read that follows it is Windows' answer about the
/// window, so nothing here waits on the player's thread.
///
/// **The box measured is the window's own and not its client area's**, because that is what a
/// placement names and what `SetWindowPos` places; mpv is launched `--no-border`, so the two are one
/// area and the trace reads as the picture's box.
///
/// **`count_the_ask` is the player's own word about whether the ask will be kept** (see
/// `adopt_the_source`): an ask made before the output is configured is a try, not one of the settle's
/// attempts, so the settle's next look is due at once rather than a retry away.
///
/// **A box the source took is half a settle**: the other half is the fill invariant, proved once for
/// the settle and once more after a nudge (see `settle_the_fill`).
fn ask_the_source_for(source: HWND, host: HWND, box_: (i32, i32, i32, i32), count_the_ask: bool) {
    // Safety: the handle is a window this module adopted, and both questions are about the window.
    if !unsafe { IsWindow(source) }.as_bool() {
        return;
    }
    let Some(now_at) = (unsafe { window_box(source) }) else {
        return;
    };
    let now = Instant::now();

    let Ok(mut held) = THUMBNAIL.lock() else {
        return;
    };
    let (x, y, width, height) = box_;
    let placement = the_source_placement();
    match the_settle_look(&mut held, box_, now_at, now) {
        SettleLook::TooSoon | SettleLook::Nothing => {}
        SettleLook::Matched(asks) => {
            if asks > 0 {
                held.matched_after = asks;
            }
            let was_filled = held.fill == Fill::Filled;
            let proved = settle_the_fill(&mut held, source, host, box_, now, count_the_ask);
            // **A settle is where the pair is re-pointed**: the source's client has just been read
            // again (the shrink a drag ended on is the picture's whole crop) and the shape with it, so
            // the registration is handed the rectangles those two make now rather than the ones it was
            // left wearing when the drag began (see `settle_the_registration`). Where that read says
            // the picture is still at another window, that window is what the crop is fitted into.
            let at = the_window_the_picture_is_at(&held, None);
            settle_the_registration(&mut held, host, at);
            // **`source settled` means both halves**: the box matched and the picture fills it. A box
            // taken again on a settle whose fill was already proved is a settle too, which is what the
            // `asks > 0` half says (the box look resets its count as soon as the box matches).
            //
            // **And it says what the drag spent.** `0 while dragging` is the invisible road's proof that
            // nothing asked the source for anything between the two ends of the gesture (see
            // `ASKS_WHILE_DRAGGING`); a handful is the visible road's catch-ups. The count is given up
            // with the line, so every settle reports its own drag and no more.
            if proved && (asks > 0 || !was_filled) {
                let while_dragging = ASKS_WHILE_DRAGGING.swap(0, Ordering::SeqCst);
                note_pin_trace!(
                    "source settled {:?} {:?} {}x{}+{}+{} host {:?} after {} asks, {} while dragging",
                    source,
                    placement,
                    width,
                    height,
                    x,
                    y,
                    host,
                    held.matched_after,
                    while_dragging,
                );
            }
        }
        SettleLook::Never => {
            note_pin_trace!(
                "source never took its box {:?} {:?} wanted {}x{}+{}+{} has {}x{}+{}+{} host {:?} after {} asks",
                source,
                placement,
                width,
                height,
                x,
                y,
                now_at.2,
                now_at.3,
                now_at.0,
                now_at.1,
                host,
                held.asks,
            );
        }
        SettleLook::Due => {
            let asks = if count_the_ask {
                held.asks += 1;
                held.asked_at = Some(now);
                held.asks
            } else {
                0
            };
            // The ask no longer decides what the picture is cropped from: the registration reads the
            // source's own client at the moment it is pointed, so a settle in flight is a crop of the
            // surface that is there rather than of the one being asked for (see `the_source_client`).
            let placed = unsafe {
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
            note_pin_trace!(
                "source placed {:?} {:?} {}x{}+{}+{} host {:?} ask {}/{} counted {} from {}x{}+{}+{} ok {}",
                source,
                placement,
                width,
                height,
                x,
                y,
                host,
                asks,
                SETTLE_ASKS,
                count_the_ask as u8,
                now_at.2,
                now_at.3,
                now_at.0,
                now_at.1,
                placed.is_ok() as u8,
            );
            // **The settle's own line, whether or not the pair moved**: an ask is the moment a box, a
            // source and a shape are all in question, and the trace says here what the Shell has been
            // handed for them (see `note_the_thumbnail_pair`).
            //
            // **And it is cropped where the picture still is**, which is the box the source was at
            // before this ask: the content is one adoption behind the window, so a crop of the window
            // this call has just asked for is a crop of a surface the Shell is not sampling (see
            // `the_window_the_picture_is_at`).
            settle_the_registration(&mut held, host, Some((now_at.2, now_at.3)));
        }
    }
}

/// Prove the fill invariant for this settle, and nudge the player's output once where it does not
/// hold (see the module note and `the_picture_fills`).
///
/// **The read is mpv's and the comparison is with the source's own client size**, which is the box the
/// Shell stretches into the destination: an output configured for another size is a picture the
/// destination shows cropped (larger), or small in a field of black (smaller), because the bitblt
/// swapchain is presented 1:1 (see the module note). `osd-dimensions` is the player's own answer about
/// the window it configured and the margins the film is drawn inside, in one read of one node (see
/// `the_drawn_picture` and `Drawn`).
///
/// **One read per settle, and one more after the nudge it makes.** A settle is the only moment on this
/// road that asks the player anything: the tick's own questions are Windows' about a window, and the
/// read is a bounded pipe request rather than a conversation (see `mpv_ipc::get_property`).
///
/// Answers whether the settle is proved: the box matched *and* the picture fills it. The caller says
/// `source settled` on that answer rather than here, because a box re-taken on a settle whose fill was
/// already proved is a settle too (see `ask_the_source_for`).
fn settle_the_fill(
    held: &mut Thumbnail,
    source: HWND,
    host: HWND,
    box_: SourceBox,
    now: Instant,
    count_the_ask: bool,
) -> bool {
    // **A free try is not the settle whose picture is judged**: the player has not yet said its
    // output is configured, so what it would answer now is about a window rather than a picture. The
    // ask that counts, on the next tick, is the one that reads (see `adopt_the_source`).
    if !count_the_ask {
        return held.fill == Fill::Filled;
    }
    let was_nudged = matches!(held.fill, Fill::Nudged(_));
    if the_fill_look(held.fill, now) != FillLook::Look {
        return held.fill == Fill::Filled;
    }
    let Some(client) = (unsafe { client_size(source) }) else {
        return held.fill == Fill::Filled;
    };
    let placement = the_source_placement();
    let (x, y, width, height) = box_;
    let drawn = the_drawn_picture(SOURCE_PROCESS.load(Ordering::SeqCst));
    // **The picture's shape is kept for the drag that may follow**: it is what every step crops the
    // source by and fills the box with, and a step asks the player nothing (see `the_thumbnail_box`).
    //
    // shortcut: this read is kept even where the settle goes on to judge it stale (`the_picture_fills`
    // below), which it is when the source's window has just been re-asked — the drag's two ends are
    // exactly that, and a shape a frame out is how the containing road came out as a bar. Upgrade it by
    // keeping only a read whose own window is the source's client; until then what a wrong shape costs is
    // a crop a pixel or a frame out, never a bar, because the box is filled from it whatever it says.
    if let Some(drawn) = drawn {
        held.picture = the_picture_shape(drawn);
        held.presented = Some(drawn.window);
    }
    let (ml, mt, mr, mb) = drawn.map_or((0, 0, 0, 0), |drawn| drawn.margins);
    match drawn {
        Some(drawn) => note_pin_trace!(
            "source fills {:?} {:?} client {}x{} drawn {}x{} picture {}x{} margins l{}t{}r{}b{} (osd-dimensions)",
            source,
            placement,
            client.0,
            client.1,
            drawn.window.0,
            drawn.window.1,
            drawn.picture.0,
            drawn.picture.1,
            ml,
            mt,
            mr,
            mb,
        ),
        None => note_pin_trace!(
            "source fills {:?} {:?} client {}x{} drawn unknown (osd-dimensions)",
            source,
            placement,
            client.0,
            client.1,
        ),
    }

    if drawn.is_some_and(|drawn| the_picture_fills(drawn, client)) {
        held.fill = Fill::Filled;
        return true;
    }
    let Some(drawn) = drawn else {
        // A player that cannot say is not a picture proven wrong: nothing to nudge and nothing to
        // retry, so the settle is written down as unproved rather than as taken.
        held.fill = Fill::Unknown;
        note_pin_trace!(
            "source box settled, picture unread {:?} {:?} {}x{}+{}+{} host {:?} after {} asks",
            source,
            placement,
            width,
            height,
            x,
            y,
            host,
            held.matched_after,
        );
        return false;
    };
    if was_nudged {
        held.fill = Fill::NeverFilled;
        note_pin_trace!(
            "source never filled {:?} {:?} drawn {}x{} picture {}x{} wanted {}x{} host {:?}",
            source,
            placement,
            drawn.window.0,
            drawn.window.1,
            drawn.picture.0,
            drawn.picture.1,
            client.0,
            client.1,
            host,
        );
        return false;
    }
    let nudged = nudge_the_source(source, box_, client);
    held.fill = Fill::Nudged(now);
    note_pin_trace!(
        "source nudged {:?} {:?} {}x{}+{}+{} drawn {}x{} client {}x{} host {:?} ok {}",
        source,
        placement,
        width,
        height,
        x,
        y,
        drawn.window.0,
        drawn.window.1,
        client.0,
        client.1,
        host,
        nudged as u8,
    );
    false
}

/// What the fill side of a settle is owed on one look: a read, a wait for the nudge's retry, or
/// nothing at all (see `settle_the_fill`).
fn the_fill_look(fill: Fill, now: Instant) -> FillLook {
    match fill {
        Fill::Unlooked => FillLook::Look,
        Fill::Nudged(at) if now.duration_since(at) < SETTLE_RETRY => FillLook::Wait,
        Fill::Nudged(_) => FillLook::Look,
        Fill::Filled | Fill::Unknown | Fill::NeverFilled => FillLook::Done,
    }
}

/// Whether what the player is presenting fills the source window the Shell is thumbnailing it
/// through: **the fill invariant**, and the one question that tells a picture the Shell may stretch
/// from one it may not.
///
/// Two halves, and both are the same claim from two sides. The picture the player draws *plus the
/// margins it draws it in* has to cover the client area — that is what makes a film wider than the box
/// pass: its margins are the black it is honestly shown with (see `Drawn`) — and the window the video
/// output configured for has to *be* that client area, because an output a size behind the window is
/// what mpv presents 1:1 with `--d3d11-flip=no`: a larger surface is cropped and a smaller one leaves
/// black. One pixel of slack either way is where an aspect worked out in integers lands.
///
/// shortcut: the margins are read as the player reports them, so a film the player crops with
/// `--video-crop`-style margins outside the OSD rectangle would read as one that does not fill; nothing
/// on this road sets a crop the box is not already at. Upgrade it by asking `video-params` for the
/// frame's own shape and comparing the fitted rectangle with it.
fn the_picture_fills(drawn: Drawn, client: (i32, i32)) -> bool {
    let (ml, mt, mr, mb) = drawn.margins;
    let covered = (drawn.picture.0 + ml + mr, drawn.picture.1 + mt + mb);
    let a_picture = drawn.picture.0 > 0 && drawn.picture.1 > 0;
    let covers_the_client = (covered.0 - client.0).abs() <= FILL_SLACK
        && (covered.1 - client.1).abs() <= FILL_SLACK;
    let is_the_client = (drawn.window.0 - client.0).abs() <= FILL_SLACK
        && (drawn.window.1 - client.1).abs() <= FILL_SLACK;
    a_picture && covers_the_client && is_the_client
}

/// What mpv says its video output is presenting, in the source window's own pixels: the window the
/// output configured for, the picture it fitted inside that window, and the letterbox margins around
/// it — `osd-dimensions`, one node property and therefore one read (see `the_drawn_picture`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
struct Drawn {
    /// The VO's window in pixels: the surface it is presenting from.
    window: (i32, i32),
    /// The rectangle the film is drawn in: the window less its four margins.
    picture: (i32, i32),
    /// The margins the film's own shape leaves, as left, top, right and bottom.
    margins: (i32, i32, i32, i32),
}

/// The shape of the picture a settle read, as a width over a height — or nothing for a player with no
/// picture in its window, which is not a shape this app may crop by (see `the_thumbnail_box`).
fn the_picture_shape(drawn: Drawn) -> Option<f64> {
    (drawn.picture.0 > 0 && drawn.picture.1 > 0)
        .then(|| drawn.picture.0 as f64 / drawn.picture.1 as f64)
}

/// mpv's own answer about the output it is presenting, in one read of `osd-dimensions` — or nothing
/// where the player has no value for it, which is a player with no configured output (see `Drawn`).
///
/// **`w`/`h` are the VO window, and `dwidth`/`dheight` are not.** `dwidth`/`dheight` are the film's
/// *display* size — what it would be drawn at were it not fitted to a window — which is why reading
/// them as the drawn picture made this settle's test impossible to pass and the nudge fire on every
/// settle (see the module note). `osd-dimensions` is the output's own geometry: measured on mpv 0.41.0
/// with a 1920×1080 film in a 640×300 window it answers
/// `{"w":640,"h":300,"par":1.000000,"aspect":2.133333,"mt":0,"mb":0,"ml":53,"mr":54}`, and the same
/// player still answers 640×300 after that window's client was made 800×400 — which is exactly the
/// stale surface this settle exists to see.
fn the_drawn_picture(pid: u32) -> Option<Drawn> {
    if pid == 0 {
        return None;
    }
    let node = mpv_ipc::get_property(pid, "osd-dimensions")?;
    let (width, height) = (a_node_field(&node, "w")?, a_node_field(&node, "h")?);
    let (ml, mt, mr, mb) = (
        a_node_field(&node, "ml").unwrap_or(0),
        a_node_field(&node, "mt").unwrap_or(0),
        a_node_field(&node, "mr").unwrap_or(0),
        a_node_field(&node, "mb").unwrap_or(0),
    );
    Some(Drawn {
        window: (width, height),
        picture: ((width - ml - mr).max(0), (height - mt - mb).max(0)),
        margins: (ml, mt, mr, mb),
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

/// Ask the player's window for a size one pixel off the box and then for the box itself: the
/// excursion that makes a `WM_SIZE` a VO stuck at an old buffer must answer.
///
/// **mpv reconfigures its output on `WM_SIZE`, and a repeated identical size produces no message**, so
/// a source whose window already reads back at the band's box can have a swapchain a size behind it
/// with nothing left to correct it (see the module note). One pixel is the smallest difference that
/// guarantees the message; the ask that follows puts the window back, and both are posted rather
/// than made (see `ask_the_source_for`). The excursion leaves 1 px of the source outside the host for
/// the retry window, which is the price of forcing the message with a size the window will take.
fn nudge_the_source(source: HWND, box_: (i32, i32, i32, i32), client: (i32, i32)) -> bool {
    let (x, y, width, height) = box_;
    let flags = SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS;
    let excursion = unsafe {
        SetWindowPos(
            source,
            HWND::default(),
            x,
            y,
            client.0.max(1) + 1,
            client.1.max(1) + 1,
            flags,
        )
    };
    let exact = unsafe { SetWindowPos(source, HWND::default(), x, y, width, height, flags) };
    excursion.is_ok() && exact.is_ok()
}

/// Follow a box a hand is resizing, coarsely and asymmetrically: a growing box is chased only once it
/// is a tenth larger than the source, and a shrinking one is asked for ahead of itself so a step of
/// lag cannot leave the source spilling out from under the host.
///
/// **A frozen source is not free, it is an upscale.** A source left at the size a drag began at is a
/// picture the Shell stretches into every box the hand passes through — 680 px grown to 2263 px is a
/// 3.3× upscale for the whole of the growth — and no swapchain model hides that. So the source does
/// follow, but only as a catch-up: a handful of asks over a long drag, throttled to one per
/// `DRAG_ASK_EVERY`, and never while the ask before it has not been taken (a second ask into an output
/// that has not configured is what the settle's cap exists to avoid).
///
/// **Grow may lag and shrink may not**, because the two are not the same defect on this placement: a
/// source smaller than the box is hidden behind the host and stretched *into* the box (geometry stays
/// right, the picture is soft), while a source larger than the box is picture standing outside the
/// pin until the next ask — the spill a `behind` source is known for. So a shrink's ask is inset from
/// the box by `the_shrink_lead` and a grow's is the box itself.
///
/// Nothing is asked on any other placement, where the source is not a window that follows the host at
/// all, nothing while no hand is on an edge (see `pin_drag::pin_is_dragging`), and **nothing at all on
/// the invisible road**, whose source is generous and unseen instead (see `the_band_was_placed`).
fn follow_the_band_while_dragging(host: HWND) {
    if !pin_drag::pin_is_dragging() || !the_band_moves_the_source(the_source_placement()) {
        return;
    }
    let Some(source) = the_source_window() else {
        return;
    };
    let Some(box_) = the_box_the_placement_names(host) else {
        return;
    };
    let Some(current) = (unsafe { window_box(source) }) else {
        return;
    };
    let Some(((width, height), reason)) =
        the_drag_catch_up((box_.2, box_.3), (current.2, current.3))
    else {
        return;
    };
    let now = Instant::now();
    let Ok(mut last) = DRAG_ASK.lock() else {
        return;
    };
    if last.is_some_and(|(asked_at, asked_box)| {
        let since = now.duration_since(asked_at);
        since < DRAG_ASK_EVERY || (current != asked_box && since < SETTLE_RETRY)
    }) {
        return;
    }
    let asked_box = (box_.0, box_.1, width, height);
    let placed = unsafe {
        SetWindowPos(
            source,
            HWND::default(),
            asked_box.0,
            asked_box.1,
            asked_box.2,
            asked_box.3,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS,
        )
    };
    *last = Some((now, asked_box));
    ASKS_WHILE_DRAGGING.fetch_add(1, Ordering::Relaxed);
    note_pin_trace!(
        "source followed {:?} {}x{}+{}+{} host {:?} from {}x{} {} ok {}",
        source,
        width,
        height,
        asked_box.0,
        asked_box.1,
        host,
        current.2,
        current.3,
        reason,
        placed.is_ok() as u8,
    );
}

/// What a drag step asks the source for, if anything: the size to ask for and the reason a trace
/// gives for it.
///
/// The two answers are the two defects of a frozen source: a box more than a tenth larger is a picture
/// the Shell stretches, and a box smaller than the source is picture the host no longer covers (see
/// `follow_the_band_while_dragging`). Shrink is asked about first, because spill is the worse of the
/// two, and a box that is neither is a source already close enough to leave alone.
fn the_drag_catch_up(
    box_size: (i32, i32),
    source_size: (i32, i32),
) -> Option<((i32, i32), &'static str)> {
    let (box_width, box_height) = box_size;
    let (source_width, source_height) = source_size;
    if source_width > box_width || source_height > box_height {
        return Some((
            (
                (box_width - the_shrink_lead(box_width)).max(1),
                (box_height - the_shrink_lead(box_height)).max(1),
            ),
            "shrink lead",
        ));
    }
    if box_width * 10 > source_width * GROWTH_OUT_OF_TENTHS
        || box_height * 10 > source_height * GROWTH_OUT_OF_TENTHS
    {
        return Some(((box_width, box_height), "grow catch-up"));
    }
    None
}

/// How far ahead of a shrinking box the source is asked for, per axis (see `the_drag_catch_up`).
fn the_shrink_lead(extent: i32) -> i32 {
    (extent / SHRINK_LEAD_PER).max(SHRINK_LEAD_MIN)
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
    let Some(source) = the_source_window() else {
        return;
    };
    if the_source_placement() != SourcePlacement::Behind {
        return;
    }
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
    /// Whether the crop had to be clamped inside the picture rather than being the extent the box's
    /// shape asks for, which a trace reads and nothing else does (see `note_the_cover`).
    clamped: bool,
}

/// **What a thumbnail of the source shows: the box, filled.** The destination is the host's whole client
/// area and the source is the rectangle of the picture that *is* the box's shape — so the film covers its
/// box and there is nothing left for the host's black to show through (see `the_crop_filling`).
///
/// `aspect` is the picture's own width over its height, as a settle's `osd-dimensions` read gave it, and
/// `client` is the box the source is at **as its window answers it at this call**. The two together are
/// the source's picture rectangle, which is why a step of a drag asks the player nothing: the shape is
/// the one the settle read and the client is Windows' answer about the window.
///
/// **A destination fitted into the box was a bar of black waiting for the two numbers to disagree**, and
/// they do disagree at exactly the two moments nothing else on this road re-reads them: a drag's
/// beginning and its release are where the picture's shape is read again (see `place_the_source`). A
/// box whose aspect was a frame out — a source mid-placement, a shape read from an output that is still
/// configuring — came out as the film squeezed into the middle of the pin with a pillar either side, and
/// that is the flash the hand sees on the press and again on the release. **Filling the box is bars made
/// impossible by construction**: whatever the two shapes are, the destination is the client and the crop
/// is the part of the picture that has the client's shape, so there is no rectangle left over to read as
/// the wrong thing.
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
/// clamped inside the picture rather than being the extent the shape asks for.
///
/// **A picture wider than the box keeps its height and gives up its sides; one no wider keeps its width
/// and gives up its top and bottom.** Integer and clamped, so the rectangle is always inside the picture
/// and never empty, whatever the two shapes are and however the rounding falls.
fn the_crop_filling(
    picture: (i32, i32, i32, i32),
    box_: (i32, i32),
) -> ((i32, i32, i32, i32), bool) {
    let (left, top, width, height) = picture;
    let (width, height) = (width.max(1), height.max(1));
    let (box_width, box_height) = (box_.0.max(1), box_.1.max(1));
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

/// The client box the source is at, read off the window at the moment the registration is pointed: the
/// source is launched `--no-border`, so its window box and its client area are one rectangle. The
/// host's own box stands in for a source that has gone.
///
/// **This is a question Windows answers about a window, not a read of the player** — the same one the
/// tick asks with `window_box` — so it spends a step nothing it is not already spending. What it
/// replaces is this app's own record of the last ask, which a settle still in flight leaves a size
/// behind: a source rectangle built against that is sampled from a surface that is no longer there, or
/// past the end of the one that is, and the Shell answers a sample outside the surface with nothing —
/// the hole the picture road showed as a band of desktop inside the pin.
fn the_source_client(source: HWND, host: (i32, i32)) -> (i32, i32) {
    match unsafe { client_size(source) } {
        Some(client) if client.0 > 0 && client.1 > 0 => client,
        _ => host,
    }
}

/// **The window the picture is actually in, where a call knows it better than the source's own
/// client.** The window's client is what this app has *asked for*; the picture is one adoption behind
/// it, so:
///
/// - a settle that has just asked the source for another box (`before_the_ask`, the box the window
///   stood at when the ask was made) crops where the content still is, not where the ask is sending it;
/// - a settle whose own read said the picture does not fill the client crops by the window the player
///   named, which is the surface the Shell is sampling.
///
/// Nothing is passed for a step of a drag, which crops by the source's live client: that is the moment
/// the window and the film move together (see `the_band_was_placed`).
fn the_window_the_picture_is_at(
    held: &Thumbnail,
    before_the_ask: Option<(i32, i32)>,
) -> Option<(i32, i32)> {
    let at = before_the_ask.or(match held.fill {
        Fill::Nudged(_) | Fill::NeverFilled => held.presented,
        _ => None,
    })?;
    (at.0 > 0 && at.1 > 0).then_some(at)
}

/// Whether a registration wearing `held.shown` has to be pointed at `wanted`: **nothing else is a
/// reason not to re-point.** Another host client, another source client or another picture shape all
/// change what the pair is, and a registration left wearing the old pair is a picture cropped off the
/// surface the Shell is sampling — the three bands of the picture road, and the one cache that
/// compared the box alone (see `settle_the_registration`).
fn the_registration_wants(held: &Thumbnail, wanted: ThumbnailBox) -> bool {
    held.registration == 0 || held.shown != Some(wanted)
}

/// Register the picture in the host's client area, or point the registration the host already wears at
/// the rectangles it is drawn with **now**: the host's live client area, the window the picture is in
/// and the film's shape as the last settle read it — all read at this call, none of them a box a step
/// cached.
///
/// **The invariant is one line long**: the destination rectangle is the host's whole client, so every
/// pixel of the host is drawn by the Shell and a bar is a rectangle nothing could have left over, and
/// the host's own erase paints that client under it — which is what a client the Shell has not yet
/// composed, or one whose source is not drawing, shows rather than the desktop. What breaks it is a
/// pair built against a box the host is no longer in, or a source rectangle reaching past the surface
/// the Shell is sampling, so both are read here and the pair is traced whether it moved or not (see
/// `note_the_thumbnail_pair` and `note_the_cover`).
///
/// `at` is the window the picture is in where the caller knows it better than the source's own client —
/// the box a settle has just asked away from, or the one the player named in its own answer — and the
/// source's live client, read here and once, everywhere else (see `the_window_the_picture_is_at`).
fn settle_the_registration(held: &mut Thumbnail, host: HWND, at: Option<(i32, i32)>) {
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
    let shown = the_thumbnail_box(client, at.unwrap_or(live), held.picture, the_picture_road());
    if !the_registration_wants(held, shown) {
        note_the_thumbnail_pair("held", host, client, live, shown);
        return;
    }

    if held.registration == 0 {
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
            "thumbnail properties failed host {:?} box {:?} placement {:?}",
            host,
            client,
            the_source_placement(),
        );
        let _ = unsafe { DwmUnregisterThumbnail(held.registration) };
        *held = Thumbnail::NOTHING;
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
/// window the picture is in (`pic`) and the exact rectangles handed over. **`host` and `box` disagree
/// only where something stands on the host's client** — a border, a frame style — which is a picture
/// that cannot line up with its pin's box, and **`cli` and `pic` disagree only where the source's window
/// and its content are a moment apart**, which is the one way a crop can be of a surface the Shell is
/// not sampling.
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

/// **The one line a flash is read from, and only at the three moments a flash can be in**: the box, the
/// picture the last settle read, and the rectangle of that picture the box is filled with — the aspect
/// of each, so a line that does not explain a bar says why it does not.
///
/// Written at a drag's beginning, at its release and at every settle, and **never on a step of a drag**:
/// those three are the moments the pair is re-pointed from a shape that was read again (see
/// `place_the_source`), and they are the whole of what this road may spend on the question.
fn note_the_cover(what: &str, host: HWND) {
    if !super::pin_trace::asked() {
        return;
    }
    let Some(box_) = (unsafe { client_size(host) }) else {
        return;
    };
    let source = HWND(SOURCE.load(Ordering::SeqCst) as *mut _);
    let client = the_source_client(source, box_);
    let held = THUMBNAIL.lock().ok();
    let aspect = held.as_ref().and_then(|held| held.picture);
    let at = held
        .as_ref()
        .and_then(|held| the_window_the_picture_is_at(held, None))
        .unwrap_or(client);
    let shown = the_thumbnail_box(box_, at, aspect, the_picture_road());
    let (_, _, width, height) = if shown.source.is_some() {
        a_picture_fitted_to(aspect, at)
    } else {
        (0, 0, at.0, at.1)
    };
    let (src_left, src_top, src_right, src_bottom) =
        shown.source.unwrap_or((0, 0, at.0, at.1));
    let an_aspect = |extent: (i32, i32)| extent.0 as f64 / extent.1.max(1) as f64;
    note_pin_trace!(
        "thumbnail cover {} box {}x{} a{:.4} cli {}x{} pic {}x{} a{:.4} src {}x{}+{}+{} clamp {} ({}={})",
        what,
        box_.0,
        box_.1,
        an_aspect(box_),
        client.0,
        client.1,
        width,
        height,
        an_aspect((width, height)),
        src_right - src_left,
        src_bottom - src_top,
        src_left,
        src_top,
        shown.clamped as u8,
        PICTURE_VARIABLE,
        the_picture_road().named(),
    );
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

    /// **A registration is pointed again at everything the rectangles are built from**: the host's
    /// client, the source's client and the film's shape — not just the host's box. A source re-asked at
    /// a settle is a different source client, and a registration left wearing the old pair is a picture
    /// cropped against a surface that has gone (see `settle_the_registration`).
    #[test]
    fn a_settle_re_points_a_registration_whatever_part_of_the_pair_changed() {
        let client = (1000, 560);
        let aspect = Some(2595.0 / 1460.0);
        let shown_at = |box_: (i32, i32), source: (i32, i32), aspect: Option<f64>| {
            the_thumbnail_box(box_, source, aspect, PictureRoad::Picture)
        };
        let generous = shown_at(client, (2731, 1536), aspect);
        let mut holding = Thumbnail {
            host: 0x2000,
            registration: 9,
            shown: Some(generous),
            picture: aspect,
            ..Thumbnail::NOTHING
        };

        assert!(
            !the_registration_wants(&holding, generous),
            "a settle that moved nothing asks nothing"
        );

        let shrunk = shown_at(client, (1000, 560), aspect);
        assert_ne!(generous.source, shrunk.source, "the source's own client changed");
        assert!(
            the_registration_wants(&holding, shrunk),
            "a source re-asked at a settle is a re-point even where the box did not move"
        );

        let read_late = shown_at(client, (1000, 560), Some(16.0 / 9.0));
        assert!(
            the_registration_wants(&holding, read_late),
            "and a shape the settle read after an unread one is another pair"
        );
        assert!(
            the_registration_wants(&holding, shown_at((1200, 560), (1000, 560), aspect)),
            "as is the host's own client"
        );

        holding.shown = Some(shrunk);
        assert!(
            !the_registration_wants(&holding, shrunk),
            "and the pair it was just handed is not handed again"
        );
        holding.registration = 0;
        assert!(
            the_registration_wants(&holding, shrunk),
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

    /// **A release is cropped where the picture still is, not where the source's window has already
    /// gone.** The measured release: `source fills … client 1332x749 drawn 2880x1460 picture 2595x1460
    /// margins l142t0r143b0`, and the pair that followed it was cropped `1331x748+0+0` — the top-left
    /// corner of a surface still holding the generous picture, whose 142 px margin is the pillar that
    /// reads as the film squished (see `the_window_the_picture_is_at`).
    #[test]
    fn a_release_crops_the_picture_where_the_player_says_it_still_is() {
        let box_ = (1332, 749);
        let live = box_;
        let surface = (2880, 1460);
        let aspect = Some(2595.0 / 1460.0);
        let mut held = Thumbnail {
            fill: Fill::Nudged(Instant::now()),
            presented: Some(surface),
            ..Thumbnail::NOTHING
        };

        let at = the_window_the_picture_is_at(&held, None).expect("the player named its window");
        assert_eq!(
            at, surface,
            "a picture that does not fill the client is where it was read"
        );
        assert_ne!(at, live, "which is not the client the window has already taken");
        let shown = the_thumbnail_box(box_, at, aspect, PictureRoad::Picture);
        assert_eq!(
            shown.source,
            Some((142, 0, 2737, 1459)),
            "the picture's shape at the surface's scale, centred: 2595x1459 at +142, never the \
             whole of the box-sized client"
        );
        assert_eq!(
            shown.destination,
            (0, 0, box_.0, box_.1),
            "and the box is still what is filled"
        );
        assert_eq!(
            the_thumbnail_box(box_, live, aspect, PictureRoad::Picture).source,
            Some((0, 0, 1331, 748)),
            "which is the pair the release was cropped with: the corner of the generous surface"
        );

        held.fill = Fill::Filled;
        assert_eq!(
            the_window_the_picture_is_at(&held, None),
            None,
            "a picture proved to fill its client is the client, and nothing has to stand in for it"
        );
    }

    /// **A drag's beginning is cropped against the window the source has, not the generous box it has
    /// been asked for.** The measured line: `thumbnail cover begin box 1920x1080 a1.7778 pic
    /// 2596x1460 a1.7781 src 2596x1460+142+0` — a crop reaching 818 px past the surface the Shell is
    /// sampling, which is the film zoomed, or black past where it really has a picture.
    #[test]
    fn a_drag_start_crops_the_window_the_source_has_not_the_plan() {
        let box_ = (1920, 1080);
        let plan = (2880, 1460);
        let content = (1920, 1080);

        let at = the_window_the_picture_is_at(&Thumbnail::NOTHING, Some(content))
            .expect("the box the source stood at before this ask is where its picture is");
        assert_eq!(
            at, content,
            "the ask has moved the window; the picture has not moved yet"
        );
        let shown = the_thumbnail_box(box_, at, Some(16.0 / 9.0), PictureRoad::Picture);
        assert_eq!(
            shown.source,
            Some((0, 0, 1920, 1080)),
            "the film's own window, whole, and never a rect past 1920"
        );
        assert_eq!(
            the_thumbnail_box(box_, plan, Some(16.0 / 9.0), PictureRoad::Picture).source,
            Some((142, 0, 2738, 1460)),
            "which is what the generous box would have cropped the window it has not taken"
        );
        assert_eq!(
            the_window_the_picture_is_at(&Thumbnail::NOTHING, None),
            None,
            "a step of a drag crops by the source's live client, which the caller reads"
        );
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

    /// A step of a drag asks nothing of the source on the two off-screen placements: what follows the
    /// host is the destination rectangle and nothing else (see `the_band_was_placed`).
    #[test]
    fn only_a_source_behind_the_host_follows_it() {
        assert!(the_band_moves_the_source(SourcePlacement::Behind));
        assert!(!the_band_moves_the_source(SourcePlacement::Offscreen));
        assert!(!the_band_moves_the_source(SourcePlacement::Edge));
    }

    /// A settle asks a bounded number of times, waits between two asks, and says once that the box was
    /// not taken — which is what tells a placement that landed from one that did not (see
    /// `ask_the_source_for`).
    #[test]
    fn a_settle_asks_a_bounded_number_of_times_and_waits_between_them() {
        let band = (10, 20, 640, 360);
        let elsewhere = (0, 0, 1920, 1080);
        let mut held = Thumbnail::NOTHING;
        let mut now = Instant::now();

        assert_eq!(
            the_settle_look(&mut held, band, elsewhere, now),
            SettleLook::Due,
            "a fresh settle asks at once"
        );
        held.asks += 1;
        held.asked_at = Some(now);
        assert_eq!(
            the_settle_look(&mut held, band, elsewhere, now),
            SettleLook::TooSoon,
            "and is not asked again inside the retry"
        );
        assert_eq!(
            the_settle_look(&mut held, band, band, now),
            SettleLook::Matched(1),
            "a box the source is in is the settle, and the count says what it took"
        );
        assert_eq!(held.asks, 0, "and the count is given up with it");

        now += SETTLE_RETRY;
        for ask in 1..=SETTLE_ASKS {
            assert_eq!(
                the_settle_look(&mut held, band, elsewhere, now),
                SettleLook::Due,
                "ask {ask} of a source that never takes the box"
            );
            held.asks = ask;
            held.asked_at = Some(now);
            now += SETTLE_RETRY;
        }
        assert_eq!(
            the_settle_look(&mut held, band, elsewhere, now),
            SettleLook::Never,
            "the cap is where the settle says the box is not being taken"
        );
        assert_eq!(
            the_settle_look(&mut held, band, elsewhere, now),
            SettleLook::Nothing,
            "and says it once, so a pin that is up does not fill the trace with it"
        );

        assert_eq!(
            the_settle_look(&mut held, (11, 20, 640, 360), elsewhere, now),
            SettleLook::Due,
            "another box is another settle, with its own count"
        );
    }

    /// **The fill invariant, which is the second half of a settle**: a picture the size of the window
    /// fills it, a picture letterboxed by the film's own shape fills it too, and an output configured
    /// for another window does not — the two symptoms being crop (a surface larger than the window) and
    /// a small picture in a field of black (smaller on both axes), which is what the bitblt swapchain
    /// presents (see `the_picture_fills`).
    #[test]
    fn a_picture_fills_its_window_exactly_or_as_a_letterbox_of_it() {
        let window = (2396, 1349);
        let drawn = |window: (i32, i32), margins: (i32, i32, i32, i32)| Drawn {
            window,
            picture: (
                window.0 - margins.0 - margins.2,
                window.1 - margins.1 - margins.3,
            ),
            margins,
        };

        assert!(
            the_picture_fills(drawn(window, (0, 0, 0, 0)), window),
            "the picture is the window it is drawn in"
        );
        assert!(
            the_picture_fills(drawn((2395, 1348), (0, 0, 0, 0)), window),
            "and an aspect worked out in integers may land a pixel out"
        );
        assert!(
            the_picture_fills(drawn(window, (0, 108, 0, 108)), window),
            "a wider film letterboxed in the box still covers it once its margins are counted: that \
             is mpv fitting the film and not an output that never resized"
        );
        assert!(
            the_picture_fills(drawn(window, (300, 0, 300, 0)), window),
            "and a taller one letterboxes on the other axis"
        );

        // Measured on mpv 0.41.0: a 1920x1080 film in a 640x300 window.
        assert!(the_picture_fills(
            Drawn {
                window: (640, 300),
                picture: (533, 300),
                margins: (53, 0, 54, 0),
            },
            (640, 300)
        ));

        assert!(
            !the_picture_fills(drawn((1920, 1080), (0, 0, 0, 0)), (867, 487)),
            "an output configured for another window is a surface a size behind it, which the Shell \
             presents cropped"
        );
        assert!(
            !the_picture_fills(drawn((441, 248), (0, 0, 0, 0)), (867, 487)),
            "or smaller, which it presents as black around a small picture"
        );
        assert!(
            !the_picture_fills(
                Drawn {
                    window: (867, 487),
                    picture: (0, 0),
                    margins: (0, 0, 0, 0),
                },
                (867, 487)
            ),
            "and a player with no picture in its window has not filled it with one"
        );
    }

    /// A settle reads the picture once and, where it does not fill, exactly one nudge is made and the
    /// retry is what waits between the two reads — the cap that keeps a stuck output from being asked
    /// for the same excursion for the life of the pin (see `the_fill_look`).
    #[test]
    fn a_settle_reads_the_fill_once_and_once_more_after_its_one_nudge() {
        let now = Instant::now();
        assert_eq!(
            the_fill_look(Fill::Unlooked, now),
            FillLook::Look,
            "a fresh settle reads the picture"
        );
        assert_eq!(
            the_fill_look(Fill::Filled, now),
            FillLook::Done,
            "and a picture that fills it needs nothing more"
        );
        assert_eq!(
            the_fill_look(Fill::Unknown, now),
            FillLook::Done,
            "a player that cannot say is not read again: there is nothing to act on"
        );
        assert_eq!(
            the_fill_look(Fill::Nudged(now), now),
            FillLook::Wait,
            "the nudge's own retry is how long the VO is given to answer it"
        );
        assert_eq!(
            the_fill_look(Fill::Nudged(now), now + SETTLE_RETRY),
            FillLook::Look,
            "and after it the one look that says whether the nudge took"
        );
        assert_eq!(
            the_fill_look(Fill::NeverFilled, now),
            FillLook::Done,
            "a source that never filled says so once and is left alone"
        );
    }

    /// A drag follows the box coarsely on the way up and leads it on the way down: a source that is a
    /// little smaller is left for the drag's end, a source a tenth smaller is chased, and a source
    /// larger than the box is asked for a box inset so a step of lag cannot spill
    /// (see `the_drag_catch_up`).
    #[test]
    fn a_drag_chases_a_growing_box_coarsely_and_leads_a_shrinking_one() {
        let source = (680, 383);
        assert_eq!(
            the_drag_catch_up(source, source),
            None,
            "a source already at the box is asked for nothing"
        );
        assert_eq!(
            the_drag_catch_up((740, 383), source),
            None,
            "and a box under a tenth larger is left for the drag's own end"
        );
        assert_eq!(
            the_drag_catch_up((760, 383), source),
            Some(((760, 383), "grow catch-up")),
            "past a tenth the source is chased, and for the box itself"
        );
        assert_eq!(
            the_drag_catch_up((680, 430), source),
            Some(((680, 430), "grow catch-up")),
            "on either axis: a stretch is a stretch whichever way the hand came"
        );
        assert_eq!(
            the_drag_catch_up((2263, 1274), source),
            Some(((2263, 1274), "grow catch-up")),
            "a long drag is a handful of asks and not one per step"
        );
        assert_eq!(
            the_drag_catch_up((600, 340), source),
            Some(((592, 332), "shrink lead")),
            "a shrink is asked for ahead of the box, so one step of lag cannot leave the source \
             outside the host"
        );
        assert_eq!(
            the_drag_catch_up((680, 340), source),
            Some(((672, 332), "shrink lead")),
            "and the lead is per axis: a box pulled shorter only leads on its height"
        );
    }

    /// The lead grows with the box and never falls under the minimum, so a pin too small for a scaled
    /// lead still gets one (`the_shrink_lead`).
    #[test]
    fn the_shrink_lead_scales_from_the_box_and_has_a_floor() {
        assert_eq!(the_shrink_lead(680), SHRINK_LEAD_MIN);
        assert!(the_shrink_lead(2263) > SHRINK_LEAD_MIN, "a big box gets a bigger lead");
        assert_eq!(the_shrink_lead(10), SHRINK_LEAD_MIN, "and a small one is not less than it");
    }

    /// **The generous box an invisible source is asked for at the beginning of a resize**: the larger
    /// of the box and a half again and the film's own display size, never under the box and never more
    /// than `GENEROUS_CAP` times it — so a small pin cannot allocate a 4K source and a big one never
    /// asks for less than the film it is showing (see `the_generous_source_size`).
    #[test]
    fn a_generous_source_is_the_box_scaled_up_or_the_films_own_size_within_a_cap() {
        // A pin smaller than the film gets the film's own size, not three times a tiny box.
        assert_eq!(the_generous_source_size((867, 487), Some((1920, 1080))), (1920, 1080));
        // A pin much larger than the film gets half again, which is more than the film.
        assert_eq!(the_generous_source_size((2263, 1274), Some((1920, 1080))), (3394, 1911));
        // A pin the film's own size gets half again as well: a growth is always worth the room.
        assert_eq!(the_generous_source_size((1920, 1080), Some((1920, 1080))), (2880, 1620));
        // A tiny pin is capped at three times itself rather than at the film's 1920.
        assert_eq!(the_generous_source_size((160, 90), Some((1920, 1080))), (480, 270));
        // And a film nothing is known about is still a half again, once only, capped.
        assert_eq!(the_generous_source_size((200, 100), None), (300, 150));
        assert_eq!(the_generous_source_size((3000, 1600), None), (4500, 2400));
        // Never under the box, and never zero for a box that is nothing.
        assert!(
            the_generous_source_size((1, 1), Some((0, 0))) == (1, 1),
            "a pin of one pixel asks for what it can be drawn at and no more"
        );
        assert_eq!(the_generous_source_size((0, 0), None), (1, 1));
        for box_ in [(640, 360), (160, 90), (2263, 1274)] {
            let (width, height) = the_generous_source_size(box_, Some((1920, 1080)));
            assert!(width >= box_.0 && height >= box_.1, "never smaller than the box: {box_:?}");
            assert!(
                width <= box_.0 * GENEROUS_CAP && height <= box_.1 * GENEROUS_CAP,
                "and never more than the cap allows: {box_:?}"
            );
        }
    }

    /// A box a settle already gave up on is not judged again — which is what makes the nudge fire
    /// once for a box and never again — while any other box (or a settled picture) is a fresh settle
    /// (see `the_settle_is_fresh`).
    #[test]
    fn a_box_that_never_filled_is_not_judged_again_until_it_changes() {
        let band = (10, 20, 640, 360);
        let gone_to = Thumbnail {
            asked: band,
            fill: Fill::NeverFilled,
            ..Thumbnail::NOTHING
        };
        assert!(
            !the_settle_is_fresh(&gone_to, band),
            "the nudge was made and did not fix it: the same box is not asked again"
        );
        assert!(
            the_settle_is_fresh(&gone_to, (11, 20, 640, 360)),
            "a box that changed is a new settle"
        );
        assert!(
            the_settle_is_fresh(
                &Thumbnail {
                    asked: band,
                    fill: Fill::Filled,
                    ..Thumbnail::NOTHING
                },
                band
            ),
            "and a picture that filled last time is worth one more look at the same box"
        );
    }

    /// The player's geometry is read once per settle, as one node: `osd-dimensions` carries the
    /// window, the margins and the picture in one answer, and a settle may spend one property read
    /// (see `the_drawn_picture`).
    #[test]
    fn the_drawn_picture_is_one_read_of_the_osd_dimensions_node() {
        let _one = super::super::pin_window::PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|held| held.into_inner());
        let pid = 0x5eed;
        mpv_ipc::remember_pipe(pid, "rhp-test-drawn-picture", false);
        mpv_ipc::clear_reads();

        // The answer mpv 0.41.0 gives a 1920x1080 film in a 640x300 window, verbatim.
        mpv_ipc::answer_a_read(
            pid,
            "osd-dimensions",
            Some(
                r#"{"w":640,"h":300,"par":1.000000,"aspect":2.133333,"mt":0,"mb":0,"ml":53,"mr":54}"#,
            ),
        );
        assert_eq!(
            the_drawn_picture(pid),
            Some(Drawn {
                window: (640, 300),
                picture: (533, 300),
                margins: (53, 0, 54, 0),
            }),
            "the window, the picture inside the letterbox and the letterbox itself come out of the \
             one answer"
        );
        assert_eq!(
            mpv_ipc::noted_reads(),
            vec![format!("{pid}:osd-dimensions")],
            "asked once, and for the node: a read per field is three more round trips on a settle"
        );

        // A player that cannot answer is a player whose picture is not judged.
        mpv_ipc::answer_a_read(pid, "osd-dimensions", None);
        assert_eq!(the_drawn_picture(pid), None);
        assert_eq!(the_drawn_picture(0), None, "a player this app does not hold is not asked");

        mpv_ipc::forget_pipe(pid);
        mpv_ipc::clear_reads();
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
        // A crop that would be wider than the picture it is cut from is clamped into it, and a trace
        // says so.
        let ((_, _, width, height), clamped) = the_crop_filling((0, 0, 4, 100), (4000, 1));
        assert_eq!((width, height), (4, 1), "never outside the picture");
        assert!(clamped, "and the clamp is what the trace reads");
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
        let drawn = Drawn {
            window: (3217, 1460),
            picture: (2595, 1460),
            margins: (311, 0, 311, 0),
        };
        let aspect = the_picture_shape(drawn);
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

    /// **The picture a settle reads is what a whole drag of steps is drawn with, and a step reads nothing
    /// of the player's**: one `osd-dimensions` read, and the same picture at every box the hand pulls the
    /// pin through — cropped to each box's shape, which is the box filled and never a bar (see
    /// `the_thumbnail_box`).
    #[test]
    fn a_drag_reuses_the_picture_its_settle_read() {
        let _one = super::super::pin_window::PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|held| held.into_inner());
        let pid = 0x5eed;
        mpv_ipc::remember_pipe(pid, "rhp-test-drag-picture", false);
        mpv_ipc::clear_reads();
        mpv_ipc::answer_a_read(
            pid,
            "osd-dimensions",
            Some(
                r#"{"w":3217,"h":1460,"par":1.000000,"aspect":1.777397,"mt":0,"mb":0,"ml":311,"mr":311}"#,
            ),
        );
        let read = the_drawn_picture(pid).expect("the player answers about its output");
        let picture = the_picture_shape(read);
        // The client the source's window answers for the step, never asked of the player: the same box
        // for the whole drag, because a hand on an edge does not resize the source (see
        // `the_source_client` and `the_band_was_placed`).
        let source = (3217, 1460);

        let boxes = [(680, 383), (1000, 560), (2263, 1274)];
        let shown: Vec<_> = boxes
            .iter()
            .map(|box_| the_thumbnail_box(*box_, source, picture, PictureRoad::Picture))
            .collect();

        assert_eq!(
            mpv_ipc::noted_reads(),
            vec![format!("{pid}:osd-dimensions")],
            "one read for the settle and nothing asked for by the steps"
        );
        let painted = a_picture_fitted_to(picture, source);
        for (step, box_) in shown.iter().zip(boxes) {
            assert_eq!(
                step.destination,
                (0, 0, box_.0, box_.1),
                "the box follows the hand and is filled at every step of it"
            );
            let (left, top, right, bottom) = step.source.expect("the picture is what is cropped by");
            assert!(
                left >= painted.0
                    && top >= painted.1
                    && right <= painted.0 + painted.2
                    && bottom <= painted.1 + painted.3,
                "and what is cropped is always inside the picture the settle read"
            );
        }

        mpv_ipc::forget_pipe(pid);
        mpv_ipc::clear_reads();
    }
}

//! The installed f3d a model is shown by: the launch, the wait for the window it makes, and the
//! frame that has to come off it.
//!
//! A model has no pixels of this app's — the scene is f3d's renderer's — so what a hover or a pin
//! shows is f3d's own window, on the roads mpv's and FFmpeg's players are already shown on:
//! `VIDEO_PID` and `VIDEO_HWND`, `engine_processes`' record of the player, the monitor that styles
//! a player's window before anything shows it, and the hover and pin hosts that draw a DWM
//! thumbnail of that window (see `start_hover_video_playback` and `pin_film_thumbnail`). Booking a
//! viewer as a player rather than beside one is the whole of what makes a hovered model killable,
//! pinnable, sweepable and promotable without a second road for any of the four.
//!
//! What is f3d's own is what no player of this app's has ever needed. **The launch is one long
//! command line** and its owner is `f3d_render` (see `launch_args`, where every flag is argued
//! for). **The window is stripped of its frame**: f3d makes a decorated top-level window and there
//! is no flag that says otherwise, so the title bar and the border come off here and the window's
//! box becomes its client — which is the rectangle a thumbnail's crop is cut from. And **the window
//! is waited for**, because a viewer that has made no window is a preview with nothing in it.
//!
//! **And a pinned model's window is adopted as a child of the pin's film host**: the scene stands inside
//! the band the pin paints around, so a hand on it is f3d's own — its orbit, its wheel, its menu, its binds —
//! while the pin's chrome stays the strip above it, and moving or resizing the pin carries the picture
//! without a call into f3d (see `adopt_the_model_window`). A hover keeps the Shell's thumbnail.

use windows::Win32::UI::WindowsAndMessaging::{
    SetParent, GWL_EXSTYLE, GWL_STYLE, SWP_ASYNCWINDOWPOS, SW_SHOW, WS_CAPTION, WS_CHILD,
    WS_EX_LAYERED, WS_EX_NOACTIVATE, WS_EX_TRANSPARENT, WS_MAXIMIZEBOX, WS_MINIMIZEBOX, WS_POPUP,
    WS_SYSMENU, WS_THICKFRAME,
};

use super::*;

/// How long f3d is given to make its window: a player's own start bound, which is what a cold disk
/// is given for a film (see `VIDEO_START_WAIT_SECS`). Measured on 4.0.1 here: visible at 182 ms and
/// drawing at 307 ms from a cold process.
#[cfg(not(test))]
const MODEL_START_WAIT: Duration = Duration::from_secs(VIDEO_START_WAIT_SECS);

/// How often the window is looked for while it is not there: the monitor's own start cadence, and
/// nothing of this app's is held while it is waited for (see `ensure_noactivate_monitor`).
#[cfg(not(test))]
const MODEL_WINDOW_POLL: Duration = Duration::from_millis(5);

/// Begin f3d on `path` at `box_`, where a hover or a pin is to draw it, and answer the process
/// every road downstream treats as a player's.
///
/// The box is the box the picture stands in — the hover's own, or the pin's band — and it is asked
/// of f3d twice: as the client size it renders at and as where its window goes. The two are the
/// same rectangle once the frame is off (see `strip_the_frame`), which is what keeps the model's
/// picture from being a crop of a window larger than the box it is shown in.
///
/// **Where the picture is drawn is asked here, before the window exists**, which is the same
/// question `start_mpv_video_playback` asks: a pin's model is a source its host thumbnails, and a
/// hover's is one when Pin Mode is on, so that pinning it is the same picture re-cropped rather
/// than a second f3d. The host itself is made by the road that keeps it — this only records which
/// process the host is to adopt (see `host_for_a_hosted_player`).
#[allow(clippy::too_many_arguments)] // The launch's facts, the same ones a player's launch is given.
pub(super) fn launch(
    path: &Path,
    box_: ScreenRegion,
    kind: Option<MediaType>,
    allow_the_hover_host: bool,
    for_a_hover: bool,
) -> Option<video_playback::PlayerProcess> {
    let program = f3d_render::program()?;
    let background = engine_render::current_background(crate::formats::routing::Backdrop::Image);

    let the_pins_film = pin_film_host::the_film_is_shown_in_a_host(path, kind);
    let hover_host = (!the_pins_film
        && allow_the_hover_host
        && hover_film_host::the_hover_film_is_kept_in_a_host(path))
    .then(|| unsafe { hover_film_host::take_up_the_hover_film_host(box_) })
    .flatten();
    let into_a_host = the_pins_film || hover_host.is_some();

    let mut command = engine_processes::hidden_command(program);
    command
        .args(f3d_render::launch_args(
            path,
            box_,
            &f3d_render::background_argument(background),
        ))
        .stdin(Stdio::null())
        .stdout(Stdio::null())
        .stderr(Stdio::null());

    let mut player = video_playback::spawn_player(&mut command)?;
    let pid = player.id();
    note_the_box(pid, box_);

    // **The window is styled before it is shown.** f3d opens the window it is going to draw in the
    // moment it has one, and a window shown without `WS_EX_NOACTIVATE` takes the foreground from the
    // file being renamed beside it — which is the one thing a hover may never do (see
    // `note_a_player_about_to_show` and `eager_ex_style`, the two bits a player's window never goes
    // on screen without). It is also what keeps the window out of the Shell's taskbar, which f3d's
    // own window would otherwise be in: it is unowned, not a tool window, and nothing of this app's
    // may put a button on the taskbar for a hover.
    ffplay_window::note_a_player_about_to_show(pid, for_a_hover);

    // **And the host holds this process and no older one**, recorded before the window exists: the
    // wait a hosted picture is revealed on is a wait for *this* process's source (see
    // `host_for_a_hosted_player`).
    if the_pins_film {
        pin_film_host::note_hosted_player(pid, Some(path));
    } else if hover_host.is_some() {
        hover_film_host::note_hover_film_host_player(pid, path);
    }

    video_playback::note_player_started(f3d_render::F3D_NAME, path, &player, box_, 0.0, false);

    match wait_for_the_window(pid, &mut player) {
        Some(window) => {
            unsafe { show_the_window(window, box_, into_a_host) };
            Some(player)
        }
        // A viewer that made no window is no preview: it is ended here rather than left standing
        // behind a hover that has nothing to show (see `take_down`).
        None => {
            // The test build's noted player is no process and no window is expected of one: what a
            // test reads back is the launch itself (see `spawn_player`).
            #[cfg(not(test))]
            {
                take_down(pid, &mut player);
                None
            }
            #[cfg(test)]
            Some(player)
        }
    }
}

/// Put a launched model's window on the screen where the box is, and say whether it is there.
///
/// Three things, in this order and for one reason each. **The frame comes off** before anything
/// else, so that no frame of a title bar is ever shown over the box. **The window is placed** at
/// the box, because `--position` named the decorated window's corner and the client has just moved
/// by a title bar's height. And **the no-activate road is the one that writes the bits and shows
/// it** (see `apply_noactivate_to_hwnd`), so a model's window is styled, raised and shown by the
/// same calls every player's window is.
///
/// A window a host is keeping is not shown: it is the source a DWM thumbnail is cut from, and the
/// Shell draws it into the host (see `pin_film_thumbnail::adopt_the_source`).
pub(super) unsafe fn show_the_window(window: HWND, box_: ScreenRegion, into_a_host: bool) {
    unsafe { strip_the_frame(window) };
    unsafe { place_the_window(window, box_) };
    if into_a_host {
        return;
    }
    // **The Shell lists a window as it is shown, so a button it has already listed is one the style
    // below cannot take away.** f3d shows its window itself, before this app has found it, so the
    // hide is what makes the show the style is read on — and the show is `SW_SHOWNOACTIVATE` on a
    // window that already carries `WS_EX_NOACTIVATE` (see `give_a_players_window_no_button`, which
    // pays the same price for the one other window of a player's that is shown before this app has
    // styled it).
    let _ = unsafe { ShowWindow(window, SW_HIDE) };
    unsafe { ffplay_window::apply_noactivate_to_hwnd(window) };
}

/// The box a model's picture is drawn at, asked of the launch: a model is the one preview whose
/// renderer draws at whatever size it is given, so the box it was begun at is the size its picture
/// is. It is what the thumbnail road is told, because the other two answers that road knows — a
/// player's own `video-params` and a probe's measurement of the film — are films' answers and a
/// scene has neither (see `pin_film_thumbnail::the_films_native_size`).
type ModelBox = (u32, (i32, i32));

static MODEL_BOX: Lazy<Mutex<Option<ModelBox>>> = Lazy::new(|| Mutex::new(None));

/// Note the box a model was begun at, beside the process it was begun as (see `MODEL_BOX`).
pub(super) fn note_the_box(pid: u32, box_: ScreenRegion) {
    if pid == 0 {
        return;
    }
    if let Ok(mut held) = MODEL_BOX.lock() {
        *held = Some((pid, ((box_.2 - box_.0).max(1), (box_.3 - box_.1).max(1))));
    }
}

/// The size a model this app began draws at — the box it was launched at — and nothing for any
/// other player (see `MODEL_BOX`).
pub(super) fn the_box_of_a_model(pid: u32) -> Option<(i32, i32)> {
    MODEL_BOX
        .lock()
        .ok()
        .and_then(|held| *held)
        .filter(|(begun, _)| *begun == pid)
        .map(|(_, box_)| box_)
}

/// End a model this road has begun and found no window for: the process, the record of it, and the
/// id every window of it would have been found by (see `clear_video_process_state`).
#[cfg(not(test))]
fn take_down(pid: u32, player: &mut video_playback::PlayerProcess) {
    let _ = player.kill();
    engine_processes::forget(pid);
    ffplay_window::clear_video_process_state(pid);
}

/// The window f3d has made for itself, waited for under the process it was begun as.
///
/// **f3d makes more windows than the one it means, and the first is not the picture.** Measured on
/// 4.0.1 here: two transient `vtkOpenGL` windows exist while it starts — one of no height at all
/// and one 300x300 — before the window it was asked for. So the pick is the largest visible window
/// of the process and it is looked for again on every pass rather than latched (see
/// `a_players_window_now`), and a process that has ended is a window that is never coming.
fn wait_for_the_window(pid: u32, player: &mut video_playback::PlayerProcess) -> Option<HWND> {
    // The test build begins no process, so there is no window to wait for and nothing to strip:
    // what a test reads back is the launch itself (see `spawn_player`).
    #[cfg(test)]
    {
        let _ = (pid, player);
        None
    }

    #[cfg(not(test))]
    {
        let deadline = Instant::now() + MODEL_START_WAIT;
        loop {
            if let Some(window) = unsafe { ffplay_window::a_players_window_now(pid) } {
                return Some(window);
            }
            if matches!(player.try_wait(), Ok(Some(_))) {
                return None;
            }
            if Instant::now() >= deadline {
                return None;
            }
            std::thread::sleep(MODEL_WINDOW_POLL);
        }
    }
}

/// Take the frame off a decorated window: no title bar, no sizing border, no system menu.
///
/// f3d 4.0.1 has no flag for this and no `ui.decorations` option — checked against `--help` and by
/// trying them — so the frame comes off the window after it exists. What it buys is that the box a
/// window is placed at and the client its picture is drawn into are one rectangle, which is what
/// the thumbnail road crops (see `pin_film_thumbnail`) and what a chrome-less preview is.
fn chrome_less_style(style: isize) -> isize {
    let frame = WS_CAPTION.0 as isize
        | WS_THICKFRAME.0 as isize
        | WS_MINIMIZEBOX.0 as isize
        | WS_MAXIMIZEBOX.0 as isize
        | WS_SYSMENU.0 as isize;

    (style & !frame) | WS_POPUP.0 as isize
}

/// Whether a style already wears neither the frame nor the popup bit that replaced it.
fn is_chrome_less(style: isize) -> bool {
    chrome_less_style(style) == style
}

/// Strip the frame from a viewer's window, where it still has one.
///
/// Safety: `window` is a window f3d made, and both calls only change its style.
pub(super) unsafe fn strip_the_frame(window: HWND) {
    let style = unsafe { GetWindowLongPtrW(window, GWL_STYLE) };
    if is_chrome_less(style) {
        return;
    }
    let _ = unsafe { SetWindowLongPtrW(window, GWL_STYLE, chrome_less_style(style)) };
    // A style written behind a window's back is not one it presents with until it is asked to read
    // it, and the frame's height comes off the window with the same call.
    let _ = unsafe {
        SetWindowPos(
            window,
            HWND::default(),
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
        )
    };
}

/// Put a viewer's window at the box, so that its client is the box exactly.
///
/// Safety: as above — the window is f3d's, and this is a placement and nothing else.
unsafe fn place_the_window(window: HWND, box_: ScreenRegion) {
    let (left, top, right, bottom) = box_;
    let _ = unsafe {
        SetWindowPos(
            window,
            HWND::default(),
            left,
            top,
            (right - left).max(1),
            (bottom - top).max(1),
            SWP_NOZORDER | SWP_NOACTIVATE,
        )
    };
}

/// The engine window a model's band stands in, where the picture standing in that band is a model's: nothing
/// for a film's player, whose window is no scene this app turns.
///
/// The kind is asked of the media on screen rather than of a flag of this module's, because that is what
/// says what the window is *for*: a player retired behind a new one, or a film begun after a model, is a
/// window whose pid and handle say nothing about what it draws.
///
/// **A pinned model's window is the child the pin's host holds** and a hover's is the one the monitor
/// published: a player begun into a host publishes nothing of the desktop's, so the published handle is
/// nothing on every pinned model (see `adopt_the_model_window` and `pin_film_host::note_hosted_player`).
pub(super) fn the_models_window() -> Option<HWND> {
    a_models_window(
        current_media_type_try(),
        adopted_window()
            .map(|adopted| adopted.source)
            .unwrap_or_else(|| VIDEO_HWND.load(Ordering::SeqCst)),
    )
}

/// Whether a pinned model's picture is still on the screen.
///
/// **f3d's window is the picture, and f3d can close it**: its own Esc or its own quit bind is a hand on the
/// scene rather than on the pin, so a handle written down when the window was there is a handle to nothing
/// afterwards. The pin is a window onto that scene, and a scene that has gone is a pin that has come apart
/// like any other (see `pin_media_is_alive`).
pub(super) fn the_model_window_is_up() -> bool {
    the_models_window().is_some_and(|window| unsafe { IsWindow(window) }.as_bool())
}

/// The rule `the_models_window` decides by, apart from the machine it reads: a model is the media on
/// screen and its player has a window. A model mid-launch is a band with nothing to turn in it yet.
fn a_models_window(kind: Option<MediaType>, window: isize) -> Option<HWND> {
    (a_model_takes_the_pointer(kind) && window != 0).then_some(HWND(window as *mut _))
}

/// **Which media types are turned by the pointer**, which is the one kind that is a window of somebody
/// else's: a scene. Every other kind either draws into this app's own window, whose presses this app
/// answers itself, or is shown through the same host but is a film — a film's band is a hand on the pin
/// (see `pin_film_host::answer_a_host`).
fn a_model_takes_the_pointer(kind: Option<MediaType>) -> bool {
    kind == Some(MediaType::Model)
}

/// The model window a pin adopted into its host, and what it wore before it was adopted.
///
/// **One at a time, and written down so it can be given back**: the styles are the window's own — f3d's
/// decorated top-level one, or the invisible layered one a hover's host was drawing — and a child of the
/// host is destroyed with it, so the restore is the first thing every road out of a pin does (see
/// `restore_the_adopted_model_window`).
#[derive(Clone, Copy)]
struct Adopted {
    process: u32,
    host: isize,
    source: isize,
    style: isize,
    ex_style: isize,
}

static ADOPTED: Lazy<Mutex<Option<Adopted>>> = Lazy::new(|| Mutex::new(None));

/// The model window a pin has adopted, while it is still a window: f3d closes its own window on its own
/// bind, and a handle written down when it was there is a handle to nothing afterwards.
fn adopted_window() -> Option<Adopted> {
    let adopted = ADOPTED.lock().ok().and_then(|held| *held)?;
    unsafe { IsWindow(HWND(adopted.source as *mut _)) }
        .as_bool()
        .then_some(adopted)
}

/// Whether the window this pin adopted is `pid`'s and still up — the question the monitor asks before it
/// goes looking for one to adopt (see `ffplay_window::try_apply_noactivate_style`).
pub(super) fn the_adopted_model_is_up(pid: u32) -> bool {
    adopted_window().is_some_and(|adopted| adopted.process == pid)
}

/// Whether `host` holds a model's window as a child of it, which is what makes the band's own press the
/// child's rather than the pin's (see `pin_film_host::answer_a_host`).
pub(super) fn a_model_child_stands_in(host: HWND) -> bool {
    adopted_window().is_some_and(|adopted| adopted.host == host.0 as isize)
}

/// Whether the picture the pin is showing is a model's and `host` is the host it is drawn in: the one
/// question that tells a model's adoption from a film's thumbnail (see `adopt_the_model_window`).
pub(super) fn a_pinned_model_holds_the_host(host: HWND) -> bool {
    pinned()
        && pin_film_host::film_host() == Some(host)
        && a_model_takes_the_pointer(current_media_type_try())
}

/// How far inside the pin's own band a model's window is placed.
///
/// **Nothing the pin paints may clip the scene.** The band is framed by the pin's edge and its corner, and a
/// scene drawn on the band's own rectangle paints over both.
///
/// shortcut: the pinned band's paint is a hairline with no radius today, so one pixel a side; upgrade to the
/// painted radius when the band's paint gains one.
const MODEL_BAND_INSET: i32 = 1;

/// The box an adopted model's window is given in its host: the host's client, inset by the pin's frame, in
/// the parent's own coordinates — which is why a move of the pin needs no call at all and a resize needs
/// this one (see `fit_the_adopted_model_window`).
fn the_adopted_box(client: (i32, i32)) -> (i32, i32, i32, i32) {
    let (left, top, right, bottom) =
        the_models_band_inset((0, 0, client.0.max(1), client.1.max(1)));
    (left, top, right - left, bottom - top)
}

/// Adopt a launched model's window into the pin's host as a child of it: **the band is the scene's, and the
/// scene stands inside the pin**.
///
/// The frame comes off, `WS_CHILD` goes on and the desktop's `WS_POPUP` off, and the thumbnail's
/// invisibility and press-through come off with them: a layered or transparent child is a scene nothing can
/// hit, and this window is the picture the hand is on. The host clips it and answers the hit test for it,
/// the pin's chrome stands above the host as it does for a film, and a move of the pin carries the child
/// because a child's box is its parent's (see `pin_film_host::answer_a_host` and
/// `pin_lifecycle::place_pinned_siblings`).
///
/// **The Shell's thumbnail of the same window is let go first**: the picture is inside the host now, and a
/// registration left standing would draw the source over it (see `pin_film_thumbnail`).
///
/// Safety: `source` is a window the caller found by `pid` and this app did not make.
pub(super) unsafe fn adopt_the_model_window(source: HWND, pid: u32, host: HWND) {
    unsafe { restore_the_adopted_model_window() };
    pin_film_thumbnail::forget_the_thumbnail_because("the model is the band's own child");

    let style = unsafe { GetWindowLongPtrW(source, GWL_STYLE) };
    let ex_style = unsafe { GetWindowLongPtrW(source, GWL_EXSTYLE) };
    if let Ok(mut held) = ADOPTED.lock() {
        *held = Some(Adopted {
            process: pid,
            host: host.0 as isize,
            source: source.0 as isize,
            style,
            ex_style,
        });
    }

    let (child, child_ex) = the_adopted_styles(style, ex_style);
    let _ = unsafe { SetWindowLongPtrW(source, GWL_STYLE, child) };
    let _ = unsafe { SetWindowLongPtrW(source, GWL_EXSTYLE, child_ex) };
    // A style written behind a window's back is not one it presents with until it is asked to read it.
    let _ = unsafe {
        SetWindowPos(
            source,
            HWND::default(),
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
        )
    };

    let _ = unsafe { SetParent(source, host) };
    fit_the_adopted_model_window(host);
    let _ = unsafe { ShowWindow(source, SW_SHOW) };
}

/// The two styles an adopted model's window wears: **the frame off, `WS_CHILD` in place of the popup, and
/// nothing of a thumbnail source left on it**. A layered child is composed rather than drawn into its parent
/// and a transparent one is skipped by the pointer, and this window is the picture the hand is on (see
/// `adopt_the_model_window`).
fn the_adopted_styles(style: isize, ex_style: isize) -> (isize, isize) {
    let child = (chrome_less_style(style) & !(WS_POPUP.0 as isize)) | WS_CHILD.0 as isize;
    let child_ex = ex_style
        & !(WS_EX_LAYERED.0 as isize)
        & !(WS_EX_TRANSPARENT.0 as isize)
        & !(WS_EX_NOACTIVATE.0 as isize);
    (child, child_ex)
}

/// Put the adopted window at the host's inner band, which is the one thing a resize owes it: the child's
/// box is its parent's, so a move of the pin is already the picture's and a resize is this re-measure.
///
/// shortcut: one posted placement per band step, never a synchronous one — a cross-process
/// `SetWindowPos` made to wait is the drag blocked on f3d's thread; upgrade to a latched box when a
/// resize is measured to lag the frame.
pub(super) fn fit_the_adopted_model_window(host: HWND) {
    let Some(adopted) = adopted_window() else {
        return;
    };
    let Some(client) = (unsafe { pin_film_thumbnail::client_size(host) }) else {
        return;
    };
    let (x, y, width, height) = the_adopted_box(client);
    let _ = unsafe {
        SetWindowPos(
            HWND(adopted.source as *mut _),
            HWND::default(),
            x,
            y,
            width,
            height,
            SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS,
        )
    };
}

/// Give an adopted model's window back: unparented first, then the styles it wore before the pin took it.
///
/// **Asked before anything else touches it.** A child destroyed with the host is a player killed under its
/// own window, and a window re-used by a later hover or pin is one that must not still be a child of a host
/// that has gone (see `pin_film_host::take_down_the_pin_film_host`).
pub(super) unsafe fn restore_the_adopted_model_window() {
    let Some(adopted) = ADOPTED.lock().ok().and_then(|mut held| held.take()) else {
        return;
    };
    let source = HWND(adopted.source as *mut _);
    if !unsafe { IsWindow(source) }.as_bool() {
        return;
    }
    let _ = unsafe { SetParent(source, None) };
    let _ = unsafe { SetWindowLongPtrW(source, GWL_STYLE, adopted.style) };
    let _ = unsafe { SetWindowLongPtrW(source, GWL_EXSTYLE, adopted.ex_style) };
    let _ = unsafe {
        SetWindowPos(
            source,
            HWND::default(),
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED,
        )
    };
}

/// Whether the band the pin is painting is a model's scene, which is the one band a press in it is not the
/// pin's: f3d's window is a child of the pin's host and takes it itself, so the hairline inside the host's
/// client is answered and carries nothing (see `adopt_the_model_window`).
pub(super) fn the_band_is_a_models_window() -> bool {
    a_model_takes_the_pointer(current_media_type_try())
}

/// The band a model's window is inset from, as a box: the pin's picture area, one pixel in from its frame
/// on every side and never past a box of one pixel on either axis (see `MODEL_BAND_INSET`).
fn the_models_band_inset(content: ScreenRegion) -> ScreenRegion {
    let (left, top, right, bottom) = content;
    let left = (left + MODEL_BAND_INSET).min(right - 1);
    let top = (top + MODEL_BAND_INSET).min(bottom - 1);
    (
        left,
        top,
        (right - MODEL_BAND_INSET).max(left + 1),
        (bottom - MODEL_BAND_INSET).max(top + 1),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    /// **A model's window is the band's own box, inset by the pin's frame**, in the host's client: the child
    /// is placed once for a band and again for a resize, and both read this one answer (see `the_adopted_box`).
    #[test]
    fn the_adopted_box_is_the_hosts_client_inside_the_pins_frame() {
        assert_eq!(the_adopted_box((400, 300)), (1, 1, 398, 298));

        // A band too small to be inset is a box of one pixel rather than an inverted one: a placement asked
        // for a negative width is no placement at all.
        assert_eq!(the_adopted_box((2, 2)), (1, 1, 1, 1));
        assert_eq!(the_adopted_box((1, 1)), (0, 0, 1, 1));
    }

    /// no window of a model's is one until it has been made.
    #[test]
    fn a_model_takes_the_pointer_and_no_other_kind_does() {
        for kind in [
            MediaType::Video,
            MediaType::NativeVideo,
            MediaType::StaticImage,
        ] {
            assert!(
                !a_model_takes_the_pointer(Some(kind)),
                "{kind:?} is no scene"
            );
            assert_eq!(
                a_models_window(Some(kind), 0x1234),
                None,
                "and no band of it is"
            );
        }
        assert!(!a_model_takes_the_pointer(None));

        assert!(a_model_takes_the_pointer(Some(MediaType::Model)));
        assert_eq!(
            a_models_window(Some(MediaType::Model), 0x1234),
            Some(HWND(0x1234 as *mut _))
        );
        assert_eq!(
            a_models_window(Some(MediaType::Model), 0),
            None,
            "a model mid-launch is a band with nothing to turn in it yet"
        );
    }

    /// **A model's window is its client**, and the bits that make that true are the frame's: every
    /// one of them goes, the popup bit the window ends up as comes on, and a window that already
    /// wears the answer is left alone so that the strip can be asked for on any pass.
    #[test]
    fn the_strip_takes_the_frame_and_leaves_the_window_alone_when_it_has_none() {
        let decorated =
            (WS_CAPTION.0 | WS_THICKFRAME.0 | WS_MINIMIZEBOX.0 | WS_MAXIMIZEBOX.0 | WS_SYSMENU.0)
                as isize;
        let bare = chrome_less_style(decorated);

        assert_eq!(bare & decorated, 0, "no bit of the frame is left on it");
        assert_eq!(
            bare & WS_POPUP.0 as isize,
            WS_POPUP.0 as isize,
            "and the window is a popup, which is what a chrome-less one is"
        );
        assert!(is_chrome_less(bare), "and a second strip is no change");
        assert!(!is_chrome_less(decorated));
    }
    /// **A band stands in the pin's own frame**, and a scene drawn on the band's own rectangle would paint
    /// over the edge and the corner the pin keeps (see `MODEL_BAND_INSET`).
    #[test]
    fn a_handed_overs_band_is_inset_by_the_pins_frame() {
        assert_eq!(
            the_models_band_inset((100, 50, 300, 250)),
            (101, 51, 299, 249)
        );

        // A band too small to be inset is a box of one pixel rather than an inverted one: a placement that
        // asks for a negative width is no placement at all.
        assert_eq!(the_models_band_inset((10, 10, 12, 12)), (11, 11, 12, 12));
        assert_eq!(the_models_band_inset((10, 10, 11, 11)), (10, 10, 11, 11));
    }

    /// **The styles an adopted model wears while the pin holds it, and the ones it is given back**: a
    /// `WS_CHILD` window of the host's, with the frame's bits and the thumbnail's two off, and the window it
    /// was before — a desktop popup — once the pin is done with it (see `adopt_the_model_window` and
    /// `restore_the_adopted_model_window`).
    #[test]
    fn an_adopted_model_is_a_child_while_the_pin_holds_it_and_a_popup_after() {
        let decorated =
            (WS_CAPTION.0 | WS_THICKFRAME.0 | WS_MINIMIZEBOX.0 | WS_MAXIMIZEBOX.0 | WS_SYSMENU.0)
                as isize;
        let layered =
            WS_EX_LAYERED.0 as isize | WS_EX_TRANSPARENT.0 as isize | WS_EX_NOACTIVATE.0 as isize;

        let (child, child_ex) = the_adopted_styles(decorated, layered);
        assert_eq!(child & WS_CHILD.0 as isize, WS_CHILD.0 as isize);
        assert_eq!(child & WS_POPUP.0 as isize, 0, "a child is no popup");
        assert_eq!(child & decorated, 0, "and the frame is off it");
        assert_eq!(child_ex & layered, 0, "and the picture is the hand's");

        // The same bits the strip leaves a model's window wearing: the child is that, one bit over.
        assert_eq!(
            child,
            (chrome_less_style(decorated) & !(WS_POPUP.0 as isize)) | WS_CHILD.0 as isize
        );
    }
}

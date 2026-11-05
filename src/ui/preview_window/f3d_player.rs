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
//! **And a pinned model's band is handed over to that window on the first press in it**: the pin draws the
//! Shell's thumbnail of it, and the press that would carry the pin makes the engine's window the window the
//! hand is on instead — parked in the band, raised over the pin and made hit-testable, with a real press
//! injected at the cursor, because a scene turns from its own captured drag and never from a posted message
//! (see `hand_the_band_over`).

use windows::Win32::UI::Input::KeyboardAndMouse::{
    SendInput, INPUT, INPUT_0, INPUT_MOUSE, MOUSEEVENTF_LEFTDOWN, MOUSEINPUT,
};
use windows::Win32::UI::WindowsAndMessaging::{
    GetCursorPos, GWL_EXSTYLE, GWL_STYLE, WS_CAPTION, WS_EX_LAYERED, WS_EX_NOACTIVATE,
    WS_EX_TOOLWINDOW, WS_EX_TRANSPARENT, WS_MAXIMIZEBOX, WS_MINIMIZEBOX, WS_SYSMENU, WS_THICKFRAME,
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
/// model's: nothing for a film's player, whose window is no scene this app turns.
///
/// The kind is asked of the media on screen rather than of a flag of this module's, because that is what
/// says what the window is *for*: a player retired behind a new one, or a film begun after a model, is a
/// window whose pid and handle say nothing about what it draws.
pub(super) fn the_models_window() -> Option<HWND> {
    a_models_window(
        current_media_type_try(),
        a_models_source(
            VIDEO_HWND.load(Ordering::SeqCst),
            pinned(),
            pin_film_thumbnail::the_source_window().map(|source| source.0 as isize),
        ),
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

/// The window a model's band is turned in: **while a pin is up it is the source that pin's host
/// thumbnails**, because a player begun into a host publishes nothing of the desktop's and the host
/// stands over the source's own rectangle — so the published handle is nothing on every pinned model,
/// and a press read off it would be a band with nothing to turn (see `pin_film_host::note_hosted_player`
/// and `pin_film_thumbnail::the_source_window`). A hover keeps the published handle, which is the
/// window the monitor found for it and the only one a hover host is not standing over.
fn a_models_source(published: isize, pinned: bool, source: Option<isize>) -> isize {
    if pinned {
        return source.unwrap_or(published);
    }
    published
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

/// The window the band was handed over to, while the gesture is the engine's — zero otherwise.
static HANDED_OVER: AtomicIsize = AtomicIsize::new(0);

/// How far inside the pin's own band a handed-over model's window is placed.
///
/// **Nothing the pin paints may clip the scene.** The band is framed by the pin's edge and its corner, and a
/// scene drawn on the band's own rectangle paints over both.
///
/// shortcut: the pinned band's paint is a hairline with no radius today, so one pixel a side; upgrade to the
/// painted radius when the band's paint gains one.
const MODEL_BAND_INSET: i32 = 1;

/// Hand the band over to the engine on the first left press in it: f3d's own window takes the hand, the
/// gesture, and everything the gesture means from here.
///
/// **A press posted into the engine is a press it never reads** — f3d turns its scene from its own captured
/// drag — so the press that hands over is a real one, injected at the cursor once the window stands above the
/// pin, hit-testable and able to activate (see `inject_a_left_press`). A window that cannot activate never gets
/// its own keyboard binds or its own Esc either, so `WS_EX_NOACTIVATE` comes off for the length of this state.
///
/// Answer whether the band under the press is a model's at all: a press it is on a model whose window is not up
/// yet carries nothing, rather than carrying the window its picture is shown in.
///
/// Safety: asked from the pin's own window procedure's press road, with that window.
pub(super) unsafe fn hand_the_band_over(hwnd: HWND) -> bool {
    if !a_model_takes_the_pointer(current_media_type_try()) {
        return false;
    }
    let Some(source) = the_models_window() else {
        return true;
    };

    // **The latch goes down before anything moves**: the window takes the press the moment it is on screen, and
    // every road that asks whether the pair is still the interaction's answers from the latch rather than from
    // where the window happens to stand (see `settle_the_band_hand_over`).
    HANDED_OVER.store(source.0 as isize, Ordering::SeqCst);

    // Above the pin's own window and below its chrome, which stands above the pin — a scene over the chrome
    // would put the caption's own buttons out of the hand (see `pin_film_host::place_the_pin_in_the_band`).
    if let Some(band) = the_pins_band() {
        unsafe { place_the_window(source, the_models_band_inset(band)) };
    }
    let _ = unsafe {
        SetWindowPos(
            source,
            hwnd,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        )
    };
    let style = unsafe { GetWindowLongPtrW(source, GWL_EXSTYLE) };
    write_the_ex_style(source, the_handed_over_ex_style(style));

    // **Our own capture is let go of before the press is injected**: a window holding the pointer is given every
    // move of the drag, and the engine would be turning a scene nothing was sent.
    if unsafe { GetCapture() } == hwnd {
        let _ = unsafe { ReleaseCapture() };
    }

    unsafe { inject_a_left_press() };
    true
}

/// Whether the band is the engine's for now, and nothing of ours claims, forwards or hit-tests it (see
/// `hand_the_band_over`).
pub(super) fn the_band_is_handed_over() -> bool {
    HANDED_OVER.load(Ordering::SeqCst) != 0
}

/// Whether `window` is the source this app handed the pointer to, which is one of ours for as long as the
/// hand-over lasts: it stands in the pin's band, over the pin's own window.
pub(super) fn the_handed_over_source_is(window: HWND) -> bool {
    let held = HANDED_OVER.load(Ordering::SeqCst);
    held != 0 && !window.is_invalid() && window.0 as isize == held
}

/// Put the band back the way the mirror leaves it: the engine's window pointer-transparent and layered again,
/// standing where the pin's host thumbnails it from, and the latch gone.
///
/// **Asked from every road that says the gesture is over**: the pair no longer the interaction's, a press on the
/// pin's own chrome, the pin put away, and the engine's window gone (see `settle_the_band_hand_over` and
/// `pin_input::pinned_press_on`).
pub(super) fn return_the_band_to_the_mirror() {
    let held = HANDED_OVER.swap(0, Ordering::SeqCst);
    if held == 0 {
        return;
    }
    let source = HWND(held as *mut _);
    if !unsafe { IsWindow(source) }.as_bool() {
        return;
    }

    // **The mirror's own road puts the source back out of sight**: the style and alpha that make it a window
    // the host thumbnails rather than one the hand is on, put on by the road that keeps it that way every tick
    // (see `pin_film_thumbnail::keep_the_source_invisible`). Nothing is asked of the player.
    unsafe { pin_film_thumbnail::keep_the_source_invisible() };

    // **Rewound out of the pointer's way**: the box the host thumbnails from is the band's, and the window goes
    // back to standing where a film's own player stands — under the host that draws it (see
    // `pin_lifecycle::place_pinned_siblings`).
    if pinned() {
        if let Some(band) = the_pins_band() {
            pin_lifecycle::ensure_pinned_sibling_box(band);
        }
    }
}

/// Whether the hand-over is over, decided apart from the desktop it reads: the pair is the interaction's while
/// the window in front is the scene's own or one of the pin's, and nothing else (see
/// `settle_the_band_hand_over`).
fn the_hand_over_is_over(pinned: bool, source_up: bool, in_front_of_the_pair: bool) -> bool {
    !pinned || !source_up || !in_front_of_the_pair
}

/// Settle the hand-over against the desktop: back to the mirror the moment the pair is not what the user is in,
/// or the engine's window is gone.
///
/// **The window in front is the whole question.** The press that handed over makes the engine's window the one
/// the user is in, and the pin's own chrome takes it back; anything else in front — the listing, another
/// program, the desktop — is the gesture over, and the band goes back to being the pin's.
pub(super) fn settle_the_band_hand_over(foreground: HWND) {
    let held = HANDED_OVER.load(Ordering::SeqCst);
    if held == 0 {
        return;
    }
    let source = HWND(held as *mut _);
    let in_front_of_the_pair = foreground == source || pin_lifecycle::a_window_of_ours(foreground);
    if the_hand_over_is_over(
        pinned(),
        unsafe { IsWindow(source) }.as_bool(),
        in_front_of_the_pair,
    ) {
        return_the_band_to_the_mirror();
    }
}

/// The box a handed-over model's window stands in, from the band the pin says its picture occupies.
///
/// Pure so the inset can be asked without a window (see `MODEL_BAND_INSET`): every side moved in by the same
/// count, and never past a box of one pixel on either axis.
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

/// The pin's own band, read from the pin's state without waiting on it (see
/// `pin_model::pin_zoom_keep_alive_try`, whose rules this follows). Nothing where no pin is readable, which
/// leaves the scene where it stands rather than placing it somewhere worse.
fn the_pins_band() -> Option<ScreenRegion> {
    pin_window::pin_state_try().and_then(|pinned| pinned.pin().map(|pin| pin.content))
}

/// The style the engine's window wears while the band is handed over: **the scene is the picture and the hand is
/// on it**, so the two bits a thumbnail source wears come off — a `WS_EX_TRANSPARENT` window is skipped by the
/// pointer and a `WS_EX_LAYERED` one is composed rather than hit — and `WS_EX_NOACTIVATE` with them, which is
/// what lets a press of the user's own activate the engine and give it its keyboard (see
/// `ffplay_window::player_ex_style_for` for the style it wears at rest).
fn the_handed_over_ex_style(source: isize) -> isize {
    (source
        & !(WS_EX_TRANSPARENT.0 as isize)
        & !(WS_EX_LAYERED.0 as isize)
        & !(WS_EX_NOACTIVATE.0 as isize))
        | WS_EX_TOOLWINDOW.0 as isize
}

/// Write an extended style and make the window read it: a style is not a thing a window is told about live (see
/// `ffplay_window::place_player_window`).
fn write_the_ex_style(window: HWND, style: isize) {
    if style == 0 || unsafe { SetWindowLongPtrW(window, GWL_EXSTYLE, style) } == 0 {
        return;
    }
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

/// Press the left button for the engine, at the cursor: the press the hand has just made is the engine's from
/// here — its own capture, its own orbit, its own binds.
///
/// **Nothing is injected where the cursor cannot be read**, and none of it re-enters this app: the window has
/// already been placed, raised and made hit-testable by the caller, so the press lands in the scene; the app's
/// own low-level mouse hook is the one place a press of its own could come back in, and it reads no injected
/// event (see `crate::shell::wheel_input`).
unsafe fn inject_a_left_press() {
    let mut point = POINT::default();
    if unsafe { GetCursorPos(&mut point) }.is_err() {
        return;
    }
    let press = INPUT {
        r#type: INPUT_MOUSE,
        Anonymous: INPUT_0 {
            mi: MOUSEINPUT {
                dx: 0,
                dy: 0,
                mouseData: 0,
                dwFlags: MOUSEEVENTF_LEFTDOWN,
                time: 0,
                dwExtraInfo: 0,
            },
        },
    };
    let _ = unsafe { SendInput(&[press], std::mem::size_of::<INPUT>() as i32) };
}

#[cfg(test)]
mod tests {
    use super::*;

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

    /// **A pinned model's band is turned in the source its host is drawing**, which is the one handle a
    /// player begun into a host never publishes: `VIDEO_HWND` is nothing for such a player, and a band
    /// read off that handle alone is a band with nothing to turn (see `the_models_window`).
    #[test]
    fn a_pinned_models_band_is_turned_in_the_source_its_host_draws() {
        assert_eq!(
            a_models_source(0, true, Some(0x1234)),
            0x1234,
            "the pin's own source stands where a hosted player published no window"
        );
        assert_eq!(
            a_models_source(0x2000, true, Some(0x1234)),
            0x1234,
            "and it is the picture on the band, not an older player's window"
        );
        assert_eq!(
            a_models_source(0, true, None),
            0,
            "a pin whose picture has not been adopted yet has nothing to turn"
        );
        assert_eq!(
            a_models_source(0x2000, false, Some(0x1234)),
            0x2000,
            "a hover keeps the window the monitor found for it"
        );
        assert_eq!(
            a_models_source(0, false, Some(0x1234)),
            0,
            "and a hover's own host is nobody's to turn"
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

    /// **The hand-over's two bits come off and the mirror's go back on**: a scene the pointer cannot hit is a
    /// scene nothing turns, and a film's source is exactly that (see `the_handed_over_ex_style` and
    /// `ffplay_window::player_ex_style_for`, which is the style the mirror leaves it wearing).
    #[test]
    fn the_engine_window_is_hit_and_can_activate_only_while_handed_over() {
        let source =
            WS_EX_TRANSPARENT.0 as isize | WS_EX_LAYERED.0 as isize | WS_EX_NOACTIVATE.0 as isize;

        let handed_over = the_handed_over_ex_style(source);
        assert_eq!(handed_over & WS_EX_TRANSPARENT.0 as isize, 0);
        assert_eq!(handed_over & WS_EX_LAYERED.0 as isize, 0);
        assert_eq!(handed_over & WS_EX_NOACTIVATE.0 as isize, 0);
        assert!(
            handed_over & WS_EX_TOOLWINDOW.0 as isize != 0,
            "and it is still a window the Shell lists no button for"
        );

        let mirror = pin_film_thumbnail::source_ex_style(handed_over, true);
        assert!(
            mirror & WS_EX_TRANSPARENT.0 as isize != 0
                && mirror & WS_EX_LAYERED.0 as isize != 0
                && mirror & WS_EX_NOACTIVATE.0 as isize != 0,
            "which is the picture a thumbnail is cut from, and the band a hand carries"
        );
    }

    /// **The hand-over is over the moment the pair is not what the user is in**, and not before: the scene's
    /// window gone, the pin gone, or a window in front that is neither the scene's nor the pin's (see
    /// `settle_the_band_hand_over`).
    #[test]
    fn the_hand_over_ends_on_every_road_but_the_pairs_own_front() {
        assert!(
            !the_hand_over_is_over(true, true, true),
            "the scene is in front"
        );
        assert!(
            the_hand_over_is_over(true, true, false),
            "another program is"
        );
        assert!(
            the_hand_over_is_over(true, false, true),
            "the engine's window is gone"
        );
        assert!(the_hand_over_is_over(false, true, true), "the pin is gone");
    }
}

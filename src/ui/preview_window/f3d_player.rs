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
//! **And the pointer is forwarded**, because a pinned model's band is not the model's own window: what a
//! hand is on there is this app's window, which thumbnails the engine's window through the pin's host (see
//! `answer_a_models_band`). The press, the drag, the release and the notch are posted on to f3d in f3d's
//! own coordinates — posted and never sent, because a sent message is this app's thread waiting on the one
//! drawing the scene.

use windows::Win32::Graphics::Gdi::ScreenToClient;
use windows::Win32::UI::Input::KeyboardAndMouse::VK_LBUTTON;
use windows::Win32::UI::WindowsAndMessaging::{
    GetClientRect, GWL_STYLE, HTCLIENT, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MOUSEMOVE, WM_MOUSEWHEEL,
    WM_NCHITTEST, WS_CAPTION, WS_MAXIMIZEBOX, WS_MINIMIZEBOX, WS_SYSMENU, WS_THICKFRAME,
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

/// The engine window a band's press and notch belong to, where the picture standing in that band is a
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

/// The window of the pair that took the down on a model's band, and nothing while no drag is on —
/// **the one that took the capture and posts the rest**: a move reaches the engine from the window the
/// press landed on and from no other of the pair's.
///
/// **Held because the release is the engine's as often as it is ours**: f3d takes the capture on the down
/// it is posted, and the capture owner is given the real messages for the rest of the drag — so this window
/// hears the up only when it kept the pointer, and a latch that waited for one would be a drag that never
/// ends: the level of the button ends it, from a move and from the loop's own tick (see
/// `settle_a_models_drag`).
static MODEL_DRAG: AtomicIsize = AtomicIsize::new(0);

/// `MK_LBUTTON`, the key-state word a press and a move of a drag carry (see `post_a_band_message`).
const MK_LBUTTON: usize = 0x0001;

/// Answer a band standing in a model, in place of the pin: the scene's own pointer, which is the press, the
/// drag and the release of it. `None` for every other message, and for every band that is not a model's.
///
/// **The band's hit test is answered here too, and it is the whole of why a press in it turns the scene**
/// rather than carrying the window: a film's band is handed down to the pin with `HTTRANSPARENT`, which is
/// what makes a hand on a film move the pin — and the same answer for a model would carry a window the hand
/// meant to turn (see `pin_film_host::answer_a_host`).
///
/// **Nothing is activated for any of it** (see `show_the_window`): f3d is given messages, no focus, and the
/// press is answered here rather than by the pin so that the caret the pin takes on a press of its own is
/// the one thing a band press does not ask for.
///
/// **The pair's two windows ask it on two roads**: the host asks it for anything landing on its own
/// client, and the pin's own window asks `press_the_models_band` once the chrome's own arms above it have
/// had the press (see `pin_input::pinned_press_on`). Whichever window took the down is the one that posts
/// the rest of the drag, and the loop ends what neither is sent (see `settle_a_models_drag`).
///
/// Safety: the answers are those of a window of this app's own, asked from its own window procedure.
pub(super) unsafe fn answer_a_models_band(
    host: HWND,
    msg: u32,
    _wparam: WPARAM,
    lparam: LPARAM,
) -> Option<LRESULT> {
    let source = the_models_window()?;

    match msg {
        // **A band's own client is the picture and nothing else** — the chrome's rows over it are the
        // pin's own window's (see `pin_input::pinned_press_on`) — so a hand on it is the scene's.
        WM_NCHITTEST => Some(LRESULT(HTCLIENT as isize)),
        WM_LBUTTONDOWN => {
            // **The pointer is taken before the down is posted**, and f3d takes it in turn: the engine's
            // own capture on the down is what hands it the real moves of the rest of the drag, and what is
            // taken here is the stretch before that down has been read (see `MODEL_DRAG`).
            MODEL_DRAG.store(host.0 as isize, Ordering::SeqCst);
            let _ = unsafe { SetCapture(host) };
            unsafe { post_a_band_message(host, source, msg, message_point(lparam)) };
            Some(LRESULT(0))
        }
        WM_MOUSEMOVE => {
            if MODEL_DRAG.load(Ordering::SeqCst) != host.0 as isize {
                return None;
            }
            // **A drag whose release went to the engine is ended here**, on the level of the button: a
            // latched drag would post a press's moves on a hand that is only over the band, and one ended
            // without its release is a scene left turning (see `end_the_models_drag`).
            if !the_left_button_is_down() {
                unsafe { end_the_models_drag(host, message_point(lparam)) };
                return None;
            }
            unsafe { post_a_band_message(host, source, msg, message_point(lparam)) };
            Some(LRESULT(0))
        }
        WM_LBUTTONUP => {
            if MODEL_DRAG.load(Ordering::SeqCst) != host.0 as isize {
                return None;
            }
            unsafe { end_the_models_drag(host, message_point(lparam)) };
            Some(LRESULT(0))
        }
        _ => None,
    }
}

/// Let go of a band the pointer was carrying: **the release is posted before the drag's own state is
/// cleared**, because a latch that goes without it is a scene left turning under a hand that is up — and the
/// capture is given back only where this window is the one holding it, since an engine that took the capture
/// from us holds it to its own release (see `settle_a_models_drag`).
unsafe fn end_the_models_drag(host: HWND, point: (i32, i32)) {
    if let Some(source) = the_models_window() {
        unsafe { post_a_band_message(host, source, WM_LBUTTONUP, point) };
    }
    MODEL_DRAG.store(0, Ordering::SeqCst);
    if unsafe { GetCapture() } == host {
        let _ = unsafe { ReleaseCapture() };
    }
}

/// End a model's turn no window of ours was sent the release of: the moves of a drag are the engine's once
/// it takes the capture on the down posted to it, so the level of the button is the one answer left — asked
/// from the loop's own tick, with the release posted to the engine before the latch goes (see
/// `answer_a_models_band`).
pub(super) fn settle_a_models_drag() {
    let host = MODEL_DRAG.load(Ordering::SeqCst);
    if host == 0 || the_left_button_is_down() {
        return;
    }

    let host = HWND(host as *mut _);
    let point = the_pointer_in_the_band(host);
    unsafe { end_the_models_drag(host, point) };
}

/// The pointer's place in a band's own client area, which is where the release of a drag no window of ours
/// was sent is posted — and the band's own corner where the pointer cannot be read at all.
fn the_pointer_in_the_band(host: HWND) -> (i32, i32) {
    let Some((x, y)) = pin_input::cursor_screen_point() else {
        return (0, 0);
    };

    let mut point = POINT { x, y };
    let _ = unsafe { ScreenToClient(host, &mut point) };
    (point.x, point.y)
}

/// Turn the scene a press on a model's band is for, and answer whether the band under the press is a
/// model's at all — a press it is on a model whose window is not up yet, which carries nothing rather than
/// the window its picture is shown in (see `answer_a_models_band`).
pub(super) unsafe fn press_the_models_band(hwnd: HWND, x: i32, y: i32) -> bool {
    if !a_model_takes_the_pointer(current_media_type_try()) {
        return false;
    }

    unsafe {
        answer_a_models_band(
            hwnd,
            WM_LBUTTONDOWN,
            WPARAM(0),
            a_point_in_a_message((x, y)),
        );
    }
    true
}

/// Whether the left button is down, read as a **level** and never for its press bit: the low bit
/// `GetAsyncKeyState` answers is spent by any reader, and it is the Explorer hook's click (see
/// `pin_drag::left_button_down`, which publishes the same state for the pin's own gestures).
fn the_left_button_is_down() -> bool {
    unsafe { GetAsyncKeyState(VK_LBUTTON.0 as i32) as u16 & 0x8000 != 0 }
}

/// Post one message of the pointer's to the engine's window, at a point in the band's own client area, and
/// answer whether it went.
unsafe fn post_a_band_message(host: HWND, source: HWND, msg: u32, point: (i32, i32)) -> bool {
    let Some(point) = (unsafe { the_point_in_the_source(host, source, point) }) else {
        return false;
    };
    // A release carries no button; the down and every move of the drag carry the one that is held —
    // `MK_LBUTTON`, which these bindings keep under a feature this app has no other use for.
    let held = if msg == WM_LBUTTONUP {
        WPARAM(0)
    } else {
        WPARAM(MK_LBUTTON)
    };
    unsafe { PostMessageW(source, msg, held, a_point_in_a_message(point)).is_ok() }
}

/// The band's point in the engine window's client: the band mapped to the screen and the screen into the
/// engine's own client (see `a_point_in_the_source`).
unsafe fn the_point_in_the_source(
    host: HWND,
    source: HWND,
    point: (i32, i32),
) -> Option<(i32, i32)> {
    let mut band = POINT::default();
    let mut origin = POINT::default();
    let mut client = RECT::default();
    if !unsafe { ClientToScreen(host, &mut band) }.as_bool() {
        return None;
    }
    if !unsafe { ClientToScreen(source, &mut origin) }.as_bool() {
        return None;
    }
    if unsafe { GetClientRect(source, &mut client) }.is_err() {
        return None;
    }
    a_point_in_the_source(
        point,
        (band.x, band.y),
        (origin.x, origin.y),
        (client.right, client.bottom),
    )
}

/// Where a point in a band's own client area falls in the engine window's client, or nothing where the
/// engine's client is no box at all.
///
/// **The two rectangles are mapped, never assumed equal.** What a band draws is the Shell's thumbnail of
/// the source, cropped to the band (see `pin_film_thumbnail`), so a resize landing between two moves or a
/// picture the host does not fill puts them out of step — and a message posted in the band's own numbers
/// would be a press near the engine's own corner.
///
/// A point outside the source is **clamped to its edge** rather than dropped or posted negative: f3d reads
/// a client point outside its box as a press outside its window, and a hand run off the band is a turn
/// still being asked for at the edge it left by (measured on 4.0.1).
fn a_point_in_the_source(
    point: (i32, i32),
    band_origin: (i32, i32),
    source_origin: (i32, i32),
    source_client: (i32, i32),
) -> Option<(i32, i32)> {
    a_point_in_a_client(
        (
            band_origin.0 + point.0 - source_origin.0,
            band_origin.1 + point.1 - source_origin.1,
        ),
        source_client,
    )
}

/// A point clamped into a box, or nothing where the box has no room for one (see `a_point_in_the_source`).
fn a_point_in_a_client(point: (i32, i32), client: (i32, i32)) -> Option<(i32, i32)> {
    if client.0 <= 0 || client.1 <= 0 {
        return None;
    }
    Some((
        point.0.clamp(0, client.0 - 1),
        point.1.clamp(0, client.1 - 1),
    ))
}

/// The `lParam` a message carrying a point is posted with: the point's two words, low one first.
fn a_point_in_a_message(point: (i32, i32)) -> LPARAM {
    LPARAM((((point.1 as u16 as u32) << 16) | point.0 as u16 as u32) as i32 as isize)
}

/// The `wParam` a wheel message is posted with: the notch in the high word, where a wheel delta sits, and
/// the key state — nothing — in the low one. Signed, so a notch back is the two's complement of one on.
fn a_notch_in_a_message(notches: i32) -> WPARAM {
    WPARAM(((notches as u16 as u32) << 16) as usize)
}

/// Hand a notch turned over a model's band to the model: f3d's wheel is its camera's zoom, so the notch is
/// the engine's and nothing of this app's (see `pin_model::the_models_band`).
///
/// **The wheel's point is a screen point where the drag's is a client one** (measured on 4.0.1): the notch is
/// read at the engine's own place on the desktop, so the point is clamped into the engine's client and given
/// back out as the screen point that client point stands at — a band larger than the source would otherwise
/// ask for a zoom outside the window the zoom is of.
pub(crate) fn post_a_notch_to_the_model(point: POINT, notches: i32) {
    let Some(source) = the_models_window() else {
        return;
    };
    let mut local = point;
    let mut client = RECT::default();
    if !unsafe { ScreenToClient(source, &mut local) }.as_bool() {
        return;
    }
    if unsafe { GetClientRect(source, &mut client) }.is_err() {
        return;
    }
    let Some((x, y)) = a_point_in_a_client((local.x, local.y), (client.right, client.bottom))
    else {
        return;
    };
    let mut packed = POINT { x, y };
    if !unsafe { ClientToScreen(source, &mut packed) }.as_bool() {
        return;
    }
    let _ = unsafe {
        PostMessageW(
            source,
            WM_MOUSEWHEEL,
            a_notch_in_a_message(notches),
            a_point_in_a_message((packed.x, packed.y)),
        )
    };
}

#[cfg(test)]
mod tests {
    use super::*;

    /// **A band's point is mapped into the engine's client and clamped at its edges**, and nothing is posted
    /// for a source with no client at all. The band and the source are the same rectangle in the ordinary
    /// case, and are not once a resize or a crop has put them out of step — which is the case the mapping
    /// exists for.
    #[test]
    fn a_bands_point_is_mapped_into_the_source_and_clamped_to_it() {
        let client = (600, 480);

        assert_eq!(
            a_point_in_the_source((10, 20), (100, 100), (100, 100), client),
            Some((10, 20)),
            "a band at the source is the source"
        );
        assert_eq!(
            a_point_in_the_source((10, 20), (200, 150), (180, 140), client),
            Some((30, 30)),
            "a band 20 right and 10 below the source is every point of it 20 right and 10 below"
        );
        assert_eq!(
            a_point_in_the_source((300, 240), (100, 100), (180, 140), client),
            Some((220, 200)),
            "and a point in the middle of the source is the middle of it however the band is placed"
        );

        assert_eq!(
            a_point_in_the_source((-40, 700), (100, 100), (100, 100), client),
            Some((0, 479)),
            "a point run off the band is the edge it left by, never a negative one"
        );
        assert_eq!(
            a_point_in_the_source((600, 480), (0, 0), (0, 0), client),
            Some((599, 479)),
            "and the last pixel of the source is the last one the mapping names"
        );
        assert_eq!(
            a_point_in_the_source((10, 10), (0, 0), (0, 0), (0, 0)),
            None,
            "a source with no client is no window to post to"
        );
    }

    /// **A press, a move and a release carry the point in the message's two words, low one first**, and a
    /// notch carries its delta in the high word — where a wheel delta sits and where the negative one lives.
    #[test]
    fn a_message_carries_its_point_and_its_notch_where_f3d_reads_them() {
        assert_eq!(a_point_in_a_message((0x1234, 0x5678)).0, 0x5678_1234);
        assert_eq!(message_point(a_point_in_a_message((300, 240))), (300, 240));

        assert_eq!(a_notch_in_a_message(120).0, 0x0078_0000);
        assert_eq!(a_notch_in_a_message(-120).0, 0xFF88_0000);
        assert_eq!(a_notch_in_a_message(0).0, 0);
    }

    /// **Only a model is a band the pointer turns.** A film's band is a hand on the pin, and every kind drawn
    /// into a window of this app's is a press this app answers itself (see `a_model_takes_the_pointer`), and
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
}

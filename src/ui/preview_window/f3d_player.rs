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
//! **And what a pin shows of one is the window itself**, standing in the pin's band and taking the hand
//! on it: f3d turns its scene from its own captured drag, so a press posted into it is a press it never
//! reads (see `place_the_model_in_the_band`). A hover keeps the Shell's thumbnail of that same
//! window.

use windows::Win32::UI::WindowsAndMessaging::{
    GWL_STYLE, WS_CAPTION, WS_MAXIMIZEBOX, WS_MINIMIZEBOX, WS_SYSMENU, WS_THICKFRAME,
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

/// The engine window a model's picture is drawn in, where the picture standing in the preview is a
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
            pin_film_thumbnail::the_source_window().map(|source| source.0 as isize),
        ),
    )
}

/// The window the model's picture is drawn in, from the two roads that name one.
///
/// **The published window is the picture's own** — the monitor writes it the moment f3d's window exists,
/// and every road that places the picture reads it there. The source is what stands in for it in one
/// case: a window a hover's host was showing publishes nothing of the desktop's until the pin's own road
/// publishes it, and a promotion from a hover is that same window rather than a scene begun again (see
/// `pin_film_thumbnail::the_source_window` and `place_the_model_in_the_band`).
fn a_models_source(published: isize, source: Option<isize>) -> isize {
    if published != 0 {
        return published;
    }
    source.unwrap_or(0)
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

/// How far inside the pin's own band a model's window is placed.
///
/// **Nothing the pin paints can clip a window that stands over it.** A pinned model's window is f3d's
/// own, standing over the pin's surface rather than under it, so the band is handed in by the frame the
/// pin draws around it: what is left outside the scene is the pin's edge and its corner, and a scene
/// drawn to the band's own rectangle would paint over both.
///
/// shortcut: the pinned band's paint is a hairline with no radius today, so one pixel a side; upgrade to
/// the painted radius when the band's paint gains one.
const MODEL_BAND_INSET: i32 = 1;

/// The box a pinned model's window stands in, from the band the pin says its picture occupies.
///
/// Pure so the inset can be asked without a window (see `MODEL_BAND_INSET`): every side moved in by the
/// same count, and never past a box of one pixel on either axis.
pub(super) fn the_models_band_inset(content: ScreenRegion) -> ScreenRegion {
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

/// Put a pinned model's window in the pin's band, and keep it there.
///
/// **The picture is f3d's own window and not the Shell's thumbnail of it.** A hand on the band is a hand
/// on the scene, and a scene behind a thumbnail is one nothing can turn — so every road that places a
/// window standing beside the pin places this one, at the band inset by the frame the pin paints around
/// it (see `the_models_band_inset`).
///
/// **Over the pin and never over its chrome.** The pin's own surface under the scene is the frame and
/// nothing else, and the band is the scene's: the pin is not put back over it as it is over a film, which
/// would be the press taken from the scene (see `place_player_window`, whose pin-back a model does not
/// get). What is drawn over the band still is: the caption's menu, raised by its own tick, and the
/// windows the hold-off below leaves the scene where it stands for (see `player_order`).
///
/// **Nothing of the pin's chrome crosses a model's band**: a model plays nothing, so there is no
/// transport row and no volume popup over it, and the caption's rows are above the band (see
/// `pin_transport_kind` and `pinned_caption_height`).
///
/// Safety: the window is f3d's, and everything here is a style, a placement and a show of it.
pub(super) unsafe fn place_the_model_in_the_band(content: ScreenRegion) {
    let Some(window) = the_models_window() else {
        return;
    };

    // **The style is the one every player's window goes through**, and the model's own answer is inside
    // it: f3d's frame off, a press reaching the scene, and the bits that keep it out of the Shell's
    // lists (see `ffplay_window::player_ex_style_for`).
    let current = unsafe { GetWindowLongPtrW(window, GWL_EXSTYLE) };
    let styled = ffplay_window::player_ex_style_for(current);
    if styled != current {
        let _ = unsafe { SetWindowLongPtrW(window, GWL_EXSTYLE, styled) };
    }

    // **The raise is held off where something of the pin's is drawn over the band**, which is the one
    // list every window in a band answers to (see `player_order`).
    if video_playback::player_order(
        video_process_image(),
        true,
        pin_chrome_up(),
        pin_volume_open(),
        pin_menu_open(),
    ) == video_playback::PlayerOrder::Leave
    {
        return;
    }

    let (left, top, right, bottom) = the_models_band_inset(content);
    let (width, height) = ((right - left).max(1), (bottom - top).max(1));
    // A window standing in front of a pin is raised by the list a running page is raised by: topmost
    // while the pin holds the front, and the top of the ordinary band where it does not (see
    // `webview_preview::raises`).
    for raise in webview_preview::raises(true, pin_holds_the_front())
        .into_iter()
        .flatten()
    {
        let _ = unsafe {
            SetWindowPos(
                window,
                raise.insert_after,
                left,
                top,
                width,
                height,
                raise.flags,
            )
        };
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    /// **The band a model's window stands in is the pin's own, taken in by the frame the pin paints
    /// around it**, because nothing the pin paints can clip a window standing over it (see
    /// `the_models_band_inset`).
    #[test]
    fn a_pinned_models_band_is_inset_inside_the_frame_the_pin_paints() {
        assert_eq!(
            the_models_band_inset((100, 50, 300, 250)),
            (101, 51, 299, 249),
            "every side of the pin's band is taken in by the frame it paints around it"
        );
        assert_eq!(
            the_models_band_inset((10, 10, 12, 12)),
            (11, 11, 12, 12),
            "and a band too small to take in is still a box of a pixel: the scene is placed, never              handed a box of nothing"
        );
    }

    /// **A model in front of a pin is raised over the pin and never over what the pin draws over the
    /// band**: the raise is the topmost one a window standing in front of a pin is raised by, and it is
    /// held off while the pin's chrome, its volume popup or the caption's menu is up (see
    /// `place_the_model_in_the_band`, `webview_preview::raises`, `video_playback::player_order`).
    #[test]
    fn a_model_is_raised_over_the_pin_and_never_over_its_chrome() {
        let raises = webview_preview::raises(true, false);
        assert_eq!(
            raises[0].map(|raise| raise.insert_after),
            Some(HWND_TOPMOST)
        );
        assert_eq!(
            raises[0].map(|raise| raise.flags),
            Some(SWP_NOACTIVATE | SWP_SHOWWINDOW),
            "a pin that does not hold the front is raised over, and shown, by the topmost call"
        );
        assert_eq!(
            raises[1].map(|raise| raise.insert_after),
            Some(HWND_NOTOPMOST),
            "and dropped to the top of the ordinary band after it"
        );
        assert_eq!(
            webview_preview::raises(true, true)[1],
            None,
            "while a pin that holds the front keeps the scene topmost with it"
        );

        for (chrome, volume, menu) in [
            (false, false, false),
            (true, false, false),
            (false, true, false),
            (false, false, true),
        ] {
            let order = video_playback::player_order("", true, chrome, volume, menu);
            let asked = if chrome || volume || menu {
                video_playback::PlayerOrder::Leave
            } else {
                video_playback::PlayerOrder::Raise
            };
            assert_eq!(
                order, asked,
                "chrome {chrome} volume {volume} menu {menu}: the band's window is raised over                  nothing the pin draws over it"
            );
        }
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

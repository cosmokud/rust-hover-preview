//! The window a pinned mpv film is drawn in: a disabled child of the pin's own window, standing in
//! the band the picture is seen through, hidden for a park and shown again after it, and going down
//! with the pin it was made for.
//!
//! Beside it, the chrome it buries: the caption and the transport are painted into the pin's own
//! layered surface, which a child composites over, so the strips of chrome live in windows of
//! their own that stand over the host (see `pin_chrome_strips`).
//!
//! Every test here stands a pin in a window of its own (see `pin_real_window`) and a film mpv is
//! playing, because the host only exists for that pair: a film FFmpeg's player draws is a window of
//! its own until its ticket, and every other kind is this app's own frame.

use super::*;
use windows::Win32::Foundation::{LPARAM, LRESULT, WPARAM};
use windows::Win32::UI::Input::KeyboardAndMouse::IsWindowEnabled;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, GetParent, GetWindowLongPtrW, IsWindowVisible,
    MoveWindow, SetWindowLongPtrW, ShowWindow, GWLP_WNDPROC, GWL_EXSTYLE, GWL_STYLE, GW_HWNDPREV,
    HWND_TOP, SWP_NOACTIVATE, SWP_NOZORDER, SW_SHOWNOACTIVATE, WM_SHOWWINDOW, WINDOW_EX_STYLE,
    WS_CHILD, WS_DISABLED, WS_EX_APPWINDOW, WS_EX_LAYERED, WS_EX_TOPMOST, WS_POPUP, WS_VISIBLE,
};

/// The show and hide messages the adopted host received, in order: whether it was shown, the style
/// it wore at that moment, and its parent. The style and parent are what tells the hide before the
/// restyle from the show after the `SetParent` (see the move order test).
static HOST_SHOWS: Lazy<Mutex<Vec<(bool, isize, isize)>>> = Lazy::new(|| Mutex::new(Vec::new()));

unsafe extern "system" fn watch_the_host_shows(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    if msg == WM_SHOWWINDOW {
        let parent = unsafe { GetParent(hwnd) }
            .ok()
            .map(|parent| parent.0 as isize)
            .unwrap_or(0);
        let style = unsafe { GetWindowLongPtrW(hwnd, GWL_STYLE) };
        if let Ok(mut shows) = HOST_SHOWS.lock() {
            shows.push((wparam.0 != 0, style, parent));
        }
    }
    unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
}

/// The place a window of a test's own is put when it must not be on the user's screen.
const NOWHERE: i32 = -32000;

/// The hover's own window: the preview a `Hide` was posted for, which is not the pin's window.
///
/// Safety: the window is the caller's to destroy.
unsafe fn a_hover_window_of_the_tests_own() -> HWND {
    let hwnd = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            NOWHERE,
            NOWHERE,
            320,
            240,
            None,
            None,
            None,
            None,
        )
    }
    .expect("a hover window of the test's own");
    let _ = unsafe { ShowWindow(hwnd, SW_SHOWNOACTIVATE) };
    hwnd
}

unsafe fn window_rect(hwnd: HWND) -> ScreenRegion {
    let mut rect = RECT::default();
    let _ = GetWindowRect(hwnd, &mut rect);
    (rect.left, rect.top, rect.right, rect.bottom)
}

/// A pin of a film standing in its own window at `box_`, with mpv as the engine that plays it and
/// the pin's pid the one this process answers to.
///
/// The process is the test's own rather than a player's, because what the host is asked for a film
/// is the pid of the child mpv would make under it — and a child this test makes is a child of the
/// same process by the same rule (see `the_hosted_player_is_drawing`).
///
/// Every global written here is put back when the stand is dropped, and the two windows are the
/// test's to destroy.
struct StandingFilm {
    box_: ScreenRegion,
    pin_hwnd: HWND,
    previous_pin: Option<PinnedPreview>,
    previous_own: bool,
    previous_window: Option<HWND>,
    previous_media: Option<MediaData>,
    previous_pid: u32,
    previous_hwnd: isize,
}

impl StandingFilm {
    /// The same stand with the film drawing in the host: `drawn` is the player whose child under it
    /// is the picture (see `the_hosted_player_is_drawing`), and it owns no window of the desktop's.
    fn drawn_in_the_host(&self, drawn: u32) {
        pin_film_host::note_hosted_player(drawn);
        VIDEO_PID.store(drawn, Ordering::SeqCst);
    }

    /// The same stand taken up on a hover's player: the film is mpv's, and the window it is drawing
    /// in is a window of the desktop's rather than the host's (see `the_pinned_player_was_begun_in_the_host`),
    /// with `hovered` the process that owns it.
    fn taken_up_from_a_hover(&self, hovered: u32) {
        pin_film_host::note_hosted_player(0);
        VIDEO_PID.store(hovered, Ordering::SeqCst);
    }

    fn new() -> StandingFilm {
        let previous_pin = take_pin_for_a_test();
        let previous_own = pin_is_its_own_window();
        let previous_window = pin_real_window::real_pin_window();
        let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
        let previous_pid = VIDEO_PID.swap(unsafe { GetCurrentProcessId() }, Ordering::SeqCst);
        // **A hosted film has no window of the desktop's**, so the road that would find one by pid
        // is stood at nothing here (see `video_window_for`).
        let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);
        // A host left by a test that took its pin's window down without taking the host with it is
        // a handle to nothing by now; the stand begins from no host at all. Taken down before the
        // player is noted, because taking it down forgets the player beside it.
        unsafe { pin_film_host::take_down_the_pin_film_host() };
        // And no strips either: they belong to the pin's window rather than to the host, so a
        // take-down of the host leaves whatever handles a test before this one stood behind.
        unsafe { pin_chrome_strips::take_the_strips_down() };
        // The host holds this test's own process, which is the pid of any child it makes: a film is
        // drawn in the host only when the player begun into it is the one this app recorded (see
        // `the_pinned_player_was_begun_in_the_host`).
        pin_film_host::note_hosted_player(unsafe { GetCurrentProcessId() });

        stand_mpv_in(Some(true));
        stand_video_engine_in(Some((VideoEngine::Mpv, true)));

        let box_ = (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240);
        let mut pin = overlay_pin(box_, PinChrome::always());
        pin.path = PathBuf::from("a-film.mkv");
        // The take-up's own question, asked of the film the way the loop asks it: the host exists
        // only for a film whose pin the classification puts in a window of its own (see
        // `pin_should_stand_in_its_own_window`).
        let its_own = pin_should_stand_in_its_own_window(Some(MediaType::Video), &pin.path);
        stand_pin(Some(pin));
        pin_stands_in_its_own_window(its_own);

        pin_real_window::stand_real_pin_window_for_a_test(None);
        let pin_hwnd =
            unsafe { pin_real_window::real_pin_window_at(box_) }.expect("the pin's own window");

        let mut media = create_loading_media(320, 240);
        media.media_type = MediaType::Video;
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = Some(media);
        }

        StandingFilm {
            box_,
            pin_hwnd,
            previous_pin,
            previous_own,
            previous_window,
            previous_media,
            previous_pid,
            previous_hwnd,
        }
    }
}

impl Drop for StandingFilm {
    fn drop(&mut self) {
        // A handover armed by a take-up that never reached its settle — a test that failed
        // before it disarmed — must not leak into the next test's chrome (see
        // `pin_handover_armed`).
        pin_film_host::disarm_pin_handover();
        unsafe {
            pin_film_host::take_down_the_pin_film_host();
            // The strips with the host: the window going down destroys them, and the handles must
            // not outlive the windows (see `take_the_strips_down`).
            pin_chrome_strips::take_the_strips_down();
        }
        pin_real_window::stand_real_pin_window_for_a_test(self.previous_window);
        unsafe {
            let _ = DestroyWindow(self.pin_hwnd);
        }
        pin_stands_in_its_own_window(self.previous_own);
        stand_pin(self.previous_pin.take());
        VIDEO_PID.store(self.previous_pid, Ordering::SeqCst);
        VIDEO_HWND.store(self.previous_hwnd, Ordering::SeqCst);
        stand_mpv_in(None);
        stand_video_engine_in(None);
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = self.previous_media.take();
        }
    }
}

/// A window of the player's own under the host, which is the film drawing where the pin can see
/// it (see `the_hosted_player_is_drawing`): made by this process, which is the player the stand's
/// host holds.
unsafe fn a_child_of_the_host(host: HWND) -> HWND {
    unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_CHILD | WS_VISIBLE,
            0,
            0,
            16,
            16,
            host,
            None,
            None,
            None,
        )
    }
    .expect("a child of the host's, in this process")
}

/// The hosted player's own pipe, and the player's word that the frame it is standing on is on
/// screen: the second the hover froze at, and a video output configured for that frame (see
/// `the_hosted_frame_is_presented`).
fn the_hosted_player_has_the_frame(pid: u32, at: &str) {
    mpv_ipc::remember_pipe(pid, "the-hosted-players-pipe", false);
    mpv_ipc::answer_a_read(pid, "time-pos", Some(at));
    mpv_ipc::answer_a_read(pid, "vo-configured", Some("true"));
}

/// A park holds the last frame the band showed, and a hosted film's band is the host: the player's
/// window is a child inside it and not a window of the desktop's, so a capture that went looking for
/// one by pid would find nothing and the cover would be painted flat — the band this app fills with
/// a colour rather than with the film (see `hold_video_window_frame`).
#[test]
fn a_hosted_park_holds_the_frame_the_host_is_showing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let drawn = 0x7FFF_FFFD;
    film.drawn_in_the_host(drawn);
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");

    // The band where the screen can be read, at a size of its own so that the frame says which
    // window it was taken from.
    unsafe {
        let _ = MoveWindow(film.pin_hwnd, 100, 100, 320, 240, true);
        let _ = SetWindowPos(host, HWND_TOP, 0, 0, 64, 48, SWP_NOZORDER | SWP_NOACTIVATE);
        let _ = ShowWindow(host, SW_SHOWNOACTIVATE);
    }
    forget_video_frame();

    assert!(
        hold_video_window_frame(),
        "the frame a hosted film's park holds is the host's, which is where the picture is"
    );
    let held = held_video_frame();
    let held = held.as_ref().expect("the frame the park holds");
    assert_eq!(
        (held.width, held.height),
        (64, 48),
        "read at the band the film is drawn in rather than at a box of the desktop's"
    );
}

/// A film in the pin's own host is a child of the pin's window: it has no place in the desktop's
/// order, no band to hold, and no style of this app's to wear. Every road that places, styles,
/// raises or tells a top-level player's window is answered for it by the host — and every one of
/// them is left working unchanged for a film that *is* a window of the desktop's (see
/// `the_pinned_player_was_begun_in_the_host`).
#[test]
fn the_top_level_roads_leave_a_hosted_player_alone() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    let previous_image = video_process_image();
    note_player_process(pid, MPV_IMAGE_NAME);
    mpv_ipc::remember_pipe(pid, "a-hosted-ontop-pipe", false);
    mpv_ipc::clear_noted();

    let player = unsafe { a_window_of_the_desktops() };
    let box_ = (200, 160, 840, 560);
    unsafe {
        let _ = SetWindowPos(
            player,
            HWND_TOP,
            box_.0,
            box_.1,
            box_.2 - box_.0,
            box_.3 - box_.1,
            SWP_NOACTIVATE,
        );
    }

    // The band a hosted film holds is the pin's own: the player's flag is not the pin's to correct.
    settle_pinned_mpv_ontop();
    assert!(
        mpv_ipc::noted_lines().is_empty(),
        "a player drawn into the pin's own host is told nothing about a band of the desktop's"
    );

    // The style list is the top-level player's: a child of a disabled host takes no press and
    // needs none of it.
    assert_eq!(
        player_ex_style_for(0),
        0,
        "a hosted player wears none of the bits a window of the desktop's is given"
    );

    assert!(
        !unsafe { apply_noactivate_to_hwnd(player) },
        "and the road that styles one answers that it styled nothing"
    );
    assert_eq!(
        unsafe { GetWindowLongPtrW(player, GWL_EXSTYLE) },
        0,
        "which leaves the window exactly as it was"
    );
    assert_eq!(
        VIDEO_HWND.load(Ordering::SeqCst),
        0,
        "and publishes no handle of a hosted film's as the player's to place"
    );

    unsafe {
        place_player_window(
            player,
            PlayerOrder::Raise,
            (10, 10, 410, 310),
            SWP_NOACTIVATE,
        );
    }
    assert_eq!(
        unsafe { window_rect(player) },
        box_,
        "and the road that places and raises a player's window does not move this one"
    );

    unsafe {
        let _ = DestroyWindow(player);
    };
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::clear_noted();
    note_player_process(pid, previous_image);
}

/// A window of the desktop's, standing nowhere in particular: what the top-level roads are asked
/// about in place of a player of this app's, which a test never begins.
unsafe fn a_window_of_the_desktops() -> HWND {
    unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            0,
            0,
            16,
            16,
            None,
            None,
            None,
            None,
        )
    }
    .expect("a stand-in window for a player's")
}

/// The film's host goes down with the pin it drew it for: a window of this app's own that stood in
/// the pin's band, and a handle left standing for the pin after it is a band of nothing the next
/// pin would take up as its picture (see `pin_film_host`).
#[test]
fn a_pin_that_ends_takes_its_film_host_with_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");

    let _ = end_pin(Reason::Closed, &Win32PinWindow);

    assert!(
        !unsafe { IsWindow(host) }.as_bool(),
        "the window the film was drawn in goes down with the pin, whichever road ended it"
    );
    assert!(
        pin_film_host::film_host().is_none(),
        "with the handle forgotten as well as the window destroyed"
    );
}

/// The film's host is the band, and the band is what a resize changes: the host follows the pin's
/// box wherever the loop is told it, without a player's geometry being touched (see
/// `place_the_pin_film_host`).
#[test]
fn the_film_host_follows_the_pins_band() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");

    // A resize of the pin's window hands the media band another box, and the loop lays the window
    // out around it: the host is the band inside that window, so it is the same box the picture
    // would be shown in.
    let wider = (
        film.box_.0,
        film.box_.1,
        film.box_.0 + 400,
        film.box_.1 + 300,
    );
    with_pin(|pin| pin.content = wider);
    place_pinned_siblings();

    assert_eq!(
        unsafe { window_rect(host) },
        wider,
        "the host stands in the band the pin now shows its film through"
    );

    // And the band keeps being followed while the film is away: a park hides the host but does not
    // stop the band it comes back at from moving, so a resize under a cover places the hidden host
    // rather than the picture it will show again (see `settle_pinned_park`).
    let taller = (wider.0, wider.1, wider.0 + 400, wider.1 + 500);
    hide_pinned_player_window();
    ensure_pinned_sibling_box(taller);
    assert_eq!(
        unsafe { window_rect(host) },
        taller,
        "the band a hidden host comes back at is the box it was placed at"
    );
    assert!(
        !unsafe { IsWindowVisible(host) }.as_bool(),
        "and placing it is not the show a park is waiting to end"
    );
}

/// A hover's mpv film drawn in the app-owned top-level host, with its player's child under it:
/// the state a take-up can move into the pin rather than begin again.
struct HostedHover {
    host: HWND,
    child: HWND,
    pid: u32,
    at: f64,
}

impl HostedHover {
    unsafe fn up(film: &StandingFilm, path: &Path, at: f64) -> HostedHover {
        let pid = unsafe { GetCurrentProcessId() };
        let host = unsafe { hover_film_host::take_up_the_hover_film_host(film.box_) }
            .expect("the hover's own host");
        hover_film_host::note_hover_film_host_player(pid, path);
        let child = unsafe { a_child_of_the_host(host) };
        unsafe { hover_film_host::reveal_the_hover_film_host(pid, film.box_) };
        note_video_loop(pid, path, film.box_, at, None);
        VIDEO_PID.store(pid, Ordering::SeqCst);
        HostedHover {
            host,
            child,
            pid,
            at,
        }
    }
}

impl Drop for HostedHover {
    fn drop(&mut self) {
        // A host the pin adopted belongs to the pin now; the stand's take-down owns that one.
        if hover_film_host::hover_film_host() == Some(self.host) {
            unsafe { hover_film_host::take_down_the_hover_film_host() };
        }
        forget_video_loop(self.pid);
    }
}

/// The move route is for an mpv hover that is the pinned file and is still drawing in the host:
/// another route, another file, or a player the host no longer holds is the fallback handover.
#[test]
fn only_a_hosted_mpv_hover_is_moved_into_the_pin() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let hover = unsafe { HostedHover::up(&film, &path, 2.0) };

    assert_eq!(
        hover_film_host::the_hover_film_to_adopt(&path),
        Some(hover.host),
        "the hosted mpv hover of the pinned file is the one that moves"
    );
    assert!(
        hover_film_host::the_hover_film_to_adopt(Path::new("a-picture.png")).is_none(),
        "a file of another route is not moved"
    );
    assert!(
        hover_film_host::the_hover_film_to_adopt(Path::new("another-film.mkv")).is_none(),
        "nor is a player begun for another file"
    );

    unsafe { hover_film_host::take_down_the_hover_film_host() };
    assert!(
        hover_film_host::the_hover_film_to_adopt(&path).is_none(),
        "and a player the host no longer holds has nothing to move"
    );
}

/// **The already-running hosted hover is moved, not begun again.** The same player and the same
/// child survive; the only window that changes is the host's, and it becomes a disabled child of
/// the pin. Nothing is frozen, read back, relaunched or held for a handover.
#[test]
fn a_film_taken_up_from_a_hosted_hover_launches_nothing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let hover = unsafe { HostedHover::up(&film, &path, 12.5) };
    clear_restart_count();
    pin_film_host::disarm_pin_handover();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();

    unsafe { show_pinned_window(film.pin_hwnd) };

    assert_eq!(
        restart_count(),
        0,
        "the hover's player is kept, not begun again"
    );
    let host = pin_film_host::film_host().expect("the film's host");
    assert_eq!(
        host, hover.host,
        "the hover's own host is the pin's host now"
    );
    assert_eq!(
        unsafe { GetParent(host) }.ok(),
        Some(film.pin_hwnd),
        "and it travels as a child of the pin's window"
    );
    assert_ne!(
        unsafe { GetWindowLongPtrW(host, GWL_STYLE) } & WS_CHILD.0 as isize,
        0,
        "with the child style it was re-styled to"
    );
    assert_eq!(
        unsafe { GetWindowLongPtrW(host, GWL_EXSTYLE) } & WS_EX_TOPMOST.0 as isize,
        0,
        "and no topmost band of its own"
    );
    assert!(
        unsafe { IsWindow(hover.child) }.as_bool(),
        "the player's child is the same window, still drawing"
    );
    assert_eq!(
        unsafe { GetParent(hover.child) }.ok(),
        Some(host),
        "under the same host, now inside the pin"
    );
    assert_eq!(
        pin_film_host::hosted_player_pid(),
        Some(hover.pid),
        "and the pin holds the player it moved"
    );
    assert_eq!(VIDEO_PID.load(Ordering::SeqCst), hover.pid);
    assert_eq!(VIDEO_HWND.load(Ordering::SeqCst), 0);
    assert_eq!(
        hover_film_host::hover_film_host(),
        None,
        "the hover's handle is handed over rather than left to destroy the pin's host"
    );
    assert!(
        !pin_film_host::pin_handover_armed(),
        "nothing is waiting for a second player's first frame"
    );
    assert_eq!(
        pin_film_host::held_hover_for_a_test(),
        None,
        "and no frozen hover is held behind the pin"
    );
    assert!(
        !mpv_ipc::noted_lines()
            .iter()
            .any(|line| line.contains("pause")),
        "the film is neither frozen nor released: it simply carries on"
    );
    assert!(
        mpv_ipc::noted_reads().is_empty(),
        "and no frame is read back off it for a second launch"
    );
    let at = video_loop_playhead(hover.pid, &path).expect("the same clock");
    assert!(
        at >= hover.at,
        "the clock the take-up read is the hover's own playhead, still advancing: {at}"
    );
}

/// An FFmpeg hover that was already in the app's host is taken up the same way mpv's is: the same
/// pid, the same child, the same host handle, moved into the pin with nothing begun again. This is
/// the seamless hover-to-pin take-up: no relaunch, no black flash.
#[test]
fn an_ffplay_hover_taken_up_is_moved_into_the_pin_not_begun() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    let path = PathBuf::from("a-film.mkv");
    let hover = unsafe { HostedHover::up(&film, &path, 7.25) };
    clear_restart_count();

    assert_eq!(
        hover_film_host::the_hover_film_to_adopt(&path),
        Some(hover.host),
        "a hosted FFmpeg hover is a film the pin can take up without beginning another"
    );
    unsafe { show_pinned_window(film.pin_hwnd) };

    assert_eq!(
        restart_count(),
        0,
        "the hover's FFmpeg player is kept, not begun again"
    );
    assert_eq!(
        pin_film_host::film_host(),
        Some(hover.host),
        "the hover's own host is the pin's host now"
    );
    assert_eq!(pin_film_host::hosted_player_pid(), Some(hover.pid));
    assert_eq!(VIDEO_PID.load(Ordering::SeqCst), hover.pid);
    assert!(
        pin_film_host::the_pinned_player_was_begun_in_the_host(),
        "the player the hover began is the pin's own from here"
    );
    assert!(
        unsafe { IsWindow(hover.child) }.as_bool(),
        "the player's child is the same window, still drawing"
    );
    assert_eq!(
        unsafe { GetParent(hover.child) }.ok(),
        Some(hover.host),
        "under the same host, now inside the pin"
    );
    assert_eq!(
        unsafe { GetParent(hover.host) }.ok(),
        Some(film.pin_hwnd),
        "and the host travels as a child of the pin's window"
    );
    assert!(
        unsafe { IsWindowVisible(hover.host) }.as_bool(),
        "the taken-up host is shown in the pin's band"
    );
    let child = unsafe { window_rect(hover.child) };
    let host = unsafe { window_rect(hover.host) };
    assert_eq!(
        (child.2 - child.0, child.3 - child.1),
        (host.2 - host.0, host.3 - host.1),
        "and the child is fitted to the pin's band rather than left at the hover's box"
    );
    assert_eq!(VIDEO_HWND.load(Ordering::SeqCst), 0);
    assert_eq!(
        hover_film_host::hover_film_host(),
        None,
        "the hover's handle is handed over rather than left to destroy the pin's host"
    );
}

/// A hover whose FFmpeg player has not put its window up at take-up cannot be moved, so the pin
/// falls back to its own hosted launch: exactly one player is begun, never two.
#[test]
fn a_hover_that_cannot_be_adopted_falls_back_to_one_pinned_launch() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    let path = PathBuf::from("a-film.mkv");
    with_pin(|pin| pin.path = path.clone());
    let hovered = start_video_playback(&path, 0, 0, 320, 180, 0.0, false, 0, None)
        .expect("the hover's ffplay");
    assert!(
        hover_film_host::the_hover_film_to_adopt(&path).is_none(),
        "a player with no child under the host has nothing to move"
    );
    clear_restart_count();

    unsafe { show_pinned_window(film.pin_hwnd) };

    assert_eq!(
        restart_count(),
        1,
        "the pin begins one player of its own, not a second beside the hover's"
    );
    assert_ne!(
        VIDEO_PID.load(Ordering::SeqCst),
        hovered.id(),
        "and the hover's player is not the one the pin now holds"
    );
}

/// The pin's own level is put on the moved player, and the transport continues from the hover's
/// playhead rather than from a frozen second.
#[test]
fn a_hosted_hover_move_keeps_its_clock_and_takes_the_pins_volume() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let hover = unsafe { HostedHover::up(&film, &path, 41.0) };
    with_pin(|pin| {
        pin.volume.audio = false;
        pin.volume.level = 37;
        pin.volume.playing_at = 0;
        let at = video_loop_playhead(hover.pid, &path).unwrap_or(41.0);
        pin.transport.started = Some((Instant::now(), at));
    });
    clear_restart_count();
    pin_film_host::disarm_pin_handover();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    mpv_ipc::remember_pipe(hover.pid, "the-hovers-moved-pipe", false);

    unsafe { show_pinned_window(film.pin_hwnd) };

    assert!(
        mpv_ipc::noted_lines()
            .iter()
            .any(|line| line.contains(r#""volume",37"#)),
        "the pin's level is applied to the moved player: {:?}",
        mpv_ipc::noted_lines()
    );
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.volume.playing_at)),
        Some(37),
        "and the bar knows the level the player is at"
    );
    let playhead =
        pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin_playhead(&pin.transport)));
    assert!(
        playhead.is_some_and(|at| at >= 41.0),
        "the transport stands at the hover's playhead rather than at zero: {playhead:?}"
    );
    assert!(
        !mpv_ipc::noted_lines()
            .iter()
            .any(|line| line.contains("pause")),
        "and the film is not paused or released by the move"
    );
    mpv_ipc::forget_pipe(hover.pid);
}

/// The pin is painted and shown before the host moves, and the chrome is still hidden over the
/// moved film: only the pointer asks a strip back.
#[test]
fn a_hosted_hover_move_shows_the_pin_before_the_host_and_keeps_the_chrome_hidden() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let hover = unsafe { HostedHover::up(&film, &path, 3.0) };
    with_pin(|pin| pin.chrome = PinChrome::hidden());
    clear_restart_count();
    pin_film_host::disarm_pin_handover();

    HOST_SHOWS.lock().unwrap_or_else(|held| held.into_inner()).clear();
    let previous = unsafe {
        SetWindowLongPtrW(hover.host, GWLP_WNDPROC, watch_the_host_shows as *const () as isize)
    };
    unsafe { show_pinned_window(film.pin_hwnd) };
    unsafe { SetWindowLongPtrW(hover.host, GWLP_WNDPROC, previous) };

    assert!(
        unsafe { IsWindowVisible(film.pin_hwnd) }.as_bool(),
        "the pin is on screen"
    );
    assert!(
        unsafe { IsWindowVisible(hover.host) }.as_bool(),
        "and the moved film is shown inside it, not hidden with a hidden parent"
    );
    let shows = HOST_SHOWS.lock().unwrap_or_else(|held| held.into_inner());
    let hidden_before_the_restyle = shows
        .iter()
        .position(|(shown, style, parent)| {
            !shown && style & WS_POPUP.0 as isize != 0 && *parent == 0
        });
    assert!(
        hidden_before_the_restyle.is_some(),
        "the host is hidden while still a top-level window, before the style and parent change: \
         {shows:?}"
    );
    assert!(
        shows.iter().all(|(shown, _, _)| !shown),
        "and nothing shows it until the move's final `SetWindowPos`: {shows:?}"
    );
    assert_eq!(
        unsafe { GetParent(hover.host) }.ok(),
        Some(film.pin_hwnd),
        "and the host ends the move shown as the pin's own child"
    );
    assert!(
        unsafe { GetWindowLongPtrW(hover.host, GWL_STYLE) } & WS_CHILD.0 as isize != 0,
        "with the child style the final show hands it"
    );
    drop(shows);
    let chrome = pin_state().and_then(|pinned| {
        pinned.pin().map(|pin| {
            (
                pin.chrome.caption,
                pin.chrome.bar,
                pin.chrome.until.is_some(),
            )
        })
    });
    assert_eq!(
        chrome,
        Some((false, false, false)),
        "the chrome stays hidden over the moved picture"
    );
    for strip in [
        pin_chrome_strips::caption_strip(),
        pin_chrome_strips::transport_strip(),
    ]
    .into_iter()
    .flatten()
    {
        assert!(
            !unsafe { IsWindowVisible(strip) }.as_bool(),
            "no strip is shown over a frame the pointer has not asked about"
        );
    }
}

/// Closing a pin that moved a hover's host takes that host down with it, and the player's child
/// goes too; a park only hides the host and the unpark shows the same one.
#[test]
fn a_moved_hover_host_goes_with_the_pin_and_survives_a_park() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let hover = unsafe { HostedHover::up(&film, &path, 6.0) };
    clear_restart_count();
    pin_film_host::disarm_pin_handover();

    unsafe { show_pinned_window(film.pin_hwnd) };

    hide_pinned_player_window();
    assert!(
        !unsafe { IsWindowVisible(hover.host) }.as_bool(),
        "a park hides the moved host rather than taking it down"
    );
    show_pinned_player_window();
    assert!(
        unsafe { IsWindowVisible(hover.host) }.as_bool(),
        "and the unpark shows the same child of the pin"
    );

    let _ = end_pin(Reason::Closed, &Win32PinWindow);

    assert!(
        !unsafe { IsWindow(hover.host) }.as_bool(),
        "the pin's close destroys the host it moved"
    );
    assert!(
        !unsafe { IsWindow(hover.child) }.as_bool(),
        "and the player's child with it"
    );
    assert!(
        pin_film_host::film_host().is_none(),
        "with the handle forgotten as well as the window destroyed"
    );
}

/// A step after a moved take-up still begins the next player into the same host: the move hands
/// the host to the pin, and the pin's launch route is unchanged.
#[test]
fn a_step_after_a_moved_hover_uses_the_same_host() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let _hover = unsafe { HostedHover::up(&film, &path, 1.0) };
    clear_restart_count();
    pin_film_host::disarm_pin_handover();

    unsafe { show_pinned_window(film.pin_hwnd) };

    let host = pin_film_host::film_host().expect("the host the move took up");
    assert_eq!(
        mpv_launch_host(&path),
        Some(host.0 as usize as u32),
        "the next player launched for the pinned film is given the same host"
    );
    assert!(
        pin_film_host::the_pinned_player_was_begun_in_the_host(),
        "and it is the pin's host the launch route finds"
    );
}

/// **The hover is frozen where it stands and the pin's player is begun on the very second its picture
/// is at** — read back off the player rather than estimated from this app's clock, because the frame
/// the user is looking at is mpv's own last one and the two pictures have to be the same frame for the
/// swap to be invisible (see `begin_the_pinned_film_at_the_hovers_frame`).
#[test]
fn the_hover_is_frozen_on_its_own_frame_and_the_pin_begins_there() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let hovered = 0x7FFF_FFFB;
    film.taken_up_from_a_hover(hovered);
    clear_restart_count();
    pin_film_host::disarm_pin_handover();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    mpv_ipc::remember_pipe(hovered, "the-hovers-pipe", false);
    mpv_ipc::answer_a_read(hovered, "time-pos", Some("412.500000"));

    unsafe { show_pinned_window(film.pin_hwnd) };

    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![format!(
            "{}\n",
            r#"{"command":["set_property","pause",true],"async":true}"#
        )],
        "the hover is told to hold where it stands, and the frame it is showing is the freeze — \
         nothing is captured, painted over or hidden"
    );
    assert_eq!(
        mpv_ipc::noted_reads(),
        vec![format!("{hovered}:time-pos")],
        "and the second that frame is at is asked of the player that is showing it"
    );
    assert_eq!(
        pin_film_host::held_hover_for_a_test(),
        Some((hovered, 412.5)),
        "the hover is what the wait holds, at the second the pin's player was begun on"
    );

    let pinned = VIDEO_PID.load(Ordering::SeqCst);
    let launched = video_loop_playhead(pinned, Path::new("a-film.mkv"))
        .expect("the clock the pin's player is measured by");
    assert!(
        (launched - 412.5).abs() < 0.5,
        "the pin begins at the second the hover froze on rather than at the clock's estimate of \
         it: {launched}"
    );
    let playhead =
        pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin_playhead(&pin.transport)));
    assert!(
        playhead.is_some_and(|at| (at - 412.5).abs() < 0.5),
        "and the bar the pin is drawn against stands there too: {playhead:?}"
    );

    pin_film_host::disarm_pin_handover();
    mpv_ipc::forget_pipe(hovered);
    mpv_ipc::forget_pipe(pinned);
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    clear_restart_count();
    drop(film);
}

/// **A read that does not answer leaves the clock as the second to begin on**, and the take-up goes
/// ahead: a frozen picture this app cannot read is still a frozen picture, and the pin is begun where
/// it would have been begun before the read existed.
#[test]
fn a_read_that_does_not_answer_leaves_the_clock_as_the_second() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let hovered = 0x7FFF_FFFB;
    film.taken_up_from_a_hover(hovered);
    clear_restart_count();
    pin_film_host::disarm_pin_handover();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    mpv_ipc::remember_pipe(hovered, "a-pipe-that-says-nothing", false);

    unsafe { show_pinned_window(film.pin_hwnd) };

    assert_eq!(
        mpv_ipc::noted_reads(),
        vec![format!("{hovered}:time-pos")],
        "the question is asked, and a player that does not answer is not waited for"
    );
    assert_eq!(restart_count(), 1, "the pin's player is begun all the same");
    assert_eq!(
        pin_film_host::held_hover_for_a_test(),
        Some((hovered, 0.0)),
        "on this app's own clock, which is what the launch had before the read existed"
    );
    assert!(
        pin_film_host::pin_handover_armed(),
        "and the handover waits for that player's frame as it always does"
    );

    pin_film_host::disarm_pin_handover();
    mpv_ipc::forget_pipe(hovered);
    mpv_ipc::forget_pipe(VIDEO_PID.load(Ordering::SeqCst));
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    clear_restart_count();
    drop(film);
}

/// **The hover is retired at the swap and not before it**: it is the picture on screen for the whole
/// wait, and a player parked for the loop's retirement would be ended the moment the pin's window was
/// under the host — which is the flash this whole road is (see `pin_film_host::FrozenHover`).
#[test]
fn the_frozen_hover_is_retired_only_once_the_pin_has_the_frame() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let previous_retirement = VIDEO_RETIREMENT
        .lock()
        .ok()
        .and_then(|mut held| held.take());
    let film = StandingFilm::new();
    let hovered = 0x7FFF_FFFB;
    film.taken_up_from_a_hover(hovered);
    let hover = unsafe { StandingHover::up() };
    clear_restart_count();
    pin_film_host::disarm_pin_handover();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    mpv_ipc::remember_pipe(hovered, "the-hovers-pipe", false);
    mpv_ipc::answer_a_read(hovered, "time-pos", Some("7.000000"));
    unsafe { show_pinned_window(film.pin_hwnd) };

    // The wait goes on with the hover as the picture: nothing has ended it and nothing has parked it
    // for the loop to end.
    let drawn = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(drawn);
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_child_of_the_host(host) };
    assert!(
        matches!(
            unsafe { pin_film_host::settle_pin_handover() },
            pin_film_host::PinHandover::Waiting
        ),
        "a window with no frame proved is still a wait"
    );
    assert!(
        hover.shown(),
        "and the hover is what the user is looking at"
    );
    assert!(
        VIDEO_RETIREMENT
            .lock()
            .ok()
            .and_then(|held| *held)
            .is_none(),
        "with the hover's player held for the handover rather than parked for the loop to end"
    );
    assert_eq!(
        pin_film_host::held_hover_for_a_test(),
        Some((hovered, 7.0)),
        "still the picture on screen"
    );

    // The frame: the swap is made, and only now is the hover let go of.
    the_hosted_player_has_the_frame(drawn, "7.000000");
    assert!(
        matches!(
            unsafe { pin_film_host::settle_pin_handover() },
            pin_film_host::PinHandover::HandedOver
        ),
        "the frame the hover froze is the frame the pin has"
    );
    assert!(!hover.shown(), "so the hover goes");
    assert_eq!(
        pin_film_host::held_hover_for_a_test(),
        None,
        "and its player with it"
    );

    if let Ok(mut held) = VIDEO_RETIREMENT.lock() {
        *held = previous_retirement;
    }
    pin_film_host::disarm_pin_handover();
    mpv_ipc::forget_pipe(hovered);
    mpv_ipc::forget_pipe(VIDEO_PID.load(Ordering::SeqCst));
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    clear_restart_count();
    drop(film);
}

/// **A film the user had held stays held after the handover.** The player is begun paused only because
/// the frame has to be the hover's; a hold that was the user's is not a presentation and is not let go
/// of here (see `PinTransport::paused_at`).
#[test]
fn a_film_the_user_held_stays_held_at_the_handover() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    film.taken_up_from_a_hover(101);
    let hover = unsafe { StandingHover::up() };
    with_pin(|pin| pin.transport.held(0.0));
    clear_restart_count();
    pin_film_host::disarm_pin_handover();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    unsafe { show_pinned_window(film.pin_hwnd) };

    let drawn = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(drawn);
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_child_of_the_host(host) };
    the_hosted_player_has_the_frame(drawn, "0.000000");

    assert!(
        matches!(
            unsafe { pin_film_host::settle_pin_handover() },
            pin_film_host::PinHandover::HandedOver
        ),
        "the swap is made on the frame, held or not"
    );
    assert!(
        !mpv_ipc::noted_lines().iter().any(|line| line
            == &format!(
                "{}\n",
                r#"{"command":["set_property","pause",false],"async":true}"#
            )),
        "and a film the user was holding is not set going by it: {:?}",
        mpv_ipc::noted_lines()
    );
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin.transport.paused_at)),
        Some(0.0),
        "the pin's own transport still says the film is held, which is what the bar is drawn from"
    );
    assert!(
        !hover.shown(),
        "with the hover down like any other handover"
    );

    pin_film_host::disarm_pin_handover();
    mpv_ipc::forget_pipe(VIDEO_PID.load(Ordering::SeqCst));
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    clear_restart_count();
    drop(film);
}

/// A pin that goes away down another road — closed, or collapsed — takes the hover it froze with it:
/// there is nothing left for that player to be the picture of, and a frozen player nobody holds is a
/// process this app would never end.
#[test]
fn a_pin_that_goes_away_takes_the_frozen_hover_with_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let hovered = 0x7FFF_FFFB;
    film.taken_up_from_a_hover(hovered);
    clear_restart_count();
    pin_film_host::disarm_pin_handover();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    mpv_ipc::remember_pipe(hovered, "the-hovers-pipe", false);
    mpv_ipc::answer_a_read(hovered, "time-pos", Some("3.000000"));
    unsafe { show_pinned_window(film.pin_hwnd) };
    assert_eq!(
        pin_film_host::held_hover_for_a_test(),
        Some((hovered, 3.0)),
        "the take-up is holding the hover"
    );

    // The pin goes away: the close owns the teardown and there is no pin for the swap to be made on.
    stand_pin(None);

    assert!(
        matches!(
            unsafe { pin_film_host::settle_pin_handover() },
            pin_film_host::PinHandover::Waiting
        ),
        "a wait with no pin behind it is not a handover"
    );
    assert_eq!(
        pin_film_host::held_hover_for_a_test(),
        None,
        "and the hover it froze is ended with the pin it was being kept for"
    );

    pin_film_host::disarm_pin_handover();
    mpv_ipc::forget_pipe(hovered);
    mpv_ipc::forget_pipe(VIDEO_PID.load(Ordering::SeqCst));
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    clear_restart_count();
    drop(film);
}

/// The picture has a window of the pin's own making, standing in the band and disabled: the pin is
/// what a press on the film reaches, and the band is the one box of the pin's the film is shown in.
#[test]
fn a_pinned_film_stands_in_a_disabled_child_of_the_pins_window() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();

    assert!(
        pin_film_host::film_host().is_none(),
        "a pin with no film has no host to show one in"
    );

    assert_eq!(
        video_route(Path::new("a-film.mkv")),
        VideoRoute::Mpv,
        "the engine that plays the film"
    );
    assert!(pin_is_its_own_window(), "the pin stands in its own window");
    assert!(pinned_content().is_some(), "the band the host stands in");

    unsafe { show_pinned_window(film.pin_hwnd) };

    let host = pin_film_host::film_host().expect("the window the film is drawn in");
    assert_eq!(
        unsafe { GetParent(host) }.ok(),
        Some(film.pin_hwnd),
        "the host is a child of the pin's own window, so it travels with it and never stands on \
         the desktop of its own"
    );
    assert!(
        !unsafe { IsWindowEnabled(host) }.as_bool(),
        "a disabled host is skipped by `WindowFromPoint`, which is what makes a press on the \
         picture a press on the pin"
    );
    assert_ne!(
        unsafe { GetWindowLongPtrW(host, GWL_STYLE) } & WS_CHILD.0 as isize,
        0,
        "and it is a child by style as well as by parent"
    );
    assert_eq!(
        unsafe { window_rect(host) },
        film.box_,
        "the host is the band: a film's chrome is drawn over its picture, so the band is the whole \
         of the window"
    );

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    assert!(
        !unsafe { IsWindow(host) }.as_bool(),
        "and the host goes down with the pin rather than standing in the next pin's band"
    );
    assert!(
        pin_film_host::film_host().is_none(),
        "with the handle forgotten as well as the window destroyed"
    );
}

/// A host the player has drawn nothing into yet is a black rectangle where the band is — the
/// film the hover was showing, blanked — so it is not shown until the player has put a child of
/// its own under it, and the tick's road is what shows it then (see `take_up_the_pin_film_host`
/// and `ensure_video_window_topmost`).
#[test]
fn a_host_with_no_film_in_it_is_not_on_screen() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");

    assert!(
        !unsafe { IsWindowVisible(host) }.as_bool(),
        "the host is made where the band is and left off the screen until there is a picture in it"
    );

    let child = unsafe { a_child_of_the_host(host) };
    assert!(
        pin_film_host::the_hosted_player_is_drawing(VIDEO_PID.load(Ordering::SeqCst)),
        "the player has put a child of its own under the host"
    );

    let _ = ensure_video_window_topmost(
        film.box_.0,
        film.box_.1,
        (film.box_.2 - film.box_.0).max(1),
        (film.box_.3 - film.box_.1).max(1),
    );
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool(),
        "and the tick's own road is what shows the band once the film is drawing in it"
    );

    unsafe {
        let _ = DestroyWindow(child);
    };
}

/// A park is the host's absence and the pin's held frame is what shows through the band; the
/// unpark is the host again.
#[test]
fn a_park_hides_the_film_host_and_the_unpark_shows_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");

    // The player's own child, and the pin asked again the way a restore asks: a host the film is
    // already drawing in is the band's picture from the take-up on.
    let child = unsafe { a_child_of_the_host(host) };
    unsafe { show_pinned_window(film.pin_hwnd) };
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool(),
        "a film that is not parked is a film on screen"
    );

    hide_pinned_player_window();
    assert!(
        !unsafe { IsWindowVisible(host) }.as_bool(),
        "a park is the picture's window hidden, whichever window the player draws it in"
    );

    show_pinned_player_window();
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool(),
        "and the unpark is the same window back"
    );

    unsafe {
        let _ = DestroyWindow(child);
    };
}

/// A hosted film's park is ended by the host's own child and not by a window of the desktop's: the
/// settle hands the band back to the same host, which is the show — no player is placed, styled or
/// raised, and nothing is relaunched (see `settle_pinned_park` and `unpark_pinned_player`).
#[test]
fn a_hosted_films_park_is_settled_by_the_host_it_was_hidden_in() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");

    // The player's own child, which is what a hosted film drawing is (see
    // `the_hosted_player_is_drawing`).
    let child = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_CHILD | WS_VISIBLE,
            0,
            0,
            16,
            16,
            host,
            None,
            None,
            None,
        )
    }
    .expect("a child of the host's, in this process");

    assert!(
        park_pinned_player_for_step(film.pin_hwnd, (film.box_.0, film.box_.1)),
        "a file step covers the band and hides the film's window"
    );
    assert!(
        !unsafe { IsWindowVisible(host) }.as_bool(),
        "which for a hosted film is the host hidden"
    );
    assert!(pin_player_is_parked(), "with the cover standing");

    assert!(
        settle_pinned_park(),
        "the settle hands the band back to the child the player is drawing in"
    );
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool(),
        "which is the same host shown again"
    );
    assert!(!pin_player_is_parked(), "with the cover down");

    unsafe {
        let _ = DestroyWindow(child);
    };
}

/// The topmost road is for players placed from the outside; the host is already where the pin is,
/// and it is the band that says where the picture goes. So the box the road was handed is not the
/// host's, and the host still comes back when a park left it hidden.
#[test]
fn the_topmost_road_neither_moves_nor_raises_the_host() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    // The player's own child is the film drawing in the band (see `the_hosted_player_is_drawing`),
    // which is the picture a park hides and the topmost road hands back.
    let child = unsafe { a_child_of_the_host(host) };

    hide_pinned_player_window();
    let elsewhere = (100, 100, 700, 500);
    let placed = ensure_video_window_topmost(
        elsewhere.0,
        elsewhere.1,
        elsewhere.2 - elsewhere.0,
        elsewhere.3 - elsewhere.1,
    );

    assert!(placed, "the band is owed its picture back");
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool(),
        "and the window the band is seen through is the host"
    );
    assert_eq!(
        unsafe { window_rect(host) },
        film.box_,
        "at the pin's band rather than at the box a top-level player would have been given"
    );

    unsafe {
        let _ = DestroyWindow(child);
    };
}

/// The wait for a film's first frame is the wait for the player's own child under the host: the
/// host carries the last film's child for as long as that process lives, so the child has to be
/// the one this app began.
#[test]
fn the_film_is_drawing_when_the_player_has_a_child_of_its_own_under_the_host() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    pin_film_host::note_hosted_player(pid);

    assert!(
        !pin_film_host::the_hosted_player_is_drawing(VIDEO_PID.load(Ordering::SeqCst)),
        "a host with no player's child in it has nothing of the film to show yet"
    );

    let child = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_CHILD | WS_VISIBLE,
            0,
            0,
            16,
            16,
            host,
            None,
            None,
            None,
        )
    }
    .expect("a child of the host's, in this process");

    assert!(
        pin_film_host::the_hosted_player_is_drawing(VIDEO_PID.load(Ordering::SeqCst)),
        "a child of the player's own process under the host is the film on screen"
    );

    unsafe {
        let _ = DestroyWindow(child);
    };

    assert!(
        !pin_film_host::the_hosted_player_is_drawing(VIDEO_PID.load(Ordering::SeqCst)),
        "and the child going is the film not being drawn anywhere"
    );
}

/// A dead hosted player's host goes off the screen with it: a black rectangle over the band is the
/// pin's held frame blanked, which is exactly the band a file step has just covered (see
/// `stop_pinned_player`).
#[test]
fn a_stopped_pinned_player_leaves_no_empty_host_over_the_band() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let drawn = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(drawn);
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let child = unsafe { a_child_of_the_host(host) };
    unsafe {
        let _ = ShowWindow(host, SW_SHOWNOACTIVATE);
    }

    // The player the pin's media holds, and the name it is checked under: a noted player is
    // killed by the media's own take-down and the pid this test stands is not a process to touch
    // (see `PlayerProcess`).
    note_player_process(drawn, MPV_IMAGE_NAME);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        if let Some(media) = current.as_mut() {
            media.video_process = Some(PlayerProcess::Noted(drawn));
        }
    }

    stop_pinned_player();

    assert!(
        !unsafe { IsWindowVisible(host) }.as_bool(),
        "the host is a black rectangle once its player is gone, so it goes with the player rather \
         than standing over the frame a file step put in its place"
    );

    unsafe {
        let _ = DestroyWindow(child);
    };
}

/// A pin shown a picture after a hosted film takes the host off the band as it takes the film
/// down: the host is a black rectangle where the picture is about to be painted, and waiting for
/// the next show to destroy it leaves that rectangle over the band in the meantime (see
/// `swap_pinned_media` and `take_down_pinned_media`).
#[test]
fn a_swap_from_a_hosted_film_to_a_picture_takes_the_host_off_the_band() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let drawn = 0x7FFF_FFFA;
    film.drawn_in_the_host(drawn);
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_child_of_the_host(host) };
    unsafe {
        let _ = ShowWindow(host, SW_SHOWNOACTIVATE);
    }
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool(),
        "the film's host is the picture the pin is showing"
    );

    let mut picture = create_loading_media(320, 240);
    picture.media_type = MediaType::StaticImage;
    let swap = swap_pinned_media(PinAnswer {
        path: PathBuf::from("a-picture.png"),
        update: PinUpdate {
            content: film.box_,
            dpi: 96,
            volume: 100,
        },
        media: Some(picture),
        walk: None,
        arc: PinArc::new(),
    });

    assert!(
        matches!(swap, PinSwap::Ready(_)),
        "a picture has no player to wait for"
    );
    assert!(
        !unsafe { IsWindowVisible(host) }.as_bool(),
        "the host the film was drawn in goes off the band with the film, rather than standing as \
         a black rectangle over the picture that replaces it"
    );
}

/// A retired player is ended when its replacement has arrived, and a hosted replacement arrives
/// when its own child is under the host — not when a top-level handle is published, which a hosted
/// player never has (see `settle_video_retirement`).
#[test]
fn a_hosted_replacement_ends_the_retirement_it_took_over_from() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let previous_pid = VIDEO_PID.load(Ordering::SeqCst);
    let previous_image = video_process_image();
    let previous_retirement = VIDEO_RETIREMENT
        .lock()
        .ok()
        .and_then(|mut held| held.take());

    let film = StandingFilm::new();
    let drawn = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(drawn);
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_child_of_the_host(host) };

    // The replacement is this test's own process, recorded under the image it runs: alive as far
    // as the check is concerned, which is what makes its *arrival* the only thing that can end the
    // wait (see `is_player_pid_alive` and `note_player_process`).
    let image = std::env::current_exe()
        .ok()
        .and_then(|exe| {
            exe.file_name()
                .map(|name| name.to_string_lossy().into_owned())
        })
        .expect("this test's own executable");
    note_player_process(drawn, Box::leak(image.into_boxed_str()));
    if let Ok(mut held) = VIDEO_RETIREMENT.lock() {
        *held = Some(VideoRetirement {
            retiring: 0x7FFF_FFFC,
            replacement: drawn,
            started: Instant::now(),
        });
    }

    settle_video_retirement();

    assert!(
        VIDEO_RETIREMENT
            .lock()
            .ok()
            .and_then(|held| *held)
            .is_none(),
        "the hosted replacement's child is the film on screen, so the player it replaced is ended \
         now rather than at the ten-second bound"
    );

    if let Ok(mut held) = VIDEO_RETIREMENT.lock() {
        *held = previous_retirement;
    }
    note_player_process(previous_pid, previous_image);
}

/// A swap onto a hosted film is answered by the child its player has put under the host, not by a
/// top-level handle a hosted player never publishes: the file is installed on the child's arrival
/// rather than let go of at the cover's bound (see `PinSwapHold::settle`).
#[test]
fn a_hosted_swaps_wait_is_answered_by_the_hosts_child() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    // The child this test makes is the test's own process's, so the player the host holds has to be
    // this process as well (see `the_hosted_player_is_drawing`).
    let drawn = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(drawn);
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_child_of_the_host(host) };

    let mut media = create_loading_media(320, 240);
    media.media_type = MediaType::Video;
    media.video_process = Some(PlayerProcess::Noted(drawn));
    let mut hold = PinSwapHold {
        file: PinInstallable {
            path: PathBuf::from("the-next-film.mkv"),
            update: PinUpdate {
                content: film.box_,
                dpi: 96,
                volume: 100,
            },
            media,
            audio: None,
            walk: None,
        },
        arc: PinArc::new(),
        started: Instant::now(),
    };

    assert_eq!(
        hold.settle(),
        Some(PinSwapWait::Arrived),
        "the hosted player's own child is the film on screen, so the swap arrives instead of \
         being given up on at the cover's six hundred milliseconds"
    );
}

/// A player begun for a pinned mpv film is told the pin's own host and nothing about a window of
/// the desktop's, while a film whose pin stands in the preview's window keeps the top-level launch
/// it has (see `start_mpv_video_playback`).
#[test]
fn a_pinned_mpv_film_is_begun_into_the_pins_film_host() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");

    let path = Path::new("a-film.mkv");
    let wid = mpv_launch_host(path).expect("the host the film is drawn in");

    assert_eq!(
        wid, host.0 as usize as u32,
        "the window the player is told to draw into is the pin's own child"
    );
    assert!(
        video_launch::mpv_args(
            path,
            0,
            0,
            320,
            240,
            0.0,
            false,
            80,
            None,
            false,
            false,
            None,
            "rhp-test",
            Some(wid),
        )
        .iter()
        .any(|arg| arg == &format!("--wid={wid}")),
        "and the command line names it"
    );

    pin_stands_in_its_own_window(false);
    assert!(
        mpv_launch_host(path).is_none(),
        "a film whose pin stands in the preview's window keeps a window of the desktop's"
    );
}

/// The geometry the strips of chrome are carried in, read the way the paint and the placement read
/// it: the window's own size, and the two strips' heights as the pin settled them.
fn the_strips_geometry() -> (i32, i32, i32, i32) {
    pin_state()
        .and_then(|pinned| {
            pinned.pin().map(|pin| {
                let (width, height) = pin.window_size();
                (
                    width,
                    height,
                    pin.caption,
                    pinned_transport_height(pin.dpi, pin.transport_bar),
                )
            })
        })
        .expect("the pin the strips are carried for")
}

/// The family of the pin's own window, walked upward from `host`: every sibling above it, nearest
/// first. The strips are siblings of the host in this family; the film itself is a child of the
/// host's and has no place in it.
unsafe fn the_family_above(host: HWND) -> Vec<isize> {
    let mut above = Vec::new();
    let mut current = unsafe { GetWindow(host, GW_HWNDPREV) }.ok();
    while let Some(hwnd) = current {
        above.push(hwnd.0 as isize);
        current = unsafe { GetWindow(hwnd, GW_HWNDPREV) }.ok();
    }
    above
}

/// The film's chrome is drawn over its picture (see `band_in_the_window`), and a child composites
/// over the layered surface its parent paints — so mpv's film stood over the strips the moment it
/// began to draw, painted and interactive and invisible. The chrome therefore has windows of its
/// own, layered children of the pin's window, which composite over the host carrying the very
/// pixels the pin paints for them (see `pin_chrome_strips`).
#[test]
fn the_chrome_of_a_hosted_film_stands_over_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let caption = pin_chrome_strips::caption_strip().expect("the caption's strip of its own");
    let transport = pin_chrome_strips::transport_strip().expect("the transport's strip of its own");
    let volume = pin_chrome_strips::volume_strip().expect("the volume popup's strip of its own");
    assert!(
        !unsafe { IsWindowVisible(caption) }.as_bool()
            && !unsafe { IsWindowVisible(transport) }.as_bool()
            && !unsafe { IsWindowVisible(volume) }.as_bool(),
        "a strip comes into the world off the screen, with nothing painted in it yet"
    );

    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");

    for strip in [caption, transport, volume] {
        assert_eq!(
            unsafe { GetParent(strip) }.ok(),
            Some(film.pin_hwnd),
            "the strips are children of the pin's own window, like the host they stand over"
        );
        assert!(
            !unsafe { IsWindowEnabled(strip) }.as_bool(),
            "a disabled strip is skipped by `WindowFromPoint`, which is what keeps a press on the \
             chrome a press on the pin"
        );
        assert_ne!(
            unsafe { GetWindowLongPtrW(strip, GWL_EXSTYLE) } & WS_EX_LAYERED.0 as isize,
            0,
            "per-pixel alpha is what a strip is composited by"
        );
        assert_ne!(
            unsafe { GetWindowLongPtrW(strip, GWL_STYLE) } & WS_DISABLED.0 as isize,
            0,
            "and it is disabled by style as well as by answer"
        );
    }

    // Above the host, and nothing else of the pin's between them: making the host puts it on top
    // of the strips, so the order is re-asserted where the host is made, and the host is placed
    // with `SWP_NOZORDER` wherever the band goes after that, so the order is the one that sticks.
    assert_eq!(
        unsafe { the_family_above(host) },
        vec![caption.0 as isize, transport.0 as isize, volume.0 as isize],
        "the strips stand over the host, nearest above it, and the volume strip highest"
    );
}

/// The strips stand where the chrome is painted: the caption across the window's first rows and
/// the transport across its last, the same rows `render_pinned_preview_at` paints them into (see
/// `pinned_paint`).
#[test]
fn the_chrome_strips_of_a_hosted_film_stand_where_the_chrome_is_painted() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    with_pin(|pin| pin.transport_bar = true);
    unsafe { show_pinned_window(film.pin_hwnd) };
    let box_ = film.box_;
    let (width, height, caption_height, transport_height) = the_strips_geometry();

    let caption = pin_chrome_strips::caption_strip().expect("the caption's strip of its own");
    assert_eq!(
        unsafe { window_rect(caption) },
        (
            box_.0,
            box_.1,
            box_.0 + width,
            box_.1 + caption_height.max(1)
        ),
        "the caption's strip is the window's first rows, which is where the caption is painted"
    );

    let transport = pin_chrome_strips::transport_strip().expect("the transport's strip of its own");
    assert_eq!(
        unsafe { window_rect(transport) },
        (
            box_.0,
            box_.1 + height - transport_height.max(1),
            box_.0 + width,
            box_.1 + height
        ),
        "the transport's strip is the window's last rows, which is where the bar is painted"
    );
}

/// The strips belong to the band, and the band is what a resize changes: they follow the pin's box
/// wherever the loop is told it, at the same point the host is re-placed at (see
/// `place_the_pin_film_host`).
#[test]
fn the_chrome_strips_of_a_hosted_film_follow_the_band() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    with_pin(|pin| pin.transport_bar = true);
    unsafe { show_pinned_window(film.pin_hwnd) };

    let wider = (
        film.box_.0,
        film.box_.1,
        film.box_.0 + 400,
        film.box_.1 + 300,
    );
    with_pin(|pin| pin.content = wider);
    place_pinned_siblings();
    let (_, _, caption_height, transport_height) = the_strips_geometry();

    let caption = pin_chrome_strips::caption_strip().expect("the caption's strip of its own");
    assert_eq!(
        unsafe { window_rect(caption) },
        (wider.0, wider.1, wider.2, wider.1 + caption_height.max(1)),
        "the caption's strip is at the head of the band the pin now shows its film through"
    );

    let transport = pin_chrome_strips::transport_strip().expect("the transport's strip of its own");
    assert_eq!(
        unsafe { window_rect(transport) },
        (wider.0, wider.3 - transport_height.max(1), wider.2, wider.3),
        "and the transport's strip is at its foot"
    );
}

/// The strips are the pin's window's, not the host's: the host goes down and they stand, and a
/// park — the host hidden — leaves the chrome standing over the cover (see `pin_chrome_strips`).
#[test]
fn the_chrome_strips_outlive_the_film_host() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    with_pin(|pin| pin.transport_bar = true);
    unsafe { show_pinned_window(film.pin_hwnd) };
    let caption = pin_chrome_strips::caption_strip().expect("the strips are made with the window");
    let transport = pin_chrome_strips::transport_strip().expect("both of them");

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    assert!(
        pin_film_host::film_host().is_none(),
        "the host is gone"
    );
    assert!(
        unsafe { IsWindow(caption) }.as_bool()
            && unsafe { IsWindow(transport) }.as_bool(),
        "and the strips stand: they are the pin window's, and the next take-up places them again"
    );
    assert!(
        pin_chrome_strips::caption_strip().is_some()
            && pin_chrome_strips::transport_strip().is_some(),
        "with the handles kept as well as the windows"
    );

    // And the on-screen half of the same question: a park hides the picture, not the chrome —
    // the bitmap paints the cover, and the strips carry the chrome over it (see
    // `set_the_pin_film_host_visible`).
    let film = StandingFilm::new();
    with_pin(|pin| pin.transport_bar = true);
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let caption =
        pin_chrome_strips::caption_strip().expect("the strips are made with the window");
    let transport =
        pin_chrome_strips::transport_strip().expect("both of them");
    let _child = unsafe { a_child_of_the_host(host) };
    unsafe { show_pinned_window(film.pin_hwnd) };
    assert!(
        unsafe { IsWindowVisible(caption) }.as_bool()
            && unsafe { IsWindowVisible(transport) }.as_bool(),
        "a film that is not parked is a film the chrome stands over"
    );

    hide_pinned_player_window();
    assert!(
        unsafe { IsWindowVisible(caption) }.as_bool()
            && unsafe { IsWindowVisible(transport) }.as_bool(),
        "a park takes the picture off the screen, and the chrome stays over the cover"
    );

    show_pinned_player_window();
    assert!(
        unsafe { IsWindowVisible(caption) }.as_bool()
            && unsafe { IsWindowVisible(transport) }.as_bool(),
        "and the unpark is the picture back under the same chrome"
    );
}

/// The strips follow the chrome's own state, apart: the caption comes and goes with
/// `pin.chrome.caption` and the transport with `pin.chrome.bar`, on the clock the pin's strips are
/// already on — a strip the pin is not painting is a strip off the screen (see `PinChrome`).
#[test]
fn the_chrome_strips_of_a_hosted_film_carry_the_chrome_state() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    with_pin(|pin| pin.transport_bar = true);
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let caption = pin_chrome_strips::caption_strip().expect("the caption's strip of its own");
    let transport = pin_chrome_strips::transport_strip().expect("the transport's strip of its own");
    let _child = unsafe { a_child_of_the_host(host) };
    unsafe { show_pinned_window(film.pin_hwnd) };
    assert!(
        unsafe { IsWindowVisible(caption) }.as_bool()
            && unsafe { IsWindowVisible(transport) }.as_bool(),
        "the chrome the pin shows the strips carry"
    );

    // The caption asked away, on the repaint the ask is answered by: the pin paints no caption,
    // and the strip that carried it goes off the screen with it.
    with_pin(|pin| pin.chrome.caption = false);
    unsafe { render_pinned_preview_at(film.pin_hwnd, film.box_.0, film.box_.1) };
    assert!(
        !unsafe { IsWindowVisible(caption) }.as_bool(),
        "a caption the pin is not painting is a caption off the screen"
    );
    assert!(
        unsafe { IsWindowVisible(transport) }.as_bool(),
        "and the transport, asked for still, stays"
    );

    // The transport the same way, and both back together.
    with_pin(|pin| {
        pin.chrome.caption = true;
        pin.chrome.bar = false;
    });
    unsafe { render_pinned_preview_at(film.pin_hwnd, film.box_.0, film.box_.1) };
    assert!(
        !unsafe { IsWindowVisible(transport) }.as_bool(),
        "a bar the pin is not painting is a bar off the screen"
    );

    with_pin(|pin| pin.chrome.bar = true);
    unsafe { render_pinned_preview_at(film.pin_hwnd, film.box_.0, film.box_.1) };
    assert!(
        unsafe { IsWindowVisible(caption) }.as_bool()
            && unsafe { IsWindowVisible(transport) }.as_bool(),
        "and what the chrome asks for again, the strips carry again"
    );
}

/// The volume popup is over a hosted film's picture the way the chrome is: the panel drawn into
/// the pin's own surface is under the film's child window, so a strip of the popup's footprint
/// stands over the host while it is open — the panel visible, and the press still the pin's (see
/// `pin_chrome_strips` and `volume_popup_bounds`).
#[test]
fn the_volume_popup_of_a_hosted_film_stands_over_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    with_pin(|pin| {
        pin.transport_bar = true;
        pin.volume.open = true;
        pin.volume.level = 40;
    });
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let volume = pin_chrome_strips::volume_strip().expect("the volume popup's strip");
    let _child = unsafe { a_child_of_the_host(host) };
    unsafe { render_pinned_preview_at(film.pin_hwnd, film.box_.0, film.box_.1) };

    assert!(
        unsafe { IsWindowVisible(volume) }.as_bool(),
        "an open popup over a hosted film is a popup the strip carries"
    );
    assert_eq!(
        unsafe { the_family_above(host) }.last().copied(),
        Some(volume.0 as isize),
        "and the strip stands over the host, highest of the strips"
    );

    let popup = pinned_volume_geometry().expect("the popup the pin has open");
    let (width, height) = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_size()))
        .expect("the pin the popup is placed in");
    let bounds = pin_chrome::volume_popup_bounds(&popup, width, height);
    assert_eq!(
        unsafe { window_rect(volume) },
        (
            film.box_.0 + bounds.left,
            film.box_.1 + bounds.top,
            film.box_.0 + bounds.right,
            film.box_.1 + bounds.bottom,
        ),
        "the strip is the popup's footprint, shadow and edge margin and all"
    );

    with_pin(|pin| pin.volume.open = false);
    unsafe { render_pinned_preview_at(film.pin_hwnd, film.box_.0, film.box_.1) };
    assert!(
        !unsafe { IsWindowVisible(volume) }.as_bool(),
        "and a popup put away is a strip off the screen again"
    );
}

/// The strips belong to the pin's own window rather than to its film's host: they are made with
/// the window, before any host, and they are there on every road — a hosted film, a park, a
/// cover, FFmpeg's player, or a picture this app draws itself (see `pin_chrome_strips`).
#[test]
fn the_strips_belong_to_the_pins_window_not_its_film_host() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    // No take-up: the host is made when the pin is shown, and the strips are already here.
    let film = StandingFilm::new();

    assert!(
        pin_film_host::film_host().is_none(),
        "no take-up, no host — and the strips are here anyway"
    );
    for strip in [
        pin_chrome_strips::caption_strip().expect("the caption's strip, made with the window"),
        pin_chrome_strips::transport_strip().expect("the transport's, made with it"),
    ] {
        assert_eq!(
            unsafe { GetParent(strip) }.ok(),
            Some(film.pin_hwnd),
            "a strip is a child of the pin's own window"
        );
        assert!(
            !unsafe { IsWindowEnabled(strip) }.as_bool(),
            "a disabled strip is skipped by `WindowFromPoint`, so a press on the chrome stays a \
             press on the pin"
        );
        assert_ne!(
            unsafe { GetWindowLongPtrW(strip, GWL_EXSTYLE) } & WS_EX_LAYERED.0 as isize,
            0,
            "per-pixel alpha is what a strip is composited by"
        );
        assert_ne!(
            unsafe { GetWindowLongPtrW(strip, GWL_STYLE) } & WS_DISABLED.0 as isize,
            0,
            "and it is disabled by style as well as by answer"
        );
        assert!(
            !unsafe { IsWindowVisible(strip) }.as_bool(),
            "a strip comes into the world off the screen, with nothing painted in it yet"
        );
    }
}

/// The pin's bitmap carries no chrome any more: the caption's and the transport's rows are the
/// pin's own alpha-1 claim on the point, and the chrome lives in the strips (see
/// `pin_chrome_strips`).
#[test]
fn the_pins_bitmap_carries_no_chrome_only_the_claim_on_its_rows() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    with_pin(|pin| pin.transport_bar = true);

    // The paint's own surface, read out of a window of the test's own rather than the pin's:
    // what is asserted is the pixels the paint composed, not the window they were blitted to.
    let hwnd = HWND(0x1000 as *mut _);
    unsafe { render_pinned_preview_at(hwnd, film.box_.0, film.box_.1) };

    let (width, height, caption, transport) = pin_state()
        .and_then(|pinned| {
            pinned.pin().map(|pin| {
                let (width, height) = pin.window_size();
                (
                    width,
                    height,
                    pin.caption,
                    pinned_transport_height(pin.dpi, pin.transport_bar),
                )
            })
        })
        .expect("the pin the paint read");
    assert!(
        caption > 0 && transport > 0,
        "the pin owes both strips, so both rows are asserted"
    );
    let bits = ensure_layered_surface(hwnd.0 as isize, width as u32, height as u32)
        .expect("the surface the paint drew into");
    let surface =
        unsafe { std::slice::from_raw_parts(bits, width as usize * height as usize * 4) };
    let pixel = |row: usize, column: usize| -> &[u8] {
        &surface[(row * width as usize + column) * 4..][..4]
    };
    for row in 0..caption as usize {
        for column in 0..width as usize {
            assert_eq!(
                pixel(row, column),
                &[1, 1, 1, 1],
                "row {row} column {column}: the caption's rows are the pin's claim, not its chrome"
            );
        }
    }
    for row in (height - transport) as usize..height as usize {
        for column in 0..width as usize {
            assert_eq!(
                pixel(row, column),
                &[1, 1, 1, 1],
                "row {row} column {column}: the transport's rows are the pin's claim, not its bar"
            );
        }
    }

    // And the chrome was still painted — into the strips' own surfaces, which is where the sync
    // carries it onto the screen from.
    CHROME_SURFACE.with(|cell| {
        assert_eq!(
            cell.borrow()
                .as_ref()
                .map(|surface| (surface.width, surface.height)),
            Some((width as u32, caption as u32)),
            "the caption is painted once, for the strips"
        );
    });
    TRANSPORT_SURFACE.with(|cell| {
        assert_eq!(
            cell.borrow()
                .as_ref()
                .map(|surface| (surface.width, surface.height)),
            Some((width as u32, transport as u32)),
            "and the bar with it"
        );
    });
}

/// A film with no host — FFmpeg's player draws in a window of the desktop's — carries its chrome
/// in the strips too: the bitmap's chrome rows are the pin's claim, and the strips stand over
/// them (see `pin_chrome_strips`).
#[test]
fn a_film_without_a_host_carries_its_chrome_in_the_strips() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    with_pin(|pin| pin.transport_bar = true);
    unsafe { show_pinned_window(film.pin_hwnd) };
    // The top-level road: the player is a window of the desktop's, so the pin keeps no host.
    unsafe { pin_film_host::take_down_the_pin_film_host() };
    assert!(
        pin_film_host::film_host().is_none(),
        "no host stands in this pin's band"
    );

    unsafe { render_pinned_preview_at(film.pin_hwnd, film.box_.0, film.box_.1) };
    let caption = pin_chrome_strips::caption_strip().expect("the caption's strip");
    let transport = pin_chrome_strips::transport_strip().expect("the transport's");
    assert!(
        unsafe { IsWindowVisible(caption) }.as_bool()
            && unsafe { IsWindowVisible(transport) }.as_bool(),
        "the chrome the pin shows the strips carry, host or no host"
    );
}

/// A park paints the cover without the chrome, and the chrome stays in its strips over it: the
/// look is the parked film with its caption and its bar, and the bitmap under them is the frame
/// the park is holding (see `park_pinned_player_inner`).
#[test]
fn a_park_keeps_the_chrome_in_the_strips_over_the_cover() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    forget_video_frame();
    let film = StandingFilm::new();
    with_pin(|pin| {
        pin.transport_bar = true;
        pin.parked = true;
    });
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_child_of_the_host(host) };
    unsafe { show_pinned_window(film.pin_hwnd) };
    hide_pinned_player_window();

    // The bitmap: the cover and nothing else — no chrome composited into it.
    let hwnd = HWND(0x1000 as *mut _);
    unsafe { render_pinned_preview_at(hwnd, film.box_.0, film.box_.1) };
    let (width, height, caption, transport) = pin_state()
        .and_then(|pinned| {
            pinned.pin().map(|pin| {
                let (width, height) = pin.window_size();
                (
                    width,
                    height,
                    pin.caption,
                    pinned_transport_height(pin.dpi, pin.transport_bar),
                )
            })
        })
        .expect("the pin the paint read");
    let bits = ensure_layered_surface(hwnd.0 as isize, width as u32, height as u32)
        .expect("the surface the paint drew into");
    let surface =
        unsafe { std::slice::from_raw_parts(bits, width as usize * height as usize * 4) };
    let pixel = |row: usize, column: usize| -> &[u8] {
        &surface[(row * width as usize + column) * 4..][..4]
    };
    for row in (0..caption as usize).chain((height - transport) as usize..height as usize) {
        for column in 0..width as usize {
            assert_eq!(
                pixel(row, column),
                &[0, 0, 0, 255],
                "row {row} column {column}: the cover stands in the chrome's rows, not the chrome"
            );
        }
    }

    // And the chrome over it, in the strips.
    let caption_strip = pin_chrome_strips::caption_strip().expect("the caption's strip");
    let transport_strip = pin_chrome_strips::transport_strip().expect("the transport's");
    assert!(
        unsafe { IsWindowVisible(caption_strip) }.as_bool()
            && unsafe { IsWindowVisible(transport_strip) }.as_bool(),
        "a park hides the picture, and the strips keep the chrome over the cover"
    );
}

/// Only a really hovered pin shows its chrome: the geometric question is asked first, and then
/// the window under the pointer — the host and the strips are disabled, so `WindowFromPoint`
/// skips them and the pin's own window comes back wherever the pointer is genuinely over the
/// pin (see `refresh_pin_chrome`).
#[test]
fn a_hover_over_the_pin_itself_asks_its_chrome_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    // The step's own state: hidden, with no arrival armed (see `take_up_pinned_window`) — on a
    // film's pin, whose bar the pointer below is near.
    with_pin(|pin| {
        pin.chrome = PinChrome::hidden();
        pin.transport_bar = true;
    });
    let hwnd = pin_real_window::real_pin_window().expect("the pin's own window");
    stand_window_at_point_in(Some(hwnd.0 as isize));

    let window = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_box()))
        .expect("a pin up");
    let changed = pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(refresh_pin_chrome(
            pin,
            Instant::now(),
            Some((window.0 + 5, window.3 - 1)),
        ))
    });
    let bar = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.chrome.bar))
        .unwrap_or(false);

    stand_window_at_point_in(None);
    drop(film);
    assert_eq!(
        changed,
        Some(true),
        "a pointer over the pin's own window is a hover, which shows the bar again"
    );
    assert!(bar, "and the bar is what the hover asked back");
}

/// The other half: a pointer geometrically near the pin but over another window — Notepad
/// standing over it — is not a hover, so the chrome stays hidden (see `refresh_pin_chrome`).
#[test]
fn a_pointer_over_a_window_covering_the_pin_shows_no_chrome() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let _film = StandingFilm::new();
    with_pin(|pin| {
        pin.chrome = PinChrome::hidden();
        pin.transport_bar = true;
    });
    // Somebody else's window under the pointer: what `WindowFromPoint` answers where a covering
    // window stands over the pin.
    stand_window_at_point_in(Some(0x2000));

    let window = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_box()))
        .expect("a pin up");
    let changed = pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(refresh_pin_chrome(
            pin,
            Instant::now(),
            Some((window.0 + 5, window.3 - 1)),
        ))
    });
    let bar = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.chrome.bar))
        .unwrap_or(true);

    stand_window_at_point_in(None);
    assert_eq!(
        changed,
        Some(false),
        "an occluded pointer asks for nothing, so no repaint is owed"
    );
    assert!(!bar, "and the bar stays hidden under the covering window");
}

/// The whole of the chrome rule on a pin's own window: it starts hidden on every spawn, the
/// pointer over the pin asks the strip near it back, and the pointer leaving puts it away again —
/// the strips follow the pin's chrome in the same pass (see `refresh_pin_chrome` and
/// `sync_the_chrome_strips`).
#[test]
fn the_chrome_stays_hidden_at_spawn_and_follows_the_pointer_over_the_pin() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    with_pin(|pin| {
        pin.chrome = PinChrome::hidden();
        pin.transport_bar = true;
    });
    unsafe { show_pinned_window(film.pin_hwnd) };

    let caption = pin_chrome_strips::caption_strip().expect("the caption's strip");
    let transport = pin_chrome_strips::transport_strip().expect("the transport's strip");
    assert!(
        !unsafe { IsWindowVisible(caption) }.as_bool()
            && !unsafe { IsWindowVisible(transport) }.as_bool(),
        "a pin comes up with neither strip on the screen"
    );

    // The pointer over the pin's own window: the window under it is the pin, so this is a hover.
    let hwnd = pin_real_window::real_pin_window().expect("the pin's own window");
    stand_window_at_point_in(Some(hwnd.0 as isize));
    let window = film.box_;
    pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(refresh_pin_chrome(
            pin,
            Instant::now(),
            Some((window.0 + 5, window.1)),
        ))
    });
    unsafe { render_pinned_preview_at(film.pin_hwnd, film.box_.0, film.box_.1) };
    assert!(
        unsafe { IsWindowVisible(caption) }.as_bool(),
        "a hover asks the caption back"
    );
    assert!(
        !unsafe { IsWindowVisible(transport) }.as_bool(),
        "and only the strip the pointer is near: the bar at the foot is not it"
    );

    // And the pointer leaving the pin puts it away in the same pass.
    pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(refresh_pin_chrome(
            pin,
            Instant::now(),
            Some((-1000, -1000)),
        ))
    });
    unsafe { render_pinned_preview_at(film.pin_hwnd, film.box_.0, film.box_.1) };
    assert!(
        !unsafe { IsWindowVisible(caption) }.as_bool(),
        "the caption goes with the pointer"
    );
    stand_window_at_point_in(None);
    drop(film);
}

/// A pin that went down with its chrome up leaves the strip's own `WS_VISIBLE` set, and the next
/// pin's first paint, still behind its hidden window, is what must put it away: asking
/// `IsWindowVisible` there answers for the hidden parent rather than for the strip, so the stale
/// strip would be revealed with the window and flash the previous pin's chrome (see
/// `set_the_strip_at`).
#[test]
fn a_strip_left_up_by_a_hidden_pin_is_not_revealed_when_the_next_one_appears() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    with_pin(|pin| pin.chrome = PinChrome::always());
    unsafe { show_pinned_window(film.pin_hwnd) };
    unsafe { render_pinned_preview_at(film.pin_hwnd, film.box_.0, film.box_.1) };
    let caption = pin_chrome_strips::caption_strip().expect("the caption's strip");
    assert!(
        unsafe { GetWindowLongPtrW(caption, GWL_STYLE) } & WS_VISIBLE.0 as isize != 0,
        "the pin came up with its chrome"
    );

    // The pin goes down the way every road takes it down: only the window is hidden, so the
    // strip's own bit is still set behind it.
    unsafe { pin_real_window::hide_real_pin_window() };
    with_pin(|pin| pin.chrome = PinChrome::hidden());
    unsafe { render_pinned_preview_at(film.pin_hwnd, film.box_.0, film.box_.1) };

    assert!(
        unsafe { GetWindowLongPtrW(caption, GWL_STYLE) } & WS_VISIBLE.0 as isize == 0,
        "a strip the hidden pin left up is put away by the paint that asks for no chrome"
    );
    // And so the next show has nothing stale to reveal.
    unsafe { let _ = ShowWindow(film.pin_hwnd, SW_SHOWNOACTIVATE); };
    assert!(
        !unsafe { IsWindowVisible(caption) }.as_bool(),
        "a pin coming back up reveals no strip the pointer has not asked about"
    );
    drop(film);
}

/// A step's arrival is the film's frame and not its window: the player's shown child is a band
/// the chrome would stand over black, so the strips wait for the player's own `vo-configured` and
/// the pointer near them asks for nothing until then. The latch that saves the pipe read is per
/// player, so a new player proves its own frame (see `the_hosted_frame_is_up`).
#[test]
fn the_chrome_waits_for_the_hosted_films_frame_not_its_window() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    // The window the player has made, with its video output not configured for a frame yet.
    let _child = unsafe { a_child_of_the_host(host) };
    with_pin(|pin| {
        pin.chrome = PinChrome::hidden();
        pin.transport_bar = true;
    });
    stand_window_at_point_in(Some(film.pin_hwnd.0 as isize));
    let window = film.box_;
    let caption_cursor = Some((window.0 + 5, window.1));
    let refresh = |cursor| {
        pin_state().and_then(|mut pinned| {
            let pin = pinned.pin_mut()?;
            Some(refresh_pin_chrome(pin, Instant::now(), cursor))
        })
    };

    assert!(
        pin_film_host::the_hosted_player_is_drawing(pid),
        "the player has a window under the host"
    );
    assert!(
        !pin_film_host::the_hosted_frame_is_up(pid),
        "which is not a frame it has drawn"
    );
    refresh(caption_cursor);
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| (pin.chrome.caption, pin.chrome.bar))),
        Some((false, false)),
        "a pointer on the caption asks for nothing over a window with no picture"
    );

    // The player's own word that the frame is drawn, and the pointer asks the caption back.
    the_hosted_player_has_the_frame(pid, "0.000000");
    assert_eq!(
        refresh(caption_cursor),
        Some(true),
        "the tick the frame is up is the tick the chrome may come back"
    );
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| (pin.chrome.caption, pin.chrome.bar))),
        Some((true, false)),
        "with the caption showing and the bar left to its own end"
    );

    // The latch answers for this player and not the next: a new player is begun with no frame of
    // its own, and the unasked pipe read is what says so.
    mpv_ipc::answer_a_read(pid, "vo-configured", None);
    pin_film_host::note_hosted_player(pid);
    assert!(
        !pin_film_host::the_hosted_frame_is_up(pid),
        "a player just begun has not inherited the last one's frame"
    );

    stand_window_at_point_in(None);
    mpv_ipc::clear_reads();
    mpv_ipc::clear_noted();
    drop(film);
}

/// The gate is for a standing pin only: a park holds its frame over the band, so its chrome stays
/// over that held frame while the incoming player has a window and no picture yet.
#[test]
fn a_parked_pin_keeps_its_chrome_before_the_new_films_frame() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_child_of_the_host(host) };
    with_pin(|pin| {
        pin.parked = true;
        pin.chrome = PinChrome::hidden();
        pin.transport_bar = true;
    });
    stand_window_at_point_in(Some(film.pin_hwnd.0 as isize));
    assert!(pin_film_host::the_hosted_player_is_drawing(pid));
    assert!(!pin_film_host::the_hosted_frame_is_up(pid));

    let window = film.box_;
    let shown = pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(refresh_pin_chrome(
            pin,
            Instant::now(),
            Some((window.0 + 5, window.1)),
        ))
    });
    assert_eq!(shown, Some(true), "a park's held frame keeps its chrome");
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| (pin.chrome.caption, pin.chrome.bar))),
        Some((true, false)),
        "with the caption over the cover while the new film has no picture"
    );

    stand_window_at_point_in(None);
    drop(film);
}

/// A pin taken up on a hover's film waits for the hosted player's first frame before it
/// appears: the take-up relaunches the adopted film into the host in the background, the hover
/// stays up and playing as the placeholder meanwhile, and the pin's window and its strips stay
/// down with the chrome hidden until the film is really drawing (see `settle_pin_handover`).
#[test]
fn a_pin_taken_up_on_a_hover_waits_for_the_first_frame() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    // Taken up from a hover: the film is the hover's player's, drawing in a window of the
    // desktop's rather than in the host's, so the take-up relaunches it into the host behind
    // the hover that is still up (see `begin_the_adopted_film_in_the_host`).
    film.taken_up_from_a_hover(101);
    let hover = unsafe { StandingHover::up() };
    // A first pin, whose arrival window is open: the take-up must not spend it over nothing.
    with_pin(|pin| pin.chrome = PinChrome::on_arrival(Instant::now()));
    forget_pin_park_swap();
    clear_park_trace();
    clear_restart_count();
    clear_pending_pinned_relaunch();
    pin_film_host::disarm_pin_handover();

    unsafe { show_pinned_window(film.pin_hwnd) };

    // No cover before the showing: nothing is captured, painted or hidden, because the hover
    // is the placeholder and it never comes down until the handover.
    assert_eq!(
        park_trace(),
        Vec::<&'static str>::new(),
        "the take-up raises no cover: the hover stays up as the placeholder"
    );
    assert_eq!(
        restart_count(),
        1,
        "but the adopted film is still relaunched into the host behind it"
    );
    assert!(
        !pin_player_is_parked(),
        "with no park standing over the relaunch"
    );
    assert!(
        !unsafe { IsWindowVisible(film.pin_hwnd) }.as_bool(),
        "and the pin's own window stays down until the hosted player draws"
    );
    assert!(
        hover.shown(),
        "while the hover stays up and playing as the placeholder"
    );
    assert!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.chrome.caption)) == Some(false),
        "with the chrome hidden rather than arriving over nothing"
    );
    assert!(
        pin_film_host::pin_handover_armed(),
        "until the handover the tick performs once the first frame lands"
    );
    for strip in [
        pin_chrome_strips::caption_strip(),
        pin_chrome_strips::transport_strip(),
    ]
    .into_iter()
    .flatten()
    {
        assert!(
            !unsafe { IsWindowVisible(strip) }.as_bool(),
            "and no strip shows regardless of the chrome state"
        );
    }

    pin_film_host::disarm_pin_handover();
    forget_pin_park_swap();
    clear_pending_pinned_relaunch();
    clear_restart_count();
    clear_park_trace();
    drop(film);
}

/// A hidden child is not a film drawing: the host carries the last player's child for as long
/// as that process lives, and a player that has not put a visible child of its own under the
/// host has nothing on screen yet — so the host stays off the screen even though a process of
/// its own exists (see `the_hosted_player_is_drawing`).
#[test]
fn a_hidden_child_is_not_a_film_drawing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let drawn = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(drawn);
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");

    // A child mpv made but has not shown yet: created hidden, the way a player whose first
    // frame has not arrived leaves it (see `host_has_a_child_of`).
    let hidden = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_CHILD,
            0,
            0,
            16,
            16,
            host,
            None,
            None,
            None,
        )
    }
    .expect("a hidden child of the host's, in this process");

    assert!(
        !pin_film_host::the_hosted_player_is_drawing(VIDEO_PID.load(Ordering::SeqCst)),
        "a hidden child is not a first frame: the engine's own report is a window it shows, \
         not one it made"
    );

    let _ = ensure_video_window_topmost(
        film.box_.0,
        film.box_.1,
        (film.box_.2 - film.box_.0).max(1),
        (film.box_.3 - film.box_.1).max(1),
    );
    assert!(
        !unsafe { IsWindowVisible(host) }.as_bool(),
        "and the tick's road leaves the host off the screen until that frame arrives"
    );

    unsafe {
        let _ = DestroyWindow(hidden);
    };
}

/// A hosted band with no film in it is opaque, never a hole: the pin's bitmap carries the
/// cover while the player has no first frame, so chrome up over an empty host is chrome over
/// a picture this app owns rather than over the desktop (see `render_pinned_preview_at`).
#[test]
fn a_hosted_band_with_no_film_in_it_is_opaque_never_a_hole() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    unsafe { show_pinned_window(film.pin_hwnd) };
    // No child: the player has put nothing under the host yet, so there is no first frame.
    assert!(
        !pin_film_host::the_hosted_player_is_drawing(VIDEO_PID.load(Ordering::SeqCst)),
        "no child is no frame"
    );

    let hwnd = film.pin_hwnd;
    unsafe { render_pinned_preview_at(hwnd, film.box_.0, film.box_.1) };
    let (width, height) = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_size()))
        .expect("a pin to have a size");
    let bits = ensure_layered_surface(hwnd.0 as isize, width as u32, height as u32)
        .expect("the surface the paint drew into");
    let painted =
        unsafe { std::slice::from_raw_parts(bits, width as usize * height as usize * 4) };
    let (band_top, band_height) = pinned_band_rows(
        height,
        pinned_caption_height(96, None),
        0,
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.overlay)) == Some(true),
    );
    let row = (band_top.max(0) as usize + band_height as usize / 2).min(height as usize - 1);
    let pixel = &painted[row * width as usize * 4..row * width as usize * 4 + 4];
    assert_eq!(
        pixel[3], 255,
        "a band with no film in it is opaque, never a hole in the desktop (row {row}: {pixel:?})"
    );
}

/// A step holds the old cover until the new film's first frame: parked over a band with no
/// visible child, the settle keeps the frame standing; the tick the child arrives hands the
/// band back with nothing transparent in between (see `settle_pinned_park`).
#[test]
fn a_step_holds_its_cover_until_the_new_films_first_frame() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");

    // The outgoing frame, stood where a real capture would have put it: there is no player on
    // the machine these run on, and the capture answers false rather than blanking a good
    // frame (see `hold_video_window_frame`).
    let red = [0u8, 0, 255, 255].repeat(64);
    stand_video_frame_for_a_test(red, 8, 8);
    assert!(
        park_pinned_player_for_step(film.pin_hwnd, (film.box_.0, film.box_.1)),
        "a file step covers the band it is leaving"
    );
    assert!(pin_player_is_parked(), "with the cover standing");
    assert!(held_video_frame().is_some(), "holding the outgoing frame");
    assert!(
        !unsafe { IsWindowVisible(host) }.as_bool(),
        "and the host hidden behind it"
    );

    // No visible child yet: the replacement has put nothing on screen, so the settle holds.
    assert!(
        !settle_pinned_park(),
        "with no first frame the band is not handed back"
    );
    assert!(pin_player_is_parked(), "so the cover stands");
    assert!(held_video_frame().is_some(), "holding the frame it stood over");

    // A child made but not shown yet is still no first frame: the settle keeps holding.
    let hidden = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_CHILD,
            0,
            0,
            16,
            16,
            host,
            None,
            None,
            None,
        )
    }
    .expect("a hidden child of the host's, in this process");
    assert!(
        !pin_film_host::the_hosted_player_is_drawing(VIDEO_PID.load(Ordering::SeqCst)),
        "a hidden child is a player still opening, not a film"
    );
    assert!(
        !settle_pinned_park(),
        "so the cover stands over a window that has drawn nothing"
    );
    assert!(pin_player_is_parked(), "holding the outgoing frame still");
    assert!(held_video_frame().is_some(), "rather than a transparent gap");
    unsafe {
        let _ = DestroyWindow(hidden);
    };

    // The new film's first frame: a visible child of the player's own process under the host.
    let _child = unsafe { a_child_of_the_host(host) };
    assert!(
        pin_film_host::the_hosted_player_is_drawing(VIDEO_PID.load(Ordering::SeqCst)),
        "a visible child is the film on screen"
    );
    assert!(
        settle_pinned_park(),
        "the tick the first frame lands hands the band back"
    );
    assert!(!pin_player_is_parked(), "with the cover down");
    assert!(held_video_frame().is_none(), "and the frame given up with it");
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool(),
        "which is the same host shown again, already drawing"
    );

    forget_video_frame();
    forget_resume_frame();
    forget_pin_park_swap();
}

/// A window of the desktop's standing in for the hover's: shown, so the take-up it hands over
/// from is a hover still up — destroyed with the guard, and the previous hover handle put back.
struct StandingHover {
    hwnd: HWND,
    previous: isize,
}

impl StandingHover {
    unsafe fn up() -> StandingHover {
        let previous = PREVIEW_HWND.load(Ordering::SeqCst);
        let hwnd = unsafe {
            CreateWindowExW(
                WINDOW_EX_STYLE::default(),
                w!("STATIC"),
                w!(""),
                WS_POPUP | WS_VISIBLE,
                NOWHERE,
                NOWHERE,
                320,
                240,
                None,
                None,
                None,
                None,
            )
        }
        .expect("a window for the hover to be up in");
        PREVIEW_HWND.store(hwnd.0 as isize, Ordering::SeqCst);
        StandingHover { hwnd, previous }
    }

    fn shown(&self) -> bool {
        unsafe { IsWindowVisible(self.hwnd).as_bool() }
    }
}

impl Drop for StandingHover {
    fn drop(&mut self) {
        PREVIEW_HWND.store(self.previous, Ordering::SeqCst);
        unsafe {
            let _ = DestroyWindow(self.hwnd);
        }
    }
}

/// A live process standing in for the player the take-up began: mpv is not on the machine these
/// run on, so the relaunch the take-up begins comes back with no handle — and a handover with no
/// player behind it is a wait that is already over rather than one still going on.
struct LiveStandInPlayer;

impl LiveStandInPlayer {
    // The guard's `Drop` kills and reaps the sleeper, which the lint cannot see from the spawn.
    #[allow(clippy::zombie_processes)]
    fn stand() -> LiveStandInPlayer {
        use std::os::windows::process::CommandExt;
        let child = std::process::Command::new("cmd")
            .args(["/c", "ping", "-n", "30", "127.0.0.1"])
            .stdin(std::process::Stdio::null())
            .stdout(std::process::Stdio::null())
            .stderr(std::process::Stdio::null())
            .creation_flags(0x08000000)
            .spawn()
            .expect("a live stand-in for the player the take-up began");
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            if let Some(media) = current.as_mut() {
                media.video_process = Some(PlayerProcess::Started(child));
            }
        }
        LiveStandInPlayer
    }
}

impl Drop for LiveStandInPlayer {
    fn drop(&mut self) {
        // Killed and reaped rather than dropped: a `Child` left to its handle outlives the
        // kill by however long the process takes to die, and a 30-second sleeper is a stray
        // process on a machine running the suite.
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            if let Some(media) = current.as_mut() {
                if let Some(PlayerProcess::Started(mut child)) = media.video_process.take() {
                    let _ = child.kill();
                    let _ = child.wait();
                }
            }
        }
    }
}

/// The handover is the only visible swap: the pin's window appears with the film already drawing
/// in it and the chrome hidden over that film, and the hover comes down in the same pass.
#[test]
fn the_handover_shows_the_pin_over_the_drawn_film_and_takes_the_hover_down() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    film.taken_up_from_a_hover(101);
    let hover = unsafe { StandingHover::up() };
    with_pin(|pin| pin.chrome = PinChrome::on_arrival(Instant::now()));
    clear_park_trace();
    clear_restart_count();
    clear_pending_pinned_relaunch();
    pin_film_host::disarm_pin_handover();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    unsafe { show_pinned_window(film.pin_hwnd) };
    assert!(
        pin_film_host::pin_handover_armed(),
        "the take-up arms the handover"
    );

    // The film's first frame: a shown child of this process under the host, owned by the player
    // the pin is holding (see `the_hosted_player_is_drawing`).
    let drawn = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(drawn);
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_child_of_the_host(host) };

    // And the player's own word that the frame it is standing on is the one the hover froze on: a
    // window under the host is not a picture, and the swap waits for both (see
    // `the_hosted_frame_is_presented`).
    the_hosted_player_has_the_frame(drawn, "0.000000");

    assert!(
        matches!(
            unsafe { pin_film_host::settle_pin_handover() },
            pin_film_host::PinHandover::HandedOver
        ),
        "the tick the first frame lands hands the band over"
    );
    assert!(
        unsafe { IsWindowVisible(film.pin_hwnd) }.as_bool(),
        "and the pin's window is the swap the user sees"
    );
    assert!(
        !hover.shown(),
        "with the hover coming down in the same pass, its film already replaced"
    );
    let chrome = pin_state().and_then(|pinned| {
        pinned.pin().map(|pin| {
            (
                pin.chrome.caption,
                pin.chrome.bar,
                pin.chrome.until.is_some(),
            )
        })
    });
    assert_eq!(
        chrome,
        Some((false, false, false)),
        "and the pin comes up hidden over the live film: only the pointer asks a strip back"
    );
    for strip in [
        pin_chrome_strips::caption_strip(),
        pin_chrome_strips::transport_strip(),
    ]
    .into_iter()
    .flatten()
    {
        assert!(
            !unsafe { IsWindowVisible(strip) }.as_bool(),
            "and no strip shows, whatever the film is drawing behind it"
        );
    }
    assert!(
        !pin_film_host::pin_handover_armed(),
        "a handover fires once: nothing is left armed behind the shown pin"
    );
    assert_eq!(
        pin_film_host::held_hover_for_a_test(),
        None,
        "and the hover the take-up froze is let go of rather than left frozen behind the pin"
    );
    assert!(
        mpv_ipc::noted_lines().iter().any(|line| line
            == &format!(
                "{}\n",
                r#"{"command":["set_property","pause",false],"async":true}"#
            )),
        "the pin's film is playing from the frame it was frozen on, rather than standing on it \
         behind a pin that looks alive: {:?}",
        mpv_ipc::noted_lines()
    );

    pin_film_host::disarm_pin_handover();
    clear_pending_pinned_relaunch();
    clear_restart_count();
    clear_park_trace();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    drop(film);
}

/// **A window under the host is not a picture**, and the swap waits for the player's own word that
/// the frame is drawn: the second it is standing at, and a video output configured for it.
#[test]
fn a_window_alone_is_not_a_handover() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    film.taken_up_from_a_hover(101);
    let hover = unsafe { StandingHover::up() };
    clear_park_trace();
    clear_restart_count();
    clear_pending_pinned_relaunch();
    pin_film_host::disarm_pin_handover();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    unsafe { show_pinned_window(film.pin_hwnd) };

    let drawn = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(drawn);
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_child_of_the_host(host) };

    // A player that has said nothing: a window with a box, and no frame this app can prove.
    mpv_ipc::remember_pipe(drawn, "the-hosted-players-pipe", false);
    assert!(
        matches!(
            unsafe { pin_film_host::settle_pin_handover() },
            pin_film_host::PinHandover::Waiting
        ),
        "a window the player has made is not a frame it has drawn"
    );
    assert!(
        !unsafe { IsWindowVisible(film.pin_hwnd) }.as_bool() && hover.shown(),
        "so the pin stays down and the hover's own frozen frame stays the picture"
    );

    // A second the hover never froze on is a frame that is not the one on screen, and the video
    // output has not configured anything yet.
    mpv_ipc::answer_a_read(drawn, "time-pos", Some("900.000000"));
    assert!(
        matches!(
            unsafe { pin_film_host::settle_pin_handover() },
            pin_film_host::PinHandover::Waiting
        ),
        "a player standing somewhere else in the film has not the frame the swap is made on"
    );

    // And the frame at that second with no video output configured for it is still not the picture:
    // the position is where the file is, and the output is what has been drawn there.
    mpv_ipc::answer_a_read(drawn, "time-pos", Some("0.000000"));
    mpv_ipc::answer_a_read(drawn, "vo-configured", None);
    assert!(
        matches!(
            unsafe { pin_film_host::settle_pin_handover() },
            pin_film_host::PinHandover::Waiting
        ),
        "a frame the video output has not configured is not a frame on screen"
    );
    assert!(hover.shown(), "and the hover is the picture for all of it");

    mpv_ipc::answer_a_read(drawn, "vo-configured", Some("true"));
    assert!(
        matches!(
            unsafe { pin_film_host::settle_pin_handover() },
            pin_film_host::PinHandover::HandedOver
        ),
        "the tick both facts hold is the tick the swap is made"
    );

    pin_film_host::disarm_pin_handover();
    clear_pending_pinned_relaunch();
    clear_restart_count();
    clear_park_trace();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    drop(film);
}

/// No handover before the first frame: the pin stays down, the hover stays up, and the chrome
/// stays hidden no matter what the pointer does.
#[test]
fn no_handover_before_the_first_frame() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    film.taken_up_from_a_hover(101);
    let hover = unsafe { StandingHover::up() };
    with_pin(|pin| pin.chrome = PinChrome::on_arrival(Instant::now()));
    clear_park_trace();
    clear_restart_count();
    clear_pending_pinned_relaunch();
    pin_film_host::disarm_pin_handover();
    unsafe { show_pinned_window(film.pin_hwnd) };
    // A player behind the wait, so the wait is still a wait: mpv is not on this machine, and a
    // relaunch that began nothing is a wait already over rather than one going on.
    let _player = LiveStandInPlayer::stand();

    assert!(
        matches!(
            unsafe { pin_film_host::settle_pin_handover() },
            pin_film_host::PinHandover::Waiting
        ),
        "a hosted player still opening is not a film yet"
    );
    assert!(
        !unsafe { IsWindowVisible(film.pin_hwnd) }.as_bool(),
        "so the pin's window stays down"
    );
    assert!(hover.shown(), "and the hover stays up and playing");
    assert!(
        pin_film_host::pin_handover_armed(),
        "with the handover still armed behind them"
    );

    // The pointer over the pin's box while it is still down: near enough to ask for the chrome,
    // which stays hidden because the frame it belongs to is not on screen yet.
    pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        let now = Instant::now();
        Some(refresh_pin_chrome(
            pin,
            now,
            Some((NOWHERE + 10, NOWHERE + 10)),
        ))
    });
    assert!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| (pin.chrome.caption, pin.chrome.bar)))
            == Some((false, false)),
        "no strip is asked for or shown over a frame that is not there"
    );

    pin_film_host::disarm_pin_handover();
    clear_pending_pinned_relaunch();
    clear_restart_count();
    clear_park_trace();
    drop(film);
}

/// A step's handover starts the new film with the chrome hidden: the swap held the old frame
/// without chrome until the new player drew, and the hidden chrome stays hidden at that handover
/// until the pointer asks for it.
#[test]
fn a_step_handover_starts_the_new_film_with_the_chrome_hidden() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    film.taken_up_from_a_hover(101);
    let hover = unsafe { StandingHover::up() };
    clear_park_trace();
    clear_restart_count();
    clear_pending_pinned_relaunch();
    pin_film_host::disarm_pin_handover();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    unsafe { show_pinned_window(film.pin_hwnd) };
    // The step's cover standing over the outgoing frame, with the chrome hidden across the swap.
    with_pin(|pin| {
        pin.parked = true;
        pin.chrome = PinChrome::hidden();
    });

    let drawn = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(drawn);
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_child_of_the_host(host) };
    the_hosted_player_has_the_frame(drawn, "0.000000");

    assert!(
        matches!(
            unsafe { pin_film_host::settle_pin_handover() },
            pin_film_host::PinHandover::HandedOver
        ),
        "the tick the new film draws hands the band over"
    );
    assert_eq!(
        pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| (pin.chrome.caption, pin.chrome.bar))),
        Some((false, false)),
        "and the bars stay hidden over the new film after that handover, not before it"
    );
    assert!(
        !hover.shown(),
        "with the hover coming down in the same pass where one is still up"
    );

    pin_film_host::disarm_pin_handover();
    clear_pending_pinned_relaunch();
    clear_restart_count();
    clear_park_trace();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    drop(film);
}

/// A film that never draws hands the pin back: the replacement is dead, so the wait is over and
/// the hidden pin comes down with the hover still up.
#[test]
fn a_film_that_never_draws_hands_the_pin_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    film.taken_up_from_a_hover(101);
    let hover = unsafe { StandingHover::up() };
    clear_park_trace();
    clear_restart_count();
    clear_pending_pinned_relaunch();
    pin_film_host::disarm_pin_handover();
    unsafe { show_pinned_window(film.pin_hwnd) };
    assert!(
        pin_film_host::pin_handover_armed(),
        "the take-up arms the handover"
    );

    // No player behind the wait: the relaunch the take-up began errored rather than starting -
    // headless runs note the player instead of beginning one, so the errored player is stood in
    // for by no player at all.
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        if let Some(media) = current.as_mut() {
            media.video_process = None;
        }
    }

    assert!(
        matches!(
            unsafe { pin_film_host::settle_pin_handover() },
            pin_film_host::PinHandover::Failed
        ),
        "a replacement that is gone is a wait with nothing left to wait for"
    );
    assert!(
        !pin_film_host::pin_handover_armed(),
        "and the handover disarms rather than sticking behind a hidden pin"
    );
    assert_eq!(
        pin_film_host::held_hover_for_a_test(),
        None,
        "and the hover it froze is let go of rather than left frozen for a swap that is never \
         coming"
    );
    assert!(
        !unsafe { IsWindowVisible(film.pin_hwnd) }.as_bool(),
        "the pin never appears"
    );
    assert!(
        hover.shown(),
        "and the hover is still up for the teardown to keep"
    );

    pin_film_host::disarm_pin_handover();
    clear_pending_pinned_relaunch();
    clear_restart_count();
    clear_park_trace();
    drop(film);
}

/// A film that keeps starting hands the pin back at the start timeout: a player that never draws
/// is one that will not, and a hidden pin waiting past the start ever does is a stuck one.
#[test]
fn a_film_that_keeps_starting_hands_the_pin_back_at_the_start_timeout() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    film.taken_up_from_a_hover(101);
    let hover = unsafe { StandingHover::up() };
    clear_park_trace();
    clear_restart_count();
    clear_pending_pinned_relaunch();
    pin_film_host::disarm_pin_handover();
    unsafe { show_pinned_window(film.pin_hwnd) };
    // A player still starting, past the start timeout: alive, but with no frame in the host.
    let _player = LiveStandInPlayer::stand();
    pin_film_host::stand_pin_handover_armed_at(Some(
        Instant::now() - Duration::from_secs(VIDEO_START_WAIT_SECS + 1),
    ));

    assert!(
        matches!(
            unsafe { pin_film_host::settle_pin_handover() },
            pin_film_host::PinHandover::Failed
        ),
        "a start past its timeout is a wait that is over"
    );
    assert!(
        !pin_film_host::pin_handover_armed(),
        "and the handover disarms rather than sticking behind a hidden pin"
    );
    assert_eq!(
        pin_film_host::held_hover_for_a_test(),
        None,
        "and the hover it froze is ended rather than left frozen behind a pin that never came"
    );
    assert!(
        !unsafe { IsWindowVisible(film.pin_hwnd) }.as_bool(),
        "the pin never appears"
    );
    assert!(
        hover.shown(),
        "and the hover is still up for the teardown to keep"
    );

    pin_film_host::disarm_pin_handover();
    clear_pending_pinned_relaunch();
    clear_restart_count();
    clear_park_trace();
    drop(film);
}

/// The move cover is a picture the pin owns: its first paint after a hosted hover is taken up
/// stands on the held frame, so the band is opaque and never the alpha-1 claim that let the
/// desktop show through while the host left the compositor (see `raise_the_move_cover`).
#[test]
fn a_move_paints_the_pins_bitmap_from_the_held_frame_not_a_hole() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let hover = unsafe { HostedHover::up(&film, &path, 3.0) };
    // A place the screen can be read from, so the capture takes the hover's own host rather than
    // failing off the edge of the desktop.
    unsafe {
        let _ = MoveWindow(hover.host, 100, 100, 320, 240, true);
        let _ = ShowWindow(hover.host, SW_SHOWNOACTIVATE);
    }
    pin_film_host::disarm_pin_handover();
    forget_video_frame();

    unsafe { show_pinned_window(film.pin_hwnd) };

    assert!(
        pin_film_host::the_move_cover_is_up(),
        "the adopt road raises the cover before the pin's first paint"
    );
    let (width, height) = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_size()))
        .expect("a pin to have a size");
    let held = held_video_frame();
    let held = held.as_ref().expect("the frame the move holds");
    let bits = ensure_layered_surface(film.pin_hwnd.0 as isize, width as u32, height as u32)
        .expect("the surface the paint drew into");
    let painted =
        unsafe { std::slice::from_raw_parts(bits, width as usize * height as usize * 4) };
    let (band_top, band_height) = pinned_band_rows(
        height,
        pinned_caption_height(96, None),
        0,
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.overlay)) == Some(true),
    );
    let row = (band_top.max(0) as usize + band_height as usize / 2).min(height as usize - 1);
    let column = width as usize / 2;
    let pixel = &painted[(row * width as usize + column) * 4..][..4];
    assert_eq!(pixel[3], 255, "the covered band is opaque, never a hole: {pixel:?}");
    assert_eq!(
        (held.width, held.height),
        (width as u32, height as u32),
        "the captured frame is the band the pin paints"
    );
    let held_pixel = &held.pixels[(row * width as usize + column) * 4..][..4];
    assert_eq!(
        pixel, held_pixel,
        "the pin's bitmap carries the held frame itself, not a flat claim"
    );
    mpv_ipc::forget_pipe(hover.pid);
}

/// The move cover is given up only once a frame has moved in the moved window: a shown child is
/// not proof (the hover's latch is carried over), so the clock has to pass the second captured at
/// take-up. Then the band goes back to its alpha-1 claim, with no stale frame left behind.
#[test]
fn a_move_cover_waits_for_a_frame_past_the_captured_second() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let hover = unsafe { HostedHover::up(&film, &path, 3.0) };
    mpv_ipc::remember_pipe(hover.pid, "a-move-cover-pipe", false);
    mpv_ipc::answer_a_read(hover.pid, "time-pos", Some("3.0"));
    pin_film_host::disarm_pin_handover();

    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the host the move took up");
    // The moved player's child is the one the hover made, and it is still shown under the host.
    assert!(
        pin_film_host::the_hosted_player_is_drawing(VIDEO_PID.load(Ordering::SeqCst)),
        "a visible child is a player drawing, and no more than that"
    );
    assert!(
        !pin_film_host::settle_the_move_cover(),
        "the same second is the same frame: the cover stands"
    );
    assert!(pin_film_host::the_move_cover_is_up(), "holding the band");
    assert!(held_video_frame().is_some(), "with the captured frame still in hand");

    mpv_ipc::answer_a_read(hover.pid, "time-pos", Some("4.0"));
    assert!(
        pin_film_host::settle_the_move_cover(),
        "a frame past the captured second is the moved window presenting"
    );
    assert!(!pin_film_host::the_move_cover_is_up(), "the cover is down");
    assert!(
        held_video_frame().is_none(),
        "and the frame goes with it, so no later park opens on a stale picture"
    );

    let (width, height) = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_size()))
        .expect("a pin to have a size");
    let bits = ensure_layered_surface(film.pin_hwnd.0 as isize, width as u32, height as u32)
        .expect("the surface the paint drew into");
    let painted =
        unsafe { std::slice::from_raw_parts(bits, width as usize * height as usize * 4) };
    let (band_top, band_height) = pinned_band_rows(
        height,
        pinned_caption_height(96, None),
        0,
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.overlay)) == Some(true),
    );
    let row = (band_top.max(0) as usize + band_height as usize / 2).min(height as usize - 1);
    let column = width as usize / 2;
    let pixel = &painted[(row * width as usize + column) * 4..][..4];
    assert_eq!(
        pixel,
        &[1, 1, 1, 1],
        "the handed-back band is the film's alpha-1 claim again: {pixel:?}"
    );
    unsafe { let _ = DestroyWindow(host); }
    mpv_ipc::forget_pipe(hover.pid);
}

/// A move that could take no frame leaves black, never alpha zero: the pin's bitmap is the only
/// guaranteed surface while the host changes parent (see `compose_parked_band`).
#[test]
fn a_move_cover_with_no_frame_is_black_never_a_hole() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    forget_video_frame();
    pin_film_host::raise_the_move_cover();
    unsafe { render_pinned_preview_at(film.pin_hwnd, film.box_.0, film.box_.1) };

    let (width, height) = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_size()))
        .expect("a pin to have a size");
    let bits = ensure_layered_surface(film.pin_hwnd.0 as isize, width as u32, height as u32)
        .expect("the surface the paint drew into");
    let painted =
        unsafe { std::slice::from_raw_parts(bits, width as usize * height as usize * 4) };
    let (band_top, band_height) = pinned_band_rows(
        height,
        pinned_caption_height(96, None),
        0,
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.overlay)) == Some(true),
    );
    let row = (band_top.max(0) as usize + band_height as usize / 2).min(height as usize - 1);
    let column = width as usize / 2;
    let pixel = &painted[(row * width as usize + column) * 4..][..4];
    assert_eq!(pixel[3], 255, "an uncaptured cover is opaque black: {pixel:?}");
    assert_eq!(&pixel[..3], &[0, 0, 0], "black rather than a transparent hole: {pixel:?}");
    pin_film_host::clear_the_move_cover();
}

/// A paused film's cover never settles, so a park, a step or a take-down must clear it: otherwise
/// the next pin would open on the previous film's frame (see `clear_the_move_cover`).
#[test]
fn a_paused_moves_cover_is_cleared_by_a_park_and_a_take_down() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let hover = unsafe { HostedHover::up(&film, &path, 3.0) };
    mpv_ipc::remember_pipe(hover.pid, "a-paused-cover-pipe", false);
    mpv_ipc::answer_a_read(hover.pid, "time-pos", Some("3.0"));
    pin_film_host::disarm_pin_handover();

    unsafe { show_pinned_window(film.pin_hwnd) };
    assert!(pin_film_host::the_move_cover_is_up(), "a paused film leaves it standing");
    assert!(
        !pin_film_host::settle_the_move_cover(),
        "and its clock never moves past the captured second"
    );

    assert!(
        park_pinned_player_for_step(film.pin_hwnd, (film.box_.0, film.box_.1)),
        "a step parks over the outgoing film"
    );
    assert!(
        !pin_film_host::the_move_cover_is_up(),
        "the park's own capture owns the band now"
    );
    give_up_pinned_park();
    forget_video_frame();
    forget_pin_park_swap();

    pin_film_host::raise_the_move_cover();
    unsafe { pin_film_host::take_down_the_pin_film_host() };
    assert!(
        !pin_film_host::the_move_cover_is_up(),
        "and a take-down leaves no cover for the next pin"
    );
    assert!(
        held_video_frame().is_none(),
        "nor the frame the next pin would otherwise open on"
    );
    mpv_ipc::forget_pipe(hover.pid);
}

/// A stand whose film FFmpeg's player draws, rather than mpv's: the route where the player's own
/// window has to be adopted into the pin's host after SDL makes it.
fn ffplay_standing_film() -> StandingFilm {
    let film = StandingFilm::new();
    stand_ffplay_in(Some(true));
    stand_mpv_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    film
}

/// A player's own top-level window of this process, framed and visible, as FFmpeg's player makes
/// it before the adoption restyles it.
unsafe fn a_players_own_window() -> HWND {
    unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP | WS_VISIBLE,
            NOWHERE,
            NOWHERE,
            320,
            180,
            None,
            None,
            None,
            None,
        )
    }
    .expect("a player's own window")
}

/// A pinned FFmpeg film is begun into the pin's host and publishes no top-level window: the
/// monitor adopts the player's own window into the host once it exists (see
/// `start_video_playback_where` and `ffplay_window::adopt_the_players_window`).
#[test]
fn a_pinned_ffplay_film_is_begun_into_the_pins_film_host() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let path = PathBuf::from("a-film.mkv");
    VIDEO_HWND.store(0x1234, Ordering::SeqCst);

    let process =
        start_video_playback(&path, film.box_.0, film.box_.1, 320, 240, 0.0, false, 80, None)
            .expect("a noted FFmpeg player");

    assert_eq!(
        pin_film_host::hosted_player_pid(),
        Some(process.id()),
        "the host records the FFmpeg player it is to hold"
    );
    assert_eq!(
        pin_film_host::film_host(),
        Some(host),
        "and it is the pin's host, not a second one"
    );
    assert_eq!(
        VIDEO_HWND.load(Ordering::SeqCst),
        0,
        "a hosted FFmpeg player publishes no window of the desktop's"
    );
}

/// A pinned FFmpeg film is launched for a layered pin window, where only SDL's Direct3D 11
/// renderer presents, and without `-noborder`, whose non-resizable window would clamp every later
/// band resize. The hover launch keeps both exactly as they were.
#[test]
fn the_hosted_ffplay_launch_asks_for_direct3d11_and_keeps_its_resizable_window() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let path = PathBuf::from("a-film.mkv");

    let _ = start_video_playback(&path, film.box_.0, film.box_.1, 320, 240, 0.0, false, 80, None)
        .expect("a noted FFmpeg player");
    let launched = take_last_spawned_command().expect("the hosted launch");
    assert!(launched.program.contains("ffplay"), "the pin's player");
    assert!(
        !launched.args.iter().any(|arg| arg == "-noborder"),
        "a hosted launch keeps SDL's resizable window, or the band could never resize it"
    );
    assert!(
        launched.env.iter().any(|(key, value)| {
            key.eq_ignore_ascii_case("SDL_RENDER_DRIVER")
                && value.as_deref() == Some("direct3d11")
        }),
        "and asks for the one renderer that presents inside a layered pin window"
    );

    // The same film with no pin and Pin Mode off is an unhosted hover: SDL's default renderer
    // and the borderless window it has always had (see `hover_film_host_window` for the hosted
    // hover's own launch).
    take_pin_for_a_test();
    pin_stands_in_its_own_window(false);
    let pin_enabled_before = CONFIG
        .lock()
        .map(|config| config.pin_enabled)
        .unwrap_or(true);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_enabled = false;
    }
    let _ = start_video_playback(&path, film.box_.0, film.box_.1, 320, 240, 0.0, false, 80, None)
        .expect("a noted FFmpeg hover");
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_enabled = pin_enabled_before;
    }
    let hovered = take_last_spawned_command().expect("the hover launch");
    assert!(
        hovered.args.iter().any(|arg| arg == "-noborder"),
        "the hover launch is unchanged and still asks for no border"
    );
    assert!(
        !hovered
            .env
            .iter()
            .any(|(key, _)| key.eq_ignore_ascii_case("SDL_RENDER_DRIVER")),
        "and it keeps SDL's own renderer choice"
    );
}

/// FFmpeg's own window is restyled to a disabled `WS_CHILD` and re-parented into the host, with
/// the frame stripped and the picture fitted to the host's client box. The pause key and the
/// window-by-pid roads then find that child (see `video_window_for`).
#[test]
fn an_ffplay_players_window_is_adopted_as_a_disabled_child_filling_the_host() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let pid = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(pid);
    let player = unsafe { a_players_own_window() };
    unsafe {
        SetWindowLongPtrW(
            player,
            GWL_EXSTYLE,
            WS_EX_TOPMOST.0 as isize | WS_EX_APPWINDOW.0 as isize,
        )
    };

    assert!(
        unsafe { ffplay_window::adopt_the_players_window(player, host) },
        "the player's window is adopted into the host"
    );

    assert_eq!(
        unsafe { GetParent(player) }.ok(),
        Some(host),
        "under the pin's own host"
    );
    let style = unsafe { GetWindowLongPtrW(player, GWL_STYLE) };
    assert_ne!(style & WS_CHILD.0 as isize, 0, "a child by style");
    assert_eq!(style & WS_POPUP.0 as isize, 0, "and no longer a popup");
    assert!(
        !unsafe { IsWindowEnabled(player) }.as_bool(),
        "disabled, so a press on the picture reaches the pin"
    );
    let ex = unsafe { GetWindowLongPtrW(player, GWL_EXSTYLE) };
    assert_eq!(ex & WS_EX_TOPMOST.0 as isize, 0, "no topmost band of its own");
    assert_eq!(ex & WS_EX_APPWINDOW.0 as isize, 0, "and no taskbar button");
    let child = unsafe { window_rect(player) };
    let host_rect = unsafe { window_rect(host) };
    assert_eq!(
        (child.2 - child.0, child.3 - child.1),
        (host_rect.2 - host_rect.0, host_rect.3 - host_rect.1),
        "fitted to the host's client box rather than left at its born size"
    );
    assert_eq!(video_window_for(pid), Some(player));
    assert!(
        video_key_pause(),
        "the pause key is posted to the child under the host"
    );

    unsafe { DestroyWindow(player) }.expect("the player's window goes");
}

/// The band moves or resizes through the same plain relayout as mpv's: the host is given the new
/// box, FFmpeg's child is fitted to it, and no player is begun again (see `relayout_pinned_media`
/// and `ensure_pinned_sibling_box`).
#[test]
fn the_ffplay_film_follows_the_pins_band_on_move_and_resize_without_a_relaunch() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let pid = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(pid);
    let path = PathBuf::from("a-film.mkv");
    pin_film_host::note_the_pins_film_route(&path);

    let moved = (film.box_.0 + 40, film.box_.1 + 30, film.box_.2 + 80, film.box_.3 + 60);
    ensure_pinned_sibling_box(moved);
    let host_rect = unsafe { window_rect(host) };
    assert_eq!(
        (host_rect.2 - host_rect.0, host_rect.3 - host_rect.1),
        (moved.2 - moved.0, moved.3 - moved.1),
        "a move gives the host the band's box, with the film in it"
    );

    clear_restart_count();
    relayout_pinned_media(&path, moved, 96, None, PinRelayoutRoad::BoxChange);
    assert_eq!(
        restart_count(),
        0,
        "a resize is a SetWindowPos on the host, not a player begun again"
    );
    let host_rect = unsafe { window_rect(host) };
    assert_eq!(
        (host_rect.2 - host_rect.0, host_rect.3 - host_rect.1),
        (moved.2 - moved.0, moved.3 - moved.1),
        "and the host stays at the box it was fitted to"
    );
}

/// The band moves or resizes: the host follows it and FFmpeg's child is fitted to the host's new
/// client box, because SDL does not follow its parent's resize on its own.
#[test]
fn a_band_resize_fits_the_hosted_ffplay_child() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    film.drawn_in_the_host(unsafe { GetCurrentProcessId() });
    let player = unsafe { a_players_own_window() };
    unsafe { ffplay_window::adopt_the_players_window(player, host) };

    let wider = (
        film.box_.0,
        film.box_.1,
        film.box_.0 + 400,
        film.box_.1 + 300,
    );
    with_pin(|pin| pin.content = wider);
    ensure_pinned_sibling_box(wider);

    assert_eq!(
        unsafe { window_rect(host) },
        wider,
        "the host stands in the band the pin now shows its film through"
    );
    let child = unsafe { window_rect(player) };
    let host_rect = unsafe { window_rect(host) };
    assert_eq!(
        (child.2 - child.0, child.3 - child.1),
        (host_rect.2 - host_rect.0, host_rect.3 - host_rect.1),
        "and the FFmpeg child is fitted to the new client box"
    );

    unsafe { DestroyWindow(player) }.expect("the player's window goes");
}

/// A park hides the host and the unpark shows the same one for FFmpeg's film as for mpv's: the
/// picture is a child of the pin, so there is no top-level window to hide instead.
#[test]
fn a_park_hides_and_the_unpark_shows_the_ffplay_film_host() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    film.drawn_in_the_host(unsafe { GetCurrentProcessId() });
    let player = unsafe { a_players_own_window() };
    unsafe { ffplay_window::adopt_the_players_window(player, host) };
    let _ = unsafe { ShowWindow(host, SW_SHOWNOACTIVATE) };

    hide_pinned_player_window();
    assert!(
        !unsafe { IsWindowVisible(host) }.as_bool(),
        "a park hides the film host"
    );
    show_pinned_player_window();
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool(),
        "and the unpark shows the same host again"
    );

    unsafe { DestroyWindow(player) }.expect("the player's window goes");
}

/// A file step begins the next FFmpeg film into the same host: the host outlives the player, and
/// the new pid is what the wait for its child is a wait for.
#[test]
fn a_step_begins_the_next_ffplay_film_into_the_same_host() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let path = PathBuf::from("a-film.mkv");
    let first = start_video_playback(
        &path,
        film.box_.0,
        film.box_.1,
        320,
        240,
        0.0,
        false,
        0,
        None,
    )
    .expect("a noted FFmpeg player");
    let next = start_video_playback(
        &path,
        film.box_.0,
        film.box_.1,
        320,
        240,
        12.0,
        false,
        0,
        None,
    )
    .expect("the next noted FFmpeg player");

    assert_eq!(
        pin_film_host::film_host(),
        Some(host),
        "the step does not make a second host"
    );
    assert_ne!(first.id(), next.id());
    assert_eq!(
        pin_film_host::hosted_player_pid(),
        Some(next.id()),
        "and the host holds the player the step began"
    );
}

/// A pin that ends takes its FFmpeg film host down with it, and the player's child goes too: the
/// player is ended first, so no window outlives the pin it was drawn for.
#[test]
fn a_pin_that_ends_takes_its_ffplay_film_host_with_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    film.drawn_in_the_host(unsafe { GetCurrentProcessId() });
    let player = unsafe { a_players_own_window() };
    unsafe { ffplay_window::adopt_the_players_window(player, host) };

    unsafe { pin_film_host::take_down_the_pin_film_host() };

    assert!(
        !unsafe { IsWindow(host) }.as_bool(),
        "the host goes down with the pin"
    );
    assert!(
        !unsafe { IsWindow(player) }.as_bool(),
        "and the player's child with the host that held it"
    );
    assert!(pin_film_host::film_host().is_none());
    assert!(pin_film_host::hosted_player_pid().is_none());
}

/// The top-level roads leave a hosted FFmpeg player alone: there is no window of the desktop's to
/// style, publish or re-parent, and the child under the host is not restyled again.
#[test]
fn the_top_level_roads_leave_a_hosted_ffplay_player_alone() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let pid = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(pid);
    let player = unsafe { a_players_own_window() };
    unsafe { ffplay_window::adopt_the_players_window(player, host) };
    let style = unsafe { GetWindowLongPtrW(player, GWL_STYLE) };
    let ex = unsafe { GetWindowLongPtrW(player, GWL_EXSTYLE) };

    assert!(
        !unsafe { try_apply_noactivate_style(pid) },
        "the monitor finds nothing of the desktop's to style"
    );
    assert!(
        !unsafe { apply_noactivate_to_hwnd(player) },
        "and a top-level styling of the child is refused"
    );
    assert_eq!(unsafe { GetParent(player) }.ok(), Some(host));
    assert_eq!(unsafe { GetWindowLongPtrW(player, GWL_STYLE) }, style);
    assert_eq!(unsafe { GetWindowLongPtrW(player, GWL_EXSTYLE) }, ex);
    assert_eq!(VIDEO_HWND.load(Ordering::SeqCst), 0);

    unsafe { DestroyWindow(player) }.expect("the player's window goes");
}

/// FFmpeg's player has no pipe to say its video output is configured, so its adopted child is the
/// frame's proof: that window is the player's whole picture and only arrives after the player made
/// and showed it (see `the_hosted_frame_is_up`). This is the fresh-pin, step and seek relaunch
/// road — without it the chrome would stay hidden for the life of the pin.
#[test]
fn a_hosted_ffplay_films_frame_is_its_child_and_the_chrome_comes_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_child_of_the_host(host) };
    with_pin(|pin| {
        pin.chrome = PinChrome::hidden();
        pin.transport_bar = true;
    });
    stand_window_at_point_in(Some(film.pin_hwnd.0 as isize));
    let window = film.box_;
    let caption_cursor = Some((window.0 + 5, window.1));

    assert!(
        pin_film_host::the_hosted_player_is_drawing(pid),
        "the FFmpeg player's own window is under the host"
    );
    assert!(
        pin_film_host::the_hosted_frame_is_up(pid),
        "and that window is the whole of its picture, with no pipe to prove it"
    );

    let shown = pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(refresh_pin_chrome(pin, Instant::now(), caption_cursor))
    });
    assert_eq!(
        shown,
        Some(true),
        "the frame is up, so the chrome is not held down as a window without a picture"
    );
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| (pin.chrome.caption, pin.chrome.bar))),
        Some((true, false)),
        "the caption comes back over the live child"
    );

    stand_window_at_point_in(None);
    drop(film);
}

/// The move cover over FFmpeg's film settles on the child that kept drawing through the reparent,
/// because its player has no clock to read: the adopt road latched the frame and the child is still
/// shown (see `settle_the_move_cover`). Without this the cover would paint a held frame under the
/// live child for the life of the pin.
#[test]
fn a_move_cover_over_a_hosted_ffplay_film_settles_on_its_child() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_child_of_the_host(host) };
    // No pipe of mpv's at all: there is no captured or moved `time-pos` for this player to read.
    mpv_ipc::clear_reads();

    pin_film_host::raise_the_move_cover();
    assert!(pin_film_host::the_move_cover_is_up(), "the move raises it");
    assert!(
        pin_film_host::settle_the_move_cover(),
        "a shown child under the host is the FFmpeg film's proof, with no clock to read"
    );
    assert!(
        !pin_film_host::the_move_cover_is_up(),
        "so the cover is down and no stale frame is left under the live child"
    );

    drop(film);
}

/// A `Hide` can already be in the channel when the pin key takes the window up — a hover's
/// dismissal, or the abandoned load a hover left behind. It is not the pin's end: the window the
/// pin stands in, the host its film is drawn through, and the media and film path the loop holds
/// all stay exactly where the take-up put them.
#[test]
fn a_late_hide_leaves_the_pins_window_host_and_media_alone() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let pid = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(pid);
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_child_of_the_host(host) };
    unsafe { pin_film_host::set_the_pin_film_host_visible(true) };
    let hover = unsafe { a_hover_window_of_the_tests_own() };
    let mut path = Some(PathBuf::from("a-film.mkv"));
    let mut pos = (film.box_.0, film.box_.1, 320, 240);
    let media_before = CURRENT_MEDIA.lock().ok().map(|media| media.is_some());

    assert!(
        unsafe { IsWindowVisible(film.pin_hwnd) }.as_bool(),
        "the pin's own window is on screen"
    );

    unsafe { pin_lifecycle::take_the_hover_down(hover, &mut path, &mut pos) };

    assert!(
        !unsafe { IsWindowVisible(hover) }.as_bool(),
        "the hover this Hide belongs to goes down"
    );
    assert!(
        unsafe { IsWindowVisible(film.pin_hwnd) }.as_bool(),
        "so the pin's own window is still on screen"
    );
    assert_eq!(
        pin_film_host::film_host(),
        Some(host),
        "and its film host is still there"
    );
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool(),
        "still on the band"
    );
    assert_eq!(
        CURRENT_MEDIA.lock().ok().map(|media| media.is_some()),
        media_before,
        "the pin's media was not taken out from under it"
    );
    assert_eq!(
        path,
        Some(PathBuf::from("a-film.mkv")),
        "nor was the film path the tick follows cleared"
    );
    assert_eq!(pos, (film.box_.0, film.box_.1, 320, 240), "nor its box");

    unsafe {
        let _ = DestroyWindow(hover);
    }
    drop(film);
}

/// With the pin drawn in the preview's own window nothing owns the window the hide takes down, so
/// the whole old teardown stands: the hover's window, the media it was playing, the film path the
/// tick follows and the hover's film host all go with it.
#[test]
fn a_hide_with_the_pin_in_the_preview_window_takes_the_hover_down() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    // The take-up did not give this pin a window of its own: it is drawn in the preview's.
    pin_stands_in_its_own_window(false);
    unsafe {
        let _ = ShowWindow(film.pin_hwnd, SW_SHOWNOACTIVATE);
    };
    let hover = unsafe { a_hover_window_of_the_tests_own() };
    let _hover_host = unsafe {
        hover_film_host::take_up_the_hover_film_host((
            NOWHERE,
            NOWHERE,
            NOWHERE + 320,
            NOWHERE + 240,
        ))
    }
    .expect("the hover's own host");
    let mut path = Some(PathBuf::from("a-film.mkv"));
    let mut pos = (film.box_.0, film.box_.1, 320, 240);

    unsafe { pin_lifecycle::take_the_hover_down(hover, &mut path, &mut pos) };

    assert!(
        !unsafe { IsWindowVisible(hover) }.as_bool(),
        "the hover's window goes down"
    );
    assert!(
        !unsafe { IsWindowVisible(film.pin_hwnd) }.as_bool(),
        "and the pin window standing in the preview's goes with it"
    );
    assert!(
        CURRENT_MEDIA
            .lock()
            .ok()
            .is_none_or(|media| media.is_none()),
        "the media is stopped"
    );
    assert_eq!(path, None, "the film path the tick follows is cleared");
    assert_eq!(pos, (0, 0, 0, 0), "and its box");
    assert!(
        hover_film_host::hover_film_host().is_none(),
        "the hover's film host is taken down"
    );

    unsafe {
        let _ = DestroyWindow(hover);
    }
    drop(film);
}

/// The reveal tick must not depend on the hover's own film path: a pinned hosted film whose
/// hover bookkeeping is empty is still owed the front, and the host is shown once the player's
/// child is under it.
#[test]
fn a_pinned_hosted_film_is_revealed_without_a_hover_path() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let pid = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(pid);
    let host = pin_film_host::film_host().expect("the film's host");
    unsafe { pin_film_host::set_the_pin_film_host_visible(false) };
    let _child = unsafe { a_child_of_the_host(host) };

    assert!(
        !unsafe { IsWindowVisible(host) }.as_bool(),
        "the host is hidden before the tick runs"
    );
    if pin_film_host::the_video_window_is_owed_the_front(None) {
        video_playback::bring_the_video_window_to_the_front(film.box_);
    }
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool(),
        "the tick shows the host the pin's film is drawing in"
    );

    drop(film);
}

/// The same for a film FFmpeg's player draws: the pin stands in its own window, the adopted child
/// is under the host, and both the window and the host are on screen with no hover path at all.
#[test]
fn a_pinned_ffplay_film_is_revealed_without_a_hover_path() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let pid = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(pid);
    let host = pin_film_host::film_host().expect("the film's host");
    unsafe { pin_film_host::set_the_pin_film_host_visible(false) };
    let _child = unsafe { a_child_of_the_host(host) };
    let path: Option<PathBuf> = None;

    if pin_film_host::the_video_window_is_owed_the_front(path.as_deref()) {
        video_playback::bring_the_video_window_to_the_front(film.box_);
    }

    assert!(
        unsafe { IsWindowVisible(film.pin_hwnd) }.as_bool(),
        "the pin's own window is up"
    );
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool(),
        "and the host FFmpeg's child is drawing in is up with it"
    );
    assert!(
        pin_film_host::the_hosted_player_is_drawing(pid),
        "with the film actually drawing in that host"
    );

    drop(film);
}

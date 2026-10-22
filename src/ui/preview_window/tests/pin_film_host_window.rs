//! The window a pinned mpv film is drawn in: a disabled child of the pin's own window, standing in
//! the band the picture is seen through, hidden for a park and shown again after it, and going down
//! with the pin it was made for.
//!
//! Every test here stands a pin in a window of its own (see `pin_real_window`) and a film mpv is
//! playing, because the host only exists for that pair: a film FFmpeg's player draws is a window of
//! its own until its ticket, and every other kind is this app's own frame.

use super::*;
use windows::Win32::UI::Input::KeyboardAndMouse::IsWindowEnabled;
use windows::Win32::UI::WindowsAndMessaging::{
    DestroyWindow, GetParent, GWL_STYLE, WINDOW_EX_STYLE, WS_CHILD, WS_VISIBLE,
};

/// The place a window of a test's own is put when it must not be on the user's screen.
const NOWHERE: i32 = -32000;

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
/// same process by the same rule (see `the_pinned_film_is_drawing`).
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
}

impl StandingFilm {
    fn new() -> StandingFilm {
        let previous_pin = take_pin_for_a_test();
        let previous_own = pin_is_its_own_window();
        let previous_window = pin_real_window::real_pin_window();
        let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
        let previous_pid = VIDEO_PID.swap(unsafe { GetCurrentProcessId() }, Ordering::SeqCst);

        stand_mpv_in(Some(true));
        stand_video_engine_in(Some((VideoEngine::Mpv, true)));

        let box_ = (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 240);
        let mut pin = overlay_pin(box_, PinChrome::always());
        pin.path = PathBuf::from("a-film.mkv");
        stand_pin(Some(pin));
        pin_stands_in_its_own_window(true);

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
        }
    }
}

impl Drop for StandingFilm {
    fn drop(&mut self) {
        unsafe {
            pin_film_host::take_down_the_pin_film_host();
        }
        pin_real_window::stand_real_pin_window_for_a_test(self.previous_window);
        unsafe {
            let _ = DestroyWindow(self.pin_hwnd);
        }
        pin_stands_in_its_own_window(self.previous_own);
        stand_pin(self.previous_pin.take());
        VIDEO_PID.store(self.previous_pid, Ordering::SeqCst);
        stand_mpv_in(None);
        stand_video_engine_in(None);
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = self.previous_media.take();
        }
    }
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
        pin_film_host::pin_film_host().is_none(),
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

    let host = pin_film_host::pin_film_host().expect("the window the film is drawn in");
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
        pin_film_host::pin_film_host().is_none(),
        "with the handle forgotten as well as the window destroyed"
    );
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
    let host = pin_film_host::pin_film_host().expect("the film's host");
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
    let host = pin_film_host::pin_film_host().expect("the film's host");

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
    let host = pin_film_host::pin_film_host().expect("the film's host");
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    pin_film_host::note_hosted_player(pid);

    assert!(
        !pin_film_host::the_pinned_film_is_drawing(),
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
        pin_film_host::the_pinned_film_is_drawing(),
        "a child of the player's own process under the host is the film on screen"
    );

    unsafe {
        let _ = DestroyWindow(child);
    };

    assert!(
        !pin_film_host::the_pinned_film_is_drawing(),
        "and the child going is the film not being drawn anywhere"
    );
}

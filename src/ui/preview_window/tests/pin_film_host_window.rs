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
    DestroyWindow, GetParent, MoveWindow, GWL_EXSTYLE, GWL_STYLE, HWND_TOP, SWP_NOACTIVATE,
    SWP_NOZORDER, SW_SHOWNOACTIVATE, WINDOW_EX_STYLE, WS_CHILD, WS_POPUP, WS_VISIBLE,
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
    previous_hwnd: isize,
}

impl StandingFilm {
    /// The same stand with the film drawing in the host: `drawn` is the player whose child under it
    /// is the picture (see `the_pinned_film_is_drawing`), and it owns no window of the desktop's.
    fn drawn_in_the_host(&self, drawn: u32) {
        pin_film_host::note_hosted_player(drawn);
        VIDEO_PID.store(drawn, Ordering::SeqCst);
    }

    /// The same stand taken up on a hover's player: the film is mpv's, and the window it is drawing
    /// in is a window of the desktop's rather than the host's (see `host_holds_the_pinned_film`),
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
        // The host holds this test's own process, which is the pid of any child it makes: a film is
        // drawn in the host only when the player begun into it is the one this app recorded (see
        // `host_holds_the_pinned_film`).
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
        VIDEO_HWND.store(self.previous_hwnd, Ordering::SeqCst);
        stand_mpv_in(None);
        stand_video_engine_in(None);
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = self.previous_media.take();
        }
    }
}

/// A window of the player's own under the host, which is the film drawing where the pin can see
/// it (see `the_pinned_film_is_drawing`): made by this process, which is the player the stand's
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
    let host = pin_film_host::pin_film_host().expect("the film's host");

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
/// `host_holds_the_pinned_film`).
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
    let host = pin_film_host::pin_film_host().expect("the film's host");

    let _ = end_pin(Reason::Closed, &Win32PinWindow);

    assert!(
        !unsafe { IsWindow(host) }.as_bool(),
        "the window the film was drawn in goes down with the pin, whichever road ended it"
    );
    assert!(
        pin_film_host::pin_film_host().is_none(),
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
    let host = pin_film_host::pin_film_host().expect("the film's host");

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

/// A hover's mpv film is a window of the desktop's, and a pin that takes it up cannot keep it
/// there: the picture belongs in the host the pin draws it in. So the take-up begins the player
/// again through the relaunch road the rest of the pin's life uses — at the second the hover had
/// reached, and held or playing as it was — and the band is covered while it does (see
/// `begin_the_adopted_film_in_the_host`).
#[test]
fn a_film_taken_up_from_a_hover_is_begun_again_in_the_host() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let hovered = 0x7FFF_FFFB;
    film.taken_up_from_a_hover(hovered);
    clear_restart_count();

    unsafe { show_pinned_window(film.pin_hwnd) };

    let host = pin_film_host::pin_film_host().expect("the film's host");
    assert_eq!(
        restart_count(),
        1,
        "the player the hover began is begun again rather than kept as it stands"
    );
    assert!(
        pin_player_is_parked(),
        "behind the cover the relaunch road raises over the player it replaces"
    );
    assert_ne!(
        VIDEO_PID.load(Ordering::SeqCst),
        hovered,
        "and the window it was drawing in is not the pin's player any more"
    );
    assert_eq!(
        mpv_launch_host(Path::new("a-film.mkv")),
        Some(host.0 as usize as u32),
        "because the player begun in its place draws in the pin's own host"
    );

    // The other half of the same question: a film the host is already drawing is a film to leave
    // alone, or every show of a hosted pin would begin its player again.
    drop(film);
    let hosted = StandingFilm::new();
    clear_restart_count();
    unsafe { show_pinned_window(hosted.pin_hwnd) };
    assert_eq!(
        restart_count(),
        0,
        "and a pin whose film the host already holds begins nothing"
    );
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
    let host = pin_film_host::pin_film_host().expect("the film's host");

    assert!(
        !unsafe { IsWindowVisible(host) }.as_bool(),
        "the host is made where the band is and left off the screen until there is a picture in it"
    );

    let child = unsafe { a_child_of_the_host(host) };
    assert!(
        pin_film_host::the_pinned_film_is_drawing(),
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
    let host = pin_film_host::pin_film_host().expect("the film's host");

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
    let host = pin_film_host::pin_film_host().expect("the film's host");

    // The player's own child, which is what a hosted film drawing is (see
    // `the_pinned_film_is_drawing`).
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
    let host = pin_film_host::pin_film_host().expect("the film's host");
    // The player's own child is the film drawing in the band (see `the_pinned_film_is_drawing`),
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
    let previous_retirement = VIDEO_RETIREMENT.lock().ok().and_then(|mut held| held.take());

    let film = StandingFilm::new();
    let drawn = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(drawn);
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::pin_film_host().expect("the film's host");
    let _child = unsafe { a_child_of_the_host(host) };

    // The replacement is this test's own process, recorded under the image it runs: alive as far
    // as the check is concerned, which is what makes its *arrival* the only thing that can end the
    // wait (see `is_player_pid_alive` and `note_player_process`).
    let image = std::env::current_exe()
        .ok()
        .and_then(|exe| exe.file_name().map(|name| name.to_string_lossy().into_owned()))
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
    let host = pin_film_host::pin_film_host().expect("the film's host");
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
    let host = pin_film_host::pin_film_host().expect("the film's host");

    let path = Path::new("a-film.mkv");
    let wid = mpv_launch_host(path).expect("the host the film is drawn in");

    assert_eq!(
        wid,
        host.0 as usize as u32,
        "the window the player is told to draw into is the pin's own child"
    );
    assert!(
        video_launch::mpv_args(
            path, 0, 0, 320, 240, 0.0, 80, None, false, false, None, "rhp-test", Some(wid),
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

//! The top-level host a hovered mpv film is drawn in, before a pin ever takes it: made for the
//! hover only where Pin Mode is on and the file is mpv's, placed at the hover's own box, revealed
//! only when the player has a child drawing, and destroyed on every road the hover ends.
//!
//! No real player is begun here — the suite's `spawn_player` notes one instead — and every window
//! is stood off the user's desktop.

use super::*;
use windows::Win32::UI::Input::KeyboardAndMouse::IsWindowEnabled;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DestroyWindow, IsWindowVisible, GWL_EXSTYLE, GWL_STYLE, WINDOW_EX_STYLE,
    WS_CHILD, WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_EX_TRANSPARENT, WS_POPUP,
    WS_VISIBLE,
};

const NOWHERE: i32 = -32000;

/// A hover's machine stood where a test needs it, with every global put back on the way out: the
/// Pin Mode switch the route is read from, mpv as the film's player, and the two handles a player
/// is found by.
struct HoverHostStand {
    pin_enabled: bool,
    previous_pid: u32,
    previous_hwnd: isize,
    previous_media: Option<MediaData>,
}

impl HoverHostStand {
    fn new(pin_enabled: bool) -> HoverHostStand {
        let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
        let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);
        let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
        let pin_enabled_before = CONFIG
            .lock()
            .map(|config| config.pin_enabled)
            .unwrap_or(true);
        if let Ok(mut config) = CONFIG.lock() {
            config.pin_enabled = pin_enabled;
        }
        // A host left by a test that failed before its guard ran is a handle to nothing by now.
        unsafe { hover_film_host::take_down_the_hover_film_host() };
        stand_mpv_in(Some(true));
        stand_ffplay_in(Some(false));
        stand_video_engine_in(Some((VideoEngine::Mpv, true)));
        HoverHostStand {
            pin_enabled: pin_enabled_before,
            previous_pid,
            previous_hwnd,
            previous_media,
        }
    }
}

impl Drop for HoverHostStand {
    fn drop(&mut self) {
        unsafe { hover_film_host::take_down_the_hover_film_host() };
        if let Ok(mut config) = CONFIG.lock() {
            config.pin_enabled = self.pin_enabled;
        }
        stand_mpv_in(None);
        stand_ffplay_in(None);
        stand_video_engine_in(None);
        VIDEO_PID.store(self.previous_pid, Ordering::SeqCst);
        VIDEO_HWND.store(self.previous_hwnd, Ordering::SeqCst);
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = self.previous_media.take();
        }
    }
}

fn a_film() -> PathBuf {
    PathBuf::from("a-hovered-film.mkv")
}

fn video_media() -> MediaData {
    let mut media = create_loading_media(16, 16);
    media.media_type = MediaType::Video;
    media
}

fn window_rect(hwnd: HWND) -> (i32, i32, i32, i32) {
    let mut rect = RECT::default();
    unsafe { GetWindowRect(hwnd, &mut rect) }.expect("a window's rectangle");
    (rect.left, rect.top, rect.right, rect.bottom)
}

/// A child of the host of the player's own process, shown and with a box: the film drawing (see
/// `the_hover_player_is_drawing`).
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
    .expect("a child of the hover's host, in this process")
}

/// An mpv hover is begun with `--wid` into a top-level host of the app's own only with Pin Mode
/// on: with it off the hover keeps today's `--ontop --geometry` launch, and no host is made.
#[test]
fn an_mpv_hover_is_begun_in_a_host_with_pin_mode_on_and_top_level_with_it_off() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());

    {
        let _stand = HoverHostStand::new(true);
        let path = a_film();
        assert!(
            hover_film_host::the_hover_film_is_kept_in_a_host(&path),
            "a film mpv plays, with Pin Mode on, is the route the host is for"
        );

        let process = start_video_playback(&path, 100, 120, 320, 180, 0.0, false, 0, None)
            .expect("a noted player");
        let pid = process.id();
        assert!(
            hover_film_host::the_hover_player_was_begun_in_the_host(pid),
            "the player is begun into the hover's host"
        );
        let host = hover_film_host::hover_film_host().expect("the hover's own host");
        let style = unsafe { GetWindowLongPtrW(host, GWL_STYLE) };
        let ex = unsafe { GetWindowLongPtrW(host, GWL_EXSTYLE) };
        assert_ne!(
            style & WS_POPUP.0 as isize,
            0,
            "the host is a top-level popup"
        );
        assert_eq!(
            style & WS_CHILD.0 as isize,
            0,
            "and not a child of anything"
        );
        for bit in [
            WS_EX_TOOLWINDOW,
            WS_EX_NOACTIVATE,
            WS_EX_TOPMOST,
            WS_EX_TRANSPARENT,
        ] {
            assert_ne!(
                ex & bit.0 as isize,
                0,
                "the host wears the hover's own bits: {bit:?}"
            );
        }
        assert!(
            !unsafe { IsWindowEnabled(host) }.as_bool(),
            "a disabled host takes no press"
        );
        assert!(
            !unsafe { IsWindowVisible(host) }.as_bool(),
            "and is off the screen until the player draws into it"
        );
        assert_eq!(
            window_rect(host),
            (100, 120, 420, 300),
            "the host is placed at the hover's own box rather than mpv's geometry"
        );
        assert!(
            !hover_film_host::the_hover_player_is_drawing(pid),
            "a player just begun has no child under the host yet"
        );
        assert_eq!(
            video_window_for(pid),
            None,
            "and so has no window of the desktop's for the top-level roads"
        );

        let mut media = video_media();
        media.video_process = Some(process);
        stop_video_playback(&mut media);
        assert!(
            hover_film_host::hover_film_host().is_none(),
            "the hover ending takes its host with it"
        );
    }

    {
        let _stand = HoverHostStand::new(false);
        let path = a_film();
        assert!(
            !hover_film_host::the_hover_film_is_kept_in_a_host(&path),
            "with Pin Mode off the hover is begun the way it always was"
        );
        let process = start_video_playback(&path, 100, 120, 320, 180, 0.0, false, 0, None)
            .expect("a noted player");
        assert!(
            hover_film_host::hover_film_host().is_none(),
            "and no host stands for it"
        );
        assert!(
            !hover_film_host::the_hover_player_was_begun_in_the_host(process.id()),
            "the player is a window of the desktop's"
        );
    }
}

/// The host follows the hover's box, stays off the screen until a shown child is in it, and is
/// what every "is the player's window up" road answers with while it is.
#[test]
fn the_hover_host_follows_the_box_and_answers_for_the_players_window() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let _stand = HoverHostStand::new(true);
    let pid = unsafe { GetCurrentProcessId() };

    let host = unsafe {
        hover_film_host::take_up_the_hover_film_host((
            NOWHERE,
            NOWHERE,
            NOWHERE + 320,
            NOWHERE + 180,
        ))
    }
    .expect("the hover's host");
    assert_eq!(
        window_rect(host),
        (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 180),
        "the host is made at the hover's box"
    );

    unsafe {
        hover_film_host::place_the_hover_film_host((
            NOWHERE,
            NOWHERE + 40,
            NOWHERE + 640,
            NOWHERE + 400,
        ))
    };
    assert_eq!(
        window_rect(host),
        (NOWHERE, NOWHERE + 40, NOWHERE + 640, NOWHERE + 400),
        "and follows it on a move or a relayout"
    );

    hover_film_host::note_hover_film_host_player(pid, &a_film());
    assert!(
        !hover_film_host::the_hover_player_is_drawing(pid),
        "recording the player is not a picture"
    );
    assert!(
        !pin_film_host::the_players_window_is_up(pid),
        "so the start wait goes on"
    );

    let child = unsafe { a_child_of_the_host(host) };
    assert!(
        hover_film_host::the_hover_player_is_drawing(pid),
        "a shown child with a box is the player drawing"
    );
    assert_eq!(
        video_window_for(pid),
        Some(child),
        "and the child is the player's window for the roads that ask by pid"
    );
    assert!(pin_film_host::the_players_window_is_up(pid));

    unsafe {
        hover_film_host::reveal_the_hover_film_host(
            pid,
            (NOWHERE, NOWHERE, NOWHERE + 320, NOWHERE + 180),
        )
    };
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool(),
        "the first frame reveals the host"
    );

    unsafe { DestroyWindow(child) }.expect("the child goes");
    assert!(
        !pin_film_host::the_players_window_is_up(pid),
        "and a player with no child is a window that is not up"
    );
}

/// The loop keeps a hover's video place as `(x, y, width, height)` and the host it reveals is placed
/// at a region: revealed from the loop's own box, the host ends at `(x, y, x + w, y + h)` rather than
/// at the box's own numbers, which is the wrong size the reveal used to put up for a frame.
#[test]
fn the_loops_hover_box_reveals_the_host_at_the_hovers_own_size() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let _stand = HoverHostStand::new(true);
    let pid = unsafe { GetCurrentProcessId() };
    let path = a_film();

    let box_ = (NOWHERE, NOWHERE + 40, 320, 180);
    let host = unsafe { hover_film_host::take_up_the_hover_film_host((0, 0, 1, 1)) }
        .expect("the hover's host");
    hover_film_host::note_hover_film_host_player(pid, &path);
    let _child = unsafe { a_child_of_the_host(host) };

    unsafe {
        hover_film_host::reveal_the_hover_film_host(
            pid,
            hover_film_host::the_hover_box_as_a_region(box_),
        )
    };

    assert_eq!(
        window_rect(host),
        (box_.0, box_.1, box_.0 + box_.2, box_.1 + box_.3),
        "the first reveal is the hover's own size and aspect"
    );
    assert_ne!(
        window_rect(host),
        (box_.0, box_.1, box_.2, box_.3),
        "and not the box's numbers read as a region, which is the wrong size"
    );
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool(),
        "the host is up once the player is drawing"
    );
}

/// A hosted hover's visible surface is the hover's own host, and a pin key over it takes that box up:
/// the preview's window is hidden behind the host and a hosted player has no window of the desktop's,
/// so neither of the two the rect used to know about is on screen.
#[test]
fn a_visible_hover_host_is_the_rect_a_pin_takes_up() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let _stand = HoverHostStand::new(true);
    let pid = unsafe { GetCurrentProcessId() };
    let path = a_film();

    let box_ = (NOWHERE, NOWHERE + 80, 320, 180);
    let host = unsafe {
        hover_film_host::take_up_the_hover_film_host(hover_film_host::the_hover_box_as_a_region(box_))
    }
    .expect("the hover's host");
    hover_film_host::note_hover_film_host_player(pid, &path);
    let _child = unsafe { a_child_of_the_host(host) };
    VIDEO_PID.store(pid, Ordering::SeqCst);
    VIDEO_HWND.store(0, Ordering::SeqCst);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(video_media());
    }

    unsafe {
        hover_film_host::reveal_the_hover_film_host(
            pid,
            hover_film_host::the_hover_box_as_a_region(box_),
        )
    };
    let host_rect = window_rect(host);

    assert_eq!(
        preview_screen_rect(),
        Some(host_rect),
        "the hover's host is the preview surface on screen"
    );

    let show = PreviewMessage::Show(path.clone(), 0, 0, None);
    match pin_what_is_on_screen(&Some(show), None) {
        Some(PreviewMessage::Pin {
            path: pinned,
            rect,
        }) => {
            assert_eq!(pinned, path, "the pin takes up the file that is on screen");
            assert_eq!(rect, host_rect, "and the box the hover's host is in");
        }
        _ => panic!("the pin key takes the hover's box up rather than finding nothing"),
    }
}

/// Teardown is owner-aware: ending another player leaves the hover's host standing, and only the
/// player it holds takes it down. A new player replaces the old one in the same host rather than
/// making a second.
#[test]
fn the_hover_host_goes_down_only_with_the_player_it_holds() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let _stand = HoverHostStand::new(true);
    let path = a_film();

    let first =
        start_video_playback(&path, 0, 0, 320, 180, 0.0, false, 0, None).expect("a noted player");
    let first_pid = first.id();
    let host = hover_film_host::hover_film_host().expect("the hover's host");

    // Another player's end is no reason to take this host down.
    let mut other = video_media();
    other.video_process = Some(PlayerProcess::Noted(first_pid.wrapping_sub(4)));
    kill_player_process(&mut other);
    assert_eq!(
        hover_film_host::hover_film_host(),
        Some(host),
        "a player the host does not hold leaves it standing"
    );

    // The next hover player is begun into the same host, which now belongs to it.
    let second =
        start_video_playback(&path, 0, 0, 320, 180, 0.0, false, 0, None).expect("a noted player");
    assert_eq!(
        hover_film_host::hover_film_host(),
        Some(host),
        "a hover that replaces another reuses its host rather than making a second"
    );
    assert!(hover_film_host::the_hover_player_was_begun_in_the_host(
        second.id()
    ));

    // The older player is no longer the host's, so ending it changes nothing.
    let mut old = video_media();
    old.video_process = Some(first);
    kill_player_process(&mut old);
    assert_eq!(hover_film_host::hover_film_host(), Some(host));

    let mut current = video_media();
    current.video_process = Some(second);
    kill_player_process(&mut current);
    assert!(
        hover_film_host::hover_film_host().is_none(),
        "the player the host holds takes it down with it"
    );
}

/// FFmpeg's player, the media engine and a picture keep the route they have: a host is mpv's and a
/// hovered film's, and nothing else enters one.
#[test]
fn ffplay_native_and_picture_routes_never_enter_a_host() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let _stand = HoverHostStand::new(true);
    let path = a_film();

    stand_ffplay_in(Some(true));
    stand_mpv_in(Some(false));
    assert!(!hover_film_host::the_hover_film_is_kept_in_a_host(&path));
    let process =
        start_video_playback(&path, 0, 0, 320, 180, 0.0, false, 0, None).expect("a noted ffplay");
    assert!(
        hover_film_host::hover_film_host().is_none(),
        "an FFmpeg hover is never given a host"
    );
    assert!(!hover_film_host::the_hover_player_was_begun_in_the_host(
        process.id()
    ));

    // The media engine's route is not this launch at all, and a picture is not a film.
    stand_video_engine_in(Some((VideoEngine::Native, true)));
    assert!(!hover_film_host::the_hover_film_is_kept_in_a_host(&path));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    assert!(
        !hover_film_host::the_hover_film_is_kept_in_a_host(Path::new("a-picture.png")),
        "a picture has no film to draw into a host"
    );
}

/// A hosted hover has exactly one ordinary relaunch behind it: alive, not yet fallen back, and
/// begun in the host. The relaunch itself makes no host, so a second failure has nothing left to
/// try (see `restart_hover_video_without_a_host`).
#[test]
fn a_hosted_hover_that_never_draws_falls_back_once_without_a_host() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let _stand = HoverHostStand::new(true);

    assert!(the_hover_host_fallback_is_owed(true, false, true));
    assert!(
        !the_hover_host_fallback_is_owed(true, true, true),
        "the fallback is spent once"
    );
    assert!(
        !the_hover_host_fallback_is_owed(false, false, true),
        "a top-level player has no host to fall back from"
    );
    assert!(
        !the_hover_host_fallback_is_owed(true, false, false),
        "and a dead player is abandoned rather than relaunched"
    );

    let path = a_film();
    unsafe { hover_film_host::take_up_the_hover_film_host((0, 0, 320, 180)) }
        .expect("the host the first player was begun in");
    let process = restart_hover_video_without_a_host(&path, 0, 0, 320, 180, 0.0, false, 0, None)
        .expect("a noted player");
    assert!(
        !hover_film_host::the_hover_player_was_begun_in_the_host(process.id()),
        "the fallback launch is an ordinary top-level one"
    );
}

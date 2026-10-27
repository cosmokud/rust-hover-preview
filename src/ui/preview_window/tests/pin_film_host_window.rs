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
    CreateWindowExW, DefWindowProcW, DestroyWindow, GetClientRect, GetParent, GetWindowLongPtrW,
    IsIconic, IsWindowVisible, MoveWindow, SendMessageW, ShowWindow, GWL_EXSTYLE, GWL_STYLE,
    HWND_TOP, SWP_NOACTIVATE, SWP_NOZORDER, SW_SHOWNOACTIVATE, WINDOW_EX_STYLE, WM_ERASEBKGND,
    WS_CHILD, WS_CLIPCHILDREN, WS_DISABLED, WS_EX_APPWINDOW, WS_EX_LAYERED, WS_POPUP, WS_VISIBLE,
};

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
        // And no hover's either: a host the stand does not own would be the one a take-up promotes.
        unsafe { hover_film_host::take_down_the_hover_film_host() };
        pin_film_thumbnail::forget_the_thumbnail();
        // And no strips either: they belong to the pin's window rather than to the host, so a
        // take-down of the host leaves whatever handles a test before this one stood behind.
        unsafe { pin_chrome_strips::take_the_strips_down() };
        // The host holds this test's own process, which is the pid of any child it makes: a film is
        // drawn in the host only when the player begun into it is the one this app recorded (see
        // `the_pinned_player_was_begun_in_the_host`).
        pin_film_host::note_hosted_player(unsafe { GetCurrentProcessId() });
        // **The stand's film is mpv's, so its player is one that answers a pipe** — which is how a
        // frame is proved to be drawn rather than merely made (see
        // `pin_film_host::the_hosted_frame_is_up`). A test that wants a frame answers it (see
        // `the_hosted_player_has_the_frame`).
        mpv_ipc::remember_pipe(
            unsafe { GetCurrentProcessId() },
            "a-standing-films-pipe",
            false,
        );

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
        mpv_ipc::forget_pipe(self.previous_pid);
        mpv_ipc::forget_pipe(unsafe { GetCurrentProcessId() });
        VIDEO_PID.store(self.previous_pid, Ordering::SeqCst);
        VIDEO_HWND.store(self.previous_hwnd, Ordering::SeqCst);
        stand_mpv_in(None);
        stand_video_engine_in(None);
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = self.previous_media.take();
        }
    }
}

/// The film drawing in the host, on the road both engines are on: a window of the player's own,
/// recorded as the source the Shell draws into `_host` (see `pin_film_thumbnail`). It is this
/// process's window, which is the player every stand in this file records, and the caller destroys
/// it.
unsafe fn a_source_of_the_host(_host: HWND) -> HWND {
    let source = unsafe { a_players_own_window() };
    pin_film_thumbnail::note_the_source(source, VIDEO_PID.load(Ordering::SeqCst));
    source
}

/// The hosted player's own pipe, and the player's word that the frame it is standing on is on
/// screen: the second the hover froze at, and a video output configured for that frame (see
/// `the_hosted_frame_is_presented`).
fn the_hosted_player_has_the_frame(pid: u32, at: &str) {
    mpv_ipc::remember_pipe(pid, "the-hosted-players-pipe", false);
    mpv_ipc::answer_a_read(pid, "time-pos", Some(at));
    mpv_ipc::answer_a_read(pid, "vo-configured", Some("true"));
}

/// Every erase a stand-in host was handed, by window: whether a point inside the child the test gave
/// it, a point outside that child, and a point beyond the box it was made at were in the device
/// context's system region (see `an_erase_is_clipped_out_of_a_childs_box_where_the_window_asks_it_to_be`).
/// One erase as the watcher recorded it: the window, whether a point inside and one outside were
/// in the clip region, and whether a point past the box was.
type EraseClip = (isize, bool, bool, bool);

static ERASE_CLIPS: Lazy<Mutex<Vec<EraseClip>>> = Lazy::new(|| Mutex::new(Vec::new()));

/// A window of a test's own that records what an erase of it is clipped to, and answers the erase the
/// way the hosts do.
unsafe extern "system" fn record_the_erase_clip(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    use windows::Win32::Graphics::Gdi::{
        CreateRectRgn, DeleteObject, GetRandomRgn, PtInRegion, HGDIOBJ, SYSRGN,
    };

    if msg == WM_ERASEBKGND {
        let dc = HDC(wparam.0 as *mut _);
        // SAFETY: the region is this call's own, and the context is the one Windows handed the erase
        // with; both are given up before this answers.
        unsafe {
            let region = CreateRectRgn(0, 0, 0, 0);
            // The system region a window is erased within is in screen coordinates, so the two
            // points the window is asked about are taken from its client box onto the desktop.
            let mut childs_corner = POINT { x: 8, y: 8 };
            let mut outside = POINT { x: 240, y: 200 };
            // Beyond the box the window is made at, so that a strip a grow has exposed is a point
            // the erase of it can be asked about.
            let mut gained = POINT { x: 400, y: 300 };
            let _ = ClientToScreen(hwnd, &mut childs_corner);
            let _ = ClientToScreen(hwnd, &mut outside);
            let _ = ClientToScreen(hwnd, &mut gained);
            // `1` is a system region copied out; `-1` is an empty one, which is no region at all.
            let copied = GetRandomRgn(dc, region, SYSRGN as i32) > 0;
            let inside = copied && PtInRegion(region, childs_corner.x, childs_corner.y).as_bool();
            let outside = copied && PtInRegion(region, outside.x, outside.y).as_bool();
            let reached = copied && PtInRegion(region, gained.x, gained.y).as_bool();
            let _ = DeleteObject(HGDIOBJ(region.0));
            if let Ok(mut clips) = ERASE_CLIPS.lock() {
                clips.push((hwnd.0 as isize, inside, outside, reached));
            }
        }
        return LRESULT(1);
    }

    unsafe { DefWindowProcW(hwnd, msg, wparam, lparam) }
}

/// **`WS_CLIPCHILDREN` is the whole of why the fill cannot touch the live film, and this is that
/// premise measured rather than assumed.** A window painted the way the hosts are — nothing but an
/// erase of its own, wearing `NULL_BRUSH` — is handed a device context by Windows whose system
/// region has no pixel of its child in it, which is what lets `paint_the_host_behind_the_film`
/// stretch a picture over the whole client box. The second window here wears no such bit and answers
/// the other way, so it is the bit and not the box that does it.
///
/// It is a window of the test's own rather than a host, because what is being joined is Windows' own
/// behaviour: the message arriving to a class with no brush at all, and the context it brings.
#[test]
fn an_erase_is_clipped_out_of_a_childs_box_where_the_window_asks_it_to_be() {
    use windows::Win32::Graphics::Gdi::{
        GetStockObject, InvalidateRect, UpdateWindow, HBRUSH, NULL_BRUSH,
    };
    use windows::Win32::UI::WindowsAndMessaging::{
        IsWindowVisible, RegisterClassExW, ShowWindow, WS_VISIBLE,
    };

    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    const ERASE_PROBE_CLASS: windows::core::PCWSTR = w!("RustHoverPreviewTestEraseProbe");
    let hinstance = unsafe { GetModuleHandleW(None) }.expect("the module's own handle");
    let registered = unsafe {
        RegisterClassExW(&WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(record_the_erase_clip),
            hInstance: hinstance.into(),
            hbrBackground: HBRUSH(GetStockObject(NULL_BRUSH).0),
            lpszClassName: ERASE_PROBE_CLASS,
            ..Default::default()
        })
    };
    assert_ne!(registered, 0, "a class of the test's own to erase");
    ERASE_CLIPS.lock().expect("the clips recorded").clear();

    let mut stands = Vec::new();
    for (index, clipped) in [true, false].into_iter().enumerate() {
        let style = if clipped {
            WS_POPUP | WS_CLIPCHILDREN
        } else {
            WS_POPUP
        };
        // SAFETY: both windows are this test's own, and both are destroyed below.
        unsafe {
            let parent = CreateWindowExW(
                WINDOW_EX_STYLE::default(),
                ERASE_PROBE_CLASS,
                w!("erase probe"),
                style,
                // Side by side, so that neither stands over the other's band.
                index as i32 * 320,
                0,
                320,
                240,
                None,
                None,
                hinstance,
                None,
            )
            .expect("a window to erase");
            // The child before the show, so that every erase this window is handed is one with a
            // child in it.
            let child = CreateWindowExW(
                WINDOW_EX_STYLE::default(),
                w!("STATIC"),
                w!(""),
                WS_CHILD | WS_VISIBLE,
                0,
                0,
                160,
                120,
                parent,
                None,
                None,
                None,
            )
            .expect("a child of the window's, in this process");
            // Shown, because a window's visible region — which is what it is erased within — is the
            // desktop's, and a window off it has none to be asked about.
            let _ = ShowWindow(parent, SW_SHOWNOACTIVATE);
            assert!(
                IsWindowVisible(parent).as_bool(),
                "a window off the desktop has no visible region to be erased within"
            );
            let _ = InvalidateRect(parent, None, true);
            let _ = UpdateWindow(parent);
            stands.push((parent, child, clipped));
        }
    }

    let clips = ERASE_CLIPS.lock().unwrap_or_else(|held| held.into_inner());
    let erases_of = |parent: HWND| {
        clips
            .iter()
            .filter(|(window, ..)| *window == parent.0 as isize)
            .map(|(_, inside, outside, _)| (*inside, *outside))
            .collect::<Vec<_>>()
    };
    let clipped = erases_of(stands[0].0);
    let unclipped = erases_of(stands[1].0);
    // Every erase of it, because a window is erased in pieces and an erase that covered its child
    // in any of them is a fill over the film.
    assert!(
        !clipped.is_empty() && clipped.iter().all(|(inside, _)| !inside),
        "no erase of a window wearing the bit covers its child: {clipped:?}"
    );
    assert!(
        clipped.iter().any(|(_, outside)| *outside),
        "and the rest of its client box is erased: {clipped:?}"
    );
    assert!(
        unclipped.iter().any(|(inside, outside)| *inside && *outside),
        "while a window wearing no such bit is erased right across its child, which is the bit doing          it: {unclipped:?}"
    );

    // **And a window that grows is erased of the strip it has gained.** The band a drag grows is
    // filled by this message and nothing else: the box is put wider and taller, and the erase that
    // follows it is asked which of the box it covered.
    let (grown, _, _) = stands[0];
    let counted = clips.len();
    drop(clips);
    // SAFETY: the window is this test's own, and the box is inside the desktop it stands on.
    unsafe {
        let _ = SetWindowPos(
            grown,
            HWND::default(),
            0,
            0,
            480,
            360,
            SWP_NOZORDER | SWP_NOACTIVATE,
        );
        let _ = UpdateWindow(grown);
    }
    let clips = ERASE_CLIPS.lock().unwrap_or_else(|held| held.into_inner());
    let after = clips[counted..]
        .iter()
        .filter(|(window, ..)| *window == grown.0 as isize)
        .map(|(_, inside, _, gained)| (*inside, *gained))
        .collect::<Vec<_>>();
    assert!(
        after.iter().any(|(_, gained)| *gained),
        "a grown window is handed an erase of the strip it gained, which is the message the fill          answers: {after:?}"
    );
    assert!(
        after.iter().all(|(inside, _)| !inside),
        "and none of it covers the child it grew around: {after:?}"
    );
    drop(clips);

    for (parent, child, _) in stands {
        // SAFETY: the two windows are this test's own, and the child goes with its parent.
        unsafe {
            let _ = DestroyWindow(child);
            let _ = DestroyWindow(parent);
        }
    }
}

/// A surface of `width` by `height` pixels in one colour, for an erase to be handed: what is left in
/// it afterwards is the whole of what an erase did.
fn a_surface_of(width: u32, height: u32, colour: [u8; 4]) -> DibSurface {
    let surface = DibSurface::create(width, height).expect("a surface to erase into");
    // SAFETY: the surface was made by this call, of the size asked for, and nothing else holds a
    // pointer to what is behind it.
    unsafe {
        let bits =
            std::slice::from_raw_parts_mut(surface.bits(), width as usize * height as usize * 4);
        for pixel in bits.as_chunks_mut::<4>().0 {
            pixel.copy_from_slice(&colour);
        }
    }
    surface
}
/// **A hover's own host is erased black whatever a pin is holding.** The two hosts share one erase
/// road — a pin taken up on a hover's running film keeps that window (see `hover_film_host`) — and
/// the road paints the picture only for the window the pin is drawing through, so a hover's band is
/// the black it always was.
#[test]
fn a_hovers_own_host_is_erased_black_with_a_pins_frame_in_hand() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    // No pin and no pin's host: this window is a hover's, and nothing else's.
    unsafe { pin_film_host::take_down_the_pin_film_host() };
    let host = unsafe { hover_film_host::take_up_the_hover_film_host((0, 0, 64, 48)) }
        .expect("the hover's host");

    let mut client = RECT::default();
    assert!(
        unsafe { GetClientRect(host, &mut client) }.is_ok(),
        "the hover's host has a client box to be erased"
    );
    let surface = a_surface_of(
        client.right.max(1) as u32,
        client.bottom.max(1) as u32,
        [0u8, 0, 255, 255],
    );

    assert_eq!(
        unsafe {
            SendMessageW(
                host,
                WM_ERASEBKGND,
                WPARAM(surface.dc.0 as usize),
                LPARAM(0),
            )
        },
        LRESULT(1),
        "the hover's host answers its own erase too"
    );
    assert!(
        surface
            .pixels()
            .as_chunks::<4>()
            .0
            .iter()
            .all(|pixel| pixel[..3] == [0, 0, 0]),
        "and the pin's frame is not painted into a hover's band, which stays black"
    );

    unsafe { hover_film_host::take_down_the_hover_film_host() };
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

    // And a host that is not on screen yet is still placed at the band it is given: the picture's
    // window is never hidden with a pin up, so a placement is a placement and never a show.
    let taller = (wider.0, wider.1, wider.0 + 400, wider.1 + 500);
    ensure_pinned_sibling_box(taller);
    assert_eq!(
        unsafe { window_rect(host) },
        taller,
        "the host stands where the band is, shown or not"
    );
}
/// A hover's film drawn in the app-owned top-level host, with the player's own window adopted as the
/// source the Shell draws into it: the state a take-up can promote into the pin rather than begin
/// again.
struct HostedHover {
    host: HWND,
    source: HWND,
    pid: u32,
}

impl HostedHover {
    unsafe fn up(film: &StandingFilm, path: &Path, at: f64) -> HostedHover {
        let pid = unsafe { GetCurrentProcessId() };
        let host = unsafe { hover_film_host::take_up_the_hover_film_host(film.box_) }
            .expect("the hover's own host");
        hover_film_host::note_hover_film_host_player(pid, path);
        let source = unsafe { a_source_of_the_host(host) };
        unsafe { hover_film_host::reveal_the_hover_film_host(pid, film.box_) };
        note_video_loop(pid, path, film.box_, at, None);
        VIDEO_PID.store(pid, Ordering::SeqCst);
        HostedHover { host, source, pid }
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
/// An FFmpeg hover that was already in the app's host is taken up the same way mpv's is: the same
/// pid, the same source, the same host handle, promoted into the pin with nothing begun again. This
/// is the seamless hover-to-pin take-up: no relaunch, no black flash.
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
        unsafe { IsWindow(hover.source) }.as_bool(),
        "the player's window is the same window, still drawing"
    );
    assert_eq!(
        pin_film_thumbnail::the_source_window(),
        Some(hover.source),
        "and it is the picture the Shell is drawing into the pin's host"
    );
    assert_eq!(
        unsafe { GetParent(hover.source) }.ok(),
        None,
        "the source is still a window of the desktop's: a promotion reparents nothing"
    );
    assert_eq!(
        unsafe { GetParent(hover.host) }.ok(),
        None,
        "and neither is the host, which is the pin's window now"
    );
    assert!(
        unsafe { IsWindowVisible(hover.host) }.as_bool(),
        "the taken-up host is shown in the pin's band"
    );
    assert_eq!(VIDEO_HWND.load(Ordering::SeqCst), 0);
    assert_eq!(
        hover_film_host::hover_film_host(),
        None,
        "the hover's handle is handed over rather than left to destroy the pin's host"
    );
}

/// **A take-up on a hover's mpv film takes the hover's own host up as the pin's host**: the same
/// window, the same player, the same child, and no second decode.
///
/// The fault this is a regression test for: the take-up on the road where the pin's chrome is owned
/// by its film's host refused to re-parent a hover's host into the pin — that was the airspace the
/// road exists to be rid of — and fell back to freezing the hover and beginning the pin's own player.
/// The pin appeared a pause after the key, and the file was opened twice. The hover's host is already
/// the window this road wants (top-level, not layered, holding the player this app began into it), so
/// the take-up promotes it instead: a style, an owner, a band and a box change, and the picture on
/// the screen is the same player's from the key to the film.
#[test]
fn an_owned_pin_takes_a_hovered_film_up_by_promoting_its_host() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let hover = unsafe { HostedHover::up(&film, &path, 9.5) };
    clear_restart_count();
    clear_pending_pinned_relaunch();
    pin_film_host::disarm_pin_handover();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    let _ = take_last_spawned_command();

    unsafe { show_pinned_window(film.pin_hwnd) };

    assert_eq!(
        restart_count(),
        0,
        "the hover's player is the pin's player: nothing is begun again behind it"
    );
    assert!(
        take_last_spawned_command().is_none(),
        "and no player process is begun for the take-up at all"
    );
    assert_eq!(
        pin_film_host::film_host(),
        Some(hover.host),
        "the hover's own host is the pin's host now"
    );
    assert_eq!(
        pin_film_host::hosted_player_pid(),
        Some(hover.pid),
        "holding the player the hover began"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        hover.pid,
        "which is still the player on screen"
    );
    assert!(
        pin_film_host::the_pinned_player_was_begun_in_the_host(),
        "and it is the host's own from here"
    );
    assert!(
        unsafe { IsWindow(hover.source) }.as_bool(),
        "the player's window is the same window, still drawing"
    );
    assert_eq!(
        pin_film_thumbnail::the_source_window(),
        Some(hover.source),
        "and it is the one picture the promoted host draws"
    );

    // The host a hover had is the host a pin has: a window of the desktop's, the Shell's to list, and
    // not the band nothing may cover or the press that falls through it.
    assert_eq!(
        unsafe { GetWindowLongPtrW(hover.host, GWL_STYLE) } & WS_CHILD.0 as isize,
        0,
        "a window of the desktop's, not a child of the pin's window"
    );
    let ex = unsafe { GetWindowLongPtrW(hover.host, GWL_EXSTYLE) };
    assert_eq!(
        ex & WS_EX_APPWINDOW.0 as isize,
        WS_EX_APPWINDOW.0 as isize,
        "the window the Shell lists for the pin"
    );
    assert_eq!(
        ex & (WS_EX_TRANSPARENT.0 | WS_EX_TOOLWINDOW.0) as isize,
        0,
        "with the press-through it wore for a hover, and its place out of the Shell's lists, gone"
    );
    assert_eq!(
        ex & WS_EX_LAYERED.0 as isize,
        0,
        "and not layered, which is the whole of why the pin's host is a window of its own"
    );
    assert_eq!(
        unsafe { GetWindow(film.pin_hwnd, GW_OWNER) },
        Ok(hover.host),
        "the pin's chrome is the host's own window"
    );
    assert_eq!(
        unsafe { window_rect(hover.host) },
        unsafe { window_rect(film.pin_hwnd) },
        "and the host is the band the pin's window is at"
    );
    assert_eq!(
        unsafe { window_holds_the_topmost_band(hover.host) },
        unsafe { window_holds_the_topmost_band(film.pin_hwnd) },
        "the pair is in one band, whichever band the setting names"
    );
    assert!(
        unsafe { IsWindowVisible(hover.host) }.as_bool(),
        "the picture never left the screen"
    );

    // Nothing of the relaunch road ran: no player was held frozen, no handover waited, no pause was
    // sent and no frame read back.
    assert!(
        !pin_film_host::pin_handover_armed(),
        "no wait for a second player's first frame"
    );
    assert_eq!(
        pin_film_host::held_hover_for_a_test(),
        None,
        "and no frozen hover behind the pin"
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
    assert_eq!(
        hover_film_host::hover_film_host(),
        None,
        "the hover's handle is handed over rather than left to destroy the pin's host"
    );

    // And a show of the pin already standing over its film — a box change, the tick's restore — is not
    // a second launch either.
    unsafe { show_pinned_window(film.pin_hwnd) };
    assert_eq!(
        restart_count(),
        0,
        "a pin already standing over its own player is not a film begun again behind it"
    );
    assert!(
        take_last_spawned_command().is_none(),
        "and nothing is spawned for it"
    );

    pin_film_host::disarm_pin_handover();
    clear_pending_pinned_relaunch();
    clear_restart_count();
    drop(hover);
    drop(film);
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
    // **A hover the pin cannot take up**: the host holds its player, and that player has no source,
    // which is a film whose window is not up yet (see `the_hover_player_is_drawing`).
    let hovered = 0x7FFF_FFFB;
    unsafe { hover_film_host::take_up_the_hover_film_host((0, 0, 320, 180)) }
        .expect("the hover's host");
    hover_film_host::note_hover_film_host_player(hovered, &path);
    VIDEO_PID.store(hovered, Ordering::SeqCst);
    assert!(
        hover_film_host::the_hover_film_to_adopt(&path).is_none(),
        "a player with no source to move has nothing to move"
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
        hovered,
        "and the hover's player is not the one the pin now holds"
    );
    unsafe { hover_film_host::take_down_the_hover_film_host() };
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
    // **The pin holds the front, so its player is owed no band ask** (see `mpv_ipc::set_ontop`):
    // what this test is about is the one line the frozen hover is told.
    let always_on_top = CONFIG.lock().ok().map(|config| config.pin_always_on_top);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
    }

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
    if let Ok(mut config) = CONFIG.lock() {
        if let Some(always_on_top) = always_on_top {
            config.pin_always_on_top = always_on_top;
        }
    }
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
    let _child = unsafe { a_source_of_the_host(host) };
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
    let _child = unsafe { a_source_of_the_host(host) };
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

    let child = unsafe { a_source_of_the_host(host) };
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

/// A hosted film's resize asks the surface for no row of slack, and a pin with no host still does.
/// The slack's whole mechanism is the *window* moving a pixel on every paint, so that the desk under
/// a strip a shrink gives up is drawn again — and a film's host is a child of that window, which
/// rides the move and bobs at paint rate with it (see
/// `a_resize_asks_the_surface_for_a_row_of_slack`). A native strip's picture is this app's own
/// pixels, so it lands on the box the hand asked for either way and keeps the ghost-clearing row.
#[test]
fn a_resize_of_a_pin_with_a_film_host_asks_the_surface_for_no_row_of_slack() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();

    // A pin with no host standing is the native road: its picture is this app's own pixels and lands
    // on the box the hand asked for either way, so it keeps the row that clears its ghost.
    assert!(
        pin_film_host::film_host().is_none(),
        "a pin that has not been shown has no host yet"
    );
    assert!(paint::a_resize_asks_the_surface_for_a_row_of_slack(
        true,
        pin_film_host::film_host().is_some()
    ));

    unsafe { show_pinned_window(film.pin_hwnd) };
    assert!(
        pin_film_host::film_host().is_some(),
        "the pin's film is drawn in a host"
    );

    assert!(
        !paint::a_resize_asks_the_surface_for_a_row_of_slack(
            true,
            pin_film_host::film_host().is_some()
        ),
        "the row a resize paints with is the row a hosted film's window would bob by"
    );
    assert!(
        !paint::a_resize_asks_the_surface_for_a_row_of_slack(
            false,
            pin_film_host::film_host().is_some()
        ),
        "and a paint that is not a resize asks for nothing at all"
    );
}

/// The topmost road is for players placed from the outside; the host is already where the pin is,
/// and it is the band that says where the picture goes. So the box the road was handed is not the
/// host's, and a host still starting is shown by it — the band is owed its picture back (see
/// `pin_film_host::set_the_pin_film_host_visible`).
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
    let child = unsafe { a_source_of_the_host(host) };

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

/// The wait for a film's first frame is the wait for the player's own window, adopted as the source
/// the Shell draws: the host carries the last film's registration for as long as that film is
/// recorded, so the source has to be the one this player made.
#[test]
fn the_film_is_drawing_when_the_players_own_window_is_the_source() {
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
        "a host with no source recorded has nothing of the film to show yet"
    );

    let source = unsafe { a_source_of_the_host(host) };

    assert!(
        pin_film_host::the_hosted_player_is_drawing(VIDEO_PID.load(Ordering::SeqCst)),
        "a window of the player's own process, adopted as the source, is the film on screen"
    );

    unsafe {
        let _ = DestroyWindow(source);
    };

    assert!(
        !pin_film_host::the_hosted_player_is_drawing(VIDEO_PID.load(Ordering::SeqCst)),
        "and the window going is the film not being drawn anywhere"
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
    let child = unsafe { a_source_of_the_host(host) };
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
    let _child = unsafe { a_source_of_the_host(host) };
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
    let _child = unsafe { a_source_of_the_host(host) };

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

/// A swap onto a hosted film is answered by the source adopted into the host, not by a top-level
/// handle a hosted player never publishes: the file is installed on the source's arrival rather than
/// let go of at the cover's bound (see `PinSwapHold::settle`).
#[test]
fn a_hosted_swaps_wait_is_answered_by_the_sources_arrival() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    // The source this test stands is the test's own process's, so the player the host holds has to
    // be this process as well (see `the_hosted_player_is_drawing`).
    let drawn = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(drawn);
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_source_of_the_host(host) };

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
/// **On the thumbnail road the pin's film is a window of the desktop's and the host only shows it**:
/// the launch names no window of this app's, and the command line it is given carries the box the
/// player's own window is made at rather than the order `--ontop` would hold it in over the pin.
#[test]
fn a_pin_drawn_through_a_thumbnail_is_begun_with_no_window_of_this_apps() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    unsafe { show_pinned_window(film.pin_hwnd) };
    pin_film_host::film_host().expect("the film's host");

    let path = Path::new("a-film.mkv");
    let args = video_launch::mpv_args(
        path, 0, 0, 320, 240, 0.0, false, 80, None, false, false, None, "rhp-test", false, false,
    );
    assert!(
        args.iter().any(|arg| arg == "--geometry=320x240+0+0"),
        "the box the player's own window is made at is still named: {args:?}"
    );
    assert!(
        !args.iter().any(|arg| arg == "--ontop"),
        "and the film is not held over the pin it is seen in: {args:?}"
    );
    assert!(
        !args.iter().any(|arg| arg.starts_with("--wid")),
        "and no window of this app's is named, because the player makes its own: {args:?}"
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

/// The film's chrome is drawn over its picture, and the host the film is drawn in composites over
/// the layered surface the pin paints — so the film stood over the strips the moment it began to
/// draw, painted and interactive and invisible. The chrome therefore has windows of its own, layered
/// children of the pin's window, which composite over the host carrying the very pixels the pin
/// paints for them (see `pin_chrome_strips`).
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

    // **Above the host**: the host is a window of the desktop's that owns the pin's window, and the
    // strips are that window's children — so the whole pin, chrome included, stands over the picture
    // rather than the other way round.
    assert_eq!(
        unsafe { GetWindow(film.pin_hwnd, GW_OWNER) },
        Ok(host),
        "the host owns the pin's window, which is what keeps the chrome over the film"
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
    assert!(pin_film_host::film_host().is_none(), "the host is gone");
    assert!(
        unsafe { IsWindow(caption) }.as_bool() && unsafe { IsWindow(transport) }.as_bool(),
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
    let caption = pin_chrome_strips::caption_strip().expect("the strips are made with the window");
    let transport = pin_chrome_strips::transport_strip().expect("both of them");
    let _child = unsafe { a_source_of_the_host(host) };
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
    let _child = unsafe { a_source_of_the_host(host) };
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
    unsafe { render_pinned_preview_at(film.pin_hwnd, film.box_.0, film.box_.1) };

    assert!(
        unsafe { IsWindowVisible(volume) }.as_bool(),
        "an open popup over a hosted film is a popup the strip carries"
    );
    assert_eq!(
        unsafe { GetWindow(film.pin_hwnd, GW_OWNER) },
        Ok(host),
        "and the strip is a child of the pin, which the host owns: the popup stands over the film"
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
    let surface = unsafe { std::slice::from_raw_parts(bits, width as usize * height as usize * 4) };
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
    unsafe {
        let _ = ShowWindow(film.pin_hwnd, SW_SHOWNOACTIVATE);
    };
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
    let _child = unsafe { a_source_of_the_host(host) };
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
    clear_restart_count();
    clear_pending_pinned_relaunch();
    pin_film_host::disarm_pin_handover();

    unsafe { show_pinned_window(film.pin_hwnd) };

    // Nothing is captured, painted or hidden before the showing: the hover is the placeholder and
    // it never comes down until the handover.
    assert_eq!(
        restart_count(),
        1,
        "the adopted film is still relaunched into the host behind the hover"
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
    clear_pending_pinned_relaunch();
    clear_restart_count();
    drop(film);
}

/// A window that is not the source is not a film drawing: the record is the player's own window
/// adopted for the host, and a window of the same process that nothing adopted is nothing the host
/// has to show — so the host stays off the screen even though a process of its own exists (see
/// `the_hosted_player_is_drawing`).
#[test]
fn a_window_that_is_not_the_source_is_not_a_film_drawing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let drawn = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(drawn);
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");

    // A window the player has made and nothing has adopted: the pass that watches a player has not
    // run over it yet, which is a film whose source is not the host's picture (see
    // `pin_film_thumbnail::adopt_the_source`).
    let unmade = unsafe {
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
    .expect("a window of the player's own, in this process");

    assert!(
        !pin_film_host::the_hosted_player_is_drawing(VIDEO_PID.load(Ordering::SeqCst)),
        "a window nothing adopted is not a source: the film is the one the record names"
    );

    let _ = ensure_video_window_topmost(
        film.box_.0,
        film.box_.1,
        (film.box_.2 - film.box_.0).max(1),
        (film.box_.3 - film.box_.1).max(1),
    );
    assert!(
        !unsafe { IsWindowVisible(host) }.as_bool(),
        "and the tick's road leaves the host off the screen until a source is adopted"
    );

    unsafe {
        let _ = DestroyWindow(unmade);
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
    let painted = unsafe { std::slice::from_raw_parts(bits, width as usize * height as usize * 4) };
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
    let _child = unsafe { a_source_of_the_host(host) };

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
    clear_restart_count();
    clear_pending_pinned_relaunch();
    pin_film_host::disarm_pin_handover();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    unsafe { show_pinned_window(film.pin_hwnd) };

    let drawn = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(drawn);
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_source_of_the_host(host) };

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
    clear_restart_count();
    clear_pending_pinned_relaunch();
    pin_film_host::disarm_pin_handover();
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    unsafe { show_pinned_window(film.pin_hwnd) };
    // The chrome hidden across the swap, which is what a step leaves behind.
    with_pin(|pin| pin.chrome = PinChrome::hidden());

    let drawn = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(drawn);
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_source_of_the_host(host) };
    the_hosted_player_has_the_frame(drawn, "0.000000");

    assert!(
        matches!(
            unsafe { pin_film_host::settle_pin_handover() },
            pin_film_host::PinHandover::HandedOver
        ),
        "the tick the new film draws hands the band over"
    );
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| (pin.chrome.caption, pin.chrome.bar))),
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
    drop(film);
}
/// A stand whose film FFmpeg's player draws, rather than mpv's: the route where the player's own
/// window has to be adopted into the pin's host after SDL makes it.
fn ffplay_standing_film() -> StandingFilm {
    let film = StandingFilm::new();
    // **FFmpeg's player answers nothing** — it is begun with no pipe at all — so the frame of an
    // FFmpeg film is its source and not a property (see `mpv_ipc::answers_properties`).
    mpv_ipc::forget_pipe(unsafe { GetCurrentProcessId() });
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

/// A pinned FFmpeg film is kept for the pin's host and publishes no top-level window of its own: the
/// monitor adopts the player's own window as the source the host draws (see
/// `start_video_playback_where` and `ffplay_window::try_apply_noactivate_style`).
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

    let process = start_video_playback(
        &path,
        film.box_.0,
        film.box_.1,
        320,
        240,
        0.0,
        false,
        80,
        None,
        None,
    )
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

/// A pinned FFmpeg film is begun with the **composed** renderer — SDL's Direct3D 9, whose blt-model
/// present the desktop compositor copies — and never with the flip-model Direct3D 11 one, whose
/// swapchain the display's adaptive sync follows: that is what made the whole panel flicker while a
/// pinned film played (see `video_playback::the_renderer_a_film_is_begun_with`). Both launches ask
/// for no border, because the source is never externally resized and its client is the box the crop
/// is cut from (see `pin_film_thumbnail`).
#[test]
fn every_ffplay_launch_asks_for_the_composed_renderer_and_a_borderless_window() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let path = PathBuf::from("a-film.mkv");

    let _ = start_video_playback(
        &path,
        film.box_.0,
        film.box_.1,
        320,
        240,
        0.0,
        false,
        80,
        None,
        None,
    )
    .expect("a noted FFmpeg player");
    let launched = take_last_spawned_command().expect("the pin's launch");
    assert!(launched.program.contains("ffplay"), "the pin's player");
    assert!(
        launched.args.iter().any(|arg| arg == "-noborder"),
        "a pinned launch is borderless, so the window's box is its client"
    );
    let composed = video_playback::the_renderer_asked_for(None);
    assert!(
        launched.env.iter().any(|(key, value)| {
            key.eq_ignore_ascii_case("SDL_RENDER_DRIVER")
                && value.as_deref() == Some(composed.as_str())
        }),
        "and asks for the composed renderer, the one no display's adaptive sync can follow"
    );

    // The same film with no pin and Pin Mode off is an unhosted hover: the same renderer, and the
    // borderless window it has always had (see `hover_film_host_window` for the hosted hover's own
    // launch).
    take_pin_for_a_test();
    pin_stands_in_its_own_window(false);
    let pin_enabled_before = CONFIG
        .lock()
        .map(|config| config.pin_enabled)
        .unwrap_or(true);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_enabled = false;
    }
    let _ = start_video_playback(
        &path,
        film.box_.0,
        film.box_.1,
        320,
        240,
        0.0,
        false,
        80,
        None,
        None,
    )
    .expect("a noted FFmpeg hover");
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_enabled = pin_enabled_before;
    }
    let hovered = take_last_spawned_command().expect("the hover launch");
    assert!(
        hovered.args.iter().any(|arg| arg == "-noborder"),
        "and so does the hover's: the source is never resized on either road"
    );
    assert!(
        hovered.env.iter().any(|(key, value)| {
            key.eq_ignore_ascii_case("SDL_RENDER_DRIVER")
                && value.as_deref() == Some(composed.as_str())
        }),
        "and is begun with the composed renderer too: its own window is a top-level one"
    );
}

/// The renderer a film is begun with is the composed one, and it is asked for by name rather than left
/// to SDL's own driver order: the composed road is this app's requirement, not a fallback (see
/// `video_playback::the_renderer_a_film_is_begun_with`). The one way it is something else is the
/// escape hatch a machine names by hand, which is passed through whole so `software`, `opengl` and
/// `direct3d11` mean what SDL says they mean — a blank name is no name at all.
#[test]
fn the_composed_renderer_is_asked_for_unless_a_machine_names_another() {
    assert_eq!(
        video_playback::the_renderer_asked_for(None),
        "direct3d",
        "the composed renderer, whose present the desktop compositor copies"
    );
    assert_eq!(
        video_playback::the_renderer_asked_for(Some("  ")),
        "direct3d",
        "a blank name is no name, and the composed road stands"
    );
    assert_eq!(
        video_playback::the_renderer_asked_for(Some(" software ")),
        "software",
        "a name given by hand is passed through, trimmed, to SDL"
    );
    assert_eq!(
        video_playback::the_renderer_asked_for(Some("direct3d11")),
        "direct3d11",
        "including the flip-model one, which is the trade the hatch exists to make deliberately"
    );
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

/// A pin that ends takes its FFmpeg film host down with it, and the picture the Shell drew goes
/// before it: the source is the player's own window and outlives the pin, but no registration and no
/// record of it is left behind (see `pin_film_thumbnail`).
#[test]
fn a_pin_that_ends_takes_its_ffplay_film_host_with_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let pid = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(pid);
    let source = unsafe { a_source_of_the_host(host) };

    unsafe { pin_film_host::take_down_the_pin_film_host() };

    assert!(
        !unsafe { IsWindow(host) }.as_bool(),
        "the host goes down with the pin"
    );
    assert!(
        !pin_film_thumbnail::the_source_is_up(pid),
        "and the picture it was showing with it"
    );
    assert!(pin_film_host::film_host().is_none());
    assert!(pin_film_host::hosted_player_pid().is_none());

    unsafe { DestroyWindow(source) }.expect("the source's window goes");
}

/// A hosted player's window is a source and not a window of the desktop's to publish: the top-level
/// roads neither enumerate it nor style it as one, and the invisibility the adoption put on it is
/// left exactly as it was (see `pin_film_thumbnail::adopt_the_source`).
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
    let source = unsafe { a_source_of_the_host(host) };

    assert!(
        !unsafe { try_apply_noactivate_style(pid) },
        "the monitor finds the host's source rather than a window of the desktop's to publish"
    );
    // **The one thing the monitor does to a source is keep it invisible** (see
    // `pin_film_thumbnail::keep_the_source_invisible`), which is what a window of the desktop's is
    // never given.
    let ex = unsafe { GetWindowLongPtrW(source, GWL_EXSTYLE) };
    assert_eq!(
        ex & WS_EX_LAYERED.0 as isize,
        WS_EX_LAYERED.0 as isize,
        "the source is kept invisible rather than published"
    );
    assert!(
        !unsafe { apply_noactivate_to_hwnd(source) },
        "and a top-level styling of the source is refused"
    );
    assert_eq!(unsafe { GetWindowLongPtrW(source, GWL_EXSTYLE) }, ex);
    assert_eq!(VIDEO_HWND.load(Ordering::SeqCst), 0);

    unsafe { DestroyWindow(source) }.expect("the source's window goes");
}

/// FFmpeg's player has no pipe to say its video output is configured, so its adopted source is the
/// frame's proof: that window is the player's whole picture and is adopted only after the player
/// made and showed it (see `the_hosted_frame_is_up`). This is the fresh-pin, step and seek relaunch
/// road — without it the chrome would stay hidden for the life of the pin.
#[test]
fn a_hosted_ffplay_films_frame_is_its_source_and_the_chrome_comes_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let pid = VIDEO_PID.load(Ordering::SeqCst);
    let host = pin_film_host::film_host().expect("the film's host");
    let _child = unsafe { a_source_of_the_host(host) };
    with_pin(|pin| {
        pin.chrome = PinChrome::hidden();
        pin.transport_bar = true;
    });
    stand_window_at_point_in(Some(film.pin_hwnd.0 as isize));
    let window = film.box_;
    let caption_cursor = Some((window.0 + 5, window.1));

    assert!(
        pin_film_host::the_hosted_player_is_drawing(pid),
        "the FFmpeg player's own window is the source the host draws"
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
    let _child = unsafe { a_source_of_the_host(host) };
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
    let _child = unsafe { a_source_of_the_host(host) };

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
    let _child = unsafe { a_source_of_the_host(host) };
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
/// A step onto an FFmpeg film begins one player, in the pin's host, and the show that follows
/// takes that player up rather than beginning a second one at zero.
///
/// **The fault this is a regression test for.** The swap takes the standing media down before the
/// incoming film is begun, and the host question used to read the emptied slot: the first launch
/// was unhosted, and the take-up that followed began a second player in the host at the start.
#[test]
fn a_step_into_an_ffplay_film_begins_it_in_the_pins_host_once() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    let _host = pin_film_host::film_host().expect("the film's host");

    // No standing player: this is the swap's own launch, and the source it is drawn from arrives
    // with the player the install begins.
    VIDEO_PID.store(0, Ordering::SeqCst);
    clear_restart_count();
    let _ = take_last_spawned_command();

    let next = PathBuf::from("the-next-film.mkv");
    let mut media = create_loading_media(320, 240);
    media.media_type = MediaType::Video;
    let swap = swap_pinned_media(PinAnswer {
        path: next.clone(),
        update: PinUpdate {
            content: film.box_,
            dpi: 96,
            volume: 0,
        },
        media: Some(media),
        walk: None,
        arc: PinArc::new(),
    });
    let file = match swap {
        PinSwap::Ready(file) => file,
        PinSwap::Holding(hold) => hold.file,
        PinSwap::Refused { .. } => panic!("the next film is a film"),
    };

    // What the install writes on the tick the swap answers: the file is the pin's own and its
    // media is what the take-up that follows is about (see `install_pinned_media`).
    with_pin(|pin| pin.path = next);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(file.media);
    }

    let launched = take_last_spawned_command().expect("the step begins one player");
    assert!(
        launched.args.iter().any(|arg| arg == "-noborder"),
        "the step's one launch is hosted and borderless, like every launch of a sourced film"
    );
    let composed = video_playback::the_renderer_asked_for(None);
    assert!(
        launched.env.iter().any(|(key, value)| {
            key.eq_ignore_ascii_case("SDL_RENDER_DRIVER")
                && value.as_deref() == Some(composed.as_str())
        }),
        "and it is begun with the composed renderer, like every other FFmpeg film"
    );

    unsafe { show_pinned_window(film.pin_hwnd) };

    assert!(
        pin_film_host::the_pinned_player_was_begun_in_the_host(),
        "the step's player is the pin's hosted one"
    );
    assert_eq!(
        restart_count(),
        0,
        "the show takes that player up rather than beginning a second at zero"
    );
    assert!(
        take_last_spawned_command().is_none(),
        "and no second player is begun"
    );
}

/// A pin put away into the taskbar leaves the film standing for the still the Shell shows of a
/// minimized window: only a handover host that is still a window of the desktop's is hidden.
///
/// **The fault this is a regression test for.** The FFmpeg arm hid the hosted picture before the
/// minimize, so the last frame the Shell had to show was an empty band.
fn a_pin_put_away_in_the_taskbar_leaves_its_film_standing(ffplay: bool) {
    let film = if ffplay {
        ffplay_standing_film()
    } else {
        StandingFilm::new()
    };
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    film.drawn_in_the_host(unsafe { GetCurrentProcessId() });
    let _child = unsafe { a_source_of_the_host(host) };
    unsafe {
        let _ = ShowWindow(host, SW_SHOWNOACTIVATE);
    }
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool(),
        "the film's host is the picture the pin is showing"
    );

    collapse_pin_to(PinMinimizeTo::Taskbar);

    // **The window the Shell lists** is the film's host, whichever engine draws it: it owns the
    // pin's chrome and carries the button (see `pin_film_host::own_the_pin_window`).
    assert!(
        unsafe { IsIconic(host) }.as_bool(),
        "the window the Shell lists is the one it minimizes"
    );
    assert!(
        unsafe { IsWindowVisible(host) }.as_bool(),
        "and the film is left standing for the still the Shell shows of it"
    );
}

#[test]
fn a_ffplay_pin_put_away_in_the_taskbar_leaves_its_film_standing_for_the_shells_still() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    a_pin_put_away_in_the_taskbar_leaves_its_film_standing(true);
}

#[test]
fn an_mpv_pin_put_away_in_the_taskbar_leaves_its_film_standing_for_the_shells_still() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    a_pin_put_away_in_the_taskbar_leaves_its_film_standing(false);
}

/// A pin's own player that a step began as a window of the desktop's is adopted by the show that
/// follows, at the same pid, rather than begun again at zero.
#[test]
fn a_show_adopts_the_pin_player_a_step_began_unhosted() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = ffplay_standing_film();
    unsafe { show_pinned_window(film.pin_hwnd) };
    VIDEO_PID.store(0, Ordering::SeqCst);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = None;
    }

    let path = PathBuf::from("a-stranded-film.mkv");
    let process = start_video_playback(&path, 0, 0, 320, 240, 0.0, false, 80, None, None)
        .expect("a noted FFmpeg player");
    let pid = process.id();
    let mut media = create_loading_media(320, 240);
    media.media_type = MediaType::Video;
    media.video_process = Some(process);
    with_pin(|pin| pin.path = path);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }
    clear_restart_count();
    let _ = take_last_spawned_command();

    unsafe { show_pinned_window(film.pin_hwnd) };

    assert_eq!(
        restart_count(),
        0,
        "the show adopts the pin's own player rather than beginning it again"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        pid,
        "the adopted player is the same process"
    );
    assert!(
        pin_film_host::the_pinned_player_was_begun_in_the_host(),
        "and it is the host's player from then on"
    );
    assert!(
        take_last_spawned_command().is_none(),
        "nothing is spawned for the take-up"
    );
}

/// **The road where the host owns the pin's chrome**: the film's host is a window of the desktop's
/// rather than a child of the pin's, and the pin's own window is owned by it — which is the whole of
/// what keeps the chrome over the film without this app moving either of them (see
/// `pin_film_host::the_pins_chrome_is_owned`).
///
/// A `WS_CHILD` is always drawn above its parent's client area, and a layered parent cannot clip a
/// hardware swapchain child; a window of the desktop's below an owned pin has the OS keep the order
/// instead, so this asserts both halves of that pair rather than the picture, which no test can see.
#[test]
fn an_owned_pin_gets_a_window_of_the_desktop_that_owns_its_chrome() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();

    unsafe { pin_film_host::take_up_the_pin_film_host(None) };
    let host = pin_film_host::film_host().expect("the film's host");

    let style = unsafe { GetWindowLongPtrW(host, GWL_STYLE) };
    assert_eq!(
        style & WS_CHILD.0 as isize,
        0,
        "the host is not a child, which is the whole of what was wrong with the one it replaces"
    );
    assert_eq!(
        style & WS_POPUP.0 as isize,
        WS_POPUP.0 as isize,
        "it is a popup: a window of the desktop's at the band the film stands in"
    );
    let ex = unsafe { GetWindowLongPtrW(host, GWL_EXSTYLE) };
    assert_eq!(
        ex & WS_EX_LAYERED.0 as isize,
        0,
        "and it is not layered: the film's swapchain needs a real window under it"
    );
    assert_eq!(
        ex & WS_EX_APPWINDOW.0 as isize,
        WS_EX_APPWINDOW.0 as isize,
        "it is the window the Shell lists and gives a button to, where the chrome it owns is not"
    );

    assert_eq!(
        unsafe { GetWindow(film.pin_hwnd, GW_OWNER) },
        Ok(host),
        "the pin's chrome is the host's own window"
    );
    assert!(
        unsafe { GetWindowLongPtrW(film.pin_hwnd, GWL_EXSTYLE) } & WS_EX_APPWINDOW.0 as isize == 0,
        "and the chrome asks for no taskbar button of its own, because a window it owns carries one"
    );
    assert_eq!(
        unsafe { window_rect(host) },
        film.box_,
        "the host stands at the pin's band, on the screen, rather than in its client coordinates"
    );

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    assert_eq!(
        unsafe { GetWindow(film.pin_hwnd, GW_OWNER) },
        Err(windows::core::Error::from_win32()),
        "and a host that goes gives the chrome back rather than taking it down with it"
    );
    assert!(
        unsafe { GetWindowLongPtrW(film.pin_hwnd, GWL_EXSTYLE) } & WS_EX_APPWINDOW.0 as isize != 0,
        "the pin is the Shell's window again"
    );
    assert!(
        unsafe { IsWindow(film.pin_hwnd) }.as_bool(),
        "and it is a window the hand can still carry, size and close"
    );

    drop(film);
}

/// **The host is the band at every step of a grow and a shrink**, whichever of the two the step
/// committed first: a window of the desktop's is not clipped by the pin, so a host left at the box
/// the step came from is a film standing outside the chrome it belongs to (see `apply_pin_drag` and
/// `pin_film_host::the_film_host_box`).
#[test]
fn an_owned_pins_host_is_the_band_at_every_step_of_a_grow_and_a_shrink() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let _stand_in = POINTER_STAND_IN.lock().expect("the pointer's own state");
    let film = StandingFilm::new();
    unsafe { show_pinned_window(film.pin_hwnd) };
    // A step places the pin on the screen its box is clamped to: the window is hidden so the test
    // owes the desktop no picture.
    unsafe {
        let _ = ShowWindow(film.pin_hwnd, SW_HIDE);
    }
    let host = pin_film_host::film_host().expect("the film's host");

    let window = a_window_at(film.box_);
    begin_pin_drag(
        film.pin_hwnd,
        &window,
        PinDragAction::Resize(edge(false, false, true, false)),
        true,
    );

    let mut width = film.box_.2 - film.box_.0;
    for (travel, growing) in [(120, true), (-160, false)] {
        super::probes::step_the_pin_drag_by(travel, 0);
        unsafe { apply_pin_drag(film.pin_hwnd) };

        let band = pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| pin.content))
            .expect("a pin up");
        let now = band.2 - band.0;
        assert_eq!(
            now > width,
            growing,
            "the step of {travel} took the window {} rather than the other way",
            if growing { "out" } else { "in" }
        );
        width = now;

        assert_eq!(
            unsafe { window_rect(film.pin_hwnd) },
            band,
            "the pin's own window is the box the step asked for"
        );
        assert_eq!(
            unsafe { window_rect(host) },
            band,
            "and the host is that band on the screen, and not the box the step came from"
        );
    }

    with_pin(|pin| pin.dragging = None);
    drop(film);
}
/// **A pin and the film it owns stand in one band, whichever window of the pair is asked.**
///
/// The fault this is a regression test for: the band was asked of the pin's chrome alone, and
/// Windows does not carry ownership's band upwards — a chrome put in the topmost band left the host
/// that owns it in the ordinary one, and a chrome given the ordinary band left the host where it was.
/// The pair was then half over another program's window and half behind it: the film under the
/// program, and the chrome above it, taking every press in a band it is nearly transparent in. That
/// is a pin mode window that looks behind a window and yet answers the hand.
///
/// What is asserted here is the pair as one unit: the band bits of the host and the chrome always
/// agree, the band is taken and given back through the host, and both windows of the pair stand on
/// the same side of a covering window — which is the order a press is delivered by, and the whole of
/// what decides whether a click over the covered band is the pin's or that window's. The press
/// itself, and `WindowFromPoint` over a desktop with somebody else's windows on it, is more than this
/// suite can answer: what a press does with the order it is given is Windows'.
#[test]
fn an_owned_pin_and_its_film_stand_in_one_band() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let always_on_top = CONFIG
        .lock()
        .map(|config| config.pin_always_on_top)
        .unwrap_or(true);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
    }

    // The pin's band somewhere on the screen, and a picture drawing in the host it is taken up with:
    // the band is alpha 1 over a film in a window under it, so a press in it is the pin's only while
    // the pin is the window Windows names at that point.
    let at = (420, 320, 760, 600);
    with_pin(|pin| {
        pin.content = at;
        pin.overlay = true;
        pin.caption = 0;
    });
    let drawn = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(drawn);
    unsafe { pin_film_host::take_up_the_pin_film_host(None) };
    let host = pin_film_host::film_host().expect("the film's host");
    let child = unsafe { a_source_of_the_host(host) };
    unsafe { show_pinned_window(film.pin_hwnd) };

    let topmost = |window: HWND| unsafe { window_holds_the_topmost_band(window) };
    assert!(
        topmost(host) && topmost(film.pin_hwnd),
        "a pin that is Always On Top takes the band with the film it owns"
    );

    // Out of the band together: the chrome's band is the owner's, so one call on the host is the
    // whole pair, and the chrome is never left standing where the film is not.
    unsafe { pin_film_host::place_the_pin_in_the_band(film.pin_hwnd, HWND_NOTOPMOST) };
    assert!(
        !topmost(host) && !topmost(film.pin_hwnd),
        "and gives it back with it, rather than the chrome out of it alone"
    );

    // Another program's window opened over the pin. The pair is moved behind it in one call, and the
    // window a press over the band would go to is the one at the point — which is that window, because
    // both windows of the pair are behind it.
    let cover = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP | WS_VISIBLE,
            at.0 + 40,
            at.1 + 40,
            200,
            120,
            None,
            None,
            None,
            None,
        )
    }
    .expect("a window of the test's own standing over the pin");
    unsafe { pin_film_host::place_the_pin_in_the_band(film.pin_hwnd, cover) };
    assert!(
        !unsafe { pin_stands_over(film.pin_hwnd, cover) }
            && !unsafe { pin_stands_over(host, cover) },
        "a pin put behind another window is behind it as one window: a press in the band is that          window's press, not a press on a pin that is not there"
    );

    // And the pin put in its band takes the pair over that window: the press is the pin's again, and
    // the film it is over is on the screen with it.
    unsafe { pin_film_host::place_the_pin_in_the_band(film.pin_hwnd, HWND_TOPMOST) };
    assert!(
        unsafe { pin_stands_over(film.pin_hwnd, cover) } && unsafe { pin_stands_over(host, cover) },
        "and a pin raised into its band raises the film it owns with it"
    );

    unsafe {
        let _ = DestroyWindow(cover);
        let _ = DestroyWindow(child);
    };
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
    }
    drop(film);
}
/// The window of the player's own that a thumbnail samples, on this test's own process: the source a
/// launch without `--wid` gives mpv, with a box of its own so a placement has somewhere to land.
unsafe fn a_window_of_the_players_own() -> HWND {
    unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP | WS_VISIBLE,
            NOWHERE,
            NOWHERE,
            1920,
            1080,
            None,
            None,
            None,
            None,
        )
    }
    .expect("a window of the player's own, in this process")
}

/// Where a hover's film's own size comes from, when the player is one that answers nothing: the
/// probe's measurement of the same film, held where the launch reads it (see `probed_as` in
/// `video_hold`).
fn probed_at(path: &Path, width: u32, height: u32) {
    video_geometry_cache().insert(
        VideoGeometryKey {
            path: path.to_path_buf(),
            version: file_version(path),
        },
        ProbedGeometry::Measured(VideoGeometry {
            width,
            height,
            frame_width: width,
            frame_height: height,
            crop: None,
            duration: None,
            subtitles: SubtitleStreams::default(),
            sidecar: None,
            derived: None,
            subtitle_codecs: Vec::new(),
            attachment_codecs: Vec::new(),
            subtitle_extraction_failed: false,
        }),
    );
}

/// A hover on the film road: its own host as the destination, its own source adopted into it, and
/// the player's one placement made. Answers the host and the source the caller destroys.
///
/// `route` is the player drawing it: **mpv answers its own `video-params` and FFmpeg's player answers
/// nothing at all**, so the film's own size comes from the probe instead (see
/// `pin_film_thumbnail::the_films_own_size`).
unsafe fn a_hover_drawn_through_a_thumbnail(
    film: &StandingFilm,
    path: &Path,
    route: VideoRoute,
) -> (HWND, HWND, u32) {
    let pid = unsafe { GetCurrentProcessId() };
    let source = unsafe { a_window_of_the_players_own() };
    // The stand begins from no picture at all: a registration left by another test is not this
    // film's, and a source window's handle is a handle the system reuses.
    pin_film_thumbnail::forget_the_thumbnail();
    let host = unsafe { hover_film_host::take_up_the_hover_film_host(film.box_) }
        .expect("the hover's own host");
    hover_film_host::note_hover_film_host_player(pid, path);
    match route {
        VideoRoute::Mpv => {
            mpv_ipc::remember_pipe(pid, "the-hover-sources-pipe", false);
            mpv_ipc::answer_a_read(
                pid,
                "video-params",
                Some(r#"{"w":1920,"h":1080,"dw":1920,"dh":1080}"#),
            );
        }
        _ => {
            mpv_ipc::forget_pipe(pid);
            probed_at(path, 1920, 1080);
        }
    }
    pin_film_thumbnail::clear_the_source_placements();
    pin_film_thumbnail::clear_the_registrations();
    // **The adoption is this test's, not the monitor's.** The pass that watches a player is a
    // thread of the whole process and would adopt a window of the process's own as the picture the
    // moment it saw a player recorded and no source; it is kept off this one by the pid it follows
    // being written after the picture is up (see `ensure_noactivate_monitor`).
    unsafe { pin_film_thumbnail::adopt_the_source(source, pid, host) };
    VIDEO_PID.store(pid, Ordering::SeqCst);
    (host, source, pid)
}

/// **A film pinned from a hover on the thumbnail road is the same player and the same source.**
/// The take-up begins nothing, asks the player for no box, and leaves the one registration where it
/// already is — in the window that was the hover's host and is the pin's (see `pin_film_thumbnail`
/// and `pin_lifecycle`). The pin's end takes the source and the registration down with it.
#[test]
fn a_hover_drawn_through_a_thumbnail_is_taken_up_without_a_player_or_a_placement() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let (hover, source, pid) =
        unsafe { a_hover_drawn_through_a_thumbnail(&film, &path, VideoRoute::Mpv) };

    assert_eq!(
        pin_film_thumbnail::the_source_placements(),
        1,
        "a source is placed once, at its own size"
    );
    assert_eq!(
        pin_film_thumbnail::the_registration_host(),
        Some(hover.0 as isize),
        "and the hover's host is where the Shell draws it"
    );
    assert!(
        pin_film_thumbnail::the_source_is_up(pid),
        "the hover's own source is the film's picture"
    );

    clear_restart_count();
    unsafe { show_pinned_window(film.pin_hwnd) };

    assert_eq!(
        restart_count(),
        0,
        "the take-up keeps the player the hover began rather than beginning one of its own"
    );
    assert_eq!(
        pin_film_host::film_host(),
        Some(hover),
        "the hover's host is the pin's host"
    );
    assert_eq!(
        pin_film_host::hosted_player_pid(),
        Some(pid),
        "and the pin holds the player it took up"
    );
    assert_eq!(
        pin_film_thumbnail::the_source_placements(),
        1,
        "nothing asks the source for another box across the promotion"
    );
    assert_eq!(
        pin_film_thumbnail::the_registration_host(),
        Some(hover.0 as isize),
        "exactly one registration, in the pin's host"
    );

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    assert!(
        !pin_film_thumbnail::the_source_is_up(pid),
        "a pin's end leaves no source behind it"
    );
    assert_eq!(
        pin_film_thumbnail::the_registration_host(),
        None,
        "and no registration either"
    );

    let _ = unsafe { DestroyWindow(source) };
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::clear_reads();
}

/// **A hover that is never pinned leaves nothing behind it**: the source it adopted and the
/// registration in its host go with the host, and no pin has to exist for either to be taken down
/// (see `hover_film_host` and `pin_film_thumbnail::forget_the_thumbnail_of`).
#[test]
fn a_hover_drawn_through_a_thumbnail_that_is_never_pinned_leaves_no_source() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let (_, source, pid) =
        unsafe { a_hover_drawn_through_a_thumbnail(&film, &path, VideoRoute::Mpv) };

    assert!(pin_film_thumbnail::the_source_is_up(pid));
    unsafe { hover_film_host::take_down_the_hover_film_host() };

    assert!(
        !pin_film_thumbnail::the_source_is_up(pid),
        "the hover ending takes the source it adopted with it"
    );
    assert_eq!(
        pin_film_thumbnail::the_registration_host(),
        None,
        "and the registration its host was drawn with"
    );

    let _ = unsafe { DestroyWindow(source) };
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::clear_reads();
}

/// **One source per film: placed once, never resized, never placed again.** A whole gesture — the
/// hover, its take-up into the pin, the press that begins a resize, both directions of its drag, the
/// release and the settle that follows it, and the teardown — leaves the source at the film's own
/// size, with exactly one placement and no second ask (see `pin_film_thumbnail`).
#[test]
fn a_film_is_placed_once_and_never_resized_across_a_whole_gesture() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    a_whole_gesture_places_a_source_once(StandingFilm::new(), VideoRoute::Mpv);
}

/// **The same gesture on the other engine**, which is the report this road was put under: dragging
/// the pin, resizing it, moving it and clicking the picture all paused an FFmpeg film, because its
/// window was fitted to the host on every step. Nothing on this road touches the source's geometry
/// after the one placement — `size_the_source` is the only call that can size it, `move_the_source`
/// is the only other that touches it and carries `SWP_NOSIZE`, and both are posted
/// (`SWP_ASYNCWINDOWPOS`) so neither can block a drag on the player's thread. What a test can hold
/// is the count and the box: one placement, no second ask, and the same size from the hover to the
/// teardown.
#[test]
fn an_ffmpeg_film_is_placed_once_and_never_resized_across_a_whole_gesture() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    a_whole_gesture_places_a_source_once(ffplay_standing_film(), VideoRoute::Ffplay);
}

/// The gesture both engines' tests above are: hover, take-up, the steps of a resize drag, the
/// release, the settle and the teardown, with the source's box and the placement count read at each
/// of them (see `a_film_is_placed_once_and_never_resized_across_a_whole_gesture`).
fn a_whole_gesture_places_a_source_once(film: StandingFilm, route: VideoRoute) {
    let path = PathBuf::from("a-film.mkv");
    let (host, source, pid) = unsafe { a_hover_drawn_through_a_thumbnail(&film, &path, route) };
    assert_eq!(
        pin_film_thumbnail::the_registrations(),
        1,
        "one source, one registration, whichever engine made it"
    );
    let placed = unsafe { window_rect(source) };
    let size = (placed.2 - placed.0, placed.3 - placed.1);
    assert_eq!(
        pin_film_thumbnail::the_source_placements(),
        1,
        "one placement, at adoption"
    );
    assert_eq!(size, (1920, 1080), "at the film's own size");
    let still = |what: &str| {
        let box_ = unsafe { window_rect(source) };
        assert_eq!(
            (box_.2 - box_.0, box_.3 - box_.1),
            size,
            "the source is not resized by {what}"
        );
        assert_eq!(
            pin_film_thumbnail::the_source_placements(),
            1,
            "nor placed again by {what}"
        );
        assert_eq!(
            pin_film_thumbnail::the_source_replacements(),
            0,
            "nor asked for a second placement by {what}"
        );
    };

    // The hover moves and is re-cropped; a box at an edge may slide the source, never size it.
    for box_ in [
        (100, 100, 740, 460),
        (900, 700, 1700, 1150),
        (0, 0, 640, 360),
    ] {
        unsafe { hover_film_host::place_the_hover_film_host(box_) };
        still("a hover moving");
    }

    // The press is a take-up: the host is promoted, nothing is begun and nothing is placed.
    clear_restart_count();
    unsafe { show_pinned_window(film.pin_hwnd) };
    assert_eq!(restart_count(), 0, "a take-up begins no player");
    assert_eq!(
        pin_film_host::film_host(),
        Some(host),
        "the hover's host is the pin's"
    );
    still("the take-up");

    // The drag, both ways: every step is a box for the host and nothing of the source.
    for box_ in [
        (0, 0, 1200, 675),
        (900, 500, 1900, 1060),
        (-100, 0, 900, 560),
        (200, 100, 700, 400),
    ] {
        unsafe { pin_film_host::place_the_pin_film_host(box_) };
        still("a step of a resize drag");
    }

    // The release and the settle: the settled box again, then the same box once more.
    for box_ in [(200, 100, 1000, 550), (200, 100, 1000, 550)] {
        unsafe { pin_film_host::place_the_pin_film_host(box_) };
        still("the release and the settle");
    }

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    assert_eq!(
        pin_film_thumbnail::the_source_placements(),
        1,
        "a teardown places nothing"
    );
    assert!(
        !pin_film_thumbnail::the_source_is_up(pid),
        "and leaves no source behind it"
    );

    let _ = unsafe { DestroyWindow(source) };
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::clear_reads();
}

/// **A window that replaced itself is a new source, and a new source is adopted once and takes the
/// picture only when it can draw it.** FFmpeg's player can make its window again — and every relaunch
/// of it is another process with another window — so the monitor's pass finds a pid whose picture it is
/// not drawing, adopts the new window and places it once. What it must not do is take the picture down
/// before that window can draw: the registration the host wears goes on drawing the window before it
/// until the new source's client is the one the placement asked for and its output is configured for a
/// frame, and then one swap is made — the new registration handed over before the old one is let go of
/// (see `pin_film_thumbnail::hand_the_picture_over`). A pass that finds the source standing asks for
/// nothing at all, which is what keeps one window from being placed twice.
#[test]
fn a_source_that_replaced_itself_is_adopted_once_and_takes_the_picture_when_it_is_ready() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let (host, first, pid) =
        unsafe { a_hover_drawn_through_a_thumbnail(&film, &path, VideoRoute::Mpv) };
    assert_eq!(
        (
            pin_film_thumbnail::the_source_placements(),
            pin_film_thumbnail::the_registrations()
        ),
        (1, 1),
        "the film's first window is placed once and registered once"
    );

    // The player's window goes and another takes its place, under the same pid: the pass that
    // watches the player adopts what it finds, and only what it finds.
    unsafe { DestroyWindow(first) }.expect("the first source goes");
    let second = unsafe { a_window_of_the_players_own() };
    pin_film_thumbnail::clear_the_source_placements();
    pin_film_thumbnail::clear_the_registrations();
    pin_film_thumbnail::clear_the_unregistrations();
    let the_monitors_pass = || {
        if !pin_film_thumbnail::the_source_is_up(pid) {
            unsafe { pin_film_thumbnail::adopt_the_source(second, pid, host) };
        } else {
            pin_film_thumbnail::place_the_source_when_the_player_answers(host);
            pin_film_thumbnail::hand_the_picture_over(host);
        }
    };
    pin_trace::clear_noted();
    // The pass that adopts the window the player made, and the pass after it — which is the one
    // that asks whether the new source can draw yet (see `hand_the_picture_over`).
    the_monitors_pass();
    the_monitors_pass();

    assert_eq!(
        pin_film_thumbnail::the_source_window(),
        Some(second),
        "the record follows the window the player made last"
    );
    assert_eq!(
        pin_film_thumbnail::the_source_placements(),
        1,
        "the new window takes its own one placement"
    );
    assert_eq!(
        pin_film_thumbnail::the_source_replacements(),
        0,
        "and is never asked for a second one"
    );
    // **The picture standing is still the window before it**, which is the whole point: the player
    // this window replaced has not had its frame taken off the screen to show one that is not there.
    assert_eq!(
        pin_film_thumbnail::the_registrations(),
        0,
        "nothing is registered for a source that cannot draw yet"
    );
    assert_eq!(
        pin_film_thumbnail::the_unregistrations(),
        0,
        "and the registration standing is not taken down for it either"
    );
    assert!(
        !pin_film_thumbnail::the_standing_source_is_the_current_one(),
        "the picture is the window before this one"
    );
    assert!(
        pin_film_thumbnail::a_source_is_waiting_to_take_the_picture(),
        "and the swap is what is waiting"
    );
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("thumbnail source not ready")),
        "and the run says so: a source that is not ready is a traced wait, not a blank pin"
    );
    assert!(
        !pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("thumbnail swapped")),
        "and nothing has changed hands"
    );

    // The player answers that its output is configured for a frame: the same pass the monitor makes
    // hands the picture over, once.
    mpv_ipc::answer_a_read(pid, "vo-configured", Some("true"));
    pin_trace::clear_noted();
    the_monitors_pass();

    assert!(
        pin_film_thumbnail::the_standing_source_is_the_current_one(),
        "the new source is the picture once it can draw"
    );
    assert!(
        !pin_film_thumbnail::a_source_is_waiting_to_take_the_picture(),
        "and nothing is waiting to take it any more"
    );
    assert_eq!(
        pin_film_thumbnail::the_registrations(),
        1,
        "exactly one registration: the swap"
    );
    assert_eq!(
        pin_film_thumbnail::the_unregistrations(),
        1,
        "and exactly one old registration let go of, after the new one was handed over"
    );
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("thumbnail swapped old")),
        "and the swap is one line: old window, new window, how long the wait was"
    );

    // The next pass finds the source standing and asks for nothing at all.
    pin_film_thumbnail::clear_the_source_placements();
    pin_film_thumbnail::clear_the_registrations();
    pin_film_thumbnail::clear_the_unregistrations();
    the_monitors_pass();
    assert_eq!(
        (
            pin_film_thumbnail::the_source_placements(),
            pin_film_thumbnail::the_registrations(),
            pin_film_thumbnail::the_unregistrations(),
        ),
        (0, 0, 0),
        "a source that is already the picture is not adopted, placed or registered again"
    );

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    assert!(
        !pin_film_thumbnail::the_source_is_up(pid),
        "and the pin's end leaves no source behind it"
    );

    let _ = unsafe { DestroyWindow(second) };
    unsafe { hover_film_host::take_down_the_hover_film_host() };
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::clear_reads();
    pin_trace::clear_noted();
}

/// **A copy landing under a pin swaps nothing.** FFmpeg's player can be told nothing, so the copy is
/// preloaded for the next relaunch rather than the picture changing hands for a subtitle: no player is
/// begun, no registration is taken down or made, and the film on screen is the one it was (see
/// `pin_playback::preload_the_subtitles_for_the_next_relaunch`).
#[test]
fn a_late_subtitle_swaps_nothing_and_rides_the_relaunch_the_user_asks_for() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let (host, _first, pid) =
        unsafe { a_hover_drawn_through_a_thumbnail(&film, &path, VideoRoute::Ffplay) };
    let _ = host;

    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    clear_restart_count();
    pin_film_thumbnail::clear_the_registrations();
    pin_film_thumbnail::clear_the_unregistrations();
    pin_trace::clear_noted();
    let hides = pin_film_host::host_hides_asked();

    assert!(
        reload_pinned_subtitles(&path),
        "a copy that lands while the pin is up is answered"
    );
    assert_eq!(
        restart_count(),
        0,
        "and no player is begun for it: a copy is not a reason to change the picture's hands"
    );
    assert_eq!(last_restart_reason(), None, "so no relaunch happened at all");
    assert!(
        pin_film_thumbnail::the_standing_source_is_the_current_one(),
        "the picture on screen is the source it was"
    );
    assert!(
        !pin_film_thumbnail::a_source_is_waiting_to_take_the_picture(),
        "and nothing is waiting to take it"
    );
    assert_eq!(
        (
            pin_film_thumbnail::the_registrations(),
            pin_film_thumbnail::the_unregistrations(),
        ),
        (0, 0),
        "nothing was registered and nothing was let go of"
    );
    assert_eq!(
        pin_film_host::host_hides_asked(),
        hides,
        "the host the picture is drawn in is never hidden for a subtitle"
    );
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("subtitle preloaded for the next relaunch")),
        "and the run says the copy is waiting for a relaunch: {:?}",
        pin_trace::noted_lines()
    );

    // The relaunch the user asks for is where it is activated, and one line says it is being carried.
    restart_pinned_player(&path, film.box_, 90.0, false, RestartReason::Seek);
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("subtitle carried into a relaunch")),
        "the copy rides the next relaunch: {:?}",
        pin_trace::noted_lines()
    );

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    unsafe { hover_film_host::take_down_the_hover_film_host() };
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    clear_restart_count();
    clear_pending_subtitles();
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::clear_reads();
    clear_pending_pinned_relaunch();
    pin_trace::clear_noted();
}

/// **A seek is one read of the numbers the frame it lands on is drawn from, and nothing of the picture
/// changes for it.** mpv is taken to the second while it runs, so there is no relaunch, no unregister
/// and no handover: the source the host is drawing is the one it was, and the crop is still the client
/// the one placement asked for — which is what says a frame drawn smaller than the pin's box on a seek
/// would be a number in this line rather than a thing to guess at (see
/// `pin_film_thumbnail::note_the_seek_picture`). FFmpeg's player is begun again, and what that costs is
/// one player and no registration: the picture standing is kept for the window that replaces it.
#[test]
fn a_seek_moves_the_second_and_not_the_picture_on_either_engine() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let (host, source, pid) =
        unsafe { a_hover_drawn_through_a_thumbnail(&film, &path, VideoRoute::Mpv) };
    let content = film.box_;
    let _ = source;
    let _ = host;

    // mpv: a line down the pipe, the bar rebased onto the second, and one read of what the engine
    // reports drawing at the frame it landed on.
    stand_mpv_in(Some(true));
    stand_ffplay_in(Some(false));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    clear_restart_count();
    pin_film_thumbnail::clear_the_unregistrations();
    mpv_ipc::clear_noted();
    pin_trace::clear_noted();
    let hides = pin_film_host::host_hides_asked();

    seek_pinned_playback(&path, content, 90.0);

    assert_eq!(
        restart_count(),
        0,
        "mpv is taken to the second while it runs"
    );
    assert_eq!(
        pin_film_host::host_hides_asked(),
        hides,
        "and nothing of the picture is hidden for a seek on that engine"
    );
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![format!(
            "{}
",
            r#"{"command":["seek",90.000,"absolute"],"async":true}"#
        )],
        "and the only command is the seek"
    );
    assert_eq!(
        pin_film_thumbnail::the_unregistrations(),
        0,
        "no registration is taken down for a seek: the source does not change"
    );
    assert!(
        pin_film_thumbnail::the_standing_source_is_the_current_one(),
        "the picture is the source it was"
    );
    assert_eq!(
        pin_film_thumbnail::the_cropped_client(),
        Some((1920, 1080)),
        "and the crop after the seek is still the client the one placement asked for"
    );
    let seek_lines: Vec<String> = pin_trace::noted_lines()
        .into_iter()
        .filter(|line| line.starts_with("seek picture"))
        .collect();
    assert_eq!(
        seek_lines.len(),
        1,
        "one read per seek, and never one per frame: {seek_lines:?}"
    );
    assert!(
        seek_lines[0].contains("cli 1920x1080")
            && seek_lines[0].contains("pic 1920x1080")
            && seek_lines[0].contains("asked 1920x1080")
            && !seek_lines[0].contains("!=asked"),
        "and the picture the engine reports is the box the source was placed at: {}",
        seek_lines[0]
    );

    // FFmpeg's player: one relaunch, and the registration the host wears is still the old source's
    // until the window that replaces it can draw (see `pin_film_thumbnail::hand_the_picture_over`).
    clear_restart_count();
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    VIDEO_PID.store(pid, Ordering::SeqCst);
    pin_film_thumbnail::clear_the_unregistrations();
    pin_trace::clear_noted();

    seek_pinned_playback(&path, content, 90.0);

    assert_eq!(
        restart_count(),
        1,
        "a seek is one player begun again on this engine"
    );
    assert_eq!(
        last_restart_reason(),
        Some(RestartReason::Seek),
        "and it says the seek is what asked for it"
    );
    assert_eq!(
        pin_film_thumbnail::the_unregistrations(),
        0,
        "and the picture standing is not taken down for the player that is still starting"
    );
    assert_eq!(
        pin_film_host::host_hides_asked(),
        hides,
        "and the host the picture is drawn in is never hidden for a seek on this engine"
    );
    assert!(
        !pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("film host hidden while a pin lives")),
        "with the run saying nothing was taken off the screen: {:?}",
        pin_trace::noted_lines()
    );
    assert!(
        pin_film_thumbnail::the_standing_source_is_the_current_one(),
        "the old source goes on drawing until the new one can"
    );

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    unsafe { hover_film_host::take_down_the_hover_film_host() };
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    clear_restart_count();
    pin_trace::clear_noted();
}

/// **A source that never reaches the client it was placed at is waited for and never blanks the pin.**
/// A short or stuck window — a player that has not taken the box it was asked for, or one that came up at
/// a size of its own — is a swap that never happens: the picture before it goes on drawing, the run says
/// `thumbnail source not ready` with both numbers, and nothing is registered, unregistered or begun for
/// it. The wait is bounded by the player-start wait, which hands the picture over anyway rather than
/// leaving the one before it to a player that is ending (see `hand_the_picture_over`).
#[test]
fn a_source_that_never_reaches_its_asked_client_does_not_blank_the_pin() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let (host, source, pid) = unsafe {
        a_hover_drawn_through_a_thumbnail(&film, &PathBuf::from("a-film.mkv"), VideoRoute::Mpv)
    };

    // The player's window comes back and is adopted in the place of the one that was drawing — and
    // then puts itself back at a box of its own, which is the client the placement did not ask for
    // and the way a player that will not take the box it was given behaves.
    unsafe { DestroyWindow(source) }.expect("the source goes");
    let short = unsafe { a_window_of_the_players_own() };
    mpv_ipc::answer_a_read(pid, "vo-configured", Some("true"));
    pin_film_thumbnail::clear_the_registrations();
    pin_film_thumbnail::clear_the_unregistrations();
    pin_trace::clear_noted();
    unsafe { pin_film_thumbnail::adopt_the_source(short, pid, host) };
    unsafe { MoveWindow(short, NOWHERE, NOWHERE, 960, 540, false) }.expect("a short window");
    pin_film_thumbnail::hand_the_picture_over(host);

    assert!(
        !pin_film_thumbnail::the_standing_source_is_the_current_one(),
        "a source at another client is not the picture"
    );
    assert_eq!(
        (
            pin_film_thumbnail::the_registrations(),
            pin_film_thumbnail::the_unregistrations(),
        ),
        (0, 0),
        "and nothing is registered or unregistered for it: the pin draws what it drew"
    );
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line == "thumbnail source not ready cli 960x540 asked 1920x1080"),
        "the wait is traced with both numbers, so a stuck window is a line and not a guess: {:?}",
        pin_trace::noted_lines()
    );

    // **And the short window does not become the picture by being asked again**: the placement is
    // the film's one, and the wait goes on rather than sizing the source a second time.
    pin_film_thumbnail::place_the_source_when_the_player_answers(host);
    pin_film_thumbnail::hand_the_picture_over(host);
    assert_eq!(
        pin_film_thumbnail::the_source_replacements(),
        0,
        "nothing is re-placed to make it fit"
    );
    assert!(
        !pin_film_thumbnail::the_standing_source_is_the_current_one(),
        "and it is still not the picture"
    );

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    let _ = unsafe { DestroyWindow(short) };
    unsafe { hover_film_host::take_down_the_hover_film_host() };
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::clear_reads();
    pin_trace::clear_noted();
}

/// **The player on screen outlives its replacement, for every reason a relaunch has.** A relaunch begins
/// another player behind the one that is drawing, adopts and places its source once, and only when the
/// swap has made that source the picture does the loop's retirement end the one before it — so the pin's
/// host is never handed the band while the registration it wears draws a window whose process is gone
/// (see `retire_replaced_player`, `pin_film_thumbnail::hand_the_picture_over`).
///
/// The order is read out of the run's own lines rather than out of a clock: the relaunch's line names the
/// player it replaced (`… from N to M`), the swap's line names both, and the retirement's line is written
/// after it and only once the picture stands — so nothing had to be stopped for the test to see the
/// order, and the one line that says a picture was ended while it stood is written by no reason at all.
#[test]
fn every_relaunch_ends_its_old_player_only_after_the_new_picture_stands() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let (host, _first, old_pid) =
        unsafe { a_hover_drawn_through_a_thumbnail(&film, &path, VideoRoute::Ffplay) };
    // The route a pinned film of this kind is relaunched on: FFmpeg's player, which answers nothing. The
    // relaunches below begin noted players, so nothing of the desktop's is spawned for them.
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));

    let reasons = [
        RestartReason::Seek,
        RestartReason::Volume,
        RestartReason::Subtitle,
        RestartReason::End,
        RestartReason::Engine,
    ];
    let mut standing = old_pid;
    let mut sources = Vec::new();
    for reason in reasons {
        let hides = pin_film_host::host_hides_asked();
        pin_trace::clear_noted();
        clear_restart_count();
        mpv_ipc::clear_noted();
        ffplay_window::clear_pause_posts();
        pin_film_thumbnail::clear_the_unregistrations();
        restart_pinned_player(&path, film.box_, 10.0, false, reason);

        let replacement = VIDEO_PID.load(Ordering::SeqCst);
        assert_ne!(replacement, 0, "a replacement player was begun");
        assert_ne!(replacement, standing, "and it is another player");
        assert_eq!(last_restart_reason(), Some(reason));

        // **The picture on screen is frozen before anything is begun over it, and stays frozen**: one
        // telling at most, because the key toggles and a second would set it going again (see
        // `freeze_the_outgoing_picture`).
        let lines = pin_trace::noted_lines();
        let at_freeze = lines
            .iter()
            .position(|line| line.starts_with(&format!("outgoing picture frozen pid {standing} ")))
            .unwrap_or_else(|| {
                panic!("the picture is frozen for a relaunch for {reason:?}: {lines:?}")
            });
        let at_restart = lines
            .iter()
            .position(|line| line.starts_with("player restarted for "))
            .unwrap_or_else(|| panic!("a restart line: {lines:?}"));
        assert!(
            at_freeze < at_restart,
            "the picture is frozen before its replacement is begun, for {reason:?}: {lines:?}"
        );
        assert!(
            ffplay_window::pause_posts() <= 1,
            "at most one key is posted to the player it replaces, for {reason:?}"
        );
        assert!(
            !mpv_ipc::noted_lines()
                .iter()
                .any(|line| line.contains(r#"["pause",false]"#)),
            "and nothing sets it going again, for {reason:?}: {:?}",
            mpv_ipc::noted_lines()
        );
        assert_eq!(
            pin_film_thumbnail::the_unregistrations(),
            0,
            "the registration the host wears is never taken down for a relaunch's wait, for \
             {reason:?}"
        );
        assert_eq!(
            pin_film_host::host_hides_asked(),
            hides,
            "and nothing hides the host for it, for {reason:?}"
        );

        assert!(
            pin_trace::noted_lines().iter().any(|line| line
                == &format!(
                    "player restarted for {} from {standing} to {replacement}",
                    reason.as_str()
                )),
            "the player on screen is still the one the relaunch replaces, so nothing ended it: {:?}",
            pin_trace::noted_lines()
        );
        assert!(
            pin_film_thumbnail::the_source_is_up(standing),
            "and its source is still the picture after the relaunch for {reason:?}"
        );
        assert!(
            !pin_trace::noted_lines()
                .iter()
                .any(|line| line.starts_with("player ended while standing")),
            "a relaunch for {reason:?} ends nothing before its replacement exists: {:?}",
            pin_trace::noted_lines()
        );

        // The replacement's own window arrives and draws: the swap is one step, and it does not end the
        // player whose picture it is replacing.
        let second = unsafe { a_window_of_the_players_own() };
        unsafe { pin_film_thumbnail::adopt_the_source(second, replacement, host) };
        pin_film_thumbnail::hand_the_picture_over(host);
        assert!(
            pin_film_thumbnail::the_standing_source_is_the_current_one(),
            "the new source is the picture once it can draw, for {reason:?}"
        );
        let swap_line = pin_trace::noted_lines()
            .into_iter()
            .find(|line| line.starts_with("thumbnail swapped old"))
            .unwrap_or_else(|| panic!("a swap is one line: {:?}", pin_trace::noted_lines()));
        assert!(
            swap_line.contains(&format!("pid {standing}"))
                && swap_line.contains(&format!("pid {replacement}")),
            "the swap names both players, so a handover onto a dead window cannot pass: {swap_line}"
        );

        assert!(
            !pin_trace::noted_lines()
                .iter()
                .any(|line| line == &format!("player retired after the swap pid {standing}")),
            "the swap itself ends nothing: the old player is still the picture's for {reason:?}"
        );

        // The loop's retirement is what ends the player whose picture has been replaced.
        settle_video_retirement();
        let lines = pin_trace::noted_lines();
        let at_swap = lines
            .iter()
            .position(|line| line.starts_with("thumbnail swapped old"))
            .expect("a swap line");
        let at_retire = lines
            .iter()
            .position(|line| line == &format!("player retired after the swap pid {standing}"))
            .unwrap_or_else(|| panic!("the retirement is one line: {lines:?}"));
        assert!(
            at_swap < at_retire,
            "the picture changed hands before the player went, for {reason:?}: {lines:?}"
        );
        assert!(
            !lines
                .iter()
                .any(|line| line.starts_with("player ended while standing")),
            "no path ends the source the host is drawing, for {reason:?}: {lines:?}"
        );

        sources.push(second);
        standing = replacement;
        mpv_ipc::clear_noted();
    }

    for source in sources {
        let _ = unsafe { DestroyWindow(source) };
    }
    unsafe { pin_film_host::take_down_the_pin_film_host() };
    unsafe { hover_film_host::take_down_the_hover_film_host() };
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    clear_restart_count();
    clear_pending_pinned_relaunch();
    pin_trace::clear_noted();
}

/// **A record that names the picture is not a relaunch in flight, and reaping it must not kill what it
/// names.** The in-flight record is what a newer gesture ends before it can publish; a record left naming
/// the source the host is drawing is the opposite of that, and it is forgotten rather than killed (see
/// `reap_superseded_relaunch`).
#[test]
fn a_relaunch_record_that_names_the_picture_is_forgotten_rather_than_killed() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let (_host, source, pid) =
        unsafe { a_hover_drawn_through_a_thumbnail(&film, &path, VideoRoute::Ffplay) };
    pin_trace::clear_noted();

    assert!(
        pin_film_thumbnail::the_standing_source_is_the_current_one(),
        "the source is the picture before anything is reaped"
    );
    note_pinned_relaunch(pid);
    assert_eq!(
        reap_superseded_relaunch(),
        None,
        "a record naming the picture reaps nothing"
    );
    assert!(
        pin_film_thumbnail::the_standing_source_is_the_current_one(),
        "and the picture it named is still on screen"
    );
    assert_eq!(
        pending_pinned_relaunch(),
        None,
        "the stale record is forgotten rather than carried into the next relaunch"
    );
    assert!(
        !pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("player ended while standing")),
        "nothing was ended: {:?}",
        pin_trace::noted_lines()
    );

    let _ = unsafe { DestroyWindow(source) };
    unsafe { pin_film_host::take_down_the_pin_film_host() };
    unsafe { hover_film_host::take_down_the_hover_film_host() };
    clear_pending_pinned_relaunch();
    pin_trace::clear_noted();
}

/// **A window the right size is not a picture, on the engine that cannot say so.** A source placed at the
/// client the one placement asked for is not handed the picture while its own pixels are dark — the window
/// is the box, and nothing has been drawn into it — and it takes the picture on the pass it draws, with the
/// wait written down once (see `pin_film_thumbnail::the_source_drew`).
#[test]
fn a_source_at_the_asked_client_takes_the_picture_only_once_it_has_drawn() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let (host, first, pid) =
        unsafe { a_hover_drawn_through_a_thumbnail(&film, &path, VideoRoute::Ffplay) };

    // The player's window goes and another takes its place, placed once at the same film's size — and not
    // one pixel of a picture in it yet.
    unsafe { DestroyWindow(first) }.expect("the first source goes");
    let second = unsafe { a_window_of_the_players_own() };
    pin_film_thumbnail::clear_the_registrations();
    pin_film_thumbnail::clear_the_unregistrations();
    pin_trace::clear_noted();
    pin_film_thumbnail::stand_the_drawn_probe_in(Some(false));
    unsafe { pin_film_thumbnail::adopt_the_source(second, pid, host) };
    pin_film_thumbnail::hand_the_picture_over(host);

    assert!(
        !pin_film_thumbnail::the_standing_source_is_the_current_one(),
        "a window at the asked client with nothing drawn in it is not the picture"
    );
    assert_eq!(
        pin_film_thumbnail::the_unregistrations(),
        0,
        "and the registration standing is not taken down for it"
    );
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line == "thumbnail source not ready cli 1920x1080 asked 1920x1080"),
        "the wait is traced with both numbers, so a dark window is a line: {:?}",
        pin_trace::noted_lines()
    );

    // The pass on which it draws is the pass the picture changes hands on, and the wait is one line.
    pin_film_thumbnail::stand_the_drawn_probe_in(Some(true));
    pin_trace::clear_noted();
    pin_film_thumbnail::hand_the_picture_over(host);
    assert!(
        pin_film_thumbnail::the_standing_source_is_the_current_one(),
        "the new source is the picture once it has drawn"
    );
    assert_eq!(
        pin_film_thumbnail::the_unregistrations(),
        1,
        "and exactly one old registration is let go of, after the new one was handed over"
    );
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("thumbnail source drew after")),
        "and the run says how long the picture before it stood: {:?}",
        pin_trace::noted_lines()
    );

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    let _ = unsafe { DestroyWindow(second) };
    unsafe { hover_film_host::take_down_the_hover_film_host() };
    pin_film_thumbnail::stand_the_drawn_probe_in(Some(true));
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::clear_reads();
    pin_trace::clear_noted();
}

/// **A source that never draws is waited for and then handed the picture anyway, out loud**: the start
/// bound is the one answer that never leaves the old picture to a player that is ending, and the line that
/// says so is the one a black opening frame is read from (see `pin_film_thumbnail::hand_the_picture_over`).
#[test]
fn a_source_that_never_draws_is_swapped_at_the_bound_and_says_so() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let (host, first, pid) =
        unsafe { a_hover_drawn_through_a_thumbnail(&film, &path, VideoRoute::Ffplay) };

    unsafe { DestroyWindow(first) }.expect("the first source goes");
    let second = unsafe { a_window_of_the_players_own() };
    pin_film_thumbnail::clear_the_registrations();
    pin_film_thumbnail::clear_the_unregistrations();
    pin_trace::clear_noted();
    pin_film_thumbnail::stand_the_drawn_probe_in(Some(false));
    unsafe { pin_film_thumbnail::adopt_the_source(second, pid, host) };
    pin_film_thumbnail::stand_the_pending_source_since(Some(
        Instant::now() - Duration::from_secs(VIDEO_START_WAIT_SECS + 1),
    ));
    pin_trace::clear_noted();
    pin_film_thumbnail::hand_the_picture_over(host);

    assert!(
        pin_film_thumbnail::the_standing_source_is_the_current_one(),
        "at the bound the picture changes hands rather than being left on a player that is ending"
    );
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("thumbnail source never drew after")
                && line.ends_with(", swapping anyway")),
        "and the line says the source never drew: {:?}",
        pin_trace::noted_lines()
    );

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    let _ = unsafe { DestroyWindow(second) };
    unsafe { hover_film_host::take_down_the_hover_film_host() };
    pin_film_thumbnail::stand_the_drawn_probe_in(Some(true));
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::clear_reads();
    pin_trace::clear_noted();
}

/// **The sample itself reads a lit client and a dark one**, which is the whole premise of the drawn
/// readiness for a player with no pipe: a few points through the source's own DC, and one lit point is
/// enough (see `pin_film_thumbnail::the_source_has_lit_pixels`).
#[test]
fn the_drawn_sample_reads_a_lit_client_and_a_dark_one() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let window = unsafe { a_window_of_the_players_own() };
    // **A window off the desktop has no visible region to be written through**, so this one stands where
    // it can be drawn in and is shown the way a player shows its own — small, and for the length of one
    // test (see `an_erase_is_clipped_out_of_a_childs_box_where_the_window_asks_it_to_be`).
    unsafe {
        let _ = MoveWindow(window, 100, 100, 320, 240, true);
        let _ = ShowWindow(window, SW_SHOWNOACTIVATE);
    }

    // A client painted black is a window the right size with nothing drawn in it.
    unsafe { fill_the_client(window, COLORREF(0)) };
    assert!(
        !unsafe { pin_film_thumbnail::the_source_has_lit_pixels(window, (320, 240)) },
        "a black client is not a picture"
    );

    unsafe { fill_the_client(window, COLORREF(0x00FF_FFFF)) };
    assert!(
        unsafe { pin_film_thumbnail::the_source_has_lit_pixels(window, (320, 240)) },
        "a lit client is one sample away from being drawn"
    );

    let _ = unsafe { DestroyWindow(window) };
}

/// Paint a window's whole client one colour, through the window's own DC: what a player would have had
/// drawn into it, made a constant the drawn sample is asked about (see `the_source_has_lit_pixels`).
unsafe fn fill_the_client(window: HWND, colour: COLORREF) {
    use windows::Win32::Graphics::Gdi::{
        CreateSolidBrush, DeleteObject, FillRect, GdiFlush, GetDC, ReleaseDC,
    };

    unsafe {
        let dc = GetDC(window);
        let brush = CreateSolidBrush(colour);
        let mut rect = RECT::default();
        let _ = GetClientRect(window, &mut rect);
        FillRect(dc, &rect, brush);
        let _ = DeleteObject(brush);
        // The surface is handed over before it is read back through another context: GDI batches, and
        // what reads it next is a different DC.
        let _ = GdiFlush();
        ReleaseDC(window, dc);
    }
}

/// **A seek on FFmpeg's player leaves the picture where it is until the new one can draw it**: the ten
/// lines a seek's band used to show came from a registration drawing the window of the player the seek had
/// just killed, so the crop after the seek is asserted against the client the one placement asked for and
/// the registration the host wears is asserted to be the same source it was (see
/// `pin_film_thumbnail::note_the_seek_picture`).
#[test]
fn a_seek_on_ffplay_ends_no_standing_source_and_crops_by_the_placed_client() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let (_host, source, pid) =
        unsafe { a_hover_drawn_through_a_thumbnail(&film, &path, VideoRoute::Ffplay) };
    let content = film.box_;

    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    VIDEO_PID.store(pid, Ordering::SeqCst);
    clear_restart_count();
    pin_film_thumbnail::clear_the_unregistrations();
    pin_trace::clear_noted();

    seek_pinned_playback(&path, content, 90.0);

    assert_eq!(restart_count(), 1, "a seek on this engine is one relaunch");
    assert_eq!(last_restart_reason(), Some(RestartReason::Seek));
    assert_eq!(
        pin_film_thumbnail::the_unregistrations(),
        0,
        "the registration the host wears is not taken down for the player that is still starting"
    );
    assert!(
        pin_film_thumbnail::the_standing_source_is_the_current_one(),
        "the source the seek left behind is still the picture"
    );
    assert!(
        pin_film_thumbnail::the_source_is_up(pid),
        "its window is still there to be drawn from: no picture gap can come from a dead registration"
    );
    assert_eq!(
        pin_film_thumbnail::the_cropped_client(),
        Some((1920, 1080)),
        "and the crop after the seek is the client the one placement asked for"
    );
    assert!(
        !pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("player ended while standing")),
        "nothing was ended under the picture: {:?}",
        pin_trace::noted_lines()
    );

    let _ = unsafe { DestroyWindow(source) };
    unsafe { pin_film_host::take_down_the_pin_film_host() };
    unsafe { hover_film_host::take_down_the_hover_film_host() };
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    clear_restart_count();
    clear_pending_pinned_relaunch();
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::clear_reads();
    pin_trace::clear_noted();
}
/// **The swap is the end of the relaunch it completes**: the in-flight record is forgotten by the step
/// that hands the picture over, so a later relaunch cannot kill the player that is now the picture as a
/// superseded one (see `pin_playback::settle_the_published_relaunch`).
#[test]
fn the_swap_forgets_the_relaunch_that_put_the_picture_up() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let (host, first, pid) =
        unsafe { a_hover_drawn_through_a_thumbnail(&film, &path, VideoRoute::Ffplay) };

    unsafe { DestroyWindow(first) }.expect("the first source goes");
    let second = unsafe { a_window_of_the_players_own() };
    pin_film_thumbnail::clear_the_registrations();
    pin_film_thumbnail::clear_the_unregistrations();
    pin_trace::clear_noted();
    unsafe { pin_film_thumbnail::adopt_the_source(second, pid, host) };
    assert!(
        pin_film_thumbnail::a_source_is_waiting_to_take_the_picture(),
        "the replacement is adopted and waiting"
    );
    note_pinned_relaunch(pid);
    assert_eq!(
        pending_pinned_relaunch().map(|pending| pending.pid),
        Some(pid),
        "and the relaunch that began it is in flight"
    );

    pin_film_thumbnail::hand_the_picture_over(host);

    assert!(
        pin_film_thumbnail::the_standing_source_is_the_current_one(),
        "the picture changed hands"
    );
    assert_eq!(
        pending_pinned_relaunch(),
        None,
        "and the relaunch is over with the swap: nothing is left for a later reap to kill"
    );

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    let _ = unsafe { DestroyWindow(second) };
    unsafe { hover_film_host::take_down_the_hover_film_host() };
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::clear_reads();
    pin_trace::clear_noted();
}

/// **The picture on screen is frozen down the pipe where the player answers on one, and nothing else
/// is asked of it.** mpv takes `set_property pause true` in microseconds and holds the frame it was
/// showing, which is the whole of what the host goes on drawing while its replacement opens the file
/// — no capture, no cover, no hide (see `freeze_the_outgoing_picture`).
#[test]
fn a_relaunch_freezes_the_picture_on_screen_down_the_pipe_the_player_answers_on() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let (_host, _source, pid) =
        unsafe { a_hover_drawn_through_a_thumbnail(&film, &path, VideoRoute::Mpv) };
    // The replacement is begun as the route a test can stand in for; what is being read here is how
    // the player *on screen* is told, which the pipe it answers on decides.
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    let hides = pin_film_host::host_hides_asked();
    clear_restart_count();
    ffplay_window::clear_pause_posts();
    mpv_ipc::clear_noted();
    pin_trace::clear_noted();

    restart_pinned_player(&path, film.box_, 10.0, false, RestartReason::Seek);

    let lines = pin_trace::noted_lines();
    assert!(
        lines
            .iter()
            .any(|line| line == &format!("outgoing picture frozen pid {pid} by the pipe")),
        "the player on screen is told to hold, and the run says how: {lines:?}"
    );
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![format!(
            "{}\n",
            r#"{"command":["set_property","pause",true],"async":true}"#
        )],
        "one command, and it is the hold: a player that answers on a pipe is not posted a key"
    );
    assert_eq!(
        ffplay_window::pause_posts(),
        0,
        "so nothing of the freeze is a key on this road"
    );
    assert_eq!(
        pin_film_host::host_hides_asked(),
        hides,
        "and nothing hides the host the frozen frame is drawn in"
    );
    assert!(
        pin_film_thumbnail::the_source_is_up(pid),
        "the registration standing goes on drawing the frame that was frozen"
    );

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    unsafe { hover_film_host::take_down_the_hover_film_host() };
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::clear_noted();
    mpv_ipc::clear_reads();
    clear_restart_count();
    pin_trace::clear_noted();
}

/// **A posted key reaches a window wearing the source's own extended style.** The freeze of a film
/// FFmpeg's player is drawing is the pause key posted to the player's own window (see
/// `freeze_the_outgoing_picture`), and that window is the source — which wears `WS_EX_TRANSPARENT`
/// and `WS_EX_NOACTIVATE` (see `player_ex_style_for`). Neither bit is about queued input: a key
/// posted to a window is delivered to that window's queue whatever it is styled, and this is that
/// measured rather than assumed. What a test cannot reach is the player's own reading of the key.
#[test]
fn a_posted_pause_key_reaches_a_window_wearing_the_sources_own_style() {
    let window = unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            0,
            0,
            64,
            32,
            None,
            None,
            None,
            None,
        )
    }
    .expect("a window to post to");
    unsafe {
        let ex = GetWindowLongPtrW(window, GWL_EXSTYLE);
        SetWindowLongPtrW(
            window,
            GWL_EXSTYLE,
            ex | WS_EX_TRANSPARENT.0 as isize | WS_EX_NOACTIVATE.0 as isize,
        );
    }

    let (vk, scan) = ffplay_window::VIDEO_PAUSE_KEY;
    let lparam = ffplay_window::ffplay_key_lparam(scan);
    unsafe {
        let _ = PostMessageW(window, WM_KEYDOWN, WPARAM(vk as usize), lparam);
        let _ = PostMessageW(window, WM_KEYUP, WPARAM(vk as usize), lparam);
    }

    let mut seen = Vec::new();
    let mut message = MSG::default();
    // SAFETY: the queue belongs to this thread and the two messages above are the only ones in this
    // range, so the loop ends on them.
    unsafe {
        while PeekMessageW(&mut message, window, WM_KEYDOWN, WM_KEYUP, PM_REMOVE).as_bool() {
            seen.push((message.message, message.lParam.0));
        }
    }
    assert_eq!(
        seen,
        vec![(WM_KEYDOWN, lparam.0), (WM_KEYUP, lparam.0)],
        "the press and the release both arrive, in order, at the window the source style was put on"
    );
    let _ = unsafe { DestroyWindow(window) };
}

/// **A film the user is holding is already a still frame, and the pause key toggles**: a relaunch of a
/// held film posts nothing to it, because a key there would set it going again — which is the one
/// thing the freeze must not do (see `the_player_on_screen_is_held`).
#[test]
fn a_relaunch_of_a_held_film_posts_no_pause_key_to_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let (_host, _source, pid) =
        unsafe { a_hover_drawn_through_a_thumbnail(&film, &path, VideoRoute::Ffplay) };
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    with_pin(|pin| pin.transport.held(30.0));
    clear_restart_count();
    ffplay_window::clear_pause_posts();
    pin_trace::clear_noted();

    restart_pinned_player(&path, film.box_, 30.0, true, RestartReason::Seek);

    assert_eq!(
        ffplay_window::pause_posts(),
        0,
        "no key is posted to a player that is already holding its film"
    );
    assert!(
        !pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("outgoing picture frozen")),
        "and nothing claims a freeze that was not needed: {:?}",
        pin_trace::noted_lines()
    );

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    unsafe { hover_film_host::take_down_the_hover_film_host() };
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::clear_reads();
    clear_restart_count();
    pin_trace::clear_noted();
}

/// **The picture that is standing is the one the tracker and the z-order follow, not the replacement
/// that has made no window yet.** A relaunch records its new player and begins it; for the whole of the
/// wait the player before it is frozen and is what the host is drawing — so the host is drawing, the
/// fit and topmost re-assert keep working, and a seek is not the instant the pin is declared empty (see
/// `pin_film_host::the_host_is_drawing` and `pin_film_thumbnail::the_standing_source_process`).
#[test]
fn the_standing_source_is_what_the_tracker_and_the_topmost_road_follow() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    unsafe { show_pinned_window(film.pin_hwnd) };
    let host = pin_film_host::film_host().expect("the film's host");
    let pid = unsafe { GetCurrentProcessId() };
    film.drawn_in_the_host(pid);

    // The pin's own player, adopted and registered: the picture the host has.
    let first = unsafe { a_source_of_the_host(host) };
    pin_film_thumbnail::the_band_was_placed(host);
    pin_trace::clear_noted();
    assert!(
        pin_film_host::the_host_is_drawing(),
        "a host with a registration standing has a picture in it"
    );
    assert!(
        pin_trace::noted_lines().iter().any(|line| line
            == &format!("hosted drawing pid {pid} host {host:?} drawing 1")),
        "and the tracker names the player whose source it is: {:?}",
        pin_trace::noted_lines()
    );

    // A seek's relaunch: another player is recorded and begun, and it has made no window at all.
    pin_trace::clear_noted();
    pin_film_thumbnail::clear_the_unregistrations();
    let hides = pin_film_host::host_hides_asked();
    restart_pinned_player(&path, film.box_, 90.0, false, RestartReason::Seek);
    let replacement = VIDEO_PID.load(Ordering::SeqCst);
    assert_ne!(replacement, pid, "a replacement player was begun");

    assert!(
        pin_film_host::the_host_is_drawing(),
        "the frozen picture the seek left behind is what the host is drawing"
    );
    assert!(
        !pin_film_host::the_hosted_player_is_drawing(replacement),
        "and the replacement is not the picture: it has made no window"
    );
    let _ = ensure_video_window_topmost(film.box_.0, film.box_.1, 320, 240);
    assert!(
        !pin_trace::noted_lines()
            .iter()
            .any(|line| line.starts_with("topmost skipped: pin host not drawing yet")),
        "so the z-order re-assert is not skipped for a seek: {:?}",
        pin_trace::noted_lines()
    );
    assert!(
        pin_trace::noted_lines()
            .iter()
            .any(|line| line == "topmost shown: pin host drawing"),
        "and the host the frozen frame is drawn in is shown: {:?}",
        pin_trace::noted_lines()
    );
    assert_eq!(
        pin_film_host::host_hides_asked(),
        hides,
        "with nothing hidden behind the freeze"
    );
    assert_eq!(
        pin_film_thumbnail::the_unregistrations(),
        0,
        "and the registration standing never taken down for the wait"
    );

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    let _ = unsafe { DestroyWindow(first) };
    unsafe { hover_film_host::take_down_the_hover_film_host() };
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    clear_restart_count();
    clear_pending_subtitles();
    clear_pending_pinned_relaunch();
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::forget_pipe(replacement);
    mpv_ipc::clear_reads();
    pin_trace::clear_noted();
}

/// **`pic none` is nothing read, not a mismatch, and a picture that really disagrees stays loud.** A
/// seek whose frame drawn by FFmpeg's player is written as one line, and the line a run is convicted by
/// — `pic!=asked` — is only written where a picture was actually read (see `note_the_seek_picture`).
#[test]
fn a_seeks_picture_line_calls_nothing_a_mismatch_that_was_never_read() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let film = StandingFilm::new();
    let path = PathBuf::from("a-film.mkv");
    let (host, _first, pid) =
        unsafe { a_hover_drawn_through_a_thumbnail(&film, &path, VideoRoute::Ffplay) };
    stand_mpv_in(Some(false));
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    clear_restart_count();
    pin_film_thumbnail::clear_the_registrations();
    pin_film_thumbnail::clear_the_unregistrations();
    pin_trace::clear_noted();

    // A seek's relaunch, and the swap onto the window that drew: FFmpeg's player answers nothing, so
    // the picture of the frame cannot be read — and that is not the loud line.
    restart_pinned_player(&path, film.box_, 90.0, false, RestartReason::Seek);
    let replacement = VIDEO_PID.load(Ordering::SeqCst);
    let second = unsafe { a_window_of_the_players_own() };
    unsafe { pin_film_thumbnail::adopt_the_source(second, replacement, host) };
    pin_film_thumbnail::hand_the_picture_over(host);

    let seek_lines: Vec<String> = pin_trace::noted_lines()
        .into_iter()
        .filter(|line| line.starts_with("seek picture"))
        .collect();
    assert_eq!(
        seek_lines.len(),
        1,
        "one read per seek, and never one per frame: {seek_lines:?}"
    );
    assert!(
        seek_lines[0].contains("cli 1920x1080")
            && seek_lines[0].contains("pic none")
            && seek_lines[0].contains("asked 1920x1080"),
        "the box is the placement's and the picture is nothing read: {}",
        seek_lines[0]
    );
    assert!(
        !seek_lines[0].contains("!=asked"),
        "a player that reports no picture is not a film fitted inside its own window: {}",
        seek_lines[0]
    );

    // A player that does answer, and answers with another size, is the loud line — asked of the player
    // whose window is the one standing, which after the swap is the replacement.
    pin_trace::clear_noted();
    mpv_ipc::remember_pipe(replacement, "the-standing-pipe", false);
    mpv_ipc::answer_a_read(
        replacement,
        "video-params",
        Some(r#"{"w":1280,"h":720,"dw":1280,"dh":720}"#),
    );
    pin_film_thumbnail::note_the_seek_picture(replacement, 90.0);
    let loud = pin_trace::noted_lines()
        .into_iter()
        .find(|line| line.starts_with("seek picture"))
        .unwrap_or_else(|| panic!("a line for the read"));
    assert!(
        loud.contains("pic 1280x720") && loud.contains("pic!=asked"),
        "a picture that really disagrees with the box stays loud: {loud}"
    );

    unsafe { pin_film_host::take_down_the_pin_film_host() };
    let _ = unsafe { DestroyWindow(second) };
    unsafe { hover_film_host::take_down_the_hover_film_host() };
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    clear_restart_count();
    clear_pending_subtitles();
    clear_pending_pinned_relaunch();
    mpv_ipc::forget_pipe(pid);
    mpv_ipc::clear_reads();
    pin_trace::clear_noted();
}

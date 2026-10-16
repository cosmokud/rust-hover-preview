//! The film's own window standing in the bubble a collapsed pin leaves: seating it in the circle,
//! taking it out again, and answering whether it is there.
//!
//! **A film mpv plays is drawn by another process, and this is the one way a bubble can show it
//! moving for nothing.** What mpv has up is a live picture at the film's own rate, already on the
//! card and already in step with the sound it is playing, so a collapse puts *that window* inside
//! the circle rather than reading the film a second time: it is carried into the bubble's square,
//! given a circular region so what shows through the ring is round rather than square, and told to
//! pan and scan so the film fills the square instead of being letterboxed inside it. The bubble
//! paints the ring, the bars and the mark around it and leaves the middle alone (see
//! `BubbleArt::Hole`).
//!
//! **The other two engines are not here.** FFmpeg's player cannot be told to fill a box — its
//! window letterboxes, so a bubble around it would show a stripe across the middle of the circle —
//! and the media engine draws into this app's own window, whose frame the bubble already has. Both
//! keep what every bubble had before a seat: a still.
//!
//! **A seat is a claim this app checks rather than believes**, because the window belongs to a
//! process that can die at any moment: a player that is no longer the one this app holds, or a
//! window that has gone, is a seat forgotten here rather than a hole painted around the desktop.

use super::*;
use windows::Win32::Graphics::Gdi::{CreateEllipticRgn, SetWindowRgn};

/// The film's window standing in the bubble, as this app seated it.
///
/// The two are one fact — a handle belongs to a process, and a window that cannot be traced back to
/// the player this app holds is a window of somebody else's — so they are kept together and verified
/// together (see `SeatedFilm`'s reader, `seated_film`). Where the window *stands* is deliberately not
/// kept here: the bubble's own window is the truth about that, and it is read where it is needed (see
/// `pin_bubble::bubble_box_on_screen`).
#[derive(Clone, Copy)]
struct SeatedFilm {
    hwnd: isize,
    pid: u32,
}

/// The film's window the bubble is standing around, or nothing while none is (see `SeatedFilm`).
static SEATED: Lazy<Mutex<Option<SeatedFilm>>> = Lazy::new(|| Mutex::new(None));

/// Whether this file's window is one a bubble can be built around.
///
/// Two questions, and both of them are about the *window* rather than the file. The film has to be
/// one somebody else is drawing, because a film this app draws is in this app's own window and its
/// bubble has the frame already; and that somebody has to be mpv, which is the one player here
/// that takes the shape of the window it is given, fills it, and can be told to (see the module
/// documentation).
pub(super) fn a_film_can_stand_in_the_bubble(kind: Option<MediaType>, route: VideoRoute) -> bool {
    matches!(kind, Some(MediaType::Video)) && route == VideoRoute::Mpv
}

/// Seat the film the pin is playing inside `box_`, answering whether one was seated.
///
/// Everything is asked of the machine here and nothing is assumed: the kind and the route decide
/// whether a seat is possible at all, and the player's own process and window decide whether there
/// is one to seat — a film whose player has gone is a film whose bubble keeps its still.
///
/// **The seat is written down before the window is touched**, so a window this app has moved is a
/// window it holds, and the record is what every reader afterwards checks the window against (see
/// `SeatedFilm`). The bubble is put back over the film as well: the picture is read *through* the
/// bubble, and every press, drag and menu is the bubble's own window's.
pub(super) fn seat_film_window(box_: ScreenRegion) -> bool {
    let Some((path, _)) = pinned_media_owner() else {
        return false;
    };
    if !a_film_can_stand_in_the_bubble(current_media_type(), video_route(&path)) {
        return false;
    }

    let pid = VIDEO_PID.load(Ordering::SeqCst);
    let hwnd = VIDEO_HWND.load(Ordering::SeqCst);
    let window = HWND(hwnd as *mut _);
    if pid == 0 || hwnd == 0 || !unsafe { IsWindow(window).as_bool() } {
        return false;
    }

    let Ok(mut seated) = SEATED.lock() else {
        return false;
    };
    *seated = Some(SeatedFilm { hwnd, pid });
    drop(seated);

    unsafe {
        put_the_film_window_in(window, box_);
    }

    // mpv fills the square rather than letterboxing inside it, which is the crop every other
    // bubble's picture is cut to as well (see `grab_video_frame`'s scale-and-crop). The answer is
    // not read: a film whose player cannot be told is a film that is seated slightly wrong rather
    // than a film that is not seated, and the picture is still the film's own.
    let _ = mpv_ipc::set_panscan(pid, true);

    raise_the_bubble();
    repaint_the_bubble();
    true
}

/// Take the film's window out of the bubble, and off the screen with it.
///
/// **The window goes with the seat rather than staying where the bubble left it**, because every
/// road that unseats is a road where the bubble is going away: the pin put back up (which places the
/// window in its band in the same call), the bubble turned into the taskbar stand-in, the pin ended,
/// and a hand that has taken hold of the bubble to carry it somewhere. A window left standing in a
/// square nobody is drawing anymore is the chrome-less film beside a bubble the app already refuses
/// in one place — and, at the collapse's own size, a film in a 44-pixel box.
///
/// What is this side's own is what the seat put on it — the region and the pan — so both come off
/// here, and the record goes with them; *where* the window goes next is still not this function's
/// business, and hiding it is not a place: it is the absence of one.
pub(super) fn unseat_film_window() {
    let Some(seated) = SEATED.lock().ok().and_then(|mut seat| seat.take()) else {
        return;
    };

    unsafe {
        let window = HWND(seated.hwnd as *mut _);
        if IsWindow(window).as_bool() {
            // **The region is cleared before the window is hidden**, because a square window with a
            // round region is a picture with its corners cut off everywhere it is put afterwards:
            // the pin's band is the first place this window goes next, and it is not a bubble (see
            // `put_the_film_window_in`).
            let _ = SetWindowRgn(window, None, true);
            let _ = ShowWindow(window, SW_HIDE);
        }
    }

    // And mpv stops filling its window for a box it no longer stands in, so a film put back in the
    // pin's band is the film that was there rather than one cropped to a bubble's square.
    let _ = mpv_ipc::set_panscan(seated.pid, false);

    repaint_the_bubble();
}

/// Keep the film the pin is showing in the bubble, answering whether the bubble has one.
///
/// This is what a tick owes a seat, and the road that puts one back: a film whose window is standing
/// in the bubble is put back in the box the circle is standing in now, and a bubble with no film in
/// it at all is given one where there is one to give — a file a walk landed on behind the bubble
/// starts its player with the pin off the screen, and nothing else would put that window anywhere
/// but the band of a window nobody can see.
///
/// **A bubble being carried by a hand is left showing its own still**, which is the same answer the
/// release below gives from the other side: a film put back under a moving pointer would be a window
/// chasing a drag (see `pin_bubble_proc` and `bubble_is_being_dragged`).
pub(super) fn keep_the_film_in_the_bubble(box_: ScreenRegion) -> bool {
    if bubble_is_being_dragged() {
        return false;
    }

    if let Some(seated) = seated_film() {
        put_it_back_were_it_has_moved(seated, box_);
        raise_the_bubble();
        return true;
    }

    // Nothing is seated, and whether a film belongs here is the seat's own question — asked again
    // rather than remembered, because what changes under a bubble is the *film*.
    if !seat_film_window(box_) {
        // Nothing to seat this time, and the bubble may be holding a hole painted for a window that
        // is no longer there: the still is what it falls back to, and it is drawn now rather than
        // whenever the bars next move (see `seated_film` and `bubble_art`).
        repaint_the_bubble();
        return false;
    }

    // And the still that goes with the film just seated, where the one in hand is of another film or
    // of none: the still is the ring's colour and the frame the bubble falls back to, and a still of
    // the film before this one is wrong for both (see `grab_bubble_frame`).
    if !the_still_in_hand_is_of_the_pinned_film() {
        grab_bubble_frame(box_);
    }

    repaint_the_bubble();
    true
}

/// Whether a film's window is standing in the bubble now.
///
/// This is the question the bubble's own painting asks, and there is no cheaper way to ask it than
/// this: a window can die between one paint and the next, and a bubble that painted a hole around a
/// window that is not there would be a circle with the desktop in the middle of it.
pub(super) fn film_is_seated() -> bool {
    seated_film().is_some()
}

/// Put a seated film's window back where the bubble has it, and the bubble back over it: what a
/// tick owes a seat something has moved.
///
/// **A window already standing in its box is left exactly where it stands**, and that is not a
/// micro-optimisation: the placement below is a call on somebody else's window, and a tick that
/// re-placed a film every pass would be a film that flashes once a frame.
fn put_it_back_were_it_has_moved(seated: SeatedFilm, box_: ScreenRegion) {
    let window = HWND(seated.hwnd as *mut _);
    unsafe {
        let mut rect = RECT::default();
        if GetWindowRect(window, &mut rect).is_err() {
            return;
        }

        if (rect.left, rect.top, rect.right, rect.bottom) != box_ {
            put_the_film_window_in(window, box_);
        }
    }
}

/// The film's window standing in the bubble, where there is one to hold, and nothing where the
/// seat has gone stale — a seat that is not holding is forgotten here rather than handed back (see
/// the module documentation).
fn seated_film() -> Option<SeatedFilm> {
    let seated = SEATED.lock().ok().and_then(|seat| *seat)?;

    // The player this app holds is the player whose window was seated; a seat whose process has
    // been replaced, or whose window has gone, is a seat standing around the desktop.
    let holds_the_player = VIDEO_PID.load(Ordering::SeqCst) == seated.pid;
    let window_is_there = unsafe { IsWindow(HWND(seated.hwnd as *mut _)).as_bool() };
    if holds_the_player && window_is_there {
        return Some(seated);
    }

    if let Ok(mut seat) = SEATED.lock() {
        *seat = None;
    }
    None
}

/// Put a film's window in a box, in the topmost band, and cut it round: the two are one placement,
/// because the region is in the window's own coordinates and is the circle of the box the window is
/// standing in.
unsafe fn put_the_film_window_in(window: HWND, box_: ScreenRegion) {
    let width = (box_.2 - box_.0).max(1);
    let height = (box_.3 - box_.1).max(1);

    // Topmost and without a z-change: a film holds the band the pin's chrome is drawn over, and the
    // bubble is put above it by the raise below rather than by this placement.
    let _ = SetWindowPos(
        window,
        HWND_TOPMOST,
        box_.0,
        box_.1,
        width,
        height,
        SWP_NOACTIVATE | SWP_NOZORDER | SWP_SHOWWINDOW,
    );

    // **After the placement and not before it**: a window region is in the window's own
    // coordinates, so an ellipse of the old size is a window whose corners are cut to the wrong
    // circle — which is what a second tick of a window mid-move would draw.
    let region = CreateEllipticRgn(0, 0, width, height);
    if !region.is_invalid() {
        // A region the window does not take is this side's to free: `SetWindowRgn` takes it over
        // only where it succeeds.
        if SetWindowRgn(window, region, true) == 0 {
            let _ = DeleteObject(HGDIOBJ(region.0));
        }
    }
}

/// Paint the bubble again for the frame the seat has just changed: what is inside the circle is the
/// film's own window or this app's own face, and which of the two is a thing to be drawn (see
/// `pin_chrome::BubbleArt::Hole`).
fn repaint_the_bubble() {
    unsafe {
        repaint_pin_bubble();
    }
}

/// Put the bubble back over the film standing in it: what the hand is on is the bubble — its click,
/// its drag, its menu and its keys are all its own window's — and the film is read through it (see
/// `BubbleArt::Hole`).
fn raise_the_bubble() {
    let hwnd = PIN_BUBBLE_HWND.load(Ordering::SeqCst);
    if hwnd == 0 {
        return;
    }

    unsafe {
        let _ = SetWindowPos(
            HWND(hwnd as *mut _),
            HWND_TOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
    }
}

#[cfg(test)]
mod tests {
    use super::super::pin_window::PIN_TESTS_ONE_AT_A_TIME;
    use super::*;
    use windows::Win32::Graphics::Gdi::{CreateRectRgn, GetWindowRgn, PtInRegion};
    use windows::Win32::UI::WindowsAndMessaging::DestroyWindow;

    /// The window a film's player would be drawing in, stood in for: a real window of this app's
    /// own, because the box and the region a seat puts on a window are asked of the machine and
    /// nowhere else — nothing this app keeps answers whether either took (see
    /// `stand_a_player_window`, the same stand-in one band over).
    ///
    /// A `STATIC` popup of the machine's own class, so nothing has to be registered.
    unsafe fn stand_a_film_window(box_: ScreenRegion) -> HWND {
        use windows::core::w;

        CreateWindowExW(
            WS_EX_TOOLWINDOW | WS_EX_TOPMOST,
            w!("STATIC"),
            w!(""),
            WS_POPUP,
            box_.0,
            box_.1,
            (box_.2 - box_.0).max(1),
            (box_.3 - box_.1).max(1),
            None,
            None,
            None,
            None,
        )
        .expect("a stand-in window for a film's player")
    }

    /// Where a window stands on the screen right now, read back off the machine.
    fn window_box_on_screen(hwnd: HWND) -> ScreenRegion {
        let mut rect = RECT::default();
        unsafe { GetWindowRect(hwnd, &mut rect) }.expect("a window standing somewhere");
        (rect.left, rect.top, rect.right, rect.bottom)
    }

    /// Whether the window's own region holds a point of the window's coordinates, and whether it
    /// has a region at all — which is what says a seat was cut round and that an unseat took it off
    /// again (see `SetWindowRgn`).
    fn the_region_holds(hwnd: HWND, x: i32, y: i32) -> Option<bool> {
        unsafe {
            let region = CreateRectRgn(0, 0, 1, 1);
            let kind = GetWindowRgn(hwnd, region);
            let holds = (kind.0 != 0).then(|| PtInRegion(region, x, y).as_bool());
            let _ = DeleteObject(HGDIOBJ(region.0));
            holds
        }
    }

    /// The machine as it stood before a test stood a film and a pin in it, so that the test can put
    /// it back: the four slots and the stand-ins a seat reads.
    struct ATestMachine {
        hwnd: isize,
        pid: u32,
        media: Option<MediaData>,
    }

    fn stand_a_film_and_a_pin() -> (HWND, ATestMachine) {
        let machine = ATestMachine {
            hwnd: VIDEO_HWND.load(Ordering::SeqCst),
            pid: VIDEO_PID.load(Ordering::SeqCst),
            media: CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take()),
        };

        // The engine and the player: mpv is the route whose window a bubble can stand around, and
        // every question a seat asks about the route is asked of these two (see `video_route`).
        stand_ffplay_in(Some(false));
        stand_mpv_in(Some(true));
        stand_video_engine_in(Some((VideoEngine::Mpv, true)));

        // The pin's own shape, written out rather than borrowed from the hover tests' `overlay_pin`:
        // that helper is theirs, and a test of a seat is not a test of a hover.
        let mut pin = PinnedPreview {
            bound: Some(64),
            content: (100, 100, 164, 164),
            chrome: PinChrome::always(),
            ..PinnedPreview::for_test()
        };
        pin.path = PathBuf::from("a-film-a-bubble-stands-around.mkv");
        pin.transport.begun(30.0, true, false);
        stand_pin(Some(pin));

        let mut video = create_loading_media(320, 240);
        video.media_type = MediaType::Video;
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = Some(video);
        }

        let window = unsafe { stand_a_film_window((0, 0, 64, 64)) };
        VIDEO_HWND.store(window.0 as isize, Ordering::SeqCst);
        note_player_process(std::process::id(), MPV_IMAGE_NAME);
        // A pipe for the seat's own line to be held and noted rather than lost: a player this app
        // holds no pipe for is a command nothing writes (see `mpv_ipc::deliver`).
        mpv_ipc::remember_pipe(std::process::id(), "a-seat-test-pipe", false);

        // The bubble's own slots, emptied: a still another test left is a still of another film, and
        // a drag left standing would keep the film out of the bubble for every test after it.
        if let Ok(mut frame) = BUBBLE_FRAME.lock() {
            *frame = None;
        }
        take_noted_grabs();
        drop_bubble_drag();

        (window, machine)
    }

    fn put_the_machine_back(window: HWND, machine: ATestMachine) {
        unseat_film_window();
        mpv_ipc::forget_pipe(std::process::id());
        take_noted_grabs();
        drop_bubble_drag();
        VIDEO_HWND.store(machine.hwnd, Ordering::SeqCst);
        note_player_process(machine.pid, MPV_IMAGE_NAME);
        stand_pin(None);
        stand_video_engine_in(None);
        stand_mpv_in(None);
        stand_ffplay_in(None);
        mpv_ipc::clear_noted();
        if let Ok(mut media) = CURRENT_MEDIA.lock() {
            *media = machine.media;
        }
        if let Ok(mut frame) = BUBBLE_FRAME.lock() {
            *frame = None;
        }
        unsafe {
            let _ = DestroyWindow(window);
        }
    }

    /// Whether a film's window is one a bubble can stand around is answered by the kind and the
    /// route and by nothing else: a film the app draws itself is in the app's own window, and
    /// FFmpeg's player cannot be told to fill a square (see `a_film_can_stand_in_the_bubble`).
    #[test]
    fn only_a_film_another_process_draws_and_only_on_mpvs_route_can_stand_in_a_bubble() {
        assert!(a_film_can_stand_in_the_bubble(
            Some(MediaType::Video),
            VideoRoute::Mpv
        ));

        for kind in [
            Some(MediaType::NativeVideo),
            Some(MediaType::Audio),
            Some(MediaType::StaticImage),
            None,
        ] {
            assert!(
                !a_film_can_stand_in_the_bubble(kind, VideoRoute::Mpv),
                "{kind:?} is not a film another process is drawing"
            );
        }

        for route in [
            VideoRoute::Ffplay,
            VideoRoute::MediaEngine,
            VideoRoute::NoPreview,
        ] {
            assert!(
                !a_film_can_stand_in_the_bubble(Some(MediaType::Video), route),
                "{route:?} is not a player whose window a bubble can be built around"
            );
        }
    }

    /// **A seated film is in the bubble's square, cut to its circle, and filling the square** — and
    /// every one of the three comes off when the film is unseated.
    ///
    /// Each is read back off the machine rather than off anything this app writes down, which is
    /// the only way either can be asserted at all (see `stand_a_film_window`).
    #[test]
    fn a_seated_film_stands_in_the_bubbles_square_and_is_cut_round() {
        let _one = PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|e| e.into_inner());
        let (window, machine) = stand_a_film_and_a_pin();
        mpv_ipc::clear_noted();

        let box_ = (300, 300, 364, 364);
        assert!(
            seat_film_window(box_),
            "a film mpv is drawing has a window a bubble can be built around"
        );
        assert!(film_is_seated(), "and the app holds it from here on");

        assert_eq!(
            window_box_on_screen(window),
            box_,
            "the film stands in the box the bubble is: its square, not the pin's band and not \
             where it was"
        );

        // The circle of the square: the middle is in the region and the corners are not, which is
        // what makes what shows through the ring round rather than square.
        assert_eq!(
            the_region_holds(window, 32, 32),
            Some(true),
            "the film fills the middle of the circle the bubble draws around it"
        );
        assert_eq!(
            the_region_holds(window, 2, 2),
            Some(false),
            "and the corners of the square are cut away, or the bubble would be a square picture \
             under a round ring"
        );

        assert_eq!(
            mpv_ipc::noted_lines(),
            vec![format!(
                "{}\n",
                r#"{"command":["set_property","panscan",1.0],"async":true}"#
            )],
            "and mpv is told to fill the square rather than letterbox a stripe inside it"
        );

        unseat_film_window();
        assert!(!film_is_seated(), "the seat is given back");
        assert!(
            !unsafe { IsWindowVisible(window).as_bool() },
            "and the film is taken off the screen with it: a 44-pixel window left standing where              the bubble was is a film in a box on a desktop that has no bubble over it"
        );
        assert_eq!(
            the_region_holds(window, 32, 32),
            None,
            "and the window is a square window again, because the pin's band it goes back into \
             is not a bubble"
        );
        assert_eq!(
            window_box_on_screen(window),
            box_,
            "while its box is left exactly where it was: where the window goes next is not the \
             seat's to decide"
        );
        assert_eq!(
            mpv_ipc::noted_lines()[1],
            format!(
                "{}\n",
                r#"{"command":["set_property","panscan",0.0],"async":true}"#
            ),
            "and mpv fits its window again"
        );

        put_the_machine_back(window, machine);
    }

    /// **A seat whose player is no longer the one this app holds is no seat**, and the bubble stops
    /// claiming one: a window can die between one paint and the next, and a hole painted around a
    /// player that has gone is a circle with the desktop in the middle of it.
    #[test]
    fn a_seat_whose_player_has_gone_is_forgotten_rather_than_painted_around() {
        let _one = PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|e| e.into_inner());
        let (window, machine) = stand_a_film_and_a_pin();

        assert!(seat_film_window((300, 300, 364, 364)));
        assert!(film_is_seated());

        // The player the app holds is no longer the one whose window was seated.
        VIDEO_PID.store(0, Ordering::SeqCst);
        assert!(
            !film_is_seated(),
            "a seat standing around a player this app no longer holds is not one the bubble \
             shows anything through"
        );

        put_the_machine_back(window, machine);
    }

    /// **A window something has moved out from under the ring is put back, and one standing where
    /// it belongs is left exactly there** — the second is not a micro-optimisation: the placement
    /// is a call on somebody else's window, and a tick that made it every pass would be a film that
    /// flashes once a frame (see `keep_the_film_in_the_bubble`).
    ///
    /// **And a film nothing has seated is taken into the bubble by the same tick**, with a still of
    /// its own asked for: a file a walk lands on behind a bubble begins its player while the pin is
    /// off the screen, and its window is otherwise put nowhere but in the band of a window nobody
    /// can see (see `grab_bubble_frame` and the module documentation).
    #[test]
    fn a_film_is_kept_in_the_bubble_a_moved_one_is_put_back_and_one_that_is_there_is_left_alone() {
        let _one = PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|e| e.into_inner());
        let (window, machine) = stand_a_film_and_a_pin();

        let box_ = (300, 300, 364, 364);
        assert!(
            keep_the_film_in_the_bubble(box_),
            "a film mpv is drawing and nothing has seated is taken into the bubble by a tick"
        );
        assert!(film_is_seated());
        assert_eq!(
            window_box_on_screen(window),
            box_,
            "at the box the circle is standing in rather than in the band of a window nobody can              see"
        );
        let asked = take_noted_grabs();
        assert_eq!(asked.len(), 1, "and a still is asked for, once");
        let (film, at, side) = &asked[0];
        assert_eq!(
            film,
            &PathBuf::from("a-film-a-bubble-stands-around.mkv"),
            "for the film the pin is showing: the one in hand is of another, or of none, and a still              is the ring's colour and the frame the bubble falls back to"
        );
        assert!(
            *at >= 30.0,
            "at the second the bar is showing ({at}), which is the transport's own clock"
        );
        assert_eq!(
            *side, 44,
            "and at the circle's own side on this display rather than at the box's width: the box is              the window's square and the still is the circle's (see `PIN_BUBBLE_PIXELS`)"
        );

        // Something moved it: another app's `SetWindowPos`, or mpv itself resizing for a film whose
        // resolution changed mid-film — whichever, the bubble's square is where it belongs.
        unsafe {
            let _ = SetWindowPos(
                window,
                HWND_TOPMOST,
                40,
                40,
                200,
                200,
                SWP_NOACTIVATE | SWP_NOZORDER,
            );
        }
        assert!(
            keep_the_film_in_the_bubble(box_),
            "a seated film is placed by a tick"
        );
        assert_eq!(window_box_on_screen(window), box_);

        let standing = window_box_on_screen(window);
        assert!(
            keep_the_film_in_the_bubble(box_),
            "and a tick over a film where it belongs answers true"
        );
        assert_eq!(
            window_box_on_screen(window),
            standing,
            "while leaving a window that is already standing there alone"
        );
        assert!(
            take_noted_grabs().is_empty(),
            "and asking for no second still: the one in hand is the film's own, which is what a              bubble that has not changed films keeps"
        );

        put_the_machine_back(window, machine);
    }

    /// **A bubble a hand is carrying keeps the film out of it.** A film's window cannot follow a
    /// circle being dragged — that is a placement of somebody else's window on every move the
    /// pointer makes — so a press takes the film out and the bubble shows its own still until the
    /// hand lets go (see `keep_the_film_in_the_bubble` and `pin_bubble_proc`).
    #[test]
    fn a_bubble_being_carried_is_left_showing_its_own_face() {
        let _one = PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|e| e.into_inner());
        let (window, machine) = stand_a_film_and_a_pin();

        assert!(seat_film_window((300, 300, 364, 364)));
        unseat_film_window();

        // The press: the drag is what has hold of the bubble from here (see `pin_bubble_proc`).
        if let Ok(mut drag) = PIN_BUBBLE_DRAG.lock() {
            *drag = Some(((0, 0), (0, 0)));
        }

        let where_it_was = window_box_on_screen(window);
        assert!(
            !keep_the_film_in_the_bubble((300, 300, 364, 364)),
            "a bubble being carried is not given a film to chase the pointer with"
        );
        assert!(!film_is_seated());
        assert_eq!(
            window_box_on_screen(window),
            where_it_was,
            "and nothing of the film's is moved for it"
        );

        // The release, which is where a drag ends in every road but the one that has lost its
        // capture: the film goes back into the circle the hand left the bubble at.
        drop_bubble_drag();
        assert!(
            keep_the_film_in_the_bubble((300, 300, 364, 364)),
            "and the film is put back once the hand has let go"
        );
        assert_eq!(window_box_on_screen(window), (300, 300, 364, 364));

        put_the_machine_back(window, machine);
    }

    /// A kind the seat is not for leaves the player's window alone — no box, no region, no pan and
    /// nothing written down: what such a film has in the bubble is the still (see `bubble_art`).
    #[test]
    fn a_film_the_app_draws_itself_is_not_seated_at_all() {
        let _one = PIN_TESTS_ONE_AT_A_TIME
            .lock()
            .unwrap_or_else(|e| e.into_inner());
        let (window, machine) = stand_a_film_and_a_pin();
        mpv_ipc::clear_noted();

        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            if let Some(media) = current.as_mut() {
                media.media_type = MediaType::NativeVideo;
            }
        }

        let where_it_was = window_box_on_screen(window);
        assert!(
            !seat_film_window((300, 300, 364, 364)),
            "a film this app draws into its own window has no window of somebody else's to seat"
        );
        assert!(!film_is_seated());
        assert_eq!(
            window_box_on_screen(window),
            where_it_was,
            "and nothing of anybody else's is moved for it"
        );
        assert!(
            mpv_ipc::noted_lines().is_empty(),
            "and mpv is told nothing: the film is not one of its"
        );

        put_the_machine_back(window, machine);
    }
}

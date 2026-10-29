use super::*;

// The stand-in for the machine's own `ReleaseCapture` re-entrancy, shared with the tests about
// the roads a release ends (see `pin_input`).
use super::pin_input::CapturingWindow;

// WS-K Phase 1: the capture discipline (a press that arms nothing must claim nothing), the
// swap's reconcile (a take-up gives up the step's leftovers), the chrome a file step raises,
// and the player's own mouse. A gesture never asks anything of a player (see `begin_pin_drag`).

/// A playing pin, as a gesture's press finds it.
fn playing_pin(content: ScreenRegion, from: f64) -> PinnedPreview {
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = PathBuf::from("ws-k-chrome-nav.mkv");
    pin.transport.begun(from, true, false);
    pin.transport.duration = Some(200.0);
    pin.transport.subtitle = Some(2);
    pin.volume.level = 40;
    pin.volume.playing_at = 40;
    pin
}

/// A video pin whose chrome stands in bands — a caption above, a transport below —
/// for the bar's own questions, which are asked of the arrangement rather than of
/// the kind.
fn banded_video_pin(content: ScreenRegion, from: f64) -> PinnedPreview {
    PinnedPreview {
        path: PathBuf::from("ws-k-chrome-nav.mkv"),
        content,
        bound: Some(400),
        zoom: None,
        restore: None,
        dpi: 96,
        audio_scale: None,
        transport_bar: true,
        transport_live: true,
        frame: PinFrame::Shaped,
        overlay: false,
        hides_chrome: false,
        caption: pinned_caption_height(96, Some(MediaType::Video)),
        chrome: PinChrome::always(),
        collapsed: false,
        bubble_pause: None,
        bubble_run: false,
        hovered: None,
        pressed: None,
        last_press: None,
        tooltip: PinTooltip::default(),
        dragging: None,
        transport: {
            let mut transport = PinTransport::default();
            transport.begun(from, true, false);
            transport.duration = Some(200.0);
            transport
        },
        volume: PinVolume::default(),
        audio_hovered: None,
        audio_pressed: None,
        audio_window_buttons: false,
        menu: PinMenu::default(),
    }
}

/// A video behind the pin, as a covered press finds it.
/// A video behind the pin, on FFmpeg's player's route rather than the machine's: with mpv
/// installed the default engine would pick mpv, and every test here that presses a box change is
/// about the kill road FFmpeg's player takes — the mpv route moves the window it has and begins
/// nothing again (see `mpv_ipc`). `restore` puts the stands back.
fn stand_video_media() -> Option<MediaData> {
    stand_ffplay_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Ffmpeg, true)));
    stand_mpv_in(Some(false));
    let previous = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }
    previous
}

fn restore_media(previous: Option<MediaData>) {
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous;
    }
}

fn restore(
    previous_pin: Option<PinnedPreview>,
    previous_pid: u32,
    previous_media: Option<MediaData>,
) {
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    clear_pending_pinned_relaunch();
    clear_restart_count();
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    restore_media(previous_media);
}

/// The state effects of an end-of-road relaunch, without starting a real
/// player: the pid the record names, the generation tag, the transport write.
fn fake_end_relaunch(pid: u32, from: f64, holding: bool) {
    VIDEO_PID.store(pid, Ordering::SeqCst);
    note_pinned_relaunch(pid);
    update_pin_transport(|transport| transport.begun(from, true, holding));
}

/// A point on the seek track of a banded pin's bar, and the proof that it is
/// on the seek track and not on the play button or the volume button beside
/// it: the bar is laid out by arithmetic a test would otherwise have to
/// reproduce to find a point on it at all.
fn seek_track_point() -> (i32, i32) {
    let bar = pinned_transport_geometry().expect("a banded pin has a bar");
    let (x, y) = (bar.width / 2, bar.top + bar.height / 2);
    assert_eq!(
        pin_chrome::transport_part_at(x, y - bar.top, bar.width, bar.height, bar.dpi, bar.live),
        Some(pin_chrome::TransportPart::Seek),
        "the middle of the bar is the seek track"
    );
    (x, y)
}

/// What a press left armed: a button, a scrub's aim, a drag, a knob.
fn armed() -> (bool, Option<f64>, bool, bool) {
    pin_state()
        .and_then(|pinned| {
            pinned.pin().map(|pin| {
                (
                    pin.pressed.is_some() || pin.transport.pressed.is_some(),
                    pin.transport.seeking,
                    pin.dragging.is_some(),
                    pin.volume.dragging,
                )
            })
        })
        .unwrap_or((false, None, false, false))
}

/// A window of this thread's own to hold the pointer, because the last-resort
/// release is guarded by `GetCapture` and nothing less than a real window is ever
/// named by that.
///
/// It is never shown, and the class it is made under is the system's own procedure,
/// so it takes nothing onto the screen. The class is registered once for the process
/// because a second registration of the same name is refused and every test that
/// wants one of these wants the same one.
fn a_real_hidden_window() -> HWND {
    use windows::Win32::UI::WindowsAndMessaging::{RegisterClassExW, WINDOW_EX_STYLE, WNDCLASSEXW};

    static CLASS: std::sync::OnceLock<()> = std::sync::OnceLock::new();
    CLASS.get_or_init(|| unsafe {
        let _ = RegisterClassExW(&WNDCLASSEXW {
            cbSize: std::mem::size_of::<WNDCLASSEXW>() as u32,
            lpfnWndProc: Some(a_procedure_of_its_own),
            hInstance: GetModuleHandleW(None).unwrap_or_default().into(),
            lpszClassName: w!("RustHoverPreviewTestWindow"),
            ..Default::default()
        });
    });

    unsafe {
        CreateWindowExW(
            WINDOW_EX_STYLE(0),
            w!("RustHoverPreviewTestWindow"),
            w!(""),
            WS_POPUP,
            0,
            0,
            4,
            4,
            None,
            None,
            None,
            None,
        )
    }
    .expect("a window of this thread's own")
}

/// The procedure the class above is registered under: the system's own, because
/// this window exists to be named by `GetCapture` and to be let go of, and nothing
/// is ever dispatched to it.
unsafe extern "system" fn a_procedure_of_its_own(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    DefWindowProcW(hwnd, msg, wparam, lparam)
}

/// K2 RELEASE at the seam: the procedure's own last resort is what gives the pointer
/// back where no arm owns it, and it is the only writer on this road — a pin that is
/// not up is not asked anything, so nothing between the message arriving and this
/// call can let the pointer go.
///
/// This is the only test here that stands a real window, because it is the only one
/// that has to: the arm is guarded by `GetCapture`, so anything less than a real
/// window would be testing the guard rather than the release.
#[test]
fn a_release_no_arm_owns_gives_the_pointer_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    // No pin: the state the watchdog and a teardown leave this window in, with a press of
    // its own still standing. No arm of a pinned window's release is even asked, and no
    // text selection and no scroll drag is in flight (see `reconcile_swap_take_up` for the
    // take-down that discards the state a press was written into).
    stand_pin(None);
    set_text_scroll_dragging(false);
    let _ = end_text_selection();

    let hwnd = a_real_hidden_window();
    unsafe {
        SetCapture(hwnd);
        assert_eq!(
            GetCapture(),
            hwnd,
            "this window is holding the pointer before the release: every mouse message on the \
             desktop is arriving here rather than at whatever it was aimed at"
        );

        window_proc(hwnd, WM_LBUTTONUP, WPARAM(0), LPARAM(0));

        assert_eq!(
            GetCapture().0,
            std::ptr::null_mut(),
            "and it is holding nothing after it: a release no arm owns is still the end of a \
             press, and the pointer goes back to the desktop"
        );

        // And the guard is still a guard, which is what makes this safe to leave in the
        // procedure rather than in an arm: a window that is not the one holding the pointer
        // does not take it from whoever is.
        SetCapture(hwnd);
        window_proc(hwnd, WM_LBUTTONUP, WPARAM(0), LPARAM(0));
        assert_eq!(
            GetCapture().0,
            std::ptr::null_mut(),
            "and a release that finds nothing owned takes nothing, having already let go"
        );
    }

    stand_pin(previous_pin);
    restore_media(previous_media);
}

/// A banded video pin wide enough for its caption to carry the whole of its row.
///
/// The walk of four is dropped whole from a caption too narrow for it rather than half of
/// it, so a pin 320 across has no `Next` on it at all, and every question about a caption
/// button would be about the three that are a window's own.
fn wide_banded_video_pin(content: ScreenRegion, from: f64) -> PinnedPreview {
    banded_video_pin((content.0, content.1, content.0 + 700, content.3), from)
}

/// A point on a named button of the caption above the pin that is up.
///
/// Found by asking the caption rather than by reproducing its arithmetic, and filtered
/// down to a point that is not also on a resize edge — because the frame is asked about
/// before the caption is (see `pinned_press`), so the two share the three places a window
/// can be resized from and only the caption keeps the rest.
fn a_caption_button_point(kind: pin_chrome::CaptionButton) -> (i32, i32) {
    let caption = pinned_caption_geometry().expect("a pin with a caption above it");
    let framed = caption.frame != PinFrame::None;

    (0..caption.height)
        .flat_map(|y| (0..caption.width).map(move |x| (x, y)))
        .find(|(x, y)| {
            pin_chrome::button_at(*x, *y, caption.width, caption.height, caption.dpi, framed)
                == Some(kind)
                && !pin_state().is_some_and(|pinned| {
                    pinned
                        .pin()
                        .is_some_and(|pin| pin.resize_edge(*x, *y).is_some())
                })
        })
        .unwrap_or_else(|| panic!("this caption draws no {kind:?} a hand can reach"))
}

/// A point on a named part of the transport bar of the pin that is up: the volume
/// button is the one the caption's own arm is not reached from, and it is a bar's
/// part rather than a caption's, so it is found the same way the seek's is.
fn a_transport_part_point(part: pin_chrome::TransportPart) -> (i32, i32) {
    let bar = pinned_transport_geometry().expect("a banded pin has a bar");
    let y = bar.top + bar.height / 2;

    (0..bar.width)
        .map(|x| (x, y))
        .find(|(x, y)| {
            pin_chrome::transport_part_at(
                *x,
                *y - bar.top,
                bar.width,
                bar.height,
                bar.dpi,
                bar.live,
            ) == Some(part)
        })
        .unwrap_or_else(|| panic!("this bar draws no {part:?} to press"))
}

/// What a press left armed on the pin: the caption's own button.
fn the_pressed_button() -> Option<pin_chrome::CaptionButton> {
    pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin.pressed))
}

/// K2 AT THE CAPTION: every button on a pinned window's caption is a press that arms
/// it and a release that asks its command, with the capture taken on the way in and
/// given back on the way out.
///
/// The release is driven through the machine's own re-entrancy (`CapturingWindow`),
/// because that is the whole of what went wrong: an arm before the caption's own that
/// lets go of the pointer re-enters `pin_capture_lost` before the caption has read what
/// the press armed, and that road drops the pressed button — so the caption finds
/// nothing pressed, and every button on a pinned window is a button that does nothing,
/// from the file walk to the maximize. A recorder that merely recorded the release would
/// hide all of it.
#[test]
fn every_caption_button_asks_its_command_on_a_press_and_a_release() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    for (kind, command) in [
        (pin_chrome::CaptionButton::Previous, PinCommand::Previous),
        (pin_chrome::CaptionButton::Next, PinCommand::Next),
        (pin_chrome::CaptionButton::Minimize, PinCommand::Minimize),
        (pin_chrome::CaptionButton::Maximize, PinCommand::Maximize),
        (pin_chrome::CaptionButton::Close, PinCommand::Close),
    ] {
        stand_pin(Some(wide_banded_video_pin((100, 80, 420, 320), 30.0)));
        take_pin_command();

        let (x, y) = a_caption_button_point(kind);
        assert!(
            unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
            "a hand on {kind:?} is the caption's to act on"
        );
        assert_eq!(
            the_pressed_button(),
            Some(kind),
            "so the press armed it, and took the capture with it"
        );

        let window = CapturingWindow::around(RecordedPinWindow::new(0x1000));
        assert!(
            unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) },
            "and the pointer is still on it, so the release is the caption's own to answer"
        );
        assert_eq!(
            take_pin_command(),
            Some(command),
            "which asks for the command that button exists to ask for"
        );
    }

    restore(previous_pin, previous_pid, previous_media);
}

#[test]
fn the_always_on_top_caption_button_toggles_pin_always_on_top() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(wide_banded_video_pin((100, 80, 420, 320), 30.0)));
    take_pin_command();

    let (x, y) = a_caption_button_point(pin_chrome::CaptionButton::AlwaysOnTop);
    assert!(
        unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on AlwaysOnTop is the caption's to act on"
    );
    assert_eq!(
        the_pressed_button(),
        Some(pin_chrome::CaptionButton::AlwaysOnTop),
        "so the press armed it, and took the capture with it"
    );

    let before = pin_always_on_top();
    let window = CapturingWindow::around(RecordedPinWindow::new(0x1000));
    assert!(
        unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) },
        "so the release is the caption's own to answer"
    );
    assert_eq!(pin_always_on_top(), !before, "AlwaysOnTop was toggled");
    assert_eq!(the_pressed_button(), None, "and the button was let go of");

    // Toggle back to restore initial state
    assert!(unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) });
    assert!(unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) });
    assert_eq!(pin_always_on_top(), before, "restored to original state");

    restore(previous_pin, previous_pid, previous_media);
}

/// K2 AT THE CAPTION, for the restore-down: a maximized pin remembers the box it had,
/// and the caption draws a restore glyph where it drew the maximize — the same button,
/// so the restore-down is dead or alive with the maximize rather than beside it.
#[test]
fn a_restore_down_button_asks_its_command_too() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    // A maximized pin is one that remembers the box it had, and the caption draws a
    // restore glyph where it drew the maximize.
    let mut pin = wide_banded_video_pin((100, 80, 420, 320), 30.0);
    pin.restore = Some((100, 80, 420, 320));
    stand_pin(Some(pin));
    let (x, y) = a_caption_button_point(pin_chrome::CaptionButton::Maximize);
    take_pin_command();

    assert!(
        unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
        "a restore-down is a caption button like any other"
    );
    assert_eq!(
        the_pressed_button(),
        Some(pin_chrome::CaptionButton::Maximize),
        "and the press armed it"
    );

    let window = CapturingWindow::around(RecordedPinWindow::new(0x1000));
    assert!(
        unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) },
        "so the release is the caption's own to answer"
    );
    assert_eq!(
        take_pin_command(),
        Some(PinCommand::Maximize),
        "and it asks for the same command the maximize does — the loop is what reads the \
         remembered box and puts it back (see `toggle_pin_maximized`)"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K2 AT THE CAPTION: the volume button is the one button of a pinned window's
/// chrome that is not on the caption, and it is dead for the same reason: the bar
/// armed it on the press, and every arm before the bar's own had something to decline.
#[test]
fn the_volume_button_opens_and_closes_its_popup() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(banded_video_pin((100, 80, 420, 320), 30.0)));

    let (x, y) = a_transport_part_point(pin_chrome::TransportPart::Volume);
    let window = CapturingWindow::around(RecordedPinWindow::new(0x1000));

    assert!(
        unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on the level is the bar's to act on"
    );
    assert!(
        unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) },
        "and the release is the bar's to answer"
    );
    assert!(
        pin_volume_open(),
        "so the popup is open, which is what a click on the button does"
    );

    // And the other click puts it away, which is the same press and the same release
    // read out of the state the first one left.
    assert!(
        unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
        "a second hand on the level is the same button"
    );
    assert!(
        unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) },
        "and the same release"
    );
    assert!(!pin_volume_open(), "so the popup is put away again");

    restore(previous_pin, previous_pid, previous_media);
}

/// K2 AT THE CAPTION, for the two hand-offs: the same press arms them and the same
/// release reaches their arm, which takes the button off the pin and hands the file to
/// the Shell. The Shell itself is not called from here — it starts a program, and a
/// test that starts one is a test that leaves a window on the user's desktop — so what
/// is read is the arm being reached: a decline by any arm before it leaves the button
/// on the pin, and the release falls out of the road as a release nothing answered.
#[test]
fn the_two_hand_offs_reach_their_arm_on_a_press_and_a_release() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(wide_banded_video_pin((100, 80, 420, 320), 30.0)));

    for kind in [
        pin_chrome::CaptionButton::OpenWith,
        pin_chrome::CaptionButton::OpenWithList,
    ] {
        let (x, y) = a_caption_button_point(kind);
        assert!(
            unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
            "a hand on {kind:?} is the caption's to act on"
        );
        assert_eq!(the_pressed_button(), Some(kind), "so the press armed it");

        // Released off the button, so the arm is reached and declines the hand-off rather
        // than answering it: the Shell is what this arm would call, and a test must not
        // start one. The button still has to come off the pin, which is what says the arm
        // was reached at all.
        let (away_x, away_y) = (0, caption_height_of_the_pin());
        let window = CapturingWindow::around(RecordedPinWindow::new(0x1000));
        assert!(
            unsafe { pinned_release(HWND(0x1000 as *mut _), away_x, away_y, &window) },
            "so the release is the caption's own to answer"
        );
        assert_eq!(
            the_pressed_button(),
            None,
            "and the button came off the pin rather than being left for a release nothing owned"
        );
    }

    restore(previous_pin, previous_pid, previous_media);
}

/// The height of the strip above the media of the pin that is up, which is where every
/// point outside it is: a caption button is answered only of a point on it (see
/// `pinned_release`), so this is how a release says "not on the button" without leaving
/// the window.
fn caption_height_of_the_pin() -> i32 {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.caption))
        .unwrap_or(0)
}

#[test]
fn a_seek_press_with_no_second_to_seek_to_claims_nothing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let mut pin = banded_video_pin((100, 80, 420, 320), 30.0);
    pin.transport.duration = None;
    stand_pin(Some(pin));

    let (x, y) = seek_track_point();
    assert!(
        !unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) },
        "a press with no second to seek to is not the bar's to act on"
    );
    assert_eq!(
        armed(),
        (false, None, false, false),
        "nothing is claimed: no button, no aim, no drag, no knob"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K1 keeps the press that *can* seek: a bar a second can be aimed on is the
/// bar's own to take hold of, capture and all.
#[test]
fn a_seek_press_with_a_second_still_arms_the_scrub() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(banded_video_pin((100, 80, 420, 320), 30.0)));

    let (x, y) = seek_track_point();
    assert!(
        unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) },
        "a bar with a length to seek on is taken hold of"
    );
    let (_, aim, _, _) = armed();
    assert!(
        aim.is_some(),
        "and the scrub has its aim, which is what the release takes the file to"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K2 RELEASE: an arm that finds nothing to answer for leaves the pointer
/// exactly as it found it. The transport arm is the one the leak was actually
/// found in, and it is asked of every release on a pinned window that is not a
/// knob and not the bar: a press on a caption button is a hand on a title bar,
/// and the bar declining it must cost nothing — the capture that press took
/// belongs to the caption's own arm, which has not been asked yet.
#[test]
fn a_transport_release_with_nothing_armed_leaves_the_pointer_alone() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(banded_video_pin((100, 80, 420, 320), 30.0)));

    let (x, y) = seek_track_point();
    let window = RecordedPinWindow::new(0x1000);
    assert!(
        !unsafe { pinned_transport_release(HWND(0x1000 as *mut _), x, y, &window) },
        "nothing armed is nothing to answer"
    );
    assert_eq!(
        window.calls(),
        Vec::new(),
        "and the pointer is not touched: a release raised by an arm with nothing to say \
         re-enters `pin_capture_lost` before the caption's arm has read what the press armed, \
         and that road drops the pressed button out from under it"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K2 RELEASE for the card's own buttons, which a rebuilt pin can empty out
/// from under a press the same way — and which are asked of every release that
/// is not a knob, the bar, or a caption button.
#[test]
fn a_card_release_with_no_button_held_leaves_the_pointer_alone() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(banded_video_pin((100, 80, 420, 320), 30.0)));

    let (x, y) = seek_track_point();
    let window = RecordedPinWindow::new(0x1000);
    assert!(
        !unsafe { pinned_audio_control_release(HWND(0x1000 as *mut _), x, y) },
        "no button of the card was held"
    );
    assert_eq!(
        window.calls(),
        Vec::new(),
        "so this arm declines in silence, leaving the capture to whichever arm did take it"
    );

    stand_pin(previous_pin);
    restore_media(previous_media);
}

/// K2 RELEASE for the road as a whole: a release whose every arm finds nothing
/// armed still ends with no capture held, because the drag's own arm is the last
/// one asked and it lets go of a capture it cannot find a drag for — and on the
/// machine, the procedure's last resort after it (see `release_pin_capture`).
///
/// Every arm declining must still add up to one release rather than to four, which is
/// what the recorder below reads: four arms each letting go of the pointer they were
/// lent is four re-entered capture-lost roads for one press.
#[test]
fn a_release_with_nothing_armed_anywhere_gives_the_pointer_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(banded_video_pin((100, 80, 420, 320), 30.0)));

    let (x, y) = seek_track_point();
    let window = RecordedPinWindow::new(0x1000);
    assert!(
        !unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) },
        "nothing anywhere is armed"
    );
    // One release, from the arm that owns the pointer — the drag's, which is the last arm
    // asked and so cannot preempt any other. What must not appear is any other window work: a
    // refused release ends here and does nothing else.
    let calls = window.calls();
    assert_eq!(
        calls,
        vec![PinWindowCall::ReleaseCapture],
        "the road lets the pointer go once and does nothing else (got {calls:?})"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K2 RELEASE for the knob's own arm: a press the swap disarmed under the
/// hand leaves `dragging` standing nowhere, and this arm is the very first of
/// them to be asked — so a release it declines must cost nothing at all.
#[test]
fn a_volume_release_with_no_knob_held_leaves_the_pointer_alone() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(banded_video_pin((100, 80, 420, 320), 30.0)));

    let window = RecordedPinWindow::new(0x1000);
    assert!(
        !unsafe { pinned_volume_release(HWND(0x1000 as *mut _), &window) },
        "no knob was held, so there is no knob to let go of"
    );
    assert_eq!(
        window.calls(),
        Vec::new(),
        "and this arm says nothing about the pointer: it is asked of every release on a pinned \
         window, so a release here means the pointer belongs to an arm further down"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K2 RELEASE for the drag's arm, which is the arm a title bar and an edge
/// both end in: the pin rebuilt under the hand takes its drag with it, and
/// this is the arm that is left holding the pointer.
#[test]
fn a_drag_end_with_no_drag_gives_the_pointer_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(banded_video_pin((100, 80, 420, 320), 30.0)));

    let window = RecordedPinWindow::new(0x1000);
    assert!(
        !finish_pin_drag(HWND(0x1000 as *mut _), &window),
        "no drag is no drag to end"
    );
    assert_eq!(
        window.calls(),
        vec![PinWindowCall::ReleaseCapture],
        "so the pointer is let go of here too: one rule, every arm"
    );

    restore(previous_pin, previous_pid, previous_media);
}
/// K3 RECONCILE: a swap under a hand gives up what a file step inherits. The in-flight relaunch
/// is reaped rather than left behind a cover nothing will hand the band to, and the box request a
/// gesture wrote belongs to the file that made it rather than to the one arriving.
#[test]
fn a_swap_reconciles_the_gesture_it_is_taking_up_with() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));
    let generation = current_pinned_generation();
    fake_end_relaunch(102, 30.0, false);
    if let Ok(mut request) = PIN_BOX_REQUEST.lock() {
        *request = Some((100, 80, 420, 320));
    }

    reconcile_swap_take_up();

    assert!(
        pending_pinned_relaunch().is_none(),
        "the relaunch behind it is reaped rather than left for the next run's sweep"
    );
    assert!(
        PIN_BOX_REQUEST
            .lock()
            .ok()
            .and_then(|request| *request)
            .is_none(),
        "and the box request goes with the file that made it"
    );
    assert!(
        current_pinned_generation() == generation,
        "and no generation is bumped for a reconcile, which is not a gesture"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "and that relaunch is the outgoing film's own, so its kill is confirmed: nothing of the          film the pin left is left standing"
    );
    assert!(
        orphaned_video_kills().is_empty(),
        "reaped rather than parked for the next run's sweep"
    );

    restore(previous_pin, previous_pid, previous_media);
}

#[test]
fn a_step_from_the_caption_leaves_the_chrome_up_until_the_file_it_asked_for_lands() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 800, 320);
    let mut pin = wide_banded_video_pin(content, 30.0);
    pin.hides_chrome = true;
    stand_pin(Some(pin));
    pin_film_host::note_hosted_player(0, None);
    take_pin_command();
    if let Ok(mut pending) = PIN_JOBS.0.lock() {
        *pending = None;
    }

    // A hand clicks Next and lets go on it: the command is the loop's, and the step is the one
    // the walk and the load will run for.
    let (x, y) = a_caption_button_point(pin_chrome::CaptionButton::Next);
    assert!(
        unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on Next is the caption's to act on"
    );
    let window = CapturingWindow::around(RecordedPinWindow::new(0x1000));
    assert!(
        unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) },
        "and its release asks the step"
    );
    assert_eq!(take_pin_command(), Some(PinCommand::Next));

    let mut wait: Option<PinWait> = None;
    assert!(
        step_pinned_file(1, &mut wait, None).is_none(),
        "the click is a walk asked, not a file yet"
    );
    assert!(wait.is_some(), "so the planner owes it an answer");
    if let Ok(mut pending) = PIN_JOBS.0.lock() {
        *pending = None;
    }

    let box_ = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_box()))
        .expect("a pin up");
    let cursor = Some((box_.0 + x, box_.1 + y));

    // No cover and no park yet — this is the gap the user sees, and the outgoing film is still
    // live under it, so the chrome stands rather than leaving a lone strip over an old frame.
    pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(refresh_pin_chrome(pin, Instant::now(), cursor))
    });
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| (pin.chrome.caption, pin.chrome.bar))),
        Some((true, false)),
        "and the step leaves the chrome up while the outgoing film is live, with the pointer still on the chevron it clicked"
    );

    // The file lands, and this stand-in is a caption *above* its media while the kind the file is
    // arrives with its chrome drawn over the film — so the step is a change of kind, and the take-up
    // arrives as the new kind's own chrome does: hidden (see `take_up_pinned_window`). A step of
    // one kind carries the chrome the pointer asked for; that is the strip tests' question (see
    // `a_step_leaves_the_chrome_that_the_pointer_is_resting_on_up`).
    let arrived = take_up_pinned_window(Path::new("ws-k-chrome-nav-next.mkv"), content);
    stand_pin(Some(arrived));

    let shown = pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(refresh_pin_chrome(pin, Instant::now(), cursor))
    });
    assert_eq!(
        shown,
        Some(true),
        "and the same pointer asks the caption back over the file that arrived"
    );
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| (pin.chrome.caption, pin.chrome.bar))),
        Some((true, false)),
        "the caption is back, with the bar left to its own end of the pin"
    );

    if let Ok(mut pending) = PIN_JOBS.0.lock() {
        *pending = None;
    }
    restore(previous_pin, previous_pid, previous_media);
}

/// A step is asked and the outgoing film is still live: the chrome stands, with no suppression
/// raised by the ask alone.
#[test]
fn a_step_ask_leaves_the_chrome_up_while_the_outgoing_film_is_live() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 800, 320);
    let mut pin = wide_banded_video_pin(content, 30.0);
    pin.hides_chrome = true;
    stand_pin(Some(pin));
    pin_film_host::note_hosted_player(0, None);
    take_pin_command();
    if let Ok(mut pending) = PIN_JOBS.0.lock() {
        *pending = None;
    }

    let (x, y) = a_caption_button_point(pin_chrome::CaptionButton::Next);
    assert!(unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) });
    let window = CapturingWindow::around(RecordedPinWindow::new(0x1000));
    assert!(unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) });
    assert_eq!(take_pin_command(), Some(PinCommand::Next));
    let mut wait = None;
    assert!(step_pinned_file(1, &mut wait, None).is_none());
    assert!(wait.is_some(), "the planner owes the step an answer");
    if let Ok(mut pending) = PIN_JOBS.0.lock() {
        *pending = None;
    }

    let box_ = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_box()))
        .expect("a pin up");
    let cursor = Some((box_.0 + x, box_.1 + y));
    pin_state().and_then(|mut pinned| {
        let pin = pinned.pin_mut()?;
        Some(refresh_pin_chrome(pin, Instant::now(), cursor))
    });
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().map(|pin| (pin.chrome.caption, pin.chrome.bar))),
        Some((true, false)),
        "and the chrome stands over the live film, with the pointer still on the chevron it clicked"
    );

    restore(previous_pin, previous_pid, previous_media);
}

#[test]
fn a_box_change_begins_no_player() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();
    clear_restart_count();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        101,
        "a player is behind the band"
    );

    for road in [PinRelayoutRoad::TakeUp, PinRelayoutRoad::BoxChange] {
        clear_restart_count();
        relayout_pinned_media(
            &PathBuf::from("ws-k-chrome-nav.mkv"),
            content,
            96,
            None,
            road,
        );

        assert_eq!(
            restart_count(),
            0,
            "{road:?} begins no player: the film follows its host"
        );
        assert_eq!(
            last_restart_reason(),
            None,
            "so no reason line is written for a box"
        );
        assert_eq!(
            VIDEO_PID.load(Ordering::SeqCst),
            101,
            "and the player is the one that was there"
        );
    }

    restore(previous_pin, previous_pid, previous_media);
}

#[test]
fn a_maximize_remembers_its_box_and_a_restore_gives_it_back() {
    let room = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };
    let content = (100, 100, 500, 400);

    let (maximized, restore) = pin_maximize_decided(PinMaximize {
        frame: PinFrame::Shaped,
        content,
        restore: None,
        shape: Some((400, 300)),
        room,
    });
    assert_eq!(
        restore,
        Some(content),
        "a maximize remembers the box it had"
    );
    assert!(
        maximized.2 - maximized.0 <= room.right - room.left
            && maximized.3 - maximized.1 <= room.bottom - room.top,
        "fitted inside the room, got {maximized:?}"
    );

    let (back, restore) = pin_maximize_decided(PinMaximize {
        frame: PinFrame::Shaped,
        content: maximized,
        restore,
        shape: Some((400, 300)),
        room,
    });
    assert_eq!(restore, None, "a restore gives the maximize up");
    assert_eq!(
        (back.2 - back.0, back.3 - back.1),
        (content.2 - content.0, content.3 - content.1),
        "and puts back the size the user had"
    );
}

/// K6 MOUSE: the player's window is made mouse-transparent where the monitor
/// thread already asserts its styles, so a hand on the band cannot reach FFmpeg's
/// own bindings. `left double-click toggle full screen` is compiled into the
/// player and no option unbinds it (verified against ffplay 9.0.2), so this has to
/// be a window property and not a flag.
#[test]
fn the_player_window_is_mouse_transparent() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();
    let previous_front = PIN_FRONT_OWED.load(std::sync::atomic::Ordering::SeqCst);
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);
    // A pin that is Always On Top, which is the state the film is topmost in.
    stand_pin(Some(overlay_pin((100, 80, 420, 320), PinChrome::always())));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
        config.pin_update_move_to_top = true;
    }
    assert!(
        pin_holds_the_front(),
        "a pin that is Always On Top holds the band, and that is the band the film is asserted into"
    );

    let style = player_ex_style_for(0);

    assert_eq!(
        style & WS_EX_TRANSPARENT.0 as isize,
        WS_EX_TRANSPARENT.0 as isize,
        "clicks on the band fall through the player's own window rather than reaching its SDL \
         handler"
    );
    assert_eq!(
        style & WS_EX_NOACTIVATE.0 as isize,
        WS_EX_NOACTIVATE.0 as isize,
        "and it still cannot take the keyboard"
    );
    assert_eq!(
        style & WS_EX_TOPMOST.0 as isize,
        WS_EX_TOPMOST.0 as isize,
        "and it is still kept above Explorer"
    );
    assert_eq!(
        style & WS_EX_TOOLWINDOW.0 as isize,
        WS_EX_TOOLWINDOW.0 as isize,
        "and it is still out of the taskbar"
    );
    assert_eq!(
        style & WS_EX_LAYERED.0 as isize,
        0,
        "and it is not made layered: a layered window draws from UpdateLayeredWindow and this \
         one is SDL's own surface, so this is the bit that could stop a film rendering"
    );
    // Idempotent, and everything the window was created with is kept: the monitor
    // thread asserts this list about five times a millisecond.
    assert_eq!(
        player_ex_style_for(style),
        style,
        "re-asserting changes nothing"
    );
    assert_eq!(
        player_ex_style_for(0x0004_0000isize) & 0x0004_0000isize,
        0x0004_0000isize,
        "and a style the player was created with is not dropped"
    );

    PIN_FRONT_OWED.store(previous_front, std::sync::atomic::Ordering::SeqCst);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// A box change on the mpv route is a move and not a relaunch: mpv rescales itself in the window
/// it already has, so nothing is parked, killed or begun again (see `relayout_pinned_media`).
///
/// The two roads a take-up and a resize walk are asserted together, because both read the same
/// player: the film on screen is moved into the new box and plays on through it.
#[test]
fn a_box_change_moves_mpv_and_begins_nothing_again() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    // mpv rather than the machine's default, which `stand_video_media` stands on FFmpeg's player.
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    stand_ffplay_in(Some(false));

    // The machine's own engine: a resize is its own call and nothing is begun again — the arm this
    // pins alongside FFmpeg's player's relaunch. Put back to the film mpv draws below.
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        if let Some(media) = media.as_mut() {
            media.media_type = MediaType::NativeVideo;
        }
    }
    clear_restart_count();
    relayout_pinned_media(
        &PathBuf::from("ws-k-chrome-nav.mkv"),
        (100, 80, 420, 320),
        96,
        None,
        PinRelayoutRoad::BoxChange,
    );
    assert_eq!(
        restart_count(),
        0,
        "the engine is resized where it stands and begins nothing"
    );
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        if let Some(media) = media.as_mut() {
            media.media_type = MediaType::Video;
        }
    }

    let content = (100, 80, 420, 320);
    let previous_own = pin_is_its_own_window();
    stand_pin(Some(playing_pin(content, 30.0)));
    // **A film pin stands in a window of its own**, which is where its host is made: a resize or a
    // take-up is a box for that host and not a player begun again (see `pin_film_host`).
    pin_stands_in_its_own_window(true);
    clear_restart_count();

    relayout_pinned_media(
        &PathBuf::from("ws-k-chrome-nav.mkv"),
        content,
        96,
        None,
        PinRelayoutRoad::BoxChange,
    );
    assert_eq!(
        restart_count(),
        0,
        "mpv rescales itself, so a resize is a move and begins nothing again"
    );

    // The take-up road answers the same way, for the same player.
    relayout_pinned_media(
        &PathBuf::from("ws-k-chrome-nav.mkv"),
        content,
        96,
        None,
        PinRelayoutRoad::TakeUp,
    );
    assert_eq!(
        restart_count(),
        0,
        "a take-up moves the window that is already showing the film and begins nothing"
    );

    pin_stands_in_its_own_window(previous_own);
    restore(previous_pin, previous_pid, previous_media);
}

/// The film's own window is told which band it is to hold, because the pin's setting and the
/// player's own flag disagree the moment a hover's film is adopted: `--ontop` was for the
/// preview's sake, and a pin that is not Always On Top is a window the user may put a program in
/// front of. The same ask answers the setting being changed under a film that is already running,
/// which is the whole of why it is an IPC property rather than a flag on a relaunch.
#[test]
fn a_pinned_mpv_film_is_told_the_pins_own_answer_about_the_topmost_band() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_image = video_process_image();
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();
    let before = pin_always_on_top();

    // The player this app holds, named as the film on screen, and the pipe it is told on.
    stand_pin(Some(playing_pin((100, 80, 420, 320), 0.0)));
    note_player_process(101, MPV_IMAGE_NAME);
    mpv_ipc::remember_pipe(101, "a-pinned-ontop-pipe", false);
    mpv_ipc::clear_noted();

    let set_top = |top: bool| {
        if let Ok(mut config) = CONFIG.lock() {
            config.pin_always_on_top = top;
        }
    };

    set_top(false);
    settle_pinned_mpv_ontop();
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![format!(
            "{}\n",
            r#"{"command":["set_property","ontop",false],"async":true}"#
        )],
        "a pin that is not Always On Top tells its film to leave the topmost band"
    );

    set_top(true);
    mpv_ipc::clear_noted();
    settle_pinned_mpv_ontop();
    assert_eq!(
        mpv_ipc::noted_lines(),
        vec![format!(
            "{}\n",
            r#"{"command":["set_property","ontop",true],"async":true}"#
        )],
        "and one that is tells it to hold the band the pin's chrome is raised over"
    );

    // FFmpeg's player has no opinion of its own to correct: this app places its window, so the
    // route that is not mpv's is told nothing at all.
    note_player_process(101, FFPLAY_IMAGE_NAME);
    mpv_ipc::clear_noted();
    settle_pinned_mpv_ontop();
    assert!(
        mpv_ipc::noted_lines().is_empty(),
        "FFmpeg's player is placed by this app and is told nothing about bands"
    );

    set_top(before);
    restore(previous_pin, previous_pid, previous_media);
    note_player_process(previous_pid, previous_image);
    mpv_ipc::forget_pipe(101);
    mpv_ipc::clear_noted();
}

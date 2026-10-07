use super::*;

// The stand-in for the machine's own `ReleaseCapture` re-entrancy and the one reader of the
// relayout slot, shared with the tests about the roads a release ends (see `pin_input`).
use super::pin_input::{take_relayout_request, CapturingWindow};

// WS-K Phase 1: the capture discipline (a press that arms nothing must claim
// nothing), the swap's reconcile (a take-up carries no gesture), the two
// covers (a file step and a box change), and the player's own mouse.
//
// RED (pre-fix): the seams below — `begin_covered_box_change`,
// `cover_step_swap_for_video`, `reconcile_swap_take_up` — do not exist yet, so
// this file does not compile; the capture tests are behavioural red through
// the arms that already exist.

/// A playing pin, as a covered road's press finds it.
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
        hovered: None,
        pressed: None,
        tooltip: PinTooltip::default(),
        dragging: None,
        parked: false,
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
fn stand_video_media() -> Option<MediaData> {
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
    take_gesture_snapshot();
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    clear_pending_pinned_relaunch();
    clear_park_swap_arm();
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

fn spend_the_swap_bound() {
    if let Ok(mut held) = PIN_PARK_SWAP.lock() {
        if let Some(swap) = held.as_mut() {
            swap.since = Some(Instant::now() - PIN_PARK_SWAP_TIMEOUT * 4);
        }
    }
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
    take_gesture_snapshot();
    forget_pin_park_swap();
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
        forget_pin_park_swap();
        forget_video_frame();
        take_gesture_snapshot();
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
    forget_pin_park_swap();
    take_gesture_snapshot();
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
    forget_pin_park_swap();
    forget_video_frame();
    take_gesture_snapshot();

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
    forget_pin_park_swap();
    forget_video_frame();
    take_gesture_snapshot();

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

/// K2 DRAG after a step: the cover a file step raises is still standing when the hand
/// goes on to carry the window or pull its edge, and the gesture must neither strand it
/// nor leave it standing for want of the relayout that ends it.
///
/// A placeholder left in the video area is that cover: a band painted with the outgoing
/// frame, held until there is a player's own window in it to see through. Both ways it can
/// be stranded are read here — `park_stranded_without_a_player` and `seek_cover_is_waiting`
/// — and so is the relayout, because a drag whose own end never runs leaves a standing
/// cover with nothing behind it to be handed, which is the cover the user was left looking
/// at. And a button afterwards, because a window whose caption has stopped answering is the
/// other half of the same report.
///
/// The resize's park is armed rather than begun, and that is not a shortcut: a begun
/// resize asks for the frame it hands the band back and spawns a player to render it,
/// which a test must not do. Its end is the same arm either way (see `finish_pin_drag`).
#[test]
fn a_drag_after_a_file_step_strands_no_cover_and_no_park() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(wide_banded_video_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    take_gesture_snapshot();
    take_relayout_request();

    assert!(
        cover_step_swap_for_video(true),
        "a step onto a video raises the cover it hands to the incoming player"
    );
    // A player standing behind the cover, published without arming a relaunch of its own:
    // the cover is waiting for a window to see through it, not for a process to begin. A
    // process id no machine hands out, so nothing here is ever a process to end.
    VIDEO_PID.store(u32::MAX, Ordering::SeqCst);

    let window = CapturingWindow::around(a_window_at(content));
    begin_pin_drag(HWND(0x1000 as *mut _), &window, PinDragAction::Move, true);
    assert!(pin_is_dragging(), "the hand is carrying the window");

    let bar = pinned_transport_geometry().expect("a banded pin has a bar");
    assert!(
        unsafe {
            pinned_release(
                HWND(0x1000 as *mut _),
                content.0 + 10,
                bar.top - 10,
                &window,
            )
        },
        "and the release is the drag's own to answer: the three arms before it each had nothing \
         armed and each declined in silence, so the drag's record was still on the pin to read"
    );
    assert!(!pin_is_dragging(), "so the drag is over");
    assert!(
        !park_stranded_without_a_player() && !seek_cover_is_waiting(),
        "and the cover is not left standing over a band nothing is going to hand back — which is \
         what a placeholder in the video area is"
    );
    assert!(
        take_relayout_request().is_none(),
        "and a move asked for no relayout of its own"
    );

    // The edge, which is the one drag that ends in a relaunch.
    with_pin(|pin| {
        pin.dragging = Some(PinDrag {
            from: (0, 0),
            window: content,
            action: PinDragAction::Resize(PinResize {
                left: false,
                top: false,
                right: true,
                bottom: true,
            }),
            delivered: true,
            carried: (i32::MIN, i32::MIN),
        });
    });
    assert!(pin_is_dragging(), "and the hand is on the corner");
    assert!(
        unsafe {
            pinned_release(
                HWND(0x1000 as *mut _),
                content.0 + 10,
                bar.top - 10,
                &window,
            )
        },
        "so the resize's release is its own too"
    );
    assert!(
        take_relayout_request().is_some(),
        "and the resize asked for its media at the box the hand settled on"
    );
    assert!(
        !park_stranded_without_a_player() && !seek_cover_is_waiting(),
        "with a relaunch of its own behind it, so nothing is stranded either"
    );

    // And the window is still a window a hand can use: the whole of the report was that
    // the caption had stopped answering.
    let (x, y) = a_caption_button_point(pin_chrome::CaptionButton::Next);
    assert!(
        unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
        "a button is still a button after a step and a drag"
    );
    assert!(
        unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) },
        "and its release is still the caption's own"
    );
    assert_eq!(
        take_pin_command(),
        Some(PinCommand::Next),
        "so the walk still moves a file along"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K1 PRESS: a seek press with no second to seek to arms nothing — no aim, no
/// button, no cover, and no capture, because every release arm answers out of
/// what the press armed and a press that armed nothing left the window holding the
/// pointer for the rest of the process (the state a pin is in between a
/// file step and the probe that answers for the file in it).
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
    forget_pin_park_swap();
    forget_video_frame();

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
    assert!(
        !pin_player_is_parked(),
        "and no cover stands over a film nothing is going to end"
    );
    assert!(
        !gesture_snapshot_active(),
        "and no dead interval is armed behind a press that armed nothing"
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
    forget_pin_park_swap();
    forget_video_frame();

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
    forget_pin_park_swap();
    take_gesture_snapshot();

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
    forget_pin_park_swap();
    take_gesture_snapshot();

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
    take_gesture_snapshot();

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
/// K3 RECONCILE: a swap under a hand reconciles what a file step inherits. A
/// snapshot nothing will ever take answers `video_drag_hold_apply` false for the
/// rest of the run, and the take-up's own `release_pin_capture` re-enters
/// `pin_capture_lost`, whose snapshot arm then relaunches the film the pin just
/// left - at its second, at its box - over the player begun for the file that
/// replaced it.
#[test]
fn a_swap_reconciles_the_gesture_it_is_taking_up_with() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));
    forget_pin_park_swap();
    assert!(
        gesture_press_freeze(),
        "a gesture killed a player of its own"
    );
    let generation = current_pinned_generation();
    fake_end_relaunch(102, 30.0, false);
    with_pin(|pin| pin.transport.drag_held = true);
    video_drag_hold_set(true);

    reconcile_swap_take_up(true);

    assert!(
        !gesture_snapshot_active(),
        "the snapshot goes with the step: no end will relaunch a file that is no longer on screen"
    );
    assert!(
        pending_pinned_relaunch().is_none(),
        "and the relaunch behind it is reaped rather than left for the next run's sweep"
    );
    assert!(
        !video_drag_holding(),
        "the claim on the film a gesture was holding goes with it: the next film to be dragged is \
         not answered with a claim about this one"
    );
    assert!(
        current_pinned_generation() == generation,
        "and no generation is bumped for a reconcile, which is not a gesture"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "and that relaunch is the outgoing film's own, so its kill is confirmed: nothing of the \
         film the pin left is left standing"
    );
    assert!(
        orphaned_video_kills().is_empty(),
        "reaped rather than parked for the next run's sweep"
    );

    video_drag_hold_set(false);
    restore(previous_pin, previous_pid, previous_media);
}

/// K3 RECONCILE: a cover standing for the file coming in is carried onto the pin
/// that file is taken up in, because the settle is what takes it down.
#[test]
fn a_swap_carries_a_cover_forward_to_the_incoming_file() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    assert!(cover_step_swap_for_video(true), "the step covers");

    reconcile_swap_take_up(true);

    assert!(
        pin_park_carried_forward(),
        "the cover survives the reconcile, so the take-up's own pin is written parked"
    );
    assert!(
        pin_player_is_parked(),
        "and the standing cover is still standing"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K3 RECONCILE: a cover with nothing behind it is given up rather than carried,
/// so a pin taken up over a picture is never left holding a frozen film.
#[test]
fn a_swap_gives_up_a_cover_nothing_is_coming_to_be_handed() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    assert!(cover_step_swap_for_video(true), "a cover stands");

    reconcile_swap_take_up(false);

    assert!(
        !pin_player_is_parked(),
        "the flag is down: nothing is coming to be handed this band"
    );
    assert!(
        !pin_park_carried_forward(),
        "and no record is left for a settle to answer"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K4 COVER: a step onto a video parks the outgoing frame - the band holds a
/// picture, never the desktop, until the new player has a window of its own.
#[test]
fn a_step_onto_a_video_parks_the_outgoing_frame() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    assert!(cover_step_swap_for_video(true), "a step onto a video parks");
    assert!(pin_player_is_parked(), "the cover stands across the swap");
    assert!(
        PIN_PARK_SWAP
            .lock()
            .ok()
            .and_then(|held| held.as_ref().map(|swap| swap.replacing))
            .unwrap_or(false),
        "for the incoming player behind it"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "and the outgoing player is killed: the cover stands over nothing"
    );
    assert!(
        orphaned_video_kills().is_empty(),
        "a confirmed kill is reaped, never parked"
    );

    // Rapid steps coalesce: a second step extends the standing cover rather than
    // stacking a second park on the first, and the newest file wins.
    fake_end_relaunch(103, 0.0, false);
    assert!(
        cover_step_swap_for_video(true),
        "a rapid second step covers again"
    );
    assert!(pin_player_is_parked(), "still one cover");
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "and the film the first step began is killed by the second"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K4 COVER: a file that paints itself is not covered, and a cover standing over
/// one is given up: there is no player coming to be handed this band, so a
/// frozen film left standing over a picture is worse than the hole it hides.
#[test]
fn a_step_onto_a_picture_raises_no_cover() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));
    forget_pin_park_swap();
    forget_video_frame();

    assert!(
        !cover_step_swap_for_video(false),
        "a picture paints itself: nothing to cover and nothing to hand the cover to"
    );
    assert!(!pin_player_is_parked(), "so no cover stands");
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        101,
        "and the outgoing player is left to the take-down that always killed it"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K4 ARM: the cover is armed for a player that is going to be put up, and a film
/// nothing on this machine will play is not one. The two roads that depend on the
/// answer - the cover and the start - are given the same one fact, because a cover
/// armed for a player that is never begun is a frozen frame of the film being left
/// behind standing in the band for the life of the pin: nothing is behind the band
/// to hand it back to, so it never comes down, and `pin_media_is_alive` reads its
/// own record as a player still on its way.
#[test]
fn a_step_onto_a_film_no_player_can_play_raises_no_cover() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));
    forget_pin_park_swap();
    forget_video_frame();

    // The machine's own answer cannot be asked for here: the arm is the one a machine with
    // FFmpeg never takes, so both answers are stood in rather than one of them read off
    // whatever machine this is running on.
    stand_ffplay_in(Some(true));
    assert!(
        pinned_player_is_coming(MediaType::Video),
        "a film on a machine with FFmpeg is a player coming"
    );
    stand_ffplay_in(Some(false));
    assert!(
        !pinned_player_is_coming(MediaType::Video),
        "and a film on a machine without one is not: there is no window to put it in"
    );
    assert!(
        !pinned_player_is_coming(MediaType::StaticImage),
        "nor is a file this app paints itself, which is the other half of the same answer"
    );

    assert!(
        !cover_step_swap_for_video(pinned_player_is_coming(MediaType::Video)),
        "so the step covers nothing"
    );
    assert!(
        !pin_player_is_parked(),
        "and no cover stands over the film it is leaving"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        101,
        "and the outgoing player is left to the take-down that always killed it"
    );

    // A cover already standing — a drag's, a seek's — is given up rather than carried onto a
    // pin whose player is never going to be begun, which is the arm the take-up reads this
    // through and the one the sibling refusal already takes for a start that failed. Raised on a
    // machine that has FFmpeg, because that is the machine a standing cover can only be left by.
    stand_ffplay_in(Some(true));
    assert!(
        cover_step_swap_for_video(pinned_player_is_coming(MediaType::Video)),
        "a cover stands over the outgoing film"
    );
    assert!(
        pin_player_is_parked(),
        "and the band is holding a picture, not the desktop"
    );

    stand_ffplay_in(Some(false));
    reconcile_swap_take_up(pinned_player_is_coming(MediaType::Video));
    assert!(
        !pin_player_is_parked(),
        "and it is given up where no player is coming: nothing is going to be handed this band"
    );
    assert!(
        !pin_park_carried_forward(),
        "no record is left for a settle to answer"
    );
    assert!(
        !pin_media_is_alive(false),
        "so the pin is a pin onto nothing rather than one waiting for a player that is never \
         coming: the cover's record is not the player's, and reading it as one kept the outgoing \
         frame on screen for the life of the pin"
    );

    stand_ffplay_in(None);
    restore(previous_pin, previous_pid, previous_media);
}

/// K5 ROAD: which of the two roads a relayout was asked on decides a video's
/// answer, and the answer is not the same for both. A box change is a player being
/// ended and begun again, so it presses; a take-up has pressed nothing, so it
/// relaunches the film that is there and reads the hold out of the pin's own
/// transport.
#[test]
fn a_take_up_relaunches_a_film_without_pressing_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();
    clear_restart_count();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        101,
        "a player is behind the band"
    );
    assert!(
        !pinned_is_held(),
        "and the film is playing rather than held"
    );

    relayout_pinned_media(
        &PathBuf::from("ws-k-chrome-nav.mkv"),
        content,
        96,
        None,
        PinRelayoutRoad::TakeUp,
    );

    assert_eq!(
        restart_count(),
        1,
        "one relaunch: the film is begun again in the box it is being shown in"
    );
    assert!(
        !pin_player_is_parked(),
        "and no cover stands over it. A take-up has pressed nothing, so there is no dead interval \
         to cover - and a cover raised here would be carried onto the pin this take-up is building \
         (see `pin_park_carried_forward`), which is a frame held for the length of a wait nothing \
         is ever going to end"
    );
    assert!(
        !gesture_snapshot_active(),
        "and no snapshot armed behind a road with no press in it: a snapshot nothing takes answers \
         `video_drag_hold_apply` false for the rest of the run"
    );
    assert!(
        !pinned_is_held(),
        "the hold is the pin's own transport, read as it was before this WS: a film that was \
         playing is playing on, not frozen under a cover no hand asked for"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K5 ROAD, the other half: the box change is the one road that does press, and its
/// relaunch carries was-held rather than the frozen transport — which is what
/// keeps a maximized playing film playing instead of coming back paused.
#[test]
fn a_box_change_presses_the_film_and_relaunches_it_behind_the_cover() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();
    clear_restart_count();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    relayout_pinned_media(
        &PathBuf::from("ws-k-chrome-nav.mkv"),
        content,
        96,
        None,
        PinRelayoutRoad::BoxChange,
    );

    assert!(
        pin_player_is_parked(),
        "the cover stands over the outgoing frame: a replacement's window is on screen within \
         milliseconds and empty until the file is open"
    );
    assert_eq!(
        restart_count(),
        1,
        "exactly one relaunch, and it is the press's own end that makes it"
    );
    assert!(
        !pinned_is_held(),
        "and it is begun playing. The film was playing when the change began it, so was-held is \
         false and the relaunch carries playing — never a pause. The frozen transport says held \
         for every film a press has touched, playing or not, which is why the relaunch reads the \
         snapshot's was-held instead"
    );
    assert!(
        !gesture_snapshot_active(),
        "and the snapshot is spent: the press's end has taken the one relaunch it owed"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K4 COVER with nothing behind it is nothing: no pin up, or no player, and there
/// is no band to cover and no frame to read.
#[test]
fn a_step_with_nothing_playing_raises_no_cover() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));
    forget_pin_park_swap();
    assert!(
        !cover_step_swap_for_video(true),
        "with no player there is no window to put away and no frame to read"
    );

    VIDEO_PID.store(101, Ordering::SeqCst);
    stand_pin(None);
    forget_pin_park_swap();
    assert!(
        !cover_step_swap_for_video(true),
        "and with no pin up there is no band to cover"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K4 COVER hands the band over: the settle waits for the incoming player's
/// window rather than swapping onto whatever merely exists, and takes the cover
/// down in the same tick it puts that window up.
#[test]
fn a_step_hands_the_cover_to_the_window_that_replaces_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    assert!(cover_step_swap_for_video(true), "the step covers");

    let window = RecordedPinWindow::new(0x1000);
    fake_end_relaunch(103, 0.0, false);

    // A replacement's window is on screen within milliseconds of being begun and
    // has decoded nothing: the cover holds until the bound, not until the window.
    assert!(
        !settle_pinned_park_where(&window, true),
        "an empty window in the band is not a band to hand back"
    );
    assert!(pin_player_is_parked(), "so the cover stands");

    spend_the_swap_bound();
    assert!(
        settle_pinned_park_where(&window, true),
        "and the bound spent with one standing, the band goes back to it"
    );
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some(content)),
            PinWindowCall::Repaint,
        ],
        "placed before shown, repainted in the same tick"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K5 PRESS: a box change snapshots, parks the cover and kills - the WS-J press
/// half, at the box the window stands at rather than at a gesture's.
#[test]
fn a_box_change_freezes_parks_and_kills() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    let generation = current_pinned_generation();

    assert!(
        unsafe { begin_covered_box_change() },
        "a maximize over a live player takes the covered road"
    );
    assert!(
        gesture_snapshot_active(),
        "the press snapshots: the end relaunch owns the resume"
    );
    assert!(pin_player_is_parked(), "and the cover stands");
    assert!(
        PIN_PARK_SWAP
            .lock()
            .ok()
            .and_then(|held| held.as_ref().map(|swap| swap.replacing))
            .unwrap_or(false),
        "for a relaunch behind it"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        0,
        "the press kills: VIDEO_PID == 0 mid-road asserts silence"
    );
    assert!(
        current_pinned_generation() > generation,
        "a newer road supersedes whatever relaunch is still in flight"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K5 PRESS with nothing to kill is the legacy road: no snapshot, no cover.
#[test]
fn a_box_change_with_no_player_stays_on_the_legacy_road() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    stand_pin(Some(playing_pin((100, 80, 420, 320), 30.0)));
    forget_pin_park_swap();

    assert!(
        !unsafe { begin_covered_box_change() },
        "with VIDEO_PID == 0 there is nothing to kill: no snapshot, no cover"
    );
    assert!(!gesture_snapshot_active(), "and no dead interval is armed");
    assert!(!pin_player_is_parked(), "and no cover stands");

    restore(previous_pin, previous_pid, previous_media);
}

/// K5 PRESS is not run twice for one relayout: a resize's drag already ran it
/// and left its cover standing, and running it again would kill the replacement
/// the first one is waiting for.
#[test]
fn a_box_change_over_a_standing_cover_does_not_press_again() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    assert!(
        park_pinned_player(HWND(0x1000 as *mut _), (content.0, content.1), true),
        "a resize's drag parks its own cover"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        101,
        "and its press has not killed yet - the release owns that"
    );

    assert!(
        !unsafe { begin_covered_box_change() },
        "so a box change over it is the same relayout and presses nothing again"
    );
    assert_eq!(
        VIDEO_PID.load(Ordering::SeqCst),
        101,
        "leaving the player the resize's own press stood behind alone"
    );
    assert!(
        pin_player_is_parked(),
        "and the standing cover still standing"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K5 END takes the snapshot once: a second end relaunches nothing, so a
/// maximize is exactly one relaunch at the final box.
#[test]
fn a_box_change_takes_the_snapshot_exactly_once() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    assert!(
        unsafe { begin_covered_box_change() },
        "the press arms one relaunch"
    );

    let first = take_gesture_snapshot();
    assert!(first.is_some(), "the end owns the one relaunch");
    assert!(
        !gesture_snapshot_active(),
        "the take disarms the dead interval"
    );
    assert_eq!(
        take_gesture_snapshot(),
        None,
        "a second end relaunches nothing: one press, one player"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K5 WHOLE ROAD: press, kill, one covered swap at the final box - a playing film
/// is playing after, placed before shown, repainted in the same tick, and never
/// at the box the press found.
#[test]
fn a_box_change_swaps_once_at_the_final_box() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 420, 320);
    let final_box = (0, 0, 800, 600);
    stand_pin(Some(playing_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    assert!(
        unsafe { begin_covered_box_change() },
        "the press snapshots, parks and kills"
    );
    let snapshot = take_gesture_snapshot().expect("the end owns the relaunch");
    fake_end_relaunch(
        102,
        snapshot.seconds,
        gesture_end_holding(snapshot.was_playing, snapshot.was_held),
    );
    spend_the_swap_bound();

    let window = RecordedPinWindow::new(0x1000);
    // The maximize has already written the new box into the pin, so the cover was
    // painted at that box rather than the one the press found.
    with_pin(|pin| pin.content = final_box);
    assert!(
        settle_pinned_park_where(&window, true),
        "the current replacement swaps"
    );
    assert!(!pin_player_is_parked(), "and the cover comes down with it");
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some(final_box)),
            PinWindowCall::Repaint,
        ],
        "one swap at the final box: placed before shown, repainted in the same tick"
    );
    assert_eq!(VIDEO_PID.load(Ordering::SeqCst), 102, "exactly one player");
    assert!(
        !settle_pinned_park_where(&window, true),
        "a spent swap never swaps twice"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// K5 BOX MATH: a maximize fits the shape to the room and remembers the box;
/// restore gives the remembered box back and forgets it. The cover is scaled to
/// whichever box the pin stands at, so this is the box it is scaled to.
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
    let style = player_ex_style(0);

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
        player_ex_style(style),
        style,
        "re-asserting changes nothing"
    );
    assert_eq!(
        player_ex_style(0x0004_0000isize) & 0x0004_0000isize,
        0x0004_0000isize,
        "and a style the player was created with is not dropped"
    );
}

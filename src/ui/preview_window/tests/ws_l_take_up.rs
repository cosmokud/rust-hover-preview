use super::*;

// The one reader of the relayout slot, shared with the tests about what a drag's end asks for
// (see `pin_input` and `ws_k_chrome_nav`).
use super::pin_input::take_relayout_request;

// WS-L: what a swap takes up with — the state beside the pin rather than in it.
//
// The report this file began from was a pin left with a placeholder in the video area after a
// step had been interrupted by a press. The placeholder is gone with the park that painted it
// (see `ARCHITECTURE.md` §"The film road"), and what survives of the report is the half that
// was never about a cover at all: a box request belongs to the file that made it, and a click in
// the picture is not a click on the caption's buttons.

/// A playing pin, wide enough for its caption to carry the whole of its row — every question
/// about a caption button would otherwise be about the three a window has of its own.
fn wide_banded_video_pin(content: ScreenRegion, from: f64) -> PinnedPreview {
    PinnedPreview {
        path: PathBuf::from("ws-l-take-up.mkv"),
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

/// A video behind the pin, as a step finds it: the kind this app plays with FFmpeg, taken away
/// and handed back by the test that stands one.
fn stand_video_media() -> Option<MediaData> {
    stand_ffplay_in(Some(true));
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
    clear_pending_pinned_relaunch();
    clear_restart_count();
    take_relayout_request();
    take_pin_command();
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    restore_media(previous_media);
}

/// The relayout a finished resize asked for, read and left where it is — a take would answer the
/// later half of the question by having already spent the thing under test.
fn the_relayout_request() -> Option<ScreenRegion> {
    PIN_BOX_REQUEST.lock().ok().and_then(|request| *request)
}

/// A press in the middle of the media band: the handle every part of it is for carrying the
/// window, which is what the title bar's own pull is (see `pinned_press`).
fn a_media_point() -> (i32, i32) {
    let content = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.content))
        .expect("a pin is up");
    let (x, y) = ((content.0 + content.2) / 2, (content.1 + content.3) / 2);
    assert!(
        pin_state().is_some_and(|pinned| {
            pinned
                .pin()
                .is_some_and(|pin| pin.resize_edge(x, y).is_none())
        }),
        "the middle of the band is not a resize edge"
    );
    (x, y)
}

/// What a press left armed on the caption.
fn the_pressed_button() -> Option<pin_chrome::CaptionButton> {
    pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin.pressed))
}

/// A click in the picture is a click in the picture, and `Previous` and `Next` live in the strip
/// above the band: the pin's own caption height is the whole of what keeps the two apart, and the
/// band's least-coverage pixel is what makes the click the window's to carry.
#[test]
fn a_band_click_is_never_a_caption_button() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 820, 620);
    stand_pin(Some(wide_banded_video_pin(content, 30.0)));
    take_pin_command();

    let (x, y) = a_media_point();
    assert!(
        unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
        "a click in the picture is the window's to carry"
    );
    assert_eq!(
        the_pressed_button(),
        None,
        "and it arms no caption button: `Previous` and `Next` live in the strip above the band, \
         which is the whole of what keeps a click in the video area off them"
    );
    assert_eq!(
        take_pin_command(),
        None,
        "so no file step is asked for by a click in the picture"
    );
    with_pin(|pin| pin.dragging = None);

    restore(previous_pin, previous_pid, previous_media);
}

/// A box request belongs to the file that made it. A resize's end writes the box the hand
/// settled on, and a swap that takes up another file before the loop drains it replays that rect
/// as a `PinBox` against a film it was never measured for (see `reconcile_swap_take_up`).
#[test]
fn a_box_request_made_before_a_swap_is_not_replayed_after_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 820, 620);
    stand_pin(Some(wide_banded_video_pin(content, 30.0)));
    take_relayout_request();
    take_pin_command();
    clear_pending_pinned_relaunch();
    clear_restart_count();

    // A resize's end, armed rather than begun: a begun resize asks for the frame it hands back
    // and spawns a player to render it, which a test must not do (see `ws_k`).
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
    let window = a_window_at(content);
    assert!(
        finish_pin_drag(HWND(0x1000 as *mut _), &window),
        "the hand lets go of the corner"
    );
    assert_eq!(
        the_relayout_request(),
        Some(content),
        "so the resize asked for its media at the box it settled on"
    );

    // The file under the pin is stepped over before the loop's next turn, which is what
    // `install_pinned_media` runs for every file.
    reconcile_swap_take_up();

    assert_eq!(
        take_relayout_request(),
        None,
        "so the box request is gone with the file that made it: a `PinBox` replayed against the \
         file that replaced this one lays that file out at this one's rect"
    );

    restore(previous_pin, previous_pid, previous_media);
}

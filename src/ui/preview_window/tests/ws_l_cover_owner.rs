use super::*;

// The one reader of the relayout slot, shared with the tests about what a drag's end asks for
// (see `pin_input` and `ws_k_chrome_nav`).
use super::pin_input::take_relayout_request;

// WS-L: a gesture may not mutate a cover it does not own.
//
// The report: a fresh pin over a video is correct, a step to the next file is correct, and then
// the FIRST thing the hand does afterwards — the title bar, the volume button, play/pause, a
// resize — leaves a placeholder in the video area, a click in that area answering as the
// next/previous button, and a pin nothing can be done with. Image preview too.
//
// RED (pre-fix):
// * a title-bar drag and a volume press write `awaiting_relaunch` onto a record a file step
//   wrote and they did not, which turns the settle's `TimedOut` arm off and leaves the cover
//   standing with a window already up behind it;
// * a resize's end writes `PIN_BOX_REQUEST` and nothing ever drains it, so it replays as a
//   `PinBox` against the file that came after;
// * a cover whose player never publishes a window is never given up, at any bound;
// * a box change over somebody else's cover falls through to a bare `restart_pinned_player`
//   with no cover and no snapshot.
//
// Three of the tests below were already green before the fix, and they are here because the report
// names them. The play/pause arm never parks at all, so it has no record to cross-write; the seek
// arm's `awaiting_relaunch` is its *own* fresh-park write, which the extend path already left
// alone; and a click in the video area is kept off `Previous`/`Next` by the caption's own height
// test rather than by anything to do with a cover.

/// A playing pin, wide enough for its caption to carry the whole of its row — every question
/// about a caption button would otherwise be about the three a window has of its own.
fn wide_banded_video_pin(content: ScreenRegion, from: f64) -> PinnedPreview {
    PinnedPreview {
        path: PathBuf::from("ws-l-cover-owner.mkv"),
        content,
        bound: Some(400),
        restore: None,
        dpi: 96,
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
    take_relayout_request();
    take_pin_command();
    stand_pin(previous_pin);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    restore_media(previous_media);
}

fn spend_the_swap_bound() {
    stand_the_park_wait_at(PIN_PARK_SWAP_TIMEOUT * 4);
}

/// Put the stamp the settle reads back where a drag of a real length would have put it, which
/// is the only part of a park a test has to supply by hand (see `park_swap_arm_for_the_band`).
fn stand_the_park_wait_at(age: Duration) {
    if let Ok(mut held) = PIN_PARK_SWAP.lock() {
        if let Some(swap) = held.as_mut() {
            swap.since = Some(Instant::now() - age);
        }
    }
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

/// A point on a named part of the transport bar of the pin that is up, found by asking the bar
/// rather than by reproducing its arithmetic.
fn a_transport_part_point(part: pin_chrome::TransportPart) -> (i32, i32) {
    let bar = pinned_transport_geometry().expect("a banded pin has a bar");
    let y = bar.top + bar.height / 2;

    (0..bar.width)
        .find(|&x| {
            pin_chrome::transport_part_at(x, y - bar.top, bar.width, bar.height, bar.dpi, bar.live)
                == Some(part)
        })
        .map(|x| (x, y))
        .unwrap_or_else(|| panic!("this bar draws no {part:?} to press"))
}

/// What a press left armed on the caption.
fn the_pressed_button() -> Option<pin_chrome::CaptionButton> {
    pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin.pressed))
}

/// The user's exact sequence up to the interaction: a fresh pin over a video, then a step onto
/// the next file. The step raises the cover it hands to the incoming player, and the record it
/// writes is a `Step` — a replacement already on its way, begun by the swap itself, so nothing
/// about that record is waiting for a release to make one.
fn stand_a_swapped_pin(content: ScreenRegion) {
    stand_pin(Some(wide_banded_video_pin(content, 30.0)));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    take_gesture_snapshot();
    take_relayout_request();
    take_pin_command();
    clear_pending_pinned_relaunch();
    clear_park_swap_arm();
    clear_restart_count();

    VIDEO_PID.store(101, Ordering::SeqCst);
    assert!(
        cover_step_swap_for_video(true),
        "a step onto a video raises the cover it hands to the incoming player"
    );
    // A player standing behind the cover, published without arming a relaunch of its own: the
    // cover is waiting for a window to see through it, not for a process to begin. A process id
    // no machine hands out, so nothing here is ever a process to end.
    VIDEO_PID.store(u32::MAX, Ordering::SeqCst);

    assert!(pin_player_is_parked(), "so the cover stands over the step");
    assert!(
        !seek_cover_is_waiting(),
        "and a step's record is not waiting for a release that has nothing to relaunch"
    );
}

/// What one interaction after that step must leave behind, at the state seam and one tick later.
///
/// Two facts. The first is where the cover is decided: a standing record left waiting on a road
/// that is not going to come can only be ended by `finish_pin_drag` or `pin_capture_lost`, and
/// both of them reach for `seek_cover_is_waiting` and then refuse. The second is the tick a real
/// loop reaches a moment after the hand let go — the window behind the cover up, the bound spent,
/// the gesture's own record gone — and it must hand the band back rather than hold a placeholder
/// over a player that is already standing there.
fn assert_the_pin_is_still_usable(content: ScreenRegion, what: &str) {
    assert!(
        !seek_cover_is_waiting(),
        "{what}: the standing record was left waiting on a release that was never going to make \
         one, so the only two roads that could end the cover both refuse it"
    );

    spend_the_swap_bound();
    let window = RecordedPinWindow::new(0x1000);
    assert!(
        settle_pinned_park_where(&window, true),
        "{what}: with the window that was already behind the cover up and the bound spent, the \
         cover settles"
    );
    assert!(
        !pin_player_is_parked(),
        "{what}: and the band is the hole the player's window is seen through again, rather than \
         an opaque rectangle whose pixels this app's own window claims every click in"
    );
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some(content)),
            PinWindowCall::Repaint
        ],
        "{what}: placed before shown, repainted in the same tick"
    );
}

/// L1 THE REPORT: a title-bar drag is the first thing a hand does after a step, and it used to
/// leave the pin bricked for the rest of its life.
#[test]
fn a_title_bar_drag_after_a_step_leaves_the_pin_usable() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 820, 620);
    stand_a_swapped_pin(content);

    // A title-bar pull resolves to this arm (`pinned_press`), begun through the window seam
    // because `begin_pin_drag` reads the box the window stands at off the machine (see `ws_k`).
    // The release is stood in rather than driven: it would begin a real player for the
    // snapshot, and what is under test is what the press left on the record.
    let window = a_window_at(content);
    begin_pin_drag(HWND(0x1000 as *mut _), &window, PinDragAction::Move, true);
    assert!(pin_is_dragging(), "so the drag began");
    with_pin(|pin| pin.dragging = None);

    assert_the_pin_is_still_usable(content, "a title-bar drag after a step");

    restore(previous_pin, previous_pid, previous_media);
}

/// L1 THE REPORT, second form: the volume button's own panel, which is the press that also
/// finds a standing record it did not write.
#[test]
fn a_volume_press_after_a_step_leaves_the_pin_usable() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 820, 620);
    stand_a_swapped_pin(content);

    // The popup is up, because the volume button was pressed first and this is the second press
    // of the same gesture: the panel is the thing the knob is taken hold of.
    with_pin(|pin| pin.volume.open = true);
    let popup = pinned_volume_geometry().expect("an open popup has a panel");
    let x = (popup.panel.left + popup.panel.right) / 2;
    let y = (popup.panel.top + popup.panel.bottom) / 2;
    assert!(
        unsafe { pinned_volume_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on the panel takes hold of the knob"
    );
    assert!(pin_volume_dragging(), "so the knob is held");

    assert_the_pin_is_still_usable(content, "a volume press after a step");

    restore(previous_pin, previous_pid, previous_media);
}

/// L1 THE REPORT, third form: the play/pause button on the bar.
#[test]
fn a_play_pause_press_after_a_step_leaves_the_pin_usable() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 820, 620);
    stand_a_swapped_pin(content);

    let (x, y) = a_transport_part_point(pin_chrome::TransportPart::Play);
    assert!(
        unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on play is the bar's to answer"
    );
    assert!(
        pin_state().is_some_and(|pinned| pinned
            .pin()
            .is_some_and(|pin| pin.transport.pressed.is_some())),
        "so the button is held"
    );

    assert_the_pin_is_still_usable(content, "a play/pause press after a step");

    restore(previous_pin, previous_pid, previous_media);
}

/// L1 THE REPORT, fourth form: a press on the seek track.
#[test]
fn a_seek_press_after_a_step_leaves_the_pin_usable() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 820, 620);
    stand_a_swapped_pin(content);

    let (x, y) = a_transport_part_point(pin_chrome::TransportPart::Seek);
    assert!(
        unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) },
        "a hand on the track puts the playhead under itself"
    );
    assert!(
        pin_state().is_some_and(|pinned| pinned
            .pin()
            .is_some_and(|pin| pin.transport.seeking.is_some())),
        "so the scrub has an aim"
    );

    assert_the_pin_is_still_usable(content, "a seek press after a step");

    restore(previous_pin, previous_pid, previous_media);
}

/// L2 OWNERSHIP, said on its own: a record a file step wrote is that step's, and a press that
/// finds it standing has no business saying a relaunch is still to come behind it — the press
/// did not begin one, and nothing about the record it found has changed.
#[test]
fn a_standing_step_record_is_not_flipped_by_a_press_that_did_not_write_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 820, 620);
    stand_a_swapped_pin(content);

    let window = a_window_at(content);
    begin_pin_drag(HWND(0x1000 as *mut _), &window, PinDragAction::Move, true);
    assert!(
        pin_player_is_parked(),
        "a press over a standing cover extends it rather than raising a second one"
    );
    assert!(
        !seek_cover_is_waiting(),
        "and it leaves the record it found exactly as it found it: a step's record is not made \
         into a cover waiting on a release by a press that began no relaunch"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// L3 THE OTHER HALF OF THE REPORT: a click in the video area is a click in the picture, and
/// `Previous` and `Next` live in the strip above the band. Checked with a cover standing — which
/// is what makes this app's own window claim that area at all — and then again once the cover is
/// down and the band is the hole it is between films.
#[test]
fn a_band_click_is_never_a_caption_button() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 820, 620);
    stand_a_swapped_pin(content);

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

    spend_the_swap_bound();
    let window = RecordedPinWindow::new(0x1000);
    assert!(
        settle_pinned_park_where(&window, true),
        "and the step's cover settles on its own deadline"
    );
    assert!(
        !pin_player_is_parked(),
        "so the band's pixels are the transparent ones a player's window is hit-tested through, \
         rather than an opaque rectangle this app's own window answers every click in"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// L4 AGGRAVATOR: a box request belongs to the file that made it. A resize's end writes the box
/// the hand settled on, and a swap that takes up another file before the loop drains it replays
/// that rect as a `PinBox` against a film it was never measured for.
#[test]
fn a_box_request_made_before_a_swap_is_not_replayed_after_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 820, 620);
    stand_a_swapped_pin(content);

    // A resize's end, armed rather than begun: a begun resize asks for the frame it hands the
    // band back and spawns a player to render it, which a test must not do (see `ws_k`).
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
    // `install_pinned_media` runs for every file but a video the media engine plays.
    reconcile_swap_take_up(true);

    assert_eq!(
        take_relayout_request(),
        None,
        "so the box request is gone with the file that made it: a `PinBox` replayed against the \
         file that replaced this one lays that file out at this one's rect"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// L5 THE INVARIANT: no cover stands for ever. A player that never publishes a window has
/// nothing for the band to be handed to, and the placeholder held over it goes on claiming every
/// click in the video area for as long as the pin stands — so the `no window at all` case has a
/// deadline of its own and gives the cover up when it goes by.
#[test]
fn a_cover_with_no_window_ever_arriving_comes_down_by_its_own_deadline() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 820, 620);
    stand_a_swapped_pin(content);
    stand_video_frame_for_a_test([0u8, 0, 0, 255].repeat(64), 8, 8);

    stand_the_park_wait_at(Duration::ZERO);
    let window = RecordedPinWindow::new(0x1000);
    assert!(
        !settle_pinned_park_where(&window, false),
        "a bound is a bound on a wait, not a delay: nothing has been waited for yet"
    );
    assert!(pin_player_is_parked(), "so the cover is still standing");

    stand_the_park_wait_at(PIN_PARK_COVER_TIMEOUT + PIN_PARK_SWAP_TIMEOUT);
    assert!(
        settle_pinned_park_where(&window, false),
        "and with no window ever arriving, the cover comes down by its own deadline"
    );
    assert!(
        !pin_player_is_parked(),
        "so no sequence of user input can leave the pin with a stranded opaque cover"
    );
    assert!(
        window.calls().contains(&PinWindowCall::Repaint),
        "and the band is painted again as the hole it is between films"
    );
    assert!(
        matches!(park_swap_last_arm(), Some((ParkSwap::Abandoned, _))),
        "and the arm is written down with the wait, which is where the next person tuning the \
         cover's bound reads it from"
    );

    restore(previous_pin, previous_pid, previous_media);
}

/// L6 AGGRAVATOR: a box change over somebody else's cover is not a bare relaunch. Refusing the
/// press leaves `relayout_pinned_media` with a `restart_pinned_player` and nothing else — a
/// player ended and another begun with the cover that was hiding the hole still standing and no
/// frame asked for, which is a hole in the shape of a video rather than a placeholder.
#[test]
fn a_box_change_over_a_step_cover_is_a_press_and_not_a_bare_relaunch() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    clear_park_swap_arm();
    let previous_pid = VIDEO_PID.swap(101, Ordering::SeqCst);
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_video_media();

    let content = (100, 80, 820, 620);
    stand_a_swapped_pin(content);
    let generation = current_pinned_generation();

    assert!(
        unsafe { begin_covered_box_change() },
        "a box change over a standing cover is a press under the cover that is already there, \
         not a refusal that leaves the relayout with nothing"
    );
    assert!(
        gesture_snapshot_active(),
        "so the press snapshots and the end relaunch owns the resume"
    );
    assert!(
        pin_player_is_parked(),
        "and the cover the relayout is about to leave a hole behind is still standing"
    );
    assert!(
        current_pinned_generation() > generation,
        "a newer road supersedes whatever relaunch is still in flight behind it"
    );

    restore(previous_pin, previous_pid, previous_media);
}

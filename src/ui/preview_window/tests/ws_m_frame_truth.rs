use super::pin_input::{take_relayout_request, CapturingWindow};
use super::*;
use std::collections::BTreeMap;

// WS-M: the frame truth. A pinned window is shown another file, and the state it
// is left in is not the state the box-change road leaves it in — so the first
// thing a hand does with it (a title-bar drag, a press of the volume button, a
// play/pause, an edge) lands on a window whose own geometry and the geometry it
// draws disagree, and a maximize followed by a restore — two box changes and
// nothing else — puts it right.
//
// The oracle is that round trip, so this file is written as a measurement
// first: one dump of every piece of pin and window state, taken at three
// points (a first take-up, a swap's take-up, and that swap followed by a
// maximize and a restore), and the fields that differ between the second and
// the third but not between the first and the second are the defect.

/// A point the dump reads nothing clock-like from: a transport's own start is
/// an `Instant`, so it is reported as "running"/"none" rather than as a value
/// that differs between every two readings.
fn playing(transport: &PinTransport) -> &'static str {
    match transport.started {
        Some(_) => "running",
        None => "none",
    }
}

/// A level or a second, to the precision two runs can be compared at.
fn rounded(value: Option<f64>) -> String {
    value
        .map(|value| format!("{:.3}", value))
        .unwrap_or_else(|| "none".to_string())
}

fn box_of(box_: ScreenRegion) -> String {
    format!("{:?}", box_)
}

/// Every piece of pin and window state that could differ between one take-up
/// and the next, as a comparable map.
///
/// It is a map rather than a struct so that a diff between two of them is a
/// list of *names*, which is the whole of what a reader of a failing test wants:
/// the field, not a wall of two structs to compare by eye.
fn pin_state_dump() -> BTreeMap<String, String> {
    let mut out = BTreeMap::new();

    // The pin's own lock is let go of before any of the statics is read: some of
    // them answer of the pin themselves (`pin_volume_open`), and a std lock is
    // not re-entrant.
    {
        let pinned = pin_state();
        let pin = pinned.as_ref().and_then(|pinned| pinned.pin());
        let Some(pin) = pin else {
            out.insert("pin".to_string(), "down".to_string());
            return out;
        };

        out.insert("pin.path".into(), format!("{:?}", pin.path));
        out.insert("pin.content".into(), box_of(pin.content));
        out.insert("pin.window_box".into(), box_of(pin.window_box()));
        out.insert("pin.window_size".into(), format!("{:?}", pin.window_size()));
        out.insert("pin.bound".into(), format!("{:?}", pin.bound));
        out.insert(
            "pin.restore".into(),
            box_of(pin.restore.unwrap_or((-1, -1, -1, -1))),
        );
        out.insert("pin.dpi".into(), pin.dpi.to_string());
        out.insert("pin.frame".into(), format!("{:?}", pin.frame));
        out.insert("pin.caption".into(), pin.caption.to_string());
        out.insert("pin.transport_bar".into(), pin.transport_bar.to_string());
        out.insert("pin.transport_live".into(), pin.transport_live.to_string());
        out.insert("pin.overlay".into(), pin.overlay.to_string());
        out.insert("pin.hides_chrome".into(), pin.hides_chrome.to_string());
        out.insert("pin.chrome.caption".into(), pin.chrome.caption.to_string());
        out.insert("pin.chrome.bar".into(), pin.chrome.bar.to_string());
        out.insert("pin.chrome.until".into(), format!("{:?}", pin.chrome.until));
        out.insert("pin.collapsed".into(), pin.collapsed.to_string());
        out.insert(
            "pin.bubble_pause".into(),
            format!("{:?}", pin.bubble_pause.is_some()),
        );
        out.insert("pin.hovered".into(), format!("{:?}", pin.hovered));
        out.insert("pin.pressed".into(), format!("{:?}", pin.pressed));
        out.insert(
            "pin.dragging".into(),
            pin.dragging
                .as_ref()
                .map(|drag| match drag.action {
                    PinDragAction::Move => "move".to_string(),
                    PinDragAction::Resize(_) => "resize".to_string(),
                })
                .unwrap_or_else(|| "none".to_string()),
        );
        out.insert("pin.parked".into(), pin.parked.to_string());
        out.insert("pin.tooltip.app".into(), pin.tooltip.default_app.clone());
        out.insert(
            "pin.tooltip.button".into(),
            format!("{:?}", pin.tooltip.button),
        );
        out.insert(
            "pin.audio_hovered".into(),
            format!("{:?}", pin.audio_hovered),
        );
        out.insert(
            "pin.audio_pressed".into(),
            format!("{:?}", pin.audio_pressed),
        );

        out.insert("pin.volume.level".into(), pin.volume.level.to_string());
        out.insert("pin.volume.audio".into(), pin.volume.audio.to_string());
        out.insert("pin.volume.open".into(), pin.volume.open.to_string());
        out.insert(
            "pin.volume.dragging".into(),
            pin.volume.dragging.to_string(),
        );
        out.insert(
            "pin.volume.playing_at".into(),
            pin.volume.playing_at.to_string(),
        );

        out.insert(
            "pin.transport.duration".into(),
            rounded(pin.transport.duration),
        );
        out.insert(
            "pin.transport.started".into(),
            playing(&pin.transport).to_string(),
        );
        out.insert(
            "pin.transport.paused_at".into(),
            rounded(pin.transport.paused_at),
        );
        out.insert(
            "pin.transport.pending_hold".into(),
            pin.transport.pending_hold.to_string(),
        );
        out.insert(
            "pin.transport.drag_held".into(),
            pin.transport.drag_held.to_string(),
        );
        out.insert(
            "pin.transport.seeking".into(),
            rounded(pin.transport.seeking),
        );
        out.insert(
            "pin.transport.hovered".into(),
            format!("{:?}", pin.transport.hovered),
        );
        out.insert(
            "pin.transport.pressed".into(),
            format!("{:?}", pin.transport.pressed),
        );
        out.insert(
            "pin.transport.subtitle".into(),
            format!("{:?}", pin.transport.subtitle),
        );
    }

    out.insert(
        "static.PIN_PARK_SWAP".into(),
        format!("{:?}", park_record_owner()),
    );
    out.insert(
        "static.PIN_PENDING_RELAUNCH".into(),
        format!(
            "{:?}",
            pending_pinned_relaunch().map(|pending| (pending.pid, pending.generation))
        ),
    );
    out.insert(
        "static.PIN_PLAYER_GENERATION".into(),
        current_pinned_generation().to_string(),
    );
    out.insert(
        "static.GESTURE_SNAPSHOT".into(),
        gesture_snapshot_active().to_string(),
    );
    out.insert(
        "static.VIDEO_DRAG_HOLDING".into(),
        video_drag_holding().to_string(),
    );
    out.insert(
        "static.VIDEO_PID".into(),
        VIDEO_PID.load(Ordering::SeqCst).to_string(),
    );
    out.insert(
        "static.VIDEO_HWND".into(),
        VIDEO_HWND.load(Ordering::SeqCst).to_string(),
    );
    out.insert("static.HELD_VIDEO_FRAME".into(), {
        let held = held_video_frame();
        held.is_some().to_string()
    });
    out.insert(
        "static.PIN_ARC".into(),
        PIN_ARC.load(Ordering::SeqCst).to_string(),
    );
    out.insert("static.PIN_BOX_REQUEST".into(), {
        let request = PIN_BOX_REQUEST
            .lock()
            .unwrap_or_else(|held| held.into_inner());
        format!("{:?}", *request)
    });
    out.insert(
        "static.CURRENT_MEDIA".into(),
        format!("{:?}", current_media_type()),
    );

    out
}

/// The fields two dumps disagree about, as `name: this | that`.
fn difference(left: &BTreeMap<String, String>, right: &BTreeMap<String, String>) -> Vec<String> {
    left.iter()
        .filter_map(|(name, value)| match right.get(name) {
            Some(other) if other != value => Some(format!("{name}: {value} | {other}")),
            _ => None,
        })
        .collect()
}

/// A video behind the pin, as the take-up reads the kind off the slot.
fn stand_a_video(width: u32, height: u32) -> Option<MediaData> {
    stand_a_kind(width, height, MediaType::Video)
}

/// A picture behind the pin: a kind whose window is its own media, so the box
/// the take-up installs and the box the window stands at are one box.
fn stand_a_picture() -> Option<MediaData> {
    stand_a_kind(640, 480, MediaType::StaticImage)
}

fn stand_a_kind(width: u32, height: u32, kind: MediaType) -> Option<MediaData> {
    let previous = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let mut media = create_loading_media(width, height);
    media.media_type = kind;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }
    previous
}

/// Every piece of a road's bookkeeping put back to nothing, so one test's
/// reads are not the next test's.
fn clean_slate() {
    take_gesture_snapshot();
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    clear_pending_pinned_relaunch();
    clear_park_swap_arm();
    clear_restart_count();
    clear_park_trace();
    stand_pin(None);
    video_drag_hold_set(false);
    pin_arc_set(None);
    VIDEO_PID.store(0, Ordering::SeqCst);
    let _ = take_pin_command();
}

/// The pin taken up for the first file, at a box on the display.
///
/// It is wide on purpose: the walk's four buttons are dropped whole from a
/// caption too narrow to carry them, and every question about a caption button
/// would then be about the three a window keeps for itself.
fn take_up_a_video_pin() {
    let first = PathBuf::from("ws-m-frame-truth-first.mkv");
    install(take_up_pinned_window(&first, (100, 80, 800, 320)));
}

/// The pin shown another file — what a caption's **Next** does — with the
/// swap's own cover armed over the outgoing frame and carried onto the pin the
/// incoming file is taken up in.
fn step_the_pin_to_another_file() -> PathBuf {
    let (space, bounds, dpi) = pin_state()
        .and_then(|pinned| {
            let pin = pinned.pin()?;
            let space = pin_swap_space(pin);
            let dpi = pin.dpi;
            Some((space, work_area_at(pin.content.0, pin.content.1), dpi))
        })
        .expect("a pin to swap away from");
    let content = pin_update_box(
        pin_swap_room(space, bounds, dpi),
        (1280, 720),
        PreviewScale::FitToScreen,
    );

    // The outgoing player, published by name rather than begun: a test must not
    // spawn one, and every question here is about the state a swap leaves
    // rather than about the process it leaves behind.
    VIDEO_PID.store(101, Ordering::SeqCst);
    forget_video_frame();
    assert!(
        cover_step_swap_for_video(true),
        "a step onto a film covers the outgoing frame"
    );
    // The incoming player, published by name rather than begun: a test must not
    // spawn one, and every question here is about the state a swap leaves
    // rather than about the process it leaves behind.
    VIDEO_PID.store(102, Ordering::SeqCst);
    reconcile_swap_take_up(true);

    let mut video = create_loading_media(1280, 720);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }

    let next = PathBuf::from("ws-m-frame-truth-second.mkv");
    install(take_up_pinned_window(&next, content));
    next
}

/// The box the pin's window is standing at, which is what the pin's own state
/// says it is standing at — the two have to agree, and the test that says so is
/// the one about the window a hand is aiming at.
fn the_window_box() -> ScreenRegion {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_box()))
        .expect("a pin up")
}

/// A point on a named button of the caption above the pin that is up, found by
/// asking the caption rather than by reproducing its arithmetic — and filtered
/// down to a point that is not also on a resize edge, because the frame is asked
/// about before the caption is.
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

/// A point in the middle of the media band of the pin that is up, in the
/// window's own coordinates: a hand on the picture, which is a handle for
/// carrying the window and nothing else.
fn a_point_in_the_media() -> (i32, i32) {
    let (width, height) = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_size()))
        .expect("a pin up");
    let caption = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.caption))
        .unwrap_or(0);

    (width / 2, caption + (height - caption) / 2)
}

/// The button the caption's press armed, read out of the pin.
fn the_pressed_button() -> Option<pin_chrome::CaptionButton> {
    pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin.pressed))
}

#[test]
fn a_dump_of_three_take_ups_is_comparable() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_media = stand_a_video(640, 360);
    clean_slate();

    // **A** — the first take-up, for the first file.
    take_up_a_video_pin();
    let a = pin_state_dump();

    // **B** — the same window shown another file.
    VIDEO_PID.store(101, Ordering::SeqCst);
    step_the_pin_to_another_file();
    let b = pin_state_dump();

    // **C** — the same window after a maximize and a restore: the two box
    // changes the user's own hands found, and nothing else. The relayout a box
    // change asks for is left out on purpose — it begins a player, and a test
    // must not spawn one — so what is measured here is the half of the road
    // that writes the pin's own state.
    let mut request = None;
    toggle_pin_maximized(&mut request);
    assert!(
        matches!(request.take(), Some(PreviewMessage::PinBox(_))),
        "a maximize asks for the box it gave the window"
    );
    toggle_pin_maximized(&mut request);
    assert!(
        matches!(request.take(), Some(PreviewMessage::PinBox(_))),
        "and the restore asks for the one it put aside"
    );
    let c = pin_state_dump();

    let ab = difference(&a, &b);
    let bc = difference(&b, &c);

    println!("--- A vs B ---\n{ab:#?}");
    println!("--- B vs C ---\n{bc:#?}");
    println!("--- A (first take-up) ---\n{a:#?}");
    println!(
        "--- repaired by the box change (B vs C but not A vs B) ---\n{:#?}",
        bc.iter()
            .filter(|line| !ab.contains(line))
            .collect::<Vec<_>>()
    );

    stand_pin(None);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
}

/// The user's own sequence: a video pinned, the next file, and then a hand on
/// the window. Every one of those four steps has to leave a window that answers.
///
/// The step this file is about is the last of them, and it is the one that
/// breaks for good: a title-bar drag on a window whose own geometry and the
/// geometry it draws have come apart moves a window whose caption is somewhere
/// else, and every press after it is answered by the wrong arm.
#[test]
fn a_video_pin_is_still_usable_after_a_title_bar_drag_that_follows_a_swap() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_a_video(640, 360);
    clean_slate();

    // 1 — a pin up on a video, and a caption that answers.
    take_up_a_video_pin();
    let window = CapturingWindow::around(a_window_at(the_window_box()));
    let (x, y) = a_caption_button_point(pin_chrome::CaptionButton::Next);
    assert!(
        unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
        "a button on a pin that has just come up is a button"
    );
    assert_eq!(
        the_pressed_button(),
        Some(pin_chrome::CaptionButton::Next),
        "so the press armed it"
    );
    assert!(
        unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) },
        "and its release is the caption's own"
    );
    assert_eq!(
        take_pin_command(),
        Some(PinCommand::Next),
        "which walks the listing along"
    );

    // 2 — the next file, which is the swap.
    step_the_pin_to_another_file();
    let window = CapturingWindow::around(a_window_at(the_window_box()));
    let (x, y) = a_caption_button_point(pin_chrome::CaptionButton::Close);
    assert!(
        unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
        "and a button on a pin shown another file is still a button"
    );
    assert!(
        unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) },
        "with its own release"
    );
    assert_eq!(
        take_pin_command(),
        Some(PinCommand::Close),
        "which asks for its own command"
    );

    // 3 — the title bar, dragged. The rest of the caption is a handle for
    // carrying the window, so this is a move and nothing else.
    let window = CapturingWindow::around(a_window_at(the_window_box()));
    let caption = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.caption))
        .unwrap_or(0);
    begin_pin_drag(HWND(0x1000 as *mut _), &window, PinDragAction::Move, true);
    assert!(
        pin_is_dragging(),
        "the hand is carrying the window, whatever the swap left behind"
    );
    assert!(
        finish_pin_drag(HWND(0x1000 as *mut _), &window),
        "and letting go of it ends the drag"
    );

    // 4 — and the window is still a window a hand can use. The picture is the
    // handle for carrying it, and a press there is the media's own: nothing is
    // armed by it, because the caption's buttons are somewhere else entirely.
    let (x, y) = a_point_in_the_media();
    assert!(
        unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
        "a press on the picture is the pin's to act on"
    );
    assert_eq!(
        the_pressed_button(),
        None,
        "and it arms no caption button: the picture is not the caption, whatever the swap left \
         behind (a press at ({x}, {y}) of a window whose caption is {caption} rows tall)"
    );

    let window = CapturingWindow::around(a_window_at(the_window_box()));
    let (x, y) = a_caption_button_point(pin_chrome::CaptionButton::Next);
    assert!(
        unsafe { pinned_press(HWND(0x1000 as *mut _), x, y) },
        "and a caption button is still a caption button"
    );
    assert_eq!(
        the_pressed_button(),
        Some(pin_chrome::CaptionButton::Next),
        "so the press arms it"
    );
    assert!(
        unsafe { pinned_release(HWND(0x1000 as *mut _), x, y, &window) },
        "with its own release"
    );
    assert_eq!(
        take_pin_command(),
        Some(PinCommand::Next),
        "which asks for its own command rather than another arm's"
    );

    stand_pin(previous_pin);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
}

/// A point on a named part of the transport bar of the pin that is up, found by
/// asking the bar rather than by reproducing its arithmetic.
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

/// The part the bar's press armed, read out of the pin.
fn the_pressed_part() -> Option<pin_chrome::TransportPart> {
    pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin.transport.pressed))
}

/// The second a scrub is aiming at, read out of the pin.
fn the_seek_aim() -> Option<f64> {
    pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin.transport.seeking))
}

/// A picture of a shape of its own, written where `media_dimensions` can read
/// it — which is what makes the swap's own measurement the one under test
/// rather than a piece of it stood in.
fn write_a_png_of_size(path: &Path, width: u32, height: u32) {
    image::RgbImage::new(width, height)
        .save(path)
        .expect("a written PNG");
}

/// A swapped pin answers the four other things a hand does with it: the level,
/// the play/pause, the edge, and the seek track. They are one test because they
/// are one report — a window that has stopped answering its caption has usually
/// stopped answering its bar for the same reason, and a fix that answered only
/// the caption would pass a test that asked about the caption.
#[test]
fn a_video_pin_is_still_usable_after_a_swap_that_is_followed_by_a_press_on_its_bar() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();
    let previous_media = stand_a_video(640, 360);
    clean_slate();

    take_up_a_video_pin();
    step_the_pin_to_another_file();

    // The level, which is a button on the bar and a popup over the picture.
    let window = CapturingWindow::around(a_window_at(the_window_box()));
    let (x, y) = a_transport_part_point(pin_chrome::TransportPart::Volume);
    assert!(
        unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) },
        "a press on the level is the bar's to act on after a swap"
    );
    assert_eq!(
        the_pressed_part(),
        Some(pin_chrome::TransportPart::Volume),
        "so the bar armed it"
    );
    assert!(
        unsafe { pinned_transport_release(HWND(0x1000 as *mut _), x, y, &window) },
        "and the release is the bar's own"
    );
    assert!(
        pin_volume_open(),
        "so the popup is up, which is what a click on the button does"
    );
    close_pin_volume();

    // The play/pause, which is the same button with a different job.
    let window = CapturingWindow::around(a_window_at(the_window_box()));
    let (x, y) = a_transport_part_point(pin_chrome::TransportPart::Play);
    assert!(
        unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) },
        "and a press on the play/pause is too"
    );
    assert_eq!(
        the_pressed_part(),
        Some(pin_chrome::TransportPart::Play),
        "so the bar armed that one"
    );
    assert!(
        unsafe { pinned_transport_release(HWND(0x1000 as *mut _), x, y, &window) },
        "with its own release"
    );

    // The seek track, which needs a length to aim at and so asks the pin's own
    // transport for one rather than waiting for a probe.
    with_pin(|pin| pin.transport.duration = Some(200.0));
    let bar = pinned_transport_geometry().expect("a banded pin has a bar");
    let (x, y) = (bar.width / 2, bar.top + bar.height / 2);
    assert!(
        unsafe { pinned_transport_press(HWND(0x1000 as *mut _), x, y) },
        "and a press on the track is the bar's own as well"
    );
    assert!(
        the_seek_aim().is_some(),
        "so the scrub has its aim, which is what the release takes the file to"
    );
    let window = CapturingWindow::around(a_window_at(the_window_box()));
    assert!(
        unsafe { pinned_transport_release(HWND(0x1000 as *mut _), x, y, &window) },
        "and its release ends the scrub"
    );

    // The edge, which is the one drag that ends in a relayout: the drag is stood
    // rather than begun, because a begun resize asks for the frame it hands the
    // band back and spawns a player to render it, which a test must not do. Its
    // end is the same arm either way (see `finish_pin_drag`) — as long as no
    // press's dead interval is standing, or the end is a relaunch rather than
    // the relayout this is about.
    take_gesture_snapshot();
    forget_resume_frame();
    take_relayout_request();
    let window = CapturingWindow::around(a_window_at(the_window_box()));
    let (width, height) = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_size()))
        .expect("a pin up");
    let (x, y) = (width - 1, height - 1);
    assert!(
        pin_state().is_some_and(|pinned| {
            pinned
                .pin()
                .is_some_and(|pin| pin.resize_edge(x, y).is_some())
        }),
        "and the corner of the window is on a resize edge"
    );
    // The box is read out *before* the guard rather than inside it: `with_pin`
    // holds the pin's own lock and the closure runs under it, so asking for
    // the window box here would be asking for that same lock a second time —
    // and a std lock is not re-entrant, so the test waits on itself forever.
    let box_ = the_window_box();
    with_pin(move |pin| {
        pin.dragging = Some(PinDrag {
            from: (x, y),
            window: box_,
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
    assert!(pin_is_dragging(), "so the resize is under way");
    assert!(
        finish_pin_drag(HWND(0x1000 as *mut _), &window),
        "and letting go of it ends the resize"
    );
    assert!(
        take_relayout_request().is_some(),
        "which asked for the media at the box the hand settled on"
    );

    stand_pin(previous_pin);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
}

/// The oracle, stated as the invariant it is: a swap leaves a window in the
/// state a box change would have left it in, so the round trip through
/// `toggle_pin_maximized` is a no-op rather than a repair.
///
/// This is the A/B/C measurement made into an assertion. `B vs C` is empty, and
/// `A vs B` is exactly the list below — which is what a swap is *supposed* to
/// change and nothing else. A field that moves on a swap and moves again on a
/// box change is the defect this file was written to find, so the list is
/// written out rather than left to a diff nobody reads.
#[test]
fn a_swap_leaves_the_state_a_box_change_leaves() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_media = stand_a_video(640, 360);
    clean_slate();

    take_up_a_video_pin();
    let a = pin_state_dump();

    VIDEO_PID.store(101, Ordering::SeqCst);
    step_the_pin_to_another_file();
    let b = pin_state_dump();

    let mut request = None;
    toggle_pin_maximized(&mut request);
    request.take();
    toggle_pin_maximized(&mut request);
    request.take();
    let c = pin_state_dump();

    assert_eq!(
        difference(&b, &c),
        Vec::<String>::new(),
        "a maximize and a restore after a swap change nothing a box change is asked to change"
    );
    assert_eq!(
        difference(&a, &b),
        vec![
            format!("pin.content: {} | {}", a["pin.content"], b["pin.content"]),
            format!("pin.parked: {} | {}", a["pin.parked"], b["pin.parked"]),
            format!("pin.path: {} | {}", a["pin.path"], b["pin.path"]),
            format!(
                "pin.window_box: {} | {}",
                a["pin.window_box"], b["pin.window_box"]
            ),
            format!(
                "pin.window_size: {} | {}",
                a["pin.window_size"], b["pin.window_size"]
            ),
            format!(
                "static.PIN_PARK_SWAP: {} | {}",
                a["static.PIN_PARK_SWAP"], b["static.PIN_PARK_SWAP"]
            ),
            format!(
                "static.PIN_PLAYER_GENERATION: {} | {}",
                a["static.PIN_PLAYER_GENERATION"], b["static.PIN_PLAYER_GENERATION"]
            ),
            format!(
                "static.VIDEO_PID: {} | {}",
                a["static.VIDEO_PID"], b["static.VIDEO_PID"]
            ),
        ],
        "and a swap changes exactly what a swap is: the file, the box that file is laid out in, and \
         the cover and the generation that go with a player being replaced"
    );

    stand_pin(None);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
}

/// The oracle, as an assertion: the box a swap measures its incoming file out
/// in is the box the take-up installs that file in.
///
/// This is the whole of what a swap's road has to get right and did not. The
/// swap begins the incoming player *before* the take-up runs, so the box it
/// hands over is the box the film is drawn in — and the take-up clamps the
/// window onto the display, so a box near the edge of a display came to be two
/// boxes: a transparent band with the player's window somewhere else inside the
/// window's own picture, a caption drawn over the top of the film, and every
/// press in the video answered by whichever window was under it.
///
/// A box change has no such gap and never had: it clamps, writes the clamped box
/// into the pin, and only then begins the replacement in it. That is why a
/// maximize and a restore put a swapped window right again.
///
/// It is asked of a picture rather than a film because a picture is measurable
/// without a probe: `pin_update_content` measures the file itself, so the road
/// here is the road the loop takes rather than a piece of it.
#[test]
fn the_box_a_swap_measures_is_the_box_the_take_up_installs() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_media = stand_a_picture();
    clean_slate();

    let folder = std::env::temp_dir().join("rhp-ws-m-frame-truth");
    let _ = std::fs::create_dir_all(&folder);
    let first = folder.join("ws-m-frame-truth-first.png");
    let next = folder.join("ws-m-frame-truth-next.png");
    write_a_png_of_size(&first, 700, 300);
    write_a_png_of_size(&next, 700, 700);

    // The display this runs on, and a pin standing at the bottom of it — which
    // is where a file of another shape overhangs, and where the window has to be
    // moved back on. A pin in the middle of a display never sees this, which is
    // why a swap that walks a listing looks right until one file is taller than
    // the last.
    let display = work_area_at(40, 40);
    let left = display.left + 40;
    let top = display.top;
    install(take_up_pinned_window(
        &first,
        (left, top, left + 700, top + 300),
    ));

    let (space, bounds, dpi) = pin_state()
        .and_then(|pinned| {
            let pin = pinned.pin()?;
            let space = pin_swap_space(pin);
            let dpi = pin.dpi;
            Some((space, work_area_at(pin.content.0, pin.content.1), dpi))
        })
        .expect("a pin to swap away from");

    let Some(PinBox::Measured(planned)) = pin_update_content(space, &next, bounds, dpi) else {
        panic!("a picture of a shape of its own is measured, not waited for");
    };
    let installed = take_up_pinned_window(&next, planned);

    assert_eq!(
        planned, installed.content,
        "the box a swap measures its file out in is the box the take-up installs it in: the swap \
         begins the incoming player in the first one, before the take-up has run"
    );

    let _ = std::fs::remove_file(&first);
    let _ = std::fs::remove_file(&next);
    stand_pin(None);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
}

/// The walk a maximized window takes through a sound's card, which is the one
/// kind a swap lays out at its own size rather than in the box the window is
/// standing in. The card is never a maximized window — there is no box for it
/// to fill, and its caption draws no maximize to press — but the maximize the
/// window was in when the card arrived is the state the file after the card is
/// laid out under: that file is fitted to the room the display has, which is
/// the box a maximized window's file is given (see `pin_update_content`, and
/// the take-up that installs what a swap planned, `take_up_pinned_window`,
/// which carries what belongs to the window rather than to the file it is
/// showing).
#[test]
fn a_maximize_a_card_steps_over_is_kept_for_the_file_after_the_card() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_media = stand_a_picture();
    clean_slate();

    let folder = std::env::temp_dir().join("rhp-ws-m-card-maximize");
    std::fs::create_dir_all(&folder).expect("a test folder");

    // The display this runs on, which is the room a maximized window's file
    // is laid out against.
    let display = work_area_at(40, 40);
    let bounds = ScreenBounds {
        left: display.left,
        top: display.top,
        right: display.right,
        bottom: display.bottom,
    };

    // The picture the pin is taken up on, at a box of its own shape — which
    // is where a pin gets the bound the files after it are fitted into.
    let first = folder.join("ws-m-card-maximize-first.png");
    let next = folder.join("ws-m-card-maximize-next.png");
    write_a_png_of_size(&first, 700, 300);
    write_a_png_of_size(&next, 4000, 3000);
    let left = bounds.left + 40;
    let top = bounds.top;
    stand_a_kind(700, 300, MediaType::StaticImage);
    install(take_up_pinned_window(
        &first,
        (left, top, left + 700, top + 300),
    ));

    // The window maximized: the room is the box, and the box it had is the
    // one a restore puts back.
    let mut request = None;
    toggle_pin_maximized(&mut request);
    assert!(
        matches!(request.take(), Some(PreviewMessage::PinBox(_))),
        "a maximize asks for the box it gave the window"
    );
    assert!(
        pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| pin.restore))
            .expect("a pin up")
            .is_some(),
        "a maximized window has a box to restore to"
    );

    // The sound the walk steps onto: a probe answered for it, and the box its
    // card came out at is held against the file the way the measure thread
    // holds it.
    let song = folder.join("ws-m-card-maximize-song.mp3");
    std::fs::write(
        &song,
        b"ID3\x04\x00\x00\x00\x00\x00\x00\x10\x00\x00\x00",
    )
    .expect("a written file");
    audio_track::remember(
        &song,
        audio_track::Probed::Track(audio_track::Track {
            player: audio_track::Player::Ffmpeg,
            codec: Some("MP3".to_string()),
            rate: Some(44_100),
            channels: Some(2),
            bitrate: Some(192_000),
            duration: Some(180.0),
        }),
    );
    let scope = {
        let options = current_audio_options();
        MeasureScope::Room {
            cap_width: (bounds.right - bounds.left).max(1) as u32,
            cap_height: bounds.height().max(1) as u32,
            dpi: 96,
            theme: options.theme,
            font_scale_percent: options.font_scale_percent,
        }
    };
    hold_box(&song, &file_version(&song), &scope, Some((240, 200)));

    // The step onto the card: the card is laid out at its own size, in the
    // middle of the box the pin has — no part of the room a maximized window
    // fills is the card's to take.
    let (space, bounds_at, dpi) = pin_state()
        .and_then(|pinned| {
            let pin = pinned.pin()?;
            let space = pin_swap_space(pin);
            let dpi = pin.dpi;
            Some((space, work_area_at(pin.content.0, pin.content.1), dpi))
        })
        .expect("a pin to swap away from");
    let Some(PinBox::Measured(card)) = pin_update_content(space, &song, bounds_at, dpi)
    else {
        panic!("a measured card is its own size, not a wait")
    };
    assert_eq!(
        (card.2 - card.0, card.3 - card.1),
        (240, 200),
        "a card is laid out at its own size even where the window is maximized"
    );

    // And the take-up of it: the card is what is on screen, at the box it was
    // planned at, and the maximize the window was in is what the take-up
    // carries — a state about the window, not about the file it is showing.
    stand_a_kind(240, 200, MediaType::Audio);
    install(take_up_pinned_window(&song, card));
    assert!(
        pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| pin.restore))
            .expect("a pin up")
            .is_some(),
        "a card is shown no maximize to stay in, but the window's own maximize is not \
         given up by the step onto it"
    );
    let card_window = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.content))
        .expect("a pin up");
    assert!(
        (card_window.2 - card_window.0) * 2 < bounds.right - bounds.left,
        "the card on screen is its own size, where the box a maximized window fills is \
         the room's"
    );

    // The step off the card onto a picture: the window is still the maximized
    // one the card stepped over, so the picture is fitted to the room the
    // display has — the box a maximized window's file is given — and not to
    // the bound the picture before the card left, which is a ceiling the room
    // is not.
    let (space, bounds_at, dpi) = pin_state()
        .and_then(|pinned| {
            let pin = pinned.pin()?;
            let space = pin_swap_space(pin);
            let dpi = pin.dpi;
            Some((space, work_area_at(pin.content.0, pin.content.1), dpi))
        })
        .expect("a pin to swap away from");
    let scale = effective_preview_scale(&next, current_hover_scales());
    let Some(PinBox::Measured(planned)) = pin_update_content(space, &next, bounds_at, dpi)
    else {
        panic!("a picture of a shape of its own is measured, not waited for")
    };
    assert_ne!(
        space.room.region(),
        pin_swap_room(space, bounds, dpi),
        "the bound the picture before the card left is a ceiling the room is not"
    );
    assert_eq!(
        planned,
        pin_update_box(space.room.region(), (4000, 3000), scale),
        "the file after the card is fitted to the room the display has, as the window \
         it follows was maximized"
    );

    // And the take-up of that picture keeps what the window is: maximized,
    // with the box to restore to it has had since before the card.
    stand_a_kind(4000, 3000, MediaType::StaticImage);
    install(take_up_pinned_window(&next, planned));
    assert!(
        pin_state()
            .and_then(|pinned| pinned.pin().map(|pin| pin.restore))
            .expect("a pin up")
            .is_some(),
        "a picture shown after a card is shown the maximized window the card stepped \
         over"
    );

    let _ = std::fs::remove_dir_all(&folder);
    stand_pin(None);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media;
    }
}

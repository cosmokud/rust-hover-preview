use super::*;

/// A press the Explorer hook publishes reaches the pin's own press handling, and a tick
/// that publishes no press leaves the count where it was.
///
/// This is the whole of what the count is for. The hook polls faster than this loop does,
/// so a press published as a bit would be overwritten by the next tick's "no press" before
/// this side ever looked, and a click that began and ended between two ticks here would
/// be a press nothing ever saw — which is the drag that would never begin, and, through
/// the press bit it was reading to find that out, the click in the Explorer listing behind
/// the pin that `Pin Mode → Update Preview` follows (see `left_button_down`).
#[test]
fn a_published_press_is_not_missed_by_a_slower_poll() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    // The two are one set for the whole process, so this test stands in with the others
    // that publish the machine's own state rather than racing them.
    let _stand_in = POINTER_STAND_IN.lock().expect("the pointer's own state");

    let before = pin_media_press_count();

    // A tick that finds nothing pressed leaves both the state and the count alone.
    publish_pin_media_press(false, false);
    assert!(!left_button_down(), "and the button is up");
    assert_eq!(
        pin_media_press_count(),
        before,
        "a tick with no press in it publishes no press"
    );

    // A press moves the count, which is what this side asks about, and leaves the button
    // down while the hand is on it.
    publish_pin_media_press(true, true);
    let pressed = pin_media_press_count();
    assert!(pressed > before, "a press is a count that moved");
    assert!(left_button_down(), "and the button is down under the hand");

    // The hand lets go and nothing else happens: the count stays where the press left it,
    // so the next poll knows that press has already been given rather than reading the
    // button being down as a second one.
    publish_pin_media_press(false, false);
    assert_eq!(pin_media_press_count(), pressed, "a release is not a press");

    // And a second press is a second move, which is what a second click has to be for the
    // drag to be begun twice.
    publish_pin_media_press(true, true);
    assert!(
        pin_media_press_count() > pressed,
        "each press moves the count again"
    );
}

/// The whole app path for one file, without Explorer: the preview loop is started,
/// the file is shown the way a hover shows it, and what happens next is reported.
/// Ignored, and driven by `RHP_APP_PROBE` —
/// `$env:RHP_APP_PROBE = "C:\art\clock.svg"; cargo test -- --ignored --nocapture app_hover_probe`
/// — for a document whose preview does not appear.
#[test]
#[ignore = "shows a preview window"]
fn app_hover_probe() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let Ok(path) = std::env::var("RHP_APP_PROBE") else {
        println!("set RHP_APP_PROBE to a path");
        return;
    };
    let path = PathBuf::from(path);

    println!(
        "engine available: {}",
        crate::engines::webview_preview::is_available()
    );
    println!(
        "document: {}",
        crate::engines::webview_preview::draws(&path)
    );
    // Which engine would play a video here, asked the way the app asks it: whether FFmpeg's
    // player is installed, which settles it on its own where it is, and — where it is not —
    // whether this file is one the media engine is asked about and can decode. Reported for
    // every file rather than only for a video, because a probe is run to find out what the
    // machine is doing. The machine's own half of it is reported by `video_engine_probe`.
    println!(
        "video: the media engine plays it = {}, and ffplay is installed = {}",
        media_engine_plays(&path),
        crate::formats::codecs::ffplay_available()
    );

    std::thread::spawn(run_preview_window);
    std::thread::sleep(Duration::from_millis(500));

    show_preview(&path, 200, 200, None);

    for step in 0..25 {
        std::thread::sleep(Duration::from_millis(200));

        let media = CURRENT_MEDIA.lock().ok().and_then(|media| {
            media.as_ref().map(|media| {
                (
                    media.current_width(),
                    media.current_height(),
                    media.frames.len(),
                    media.media_type.is_native_video(),
                    // Whether the placeholder a video preview is opened with has been
                    // replaced by a frame of it, which is the one thing a video that
                    // plays and a video that does not look different in.
                    media
                        .current_pixels()
                        .as_chunks::<4>()
                        .0
                        .iter()
                        .any(|pixel| pixel[..3] != [40, 40, 40]),
                )
            })
        });

        println!(
            "{:>5} ms: engine_showing={} (w, h, frames, native video, framed)={:?}",
            (step + 1) * 200,
            crate::engines::webview_preview::is_showing(),
            media
        );

        if crate::engines::webview_preview::is_showing() {
            break;
        }
    }

    // Kept up long enough for the screen to be looked at, for the probes that
    // measure what is on it rather than how long it took to get there.
    let hold = std::env::var("RHP_APP_PROBE_HOLD_MS")
        .ok()
        .and_then(|ms| ms.trim().parse().ok())
        .unwrap_or(0);
    if hold > 0 {
        std::thread::sleep(Duration::from_millis(hold));
    }

    hide_preview();
    std::thread::sleep(Duration::from_millis(300));
    println!(
        "after hide: engine_showing={}",
        crate::engines::webview_preview::is_showing()
    );
}

/// Put the pointer at a point on the screen, as a hand would: a real `SendInput` move,
/// because a drag is a sequence of moves delivered to whatever window is under the pointer
/// and nothing less is one.
///
/// Only the `#[ignore]`d drag probes use this, because a real move is the user's own pointer
/// teleported and left wherever the probe finished: a plain test takes its step with
/// `step_the_pin_drag_by` instead, which asks the machine nothing but where the cursor already is.
pub(super) fn send_pointer_to(x: i32, y: i32) {
    use windows::Win32::UI::WindowsAndMessaging::{
        GetSystemMetrics, SM_CXVIRTUALSCREEN, SM_CYVIRTUALSCREEN, SM_XVIRTUALSCREEN,
        SM_YVIRTUALSCREEN,
    };

    unsafe {
        let (left, top) = (
            GetSystemMetrics(SM_XVIRTUALSCREEN),
            GetSystemMetrics(SM_YVIRTUALSCREEN),
        );
        let (wide, high) = (
            GetSystemMetrics(SM_CXVIRTUALSCREEN),
            GetSystemMetrics(SM_CYVIRTUALSCREEN),
        );
        let scaled =
            |value: i32, from: i32, over: i32| ((value - from).max(0) * 65535) / (over - 1).max(1);

        send_inputs(&[mouse_input(
            scaled(x, left, wide),
            scaled(y, top, high),
            MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE,
        )]);
    }
}

/// Take a step of `(dx, dy)` with a drag that is going, without moving the pointer: a step is
/// `(where the cursor is) - (the drag's own anchor)`, and the cursor is the one end of that a
/// stand-in window cannot answer for — `apply_pin_drag` reads `GetCursorPos` itself
/// (see `PinWindow::pointer`). So the anchor is put a step's travel behind wherever the cursor
/// really is, the callers' own `apply_pin_drag` measures the step from there, and the hand's
/// pointer stays where the user's hand left it — which is what a plain `cargo test` owes the
/// desktop (see `a_border_stroke` in the edges' own tests).
pub(super) fn step_the_pin_drag_by(dx: i32, dy: i32) {
    let at = cursor_screen_point().expect("the pointer is somewhere on the desktop");
    let mut anchored = false;
    with_pin(|pin| {
        if let Some(drag) = pin.dragging.as_mut() {
            drag.from = (at.0 - dx, at.1 - dy);
            // The guard reads the point the last step was measured at, and a hand that has not
            // moved would otherwise be read as the same point twice and the step dropped.
            drag.carried = (i32::MIN, i32::MIN);
            anchored = true;
        }
    });
    assert!(anchored, "a step is asked for with a drag going");
}

fn send_left_button(down: bool) {
    send_inputs(&[mouse_input(
        0,
        0,
        if down {
            MOUSEEVENTF_LEFTDOWN
        } else {
            MOUSEEVENTF_LEFTUP
        },
    )]);
}

/// The pin key, as a hand would press it — the low-level hook counts real key-downs and this
/// is the only way to give it one without a hand.
fn send_pin_key() {
    use windows::Win32::UI::Input::KeyboardAndMouse::{KEYBDINPUT, KEYBD_EVENT_FLAGS, VIRTUAL_KEY};

    fn key(vk: VIRTUAL_KEY, flags: KEYBD_EVENT_FLAGS) -> INPUT {
        INPUT {
            r#type: INPUT_TYPE(1),
            Anonymous: INPUT_0 {
                ki: KEYBDINPUT {
                    wVk: vk,
                    wScan: 0,
                    dwFlags: flags,
                    time: 0,
                    dwExtraInfo: 0,
                },
            },
        }
    }

    send_inputs(&[
        key(VIRTUAL_KEY(0x20), KEYBD_EVENT_FLAGS(0)),
        key(VIRTUAL_KEY(0x20), KEYBD_EVENT_FLAGS(0x0002)),
    ]);
}

fn mouse_input(dx: i32, dy: i32, flags: MOUSE_EVENT_FLAGS) -> INPUT {
    INPUT {
        r#type: INPUT_TYPE(0),
        Anonymous: INPUT_0 {
            mi: MOUSEINPUT {
                dx,
                dy,
                mouseData: 0,
                dwFlags: flags,
                time: 0,
                dwExtraInfo: 0,
            },
        },
    }
}

fn send_inputs(inputs: &[INPUT]) {
    unsafe {
        let sent = SendInput(inputs, std::mem::size_of::<INPUT>() as i32);
        assert_eq!(
            sent as usize,
            inputs.len(),
            "SendInput refused {sent} of {} — another program is holding the input desktop",
            inputs.len()
        );
    }
}

/// The document the drag probe drags: whatever `RHP_APP_PROBE` names, and a plain drawing
/// written out where it names nothing, so the probe can be run with nothing set up at all.
fn probe_document() -> PathBuf {
    let path = match std::env::var("RHP_APP_PROBE") {
        Ok(path) => PathBuf::from(path),
        Err(_) => std::env::temp_dir().join("rhp-pin-drag-probe.svg"),
    };
    if !path.exists() {
        std::fs::write(
            &path,
            b"<svg xmlns='http://www.w3.org/2000/svg' width='200' height='200'>\
                  <rect width='200' height='200' fill='#2b6'/><circle cx='100' cy='100' r='60' fill='#fff'/>\
                  </svg>",
        )
        .expect("a written document");
    }
    path
}

/// Stand a document up in a pin the way a hover and then the pin key would, and hand back the
/// box its drawing is in — which is the place a hand that means to carry the window puts
/// itself, since a press anywhere on the media of a pin means to move it (`pinned_engine_press_action`).
fn pin_a_document(path: &std::path::Path) -> ScreenRegion {
    std::thread::spawn(run_preview_window);
    std::thread::sleep(Duration::from_millis(500));

    show_preview(path, 200, 200, None);

    let mut landed = false;
    for _ in 0..100 {
        std::thread::sleep(Duration::from_millis(100));
        if crate::engines::webview_preview::showing_path().as_deref() == Some(path) {
            landed = true;
            break;
        }
    }
    assert!(landed, "the document never reached the browser");

    crate::shell::key_input::spawn_key_watcher();
    crate::shell::key_input::refresh();
    std::thread::sleep(Duration::from_millis(200));
    send_pin_key();

    let mut up = false;
    for _ in 0..50 {
        std::thread::sleep(Duration::from_millis(100));
        if pinned() {
            up = true;
            break;
        }
    }
    assert!(up, "the pin never came up");

    let Some(content) = pinned_content() else {
        panic!("a pinned window with no media box");
    };
    assert!(
        engine_window_is_at((content.0 + content.2) / 2, (content.1 + content.3) / 2),
        "the engine's own window is not where the drawing is",
    );

    content
}

/// The whole of a pinned drawing's fault, end to end and without a hand: a document the
/// engine draws is taken up in a pin, a press is made over the drawing and carried across
/// it at the hand's own rate, and the pin is then asked to answer a press on its own caption.
///
/// The last question is the one that matters and the one nothing else can answer. A drag of
/// the drawing is *supposed* to move the window, so a window that moved says nothing about
/// whether it can be moved again; a caption button that answers afterwards says the window is
/// still a window. The two symptoms are what a hand reports — the drawing comes away in the
/// cursor, the window follows it for a while, and afterwards nothing on the window answers.
///
/// The middle question is the one before it, and it is a question of rate rather than of
/// arrival: the press is carried on, but carried by the loop rather than by messages, so the
/// window follows the hand at whatever pace the loop turns at (`drag_places_per_second`). Both
/// halves are in this one probe because both are the same fault read two ways — a drawing whose
/// press is taken by somebody else's window — and because the engine is a per-process
/// singleton, so a second probe standing up its own document could not have got this far.
///
/// The press over the drawing is published rather than waited for. What the Explorer hook
/// does is count presses and publish the button's state (`publish_pin_media_press`), and the
/// loop answers that publication rather than a message, because the engine's window is
/// somebody else's and a message of this window's never arrives. A probe with no Explorer
/// behind it publishes what the hook would have published, and everything downstream of that
/// is the production path.
///
/// Ignored, and driven by `RHP_APP_PROBE` —
/// `$env:RHP_APP_PROBE = "C:\art\clock.svg"; cargo test -- --ignored --nocapture pin_drawing_drag_probe`
#[test]
#[ignore = "drives the pointer over a pinned drawing"]
fn pin_drawing_drag_probe() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let path = probe_document();
    let content = pin_a_document(&path);
    println!("drawing at {content:?}");

    let before = pinned_window_box().map(|(window, ..)| window);
    let latched = |()| pin_state().and_then(|state| state.pin().map(|pin| pin.dragging.is_some()));

    // The drag: a press on the drawing, carried across it a step at a time, and let go of.
    let (cx, cy) = ((content.0 + content.2) / 2, (content.1 + content.3) / 2);
    send_pointer_to(cx, cy);
    std::thread::sleep(Duration::from_millis(120));
    send_left_button(true);
    publish_pin_media_press(true, true);
    std::thread::sleep(Duration::from_millis(120));
    for step in 1..=24 {
        send_pointer_to(cx - step * 8, cy - step * 4);
        publish_pin_media_press(true, false);
        std::thread::sleep(Duration::from_millis(25));
    }
    send_left_button(false);
    publish_pin_media_press(false, false);
    std::thread::sleep(Duration::from_millis(400));

    let after = pinned_window_box().map(|(window, ..)| window);
    println!(
        "after the drag: pinned={}, window {before:?} -> {after:?}, drag still latched={:?}",
        pinned(),
        latched(()),
    );

    // A drag of the drawing is what a hand on a drawing means, so the window following it is
    // not the bug — it is the feature, and the feature is what a fix that simply refused the
    // drag would have thrown away with it. Asserted before the window is asked anything else,
    // because a window that has followed a hand is a window that was listening.
    assert!(
        before != after,
        "a drag of the drawing did not move the window"
    );

    // And now the question: is this still a window?
    assert!(
        pinned(),
        "the pin did not survive a drag of its own drawing — it was taken down rather than left"
    );

    // The second drag is the other fault, and the one a hand complains about first. That the
    // window moved at all is not in question any more; how it moved is. A drawing is carried by
    // the loop rather than by messages delivered to the window, so the rate it follows the hand
    // at is the rate the loop turns at — and the loop turns every `STATIC_PIN_WAIT_MS` for a
    // pinned static document, which is a window thrown after the pointer rather than carried by
    // it. Asked of the drawing's *current* place, since the drag above moved it.
    let Some(content) = pinned_content() else {
        panic!("a pinned window with no media box");
    };
    let rate = drag_places_per_second(((content.0 + content.2) / 2, (content.1 + content.3) / 2));
    // The hand was moved faster than any display turns, so a window that kept up with it was
    // put at a new place at very nearly the pointer's own rate. A window carried at the loop's
    // pace instead is put at one place per tick, however fast the hand is going.
    assert!(
        rate >= 100.0,
        "a pinned drawing followed the pointer at only {rate:.0} places a second"
    );

    let close = close_button_on_screen();
    send_pointer_to(close.0, close.1);
    std::thread::sleep(Duration::from_millis(120));
    // A button that is not under the pointer after the pointer was put there is a button the
    // window cannot be asked about, and that is the same answer however it was arrived at -
    // so it is said rather than left to be the close click's own confusing failure.
    assert_eq!(
        cursor_screen_point(),
        Some(close),
        "the close button is not where the window was asked to put it",
    );
    send_left_button(true);
    std::thread::sleep(Duration::from_millis(120));
    send_left_button(false);

    let mut closed = false;
    for _ in 0..30 {
        std::thread::sleep(Duration::from_millis(100));
        if !pinned() {
            closed = true;
            break;
        }
    }
    assert!(
        closed,
        "a drag of the drawing left a window whose caption answers nothing"
    );

    hide_preview();
}

/// How many places a second the pinned window was actually put at while a hand carried it from
/// a point, which is the rate a hand reads as smooth and which nothing else here can report.
///
/// Every other kind of pinned window has its drag carried by `WM_MOUSEMOVE` messages
/// delivered to it, and the wait at the end of the loop wakes the moment one is queued, so it
/// follows at the pointer's own rate whatever that rate is. A drawing's press is taken by the
/// engine's window, so its drag is carried by the loop instead, and a loop carries a drag at
/// the loop's own pace — one place per tick, and a tick of a pinned static document is
/// `STATIC_PIN_WAIT_MS` away. The count is therefore a direct reading of the fault, and of
/// nothing else.
///
/// So what is counted is distinct places the window stood at, watched from a thread of its own
/// so that what is measured is the window rather than the loop that is meant to be moving it.
/// The sampler polls far faster than the app can move the window, so every place the app put
/// the window in is one the sampler sees; and the pointer is stepped faster than any display
/// turns, so what the count is bounded by is the loop and not the hand.
fn drag_places_per_second(from: (i32, i32)) -> f64 {
    const STEPS: i32 = 120;
    const STEP_EVERY: Duration = Duration::from_millis(4);

    let (cx, cy) = from;
    send_pointer_to(cx, cy);
    std::thread::sleep(Duration::from_millis(120));
    send_left_button(true);
    publish_pin_media_press(true, true);
    std::thread::sleep(Duration::from_millis(120));

    let watching = std::sync::Arc::new(AtomicBool::new(true));
    let places = std::sync::Arc::new(Mutex::new(0usize));
    let watcher = {
        let watching = watching.clone();
        let places = places.clone();
        let hwnd = PREVIEW_HWND.load(Ordering::SeqCst);
        std::thread::spawn(move || {
            let hwnd = HWND(hwnd as *mut _);
            let mut last = None;
            while watching.load(Ordering::Acquire) {
                if let Some((left, top, ..)) = window_origin(hwnd) {
                    if last != Some((left, top)) {
                        last = Some((left, top));
                        *places.lock().expect("the places seen") += 1;
                    }
                }
                std::thread::sleep(Duration::from_micros(500));
            }
        })
    };

    let began = Instant::now();
    for step in 1..=STEPS {
        // To the right and down, where the pin has room to go: a hand that walks a window off
        // the edge of the screen stops moving for reasons that have nothing to do with the pace
        // being measured.
        send_pointer_to(cx + step * 3, cy + step * 2);
        publish_pin_media_press(true, false);
        std::thread::sleep(STEP_EVERY);
    }
    let took = began.elapsed();
    send_left_button(false);
    publish_pin_media_press(false, false);
    watching.store(false, Ordering::Release);
    watcher.join().expect("the watcher");

    let places = places.lock().expect("the places seen");
    let rate = *places as f64 / took.as_secs_f64();
    println!(
        "a {took:?} drag of {STEPS} steps put the window at {} places — {rate:.0}/s, with the \
             pointer itself moving at {:.0}/s",
        *places,
        STEPS as f64 / took.as_secs_f64(),
    );
    rate
}

/// Where the caption's close button is on the screen right now, found by asking the chrome
/// that draws it rather than by a metric that would be right only on this display, and
/// placed against where the window *is* rather than where the pin's state last said it was:
/// a click is delivered to a real window, and to nothing else.
fn close_button_on_screen() -> (i32, i32) {
    let hwnd = HWND(PREVIEW_HWND.load(Ordering::SeqCst) as *mut _);
    let Some((left, top, width, _)) = window_origin(hwnd) else {
        panic!("no pinned window to find a button on");
    };
    let Some(caption) = pinned_caption_geometry() else {
        panic!("a pinned window with no caption");
    };
    let y = caption.height / 2;
    let closes: Vec<i32> = (0..width)
        .filter(|x| {
            pin_chrome::button_at(*x, y, caption.width, caption.height, caption.dpi, false)
                == Some(pin_chrome::CaptionButton::Close)
        })
        .collect();
    let Some(span) = closes.first().zip(closes.last()) else {
        panic!("no close button on this caption")
    };

    // The middle of the button rather than its first pixel: the band along the top of a
    // framed window is a resize, and it is asked about before the caption is (see
    // `pinned_press`), so a press on the button's outermost column is a resize rather than a
    // click on it.
    (left + (span.0 + span.1) / 2, top + y)
}

/// A video replaced in place is probed again rather than cropped and sized by
/// the answer about the file it used to be: the version is part of what a probed
/// geometry is held for.
#[test]
fn keys_a_probed_geometry_by_the_files_version() {
    // A folder of this module's own: the tests run beside each other, and one
    // of them clearing its fixtures must not take another's with it.
    let folder = std::env::temp_dir()
        .join("rust-hover-preview-video-tests")
        .join("geometry");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("keyed.mp4");
    std::fs::write(&path, b"one").expect("a written file");

    let key = |path: &PathBuf| VideoGeometryKey {
        path: path.clone(),
        version: file_version(path),
    };

    let first = key(&path);
    assert!(first == key(&path), "the same file is the same key");

    std::fs::write(&path, b"a longer file").expect("a rewritten file");
    assert!(first != key(&path), "a rewritten file is another key");

    let _ = std::fs::remove_file(&path);
}

/// The report an `ffprobe` pass writes is read for the sound in it and nothing of the
/// picture before it: a film with a soundtrack is not a sound, and the fields a card is
/// drawn with are the ones the sound's own stream carries. A stream's codec name is
/// written ahead of the `codec_type` line that says what the stream is, which is the
/// order an `ffprobe` pass writes its fields in.
#[test]
fn reads_a_sound_out_of_an_ffprobe_report() {
    let report = "codec_name=h264\ncodec_type=video\nwidth=1920\n\
                      codec_name=flac\ncodec_type=audio\nsample_rate=44100\nchannels=2\n\
                      bit_rate=1006000\nduration=562.31\n";

    let track = audio_track_from_report(report).expect("a sound in the report");
    assert_eq!(track.player, Player::Ffmpeg);
    assert_eq!(track.codec.as_deref(), Some("FLAC"));
    assert_eq!(track.rate, Some(44_100));
    assert_eq!(track.channels, Some(2));
    assert_eq!(track.bitrate, Some(1_006_000));
    assert_eq!(track.duration, Some(562.31));

    assert_eq!(
        audio_track_from_report("codec_name=h264\ncodec_type=video\nduration=10.0\n"),
        None,
        "a file with no sound stream in it is not a sound, however it is named"
    );
}

/// What this machine has for the sounds named in `RHP_AUDIO_PROBE`, and what their cards
/// come out as.
///
/// Ignored by default, like every other probe here: it reads real files, it starts no
/// player and it plays nothing — what it asks is the question a hover asks before a card is
/// drawn, and what it prints is the answer beside the size the card was painted at.
///
/// ```text
/// $env:RHP_AUDIO_PROBE = "C:\music\track.flac;C:\music\podcast.opus"
/// cargo test audio_probe -- --ignored --nocapture
/// ```
#[test]
#[ignore = "reads the files named in RHP_AUDIO_PROBE"]
fn audio_probe() {
    let Ok(list) = std::env::var("RHP_AUDIO_PROBE") else {
        println!("set RHP_AUDIO_PROBE to one or more paths, separated by ';'");
        return;
    };

    for path in list
        .split(';')
        .map(str::trim)
        .filter(|path| !path.is_empty())
        .map(PathBuf::from)
    {
        println!("\n--- {} ---", path.display());
        println!(
            "the lists call it a sound: {}, and the preview is shown: {}",
            named_as(&path, PreviewType::Audio),
            drawn_as_audio(&path)
        );

        let started = Instant::now();
        let probed = probe_audio_track(&path);
        println!("probe: {probed:?} ({} ms)", started.elapsed().as_millis());

        // A container of a video's name is the case the whole verdict exists for: what the
        // video probe finds in it, and what the router says once that probe has answered.
        let named_video = named_as(&path, PreviewType::Videos);
        if named_video {
            println!(
                "the video probe answered {}",
                match probe_video_geometry(&path) {
                    ProbedGeometry::Measured(_) => "a shape",
                    ProbedGeometry::Unmeasurable => "nothing to measure",
                }
            );
            println!(
                "and the router calls it {:?}",
                CONFIG
                    .lock()
                    .ok()
                    .and_then(|config| crate::formats::routing::kind_of(&path, &config))
            );
            println!("its card is drawn: {}", drawn_as_audio(&path));
        }

        let Some(track) = probed else {
            println!("nothing here plays this file");
            continue;
        };

        println!(
            "the card says: {:?}",
            audio_preview::facts_of(&track, &path)
        );
        for (elapsed, duration) in [
            (None, None),
            (Some(67.0), track.duration),
            (Some(0.5), None),
        ] {
            let Some(card) = audio_card(&path, elapsed, duration, 0, None) else {
                continue;
            };
            let options = current_audio_options();
            let (width, height) =
                audio_preview::measure(&card, 4096, 2160, 96, options).expect("a measured card");
            let painted =
                audio_preview::render(&card, width, height, 96, options).expect("a painted card");

            println!(
                "card at {elapsed:?} / {duration:?}: {width}x{height}, {} bytes of frame",
                painted.0.len()
            );
        }

        // And what the engine actually does with the file, which is the question the probe
        // beside it does not answer: a probe says this machine has a decoder, and what a
        // hover needs is a player that gets somewhere. The two came apart once — a name the
        // engine resolved as a URL was refused after `Play` had already answered, so a file
        // probed as playable drew a card whose clock never moved (see `Session::begin`) —
        // and what tells that apart from a file that plays is the position below: a session
        // that failed reports a state of not playing and a position stuck at nothing, and
        // one that is playing reports both moving. Only the engine is asked: FFmpeg's
        // player is a process of its own and nothing on this side is drawn from it.
        if track.player == Player::Native {
            let volume = current_audio_volume().max(1);
            let seek = current_audio_seek();
            let start = audio_seek::start_position(&path, seek, track.duration);
            println!("starting at {start:.3}s, by `Volume → Audio Seek`");

            video_player::play_audio(&path, volume, start, true);
            std::thread::sleep(Duration::from_millis(500));

            println!(
                "playing natively: {}, at {:?} of {:?} ({})",
                video_player::is_playing(),
                video_player::position(),
                video_player::duration(),
                if video_player::playing_path().as_deref() == Some(path.as_path()) {
                    "the file that was asked for"
                } else {
                    "not the file that was asked for"
                },
            );
            video_player::stop();
        }
    }
}

/// The player's window, stood in for the one mpv would be drawing in: a real window of this
/// app's own, because the box a player's window is carried to is only ever asked of the
/// machine — `ensure_pinned_sibling_box` moves a window this app does not own and keeps no
/// record of the box it asked for anywhere a test can read.
///
/// A `STATIC` popup of the machine's own class, so nothing has to be registered; it is never
/// shown by the stand-in itself, and the placement that is being watched is the only thing that
/// can put it on screen.
unsafe fn stand_a_player_window(box_: ScreenRegion) -> HWND {
    use windows::core::w;

    // The pin's own creation style, which is what puts a window in the topmost band on a
    // machine that is refusing raises: the shell answers `HWND_TOPMOST` in `SetWindowPos`
    // only to a window that owns the foreground, and a stand-in raised that way — a plain
    // popup — is left in the ordinary band with the bit set but the order refused, so two
    // stand-ins never stand in any order at all. Born topmost, both are in the band from the
    // first placement on, the way the pin and the film it stands over are (see
    // `show_pinned_window`).
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
    .expect("a stand-in window for the player")
}

/// Where a window stands on the screen right now, read back off the machine — the very box
/// `ensure_video_window_topmost` asks for with its one `SetWindowPos`.
fn window_box_on_screen(hwnd: HWND) -> ScreenRegion {
    let (left, top, width, height) = window_origin(hwnd).expect("a window standing somewhere");
    (left, top, left + width, top + height)
}

/// A hand carrying an mpv'd pin by its chrome: the player's own window goes with the band on
/// every step, because nothing else stands in the band for it — mpv is left drawing, nothing is
/// parked, and the band over it is a hole.
///
/// **What it proves.** `apply_pin_drag` moves the pin and then asks `ensure_pinned_sibling_box`
/// to put whatever stands in the media band at the band's new box; for a film mpv is drawing,
/// that window *is* the picture, and a step that does not carry it leaves the band's hole over
/// bare desktop — the video goes, the desktop shows behind it, and it comes back when the drag
/// ends. The player's window is stood in by a real window of this app's own, and its box is read
/// back off the machine, because the place a player's window is put at is asked of the machine
/// and nothing else in this view keeps it. The record of which player's window that is says mpv,
/// which is the answer the placement is read off, and `VIDEO_PID` is left zero with it so the style
/// monitor is not asked to find a window of a player's process and cannot take the stand-in's
/// handle away (see `note_player_process`, `try_apply_noactivate_style`).
///
/// **Why it is a plain test.** It spawns no player and takes no player's window: the box it
/// carries is a stand-in window of this app's own, and which player that window belongs to is
/// stood in, so the whole of the loop is arithmetic, a pointer move and one `SetWindowPos` per
/// step. What it does touch is the machine's own pointer, because a drag is measured from where
/// the cursor is (`send_pointer_to`), which is why it takes the pointer's own lock:
///
/// `cargo test --lib a_dragged_mpv_pin_carries_its_players_window_on_every_move_step`
///
/// **What a red result discriminates.** It is red if the player's window is left behind while the
/// band moves, which is the first hypothesis: `ensure_video_window_topmost` refusing every
/// placement while the pin's chrome is up, with no arm that keeps a window mpv is drawing moving
/// with the band. It is green under the hide-the-player hypotheses — a park, or an `SW_HIDE` the
/// mpv gate missed — because a window that is hidden still carries the box this asserts.
#[test]
fn a_dragged_mpv_pin_carries_its_players_window_on_every_move_step() {
    use windows::Win32::UI::WindowsAndMessaging::DestroyWindow;

    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _stand_in = POINTER_STAND_IN.lock().expect("the pointer's own state");

    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);
    let previous_image = video_process_image();
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let content = (100, 80, 420, 320);

    // The player's window, the one thing here that can be asked where the film is.
    let player = unsafe { stand_a_player_window(content) };
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);

    // And whose window it is, as the placement reads it: this app writes the player down as it
    // begins one, and nothing here begins one (see `note_player_process`).
    note_player_process(0, MPV_IMAGE_NAME);

    // The mpv route, chrome up: the press leaves the film playing and its window the one the
    // drag moves — nothing is parked, which is what puts the band's hole over the desktop.
    stand_ffplay_in(Some(false));
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = PathBuf::from("a-pin-whose-player-is-mpv.mkv");
    pin.transport.begun(30.0, true, false);
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }

    let hwnd = HWND(0x1000 as *mut _);
    let window = a_window_at(content);
    begin_pin_drag(hwnd, &window, PinDragAction::Move, true);

    assert!(
        pin_chrome_up(),
        "the caption the hand is on is up over the band"
    );
    assert!(
        !pin_player_is_parked(),
        "and mpv's window is the one the drag has to carry"
    );

    for step in 1..=3 {
        let (dx, dy) = (step * 20, step * 10);
        step_the_pin_drag_by(dx, dy);

        // The stand-in is re-asserted here rather than left standing across the pair, because the
        // record of a player's window is the whole process's and not this test's: the monitor
        // thread that follows `VIDEO_PID` forgets the window it names the moment it sees a player
        // end, and in a run of the whole suite the player it sees end is another test's (see
        // `ensure_noactivate_monitor`).
        VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);

        unsafe { apply_pin_drag(hwnd) };

        let band = pinned_content().expect("a pin with a media band");
        let carried = window_box_on_screen(player);
        assert_eq!(
            carried, band,
            "move {step}: the player's window is not where the band is — the film has been left \
             behind, and the band's hole is the desktop showing through"
        );
    }

    finish_pin_drag(hwnd, &window);

    unsafe { DestroyWindow(player).expect("the stand-in window let go of") };
    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
    note_player_process(previous_pid, previous_image);
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A resize of a pin whose film mpv draws leaves the player's window exactly at the band on every
/// step, and at the band a whole grow and a whole shrink end on.
///
/// **What it proves.** The band a resized pin shows is this app's own alpha-1 claim, and the film is
/// the window under it (see `fill_band_hit_testable`): a step that put the pin at its new box and
/// left the player's window at an old one leaves the band a hole over the desktop for the whole of
/// the gesture, and a window that lags the band by a step is that same hole drawn smaller. So the
/// promise is per step and not per drag, in both directions — the drag that grows and the drag that
/// shrinks expose and give up a strip on different sides of the box, and a step of either is the
/// same `SetWindowPos` (see `ensure_pinned_sibling_box`).
///
/// The player's window is a stand-in of this app's own (`stand_a_player_window`) and its box is read
/// back off the machine, because the place a player's window is put at is asked of the machine and
/// nothing else in this view keeps a record of it.
///
/// `cargo test --lib a_resized_mpv_pin_leaves_its_players_window_at_the_band_on_every_step`
#[test]
fn a_resized_mpv_pin_leaves_its_players_window_at_the_band_on_every_step() {
    use windows::Win32::UI::WindowsAndMessaging::DestroyWindow;

    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _stand_in = POINTER_STAND_IN.lock().expect("the pointer's own state");

    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);
    let previous_image = video_process_image();
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let content = (100, 80, 420, 320);
    let player = unsafe { stand_a_player_window(content) };
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);
    note_player_process(0, MPV_IMAGE_NAME);

    stand_ffplay_in(Some(false));
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = PathBuf::from("a-resized-pin-whose-player-is-mpv.mkv");
    pin.transport.begun(30.0, true, false);
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    if let Ok(mut request) = PIN_BOX_REQUEST.lock() {
        *request = None;
    }

    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }

    let hwnd = HWND(0x1000 as *mut _);
    let window = a_window_at(content);
    let corner = edge(false, false, true, true);

    // **The grow.** Four steps out from the box the press found, each one asked where the film is.
    begin_pin_drag(hwnd, &window, PinDragAction::Resize(corner), true);
    assert!(
        !pin_player_is_parked(),
        "mpv is left drawing through a resize: the window is moved and the film plays on"
    );
    for step in 1..=4 {
        step_the_pin_drag_by(step * 30, step * 24);
        // Re-asserted per step for the reason the move test gives: the monitor thread of the whole
        // process follows `VIDEO_PID`, and another test's player ending is a window forgotten.
        VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);
        unsafe { apply_pin_drag(hwnd) };

        let band = pinned_content().expect("a pin with a media band");
        assert_eq!(
            window_box_on_screen(player),
            band,
            "grow {step}: the player's window is not where the band is — the film has been left \
             behind and the band's hole is the desktop showing through"
        );
    }
    finish_pin_drag(hwnd, &window);
    let grown = pinned_content().expect("a pin with a media band");
    assert!(
        grown.2 - grown.0 > content.2 - content.0,
        "the grow really grew: {grown:?} against {content:?}"
    );
    assert_eq!(
        window_box_on_screen(player),
        grown,
        "and the drag ends with the window at the band it ended on"
    );

    // **The shrink**, from the grown box and back past where the whole drag began.
    begin_pin_drag(hwnd, &window, PinDragAction::Resize(corner), true);
    for step in 1..=4 {
        step_the_pin_drag_by(-step * 60, -step * 48);
        VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);
        unsafe { apply_pin_drag(hwnd) };

        let band = pinned_content().expect("a pin with a media band");
        assert_eq!(
            window_box_on_screen(player),
            band,
            "shrink {step}: the player's window is not where the band is — a window bigger than \
             the band is the film under the chrome it has just given up"
        );
    }
    finish_pin_drag(hwnd, &window);
    let shrunk = pinned_content().expect("a pin with a media band");
    assert!(
        shrunk.2 - shrunk.0 < content.2 - content.0,
        "the shrink really shrank, past the box the press found: {shrunk:?} against {content:?}"
    );
    assert_eq!(
        window_box_on_screen(player),
        shrunk,
        "and the shrink ends with the window at the band it ended on"
    );

    unsafe { DestroyWindow(player).expect("the stand-in window let go of") };
    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
    note_player_process(previous_pid, previous_image);
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// The window an mpv film is drawn in goes down with the pin and is put back in its band with it,
/// and it travels with the band on a step of a drag.
///
/// **What it proves.** A child of the pin's window is hidden, shown and moved by the compositor for
/// nothing; a window of the desktop's is the pin's to do all three with, and every road that does one
/// of them for the pin already exists for a film that is a window of its own (see
/// `hide_pinned_siblings` and `ensure_pinned_sibling_box`). So what this pins is that an mpv film on
/// the top-level road is still on all three lists — the failure it guards is a film left standing on
/// screen after the pin was put away, and a film left behind by a step.
///
/// `VIDEO_PID` is left zero, as the placement tests leave it: the roads are read off this app's own
/// record of the window, and no process of the machine's is enumerated or styled.
#[test]
fn an_mpv_pins_player_window_goes_down_with_the_pin_and_travels_with_its_band() {
    use windows::Win32::UI::WindowsAndMessaging::{DestroyWindow, IsWindowVisible};

    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);
    let previous_image = video_process_image();
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let band = (140, 120, 460, 360);

    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = PathBuf::from("a-pin-whose-player-is-mpv.mkv");
    stand_pin(Some(pin));
    note_player_process(0, MPV_IMAGE_NAME);

    let player = unsafe { stand_a_player_window(content) };
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);
    unsafe {
        let _ = ShowWindow(player, SW_SHOWNOACTIVATE);
    }

    // The pin is put away: the film goes with it rather than standing alone on the desktop.
    unsafe { hide_pinned_siblings() };
    assert!(
        !unsafe { IsWindowVisible(player) }.as_bool(),
        "a pin that goes down leaves no film of its own on the screen"
    );

    // And the band the pin comes back to is the band the film is put back in, which is the same
    // call a step of a drag makes: the place and the show are one `SetWindowPos`.
    ensure_pinned_sibling_box(band);
    assert!(
        unsafe { IsWindowVisible(player) }.as_bool(),
        "the band a pin comes back to is a hole until the film is put back in it"
    );
    assert_eq!(
        window_box_on_screen(player),
        band,
        "and it is put back where the band is, not where it stood when the pin went away"
    );

    unsafe { DestroyWindow(player).expect("the stand-in window let go of") };
    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
    note_player_process(previous_pid, previous_image);
    stand_pin(previous_pin);
}

/// A resize drag asks for the pin's band where it begins and where it ends, and never on a step
/// between the two.
///
/// **The fault this is a regression test for is a `SetWindowPos` per pointer move.** The band a pin
/// belongs in is a question about the setting rather than about its box (`band_place`), so a resize
/// that re-asserted it on every step was a DWM reorder per move for a window already in the band it
/// named. What a drag owes is the band at its press and at its release — so this reads both, one
/// direction for each setting, and reads the band a step leaves behind in between.
///
/// The window is a real one of this app's own (the same stand-in the player placements are read
/// off), because a band is a fact about a window and nothing in this view keeps a record of it: the
/// raised state is read back off the machine with `window_holds_the_topmost_band`.
#[test]
fn a_resize_drag_bands_the_pin_at_its_two_ends_and_not_on_a_step_between_them() {
    use windows::Win32::UI::WindowsAndMessaging::DestroyWindow;

    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _stand_in = POINTER_STAND_IN.lock().expect("the pointer's own state");

    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_band = PIN_BAND_HELD.swap(false, Ordering::SeqCst);
    let always_on_top = pin_always_on_top();

    let content = (100, 80, 420, 320);
    // A real window of the machine's own, born topmost — the state this test puts the pin in the
    // wrong band of is its own, so what is read after the press is this app's answer and not the
    // window's creation.
    let pin_window = unsafe { stand_a_player_window(content) };
    assert!(
        unsafe { window_holds_the_topmost_band(pin_window) },
        "the stand-in is born in the topmost band, so a window read as ordinary after the press \
         was put there by the drag"
    );

    // **`Always On Top` off: the press and the release are the band's, and a step's is not.** The
    // band is put back in the topmost one by hand mid-drag, which is any road that put the window in
    // the band `band_place` does not name: a step that re-asked the band would take it away again.
    let banded = |window: HWND| unsafe { window_holds_the_topmost_band(window) };
    let band_the_pin = |window: HWND, band: HWND| unsafe {
        let _ = SetWindowPos(
            window,
            band,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
        );
        let _ = DwmFlush();
    };

    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = false;
    }
    stand_pin(Some(overlay_pin(content, PinChrome::always())));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    let hwnd = pin_window;
    let window = a_window_at(content);
    begin_pin_drag(
        hwnd,
        &window,
        PinDragAction::Resize(edge(false, false, true, true)),
        true,
    );
    assert!(
        !banded(pin_window),
        "the press puts the pin in the band the setting names, which with `Always On Top` off is \
         the ordinary one"
    );

    band_the_pin(pin_window, HWND_TOPMOST);
    assert!(banded(pin_window), "the window is back in the wrong band");

    for step in 1..=3 {
        step_the_pin_drag_by(step * 12, step * 9);
        unsafe { apply_pin_drag(hwnd) };
        assert!(
            banded(pin_window),
            "step {step}: a step of a resize asked for the band again, which is a DWM reorder per \
             pointer move for a window already in it"
        );
    }

    finish_pin_drag(hwnd, &window);
    assert!(
        !banded(pin_window),
        "and the release is the other end the band is asked at: the pin goes back to the ordinary \
         band however a step left it"
    );

    // **`Always On Top` on: the same two ends, the other way.** The window is put in the ordinary
    // band first, so the topmost read at each end is the drag's answer rather than the band it began
    // in.
    if let Ok(mut request) = PIN_BOX_REQUEST.lock() {
        *request = None;
    }
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
    }
    stand_pin(Some(overlay_pin(content, PinChrome::always())));
    band_the_pin(pin_window, HWND_NOTOPMOST);
    assert!(
        !banded(pin_window),
        "the window is out of the band the setting names"
    );

    begin_pin_drag(
        hwnd,
        &window,
        PinDragAction::Resize(edge(false, false, true, true)),
        true,
    );
    assert!(
        banded(pin_window),
        "the press of a resize puts the pin in the topmost band where the setting asks for it"
    );
    step_the_pin_drag_by(8, 6);
    unsafe { apply_pin_drag(hwnd) };
    assert!(
        banded(pin_window),
        "and a step leaves the band it found alone"
    );
    finish_pin_drag(hwnd, &window);
    assert!(
        banded(pin_window),
        "and the release leaves the pin in the band the setting asks for"
    );

    if let Ok(mut request) = PIN_BOX_REQUEST.lock() {
        *request = None;
    }
    unsafe { DestroyWindow(pin_window).expect("the stand-in window let go of") };
    PIN_BAND_HELD.store(previous_band, Ordering::SeqCst);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
    }
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// Only a window mpv is drawing is moved with a dragged pin where the pin's chrome is up: every
/// other answer the placement can give is the one it gave before.
///
/// **Why it is asked of the placement rather than of a drag.** The move `ensure_pinned_sibling_box` makes is
/// one condition — whether the recorded player's window is mpv's (`video_process_image`) — and the
/// answers it is not are the ones the hold-off is for. FFmpeg's window is the film a park hides and
/// a strip of the pin's chrome stands over, so it is left where it is; and a band with no window of
/// a player's in it, which is every preview this app draws itself, has nothing to place at all. The
/// move travels without a raise, so the pin stays over the film it was standing over and no frame
/// puts the film over the chrome. A drag for each would be the same arithmetic three times, so the
/// drag the mpv arm gets is the one above and the rest are pinned here.
///
/// Both windows are stand-ins of this app's own (`stand_a_player_window`) and `VIDEO_PID` is left
/// zero, so no player's process is looked for and nothing of this machine's is styled.
#[test]
fn only_an_mpv_players_window_is_placed_with_the_pins_chrome_up() {
    use windows::Win32::UI::WindowsAndMessaging::DestroyWindow;

    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);
    let previous_image = video_process_image();
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let moved = (200, 180, 520, 420);

    // The pin the placement is asked about: a strip of its chrome is up over the band, which is
    // the whole of the hold-off the mpv film is an exception to.
    stand_pin(Some(overlay_pin(content, PinChrome::always())));

    let mpv_player = unsafe { stand_a_player_window(content) };
    let ffplay_player = unsafe { stand_a_player_window(content) };

    VIDEO_HWND.store(mpv_player.0 as isize, Ordering::SeqCst);
    note_player_process(0, MPV_IMAGE_NAME);
    ensure_pinned_sibling_box(moved);
    assert_eq!(
        window_box_on_screen(mpv_player),
        moved,
        "a film mpv is drawing is the band's picture, so its window travels with the band"
    );

    VIDEO_HWND.store(ffplay_player.0 as isize, Ordering::SeqCst);
    note_player_process(0, FFPLAY_IMAGE_NAME);
    // On screen, the way the film this stands for is, and that is the whole of what the answer
    // below is about: what the hold-off refuses FFmpeg's window is the *place*, while a window that
    // is not on screen is owed one wherever the chrome is (see
    // `put_a_player_that_is_not_on_screen_back`).
    unsafe {
        let _ = ShowWindow(ffplay_player, SW_SHOWNOACTIVATE);
    }
    ensure_pinned_sibling_box(moved);
    assert_eq!(
        window_box_on_screen(ffplay_player),
        content,
        "FFmpeg's window is refused where the band has moved to: a drag hides it, and the pin's \
         chrome covers the band it was standing in"
    );

    // And there is nothing to place where there is no window of a player's in the band at all,
    // which is every preview this app draws into its own window — the media engine's route, where
    // a pin's band is this app's own pixels rather than a hole onto somebody else's.
    VIDEO_HWND.store(0, Ordering::SeqCst);
    note_player_process(0, MPV_IMAGE_NAME);
    assert!(
        !ensure_video_window_topmost(moved.0, moved.1, moved.2 - moved.0, moved.3 - moved.1),
        "the mpv arm finds no window of a player's to place, so it places nothing either"
    );

    unsafe {
        DestroyWindow(mpv_player).expect("the stand-in window let go of");
        DestroyWindow(ffplay_player).expect("the stand-in window let go of");
    }
    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
    note_player_process(previous_pid, previous_image);
    stand_pin(previous_pin);
}

/// The band a hold-off still owes: a player's window that is not on screen goes back into it, with
/// the raise held off and with a window that is already standing there left exactly where it stands.
///
/// **This is the state a park leaves behind, and the swap out of one is made of this place** (see
/// `put_a_player_that_is_not_on_screen_back`): a seek parks the film and hides its window, the cover
/// comes down and the band under it is this app's own transparent rows — a place that only raised a
/// window would leave the desktop standing in the band, which is what a pinned film does now while a
/// strip of its chrome is up and the pointer is on it, that being where a seek starts.
///
/// The window is a stand-in of this app's own (`stand_a_player_window`), which is never shown by
/// itself: what is asked here is whether the placement shows it, and at which box.
///
/// `cargo test --lib a_players_window_that_is_not_on_screen_is_put_back_in_the_band`
#[test]
fn a_players_window_that_is_not_on_screen_is_put_back_in_the_band() {
    use windows::Win32::UI::WindowsAndMessaging::DestroyWindow;

    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);
    let previous_image = video_process_image();
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let band = (200, 180, 520, 420);

    // The pin the placement is asked about: a strip of its chrome is up over the band, which is the
    // hold-off this window is owed the band through.
    stand_pin(Some(overlay_pin(content, PinChrome::always())));
    let _ = take_pin_re_raises();

    let player = unsafe { stand_a_player_window(content) };
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);
    note_player_process(0, FFPLAY_IMAGE_NAME);

    assert!(
        !ensure_video_window_topmost(band.0, band.1, band.2 - band.0, band.3 - band.1),
        "the hold-off is the raise, so nothing was raised over the strip the pin is showing"
    );
    assert!(
        unsafe { IsWindowVisible(player).as_bool() },
        "and the window is on screen: the band's own rows are transparent, so a window left hidden \
         in them is the desktop behind the pin"
    );
    assert_eq!(
        window_box_on_screen(player),
        band,
        "put at the box the band is in, which is the place the swap out of a park owes"
    );
    assert_eq!(
        take_pin_re_raises(),
        0,
        "and nothing was put back over a raise that was never made"
    );

    // A window that is standing there already is left exactly where it stands: what the hold-off
    // refuses for a film on screen is the place as well as the raise, which is what the chrome being
    // up over the band means.
    let _ = ensure_video_window_topmost(
        content.0,
        content.1,
        content.2 - content.0,
        content.3 - content.1,
    );
    assert_eq!(
        window_box_on_screen(player),
        band,
        "a window that is on screen is not moved by a hold-off"
    );

    unsafe { DestroyWindow(player).expect("the stand-in window let go of") };
    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
    note_player_process(previous_pid, previous_image);
    stand_pin(previous_pin);
}

/// A collapsed pin places nothing: its band went with the window, so even a hidden film is left
/// hidden rather than put back into a band nobody can see.
///
/// A collapsed pin has two asks of every road that puts a player's window somewhere: `hide_pinned_windows`
/// takes the film off the screen with the window, and what must then not happen is any road
/// putting it back — the tick's re-assertion, the monitor's pass, or a drag's move — because a
/// place at the pin's box is a chrome-less film over Explorer beside the bubble. Both roads are
/// asked here through the same stand-in window the band tests stand: the tick's road through
/// `ensure_video_window_topmost` and a drag's through `ensure_pinned_sibling_box` (see
/// `hide_pinned_windows`).
#[test]
fn a_collapsed_pin_places_no_players_window() {
    use windows::Win32::UI::WindowsAndMessaging::{DestroyWindow, ShowWindow, SW_HIDE};

    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_image = video_process_image();
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let band = (200, 180, 520, 420);

    // The pin as a collapse leaves it: up, with its window off the screen.
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.collapsed = true;
    stand_pin(Some(pin));
    note_player_process(0, MPV_IMAGE_NAME);

    // The tick's road, over a film that is still on screen.
    let player = unsafe { stand_a_player_window(content) };
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);
    unsafe {
        let _ = ShowWindow(player, SW_SHOWNOACTIVATE);
    }
    assert!(
        !ensure_video_window_topmost(band.0, band.1, band.2 - band.0, band.3 - band.1),
        "a collapsed pin places nothing: the band went with the window, so the film stays where it stands"
    );
    assert_eq!(
        window_box_on_screen(player),
        content,
        "and a film on screen is not moved into a band nobody can see"
    );

    // And the hold-off's own exception, over a film that is not: a hidden window is still owed no
    // band while the pin is collapsed.
    unsafe {
        let _ = ShowWindow(player, SW_HIDE);
    }
    assert!(
        !ensure_video_window_topmost(band.0, band.1, band.2 - band.0, band.3 - band.1),
        "a collapsed pin owes no band either: the place beside the bubble is a chrome-less film over Explorer"
    );
    assert!(
        !unsafe { IsWindowVisible(player).as_bool() },
        "so a hidden film stays hidden"
    );

    // A drag's move asks the same question of the same pin and gets the same answer.
    unsafe {
        let _ = ShowWindow(player, SW_SHOWNOACTIVATE);
    }
    ensure_pinned_sibling_box(band);
    assert_eq!(
        window_box_on_screen(player),
        content,
        "a drag's move carries nothing while the pin is collapsed"
    );

    unsafe { DestroyWindow(player).expect("the stand-in window let go of") };
    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    note_player_process(previous_pid, previous_image);
    stand_pin(previous_pin);
}

/// Whether a window carries the extended style that makes a press on it fall through to whatever is
/// behind it, read back off the machine rather than remembered from what was asked for.
fn is_click_through(hwnd: HWND) -> bool {
    unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) & WS_EX_TRANSPARENT.0 as isize != 0 }
}

/// A film begun *while the pin is a bubble* comes off the screen, which is the half of a collapse
/// that `hide_pinned_windows` cannot do: that call hides the window the collapse found, and a player
/// begun afterwards — a walk's next file, an end-of-file step — puts up a window of its own that
/// nothing else hides.
///
/// **The road is the monitor's, because it is the one that finds a window a player showed itself.**
/// Neither player takes a switch that means "come up off the screen", so the film a walk lands on
/// behind a bubble stands chrome-less at the box the pin's own band had, with the bubble beside it
/// painting the same film out of frames read from the file. What a *windowed* pin is owed is the
/// opposite — the same pass leaves it on the screen it is shown in, and that is asserted here too,
/// because the hide is a fact about the bubble and not about a player's window.
#[test]
fn a_collapsed_pin_takes_a_players_window_off_the_screen() {
    use windows::Win32::UI::WindowsAndMessaging::{
        DestroyWindow, ShowWindow, SW_HIDE, SW_SHOWNOACTIVATE,
    };

    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);
    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_image = video_process_image();
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let player = unsafe { stand_a_player_window(content) };
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);

    // The pin as a collapse leaves it, over a film a player has just put its window up for.
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.collapsed = true;
    stand_pin(Some(pin));
    note_player_process(0, MPV_IMAGE_NAME);
    unsafe {
        let _ = ShowWindow(player, SW_SHOWNOACTIVATE);
    }
    assert!(
        unsafe { IsWindowVisible(player).as_bool() },
        "the window is up, as a player leaves it: a film put up beside a bubble is the defect"
    );

    unsafe { apply_noactivate_to_hwnd(player) };
    assert!(
        !unsafe { IsWindowVisible(player).as_bool() },
        "and the monitor's pass takes it off the screen: the picture a bubble shows is the frames \
         read beside the player, and a chrome-less film beside it is a second picture of one film"
    );

    // And the pin put back up: the same pass places the film in the band it is shown in rather than
    // hiding it, so what was asserted above is the bubble's rule and not the monitor's.
    let mut up = take_pin_for_a_test().expect("the pin is up");
    up.collapsed = false;
    stand_pin(Some(up));
    unsafe {
        let _ = ShowWindow(player, SW_SHOWNOACTIVATE);
    }
    unsafe { apply_noactivate_to_hwnd(player) };
    assert!(
        unsafe { IsWindowVisible(player).as_bool() },
        "a windowed pin's film is the picture and is left on the screen under it"
    );

    unsafe {
        let _ = ShowWindow(player, SW_HIDE);
        DestroyWindow(player).expect("the stand-in window let go of");
    }
    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    note_player_process(previous_pid, previous_image);
    stand_pin(previous_pin);
}

/// A player begun behind a bubble is begun off the desktop, which is what keeps its window from
/// being seen at all: the monitor's hide is a pass of a thread later, and a window that has just
/// been shown is a black rectangle over the desktop in the meantime (see `bubble_players_box`).
///
/// What the box is *not* is a different size: a player renders at the size it was begun at, so the
/// width the band has is kept and only the place moves — and a box asked for with no bubble up, or
/// with a pin that is not a bubble, is the box it always was.
#[test]
fn a_player_begun_behind_a_bubble_is_begun_off_the_desktop() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pin = take_pin_for_a_test();
    let content = (100, 80, 420, 320);
    let width = content.2 - content.0;

    // No pin at all: a hover's player, at the box the preview was laid out at.
    stand_pin(None);
    assert_eq!(
        bubble_players_box(content.0, content.1, width),
        (content.0, content.1),
        "a hover's player is the picture and is begun where the preview was laid out"
    );

    // A pin whose window is up: the film's window stands in the band, which is where it goes.
    stand_pin(Some(overlay_pin(content, PinChrome::always())));
    assert_eq!(
        bubble_players_box(content.0, content.1, width),
        (content.0, content.1),
        "a windowed pin's player is begun in the band it is watched in"
    );

    // And a collapsed pin: the same box, moved clear of every display the desktop has.
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.collapsed = true;
    stand_pin(Some(pin));

    let left = unsafe { GetSystemMetrics(SM_XVIRTUALSCREEN) };
    let (x, y) = bubble_players_box(content.0, content.1, width);
    assert_eq!(y, content.1, "only the place across the desktop moves");
    assert!(
        x + width < left,
        "a window of the band's own width wholly beyond the desktop's left edge ({x} + {width} \
         against {left}) is over no display at all, so a player that shows its window the moment it \
         has one shows nothing"
    );

    stand_pin(previous_pin);
}

/// A press on the film of a pin whose player is mpv is a move of the pin, from the pin's own box.
///
/// This is the far end of the band such a pin now wins the point of: what a press that arrives
/// there becomes. It is the road every other pin's body already takes — a media press is the pin's
/// by `pinned_content_is_the_pins`, and the caption, the buttons, the edges and the corners are
/// asked about first and win where they are drawn — so the film is a handle rather than a gesture
/// of its own. Nothing is parked and nothing is held, because mpv's window *is* the film and is
/// carried rather than put away (see `gesture_press_freeze`).
#[test]
fn a_press_on_the_film_of_an_mpv_pin_moves_the_pin_from_its_own_box() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);
    let previous_image = video_process_image();
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let content = (100, 80, 420, 320);

    // The mpv route with a pin's own file: the press leaves the film playing and the drag is a move
    // of the window it already has (see `stand_mpv_in`).
    stand_ffplay_in(Some(false));
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = PathBuf::from("a-film-mpv-draws.mkv");
    let caption = pin.caption;
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    // This app writes the player down as it begins one, and nothing here begins one (see
    // `note_player_process`).
    note_player_process(0, MPV_IMAGE_NAME);

    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }

    // A point inside the picture rather than on the strip above it: the caption is a handle of its
    // own, and the press this test is about is the one the band now lets through.
    let (x, y) = (content.0 + 10, content.1 + caption + 10);
    let hwnd = HWND(0x1000 as *mut _);
    assert!(
        unsafe { pinned_press(hwnd, x, y) },
        "a press on the picture is the pin's to act on, and mpv's film is the picture"
    );
    assert_eq!(
        pin_state().and_then(|pinned| pinned.pin().and_then(|pin| pin.pressed)),
        None,
        "and it is the picture's own road rather than the caption's: a press below the strip arms \
         no button, whatever is drawn there"
    );

    // The road that press takes, begun from a box a test can name: `pinned_press` answers a press
    // through `Win32PinWindow`, which reads the machine for the real pointer and the real window,
    // so what the drag reports is asked here instead.
    let window = a_window_at(content);
    begin_pin_drag(hwnd, &window, PinDragAction::Move, true);

    let drag = pin_state()
        .and_then(|pinned| pinned.pin().and_then(|pin| pin.dragging))
        .expect("the press began the pin's own drag");
    assert!(
        matches!(drag.action, PinDragAction::Move),
        "and what it began is a move of the window: no resume, no click-to-pause, nothing else"
    );
    assert_eq!(
        drag.window, content,
        "from the box the pin is standing at, which is the box the drag reports and the box the \
         window and the film are carried from"
    );
    assert_eq!(
        pinned_content(),
        Some(content),
        "and the band the press landed in is that same box"
    );
    assert!(
        !pin_player_is_parked(),
        "mpv is left drawing: a gesture writes commands down its pipe rather than ending its player"
    );

    finish_pin_drag(hwnd, &window);

    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
    note_player_process(previous_pid, previous_image);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// The window a player draws in keeps the style that makes a press fall through it, through every
/// road this app takes to that window.
///
/// That bit is what lets a press on the film past the player and into the pin (see
/// `film_band_is_hit_testable`), and four roads write that window's extended style: the monitor's
/// re-assert, a tick, a drag's own step and a re-assert again. Each is asked here to keep what it
/// found rather than to re-decide it — none of them may take the bit away, and the player is
/// created without it, which is what makes the four assertions mean anything.
#[test]
fn a_players_window_keeps_its_click_through_style_through_a_tick_a_drag_and_a_re_assert() {
    use windows::Win32::UI::WindowsAndMessaging::DestroyWindow;

    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _stand_in = POINTER_STAND_IN.lock().expect("the pointer's own state");

    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);
    let previous_image = video_process_image();
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let content = (100, 80, 420, 320);
    let player = unsafe { stand_a_player_window(content) };
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);
    note_player_process(0, MPV_IMAGE_NAME);

    stand_ffplay_in(Some(false));
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.path = PathBuf::from("a-film-whose-window-is-clicked-through.mkv");
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }

    // The list this app asserts on a player's window carries the bit, which is what makes the roads
    // below unable to be the road that loses it.
    assert!(
        player_ex_style_for(0) & WS_EX_TRANSPARENT.0 as isize != 0,
        "the style this app puts on a player's window is click-through among the rest"
    );

    // As a player creates its window, and this app is the only thing that ever puts the bit on.
    unsafe { SetWindowLongPtrW(player, GWL_EXSTYLE, WS_EX_TOOLWINDOW.0 as isize) };
    assert!(
        !is_click_through(player),
        "a player's window is created without the bit, so what is asserted below is put there by \
         this app and not by the machine"
    );

    // The monitor's own road.
    unsafe { apply_noactivate_to_hwnd(player) };
    assert!(
        is_click_through(player),
        "the monitor's re-assert is what puts the click-through style on"
    );

    // A tick.
    ensure_video_window_topmost(
        content.0,
        content.1,
        content.2 - content.0,
        content.3 - content.1,
    );
    assert!(
        is_click_through(player),
        "a tick re-asserts the three bits it knows about and leaves the rest of the style alone"
    );

    // A drag, and one step of it.
    let hwnd = HWND(0x1000 as *mut _);
    let window = a_window_at(content);
    begin_pin_drag(hwnd, &window, PinDragAction::Move, true);
    assert!(pin_is_dragging(), "the hand is carrying the window");
    step_the_pin_drag_by(20, 10);

    // The stand-in is re-asserted here rather than left standing across the pair, because the
    // record of a player's window is the whole process's and not this test's (see
    // `ensure_noactivate_monitor`).
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);

    unsafe { apply_pin_drag(hwnd) };
    assert!(
        is_click_through(player),
        "and a drag's step carries the player's window without taking its click-through away, \
         which is what the band's press has to reach the pin through"
    );

    // And a re-assert over what the drag left.
    unsafe { apply_noactivate_to_hwnd(player) };
    assert!(
        is_click_through(player),
        "a re-assert over a style this app last wrote keeps it"
    );

    finish_pin_drag(hwnd, &window);

    unsafe { DestroyWindow(player).expect("the stand-in window let go of") };
    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
    note_player_process(previous_pid, previous_image);
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A pinned film mpv draws is raised like every other film — and held off like every other film.
///
/// **The band a film mpv draws stands in is the pin's own pixel** (see `film_band_is_hit_testable`),
/// so what a press on the picture lands in is the pin, whatever band the player's window is in — the
/// player's window being click-through besides (see `player_ex_style_for`). A player left anywhere
/// but the topmost band is therefore not a pin that keeps the press: it is a film Explorer covers
/// the moment the user clicks into it, with the band over it going on taking the press. The raise
/// is asked for on every road while nothing of the pin's own is drawn over the band, and the pin
/// is put back over the raised film in the same breath (see `place_player_window`).
///
/// **A strip of the pin's chrome, the volume popup and the card's menu hold the raise off for mpv's
/// film as they do for FFmpeg's**: a raise there puts the film over what the user is looking at,
/// then the pin back over it — a two-step reorder that blinks the chrome, and one that blinks a
/// menu's panel every time a raise lands between one of `render_pin_menu`'s raises and the next.
/// What keeps the film visible under them
/// instead is mpv's own window holding its band (`--ontop`), and what keeps the press the pin's is
/// the pin's own raise; what a band that has moved still owes is the *place*, which travels without
/// a raise (see `only_an_mpv_players_window_is_placed_with_the_pins_chrome_up`). A park is refused
/// for every window there is, and the media engine's route draws its band in this app's own window
/// and begins no player, so the decision never reaches it. This asserts the decision alone; the
/// raisers that read it are asked in
/// `a_pinned_mpv_film_is_placed_by_a_tick_a_drag_and_a_monitor_pass`.
#[test]
fn a_pinned_film_mpv_draws_is_raised_like_every_other_film() {
    assert_eq!(
        player_order(MPV_IMAGE_NAME, true, false, false, false, false),
        PlayerOrder::Raise,
        "an mpv pin's film is raised with nothing of the pin's own drawn over the band: the pin \
         is put back over it right after"
    );
    for (chrome_up, volume_open, menu_open, what) in [
        (
            true,
            false,
            false,
            "with a strip of the pin's chrome up over it",
        ),
        (false, true, false, "with the volume popup open over it"),
        (false, false, true, "with the card's menu up over it"),
    ] {
        assert_eq!(
            player_order(
                MPV_IMAGE_NAME,
                true,
                chrome_up,
                volume_open,
                menu_open,
                false
            ),
            PlayerOrder::Leave,
            "an mpv pin's film is left exactly where it is {what}: a raise would put the film \
             over what the user is looking at and blink it, while mpv's own window holds the band \
             and the pin's own raise carries the press"
        );
    }

    assert_eq!(
        player_order(MPV_IMAGE_NAME, false, false, false, false, false),
        PlayerOrder::Raise,
        "the same film on a hover is raised as it always was: a hover must keep falling through to \
         what is behind it"
    );

    assert_eq!(
        player_order(FFPLAY_IMAGE_NAME, true, false, false, false, false),
        PlayerOrder::Raise,
        "an ffplay pin raises its player over the band as it always has"
    );
    for (chrome_up, volume_open, menu_open, what) in [
        (
            true,
            false,
            false,
            "while the pin's own chrome is drawn over the film",
        ),
        (false, true, false, "while the volume popup is open over it"),
        (false, false, true, "while the card's menu floats over it"),
    ] {
        assert_eq!(
            player_order(
                FFPLAY_IMAGE_NAME,
                true,
                chrome_up,
                volume_open,
                menu_open,
                false
            ),
            PlayerOrder::Leave,
            "an ffplay pin's player is left exactly where it is {what}: on that route the band is \
             that very window, and a raise would bury what this app is drawing there"
        );
    }

    assert_eq!(
        player_order(FFPLAY_IMAGE_NAME, true, false, false, false, true),
        PlayerOrder::Leave,
        "and a film parked for a gesture is left alone: a raise here would show it back over the \
         cover the park painted"
    );
    assert_eq!(
        player_order(MPV_IMAGE_NAME, true, false, false, false, true),
        PlayerOrder::Leave,
        "mpv's own film included — nothing on screen is raised, and a parked band would be the \
         desktop under the pin's own pixel"
    );
}

/// A player's window is asserted the styles **with** the topmost bit, and the click-through style
/// besides. The film is topmost for as long as the pin is Always On Top, because a pin's band is the
/// band its film stands in (see `pin_holds_the_front`).
///
/// Topmost is what keeps a film above every ordinary window — Explorer, and whatever else the user
/// brings up — which is the whole of why a band painted over a film is not a band over a film nobody
/// can see. The press on the picture is the pin's because the band is the pin's own pixel and the
/// pin is put back over the film right after it is raised (see `film_band_is_hit_testable`,
/// `place_player_window`), so what the style must not do is drop the click-through bit the press
/// falls through the player by. The raise half of the same invariant is
/// `a_pinned_players_placement_puts_the_pin_back_over_it`; the style half is this test.
#[test]
fn a_players_window_is_asserted_topmost() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_pin = take_pin_for_a_test();
    let previous_front = PIN_FRONT_OWED.load(Ordering::SeqCst);
    let always_on_top = pin_always_on_top();
    let move_to_top = CONFIG
        .lock()
        .map(|config| config.pin_update_move_to_top)
        .unwrap_or(true);
    stand_pin(Some(overlay_pin((100, 80, 420, 320), PinChrome::always())));
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
        config.pin_update_move_to_top = true;
    }
    assert!(
        pin_holds_the_front(),
        "a pin that is Always On Top holds the band, and that is the band the film is asserted into"
    );

    let a_style_the_player_needs = 0x0004_0000isize;
    let asserted = player_ex_style_for(WS_EX_TOPMOST.0 as isize | a_style_the_player_needs);

    assert_eq!(
        asserted & WS_EX_TOPMOST.0 as isize,
        WS_EX_TOPMOST.0 as isize,
        "a player's window is asserted topmost: anything less and the film is under the first \
         ordinary window the user activates, while the band over it goes on taking the press"
    );
    assert_eq!(
        asserted & a_style_the_player_needs,
        a_style_the_player_needs,
        "and every style the player itself needs is kept"
    );
    assert_ne!(
        asserted & WS_EX_TRANSPARENT.0 as isize,
        0,
        "and the click-through bit with it, which is what a press on the film reaches the pin \
         through"
    );

    PIN_FRONT_OWED.store(previous_front, Ordering::SeqCst);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
        config.pin_update_move_to_top = move_to_top;
    }
    stand_pin(previous_pin);
}

/// The film of a pinned mpv preview is placed where the band is, and put under the pin, through
/// the roads that put a player's window somewhere.
///
/// The tick's re-assertion and the monitor's pass are the two roads that raise a player's window
/// somewhere, and each is asked here to have read the placement decision (see `player_order`): the
/// answer it came to is recorded for a test's eye, because a test run has no screen to read a
/// z-order off and the answer is the whole of what a raiser decides before it places. A drag's move
/// carries the film without reading it — the band's new box with the order left alone (see
/// `ensure_pinned_sibling_box`) — so what this asks of the drag is the carry the sibling test pins
/// down, and the decision it asserts is the one the tick and the monitor came to. A `Leave` in any
/// one of them is what this catches — a film left where it stands is a film
/// under Explorer the moment the user clicks into it, while the band painted over it goes on taking
/// the press — and what the answer does with the placement is
/// `a_pinned_players_placement_puts_the_pin_back_over_it`.
#[test]
fn a_pinned_mpv_film_is_placed_by_a_tick_a_drag_and_a_monitor_pass() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let _stand_in = POINTER_STAND_IN.lock().expect("the pointer's own state");

    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);
    let previous_image = video_process_image();
    let previous_pin = take_pin_for_a_test();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let content = (100, 80, 420, 320);

    // The mpv route under a pin whose chrome is down — the band and the film and nothing of the
    // pin's own over it.
    note_player_process(0, MPV_IMAGE_NAME);
    stand_ffplay_in(Some(false));
    stand_mpv_in(Some(true));
    stand_video_engine_in(Some((VideoEngine::Mpv, true)));
    let mut pin = overlay_pin(
        content,
        PinChrome {
            caption: false,
            bar: false,
            until: None,
            step_in_flight: false,
        },
    );
    pin.path = PathBuf::from("a-film-the-pin-stands-above.mkv");
    pin.transport.begun(30.0, true, false);
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    let mut video = create_loading_media(320, 240);
    video.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(video);
    }

    // The tick's own re-assertion.
    let _ = ensure_video_window_topmost(
        content.0,
        content.1,
        content.2 - content.0,
        content.3 - content.1,
    );
    assert_eq!(
        last_player_order(),
        Some(PlayerOrder::Raise),
        "the tick's re-assertion raises a pinned mpv film over every ordinary window, and the pin \
         is put back over it with it"
    );

    // The monitor's pass, the one nobody can hold off from outside.
    let hwnd = HWND(0x1000 as *mut _);
    unsafe { apply_noactivate_to_hwnd(hwnd) };
    assert_eq!(
        last_player_order(),
        Some(PlayerOrder::Raise),
        "and the monitor's own pass, which no caller can hold off, raises it out of Explorer's \
         reach the same way"
    );

    // A drag's move: the player travels with the band rather than being left behind it.
    let window = a_window_at(content);
    begin_pin_drag(hwnd, &window, PinDragAction::Move, true);
    step_the_pin_drag_by(10, 10);
    unsafe { apply_pin_drag(hwnd) };
    assert_eq!(
        last_player_order(),
        Some(PlayerOrder::Raise),
        "and a drag's move takes the film with the band: a window left behind a band that has \
         moved is the desktop showing through the hole"
    );
    finish_pin_drag(hwnd, &window);

    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
    note_player_process(previous_pid, previous_image);
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_video_engine_in(None);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A player's window is raised and the pin is put back over it — and a hover's, or an answer that
/// places nothing, is not.
///
/// **This is the half of the invariant that has nothing but a screen to be read off**: both windows
/// end up topmost, so which of the two stands in front is the window manager's answer and a test run
/// has no screen. What is asked instead is the decision the placement came to, counted where it is
/// made (see `place_player_window`, `take_pin_re_raises`). The pin's window is the one that has to
/// be put back there, because what the film would otherwise cover is the pin's own pixels — the
/// caption and the transport bar drawn into the band over the media, and the press on the picture.
///
/// The answers that are *not* a pinned raise are asked with it, because they are the whole reason
/// the count is taken at the raisers rather than of `player_order`: a `Leave` never reaches a
/// placement at all (both raisers refuse it before asking), and a hover's raise has no pin to put
/// over anything.
#[test]
fn a_pinned_players_placement_puts_the_pin_back_over_it() {
    use windows::Win32::UI::WindowsAndMessaging::DestroyWindow;

    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);
    let previous_image = video_process_image();
    let previous_pin = take_pin_for_a_test();
    let always_on_top = pin_always_on_top();
    // The pair the re-raise is about is a Topmost one: a pin with the setting off puts its film
    // *below* itself, so there is nothing to put the pin back over (see `place_player_window`).
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = true;
    }

    let content = (100, 80, 420, 320);
    let player = unsafe { stand_a_player_window(content) };
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);

    let _ = take_pin_re_raises();

    // A pinned film mpv draws, with nothing of the pin's own up: raised, and the pin put back
    // over the film in the same placement — which is where the pin's chrome is drawn.
    note_player_process(0, MPV_IMAGE_NAME);
    stand_pin(Some(overlay_pin(
        content,
        PinChrome {
            caption: false,
            bar: false,
            until: None,
            step_in_flight: false,
        },
    )));
    assert!(
        ensure_video_window_topmost(
            content.0,
            content.1,
            content.2 - content.0,
            content.3 - content.1,
        ),
        "the band the player's window stands in is placed while nothing of the pin's own is up"
    );
    assert_eq!(
        take_pin_re_raises(),
        1,
        "and the pin is put back over the film it raised: the pin's chrome and the press on the \
         picture are a raise of the pin's own window and not a fact about the film"
    );

    // A hover: the same raise, and nothing to put over it.
    stand_pin(None);
    let _ = ensure_video_window_topmost(
        content.0,
        content.1,
        content.2 - content.0,
        content.3 - content.1,
    );
    assert_eq!(
        take_pin_re_raises(),
        0,
        "a hover's player is raised and nothing is put over it: there is no pin standing there"
    );

    // And an answer that places nothing is followed by nothing: either player's window with a
    // strip of the pin's chrome up is held off before the placement (see `player_order`).
    note_player_process(0, FFPLAY_IMAGE_NAME);
    stand_pin(Some(overlay_pin(content, PinChrome::always())));
    assert!(
        !ensure_video_window_topmost(
            content.0,
            content.1,
            content.2 - content.0,
            content.3 - content.1,
        ),
        "a strip of the pin's chrome holds FFmpeg's own window off: the band that route stands in \
         is that very window, so a raise would cover what the user is looking at"
    );
    assert_eq!(
        take_pin_re_raises(),
        0,
        "and nothing was placed, so nothing was put back over anything"
    );

    // mpv's film with the same strip up is held off the same way: the chrome would blink under a
    // raise, mpv's own window holds the band, and the pin's own raise carries the press — while a
    // band that has moved is placed without a raise (see `ensure_pinned_sibling_box`).
    note_player_process(0, MPV_IMAGE_NAME);
    stand_pin(Some(overlay_pin(content, PinChrome::always())));
    assert!(
        !ensure_video_window_topmost(
            content.0,
            content.1,
            content.2 - content.0,
            content.3 - content.1,
        ),
        "a strip of the pin's chrome holds mpv's film off too: a raise would put the film over \
         the chrome for a frame, and the place a moved band owes travels without one"
    );
    assert_eq!(
        take_pin_re_raises(),
        0,
        "and nothing was placed, so nothing was put back over anything either"
    );

    // And the card's menu is a hold-off of the same kind, for a window of its own: every raise the
    // panel gets is `render_pin_menu`'s own, and a raise here would leave the film over the panel
    // until the next one — the blink a hand reads off the rows of the menu that stand over the
    // picture (see `pin_menu_open`).
    note_player_process(0, MPV_IMAGE_NAME);
    let mut with_menu = overlay_pin(
        content,
        PinChrome {
            caption: false,
            bar: false,
            until: None,
            step_in_flight: false,
        },
    );
    with_menu.menu.open = true;
    stand_pin(Some(with_menu));
    assert!(
        !ensure_video_window_topmost(
            content.0,
            content.1,
            content.2 - content.0,
            content.3 - content.1,
        ),
        "the card's menu holds the film off too: a raise puts the film over the panel the user is \
         reading"
    );
    assert_eq!(
        take_pin_re_raises(),
        0,
        "and nothing was placed, so nothing was put back over the panel either"
    );

    unsafe { DestroyWindow(player).expect("the stand-in window let go of") };
    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
    note_player_process(previous_pid, previous_image);
    stand_pin(previous_pin);
    if let Ok(mut config) = CONFIG.lock() {
        config.pin_always_on_top = always_on_top;
    }
}

/// The answer the chrome block's raise is asked against: whether the player's window is standing
/// over the pin, which is a fact of the machine's z-order rather than of this app's bookkeeping —
/// and one that changes without this app doing anything, because the film stands itself back up on
/// its own band (see `player_stands_over_the_pin`, `mpv_args`).
///
/// **Two windows of this app's own are what the answer is read off.** Both raisers put the film
/// *under* the pin and leave it there, so the state that has the chrome covered is one no placement
/// road reaches; it is put in here with the one thing that does reach it, a raise of the film, and
/// the two answers are asked against each other. What a red result discriminates: a walk that goes
/// the wrong way down the z-order answers `true` for the order both raisers put in, which is a raise
/// of the pin per tick instead of none — the churn the hold-offs are for — and the other way round
/// it answers `false` for a film that is over the chrome, which is the chrome left under the film
/// until the pointer asks for a strip of it again.
///
/// `cargo test --lib the_player_stands_over_the_pin_only_while_it_is_raised_over_it`
#[test]
fn the_player_stands_over_the_pin_only_while_it_is_raised_over_it() {
    use windows::Win32::UI::WindowsAndMessaging::{DestroyWindow, ShowWindow, SW_HIDE};

    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);

    let content = (100, 80, 420, 320);
    let pin = unsafe { stand_a_player_window(content) };
    let player = unsafe { stand_a_player_window(content) };
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);

    // A raise, asked of the machine the way `place_player_window` asks it: top of the band for a
    // window that is already in it, its own box, and on screen — a window that was never shown is
    // not one that covers anything, which is the stand-in's own answer further down.
    unsafe fn raise(hwnd: HWND, content: ScreenRegion) {
        let _ = SetWindowPos(
            hwnd,
            HWND_TOPMOST,
            content.0,
            content.1,
            content.2 - content.0,
            content.3 - content.1,
            SWP_NOACTIVATE | SWP_SHOWWINDOW,
        );
    }

    // A window put in the topmost band at all, which one `HWND_TOPMOST` does not do: the first
    // raise puts a window that is not yet topmost at the *foot* of that band, and the raise after
    // it is the ordering raise both raisers make (see `place_player_window`). Both of the app's two
    // windows are in the band from their first placement on, so this is the state each of the
    // stand-ins is put in before anything is asked about their order.
    unsafe fn make_topmost(hwnd: HWND, content: ScreenRegion) {
        raise(hwnd, content);
        raise(hwnd, content);
    }

    // The order both raisers put in: the film placed, and the pin put back over it.
    unsafe {
        make_topmost(player, content);
        make_topmost(pin, content);
    }
    assert!(
        !unsafe { player_stands_over_the_pin(pin) },
        "a film under the pin is the order every placement leaves standing, and a raise of the pin \
         for it is a raise the chrome already has"
    );

    // And the film standing itself back up over the pin, which is what mpv's own band does from the
    // first frame it draws: the pin owes a raise for as long as a strip of its chrome is up, which
    // is the raise the chrome block makes.
    unsafe { raise(player, content) };
    assert!(
        unsafe { player_stands_over_the_pin(pin) },
        "a film raised over the pin is the one that covers the caption and the transport bar the \
         pin draws into the band over it"
    );

    // The pin put back over the film: the film is under it again, and there is nothing to raise
    // for until the film stands itself up again.
    unsafe { raise(pin, content) };
    assert!(
        !unsafe { player_stands_over_the_pin(pin) },
        "the pin's own raise is what puts it back over the film, and the answer is the order that \
         raise left standing"
    );

    // A parked film is not on screen, and a window that is not on screen covers nothing — the
    // answer a drag's own moves owe while the window is hidden (see `pin_player_is_parked`).
    unsafe {
        raise(player, content);
        let _ = ShowWindow(player, SW_HIDE);
    }
    assert!(
        !unsafe { player_stands_over_the_pin(pin) },
        "a hidden film stands over nothing, whatever its place in the band"
    );

    // And a band with no player's window in it at all: a hover, or the media engine painting this
    // app's own rows.
    VIDEO_HWND.store(0, Ordering::SeqCst);
    assert!(
        !unsafe { player_stands_over_the_pin(pin) },
        "no player's window is nobody standing over the pin"
    );

    unsafe {
        let _ = DestroyWindow(player);
        let _ = DestroyWindow(pin);
    }
    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
}

/// The rule a press on a document the engine draws is read by, without a desktop under it: a
/// window the engine's own browser owns is the page, whatever the host's tree answers for.
///
/// The case that matters is the third one. A press on a picture was read as a press on somebody
/// else's window because the browser's window is not a child of the host's in the tree the walk
/// sees, and the pin was left where it was by a click that was meant to bring it forward.
#[test]
fn a_window_of_the_engines_browser_is_the_page() {
    // The host window itself, and a window drawn inside it.
    assert!(the_engines_window(true, false, 4242, 4242));
    assert!(the_engines_window(false, true, 0, 4242));
    // The browser's own window, which the host's tree does not answer for.
    assert!(the_engines_window(false, false, 4242, 4242));
    // Another program's window, a browser this app never started, and nothing at all.
    assert!(!the_engines_window(false, false, 9001, 4242));
    assert!(!the_engines_window(false, false, 4242, 0));
    assert!(!the_engines_window(false, false, 0, 0));
}

/// **Every road that puts a player's window somewhere asks for it rather than waiting on the
/// player's thread**, and says so in the call it makes.
///
/// This is the whole of the hang a pinned film had. A `SetWindowPos` on a window another process
/// owns is a blocking send to the thread that draws it, and that thread is not passive: mpv
/// reorders and repaints its own top-level window as it plays, and a reorder under a window of this
/// app's standing in the same band sends a message back into the thread that is waiting. Two
/// threads waiting on each other is a frozen film and a pin that has stopped following the hand,
/// with only the tray — a thread of its own — still answering (see `place_a_players_window`).
///
/// **What this can and cannot prove from inside one test.** A blocking wait needs a second thread
/// that will not answer, and the whole of the deadlock is that it never does; so what is asserted
/// here is the shape of the call the three roads make — the flag that asks for the placement rather
/// than waiting for it — for the band a drag carries, the band the tick re-asserts, and the place
/// under the pin a press asks for. The wait itself is asserted by
/// `a_players_window_is_placed_without_waiting_for_the_thread_that_owns_it`.
#[test]
fn every_players_window_placement_is_asked_for_rather_than_waited_on() {
    use windows::Win32::UI::WindowsAndMessaging::DestroyWindow;

    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_pid = VIDEO_PID.swap(0, Ordering::SeqCst);
    let previous_hwnd = VIDEO_HWND.swap(0, Ordering::SeqCst);
    let previous_preview = PREVIEW_HWND.load(Ordering::SeqCst);
    let previous_image = video_process_image();
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let moved = (140, 120, 460, 360);
    let player = unsafe { stand_a_player_window(content) };
    let pin = unsafe { stand_a_player_window(content) };

    // The chrome down, so that the tick's own raise is not held off: what is being read is the
    // placement each road makes, and a hold-off is a road that makes none.
    stand_pin(Some(overlay_pin(content, PinChrome::hidden())));
    PREVIEW_HWND.store(pin.0 as isize, Ordering::SeqCst);
    VIDEO_HWND.store(player.0 as isize, Ordering::SeqCst);
    note_player_process(0, MPV_IMAGE_NAME);

    // The band a step of a drag is carried to, which is the road a film frozen under the hand was.
    let _ = video_playback::take_the_players_window_placement();
    ensure_pinned_sibling_box(moved);
    let asked = video_playback::take_the_players_window_placement();
    assert_ne!(asked, 0, "the mpv arm placed the film's window");
    assert_ne!(
        asked & SWP_ASYNCWINDOWPOS.0,
        0,
        "and asked for the box rather than waiting for the thread that owns the window: {asked:#x}"
    );

    // The band the tick re-asserts, which is the same placement from the other raiser.
    let _ = video_playback::take_the_players_window_placement();
    assert!(
        ensure_video_window_topmost(moved.0, moved.1, moved.2 - moved.0, moved.3 - moved.1),
        "the tick puts the film in its band"
    );
    let asked = video_playback::take_the_players_window_placement();
    assert_ne!(
        asked & SWP_ASYNCWINDOWPOS.0,
        0,
        "and asks for it rather than waiting for it: {asked:#x}"
    );

    // The place under the pin a press asks for, which is a reorder of the player's window and so
    // the same wait by another name.
    let _ = video_playback::take_the_players_window_placement();
    unsafe { put_player_under_the_pin() };
    let asked = video_playback::take_the_players_window_placement();
    assert_ne!(
        asked & SWP_ASYNCWINDOWPOS.0,
        0,
        "and the place under the pin is asked for too: {asked:#x}"
    );

    unsafe {
        let _ = DestroyWindow(player);
        let _ = DestroyWindow(pin);
    }
    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
    PREVIEW_HWND.store(previous_preview, Ordering::SeqCst);
    note_player_process(previous_pid, previous_image);
    stand_pin(previous_pin);
}

/// **A placement of a window this app does not own comes back without the thread that owns it.**
///
/// The window here belongs to a thread that makes it and then stops reading its own queue — which
/// is the player's thread on the day the deadlock would happen, with nothing left to hide it. An
/// `SWP_ASYNCWINDOWPOS` request is posted to that thread and answered here at once; a placement
/// that waited on it would never come back at all, so the call is made on a thread of the test's
/// own and joined with a bound: a regression fails this test rather than wedging the suite.
#[test]
fn a_players_window_is_placed_without_waiting_for_the_thread_that_owns_it() {
    use std::sync::mpsc;
    use std::time::Duration;

    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let (made_at, made) = mpsc::channel::<isize>();
    let (held, hold) = mpsc::channel::<()>();
    let owner = std::thread::spawn(move || {
        // A window of this thread's own: from here on it is the thread's queue that answers for
        // it, and this thread never reads that queue again.
        let hwnd = unsafe { stand_a_player_window((0, 0, 32, 24)) };
        let _ = made_at.send(hwnd.0 as isize);
        let _ = hold.recv();
        unsafe {
            let _ = windows::Win32::UI::WindowsAndMessaging::DestroyWindow(hwnd);
        }
    });
    let hwnd = made.recv().expect("a window of the other thread's");

    let (placed_at, placed) = mpsc::channel::<u32>();
    std::thread::spawn(move || {
        // A handle is not `Send`, and it is not one of this app's either: the window is another
        // thread's, which is the whole of what is being asserted.
        let hwnd = HWND(hwnd as *mut _);
        let _ = unsafe {
            video_playback::place_a_players_window(
                hwnd,
                HWND::default(),
                1,
                2,
                32,
                24,
                SWP_NOZORDER | SWP_NOACTIVATE,
            )
        };
        let _ = placed_at.send(video_playback::take_the_players_window_placement());
    });

    let asked = placed
        .recv_timeout(Duration::from_secs(5))
        .expect("a placement of another thread's window comes back");
    assert_ne!(
        asked & SWP_ASYNCWINDOWPOS.0,
        0,
        "because it was asked for rather than waited on: {asked:#x}"
    );

    let _ = held.send(());
    let _ = owner.join();
}

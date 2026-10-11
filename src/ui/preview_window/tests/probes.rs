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
fn send_pointer_to(x: i32, y: i32) {
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

    CreateWindowExW(
        WS_EX_TOOLWINDOW,
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
/// which is the arm the carry is read off, and `VIDEO_PID` is left zero with it so the style
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

    // And whose window it is, as the carry reads it: this app writes the player down as it begins
    // one, and nothing here begins one (see `note_player_process`).
    note_player_process(0, MPV_IMAGE_NAME);

    // The mpv route, chrome up: the press leaves the film playing and its window the one the
    // drag moves — nothing is parked, which is what puts the band's hole over the desktop.
    stand_ffplay_in(Some(false));
    stand_mpv_in(Some(true));
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
        send_pointer_to(content.0 + dx, content.1 + dy);
        std::thread::sleep(Duration::from_millis(25));

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
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// Only a window mpv is drawing is carried with a dragged pin: every other answer the placement
/// can give is the one it gave before.
///
/// **Why it is asked of the placement rather than of a drag.** The carry is one condition —
/// whether the recorded player's window is mpv's (`video_process_image`) — and the answers it is
/// not are the ones the refusal is for. FFmpeg's window is the film a park hides and a strip of
/// the pin's chrome stands over, so it is refused; and a band with no window of a player's in it,
/// which is every preview this app draws itself, has nothing to carry at all. A drag for each
/// would be the same arithmetic three times, so the drag the mpv arm gets is the one above and
/// the rest are pinned here.
///
/// Both windows are stand-ins of this app's own (`stand_a_player_window`) and `VIDEO_PID` is left
/// zero, so no player's process is looked for and nothing of this machine's is styled.
#[test]
fn only_an_mpv_players_window_is_carried_under_the_pins_chrome() {
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
    // the whole of the refusal the carry is an exception to.
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
    ensure_pinned_sibling_box(moved);
    assert_eq!(
        window_box_on_screen(ffplay_player),
        content,
        "FFmpeg's window is refused where the band has moved to: a drag hides it, and the pin's \
         chrome covers the band it was standing in"
    );

    // And the carry has nothing to carry where there is no window of a player's in the band at
    // all, which is every preview this app draws into its own window — the media engine's route,
    // where a pin's band is this app's own pixels rather than a hole onto somebody else's.
    VIDEO_HWND.store(0, Ordering::SeqCst);
    note_player_process(0, MPV_IMAGE_NAME);
    assert!(
        !ensure_video_window_topmost(moved.0, moved.1, moved.2 - moved.0, moved.3 - moved.1),
        "the mpv arm finds no window of a player's to carry, so it places nothing either"
    );

    unsafe {
        DestroyWindow(mpv_player).expect("the stand-in window let go of");
        DestroyWindow(ffplay_player).expect("the stand-in window let go of");
    }
    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
    note_player_process(previous_pid, previous_image);
    stand_pin(previous_pin);
}

/// Whether a window carries the extended style that makes a press on it fall through to whatever is
/// behind it, read back off the machine rather than remembered from what was asked for.
fn is_click_through(hwnd: HWND) -> bool {
    unsafe { GetWindowLongPtrW(hwnd, GWL_EXSTYLE) & WS_EX_TRANSPARENT.0 as isize != 0 }
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
        player_ex_style_for(0, false) & WS_EX_TRANSPARENT.0 as isize != 0,
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
    send_pointer_to(content.0 + 20, content.1 + 10);
    std::thread::sleep(Duration::from_millis(25));

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
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// A pinned film mpv draws is the one answer the player is never raised over — and every other
/// route keeps the answer it has always had.
///
/// The exception is the one band Windows hands to the pin rather than to the window behind it —
/// an mpv film (see `film_band_is_hit_testable`) — because a player raised over the pin takes the
/// press the pin is owed. FFmpeg's player keeps its hole, a hover must go on falling through, and
/// a film parked for a gesture stays hidden. The media engine's route draws its band in this app's
/// own window and begins no player, so there is no window of somebody else's in the band to place
/// and the decision never reaches it (see
/// `only_an_mpv_players_window_is_carried_under_the_pins_chrome`). This asserts the decision alone;
/// the raisers that read it are asked in
/// `a_pinned_mpv_film_is_never_put_above_the_pin_by_a_tick_a_drag_or_a_monitor_pass`.
#[test]
fn a_pinned_film_mpv_draws_is_the_one_player_never_raised_over_the_pin() {
    for (chrome_up, volume_open, what) in [
        (false, false, "with nothing of the pin's own drawn over the band"),
        (true, false, "with a strip of the pin's chrome up over it"),
        (false, true, "with the volume popup open over it"),
    ] {
        assert_eq!(
            player_order(MPV_IMAGE_NAME, true, chrome_up, volume_open, false),
            PlayerOrder::Carry,
            "an mpv pin's film is carried and never raised over the pin, {what}"
        );
    }

    assert_eq!(
        player_order(MPV_IMAGE_NAME, false, false, false, false),
        PlayerOrder::Raise,
        "the same film on a hover is raised over the hover as it always was: a hover must keep \
         falling through to what is behind it"
    );

    assert_eq!(
        player_order(FFPLAY_IMAGE_NAME, true, false, false, false),
        PlayerOrder::Raise,
        "an ffplay pin raises its player over the band as it always has"
    );
    for (chrome_up, volume_open) in [(true, false), (false, true)] {
        assert_eq!(
            player_order(FFPLAY_IMAGE_NAME, true, chrome_up, volume_open, false),
            PlayerOrder::Leave,
            "and is left exactly where it is while the pin's own chrome is drawn over the film"
        );
    }

    assert_eq!(
        player_order(FFPLAY_IMAGE_NAME, true, false, false, true),
        PlayerOrder::Leave,
        "and a film parked for a gesture is left alone: a raise here would show it back over the \
         cover the park painted"
    );
}

/// A carried player is asserted the styles **without** the topmost bit — and that is a second place
/// the order is decided, not a detail of the raise.
///
/// A topmost window can never be carried under a topmost pin, so a style that keeps the bit puts the
/// player back in front of the pin whatever the raise is held off from doing: the press on the
/// picture is taken from the pin again, and the film is no longer seen through the pin's own band
/// but standing in front of it. The raise half of this is guarded by
/// `a_pinned_mpv_film_is_never_put_above_the_pin_by_a_tick_a_drag_or_a_monitor_pass`; the style half
/// is this test, and it is the half a raise-only guard leaves open.
#[test]
fn a_carried_players_window_is_never_asserted_topmost() {
    let a_style_the_player_needs = 0x0004_0000isize;
    let carried = player_ex_style_for(WS_EX_TOPMOST.0 as isize | a_style_the_player_needs, true);

    assert_eq!(
        carried & WS_EX_TOPMOST.0 as isize,
        0,
        "a carried player is asserted no topmost bit: the pin's band is the pixel a press has to \
         land in, and a topmost player is a player in front of it"
    );
    assert_eq!(
        carried & a_style_the_player_needs,
        a_style_the_player_needs,
        "and every style the player itself needs is kept"
    );

    assert_ne!(
        player_ex_style_for(0, false) & WS_EX_TOPMOST.0 as isize,
        0,
        "while a player raised over the band keeps the bit as it always has"
    );
    assert_ne!(
        player_ex_style_for(0, false) & WS_EX_TRANSPARENT.0 as isize,
        0,
        "and keeps the click-through style with it"
    );
}

/// The film of a pinned mpv preview is never put above the pin, through every road that puts a
/// player's window somewhere.
///
/// This is the point the user's press has to reach the pin through. The tick's re-assertion, the
/// monitor's pass and a drag's move are the three roads that put a player's window somewhere, and
/// each is asked here to have read the placement decision (see `player_order`): the answer it came
/// to is recorded for a test's eye, because a test run has no screen to read a z-order off and the
/// answer is the whole of what a raiser decides. A raise in any one of them is what this catches.
#[test]
fn a_pinned_mpv_film_is_never_put_above_the_pin_by_a_tick_a_drag_or_a_monitor_pass() {
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
    // pin's own over it, which is the case the raise is not already held off by (see
    // `pin_chrome_up`).
    note_player_process(0, MPV_IMAGE_NAME);
    stand_ffplay_in(Some(false));
    stand_mpv_in(Some(true));
    let mut pin = overlay_pin(
        content,
        PinChrome {
            caption: false,
            bar: false,
            until: None,
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
        Some(PlayerOrder::Carry),
        "the tick's re-assertion carries a pinned mpv film rather than raising it over the pin"
    );

    // The monitor's pass, the one nobody can hold off from outside.
    let hwnd = HWND(0x1000 as *mut _);
    unsafe { apply_noactivate_to_hwnd(hwnd) };
    assert_eq!(
        last_player_order(),
        Some(PlayerOrder::Carry),
        "and the monitor's own pass, which no caller can hold off, leaves it under the pin too"
    );

    // A drag's move: the player is carried with the band and not raised over the pin.
    let window = a_window_at(content);
    begin_pin_drag(hwnd, &window, PinDragAction::Move, true);
    send_pointer_to(content.0 + 10, content.1 + 10);
    std::thread::sleep(Duration::from_millis(25));
    unsafe { apply_pin_drag(hwnd) };
    assert_eq!(
        last_player_order(),
        Some(PlayerOrder::Carry),
        "and a drag's move carries the film with the band rather than raising it over the pin"
    );
    finish_pin_drag(hwnd, &window);

    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);
    note_player_process(previous_pid, previous_image);
    stand_mpv_in(None);
    stand_ffplay_in(None);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

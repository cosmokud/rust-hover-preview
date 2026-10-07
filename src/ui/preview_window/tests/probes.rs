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

            video_player::play_audio(&path, volume, start);
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

use super::*;
use std::time::{Duration, Instant};

// Repro for the fatal pin bug: a pin showing a 4K film through ffplay,
// swapped onto a 1080p film the media engine plays, freezes the app —
// audio of the new film plays, no video ever shows, the window stops
// answering and the process has to be killed.
//
// The test drives the real handoff `swap_pinned_media` performs on its
// native arm (`stop_pinned_player`, then `video_player::play`, then the
// held wait for the first frame), with real files, a real ffplay and the
// real media engine, logging where it stalls.
//
// Env-gated like `video_take_cost`: `RHP_REPRO_HD` names a ~1080p film
// the engine plays natively, `RHP_REPRO_UHD` a ~4K one handed to ffplay.
// Without both, the test reports what to set and returns.
fn repro_files() -> Option<(PathBuf, PathBuf)> {
    let hd = std::env::var("RHP_REPRO_HD").ok().map(PathBuf::from)?;
    let uhd = std::env::var("RHP_REPRO_UHD").ok().map(PathBuf::from)?;
    (hd.is_file() && uhd.is_file()).then_some((hd, uhd))
}

/// Load the way a pin swap loads: the real route (`Best` at the default
/// configuration) and the real probe, so a mis-route fails loudly here
/// rather than wedging the wait further down.
fn load_for_swap(path: &PathBuf) -> Option<MediaData> {
    load_video_thumbnail(path, 960, 600, PreviewScale::FitToScreen)
}

/// Wait for the engine's first frame the way `PinSwapHold::settle` does,
/// with a deadline so a stall fails instead of hanging the suite. Returns
/// what the wait ended with: the frame, or a description of the stall.
fn wait_for_first_frame(held: &mut MediaData, what: &str, deadline: Duration) -> Option<String> {
    let started = Instant::now();
    let mut logged = Instant::now() - Duration::from_secs(99);
    let mut last_playing = None;
    let mut last_failing: Option<bool> = None;

    loop {
        // The hold's own take, into the held file's buffer, on this thread —
        // the preview thread in the app.
        let frame = held.take_native_video_frame();
        let playing = video_player::is_playing();
        let failing = video_player::failing_before_a_frame().is_some();

        if frame {
            return None;
        }
        if !playing || failing {
            return Some(format!(
                "{what}: wait abandoned after {} ms (playing={playing}, failing={failing})",
                started.elapsed().as_millis(),
            ));
        }
        if started.elapsed() >= deadline {
            return Some(format!(
                "{what}: NO first frame in {} ms and the wait never ended \
                 (playing={playing}, failing={failing}) — the pin would sit here forever",
                started.elapsed().as_millis(),
            ));
        }

        if last_playing != Some(playing) || last_failing != Some(failing) {
            eprintln!(
                "repro {what}: playing={playing} failing={failing} at {} ms",
                started.elapsed().as_millis(),
            );
            last_playing = Some(playing);
            last_failing = Some(failing);
        }
        if logged.elapsed() >= Duration::from_secs(1) {
            logged = Instant::now();
            eprintln!(
                "repro {what}: still waiting at {} ms (playing={playing}, failing={failing})",
                started.elapsed().as_millis(),
            );
        }
        std::thread::sleep(Duration::from_millis(16));
    }
}

#[test]
fn ffplay_to_native_swap_delivers_a_frame() {
    let Some((hd, uhd)) = repro_files() else {
        println!("set RHP_REPRO_HD to a 1080p film and RHP_REPRO_UHD to a 4K one to run this");
        return;
    };
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    pdf_preview::initialize_apartment();
    wic_image::initialize_apartment();

    eprintln!("repro: ffplay_available={} mf_started={} hwaccel={:?}", codecs::ffplay_available(), codecs::mf_started(), video_hw_accel_device());

    // Save the process-wide state this touches, so the next test walks
    // into what it walked into before.
    let mut previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut m| m.take());
    let previous_pid = VIDEO_PID.load(Ordering::SeqCst);
    let previous_hwnd = VIDEO_HWND.load(Ordering::SeqCst);

    let mut failure: Option<String> = None;

    // Prime the geometry the way hovers do before any pin exists: the
    // route reads the *cached* pixel count, and an unmeasured film is
    // handed to ffplay by design. Without this the fixtures would both
    // route to ffplay and the test would prove nothing about the swap.
    eprintln!("repro: probing both fixtures (real ffprobe)");
    match probe_video_geometry(&hd) {
        ProbedGeometry::Measured(g) => {
            eprintln!("repro: 1080p probe measured {}x{}", g.width, g.height)
        }
        ProbedGeometry::Unmeasurable => eprintln!("repro: 1080p probe UNMEASURABLE"),
    }
    match probe_video_geometry(&uhd) {
        ProbedGeometry::Measured(g) => {
            eprintln!("repro: 4K probe measured {}x{}", g.width, g.height)
        }
        ProbedGeometry::Unmeasurable => eprintln!("repro: 4K probe UNMEASURABLE"),
    }

    // The route has to split the fixtures the way the bug needs: the 4K
    // film through ffplay, the 1080p one through the media engine.
    let uhd_media = load_for_swap(&uhd);
    let hd_media = load_for_swap(&hd);
    eprintln!(
        "repro: 4K route={:?}, 1080p route={:?}",
        uhd_media.as_ref().map(|m| m.media_type),
        hd_media.as_ref().map(|m| m.media_type),
    );
    eprintln!(
        "repro: plays(1080p)={} engine_route(1080p)={} plays(4K)={} engine_route(4K)={}",
        video_player::plays(&hd),
        named_for_the_media_engine(&hd),
        video_player::plays(&uhd),
        named_for_the_media_engine(&uhd),
    );
    if uhd_media.as_ref().is_none_or(|m| m.media_type != MediaType::Video)
        || hd_media.as_ref().is_none_or(|m| m.media_type != MediaType::NativeVideo)
    {
        failure = Some(format!(
            "fixtures do not split across the engines (4K={:?}, 1080p={:?}); \
             check Video Engine=Best and that ffplay is installed",
            uhd_media.as_ref().map(|m| m.media_type),
            hd_media.as_ref().map(|m| m.media_type),
        ));
    }

    // Baseline: the 1080p film straight into the engine, with no ffplay
    // before it. If this already stalls, the machine (not the swap) is
    // what cannot play the file here.
    if failure.is_none() {
        if let Some(mut held) = hd_media {
            let (w, h) = (held.current_width(), held.current_height());
            eprintln!("repro baseline: play 1080p at {w}x{h}");
            video_player::play(&hd, w, h, 0, probed_picture(&hd, w, h));
            eprintln!("repro baseline: is_playing={}", video_player::is_playing());
            failure = wait_for_first_frame(&mut held, "baseline", Duration::from_secs(15));
            if failure.is_none() {
                eprintln!("repro baseline: first frame arrived");
            }
            video_player::stop();
        }
    }

    // The pin standing on the 1080p film first (the live native session
    // the 1080p -> 4K leg of the bug tears down), then swapped onto the
    // 4K film through ffplay, then swapped back: the full history.
    if failure.is_none() {
        // Leg 1: pin shows 1080p natively, frames flowing.
        let mut live = load_for_swap(&hd);
        if live.as_ref().is_none_or(|m| m.media_type != MediaType::NativeVideo) {
            failure = Some("1080p film would not load for leg 1".to_string());
        } else if let Some(ref media) = live {
            let (w, h) = (media.current_width(), media.current_height());
            eprintln!("repro leg 1: play 1080p at {w}x{h}");
            video_player::play(&hd, w, h, 0, probed_picture(&hd, w, h));
            let mut flowing = live.take().unwrap();
            failure = wait_for_first_frame(&mut flowing, "leg1", Duration::from_secs(15));
            if failure.is_none() {
                eprintln!("repro leg 1: 1080p playing natively, frames flowing");
                if let Ok(mut current) = CURRENT_MEDIA.lock() {
                    *current = Some(flowing);
                }
            }
        }
    }

    // Leg 2: swap onto the 4K film — the ffplay arm's own take-down of a
    // LIVE native session, then a real ffplay.
    if failure.is_none() {
        match load_for_swap(&uhd) {
            None => {
                failure = Some("4K film would not load on the second pass".to_string());
            }
            Some(mut standing) => {
        // The ffplay arm tears the whole standing media down first —
        // this ends the LIVE native session leg 1 started.
        eprintln!("repro leg 2: take_down_pinned_media ends the live session");
        take_down_pinned_media();
        eprintln!(
            "repro leg 2: after take-down is_playing={}",
            video_player::is_playing(),
        );
        eprintln!("repro leg 2: start ffplay for the 4K film");
        let process = start_video_playback(&uhd, 100, 100, 640, 360, 0.0, 0, None);
        eprintln!(
            "repro: ffplay spawned pid={:?}, VIDEO_PID={}",
            process.as_ref().map(|c| c.id()),
            VIDEO_PID.load(Ordering::SeqCst),
        );
        if process.is_none() {
            failure = Some("ffplay would not start for the 4K film".to_string());
        } else {
            standing.video_process = process;
            if let Ok(mut current) = CURRENT_MEDIA.lock() {
                *current = Some(standing);
            }

            // The swap onto the 1080p film: the native arm's own two
            // calls, then the hold's own wait.
            if let Some(mut held) = load_for_swap(&hd) {
                let (w, h) = (held.current_width(), held.current_height());
                eprintln!("repro swap: stop_pinned_player, then play 1080p at {w}x{h}");
                stop_pinned_player();
                eprintln!(
                    "repro swap: after stop VIDEO_PID={}",
                    VIDEO_PID.load(Ordering::SeqCst),
                );
                video_player::play(&hd, w, h, 0, probed_picture(&hd, w, h));
                eprintln!("repro swap: is_playing={}", video_player::is_playing());
                failure = wait_for_first_frame(&mut held, "swap", Duration::from_secs(20));
                if failure.is_none() {
                    eprintln!("repro swap: first frame arrived — the handoff works");
                    // The install: the held file becomes what is on screen,
                    // and every tick after takes into it the way the loop's
                    // own take does once the hold is over.
                    if let Ok(mut current) = CURRENT_MEDIA.lock() {
                        *current = Some(held);
                    } else {
                        failure = Some("CURRENT_MEDIA lock unavailable after swap".to_string());
                    }
                }
            } else {
                failure = Some("1080p film would not load for the swap".to_string());
            }
        }
            }
        }
    }

    // Phase E: rapid re-swap — abandon an OPENING session and begin
    // another at once, the way a second pick supersedes a hold whose
    // first frame has not landed. A Shutdown racing an async open that
    // wedges the engine would hang here, not in the settled path above.
    if failure.is_none() {
        eprintln!("repro rapid: play 1080p and abandon before any frame");
        if let Some(mut opening) = load_for_swap(&hd) {
            let (w, h) = (opening.current_width(), opening.current_height());
            video_player::play(&hd, w, h, 0, probed_picture(&hd, w, h));
            // No take at all: straight to the abandon a newer pick performs.
            stop_video_playback(&mut opening);
            eprintln!(
                "repro rapid: abandoned while opening, playing={}",
                video_player::is_playing(),
            );
            if let Some(mut held) = load_for_swap(&hd) {
                let (w2, h2) = (held.current_width(), held.current_height());
                video_player::play(&hd, w2, h2, 0, probed_picture(&hd, w2, h2));
                eprintln!("repro rapid: replayed, playing={}", video_player::is_playing());
                failure = wait_for_first_frame(&mut held, "rapid", Duration::from_secs(20));
                if failure.is_none() {
                    eprintln!("repro rapid: first frame arrived after abandon");
                    if let Ok(mut current) = CURRENT_MEDIA.lock() {
                        *current = Some(held);
                    }
                }
            } else {
                failure = Some("1080p film would not load for the rapid replay".to_string());
            }
        } else {
            failure = Some("1080p film would not load for the rapid abandon".to_string());
        }
    }
    if failure.is_none() {
        eprintln!("repro sustained: 10 s of per-tick takes into the installed media");
        let started = Instant::now();
        let mut logged = Instant::now();
        let mut frames = 0u32;
        let mut ticks = 0u32;
        while started.elapsed() < Duration::from_secs(10) {
            ticks += 1;
            eprintln!("repro sustained: take #{ticks}");
            let took = if let Ok(mut current) = CURRENT_MEDIA.lock() {
                current.as_mut().map(|media| media.take_native_video_frame())
            } else {
                None
            };
            match took {
                Some(true) => frames += 1,
                Some(false) => {}
                None => {
                    failure = Some("CURRENT_MEDIA lock unavailable during sustained takes".to_string());
                    break;
                }
            }
            eprintln!("repro sustained: take #{ticks} done (frames so far: {frames})");
            if logged.elapsed() >= Duration::from_secs(1) {
                logged = Instant::now();
                eprintln!(
                    "repro sustained: {} ms, {frames} frames, playing={}",
                    started.elapsed().as_millis(),
                    video_player::is_playing(),
                );
            }
            std::thread::sleep(Duration::from_millis(16));
        }
        if failure.is_none() {
            eprintln!("repro sustained: done, {frames} frames over {ticks} takes");
        }
    }

    video_player::stop();
    kill_stray_video_process();
    VIDEO_PID.store(0, Ordering::SeqCst);
    VIDEO_HWND.store(0, Ordering::SeqCst);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous_media.take();
    }
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    VIDEO_HWND.store(previous_hwnd, Ordering::SeqCst);

    assert!(failure.is_none(), "repro failed: {failure:?}");
}

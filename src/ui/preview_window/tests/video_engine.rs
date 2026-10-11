use super::*;

/// The router's question about a video, against the machine it runs on: whether FFmpeg's
/// player is installed here, whether the media engine can decode this file, which of the
/// three answers the router reaches, and what the geometry probe has to say about it.
///
/// It is the probe beside this one's counterpart for pictures, and it is ignored for the same
/// reason — it reads real files and starts a real media stack.
///
/// Run it by hand:
///
/// ```text
/// cargo test video_engine_probe -- --ignored --nocapture
/// ```
#[test]
#[ignore = "reads the files named in RHP_VIDEO_PROBE and starts the media stack"]
fn video_engine_probe() {
    let Ok(list) = std::env::var("RHP_VIDEO_PROBE") else {
        println!("set RHP_VIDEO_PROBE to one or more paths, separated by ';'");
        return;
    };

    // Both halves, because the name reads as a question about whether this machine can play
    // video and is not one: what it answers is whether the media engine is the *only* player
    // here, which is the absence of FFmpeg's rather than the presence of a decoder. Printed
    // together so the two cannot be read as contradicting each other, and it is the first of
    // the two that settles the routing wherever FFmpeg is installed — so where it reads true,
    // everything below it is read for the record rather than because the answer is in doubt.
    println!(
        "ffmpeg installed = {}, so plays_video_natively = {}",
        crate::formats::codecs::ffplay_available(),
        crate::formats::codecs::plays_video_natively()
    );

    for path in list
        .split(';')
        .map(str::trim)
        .filter(|path| !path.is_empty())
        .map(PathBuf::from)
    {
        println!("\n--- {} ---", path.display());
        println!(
            "the preview is shown: {}",
            HoverFacts::read(&path).is_video()
        );

        let started = Instant::now();
        let can_play = video_player::can_play(&path);
        println!(
            "the engine can decode it: {can_play} ({} ms, uncached)",
            started.elapsed().as_millis()
        );
        println!(
            "and the router reads that back: {}",
            video_player::plays(&path)
        );

        let geometry = probe_video_geometry(&path);
        println!(
            "the geometry probe: {}",
            match &geometry {
                ProbedGeometry::Measured(geometry) => format!(
                    "{}x{}, duration {:?}",
                    geometry.width, geometry.height, geometry.duration
                ),
                ProbedGeometry::Unmeasurable => "nothing to measure".to_string(),
            }
        );
        println!(
            "and the name is one the `[video]` list asks the engine for: {}",
            named_as(&path, PreviewType::Videos)
        );
        println!(
            "so a preview of it is played by {}",
            match video_route(&path) {
                VideoRoute::MediaEngine => "the media engine Windows has",
                VideoRoute::Ffplay => "FFmpeg's player",
                VideoRoute::Mpv => "mpv",
                VideoRoute::NoPreview => "nothing at all, so there is no preview",
            }
        );

        // And the half of the question no probe answers: a file the engine can decode is a
        // file a preview of which is nothing at all if the engine cannot draw it either. So
        // the engine is started over the file the way the preview starts one, frames are
        // asked for the way the preview loop asks for them, and what comes back is counted.
        // The volume is nothing, so nothing is heard of this.
        if media_engine_plays(&path) {
            video_player::play(&path, 320, 240, 0, probed_picture(&path, 320, 240));
            println!("a session over it started: {}", video_player::is_playing());

            let mut pixels = Vec::new();
            let mut frames = 0;
            let waited = Instant::now();

            // A file the engine cannot draw has no frame to take however long it is watched,
            // and the app only finds out when its give-up has run — so a file with no frame
            // after the shorter wait is watched for the longer one, which is where the answer
            // `failing_path` gives is due (see `video_player::FIRST_FRAME_GIVE_UP`).
            for limit in [600u64, 3600] {
                while waited.elapsed() < Duration::from_millis(limit) {
                    if video_player::copy_frame_into(&mut pixels).is_some() {
                        frames += 1;
                    }
                    std::thread::sleep(Duration::from_millis(10));
                }

                if frames > 0 {
                    break;
                }
            }

            println!(
                "and the engine drew {frames} frames of it in {} ms",
                waited.elapsed().as_millis()
            );

            // And a seek, which is what dragging a pinned video's bar does — asked for late in
            // the session on purpose, because that is when a bar is dragged, and a seek that
            // is dropped for arriving after a give-up is a bar that does nothing at all.
            if frames > 0 {
                while waited.elapsed() < Duration::from_millis(3300) {
                    video_player::copy_frame_into(&mut pixels);
                    std::thread::sleep(Duration::from_millis(10));
                }

                let target = video_player::duration()
                    .map(|duration| duration * 0.7)
                    .unwrap_or(7.0);

                video_player::seek(target);
                std::thread::sleep(Duration::from_millis(250));
                video_player::copy_frame_into(&mut pixels);

                println!(
                    "and a seek to {target:.2}s, asked {:.1}s into the session, reads back as {:?}",
                    waited.elapsed().as_secs_f32(),
                    video_player::position()
                );

                // And the same seek with the file held where it is, which is the other half of
                // what a bar dragged on a paused pin has to do: the second it is drawn at moves
                // with the hand (the app writes that one down itself), and the picture has to
                // follow it too — which is a frame the engine hands over while it is paused.
                let held_target = target * 0.3;
                video_player::set_paused(true);
                video_player::seek(held_target);

                let held = Instant::now();
                let mut held_frames = 0;
                while held.elapsed() < Duration::from_millis(400) {
                    if video_player::copy_frame_into(&mut pixels).is_some() {
                        held_frames += 1;
                    }
                    std::thread::sleep(Duration::from_millis(10));
                }

                println!(
                    "and a seek to {held_target:.2}s with the file paused drew {held_frames} frames and reads back as {:?}",
                    video_player::position()
                );

                video_player::set_paused(false);
            }

            if frames == 0 {
                let failing = video_player::failing_path();
                println!("and the engine failing at it reads as: {failing:?}");

                if let Some(failing) = failing {
                    video_player::mark_unplayable(&failing);
                }

                println!(
                    "so what plays it reads as {} (plays = {}, route = {:?})",
                    match video_route(&path) {
                        VideoRoute::MediaEngine => "the media engine Windows has",
                        VideoRoute::Ffplay => "FFmpeg's player",
                        VideoRoute::Mpv => "mpv",
                        VideoRoute::NoPreview => "nothing at all, so there is no preview",
                    },
                    video_player::plays(&path),
                    video_route(&path),
                );
            }

            video_player::stop();
        }
    }
}

/// A video that has not been probed yet is a hover that is waiting, so its box is the
/// wait's — and the probe's answer, whatever it is, is what the box becomes: a shape
/// is the shape, and a file with nothing to measure is the box FFmpeg's player is
/// given rather than one the probe is asked for again.
#[test]
fn a_video_that_has_not_been_probed_waits_in_the_waiting_box() {
    // A folder of this module's own, and a name no other test uses: the cache is
    // shared, and what this test puts in it must be its own key.
    let folder = std::env::temp_dir()
        .join("rust-hover-preview-video-tests")
        .join("probe-box");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("waited-for.mp4");
    std::fs::write(&path, b"a file the probe has not seen").expect("a written file");

    let key = VideoGeometryKey {
        path: path.clone(),
        version: file_version(&path),
    };
    assert!(
        cached_video_geometry(&path).is_none(),
        "nothing has probed this file"
    );
    assert_eq!(
        video_box(&path),
        Some((office_preview::WAITING_BOX, office_preview::WAITING_BOX)),
        "an unprobed video is placed as the wait for its probe"
    );

    // The answer the probe gives is what the hover is placed at, and it is read from
    // the cache rather than measured again.
    video_geometry_cache().insert(
        key.clone(),
        ProbedGeometry::Measured(VideoGeometry {
            width: 640,
            height: 360,
            frame_width: 640,
            frame_height: 360,
            crop: None,
            duration: None,
            subtitles: SubtitleStreams::default(),
            sidecar: None,
            derived: None,
            subtitle_codecs: Vec::new(),
            attachment_codecs: Vec::new(),
            subtitle_extraction_failed: false,
        }),
    );
    assert_eq!(video_box(&path), Some((640, 360)));

    // A file the probe could not measure is not a file to probe again: where a player is
    // going to be handed it anyway it is the box that player is given, and where no player
    // will take it at all the hover is dropped rather than laid out.
    video_geometry_cache().insert(key, ProbedGeometry::Unmeasurable);
    assert_eq!(
        video_box(&path),
        match video_route(&path) {
            VideoRoute::Ffplay | VideoRoute::Mpv => Some((1920, 1080)),
            VideoRoute::MediaEngine | VideoRoute::NoPreview => None,
        },
        "an unmeasurable video is placed at the 16:9 box where a player is what will play it, and dropped where no player will"
    );

    let _ = std::fs::remove_file(&path);
}

/// Which engine plays a video is decided by the choice, the fallback switch, the machine and the
/// file's size, and the whole of the rule is `resolve_video_engine` — so this is the rule asked with
/// stub answers rather than with the machine the test happens to run on.
///
/// `Best` prefers mpv wherever the machine has it, whatever the file weighs and whether it could be
/// measured at all; a machine without it answers the file's own rule — the media engine for a film
/// at or below the pixel threshold, and FFmpeg's player for a bigger one or for a film nobody
/// measured. `Hybrid` is that size rule stated plainly, so mpv is not part of it. An explicit choice
/// is itself alone with the fallback off and a walk of the order with it on, and a choice naming an
/// engine the machine has not got is ignored as `Best`, since a row that is greyed names a player
/// that is not here.
#[test]
fn a_video_is_played_by_whichever_engine_the_choice_and_the_machine_leave_to_it() {
    let all_installed = |_| true;
    let all_play = |_| true;
    let nothing_plays = |_| false;
    // A machine without mpv, which is the machine this rule was written for: every other engine is
    // here, and the answers below are the ones it gave before mpv existed.
    let without_mpv = |engine: VideoEngine| engine != VideoEngine::Mpv;
    let big = Some(VIDEO_FFMPEG_ABOVE_PIXELS + 1);
    let small = Some(VIDEO_FFMPEG_ABOVE_PIXELS - 1);

    // `Best` is mpv wherever it is installed, whatever the file weighs and whether it could be
    // measured at all.
    for (pixels, what) in [
        (big, "a film above the threshold"),
        (small, "a film at or below it"),
        (None, "a film nobody measured"),
    ] {
        assert_eq!(
            resolve_video_engine(VideoEngine::Best, true, pixels, all_installed, all_play),
            Some(VideoEngine::Mpv),
            "{what} is played by mpv where the machine has one"
        );
    }

    // And the hybrid, which is the size rule stated plainly: the size of the file decides which
    // engine is preferred, and mpv is not part of it.
    assert_eq!(
        resolve_video_engine(VideoEngine::Hybrid, false, big, all_installed, all_play),
        Some(VideoEngine::Ffmpeg),
        "the hybrid is that rule whether or not the fallback is on, which is what it is for"
    );
    assert_eq!(
        resolve_video_engine(VideoEngine::Hybrid, true, small, all_installed, all_play),
        Some(VideoEngine::Native),
        "and mpv is not in it: a film at or below the threshold is drawn by the media engine"
    );
    assert_eq!(
        resolve_video_engine(VideoEngine::Hybrid, true, None, all_installed, all_play),
        Some(VideoEngine::Ffmpeg),
        "and a film nobody measured is handed over, since the probe that would have weighed it \
         could not read it"
    );
    assert_eq!(
        resolve_video_engine(
            VideoEngine::Hybrid,
            true,
            None,
            |engine| engine == VideoEngine::Native,
            all_play,
        ),
        Some(VideoEngine::Native),
        "unless there is no FFmpeg to hand it to, and the engine is all there is"
    );

    // Where there is no mpv, `Best` is the answer it has always been.
    assert_eq!(
        resolve_video_engine(VideoEngine::Best, true, big, without_mpv, all_play),
        Some(VideoEngine::Ffmpeg),
        "a film above the threshold is handed to FFmpeg's player"
    );
    assert_eq!(
        resolve_video_engine(VideoEngine::Best, true, small, without_mpv, all_play),
        Some(VideoEngine::Native),
        "a film at or below it is drawn by the media engine"
    );

    // An explicit engine is itself, and the size has nothing to do with it.
    assert_eq!(
        resolve_video_engine(VideoEngine::Native, true, big, all_installed, all_play),
        Some(VideoEngine::Native),
        "a chosen media engine plays a large film itself rather than handing it over"
    );
    assert_eq!(
        resolve_video_engine(VideoEngine::Mpv, false, big, all_installed, all_play),
        Some(VideoEngine::Mpv),
        "and a chosen mpv plays it whether or not the fallback is on"
    );

    // A machine with no FFmpeg on it loses the FFmpeg half of every order.
    assert_eq!(
        resolve_video_engine(
            VideoEngine::Best,
            true,
            big,
            |engine| engine == VideoEngine::Native,
            |engine| engine == VideoEngine::Native,
        ),
        Some(VideoEngine::Native),
        "a machine with no FFmpeg on it plays a large film with the media engine as well"
    );

    assert_eq!(
        resolve_video_engine(VideoEngine::Best, true, big, all_installed, nothing_plays),
        None,
        "a file no installed engine will take has no player at all, which is a preview-less video \
         rather than a broken one"
    );

    // The fallback decides what an explicit choice does with a file it cannot play.
    assert_eq!(
        resolve_video_engine(VideoEngine::Native, false, small, all_installed, |engine| {
            engine == VideoEngine::Ffmpeg
        }),
        None,
        "with the fallback off a chosen engine that cannot play the file is the end of the rule, \
         so there is no preview of it"
    );
    assert_eq!(
        resolve_video_engine(VideoEngine::Native, true, small, all_installed, |engine| {
            engine == VideoEngine::Ffmpeg
        }),
        Some(VideoEngine::Ffmpeg),
        "with the fallback on the same file is played by the next engine of the order instead"
    );
    // An explicit choice keeps the meaning it had before mpv existed: its fall-through is the pair
    // it always walked, so mpv is playable here and is still never the answer.
    assert_eq!(
        resolve_video_engine(VideoEngine::Native, true, small, all_installed, |engine| {
            engine == VideoEngine::Ffmpeg || engine == VideoEngine::Mpv
        }),
        Some(VideoEngine::Ffmpeg),
        "an explicit native choice that cannot play the film reaches FFmpeg's player and never mpv"
    );
    // A chosen mpv is the exception: it still falls through to both the others.
    assert_eq!(
        resolve_video_engine(VideoEngine::Mpv, true, small, all_installed, |engine| {
            engine == VideoEngine::Ffmpeg || engine == VideoEngine::Native
        }),
        Some(VideoEngine::Ffmpeg),
        "a chosen mpv that cannot play the film falls through to FFmpeg's player"
    );
    assert_eq!(
        resolve_video_engine(VideoEngine::Mpv, true, small, all_installed, |engine| {
            engine == VideoEngine::Native
        }),
        Some(VideoEngine::Native),
        "and reaches the media engine past it"
    );

    assert_eq!(
        resolve_video_engine(
            VideoEngine::Ffmpeg,
            false,
            big,
            |engine| engine == VideoEngine::Native,
            |engine| engine == VideoEngine::Native,
        ),
        Some(VideoEngine::Native),
        "a choice naming an engine this machine has not got is answered as `Best`"
    );

    assert_eq!(
        resolve_video_engine(VideoEngine::Mpv, false, big, without_mpv, all_play),
        Some(VideoEngine::Ffmpeg),
        "mpv named on a machine without it is answered as `Best` too"
    );
}

/// The media engine is not asked about a file the route settles on FFmpeg's player for, and what
/// that saves is not only a lock: the two extensions the video list shares with the text lists are
/// settled by whether the file holds MPEG-TS packets, which is an open and a read of it. So a film
/// above the threshold on a machine with FFmpeg on it answers its routing from the install alone,
/// and the decoder chain the engine would be asked to build is never built for a file FFmpeg's
/// player will take.
#[test]
fn the_media_engine_is_not_asked_where_the_route_settles_on_ffplay() {
    let engine = resolve_video_engine(
        VideoEngine::Best,
        true,
        Some(VIDEO_FFMPEG_ABOVE_PIXELS + 1),
        |engine| engine != VideoEngine::Mpv,
        |engine| match engine {
            VideoEngine::Ffmpeg => true,
            VideoEngine::Native => {
                panic!("the media engine was asked about a file FFmpeg's player takes anyway")
            }
            VideoEngine::Best | VideoEngine::Hybrid | VideoEngine::Mpv => false,
        },
    );

    assert_eq!(
        engine,
        Some(VideoEngine::Ffmpeg),
        "FFmpeg's player takes the file without the media engine being consulted"
    );
}

/// The geometry probe runs FFmpeg's passes for every choice that may hand the film to FFmpeg's
/// player, so that a film it reaches is given a geometry; an explicit native choice with nothing to
/// hand the film on is the engine alone, and a machine with no `ffplay` runs none of them.
#[test]
fn the_probe_runs_ffmpeg_for_every_choice_that_may_hand_the_film_over() {
    assert!(
        probe_runs_ffmpeg(VideoEngine::Best, true, true, || false),
        "the hybrid probe is what reads the size it decides on"
    );
    assert!(
        probe_runs_ffmpeg(VideoEngine::Hybrid, true, true, || false),
        "and it does the same as `Best`, which is the same rule"
    );
    assert!(
        probe_runs_ffmpeg(VideoEngine::Ffmpeg, false, true, || false),
        "an explicit FFmpeg choice is FFmpeg's player and needs its passes"
    );

    // An explicit native choice is the engine alone: with the fallback off the film is never handed
    // on, so the engine is not even asked, and with it on it is asked only to learn the answer.
    assert!(
        !probe_runs_ffmpeg(VideoEngine::Native, false, true, || {
            panic!("the engine was asked with the fallback off")
        }),
        "a native choice with the fallback off runs none of FFmpeg's passes"
    );
    assert!(
        !probe_runs_ffmpeg(VideoEngine::Native, true, true, || true),
        "and one whose engine takes the film is alone with the fallback on too"
    );
    assert!(
        probe_runs_ffmpeg(VideoEngine::Native, true, true, || false),
        "but a film it will not take is handed on, so FFmpeg's passes run for the player that gets it"
    );

    assert!(
        !probe_runs_ffmpeg(VideoEngine::Best, true, false, || false),
        "a machine with no ffplay runs none of FFmpeg's passes"
    );
}

/// The path a launch runs comes through the same seam the router picks an engine at: the setting's
/// own text where it names a file that is still there, and the first install where it is empty or
/// has gone. Nothing is spawned here — this is the answer a launch is handed, not the launch.
#[test]
fn a_launch_resolves_the_picked_mpv_through_the_same_seam() {
    let installs = vec![
        PathBuf::from("first-install").join("mpv.exe"),
        PathBuf::from("second-install").join("mpv.exe"),
    ];

    assert_eq!(
        video_hw::mpv_program("", &installs),
        Some(installs[0].clone()),
        "an empty setting is the first install found"
    );
    assert_eq!(
        video_hw::mpv_program("   ", &installs),
        Some(installs[0].clone()),
        "a setting that is only spaces is nobody's pick"
    );
    assert_eq!(
        video_hw::mpv_program(r"a-pick-that-has-gone\mpv.exe", &installs),
        Some(installs[0].clone()),
        "a picked path that has gone falls back to the first install"
    );

    let picked = std::env::current_exe().expect("this test's own executable");
    assert_eq!(
        video_hw::mpv_program(&picked.to_string_lossy(), &installs),
        Some(picked),
        "a picked path that is still a file is what a launch runs"
    );
}

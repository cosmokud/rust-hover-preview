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
            VideoRoute::Ffplay => Some((1920, 1080)),
            VideoRoute::MediaEngine | VideoRoute::NoPreview => None,
        },
        "an unmeasurable video is placed at the 16:9 box where FFmpeg's player is what will play it, and dropped where no player will"
    );

    let _ = std::fs::remove_file(&path);
}

/// Which engine plays a video is decided by the choice, the fallback switch, the machine and the
/// file's size, and the whole of the rule is `resolve_video_engine` — so this is the rule asked with
/// stub answers rather than with the machine the test happens to run on.
///
/// `Best` and `Hybrid` are one rule between them and it is the file's own: the media engine for a
/// film at or below the pixel threshold, FFmpeg's player for a bigger one, and the media engine
/// where the size was never read. An explicit choice is itself alone with the fallback off and a
/// walk of the order with it on, and a choice naming an engine the machine has not got is ignored
/// as `Best`, since a row that is greyed names a player that is not here.
#[test]
fn a_video_is_played_by_whichever_engine_the_choice_and_the_machine_leave_to_it() {
    let all_installed = |_| true;
    let all_play = |_| true;
    let nothing_plays = |_| false;
    let big = Some(VIDEO_FFMPEG_ABOVE_PIXELS + 1);
    let small = Some(VIDEO_FFMPEG_ABOVE_PIXELS - 1);

    // The hybrid, which `Best` is: the size of the file decides which engine is preferred.
    assert_eq!(
        resolve_video_engine(VideoEngine::Best, true, big, all_installed, all_play),
        Some(VideoEngine::Ffmpeg),
        "a film above the threshold is handed to FFmpeg's player"
    );
    assert_eq!(
        resolve_video_engine(VideoEngine::Hybrid, false, big, all_installed, all_play),
        Some(VideoEngine::Ffmpeg),
        "and the hybrid is that rule whether or not the fallback is on, which is what it is for"
    );
    assert_eq!(
        resolve_video_engine(VideoEngine::Best, true, small, all_installed, all_play),
        Some(VideoEngine::Native),
        "a film at or below it is drawn by the media engine"
    );
    assert_eq!(
        resolve_video_engine(VideoEngine::Hybrid, true, None, all_installed, all_play),
        Some(VideoEngine::Native),
        "and a film nobody measured is taken for a small one"
    );

    // An explicit engine is itself, and the size has nothing to do with it.
    assert_eq!(
        resolve_video_engine(VideoEngine::Native, true, big, all_installed, all_play),
        Some(VideoEngine::Native),
        "a chosen media engine plays a large film itself rather than handing it over"
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
        |_| true,
        |engine| match engine {
            VideoEngine::Ffmpeg => true,
            VideoEngine::Native => {
                panic!("the media engine was asked about a file FFmpeg's player takes anyway")
            }
            VideoEngine::Best | VideoEngine::Hybrid => false,
        },
    );

    assert_eq!(
        engine,
        Some(VideoEngine::Ffmpeg),
        "FFmpeg's player takes the file without the media engine being consulted"
    );
}

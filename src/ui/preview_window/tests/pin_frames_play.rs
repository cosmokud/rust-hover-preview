//! A pinned picture that plays: the bar it draws, the frames its track names, and the pace its
//! own button sets.

use super::*;

/// The kinds a bar with a track of frames is drawn for, and the ones it is not: a picture that
/// plays carries one, and a held sequence this app hands to the engine — a HEIF or an AVIF, or
/// an animated JXL, none of which is decoded frame by frame here — does not.
#[test]
fn only_the_pictures_this_app_decodes_itself_carry_a_bar_of_frames() {
    for kind in [
        MediaType::AnimatedGif,
        MediaType::AnimatedApng,
        MediaType::AnimatedWebP,
    ] {
        assert!(pin_transport_kind(Some(kind)), "{kind:?} carries a bar");
        assert!(
            is_played_picture(Some(kind)),
            "{kind:?} is the app's own frames"
        );
        assert!(
            !pin_transport_live(Some(kind)),
            "{kind:?} has no player behind it, so its button is not drawn"
        );
    }

    for kind in [MediaType::AnimatedHeif, MediaType::AnimatedJxl] {
        assert!(!is_played_picture(Some(kind)), "{kind:?} is the engine's");
    }

    // A film keeps its own bar and its own button, and a picture, a page and a sound carry none.
    assert!(pin_transport_live(Some(MediaType::Video)));
    assert!(pin_transport_live(Some(MediaType::NativeVideo)));
    assert!(!pin_transport_kind(Some(MediaType::Audio)));
    assert!(!pin_transport_kind(Some(MediaType::StaticImage)));
    assert!(!pin_transport_kind(None));
}

/// A picture that plays is pinned with the two things its bar is drawn from read off its own
/// frames: how long its run is, and the pace it keeps.
#[test]
fn a_pinned_picture_takes_its_length_and_its_pace_from_its_own_frames() {
    let mut pin = overlay_pin((100, 100, 500, 400), PinChrome::always());
    pin.transport_bar = true;
    pin.transport_animated = true;
    pin.fps = PinFps {
        detected: 12.5,
        wanted: None,
        ..PinFps::default()
    };
    pin.transport = PinTransport {
        duration: Some(2.4),
        ..PinTransport::default()
    };

    assert_eq!(
        pin.fps.pace(),
        13,
        "the pace the slider opens at is the file's own, in whole frames a second"
    );
    assert_eq!(
        pin_duration(&pin.transport),
        Some(2.4),
        "and its length is the run of its own frames"
    );

    // A hand that moved the slider is the pace the picture plays at, and the file's own is what is
    // under it — a picture put back to its own pace on every scrub would be a slider that snapped.
    pin.fps.wanted = Some(30);
    assert_eq!(pin.fps.pace(), 30);
}

/// The pace a hand asks for scales every frame's own delay rather than replacing it: a picture paced
/// at twice its own rate halves every delay, and one paced at half of it doubles them, so the
/// relationship between the frames stays where the file's author put it.
#[test]
fn the_pace_a_hand_asks_for_scales_every_delay_the_file_gives() {
    let mut playback = AnimationPlayback {
        detected: 10.0,
        wanted: None,
    };

    assert_eq!(
        playback.delay_scale(),
        1.0,
        "a picture nobody has touched plays at its own pace"
    );
    assert_eq!(
        effective_frame_delay_ms(&MediaType::AnimatedGif, 100, playback),
        100
    );

    playback.wanted = Some(20);
    assert!(
        (playback.delay_scale() - 0.5).abs() < 0.001,
        "twice the file's own pace is half of every wait"
    );
    assert_eq!(
        effective_frame_delay_ms(&MediaType::AnimatedGif, 100, playback),
        50
    );
    assert_eq!(
        effective_frame_delay_ms(&MediaType::AnimatedGif, 200, playback),
        100,
        "and a frame authored to wait twice as long still waits twice as long"
    );

    playback.wanted = Some(5);
    assert_eq!(
        effective_frame_delay_ms(&MediaType::AnimatedGif, 100, playback),
        200,
        "half the pace is twice the wait"
    );

    // And a picture whose own pace was never read — a file whose container could not be walked —
    // plays at the delays it was given, whatever a hand asked for: a pace has no base to scale
    // against, and a guessed base would be a picture played at a rate nobody chose.
    let unknown = AnimationPlayback {
        detected: 0.0,
        wanted: Some(30),
    };
    assert_eq!(unknown.delay_scale(), 1.0);
    assert_eq!(
        effective_frame_delay_ms(&MediaType::AnimatedGif, 100, unknown),
        100
    );
}

/// The frame a share of the bar names is a place in the file rather than a second of a player.
#[test]
fn a_share_of_the_track_names_a_frame_of_the_file() {
    let two_frames = crate::formats::animation_frames::FrameTimeline::of_delays(&[100, 100]);

    assert_eq!(two_frames.frames(), 2);
    assert_eq!(
        two_frames.index_at_share(0.0),
        0,
        "the bar's start is the first"
    );
    assert_eq!(two_frames.index_at_share(0.5), 1, "half way is the second");
    assert_eq!(two_frames.index_at_share(1.0), 1, "and its end is the last");

    // A run of a single frame is not one a track can be dragged along, which is what the press's
    // own gate refuses (see `pin_frame_share`).
    let one_frame = crate::formats::animation_frames::FrameTimeline::of_delays(&[100]);
    assert_eq!(one_frame.frames(), 1);
    assert_eq!(one_frame.index_at_share(1.0), 0);
}

/// **A hand on the track paints the frame it named and never replaces the media**: a frame the
/// scrub holds is shown beside the run in hand rather than in place of it, so the playhead the
/// picture plays from stays where the file left it — and a release drops the frame it was shown
/// and plays on from the nearest frame the file has given it (see `show_the_frame_of` and
/// `settle_the_picture_on`).
#[test]
fn a_hand_on_the_track_paints_the_frame_it_named_without_replacing_the_media() {
    // The media and the runs beside it are one set for the whole process, so the tests that stand
    // their own media in it take the room to themselves (see `the_room_to_itself`).
    let _room = the_room_to_itself();
    let folder = std::env::temp_dir().join("pin-scrub-shown-tests");
    std::fs::create_dir_all(&folder).expect("a fixture folder");
    let path = folder.join("run.gif");
    std::fs::write(&path, b"GIF89a").expect("a fixture file");

    let a_run = |indices: &[usize]| {
        let frame = |file_index: usize| {
            let mut frame = ImageFrame::new(vec![0u8; 4], 1, 1, 100);
            frame.file_index = file_index;
            Arc::new(frame)
        };
        let mut media = static_image_media(frame(indices[0]), MediaType::AnimatedGif);
        media.frames = indices.iter().copied().map(frame).collect();
        media
    };
    let stand = |media: MediaData| {
        *CURRENT_MEDIA.lock().expect("the media's lock") = Some(media);
    };
    let playhead = || {
        CURRENT_MEDIA.lock().ok().and_then(|media| {
            media
                .as_ref()
                .and_then(|media| media.current_frame_of_file())
        })
    };
    let shown = || {
        CURRENT_MEDIA.lock().ok().and_then(|media| {
            media.as_ref().and_then(|media| {
                media
                    .shown_for_the_hand
                    .as_ref()
                    .map(|frame| frame.file_index)
            })
        })
    };
    let previous = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    pin_scrub::forget();

    // A frame of the run in hand is a playhead put down on it, which is the cheapest answer there
    // is and the one place a hand moves the picture itself.
    stand(a_run(&[40, 41, 42]));
    assert!(show_the_frame_of(&path, 41), "a frame of the run in hand");
    assert_eq!(playhead(), Some(41), "and the playhead is on it");
    assert_eq!(shown(), None, "with no frame painted beside the run");

    // A frame this pin decoded before is that frame *shown*, and the media is untouched: the run
    // of the file is not walked again for a place this pin has already been, and the playhead the
    // picture plays from stays where it was.
    pin_scrub::remember(&path, &a_run(&[10, 11, 12]).frames);
    stand(a_run(&[40, 41, 42]));
    assert!(
        show_the_frame_of(&path, 11),
        "a frame the scrub holds is a frame shown"
    );
    assert_eq!(shown(), Some(11), "and it is the frame painted");
    assert_eq!(
        playhead(),
        Some(40),
        "while the playhead the picture plays from is left where the file had it"
    );
    assert_eq!(
        CURRENT_MEDIA
            .lock()
            .ok()
            .and_then(|media| media.as_ref().map(|media| media.frames.len())),
        Some(3),
        "and the run in hand is the whole of what the media holds"
    );

    // And a frame nothing at all holds is the nearest frame of the run in hand, which is the one
    // place a hand moves the picture itself: there is nothing to paint beside it, so the playhead is
    // put down where the file has got to and the caller is told the exact frame is not there yet.
    pin_scrub::forget();
    stand(a_run(&[40, 41, 42]));
    assert!(
        !show_the_frame_of(&path, 300),
        "a frame neither the run nor the scrub holds is not shown"
    );
    assert_eq!(shown(), None, "with nothing painted beside the run");
    assert_eq!(
        playhead(),
        Some(42),
        "the picture is taken to the nearest frame that is here instead"
    );

    assert!(
        !show_the_frame_of(&path, 5),
        "and a place nothing holds is answered the same way"
    );
    assert_eq!(shown(), None);
    assert_eq!(playhead(), Some(40), "the nearest frame of the run");

    // **A release settles on the frame in hand**: the frame painted beside the run is dropped, and
    // the playhead goes to the nearest frame the file has given the picture — nothing is decoded,
    // and no media is installed for it.
    settle_the_picture_on(5);
    assert_eq!(shown(), None, "the hand's frame goes with the hand");
    assert_eq!(
        playhead(),
        Some(40),
        "and the picture plays on from the nearest frame it holds"
    );

    pin_scrub::forget();
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous;
    }
}

/// **A picture the button at the start of its bar holds stands on the frame it was held on**, and
/// letting go plays on from that frame rather than from where its own clock had got to: the walk
/// is not asked to move a held picture at all, and the catch-up a late playhead is given is not
/// run on through the frames a hold lasted (see `hold_the_picture_itself` and `advance_frame`).
#[test]
fn a_picture_held_by_its_button_stands_on_its_frame_and_plays_on_from_it() {
    let frames = |delays: &[u32]| {
        let media = MediaData {
            frames: delays
                .iter()
                .map(|delay| Arc::new(ImageFrame::new(vec![0u8; 4], 1, 1, *delay)))
                .collect(),
            shared_frames: None,
            all_frames_loaded: Some(Arc::new(AtomicBool::new(true))),
            current_frame: 0,
            last_frame_time: Instant::now(),
            media_type: MediaType::AnimatedGif,
            stream_cancel: None,
            video_process: None,
            loading_start: None,
            text_state: None,
            audio_scale: None,
            animation: AnimationPlayback {
                detected: 60.0,
                wanted: None,
            },
            paused: false,
            shown_for_the_hand: None,
        };
        media
    };

    let mut media = frames(&[16, 16, 16, 16]);

    // Held: the walk leaves it where it is, however long the hold lasts.
    media.paused = true;
    media.last_frame_time = Instant::now() - Duration::from_secs(5);
    assert!(
        !media.advance_frame(),
        "a picture a hand paused does not move"
    );
    assert_eq!(media.current_frame, 0, "and stands where it was held");

    // Let go: it plays on from that frame rather than running through the ones a hold would have
    // played, which is what re-basing the clock on the release is for.
    media.paused = false;
    media.last_frame_time = Instant::now() - Duration::from_millis(20);
    assert!(media.advance_frame(), "and one nobody holds plays on");
    assert_eq!(
        media.current_frame, 1,
        "one frame past the one it was held on rather than the whole run"
    );
}

/// **A run that has played to its end wraps round to the frame carrying the file's first frame**,
/// which is the loop a picture that fits the pin's budget plays for as long as it is up: the run is
/// held whole, plays from beginning to end, and starts again at the file's first frame whatever
/// place that frame is at in the list (see `MediaData::frame_of_file_start` and `advance_frame`).
#[test]
fn a_run_that_has_played_to_its_end_and_holds_its_head_wraps_to_the_files_first_frame() {
    let a_run = |indices: &[usize]| {
        let frame = |file_index: usize| {
            let mut frame = ImageFrame::new(vec![0u8; 4], 1, 1, 100);
            frame.file_index = file_index;
            Arc::new(frame)
        };
        let mut media = static_image_media(frame(indices[0]), MediaType::AnimatedGif);
        media.frames = indices.iter().copied().map(frame).collect();
        media.all_frames_loaded = Some(Arc::new(AtomicBool::new(true)));
        media
    };

    // A decode begun part-way through the file answers the file's tail and then its head, so the
    // frame carrying the file's first frame is not the list's own first one.
    let mut media = a_run(&[180, 181, 0, 1]);
    media.current_frame = 3;
    // The clock owes one frame and not a stretch of them: a playhead left late is caught up through
    // several frames at once, and what wraps is the one it lands on (see `advance_frame`).
    media.last_frame_time = Instant::now() - Duration::from_millis(100);
    assert!(media.advance_frame(), "a run at its end plays on");
    assert_eq!(
        media.current_frame_of_file(),
        Some(0),
        "and the picture starts again at the frame the file begins with"
    );
}

/// **A run whose head was never delivered holds its last frame instead of replaying its own tail**,
/// which is the cycle a cancelled head pass used to turn into: the run holds the file's tail and no
/// frame carrying `0`, so there is no start to wrap to and nothing wraps (see `advance_frame`).
#[test]
fn a_run_whose_head_was_never_delivered_holds_its_last_frame_instead_of_replaying_it() {
    let mut media = static_image_media(
        {
            let mut frame = ImageFrame::new(vec![0u8; 4], 1, 1, 100);
            frame.file_index = 180;
            Arc::new(frame)
        },
        MediaType::AnimatedGif,
    );
    media.frames = (180..200)
        .map(|file_index| {
            let mut frame = ImageFrame::new(vec![0u8; 4], 1, 1, 100);
            frame.file_index = file_index;
            Arc::new(frame)
        })
        .collect();
    media.all_frames_loaded = Some(Arc::new(AtomicBool::new(true)));
    media.current_frame = media.frames.len() - 1;
    media.last_frame_time = Instant::now() - Duration::from_millis(500);

    assert!(
        !media.advance_frame(),
        "a run with no start holds where it is"
    );
    assert_eq!(
        media.current_frame_of_file(),
        Some(199),
        "the last frame of the tail is the frame it holds"
    );

    // And it holds it however long it is asked again: nothing here walks the tail round.
    media.last_frame_time = Instant::now() - Duration::from_millis(500);
    assert!(!media.advance_frame());
    assert_eq!(media.current_frame_of_file(), Some(199));
}

/// **A run whose stream was cancelled holds its last frame as well**, which is the shape the fault
/// took: a decode begun at a place answered the file's tail, its head pass was stopped part-way, and
/// the flag the decoder leaves behind on its way out says the file is fully loaded — so the walk
/// takes the wrap arm with a run that holds no head at all (see `cancel_background_work` and
/// `MediaData::frame_of_file_start`).
#[test]
fn a_run_whose_stream_was_cancelled_holds_its_last_frame_and_never_returns_to_its_tail() {
    let mut media = static_image_media(
        {
            let mut frame = ImageFrame::new(vec![0u8; 4], 1, 1, 100);
            frame.file_index = 180;
            Arc::new(frame)
        },
        MediaType::AnimatedGif,
    );
    media.frames = (180..200)
        .map(|file_index| {
            let mut frame = ImageFrame::new(vec![0u8; 4], 1, 1, 100);
            frame.file_index = file_index;
            Arc::new(frame)
        })
        .collect();
    media.shared_frames = Some(Arc::new(Mutex::new(StreamedFrames {
        queue: VecDeque::new(),
        released: false,
        decoded: false,
    })));
    media.all_frames_loaded = Some(Arc::new(AtomicBool::new(true)));
    media.stream_cancel = Some(Arc::new(AtomicBool::new(false)));
    media.current_frame = media.frames.len() - 1;
    media.last_frame_time = Instant::now() - Duration::from_millis(500);

    media.cancel_background_work();

    assert!(
        !media.advance_frame(),
        "a run whose head pass was cancelled holds where it is"
    );
    assert_eq!(
        media.current_frame_of_file(),
        Some(199),
        "on the last frame of the tail, and never back at its first"
    );

    media.last_frame_time = Instant::now() - Duration::from_millis(500);
    assert!(!media.advance_frame());
    assert_eq!(
        media.current_frame_of_file(),
        Some(199),
        "so the frames the picture would have cycled through are not played again"
    );
}

/// **A pinned picture's held run is its own budget, and a hover keeps the constant**: the room a
/// release gives frames back in is the tray's `Performance → Memory → Animated WEBP` for a picture
/// that has been pinned — which is what lets the player hold a run whole and serve a drag from it —
/// and the module's constant for a hover, which is never dragged (see `MediaData::retained_room`).
#[test]
fn a_pinned_pictures_retained_room_is_its_own_budget_and_a_hover_keeps_the_constant() {
    let a_picture = |kind: MediaType| {
        static_image_media(Arc::new(ImageFrame::new(vec![0u8; 4], 1, 1, 0)), kind)
    };
    let budget = crate::config::config::animated_webp_budget_bytes();
    let previous_pin = take_pin_for_a_test();

    // A hover: a picture that plays, held to the constant every hover's animation is held to.
    assert_eq!(
        a_picture(MediaType::AnimatedGif).retained_room(),
        ANIMATION_RETAINED_BYTES,
        "a hover's run keeps the constant it always had"
    );

    // A pin: the same picture, held to the budget the tray's own row names.
    let mut pin = overlay_pin((100, 100, 400, 340), PinChrome::always());
    pin.path = std::env::temp_dir()
        .join("pin-index-tests")
        .join("run.webp");
    stand_pin(Some(pin));
    assert_eq!(
        a_picture(MediaType::AnimatedWebP).retained_room(),
        budget,
        "a pinned picture's room is the budget set for it"
    );
    assert_eq!(
        a_picture(MediaType::AnimatedGif).retained_room(),
        budget,
        "whatever kind of run it is"
    );
    stand_pin(None);

    // And a still is a picture that does not play: it is held whole either way, so its room is
    // the constant and never the budget of a run it does not have.
    assert_eq!(
        a_picture(MediaType::StaticImage).retained_room(),
        ANIMATION_RETAINED_BYTES
    );

    stand_pin(previous_pin);
}

/// **A release seeks after a resize, and it seeks exactly, whether or not the file has finished
/// streaming into the run.** A box change leaves the whole of the run the pin was playing, with the
/// frames its own decode answers with taking their places in it (see `keep_animation_place`), so the
/// place a hand lets go on is a frame the run holds — and a release that left the playhead alone
/// while the stream was still filling left the picture playing on from wherever the resize had put
/// it, which is a release that sought nowhere at all.
#[test]
fn a_release_after_a_resize_seeks_to_the_frame_the_hand_left() {
    // The media is one slot for the whole process, so the tests that stand their own in it take the
    // room to themselves (see `the_room_to_itself`).
    let _room = the_room_to_itself();
    let frame = |file_index: usize, side: u32| {
        let mut frame = ImageFrame::new(vec![0u8; 4], side, side, 100);
        frame.file_index = file_index;
        Arc::new(frame)
    };
    let previous = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    // The run the box change left: the file's places, the frame the pin was on drawn at the box the
    // window is now, and a stream still filling the rest of it.
    let mut media = static_image_media(frame(140, 8), MediaType::AnimatedGif);
    media.frames = (0..300).map(|index| frame(index, 4)).collect();
    media.frames[140] = frame(140, 8);
    media.current_frame = 140;
    media.all_frames_loaded = Some(Arc::new(AtomicBool::new(false)));
    media.shared_frames = Some(Arc::new(Mutex::new(StreamedFrames {
        queue: VecDeque::new(),
        released: false,
        decoded: false,
    })));
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }

    settle_the_picture_on(300 - 1);
    let landed = CURRENT_MEDIA.lock().ok().and_then(|media| {
        media.as_ref().map(|media| {
            (
                media.current_frame_of_file(),
                media.current_width(),
                media.shown_for_the_hand.is_some(),
            )
        })
    });
    assert_eq!(
        landed,
        Some((Some(300 - 1), 4, false)),
        "the place the hand let go on is the place the picture carries on from, and nothing is \
         painted beside the media any longer"
    );

    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous;
    }
}

/// **And a hand that names a place mid-drag is shown that place as soon as the run holds it**, rather
/// than only once the whole of the file has streamed into the run: the decode a box change asks for
/// is answered a frame at a time, and every place the file has is a place the run the resize left
/// still holds (see `show_the_frame_of`).
#[test]
fn a_hand_on_the_track_is_shown_the_place_it_names_while_the_stream_is_still_filling() {
    // The media is one slot for the whole process, so the tests that stand their own in it take the
    // room to themselves (see `the_room_to_itself`).
    let _room = the_room_to_itself();
    let frame = |file_index: usize, side: u32| {
        let mut frame = ImageFrame::new(vec![0u8; 4], side, side, 100);
        frame.file_index = file_index;
        Arc::new(frame)
    };
    let previous = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let mut media = static_image_media(frame(140, 8), MediaType::AnimatedGif);
    media.frames = (0..300).map(|index| frame(index, 4)).collect();
    media.frames[140] = frame(140, 8);
    media.current_frame = 140;
    media.all_frames_loaded = Some(Arc::new(AtomicBool::new(false)));
    media.shared_frames = Some(Arc::new(Mutex::new(StreamedFrames {
        queue: VecDeque::new(),
        released: false,
        decoded: false,
    })));
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }

    let path = std::env::temp_dir()
        .join("pin-scrub-shown-tests")
        .join("run.gif");
    std::fs::create_dir_all(path.parent().expect("a folder")).expect("a fixture folder");
    std::fs::write(&path, b"GIF89a").expect("a fixture file");
    pin_scrub::forget();
    pin_index::forget();

    let shown = |target: usize| {
        let exact = show_the_frame_of(&path, target);
        let answer = CURRENT_MEDIA.lock().ok().and_then(|media| {
            media.as_ref().map(|media| {
                (
                    media.current_frame_of_file(),
                    media.current_frame().map(|frame| frame.file_index),
                    media
                        .shown_for_the_hand
                        .as_ref()
                        .map(|frame| frame.file_index),
                )
            })
        });
        (exact, answer)
    };

    assert_eq!(
        shown(10),
        (true, Some((Some(10), Some(10), None))),
        "a place the run holds is the playhead put down on it, without waiting for the rest of the \
         file"
    );
    assert_eq!(
        shown(200),
        (true, Some((Some(200), Some(200), None))),
        "and one from the far end of the file, which the stream has not reached yet"
    );

    pin_scrub::forget();
    pin_index::forget();
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = previous;
    }
}

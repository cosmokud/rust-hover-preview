use super::*;

/// One share for every kind, so a test can say what a single kind's arm does by changing one
/// setting and leaving the other twelve where they are.
///
/// Every scale is distinct on purpose: a test that set them all alike could not tell an
/// arm that read the wrong field from one that read the right one, and the whole of what
/// this table decides is *which* setting each kind follows.
fn one_scale_per_kind() -> HoverScales {
    HoverScales {
        picture: PreviewScale::Percent(11),
        animated: PreviewScale::Percent(12),
        video: PreviewScale::Percent(13),
        ebook: PreviewScale::Percent(14),
        document: PreviewScale::Percent(15),
        vector: PreviewScale::Percent(16),
        design: PreviewScale::Percent(17),
        font: PreviewScale::Percent(18),
    }
}

/// Every kind follows the one setting it names, and the share that comes back says which.
///
/// The table is the reason this function exists at all. It was one share written into each
/// arm of a chain, and the chain's output disagreed with the loader's: a picture laid out
/// at the vector share is a preview placed against a box nothing will ever draw it into.
/// Writing the setting into the arm and reading it back through the placement the caller
/// uses is what makes "every kind follows its own" a claim rather than a hope — and the
/// shares are all distinct precisely so that an arm reading a neighbour's field fails.
///
/// Every share here is below `100%`, so each kind is answered with a *reduced fit* — a
/// share of the display's room rather than of its own size — and the two are told apart by
/// the number rather than by the variant, which is the distinction a kind's arm exists to
/// draw.
#[test]
fn every_kind_follows_the_one_setting_it_names() {
    let scales = one_scale_per_kind();

    for (kind, setting, expected) in [
        (PreviewType::Ebook, "ebook_scale", scales.ebook),
        (PreviewType::Libre, "document_scale", scales.document),
        (PreviewType::Calibre, "ebook_scale", scales.ebook),
        (PreviewType::Design, "design_scale", scales.design),
        (PreviewType::Vector, "vector_scale", scales.vector),
        (PreviewType::Fonts, "font_scale", scales.font),
    ] {
        let path = std::env::temp_dir().join(format!("scale-of-kind-{setting}"));
        // The path does not have to exist: these arms ask the file nothing, and the share is
        // the display's room either way.
        let hover = HoverFacts::read(&path);
        assert_eq!(
            scale_of_kind(kind, &hover, scales),
            fit_reduced(expected),
            "{kind:?} follows {setting} and nothing else"
        );
    }
}

/// A page takes the display's room and a bitmap takes a share of its own size, and a fit is
/// the only place the two rules come apart.
///
/// They are not a preference. A page has no pixels of its own — it is laid out at whatever
/// box it is given, so the display's room is free quality — while a bitmap is only ever as
/// good as the pixels it holds, and stretching a worksheet's corner over a display produces
/// a preview that is larger and no more readable. So `Fit to Screen` means the whole of the
/// room for a page and the picture at the size it is for a bitmap, and the difference is the
/// whole of what `bitmap_at_display_scale` is for. A reduced fit is the one share the two
/// rules read differently: half the room for a page, half the picture for a bitmap.
///
/// Asserted against the two rules rather than against the table of kinds above, because they
/// live in different functions and the bug this rules out is one being reached where the
/// other belongs: a document handed a bitmap's share is placed against a box nothing is
/// going to draw it into.
#[test]
fn a_page_takes_the_display_and_a_bitmap_takes_its_own_size() {
    for configured in [
        PreviewScale::Percent(100),
        PreviewScale::Percent(250),
        PreviewScale::FitToScreen,
        PreviewScale::FitToScreenReduced(40),
    ] {
        assert_eq!(
            fit_reduced(configured),
            match configured {
                PreviewScale::Percent(percent) if percent < 100 => {
                    PreviewScale::FitToScreenReduced(percent)
                }
                // A reduced fit the configuration itself holds is the
                // answer already, kept as it is.
                PreviewScale::FitToScreenReduced(_) => configured,
                _ => PreviewScale::FitToScreen,
            },
            "a page at {configured:?} is a share of the display's room"
        );
        assert_eq!(
            bitmap_at_display_scale(configured),
            match configured {
                PreviewScale::Percent(percent) => PreviewScale::Percent(percent),
                // Which is the half of the rule a fit is: the whole of the display as the
                // picture's own size, and never more of the picture than it has — a
                // reduced fit included, whose share is of that own size.
                PreviewScale::FitToScreenReduced(percent) => PreviewScale::Percent(percent),
                PreviewScale::FitToScreen => PreviewScale::Percent(100),
            },
            "and a bitmap at {configured:?} is a share of its own, so a fit never enlarges \
                 one to fill a display"
        );
    }

    // And the one place the two rules are genuinely different, named on both sides so a
    // reader can see that these are not the same rule written twice.
    assert_ne!(
        fit_reduced(PreviewScale::FitToScreen),
        bitmap_at_display_scale(PreviewScale::FitToScreen),
        "a fit is the room for a page and the picture's own size for a bitmap"
    );
}

/// A picture keeps the picture's share, and an animation is the one thing that asks the file
/// — and it asks it last, because the answer cannot matter for any other kind.
///
/// The ordering is a cost decision with a bug in it if it is wrong: `image_is_animated` is
/// the only arm of the table that reads the file, so a kind answered by one of its own arms
/// and then asked again would pay two file reads for one hover, on the thread that pumps
/// this window's messages. The cheap exit is that a user who has given animations the same
/// size as pictures never pays the probe at all — which is the state a fresh install is in,
/// because both settings start at `100%`.
#[test]
fn a_picture_asks_the_file_about_moving_only_when_the_two_sizes_differ() {
    let animated = std::env::temp_dir().join("scale-of-kind-animated.gif");
    write_test_gif(&animated, 2);
    let still = std::env::temp_dir().join("scale-of-kind-still.gif");
    write_test_gif(&still, 1);

    let differing = HoverScales {
        picture: PreviewScale::Percent(100),
        animated: PreviewScale::Percent(25),
        ..one_scale_per_kind()
    };
    assert_eq!(
        scale_of_kind(PreviewType::Images, &HoverFacts::read(&animated), differing),
        PreviewScale::Percent(25),
        "a file whose own head says it moves is drawn at the animation's share"
    );
    assert_eq!(
        scale_of_kind(PreviewType::Images, &HoverFacts::read(&still), differing),
        PreviewScale::Percent(100),
        "and one holding a single frame is a picture like any other"
    );

    // The same two files where the probe cannot change the answer, which is what makes the
    // question affordable to ask on every hover.
    let alike = HoverScales {
        picture: PreviewScale::Percent(100),
        animated: PreviewScale::Percent(100),
        ..one_scale_per_kind()
    };
    for path in [&animated, &still] {
        assert_eq!(
            scale_of_kind(PreviewType::Images, &HoverFacts::read(path), alike),
            PreviewScale::Percent(100),
            "{} follows the picture's share, and the file is not read to find that out",
            path.display()
        );
    }
}

// The lock the tests below take over the preview's own process-wide state is
// `pin_window::PIN_TESTS_ONE_AT_A_TIME`, declared there beside the slot it guards and
// documented at its definition: there is one lock for that state, not one per module,
// because a second lock over the same slot is no lock at all.

/// A little GIF of one pixel, written with `frames` frame blocks in it: an
/// animation when there is more than one frame and a still picture when there is
/// one, which is the whole of what tells the two apart. The bytes are a real file
/// rather than a shape, because that is what the probe and the decoder both read —
/// two colours, one pixel, and the frames of it.
fn write_test_gif(path: &Path, frames: usize) {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(b"GIF89a");
    bytes.extend_from_slice(&[1, 0, 1, 0]); // one pixel square
    bytes.push(0xF0); // a global colour table, two entries
    bytes.push(0); // background colour index
    bytes.push(0); // pixel aspect ratio
    bytes.extend_from_slice(&[0, 0, 0, 255, 255, 255]); // black and white

    for _ in 0..frames {
        bytes.push(0x2C); // an image descriptor
        bytes.extend_from_slice(&[0, 0, 0, 0, 1, 0, 1, 0, 0]); // at 0,0, one pixel
        bytes.push(2); // the LZW code size
        bytes.extend_from_slice(&[2, 0x4C, 0x01, 0]); // one block of one code, then the end
    }

    bytes.push(0x3B); // the trailer
    std::fs::write(path, bytes).expect("a written GIF");
}

/// A video's size is a setting of its own: a video the probe has answered for keeps
/// the share `video_scale` names whatever the pictures beside it are drawn at. A
/// video the probe has not answered for is not laid out at any share yet — what is
/// on screen is the wait for the probe — so both settings leave that hover in the
/// spinner's own box (see `video_probe_due`).
#[test]
fn a_video_follows_its_own_scale() {
    let video = PathBuf::from(r"C:\clips\holiday.mp4");
    let unprobed = PathBuf::from(r"C:\clips\not-probed-yet.mp4");

    assert_eq!(
        effective_preview_scale(
            &unprobed,
            HoverScales {
                video: PreviewScale::Percent(50),
                picture: PreviewScale::Percent(400),
                ..hover_scales()
            }
        ),
        PreviewScale::Percent(100),
        "a video the probe has not answered for waits in the spinner's own box"
    );

    // The probe's answer is held per file and version, which is the state a hover
    // the probe has already answered for is in by the time its replay is laid out.
    video_geometry_cache().insert(
        VideoGeometryKey {
            path: video.clone(),
            version: file_version(&video),
        },
        ProbedGeometry::Measured(VideoGeometry {
            width: 1920,
            height: 1080,
            frame_width: 1920,
            frame_height: 1080,
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

    assert_eq!(
        effective_preview_scale(
            &video,
            HoverScales {
                video: PreviewScale::Percent(50),
                picture: PreviewScale::Percent(400),
                ..hover_scales()
            }
        ),
        PreviewScale::Percent(50),
        "a measured video follows its own share, not the picture's"
    );

    assert_eq!(
        effective_preview_scale(
            &video,
            HoverScales {
                video: PreviewScale::FitToScreen,
                ..hover_scales()
            }
        ),
        PreviewScale::FitToScreen,
        "and taking the whole of its own size is a share it is offered too"
    );
}

/// A hovered video at a fit fills the room it is given, and at a share of its own size takes that
/// share — the two bases the `Scaling → Video` menu offers, told apart.
///
/// A video is not a bitmap for this question, whatever engine plays it: what its preview is, once
/// its player has a window up, is a box the film is drawn into — the media engine's frames landing
/// in this app's own window, FFmpeg's or mpv's player putting up a window of its own at the same
/// box — so a fit is the size the film is watched at, up as well as down. It is the fit a pinned
/// window has always taken, and the assertion below that the two roads agree is the whole of what
/// changed: a film smaller than the display used to sit at its own size on a hover while a pin
/// maximized over the same display enlarged it.
///
/// The shares below the fit are the file's own size for a video as they are for a picture, which is
/// the half of this that did not change: `By Own Size 150%` enlarges because a user asked for it,
/// and `By Screen 50%` is half of the film rather than half of the room.
#[test]
fn a_hovered_video_fills_the_room_at_a_fit_and_takes_its_own_size_below_one() {
    let small = PathBuf::from(r"C:\clips\holiday.mp4");
    let large = PathBuf::from(r"C:\clips\concert-4k.mp4");
    let room = (3840u32, 2160u32);

    for (path, shape) in [(small.clone(), (1920u32, 1080u32)), (large, (7680, 4320))] {
        video_geometry_cache().insert(
            VideoGeometryKey {
                path: path.clone(),
                version: file_version(&path),
            },
            ProbedGeometry::Measured(VideoGeometry {
                width: shape.0,
                height: shape.1,
                frame_width: shape.0,
                frame_height: shape.1,
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
    }

    let at_fit = HoverScales {
        video: PreviewScale::FitToScreen,
        ..hover_scales()
    };
    let hover = HoverFacts::read(&small);
    let hover_large = HoverFacts::read(Path::new(r"C:\clips\concert-4k.mp4"));

    // Smaller than the room: the room, which is what a fit is for something drawn into the box it
    // is given rather than scaled out of a picture of its own. The box is asked of
    // `scale_in_room` rather than of the scale alone, because the scale is only half the claim —
    // `FitToScreen` of a film is the whole work area only because nothing caps it there.
    assert_eq!(
        scale_dimensions(
            1920,
            1080,
            room.0,
            room.1,
            hover_preview_scale_of(&hover, at_fit)
        ),
        (3840, 2160),
        "a 1080p film fitted to a 4K display is stretched over the whole work area, which is the \
             size it is watched at"
    );

    // And the pinned window's own fit is the same answer, which is what makes the hover and the
    // maximize one rule rather than two: a film fitted in a pinned window fills the room.
    assert_eq!(
        hover_preview_scale_of(&hover, at_fit),
        effective_preview_scale(&small, at_fit),
        "a hover's fit and a pin's fit are the same fit, so a film does not change size when it is \
             pinned"
    );

    // Larger than the room: still fitted down, because shrinking to fit is what a fit is for.
    assert_eq!(
        scale_dimensions(
            7680,
            4320,
            room.0,
            room.1,
            hover_preview_scale_of(&hover_large, at_fit)
        ),
        (3840, 2160),
        "a 8K film fitted to a 4K display is still fitted down to it"
    );

    assert_eq!(
        hover_preview_scale_of(
            &hover,
            HoverScales {
                video: PreviewScale::Percent(150),
                ..hover_scales()
            }
        ),
        PreviewScale::Percent(150),
        "and a percentage above 100 enlarges, because that is the user asking for it rather \
             than a fit guessing"
    );

    assert_eq!(
        hover_preview_scale_of(&hover, hover_scales()),
        PreviewScale::Percent(100),
        "which is also what the app starts at, so a fresh install shows a film at its own size \
             and still does"
    );

    // A reduced fit of the display is the room's fitted size reduced to the
    // share it names, and a bitmap's fitted size is its own — so a hover is
    // asked for that share of the film's own size, rather than half of the room.
    let at_half_fit = HoverScales {
        video: PreviewScale::FitToScreenReduced(50),
        ..hover_scales()
    };
    assert_eq!(
        hover_preview_scale_of(&hover, at_half_fit),
        PreviewScale::Percent(50),
        "a reduced fit asks a video for half of its own size, not half of the room"
    );
    assert_eq!(
        scale_dimensions(
            1920,
            1080,
            room.0,
            room.1,
            hover_preview_scale_of(&hover, at_half_fit)
        ),
        (960, 540),
        "a 1080p film at `By Screen 50%` is 960x540: half of the film, not half of the room, \
             which is the base the menu's percentage rows are read against"
    );

    // And a pinned window reads the fit the same way the hover does — asserted beside it above —
    // while the two still part company at a reduced fit, which is the answer `pinned_media_box`
    // is written for and the one thing this change did not move.
    assert_eq!(
        scale_dimensions(
            1920,
            1080,
            room.0,
            room.1,
            effective_preview_scale(&small, at_fit)
        ),
        (3840, 2160),
        "a pinned window is fitted to its media at a fit, media and all, and a maximize is \
             a maximize"
    );

    // And the pin keeps the whole of the fit's bargain at a reduced fit too: the
    // room is what is reduced, so a pinned film at half of a fitted room is
    // half of the display rather than half of the film. It is the one place the two roads are
    // still apart, and deliberately: a pinned window is fitted to the media it shows, and a
    // reduced fit of the display is the room a maximize takes its share of.
    assert_eq!(
        scale_dimensions(
            1920,
            1080,
            room.0,
            room.1,
            effective_preview_scale(&small, at_half_fit)
        ),
        (1920, 1080),
        "a pinned window is fitted to the room a reduced fit names, reduced"
    );
}

/// A GIF that moves, a WebP that moves and a PNG that moves are one kind of preview —
/// the kind whose size `animated_scale` names — while a GIF or a PNG that holds a
/// single frame is a picture like any other and keeps the picture scale. What tells
/// them apart is the file's own content, not its name: an animated PNG is very often
/// called `.png`, and a `.gif` written by a still encoder is a picture.
#[test]
fn an_animation_follows_its_own_scale_and_a_still_keeps_the_pictures() {
    let folder = std::env::temp_dir().join("rhp-animated-scale-fixtures");
    std::fs::create_dir_all(&folder).expect("a fixture folder");

    let animated_gif = folder.join("animated.gif");
    write_test_gif(&animated_gif, 2);
    let still_gif = folder.join("still.gif");
    write_test_gif(&still_gif, 1);
    let animated_apng = folder.join("animated.png");
    write_test_png(&animated_apng, true);
    let still_png = folder.join("still.png");
    write_test_png(&still_png, false);

    let scales = HoverScales {
        picture: PreviewScale::Percent(100),
        animated: PreviewScale::Percent(25),
        ..hover_scales()
    };

    assert_eq!(
        effective_preview_scale(&animated_gif, scales),
        PreviewScale::Percent(25),
        "a GIF with two frames is an animation"
    );
    assert_eq!(
        effective_preview_scale(&still_gif, scales),
        PreviewScale::Percent(100),
        "a GIF with one frame is a picture"
    );
    assert_eq!(
        effective_preview_scale(&animated_apng, scales),
        PreviewScale::Percent(25),
        "a PNG whose chunks say it animates is an animation"
    );
    assert_eq!(
        effective_preview_scale(&still_png, scales),
        PreviewScale::Percent(100),
        "a PNG without the animation control chunk is a picture"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// A WebP's animation is a chunk of its container, and the probe reads it there: an
/// `ANIM` chunk ahead of the picture data is an animation, while a plain `VP8 ` or
/// `VP8L` WebP is a picture whatever else follows it. Nothing checksummed is written
/// here — the probe walks chunk lengths, and the frames those chunks hold are the
/// loader's business rather than this question's.
#[test]
fn a_webp_is_an_animation_when_its_container_says_so() {
    let folder = std::env::temp_dir().join("rhp-webp-probe-fixtures");
    std::fs::create_dir_all(&folder).expect("a fixture folder");

    let chunk = |kind: &[u8; 4], body: &[u8]| {
        let mut bytes = kind.to_vec();
        bytes.extend_from_slice(&(body.len() as u32).to_le_bytes());
        bytes.extend_from_slice(body);
        if body.len() % 2 == 1 {
            bytes.push(0); // chunks are padded to an even length
        }
        bytes
    };
    let riff = |chunks: Vec<u8>| {
        let mut bytes = b"RIFF".to_vec();
        bytes.extend_from_slice(&((chunks.len() + 4) as u32).to_le_bytes());
        bytes.extend_from_slice(b"WEBP");
        bytes.extend_from_slice(&chunks);
        bytes
    };

    let mut animated_chunks = chunk(b"VP8X", &[0x02, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
    animated_chunks.extend_from_slice(&chunk(b"ANIM", &[0; 6]));
    animated_chunks.extend_from_slice(&chunk(b"ANMF", &[0; 16]));
    let animated_path = folder.join("animated.webp");
    std::fs::write(&animated_path, riff(animated_chunks)).expect("a written WebP");

    let still_path = folder.join("still.webp");
    std::fs::write(&still_path, riff(chunk(b"VP8 ", &[0; 4]))).expect("a written WebP");

    assert!(image_is_animated(&animated_path), "an `ANIM` chunk");
    assert!(!image_is_animated(&still_path), "a plain picture");
    assert!(
        !image_is_animated(&folder.join("missing.webp")),
        "a file that is not there is not an animation"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// Every animated kind this app has is a kind whose queue is drained while it plays.
///
/// This is here because getting it wrong does not look like getting it wrong. A new
/// animated `MediaType` left out of `is_streaming` still decodes and still advances:
/// what it does not do is drain its queue, so the decoder waits forever in
/// `await_frame_queue_room` and the animation plays its first two frames and stops —
/// a hang, in a thread, that no test above would catch and no compiler enforces,
/// because `MediaType` is not `#[non_exhaustive]`.
#[test]
fn every_animated_kind_drains_its_queue() {
    for media_type in [
        MediaType::AnimatedGif,
        MediaType::AnimatedApng,
        MediaType::AnimatedWebP,
        MediaType::AnimatedHeif,
        MediaType::AnimatedJxl,
    ] {
        // A player whose decode is finished is not streaming whatever its kind is, so
        // the kind is asked of one that has not finished.
        let media = MediaData {
            frames: vec![],
            shared_frames: None,
            all_frames_loaded: Some(Arc::new(AtomicBool::new(false))),
            current_frame: 0,
            last_frame_time: Instant::now(),
            media_type,
            stream_cancel: None,
            video_process: None,
            loading_start: None,
            text_state: None,
            audio_scale: None,
        };

        assert!(
            media.is_streaming(),
            "{media_type:?} is decoded on its own thread and its queue must be drained"
        );
        assert_eq!(
            media.media_type.kind(),
            Some(PreviewType::Images),
            "{media_type:?} is a moving picture, and a moving picture is a picture"
        );
        assert!(
            media.media_type.has_bubble_picture(),
            "{media_type:?} has a frame, and a frame is what a bubble draws"
        );
    }
}

#[test]
fn only_a_gif_with_a_second_frame_counts_as_one() {
    let folder = std::env::temp_dir().join("rhp-animated-probe-fixtures");
    std::fs::create_dir_all(&folder).expect("a fixture folder");

    let gif = folder.join("animation.gif");
    write_test_gif(&gif, 2);
    let still = folder.join("still.gif");
    write_test_gif(&still, 1);

    assert!(image_is_animated(&gif), "two frame blocks in the file");
    assert!(!image_is_animated(&still), "one frame block in the file");

    // The same file the loader sees, so the size a hover is placed at is the size its
    // frames are decoded at: two frames are what makes it an animation there as well.
    let loaded = load_animated_gif(
        &gif,
        64,
        64,
        PreviewScale::Percent(100),
        Arc::new(AtomicBool::new(false)),
    )
    .expect("the two-frame file loads as an animation");
    assert_eq!(
        loaded.media_type.kind(),
        Some(PreviewType::Images),
        "it is a preview of the `Images` kind"
    );

    assert!(
        load_animated_gif(
            &still,
            64,
            64,
            PreviewScale::Percent(100),
            Arc::new(AtomicBool::new(false)),
        )
        .is_none(),
        "a single frame is not an animation, so the loader turns it down"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

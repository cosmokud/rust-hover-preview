use super::*;

/// A band the frame is not the size of is filled by the frame, sampled: it is what a pinned
/// window draws while an edge of it is being dragged, and what it may never draw is the frame
/// standing at its own size with nothing around it.
#[test]
fn a_frame_of_another_size_is_sampled_into_the_band_it_is_drawn_in() {
    fn band(
        frame: &[u8],
        source: (u32, u32),
        size: (u32, u32),
        background: TransparentBackground,
        opaque: bool,
    ) -> Vec<u8> {
        let mut out = vec![0u8; size.0 as usize * size.1 as usize * 4];
        resample_into_band(
            frame,
            source,
            size,
            background,
            opaque,
            BandTarget {
                out: &mut out,
                // The sampler is the road that has no surface of its own — what the frame is
                // scaled *out of* is the fast road's business, and this is the test of the
                // road that scales the frame itself (see `stretch_into_band`).
                dc: HDC(std::ptr::null_mut()),
                width: size.0,
                origin_y: 0,
                height: size.1,
            },
        );
        out
    }

    // A frame of one color, into a band twice its size: every pixel of the band is the media.
    let red = [0u8, 0, 255, 255].repeat(4);
    let scaled = band(
        &red,
        (2, 2),
        (4, 4),
        TransparentBackground::Transparent,
        true,
    );
    assert_eq!(scaled.len(), 4 * 4 * 4);
    assert!(
        scaled
            .as_chunks::<4>()
            .0
            .iter()
            .all(|pixel| pixel == &[0, 0, 255, 255]),
        "the whole band is the media"
    );

    // Four pixels of four colors into a band of one: the sample is the middle of the square
    // they make, which is a quarter of each of them.
    let quarters = [
        0u8, 0, 0, 255, // black
        255, 0, 0, 255, // blue
        0, 255, 0, 255, // green
        0, 0, 255, 255, // red
    ];
    assert_eq!(
        band(
            &quarters,
            (2, 2),
            (1, 1),
            TransparentBackground::Transparent,
            true
        ),
        [64, 64, 64, 255]
    );

    // And a frame that is not opaque is composited where it lands: half a white pixel over
    // black is the grey between them, and the band it lands in is opaque either way.
    let half = [255u8, 255, 255, 128].repeat(4);
    assert_eq!(
        band(&half, (2, 2), (2, 2), TransparentBackground::Black, false),
        [128, 128, 128, 255].repeat(4)
    );

    // What the backdrop is placed by is where a pixel is in the *band*: the squares of a
    // checkerboard do not move with the picture being dragged over them.
    let squares = band(
        &half,
        (2, 2),
        (32, 32),
        TransparentBackground::Checkerboard,
        false,
    );
    let corner = &squares[..4];
    let sixteen_across = &squares[16 * 4..16 * 4 + 4];
    assert_ne!(corner, sixteen_across, "the next square is another shade");
    assert_eq!(
        corner,
        &squares[..4],
        "and the first square is where it was"
    );
}

/// And the other road a scaled band is filled by, which is the one an opaque frame takes: GDI
/// stretches the frame into the window's own surface, which is what keeps a window a hand is
/// dragging by its edge at the pace of the hand rather than at the pace of a sample per pixel
/// of a display's worth of them (see `stretch_into_band`).
///
/// What is asked of the result is what a layered window needs of it: the frame at the band's
/// size, in the band's rows and nowhere else, with every pixel of it opaque. The last is the
/// part worth a test of its own — a 32-bit `BI_RGB` surface has no alpha channel as far as GDI
/// is concerned, so the byte it leaves there is one this side has to write rather than trust.
#[test]
fn a_stretched_band_is_the_frame_at_the_size_it_is_drawn_at() {
    let (width, height, band_height, origin_y) = (8u32, 6u32, 4u32, 2u32);
    let surface = DibSurface::create(width, height).expect("a surface to stretch into");
    let out =
        unsafe { std::slice::from_raw_parts_mut(surface.bits(), (width * height * 4) as usize) };

    let media = [32u8, 64, 128, 255].repeat(4);
    assert!(
        stretch_into_band(
            &media,
            (2, 2),
            (width, band_height),
            &mut BandTarget {
                out,
                dc: surface.dc,
                width,
                origin_y,
                height: band_height,
            },
        ),
        "a frame of the surface's own format is one GDI stretches"
    );

    for (row, line) in out.chunks(width as usize * 4).enumerate() {
        for (column, pixel) in line.as_chunks::<4>().0.iter().enumerate() {
            if row < origin_y as usize {
                assert_eq!(
                    pixel,
                    &[0, 0, 0, 0],
                    "row {row} is above the band and is nothing at all"
                );
                continue;
            }

            assert_eq!(pixel[3], 255, "pixel {row}:{column} is opaque");
            for (channel, wanted) in [32i32, 64, 128].iter().enumerate() {
                let got = pixel[channel] as i32;
                assert!(
                    (got - wanted).abs() <= 4,
                    "pixel {row}:{column} is the media: {pixel:?}"
                );
            }
        }
    }
}

/// A page that is laid out to its box — a document, a listing — is resized one dimension at
/// a time: a wider window is a longer line, and there is no shape of the file's to keep
/// because the box the text is poured into *is* the layout.
#[test]
fn a_page_that_is_laid_out_to_its_box_is_resized_one_edge_at_a_time() {
    let window = (0, 0, 400, 300 + pinned_caption_height(96, None));
    let resized = resize_pinned_window(
        window,
        PinResize {
            left: false,
            top: false,
            right: true,
            bottom: false,
        },
        200,
        0,
        96,
        false,
        false,
        PinFrame::Free,
        pinned_caption_height(96, None),
    );

    let content = content_box_of(resized, 96, false, false, pinned_caption_height(96, None));
    assert_eq!(content, (0, 30, 600, 330));

    // And its own floor holds: a page cannot be dragged to nothing.
    let resized = resize_pinned_window(
        window,
        PinResize {
            left: false,
            top: false,
            right: false,
            bottom: true,
        },
        0,
        -10_000,
        96,
        false,
        false,
        PinFrame::Free,
        pinned_caption_height(96, None),
    );
    let content = content_box_of(resized, 96, false, false, pinned_caption_height(96, None));
    assert_eq!(content, (0, 30, 400, 30 + PIN_MIN_MEDIA_PIXELS as i32));
}

#[test]
fn what_a_kind_is_framed_by_follows_what_is_inside_it() {
    // What is drawn is the file's own shape, so the box can only be a box of that shape:
    // a picture, a page an engine rendered, and a video — whichever of the two players is
    // the one drawing it. A video FFmpeg plays is framed like any other picture because its
    // pixels are still the file's own shape, and a window of somebody else's can be asked
    // for a different size and even begun again at one (see `relayout_pinned_media`).
    assert_eq!(pin_frame(Some(MediaType::StaticImage)), PinFrame::Shaped);
    assert_eq!(pin_frame(Some(MediaType::AnimatedGif)), PinFrame::Shaped);
    assert_eq!(pin_frame(Some(MediaType::NativeVideo)), PinFrame::Shaped);
    assert_eq!(pin_frame(Some(MediaType::Video)), PinFrame::Shaped);
    assert_eq!(pin_frame(Some(MediaType::Pdf)), PinFrame::Shaped);

    // A page of text is poured into whatever box it is given, at any shape.
    assert_eq!(pin_frame(Some(MediaType::Text)), PinFrame::Free);
    assert_eq!(pin_frame(Some(MediaType::Archive)), PinFrame::Free);

    // And one kind is framed by nothing: a sound's card is its own size, and a window around
    // it would only be a window with room in it.
    assert_eq!(pin_frame(Some(MediaType::Audio)), PinFrame::None);
}

#[test]
fn a_video_played_by_ffmpeg_carries_a_transport_bar_that_does_something() {
    // Both kinds of video now carry a live bar, for opposite reasons — the engine answers
    // every question it is asked, and FFmpeg's player answers the two that are keys posted to
    // its window — so the button, the track and the bar underneath are all the pointer's.
    assert_eq!(
        pin_transport_kind(Some(MediaType::Video)),
        pin_transport_live(Some(MediaType::Video)),
        "a bar that cannot be told anything must not draw a button it cannot honour"
    );
    assert!(pin_transport_live(Some(MediaType::NativeVideo)));
    assert!(
        !pin_transport_live(Some(MediaType::Audio)),
        "a sound's controls are the card's, not a bar's"
    );
    assert!(!pin_transport_live(None));
}

#[test]
fn a_maximized_box_is_the_display_by_the_medias_own_shape() {
    let room = ScreenBounds {
        left: 0,
        top: 0,
        right: 1000,
        bottom: 1000,
    };

    // Maximizing is Fit to Screen: the media is given the largest box of its own shape the
    // room has, which is the room with a margin at two of its sides rather than the room
    // itself — the letterbox a box of the display's shape would leave is the one thing a
    // pinned window is never given.
    assert_eq!(
        pinned_media_box((4000, 3000), room, PreviewScale::FitToScreen),
        (1000, 750)
    );
    assert_eq!(
        pinned_media_box((3000, 4000), room, PreviewScale::FitToScreen),
        (750, 1000)
    );

    // And one smaller than the display is scaled up to it rather than left at its own size
    // in the middle of the screen: the button says "as large as this display can show it",
    // and a box that refused to grow would be a button that did nothing on a small picture.
    assert_eq!(
        pinned_media_box((400, 300), room, PreviewScale::FitToScreen),
        (1000, 750)
    );

    // What a display change asks for is not the same question: a pin that still fits keeps
    // the size the user chose for it, and only one that no longer fits is brought down to
    // the room.
    assert_eq!(
        pinned_media_box((400, 300), room, PreviewScale::Percent(100)),
        (400, 300)
    );
    assert_eq!(
        pinned_media_box((4000, 3000), room, PreviewScale::Percent(100)),
        (1000, 750)
    );
}

#[test]
fn a_box_is_centred_in_a_room_and_around_a_point() {
    let room = ScreenBounds {
        left: 100,
        top: 50,
        right: 1100,
        bottom: 850,
    };

    // A box smaller than the room is put in the middle of it, and what the two sides are left
    // with differs by at most the odd pixel of an odd difference.
    assert_eq!(centred_box((400, 300), room), (400, 300, 800, 600));
    assert_eq!(centred_box((401, 301), room), (399, 299, 800, 600));

    // And one the room cannot hold goes against the room's own top-left corner rather than half
    // past it: there is no middle to be in when the box is bigger than the place it is put in.
    assert_eq!(centred_box((1200, 900), room), (100, 50, 1300, 950));

    // A size centred on a point — where a swap puts another file's shape — is put around that
    // point, wherever it is: a middle is not a place, so nothing here holds the box on a
    // display.
    assert_eq!(centred_at((400, 300), (600, 400)), (400, 250, 800, 550));
    assert_eq!(centred_at((401, 301), (600, 400)), (400, 250, 801, 551));
    assert_eq!(centred_at((400, 300), (-100, 40)), (-300, -110, 100, 190));
}

/// What a restore down owes after the hand has had a maximized window: the box the window had
/// before it was maximized while nothing has moved it, and nothing at all the moment the hand
/// has moved it — carried to another place or pulled to a size, a maximize the user has taken
/// hold of is one there is nothing left to undo, so the caption draws a maximize again (see
/// `pin_restore_box`).
#[test]
fn a_restore_puts_back_the_size_the_window_came_from() {
    // The box the window had before the maximize button was pressed, and the box maximize put
    // it in.
    let before = (400, 300, 800, 600);
    let maximized = (100, 100, 1100, 850);

    // Untouched since the maximize — a press the hand did not move is no drag — and the box
    // the maximize put aside is the box a restore puts back.
    assert_eq!(
        pin_restore_box(Some(before), maximized, maximized),
        Some(before)
    );

    // Carried somewhere: the maximize is given up rather than brought along, which is what puts
    // a maximize back in the caption and makes the button maximize again.
    assert_eq!(
        pin_restore_box(Some(before), maximized, (150, 400, 1150, 1150)),
        None
    );

    // Resized: the same answer, for the same reason — the size the maximize gave the window has
    // been replaced by one the hand asked for. And a carry that follows either one has nothing
    // left to carry, there being no restore any more.
    let resized = (100, 100, 700, 550);
    assert_eq!(pin_restore_box(Some(before), maximized, resized), None);
    assert_eq!(pin_restore_box(None, resized, (300, 700, 900, 1150)), None);
}

#[test]
fn only_the_kinds_that_play_carry_a_transport_bar() {
    // A video of either engine is the kind with a playhead to show. A sound is not: its card
    // carries a bar of its own, drawn by the page that paints it.
    assert!(pin_transport_kind(Some(MediaType::Video)));
    assert!(pin_transport_kind(Some(MediaType::NativeVideo)));
    assert!(!pin_transport_kind(Some(MediaType::Audio)));
    assert!(!pin_transport_kind(Some(MediaType::StaticImage)));
    assert!(!pin_transport_kind(None));
}

/// asking whether the player behind a pinned preview is still alive walks /// the media to reach it, so the media must not be held while the walk is made.
#[test]
fn asking_after_a_pinned_players_liveness_does_not_stall_on_the_media_it_asks_about() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());

    let mut video = create_loading_media(8, 8);
    video.media_type = MediaType::Video;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(video);
    }

    let (sender, receiver) = std::sync::mpsc::channel();
    std::thread::spawn(move || {
        let _ = sender.send(pin_media_is_alive(false));
    });

    let answer = receiver.recv_timeout(Duration::from_secs(5)).expect(
        "the ask came back: a player is reached *through* the media, so the media must not \
             still be held when it is asked about",
    );
    assert!(!answer, "a video with no player behind it is not alive");

    // And the same question asked of the same player while a next file is already in hand
    // has the other answer: nothing is running because a swap was asked for, and a pin that
    // is on its way to being shown something else is not a window onto nothing.
    assert!(
        pin_media_is_alive(true),
        "a pin already being shown another file is not a pin that came apart"
    );

    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous;
    }
}

/// a pinned film the media engine took and then failed at, with the walk that steps /// over it already queued.
///
/// This is the case the liveness close was wrong about: the swap that reached the file
/// emptied the media slot before it knew the engine would take it, so the player being gone
/// is what this app asked for rather than a window onto nothing — and reading it the other
/// way took the window down a tick before the queued walk could be taken up.
#[test]
fn a_pinned_film_the_engine_failed_at_is_stepped_over_rather_than_closed() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let mut film = create_loading_media(8, 8);
    film.media_type = MediaType::NativeVideo;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(film);
    }

    let mut pin = overlay_pin((0, 0, 80, 60), PinChrome::always());
    pin.path = PathBuf::from("C:\\folder\\broken.mp4");
    stand_pin(Some(pin));

    // The loop's own flag, from the four slots that mean a next file is already pending. The
    // walk is the one that matters here: the tick that finds the failure out queues it here,
    // and `pin_command_request` asks whether the pin is still there on the very next tick.
    let walk: Option<PinStep> = Some(PinStep {
        at: PathBuf::from("C:\\folder\\next.mp4"),
        from: PathBuf::from("C:\\folder\\broken.mp4"),
        step: 1,
        left: 1,
        list: vec![
            PathBuf::from("C:\\folder\\broken.mp4"),
            PathBuf::from("C:\\folder\\next.mp4"),
        ],
    });
    let load: Option<PinLoad> = None;
    let awaiting_box: Option<PathBuf> = None;
    let held_pick: Option<PathBuf> = None;

    assert!(
        !pin_media_is_alive(false),
        "a film the engine is failing at, with nothing queued to replace it, is still a pin \
             that has come apart: this is the close that stays"
    );
    assert!(
        pin_media_is_alive(
            walk.is_some() || load.is_some() || awaiting_box.is_some() || held_pick.is_some()
        ),
        "and the same film with the walk queued is not a pin that came apart: the player is \
             gone because the pin is on its way to being shown something else, so the walk gets \
             to be taken up rather than the window coming down under it"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// a pin shown a document the browser has to *start* for.
#[test]
fn an_engine_coming_up_for_a_pinned_document_is_not_a_pin_that_came_apart() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let folder = std::env::temp_dir().join("rust-hover-preview-pin-engine");
    std::fs::create_dir_all(&folder).expect("a test folder");
    let path = folder.join("drawing.svg");
    std::fs::write(
        &path,
        br#"<svg xmlns="http://www.w3.org/2000/svg" width="40" height="24"/>"#,
    )
    .expect("a written file");

    let mut media = create_loading_media(8, 8);
    media.media_type = MediaType::EngineSvg;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }

    let mut pin = overlay_pin((0, 0, 80, 60), PinChrome::always());
    pin.path = path.clone();
    stand_pin(Some(pin));

    // The browser is coming up for this document: the engine owes it, and no window exists yet.
    webview_preview::publish_want_for_test(&path);
    assert!(
        pin_media_is_alive(false),
        "a browser that is coming up is the thing the pin is a window onto"
    );

    // And a document nothing is owed for and nothing is showing is what it always was: the pin
    // comes down.
    webview_preview::clear_want_for_test();
    assert!(
        !pin_media_is_alive(false),
        "a document the engine no longer owes and no window shows is a window onto nothing"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    let _ = std::fs::remove_dir_all(&folder);
}

/// a pin collapsed into its bubble whose player this app has parked, which /// is a player that is deliberately not running.
#[test]
fn a_player_the_bubble_parked_is_not_a_pin_that_came_apart() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let mut video = create_loading_media(8, 8);
    video.media_type = MediaType::Video;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(video);
    }

    let mut pin = overlay_pin((0, 0, 80, 60), PinChrome::always());
    pin.collapsed = true;
    pin.bubble_pause = Some(BubblePause::Player(12.5));
    stand_pin(Some(pin));

    assert!(
        pin_media_is_alive(false),
        "the player is parked rather than gone: the pin is a bubble standing for it"
    );

    update_pin_bubble_pause(None);
    assert!(
        !pin_media_is_alive(false),
        "and a video whose player is simply gone is what it always was"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// a pinned window playing a sound, whose player plays the pass it was /// given and stops at the end of it.
#[test]
fn a_pinned_sound_is_not_a_pin_that_came_apart_between_passes() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());

    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let folder = std::env::temp_dir().join("rust-hover-preview-pin-sound");
    let path = folder.join("pass.mp3");
    std::fs::create_dir_all(&folder).expect("a test folder");
    std::fs::write(&path, b"not really a sound").expect("a written file");

    let mut sound = create_loading_media(8, 8);
    sound.media_type = MediaType::Audio;
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(sound);
    }

    let mut pin = overlay_pin((0, 0, 80, 60), PinChrome::always());
    pin.path = path.clone();
    stand_pin(Some(pin));

    assert!(
        pin_media_is_alive(false),
        "a card is this app's own text: a player between two passes is not a window onto \
             nothing, and taking the window down for it closed a looping sound at the end of \
             every pass"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
    let _ = std::fs::remove_dir_all(&folder);
}

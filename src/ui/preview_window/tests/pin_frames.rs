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

/// A drag of a pinned film hides the player's window for its whole length, and the band that
/// window was standing in is this app's to fill for that length. What it is filled with is the
/// last frame the player had on screen, scaled to the box being dragged to — a hand choosing a
/// size needs the film in the box, not a rectangle the film is behind.
///
/// So this is a test of the three answers the fill can give, and the middle one is the one that
/// used to be the only one: a frame grows with the box and shrinks into it, both of them opaque,
/// and a frame that could not be taken leaves black rather than a hole. Black rather than
/// transparent is not a smaller version of the same wrong thing — a layered window's hit testing
/// and its compositing are both answered by the shape of its pixels, so a band of transparent
/// ones is the desktop, and a drag that leaves one is a hole in the desktop shaped like a video
/// (see `compose_parked_band` and `render_pinned_preview_at`).
#[test]
fn a_band_whose_player_is_parked_is_the_last_frame_scaled_and_never_a_hole() {
    // A window's own surface with a band of it four rows down, so what is asserted is the band and
    // the rows around it: a fill that reached past the band would be as wrong as one that stopped
    // short of it. The band is always the window's own width — that is what a band is, and what
    // `stretch_into_band` stretches into — so the dimension a drag changes is the height.
    fn painted(
        held: Option<(&[u8], (u32, u32))>,
        window: (u32, u32),
        band_height: u32,
    ) -> (Vec<u8>, bool) {
        let (width, height) = window;
        let origin_y = 4u32;
        let surface = DibSurface::create(width, height).expect("a surface to paint the band into");
        let out = unsafe {
            std::slice::from_raw_parts_mut(surface.bits(), (width * height * 4) as usize)
        };

        let filled = compose_parked_band(
            held,
            width,
            BandTarget {
                out,
                dc: surface.dc,
                width,
                origin_y,
                height: band_height,
            },
        );
        (out.to_vec(), filled)
    }

    // A red frame, so a band that is anything but the film says so.
    let red = [0u8, 0, 255, 255].repeat(64);

    // Grown: a four-by-four frame into a band eight rows high is the frame at the band's size,
    // which is what makes a hand dragging an edge downwards see the film grow.
    let (grown, filled) = painted(Some((&red, (4, 4))), (16, 12), 8);
    assert!(filled, "a held frame fills the band it is scaled into");

    // Shrunk: an eight-by-four frame into a band two rows high, sampled down rather than cropped.
    let (shrunk, filled) = painted(Some((&red, (8, 4))), (8, 8), 2);
    assert!(
        filled,
        "and it fills a band it is being shrunk into just as well"
    );

    // Both are opaque everywhere and are the film, which is the whole of what the user is owed:
    // scaled, and never a hole to see the desktop through.
    for (surface, window, band_height, what) in [
        (&grown, (16u32, 12u32), 8u32, "grown"),
        (&shrunk, (8u32, 8u32), 2u32, "shrunk"),
    ] {
        let (width, height) = window;
        for row in 0..height as usize {
            for column in 0..width as usize {
                let pixel = surface_pixel(surface, window, row, column);
                let in_band = row >= 4 && row < 4 + band_height as usize;

                if !in_band {
                    assert_eq!(
                        pixel,
                        &[0, 0, 0, 0],
                        "{what}: pixel {row}:{column} is outside the band and is nothing at all"
                    );
                    continue;
                }

                assert_eq!(pixel[3], 255, "{what}: pixel {row}:{column} is opaque");
                for (channel, wanted) in [0u8, 0, 255].iter().enumerate() {
                    assert!(
                        (pixel[channel] as i32 - *wanted as i32).abs() <= 8,
                        "{what}: pixel {row}:{column} is the film scaled: {pixel:?}"
                    );
                }
            }
        }
    }

    // And with nothing held there is still no hole: black is a picture this app owns, transparent
    // is the desktop showing through a window that has nothing behind it.
    let (bare, filled) = painted(None, (16, 12), 8);
    assert!(
        !filled,
        "a park that took no frame says so, rather than answering with an empty frame"
    );
    for row in 4..12usize {
        for pixel in bare[(row * 16 * 4)..(row * 16 * 4 + 16 * 4)]
            .as_chunks::<4>()
            .0
        {
            assert_eq!(
                pixel,
                &[0, 0, 0, 255],
                "row {row} is black and opaque, never a hole in the desktop"
            );
        }
    }
}

/// One pixel of a painted surface, by row and column: the windows these tests paint are of two
/// sizes, so a pixel cannot be spelled out as an offset into a fixed stride.
fn surface_pixel(surface: &[u8], window: (u32, u32), row: usize, column: usize) -> &[u8] {
    let stride = window.0 as usize * 4;
    &surface[row * stride + column * 4..row * stride + column * 4 + 4]
}

/// **The first of the two flashes a drag used to have, and it is an ordering rather than a
/// mechanism.** A layered window is drawn from the surface it was last painted into, so a park that
/// hid the player's window and left the paint to the next repaint spent the gap between the two
/// with the band's last painted pixels still on screen — and for a film playing those are
/// transparent ones, because the band of a video is left empty on purpose so FFmpeg's own window
/// shows through it. Transparent with nothing behind it is the desktop: the window the hand was
/// dragging out of is what the user sees, for as long as the compositor takes to reach a repaint,
/// and a *move* answers none at all until the transport bar's own two-hundred-millisecond one.
///
/// So the park is a capture, a paint and a hide, in that order, and this asserts both halves of it:
/// the order the three happened in, and what the window's surface holds the moment the park
/// returns — which is the pixels the compositor is about to be handed, and is the only place the
/// user's first frame of a drag actually exists on a machine with no player of this app's.
#[test]
fn a_park_paints_the_frame_it_took_before_it_hides_the_players_window() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = PathBuf::from("parked-before-it-is-hidden.mkv");
    pin.dpi = 96;
    stand_pin(Some(pin));
    forget_pin_park_swap();
    clear_park_trace();

    // The frame a real park would have taken off the desktop is stood in for it: there is no
    // player of this app's on the machine these run on, and the capture answers false rather than
    // writing a blank frame over a good one (see `hold_video_window_frame`).
    let red = [0u8, 0, 255, 255].repeat(64);
    stand_video_frame_for_a_test(red.clone(), 8, 8);

    let hwnd = HWND(0x1000 as *mut _);
    assert!(
        park_pinned_player(hwnd, (100, 80), false),
        "the first pointer message of a drag parks the player"
    );

    assert_eq!(
        park_trace(),
        vec!["capture", "paint", "hide"],
        "and it takes the frame, paints it and only then puts the window away — the reverse order \
         spends the gap between the hide and the paint with the band's transparent pixels still \
         composited, which is the desktop showing through a window being dragged"
    );

    // And the surface the compositor is handed is opaque where the film is. The window here is the
    // paint's own key, so the very surface the park painted is the one read back.
    let (width, height) = pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_size()))
        .expect("a pin to have a size");
    let bits = ensure_layered_surface(hwnd.0 as isize, width as u32, height as u32)
        .expect("the surface the park painted");
    let painted = unsafe { std::slice::from_raw_parts(bits, width as usize * height as usize * 4) };

    let (band_top, band_height) = pinned_band_rows(
        height,
        pinned_caption_height(96, None),
        0,
        pin_state().and_then(|pinned| pinned.pin().map(|pin| pin.overlay)) == Some(true),
    );
    let mut band_pixels = 0;
    for row in 0..height as usize {
        let in_band = row >= band_top.max(0) as usize
            && row < band_top.max(0) as usize + band_height as usize;
        let pixel = surface_pixel(painted, (width as u32, height as u32), row, 0);
        if in_band {
            assert_eq!(
                pixel[3], 255,
                "row {row} is in the band and is opaque, not a hole in the desktop"
            );
            band_pixels += 1;
        }
    }
    assert!(
        band_pixels > 0,
        "and the band is a band: a paint that put nothing there has not filled anything"
    );

    forget_video_frame();
    forget_resume_frame();
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// **The second flash is the drag's end, and it is the same mistake pointed the other way: handing
/// the band back before there is anything to see through it.** A move has no relaunch behind it, so
/// the release used to show the player's window itself — the flag down, the band transparent, and
/// the compositor's first look at a window that had been hidden for the length of a drag still
/// filling in. A resize's release already left the settle to do it, and the settle asked only
/// whether the window was visible, which a replacement's window is within milliseconds of being
/// begun and empty until the file is open.
///
/// So the release answers neither: it drops the drag and leaves the band this app's to fill, and the
/// settle takes the park back on the arm it decides (see `park_swap_arm`). What this asserts is the
/// half that is observable here — a release that used to put the picture back on its own message
/// and now leaves the frame standing for the settle to be answered against.
#[test]
fn a_drag_that_ends_leaves_the_frame_standing_for_the_settle_to_answer() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    stand_pin(Some(PinnedPreview {
        path: PathBuf::from("ended-by-a-tick-rather-than-by-a-message.mkv"),
        ..PinnedPreview::for_test()
    }));

    let mut media = create_loading_media(320, 240);
    media.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }

    let window = RecordedPinWindow::new(0x1000);
    assert!(park_pinned_player(HWND(0x1000 as *mut _), (0, 0), false));
    stand_video_frame_for_a_test([0u8, 0, 255, 255].repeat(64), 8, 8);

    with_pin(|pin| {
        pin.dragging = Some(PinDrag {
            from: (0, 0),
            window: (0, 0, 320, 240),
            action: PinDragAction::Move,
            delivered: true,
            carried: (i32::MIN, i32::MIN),
        });
    });
    assert!(
        finish_pin_drag(HWND(0x1000 as *mut _), &window),
        "a move's release is a drag to let go of"
    );

    assert!(
        pin_player_is_parked(),
        "and it does not put the player's window back on its own message: the flag going down is \
         what makes the band transparent, and the only thing behind a transparent band is whatever \
         has composited so far — which for a window hidden for the length of a drag is nothing"
    );
    assert!(
        !window
            .calls()
            .iter()
            .any(|call| matches!(call, PinWindowCall::UnparkPlayerWindow(_))),
        "and nothing is asked of the player's window either"
    );
    assert!(
        held_video_frame().is_some(),
        "the frame the drag was holding is still the band's picture, because the park is still \
         standing"
    );

    // And the settle is the one place that takes it back — the flag and the window in one tick,
    // through the one writer (see `unpark_pinned_player`).
    assert!(
        unpark_pinned_player(&window),
        "the swap is the settle's, and it goes through the same line a test can reach"
    );
    assert!(
        !pin_player_is_parked() && held_video_frame().is_none(),
        "which puts the flag down and gives the frame up in the same tick: a band that has a \
         window of somebody else's in it needs nothing kept behind it"
    );

    forget_video_frame();
    forget_resume_frame();
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// The arms a park's end can be taken on, and the answer that is not one of them.
///
/// **Visible is not presented, and that is the whole of the end-of-drag flash.** A replacement's
/// window is on screen and correctly sized within a few milliseconds of being begun — SDL makes it
/// before it opens the file — and it is empty until the seek is done and the first frame is
/// decoded. So the arm that swaps at once is not "the window is up" but "the player standing in
/// the band is the one that was playing before the pointer went down", which a drag holds at its
/// first pointer message and which therefore has the very frame the band is holding to show again
/// the instant it is shown. Everything else waits, and the wait is bounded — a dark scene, a
/// player that died mid-relaunch, or a machine with no decoder must not be able to hold a black
/// band against a window for ever.
///
/// **And no window at all is not a swap, but it cannot be a cover for ever either.** The band is
/// transparent because a player's window stands in it, so a band handed back with nothing behind it
/// is the desktop — and a replacement within its bound of publishing a window is answered out of
/// the hand by every place that could put it up, so swapping on the bound alone would take the
/// parked flag down and show nothing. What an opaque band *is*, though, is a shape this app's own
/// window hit-tests: a placeholder standing over the video area answers every click in it and
/// answers it wrong. So the `no window` case has a bound of its own and ends by giving the cover
/// up (see `PIN_PARK_COVER_TIMEOUT` and `ParkSwap::Abandoned`).
#[test]
fn a_band_is_handed_back_on_a_player_and_given_up_on_a_cover_that_waits_for_neither() {
    // The player that was playing is back, and it is the same process: nothing to wait for, at any
    // age — including immediately, which is what a move's release is.
    assert_eq!(
        park_swap_arm(true, false, Duration::ZERO),
        Some(ParkSwap::Presented),
        "a move's park is answered by the player that is holding the frame the band is showing"
    );
    assert_eq!(
        park_swap_arm(true, false, PIN_PARK_SWAP_TIMEOUT * 2),
        Some(ParkSwap::Presented),
        "and the age of the wait changes nothing about that answer"
    );

    // A replacement is up but has decoded nothing: waiting, until the bound says stop.
    assert_eq!(
        park_swap_arm(true, true, Duration::ZERO),
        None,
        "a replacement's window being visible is not a frame being on it"
    );
    assert_eq!(
        park_swap_arm(true, true, PIN_PARK_SWAP_TIMEOUT - Duration::from_millis(1)),
        None,
        "nor is one that is one millisecond short of the bound"
    );
    assert_eq!(
        park_swap_arm(true, true, PIN_PARK_SWAP_TIMEOUT),
        Some(ParkSwap::TimedOut),
        "and the bound hands the band back rather than holding a placeholder against an empty \
         window for ever — which is what stops a dark scene, a dead player or a missing decoder \
         from being a black band that never becomes a picture"
    );

    // No window at all is the case the swap bound does *not* cover, and the one it most certainly
    // must not: there is nothing to hand the band to, and the placeholder it is holding is opaque.
    assert_eq!(
        park_swap_arm(false, false, Duration::ZERO),
        None,
        "a player with no window cannot have presented"
    );
    assert_eq!(
        park_swap_arm(false, true, PIN_PARK_SWAP_TIMEOUT),
        None,
        "and the swap bound is not a reason to swap one that never published a window: the swap \
         takes the parked flag down and the window up together, and with no window up it is the \
         desktop"
    );

    // The cover's own bound is the case that must not be for ever, because an opaque band this
    // app's own window hit-tests answers every click in the video area for the life of the pin.
    assert_eq!(
        park_swap_arm(false, true, PIN_PARK_COVER_TIMEOUT),
        Some(ParkSwap::Abandoned),
        "past the cover's own bound the cover is given up rather than held: the band goes back to \
         being the hole it is between films, which a player's window fills the moment it has one"
    );
    assert_eq!(
        park_swap_arm(false, true, PIN_PARK_SWAP_TIMEOUT * 4),
        Some(ParkSwap::Abandoned),
        "nor is four times the swap bound any different — the wait is extended rather than \
         restarted, so a window arriving late is still handed the band on the first tick that \
         finds it, and a window that never arrives does not keep the placeholder"
    );
}

/// **The settle itself, on the arm the table above refuses: a park that has waited out its bound
/// with nothing behind the band.** The arm is a table, and a table is the place a defect like this
/// hides — the swap reads as being about *when* and is in fact about *whether there is anything to
/// show*: `unpark_pinned_player` takes the parked flag down through the one writer that takes it,
/// and every place that could put a replacement's window up is answered out of the hand while the
/// park stands, so a swap with no window in the band leaves the band transparent over the desktop,
/// for as long as the user looks at the window they were dragging.
///
/// So the bound with no window holds, and this drives the whole settle rather than the arm alone:
/// a park held past its bound keeps the placeholder and asks nothing of the player's window; a
/// frame a background rendered for it is installed and painted, because a park that is standing has
/// no drag in flight to repaint it; and the tick that finds a window up hands it the band at once,
/// the wait having been extended rather than restarted (see `settle_pinned_park_where`).
#[test]
fn a_park_waited_out_with_no_window_standing_holds_the_placeholder_until_one_is() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    clear_park_swap_arm();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = PathBuf::from("waited-out-with-nothing-behind-it.mkv");
    pin.dpi = 96;
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    // A resize's park, because that is the one with a replacement to wait for. Nothing is prepared
    // for it: the media is taken out above, so there is no playhead to render a frame at (see
    // `prepare_resume_frame`).
    assert!(
        park_pinned_player(HWND(0x1000 as *mut _), (100, 80), true),
        "the park stands, with a replacement behind it and the band holding what the drag took"
    );
    stand_video_frame_for_a_test([0u8, 0, 255, 255].repeat(64), 8, 8);

    // The bound spent: the stamp the settle takes is put back where a drag of a real length would
    // have put it, which is the only part of a park a test has to supply by hand (see
    // `park_swap_arm_for_the_band`).
    if let Ok(mut held) = PIN_PARK_SWAP.lock() {
        if let Some(swap) = held.as_mut() {
            swap.since = Some(Instant::now() - PIN_PARK_SWAP_TIMEOUT);
        }
    }

    // And a frame a background finished inside the drag, which is what a park waiting on a window
    // is most likely to be upgraded by.
    let prepared = [255u8, 0, 0, 255].repeat(64);
    stand_resume_frame_for_a_test(prepared.clone(), 8, 8);

    let window = RecordedPinWindow::new(0x1000);
    assert!(
        !settle_pinned_park_where(&window, false),
        "a bound spent with no window behind the band is not a band to hand back"
    );
    assert!(
        pin_player_is_parked(),
        "so the flag stays down, which is the whole of what makes the band this app's picture rather \
         than a hole in the desktop"
    );
    assert_eq!(
        held_video_frame().as_ref().map(|held| held.pixels.clone()),
        Some(prepared),
        "and the frame the background rendered for this park is put in place of the one the drag \
         took: the same picture at the second the film goes on from"
    );
    assert_eq!(
        window.calls(),
        vec![PinWindowCall::Repaint],
        "and nothing else is asked of the window — the player's own is not put up, and the band is \
         painted because an upgrade nothing repaints is a frame that is held and never shown"
    );
    assert_eq!(
        park_swap_last_arm(),
        None,
        "no arm was taken, so there is nothing to read back about a park that is still standing"
    );

    // Asked again with nothing new to install, which is what the tick does while it waits: the
    // frame is already in place, so there is nothing to paint and nothing to swap.
    assert!(
        !settle_pinned_park_where(&window, false),
        "and a park still waiting is asked again with the same answer"
    );
    assert_eq!(
        window.calls(),
        vec![PinWindowCall::Repaint],
        "without repainting a band that is already showing the frame in place"
    );
    assert!(
        pin_player_is_parked() && held_video_frame().is_some(),
        "and still holding it"
    );

    // The tick that finds a window: the wait was extended by the hold rather than restarted, so a
    // replacement that publishes its window late is handed the band on the first tick that sees it.
    assert!(
        settle_pinned_park_where(&window, true),
        "a window standing in the band is the thing the settle was waiting for, and the flag and \
         the window go in one tick"
    );
    assert!(
        !pin_player_is_parked() && held_video_frame().is_none(),
        "which puts the flag down and gives the frame up together: a band with a window of \
         somebody else's in it needs nothing kept behind it"
    );
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::Repaint,
            PinWindowCall::UnparkPlayerWindow(Some((100, 80, 420, 320))),
            PinWindowCall::Repaint,
        ],
        "and the player's window is put up at where the band is, rather than at where the drag \
         began, and the band is painted through it in the same tick — the window first, because a \
         paint with nothing behind it is the desktop, and the paint at all because the pixels the \
         compositor is holding are still the placeholder this park went on holding"
    );
    assert!(
        matches!(
            park_swap_last_arm(),
            Some((ParkSwap::TimedOut, waited)) if waited >= PIN_PARK_SWAP_TIMEOUT
        ),
        "and the arm is written down with the wait that was actually spent, which is where the next \
         person tuning the bound reads it from"
    );

    forget_video_frame();
    forget_resume_frame();
    forget_pin_park_swap();
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// What the pid comparison is actually for: telling a player carrying the captured picture from one
/// that opened the file after the drag began, which a window's own facts cannot do.
///
/// It takes the same lock its siblings do, because `VIDEO_PID` is one slot for the whole process
/// and this test writes a pid into it rather than standing a pin and giving it back.
#[test]
fn a_replacement_is_told_from_the_process_and_not_from_the_window() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    let previous = VIDEO_PID.swap(4242, Ordering::SeqCst);
    assert!(
        !player_replaced_since(4242),
        "the same player is the one the park captured for, whatever its window looks like"
    );
    assert!(
        player_replaced_since(4241),
        "and a different one is a replacement that has decoded nothing yet"
    );

    VIDEO_PID.store(0, Ordering::SeqCst);
    assert!(
        !player_replaced_since(4242),
        "a player that is not there is not a replacement: `VIDEO_PID` is cleared by every path \
         that ends one, so a cleared pid says the pin has nothing playing rather than that \
         something new has taken its place"
    );
    assert!(
        !player_replaced_since(0),
        "and a park that captured nothing has nothing to compare against"
    );

    VIDEO_PID.store(previous, Ordering::SeqCst);
}

/// The frame a resize's park prepares while the drag lasts, and the three answers it can have.
///
/// It is an upgrade and never a gate, which is what the whole of it is: the band is opaque from the
/// pointer message onwards whatever this returns, so the failure cases are not a blank band but the
/// stale frame still standing where it was.
#[test]
fn a_prepared_frame_upgrades_the_placeholder_and_never_gates_it() {
    assert!(
        resume_frame_upgrades(true, true, true),
        "a frame that landed inside the drag is put in place of the one taken at its first pointer \
         message — the same picture, at the second the film goes on from"
    );

    assert!(
        !resume_frame_upgrades(false, true, true),
        "a background that has not answered leaves the stale frame standing: the band is painted \
         from what it has, and what it has is a picture"
    );
    assert!(
        !resume_frame_upgrades(true, false, true),
        "a frame that arrives after the park has ended is dropped rather than kept: a drag that is \
         over has a player in the band, and a frame decoded for it would be the second a previous \
         drag let go at"
    );
    assert!(
        !resume_frame_upgrades(true, true, false),
        "and a park that took no frame has nothing to upgrade — the flat fill is what is standing \
         in for the picture, and putting a frame under nothing leaves the band reading a frame it \
         is no longer painting from"
    );
}

/// And the install is the one that takes it, so a paint is never swapped under itself.
///
/// It is asked with the park's own flag both ways round, because that is the one fact `install` has
/// no other way of learning: a slot filled by a background that finished a drag ago is a slot whose
/// contents must not reach the next drag's band (see `resume_frame_upgrades`).
#[test]
fn a_prepared_frame_is_installed_only_while_its_park_is_standing() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    let previous_pin = take_pin_for_a_test();

    let stale = [255u8, 255, 255, 255].repeat(64);
    let prepared = [0u8, 0, 255, 255].repeat(64);

    stand_pin(Some(PinnedPreview {
        path: PathBuf::from("prepared-while-the-drag-lasts.mkv"),
        ..PinnedPreview::for_test()
    }));

    stand_video_frame_for_a_test(stale.clone(), 8, 8);
    stand_resume_frame_for_a_test(prepared, 8, 8);
    assert!(
        !pin_player_is_parked(),
        "nothing is parked yet: the frame the background prepared has a park to be inside or it is \
         a decode for a drag that is over"
    );
    assert!(
        !install_resume_frame(),
        "so it is not put over a band nobody is painting"
    );
    assert_eq!(
        held_video_frame().as_ref().map(|held| held.width),
        Some(8),
        "and the stale frame is left exactly where it was, rather than a slot emptied for nothing"
    );

    // Parked, with the stale frame standing: this is the arm the install exists for.
    with_pin(|pin| pin.parked = true);
    stand_video_frame_for_a_test(stale, 8, 8);
    stand_resume_frame_for_a_test([0u8, 0, 255, 255].repeat(64), 8, 8);
    assert!(
        install_resume_frame(),
        "and one that lands inside a park is put in place of the stale frame"
    );
    let (width, height) = held_video_frame()
        .as_ref()
        .map(|held| (held.width, held.height))
        .expect("a frame for the band to scale");
    assert_eq!(
        (width, height),
        (8, 8),
        "at the size the background rendered it rather than the size it was captured at"
    );

    // A park that took no frame has the flat fill standing in for the picture, and there is nothing
    // for a prepared frame to be an upgrade of.
    forget_video_frame();
    stand_resume_frame_for_a_test([0u8, 0, 255, 255].repeat(64), 8, 8);
    assert!(
        !install_resume_frame(),
        "a band with nothing held is not upgraded by a frame, because the flat fill is the picture \
         it is painted from"
    );

    forget_video_frame();
    forget_resume_frame();
    stand_pin(previous_pin);
}

/// **A render belongs to the park that asked for it, and a park is not the only one that can be in
/// flight.** A park's end gives up its frame and the background rendering it is up to
/// `RESUME_FRAME_WAIT` of FFmpeg away from answering — so a second park begun inside that window
/// has a render already running for the first one. With one flag between them the second park took
/// the refusal back down, the first render read a park that had ended as one that had not, and put
/// a frame taken at the second the *previous* drag let go at into the band the new park was
/// holding for a picture of its own.
///
/// So a render is opened for a named park and asks twice whether that is still the park this app
/// is waiting for, which two overlapping parks can never both be. The two asks are the same fact
/// read at the two moments it matters: once between the render's reads, so a park that ended is not
/// waited out to the bound, and once with the frame in hand, so a stale render cannot fill the slot
/// a later park installs from (see `resume_frame_asked_for`).
#[test]
fn a_render_answers_only_the_park_that_asked_for_it() {
    let first = begin_resume_frame();
    assert!(
        resume_frame_asked_for(first),
        "a render that has just been opened is the one this app is waiting for"
    );

    // A second park, opened while the first is still decoding — the overlap the flag could not
    // carry, because a flag says whether *some* park has ended and this is a park that has begun.
    let second = begin_resume_frame();
    assert!(
        !resume_frame_asked_for(first),
        "so the first render finds out that its park is not the current one, and drops what it has \
         rather than putting it where a band is waiting for it"
    );
    assert!(
        resume_frame_asked_for(second),
        "and the render the second park opened is the one it is waiting for"
    );

    // And the other end of the first park: its frame is given up rather than wanted.
    abandon_resume_frame();
    assert!(
        !resume_frame_asked_for(second),
        "a park that has been given back has a player in the band and no use for a decode, so the \
         render it opened finds out too"
    );
    assert!(
        resume_frame_asked_for(begin_resume_frame()),
        "and the next park's render is wanted again — a refusal is a park's end, not a stop"
    );
}

/// **The flash that is left after the swap, and it is not an ordering between the two ends of a
/// drag: it is a paint that never happens.** The swap takes the parked flag down and puts the
/// player's window up in one tick, and both of those are right — but a layered window is drawn
/// from the surface it was last painted into, and the last paint was the one `park_pinned_player`
/// made with the band held flat under the frozen frame. The flag going down changes what the
/// *next* paint would draw; it changes nothing about the pixels the compositor is already holding.
///
/// So the frame the drag was holding stays composited over the player's window until something
/// paints the band transparent again, and for a pinned video the only thing that does is the
/// transport bar's own quarter-of-a-second cadence: a flash of exactly the frozen image on every
/// release, for as long as that cadence is owed rather than a frame.
///
/// The paint belongs inside the swap, after the window is up — in that order, because a repaint
/// with nothing behind the band is the hole this whole arrangement exists not to leave.
#[test]
fn a_swap_repaints_the_band_it_hands_back_in_the_same_tick() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    clear_park_swap_arm();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = PathBuf::from("repainted-as-it-is-handed-back.mkv");
    pin.dpi = 96;
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    let hwnd = HWND(0x1000 as *mut _);
    assert!(
        park_pinned_player(hwnd, (100, 80), false),
        "a move's park stands, and the band is holding the frame the drag took"
    );
    stand_video_frame_for_a_test([0u8, 0, 255, 255].repeat(64), 8, 8);

    let window = RecordedPinWindow::new(0x1000);
    assert!(
        settle_pinned_park_where(&window, true),
        "the same player is standing in the band, so there is nothing to wait for and the swap is \
         taken on the first tick that asks"
    );
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some((100, 80, 420, 320))),
            PinWindowCall::Repaint,
        ],
        "the player's window is put up at where the band is and the band is then painted through it \
         — in that order, and in the same tick, because a paint before the window is up is the \
         desktop showing through a band that has nothing in it, and a paint in a later tick leaves \
         the frame the drag was holding composited over the window for as long as the next repaint \
         is owed"
    );

    // And a park that has already been handed back is not a park: the second settle has nothing to
    // do and must not put a window up or paint a band that is no longer this app's.
    assert!(
        !settle_pinned_park_where(&window, true),
        "a swap that has happened has nothing left to swap"
    );
    assert_eq!(
        window.calls().len(),
        2,
        "and it neither raises the player's window again nor repaints a band that is transparent now"
    );

    forget_video_frame();
    forget_resume_frame();
    forget_pin_park_swap();
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// The same repaint on the arm a resize settles on — the swap a replacement is waited for, which is
/// the one a relaunch is behind and therefore the one whose band has most to be repainted.
///
/// It is the same two calls in the same order; it is driven separately because the two arms are
/// driven by two different facts and a test of one says nothing about the other.
#[test]
fn a_swap_onto_a_replacement_repaints_the_band_too() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    clear_park_swap_arm();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = PathBuf::from("repainted-as-a-replacement-takes-it.mkv");
    pin.dpi = 96;
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    let hwnd = HWND(0x1000 as *mut _);
    assert!(
        park_pinned_player(hwnd, (100, 80), true),
        "a resize's park stands with a replacement behind it"
    );
    stand_video_frame_for_a_test([0u8, 0, 255, 255].repeat(64), 8, 8);
    if let Ok(mut held) = PIN_PARK_SWAP.lock() {
        if let Some(swap) = held.as_mut() {
            swap.since = Some(Instant::now() - PIN_PARK_SWAP_TIMEOUT);
        }
    }

    let window = RecordedPinWindow::new(0x1000);
    assert!(
        settle_pinned_park_where(&window, true),
        "and the bound spent with a window standing in the band hands it the band"
    );
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some((100, 80, 420, 320))),
            PinWindowCall::Repaint,
        ],
        "with the same repaint in the same tick: the placeholder this park was holding is opaque \
         pixels on the band's own surface, and they are still there whatever replaced the window \
         behind them"
    );
    assert!(
        matches!(
            park_swap_last_arm(),
            Some((ParkSwap::TimedOut, waited)) if waited >= PIN_PARK_SWAP_TIMEOUT
        ),
        "and it is the timeout arm that took it, which is where the next person reads the bound from"
    );

    forget_video_frame();
    forget_resume_frame();
    forget_pin_park_swap();
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// **A park with no record of itself is still a park, and it is the one that must never be held for
/// ever.** The flag and the record that says what to wait for are two facts written under two locks
/// on two threads — the flag on the pointer message that begins a drag, the record read by the tick
/// that ends it — and a settle that has already read "no park standing" can give the record up
/// afterwards, after a drag has begun and written its own. The park that begins in that gap has its
/// record taken away under it, and nothing will ever write it again: `park_pinned_player` answers
/// false to a park that is already standing, so the second write it would have made is the one it
/// does not make.
///
/// What is left is a band this app is painting flat for ever, over a film the drag's hold put on
/// pause, with no tick that will ever hand it back. So the arm that answers is the one that needs
/// nothing: a park this app cannot account for is handed back on the first tick that finds a window
/// standing in the band, because the alternative is the picture staying behind an opaque rectangle
/// for the life of the pin. (An extend over a standing park refreshes the record rather than
/// replacing it, so a park begun in the gap still writes — this arm answers the park that has
/// lost its record anyway.)
#[test]
fn a_park_that_lost_its_record_is_handed_back_rather_than_held_for_ever() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = PathBuf::from("left-standing-with-nothing-to-settle-it.mkv");
    pin.dpi = 96;
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    let hwnd = HWND(0x1000 as *mut _);
    assert!(
        park_pinned_player(hwnd, (100, 80), false),
        "the drag's park stands"
    );
    stand_video_frame_for_a_test([0u8, 0, 255, 255].repeat(64), 8, 8);

    // The record goes, under the flag: what a settle that has already read the flag as down and
    // then given the record up on its own leaves behind. Nothing in the code may produce it after
    // this WS — the flag and the record are written in one critical section — and a park that has
    // it anyway is the stuck placeholder this arm answers.
    forget_pin_park_swap();
    assert!(
        pin_player_is_parked(),
        "which is a band painted flat over a film nothing is going to be shown behind"
    );

    let window = RecordedPinWindow::new(0x1000);
    assert!(
        settle_pinned_park_where(&window, true),
        "a park with nothing left to wait for is handed back on the first tick that finds a window \
         standing in the band — there is no replacement it is waiting for and no bound that has \
         anything to do with it"
    );
    assert!(
        !pin_player_is_parked(),
        "and the flag goes down with it, rather than a band this app is painting flat staying up \
         over a paused film for the rest of the pin's life"
    );

    forget_video_frame();
    forget_resume_frame();
    forget_pin_park_swap();
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// **A park and its record are one fact, and this walks every write of either to say so.** They are
/// a flag on the pin and a slot behind a lock of their own, and the tick that ends a park and the
/// pointer message that begins one are two threads — so the only thing that keeps a stuck
/// placeholder out is that no read of one can be answered without the other, which means both are
/// written with the pin held.
///
/// Each step below is one of those writes, in the order the loop and the window procedure make
/// them, and the assertion after each is the whole of what the invariant is: a band this app is
/// painting flat is a park there is a record of, and a park with no record is not one this app is
/// painting.
#[test]
fn a_park_and_its_record_are_never_one_without_the_other() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = PathBuf::from("a-park-and-its-record.mkv");
    pin.dpi = 96;
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    let hwnd = HWND(0x1000 as *mut _);
    let recorded = || {
        PIN_PARK_SWAP
            .lock()
            .map(|held| held.is_some())
            .unwrap_or(false)
    };
    let agrees = || pin_player_is_parked() == recorded();

    assert!(
        !pin_player_is_parked() && !recorded(),
        "a pin at rest has no park and no record"
    );

    // The pointer message: flag up and record written in the same breath, or the swap has nothing
    // to answer with the moment the flag stops being true.
    assert!(park_pinned_player(hwnd, (100, 80), false));
    assert!(agrees(), "the park writes both or neither");

    // A tick that finds the drag still in flight is a park doing its job, and it writes nothing.
    assert!(!settle_pinned_park_where(
        &RecordedPinWindow::new(0x1000),
        false
    ));
    assert!(agrees(), "and a settle with nowhere to go changes neither");

    // The swap: flag down and record given up together.
    assert!(settle_pinned_park_where(
        &RecordedPinWindow::new(0x1000),
        true
    ));
    assert!(
        agrees(),
        "and the swap takes both down, because a record left behind is the next park's `since` and \
         its `replacing` — a resize that inherited a move's record would wait 600 ms for a \
         replacement that was never begun"
    );

    // And a settle of a pin with nothing parked is a no-op rather than a write.
    assert!(!settle_pinned_park_where(
        &RecordedPinWindow::new(0x1000),
        true
    ));
    assert!(
        agrees(),
        "and a settle with no park standing changes neither"
    );

    forget_video_frame();
    forget_resume_frame();
    forget_pin_park_swap();
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// **A gesture ends its own hold, on its own message, and the two halves of it are what a relaunch
/// behind it has to be told.** The hold is what stops the film for the length of a drag, and it was
/// both begun and ended by the tick — the press on one tick and the release on another, so the film
/// played on for the whole of the tick between the hand going down and the tick noticing.
///
/// The cost is not the tick. It is the resize: the release asks for its media to be laid out again
/// at the box it settled on, the relaunch begins a player that is owed the hold, and the hold the
/// gesture took is still standing — so the tick that delivers the owed hold and the tick that lets
/// the gesture go of it are the same tick, and a pause key is a toggle. Two of them is a film
/// playing over a bar with a pause glyph on it.
///
/// So the claim a gesture took is given back by the gesture's own end, and the film's state at that
/// end is the state the relaunch reads: a film that was playing is playing, so nothing is owed it,
/// and a film the user had paused is owed one hold by the player that replaced it and by nothing
/// else.
#[test]
fn a_release_ends_the_hold_its_own_gesture_took() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let mut transport = PinTransport::default();
    transport.begun(3.0, true, false);
    // The hold a tick has taken, stood in: the key it posted cannot land on this machine, but the
    // claim and the second are what every later reader — the relayout included — goes by. The flag
    // beside it is stood with it because the two are written by one call (see `settle_video_drag_hold`),
    // and a half of that pair is a state no gesture has ever produced.
    transport.drag_held = true;
    transport.held(3.0);
    video_drag_hold_set(true);

    stand_pin(Some(PinnedPreview {
        path: PathBuf::from("a-gesture-that-ended-its-own-hold.mkv"),
        transport,
        ..PinnedPreview::for_test()
    }));

    let mut media = create_loading_media(320, 240);
    media.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }

    let window = RecordedPinWindow::new(0x1000);
    assert!(
        park_pinned_player(HWND(0x1000 as *mut _), (0, 0), false),
        "the drag's park stands and the film is held for the length of the gesture. It is not \
         parked as a resize's is: that asks for a frame to be rendered off the drag's own time \
         (see `prepare_resume_frame`), which is two external processes a test has no business \
         starting — and the resize is in the *drag*, which is what the release below reads"
    );
    with_pin(|pin| {
        pin.dragging = Some(PinDrag {
            from: (0, 0),
            window: (0, 0, 320, 240),
            action: PinDragAction::Resize(PinResize {
                left: false,
                top: false,
                right: true,
                bottom: false,
            }),
            delivered: true,
            carried: (i32::MIN, i32::MIN),
        });
    });
    assert!(
        finish_pin_drag(HWND(0x1000 as *mut _), &window),
        "and the release is a resize's, which is the one that asks for its media again"
    );
    if let Ok(mut request) = PIN_BOX_REQUEST.lock() {
        *request = None;
    }

    let claim = || {
        pin_state().and_then(|pinned| {
            pinned
                .pin()
                .map(|pin| (pin.transport.drag_held, pin.transport.pending_hold))
        })
    };
    assert_eq!(
        claim(),
        Some((false, false)),
        "the release gives the claim back, and it does so on its own message rather than on the next \
         tick: a claim a gesture has ended and nothing else has is a claim the next film to be \
         dragged is answered with — so that drag refuses the hold it exists for, and its release \
         posts the toggle onto a film nobody stopped"
    );
    assert!(
        !video_drag_holding(),
        "and the flag beside it agrees, or the tick would find a gesture it believes it has already \
         let go of and answer it a second time"
    );

    forget_video_frame();
    forget_resume_frame();
    forget_pin_park_swap();
    video_drag_hold_set(false);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// **A hold a relaunch owes is delivered by the relaunch's own player, and never twice.** The two
/// keys are the same key: the gesture's hold is the hold, and a replacement that begins under it is
/// owed it — so a tick that finds both the owed hold and a gesture still holding the film delivers
/// nothing, because the gesture's own end is the delivery and posting both is two toggles on one
/// player, which is a film playing over a bar with a pause glyph on it.
///
/// The claim is the one fact that tells the two apart, and it is the transport's own rather than a
/// flag beside it, which is what makes every path that ends the player it was made against take it
/// along.
#[test]
fn a_hold_a_gesture_is_still_owing_is_not_delivered_a_second_time() {
    assert!(
        pending_hold_delivers(true, false),
        "a hold owed to a replacement with no gesture over it is what the settle exists for: the \
         player that has just begun has no window to be given a key through yet"
    );
    assert!(
        !pending_hold_delivers(true, true),
        "and a gesture that is still holding the film delivers it instead — the gesture's own end is \
         the one key this hold is waiting for, and a second one toggles the player back into playing"
    );
    assert!(
        !pending_hold_delivers(false, true),
        "a claim over a hold that is not owed is a claim over nothing, and it delivers nothing"
    );
    assert!(
        !pending_hold_delivers(false, false),
        "and there is no hold to deliver"
    );
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

/// WS-F R2: a begin over a standing park extends it rather than orphaning the settle.
///
/// Release-then-instant-regrab sticks the frozen placeholder while audio and video play
/// underneath: the second `begin` while a park stands is refused, while the in-flight
/// settle's record is superseded — nobody owns the swap. A second begin must extend the
/// standing park (same generation chain, record kept current, capture kept from the first
/// begin) and the settle must drain whatever generation is current, exactly once.
#[test]
fn a_begin_over_a_standing_park_extends_it_rather_than_orphaning_the_settle() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    clear_park_swap_arm();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();
    let previous_pid = VIDEO_PID.swap(4242, Ordering::SeqCst);

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = PathBuf::from("extended-rather-than-orphaned.mkv");
    pin.dpi = 96;
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();
    clear_park_trace();

    let hwnd = HWND(0x1000 as *mut _);
    assert!(
        park_pinned_player(hwnd, (100, 80), true),
        "a resize's park stands with a replacement behind it"
    );

    // The release's relaunch, stood in for: the replacement is begun behind the cover, so the
    // pid the park captured for is no longer the player on screen.
    VIDEO_PID.store(4243, Ordering::SeqCst);
    clear_park_trace();

    assert!(
        park_pinned_player(hwnd, (100, 80), false),
        "a begin over a standing park extends it: refusing it leaves the in-flight settle's \
         record superseded with nobody owning the swap, which is the frozen placeholder standing \
         over a live player until the next release"
    );
    assert_eq!(
        park_trace(),
        vec!["extend"],
        "and the extend keeps the first begin's capture: the window was fully visible then and \
         there is nothing to read once it has gone, so a second capture, paint and hide would \
         only re-hide a window that is already hidden"
    );
    assert!(
        PIN_PARK_SWAP.lock().is_ok_and(|held| held
            .map(|swap| swap.player == 4243 && swap.replacing)
            .unwrap_or(false)),
        "and the record is kept current: the player named is the one the relaunch began, with the \
         replacement still awaited — a record left naming the retired player is a swap nobody owns"
    );

    // The final release's settle: whatever generation is current drains, exactly once.
    if let Ok(mut held) = PIN_PARK_SWAP.lock() {
        if let Some(swap) = held.as_mut() {
            swap.since = Some(Instant::now() - PIN_PARK_SWAP_TIMEOUT);
        }
    }
    let window = RecordedPinWindow::new(0x1000);
    assert!(
        settle_pinned_park_where(&window, true),
        "the settle drains the current generation once the replacement's window is standing in \
         the band"
    );
    assert!(
        !pin_player_is_parked(),
        "so no placeholder is left standing over the live player"
    );
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some((100, 80, 420, 320))),
            PinWindowCall::Repaint,
        ],
        "with the window put up at the final box and the band painted through it in the same tick"
    );
    assert!(
        !settle_pinned_park_where(&window, true),
        "and a second settle swaps nothing: one cover, one swap"
    );

    forget_video_frame();
    forget_resume_frame();
    forget_pin_park_swap();
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// WS-F R3: what the swap is watched for — whether the cover was still up when the player's
/// window was put back, and at what box. No call list can show it: the recorder records the
/// place either way, so the witness reads the flag from inside the place itself.
struct CoverOrderWindow {
    parked_when_placed: Mutex<Option<bool>>,
    placed_band: Mutex<Option<Option<ScreenRegion>>>,
    parked_when_repainted: Mutex<Option<bool>>,
    calls: Mutex<Vec<&'static str>>,
}

impl CoverOrderWindow {
    fn new() -> Self {
        Self {
            parked_when_placed: Mutex::new(None),
            placed_band: Mutex::new(None),
            parked_when_repainted: Mutex::new(None),
            calls: Mutex::new(Vec::new()),
        }
    }

    fn record(&self, call: &'static str) {
        if let Ok(mut calls) = self.calls.lock() {
            calls.push(call);
        }
    }
}

impl PinWindow for CoverOrderWindow {
    fn hwnd(&self) -> isize {
        0x1000
    }

    fn pointer(&self) -> Option<(i32, i32)> {
        None
    }

    fn window_box(&self, _hwnd: isize) -> Option<ScreenRegion> {
        None
    }

    fn capture(&self, _hwnd: isize) {}

    fn release_capture(&self, _hwnd: isize) {}

    fn set_focusable(&self, _hwnd: isize, _focusable: bool) {}

    fn set_focus(&self, _hwnd: isize) {}

    fn set_foreground(&self, _hwnd: isize) {}

    fn hide_pin_windows(&self) {}

    fn hide_pin_bubble(&self) {}

    fn unpark_player_window(&self, band: Option<ScreenRegion>) {
        if let Ok(mut parked) = self.parked_when_placed.lock() {
            *parked = Some(pin_player_is_parked());
        }
        if let Ok(mut placed) = self.placed_band.lock() {
            *placed = Some(band);
        }
        self.record("unpark");
    }

    fn repaint(&self) {
        if let Ok(mut parked) = self.parked_when_repainted.lock() {
            *parked = Some(pin_player_is_parked());
        }
        self.record("repaint");
    }

    fn post(&self, _hwnd: isize, _message: u32) {}
}

/// WS-F R3, same-player exit: the unpark places the player at the final box before the flag
/// goes down, and repaints in the same tick.
#[test]
fn an_unpark_places_the_player_before_it_drops_the_flag_on_the_same_player() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    clear_park_swap_arm();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();
    let previous_pid = VIDEO_PID.swap(4242, Ordering::SeqCst);

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = PathBuf::from("placed-before-it-is-dropped.mkv");
    pin.dpi = 96;
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    let hwnd = HWND(0x1000 as *mut _);
    assert!(
        park_pinned_player(hwnd, (100, 80), false),
        "a move's park stands over the player that was playing"
    );
    // The hand carries the window elsewhere while the park stands; every place that would
    // have told the player is answered out of the hand, so the final box is read at uncover.
    with_pin(|pin| pin.content = (200, 160, 520, 400));

    let window = CoverOrderWindow::new();
    assert!(
        settle_pinned_park_where(&window, true),
        "the same player is standing in the band, so the swap is taken at once"
    );
    assert_eq!(
        window
            .parked_when_placed
            .lock()
            .ok()
            .and_then(|parked| *parked),
        Some(true),
        "the player's window is put back while the cover is still up: dropping the flag first \
         leaves a transparent band over a window whose first composited frame has not landed"
    );
    assert_eq!(
        window.placed_band.lock().ok().and_then(|band| *band),
        Some(Some((200, 160, 520, 400))),
        "at the final box the drag settled on, not at the box the drag began from"
    );
    assert_eq!(
        window
            .parked_when_repainted
            .lock()
            .ok()
            .and_then(|parked| *parked),
        Some(false),
        "and the band is repainted after the flag goes down, in the same tick: the pixels the \
         compositor is holding are still the placeholder"
    );
    assert_eq!(
        window.calls.lock().ok().map(|calls| calls.clone()),
        Some(vec!["unpark", "repaint"]),
        "place, flag down, repaint — strictly ordered, same tick"
    );

    forget_video_frame();
    forget_resume_frame();
    forget_pin_park_swap();
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// WS-F R3, replacement exit: the same order when the swap hands the band to a relaunch.
#[test]
fn an_unpark_places_the_player_before_it_drops_the_flag_on_a_replacement() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    clear_park_swap_arm();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();
    let previous_pid = VIDEO_PID.swap(4242, Ordering::SeqCst);

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = PathBuf::from("placed-before-it-is-dropped-behind-a-relaunch.mkv");
    pin.dpi = 96;
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    let hwnd = HWND(0x1000 as *mut _);
    assert!(
        park_pinned_player(hwnd, (100, 80), true),
        "a resize's park stands with a replacement behind it"
    );
    with_pin(|pin| pin.content = (200, 160, 560, 430));

    // The release's relaunch, stood in for, and the bound spent waiting for its window.
    VIDEO_PID.store(4243, Ordering::SeqCst);
    if let Ok(mut held) = PIN_PARK_SWAP.lock() {
        if let Some(swap) = held.as_mut() {
            swap.since = Some(Instant::now() - PIN_PARK_SWAP_TIMEOUT);
        }
    }

    let window = CoverOrderWindow::new();
    assert!(
        settle_pinned_park_where(&window, true),
        "the bound spent with a window standing in the band hands it the band"
    );
    assert_eq!(
        window
            .parked_when_placed
            .lock()
            .ok()
            .and_then(|parked| *parked),
        Some(true),
        "the replacement's window is put up while the cover is still up, on this exit too"
    );
    assert_eq!(
        window.placed_band.lock().ok().and_then(|band| *band),
        Some(Some((200, 160, 560, 430))),
        "at the final box the resize settled on"
    );
    assert_eq!(
        window
            .parked_when_repainted
            .lock()
            .ok()
            .and_then(|parked| *parked),
        Some(false),
        "and the repaint follows the flag down in the same tick"
    );

    forget_video_frame();
    forget_resume_frame();
    forget_pin_park_swap();
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// WS-F R4: the gesture's hold is taken before the park it belongs to, on both entries.
///
/// The hold is a pause key and the park below reads the picture off the screen, so a hold
/// that lands a tick later is a band holding a frame the player has already moved on from.
/// Both entries go through one call, and the trace is the seam that says which came first.
#[test]
fn a_gestures_hold_is_taken_before_the_park_it_belongs_to() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();
    let holding = video_drag_holding();

    for (name, action) in [
        ("move", PinDragAction::Move),
        (
            "resize",
            PinDragAction::Resize(PinResize {
                left: false,
                top: false,
                right: true,
                bottom: true,
            }),
        ),
    ] {
        let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
        pin.path = PathBuf::from("hold-before-capture.mkv");
        pin.dpi = 96;
        stand_pin(Some(pin));
        forget_pin_park_swap();
        forget_video_frame();
        forget_resume_frame();
        clear_park_trace();

        let window = a_window_at((100, 80, 420, 320));
        begin_pin_drag(HWND(0x1000 as *mut _), &window, action, true);
        assert_eq!(
            park_trace(),
            vec!["hold", "capture", "paint", "hide"],
            "a {name}'s hold is taken before the park reads the frame it is going to hold"
        );

        video_drag_hold_set(false);
    }

    video_drag_hold_set(holding);
    forget_video_frame();
    forget_resume_frame();
    forget_pin_park_swap();
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// WS-F R1: a seek parks the stale frame it is about to replace.
///
/// The seekbar seek flashes desktop because the seek relaunch runs with no park cover: the
/// old player is parked for retirement at relaunch with the hole transparent and the
/// replacement up later. Parking the stale frame at seek-start puts the relaunch behind a
/// cover for its single swap — while the film keeps playing (no hold: a scrub is not a drag
/// and pausing it is the defect rather than the fix), and with no resume render asked (the
/// playhead is moving, so a frame rendered at the second the hand started from is a frame of
/// the wrong second).
#[test]
fn a_seek_parks_the_stale_frame_it_is_about_to_replace() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    clear_park_swap_arm();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();
    let previous_pid = VIDEO_PID.swap(4242, Ordering::SeqCst);

    let mut transport = PinTransport::default();
    transport.begun(3.0, true, false);
    stand_pin(Some(PinnedPreview {
        path: PathBuf::from("parked-for-a-seek.mkv"),
        transport,
        ..PinnedPreview::for_test()
    }));

    let mut media = create_loading_media(320, 240);
    media.media_type = MediaType::Video;
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }

    let playing = || {
        pin_state().and_then(|state| {
            state.pin().map(|pin| {
                (
                    pin.transport.paused_at,
                    pin.transport.pending_hold,
                    pin.transport.drag_held,
                    pin.transport.started.is_some(),
                )
            })
        })
    };
    assert_eq!(
        playing(),
        Some((None, false, false, true)),
        "the premise: the film is playing and held by nothing"
    );

    clear_park_trace();
    let hwnd = HWND(0x1000 as *mut _);
    assert!(
        park_pinned_player_for_seek(hwnd, (100, 80)),
        "a seek parks the stale frame at seek-start, so the relaunch the release makes runs \
         behind a cover rather than through a transparent hole"
    );
    assert_eq!(
        park_trace(),
        vec!["capture", "paint", "hide"],
        "with the same three steps a drag's park takes, in the same order"
    );
    assert!(
        PIN_PARK_SWAP.lock().is_ok_and(|held| held
            .map(|swap| swap.player == 4242 && swap.replacing)
            .unwrap_or(false)),
        "and the record awaits a replacement: the seek ends in a player that has decoded nothing"
    );
    assert_eq!(
        playing(),
        Some((None, false, false, true)),
        "while the film keeps playing: the cover carries no hold, so a scrub never pauses"
    );
    assert!(
        !video_drag_holding(),
        "and the gesture flag is untouched for the same reason"
    );

    // A scrub aims without relaunching: rapid steps coalesce into the one relaunch the release
    // makes at the latest playhead, so intermediate seconds cost no player.
    update_pin_transport(|transport| transport.seeking = Some(90.0));
    update_pin_transport(|transport| transport.seeking = Some(95.0));
    assert!(
        pin_player_is_parked()
            && VIDEO_PID.load(Ordering::SeqCst) == 4242
            && playing() == Some((None, false, false, true)),
        "aiming moves only the second under the hand: no relaunch, no hold, cover standing"
    );

    // A second press over the standing seek cover extends rather than re-captures.
    clear_park_trace();
    assert!(
        park_pinned_player_for_seek(hwnd, (100, 80)),
        "a second seek press extends the standing cover"
    );
    assert_eq!(
        park_trace(),
        vec!["extend"],
        "keeping the first press's capture rather than reading the screen with nothing behind it"
    );

    forget_video_frame();
    forget_resume_frame();
    forget_pin_park_swap();
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// WS-F R1, scrub: a seek's cover waits for its relaunch rather than spending its bound.
///
/// A scrub aims without relaunching, and a slow hand holds past the swap bound — so the bound
/// cannot run from the press. Until the release relaunches behind the cover there is nothing to
/// swap to: the settle holds even with a window standing in the band, stamps no budget, and
/// drains once the relaunch is begun.
#[test]
fn a_seek_cover_waits_for_its_relaunch_rather_than_spending_its_bound() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    clear_park_swap_arm();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();
    let previous_pid = VIDEO_PID.swap(4242, Ordering::SeqCst);

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = PathBuf::from("waiting-for-its-relaunch.mkv");
    pin.dpi = 96;
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    let hwnd = HWND(0x1000 as *mut _);
    assert!(
        park_pinned_player_for_seek(hwnd, (100, 80)),
        "a seek parks its cover at seek-start"
    );
    // The scrub itself, which is what this test is about: the hold below is a tick *mid-scrub*, so
    // the aim the press stored has to be standing. A cover with no scrub in flight is a cover whose
    // release has already come, and the tick does not hold that one (see
    // `cover_relaunch_is_still_possible`).
    update_pin_transport(|transport| transport.seeking = Some(90.0));
    assert!(
        seek_cover_is_waiting(),
        "which waits for the release's relaunch: no player has been begun behind it yet"
    );

    let window = RecordedPinWindow::new(0x1000);
    assert!(
        !settle_pinned_park_where(&window, true),
        "so a tick mid-scrub holds even with a window standing in the band: there is nothing to \
         swap to, and swapping to the player the cover stands over would spend the cover before \
         the relaunch it was parked for"
    );
    assert!(
        pin_player_is_parked() && window.calls().is_empty(),
        "and the hold asks nothing of any window and paints nothing: a scrub held past the bound \
         must neither spend it nor show the player it is still aiming over"
    );
    assert!(
        PIN_PARK_SWAP
            .lock()
            .is_ok_and(|held| held.as_ref().is_some_and(|swap| swap.since.is_none())),
        "with the bound left unstamped: it runs from the relaunch, not from the press"
    );
    // The release's relaunch, stood in for, and the bound spent waiting for its window.
    VIDEO_PID.store(4243, Ordering::SeqCst);
    assert!(
        !seek_cover_is_waiting(),
        "a relaunch begun behind the cover ends the wait"
    );
    if let Ok(mut held) = PIN_PARK_SWAP.lock() {
        if let Some(swap) = held.as_mut() {
            swap.since = Some(Instant::now() - PIN_PARK_SWAP_TIMEOUT);
        }
    }
    assert!(
        settle_pinned_park_where(&window, true),
        "and the bound spent with the replacement's window standing in the band hands it the band"
    );
    assert!(!pin_player_is_parked(), "exactly once: one cover, one swap");
    assert_eq!(
        window.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some((100, 80, 420, 320))),
            PinWindowCall::Repaint,
        ],
        "with the window put up at the final box and the band painted through it in the same tick"
    );

    forget_video_frame();
    forget_resume_frame();
    forget_pin_park_swap();
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// WS-F R1, abandon: a seek's cover ended without a relaunch is handed back, not held.
///
/// A scrub that never releases — a capture stolen mid-aim, a move let go of over it — begins
/// no player behind the cover, so the settle would hold it for ever. Both ends hand the band
/// back to the player the cover stands over instead.
#[test]
fn a_seek_cover_ended_without_a_relaunch_is_handed_back() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    clear_park_swap_arm();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();
    let previous_pid = VIDEO_PID.swap(4242, Ordering::SeqCst);

    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = PathBuf::from("handed-back-without-a-relaunch.mkv");
    pin.dpi = 96;
    stand_pin(Some(pin));
    forget_pin_park_swap();
    forget_video_frame();
    forget_resume_frame();

    let hwnd = HWND(0x1000 as *mut _);
    assert!(
        park_pinned_player_for_seek(hwnd, (100, 80)),
        "a seek parks its cover at seek-start"
    );
    update_pin_transport(|transport| transport.seeking = Some(90.0));

    // The capture stolen mid-aim: the aim goes with it, and the cover with the aim.
    let stolen = RecordedPinWindow::new(0x1000);
    pin_capture_lost(&stolen);
    assert!(
        !pin_player_is_parked(),
        "a seek abandoned mid-aim never relaunches, so the cover is handed back rather than held \
         for a player that is never begun"
    );
    assert_eq!(
        stolen.calls(),
        vec![
            PinWindowCall::UnparkPlayerWindow(Some((100, 80, 420, 320))),
            PinWindowCall::Repaint,
        ],
        "to the player it covers, at the band it stands in, painted through in the same tick"
    );

    // The move let go of over a scrub: a gesture that relaunches nothing ends the wait the same
    // way. A resize always relaunches below the cover and needs no such end.
    let mut pin = overlay_pin((100, 80, 420, 320), PinChrome::always());
    pin.path = PathBuf::from("handed-back-by-a-move.mkv");
    pin.dpi = 96;
    stand_pin(Some(pin));
    forget_pin_park_swap();
    assert!(
        park_pinned_player_for_seek(hwnd, (100, 80)),
        "and the cover stands again over the next seek"
    );
    update_pin_transport(|transport| transport.seeking = Some(95.0));
    with_pin(|pin| {
        pin.dragging = Some(PinDrag {
            from: (0, 0),
            window: (100, 80, 420, 320),
            action: PinDragAction::Move,
            delivered: true,
            carried: (i32::MIN, i32::MIN),
        });
    });

    let released = RecordedPinWindow::new(0x1000);
    assert!(
        finish_pin_drag(hwnd, &released),
        "a move is let go of through its release"
    );
    assert!(
        !pin_player_is_parked(),
        "which hands a waiting seek cover back: a move relaunches nothing, so nothing is coming \
         for the settle to swap to"
    );
    assert!(
        released
            .calls()
            .contains(&PinWindowCall::UnparkPlayerWindow(Some((
                100, 80, 420, 320
            )))),
        "at the band the move left behind"
    );

    forget_video_frame();
    forget_resume_frame();
    forget_pin_park_swap();
    VIDEO_PID.store(previous_pid, Ordering::SeqCst);
    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

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

/// The two arms a park's end can be taken on, and the answer that is not one of them.
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
/// **And no window at all is not a third arm, which is the hole this whole arrangement exists not
/// to leave.** The band is transparent because a player's window stands in it, so a band handed
/// back with nothing behind it is the desktop — and a replacement within its bound of publishing a
/// window is answered out of the hand by every place that could put it up, so swapping on the bound
/// alone would take the parked flag down and show nothing.
#[test]
fn a_band_is_handed_back_on_a_player_or_on_a_wait_and_never_on_neither() {
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

    // No window at all is the case the bound does *not* cover, and the one it most certainly must
    // not: there is nothing to hand the band to, and the placeholder it is holding is opaque.
    assert_eq!(
        park_swap_arm(false, false, Duration::ZERO),
        None,
        "a player with no window cannot have presented"
    );
    assert_eq!(
        park_swap_arm(false, true, PIN_PARK_SWAP_TIMEOUT),
        None,
        "and the bound is not a reason to swap one that never published a window: the swap takes \
         the parked flag down and the window up together, and with no window up it is the desktop"
    );
    assert_eq!(
        park_swap_arm(false, true, PIN_PARK_SWAP_TIMEOUT * 4),
        None,
        "nor is four times the bound — the placeholder stands until there is a player to see \
         through it, and the wait is extended rather than restarted so that a window arriving late \
         is handed the band on the first tick that finds it"
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
        ],
        "and the player's window is put up at where the band is, rather than at where the drag \
         began"
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

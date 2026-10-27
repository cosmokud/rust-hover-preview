use super::*;

/// A band tall enough to be worth splitting is scaled by more than one thread, row for row where the
/// single-threaded walk put them and nowhere else.
///
/// Two things are asked of it that a band of a handful of rows cannot ask. The rows are split across
/// threads at `resample_threads`' floor, so a band this tall is written by several walks at once; and
/// the band is *not* the whole surface, so a walk that took its rows from the surface rather than from
/// the band would read the frame's box a band's height out, which is a picture drawn at the wrong row
/// of itself rather than an empty band (see `BandTarget::frame`).
///
/// The frame is its own size scaled one for one, so every row of the source is a row of the band and
/// the answer is the source itself: a row of the destination is a value written down here and not an
/// arithmetic result of the sampler's.
#[test]
fn a_band_tall_enough_to_split_is_scaled_row_for_row_where_the_band_is() {
    let (width, height) = (4u32, 80u32);
    let (band_height, origin_y) = (64u32, 8u32);
    assert!(
        resample_threads(band_height as usize) > 1,
        "a band of {band_height} rows is one the sampler splits across threads"
    );

    let source_height = band_height;
    let mut media = Vec::new();
    for row in 0..source_height {
        for _ in 0..width {
            media.extend_from_slice(&[row as u8 * 4, row as u8 * 4, row as u8 * 4, 255]);
        }
    }

    let surface = DibSurface::create(width, height).expect("a surface to scale into");
    let out =
        unsafe { std::slice::from_raw_parts_mut(surface.bits(), (width * height * 4) as usize) };

    resample_into_band(
        &media,
        (width, source_height),
        (width, band_height),
        // Never reached: an opaque frame covers the box it lands in.
        TransparentBackground::Black,
        true,
        BandTarget {
            out,
            width,
            origin_y,
            height: band_height,
            frame: (0, 0, width as i32, band_height as i32),
            exact: (0.0, 0.0, width as f32, band_height as f32),
        },
    );

    for row in 0..height as usize {
        let in_the_band =
            (origin_y as usize..origin_y as usize + band_height as usize).contains(&row);
        let line = &out[row * width as usize * 4..(row + 1) * width as usize * 4];

        if !in_the_band {
            assert_eq!(
                line,
                vec![0u8; width as usize * 4],
                "row {row} is outside the band and is nothing at all"
            );
            continue;
        }

        let source_row = (row - origin_y as usize) as u8 * 4;
        for pixel in line.as_chunks::<4>().0 {
            assert_eq!(
                pixel,
                &[source_row, source_row, source_row, 255],
                "row {row} is row {} of the picture",
                row - origin_y as usize
            );
        }
    }
}

/// A band the frame is not the size of is filled by the frame, sampled: it is what a pinned
/// window draws while an edge of it is being dragged, and what it may never draw is the frame
/// standing at its own size with nothing around it.
#[test]
fn a_frame_drawn_at_a_size_of_its_own_is_sampled_through_the_place_it_really_lands_on() {
    // A picture of two columns, black and white, drawn into a band four columns wide at a place of a
    // quarter of a pixel: the band's columns are the blends the place asks for, and not the blends a box
    // of whole pixels — which is where this picture begins — would give.
    //
    // Worked example, the picture's left edge a quarter of a pixel in from the band's left and the band
    // twice the picture's own width: column *i* of the band samples the picture at `(i + 0.25) / 2 - 0.5`,
    // which is 0.125, 0.625 and 1.125 for the three columns the picture's own pixels reach — a line whose
    // grey goes 32, 159, 255 in 256ths — where the box alone would sample 0.25 and 0.75 and give 64 and
    // 191. Measured through `render_pinned_preview_at`, that difference is the whole of what a picture in
    // motion owes the wheel's own anchor (see `PinZoom::picture_place`).
    let picture = [0u8, 0, 0, 255, 255, 255, 255, 255];
    let mut out = vec![0u8; 4 * 4];

    resample_into_band(
        &picture,
        (2, 1),
        (4, 1),
        TransparentBackground::Black,
        true,
        BandTarget {
            out: &mut out,
            width: 4,
            origin_y: 0,
            height: 1,
            frame: (0, 0, 4, 1),
            exact: (0.25, 0.0, 4.0, 1.0),
        },
    );

    let grey = |column: usize| out[column * 4];
    for (column, wanted) in [(0usize, 0u8), (1, 32), (2, 159), (3, 255)] {
        assert!(
            grey(column).abs_diff(wanted) <= 1,
            "column {column} of the band is the picture sampled where it really begins: {} against {wanted}",
            grey(column)
        );
    }
}

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
                width: size.0,
                origin_y: 0,
                height: size.1,
                frame: (0, 0, size.0 as i32, size.1 as i32),
                exact: (0.0, 0.0, size.0 as f32, size.1 as f32),
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

/// And the same frame with every pixel of it opaque, which is the kind every photograph and every
/// screenshot the wheel is tried on is: the frame fills the band's rows and nowhere else, and every
/// pixel of it is opaque.
///
/// The last is the part worth asking about: a layered window is composited from premultiplied
/// coverage, so an opaque frame's alpha byte is 255 everywhere — and it comes from the frame itself
/// rather than from the road that drew it (see `pixels_are_opaque`).
#[test]
fn an_opaque_frame_of_another_size_fills_the_band_it_is_drawn_in() {
    let (width, height, band_height, origin_y) = (8u32, 6u32, 4u32, 2u32);
    let surface = DibSurface::create(width, height).expect("a surface to fill");
    let out =
        unsafe { std::slice::from_raw_parts_mut(surface.bits(), (width * height * 4) as usize) };

    let media = [32u8, 64, 128, 255].repeat(4);
    compose_media_into_band(
        &media,
        2,
        2,
        (width, band_height),
        // Never reached: an opaque frame covers the box it lands in.
        TransparentBackground::Black,
        true,
        BandTarget {
            out,
            width,
            origin_y,
            height: band_height,
            frame: (0, 0, width as i32, band_height as i32),
            exact: (0.0, 0.0, width as f32, band_height as f32),
        },
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

/// A picture drawn larger than the window it is in is the size it is *drawn* at, and not the size of
/// the part of it the window has room for.
///
/// That is the whole of what a zoom past the fit-to-screen cap is: the window stops and the picture
/// goes on growing inside it, so what a repaint is handed is a frame larger than its band in one or
/// both directions. A road that stretched the frame into the rows still on screen would be a picture
/// squashed by however much of it had been carried off — and a picture that changes shape as the
/// hand carries it is the one thing a zoom may never be.
///
/// It is the shape every photograph and every screenshot a zoom is tried on takes: an opaque frame
/// drawn larger than the window it is in (see `compose_media_into_band`).
#[test]
fn a_frame_bigger_than_its_band_is_drawn_at_the_size_it_is_drawn_at() {
    // Four rows, the last of them white and the rest black. Drawn eight times as tall as the band
    // and centred on it, the band is shown the source's *middle* two rows, which are black; a frame
    // stretched into the rows the band has room for would be shown the source's last row instead,
    // which is white.
    let (source_width, source_height) = (2u32, 4u32);
    let mut media = Vec::new();
    for row in 0..source_height {
        let value = if row == source_height - 1 { 255u8 } else { 0 };
        for _ in 0..source_width {
            media.extend_from_slice(&[value, value, value, 255]);
        }
    }

    let (width, height) = (2u32, 8u32);
    let surface = DibSurface::create(width, height).expect("a surface to fill");
    let out =
        unsafe { std::slice::from_raw_parts_mut(surface.bits(), (width * height * 4) as usize) };

    // A band the whole of the surface, which is what a picture's band is: a pin whose chrome is
    // drawn over its media is media the whole of its window down (see `pinned_band_rows`).
    compose_media_into_band(
        &media,
        source_width,
        source_height,
        (width, height),
        // Never reached: an opaque frame covers the box it lands in.
        TransparentBackground::Black,
        true,
        BandTarget {
            out,
            width,
            origin_y: 0,
            height,
            frame: (0, (height as i32 - 32) / 2, source_width as i32, 32),
            exact: (
                0.0,
                ((height as i32 - 32) / 2) as f32,
                source_width as f32,
                32.0,
            ),
        },
    );

    for (row, line) in out.chunks(width as usize * 4).enumerate() {
        for pixel in line.as_chunks::<4>().0 {
            assert!(
                pixel[0] < 16,
                "row {row} of the band is the middle of the picture and not the end of it: {pixel:?}"
            );
        }
    }
}

/// The two backdrops a pinned picture's band can be filled with, stood where a test wants them and put
/// back when the guard is dropped — the house pattern for a setting a test has to move (see
/// `PausePreviewSettings`).
///
/// Both are stood together because the test's whole claim is that the *right one of the two* is read: a
/// test that moved one of them would pass with the branch missing entirely.
struct BandBackdrops {
    image: TransparentBackground,
    zoom: TransparentBackground,
}

impl BandBackdrops {
    fn stood_at(image: TransparentBackground, zoom: TransparentBackground) -> Self {
        let mut config = crate::CONFIG.lock().expect("the configuration");
        let was = BandBackdrops {
            image: config.image_background,
            zoom: config.zoom_background,
        };
        config.image_background = image;
        config.zoom_background = zoom;

        was
    }
}

impl Drop for BandBackdrops {
    fn drop(&mut self) {
        if let Ok(mut config) = crate::CONFIG.lock() {
            config.image_background = self.image;
            config.zoom_background = self.zoom;
        }
    }
}

/// A zoomed picture with no transparency in it draws its band over the tray's own backdrop for a zoom;
/// a picture with a transparent pixel in it, and every picture that fills its window, over the backdrop
/// its own kind asks for.
///
/// A picture's transparency is the picture's own business whatever the wheel has done to it, so the
/// `Background` half's setting — `Transparent` included — stands behind the transparent pixels and the
/// band beside them alike; the band is the window's rather than the picture's only where there is no
/// transparency for the picture's setting to answer (see `Image → Pin Mode → Zoom Background` in the
/// tray and `pinned_band_backdrop`).
#[test]
fn a_zoomed_pictures_band_is_the_zooms_backdrop_until_the_picture_has_transparency() {
    let _stood = BandBackdrops::stood_at(
        TransparentBackground::Checkerboard,
        TransparentBackground::White,
    );

    assert_eq!(
        pinned_band_backdrop(MediaType::StaticImage, true, true),
        TransparentBackground::White,
        "an opaque picture the wheel has taken past its window is drawn over the zoom's own backdrop"
    );
    assert_eq!(
        pinned_band_backdrop(MediaType::StaticImage, true, false),
        TransparentBackground::Checkerboard,
        "and a picture with transparency in it over the one its own kind asks for, the width of the band included"
    );
    assert_eq!(
        pinned_band_backdrop(MediaType::StaticImage, false, false),
        TransparentBackground::Checkerboard,
        "which is the backdrop it was drawn over at the size of its window all along"
    );
    assert_eq!(
        pinned_band_backdrop(MediaType::Dds, true, true),
        TransparentBackground::White,
        "and the band is the window's whatever kind of picture the window is standing on"
    );
}

/// A picture drawn at the size its own frame is — which is every picture a zoom has settled on — is a
/// crop of that frame and nothing else: no stretch, no sample, and no walk over the rows the band has
/// not got.
///
/// It is the road a deep zoom's pan takes, and what it is worth is what the other two cost: the
/// stretch hands GDI the whole frame through a surface of this app's own, and the sampler asks a
/// per-pixel question of a frame that has nothing to be sampled. A hundred-megabyte frame carried by a
/// hand is a hundred megabytes copied per move of it under the one, and the band's own rows under this.
#[test]
fn a_frame_drawn_at_its_own_size_is_a_crop_of_it() {
    // A frame of sixteen rows drawn at its own size, standing in a band of eight: the band shows the
    // frame's middle eight rows, which is where the picture's own middle is.
    let (width, height) = (4u32, 8u32);
    let mut out = vec![0u8; (width * height * 4) as usize];

    // Every row of the frame is its own shade, so a crop that took the wrong rows says so.
    let mut media = Vec::new();
    for row in 0..16u32 {
        for _ in 0..width {
            media.extend_from_slice(&[row as u8 * 16, 0, 0, 255]);
        }
    }

    let drawn = |left: i32, out: &mut [u8]| {
        copy_frame_into_band(
            &media,
            (width, 16),
            true,
            &mut BandTarget {
                out,
                width,
                origin_y: 0,
                height,
                frame: (left, -4, width as i32, 16),
                exact: (left as f32, -4.0, width as f32, 16.0),
            },
        )
    };

    assert!(drawn(0, &mut out), "a frame of its own size is one to copy");
    for row in 0..height as usize {
        let pixel = &out[row * width as usize * 4..row * width as usize * 4 + 4];
        assert_eq!(
            pixel[0],
            (row + 4) as u8 * 16,
            "row {row} of the band is row {} of the picture",
            row + 4
        );
    }

    // A picture carried off the side fills only the columns of the band it still covers, and the ones
    // it has left are the caller's — which is the backdrop a hand sees beside a picture it has carried.
    let mut out = vec![0u8; (width * height * 4) as usize];
    assert!(drawn(-2, &mut out), "a frame carried off the side is one");
    for row in 0..height as usize {
        let line = &out[row * width as usize * 4..(row + 1) * width as usize * 4];
        assert_eq!(
            line.as_chunks::<4>().0[0],
            [(row + 4) as u8 * 16, 0, 0, 255],
            "the columns it still covers are the picture's"
        );
        assert_eq!(
            line.as_chunks::<4>().0[3],
            [0, 0, 0, 0],
            "and the columns it has left are exactly as the caller left them"
        );
    }

    // And a frame that is not the size of the box it lands in is not this road's: it is one the band
    // has to be filled *from*, and the sampler is the one that knows how.
    let mut out = vec![0u8; (width * height * 4) as usize];
    assert!(
        !copy_frame_into_band(
            &media,
            (width, 16),
            true,
            &mut BandTarget {
                out: &mut out,
                width,
                origin_y: 0,
                height,
                frame: (0, -4, width as i32, 8),
                exact: (0.0, -4.0, width as f32, 8.0),
            },
        ),
        "a frame drawn at another size is a scale and not a copy"
    );

    // And neither is a frame with an alpha channel, whatever size it is: what stands behind it is a
    // per-pixel question and this is a copy.
    let mut out = vec![0u8; (width * height * 4) as usize];
    assert!(
        !copy_frame_into_band(
            &media,
            (width, 16),
            false,
            &mut BandTarget {
                out: &mut out,
                width,
                origin_y: 0,
                height,
                frame: (0, -4, width as i32, 16),
                exact: (0.0, -4.0, width as f32, 16.0),
            },
        ),
        "a frame to be composited is the sampler's"
    );
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

/// One pixel of a painted surface, by row and column: the windows these tests paint are of two
/// sizes, so a pixel cannot be spelled out as an offset into a fixed stride.
fn surface_pixel(surface: &[u8], window: (u32, u32), row: usize, column: usize) -> &[u8] {
    let stride = window.0 as usize * 4;
    &surface[row * stride + column * 4..row * stride + column * 4 + 4]
}

#[test]
fn a_picture_the_wheel_has_taken_past_its_window_is_drawn_where_the_zoom_says_it_is() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let content = (100, 80, 420, 320);
    let (window_width, window_height) = region_size(content);
    // A quarter of a pixel of pan, so the place and the box differ by exactly the fraction this is about:
    // the box begins at -123 and the place at -123.25.
    let pan = 37.25;
    let (drawn_width, drawn_height) = ((window_width * 2) as u32, (window_height * 2) as u32);

    let mut pin = overlay_pin(
        content,
        PinChrome {
            caption: false,
            bar: false,
            until: None,
        },
    );
    pin.zoom = Some(PinZoom::still(
        (drawn_width as f32, drawn_height as f32),
        (pan, 0.0),
        centred_on(pin.content),
    ));
    stand_pin(Some(pin));

    let hwnd = HWND(0x1000 as *mut _);
    let row = window_height as usize / 2;
    let painted = |media: MediaData| {
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = Some(media);
        }
        unsafe { render_pinned_preview_at(hwnd, content.0, content.1) };

        let bits =
            ensure_layered_surface(hwnd.0 as isize, window_width as u32, window_height as u32)
                .expect("the surface the paint drew into");
        let surface = unsafe {
            std::slice::from_raw_parts(bits, window_width as usize * window_height as usize * 4)
        };

        (0..window_width as usize)
            .map(|column| {
                surface_pixel(
                    surface,
                    (window_width as u32, window_height as u32),
                    row,
                    column,
                )[0]
            })
            .collect::<Vec<u8>>()
    };

    // The frame a zoom's own decode hands back for a box: the picture at the size of the box, with every
    // column its own shade, so a crop taken from the wrong end of it says so. The box begins at -123, so
    // column *x* of the window is the picture's own column `x + 123`.
    let mut at_the_box = Vec::with_capacity((drawn_width * drawn_height * 4) as usize);
    for _ in 0..drawn_height {
        for column in 0..drawn_width {
            at_the_box.extend_from_slice(&[((column % 250) + 1) as u8, 0, 0, 255]);
        }
    }

    let mut media = create_loading_media(drawn_width, drawn_height);
    media.frames[0] = Arc::new(ImageFrame::new(at_the_box, drawn_width, drawn_height, 33));
    media.media_type = MediaType::StaticImage;

    for (column, drawn) in painted(media).into_iter().enumerate() {
        let of_the_picture = ((column as i32 + 123).rem_euclid(250)) as u8 + 1;

        assert_eq!(
            drawn, of_the_picture,
            "column {column} of the window is the picture's own column {of_the_picture}, not {drawn}"
        );
    }

    // And the same window from a frame that is not the size of the box, which the sampler takes: a picture
    // of stripes, one column of each, so where the place falls between two of them is what every column of
    // the window says. The place begins at -123.25, so column *x* samples the picture at `x + 122.75`: a
    // quarter of the way onto the column after it, which is 191 of 255 on a window column whose two
    // picture columns are black then white and 64 where they are white then black.
    let stripes: Vec<u8> = (0..drawn_width)
        .flat_map(|column| [(column % 2) as u8 * 255, 0, 0, 255])
        .collect();
    let mut media = create_loading_media(drawn_width, 1);
    media.frames[0] = Arc::new(ImageFrame::new(stripes, drawn_width, 1, 33));
    media.media_type = MediaType::StaticImage;

    for (column, drawn) in painted(media).into_iter().enumerate() {
        let wanted = if column % 2 == 0 { 191 } else { 64 };

        assert!(
            drawn.abs_diff(wanted) <= 1,
            "column {column} of the window is the picture sampled a quarter of a pixel in: {drawn} against {wanted}"
        );
    }

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
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
        order: vec![
            PathBuf::from("C:\\folder\\broken.mp4"),
            PathBuf::from("C:\\folder\\next.mp4"),
        ],
        missed: 0,
        shuffle: false,
        anchor: None,
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

    // And the park going is not what makes a pin come apart: a bubble is a picture this app holds
    // rather than a hole somebody else's window stands in, so a player that has gone leaves the
    // circle on the frame it has (see `a_film_behind_a_bubble_that_has_gone_leaves_the_bubble_standing`).
    update_pin_bubble_pause(None);
    assert!(
        pin_media_is_alive(false),
        "a video whose player is simply gone is still a picture in the circle, so the bubble stands"
    );

    // What does come apart is the same film as a *window*: there the player's window is the
    // picture, and a band with nothing in it is a window onto nothing — which is where a bubble
    // that has been restored is read again.
    let mut up = take_pin_for_a_test().expect("the pin is up");
    up.collapsed = false;
    stand_pin(Some(up));
    assert!(
        !pin_media_is_alive(false),
        "a video whose player is gone is a window onto nothing"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// a film behind a bubble that has gone leaves the circle standing on the picture it has, /// rather
/// than taking the whole pin down with it.
///
/// **This is the walk behind a bubble.** A step's file whose player never came up — one the
/// engine refused, one a swap for another film ended on its way — leaves the pin's own player
/// gone with nothing in hand to replace it, and a liveness read that took the pin down over it
/// closed a bubble for a press that was only asking for the next file (see `pin_media_is_alive`).
/// The exemption is the bubble's and not the pin's: the same film behind a window is a window
/// onto nothing, which is the answer it always had.
#[test]
fn a_film_behind_a_bubble_that_has_gone_leaves_the_bubble_standing() {
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
    pin.path = PathBuf::from("C:\\folder\\a-film-the-engine-refused.mp4");
    pin.collapsed = true;
    stand_pin(Some(pin));

    assert!(
        pin_media_is_alive(false),
        "a film behind a bubble with no player running is not a pin that came apart: the circle \
         shows the frame it has, and a bubble that vanished would take with it the only way back to \
         the file it was showing"
    );

    // And the same pin put back up is a window onto that film again, which is where it does come
    // apart: what was forgiven behind the bubble is the band nobody could see.
    let mut up = take_pin_for_a_test().expect("the pin is up");
    up.collapsed = false;
    stand_pin(Some(up));
    assert!(
        !pin_media_is_alive(false),
        "a film whose player is gone is still a window onto nothing where the window is what is on \
         screen"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// a swap that takes the standing film down leaves a bubble the film it is showing: the media /// is
/// the picture inside the circle, and taking it away takes the bubble's face with it.
///
/// **A step that cannot be shown is what this is for.** A player that would not start, an engine
/// that refused the file, a load that never arrives: a windowed pin owes the incoming file its band
/// and the outgoing frame is not the picture any more, but a bubble has no band — what is on the
/// screen is a frame out of this very media — so it keeps the film the press was made over until
/// another one is really installed (see `take_the_pinned_media_down_for_a_swap`).
#[test]
fn a_bubble_keeps_the_film_it_is_showing_when_a_swap_takes_the_standing_one_down() {
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
    pin.collapsed = true;
    stand_pin(Some(pin));

    take_the_pinned_media_down_for_a_swap();
    assert!(
        CURRENT_MEDIA
            .lock()
            .ok()
            .is_some_and(|media| media.is_some()),
        "the media a bubble is drawing its picture out of is not taken off the bubble by a swap \
         that has not happened yet"
    );

    // And the same take-down on a windowed pin is what it always was: the band the film was in
    // belongs to the file replacing it.
    let mut up = take_pin_for_a_test().expect("the pin is up");
    up.collapsed = false;
    stand_pin(Some(up));

    take_the_pinned_media_down_for_a_swap();
    assert!(
        CURRENT_MEDIA
            .lock()
            .ok()
            .is_some_and(|media| media.is_none()),
        "a window is a hole for the incoming player's own window, so the outgoing frame goes with \
         the file it was of"
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

#[test]
fn every_film_has_a_band_the_pin_wins_the_point_of() {
    assert!(
        film_band_is_hit_testable(MediaType::Video),
        "a film is a picture the pin is carried by whichever player draws it — mpv's or          FFmpeg's — so a press on it is the pin's and not the player's window's"
    );

    for (kind, what) in [
        (
            MediaType::NativeVideo,
            "a video the media engine draws into this app's own band",
        ),
        (
            MediaType::StaticImage,
            "a picture, whose band is this window's pixels already",
        ),
        (
            MediaType::EngineSvg,
            "a document the browser engine draws in a window of its own",
        ),
        (
            MediaType::EngineFont,
            "a specimen the browser engine draws the same way",
        ),
    ] {
        assert!(
            !film_band_is_hit_testable(kind),
            "{what} is left exactly as it was, whatever player was last begun"
        );
    }
}

/// A band the pin wins the point of is one part in 255, over the band and nowhere else, and that
/// part is a change no eye can see.
///
/// Both halves are the point. The coverage is the whole of what such a band is given, so a fill
/// that reached one row past it would claim a pixel that is not the pin's — which is the click this
/// may not steal — and a fill of nothing would claim none at all. And what one part in 255 is worth
/// on screen is the app's own compositing of a premultiplied pixel over an opaque backdrop, at
/// every backdrop a film can have behind it: at most one step of 255, and that one step only at the
/// darkest end.
#[test]
fn a_band_the_pin_wins_is_one_part_in_255_and_only_over_the_band() {
    let (width, height) = (8u32, 8u32);
    let (origin_y, band_height) = (2u32, 3u32);
    let mut out = vec![0u8; (width * height * 4) as usize];

    fill_band_hit_testable(&mut out, width, origin_y, band_height);

    for row in 0..height as usize {
        let in_band = row >= origin_y as usize && row < (origin_y + band_height) as usize;
        for column in 0..width as usize {
            assert_eq!(
                surface_pixel(&out, (width, height), row, column),
                if in_band {
                    &[1, 1, 1, 1]
                } else {
                    &[0, 0, 0, 0]
                },
                "row {row} column {column}: the band is the least a pixel can be, and every row \
                 outside it is left exactly as the window painted it"
            );
        }
    }

    for backdrop in 0u32..=255 {
        let mut pixel = [backdrop as u8, backdrop as u8, backdrop as u8, 255];
        blend_pixel_over(&[1, 1, 1, 1], &mut pixel, backdrop, backdrop, backdrop);
        assert!(
            (pixel[0] as i32 - backdrop as i32).abs() <= 1,
            "one part in 255 over a backdrop of {backdrop} leaves {}, which is more than the one \
             step of 255 nothing can see",
            pixel[0]
        );
    }
}

/// A double-click puts a pinned window in the middle of the screen it stands on, and the size it
/// had is the size it keeps: what is moved is the window's own box, bands and all, and for a window
/// fitted to the screen that box is already the middle of it (see `centred_pin_box`).
#[test]
fn a_double_click_puts_a_pin_in_the_middle_of_its_screen() {
    let room = ScreenBounds {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1040,
    };

    // A kind whose chrome is drawn over its media is one box, so a fit to the screen — the whole
    // of the room's height, which is what a fit leaves a picture — is put back exactly where it
    // was: the middle of the room is the place it already stands in, so no gap is opened above it
    // and none below.
    let fitted = (135, 0, 1785, 1040);
    assert_eq!(
        centred_pin_box(fitted, 96, false, true, 0, room),
        fitted,
        "a window fitted to the screen was moved off the middle of it, or given a gap above it or \
         below"
    );

    // A window smaller than the room is moved to the middle of it, size and all: a 400 by 300 box
    // in the corner comes out 760 across and 370 down.
    assert_eq!(
        centred_pin_box((0, 0, 400, 300), 96, false, true, 0, room),
        (760, 370, 1160, 670),
        "a window smaller than the screen was not put in the middle of it"
    );

    // A kind whose chrome is in bands is centred as the window, bands and all: the media box that
    // comes out of it is the one that leaves the window's own box in the middle, which is why the
    // media of such a pin is not simply the middle of the room.
    let bands = 30 + pinned_transport_height(96, true);
    let banded_window = (0, 0, 400, 300 + bands);
    let banded = centred_pin_box(banded_window, 96, true, false, 30, room);
    let window = pinned_window_box_of(banded, 96, true, false, 30);

    assert_eq!(
        (window.2 - window.0, window.3 - window.1),
        (400, 300 + bands),
        "centring changed the size of a window whose chrome is in bands"
    );
    assert_eq!(
        window.1,
        (room.bottom - room.top - (300 + bands)) / 2,
        "the box put in the middle is the window's own, bands and all, rather than the media's"
    );
}

/// The backdrop a zoomed picture is drawn over fills the band *less* the box the picture lands on, and
/// nothing else: the pixels the picture covers are written by the compose that follows, and on a band of
/// a display's own size filling them first was 2.0 ms of the 6.94 ms a 144 Hz frame has.
///
/// The band is a surface of the window's own size, as the paint's is, and what is not filled is a
/// sentinel: the answer this asks for is exactly which pixels the fill touched.
#[test]
fn a_backdrops_fill_leaves_the_box_the_picture_lands_on_alone() {
    let (width, height) = (24u32, 8u32);
    let surface = DibSurface::create(width, height).expect("a band to fill");
    let out =
        unsafe { std::slice::from_raw_parts_mut(surface.bits(), (width * height * 4) as usize) };

    let sentinel = [7u8, 7, 7, 7];
    for pixel in out.as_chunks_mut::<4>().0.iter_mut() {
        *pixel = sentinel;
    }

    // A picture landing on the middle of the band: rows 2 to 5, columns 8 to 15.
    fill_band_backdrop(
        out,
        width,
        0,
        height,
        (8, 2, 8, 4),
        TransparentBackground::White,
    );

    let at = |x: u32, y: u32| {
        let start = (y * width + x) as usize * 4;
        [out[start], out[start + 1], out[start + 2], out[start + 3]]
    };
    let backdrop = [255u8, 255, 255, 255];

    assert_eq!(
        at(8, 0),
        backdrop,
        "the rows above the picture are the backdrop"
    );
    assert_eq!(
        at(0, 3),
        backdrop,
        "and the columns before it on a row it lands on"
    );
    assert_eq!(at(23, 3), backdrop, "and the columns after it");
    assert_eq!(at(8, 7), backdrop, "and the rows below it");

    assert_eq!(
        at(8, 2),
        sentinel,
        "the picture's own box is left exactly as it was"
    );
    assert_eq!(at(15, 5), sentinel, "to its last row and column");
    assert_eq!(
        at(7, 3),
        backdrop,
        "and the column it starts at is the first of the backdrop"
    );
    assert_eq!(
        at(16, 3),
        backdrop,
        "as the one it stops at is the first of the backdrop after it"
    );
}

/// A fill of a transparent backdrop writes nothing at all: the band a picture with transparency in it
/// leaves behind is the window's own emptiness, which is what shows the desktop through a layered window
/// and what the caller's surface already holds (see `fill_band_backdrop`).
#[test]
fn a_transparent_backdrop_is_left_as_the_band_already_was() {
    let (width, height) = (8u32, 4u32);
    let surface = DibSurface::create(width, height).expect("a band to fill");
    let out =
        unsafe { std::slice::from_raw_parts_mut(surface.bits(), (width * height * 4) as usize) };

    let sentinel = [7u8, 7, 7, 7];
    for pixel in out.as_chunks_mut::<4>().0.iter_mut() {
        *pixel = sentinel;
    }

    // A picture landing on the middle of the band, the same box the test above fills around.
    fill_band_backdrop(
        out,
        width,
        0,
        height,
        (2, 1, 4, 2),
        TransparentBackground::Transparent,
    );

    assert!(
        out.as_chunks::<4>()
            .0
            .iter()
            .all(|pixel| *pixel == sentinel),
        "a transparent backdrop wrote over pixels the caller had already placed"
    );
}

/// A checkerboard backdrop keeps the squares it had wherever the picture leaves it: the colour of a pixel
/// is a question about the pixel's own column of the *band*, and the strip beside a picture is filled in
/// two runs — one before the picture and one after it — so a run that counted its own columns would
/// start the second strip's squares again from its own edge.
#[test]
fn a_checkerboard_backdrop_keeps_its_squares_wherever_the_picture_leaves_it() {
    let (width, height) = (48u32, 4u32);
    let surface = DibSurface::create(width, height).expect("a band to fill");
    let out =
        unsafe { std::slice::from_raw_parts_mut(surface.bits(), (width * height * 4) as usize) };
    out.fill(0);

    // A picture across the band's first sixteen columns, so the strip that is filled starts at column 16
    // — which is the second square of a checkerboard sixteen pixels across.
    fill_band_backdrop(
        out,
        width,
        0,
        height,
        (0, 0, 16, height as i32),
        TransparentBackground::Checkerboard,
    );

    let at = |x: u32, y: u32| {
        let start = (y * width + x) as usize * 4;
        [out[start], out[start + 1], out[start + 2], out[start + 3]]
    };

    // A square is sixteen pixels, and the colours are the pair `checkerboard_color` names: the squares of
    // the band's own first row start dark and the next one along is light.
    assert_eq!(
        at(16, 0),
        [144, 144, 144, 255],
        "the strip begins on the band's own second square"
    );
    assert_eq!(
        at(31, 0),
        [144, 144, 144, 255],
        "which runs to the end of its sixteen"
    );
    assert_eq!(
        at(32, 0),
        [224, 224, 224, 255],
        "and the square after it is the other colour"
    );
    assert_eq!(
        at(0, 0),
        [0, 0, 0, 0],
        "the picture's own columns were left as the caller left them"
    );
}

/// **A page too narrow for its caption is still the whole of a window.**
///
/// A caption too narrow to be a command bar is drawn none at all, and the rows it would have taken
/// are rows the window still has (see `caption_fits` and `pinned_band_rows`). A picture draws its
/// chrome *over* its media, so the media is under those rows and owes nothing; a page is drawn
/// *between* its chrome, so a paint that stops at the band would leave them as the surface was
/// emptied — a hole across the top of the window, with the desktop showing through the pin and a
/// press falling through it, since a layered window is hit-tested by the shape of its pixels. What
/// stands in them is the pin's own claim instead: the strips carry the chrome, and the bitmap owes
/// those rows only the least pixel a press can land on (see `chrome_rows_are_hit_testable`).
#[test]
fn a_page_too_narrow_for_its_caption_is_still_a_whole_window() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME.lock();
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    // A page on a band of its own colour, so every row of the window says where it came from.
    let page = [200u8, 200, 200, 255];

    // A box the hand has narrowed a page to: there is no room in it for the caption's own buttons.
    let content = (100, 80, 220, 300);
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.dpi = 96;
    pin.frame = PinFrame::Free;
    pin.overlay = false;
    pin.zoom = None;
    let caption = pin.caption;
    stand_pin(Some(pin));

    let (width, height) = the_pins_window_size();
    assert!(
        caption > 0 && !pin_chrome::caption_fits(width, 96, true),
        "a page this narrow has no room for its caption's buttons, so none of it is drawn"
    );

    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(static_image_media(
            Arc::new(ImageFrame::new(page.repeat(120 * 220), 120, 220, 0)),
            MediaType::Text,
        ));
    }
    let hwnd = HWND(0x1000 as *mut _);
    unsafe { render_pinned_preview_at(hwnd, content.0, content.1) };

    let painted = the_painted_surface(hwnd, width, height);
    for row in 0..caption as usize {
        for column in 0..width as usize {
            assert_eq!(
                surface_pixel(&painted, (width as u32, height as u32), row, column),
                &[1, 1, 1, 1],
                "row {row} column {column} is the pin's own claim even though the caption is not \
                 drawn, rather than a hole across the top of the window"
            );
        }
    }
    assert_eq!(
        surface_pixel(&painted, (width as u32, height as u32), caption as usize, 0),
        page,
        "while the media's own rows under the bar are the page's, not the bar's"
    );

    // And a picture of the same narrowness owes none of it: its chrome is drawn over the media, so
    // the rows the caption gave up are the picture's own — filling them would paint the bar across
    // the top of somebody's picture, which is the band the strip was given up not to take.
    with_pin(|pin| {
        pin.frame = PinFrame::Shaped;
        pin.overlay = true;
    });
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = Some(static_image_media(
            Arc::new(ImageFrame::new(page.repeat(120 * 220), 120, 220, 0)),
            MediaType::StaticImage,
        ));
    }
    let (width, height) = the_pins_window_size();
    unsafe { render_pinned_preview_at(hwnd, content.0, content.1) };
    let painted = the_painted_surface(hwnd, width, height);
    assert_eq!(
        surface_pixel(&painted, (width as u32, height as u32), 0, 0),
        page,
        "a picture's own pixels stand in the rows a caption that does not fit would have taken"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}

/// The size of the pin standing now, as the window and the paint are given it.
fn the_pins_window_size() -> (i32, i32) {
    pin_state()
        .and_then(|pinned| pinned.pin().map(|pin| pin.window_size()))
        .expect("a pin up")
}

/// The surface the paint of the pin standing now was written into.
fn the_painted_surface(hwnd: HWND, width: i32, height: i32) -> Vec<u8> {
    let bits = ensure_layered_surface(hwnd.0 as isize, width as u32, height as u32)
        .expect("the surface the paint wrote");
    unsafe { std::slice::from_raw_parts(bits, width as usize * height as usize * 4) }.to_vec()
}

/// A resize paints its picture into a surface with a row of slack in it, and stands the window a row
/// higher for every row it pushes the picture down into it (see `the_slack_a_resize_paints_with`).
///
/// The two halves of that are the whole of it, and both are asserted here: the picture lands on the
/// box the pointer asked for on **every** paint of the drag, and the window itself changes place on
/// every one of them. The first is what keeps the drag smooth — a step that moved the window and took
/// the picture with it is a picture a pixel off, or shivering, for as long as the drag gives the window
/// less than it had. The second is what keeps the ghost away: a window whose only change is a size
/// leaves the strip it gives up standing, and a window that moves has the desk under the place it has
/// left drawn again.
#[test]
fn a_resize_paints_a_row_of_slack_so_the_window_moves_and_the_picture_does_not() {
    // A box 800 by 600 standing at 100,100. What the window is asked for, paint after paint.
    let (x, y, height) = (100, 100, 600);

    let mut places = Vec::new();
    let mut landed = Vec::new();
    for _ in 0..6 {
        let (padding, slack) = paint::the_slack_a_resize_paints_with(true);
        assert_eq!(
            slack, 1,
            "a resize's surface is one row taller than the pin"
        );

        let place = (x, y - padding);
        places.push(place);
        // The picture's own top and bottom on the screen: the window's place plus where it stands in
        // the surface, which is the row the paint gave it.
        landed.push((place.1 + padding, place.1 + padding + height));
    }

    for (top, bottom) in &landed {
        assert_eq!(
            (*top, *bottom),
            (y, y + height),
            "the picture was not on the box the pointer asked for"
        );
    }
    for pair in places.windows(2) {
        assert_ne!(
            pair[0], pair[1],
            "the window stood still from one paint to the next, which leaves the strip it gave up"
        );
        assert_eq!(
            (pair[0].1 - pair[1].1).abs(),
            1,
            "the window moved by more than the row of slack it carries"
        );
    }

    // And a paint that is not a resize has no slack in it at all: the window stands on its own box,
    // where the pin says it does.
    assert_eq!(
        paint::the_slack_a_resize_paints_with(false),
        (0, 0),
        "a paint that is not a resize asks the surface for nothing"
    );
    assert_eq!(paint::the_slack_a_resize_paints_with(false), (0, 0));
}

/// The children the pin's window carries are placed in its client coordinates and so ride every move
/// of it — including the row a resize moves it up by. The row the last paint chose is what they are
/// put back down by, and it answers only while a resize is the drag in flight, so a release that
/// never repaints cannot leave a row standing over a window that has stopped moving (see
/// `the_row_a_resize_paints_the_picture_down`).
#[test]
fn the_children_are_put_back_down_by_the_row_a_resize_paints_the_picture_down() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|held| held.into_inner());
    let previous_pin = take_pin_for_a_test();

    let content = (100, 100, 500, 400);
    let mut pin = overlay_pin(content, PinChrome::always());
    pin.caption = 30;
    pin.overlay = false;
    stand_pin(Some(pin));

    // A paint that is not a resize has no row, even one made just after a resize's own paint chose
    // one: the drag in flight is the whole of what the answer stands for.
    assert_eq!(paint::the_row_a_resize_paints_the_picture_down(), 0);

    with_pin(|pin| {
        pin.dragging = Some(PinDrag {
            from: (0, 0),
            window: content,
            action: PinDragAction::Resize(edge(false, false, true, false)),
            delivered: true,
            carried: (i32::MIN, i32::MIN),
        });
    });

    // Every paint of the drag asks for its row, and the chrome's strips are placed by that same row:
    // the pin's picture lands on the box the pointer asked for either way (see
    // `the_slack_a_resize_paints_with`).
    for _ in 0..4 {
        let (padding, slack) = paint::the_slack_a_resize_paints_with(true);
        assert_eq!(
            slack, 1,
            "a resize's surface is one row taller than the pin"
        );
        assert_eq!(
            paint::the_row_a_resize_paints_the_picture_down(),
            padding,
            "the children are owed the row this paint just chose"
        );
    }

    // A release that never repaints answers nothing: the row is not left applied to children whose
    // window has stopped moving.
    with_pin(|pin| pin.dragging = None);
    assert_eq!(paint::the_row_a_resize_paints_the_picture_down(), 0);

    stand_pin(previous_pin);
}

/// Every road that installs a file in a pin — a first pin, the same file pinned again, a step to
/// the next or previous one — brings the chrome up hidden: a stable below shows without chrome,
/// no arrival is armed, and the pointer asks the strips back afterwards (see
/// `take_up_pinned_window`).
#[test]
fn every_spawn_hides_the_pins_chrome_until_the_pointer_asks_for_it() {
    let _one = pin_window::PIN_TESTS_ONE_AT_A_TIME
        .lock()
        .unwrap_or_else(|e| e.into_inner());
    let previous_media = CURRENT_MEDIA.lock().ok().and_then(|mut media| media.take());
    let previous_pin = take_pin_for_a_test();

    let stand_media = |kind| {
        let mut media = create_loading_media(320, 240);
        media.media_type = kind;
        if let Ok(mut current) = CURRENT_MEDIA.lock() {
            *current = Some(media);
        }
    };

    // A first pin comes up with neither strip showing and no arrival armed behind it.
    stand_media(MediaType::StaticImage);
    let arrived = take_up_pinned_window(&PathBuf::from("first-picture.png"), (100, 80, 420, 320));
    assert!(
        !arrived.chrome.caption && !arrived.chrome.bar && arrived.chrome.until.is_none(),
        "a first pin comes up hidden: caption {}, bar {}, arrival {:?}",
        arrived.chrome.caption,
        arrived.chrome.bar,
        arrived.chrome.until
    );

    // Pinning the file that is already up carries no chrome either, for the same reason a step
    // does not: it is an install, and an install brings the chrome up hidden.
    stand_pin(Some(arrived));
    stand_media(MediaType::StaticImage);
    let again = take_up_pinned_window(&PathBuf::from("first-picture.png"), (100, 80, 420, 320));
    assert!(
        !again.chrome.caption && !again.chrome.bar && again.chrome.until.is_none(),
        "pinning the file again comes up hidden: caption {}, bar {}, arrival {:?}",
        again.chrome.caption,
        again.chrome.bar,
        again.chrome.until
    );

    // And a step onto another picture carries none of that over: the cover the swap stands shows
    // without chrome, and the arrival window is not armed again.
    stand_pin(Some(again));
    stand_media(MediaType::StaticImage);
    let mut stepped =
        take_up_pinned_window(&PathBuf::from("second-picture.png"), (100, 80, 420, 320));
    assert!(
        !stepped.chrome.caption && !stepped.chrome.bar && stepped.chrome.until.is_none(),
        "a step hides the chrome outright: caption {}, bar {}, arrival {:?}",
        stepped.chrome.caption,
        stepped.chrome.bar,
        stepped.chrome.until
    );

    // The pointer asks it back: near the caption after the arrival would have closed, the strip
    // comes out, and away from the window it stays in.
    let window = stepped.window_box();
    let after_arrival = Instant::now() + Duration::from_secs(2);
    let changed = refresh_pin_chrome(&mut stepped, after_arrival, Some((window.0 + 5, window.1)));
    assert!(
        changed && stepped.chrome.caption,
        "a hover after the step shows the caption again"
    );
    refresh_pin_chrome(&mut stepped, after_arrival, Some((-1000, -1000)));
    assert!(
        !stepped.chrome.caption,
        "while away from the window the stepped-on chrome stays hidden"
    );

    // And a step across kinds arms no arrival either: a pin of a bands kind carries an always-on
    // caption, and the film stepped onto comes up hidden all the same.
    stand_pin(Some(stepped));
    stand_media(MediaType::Text);
    let bands = take_up_pinned_window(&PathBuf::from("note.txt"), (100, 80, 420, 320));
    assert!(
        !bands.hides_chrome,
        "a text pin keeps its chrome in bands, which is the carry this step breaks"
    );
    stand_pin(Some(bands));
    stand_media(MediaType::Video);
    let crossed = take_up_pinned_window(&PathBuf::from("film.mkv"), (100, 80, 420, 320));
    assert!(
        !crossed.chrome.caption && !crossed.chrome.bar && crossed.chrome.until.is_none(),
        "a step onto a film hides the chrome rather than re-arming the arrival"
    );

    stand_pin(previous_pin);
    if let Ok(mut media) = CURRENT_MEDIA.lock() {
        *media = previous_media;
    }
}
